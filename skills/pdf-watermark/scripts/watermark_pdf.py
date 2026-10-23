import argparse
import sys
from io import BytesIO
from pathlib import Path

import pymupdf
from PIL import Image


def _is_pdf(path: Path) -> bool:
    try:
        with open(path, "rb") as fh:
            return fh.read(5) == b"%PDF-"
    except OSError:
        return False


def docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
        doc.SaveAs(str(pdf_path.resolve()), FileFormat=17)
        doc.Close(False)
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def load_logo(path: Path, white_to_alpha: bool) -> Image.Image:
    logo = Image.open(path).convert("RGBA")
    if white_to_alpha:
        import numpy as np

        arr = np.array(logo)
        r, g, b, a = arr[..., 0], arr[..., 1], arr[..., 2], arr[..., 3]
        white = (r > 240) & (g > 240) & (b > 240)
        arr[..., 3] = np.where(white, 0, a)
        logo = Image.fromarray(arr, "RGBA")
    bbox = logo.getchannel("A").getbbox()
    return logo.crop(bbox) if bbox else logo


def build_tile(logo: Image.Image, target_w: int, rotate: float, opacity: float) -> Image.Image:
    target_w = max(1, target_w)
    target_h = max(1, round(logo.height * target_w / logo.width))
    tile = logo.resize((target_w, target_h), Image.LANCZOS)
    tile = tile.rotate(rotate, expand=True, resample=Image.BICUBIC)
    if opacity < 1.0:
        alpha = tile.getchannel("A").point(lambda v: int(v * opacity))
        tile.putalpha(alpha)
    return tile


def build_overlay(width: int, height: int, tiles, gap: float, stagger: bool) -> Image.Image:
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    for tile in tiles:
        step_x = max(1, round(tile.width * (1 + gap)))
        step_y = max(1, round(tile.height * (1 + gap)))
        row = 0
        y = -tile.height
        while y < height:
            offset = step_x // 2 if (stagger and row % 2) else 0
            x = -tile.width + offset
            while x < width:
                overlay.paste(tile, (x, y), tile)
                x += step_x
            y += step_y
            row += 1
    return overlay


def process_pdf(
    pdf_path: Path,
    logo_path: Path,
    out_path: Path,
    dpi: int,
    tile_size: float,
    rotate: float,
    opacity: float,
    gap: float,
    stagger: bool,
    white_to_alpha: bool,
    crosshatch: bool = False,
    angle2: float = -45.0,
    opacity2: float = None,
) -> None:
    src = pymupdf.open(str(pdf_path))
    out = pymupdf.open()
    logo = load_logo(logo_path, white_to_alpha)
    zoom = dpi / 72.0
    total = src.page_count
    if opacity2 is None:
        opacity2 = opacity
    for index in range(total):
        page = src[index]
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
        image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).convert("RGBA")
        tile_w = max(1, round(image.width * tile_size))
        tiles = [build_tile(logo, tile_w, rotate, opacity)]
        if crosshatch:
            tiles.append(build_tile(logo, tile_w, angle2, opacity2))
        overlay = build_overlay(image.width, image.height, tiles, gap, stagger)
        image.alpha_composite(overlay)
        image = image.convert("RGB")
        buffer = BytesIO()
        image.save(buffer, format="PNG")
        new_page = out.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(new_page.rect, stream=buffer.getvalue())
        print(f"[{index + 1}/{total}] done  ({image.width}x{image.height}px)", flush=True)
    out.save(str(out_path), deflate=True, garbage=3)
    out.close()
    src.close()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert a Word/PDF to a rasterized PDF with a tiled, rotated PNG watermark burned into every page."
    )
    parser.add_argument("input", type=Path, help="input file (.docx or .pdf)")
    parser.add_argument("-o", "--output", type=Path, help="output PDF path")
    parser.add_argument("-l", "--logo", type=Path, required=True, help="watermark logo (.png, transparent)")
    parser.add_argument("-d", "--dpi", type=int, default=300, help="rasterize DPI (default 300, print quality; 150-200 for screen)")
    parser.add_argument("--tile-size", type=float, default=0.22, help="watermark tile width as fraction of page width (default 0.22)")
    parser.add_argument("--rotate", type=float, default=45.0, help="rotation angle in degrees (default 45)")
    parser.add_argument("--opacity", type=float, default=0.12, help="watermark opacity 0-1 (default 0.12)")
    parser.add_argument("--gap", type=float, default=0.15, help="gap between tiles as fraction of tile size (default 0.15)")
    parser.add_argument("--no-stagger", action="store_true", help="disable brick-style row offset")
    parser.add_argument("--white-to-alpha", action="store_true", help="treat near-white logo pixels as transparent")
    parser.add_argument("--crosshatch", action="store_true", help="overlay a second tiled layer in the opposite direction (harder to remove)")
    parser.add_argument("--angle2", type=float, default=-45.0, help="second layer rotation angle for --crosshatch (default -45)")
    parser.add_argument("--opacity2", type=float, default=None, help="second layer opacity for --crosshatch (default = --opacity)")
    args = parser.parse_args()

    src = args.input
    if not src.exists():
        parser.error(f"input not found: {src}")
    if not args.logo.exists():
        parser.error(f"logo not found: {args.logo}")

    out = args.output or src.with_name(f"{src.stem}_watermarked.pdf")
    suffix = src.suffix.lower()
    tmp_pdf = None
    try:
        if suffix in (".docx", ".doc"):
            tmp_pdf = src.with_name(f"{src.stem}_tmp.pdf")
            print("converting Word -> PDF ...", flush=True)
            docx_to_pdf(src, tmp_pdf)
            pdf_in = tmp_pdf
        elif suffix == ".pdf":
            pdf_in = src
        elif _is_pdf(src):
            print("input has no .pdf suffix but looks like a PDF, processing as PDF ...", flush=True)
            pdf_in = src
        else:
            parser.error("input must be .docx, .doc or .pdf")

        print(f"watermarking {pdf_in.name} -> {out.name} @ {args.dpi} DPI ...", flush=True)
        process_pdf(
            pdf_in,
            args.logo,
            out,
            args.dpi,
            args.tile_size,
            args.rotate,
            args.opacity,
            args.gap,
            not args.no_stagger,
            args.white_to_alpha,
            args.crosshatch,
            args.angle2,
            args.opacity2,
        )
        print(f"\nsaved: {out}")
    finally:
        if tmp_pdf and tmp_pdf.exists():
            tmp_pdf.unlink()
    return 0


if __name__ == "__main__":
    sys.exit(main())

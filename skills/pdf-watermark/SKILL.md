---
name: pdf-watermark
description: Burn a tiled, rotated logo watermark into every page of a PDF or Word document by rasterizing each page, so the watermark becomes part of the image pixels and cannot be removed as a separate PDF layer. Use when the user wants to watermark PDF/DOCX files, protect documents from watermark removal, or batch-apply a company or brand logo watermark locally.
license: Apache-2.0
compatibility: Requires Python 3.9+ with PyMuPDF and Pillow. Converting .docx/.doc to PDF requires Windows with Microsoft Word installed and pywin32.
metadata:
  version: "1.0.0"
  author: watermark-tool contributors
  opencode/slash: "false"
---

# PDF Watermark

Rasterize each page of a PDF (or Word document), composite a tiled logo watermark onto the pixels, and rebuild the PDF. Because the watermark is part of the page image, PDF object editors cannot delete it as a layer; removal requires expensive per-page image editing.

## When to use

- The user asks to add a logo/watermark to PDFs or Word documents.
- The document must be hard to un-watermark with Acrobat, pdfcpu, or similar PDF tools.
- Files must be processed locally and not uploaded to a third party.

Do not use when the output must keep selectable/searchable text, accessibility, or a small file size. Rasterization destroys the text layer.

## Requirements

- Python 3.9+.
- `pip install -r requirements.txt` from the repository root, or `pip install PyMuPDF Pillow numpy`.
- `.docx` / `.doc` input: Windows + Microsoft Word + `pywin32`.

## Workflow

1. Confirm the input file, logo image, and desired output path.
2. If the user has no preference, use the recommended preset:
   `--dpi 300 --opacity 0.24 --tile-size 0.26 --gap 0.08 --rotate 45`
3. Add `--white-to-alpha` when the logo is a JPG with a white background.
4. Run the script from this skill directory, or use absolute paths:

   ```bash
   python scripts/watermark_pdf.py "input/report.pdf" \
     -l "assets/logo.png" \
     -o "output/report_watermarked.pdf" \
     --white-to-alpha --dpi 300 --opacity 0.24 --tile-size 0.26 --gap 0.08
   ```

5. Verify the output exists, then report its path, page count, and file size.
6. Remind the user that text is no longer selectable/searchable and the file is larger.

## Input handling

| Input | Behavior |
|---|---|
| `.docx` / `.doc` | Converts to PDF through Microsoft Word, then watermarks |
| `.pdf` | Rasterizes and watermarks directly |
| No extension but starts with `%PDF-` | Treated as PDF |

## Parameters

| Flag | Default | Meaning |
|---|---|---|
| `input` | required | Input `.docx`, `.doc`, or `.pdf` |
| `-o, --output` | `<name>_watermarked.pdf` | Output PDF path |
| `-l, --logo` | required | Logo image; transparent PNG works best |
| `-d, --dpi` | `300` | Rasterization DPI; 150-200 is enough for screen, 300 for print |
| `--tile-size` | `0.22` | Tile width as a fraction of page width; smaller = denser |
| `--rotate` | `45` | Tile rotation in degrees |
| `--opacity` | `0.12` | Watermark opacity from 0 to 1 |
| `--gap` | `0.15` | Gap between tiles as a fraction of tile size |
| `--no-stagger` | off | Disable brick-style row offset |
| `--white-to-alpha` | off | Make near-white pixels transparent (use for white-background JPG logos) |
| `--crosshatch` | off | Add a second layer in the opposite direction; harder to remove but logos overlap |
| `--angle2` | `-45` | Second-layer angle for `--crosshatch` |
| `--opacity2` | same as `--opacity` | Second-layer opacity for `--crosshatch` |

## Recommended presets

- Clear brand, single layer (recommended):
  `--opacity 0.24 --tile-size 0.26 --gap 0.08 --rotate 45`
- Harder to remove, crosshatched:
  `--crosshatch --opacity 0.22 --opacity2 0.15 --tile-size 0.22 --gap 0.08`

## Limitations

- No watermarking scheme is 100% removal-proof. AI/Photoshop inpainting can still remove part of the mark, but only page by page and with quality loss.
- Faint watermarks can be reduced by global contrast adjustments; more visible marks are harder to remove.
- Rasterized output cannot be copied, searched, or read by accessibility tools.
- On Windows, if Word COM raises an error, close leftover `WINWORD` processes and retry.

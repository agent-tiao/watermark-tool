# watermark-tool

[简体中文](README.md) | English

A command-line tool that converts **Word / PDF** files into **rasterized PDFs** and **burns a logo watermark into every page's pixels**, making the watermark hard to remove. It also ships as a standard **Agent Skill** for Claude Code, opencode, and other AI clients.

## The problem: a watermark that can be deleted in one click is not a watermark

People watermark documents to stop them from being casually forwarded or leaked. But a normal PDF watermark is just a **separate layer/object**—a recipient can remove it in seconds with Acrobat, pdfcpu, or an online tool, with no technical skill required.

Once the watermark is gone, the document looks exactly like the original. When it does leak:

- You **don't know who leaked it**.
- You **don't know how it leaked**.
- You **cannot even prove the file was ever watermarked**.

A watermark should be a deterrent and a trace. A one-click-removable watermark is just decoration. **That is why this tool exists.**

**Typical use cases**: quotations, technical documentation, training material, design drafts, bids, and any partner-facing document whose distribution needs to be controlled.

## The solution

`watermark-tool` does not treat the watermark as a layer. It **burns it into every page's pixels**:

1. Render every page to a bitmap.
2. Composite a semi-transparent logo across the page at an angle, with staggered tiling.
3. Repack the bitmaps into a PDF.

The watermark is now part of the page image. Removing it requires page-by-page image editing, which is lossy, slow, and expensive.

**Features**

- Handles `.docx` / `.doc` / `.pdf` and picks the right pipeline automatically
- Adjustable opacity, density, and rotation, plus optional crosshatch tiling
- Runs entirely locally; files are never uploaded
- Includes an Agent Skill so AI clients can batch-watermark documents
- Apache-2.0 licensed; free to modify and redistribute

**Trade-offs**: text is no longer selectable or searchable, and output files are larger.

## How it works

```
Word/PDF  ->  (Word to PDF)  ->  rasterize each page (DPI)  ->  overlay tiled watermark  ->  repack as PDF
```

- **Rasterization**: content and watermark both become pixels, so the watermark can no longer be separated or deleted in one click.
- **Tiling + rotation + transparency**: covers the whole page, cannot be cropped away, does not hide the content, but remains clearly visible.
- **Trade-offs**: no text selection/search, larger files.

## Repository layout

```
watermark-tool/
├─ skills/
│  └─ pdf-watermark/         # Agent Skill
│     ├─ SKILL.md
│     └─ scripts/
│        └─ watermark_pdf.py # main script
├─ assets/
│  └─ sample_logo.png        # generic sample watermark
├─ input/                    # source files (gitignored)
├─ output/                   # generated PDFs (gitignored)
├─ requirements.txt
├─ LICENSE
├─ README.md
├─ README.en.md
└─ .gitignore
```

> Before publishing to a public repository, confirm that `input/`, `output/`, and your brand logo are excluded by `.gitignore` to avoid leaking customer documents or internal assets.

## Requirements

```bash
python -m pip install -r requirements.txt
```

- **Windows with Microsoft Word** (for `.docx -> .pdf` conversion; not needed when processing only `.pdf`)
- Python packages: PyMuPDF, Pillow, pywin32, numpy

## Quick start

Run from the repository root:

```bash
# Word input
python skills/pdf-watermark/scripts/watermark_pdf.py "input/xxx.docx" -l "assets/sample_logo.png" --white-to-alpha -d 300 --opacity 0.24 --tile-size 0.26 --gap 0.08 -o "output/xxx_watermarked.pdf"

# PDF input (rasterize and watermark directly)
python skills/pdf-watermark/scripts/watermark_pdf.py "input/xxx.pdf" -l "assets/sample_logo.png" -d 300 --opacity 0.24 --tile-size 0.26 --gap 0.08 -o "output/xxx_watermarked.pdf"
```

Replace `assets/sample_logo.png` with your own logo (a transparent PNG works best; add `--white-to-alpha` for JPG logos with a white background).

## Using as an Agent Skill

`skills/pdf-watermark/` follows the [Agent Skills](https://agentskills.io) specification (`SKILL.md` + scripts) and works with Claude Code, opencode, and other compatible clients.

**Install globally**

```bash
# macOS / Linux
mkdir -p ~/.claude/skills
cp -r skills/pdf-watermark ~/.claude/skills/

# opencode
mkdir -p ~/.config/opencode/skills
cp -r skills/pdf-watermark ~/.config/opencode/skills/
```

```powershell
# Windows PowerShell
New-Item -ItemType Directory -Force "$HOME\.claude\skills" | Out-Null
Copy-Item -Recurse "skills\pdf-watermark" "$HOME\.claude\skills\"
```

You can also install it per project under `.claude/skills/` or `.opencode/skills/`. Restart the client after installing.

**Usage**: simply tell the AI

> Use pdf-watermark to watermark `input/report.pdf` with opacity 0.24 and write the result to `output/`.

The AI follows the workflow in `SKILL.md`, chooses parameters, runs the script, and verifies the result.

## Supported inputs

The script selects the pipeline based on the input type:

| Input | Behavior |
|---|---|
| `.docx` / `.doc` | Convert to PDF with Microsoft Word, then watermark |
| `.pdf` | Rasterize and watermark directly |
| No extension but starts with `%PDF-` | Treated as PDF (handles missing extensions) |

## Parameters

| Flag | Default | Description |
|---|---|---|
| `input` | required | Input file (`.docx` / `.doc` / `.pdf`) |
| `-o, --output` | `<name>_watermarked.pdf` | Output PDF path |
| `-l, --logo` | required | Watermark logo image (transparent PNG preferred) |
| `-d, --dpi` | `300` | Rasterization DPI. 300 for print; 150-200 is enough for screen and produces smaller files |
| `--tile-size` | `0.22` | Watermark width as a fraction of page width; smaller = denser |
| `--rotate` | `45` | Watermark rotation in degrees |
| `--opacity` | `0.12` | Watermark opacity from 0 to 1; higher = more visible and harder to remove |
| `--gap` | `0.15` | Gap between tiles as a fraction of tile size |
| `--no-stagger` | off | Disable brick-style row offset |
| `--white-to-alpha` | off | Treat near-white pixels as transparent (**required for white-background JPG logos**) |
| `--crosshatch` | off | Crosshatch tiling: overlay two directions. Harder to remove, but logos overlap |
| `--angle2` | `-45` | Second-layer angle for `--crosshatch` |
| `--opacity2` | same as `--opacity` | Second-layer opacity for `--crosshatch` |

## Recommended presets

- **Single layer, no overlap (recommended for brand clarity)**
  ```bash
  --opacity 0.24 --tile-size 0.26 --gap 0.08 --rotate 45
  ```
- **Crosshatch (harder to remove, but logos overlap and brand clarity drops)**
  ```bash
  --crosshatch --opacity 0.22 --opacity2 0.15 --tile-size 0.22 --gap 0.08
  ```

## Usage notes and limitations

- **No watermark is 100% removal-proof.** Pure PDF watermark removers (Acrobat, pdfcpu, etc.) **do not work** on this tool's output because there is no removable object. AI/Photoshop inpainting **can still remove part of the watermark**, but it is lossy, must be done page by page, and is expensive.
- **The more visible the watermark, the harder it is to remove.** A faint watermark can be reduced with global contrast adjustments.
- After rasterization, text **cannot be copied or searched**, and accessibility support is lost. This approach is not suitable when a searchable document is required.
- When sending the file through chat tools, choose **"Send file"** rather than image/link conversion to avoid recompression.
- On Windows, if Word occasionally raises a COM error, close any leftover `WINWORD` processes and retry.

## License

[Apache-2.0](LICENSE).

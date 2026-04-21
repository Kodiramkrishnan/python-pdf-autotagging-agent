# Universal PDF Tagging Agent

Python-based PDF accessibility pipeline for converting digital/scanned PDFs into tagged outputs targeting PDF/UA-1 and PAC/WCAG-oriented checks.

Main script:
- `universal_pdf_tagging_agent.py`

## What This Project Does

- Vision-first layout analysis (DocTR primary, LayoutLMv3 fallback)
- Semantic structure reconstruction (headings, paragraphs, lists, tables, links, figures)
- Artifact handling for non-semantic content
- Metadata injection (title, language, PDF/UA markers)
- Bookmark generation from headings
- PAC checkpoint helper probes (font/alt/embedded file related checks)
- Optional contrast-enhancement pipeline for low-contrast scans
- Single-file and batch processing modes

## Setup

### 1) Python

- Recommended: Python 3.11+

### 2) Install Python dependencies

```powershell
python -m pip install --upgrade pip
python -m pip install pikepdf pymupdf transformers pillow pytesseract python-doctr
```

### 3) Install Tesseract OCR (Windows)

```powershell
winget install --id tesseract-ocr.tesseract -e --accept-package-agreements --accept-source-agreements
```

If needed, add Tesseract to PATH for current terminal:

```powershell
$env:PATH = "C:\Program Files\Tesseract-OCR;" + $env:PATH
```

## Title Behavior (Latest)

When `--title` is not provided, the agent now resolves document title in this order:

1. Source PDF `/Title` metadata
2. Source PDF XMP `dc:title`
3. First-page visual title heuristic (dominant large text)
4. Filename stem fallback

If you pass `--title`, that value is always used.

## Single File Usage

Basic:

```powershell
python universal_pdf_tagging_agent.py "input.pdf" "output_tagged.pdf"
```

Recommended stable run profile:

```powershell
python universal_pdf_tagging_agent.py "input.pdf" "output_tagged.pdf" --enhance-contrast --force-pac-checkpoints --force-pac-font-check --force-pac-alt-checkpoints
```

Explicit title override:

```powershell
python universal_pdf_tagging_agent.py "input.pdf" "output_tagged.pdf" --title "My Custom Title"
```

## Batch Processing (Latest)

Batch mode processes all `.pdf` files inside an input directory and writes outputs into per-file subfolders.

```powershell
python universal_pdf_tagging_agent.py --input-dir "c:\path\to\input" --output-dir "c:\path\to\output" --enhance-contrast --force-pac-checkpoints --force-pac-font-check --force-pac-alt-checkpoints
```

### Batch Output Structure

For each input file `SomeFile.pdf`, output is:

- `output\SomeFile\SomeFile.pdf` (same filename as input)
- `output\SomeFile\SomeFile_report.json`

### Report JSON Contents

Each `*_report.json` includes:

- `input_file`, `output_file`
- `status` (`success` / `failed`)
- `started_at_utc`, `finished_at_utc`, `duration_seconds`
- run `options`
- on success: `output_size_bytes`, `output_pages`, `output_title`
- on failure: `error`, `error_type`

## CLI Options

- `--lang` (default: `en-US`)
- `--title` (optional explicit title)
- `--enhance-contrast`
- `--force-pac-checkpoints`
- `--force-pac-font-check`
- `--force-pac-font-matrix` (experimental, can increase file size)
- `--force-pac-alt-checkpoints`
- `--input-dir` + `--output-dir` (batch mode)

## Validation

### PAC

Check generated PDFs in PAC:
- PDF/UA
- WCAG
- Quality
- AI

### veraPDF

Example (if installed locally):

```powershell
& "c:\Users\KRISHNA\Documents\AI-AGent\verapdf\verapdf.bat" --flavour ua1 --format text "output_tagged.pdf"
```

## Troubleshooting

- `ImportError` / OCR failures: verify Python deps + Tesseract install
- `tesseract` not found: restart shell or update PATH
- Very large output files: avoid `--force-pac-font-matrix` unless needed
- Batch mode error about arguments: do not mix positional `input_pdf output_pdf` with `--input-dir/--output-dir`

## Notes

- PAC AI and some quality checks are document-dependent, especially for complex scans.
- The code is modular; heuristics can be tuned per document family.

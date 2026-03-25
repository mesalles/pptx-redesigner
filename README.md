# pptx-redesigner

A Python tool that reads an existing `.pptx` presentation, extracts all text
content (titles, body text, bullet points, tables, notes), and generates a
brand-new, visually redesigned presentation with a **corporate/professional
blue theme** — without modifying the original file.

## Features

- Extracts **all text** from every slide: text frames, placeholders, tables,
  grouped shapes, and speaker notes
- Preserves text hierarchy (title, subtitle, body, bullet levels)
- Detects slide types automatically (cover, section header, content, closing)
- Applies a clean **corporate blue design**:
  - Dark navy cover and closing slides
  - Corporate blue section dividers
  - White/light-gray content slides with blue header bars
  - Accent lines, side bars, and footer with slide numbers
- Works with **any** `.pptx` file, not just the included sample
- Handles Spanish (UTF-8) and other non-ASCII text correctly

## Requirements

- Python 3.7+
- [python-pptx](https://python-pptx.readthedocs.io/)

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
# Redesign the default file (presentacion_final_IA.pptx)
python redesign_pptx.py

# Specify custom input and output files
python redesign_pptx.py my_presentation.pptx my_presentation_redesigned.pptx
```

The script will print progress for each slide and save the result to
`presentacion_final_IA_redesigned.pptx` (or the output path you specified).

## What the script does

1. Opens the source `.pptx` with `python-pptx`
2. Walks every slide and shape (including groups and tables) to collect text
3. Classifies each slide as *cover*, *section*, *content*, or *closing*
4. Creates a new blank widescreen (16:9) presentation
5. Rebuilds each slide with the appropriate layout and corporate blue styling
6. Saves the result — the original file is **never modified**

## Customization

All colors, fonts, and sizes are defined as constants at the top of
`redesign_pptx.py` under the **CONFIGURATION** section. Edit them to match
your brand:

| Constant         | Default     | Purpose                              |
|------------------|-------------|--------------------------------------|
| `C_NAVY`         | `#1B2A4A`   | Cover/closing slide background       |
| `C_CORP_BLUE`    | `#2E5090`   | Header bars and section backgrounds  |
| `C_ACCENT`       | `#4A90D9`   | Accent lines and side bar            |
| `FONT_TITLE`     | `Calibri`   | Font used for titles                 |
| `FONT_BODY`      | `Calibri`   | Font used for body text              |
| `PT_COVER_TITLE` | `36pt`      | Title font size on the cover slide   |
| `PT_BODY`        | `16pt`      | Body text font size                  |

Change `INPUT_FILE` and `OUTPUT_FILE` constants to set different default paths.
# DOCX Redline

A macOS tool for comparing two `.docx` files and generating a redlined document with tracked changes and a change report.

**Redline markup:**
- ~~Strikethrough red~~ = deleted text
- <u>Underlined red</u> = inserted text
- Yellow highlight = formatting changes

---

## Quick Install

Open Terminal and run:

```bash
curl -sL https://raw.githubusercontent.com/sethsaler/docx_redline/main/install.sh | bash
```

This downloads the tool to your Desktop and runs setup. If Python is missing, it will offer to install it for you.

Once installed, launch it anytime by double-clicking **`run.command`** in the `docx_redline` folder.

---

## Manual Install

If you already have the folder (copied from a USB drive, AirDrop, etc.):

1. Open **Terminal**
2. Run:
   ```bash
   cd /path/to/docx_redline
   bash setup.sh
   ```
3. Done. Double-click **`run.command`** to use the tool.

**Requirements:** macOS with Python 3.9 or later. If Python is not installed, `setup.sh` will download and install it automatically.

---

## How to Use

1. Launch the tool (double-click `run.command` or run `./run.command`)
2. Enter the path to the **original** `.docx` file
3. Enter the path to the **modified** `.docx` file
4. Press Enter to accept the default output path (or type your own)
5. The tool generates a redlined `.docx` file with inline markup and an "Exhibit A — Change Report" at the end

### Command-line mode

You can also run it non-interactively:

```bash
.venv/bin/python -m docx_redline.cli original.docx modified.docx -o output.docx
```

---

## Folder Structure

```
docx_redline/
├── docx_redline/          Python package source
├── setup.sh               One-time installer
├── install.sh             Bootstrap script (for curl install)
├── run.command            Double-click to launch
├── pyproject.toml         Package metadata
└── README.md
```

The `.venv/` folder is created automatically during setup and contains all dependencies.

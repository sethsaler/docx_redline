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

The installer tries **git clone** first; if that fails (VPN, firewall, or corporate network blocking Git), it **automatically falls back** to downloading a **ZIP** from GitHub over HTTPS.

Once installed, launch it anytime by double-clicking **`run.command`** in the `docx_redline` folder.

---

## Troubleshooting

**“The graphical file picker requires Tk (tkinter)”** — Your Python was built without Tk. Install a Python that includes Tcl/Tk (the official macOS installer from [python.org](https://www.python.org/downloads/) does), or on Linux install `python3-tk`, then delete the project’s `.venv` folder and run `bash setup.sh` again.

**Git or curl install errors** — If `git clone` fails, use the latest `install.sh` from this repo (it falls back to ZIP). You can also download the ZIP manually: open `https://github.com/sethsaler/docx_redline/archive/refs/heads/main.zip`, unzip it, rename the folder to `docx_redline` on your Desktop, and run `bash setup.sh` inside it.

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

**Requirements:** macOS with Python 3.9 or later. If Python is not installed, `setup.sh` will download and install it automatically. The double-click UI uses **Tk** (included with python.org Python on macOS). On Linux, install `python3-tk` if the window does not appear and the tool falls back to terminal prompts.

---

## How to Use

1. Launch the tool (double-click `run.command` or run `./run.command`)
2. Click **Browse…** to pick the **original** and **modified** `.docx` files (no need to copy paths into a terminal)
3. Optionally set the output path with **Browse…**, or leave it blank to use the default next to **Generate redline**
4. Click **Generate redline**
5. The tool writes a redlined `.docx` with inline markup and an "Exhibit A — Change Report" at the end

**Terminal prompts (no file picker):** If you prefer typing paths, run:

```bash
.venv/bin/python -m docx_redline.cli_interactive
```

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

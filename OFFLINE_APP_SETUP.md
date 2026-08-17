# Leave Document Generator — Offline Desktop App

This turns the tool into a double-clickable window with a "Generate" button —
no terminal, no typing commands. Everything runs on your machine; no internet
is used while generating documents.

## One-time setup (needs internet once, then it's fully offline)

### 1. Install Python (skip if you already have it)
Download from https://www.python.org/downloads/ (3.10 or newer).
During install, **tick "Add python.exe to PATH"**. The "tcl/tk" component
(needed for the app window) is included by default — don't uncheck it.

### 2. Install the required packages
Open Command Prompt and run:
```
pip install openpyxl python-docx pypdf reportlab
```

### 3. Install LibreOffice
Download and install from https://www.libreoffice.org/download/download/
(this is what converts the filled forms to PDF). Any recent version works.

That's it — steps 1–3 only need to be done once. After that, no internet
connection is required to generate documents.

## Running the app

**Windows:** double-click `Run Leave Generator.bat` in this folder.

**Mac/Linux:** open a terminal in this folder and run:
```
python3 leave_generator_app.py
```

A window opens with:
1. **Browse…** next to "Input Excel file" → pick your filled employee sheet
2. **Browse…** next to "Output folder" → pick where documents should be saved
3. **Generate Leave Documents** button → runs everything and shows progress in the log
4. **Open Output Folder** → jumps straight to the results when done

Keep reusing the same input file (add/edit employee rows) and just click
Generate again each time.

## Turning it into a true standalone .exe (no Python needed to run it)

If you want to hand this to someone without them installing Python at all,
package it once with PyInstaller **on a Windows machine**:

```
pip install pyinstaller
pyinstaller --noconsole --onefile --name "LeaveDocumentGenerator" ^
    --add-data "templates;templates" ^
    --add-data "scripts;scripts" ^
    leave_generator_app.py
```

This produces `dist/LeaveDocumentGenerator.exe` — a single file you can copy
anywhere and double-click. (LibreOffice still needs to be installed on that
machine, since the .exe calls it to produce PDFs.)

## Troubleshooting

- **"No module named tkinter"** (Linux only) — install it with
  `sudo apt install python3-tk`, then try again.
- **Nothing happens when converting to PDF** — check LibreOffice is
  installed and that `soffice` runs from a terminal. If LibreOffice was
  installed to a non-default location, add its folder to your system PATH.
- **"Invalid Sponsor ID format" for everyone** — double check the `ID
  NUMBER` column values start with `7`, `C`, or `EM-`.

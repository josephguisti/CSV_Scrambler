## How to Turn a Microsoft List Export into an Email (Windows / PC Guide)

This guide is for non-technical staff. It explains how to take a **Critical Incident Microsoft List export (CSV)** and turn it into an **email-friendly text file**.

The output looks like this (one incident per row in the CSV):

Status:
New

Date Submitted:
3/3/2026

----------

### What this does (in plain English)
- Reads your exported CSV
- For each incident (each row), it prints each column as:
  - the **column name**
  - then the **value** on the next line
- Adds `----------` between incidents
- Correctly handles fields that contain multiple lines (like long incident descriptions)

---

## What you need
- The folder that contains:
  - `format_critical_incidents_for_email.py`
  - your exported CSV file (from Microsoft Lists)
- **Python** installed on your computer

---

## Step 1: Export the Microsoft List to CSV
In Microsoft Lists, export/download the list as a **CSV** and save it somewhere you can find it.

Recommended: save it into the same folder that contains the Python scripts (or copy it there).

---

## Step 2: Find and open “Terminal” on Windows
Windows has a few “Terminal” apps. Any of these work:
- **Command Prompt**
- **Windows PowerShell**
- **Windows Terminal** (newer app that can open PowerShell/Command Prompt tabs)

How to open:
- Click the **Start** menu
- Type **PowerShell** (or **Command Prompt**)
- Click the app to open it

You will see a window with a blinking cursor where you can type commands.

---

## Step 3: Confirm Python is installed
In the Terminal window, type this and press **Enter**:

```bash
python --version
```

If you see something like `Python 3.x.x`, you’re good.

If you see an error like “python is not recognized”:
- Install Python 3 from the Microsoft Store (search **Python**), **or**
- Ask IT / a technical staff member to install Python 3

Then re-run:

```bash
python --version
```

---

## Step 4: Point Terminal at the right folder (VERY IMPORTANT)
Most “it doesn’t work” issues happen because Terminal is in the wrong folder.

### 4A) Find the folder path
1. Open **File Explorer**
2. Navigate to the folder that contains `format_critical_incidents_for_email.py`
3. Click the address bar at the top (where the folder location shows)
4. Copy the full path (it will look like something starting with `C:\Users\...`)

### 4B) Change Terminal to that folder
In Terminal, type `cd` (change directory), then paste your folder path in quotes:

```bash
cd "PASTE-YOUR-FOLDER-PATH-HERE"
```

Example (yours may be different):

```bash
cd "C:\Users\YourName\Documents\Python\CSV scrambler and exporter"
```

### 4C) Double-check you’re in the right place
Run:

```bash
dir
```

You should see a list of files including:
- `format_critical_incidents_for_email.py`
- your CSV file

If you do **not** see those files, you are in the wrong folder:
- Repeat Step 4A and 4B
- Then run `dir` again

---

## Step 5: Close the CSV in Excel first
If your CSV is open in Excel, Windows may block the script from reading it (or block writing the output).

Close Excel (or at least close the CSV) before running the command.

---

## Step 6: Run the formatter (creates the email-ready text file)
Type this and press **Enter**:

```bash
python format_critical_incidents_for_email.py
```

The script will:
- Ask you to **select the CSV file** (a file picker window opens)
- Ask you where to **save the email-ready .txt file** (a “Save As” window opens)

Recommended: when you save the output, name it something like:
- `Critical Incident Email Output.txt`

Example:

```bash
python format_critical_incidents_for_email.py
```

When it finishes, you’ll see something like:
- `Format complete.`
- `Output: ...\Critical Incident Email Output.txt`

### If the file picker windows do NOT appear (backup method)
Sometimes pop-up windows are blocked by your computer settings. If that happens, you can run it by typing the file names yourself:

```bash
python format_critical_incidents_for_email.py --input "YOUR_EXPORT.csv" --output "Critical Incident Email Output.txt" --skip-empty-fields
```

---

## Step 7: Send the email
1. Find the output file: `Critical Incident Email Output.txt`
   - It will be in the same folder you `cd`’d into (unless you used a different output path).
2. Open the file (double-click it)
3. Press **Ctrl + A** (select all)
4. Press **Ctrl + C** (copy)
5. Paste it into your email

---

## Optional options (only if you need them)

### Include extra columns (if the CSV has new columns)
If Microsoft Lists adds extra columns over time, you can include them:

```bash
python format_critical_incidents_for_email.py --input "YOUR_EXPORT.csv" --output "Critical Incident Email Output.txt" --skip-empty-fields --include-extras
```

### Stop with an error if the CSV is missing expected columns
Use this if you want a strict “fail fast” check:

```bash
python format_critical_incidents_for_email.py --input "YOUR_EXPORT.csv" --output "Critical Incident Email Output.txt" --strict-schema
```

---

## Troubleshooting (common SNAFUs)

### “python is not recognized…”
- Python is not installed, or Windows can’t find it.
- Install Python 3 (Microsoft Store), or ask IT.

### “File not found” / “No such file”
Most likely you’re in the wrong folder.

Fix:
- Do Step 4 again (`cd ...`)
- Run `dir`
- Confirm you see:
  - `format_critical_incidents_for_email.py`
  - your CSV file name exactly as you typed it

### It runs but the output file doesn’t appear
- Run `dir` and look for `Critical Incident Email Output.txt`
- Make sure you closed Excel
- Make sure you have permission to write to that folder

### OneDrive / permission / syncing issues
If this folder is inside OneDrive, sometimes files can be locked or half-synced.

Fix:
- Copy the CSV into a local folder like `C:\Temp\`
- Run the script from there

Example:

```bash
cd "C:\Temp"
python "C:\PATH\TO\format_critical_incidents_for_email.py" --input "YOUR_EXPORT.csv" --output "Critical Incident Email Output.txt" --skip-empty-fields
```

### The formatting looks “weird” (extra spaces, blank lines, etc.)
- Try using `--skip-empty-fields` (recommended)
- If the CSV has brand-new columns, try `--include-extras`


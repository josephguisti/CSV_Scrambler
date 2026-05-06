# How to Scramble a CSV (Windows / PC Guide)

This guide is for non-technical staff. It explains how to run a Python script that **scrambles names, phone numbers, and other PII** in a CSV before you upload it into an AI-enabled tool.

## What you need

- The folder containing:
  - `scramble_csv.py`
  - your input CSV (the file you want to scrub)
- **Python** installed on your computer

## Step 1: Find and open “Terminal” on Windows

Windows has a few “Terminal” apps. Any of these work:

- **Command Prompt**
- **Windows PowerShell**
- **Windows Terminal** (a newer app that can open PowerShell/Command Prompt tabs)

How to open:

- Click the **Start** menu
- Type **PowerShell** (or **Command Prompt**)
- Click the app to open it

You will see a window with a blinking cursor where you can type commands.

## Step 2: Confirm Python is installed

In the Terminal window, type this and press **Enter**:

```bash
python --version
```

If you see something like `Python 3.x.x`, you’re good.

If you see an error like “python is not recognized”:

- Install Python 3 from the Microsoft Store (search **Python** in the Store), **or**
- Ask IT / a technical staff member to install Python 3

Then re-run:

```bash
python --version
```

## Step 3: Go to the folder that contains the script and your CSV

You must “point” the Terminal to the correct folder (directory).

### 3A) Find the folder path

1. Open **File Explorer**
2. Navigate to the folder that contains `scramble_csv.py`
3. Click the address bar at the top (where the folder location shows)
4. Copy the full path (it will look like something starting with `C:\Users\...`)

### 3B) Change the Terminal to that folder

In Terminal, type `cd` (change directory), then paste your folder path in quotes:

```bash
cd "PASTE-YOUR-FOLDER-PATH-HERE"
```

Example (yours may be different):

```bash
cd "C:\Users\YourName\Documents\Python\CSV scrambler and exporter"
```

### 3C) Check you’re in the right place

Run:

```bash
dir
```

You should see a list of files including:
- `scramble_csv.py`
- your CSV file

If you do **not** see `scramble_csv.py`, you are in the wrong folder—repeat Step 3 with the correct folder path.

## Step 4: Run the scrambler

You’ll run the script with:
- an **input** CSV (original)
- an **output** CSV (scrubbed copy)

### Option 1 (easiest): Use the file picker windows

In the same folder as `scramble_csv.py` (see Step 3), run:

```bash
python scramble_csv.py
```

Two windows will appear:
- A chooser window to select the **CSV file to scramble**
- A “Save as…” window to choose the **name and location** of the scrambled CSV

### 4A) IMPORTANT: Close the CSV in Excel first

If your CSV is open in Excel, Windows may block the script from saving the new output file.

Close Excel (or at least close the CSV) before running the script.

### Option 2: Run the command by typing filenames (advanced)

In Terminal, type:

```bash
python scramble_csv.py --input "YOUR_INPUT_FILE.csv" --output "YOUR_OUTPUT_FILE.csv"
```

Example:

```bash
python scramble_csv.py --input "Critical Incident Report Pre-Scramble.csv" --output "Critical Incident Report Scrambled.csv"
```

When it finishes, it prints a message that includes:
- the **output file path**
- a **Seed**

## Step 5: (Optional) Keep the same scramble every time (recommended)

If you use the same **seed**, you get the same scrambled results every time for the same input.

Run:

```bash
python scramble_csv.py --input "YOUR_INPUT_FILE.csv" --output "YOUR_OUTPUT_FILE.csv" --seed "PASTE-SEED-HERE"
```

Tip: Save the seed somewhere safe (NOT in the scrubbed CSV).

## Step 6: Find the scrubbed CSV

The scrubbed file is wherever you told the script to write `--output`.

- If you used a simple name like `--output "Scrubbed.csv"`, it will be in the same folder you `cd`’d into.
- You can also search for the output name in File Explorer.

## Step 7: (Optional) Confirm minor/TAY/adult counts stayed the same

If your CSV has a DOB column (column header containing “age”, “birth”, or “dob”), you can run a check:

```bash
python age_bucket_report.py --input "YOUR_INPUT_FILE.csv" --compare "YOUR_OUTPUT_FILE.csv"
```

Example:

```bash
python age_bucket_report.py --input "Critical Incident Report Pre-Scramble.csv" --compare "Critical Incident Report Scrambled.csv"
```

You should see **no differences** in the “Diff” section.

## Troubleshooting

### “python is not recognized…”
- Python is not installed, or Windows can’t find it.
- Install Python 3 (Microsoft Store), or ask IT.

### “File not found” / “No such file”
- You are probably in the wrong folder.
- Repeat Step 3 and confirm with `dir` that your CSV and `scramble_csv.py` are listed.

### Permission errors / OneDrive issues
- If the folder is in OneDrive, it may be syncing or locked.
- Try copying the CSV into a local folder like `C:\Temp\` and run the script there.

### Output file won’t save
- Close the CSV in Excel.
- Make sure you have permission to write to that folder.

### The output looks “too scrambled”
- That’s expected: letters and digits are scrambled to remove PII.
- “Program” and “Project” columns (except those containing “Date” in the header) are intentionally left unchanged for error-checking.

## Bonus: Turn the Microsoft List export into an email (Critical Incident format)
If the automator breaks, you may need to email Critical Incident entries to directors/managers in a readable format.
This script turns each CSV row into a clean block like:

- `COLUMN HEADER:`
- the value on the next line

…with `----------` between incidents.

### Step A: Export the Microsoft List to CSV
In Microsoft Lists, export/download the list as a **CSV** and save it into the same folder as the scripts (or copy it there).

### Step B: Close the CSV in Excel first
If the CSV is open in Excel, Windows may block the script from writing the output file.
Close Excel (or at least close the CSV) before running the command.

### Step C: Run the formatter
From the same folder (see Step 3 earlier), run:

```bash
python format_critical_incidents_for_email.py --input "YOUR_EXPORT.csv" --output "Critical Incident Email Output.txt"
```

Then open `Critical Incident Email Output.txt`, copy it, and paste into your email.

Optional (recommended): skip blank fields for cleaner emails:

```bash
python format_critical_incidents_for_email.py --input "YOUR_EXPORT.csv" --output "Critical Incident Email Output.txt" --skip-empty-fields
```

## Appendix: Business rules (what the script does)

This section lists the exact rules baked into `scramble_csv.py`.

### What gets scrambled (by default)

- **Letters**: All A–Z characters in the CSV are scrambled.
  - **Case is preserved** (uppercase stays uppercase, lowercase stays lowercase).
  - **Length is preserved** (the scrambled text has the same number of characters).
- **Digits**: All 0–9 digits are scrambled.
  - This includes digits inside phone numbers, addresses, IDs, etc.
- **Punctuation and spacing**: preserved exactly.
  - Examples of things that are preserved: spaces, commas, periods, parentheses, newlines, and hyphens.
  - Hyphenated names/words keep the hyphen in the same place (the letters on each side are scrambled, but the hyphen remains).

### Special rule: Program/Project columns stay stable

- If a **column header contains `Program` or `Project`** (any capitalization)
  - and the header does **NOT** contain the word `Date`
  - then **that entire column is left unchanged** (no scrambling at all).

Examples:
- `Programs Involved` → unchanged
- `Project Assigned Staff` → unchanged
- `Project Start Date` → scrambled (because it contains `Date`)

### Special rule: Dates are scrambled, but stay in the same month + year

When the script finds a date inside a cell, it changes only the **day** number while keeping:
- the **month** the same
- the **year** the same

Supported date patterns:
- `M/D/YYYY` (example: `3/3/2026`)
- `YYYY-MM-DD` (example: `2026-03-03`)
  - Works even if the date is immediately followed by time text (example: `2026-03-0310:40AM`)

### Special rule: Time-of-day is NOT scrambled

You requested “just the date,” so the script preserves time-of-day substrings like:
- `10:40AM`, `3:00pm`
- `19:06`, `22:12`

### Special rule: DOB / birth / age columns keep minor/TAY/adult classification (as-of today)

For columns whose **header contains** any of these words (any capitalization):
- `age`
- `birth`
- `dob`

…the script scrambles the **day of the DOB** (still keeping DOB month+year the same) in a way that preserves the person’s category **as-of today’s date**:
- **Minor**: under 18
- **TAY**: 18–24 inclusive
- **Adult**: 25+

### Repeatability (important)

- If you provide `--seed "something"`, the scrambling is **deterministic** (you’ll get the same scrambled output again for the same input).
- If you do **not** provide a seed, the script generates one and **prints it** at the end so you can copy/save it.

### Known limitations (good to know)

- The script only recognizes the date formats listed above. Other formats (like `03-03-2026` or `March 3, 2026`) will be treated like normal text and scrambled.
- Everything is processed “as text” unless it matches the special rules above. For example, numeric columns like `Days Since Submission` will have their digits scrambled unless you decide to exempt that column in a future update.


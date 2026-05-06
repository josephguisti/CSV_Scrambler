# How to Use the CSV Scrambler App (Windows / PC Guide)

This guide is for non-technical staff. It explains how to use the **CSV Scrambler app** (a double-click `.exe`) to remove PII before uploading data into AI-enabled tools.

## What you need

- The app file: **`CSV Scrambler.exe`**
- The CSV you want to scrub

## Step 1: Open the app

1. Find `CSV Scrambler.exe` in File Explorer
2. Double-click it

If Windows shows a security warning (“Windows protected your PC”):
- Click **More info**
- Click **Run anyway**

## Step 2: Choose the file to scramble (CSV or Excel)

Click **Choose File…** and pick the file you want to scrub.

After you choose the file, you’ll see a list of all column headers with dropdowns.

### Excel requirements (important)

If you use an Excel file (`.xlsx`):
- **Headers are in row 1 on every worksheet tab** you want scrambled
- The app lists **all column names** from **all** tabs together; the rules you pick apply to every tab that has that column
- **Shuffle values** is done **inside each sheet** (not mixed across tabs)
- Chart-only tabs are **not** included (only regular worksheet grids)
- Avoid merged cells / tables (filters are OK)

The app writes a **new Excel file with values only** (it does not preserve formulas/formatting), and any filters will be removed in the output.

## Step 3: Choose where the scrambled CSV will be saved

Click **Choose Output…** and pick the **name and location** for the scrambled file.

Tip: save it somewhere easy to find (like your Desktop) if you’re going to upload it right away.

## Step 4: Review the column rules (dropdowns)

Each column has a dropdown showing how it will be handled:

- **Full scramble**: scrambles letters and digits in that column (including any dates/times)
- **Date scramble**: scrambles letters/digits, but:
  - dates keep the **same month + year** (only the day changes)
  - DOB-like columns (headers containing `age`, `birth`, or `dob`) keep the same bucket **as-of today** (minor / TAY / adult)
  - time-of-day like `10:40AM` is preserved
- **Shuffle values**: keeps the set of values in the column, but reassigns them across rows so **each non-blank row gets a different value than it had**
  - blanks stay blank
  - if all rows have the same value, the app will show an error
- **No scramble**: leaves the column unchanged
- **Delete data**: blanks out the column values (the column stays in the CSV)

## Step 5 (Optional): Create Keyed Normal Version

If you check **Create Keyed Normal Version**, the app will produce **two files**:

- the **scrambled CSV**
- a **keyed original CSV** (unchanged original values), both containing a `Row ID` column

This can be used to join the two files later if you need to recover the original values.

## Step 6: Run it

Click **Run Scramble**.

When it finishes, the app will show:
- the path to the scrambled file
- and (if enabled) the path to the keyed original file

## Troubleshooting

### The output file won’t save
- Close the CSV in Excel (Excel can lock the file).
- Make sure you have permission to write to that folder.

### Shuffle values error
- On Excel, the message usually names the **sheet** and **column**.
- This usually means that column has too little variety **on that sheet** to guarantee every row changes (for example, every row has the same value, or only one row has data).


# How to Use the CSV / Excel Scrambler App (Windows / PC Guide)

This guide is for non-technical staff. It explains how to use the **CSV Scrambler app** (a double-click `.exe`) to remove PII from CSV or Excel files before uploading data into AI-enabled tools.

## What you need

- The app file: **`CSV Scrambler.exe`**
- The CSV or Excel file you want to scrub

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

If you use an Excel file (`.xlsx`) with **more than one worksheet**, when you click **Run Scramble** the app opens a **sheet picker**: choose which tabs to include in the export (you can uncheck tabs you do not need). If conditional formatting on an included sheet references an excluded sheet, you may see a **warning** listing those relationships—you can still proceed.

If you use an Excel file with **only one worksheet**, there is no sheet picker.

If you use an Excel file (`.xlsx`):
- **Headers** are detected automatically on each tab (often row 1; see the service manual if your file has title rows above the headers).
- When **all** tabs are loaded for rules, the app lists column names from every tab together (keys look like `ColumnName (SheetName)`); after you pick a subset of sheets at Run time, the rule list shows **only** columns for those tabs.
- **Shuffle values** is done **inside each sheet** (not mixed across tabs)
- Chart-only tabs are **not** included (only regular worksheet grids)
- Avoid merged cells / tables where possible (filters are OK)

The app writes a **new Excel file** with scrambled values; static cell coloring and conditional formatting are copied where possible (see the service manual). Formulas are not preserved.

### Password-protected Excel files

If you open a password-protected `.xlsx` file, the app will automatically prompt you for the password. Enter it and click OK. The password is used only to open the file — the output file will **not** be password-protected.

## Step 3: Choose where the scrambled output will be saved

Click **Choose Output…** and pick the **name and location** for the scrambled file.

Tip: save it somewhere easy to find (like your Desktop) if you’re going to upload it right away.

For Excel files:
- Use `.xlsx` output when exporting more than one worksheet.
- `.csv` output is allowed only when exporting one worksheet.

## Step 4: Review the column rules (dropdowns)

By default, **every column starts as No scramble**—change the dropdown only for columns you want to transform.

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

- the **scrambled output**
- a **keyed original** file (unchanged original values), both containing a `Row ID` column

This can be used to join the two files later if you need to recover the original values.

## Step 6: Run it

Click **Run Scramble**.

When it finishes, the app will show:
- the path to the scrambled file
- and (if enabled) the path to the keyed original file

## Troubleshooting

### The output file won’t save
- Close the output file in Excel (Excel can lock the file).
- Make sure you have permission to write to that folder.
- If the app says the output file is open, close the named file and run the scramble again.

### "Wrong password or corrupt file"
- The password you entered was incorrect, or the file is not a valid encrypted Excel file.
- Re-select the file and enter the correct password.

### Shuffle values error
- On Excel, the message usually names the **sheet** and **column**.
- This usually means that column has too little variety **on that sheet** to guarantee every row changes (for example, every row has the same value, or only one row has data).


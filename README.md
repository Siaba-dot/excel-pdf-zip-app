# excel-zip-app
📄 Excel Processing Streamlit App

✨ Features

Upload a folder as ZIP (with subfolders)

Automatically update C5 cell with the last date of the current month

Replace month names in A9 cell with the current month (Lithuanian)

Rename files to match the YYYY_MM pattern

Download all updated Excel files as a new ZIP

📅 Why date replacement matters

Updating C5 cell to the last day of the current month ensures the document always reflects the correct reporting date

This change also recalculates the number of days in that month (28, 30, or 31)

Dependent formulas are updated automatically → resulting in correct totals and sums

A9 cell is updated with the correct month name in Lithuanian, keeping the documents consistent with the actual reporting period

🚀 Usage

Open the deployed app (Streamlit Cloud link).

Click "Upload ZIP" and select a folder (compressed as ZIP) containing your Excel files.

The app will start processing:

Updating C5 and A9 cells

Renaming files if necessary

Showing live progress in the log window

When processing is finished, click "Download updated Excel (.zip)" to get your results.

📦 Input requirements

Files must be .xlsx or .xlsm.

You can include subfolders inside the ZIP.

Non-Excel files are ignored.

🛠️ Tech stack

Streamlit

OpenPyXL

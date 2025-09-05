# excel-zip-app
📄 Excel Processing Streamlit App

✨ Features

Upload a folder as ZIP (with subfolders)

Automatically update C5 cell with the last date of the current month

Replace month names in A9 cell with the current month (Lithuanian)

Rename files to match the YYYY_MM pattern

Download all updated Excel files as a new ZIP

📅 Why date replacement matters

C5 cell is automatically updated to the last day of the current month (e.g. 2025-09-30 in September, 2025-10-31 in October)

This ensures that each month the data refreshes according to the current reporting period

Changing the date also triggers recalculation of the number of days in that month (28, 30, or 31)

Dependent formulas automatically update → producing correct totals and sums

A9 cell is updated with the Lithuanian month name, so documents always match the actual month

🚀 Usage

Open the deployed app (Streamlit Cloud link).

Click "Upload ZIP" and select a folder (compressed as ZIP) containing your Excel files.

The app will start processing:

Updating C5 and A9 cells

Renaming files 

Showing live progress in the log window

When processing is finished, click "Download updated Excel (.zip)" to get your results.

📦 Input requirements

Files must be .xlsx or .xlsm.

You can include subfolders inside the ZIP.

Non-Excel files are ignored.

🛠️ Tech stack

Streamlit

OpenPyXL

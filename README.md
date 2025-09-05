# excel-zip-app
📄 Excel Processing Streamlit App

Streamlit app for processing Excel files:

Upload a folder as ZIP (with subfolders)

Automatically update C5 cell with the last date of the current month

Replace month names in A9 cell with the current month (Lithuanian)

Rename files to match the YYYY_MM pattern

Download all updated Excel files as a new ZIP

✨ Features

Upload a folder as ZIP (with subfolders)

Automatically update C5 cell with the last date of the current month

Replace month names in A9 cell with the current month (Lithuanian)

Rename files to match the YYYY_MM pattern

Download all updated Excel files as a new ZIP

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

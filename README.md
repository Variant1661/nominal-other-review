# Nominal OTHER Review Streamlit App

Browser-based version of the Nominal OTHER Review tool.

## Files

- `streamlit_app.py` - Streamlit web app entrypoint
- `nominal_logic.py` - payroll parsing, staff matching, validation, and Excel output
- `requirements.txt` - packages Streamlit Cloud installs

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload these three files to the repository root:
   - `streamlit_app.py`
   - `nominal_logic.py`
   - `requirements.txt`
3. In Streamlit Community Cloud, click **Create app**.
4. Choose the GitHub repository.
5. Set branch to `main`.
6. Set main file path to `streamlit_app.py`.
7. Click **Deploy**.

## Data note

This app processes uploaded files in memory/session temp storage and returns an Excel report for download. For payroll data, deploy inside an approved company environment or restrict access to approved users.

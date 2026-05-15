# Nominal OTHER Review Streamlit App

Browser-based version of the Nominal OTHER Review tool.

## Files

- `streamlit_app.py` - Streamlit web app entrypoint
- `nominal_logic.py` - payroll parsing, staff matching, validation, and Excel output
- `requirements.txt` - packages Streamlit Cloud installs
- `assets/company_logo.gif` - Apex Care Homes logo used in the app header/sidebar

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload these three files to the repository root:
   - `streamlit_app.py`
   - `nominal_logic.py`
   - `requirements.txt`
3. Upload the `assets` folder to the repository root as well.
4. In Streamlit Community Cloud, click **Create app**.
5. Choose the GitHub repository.
6. Set branch to `main`.
7. Set main file path to `streamlit_app.py`.
8. Click **Deploy**.

## Missing nominal workflow

The missing nominal screen is staff-level. If you enter `504` once for a staff member, the app applies `504` to every blank OTHER/NIC row for that same staff member.

The Dashboard tab updates live as corrections are saved.

## Data note

This app processes uploaded files in memory/session temp storage and returns an Excel report for download. For payroll data, deploy inside an approved company environment or restrict access to approved users.

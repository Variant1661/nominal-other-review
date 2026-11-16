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

## Behaviour notes

- A payroll row with no Code is never given the nominal of the row above; it appears in the Missing Nominals queue.
- Amounts are signed (debit minus credit), so reversals reduce transfer totals.
- Correction keys include the source file and row number, so identical lines are corrected independently.
- Staff base homes that are blank or on an unmapped nominal prefix block report generation.
- A generated report is only offered for download while it still matches the current inputs.

## Tests

```
pip install -r requirements.txt -r requirements-dev.txt
python -m pytest
```
Tests use fictional data only.

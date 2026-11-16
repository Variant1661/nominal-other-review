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

## Department costing page

`pages/1_Department_Costing.py` (logic in `department_costing.py`) re-splits cost between homes after wages are posted.

- **Pool rules** - a shared department or role (e.g. kitchen suffix `05`, paying homes CRE+PET, 50/50). All matching lines are pooled, then split by the shares.
- **Person rules** - one person across homes by percentage or days (2/2/1 days = 40/40/20%). These override pool rules.
- Shares are normalised, so percentages and days both work. Rules can be saved/loaded as JSON.
- Output: net journal (credit nominal -> debit nominal per cost type), allocation detail, staff included, and warnings. Lines with no nominal code are listed, never allocated.
- Payroll names ("Surname Forename X") are matched to the staff master by words; an ambiguous or unmatched name is flagged, not guessed.

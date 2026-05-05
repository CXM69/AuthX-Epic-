# AuthX Epic Targeting Engine

A Streamlit app for ranking health system accounts from:

- `EPIC Organization list.xlsx`
- `Health Systems by EHR.xlsx`

The app reads every worksheet from both workbooks, normalizes account names, deduplicates accounts, flags Epic and Imprivata signals, scores accounts from 0 to 100, assigns targeting tiers, and exports a ranked Excel workbook.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

Open the local Streamlit URL shown in the terminal, upload both Excel files, then download the ranked output.

## Scoring

Accounts are scored with these signals:

- Epic customer: 35 points
- New Epic system signal: 30 points
- Imprivata customer signal: 20 points
- Multiple source files: up to 10 points
- Multiple source tabs: up to 5 points
- State or region present: 3 points

Scores are capped at 100. Accounts with exclude signals receive a score of 0 and are assigned `Exclude`.

## Tiers

- `Tier 1`: score of 80 or higher
- `Tier 2`: score of 60 to 79
- `Tier 3`: score of 40 to 59
- `Hold`: score below 40
- `Exclude`: account has an exclusion signal

## Expected Columns

The app automatically looks for common account-name columns such as:

- `Account Name`
- `Organization`
- `Health System`
- `Customer`
- `Name`

It also detects common EHR, status, and state columns when present. If an uploaded workbook uses different names, update the column candidate lists in `scoring.py`.

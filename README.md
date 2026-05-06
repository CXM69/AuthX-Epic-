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

Open the local Streamlit URL shown in the terminal, upload both Excel files, click `Generate ranked accounts`, then download the ranked output.

## Deploy On Render

This repo includes `render.yaml` for Render Blueprint deploys.

Use these settings if creating the service manually:

- Runtime: Python
- Build command: `pip install -r requirements.txt`
- Start command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0 --server.headless true`

## Scoring

Accounts are scored with these signals:

- Epic customer: 35 points
- New Epic system signal: 30 points
- Imprivata customer signal: 20 points
- Multiple source files: up to 10 points
- Multiple source tabs: up to 5 points
- State or region present: 3 points

Fit scores are capped at 100. Accounts with exclude signals receive a fit score of 0 and keep `Exclude Flag` marked true.

## Tiers

- `1`: highest priority, fit score of 60 or higher
- `2`: secondary priority, fit score of 40 to 59
- `Hold`: lower priority, fit score below 40

The `Exclude Flag` column separately marks accounts with exclusion signals.

## Expected Columns

The app automatically looks for common account-name columns such as:

- `Account Name`
- `Organization`
- `Health System`
- `Customer`
- `Name`

It also detects common EHR, status, and state columns when present. If an uploaded workbook uses different names, update the column candidate lists in `scoring.py`.

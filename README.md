# AuthX Epic Front-Door Targeting Engine

A simple Streamlit MVP that helps sales teams rank healthcare accounts for AuthX as the Epic front-door authentication and access layer.

The app reads every worksheet from one or more uploaded Excel workbooks. The original source files are supported examples, not hard requirements:

- `EPIC Organization list.xlsx`
- `Health Systems by EHR.xlsx`

You can also upload other account, EHR, Epic, Imprivata, security, workflow, or target-account spreadsheets. The app normalizes account names, deduplicates accounts across every uploaded tab, identifies Epic and competitive signals, scores each account from 0 to 100, recommends the best first buyer role, and exports a ranked Excel workbook.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run Locally

```bash
streamlit run app.py
```

Open the local Streamlit URL shown in the terminal, upload one or more Excel workbooks, click `Generate ranked list`, then download the ranked output.

## Scoring Model

- Epic Status: 25 points
- Timing Trigger: 20 points
- Enterprise Value: 15 points
- Imprivata / Competitive Signal: 15 points
- Front-Door Workflow Fit: 10 points
- Security / Compliance Urgency: 10 points
- Buyer Role Fit: 5 points

## Tiers

- Tier 1: 85-100
- Tier 2: 70-84
- Tier 3: 50-69
- Hold: 30-49
- Exclude: 0-29

## Guardrail

AuthX is positioned as the authentication and access layer around Epic: workstation access, SSO, MFA, badge tap, biometrics, passkeys, VDI/thin-client access, and step-up re-authentication.

Do not position AuthX as replacing Epic role assignment, identity governance, provisioning, deprovisioning, or internal Epic authorization.

## Excel Export

The downloaded workbook includes:

- `All Scored Accounts`
- `Tier 1 Targets`
- `Tier 2 Targets`
- `Role Strategy`
- `Scoring Summary`

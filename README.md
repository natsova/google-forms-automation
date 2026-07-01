# Google Forms → Local Documents Automation

Automatically transforms Google Form responses into personalised Word documents,
saved locally - no subscriptions required.

**Pipeline:** Google Form → Google Sheets → Python → Local filesystem

Built and documented by [Natalie Sova](https://natsova.com).

## Series Overview

This repository follows the YouTube series that takes a working demo script
through five layers of production hardening.

| Part | Topic | Description |
|------|-------|-------------|
| 1 | Demo | Working proof-of-concept - Google Sheets to local .docx |
| 2 | Infrastructure | Preflight checks, environment validation, directory provisioning |
| 3 | Security | Credential handling, input sanitisation, path confinement |
| 4 | Idempotency | Deduplication markers, atomic file claiming |
| 5 | Reliability | Atomic writes, orphan cleanup, exception rollback |
| 6 | Observability | Structured logging, audit trail, run summaries |

Each folder contains the complete working script at that stage - drop into any
part and run it independently.

## Setup

**1. Clone and install**

Run the following from the root of the repository:

```bash
git clone https://github.com/natsova/google-forms-automation.git
cd google-forms-automation
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

**2. Add credentials**

Place your service account key at `keys/service-account.json` (and add `keys/` to `.gitignore`)

**3. Configure `.env`**

```env
SPREADSHEET_ID=your_google_sheet_id
GOOGLE_SHEETS_KEY=keys/service-account.json
```

(and add `.env` to `.gitignore`)

**4. Navigate to the part you want to run**

```bash
cd part2-infrastructure
python3 main.py
```

**Note:** Part 1 uses `generate.py`

## How it works

1. Loads config from `.env`
2. Authenticates with Google Sheets via service account
3. Fetches all rows from the linked Google Sheet
4. Generates one `.docx` per row, saved under `submissions/`

## Output

```
submissions/
  customers/
    jane_doe/
      2026-05-06_120102_123456_submission.docx
    john_smith/
      2026-05-06_120245_654321_submission.docx
```

## Requirements

- Python 3.9+
- A Google Cloud service account with Sheets API enabled
- A Google Form linked to a Google Sheet

## Links

- YouTube series: [Hardened Google Forms Automation](https://www.youtube.com/watch?v=27FuXXj7R5k&list=PL5SSMCIUhqDij99pd791uzobafUR3FGgF)
- Engineering write-up: [natsova.com](https://natsova.com/google-forms-automation)

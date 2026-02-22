# Usage: Any Customer

This guide runs the full flow for any customer and engagement.

## 1) Prerequisites

- Run from repository root:
  - `/Users/pravin.varma/Documents/Demo/pm-automation`
- Python 3 and Node.js installed
- Dependencies installed:

```bash
npm install
```

- MCP/tools authenticated (if pulling live): Slack, Google Drive, Salesforce, Glean/Genie

## 2) Variables

Set these for each run:

- `CUSTOMER_NAME` (example: `ANZ`)
- `ENGAGEMENT_NAME` (example: `Professional Services`)
- `PERIOD_LABEL` (example: `2026-02-02 to 2026-02-06`)

## 3) Option A: Skills export workflow (recommended)

1. Use Claude Code to fetch real data and write:
   - `inputs/skills_exports/slack.json`
   - `inputs/skills_exports/salesforce.json`
   - `inputs/skills_exports/gdrive.json`
   - `inputs/skills_exports/glean.json`

2. Build reports:

```bash
python3 pm_report_automation.py \
  --project-root . \
  --ingest-skills-exports \
  --skills-dir "inputs/skills_exports" \
  --period-label "2026-02-02 to 2026-02-06" \
  --customer-name "ANZ" \
  --engagement-name "Professional Services"
```

3. Build PPT:

```bash
npm run report:ppt -- \
  --period "2026-02-02 to 2026-02-06" \
  --customer "ANZ" \
  --engagement "Professional Services"
```

## 4) Option B: Live MCP fetch workflow

```bash
python3 pm_report_automation.py \
  --project-root . \
  --fetch-live \
  --lookback-days 7 \
  --mcp-config "~/.claude.json" \
  --period-label "2026-02-02 to 2026-02-06" \
  --customer-name "ANZ" \
  --engagement-name "Professional Services"
```

Then generate PPT:

```bash
npm run report:ppt -- \
  --period "2026-02-02 to 2026-02-06" \
  --customer "ANZ" \
  --engagement "Professional Services"
```

## 5) Outputs

- Generic report:
  - `outputs/pm_weekly_report.md`
  - `outputs/pm_weekly_report.html`
- Customer report:
  - `outputs/<customer_slug>_anz_ps_weekly_report.md`
  - `outputs/<customer_slug>_anz_ps_weekly_report.html`
- PPT:
  - `outputs/<customer_slug>_<engagement_slug>_databricks_ps_status_report_<YYYY-MM-DD>.pptx`

## 6) Troubleshooting

- If live MCP fetch times out or fails auth, use Option A skills export workflow.
- If Salesforce is not available via MCP, export Salesforce data via skills and write `inputs/skills_exports/salesforce.json`.

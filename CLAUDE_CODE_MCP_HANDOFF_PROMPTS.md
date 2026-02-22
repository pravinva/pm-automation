# Claude Code Handoff: Real Data Pull (Any Customer) + Report + PPT

Use this in Claude Code when MCP/tools are available. This is customer-agnostic.

## Inputs to provide

- `CUSTOMER_NAME` (required)
- `ENGAGEMENT_NAME` (required)
- `PERIOD_LABEL` (example: `2026-02-02 to 2026-02-06`)
- `LOOKBACK_DAYS` (default `7`)

## One-shot prompt for Claude Code

```text
You are working in:
/Users/pravin.varma/Documents/Demo/pm-automation

Variables:
CUSTOMER_NAME = "<set customer>"
ENGAGEMENT_NAME = "<set engagement>"
PERIOD_LABEL = "<set period>"
LOOKBACK_DAYS = 7

Goal:
Generate weekly PS status outputs using REAL data for CUSTOMER_NAME from Slack, Glean/Genie, Salesforce, and Google Drive.

Steps:
1) Fetch source data for CUSTOMER_NAME from last LOOKBACK_DAYS days:
   - Slack: updates, blockers, decisions, owners, actions
   - Glean/Genie: relevant insights/docs
   - Salesforce: account/opportunity updates, stage, risk, next steps
   - Google Drive: RAID/plans/status docs updates

2) Normalize and write to:
   - inputs/skills_exports/slack.json      (top key: messages)
   - inputs/skills_exports/glean.json      (top key: insights)
   - inputs/skills_exports/salesforce.json (top key: opportunities)
   - inputs/skills_exports/gdrive.json     (top key: documents)

3) Run report generation:
   python3 pm_report_automation.py --project-root . --ingest-skills-exports --skills-dir "inputs/skills_exports" --period-label "$PERIOD_LABEL" --customer-name "$CUSTOMER_NAME" --engagement-name "$ENGAGEMENT_NAME"

4) Run PPT generation:
   npm run report:ppt -- --period "$PERIOD_LABEL" --customer "$CUSTOMER_NAME" --engagement "$ENGAGEMENT_NAME"

5) Return:
   - fetched record counts per source
   - auth/tool gaps (if any)
   - output files created

Constraints:
- Do not fabricate data.
- If a source is unavailable, write a valid empty payload for that source and report the limitation.
```

## Expected source schemas

- `slack.json`:
  - `{"messages":[{"title":"","owner":"","status":"","impact":"Low|Medium|High|Critical","detail":"","action":""}]}`
- `glean.json`:
  - `{"insights":[{"topic":"","owner":"","state":"","priority":"Low|Medium|High|Critical","summary":"","follow_up":""}]}`
- `salesforce.json`:
  - `{"opportunities":[{"account":"","name":"","owner":"","stage":"","risk":"Low|Medium|High|Critical","detail":"","next_step":""}]}`
- `gdrive.json`:
  - `{"documents":[{"title":"","owner":"","state":"","priority":"Low|Medium|High|Critical","summary":"","required_action":""}]}`

## Validation checklist

- Ensure all four exports exist in `inputs/skills_exports/`.
- Ensure each export is valid JSON and has expected top-level key.
- Ensure final outputs exist in `outputs/`:
  - `pm_weekly_report.md`
  - `pm_weekly_report.html`
  - `<customer_slug>_anz_ps_weekly_report.md`
  - `<customer_slug>_anz_ps_weekly_report.html`
  - `<customer_slug>_<engagement_slug>_databricks_ps_status_report_<date>.pptx`

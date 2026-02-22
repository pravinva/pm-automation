# Plan: Pull Real Data and Generate Ventia Report

This plan is for running in an environment where Claude Code has working access to Slack, Genie/Glean, Salesforce, and Google Drive tools.

## Goal

Generate a Ventia PM status report for ANZ Professional Services using real source data from:

- Slack
- Genie (Glean)
- Salesforce
- Google Drive

## Scope

- Fetch latest weekly signals from each source.
- Normalize into the report input schema.
- Produce Markdown and HTML report outputs for Ventia.
- Keep a copy of raw source exports for traceability.

## Step-by-Step Execution

1. **Validate tool access**
   - Confirm Slack, Glean/Genie, Salesforce, and Google tools are authenticated.
   - Confirm each source returns data for the last 7 days.

2. **Export raw source data**
   - Save raw extracts in `inputs/skills_exports/`:
     - `slack.json` (or `.md`/`.txt`)
     - `glean.json` (or `.md`/`.txt`)
     - `salesforce.json` (or `.md`/`.txt`)
     - `gdrive.json` (or `.md`/`.txt`)

3. **Run ingestion and report generation**
   - Run:
     - `python3 pm_report_automation.py --project-root . --ingest-skills-exports --skills-dir "inputs/skills_exports"`

4. **Create Ventia-specific report copies**
   - Copy generated files to:
     - `outputs/ventia_anz_ps_weekly_report.md`
     - `outputs/ventia_anz_ps_weekly_report.html`

5. **Review and enrich**
   - Check top risks, blockers, owners, and action items are accurate.
   - Ensure Salesforce commercial signals are reflected in executive summary.
   - Ensure Slack and GDrive progress notes are represented in highlights.

6. **Deliverable check**
   - Confirm these files exist:
     - `outputs/pm_weekly_report.md`
     - `outputs/pm_weekly_report.html`
     - `outputs/ventia_anz_ps_weekly_report.md`
     - `outputs/ventia_anz_ps_weekly_report.html`

## Commands (Quick Run)

```bash
cd "/Users/pravin.varma/Documents/Demo/pm-automation"
python3 pm_report_automation.py --project-root . --ingest-skills-exports --skills-dir "inputs/skills_exports"
cp "outputs/pm_weekly_report.md" "outputs/ventia_anz_ps_weekly_report.md"
cp "outputs/pm_weekly_report.html" "outputs/ventia_anz_ps_weekly_report.html"
```

## Notes

- If any source is missing, the script falls back to an empty/default shape for that source.
- For best output quality, provide JSON exports using the expected keys:
  - Slack: `messages`
  - Salesforce: `opportunities`
  - GDrive: `documents`
  - Glean: `insights`

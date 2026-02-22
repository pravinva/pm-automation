# PM Automation

This project generates a weekly PM report by consolidating updates from:

- Slack MCP
- Google Drive MCP
- Salesforce MCP
- Glean MCP

## What is already configured

From your current local config, the MCP setup includes:

- `slack`
- `glean`
- `google` (use this for Drive/Docs/Sheets)

`salesforce` is not currently present in the MCP config files, so either add it as an MCP server or export Salesforce data through your existing workflow into `inputs/salesforce.json`.

## Quick start

1. Run with live MCP fetch enabled:

```bash
python3 pm_report_automation.py --project-root . --fetch-live --lookback-days 7
```

2. Outputs are written to:
   - `outputs/pm_weekly_report.md`
   - `outputs/pm_weekly_report.html`
   - `outputs/<customer_slug>_anz_ps_weekly_report.md`
   - `outputs/<customer_slug>_anz_ps_weekly_report.html`
   - `outputs/<customer_slug>_<engagement_slug>_databricks_ps_status_report_YYYY-MM-DD.pptx` (via PptxGenJS)

3. Live fetch writes latest source payloads to:
   - `inputs/slack.json`
   - `inputs/salesforce.json`
   - `inputs/gdrive.json`
   - `inputs/glean.json`

### Optional

- If your MCP config is in a custom path:

```bash
python3 pm_report_automation.py --project-root . --fetch-live --mcp-config "~/.config/mcp/config.json"
```

- To run from existing local JSON files only (no live fetch):

```bash
python3 pm_report_automation.py --project-root .
```

- To set the report period label shown in engagement sections:

```bash
python3 pm_report_automation.py --project-root . --period-label "2026-02-02 to 2026-02-06" --customer-name "ANZ" --engagement-name "Professional Services"
```

- To ingest Claude skill exports first (no Salesforce MCP required):

```bash
python3 pm_report_automation.py --project-root . --ingest-skills-exports --skills-dir "inputs/skills_exports"
```

Skill export directory supports files named:
- `slack.json|md|txt`
- `salesforce.json|md|txt`
- `gdrive.json|md|txt`
- `glean.json|md|txt`

## Input schema

The script expects the following top-level keys:

- Slack: `messages`
- Salesforce: `opportunities`
- GDrive: `documents`
- Glean: `insights`

Sample payloads are included in `inputs/`.

## Notes for MCP wiring

- Live fetch is implemented via `mcp_live_fetch.py` using stdio MCP JSON-RPC.
- Config discovery order:
  1. `~/.config/mcp/config.json`
  2. `~/.claude.json` (`mcpServers`)
- Current expected server names:
  - Slack: `slack`
  - GDrive: `google`
  - Salesforce: `salesforce`
  - Glean: `glean`
- If a server is missing or fetch fails, the script keeps/falls back to local `inputs/*.json` data.

## Troubleshooting live fetch

- If you see permission errors like:
  - `MCP server binary is not readable: .../google_mcp_deploy.pex`
  - `Permission denied: ~/.local/state/mcp-servers/...`
- Fix ownership/permissions (example commands):

```bash
sudo chown -R "$USER":staff ~/mcp/servers
sudo chmod -R u+rwX ~/mcp/servers
sudo chown -R "$USER":staff ~/.local/state/mcp-servers
sudo chmod -R u+rwX ~/.local/state/mcp-servers
```

- If Salesforce is not discovered, add a `salesforce` MCP server entry to your MCP config and rerun with `--fetch-live`.

## PPT generation (PptxGenJS)

Install dependencies:

```bash
npm install
```

Generate Ventia-format PPT:

```bash
npm run report:ppt -- --customer "Ventia" --engagement "Supply Chain"
```

Optional period label for script context:

```bash
npm run report:ppt -- --period "2026-02-02 to 2026-02-06" --customer "ANZ" --engagement "Professional Services"
```

## Auth requirements

- **Skills mode (`--ingest-skills-exports`)**: this script does not directly authenticate. It reads files produced by your skill runs.
- **Live MCP mode (`--fetch-live`)**: authentication is required by each MCP server (Slack/Google/Glean/Salesforce) in your local environment.
- For Salesforce specifically:
  - If you use **skills export**, auth is handled when the skill fetches data.
  - If you use **MCP live fetch**, you need a configured `salesforce` MCP server plus valid auth.

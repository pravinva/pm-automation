#!/usr/bin/env python3
"""Generate a PM status report from MCP-exported source data."""

from __future__ import annotations

import argparse
import datetime as dt
import html
import json
from pathlib import Path
from typing import Any
from mcp_live_fetch import load_mcp_servers, fetch_source_data, save_json, print_server_summary


def read_json(path: Path, fallback: dict[str, Any]) -> dict[str, Any]:
    if not path.exists():
        return fallback
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _read_text_file(path: Path) -> str:
    if not path.exists():
        return ""
    return path.read_text(encoding="utf-8").strip()


def _slugify(value: str) -> str:
    cleaned = "".join(ch.lower() if ch.isalnum() else "-" for ch in value.strip())
    while "--" in cleaned:
        cleaned = cleaned.replace("--", "-")
    return cleaned.strip("-") or "customer"


def _extract_data(payload: dict[str, Any]) -> dict[str, Any]:
    """Handle both plain schema and live-fetch wrapped payloads."""
    if "data" in payload and isinstance(payload.get("data"), dict):
        return payload["data"]
    return payload


def _raw_to_messages(raw: dict[str, Any]) -> list[dict[str, str]]:
    text = raw.get("raw_text", "")
    if not text:
        return []
    return [
        {
            "title": "Slack weekly updates",
            "owner": "Slack MCP",
            "status": "info",
            "impact": "Medium",
            "detail": text[:800],
            "action": "",
        }
    ]


def _raw_to_documents(raw: dict[str, Any], source_name: str) -> list[dict[str, str]]:
    text = raw.get("raw_text", "")
    if not text:
        return []
    return [
        {
            "title": f"{source_name} weekly extract",
            "owner": f"{source_name} MCP",
            "state": "updated",
            "priority": "Medium",
            "summary": text[:800],
            "required_action": "",
        }
    ]


def normalize_slack(payload: dict[str, Any]) -> list[dict[str, str]]:
    payload = _extract_data(payload)
    if "raw_text" in payload and "messages" not in payload:
        payload = {"messages": _raw_to_messages(payload)}
    items = []
    for msg in payload.get("messages", []):
        items.append(
            {
                "source": "Slack",
                "title": msg.get("title", "Channel update"),
                "owner": msg.get("owner", "Unknown"),
                "status": msg.get("status", "update"),
                "impact": msg.get("impact", "Medium"),
                "detail": msg.get("detail", ""),
                "action": msg.get("action", ""),
            }
        )
    return items


def normalize_salesforce(payload: dict[str, Any]) -> list[dict[str, str]]:
    payload = _extract_data(payload)
    if "raw_text" in payload and "opportunities" not in payload:
        payload = {
            "opportunities": [
                {
                    "account": "Salesforce extract",
                    "name": "Pipeline signal",
                    "owner": "Salesforce MCP",
                    "stage": "Unknown",
                    "risk": "Medium",
                    "detail": payload["raw_text"][:800],
                    "next_step": "",
                }
            ]
        }
    items = []
    for opp in payload.get("opportunities", []):
        items.append(
            {
                "source": "Salesforce",
                "title": f"{opp.get('account', 'Account')} - {opp.get('name', 'Opportunity')}",
                "owner": opp.get("owner", "Unknown"),
                "status": opp.get("stage", "Unknown"),
                "impact": opp.get("risk", "Medium"),
                "detail": opp.get("detail", ""),
                "action": opp.get("next_step", ""),
            }
        )
    return items


def normalize_gdrive(payload: dict[str, Any]) -> list[dict[str, str]]:
    payload = _extract_data(payload)
    if "raw_text" in payload and "documents" not in payload:
        payload = {"documents": _raw_to_documents(payload, "Google Drive")}
    items = []
    for doc in payload.get("documents", []):
        items.append(
            {
                "source": "GDrive",
                "title": doc.get("title", "Document update"),
                "owner": doc.get("owner", "Unknown"),
                "status": doc.get("state", "updated"),
                "impact": doc.get("priority", "Medium"),
                "detail": doc.get("summary", ""),
                "action": doc.get("required_action", ""),
            }
        )
    return items


def normalize_glean(payload: dict[str, Any]) -> list[dict[str, str]]:
    payload = _extract_data(payload)
    if "raw_text" in payload and "insights" not in payload:
        payload = {
            "insights": [
                {
                    "topic": "Glean weekly search signal",
                    "owner": "Glean MCP",
                    "state": "info",
                    "priority": "Medium",
                    "summary": payload["raw_text"][:800],
                    "follow_up": "",
                }
            ]
        }
    items = []
    for insight in payload.get("insights", []):
        items.append(
            {
                "source": "Glean",
                "title": insight.get("topic", "Knowledge signal"),
                "owner": insight.get("owner", "Unknown"),
                "status": insight.get("state", "info"),
                "impact": insight.get("priority", "Medium"),
                "detail": insight.get("summary", ""),
                "action": insight.get("follow_up", ""),
            }
        )
    return items


def priority_weight(priority: str) -> int:
    rank = {"Critical": 3, "High": 2, "Medium": 1, "Low": 0}
    return rank.get(priority, 1)


def generate_markdown(items: list[dict[str, str]], report_date: str) -> str:
    top_risks = [i for i in items if i["impact"] in {"Critical", "High"}][:5]
    completed = [i for i in items if i["status"].lower() in {"closed", "done", "completed"}][:5]
    in_flight = [i for i in items if i["status"].lower() not in {"closed", "done", "completed"}][:8]

    lines = [
        f"# PM Weekly Report ({report_date})",
        "",
        "## Executive Snapshot",
        f"- Total updates captured: **{len(items)}**",
        f"- Critical/High priority items: **{len([i for i in items if i['impact'] in {'Critical', 'High'}])}**",
        f"- Sources: **{', '.join(sorted({i['source'] for i in items}))}**",
        "",
        "## Top Risks",
    ]
    if not top_risks:
        lines.append("- No high-severity risks identified.")
    else:
        for item in top_risks:
            lines.append(
                f"- **[{item['source']}] {item['title']}** ({item['owner']}) - {item['detail']} | Next: {item['action'] or 'TBD'}"
            )

    lines.extend(["", "## Progress Highlights"])
    if not completed:
        lines.append("- No completed items reported this cycle.")
    else:
        for item in completed:
            lines.append(f"- **[{item['source']}] {item['title']}** - {item['detail']}")

    lines.extend(["", "## In-Flight Work"])
    for item in in_flight:
        lines.append(
            f"- **[{item['source']}] {item['title']}** ({item['status']}, {item['impact']}) - {item['detail']} | Owner: {item['owner']}"
        )

    lines.extend(["", "## Action Register"])
    for item in items:
        if item["action"]:
            lines.append(f"- {item['title']}: {item['action']}")

    return "\n".join(lines) + "\n"


def generate_html(items: list[dict[str, str]], report_date: str) -> str:
    rows = []
    for item in items:
        rows.append(
            "<tr>"
            f"<td>{html.escape(item['source'])}</td>"
            f"<td>{html.escape(item['title'])}</td>"
            f"<td>{html.escape(item['owner'])}</td>"
            f"<td>{html.escape(item['status'])}</td>"
            f"<td>{html.escape(item['impact'])}</td>"
            f"<td>{html.escape(item['detail'])}</td>"
            f"<td>{html.escape(item['action'])}</td>"
            "</tr>"
        )

    return f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>PM Report {html.escape(report_date)}</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 24px; color: #222; }}
    h1 {{ margin-bottom: 8px; }}
    .meta {{ color: #666; margin-bottom: 20px; }}
    table {{ border-collapse: collapse; width: 100%; }}
    th, td {{ border: 1px solid #ddd; padding: 8px; vertical-align: top; }}
    th {{ background: #f5f5f5; text-align: left; }}
  </style>
</head>
<body>
  <h1>PM Weekly Report</h1>
  <div class="meta">Generated: {html.escape(report_date)} | Records: {len(items)}</div>
  <table>
    <thead>
      <tr>
        <th>Source</th><th>Title</th><th>Owner</th><th>Status</th><th>Impact</th><th>Detail</th><th>Action</th>
      </tr>
    </thead>
    <tbody>
      {''.join(rows)}
    </tbody>
  </table>
</body>
</html>
"""


def _status_rag_from_items(items: list[dict[str, str]]) -> tuple[str, str, str, str]:
    has_critical = any(i["impact"] == "Critical" for i in items)
    has_high = any(i["impact"] == "High" for i in items)
    overall = "Red" if has_critical else ("Amber" if has_high else "Green")
    scope = "Amber" if has_high else "Green"
    schedule = "Amber" if has_high else "Green"
    make_it_right = "Green"
    return overall, scope, schedule, make_it_right


def _status_bucket(status: str, impact: str) -> str:
    s = status.lower()
    if s in {"done", "completed", "closed"}:
        return "Complete"
    if s in {"blocked"}:
        return "Blocked"
    if impact in {"Critical", "High"}:
        return "At Risk"
    if s in {"not started", "todo"}:
        return "Not Started"
    return "In Progress"


def generate_ventia_markdown(
    items: list[dict[str, str]], report_date: str, period_label: str, customer_name: str, engagement_name: str
) -> str:
    overall, scope, schedule, make_it_right = _status_rag_from_items(items)
    risks = [i for i in items if i["impact"] in {"Critical", "High"}]
    actions = [i for i in items if i["action"]]
    sources = sorted({i["source"] for i in items})

    lines = [
        f"# {customer_name} {engagement_name} - Databricks PS Engagement",
        "## Weekly Status Report",
        f"**Date:** {report_date}",
        "",
        "## Agenda",
        "- Teams",
        "- Status Updates / Issues, risks",
        "- High Level Plan",
        "- Resource Plan",
        "- Key Points to Discuss",
        "",
        "## 1) Teams",
        f"- {customer_name}: Customer sponsor, data lead, engineering lead",
        "- Databricks / Partner: RSA, Senior PM, Delivery Engineers, Account Team",
        "",
        "## 2) Status Updates / Issues, Risks",
        "### Action Items",
        "| S. No. | Date | Description | Owner | Status | Comments |",
        "|---|---|---|---|---|---|",
    ]

    for idx, item in enumerate(actions[:12], start=1):
        status = _status_bucket(item["status"], item["impact"])
        lines.append(
            f"| {idx} | {report_date} | {item['title']} | {item['owner']} | {status} | {item['action']} |"
        )
    if len(actions) == 0:
        lines.append("| 1 | - | No actions captured | - | - | - |")

    lines.extend(
        [
            "",
            "### Engagement Status",
            f"- **Project Status (Overall):** {overall}",
            f"- **Scope:** {scope}",
            f"- **Schedule:** {schedule}",
            f"- **Make-It-Right:** {make_it_right}",
            f"- **Status Summary:** Weekly data consolidated from {', '.join(sources) if sources else 'source systems'} for period {period_label}.",
            "",
            "### Points to discuss",
        ]
    )
    for r in risks[:5]:
        lines.append(f"- {r['title']}: {r['detail']}")
    if len(risks) == 0:
        lines.append("- No high-severity points to discuss this period.")

    lines.extend(
        [
            "",
            "### Milestones / Phases / Deliverables",
            "| Item | Status | Target Date |",
            "|---|---|---|",
            f"| Discovery & Design Alignment | {'In Progress' if items else 'Not Started'} | TBC |",
            f"| Build / Validation Stream | {'In Progress' if items else 'Not Started'} | TBC |",
            f"| Reporting & Handover | {'In Progress' if items else 'Not Started'} | TBC |",
            "",
            "**Legend:** Complete | In Progress | At Risk | Blocked | Not Started",
            "",
            "### Key Accomplishments & Next Steps",
            f"**Accomplishments this period ({period_label})**",
        ]
    )
    for i in items[:8]:
        lines.append(f"- {i['title']} ({i['source']}): {i['detail']}")
    if len(items) == 0:
        lines.append("- No source updates captured.")

    lines.append("")
    lines.append("**Activities for next period**")
    for a in actions[:8]:
        lines.append(f"- {a['title']}: {a['action']}")
    if len(actions) == 0:
        lines.append("- Confirm source updates and define action owners.")

    lines.extend(
        [
            "",
            "## Risk & Issue",
            "| ID | Type | Description | Impact | Probability | Action(s) - Owner | Status |",
            "|---|---|---|---|---|---|---|",
        ]
    )
    for idx, r in enumerate(risks[:10], start=1):
        lines.append(
            f"| {idx:02d} | Risk | {r['title']} - {r['detail']} | {r['impact']} | Med | {r['action'] or 'Mitigation TBD'} - {r['owner']} | {r['status']} |"
        )
    if len(risks) == 0:
        lines.append("| 01 | Risk | No high-severity risks captured | Low | Low | Continue monitoring - PM | Open |")

    lines.extend(
        [
            "",
            "## 3) High Level Plan",
            "| Item | Current Status | Notes |",
            "|---|---|---|",
            f"| Requirements & Design | {'In Progress' if items else 'Not Started'} | Design decisions and stakeholder approvals in progress. |",
            f"| Build & Validation | {'In Progress' if items else 'Not Started'} | Weekly execution and quality checks across workstreams. |",
            f"| Readout & Handover | {'In Progress' if items else 'Not Started'} | PM reporting cadence and governance updates. |",
            "",
            "## 4) Resource Plan",
            "| Name | Role | Hours | 19/1 | 26/1 | 2/2 | 9/2 | 16/2 | 23/2 |",
            "|---|---|---:|---:|---:|---:|---:|---:|---:|",
            "| Delivery Lead | Project Delivery | 128 | 16 | 16 | 16 | 16 | 16 | 16 |",
            "| Data Engineer | Engineering | 240 | 40 | 40 | 40 | 40 | 40 | 40 |",
            "| PM | Project Management | 64 | 8 | 8 | 8 | 8 | 8 | 8 |",
            "| RSA | Architecture / Advisory | 32 | 8 | 8 | 8 | 8 | - | - |",
            "",
            "## Plan Tracking",
            "- Tracked in customer Jira / agreed work tracking board.",
            "",
            "## 5) Key Points to Discuss",
            "- Confirm acceptance criteria and sign-off windows.",
            "- Confirm dependency closure dates and owners.",
            "- Confirm next-week priorities and stakeholder readiness.",
            "",
            "## Appendix",
            "- Generated from PM automation pipeline.",
        ]
    )
    return "\n".join(lines) + "\n"


def generate_ventia_html(
    items: list[dict[str, str]], report_date: str, period_label: str, customer_name: str, engagement_name: str
) -> str:
    md = generate_ventia_markdown(items, report_date, period_label, customer_name, engagement_name)
    escaped = html.escape(md)
    return f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>{html.escape(customer_name)} {html.escape(engagement_name)} - PS Weekly Status</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 28px; color: #222; }}
    h1 {{ margin-bottom: 8px; }}
    .note {{ color: #555; margin-bottom: 16px; }}
    pre {{ white-space: pre-wrap; line-height: 1.35; }}
  </style>
</head>
<body>
  <div class="note">{html.escape(customer_name)} template format view ({html.escape(report_date)} / {html.escape(period_label)})</div>
  <pre>{escaped}</pre>
</body>
</html>
"""


def fetch_live_inputs(
    project_root: Path, lookback_days: int, config_path: Path | None = None, customer_name: str = ""
) -> None:
    """Fetch latest source data via configured MCP servers into inputs/*.json."""
    inputs_dir = project_root / "inputs"
    servers = load_mcp_servers(config_path)
    print_server_summary(servers)

    source_server_pairs = {
        "slack": "slack",
        "gdrive": "google",
        "salesforce": "salesforce",
        "glean": "glean",
    }
    fallback_shapes = {
        "slack": {"messages": []},
        "gdrive": {"documents": []},
        "salesforce": {"opportunities": []},
        "glean": {"insights": []},
    }

    for source, server_name in source_server_pairs.items():
        out_path = inputs_dir / f"{source}.json"
        if server_name not in servers:
            print(f"[warn] MCP server '{server_name}' not configured, keeping {out_path.name} fallback.")
            if not out_path.exists():
                save_json(out_path, fallback_shapes[source])
            continue
        try:
            payload = fetch_source_data(servers[server_name], source, lookback_days, customer_name=customer_name)
            save_json(out_path, payload)
            print(f"[ok] Fetched {source} into {out_path}")
        except Exception as exc:
            print(f"[warn] Failed live fetch for {source}: {exc}")
            if not out_path.exists():
                save_json(out_path, fallback_shapes[source])


def ingest_skill_exports(project_root: Path, skills_dir: Path, overwrite: bool = True) -> None:
    """Ingest Claude skill export files into inputs/*.json.

    Supported filenames per source:
    - <source>.json or <source>.md or <source>.txt
    where source in: slack, salesforce, gdrive, glean
    """
    inputs_dir = project_root / "inputs"
    inputs_dir.mkdir(parents=True, exist_ok=True)

    shapes = {
        "slack": ("messages", "Slack skill export"),
        "salesforce": ("opportunities", "Salesforce skill export"),
        "gdrive": ("documents", "Google Drive skill export"),
        "glean": ("insights", "Glean skill export"),
    }

    for source, (key, label) in shapes.items():
        json_src = skills_dir / f"{source}.json"
        md_src = skills_dir / f"{source}.md"
        txt_src = skills_dir / f"{source}.txt"
        out_path = inputs_dir / f"{source}.json"

        if json_src.exists():
            try:
                payload = json.loads(json_src.read_text(encoding="utf-8"))
                if not isinstance(payload, dict):
                    raise ValueError("JSON export must be an object")
                if overwrite or not out_path.exists():
                    save_json(out_path, payload)
                print(f"[ok] Imported {source} from {json_src}")
                continue
            except Exception as exc:
                print(f"[warn] Could not parse {json_src}: {exc}")

        text = _read_text_file(md_src) or _read_text_file(txt_src)
        if text:
            payload = {"raw_text": text}
            if overwrite or not out_path.exists():
                save_json(out_path, payload)
            print(f"[ok] Imported {source} text export into {out_path}")
            continue

        if not out_path.exists():
            save_json(out_path, {key: []})
        print(f"[warn] No skill export found for {source} in {skills_dir}; using fallback.")


def build_report(
    project_root: Path, period_label: str, customer_name: str, engagement_name: str
) -> tuple[Path, Path, Path, Path]:
    inputs_dir = project_root / "inputs"
    outputs_dir = project_root / "outputs"
    outputs_dir.mkdir(parents=True, exist_ok=True)

    slack = read_json(inputs_dir / "slack.json", {"messages": []})
    salesforce = read_json(inputs_dir / "salesforce.json", {"opportunities": []})
    gdrive = read_json(inputs_dir / "gdrive.json", {"documents": []})
    glean = read_json(inputs_dir / "glean.json", {"insights": []})

    items = (
        normalize_slack(slack)
        + normalize_salesforce(salesforce)
        + normalize_gdrive(gdrive)
        + normalize_glean(glean)
    )
    items.sort(key=lambda x: priority_weight(x["impact"]), reverse=True)

    report_date = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    md_path = outputs_dir / "pm_weekly_report.md"
    html_path = outputs_dir / "pm_weekly_report.html"
    customer_slug = _slugify(customer_name)
    ventia_md_path = outputs_dir / f"{customer_slug}_anz_ps_weekly_report.md"
    ventia_html_path = outputs_dir / f"{customer_slug}_anz_ps_weekly_report.html"

    md_path.write_text(generate_markdown(items, report_date), encoding="utf-8")
    html_path.write_text(generate_html(items, report_date), encoding="utf-8")
    ventia_md_path.write_text(
        generate_ventia_markdown(items, report_date, period_label, customer_name, engagement_name), encoding="utf-8"
    )
    ventia_html_path.write_text(
        generate_ventia_html(items, report_date, period_label, customer_name, engagement_name), encoding="utf-8"
    )
    return md_path, html_path, ventia_md_path, ventia_html_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate PM report from MCP exports.")
    parser.add_argument("--project-root", default=".", help="Project root directory")
    parser.add_argument("--fetch-live", action="store_true", help="Fetch live MCP data before report generation")
    parser.add_argument("--lookback-days", type=int, default=7, help="How many days to look back for MCP queries")
    parser.add_argument("--mcp-config", default="", help="Optional path to MCP config JSON")
    parser.add_argument(
        "--ingest-skills-exports",
        action="store_true",
        help="Import skill exports from a directory into inputs/*.json",
    )
    parser.add_argument(
        "--skills-dir",
        default="inputs/skills_exports",
        help="Directory containing slack/salesforce/gdrive/glean exports (.json/.md/.txt)",
    )
    parser.add_argument("--period-label", default="current reporting period", help="Label for report period text")
    parser.add_argument("--customer-name", default="Ventia", help="Customer name used in report titles")
    parser.add_argument("--engagement-name", default="Supply Chain", help="Engagement name used in report titles")
    args = parser.parse_args()

    project_root = Path(args.project_root).resolve()
    mcp_config = Path(args.mcp_config).expanduser() if args.mcp_config else None
    if args.ingest_skills_exports:
        ingest_skill_exports(project_root, (project_root / args.skills_dir).resolve())
    if args.fetch_live:
        fetch_live_inputs(project_root, args.lookback_days, mcp_config, customer_name=args.customer_name)

    md_path, html_path, ventia_md_path, ventia_html_path = build_report(
        project_root, args.period_label, args.customer_name, args.engagement_name
    )
    print(f"Generated Markdown report: {md_path}")
    print(f"Generated HTML report: {html_path}")
    print(f"Generated Ventia Markdown report: {ventia_md_path}")
    print(f"Generated Ventia HTML report: {ventia_html_path}")


if __name__ == "__main__":
    main()

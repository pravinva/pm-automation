#!/usr/bin/env python3
"""Minimal stdio MCP client and source fetch helpers."""

from __future__ import annotations

import json
import os
import shlex
import select
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any


def _expand_path(value: str) -> str:
    return os.path.expandvars(os.path.expanduser(value))


@dataclass
class ServerConfig:
    name: str
    command: str
    args: list[str]
    env: dict[str, str]
    server_type: str

    @classmethod
    def from_dict(cls, name: str, payload: dict[str, Any]) -> "ServerConfig":
        return cls(
            name=name,
            command=_expand_path(str(payload.get("command", ""))),
            args=[_expand_path(str(x)) for x in payload.get("args", [])],
            env={k: str(v) for k, v in (payload.get("env") or {}).items()},
            server_type=str(payload.get("type", "stdio")),
        )


def load_mcp_servers(config_path: Path | None = None) -> dict[str, ServerConfig]:
    """Load MCP server configs from local files.

    Priority:
    1) ~/.config/mcp/config.json (vibe/global style)
    2) ~/.claude.json (mcpServers section)
    """
    candidates = []
    if config_path:
        candidates.append(config_path)
    else:
        candidates.extend(
            [
                Path("~/.config/mcp/config.json").expanduser(),
                Path("~/.claude.json").expanduser(),
            ]
        )

    for path in candidates:
        if not path.exists():
            continue
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue

        # ~/.config/mcp/config.json style:
        # { "claude-code": { "slack": {command,args,...} } }
        if "claude-code" in payload and isinstance(payload["claude-code"], dict):
            return {
                name: ServerConfig.from_dict(name, cfg)
                for name, cfg in payload["claude-code"].items()
                if isinstance(cfg, dict) and cfg.get("enabled", True)
            }

        # ~/.claude.json style:
        # { "mcpServers": { "slack": {command,args,...} } }
        if "mcpServers" in payload and isinstance(payload["mcpServers"], dict):
            return {
                name: ServerConfig.from_dict(name, cfg)
                for name, cfg in payload["mcpServers"].items()
                if isinstance(cfg, dict)
            }

    return {}


class McpClient:
    """Very small JSON-RPC over stdio client for MCP servers."""

    def __init__(self, server: ServerConfig, timeout_s: int = 30):
        self.server = server
        self.timeout_s = timeout_s
        self.proc: subprocess.Popen[bytes] | None = None
        self._id = 0

    def __enter__(self) -> "McpClient":
        env = os.environ.copy()
        env.update(self.server.env)
        # Some MCP server builds write logs under XDG state paths.
        # Force a user-writable location to avoid permission failures.
        state_home = Path.home() / ".pm-automation" / "state"
        state_home.mkdir(parents=True, exist_ok=True)
        env["XDG_STATE_HOME"] = str(state_home)
        if self.server.args:
            first_arg = self.server.args[0]
            arg_path = Path(first_arg)
            if arg_path.is_absolute() and arg_path.exists() and not os.access(arg_path, os.R_OK):
                raise RuntimeError(
                    f"MCP server binary is not readable: {arg_path}. "
                    "Fix file permissions for your user before running live fetch."
                )
        self.proc = subprocess.Popen(
            [self.server.command, *self.server.args],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=env,
        )
        self.initialize()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if not self.proc:
            return
        try:
            self.proc.terminate()
            self.proc.wait(timeout=2)
        except Exception:
            self.proc.kill()

    def _write_message(self, payload: dict[str, Any]) -> None:
        if not self.proc or not self.proc.stdin:
            raise RuntimeError("MCP process is not running")
        data = json.dumps(payload).encode("utf-8")
        header = f"Content-Length: {len(data)}\r\n\r\n".encode("ascii")
        self.proc.stdin.write(header + data)
        self.proc.stdin.flush()

    def _read_message(self) -> dict[str, Any]:
        if not self.proc or not self.proc.stdout:
            raise RuntimeError("MCP process is not running")
        deadline = time.monotonic() + self.timeout_s
        header = b""
        while b"\r\n\r\n" not in header:
            remaining = max(0.0, deadline - time.monotonic())
            if remaining == 0.0:
                raise RuntimeError(f"MCP read timeout after {self.timeout_s}s while reading header")
            ready, _, _ = select.select([self.proc.stdout], [], [], remaining)
            if not ready:
                raise RuntimeError(f"MCP read timeout after {self.timeout_s}s while reading header")
            chunk = self.proc.stdout.read(1)
            if not chunk:
                err = ""
                if self.proc.stderr:
                    try:
                        err = self.proc.stderr.read().decode("utf-8", errors="ignore").strip()
                    except Exception:
                        err = ""
                suffix = f" stderr: {err}" if err else ""
                raise RuntimeError(f"MCP stream closed while reading header.{suffix}")
            header += chunk
        header_text = header.decode("ascii", errors="ignore")
        length = None
        for line in header_text.split("\r\n"):
            if line.lower().startswith("content-length:"):
                length = int(line.split(":", 1)[1].strip())
                break
        if length is None:
            raise RuntimeError("Missing Content-Length header from MCP server")
        body = b""
        while len(body) < length:
            remaining = max(0.0, deadline - time.monotonic())
            if remaining == 0.0:
                raise RuntimeError(f"MCP read timeout after {self.timeout_s}s while reading body")
            ready, _, _ = select.select([self.proc.stdout], [], [], remaining)
            if not ready:
                raise RuntimeError(f"MCP read timeout after {self.timeout_s}s while reading body")
            chunk = self.proc.stdout.read(length - len(body))
            if not chunk:
                raise RuntimeError("MCP stream closed while reading body")
            body += chunk
        return json.loads(body.decode("utf-8"))

    def _request(self, method: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        self._id += 1
        req = {"jsonrpc": "2.0", "id": self._id, "method": method, "params": params or {}}
        self._write_message(req)
        while True:
            msg = self._read_message()
            if "id" not in msg:
                # Notification; ignore.
                continue
            if msg.get("id") != self._id:
                continue
            if "error" in msg:
                raise RuntimeError(f"{method} failed: {msg['error']}")
            return msg.get("result", {})

    def initialize(self) -> None:
        self._request(
            "initialize",
            {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "pm-automation", "version": "0.1.0"},
            },
        )
        # Best-effort notification expected by many servers.
        try:
            self._write_message({"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}})
        except Exception:
            pass

    def list_tools(self) -> list[dict[str, Any]]:
        result = self._request("tools/list", {})
        return result.get("tools", [])

    def call_tool(self, name: str, arguments: dict[str, Any] | None = None) -> dict[str, Any]:
        return self._request("tools/call", {"name": name, "arguments": arguments or {}})


def _parse_text_content(call_result: dict[str, Any]) -> str:
    content = call_result.get("content", [])
    parts: list[str] = []
    for item in content:
        if isinstance(item, dict):
            if "text" in item and item["text"] is not None:
                parts.append(str(item["text"]))
            elif "json" in item:
                parts.append(json.dumps(item["json"]))
    return "\n".join(parts).strip()


def _safe_json_from_result(call_result: dict[str, Any]) -> dict[str, Any]:
    text = _parse_text_content(call_result)
    if not text:
        return {}
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        # Keep raw text when a server returns plain text.
        return {"raw_text": text}


def _best_tool_name(tools: list[dict[str, Any]], preferences: list[str]) -> str | None:
    available = [t.get("name", "") for t in tools]
    lowered = {name.lower(): name for name in available}
    for pref in preferences:
        if pref.lower() in lowered:
            return lowered[pref.lower()]
    for name in available:
        lname = name.lower()
        if "search" in lname or "list" in lname or "query" in lname:
            return name
    return available[0] if available else None


def _call_with_attempts(client: McpClient, tool_name: str, arg_options: list[dict[str, Any]]) -> dict[str, Any]:
    last_error = None
    for args in arg_options:
        try:
            return client.call_tool(tool_name, args)
        except Exception as exc:
            last_error = exc
            continue
    if last_error:
        raise RuntimeError(str(last_error))
    raise RuntimeError("No tool call attempts were made")


def fetch_source_data(server_cfg: ServerConfig, source: str, lookback_days: int, customer_name: str = "") -> dict[str, Any]:
    with McpClient(server_cfg) as client:
        tools = client.list_tools()
        if not tools:
            return {"error": f"No tools exposed by {server_cfg.name}"}

        preferences = {
            "slack": ["slack_read_api_call", "search_messages", "list_channels"],
            "google": ["google_drive_search", "drive_search", "google_docs_search"],
            "salesforce": ["salesforce_query", "soql_query", "opportunity_list"],
            "glean": ["glean_read_api_call", "search", "search_documents"],
        }.get(source, [])

        selected = _best_tool_name(tools, preferences)
        if not selected:
            return {"error": f"No usable tool for {source}"}

        customer_query = f"{customer_name} weekly project status updates".strip() if customer_name else "weekly project status updates"
        call_result = _call_with_attempts(
            client,
            selected,
            [
                {"query": customer_query, "lookback_days": lookback_days},
                {"q": customer_query, "days": lookback_days},
                {"lookback_days": lookback_days},
                {},
            ],
        )
        parsed = _safe_json_from_result(call_result)
        return {
            "fetched_via": server_cfg.name,
            "tool": selected,
            "fetched_at": dt_now_iso(),
            "data": parsed,
        }


def dt_now_iso() -> str:
    import datetime as dt

    return dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def save_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def print_server_summary(servers: dict[str, ServerConfig]) -> None:
    if not servers:
        print("No MCP servers discovered from local config.")
        return
    print("Discovered MCP servers:")
    for name, cfg in servers.items():
        cmd = " ".join(shlex.quote(x) for x in [cfg.command, *cfg.args])
        print(f"- {name}: {cmd}")


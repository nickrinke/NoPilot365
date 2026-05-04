#!/usr/bin/env python3
# /// script
# dependencies = [
#   "anthropic>=0.92.0",
#   "msal>=1.31.0",
#   "requests>=2.32.0",
#   "rich>=13.0.0",
# ]
# ///
"""NoPilot 365 — natural language Microsoft 365 administration via Graph API."""

import atexit
import csv
import io
import json
import os
import shutil
import subprocess
import sys
import threading
from datetime import datetime, timezone
from pathlib import Path

import anthropic
import msal
import requests
from rich.console import Console
from rich.markdown import Markdown
from rich.panel import Panel
from rich.prompt import Confirm, Prompt

console = Console()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
CLI_DIR = Path.home() / ".m365cli"

SCOPES = [
    "https://graph.microsoft.com/User.ReadWrite.All",
    "https://graph.microsoft.com/Group.ReadWrite.All",
    "https://graph.microsoft.com/Directory.ReadWrite.All",
    "https://graph.microsoft.com/Organization.Read.All",
]

SYSTEM_PROMPT = """You are an expert Microsoft 365 administrator assistant. \
Help M365 admins accomplish tasks using natural language by translating \
requests into Microsoft Graph API calls or Exchange Online PowerShell commands.

When given a task:
1. Use search_m365_docs to find the correct approach when unsure
2. Prefer execute_graph_api for most tasks — it is faster and needs no extra setup
3. Use execute_powershell for tasks that require Exchange Online PowerShell, such as:
   - Message traces (Get-MessageTrace)
   - Transport rules (Get/New/Set-TransportRule)
   - Mail flow connectors (Get/New-InboundConnector, Get/New-OutboundConnector)
   - Detailed mailbox permissions (Add-MailboxPermission, Add-RecipientPermission)
   - Litigation hold, mailbox auditing, journaling
   - Email forwarding via Set-Mailbox -ForwardingSmtpAddress
   - Distribution group settings not available in Graph
4. Tell the user what you will do before doing it
5. For write operations, the system will ask for confirmation — do not ask again yourself
6. Present results in a clear, human-readable format

For bulk operations, use read_local_file to read CSV files, then process each row.

For offboarding a user, perform in order:
  1. Revoke sign-in sessions (POST /users/{id}/revokeSignInSessions)
  2. Disable account (PATCH /users/{id} accountEnabled: false)
  3. Remove all licenses (POST /users/{id}/assignLicense)
  4. Convert mailbox to shared if needed (execute_powershell: Set-Mailbox)

Always be precise. If parameters are ambiguous, ask for clarification.

After completing a task, report the results directly and concisely. Do not offer follow-up options, numbered lists of suggested next steps, or ask "Would you like me to..." questions. If the user wants something else, they will ask."""

TOOLS = [
    {
        "name": "search_m365_docs",
        "description": "Search Microsoft 365 documentation to find Graph API endpoints, required parameters, and permissions.",
        "input_schema": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "Search query"}
            },
            "required": ["query"],
        },
    },
    {
        "name": "execute_graph_api",
        "description": "Execute a Microsoft Graph API call. Write operations require user confirmation.",
        "input_schema": {
            "type": "object",
            "properties": {
                "method": {
                    "type": "string",
                    "enum": ["GET", "POST", "PATCH", "PUT", "DELETE"],
                },
                "endpoint": {
                    "type": "string",
                    "description": "Graph API path, e.g. '/users' or '/users/{id}/assignLicense'",
                },
                "body": {
                    "type": "object",
                    "description": "Request body for POST/PATCH/PUT requests",
                },
                "description": {
                    "type": "string",
                    "description": "Human-readable description shown to the user before confirmation",
                },
            },
            "required": ["method", "endpoint", "description"],
        },
    },
    {
        "name": "read_local_file",
        "description": "Read a local CSV or text file from the user's computer for bulk operations.",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "Absolute or ~ path to the file"},
            },
            "required": ["path"],
        },
    },
    {
        "name": "execute_powershell",
        "description": (
            "Run an Exchange Online PowerShell command for tasks not available in Graph API: "
            "message traces, transport rules, mail flow connectors, detailed mailbox permissions, "
            "litigation hold, journaling, and other EXO-only operations. "
            "Write commands require user confirmation."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "command": {
                    "type": "string",
                    "description": "PowerShell command to execute",
                },
                "description": {
                    "type": "string",
                    "description": "Human-readable description of what this command does",
                },
                "is_write": {
                    "type": "boolean",
                    "description": "True if this command makes changes (New-, Set-, Add-, Remove-, Enable-, Disable-)",
                },
            },
            "required": ["command", "description", "is_write"],
        },
    },
]

HELP_TEXT = """
[bold]Commands[/bold]

  [green]help[/green]              Show this help
  [green]clear[/green]             Reset conversation context
  [green]readonly[/green]          Toggle read-only mode (blocks all write operations)
  [green]dryrun[/green]            Toggle dry-run mode (shows what would happen, doesn't execute)
  [green]export [file][/green]     Export last result to CSV (default: export.csv)
  [green]licenses[/green]          Show license inventory dashboard
  [green]offboard <email>[/green]  Run full offboarding workflow for a user
  [green]report passwords[/green]  Show users with passwords expiring soon
  [green]disconnect exchange[/green] Disconnect the Exchange Online PowerShell session
  [green]exit[/green]              Quit

[bold]Examples[/bold]

  Create a user named Jane Smith with email jsmith@contoso.com
  Assign an E3 license to john@contoso.com
  Show all members of the IT Admins group
  Reset the password for sarah@contoso.com
  List all users who haven't signed in for 90 days
"""


# ─── STATE ────────────────────────────────────────────────────────────────────

class State:
    readonly: bool = False
    dryrun: bool = False
    last_result: list[dict] | None = None


state = State()


# ─── EXCHANGE ONLINE SESSION ──────────────────────────────────────────────────

class ExchangeSession:
    """Persistent Exchange Online PowerShell session."""

    _SENTINEL = "<<<EXO_DONE_a3f9>>>"

    def __init__(self):
        self._proc: subprocess.Popen | None = None
        self._lock = threading.Lock()

    @property
    def connected(self) -> bool:
        return self._proc is not None and self._proc.poll() is None

    def connect(self, upn: str) -> str:
        if not shutil.which("pwsh"):
            return (
                "PowerShell 7 is not installed.\n"
                "Install it with: brew install powershell\n"
                "Then restart the app."
            )

        self._proc = subprocess.Popen(
            ["pwsh", "-NoLogo", "-NoExit", "-Command", "-"],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
        )

        connect_cmd = (
            "if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) { "
            "  Write-Host 'Installing ExchangeOnlineManagement module...'; "
            "  Install-Module ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser "
            "} "
            f"Connect-ExchangeOnline -UserPrincipalName '{upn}' -ShowBanner:$false"
        )
        return self._exec(connect_cmd, timeout=120)

    def run(self, command: str, timeout: int = 60) -> str:
        if not self.connected:
            return json.dumps({"error": "Not connected to Exchange Online."})
        return self._exec(command, timeout=timeout)

    def _exec(self, command: str, timeout: int = 60) -> str:
        with self._lock:
            wrapped = (
                f"try {{ {command} | ConvertTo-Json -Depth 4 -WarningAction SilentlyContinue }} "
                f"catch {{ Write-Output ('{{'+'\"error\": \"' + $_.Exception.Message + '\"'+'}}') }}; "
                f"Write-Output '{self._SENTINEL}'\n"
            )
            self._proc.stdin.write(wrapped)
            self._proc.stdin.flush()

            lines: list[str] = []
            done = threading.Event()

            def _read():
                for line in self._proc.stdout:
                    if self._SENTINEL in line:
                        done.set()
                        return
                    lines.append(line)

            t = threading.Thread(target=_read, daemon=True)
            t.start()
            if not done.wait(timeout=timeout):
                return json.dumps({"error": f"Command timed out after {timeout}s"})

            output = "".join(lines).strip()
            return output if output else json.dumps({"success": True})

    def disconnect(self):
        if self.connected:
            try:
                self._exec("Disconnect-ExchangeOnline -Confirm:$false", timeout=15)
                self._proc.terminate()
            except Exception:
                pass
        self._proc = None


exo_session = ExchangeSession()


# ─── AUDIT LOG ────────────────────────────────────────────────────────────────

def audit(command: str, method: str, endpoint: str, status: int, dry: bool = False) -> None:
    CLI_DIR.mkdir(exist_ok=True)
    entry = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "command": command,
        "method": method,
        "endpoint": endpoint,
        "status_code": status,
        "dry_run": dry,
    }
    with open(CLI_DIR / "audit.log", "a") as f:
        f.write(json.dumps(entry) + "\n")


# ─── AUTH ─────────────────────────────────────────────────────────────────────

def authenticate(client_id: str, tenant_id: str) -> str:
    cache = msal.SerializableTokenCache()
    cache_path = CLI_DIR / "token_cache.json"
    CLI_DIR.mkdir(exist_ok=True)
    if cache_path.exists():
        cache.deserialize(cache_path.read_text())

    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=cache,
    )

    def _save():
        if cache.has_state_changed:
            cache_path.write_text(cache.serialize())

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save()
            console.print("[green]✓ Authenticated (cached token)[/green]")
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Device flow failed: {flow.get('error_description')}")

    console.print(
        Panel(
            f"Open [link=https://microsoft.com/devicelogin]https://microsoft.com/devicelogin[/link]\n"
            f"and enter code: [bold green]{flow['user_code']}[/bold green]",
            title="🔐 Sign in to Microsoft 365",
        )
    )

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Authentication failed: {result.get('error_description')}")

    _save()
    console.print("[green]✓ Authenticated[/green]")
    return result["access_token"]


# ─── TOOL IMPLEMENTATIONS ─────────────────────────────────────────────────────

def search_m365_docs(query: str) -> str:
    try:
        resp = requests.get(
            "https://learn.microsoft.com/api/search",
            params={"search": query, "locale": "en-us", "$top": 5},
            timeout=10,
        )
        resp.raise_for_status()
        items = resp.json().get("results", [])
        results = [
            {
                "title": item.get("title", ""),
                "url": item.get("url", ""),
                "excerpt": item.get("description", "")[:400],
            }
            for item in items[:5]
        ]
        return json.dumps(results, indent=2) if results else "No documentation results found."
    except Exception as e:
        return f"Search error: {e}"


def call_graph(token: str, method: str, endpoint: str, body: dict | None) -> dict:
    url = endpoint if endpoint.startswith("http") else f"{GRAPH_BASE}{endpoint}"
    resp = requests.request(
        method,
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=body,
        timeout=30,
    )
    try:
        data = resp.json()
    except Exception:
        data = {"raw_response": resp.text}
    data["_status_code"] = resp.status_code
    return data


def read_local_file(path: str) -> str:
    try:
        p = Path(path).expanduser()
        if not p.exists():
            return f"File not found: {path}"
        text = p.read_text()
        if p.suffix.lower() == ".csv":
            reader = csv.DictReader(io.StringIO(text))
            rows = list(reader)
            return json.dumps(rows, indent=2)
        return text
    except Exception as e:
        return f"Error reading file: {e}"


def export_last_result(filename: str = "export.csv", log_fn=None) -> None:
    if log_fn is None:
        log_fn = console.print
    if not state.last_result:
        log_fn("[yellow]No data to export. Run a list command first.[/yellow]")
        return
    path = Path(filename).expanduser()
    keys = list(state.last_result[0].keys())
    with open(path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=keys, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(state.last_result)
    log_fn(f"[green]✓ Exported {len(state.last_result)} rows to {path}[/green]")


# ─── AGENT LOOP ───────────────────────────────────────────────────────────────

def _maybe_status(label: str, fn):
    with console.status(f"[dim]{label}[/dim]"):
        return fn()


def run_command(
    user_input: str,
    token: str,
    client: anthropic.Anthropic,
    messages: list[dict],
    upn: str = "",
    log_fn=None,
    confirm_fn=None,
) -> None:
    if log_fn is None:
        log_fn = console.print
    if confirm_fn is None:
        confirm_fn = lambda p: Confirm.ask(f"[yellow]{p}[/yellow]")

    messages.append({"role": "user", "content": user_input})

    while True:
        response = _maybe_status(
            "Thinking...",
            lambda: client.messages.stream(
                model="claude-opus-4-7",
                max_tokens=4096,
                thinking={"type": "adaptive"},
                output_config={"effort": "high"},
                system=[{"type": "text", "text": SYSTEM_PROMPT, "cache_control": {"type": "ephemeral"}}],
                tools=TOOLS,
                messages=messages,
            ).__enter__().get_final_message(),
        )

        for block in response.content:
            if block.type == "text" and block.text.strip():
                log_fn(Markdown(block.text))

        messages.append({"role": "assistant", "content": response.content})

        if response.stop_reason != "tool_use":
            break

        tool_results = []

        for block in response.content:
            if block.type != "tool_use":
                continue

            result: str

            if block.name == "search_m365_docs":
                result = _maybe_status(
                    f"Searching docs: {block.input['query']}",
                    lambda: search_m365_docs(block.input["query"]),
                )

            elif block.name == "read_local_file":
                result = read_local_file(block.input["path"])

            elif block.name == "execute_powershell":
                command = block.input["command"]
                description = block.input.get("description", command)
                is_write = block.input.get("is_write", False)

                log_fn(Panel(
                    f"[bold]PowerShell[/bold]\n\n{description}\n\n"
                    f"```powershell\n{command}\n```",
                    title="💻 Exchange Online",
                    border_style="magenta",
                ))

                if is_write and state.readonly:
                    log_fn("[red]⛔ Blocked — read-only mode is enabled.[/red]")
                    result = json.dumps({"blocked": True, "reason": "Read-only mode"})
                    tool_results.append({"type": "tool_result", "tool_use_id": block.id, "content": result})
                    continue

                if is_write and state.dryrun:
                    log_fn("[yellow]🔍 Dry run — not executed.[/yellow]")
                    result = json.dumps({"dry_run": True, "would_execute": command})
                    tool_results.append({"type": "tool_result", "tool_use_id": block.id, "content": result})
                    continue

                if is_write:
                    if not confirm_fn("Execute this PowerShell operation?"):
                        result = json.dumps({"cancelled": True})
                        tool_results.append({"type": "tool_result", "tool_use_id": block.id, "content": result})
                        continue

                if not exo_session.connected:
                    log_fn("[dim]Connecting to Exchange Online (browser sign-in required)...[/dim]")
                    conn_result = exo_session.connect(upn)
                    if "error" in conn_result.lower() or "not installed" in conn_result.lower():
                        result = json.dumps({"error": conn_result})
                        tool_results.append({"type": "tool_result", "tool_use_id": block.id, "content": result})
                        continue
                    log_fn("[green]✓ Connected to Exchange Online[/green]")

                result = _maybe_status(
                    "Running PowerShell command...",
                    lambda: exo_session.run(command),
                )
                audit(user_input, "POWERSHELL", command[:80], 0)

            elif block.name == "execute_graph_api":
                method = block.input["method"].upper()
                endpoint = block.input["endpoint"]
                body = block.input.get("body")
                description = block.input.get("description", f"{method} {endpoint}")

                body_section = f"\n\n```json\n{json.dumps(body, indent=2)}\n```" if body else ""
                log_fn(Panel(
                    f"[bold]{method}[/bold] `{endpoint}`\n\n{description}{body_section}",
                    title="📡 Graph API Call",
                    border_style="cyan",
                ))

                is_write = method in ("POST", "PATCH", "PUT", "DELETE")

                if is_write and state.readonly:
                    log_fn("[red]⛔ Blocked — read-only mode is enabled.[/red]")
                    result = json.dumps({"blocked": True, "reason": "Read-only mode"})
                    tool_results.append({"type": "tool_result", "tool_use_id": block.id, "content": result})
                    continue

                if is_write and state.dryrun:
                    log_fn("[yellow]🔍 Dry run — not executed.[/yellow]")
                    audit(user_input, method, endpoint, 0, dry=True)
                    result = json.dumps({"dry_run": True, "would_execute": {"method": method, "endpoint": endpoint, "body": body}})
                    tool_results.append({"type": "tool_result", "tool_use_id": block.id, "content": result})
                    continue

                if is_write:
                    if not confirm_fn("Execute this Graph API operation?"):
                        result = json.dumps({"cancelled": True, "reason": "User declined"})
                        tool_results.append({"type": "tool_result", "tool_use_id": block.id, "content": result})
                        continue

                data = _maybe_status(
                    "Calling Graph API...",
                    lambda: call_graph(token, method, endpoint, body),
                )
                audit(user_input, method, endpoint, data.get("_status_code", 0))

                if "value" in data and isinstance(data["value"], list):
                    state.last_result = data["value"]

                result = json.dumps(data, indent=2)

            else:
                result = json.dumps({"error": f"Unknown tool: {block.name}"})

            tool_results.append({"type": "tool_result", "tool_use_id": block.id, "content": result})

        messages.append({"role": "user", "content": tool_results})


# ─── BUILT-IN COMMANDS ────────────────────────────────────────────────────────

def handle_builtin(
    cmd: str,
    token: str,
    client: anthropic.Anthropic,
    messages: list[dict],
    upn: str = "",
    log_fn=None,
    confirm_fn=None,
) -> bool:
    if log_fn is None:
        log_fn = console.print
    if confirm_fn is None:
        confirm_fn = lambda p: Confirm.ask(f"[yellow]{p}[/yellow]")

    parts = cmd.strip().split()
    keyword = parts[0].lower()

    if keyword == "help":
        log_fn(HELP_TEXT)
        return True

    if keyword == "clear":
        messages.clear()
        log_fn("[green]✓ Conversation context cleared.[/green]")
        return True

    if keyword == "readonly":
        state.readonly = not state.readonly
        status = "[red]ON[/red]" if state.readonly else "[green]OFF[/green]"
        log_fn(f"Read-only mode: {status}")
        return True

    if keyword == "dryrun":
        state.dryrun = not state.dryrun
        status = "[yellow]ON[/yellow]" if state.dryrun else "[green]OFF[/green]"
        log_fn(f"Dry-run mode: {status}")
        return True

    if keyword == "export":
        filename = parts[1] if len(parts) > 1 else "export.csv"
        export_last_result(filename, log_fn=log_fn)
        return True

    if keyword == "licenses":
        run_command(
            "Show me a license inventory dashboard: list all license SKUs with the product name, "
            "total purchased seats, consumed seats, and available seats remaining.",
            token, client, messages, upn, log_fn=log_fn, confirm_fn=confirm_fn,
        )
        return True

    if keyword == "offboard" and len(parts) > 1:
        email = parts[1]
        run_command(
            f"Offboard the user {email}: revoke their sign-in sessions, disable their account, "
            "remove all licenses, and report when done.",
            token, client, messages, upn, log_fn=log_fn, confirm_fn=confirm_fn,
        )
        return True

    if keyword == "report" and len(parts) > 1 and parts[1].lower() == "passwords":
        run_command(
            "Show me a report of all users whose passwords are expiring within the next 30 days, "
            "including their display name, email, and password expiry date.",
            token, client, messages, upn, log_fn=log_fn, confirm_fn=confirm_fn,
        )
        return True

    if keyword == "disconnect" and len(parts) > 1 and parts[1].lower() == "exchange":
        if exo_session.connected:
            exo_session.disconnect()
            log_fn("[green]✓ Disconnected from Exchange Online[/green]")
        else:
            log_fn("[dim]Not connected to Exchange Online[/dim]")
        return True

    return False


# ─── CONFIG ───────────────────────────────────────────────────────────────────

CONFIG_PATH = CLI_DIR / "config.json"


def load_config() -> dict:
    if CONFIG_PATH.exists():
        return json.loads(CONFIG_PATH.read_text())
    return {}


def save_config(config: dict) -> None:
    CLI_DIR.mkdir(exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(config, indent=2))


def prompt_required(label: str, **kwargs) -> str:
    while True:
        value = Prompt.ask(label, **kwargs).strip()
        if value:
            return value
        console.print("[red]This field is required.[/red]")


def setup_wizard() -> dict:
    console.print(
        Panel(
            "[bold]Welcome to NoPilot 365[/bold]\n\n"
            "Let's get you set up. You'll only need to do this once.\n\n"
            "[dim]You'll need:\n"
            "  • An Anthropic API key (console.anthropic.com)\n"
            "  • An Azure AD App Registration (portal.azure.com)[/dim]",
            title="⚙️  First-time Setup",
            border_style="blue",
        )
    )

    console.print("\n[bold]Step 1 of 3 — Anthropic API Key[/bold]")
    console.print("[dim]Get yours at console.anthropic.com → API Keys[/dim]")
    anthropic_key = prompt_required("Anthropic API key", password=True)

    console.print("\n[bold]Step 2 of 3 — Azure App Client ID[/bold]")
    console.print("[dim]Found in Azure portal → App registrations → your app → Overview[/dim]")
    client_id = prompt_required("Client ID")

    console.print("\n[bold]Step 3 of 3 — Azure Tenant ID[/bold]")
    console.print("[dim]Found in the same Overview page, just below Client ID[/dim]")
    tenant_id = prompt_required("Tenant ID")

    config = {"anthropic_api_key": anthropic_key, "client_id": client_id, "tenant_id": tenant_id}
    save_config(config)
    console.print("\n[green]✓ Configuration saved. You won't need to enter these again.[/green]\n")
    return config


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main() -> None:
    config = load_config()
    if not all(k in config for k in ("anthropic_api_key", "client_id", "tenant_id")):
        config = setup_wizard()

    console.print("[dim]Connecting to Microsoft 365...[/dim]")

    try:
        token = authenticate(config["client_id"], config["tenant_id"])
    except Exception as e:
        console.print(f"[red]Authentication failed: {e}[/red]")
        sys.exit(1)

    try:
        me = call_graph(token, "GET", "/me?$select=userPrincipalName,displayName", None)
        upn = me.get("userPrincipalName", "")
        display_name = me.get("displayName", "")
    except Exception:
        upn = ""
        display_name = ""

    client = anthropic.Anthropic(api_key=config["anthropic_api_key"])
    messages: list[dict] = []

    atexit.register(exo_session.disconnect)

    console.print(Panel(
        f"[dim]Signed in as [bold]{display_name}[/bold] ({upn})\n"
        "Type [bold]help[/bold] for available commands, or describe what you want to do.[/dim]",
        title="[bold blue]NoPilot 365[/bold blue]",
        border_style="blue",
    ))

    while True:
        try:
            user_input = Prompt.ask("\n[bold green]M365>[/bold green]").strip()
        except (EOFError, KeyboardInterrupt):
            console.print("\n[dim]Goodbye![/dim]")
            break

        if not user_input:
            continue

        if user_input.lower() in ("exit", "quit", "q"):
            console.print("[dim]Goodbye![/dim]")
            break

        try:
            handled = handle_builtin(user_input, token, client, messages, upn)
            if not handled:
                run_command(user_input, token, client, messages, upn)
        except Exception as e:
            console.print(f"[red]Error: {e}[/red]")


if __name__ == "__main__":
    main()

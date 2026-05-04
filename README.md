# NoPilot 365

Natural language Microsoft 365 administration. Describe what you want to do in plain English — NoPilot translates it into Microsoft Graph API calls and Exchange Online PowerShell commands, asks for your confirmation on write operations, and executes them.

## What it does

- Authenticates to Microsoft 365 using device flow (no service account needed)
- Translates natural language requests into Graph API calls or Exchange Online PowerShell
- Requires user confirmation before any write operation
- Supports bulk operations via local CSV files
- Logs every write operation to `~/.m365cli/audit.log`
- Read-only and dry-run modes for safe exploration

## Requirements

- Python 3.11+
- [`uv`](https://github.com/astral-sh/uv) (recommended) or `pip`
- An [Anthropic API key](https://console.anthropic.com/)
- A Microsoft 365 tenant with an Azure App Registration

## Setup

### 1. Clone the repo

```bash
git clone https://github.com/nickrinke/NoPilot365.git
cd NoPilot365
```

### 2. Create an app registration in Azure

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name it anything (e.g. `NoPilot365`), leave redirect URI blank, click **Register**
3. Note the **Application (client) ID** and **Directory (tenant) ID** from the Overview page
4. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**, and add:
   - `User.ReadWrite.All`
   - `Group.ReadWrite.All`
   - `Directory.ReadWrite.All`
   - `Organization.Read.All`
5. Click **Grant admin consent**
6. Under **Authentication**, enable **Allow public client flows**

### 3. Run the app

**With uv (recommended):**
```bash
uv run m365_admin.py
```

**With pip:**
```bash
pip install -r requirements.txt
python m365_admin.py
```

On first run, a setup wizard will prompt for your Anthropic API key, Client ID, and Tenant ID. These are saved to `~/.m365cli/config.json` — you won't need to enter them again.

## Usage

Just describe what you want at the prompt:

```
M365> Show all users who haven't signed in for 90 days
M365> Assign an E3 license to jsmith@contoso.com
M365> Create a shared mailbox for the finance team
M365> Reset the password for sarah@contoso.com
M365> Offboard john@contoso.com
```

### Built-in commands

| Command | Description |
|---|---|
| `help` | Show available commands |
| `clear` | Reset conversation context |
| `readonly` | Toggle read-only mode (blocks all write operations) |
| `dryrun` | Toggle dry-run mode (shows what would happen without executing) |
| `export [file]` | Export last result to CSV (default: `export.csv`) |
| `licenses` | Show license inventory dashboard |
| `offboard <email>` | Run full offboarding workflow |
| `report passwords` | Users with passwords expiring in 30 days |
| `disconnect exchange` | Disconnect the Exchange Online PowerShell session |
| `exit` | Quit |

## Example Output

```
╭─ NoPilot 365 ──────────────────────────────────────╮
│ Signed in as Nick Rinke (nick@contoso.com)          │
│ Type help for available commands, or describe what  │
│ you want to do.                                     │
╰─────────────────────────────────────────────────────╯

M365> Disable the account for jsmith@contoso.com

╭─ 📡 Graph API Call ─────────────────────────────────╮
│ PATCH `/users/jsmith@contoso.com`                   │
│                                                     │
│ Disable sign-in for jsmith@contoso.com              │
│                                                     │
│ ```json                                             │
│ { "accountEnabled": false }                         │
│ ```                                                 │
╰─────────────────────────────────────────────────────╯

Execute this Graph API operation? [y/n]: y
✓ Account disabled.
```

## Exchange Online PowerShell

For tasks not available in the Graph API — message traces, transport rules, mail flow connectors, litigation hold — NoPilot uses Exchange Online PowerShell via `pwsh`.

Install PowerShell 7 if needed:
```bash
brew install powershell
```

NoPilot will install the `ExchangeOnlineManagement` module and open a browser sign-in automatically the first time a PowerShell command is needed.

## Audit log

Every write operation is logged to `~/.m365cli/audit.log` as newline-delimited JSON:

```json
{"timestamp": "2026-05-03T18:00:00Z", "command": "Disable jsmith", "method": "PATCH", "endpoint": "/users/jsmith@contoso.com", "status_code": 204, "dry_run": false}
```

## Tech Stack

- [MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-python) — Microsoft authentication and token caching
- [Requests](https://requests.readthedocs.io/) — HTTP client for Graph API calls
- [Anthropic Python SDK](https://github.com/anthropics/anthropic-sdk-python) — Claude integration
- [Rich](https://github.com/Textualize/rich) — Terminal UI

## License

MIT

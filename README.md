# NoPilot 365

Natural language Microsoft 365 administration via the Graph API and Exchange Online PowerShell. Describe what you want to do in plain English — NoPilot translates it into API calls and runs them with your confirmation.

## Requirements

- Python 3.11+
- [`uv`](https://github.com/astral-sh/uv) (recommended) or `pip`
- An [Anthropic API key](https://console.anthropic.com/)
- An Azure AD App Registration with the permissions below

## Azure App Registration

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

## Setup

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

```
M365> Show all users who haven't signed in for 90 days
M365> Assign an E3 license to jsmith@contoso.com
M365> Offboard sarah@contoso.com
M365> Create a shared mailbox for the finance team
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

## Exchange Online PowerShell

For tasks not available in the Graph API (message traces, transport rules, mail flow connectors, litigation hold), NoPilot uses Exchange Online PowerShell via `pwsh`.

Install PowerShell 7 if needed:
```bash
brew install powershell
```

NoPilot will install the `ExchangeOnlineManagement` module and open a browser sign-in automatically the first time a PowerShell command is needed.

## Audit log

Every write operation is logged to `~/.m365cli/audit.log` as newline-delimited JSON.

## License

MIT

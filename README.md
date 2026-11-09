# outlook-desktop-mcp

[![PyPI](https://img.shields.io/pypi/v/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Python](https://img.shields.io/pypi/pyversions/outlook-desktop-mcp)](https://pypi.org/project/outlook-desktop-mcp/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-blue)]()
[![Security](https://img.shields.io/badge/security-Touch%20ID%20Gated-brightgreen)]()
[![License: AGPL-3.0](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](LICENSE)

**Turn your running Outlook Desktop into an MCP server.** You do not need Microsoft Graph API, Entra app registration, or OAuth tokens. The server connects directly to your local Outlook session.

Any Model Context Protocol (MCP) client (Claude Code, Claude Desktop, Cursor) can read emails, search mailboxes, manage calendars, and organize tasks through your existing Outlook application.

> **Upstream Attribution & Extensions**:
> This project is forked from the original work by **Andreas Martin Aanerud** ([aaanerud/outlook-desktop-mcp](https://github.com/aaanerud/outlook-desktop-mcp)).
> Maintained and extended by **Grobler Engineering** ([hemelbesem90s/outlook-desktop-mcp](https://github.com/hemelbesem90s/outlook-desktop-mcp)) to add native macOS Touch ID biometric security gating for mutating operations, delete message tools, and zero-trust verification.

## Biometric Security Gate (macOS)

AI agents should read data freely but ask for explicit human permission before altering your mailbox. On macOS, this server integrates native **Apple Touch ID** via the `LocalAuthentication` framework:

- **Frictionless Reading**: Read-only tools (`list_emails`, `read_email`, `search_emails`, `list_events`, `list_folders`, `list_tasks`) run immediately without biometric friction.
- **Biometric Gate on Mutations**: Mutating operations trigger a native Touch ID prompt with full operational context:
  - `send_email` and `reply_email` display the recipient and subject.
  - `create_event`, `create_calendar_event`, `create_meeting`, and `update_event` display event details and attendees.
  - `delete_event`, `delete_calendar_event`, `delete_message`, `delete_email`, and `delete_task` show the specific item ID before removal.
- **Fail-Safe Fallback**: If biometric evaluation is unavailable, the server falls back to an explicit macOS system confirmation dialog.
- **CI / Testing Override**: Set `OUTLOOK_MCP_DISABLE_BIOMETRICS=1` in your environment to bypass prompts during automated test runs.

## Quick Start

**1. Install** (requires Python 3.12+):

```bash
pip install outlook-desktop-mcp
```

**2. Register with Claude Code:**

```bash
claude mcp add outlook-desktop -- outlook-desktop-mcp
```

**3. Open Outlook and start a Claude Code session.** Tools become available immediately.

## How It Works — Platform Routing

When the server starts, it checks your operating system and routes commands to the appropriate bridge:

```
                        outlook-desktop-mcp starts
                                  |
                          sys.platform check
                         /                  \
                   "win32"                "darwin"
                      |                      |
              ┌───────┴────────┐    ┌────────┴────────┐
              │  server.py     │    │  server_mac.py   │
              │  COM Bridge    │    │  AppleScript     │
              │  (29 tools)    │    │  + Touch ID Gate │
              │                │    │  (27 tools)      │
              └───────┬────────┘    └────────┬─────────┘
                      |                      |
              OUTLOOK.EXE via         Microsoft Outlook
              COM / STA thread        via osascript
                      |                      |
              Exchange / M365         Exchange / M365
```

**Both paths use your locally running Outlook app and its existing authenticated session.** No cloud credentials, no Graph API tokens — the server inherits whatever account Outlook is signed into.

### Why two paths?

Windows Outlook (Classic) exposes a rich COM automation interface — the Outlook Object Model (`MSOUTL.OLB`). This has been the standard way to programmatically control Outlook on Windows for over 20 years. It provides deep access to mail rules, categories, MAPI properties, and the full folder hierarchy.

Mac Outlook does not support COM. Instead, it exposes an AppleScript dictionary that can be driven via the `osascript` command. The AppleScript interface covers the core operations — email, calendar, tasks — but does not expose rules, categories, or certain advanced MAPI features. This is a limitation of what Microsoft chose to include in Outlook for Mac's scripting dictionary, not a limitation of this project.

The server is structured as two parallel implementations with identical tool names and signatures, so MCP clients see the same interface regardless of platform. Tools that are not available on a given platform are simply not registered.

## Requirements

### Windows

- **Outlook Desktop (Classic)** — the `OUTLOOK.EXE` that comes with Microsoft 365 / Office. The new "modern" Outlook (`olk.exe`) does **not** support COM
- **Python 3.12+** (x64 or ARM64)
- **Outlook must be running** when the MCP server starts

Both x64 and ARM64 Windows are supported. On ARM64, all dependencies (`pywin32`, `mcp`, `pydantic-core`, `cryptography`, `cffi`, `rpds-py`) have prebuilt `win_arm64` wheels — see the [ARM64 install notes](#arm64-windows) below for the one extra `pip` flag you need.

#### Outlook "Programmatic Access" security prompts

When the MCP server first touches Outlook's COM API, Outlook may show a dialog: *"A program is trying to access email address information stored in Outlook"* (the Object Model Guard / Programmatic Access prompt). This is Outlook's protection against malicious automation.

You have three options:

1. **Click "Allow access for 10 minutes"** every time you start a session. Fine for casual use.
2. **Get the antivirus status to "Valid"** in *File > Options > Trust Center > Trust Center Settings > Programmatic Access*. When that line reads `Valid`, the "Never warn me about suspicious activity" radio becomes selectable and the prompts go away. On most personal machines with current Defender this works out of the box.
3. **Apply the registry policy** in [`docs/suppress-outlook-oom-prompts.reg`](docs/suppress-outlook-oom-prompts.reg). From an **elevated** PowerShell or Command Prompt (Win+X → *Terminal (Admin)*), run:

   ```powershell
   reg import "C:\path\to\outlook-desktop-mcp\docs\suppress-outlook-oom-prompts.reg"
   ```

   Then fully quit Outlook (check Task Manager for stray `OUTLOOK.EXE` processes) and reopen it. This writes `AdminSecurityMode=3` and approves all `PromptOOM*` categories under `HKLM\Software\Policies\Microsoft\Office\16.0\Outlook\Security`, which Outlook honors regardless of AV status. On Intune/MDM-managed corporate devices, `HKCU\Software\Policies\...\Outlook` is locked and the HKLM keys may be overwritten on next policy sync — if `reg import` fails or the prompts come back, ask IT to push the equivalent settings via Group Policy.

> **Windows 11 ARM64 note:** Defender does not register with Outlook's `IOfficeAntiVirus` interface on ARM64, so Trust Center shows *"Antivirus status: Invalid"* and the *"Never warn me about suspicious activity"* radio stays greyed out **even when Outlook is launched as Administrator**. Option 2 is unavailable on ARM64; the `reg import` from option 3 is the only durable suppression path.

### macOS

- **Microsoft Outlook for Mac** — version 16.x or later
- **Python 3.12+**
- **Outlook must be running** when the MCP server starts

#### Required macOS permissions

The first time a tool runs, macOS will show **two permission prompts** that you must approve:

1. **Privacy & Automation** — a system dialog asks: *"python3.12 wants to control Microsoft Outlook"*. Click **Allow** to let the server send AppleScript commands to Outlook.

2. **Accessibility** — to read your Exchange/M365 inbox, the server uses macOS UI scripting (System Events). This requires Accessibility access for `python3.12`:
   - Open **System Settings > Privacy & Security > Accessibility**
   - Find **python3.12** in the list (it appears after the first prompt)
   - Toggle it **on**

   Without Accessibility enabled, calendar, tasks, and local folder tools will work, but listing Exchange inbox messages will return empty results.

Both permissions are one-time setup — macOS remembers them for future sessions.

## Available Tools by Platform

### Email

| Tool | Windows | macOS | Biometric Gate | Description |
|------|:-------:|:-----:|:--------------:|-------------|
| `send_email` | yes | yes | **Touch ID** | Send an email with To/CC/BCC, plain text or HTML body |
| `reply_email` | yes | yes | **Touch ID** | Reply or reply-all, preserving the conversation thread |
| `delete_message` | — | yes | **Touch ID** | Permanently delete an email message by numeric ID |
| `delete_email` | — | yes | **Touch ID** | Alias for `delete_message` |
| `list_emails` | yes | yes | — | List recent emails from any folder, with optional unread filter |
| `read_email` | yes | yes | — | Read full email content by entry ID or subject search |
| `search_emails` | yes | yes | — | Full-text search across email subjects and bodies |
| `mark_as_read` | yes | yes | — | Mark a specific email as read |
| `mark_as_unread` | yes | yes | — | Mark a specific email as unread |
| `move_email` | yes | yes | — | Move an email to Archive, Trash, or any folder |
| `list_folders` | yes | yes | — | Browse the folder hierarchy with item counts |

### Calendar

| Tool | Windows | macOS | Biometric Gate | Description |
|------|:-------:|:-----:|:--------------:|-------------|
| `create_event` | yes | yes | **Touch ID** | Create a personal calendar appointment |
| `create_calendar_event`| — | yes | **Touch ID** | Alias for `create_event` |
| `create_meeting` | yes | yes | **Touch ID** | Create a meeting and send invitations to attendees |
| `update_event` | yes | yes | **Touch ID** | Modify event properties. If attendees exist, halts and requires `confirm_send_updates=True` to prevent unexpected emails |
| `delete_event` | yes | yes | **Touch ID** | Delete an appointment or cancel a meeting |
| `delete_calendar_event`| — | yes | **Touch ID** | Alias for `delete_event` |
| `list_events` | yes | yes | — | List upcoming events within a date range |
| `get_event` | yes | yes | — | Read full event details by entry ID |
| `search_events` | yes | yes | — | Search calendar events by keyword within a date range |
| `respond_to_meeting` | yes | — | — | Accept, decline, or tentatively accept a meeting invite |

### Tasks

| Tool | Windows | macOS | Biometric Gate | Description |
|------|:-------:|:-----:|:--------------:|-------------|
| `delete_task` | yes | yes | **Touch ID** | Remove a task |
| `list_tasks` | yes | yes | — | List pending or completed tasks, sorted by due date |
| `get_task` | yes | yes | — | Read full task details including body and completion status |
| `create_task` | yes | yes | — | Create a new task with subject, due date, importance |
| `complete_task` | yes | yes | — | Mark a task as complete |

### Attachments

| Tool | Windows | macOS | Biometric Gate | Description |
|------|:-------:|:-----:|:--------------:|-------------|
| `list_attachments` | yes | yes | — | List all attachments on an email or calendar event |
| `save_attachment` | yes | yes | — | Download an attachment to a local directory |

### Categories, Rules, Out of Office (Windows only)

These tools rely on COM-specific APIs (MAPI property accessors, the Rules object model, and the Categories collection) that Outlook for Mac does not expose through AppleScript.

| Tool | Windows | macOS | Description |
|------|:-------:|:-----:|-------------|
| `list_categories` | yes | — | List all available color categories in Outlook |
| `set_category` | yes | — | Set or clear categories on any email, event, or task |
| `list_rules` | yes | — | List all mail rules with enabled/disabled status |
| `toggle_rule` | yes | — | Enable or disable a mail rule by name |
| `get_out_of_office` | yes | — | Check whether Out of Office auto-reply is on or off |

**Total: 29 tools on Windows, 27 tools on macOS.**

## Architecture Details

### Windows: COM Bridge (`com_bridge.py`)

All Outlook COM operations run on a dedicated thread using the Single-Threaded Apartment (STA) model, as required by COM. The async MCP event loop dispatches tool calls to this thread via a queue and awaits results, keeping COM threading rules respected and the MCP protocol non-blocking.

```
MCP tool call (async)
  → bridge.call(func, args)
    → queued to STA thread
      → func(outlook, namespace, args) executes on COM thread
    → result returned via threading.Event
  → JSON response back to MCP client
```

Each tool's inner function receives the live `Outlook.Application` and `MAPI.Namespace` COM objects and works directly with the Outlook Object Model — `GetItemFromID`, `CreateItem`, `Items.Restrict` with DASL filters, and so on.

### macOS: AppleScript Bridge (`applescript_bridge.py`)

Each tool call builds an AppleScript string and executes it as a subprocess via `osascript`. There is no persistent connection — every call is stateless.

```
MCP tool call (async)
  → build AppleScript string
    → asyncio.create_subprocess_exec("osascript", "-e", script)
    → parse stdout text into structured data
  → JSON response back to MCP client
```

Each tool constructs a single AppleScript that fetches all needed data in one `osascript` call (no per-message subprocess loops). Results come back as delimited text, which the server parses into the same JSON structure the Windows server produces.

**Key differences from Windows:**

- Entry IDs on macOS are **numeric** (e.g. `42`), not hex strings. They identify items within their folder context.
- Folder references use AppleScript's **locale-independent keywords** (`inbox`, `sent items`, `drafts`, `deleted items`) rather than localized folder names.
- Search uses AppleScript's `whose` clause (e.g. `messages whose subject contains "query"`) instead of DASL filters.
- User input is escaped for safe embedding in AppleScript strings to prevent script injection.

## Install from Source

### Windows (x64)

```bash
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python -m venv .venv
.venv\Scripts\activate
pip install pywin32 "mcp[cli]" -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

Register from source using the launcher script:

```bash
claude mcp add outlook-desktop -- powershell.exe -Command "& 'C:\path\to\outlook-desktop-mcp\outlook-desktop-mcp.cmd' mcp"
```

### Windows (ARM64)

The `[cli]` extra of `mcp` transitively pulls in `cryptography`, and pip's default resolver may pick a version that lacks a `win_arm64` wheel — which then fails to build because it requires a Rust toolchain plus OpenSSL. Install without the `cli` extra and force wheels-only resolution:

```powershell
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
& "C:\Program Files\Python312-arm64\python.exe" -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install --only-binary=:all: pywin32 mcp
python -m pip install --no-deps -e .
python .venv\Scripts\pywin32_postinstall.py -install
```

The base `mcp` package is sufficient for running the stdio server — the `[cli]` extra is only needed for the `mcp` developer CLI tools (`mcp dev`, `mcp inspector`), which aren't used at runtime.

Register from source the same way as x64:

```bash
claude mcp add outlook-desktop -- powershell.exe -Command "& 'C:\path\to\outlook-desktop-mcp\outlook-desktop-mcp.cmd' mcp"
```

### macOS

```bash
git clone https://github.com/Aanerud/outlook-desktop-mcp.git
cd outlook-desktop-mcp
python3 -m venv .venv
source .venv/bin/activate
pip install "mcp[cli]" -e .
```

Register from source:

```bash
claude mcp add outlook-desktop -- /path/to/outlook-desktop-mcp/.venv/bin/python -m outlook_desktop_mcp
```

## Usage Examples

Once registered, just talk to Claude naturally:

- *"Show me my 10 most recent inbox emails"*
- *"Read the email from Taylor about MLADS"*
- *"Send an email to alice@example.com about the project update"*
- *"What's on my calendar this week?"*
- *"Create a meeting with bob@example.com tomorrow at 2pm for 30 minutes"*
- *"Save the attachment from that email to my Downloads folder"*
- *"Create a task to review the quarterly report, due Friday, high importance"*
- *"Mark that email as read and move it to archive"*

Windows-only examples:

- *"What categories do I have? Set this email to 'Follow-up'"*
- *"List my mail rules"*
- *"Am I set as Out of Office?"*

## Why Not Microsoft Graph?

| | Microsoft Graph | outlook-desktop-mcp |
|---|---|---|
| Entra app registration | Required | Not needed |
| Admin consent | Required for mail permissions | Not needed |
| OAuth token management | You handle refresh tokens | Not needed |
| Tenant configuration | Required | Not needed |
| Works offline / cached | No | Yes (reads from local cache) |
| Setup time | 30-60 minutes | 2 minutes |
| Auth requirement | **Your own OAuth flow** | **Outlook is open** |

## Project Structure

```
outlook-desktop-mcp/
  src/outlook_desktop_mcp/
    entrypoint.py            # Platform detection → routes to correct server
    server.py                # Windows MCP server (29 tools, COM automation)
    server_mac.py            # macOS MCP server (27 tools, AppleScript + Touch ID)
    biometrics.py            # Native macOS Touch ID gate (Swift LocalAuthentication)
    com_bridge.py            # Async-to-COM threading bridge (Windows)
    applescript_bridge.py    # Async osascript execution (macOS)
    tools/
      _folder_constants.py   # Outlook enums and constants (Windows)
    utils/
      formatting.py          # Email/event/task data extraction (Windows)
      errors.py              # COM error formatting (Windows)
      applescript_helpers.py # AppleScript escaping, date formatting (macOS)
  tests/
    test_biometrics_mac.py   # Unit tests for Touch ID gating, fallbacks, and tool blocking
    phase1_com_test.py       # Email COM validation (Windows)
    phase3_mcp_test.py       # Email MCP test (Windows)
    calendar_com_test.py     # Calendar COM validation (Windows)
    calendar_mcp_test.py     # Calendar MCP test (Windows)
    extras_com_test.py       # Tasks/attachments/categories/rules/OOF COM test (Windows)
    extras_mcp_test.py       # Tasks/attachments/categories/rules/OOF MCP test (Windows)
  outlook-desktop-mcp.cmd   # Windows launcher script
  pyproject.toml
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for the branching strategy and development setup.

## Attribution & License

- **Original Author**: Andreas Martin Aanerud ([aaanerud/outlook-desktop-mcp](https://github.com/aaanerud/outlook-desktop-mcp))
- **Biometric Extensions**: Jacob Grobler / [Grobler Engineering](https://grobler.nl) ([hemelbesem90s/outlook-desktop-mcp](https://github.com/hemelbesem90s/outlook-desktop-mcp))
- **License**: GNU Affero General Public License v3.0 ([AGPL-3.0](LICENSE)).

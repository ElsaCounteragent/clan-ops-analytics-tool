# Clan Operations & Analytics Automation Tool

**Excel/VBA + Python automation that replaced repetitive manual coordination for a 50+ member online community.**

I built this to solve a real workflow problem: running a large, active Clash of Clans group meant hours of copy-pasting assignments, hand-checking qualification rules, cleaning reports, and manually verifying new Discord members. Over several months, I turned that process into a structured internal operations system spanning Excel, VBA, REST APIs, and a Python Discord bot.

> Personal project. Not affiliated with Supercell or Discord.

**Core focus:** process automation · VBA · Python (`asyncio`) · REST API integration · webhook messaging · data validation · environment-based configuration

---

## The Problem

Managing the group manually created recurring friction:

- Assignments were copied and updated by hand every cycle.
- Member performance was tracked manually across multiple events.
- Qualification rules were checked by eye, which was slow and error-prone.
- Public reports needed manual cleanup to hide admin notes and formulas.
- Discord announcements were drafted and sent by hand, with a risk of spamming the channel.
- New members had to be manually asked to provide their in-game username, which took time and was unsafe.

As the roster changed week to week, this became slow and easy to get wrong.

## The Solution

The system centralizes operational data and automates each repetitive step. A typical cycle:

1. Update the roster or event data.
2. Run the assignment generator.
3. Review generated assignments and performance flags.
4. Generate a clean public report.
5. Send grouped Discord notifications via webhook.
6. New members self-verify through the Discord bot, with no moderator needed.

The goal isn't to remove human judgment. It's to remove copy-paste, manual formatting, and admin overhead so decisions can happen faster.

---

## Core Features

### 1. Assignment Automation (VBA)
- Bulk assignment generation in a single run
- Handles Unicode names and special characters
- Supports variable roster sizes without restructuring the workbook
- Reusable worksheet structure with prompt-driven defaults

### 2. Performance Tracking (Excel + VBA)
- Dynamic arrays (`UNIQUE`, `FILTER`) for roster deduplication and filtering
- `XLOOKUP` to match roster entries to historical records
- `COUNTIFS` for participation and performance metrics
- Rule-based logic that automatically flags members who meet qualification conditions across multiple events

### 3. Public Report Generation (VBA)
- Duplicates reporting sheets into a new workbook
- Converts formulas to static values
- Hides admin columns and removes unused tabs
- Keeps internal logic and notes separate from what gets shared

### 4. Discord Webhook Notifications (VBA)
- JSON payload construction and HTTP requests
- Grouped, priority-ordered assignment messages
- Batched mentions to avoid flooding the channel
- Delays between requests to respect rate limits
- Optional scheduled follow-ups

### 5. Discord Verification Bot (Python)
A standalone bot (`bot/verify_bot.py`) that verifies players against the official Clash of Clans API and syncs their Discord profile.

**How it works**
1. An admin runs `!spawn_verify` to post a persistent verification panel.
2. A member clicks **Verify Account** and fills in a modal with their player tag and in-game API token.
3. The bot calls the Clash of Clans `verifytoken` endpoint to confirm the user actually owns the account.
4. On success, it fetches the player profile, assigns the verified role, and sets the member's nickname to their in-game name.

**Technical implementation**
- **Async architecture:** Built on `discord.py` and `aiohttp` so API calls never block the bot.
- **Persistent views:** The verification panel uses `timeout=None` with fixed `custom_id`s, registered in `setup_hook`, so buttons keep working after a bot restart.
- **Idempotent updates:** Re-verifying doesn't duplicate roles. The bot detects whether the member is already verified, only needs a nickname refresh, or needs full setup.
- **Input handling:** Normalizes tags (uppercase, adds `#`), URL-encodes them for the API, and truncates names to Discord's 32-character nickname limit.
- **Defensive error handling:** Distinct handling for API failures (including 403 errors from an outdated IP whitelist), network errors, Discord permission errors, and unexpected exceptions, each with a clear ephemeral message and a log entry.
- **Privacy-conscious UX:** All responses are ephemeral, so tags and tokens are never exposed in public channels.
- **Fail-fast configuration:** The bot exits with a critical log if required environment variables are missing.

---

## Technical Stack

| Area | Technologies |
|------|-------------|
| Core Automation | Excel, VBA |
| Discord Bot | Python 3, `discord.py`, `aiohttp`, `asyncio` |
| Data & Analytics | Dynamic arrays, `XLOOKUP`, `COUNTIFS`, `Scripting.Dictionary` |
| APIs & Web | REST APIs (Supercell, Discord), JSON payloads, `WinHttpRequest`, `curl.exe`, webhooks |
| Practices | Environment-based secrets, logging, idempotent operations, rate-limit handling |
| Environment | Windows, Discord Developer Portal |

---

## Repository Structure

```text
├── vba/
│   ├── WarAssignments.bas        # Multi-phase war target optimizer
│   ├── WarApiTracker.bas         # Live war API tracker and performance logger
│   ├── CwlStrategyAssignment.bas # Bulk CWL strategy assignment tool
│   ├── PublicReportPublisher.bas # Public report generator and Discord file uploader
│   ├── DiscordNotifications.bas  # Batched Discord webhook notifications
│   ├── DatabaseSync.bas          # Master database ETL synchronization
│   └── CwlApiSync.bas            # Clan War League API data synchronization
├── bot/
│   └── verify_bot.py             # Discord verification bot (Python)
├── .env.example                  # Template for required environment variables
└── README.md
## Download
Download the latest workbook from the [Releases page](https://github.com/ElsaCounteragent/clan-ops-analytics-tool/releases).

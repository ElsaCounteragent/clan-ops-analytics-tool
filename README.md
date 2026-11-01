# Clan Operations & Analytics Automation Tool

An Excel/VBA operations tool and Python Discord bot originally built to manage logistics, performance tracking, reporting, and server administration for a 50+ member Clash of Clans group.

The project started as a way to reduce repetitive manual coordination work and evolved into a structured internal operations system. It handles roster assignments, event-based performance tracking, clean public report generation, webhook-based Discord notifications, and automated server verification.

This is a personal project. It is not affiliated with Supercell or Discord.

---

## The Problem

Managing a large active group manually created several recurring friction points:

- Assignments had to be copied and updated by hand every cycle.
- Member performance had to be tracked manually across multiple events.
- Qualification rules had to be checked by eye, which was slow and error-prone.
- Public reports had to be cleaned before sharing to hide internal admin notes and formulas.
- Discord notifications had to be drafted and sent manually without spamming the channel.
- Verifying new members and keeping Discord nicknames synchronized with in-game names had to be done manually by moderators.

When the roster changed frequently, this manual process became slow and easy to mess up.

---

## The Solution

The system centralizes operational data and automates the most repetitive steps across both Excel and Discord. A typical cycle looks like this:

1. Update the active roster or event data.
2. Run the assignment generator.
3. Review generated assignments and performance flags.
4. Generate a clean public report.
5. Send grouped Discord notifications through the webhook tool.
6. Allow new members to verify their accounts via the Discord bot.

The goal is not to remove human decision-making. It is to remove repetitive copy-paste, manual formatting, and administrative overhead so that decisions and community management can happen faster.

---

## Core Features

### 1. Assignment Automation
Generates assignments across the active roster based on the selected timeframe and operator input.
- Bulk assignment generation with a single run
- Unicode names and special character support
- Variable roster sizes without restructuring
- Reusable worksheet structure
- Default assignment population based on user prompts

### 2. Performance Tracking
Tracks participation and performance across events.
- Dynamic arrays (`UNIQUE`, `FILTER`) for roster deduplication and filtering
- `XLOOKUP` for matching roster entries to historical records
- `COUNTIFS` for participation and performance metrics
- Rule-based logic for qualification or status changes

If a member meets a defined set of performance conditions across multiple events, the tool flags or indexes them as qualified automatically.

### 3. Public Report Generation
Exports a clean public-facing version of the internal workbook.
- Duplicates active reporting sheets into a new workbook
- Converts formula-driven output into static values where appropriate
- Hides administrative columns
- Removes unused tabs
- Produces a clean file suitable for sharing

This keeps internal logic and admin-only fields separate from the public report.

### 4. Discord Webhook Notifications
Sends structured messages to Discord using webhook requests.
- JSON payload formatting
- Grouped assignment notifications
- Priority-based ordering
- Batched mentions to avoid flooding the channel
- Delays between requests to respect rate limits
- Optional scheduled follow-up messages

### 5. Discord Verification Bot (Python)
A standalone bot that manages server access and member verification using the official Clash of Clans API.
- Uses Discord UI components (Buttons and Modals) for a streamlined verification flow.
- Validates player ownership by requiring users to generate and submit a temporary in-game API token.
- Automatically assigns a verified role and updates the member's server nickname to their current in-game name.
- Handles API rate limits, HTTP errors, and Discord permission boundaries gracefully.

---

## Technical Stack

| Area | Technologies |
|------|-------------|
| Core Automation | Excel, VBA |
| Discord Bot | Python, `discord.py`, `aiohttp` |
| Data & Analytics | Dynamic arrays, `XLOOKUP`, `COUNTIFS`, `Scripting.Dictionary` |
| API & Web | REST APIs (Supercell, Discord), JSON payload construction, HTTP requests via `WinHttpRequest` and `curl.exe` |
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
│   ├── DatabaseSync.bas          # Master database synchronization
│   └── CwlApiSync.bas            # Clan War League API data synchronization
├── bot/

│   └── verify_bot.py             # Discord verification bot (Python)
├── .env.example                  # Template for required environment variables
└── README.md
```

LinkedIn: https://www.linkedin.com/in/maxim-kolosov-92b526406/

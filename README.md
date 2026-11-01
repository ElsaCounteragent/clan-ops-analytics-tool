# Clan Operations & Analytics Automation Tool

An Excel/VBA operations tool built to manage logistics, performance tracking, reporting, and communication for a 50+ member Clash of Clans group.

What started as a way to cut down repetitive manual coordination work evolved into a structured internal operations system. The tool handles roster assignments, event-based performance tracking, clean public report generation, and webhook-based Discord notifications.

This is a personal project. It is not affiliated with Supercell or Discord.

---

## The Problem

Managing a large active group manually created several recurring friction points:

- Assignments had to be copied and updated by hand every cycle.
- Member performance had to be tracked manually across multiple events.
- Qualification rules had to be checked by eye, which was slow and error-prone.
- Public reports had to be cleaned before sharing to hide internal admin notes and formulas.
- Discord notifications had to be drafted and sent one at a time.

When the roster changed frequently, this manual process became slow and easy to mess up.

---

## The Solution

The workbook centralizes operational data and automates the most repetitive steps. A typical cycle looks like this:

1. Update the active roster or event data.
2. Run the assignment generator.
3. Review generated assignments and performance flags.
4. Generate a clean public report for sharing.
5. Send grouped Discord notifications through the webhook tool.

The goal is not to remove human decision-making. It is to remove repetitive copy-paste and manual formatting work so that decisions can be made faster.

---

## Core Features

### Assignment Automation
Generates assignments across the active roster based on the selected timeframe and operator input.

- Bulk assignment generation with a single run
- Unicode names and special character support
- Variable roster sizes without restructuring
- Reusable worksheet structure
- Default assignment population based on user prompts

### Performance Tracking
Tracks participation and performance across events.

- Dynamic arrays (`UNIQUE`, `FILTER`) for roster deduplication and filtering
- `XLOOKUP` for matching roster entries to historical records
- `COUNTIFS` for participation and performance metrics
- Rule-based logic for qualification or status changes

If a member meets a defined set of performance conditions across multiple events, the tool flags or indexes them as qualified automatically.

### Public Report Generation
Exports a clean public-facing version of the internal workbook.

- Duplicates active reporting sheets into a new workbook
- Converts formula-driven output into static values where appropriate
- Hides administrative columns
- Removes unused tabs
- Produces a clean file suitable for sharing

This keeps internal logic and admin-only fields separate from the public report.

### Discord Webhook Notifications
Sends structured messages to Discord using webhook requests.

- JSON payload formatting
- Grouped assignment notifications
- Priority-based ordering
- Batched mentions to avoid flooding the channel
- Delays between requests to respect rate limits
- Optional scheduled follow-up messages

---

## Technical Stack

| Area | Technologies |
|------|-------------|
| Core | Excel, VBA |
| Data | Dynamic arrays, XLOOKUP, COUNTIFS, Scripting.Dictionary |
| Reporting | Workbook export, formula-to-value conversion, column hiding |
| API | Discord Webhooks, JSON payload construction, HTTP requests |
| Structure | Modular VBA design, reusable worksheet templates |

---

## Limitations

This is a personal operations tool, not a production SaaS product.

**Designed for:**
- Small to medium online groups
- Excel-based workflows
- Manual review supported by automation
- One operator managing the workbook

**Not designed for:**
- Multi-user concurrent editing
- Cloud-scale operations
- Sensitive financial or personal data processing
- Fully autonomous decision-making

---

## What I Learned Building This

I spent several months iterating on this tool based on real operational needs, not hypothetical ones. The main lessons were practical:

- **Making repetitive work repeatable.** If a task has to be done every week, it should be done the same way every week.
- **Designing around change.** Roster sizes shift. The tool had to handle that without breaking.
- **Separating internal from public.** Admin notes, formulas, and working logic should never leak into a shared report.
- **Handling APIs carefully.** Rate limits, batching, and graceful failure matter more than speed.
- **Building tools people actually use.** A perfect tool nobody opens is worthless. A good tool people rely on daily is valuable.

---

## Repository Structure

[confirm: Add a brief file/folder map here once you have the final script set]

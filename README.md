# Clan Operations & Analytics Automation Tool

An Excel/VBA operations tool originally built to manage logistics, performance tracking, reporting, and communication for a 50+ member Clash of Clans group.

The project started as a way to reduce repetitive manual coordination work and evolved into a structured internal operations system. It handles roster assignments, event-based performance tracking, clean public report generation, and webhook-based Discord notifications.

This is a personal project. It is not affiliated with Supercell or Discord.

---

## What the Tool Does

The tool combines Excel-based analytics with VBA automation to support recurring clan operations.

Main functions:

- Roster and assignment generation
- Event participation and performance tracking
- Rule-based status or qualification updates
- Clean public report generation
- Discord webhook notifications with batching and delays
- Scalable workbook structure for variable roster sizes

---

## The Problem

Managing a large active group manually created several recurring problems:

- Assignments had to be copied and updated by hand.
- Member performance had to be tracked across multiple events.
- Qualification rules had to be checked manually.
- Public reports had to be cleaned before sharing.
- Discord notifications had to be drafted and sent manually.
- Internal admin notes and formulas had to be hidden before publishing reports.

Doing this manually was slow and easy to mess up, especially when the roster changed frequently.

---

## The Solution

The workbook centralizes the operational data and automates the most repetitive steps.

A typical workflow looks like this:

1. Update the active roster or event data.
2. Run the assignment generator.
3. Review generated assignments and performance flags.
4. Generate a clean public report.
5. Send grouped Discord notifications through the webhook tool.

The goal is not to remove human decision-making, but to remove repetitive copy-paste and manual formatting work.

---

## Core Features

### 1. Assignment Automation

The tool can generate assignments across the active roster based on the selected timeframe and user input.

It supports:

- Bulk assignment generation
- Unicode names and special characters
- Variable roster sizes
- Reusable worksheet structure
- Default assignment population based on prompts

---

### 2. Performance Tracking

The tool tracks participation and performance across events.

It uses:

- Dynamic arrays such as UNIQUE and FILTER
- XLOOKUP for matching roster entries to historical records
- COUNTIFS for participation and performance metrics
- Rule-based logic for qualification or status changes

For example, if a member meets a defined set of performance conditions across multiple events, the tool can flag or index them as qualified.

---

### 3. Public Report Generation

The tool can export a clean public-facing version of the internal workbook.

The export process can:

- Duplicate active reporting sheets into a new workbook
- Convert formula-driven output into static values where appropriate
- Hide administrative columns
- Remove unused tabs
- Produce a clean file suitable for sharing

This keeps internal logic and admin-only fields separate from the public report.

---

### 4. Discord Webhook Notifications

The tool can send structured messages to Discord using webhook requests.

Features include:

- JSON payload formatting
- Grouped assignment notifications
- Priority-based ordering
- Batched mentions to avoid flooding the channel
- Delays between requests to respect rate limits
- Optional scheduled follow-up messages

---


## Limitations

This is a personal operations tool, not a production SaaS product.

It is designed for:

- Small to medium online groups
- Excel-based workflows
- Manual review supported by automation
- One operator managing the workbook

It is not designed for:

- Multi-user concurrent editing
- Cloud-scale operations
- Sensitive financial or personal data processing
- Fully autonomous decision-making

---

## Why This Project Matters to Me

I spent several months iterating on this tool based on real operational needs. The main lessons were:

- Making repetitive work repeatable
- Designing around changing roster sizes
- Separating internal data from public reports
- Handling API notifications carefully
- Building practical automation instead of theoretical automation

The project taught me a lot about workflow design, data hygiene, and making tools that people actually use.

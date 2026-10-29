# AppDevt Finals

LAMP Office Student Assistant Daily Time Record (DTR) automation built with UiPath Studio.

## Overview

This process asks for a reporting month, calculates the second Wednesday deadline of the following month, evaluates student DTR submissions, rebuilds the tracker, and records missing-submission reminders.

The workflow entry point is `Main.xaml`.

## Project Information

- **Project type:** UiPath Process
- **Target framework:** Windows
- **Expression language:** Visual Basic
- **Studio version:** 26.0.204
- **Entry point:** `Main.xaml`
- **Workbook:** `LAMP-DTR-Tracker.xlsx`

## Main Workflow

`Main.xaml` performs these steps:

1. Prompts for a reporting month in `MMMM yyyy` format.
2. Calculates the second Wednesday of the following month as the submission deadline.
3. Reads the `Roster`, `Intake`, `Tracker`, and `EmailLog` worksheets.
4. Matches roster students with intake records for the selected month.
5. Classifies each student as:
   - `Submitted`
   - `Submitted Late`
   - `Missing`
6. Rebuilds the `Tracker` worksheet and updates summary counts.
7. Identifies missing DTR records.
8. Checks email validity and previous reminder records.
9. Uses the Google Workspace Gmail connection activity for eligible reminders.
10. Writes the updated reminder history to `EmailLog`.

## Workbook Structure

### `Roster`

| Column | Description |
|---|---|
| `Student ID` | Student identifier |
| `Name` | Student name |
| `Email` | Student email address |

### `Intake`

| Column | Description |
|---|---|
| `Student ID` | Student identifier |
| `Reporting Month` | Month represented in `MMMM yyyy` format |
| `Date Received` | Date the signed DTR was received |

### `Tracker`

The main tracker table uses columns `A:F`:

`Student ID`, `Name`, `Email`, `Reporting Month`, `Date Received`, `Status`

The summary table uses `H:I`:

| Cell | Value |
|---|---|
| `I2` | Submitted count |
| `I3` | Submitted Late count |
| `I4` | Missing count |

### `EmailLog`

| Column | Description |
|---|---|
| `Student ID` | Student identifier |
| `Name` | Student name |
| `Email` | Recipient email address |
| `Reporting Month` | Reminder month |
| `Date Sent` | Date the reminder was recorded |
| `Email Status` | Current status, such as `Sent` |

## Dependencies

The project uses these UiPath packages:

- `UiPath.System.Activities` `26.10.0`
- `UiPath.Excel.Activities` `3.7.0-preview`
- `UiPath.GSuite.Activities` `3.12.10`
- `UiPath.Mail.Activities` `2.12.10`

## Setup

1. Open the project in UiPath Studio.
2. Confirm that `LAMP-DTR-Tracker.xlsx` is in the project directory.
3. Confirm that all four worksheets exist with the expected names and headers.
4. Restore the project dependencies.
5. Configure and verify the Google Workspace Gmail connection used by `Send Gmail Reminder (disabled by default)`.
6. Review the recipient expression and reminder conditions before any live test.

The workflow does not store Gmail passwords, OAuth secrets, or access tokens in source files.

## Email Safety

Do not run the workflow against real student recipients until the Gmail connection, recipient data, reporting month, and message body have been reviewed.

The workflow contains a connection-based Gmail activity. A controlled test should use a safe test recipient or draft mode before live dispatch. Confirm the `SaveAsDraft` setting and the `bool_EnableReminders` variable in `Main.xaml` before testing, because these settings control whether reminders can be sent.

## Validation and Build

From the project directory, validate the entry workflow:

```powershell
uip rpa validate --file-path "Main.xaml" --project-dir "." --output json
```

Build the complete project:

```powershell
uip rpa build "." --output json
```

Do not run the workflow until the validation and build commands complete without errors and the email safety checks have been completed.

## Git and Local Files

The repository `.gitignore` excludes generated UiPath output and local AI-assistant files, including:

- `.local/`
- `.claude/`
- `AGENTS.md`
- `CLAUDE.md`
- Other configured local AI workspace folders

Project source, workflow files, workbook files, settings, bindings, and documentation remain available for version control.

## Project Files

- `Main.xaml` — process entry point
- `project.json` — project manifest and dependencies
- `project.uiproj` — Studio project metadata
- `LAMP-DTR-Tracker.xlsx` — source and output workbook
- `.objects/` — Object Repository metadata
- `.settings/` — project settings
- `bindings_v2.json` — connection binding metadata
- `entry-points.json` — entry-point metadata
- `rebuild-instructions.md` — workflow construction notes
- `README.md` — project documentation

# Main.xaml Current-State Instructions

## Project

- **Project:** AppDevt Finals
- **Root:** `C:\Users\justi\Documents\UiPath\AppDevt Finals`
- **Workflow:** `Main.xaml`
- **Framework:** Windows
- **Expression language:** Visual Basic
- **Workbook:** `LAMP-DTR-Tracker.xlsx`
- **Purpose:** Rebuild the LAMP Office monthly Student Assistant DTR tracker and process missing-DTR reminders.

These instructions describe the current `Main.xaml` implementation. They are not a greenfield construction specification.

## Dependencies

The current `project.json` declares:

- `UiPath.System.Activities` `[26.10.0]`
- `UiPath.Excel.Activities` `[3.7.0-preview]`
- `UiPath.GSuite.Activities` `[3.12.10]`
- `UiPath.Mail.Activities` `[2.12.10]`

The Gmail activity is `UiPath.GSuite.Activities.SendEmailConnections` and uses the Google Workspace connection service.

## Workbook Contract

`Main.xaml` resolves the workbook with:

```vb
System.IO.Path.Combine(Environment.CurrentDirectory, "LAMP-DTR-Tracker.xlsx")
```

Required worksheets and columns:

### `Roster`

- `Student ID`
- `Name`
- `Email`

### `Intake`

- `Student ID`
- `Reporting Month`
- `Date Received`

### `Tracker`

The workflow reads headers from `A1:F1`, clears the in-memory rows, and writes rebuilt rows beginning at `A2`.

Expected columns:

`Student ID`, `Name`, `Email`, `Reporting Month`, `Date Received`, `Status`

Summary counts are written to:

- `I2` — submitted count
- `I3` — submitted-late count
- `I4` — missing count

### `EmailLog`

The workflow reads from `A1`, preserves the loaded table schema, clears worksheet rows `A2:F1048576`, and writes the updated table beginning at `A2`.

Expected columns:

`Student ID`, `Name`, `Email`, `Reporting Month`, `Date Sent`, `Email Status`

## Current Execution Sequence

The root `Main Sequence` performs the following steps:

1. Repeats an `Input Dialog` until the reporting month is valid.
2. Trims the input and validates it with `DateTime.TryParseExact` using:
   - format: `MMMM yyyy`
   - `InvariantCulture`
   - `AllowWhiteSpaces`
3. Shows an invalid-format message and changes the next prompt when validation fails. It does not throw an exception for invalid user input.
4. Parses the validated month with `DateTime.ParseExact`.
5. Calculates the first day of the following month.
6. Calculates the second Wednesday of the following month.
7. Resolves the workbook path from `Environment.CurrentDirectory`.
8. Opens an `ExcelProcessScopeX` configured with:
   - `ProcessMode`: `AlwaysCreateNew`
   - `ShowExcelWindow`: `False`
   - `DisplayAlerts`: `False`
9. Opens the workbook through `ExcelApplicationCard`.
10. Reads `Roster`, `Intake`, `Tracker` headers, and `EmailLog`.
11. Clears the in-memory tracker table and resets all three counters.
12. Iterates through `dt_Roster`.
13. For each roster row, finds the first matching intake row with `DataTable.Select` using `Student ID` and `Reporting Month`. Single quotes in values are escaped with `Replace("'", "''")` for the filter expression.
14. For a match, converts `Date Received` to `DateTime`, formats it as `yyyy-MM-dd`, and classifies it as:
   - `Submitted` when the receipt date is on or before the deadline
   - `Submitted Late` otherwise
15. For no match, sets the date to an empty string and status to `Missing`.
16. Adds the roster row and calculated values to `dt_Tracker`.
17. Clears old tracker worksheet rows and writes the rebuilt tracker table.
18. Writes the three summary counts to `Tracker!I2:I4`.
19. Filters `dt_Tracker` to rows where `Status = Missing`, outputting `dt_MissingDTRs`.
20. Iterates through missing-DTR rows.
21. Validates the email by checking that it is nonblank and contains `@`.
22. Checks `dt_EmailLog` for an existing `Sent` record matching the student and reporting month.
23. Sends a Gmail reminder only when all of the following are true:
   - `bool_EnableReminders` is `True`
   - the email is considered valid
   - no matching sent reminder exists
   - `DateTime.Today` is later than the deadline date
24. Appends a `Sent` row to `dt_EmailLog` after the send activity completes and logs an informational message.
25. Logs a warning when a reminder is skipped.
26. Clears old EmailLog worksheet rows and writes the updated `dt_EmailLog` table.

## Root Variables

All variables are declared at the root `Main Sequence` so nested activities can access them:

- `str_ReportingMonth` — `String`
- `dt_ReportingMonthDate` — `DateTime`
- `dt_NextMonthFirstDate` — `DateTime`
- `int_DaysToFirstWednesday` — `Int32`, default `0`
- `dt_DeadlineDate` — `DateTime`
- `str_WorkbookPath` — `String`
- `dt_Roster` — `DataTable`
- `dt_Intake` — `DataTable`
- `dt_Tracker` — `DataTable`
- `dt_EmailLog` — `DataTable`
- `int_SubmittedCount` — `Int32`, default `0`
- `int_LateCount` — `Int32`, default `0`
- `int_MissingCount` — `Int32`, default `0`
- `dr_MatchedIntake` — `DataRow`
- `str_DateReceived` — `String`
- `str_Status` — `String`
- `dt_ReceiptDate` — `DateTime`
- `dt_MissingDTRs` — `DataTable`
- `bool_HasValidEmail` — `Boolean`, default `False`
- `bool_AlreadyReminded` — `Boolean`, default `False`
- `bool_EnableReminders` — `Boolean`, default `True`
- `bool_IsValidReportingMonth` — `Boolean`, default `False`
- `str_ReportingMonthPrompt` — `String`, default prompt text

## Gmail Configuration in the Current File

The current activity is:

- **Display name:** `Send Gmail Reminder (disabled by default)`
- **Class:** `SendEmailConnections`
- **Input type:** `TEXT`
- **Save as draft:** `False`
- **Use connection service:** `True`
- **Connection ID:** literal configured connection identifier in the XAML
- **To:** `New String() {CurrentRow("Email").ToString.Trim}`

The typed `New String()` array is required because `To` expects `IEnumerable(Of String)`. Do not change it back to an untyped collection literal.

The display name says “disabled by default,” but the current variable `bool_EnableReminders` defaults to `True`. Runtime eligibility still requires a valid email, no duplicate sent record, and a date after the deadline. Because `SaveAsDraft` is `False`, an eligible row can send a real email.

Before any runtime test:

1. Verify the configured Google connection.
2. Review the connection identifier and recipient data.
3. Set `bool_EnableReminders` to `False` or otherwise use a controlled test configuration when live sending is not intended.
4. Do not test against real student recipients until the message and recipients are approved.

No Gmail password, OAuth client secret, or access token belongs in `Main.xaml`.

## Important Current-State Notes

- The workflow uses the classic `DataTable.Select` matching expressions currently serialized in `Main.xaml`; do not replace them with a different LINQ implementation unless the XAML and validation are deliberately updated.
- `dt_Tracker` is initialized by reading the tracker header range before `ClearDataTable`; the six-column schema is needed by `AddDataRow` and `WriteRangeX`.
- `dt_EmailLog` is read before reminder processing so duplicate checks and appended rows use the same table.
- The workbook is modified inside the existing Excel scope. Do not move the EmailLog write outside the workbook body.
- The workflow was last structurally repaired by changing the Gmail `To` binding to an explicitly typed string array.

## Validation and Build

From the project directory, run per-file validation:

```powershell
uip rpa validate --file-path "C:\Users\justi\Documents\UiPath\AppDevt Finals\Main.xaml" --project-dir "C:\Users\justi\Documents\UiPath\AppDevt Finals" --output json
```

Then build the complete project:

```powershell
uip rpa build "C:\Users\justi\Documents\UiPath\AppDevt Finals" --output json
```

Do not run the workflow solely as a construction check: it reads and writes the workbook and may send Gmail reminders. Run it only after the workbook, reporting month, connection, reminder flag, and recipient safety checks are complete.

## Rebuild Guidance

If `Main.xaml` must be recreated in Studio:

1. Confirm the project is Windows/VB and restore the declared dependencies.
2. Recreate the root variables and preserve their scope.
3. Build the input/deadline sequence, including the retry loop and invalid-input message.
4. Build the Excel scope and keep all workbook reads/writes inside the workbook body.
5. Recreate the roster classification and counter logic.
6. Recreate the missing-DTR filter and reminder gate.
7. Configure the Gmail activity through Studio and verify its connection service settings.
8. Ensure the `To` value is an `IEnumerable(Of String)` expression, serialized in the current form as `New String() { ... }`.
9. Preserve the workbook headers and worksheet names.
10. Validate `Main.xaml`, then build the project before any runtime test.

Do not hand-edit Studio-generated view-state metadata or replace the connection identifier with a password or token.

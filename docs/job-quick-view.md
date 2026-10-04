# SQL Agent Quick Manage

In Object Explorer, right-click a job under **SQL Server Agent > Jobs** and select **Quick Manage**. The same window is available from **Axial SQL Tools toolbar > Tools > Quick Manage** when a job is selected.

Quick Manage opens in the SSMS document area alongside query tabs, titled **Job name - Quick Manage**. Reopening the same job on the same server and authentication identity activates its existing tab without replacing drafts. Different jobs can remain open simultaneously. Tabs are session-only and are not restored when SSMS restarts.

The tab shows the job name and server, current status, schedules, and the latest execution result. Details load asynchronously. The Steps tab keeps the step list beside the command editor. Job details combines description, category, creation/modification dates, job ID, and schedules. The Execution tab shows runtime details and messages. A compact header keeps job identity, status badges, and actions together, followed by a single-line last/next execution strip. Narrow or split document groups scroll the content to keep actions reachable.

## Common tasks

- **Start / Stop** requests a job start or stop through SQL Server Agent. Starting uses the job's configured starting step and its saved commands.
- **Enable / Disable** changes the job's enabled flag. Individual schedules and other job properties are not changed.
- **Refresh** reloads server details while retaining unsaved command drafts.
- Select a step to inspect its command, subsystem, database, and success/failure routing. A flag marks the configured starting step, which may be a step other than step 1. The marker follows server changes on refresh.
- Edit the command and select **Save step** to persist that step only. **Discard** restores the saved command.
- **Format SQL** applies the existing Axial SQL Tools formatter to a T-SQL command. Formatting is an editable, undoable change and is not saved automatically.

The editor includes syntax highlighting, line numbers, search, and undo/redo. Non-SQL commands are not passed to the T-SQL formatter. Commands that the formatter cannot parse, including some SQL Agent token expressions, remain unchanged if formatting fails.

Schedules appear in **Job details**, with SQL Server's human-readable descriptions from `msdb.dbo.sp_help_jobschedule @include_description = 1`, including frequency and active-date information.

## Status and edits

Job information loads when the window opens and refreshes after job actions. Use **Refresh** or **F5** to request another snapshot; there is no timed refresh. Refreshing preserves unsaved command drafts. Command edits, including formatting, show **Unsaved changes** in bold red and enable **Save step** and **Discard**. Closing with unsaved changes asks whether to discard them. Refresh failures are shown in the window instead of replacing the last successful snapshot with empty data.

Before a command is saved, Quick Manage checks that the same step still exists and that its command, subsystem, and database have not changed on the server. A conflict keeps the local draft and requires reconciling it with the current server version. A save changes only the command through `msdb.dbo.sp_update_jobstep`; it does not recreate the job or its steps.

Dates and times are shown in the SQL Server's local time. SQL Agent's cached next-run values can lag schedule edits, so the next-run time is an estimate. Job status is a snapshot, and a start or stop request can take time to appear in the next refresh.

Quick Manage uses the selected Object Explorer connection and does not elevate permissions. Viewing full commands and safely saving edits requires SELECT access to `msdb.dbo.sysjobsteps` and `msdb.dbo.sysjobs_view`; execution messages use `msdb.dbo.sysjobhistory`. Job actions also require the usual SQL Server Agent permissions. Sysadmin accounts satisfy these requirements. Permission failures and unavailable Agent status are reported in the window.

## Validation in SSMS 22

Use a disposable test job when checking actions and edits:

1. Open Quick Manage with no query editor open. Verify the settings icon and command name in both menus, action icons, colored status badges, SQL highlighting, and the Steps, Job details, and Execution tabs. Compare the readable schedule description with the native job properties dialog; verify category and created/modified dates. Confirm redundant header fields are absent from Job details.
2. Check a job with several steps, several schedules, no schedule, no history, and a long command. Configure a starting step other than step 1; verify its flag, change the starting step externally, and refresh. Ensure the flag follows the configuration while unsaved commands stay intact. Resize the window and check light, dark, and high-contrast themes.
3. Format a previously unmodified step. Confirm the bold red unsaved state and enabled Save/Discard buttons; undo to restore the clean state. Edit two steps, switch between them, and refresh. Confirm both drafts remain. Save one step and verify all other job settings are unchanged. Confirm closing warns about the remaining draft.
4. Change a step externally while a local draft is open. A save must report a conflict rather than overwrite the external change. Repeat with a deleted/recreated or renumbered step.
5. Leave the window idle and confirm it does not poll. Enable, disable, start, and stop the test job. Verify observed status and the latest completion result after refresh. Confirm a start uses saved commands.
6. Test an account without modification permissions, an unreachable server, a stopped SQL Server Agent, and closing during an asynchronous load. Confirm errors remain actionable and SSMS stays responsive.
7. Open the job context menu repeatedly, then open a non-job context menu. Confirm Quick Manage appears only once and only for a single job.
8. Open two different jobs and confirm separate document tabs. Open the first again and confirm its existing tab and drafts are retained. Switch to a SQL query and back; verify no reload or lost edits. Check Ctrl+F, Ctrl+S, and Ctrl+Shift+F in the hosted editor.
9. Close a dirty tab and choose No: it must remain open with its drafts. Repeat through Close All. Choose Yes and reopen: it must load fresh data. Attempt to close during a save or job action: closing must be blocked until the write finishes. Close during initial loading and verify safe cancellation. Resize/split the document group and check both scrollbars and editor scrolling. Confirm closed tabs are not restored after restarting SSMS.

Builds are started manually from GitHub Actions. This feature does not add automatic pull request or push triggers.

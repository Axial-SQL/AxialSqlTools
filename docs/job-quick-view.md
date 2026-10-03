# SQL Agent Job Quick View

In Object Explorer, right-click a job under **SQL Server Agent > Jobs** and select **Quick View**. The same window is available from **Axial SQL Tools toolbar > Tools > Job Quick View** when a job is selected.

Quick View opens a resizable window with the job name and server, current status, schedules, and the latest execution result. Details load asynchronously. The step list stays beside the command editor so you can move between steps without opening separate property dialogs.

## Common tasks

- **Start / Stop** requests a job start or stop through SQL Server Agent. Starting uses the job's configured starting step and its saved commands.
- **Enable / Disable** changes the job's enabled flag. Individual schedules and other job properties are not changed.
- **Refresh** reloads server details while retaining unsaved command drafts.
- Select a step to inspect its command, subsystem, database, and success/failure routing.
- Edit the command and select **Save step** to persist that step only. **Discard** restores the saved command.
- **Format SQL** applies the existing Axial SQL Tools formatter to a T-SQL command. Formatting is an editable, undoable change and is not saved automatically.

The editor includes syntax highlighting, line numbers, search, and undo/redo. Non-SQL commands are not passed to the T-SQL formatter. Commands that the formatter cannot parse, including some SQL Agent token expressions, remain unchanged if formatting fails.

## Status and edits

The window periodically refreshes job status without replacing unsaved command drafts. Closing with unsaved changes asks whether to discard them. Refresh failures are shown in the window instead of replacing the last successful snapshot with empty data.

Before a command is saved, Quick View checks that the same step still exists and that its command, subsystem, and database have not changed on the server. A conflict keeps the local draft and requires reconciling it with the current server version. A save changes only the command through `msdb.dbo.sp_update_jobstep`; it does not recreate the job or its steps.

Dates and times are shown in the SQL Server's local time. SQL Agent's cached next-run values can lag schedule edits, so the next-run time is an estimate. Job status is a snapshot, and a start or stop request can take time to appear in the next refresh.

Quick View uses the selected Object Explorer connection and does not elevate permissions. Viewing full commands and safely saving edits requires SELECT access to `msdb.dbo.sysjobsteps` and `msdb.dbo.sysjobs_view`; execution messages use `msdb.dbo.sysjobhistory`. Job actions also require the usual SQL Server Agent permissions. Sysadmin accounts satisfy these requirements. Permission failures and unavailable Agent status are reported in the window.

## Validation in SSMS 22

Use a disposable test job when checking actions and edits:

1. Open Quick View with no query editor open. Verify job selection, SQL highlighting, formatting, schedule descriptions, and the last-run message.
2. Check a job with several steps, several schedules, no schedule, no history, and a long command. Resize the window and check light, dark, and high-contrast themes.
3. Edit two steps, switch between them, and refresh. Confirm both drafts remain. Save one step and verify all other job settings are unchanged. Confirm closing warns about the remaining draft.
4. Change a step externally while a local draft is open. A save must report a conflict rather than overwrite the external change. Repeat with a deleted/recreated or renumbered step.
5. Enable, disable, start, and stop the test job. Verify observed status and the latest completion result after refresh. Confirm a start uses saved commands.
6. Test an account without modification permissions, an unreachable server, a stopped SQL Server Agent, and closing during an asynchronous load. Confirm errors remain actionable and SSMS stays responsive.
7. Open the job context menu repeatedly, then open a non-job context menu. Confirm Quick View appears only once and only for a single job.

Builds are started manually from GitHub Actions. This feature does not add automatic pull request or push triggers.

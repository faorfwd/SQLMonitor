---
name: collection-layer-snippets
description: Use when writing or modifying collection stored procedures (SCH-usp_collect_*.sql), SQL Agent job scripts (SCH-Job-*.sql), or PowerShell collectors in SQLMonitor/ — provides canonical patterns required by collection layer architecture decisions
---

# Collection Layer Snippets

## Overview

Canonical boilerplate for collection layer files. Decisions are locked in `docs/architecture/collection-layer-decisions.md` — these patterns enforce them. Use as a starting skeleton; fill in `<metric>`, `<table>`, and `<columns>`.

---

## Pattern A — Collection Stored Procedure

```sql
CREATE OR ALTER PROCEDURE [dbo].[usp_collect_<metric>]
    @data_destination_server SYSNAME = NULL
AS
BEGIN
    SET NOCOUNT ON;

    IF @data_destination_server IS NULL
        SELECT @data_destination_server = data_destination_sql_instance
        FROM dbo.instance_details;

    SELECT <columns>
    INTO #collected
    FROM <source>;

    IF @data_destination_server = @@SERVERNAME
        INSERT dbo.<table>
        SELECT * FROM #collected;
    ELSE
        INSERT [$(InventoryServer)].[DBA].[dbo].[<table>]
        SELECT * FROM #collected;
END
GO
```

**Rules:**
- `CREATE OR ALTER PROCEDURE` — never bare `CREATE PROCEDURE`
- `@data_destination_server SYSNAME = NULL` — always the first parameter
- Collect into `#tmp` locally first; then route — never INSERT directly from the DMV
- `[$(InventoryServer)]` is a SQLCMD variable passed via `Invoke-DbaQuery -Variable`
- No `OPENQUERY` for writes; no `EXEC ... AT`

---

## Pattern B — SQL Agent Job Upsert

```sql
/* ── Job ── */
IF NOT EXISTS (SELECT 1 FROM msdb.dbo.sysjobs WHERE name = N'(dba) Collect-X')
    EXEC msdb.dbo.sp_add_job
        @job_name      = N'(dba) Collect-X',
        @category_name = N'(dba) SQLMonitor',
        @enabled       = 1,
        @description   = N'<description>';
ELSE
    EXEC msdb.dbo.sp_update_job
        @job_name      = N'(dba) Collect-X',
        @enabled       = 1,
        @description   = N'<description>';

/* ── Step ── */
IF NOT EXISTS (
    SELECT 1
    FROM msdb.dbo.sysjobsteps js
    JOIN msdb.dbo.sysjobs     j  ON js.job_id   = j.job_id
    WHERE j.name = N'(dba) Collect-X' AND js.step_name = N'Collect'
)
    EXEC msdb.dbo.sp_add_jobstep
        @job_name          = N'(dba) Collect-X',
        @step_name         = N'Collect',
        @subsystem         = N'TSQL',   /* or N'CmdExec' for PowerShell */
        @command           = N'EXEC dbo.usp_collect_<metric>;',
        @on_success_action = 1,
        @on_fail_action    = 2;
ELSE
    EXEC msdb.dbo.sp_update_jobstep
        @job_name  = N'(dba) Collect-X',
        @step_name = N'Collect',
        @command   = N'EXEC dbo.usp_collect_<metric>;';

/* ── Schedule (add only; never update — schedules are environment-specific) ── */
IF NOT EXISTS (
    SELECT 1
    FROM msdb.dbo.sysjobschedules js
    JOIN msdb.dbo.sysjobs         j  ON js.job_id     = j.job_id
    JOIN msdb.dbo.sysschedules    s  ON js.schedule_id = s.schedule_id
    WHERE j.name = N'(dba) Collect-X'
)
    EXEC msdb.dbo.sp_add_jobschedule
        @job_name             = N'(dba) Collect-X',
        @name                 = N'Every 1 minute',
        @freq_type            = 4,  /* Daily */
        @freq_interval        = 1,
        @freq_subday_type     = 4,  /* Minutes */
        @freq_subday_interval = 1,
        @active_start_time    = 0;
```

**Rules:**
- No `APP_NAME() = 'SSMS'` gate — removed
- All three states handled: create / update / no-op (EXISTS check makes it a no-op)
- `@subsystem = N'TSQL'` for T-SQL steps; `N'CmdExec'` for PowerShell — never mixed within a step
- Schedule block: add-only; `sp_update_jobschedule` is never used (schedules vary per environment)

---

## Pattern C — PowerShell Collector Push

```powershell
# $inventoryServer comes from param or environment; never hardcode localhost
Write-DbaDbTableData -SqlInstance $inventoryServer `
                     -Database    'DBA' `
                     -Table       '<table>' `
                     -Schema      'dbo' `
                     -InputObject $collectedData `
                     -AutoCreateTable:$false
```

**Rules:**
- `-SqlInstance` = inventory server name, never `localhost` or `$env:COMPUTERNAME`
- `Write-DbaDbTableData` only — no `Invoke-Sqlcmd`, no `SqlConnection`/`SqlCommand`
- SQL auth: retrieve via `usp_get_credential` → construct `[pscredential]`; Windows auth is the default

---

## Quick Reference

| Need | Pattern | Key constraint |
|------|---------|----------------|
| New collect proc | A | `CREATE OR ALTER` + `@data_destination_server` |
| Modify existing proc | A | Migrate to `CREATE OR ALTER` when touched |
| New or updated job | B | Full upsert; no SSMS gate |
| PS collector | C | `Write-DbaDbTableData`; inventory as target |

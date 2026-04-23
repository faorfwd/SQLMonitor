# Tech — Optional Email Delivery (Global Kill Switch)

## Open Questions — Resolved

### 1. Caching the `email_delivery_enabled` flag
**Decision: No caching.** The lookup is a clustered PK point-read on a tiny temporal
table. SQL Server keeps that page in buffer pool indefinitely. Even at 60-second
intervals (`usp_collect_wait_stats`) the cost is one logical read from cache —
unmeasurable against the collection work that surrounds it. Session-scoped cache would
add complexity and a correctness hazard: a manual `UPDATE` to re-enable mail mid-session
would not take effect until the next SQL Agent job cycle. Inline SELECT every time.

### 2. PowerShell collectors with mail hooks
**Decision: No action required.** A survey of `SQLMonitor/*.ps1` found no
`sp_send_dbmail` calls or mail invocations. The Python alert engine in `Alerting/`
manages its own notification channels (Slack, PagerDuty, Email) independently and is
out of scope.

### 3. Combine the two `sma_params` reads in the two wrappers
**Decision: Implement.** Both `usp_wrapper_GetAllServerInfo` (lines 86–87) and
`usp_wrapper_GetAllServerCollectedData` (lines 91–92) currently fire two separate
point-reads for `dba_team_email_id` and `send_sqlmonitor_job_failure_mail`. Adding the
global switch naively would make it three. Folding all three into one pivoted SELECT
eliminates the N+1 pattern and makes the new preamble zero-marginal-cost in these files.

---

## Installation Modes

### Mode definitions

SQLMonitor supports two topologies. The same SQL and PowerShell code paths serve both;
the difference is which server the installer targets and where jobs execute.

**Distributed** — every baseline is also its own inventory. Each server hosts a
complete install (`SCH-Create-All-Objects.sql` + `SCH-Create-Inventory-Specific-Objects.sql`).
T-SQL and PowerShell jobs run locally. The installer is invoked once per server with
`-SqlInstanceToBaseline <server> -InventoryServer <same-server>`.

**Centralized** — one dedicated inventory server; one or more remote baselines. Only
the inventory receives `SCH-Create-Inventory-Specific-Objects.sql`. Remote baselines
receive `SCH-Create-All-Objects.sql` only. All SQL Agent jobs execute at the inventory
(`-SqlInstanceForTsqlJobs <inventory> -SqlInstanceForPowershellJobs <inventory>`). The
installer is invoked once with `$SqlInstanceToBaseline = $InventoryServer` to install
the inventory, then once per remote baseline with `$SqlInstanceToBaseline = <baseline>`
and both job-server params pointing back to the inventory.

### Where `dbo.sma_params` and the flag live

`SCH-Create-Inventory-Specific-Objects.sql` (which seeds `email_delivery_enabled`) runs
only when `$InventoryServer -eq $SqlInstanceToBaseline`. This means:

| Mode | Where the flag row lives |
|------|--------------------------|
| Distributed | Local `DBA.dbo.sma_params` on each server (each is its own inventory) |
| Centralized | `DBA.dbo.sma_params` on the inventory server only |

### Where collection procs execute and read the flag

Every email-sending proc reads `dbo.sma_params` with a two-part name (no server
qualifier). The name resolves to the database of the connection that is executing the
proc — always the server where T-SQL jobs run (`$SqlInstanceForTsqlJobs`).

| Mode | Proc executes at | `dbo.sma_params` resolved | Flag row present? |
|------|-----------------|--------------------------|-------------------|
| Distributed | Local server (= inventory) | Local `DBA` | Yes ✓ |
| Centralized — inventory install | Inventory | Inventory `DBA` | Yes ✓ |
| Centralized — remote baseline install | Inventory (jobs run there) | Inventory `DBA` | Yes ✓ |

No preamble SQL changes are needed to support either mode.

### `SkipMailProfileCheck` (E2) across modes

The mail-profile validation at line 2260 runs against `$conSqlInstanceToBaseline`.

- **Distributed**: baseline = inventory → the check targets the mail-enabled server. When
  `-EnableEmailAlerts:$false`, E2 sets `$SkipMailProfileCheck = $true` and the check is
  bypassed. Correct.
- **Centralized, inventory install**: `$conSqlInstanceToBaseline` = inventory → same
  server that would send mail. E2 bypass is still correct.
- **Centralized, remote baseline install**: `$conSqlInstanceToBaseline` = remote
  baseline, which typically has no mail profile. The check would always fail or be
  vacuous anyway; when `-EnableEmailAlerts:$false`, E2 skips it. Correct (the mail
  profile that matters lives on the inventory, which was validated during the inventory
  install).

---

## Data Layer

### `dbo.sma_params` seed row

One new row in the idempotent VALUES block inside
`DDLs/SCH-Create-Inventory-Specific-Objects.sql` (after `send_sqlmonitor_job_failure_mail`,
around line 1550):

```sql
('email_delivery_enabled','1',
 'Global kill switch: 0 disables all sp_send_dbmail calls across SQLMonitor'),
```

The existing pattern:
```sql
insert dbo.sma_params (param_key, param_value, remarks)
select ...
from (values (...)) my_keys (param_key, param_value, remarks)
left join dbo.sma_params p on p.param_key = my_keys.param_key
where p.param_key is null;
```
ensures this is a no-op on upgrade runs — the stored value is never clobbered by DDL
re-runs.

---

## Runtime Gate (SQL)

### Standard preamble

Inserted after the initial `DECLARE` block in each proc, before the first business-logic
statement. Uses `ISNULL(..., 1)` so a missing row defaults to enabled (preserves current
behavior on environments that haven't run the new DDL yet):

```sql
DECLARE @email_delivery_enabled BIT = ISNULL(
    (SELECT TOP 1 CONVERT(BIT, param_value)
     FROM dbo.sma_params
     WHERE param_key = 'email_delivery_enabled'),
    1  -- absent row → treat as enabled
);
```

The coercion line that follows uses whichever local flag the proc owns:

| Local flag | Coercion |
|------------|----------|
| `@send_mail` | `IF @email_delivery_enabled = 0 SET @send_mail = 0;` |
| `@send_error_mail` | `IF @email_delivery_enabled = 0 SET @send_error_mail = 0;` |
| `@send_email` | `IF @email_delivery_enabled = 0 SET @send_email = 0;` |

No `sp_send_dbmail` calls are modified in place; every existing `IF @send_mail = 1`
(or equivalent) gate naturally short-circuits.

### Combined pivot (the two wrappers only)

For `usp_wrapper_GetAllServerInfo` and `usp_wrapper_GetAllServerCollectedData`, replace
the two existing separate `SELECT` statements with a single pivoted read:

```sql
-- declaration added to the DECLARE block above
DECLARE @email_delivery_enabled BIT;

-- replaces the two separate point-reads
SELECT
    @recipients              = MAX(CASE WHEN param_key = 'dba_team_email_id'
                                   THEN param_value END),
    @send_error_mail         = MAX(CASE WHEN param_key = 'send_sqlmonitor_job_failure_mail'
                                   THEN CONVERT(BIT, param_value) END),
    @email_delivery_enabled  = ISNULL(MAX(CASE WHEN param_key = 'email_delivery_enabled'
                                   THEN CONVERT(BIT, param_value) END), 1)
FROM dbo.sma_params
WHERE param_key IN (
    'dba_team_email_id',
    'send_sqlmonitor_job_failure_mail',
    'email_delivery_enabled'
);
IF @email_delivery_enabled = 0 SET @send_error_mail = 0;
```

This replaces lines 86–87 in `GetAllServerInfo` and lines 91–92 in
`GetAllServerCollectedData`. The `send_sqlmonitor_job_failure_mail` semantics are
unchanged — the pivot just returns both values in one pass.

---

## Proc Inventory

### CREATE OR ALTER conversion

All modified procs must use `CREATE OR ALTER PROCEDURE` (Collection Layer Decision 2).
The current stub pattern:
```sql
if object_id('dbo.usp_xxx') is null exec ('create proc dbo.usp_xxx as select 1');
go
alter procedure dbo.usp_xxx
```
becomes:
```sql
create or alter procedure dbo.usp_xxx
```
(two lines removed, one changed).

### Per-file action table

| File | Current form | Action | Flag coerced |
|------|-------------|--------|--------------|
| `SCH-usp_send_login_expiry_emails.sql` | `CREATE OR ALTER` | Preamble only | `@send_mail` |
| `SCH-usp_GetAllServerDashboardMail.sql` | stub+`ALTER` | Convert + preamble | `@send_mail` |
| `SCH-usp_wrapper_populate_sma_sql_instance.sql` | `CREATE OR ALTER` | Preamble only | `@send_mail` |
| `SCH-usp_capture_alert_messages.sql` | stub+`ALTER` | Convert + preamble | `@send_error_mail` |
| `SCH-usp_check_sql_agent_jobs.sql` | stub+`ALTER` | Convert + preamble | `@send_error_mail` |
| `SCH-usp_collect_wait_stats.sql` | stub+`ALTER` | Convert + preamble | `@send_error_mail` |
| `SCH-usp_collect_file_io_stats.sql` | stub+`ALTER` | Convert + preamble | `@send_error_mail` |
| `SCH-usp_collect_memory_clerks.sql` | stub+`ALTER` | Convert + preamble | `@send_error_mail` |
| `SCH-usp_collect_ag_health_state.sql` | stub+`ALTER` | Convert + preamble | `@send_error_mail` |
| `SCH-usp_run_WhoIsActive.sql` | stub+`ALTER` | Convert + preamble | `@send_error_mail` |
| `SCH-usp_wrapper_CollectPrivilegedInfo.sql` | stub+`ALTER` | Convert + preamble | `@send_error_mail` |
| `SCH-usp_wrapper_GetAllServerInfo.sql` | stub+`ALTER` | Convert + combined pivot | `@send_error_mail` |
| `SCH-usp_wrapper_GetAllServerCollectedData.sql` | stub+`ALTER` | Convert + combined pivot | `@send_error_mail` |
| `SCH-usp_LogSaver.sql` | stub+`ALTER` | Convert + preamble | `@send_email` |
| `SCH-usp_TempDbSaver.sql` | stub+`ALTER` | Convert + preamble | `@send_email` |

> **Note:** The spec lists "17 total" but enumerates 15 files. These 15 are the
> authoritative set; implementation covers exactly them.

---

## Installer Changes (`Install-SQLMonitor.ps1`)

### E1 — New parameter (alongside existing `$SkipMailProfileCheck`, line ~189)

```powershell
[Parameter(Mandatory=$false)]
[bool]$EnableEmailAlerts = $true,
```

Defaults to `$true` — existing `Wrapper-InstallSQLMonitor.ps1` scripts keep working
with no change.

### E2 — Auto-skip mail profile check (before line ~2259)

```powershell
if (-not $EnableEmailAlerts) {
    $SkipMailProfileCheck = $true
}
```

Must execute before the `if(-not $SkipMailProfileCheck)` block that validates
`msdb.dbo.sysmail_profile`.

### E3 — Write to `sma_params` (after the DDL step that runs SCH-Create-Inventory-Specific-Objects.sql)

Uses `$conInventoryServer` — the same connection used to execute
`SCH-Create-Inventory-Specific-Objects.sql` (established at line ~2618 in the current
installer). This is correct for both installation modes:

- **Distributed** (`$InventoryServer = $SqlInstanceToBaseline`): `$conInventoryServer`
  resolves to the local server, which is both the baseline and the inventory.
- **Centralized** (`$InventoryServer ≠ $SqlInstanceToBaseline`): `$conInventoryServer`
  resolves to the central inventory, where collection procs execute and read
  `dbo.sma_params`. When installing a remote baseline in centralized mode, E3 still
  writes to the inventory — not the baseline — so the flag is always on the server
  that enforces it.

Only fires when the caller explicitly passed the flag (detected via
`$PSBoundParameters.ContainsKey`):

```powershell
if ($PSBoundParameters.ContainsKey('EnableEmailAlerts')) {
    $emailValue  = if ($EnableEmailAlerts) { '1' } else { '0' }
    $emailReason = "Set by Install-SQLMonitor.ps1 on $(Get-Date -Format 'u')"
    $sqlUpsert   = @"
UPDATE dbo.sma_params
SET    param_value = '$emailValue',
       remarks     = '$emailReason'
WHERE  param_key   = 'email_delivery_enabled';
"@
    $conInventoryServer |
        Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlUpsert `
                        -EnableException -Verbose:$false -Debug:$false
}
```

Fresh installs: the DDL seeds `'1'`; the UPDATE then flips to `'0'` if needed.
Upgrade runs without the flag: DDL is a no-op (row exists); UPDATE is skipped → value
preserved. Upgrade runs with the flag: UPDATE overwrites to the new value.

---

## Wrapper Sample

In `Wrapper-Samples/Wrapper-InstallSQLMonitor.ps1`, add a commented line in the
`$params` hashtable alongside `#SkipMailProfileCheck`:

```powershell
#EnableEmailAlerts = $false   # uncomment to disable all sp_send_dbmail calls at runtime
```

---

## Risks

| Risk | Mitigation |
|------|-----------|
| Missing `email_delivery_enabled` row (pre-DDL upgrade) | `ISNULL(..., 1)` in preamble — absent row = enabled, no behavior change |
| Developer adds a new mail-sending proc and omits the preamble | Verification step 3 (`grep email_delivery_enabled`) catches it during code review |
| Combined pivot returns wrong `@send_error_mail` if row has no `send_sqlmonitor_job_failure_mail` | `MAX(CASE ...)` returns NULL → downstream `IF @send_error_mail = 1` treats NULL as false — safe; same as today |
| Centralized: E3 uses wrong connection (`$conSqlInstanceToBaseline` instead of `$conInventoryServer`) | E3 explicitly uses `$conInventoryServer -Database $InventoryDatabase`; writes always land on inventory |
| Centralized: remote baseline install passes `-EnableEmailAlerts:$false` and writes to inventory; later inventory re-install overwrites | Idempotent UPDATE; last explicit pass of the flag wins. Document in wrapper sample that the flag is inventory-scoped. |

---

## Verification Checklist

1. `grep -rn "CREATE PROCEDURE " DDLs/SCH-usp_*.sql` — expect **zero** hits.
2. `grep -rn "CREATE OR ALTER PROCEDURE" DDLs/SCH-usp_*.sql` — expect **15** files.
3. `grep -rn "email_delivery_enabled" DDLs/` — expect **1** seed row + **15** preambles/pivots.
4. `grep -n "EnableEmailAlerts" SQLMonitor/Install-SQLMonitor.ps1` — expect **3** hits
   (declaration, auto-skip block, sma_params UPDATE block).
5. `grep -n "conInventoryServer" SQLMonitor/Install-SQLMonitor.ps1` — E3 block must
   reference `$conInventoryServer`, not `$conSqlInstanceToBaseline`.
6. Manual smoke test:
   ```sql
   UPDATE dbo.sma_params SET param_value = '0' WHERE param_key = 'email_delivery_enabled';
   DECLARE @before DATETIME = GETDATE();
   EXEC dbo.usp_GetAllServerDashboardMail @send_mail = 1, @recipients = 'test@example.com';
   SELECT COUNT(*) FROM msdb.dbo.sysmail_allitems WHERE sent_date > @before;
   -- expect: 0
   ```

---

## Next in Flow

Continue with [`/kyos:tasks`](./tasks.md) to break the implementation into ordered
execution slices.

# Tasks — Optional Email Delivery (Global Kill Switch)

Execution slices derived from [`tech.md`](./tech.md). All slices target the
`email-disabling` branch. T3, T4, and T5 can run in parallel once T1 is done.

---

## T1 — DDL: seed `email_delivery_enabled` row

**Why:** The kill switch row in `dbo.sma_params` is the source of truth for every
runtime gate. Nothing else can land until this row exists.

**Depends on:** nothing

**File:** `DDLs/SCH-Create-Inventory-Specific-Objects.sql`

**What to do:**  
Add one row to the idempotent `INSERT … SELECT … LEFT JOIN … WHERE p.param_key IS NULL`
block (after the `send_sqlmonitor_job_failure_mail` row, ~line 1550):

```sql
('email_delivery_enabled','1',
 'Global kill switch: 0 disables all sp_send_dbmail calls across SQLMonitor'),
```

**Done when:**
- The row is present in the VALUES block.
- `grep -rn "email_delivery_enabled" DDLs/SCH-Create-Inventory-Specific-Objects.sql` → 1 hit.
- Re-running the DDL on a server that already has the row produces no error and no
  data change (idempotency).

---

## T2 — Installer: `$EnableEmailAlerts` parameter + E2 + E3 + wrapper sample comment

**Why:** Exposes the kill switch to the install workflow. E3 is the only place outside
of a manual UPDATE where the flag value is written; it must use `$conInventoryServer`
so it writes to the right server in both distributed and centralized topologies.

**Depends on:** T1 (E3 UPDATEs the row seeded by T1; DDL must exist first)

**Files:**
- `SQLMonitor/Install-SQLMonitor.ps1`
- `Wrapper-Samples/Wrapper-InstallSQLMonitor.ps1`

**What to do:**

E1 — add parameter alongside `$SkipMailProfileCheck` (~line 189):
```powershell
[Parameter(Mandatory=$false)]
[bool]$EnableEmailAlerts = $true,
```

E2 — add auto-skip block immediately before the `if(-not $SkipMailProfileCheck)` mail
profile validation (~line 2259):
```powershell
if (-not $EnableEmailAlerts) {
    $SkipMailProfileCheck = $true
}
```

E3 — add upsert block after the step that runs
`SCH-Create-Inventory-Specific-Objects.sql` (~line 2618), using `$conInventoryServer`:
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

Wrapper sample — add a commented line in the `$params` hashtable alongside
`#SkipMailProfileCheck`:
```powershell
#EnableEmailAlerts = $false   # uncomment to disable all sp_send_dbmail calls at runtime
```

**Done when:**
- `grep -n "EnableEmailAlerts" SQLMonitor/Install-SQLMonitor.ps1` → exactly 3 hits
  (declaration, E2 block, E3 block).
- `grep -n "conInventoryServer" SQLMonitor/Install-SQLMonitor.ps1` shows E3 uses
  `$conInventoryServer`, not `$conSqlInstanceToBaseline`.
- Commented line present in `Wrapper-Samples/Wrapper-InstallSQLMonitor.ps1`.

---

## T3 — Preamble: procs already in `CREATE OR ALTER` form (2 files)

**Why:** These two procs need only the standard preamble block — no stub conversion
required. They are the simplest change and good to do first to validate the preamble
pattern before touching the larger group.

**Depends on:** T1

**Files:**

| File | Local flag |
|------|-----------|
| `DDLs/SCH-usp_send_login_expiry_emails.sql` | `@send_mail` |
| `DDLs/SCH-usp_wrapper_populate_sma_sql_instance.sql` | `@send_mail` |

**What to do:**  
In each file, insert the standard preamble after the initial `DECLARE` block, before the
first business-logic statement:

```sql
DECLARE @email_delivery_enabled BIT = ISNULL(
    (SELECT TOP 1 CONVERT(BIT, param_value)
     FROM dbo.sma_params
     WHERE param_key = 'email_delivery_enabled'),
    1
);
IF @email_delivery_enabled = 0 SET @send_mail = 0;
```

**Done when:**
- Both files contain `email_delivery_enabled`.
- `grep -n "CREATE PROCEDURE " DDLs/SCH-usp_send_login_expiry_emails.sql DDLs/SCH-usp_wrapper_populate_sma_sql_instance.sql` → 0 hits (confirm no regression).

---

## T4 — Convert + preamble: stub+ALTER procs (11 files)

**Why:** Eleven procs still use the old `IF object_id … EXEC ('create proc …'); GO ALTER`
stub pattern. Converting them to `CREATE OR ALTER` is a required prerequisite (Collection
Layer Decision 2) and a two-line mechanical change per file. The preamble drops in
immediately after.

**Depends on:** T1 (can run in parallel with T3 and T5)

**Files and local flags:**

| File | Local flag |
|------|-----------|
| `DDLs/SCH-usp_GetAllServerDashboardMail.sql` | `@send_mail` |
| `DDLs/SCH-usp_capture_alert_messages.sql` | `@send_error_mail` |
| `DDLs/SCH-usp_check_sql_agent_jobs.sql` | `@send_error_mail` |
| `DDLs/SCH-usp_collect_wait_stats.sql` | `@send_error_mail` |
| `DDLs/SCH-usp_collect_file_io_stats.sql` | `@send_error_mail` |
| `DDLs/SCH-usp_collect_memory_clerks.sql` | `@send_error_mail` |
| `DDLs/SCH-usp_collect_ag_health_state.sql` | `@send_error_mail` |
| `DDLs/SCH-usp_run_WhoIsActive.sql` | `@send_error_mail` |
| `DDLs/SCH-usp_wrapper_CollectPrivilegedInfo.sql` | `@send_error_mail` |
| `DDLs/SCH-usp_LogSaver.sql` | `@send_email` |
| `DDLs/SCH-usp_TempDbSaver.sql` | `@send_email` |

**What to do per file:**

1. Replace the stub pattern:
   ```sql
   -- remove these two lines:
   if object_id('dbo.usp_xxx') is null exec ('create proc dbo.usp_xxx as select 1');
   go
   -- change this line:
   alter procedure dbo.usp_xxx
   -- to:
   create or alter procedure dbo.usp_xxx
   ```

2. Insert the standard preamble (after the initial DECLARE block) with the correct
   coercion flag for the file:
   ```sql
   DECLARE @email_delivery_enabled BIT = ISNULL(
       (SELECT TOP 1 CONVERT(BIT, param_value)
        FROM dbo.sma_params
        WHERE param_key = 'email_delivery_enabled'),
       1
   );
   IF @email_delivery_enabled = 0 SET @send_error_mail = 0;  -- or @send_mail / @send_email
   ```

**Done when:**
- `grep -rn "CREATE PROCEDURE " DDLs/SCH-usp_*.sql` → 0 hits.
- All 11 files contain `email_delivery_enabled`.

---

## T5 — Combined pivot: two wrapper procs

**Why:** `usp_wrapper_GetAllServerInfo` and `usp_wrapper_GetAllServerCollectedData` each
already do two separate `sma_params` point-reads for `dba_team_email_id` and
`send_sqlmonitor_job_failure_mail`. Adding a third read naively would create N+1 reads.
The tech plan folds all three into one pivoted SELECT — the preamble here looks different
from the standard pattern.

**Depends on:** T1 (can run in parallel with T3 and T4)

**Files:**
- `DDLs/SCH-usp_wrapper_GetAllServerInfo.sql`
- `DDLs/SCH-usp_wrapper_GetAllServerCollectedData.sql`

**What to do per file:**

1. Convert stub+ALTER to `CREATE OR ALTER PROCEDURE` (same as T4).

2. Add `DECLARE @email_delivery_enabled BIT;` to the existing DECLARE block (do **not**
   use the inline assignment form — it comes from the pivot below).

3. Replace the two existing separate `SELECT` reads (lines ~86–87 in GetAllServerInfo,
   ~91–92 in GetAllServerCollectedData) with the single pivoted read:

```sql
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

**Done when:**
- Both files contain the pivoted SELECT with all three `param_key` values.
- No standalone `SELECT @recipients = …` or `SELECT @send_error_mail = …` reads remain.
- Both files use `CREATE OR ALTER PROCEDURE`.

---

## T6 — Verification sweep

**Why:** Confirms all slices landed correctly and the kill switch works end-to-end.
Run after T1–T5 are all merged to the branch.

**Depends on:** T1, T2, T3, T4, T5

**Checks (from tech.md verification checklist):**

```bash
# 1. No bare CREATE PROCEDURE in modified procs
grep -rn "CREATE PROCEDURE " DDLs/SCH-usp_*.sql
# expect: 0 hits

# 2. All 15 proc files use CREATE OR ALTER
grep -rn "CREATE OR ALTER PROCEDURE" DDLs/SCH-usp_*.sql
# expect: 15 files

# 3. Seed row + 15 preambles/pivots
grep -rn "email_delivery_enabled" DDLs/
# expect: 1 seed row (in SCH-Create-Inventory-Specific-Objects.sql) + 15 proc occurrences = 16 total

# 4. Installer has exactly 3 EnableEmailAlerts hits
grep -n "EnableEmailAlerts" SQLMonitor/Install-SQLMonitor.ps1
# expect: 3 hits

# 5. E3 uses $conInventoryServer
grep -n "conInventoryServer" SQLMonitor/Install-SQLMonitor.ps1
# expect: E3 block references $conInventoryServer
```

**Manual smoke test** (run on a dev instance after deploying the DDLs):

```sql
UPDATE dbo.sma_params SET param_value = '0' WHERE param_key = 'email_delivery_enabled';
DECLARE @before DATETIME = GETDATE();
EXEC dbo.usp_GetAllServerDashboardMail @send_mail = 1, @recipients = 'test@example.com';
SELECT COUNT(*) FROM msdb.dbo.sysmail_allitems WHERE sent_date > @before;
-- expect: 0

UPDATE dbo.sma_params SET param_value = '1' WHERE param_key = 'email_delivery_enabled';
-- restore after test
```

**Done when:** All 5 grep checks pass. Manual smoke test confirms 0 mails sent with
flag off and normal send with flag on.

---

## Dependency map

```
T1 (DDL seed)
├── T2 (Installer)         ← run after T1; touches different files
├── T3 (Preamble only)     ← run after T1; can run in parallel with T4 and T5
├── T4 (Convert+preamble)  ← run after T1; can run in parallel with T3 and T5
└── T5 (Combined pivot)    ← run after T1; can run in parallel with T3 and T4
         └── T6 (Verify)   ← run after all of the above
```

---

## Next in flow

Continue with [`/kyos:implement`](./implement.md) — execute slices one by one, starting
with T1.

# Tasks — Optional Email Delivery (Global Kill Switch)

> Execution slices in dependency order. Slices 3a/3b/3c are independent and can run in
> parallel. No automated test harness exists for T-SQL stored procedures in this repo —
> verification relies on grep checks and manual SQL smoke tests.

---

## Slice 1 — DDL seed row

**Why:** Every proc preamble reads `dbo.sma_params WHERE param_key = 'email_delivery_enabled'`. The row must exist before the runtime gate is meaningful. Must land first.

**Depends on:** nothing

**Files:**
- `DDLs/SCH-Create-Inventory-Specific-Objects.sql`

**Change:** Add one row to the idempotent VALUES block (after `send_sqlmonitor_job_failure_mail`, ~line 1550):
```sql
('email_delivery_enabled','1',
 'Global kill switch: 0 disables all sp_send_dbmail calls across SQLMonitor'),
```

**Done when:** Row is in the VALUES list; the surrounding LEFT JOIN `WHERE p.param_key IS NULL` pattern is preserved intact.

**Verify:**
```bash
grep -n "email_delivery_enabled" DDLs/SCH-Create-Inventory-Specific-Objects.sql
# expect: 1 hit in the VALUES block
```

---

## Slice 2 — Wrapper procs: combined pivot

**Why:** `usp_wrapper_GetAllServerInfo` and `usp_wrapper_GetAllServerCollectedData` already read two `sma_params` keys separately. Adding the global switch naively would make it three reads. The combined pivot folds all three into one SELECT — establishing the canonical pattern for these two high-value procs.

**Depends on:** Slice 1

**Files:**
- `DDLs/SCH-usp_wrapper_GetAllServerInfo.sql`
- `DDLs/SCH-usp_wrapper_GetAllServerCollectedData.sql`

**Changes per file:**
1. Remove the two-line `CREATE PROC` stub (lines 13–14 in each file).
2. Change `ALTER PROCEDURE` → `CREATE OR ALTER PROCEDURE`.
3. Add `DECLARE @email_delivery_enabled BIT;` to the DECLARE block.
4. Replace the two existing single-key SELECTs (GetAllServerInfo lines 86–87; GetAllServerCollectedData lines 91–92) with the combined pivot from tech.md §"Combined pivot".
5. Add `IF @email_delivery_enabled = 0 SET @send_error_mail = 0;` immediately after the pivot.

**Done when:** Both files use `CREATE OR ALTER PROCEDURE`; each contains exactly one `FROM dbo.sma_params` block that returns all three keys; the coercion line follows it.

**Verify:**
```bash
grep -n "CREATE OR ALTER" DDLs/SCH-usp_wrapper_GetAllServerInfo.sql DDLs/SCH-usp_wrapper_GetAllServerCollectedData.sql
grep -n "email_delivery_enabled" DDLs/SCH-usp_wrapper_GetAllServerInfo.sql DDLs/SCH-usp_wrapper_GetAllServerCollectedData.sql
grep -n "CREATE PROC\b" DDLs/SCH-usp_wrapper_GetAllServerInfo.sql DDLs/SCH-usp_wrapper_GetAllServerCollectedData.sql
# last grep: expect zero hits
```

---

## Slice 3a — `@send_error_mail` procs (8 files, standard preamble)

**Why:** Largest group. All use `@send_error_mail` and currently use the stub+ALTER pattern.

**Depends on:** Slice 1 (conceptually; files are independent of Slice 2)

**Files:**
- `DDLs/SCH-usp_capture_alert_messages.sql`
- `DDLs/SCH-usp_check_sql_agent_jobs.sql`
- `DDLs/SCH-usp_collect_wait_stats.sql`
- `DDLs/SCH-usp_collect_file_io_stats.sql`
- `DDLs/SCH-usp_collect_memory_clerks.sql`
- `DDLs/SCH-usp_collect_ag_health_state.sql`
- `DDLs/SCH-usp_run_WhoIsActive.sql`
- `DDLs/SCH-usp_wrapper_CollectPrivilegedInfo.sql`

**Change per file:**
1. Remove the two-line `CREATE PROC` stub.
2. Change `ALTER PROCEDURE` → `CREATE OR ALTER PROCEDURE`.
3. Insert the standard preamble (from tech.md §"Standard preamble") after the initial DECLARE block, before the first business-logic statement:
   ```sql
   DECLARE @email_delivery_enabled BIT = ISNULL(
       (SELECT TOP 1 CONVERT(BIT, param_value)
        FROM dbo.sma_params
        WHERE param_key = 'email_delivery_enabled'),
       1
   );
   IF @email_delivery_enabled = 0 SET @send_error_mail = 0;
   ```

**Done when:** All 8 files use `CREATE OR ALTER PROCEDURE` and contain the preamble.

**Verify:**
```bash
grep -l "CREATE OR ALTER" DDLs/SCH-usp_capture_alert_messages.sql DDLs/SCH-usp_check_sql_agent_jobs.sql DDLs/SCH-usp_collect_wait_stats.sql DDLs/SCH-usp_collect_file_io_stats.sql DDLs/SCH-usp_collect_memory_clerks.sql DDLs/SCH-usp_collect_ag_health_state.sql DDLs/SCH-usp_run_WhoIsActive.sql DDLs/SCH-usp_wrapper_CollectPrivilegedInfo.sql
# expect: all 8 filenames echoed

grep -l "email_delivery_enabled" DDLs/SCH-usp_capture_alert_messages.sql DDLs/SCH-usp_check_sql_agent_jobs.sql DDLs/SCH-usp_collect_wait_stats.sql DDLs/SCH-usp_collect_file_io_stats.sql DDLs/SCH-usp_collect_memory_clerks.sql DDLs/SCH-usp_collect_ag_health_state.sql DDLs/SCH-usp_run_WhoIsActive.sql DDLs/SCH-usp_wrapper_CollectPrivilegedInfo.sql
# expect: all 8 filenames echoed
```

---

## Slice 3b — `@send_mail` procs (3 files, standard preamble)

**Why:** Two already use `CREATE OR ALTER`; only `usp_GetAllServerDashboardMail` needs the stub conversion.

**Depends on:** Slice 1

**Files:**
- `DDLs/SCH-usp_send_login_expiry_emails.sql` — preamble only
- `DDLs/SCH-usp_GetAllServerDashboardMail.sql` — convert + preamble
- `DDLs/SCH-usp_wrapper_populate_sma_sql_instance.sql` — preamble only

**Change per file:**
- For `usp_GetAllServerDashboardMail`: remove stub, change to `CREATE OR ALTER`.
- All three: insert preamble with `SET @send_mail = 0` coercion after the initial DECLARE block.

**Done when:** All 3 files contain `CREATE OR ALTER PROCEDURE` and the `@send_mail` preamble.

**Verify:**
```bash
grep -l "email_delivery_enabled" DDLs/SCH-usp_send_login_expiry_emails.sql DDLs/SCH-usp_GetAllServerDashboardMail.sql DDLs/SCH-usp_wrapper_populate_sma_sql_instance.sql
grep -n "CREATE PROC\b" DDLs/SCH-usp_GetAllServerDashboardMail.sql
# last grep: expect zero hits
```

---

## Slice 3c — `@send_email` procs (2 files, standard preamble)

**Why:** `usp_LogSaver` and `usp_TempDbSaver` use `@send_email` (not `@send_mail`), so they get a distinct coercion line.

**Depends on:** Slice 1

**Files:**
- `DDLs/SCH-usp_LogSaver.sql`
- `DDLs/SCH-usp_TempDbSaver.sql`

**Change per file:**
1. Remove stub; change to `CREATE OR ALTER PROCEDURE`.
2. Insert preamble with `SET @send_email = 0` coercion.

**Done when:** Both files use `CREATE OR ALTER PROCEDURE` and contain the `@send_email` preamble.

**Verify:**
```bash
grep -l "email_delivery_enabled" DDLs/SCH-usp_LogSaver.sql DDLs/SCH-usp_TempDbSaver.sql
# expect: both filenames
```

---

## Slice 4 — Installer changes

**Why:** Exposes `EnableEmailAlerts` to operators; auto-skips the mail profile prerequisite when disabled; writes the flag to `sma_params` at install time with correct upgrade semantics.

**Depends on:** Slice 1 (sma_params row must exist for the UPDATE to target)

**Files:**
- `SQLMonitor/Install-SQLMonitor.ps1`

**Three changes (E1, E2, E3 from tech.md):**

E1 — Add parameter alongside `$SkipMailProfileCheck` (~line 189):
```powershell
[Parameter(Mandatory=$false)]
[bool]$EnableEmailAlerts = $true,
```

E2 — Auto-skip block, placed before the `if(-not $SkipMailProfileCheck)` validation at ~line 2259:
```powershell
if (-not $EnableEmailAlerts) {
    $SkipMailProfileCheck = $true
}
```

E3 — sma_params UPDATE block, placed directly after the step that executes
`SCH-Create-Inventory-Specific-Objects.sql`:
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
    $conSqlInstanceToBaseline |
        Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlUpsert `
                        -EnableException -Verbose:$false -Debug:$false
}
```

**Done when:** `grep -n "EnableEmailAlerts" SQLMonitor/Install-SQLMonitor.ps1` returns exactly 3 hits (declaration, E2 block, E3 block) plus the help comment at the .PARAMETER section.

**Verify:**
```bash
grep -n "EnableEmailAlerts" SQLMonitor/Install-SQLMonitor.ps1
# expect: 3-4 hits (param declaration, help text, E2 block, E3 block)
grep -n "SkipMailProfileCheck" SQLMonitor/Install-SQLMonitor.ps1
# confirm E2 block appears before the SkipMailProfileCheck validation block
```

---

## Slice 5 — Wrapper sample

**Why:** Documents the new flag to operators copying the wrapper template.

**Depends on:** nothing (docs only)

**Files:**
- `Wrapper-Samples/Wrapper-InstallSQLMonitor.ps1`

**Change:** Add one commented line in the `$params` hashtable alongside `#SkipMailProfileCheck`:
```powershell
#EnableEmailAlerts = $false   # uncomment to disable all sp_send_dbmail calls at runtime
```

**Done when:** Line is present in the wrapper sample.

**Verify:**
```bash
grep "EnableEmailAlerts" Wrapper-Samples/Wrapper-InstallSQLMonitor.ps1
# expect: 1 hit
```

---

## Slice 6 — End-to-end verification

**Why:** Structural checks pass, but only a live SQL run can confirm the preamble actually suppresses mail.

**Depends on:** Slices 1–5 all merged

**Checks:**

```bash
# 1. No bare CREATE PROCEDURE in any proc file
grep -rn "^create procedure " DDLs/SCH-usp_*.sql
# expect: zero hits

# 2. All 15 modified files now use CREATE OR ALTER
grep -rl "create or alter procedure" DDLs/SCH-usp_*.sql | wc -l
# expect: ≥ 15 (may be more if unrelated files were already converted)

# 3. All 15 modified files contain the email gate
grep -rl "email_delivery_enabled" DDLs/ | wc -l
# expect: 16 (1 DDL seed + 15 proc preambles)

# 4. Installer has 3+ hits for the new flag
grep -n "EnableEmailAlerts" SQLMonitor/Install-SQLMonitor.ps1 | wc -l
# expect: ≥ 3
```

**Manual smoke test** (requires a SQL Server with the updated objects deployed):
```sql
UPDATE dbo.sma_params SET param_value = '0' WHERE param_key = 'email_delivery_enabled';

DECLARE @before DATETIME = GETDATE();
EXEC dbo.usp_GetAllServerDashboardMail @send_mail = 1, @recipients = 'test@example.com';
SELECT COUNT(*) AS mails_sent
FROM msdb.dbo.sysmail_allitems
WHERE sent_date > @before;
-- expect: 0

UPDATE dbo.sma_params SET param_value = '1' WHERE param_key = 'email_delivery_enabled';
-- re-enable and confirm mail flows with a working Database Mail profile
```

---

## Parallelism map

```
Slice 1 (DDL seed)
    └── Slice 2 (wrapper pivot)      ─┐
    └── Slice 3a (@send_error_mail)   ├── all can run in parallel once Slice 1 is done
    └── Slice 3b (@send_mail)         │
    └── Slice 3c (@send_email)       ─┘
    └── Slice 4 (installer)
Slice 5 (wrapper sample)             -- independent, any time
Slice 6 (verification)               -- after 1–5
```

---

## Next in flow

Continue with [`/kyos:implement`](./implement.md) to execute the slices.

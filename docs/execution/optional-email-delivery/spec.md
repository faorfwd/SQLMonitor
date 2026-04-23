# Spec — Optional Email Delivery (Global Kill Switch)

## Context

Every SQLMonitor stored procedure that raises a notification uses
`msdb.dbo.sp_send_dbmail` directly, and all of them default to sending mail
(`@send_mail = 1`, `@send_error_mail = 1`, `@send_email = 1`). SQL Agent job
steps hardcode those flags to `1` in their `sqlcmd` / `EXEC` commands.

Operators who deploy SQLMonitor into environments where email alerting is not
wanted today have only two options: (a) leave the placeholder recipient
`dba_team@gmail.com` in place and hope Database Mail is not configured, or
(b) edit every job script by hand to pass `@send_mail = 0`. The install-time
`-SkipMultiMailJobSteps` flag only drops 2 of the mail-sending jobs (login
expiry + dashboard mail); it does not silence the other 9 procs or their
wrappers.

Goal: add a single install-time switch that turns email delivery off for the
entire SQLMonitor surface without touching individual job scripts. The
per-proc `@send_mail` parameters remain and still override locally, but a
global kill-switch shortens every code path to a no-op when off.

---

## Feature (plain language)

> As an operator installing SQLMonitor, I can decide at install time whether
> the monitoring stack is allowed to send email. If I turn email off, no
> stored procedure in the DBA database ever calls `sp_send_dbmail`, and the
> installer does not fail prerequisite checks that only matter when mail is
> enabled.

---

## Decisions (locked with user)

| # | Decision | Chosen |
|---|----------|--------|
| 1 | Toggle granularity | **Global kill switch only** — one key, applies everywhere |
| 2 | Storage of the flag | **`dbo.sma_params`** — existing temporal config table |
| 3 | Database Mail profile creation | **Stay prerequisite-only**, but the `SkipMailProfileCheck` validation is auto-skipped when email is disabled |
| 4 | Default when installer runs without the new flag | **Enabled** — preserves current behavior; opt-out, not opt-in |

---

## Behavior

### Install-time

1. `Install-SQLMonitor.ps1` gains one new switch parameter:

   ```powershell
   [bool]$EnableEmailAlerts = $true
   ```

   Defaults to `$true` so existing `Wrapper-InstallSQLMonitor.ps1` scripts
   keep working with no change.

2. During the install step that populates `dbo.sma_params` defaults
   (inside `DDLs/SCH-Create-Inventory-Specific-Objects.sql`, around the block
   that inserts `dba_team_email_id`, `smtp_server`, etc.), a new row is
   upserted:

   | param_key | param_value |
   |-----------|-------------|
   | `email_delivery_enabled` | `'1'` if `$EnableEmailAlerts`, else `'0'` |
   | `email_delivery_enabled_reason` | `'Set by Install-SQLMonitor.ps1 on <utc timestamp>'` |

   The `sma_params` table is temporal, so every flip keeps history.

3. When `-EnableEmailAlerts:$false`, the installer:
   - Sets `$SkipMailProfileCheck = $true` internally, bypassing the
     `msdb.dbo.sysmail_profile` validation at ~line 2260 of
     `Install-SQLMonitor.ps1`.
   - Still installs every mail-related job (scheduling stays the same); the
     runtime gate handles the silencing. No step is skipped based on the new
     flag; existing `-SkipMultiMailJobSteps` is independent and unchanged.

4. On upgrade runs (row already exists in `sma_params`):
   - If the operator **did not** pass `-EnableEmailAlerts`, the existing
     value in `sma_params` is **preserved** (do not clobber with default).
   - If the operator **did** pass the flag explicitly, the stored value is
     overwritten with the new value.

### Interaction with the existing `send_sqlmonitor_job_failure_mail` key

`dbo.sma_params` already ships with a narrower, per-category switch seeded
in `DDLs/SCH-Create-Inventory-Specific-Objects.sql:1550`:

```
('send_sqlmonitor_job_failure_mail','1','When enabled, then job failure
 mail is send to DBA team'),
```

This key is read today by exactly two procs:

- `DDLs/SCH-usp_wrapper_GetAllServerInfo.sql:87`
- `DDLs/SCH-usp_wrapper_GetAllServerCollectedData.sql:92`

Both do:

```sql
select @send_error_mail = convert(bit,p.param_value)
from dbo.sma_params p
where p.param_key = 'send_sqlmonitor_job_failure_mail';
```

This is the same pattern the new spec generalizes — so it acts as the
reference implementation. The new global switch does **not** replace or
rename it; the two keys layer, with the global switch winning.

**Layering rule (AND-semantics, global wins):**

```sql
-- preamble used by every email-sending proc
DECLARE @email_enabled BIT =
    CASE WHEN (SELECT TOP 1 param_value FROM dbo.sma_params
               WHERE param_key = 'email_delivery_enabled') = '1'
         THEN 1 ELSE 0 END;

IF @email_enabled = 0 SET @send_mail = 0;
-- (procs that also consult per-category keys like
-- send_sqlmonitor_job_failure_mail still do so after this line)
```

| `email_delivery_enabled` | `send_sqlmonitor_job_failure_mail` | Job-failure mail sent? |
|--------------------------|------------------------------------|------------------------|
| `'0'`                    | `'0'`                              | No |
| `'0'`                    | `'1'`                              | No (global wins) |
| `'1'`                    | `'0'`                              | No (category off) |
| `'1'`                    | `'1'`                              | Yes (current behavior) |

**Behavior-change summary for this key:**

- **Semantics preserved.** Operators who already rely on
  `send_sqlmonitor_job_failure_mail` see no change unless they also set
  the new global flag to `'0'`.
- **No rename, no removal.** The key stays in place with its default of
  `'1'` and its remark text unchanged.
- **Two wrappers now run two lookups.** The two wrappers above gain the
  global-switch preamble *in addition to* their existing category
  lookup. That is two `sma_params` reads per invocation (fires at most
  once per schedule tick; cost is negligible, but `/kyos:tech` should
  consider folding both into a single read).
- **Naming guidance for future keys** (non-blocking, informational):
  follow the existing `send_<something>_mail` prefix convention when
  adding per-category switches, so operators can discover related keys
  with a prefix query (`WHERE param_key LIKE 'send_%_mail'`).

### Runtime (every email-sending proc)

Every stored procedure in the inventory below gains a short preamble:

```sql
DECLARE @email_delivery_enabled BIT =
    CASE WHEN (SELECT TOP 1 param_value
               FROM dbo.sma_params
               WHERE param_key = 'email_delivery_enabled') = '1'
         THEN 1 ELSE 0 END;

IF @email_delivery_enabled = 0
BEGIN
    SET @send_mail = 0;   -- or @send_error_mail / @send_email
END
```

Only the existing local flag is coerced to `0`; the rest of the proc's logic
runs unchanged. Every existing gate (`IF @send_mail = 1`) naturally
short-circuits. No `sp_send_dbmail` calls are modified in place.

Procs/wrappers that must gain this preamble (17 total):

- `DDLs/SCH-usp_send_login_expiry_emails.sql`
- `DDLs/SCH-usp_GetAllServerDashboardMail.sql`
- `DDLs/SCH-usp_capture_alert_messages.sql`
- `DDLs/SCH-usp_check_sql_agent_jobs.sql`
- `DDLs/SCH-usp_LogSaver.sql`
- `DDLs/SCH-usp_TempDbSaver.sql`
- `DDLs/SCH-usp_collect_wait_stats.sql`
- `DDLs/SCH-usp_collect_file_io_stats.sql`
- `DDLs/SCH-usp_collect_memory_clerks.sql`
- `DDLs/SCH-usp_collect_ag_health_state.sql`
- `DDLs/SCH-usp_run_WhoIsActive.sql`
- `DDLs/SCH-usp_wrapper_GetAllServerCollectedData.sql`
- `DDLs/SCH-usp_wrapper_GetAllServerInfo.sql`
- `DDLs/SCH-usp_wrapper_CollectPrivilegedInfo.sql`
- `DDLs/SCH-usp_wrapper_populate_sma_sql_instance.sql`

Each update follows **Collection Layer Decision 2** (`CREATE OR ALTER
PROCEDURE`), so no stub+ALTER steps are introduced.

### Consumer behavior (unchanged)

- Grafana, sql_exporter, and the Python alert engine are unaffected — none
  of them call `sp_send_dbmail`. The alert engine has its own notification
  channels (Slack, PagerDuty, Email) configured separately in
  `Alerting/SmaAlertPackage/`.

---

## Acceptance criteria

1. **New install param exists and defaults safely.**
   `-EnableEmailAlerts` is documented in `Install-SQLMonitor.ps1`'s param
   help, defaults to `$true`, and `Wrapper-InstallSQLMonitor.ps1` still
   installs cleanly without mentioning it.

2. **Install-time write.**
   After a fresh install with `-EnableEmailAlerts:$false`, querying
   `dbo.sma_params` on the inventory server shows
   `email_delivery_enabled = '0'`.

3. **Runtime gate works.**
   With `email_delivery_enabled = '0'`, executing each of the 17 procs with
   valid parameters must not produce a row in
   `msdb.dbo.sysmail_allitems` (verify `SELECT COUNT(*) FROM
   msdb.dbo.sysmail_allitems WHERE sent_date > @before` is unchanged).

4. **Runtime gate reversibility.**
   Flipping the key to `'1'` (`UPDATE dbo.sma_params SET param_value = '1'
   WHERE param_key = 'email_delivery_enabled'`) and re-running the same
   procs produces mail rows in `sysmail_allitems` (given a working Database
   Mail profile).

5. **Mail profile prerequisite auto-skipped when disabled.**
   With `-EnableEmailAlerts:$false` on an instance that has **no** Database
   Mail profile, `Install-SQLMonitor.ps1` completes step by step without
   raising the "no default mail profile" error.

6. **Per-proc override still wins locally.**
   Calling a proc with `@send_mail = 1` while
   `email_delivery_enabled = '0'` must still **not** send mail — the
   global switch is authoritative. (The preamble forces `@send_mail = 0`.)

7. **Existing `send_sqlmonitor_job_failure_mail` behavior preserved.**
   With `email_delivery_enabled = '1'` and
   `send_sqlmonitor_job_failure_mail` left at its default `'1'`, the two
   wrappers (`usp_wrapper_GetAllServerInfo`,
   `usp_wrapper_GetAllServerCollectedData`) continue to send job-failure
   mail exactly as before. Flipping the category key to `'0'` suppresses
   only those job-failure paths; other alerts still flow.

8. **Global overrides category.**
   With `email_delivery_enabled = '0'`, neither wrapper sends
   job-failure mail regardless of `send_sqlmonitor_job_failure_mail`.

9. **Upgrade preserves operator choice.**
   Running the installer a second time without passing the flag does not
   overwrite an existing `email_delivery_enabled` value in `sma_params`.
   It also does not modify `send_sqlmonitor_job_failure_mail`.

10. **Collection-layer decisions honored.**
    - All modified procs use `CREATE OR ALTER`.
    - No `OPENQUERY` introduced.
    - If any SQL file needs install-time substitution, it uses SQLCMD
      variables, not `.Replace()`.

---

## Unresolved questions

- None that block spec sign-off. Three to revisit at `/kyos:tech` time:
  - **Caching the flag:** reading `sma_params` every collection run is
    cheap, but for high-frequency procs (e.g., `usp_collect_wait_stats`
    every 60 s) we may prefer a session-scoped cache or a fallback default
    when the row is missing. Decide during tech design.
  - **PowerShell collectors with mail hooks (if any):** the current survey
    found none, but tech phase should re-confirm that no `*.ps1` under
    `SQLMonitor/` sends mail directly. If any are found, they must honor
    the same flag (read via `dbo.usp_get_sma_param` or equivalent).
  - **Combine the two `sma_params` reads** in
    `usp_wrapper_GetAllServerInfo` and
    `usp_wrapper_GetAllServerCollectedData` into a single pivoted lookup
    (`email_delivery_enabled` + `send_sqlmonitor_job_failure_mail`) so the
    new preamble does not become N+1 reads per run.

---

## Critical files

| File | Change |
|------|--------|
| `SQLMonitor/Install-SQLMonitor.ps1` | Add `-EnableEmailAlerts` param; write to `sma_params`; conditionally set `$SkipMailProfileCheck` |
| `DDLs/SCH-Create-Inventory-Specific-Objects.sql` | Seed the `email_delivery_enabled` row in `sma_params` defaults block |
| `DDLs/SCH-usp_*.sql` (17 files listed above) | Add preamble that forces local `@send_mail` off when the global key is `'0'`; convert to `CREATE OR ALTER` if not already |
| `Wrapper-Samples/Wrapper-InstallSQLMonitor.ps1` | Add commented-out example of `-EnableEmailAlerts:$false` |

No changes expected in `DDLs/SCH-Job-*.sql` — jobs continue to pass
`@send_mail = 1`; the runtime gate overrides.

---

## Verification plan

1. **Fresh install, email disabled.**
   Run `Install-SQLMonitor.ps1 -EnableEmailAlerts:$false` against a test
   instance that has no Database Mail profile. Expect: no prerequisite
   error; `sma_params.email_delivery_enabled = '0'`.

2. **Runtime silence.**
   `EXEC dbo.usp_GetAllServerDashboardMail @send_mail = 1, @recipients =
   'test@example.com'`. Confirm no new row in `msdb.dbo.sysmail_allitems`.

3. **Re-enable and resend.**
   `UPDATE dbo.sma_params SET param_value = '1' WHERE param_key =
   'email_delivery_enabled'`. Repeat step 2. Confirm a row appears in
   `sysmail_allitems` with status `sent`.

4. **Upgrade preservation.**
   With `email_delivery_enabled = '0'`, re-run installer without the flag.
   Confirm value is still `'0'`.

5. **Regression sweep.**
   Run `grep -n "CREATE PROCEDURE " DDLs/SCH-usp_*.sql` and confirm every
   modified file now reads `CREATE OR ALTER PROCEDURE` (Collection Layer
   Decision 2).

---

## Next in flow

Once this spec is approved, continue with `/kyos:tech` to design:
- the exact preamble helper (inline vs. a new `dbo.usp_is_email_enabled`)
- `sma_params` read caching strategy for high-frequency procs
- PowerShell coverage question

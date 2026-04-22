# Technical Design — Remove hardcoded `grafana` SQL login password

Spec: [hardcoded-grafana-password-removal.md](hardcoded-grafana-password-removal.md)

---

## 1. Approach summary

Eight files change across three independent tracks that can be implemented in any order and verified independently.

| Track | Files | Key technique |
|-------|-------|---------------|
| A — Installer | `Install-SQLMonitor.ps1`, `DCL-[grafana-login].sql`, `Wrapper-InstallSQLMonitor.ps1` | New params + dbatools `New-DbaLogin`/`Set-DbaLogin` + `usp_add_credential`/`usp_update_credential` |
| B — Monitoring job | `check-instance-availability.ps1` | `usp_get_credential` OUTPUT param → `SecureString` → `PSCredential`; one-release fallback |
| C — Linked-server templates | `SCH-Create-LinkedServers-4-AgListenerIP.sql`, `SCH-Create-LinkedServers-4-AgListenerName.sql`, `SCH-Linked-Servers-Sample.sql` | SQLCMD `:setvar` header + `$(GrafanaLoginPassword)` substitution |

---

## 2. Track A — Installer

### 2.1 New parameters (Install-SQLMonitor.ps1 `Param()` block)

Add immediately after the existing `[PSCredential]$SqlCredential` param:

```powershell
[Parameter(Mandatory=$false)]
[securestring]$GrafanaLoginPassword,

[Parameter(Mandatory=$false)]
[switch]$RotateGrafanaLoginPassword
```

`$GrafanaLoginPassword` is `[securestring]` not `[PSCredential]` — the grafana login is a SQL login with a password only, not a credential pair.

### 2.2 Fail-fast validation (Install-SQLMonitor.ps1, after `$Steps2Execute` is finalized)

`$Steps2Execute` is built around line 830. Insert the block below **between** the `$Steps2Execute` finalization and the first `Connect-DbaInstance` call (~line 837). At that point both `$SqlInstanceToBaselineWithOutPort` and `$InventoryServerWithOutPort` are already resolved.

```powershell
# Fail fast: step 56 on inventory server requires GrafanaLoginPassword
$_isInventoryBaseline = ($SqlInstanceToBaselineWithOutPort -eq $InventoryServerWithOutPort)
if ('56__GrafanaLogin' -in $Steps2Execute -and $_isInventoryBaseline -and ($null -eq $GrafanaLoginPassword)) {
    Write-Error ("Install-SQLMonitor.ps1: step '56__GrafanaLogin' on the inventory server requires " +
                 "-GrafanaLoginPassword <securestring>. Supply the parameter or exclude step 56 with -SkipSteps.")
}
```

`Write-Error` with `$ErrorActionPreference='Stop'` (the script's default) terminates immediately. No DDL runs.

### 2.3 DCL file surgery (DDLs/DCL-[grafana-login].sql)

Delete lines 1–4 (the `CREATE LOGIN` block). File after the change starts at the first blank line before `use [master];` (current line 6):

```sql
-- Lines removed (now managed by Install-SQLMonitor.ps1 via New-DbaLogin / Set-DbaLogin):
-- use [master]
-- if not exists (select * from sys.syslogins where name = 'grafana')
--     exec('create login [grafana] with password=N''grafana'', ...');
-- go

use [master];
if exists (select * from sys.sysusers where name = 'grafana')
    exec('drop user [grafana]')
-- ... rest of file unchanged
```

The remaining sections are all idempotent DROP/CREATE user + role assignment + GRANT — safe to re-run on every install.

### 2.4 Step 56 rewrite (Install-SQLMonitor.ps1 ~line 8566)

Replace the current step 56 block with the logic below. The existing `$GrafanaLoginFilePath` variable points to the DCL file, which now contains only the residual DCL.

```powershell
# 56__GrafanaLogin
$stepName = '56__GrafanaLogin'
if ($stepName -in $Steps2Execute) {
    "`n$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', "*****Working on step '$stepName'.."

    $sqlResidualDcl = [System.IO.File]::ReadAllText($GrafanaLoginFilePath).Replace("[DBA]", "[$DbaDatabase]")

    if ($SqlInstanceToBaselineWithOutPort -ne $InventoryServerWithOutPort) {
        # ── Non-inventory: copy login from inventory, then apply residual DCL ──
        try {
            Copy-DbaLogin -Source $conInventoryServer -Destination $conSqlInstanceToBaseline `
                -Login 'grafana' -EnableException
        }
        catch {
            $errMessage = $_.Exception.Message
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'ERROR:', `
                "Error occurred while using Copy-DbaLogin for [grafana].`n`n$errMessage" |
                Write-Host -ForegroundColor Red
            "STOP here, and fix above issue." | Write-Error
        }
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database master -Query $sqlResidualDcl -EnableException
    }
    else {
        # ── Inventory server: create or conditionally rotate ──
        $grafanaLogin = Get-DbaLogin -SqlInstance $conSqlInstanceToBaseline -Login 'grafana'
        $shouldUpsertCredential = $false

        if ($null -eq $grafanaLogin) {
            # Branch 1 — fresh install: create login
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', `
                "Creating [grafana] login (password value redacted).."
            New-DbaLogin -SqlInstance $conSqlInstanceToBaseline `
                -Login 'grafana' `
                -SecurePassword $GrafanaLoginPassword `
                -DefaultDatabase $DbaDatabase `
                -DisablePasswordPolicy `
                -DisablePasswordExpiration `
                -EnableException
            $shouldUpsertCredential = $true
            $credentialIsNew = $true
        }
        elseif ($RotateGrafanaLoginPassword) {
            # Branch 2 — rotation
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', `
                "Rotating [grafana] login password (value redacted).."
            Set-DbaLogin -SqlInstance $conSqlInstanceToBaseline `
                -Login 'grafana' `
                -SecurePassword $GrafanaLoginPassword `
                -EnableException
            $shouldUpsertCredential = $true
            $credentialIsNew = $false
        }
        else {
            # Branch 3 — login exists, no rotation flag: leave login untouched
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', `
                "[grafana] login exists; -RotateGrafanaLoginPassword not set. Login left untouched."
        }

        # Residual DCL — always; recovers grant drift idempotently
        $conSqlInstanceToBaseline | Invoke-DbaQuery -Database master -Query $sqlResidualDcl -EnableException

        # Upsert credential_manager (branches 1 and 2 only)
        if ($shouldUpsertCredential) {
            # Decrypt SecureString to plaintext only for the proc call; nulled immediately after
            $bstr   = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($GrafanaLoginPassword)
            $grafanaPwdPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

            try {
                # Determine add vs update by checking for an existing row
                $credRowCount = ($conInventoryServer | Invoke-DbaQuery -Database $DbaDatabase -Query `
                    "SELECT COUNT(*) AS n FROM dbo.credential_manager WHERE server_ip='*                  ' AND [user_name]='grafana'" `
                    -EnableException).n

                if ($credRowCount -eq 0 -or $credentialIsNew) {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', `
                        "Adding [grafana] credential to dbo.credential_manager (value redacted).."
                    $conInventoryServer | Invoke-DbaQuery -Database $DbaDatabase -Query `
                        "EXEC dbo.usp_add_credential @server_ip='*', @user_name='grafana', @password_string=@pwd, @is_sql_user=1, @remarks='Grafana SQL login managed by Install-SQLMonitor.ps1'" `
                        -SqlParameter @{pwd=$grafanaPwdPlain} -EnableException
                }
                else {
                    "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', `
                        "Updating [grafana] credential in dbo.credential_manager (value redacted).."
                    $conInventoryServer | Invoke-DbaQuery -Database $DbaDatabase -Query `
                        "EXEC dbo.usp_update_credential @server_ip='*', @user_name='grafana', @new_password_string=@pwd, @confirm_forgot_password=1" `
                        -SqlParameter @{pwd=$grafanaPwdPlain} -EnableException
                }
            }
            finally {
                $grafanaPwdPlain = $null
            }
            "$(Get-Date -Format yyyyMMMdd_HHmm) {0,-10} {1}" -f 'INFO:', `
                "dbo.credential_manager upsert for [grafana] complete (value redacted).."
        }
    }
}
```

**Why `@server_ip='*'` padded to 25 chars:** `dbo.credential_manager.server_ip` is `char(25)` — a literal `'*'` stored is right-padded to 25 chars by SQL Server. The `WHERE server_ip='*                  '` check above must match that exact representation, or use `RTRIM`. Use `RTRIM(server_ip) = '*'` in the existence check query to be safe.

**Why `@confirm_forgot_password=1`:** The installer runs as sysadmin. `usp_update_credential` requires sysadmin + `@confirm_forgot_password=1` when `@old_password_string` is not provided. This is the correct path — the installer owns the value, it doesn't need to prove the old password.

**`[System.Runtime.InteropServices.Marshal]` pattern:** Converts `SecureString` → BSTR → plain string in managed memory, then immediately zero-frees the BSTR. The plain string itself lives only for the duration of the `try` block and is nulled in `finally`. This is the standard PowerShell idiom for SecureString materialisation; it minimises the window during which the plaintext is in memory.

### 2.5 Wrapper-InstallSQLMonitor.ps1

Add to the `$params` hashtable:

```powershell
GrafanaLoginPassword = (Read-Host -AsSecureString -Prompt 'Grafana SQL login password')
#RotateGrafanaLoginPassword = $true   # uncomment to rotate on re-run
```

Add a comment block above the call:

```powershell
<#
  GrafanaLoginPassword  — required when step 56__GrafanaLogin is included on the inventory server.
                          Accepts a SecureString. Alternatives:
                            $params.GrafanaLoginPassword = (Get-Credential -UserName grafana -Message 'Grafana password').Password
                            $params.GrafanaLoginPassword = ConvertTo-SecureString $env:GRAFANA_SQL_PWD -AsPlainText -Force
  RotateGrafanaLoginPassword — add this switch to also rotate the live login on a re-run.
                               Without it, an existing login is left untouched (safe default).
  NOTE: After rotating, update the Grafana datasource password in the Grafana UI.
        The installer does NOT propagate the rotation to Grafana automatically.
#>
```

---

## 3. Track B — Monitoring job (check-instance-availability.ps1)

### 3.1 Locate the replacement site

Current lines 50–53:

```powershell
# Create Grafana Credential
$username = "grafana"
$password = ConvertTo-SecureString "grafana" -AsPlainText -Force
$sqlCredential = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $password)
```

Replace with:

```powershell
# Build Grafana PSCredential from dbo.credential_manager (no literal password)
$username = "grafana"
$sqlGetGrafanaCred = @"
DECLARE @pwd varchar(256);
EXEC dbo.usp_get_credential @server_ip='*', @user_name='grafana', @password=@pwd OUTPUT;
SELECT [password] = @pwd;
"@
try {
    $credRow = $conInventoryServer | Invoke-DbaQuery -Database $InventoryDatabase -Query $sqlGetGrafanaCred -EnableException
    $password = ConvertTo-SecureString $credRow.password -AsPlainText -Force
}
catch {
    $errMsg = $_.Exception.Message
    if ($errMsg -like '*No matching credentials found*') {
        # TEMPORARY FALLBACK — remove after all deployments have run the updated installer
        Write-Warning ("$(Get-Date -Format yyyyMMMdd_HHmm) [grafana] credential not found in " +
                       "dbo.credential_manager. Using built-in default for this cycle. " +
                       "Re-run Install-SQLMonitor.ps1 with -GrafanaLoginPassword to register the credential.")
        $password = ConvertTo-SecureString "grafana" -AsPlainText -Force
    }
    else { throw }
}
$sqlCredential = New-Object System.Management.Automation.PSCredential -ArgumentList ($username, $password)
```

The `$conInventoryServer` object is already established at line 30 (a `Connect-DbaInstance` result, authenticated as Windows identity). `$InventoryDatabase` is already a parameter. `$blockGetServerHealth` (line 62+) uses `$Using:sqlCredential` — unchanged.

### 3.2 Why OUTPUT param via T-SQL wrapper, not SqlParameter

`usp_get_credential` is `WITH ENCRYPTION` and takes `@password` as an OUTPUT parameter. `Invoke-DbaQuery` does not support OUTPUT parameters natively (it wraps `SqlCommand.ExecuteReader`, which ignores OUTPUT params). The wrapper `DECLARE @pwd … EXEC … SELECT @pwd` collapses the OUTPUT param into a result set column, which `Invoke-DbaQuery` returns as a normal row.

### 3.3 Fallback removal plan

The fallback `ConvertTo-SecureString "grafana"` literal is a deliberate, scoped survivor for one release cycle. Remove it in the commit that follows the upgrade announcement. The `Write-Warning` message tells operators exactly what action to take. The fallback is annotated with `# TEMPORARY FALLBACK` so it's searchable.

---

## 4. Track C — Linked-server templates

### 4.1 Pattern applied to all three files

Prepend a SQLCMD header block and replace the `'grafana'` password literal. The three files follow different patterns.

**`Sql-Queries/SCH-Create-LinkedServers-4-AgListenerIP.sql`** — uses T-SQL `DECLARE` variables, not SQLCMD. Change is SQLCMD-only for the password:

```sql
-- Add at top of file (before USE [master]):
/*
  SQLCMD variable required before executing this script:
    :setvar GrafanaLoginPassword <YOUR_PASSWORD_HERE>
  Note: if the password contains single quotes, double them (e.g. it''s).
*/
:setvar GrafanaLoginPassword <YOUR_PASSWORD_HERE>
```

Line 9 change:

```sql
-- Before:
EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=@_alias_server,...,@rmtpassword='grafana';
-- After:
EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=@_alias_server,...,@rmtpassword='$(GrafanaLoginPassword)';
```

**`Sql-Queries/SCH-Create-LinkedServers-4-AgListenerName.sql`** — same treatment, line 10.

**`DDLs/SCH-Linked-Servers-Sample.sql`** — uses N-quoted literals throughout:

```sql
-- Add at top:
/*
  SQLCMD variable required:
    :setvar GrafanaLoginPassword <YOUR_PASSWORD_HERE>
  Note: if the password contains single quotes, double them (e.g. it''s).
*/
:setvar GrafanaLoginPassword <YOUR_PASSWORD_HERE>
```

Line 6 change:

```sql
-- Before:
EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=N'YourSqlInstanceNameHere',...,@rmtpassword='grafana'
-- After:
EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=N'YourSqlInstanceNameHere',...,@rmtpassword='$(GrafanaLoginPassword)'
```

Line 11 (`@optvalue = 'Encrypt=yes;TrustServerCertificate=yes;User ID=grafana'`) — `User ID=grafana` is a username reference, not a password. Leave unchanged per spec.

**SQLCMD variable + single-quote escaping:** SQLCMD replaces `$(GrafanaLoginPassword)` as a raw string inside the SQL text before it is sent to the server. A password containing `'` must be doubled by the operator before setting the variable (documented in the header comment). This is the known limitation for SQLCMD-variable injection into SQL string literals, per Decision 4.

---

## 5. Data / contracts

### 5.1 dbatools calls

| Call | Parameters | Notes |
|------|-----------|-------|
| `New-DbaLogin` | `-SqlInstance $con -Login 'grafana' -SecurePassword $GrafanaLoginPassword -DefaultDatabase $DbaDatabase -DisablePasswordPolicy -DisablePasswordExpiration -EnableException` | Creates login. `CHECK_POLICY=OFF, CHECK_EXPIRATION=OFF` preserved per non-goal. |
| `Set-DbaLogin` | `-SqlInstance $con -Login 'grafana' -SecurePassword $GrafanaLoginPassword -EnableException` | Rotates password. dbatools handles T-SQL escaping; no injection risk. |
| `Get-DbaLogin` | `-SqlInstance $con -Login 'grafana'` | Returns `$null` if login absent. Used to branch create vs rotate. |
| `Copy-DbaLogin` | `-Source $conInventoryServer -Destination $con -Login 'grafana' -EnableException` | Non-inventory path unchanged. |

### 5.2 credential_manager proc signatures used

**Add** (first install):
```sql
EXEC dbo.usp_add_credential
    @server_ip    = '*',          -- char(25), right-padded by SQL Server
    @user_name    = 'grafana',
    @password_string = @pwd,      -- passed as SqlParameter, never interpolated
    @is_sql_user  = 1,
    @remarks      = 'Grafana SQL login managed by Install-SQLMonitor.ps1'
```

**Update** (rotation — sysadmin path):
```sql
EXEC dbo.usp_update_credential
    @server_ip              = '*',
    @user_name              = 'grafana',
    @new_password_string    = @pwd,       -- SqlParameter
    @confirm_forgot_password = 1          -- sysadmin bypass; no old_password needed
```

**Existence check** (to decide add vs update):
```sql
SELECT COUNT(*) AS n
FROM dbo.credential_manager
WHERE RTRIM(server_ip) = '*' AND [user_name] = 'grafana'
```

**Get** (check-instance-availability.ps1):
```sql
DECLARE @pwd varchar(256);
EXEC dbo.usp_get_credential @server_ip='*', @user_name='grafana', @password=@pwd OUTPUT;
SELECT [password] = @pwd;
```

### 5.3 Caller permissions for usp_get_credential in the monitoring job

`usp_get_credential` checks `IS_SRVROLEMEMBER('SYSADMIN', caller)`. SQL Agent on the inventory server typically runs under a sysadmin-equivalent service account. If the account is NOT sysadmin, the caller must be `created_by`, `updated_by`, or a `delegate_login_0x` on the row. The installer writes the row with `SUSER_NAME()` as `created_by` (the sysadmin running the installer). If SQL Agent is non-sysadmin, the operator must patch the row:

```sql
EXEC dbo.usp_update_credential
    @server_ip = '*', @user_name = 'grafana',
    @delegate_login_01 = '<SqlAgentServiceAccount>',
    @confirm_forgot_password = 1
```

Document this in `docs/deployment/install.md`.

---

## 6. Risk list

| # | Risk | Severity | Mitigation |
|---|------|----------|------------|
| R1 | `usp_update_credential` row absent on rotation (login created manually, never via installer) | Medium | Existence check before deciding add vs update |
| R2 | `char(25)` padding on `server_ip='*'` causes WHERE mismatch | Medium | Use `RTRIM(server_ip) = '*'` in the existence check |
| R3 | `Invoke-DbaQuery -SqlParameter` not available in older dbatools versions | Low | dbatools >= 1.0 supports `-SqlParameter`; the project already requires a current dbatools |
| R4 | `OnlySteps 56__GrafanaLogin` on inventory without `$GrafanaLoginPassword` — existing automation breaks if validation was not added | High | Fail-fast validation block (section 2.2) — must be confirmed present before merging |
| R5 | `$conInventoryServer` not yet established when step 56 runs on a non-inventory baseline | Low | The connection is set up early in the installer; the step 56 block is after line 837 |
| R6 | Password with `$(` in SQLCMD template breaks variable expansion | Medium | Documented in header comment; operator must double-quote or avoid such chars |
| R7 | SQL Agent runs as non-sysadmin; `usp_get_credential` denies access | Medium | Documented; operator adds `delegate_login_01` after install |
| R8 | `SecureStringToBSTR` materialises password in managed heap; memory not immediately GC'd | Low | `ZeroFreeBSTR` called in `finally` to zero the BSTR; residual risk is GC timing. Acceptable for an on-prem installer. |
| R9 | Fallback literal in `check-instance-availability.ps1` not removed in follow-up commit | Medium | Annotated `# TEMPORARY FALLBACK` + `Write-Warning` message; tracked in security-review.md as open item until removed |

---

## 7. Test strategy

### 7.1 Grep gates (automated, run before merge)

```powershell
# Must all return 0 results
rg -n "password=N''grafana''" DDLs/
rg -n 'create login \[grafana\]' DDLs/ --ignore-case
rg -n 'ConvertTo-SecureString "grafana"' SQLMonitor/
rg -n "@rmtpassword='grafana'" Sql-Queries/ DDLs/
```

### 7.2 Unit-level (no SQL Server required)

- Dot-source `Install-SQLMonitor.ps1` in a pester test harness; call with `-OnlySteps 56__GrafanaLogin` and no `-GrafanaLoginPassword`. Assert `Write-Error` fires (or assert that `$error` is non-empty).
- Confirm `$RotateGrafanaLoginPassword` defaults to `$false` (switch default).

### 7.3 Integration (throwaway SQL Server instance)

Follow the end-to-end steps from the spec verbatim:

1. Fresh inventory install → verify `sys.syslogins`, SSMS connect, `usp_get_credential` output.
2. Run `(dba) Check-InstanceAvailability` → verify no errors, `dbo.instance_details.is_available` updates.
3. Rotate via `-RotateGrafanaLoginPassword` → repeat step 1 with new password; confirm old password fails.
4. Re-run without rotation flag, supplying a *different* password → confirm login password unchanged (old password still works).
5. `OnlySteps 56__GrafanaLogin` with no password param → confirm fails before any DDL.
6. `SkipSteps 56__GrafanaLogin` with no password param → confirm success.
7. Simulate pre-upgrade: manually delete the `credential_manager` row; run `check-instance-availability.ps1` → confirm `Write-Warning` fires, job completes with fallback.

### 7.4 Special-character password test

Use password `P@ss'w"ord$(1)` in steps 1 and 3. Verify:
- Installer path (dbatools): no error — dbatools handles T-SQL escaping.
- `usp_add_credential` / `usp_get_credential` round-trip: returned password matches what was supplied.
- SQLCMD templates: document that this password would need manual escaping (`P@ss''w"ord$(1)` — double the single quote); **do not** test this automatically, just verify the header comment instruction is present.

---

## 8. Implementation order

1. **DCL file** — delete `CREATE LOGIN` lines. Smallest change, no dependencies.
2. **Install-SQLMonitor.ps1** — add params → add validation block → rewrite step 56. Atomic in a single file.
3. **Wrapper-InstallSQLMonitor.ps1** — add param example. Depends on (2) only for correctness.
4. **check-instance-availability.ps1** — independent of (2)/(3). Fallback keeps it safe to merge before the installer is updated in all envs.
5. **Three linked-server templates** — independent of everything. Pure text substitution.
6. **Docs** (`install.md`, `prerequisites.md`, `security-review.md`) — after (2)–(5) are working.

Next: `/kyos:tasks` to break implementation order into discrete tasks.

# Spec — Remove hardcoded `grafana` SQL login password

## Context

The `grafana` SQL login is currently created with the literal password `'grafana'` in [DDLs/DCL-[grafana-login].sql:3](DDLs/DCL-%5Bgrafana-login%5D.sql#L3), with `CHECK_POLICY=off` and `CHECK_EXPIRATION=off`. The same literal is reused in [SQLMonitor/check-instance-availability.ps1:51-52](SQLMonitor/check-instance-availability.ps1#L51-L52) — a script executed every cycle by SQL Agent job `(dba) Check-InstanceAvailability` to probe every monitored instance — and in three linked-server SCH templates. This is finding #1 in [docs/architecture/security-review.md](docs/architecture/security-review.md) (hardcoded credential) and a secondary instance of finding #2 (credentials echoed in verbose logs).

Goal: the password becomes an installer parameter. The active monitoring job that currently hardcodes the same literal is rewired to retrieve the password from the project's existing `dbo.credential_manager` store, so a rotation via the installer does not break it.

Reference: [docs/architecture/collection-layer-decisions.md](docs/architecture/collection-layer-decisions.md) — Decision 4 (SQLCMD variables) applies to the linked-server templates; Decision 3 (upsert idempotency) applies to the credential_manager seed step.

## Goals

1. `Install-SQLMonitor.ps1` accepts the `grafana` login password as a secure parameter instead of reading a literal from a DDL file.
2. No literal password `'grafana'` remains in any file that executes in production (installer-run DDL, the active PS script, and linked-server SCH templates).
3. Password rotation on re-install is **opt-in**. Default re-install behavior stays "create only if missing" — so existing deployments keep working during rollout.
4. The `(dba) Check-InstanceAvailability` job survives a password rotation without requiring a redeploy of the PS script — credentials live in `dbo.credential_manager`, not in the script.
5. The password is never echoed to console, verbose logs, debug logs, or SQL Agent job step text.

## Non-goals

- CHECK_POLICY / CHECK_EXPIRATION hardening on the `grafana` login — kept **off** per current behavior; documented as a separate follow-up in security-review.md.
- Cleanup of [Python-Scripts/connect-2-sqlserver.py](Python-Scripts/connect-2-sqlserver.py) (sample only, not run by the installer).
- Other security-review findings (SMTP password in `DDLs/DDL-[DatabaseMail_Using_GMail].sql`, `AI-Agent/.env__bak`, `xp_cmdshell`, HTTP-in-URL patterns, etc.).
- Grafana UI-side datasource config — Grafana stores its own copy of the password in its provisioning file; rotating via the installer requires the operator to update Grafana's datasource out of band. Documented, not automated.
- Credential-manager schema changes.

## In scope — files that will change

| File | Change |
|------|--------|
| [DDLs/DCL-[grafana-login].sql](DDLs/DCL-%5Bgrafana-login%5D.sql) | Remove `CREATE LOGIN` block (moves into installer). Keep the user / role / grant sections. |
| [SQLMonitor/Install-SQLMonitor.ps1](SQLMonitor/Install-SQLMonitor.ps1) | Add `[securestring]$GrafanaLoginPassword` + `[switch]$RotateGrafanaLoginPassword` params. Step `56__GrafanaLogin` creates/alters login via `New-DbaLogin` / `Set-DbaLogin`, runs the residual DCL, and upserts the credential into `dbo.credential_manager`. |
| [Wrapper-Samples/Wrapper-InstallSQLMonitor.ps1](Wrapper-Samples/Wrapper-InstallSQLMonitor.ps1) | Show how to supply the param securely (`Read-Host -AsSecureString` and commented `Get-Credential` alternatives). |
| [SQLMonitor/check-instance-availability.ps1](SQLMonitor/check-instance-availability.ps1) | Replace hardcoded `ConvertTo-SecureString "grafana"` with a call that pulls the password from `dbo.usp_get_credential` on the inventory server (using the existing `$conInventoryServer` connection) and builds the `PSCredential`. |
| [Sql-Queries/SCH-Create-LinkedServers-4-AgListenerIP.sql](Sql-Queries/SCH-Create-LinkedServers-4-AgListenerIP.sql) | Replace `@rmtpassword='grafana'` with `@rmtpassword='$(GrafanaLoginPassword)'`; add a `:setvar GrafanaLoginPassword <YOUR_PASSWORD_HERE>` header + comment. |
| [Sql-Queries/SCH-Create-LinkedServers-4-AgListenerName.sql](Sql-Queries/SCH-Create-LinkedServers-4-AgListenerName.sql) | Same SQLCMD-variable treatment. |
| [DDLs/SCH-Linked-Servers-Sample.sql](DDLs/SCH-Linked-Servers-Sample.sql) | Same SQLCMD-variable treatment for `@rmtpassword` and the `User ID=grafana` provider string line. |
| [docs/deployment/install.md](docs/deployment/install.md) and [docs/deployment/prerequisites.md](docs/deployment/prerequisites.md) | Document the new parameter, rotation switch, Grafana datasource step, and credential_manager flow. |
| [docs/architecture/security-review.md](docs/architecture/security-review.md) | Mark finding #1 (grafana login) resolved for the installer path; note remaining items still open. |

## Acceptance criteria

### Installer (SQLMonitor/Install-SQLMonitor.ps1)

- New parameter `[Parameter(Mandatory=$false)] [securestring]$GrafanaLoginPassword`.
- New parameter `[Parameter(Mandatory=$false)] [switch]$RotateGrafanaLoginPassword`.
- If step `56__GrafanaLogin` is in `$Steps2Execute` **and** the target is the inventory server **and** `$GrafanaLoginPassword` is `$null`, the script fails fast with a clear message naming the parameter and the step. Fail happens during parameter validation (before any DDL runs), not mid-install.
- If step 56 is skipped (via `SkipSteps` or `OnlySteps` exclusion), the parameter is not required.
- On non-inventory baselines, step 56 stays `Copy-DbaLogin`-only; parameter is optional and ignored.
- On the inventory server, step 56:
  1. If the login does not exist → creates it via dbatools `New-DbaLogin -SecurePassword $GrafanaLoginPassword -DefaultDatabase $DbaDatabase -DisablePasswordPolicy -DisablePasswordExpiration`. Then runs the residual DCL (users, roles, grants).
  2. If the login exists and `-RotateGrafanaLoginPassword` is set → issues `Set-DbaLogin -SqlInstance ... -Login grafana -SecurePassword $GrafanaLoginPassword`. Then runs the residual DCL (idempotent).
  3. If the login exists and `-RotateGrafanaLoginPassword` is **not** set → leaves the login untouched. Still runs the residual DCL (grant drift recovery stays idempotent). `$GrafanaLoginPassword` is ignored — nothing is written to `dbo.credential_manager` in this branch, so a mismatched supplied password cannot desync the stored credential from the live login.
  4. In branches (1) and (2), upsert the credential into `dbo.credential_manager` via `dbo.usp_add_credential` (first install) or `dbo.usp_update_credential` (rotation), with `@server_ip='*'`, `@user_name='grafana'`, `@is_sql_user=1`, `@remarks='Grafana SQL login managed by Install-SQLMonitor.ps1'`.
- The password value is never interpolated into log strings, verbose output, or `Write-Debug`. Logs may say "Setting grafana password (value redacted)", never the value.

### DCL file (DDLs/DCL-[grafana-login].sql)

- Lines 1–4 (`CREATE LOGIN [grafana] WITH PASSWORD=N'grafana', …`) deleted. Password creation is handled by the installer via dbatools (safer than SQLCMD-variable interpolation for arbitrary password strings — no escaping concern).
- Remaining sections (users, roles, grants) unchanged.
- Grep check: `rg -n "create login \[grafana\]" DDLs/` returns zero results.

### Monitoring job script (SQLMonitor/check-instance-availability.ps1)

- Lines 50–53 hardcoded credential block replaced by a call that uses the existing `$conInventoryServer` (already a `Connect-DbaInstance` object, already authenticated to the inventory) to execute `dbo.usp_get_credential @server_ip='*', @user_name='grafana'` and capture the output parameter.
- Converts the returned plaintext into a `[securestring]` then into a `[pscredential]` — same variable name (`$sqlCredential`) so downstream `$blockGetServerHealth` is unchanged.
- Plaintext password never appears in `Write-Output` / `Write-Host` / `Write-Verbose` / job history. Guarded by `$ErrorActionPreference='Stop'` so a missing credential row fails the job loudly.
- Grep check: `rg -n "ConvertTo-SecureString .grafana." SQLMonitor/` returns zero results.

### Linked-server SCH templates

- Each file gains a `:setvar GrafanaLoginPassword <YOUR_PASSWORD_HERE>` header and a leading block comment instructing the operator to override the value before execution (matching Decision 4's pattern).
- The literal `'grafana'` in `@rmtpassword='grafana'` becomes `@rmtpassword='$(GrafanaLoginPassword)'`.
- [DDLs/SCH-Linked-Servers-Sample.sql:11](DDLs/SCH-Linked-Servers-Sample.sql#L11) provider-string `User ID=grafana` is not a password reference — left unchanged; but its password equivalent (if any follow-up line contains one) is parameterized.
- Grep check: `rg -n "@rmtpassword='grafana'" Sql-Queries/ DDLs/` returns zero results.

## Edge cases

- **Fresh install on inventory** — param required; login created; `credential_manager` row created. Grafana operator is told to set the same password in the Grafana datasource UI (docs step).
- **Re-run of installer, same password, no rotation switch** — login untouched, residual DCL re-applied (idempotent grants), credential_manager untouched. No-op from the user's perspective.
- **Re-run with rotation switch but same password** — `Set-DbaLogin` is a no-op for SQL internally; credential_manager update rewrites the same encrypted blob. Harmless.
- **Re-run with rotation switch and new password** — login's password changes; credential_manager row updates; the next execution of `(dba) Check-InstanceAvailability` picks up the new password on its next fire without restart. Grafana UI datasource will fail until operator rotates it there — documented consequence.
- **Non-inventory baseline** — step 56 performs `Copy-DbaLogin` from inventory (existing behavior). No param required. No credential_manager write (credential_manager is an inventory-only table in practice).
- **`OnlySteps 56__GrafanaLogin` on inventory without param** — must fail in parameter validation with a clear message. This is the main regression risk to existing automation.
- **`SkipSteps 56__GrafanaLogin`** — parameter becomes optional; installer proceeds normally. Wrapper scripts that currently use `SkipSteps` to avoid the grafana step keep working unchanged.
- **Passwords with special chars** (`'`, `"`, `$(`, `;`, backtick) — safe in the installer path because `New-DbaLogin -SecurePassword` takes a `SecureString` and dbatools handles the T-SQL escaping; safe in `usp_add_credential` because it takes a `varchar` parameter (no string interpolation). Unsafe in the linked-server SQLCMD templates — documented in the template header that operators must double any single quotes in the password value before running.
- **CHECK_POLICY stays off** — by explicit decision; SQL accepts any non-null value for the password. A weak password is an operator responsibility. Documented in security-review.md follow-up.
- **`(dba) Check-InstanceAvailability` runs under SQL Agent service account** — that account executes `usp_get_credential` via the signed-procedure path. SQL Agent default is SYSADMIN-level on the inventory, so the owner/delegate checks in [usp_get_credential](Credential-Manager/SCH-%5Bdbo%5D.%5Busp_get_credential%5D.sql) bypass cleanly. If the operator has demoted SQL Agent off sysadmin, they must set `@delegate_login_01` on the row to the SQL Agent account — documented.
- **Operator upgrades to this version without running installer yet** — existing deployments keep using literal `'grafana'` as the login password because nothing changes the login; the monitoring job still works because the credential_manager row will be absent, and we must detect this case in `check-instance-availability.ps1`. **Handling:** if `usp_get_credential` throws `'No matching credentials found.'`, the script falls back to creating the PSCredential with the literal `"grafana"` password **for one release cycle only**, writing a single `Write-Warning` line per run instructing the operator to re-run the installer. The fallback is removed in a follow-up commit once rollout completes. This is the only place the literal survives — deliberately, to avoid breaking existing installations.

## Verification

### End-to-end (manual, on a throwaway instance)

1. Fresh install on a new inventory: `Install-SQLMonitor.ps1 -GrafanaLoginPassword (Read-Host -AsSecureString) -InventoryServer localhost …`. Confirm:
   - `SELECT * FROM sys.syslogins WHERE name='grafana'` returns one row.
   - Connect via SSMS with `grafana` + new password → succeeds.
   - `DECLARE @p varchar(256); EXEC dbo.usp_get_credential @server_ip='*', @user_name='grafana', @password=@p output; SELECT @p` returns the new password.
2. Run `(dba) Check-InstanceAvailability` manually from SSMS: confirm it completes without errors and `dbo.instance_details.is_available` flags update.
3. Re-run installer with `-RotateGrafanaLoginPassword` and a new password. Repeat steps 1.2–1.3 with the new password.
4. Re-run installer **without** the rotation switch, supplying a different password. Confirm login password did **not** change (SSMS login with original password still works; the supplied value is ignored).
5. Run installer with `-OnlySteps 56__GrafanaLogin` and no `-GrafanaLoginPassword`: fails during param validation.
6. Run installer with `-SkipSteps 56__GrafanaLogin` and no `-GrafanaLoginPassword`: succeeds.

### Grep gates

```
rg -n "password=N''grafana''" DDLs/                  # 0 results
rg -n 'create login \[grafana\]' DDLs/               # 0 results
rg -n 'ConvertTo-SecureString "grafana"' SQLMonitor/ # 0 results
rg -n "@rmtpassword='grafana'" Sql-Queries/ DDLs/    # 0 results
```

### Reused utilities

- [`dbo.usp_add_credential`](Credential-Manager/SCH-%5Bdbo%5D.%5Busp_add_credential%5D.sql), [`dbo.usp_update_credential`](Credential-Manager/SCH-%5Bdbo%5D.%5Busp_update_credential%5D.sql), [`dbo.usp_get_credential`](Credential-Manager/SCH-%5Bdbo%5D.%5Busp_get_credential%5D.sql) — existing encrypted credential store. `@server_ip='*'` wildcard is already the project convention (see [QRY-Usages-CredentialManager.sql:17](Credential-Manager/QRY-Usages-CredentialManager.sql#L17)).
- dbatools `New-DbaLogin`, `Set-DbaLogin`, `Copy-DbaLogin` — already used in `Install-SQLMonitor.ps1`.
- SQLCMD-variable parameterization — per [docs/architecture/collection-layer-decisions.md](docs/architecture/collection-layer-decisions.md) Decision 4.

## Next step

Run `/kyos:tech` to produce the technical design for the implementation (parameter validation placement, exact dbatools calls, credential_manager upsert semantics, fallback-and-remove plan for `check-instance-availability.ps1`).

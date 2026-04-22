# Tasks — Remove hardcoded `grafana` SQL login password

Spec: [hardcoded-grafana-password-removal.md](hardcoded-grafana-password-removal.md)
Tech design: [hardcoded-grafana-password-removal-tech.md](hardcoded-grafana-password-removal-tech.md)

---

## Implementation order

Tasks 1, 2, 5, 6 are independent and can run in parallel.
Task 3 requires Task 2 complete. Tasks 7–8 run last.

| # | Subject | Track | File(s) | Status | Blocked by |
|---|---------|-------|---------|--------|------------|
| 1 | Strip `CREATE LOGIN` block | A | `DDLs/DCL-[grafana-login].sql` | pending | — |
| 2 | Add `$GrafanaLoginPassword` param + fail-fast validation | A | `SQLMonitor/Install-SQLMonitor.ps1` | pending | — |
| 3 | Rewrite step 56 (`New-DbaLogin` / `Set-DbaLogin` / credential_manager upsert) | A | `SQLMonitor/Install-SQLMonitor.ps1` | pending | #2 |
| 4 | Update wrapper with `GrafanaLoginPassword` param example | A | `Wrapper-Samples/Wrapper-InstallSQLMonitor.ps1` | pending | #2 |
| 5 | Replace hardcoded password with `credential_manager` lookup + fallback | B | `SQLMonitor/check-instance-availability.ps1` | pending | — |
| 6 | Add SQLCMD `$(GrafanaLoginPassword)` to linked-server templates | C | `Sql-Queries/SCH-Create-LinkedServers-4-AgListenerIP.sql`<br>`Sql-Queries/SCH-Create-LinkedServers-4-AgListenerName.sql`<br>`DDLs/SCH-Linked-Servers-Sample.sql` | pending | — |
| 7 | Grep gates — confirm zero literal passwords remain | QA | `DDLs/`, `SQLMonitor/`, `Sql-Queries/` | pending | #1–6 |
| 8 | Docs: delegate_login note + fallback-removal open item | Docs | `docs/upcomming-changes/hardcoded-grafana-password-removal.md` | pending | #1–6 |

---

## Task detail

### Task 1 — Strip `CREATE LOGIN` block (§2.3)
Delete lines 1–4 from `DDLs/DCL-[grafana-login].sql` (the `CREATE LOGIN [grafana] WITH PASSWORD=N'grafana'` block). File after the change starts at the blank line before `use [master];`.

### Task 2 — New params + fail-fast (§2.1, §2.2)
In `SQLMonitor/Install-SQLMonitor.ps1`:
- Add `[securestring]$GrafanaLoginPassword` and `[switch]$RotateGrafanaLoginPassword` to the `Param()` block after `$SqlCredential`.
- Insert fail-fast guard between `$Steps2Execute` finalization (~line 830) and first `Connect-DbaInstance` (~line 837).

### Task 3 — Rewrite step 56 (§2.4, §5.1, §5.2)
Replace the step 56 block (~line 8566) with three-branch logic:
- Non-inventory: `Copy-DbaLogin` + residual DCL.
- Inventory / fresh: `New-DbaLogin` → `usp_add_credential`.
- Inventory / rotate: `Set-DbaLogin` → `usp_update_credential @confirm_forgot_password=1`.
- Inventory / no rotate: leave login, apply residual DCL.
- SecureString → BSTR → plaintext in `try/finally`; zero-free the BSTR immediately.
- Existence check: `RTRIM(server_ip) = '*'`.

### Task 4 — Update wrapper (§2.5)
Add `GrafanaLoginPassword = (Read-Host -AsSecureString ...)` to `$params` hashtable and the explanatory comment block above the call.

### Task 5 — check-instance-availability.ps1 (§3.1–3.3)
Replace lines 50–53 with `usp_get_credential` OUTPUT-param wrapper query + `try/catch` fallback annotated `# TEMPORARY FALLBACK`.

### Task 6 — Linked-server templates (§4.1)
Prepend `:setvar GrafanaLoginPassword` header to all three files; replace `@rmtpassword='grafana'` with `@rmtpassword='$(GrafanaLoginPassword)'`. Leave `User ID=grafana` on line 11 of the Sample file untouched.

### Task 7 — Grep gates (§7.1)
```powershell
rg -n "password=N''grafana''" DDLs/
rg -n 'create login \[grafana\]' DDLs/ --ignore-case
rg -n 'ConvertTo-SecureString "grafana"' SQLMonitor/
rg -n "@rmtpassword='grafana'" Sql-Queries/ DDLs/
```
All must return 0 results.

### Task 8 — Docs (§5.3, R7, R9)
- Document the non-sysadmin SQL Agent `delegate_login_01` step.
- Flag `# TEMPORARY FALLBACK` in `check-instance-availability.ps1` as an open item to remove after all deployments are upgraded.

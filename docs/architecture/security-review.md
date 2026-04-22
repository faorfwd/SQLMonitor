# SQLMonitor Repo Security Triage

Date: 2026-04-22

This is a lightweight, repository-level security triage (not a full pentest). Findings are prioritized by likely real-world risk and fix effort.

## Summary (highest priority first)

1. Remove committed credential-like files and hardcoded passwords; add ignore rules to prevent reintroduction.
2. Stop logging secrets (even in “verbose” mode); treat logs as sensitive and assume they can leak.
3. Avoid putting credentials in URLs; use HTTPS for external calls; reduce exposure to MITM/log scraping.
4. Avoid machine-wide PowerShell policy changes and repeated `ExecutionPolicy Bypass` patterns; tighten script execution model.
5. Treat `xp_cmdshell` as high-risk: default disabled, gated, and tightly permissioned if absolutely required.
6. Avoid “download then execute” without pinning/integrity verification.

## Findings

### 1) Committed credential-like material / hardcoded passwords

**Severity:** High

**Impact:** Credential disclosure, unauthorized access, lateral movement.

**Evidence (examples):**
- `AI-Agent/.env__bak` contains placeholder API key-like values and a SQL login password (e.g. `SQLMONITOR_LOGIN_PASSWORD=...`).
- `DDLs/DCL-[grafana-login].sql` creates a SQL login `grafana` with password `grafana` and disables password policy/expiration.
- `DDLs/DDL-[DatabaseMail_Using_GMail].sql` includes a hardcoded SMTP password.

**Fix:**
- Delete or redact these files before distribution; replace with templates (e.g., `.env.example`) that contain no real secrets.
- Expand ignore rules to cover backup/env variants (e.g., `.env*`, `*.bak`, `*__bak`), and add a secret-scanning pre-commit/CI gate.
- For SQL logins: require strong passwords, policy enforcement, and consider using integrated auth / service principals where possible.

**Verification:**
- `rg -n -S -i "(password|api[_-]?key|secret|token|-----BEGIN|\\bsk-[A-Za-z0-9_-]{10,})" .`
- Ensure CI fails if secrets are detected.

### 2) Secret leakage via logging (verbose output)

**Severity:** High

**Impact:** Password/token leakage into console logs, CI logs, or log files.

**Evidence (example):**
- `Alerting/SQLMonitorAlertEngineApp.py` prints `login_password` when `verbose` is enabled.

**Fix:**
- Never log secrets; mask/redact values (`****`) or omit entirely.
- Ensure “debug/verbose” modes do not print sensitive data.

**Verification:**
- Run with verbose enabled and confirm no secrets appear.

### 3) Insecure external call + credentials in URL

**Severity:** Medium–High

**Impact:** Credentials can leak via terminal history, proxies, logs, and referrers; HTTP calls can be MITM’d.

**Evidence (example):**
- `SQLMonitor/Update-SQLMonitorIP.ps1` calls `http://ipinfo.io/json` (no TLS).
- The same script builds a URL that includes a password in the query string and echoes the URL.

**Fix:**
- Use HTTPS for external calls.
- Avoid credentials in URLs; use headers/body, or a client that supports basic auth safely.
- Do not print full URLs containing secrets.

**Verification:**
- Confirm scripts only use `https://` and do not output credential-bearing values.

### 4) Unsafe PowerShell execution defaults

**Severity:** Medium

**Impact:** Broadly weakens execution posture and increases the blast radius of script tampering.

**Evidence (examples):**
- `SQLMonitor/Install-SQLMonitor.ps1` sets machine scope execution policy to `Unrestricted`.
- Various job scripts run PowerShell with `-executionpolicy bypass`.

**Fix:**
- Avoid `Set-ExecutionPolicy -Scope LocalMachine ...` in installer flows.
- Prefer code signing + `AllSigned`/`RemoteSigned`, or scoped execution policy for the current process only.
- Reduce reliance on `ExecutionPolicy Bypass` for scheduled tasks; use a signed deployment story.

**Verification:**
- Install flow does not change machine policy; scheduled execution works without bypass.

### 5) Use of `xp_cmdshell` / high-priv SQL features

**Severity:** Medium–High (context-dependent)

**Impact:** If an attacker gains SQL-level execution, `xp_cmdshell` can enable OS-level code execution.

**Evidence:**
- Repo contains scripts that assume or reference `xp_cmdshell` usage.

**Fix:**
- Default: keep `xp_cmdshell` disabled.
- If required: tightly restrict permissions, isolate the instance, monitor usage, and document compensating controls.
- Prefer agent/collector architectures that don’t require SQL Server to shell out.

**Verification:**
- Confirm `xp_cmdshell` remains disabled in non-lab deployments; audit/alert on enablement and usage.

### 6) “Download then execute” without pinning/integrity checks

**Severity:** Medium

**Impact:** Supply-chain compromise could lead to remote code execution or malicious SQL execution.

**Evidence (example):**
- `SQLMonitor/sqlserver-versions-update.ps1` downloads SQL from GitHub raw content and executes it.

**Fix:**
- Pin to a specific commit/tag and validate integrity (checksum/signature) before executing.
- Prefer vendoring a reviewed copy of critical scripts, or fetching releases with verified signatures.

**Verification:**
- Script fails closed on integrity mismatch; runtime uses pinned references.

## Recommended next steps

1. Remove/redact committed secret-like files and hardcoded passwords; add template replacements.
2. Add automated secret scanning (CI + pre-commit).
3. Redact sensitive logging in Python/PowerShell; add a “safe logging” helper.
4. Review installer defaults: eliminate machine-wide policy changes; document a signed execution model.
5. Document and gate any `xp_cmdshell`-dependent paths, with a safe default off.


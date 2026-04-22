# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

---

## Managed bootstrap (kyos-cli)

- `.kyos/claude/` is the **managed source layer** — do not edit files there directly.
- `.claude/` is the **repo-owned customization layer** for overrides and local rules.
- `npx kyos-cli --init` to install or analyze the base structure.
- `npx kyos-cli --apply` after reviewing a proposal.

---

## Architecture overview

SQLMonitor is an open-source SQL Server monitoring stack with three tiers:

1. **Collection tier** — SQL Agent jobs (T-SQL, PowerShell, Python) capture metrics from each monitored instance's DMVs, OS/Perfmon, XEvents.
2. **Storage tier** — Metrics land in the `DBA` database: hourly-partitioned, page-compressed tables per monitored instance; inventory aggregates in `dbo.all_server_*` Memory-Optimized tables on the inventory server.
3. **Consumption tier** — Grafana dashboards, the Python alert engine, `sql_exporter` (Prometheus), and an AI agent all read from those tables.

```
Monitored instance DBA DB  -->  Linked server  -->  Inventory DBA DB  -->  Grafana / Alert Engine / Prometheus
                                                                       -->  AI Agent (LangChain + Ollama)
```

---

## Key directories

| Directory | What lives there |
|---|---|
| `SQLMonitor/` | Deployment orchestration. `Install-SQLMonitor.ps1` (59 named steps), `Remove-SQLMonitor.ps1` (145+ steps), all PowerShell metric collectors. |
| `DDLs/` | All SQL objects: `SCH-Create-All-Objects.sql`, `SCH-Create-Inventory-Specific-Objects.sql`, 60+ `SCH-Job-*.sql`, 40+ `SCH-usp_*.sql`, partition scheme scripts. |
| `Grafana-Dashboards/` | 17+ dashboard JSONs + Grafana alert rules/contact-points YAML. |
| `Alerting/` | Python alert engine: Flask + APScheduler daemon (`SQLMonitorAlertEngineApp.py`), `SmaAlertPackage/` (alert classes + common functions), `Dockerfile`, `requirements.txt`. |
| `sql_exporter/` | Optional Prometheus path: `sql_exporter.yml` + `mssql_*.collector.yml` + `windows_exporter_config.yml`. |
| `TSQLTextNormalizer/` | C# CLR assembly that normalizes and hashes T-SQL text; used by XEvent workload collectors. |
| `AI-Agent/` | LangChain + Ollama + Streamlit agent over the inventory database. |
| `Credential-Manager/` | SQL-based credential store (`dbo.credential_manager` + stored procs) used by Python scripts. |
| `Wrapper-Samples/` | Template wrapper scripts. **Copy to `Private/` and fill in your environment before running**; `Private/` is git-ignored. |
| `Python-Scripts/` | Ancillary scripts: multi-server runners, alert raisers, Cloudflare DDNS, etc. |

---

## Install / remove

```powershell
# Copy the template, fill in credentials and instance name, then run:
Copy-Item .\Wrapper-Samples\Wrapper-InstallSQLMonitor.ps1 .\Private\Wrapper-InstallSQLMonitor.ps1
.\Private\Wrapper-InstallSQLMonitor.ps1 -Verbose

# To remove:
Copy-Item .\Wrapper-Samples\Wrapper-RemoveSQLMonitor.ps1 .\Private\Wrapper-RemoveSQLMonitor.ps1
.\Private\Wrapper-RemoveSQLMonitor.ps1 -Verbose
```

`Install-SQLMonitor.ps1` supports `SkipSteps`, `OnlySteps`, `StartAtStep`, and `StopAtStep` to run a subset of the 59 install steps.

---

## Alert engine

**Run manually (development):**
```bash
cd Alerting
python -m venv AlertEngineVenv
AlertEngineVenv\Scripts\activate.bat   # Windows
# source AlertEngineVenv/bin/activate  # Linux/Mac
pip install -r requirements.txt
python SQLMonitorAlertEngineApp.py --inventory_server <host> --verbose True --login_password '<pwd>'
```

**Run in a container (Podman):**
```bash
cd Alerting
podman build -t sqlmonitor-alert-engine .
podman run --replace -d \
  -e inventory_server='sqlmonitor' \
  -e login_password="$MSSQLPASSWORD" \
  --name sqlmonitor-alert-engine -p 5000:5000 sqlmonitor-alert-engine
podman logs -f sqlmonitor-alert-engine
```

Alert classes live in `Alerting/SmaAlertPackage/AlertClasses/` (one file per alert, e.g. `SmaCpuAlert.py`). Common helpers (DB connection, credential lookup, Slack/PagerDuty sending) are in `Alerting/SmaAlertPackage/CommonFunctions/`.

---

## Prometheus / sql_exporter

`sql_exporter` runs as a Windows service or a `launchd` daemon on Mac:
```powershell
# Windows service
New-Service -Name "sql_exporter" `
  -BinaryPathName "E:\Github\SQLMonitor\sql_exporter\sql_exporter.exe --config.file E:\Github\SQLMonitor\sql_exporter\sql_exporter.yml" `
  -StartupType Automatic -DisplayName "SQL Exporter for Prometheus"
Start-Service sql_exporter
# Validate: http://localhost:9399/metrics
```

Collector files (`mssql_*.collector.yml`) define the PromQL metric name, SQL query, and labels. The main config file (`sql_exporter.yml`) lists which collectors to load.

---

## AI Agent

```bash
cd AI-Agent
python -m venv venv
source venv/bin/activate   # or venv\Scripts\activate on Windows
pip install -r requirements.txt
# Set env vars: OPENAI_API_KEY_SECRET or configure Ollama endpoint
streamlit run sql_db_agent_4_sqlmonitor_using_ollama.py
```

---

## Database naming conventions

- **Tables** — `dbo.snake_case` (e.g. `dbo.wait_stats`, `dbo.file_io_stats`); inventory aggregates use `dbo.all_server_*`.
- **Stored procedures** — `dbo.usp_collect_*`, `dbo.usp_get_*`, `dbo.usp_compute_*`, `dbo.usp_run_*`, `dbo.usp_purge_*`.
- **SQL Agent jobs** — `(dba) <Verb>-<Subject>`, all in category `(dba) SQLMonitor`.
- **Grafana login** — `grafana` with `db_datareader` on `DBA`.

## CLR assembly

`TSQLTextNormalizer.dll` is a signed CLR assembly that exposes `dbo.fn_get_hash_for_string` and is used by `dbo.usp_collect_xevent_metrics_hashed` to normalize and hash T-SQL text (so workload views group by logical query, not parameter literals). It is registered during install step `2__AllDatabaseObjects` via `TSQLTextNormalizer/SCH-Assembly-[SQLMonitorAssembly].sql`.

---

## Docs site

Full documentation: <https://imajaydwivedi.github.io/SQLMonitor/>

Built with MkDocs Material (`mkdocs.yml`). Sources live in `docs/`. The architecture pages (`docs/architecture/`) are the best place to understand data flow and component relationships.

# Collection Layer Architecture Decisions

Decisions locked in for the SQL monitoring collection layer. These cover stored procedure management, remote SQL execution, SQL Agent jobs (T-SQL + PowerShell), and linked servers. They should not be re-debated per-task.

---

## Current State (key facts)

| Area | Current pattern | Problem |
|------|----------------|---------|
| **Data flow** | Inventory PULLs from monitored via linked server + OPENQUERY | Inventory must reach every monitored instance; firewall rules are one-directional; outage on inventory breaks aggregation |
| **Consumer scope** | Grafana/Prometheus already query inventory only | Already correct; no change needed |
| Proc deployment | Stub + `ALTER PROCEDURE` | Two-step, only partially idempotent |
| Job idempotency | Drop only if `APP_NAME() = SSMS` | Automation path (`Invoke-DbaQuery`) never drops → stale jobs survive upgrades |
| Job parameterization | `.Replace()` on raw file content in `Install-SQLMonitor.ps1` | Fragile; breaks if strings appear elsewhere in DDL |
| PowerShell execution | dbatools `Invoke-DbaQuery` / `Write-DbaDbTableData` | Consistent — keep it |
| Credential handling | `dbo.credential_manager` temporal table | Consistent — keep it |

---

## Decision 1 — Data Flow: PUSH from Monitored Instances to Inventory

**Rule:** Collected metrics flow FROM each monitored instance TO the inventory server. The inventory server does not reach into monitored instances to collect data.

**Target topology:**

```
Monitored Instance
  ├─ T-SQL Agent jobs → INSERT into inventory.DBA.dbo.<table>  (via linked server TO inventory)
  └─ PowerShell collectors → Write-DbaDbTableData -SqlInstance <InventoryServer>

Inventory Server (DBA database)
  ├─ Receives pushed data; no OPENQUERY collection jobs
  └─ Exposes all_server_* aggregation tables to consumers

Consumer layer (Grafana, sql_exporter/Prometheus, alert engine)
  └─ Query inventory server only — never individual monitored instances
```

**Mechanism by collector type:**

| Collector type | Push mechanism | Change required |
|----------------|---------------|----------------|
| PowerShell (`disk-space-collector.ps1`, `perfmon-collector-push-to-sqlserver.ps1`) | `Write-DbaDbTableData -SqlInstance $InventoryServer` | Change job step `-SqlInstance` from `localhost` to inventory server name |
| T-SQL (`usp_collect_*`) | INSERT via reverse linked server (monitored → inventory) | Add `@data_destination_server` parameter; route INSERT through linked server when destination ≠ local |

**`data_destination_sql_instance` is authoritative:** The column in `dbo.instance_details` already exists. Its default changes from `@@SERVERNAME` (local) to the inventory server name. All collection procs and job steps read this value to determine where to write.

**What is eliminated:**

- `usp_wrapper_GetAllServerCollectedData` OPENQUERY pull loop — replaced by simpler aggregation of already-landed data
- Linked servers FROM inventory TO monitored instances are deprecated for collection purposes; kept only for metadata/discovery (`usp_populate_sma_sql_instance` for instance registration)

**Reverse linked server (monitored → inventory):**

- Each monitored instance creates a linked server pointing TO the inventory server (new install step)
- The linked server account needs INSERT permission on the relevant DBA tables on inventory

**Consumer isolation is already correct:** Grafana uses a single datasource (inventory server), sql_exporter targets inventory `all_server_*` tables, and the alert engine queries inventory. No changes needed on the consumer side.

---

## Decision 2 — Stored Procedure Deployment: `CREATE OR ALTER`

**Rule:** All new and modified procedures use `CREATE OR ALTER PROCEDURE` (SQL Server 2016+). The stub+ALTER two-step is retired for new work. Existing procedures are migrated opportunistically when touched, not in bulk.

**Why:** Eliminates the race condition where a stub runs without error but ALTER fails silently. `CREATE OR ALTER` is atomic and works identically in `Invoke-DbaQuery`, SSMS, and `sqlcmd`.

**Boundary:** Scripts in `DDLs/SCH-usp_*.sql` are the authoritative source. No procedure logic in `Install-SQLMonitor.ps1`.

---

## Decision 3 — SQL Agent Job Idempotency: Explicit Upsert

**Rule:** Job scripts must handle all three states: does not exist (create), exists with same definition (no-op), exists with changed definition (update). Remove the `APP_NAME() = SSMS` gate.

```sql
IF NOT EXISTS (SELECT 1 FROM msdb.dbo.sysjobs WHERE name = N'(dba) Collect-X')
    EXEC msdb.dbo.sp_add_job @job_name = N'(dba) Collect-X', ...
ELSE
    EXEC msdb.dbo.sp_update_job @job_name = N'(dba) Collect-X', ...
-- same for sp_add_jobstep / sp_update_jobstep
```

**Why:** Automation deployments (CI, re-runs of `Install-SQLMonitor.ps1` with `OnlySteps`) must be idempotent without human intervention.

**Boundary:** PowerShell job steps use `CmdExec` subsystem. T-SQL job steps use `TSQL` subsystem. No mixing within a single step.

---

## Decision 4 — SQL Script Parameterization: SQLCMD Variables

**Rule:** New SQL scripts that need environment-specific values (instance name, database name, file paths) use SQLCMD variables (`:setvar` / `$(VarName)`) instead of placeholder strings. `Install-SQLMonitor.ps1` passes these via `Invoke-DbaQuery -Variable` or `sqlcmd -v`.

```sql
:setvar DbaDatabase "DBA"
:setvar InventoryServer "inventory-server"
USE [$(DbaDatabase)]
INSERT [$(InventoryServer)].[$(DbaDatabase)].[dbo].[wait_stats] ...
```

**Why:** String `.Replace()` in PowerShell matches substrings, cannot validate completeness, and breaks when DDL refactoring changes the literal text. SQLCMD variables fail loudly if undefined.

**Migration:** Existing scripts are migrated when touched for another reason, not in bulk.

---

## Decision 5 — Remote Execution: Four-Part INSERT for T-SQL Push

**Rule:** T-SQL collection procedures that push to a remote destination use direct `INSERT [LinkedServer].[DBA].[dbo].[table]` (four-part name). `OPENQUERY` for data insertion is prohibited.

**OPENQUERY stays only for:** Instance discovery queries in `usp_populate_sma_sql_instance` where the query text must be built dynamically.

**`EXEC ... AT [linked_server]`** is prohibited in all new code.

---

## Decision 6 — PowerShell Remote Execution: dbatools Only

**Rule:** All PowerShell-based SQL interaction uses dbatools (`Connect-DbaInstance`, `Invoke-DbaQuery`, `Write-DbaDbTableData`). Raw `Invoke-Sqlcmd`, `SqlConnection`/`SqlCommand`, and `Invoke-Command` for SQL execution are prohibited in new collectors.

**Credential binding:** Collector scripts that need SQL auth retrieve credentials from `dbo.credential_manager` via `usp_get_credential` and construct a `[pscredential]` object. Windows auth is the default; SQL auth is the exception.

---

## Out of Scope

- **Prometheus / sql_exporter** — separate layer; no changes needed (already queries inventory only)
- **Alert engine (Python/Flask)** — separate process; its DB access pattern is not affected
- **Grafana dashboards** — no changes needed (already queries inventory only)
- **Linked server provider (SQLNCLI → MSOLEDBSQL)** — separate change request
- **Credential-Manager schema** — stable; no changes

---

## Critical Files

| File | Change needed |
|------|--------------|
| `DDLs/SCH-usp_collect_*.sql` (new/modified) | Add `@data_destination_server` param; four-part INSERT when destination ≠ local; `CREATE OR ALTER` |
| `DDLs/SCH-Job-*.sql` (new/modified) | Replace SSMS gate with explicit upsert; pass inventory server to PowerShell steps |
| `SQLMonitor/disk-space-collector.ps1` | Job step passes inventory server as `-SqlInstance` |
| `SQLMonitor/perfmon-collector-push-to-sqlserver.ps1` | Same `-SqlInstance` change |
| `SQLMonitor/Install-SQLMonitor.ps1` | New install step: reverse linked server (monitored → inventory); retire OPENQUERY aggregation jobs |
| `DDLs/SCH-Create-Inventory-Specific-Objects.sql` | Change default of `data_destination_sql_instance` from `@@SERVERNAME` to inventory server name |

---

## Verification Checklist

When reviewing any new or modified collection layer file:

- [ ] **PUSH confirmed** — query inventory tables after a collection run; data must appear without a manual aggregation job
- [ ] **No OPENQUERY for writes** — grep for `OPENQUERY` in new `SCH-usp_collect_*.sql` files → zero results
- [ ] **`CREATE OR ALTER`** — grep for `CREATE PROCEDURE` without `OR ALTER` in new/modified files → zero results
- [ ] **No SSMS gate** — grep for `APP_NAME()` in new/modified job files → zero results
- [ ] **dbatools only** — grep for `Invoke-Sqlcmd` or `new SqlConnection` in new collectors → zero results
- [ ] **Consumer isolation** — Grafana and sql_exporter datasources point only to inventory server

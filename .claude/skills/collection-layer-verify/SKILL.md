---
name: collection-layer-verify
description: Use when modifying DDLs/SCH-usp_collect_*.sql, DDLs/SCH-Job-*.sql, or SQLMonitor/*.ps1 collectors — before committing or claiming collection layer work is complete
---

# Collection Layer Verify

## Overview

Run the 6-item compliance checklist from `docs/architecture/collection-layer-decisions.md` against modified files. Any FAIL blocks the commit.

## Checks

Use the Grep tool for each check. Report results in a table.

### 1. No OPENQUERY in collect procs
```
pattern: OPENQUERY
path: DDLs/
glob: SCH-usp_collect_*.sql
```
Expect: 0 matches → PASS

### 2. CREATE OR ALTER only (no bare CREATE PROCEDURE)
```
pattern: CREATE PROCEDURE
path: DDLs/
glob: SCH-usp_collect_*.sql
```
Expect: 0 matches → PASS  
*(All procs must use `CREATE OR ALTER PROCEDURE`; bare `CREATE` never matches this check when done right.)*

### 3. No APP_NAME() gate in job files
```
pattern: APP_NAME\(\)
path: DDLs/
glob: SCH-Job-*.sql
```
Expect: 0 matches → PASS

### 4. No Invoke-Sqlcmd in collectors
```
pattern: Invoke-Sqlcmd
path: SQLMonitor/
glob: *.ps1
```
Expect: 0 matches → PASS

### 5. No raw SqlConnection in collectors
```
pattern: new SqlConnection
path: SQLMonitor/
glob: *.ps1
```
Expect: 0 matches → PASS

### 6. @data_destination_server in every modified collect proc
For each modified `SCH-usp_collect_*.sql`, grep the file for `@data_destination_server`. Expect ≥ 1 match per file → PASS.

## Output Format

| Check | Result | Status |
|-------|--------|--------|
| No OPENQUERY | 0 matches | PASS |
| CREATE OR ALTER | 0 matches | PASS |
| No APP_NAME() | 0 matches | PASS |
| No Invoke-Sqlcmd | 0 matches | PASS |
| No SqlConnection | 0 matches | PASS |
| @data_destination_server | present in all modified procs | PASS |

Any FAIL = do not commit. Fix the violation, re-run checks.

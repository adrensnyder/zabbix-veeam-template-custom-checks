# VeeamCustomChecks_ZabbixTemplate

I'm not satisfied with the [official template of Veeam Backup and Replication](https://www.zabbix.com/integrations/veeam) for Zabbix, as the API does not provide all the information needed to properly monitor backups in a efficiently way.
For this reason, I created my own template, which connects directly to MS SQL or PostgreSQL.

## Database Versions
- **v11** uses mostly SQL Server.
- **v12/v13** uses PostgreSQL and requires psqlodbc to be installed in both 32-bit and 64-bit versions from [here](https://www.postgresql.org/ftp/odbc/releases/).
  - Default port: 5432
  - Default user: `postgres` (without password)

## Status / Result / State Codes

The template uses value maps to normalize the meaning of numeric values:

### Result
- `0` = Success
- `1` = Warning
- `2` = Failed
- `-1` = None / Unknown

### Status
- `0` = Success
- `2` = Failed
- `3` = Failed / Retry / Connection error
- `6` = Other / Unknown

### State
- `5` = In progress
- `-1` = Idle / Unknown

### Monitoring
- `1` = Monitoring enabled
- `0` = Disabled / Deleted
- `-1` = Manually disabled by macro (Blacklist)

## Configurable MACROS

The template includes several configurable MACROS with descriptions of their usage, including:

### Connection settings
- `{$VEEAM.CST.DRIVER}`  
  Example values: `SQL Server`, `PostgreSQL ANSI`, `PostgreSQL Unicode`.  
- `{$VEEAM.CST.SRV}`  
  For SQL Server you may need to include the instance (e.g. `localhost\VEEAMSQLxxxx`).  
- `{$VEEAM.CST.PORT}`  
  PostgreSQL default is `5432`. For SQL Server you typically leave it empty.
- `{$VEEAM.CST.DB}`  
  Database name (commonly `VeeamBackup`).
- `{$VEEAM.CST.USER}` / `{$VEEAM.CST.PASS}`  
  If empty, the collector tries to use Windows Authentication (where applicable).

### Manual exclusions (Blacklist)
- `{$VEEAM.CST.BLACKLIST}`  
  A pipe-separated list of names to treat as **manually disabled** without removing them from discovery.

### Alert tuning (time windows)
- `{$VEEAM.CST_ERR_CHECK01..04}`  
  Time window used for error checks (the template uses `min()` in triggers).
- `{$VEEAM.CST_NODATA01..04}`  
  Time window used for `nodata()` checks.
- `{$VEEAM.CST_DLY_CHECK01..04}`  
  Delay thresholds (days) for “delayed execution” triggers.

## Paths and Configuration

- Change the path of the executable in the item `[VEEAM-CST] Collect Data` if needed.
- By default the template runs:
  - `c:\zabbix_agent\VeeamCustomChecks.exe`
  - with parameters: `--driver`, `--server`, `--port`, `--database`, `--user`, `--password`, `--blacklist`.  

> The `[VEEAM-CST] Collect Data` item is scheduled and is responsible for producing the data consumed by the trapper items and discovery rules.

## Low Level Discovery (LLD)

The template uses multiple LLD rules to discover and monitor different Veeam object types, including:

- **Agent Backup**
- **Agent Policy**
- **BackupJob**
- **BackupSync**
- **Endpoint**
- **Repository**
- **Tape File**
- **Tape VM**
- **VM** (per job / per VM)

Each discovered object exposes a consistent set of signals (where applicable), typically:
- `monitored` (Monitoring value map)
- `result` / `status` (Result/Status value map)
- `state` (State value map)
- `reason` (text)
- `creation time`, `end time`, `duration`, `datediff`, `next_run_time`

## Items (summary)

To keep the template lightweight but useful, items are structured in two layers:

### 1) Global counters / health
- Counts of discovered objects (and monitored objects) for each category (jobs, endpoints, repositories, tape, etc.).
- A text item for collector errors:
  - `backup.veeam.customchecks.dataerrors`

### 2) Per-object items (via LLD)
Per-object monitoring focuses on:
- latest status/result/state
- reason/message correlation
- delay detection (`datediff`)
- monitored state (including manual disable via blacklist)

## Triggers (noise-resistant)

Triggers are designed to:
- detect failures using numeric result/status values,
- reduce false positives by using `reason` correlation (e.g. filtering transitional “Processing” states where relevant),
- respect the `monitored` state so disabled/blacklisted objects do not alert,
- provide escalating severity via configurable time windows (`ERR_CHECK*`, `NODATA*`, `DLY_CHECK*`).

The differences from the official template are that this version allows you to directly monitor the latest backup status, identify any jobs that ended with errors or warnings, and detect any backups that are delayed for any reason.
The original template does not monitor all types (IDs) of backup jobs and only retrieves the various sessions. If a session for a job returns an error but the next one is successful, the alert remains in Zabbix for a while, resulting in a false positive.
 
# Copyright
This project is an unofficial Zabbix template for monitoring Veeam backup jobs. It is not affiliated with, endorsed, or supported by Veeam Software in any way.  
“Veeam” and related trademarks are the property of Veeam Software.  

This template and the included executable do not modify Veeam software in any way, nor do they interfere with its licensing.  
All data is retrieved in read-only mode from the internal database and sent to Zabbix for monitoring purposes.  

The goal is to provide more detailed monitoring until Veeam expands its public API capabilities, at which point an API-only solution may be developed.

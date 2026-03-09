;###################################################################
;# Copyright (c) 2026 AdrenSnyder https://github.com/adrensnyder
;#
;# Permission is hereby granted, free of charge, to any person
;# obtaining a copy of this software and associated documentation
;# files (the "Software"), to deal in the Software without
;# restriction, including without limitation the rights to use,
;# copy, modify, merge, publish, distribute, sublicense, and/or sell
;# copies of the Software, and to permit persons to whom the
;# Software is furnished to do so, subject to the following
;# conditions:
;#
;# The above copyright notice and this permission notice shall be
;# included in all copies or substantial portions of the Software.
;#
;# DISCLAIMER:
;# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
;# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
;# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
;# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
;# HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
;# WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
;# FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
;# OTHER DEALINGS IN THE SOFTWARE.
;###################################################################

#AutoIt3Wrapper_Res_Description=VeeamCustomChecks
#AutoIt3Wrapper_Res_Fileversion=1.5
#AutoIt3Wrapper_Res_ProductVersion=
#AutoIt3Wrapper_Res_Language=
#AutoIt3Wrapper_Res_LegalCopyright=Created by AdrenSnyder

#include <File.au3>
#include <Array.au3>
#include <Constants.au3>
#include <Date.au3>
#include <ADO.au3>
#AutoIt3Wrapper_Change2CUI=y
#RequireAdmin

#Region Globals
Global $vDefaultConf[2]
Global $ConfString = ""
Global $vDefaultConf_Size

Global $LogName = ""
Global $LogDir = ""
Global $LogFile = ""

Global $JsonName = ""
Global $JsonDir = ""
Global $JsonFile = ""

Global $ConfName = ""
Global $ConfDir = ""
Global $ConfFile = ""

Global $ZabbixDataFile = ""
Global $vZabbix_Conf = ""
Global $vZabbix_Sender_Exe = ""
Global $Zabbix_Items = ""

Global $ZabbixBasePath = "c:\zabbix_agent"

Global $DataErrors = ""

Global $JobsCount = 0
Global $BackupJobsCount = 0
Global $RepoCount = 0
Global $VmByJobCount = 0
Global $TapeFileCount = 0
Global $TapeVmCount = 0
Global $BackupSyncCount = 0
Global $EndpointCount = 0
Global $AgentPolicyCount = 0
Global $AgentBackupCount = 0
Global $VmByJobMonitoredCount = 0
Global $TapeFileMonitoredCount = 0
Global $TapeVmMonitoredCount = 0
Global $BackupSyncMonitoredCount = 0
Global $EndpointMonitoredCount = 0
Global $AgentPolicyMonitoredCount = 0
Global $AgentBackupMonitoredCount = 0

Global $Array_Disc = ""
Global $Array_Disc_Tmp = ""
Global $Comma = ""
Global $Array_Disc_Repo = ""
Global $Array_Disc_Repo_Tmp = ""
Global $Comma_Repo = ""
Global $Array_Disc_VM = ""
Global $Array_Disc_VM_Tmp = ""
Global $Comma_VM = ""
Global $Array_Disc_TapeF = ""
Global $Array_Disc_TapeF_Tmp = ""
Global $Comma_TapeF = ""
Global $Array_Disc_TapeV = ""
Global $Array_Disc_TapeV_Tmp = ""
Global $Comma_TapeV = ""
Global $Array_Disc_BackupSync = ""
Global $Array_Disc_BackupSync_Tmp = ""
Global $Comma_BackupSync = ""
Global $Array_Disc_Endpoint = ""
Global $Array_Disc_Endpoint_Tmp = ""
Global $Comma_Endpoint = ""
Global $Array_Disc_AgentPolicy = ""
Global $Array_Disc_AgentPolicy_Tmp = ""
Global $Comma_AgentPolicy = ""
Global $Array_Disc_AgentBackup = ""
Global $Array_Disc_AgentBackup_Tmp = ""
Global $Comma_AgentBackup = ""

Global $sConnectionString
Global $oConnection
Global $Debug = 0
Global $sDriver = ""
Global $sDatabase = ""
Global $sServer = ""
Global $sPort = ""
Global $sUID = ""
Global $sPWD = ""
Global $BlacklistPatterns = ""
Global $MSSQL_JobsView = "dbo.[JobsView]"
Global $MSSQL_ObjectsInJobsView = "dbo.[ObjectsInJobsView]"
Global $MSSQL_ObjectsView = "dbo.[Backup.Model.ObjectsView]"
Global $MSSQL_BackupRepositories = "dbo.[BackupRepositories]"
Global $MSSQL_BackupRepositoryContainer = "dbo.[BackupRepositoryContainer]"
Global $MSSQL_BackupRepositoryContainerRepos = "dbo.[BackupRepositoryContainer.Repositories]"
Global $MSSQL_JobSessions = "dbo.[Backup.Model.JobSessions]"
Global $MSSQL_BackupTaskSessions = "dbo.[Backup.Model.BackupTaskSessions]"
Global $MSSQL_TapeJobs = "dbo.[Tape.jobs]"
Global $MSSQL_SessionLog = "dbo.[SessionLog]"
Global Const $BackupConfigurationJobType = 100
Global Const $MonitoredBackupJobTypes = ",0,1,12003,"

#EndRegion Globals

#Region Check Parameters
	For $i = 1 To $CmdLine[0]
	    Switch $CmdLine[$i]
		Case StringInStr($CmdLine[$i],"--debug=") <> 0
			$Debug = GetParameter($CmdLine[$i])
		Case StringInStr($CmdLine[$i],"--driver=") <> 0
			$sDriver = GetParameter($CmdLine[$i])
        Case StringInStr($CmdLine[$i],"--database=") <> 0
			$sDatabase = GetParameter($CmdLine[$i])
		Case StringInStr($CmdLine[$i],"--server=") <> 0
			$sServer = GetParameter($CmdLine[$i])
		Case StringInStr($CmdLine[$i],"--port=") <> 0
			$sPort = GetParameter($CmdLine[$i])
		Case StringInStr($CmdLine[$i],"--user=") <> 0
			$sUID = GetParameter($CmdLine[$i])
		Case StringInStr($CmdLine[$i],"--password=") <> 0
			$sPWD = GetParameter($CmdLine[$i])
		Case StringInStr($CmdLine[$i],"--blacklist=") <> 0
			_SetupBlacklist(GetParameter($CmdLine[$i]))
	EndSwitch
Next

Func GetParameter($string)
    Local $result = StringRegExp($string, "=(.*)", 1)

    If @error Then
        Return ""
    Else
        Return $result[0]
    EndIf
EndFunc

Func _NormalizeForBlacklist($value)
	Local $clean = StringReplace($value, ",", "_")
	$clean = StringStripWS($clean, 3)
	Return StringLower($clean)
EndFunc

Func _BuildBlacklistPattern($value)
	Local $normalized = _NormalizeForBlacklist($value)
	If $normalized = "" Then Return ""
	Local $placeholder = "__ASTERISK__"
	Local $tmp = StringReplace($normalized, "*", $placeholder)
	Local $escaped = StringRegExpReplace($tmp, "([\[\]\^\$\.\|\?\+\(\)\{\}\\])", "\\$1")
	Local $final = StringReplace($escaped, $placeholder, ".*")
	Return "^" & $final & "$"
EndFunc

Func _SetupBlacklist($value)
	Local $raw = StringStripWS($value, 3)
	$raw = StringReplace($raw, '"', "")
	If $raw = "" Then Return
	Local $parts = StringSplit($raw, "|")
	Local $builder = ""
	For $i = 1 To $parts[0]
		Local $entry = StringStripWS($parts[$i], 3)
		If $entry <> "" Then
			$builder &= "|" & _BuildBlacklistPattern($entry)
		EndIf
	Next
	If $builder <> "" Then
		$BlacklistPatterns &= $builder & "|"
		If $Debug > 0 Then
			_logmsg($LogFile,"Blacklist patterns: " & $BlacklistPatterns,true,true)
		EndIf
	EndIf
EndFunc

Func _IsBlacklistedName($name)
	If $BlacklistPatterns = "" Then Return False
	Local $target = _NormalizeForBlacklist($name)
	If $target = "" Then Return False
	Local $patterns = StringSplit(StringTrimLeft($BlacklistPatterns, 1), "|")
	For $i = 1 To $patterns[0]
		If $patterns[$i] = "" Then ContinueLoop
		Local $regex = $patterns[$i]
		If StringRegExp($target, $regex) Then
			If $Debug > 0 Then
				_logmsg($LogFile,"Blacklist match '" & $regex & "' for " & $name,true,true)
			EndIf
			Return True
		EndIf
	Next
	Return False
EndFunc

Func _IsMonitoredBackupJobType($jobType)
	Local $normalized_type = "," & StringStripWS(String($jobType), 8) & ","
	Return StringInStr($MonitoredBackupJobTypes, $normalized_type) > 0
EndFunc

Func _ToBool($value)
	If $value = Null Or $value = "" Then Return False
	Return _IsTrueValue($value)
EndFunc

Func _IsJobMonitorEnabled($schedule_enabled, $is_deleted, $run_manually)
	Local $schedule_ok = _IsTrueValue($schedule_enabled)
	Local $not_deleted = Not _IsTrueValue($is_deleted)
	Local $not_manual = Not _IsTrueValue($run_manually)
	Return ($schedule_ok And $not_deleted And $not_manual) ? 1 : 0
EndFunc

Func _ApplyBlacklistMonitorEnabled($monitorEnabled, $name1, $name2 = "")
	If _IsBlacklistedName($name1) Then Return -1
	If $name2 <> "" And _IsBlacklistedName($name2) Then Return -1
	Return $monitorEnabled
EndFunc

Func _IsEmptyGuid($value)
	If $value = Null Then Return True
	Local $normalized = StringStripWS(String($value), 8)
	Return $normalized = "" Or $normalized = "00000000-0000-0000-0000-000000000000"
EndFunc

Func _IsAgentPolicyJob($job_type, $parent_job_id = "")
	If $job_type = 12000 Or $job_type = 12002 Or $job_type = 12003 Then Return True
	If $job_type = 4000 And Not _IsEmptyGuid($parent_job_id) Then Return True
	Return False
EndFunc

Func _ResolveAgentMetricPrefix($job_type, $parent_job_id = "")
	If _IsAgentPolicyJob($job_type, $parent_job_id) Then
		Return "backup.veeam.customchecks.agent.policy"
	EndIf
	Return "backup.veeam.customchecks.agent.backup"
EndFunc

if ($sDriver = "" or $sDatabase = "" or $sServer = "") then
	ConsoleWrite(@CRLF & "Error: Missing parameters")

	$msg_usage = "Usage:" & @CRLF & _
	"--debug=[0/1/2] [Default 0]: (1) Display some infos, (2) Display SQL queries" & @CRLF & _
	"--driver=[string]: ODBC Driver name. ['SQL Server' for MS SQL. 'PostgreSQL ANSI' or 'PostgreSQL Unicode' for PostgreSQL'" & @CRLF & _
	"--server=[string]: Instance and server [Ex. MS SQL: localhost\VEEAMSQL PostgreSQL: localhost]" & @CRLF & _
	"--port=[string]: Port if required. For PostgreSQL is 5432. Not needed usually for MS SQL" & @CRLF & _
	"--database=[string]: Database Name" & @CRLF & _
	"--user=[string]: Username if needed. Per PostgreSQL the default is 'postgres'" & @CRLF & _
	"--password[string]: Password if needed"

	ConsoleWrite(@CRLF & @CRLF & $msg_usage)

	exit
EndIf
#EndRegion Check Parameters

#Region Files Handling
; Log
$LogName = @ScriptName & ".log"
$LogDir = $ZabbixBasePath & "\log"
$LogFile = $LogDir & "\" & $LogName

if Not FileExists($LogDir) Then
	DirCreate($LogDir)
endif

if FileExists($LogFile) then
	FileDelete($LogFile)
endif

FileDelete($LogFile)

; Json
$JsonName = @ScriptName & ".json"
$JsonDir = $ZabbixBasePath & "\data_apps"
$JsonFile = $JSONDIR & "\" & $JsonName

if Not FileExists($JSONDIR) Then
	DirCreate($JSONDIR)
endif

FileDelete($JsonFile)

; Data file for zabbix
$ZabbixDataFile = $JSONDIR & "\" & @ScriptName & ".data"

if FileExists($ZabbixDataFile) then
	FileDelete($ZabbixDataFile)
endif

; Conf File
$ConfName = @ScriptName & ".conf"

$ConfDir = $ZabbixBasePath & "\data_apps"
$ConfFile = $ConfDir & "\" & $ConfName

if Not FileExists($ConfDir) Then
	DirCreate($ConfDir)
endif
#EndRegion Files Handiling

#Region Errors Handler
Local $oErrorHandler = ObjEvent("AutoIt.Error", "_ErrFunc")
Func _ErrFunc($oError)
	local $errmsg = @CRLF & "err.number is: " & @TAB & $oError.number & @CRLF & _
            "err.windescription:" & @TAB & $oError.windescription & @CRLF & _
            "err.description is: " & @TAB & $oError.description & @CRLF & _
            "err.source is: " & @TAB & $oError.source & @CRLF & _
            "err.helpfile is: " & @TAB & $oError.helpfile & @CRLF & _
            "err.helpcontext is: " & @TAB & $oError.helpcontext & @CRLF & _
            "err.lastdllerror is: " & @TAB & $oError.lastdllerror & @CRLF & _
            "err.scriptline is: " & @TAB & $oError.scriptline & @CRLF & _
            "err.retcode is: " & @TAB & $oError.retcode & @CRLF & @CRLF
	_logmsg($LogFile,$errmsg,true,true)
EndFunc   ;==>_ErrFunc

; Internal ADO.au3 UDF COMError Handler
_ADO_ComErrorHandler_UserFunction(_ADO_COMErrorHandler_Function)
#EndRegion Errors Handler

#Region Configuration
; Default conf
$vDefaultConf[1] = "Veeam Custom Checks Configuration"
$ConfString = "# Zabbix_Sender.exe Position"
$ConfString &= "|" & "zabbix_sender_Var=c:\zabbix_agent\zabbix_sender.exe"
$ConfString &= "|" & ""
$ConfString &= "|" & "# Zabbix Configuration File"
$ConfString &= "|" & "zabbix_conf_Var=c:\zabbix_agent\data_apps\zabbix_agentd.win.conf"

_ArrayAdd($vDefaultConf,$ConfString)
$vDefaultConf_Size = Ubound($vDefaultConf) - 1

; Load Conf
if FileExists($ConfFile) = false then
	Local $hFileOpen = FileOpen($ConfFile, $FO_APPEND)
	If $hFileOpen = -1 Then
		MsgBox($MB_SYSTEMMODAL, "", "Configuration file cannot be opened", 10)
	EndIf

	for $i = 1 to $vDefaultConf_Size step 1
		FileWriteLine($hFileOpen,$vDefaultConf[$i])
	Next

	FileClose($hFileOpen)

	_logmsg($LogFile,"New configuration file created",true,true)
endif

For $i = 1 to _FileCountLines($ConfFile)
	$line = FileReadLine($ConfFile, $i)
	;msgbox (0,'','the line ' & $i & ' is ' & $line)
	$vResult = StringSplit($line,"=")
	;msgbox (0,'','the line ' & $i & ' is ' & $vResult)
    if $vResult[0] > 1 then
		Select
			Case StringInStr($vResult[1],"zabbix_conf_Var")
				$vZabbix_Conf = $vResult[2]
			Case StringInStr($vResult[1],"zabbix_sender_Var")
				$vZabbix_Sender_Exe = $vResult[2]
		endselect
	endif
Next
#EndRegion Configuration

#Region Main Function
_VeeamDataSearch()

Func _VeeamDataSearch()

	Local $result, $Result_Discovery, $Result_Data

	$result = _SQLConnection()
	if $result <> 0 then
		_logmsg($LogFile,"Connection Error",true,true)
		exit
	EndIf
	_InitMssqlObjectMap()

	; SQL MS SQL

	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 then
		$sqldiscovery = "WITH HostsConcatenated AS (" & @CRLF & _
				"    SELECT" & @CRLF & _
				"        ij.job_id," & @CRLF & _
				"        STUFF((" & @CRLF & _
				"            SELECT '|' + o.object_name + ',' + COALESCE(o.viobject_type, 'NoType')" & @CRLF & _
				"            FROM " & $MSSQL_ObjectsInJobsView & " ij2" & @CRLF & _
				"            JOIN " & $MSSQL_ObjectsView & " o ON o.id = ij2.object_id" & @CRLF & _
				"            WHERE ij2.job_id = ij.job_id" & @CRLF & _
				"            FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 1, '') AS hosts" & @CRLF & _
				"    FROM " & $MSSQL_ObjectsInJobsView & " ij" & @CRLF & _
				"    GROUP BY ij.job_id" & @CRLF & _
				")," & @CRLF & _
				"FirstSelection AS (" & @CRLF & _
				"    SELECT" & @CRLF & _
				"        CAST(j.id AS VARCHAR(255)) AS job_id," & @CRLF & _
				"        CAST(COALESCE(pj.name, j.name) AS VARCHAR(255)) AS job_name," & @CRLF & _
				"        CAST(j.repository_id AS VARCHAR(255)) AS repository_id," & @CRLF & _
				"        CAST(r.name AS VARCHAR(255)) AS repository_name," & @CRLF & _
				"        CAST(j.type AS VARCHAR(255)) AS job_type," & @CRLF & _
				"        CAST(j.is_deleted AS VARCHAR(255)) AS is_job_deleted," & @CRLF & _
				"        CAST(j.latest_result AS VARCHAR(255)) AS latest_job_result," & @CRLF & _
				"        CAST(j.schedule_enabled AS VARCHAR(255)) AS is_schedule_enabled," & @CRLF & _
				"        CAST(j.options.value('(//JobOptionsRoot/RunManually/text())[1]', 'VARCHAR(MAX)') AS VARCHAR(255)) AS job_options_runmanually," & @CRLF & _
				"        CAST(j.parent_job_id AS VARCHAR(255)) AS parent_job_id," & @CRLF & _
				"        CAST(j.parent_schedule_id AS VARCHAR(255)) AS parent_schedule_id," & @CRLF & _
				"        CAST(hc.hosts AS VARCHAR(255)) AS backup_hosts" & @CRLF & _
				"    FROM" & @CRLF & _
				"        " & $MSSQL_JobsView & " j" & @CRLF & _
				"    LEFT JOIN" & @CRLF & _
				"        HostsConcatenated hc ON j.id = hc.job_id" & @CRLF & _
				"    LEFT JOIN" & @CRLF & _
				"        " & $MSSQL_JobsView & " pj ON j.parent_job_id = pj.id" & @CRLF & _
				"    LEFT JOIN" & @CRLF & _
				"        " & $MSSQL_BackupRepositories & " r ON j.repository_id = r.id" & @CRLF & _
				")" & @CRLF & _
				"SELECT *" & @CRLF & _
				"FROM FirstSelection" & @CRLF & _
				"WHERE " & @CRLF & _
				"    job_id NOT IN (SELECT DISTINCT parent_job_id FROM FirstSelection WHERE parent_job_id IS NOT NULL);"

		$sql = "WITH HostsConcatenated AS (" & @CRLF & _
				"    SELECT " & @CRLF & _
				"        ij.job_id, " & @CRLF & _
				"        STUFF((" & @CRLF & _
				"            SELECT '|' + o.object_name + ',' + COALESCE(o.viobject_type, 'NoType') " & @CRLF & _
				"            FROM " & $MSSQL_ObjectsInJobsView & " ij2 " & @CRLF & _
				"            JOIN " & $MSSQL_ObjectsView & " o ON o.id = ij2.object_id " & @CRLF & _
				"            WHERE ij2.job_id = ij.job_id " & @CRLF & _
				"            FOR XML PATH('')" & @CRLF & _
				"        ), 1, 1, '') AS hosts " & @CRLF & _
				"    FROM " & $MSSQL_ObjectsInJobsView & " ij " & @CRLF & _
				"    GROUP BY ij.job_id " & @CRLF & _
				"), " & @CRLF & _
				"LatestBackupState AS (" & @CRLF & _
				"    SELECT " & @CRLF & _
				"        bs.job_id, " & @CRLF & _
				"        bs.job_type, " & @CRLF & _
				"        bs.state, " & @CRLF & _
				"        bs.result, " & @CRLF & _
				"        bs.reason, " & @CRLF & _
				"        bs.creation_time, " & @CRLF & _
				"        bs.end_time, " & @CRLF & _
				"        ROW_NUMBER() OVER (PARTITION BY bs.job_id ORDER BY bs.creation_time DESC) AS rn " & @CRLF & _
				"    FROM " & $MSSQL_JobSessions & " bs " & @CRLF & _
				"), " & @CRLF & _
				"LatestBackupTaskSession AS (" & @CRLF & _
				"    SELECT " & @CRLF & _
				"        bts.session_id, " & @CRLF & _
				"        bts.status, " & @CRLF & _
				"        bts.reason, " & @CRLF & _
				"        bts.creation_time, " & @CRLF & _
				"        js.job_id, " & @CRLF & _
				"        ROW_NUMBER() OVER (PARTITION BY js.job_id ORDER BY bts.creation_time DESC) AS rn " & @CRLF & _
				"    FROM " & $MSSQL_BackupTaskSessions & " bts " & @CRLF & _
				"    JOIN " & $MSSQL_JobSessions & " js ON bts.session_id = js.id " & @CRLF & _
				") " & @CRLF & _
				"SELECT " & @CRLF & _
				"	CAST(j.id AS VARCHAR(255)) AS job_id, " & @CRLF  & _
				"	CAST(COALESCE(pj.name, j.name) AS VARCHAR(255)) AS job_name, " & @CRLF  & _
				"	CAST(j.repository_id AS VARCHAR(255)) AS repository_id, " & @CRLF  & _
				"	CAST(r.name AS VARCHAR(255)) AS repository_name, " & @CRLF  & _
				"	CAST(j.type AS VARCHAR(255)) AS job_type, " & @CRLF  & _
				"	CAST(j.is_deleted AS VARCHAR(255)) AS is_job_deleted, " & @CRLF  & _
				"	CAST(j.latest_result AS VARCHAR(255)) AS latest_job_result, " & @CRLF  & _
				"	CAST(j.schedule_enabled AS VARCHAR(255)) AS is_schedule_enabled, " & @CRLF  & _
				"	CAST(j.parent_job_id AS VARCHAR(255)) AS parent_job_id, " & @CRLF  & _
				"	CAST(j.parent_schedule_id AS VARCHAR(255)) AS parent_schedule_id, " & @CRLF  & _
				"	CAST(hc.hosts AS VARCHAR(255)) AS backup_hosts, " & @CRLF  & _
				"	CAST(lbs.job_type AS VARCHAR(255)) AS backup_job_type, " & @CRLF  & _
				"	CAST(lbs.state AS VARCHAR(255)) AS backup_state, " & @CRLF  & _
				"	CAST(lbs.result AS VARCHAR(255)) AS backup_result, " & @CRLF  & _
				"	CAST(lbs.reason AS VARCHAR(255)) AS backup_reason, " & @CRLF  & _
				"	CAST(lbs.creation_time AS VARCHAR(255)) AS backup_creation_time, " & @CRLF  & _
				"	CAST(lbs.end_time AS VARCHAR(255)) AS backup_end_time, " & @CRLF  & _
				"	CAST(bts.status AS VARCHAR(255)) AS backup_task_status, " & @CRLF  & _
				"	CAST(bts.reason AS VARCHAR(255)) AS backup_task_reason, " & @CRLF  & _
				"	CAST(bts.session_id AS VARCHAR(255)) AS backup_task_session_id, " & @CRLF  & _
				"	CAST(j.schedule.value('(//OptionsScheduleAfterJob/IsEnabled/text())[1]', 'VARCHAR(MAX)') AS VARCHAR(255)) AS job_schedule_afterjob_enabled, " & @CRLF  & _
				"	CAST(j.schedule.value('(//OptionsDaily/Enabled/text())[1]', 'VARCHAR(MAX)') AS VARCHAR(255)) AS job_schedule_daily_enabled, " & @CRLF  & _
				"	CAST(j.schedule.value('(//OptionsDaily/Kind/text())[1]', 'VARCHAR(MAX)') AS VARCHAR(255)) AS job_schedule_daily_kind, " & @CRLF  & _
				"	CAST(STUFF(( " & @CRLF  & _
				"		SELECT ', ' + x.EMonth.value('.', 'VARCHAR(MAX)') " & @CRLF  & _
				"		FROM j.schedule.nodes('(//OptionsDaily/Days/DayOfWeek)') AS x(EMonth) " & @CRLF  & _
				"		FOR XML PATH('') " & @CRLF  & _
				"	), 1, 2, '') AS VARCHAR(255)) AS job_schedule_daily_days, " & @CRLF  & _
				"	CAST(j.schedule.value('(//OptionsPeriodically/Enabled/text())[1]', 'VARCHAR(MAX)') AS VARCHAR(255)) AS job_schedule_periodically_enabled, " & @CRLF  & _
				"	CAST(j.schedule.value('(//OptionsMonthly/Enabled/text())[1]', 'VARCHAR(MAX)') AS VARCHAR(255)) AS job_schedule_monthly_enabled, " & @CRLF  & _
				"	CAST(STUFF(( " & @CRLF  & _
				"		SELECT ', ' + x.EMonth.value('.', 'VARCHAR(MAX)') " & @CRLF  & _
				"		FROM j.schedule.nodes('(//OptionsMonthly/Months/EMonth)') AS x(EMonth) " & @CRLF  & _
				"		FOR XML PATH('') " & @CRLF  & _
				"	), 1, 2, '') AS VARCHAR(255)) AS job_schedule_monthly_months " & @CRLF  & _
				"FROM " & @CRLF & _
				"    " & $MSSQL_JobsView & " j " & @CRLF & _
				"LEFT JOIN " & @CRLF & _
				"    HostsConcatenated hc ON j.id = hc.job_id " & @CRLF & _
				"LEFT JOIN " & @CRLF & _
				"    " & $MSSQL_JobsView & " pj ON j.parent_job_id = pj.id " & @CRLF & _
				"LEFT JOIN " & @CRLF & _
				"    " & $MSSQL_BackupRepositories & " r ON j.repository_id = r.id " & @CRLF & _
				"LEFT JOIN " & @CRLF & _
				"    LatestBackupState lbs ON j.id = lbs.job_id AND lbs.rn = 1 " & @CRLF & _
				"LEFT JOIN " & @CRLF & _
				"    LatestBackupTaskSession bts ON j.id = bts.job_id AND bts.rn = 1 " & @CRLF & _
				"WHERE " & @CRLF & _
				"    bts.session_id IS NOT NULL;"
	Endif

	; SQL PostgreSQL
	If StringInStr($sDriver,"PostgreSQL") <> 0 then
		$sqldiscovery = "WITH HostsConcatenated AS (" & @CRLF & _
						"    SELECT" & @CRLF & _
						"        ij.job_id," & @CRLF & _
						"        STRING_AGG(o.object_name || ',' || COALESCE(o.viobject_type, 'NoType'), '|') AS hosts" & @CRLF & _
						"    FROM public." & chr(34) & "objectsinjobsview" & chr(34) & " ij " & @CRLF  & _
						"    JOIN public." & chr(34) & "objectsview" & chr(34) & " o ON o.id = ij.object_id " & @CRLF  & _
						"    GROUP BY ij.job_id" & @CRLF & _
						")," & @CRLF & _
						"FirstSelection AS (" & @CRLF & _
						"    SELECT" & @CRLF & _
						"        j.id::TEXT AS job_id," & @CRLF & _
						"        j.name::TEXT AS job_name," & @CRLF & _
						"        j.repository_id::TEXT AS repository_id," & @CRLF & _
						"        r.name::TEXT AS repository_name," & @CRLF & _
						"        j.type::TEXT AS job_type," & @CRLF & _
						"        j.is_deleted::TEXT AS is_job_deleted," & @CRLF & _
						"        j.latest_result::TEXT AS latest_job_result," & @CRLF & _
						"        j.schedule_enabled::TEXT AS is_schedule_enabled," & @CRLF & _
						"        (xpath('//JobOptionsRoot/RunManually/text()', xmlparse(document j.options)))[1]::text AS job_options_runmanually," & @CRLF & _
						"        j.parent_job_id::TEXT AS parent_job_id," & @CRLF & _
						"        j.parent_schedule_id::TEXT AS parent_schedule_id," & @CRLF & _
						"        hc.hosts::TEXT AS backup_hosts" & @CRLF & _
						"    FROM public." & chr(34) & "jobsview" & chr(34) & " j " & @CRLF  & _
						"    LEFT JOIN HostsConcatenated hc ON j.id = hc.job_id" & @CRLF & _
						"    LEFT JOIN public." & chr(34) & "backuprepositories" & chr(34) & " r ON j.repository_id = r.id " & @CRLF  & _
						")" & @CRLF & _
						"SELECT *" & @CRLF & _
						"FROM FirstSelection" & @CRLF & _
						"WHERE " & @CRLF & _
						"    job_id NOT IN (SELECT DISTINCT parent_job_id FROM FirstSelection WHERE parent_job_id IS NOT NULL);"

		$sql = "WITH HostsConcatenated AS (" & @CRLF  & _
				"    SELECT " & @CRLF  & _
				"        ij.job_id, " & @CRLF  & _
				"        string_agg(o.object_name || ',' || COALESCE(o.viobject_type, 'NoType'), '|') AS hosts " & @CRLF  & _
				"    FROM " & @CRLF  & _
				"        public." & chr(34) & "objectsinjobsview" & chr(34) & " ij " & @CRLF  & _
				"    JOIN " & @CRLF  & _
				"        public." & chr(34) & "objectsview" & chr(34) & " o ON o.id = ij.object_id " & @CRLF  & _
				"    GROUP BY " & @CRLF  & _
				"        ij.job_id " & @CRLF  & _
				"), " & @CRLF  & _
				"LatestBackupState AS (" & @CRLF  & _
				"    SELECT " & @CRLF  & _
				"        bs.job_id, " & @CRLF  & _
				"        bs.job_type, " & @CRLF  & _
				"        bs.state, " & @CRLF  & _
				"        bs.result, " & @CRLF  & _
				"        bs.reason, " & @CRLF  & _
				"        bs.creation_time, " & @CRLF  & _
				"        bs.end_time, " & @CRLF  & _
				"        ROW_NUMBER() OVER (PARTITION BY bs.job_id ORDER BY bs.creation_time DESC) AS rn " & @CRLF  & _
				"    FROM " & @CRLF  & _
				"        public." & chr(34) & "backup.model.jobsessions" & chr(34) & " bs " & @CRLF  & _
				"), " & @CRLF  & _
				"LatestBackupTaskSession AS (" & @CRLF  & _
				"    SELECT " & @CRLF  & _
				"        bts.session_id, " & @CRLF  & _
				"        bts.status, " & @CRLF  & _
				"        bts.reason, " & @CRLF  & _
				"        bts.creation_time, " & @CRLF  & _
				"        js.job_id, " & @CRLF  & _
				"        ROW_NUMBER() OVER (PARTITION BY js.job_id ORDER BY bts.creation_time DESC) AS rn " & @CRLF  & _
				"    FROM " & @CRLF  & _
				"        public." & chr(34) & "backup.model.backuptasksessions" & chr(34) & " bts " & @CRLF  & _
				"    JOIN " & @CRLF  & _
				"        public." & chr(34) & "backup.model.jobsessions" & chr(34) & " js ON bts.session_id = js.id " & @CRLF  & _
				") " & @CRLF  & _
				"SELECT " & @CRLF  & _
				"    j.id AS job_id, " & @CRLF  & _
				"    j.name AS job_name, " & @CRLF  & _
				"    j.repository_id, " & @CRLF  & _
				"    r.name AS repository_name, " & @CRLF  & _
				"    j.type AS job_type, " & @CRLF  & _
				"    j.is_deleted AS is_job_deleted, " & @CRLF  & _
				"    j.latest_result AS latest_job_result, " & @CRLF  & _
				"    j.schedule_enabled AS is_schedule_enabled, " & @CRLF  & _
				"    j.parent_job_id, " & @CRLF  & _
				"    j.parent_schedule_id, " & @CRLF  & _
				"    hc.hosts AS backup_hosts, " & @CRLF  & _
				"    lbs.job_type AS backup_job_type, " & @CRLF  & _
				"    lbs.state AS backup_state, " & @CRLF  & _
				"    lbs.result AS backup_result, " & @CRLF  & _
				"    lbs.reason AS backup_reason, " & @CRLF  & _
				"    lbs.creation_time AS backup_creation_time, " & @CRLF  & _
				"    lbs.end_time AS backup_end_time, " & @CRLF  & _
				"    bts.status AS backup_task_status, " & @CRLF  & _
				"    bts.reason AS backup_task_reason, " & @CRLF  & _
				"    bts.session_id AS backup_task_session_id, " & @CRLF  & _
				"    (xpath('//OptionsScheduleAfterJob/IsEnabled/text()', xmlparse(document j.schedule)))[1]::text AS job_schedule_afterjob_enabled, " & @CRLF  & _
				"    (xpath('//OptionsDaily/Enabled/text()', xmlparse(document j.schedule)))[1]::text AS job_schedule_daily_enabled, " & @CRLF  & _
				"    (xpath('//OptionsDaily/Kind/text()', xmlparse(document j.schedule)))[1]::text AS job_schedule_daily_kind, " & @CRLF  & _
				"    array_to_string(" & @CRLF  & _
				"        array(" & @CRLF  & _
				"            SELECT unnest(xpath('//OptionsDaily/Days/DayOfWeek/text()', xmlparse(document j.schedule)))" & @CRLF  & _
				"        ), ', ' " & @CRLF  & _
				"    ) AS job_schedule_daily_days, " & @CRLF  & _
				"    (xpath('//OptionsMonthly/Enabled/text()', xmlparse(document j.schedule)))[1]::text AS job_schedule_monthly_enabled, " & @CRLF  & _
				"    (xpath('//OptionsPeriodically/Enabled/text()', xmlparse(document j.schedule)))[1]::text AS job_schedule_periodically_enabled, " & @CRLF  & _
				"    array_to_string(" & @CRLF  & _
				"        array(" & @CRLF  & _
				"            SELECT unnest(xpath('//OptionsMonthly/Months/EMonth/text()', xmlparse(document j.schedule)))" & @CRLF  & _
				"        ), ', ' " & @CRLF  & _
				"    ) AS job_schedule_monthly_months " & @CRLF  & _
				"FROM " & @CRLF  & _
				"    public." & chr(34) & "jobsview" & chr(34) & " j " & @CRLF  & _
				"LEFT JOIN " & @CRLF  & _
				"    HostsConcatenated hc ON j.id = hc.job_id " & @CRLF  & _
				"LEFT JOIN " & @CRLF  & _
				"    public." & chr(34) & "backuprepositories" & chr(34) & " r ON j.repository_id = r.id " & @CRLF  & _
				"LEFT JOIN " & @CRLF  & _
				"    LatestBackupState lbs ON j.id = lbs.job_id AND lbs.rn = 1 " & @CRLF  & _
				"LEFT JOIN " & @CRLF  & _
				"    LatestBackupTaskSession bts ON j.id = bts.job_id AND bts.rn = 1 " & @CRLF  & _
				"WHERE " & @CRLF  & _
				"    bts.session_id IS NOT NULL;"

	EndIf

	$Array_Disc = "{" & chr(34) & "data" & chr(34) & ":["
	$Array_Disc_Tmp = ""
	$Comma = ""
	$Array_Disc_Repo = "{" & chr(34) & "data" & chr(34) & ":["
	$Array_Disc_Repo_Tmp = ""
	$Comma_Repo = ""
	$Array_Disc_VM = "{" & chr(34) & "data" & chr(34) & ":["
	$Array_Disc_VM_Tmp = ""
	$Comma_VM = ""
	$Array_Disc_TapeF = "{" & chr(34) & "data" & chr(34) & ":["
	$Array_Disc_TapeF_Tmp = ""
	$Comma_TapeF = ""
	$Array_Disc_TapeV = "{" & chr(34) & "data" & chr(34) & ":["
	$Array_Disc_TapeV_Tmp = ""
	$Comma_TapeV = ""
	$Array_Disc_BackupSync = "{" & chr(34) & "data" & chr(34) & ":["
	$Array_Disc_BackupSync_Tmp = ""
	$Comma_BackupSync = ""
	$Array_Disc_Endpoint = "{" & chr(34) & "data" & chr(34) & ":["
	$Array_Disc_Endpoint_Tmp = ""
	$Comma_Endpoint = ""
	$Array_Disc_AgentPolicy = "{" & chr(34) & "data" & chr(34) & ":["
	$Array_Disc_AgentPolicy_Tmp = ""
	$Comma_AgentPolicy = ""
	$Array_Disc_AgentBackup = "{" & chr(34) & "data" & chr(34) & ":["
	$Array_Disc_AgentBackup_Tmp = ""
	$Comma_AgentBackup = ""

	; Get jobs for discovery
	$Result_Discovery = _SqlRetrieveData($sqldiscovery)
	If IsObj($Result_Discovery) then
		DiscoveryData($Result_Discovery,$sDriver)
	else
		_logmsg($LogFile,"Error Main SQL: " & $Result_Discovery,true,true)
	Endif

	; Repositories discovery
	Local $sql_repo_discovery = _SqlRepoDiscovery($sDriver)
	Local $Result_RepoDiscovery = _SqlRetrieveData($sql_repo_discovery)
	If IsObj($Result_RepoDiscovery) Then
		DiscoveryRepoData($Result_RepoDiscovery)
	Else
		_logmsg($LogFile,"Error Repo Discovery SQL: " & $Result_RepoDiscovery,true,true)
	EndIf

	; VM-by-Job discovery
	Local $sql_vm_discovery = _SqlVmByJobDiscovery($sDriver)
	Local $Result_VmDiscovery = _SqlRetrieveData($sql_vm_discovery)
	If IsObj($Result_VmDiscovery) Then
		DiscoveryVmByJobData($Result_VmDiscovery)
	Else
		_logmsg($LogFile,"Error VM Discovery SQL: " & $Result_VmDiscovery,true,true)
	EndIf

	; Tape discovery (file-to-tape and VM-tape)
	Local $sql_tape_file = _SqlTapeDiscovery($sDriver, 24)
	Local $Result_TapeFile = _SqlRetrieveData($sql_tape_file)
	If IsObj($Result_TapeFile) Then
		DiscoveryTapeData($Result_TapeFile, 24)
	Else
		_logmsg($LogFile,"Error Tape File Discovery SQL: " & $Result_TapeFile,true,true)
	EndIf

	Local $sql_tape_vm = _SqlTapeDiscovery($sDriver, 28)
	Local $Result_TapeVm = _SqlRetrieveData($sql_tape_vm)
	If IsObj($Result_TapeVm) Then
		DiscoveryTapeData($Result_TapeVm, 28)
	Else
		_logmsg($LogFile,"Error Tape VM Discovery SQL: " & $Result_TapeVm,true,true)
	EndIf

	; BackupSync discovery
	Local $sql_backupsync = _SqlBackupSyncDiscovery($sDriver)
	Local $Result_BackupSync = _SqlRetrieveData($sql_backupsync)
	If IsObj($Result_BackupSync) Then
		DiscoveryBackupSyncData($Result_BackupSync,$sDriver)
	Else
		_logmsg($LogFile,"Error BackupSync Discovery SQL: " & $Result_BackupSync,true,true)
	EndIf

	; Endpoint/Agent discovery
	Local $sql_endpoint = _SqlEndpointDiscovery($sDriver)
	Local $Result_Endpoint = _SqlRetrieveData($sql_endpoint)
	If IsObj($Result_Endpoint) Then
		DiscoveryEndpointData($Result_Endpoint)
	Else
		_logmsg($LogFile,"Error Endpoint Discovery SQL: " & $Result_Endpoint,true,true)
	EndIf

	Local $sql_agent = _SqlAgentDiscovery($sDriver)
	Local $Result_Agent = _SqlRetrieveData($sql_agent)
	If IsObj($Result_Agent) Then
		DiscoveryAgentData($Result_Agent)
	Else
		_logmsg($LogFile,"Error Agent Discovery SQL: " & $Result_Agent,true,true)
	EndIf

	; Backup Configuration Job (PostgreSQL only)
	If StringInStr($sDriver,"PostgreSQL") <> 0 Then
		Local $sql_backup_config = _SqlBackupConfigurationJobPostgres()
		Local $Result_BackupConfig = _SqlRetrieveData($sql_backup_config)
		If IsObj($Result_BackupConfig) Then
			BackupConfigurationJobData($Result_BackupConfig)
		Else
			_logmsg($LogFile,"Error Backup Configuration Job SQL: " & $Result_BackupConfig,true,true)
		EndIf
	EndIf

	; Repo space metrics
	_logmsg($LogFile,"=== METRICS ===",true,true)
	Local $sql_repo_metrics = _SqlRepoMetrics($sDriver)
	Local $Result_RepoMetrics = _SqlRetrieveData($sql_repo_metrics)
	If IsObj($Result_RepoMetrics) Then
		RepoData($Result_RepoMetrics)
	Else
		_logmsg($LogFile,"Error Repo Metrics SQL: " & $Result_RepoMetrics,true,true)
	EndIf

	_logmsg($LogFile,"=== STATUS ===",true,true)

	; Core backupjob status
	$Result_Data = _SqlRetrieveData($sql)
	If IsObj($Result_Data) then
		BackupData($Result_Data,$sDriver)
	else
		_logmsg($LogFile,"Error Main SQL: " & $Result_Data,true,true)
	Endif

	; VM task status with retry awareness
	Local $sql_vm_tasks = _SqlVmTasksWithRetry($sDriver)
	Local $Result_VmTasks = _SqlRetrieveData($sql_vm_tasks)
	If _RecordsetHasRows($Result_VmTasks) Then
		VmTaskData($Result_VmTasks)
	ElseIf StringInStr($sDriver,"SQL Server") <> 0 Then
		Local $sql_vm_tasks_fallback = _SqlVmTasksFallbackMssql($sDriver)
		Local $Result_VmTasksFallback = _SqlRetrieveData($sql_vm_tasks_fallback)
		If _RecordsetHasRows($Result_VmTasksFallback) Then
			_logmsg($LogFile,"VM Tasks fallback MSSQL: using simplified query.",true,true)
			VmTaskData($Result_VmTasksFallback)
		Else
			Local $Result_VmDiscoveryDefaults = _SqlRetrieveData(_SqlVmByJobDiscovery($sDriver))
			If _RecordsetHasRows($Result_VmDiscoveryDefaults) Then
				_logmsg($LogFile,"VM Tasks fallback MSSQL: no task rows, emitting defaults from discovery.",true,true)
				VmTaskDefaultData($Result_VmDiscoveryDefaults)
			Else
				_logmsg($LogFile,"No VM task/discovery rows available for VM status items.",true,true)
			EndIf
		EndIf
	Else
		Local $Result_VmDiscoveryDefaults = _SqlRetrieveData(_SqlVmByJobDiscovery($sDriver))
		If _RecordsetHasRows($Result_VmDiscoveryDefaults) Then
			_logmsg($LogFile,"VM Tasks fallback: no task rows, emitting defaults from discovery.",true,true)
			VmTaskDefaultData($Result_VmDiscoveryDefaults)
		Else
			_logmsg($LogFile,"No VM task/discovery rows available for VM status items.",true,true)
		EndIf
	EndIf

	; Tape status
	Local $sql_tape_status = _SqlTapeStatus($sDriver)
	Local $Result_TapeStatus = _SqlRetrieveData($sql_tape_status)
	If IsObj($Result_TapeStatus) Then
		TapeStatusData($Result_TapeStatus)
	Else
		_logmsg($LogFile,"Error Tape Status SQL: " & $Result_TapeStatus,true,true)
	EndIf

	; BackupSync status
	Local $sql_backupsync_status = _SqlBackupSyncStatus($sDriver)
	Local $Result_BackupSyncStatus = _SqlRetrieveData($sql_backupsync_status)
	If IsObj($Result_BackupSyncStatus) Then
		BackupSyncStatusData($Result_BackupSyncStatus)
	Else
		_logmsg($LogFile,"Error BackupSync Status SQL: " & $Result_BackupSyncStatus,true,true)
	EndIf

	; Endpoint/Agent status
	Local $sql_endpoint_status = _SqlEndpointStatus($sDriver)
	Local $Result_EndpointStatus = _SqlRetrieveData($sql_endpoint_status)
	If IsObj($Result_EndpointStatus) Then
		EndpointStatusData($Result_EndpointStatus)
	Else
		_logmsg($LogFile,"Error Endpoint Status SQL: " & $Result_EndpointStatus,true,true)
	EndIf

	$Array_Disc &= $Array_Disc_Tmp & "]}"
	$Array_Disc_Repo &= $Array_Disc_Repo_Tmp & "]}"
	$Array_Disc_VM &= $Array_Disc_VM_Tmp & "]}"
	$Array_Disc_TapeF &= $Array_Disc_TapeF_Tmp & "]}"
	$Array_Disc_TapeV &= $Array_Disc_TapeV_Tmp & "]}"
	$Array_Disc_BackupSync &= $Array_Disc_BackupSync_Tmp & "]}"
	$Array_Disc_Endpoint &= $Array_Disc_Endpoint_Tmp & "]}"
	$Array_Disc_AgentPolicy &= $Array_Disc_AgentPolicy_Tmp & "]}"
	$Array_Disc_AgentBackup &= $Array_Disc_AgentBackup_Tmp & "]}"

	; Add DataErrors to zabbix data
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.dataerrors",$DataErrors)

	; Send discovery data to Zabbix
	_logmsg($LogFile,"Zabbix - Discovery",false,true)
	Local $DiscoveryPayload = " - backup.veeam.customchecks.backupjob.discovery " & $array_disc & @CRLF & _
		" - backup.veeam.customchecks.repo.discovery " & $Array_Disc_Repo & @CRLF & _
		" - backup.veeam.customchecks.vm.discovery " & $Array_Disc_VM & @CRLF & _
		" - backup.veeam.customchecks.tape.file.discovery " & $Array_Disc_TapeF & @CRLF & _
		" - backup.veeam.customchecks.tape.vm.discovery " & $Array_Disc_TapeV & @CRLF & _
		" - backup.veeam.customchecks.backupsync.discovery " & $Array_Disc_BackupSync & @CRLF & _
		" - backup.veeam.customchecks.endpoint.discovery " & $Array_Disc_Endpoint & @CRLF & _
		" - backup.veeam.customchecks.agent.policy.discovery " & $Array_Disc_AgentPolicy & @CRLF & _
		" - backup.veeam.customchecks.agent.backup.discovery " & $Array_Disc_AgentBackup
	FileWrite($JsonFile, $DiscoveryPayload)
	$ZabbixSend = $vZabbix_Sender_Exe & " -vv -c " & $vZabbix_Conf & " -i " & $JsonFile
	RunWait($ZabbixSend,$ZabbixBasePath,@SW_HIDE)

	; Jobs Count
	_logmsg($LogFile,"Zabbix - Jobs Count",false,true)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.monitored.count",$JobsCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.count",$BackupJobsCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.repo.count",$RepoCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vmbyjob.count",$VmByJobCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.tape.file.count",$TapeFileCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.tape.vm.count",$TapeVmCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupsync.count",$BackupSyncCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.endpoint.count",$EndpointCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.agent.policy.count",$AgentPolicyCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.agent.backup.count",$AgentBackupCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vmbyjob.monitored.count",$VmByJobMonitoredCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.tape.file.monitored.count",$TapeFileMonitoredCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.tape.vm.monitored.count",$TapeVmMonitoredCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupsync.monitored.count",$BackupSyncMonitoredCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.endpoint.monitored.count",$EndpointMonitoredCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.agent.policy.monitored.count",$AgentPolicyMonitoredCount)
	$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.agent.backup.monitored.count",$AgentBackupMonitoredCount)

	; Send data to Zabbix
	_logmsg($LogFile,"Zabbix - Data",false,true)
	FileWrite($ZabbixDataFile, $Zabbix_Items)
	$ZabbixSend = $vZabbix_Sender_Exe & " -vv -c " & $vZabbix_Conf & " -i " & $ZabbixDataFile
	RunWait($ZabbixSend,$ZabbixBasePath,@SW_HIDE)

	; CleanUp
	$oRecordset = Null
	_ADO_Connection_Close($oConnection)
	$oConnection = Null

EndFunc
#EndRegion Main Function

#Region Functions
Func _InitMssqlObjectMap()
	If StringInStr($sDriver, "SQL Server") = 0 Then Return

	$MSSQL_JobsView = _MssqlResolveObject("JobsView", "V", "WmiServer.JobsView", "V", "BJobs", "U")
	$MSSQL_ObjectsInJobsView = _MssqlResolveObject("ObjectsInJobsView", "V", "WmiServer.FileSystemObjectsInJobsView", "V", "ObjectsInJobs", "U")
	$MSSQL_ObjectsView = _MssqlResolveObject("Backup.Model.ObjectsView", "V", "ObjectsView", "V")
	$MSSQL_BackupRepositories = _MssqlResolveObject("BackupRepositories", "U", "backuprepositories", "U")
	$MSSQL_BackupRepositoryContainer = _MssqlResolveObject("BackupRepositoryContainer", "U", "backuprepositorycontainer", "U")
	$MSSQL_BackupRepositoryContainerRepos = _MssqlResolveObject("BackupRepositoryContainer.Repositories", "U", "backuprepositorycontainer.repositories", "U")
	$MSSQL_JobSessions = _MssqlResolveObject("Backup.Model.JobSessions", "U", "backup.model.jobsessions", "U")
	$MSSQL_BackupTaskSessions = _MssqlResolveObject("Backup.Model.BackupTaskSessions", "U", "backup.model.backuptasksessions", "U")
	$MSSQL_TapeJobs = _MssqlResolveObject("Tape.jobs", "U", "tape.jobs", "U")
	$MSSQL_SessionLog = _MssqlResolveObject("SessionLog", "U", "sessionlog", "U")
EndFunc

Func _MssqlResolveObject($name1, $type1, $name2 = "", $type2 = "", $name3 = "", $type3 = "")
	If _MssqlObjectExists($name1, $type1) Then Return "dbo.[" & $name1 & "]"
	If $name2 <> "" And _MssqlObjectExists($name2, $type2) Then Return "dbo.[" & $name2 & "]"
	If $name3 <> "" And _MssqlObjectExists($name3, $type3) Then Return "dbo.[" & $name3 & "]"
	Return "dbo.[" & $name1 & "]"
EndFunc

Func _MssqlObjectExists($name, $objType)
	Local $sql = "SELECT CASE WHEN OBJECT_ID(N'[dbo].[" & $name & "]', '" & $objType & "') IS NULL THEN 0 ELSE 1 END AS obj_exists;"
	Local $rs = _SqlRetrieveData($sql)
	If Not IsObj($rs) Then Return False
	If $rs.EOF Then Return False
	Local $exists = $rs.Fields("obj_exists").Value
	Return ($exists = 1)
EndFunc

; Connection to SQL
Func _SQLConnection()
	Local $port_string = ""
	If $sPort <> "" then
		$port_string = 'PORT=' & $sPort & ';'
	EndIf

	$TrustedConn = ""

	if stringinstr($sDriver,"SQL Server") <> 0 then
		$TrustedConn= "Trusted_Connection=Yes"
	endif

	$sConnectionString = 'DRIVER={' & $sDriver & '};SERVER=' & $sServer & ';DATABASE=' & $sDatabase & ';UID=' & $sUID & ';PWD=' & $sPWD & ';' & $port_string & ";" & $TrustedConn

	$oConnection = _ADO_Connection_Create()
	_logmsg($LogFile,"Connection to " & $sDriver,false,true)

	If $Debug = 2 Then
		_logmsg($LogFile,"ConnectionString: " & $sConnectionString,true,true)
	EndIf

	_ADO_Connection_OpenConString($oConnection, $sConnectionString)

	If @error Then
		_logmsg($LogFile,"Connection Error: " & @error & " - " & @extended & " - " & $ADO_RET_FAILURE,true,true)
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	EndIf
EndFunc

Func _SqlRetrieveData($sql)

	if $sql = "" then
		return null
	endif

	Local $oRecordset = _ADO_Execute($oConnection, $sql)
	If @error Then
		Local $sErrorDetail = ""
		If IsObj($oConnection.Errors) Then
			For $oError In $oConnection.Errors
				$sErrorDetail &= "Descrizione: " & $oError.Description & @CRLF
				$sErrorDetail &= "Numero: " & $oError.Number & @CRLF
				$sErrorDetail &= "Origine: " & $oError.Source & @CRLF
			Next
		Else
			$sErrorDetail = "Nessun dettaglio disponibile"
		EndIf

		_logmsg($LogFile, "Retrieve data error: " & @error & " - " & @extended & @CRLF & $sErrorDetail, True, True)
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	EndIf

	If $Debug = 2 Then
		_logmsg($LogFile, "SQL: " & $sql, True, True)
	EndIf

	Return $oRecordset
EndFunc

Func _IsTrueValue($value)
	If IsBool($value) Then Return $value
	Local $s = StringLower(StringStripWS(String($value), 3))
	Return ($s = "1" Or $s = "true" Or $s = "t" Or $s = "yes")
EndFunc

Func _IsFalseValue($value)
	If IsBool($value) Then Return Not $value
	Local $s = StringLower(StringStripWS(String($value), 3))
	Return ($s = "0" Or $s = "false" Or $s = "f" Or $s = "no")
EndFunc

Func _NullToZero($value)
	If $value = Null Or $value = "" Then Return 0
	Return $value
EndFunc

Func _RecordsetHasRows($Recordset)
	If Not IsObj($Recordset) Then Return False
	Return Not ($Recordset.BOF And $Recordset.EOF)
EndFunc

Func _IsNumericValue($value)
	If IsNumber($value) Then Return True
	Local $s = StringStripWS(String($value), 3)
	If $s = "" Then Return False
	Return StringRegExp($s, "^-?\d+(\.\d+)?$")
EndFunc

Func _AppendReason($base, $extra)
	Local $b = StringStripWS(String($base), 3)
	Local $e = StringStripWS(String($extra), 3)
	If $e = "" Or $e = "\N" Then Return $b
	If $b = "" Or $b = "\N" Then Return $e
	Local $parts = StringSplit($b, ";")
	For $i = 1 To $parts[0]
		Local $p = StringStripWS($parts[$i], 3)
		If StringLower($p) = StringLower($e) Then
			Return $b
		EndIf
	Next
	Return $b & "; " & $e
EndFunc

Func _ComputeDurationMinutes($startDate, $endDate)
	If _DateIsValid($startDate) And _DateIsValid($endDate) Then
		If $endDate > $startDate Then
			Return _DateDiff('n', $startDate, $endDate)
		EndIf
	EndIf
	Return -1
EndFunc

Func _ComputeDateDiffDays($startDate)
	If _DateIsValid($startDate) Then
		Return _DateDiff('D', $startDate, _NowCalc())
	EndIf
	Return -1
EndFunc

; Compute next run time using the same schedule logic as backupjob
Func _ComputeNextRunValue($sDriver, $job_id, $backup_creation_time_date)
	If $backup_creation_time_date = -1 Then Return -1
	If $job_id = "" Or $job_id = Null Then Return -1

	Local $job_schedule_afterjob_enabled = ""
	Local $job_schedule_afterjob_name = ""
	Local $job_schedule_daily_enabled = ""
	Local $job_schedule_daily_kind = ""
	Local $job_schedule_daily_days = ""
	Local $job_schedule_monthly_enabled = ""
	Local $job_schedule_monthly_months = ""
	Local $job_schedule_periodically_enabled = ""
	Local $parent_schedule_id = ""
	Local $parent_job_id = ""

	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 Then
		$sql = "SELECT id, name, parent_job_id, parent_schedule_id," & @CRLF & _
				"    schedule.value('(//OptionsScheduleAfterJob/IsEnabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_afterjob_enabled," & @CRLF & _
				"    schedule.value('(//OptionsDaily/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_daily_enabled," & @CRLF & _
				"    schedule.value('(//OptionsDaily/Kind/text())[1]', 'VARCHAR(MAX)') AS job_schedule_daily_kind," & @CRLF & _
				"    STUFF((" & @CRLF & _
				"        SELECT ', ' + x.EMonth.value('.', 'VARCHAR(MAX)')" & @CRLF & _
				"        FROM schedule.nodes('(//OptionsDaily/Days/DayOfWeek)') AS x(EMonth)" & @CRLF & _
				"        FOR XML PATH('')" & @CRLF & _
				"    ), 1, 2, '') AS job_schedule_daily_days," & @CRLF & _
				"    schedule.value('(//OptionsPeriodically/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_periodically_enabled," & @CRLF & _
				"    schedule.value('(//OptionsMonthly/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_monthly_enabled," & @CRLF & _
				"    STUFF((" & @CRLF & _
				"        SELECT ', ' + x.EMonth.value('.', 'VARCHAR(MAX)')" & @CRLF & _
				"        FROM schedule.nodes('(//OptionsMonthly/Months/EMonth)') AS x(EMonth)" & @CRLF & _
				"        FOR XML PATH('')" & @CRLF & _
				"    ), 1, 2, '') AS job_schedule_monthly_months" & @CRLF & _
				"FROM " & $MSSQL_JobsView & " WHERE id = '" & $job_id & "';"
	Else
		$sql = "SELECT id, name, parent_job_id, parent_schedule_id," & @CRLF & _
				"    (xpath('//OptionsScheduleAfterJob/IsEnabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_afterjob_enabled," & @CRLF & _
				"    (xpath('//OptionsDaily/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_daily_enabled," & @CRLF & _
				"    (xpath('//OptionsDaily/Kind/text()', xmlparse(document schedule)))[1]::text AS job_schedule_daily_kind," & @CRLF & _
				"    array_to_string(array(SELECT unnest(xpath('//OptionsDaily/Days/DayOfWeek/text()', xmlparse(document schedule)))), ', ') AS job_schedule_daily_days," & @CRLF & _
				"    (xpath('//OptionsMonthly/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_monthly_enabled," & @CRLF & _
				"    (xpath('//OptionsPeriodically/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_periodically_enabled," & @CRLF & _
				"    array_to_string(array(SELECT unnest(xpath('//OptionsMonthly/Months/EMonth/text()', xmlparse(document schedule)))), ', ') AS job_schedule_monthly_months" & @CRLF & _
				"FROM public.jobsview WHERE id = '" & $job_id & "';"
	EndIf

	Local $oRecordset_Job = _SqlRetrieveData($sql)
	If IsObj($oRecordset_Job) Then
		While Not $oRecordset_Job.EOF
			$parent_job_id = $oRecordset_Job.Fields("parent_job_id").Value
			$parent_schedule_id = $oRecordset_Job.Fields("parent_schedule_id").Value
			$job_schedule_afterjob_enabled = $oRecordset_Job.Fields("job_schedule_afterjob_enabled").Value
			$job_schedule_afterjob_name = $oRecordset_Job.Fields("name").Value
			$job_schedule_daily_enabled = $oRecordset_Job.Fields("job_schedule_daily_enabled").Value
			$job_schedule_daily_kind = $oRecordset_Job.Fields("job_schedule_daily_kind").Value
			$job_schedule_daily_days = StringReplace($oRecordset_Job.Fields("job_schedule_daily_days").Value," ","")
			$job_schedule_monthly_enabled = $oRecordset_Job.Fields("job_schedule_monthly_enabled").Value
			$job_schedule_monthly_months = StringReplace($oRecordset_Job.Fields("job_schedule_monthly_months").Value," ","")
			$job_schedule_periodically_enabled = $oRecordset_Job.Fields("job_schedule_periodically_enabled").Value
			$oRecordset_Job.MoveNext()
		WEnd
	Else
		Return -1
	EndIf

	Local $job_schedule_daily_days_array = StringSplit($job_schedule_daily_days,",")
	Local $job_schedule_monthly_months_array = StringSplit($job_schedule_monthly_months,",")

	If $parent_job_id <> Null Then
		$sql = ""
		If StringInStr($sDriver,"SQL Server") <> 0 Then
			$sql = "SELECT id, name, parent_schedule_id," & @CRLF & _
					"    schedule.value('(//OptionsScheduleAfterJob/IsEnabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_afterjob_enabled," & @CRLF & _
					"    schedule.value('(//OptionsDaily/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_daily_enabled," & @CRLF & _
					"    schedule.value('(//OptionsDaily/Kind/text())[1]', 'VARCHAR(MAX)') AS job_schedule_daily_kind," & @CRLF & _
					"    STUFF((" & @CRLF & _
					"        SELECT ', ' + x.EMonth.value('.', 'VARCHAR(MAX)')" & @CRLF & _
					"        FROM schedule.nodes('(//OptionsDaily/Days/DayOfWeek)') AS x(EMonth)" & @CRLF & _
					"        FOR XML PATH('')" & @CRLF & _
					"    ), 1, 2, '') AS job_schedule_daily_days," & @CRLF & _
					"    schedule.value('(//OptionsPeriodically/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_periodically_enabled," & @CRLF & _
					"    schedule.value('(//OptionsMonthly/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_monthly_enabled," & @CRLF & _
					"    STUFF((" & @CRLF & _
					"        SELECT ', ' + x.EMonth.value('.', 'VARCHAR(MAX)')" & @CRLF & _
					"        FROM schedule.nodes('(//OptionsMonthly/Months/EMonth)') AS x(EMonth)" & @CRLF & _
					"        FOR XML PATH('')" & @CRLF & _
					"    ), 1, 2, '') AS job_schedule_monthly_months" & @CRLF & _
					"FROM " & $MSSQL_JobsView & " WHERE id = '" & $parent_job_id & "';"
		Else
			$sql = "SELECT id, name, parent_schedule_id," & @CRLF & _
					"    (xpath('//OptionsScheduleAfterJob/IsEnabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_afterjob_enabled," & @CRLF & _
					"    (xpath('//OptionsDaily/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_daily_enabled," & @CRLF & _
					"    (xpath('//OptionsDaily/Kind/text()', xmlparse(document schedule)))[1]::text AS job_schedule_daily_kind," & @CRLF & _
					"    array_to_string(array(SELECT unnest(xpath('//OptionsDaily/Days/DayOfWeek/text()', xmlparse(document schedule)))), ', ') AS job_schedule_daily_days," & @CRLF & _
					"    (xpath('//OptionsMonthly/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_monthly_enabled," & @CRLF & _
					"    (xpath('//OptionsPeriodically/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_periodically_enabled," & @CRLF & _
					"    array_to_string(array(SELECT unnest(xpath('//OptionsMonthly/Months/EMonth/text()', xmlparse(document schedule)))), ', ') AS job_schedule_monthly_months" & @CRLF & _
					"FROM public.jobsview WHERE id = '" & $parent_job_id & "';"
		EndIf

		$oRecordset_Job = _SqlRetrieveData($sql)
		If IsObj($oRecordset_Job) Then
			While Not $oRecordset_Job.EOF
				$parent_schedule_id = $oRecordset_Job.Fields("parent_schedule_id").Value
				$job_schedule_afterjob_enabled = $oRecordset_Job.Fields("job_schedule_afterjob_enabled").Value
				$job_schedule_afterjob_name = $oRecordset_Job.Fields("name").Value
				$job_schedule_daily_enabled = $oRecordset_Job.Fields("job_schedule_daily_enabled").Value
				$job_schedule_daily_kind = $oRecordset_Job.Fields("job_schedule_daily_kind").Value
				$job_schedule_daily_days = StringReplace($oRecordset_Job.Fields("job_schedule_daily_days").Value," ","")
				$job_schedule_daily_days_array = StringSplit($job_schedule_daily_days,",")
				$job_schedule_monthly_enabled = $oRecordset_Job.Fields("job_schedule_monthly_enabled").Value
				$job_schedule_monthly_months = StringReplace($oRecordset_Job.Fields("job_schedule_monthly_months").Value," ","")
				$job_schedule_monthly_months_array = StringSplit($job_schedule_monthly_months,",")
				$job_schedule_periodically_enabled = $oRecordset_Job.Fields("job_schedule_periodically_enabled").Value
				$oRecordset_Job.MoveNext()
			WEnd
		EndIf
	EndIf

	If $job_schedule_afterjob_enabled = "true" And $parent_schedule_id <> Null And $parent_schedule_id <> "00000000-0000-0000-0000-000000000000" Then
		$sql = ""
		If StringInStr($sDriver,"SQL Server") <> 0 Then
			$sql = "WITH ParentHierarchy AS (" & @CRLF & _
					"    SELECT id, name, parent_schedule_id, schedule" & @CRLF & _
					"    FROM " & $MSSQL_JobsView & "" & @CRLF & _
					"    WHERE id = '" & $parent_schedule_id & "'" & @CRLF & _
					"    UNION ALL" & @CRLF & _
					"    SELECT j.id, j.name, j.parent_schedule_id, j.schedule" & @CRLF & _
					"    FROM " & $MSSQL_JobsView & " j" & @CRLF & _
					"    INNER JOIN ParentHierarchy ph ON j.id = ph.parent_schedule_id" & @CRLF & _
					")" & @CRLF & _
					"SELECT id, name," & @CRLF & _
					"    schedule.value('(//OptionsScheduleAfterJob/IsEnabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_afterjob_enabled," & @CRLF & _
					"    schedule.value('(//OptionsDaily/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_daily_enabled," & @CRLF & _
					"    schedule.value('(//OptionsDaily/Kind/text())[1]', 'VARCHAR(MAX)') AS job_schedule_daily_kind," & @CRLF & _
					"    STUFF((" & @CRLF & _
					"        SELECT ', ' + x.EMonth.value('.', 'VARCHAR(MAX)')" & @CRLF & _
					"        FROM schedule.nodes('(//OptionsDaily/Days/DayOfWeek)') AS x(EMonth)" & @CRLF & _
					"        FOR XML PATH('')" & @CRLF & _
					"    ), 1, 2, '') AS job_schedule_daily_days," & @CRLF & _
					"    schedule.value('(//OptionsPeriodically/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_periodically_enabled," & @CRLF & _
					"    schedule.value('(//OptionsMonthly/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_monthly_enabled," & @CRLF & _
					"    STUFF((" & @CRLF & _
					"        SELECT ', ' + x.EMonth.value('.', 'VARCHAR(MAX)')" & @CRLF & _
					"        FROM schedule.nodes('(//OptionsMonthly/Months/EMonth)') AS x(EMonth)" & @CRLF & _
					"        FOR XML PATH('')" & @CRLF & _
					"    ), 1, 2, '') AS job_schedule_monthly_months" & @CRLF & _
					"FROM ParentHierarchy WHERE parent_schedule_id IS NULL;"
		Else
			$sql = "WITH RECURSIVE ParentHierarchy AS (" & @CRLF & _
					"    SELECT id, name, parent_schedule_id, schedule" & @CRLF & _
					"    FROM public.jobsview" & @CRLF & _
					"    WHERE id = '" & $parent_schedule_id & "'" & @CRLF & _
					"    UNION ALL" & @CRLF & _
					"    SELECT j.id, j.name, j.parent_schedule_id, j.schedule" & @CRLF & _
					"    FROM public.jobsview j" & @CRLF & _
					"    INNER JOIN ParentHierarchy ph ON j.id = ph.parent_schedule_id" & @CRLF & _
					"    WHERE ph.parent_schedule_id IS NOT NULL" & @CRLF & _
					")" & @CRLF & _
					"SELECT id, name," & @CRLF & _
					"    (xpath('//OptionsScheduleAfterJob/IsEnabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_afterjob_enabled," & @CRLF & _
					"    (xpath('//OptionsDaily/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_daily_enabled," & @CRLF & _
					"    (xpath('//OptionsDaily/Kind/text()', xmlparse(document schedule)))[1]::text AS job_schedule_daily_kind," & @CRLF & _
					"    array_to_string(array(SELECT unnest(xpath('//OptionsDaily/Days/DayOfWeek/text()', xmlparse(document schedule)))), ', ') AS job_schedule_daily_days," & @CRLF & _
					"    (xpath('//OptionsMonthly/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_monthly_enabled," & @CRLF & _
					"    (xpath('//OptionsPeriodically/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_periodically_enabled," & @CRLF & _
					"    array_to_string(array(SELECT unnest(xpath('//OptionsMonthly/Months/EMonth/text()', xmlparse(document schedule)))), ', ') AS job_schedule_monthly_months" & @CRLF & _
					"FROM ParentHierarchy WHERE parent_schedule_id IS NULL;"
		EndIf

		Local $oRecordset_Schedule = _SqlRetrieveData($sql)
		If IsObj($oRecordset_Schedule) Then
			While Not $oRecordset_Schedule.EOF
				$job_schedule_afterjob_name = $oRecordset_Schedule.Fields("name").Value
				$job_schedule_daily_enabled = $oRecordset_Schedule.Fields("job_schedule_daily_enabled").Value
				$job_schedule_daily_kind = $oRecordset_Schedule.Fields("job_schedule_daily_kind").Value
				$job_schedule_daily_days = StringReplace($oRecordset_Schedule.Fields("job_schedule_daily_days").Value," ","")
				$job_schedule_daily_days_array = StringSplit($job_schedule_daily_days,",")
				$job_schedule_monthly_enabled = $oRecordset_Schedule.Fields("job_schedule_monthly_enabled").Value
				$job_schedule_monthly_months = StringReplace($oRecordset_Schedule.Fields("job_schedule_monthly_months").Value," ","")
				$job_schedule_monthly_months_array = StringSplit($job_schedule_monthly_months,",")
				$job_schedule_periodically_enabled = $oRecordset_Schedule.Fields("job_schedule_periodically_enabled").Value
				$oRecordset_Schedule.MoveNext()
			WEnd
		EndIf
	EndIf

	Local $nextBackupDate = ""
	If $job_schedule_daily_enabled = "true" Then
		If $job_schedule_daily_kind = "Everyday" Then
			$nextBackupDate = _DateAdd("D",1,$backup_creation_time_date)
		Else
			$nextBackupDate = CalculateNextBackupDate($backup_creation_time_date,"D",$job_schedule_daily_days_array)
		EndIf
	EndIf
	If $job_schedule_periodically_enabled = "true" Then
		$nextBackupDate = _DateAdd("D",1,$backup_creation_time_date)
	EndIf
	If $job_schedule_monthly_enabled = "true" Then
		$nextBackupDate = CalculateNextBackupDate($backup_creation_time_date,"M",$job_schedule_monthly_months_array)
	EndIf

	If $nextBackupDate = "" Then Return -1

	Local $next_run_value = $nextBackupDate
	If $job_schedule_afterjob_enabled = "true" And $job_schedule_afterjob_name <> "" Then
		$next_run_value = "After Job: " & $job_schedule_afterjob_name
	EndIf

	Return $next_run_value
EndFunc

Func _SqlBackupConfigurationJobPostgres()
	Local $sql = ""
	$sql = "WITH latest_js AS (" & @CRLF & _
		   "    SELECT" & @CRLF & _
		   "        js.id," & @CRLF & _
		   "        js.job_id," & @CRLF & _
		   "        js.job_name," & @CRLF & _
		   "        js.job_type," & @CRLF & _
		   "        js.creation_time," & @CRLF & _
		   "        js.end_time," & @CRLF & _
		   "        js.state," & @CRLF & _
		   "        js.result," & @CRLF & _
		   "        js.reason" & @CRLF & _
		   "    FROM public." & chr(34) & "backup.model.jobsessions" & chr(34) & " js" & @CRLF & _
		   "    WHERE js.job_type = " & $BackupConfigurationJobType & @CRLF & _
		   "    ORDER BY js.creation_time DESC" & @CRLF & _
		   "    LIMIT 1" & @CRLF & _
		   ")" & @CRLF & _
		   "SELECT" & @CRLF & _
		   "    js.job_id::TEXT AS job_id," & @CRLF & _
		   "    js.job_name::TEXT AS job_name," & @CRLF & _
		   "    js.job_type::TEXT AS job_type," & @CRLF & _
		   "    js.creation_time::TEXT AS creation_time," & @CRLF & _
		   "    js.end_time::TEXT AS end_time," & @CRLF & _
		   "    js.state::TEXT AS job_state," & @CRLF & _
		   "    js.result::TEXT AS job_result," & @CRLF & _
		   "    js.reason::TEXT AS job_reason," & @CRLF & _
		   "    sl.status::TEXT AS log_status," & @CRLF & _
		   "    sl.title::TEXT AS log_title," & @CRLF & _
		   "    sl." & chr(34) & "desc" & chr(34) & "::TEXT AS log_desc," & @CRLF & _
		   "    sl.starttimeutc::TEXT AS log_starttimeutc," & @CRLF & _
		   "    sl.updatetimeutc::TEXT AS log_updatetimeutc" & @CRLF & _
		   "FROM latest_js js" & @CRLF & _
		   "LEFT JOIN LATERAL (" & @CRLF & _
		   "    SELECT status, title, " & chr(34) & "desc" & chr(34) & ", starttimeutc, updatetimeutc" & @CRLF & _
		   "    FROM public.sessionlog" & @CRLF & _
		   "    WHERE sessionid = js.id" & @CRLF & _
		   "    ORDER BY updatetimeutc DESC NULLS LAST" & @CRLF & _
		   "    LIMIT 1" & @CRLF & _
		   ") sl ON true;"
	Return $sql
EndFunc

Func _SqlRepoDiscovery($sDriver)
	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 Then
		$sql = "SELECT CAST(r.name AS VARCHAR(255)) AS repo_name " & @CRLF & _
			   "FROM " & $MSSQL_BackupRepositories & " r;"
	Else
		$sql = "SELECT name::TEXT AS repo_name FROM public.backuprepositories;"
	EndIf
	Return $sql
EndFunc

Func _SqlRepoMetrics($sDriver)
	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 Then
		$sql = "SELECT " & @CRLF & _
			   "    CAST(r.name AS VARCHAR(255)) AS repo_name, " & @CRLF & _
			   "    CAST(ISNULL(c.totalspace, 0) AS BIGINT) AS total_space, " & @CRLF & _
			   "    CAST(ISNULL(c.freespace, 0) AS BIGINT) AS free_space " & @CRLF & _
			   "FROM " & $MSSQL_BackupRepositories & " r " & @CRLF & _
			   "LEFT JOIN " & $MSSQL_BackupRepositoryContainerRepos & " cr ON cr.repositoryid = r.id " & @CRLF & _
			   "LEFT JOIN " & $MSSQL_BackupRepositoryContainer & " c ON c.id = cr.id;"
	Else
		$sql = "SELECT name::TEXT AS repo_name, total_space::BIGINT AS total_space, free_space::BIGINT AS free_space " & @CRLF & _
			   "FROM public.backuprepositoriesview;"
	EndIf
	Return $sql
EndFunc

Func _SqlVmByJobDiscovery($sDriver)
	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 Then
		$sql = "SELECT " & @CRLF & _
			   "    CAST(COALESCE(pj.name, j.name) AS VARCHAR(255)) AS job_name, " & @CRLF & _
			   "    CAST(o.object_name AS VARCHAR(255)) AS vm_name, " & @CRLF & _
			   "    CAST(j.schedule_enabled AS VARCHAR(10)) AS schedule_enabled, " & @CRLF & _
			   "    CAST(j.is_deleted AS VARCHAR(10)) AS is_job_deleted, " & @CRLF & _
			   "    CAST(ISNULL(j.options.value('(//JobOptionsRoot/RunManually/text())[1]', 'VARCHAR(MAX)'), 'false') AS VARCHAR(255)) AS job_options_runmanually " & @CRLF & _
			   "FROM " & $MSSQL_ObjectsInJobsView & " ij " & @CRLF & _
			   "JOIN " & $MSSQL_JobsView & " j ON j.id = ij.job_id " & @CRLF & _
			   "LEFT JOIN " & $MSSQL_JobsView & " pj ON j.parent_job_id = pj.id " & @CRLF & _
			   "JOIN " & $MSSQL_ObjectsView & " o ON o.id = ij.object_id " & @CRLF & _
			   "WHERE LOWER(o.viobject_type) = 'vm';"
	Else
		$sql = "SELECT " & @CRLF & _
			   "    COALESCE(pj.name, j.name)::TEXT AS job_name, " & @CRLF & _
			   "    o.object_name::TEXT AS vm_name, " & @CRLF & _
			   "    j.schedule_enabled::TEXT AS schedule_enabled, " & @CRLF & _
			   "    j.is_deleted::TEXT AS is_job_deleted, " & @CRLF & _
			   "    (xpath('//JobOptionsRoot/RunManually/text()', xmlparse(document j.options)))[1]::text AS job_options_runmanually " & @CRLF & _
			   "FROM public." & chr(34) & "objectsinjobsview" & chr(34) & " ij " & @CRLF & _
			   "JOIN public." & chr(34) & "jobsview" & chr(34) & " j ON j.id = ij.job_id " & @CRLF & _
			   "LEFT JOIN public." & chr(34) & "jobsview" & chr(34) & " pj ON j.parent_job_id = pj.id " & @CRLF & _
			   "JOIN public." & chr(34) & "objectsview" & chr(34) & " o ON o.id = ij.object_id " & @CRLF & _
			   "WHERE LOWER(o.viobject_type) = 'vm';"
	EndIf
	Return $sql
EndFunc

Func _SqlVmTasksWithRetry($sDriver)
	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 Then
		$sql = "WITH LatestOriginal AS (" & @CRLF & _
			   "    SELECT js.id, js.job_id, js.creation_time, " & @CRLF & _
			   "           ROW_NUMBER() OVER (PARTITION BY js.job_id ORDER BY js.creation_time DESC) AS rn " & @CRLF & _
			   "    FROM " & $MSSQL_JobSessions & " js " & @CRLF & _
			   "    WHERE js.orig_session_id IS NULL " & @CRLF & _
			   "), SessionGroup AS (" & @CRLF & _
			   "    SELECT js.id, js.job_id, lo.id AS orig_session_id " & @CRLF & _
			   "    FROM " & $MSSQL_JobSessions & " js " & @CRLF & _
			   "    JOIN LatestOriginal lo ON (js.id = lo.id OR js.orig_session_id = lo.id) " & @CRLF & _
			   "    WHERE lo.rn = 1 " & @CRLF & _
			   "), TaskRanked AS (" & @CRLF & _
			   "    SELECT sg.job_id, COALESCE(pj.name, j.name) AS job_name, bts.object_name AS vm_name, " & @CRLF & _
			   "           bts.status, bts.reason, bts.creation_time, bts.end_time, " & @CRLF & _
			   "           ROW_NUMBER() OVER (PARTITION BY sg.job_id, bts.object_name ORDER BY bts.creation_time DESC) AS rn " & @CRLF & _
			   "    FROM SessionGroup sg " & @CRLF & _
			   "    JOIN " & $MSSQL_BackupTaskSessions & " bts ON bts.session_id = sg.id " & @CRLF & _
			   "    JOIN " & $MSSQL_JobsView & " j ON j.id = sg.job_id " & @CRLF & _
			   "    LEFT JOIN " & $MSSQL_JobsView & " pj ON j.parent_job_id = pj.id " & @CRLF & _
			   ") " & @CRLF & _
			   "SELECT CAST(job_name AS VARCHAR(255)) AS job_name, " & @CRLF & _
			   "       CAST(vm_name AS VARCHAR(255)) AS vm_name, " & @CRLF & _
			   "       CAST(status AS VARCHAR(255)) AS status, " & @CRLF & _
			   "       CAST(reason AS VARCHAR(255)) AS reason, " & @CRLF & _
			   "       CAST(creation_time AS VARCHAR(255)) AS creation_time, " & @CRLF & _
			   "       CAST(end_time AS VARCHAR(255)) AS end_time " & @CRLF & _
			   "FROM TaskRanked WHERE rn = 1;"
	Else
		$sql = "WITH LatestOriginal AS (" & @CRLF & _
			   "    SELECT js.id, js.job_id, js.creation_time, " & @CRLF & _
			   "           ROW_NUMBER() OVER (PARTITION BY js.job_id ORDER BY js.creation_time DESC) AS rn " & @CRLF & _
			   "    FROM public." & chr(34) & "backup.model.jobsessions" & chr(34) & " js " & @CRLF & _
			   "    WHERE js.orig_session_id IS NULL " & @CRLF & _
			   "), SessionGroup AS (" & @CRLF & _
			   "    SELECT js.id, js.job_id, lo.id AS orig_session_id " & @CRLF & _
			   "    FROM public." & chr(34) & "backup.model.jobsessions" & chr(34) & " js " & @CRLF & _
			   "    JOIN LatestOriginal lo ON (js.id = lo.id OR js.orig_session_id = lo.id) " & @CRLF & _
			   "    WHERE lo.rn = 1 " & @CRLF & _
			   "), TaskRanked AS (" & @CRLF & _
			   "    SELECT sg.job_id, COALESCE(pj.name, j.name) AS job_name, bts.object_name AS vm_name, " & @CRLF & _
			   "           bts.status, bts.reason, bts.creation_time, bts.end_time, " & @CRLF & _
			   "           ROW_NUMBER() OVER (PARTITION BY sg.job_id, bts.object_name ORDER BY bts.creation_time DESC) AS rn " & @CRLF & _
			   "    FROM SessionGroup sg " & @CRLF & _
			   "    JOIN public." & chr(34) & "backup.model.backuptasksessions" & chr(34) & " bts ON bts.session_id = sg.id " & @CRLF & _
			   "    JOIN public." & chr(34) & "jobsview" & chr(34) & " j ON j.id = sg.job_id " & @CRLF & _
			   "    LEFT JOIN public." & chr(34) & "jobsview" & chr(34) & " pj ON j.parent_job_id = pj.id " & @CRLF & _
			   ") " & @CRLF & _
			   "SELECT job_name::TEXT AS job_name, vm_name::TEXT AS vm_name, " & @CRLF & _
			   "       status::TEXT AS status, reason::TEXT AS reason, creation_time::TEXT AS creation_time, end_time::TEXT AS end_time " & @CRLF & _
			   "FROM TaskRanked WHERE rn = 1;"
	EndIf
	Return $sql
EndFunc

Func _SqlVmTasksFallbackMssql($sDriver)
	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 Then
		$sql = "WITH TaskRanked AS (" & @CRLF & _
			   "    SELECT " & @CRLF & _
			   "        CAST(COALESCE(pj.name, j.name) AS VARCHAR(255)) AS job_name, " & @CRLF & _
			   "        CAST(bts.object_name AS VARCHAR(255)) AS vm_name, " & @CRLF & _
			   "        CAST(ISNULL(bts.status, -1) AS VARCHAR(255)) AS status, " & @CRLF & _
			   "        CAST(ISNULL(bts.reason, '') AS VARCHAR(255)) AS reason, " & @CRLF & _
			   "        CAST(bts.creation_time AS VARCHAR(255)) AS creation_time, " & @CRLF & _
			   "        CAST(bts.end_time AS VARCHAR(255)) AS end_time, " & @CRLF & _
			   "        ROW_NUMBER() OVER (PARTITION BY j.id, bts.object_name ORDER BY bts.creation_time DESC) AS rn " & @CRLF & _
			   "    FROM " & $MSSQL_BackupTaskSessions & " bts " & @CRLF & _
			   "    JOIN " & $MSSQL_JobSessions & " js ON js.id = bts.session_id " & @CRLF & _
			   "    JOIN " & $MSSQL_JobsView & " j ON j.id = js.job_id " & @CRLF & _
			   "    LEFT JOIN " & $MSSQL_JobsView & " pj ON j.parent_job_id = pj.id " & @CRLF & _
			   ") " & @CRLF & _
			   "SELECT job_name, vm_name, status, reason, creation_time, end_time " & @CRLF & _
			   "FROM TaskRanked WHERE rn = 1;"
	EndIf
	Return $sql
EndFunc

Func _SqlTapeDiscovery($sDriver, $tapeType)
	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 Then
		$sql = "SELECT CAST(t.name AS VARCHAR(255)) AS job_name, CAST(t.type AS INT) AS job_type, " & @CRLF & _
			   "       CAST(t.schedule_enabled AS VARCHAR(10)) AS schedule_enabled, " & @CRLF & _
			   "       CAST(0 AS VARCHAR(10)) AS is_job_deleted, " & @CRLF & _
			   "       'false' AS job_options_runmanually " & @CRLF & _
			   "FROM " & $MSSQL_TapeJobs & " t WHERE t.type = " & $tapeType & ";"
	Else
		$sql = "SELECT name::TEXT AS job_name, type::INT AS job_type, " & @CRLF & _
			   "       schedule_enabled::TEXT AS schedule_enabled, " & @CRLF & _
			   "       'false'::TEXT AS is_job_deleted, " & @CRLF & _
			   "       'false'::TEXT AS job_options_runmanually " & @CRLF & _
			   "FROM public." & chr(34) & "tape.jobs" & chr(34) & " WHERE type = " & $tapeType & ";"
	EndIf
	Return $sql
EndFunc

Func _SqlTapeStatus($sDriver)
	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 Then
		$sql = "WITH LastSession AS (" & @CRLF & _
			   "    SELECT js.job_id, js.job_type, js.result, js.state, js.description, js.creation_time, js.end_time, " & @CRLF & _
			   "           ROW_NUMBER() OVER (PARTITION BY js.job_id ORDER BY js.creation_time DESC) AS rn " & @CRLF & _
			   "    FROM " & $MSSQL_JobSessions & " js " & @CRLF & _
			   "    WHERE js.job_type IN (24,28) " & @CRLF & _
			   ") " & @CRLF & _
			   "SELECT CAST(t.name AS VARCHAR(255)) AS job_name, " & @CRLF & _
			   "       CAST(ls.job_type AS INT) AS job_type, " & @CRLF & _
			   "       CAST(ISNULL(ls.result, -1) AS VARCHAR(255)) AS job_result, " & @CRLF & _
			   "       CAST(ISNULL(ls.state, -1) AS VARCHAR(255)) AS job_state, " & @CRLF & _
			   "       CAST(ISNULL(ls.description, '') AS VARCHAR(255)) AS job_reason, " & @CRLF & _
			   "       CAST(ls.creation_time AS VARCHAR(255)) AS creation_time, " & @CRLF & _
			   "       CAST(ls.end_time AS VARCHAR(255)) AS end_time " & @CRLF & _
			   "FROM " & $MSSQL_TapeJobs & " t " & @CRLF & _
			   "LEFT JOIN LastSession ls ON ls.job_id = t.id AND ls.rn = 1 " & @CRLF & _
			   "WHERE t.type IN (24,28);"
	Else
		$sql = "WITH LastSession AS (" & @CRLF & _
			   "    SELECT js.job_id, js.job_type, js.result, js.state, js.description, js.creation_time, js.end_time, " & @CRLF & _
			   "           ROW_NUMBER() OVER (PARTITION BY js.job_id ORDER BY js.creation_time DESC) AS rn " & @CRLF & _
			   "    FROM public." & chr(34) & "backup.model.jobsessions" & chr(34) & " js " & @CRLF & _
			   "    WHERE js.job_type IN (24,28) " & @CRLF & _
			   ") " & @CRLF & _
			   "SELECT t.name::TEXT AS job_name, " & @CRLF & _
			   "       ls.job_type::INT AS job_type, " & @CRLF & _
			   "       COALESCE(ls.result, -1)::TEXT AS job_result, " & @CRLF & _
			   "       COALESCE(ls.state, -1)::TEXT AS job_state, " & @CRLF & _
			   "       COALESCE(ls.description, '')::TEXT AS job_reason, " & @CRLF & _
			   "       ls.creation_time::TEXT AS creation_time, " & @CRLF & _
			   "       ls.end_time::TEXT AS end_time " & @CRLF & _
			   "FROM public." & chr(34) & "tape.jobs" & chr(34) & " t " & @CRLF & _
			   "LEFT JOIN LastSession ls ON ls.job_id = t.id AND ls.rn = 1 " & @CRLF & _
			   "WHERE t.type IN (24,28);"
	EndIf
	Return $sql
EndFunc

Func _SqlBackupSyncDiscovery($sDriver)
	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 Then
		$sql = "SELECT CAST(j.name AS VARCHAR(255)) AS job_name, CAST(j.id AS VARCHAR(255)) AS job_id, " & @CRLF & _
			   "       CAST(j.schedule_enabled AS VARCHAR(10)) AS schedule_enabled, " & @CRLF & _
			   "       CAST(j.is_deleted AS VARCHAR(10)) AS is_job_deleted, " & @CRLF & _
			   "       CAST(ISNULL(j.options.value('(//JobOptionsRoot/RunManually/text())[1]', 'VARCHAR(MAX)'), 'false') AS VARCHAR(255)) AS job_options_runmanually " & @CRLF & _
			   "FROM " & $MSSQL_JobsView & " j WHERE j.type = 51 AND j.is_deleted = 0;"
	Else
		$sql = "SELECT name::TEXT AS job_name, id::TEXT AS job_id, " & @CRLF & _
			   "       schedule_enabled::TEXT AS schedule_enabled, " & @CRLF & _
			   "       is_deleted::TEXT AS is_job_deleted, " & @CRLF & _
			   "       (xpath('//JobOptionsRoot/RunManually/text()', xmlparse(document options)))[1]::text AS job_options_runmanually " & @CRLF & _
			   "FROM public." & chr(34) & "jobsview" & chr(34) & " WHERE type = 51 AND is_deleted = false;"
	EndIf
	Return $sql
EndFunc

Func _SqlBackupSyncStatus($sDriver)
	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 Then
		$sql = "WITH LastSession AS (" & @CRLF & _
			   "    SELECT js.job_id, js.job_type, js.result, js.state, js.description, js.creation_time, js.end_time, " & @CRLF & _
			   "           ROW_NUMBER() OVER (PARTITION BY js.job_id ORDER BY js.creation_time DESC) AS rn " & @CRLF & _
			   "    FROM " & $MSSQL_JobSessions & " js " & @CRLF & _
			   "    WHERE js.job_type = 51 " & @CRLF & _
			   ") " & @CRLF & _
			   "SELECT CAST(j.id AS VARCHAR(255)) AS job_id, " & @CRLF & _
			   "       CAST(j.name AS VARCHAR(255)) AS job_name, " & @CRLF & _
			   "       CAST(ISNULL(ls.result, -1) AS VARCHAR(255)) AS job_result, " & @CRLF & _
			   "       CAST(ISNULL(ls.state, -1) AS VARCHAR(255)) AS job_state, " & @CRLF & _
			   "       CAST(ISNULL(ls.description, '') AS VARCHAR(255)) AS job_reason, " & @CRLF & _
			   "       CAST(ls.creation_time AS VARCHAR(255)) AS creation_time, " & @CRLF & _
			   "       CAST(ls.end_time AS VARCHAR(255)) AS end_time " & @CRLF & _
			   "FROM " & $MSSQL_JobsView & " j " & @CRLF & _
			   "LEFT JOIN LastSession ls ON ls.job_id = j.id AND ls.rn = 1 " & @CRLF & _
			   "WHERE j.type = 51;"
	Else
		$sql = "WITH LastSession AS (" & @CRLF & _
			   "    SELECT js.job_id, js.job_type, js.result, js.state, js.description, js.creation_time, js.end_time, " & @CRLF & _
			   "           ROW_NUMBER() OVER (PARTITION BY js.job_id ORDER BY js.creation_time DESC) AS rn " & @CRLF & _
			   "    FROM public." & chr(34) & "backup.model.jobsessions" & chr(34) & " js " & @CRLF & _
			   "    WHERE js.job_type = 51 " & @CRLF & _
			   ") " & @CRLF & _
			   "SELECT j.id::TEXT AS job_id, " & @CRLF & _
			   "       j.name::TEXT AS job_name, " & @CRLF & _
			   "       COALESCE(ls.result, -1)::TEXT AS job_result, " & @CRLF & _
			   "       COALESCE(ls.state, -1)::TEXT AS job_state, " & @CRLF & _
			   "       COALESCE(ls.description, '')::TEXT AS job_reason, " & @CRLF & _
			   "       ls.creation_time::TEXT AS creation_time, " & @CRLF & _
			   "       ls.end_time::TEXT AS end_time " & @CRLF & _
			   "FROM public." & chr(34) & "jobsview" & chr(34) & " j " & @CRLF & _
			   "LEFT JOIN LastSession ls ON ls.job_id = j.id AND ls.rn = 1 " & @CRLF & _
			   "WHERE j.type = 51;"
	EndIf
	Return $sql
EndFunc

Func _SqlEndpointDiscovery($sDriver)
	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 Then
		$sql = "SELECT CAST(j.name AS VARCHAR(255)) AS job_name, CAST(j.type AS INT) AS job_type, " & @CRLF & _
			   "       CAST(j.schedule_enabled AS VARCHAR(10)) AS schedule_enabled, " & @CRLF & _
			   "       CAST(j.is_deleted AS VARCHAR(10)) AS is_job_deleted, " & @CRLF & _
			   "       CAST(ISNULL(j.options.value('(//JobOptionsRoot/RunManually/text())[1]', 'VARCHAR(MAX)'), 'false') AS VARCHAR(255)) AS job_options_runmanually " & @CRLF & _
			   "FROM " & $MSSQL_JobsView & " j WHERE j.type IN (12111,22000,33000) AND j.is_deleted = 0;"
	Else
		$sql = "SELECT name::TEXT AS job_name, type::INT AS job_type, " & @CRLF & _
			   "       schedule_enabled::TEXT AS schedule_enabled, " & @CRLF & _
			   "       is_deleted::TEXT AS is_job_deleted, " & @CRLF & _
			   "       (xpath('//JobOptionsRoot/RunManually/text()', xmlparse(document options)))[1]::text AS job_options_runmanually " & @CRLF & _
			   "FROM public." & chr(34) & "jobsview" & chr(34) & " " & @CRLF & _
			   "WHERE type IN (12111,22000,33000) AND is_deleted = false;"
	EndIf
	Return $sql
EndFunc

Func _SqlAgentDiscovery($sDriver)
	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 Then
		$sql = "SELECT CAST(j.id AS VARCHAR(255)) AS job_id, CAST(j.parent_job_id AS VARCHAR(255)) AS parent_job_id, " & @CRLF & _
			   "       CAST(j.name AS VARCHAR(255)) AS job_name, CAST(j.type AS INT) AS job_type, " & @CRLF & _
			   "       CAST(j.schedule_enabled AS VARCHAR(10)) AS schedule_enabled, " & @CRLF & _
			   "       CAST(j.is_deleted AS VARCHAR(10)) AS is_job_deleted, " & @CRLF & _
			   "       CAST(ISNULL(j.options.value('(//JobOptionsRoot/RunManually/text())[1]', 'VARCHAR(MAX)'), 'false') AS VARCHAR(255)) AS job_options_runmanually " & @CRLF & _
			   "FROM " & $MSSQL_JobsView & " j " & @CRLF & _
			   "WHERE j.type = 4000 AND j.is_deleted = 0 " & @CRLF & _
			   "UNION " & @CRLF & _
		"SELECT CAST(j.id AS VARCHAR(255)) AS job_id, CAST(j.parent_job_id AS VARCHAR(255)) AS parent_job_id, " & @CRLF & _
		"       CAST(j.name AS VARCHAR(255)) AS job_name, CAST(j.type AS INT) AS job_type, " & @CRLF & _
		"       CAST(j.schedule_enabled AS VARCHAR(10)) AS schedule_enabled, " & @CRLF & _
		"       CAST(j.is_deleted AS VARCHAR(10)) AS is_job_deleted, " & @CRLF & _
		"       CAST(ISNULL(j.options.value('(//JobOptionsRoot/RunManually/text())[1]', 'VARCHAR(MAX)'), 'false') AS VARCHAR(255)) AS job_options_runmanually " & @CRLF & _
			   "FROM " & $MSSQL_JobsView & " j " & @CRLF & _
			   "WHERE j.type = 12000 AND j.is_deleted = 0 AND j.parent_job_id IS NULL " & @CRLF & _
			   "UNION " & @CRLF & _
		"SELECT CAST(j.id AS VARCHAR(255)) AS job_id, CAST(j.parent_job_id AS VARCHAR(255)) AS parent_job_id, " & @CRLF & _
		"       CAST(j.name AS VARCHAR(255)) AS job_name, CAST(j.type AS INT) AS job_type, " & @CRLF & _
		"       CAST(j.schedule_enabled AS VARCHAR(10)) AS schedule_enabled, " & @CRLF & _
		"       CAST(j.is_deleted AS VARCHAR(10)) AS is_job_deleted, " & @CRLF & _
		"       CAST(ISNULL(j.options.value('(//JobOptionsRoot/RunManually/text())[1]', 'VARCHAR(MAX)'), 'false') AS VARCHAR(255)) AS job_options_runmanually " & @CRLF & _
			   "FROM " & $MSSQL_JobsView & " j " & @CRLF & _
			   "WHERE j.type = 12002 AND j.is_deleted = 0 " & @CRLF & _
			   "  AND NOT EXISTS (SELECT 1 FROM " & $MSSQL_JobsView & " j2 " & @CRLF & _
			   "                  WHERE j2.type = 4000 AND j2.is_deleted = 0 " & @CRLF & _
			   "                    AND j2.name LIKE j.name + ' - %') " & @CRLF & _
			   "UNION " & @CRLF & _
		"SELECT CAST(j.id AS VARCHAR(255)) AS job_id, CAST(j.parent_job_id AS VARCHAR(255)) AS parent_job_id, " & @CRLF & _
		"       CAST(j.name AS VARCHAR(255)) AS job_name, CAST(j.type AS INT) AS job_type, " & @CRLF & _
		"       CAST(j.schedule_enabled AS VARCHAR(10)) AS schedule_enabled, " & @CRLF & _
		"       CAST(j.is_deleted AS VARCHAR(10)) AS is_job_deleted, " & @CRLF & _
		"       CAST(ISNULL(j.options.value('(//JobOptionsRoot/RunManually/text())[1]', 'VARCHAR(MAX)'), 'false') AS VARCHAR(255)) AS job_options_runmanually " & @CRLF & _
			   "FROM " & $MSSQL_JobsView & " j " & @CRLF & _
			   "WHERE j.type = 12003 AND j.is_deleted = 0;"
	Else
		$sql = "SELECT id::TEXT AS job_id, parent_job_id::TEXT AS parent_job_id, " & @CRLF & _
			   "       name::TEXT AS job_name, type::INT AS job_type, " & @CRLF & _
			   "       schedule_enabled::TEXT AS schedule_enabled, " & @CRLF & _
			   "       is_deleted::TEXT AS is_job_deleted, " & @CRLF & _
			   "       (xpath('//JobOptionsRoot/RunManually/text()', xmlparse(document options)))[1]::text AS job_options_runmanually " & @CRLF & _
			   "FROM public." & chr(34) & "jobsview" & chr(34) & " " & @CRLF & _
			   "WHERE type = 4000 AND is_deleted = false " & @CRLF & _
			   "UNION " & @CRLF & _
		"SELECT j.id::TEXT AS job_id, j.parent_job_id::TEXT AS parent_job_id, " & @CRLF & _
		"       j.name::TEXT AS job_name, j.type::INT AS job_type, " & @CRLF & _
		"       j.schedule_enabled::TEXT AS schedule_enabled, " & @CRLF & _
		"       j.is_deleted::TEXT AS is_job_deleted, " & @CRLF & _
		"       (xpath('//JobOptionsRoot/RunManually/text()', xmlparse(document j.options)))[1]::text AS job_options_runmanually " & @CRLF & _
			   "FROM public." & chr(34) & "jobsview" & chr(34) & " j " & @CRLF & _
			   "WHERE j.type = 12000 AND j.is_deleted = false AND j.parent_job_id IS NULL " & @CRLF & _
			   "UNION " & @CRLF & _
		"SELECT j.id::TEXT AS job_id, j.parent_job_id::TEXT AS parent_job_id, " & @CRLF & _
		"       j.name::TEXT AS job_name, j.type::INT AS job_type, " & @CRLF & _
		"       j.schedule_enabled::TEXT AS schedule_enabled, " & @CRLF & _
		"       j.is_deleted::TEXT AS is_job_deleted, " & @CRLF & _
		"       (xpath('//JobOptionsRoot/RunManually/text()', xmlparse(document j.options)))[1]::text AS job_options_runmanually " & @CRLF & _
			   "FROM public." & chr(34) & "jobsview" & chr(34) & " j " & @CRLF & _
			   "WHERE j.type = 12002 AND j.is_deleted = false " & @CRLF & _
			   "  AND NOT EXISTS (SELECT 1 FROM public." & chr(34) & "jobsview" & chr(34) & " j2 " & @CRLF & _
			   "                  WHERE j2.type = 4000 AND j2.is_deleted = false " & @CRLF & _
			   "                    AND j2.name LIKE (j.name || ' - %')) " & @CRLF & _
			   "UNION " & @CRLF & _
		"SELECT j.id::TEXT AS job_id, j.parent_job_id::TEXT AS parent_job_id, " & @CRLF & _
		"       j.name::TEXT AS job_name, j.type::INT AS job_type, " & @CRLF & _
		"       j.schedule_enabled::TEXT AS schedule_enabled, " & @CRLF & _
		"       j.is_deleted::TEXT AS is_job_deleted, " & @CRLF & _
		"       (xpath('//JobOptionsRoot/RunManually/text()', xmlparse(document j.options)))[1]::text AS job_options_runmanually " & @CRLF & _
			   "FROM public." & chr(34) & "jobsview" & chr(34) & " j " & @CRLF & _
			   "WHERE j.type = 12003 AND j.is_deleted = false;"
	EndIf
	Return $sql
EndFunc

Func _SqlEndpointStatus($sDriver)
	Local $sql = ""
	If StringInStr($sDriver,"SQL Server") <> 0 Then
		$sql = "WITH LastSession AS (" & @CRLF & _
			   "    SELECT js.job_id, js.job_type, js.result, js.state, js.description, js.creation_time, js.end_time, " & @CRLF & _
			   "           ROW_NUMBER() OVER (PARTITION BY js.job_id ORDER BY js.creation_time DESC) AS rn " & @CRLF & _
			   "    FROM " & $MSSQL_JobSessions & " js " & @CRLF & _
			   "    WHERE js.job_type IN (4000,12000,12002,12003,12005,12006,12007,12008,12009,12111,22000,33000) " & @CRLF & _
			   ") " & @CRLF & _
			   "SELECT CAST(j.id AS VARCHAR(255)) AS job_id, " & @CRLF & _
			   "       CAST(j.name AS VARCHAR(255)) AS job_name, " & @CRLF & _
			   "       CAST(j.type AS INT) AS base_job_type, " & @CRLF & _
			   "       CAST(j.parent_job_id AS VARCHAR(255)) AS parent_job_id, " & @CRLF & _
			   "       CAST(ls.job_type AS INT) AS job_type, " & @CRLF & _
			   "       CAST(ISNULL(ls.result, -1) AS VARCHAR(255)) AS job_result, " & @CRLF & _
			   "       CAST(ISNULL(ls.state, -1) AS VARCHAR(255)) AS job_state, " & @CRLF & _
			   "       CAST(ISNULL(ls.description, '') AS VARCHAR(255)) AS job_reason, " & @CRLF & _
			   "       CAST(ls.creation_time AS VARCHAR(255)) AS creation_time, " & @CRLF & _
			   "       CAST(ls.end_time AS VARCHAR(255)) AS end_time " & @CRLF & _
			   "FROM " & $MSSQL_JobsView & " j " & @CRLF & _
			   "LEFT JOIN LastSession ls ON ls.job_id = j.id AND ls.rn = 1 " & @CRLF & _
			   "WHERE j.type IN (4000,12000,12002,12003,12005,12006,12007,12008,12009,12111,22000,33000) " & @CRLF & _
			   "  AND NOT (j.type = 12000 AND j.parent_job_id IS NOT NULL) " & @CRLF & _
			   "  AND NOT (j.type = 12002 AND EXISTS (SELECT 1 FROM " & $MSSQL_JobsView & " j2 " & @CRLF & _
			   "           WHERE j2.type = 4000 AND j2.is_deleted = 0 AND j2.name LIKE j.name + ' - %'));"
	Else
		$sql = "WITH LastSession AS (" & @CRLF & _
			   "    SELECT js.job_id, js.job_type, js.result, js.state, js.description, js.creation_time, js.end_time, " & @CRLF & _
			   "           ROW_NUMBER() OVER (PARTITION BY js.job_id ORDER BY js.creation_time DESC) AS rn " & @CRLF & _
			   "    FROM public." & chr(34) & "backup.model.jobsessions" & chr(34) & " js " & @CRLF & _
			   "    WHERE js.job_type IN (4000,12000,12002,12003,12005,12006,12007,12008,12009,12111,22000,33000) " & @CRLF & _
			   ") " & @CRLF & _
			   "SELECT j.id::TEXT AS job_id, " & @CRLF & _
			   "       j.name::TEXT AS job_name, " & @CRLF & _
			   "       j.type::INT AS base_job_type, " & @CRLF & _
			   "       j.parent_job_id::TEXT AS parent_job_id, " & @CRLF & _
			   "       ls.job_type::INT AS job_type, " & @CRLF & _
			   "       COALESCE(ls.result, -1)::TEXT AS job_result, " & @CRLF & _
			   "       COALESCE(ls.state, -1)::TEXT AS job_state, " & @CRLF & _
			   "       COALESCE(ls.description, '')::TEXT AS job_reason, " & @CRLF & _
			   "       ls.creation_time::TEXT AS creation_time, " & @CRLF & _
			   "       ls.end_time::TEXT AS end_time " & @CRLF & _
			   "FROM public." & chr(34) & "jobsview" & chr(34) & " j " & @CRLF & _
			   "LEFT JOIN LastSession ls ON ls.job_id = j.id AND ls.rn = 1 " & @CRLF & _
			   "WHERE j.type IN (4000,12000,12002,12003,12005,12006,12007,12008,12009,12111,22000,33000) " & @CRLF & _
			   "  AND NOT (j.type = 12000 AND j.parent_job_id IS NOT NULL) " & @CRLF & _
			   "  AND NOT (j.type = 12002 AND EXISTS (SELECT 1 FROM public." & chr(34) & "jobsview" & chr(34) & " j2 " & @CRLF & _
			   "           WHERE j2.type = 4000 AND j2.is_deleted = false AND j2.name LIKE (j.name || ' - %')));"
	EndIf
	Return $sql
EndFunc

Func BackupConfigurationJobData($Recordset)
	If $Recordset.EOF Then Return

	While Not $Recordset.EOF

		Local $MonitorEnabled = 1

		Local $job_name_original = $Recordset.Fields("job_name").Value
		Local $job_name = StringReplace($job_name_original,",","_")

		Local $backup_creation_time = $Recordset.Fields("creation_time").Value
		Local $backup_creation_time_date = _DateVeeamFormat($backup_creation_time)

		Local $backup_end_time = $Recordset.Fields("end_time").Value
		Local $backup_end_time_date = _DateVeeamFormat($backup_end_time)

		Local $backup_state = $Recordset.Fields("job_state").Value
		Local $backup_result = $Recordset.Fields("job_result").Value
		Local $backup_reason = $Recordset.Fields("job_reason").Value

		Local $Duration = 0
		If $backup_end_time_date > $backup_creation_time_date Then
			$Duration = _DateDiff('n',$backup_creation_time_date,$backup_end_time_date)
		Else
			Local $duration_message = "Backup starting or in progress"
			$backup_reason = _AppendReason($backup_reason, $duration_message)
			$Duration = $duration_message
		EndIf

		Local $backup_task_status = $Recordset.Fields("log_status").Value
		Local $backup_task_reason = $Recordset.Fields("log_title").Value
		If $backup_task_reason = "" Then
			$backup_task_reason = $Recordset.Fields("log_desc").Value
		EndIf

		Local $DateDiff = _DateDiff('D',$backup_creation_time_date,_NowCalc())

		$JobsCount += 1

		$Array_Disc_Tmp &= $Comma & "{" & chr(34) & "{#VEEAMJOB}" & chr(34) & ":" & chr(34) & $job_name & "" & chr(34) & "}"
		$Comma = ","

		If Not _IsNumericValue($Duration) Then
			Local $computed = _ComputeDurationMinutes($backup_creation_time_date, $backup_end_time_date)
			If $computed >= 0 Then
				$Duration = $computed
			Else
				$backup_reason = _AppendReason($backup_reason, $Duration)
				$Duration = -1
			EndIf
		EndIf

		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.monitored[" & $job_name & "]",$MonitorEnabled)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.job.state[" & $job_name & "]",$backup_state)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.job.result[" & $job_name & "]",$backup_result)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.job.reason[" & $job_name & "]",$backup_reason)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.status[" & $job_name & "]",$backup_task_status)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.reason[" & $job_name & "]",$backup_task_reason)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.creationtime[" & $job_name & "]",$backup_creation_time_date)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.endtime[" & $job_name & "]",$backup_end_time_date)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.datediff[" & $job_name & "]",$DateDiff)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.duration[" & $job_name & "]",$Duration)

		$Recordset.MoveNext()
	WEnd
EndFunc

Func DiscoveryData($Recordset,$sDriver)
	$count = 0
	Local $BackupJobDiscoveryLog = ""
	Local $seen_jobs = "|"

	While Not $Recordset.EOF

		Local $MonitorEnabled = 0

		$count += 1

		Local $job_name_original = $Recordset.Fields("job_name").Value
		Local $job_name = StringReplace($job_name_original,",","_")

		Local $is_schedule_enabled = $Recordset.Fields("is_schedule_enabled").Value
		Local $is_job_deleted = $Recordset.Fields("is_job_deleted").Value
		Local $backup_job_type = $Recordset.Fields("job_type").Value
		Local $job_options_runmanually = $Recordset.Fields("job_options_runmanually").Value

		Local $CheckJobType = _IsMonitoredBackupJobType($backup_job_type)

		Local $MonitorEnabled = 0
		If $CheckJobType Then
			$MonitorEnabled = _IsJobMonitorEnabled($is_schedule_enabled, $is_job_deleted, $job_options_runmanually)
		EndIf

		$MonitorEnabled = _ApplyBlacklistMonitorEnabled($MonitorEnabled, $job_name)

		If $MonitorEnabled = 1 Then
			$JobsCount += 1
		EndIf

		If (Not _ToBool($is_job_deleted) And $CheckJobType) Then
			$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.monitored[" & $job_name & "]",$MonitorEnabled)
			If StringInStr($seen_jobs, "|" & $job_name & "|") = 0 Then
				$seen_jobs &= $job_name & "|"
				$BackupJobsCount += 1
				$Array_Disc_Tmp &= $Comma & "{" & chr(34) & "{#VEEAMJOB}" & chr(34) & ":" & chr(34) & $job_name & "" & chr(34) & "}"
				$Comma = ","
				$BackupJobDiscoveryLog &= "-> BackupJob: " & $job_name & @CRLF
			EndIf
		Endif

		If $Debug > 0 then
			_logmsg($LogFile,"   IsScheduleEnabled: " & $is_schedule_enabled,true,true)
			_logmsg($LogFile,"   IsJobDeleted: " & $is_job_deleted,true,true)
			_logmsg($LogFile,"   MonitoredType (" & $backup_job_type & "): " & $CheckJobType,true,true)
		EndIf

		$Recordset.MoveNext()
	WEnd

EndFunc

Func DiscoveryRepoData($Recordset)
	While Not $Recordset.EOF
		Local $repo_name_original = $Recordset.Fields("repo_name").Value
		Local $repo_name = StringReplace($repo_name_original,",","_")
		$RepoCount += 1

		$Array_Disc_Repo_Tmp &= $Comma_Repo & "{" & chr(34) & "{#VEEAMREPO}" & chr(34) & ":" & chr(34) & $repo_name & "" & chr(34) & "}"
		$Comma_Repo = ","

		$Recordset.MoveNext()
	WEnd
EndFunc

Func DiscoveryVmByJobData($Recordset)
	Local $seen_pairs = "|"
	While Not $Recordset.EOF
		Local $job_name_original = $Recordset.Fields("job_name").Value
		Local $job_name = StringReplace($job_name_original,",","_")
		Local $vm_name_original = $Recordset.Fields("vm_name").Value
		Local $vm_name = StringReplace($vm_name_original,",","_")
		Local $schedule_enabled = $Recordset.Fields("schedule_enabled").Value
		Local $is_job_deleted = $Recordset.Fields("is_job_deleted").Value
		Local $job_options_runmanually = $Recordset.Fields("job_options_runmanually").Value
		Local $MonitorEnabled = _IsJobMonitorEnabled($schedule_enabled, $is_job_deleted, $job_options_runmanually)
		$MonitorEnabled = _ApplyBlacklistMonitorEnabled($MonitorEnabled, $job_name, $vm_name)
		Local $pair_key = $job_name & "|" & $vm_name
		If StringInStr($seen_pairs, "|" & $pair_key & "|") = 0 Then
			$seen_pairs &= $pair_key & "|"
			$VmByJobCount += 1
			If $MonitorEnabled = 1 Then $VmByJobMonitoredCount += 1
			$Array_Disc_VM_Tmp &= $Comma_VM & "{" & chr(34) & "{#VEEAMJOB}" & chr(34) & ":" & chr(34) & $job_name & "" & chr(34) & "," & _
				chr(34) & "{#VEEAMVM}" & chr(34) & ":" & chr(34) & $vm_name & "" & chr(34) & "}"
			$Comma_VM = ","
			$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.monitored[" & $job_name & "," & $vm_name & "]",$MonitorEnabled)
		EndIf

		$Recordset.MoveNext()
	WEnd
EndFunc

Func DiscoveryTapeData($Recordset, $tapeType)
	Local $label = "Tape"
	If $tapeType = 24 Then
		$label = "Tape File-to-Tape"
	ElseIf $tapeType = 28 Then
		$label = "Tape VM"
	EndIf

	While Not $Recordset.EOF
		Local $job_name_original = $Recordset.Fields("job_name").Value
		Local $job_name = StringReplace($job_name_original,",","_")
		Local $job_type = $tapeType
		If $tapeType = 24 Then
			$TapeFileCount += 1
		Else
			$TapeVmCount += 1
		EndIf

		If $Debug > 0 Then
			_logmsg($LogFile,"Tape Discovery -> JobName: " & $job_name_original & " | Type: " & $job_type,true,true)
		EndIf

		If $tapeType = 24 Then
			$Array_Disc_TapeF_Tmp &= $Comma_TapeF & "{" & chr(34) & "{#TAPEJOB}" & chr(34) & ":" & chr(34) & $job_name & "" & chr(34) & "}"
			$Comma_TapeF = ","
		Else
			$Array_Disc_TapeV_Tmp &= $Comma_TapeV & "{" & chr(34) & "{#TAPEJOB}" & chr(34) & ":" & chr(34) & $job_name & "" & chr(34) & "}"
			$Comma_TapeV = ","
		EndIf

		$Recordset.MoveNext()
	WEnd
EndFunc

Func DiscoveryBackupSyncData($Recordset,$sDriver)
	While Not $Recordset.EOF
		Local $job_id = $Recordset.Fields("job_id").Value
		Local $parent_job_id = $Recordset.Fields("parent_job_id").Value
		Local $job_name_original = $Recordset.Fields("job_name").Value
		Local $job_name = StringReplace($job_name_original,",","_")
		Local $job_type = $Recordset.Fields("job_type").Value
		Local $schedule_enabled = $Recordset.Fields("schedule_enabled").Value
		Local $is_job_deleted = $Recordset.Fields("is_job_deleted").Value
		Local $job_options_runmanually = $Recordset.Fields("job_options_runmanually").Value
		Local $MonitorEnabled = _IsJobMonitorEnabled($schedule_enabled, $is_job_deleted, $job_options_runmanually)
		$MonitorEnabled = _ApplyBlacklistMonitorEnabled($MonitorEnabled, $job_name)
		$BackupSyncCount += 1
		If $MonitorEnabled = 1 Then $BackupSyncMonitoredCount += 1

		If $Debug > 0 Then
			_logmsg($LogFile,"BackupSync Discovery -> JobName: " & $job_name_original & " | Type: " & $job_type,true,true)
		EndIf

		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupsync.monitored[" & $job_name & "]",$MonitorEnabled)

		$Array_Disc_BackupSync_Tmp &= $Comma_BackupSync & "{" & chr(34) & "{#BACKUPSYNCJOB}" & chr(34) & ":" & chr(34) & $job_name & "" & chr(34) & "}"
		$Comma_BackupSync = ","

		$Recordset.MoveNext()
	WEnd
EndFunc

Func DiscoveryEndpointData($Recordset)
	While Not $Recordset.EOF
		Local $job_id = $Recordset.Fields("job_id").Value
		Local $parent_job_id = $Recordset.Fields("parent_job_id").Value
		Local $job_name_original = $Recordset.Fields("job_name").Value
		Local $job_name = StringReplace($job_name_original,",","_")
		Local $job_type = $Recordset.Fields("job_type").Value
		Local $schedule_enabled = $Recordset.Fields("schedule_enabled").Value
		Local $is_job_deleted = $Recordset.Fields("is_job_deleted").Value
		Local $job_options_runmanually = $Recordset.Fields("job_options_runmanually").Value
		Local $MonitorEnabled = _IsJobMonitorEnabled($schedule_enabled, $is_job_deleted, $job_options_runmanually)
		$MonitorEnabled = _ApplyBlacklistMonitorEnabled($MonitorEnabled, $job_name)
		If $MonitorEnabled = 1 Then $EndpointMonitoredCount += 1
		$EndpointCount += 1

		If $Debug > 0 Then
			_logmsg($LogFile,"Endpoint Discovery -> JobName: " & $job_name_original & " | Type: " & $job_type,true,true)
		EndIf

		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.endpoint.monitored[" & $job_name & "]",$MonitorEnabled)

		$Array_Disc_Endpoint_Tmp &= $Comma_Endpoint & "{" & chr(34) & "{#ENDPOINTJOB}" & chr(34) & ":" & chr(34) & $job_name & "" & chr(34) & "}"
		$Comma_Endpoint = ","

		$Recordset.MoveNext()
	WEnd
EndFunc

Func DiscoveryAgentData($Recordset)
	Local $PolicyLog = ""
	Local $BackupLog = ""

	While Not $Recordset.EOF
		Local $job_id = $Recordset.Fields("job_id").Value
		Local $parent_job_id = $Recordset.Fields("parent_job_id").Value
		Local $job_name_original = $Recordset.Fields("job_name").Value
		Local $job_name = StringReplace($job_name_original,",","_")
		Local $job_type = $Recordset.Fields("job_type").Value
		Local $schedule_enabled = $Recordset.Fields("schedule_enabled").Value
		Local $is_job_deleted = $Recordset.Fields("is_job_deleted").Value
		Local $job_options_runmanually = $Recordset.Fields("job_options_runmanually").Value
		Local $MonitorEnabled = _IsJobMonitorEnabled($schedule_enabled, $is_job_deleted, $job_options_runmanually)
		$MonitorEnabled = _ApplyBlacklistMonitorEnabled($MonitorEnabled, $job_name)
		If $Debug > 0 Then
			_logmsg($LogFile,"Agent Discovery -> JobName: " & $job_name_original & " | Type: " & $job_type & _
					" | ParentJobId: " & $parent_job_id & " | JobId: " & $job_id,true,true)
		EndIf

		Local $agent_prefix = _ResolveAgentMetricPrefix($job_type, $parent_job_id)
		If $MonitorEnabled = 1 Then
			If $agent_prefix = "backup.veeam.customchecks.agent.policy" Then
				$AgentPolicyMonitoredCount += 1
			Else
				$AgentBackupMonitoredCount += 1
			EndIf
		EndIf
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$agent_prefix & ".monitored[" & $job_name & "]",$MonitorEnabled)

		If $agent_prefix = "backup.veeam.customchecks.agent.policy" Then
			$AgentPolicyCount += 1
			$Array_Disc_AgentPolicy_Tmp &= $Comma_AgentPolicy & "{" & chr(34) & "{#AGENTJOB}" & chr(34) & ":" & chr(34) & $job_name & "" & chr(34) & "}"
			$Comma_AgentPolicy = ","
			$PolicyLog &= "-> Agent Policy: " & $job_name_original & @CRLF
		Else
			$AgentBackupCount += 1
			$Array_Disc_AgentBackup_Tmp &= $Comma_AgentBackup & "{" & chr(34) & "{#AGENTJOB}" & chr(34) & ":" & chr(34) & $job_name & "" & chr(34) & "}"
			$Comma_AgentBackup = ","
			$BackupLog &= "-> Agent Backup: " & $job_name_original & @CRLF
		EndIf

		$Recordset.MoveNext()
	WEnd

EndFunc

Func BackupData($Recordset,$sDriver)
	_logmsg($LogFile,"",true,true)
	_logmsg($LogFile,"List BackupJob jobs found:",true,true)
	Local $BackupJobStatusLog = ""

;~ 	; Old Code
;~ 	Local $aRecordsetArray = _ADO_Recordset_ToArray($Recordset, False)
;~ 	Local $aRecordset_inner = _ADO_RecordsetArray_GetContent($aRecordsetArray)
;~ 	Local $iColumn_count = UBound($aRecordset_inner, $UBOUND_COLUMNS)

;~ 	; Ottieni il numero di righe e colonne
;~ 	Local $iRowCount = UBound($aRecordset_inner, $UBOUND_ROWS)
;~ 	Local $iColumnCount = UBound($aRecordset_inner, $UBOUND_COLUMNS)

;~ 	; Cicla ogni riga
;~ 	For $iRow = 0 To $iRowCount - 1
;~ 		; Cicla ogni colonna nella riga
;~ 		For $iCol = 0 To $iColumnCount - 1
;~ 			; Leggi il valore della cella
;~ 			Local $value = $aRecordset_inner[$iRow][$iCol]
;~ 			ConsoleWrite("Valore alla riga " & $iRow & ", colonna " & $iCol & ": " & $value & @CRLF)
;~ 		Next
;~ 	Next
	;ConsoleWrite(@CRLF & "R: " & UBound($aRecordset_inner) & " C: " & $iColumn_count & @CRLF & @CRLF)

	; job_id
	; job_name
	; repository_id
	; repository_name
	; job_type
	; job_schedule
	; is_job_deleted
	; latest_job_result
	; is_schedule_enabled
	; parent_schedule_id
	; backup_hosts
	; backup_job_type
	; backup_state
	; backup_result
	; backup_creation_time
	; backup_end_time
	; backup_task_status
	; backup_task_reason
	; backup_task_session_id
	; job_schedule_afterjob_enabled
	; job_schedule_daily_enabled
	; job_schedule_daily_kind
	; job_schedule_daily_days
	; job_schedule_monthly_enabled
	; job_schedule_periodically_enabled
	; job_schedule_monthly_months

	While Not $Recordset.EOF

		;Local $MonitorEnabled = 0

		Local $job_name_original = $Recordset.Fields("job_name").Value
		Local $job_name = StringReplace($job_name_original,",","_")

		Local $backup_creation_time = $Recordset.Fields("backup_creation_time").Value
		Local $backup_creation_time_date = _DateVeeamFormat($backup_creation_time)

		Local $backup_end_time = $Recordset.Fields("backup_end_time").Value
		Local $backup_end_time_date = _DateVeeamFormat($backup_end_time)

		Local $Duration = ""
		If $backup_end_time_date > $backup_creation_time_date then
			$Duration = _DateDiff('n',$backup_creation_time_date,$backup_end_time_date)
		Else
			$Duration = "Backup starting or in progress"
		Endif

		Local $backup_state = $Recordset.Fields("backup_state").Value
		Local $backup_result = $Recordset.Fields("backup_result").Value
		Local $backup_reason = $Recordset.Fields("backup_reason").Value

		Local $backup_task_status = $Recordset.Fields("backup_task_status").Value
		Local $backup_task_reason = $Recordset.Fields("backup_task_reason").Value

		Local $is_schedule_enabled = $Recordset.Fields("is_schedule_enabled").Value
		Local $is_job_deleted = $Recordset.Fields("is_job_deleted").Value

		Local $parent_job_id = $Recordset.Fields("parent_job_id").Value

		Local $job_type = $Recordset.Fields("job_type").Value
		Local $backup_job_type = $Recordset.Fields("backup_job_type").Value

		; Filter by internal monitored backup job types
		Local $CheckJobType = _IsMonitoredBackupJobType($backup_job_type) Or _IsMonitoredBackupJobType($job_type)
		If Not $CheckJobType Then
			$Recordset.MoveNext()
			ContinueLoop
		EndIf

;~ 		If ( $is_schedule_enabled = 1 and $is_job_deleted = 0 and $CheckJobType > 0 ) then
;~ 			$MonitorEnabled = 1
;~ 		Endif

;~ 		If $MonitorEnabled = 1 Then
;~ 			$JobsCount += 1
;~ 		EndIf

		;for $i = 0 to 24
		;	ConsoleWrite(@CRLF & $Recordset.Fields($i).Value)
		;Next

		Local $job_schedule_afterjob_enabled = $Recordset.Fields("job_schedule_afterjob_enabled").Value
		Local $job_schedule_afterjob_name = ""
		Local $job_schedule_daily_enabled = $Recordset.Fields("job_schedule_daily_enabled").Value
		Local $job_schedule_daily_kind = $Recordset.Fields("job_schedule_daily_kind").Value
		Local $job_schedule_daily_days = StringReplace($Recordset.Fields("job_schedule_daily_days").Value," ","")
		Local $job_schedule_daily_days_array = StringSplit($job_schedule_daily_days,",")
		Local $job_schedule_monthly_enabled = $Recordset.Fields("job_schedule_monthly_enabled").Value
		Local $job_schedule_monthly_months = StringReplace($Recordset.Fields("job_schedule_monthly_months").Value," ","")
		Local $job_schedule_monthly_months_array = StringSplit($job_schedule_monthly_months,",")
		Local $job_schedule_periodically_enabled = $Recordset.Fields("job_schedule_periodically_enabled").Value

		Local $parent_schedule_id = $Recordset.Fields("parent_schedule_id").Value

		;ConsoleWrite(@CRLF & "P:" & $parent_job_id & @CRLF)

		If $parent_job_id <> Null then
			$sql = ""
			If StringInStr($sDriver,"SQL Server") <> 0 then
				$sql = "SELECT" & @CRLF & _
						"    id," & @CRLF & _
						"    name," & @CRLF & _
						"	 parent_schedule_id," & @CRLF & _
						"    schedule.value('(//OptionsScheduleAfterJob/IsEnabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_afterjob_enabled," & @CRLF & _
						"    schedule.value('(//OptionsDaily/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_daily_enabled," & @CRLF & _
						"    schedule.value('(//OptionsDaily/Kind/text())[1]', 'VARCHAR(MAX)') AS job_schedule_daily_kind," & @CRLF & _
						"    STUFF((" & @CRLF & _
						"        SELECT ', ' + x.EMonth.value('.', 'VARCHAR(MAX)')" & @CRLF & _
						"        FROM schedule.nodes('(//OptionsDaily/Days/DayOfWeek)') AS x(EMonth)" & @CRLF & _
						"        FOR XML PATH('')" & @CRLF & _
						"    ), 1, 2, '') AS job_schedule_daily_days," & @CRLF & _
						"    schedule.value('(//OptionsPeriodically/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_periodically_enabled," & @CRLF & _
						"    schedule.value('(//OptionsMonthly/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_monthly_enabled," & @CRLF & _
						"    STUFF((" & @CRLF & _
						"        SELECT ', ' + x.EMonth.value('.', 'VARCHAR(MAX)')" & @CRLF & _
						"        FROM schedule.nodes('(//OptionsMonthly/Months/EMonth)') AS x(EMonth)" & @CRLF & _
						"        FOR XML PATH('')" & @CRLF & _
						"    ), 1, 2, '') AS job_schedule_monthly_months" & @CRLF & _
						"FROM " & $MSSQL_JobsView & "" & @CRLF & _
						"WHERE id = '" & $parent_job_id & "';"

			Endif

			If StringInStr($sDriver,"PostgreSQL") <> 0 then
				$sql = "SELECT " & @CRLF & _
					   "    id," & @CRLF & _
					   "    name," & @CRLF & _
					   "	parent_schedule_id," & @CRLF & _
					   "    (xpath('//OptionsScheduleAfterJob/IsEnabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_afterjob_enabled," & @CRLF & _
					   "    (xpath('//OptionsDaily/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_daily_enabled," & @CRLF & _
					   "    (xpath('//OptionsDaily/Kind/text()', xmlparse(document schedule)))[1]::text AS job_schedule_daily_kind," & @CRLF & _
					   "    array_to_string(" & @CRLF & _
					   "        array(" & @CRLF & _
					   "            SELECT unnest(xpath('//OptionsDaily/Days/DayOfWeek/text()', xmlparse(document schedule)))" & @CRLF & _
					   "        ), ', '" & @CRLF & _
					   "    ) AS job_schedule_daily_days," & @CRLF & _
					   "    (xpath('//OptionsMonthly/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_monthly_enabled," & @CRLF & _
					   "    (xpath('//OptionsPeriodically/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_periodically_enabled," & @CRLF & _
					   "    array_to_string(" & @CRLF & _
					   "        array(" & @CRLF & _
					   "            SELECT unnest(xpath('//OptionsMonthly/Months/EMonth/text()', xmlparse(document schedule)))" & @CRLF & _
					   "        ), ', '" & @CRLF & _
					   "    ) AS job_schedule_monthly_months" & @CRLF & _
					   "FROM public.jobsview" & @CRLF & _
					   "WHERE id = '" & $parent_job_id & "';"
			Endif

			;ConsoleWrite(@CRLF & $sql & @CRLF)
			;exit
			$oRecordset_Job = _SqlRetrieveData($sql)

			If IsObj($oRecordset_Job) then
				While Not $oRecordset_Job.EOF

					$parent_schedule_id = $oRecordset_Job.Fields("parent_schedule_id").Value
					$job_schedule_afterjob_enabled = $oRecordset_Job.Fields("job_schedule_afterjob_enabled").Value
					$job_schedule_afterjob_name = $oRecordset_Job.Fields("name").Value
					$job_schedule_daily_enabled = $oRecordset_Job.Fields("job_schedule_daily_enabled").Value
					$job_schedule_daily_kind = $oRecordset_Job.Fields("job_schedule_daily_kind").Value
					$job_schedule_daily_days = StringReplace($oRecordset_Job.Fields("job_schedule_daily_days").Value," ","")
					$job_schedule_daily_days_array = StringSplit($job_schedule_daily_days,",")
					$job_schedule_monthly_enabled = $oRecordset_Job.Fields("job_schedule_monthly_enabled").Value
					$job_schedule_monthly_months = StringReplace($oRecordset_Job.Fields("job_schedule_monthly_months").Value," ","")
					$job_schedule_monthly_months_array = StringSplit($job_schedule_monthly_months,",")
					$job_schedule_periodically_enabled = $oRecordset_Job.Fields("job_schedule_periodically_enabled").Value

					$oRecordset_Job.MoveNext()
				WEnd
			else
				_logmsg($LogFile,"Error Parent Job SQL: " & $oRecordset_Job,true,true)
			Endif
		Endif

		If $job_schedule_afterjob_enabled = "true" and $parent_schedule_id <> Null and $parent_schedule_id <> "00000000-0000-0000-0000-000000000000" then
			$sql = ""
			If StringInStr($sDriver,"SQL Server") <> 0 then
				$sql = "WITH ParentHierarchy AS (" & @CRLF & _
						"    -- First level: Retrieve initial record" & @CRLF & _
						"    SELECT" & @CRLF & _
						"        id," & @CRLF & _
						"        parent_schedule_id," & @CRLF & _
						"        name," & @CRLF & _
						"        schedule" & @CRLF & _
						"    FROM" & @CRLF & _
						"        dbo.jobsview" & @CRLF & _
						"    WHERE" & @CRLF & _
						"        id = '" & $parent_schedule_id & "'" & @CRLF & _
						"    UNION ALL" & @CRLF & _
						"    -- Next levels: Search for parent_schedule_id" & @CRLF & _
						"    SELECT" & @CRLF & _
						"        j.id," & @CRLF & _
						"        j.parent_schedule_id," & @CRLF & _
						"        j.name," & @CRLF & _
						"        j.schedule" & @CRLF & _
						"    FROM" & @CRLF & _
						"        dbo.jobsview j" & @CRLF & _
						"    INNER JOIN" & @CRLF & _
						"        ParentHierarchy ph ON j.id = ph.parent_schedule_id" & @CRLF & _
						")" & @CRLF & _
						"-- Select last record found (parent_schedule_id = NULL)" & @CRLF & _
						"SELECT" & @CRLF & _
						"    id," & @CRLF & _
						"    name," & @CRLF & _
						"    schedule.value('(//OptionsScheduleAfterJob/IsEnabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_afterjob_enabled," & @CRLF & _
						"    schedule.value('(//OptionsDaily/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_daily_enabled," & @CRLF & _
						"    schedule.value('(//OptionsDaily/Kind/text())[1]', 'VARCHAR(MAX)') AS job_schedule_daily_kind," & @CRLF & _
						"    STUFF((" & @CRLF & _
						"        SELECT ', ' + x.EMonth.value('.', 'VARCHAR(MAX)')" & @CRLF & _
						"        FROM schedule.nodes('(//OptionsDaily/Days/DayOfWeek)') AS x(EMonth)" & @CRLF & _
						"        FOR XML PATH('')" & @CRLF & _
						"    ), 1, 2, '') AS job_schedule_daily_days," & @CRLF & _
						"    schedule.value('(//OptionsPeriodically/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_periodically_enabled," & @CRLF & _
						"    schedule.value('(//OptionsMonthly/Enabled/text())[1]', 'VARCHAR(MAX)') AS job_schedule_monthly_enabled," & @CRLF & _
						"    STUFF((" & @CRLF & _
						"        SELECT ', ' + x.EMonth.value('.', 'VARCHAR(MAX)')" & @CRLF & _
						"        FROM schedule.nodes('(//OptionsMonthly/Months/EMonth)') AS x(EMonth)" & @CRLF & _
						"        FOR XML PATH('')" & @CRLF & _
						"    ), 1, 2, '') AS job_schedule_monthly_months" & @CRLF & _
						"FROM ParentHierarchy" & @CRLF & _
						"WHERE parent_schedule_id IS NULL;"

			Endif

			If StringInStr($sDriver,"PostgreSQL") <> 0 then
				$sql = "WITH RECURSIVE ParentHierarchy AS (" & @CRLF & _
					   "    -- First level: Retrieve initial record" & @CRLF & _
					   "    SELECT " & @CRLF & _
					   "        id," & @CRLF & _
					   "        parent_schedule_id," & @CRLF & _
					   "        name," & @CRLF & _
					   "        schedule" & @CRLF & _
					   "    FROM " & @CRLF & _
					   "        public.jobsview" & @CRLF & _
					   "    WHERE " & @CRLF & _
					   "        id = '" & $parent_schedule_id & "'" & @CRLF & _
					   "    UNION ALL" & @CRLF & _
					   "    -- Next levels: Search for parent_schedule_id" & @CRLF & _
					   "    SELECT " & @CRLF & _
					   "        j.id," & @CRLF & _
					   "        j.parent_schedule_id," & @CRLF & _
					   "        j.name," & @CRLF & _
					   "        j.schedule" & @CRLF & _
					   "    FROM " & @CRLF & _
					   "        public.jobsview j" & @CRLF & _
					   "    INNER JOIN " & @CRLF & _
					   "        ParentHierarchy ph ON j.id = ph.parent_schedule_id" & @CRLF & _
					   "    WHERE " & @CRLF & _
					   "        ph.parent_schedule_id IS NOT NULL" & @CRLF & _
					   ")" & @CRLF & _
					   "-- Select last record found (parent_schedule_id = NULL)" & @CRLF & _
					   "SELECT " & @CRLF & _
					   "    id," & @CRLF & _
					   "    name," & @CRLF & _
					   "    (xpath('//OptionsScheduleAfterJob/IsEnabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_afterjob_enabled," & @CRLF & _
					   "    (xpath('//OptionsDaily/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_daily_enabled," & @CRLF & _
					   "    (xpath('//OptionsDaily/Kind/text()', xmlparse(document schedule)))[1]::text AS job_schedule_daily_kind," & @CRLF & _
					   "    array_to_string(" & @CRLF & _
					   "        array(" & @CRLF & _
					   "            SELECT unnest(xpath('//OptionsDaily/Days/DayOfWeek/text()', xmlparse(document schedule)))" & @CRLF & _
					   "        ), ', '" & @CRLF & _
					   "    ) AS job_schedule_daily_days," & @CRLF & _
					   "    (xpath('//OptionsMonthly/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_monthly_enabled," & @CRLF & _
					   "    (xpath('//OptionsPeriodically/Enabled/text()', xmlparse(document schedule)))[1]::text AS job_schedule_periodically_enabled," & @CRLF & _
					   "    array_to_string(" & @CRLF & _
					   "        array(" & @CRLF & _
					   "            SELECT unnest(xpath('//OptionsMonthly/Months/EMonth/text()', xmlparse(document schedule)))" & @CRLF & _
					   "        ), ', '" & @CRLF & _
					   "    ) AS job_schedule_monthly_months" & @CRLF & _
					   "FROM ParentHierarchy" & @CRLF & _
					   "WHERE parent_schedule_id IS NULL;"
			Endif

			$oRecordset_Schedule = _SqlRetrieveData($sql)
			If IsObj($oRecordset_Schedule) then
				While Not $oRecordset_Schedule.EOF
					$job_schedule_afterjob_name = $oRecordset_Schedule.Fields("name").Value
					$job_schedule_daily_enabled = $oRecordset_Schedule.Fields("job_schedule_daily_enabled").Value
					$job_schedule_daily_kind = $oRecordset_Schedule.Fields("job_schedule_daily_kind").Value
					$job_schedule_daily_days = StringReplace($oRecordset_Schedule.Fields("job_schedule_daily_days").Value," ","")
					$job_schedule_daily_days_array = StringSplit($job_schedule_daily_days,",")
					$job_schedule_monthly_enabled = $oRecordset_Schedule.Fields("job_schedule_monthly_enabled").Value
					$job_schedule_monthly_months = StringReplace($oRecordset_Schedule.Fields("job_schedule_monthly_months").Value," ","")
					$job_schedule_monthly_months_array = StringSplit($job_schedule_monthly_months,",")
					$job_schedule_periodically_enabled = $oRecordset_Schedule.Fields("job_schedule_periodically_enabled").Value

					$oRecordset_Schedule.MoveNext()
				WEnd
			else
				_logmsg($LogFile,"Error Parent Schedule SQL: " & $oRecordset_Schedule,true,true)
			Endif
		Endif

		$nextBackupDate = ""

		If $job_schedule_daily_enabled = "true" then
			If $job_schedule_daily_kind = "Everyday" then
				$nextBackupDate = _DateAdd("D",1,$backup_creation_time_date)
			Else
				$nextBackupDate = CalculateNextBackupDate($backup_creation_time_date,"D",$job_schedule_daily_days_array)
			EndIf
		EndIf

		If $job_schedule_periodically_enabled = "true" then
			$nextBackupDate = _DateAdd("D",1,$backup_creation_time_date)
		EndIf

		If $job_schedule_monthly_enabled = "true" then
			Local $nextBackupDate = CalculateNextBackupDate($backup_creation_time_date,"M",$job_schedule_monthly_months_array)
		EndIf

		; Old code that add 1+ day at the nextschedule
		;$checkBackupDateLate = _Dateadd("D",1,$nextBackupDate)

		;ConsoleWrite(@CRLF & $backup_creation_time_date & " - " & $nextBackupDate & @CRLF)
		$DateDiff_Check = _DateDiff('D',$backup_creation_time_date,$nextBackupDate)

		$DateDiff = _DateDiff('D',$backup_creation_time_date,_NowCalc())
		$DateDiff = $DateDiff - $DateDiff_Check

		;$Array_Disc_Tmp &= $Comma & "{" & chr(34) & "{#VEEAMJOB}" & chr(34) & ":" & chr(34) & $job_name & "" & chr(34) & "}"
		;$Comma = ","

		Local $next_run_value = $nextBackupDate
		If $job_schedule_afterjob_enabled = "true" And $job_schedule_afterjob_name <> "" Then
			$next_run_value = "After Job: " & $job_schedule_afterjob_name
		EndIf

		If Not _IsNumericValue($Duration) Then
			Local $computed = _ComputeDurationMinutes($backup_creation_time_date, $backup_end_time_date)
			If $computed >= 0 Then
				$Duration = $computed
			Else
				$backup_reason = _AppendReason($backup_reason, $Duration)
				$Duration = -1
			EndIf
		EndIf

		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.job.state[" & $job_name & "]",$backup_state)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.job.result[" & $job_name & "]",$backup_result)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.job.reason[" & $job_name & "]",$backup_reason)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.status[" & $job_name & "]",$backup_task_status)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.reason[" & $job_name & "]",$backup_task_reason)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.creationtime[" & $job_name & "]",$backup_creation_time_date)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.endtime[" & $job_name & "]",$backup_end_time_date)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.datediff[" & $job_name & "]",$DateDiff)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.duration[" & $job_name & "]",$Duration)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupjob.next_run_time[" & $job_name & "]",$next_run_value)
		$BackupJobStatusLog &= "-> BackupJob: " & $job_name_original & " Type: " & $backup_job_type & " Result: " & $backup_result & @CRLF

		; Debug
		if $Debug = 1 then
			_logmsg($LogFile,"JobName: " & $job_name_original,true,true)
			_logmsg($LogFile,"JobNameMonitoring: " & $job_name,true,true)
			;_logmsg($LogFile,"MonitoringEnabled: " & $MonitorEnabled,true,true)
			_logmsg($LogFile,"JobType: " & $backup_job_type,true,true)
			_logmsg($LogFile,"SessionState: " & $backup_state,true,true)
			_logmsg($LogFile,"SessionResult: " & $backup_result,true,true)
			_logmsg($LogFile,"SessionReason: " & $backup_reason,true,true)
			_logmsg($LogFile,"Status: " & $backup_task_status,true,true)
			_logmsg($LogFile,"Reason: " & $backup_task_reason,true,true)
			_logmsg($LogFile,"CreationTime: " & $backup_creation_time_date & " (" & $backup_creation_time & ")",true,true)
			_logmsg($LogFile,"EndTime: " & $backup_end_time_date & " (" & $backup_end_time & ")",true,true)
			_logmsg($LogFile,"Duration (Min): " & $Duration,true,true)
			_logmsg($LogFile,"DateDiff: " & $DateDiff,true,true)
			_logmsg($LogFile,"ScheduleEnabled: " & $is_schedule_enabled,true,true)
			_logmsg($LogFile,"NextSchedule: " & $nextBackupDate,true,true)
			_logmsg($LogFile,"DateDiffCheck: " & $DateDiff_Check,true,true)
			_logmsg($LogFile,"AfterJobEnabled: " & $job_schedule_afterjob_enabled,true,true)
			_logmsg($LogFile,"AfterJobName: " & $job_schedule_afterjob_name,true,true)
			_logmsg($LogFile,"DailyEnabled: " & $job_schedule_daily_enabled,true,true)
			_logmsg($LogFile,"MonthlyEnabled: " & $job_schedule_monthly_enabled,true,true)
			_logmsg($LogFile,"PeriodicallyEnabled: " & $job_schedule_periodically_enabled,true,true)
			_logmsg($LogFile,"ParentScheduleID: " & $parent_schedule_id,true,true)
			_logmsg($LogFile,"JobDeleted: " & $is_job_deleted,true,true)
		endif

		$Recordset.MoveNext()
	WEnd

	If $BackupJobStatusLog <> "" Then
		_logmsg($LogFile,StringTrimRight($BackupJobStatusLog, 2),true,true)
	EndIf

	_logmsg($LogFile,"",true,true)

EndFunc

Func RepoData($Recordset)
	_logmsg($LogFile,"",true,true)
	_logmsg($LogFile,"Repository metrics:",true,true)
	_logmsg($LogFile,"List repositories found:",true,true)

	While Not $Recordset.EOF
		Local $repo_name_original = $Recordset.Fields("repo_name").Value
		Local $repo_name = StringReplace($repo_name_original,",","_")
		Local $total_space = _NullToZero($Recordset.Fields("total_space").Value)
		Local $free_space = _NullToZero($Recordset.Fields("free_space").Value)

		Local $used_percent = 0
		If Number($total_space) > 0 Then
			$used_percent = Round(((Number($total_space) - Number($free_space)) / Number($total_space)) * 100, 2)
		EndIf

		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.repo.total_space[" & $repo_name & "]",$total_space)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.repo.free_space[" & $repo_name & "]",$free_space)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.repo.used_percent[" & $repo_name & "]",$used_percent)

		_logmsg($LogFile,"-> Repo: " & $repo_name_original & " Total: " & $total_space & " Free: " & $free_space & " Used%: " & $used_percent,true,true)
		$Recordset.MoveNext()
	WEnd

	_logmsg($LogFile,"",true,true)
EndFunc

Func VmTaskData($Recordset)
	_logmsg($LogFile,"",true,true)
	_logmsg($LogFile,"List VMs by job found:",true,true)

	While Not $Recordset.EOF
		Local $job_name_original = $Recordset.Fields("job_name").Value
		Local $job_name = StringReplace($job_name_original,",","_")
		Local $vm_name_original = $Recordset.Fields("vm_name").Value
		Local $vm_name = StringReplace($vm_name_original,",","_")
		Local $status = $Recordset.Fields("status").Value
		Local $reason = $Recordset.Fields("reason").Value
		Local $creation_time = $Recordset.Fields("creation_time").Value
		Local $end_time = $Recordset.Fields("end_time").Value
		Local $creation_time_date = _DateVeeamFormat($creation_time)
		Local $end_time_date = _DateVeeamFormat($end_time)
		Local $duration = _ComputeDurationMinutes($creation_time_date, $end_time_date)
		Local $datediff = _ComputeDateDiffDays($creation_time_date)

		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.status[" & $job_name & "," & $vm_name & "]",$status)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.reason[" & $job_name & "," & $vm_name & "]",$reason)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.creationtime[" & $job_name & "," & $vm_name & "]",$creation_time_date)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.endtime[" & $job_name & "," & $vm_name & "]",$end_time_date)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.duration[" & $job_name & "," & $vm_name & "]",$duration)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.datediff[" & $job_name & "," & $vm_name & "]",$datediff)

		_logmsg($LogFile,"-> Job: " & $job_name_original & " VM: " & $vm_name_original & " Status: " & $status,true,true)
		$Recordset.MoveNext()
	WEnd

	_logmsg($LogFile,"",true,true)
EndFunc

Func TapeStatusData($Recordset)
	_logmsg($LogFile,"",true,true)
	Local $TapeFileLog = ""
	Local $TapeVmLog = ""

	While Not $Recordset.EOF
		Local $job_id = $Recordset.Fields("job_id").Value
		Local $job_name_original = $Recordset.Fields("job_name").Value
		Local $job_name = StringReplace($job_name_original,",","_")
		Local $job_type = $Recordset.Fields("job_type").Value
		Local $base_job_type = $Recordset.Fields("base_job_type").Value
		Local $parent_job_id = $Recordset.Fields("parent_job_id").Value
		Local $job_result = $Recordset.Fields("job_result").Value
		Local $job_state = $Recordset.Fields("job_state").Value
		Local $job_reason = $Recordset.Fields("job_reason").Value
		Local $creation_time = $Recordset.Fields("creation_time").Value
		Local $end_time = $Recordset.Fields("end_time").Value
		Local $schedule_enabled = $Recordset.Fields("schedule_enabled").Value
		Local $is_job_deleted = $Recordset.Fields("is_job_deleted").Value
		Local $job_options_runmanually = $Recordset.Fields("job_options_runmanually").Value
		Local $MonitorEnabled = _IsJobMonitorEnabled($schedule_enabled, $is_job_deleted, $job_options_runmanually)
		$MonitorEnabled = _ApplyBlacklistMonitorEnabled($MonitorEnabled, $job_name)
		Local $creation_time_date = _DateVeeamFormat($creation_time)
		Local $end_time_date = _DateVeeamFormat($end_time)
		Local $duration = _ComputeDurationMinutes($creation_time_date, $end_time_date)
		Local $datediff = _ComputeDateDiffDays($creation_time_date)
		Local $next_run_value = _ComputeNextRunValue($sDriver, $job_id, $creation_time_date)

		If $Debug > 0 Then
			_logmsg($LogFile,"Tape Status Raw -> JobName: " & $job_name_original & " | Type: " & $job_type & _
					" | Result: " & $job_result & " | State: " & $job_state,true,true)
			_logmsg($LogFile,"Tape Status Raw -> Creation: " & $creation_time & " | End: " & $end_time,true,true)
			_logmsg($LogFile,"Tape Status Calc -> CreationFmt: " & $creation_time_date & " | EndFmt: " & $end_time_date & _
					" | Duration: " & $duration & " | DateDiff: " & $datediff,true,true)
		EndIf

		Local $prefix = "backup.veeam.customchecks.tape.vm"
		If $job_type = 24 Then
			$prefix = "backup.veeam.customchecks.tape.file"
		EndIf

		If $MonitorEnabled = 1 Then
			If $job_type = 24 Then
				$TapeFileMonitoredCount += 1
			ElseIf $job_type = 28 Then
				$TapeVmMonitoredCount += 1
			EndIf
		EndIf

		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".monitored[" & $job_name & "]",$MonitorEnabled)

		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".result[" & $job_name & "]",$job_result)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".state[" & $job_name & "]",$job_state)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".reason[" & $job_name & "]",$job_reason)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".creationtime[" & $job_name & "]",$creation_time_date)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".endtime[" & $job_name & "]",$end_time_date)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".duration[" & $job_name & "]",$duration)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".datediff[" & $job_name & "]",$datediff)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".next_run_time[" & $job_name & "]",$next_run_value)

		If $job_type = 24 Then
			$TapeFileLog &= "-> Tape File-to-Tape: " & $job_name_original & " Result: " & $job_result & @CRLF
		ElseIf $job_type = 28 Then
			$TapeVmLog &= "-> Tape VM: " & $job_name_original & " Result: " & $job_result & @CRLF
		EndIf
		$Recordset.MoveNext()
	WEnd

	_logmsg($LogFile,"List Tape File-to-Tape jobs found:",true,true)
	If $TapeFileLog <> "" Then
		_logmsg($LogFile,StringTrimRight($TapeFileLog, 2),true,true)
	EndIf
	_logmsg($LogFile,"",true,true)

	_logmsg($LogFile,"List Tape VM jobs found:",true,true)
	If $TapeVmLog <> "" Then
		_logmsg($LogFile,StringTrimRight($TapeVmLog, 2),true,true)
	EndIf
	_logmsg($LogFile,"",true,true)
EndFunc

Func VmTaskDefaultData($Recordset)
	_logmsg($LogFile,"",true,true)
	_logmsg($LogFile,"List VMs by job found (default values):",true,true)
	Local $seen_pairs = "|"

	While Not $Recordset.EOF
		Local $job_name_original = $Recordset.Fields("job_name").Value
		Local $job_name = StringReplace($job_name_original,",","_")
		Local $vm_name_original = $Recordset.Fields("vm_name").Value
		Local $vm_name = StringReplace($vm_name_original,",","_")
		Local $pair_key = $job_name & "|" & $vm_name

		If StringInStr($seen_pairs, "|" & $pair_key & "|") = 0 Then
			$seen_pairs &= $pair_key & "|"
			$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.status[" & $job_name & "," & $vm_name & "]",-1)
			$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.reason[" & $job_name & "," & $vm_name & "]","No session data")
			$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.creationtime[" & $job_name & "," & $vm_name & "]",-1)
			$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.endtime[" & $job_name & "," & $vm_name & "]",-1)
			$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.duration[" & $job_name & "," & $vm_name & "]",-1)
			$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.datediff[" & $job_name & "," & $vm_name & "]",-1)
			Local $vmEnabledDefault = _ApplyBlacklistMonitorEnabled(0, $job_name, $vm_name)
			$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.vm.monitored[" & $job_name & "," & $vm_name & "]",$vmEnabledDefault)
			_logmsg($LogFile,"-> Job: " & $job_name_original & " VM: " & $vm_name_original & " Status: -1 (default)",true,true)
		EndIf

		$Recordset.MoveNext()
	WEnd

	_logmsg($LogFile,"",true,true)
EndFunc

Func BackupSyncStatusData($Recordset)
	_logmsg($LogFile,"",true,true)
	_logmsg($LogFile,"List BackupSync jobs found:",true,true)

	While Not $Recordset.EOF
		Local $job_id = $Recordset.Fields("job_id").Value
		Local $job_name_original = $Recordset.Fields("job_name").Value
		Local $job_name = StringReplace($job_name_original,",","_")
		Local $job_result = $Recordset.Fields("job_result").Value
		Local $job_state = $Recordset.Fields("job_state").Value
		Local $job_reason = $Recordset.Fields("job_reason").Value
		Local $creation_time = $Recordset.Fields("creation_time").Value
		Local $end_time = $Recordset.Fields("end_time").Value
		Local $creation_time_date = _DateVeeamFormat($creation_time)
		Local $end_time_date = _DateVeeamFormat($end_time)
		Local $duration = _ComputeDurationMinutes($creation_time_date, $end_time_date)
		Local $datediff = _ComputeDateDiffDays($creation_time_date)
		Local $next_run_value = _ComputeNextRunValue($sDriver, $job_id, $creation_time_date)

		If $Debug > 0 Then
			_logmsg($LogFile,"BackupSync Status Raw -> JobName: " & $job_name_original & _
					" | Result: " & $job_result & " | State: " & $job_state,true,true)
			_logmsg($LogFile,"BackupSync Status Raw -> Creation: " & $creation_time & " | End: " & $end_time,true,true)
			_logmsg($LogFile,"BackupSync Status Calc -> CreationFmt: " & $creation_time_date & " | EndFmt: " & $end_time_date & _
					" | Duration: " & $duration & " | DateDiff: " & $datediff,true,true)
		EndIf

		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupsync.result[" & $job_name & "]",$job_result)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupsync.state[" & $job_name & "]",$job_state)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupsync.reason[" & $job_name & "]",$job_reason)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupsync.creationtime[" & $job_name & "]",$creation_time_date)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupsync.endtime[" & $job_name & "]",$end_time_date)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupsync.duration[" & $job_name & "]",$duration)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupsync.datediff[" & $job_name & "]",$datediff)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,"backup.veeam.customchecks.backupsync.next_run_time[" & $job_name & "]",$next_run_value)

		_logmsg($LogFile,"-> BackupSync: " & $job_name_original & " Result: " & $job_result,true,true)
		$Recordset.MoveNext()
	WEnd

	_logmsg($LogFile,"",true,true)
EndFunc

Func EndpointStatusData($Recordset)
	_logmsg($LogFile,"",true,true)

	Local $EndpointLog = ""
	Local $AgentPolicyLog = ""
	Local $AgentBackupLog = ""

	While Not $Recordset.EOF
		Local $job_id = $Recordset.Fields("job_id").Value
		Local $job_name_original = $Recordset.Fields("job_name").Value
		Local $job_name = StringReplace($job_name_original,",","_")
		Local $job_type = $Recordset.Fields("job_type").Value
		Local $base_job_type = $Recordset.Fields("base_job_type").Value
		Local $parent_job_id = $Recordset.Fields("parent_job_id").Value
		Local $job_result = $Recordset.Fields("job_result").Value
		Local $job_state = $Recordset.Fields("job_state").Value
		Local $job_reason = $Recordset.Fields("job_reason").Value
		Local $creation_time = $Recordset.Fields("creation_time").Value
		Local $end_time = $Recordset.Fields("end_time").Value
		Local $creation_time_date = _DateVeeamFormat($creation_time)
		Local $end_time_date = _DateVeeamFormat($end_time)
		Local $duration = _ComputeDurationMinutes($creation_time_date, $end_time_date)
		Local $datediff = _ComputeDateDiffDays($creation_time_date)
		Local $next_run_value = _ComputeNextRunValue($sDriver, $job_id, $creation_time_date)

		If $Debug > 0 Then
			_logmsg($LogFile,"Endpoint/Agent Status Raw -> JobName: " & $job_name_original & _
					" | SessionType: " & $job_type & " | BaseType: " & $base_job_type & _
					" | ParentJobId: " & $parent_job_id & " | Result: " & $job_result & " | State: " & $job_state,true,true)
			_logmsg($LogFile,"Endpoint/Agent Status Raw -> Creation: " & $creation_time & " | End: " & $end_time,true,true)
			_logmsg($LogFile,"Endpoint/Agent Status Calc -> CreationFmt: " & $creation_time_date & " | EndFmt: " & $end_time_date & _
					" | Duration: " & $duration & " | DateDiff: " & $datediff,true,true)
		EndIf

		Local $prefix = "backup.veeam.customchecks.endpoint"
		If $base_job_type = 4000 Or $base_job_type = 12000 Or $base_job_type = 12002 Or $base_job_type = 12003 Then
			$prefix = _ResolveAgentMetricPrefix($base_job_type, $parent_job_id)
		EndIf

		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".result[" & $job_name & "]",$job_result)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".state[" & $job_name & "]",$job_state)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".reason[" & $job_name & "]",$job_reason)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".creationtime[" & $job_name & "]",$creation_time_date)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".endtime[" & $job_name & "]",$end_time_date)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".duration[" & $job_name & "]",$duration)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".datediff[" & $job_name & "]",$datediff)
		$Zabbix_Items = _add_item_zabbix($Zabbix_Items,$prefix & ".next_run_time[" & $job_name & "]",$next_run_value)

		If $prefix = "backup.veeam.customchecks.endpoint" Then
			$EndpointLog &= "-> Endpoint: " & $job_name_original & " Type: " & $job_type & " Result: " & $job_result & @CRLF
		ElseIf $prefix = "backup.veeam.customchecks.agent.policy" Then
			$AgentPolicyLog &= "-> Agent Policy: " & $job_name_original & " Type: " & $job_type & " Result: " & $job_result & @CRLF
		ElseIf $prefix = "backup.veeam.customchecks.agent.backup" Then
			$AgentBackupLog &= "-> Agent Backup: " & $job_name_original & " Type: " & $job_type & " Result: " & $job_result & @CRLF
		EndIf

		$Recordset.MoveNext()
	WEnd

	_logmsg($LogFile,"List Endpoint jobs found:",true,true)
	If $EndpointLog <> "" Then
		_logmsg($LogFile,StringTrimRight($EndpointLog, 2),true,true)
	EndIf
	_logmsg($LogFile,"",true,true)

	_logmsg($LogFile,"List Agent Policy jobs found:",true,true)
	If $AgentPolicyLog <> "" Then
		_logmsg($LogFile,StringTrimRight($AgentPolicyLog, 2),true,true)
	EndIf
	_logmsg($LogFile,"",true,true)

	_logmsg($LogFile,"List Agent Backup jobs found:",true,true)
	If $AgentBackupLog <> "" Then
		_logmsg($LogFile,StringTrimRight($AgentBackupLog, 2),true,true)
	EndIf

	_logmsg($LogFile,"",true,true)
EndFunc

Func GetDaysInMonth($year, $month)
    Local $days = Int(@MON[$month])
    If $month = 2 And Mod($year, 4) = 0 And (Mod($year, 100) <> 0 Or Mod($year, 400) = 0) Then
        $days = 29
    EndIf
    Return $days
EndFunc

Func IsValidDate($year, $month, $day)
    Return $day <= GetDaysInMonth($year, $month)
EndFunc

; Convert Veeam data to usable format
Func _DateVeeamFormat($date_string)
    $date = -1

	$date_string = StringStripWS($date_string,$STR_STRIPSPACES)

	; Format YYYY-MM-DD HH:MM:SS(.ffffff)
	if StringRegExp($date_string, "^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}(\.\d+)?$") Then
		local $parts = StringSplit($date_string, " ")
		if $parts[0] = 2 Then
			local $date_part = StringReplace($parts[1], "-", "/")
			local $time_part = $parts[2]
			; strip fractional seconds if present
			if StringInStr($time_part, ".") Then
				$time_part = StringLeft($time_part, StringInStr($time_part, ".") - 1)
			EndIf
			$date = $date_part & " " & $time_part
			Return $date
		EndIf
	EndIf

    ; Format AM/PM
    If StringInStr($date_string, "AM") Or StringInStr($date_string, "PM") Then
		local $parts = StringSplit($date_string, " ")

		local $month_str = $parts[1]
		local $day = $parts[2]
		local $year = $parts[3]
		local $time_str = $parts[4]

		local $months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
		local $month = 0
		For $i = 0 To 11
			If $months[$i] = $month_str Then
				$month = $i + 1
				ExitLoop
			EndIf
		Next

		local $time_parts = StringSplit($time_str, ":")
		local $hour = $time_parts[1]
		local $min = $time_parts[2]
		local $ampm = StringRight($time_str, 2)

		$hour = Int($hour)

		If $ampm == "PM" Then
			If $hour <> 12 Then
				$hour = $hour + 12
			EndIf
		ElseIf $ampm == "AM" Then
			If $hour == 12 Then
				$hour = 0
			EndIf
		EndIf

		$date = $year & "/" & StringFormat("%02d", $month) & "/" & StringFormat("%02d", $day) & " " & StringFormat("%02d", $hour) & ":" & StringFormat("%02d", $min) & ":00"

    ElseIf StringLen($date_string) = 14 Then
		; parsig format 20241125100000
        local $year = StringLeft($date_string,4)
		local $month = StringMid($date_string,5,2)
		local $day = StringMid($date_string,7,2)
		local $hour = StringMid($date_string,9,2)
		local $min = StringMid($date_string,11,2)
		local $sec = StringMid($date_string,13,2)
		$date = $year & "/" & $month & "/" & $day & " " & $hour & ":" & $min & ":" & $sec
    EndIf

    Return $date
EndFunc
Func MapMonthToNumber($monthName)
    Local $monthsNames[12] = ["January", "February", "March", "April", "May", "June", _
                               "July", "August", "September", "October", "November", "December"]
    For $i = 0 To UBound($monthsNames) - 1
        If $monthsNames[$i] = $monthName Then
            Return $i + 1
        EndIf
    Next
    Return -1
EndFunc

Func MapDayToNumber($dayName)
    Local $daysNames[12] = ["Sunday" , "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    For $i = 0 To UBound($daysNames) - 1
        If $daysNames[$i] = $dayName Then
            Return $i + 1
        EndIf
    Next
    Return -1
EndFunc

Func ParseMonthsFromXML($xml)
    Local $oXML = ObjCreate("Microsoft.XMLDOM")
    Local $monthsList = []
    If $oXML.loadXML($xml) Then
        Local $nodes = $oXML.SelectNodes("//ScheduleOptions/OptionsMonthly/Months/EMonth")
        For $node In $nodes
            _ArrayAdd($monthsList, $node.text)
        Next
	Else
		ConsoleWrite(@CRLF & "Errore: " & $oXML.parseError.reason & @CRLF & @CRLF & $xml & @CRLF & @CRLF)
    EndIf
    Return $monthsList
EndFunc

Func CalculateNextBackupDate($lastBackupDate, $type, $list)

	Local $nextBackupDate = $lastBackupDate
	Local $MaxTries = 1

	If $type ="M" then
		$MaxTries = 12
	EndIf
	If $type ="D" then
		$MaxTries = 7
	EndIf

    Local $temporalNames[1] = [0]
    For $i = 1 To $list[0]
		If $type ="M" then
			Local $Num = MapMonthToNumber($list[$i])
		EndIf
		If $type ="D" then
			Local $Num = MapDayToNumber($list[$i])
		EndIf

        If $Num > 0 Then
            _ArrayAdd($temporalNames, $Num)
        EndIf
    Next

	$temporalNames[0] = UBound($temporalNames) - 1

    _ArraySort($temporalNames,0,1)

    if UBound($temporalNames) = 0 then
		return $nextBackupDate
	endif

	$count = 0

    Local $found = False
    While Not $found
		$count += 1
		if $count > $MaxTries then
			$found = True

			Local $nextBackupDate = $lastBackupDate
			$nextBackupDate = _DateAdd($type, 1, $nextBackupDate)
			Local $datearray = StringSplit($nextBackupDate,"/")
			Local $nextName
			If $type = "M" then
				$nextName = Number($datearray[2])
			EndIf
			If $type = "D" then
				$nextName = _DateToDayOfWeek($datearray[1],$datearray[2],$datearray[3])
			EndIf

			_logmsg($LogFile,"Search " & $type & " failed. Exit with first result",true,true)
			_logmsg($LogFile,"BackupDate: " & $lastBackupDate & " NextBackupDate: " & $nextBackupDate & " (Tries: " & $count & ")",true,true)
			ExitLoop
		endif

        $nextBackupDate = _DateAdd($type, 1, $nextBackupDate)

		Local $datearray = StringSplit($nextBackupDate,"/")
        Local $nextName
		If $type = "M" then
			$nextName = Number($datearray[2])
		EndIf
		If $type = "D" then
			$nextName = _DateToDayOfWeek($datearray[1],$datearray[2],$datearray[3])
		EndIf

        For $i = 1 To $temporalNames[0]
            If $temporalNames[$i] = $nextName Then
                $found = True
                ExitLoop
            EndIf
        Next

    WEnd

    Return $nextBackupDate
EndFunc

; Extract data from xml
Func _XMLExtractValue($xml,$search)
	local $XMLValue = 0

	Local $oXML = ObjCreate("Microsoft.XMLDOM")
	If $oXML.loadXML($xml) Then
		if $Debug = 1 then
			_logmsg($LogFile,"XML Loaded",true,true)
			_logmsg($LogFile,"XML: " & $oXML.xml,true,true)
		EndIf

		Local $XMLNode = $oXML.SelectSingleNode($search)

		if IsObj($XMLNode) Then
			if $Debug = 1 then
				_logmsg($LogFile,"$XMLValue: " & $XMLNode.text,true,true)
			EndIf
			$XMLValue = $XMLNode.text
		else
			if $Debug = 1 then
				_logmsg($LogFile,"XML Node Not Obj ",true,true)
			EndIf
		EndIf

	else
		if $Debug = 1 then
			_logmsg($LogFile,"XML Not Loaded",true,true)
		EndIf
	EndIF

	return $XMLValue
EndFunc

; Joint items for zabbix
func _add_item_zabbix($zabbix_items,$key,$value)
	if $zabbix_items <> "" Then
		$zabbix_items &= @CRLF
	endif

	$value = StringReplace($value,@CRLF," ")
	$value = StringReplace($value,@CR," ")
	$value = StringReplace($value,@LF," ")
	;$value = RemoveControlChars($value)
	$zabbix_items = $zabbix_items & " - " & chr(34) & $key & chr(34) & " " & chr(34) & $value & chr(34)

	return $zabbix_items
endfunc

; Function to remove control characters from a string
Func RemoveControlChars($text)
    ; Define the control characters to remove
    Local $controlChars = ["\0", "\a", "\b", "\t", "\v", "\f", "\r", "\n"]
    ; Loop through each control character and replace it with an empty string
    For $i = 0 To UBound($controlChars) - 1
        $text = StringReplace($text, $controlChars[$i], "")
    Next
    Return $text
EndFunc

; Log to file/console
func _logmsg($logfile,$msg,$file = false,$console = true)
	if $file = true then
		_FileWriteLog($logfile,$msg & @CRLF)
	endif

	if $console = true Then
		ConsoleWrite($msg & @CRLF)
	endif
EndFunc
#EndRegion Functions

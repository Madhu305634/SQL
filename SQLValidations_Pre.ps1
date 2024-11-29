[CmdletBinding()]
param (
	[object]$EmailList
	    )

#$EmailList = 'prashanthkumar.runja@epiqglobal.com'

$HostServer = "P054SQLMGMT03\SQLADMIN"
$HostDB = "DBASUPPORT"

$ServerList = "select name from msdb.dbo.sysmanagement_shared_registered_servers_internal where server_group_id in('20','23','1238') and (name not like '%SOLARWINDS' and name not like '%CIC' and name not like 'HKG%') order by name"
$Servers = Invoke-Sqlcmd -ServerInstance $HostServer -Query $ServerList | Select-Object name -ExpandProperty name | Sort-Object name
$Servers = $Servers.name
#$Servers="S061ESRSQLS01.amer.EPIQCORP.COM\EPIQDIRECTORY","P054SQLMGMT03\SQLADMIN"

$PackageStartTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), 'India Standard Time') 
$SvcStatGrid = @() 
$DBStatGrid=@()
$DBCountStatGrid=@()
$PatchStatGrid=@()
$PreferredNodeStatGrid=@()
$UptimeStatGrid=@()
$CohAgtGrid = @()

$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #9FD4DC;text-align: left;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@

$path_Pre = "D:\PowerShell\SQLValidations\Pre"

#Delete files
Get-ChildItem -Path $path_Pre -Include *.* -File -Recurse | foreach { $_.Delete()}

#FilesCopy
$OutFileLog = "$path_Pre\AMER_SQLValidations_Pre.xlsx" 


Foreach($SQLInstance in $Servers)
{
Write-Host $SQLInstance

$instance = $SQLInstance
#$computer= $instance -replace "\*",""
if($instance.IndexOf("\") -gt 0) {
        $computer = $instance.Substring(0, $instance.IndexOf("\"))
} else {
        $computer = $instance
}


#Databases status
    $DbQuery= "SELECT @@SERVERNAME AS [ServerName], NAME AS [DatabaseName], DATABASEPROPERTYEX(NAME, 'Status') AS [Status] FROM dbo.sysdatabases ORDER BY NAME ASC"
    $DBStatGrid += Invoke-Sqlcmd -ServerInstance $instance -Query $DbQuery | select ServerName, DatabaseName,Status 

#Databases Count
    $DbCountQuery= "SELECT @@servername As ServerName,count(name) As DBCount from sys.databases"
    $DBCountStatGrid += Invoke-Sqlcmd -ServerInstance $instance -Query $DbCountQuery | select ServerName,DBCount 

#SQL Patch
    $PatchQuery= @"
    SELECT
                @@SERVERNAME as SERVERNAME,
                CASE
                WHEN CONVERT(VARCHAR(128), SERVERPROPERTY ('productversion')) like '8%' THEN 'SQL2000'
                WHEN CONVERT(VARCHAR(128), SERVERPROPERTY ('productversion')) like '9%' THEN 'SQL2005'
                WHEN CONVERT(VARCHAR(128), SERVERPROPERTY ('productversion')) like '10.0%' THEN 'SQL2008'
                WHEN CONVERT(VARCHAR(128), SERVERPROPERTY ('productversion')) like '10.5%' THEN 'SQL2008 R2'
                WHEN CONVERT(VARCHAR(128), SERVERPROPERTY ('productversion')) like '11%' THEN 'SQL2012'
                WHEN CONVERT(VARCHAR(128), SERVERPROPERTY ('productversion')) like '12%' THEN 'SQL2014'
                WHEN CONVERT(VARCHAR(128), SERVERPROPERTY ('productversion')) like '13%' THEN 'SQL2016'
                WHEN CONVERT(VARCHAR(128), SERVERPROPERTY ('productversion')) like '14%' THEN 'SQL2017'
                WHEN CONVERT(VARCHAR(128), SERVERPROPERTY ('productversion')) like '15%' THEN 'SQL2019'
                WHEN CONVERT(VARCHAR(128), SERVERPROPERTY ('productversion')) like '16%' THEN 'SQL2022'
                ELSE 'unknown'
                END AS Version, SERVERPROPERTY('ProductLevel') AS ProductLevel, SERVERPROPERTY('Edition') AS Edition, SERVERPROPERTY('ProductVersion') AS ProductVersion, SERVERPROPERTY('IsClustered') AS IsClustered, SERVERPROPERTY('ProductUpdateLevel') AS CurrentCU
"@
         $PatchStatGrid += Invoke-Sqlcmd -ServerInstance $instance -Query $PatchQuery | select SERVERNAME,Version,ProductLevel,Edition,ProductVersion,CurrentCU,IsClustered
             
#SQL Services 
    
   #$SvcStatGrid += Get-Service -computer $computer -Name $Name | Where-Object{$_.DisplayName -like "SQL*"} | select MachineName, Name, DisplayName, Status
    $ServicesQuery= "DECLARE @WINSCCMD TABLE (ID INT IDENTITY (1,1) PRIMARY KEY NOT NULL, Line VARCHAR(MAX))
                INSERT INTO @WINSCCMD(Line) EXEC master.dbo.xp_cmdshell 'sc queryex type= service state= all'
 
            SELECT  SERVERPROPERTY ( 'Machinename' ) As Machinename  
            ,LTRIM (SUBSTRING (W1.Line, 15, 100)) AS ServiceName
            , LTRIM (SUBSTRING (W2.Line, 15, 100)) AS DisplayName
            , LTRIM (SUBSTRING (W3.Line, 33, 100)) AS ServiceState
	            FROM @WINSCCMD W1, @WINSCCMD W2, @WINSCCMD W3
            WHERE W1.ID = W2.ID - 1 AND
            W3.ID - 3 = W1.ID AND
            LTRIM (SUBSTRING (W1.Line, 15, 100)) is not null AND
            LTRIM (SUBSTRING (W2.Line, 15, 100))  like 'SQL%';"
    $SvcStatGrid += Invoke-Sqlcmd -ServerInstance $instance -Query $ServicesQuery | select Machinename, ServiceName,DisplayName,ServiceState


#SQL preferred Node
    $NodeQuery= "SELECT
                @@SERVERNAME As Servername,
                [ClusterName] = SUBSTRING(@@SERVERNAME,0,CHARINDEX('\',@@SERVERNAME)),
                [Nodes] = NodeName,
                [IsActiveNode] = CASE WHEN NodeName = SERVERPROPERTY('ComputerNamePhysicalNetBIOS') THEN '1' ELSE '0' END
                FROM sys.dm_os_cluster_nodes
                WHERE SERVERPROPERTY('ComputerNamePhysicalNetBIOS') <> SUBSTRING(@@SERVERNAME,0,CHARINDEX('\',@@SERVERNAME))
                AND @@SERVERNAME <> SERVERPROPERTY('ComputerNamePhysicalNetBIOS');"
    $PreferredNodeStatGrid += Invoke-Sqlcmd -ServerInstance $instance -Query $NodeQuery | select Servername, ClusterName,Nodes,IsActiveNode 

#Server Uptime
    $ServeruptimeQuery= @"
    SELECT 
           SERVERPROPERTY('ComputerNamePhysicalNetBIOS') AS 'Current_NodeName',
           [OSStartTime]   = convert(varchar(23),b.OS_Start,121),
           [SQLServerStartTime]   = convert(varchar(23),a.SQL_Start,121),
           [SQLAgentStartTime]   = convert(varchar(23),a.Agent_Start,121),
           [OSUptime] = convert(varchar(15), right(10000000+datediff(dd,0,getdate()-b.OS_Start),4)+' '+ convert(varchar(20),getdate()-b.OS_Start,108)),
           [SQLUptime] = convert(varchar(15), right(10000000+datediff(dd,0,getdate()-a.SQL_Start),4)+' '+ convert(varchar(20),getdate()-a.SQL_Start,108)) ,
           [AgentUptime] = convert(varchar(15), right(10000000+datediff(dd,0,getdate()-a.Agent_Start),4)+' '+ convert(varchar(20),getdate()-a.Agent_Start,108))
            from
           (
           Select SQL_Start = min(aa.login_time),Agent_Start = nullif(min(case when aa.program_name like 'SQLAgent %' then aa.login_time else '99990101' end), convert(datetime,'99990101'))
           from  master.dbo.sysprocesses aa
           where aa.login_time > '20000101') a
           cross join
           (
           select OS_Start = dateadd(ss,bb.[ms_ticks]/-1000,getdate())
       from sys.[dm_os_sys_info] bb) b
"@
         $UptimeStatGrid += Invoke-Sqlcmd -ServerInstance $instance -Query $ServeruptimeQuery | select Current_NodeName,OSStartTime,SQLServerStartTime,SQLAgentStartTime,OSUptime,SQLUptime,AgentUptime   

#Cohesity Agent Services 
    
   $CohQuery= "DECLARE @WINSCCMD TABLE (ID INT IDENTITY (1,1) PRIMARY KEY NOT NULL, Line VARCHAR(MAX))
                INSERT INTO @WINSCCMD(Line) EXEC master.dbo.xp_cmdshell 'sc queryex type= service state= all'
 
                SELECT  SERVERPROPERTY ( 'Machinename' ) As Machinename  
                ,RTRIM(LTRIM (SUBSTRING (W1.Line, 15, 100))) AS ServiceName
                , RTRIM(LTRIM (SUBSTRING (W2.Line, 15, 100))) AS DisplayName
                , RTRIM(LTRIM (SUBSTRING (W3.Line, 33, 100))) AS ServiceState
                FROM @WINSCCMD W1, @WINSCCMD W2, @WINSCCMD W3
                WHERE W1.ID = W2.ID - 1 AND
                W3.ID - 3 = W1.ID AND
                RTRIM(LTRIM (SUBSTRING (W1.Line, 15, 100))) is not null AND
                --(RTRIM(LTRIM (SUBSTRING (W2.Line, 15, 100)))  like 'SQL%' OR 
                RTRIM(LTRIM (SUBSTRING (W2.Line, 15, 100)))  like 'Cohesity%';"
    $CohAgtGrid += Invoke-Sqlcmd -ServerInstance $instance -Query $CohQuery | select Machinename, ServiceName,DisplayName,ServiceState
} 

$DBStatGrid | Export-Csv -Path "$path_Pre\DBStatusPre.csv"
$DBCountStatGrid | Export-Csv -Path "$path_Pre\DBCountPre.csv"
$PatchStatGrid | Export-Csv -Path "$path_Pre\SQLPatchPre.csv"
$SvcStatGrid | Export-Csv -Path "$path_Pre\ServicesPre.csv"
$PreferredNodeStatGrid | Export-Csv -Path "$path_Pre\PreferredNodePre.csv"
$UptimeStatGrid | Export-Csv -Path "$path_Pre\SQLUptimePre.csv"
$CohAgtGrid | Export-Csv -Path "$path_Pre\COHServicePre.csv"

$DBStatGrid | Export-Excel -Path $OutFileLog -AutoSize -TableName DBStatus -WorksheetName DBStatus
$DBCountStatGrid | Export-Excel -Path $OutFileLog -AutoSize -TableName DBCount -WorksheetName DBCount
$PatchStatGrid | Export-Excel -Path $OutFileLog -AutoSize -TableName SQLPatch -WorksheetName SQLPatch
$SvcStatGrid | Export-Excel -Path $OutFileLog -AutoSize -TableName Services -WorksheetName Services
$PreferredNodeStatGrid | Export-Excel -Path $OutFileLog -AutoSize -TableName PreferredNode -WorksheetName PreferredNode
$UptimeStatGrid | Export-Excel -Path $OutFileLog -AutoSize -TableName ServerUptime -WorksheetName ServerUptime
$CohAgtGrid | Export-Excel -Path $OutFileLog -AutoSize -TableName Cohesity -WorksheetName Cohesity

Send-MailMessage -Body "Please find the attached spread sheet for LVDC Amer - SQLValidatons</br></br>" -To $EmailList -From DL-SQLDatabaseSupportL1@epiqglobal.com -SmtpServer mailrelay.amer.epiqcorp.com -Subject "SQLValidations:LVDC-Amer - $PackageStartTime"  -BodyAsHtml -Attachments $OutFileLog

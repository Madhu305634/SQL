[CmdletBinding()]
param (
	[object]$EmailList
	    )

$HostServer = "P054SQLMGMT03\SQLADMIN"
$HostDB = "DBASUPPORT"

$ServerList = "select name from msdb.dbo.sysmanagement_shared_registered_servers_internal where server_group_id in('20','23','1238') and (name not like '%SOLARWINDS' and name not like '%CIC' and name not like 'HKG%') order by name"
$Servers = Invoke-Sqlcmd -ServerInstance $HostServer -Query $ServerList | Select-Object name -ExpandProperty name | Sort-Object name
$Servers = $Servers.name

$PackageStartTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId((Get-Date), 'India Standard Time') 
$SvcStatGrid = @() 
$DBStatGrid=@()
$DBCountStatGrid=@()
$PatchStatGrid=@()
$preferredNodeStatGrid=@()
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
$path_Post = "D:\PowerShell\SQLValidations\Post"

#Delete files
Get-ChildItem -Path $path_Post -Include *.* -File -Recurse | foreach { $_.Delete()}

#FilesCopy
$OutFileLog_Pre = "$path_Pre\AMER_SQLValidations_Pre.xlsx" 
$OutFileLog_Post = "$path_Post\AMER_SQLValidations_Post.xlsx" 

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
    $DBStatGrid += Invoke-Sqlcmd -ServerInstance $instance -Query $DbQuery -QueryTimeout 0 | select ServerName, DatabaseName,Status 

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
         $PatchStatGrid += Invoke-Sqlcmd -ServerInstance $instance -Query $PatchQuery -QueryTimeout 0 | select SERVERNAME,Version,ProductLevel,Edition,ProductVersion,CurrentCU,IsClustered
             
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
    $SvcStatGrid += Invoke-Sqlcmd -ServerInstance $instance -Query $ServicesQuery -QueryTimeout 0 | select Machinename, ServiceName,DisplayName,ServiceState


#SQL Preferred Node
    $NodeQuery= "SELECT
                @@SERVERNAME As Servername,
                [ClusterName] = SUBSTRING(@@SERVERNAME,0,CHARINDEX('\',@@SERVERNAME)),
                [Nodes] = NodeName,
                [IsActiveNode] = CASE WHEN NodeName = SERVERPROPERTY('ComputerNamePhysicalNetBIOS') THEN '1' ELSE '0' END
                FROM sys.dm_os_cluster_nodes
                WHERE SERVERPROPERTY('ComputerNamePhysicalNetBIOS') <> SUBSTRING(@@SERVERNAME,0,CHARINDEX('\',@@SERVERNAME))
                AND @@SERVERNAME <> SERVERPROPERTY('ComputerNamePhysicalNetBIOS');"
    $PreferredNodeStatGrid += Invoke-Sqlcmd -ServerInstance $instance -Query $NodeQuery -QueryTimeout 0 | select Servername, ClusterName,Nodes,IsActiveNode 

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
         $UptimeStatGrid += Invoke-Sqlcmd -ServerInstance $instance -Query $ServeruptimeQuery -QueryTimeout 0 | select Current_NodeName,OSStartTime,SQLServerStartTime,SQLAgentStartTime,OSUptime,SQLUptime,AgentUptime   

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
    $CohAgtGrid += Invoke-Sqlcmd -ServerInstance $instance -Query $CohQuery -QueryTimeout 0 | select Machinename, ServiceName,DisplayName,ServiceState
} 

$DBStatGrid | Export-Csv -Path "$path_Post\DBStatusPost.csv"
$DBCountStatGrid | Export-Csv -Path "$path_Post\DBCountPost.csv"
$PatchStatGrid | Export-Csv -Path "$path_Post\SQLPatchPost.csv"
$SvcStatGrid | Export-Csv -Path "$path_Post\ServicesPost.csv"
$preferredNodeStatGrid | Export-Csv -Path "$path_Post\PreferredNodePost.csv"
$UptimeStatGrid | Export-Csv -Path "$path_Post\SQLUptimePost.csv"
$CohAgtGrid | Export-Csv -Path "$path_Post\COHServicePost.csv"

$DBStatGrid | Export-Excel -Path $OutFileLog_Post -AutoSize -TableName DBStatus -WorksheetName DBStatus
$DBCountStatGrid | Export-Excel -Path $OutFileLog_Post -AutoSize -TableName DBCount -WorksheetName DBCount
$PatchStatGrid | Export-Excel -Path $OutFileLog_Post -AutoSize -TableName SQLPatch -WorksheetName SQLPatch
$SvcStatGrid | Export-Excel -Path $OutFileLog_Post -AutoSize -TableName Services -WorksheetName Services
$preferredNodeStatGrid | Export-Excel -Path $OutFileLog_Post -AutoSize -TableName PreferredNode -WorksheetName PreferredNode
$UptimeStatGrid | Export-Excel -Path $OutFileLog_Post -AutoSize -TableName ServerUptime -WorksheetName ServerUptime
$CohAgtGrid | Export-Excel -Path $OutFileLog_Post -AutoSize -TableName Cohesity -WorksheetName Cohesity

#Differences comparision
#DB Count
$CountPre = Import-Csv "$path_Pre\DBCountPre.csv" 
$countPost = Import-Csv "$path_Post\DBCountPost.csv"

$Dbcomp =Compare-Object $countPost $CountPre  -Property ServerName, DBCount | Select-Object ServerName,DBCount | Group-Object ServerName | Select @{name=”ServerName”;expression={$_.name}},@{Name="PreValue";Expression={$_.Group[0].DBCount}},@{Name="PostValue";Expression={$_.Group[1].DBCount}} | Sort-Object ServerName | ConvertTo-Html -Head $Header
$Dbcomp

#Database Status
$StatusPre = Import-Csv "$path_Pre\DBStatusPre.csv" 
$StatusPost = Import-Csv "$path_Post\DBStatusPost.csv"

$Statuscomp=Compare-Object $StatusPost $StatusPre -Property ServerName,DatabaseName,Status | Group-Object ServerName,DatabaseName | Select @{Name="ServerName";Expression={$_.Group[0].ServerName}},@{Name="DatabaseName";Expression={$_.Group[0].DatabaseName}},@{Name="PreValue";Expression={$_.Group[0].Status}},@{Name="PostValue";Expression={$_.Group[1].Status}}| Sort-Object ServerName| ConvertTo-Html -Head $Header
$Statuscomp

#Services
$ServicesPre = Import-Csv "$path_Pre\ServicesPre.csv"
$ServicesPost = Import-Csv "$path_Post\ServicesPost.csv"

$Servicescomp =Compare-Object $ServicesPost $ServicesPre -Property Machinename,ServiceName,ServiceState | Group-Object Machinename,ServiceName | Select @{Name="Machinename";Expression={$_.Group[0].Machinename}},@{Name="ServiceName";Expression={$_.Group[0].ServiceName}},@{Name="PreValue";Expression={$_.Group[0].ServiceState}},@{Name="PostValue";Expression={$_.Group[1].ServiceState}} | Sort-Object Machinename| ConvertTo-Html -Head $Header
$Servicescomp

#PatchList
$CUlist = Invoke-Sqlcmd -ServerInstance $HostServer -Database $HostDB -Query "select top 1 with ties * from [SQLBuilds]
                                                                                                            where  SQLSERVER in ('2022','2019','2017','2016','2014','2012')
                                                                                                            order by row_number() over (partition by SQLSERVER order by SQLSERVER desc)" | Select-Object SQLServer,Build,Description,Link,@{name=”ReleaseDate”;expression={$_.ReleaseDate.tostring("MM-dd-yyyy")}} | ConvertTo-Html -Head $Header

#Patch
$RecPatch = Invoke-Sqlcmd -ServerInstance $HostServer -Database $HostDB -Query "select top 1 with ties * from [SQLBuilds]
                                                                                                            --where  SQLSERVER in ('2022','2019','2017','2016','2014','2012')
                                                                                                            order by row_number() over (partition by SQLSERVER order by SQLSERVER desc)"

$PatchPost = Import-Csv "$path_Post\SQLpatchPost.csv" 
$PatchResults=@()
$PatchPost | ForEach-Object {
        if ($_.ProductVersion -notin $RecPatch.Build) 
        {
         $PatchResults +=[PSCustomObject]@{
         ServerName = $_.ServerName 
         Version =$_.Version
         ProductVersion=$_.ProductVersion
         CurrentCU=$_.CurrentCU }
          
        }
        
}
$PatchResults=$PatchResults | ConvertTo-Html -Head $Header

#Preferred Node
$CluNodePre = Import-Csv "$path_Pre\PreferredNodePre.csv" 
$CluNodePost = Import-Csv "$path_Post\PreferredNodePost.csv"
$NodeComp =Compare-Object $CluNodePost $CluNodePre  -Property ServerName,ClusterName, Nodes,IsActiveNode| select  ServerName,ClusterName,Nodes,IsActiveNode  | Group-Object ServerName,ClusterName,Nodes | Select @{name=”ServerName”;expression={$_.Group[0].ServerName}},@{name=”ClusterName”;expression={$_.Group[0].ClusterName}},@{name=”Nodes”;expression={$_.Group[0].Nodes}},@{Name="PreValue";Expression={$_.Group[0].IsActiveNode}},@{Name="PostValue";Expression={$_.Group[1].IsActiveNode}}| ConvertTo-Html -Head $Header
$NodeComp

#OS Reboot time
$UptimePost = Import-Csv "$path_Post\SQLUptimePost.csv"
$today=Get-Date -format "yyyy-MM-dd"
$Yday=(Get-Date).AddDays(-1).ToString('yyyy-MM-dd')
$Nday=(Get-Date).AddDays(+1).ToString('yyyy-MM-dd')
$OSStartTime=$UptimePost | select @{name=”ServerName”;expression={$_.Current_NodeName}},OSStartTime| Where-Object {$_.OSStartTime.ToString() -notmatch $today.ToString() -and $_.OSStartTime.ToString() -notmatch $Yday.ToString() -and $_.OSStartTime.ToString() -notmatch $Nday.ToString()} | Sort-Object ServerName
$OSStartTime=$OSStartTime | ConvertTo-Html -Head $Header

#COHESITY
$COHPre = Import-Csv "$path_Pre\COHServicePre.csv"  
$COHPost = Import-Csv "$path_Post\COHServicePost.csv" 

$COHcomp=Compare-Object $COHPost $COHPre -Property Machinename,ServiceName,ServiceState | Group-Object Machinename,ServiceName | Select @{Name="Machinename";Expression={$_.Group[0].Machinename}},@{Name="ServiceName";Expression={$_.Group[0].ServiceName}},@{Name="PreValue";Expression={$_.Group[0].ServiceState}},@{Name="PostValue";Expression={$_.Group[1].ServiceState}} | Sort-Object Machinename | ConvertTo-Html -Head $Header
$COHcomp

<#
$StatusComp | Export-Excel -Path $OutFileLog_Diff -AutoSize -TableName DBStatus -WorksheetName DBStatus
$DBCountComp | Export-Excel -Path $OutFileLog_Diff -AutoSize -TableName DBCount -WorksheetName DBCount
$PatchComp | Export-Excel -Path $OutFileLog_Diff -AutoSize -TableName SQLPatch -WorksheetName SQLPatch
$ServicesComp | Export-Excel -Path $OutFileLog_Diff -AutoSize -TableName Services -WorksheetName Services
$NodeComp | Export-Excel -Path $OutFileLog_Diff -AutoSize -TableName preferredNode -WorksheetName preferredNode
$UptimeComp | Export-Excel -Path $OutFileLog_Diff -AutoSize -TableName ServerUptime -WorksheetName ServerUptime
#>

Send-MailMessage -Body "Please find the validation report below.Detailed information can be found in the attached sheets</br></br><B> DBcount difference below:</B></br></br>$Dbcomp</br></br><B>Databases which have changed their status are listed below:</B></br></br>$Statuscomp</br></br><B>Services which have changed their status are listed below:</B></br></br>$Servicescomp</br></br><B>Recent patches list below:</B></br></br>$CUlist</br><B>Below servers do not have the latest SQL patch:</B></br></br>$PatchResults</br></br><B>Below machines changed their preferred node:</B></br></br>$NodeComp</br></br><B>Below machines have not rebooted:</B></br></br>$OSStartTime</br></br><B>COH Agent service status below:</B></br></br>$COHcomp</br>"`
 -To $EmailList -From DL-SQLDatabaseSupportL1@epiqglobal.com -SmtpServer mailrelay.amer.epiqcorp.com -Subject "SQLValidations:LVDC-Amer - $PackageStartTime"  -BodyAsHtml -Attachments $OutFileLog_Pre, $OutFileLog_Post

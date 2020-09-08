# ===============
# Load Config.xml
# ===============
# Load XML config file with client settings:
$Path = "C:\script\Reporting\Config.xml"
# load it into an XML object:
$xml = New-Object -TypeName XML
$xml.Load($Path)

# ============================
# Connect to SQL and run QUERY 
# ============================

#Generate the Encrypted Password
#$credential = Get-Credential
#$credential.Password | ConvertFrom-SecureString | Set-Content c:\script\encrypted_password.txt
$EncryptedPassword = $xml.configuration.Settings.SCCM.Password
$UserName = $xml.configuration.Settings.SCCM.Username
$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, ($EncryptedPassword | ConvertTo-SecureString)
$client = $xml.configuration.export.client
$SQLServer = $xml.configuration.Settings.SCCM.Server
$SQLDBName = $xml.configuration.Settings.SCCM.Database
$SQLUsername = $credentials.UserName
$SQLPassword = $Credentials.GetNetworkCredential().password
$OutputPath = "C:\Script\Reporting\Export\" + $client + "SCCM"

# ==========================
# Applications
# ==========================
$Applications = $OutputPath + "Applications"
$SqlQuery = "select  all 

App.CI_ID,
App.DisplayName [Application],
App.Manufacturer,
App.CIVersion [Revision],
App.CreatedBy [Created By],
App.DateCreated [Date Created],
App.DateLastModified [Date Last Modified],
--App.HasContent,
--App.IsBundle,
--App.IsDeployable,
App.IsDeployed [Deployed],
App.IsEnabled [Enabled],
App.IsExpired [Expired],
App.IsHidden [Hidden],
--App.IsLatest,
App.IsQuarantined [Quarantined],
App.IsSuperseded [Superseded],
App.IsSuperseding [Superseding],
App.LastModifiedBy [Last Modified By],
--App.Description,
App.NumberOfDependedDTs [Depended Deployment Types],
App.NumberOfDependentTS [Dependent TS],
App.NumberOfDeployments [Deployments],
App.NumberOfDeploymentTypes [Deployment Types],
App.NumberOfDevicesWithApp [Devices with App],
App.NumberOfDevicesWithFailure [Devices with Failure],
App.NumberOfSettings [Settings],
App.NumberOfUsersWithApp [Users with App],
App.NumberOfUsersWithFailure [Users with Failure],
App.NumberOfUsersWithRequest [Users with Request],
--App.NumberOfVirtualEnvironments,
App.SoftwareVersion from fn_ListLatestApplicationCIs_List(1033) AS App  

where (App.ModelName not  in (select  all Folder##Alias##810314.InstanceKey from vFolderMembers AS Folder##Alias##810314  where Folder##Alias##810314.ObjectTypeName = N'App') AND App.IsHidden = 0 ) order by App.DisplayName"
  
## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $applications'.csv' -NoTypeInformation -Force

# ==========================
# Applications - O365
# ==========================
$ApplicationsO365 = $OutputPath + "ApplicationsO365"
$SqlQuery = "SELECT DISTINCT
  ResourceID,
  GroupID,
  RevisionID,
  AgentID,
  [TimeStamp],
  AutoUpgrade0,
  CCMManaged0,
  CASE
    WHEN CDNBaseUrl0 = 'http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60' THEN 'Monthly Channel'
    WHEN CDNBaseUrl0 = 'http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114' THEN 'Semi-Annual Channel'
    WHEN CDNBaseUrl0 = 'http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be' THEN 'Monthly Channel (Targeted)'
    WHEN CDNBaseUrl0 = 'http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf' THEN 'Semi-Annual Channel (Targeted)'
    ELSE ''
  END AS 'Base Channel',
  cfgUpdateChannel0,
  ClientCulture0 AS [Client Culture],
  ClientFolder0,
  GPOChannel0,
  GPOOfficeMgmtCOM0,
  InstallationPath0,
  KeyName0,
  LastScenario0,
  LastScenarioResult0,
  CASE
    WHEN OfficeMgmtCOM0 = 'True' THEN 'True'
    ELSE 'False'
  END AS [ConfigMgr Management],
  Platform0 AS [Platform],
  CASE
    WHEN SharedComputerLicensing0 = 1 THEN 'Shared'
    ELSE 'User'
  END AS 'Licensing Model',
  UpdateChannel0,
  UpdatePath0,
  UpdatesEnabled0 AS [Updates Enabled],
  UpdateUrl0,
  VersionToReport0 AS [Version]
FROM v_GS_OFFICE365PROPLUSCONFIGURATIONS
WHERE VersionToReport0 IS NOT NULL"
  
## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $applicationsO365'.csv' -NoTypeInformation -Force

# ==========================
# Client Download History
# ==========================
$ClientDownload = $OutputPath + "ClientDownload"
$SqlQuery = "SELECT [ID]
      ,[DownloadHistoryID] [Download History ID]
      ,[HostName] [Host Name]
      ,[BytesDownloaded] [Bytes Downloaded]
      ,[DistributionPointType] [DP Type]
      ,[DownloadType] [Download Type]
      ,[ContentID]
  FROM [ClientDownloadHistorySources] "
  
## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $ClientDownload'.csv' -NoTypeInformation -Force

# ==========================
# Client Health Summary
# ==========================
$ClientHealth = $OutputPath + "ClientHealth"
$SqlQuery = "DECLARE @Now DateTime = GetDate()
SELECT [ResourceID]
      --,[LastOnline]
      ,[LastMPServerName] [Last MP Server]
      ,[LastDDR] [Last Heartbeat Discovery]
	,Case
		When DATEDIFF(dd,[LastDDR],@Now) between 0 and 7 then 'Past Week'
		When DATEDIFF(dd,[LastDDR],@Now) between 8 and 14 then 'Last 2 Weeks'
		When DATEDIFF(dd,[LastDDR],@Now) between 15 and 21 then 'Last 3 Weeks'
		When DATEDIFF(dd,[LastDDR],@Now) between 22 and 29 then 'Last 4 Weeks'
		When DATEDIFF(dd,[LastDDR],@Now) between 30 and 60 then 'Last 2 Months'
		When DATEDIFF(dd,[LastDDR],@Now) between 61 and 89 then 'Last 3 Months'
		When DATEDIFF(dd,[LastDDR],@Now) >= 90 then 'Over 3 Months'
	 Else 'Never'
	 End As 'Days Since Last DDR Scan'
      ,[LastHW] [Last HW Inventory]
    ,Case
		When DATEDIFF(dd,[LastHW],@Now) between 0 and 7 then 'Past Week'
		When DATEDIFF(dd,[LastHW],@Now) between 8 and 14 then 'Last 2 Weeks'
		When DATEDIFF(dd,[LastHW],@Now) between 15 and 21 then 'Last 3 Weeks'
		When DATEDIFF(dd,[LastHW],@Now) between 22 and 29 then 'Last 4 Weeks'
		When DATEDIFF(dd,[LastHW],@Now) between 30 and 60 then 'Last 2 Months'
		When DATEDIFF(dd,[LastHW],@Now) between 61 and 89 then 'Last 3 Months'
		When DATEDIFF(dd,[LastHW],@Now) >= 90 then 'Over 3 Months'
	 Else 'Never'
	 End As 'Days Since Last HW Scan'
      ,[LastSW] [Last SW Inventory]
	,Case
		When DATEDIFF(dd,[LastSW],@Now) between 0 and 7 then 'Past Week'
		When DATEDIFF(dd,[LastSW],@Now) between 8 and 14 then 'Last 2 Weeks'
		When DATEDIFF(dd,[LastSW],@Now) between 15 and 21 then 'Last 3 Weeks'
		When DATEDIFF(dd,[LastSW],@Now) between 22 and 29 then 'Last 4 Weeks'
		When DATEDIFF(dd,[LastSW],@Now) between 30 and 60 then 'Last 2 Months'
		When DATEDIFF(dd,[LastSW],@Now) between 61 and 89 then 'Last 3 Months'
		When DATEDIFF(dd,[LastSW],@Now) >= 90 then 'Over 3 Months'
	Else 'Never'
	End As 'Days Since Last SW Scan'
      ,[LastStatusMessage] [Last Status Message]
      ,[LastPolicyRequest] [Last Policy Request]
      ,[LastHealthEvaluation] [Last Policy Request]
      ,[LastHealthEvaluationResult] [Last Health Evaluation Result]
      --,[IsActiveDDR]
      --,[IsActiveHW]
      --,[IsActiveSW]
      --,[IsActivePolicyRequest]
      --,[IsActiveStatusMessages]
      --,[LastEvaluationHealthy]
      ,[ClientRemediationSuccess] [Client Remediation Success]
      ,[LastActiveTime] [Last Active Time]
      ,[ClientActiveStatus] [Client Active Status]
      ,[ClientState] [Client State]
      ,[ClientStateDescription] [Client State Description]
      ,[ExpectedNextPolicyRequest] [Expected Next Policy Request]
  FROM [v_CH_ClientSummary]"
 
## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $ClientHealth'.csv' -NoTypeInformation -Force

# ==========================
# Collection Membership
# ==========================
$Colmembership = $OutputPath + "Colmembership"
$SqlQuery = "SELECT Mem.[CollectionID]
      ,[ResourceID]
	  ,Mem.Name [Member]
	  --,Mem.ResourceType
	  ,Case
	  When Mem.ResourceType = 5 Then 'Device'
	  When Mem.ResourceType = 4 Then 'User'
	  Else
	  'unknown'
	  End As 'Member Type'
      ,Col.[Name] [Collection]
      ,[Domain]
      ,[SiteCode]
  FROM v_FullCollectionMembership Mem
  LEFT JOIN v_Collection Col on Mem.CollectionID = Col.CollectionID
  Where mem.ResourceType <> 2"

## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $colmembership'.csv' -NoTypeInformation -Force

# ==========================
# Collection
# ==========================
$Collection = $OutputPath + "Collection"
$SqlQuery = "Select 
--col.CollectionID,
col.SiteID [Collection ID],
CASE
WHEN Col.CollectionType = 2 THEN 'Device'
WHEN Col.CollectionType = 1 THEN 'User'
ELSE NULL
END AS 'Collection Type',
CollectionName [Collection Name],
CollectionComment [Collection Comment],
LimitToCollectionID,
limcol.Name [Limiting Collection],
Col.EvaluationStartTime [Evaluation Start Time],
Col.LastRefreshTime [Last Refresh Time],
Col.LastRefreshRequest [Last Refresh Request],
Col.LastIncrementalRefreshTime [Last Incremental Refresh Time],
Col.MemberCount [Member Count],
(CAST(ColL.EvaluationLength as float)/1000) as 'Time Spent On Full Eval', 
(CAST(ColL.IncrementalEvaluationLength as float)/1000) as 'Time Spent On Incremental Eval' 
 from V_Collections col
 Left Join v_Collection limcol on col.LimitToCollectionID = limcol.CollectionID
 Left Join Collections_L ColL on limcol.CollID = ColL.CollectionID
"

## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $Collection'.csv' -NoTypeInformation -Force

# ==========================
# Devices
# ==========================
$Devices = $OutputPath + "Devices"
$SqlQuery = "DECLARE @Now DateTime = GetDate()

;WITH Devices AS
(

SELECT distinct 
			ci.MachineID [ResourceID],
			ci.Name,
			ci.[ClientVersion] [Client Version] ,
			--ci.[IsClient] [Client],
			ci.[LastStatusMessage] [Last Status Message],
			ci.[LastPolicyRequest] [Last Policy Request],
			ci.[LastDDR] [Last DDR],
			ci.[LastHardwareScan] [Last Hardware Scan],
			ci.[LastSoftwareScan] [Last Software Scan],
			ci.[LastMPServerName] [Last MP Server],
			CH.ClientActiveStatus [Client Active Status],
			CH.ClientStateDescription [Client State],
			CH.ExpectedNextPolicyRequest [Expected Next Policy Request],
			ci.[Domain],
			ci.ADSiteName [AD Site],
			--SYS.AD_Site_Name0 [AD Site], 
			ci.LastActiveTime [Last Active Time],
			ci.isvirtualmachine [Is VM],
			case
			When DATEDIFF(dd,ci.LastActiveTime,@Now) between 0 and 7 then 'Past Week'
			When DATEDIFF(dd,ci.LastActiveTime,@Now) between 8 and 14 then 'Last 2 Weeks'
			When DATEDIFF(dd,ci.LastActiveTime,@Now) between 15 and 21 then 'Last 3 Weeks'
			When DATEDIFF(dd,ci.LastActiveTime,@Now) between 22 and 29 then 'Last 4 Weeks'
			When DATEDIFF(dd,ci.LastActiveTime,@Now) between 30 and 60 then 'Last 2 Months'
			When DATEDIFF(dd,ci.LastActiveTime,@Now) between 61 and 89 then 'Last 3 Months'
			When DATEDIFF(dd,ci.LastActiveTime,@Now) >= 90 then 'Over 3 Months'
			Else 'Never'
			End As 'Last Active',
			ci.Username [Last Logon Username],
			ci.CNLastOnlineTime [Last Online Time],
			ci.CNLastOfflineTime [Last Offline Time],
				SYS.Creation_Date0 [Creation Date], 
			OPSYS.Caption0 as OS, 

			Case
			When OPSYS.Caption0 like '%Windows 7%' Then 'Windows 7 Professional'
			When OPSYS.Caption0 like '%Windows 8.1%' Then 'Windows 8.1'
			When OPSYS.Caption0 like '%2015 LTSB%' Then 'Windows 10 2015 LTSB'
			When OPSYS.Caption0 like '%2016 LTSB%' Then 'Windows 10 2016 LTSB'
			When OPSYS.Caption0 like '%Windows 10%' Then 'Windows 10'
			Else
				''
				End As 'OS Support Name'
			,
			OPSYS.Version0 [OS Version],
		
			CASE
				WHEN ci.[DeviceOS] like '%Workstation%' THEN 'Workstation'
				WHEN ci.[DeviceOS] like '%Server%' THEN 'Server'
			ELSE NULL
			END AS 'OS Type',
            STUFF((SELECT (N','+IPAddr.IP_Addresses0) AS [text()]
            FROM v_RA_System_IPAddresses  IPAddr
            WHERE ci.MachineID = IPAddr.ResourceID for xml path(N''))
            ,1,1,N'') as [IP Addresses], -- if there are multiple IP address then combine them together
            --CSYS.Manufacturer0 [Manufacturer],
			CASE
				when CSYS.Manufacturer0 Like 'Microsoft%' then 'Microsoft'
				when CSYS.Manufacturer0 = 'LENOVO' then 'Lenovo'
				when CSYS.Manufacturer0 Like 'Dell%' then 'Dell'
				when CSYS.Manufacturer0 = 'Hewlett-Packard' then 'HP'
			ELSE CSYS.Manufacturer0
			END AS 'Manufacturer',
            CSYS.Model0 [Model], 
			--Processor.Name0 [CPU], 
			--CASE
				--when Processor.Manufacturer0 like '%Intel%' then 'Intel'
				--when Processor.Manufacturer0 Like '%AMD%' then 'AMD'

			--ELSE Processor.Manufacturer0
			--END AS 'Processor',
			bios.SerialNumber0 'Serial Number',
			--Processor.Manufacturer0 [Processor],
			--Processor.MaxClockSpeed0 [Max Clock Speed],
			MEM.TotalPhysicalMemory0 / 1024 As [Physical Memory (MB)],
			Ram.[Memory Slots],
			CASE
				WHEN chs.ChassisTypes0 = 1 THEN 'Other'
				WHEN chs.ChassisTypes0 = 2 THEN 'Unknown'
				WHEN chs.ChassisTypes0 = 3 THEN 'Desktop'
				WHEN chs.ChassisTypes0 = 4 THEN 'Low Profile Desktop'
				WHEN chs.ChassisTypes0 = 5 THEN 'Pizza Box'
				WHEN chs.ChassisTypes0 = 6 THEN 'Mini Tower'
				WHEN chs.ChassisTypes0 = 7 THEN 'Tower'
				WHEN chs.ChassisTypes0 = 8 THEN 'Portable'
				WHEN chs.ChassisTypes0 = 9 THEN 'Laptop'
				WHEN chs.ChassisTypes0 = 10 THEN 'Notebook'
				WHEN chs.ChassisTypes0 = 11 THEN 'Hand Held'
				WHEN chs.ChassisTypes0 = 12 THEN 'Docking Station'
				WHEN chs.ChassisTypes0 = 13 THEN 'All in One'
				WHEN chs.ChassisTypes0 = 14 THEN 'Sub Notebook'
				WHEN chs.ChassisTypes0 = 15 THEN 'Space-Saving'
				WHEN chs.ChassisTypes0 = 16 THEN 'Lunch-Box'
				WHEN chs.ChassisTypes0 = 17 THEN 'Main System Chassis'
				WHEN chs.ChassisTypes0 = 18 THEN 'Expansion Chassis'
				WHEN chs.ChassisTypes0 = 19 THEN 'Sub Chassis'
				WHEN chs.ChassisTypes0 = 20 THEN 'Bus Expansion Chassis'
				WHEN chs.ChassisTypes0 = 21 THEN 'Peripheral Chassis'
				WHEN chs.ChassisTypes0 = 22 THEN 'Storage Chassis'
				WHEN chs.ChassisTypes0 = 23 THEN 'Rack Mount Chassis'
				WHEN chs.ChassisTypes0 = 24 THEN 'Sealed-Case PC'
			ELSE NULL
			END AS 'Chassis Type',
			--chs.SerialNumber0 [Serial Number ID],
			--chs.[SMBIOSAssetTag0],

			scu.TopConsoleUser0 [Top Console User],
			frm.SecureBoot0 [Secure Boot], 
			frm.UEFI0 [UEFI],
			tpm.SpecVersion0 [TPM Spec Version], 
			tpm.IsActivated_InitialValue0 [TPM Activated], 
			tpm.IsEnabled_InitialValue0 [TPM Enabled], 
			tpm.IsOwned_InitialValue0 [TPM Owned], 
			tps.IsReady0 [TPM Ready], 
			tps.Information0 [TPM Information], 
			tps.IsApplicable0 [TPM Applicable],
            --csys.systemtype0 [Architecture],
			REPLACE (csys.systemtype0,'-based PC','') [Architecture],
			bios.Manufacturer0 [Bios Manufacturer],
			bios.SMBIOSBIOSVersion0 [Bios Version],
			bios.ReleaseDate0 [Bios Released Date],
			ci.IsObsolete [Obsolete]
            FROM vSMS_CombinedDeviceResources ci 


			LEFT JOIN v_R_System SYS on ci.MachineID = SYS.ResourceID
            LEFT JOIN  v_GS_COMPUTER_SYSTEM CSYS on SYS.ResourceID = CSYS.ResourceID
            --LEFT JOIN  v_GS_PROCESSOR  Processor  on Processor.ResourceID = SYS.ResourceID
           

			LEFT JOIN 
			(SELECT a.Resourceid, a.RevisionID, Caption0, Version0, a.GroupID, a.TimeStamp FROM v_GS_OPERATING_SYSTEM  a
			INNER JOIN (SELECT Distinct Resourceid, MAX(TimeStamp) rev FROM v_GS_OPERATING_SYSTEM GROUP BY Resourceid) b ON a.Resourceid = b.Resourceid AND a.TimeStamp = b.rev	
			) OPSYS 
			ON OPSYS.ResourceID = SYS.ResourceID


			LEFT JOIN 
			(SELECT a.Resourceid, a.RevisionID, TotalPhysicalMemory0, a.GroupID, a.TimeStamp FROM v_GS_X86_PC_MEMORY a
			INNER JOIN (SELECT Distinct Resourceid, MAX(TimeStamp) rev FROM v_GS_X86_PC_MEMORY GROUP BY Resourceid) b ON a.Resourceid = b.Resourceid AND a.TimeStamp = b.rev	
			) MEM 
			ON MEM.ResourceID = SYS.ResourceID


			LEFT JOIN 
			(SELECT a.Resourceid, a.RevisionID, a.ChassisTypes0, a.SerialNumber0, a.SMBIOSAssetTag0, a.GroupID, a.TimeStamp FROM v_GS_SYSTEM_ENCLOSURE a
			INNER JOIN (SELECT Resourceid, MAX(TimeStamp) rev FROM v_GS_SYSTEM_ENCLOSURE GROUP BY Resourceid) b ON a.Resourceid = b.Resourceid AND a.TimeStamp = b.rev	
			Where GroupID = 1) chs 
			ON chs.ResourceID = sys.ResourceID

			LEFT JOIN (select ResourceID, COUNT(RAM.ResourceID) [Memory Slots] from v_GS_PHYSICAL_MEMORY RAM
			Group By ResourceID) Ram ON Sys.ResourceID = Ram.ResourceID 

			Left JOIN 
			(SELECT a.Resourceid, a.RevisionID, a.[TopConsoleUser0], a.TimeStamp FROM v_GS_SYSTEM_CONSOLE_USAGE a
			INNER JOIN (SELECT Resourceid, MAX(TimeStamp) rev FROM v_GS_SYSTEM_CONSOLE_USAGE GROUP BY Resourceid) b ON a.Resourceid = b.Resourceid AND a.TimeStamp = b.rev
						) scu 
			ON scu.ResourceID = sys.ResourceID
			
			LEFT JOIN
			(SELECT a.Resourceid, a.RevisionID, a.SecureBoot0, a.UEFI0, a.TimeStamp FROM v_GS_FIRMWARE a INNER JOIN (SELECT Resourceid, MAX(TimeStamp) rev FROM v_GS_FIRMWARE GROUP BY Resourceid) b 
			ON a.Resourceid = b.Resourceid AND a.TimeStamp = b.rev) frm
			ON SYS.ResourceID = frm.ResourceID

			LEFT JOIN 
			(SELECT a.Resourceid, a.RevisionID, a.IsReady0, a.Information0, a.IsApplicable0, a.TimeStamp FROM v_GS_TPM_STATUS a INNER JOIN (SELECT Resourceid, MAX(TimeStamp) rev 
			FROM v_GS_TPM_STATUS GROUP BY Resourceid) b ON a.Resourceid = b.Resourceid AND a.TimeStamp = b.rev) tps 
			ON SYS.ResourceID = tps.ResourceID

			LEFT JOIN 
			(SELECT a.Resourceid, a.RevisionID, a.SpecVersion0, a.IsActivated_InitialValue0, a.IsEnabled_InitialValue0, a.IsOwned_InitialValue0, a.TimeStamp FROM v_GS_TPM a
			INNER JOIN (SELECT Resourceid, MAX(TimeStamp) rev FROM v_GS_TPM GROUP BY Resourceid) b ON a.Resourceid = b.Resourceid AND a.TimeStamp = b.rev) tpm 
			ON SYS.ResourceID = tpm.ResourceID
			
			LEFT JOIN 
			(SELECT a.Resourceid, a.RevisionID, a.[Manufacturer0], a.[ReleaseDate0], a.[SMBIOSBIOSVersion0], a.SerialNumber0, a.TimeStamp FROM [v_GS_PC_BIOS] a
			INNER JOIN (SELECT Resourceid, MAX(TimeStamp) rev FROM v_GS_PC_BIOS GROUP BY Resourceid) b ON a.Resourceid = b.Resourceid AND a.TimeStamp = b.rev) bios 
			ON SYS.ResourceID = bios.ResourceID

			LEFT JOIN v_CH_ClientSummary CH on SYS.ResourceID = CH.ResourceID
			Where ci.ClientType =1 and ci.EAS_DeviceID IS NULL
)

Select Distinct * from Devices
"

## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $Devices'.csv' -NoTypeInformation -Force

# ==========================
# Devices Primary Users
# ==========================
$DevicesPUsers = $OutputPath + "DevicesPUsers"
$SqlQuery = "SELECT --[RelationshipResourceID],
	   [UniqueUserName] [Primary User]
      ,[MachineResourceID] [ResourceID]
      ,[RelationActive] [Relation Active]
      ,[CreationTime] [Created Time]
      ,[MachineResourceName] [Device]
      --,[MachineResourceClientType]
	  ,Obsolete0 [Obsolete]
  FROM [v_UserMachineRelationship] vUMR
  Left Outer Join v_R_System SYS on vUMR.MachineResourceID = SYS.ResourceID
    Where SYS.Obsolete0 <> 1"

## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $DevicesPUsers'.csv' -NoTypeInformation -Force

# ==========================
# Devices Encryption Volumes
# ==========================
$DevicesEV = $OutputPath + "DevicesEV"
$SqlQuery = "SELECT 
ev.ResourceID,
ev.DriveLetter0 [Drive Letter], 
--ev.ProtectionStatus0 [Protection Status],
CASE
	When ev.ProtectionStatus0 = 1 THEN 'True'
	When ev.ProtectionStatus0 = 0 THEN 'False'
ELSE	'Unknown'
End As 'Protection Status'


FROM V_GS_ENCRYPTABLE_VOLUME as ev 
WHERE ev.DriveLetter0 IS NOT NULL;"

## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $DevicesEV'.csv' -NoTypeInformation -Force

# ==========================
# Devices User Intelligence
# ==========================
$DevicesUI = $OutputPath + "DevicesUI"
$SqlQuery = "SELECT [MachineResourceID] [ResourceID]
      ,[UniqueUserName] [Last Login User]
      ,[MachineResourceName] [Device]
      --,[MachineResourceClientType]
      ,[NumberOfLogins] [Number of Logins]
      ,[LastLoginTime] [Last Login Time]
      ,[ConsoleMinutes] [Console Minutes]
  FROM v_UserMachineIntelligence"
   
## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $DevicesUI'.csv' -NoTypeInformation -Force

# ==========================
# Devices - W10 Servicing
# ==========================
$DevicesW10 = $OutputPath + "DevicesW10"
$SqlQuery = "select 

SYS.ResourceID,
vOS.Caption0 [Operating System], 
sys.Build01 [OS Version],
WSLN.Value Version,
CASE WSS.Branch  
         WHEN '0' THEN 'Current Branch'  
         WHEN '1' THEN 'Current Branch for Business'  
         WHEN '2' THEN 'Long Term Servicing Branch'  
END as 'Branch',  
CASE WSS.State 
	WHEN '1' THEN 'Release Ready'  
	WHEN '2' THEN 'Business Ready'
	WHEN '3' THEN 'Expiring Soon'  
	WHEN '4' THEN 'Near End of Life'   
END as 'State',
CASE WSLN.Value
	WHEN '2015 LTSB' THEN 'Windows 10, 2015 LTSB'
	WHEN '2016 LTSB' THEN 'Windows 10, 2016 LTSB'    
	WHEN '1507' THEN 'Windows 10, released July 2015 (version 1507)'  
	WHEN '1511' THEN 'Windows 10, version 1511'
	WHEN '1607' THEN 'Windows 10, version 1607'  
	WHEN '1703' THEN 'Windows 10, version 1703'  
	WHEN '1709' THEN 'Windows 10, version 1709'   
	WHEN '1803' THEN 'Windows 10, version 1803'
END as 'OS Support Name',
WSLN.LocaleID,
WSLN.Name
--SYS.User_Name0 [Primary User], 

from v_R_System SYS
JOIN vSMS_WindowsServicingStates WSS on SYS.OSBranch01 = WSS.Branch and SYS.Build01 = WSS.Build
JOIN vSMS_WindowsServicingLocalizedNames WSLN on WSS.Name = WSLN.Name
join v_GS_OPERATING_SYSTEM vOS on SYS.ResourceID = vOS.ResourceID and vOS.Caption0 like '%Windows 10%'
"
   
## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $DevicesW10'.csv' -NoTypeInformation -Force

# ==========================
# Site
# ==========================
$Site = $OutputPath + "Site"
$SqlQuery = "select top 1
SiteCode
,GETDATE() [Refresh Date]
,SiteName
,Version
,BuildNumber
,Upd.PackageGuid
,LUpd.Name [Site Branch Version]
,REPLACE (LUpd.Name,'Configuration Manager ','') [Branch Version]
 from v_Site Ste
Inner Join [vSMS_CM_UpdatePackages] Upd ON Ste.Version = Upd.FullVersion
Inner Join v_LocalizedUpdatePackageMetaData_SiteLoc LUpd ON LUpd.PackageGuid = Upd.PackageGuid"

## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $Site'.csv' -NoTypeInformation -Force

# ==========================
# Site Status
# ==========================
$SiteS = $OutputPath + "SiteStatus"
$SqlQuery = "SELECT 

	SiteSystem as [Site System]
	,SiteCode as [Site Code]
	,SiteObject as [Site Object]
	,Case 
		When [ObjectType] = 1 Then ' Transaction Log'
		When [ObjectType] = 2 Then ' Database'
		Else ''
	End AS 'SQL Type'
    ,Role
	,Status
    ,Case
		WHEN Status = 0 THEN 'OK'
		WHEN Status = 1 THEN 'Warning'
		WHEN Status = 2 THEN 'Critical'
	 END As [Status Value]

	 
	,BytesTotal as [Byte Total]
	,BytesFree as [Byte Free]
	,PercentFree as [Percent Free]
     
  FROM v_SiteSystemSummarizer
"

## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $SiteS'.csv' -NoTypeInformation -Force

# ==========================
# Site Updates
# ==========================
$SiteU = $OutputPath + "SiteUpdates"
$SqlQuery = "Select UP.Name
,Upd.FullVersion
,Upd.DateReleased
,Upd.MoreInfoLink
,Case
When CAST( Upd.State AS varchar) = '2'	  then 'Checking Prerequisites'
When CAST( Upd.State AS varchar) = '65538'  then 'Checking Prerequisites'
When CAST( Upd.State AS varchar) = '131074' then 'Prerequisite check passed'
When CAST( Upd.State AS varchar) = '196609' then 'Installing'
When CAST( Upd.State AS varchar) = '196612' then 'Installed'
When CAST( Upd.State AS varchar) = '196619' then 'Installing'
When CAST( Upd.State AS varchar) = '262146' then 'Ready to Install'
When CAST( Upd.State AS varchar) = '327682' then 'Available to Download'
Else CAST( Upd.State AS varchar)
End As 'Status'
,Case
When CAST( Upd.UpdateType AS varchar) = 0 then 'Core'
When CAST( Upd.UpdateType AS varchar) = 2 then 'Hotfix'
Else CAST( Upd.UpdateType AS varchar)
End As 'Update Type'
,Upd.ClientVersion [Client Version]
 from  [vSMS_CM_UpdatePackages] Upd
Inner Join (
select  all SMS_CM_UpdatePackages.Name,SMS_CM_UpdatePackages.PackageGuid from fn_ListCMUpdatePackages(1033) AS SMS_CM_UpdatePackages) UP on UP.PackageGuid = Upd.PackageGuid

"

## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $SiteU'.csv' -NoTypeInformation -Force

# ==========================
# Site Updates History
# ==========================
$SiteSH = $OutputPath + "SiteSH"
$SqlQuery = "/****** Script for SelectTopNRows command from SSMS  ******/
SELECT Upd.[PackageGuid]
      --,[SiteNumber]
      --,[State]
      ,[LastUpdateTime] [Install Date]
      
	  ,LUpd.Name [Installed Update]
  FROM vSMS_CM_UpdatePackageSiteStatus Upd
  Inner Join v_LocalizedUpdatePackageMetaData_SiteLoc LUpd ON LUpd.PackageGuid = Upd.PackageGuid
  Where State = 196612"
  
## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $SiteSH'.csv' -NoTypeInformation -Force

# ==========================
# Update Compliance
# ==========================
$UpdateC = $OutputPath + "UpdateCom"
$SqlQuery = "Select Distinct
vUCS.ResourceID,
vUI.CI_ID [Update CI_ID],					
vUI.ArticleID	[Article ID],
vUI.BulletinID	[Bulletin ID],
vUI.Severity,	
--vUI.UpdateSource_ID,
vSUS.UpdateSourceName [Update Source],
vUI.Title [Update Name],
vUI.Description [Update Description],
vUI.IsEnabled [Enabled],
vUI.IsDeployed [Deployed],
vUI.IsSuperseded [Superseded],
vUI.InfoURL [Info URL],
vUCS.Status [Status ID],
vSNStatus.StateName [Update Status],
vUCS.LastStatusCheckTime [Last Status Check Time],
vUCS.LastEnforcementMessageTime [Last Enforcement Message Time],
--vUCS.LastEnforcementMessageID,
vSNEM.StateName [Enforcement Status],
VCICA.CategoryInstanceName [Update Category],
--VCICA.LocaleID,
vUSS.ScanTime [Scan Time],
--vUSS.LastScanState,
vSNScanState.StateName [Scan Status],
vUSS.LastErrorCode [Last Error Code]

From v_Update_ComplianceStatus vUCS
Left JOIN v_UpdateInfo vUI	on vUCS.CI_ID = vUI.CI_ID 
Left JOIN v_CICategoryInfo_All vCICA		on vUI.CI_ID = vCICA.CI_ID AND VCICA.CategoryTypeName = 'UpdateClassification'
JOIN v_BundledConfigurationItems vBCI		on vUI.CI_ID = vBCI.BundledCI_ID 
JOIN v_UpdateScanStatus vUSS				on vUSS.ResourceID = vUCS.ResourceID
Left JOIN v_StateNames vSNEM				on vUCS.LastEnforcementMessageID = vSNEM.StateID and vSNEM.TopicType = 402
Left JOIN v_StateNames vSNStatus			on vUCS.Status = vSNStatus.StateID and vSNStatus.TopicType = 500
Left JOIN v_StateNames vSNScanState			on vUSS.LastScanState = vSNScanState.StateID and vSNScanState.TopicType = 501
Left JOIN v_SoftwareUpdateSource vSUS		on vSUS.UpdateSource_ID = vUI.UpdateSource_ID"

## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $UpdateC'.csv' -NoTypeInformation -Force

# ==========================
# Update Assignment Compliance
# ==========================
$UpdateAC = $OutputPath + "UpdateACom"
$SqlQuery = "Select

vCIA.AssignmentID,	
vUASL.ResourceID,
vALI.Title [SUG Name],	
vALI.Description [SUG Description],
IsCompliant [Compliant],
ComplianceTime [Compliance Time],
--vSNCo.StateName [Compliance State],
CASE 
	WHEN vSNCo.StateName IS NULL THEN 'Unknown'
	ELSE vSNCo.StateName
END As [Compliance State],
--LastComplianceMessageID,
--LastComplianceMessageTime,
vSNEn.StateName [Last Enforcement State],
--LastEnforcementMessageID,
--LastEnforcementErrorID,
LastEnforcementErrorCode [Last Enforcement Error Code],
--vUASL.LastEnforcementMessageTime,
--LastEnforcementIsRebootSupressed,
vSNEv.StateName [Last Evaluation State],
--LastEvaluationMessageID,
LastEvaluationErrorID,
LastEvaluationErrorCode,
LastEvaluationMessageTime,


AssignmentName [Assignment Name],
vCIA.Description [Assignment Description],
CollectionID [Collection ID],
CollectionName [Collection Name],
AssignmentAction [Assignment Action],
ExpirationTime [Expiration Time],
StartTime [Start Time],
EnforcementDeadline [Enforcement Deadline],
--vCIA.LastModifiedBy,
--UserUIExperience,
--AssignmentEnabled,
--LimitStateMessageVerbosity,
--StateMessageVerbosity,	
--StateMessagePriority,
--UseBranchCache,
--RequirePostRebootFullScan,
vALI.CI_ID [SUG CI_ID]
FROM  v_UpdateAssignmentStatus_Live vUASL

JOIN		v_CIAssignment vCIA			 on vUASL.AssignmentID = vCIA.AssignmentID
LEFT JOIN	v_CIAssignmentToGroup vCATG  on vUASL.AssignmentID = vCATG.AssignmentID
LEFT JOIN   v_AuthListInfo vALI			 on vALI.CI_ID = vCATG.AssignedUpdateGroup
LEFT JOIN v_StateNames vSNCo			 on vUASL.LastComplianceMessageID  = vSNCo.StateID and vSNCo.TopicType=300
LEFT JOIN v_StateNames vSNEn			 on vUASL.LastEnforcementMessageID = vSNEn.StateID and vSNEn.TopicType=402
LEFT JOIN v_StateNames vSNEv			 on vUASL.LastEvaluationMessageID = vSNEv.StateID and vSNEv.TopicType=302"
   
## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $UpdateAC'.csv' -NoTypeInformation -Force

# ==========================
# Database Framentation
# ==========================
$DBFrag = $OutputPath + "DBFrag"
$SqlQuery = "Select DB_NAME (database_id) AS [Database Name], OBJECT_NAME(ps.object_id) AS [Object Name],
       i.name AS [Index Name], ps.index_id, index_type_desc,
       avg_fragmentation_in_percent, fragment_count, page_count
FROM sys.dm_db_index_physical_stats(DB_ID(),NULL, NULL, NULL ,N'LIMITED') AS ps
       INNER JOIN sys.indexes AS i WITH (NOLOCK)
       ON ps.[object_id] = i.[object_id] AND ps.index_id = i.index_id
WHERE database_id = DB_ID()
AND page_count > 1500
ORDER BY avg_fragmentation_in_percent DESC OPTION (RECOMPILE);"

## - Connect to SQL Server using non-SMO class 'System.Data': 
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $SQLUsername; Password = $SQLPassword"  
$SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
$SqlCmd.CommandText = $SqlQuery
$SqlCmd.Connection = $SqlConnection  
$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
$SqlAdapter.SelectCommand = $SqlCmd  
$DataSet = New-Object System.Data.DataSet 
$SqlAdapter.Fill($DataSet) 
$SqlConnection.Close()  
#Output RESULTS to CSV
$DataSet.Tables[0] | Export-Csv -path $DBFrag'.csv' -NoTypeInformation -Force

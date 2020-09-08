# ===============
# Load Config.xml
# ===============
# Load XML config file with client settings:
$Path = "C:\script\Reporting\Config.xml"
# load it into an XML object:
$xml = New-Object -TypeName XML
$xml.Load($Path)

# ==========================
# Modules to Install/import
# ==========================

# Import Active Directory
If ($xml.configuration.Settings.AD.Enabled -eq "True")
{
Import-Module activedirectory
}

# ==========
# Variables
# ==========

$exportpath = "C:\Script\Reporting\Export\"

$client = $xml.configuration.export.client

# =======================
# AD Info Export
# =======================

If ($xml.configuration.Settings.AD.Enabled -eq "True")
{
$ADInfoExportpath = $exportpath + $client + "ADInfo"

$CompanyInfoTable = New-Object 'System.Collections.Generic.List[System.Object]'

$ADRecycleBinStatus = (Get-ADOptionalFeature -Filter 'name -like "Recycle Bin Feature"').EnabledScopes
if ($ADRecycleBinStatus.Count -lt 1)
{
$ADRecycleBin = "Disabled"
}
else
{
$ADRecycleBin = "Enabled"
}
$ForestObj = Get-ADForest
$DomainControllerobj = Get-ADDomain
$InfrastructureMaster = $DomainControllerobj.InfrastructureMaster
$RIDMaster = $DomainControllerobj.RIDMaster
$PDCEmulator = $DomainControllerobj.PDCEmulator
$DomainNamingMaster = $ForestObj.DomainNamingMaster
$SchemaMaster = $ForestObj.SchemaMaster
$PwdPolicy = Get-ADDefaultDomainPasswordPolicy

$obj = [PSCustomObject]@{
'Client' = $client
'Customer' = $xml.configuration.export.Customer
'AD Recycle Bin' = $ADRecycleBin
'Infrastructure Master' = $InfrastructureMaster
'RID Master' = $RIDMaster
'PDC Emulator' = $PDCEmulator
'Domain Naming Master' = $DomainNamingMaster
'Schema Master' = $SchemaMaster
'ComplexityEnabled' = $PwdPolicy.ComplexityEnabled 
'LockoutDuration' = $PwdPolicy.LockoutDuration.Minutes
'LockoutThreshold' = $PwdPolicy.LockoutThreshold
'MaxPasswordAge' = $PwdPolicy.MaxPasswordAge.Days
'MinPasswordAge' = $PwdPolicy.MinPasswordAge.Days
'MinPasswordLength' = $PwdPolicy.MinPasswordLength
'PasswordHistoryCount' = $PwdPolicy.PasswordHistoryCount
'ReversibleEncryptionEnabled' = $PwdPolicy.ReversibleEncryptionEnabled
}
$CompanyInfoTable.Add($obj)
$CompanyInfoTable | Export-Csv -Path $ADInfoExportpath".csv" -NoTypeInformation
}

# =======================
# AD User Export 
# =======================

If ($xml.configuration.Settings.AD.Enabled -eq "True")
{
    $ADUserExportpath = $exportpath + $client + "ADUsers"
    $ADUsers = Get-ADUser -Filter *
    $dcs = Get-ADDomainController -Filter {Name -like "*"}
    $TheUsers = foreach($ADUser in $ADUsers)
    {
      $time = 0
      foreach($dc in $dcs)
      { 
        $DCIP = $dc.IPv4Address
        $user = Get-ADUser $ADUser -Server $DCIP -prop *
        if($user.lastLogon -gt $time) 
        {
          $time = $user.lastLogon
        }
        if($user.lastlogontimestamp -gt $time)
        {
            $time = $user.lastlogontimestamp
        }
      }
      $user | Select-Object @{n='Client';e={$client}},`
                    @{n='Name';e={$_.Name}},`
                   @{n='First Name';e={$_.givenName}},`
                   @{n='Surname';e={$_.sn}}, `
                   @{n='Display Name';e={$_.DisplayName}}, `
                   @{n='Username';e={$_.sAMAccountName}}, `
                   @{n='Enabled';e={$_.enabled}}, `
                   @{n='Description';e={$_.description}}, `
                   @{n='Emailaddress';e={$_.mail}}, `
                   @{n='Job Title';e={$_.title}}, `
                   @{n='Department';e={$_.Department}}, `
                   @{n='Division';e={$_.Division}}, `
                   @{n='Creation Date';e={$_.whenCreated}}, `
                   @{n='Modification Date';e={$_.Modified}}, `
                   @{n='Expiration Date';e={$_.AccountExpirationdate}}, `
                   @{n='Allow Log On To';e={$_.Logonworkstations}}, `
                   @{n='Account Is Locked Out';e={$_.Lockedout}}, `
                   @{n="LastLogonTime";e={if($time -ne 0) {[datetime]::FromFileTime($time)}}}, `
                   @{n='Password Last Set';e={if($_.PwdLastSet -ne $null) {[datetime]::FromFileTime($_.PwdLastSet)}}}, `
                   @{n='Password Age';e={if($_.PwdLastSet -ne 0){(new-TimeSpan([datetime]::FromFileTimeUTC($_.PwdLastSet)) $(Get-Date)).days}else{0}}}, `
                   @{n='Password Never Expires';e={$_.Passwordneverexpires}}, `
                   @{n='SystemObject';e={$_.isCriticalSystemObject}}, `
                   @{n='Parent Container';e={$_.canonicalname -replace "/$($_.cn)",""}}             
    }
    $Theusers | Export-Csv -Path $ADUserexportpath".csv" -NoTypeInformation
}

# =======================
# AD Computers Export
# =======================

If ($xml.configuration.Settings.AD.Enabled -eq "True")
{
    $ADComputerExportpath = $exportpath + $client + "ADComputer"
    $ADComputers =  Get-ADComputer -Filter *
    $dcs = Get-ADDomainController -Filter {Name -like "*"}
    $TheComputers = foreach($ADComputer in $ADComputers)
    {
      $time = 0
      foreach($dc in $dcs)
      { 
        $DCIP = $dc.IPv4Address
        $computer = Get-ADComputer $ADComputer -Server $DCIP -prop *
        if($computer.lastLogon -gt $time) 
        {
          $time = $computer.lastLogon
        }
        if ($computer.lastlogontimestamp -gt $time)
        {
            $time = $computer.lastlogontimestamp
        }
      }
          $computer | Select-Object @{n='Client';e={$client}},`
                            @{n='Name';e={$_.Name}}, `
                            @{n='Username';e={$_.sAMAccountName}}, `
                            @{n='Description';e={$_.description}}, `
                            @{n='Enabled';e={$_.enabled}}, `
                            @{n="LastLogonTime";e={if($time -ne $null) {[datetime]::FromFileTime($time)}}}, `
                            @{n='Creation Date';e={$_.whenCreated}}, `
                            @{n='Modification Date';e={$_.whenChanged}}, `
                            @{n='Parent Container';e={$_.canonicalname -replace "/$($_.cn)",""}}, `
                            @{n='Operating System';e={$_.operatingSystem}}, `
                            @{n='O/S Version Number';e={$_.operatingSystemVersion}}, `
                            @{n='O/S Service Pack';e={$_.operatingSystemServicePack}}, `
                            @{n='DNS Host Name';e={$_.dNSHostName}}, `
                            @{n='IPv4';e={$_.IPv4Address}}, `
                            @{n='SystemObject';e={$_.isCriticalSystemObject}}, `
                            @{n='KerberosEncryptionType';e={$_.KerberosEncryptionType}}, `
                            @{n='SupportedEncryptionTypes';e={$_.'msDS-SupportedEncryptionTypes'}}                                   
    }
    $TheComputers | Export-Csv -Path $ADComputerExportpath".csv" -NoTypeInformation
}

# =======================
# AD Groups Export
# =======================

If ($xml.configuration.Settings.AD.Enabled -eq "True")
{
    $ADGroupExportpath = $exportpath + $client + "ADGroup"
    $ADGroup = Get-ADGroup -Filter * -Properties *
    $GroupTable = New-Object 'System.Collections.Generic.List[System.Object]'
    foreach($group in $ADGroup)
    {
        
        $members = Get-ADGroupMember -Identity $group.SamAccountName -Recursive
        If ($members.length -eq "0")
        {
            $obj = [PSCustomObject]@{
                    'client' = $client
                    'Group Name' = $group.Name
                    'Group Description' = $group.Description
                    'Group Type' = $group.GroupCategory
                    'Group Scope' = $group.GroupScope
                    'Group Created' = $group.Created
                    'Group Modified' = $group.Modified
                    'Member Name' = "None"
                    'Member Username' = "None"
                    'Member Last Logon' = "None"
                    'Member Enabled' = "None"
                    'Parent Container' = $group.CanonicalName -replace "/$($group.cn)",""
                    }
                $GroupTable.add($obj)
        }
        ELSE
        {
            foreach ($member in $members)
            {
                    $obj = [PSCustomObject]@{
                    'client' = $client
                    'Group Name' = $group.Name
                    'Group Description' = $group.Description
                    'Group Type' = $group.GroupCategory
                    'Group Scope' = $group.GroupScope
                    'Group Created' = $group.Created
                    'Group Modified' = $group.Modified
                    'Member Name' = $member.name
                    'Member Username' = $member.SamAccountName
                    'Member Last Logon' = ($theusers | Where-Object {$_.username -eq $member.SamaccountName}).LastLogonTime
                    'Member Enabled' = ($theusers | Where-Object {$_.username -eq $member.SamaccountName}).Enabled
                    'Parent Container' = $group.CanonicalName -replace "/$($group.cn)",""
                    }
                $GroupTable.add($obj)
            }
        }
    }
    $GroupTable | Export-Csv -Path $ADGroupexportpath".csv" -NoTypeInformation
}
   
# =======================
# GPO export
# =======================

If ($xml.configuration.Settings.AD.GPO.Enabled -eq "True")
{
$GPOExportpath = $exportpath + $client + "GPO"

$GPOTable = New-Object 'System.Collections.Generic.List[System.Object]'

$GPOs = Get-GPO -All

foreach ($GPO in $GPOS)
        {
         If ($null -ne ($gpo | Get-GPOReport -ReportType XML | Select-String -NotMatch "<LinksTo>"))
         {
            $notlinked = "True"
         }
         ELSE
         {
            $notlinked = "False"
         }
         if ((Get-GPPermission -Guid $gpo.id -All).permission -notcontains "GpoApply")
         {
            $HasSecurityFilters = "False"
         }
         Else
         {
            $HasSecurityFilters = "True"
         }

         $obj = [PSCustomObject]@{
            'Client' = $client
            'Name' = $GPO.DisplayName
            'Status' = $GPO.GpoStatus
            'Modified Date' = $GPO.ModificationTime
            'Created Date' = $GPO.creationTime
            'User Version' = $GPO.User.DSVersion
            'Computer Version' = $GPO.Computer.dsversion
            'Not Linked' = $notlinked
            'Has Security Filters' = $HasSecurityFilters
            }
        $GPOTable.Add($obj)
        }
    $GPOTable | Export-Csv -Path $GPOExportpath".csv" -NoTypeInformation
}

# =======================
# AD Deleted Export 
# =======================

If ($xml.configuration.Settings.AD.Enabled -eq "True")
{
    $ADDeletedExportpath = $exportpath + $client + "ADDeleted"
    $DeletedTable = New-Object 'System.Collections.Generic.List[System.Object]'
    $dcs = Get-ADDomainController -Filter {Name -like "*"}
    $ADDeleted = get-adobject -filter {IsDeleted -eq $True} -IncludeDeletedObjects
    foreach($Deleted in $ADDeleted)
    {
        $time = 0
        foreach($dc in $dcs)
        { 
            $DCIP = $dc.IPv4Address
            $Del = get-adobject -filter {name -eq $deleted.name} -IncludeDeletedObjects -Properties * -Server $DCIP
            if($Del.lastLogon -gt $time) 
            {
                $time = $Del.lastLogon
            }
            if ($Del.lastlogontimestamp -gt $time)
            {
                $time = $Del.lastlogontimestamp
            }
        }
            $obj = [PSCustomObject]@{
            'Client' = $client
            'Name' = $Del.Name
            'Username' = $Del.samaccountname
            'Description' = $del.description
            'Modified Date' = $Del.Modified
            'Created Date' = $Del.Created
            'LastLogonTime' = [datetime]::FromFileTime($time)
            'Object Type' = $Del.ObjectClass
            'Deleted' = $del.isDeleted
            'Last Known OU' = $del.LastKnownParent
            }
            $DeletedTable.add($obj)          
    }
    $Deletedtable | export-csv -Path $ADDeletedExportpath".csv" -NoTypeInformation
}

# =======================
# AD Deleted Export 
# =======================

If ($xml.configuration.Settings.AD.Enabled -eq "True")
{
    $ADOUExportpath = $exportpath + $client + "ADOU"
    $OUTable = New-Object 'System.Collections.Generic.List[System.Object]'
    $ADOU = Get-ADOrganizationalUnit -Filter * -Properties *

    foreach($OU in $ADOU) 
    {
            $objects = Get-adobject -SearchBase $ou.DistinguishedName -Filter * -Properties *
        
            $obj = [PSCustomObject]@{
            'Client' = $client
            'Name' = $ou.Name
            'Created' = $ou.createTimeStamp
            'Modified' = $ou.modifyTimeStamp
            'Number of Linked GPO' = $ou.LinkedGroupPolicyObjects.count
            'Protected' = $ou.ProtectedFromAccidentalDeletion
            'TotalObj' = $objects.count
            'ComputerObj' = ($objects | Where-Object {$_.objectclass -eq "computer"}).count
            'OUObj' = ($objects | Where-Object {$_.objectclass -eq "organizationalUnit"}).count
            'UserObj' = ($objects | Where-Object {$_.objectclass -eq "user"}).count
            'groupObj' = ($objects | Where-Object {$_.objectclass -eq "group"}).count
            'Parent Container' = $ou.CanonicalName -replace "$($ou.cn)",""
            }
            $OUTable.add($obj)          
    }
    $OUtable | export-csv -Path $ADOUExportpath".csv" -NoTypeInformation
}

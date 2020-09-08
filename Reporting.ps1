# ===============
# Load Config.xml
# ===============
# Load XML config file with client settings:
$Path = "C:\script\Reporting\Config.xml"
# load it into an XML object:
$xml = New-Object -TypeName XML
$xml.Load($Path)

# =================
# Check for Update
# =================

If ($xml.configuration.Update.Enabled -eq "True")
{
    $GithubUrl = "https://raw.githubusercontent.com/ELLISB1000/Reporting/master/Config_Template.xml"
    $currentversion = $xml.configuration.update.Version
    $nextversion = $null
    try
    {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        [XML]$Gitversion = (New-Object System.Net.WebClient).Downloadstring($githuburl)
        $nextversion = $gitversion.configuration.Update.Version
    }
    Catch
    {
        Write-Verbose "Unable to Download version from Github"
    }
    If ($null -eq $currentversion -or $currentversion -ne $nextversion)
        {
            Write-Verbose "Updating script from $($currentversion) to $($nextversion)"
            $updatepath = "C:\script\Reporting\Update.zip"
            $updatefile = "https://github.com/ELLISB1000/Reporting/archive/master.zip"
            Remove-Item -Path "C:\script\Reporting" -Force -confirm:$false
            new-item -Name "Reporting" -path "C:\script" -itemtype "directory"
            (New-Object System.Net.WebClient).DownloadFile($updatefile, $updatepath)
            Expand-Archive $updatepath -Force -DestinationPath "C:\script\Reporting"
            Remove-Item -Path $updatepath -Force -confirm:$false
            $items = Get-ChildItem -Path "C:\script\Reporting\Reporting-Master"
            Foreach ($item in $items)
                {
                    If ($item.name -ne "Installer")
                    {
                        Copy-Item $item.FullName -Destination C:\script\Reporting -Force -Recurse -confirm:$false
                    }
                }
            Remove-Item -Path "C:\script\reporting\Reporting-master" -Recurse -Force
            $xml.configuration.update.Version = $nextversion
            $xml.Save($path)
       }
}

# ==========================
# Modules to Install/import
# ==========================

# Import Active Directory
If ($xml.configuration.Settings.AD.Enabled -eq "True")
{
    try
    {
        Get-Module -ListAvailable -Name "activedirectory"
    }
    Catch
    {
        Write-Verbose "Unable to Import Active Directory Powershell module"
    }
}

# Import CA Powershell Module
If ($xml.configuration.Settings.CA.Enabled -eq "True")
{
    try
    {
        $CAMod = Get-Module -ListAvailable -Name "pspki"        
        If ($null -eq $CAMod)
        {
            Install-Module -Name pspki -Force
        }
    }
    Catch
    {
        Write-Verbose "Unable to Import CA Powershell module"
    }
}


# Import SMTP Relay Powershell Module
If ($xml.configuration.Settings.SMTPRelay.Enabled -eq "True")
{
    Import-Module ServerManager 
    Foreach ($server in $xml.configuration.Settings.SMTPRelay.servers)
    {
        $wmi = get-windowsfeature WEB-WMI -cn $server
        If ($wmi.installstate -ne "Available")
        {
            try
            {
                # add-windowsfeature WEB-WMI -Confirm:$false -cn $server
            }
            Catch
            {
                Write-Verbose "Unable to Install IIS WMI Feature"
            }
        }
    }
}

# ==========
# Variables
# ==========

#CREATE EMPTY DIRECTORY
$exportpath = "C:\Script\Reporting\Export\"

if ((Test-Path -Path $exportpath) -ne $true) {
    Write-Verbose "Created Export file Path $($exportpath)"
    New-Item -Path '$exportpath' -ItemType Directory
}

Get-ChildItem -path $exportpath | Remove-Item -Force -Confirm:$false

$client = $xml.configuration.export.client

# =======================
# =======================
# =======================
# Dashboard Export 
# =======================
# =======================
# =======================

# =======================
# Start Jobs
# =======================

# Remove old jobs
If(Get-Job -name "Reporting*") 
{ 
    Write-Verbose "Removing old jobs" 
    Get-Job -name "Reporting*" | Remove-Job -Force 
} 

# Run Active Directory Script
If ($xml.configuration.Settings.AD.Enabled -eq "True")
{
    Write-Verbose "Starting Job ReportingAD"
    Start-Job -Name "ReportingAD" -ScriptBlock {Invoke-Expression -Command "C:\Script\Reporting\Functions\AD.ps1"}
}

# Run DC Events Script
If ($xml.configuration.Settings.AD.Enabled -eq "True")
{
    Write-Verbose "Starting Job DC Events"
    Start-Job -Name "ReportingDCEvents" -ScriptBlock {Invoke-Expression -Command "C:\Script\Reporting\Functions\DCEvents.ps1"}
}

# Run CA Script
If ($xml.configuration.Settings.CA.Enabled -eq "True")
{
    Write-Verbose "Starting Job ReportingCA"
    Start-Job -Name "ReportingCA" -ScriptBlock {Invoke-Expression -Command "C:\Script\Reporting\Functions\CA.ps1"}
}

# Run SMTP Relay Script
If ($xml.configuration.Settings.SMTPRelay.Enabled -eq "True")
{
    Write-Verbose "Starting Job ReportingSMTP"
    Start-Job -Name "ReportingSMTP" -ScriptBlock {Invoke-Expression -Command "C:\Script\Reporting\Functions\SMTP.ps1"}
}

# Run SCCM Script
If ($xml.configuration.Settings.SCCM.Enabled -eq "True")
{
    Write-Verbose "Starting Job ReportingSCCM"
    Start-Job -Name "ReportingSCCM" -ScriptBlock {Invoke-Expression -Command "C:\Script\Reporting\Functions\SCCM.ps1"}
}

# Run Nmap Script
If ($xml.configuration.Settings.Nmap.Enabled -eq "True")
{
    Write-Verbose "Starting Job ReportingNmap"
    Start-Job -Name "ReportingNmap" -ScriptBlock {Invoke-Expression -Command "C:\Script\Reporting\Functions\NMAP.ps1"}
}

# =======================
# Check Jobs
# =======================
$JobExportpath = $exportpath + $client + "Job"
$JobTable = New-Object 'System.Collections.Generic.List[System.Object]'
do
{
    $running = Get-Job -name "Reporting*" | Where-Object {$_.State -eq "Running"}
    Write-Verbose "Currently $($running.count) jobs are running"
    $Jobs = Get-Job -name "Reporting*" -ErrorAction SilentlyContinue | Where-Object {$_.State -eq "Completed"}
    foreach ($job in $jobs)
    {
        Write-Verbose "Job $($job.name) Completed"
        $obj = [PSCustomObject]@{
        'Client' = $client
        'Job Name' = $job.Name
        'Job State' = $job.jobstateinfo
        'Job Start' = $job.PSBegintime
        'Job End' = $job.psendtime
        'Runtime' = New-TimeSpan -Start $job.PSBeginTime -End $job.PSEndTime -ErrorAction SilentlyContinue
        }
        $Jobtable.Add($obj)
    Write-Verbose "Removing job $($job.name)"
    Remove-Job -id $job.id -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 60
    $job = $null
    }
}
While(Get-job -name "Reporting*")
{
    Write-Verbose "Scan Finished"
}
$JobTable | export-csv -Path $jobExportpath".csv" -NoTypeInformation   
# =======================
# Config Export
# =======================

    $ConfigExportpath = $exportpath + $client + "config"
    $ConfigTable = New-Object 'System.Collections.Generic.List[System.Object]'
    Import-Module ScheduledTasks -Force
    $Task = Get-ScheduledTask -TaskName "Reporting"  
    $ConfigObj = [PSCustomObject]@{
            'Client' = $client
            'Certificate Autority' = $xml.configuration.Settings.CA.Enabled
            'Certificate Autority Servers' = $xml.configuration.Settings.CA.Servers
            'SMTP Relay' = $xml.configuration.Settings.SMTPRelay.Enabled
            'SMTP Relay Servers'= $xml.configuration.Settings.SMTPRelay.Servers
            'AD' = $xml.configuration.Settings.AD.Enabled
            'GPO' = $xml.configuration.Settings.Ad.GPO.enabled
            'SCCM' = $xml.configuration.settings.SCCM.enabled
            'Running on' = $env:COMPUTERNAME
            'Last Run' = Get-Date -Format "MM/dd/yyyy HH:mm"
            'Auto-Update' = $xml.configuration.Update.Enabled
            'Version' = $xml.configuration.Update.Version
            'Task Account' = $task.principal.userid
            'Last Run Time' = ($task | Get-ScheduledTaskInfo).lastruntime
            'Last Run Result' = ($task | Get-ScheduledTaskInfo).lasttaskresult
            'Nest Run Time' = ($task | Get-ScheduledTaskInfo).nextruntime
            'Number of Missed Runs' = ($task | Get-ScheduledTaskInfo).numberofmissedruns
            'Powershell Version' = $PSVersionTable.Psversion
            }
        $ConfigTable.Add($configobj)
    $ConfigTable | export-csv -Path $ConfigExportpath".csv" -NoTypeInformation





        
        
        


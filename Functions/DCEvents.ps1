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
# DC Events Export 
# =======================

If ($xml.configuration.Settings.AD.Enabled -eq "True")
{
    $eventsExportpath = $exportpath + $client + "DCEvents"
    $EventsTable = New-Object 'System.Collections.Generic.List[System.Object]'
    $dcs = Get-ADDomainController -Filter {Name -like "*"}
    foreach($dc in $dcs)
    {
        $lockedoutEvents = Get-WinEvent -ComputerName $DC.IPv4Address -FilterHashtable @{LogName='Security';Id=4724,4740,4725,4722,4776,4771;StartTime=[datetime]::Today.AddDays(-3);EndTime=$(Get-Date)} -ErrorAction SilentlyContinue
    
        Foreach($Event in $LockedOutEvents)
        {
            If($event.id -eq "4724")
            {
                $obj = [PSCustomObject]@{
                'Client' = $client
                'User' = $event.Properties[4].Value
                'Target' = $event.Properties[0].Value
                'DomainController' = $event.MachineName
                'EventId' = $event.Id
                'EventTimestamp' = $event.TimeCreated
                'Message' = $event.Message -split "`r" | Select-Object -First 1
                }
            }
            elseif($event.id -eq "4740")
            {
                $obj = [PSCustomObject]@{
                'Client' = $client
                'User' = $event.Properties[0].Value
                'Target' = $event.Properties[1].Value
                'DomainController' = $event.MachineName
                'EventId' = $event.Id
                'EventTimestamp' = $event.TimeCreated
                'Message' = $event.Message -split "`r" | Select-Object -First 1
                }
            }
            elseif($event.id -eq "4725")
            {
                $obj = [PSCustomObject]@{
                'Client' = $client
                'User' = $event.Properties[4].Value
                'Target' = $event.Properties[0].Value
                'DomainController' = $event.MachineName
                'EventId' = $event.Id
                'EventTimestamp' = $event.TimeCreated
                'Message' = $event.Message -split "`r" | Select-Object -First 1
                }
            }
            elseif($event.id -eq "4722")
            {
                $obj = [PSCustomObject]@{
                'Client' = $client
                'User' = $event.Properties[4].Value
                'Target' = $event.Properties[0].Value
                'DomainController' = $event.MachineName
                'EventId' = $event.Id
                'EventTimestamp' = $event.TimeCreated
                'Message' = $event.Message -split "`r" | Select-Object -First 1
                }
            }
            elseif($event.id -eq "4776")
            {
                $obj = [PSCustomObject]@{
                'Client' = $client
                'User' = $event.Properties[1].Value
                'Target' = $event.Properties[2].Value
                'DomainController' = $event.MachineName
                'EventId' = $event.Id
                'EventTimestamp' = $event.TimeCreated
                'Message' = $event.Message -split "`r" | Select-Object -First 1
                }
            }
            elseIf($event.id -eq "4771")
            {
                $obj = [PSCustomObject]@{
                'Client' = $client
                'User' = $event.Properties[0].Value
                'Target' = $event.Properties[6].Value -split "`:" | Select-Object -Last 1
                'DomainController' = $event.MachineName
                'EventId' = $event.Id
                'EventTimestamp' = $event.TimeCreated
                'Message' = $event.Message -split "`r" | Select-Object -First 1
                }
            }

            $EventsTable.Add($obj)
        }        
    }
    $Eventstable | Sort-Object -Property Eventtimestamp -Descending | export-csv -Path $eventsExportpath".csv" -NoTypeInformation
}
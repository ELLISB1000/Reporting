#Security Logs
$DomainControllers = "GR2GPCMAFSV01", "SL1GPCMAFSV01"
$SecurityEventTable = New-Object 'System.Collections.Generic.List[System.Object]'

Foreach($DC in $DomainControllers)
        {   
            $LockedOutEvents = Get-WinEvent -ComputerName $DC -FilterHashtable @{LogName='Security';Id=411} -ErrorAction Stop | Sort-Object -Property TimeCreated -Descending
        }

Foreach($Event in $LockedOutEvents)
        {
            $obj = [PSCustomObject]@{
            'User' = $eventobj.user
            'ADFS Server' = $event.MachineName
            'Timestamp' = $Event.timecreated
            'Event ID' = $Event.Id
            'Source IP' = $Event.Properties[4].Value
            'Account' = $event.Properties[2].Value -split "-" | Select-Object -First 1
            'Error Message' = $event.Properties[2].Value -split "-" | Select-Object -last 1
            }

            $SecurityEventTable.Add($obj)
        }
                      


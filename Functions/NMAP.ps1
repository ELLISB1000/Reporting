# ===============
# Load Config.xml
# ===============
# Load XML config file with client settings:
$Path = "C:\script\Reporting\Config.xml"
# load it into an XML object:
$xml = New-Object -TypeName XML
$xml.Load($Path)

# ==========
# Variables
# ==========

$exportpath = "C:\Script\Reporting\Export\"
$client = $xml.configuration.export.client

if ((Test-Path -Path "C:\Script\nmap-7.80\") -ne $true) 
    {
        write-host "Nmap folder missing from C:\Script\nmap-7.80\"
        Start-Sleep -Seconds 10
        exit
    }

$ips = $xml.configuration.Settings.nmap.IPrange -split "`,"
$NmapTable = New-Object 'System.Collections.Generic.List[System.Object]'
$NmapExportpath = $exportpath + $client + "Nmap"
Set-Location "C:\Script\nmap-7.80\"

# ==========
# Namp run
# ==========

If ($xml.configuration.Settings.nmap.Enabled -eq "True")
{
    Foreach ($ip in $ips)
    {
    Remove-Item -Path "C:\Script\nmap-7.80\scan.xml" -Force
    .\nmap.exe -sV -O --script ssl-cert $ip -v -oX "C:\Script\nmap-7.80\scan.xml" 

    # ===============
    # Load Scan.xml
    # ===============
    # Load XML config file with client settings:
    $sPath = "C:\Script\nmap-7.80\scan.xml"
    # load it into an XML object:
    $sxml = New-Object -TypeName XML
    $sxml.Load($sPath)

    # =======================
    # Build and export table
    # =======================

    $hosts = $sxml.nmaprun.host
  
    Foreach ($h in $hosts)
        {
            if ($h.status.state -eq "up")
            {
                $ports = $h.ports.port
                foreach ($p in $ports)
                {
                    $Obj = [PSCustomObject]@{
                    'Client' = $client
                    'Hostname' = $h.hostnames.hostname.name
                    'IP address' = $h.address.addr
                    'OS Match' = $h.os.osmatch.name[0]
                    'OS Accuracy' = $h.os.osmatch.accuracy[0]
                    'Uptime' = $h.uptime.seconds
                    'State' = $h.status.state 
                    'State Reason' = $h.status.reason
                    'PortId' = $p.portid
                    'Port Protocol' = $p.protocol
                    'Port State' = $p.state.state
                    'Port Service' = $p.service.name
                    'Port cert Name' = if ($null -ne $p.script.table){($p.script.table.elem | Where-Object {$_.key -eq 'CommonName'}).'#text' | Select-Object -First 1}ELSE{""}
                    'Port cert Issuer' = if ($null -ne $p.script.table){($p.script.table.elem | Where-Object {$_.key -eq 'CommonName'}).'#text' | Select-Object -First 1}ELSE{""}
                    'Port cert Type' = if ($null -ne $p.script.table){($p.script.table | Where-Object {$_.key -eq 'pubkey'}).elem.'#text'| Select-Object -Last 1}ELSE{""}
                    'Port cert Valid Before' = if ($null -ne $p.script.table){($p.script.table | Where-Object {$_.key -eq 'validity'}).elem.'#text' | Select-Object -First 1}ELSE{""}
                    'Port cert Valid After' = if ($null -ne $p.script.table){($p.script.table | Where-Object {$_.key -eq 'validity'}).elem.'#text'| Select-Object -Last 1 }ELSE{""}
                    }
                    $nmapTable.Add($obj)
                }
            }
            ELSE
            {
                    $Obj = [PSCustomObject]@{
                    'Client' = $client
                    'Hostname' = ""
                    'IP address' = $h.address.addr
                    'OS Match' = ""
                    'OS Accuracy' = ""
                    'Uptime' = ""
                    'State' = $h.status.state 
                    'State Reason' = $h.status.reason
                    'PortId' = ""
                    'Port Protocol' = ""
                    'Port State' = ""
                    'Port Service' = ""
                    'Port cert Name' = ""
                    'Port cert Issuer' = ""
                    'Port cert Type' = ""
                    'Port cert Valid Before' = ""
                    'Port cert Valid After' = ""
                    }
                    $nmapTable.Add($obj)
            }
            
        }
    }

    $nmapTable | export-csv -Path $nmapExportpath".csv" -NoTypeInformation -Force  
}
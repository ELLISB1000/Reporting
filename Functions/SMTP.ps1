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

# =======================
# SMTP Relay Export
# =======================

If ($xml.configuration.Settings.SMTPRelay.Enabled -eq "True")
{
    $SMTPExportpath = $exportpath + $client + "SMTPRelay"
    $SMTPTable = New-Object 'System.Collections.Generic.List[System.Object]'
    $SMTPs = $xml.configuration.Settings.SMTPRelay.Servers -split "`,"
    Foreach ($SMTP in $SMTPs)
    {
        $S = Get-WmiObject IISSmtpServerSetting -namespace "ROOT\MicrosoftIISv2" -cn $SMTP | Where-Object { $_.name -like "SmtpSVC/1"}
        $RelayIPList = @()    

        For ($CurrentIt = 0; $CurrentIt -lt $S.RelayIpList.Count; $CurrentIt += 4) 
        {
            $EndIndex = $CurrentIt + 3
            $RelayIPList += $S.RelayIpList[$CurrentIt..$EndIndex] -join "."
        }
        $RelayIPList = $RelayIPList[20..($RelayIPList.Count - 1)]
        $RelayIPList = $RelayIPList | Where-Object {($_ -notlike "255.255.*") -and ($_ -notlike "*.0.0.*")}
        Foreach ($relayip in $RelayIPList)
        {
            $obj = [PSCustomObject]@{
            'Client' = $client
            'IPAddress' = $relayip
            'HostName' = $([System.Net.Dns]::gethostentry($RelayIP).HostName)
            'Server' = $smtp
            }
        $SMTPTable.Add($obj)
        }  
    }    
    $SMTPTable | export-csv -Path $SMTPExportpath".csv" -NoTypeInformation
}
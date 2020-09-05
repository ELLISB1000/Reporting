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

# Import CA Powershell Module
If ($xml.configuration.Settings.CA.Enabled -eq "True")
{
    import-Module -Name pspki
}

# ==========
# Variables
# ==========

$exportpath = "C:\Script\Reporting\Export\"

$client = $xml.configuration.export.client

# =======================
# CA Export
# =======================

If ($xml.configuration.Settings.CA.Enabled -eq "True")
{
    $CAExportpath = $exportpath + $client + "CA"
    $CAs = $xml.configuration.Settings.CA.Servers -split "`,"
    $caTable = New-Object 'System.Collections.Generic.List[System.Object]'
    Foreach ($caserver in $CAs)
    {
    $c = Get-IssuedRequest -CertificationAuthority $CAserver   
    Foreach ($caobj in $c)
        {
        $CertObj = [PSCustomObject]@{
            'Client' = $client
            'Name' = $caobj.CommonName
            'Valid From' = $caobj.notbefore
            'Valid To' = $caobj.notafter
            'Template' = $caobj.CertificateTemplate
            'Serial Number' = $caobj.SerialNumber
            'Config' = $caobj.configstring
            
            }
        $caTable.Add($certobj)
        }

    }
    $caTable | export-csv -Path $CAExportpath".csv" -NoTypeInformation
}




        
        
        


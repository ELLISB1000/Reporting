Write-Host "Starting Install..."
Write-Host "Checking Powershell Version"
If ($PSVersionTable.PSVersion.major -lt "5")
    {
        Write-host "Powershell version is not later than version 5.0. Current Version $($psversiontable.PSVersion)"
        Write-host "Install the following: https://www.microsoft.com/en-us/download/confirmation.aspx?id=54616"
        Start-Sleep -Seconds 60
        Exit
    }
    ELSE
    {  
        Write-Host "Copying files"
        $updatepath = "C:\script\Reporting\"
        if ((Test-Path -Path $updatepath -ErrorAction SilentlyContinue) -ne $true)
            {
                New-Item -Path $updatepath -ItemType directory -ErrorAction SilentlyContinue
            }
        $items = Get-ChildItem $PSScriptRoot -Recurse
        Foreach ($item in $items)
            {
                If ($item.name -ne "Installer")
                    {
                    $End = $item.VersionInfo.FileName -split "`Reporting\\" | Select-Object -last 1
                    $exportpath = "C:\script\Reporting\" + $end
                    Copy-Item $item.FullName -Destination $exportpath -force -confirm:$false -ErrorAction SilentlyContinue   
                    }
            }
        Write-Host "Completed" -fore Green

        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        Write-host "Checking Powershell Modules Are installed"
        # Import Active Directory
            try
            {
                Write-Host "Installing AD Powershell Module"
                Get-Module -ListAvailable -Name "activedirectory"
            }
            Catch
            {
                Write-host "Unable to Import Active Directory Powershell module" -ForegroundColor Red
            }


        # Import CA Powershell Module
            try
            {
                $CAMod = Get-Module -ListAvailable -Name "pspki"        
                If ($null -eq $CAMod)
                {
                    Write-Host "Installing CA Powershell Module"
                    Install-Module -Name pspki -Force
                }
            }
            Catch
            {
                Write-host "Unable to Import CA Powershell module" -fore Red
            }

        Write-Host "Completed" -fore Green
        Write-Host "Creating Scheduled Task"
        try
        {
            Import-Module ScheduledTasks -Force
            $action = New-ScheduledTaskAction -Execute powershell.exe -Argument "-file C:\script\reporting\reporting.ps1"
            $trigger = New-ScheduledTaskTrigger -Daily -At "01:00"
            $user = Read-Host "Task Username:"
            $password = Read-Host "Task Password:"
            $settings = New-ScheduledTaskSettingsSet -MultipleInstances Parallel
            Register-ScheduledTask -TaskName "Reporting" -Description "BI Reporting - Contact Ellis Barrett (A365)" -TaskPath "\" -Action $action -Trigger $trigger -Settings $settings -User $user -Password $password -RunLevel Highest
        }
        Catch
        {
            Write-Host "Unable to create Scheduled Task" -fore Red 
        }
        Write-Host "Completed" -fore Green
        Write-Host "Set-up Config file"
        Pause
    }






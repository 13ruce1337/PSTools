#####
# Alfred Pennyworth
# Makes troubleshooting basic tasks easier.
# Author: Trevor Cooper
# Commmands - install, uninstall, reinstall, update
#####

# This param must be at the top of the script. It defines the inputs.
param($command, $application)

# Helper functions
function UninstallApplication {
    param($name)

    $Apps = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -like "*$name*"}

    foreach ($app in $apps) {
        $App.Uninstall()
    }
}
function InstallMSI {
    param($msi)

    Write-Host "Installing $msi" -ForegroundColor Yellow
    Start-Process msiexec -ArgumentList "/i $msi /qn" -Wait
    Write-Host "Installation complete." -ForegroundColor Green
}
function ODTExec {
    param($config)

    $odt_url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17126-20132.exe"
    $odt = "officedeploymenttool_17126-20132.exe"
    $config_name = "office_config.xml"
    $processes = Get-CimInstance -ClassName Win32_Process -Filter "ExecutablePath like '%Microsoft Office%'"

    Write-Host "Closing Office Applications."
    $processes | Invoke-CimMethod -MethodName "Terminate"
    Write-Host "Creating configuration file."
    $config | New-Item -Path $env:TEMP -Name $config_name -Force
    Write-Host "Downloading ODT from Microsoft."
    Invoke-WebRequest -Uri $odt_url -OutFile $env:TEMP\$odt
    Write-Host "Extracting the ODT then running with the following config:"
    Write-Host $config
    Write-Host "This may take some time..." -ForegroundColor Yellow
    Start-Process -FilePath $env:TEMP\$odt -ArgumentList "/norestart /passive /quiet /extract:$env:TEMP" -Wait
    Start-Process "$env:TEMP\setup.exe" -ArgumentList "/configure $env:TEMP\$config_name" -Wait
}
function DownloadTeams {
    
    $teams_url = "https://teams.microsoft.com/downloads/desktopurl?env=production&plat=windows&arch=x64&managedInstaller=true&download=true"
    $teams = "Teams_windows_x64.msi"
    if (-Not (Test-Path -Path $env:TEMP\$teams)) {
        Invoke-WebRequest -Uri $teams_url -OutFile $env:TEMP\$teams
    }
    return "$env:TEMP\$teams"
}

# Specific application functions
# Office functions 
function InstallOffice {
    $config = '<Configuration> 
    <Add OfficeClientEdition="64" Channel="Current">
        <Product ID="O365ProPlusRetail" > 
            <Language ID="en-us" />        
        </Product> 
    </Add> 
    <Display Level="None" AcceptEULA="TRUE" />
</Configuration>'

    Write-Host "Installing Office applications." -ForegroundColor Cyan
    ODTExec $config
    Write-Host "Office Applications have been installed." -ForegroundColor Green
}
function UninstallOffice {
    $config = '<Configuration>
    <Remove All="TRUE"/>
    <Display Level="None" AcceptEULA="TRUE" />
</Configuration>'

    Write-Host "Uninstalling Office applications." -ForegroundColor Cyan
    ODTExec $config
    Write-Host "Completed Office removal." -ForegroundColor Green
}
function ReinstallOffice {
    Write-Host "Reinstalling Office applications." -ForegroundColor Cyan
    UninstallOffice
    InstallOffice
    Write-Host "Completed the uninstallation and reinstallation process for Office applications." -ForegroundColor Green
}

# Teams functions
function InstallTeams {
    Write-Host "Installing Teams."
    $installer = DownloadTeams
    Start-Process -FilePath "$installer" -ArgumentList "/qn" -Wait
    Write-Host "Completed installing Teams."
}
function UninstallTeams {
    Write-Host "Uninstalling Teams." -ForegroundColor Cyan
    Stop-Process -Name "ms-teams" -Force -ErrorAction SilentlyContinue
    $installer = DownloadTeams
    try {
        $process = Start-Process -FilePath "$installer" -ArgumentList "/uninstall /s" -PassThru -Wait -ErrorAction STOP
        if ($process.ExitCode -ne 0)
        {
            Write-Error "Uninstallation failed with exit code  $($process.ExitCode)."
        }
    }
    catch {
        Write-Error $_.Exception.Message
        Exit
    }
    UninstallApplication "Teams Machine-Wide Installer"
    Write-Host "Completed uninstalling Teams." -ForegroundColor Green
}
function ReinstallTeams {
    Write-Host "Reinstalling Teams." -ForegroundColor Cyan
    UninstallTeams
    InstallTeams
    Write-Host "Completed uninstalling and reinstalling Teams."
}   

# Functions for initial input arguments
function Install {
    switch ($application)
    {
        "office" {
            InstallOffice
        }
        "teams" {
            InstallTeams
        }
    }
}
function Uninstall {
    switch ($application)
    {
        "office" {
            UninstallOffice
        }
        "teams" {
            UninstallTeams
        }
        Default {
            UninstallApplication $application
        }
    }
}
function Reinstall {
    switch ($application)
    {
        "office" {
            ReinstallOffice
        }
        "teams" {
            ReinstallTeams
        }
    }
}

# Initial arguments switch
switch ($command)
{
    "install" {Install}
    "uninstall" {Uninstall}
    "reinstall" {Reinstall}
}
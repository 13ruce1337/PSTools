<#
.SYNOPSIS
Alfred Pennyworth

.DESCRIPTION
VERSION 0.3

.INPUTS
COMMANDS
    install
    uninstall
    reinstall
    nuke
    update
    troubleshoot
    sd
APPLICATIONS
    office
    teams
    visio
    project
    onenote
    adobedc

.OUTPUTS
This command currently doesn't output any data but will display status.

.EXAMPLE
.\alfred install office

.LINK
https://github.com/13ruce1337/pstools
#>

# This param must be at the top of the script. It defines the inputs.
param($command, $application)

$ProgressPreference = 'SilentlyContinue'

# Helper functions
function AdminCheck {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
function UninstallApplication {
    param($name)

    $Apps = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -like "*$name*"}

    foreach ($app in $apps) {
        $App.Uninstall()
    }
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
    if (-not (Test-Path -Path $env:TEMP\$odt)) {
        Write-Host "Downloading ODT from Microsoft..."
        Invoke-WebRequest -Uri $odt_url -OutFile $env:TEMP\$odt
    } 
    Write-Host "Extracting the ODT then running with the following config:"
    Write-Host $config
    Write-Host "This may take some time..." -ForegroundColor Yellow
    Start-Process -FilePath $env:TEMP\$odt -ArgumentList "/norestart /passive /quiet /extract:$env:TEMP" -Wait
    Start-Process "$env:TEMP\setup.exe" -ArgumentList "/configure $env:TEMP\$config_name" -Wait
}
function ODTInstallConfig {
    param($product)

    $config = '<Configuration> 
    <Add OfficeClientEdition="64" Channel="Current">
        <Product ID="'+$product+'" > 
            <Language ID="en-us" />        
        </Product> 
    </Add> 
    <Display Level="None" AcceptEULA="TRUE" />
</Configuration>'

    Write-Host "Installing $product." -ForegroundColor Cyan
    ODTExec $config
    Write-Host "$product has been installed." -ForegroundColor Green
}
function ODTUninstallConfig {
    param($product)

    $config = '<Configuration>
    <Remove All="FALSE">
        <Product ID="'+$product+'" >
        </Product>
    </Remove>
    <Display Level="None" AcceptEULA="TRUE" />
</Configuration>'

    Write-Host "Uninstalling $product." -ForegroundColor Cyan
    ODTExec $config
    Write-Host "$product has been removed." -ForegroundColor Green
}
function DownloadAdobeDC {
    $url = "https://trials.adobe.com/AdobeProducts/APRO/Acrobat_HelpX/win32/Acrobat_DC_Web_x64_WWMUI.zip"
    $installer = "Acrobat_DC_Web_x64_WWMUI.zip"
    if (-not (Test-Path -Path $env:TEMP\$installer)) {
        Write-Host "Downloading Adobe STD..."
        Invoke-WebRequest -Uri $url -OutFile $env:TEMP\$installer
        Expand-Archive -Path $env:TEMP\$installer -DestinationPath $env:TEMP\ -Force
    }
    return "$env:temp\Adobe Acrobat\AcroPro.msi"
}

# Specific application functions
# Office functions 
function InstallOffice {
    ODTInstallConfig("O365ProPlusRetail")
}
function UninstallOffice {
    ODTUninstallConfig("O365ProPlusRetail")
}
function NukeOffice {
    $config = '<Configuration>
    <Remove All="TRUE"/>
    <Display Level="None" AcceptEULA="TRUE" />
</Configuration>'

    Write-Host "Uninstalling all Office applications." -ForegroundColor Cyan
    ODTExec $config
    Write-Host "Completed Office removal." -ForegroundColor Green
}
function ReinstallOffice {
    Write-Host "Reinstalling Office applications." -ForegroundColor Cyan
    UninstallOffice
    InstallOffice
    Write-Host "Completed the uninstallation and reinstallation process for Office applications." -ForegroundColor Green
}

# Visio functions
function InstallVisio {
    ODTInstallConfig("VisioPro2021Retail")
}
function UninstallVisio {
    ODTUninstallConfig("VisioPro2021Retail")
}
function ReinstallVisio {
    UninstallVisio
    InstallVisio
}

# Project functions
function InstallProject {
    ODTInstallConfig("ProjectPro2021Retail")
}
function UninstallProject {
    ODTUninstallConfig("ProjectPro2021Retail")
}
function ReinstallProject {
    UninstallProject
    InstallProject
}

# OneNote functions
function InstallOneNote {
    ODTInstallConfig("OneNoteRetail")
}
function UninstallOneNote {
    ODTUninstallConfig("OneNoteRetail")
}
function ReinstallOneNote {
    UninstallOneNote
    InstallOneNote
}

# Teams functions
function InstallTeams {
    if (AdminCheck) {
        Write-Host "This specific task cannot be done as the administrator" -ForegroundColor Red
        return
    }
    $url = "https://go.microsoft.com/fwlink/?linkid=2196106&clcid=0x409&culture=en-us&country=us"
    $installer = "teams_installer.msix"
    Write-Host "Installing Teams."
    if (-not (Test-Path -Path $env:TEMP\$installer)) {
        Write-Host "Downloading Teams..."
        Invoke-WebRequest -uri $url -OutFile "$env:TEMP\$installer"
    }
    Add-AppPackage -path "$env:TEMP\$installer"
    Write-Host "Completed installing Teams." -ForegroundColor Green
}
function UninstallTeams {
    if (AdminCheck) {
        Write-Host "This specific task cannot be done as the administrator" -ForegroundColor Red
        return
    }
    Write-Host "Uninstalling Teams." -ForegroundColor Cyan
    Stop-Process -Name "ms-teams" -Force -ErrorAction SilentlyContinue
    try {
        Get-AppxPackage MicrosoftTeams | Remove-AppxPackage
        Get-AppxPackage MSTeams | Remove-AppxPackage
    }
    catch {
        Write-Error $_.Exception.Message
        Exit
    }
    UninstallApplication "Teams Machine-Wide Installer"
    Write-Host "Completed uninstalling Teams." -ForegroundColor Green
}
function ReinstallTeams {
    if (AdminCheck) {
        Write-Host "This specific task cannot be done as the administrator" -ForegroundColor Red
        return
    }
    Write-Host "Reinstalling Teams." -ForegroundColor Cyan
    UninstallTeams
    InstallTeams
    Write-Host "Completed uninstalling and reinstalling Teams." -ForegroundColor Green
}

# Adobe functions
function InstallAdobeDC {
    Write-Host "Installing Adobe DC." -ForegroundColor Cyan
    $msi = DownloadAdobeDC
    try {
        Start-Process "msiexec.exe" -Argument "/i `"$msi`"  /qn" -Verbose -Wait
    }
    catch {
        Write-Host $_
    }
    Write-Host "Completed installing Adobe DC." -ForegroundColor Green
}
function UninstallAdobeDC {
    Write-Host "Uninstalling Adobe DC." -ForegroundColor Cyan
    $msi = DownloadAdobeDC
    try {
        Start-process "msiexec.exe" -Argument "/x `"$msi`" /qn" -Verbose -Wait
    }
    catch {
        Write-Host $_
    }
    Write-Host "Adobe DC has been uninstalled." -ForegroundColor Green
}
function ReinstallAdobeDC {
    UninstallAdobeDC
    InstallAdobeDC
}

# Troubleshooting functions
function TroubleshootNetwork {
    Write-Host "Attempting basic network fixes." -ForegroundColor Cyan
    ipconfig /release
    ipconfig /renew
    ipconfig /flushdns
    ipconfig /registerdns
    Write-Host "Finished basic networking fixes, below is the latest IP info:" -ForegroundColor Green
    ipconfig /all 
}
function TroubleshootWindows {
    Write-Host "Attempting basic Windows fixes." -ForegroundColor Cyan
    UpdateWindows
    Write-Host "Optimizing the OS volume."
    Optimize-Volume -DriveLetter C -Analyze -Confirm -Defrag -ReTrim -SlabConsolidate -Verbose
    Write-Host "Cleaning up the OS image."
    DISM /Online /Cleanup-Image /RestoreHealth
    Write-Host "Running System File Checker."
    sfc /scannow
    Write-Host "Finished basic Windows fixes." -ForegroundColor Green
}

# Update functions
function UpdateWindows {
    Write-Host "Updating Windows." -ForegroundColor Cyan
    if (-not (Get-PackageProvider -ListAvailable -Name "NuGet" -ErrorAction "Ignore")) {
        Write-Host "NuGet is not installed, installing now..." -ForegroundColor Red
		Install-PackageProvider -Name "NuGet" -Force
	}
	if (-not (Get-InstalledModule -Name "PSWindowsUpdate" -ErrorAction "Ignore")) {
    	Write-Host "PSWindowsUpdate is not installed, installing now..." -ForegroundColor Red
	    Install-Module -Name "PSWindowsUpdate" -Force
    }
    Write-Host "Checking for Windows updates..."
    if (Get-WindowsUpdate) {
    	Write-Host "Installing updates..."
        Get-WindowsUpdate -AcceptAll -Download -Install
        Write-Host "Updates have been installed. You should reboot the system." -ForegroundColor Green
    } else {
        Write-Host "Windows is up to date." -ForegroundColor Green
    }
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
        "visio" {
            InstallVisio
        }
        "project" {
            InstallProject
        }
        "onenote" {
            InstallOneNote
        }
        "adobedc" {
            InstallAdobeDC
        }
        default {
            Write-Host "I do not have the ability to install that application." -ForegroundColor Red
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
        "visio" {
            UninstallVisio
        }
        "project" {
            UninstallProject
        }
        "onenote" {
            UninstallOneNote
        }
        "adobedc" {
            UninstallAdobeDC
        }
        Default {
            Write-Host "I do not have a way to uninstall that application." -ForegroundColor Red
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
        "onenote" {
            ReinstallOneNote
        }
        "project" {
            ReinstallProject
        }
        "visio" {
            ReinstallVisio
        }
        "adobedc" {
            ReinstallAdobeDC
        }
    }
}
function Nuke {
    switch ($application)
    {
        "office" {
            NukeOffice
        }
    }
}
function Troubleshoot {
    switch ($application)
    {
        "network" {
            TroubleshootNetwork
        }
        "windows" {
            TroubleshootWindows
        }
    }
}
function Update {
    switch ($application) {
        "windows" { 
            UpdateWindows
         }
        Default {}
    }
    
}
function SelfDestruct {
    Remove-Item $PSCommandPath -Force 
}
# Initial argument (verb) switch
switch ($command)
{
    "install" {Install}
    "uninstall" {Uninstall}
    "reinstall" {Reinstall}
    "nuke" {Nuke}
    "troubleshoot" {Troubleshoot}
    "update" {Update}
    "sd" {SelfDestruct}
}
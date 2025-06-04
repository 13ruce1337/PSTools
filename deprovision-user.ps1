<#
.SYNOPSIS
    Deprovisions User

.DESCRIPTION
    IMPORTANT: You must have RSAT installed and you must run this from an elevated terminal
    IMPORTANT: This script expects a hybrid AD/Entra environment

.NOTES
    Version:        1.7
    Author:         Trevor Cooper
    Creation Date:  10/17/24
    Purpose/Change: Initial script development
    Mod 5/7/25: Updated QOL, comments, function for connecting, module check/install
    Mod 5/14/25: Added function for main components, added abillity to take a CSV

.EXAMPLE
    .\deprovision-user.ps1 -user trevor.cooper -manager tuan.pham -date "10/10/24"
    .\deprovision-user.ps1

.LINK
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

[cmdletbinding()]
param(
    [string]$user,
    [string]$manager,
    [string]$date
)

#Preferences
$ProgressPreference = "SilentlyContinue"

#Assemblies
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")

#Script Variables
$dependency_modules = @(
    "Graph",
    "Microsoft.Online.SharePoint.Powershell",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Users.Actions",
    "ExchangeOnlineManagement",
    "AzureAD"
)
$script:id
$script:managerid
$script:ooom
$script:multiple_users = $false
$script:user_list
$sharepoint_site = "https://YOURSITE.sharepoint.com/"

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
[float]$script:script_version = 1.6

#Logging
[string]$script:log_path = "C:\Windows\Temp"
[string]$script:log_name = "deprovision_user.log"
[string]$script:log_file = Join-Path -Path $script:log_path -ChildPath $script:log_name
Start-Transcript -Force -IncludeInvocationHeader -Path $script:log_file

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Test-SharePointSiteVariable {
    <#
    .SYNOPSIS
        Checks if the SharePoint site variable has been updated.

    .DESCRIPTION
        This function verifies if the $sharepoint_site variable still contains the placeholder value.
        If it does, it outputs an error message and exits the script.
    #>
    if ($script:sharepoint_site -like "*YOURSITE.sharepoint.com*") {
        Write-Error "Error: Please update the `$sharepoint_site` variable in the script with your actual SharePoint site URL."
        Write-Host "Example: `$sharepoint_site = 'https://yourcompany.sharepoint.com/'`" -ForegroundColor Yellow
        Exit 1
    }
}

function connect_services {
    Write-Host "Due to MFA, we cannot store credentials. So you will be profusely asked for them while connecting to the different platforms." -ForegroundColor Yellow
    Connect-Graph -Scopes User.ReadWrite.All, Organization.Read.All
    Connect-ExchangeOnline
    Connect-SPOService -Url $sharepoint_site
    Connect-AzureAD
}
function initialize_vars {
    param(
        [string]$user,
        [string]$manager,
        [string]$date
    )
    $script:id = Get-ADUser $user -Properties *
    $script:managerid = Get-ADUser $manager -Properties *
    $script:ooom = "As of $date, " + $script:id.Name + " no longer works for this organization. Please direct all inquiries to " + $script:managerid.Name + " - " + $script:managerid.EmailAddress
}
function mod_check($mod) {
    if (!(Get-Module -ListAvailable -Name $mod)) {
        Write-Host "$mod is missing, instaling now..."
        Install-Module $mod
    }
    else {
        Write-Host "$mod is installed."
    }
}
function dependency_check {
    #check for admin
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not ($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))){
        throw "You must run this script as admin. Right click and run Powershell.exe as admin."
    }
    #check for RSAT
    $rsat = Get-WindowsCapability -Online | Where-Object Name -match "Rsat.ServerManager.Tools*"
    if ($rsat.State -eq "NotPresent") {
        throw "You must install RSAT to use this script. Switch to your DA account and install the feature."
    }
    else {
        Import-Module ActiveDirectory
    }
    foreach ($mod in $dependency_modules) {
        mod_check $mod
        if (! (Get-Module $mod)) {
            Import-Module $mod
        }
    }
}
function remove_groups {
    $aduser = Get-AzureADUser -All $true | Where-Object {$_.UserPrincipalName -eq $script:id.EmailAddress}
    $groups = Get-AzureADUserMembership -ObjectId $aduser.ObjectId
    #find and remove from DLs
    Get-DistributionGroup -ResultSize unlimited | Where-Object { (Get-DistributionGroupMember $_.Identity).PrimarySmtpAddress -contains $script:id.EmailAddress } | Remove-DistributionGroupMember -Member $script:id.Name
    #find and remove from Azure groups
    foreach($group in $groups) {
        try {
            if($group.Name -ne "All Users") {
                Remove-AzureADGroupMember -ObjectId $group.ObjectId -MemberId $aduser.ObjectId
            }
        }
        catch {
            Write-Host "Error removing $($aduser.DisplayName) from $($group.DisplayName): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}
function remove_licenses {
    $licenses = Get-MgUserLicenseDetail -UserId $script:id.EmailAddress
    $licenses_filtered = @()
    foreach($license in $licenses) {
        $licenses_filtered += $license.SkuId
    }
    Set-MgUserLicense -UserId $script:id.EmailAddress -RemoveLicenses $licenses_filtered -AddLicenses @{}
}
function arg_check {
    if (!$args) {
        $script:multiple_users = $true
    } elseif (!($args -eq 2)) {
        Write-Error "Make sure to provide a user, manager and date for arguments"
        Stop-Transcript
        exit 1
    }
}
function deprovision_ad {
    Write-Host "Deprovisioning $user in AD"
    #Active Directory
    #hide from exchange address list
    Set-ADUser $script:id.DistinguishedName -Add @{msExchHideFromAddressLists="TRUE"}
    #remove from groups in AD
    Get-AdPrincipalGroupMembership -Identity $script:id.DistinguishedName | Where-Object -Property Name -Ne -Value 'Domain Users' | Remove-AdGroupMember -Members $user
    #remove manager from AD
    Set-ADUser $script:id -clear manager
    #remove direct reports from AD
    Get-ADUser $script:id -Properties DirectReports | Select-Object -ExpandProperty DirectReports | Set-ADUser -Clear Manager
    #remove employee type
    Set-ADUser $script:id.DistinguishedName -Clear employeeType
}
function deprovision_msol {
    Write-Host "Deprovisioning $user in MSOL"

    #Exchange Online
    #set out of office
    Set-MailboxAutoReplyConfiguration -Identity $script:id.EmailAddress -AutoReplyState Enabled -InternalMessage $script:ooom -ExternalMessage $script:ooom -ExternalAudience All
    #convert to shared mailbox
    Set-Mailbox -Identity $script:id.EmailAddress -Type Shared
    #give access to manager
    Add-MailboxPermission -Identity $script:id.EmailAddress -User $script:managerid.EmailAddress -AccessRights FullAccess -InheritanceType All
    #remove from 365 groups
    remove_groups
    #remove devices
    (Get-MobileDevice -Mailbox $script:id.EmailAddress).DistinguishedName | Remove-MobileDevice -Confirm:$false
    #remove 365 licenses
    remove_licenses
}
function get_csv {
    $file_dialog = New-Object System.Windows.Forms.OpenFileDialog
    $file_dialog.Filter = "CSV (*.csv) | *.csv"
    $file_dialog.ShowDialog()
    $path = $file_dialog.Filename
    $script:user_list = Import-CSV -Path $path
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

#check for dependcies
try {
    dependency_check
}
catch {
    Write-Error $_
    Stop-Transcript
    exit 1
}

# Check if $sharepoint_site has been updated
Test-SharePointSiteVariable

connect_services
arg_check

if ($multiple_users) {
    Write-Host "Running for multiple users"
    try {
        get_csv
    }
    catch {
        Write-Error "Could not load the file"
        Throw $_
        Stop-Transcript
    }
    foreach ($emp in $script:user_list) {
        Write-Host "Deprovisioning" $emp.user
        initialize_vars -user $emp.user -manager $emp.manager -date $emp.date
        deprovision_ad
        deprovision_msol
    }
} else {
    Write-Host "Running for a single user"
    initialize_vars -user $user -manager $manager -date $date
    deprovision_ad
    deprovision_msol
}

Write-Host "Disconnecting from services"
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-AzureAD -Confirm:$false
Disconnect-Graph
Disconnect-SPOService

Write-Host "Transcript can be found: " $script:log_file
Stop-Transcript
Exit 0
function Disable-M365User {
    <#
    .SYNOPSIS
        Deprovisions a user in Microsoft 365 (Exchange Online + Entra ID).

    .DESCRIPTION
        - Sets an out-of-office reply pointing to the manager
        - Converts the mailbox to a shared mailbox
        - Grants the manager FullAccess to the mailbox
        - Removes the user from cloud distribution lists
        - Removes the user from cloud-managed Entra/M365 groups
        - Removes mobile device partnerships
        - Removes directly assigned licenses

        Groups synced from on-prem AD are skipped (handle them with Disable-ADUserAccount).
        Dynamic groups and group-based licenses cannot be removed directly and are skipped/warned.

        IMPORTANT: You must already be connected to Microsoft Graph with the required scopes:
            Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All, Group.Read.All, GroupMember.ReadWrite.All

        An existing Exchange Online session is reused; if there is none, you will be prompted to sign in.

    .PARAMETER UserPrincipalName
        UPN of the user being deprovisioned.

    .PARAMETER ManagerUserPrincipalName
        UPN of the manager who will receive mailbox access and be named in the out-of-office message.

    .PARAMETER Date
        Date to show in the out-of-office message. Defaults to today.

    .PARAMETER ExcludeGroup
        Display names of groups to never remove the user from.

    .NOTES
        Version:        2.1
        Author:         Trevor Cooper
        Creation Date:  10/17/24
        Mod 5/7/25:  Updated QOL, comments, function for connecting, module check/install
        Mod 5/14/25: Added function for main components, added ability to take a CSV
        Mod 2.0:     Split from Disable-User into M365-only function, replaced AzureAD with Graph
        Mod 2.1:     Use Assert-ExchangeOnlineConnection for the Exchange Online session

    .EXAMPLE
        Disable-M365User -UserPrincipalName trevor.cooper@contoso.com -ManagerUserPrincipalName tuan.pham@contoso.com -Date "10/10/24"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('UPN')]
        [string]$UserPrincipalName,

        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('ManagerUPN')]
        [string]$ManagerUserPrincipalName,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$Date = (Get-Date).ToShortDateString(),

        [string[]]$ExcludeGroup = @('All Users')
    )

    begin {
        Assert-Module -Name @(
            'Microsoft.Graph.Authentication',
            'Microsoft.Graph.Users',
            'Microsoft.Graph.Users.Actions',
            'Microsoft.Graph.Groups',
            'ExchangeOnlineManagement'
        )

        if (-not (Get-MgContext)) {
            throw "Not connected to Microsoft Graph. Run: Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All, Group.Read.All, GroupMember.ReadWrite.All"
        }

        Assert-ExchangeOnlineConnection
    }

    process {
        $mgUser    = Get-MgUser -UserId $UserPrincipalName -Property Id, DisplayName, UserPrincipalName -ErrorAction Stop
        $mgManager = Get-MgUser -UserId $ManagerUserPrincipalName -Property Id, DisplayName, Mail, UserPrincipalName -ErrorAction Stop
        $managerEmail = if ($mgManager.Mail) { $mgManager.Mail } else { $mgManager.UserPrincipalName }
        $upn = $mgUser.UserPrincipalName

        Write-Host "Deprovisioning $upn in Microsoft 365" -ForegroundColor Cyan

        #----------------------------------------------------------[Exchange Online]----------------------------------------------------------

        # Out of office
        Write-Host "  Setting out-of-office reply"
        $message = "As of $Date, $($mgUser.DisplayName) no longer works for this organization. Please direct all inquiries to $($mgManager.DisplayName) - $managerEmail"
        Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled -InternalMessage $message -ExternalMessage $message -ExternalAudience All -ErrorAction Stop

        # Convert to shared mailbox (must happen before removing licenses)
        Write-Host "  Converting mailbox to shared"
        Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop

        # Grant manager access
        Write-Host "  Granting $($mgManager.UserPrincipalName) FullAccess"
        Add-MailboxPermission -Identity $upn -User $mgManager.UserPrincipalName -AccessRights FullAccess -InheritanceType All -ErrorAction Stop | Out-Null

        # Remove from cloud distribution lists / mail-enabled security groups
        Write-Host "  Removing from distribution lists"
        $mailbox = Get-Mailbox -Identity $upn -ErrorAction Stop
        $escapedDn = $mailbox.DistinguishedName -replace "'", "''"
        $distributionGroups = Get-DistributionGroup -ResultSize Unlimited -Filter "Members -eq '$escapedDn'"
        foreach ($dl in $distributionGroups) {
            if ($dl.IsDirSynced) {
                Write-Host "    Skipping $($dl.DisplayName) (synced from AD)"
                continue
            }
            try {
                Remove-DistributionGroupMember -Identity $dl.Identity -Member $mailbox.Identity -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                Write-Host "    Removed from $($dl.DisplayName)"
            }
            catch {
                Write-Warning "    Could not remove $upn from $($dl.DisplayName): $($_.Exception.Message)"
            }
        }

        # Remove mobile devices
        Write-Host "  Removing mobile devices"
        $devices = Get-MobileDevice -Mailbox $mailbox.Identity -ErrorAction SilentlyContinue
        foreach ($device in $devices) {
            try {
                Remove-MobileDevice -Identity $device.Identity -Confirm:$false -ErrorAction Stop
                Write-Host "    Removed $($device.FriendlyName) ($($device.DeviceModel))"
            }
            catch {
                Write-Warning "    Could not remove device $($device.Identity): $($_.Exception.Message)"
            }
        }

        #----------------------------------------------------------[Entra ID / Graph]----------------------------------------------------------

        # Remove from cloud-managed groups (M365 groups, cloud security groups)
        Write-Host "  Removing from Entra/M365 groups"
        $groups = Get-MgUserMemberOfAsGroup -UserId $mgUser.Id -All -Property Id, DisplayName, GroupTypes, MailEnabled, SecurityEnabled, OnPremisesSyncEnabled
        foreach ($group in $groups) {
            if ($ExcludeGroup -contains $group.DisplayName) {
                Write-Host "    Skipping $($group.DisplayName) (excluded)"
                continue
            }
            if ($group.OnPremisesSyncEnabled) {
                Write-Host "    Skipping $($group.DisplayName) (synced from AD)"
                continue
            }
            if ($group.GroupTypes -contains 'DynamicMembership') {
                Write-Host "    Skipping $($group.DisplayName) (dynamic membership)"
                continue
            }
            if ($group.MailEnabled -and ($group.GroupTypes -notcontains 'Unified')) {
                # Distribution lists / mail-enabled security groups are handled by Exchange above
                continue
            }
            try {
                Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $mgUser.Id -ErrorAction Stop
                Write-Host "    Removed from $($group.DisplayName)"
            }
            catch {
                Write-Warning "    Could not remove $upn from $($group.DisplayName): $($_.Exception.Message)"
            }
        }

        # Remove directly assigned licenses
        Write-Host "  Removing licenses"
        $skuIds = @((Get-MgUserLicenseDetail -UserId $mgUser.Id).SkuId)
        if ($skuIds.Count -gt 0) {
            try {
                Set-MgUserLicense -UserId $mgUser.Id -RemoveLicenses $skuIds -AddLicenses @() -ErrorAction Stop | Out-Null
            }
            catch {
                Write-Warning "    Could not remove one or more licenses (group-based licenses must be removed via group membership): $($_.Exception.Message)"
            }
        }
        else {
            Write-Host "    No licenses assigned"
        }
    }
}

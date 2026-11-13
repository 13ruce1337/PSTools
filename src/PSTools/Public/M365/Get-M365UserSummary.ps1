function Get-M365UserSummary {
    <#
    .SYNOPSIS
        Shows a brief overview of a Microsoft 365 user, similar to the user page
        in the Microsoft 365 admin center.

    .DESCRIPTION
        Collects the key facts about a user from Microsoft Graph and Exchange Online:

            Account         Enabled / blocked, user type, cloud or synced, created,
                            job title, department, office, phone, manager
            Password        Last set, age in days, never-expires policy
            Sign-in         Last interactive / non-interactive / successful sign-in
                            (requires Entra ID P1 and AuditLog.Read.All)
            Auth methods    Registered authentication (MFA) methods
            Licenses        Assigned SKUs, direct or group-based, assignment errors
            Groups / roles  Group memberships and activated directory roles
            Mailbox         Type, size, item count, last logon, archive, litigation
                            hold, hidden from GAL, forwarding
            Mailbox access  FullAccess, SendAs and SendOnBehalf delegates
            Out of office   State, schedule, audience and messages
            Mobile devices  Devices with an Exchange ActiveSync partnership

        Each section is retrieved independently; if one cannot be read a note is
        shown and the rest of the report continues.

        The report is written to the console. Use -PassThru to also return a
        report object.

    .PARAMETER Identity
        The user to look up: UPN, object ID or email address (primary or alias).
        Accepts pipeline input.

    .PARAMETER SkipMailbox
        Skip the Exchange Online sections (no Exchange Online connection is made).

    .PARAMETER PassThru
        Return a report object in addition to the on-screen report.

    .EXAMPLE
        Get-M365UserSummary -Identity trevor.cooper@contoso.com

    .EXAMPLE
        Get-M365UserSummary trevor.cooper@contoso.com -SkipMailbox

    .EXAMPLE
        $summary = Get-M365UserSummary -Identity trevor.cooper@contoso.com -PassThru
        $summary.MailboxPermissions

    .EXAMPLE
        'user1@contoso.com', 'user2@contoso.com' | Get-M365UserSummary

    .NOTES
        Requires the Microsoft.Graph.Authentication module and the delegated scopes:
            User.Read.All, GroupMember.Read.All, UserAuthenticationMethod.Read.All,
            AuditLog.Read.All
        Mailbox sections require the ExchangeOnlineManagement module.
        Existing Graph and Exchange Online sessions are reused.

        For users synced from on-premises AD, the password last set date reflects
        the last password change synced by Entra Connect (password hash sync).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('UserPrincipalName', 'UPN', 'PrimarySmtpAddress', 'Mail', 'Id')]
        [ValidateNotNullOrEmpty()]
        [string[]]$Identity,

        [switch]$SkipMailbox,

        [switch]$PassThru
    )

    begin {
        # Graph first so its MSAL assembly is loaded before Exchange Online's.
        $graphScopes = @(
            'User.Read.All'
            'GroupMember.Read.All'
            'UserAuthenticationMethod.Read.All'
            'AuditLog.Read.All'
        )
        Assert-MgGraphConnection -Scopes $graphScopes

        if (-not $SkipMailbox) {
            Assert-ExchangeOnlineConnection
        }

        #----------------------------------------------------------[Inner helpers]----------------------------------------------------------
        # Defined inline so they are not exported as part of the module.

        # Emits every item from a paged Graph collection.
        function Invoke-GraphPagedRequest {
            param([Parameter(Mandatory)][string]$Uri)
            while ($Uri) {
                $page = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
                if ($page.value) {
                    foreach ($entry in @($page.value)) { $entry }
                }
                $Uri = $page.'@odata.nextLink'
            }
        }

        # Converts a Graph date (DateTime or ISO 8601 string) to local time, or $null.
        function ConvertTo-LocalDateTime {
            param($Value)
            if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }

            if ($Value -is [datetime]) {
                $dt = $Value
            }
            else {
                try {
                    $dt = [datetime]::Parse(
                        [string]$Value,
                        [System.Globalization.CultureInfo]::InvariantCulture,
                        [System.Globalization.DateTimeStyles]::RoundtripKind
                    )
                }
                catch {
                    return $null
                }
            }

            if ($dt.Kind -eq [System.DateTimeKind]::Utc) { return $dt.ToLocalTime() }
            return $dt
        }

        # Strips HTML (as used in out-of-office messages) down to plain text.
        function ConvertTo-PlainText {
            param([string]$Html)
            if ([string]::IsNullOrWhiteSpace($Html)) { return $null }
            $text = $Html -replace '(?is)<(style|script|head)[^>]*>.*?</\1>', ''
            $text = $text -replace '(?i)<br\s*/?>|</p>|</div>', ' '
            $text = $text -replace '<[^>]+>', ''
            $text = [System.Net.WebUtility]::HtmlDecode($text)
            $text = ($text -replace '\s+', ' ').Trim()
            if ([string]::IsNullOrWhiteSpace($text)) { return $null }
            return $text
        }

        # Shortens long text for on-screen display.
        function Get-ShortText {
            param([string]$Text, [int]$Length = 200)
            if ([string]::IsNullOrEmpty($Text)) { return $Text }
            if ($Text.Length -le $Length) { return $Text }
            return $Text.Substring(0, $Length) + '...'
        }

        function Write-Section {
            param([string]$Title)
            Write-Host "`n--- $Title ---" -ForegroundColor Yellow
        }

        #----------------------------------------------------------[Static data]----------------------------------------------------------

        $emailPattern = '^[^@\s]+@[^@\s]+\.[^@\s]+$'

        # Permission entries for built-in / orphaned principals are noise in the report.
        $systemPrincipalPattern = '^(NT AUTHORITY\\|S-1-5-)'

        $userProperties = @(
            'id'
            'displayName'
            'userPrincipalName'
            'mail'
            'accountEnabled'
            'userType'
            'createdDateTime'
            'lastPasswordChangeDateTime'
            'passwordPolicies'
            'onPremisesSyncEnabled'
            'onPremisesLastSyncDateTime'
            'jobTitle'
            'department'
            'officeLocation'
            'usageLocation'
            'businessPhones'
            'mobilePhone'
            'licenseAssignmentStates'
        )
        $userSelect = $userProperties -join ','

        $authMethodNames = @{
            password                            = 'Password'
            phone                               = 'Phone (SMS/Voice)'
            email                               = 'Email'
            fido2                               = 'FIDO2 Security Key'
            microsoftAuthenticator              = 'Microsoft Authenticator'
            passwordlessMicrosoftAuthenticator  = 'Microsoft Authenticator (Passwordless)'
            softwareOath                        = 'Software OATH Token'
            hardwareOath                        = 'Hardware OATH Token'
            windowsHelloForBusiness             = 'Windows Hello for Business'
            temporaryAccessPass                 = 'Temporary Access Pass'
            x509Certificate                     = 'Certificate (CBA)'
            platformCredential                  = 'Platform Credential'
        }

        $mailboxProperties = @(
            'RecipientTypeDetails'
            'ForwardingAddress'
            'ForwardingSmtpAddress'
            'DeliverToMailboxAndForward'
            'GrantSendOnBehalfTo'
            'HiddenFromAddressListsEnabled'
            'LitigationHoldEnabled'
            'ArchiveGuid'
            'ArchiveStatus'
            'ProhibitSendQuota'
            'WhenMailboxCreated'
        )

        $statsProperties = @(
            'LastLogonTime'
        )

        $groupTableProperties = @(
            'DisplayName'
            'Type'
        )

        $permissionSortProperties = @(
            'User'
            'Permission'
        )

        $permissionTableProperties = @(
            'User'
            'Permission'
        )

        $deviceProperties = @(
            'FriendlyName'
            'DeviceModel'
            'DeviceOS'
            'DeviceType'
            'FirstSyncTime'
            'DeviceAccessState'
        )

        $deviceTableProperties = @(
            'FriendlyName'
            'DeviceModel'
            'DeviceOS'
            'FirstSyncTime'
        )
    }

    process {
        foreach ($inputId in $Identity) {
            $lookupValue = $inputId.Trim()

            #----------------------------------------------------------[Find user]----------------------------------------------------------

            $user        = $null
            $lookupError = $null

            try {
                $userUri = "v1.0/users/$([uri]::EscapeDataString($lookupValue))?`$select=$userSelect"
                $user = Invoke-MgGraphRequest -Method GET -Uri $userUri -OutputType PSObject -ErrorAction Stop
            }
            catch {
                $lookupError = $_.Exception.Message
                Write-Verbose "Direct lookup of '$lookupValue' failed: $lookupError"
            }

            # Not a UPN / object ID: try primary mail address, then proxy (alias) addresses.
            if (-not $user -and $lookupValue -match $emailPattern) {
                $safeValue = $lookupValue -replace "'", "''"
                $filters = @(
                    "mail eq '$safeValue'"
                    "proxyAddresses/any(p:p eq 'smtp:$safeValue')"
                )

                foreach ($filter in $filters) {
                    $searchUri = "v1.0/users?`$filter=$([uri]::EscapeDataString($filter))&`$select=$userSelect"
                    try {
                        $found = @(Invoke-GraphPagedRequest -Uri $searchUri)
                    }
                    catch {
                        $lookupError = $_.Exception.Message
                        Write-Verbose "Lookup of '$lookupValue' with filter [$filter] failed: $lookupError"
                        continue
                    }

                    if ($found.Count -gt 0) {
                        if ($found.Count -gt 1) {
                            Write-Warning "Multiple users match '$lookupValue'. Using the first one ($($found[0].userPrincipalName))."
                        }
                        $user = $found[0]
                        break
                    }
                }
            }

            if (-not $user) {
                if ($lookupError) {
                    Write-Error "Unable to find user '$lookupValue': $lookupError"
                }
                else {
                    Write-Error "User '$lookupValue' not found."
                }
                continue
            }

            $upn = [string]$user.userPrincipalName

            Write-Host "`n===================================================" -ForegroundColor Cyan
            Write-Host " User Summary for: $upn" -ForegroundColor Cyan
            Write-Host "===================================================" -ForegroundColor Cyan

            #----------------------------------------------------------[Manager]----------------------------------------------------------

            $managerText = 'None'
            try {
                $managerUri = "v1.0/users/$($user.id)/manager?`$select=displayName,userPrincipalName,mail"
                $mgr = Invoke-MgGraphRequest -Method GET -Uri $managerUri -OutputType PSObject -ErrorAction Stop
                if ($mgr) {
                    $mgrAddress = if ($mgr.userPrincipalName) { $mgr.userPrincipalName } else { $mgr.mail }
                    $managerText = "$($mgr.displayName) ($mgrAddress)"
                }
            }
            catch {
                # Graph returns 404 when no manager is set.
                Write-Verbose "No manager returned for $($upn): $($_.Exception.Message)"
            }

            #----------------------------------------------------------[Sign-in activity]----------------------------------------------------------

            $signInAvailable = $false
            $signIn          = $null
            try {
                $signInUri = "v1.0/users/$($user.id)?`$select=signInActivity"
                $signInResult = Invoke-MgGraphRequest -Method GET -Uri $signInUri -OutputType PSObject -ErrorAction Stop
                $signIn = $signInResult.signInActivity
                $signInAvailable = $true
            }
            catch {
                Write-Verbose "Sign-in activity unavailable for $($upn): $($_.Exception.Message)"
            }

            #----------------------------------------------------------[Account overview]----------------------------------------------------------

            $passwordLastSet = ConvertTo-LocalDateTime $user.lastPasswordChangeDateTime
            $passwordAgeDays = $null
            if ($passwordLastSet) {
                $passwordAgeDays = [int][math]::Floor(((Get-Date) - $passwordLastSet).TotalDays)
            }

            $accountStatus = if ($user.accountEnabled) { 'Enabled' } else { 'Disabled (sign-in blocked)' }
            $source        = if ($user.onPremisesSyncEnabled) { 'Synced from on-premises AD' } else { 'Cloud' }

            $phones = @()
            foreach ($phone in @($user.businessPhones)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$phone)) { $phones += [string]$phone }
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$user.mobilePhone)) {
                $phones += "$($user.mobilePhone) (mobile)"
            }

            $signInUnavailableText = 'Unavailable (requires Entra ID P1 and AuditLog.Read.All)'
            if ($signInAvailable) {
                $lastSignIn               = ConvertTo-LocalDateTime $signIn.lastSignInDateTime
                $lastNonInteractiveSignIn = ConvertTo-LocalDateTime $signIn.lastNonInteractiveSignInDateTime
                $lastSuccessfulSignIn     = ConvertTo-LocalDateTime $signIn.lastSuccessfulSignInDateTime
            }

            $account = [pscustomobject][ordered]@{
                DisplayName              = $user.displayName
                UserPrincipalName        = $upn
                Mail                     = $user.mail
                AccountStatus            = $accountStatus
                UserType                 = $user.userType
                Source                   = $source
                LastDirSync              = ConvertTo-LocalDateTime $user.onPremisesLastSyncDateTime
                Created                  = ConvertTo-LocalDateTime $user.createdDateTime
                JobTitle                 = $user.jobTitle
                Department               = $user.department
                Office                   = $user.officeLocation
                UsageLocation            = $user.usageLocation
                Phone                    = $phones -join '; '
                Manager                  = $managerText
                PasswordLastSet          = $passwordLastSet
                PasswordAgeDays          = $passwordAgeDays
                PasswordNeverExpires     = ([string]$user.passwordPolicies -match 'DisablePasswordExpiration')
                LastSignIn               = if ($signInAvailable) { $lastSignIn } else { $signInUnavailableText }
                LastNonInteractiveSignIn = if ($signInAvailable) { $lastNonInteractiveSignIn } else { $signInUnavailableText }
                LastSuccessfulSignIn     = if ($signInAvailable) { $lastSuccessfulSignIn } else { $signInUnavailableText }
            }

            Write-Section 'Account'
            $account | Format-List | Out-Host

            #----------------------------------------------------------[Authentication methods]----------------------------------------------------------

            Write-Section 'Authentication Methods'
            $authMethods = @()
            try {
                $methods = @(Invoke-GraphPagedRequest -Uri "v1.0/users/$($user.id)/authentication/methods")
                $methodList = [System.Collections.Generic.List[string]]::new()
                foreach ($method in $methods) {
                    $short = [string]$method.'@odata.type' -replace '^#microsoft\.graph\.', '' -replace 'AuthenticationMethod$', ''
                    $friendly = if ($authMethodNames.ContainsKey($short)) { $authMethodNames[$short] } else { $short }
                    if (-not $methodList.Contains($friendly)) { $methodList.Add($friendly) }
                }
                $authMethods = @($methodList | Sort-Object)

                if ($authMethods.Count -gt 0) {
                    foreach ($m in $authMethods) { Write-Host "  $m" }
                    if (-not ($authMethods | Where-Object { $_ -ne 'Password' -and $_ -ne 'Email' })) {
                        Write-Host '  (No MFA method registered)' -ForegroundColor Red
                    }
                }
                else {
                    Write-Host '  None registered' -ForegroundColor Gray
                }
            }
            catch {
                Write-Warning "Unable to read authentication methods: $($_.Exception.Message)"
            }

            #----------------------------------------------------------[Licenses]----------------------------------------------------------

            Write-Section 'Licenses'
            $licenses = @()
            try {
                $licenseDetails = @(Invoke-GraphPagedRequest -Uri "v1.0/users/$($user.id)/licenseDetails?`$select=skuId,skuPartNumber")
                $states         = @($user.licenseAssignmentStates)

                $licenses = @(foreach ($detail in $licenseDetails) {
                    $skuStates = @($states | Where-Object { [string]$_.skuId -eq [string]$detail.skuId })

                    $assignedBy = @($skuStates | ForEach-Object {
                        if ($_.assignedByGroup) { 'Group' } else { 'Direct' }
                    } | Sort-Object -Unique)

                    $errors = @($skuStates | Where-Object { $_.error -and $_.error -ne 'None' } |
                        ForEach-Object { [string]$_.error } | Sort-Object -Unique)

                    [pscustomobject]@{
                        License    = $detail.skuPartNumber
                        AssignedBy = if ($assignedBy.Count -gt 0) { $assignedBy -join ', ' } else { 'Unknown' }
                        Error      = if ($errors.Count -gt 0) { $errors -join ', ' } else { $null }
                    }
                })
                $licenses = @($licenses | Sort-Object -Property License)

                if ($licenses.Count -gt 0) {
                    $licenses | Format-Table -AutoSize | Out-Host
                }
                else {
                    Write-Host '  No licenses assigned' -ForegroundColor Gray
                }
            }
            catch {
                Write-Warning "Unable to read licenses: $($_.Exception.Message)"
            }

            #----------------------------------------------------------[Groups and roles]----------------------------------------------------------

            Write-Section 'Group Memberships and Roles'
            $groups = @()
            $roles  = @()
            try {
                $memberOfUri = "v1.0/users/$($user.id)/memberOf?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled,onPremisesSyncEnabled"
                $memberOf = @(Invoke-GraphPagedRequest -Uri $memberOfUri)

                $groups = @(foreach ($entry in $memberOf) {
                    if ($entry.'@odata.type' -ne '#microsoft.graph.group') { continue }

                    $kind = if ($entry.groupTypes -contains 'Unified') { 'Microsoft 365' }
                            elseif ($entry.mailEnabled -and $entry.securityEnabled) { 'Mail-enabled security' }
                            elseif ($entry.mailEnabled) { 'Distribution' }
                            else { 'Security' }
                    if ($entry.groupTypes -contains 'DynamicMembership') { $kind += ' (dynamic)' }
                    if ($entry.onPremisesSyncEnabled) { $kind += ' (synced)' }

                    [pscustomobject]@{
                        DisplayName = $entry.displayName
                        Type        = $kind
                        Id          = $entry.id
                    }
                })
                $groups = @($groups | Sort-Object -Property DisplayName)

                $roles = @($memberOf |
                    Where-Object { $_.'@odata.type' -eq '#microsoft.graph.directoryRole' } |
                    ForEach-Object { [string]$_.displayName } |
                    Sort-Object -Unique)

                if ($roles.Count -gt 0) {
                    Write-Host "  Admin roles: $($roles -join ', ')" -ForegroundColor Magenta
                }
                else {
                    Write-Host '  Admin roles: None' -ForegroundColor Gray
                }

                if ($groups.Count -gt 0) {
                    Write-Host "  Member of $($groups.Count) group(s):"
                    $groups | Format-Table -Property $groupTableProperties -AutoSize | Out-Host
                }
                else {
                    Write-Host '  Not a member of any groups' -ForegroundColor Gray
                }
            }
            catch {
                Write-Warning "Unable to read group memberships: $($_.Exception.Message)"
            }

            #----------------------------------------------------------[Exchange Online]----------------------------------------------------------

            $mailboxInfo        = $null
            $mailboxPermissions = @()
            $autoReply          = $null
            $mobileDevices      = @()

            if (-not $SkipMailbox) {
                Write-Section 'Mailbox'

                $mailbox = $null
                try {
                    $mailbox = Get-EXOMailbox -Identity $upn -Properties $mailboxProperties -ErrorAction Stop
                }
                catch {
                    Write-Host "  No Exchange Online mailbox found ($($_.Exception.Message))" -ForegroundColor Gray
                }

                if ($mailbox) {
                    $address = [string]$mailbox.PrimarySmtpAddress

                    # --- Overview ---
                    $stats = $null
                    try {
                        $stats = Get-EXOMailboxStatistics -Identity $address -Properties $statsProperties -ErrorAction Stop
                    }
                    catch {
                        Write-Warning "Unable to read mailbox statistics: $($_.Exception.Message)"
                    }

                    $forwardTargets = @()
                    if ($mailbox.ForwardingSmtpAddress) {
                        $forwardTargets += ([string]$mailbox.ForwardingSmtpAddress -replace '^smtp:', '')
                    }
                    if ($mailbox.ForwardingAddress) {
                        $forwardTargets += [string]$mailbox.ForwardingAddress
                    }
                    $forwarding = if ($forwardTargets.Count -gt 0) { $forwardTargets -join '; ' } else { 'None' }

                    $hasArchive = $mailbox.ArchiveGuid -and ([guid]$mailbox.ArchiveGuid -ne [guid]::Empty)

                    $mailboxInfo = [pscustomobject][ordered]@{
                        PrimarySmtpAddress  = $address
                        MailboxType         = $mailbox.RecipientTypeDetails
                        Created             = $mailbox.WhenMailboxCreated
                        TotalItemSize       = if ($stats) { $stats.TotalItemSize } else { $null }
                        ItemCount           = if ($stats) { $stats.ItemCount } else { $null }
                        ProhibitSendQuota   = $mailbox.ProhibitSendQuota
                        LastLogonTime       = if ($stats) { $stats.LastLogonTime } else { $null }
                        ArchiveEnabled      = [bool]$hasArchive
                        LitigationHold      = $mailbox.LitigationHoldEnabled
                        HiddenFromGAL       = $mailbox.HiddenFromAddressListsEnabled
                        Forwarding          = $forwarding
                        KeepCopyOnForward   = if ($forwardTargets.Count -gt 0) { $mailbox.DeliverToMailboxAndForward } else { $null }
                    }
                    $mailboxInfo | Format-List | Out-Host

                    if ($forwardTargets.Count -gt 0) {
                        Write-Host "  Mail is being forwarded to: $forwarding" -ForegroundColor Red
                    }

                    # --- Permissions ---
                    Write-Section 'Mailbox Access'
                    $permissionList = [System.Collections.Generic.List[object]]::new()

                    try {
                        $fullAccess = @(Get-EXOMailboxPermission -Identity $address -ErrorAction Stop |
                            Where-Object { -not $_.Deny -and -not $_.IsInherited -and ([string]$_.User -notmatch $systemPrincipalPattern) })
                        foreach ($perm in $fullAccess) {
                            $permissionList.Add([pscustomobject]@{
                                User       = [string]$perm.User
                                Permission = (@($perm.AccessRights) -join ', ')
                            })
                        }
                    }
                    catch {
                        Write-Warning "Unable to read mailbox permissions: $($_.Exception.Message)"
                    }

                    try {
                        $sendAs = @(Get-EXORecipientPermission -Identity $address -ErrorAction Stop |
                            Where-Object { $_.AccessControlType -eq 'Allow' -and ([string]$_.Trustee -notmatch $systemPrincipalPattern) })
                        foreach ($perm in $sendAs) {
                            $permissionList.Add([pscustomobject]@{
                                User       = [string]$perm.Trustee
                                Permission = (@($perm.AccessRights) -join ', ')
                            })
                        }
                    }
                    catch {
                        Write-Warning "Unable to read SendAs permissions: $($_.Exception.Message)"
                    }

                    foreach ($delegate in @($mailbox.GrantSendOnBehalfTo)) {
                        if ([string]::IsNullOrWhiteSpace([string]$delegate)) { continue }
                        $permissionList.Add([pscustomobject]@{
                            User       = [string]$delegate
                            Permission = 'SendOnBehalf'
                        })
                    }

                    $mailboxPermissions = @($permissionList | Sort-Object -Property $permissionSortProperties)

                    if ($mailboxPermissions.Count -gt 0) {
                        $mailboxPermissions | Format-Table -Property $permissionTableProperties -AutoSize | Out-Host
                    }
                    else {
                        Write-Host '  No delegates have access to this mailbox' -ForegroundColor Gray
                    }

                    # --- Out of office ---
                    Write-Section 'Out of Office'
                    try {
                        $oof = Get-MailboxAutoReplyConfiguration -Identity $address -ErrorAction Stop
                        $state = [string]$oof.AutoReplyState

                        $autoReply = [pscustomobject][ordered]@{
                            State            = $state
                            StartTime        = if ($state -eq 'Scheduled') { $oof.StartTime } else { $null }
                            EndTime          = if ($state -eq 'Scheduled') { $oof.EndTime } else { $null }
                            ExternalAudience = $oof.ExternalAudience
                            InternalMessage  = ConvertTo-PlainText ([string]$oof.InternalMessage)
                            ExternalMessage  = ConvertTo-PlainText ([string]$oof.ExternalMessage)
                        }

                        if ($state -eq 'Disabled') {
                            Write-Host '  Out of office is not set' -ForegroundColor Gray
                        }
                        else {
                            [pscustomobject][ordered]@{
                                State            = $autoReply.State
                                StartTime        = $autoReply.StartTime
                                EndTime          = $autoReply.EndTime
                                ExternalAudience = $autoReply.ExternalAudience
                                InternalMessage  = Get-ShortText $autoReply.InternalMessage
                                ExternalMessage  = Get-ShortText $autoReply.ExternalMessage
                            } | Format-List | Out-Host
                        }
                    }
                    catch {
                        Write-Warning "Unable to read out-of-office settings: $($_.Exception.Message)"
                    }

                    # --- Mobile devices ---
                    Write-Section 'Mobile Devices'
                    try {
                        $mobileDevices = @(Get-MobileDevice -Mailbox $address -ErrorAction Stop |
                            Select-Object -Property $deviceProperties)

                        if ($mobileDevices.Count -gt 0) {
                            $mobileDevices | Format-Table -Property $deviceTableProperties -AutoSize | Out-Host
                        }
                        else {
                            Write-Host '  No mobile devices' -ForegroundColor Gray
                        }
                    }
                    catch {
                        Write-Warning "Unable to read mobile devices: $($_.Exception.Message)"
                    }
                }
            }

            Write-Host "`n===================================================" -ForegroundColor Cyan
            Write-Host " Summary Complete" -ForegroundColor Cyan
            Write-Host "===================================================`n" -ForegroundColor Cyan

            if ($PassThru) {
                [pscustomobject]@{
                    UserPrincipalName   = $upn
                    Account             = $account
                    AuthMethods         = $authMethods
                    Licenses            = $licenses
                    AdminRoles          = $roles
                    Groups              = $groups
                    Mailbox             = $mailboxInfo
                    MailboxPermissions  = $mailboxPermissions
                    OutOfOffice         = $autoReply
                    MobileDevices       = $mobileDevices
                }
            }
        }
    }
}

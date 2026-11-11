function Get-M365AdminMfaReport {
    <#
    .SYNOPSIS
        Reports the MFA registration status of every Entra ID directory role holder.

    .DESCRIPTION
        Finds all users holding an Entra ID directory role (Global Administrator,
        Exchange Administrator, etc.), including users who receive a role through
        a role-assignable group, then checks their registered authentication
        methods and classifies their MFA strength:

            Phishing-resistant  FIDO2, Windows Hello for Business, platform credential, certificate
            Enabled             Microsoft Authenticator, software/hardware OATH token
            Weak (phone only)   SMS / voice only
            Disabled            Only password / email / Temporary Access Pass
            Unknown             Authentication methods could not be read

        Report objects are written to the pipeline and exported to CSV in
        -OutputPath unless -NoExport is specified.

    .PARAMETER IncludeEligible
        Also include PIM eligible role assignments (requires Entra ID P2).

    .PARAMETER OutputPath
        Folder for the CSV export. Created if it does not exist.
        Default: $env:TEMP\PSTools

    .PARAMETER NoExport
        Do not write a CSV file.

    .EXAMPLE
        Get-M365AdminMfaReport | Format-Table -AutoSize

    .EXAMPLE
        Get-M365AdminMfaReport -IncludeEligible -OutputPath C:\Temp

    .EXAMPLE
        Get-M365AdminMfaReport -NoExport | Where-Object MfaStatus -notin 'Enabled', 'Phishing-resistant'

    .NOTES
        Requires the Microsoft.Graph.Authentication module and the delegated scopes:
            User.Read.All, RoleManagement.Read.Directory, GroupMember.Read.All,
            UserAuthenticationMethod.Read.All
        An existing Graph session with these scopes is reused.
    #>
    [CmdletBinding()]
    param(
        [switch]$IncludeEligible,

        [ValidateNotNullOrEmpty()]
        [string]$OutputPath = (Join-Path $env:TEMP 'PSTools'),

        [switch]$NoExport
    )

    $graphScopes = @(
        'User.Read.All'
        'RoleManagement.Read.Directory'
        'GroupMember.Read.All'
        'UserAuthenticationMethod.Read.All'
    )
    Assert-MgGraphConnection -Scopes $graphScopes

    #----------------------------------------------------------[Inner helpers]----------------------------------------------------------

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

    # Records a role label against a user id.
    function Add-AdminRole {
        param([string]$UserId, [string]$RoleLabel)
        if (-not $admins.ContainsKey($UserId)) {
            $admins[$UserId] = [System.Collections.Generic.List[string]]::new()
        }
        if (-not $admins[$UserId].Contains($RoleLabel)) {
            $admins[$UserId].Add($RoleLabel)
        }
    }

    # Returns the user ids of all (transitive) members of a group, cached.
    function Get-GroupUserIds {
        param([string]$GroupId)
        if (-not $groupMemberCache.ContainsKey($GroupId)) {
            $uri = "v1.0/groups/$($GroupId)/transitiveMembers?`$select=id"
            $memberIds = @(Invoke-GraphPagedRequest -Uri $uri |
                Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' } |
                ForEach-Object { $_.id })
            $groupMemberCache[$GroupId] = $memberIds
        }
        return $groupMemberCache[$GroupId]
    }

    # Resolves an assignment's principal to user(s) and records the role.
    function Register-RoleAssignment {
        param($Assignment, [switch]$Eligible)

        $roleName = $roleNames[[string]$Assignment.roleDefinitionId]
        if (-not $roleName) { $roleName = [string]$Assignment.roleDefinitionId }

        $label = $roleName
        if ($Assignment.directoryScopeId -and $Assignment.directoryScopeId -ne '/') { $label += ' [scoped]' }
        if ($Eligible) { $label += ' (Eligible)' }

        $principal = $Assignment.principal
        if (-not $principal) {
            try {
                $principal = Invoke-MgGraphRequest -Method GET -Uri "v1.0/directoryObjects/$($Assignment.principalId)" -OutputType PSObject -ErrorAction Stop
            }
            catch {
                Write-Verbose "Could not resolve principal $($Assignment.principalId): $($_.Exception.Message)"
                return
            }
        }

        switch ([string]$principal.'@odata.type') {
            '#microsoft.graph.user' {
                Add-AdminRole -UserId $Assignment.principalId -RoleLabel $label
            }
            '#microsoft.graph.group' {
                $groupLabel = "$label via $($principal.displayName)"
                try {
                    foreach ($memberId in (Get-GroupUserIds -GroupId $Assignment.principalId)) {
                        Add-AdminRole -UserId $memberId -RoleLabel $groupLabel
                    }
                }
                catch {
                    Write-Warning "Could not expand members of role group '$($principal.displayName)': $($_.Exception.Message)"
                }
            }
            default {
                Write-Verbose "Skipping non-user principal $($Assignment.principalId) ($($principal.'@odata.type')) for role '$roleName'."
            }
        }
    }

    #----------------------------------------------------------[Method catalog]----------------------------------------------------------

    # Strength: 3 = phishing-resistant, 2 = app/token MFA, 1 = phone (weak), 0 = not MFA
    $methodCatalog = @{
        '#microsoft.graph.fido2AuthenticationMethod'                          = @{ Name = 'FIDO2 Security Key';                         Strength = 3 }
        '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod'        = @{ Name = 'Windows Hello for Business';                 Strength = 3 }
        '#microsoft.graph.platformCredentialAuthenticationMethod'             = @{ Name = 'Platform Credential';                        Strength = 3 }
        '#microsoft.graph.x509CertificateAuthenticationMethod'                = @{ Name = 'Certificate (CBA)';                          Strength = 3 }
        '#microsoft.graph.passwordlessMicrosoftAuthenticatorAuthenticationMethod' = @{ Name = 'Microsoft Authenticator App (Passwordless)'; Strength = 2 }
        '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod'         = @{ Name = 'Microsoft Authenticator App';                Strength = 2 }
        '#microsoft.graph.softwareOathAuthenticationMethod'                   = @{ Name = 'Software OATH Token';                        Strength = 2 }
        '#microsoft.graph.hardwareOathAuthenticationMethod'                   = @{ Name = 'Hardware OATH Token';                        Strength = 2 }
        '#microsoft.graph.phoneAuthenticationMethod'                          = @{ Name = 'Phone (SMS/Voice)';                          Strength = 1 }
        '#microsoft.graph.emailAuthenticationMethod'                          = @{ Name = 'Email (SSPR only)';                          Strength = 0 }
        '#microsoft.graph.temporaryAccessPassAuthenticationMethod'            = @{ Name = 'Temporary Access Pass';                      Strength = 0 }
        '#microsoft.graph.passwordAuthenticationMethod'                       = @{ Name = 'Password';                                   Strength = 0 }
    }

    $strengthLabels = @{
        3 = 'Phishing-resistant'
        2 = 'Enabled'
        1 = 'Weak (phone only)'
        0 = 'Disabled'
    }

    #----------------------------------------------------------[Collect role holders]----------------------------------------------------------

    $admins           = @{}
    $groupMemberCache = @{}
    $roleNames        = @{}

    Write-Host 'Reading directory role definitions...'
    foreach ($definition in (Invoke-GraphPagedRequest -Uri 'v1.0/roleManagement/directory/roleDefinitions?$select=id,displayName')) {
        $roleNames[[string]$definition.id] = $definition.displayName
    }

    Write-Host 'Reading active role assignments...'
    $activeAssignments = @(Invoke-GraphPagedRequest -Uri 'v1.0/roleManagement/directory/roleAssignments?$expand=principal')
    foreach ($assignment in $activeAssignments) {
        Register-RoleAssignment -Assignment $assignment
    }

    if ($IncludeEligible) {
        Write-Host 'Reading PIM eligible role assignments...'
        try {
            $eligibleAssignments = @(Invoke-GraphPagedRequest -Uri 'v1.0/roleManagement/directory/roleEligibilitySchedules?$expand=principal')
            foreach ($assignment in $eligibleAssignments) {
                Register-RoleAssignment -Assignment $assignment -Eligible
            }
        }
        catch {
            Write-Warning "Could not read PIM eligible assignments (requires Entra ID P2): $($_.Exception.Message)"
        }
    }

    if ($admins.Count -eq 0) {
        Write-Warning 'No users with directory roles were found.'
        return
    }

    #----------------------------------------------------------[Check MFA]----------------------------------------------------------

    Write-Host "Checking MFA status for $($admins.Count) administrative user(s)..."

    $report = [System.Collections.Generic.List[object]]::new()

    foreach ($userId in @($admins.Keys)) {
        $displayName    = $userId
        $upn            = $null
        $accountEnabled = $null

        try {
            $user = Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$($userId)?`$select=id,displayName,userPrincipalName,accountEnabled" -OutputType PSObject -ErrorAction Stop
            $displayName    = $user.displayName
            $upn            = $user.userPrincipalName
            $accountEnabled = $user.accountEnabled
        }
        catch {
            Write-Warning "Could not read user $($userId): $($_.Exception.Message)"
        }

        Write-Verbose "Checking authentication methods for $(if ($upn) { $upn } else { $userId })"

        $mfaStatus       = 'Unknown'
        $strongestMethod = $null
        $methodNames     = @()

        try {
            $methods = @(Invoke-GraphPagedRequest -Uri "v1.0/users/$($userId)/authentication/methods")

            $bestStrength = 0
            $names = [System.Collections.Generic.List[string]]::new()

            foreach ($method in $methods) {
                $type  = [string]$method.'@odata.type'
                $entry = $methodCatalog[$type]
                if (-not $entry) {
                    $friendly = $type -replace '^#microsoft\.graph\.', '' -replace 'AuthenticationMethod$', ''
                    $entry = @{ Name = $friendly; Strength = 0 }
                }

                if (-not $names.Contains($entry.Name)) { $names.Add($entry.Name) }

                if ($entry.Strength -gt $bestStrength) {
                    $bestStrength    = $entry.Strength
                    $strongestMethod = $entry.Name
                }
            }

            $mfaStatus   = $strengthLabels[$bestStrength]
            $methodNames = @($names)
            if (-not $strongestMethod) { $strongestMethod = 'None' }
        }
        catch {
            Write-Warning "Could not read authentication methods for $(if ($upn) { $upn } else { $userId }): $($_.Exception.Message)"
        }

        $report.Add([pscustomobject]@{
            DisplayName       = $displayName
            UserPrincipalName = $upn
            AccountEnabled    = $accountEnabled
            Roles             = (@($admins[$userId]) | Sort-Object) -join '; '
            MfaStatus         = $mfaStatus
            StrongestMethod   = $strongestMethod
            RegisteredMethods = $methodNames -join '; '
        })
    }

    $sorted = @($report | Sort-Object -Property DisplayName)

    #----------------------------------------------------------[Summary / export]----------------------------------------------------------

    $atRisk = @($sorted | Where-Object { $_.MfaStatus -eq 'Disabled' -or $_.MfaStatus -eq 'Weak (phone only)' })
    if ($atRisk.Count -gt 0) {
        Write-Warning "$($atRisk.Count) admin(s) have no app/token-based MFA registered."
    }

    if (-not $NoExport) {
        if (-not (Test-Path -LiteralPath $OutputPath)) {
            $null = New-Item -ItemType Directory -Path $OutputPath -Force
        }
        $exportPath = Join-Path $OutputPath "AdminMfaStatus_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        $sorted | Export-Csv -Path $exportPath -NoTypeInformation
        Write-Host "Admin MFA report exported to: $exportPath" -ForegroundColor Green
    }

    $sorted
}

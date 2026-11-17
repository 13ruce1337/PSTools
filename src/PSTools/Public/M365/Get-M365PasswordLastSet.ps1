function Get-M365PasswordLastSet {
    <#
    .SYNOPSIS
        Shows the Microsoft 365 users whose password was set longest ago.

    .DESCRIPTION
        Reads the password last set date of every user in the tenant (or the
        users given with -Identity) from Microsoft Graph. It then prints the
        -Top users (default 30) with the oldest password, longest first.
        Users with no password last set date are listed last.

        Nothing is written to disk unless -Export or -OutputPath is specified.
        When exporting, the full sorted list (not only the top entries) is
        written to a CSV file in -OutputPath.

        Use -PassThru to also return every record to the pipeline.

    .PARAMETER Identity
        Optional. Specific users to report on: UPN, object ID or email address
        (primary or alias). Accepts pipeline input. If omitted, all users in the
        tenant are read.

    .PARAMETER Top
        Number of users to show on screen. Default 30.

    .PARAMETER EnabledOnly
        Only include enabled accounts.

    .PARAMETER Export
        Export the full results to a CSV file in -OutputPath.

    .PARAMETER OutputPath
        Folder for the CSV export. Created if it does not exist. Specifying this
        parameter also turns on the export.
        Default: $env:TEMP\PSTools

    .PARAMETER PassThru
        Return all records to the pipeline, in addition to the on-screen report.

    .EXAMPLE
        Get-M365PasswordLastSet

    .EXAMPLE
        Get-M365PasswordLastSet -Top 50 -EnabledOnly

    .EXAMPLE
        Get-M365PasswordLastSet -OutputPath C:\Reports

    .EXAMPLE
        Get-M365PasswordLastSet -Identity user@contoso.com

    .EXAMPLE
        Get-M365PasswordLastSet -PassThru | Where-Object { $_.IsActive -and $_.PasswordAgeDays -gt 365 }

    .NOTES
        Requires the Microsoft.Graph.Authentication module and the delegated scope:
            User.Read.All
        An existing Graph session with this scope is reused.

        For users synced from on-premises AD, the password last set date reflects
        the last password change synced by Entra Connect (password hash sync).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('UserPrincipalName', 'UPN', 'PrimarySmtpAddress', 'Mail', 'Id')]
        [ValidateNotNullOrEmpty()]
        [string[]]$Identity,

        [ValidateRange(1, 2147483647)]
        [int]$Top = 30,

        [switch]$EnabledOnly,

        [switch]$Export,

        [ValidateNotNullOrEmpty()]
        [string]$OutputPath = (Join-Path $env:TEMP 'PSTools'),

        [switch]$PassThru
    )

    begin {
        $graphScopes = @(
            'User.Read.All'
        )
        Assert-MgGraphConnection -Scopes $graphScopes

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

        # Builds a report record from a Graph user object.
        function ConvertTo-PasswordRecord {
            param($User)

            $lastSet = ConvertTo-LocalDateTime $User.lastPasswordChangeDateTime
            $ageDays = $null
            if ($lastSet) {
                $ageDays = [int][math]::Floor(((Get-Date) - $lastSet).TotalDays)
            }

            [pscustomobject][ordered]@{
                DisplayName          = $User.displayName
                UserPrincipalName    = $User.userPrincipalName
                Email                = $User.mail
                FirstName            = $User.givenName
                LastName             = $User.surname
                PasswordLastSet      = $lastSet
                PasswordAgeDays      = $ageDays
                IsActive             = [bool]$User.accountEnabled
                PasswordNeverExpires = ([string]$User.passwordPolicies -match 'DisablePasswordExpiration')
            }
        }

        #----------------------------------------------------------[Static data]----------------------------------------------------------

        $emailPattern = '^[^@\s]+@[^@\s]+\.[^@\s]+$'

        $userProperties = @(
            'id'
            'displayName'
            'userPrincipalName'
            'mail'
            'givenName'
            'surname'
            'lastPasswordChangeDateTime'
            'accountEnabled'
            'passwordPolicies'
        )
        $userSelect = $userProperties -join ','

        # Oldest password first; users without a date go last.
        $sortProperties = @(
            @{ Expression = { $null -eq $_.PasswordLastSet }; Ascending = $true }
            @{ Expression = { $_.PasswordLastSet };           Ascending = $true }
            @{ Expression = { $_.DisplayName };               Ascending = $true }
        )

        $tableProperties = @(
            'DisplayName'
            'UserPrincipalName'
            @{ Label = 'PasswordLastSet'; Expression = { if ($_.PasswordLastSet) { $_.PasswordLastSet } else { 'Never / unknown' } } }
            @{ Label = 'AgeDays';         Expression = { $_.PasswordAgeDays }; Alignment = 'Right' }
            'IsActive'
            'PasswordNeverExpires'
        )

        $doExport = $Export -or $PSBoundParameters.ContainsKey('OutputPath')

        if ($doExport -and -not (Test-Path -LiteralPath $OutputPath)) {
            $null = New-Item -ItemType Directory -Path $OutputPath -Force
        }

        $records      = [System.Collections.Generic.List[object]]::new()
        $seenIds      = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $identityMode = $false
    }

    process {
        if (-not $Identity) { return }
        $identityMode = $true

        foreach ($inputId in $Identity) {
            $lookupValue = $inputId.Trim()

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

            if ($seenIds.Add([string]$user.id)) {
                $records.Add((ConvertTo-PasswordRecord $user))
            }
        }
    }

    end {
        #----------------------------------------------------------[All users]----------------------------------------------------------

        if (-not $identityMode) {
            Write-Host 'Reading users from Microsoft Graph...' -ForegroundColor DarkGray
            $allUsersUri = "v1.0/users?`$select=$userSelect&`$top=999"

            try {
                foreach ($user in (Invoke-GraphPagedRequest -Uri $allUsersUri)) {
                    if ($seenIds.Add([string]$user.id)) {
                        $records.Add((ConvertTo-PasswordRecord $user))
                    }
                }
            }
            catch {
                Write-Error "Unable to read users: $($_.Exception.Message)"
                return
            }
        }

        $filtered = @($records)
        if ($EnabledOnly) {
            $filtered = @($filtered | Where-Object { $_.IsActive })
        }

        if ($filtered.Count -eq 0) {
            Write-Warning 'No users found.'
            return
        }

        $sorted = @($filtered | Sort-Object -Property $sortProperties)

        #----------------------------------------------------------[Display]----------------------------------------------------------

        $shown = @($sorted | Select-Object -First $Top)
        $scope = if ($EnabledOnly) { 'enabled user(s)' } else { 'user(s)' }

        Write-Host ''
        Write-Host "=== Top $($shown.Count) of $($sorted.Count) $scope by oldest password ===" -ForegroundColor Green
        $shown | Format-Table -Property $tableProperties -AutoSize | Out-Host

        $neverSet = @($sorted | Where-Object { $null -eq $_.PasswordLastSet }).Count
        if ($neverSet -gt 0) {
            Write-Host "$neverSet user(s) have no password last set date." -ForegroundColor Yellow
        }

        #----------------------------------------------------------[Export]----------------------------------------------------------

        if ($doExport) {
            $exportPath = Join-Path $OutputPath "PasswordLastSet_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
            $sorted | Export-Csv -Path $exportPath -NoTypeInformation
            Write-Host "Exported $($sorted.Count) user(s) to: $exportPath" -ForegroundColor Green
        }

        if ($PassThru) {
            $sorted
        }
    }
}

function Get-M365LicenseActivity {
    <#
    .SYNOPSIS
        Reports the last sign-in of licensed Microsoft 365 users, to find licenses
        that are assigned but not used.

    .DESCRIPTION
        Reads every user with at least one license assigned from Microsoft Graph,
        resolves license SKU IDs to their part numbers, and works out each user's
        most recent sign-in (interactive or non-interactive). Each user gets a
        status based on -DaysInactive:

            Active          Signed in within the last -DaysInactive days
            Inactive        Last sign-in is older than -DaysInactive days
            NeverSignedIn   No recorded sign-in
            New             No sign-in yet, but created within the last
                            -DaysInactive days (grace period)

        Disabled accounts are skipped unless -IncludeDisabled is specified.

        The console shows a summary by status and a table of the users that look
        unused (Inactive and NeverSignedIn), longest without a sign-in first.

        Nothing is written to disk unless -Export or -OutputPath is specified.
        When exporting, every licensed user (not only the unused ones) is written
        to a CSV file in -OutputPath.

        Use -PassThru to also return every record to the pipeline.

    .PARAMETER DaysInactive
        Number of days without a sign-in before a licensed user is considered
        inactive. Default 30.

    .PARAMETER IncludeDisabled
        Also include disabled (sign-in blocked) accounts that still hold licenses.

    .PARAMETER Export
        Export the full results to a CSV file in -OutputPath.

    .PARAMETER OutputPath
        Folder for the CSV export. Created if it does not exist. Specifying this
        parameter also turns on the export.
        Default: $env:TEMP\PSTools

    .PARAMETER PassThru
        Return all records to the pipeline, in addition to the on-screen report.

    .EXAMPLE
        Get-M365LicenseActivity

    .EXAMPLE
        Get-M365LicenseActivity -DaysInactive 90 -IncludeDisabled

    .EXAMPLE
        Get-M365LicenseActivity -DaysInactive 90 -OutputPath C:\Reports

    .EXAMPLE
        Get-M365LicenseActivity -PassThru | Where-Object Licenses -match 'SPE_E3'

    .NOTES
        Requires the Microsoft.Graph.Authentication module and the delegated scopes:
            User.Read.All, AuditLog.Read.All, Organization.Read.All
        An existing Graph session with these scopes is reused.

        Sign-in activity (signInActivity) requires Entra ID P1 or P2. It is not
        updated in real time and can lag by several hours.
    #>
    [CmdletBinding()]
    param(
        [ValidateRange(1, 3650)]
        [int]$DaysInactive = 30,

        [switch]$IncludeDisabled,

        [switch]$Export,

        [ValidateNotNullOrEmpty()]
        [string]$OutputPath = (Join-Path $env:TEMP 'PSTools'),

        [switch]$PassThru
    )

    $graphScopes = @(
        'User.Read.All'
        'AuditLog.Read.All'
        'Organization.Read.All'
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

    #----------------------------------------------------------[Static data]----------------------------------------------------------

    $userProperties = @(
        'id'
        'displayName'
        'userPrincipalName'
        'userType'
        'accountEnabled'
        'assignedLicenses'
        'createdDateTime'
        'signInActivity'
    )
    $userSelect = $userProperties -join ','

    $statusOrder = @(
        'Active'
        'Inactive'
        'NeverSignedIn'
        'New'
    )

    $unusedStatuses = @(
        'Inactive'
        'NeverSignedIn'
    )

    # Longest without a sign-in first; never signed in counts as longest.
    $sortProperties = @(
        @{ Expression = { if ($null -eq $_.DaysSinceLastSignIn) { [int]::MaxValue } else { $_.DaysSinceLastSignIn } }; Descending = $true }
        @{ Expression = { $_.DisplayName }; Ascending = $true }
    )

    $tableProperties = [System.Collections.Generic.List[object]]::new()
    $tableProperties.Add('DisplayName')
    $tableProperties.Add('UserPrincipalName')
    $tableProperties.Add('Licenses')
    $tableProperties.Add(@{ Label = 'LastSignIn'; Expression = { if ($_.LastSignIn) { $_.LastSignIn } else { 'Never' } } })
    $tableProperties.Add(@{ Label = 'Days'; Expression = { $_.DaysSinceLastSignIn }; Alignment = 'Right' })
    $tableProperties.Add('Status')
    if ($IncludeDisabled) {
        $tableProperties.Add('AccountEnabled')
    }

    $doExport = $Export -or $PSBoundParameters.ContainsKey('OutputPath')

    if ($doExport -and -not (Test-Path -LiteralPath $OutputPath)) {
        $null = New-Item -ItemType Directory -Path $OutputPath -Force
    }

    #----------------------------------------------------------[License names]----------------------------------------------------------

    Write-Host 'Reading subscribed licenses...' -ForegroundColor DarkGray
    $skuLookup = @{}
    try {
        foreach ($sku in (Invoke-GraphPagedRequest -Uri 'v1.0/subscribedSkus')) {
            $skuLookup[[string]$sku.skuId] = [string]$sku.skuPartNumber
        }
    }
    catch {
        Write-Warning "Unable to read subscribed licenses; SKU IDs will be shown instead of names: $($_.Exception.Message)"
    }

    #----------------------------------------------------------[Users]----------------------------------------------------------

    Write-Host 'Reading users and sign-in activity from Microsoft Graph (this may take a while)...' -ForegroundColor DarkGray
    $usersUri = "v1.0/users?`$select=$userSelect&`$top=999"

    try {
        $allUsers = @(Invoke-GraphPagedRequest -Uri $usersUri)
    }
    catch {
        $message = [string]$_.Exception.Message
        if ($message -match 'NonPremium') {
            Write-Error "Unable to read sign-in activity. It requires Entra ID P1 or P2 in this tenant. $message"
        }
        else {
            Write-Error "Unable to read users: $message"
        }
        return
    }

    $users = @($allUsers | Where-Object { $_.assignedLicenses -and @($_.assignedLicenses).Count -gt 0 })

    if (-not $IncludeDisabled) {
        $users = @($users | Where-Object { $_.accountEnabled -eq $true })
    }

    if ($users.Count -eq 0) {
        Write-Warning 'No licensed users found.'
        return
    }

    $scope = if ($IncludeDisabled) { 'licensed user(s), including disabled accounts' } else { 'enabled licensed user(s)' }
    Write-Host "Found $($users.Count) $scope." -ForegroundColor DarkGray

    #----------------------------------------------------------[Build records]----------------------------------------------------------

    $now    = Get-Date
    $cutoff = $now.AddDays(-$DaysInactive)

    $records = foreach ($user in $users) {
        $licenseNames = @(foreach ($license in @($user.assignedLicenses)) {
            $skuId = [string]$license.skuId
            if ($skuLookup.ContainsKey($skuId)) { $skuLookup[$skuId] } else { $skuId }
        })

        $interactive    = ConvertTo-LocalDateTime $user.signInActivity.lastSignInDateTime
        $nonInteractive = ConvertTo-LocalDateTime $user.signInActivity.lastNonInteractiveSignInDateTime
        $created        = ConvertTo-LocalDateTime $user.createdDateTime

        $lastSignIn = @($interactive, $nonInteractive) |
            Where-Object { $null -ne $_ } |
            Sort-Object -Descending |
            Select-Object -First 1

        $daysSince = $null
        if ($null -eq $lastSignIn) {
            $status = if ($created -and $created -gt $cutoff) { 'New' } else { 'NeverSignedIn' }
        }
        else {
            $daysSince = [int][math]::Floor(($now - $lastSignIn).TotalDays)
            $status    = if ($lastSignIn -lt $cutoff) { 'Inactive' } else { 'Active' }
        }

        [pscustomobject][ordered]@{
            DisplayName              = $user.displayName
            UserPrincipalName        = $user.userPrincipalName
            UserType                 = $user.userType
            AccountEnabled           = [bool]$user.accountEnabled
            Licenses                 = (@($licenseNames | Sort-Object) -join '; ')
            LicenseCount             = $licenseNames.Count
            LastSignIn               = $lastSignIn
            LastInteractiveSignIn    = $interactive
            LastNonInteractiveSignIn = $nonInteractive
            DaysSinceLastSignIn      = $daysSince
            Created                  = $created
            Status                   = $status
        }
    }

    $results = @($records | Sort-Object -Property $sortProperties)
    $unused  = @($results | Where-Object { $unusedStatuses -contains $_.Status })

    #----------------------------------------------------------[Display]----------------------------------------------------------

    Write-Host ''
    Write-Host "=== License activity summary (inactive after $DaysInactive days) ===" -ForegroundColor Green
    foreach ($statusName in $statusOrder) {
        $count = @($results | Where-Object { $_.Status -eq $statusName }).Count
        Write-Host ('  {0,-15} {1}' -f $statusName, $count)
    }

    Write-Host ''
    if ($unused.Count -gt 0) {
        $unusedLicenses = ($unused | Measure-Object -Property LicenseCount -Sum).Sum
        Write-Host "=== $($unused.Count) licensed user(s) not actively used, holding $unusedLicenses license(s) ===" -ForegroundColor Yellow
        $unused | Format-Table -Property $tableProperties.ToArray() -AutoSize | Out-Host
    }
    else {
        Write-Host 'No inactive licensed users found.' -ForegroundColor Green
    }

    #----------------------------------------------------------[Export]----------------------------------------------------------

    if ($doExport) {
        $exportPath = Join-Path $OutputPath "LicenseActivity_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        $results | Export-Csv -Path $exportPath -NoTypeInformation
        Write-Host "Exported $($results.Count) user(s) to: $exportPath" -ForegroundColor Green
    }

    if ($PassThru) {
        $results
    }
}

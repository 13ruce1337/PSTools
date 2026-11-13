function Get-M365GroupMember {
    <#
    .SYNOPSIS
        Lists the user members of an Entra ID / Microsoft 365 group, optionally exporting them to CSV.

    .DESCRIPTION
        Looks up a group by display name or email address and returns the first
        name, last name and email address of each user member. Devices, contacts
        and nested groups are skipped. Use -Transitive to also include users in
        nested groups.

        If the value looks like an email address it is matched against the
        group's primary mail address first, then its proxy (alias) addresses,
        then its display name. Otherwise it is matched against the display name.
        Additional lookups only happen when the earlier ones find nothing, so the
        common case is a single request.

        Results are written to the pipeline. Use -Export to also write them to a
        CSV file in -OutputPath.

    .PARAMETER GroupName
        Display name or email address of the group. Accepts pipeline input.

    .PARAMETER Transitive
        Include users who are members through nested groups.

    .PARAMETER Export
        Also export the results to a CSV file in -OutputPath.

    .PARAMETER OutputPath
        Folder for the CSV export (used with -Export). Created if it does not exist.
        Default: $env:TEMP\PSTools

    .EXAMPLE
        Get-M365GroupMember -GroupName 'Finance Team'

    .EXAMPLE
        Get-M365GroupMember -GroupName finance@contoso.com

    .EXAMPLE
        Get-M365GroupMember -GroupName 'All Staff' -Transitive -Export -OutputPath C:\Temp

    .EXAMPLE
        'Sales', 'marketing@contoso.com' | Get-M365GroupMember | Format-Table

    .NOTES
        Requires the Microsoft.Graph.Authentication module and the delegated scopes:
            User.Read.All, GroupMember.Read.All
        An existing Graph session with these scopes is reused.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('DisplayName', 'Group', 'Identity')]
        [ValidateNotNullOrEmpty()]
        [string[]]$GroupName,

        [switch]$Transitive,

        [switch]$Export,

        [ValidateNotNullOrEmpty()]
        [string]$OutputPath = (Join-Path $env:TEMP 'PSTools')
    )

    begin {
        $graphScopes = @(
            'User.Read.All'
            'GroupMember.Read.All'
        )
        Assert-MgGraphConnection -Scopes $graphScopes

        # Defined inline so it is not exported as part of the module.
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

        if ($Export -and -not (Test-Path -LiteralPath $OutputPath)) {
            $null = New-Item -ItemType Directory -Path $OutputPath -Force
        }

        $memberSegment  = if ($Transitive) { 'transitiveMembers' } else { 'members' }
        $invalidChars   = [System.IO.Path]::GetInvalidFileNameChars()
        $sortProperties = @('LastName', 'FirstName')
        $emailPattern   = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
    }

    process {
        foreach ($name in $GroupName) {

            #----------------------------------------------------------[Find group]----------------------------------------------------------

            # Escape single quotes for the OData filter values.
            $safeName = $name.Trim() -replace "'", "''"

            # Ordered list of lookups; later ones only run if earlier ones find nothing.
            $lookups = [System.Collections.Generic.List[object]]::new()
            if ($name.Trim() -match $emailPattern) {
                $lookups.Add(@{ Label = 'email address'; Filter = "mail eq '$safeName'" })
                $lookups.Add(@{ Label = 'proxy address'; Filter = "proxyAddresses/any(p:p eq 'smtp:$safeName')" })
            }
            $lookups.Add(@{ Label = 'display name'; Filter = "displayName eq '$safeName'" })

            $groups      = @()
            $matchedBy   = $null
            $lookupError = $null

            foreach ($lookup in $lookups) {
                $filter    = [uri]::EscapeDataString($lookup.Filter)
                $groupsUri = "v1.0/groups?`$filter=$filter&`$select=id,displayName,mail"

                try {
                    $groups = @(Invoke-GraphPagedRequest -Uri $groupsUri)
                }
                catch {
                    $lookupError = $_.Exception.Message
                    Write-Verbose "Lookup of '$name' by $($lookup.Label) failed: $lookupError"
                    continue
                }

                if ($groups.Count -gt 0) {
                    $matchedBy = $lookup.Label
                    break
                }
            }

            if ($groups.Count -eq 0) {
                if ($lookupError) {
                    Write-Error "Unable to find group '$name': $lookupError"
                }
                else {
                    Write-Error "Group '$name' not found."
                }
                continue
            }

            if ($groups.Count -gt 1) {
                Write-Warning "Multiple groups with $matchedBy '$name' found. Using the first one ($($groups[0].displayName), $($groups[0].id))."
            }
            $group = $groups[0]
            Write-Verbose "Matched '$name' by $matchedBy to group '$($group.displayName)' ($($group.id))."

            #----------------------------------------------------------[Members]----------------------------------------------------------

            # Keep only users (skip devices, contacts, nested groups).
            $membersUri = "v1.0/groups/$($group.id)/$($memberSegment)?`$select=id,givenName,surname,mail"

            try {
                $members = @(Invoke-GraphPagedRequest -Uri $membersUri |
                    Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' })
            }
            catch {
                Write-Error "Unable to list members of '$($group.displayName)': $($_.Exception.Message)"
                continue
            }

            # Transitive queries can return the same user more than once; de-duplicate by id.
            $seenIds  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
            $unsorted = [System.Collections.Generic.List[object]]::new()

            foreach ($member in $members) {
                if (-not $seenIds.Add([string]$member.id)) { continue }

                $unsorted.Add([pscustomobject]@{
                    FirstName = $member.givenName
                    LastName  = $member.surname
                    Email     = $member.mail
                })
            }

            $results = @($unsorted | Sort-Object -Property $sortProperties)

            #----------------------------------------------------------[Export]----------------------------------------------------------

            if ($Export) {
                $fileName   = -join ($group.displayName.ToCharArray() | ForEach-Object { if ($invalidChars -contains $_) { '_' } else { $_ } })
                $exportPath = Join-Path $OutputPath "GroupMembers_$($fileName)_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
                $results | Export-Csv -Path $exportPath -NoTypeInformation
                Write-Host "Exported $($results.Count) member(s) of '$($group.displayName)' to: $exportPath" -ForegroundColor Green
            }

            $results
        }
    }
}

function Get-M365GroupMember {
    <#
    .SYNOPSIS
        Lists the user members of an Entra ID / Microsoft 365 group, optionally exporting them to CSV.

    .DESCRIPTION
        Looks up a group by display name and returns the first name, last name
        and email address of each user member. Devices, contacts and nested
        groups are skipped. Use -Transitive to also include users in nested groups.

        Results are written to the pipeline. Use -Export to also write them to a
        CSV file in -OutputPath.

    .PARAMETER GroupName
        Display name of the group. Accepts pipeline input.

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
        Get-M365GroupMember -GroupName 'All Staff' -Transitive -Export -OutputPath C:\Temp

    .EXAMPLE
        'Sales', 'Marketing' | Get-M365GroupMember | Format-Table

    .NOTES
        Requires the Microsoft.Graph.Authentication module and the delegated scopes:
            User.Read.All, GroupMember.Read.All
        An existing Graph session with these scopes is reused.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('DisplayName', 'Group')]
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

        $memberSegment = if ($Transitive) { 'transitiveMembers' } else { 'members' }
        $invalidChars  = [System.IO.Path]::GetInvalidFileNameChars()
    }

    process {
        foreach ($name in $GroupName) {

            #----------------------------------------------------------[Find group]----------------------------------------------------------

            # Escape single quotes for the OData filter, then URL-encode it.
            $safeName  = $name -replace "'", "''"
            $filter    = [uri]::EscapeDataString("displayName eq '$safeName'")
            $groupsUri = "v1.0/groups?`$filter=$filter&`$select=id,displayName"

            try {
                $groups = @(Invoke-GraphPagedRequest -Uri $groupsUri)
            }
            catch {
                Write-Error "Unable to search for group '$name': $($_.Exception.Message)"
                continue
            }

            if ($groups.Count -eq 0) {
                Write-Error "Group '$name' not found."
                continue
            }

            if ($groups.Count -gt 1) {
                Write-Warning "Multiple groups named '$name' found. Using the first one ($($groups[0].id))."
            }
            $group = $groups[0]

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

            # Transitive queries can return the same user more than once.
            $results = @($members |
                Sort-Object -Property id -Unique |
                ForEach-Object {
                    [pscustomobject]@{
                        FirstName = $_.givenName
                        LastName  = $_.surname
                        Email     = $_.mail
                    }
                } |
                Sort-Object -Property LastName, FirstName)

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

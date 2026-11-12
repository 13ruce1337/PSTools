function Disable-ADUser {
    <#
    .SYNOPSIS
        Deprovisions a user in on-premises Active Directory.

    .DESCRIPTION
        For each user:
          1. Disables the account
          2. Stamps the description with the disable date and operator
             (the previous description is kept after it)
          3. Removes the user from all AD groups (except the primary group and
             any -ExcludeGroup entries)
          4. Optionally resets the password to a random value (-ResetPassword)
          5. Optionally hides the user from Exchange address lists
             (-HideFromAddressList; syncs to Exchange Online in hybrid setups)
          6. Optionally moves the user to another OU (-TargetOU)

        Failures removing individual groups are reported and do not stop the
        remaining steps. Use -PassThru to get a result object listing the groups
        that were removed (useful for auditing or rollback).

        Microsoft 365 deprovisioning is handled separately by Disable-M365User.

        Requires RSAT (the ActiveDirectory module).

    .PARAMETER Identity
        User to deprovision: SamAccountName, DistinguishedName, GUID or SID.
        Accepts pipeline input (strings, AD user objects, or objects with a
        SamAccountName or user property, e.g. from Import-Csv).

    .PARAMETER Date
        Date written to the account description. Defaults to today.

    .PARAMETER ExcludeGroup
        Names (or SamAccountNames) of groups to never remove the user from.

    .PARAMETER TargetOU
        Distinguished name of the OU or container to move the user into.

    .PARAMETER ResetPassword
        Reset the account password to a random 32-character value.

    .PARAMETER HideFromAddressList
        Set msExchHideFromAddressLists on the user. Requires the Exchange schema
        extensions in AD.

    .PARAMETER Server
        Domain controller (or domain) to run the AD operations against.

    .PARAMETER Credential
        Credential to use for the AD operations.

    .PARAMETER PassThru
        Return a result object for each user.

    .NOTES
        Requires RSAT (the ActiveDirectory module).

    .EXAMPLE
        Disable-ADUser -Identity trevor.cooper

    .EXAMPLE
        Disable-ADUser -Identity trevor.cooper -ResetPassword -HideFromAddressList -TargetOU 'OU=Disabled Users,DC=contoso,DC=com'

    .EXAMPLE
        Import-Csv .\leavers.csv | Disable-ADUser -WhatIf

    .EXAMPLE
        $result = Disable-ADUser -Identity trevor.cooper -PassThru
        $result.RemovedGroups
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('SamAccountName', 'User')]
        [ValidateNotNullOrEmpty()]
        [string[]]$Identity,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]$Date = (Get-Date).ToShortDateString(),

        [string[]]$ExcludeGroup = @(),

        [ValidateNotNullOrEmpty()]
        [string]$TargetOU,

        [switch]$ResetPassword,

        [switch]$HideFromAddressList,

        [ValidateNotNullOrEmpty()]
        [string]$Server,

        [pscredential]$Credential,

        [switch]$PassThru
    )

    begin {
        Assert-Module -Name ActiveDirectory -NoInstall
        Import-Module ActiveDirectory -ErrorAction Stop -Verbose:$false

        # Common parameters passed to every AD cmdlet.
        $adParams = @{}
        if ($Server)     { $adParams.Server     = $Server }
        if ($Credential) { $adParams.Credential = $Credential }

        # Extra user attributes to retrieve (explicit array rather than an inline comma list).
        $userProperties = @(
            'Description'
            'MemberOf'
            'UserPrincipalName'
        )

        # Validate the target OU once, before touching any user.
        if ($TargetOU) {
            try {
                $null = Get-ADObject -Identity $TargetOU @adParams -ErrorAction Stop
            }
            catch {
                throw "Target OU not found: $TargetOU ($($_.Exception.Message))"
            }
        }

        # Defined inline so it is not exported as part of the module.
        # Returns a random SecureString containing upper, lower, digit and symbol characters.
        function New-RandomPassword {
            param([int]$Length = 32)

            $charSets = @(
                'ABCDEFGHJKLMNPQRSTUVWXYZ'
                'abcdefghijkmnopqrstuvwxyz'
                '23456789'
                '!@#$%^&*-_=+?'
            )
            $allChars = -join $charSets

            $rng   = [System.Security.Cryptography.RandomNumberGenerator]::Create()
            $bytes = [byte[]]::new(4)
            try {
                $nextIndex = {
                    param([int]$Max)
                    $rng.GetBytes($bytes)
                    [int]([BitConverter]::ToUInt32($bytes, 0) % [uint32]$Max)
                }

                $chars = [System.Collections.Generic.List[char]]::new()

                # Guarantee at least one character from each set (AD complexity rules).
                foreach ($set in $charSets) {
                    $chars.Add($set[(& $nextIndex $set.Length)])
                }
                while ($chars.Count -lt $Length) {
                    $chars.Add($allChars[(& $nextIndex $allChars.Length)])
                }

                # Fisher-Yates shuffle so the guaranteed characters are not always first.
                for ($i = $chars.Count - 1; $i -gt 0; $i--) {
                    $j = & $nextIndex ($i + 1)
                    $temp      = $chars[$i]
                    $chars[$i] = $chars[$j]
                    $chars[$j] = $temp
                }

                $secure = New-Object System.Security.SecureString
                foreach ($c in $chars) { $secure.AppendChar($c) }
                $secure.MakeReadOnly()
                return $secure
            }
            finally {
                $rng.Dispose()
                [Array]::Clear($bytes, 0, $bytes.Length)
            }
        }
    }

    process {
        foreach ($id in $Identity) {

            #----------------------------------------------------------[Lookup]----------------------------------------------------------

            try {
                $adUser = Get-ADUser -Identity $id -Properties $userProperties @adParams -ErrorAction Stop
            }
            catch {
                Write-Error "Unable to find AD user '$id': $($_.Exception.Message)"
                continue
            }

            $sam = $adUser.SamAccountName
            Write-Host "Deprovisioning $sam in Active Directory" -ForegroundColor Cyan

            $removedGroups = [System.Collections.Generic.List[string]]::new()
            $failedGroups  = [System.Collections.Generic.List[string]]::new()
            $disabled      = $false
            $passwordReset = $false
            $hidden        = $false
            $movedTo       = $null

            #----------------------------------------------------------[Disable]----------------------------------------------------------

            if ($PSCmdlet.ShouldProcess($sam, 'Disable AD account')) {
                try {
                    Write-Host '  Disabling account'
                    Disable-ADAccount -Identity $adUser.DistinguishedName @adParams -ErrorAction Stop
                    $disabled = $true
                }
                catch {
                    Write-Error "Could not disable '$sam': $($_.Exception.Message)"
                    continue
                }
            }

            #----------------------------------------------------------[Description]----------------------------------------------------------

            $newDescription = "Disabled on $Date by $env:USERNAME"
            if (-not [string]::IsNullOrWhiteSpace($adUser.Description)) {
                $newDescription += " | Previous: $($adUser.Description)"
            }
            # AD limits the description attribute to 1024 characters.
            if ($newDescription.Length -gt 1024) {
                $newDescription = $newDescription.Substring(0, 1024)
            }

            if ($PSCmdlet.ShouldProcess($sam, "Set description to '$newDescription'")) {
                try {
                    Write-Host '  Updating description'
                    Set-ADUser -Identity $adUser.DistinguishedName -Description $newDescription @adParams -ErrorAction Stop
                }
                catch {
                    Write-Warning "  Could not update description for '$sam': $($_.Exception.Message)"
                }
            }

            #----------------------------------------------------------[Groups]----------------------------------------------------------

            # MemberOf does not include the primary group, which AD will not let us remove anyway.
            Write-Host '  Removing from AD groups'
            $groupDns = @($adUser.MemberOf)

            if ($groupDns.Count -eq 0) {
                Write-Host '    No group memberships'
            }

            foreach ($groupDn in $groupDns) {
                try {
                    $group = Get-ADGroup -Identity $groupDn @adParams -ErrorAction Stop
                }
                catch {
                    Write-Warning "    Could not read group '$groupDn': $($_.Exception.Message)"
                    $failedGroups.Add($groupDn)
                    continue
                }

                if (($ExcludeGroup -contains $group.Name) -or ($ExcludeGroup -contains $group.SamAccountName)) {
                    Write-Host "    Skipping $($group.Name) (excluded)"
                    continue
                }

                if (-not $PSCmdlet.ShouldProcess($sam, "Remove from group '$($group.Name)'")) {
                    continue
                }

                try {
                    Remove-ADGroupMember -Identity $group.DistinguishedName -Members $adUser.DistinguishedName -Confirm:$false @adParams -ErrorAction Stop
                    $removedGroups.Add($group.Name)
                    Write-Host "    Removed from $($group.Name)"
                }
                catch {
                    $failedGroups.Add($group.Name)
                    Write-Warning "    Could not remove $sam from $($group.Name): $($_.Exception.Message)"
                }
            }

            #----------------------------------------------------------[Password]----------------------------------------------------------

            if ($ResetPassword -and $PSCmdlet.ShouldProcess($sam, 'Reset password to a random value')) {
                try {
                    Write-Host '  Resetting password'
                    Set-ADAccountPassword -Identity $adUser.DistinguishedName -Reset -NewPassword (New-RandomPassword) @adParams -ErrorAction Stop
                    $passwordReset = $true
                }
                catch {
                    Write-Warning "  Could not reset password for '$sam': $($_.Exception.Message)"
                }
            }

            #----------------------------------------------------------[Address lists]----------------------------------------------------------

            if ($HideFromAddressList -and $PSCmdlet.ShouldProcess($sam, 'Hide from Exchange address lists')) {
                try {
                    Write-Host '  Hiding from address lists'
                    Set-ADUser -Identity $adUser.DistinguishedName -Replace @{ msExchHideFromAddressLists = $true } @adParams -ErrorAction Stop
                    $hidden = $true
                }
                catch {
                    Write-Warning "  Could not hide '$sam' from address lists (is the Exchange schema present?): $($_.Exception.Message)"
                }
            }

            #----------------------------------------------------------[Move]----------------------------------------------------------

            # Done last: moving changes the user's DistinguishedName.
            $finalDn = $adUser.DistinguishedName
            if ($TargetOU -and $PSCmdlet.ShouldProcess($sam, "Move to '$TargetOU'")) {
                try {
                    Write-Host "  Moving to $TargetOU"
                    $moved   = Move-ADObject -Identity $adUser.DistinguishedName -TargetPath $TargetOU -PassThru @adParams -ErrorAction Stop
                    $finalDn = $moved.DistinguishedName
                    $movedTo = $TargetOU
                }
                catch {
                    Write-Warning "  Could not move '$sam' to '$TargetOU': $($_.Exception.Message)"
                }
            }

            if ($failedGroups.Count -gt 0) {
                Write-Warning "  $sam is still a member of $($failedGroups.Count) group(s): $($failedGroups -join ', ')"
            }

            if ($PassThru) {
                [pscustomobject]@{
                    SamAccountName         = $sam
                    UserPrincipalName      = $adUser.UserPrincipalName
                    DistinguishedName      = $finalDn
                    Disabled               = $disabled
                    Description            = $newDescription
                    RemovedGroups          = @($removedGroups)
                    FailedGroups           = @($failedGroups)
                    PasswordReset          = $passwordReset
                    HiddenFromAddressLists = $hidden
                    MovedTo                = $movedTo
                }
            }
        }
    }
}

function Remove-IntuneDevice
{
    <#
    .SYNOPSIS
        Removes devices from Intune by serial number, and optionally from Windows Autopilot and Entra ID.

    .DESCRIPTION
        Finds Intune managed devices by serial number and deletes them. Optionally
        also deregisters the device from Windows Autopilot and deletes the device
        object from Entra ID.

        Records are always removed in the order Microsoft recommends:
          1. Intune
          2. Windows Autopilot   (-IncludeAutopilot, or implied by -IncludeEntra)
          3. Entra ID            (-IncludeEntra)

        -IncludeEntra always runs the Autopilot step first. The Entra step only
        starts once the Autopilot identity is confirmed to be gone. Autopilot
        deletion is asynchronous, so the command polls until the identity
        disappears or -TimeoutSeconds is reached. If a step fails, is declined,
        or times out, the later steps are skipped for that device.

        -Hostname is an optional safety check. If the serial number has Intune
        records but none of them have that device name, the device is skipped.
        If the serial number has no Intune record, the hostname cannot be
        verified and processing continues with Autopilot / Entra.

        Entra devices are found through the Entra device IDs on the Intune and
        Autopilot records. If none are found and a hostname was given, the Entra
        device is looked up by display name, but only if exactly one device has
        that name. Synced (hybrid joined) Entra devices are not deleted; remove
        them from on-premises Active Directory instead.

        Batches (CSV or pipeline input) are processed in phases: every Intune
        deletion, then every Autopilot deletion, then one shared wait, then every
        Entra deletion. You wait once for the whole batch, not once per device.

        Requested Microsoft Graph scopes:
          DeviceManagementManagedDevices.ReadWrite.All   (always)
          DeviceManagementServiceConfig.ReadWrite.All    (Autopilot)
          Device.ReadWrite.All                           (Entra)

    .PARAMETER SerialNumber
        Serial number of the device. Accepts pipeline input by property name
        (alias: Serial).

    .PARAMETER Hostname
        Optional device name used as a safety check against the Intune record.
        Accepts pipeline input by property name (aliases: DeviceName, ComputerName).

    .PARAMETER CsvPath
        Path to a CSV file with a 'SerialNumber' (or 'Serial') column and an
        optional 'Hostname' (or 'DeviceName' / 'ComputerName') column.

    .PARAMETER IncludeAutopilot
        Also deregister the device from Windows Autopilot (after Intune).

    .PARAMETER IncludeEntra
        Also delete the device from Entra ID. Implies -IncludeAutopilot; Entra
        deletion only runs after the Autopilot identity is confirmed removed.

    .PARAMETER TimeoutSeconds
        How long to wait for Autopilot deletions to complete (and to retry
        Autopilot deletion while a just-deleted Intune record is still being
        processed). Default 600.

    .PARAMETER PollIntervalSeconds
        Delay between Autopilot status checks / retries. Default 15.

    .PARAMETER Force
        Do not prompt for confirmation before each deletion.

    .PARAMETER UseWAM
        Passed to Assert-MgGraphConnection: do not disable the Windows Web
        Account Manager broker for sign-in.

    .OUTPUTS
        One object per input row with SerialNumber, Hostname, IntuneDevice,
        IntuneStatus, AutopilotStatus, EntraDevice, EntraStatus, Notes and Source.
        Status values: Removed, NotFound, WhatIf, Skipped, Failed, TimedOut,
        NotRequested.

    .EXAMPLE
        Remove-IntuneDevice -SerialNumber 5CG1234ABC -WhatIf

        Shows which Intune records would be deleted for the serial number.

    .EXAMPLE
        Remove-IntuneDevice -SerialNumber 5CG1234ABC -Hostname LAPTOP-042 -IncludeEntra

        Deletes the Intune record (only if it is named LAPTOP-042), deregisters
        the device from Autopilot, waits for that to finish, then deletes the
        Entra device.

    .EXAMPLE
        Remove-IntuneDevice -CsvPath .\retired.csv -IncludeEntra -Force | Export-Csv .\retired-results.csv -NoTypeInformation

        Batch removal from a CSV without prompting, saving the results.

    .EXAMPLE
        Import-Csv .\retired.csv | Remove-IntuneDevice -IncludeAutopilot

        Same as -CsvPath, using pipeline input.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'Device')]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Device', Position = 0, ValueFromPipelineByPropertyName)]
        [Alias('Serial')]
        [AllowEmptyString()]
        [string]$SerialNumber,

        [Parameter(ParameterSetName = 'Device', Position = 1, ValueFromPipelineByPropertyName)]
        [Alias('DeviceName', 'ComputerName')]
        [AllowEmptyString()]
        [AllowNull()]
        [string]$Hostname,

        [Parameter(Mandatory, ParameterSetName = 'Csv')]
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
        [string]$CsvPath,

        [switch]$IncludeAutopilot,

        [switch]$IncludeEntra,

        [ValidateRange(30, 3600)]
        [int]$TimeoutSeconds = 600,

        [ValidateRange(5, 300)]
        [int]$PollIntervalSeconds = 15,

        [switch]$Force,

        [switch]$UseWAM
    )

    begin
    {
        if ($Force -and -not $PSBoundParameters.ContainsKey('Confirm'))
        {
            $ConfirmPreference = 'None'
        }

        $doEntra     = [bool]$IncludeEntra
        $doAutopilot = [bool]$IncludeAutopilot -or $doEntra

        if ($doEntra -and -not $IncludeAutopilot)
        {
            Write-Verbose '-IncludeEntra specified: Autopilot removal will run before Entra removal.'
        }

        $requests = [System.Collections.Generic.List[object]]::new()
    }

    process
    {
        if ($PSCmdlet.ParameterSetName -eq 'Csv')
        {
            $rows = @(Import-Csv -LiteralPath $CsvPath)
            if ($rows.Count -eq 0)
            {
                Write-Warning "CSV '$CsvPath' contains no rows."
                return
            }

            $columns      = @($rows[0].PSObject.Properties.Name)
            $serialColumn = $columns | Where-Object { $_ -in 'SerialNumber', 'Serial', 'Serial Number' } | Select-Object -First 1
            $hostColumn   = $columns | Where-Object { $_ -in 'Hostname', 'Host Name', 'DeviceName', 'Device Name', 'ComputerName' } | Select-Object -First 1

            if (-not $serialColumn)
            {
                throw "CSV '$CsvPath' must have a 'SerialNumber' column. Columns found: $($columns -join ', ')."
            }

            $rowNumber = 0
            foreach ($row in $rows)
            {
                $rowNumber++
                $rowHost = ''
                if ($hostColumn) { $rowHost = ([string]$row.$hostColumn).Trim() }

                $requests.Add([pscustomobject]@{
                    SerialNumber = ([string]$row.$serialColumn).Trim()
                    Hostname     = $rowHost
                    Source       = "$(Split-Path -Path $CsvPath -Leaf) row $rowNumber"
                })
            }
        }
        else
        {
            $requests.Add([pscustomobject]@{
                SerialNumber = ([string]$SerialNumber).Trim()
                Hostname     = ([string]$Hostname).Trim()
                Source       = if ($MyInvocation.ExpectingInput) { 'pipeline' } else { 'parameter' }
            })
        }
    }

    end
    {
        if ($requests.Count -eq 0) { return }

        #region Helpers

        $baseUri = 'https://graph.microsoft.com/v1.0'

        # GET with paging; outputs each item (hashtable).
        $getAll = {
            param([string]$Uri)
            $next = $Uri
            while ($next)
            {
                $response = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType HashTable -ErrorAction Stop
                foreach ($item in @($response['value']))
                {
                    if ($null -ne $item) { $item }
                }
                $next = $response['@odata.nextLink']
            }
        }

        $odataFilter = {
            param([string]$Expression)
            '?$filter=' + [uri]::EscapeDataString($Expression)
        }

        $quote = {
            param([string]$Value)
            $Value.Replace("'", "''")
        }

        $isNotFound = {
            param($ErrorRecord)
            $response = $ErrorRecord.Exception.Response
            if ($response -and $response.StatusCode)
            {
                try { return ([int]$response.StatusCode -eq 404) } catch { }
            }
            return ([string]$ErrorRecord.Exception.Message -match '\b404\b|NotFound|ResourceNotFound')
        }

        $newCount = { @{ Removed = 0; Failed = 0; Declined = 0; TimedOut = 0; Skipped = 0 } }

        $summarize = {
            param([hashtable]$Count)
            if ($Count.Failed -gt 0)   { return 'Failed' }
            if ($Count.TimedOut -gt 0) { return 'TimedOut' }
            if ($Count.Declined -gt 0)
            {
                if ($WhatIfPreference) { return 'WhatIf' }
                return 'Skipped'
            }
            if ($Count.Skipped -gt 0)  { return 'Skipped' }
            return 'Removed'
        }

        # A later step may only run once the earlier step is cleared.
        # Under -WhatIf, continue so every step is previewed.
        $canProceed = {
            param([string]$Status)
            ($Status -in 'Removed', 'NotFound') -or ($WhatIfPreference -and $Status -eq 'WhatIf')
        }

        $addNote = {
            param($Record, [string]$Text, [switch]$Warn)
            $Record.Notes.Add($Text)
            if ($Warn) { Write-Warning "[$($Record.SerialNumber)] $Text" }
        }

        #endregion

        $scopes = @('DeviceManagementManagedDevices.ReadWrite.All')
        if ($doAutopilot) { $scopes += 'DeviceManagementServiceConfig.ReadWrite.All' }
        if ($doEntra)     { $scopes += 'Device.ReadWrite.All' }

        Assert-MgGraphConnection -Scopes $scopes -UseWAM:$UseWAM

        $steps = @('Intune')
        if ($doAutopilot) { $steps += 'Autopilot' }
        if ($doEntra)     { $steps += 'Entra ID' }
        Write-Host "Processing $($requests.Count) device(s): $($steps -join ' -> ')" -ForegroundColor DarkCyan

        $activity = 'Remove-IntuneDevice'

        #region Phase 1: look up records

        $records = [System.Collections.Generic.List[object]]::new()
        $seen    = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $index   = 0

        foreach ($request in $requests)
        {
            $index++
            Write-Progress -Activity $activity -Status "Looking up $($request.SerialNumber) ($index of $($requests.Count))" `
                -PercentComplete ([int](($index / $requests.Count) * 100))

            $record = @{
                SerialNumber     = $request.SerialNumber
                Hostname         = $request.Hostname
                Source           = $request.Source
                IntuneDevices    = @()
                AutopilotDevices = @()
                EntraDevices     = @()
                IntuneStatus     = 'Queued'
                AutopilotStatus  = if ($doAutopilot) { 'Queued' } else { 'NotRequested' }
                EntraStatus      = if ($doEntra) { 'Queued' } else { 'NotRequested' }
                Counts           = @{ Intune = (& $newCount); Autopilot = (& $newCount); Entra = (& $newCount) }
                Notes            = [System.Collections.Generic.List[string]]::new()
            }
            $records.Add($record)

            $skipAll = {
                param($Record, [string]$Status)
                $Record.IntuneStatus = $Status
                if ($doAutopilot) { $Record.AutopilotStatus = $Status }
                if ($doEntra)     { $Record.EntraStatus = $Status }
            }

            if ([string]::IsNullOrWhiteSpace($record.SerialNumber))
            {
                & $skipAll $record 'Skipped'
                & $addNote $record "No serial number ($($record.Source))." -Warn
                continue
            }

            if (-not $seen.Add("$($record.SerialNumber)|$($record.Hostname)"))
            {
                & $skipAll $record 'Skipped'
                & $addNote $record 'Duplicate of an earlier row; skipped.'
                continue
            }

            try
            {
                $serialQ = & $quote $record.SerialNumber

                # Intune
                $intuneSelect = '&$select=id,deviceName,serialNumber,azureADDeviceId,operatingSystem,lastSyncDateTime,userPrincipalName'
                $intune = @(& $getAll ("$baseUri/deviceManagement/managedDevices" + (& $odataFilter "serialNumber eq '$serialQ'") + $intuneSelect))
                $intune = @($intune | Where-Object { [string]$_['serialNumber'] -eq $record.SerialNumber })

                if ($record.Hostname)
                {
                    $matching = @($intune | Where-Object { [string]$_['deviceName'] -eq $record.Hostname })
                    if ($intune.Count -gt 0 -and $matching.Count -eq 0)
                    {
                        $names = @($intune | ForEach-Object { [string]$_['deviceName'] }) -join ', '
                        & $skipAll $record 'Skipped'
                        & $addNote $record "Hostname '$($record.Hostname)' does not match the Intune device name(s) for this serial ($names). Nothing removed." -Warn
                        continue
                    }
                    if ($intune.Count -eq 0)
                    {
                        & $addNote $record 'No Intune record; hostname could not be verified.'
                    }
                    $intune = $matching
                }

                if ($intune.Count -gt 1)
                {
                    & $addNote $record "$($intune.Count) Intune records found for this serial."
                }

                $record.IntuneDevices = $intune
                if ($intune.Count -eq 0) { $record.IntuneStatus = 'NotFound' }

                # Autopilot (the service only supports contains() on serialNumber; match exactly here).
                if ($doAutopilot)
                {
                    $autopilot = @(& $getAll ("$baseUri/deviceManagement/windowsAutopilotDeviceIdentities" + (& $odataFilter "contains(serialNumber,'$serialQ')")))
                    $autopilot = @($autopilot | Where-Object { [string]$_['serialNumber'] -eq $record.SerialNumber })
                    $record.AutopilotDevices = $autopilot
                    if ($autopilot.Count -eq 0) { $record.AutopilotStatus = 'NotFound' }
                }

                # Entra ID
                if ($doEntra)
                {
                    $entraSelect = '&$select=id,deviceId,displayName,operatingSystem,trustType,onPremisesSyncEnabled'
                    $emptyGuid   = '00000000-0000-0000-0000-000000000000'

                    $entraIds = @(@($record.IntuneDevices | ForEach-Object { [string]$_['azureADDeviceId'] }) +
                                  @($record.AutopilotDevices | ForEach-Object { [string]$_['azureActiveDirectoryDeviceId'] }))
                    $entraIds = @($entraIds | Where-Object { $_ -and $_ -ne $emptyGuid } | Sort-Object -Unique)

                    $entra = @(foreach ($entraId in $entraIds)
                    {
                        & $getAll ("$baseUri/devices" + (& $odataFilter "deviceId eq '$(& $quote $entraId)'") + $entraSelect)
                    })

                    if ($entra.Count -eq 0 -and $record.Hostname)
                    {
                        $byName = @(& $getAll ("$baseUri/devices" + (& $odataFilter "displayName eq '$(& $quote $record.Hostname)'") + $entraSelect))
                        if ($byName.Count -eq 1)
                        {
                            $entra = $byName
                            & $addNote $record 'Entra device matched by hostname.'
                        }
                        elseif ($byName.Count -gt 1)
                        {
                            & $addNote $record "$($byName.Count) Entra devices are named '$($record.Hostname)'; none removed. Remove the correct one manually." -Warn
                        }
                    }

                    # De-duplicate by object id.
                    $unique = @{}
                    foreach ($device in $entra) { $unique[[string]$device['id']] = $device }
                    $record.EntraDevices = @($unique.Values)

                    if ($record.EntraDevices.Count -eq 0) { $record.EntraStatus = 'NotFound' }
                }
            }
            catch
            {
                & $skipAll $record 'Failed'
                & $addNote $record "Lookup failed: $($_.Exception.Message)" -Warn
            }
        }

        #endregion

        #region Phase 2: Intune

        foreach ($record in $records)
        {
            if ($record.IntuneStatus -ne 'Queued') { continue }
            $count = $record.Counts.Intune

            foreach ($device in $record.IntuneDevices)
            {
                $target = "Intune device '$($device['deviceName'])' (serial $($record.SerialNumber), user $($device['userPrincipalName']), id $($device['id']))"
                if (-not $PSCmdlet.ShouldProcess($target, 'Delete'))
                {
                    $count.Declined++
                    continue
                }

                try
                {
                    Invoke-MgGraphRequest -Method DELETE -Uri "$baseUri/deviceManagement/managedDevices/$($device['id'])" -ErrorAction Stop | Out-Null
                    Write-Verbose "Deleted $target."
                    $count.Removed++
                }
                catch
                {
                    if (& $isNotFound $_)
                    {
                        $count.Removed++
                    }
                    else
                    {
                        $count.Failed++
                        & $addNote $record "Intune delete failed for '$($device['deviceName'])': $($_.Exception.Message)" -Warn
                    }
                }
            }

            $record.IntuneStatus = & $summarize $count
        }

        #endregion

        #region Phase 3 + 4: Autopilot (delete, then wait until gone)

        if ($doAutopilot)
        {
            $pending = [System.Collections.Generic.List[object]]::new()

            foreach ($record in $records)
            {
                if ($record.AutopilotStatus -ne 'Queued') { continue }
                $count = $record.Counts.Autopilot

                if (-not (& $canProceed $record.IntuneStatus))
                {
                    $record.AutopilotStatus = 'Skipped'
                    & $addNote $record "Intune removal did not complete ($($record.IntuneStatus)); Autopilot not removed."
                    continue
                }

                # Retry only when an Intune record was just deleted; Autopilot may reject
                # the deletion until Intune finishes processing it.
                $retry = $record.Counts.Intune.Removed -gt 0

                foreach ($identity in $record.AutopilotDevices)
                {
                    $target = "Autopilot device (serial $($record.SerialNumber), model $($identity['model']), id $($identity['id']))"
                    if (-not $PSCmdlet.ShouldProcess($target, 'Deregister'))
                    {
                        $count.Declined++
                        continue
                    }

                    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
                    $done     = $false
                    while (-not $done)
                    {
                        try
                        {
                            Invoke-MgGraphRequest -Method DELETE -Uri "$baseUri/deviceManagement/windowsAutopilotDeviceIdentities/$($identity['id'])" -ErrorAction Stop | Out-Null
                            Write-Verbose "Requested deletion of $target."
                            $pending.Add(@{ Record = $record; Id = [string]$identity['id'] })
                            $done = $true
                        }
                        catch
                        {
                            if (& $isNotFound $_)
                            {
                                $count.Removed++
                                $done = $true
                            }
                            elseif (-not $retry -or (Get-Date) -ge $deadline)
                            {
                                $count.Failed++
                                & $addNote $record "Autopilot delete failed: $($_.Exception.Message)" -Warn
                                $done = $true
                            }
                            else
                            {
                                Write-Verbose "Autopilot delete for $($record.SerialNumber) not accepted yet ($($_.Exception.Message)); retrying in $PollIntervalSeconds s."
                                Write-Progress -Activity $activity -Status "Waiting for Intune to release $($record.SerialNumber) before Autopilot removal"
                                Start-Sleep -Seconds $PollIntervalSeconds
                            }
                        }
                    }
                }
            }

            if ($pending.Count -gt 0)
            {
                Write-Host "Waiting for $($pending.Count) Autopilot deletion(s) to complete (timeout $TimeoutSeconds s)..." -ForegroundColor DarkCyan
                $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
                $total    = $pending.Count

                while ($pending.Count -gt 0)
                {
                    foreach ($item in @($pending))
                    {
                        try
                        {
                            Invoke-MgGraphRequest -Method GET -Uri "$baseUri/deviceManagement/windowsAutopilotDeviceIdentities/$($item.Id)" -OutputType HashTable -ErrorAction Stop | Out-Null
                        }
                        catch
                        {
                            if (& $isNotFound $_)
                            {
                                $item.Record.Counts.Autopilot.Removed++
                                [void]$pending.Remove($item)
                            }
                            else
                            {
                                Write-Verbose "Autopilot status check failed for $($item.Record.SerialNumber): $($_.Exception.Message)"
                            }
                        }
                    }

                    if ($pending.Count -eq 0 -or (Get-Date) -ge $deadline) { break }

                    $doneCount = $total - $pending.Count
                    Write-Progress -Activity $activity -Status "Autopilot deletions confirmed: $doneCount of $total" `
                        -PercentComplete ([int](($doneCount / $total) * 100))
                    Start-Sleep -Seconds $PollIntervalSeconds
                }

                foreach ($item in $pending)
                {
                    $item.Record.Counts.Autopilot.TimedOut++
                    & $addNote $item.Record "Autopilot deletion requested but still present after $TimeoutSeconds s. Run again later to finish." -Warn
                }
            }

            foreach ($record in $records)
            {
                if ($record.AutopilotStatus -eq 'Queued')
                {
                    $record.AutopilotStatus = & $summarize $record.Counts.Autopilot
                }
            }
        }

        #endregion

        #region Phase 5: Entra ID

        if ($doEntra)
        {
            foreach ($record in $records)
            {
                if ($record.EntraStatus -ne 'Queued') { continue }
                $count = $record.Counts.Entra

                if (-not (& $canProceed $record.AutopilotStatus))
                {
                    $record.EntraStatus = 'Skipped'
                    & $addNote $record "Autopilot removal not confirmed ($($record.AutopilotStatus)); Entra device not removed."
                    continue
                }

                foreach ($device in $record.EntraDevices)
                {
                    if ($device['onPremisesSyncEnabled'])
                    {
                        $count.Skipped++
                        & $addNote $record "Entra device '$($device['displayName'])' is synced from on-premises AD (hybrid joined); remove it from AD instead." -Warn
                        continue
                    }

                    $target = "Entra device '$($device['displayName'])' (serial $($record.SerialNumber), deviceId $($device['deviceId']))"
                    if (-not $PSCmdlet.ShouldProcess($target, 'Delete'))
                    {
                        $count.Declined++
                        continue
                    }

                    try
                    {
                        Invoke-MgGraphRequest -Method DELETE -Uri "$baseUri/devices/$($device['id'])" -ErrorAction Stop | Out-Null
                        Write-Verbose "Deleted $target."
                        $count.Removed++
                    }
                    catch
                    {
                        if (& $isNotFound $_)
                        {
                            $count.Removed++
                        }
                        else
                        {
                            $count.Failed++
                            & $addNote $record "Entra delete failed for '$($device['displayName'])': $($_.Exception.Message)" -Warn
                        }
                    }
                }

                $record.EntraStatus = & $summarize $count
            }
        }

        #endregion

        Write-Progress -Activity $activity -Completed

        foreach ($record in $records)
        {
            [pscustomobject]@{
                SerialNumber    = $record.SerialNumber
                Hostname        = $record.Hostname
                IntuneDevice    = @($record.IntuneDevices | ForEach-Object { [string]$_['deviceName'] }) -join ', '
                IntuneStatus    = $record.IntuneStatus
                AutopilotStatus = $record.AutopilotStatus
                EntraDevice     = @($record.EntraDevices | ForEach-Object { [string]$_['displayName'] }) -join ', '
                EntraStatus     = $record.EntraStatus
                Notes           = $record.Notes -join ' '
                Source          = $record.Source
            }
        }
    }
}

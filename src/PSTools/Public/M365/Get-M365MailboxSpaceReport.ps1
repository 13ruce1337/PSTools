function Get-M365MailboxSpaceReport {
    <#
    .SYNOPSIS
        Analyzes where space is being used in an Exchange Online mailbox.

    .DESCRIPTION
        Produces a mailbox space report containing:
            1. Mailbox overview (size, item counts, deleted items, last logon, quotas)
            2. All folders sorted by size (largest first)
            3. Recoverable Items folders (a common hidden space hog)
            4. Archive mailbox statistics (if an archive is enabled)
            5. Sent Items folder detail
            6. Large sent messages over the last -Days days, via Get-MessageTraceV2
               (queried in chunks of -ChunkDays, since MessageTraceV2 allows at
               most 10 days per query and 90 days of history)

        Message trace returns one row per recipient; rows are grouped by
        MessageId so each message is listed once with its recipients.

        Folder statistics and large sent messages are exported to CSV in
        -OutputPath unless -NoExport is specified.

    .PARAMETER Identity
        The mailbox to analyze (email address, UPN or alias). Accepts pipeline
        input by PrimarySmtpAddress.

    .PARAMETER Days
        How many days of message trace history to search (1-90). Default 90.

    .PARAMETER MinimumSizeMB
        Only sent messages larger than this size (in MB) are reported. Default 5.

    .PARAMETER ChunkDays
        Size of each message trace query window in days (1-10). Default 10.

    .PARAMETER OutputPath
        Folder for the CSV exports. Created if it does not exist.
        Default: $env:TEMP\PSTools

    .PARAMETER NoExport
        Do not write CSV files.

    .PARAMETER PassThru
        Return a report object in addition to the on-screen report.

    .PARAMETER UserPrincipalName
        Optional UPN used if a new Exchange Online connection is required.

    .EXAMPLE
        Get-M365MailboxSpaceReport -Identity trevor.cooper@contoso.com

    .EXAMPLE
        Get-M365MailboxSpaceReport -Identity trevor.cooper@contoso.com -Days 30 -MinimumSizeMB 10 -OutputPath C:\Temp

    .EXAMPLE
        $report = Get-M365MailboxSpaceReport -Identity trevor.cooper@contoso.com -NoExport -PassThru
        $report.Folders | Select-Object -First 5

    .NOTES
        Requires ExchangeOnlineManagement 3.7.0+ for Get-MessageTraceV2. If it is
        not available, the large sent messages section is skipped.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('UserEmail', 'PrimarySmtpAddress', 'Mailbox')]
        [string]$Identity,

        [ValidateRange(1, 90)]
        [int]$Days = 90,

        [ValidateRange(0, 150)]
        [double]$MinimumSizeMB = 5,

        [ValidateRange(1, 10)]
        [int]$ChunkDays = 10,

        [ValidateNotNullOrEmpty()]
        [string]$OutputPath = (Join-Path $env:TEMP 'PSTools'),

        [switch]$NoExport,

        [switch]$PassThru,

        [string]$UserPrincipalName
    )

    begin {
        Assert-ExchangeOnlineConnection -UserPrincipalName $UserPrincipalName

        # Explicit byte multiplier (avoids relying on the 1MB numeric literal suffix).
        [int64]$bytesPerMB = 1048576

        # Parses "1.23 MB (1,234,567 bytes)" (ByteQuantifiedSize or its string form) into bytes.
        function ConvertTo-ByteCount {
            param($Value)
            if ($null -eq $Value) { return [int64]0 }
            $match = [regex]::Match([string]$Value, '\(([\d,]+) bytes\)')
            if ($match.Success) {
                return [int64]($match.Groups[1].Value -replace ',', '')
            }
            return [int64]0
        }

        $hasMessageTraceV2 = [bool](Get-Command -Name Get-MessageTraceV2 -ErrorAction SilentlyContinue)
        if (-not $hasMessageTraceV2) {
            Write-Warning 'Get-MessageTraceV2 is not available (requires ExchangeOnlineManagement 3.7.0+). The large sent messages section will be skipped.'
        }

        if (-not $NoExport -and -not (Test-Path -LiteralPath $OutputPath)) {
            $null = New-Item -ItemType Directory -Path $OutputPath -Force
        }

        $minimumBytes = [int64]($MinimumSizeMB * $bytesPerMB)

        # Property lists (kept as explicit arrays rather than inline comma lists).
        $mailboxProperties = @(
            'IssueWarningQuota'
            'ProhibitSendQuota'
            'ProhibitSendReceiveQuota'
            'RecoverableItemsQuota'
            'RecoverableItemsWarningQuota'
            'ArchiveGuid'
            'ArchiveStatus'
            'ArchiveQuota'
        )

        # LastLogonTime is not returned by Get-EXOMailboxStatistics unless requested.
        $statsProperties = @(
            'LastLogonTime'
        )

        $folderSizeMBProperty = @{
            Name       = 'SizeMB'
            Expression = { [math]::Round((ConvertTo-ByteCount $_.FolderSize) / $bytesPerMB, 2) }
        }

        $folderProperties = @(
            'Name'
            'FolderPath'
            'FolderType'
            'ItemsInFolder'
            'FolderSize'
            $folderSizeMBProperty
        )

        $folderTableSizeColumn = @{
            Label      = 'Size(MB)'
            Expression = { $_.SizeMB }
        }

        $folderTableProperties = @(
            'Name'
            'FolderPath'
            'ItemsInFolder'
            'FolderSize'
            $folderTableSizeColumn
        )

        $recoverableSizeMBProperty = @{
            Name       = 'SizeMB'
            Expression = { [math]::Round((ConvertTo-ByteCount $_.FolderAndSubfolderSize) / $bytesPerMB, 2) }
        }

        $recoverableProperties = @(
            'Name'
            'FolderPath'
            'ItemsInFolderAndSubfolders'
            'FolderAndSubfolderSize'
            $recoverableSizeMBProperty
        )

        $sentItemsProperties = @(
            'Name'
            'FolderPath'
            'ItemsInFolder'
            'FolderSize'
            'FolderAndSubfolderSize'
        )

        $largeSentTableProperties = @(
            'Received'
            'SizeMB'
            'RecipientCount'
            'Subject'
            'Recipients'
        )
    }

    process {
        try {
            $mailbox = Get-EXOMailbox -Identity $Identity -Properties $mailboxProperties -ErrorAction Stop
        }
        catch {
            Write-Error "Unable to find mailbox '$Identity': $($_.Exception.Message)"
            return
        }

        $address   = [string]$mailbox.PrimarySmtpAddress
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmm'
        $fileStem  = $address.Split('@')[0]

        Write-Host "`n===================================================" -ForegroundColor Cyan
        Write-Host " Mailbox Space Report for: $address" -ForegroundColor Cyan
        Write-Host "===================================================`n" -ForegroundColor Cyan

        #----------------------------------------------------------[1. Overview]----------------------------------------------------------

        Write-Host "`n--- Mailbox Overview ---" -ForegroundColor Yellow
        $stats = Get-EXOMailboxStatistics -Identity $address -Properties $statsProperties -ErrorAction Stop

        $overview = [pscustomobject]@{
            DisplayName                  = $stats.DisplayName
            PrimarySmtpAddress           = $address
            LastLogonTime                = $stats.LastLogonTime
            TotalItemSize                = $stats.TotalItemSize
            ItemCount                    = $stats.ItemCount
            TotalDeletedItemSize         = $stats.TotalDeletedItemSize
            DeletedItemCount             = $stats.DeletedItemCount
            IssueWarningQuota            = $mailbox.IssueWarningQuota
            ProhibitSendQuota            = $mailbox.ProhibitSendQuota
            ProhibitSendReceiveQuota     = $mailbox.ProhibitSendReceiveQuota
            RecoverableItemsWarningQuota = $mailbox.RecoverableItemsWarningQuota
            RecoverableItemsQuota        = $mailbox.RecoverableItemsQuota
        }
        $overview | Format-List | Out-Host

        #----------------------------------------------------------[2. Folders by size]----------------------------------------------------------

        Write-Host "`n--- All Folders Sorted by Size (Largest First) ---" -ForegroundColor Yellow
        $rawFolders  = Get-EXOMailboxFolderStatistics -Identity $address -FolderScope All -ErrorAction Stop
        $folderStats = @($rawFolders | Select-Object -Property $folderProperties | Sort-Object -Property SizeMB -Descending)

        $folderStats | Format-Table -Property $folderTableProperties -AutoSize | Out-Host

        #----------------------------------------------------------[3. Recoverable Items]----------------------------------------------------------

        Write-Host "`n--- Recoverable Items Folders ---" -ForegroundColor Yellow
        $recoverableItems = @()
        try {
            $rawRecoverable   = Get-EXOMailboxFolderStatistics -Identity $address -FolderScope RecoverableItems -ErrorAction Stop
            $recoverableItems = @($rawRecoverable | Select-Object -Property $recoverableProperties | Sort-Object -Property SizeMB -Descending)
            $recoverableItems | Format-Table -AutoSize | Out-Host
        }
        catch {
            Write-Warning "Unable to read Recoverable Items folders: $($_.Exception.Message)"
        }

        #----------------------------------------------------------[4. Archive mailbox]----------------------------------------------------------

        Write-Host "`n--- Archive Mailbox ---" -ForegroundColor Yellow
        $archive = $null

        # ArchiveStatus is unreliable (often 'None' for active archives); ArchiveGuid is the dependable check.
        $hasArchive = $mailbox.ArchiveGuid -and ([guid]$mailbox.ArchiveGuid -ne [guid]::Empty)

        if ($hasArchive) {
            try {
                $archiveStats = Get-EXOMailboxStatistics -Identity $address -Archive -ErrorAction Stop
                $archive = [pscustomobject]@{
                    DisplayName          = $archiveStats.DisplayName
                    TotalItemSize        = $archiveStats.TotalItemSize
                    ItemCount            = $archiveStats.ItemCount
                    TotalDeletedItemSize = $archiveStats.TotalDeletedItemSize
                    DeletedItemCount     = $archiveStats.DeletedItemCount
                    ArchiveStatus        = $mailbox.ArchiveStatus
                    ArchiveQuota         = $mailbox.ArchiveQuota
                }
                $archive | Format-List | Out-Host
            }
            catch {
                Write-Warning "Archive appears to be enabled but statistics could not be retrieved: $($_.Exception.Message)"
            }
        }
        else {
            Write-Host "No archive mailbox enabled." -ForegroundColor Gray
        }

        #----------------------------------------------------------[5. Sent Items]----------------------------------------------------------

        Write-Host "`n--- Sent Items Folder Detail ---" -ForegroundColor Yellow
        $rawSentItems = Get-EXOMailboxFolderStatistics -Identity $address -FolderScope SentItems -ErrorAction Stop
        $sentItems    = @($rawSentItems | Select-Object -Property $sentItemsProperties)

        $sentItems | Format-Table -AutoSize | Out-Host

        #----------------------------------------------------------[6. Large sent messages]----------------------------------------------------------

        $largeSent = @()

        if ($hasMessageTraceV2) {
            Write-Host "`n--- Large Sent Messages (Last $Days Days, >$($MinimumSizeMB)MB) via MessageTraceV2 ---" -ForegroundColor Yellow

            $traceResults = [System.Collections.Generic.List[object]]::new()
            $traceEnd     = Get-Date
            $chunkStart   = $traceEnd.AddDays(-$Days)

            while ($chunkStart -lt $traceEnd) {
                $chunkEnd = $chunkStart.AddDays($ChunkDays)
                if ($chunkEnd -gt $traceEnd) { $chunkEnd = $traceEnd }

                Write-Host "  Querying $($chunkStart.ToString('yyyy-MM-dd')) to $($chunkEnd.ToString('yyyy-MM-dd'))..." -ForegroundColor DarkGray

                try {
                    $results = @(Get-MessageTraceV2 -SenderAddress $address -StartDate $chunkStart -EndDate $chunkEnd -ResultSize 5000 -ErrorAction Stop)
                    if ($results.Count -ge 5000) {
                        Write-Warning "  Window $($chunkStart.ToString('yyyy-MM-dd')) to $($chunkEnd.ToString('yyyy-MM-dd')) hit the 5000 result limit; results may be incomplete. Try a smaller -ChunkDays."
                    }
                    foreach ($r in $results) { $traceResults.Add($r) }
                }
                catch {
                    Write-Warning "  Message trace failed for $($chunkStart.ToString('yyyy-MM-dd')) to $($chunkEnd.ToString('yyyy-MM-dd')): $($_.Exception.Message)"
                }

                $chunkStart = $chunkEnd
            }

            # Message trace returns one row per recipient; collapse to one row per message.
            $largeSent = @($traceResults |
                Where-Object { $_.Size -gt $minimumBytes } |
                Group-Object -Property { if ($_.MessageId) { $_.MessageId } else { $_.MessageTraceId } } |
                ForEach-Object {
                    $first      = $_.Group | Sort-Object -Property Received | Select-Object -First 1
                    $recipients = @($_.Group.RecipientAddress | Sort-Object -Unique)
                    [pscustomobject]@{
                        Received       = $first.Received
                        Subject        = $first.Subject
                        SizeMB         = [math]::Round($first.Size / $bytesPerMB, 2)
                        RecipientCount = $recipients.Count
                        Recipients     = $recipients -join '; '
                        MessageId      = $first.MessageId
                    }
                } |
                Sort-Object -Property SizeMB -Descending)

            if ($largeSent.Count -gt 0) {
                $largeSent | Format-Table -Property $largeSentTableProperties -AutoSize | Out-Host
            }
            else {
                Write-Host "No individual messages over $($MinimumSizeMB)MB found in the last $Days days via MessageTraceV2." -ForegroundColor Gray
            }
        }

        #----------------------------------------------------------[7. Export]----------------------------------------------------------

        $folderExportPath    = $null
        $largeSentExportPath = $null

        if (-not $NoExport) {
            $folderExportPath = Join-Path $OutputPath "$($fileStem)_MailboxSpaceReport_$timestamp.csv"
            $folderStats | Export-Csv -Path $folderExportPath -NoTypeInformation
            Write-Host "`nFolder report exported to: $folderExportPath" -ForegroundColor Green

            if ($largeSent.Count -gt 0) {
                $largeSentExportPath = Join-Path $OutputPath "$($fileStem)_LargeSentMessages_$timestamp.csv"
                $largeSent | Export-Csv -Path $largeSentExportPath -NoTypeInformation
                Write-Host "Large sent messages exported to: $largeSentExportPath" -ForegroundColor Green
            }
        }

        Write-Host "`n===================================================" -ForegroundColor Cyan
        Write-Host " Report Complete" -ForegroundColor Cyan
        Write-Host "===================================================`n" -ForegroundColor Cyan

        if ($PassThru) {
            [pscustomobject]@{
                Mailbox             = $address
                Overview            = $overview
                Folders             = $folderStats
                RecoverableItems    = $recoverableItems
                Archive             = $archive
                SentItems           = $sentItems
                LargeSentMessages   = $largeSent
                FolderExportPath    = $folderExportPath
                LargeSentExportPath = $largeSentExportPath
            }
        }
    }
}

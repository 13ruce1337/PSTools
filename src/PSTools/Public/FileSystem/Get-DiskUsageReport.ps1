function Get-DiskUsageReport {
    <#
    .SYNOPSIS
        Shows the largest folders (recursively) and the largest files under a path.

    .DESCRIPTION
        Starting from the current directory (or -Path), scans all subfolders and reports:
          1. The total size and file count of the starting folder.
          2. The top N largest subfolders, with their total size
             (including all subfolders) and file count.
          3. The top N largest files with name, location and size.

        Output is written to the console. Use -PassThru to also return a report
        object.

        Items that cannot be read (access denied, path too long) are skipped and
        counted; sizes may be understated in that case. Run elevated for full
        results.

    .PARAMETER Path
        Folder(s) to scan. Defaults to the current directory. Accepts pipeline
        input (strings or objects with a FullName property, e.g. Get-ChildItem).

    .PARAMETER Top
        Number of largest folders and largest files to list. Default 10.

    .PARAMETER PassThru
        Return a report object in addition to the on-screen report.

    .EXAMPLE
        Get-DiskUsageReport

    .EXAMPLE
        Get-DiskUsageReport -Path D:\Shares -Top 25

    .EXAMPLE
        Get-ChildItem C:\Users -Directory | Get-DiskUsageReport -Top 5

    .EXAMPLE
        $report = Get-DiskUsageReport -Path C:\Data -PassThru
        $report.TopFolders | Select-Object -First 3
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [ValidateNotNullOrEmpty()]
        [string[]]$Path = '.',

        [ValidateRange(1, 2147483647)]
        [int]$Top = 10,

        [switch]$PassThru
    )

    begin {
        # Explicit byte multipliers (avoids relying on the KB/MB/GB/TB numeric literal suffixes).
        [int64]$bytesPerKB = 1024
        [int64]$bytesPerMB = 1048576
        [int64]$bytesPerGB = 1073741824
        [int64]$bytesPerTB = 1099511627776

        # Defined inline so it is not exported as part of the module.
        function Format-Size {
            param([int64]$Bytes)
            if ($Bytes -ge $bytesPerTB) { return ('{0:N2} TB' -f ($Bytes / $bytesPerTB)) }
            if ($Bytes -ge $bytesPerGB) { return ('{0:N2} GB' -f ($Bytes / $bytesPerGB)) }
            if ($Bytes -ge $bytesPerMB) { return ('{0:N2} MB' -f ($Bytes / $bytesPerMB)) }
            if ($Bytes -ge $bytesPerKB) { return ('{0:N2} KB' -f ($Bytes / $bytesPerKB)) }
            return "$Bytes B"
        }

        $folderColumns = @(
            @{ Label = 'Size';   Expression = { $_.Size };   Alignment = 'Right' }
            @{ Label = 'Files';  Expression = { $_.Files };  Alignment = 'Right' }
            @{ Label = 'Folder'; Expression = { $_.Folder } }
        )

        $fileColumns = @(
            @{ Label = 'Size';     Expression = { $_.Size }; Alignment = 'Right' }
            @{ Label = 'Name';     Expression = { $_.Name } }
            @{ Label = 'Location'; Expression = { $_.Location } }
        )
    }

    process {
        foreach ($inputPath in $Path) {

            #----------------------------------------------------------[Resolve path]----------------------------------------------------------

            try {
                $resolved = Resolve-Path -LiteralPath $inputPath -ErrorAction Stop | Select-Object -First 1
            }
            catch {
                Write-Error "Path not found: $inputPath"
                continue
            }

            if ($resolved.Provider.Name -ne 'FileSystem') {
                Write-Error "Path is not on the file system: $inputPath"
                continue
            }

            $root = $resolved.ProviderPath

            if (-not (Test-Path -LiteralPath $root -PathType Container)) {
                Write-Error "Path is not a folder: $root"
                continue
            }

            # Strip a trailing slash unless the path is a volume/share root (e.g. "C:\").
            if ([System.IO.Path]::GetPathRoot($root) -ne $root) {
                $root = $root.TrimEnd('\')
            }

            Write-Host "Scanning $root ..." -ForegroundColor Cyan

            #----------------------------------------------------------[Scan]----------------------------------------------------------

            $comparer    = [System.StringComparer]::OrdinalIgnoreCase
            $folderSize  = [System.Collections.Generic.Dictionary[string, long]]::new($comparer)
            $folderFiles = [System.Collections.Generic.Dictionary[string, int]]::new($comparer)
            $files       = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

            $folderSize[$root]  = 0
            $folderFiles[$root] = 0

            $scanErrors = $null
            $items = Get-ChildItem -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable scanErrors

            # Pass 1: record each file against its own folder only.
            foreach ($item in $items) {
                if ($item.PSIsContainer) {
                    if (-not $folderSize.ContainsKey($item.FullName)) {
                        $folderSize[$item.FullName]  = 0
                        $folderFiles[$item.FullName] = 0
                    }
                    continue
                }

                $files.Add($item)
                $dir = $item.DirectoryName

                if ($folderSize.ContainsKey($dir)) {
                    $folderSize[$dir]  += $item.Length
                    $folderFiles[$dir] += 1
                }
                else {
                    $folderSize[$dir]  = $item.Length
                    $folderFiles[$dir] = 1
                }
            }

            # Pass 2: roll totals up to parents. A child path is always longer than
            # its parent, so processing longest-first finalizes every child before
            # it is added to its parent.
            $folderKeys = @($folderSize.Keys | Sort-Object -Property Length -Descending)
            foreach ($key in $folderKeys) {
                if ($comparer.Equals($key, $root)) { continue }

                $parent = [System.IO.Path]::GetDirectoryName($key)
                if (-not $parent -or -not $folderSize.ContainsKey($parent)) { continue }

                $folderSize[$parent]  += $folderSize[$key]
                $folderFiles[$parent] += $folderFiles[$key]
            }

            #----------------------------------------------------------[Build report]----------------------------------------------------------

            $rootPrefixLength = if ($root.EndsWith('\')) { $root.Length } else { $root.Length + 1 }
            $subfolderCount   = $folderSize.Count - 1

            $topFolders = @($folderKeys |
                Where-Object { -not $comparer.Equals($_, $root) } |
                Sort-Object -Property { $folderSize[$_] } -Descending |
                Select-Object -First $Top |
                ForEach-Object {
                    [pscustomobject]@{
                        Folder   = '.\' + $_.Substring($rootPrefixLength)
                        Size     = Format-Size $folderSize[$_]
                        Files    = $folderFiles[$_]
                        Bytes    = $folderSize[$_]
                        FullName = $_
                    }
                })

            $topFiles = @($files |
                Sort-Object -Property Length -Descending |
                Select-Object -First $Top |
                ForEach-Object {
                    [pscustomobject]@{
                        Name     = $_.Name
                        Location = $_.DirectoryName
                        Size     = Format-Size $_.Length
                        Bytes    = $_.Length
                        FullName = $_.FullName
                    }
                })

            $errorCount = @($scanErrors).Count

            #----------------------------------------------------------[Display]----------------------------------------------------------

            Write-Host ''
            Write-Host "=== $root : total $(Format-Size $folderSize[$root]), $($folderFiles[$root]) files, $subfolderCount subfolders ===" -ForegroundColor Green
            Write-Host ''
            Write-Host "=== Top $Top largest folders ===" -ForegroundColor Green
            if ($topFolders.Count -gt 0) {
                $topFolders | Format-Table -AutoSize -Property $folderColumns | Out-Host
            }
            else {
                Write-Host 'No subfolders found.' -ForegroundColor Yellow
                Write-Host ''
            }

            Write-Host "=== Top $Top largest files ===" -ForegroundColor Green
            if ($topFiles.Count -gt 0) {
                $topFiles | Format-Table -AutoSize -Property $fileColumns | Out-Host
            }
            else {
                Write-Host 'No files found.' -ForegroundColor Yellow
            }

            if ($errorCount -gt 0) {
                Write-Warning "$errorCount item(s) could not be read (access denied or path too long). Sizes may be understated. Run as administrator for full results."
            }

            if ($PassThru) {
                [pscustomobject]@{
                    Path           = $root
                    TotalBytes     = $folderSize[$root]
                    TotalSize      = Format-Size $folderSize[$root]
                    FileCount      = $folderFiles[$root]
                    SubfolderCount = $subfolderCount
                    TopFolders     = $topFolders
                    TopFiles       = $topFiles
                    ErrorCount     = $errorCount
                }
            }
        }
    }
}

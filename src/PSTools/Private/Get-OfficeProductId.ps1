# Friendly product name -> Office Deployment Tool (ODT) Product ID.
# Single source of truth for SKUs. If you add or rename a key here, update the
# [ValidateSet()] on -Product in Install-, Uninstall- and Repair-OfficeProduct.
# Product IDs: https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/installation/product-ids-supported-office-deployment-click-to-run
$script:OfficeProductMap = [ordered]@{
    Office  = 'O365ProPlusRetail'
    Visio   = 'VisioPro2021Retail'
    Project = 'ProjectPro2021Retail'
    OneNote = 'OneNoteRetail'
    Excel   = 'Excel2019Retail'
    Outlook = 'OutlookRetail'
}

function Get-OfficeProductId {
    <#
    .SYNOPSIS
        Translates a friendly Office product name into its ODT Product ID.

    .EXAMPLE
        Get-OfficeProductId Visio
        # VisioPro2021Retail

    .EXAMPLE
        'Office', 'Project' | Get-OfficeProductId
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Name
    )

    process {
        foreach ($n in $Name) {
            if (-not $script:OfficeProductMap.Contains($n)) {
                throw "Unknown Office product '$n'. Valid values: $($script:OfficeProductMap.Keys -join ', ')."
            }
            $script:OfficeProductMap[$n]
        }
    }
}

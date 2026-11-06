function New-OdtConfiguration {
    <#
    .SYNOPSIS
        Builds an Office Deployment Tool configuration XML string.

    .DESCRIPTION
        Generates the XML consumed by 'setup.exe /configure'. Uses XmlWriter so
        every value is properly escaped.

          Add       - install one or more products
          Remove    - remove specific products (Remove All="FALSE")
          RemoveAll - remove every Click-to-Run Office product (Remove All="TRUE")

    .LINK
        https://learn.microsoft.com/en-us/microsoft-365-apps/deploy/office-deployment-tool-configuration-options
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Add', 'Remove', 'RemoveAll')]
        [string]$Action,

        [string[]]$ProductId,

        [ValidateNotNullOrEmpty()]
        [string]$Channel = 'Current',

        [ValidateSet('32', '64')]
        [string]$Architecture = '64',

        [ValidateNotNullOrEmpty()]
        [string[]]$Language = @('en-us'),

        [string[]]$ExcludeApp,

        # Folder ODT writes its logs to.
        [string]$LogPath,

        # Adds FORCEAPPSHUTDOWN=TRUE so ODT closes Office apps itself.
        [switch]$ForceAppShutdown
    )

    if ($Action -ne 'RemoveAll' -and -not $ProductId) {
        throw "-ProductId is required when -Action is '$Action'."
    }

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent             = $true
    $settings.OmitXmlDeclaration = $true

    $sb     = New-Object System.Text.StringBuilder
    $writer = [System.Xml.XmlWriter]::Create($sb, $settings)

    try {
        $writer.WriteStartElement('Configuration')

        switch ($Action) {
            'Add' {
                $writer.WriteStartElement('Add')
                $writer.WriteAttributeString('OfficeClientEdition', $Architecture)
                $writer.WriteAttributeString('Channel', $Channel)

                foreach ($id in $ProductId) {
                    $writer.WriteStartElement('Product')
                    $writer.WriteAttributeString('ID', $id)

                    foreach ($lang in $Language) {
                        $writer.WriteStartElement('Language')
                        $writer.WriteAttributeString('ID', $lang)
                        $writer.WriteEndElement()
                    }

                    foreach ($app in $ExcludeApp) {
                        $writer.WriteStartElement('ExcludeApp')
                        $writer.WriteAttributeString('ID', $app)
                        $writer.WriteEndElement()
                    }

                    $writer.WriteEndElement()   # Product
                }

                $writer.WriteEndElement()       # Add
            }

            'Remove' {
                $writer.WriteStartElement('Remove')
                $writer.WriteAttributeString('All', 'FALSE')

                foreach ($id in $ProductId) {
                    $writer.WriteStartElement('Product')
                    $writer.WriteAttributeString('ID', $id)
                    $writer.WriteEndElement()
                }

                $writer.WriteEndElement()       # Remove
            }

            'RemoveAll' {
                $writer.WriteStartElement('Remove')
                $writer.WriteAttributeString('All', 'TRUE')
                $writer.WriteEndElement()
            }
        }

        if ($ForceAppShutdown) {
            $writer.WriteStartElement('Property')
            $writer.WriteAttributeString('Name', 'FORCEAPPSHUTDOWN')
            $writer.WriteAttributeString('Value', 'TRUE')
            $writer.WriteEndElement()
        }

        $writer.WriteStartElement('Display')
        $writer.WriteAttributeString('Level', 'None')
        $writer.WriteAttributeString('AcceptEULA', 'TRUE')
        $writer.WriteEndElement()

        if ($LogPath) {
            $writer.WriteStartElement('Logging')
            $writer.WriteAttributeString('Level', 'Standard')
            $writer.WriteAttributeString('Path', $LogPath)
            $writer.WriteEndElement()
        }

        $writer.WriteEndElement()               # Configuration
    }
    finally {
        $writer.Flush()
        $writer.Close()
    }

    $sb.ToString()
}

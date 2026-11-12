@{
    RootModule            = 'PSTools.psm1'

    ModuleVersion         = '2.4.1'

    GUID                  = 'b7c1f6a2-3d54-4e8f-9a1b-6c2e0d7f4a93'

    Author                = 'Trevor Cooper'
    CompanyName           = 'Community'
    Copyright             = '(c) Trevor Cooper. All rights reserved.'
    Description           = 'PowerShell IT toolkit: Hyper-V provisioning, Microsoft 365 administration, user deprovisioning, disk usage reporting and more.'

    PowerShellVersion     = '5.1'
    CompatiblePSEditions  = @('Desktop', 'Core')

    RequiredModules       = @()

    # Export explicitly (faster command discovery than wildcards).
    # Keep this list in sync with the files in Public\ (including subfolders).
    FunctionsToExport     = @(
        # Users
        'Disable-User'

        # FileSystem
        'Get-DiskUsageReport'

        # HyperV
        'New-DebianVM'
        'New-WindowsVM'

        # M365
        'Disable-M365User'
        'Get-M365AdminMfaReport'
        'Get-M365MailboxSpaceReport'
        'Get-M365Room'
        'Get-M365RoomList'

        # Network
        'Get-MerakiFirmware'
    )
    CmdletsToExport       = @()
    VariablesToExport     = @()
    AliasesToExport       = @()

    PrivateData           = @{
        PSData = @{
            Tags       = @('IT', 'Admin', 'HyperV', 'Debian', 'Automation', 'M365', 'ExchangeOnline', 'EntraID', 'MFA', 'DiskUsage', 'Storage', 'Meraki', 'Networking')
            ProjectUri = 'https://github.com/13ruce1337/PSTools'
        }
    }
}

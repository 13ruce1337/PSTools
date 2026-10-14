@{
    # Script module or binary module file associated with this manifest.
    RootModule            = 'PSTools.psm1'

    # Version number of this module (single source of truth for versioning).
    ModuleVersion         = '2.0.0'

    # Unique ID for this module.
    GUID                  = 'b7c1f6a2-3d54-4e8f-9a1b-6c2e0d7f4a93'

    Author                = 'Trevor Cooper'
    CompanyName           = 'Community'
    Copyright             = '(c) Trevor Cooper. All rights reserved.'
    Description           = 'PowerShell IT toolkit: Hyper-V provisioning, driver staging, user deprovisioning and more.'

    # Windows PowerShell 5.1 and PowerShell 7+ (Hyper-V cmdlets are Windows-only).
    PowerShellVersion     = '5.1'
    CompatiblePSEditions  = @('Desktop', 'Core')

    # Heavy or optional dependencies (Hyper-V, Graph, ExchangeOnline, ...) are
    # checked inside the functions that need them, so importing PSTools stays
    # lightweight on machines that don't have them.
    RequiredModules       = @()

    # Export explicitly (faster command discovery than wildcards).
    # Keep this list in sync with the files in Public/.
    FunctionsToExport     = @(
        'New-DebianVM'
    )
    CmdletsToExport       = @()
    VariablesToExport     = @()
    AliasesToExport       = @()

    PrivateData           = @{
        PSData = @{
            Tags       = @('IT', 'Admin', 'HyperV', 'Debian', 'Automation')
            ProjectUri = 'https://github.com/13ruce1337/PSTools'
            # LicenseUri = 'https://github.com/13ruce1337/PSTools/blob/main/LICENSE'
        }
    }
}

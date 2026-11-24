# PSTools — PowerShell IT Toolkit

PSTools is a PowerShell module of IT admin commands for on-premises Active Directory, Microsoft 365, Intune, Hyper-V, file systems and networking. Each command is a self-contained advanced function with built-in help.

## Installation

Clone the repo and import the module from the manifest:

```powershell
git clone https://github.com/13ruce1337/PSTools.git
Import-Module .\PSTools\src\PSTools\PSTools.psd1
Get-Command -Module PSTools
```

> **Note:** If the import fails with *"running scripts is disabled on this system"*, change the execution policy for the current terminal only:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

---

## Areas

Each area links to its command reference in [`docs/`](docs/README.md).

- **[Active Directory](docs/ActiveDirectory/)**: User deprovisioning for on-premises AD. Requires RSAT (the `ActiveDirectory` module).
- **[File System](docs/FileSystem/)**: Disk usage reports. No extra requirements.
- **[Hyper-V](docs/HyperV/)**: Debian and Windows VM provisioning. Run from an elevated session with the Hyper-V module installed.
- **[Intune](docs/Intune/)**: Device cleanup across Intune, Windows Autopilot and Entra ID. Uses Microsoft Graph.
- **[Microsoft 365](docs/M365/)**: User lookup, deprovisioning, groups and rooms, plus license, password, MFA and mailbox reports. Uses Microsoft Graph and Exchange Online.
- **[Network](docs/Network/)**: Meraki firmware reports. Requires a Meraki Dashboard API key.

---

## Requirements

| Requirement | Details |
|---|---|
| **OS** | Windows 10 / 11 |
| **PowerShell** | Windows PowerShell 5.1 or PowerShell 7+ |
| **Execution policy** | Scripts must be allowed to run. See the note under [Installation](#installation) |

Missing modules are detected at run time, and you'll be asked whether to install them from the PowerShell Gallery. Commands that use Microsoft Graph list the permissions they request in their help.

---

## Getting help

Every command includes built-in help with parameters and examples:

```powershell
# List all commands in the module
Get-Command -Module PSTools

# Show usage examples for a command
Get-Help Remove-IntuneDevice -Examples

# Show full help, including every parameter and the permissions it needs
Get-Help Remove-IntuneDevice -Full
```

The same help is published as Markdown in [`docs/`](docs/README.md). It is generated from the code, so it always matches the current version.

Commands that make changes support `-WhatIf`, so you can preview what they'll do before running them for real.

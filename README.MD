# PSTools — PowerShell IT Toolkit

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

## Commands

Each command is a self-contained advanced function with comment-based help. The descriptions below are summaries. See [Getting help](#getting-help) for parameters and examples.

**Areas:** [Active Directory](#active-directory) · [File System](#file-system) · [Hyper-V](#hyper-v) · [Intune](#intune) · [Microsoft 365](#microsoft-365) · [Network](#network)

### Active Directory

Requires RSAT (the `ActiveDirectory` module).

| Command | Description |
|---|---|
| `Disable-ADUser` | Disables an on-prem AD account, stamps the description and removes group memberships. Can also reset the password, hide the account from address lists and move it to another OU. |

### File System

No extra requirements.

| Command | Description |
|---|---|
| `Get-DiskUsageReport` | Shows the largest folders (recursive totals) and largest files under a path. |

### Hyper-V

Run from an elevated session with the Hyper-V PowerShell module installed. `New-DebianVM` also uses QEMU and the Windows ADK, and installs them via winget if they're missing.

| Command | Description |
|---|---|
| `New-DebianVM` | Creates a Gen 2 Debian VM from the Debian cloud image with a cloud-init seed (hostname, root password, packages). |
| `New-WindowsVM` | Creates a Gen 2 Windows 11 reference VM (vTPM, Secure Boot, no checkpoints) on an isolated switch and boots it from an ISO. |

### Intune

Uses Microsoft Graph. Your account needs a role that can delete devices, such as Intune Administrator. See [Permissions](#permissions) for details.

| Command | Description |
|---|---|
| `Remove-IntuneDevice` | Removes devices from Intune by serial number, with an optional hostname check. Can also remove them from Windows Autopilot and Entra ID in the recommended order (Intune, Autopilot, Entra ID). Accepts CSV or pipeline input and supports `-WhatIf`. |

### Microsoft 365

Uses Microsoft Graph and Exchange Online. Missing modules are offered for install the first time you run a command, and you'll be asked to sign in. See [Permissions](#permissions) for details.

#### Users

| Command | Description |
|---|---|
| `Get-M365User` | Admin-center style overview of a user: account, password, sign-ins, MFA, licenses, groups/roles, mailbox delegates, forwarding, out-of-office and mobile devices. |
| `Disable-M365User` | Deprovisions a user: sets out-of-office, converts the mailbox to shared, grants the manager access, removes group/DL memberships and mobile devices, and removes licenses. |

#### Groups & rooms

| Command | Description |
|---|---|
| `Get-M365GroupMember` | Lists the user members of a group by display name or email. Supports nested membership and CSV export. |
| `Get-M365RoomList` | Lists room lists (locations/buildings). Pipe the output to `Get-M365Room`. |
| `Get-M365Room` | Lists room mailboxes, optionally filtered by room list. |

#### Reports

| Command | Description |
|---|---|
| `Get-M365AdminMfaReport` | MFA registration strength for every Entra ID directory role holder, including group-assigned and PIM-eligible roles. |
| `Get-M365License` | Licensed users who haven't signed in recently (or ever), to find assigned but unused licenses. |
| `Get-M365PasswordLastSet` | Users whose password was set longest ago, with password age and account status. |
| `Get-M365MailboxSpaceReport` | Breaks down mailbox space: folders, Recoverable Items, archive and large sent messages. |

### Network

Requires a Meraki Dashboard API key.

| Command | Description |
|---|---|
| `Get-MerakiFirmware` | Exports Meraki device firmware versions and the latest available firmware per product type to CSV. |

---

## Requirements

The table below applies to every command. Each area under [Commands](#commands) also has a short note on what it needs, and [Permissions](#permissions) lists the Graph permissions each command requests.

| Requirement | Details |
|---|---|
| **OS** | Windows 10 / 11 |
| **PowerShell** | Windows PowerShell 5.1 or PowerShell 7+ |
| **Execution policy** | Scripts must be allowed to run. See the note under [Installation](#installation) |

Missing modules are detected at run time, and you'll be asked whether to install them from the PowerShell Gallery.

### Permissions

Graph commands sign in with `Microsoft.Graph.Authentication` and request the permissions below. You'll see them on the consent prompt the first time you sign in. Most only request read access.

| Command | Graph permissions | Notes |
|---|---|---|
| `Remove-IntuneDevice` | `DeviceManagementManagedDevices.ReadWrite.All`, plus `DeviceManagementServiceConfig.ReadWrite.All` for Autopilot and `Device.ReadWrite.All` for Entra ID | Needs a role that can delete these records, such as Intune Administrator |
| `Get-M365User` | `User.Read.All`, `GroupMember.Read.All`, `UserAuthenticationMethod.Read.All`, `AuditLog.Read.All` | Sign-in activity needs Entra ID P1 |
| `Get-M365License` | `User.Read.All`, `AuditLog.Read.All`, `Organization.Read.All` | Sign-in activity needs Entra ID P1 |
| `Get-M365PasswordLastSet` | `User.Read.All` | |

`Disable-M365User` also uses the `Microsoft.Graph.Users`, `Microsoft.Graph.Users.Actions` and `Microsoft.Graph.Groups` modules. Exchange Online commands use `ExchangeOnlineManagement`. The large-sent-messages section of `Get-M365MailboxSpaceReport` needs version 3.7.0 or later.

---

## Getting help

Every command includes built-in help with parameters and examples:

```powershell
# List all commands in the module
Get-Command -Module PSTools

# Show usage examples for a command
Get-Help Remove-IntuneDevice -Examples

# Show full help, including every parameter
Get-Help Remove-IntuneDevice -Full
```

Commands that make changes support `-WhatIf`, so you can preview what they'll do before running them for real.

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

Each command is a self-contained advanced function with comment-based help. The descriptions below are summaries. Run `Get-Help <Command> -Full` for parameters, details and examples.

**Areas:** [Active Directory](#active-directory) · [File System](#file-system) · [Hyper-V](#hyper-v) · [Intune](#intune) · [Microsoft 365](#microsoft-365) · [Network](#network)

### Active Directory

> **Requires:** RSAT (the `ActiveDirectory` module).

| Command | Description |
|---|---|
| `Disable-ADUser` | Disables an on-prem AD account, stamps the description and removes group memberships. Can also reset the password, hide the account from address lists and move it to another OU. |

### File System

> **Requires:** nothing beyond the general [requirements](#requirements).

| Command | Description |
|---|---|
| `Get-DiskUsageReport` | Shows the largest folders (recursive totals) and largest files under a path. |

### Hyper-V

> **Requires:** an elevated session and the Hyper-V PowerShell module. `New-DebianVM` installs QEMU (`qemu-img`) and the Windows ADK Deployment Tools (`oscdimg`) via winget if they are missing.

| Command | Description |
|---|---|
| `New-DebianVM` | Creates a Gen 2 Debian VM from the Debian cloud image with a cloud-init seed (hostname, root password, packages). |
| `New-WindowsVM` | Creates a Gen 2 Windows 11 reference VM (vTPM, Secure Boot, no checkpoints) on an isolated switch and boots it from an ISO. |

### Intune

> **Requires:** `Microsoft.Graph.Authentication`. Requests `DeviceManagementManagedDevices.ReadWrite.All`, plus `DeviceManagementServiceConfig.ReadWrite.All` for Autopilot and `Device.ReadWrite.All` for Entra ID. The signed-in account needs a role that can delete these records (for example Intune Administrator).

| Command | Description |
|---|---|
| `Remove-IntuneDevice` | Removes devices from Intune by serial number, with an optional hostname check. Can also remove them from Windows Autopilot and Entra ID in the recommended order (Intune, Autopilot, Entra ID). Accepts CSV or pipeline input and supports `-WhatIf`. |

### Microsoft 365

> **Requires** (applies to all Microsoft 365 commands below):
>
> - **Graph commands:** `Microsoft.Graph.Authentication`.
>   - `Disable-M365User` also needs `Microsoft.Graph.Users`, `Microsoft.Graph.Users.Actions` and `Microsoft.Graph.Groups`.
>   - `Get-M365License` requests `User.Read.All`, `AuditLog.Read.All` and `Organization.Read.All`. Sign-in activity also needs Entra ID P1.
>   - `Get-M365PasswordLastSet` requests `User.Read.All`.
>   - `Get-M365User` requests `User.Read.All`, `GroupMember.Read.All`, `UserAuthenticationMethod.Read.All` and `AuditLog.Read.All`. Sign-in activity also needs Entra ID P1.
> - **Exchange commands:** `ExchangeOnlineManagement`. The large-sent-messages section of `Get-M365MailboxSpaceReport` needs version 3.7.0+ (`Get-MessageTraceV2`).

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

> **Requires:** a Meraki Dashboard API key.

| Command | Description |
|---|---|
| `Get-MerakiFirmware` | Exports Meraki device firmware versions and the latest available firmware per product type to CSV. |

---

## Requirements

These apply to every command. Area-specific requirements are listed at the top of each area under [Commands](#commands).

| Requirement | Details |
|---|---|
| **OS** | Windows 10 / 11 |
| **PowerShell** | Windows PowerShell 5.1 or PowerShell 7+ |
| **Execution policy** | Scripts must be allowed to run. See the note under [Installation](#installation) |

Missing modules are detected at run time, and you'll be asked whether to install them from the PowerShell Gallery.

---

## Usage examples

**Active Directory & Microsoft 365 leavers**

```powershell
# Hybrid leaver: on-prem AD first, then Microsoft 365
Disable-ADUser -Identity jdoe -ResetPassword -HideFromAddressList -TargetOU 'OU=Disabled Users,DC=contoso,DC=com'
Disable-M365User -UserPrincipalName jdoe@contoso.com -ManagerUserPrincipalName manager@contoso.com

# Bulk AD deprovisioning from CSV (preview first)
Import-Csv .\leavers.csv | Disable-ADUser -WhatIf
```

**File System**

```powershell
# Disk usage
Get-DiskUsageReport -Path D:\Shares -Top 25
```

**Hyper-V**

```powershell
# Debian VM on Hyper-V
$pw = Read-Host -AsSecureString -Prompt 'Root password'
New-DebianVM -Name web01 -CpuCount 2 -MemoryGB 4 -DiskGB 40 -RootPassword $pw
```

**Intune**

```powershell
# Preview removing a device from Intune, Autopilot and Entra ID
Remove-IntuneDevice -SerialNumber 5CG1234ABC -Hostname LAPTOP-042 -IncludeEntra -WhatIf

# Retire a batch of devices from a CSV (SerialNumber, Hostname columns) without prompting, saving the results
Remove-IntuneDevice -CsvPath .\retired.csv -IncludeEntra -Force | Export-Csv .\retired-results.csv -NoTypeInformation
```

**Microsoft 365**

```powershell
# Quick overview of a user (account, password, licenses, mailbox access, out-of-office...)
Get-M365UserSummary -Identity user@contoso.com

# The 30 users with the oldest passwords
Get-M365PasswordLastSet

# Password last set for one user
Get-M365PasswordLastSet -Identity user@contoso.com

# Top 50 enabled accounts by password age, with the full list exported to C:\Reports
Get-M365PasswordLastSet -Top 50 -EnabledOnly -OutputPath C:\Reports

# Active accounts whose password is over a year old
Get-M365PasswordLastSet -PassThru | Where-Object { $_.IsActive -and $_.PasswordAgeDays -gt 365 } | Format-Table

# Licensed users with no sign-in in the last 30 days
Get-M365LicenseActivity

# Same with a 90-day threshold, including disabled accounts, exported to C:\Reports
Get-M365LicenseActivity -DaysInactive 90 -IncludeDisabled -OutputPath C:\Reports

# Group membership, including nested groups, exported to CSV
Get-M365GroupMember -GroupName 'All Staff' -Transitive -Export

# Admin MFA posture: show only admins without strong MFA
Get-M365AdminMfaReport -NoExport | Where-Object MfaStatus -notin 'Enabled', 'Phishing-resistant'

# Mailbox space investigation
Get-M365MailboxSpaceReport -Identity user@contoso.com -Days 30 -MinimumSizeMB 10

# Rooms in a specific location
Get-M365RoomList -Name '*Dallas*' | Get-M365Room
```

# PSTools — PowerShell IT Toolkit

## Installation

Clone the repo and import the module from the manifest:

```powershell
git clone https://github.com/13ruce1337/PSTools.git
Import-Module .\PSTools\src\PSTools\PSTools.psd1
Get-Command -Module PSTools
```

To make it available in every session, copy `src\PSTools` into a folder on your `$env:PSModulePath`, for example `$HOME\Documents\PowerShell\Modules\PSTools`. Then run `Import-Module PSTools`.

---

## Commands

Each command is a self-contained advanced function with comment-based help. Run `Get-Help <Command> -Full` for details and examples.

| Area | Command | Description |
|---|---|---|
| **ActiveDirectory** | `Disable-ADUser` | Disables an on-prem AD account, stamps the description, removes group memberships, and optionally resets the password, hides it from address lists, and moves it to another OU. |
| **FileSystem** | `Get-DiskUsageReport` | Shows the largest folders (recursive totals) and largest files under a path. |
| **HyperV** | `New-DebianVM` | Creates a Gen 2 Debian VM from the Debian cloud image with a cloud-init seed (hostname, root password, packages). |
| **HyperV** | `New-WindowsVM` | Creates a Gen 2 Windows 11 reference VM (vTPM, Secure Boot, no checkpoints) on an isolated switch and boots it from an ISO. |
| **Intune** | `Remove-IntuneDevice` | Removes devices from Intune by serial number (with an optional hostname check), and optionally from Windows Autopilot and Entra ID in the correct order (Intune, Autopilot, Entra). Accepts a CSV for batches and supports `-WhatIf`. |
| **M365** | `Disable-M365User` | Deprovisions a user in Microsoft 365: sets out-of-office, converts to shared, grants the manager access, removes group/DL memberships, removes mobile devices, and removes licenses. |
| **M365** | `Get-M365AdminMfaReport` | Reports MFA registration strength for every Entra ID directory role holder, including group-assigned and PIM-eligible roles. |
| **M365** | `Get-M365GroupMember` | Lists the user members of a group by display name or email, with optional nested membership and CSV export. |
| **M365** | `Get-M365License` | Shows licensed users who haven't signed in recently (or ever), with their licenses and last sign-in, to find assigned but unused licenses. Optional CSV export of all licensed users. |
| **M365** | `Get-M365MailboxSpaceReport` | Breaks down mailbox space usage: folders, Recoverable Items, archive, and large sent messages. |
| **M365** | `Get-M365PasswordLastSet` | Shows the users whose password was set longest ago (top 30 by default, with password age and account status) for all users or specific UPNs/emails, with optional CSV export of the full list. |
| **M365** | `Get-M365Room` | Lists room mailboxes, optionally filtered by room list. |
| **M365** | `Get-M365RoomList` | Lists room lists (locations/buildings). Pipe the output to `Get-M365Room`. |
| **M365** | `Get-M365User` | Admin-center style overview of a user: account status, password last set, sign-ins, MFA methods, licenses, groups/roles, mailbox delegates, forwarding, out-of-office and mobile devices. |
| **Network** | `Get-MerakiFirmware` | Exports Meraki device firmware versions and the latest available firmware per product type to CSV. |

---

## Requirements

**General**

| Requirement | Details |
|---|---|
| **OS** | Windows 10 / 11 |
| **PowerShell** | Windows PowerShell 5.1 or PowerShell 7+ |
| **Execution policy** | If imports are blocked, run `Set-ExecutionPolicy -Scope Process Bypass` in the current session |

**Per area**

| Area | Needs |
|---|---|
| **ActiveDirectory** | RSAT (the `ActiveDirectory` module) |
| **HyperV** | Elevated session and the Hyper-V PowerShell module. `New-DebianVM` installs QEMU (`qemu-img`) and the Windows ADK Deployment Tools (`oscdimg`) via winget if they are missing. |
| **Intune (Graph)** | `Microsoft.Graph.Authentication`. `Remove-IntuneDevice` requests `DeviceManagementManagedDevices.ReadWrite.All`, plus `DeviceManagementServiceConfig.ReadWrite.All` for Autopilot and `Device.ReadWrite.All` for Entra ID. The signed-in account needs a role that can delete these records (for example Intune Administrator). |
| **M365 (Graph)** | `Microsoft.Graph.Authentication`. `Disable-M365User` also needs `Microsoft.Graph.Users`, `Microsoft.Graph.Users.Actions` and `Microsoft.Graph.Groups`. `Get-M365LicenseActivity` requests `User.Read.All`, `AuditLog.Read.All` and `Organization.Read.All` (sign-in activity also needs Entra ID P1). `Get-M365PasswordLastSet` requests `User.Read.All`. `Get-M365UserSummary` requests `User.Read.All`, `GroupMember.Read.All`, `UserAuthenticationMethod.Read.All` and `AuditLog.Read.All` (sign-in activity also needs Entra ID P1). |
| **M365 (Exchange)** | `ExchangeOnlineManagement`. The large-sent-messages section of `Get-M365MailboxSpaceReport` needs version 3.7.0+ (`Get-MessageTraceV2`). |
| **Network** | A Meraki Dashboard API key |


---

## Usage examples

```powershell
# Disk usage
Get-DiskUsageReport -Path D:\Shares -Top 25

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

# Hybrid leaver: on-prem AD first, then Microsoft 365
Disable-ADUser -Identity jdoe -ResetPassword -HideFromAddressList -TargetOU 'OU=Disabled Users,DC=contoso,DC=com'
Disable-M365User -UserPrincipalName jdoe@contoso.com -ManagerUserPrincipalName manager@contoso.com

# Bulk AD deprovisioning from CSV (preview first)
Import-Csv .\leavers.csv | Disable-ADUser -WhatIf

# Preview removing a device from Intune, Autopilot and Entra ID
Remove-IntuneDevice -SerialNumber 5CG1234ABC -Hostname LAPTOP-042 -IncludeEntra -WhatIf

# Retire a batch of devices from a CSV (SerialNumber, Hostname columns) without prompting, saving the results
Remove-IntuneDevice -CsvPath .\retired.csv -IncludeEntra -Force | Export-Csv .\retired-results.csv -NoTypeInformation

# Debian VM on Hyper-V
$pw = Read-Host -AsSecureString -Prompt 'Root password'
New-DebianVM -Name web01 -CpuCount 2 -MemoryGB 4 -DiskGB 40 -RootPassword $pw
```

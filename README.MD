# PSTools — PowerShell IT Toolkit

- Microsoft 365 / Exchange Online administration and reporting
- User deprovisioning in hybrid Active Directory / Entra ID environments
- Hyper-V VM provisioning
- Disk usage reporting
- Meraki firmware inventory

Each command is a self-contained advanced function with comment-based help. Run `Get-Help <Command> -Full` for details and examples.

---

## Commands

| Area | Command | Description |
|---|---|---|
| **ActiveDirectory** | `Disable-ADUser` | Disables an on-prem AD account, stamps the description, removes group memberships, and optionally resets the password, hides it from address lists, and moves it to another OU. |
| **FileSystem** | `Get-DiskUsageReport` | Shows the largest folders (recursive totals) and largest files under a path. |
| **HyperV** | `New-DebianVM` | Creates a Gen 2 Debian VM from the Debian cloud image with a cloud-init seed (hostname, root password, packages). |
| **HyperV** | `New-WindowsVM` | Creates a Gen 2 Windows 11 reference VM (vTPM, Secure Boot, no checkpoints) on an isolated switch and boots it from an ISO. |
| **M365** | `Disable-M365User` | Deprovisions a user in Microsoft 365: sets out-of-office, converts to shared, grants the manager access, removes group/DL memberships, removes mobile devices, and removes licenses. |
| **M365** | `Get-M365AdminMfaReport` | Reports MFA registration strength for every Entra ID directory role holder, including group-assigned and PIM-eligible roles. |
| **M365** | `Get-M365GroupMember` | Lists the user members of a group by display name or email, with optional nested membership and CSV export. |
| **M365** | `Get-M365MailboxSpaceReport` | Breaks down mailbox space usage: folders, Recoverable Items, archive, and large sent messages. |
| **M365** | `Get-M365Room` | Lists room mailboxes, optionally filtered by room list. |
| **M365** | `Get-M365RoomList` | Lists room lists (locations/buildings). Pipe the output to `Get-M365Room`. |
| **M365** | `Get-M365UserSummary` | Admin-center style overview of a user: account status, password last set, sign-ins, MFA methods, licenses, groups/roles, mailbox delegates, forwarding, out-of-office and mobile devices. |
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
| **M365 (Graph)** | `Microsoft.Graph.Authentication`. `Disable-M365User` also needs `Microsoft.Graph.Users`, `Microsoft.Graph.Users.Actions` and `Microsoft.Graph.Groups`. `Get-M365UserSummary` requests `User.Read.All`, `GroupMember.Read.All`, `UserAuthenticationMethod.Read.All` and `AuditLog.Read.All` (sign-in activity also needs Entra ID P1). |
| **M365 (Exchange)** | `ExchangeOnlineManagement`. The large-sent-messages section of `Get-M365MailboxSpaceReport` needs version 3.7.0+ (`Get-MessageTraceV2`). |
| **Network** | A Meraki Dashboard API key |

If a required Microsoft Graph or Exchange Online module is missing, you are asked whether to install it from the PowerShell Gallery (CurrentUser scope, no elevation needed). If you decline, the command stops and shows the `Install-Module` command to run.

Existing Microsoft Graph and Exchange Online sessions are reused when they have the required scopes. Otherwise you are prompted to sign in. The signed-in account and tenant are always shown before any changes are made.

---

## Installation

Clone the repo and import the module from the manifest:

```powershell
git clone https://github.com/13ruce1337/PSTools.git
Import-Module .\PSTools\src\PSTools\PSTools.psd1

Get-Command -Module PSTools
```

To make it available in every session, copy `src\PSTools` into a folder on your `$env:PSModulePath`, for example `$HOME\Documents\PowerShell\Modules\PSTools`. Then run `Import-Module PSTools`.

---

## Usage examples

```powershell
# Disk usage
Get-DiskUsageReport -Path D:\Shares -Top 25

# Quick overview of a user (account, password, licenses, mailbox access, out-of-office...)
Get-M365UserSummary -Identity user@contoso.com

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

# Debian VM on Hyper-V
$pw = Read-Host -AsSecureString -Prompt 'Root password'
New-DebianVM -Name web01 -CpuCount 2 -MemoryGB 4 -DiskGB 40 -RootPassword $pw
```

### Notes

- **Output location:** report commands write CSVs to `$env:TEMP\PSTools` by default. Use `-OutputPath` to change the folder, or `-NoExport` to skip the CSV.
- **`New-WindowsVM`** has no parameters yet. The VM name, ISO path (`C:\ISOs\Win11_Unattended.iso` by default), disk size, CPU count and switch names are set at the top of the function. The VM starts on an isolated private switch, and the summary at the end prints the command to connect it to a real network.
- **`Get-MerakiFirmware`** writes `Meraki_Firmware_Report.csv` to the current directory and a transcript to `C:\Windows\Temp\get_meraki_firmware.log`. It uses the deprecated `/networks/{id}/devices` endpoint; see [`CHANGELOG.md`](CHANGELOG.md).
- **Preview mode:** commands that make changes support `-WhatIf` and `-Confirm` where noted in their help.

---
Adding a command

1. Create a file named after the function (for example `Get-Thing.ps1`) in the appropriate `Public\<Area>\` folder. The file must contain only the function definition, because any top-level code runs at import time.
2. Add the function name to `FunctionsToExport` in `PSTools.psd1`.
3. Add an entry under `[Unreleased]` in `CHANGELOG.md`.

Cutting a release:

```powershell
Update-ModuleManifest -Path .\src\PSTools\PSTools.psd1 -ModuleVersion 0.7.0
# Move the [Unreleased] entries in CHANGELOG.md under a new [0.7.0] heading, then:
git commit -am "Release 0.7.0"
git tag v0.7.0
git push --follow-tags
```

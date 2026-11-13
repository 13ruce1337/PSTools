# PSTools — PowerShell IT Toolkit

PSTools is a collection of standalone PowerShell scripts for everyday IT work. The library covers the application lifecycle (install, uninstall, reinstall, and troubleshoot), user deprovisioning in hybrid AD/Entra environments, provisioning Windows and Debian Hyper-V reference VMs, staging Lenovo drivers and pulling Meraki firmware inventory, and building unattended Windows installation media. Each script runs independently from the command line, so you can use them individually or wire them into your own automation.

---

## Versioning

PSTools follows [Semantic Versioning](https://semver.org) (`MAJOR.MINOR.PATCH`).

- The module version is defined in **one place only**: `ModuleVersion` in [`src/PSTools/PSTools.psd1`](src/PSTools/PSTools.psd1). Individual functions are not versioned separately.
- All notable changes are recorded in [`CHANGELOG.md`](CHANGELOG.md).
- **Pre-1.0 (`0.x.y`)**: the module is still being consolidated. `MINOR` bumps add commands and *may* include breaking changes; `PATCH` bumps are fixes only.
- **1.0.0** will mark a stable public command set with documentation, published as a GitHub release (and potentially to the PowerShell Gallery). After that, breaking changes require a `MAJOR` bump.

Check the version you have:

```powershell
# From a clone of the repo
(Import-PowerShellDataFile .\src\PSTools\PSTools.psd1).ModuleVersion

# From an imported / installed module
Get-Module PSTools -ListAvailable | Select-Object Name, Version, Path
```

Cutting a release:

```powershell
Update-ModuleManifest -Path .\src\PSTools\PSTools.psd1 -ModuleVersion 0.7.0
# Move the [Unreleased] entries in CHANGELOG.md under a new [0.7.0] heading, then:
git commit -am "Release 0.7.0"
git tag v0.7.0
git push --follow-tags
```

---

## Requirements

| Requirement | Details |
|---|---|
| **Permissions** | Must be run as **Administrator** |
| **Shell** | PowerShell (Execution Policy Bypass required) |
| **OS** | Windows 10 / 11 |

---

# Windows Unattended Reference Image

This guide details how to build an automated, bootable Windows 10/11 installation ISO using an answer file (`autounattend.xml`) and Microsoft's `oscdimg` deployment utility.


### Prerequisites
1. **Windows Assessment and Deployment Kit (ADK)**:
   Install the Windows ADK (which includes `oscdimg.exe` under Deployment Tools) via winget:
   ```powershell
   winget install --id Microsoft.WindowsADK --exact --accept-package-agreements --accept-source-agreements
   ```
2. **Source Media & Answer File**:
   A directory containing the extracted contents of the Windows installation ISO with your configured `autounattend.xml` placed in the root directory.

## 1. Creating the Bootable Unattended ISO
Use Microsoft's `oscdimg` utility (included with the Windows ADK Deployment Tools) to package the extracted Windows install files and `autounattend.xml` into a UEFI/BIOS dual-boot ISO with full UDF long-filename support.

### Command

Run the following command in **PowerShell (Run as Administrator)**:

```powershell
& "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe" -m -o -u2 -udfver102 "-bootdata:2#p0,e,b<SOURCE_DIR>\boot\etfsboot.com#pEF,e,b<SOURCE_DIR>\efi\microsoft\boot\efisys_noprompt.bin" "<SOURCE_DIR>" "<OUTPUT_DIR>\<ISO_NAME>"
```

### Parameter Placeholders

| Placeholder | Description | Example |
| :--- | :--- | :--- |
| `<SOURCE_DIR>` | Root directory containing the extracted ISO files and `autounattend.xml` | `D:\ISO_Build\Win10_Source` |
| `<OUTPUT_DIR>` | Target folder where the generated `.iso` will be saved | `D:\ISOs` |
| `<ISO_NAME>` | Desired filename for the output ISO | `Win10Pro_Unattended` |

### Flag Breakdown

* **`-m`**: Disables the standard maximum CD/DVD size barrier (required for multi-gigabyte modern WIM/ESD images).
* **`-o`**: Enables binary hashing optimization to encode duplicate files only once.
* **`-u2 -udfs`**: Forces a pure UDF filesystem (v1.02) without legacy ISO 9660 restrictions. This ensures long file names are preserved and prevents `autounattend.xml` from truncating to 8.3 format (such as `AUTOUN~1.XML`).
* **`-bootdata:2...`**: Configures multi-platform boot headers:
  * `p0,e,b"...etfsboot.com"`: Legacy BIOS boot sector (Platform `0`).
  * `pEF,e,b"...efisys.bin"`: UEFI boot sector (Platform `EF`).

## 2. Provisioning the Windows VM
* Run the provision_windows.ps1 script after filling in the variables.
* This same script can be ran on the downloaded iso from windows directly if needed.

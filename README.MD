# Alfred Pennyworth
**Must be ran as administrator**
```
Invoke-WebRequest https://raw.githubusercontent.com/13ruce1337/PSTools/main/alfred.ps1 -OutFile "$env:TEMP\alfred.ps1"
```
```
start-process powershell.exe -Argumentlist "-ExecutionPolicy Bypass", "-File $env:TEMP\alfred.ps1", "update windows" -NoNewWindow
```

## COMMANDS
- install
- uninstall
- reinstall
- nuke
- update
- troubleshoot
- sd
## APPLICATIONS
- office
- teams
- project
- visio
- onenote
- onedrive
- adobedc
- network

---

# Windows Unattended Reference Image & FFU Preparation Guide

This guide details how to build an automated, bootable Windows 10/11 installation ISO using an answer file (`autounattend.xml`) and Microsoft's `oscdimg` deployment utility. The resulting VM boots straight to the desktop for rapid configuration prior to Sysprep and Full Flash Update (FFU) capture.

## 1. Creating the Bootable Unattended ISO

Use Microsoft's `oscdimg` utility (included with the Windows ADK Deployment Tools) to package the extracted Windows install files and `autounattend.xml` into a UEFI/BIOS dual-boot ISO with full UDF long-filename support.

### Prerequisites

1. **Windows Assessment and Deployment Kit (ADK)**:
   Install the Windows ADK (which includes `oscdimg.exe` under Deployment Tools) via winget:
   ```powershell
   winget install --id Microsoft.WindowsADK --exact --accept-package-agreements --accept-source-agreements
   ```
2. **Source Media & Answer File**:
   A directory containing the extracted contents of the Windows installation ISO with your configured `autounattend.xml` placed in the root directory.

### Command

Run the following command in **PowerShell (Run as Administrator)**:

```powershell
& "<PATH_TO_OSCDIMG>\oscdimg.exe" `
  -m `
  -o `
  -u2 `
  -udfs `
  -bootdata:2#p0,e,b"<SOURCE_DIR>\boot\etfsboot.com"#pEF,e,b"<SOURCE_DIR>\efi\microsoft\boot\efisys.bin" `
  "<SOURCE_DIR>" `
  "<OUTPUT_DIR>\<ISO_NAME>.iso"
```

### Parameter Placeholders

| Placeholder | Description | Example |
| :--- | :--- | :--- |
| `<PATH_TO_OSCDIMG>` | Full path to the directory containing `oscdimg.exe` | `C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg` |
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

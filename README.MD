# Alfred Pennyworth — IT Automation Script
> **Version 1.01** | PowerShell Automation Toolkit

Alfred is a PowerShell script that handles common IT tasks — installing, uninstalling, reinstalling, and troubleshooting applications and Windows — all from a single command.

---

## Requirements

| Requirement | Details |
|---|---|
| **Permissions** | Must be run as **Administrator** |
| **Shell** | PowerShell (Execution Policy Bypass required) |
| **OS** | Windows 10 / 11 |

> ⚠️ **Teams** and **OneDrive** must be run as a **standard user** (not Administrator). Alfred will warn you if you attempt otherwise.

---

## Quick Start

Download and execute Alfred directly from GitHub in one step:

```powershell
Invoke-WebRequest https://raw.githubusercontent.com/13ruce1337/PSTools/main/alfred.ps1 -OutFile "$env:TEMP\alfred.ps1"
Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass", "-File $env:TEMP\alfred.ps1", "<command>", "<application>" -NoNewWindow
```

Replace `<command>` and `<application>` with values from the tables below.

---

## Usage

```
alfred.ps1 <command> <application>
```

### Examples

```powershell
# Install Microsoft Office
.\alfred.ps1 install office

# Reinstall Teams
.\alfred.ps1 reinstall teams

# Uninstall Adobe Acrobat DC
.\alfred.ps1 uninstall adobedc

# Run network troubleshooting
.\alfred.ps1 troubleshoot network

# Update Windows
.\alfred.ps1 update windows

# Nuke (remove ALL) Office applications
.\alfred.ps1 nuke office

# Self-destruct (delete the script)
.\alfred.ps1 sd
```

---

## Supported Applications

| Application | `install` | `uninstall` | `reinstall` | Notes |
|---|:---:|:---:|:---:|---|
| `office` | ✅ | ✅ | ✅ | Microsoft 365 ProPlus (64-bit, Current Channel) |
| `teams` | ✅ | ✅ | ✅ | ⚠️ Must run as **standard user** |
| `visio` | ✅ | ✅ | ✅ | Visio Pro 2021 |
| `project` | ✅ | ✅ | ✅ | Project Pro 2021 |
| `onenote` | ✅ | ✅ | ✅ | OneNote Retail |
| `onedrive` | ✅ | ✅ | ❌ | ⚠️ Must run as **standard user** |
| `adobedc` | ✅ | ✅ | ✅ | Adobe Acrobat DC (Standard, 64-bit) |
| `network` | ❌ | ❌ | ❌ | Used with `troubleshoot` only |

---

## Troubleshoot Targets

| Target | Command | What It Does |
|---|---|---|
| `network` | `troubleshoot network` | Releases/renews IP, flushes & registers DNS, displays full `ipconfig /all` |
| `windows` | `troubleshoot windows` | Runs Windows Update, Disk Optimization, DISM image repair, and SFC scan |

---

## Nuke Targets

> ⚠️ These actions are destructive and irreversible.

| Target | Command | What It Does |
|---|---|---|
| `office` | `nuke office` | Removes **all** Microsoft Office products using ODT `Remove All="TRUE"` |
| `self` | `nuke self` | Removes PSWindowsUpdate module and deletes the script |

---

## Logging

Alfred automatically logs all activity to:

```
C:\Windows\Temp\alfred.log
```

A transcript is started at launch and stopped on exit. Check this file for detailed output and error messages.
  
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

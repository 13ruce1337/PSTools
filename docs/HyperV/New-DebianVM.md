# New-DebianVM

## SYNOPSIS

Creates and starts a Generation 2 Hyper-V VM from a Debian cloud image.

## SYNTAX

### __AllParameterSets

```
New-DebianVM [-Name] <string> -RootPassword <securestring> [-CpuCount <int>] [-MemoryGB <int>]
 [-DiskGB <int>] [-BaseVhdxPath <string>] [-BaseImageUrl <string>] [-RebuildBaseImage]
 [-SeedsPath <string>] [-SwitchName <string>] [-OscdimgPath <string>] [-IsoBootDelaySeconds <int>]
 [-WhatIf] [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

Copies a base Debian cloud VHDX, resizes it, builds a cloud-init NoCloud
seed ISO with oscdimg (Windows ADK Deployment Tools), creates the VM,
attaches both, and starts it.

If the base VHDX does not exist (or -RebuildBaseImage is specified), it
is created automatically: qemu-img is located (QEMU is installed via
winget if needed), the Debian cloud qcow2 image is downloaded, and it is
converted to VHDX at -BaseVhdxPath.

Requirements:
  - Elevated session with the Hyper-V PowerShell module
  - Windows ADK "Deployment Tools" (oscdimg.exe); auto-installed via
    winget (Microsoft.WindowsADK) if missing
  - winget (needed if oscdimg.exe or QEMU is not already installed)

## EXAMPLES

### EXAMPLE 1

$pw = Read-Host -AsSecureString -Prompt 'Root password'
New-DebianVM -Name web01 -CpuCount 2 -MemoryGB 4 -DiskGB 40 -RootPassword $pw

### EXAMPLE 2

New-DebianVM -Name test01 -RootPassword $pw -WhatIf

### EXAMPLE 3

New-DebianVM -Name test02 -RootPassword $pw -RebuildBaseImage

## PARAMETERS

### -BaseImageUrl

URL of the Debian cloud qcow2 image used to build the base VHDX.

```yaml
Type: System.String
DefaultValue: https://cloud.debian.org/images/cloud/bookworm/latest/debian-12-genericcloud-amd64.qcow2
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -BaseVhdxPath

Path to the base Debian VHDX.
Created automatically if it does not exist.

```yaml
Type: System.String
DefaultValue: (Join-Path $env:USERPROFILE 'Images\debian-12-base.vhdx')
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Confirm

Prompts you for confirmation before running the cmdlet.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: ''
SupportsWildcards: false
Aliases:
- cf
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -CpuCount

Number of virtual processors.
Default 1.

```yaml
Type: System.Int32
DefaultValue: 1
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -DiskGB

Target disk size in GB.
If smaller than the base image, the base image
size is used instead.

```yaml
Type: System.Int32
DefaultValue: 20
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -IsoBootDelaySeconds

Seconds to wait after the first boot reaches Running, before the
one-time reboot that lets cloud-init consume the NoCloud seed.
Default 4.
Set to 0 to reboot as soon as the VM is Running.

```yaml
Type: System.Int32
DefaultValue: 4
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -MemoryGB

Startup memory in GB.
Default 2.

```yaml
Type: System.Int32
DefaultValue: 2
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Name

VM name.
Also used as the guest hostname.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: 0
  IsRequired: true
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -OscdimgPath

Path to oscdimg.exe.
If it does not exist, the Windows ADK is
installed automatically via winget before this path is re-checked.

```yaml
Type: System.String
DefaultValue: C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -RebuildBaseImage

Download and convert the base image again, replacing any existing
base VHDX.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: False
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -RootPassword

Root password for the guest, as a SecureString.

```yaml
Type: System.Security.SecureString
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: true
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -SeedsPath

Folder where the cloud-init seed ISO is stored.

```yaml
Type: System.String
DefaultValue: (Join-Path $env:USERPROFILE 'Seeds')
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -SwitchName

Hyper-V virtual switch to connect to.

```yaml
Type: System.String
DefaultValue: Default Switch
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -WhatIf

Runs the command in a mode that only reports what would happen without performing the actions.

```yaml
Type: System.Management.Automation.SwitchParameter
DefaultValue: ''
SupportsWildcards: false
Aliases:
- wi
ParameterSets:
- Name: (All)
  Position: Named
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### Microsoft.HyperV.PowerShell.VirtualMachine

{{ Fill in the Description }}

## NOTES

## RELATED LINKS

{{ Fill in the related links here }}


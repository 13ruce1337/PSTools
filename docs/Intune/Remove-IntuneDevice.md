---
document type: cmdlet
external help file: PSTools-Help.xml
HelpUri: ''
Locale: en-US
Module Name: PSTools
PlatyPS schema version: 2024-05-01
title: Remove-IntuneDevice
---

# Remove-IntuneDevice

## SYNOPSIS

Removes devices from Intune by serial number, and optionally from Windows Autopilot and Entra ID.

## SYNTAX

### Device (Default)

```
Remove-IntuneDevice [-SerialNumber] <string> [[-Hostname] <string>] [-IncludeAutopilot]
 [-IncludeEntra] [-TimeoutSeconds <int>] [-PollIntervalSeconds <int>] [-Force] [-UseWAM] [-WhatIf]
 [-Confirm]
```

### Csv

```
Remove-IntuneDevice -CsvPath <string> [-IncludeAutopilot] [-IncludeEntra] [-TimeoutSeconds <int>]
 [-PollIntervalSeconds <int>] [-Force] [-UseWAM] [-WhatIf] [-Confirm]
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

Finds Intune managed devices by serial number and deletes them.
Optionally
also deregisters the device from Windows Autopilot and deletes the device
object from Entra ID.

Records are always removed in the order Microsoft recommends:
  1.
Intune
  2.
Windows Autopilot   (-IncludeAutopilot, or implied by -IncludeEntra)
  3.
Entra ID            (-IncludeEntra)

-IncludeEntra always runs the Autopilot step first.
The Entra step only
starts once the Autopilot identity is confirmed to be gone.
Autopilot
deletion is asynchronous, so the command polls until the identity
disappears or -TimeoutSeconds is reached.
If a step fails, is declined,
or times out, the later steps are skipped for that device.
Running the
command again later finishes the remaining steps.

-Hostname is an optional safety check.
If the serial number has Intune
records but none of them have that device name, the device is skipped.
If the serial number has no Intune record, the hostname cannot be
verified and processing continues with Autopilot / Entra.

Entra devices are found through the Entra device IDs on the Intune and
Autopilot records.
If none are found and a hostname was given, the Entra
device is looked up by display name, but only if exactly one device has
that name.
Synced (hybrid joined) Entra devices are not deleted; remove
them from on-premises Active Directory instead.

Batches (CSV or pipeline input) are processed in phases: every Intune
deletion, then every Autopilot deletion request, then one shared wait,
then every Entra deletion.
You wait once for the whole batch, not once
per device.

Requested Microsoft Graph scopes:
  DeviceManagementManagedDevices.ReadWrite.All   (always)
  DeviceManagementServiceConfig.ReadWrite.All    (Autopilot)
  Device.ReadWrite.All                           (Entra)

## EXAMPLES

### EXAMPLE 1

Remove-IntuneDevice -SerialNumber 5CG1234ABC -WhatIf

Shows which Intune records would be deleted for the serial number.

### EXAMPLE 2

Remove-IntuneDevice -SerialNumber 5CG1234ABC -Hostname LAPTOP-042 -IncludeEntra

Deletes the Intune record (only if it is named LAPTOP-042), deregisters
the device from Autopilot, waits for that to finish, then deletes the
Entra device.

### EXAMPLE 3

Remove-IntuneDevice -CsvPath .\retired.csv -IncludeEntra -Force -TimeoutSeconds 1800 | Export-Csv .\retired-results.csv -NoTypeInformation

Batch removal from a CSV without prompting, allowing up to 30 minutes in
total for Autopilot deletions, and saving the results.

### EXAMPLE 4

Import-Csv .\retired.csv | Remove-IntuneDevice -IncludeAutopilot

Same as -CsvPath, using pipeline input.

## PARAMETERS

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

### -CsvPath

Path to a CSV file with a 'SerialNumber' (or 'Serial') column and an
optional 'Hostname' (or 'DeviceName' / 'ComputerName') column.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: Csv
  Position: Named
  IsRequired: true
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Force

Do not prompt for confirmation before each deletion.

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

### -Hostname

Optional device name used as a safety check against the Intune record.
Accepts pipeline input by property name (aliases: DeviceName, ComputerName).

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases:
- DeviceName
- ComputerName
ParameterSets:
- Name: Device
  Position: 1
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -IncludeAutopilot

Also deregister the device from Windows Autopilot (after Intune).

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

### -IncludeEntra

Also delete the device from Entra ID.
Implies -IncludeAutopilot; Entra
deletion only runs after the Autopilot identity is confirmed removed.

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

### -PollIntervalSeconds

Delay between Autopilot status checks / retries.
Default 15.

```yaml
Type: System.Int32
DefaultValue: 15
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

### -SerialNumber

Serial number of the device.
Accepts pipeline input by property name
(alias: Serial).

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases:
- Serial
ParameterSets:
- Name: Device
  Position: 0
  IsRequired: true
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -TimeoutSeconds

Maximum total time, in seconds, to wait for Autopilot deletions.
This is
one shared limit for the whole run (a single device or an entire batch),
not a per-device limit.
It covers both retrying deletes that Autopilot
rejects while a just-deleted Intune record is still being processed, and
waiting for accepted deletes to disappear.
Waiting stops as soon as every
identity is gone.
Default 600 (10 minutes).
Autopilot deletion usually
completes within a few minutes but can take longer; anything still
pending is reported as TimedOut and can be finished by running the
command again.

```yaml
Type: System.Int32
DefaultValue: 600
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

### -UseWAM

Passed to Assert-MgGraphConnection: do not disable the Windows Web
Account Manager broker for sign-in.

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

### System.String

{{ Fill in the Description }}

## OUTPUTS

### One object per input row with SerialNumber

{{ Fill in the Description }}

### System.Management.Automation.PSObject

{{ Fill in the Description }}

## NOTES

## RELATED LINKS

{{ Fill in the related links here }}


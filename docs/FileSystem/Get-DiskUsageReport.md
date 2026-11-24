---
document type: cmdlet
external help file: PSTools-Help.xml
HelpUri: ''
Locale: en-US
Module Name: PSTools
PlatyPS schema version: 2024-05-01
title: Get-DiskUsageReport
---

# Get-DiskUsageReport

## SYNOPSIS

Shows the largest folders (recursively) and the largest files under a path.

## SYNTAX

### __AllParameterSets

```
Get-DiskUsageReport [[-Path] <string[]>] [-Top <int>] [-PassThru]
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

Starting from the current directory (or -Path), scans all subfolders and reports:
  1.
The total size and file count of the starting folder.
  2.
The top N largest subfolders, with their total size
     (including all subfolders) and file count.
  3.
The top N largest files with name, location and size.

Output is written to the console.
Use -PassThru to also return a report
object.

Items that cannot be read (access denied, path too long) are skipped and
counted; sizes may be understated in that case.
Run elevated for full
results.

## EXAMPLES

### EXAMPLE 1

Get-DiskUsageReport

### EXAMPLE 2

Get-DiskUsageReport -Path D:\Shares -Top 25

### EXAMPLE 3

Get-ChildItem C:\Users -Directory | Get-DiskUsageReport -Top 5

### EXAMPLE 4

$report = Get-DiskUsageReport -Path C:\Data -PassThru
$report.TopFolders | Select-Object -First 3

## PARAMETERS

### -PassThru

Return a report object in addition to the on-screen report.

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

### -Path

Folder(s) to scan.
Defaults to the current directory.
Accepts pipeline
input (strings or objects with a FullName property, e.g.
Get-ChildItem).

```yaml
Type: System.String[]
DefaultValue: .
SupportsWildcards: false
Aliases:
- FullName
ParameterSets:
- Name: (All)
  Position: 0
  IsRequired: false
  ValueFromPipeline: true
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Top

Number of largest folders and largest files to list.
Default 10.

```yaml
Type: System.Int32
DefaultValue: 10
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

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

### System.String[]

{{ Fill in the Description }}

## OUTPUTS

## NOTES

## RELATED LINKS

{{ Fill in the related links here }}


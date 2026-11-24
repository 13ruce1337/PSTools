# Get-M365MailboxSpaceReport

## SYNOPSIS

Analyzes where space is being used in an Exchange Online mailbox.

## SYNTAX

### __AllParameterSets

```
Get-M365MailboxSpaceReport [-Identity] <string> [-Days <int>] [-MinimumSizeMB <double>]
 [-ChunkDays <int>] [-OutputPath <string>] [-NoExport] [-PassThru] [-UserPrincipalName <string>]
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

Produces a mailbox space report containing:
    1.
Mailbox overview (size, item counts, deleted items, last logon, quotas)
    2.
All folders sorted by size (largest first)
    3.
Recoverable Items folders (a common hidden space hog)
    4.
Archive mailbox statistics (if an archive is enabled)
    5.
Sent Items folder detail
    6.
Large sent messages over the last -Days days, via Get-MessageTraceV2
       (queried in chunks of -ChunkDays, since MessageTraceV2 allows at
       most 10 days per query and 90 days of history)

Message trace returns one row per recipient; rows are grouped by
MessageId so each message is listed once with its recipients.

Folder statistics and large sent messages are exported to CSV in
-OutputPath unless -NoExport is specified.

## EXAMPLES

### EXAMPLE 1

Get-M365MailboxSpaceReport -Identity trevor.cooper@contoso.com

### EXAMPLE 2

Get-M365MailboxSpaceReport -Identity trevor.cooper@contoso.com -Days 30 -MinimumSizeMB 10 -OutputPath C:\Temp

### EXAMPLE 3

$report = Get-M365MailboxSpaceReport -Identity trevor.cooper@contoso.com -NoExport -PassThru
$report.Folders | Select-Object -First 5

## PARAMETERS

### -ChunkDays

Size of each message trace query window in days (1-10).
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

### -Days

How many days of message trace history to search (1-90).
Default 90.

```yaml
Type: System.Int32
DefaultValue: 90
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

### -Identity

The mailbox to analyze (email address, UPN or alias).
Accepts pipeline
input by PrimarySmtpAddress.

```yaml
Type: System.String
DefaultValue: ''
SupportsWildcards: false
Aliases:
- UserEmail
- PrimarySmtpAddress
- Mailbox
ParameterSets:
- Name: (All)
  Position: 0
  IsRequired: true
  ValueFromPipeline: true
  ValueFromPipelineByPropertyName: true
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -MinimumSizeMB

Only sent messages larger than this size (in MB) are reported.
Default 5.

```yaml
Type: System.Double
DefaultValue: 5
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

### -NoExport

Do not write CSV files.

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

### -OutputPath

Folder for the CSV exports.
Created if it does not exist.
Default: $env:TEMP\PSTools

```yaml
Type: System.String
DefaultValue: (Join-Path $env:TEMP 'PSTools')
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

### -UserPrincipalName

Optional UPN used if a new Exchange Online connection is required.

```yaml
Type: System.String
DefaultValue: ''
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

### System.String

{{ Fill in the Description }}

## OUTPUTS

## NOTES

Requires ExchangeOnlineManagement 3.7.0+ for Get-MessageTraceV2.
If it is
not available, the large sent messages section is skipped.


## RELATED LINKS

{{ Fill in the related links here }}


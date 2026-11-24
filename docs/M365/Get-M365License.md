# Get-M365License

## SYNOPSIS

Reports the last sign-in of licensed Microsoft 365 users, to find licenses
that are assigned but not used.

## SYNTAX

### __AllParameterSets

```
Get-M365License [[-DaysInactive] <int>] [[-OutputPath] <string>] [-IncludeDisabled] [-Export]
 [-PassThru]
```

## ALIASES

This cmdlet has the following aliases,
  {{Insert list of aliases}}

## DESCRIPTION

Reads every user with at least one license assigned from Microsoft Graph,
resolves license SKU IDs to their part numbers, and works out each user's
most recent sign-in (interactive or non-interactive).
Each user gets a
status based on -DaysInactive:

    Active          Signed in within the last -DaysInactive days
    Inactive        Last sign-in is older than -DaysInactive days
    NeverSignedIn   No recorded sign-in
    New             No sign-in yet, but created within the last
                    -DaysInactive days (grace period)

Disabled accounts are skipped unless -IncludeDisabled is specified.

The console shows a summary by status and a table of the users that look
unused (Inactive and NeverSignedIn).
The table is sorted by last sign-in:
the most recent sign-in is at the top, and the longest time since last
sign-in is at the bottom.
Users who never signed in are listed last.

Nothing is written to disk unless -Export or -OutputPath is specified.
When exporting, every licensed user (not only the unused ones) is written
to a CSV file in -OutputPath, in the same order.

Use -PassThru to also return every record to the pipeline.

## EXAMPLES

### EXAMPLE 1

Get-M365LicenseActivity

### EXAMPLE 2

Get-M365LicenseActivity -DaysInactive 90 -IncludeDisabled

### EXAMPLE 3

Get-M365LicenseActivity -DaysInactive 90 -OutputPath C:\Reports

### EXAMPLE 4

Get-M365LicenseActivity -PassThru | Where-Object Licenses -match 'SPE_E3'

## PARAMETERS

### -DaysInactive

Number of days without a sign-in before a licensed user is considered
inactive.
Default 30.

```yaml
Type: System.Int32
DefaultValue: 30
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: 0
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -Export

Export the full results to a CSV file in -OutputPath.

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

### -IncludeDisabled

Also include disabled (sign-in blocked) accounts that still hold licenses.

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

Folder for the CSV export.
Created if it does not exist.
Specifying this
parameter also turns on the export.
Default: $env:TEMP\PSTools

```yaml
Type: System.String
DefaultValue: (Join-Path $env:TEMP 'PSTools')
SupportsWildcards: false
Aliases: []
ParameterSets:
- Name: (All)
  Position: 1
  IsRequired: false
  ValueFromPipeline: false
  ValueFromPipelineByPropertyName: false
  ValueFromRemainingArguments: false
DontShow: false
AcceptedValues: []
HelpMessage: ''
```

### -PassThru

Return all records to the pipeline, in addition to the on-screen report.

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

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable,
-InformationAction, -InformationVariable, -OutBuffer, -OutVariable, -PipelineVariable,
-ProgressAction, -Verbose, -WarningAction, and -WarningVariable. For more information, see
[about_CommonParameters](https://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

Requires the Microsoft.Graph.Authentication module and the delegated scopes:
    User.Read.All, AuditLog.Read.All, Organization.Read.All
An existing Graph session with these scopes is reused.

Sign-in activity (signInActivity) requires Entra ID P1 or P2.
It is not
updated in real time and can lag by several hours.


## RELATED LINKS

{{ Fill in the related links here }}


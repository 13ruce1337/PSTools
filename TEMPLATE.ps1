<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER <Parameter_Name>

.INPUTS

.OUTPUTS

.NOTES
  Version:        1.0
  Author:         <Name>
  Creation Date:  <Date>
  Purpose/Change: Initial script development
  
.EXAMPLE
  
.LINK
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

[cmdletbinding()]
param()

#Assemblies

#Preferences
#$ErrorActionPreference = "SilentlyContinue"
#$DebugPreference = 'SilentlyContinue'
$ProgressPreference = $false

#Dot Source required Function Libraries

#Script Variables

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
[float]$script:script_version = 1.0

#Logging
[string]$script:log_path = "C:\Windows\Temp"
[string]$script:log_name = "<script_name>.log"
[string]$script:log_file = Join-Path -Path $script:log_path -ChildPath $script:log_name
Start-Transcript -Force -IncludeInvocationHeader -Path $script:log_file

#-----------------------------------------------------------[Functions]------------------------------------------------------------



#-----------------------------------------------------------[Execution]------------------------------------------------------------

Stop-Transcript
Exit 0
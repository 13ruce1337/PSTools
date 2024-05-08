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

#$ErrorActionPreference = "SilentlyContinue"
$ProgressPreference = $false

#Dot Source required Function Libraries

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
[float]$script:script_version = 1.0

#Log File Info
[string]$script:log_path = "C:\Windows\Temp"
[string]$script:log_name = "<script_name>.log"
[string]$script:log_file = Join-Path -Path $sLogPath -ChildPath $sLogName

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Log {
  param(
      [Parameter(Mandatory=$true)][String]$private:msg
  )

  [string]$private:time = Get-Date -Format "HH:mm:ss"
  [string]$private:text = "[$private:time]:$private:msg"
  
  Write-Host $private:text
  Add-Content -Path $script:log_file -Value "$private:text"
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

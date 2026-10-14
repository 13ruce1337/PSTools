function Test-IsAdmin {
    <#
    .SYNOPSIS
        Returns $true when the current session is elevated.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $principal = New-Object Security.Principal.WindowsPrincipal(
        [Security.Principal.WindowsIdentity]::GetCurrent()
    )
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

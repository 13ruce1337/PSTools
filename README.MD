# Alfred Pennyworth
### Makes troubleshooting basic tasks easier.
### Version 0.2

==Must be ran as administrator.==

`Invoke-WebRequest https://raw.githubusercontent.com/13ruce1337/PSTools/main/alfred.ps1 -OutFile "$env:TEMP\alfred.ps1"`\
`start-process powershell.exe -Argumentlist "-ExecutionPolicy Bypass", "-File $env:TEMP\alfred.ps1", "install teams" -NoNewWindow`

## COMMANDS
- install
- uninstall
- reinstall
- update
- troubleshoot
- sd
## APPLICATIONS
- office
- teams
- visio
# Alfred Pennyworth
**Must be ran as administrator**

`Invoke-WebRequest https://raw.githubusercontent.com/13ruce1337/PSTools/main/alfred.ps1 -OutFile "$env:TEMP\alfred.ps1"`\
`start-process powershell.exe -Argumentlist "-ExecutionPolicy Bypass", "-File $env:TEMP\alfred.ps1", "update windows" -NoNewWindow`

## COMMANDS
- install
- uninstall
- reinstall
- nuke
- update
- troubleshoot
- sd
## APPLICATIONS
- office
- teams
- project
- visio
- onenote
- adobedc

## Roadmap
### 0.5
- full odt capabilities written out
- add any additional Microsoft enterprise apps
    - OneDrive
- better error reporting

### 0.7
- any additional 3rd party enterprise apps like Zoom
- generic system report
- logs

### 0.9
- polish on grammar/spelling
- uniformity
- comments in code
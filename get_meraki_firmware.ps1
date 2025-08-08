<#
.SYNOPSIS
    Get Meraki Firmware

.DESCRIPTION
    Using Meraki's API, it grabs all devices and lists some information including the firmware. 
    It also grabs the latest firmware version.
    !Remember to enter your API key!

.OUTPUTS
    Outputs a CSV file in the same directory as this file.

.NOTES
    Version:        1.0
    Author:         Trevor Cooper
    Creation Date:  8/7/2025
    Purpose/Change: Initial script development

    Addendum: I'm aware /networks/$($network.id)/devices is deprecated and am working to replace.
  
.LINK
    https://developer.cisco.com/meraki/api-v1/
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

[cmdletbinding()]
param()

#Preferences
$ProgressPreference = "SilentlyContinue"

#Script Variables
################
$api_key = ""
################

$base_url = "https://api.meraki.com/api/v1"
$headers = @{
    "X-Cisco-Meraki-API-Key" = $api_key
    "Content-Type" = "application/json"
}
$device_list = @()

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
[float]$script:script_version = 1.0

#Logging
[string]$script:log_path = "C:\Windows\Temp"
[string]$script:log_name = "get_meraki_firmware.log"
[string]$script:log_file = Join-Path -Path $script:log_path -ChildPath $script:log_name
Start-Transcript -Force -IncludeInvocationHeader -Path $script:log_file

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function main {
    $orgs_response = Invoke-RestMethod -Method Get -Uri "$base_url/organizations" -Headers $headers
    $orgs = $orgs_response | Where-Object { $null -ne $_.id }

    $device_list = @()

    foreach ($org in $orgs) {
        $networks_uri = "$base_url/organizations/$($org.id)/networks"
        $devices_uri = "$base_url/organizations/$($org.id)/devices"
        try {
            $networks = Invoke-RestMethod -Method Get -Uri $networks_uri -Headers $headers
            foreach ($network in $networks) {
                Write-Host "Getting information from $($network.name)"
                $firmware_info_uri = "$base_url/networks/$($network.id)/firmwareUpgrades"
                $devices_uri = "$base_url/networks/$($network.id)/devices"
                $firmware_info = Invoke-RestMethod -Method Get -Uri $firmware_info_uri -Headers $headers
                $devices = Invoke-RestMethod -Method Get -Uri $devices_uri -Headers $headers
                Write-Host "Found $($devices.Count) device(s)"
                foreach ($device in $devices) {
                    $device_list += [PSCustomObject]@{
                        org         = $org.name
                        name        = $device.name
                        model       = $device.model
                        location    = $device.address
                        firmware    = $device.firmware
                        serial      = $device.serial
                        mac         = $device.mac
                    }
                }
                $product_types = @("appliance", "cellularGateway", "sensor", "switch", "switchCatalyst", "wireless")
                $results = @()
                foreach ($type in $product_types) {
                    if ($firmware_info.products.$type) {
                        Write-Host "Getting lastest firmware version for $type"
                        $latest = $firmware_info.products.$type.availableVersions[-1].shortName
                        $results += [PSCustomObject]@{
                            name     = "Latest $type"
                            firmware = $latest
                        }
                    }
                }
                $device_list += $results
                $device_list += [PSCustomObject]@{}
            } 
        } catch {
            Write-Warning "Failed to fetch firmware info for organization $($org.name): $_"
        }
    }

    $csv_path = "Meraki_Firmware_Report.csv"
    $device_list | Export-Csv -Path $csv_path -NoTypeInformation
    Write-Host "Firmware report exported to $csv_path"

}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

main

Stop-Transcript
Exit 0
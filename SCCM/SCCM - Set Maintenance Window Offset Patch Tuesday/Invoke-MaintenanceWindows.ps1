<#.SYNOPSISThis script invokes a Maintenance Window script with different parameters    .DESCRIPTIONThis script invokes a Maintenance Window script with different parametersAuthor: Daniel Classon (www.danielclasson.com)Version: 1.0Date: 2018-12-11    .EXAMPLE.\Invoke-MaintenanceWindows -CollID1 P01000AD -CollID2 P01000ACWill invoke script for two different collections    .DISCLAIMERAll scripts and other Powershell references are offered AS IS with no warranty.These script and functions are tested in my environment and it is recommended that you test these scripts in a test environment before using in your production environment.#>    PARAM(
    [string]$CollID1,
    [string]$CollID2
    )  


$ErrorMessage = $_.Exception.Message

Function Write-Log
{
    PARAM(
    [String]$Message,
    [int]$Severity,
    [string]$Component
    )

    Set-Location $PSScriptRoot
    $Logpath = "\\prod-p01\logs$\Set Maintenance Windows"
    $TimeZoneBias = Get-WMIObject -Query "Select Bias from Win32_TimeZone"
    $Date= Get-Date -Format "HH:mm:ss.fff"
    $Date2= Get-Date -Format "MM-dd-yyyy"
    $Type=1
    "<![LOG[$Message]LOG]!><time=$([char]34)$Date$($TimeZoneBias.bias)$([char]34) date=$([char]34)$date2$([char]34) component=$([char]34)$Component$([char]34) context=$([char]34)$([char]34) type=$([char]34)$Severity$([char]34) thread=$([char]34)$([char]34) file=$([char]34)$([char]34)>"| Out-File -FilePath "$LogPath\Invoke-MaintenanceWindows.log" -Append -NoClobber -Encoding default
}

Try {

    #Invoke Maintenance Windows creation script for collection 1
    & "$PSScriptRoot\Set-MaintenanceWindow.ps1" -OffSetWeeks 1 -OffSetDays 3 -AddStartHour 0 -AddStartMinutes 0 -AddEndHour 23 -AddEndMinutes 59 -CollID $CollID1 -ErrorAction Stop
    & "$PSScriptRoot\Set-MaintenanceWindow.ps1" -OffSetWeeks 1 -OffSetDays 4 -AddStartHour 0 -AddStartMinutes 0 -AddEndHour 23 -AddEndMinutes 59 -CollID $CollID1 -ErrorAction Stop
    & "$PSScriptRoot\Set-MaintenanceWindow.ps1" -OffSetWeeks 1 -OffSetDays 5 -AddStartHour 0 -AddStartMinutes 0 -AddEndHour 23 -AddEndMinutes 59 -CollID $CollID1 -ErrorAction Stop
}

Catch {
    Write-Warning "$_.Exception.Message"
    Write-Log -Message "Error: $_.Exception.Message" -Severity 3 -Component "Invoke Maintenance Windows script"
}

Try {

    #Invoke Maintenance Windows creation script for collection 2
    & "$PSScriptRoot\Set-MaintenanceWindow.ps1" -OffSetWeeks 2 -OffSetDays 3 -AddStartHour 0 -AddStartMinutes 0 -AddEndHour 23 -AddEndMinutes 59 -CollID $CollID2 -ErrorAction Stop
    & "$PSScriptRoot\Set-MaintenanceWindow.ps1" -OffSetWeeks 2 -OffSetDays 4 -AddStartHour 0 -AddStartMinutes 0 -AddEndHour 23 -AddEndMinutes 59 -CollID $CollID2 -ErrorAction Stop
    & "$PSScriptRoot\Set-MaintenanceWindow.ps1" -OffSetWeeks 2 -OffSetDays 5 -AddStartHour 0 -AddStartMinutes 0 -AddEndHour 23 -AddEndMinutes 59 -CollID $CollID2 -ErrorAction Stop

}

Catch {
    Write-Warning "$_.Exception.Message"
    Write-Log -Message "Error: $_.Exception.Message" -Severity 3 -Component "Invoke Maintenance Windows script"
}
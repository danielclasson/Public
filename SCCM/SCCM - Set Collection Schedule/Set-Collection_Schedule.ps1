<#
.SYNOPSIS
Sets the collection refresh type for all collections in a text file.

.PARAMETER CollectionType
Sets the type of Collection, either Device or User

.PARAMETER DayofWeek
Sets the day of the week

.PARAMETER RecurCount
Sets how often the schudule should
    
.DESCRIPTION

Author: Daniel Classon
Version: 1.0
Date: 2015/05/19
    
.EXAMPLE
.\Set-Collection_Schedule.ps1 -CollectionType Device -Interval Day -Recurrance 3
Will set the device collection refresh schedule to run every 3 days

.DISCLAIMER
All scripts and other powershell references are offered AS IS with no warranty.
These script and functions are tested in my environment and it is recommended that you test these scripts in a test environment before using in your production environment.
#>

[CmdletBinding()]Param(    [Parameter(Mandatory=$True, Helpmessage="Enter the Collection Type (Device/User)")]    [string]$CollectionType,    [Parameter(Mandatory=$True, Helpmessage="Enter the day of week")]    [string]$DayofWeek,    [Parameter(Mandatory=$True, Helpmessage="Enter how often schedule should run")]    [int]$RecurCount)

Begin {
    #Checks if the user is in the administrator group. Warns and stops if the user is not.
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "You are not running this as local administrator. Run it again in an elevated prompt." ; break
    }
    try {
        Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)
    }
    catch [System.UnauthorizedAccessException] {
        Write-Warning -Message "Access denied" ; break
    }
    catch [System.Exception] {
        Write-Warning "Unable to load the Configuration Manager Powershell module from $env:SMS_ADMIN_UI_PATH" ; break
    }
}
Process {
    try {
        $CollectionIDs = Get-Content "collections.txt" -ErrorAction Stop
    }
    catch [System.Exception] {
        Write-Warning "Unable to find collections.txt. Make sure to place it in the script directory."
    }
    $SiteCode = Get-PSDrive -PSProvider CMSITE
    Set-Location -Path "$($SiteCode.Name):\"

    $Schedule = New-CMSchedule -DayOfWeek $DayofWeek -RecurCount $RecurCount
    $Count = 0

    Foreach ($CollectionID in $CollectionIDs) {
        $Count++
        if ($CollectionType -eq "Device") {
            $Collection = Get-CMDeviceCollection -CollectionId $CollectionID
            Set-CMDeviceCollection -CollectionId $CollectionID -RefreshSchedule $Schedule
                Write-Progress -Activity "Setting Refresh Schedule  on Collection ID: $CollectionID" -Status "Modified $Count of $($CollectionIDs.Count) collections" -PercentComplete ($Count / $CollectionIDs.count * 100)
                Write-Host "Setting Refresh Schedule for Collection : " $Collection.CollectionID "`t" $Collection.Name -ForegroundColor Yellow
            }
        } 
        if ($CollectionType -eq "User") {
            $Collection = Get-CMUserCollection -CollectionId $CollectionID
            Set-CMUserCollection -CollectionId $CollectionID -RefreshSchedule $Schedule
            Write-Progress -Activity "Setting Refresh Schedule  on Collection ID: $CollectionID" -Status "Modified $Count of $($CollectionIDs.Count) collections" -PercentComplete ($Count / $CollectionIDs.count * 100)
            Write-Host "Setting Refresh Schedule for Collection : " $Collection.CollectionID "`t" $Collection.Name -ForegroundColor Yellow
        } 
}  
End {
    Write-Host "$Count $CollectionType collections were updated"  
    Set-Location -Path $env:SystemDrive
}
<#
.PARAMETER COLLECTIONTYPE
Define the type of collections in the collections.txt file. Valid inputs are Device and User

.PARAMETER REFRESHTYPE
Define the Refresh type for the collection. Valid inputs are:

# The following refresh types exist for ConfigMgr collections 
# 6 = Incremental and Scheduled Updates 
# 4 = Incremental Updates Only 
# 2 = Scheduled Updates only 
# 1 = Manual Update only 
    
.DESCRIPTION
Sets the collection refresh type for all collections defined in a text file.

.NOTES
Author: Daniel Classon
Version: 1.0
Date: 2015/05/17
    
.EXAMPLE
.\Set-Collection_Updates.ps1 -CollectionType Device -RefreshType 2
Will set the collections in collections.txt to "Scheduled Updates only".

.DISCLAIMER
All scripts and other powershell references are offered AS IS with no warranty.
These script and functions are tested in my environment and it is recommended that you test these scripts in a test environment before using in your production environment.
#>

[CmdletBinding()]

Param(
    [Parameter(Mandatory=$True, Helpmessage="Enter the Collection Type")]
    [string]$CollectionType,
    [Parameter(Mandatory=$True, Helpmessage="Enter the Refresh Type")]
    [string]$RefreshType
)

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
    
    $Count = 0

    Foreach ($CollectionID in $CollectionIDs) {
        $Count++
        if ($CollectionType -eq "Device") {
            $Collection = Get-CMDeviceCollection -CollectionId $CollectionID
            $Collection.RefreshType = $RefreshType
            $Collection.Put()
            if ($RefreshType -eq 1) {
                Write-Progress -Activity "Enabling Manual Update only  on $CollectionType Collection ID: $CollectionID" -Status "Modified $Count of $($CollectionIDs.Count) collections" -PercentComplete ($Count / $CollectionIDs.count * 100)
                Write-Host "Enabling Manual Update only on: " $Collection.CollectionID "`t" $Collection.Name -ForegroundColor Yellow
            }
            elseif ($RefreshType -eq 2) {
                Write-Progress -Activity "Enabling Scheduled Updates only on $CollectionType Collection ID: $CollectionID" -Status "Modified $Count of $($CollectionIDs.Count) collections" -PercentComplete ($Count / $CollectionIDs.count * 100)
                Write-Host "Enabling Scheduled Updates only on: " $Collection.CollectionID "`t" $Collection.Name -ForegroundColor Yellow
            }
            elseif ($RefreshType -eq 4) {
                Write-Progress -Activity "Enabling Incremental Updates only on $CollectionType Collection ID: $CollectionID" -Status "Modified $Count of $($CollectionIDs.Count) collections" -PercentComplete ($Count / $CollectionIDs.count * 100)
                Write-Host "Enabling Incremental Updates Only on: " $Collection.CollectionID "`t" $Collection.Name -ForegroundColor Yellow
            }
            elseif ($RefreshType -eq 6) {
                Write-Progress -Activity "Enabling Incremental and Scheduled Updates on $CollectionType Collection ID: $CollectionID" -Status "Modified $Count of $($CollectionIDs.Count) collections" -PercentComplete ($Count / $CollectionIDs.count * 100)
                Write-Host "Enabling Incremental and Scheduled Updates on: " $Collection.CollectionID "`t" $Collection.Name -ForegroundColor Yellow
        }
        } 
        if ($CollectionType -eq "User") {
            $Collection = Get-CMUserCollection -CollectionId $CollectionID
            $Collection.RefreshType = $RefreshType
            $Collection.Put()
            if ($RefreshType -eq 1) {
                Write-Progress -Activity "Enabling Manual Update only  on $CollectionType Collection ID: $CollectionID" -Status "Modified $Count of $($CollectionIDs.Count) collections" -PercentComplete ($Count / $CollectionIDs.count * 100)
                Write-Host "Enabling Manual Update only on: " $Collection.CollectionID "`t" $Collection.Name -ForegroundColor Yellow
            }
            elseif ($RefreshType -eq 2) {
                Write-Progress -Activity "Enabling Scheduled Updates only on $CollectionType Collection ID: $CollectionID" -Status "Modified $Count of $($CollectionIDs.Count) collections" -PercentComplete ($Count / $CollectionIDs.count * 100)
                Write-Host "Enabling Scheduled Updates only on: " $Collection.CollectionID "`t" $Collection.Name -ForegroundColor Yellow
            }
            elseif ($RefreshType -eq 4) {
                Write-Progress -Activity "Enabling Incremental Updates only on $CollectionType Collection ID: $CollectionID" -Status "Modified $Count of $($CollectionIDs.Count) collections" -PercentComplete ($Count / $CollectionIDs.count * 100)
                Write-Host "Enabling Incremental Updates Only on: " $Collection.CollectionID "`t" $Collection.Name -ForegroundColor Yellow
            }
            elseif ($RefreshType -eq 6) {
                Write-Progress -Activity "Enabling Incremental and Scheduled Updates on $CollectionType Collection ID: $CollectionID" -Status "Modified $Count of $($CollectionIDs.Count) collections" -PercentComplete ($Count / $CollectionIDs.count * 100)
                Write-Host "Enabling Incremental and Scheduled Updates on: " $Collection.CollectionID "`t" $Collection.Name -ForegroundColor Yellow
        }
        } 
} 
}  
End {
    Write-Host "$Count $CollectionType collections were updated"  
    Set-Location -Path $env:SystemDrive
}
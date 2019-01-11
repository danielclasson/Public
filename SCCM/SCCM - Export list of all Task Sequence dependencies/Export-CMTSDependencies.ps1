<#
.SYNOPSIS
Exports a CSV of all the depedencies for a given Task Sequence.
    
.DESCRIPTION
Exports a CSV of all the depedencies for a given Task Sequence.

Author: Daniel Classon
Version: 1.0
Date: 2015/05/12
    
.EXAMPLE
.\Export-Task_Sequence_Dependencies.ps1 -TSID P0000061
Exports a CSV of all the depedencies for Task Sequence with ID P0000061.
    
.DISCLAIMER
All scripts and other powershell references are offered AS IS with no warranty.
These script and functions are tested in my environment and it is recommended that you test these scripts in a test environment before using in your production environment.
#>[CmdletBinding()]param(        [Parameter(Mandatory=$True, Helpmessage="Provide the Task Sequence Package ID")]    [string]$TSID)Begin {    #Checks if the user is in the administrator group. Warns and stops if the user is not.    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "You are not running this as local administrator. Run it again in an elevated prompt." ; break
    }    try {        Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)    }    catch [System.UnauthorizedAccessException] {
        Write-Warning -Message "Access denied" ; break
    }    catch [System.Exception] {        Write-Warning "Unable to load the Configuration Manager Powershell module from $env:SMS_ADMIN_UI_PATH" ; break    }}Process {    $SiteCode = Get-PSDrive -PSProvider CMSITE    Set-Location -Path "$($SiteCode.Name):\"    $TS = Get-CMTaskSequence -TaskSequencePackageId $TSID}End {    Set-Location -Path $env:SystemDrive    $TS.references | select Package | Export-CSV "dependencies.csv"}
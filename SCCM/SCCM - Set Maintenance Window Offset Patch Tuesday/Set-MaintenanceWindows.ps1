<#


    PARAM(
    [int]$OffSetWeeks,
    [int]$OffSetDays,
    [Parameter(Mandatory=$True)]
    [int]$AddStartHour,
    [Parameter(Mandatory=$True)]
    [int]$AddStartMinutes,
    [Parameter(Mandatory=$True)]
    [int]$AddEndHour,
    [Parameter(Mandatory=$True)]
    [int]$AddEndMinutes,
    [Parameter(Mandatory=$True,Helpmessage="Provide the ID of the collection")]
    [string]$CollID
    )  

#region Initialize

#Load SCCM Powershell module
Try {
    Write-Warning -Message "Access denied" ; break
}

#endregion

#region functions

#Set Patch Tuesday for a Month 
Function Get-PatchTuesday ($Month,$Year)  
 { 
    $FindNthDay=2 #Aka Second occurence 
    $WeekDay='Tuesday' 
    $todayM=($Month).ToString()
    $todayY=($Year).ToString()
    $StrtMonth=$todayM+'/1/'+$todayY 
    [datetime]$StrtMonth=$todayM+'/1/'+$todayY 
    while ($StrtMonth.DayofWeek -ine $WeekDay ) { $StrtMonth=$StrtMonth.AddDays(1) } 
    $PatchDay=$StrtMonth.AddDays(7*($FindNthDay-1)) 
    return $PatchDay
    Write-Log -Message "Patch Tuesday this month is $PatchDay" -Severity 1 -Component "Set Patch Tuesday"
    Write-Output "Patch Tuesday this month is $PatchDay"
 }  
 
#Remove all existing Maintenance Windows for a Collection 
Function Remove-MaintenanceWindow {
    PARAM(
    [string]$CollID
    )
    Get-CMMaintenanceWindow -CollectionId $CollID | ? {$_.StartTime -lt (Get-Date)} | ForEach-Object { 
    
    Try {
        Remove-CMMaintenanceWindow -CollectionID $CollID -Name $_.Name -Force -ErrorAction Stop
        $Coll=Get-CMDeviceCollection -CollectionId $CollID -ErrorAction Stop
        Write-Log -Message "Removing $($OldMW.Name) from collection $MWCollection" -Severity 1 -Component "Remove Maintenance Window"
        Write-Output "Removing $($OldMW.Name) from collection $MWCollection"
    }
    Catch {
        Write-Log -Message "Unable to remove $($OldMW.Name) from collection $MWCollection" -Severity 3 -Component "Remove Maintenance Window"
        Write-Warning "Unable to remove $($OldMW.Name) from collection $MWCollection. Error: $_.Exception.Message"   
    } 
}
} 

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
    "<![LOG[$Message]LOG]!><time=$([char]34)$Date$($TimeZoneBias.bias)$([char]34) date=$([char]34)$date2$([char]34) component=$([char]34)$Component$([char]34) context=$([char]34)$([char]34) type=$([char]34)$Severity$([char]34) thread=$([char]34)$([char]34) file=$([char]34)$([char]34)>"| Out-File -FilePath "$Logpath\Set-MaintenanceWindows.log" -Append -NoClobber -Encoding default
    $SiteCode = Get-PSDrive -PSProvider CMSITE
    Set-Location -Path "$($SiteCode.Name):\"
}

#endregion

#region Parameters

$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-Location -Path "$($SiteCode.Name):\"
$GetCollection = Get-CMDeviceCollection -ID $CollID
$MWCollection = $GetCollection.Name
$ErrorMessage = $_.Exception.Message

#endregion

$MonthArray = New-Object System.Globalization.DateTimeFormatInfo 
$MonthNames = $MonthArray.MonthNames 
$PatchMonth = (Get-Date).Month
$PatchYear = (Get-Date).Year
$LastDay = ((Get-Date).Addmonths(1)).Adddays(-(Get-Date ((Get-Date).Addmonths(1)) -Format dd)).Day

#Set Patch Tuesday for each Month 
$PatchDay = Get-PatchTuesday $PatchMonth $PatchYear
    
#Set Date value For The Next Month if Required
if ((($PatchDay.day + $OffSetDays + ($OffSetWeeks*7)) -gt (Get-date).Day) -ne $True) {
    $PatchMonth = ((Get-Date).AddMonths(1)).Month
    if($PatchMonth -eq '1'){
    $PatchYear=((Get-Date).AddYears(1)).Year
    }
$PatchDay = Get-PatchTuesday $PatchMonth $PatchYear
}
                 
#Set Maintenance Window Naming Convention (Months array starting from 0 hence the -1) 
$NewMWName =  "MW-SUM_"+$MonthNames[$PatchMonth-1]+"_OffsetWeeks"+$OffSetWeeks+"_OffSetDays"+$OffSetDays
$OldMW = Get-CMMaintenanceWindow -CollectionId $CollID
 
#Set Device Collection Maintenace interval  
$StartTime=$PatchDay.AddDays($OffSetDays).AddHours($AddStartHour).AddMinutes($AddStartMinutes)
$EndTime=$StartTime.Addhours(0).AddHours($AddEndHour).AddMinutes($AddEndMinutes)

#Delete Previous MW 

Try {
    If(($OldMW).StartTime -lt (Get-Date) -eq $True){    
        Set-Location -Path "$($SiteCode.Name):\" -ErrorAction Stop
        Remove-MaintenanceWindow $CollID -ErrorAction Stop                  
    }
}
Catch {
    Write-Log -Message "$_.Exception.Message" -Severity 3 -Component "Remove Maintenance Window"
}
  
Try {
    #Create The Schedule Token  
    $Schedule = New-CMSchedule -Nonrecurring -Start $StartTime.AddDays($OffSetWeeks*7) -End $EndTime.AddDays($OffSetWeeks*7) -ErrorAction Stop
    
    New-CMMaintenanceWindow -CollectionID $CollID -Schedule $Schedule -Name $NewMWName -ApplyTo SoftwareUpdatesOnly -ErrorAction Stop
    Write-Log -Message "Created Maintenance Window $NewMWName for Collection $MWCollection" -Severity 1 -Component "New Maintenance Window"
    Write-Output "Created Maintenance Window $NewMWName for Collection $MWCollection" 
}
Catch {
    Write-Warning "$_.Exception.Message"
    Write-Log -Message "$_.Exception.Message" -Severity 3 -Component "Create new Maintenance Window"
}

   
If([int]$OffSetDays -gt 7){
    $OffSetWeeks = "{0:N0}" -f ([int]$OffSetDays / 7)   
    $OffSetDays = [int]$OffSetDays - ([int]$OffSetWeeks*7)                
}
Else{
    $OffSetWeeks = 0
}
Set-Location $PSScriptRoot
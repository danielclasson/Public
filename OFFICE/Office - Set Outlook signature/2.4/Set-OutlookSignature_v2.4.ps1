<#.SYNOPSISScript to set Outlook 2010/2013 e-mail signature using Active Directory information.DESCRIPTIONThis script will set the Outlook 2010/2013 e-mail signature on the local client using Active Directory information. The template is created with a Word document, where images can be inserted and AD values can be provided.Author: Daniel ClassonVersion 2.4.DISCLAIMERAll scripts and other powershell references are offered AS IS with no warranty.These script and functions are tested in my environment and it is recommended that you test these scripts in a test environment before using in your production environment.#>


#Custom variables
$SignatureName = '' #insert the company name (no spaces)
$SigSource = "" #Change if desired for signature templates
$SignatureVersion = "1.0" #Change this if you have updated the signature. If you do not change it, the script will quit after checking for the version already on the machine
$ForceSignature = $False
 
#Environment variables
$AppData=(Get-Item env:appdata).value
$SigPath = '\Microsoft\Signatures' #This is different depending on system language. I.e Swedish is 'Microsoft\Signaturer'
$LocalSignaturePath = $AppData+$SigPath
$RemoteSignaturePathFull = $SigSource

#Copy version file
If (-not(Test-Path -Path "C:\ProgramData\Microsoft\OFFICE\Signature\Version\$SignatureVersion")) {
    New-Item -Path "C:\ProgramData\Microsoft\OFFICE\Signature\Version\$SignatureVersion" -ItemType Directory
}
Elseif (Test-Path -Path "C:\ProgramData\Microsoft\OFFICE\Signature\Version\$SignatureVersion") {
    Write-Output "Latest signature already exists"
    break
}

#Check signature path (needs to be created if a signature has never been created for the profile
If (-not(Test-Path -path $LocalSignaturePath)) {
	New-Item $LocalSignaturePath -Type Directory
}

#Get Active Directory information for current user
$UserName = $env:username
$Filter = "(&(objectCategory=User)(samAccountName=$UserName))"
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.Filter = $Filter
$ADUserPath = $Searcher.FindOne()
$ADUser = $ADUserPath.GetDirectoryEntry()
$ADDisplayName = $ADUser.DisplayName
$ADEmailAddress = $ADUser.mail
$ADTitle = $ADUser.title
$ADDescription = $ADUser.description
$ADTelePhoneNumber = $ADUser.TelephoneNumber
$ADMobile = $ADUser.mobile
$ADStreetAddress = $ADUser.streetaddress
$ADCity = $ADUser.l
$ADDepartment = $ADUser.department
$ADCustomAttribute1 = $ADUser.extensionAttribute1
$ADCustomAttribute2 = $ADUser.extensionAttribute2
$ADCustomAttribute3 = $ADUser.extensionAttribute3
$ADModify = $ADUser.whenChanged

#Copy signature templates from source to local Signature-folder
Write-Output "Copying Signatures"
Copy-Item "$Sigsource" $LocalSignaturePath -Recurse -Force
$ReplaceAll = 2
$FindContinue = 1
$MatchCase = $False
$MatchWholeWord = $True
$MatchWildcards = $False
$MatchSoundsLike = $False
$MatchAllWordForms = $False
$Forward = $True
$Wrap = $FindContinue
$Format = $False
	
#Insert variables from Active Directory to rtf signature-file
$MSWord = New-Object -ComObject word.application
$fullPath = $LocalSignaturePath+'\'+$SignatureName+'.docx'
$MSWord.Documents.Open($fullPath)
	
#ExtensionAttribute2
Write-Host "Value is $ADCustomAttribute2"
$FindText = "ExtensionAttribute2"
$ReplaceText = $ADCustomAttribute2.ToString()
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)

#ExtensionAttribute3
$FindText = "ExtensionAttribute3"
$ReplaceText = $ADCustomAttribute3.ToString()
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)

#User Name $ Designation 
$FindText = "DisplayName" 
$Designation = $ADCustomAttribute1.ToString()
If ($Designation -ne '') { 
	$Name = $ADDisplayName.ToString()
	$ReplaceText = $Name+', '+$Designation
}
Else {
	$ReplaceText = $ADDisplayName.ToString() 
}
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)	

#Department
$FindText = "Department"
$ReplaceText = $ADDepartment.ToString()
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)

#Title		
$FindText = "Title"
$ReplaceText = $ADTitle.ToString()
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	
#Description
If ($ADDescription -ne '') { 
   	$FindText = "Description"
   	$ReplaceText = $ADDescription.ToString()
}
Else {
	$FindText = " | Description "
   	$ReplaceText = "".ToString()
}
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
#$LogInfo += $NL+'Description: '+$ReplaceText
   	
#Street Address
If ($ADStreetAddress -ne '') { 
       $FindText = "StreetAddress"
    $ReplaceText = $ADStreetAddress.ToString()
}
Else {
    $FindText = "StreetAddress"
    $ReplaceText = $DefaultAddress
    }
	$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)

#City
If ($ADCity -ne '') { 
    $FindText = "City"
       $ReplaceText = $ADCity.ToString()
   }
Else {
    $FindText = "City"
    $ReplaceText = $DefaultCity 
   }
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	
#Telephone
If ($ADTelephoneNumber -ne "") { 
	$FindText = "TelephoneNumber"
	$ReplaceText = $ADTelephoneNumber.ToString()
   }
Else {
	$FindText = "TelephoneNumber"
    $ReplaceText = $DefaultTelephone
	}
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	
#Mobile
If ($ADMobile -ne "") { 
	$FindText = "MobileNumber"
	$ReplaceText = $ADMobile.ToString()
   }
Else {
	$FindText = "| Mob MobileNumber "
    $ReplaceText = "".ToString()
	}
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)

#Save new message signature 
Write-Output "Saving signatures"
#Save HTML
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML");
$path = $LocalSignaturePath+'\'+$SignatureName+".htm"
$MSWord.ActiveDocument.saveas([ref]$path, [ref]$saveFormat)
    
#Save RTF 
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatRTF");
$path = $LocalSignaturePath+'\'+$SignatureName+".rtf"
$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$saveFormat)
	
#Save TXT    
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatText");
$path = $LocalSignaturePath+'\'+$SignatureName+".txt"
$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$SaveFormat)
$MSWord.ActiveDocument.Close()
$MSWord.Quit()
	

#Office 2010
If ((Test-Path HKCU:'\Software\Microsoft\Office\14.0') -and ($ForceSignature -eq $false))  {
    Write-Output "Setting Office 2010 signature as available"
    Try {
        Remove-ItemProperty -Path HKCU:\Software\Microsoft\Office\14.0\Common\MailSettings -Name ReplySignature -Force -ErrorAction SilentlyContinue -Verbose
        Remove-ItemProperty -Path HKCU:\Software\Microsoft\Office\14.0\Common\MailSettings -Name NewSignature -Force -ErrorAction SilentlyContinue -Verbose
    }
    Catch {
    }
    $MSWord = New-Object -comobject word.application
    $EmailOptions = $MSWord.EmailOptions
    $EmailSignature = $EmailOptions.EmailSignature
    $EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
    $EmailSignature.NewMessageSignature=$SignatureName
    $EmailSignature.ReplyMessageSignature=$SignatureName
}

If ((Test-Path HKCU:'\Software\Microsoft\Office\14.0') -and ($ForceSignature -eq $True)) {
    Write-Output "Setting signature for Office 2010 as forced"
    New-ItemProperty HKCU:'\Software\Microsoft\Office\14.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force
    New-ItemProperty HKCU:'\Software\Microsoft\Office\14.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force
}

#Office 2013 signature

If ((Test-Path HKCU:'\Software\Microsoft\Office\15.0') -and ($ForceSignature -eq $False)) {
    Write-Output "Setting Office 2013 signature as available"
    Try {
        Remove-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Common\MailSettings -Name ReplySignature -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -Path HKCU:\Software\Microsoft\Office\15.0\Common\MailSettings -Name NewSignature -Force -ErrorAction SilentlyContinue
    }
    Catch {
    }
    $MSWord = New-Object -comobject word.application
    $EmailOptions = $MSWord.EmailOptions
    $EmailSignature = $EmailOptions.EmailSignature
    $EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
    #$EmailSignature.NewMessageSignature=$SignatureName
    #$EmailSignature.ReplyMessageSignature=$SignatureName
}

If ((Test-Path HKCU:'\Software\Microsoft\Office\15.0') -and ($ForceSignature -eq $true)) {
    Write-Output "Setting signature for Office 2013 as forced"
    If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings') { } 
    Else { 
        New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'NewSignature' -Value $SignatureName -PropertyType 'String' -Force 
    } 
    If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings') { } 
    Else { 
        New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'ReplySignature' -Value $SignatureName -PropertyType 'String' -Force
    } 
}
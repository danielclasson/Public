###########################################################################"
#
# NAME: Set-OutlookSignature.ps1
#
# AUTHOR: Jan Egil Ring
# Modifications by Darren Kattan
# Further Modifications by Jamie McKillop - http://jamiemckillop.wordpress.com/
#
# COMMENT: Script to create an Outlook signature based on user information from Active Directory.
#          Adjust the variables in the "Custom variables"-section
#          Create an Outlook-signature from Microsoft Word (logo, fonts etc) and copy this signature to \\domain\NETLOGON\Signatures\$SignatureName\$SignatureName.docx
#		   This script supports the following keywords:
#		   	DisplayName
#			Title
#			Email
#                      Description
#                      TelephoneNumber
#                      facsimileTelephoneNumber
#                      mobile
#                      streetaddress
#                      City
#                     postofficebox
#                     extensionAttribute1
#   See the following blog-posts for more information: 
# http://blog.crayon.no/blogs/janegil/archive/2010/01/09/outlook-signature-based-on-user-information-from-active-directory.aspx
# http://gallery.technet.microsoft.com/office/6f7eee4b-1f42-499e-ae59-1aceb26100de
# http://www.experts-exchange.com/Software/Server_Software/Email_Servers/Exchange/Q_28035335.html
# http://jamiemckillop.wordpress.com/category/powershell/
# http://www.immense.net/deploying-unified-email-signature-template-outlook/
#
#          Tested on Office 2003,2007,2010 and 2013
#
# You have a royalty-free right to use, modify, reproduce, and
# distribute this script file in any way you find useful, provided that
# you agree that the creator, owner above has no warranty, obligations,
# or liability for such use.
#
# VERSION HISTORY:
# 1.0 09.01.2010 - Initial release
# 1.1 11.09.2010 - Modified by Darren Kattan
#	- Removed bookmarks. Now uses simple find and replace for DisplayName, Title, and Email.
#	- Email address is generated as a link
#	- Signature is generated from a single .docx file
#	- Removed version numbers for script to run. Script runs at boot up when it sees a change in the "Date Modified" property of your signature template.
# 1.11 11.15.2010 - Revised by Darren Kattan
#   - Fixed glitch with text signatures
# 1.2 07.06.2012 - Revised by Jamie McKillop
#   - Modified script so that Force Signature settings are set on first run of script
#	- Added variables to allow setting of default signature on creation of signature but not force the signature on each script run
#	- Used variables defined in script for $ForceSignatureNew and $ForceSignatureReplyForward instead of pulling values from the registry
# 1.3 01.13.2014 - Revised by Dominic Whyle
#   - Modified script so Include logging
#	- Added variables to allow setting of default signature address, telephone, fax and city
#	- Modifed script to replace unused fields (Mobile or Description)
#	- Added force script to run in x86mode for x86 versions of office
# 1.4 01.16.2014 - Revised by Dominic Whyle
#   - Added variable for AD account whenChanged to allow automatic updating of signauture when AD account changes 
#
###########################################################################"

#Run Script in x86 Mode
if ($env:Processor_Architecture -ne "x86")
{ write-warning 'Launching x86 PowerShell'
&"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" -noninteractive -noprofile -file $myinvocation.Mycommand.path -executionpolicy bypass
exit
}
"Always running in 32bit PowerShell at this point."
$env:Processor_Architecture
[IntPtr]::Size

#Custom variables
#$ModulePath = '\\'+$DomainName+'\Netlogon\Signatures' #insert log module path
$SignatureName = 'Ale Kommun' #insert the company name (no spaces) - could be signature name if more than one sig needed
$DomainName = 'ale.local' #insert the domain name
$SigSource = "C:\test\$SignatureName" #Change if desired for signature templates
$ForceSignatureNew = '1' #When the signature is forced it sets the default signature for new messages each time the script runs. 0 = no force, 1 = force
$ForceSignatureReplyForward = '1' #When the signature is forced it sets the default signature for reply/forward messages each time the script runs. 0 = no force, 1 = force
$SetSignatureNew = '1' #Determines wheter to set the signature as the default for new messages on first run. This is overridden if $ForceSignatureNew = 1. 0 = don't set, 1 = set
$SetSignatureReplyForward = '1' #Determines wheter to set the signature as the default for reply/forward messages on first run. This is overridden if $ForceSignatureReplyForward = 1. 0 = don't set, 1 = set
$DefaultAddress = 'My Company Address #insert default address'
$DefaultPOBox = 'PO Box 666' #insert default PO Box
$DefaultCity = 'Somewhere' #insert default city
$DefaultTelephone = '123456' #insert default phone number
$DefaultFax = '123456' #insert default fax number

#Modules
#New-PSDrive -Name O -PSProvider FileSystem -Root $ModulePath #Map the modules folder for PS to the O: drive
#. O:\LogData.ps1 #Add logging module

#Log data to the $LogInfo variable
$LogInfo = '' #clear the log variable
$NL = [Environment]::NewLine #new line variable for ease of use
$Date = Get-Date
$LogInfo = 'Signature Script - '+$Date
$LogInfo += $NL+'Signature Source: '+$SigSource
 
#Environment variables
$AppData=(Get-Item env:appdata).value
$SigPath = '\Microsoft\Signaturer'
$LocalSignaturePath = $AppData+$SigPath
$RemoteSignaturePathFull = $SigSource+'\'+$SignatureName+'.docx'

#Check signature path (needs to be created if a signature has never been created for the profile
if (!(Test-Path -path $LocalSignaturePath)) {
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
$ADFax = $ADUser.facsimileTelephoneNumber
$ADMobile = $ADUser.mobile
$ADStreetAddress = $ADUser.streetaddress
$ADCity = $ADUser.l
$ADPOBox = $ADUser.postofficebox
$ADCustomAttribute1 = $ADUser.extensionAttribute1
$ADModify = $ADUser.whenChanged

#Setting registry information for the current user
$CompanyRegPath = "HKCU:\Software\"+$DomainName
$SignatureRegPath = $CompanyRegPath+'\'+$SignatureName

if (Test-Path $SignatureRegPath) { Echo 'Company Registry Exists'
}
else {
	New-Item -path "HKCU:\Software" -name $DomainName
    New-Item -path $CompanyRegPath -name $SignatureName
}

if (Test-Path $SignatureRegPath'\Outlook Signature Settings') { Echo 'Outlook Settings Exist'
}
else {
	New-Item -path $SignatureRegPath -name "Outlook Signature Settings"
}

$SigVersion = (gci $RemoteSignaturePathFull).LastWriteTime  #When was the last time the signature was written
$LogInfo += $NL+'Master Signature Version: '+$SigVersion
$SignatureVersion = (Get-ItemProperty $SignatureRegPath'\Outlook Signature Settings').SignatureVersion
$LogInfo += $NL+'Local Signature Version: '+$SignatureVersion
Set-ItemProperty $SignatureRegPath'\Outlook Signature Settings' -name SignatureSourceFiles -Value $SigSource
$SignatureSourceFiles = (Get-ItemProperty $SignatureRegPath'\Outlook Signature Settings').SignatureSourceFiles
Set-ItemProperty $SignatureRegPath'\Outlook Signature Settings' -name UserAccountModifyDate -Value $ADModify.ToString()
$UserModify = (Get-ItemProperty $SignatureRegPath'\Outlook Signature Settings').UserAccountModifyDate

#Copying signature sourcefiles and creating signature if signature-version are different from local version
if (($SignatureVersion -eq $SigVersion) -or ($UserModify -eq $ADModify))
    {
	Echo 'Primary signature is up to date'
	$LogInfo += $NL+'Signature up to date' 
}
else
{

	Echo 'Running Main Script'

	#Copy signature templates from domain to local Signature-folder
	Copy-Item "$SignatureSourceFiles\*" $LocalSignaturePath -Recurse -Force
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
	
	#User Name $ Designation 
	$FindText = "DisplayName" 
	$Designation = $ADCustomAttribute1.ToString() #designations in Exchange custom attribute 1
	If ($Designation -ne '') { 
		$Name = $ADDisplayName.ToString()
		$ReplaceText = $Name+', '+$Designation
	}
	Else {
		$ReplaceText = $ADDisplayName.ToString() 
	}
	$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	$LogInfo += $NL+'Username: '+$ReplaceText	

	#Title		
	$FindText = "Title"
	$ReplaceText = $ADTitle.ToString()
	$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	$LogInfo += $NL+'Title: '+$ReplaceText
	
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
	$LogInfo += $NL+'Description: '+$ReplaceText
   	
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
	$LogInfo += $NL+'Street Address: '+$ReplaceText
		
	#PostofficeBox
	If ($ADPOBox -ne '') { 
        $FindText = "PostofficeBox"
        $ReplaceText = $ADPOBox.ToString()
    }
    Else {
	    $FindText = "PostofficeBox"
	    $ReplaceText = $DefaultPOBox 
    }
	$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	$LogInfo += $NL+'PostofficeBox: '+$ReplaceText

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
	$LogInfo += $NL+'City: '+$ReplaceText
	
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
	$LogInfo += $NL+'Telephone: '+$ReplaceText
	
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
    $LogInfo += $NL+'MobileNumber: '+$ReplaceText

	#Fax
    If ($ADFaxNumber -ne '') { 
    	$FindText = "FaxNumber"
        $ReplaceText = $ADFaxNumber 
    }
    Else {
	    $FindText = "FaxNumber"
        $ReplaceText = $DefaultFax 
    }
	$ReplaceText = $ADFax.ToString()
	$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,	$MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,	$Format, $ReplaceText, $ReplaceAll	)
	$LogInfo += $NL+'Fax Number: '+$ReplaceText

	#$MSWord.Selection.Find.Execute("Email")
	#$MSWord.ActiveDocument.Hyperlinks.Add($MSWord.Selection.Range, "mailto:"+$ADEmailAddress.ToString(), $missing, $missing, $ADEmailAddress.ToString())

    #Save new message signature 
    Echo 'Saving Signatures'
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
	
    #Save new message signature 
	#Echo 'Saving Signatures'
	#Save HTML
	#$SaveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML"); 
	#[ref]$BrowserLevel = "microsoft.office.interop.word.WdBrowserLevel" -as [type] 
	#$New.WebOptions.OrganizeInFolder = $true 
	#$New.WebOptions.UseLongFileNames = $true 
	#$New.WebOptions.BrowserLevel = $BrowserLevel::wdBrowserLevelMicrosoftInternetExplorer6 
	#$Path = $LocalSignaturePath+'\'+$CompanyName+'.htm' 
	#If ((Test-Path $Path) -eq '1') { (Get-Item $Path).attributes = 'Archive' }
	#$New.SaveAs([ref]$Path, [ref]$saveFormat) 
 
	#Set signature for new mesages if enabled
	#if ($SetSignatureNew -eq '1') {
	#	#Set company signature as default for New messages
	#	$MSWord = New-Object -ComObject word.application
	#	$EmailOptions = $MSWord.EmailOptions
	#	$EmailSignature = $EmailOptions.EmailSignature
	#	$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
	#	$EmailSignature.NewMessageSignature=$SignatureName
	#	$MSWord.Quit()
	#}
	
	#Set signature for reply/forward messages if enabled
	#if ($SetSignatureReplyForward -eq '1') {
	#	#Set company signature as default for Reply/Forward messages
	#	$MSWord = New-Object -ComObject word.application
	#	$EmailOptions = $MSWord.EmailOptions
	#	$EmailSignature = $EmailOptions.EmailSignature
	#	$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
	#	$EmailSignature.ReplyMessageSignature=$SignatureName
	#	$MSWord.Quit()
	#}
}

#Stamp registry-values for Outlook Signature Settings if they doesn`t match the initial script variables. Note that these will apply after the second script run when changes are made in the "Custom variables"-section.
if ($ForcedSignatureNew -eq $ForceSignatureNew){
}
else {
	Set-ItemProperty $SignatureRegPath'\Outlook Signature Settings' -name ForcedSignatureNew -Value $ForceSignatureNew
}

if ($ForcedSignatureReplyForward -eq $ForceSignatureReplyForward){
}
else {
	Set-ItemProperty $SignatureRegPath'\Outlook Signature Settings' -name ForcedSignatureReplyForward -Value $ForceSignatureReplyForward
}

if ($SignatureVersion -eq $SigVersion){
}
else {
	Set-ItemProperty $SignatureRegPath'\Outlook Signature Settings' -name SignatureVersion -Value $SigVersion
}

#Forcing signature for new messages if enabled
#if ($ForceSignatureNew -eq '1') {
#	#Set company signature as default for New messages
#	$MSWord = New-Object -ComObject word.application
#	$EmailOptions = $MSWord.EmailOptions
#	$EmailSignature = $EmailOptions.EmailSignature
#	$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
#	$EmailSignature.NewMessageSignature="$SignatureName"
#	$MSWord.Quit()
#}

#Forcing signature for reply/forward messages if enabled
#Office 2013 
If ((Test-Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings') -eq 1) {  
    If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
    Else { New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'NewSignature' -Value 'ale' -PropertyType 'String' } 
    If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
    Else { New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'ReplySignature' -Value 'ale' -PropertyType 'String' } 
    $LogInfo += $NL+'Office 2013 Registry values added (HKCU:\Software\Microsoft\Office\15.0\Common\MailSettings' 
} 

#if ($ForceSignatureReplyForward -eq '1') {
#	#Set company signature as default for Reply/Forward messages
#	$MSWord = New-Object -ComObject word.application
#	$EmailOptions = $MSWord.EmailOptions
#	$EmailSignature = $EmailOptions.EmailSignature
#	$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
#	$EmailSignature.ReplyMessageSignature=$SignatureName
#	$MSWord.Quit()
#}
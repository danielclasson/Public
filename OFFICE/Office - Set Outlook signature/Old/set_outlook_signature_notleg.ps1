<# ------------------------------------------------------------------------------------------ 
Signature Creation Script - Lock signatures after creation 
 
Created: 11/30/2012  
Created by: Drew Heath 
 
Detail:  Script copies signature templates from \\<domain>\netlogon\Signatures  
to the local %appdata%\Microsoft\Signatures folder if they do not exist then  
creates custom signatures for the user with data from AD.  If there are  
already signiture files in place, the script compares the modified date from 
netlogon to the local files and will re-run the script if netlogon is newer.  
 
Credit: Portions of this script have been taken from the following link by Jan Egil Ring: 
http://gallery.technet.microsoft.com/office/6f7eee4b-1f42-499e-ae59-1aceb26100de 
------------------------------------------------------------------------------------------#> 
 
#Paths and variables 
#$ModulePath = '' #insert log module path 
$CompanyName = 'ale' #insert the company name (no spaces) 
$DomainName = 'ale.local' #insert the domain name 
$SigSource =  'c:\test' #Change if desired for signature templates  
$DefaultTelephone = '' #insert default phone number 
$DefaultFax = '' #insert default fax number 
 
#Modules 
#New-PSDrive -Name P -PSProvider FileSystem -Root $ModulePath #Map the modules folder for PS to the P: drive 
#. P:\LogData.ps1 #Add logging module 
 
#Log data to the $LogInfo variable 
#$LogInfo = '' #clear the log variable 
#$NL = [Environment]::NewLine #new line variable for ease of use 
#$Date = Get-Date 
#$LogInfo = 'Signature Script - '+$Date 
#$LogInfo += $NL+'Signature Source: '+$SigSource 
  
#Environment variables  
$AppData=(Get-Item env:appdata).value  
$SigPath = '\Microsoft\Signaturer'  
$LocalSignaturePath = $AppData+$SigPath  
$RemoteSignaturePathFull = $SigSource+'\'+$CompanyName+'.docx' 
 
#Check signature path (needs to be created if a signature has never been created for the profile 
If ((Test-Path $LocalSignaturePath) -eq 0) { New-Item -Path $AppData'\Microsoft' -Name "Signaturer" -Type Directory }  
     
#Get Active Directory information for current user (designations are set in Custom Attribute 1 in Exchange. 
$UserName = $env:username  
$Filter = "(&(objectCategory=User)(samAccountName=$UserName))" 
$Searcher     = New-Object System.DirectoryServices.DirectorySearcher  
$Searcher.Filter = $Filter  
$ADUserPath = $Searcher.FindOne()  
$ADUser = $ADUserPath.GetDirectoryEntry()  
$ADDisplayName = $ADUser.displayName  
$ADEmailAddress = $ADUser.mail  
$ADTitle = $ADUser.title  
$ADTelePhoneNumber = $ADUser.telephoneNumber 
$ADFaxNumber = $ADUser.facsimileTelephoneNumber 
$ADMobileNumber = $ADUser.mobile 
$ADStreetAddress = $ADUser.streetAddress 
$ADCity = $ADUser.l 
$ADState = $ADUser.st 
$ADZip = $ADUser.postalCode 
$ADCustomAttribute1 = $ADUser.extensionAttribute1 
 
#Set default telephone/fax if blank 
If ($ADTelePhoneNumber -eq '') { $ADTelePhoneNumber = $DefaultTelephone } 
If ($ADFaxNumber -eq '') { $ADFaxNumber = $DefaultFax } 
 
#Setup registry information for the current user  
$CompanyRegPath = "HKCU:\Software\"+$CompanyName  
If (Test-Path $CompanyRegPath) { Echo 'Company Registry Exists' }  
Else {New-Item -path 'HKCU:\Software' -name $CompanyName}  
  
If (Test-Path $CompanyRegPath'\Outlook Signature Settings') { Echo 'Outlook Settings Exist' }  
Else {New-Item -path $CompanyRegPath -name 'Outlook Signature Settings'}  
  
$SigVersion = (GCI $RemoteSignaturePathFull).LastWriteTime #When was the last time the signature was written 
$LogInfo += $NL+'Master Signature Version: '+$SigVersion 
$ForcedSignatureNew = (Get-ItemProperty $CompanyRegPath'\Outlook Signature Settings').ForcedSignatureNew  
$ForcedSignatureReplyForward = (Get-ItemProperty $CompanyRegPath'\Outlook Signature Settings').ForcedSignatureReplyForward  
$SignatureVersion = (Get-ItemProperty $CompanyRegPath'\Outlook Signature Settings').SignatureVersion 
$LogInfo += $NL+'Local Signature Version: '+$SignatureVersion 
Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name SignatureSourceFiles -Value $SigSource  
$SignatureSourceFiles = (Get-ItemProperty $CompanyRegPath'\Outlook Signature Settings').SignatureSourceFiles  
  
#Copying signature sourcefiles and creating signature if signature-version are different from local version  
$ExistingSignature = $LocalSignaturePath+'\'+$CompanyName+'.htm' 
If (($SignatureVersion -eq $SigVersion) -and ((Test-Path $ExistingSignature) -eq 1)) {  
    Echo 'Primary signature is up to date' 
    $LogInfo += $NL+'Signature up to date'  
}  
Else {  
    Echo 'Running Main Script' 
     
    #The signature is either new or non-existant, force an update of the association signatures 
    $NoAssocSignature = $true 
    #Copy signature templates from domain to local Signature-folder  
    Copy-Item "$SignatureSourceFiles\*" $LocalSignaturePath -Recurse -Force  
      
    #Create Word objects for the new message and reply message signatures 
    $MSWordNew = New-Object -com word.application  
    $MSWordReply = New-Object -com word.application 
     
    #Check for a mobile number in AD, open temple with mobile number if present 
    $Mobile = $ADMobileNumber.ToString() 
    If ($Mobile -eq '') { $FullPathNM = $LocalSignaturePath+'\'+$CompanyName+'.docx' } 
    Else { $FullPathNM = $LocalSignaturePath+'\'+$CompanyName+'Mobile.docx' } 
    If ($Mobile -eq '') { $FullPathRM = $LocalSignaturePath+'\'+$CompanyName+'Reply.docx' } 
    Else { $FullPathRM = $LocalSignaturePath+'\'+$CompanyName+'ReplyMobile.docx' } 
     
    #Open the documents with the corresponding Word objects 
    $New     = $MSWordNew.Documents.Open($FullPathNM) 
    $Reply     = $MSWordReply.Documents.Open($FullPathRM) 
      
    #Replace the bookmarks in the opened documents with AD information  
    #User Name $ Designation  
    $Bookmark = "DisplayName"  
    $Designation = $ADCustomAttribute1.ToString() #designations in Exchange custom attribute 1 
    If ($Designation -ne '') {  
        $Name = $ADDisplayName.ToString() 
        $ReplaceText = $Name+', '+$Designation 
    } 
    Else { 
        $ReplaceText = $ADDisplayName.ToString()  
    } 
    $LogInfo += $NL+'Username: '+$ReplaceText 
    $RangeNew = $New.Bookmarks.Item($Bookmark).Range 
    $RangeNew.Text = $ReplaceText 
    $New.Bookmarks.Add($Bookmark,$RangeNew) 
    $RangeReply = $Reply.Bookmarks.Item($Bookmark).Range 
    $RangeReply.Text = $ReplaceText 
    $Reply.Bookmarks.Add($Bookmark,$RangeReply)  
      
    #Title 
    $Bookmark = "Title" 
    $ReplaceText = $ADTitle.ToString()  
    $LogInfo += $NL+'Title: '+$ReplaceText 
    $RangeNew = $New.Bookmarks.Item($Bookmark).Range 
    $RangeNew.Text = $ReplaceText 
    $New.Bookmarks.Add($Bookmark,$RangeNew) 
    $RangeReply = $Reply.Bookmarks.Item($Bookmark).Range 
    $RangeReply.Text = $ReplaceText 
    $Reply.Bookmarks.Add($Bookmark,$RangeReply)  
 
    #Street Address 
    $Bookmark = "StreetAddress"  
    $ReplaceText = $ADStreetAddress.ToString() 
    $LogInfo += $NL+'Street Address: '+$ReplaceText 
    $RangeNew = $New.Bookmarks.Item($Bookmark).Range 
    $RangeNew.Text = $ReplaceText 
    $New.Bookmarks.Add($Bookmark,$RangeNew) 
    $RangeReply = $Reply.Bookmarks.Item($Bookmark).Range 
    $RangeReply.Text = $ReplaceText 
    $Reply.Bookmarks.Add($Bookmark,$RangeReply)  
 
    #City 
    $Bookmark = "City"  
    $ReplaceText = $ADCity.ToString() 
    $LogInfo += $NL+'City: '+$ReplaceText 
    $RangeNew = $New.Bookmarks.Item($Bookmark).Range 
    $RangeNew.Text = $ReplaceText 
    $New.Bookmarks.Add($Bookmark,$RangeNew) 
    $RangeReply = $Reply.Bookmarks.Item($Bookmark).Range 
    $RangeReply.Text = $ReplaceText 
    $Reply.Bookmarks.Add($Bookmark,$RangeReply)  
     
    #State 
    $Bookmark = "State"  
    $ReplaceText = $ADState.ToString() 
    $LogInfo += $NL+'State: '+$ReplaceText 
    $RangeNew = $New.Bookmarks.Item($Bookmark).Range 
    $RangeNew.Text = $ReplaceText 
    $New.Bookmarks.Add($Bookmark,$RangeNew) 
    $RangeReply = $Reply.Bookmarks.Item($Bookmark).Range 
    $RangeReply.Text = $ReplaceText 
    $Reply.Bookmarks.Add($Bookmark,$RangeReply)  
     
    #ZIP 
    $Bookmark = "ZIP"  
    $ReplaceText = $ADZip.ToString() 
    $LogInfo += $NL+'Zip Code: '+$ReplaceText 
    $RangeNew = $New.Bookmarks.Item($Bookmark).Range 
    $RangeNew.Text = $ReplaceText 
    $New.Bookmarks.Add($Bookmark,$RangeNew) 
    $RangeReply = $Reply.Bookmarks.Item($Bookmark).Range 
    $RangeReply.Text = $ReplaceText 
    $Reply.Bookmarks.Add($Bookmark,$RangeReply)  
     
    #Telephone 
    $Bookmark = "Telephone"  
    $ReplaceText = $ADTelePhoneNumber.ToString() 
    $LogInfo += $NL+'Telephone: '+$ReplaceText 
    $RangeNew = $New.Bookmarks.Item($Bookmark).Range 
    $RangeNew.Text = $ReplaceText 
    $New.Bookmarks.Add($Bookmark,$RangeNew) 
    $RangeReply = $Reply.Bookmarks.Item($Bookmark).Range 
    $RangeReply.Text = $ReplaceText 
    $Reply.Bookmarks.Add($Bookmark,$RangeReply)  
     
    #Fax 
    $Bookmark = "Fax"  
    $ReplaceText = $ADFaxNumber.ToString() 
    $LogInfo += $NL+'Fax Number: '+$ReplaceText 
    $RangeNew = $New.Bookmarks.Item($Bookmark).Range 
    $RangeNew.Text = $ReplaceText 
    $New.Bookmarks.Add($Bookmark,$RangeNew) 
    $RangeReply = $Reply.Bookmarks.Item($Bookmark).Range 
    $RangeReply.Text = $ReplaceText 
    $Reply.Bookmarks.Add($Bookmark,$RangeReply)  
     
    #Mobile 
    If ($Mobile -ne '') { 
        $Bookmark = "Mobile"  
        $ReplaceText = $Mobile 
        $LogInfo += $NL+'MobileNumber: '+$ReplaceText 
        $RangeNew = $New.Bookmarks.Item($Bookmark).Range 
        $RangeNew.Text = $ReplaceText 
        $New.Bookmarks.Add($Bookmark,$RangeNew) 
        $RangeReply = $Reply.Bookmarks.Item($Bookmark).Range 
        $RangeReply.Text = $ReplaceText 
        $Reply.Bookmarks.Add($Bookmark,$RangeReply)   
    } 
 
    #Email Address 
    $Bookmark = 'Email' 
    $ReplaceText = $ADEmailAddress.ToString() 
    $LogInfo += $NL+'Email Address: '+$ReplaceText 
    $RangeNew = $New.Bookmarks.Item($Bookmark).Range 
    $Link = $New.HyperLinks.Add($RangeNew,'mailto:'+$ReplaceText,$null,$null,$ReplaceText) 
    $New.Bookmarks.Add($Bookmark,$RangeNew) 
    $RangeReply = $Reply.Bookmarks.Item($Bookmark).Range 
    $Link = $Reply.HyperLinks.Add($RangeReply,'mailto:'+$ReplaceText,$null,$null,$ReplaceText) 
    $Reply.Bookmarks.Add($Bookmark,$RangeReply) 
     
              
    #Save new message signature  
    Echo 'Saving Signatures' 
    #Save HTML 
    $SaveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML");  
    [ref]$BrowserLevel = "microsoft.office.interop.word.WdBrowserLevel" -as [type]  
    $New.WebOptions.OrganizeInFolder = $true  
    $New.WebOptions.UseLongFileNames = $true  
    $New.WebOptions.BrowserLevel = $BrowserLevel::wdBrowserLevelMicrosoftInternetExplorer6  
    $Path = $LocalSignaturePath+'\'+$CompanyName+'.htm'  
    If ((Test-Path $Path) -eq '1') { (Get-Item $Path).attributes = 'Archive' } 
    $New.SaveAs([ref]$Path, [ref]$saveFormat)  
  
    #Save RTF  
    $SaveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatRTF");  
    $Path = $LocalSignaturePath+'\'+$CompanyName+'.rtf'  
    If ((Test-Path $Path) -eq '1') { (Get-Item $Path).attributes = 'Archive' } 
    $New.SaveAs([ref] $Path, [ref]$saveFormat)  
     
    #Save TXT 
    $SaveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatText");  
    $Path = $LocalSignaturePath+'\'+$CompanyName+'.txt'  
    If ((Test-Path $Path) -eq '1') { (Get-Item $Path).attributes = 'Archive' } 
    $New.SaveAs([ref] $path, [ref]$SaveFormat)  
 
    #Save reply message signature  
    #Save HTML 
    $SaveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML");  
    [ref]$BrowserLevel = "microsoft.office.interop.word.WdBrowserLevel" -as [type]  
    $Reply.WebOptions.OrganizeInFolder = $true  
    $Reply.WebOptions.UseLongFileNames = $true  
    $Reply.WebOptions.BrowserLevel = $BrowserLevel::wdBrowserLevelMicrosoftInternetExplorer6  
    $Path = $LocalSignaturePath+'\'+$CompanyName+'Reply.htm'  
    If ((Test-Path $Path) -eq '1') { (Get-Item $Path).attributes = 'Archive' } 
    $Reply.SaveAs([ref]$Path, [ref]$saveFormat)  
      
    #Save RTF  
    $SaveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatRTF");  
    $Path = $LocalSignaturePath+'\'+$CompanyName+'Reply.rtf'  
    If ((Test-Path $Path) -eq '1') { (Get-Item $Path).attributes = 'Archive' } 
    $Reply.SaveAs([ref] $Path, [ref]$saveFormat)  
     
    #Save TXT 
    $SaveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatText");  
    $Path = $LocalSignaturePath+'\'+$CompanyName+'Reply.txt'  
    If ((Test-Path $Path) -eq '1') { (Get-Item $Path).attributes = 'Archive' } 
    $Reply.SaveAs([ref]$Path, [ref]$SaveFormat)  
     
    #Close Word objects 
    $New.Close() 
    $Reply.Close() 
}  
 
Echo 'Setting Registry Settings' 
#Stamp registry-values for Outlook Signature Settings if they don`t match the initial script variables. Note that these will apply after the second script run when changes are made in the "Custom variables"-section.  
If ($ForcedSignatureNew -eq $ForceSignatureNew){}  
Else { Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name ForcedSignatureNew -Value $ForceSignatureNew }  
  
If ($ForcedSignatureReplyForward -eq $ForceSignatureReplyForward){}  
Else { Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name ForcedSignatureReplyForward -Value $ForceSignatureReplyForward }  
  
If ($SignatureVersion -eq $SigVersion){}  
Else { Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name SignatureVersion -Value $SigVersion } 
 
#Set Registry for signature files 
#Office 2007 
If ((Test-Path HKCU:'\Software\Microsoft\Office\13.0\Common\MailSettings') -eq 1) {  
    If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\13.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
    Else { New-ItemProperty HKCU:'\Software\Microsoft\Office\13.0\Common\MailSettings' -Name 'NewSignature' -Value 'TerraWest' -PropertyType 'String' } 
    If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\13.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
    Else { New-ItemProperty HKCU:'\Software\Microsoft\Office\13.0\Common\MailSettings' -Name 'ReplySignature' -Value 'TerraWestReply' -PropertyType 'String' } 
    $LogInfo += $NL+'Office 2007 Registry values added (HKCU:\Software\Microsoft\Office\13.0\Common\MailSettings' 
} 
 
#Office 2010 
If ((Test-Path HKCU:'\Software\Microsoft\Office\14.0\Common\MailSettings') -eq 1) {  
    If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\14.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
    Else { New-ItemProperty HKCU:'\Software\Microsoft\Office\14.0\Common\MailSettings' -Name 'NewSignature' -Value 'TerraWest' -PropertyType 'String' } 
    If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\14.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
    Else { New-ItemProperty HKCU:'\Software\Microsoft\Office\14.0\Common\MailSettings' -Name 'ReplySignature' -Value 'TerraWestReply' -PropertyType 'String' } 
    $LogInfo += $NL+'Office 2010 Registry values added (HKCU:\Software\Microsoft\Office\14.0\Common\MailSettings' 
} 
 
#Office 2013 
If ((Test-Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings') -eq 1) {  
    If (Get-ItemProperty -Name 'NewSignature' -Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
    Else { New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'NewSignature' -Value 'ale' -PropertyType 'String' } 
    If (Get-ItemProperty -Name 'ReplySignature' -Path HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -ErrorAction SilentlyContinue) { } 
    Else { New-ItemProperty HKCU:'\Software\Microsoft\Office\15.0\Common\MailSettings' -Name 'ReplySignature' -Value 'alereply' -PropertyType 'String' } 
    $LogInfo += $NL+'Office 2013 Registry values added (HKCU:\Software\Microsoft\Office\15.0\Common\MailSettings' 
} 
 
#Log Data 
#LogData 'Signature' $LogInfo 
 
#Cleanup mapped drive and imported functions 
#Remove-PSDrive -Name P 
#Remove-Item Function:\LogData
#$banner = "http://FQDN/banner.jpg" 

$strName = $env:username

$strFilter = "(&(objectCategory=User)(samAccountName=$strName))"

$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.Filter = $strFilter

$objPath = $objSearcher.FindOne()
$objUser = $objPath.GetDirectoryEntry()


$strName = $objUser.FullName
$strTitle = $objUser.Title
$strCompany = $objUser.Company
$strCred = $objUser.info
$strStreet = $objUser.StreetAddress
$strPhone = $objUser.telephoneNumber
$strCity =  $objUser.l
$strPostCode = $objUser.PostalCode
$strCountry = $objUser.co
if ($strCountry[0] = "Netherlands") {$strCountry[0]="The Netherlands"}
$strEmail = $objUser.mail
$strWebsite = $objUser.wWWHomePage

$UserDataPath = $Env:appdata
if (test-path "HKCU:\\Software\\Microsoft\\Office\\11.0\\Common\\General") {
  get-item -path HKCU:\\Software\\Microsoft\\Office\\11.0\\Common\\General | new-Itemproperty -name Signatures -value signaturesCompany -propertytype string -force
  

if (test-path "HKCU:\\Software\\Microsoft\\Office\\12.0\\Common\\General") {
  get-item -path HKCU:\\Software\\Microsoft\\Office\\12.0\\Common\\General | new-Itemproperty -name Signatures -value signaturesCompany -propertytype string -force
}
if (test-path "HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General") {
  get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General | new-Itemproperty -name Signatures -value signaturesCompany -propertytype string -force
}
$FolderLocation = $UserDataPath + '\\Microsoft\\signaturesCompany'  
mkdir $FolderLocation -force

$stream = [System.IO.StreamWriter] "$FolderLocation\\$strName.htm"
$stream.WriteLine("<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.0 Transitional//EN`">")
$stream.WriteLine("<HTML><HEAD><TITLE>Signature</TITLE>")
$stream.WriteLine("<META http-equiv=Content-Type content=`"text/html; charset=windows-1252`">")
$stream.WriteLine("<BODY>")
$stream.WriteLine("<SPAN style=`"FONT-SIZE: 10pt; COLOR: black; FONT-FAMILY: `'Trebuchet MS`'`">")
$stream.WriteLine("<BR><BR>")
$stream.WriteLine("<B><SPAN style=`"FONT-SIZE: 9pt; COLOR: gray; FONT-FAMILY: `'Trebuchet MS`'`">" + ($strName.ToUpper()).Substring(0,$strName.Length-7) + "</SPAN></B>")
$stream.WriteLine("<SPAN style=`"FONT-SIZE: 9pt; COLOR: gray; FONT-FAMILY: `'Trebuchet MS`'`"> - "+ $strTitle[0].ToUpper())
$stream.WriteLine("</SPAN><BR>")
$stream.WriteLine("<SPAN style=`"FONT-SIZE: 9pt; COLOR: gray; FONT-FAMILY: `'Trebuchet MS`'`">")
$stream.WriteLine($strCompany[0] )
$stream.WriteLine(" - " + $strStreet[0].ToUpper() + " - " + $strPostCode + " - " + $strCity[0].ToUpper() +" - " + $strCountry[0].ToUpper() + "</SPAN><BR>")
$stream.WriteLine("<SPAN style=`"FONT-SIZE: 9pt; COLOR: gray; FONT-FAMILY: `'Trebuchet MS`'`">T " + $strPhone)
$stream.WriteLine(" - <A href=`"mailto:"+ $strEmail +"`"><SPAN title=" + $strEmail + " style=`"COLOR: gray; TEXT-DECORATION: none; text-underline: none; FONT-FAMILY: `'Trebuchet MS`'`">" + $strEmail[0].ToUpper() + "</SPAN></A>")
$stream.WriteLine("<SPAN style=`"FONT-SIZE: 9pt; COLOR: gray; FONT-FAMILY: `'Trebuchet MS`'`"> -  </SPAN>")
$stream.WriteLine("<A href=`"http://" + $strWebsite +"`"><SPAN title=" + $strWebsite + " style=`"COLOR: gray; TEXT-DECORATION: none; text-underline: none; FONT-FAMILY: `'Trebuchet MS`'`">" + $strWebsite[0].ToUpper() + "</SPAN></A><BR><BR>")
$stream.WriteLine("<a href=`"http://nl.linkedin.com/in/Company`" target=`"_TOP`"><img src=`"http://FQDN/linkedin.png`" width=`"20`" height=`"15`" alt=`"View Company`'s LinkedIn profile`" style=`"vertical-align:middle`" border=`"0`"></a>")
$stream.WriteLine("<a href=`"http://www.facebook.com/pages/Company`" target=`"_TOP`" title=`"Company`"><img src=`"http://FQDN/facebooklogo.jpg`" width=`"15`" height=`"15`" alt=`"View Company's Facebook profile`" style=`"vertical-align:middle`" border=`"0`"></a>&nbsp")
$stream.WriteLine("<a href=`"http://twitter.com/Company`" target=`"_TOP`"><img src=`"http://FQDN/twitter.png`" width=`"15`" height=`"15`" alt=`"View Company`'s Twitter profile`" style=`"vertical-align:middle`" border=`"0`"></a>")
$stream.WriteLine("<BR><BR>")
$stream.WriteLine("<SPAN style=`"COLOR: gray; TEXT-DECORATION: none; text-underline: none; FONT-FAMILY: `'Trebuchet MS`'`"><IMG height=130 src=`"" + $banner + "`" border=`"0`" width=550></SPAN></A>")
$stream.WriteLine("<BR>")
$stream.WriteLine("</BODY>")
$stream.WriteLine("</HTML>")
$stream.close()
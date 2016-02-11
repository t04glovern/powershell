#==================| Nathan Glover |======================================
#
#            Author  :  Nathan Glover
#            E-Mail  :  nathan@glovers.id.au
#            website :  nathanglover.com
#
#            Created : 11-02-2016
#            Purpose : Retrieve version numbers for Office16 components
#            Version : 1
#
#========================================================================

# This script was originally written for Office 2010:
# https://gallery.technet.microsoft.com/scriptcenter/Get-Offie-Applications-5c0bb2e1

# I have updated it for Microsoft Office 2016 32bit
# Office location and version will be different for 64bit

#-file Locations-----------------------------------------------------

#Ms Word
$wordPath = 'C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE'

#Ms Excel
$excelPath = 'C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE'

#Ms Powerpoint
$powepointPath = 'C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE'

#Ms Outlook
$outlookPath = 'C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE'

#Ms Publisher
$publisherPath = 'C:\Program Files (x86)\Microsoft Office\root\Office16\MSPUB.EXE'

#Ms OneNote
$onenotePath = 'C:\Program Files (x86)\Microsoft Office\root\Office16\ONENOTE.EXE'

#Ms Visio
$visioPath = 'C:\Program Files (x86)\Microsoft Office\root\Office16\VISIO.EXE'

#-File Properties----------------------------------------------------

#word
$wordProperty = Get-ItemProperty $wordPath

#excel
$excelProperty = Get-ItemProperty $excelPath

#outlook
$outLookProperty = Get-ItemProperty $outlookPath

#powerpoint
$powerPointProperty = Get-ItemProperty $powepointPath

#powerpoint
$publisherProperty = Get-ItemProperty $publisherPath

#powerpoint
$onenoteProperty = Get-ItemProperty $onenotePath

#powerpoint
$visioProperty = Get-ItemProperty $visioPath

#-Print Them---------------------------------------------------------

#Creating a new PS-Object
$myObject = New-Object  Psobject

#Word
$myObject | Add-Member Noteproperty -name "Word Product Version " -Value $wordProperty.VersionInfo.ProductVersion
$myObject | Add-Member Noteproperty -name "Word File Version" -Value $wordProperty.VersionInfo.FileVersion

#Excel
$myObject | Add-Member NoteProperty -Name "Excel Product Version" -Value $excelProperty.VersionInfo.ProductVersion
$myObject | Add-Member NoteProperty -Name "Excel File Version" $excelProperty.VersionInfo.FileVersion

#Powerpoint
$myObject | Add-Member NoteProperty -Name "Powerpoint Product Version" -Value $powerPointProperty.VersionInfo.ProductVersion
$myObject | Add-Member NoteProperty -Name "Powerpoint File Version" -Value $powerPointProperty.VersionInfo.FileVersion

#Outlook
$myObject | Add-Member NoteProperty -Name "Outlook Product Version" -Value  $outLookProperty.VersionInfo.ProductVersion
$myObject | Add-Member NoteProperty -Name "Outlook File Version" -Value  $outLookProperty.VersionInfo.FileVersion

#Publisher
$myObject | Add-Member NoteProperty -Name "Publisher Product Version" -Value  $publisherProperty.VersionInfo.ProductVersion
$myObject | Add-Member NoteProperty -Name "Publisher File Version" -Value  $publisherProperty.VersionInfo.FileVersion

#OneNote
$myObject | Add-Member NoteProperty -Name "OneNote Product Version" -Value  $onenoteProperty.VersionInfo.ProductVersion
$myObject | Add-Member NoteProperty -Name "OneNote File Version" -Value  $onenoteProperty.VersionInfo.FileVersion

#Visio
$myObject | Add-Member NoteProperty -Name "Visio Product Version" -Value  $visioProperty.VersionInfo.ProductVersion
$myObject | Add-Member NoteProperty -Name "Visio File Version" -Value  $visioProperty.VersionInfo.FileVersion

# Publishing Product informations.
$myObject

#End of the script.
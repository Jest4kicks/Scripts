###########################################################
##                                                       ##
##                  UnpackPESfiles.ps1                   ##
##      Written by Michael Troutman for Gail Lesnick     ##
##                                                       ##
###########################################################

#This script extracts a .zip file, removes files not matching
#an extension filter, cleans-up empty folders, and moves the
#result to a static folder.
#
#Preparation
#1. The script runs best when the temp directory used is 
#excluded from real-time virus protection.
#
#2. To add this script's functionality to the windows 
#context menu, open regedit and drill down to:
# "HKEY_CLASSES_ROOT\*\shell". Within that key, create a 
#new key called "Unpack PES Files". Then create
#one subkey here called "Command" and change its default
#<REG_SZ> value to:
#"C:\windows\system32\WindowsPowerShell\v1.0\powershell.exe" "c:\scripts\unpackpesfiles.ps1 \"%1\""
#(adjust the locations of the powershell exe, and script
#location, appropriately.

param(
$strzipfile
)

#Excluded file extensions
$strexclude = @("*.pes","*.jpg","*.pdf","*.zip")

Write-host $strzipfile

If ($strzipfile -notlike "*.zip") {
  Write-host "Target is not a zip file.  Closing in 5 seconds."
  Start-sleep -Seconds 5
  Exit
  }

$objzipfile = Get-Item $strzipfile
$strzipshortname = [System.IO.Path]::GetFileNameWithoutExtension($objzipfile)
$strtempdest = $env:TEMP + "\" + $strzipshortname + (get-date -format MMddhhmmss) 
$strdest = "C:\users\Gail\Downloads"

#Phase 1: Unzip the File
Add-Type -assembly "system.io.compression.filesystem"
Write-Host "Unzipping the file."
[io.compression.zipfile]::ExtractToDirectory($strzipfile, $strtempdest)

#Phase 2: Remove Unwanted File Extensions
Write-Host "Removing everything that's not a PDF, JPG, ZIP, or PES file."
$arrbadfiles = Get-Childitem $strtempdest -recurse -exclude $strexclude | 
  where-object { $_.PSIsContainer -eq $False } 

ForEach ( $objbadfile in $arrbadfiles ) {
  #Write-host "Deleting: " $objbadfile.fullname
  Remove-item $objbadfile.fullname
  }

#Phase 3: Remove Empty Folders
Write-host "Removing empty folders."
do {
$repeatflag = 0
$arrFolders = Get-ChildItem $strtempdest -recurse | Where-Object {$_.PSIsContainer -eq $True}

ForEach ( $objsubfolder in $arrFolders ) {
  #Write-host "Checking: " $objsubfolder.FullName

  If ( $objsubfolder.GetFiles().Count -eq 0 -and $objsubfolder.GetDirectories().Count -eq 0 ) {
    #Write-host "Removing empty folder: " $objsubfolder
    Remove-Item $objsubfolder.fullname
    $repeatflag = 1
    }
  }

}
until ($repeatflag -eq 0)
 

#Phase 4: Move Result back to Downloads
Write-Host "Dropping the folder in Downloads."
Move-Item $strtempdest $strdest

Write-host "Complete. Closing."
Start-sleep -Seconds 5
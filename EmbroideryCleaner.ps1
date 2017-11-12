###########################################################
##                                                       ##
##                 EmbroideryCleaner.ps1                 ##
##      Written by Michael Troutman for Gail Lesnick     ##
##                                                       ##
###########################################################

<##########################################################
This script scans a directory tree for files not matching
a defined list of extension-based exclusions.  It then
removes them and cleans-up any resulting empty folders.
##########################################################>

param(
$strfolderroot
)

$repeatflag = 1

#Excluded file extensions
$strexclude = @("*.pes","*.jpg","*.pdf","*.zip")

#Phase 1: Remove Unwanted File Extensions

#Find files
Write-Host "Scanning..."
$arrbadfiles = Get-Childitem $strfolderroot -recurse -exclude $strexclude |
  where-object { $_.PSIsContainer -eq $False }

#Report findings
Write-Host "Found " $arrbadfiles.count " files."
Start-Sleep -Seconds 2

#Opportunity to abort  #lifebeginsatrightclicknew
Write-Host "Starting delete in 5 seconds.  Last chance to Ctrl+C to cancel."
Start-Sleep -Seconds 5

#Delete files
Write-Host "Deleting..."
$arrbadfiles | Remove-Item 


#Phase 2: Remove Empty Folders
Write-host "Removing empty folders."
do {
$repeatflag = 0
$arrFolders = Get-ChildItem $strfolderroot -recurse | Where-Object {$_.PSIsContainer -eq $True}

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

Write-Host "Complete.  Deleted " $arrbadfiles.count " files."
Start-Sleep -Seconds 2

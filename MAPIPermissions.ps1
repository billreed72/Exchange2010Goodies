Write-Host 'Mailbox MAPI Permissions' -ForegroundColor Magenta
Write-Host '------------------------' -ForegroundColor Green
Import-CSV 'c:\userList-ADID-SMTP.csv' |
# The CSV file contains 2 columns of data:
# "PrimarySmtpAddress","Identity"
# "btester@dex10.net","dex10.net/Staff/Bob Tester"
foreach {
$Email = $_.PrimarySmtpAddress
$sam = $_.Identity
########## MAILBOX FOLDERS ############
$folders = Get-MailboxFolderStatistics -Identity $Email | Where {$_.Foldertype -ne "SyncIssues" -and $_.Foldertype -ne "Conflicts" -and $_.Foldertype -ne "LocalFailures" -and $_.Foldertype -ne "ServerFailures" -and $_.Foldertype -ne "RecoverableItemsRoot" -and $_.Foldertype -ne "RecoverableItemsDeletions" -and $_.Foldertype -ne "RecoverableItemsPurges" -and $_.Foldertype -ne "RecoverableItemsVersions" -and $_.Foldertype -ne "Root"} | select folderpath,ItemsInFolder,FolderSize
#Write-Host "Mailbox Folders:" -ForegroundColor Green
Foreach ($Folder in $folders)
 {
#  Write-Host 'Folder:'$Folder.Folderpath 'Items:'$Folder.ItemsInFolder 'Size:'$Folder.FolderSize
 }
########## MAPI PERMISSIONS ###########
Write-Host "MAPI Permissions: " -ForegroundColor Green -NoNewLine; Write-Host $Email -ForegroundColor Yellow -BackgroundColor DarkBlue
get-mailboxfolderpermission -identity $Email":\" | ft foldername, User, AccessRights
Foreach ($Folder in $folders)
 {
  $NormalizedFolder = $Folder.FolderPath.Replace("/","\")
  $NormalizedIdentity = $Email + ":" + $NormalizedFolder
  get-mailboxfolderpermission -identity $NormalizedIdentity | ft foldername, User, AccessRights
 }
}

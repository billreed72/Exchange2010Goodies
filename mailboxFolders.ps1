Write-Host 'Mailbox Folder Statistics' -ForegroundColor Magenta
Write-Host '-------------------------' -ForegroundColor Green
Import-CSV 'c:\userList-ADID-SMTP.csv' |
# The CSV file contains 2 columns of data:
# "PrimarySmtpAddress","Identity"
# "btester@dex10.net","dex10.net/Staff/Bob Tester"
foreach {
$Email = $_.PrimarySmtpAddress
$sam = $_.Identity
########## MAILBOX FOLDERS ############
$folders = Get-MailboxFolderStatistics -Identity $Email | Where {$_.Foldertype -ne "SyncIssues" -and $_.Foldertype -ne "Conflicts" -and $_.Foldertype -ne "LocalFailures" -and $_.Foldertype -ne "ServerFailures" -and $_.Foldertype -ne "RecoverableItemsRoot" -and $_.Foldertype -ne "RecoverableItemsDeletions" -and $_.Foldertype -ne "RecoverableItemsPurges" -and $_.Foldertype -ne "RecoverableItemsVersions" -and $_.Foldertype -ne "Root"} | select folderpath,ItemsInFolder,FolderSize
Foreach ($Folder in $folders)
 {
  Write-Host $Email -foregroundcolor Yellow -NoNewLine; Write-Host ' :Folder: '$Folder.Folderpath ' :Items: '$Folder.ItemsInFolder ' :Size: '$Folder.FolderSize
 }
}

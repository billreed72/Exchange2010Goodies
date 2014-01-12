Write-Host 'INPUT filename.' -ForegroundColor Cyan -BackgroundColor DarkBlue;
$UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
$OutData = @()
$UserList = Get-Content $UserListFile
$CurProcMbx = 1
####### Loop 1: Query user Mailbox Folder Stats
Foreach ($UserID in $UserList) {
  write-progress "Processing Mailbox num: " $CurProcMbx
  $FolderData = Get-MailboxFolderStatistics -Identity $UserID | Where {$_.Foldertype -ne "SyncIssues" -and $_.Foldertype -ne "Conflicts" -and $_.Foldertype -ne "LocalFailures" -and $_.Foldertype -ne "ServerFailures" -and $_.Foldertype -ne "RecoverableItemsRoot" -and $_.Foldertype -ne "RecoverableItemsDeletions" -and $_.Foldertype -ne "RecoverableItemsPurges" -and $_.Foldertype -ne "RecoverableItemsVersions" -and $_.Foldertype -ne "Root"} | select Identity,FolderPath,ItemsInFolder,FolderSize
  $OutData += $FolderData
  $CurProcMbx++
}
$Outdata | Export-csv  -Path ('MailboxFolders-{1:yyyyMMddHHmm}.csv' -f $env:COMPUTERNAME,(Get-Date))

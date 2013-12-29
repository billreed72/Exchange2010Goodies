#--------------------------------------------------------------#
#------------------- READ CAREFULLY!!! ------------------------#
# 
# Step 1: Create a user list with the Active Directory Identity
#   Powershell file: "createUserList-ADID.ps1"
#     Get-MailboxPermission -Identity * | where {$_.user.tostring() -eq "NT AUTHORITY\SELF"} | select Identity | Export-csv -path 'c:\userList-ADID.csv'
#   Execute script: [PS] C:\>.\createUserList-ADID.ps1
# 
# Step 2: Create a user list with the Active Directory Identity and the user's primary SMTP Address
#   Powershell file: "createUserList-ADID-SMTP.ps1"
#     Import-CSV c:\userList-ADID.csv |
#     foreach {
#     Get-Mailbox -Identity $_.Identity | select primarySMTPAddress,Identity
#     }
#   Execute script: [PS] C:\>.\createUserList-ADID-SMTP.ps1 | export-csv -path 'c:\userList-ADID-SMTP.csv'
#
# Step 3: Query Exchange user permissions for:
#         Full Access, Send-As, Send On Behalf, Mailbox Folders, & MAPI
#   Powershell file: "all-in-1-Permissions.ps1"
#   Execute script: [PS] C:\>.\all-in-1-Permissions.ps1
#--------------------------------------------------------------#
Import-CSV 'c:\userList-ADID-SMTP.csv' |
foreach {
$Email = $_.PrimarySmtpAddress
$sam = $_.Identity
Write-Host '  Get Mailbox Permission for: ' -ForegroundColor Magenta -NoNewLine; Write-Host $Email -ForegroundColor Yellow
$FullAccess = Get-MailboxPermission -Identity $Email | where { ($_.AccessRights -eq 'FullAccess' -and $_.IsInherited -eq $False -and $_.User -notlike 'NT AUTHORITY\SELF') } | Select User
Write-Host 'Full Access: ' -ForegroundColor Green
foreach ($User in $FullAccess)
 {
  Write-Host $User.User
 }
$SendAs = Get-ADPermission -Identity $sam | Where {$_.ExtendedRights -like 'Send-As' -and $_.User -notlike 'NT AUTHORITY\SELF' -and $_.Deny -eq $False} | Select User
Write-Host 'Send As: ' -ForegroundColor Green
foreach ($User in $SendAs)
 {
  Write-Host $User.User
 }
$SendOnBehalf = Get-Mailbox -Identity $Email | Select @{Name='SendOnBehalf';Expression={$_.'GrantSendOnBehalfTo'}}
Write-Host 'Send On Behalf: ' -ForegroundColor Green
foreach ($User in $SendOnBehalf)
 {
  Write-Host $User.SendOnBehalf
 }
$folders = Get-MailboxFolderStatistics -Identity $Email | Where {$_.FolderType -ne 'SyncIssues' -and $_.FolderType -ne 'Conflicts' -and $_.FolderType -ne 'LocalFailures' -and $_.FolderType -ne 'ServerFailures' -and $_.FolderType -ne 'RecoverableItemsRoot' -and $_.FolderType -ne 'RecoverableItemsDeletions' -and $_.FolderType -ne 'RecoverableItemsPurges' -and $_.FolderType -ne 'RecoverableItemsVersions' -and $_.Foldertype -ne 'Root'} | Select FolderPath,ItemsInFolder,FolderSize
Write-Host 'Mailbox Folders: ' -ForegroundColor Green
Foreach ($Folder in $folders)
 {
  Write-Host 'Folder: '$Folder.FolderPath 'Items: '$Folder.ItemsInFolder 'Size: '$Folder.FolderSize
 }
Write-Host 'MAPI Permissions: ' -ForegroundColor Green
Get-MailboxFolderPermissions -Identity $Email":\" | ft FolderName, User, AccessRights
foreach ($Folder in $folders)
 {
  $NormalizedFolder = $Folder.FolderPath.Replace("/","\")
  $NormalizedIdentity = $Email + ':' + $NormalizedFolder
  Get-MailboxFolderPermissions -Identity $NormalizedIdentity | ft FolderName, User, AccessRights
 }
}

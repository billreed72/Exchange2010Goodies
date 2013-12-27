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

Write-Host "*******************************************************************************" -foregroundcolor magenta
Write-Host "  Get Mailbox Permission for: "-foregroundcolor magenta
Write-Host "                             $Email" -foregroundcolor yellow
Write-Host "*******************************************************************************" -foregroundcolor magenta

########## FULL ACCESS ################
$FullAccess = Get-MailboxPermission -Identity $Email | where { ($_.AccessRights -eq "FullAccess" -and $_.IsInherited -eq $false -and $_.User -notlike "NT AUTHORITY\SELF") } | Select User
Write-Host "Full Access:" -foregroundcolor Green
Foreach ($User in $FullAccess)
 {
  Write-Host $User.User
 }

########## SEND-AS ####################
$SendAs = Get-ADPermission -Identity $sam | Where {$_.ExtendedRights -like "Send-As" -and $_.User -notlike "NT AUTHORITY\SELF" -and $_.Deny -eq $false} | Select User
Write-Host "Send As:" -foregroundcolor Green
Foreach ($User in $SendAs)
 {
  Write-Host $User.User
 }

########## SEND ON BEHALF #############
#$SendOnBehalf = get-mailbox -Identity $Email | Select GrantSendOnBehalfTo
$SendOnBehalf = get-mailbox -Identity $Email | Select @{Name="SendOnBehalf";Expression={$_."GrantSendOnBehalfTo"}}
Write-Host "Send On Behalf:" -foregroundcolor Green
Foreach ($User in $SendOnBehalf)
 {
  Write-Host $User.SendOnBehalf
 }

########## MAILBOX FOLDERS ############
$folders = Get-MailboxFolderStatistics -Identity $Email | Where {$_.Foldertype -ne "SyncIssues" -and $_.Foldertype -ne "Conflicts" -and $_.Foldertype -ne "LocalFailures" -and $_.Foldertype -ne "ServerFailures" -and $_.Foldertype -ne "RecoverableItemsRoot" -and $_.Foldertype -ne "RecoverableItemsDeletions" -and $_.Foldertype -ne "RecoverableItemsPurges" -and $_.Foldertype -ne "RecoverableItemsVersions" -and $_.Foldertype -ne "Root"} | select folderpath,ItemsInFolder,FolderSize
Write-Host "Mailbox Folders:" -foregroundcolor Green
Foreach ($Folder in $folders)
 {
  Write-Host 'Folder:'$Folder.Folderpath 'Items:'$Folder.ItemsInFolder 'Size:'$Folder.FolderSize
 }

########## MAPI PERMISSIONS ###########
Write-Host "MAPI Permissions:" -foregroundcolor Green
get-mailboxfolderpermission -identity $Email":\" | ft foldername, User, AccessRights
Foreach ($Folder in $folders)
 {
  $NormalizedFolder = $Folder.FolderPath.Replace("/","\")
  $NormalizedIdentity = $Email + ":" + $NormalizedFolder
  get-mailboxfolderpermission -identity $NormalizedIdentity | ft foldername, User, AccessRights
 }
}

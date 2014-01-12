Write-Host 'INPUT filename.' -ForegroundColor Cyan -BackgroundColor DarkBlue;
$UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
$OutData = @()
$UserList = Get-Content $UserListFile
$CurProcMbx = 1
####### Loop 1: Query user Mailbox Permissions for Full Access
Foreach ($UserID in $UserList) {
  write-progress "Processing Mailbox num: " $CurProcMbx
  $GrantedFullAccessList = @()
  $FullAccessUserID = Get-MailboxPermission -Identity $UserID | Where { ($_.AccessRights -eq 'FullAccess' -and $_.IsInherited -eq $False -and $_.User -notlike 'NT AUTHORITY\SELF') } | Select User
  $GrantedFullAccessList += $FullAccessUserID
###### Loop 2: For each FullAccessUserID that's not NULL, list the grantor & grantee
Foreach ($FullAccessUserID in $GrantedFullAccessList) {
  If ($FullAccessUserID -ne $NULL) {
    $OutObject = "" | select Mailbox, FullAccess
    $OutObject.Mailbox = $UserID
    $OutObject.FullAccess = (Get-recipient $FullAccessUserID.User).PrimarySmtpAddress.ToString()
    $Outdata += $OutObject
    $OutObject
  }
  $CurProcMbx++
}
}
$Outdata | Export-csv  -Path ('FullAccessOutput-{1:yyyyMMddHHmm}.csv' -f $env:COMPUTERNAME,(Get-Date))

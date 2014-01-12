Write-Host 'INPUT filename.' -ForegroundColor Cyan -BackgroundColor DarkBlue;
$UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
$OutData = @()
$UserList = Get-Content $UserListFile
$CurProcMbx = 1
####### Loop 1: Convert smtp to ADIDs
Foreach ($UserID in $UserList) {
  write-progress "Processing Mailbox num: " $CurProcMbx
  $ADIDList = @()
  $UserADID = Get-mailbox $UserID | select PrimarySMTPAddress,Identity
  $ADIDList += $UserADID
####### Loop 2: Query user AD Permissions for Send-As
Foreach ($Identity in $ADIDList) {
  $GrantedSendAsList = @()
  $SendAsUserAD = Get-ADPermission $UserADID.Identity | Where {$_.ExtendedRights -like 'Send-As' -and $_.IsInherited -eq $False -and $_.User -notlike 'NT AUTHORITY\SELF'} | Select User
  $GrantedSendAsList += $SendAsUserAD
####### Loop 3: For each SendasUserAD, that's not NULL,, list the grantor & grantee
Foreach ($SendAsUserAD in $GrantedSendAsList) {
  If ($SendAsUserAD -ne $NULL) {
    $OutObject = "" | select Mailbox, SendAsAccess
    $OutObject.Mailbox = $UserID
    $OutObject.SendAsAccess = (Get-recipient $SendAsUserAD.User).PrimarySmtpAddress.ToString()
    $Outdata += $OutObject
    $OutObject
    }
  }
  $CurProcMbx++
}
}
$Outdata | Export-csv  -Path ('SendAs-{1:yyyyMMddHHmm}.csv' -f $env:COMPUTERNAME,(Get-Date))

Write-Host 'INPUT filename:' -ForegroundColor Cyan -BackgroundColor DarkBlue;
$UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
$OutData = @()
$UserList = get-content $UserListFile
$CurProcMbx = 1
###### Function 1: Need to gain understanding
Function RecExpand ($grpn) {
$grpfinal= @()
$grp = Get-DistributionGroupMember -Identity $grpn -ResultSize unlimited
###### Loop w/in Function: Expand group members into array
Foreach ($g in $grp) {
  if($g.RecipientType -like "*group*"){$grpfinal += RecExpand $g.Tostring()}
  else{$grpfinal += $g.Tostring()}
 }
 Return $grpfinal
}
###### Loop 1: Get the UserID, save to array
Foreach ($UserID in $UserList) {
  write-progress "Processing Mailbox num: " $CurProcMbx
  $FinalList = @()
  $User = Get-mailbox $UserID
  $InitialList = $User.GrantSendOnBehalfTo
###### Loop 2: Query for groups to expand, otherwise, save recipients to array
Foreach ($recipient in $InitialList) {
    $type = (Get-recipient $recipient.Name).RecipientType
    If ($type -like "*group*") {$FinalList += RecExpand ($recipient.Name)}
    Else {$FinalList += $recipient}
  }
###### Loop 3: 
  Foreach ($recipient in $FinalList) {
    If ($recipient -ne $NULL) {
      $OutObject = "" | select Mailbox, SendOnBehalfAccess
      $OutObject.Mailbox = $User.PrimarySmtpAddress
      $OutObject.SendOnBehalfAccess = (Get-Recipient $recipient).PrimarySmtpAddress.ToString()
      $Outdata += $OutObject
      $OutObject
    }
  }
  $CurProcMbx++
}
$Outdata | Export-csv  -Path ('SendOnBehalf-{1:yyyyMMddHHmm}.csv' -f $env:COMPUTERNAME,(Get-Date));

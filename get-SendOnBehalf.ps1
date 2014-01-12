Write-Host 'INPUT filename:' -ForegroundColor Cyan -BackgroundColor DarkBlue;
$UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
$OutData = @()
Function RecExpand ($grpn) {
$grpfinal= @()
$grp = Get-DistributionGroupMember -Identity $grpn -ResultSize unlimited
Foreach ($g in $grp) {
  if($g.RecipientType -like "*group*"){$grpfinal += RecExpand $g.Tostring()}
  else{$grpfinal += $g.Tostring()}
 }
 Return $grpfinal
}
$UserList = get-content $UserListFile
$CurProcMbx = 1
Foreach ($UserID in $UserList) {
  write-progress "Processing Mailbox num: " $CurProcMbx
  $FinalList = @()
  $User = Get-mailbox $UserID
  $InitialList = $User.GrantSendOnBehalfTo
  Foreach ($recipient in $InitialList) {
    $type = (Get-recipient $recipient.Name).RecipientType
    If ($type -like "*group*") {$FinalList += RecExpand ($recipient.Name)}
    Else {$FinalList += $recipient}
  }
  Foreach ($recipient in $FinalList) {
    If ($recipient -ne $NULL) {
      $OutObject = "" | select Mailbox, Delegate
      $OutObject.Mailbox = $User.PrimarySmtpAddress
      $OutObject.Delegate = (Get-Recipient $recipient).PrimarySmtpAddress.ToString()
      $Outdata += $OutObject
      $OutObject
    }
  }
  $CurProcMbx++
}
$Outdata | Export-csv  -Path ('SendOnBehalfOutput-{1:yyyyMMddHHmm}.csv' -f $env:COMPUTERNAME,(Get-Date));
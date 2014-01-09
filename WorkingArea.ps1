Write-Host 'Mailbox                        Send On Behalf Granted' -ForegroundColor Magenta
Write-Host '-------                        ----------------------' -ForegroundColor Green
$userList = Read-Host 'Please specify a userList csv file.'
Import-CSV '$userList' |
# The CSV file contains 2 columns of data:
# "PrimarySmtpAddress","Identity"
# "btester@dex10.net","dex10.net/Staff/Bob Tester"
foreach {
$Email = $_.PrimarySmtpAddress
$sam = $_.Identity
########## SEND ON BEHALF #############
#$SendOnBehalf = get-mailbox -Identity $Email | Select GrantSendOnBehalfTo
$SendOnBehalf = get-mailbox -Identity $Email | Select @{Name="SendOnBehalf";Expression={$_."GrantSendOnBehalfTo"}}
Foreach ($User in $SendOnBehalf)
 {
  Write-Host $Email -ForegroundColor Magenta -NoNewLine; Write-Host " - " $User.SendOnBehalf  -ForegroundColor Cyan
 }
}

param([string]$UserListFile=$(throw "User List needed!"))
$OutData = @()
$OutFile = "SharedMailboxDia-Deleg.csv"
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

$OutData | export-csv $Outfile

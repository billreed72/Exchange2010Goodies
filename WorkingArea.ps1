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

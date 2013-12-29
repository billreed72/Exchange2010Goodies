Write-Host 'Mailbox			Send On Behalf Granted' -ForegroundColor Magenta
Write-Host '-------			----------------------' -ForegroundColor Green
Import-CSV 'c:\userList-ADID-SMTP.csv' |
foreach {
$Email = $_.PrimarySmtpAddress
$sam = $_.Identity
########## SEND ON BEHALF #############
#$SendOnBehalf = get-mailbox -Identity $Email | Select GrantSendOnBehalfTo
$SendOnBehalf = get-mailbox -Identity $Email | Select @{Name="SendOnBehalf";Expression={$_."GrantSendOnBehalfTo"}}
Foreach ($User in $SendOnBehalf)
 {
  Write-Host $Email -ForegroundColor Magenta -NoNewLine; Write-Host " - " $User.User  -ForegroundColor Cyan
 }
}

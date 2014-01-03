Write-Host 'Mailbox___Full Access Granted' -ForegroundColor Magenta
Write-Host '-----------------------------' -ForegroundColor Green
Import-CSV 'c:\userList-ADID-SMTP.csv' |
# The CSV file contains 2 columns of data:
# "PrimarySmtpAddress","Identity"
# "btester@dex10.net","dex10.net/Staff/Bob Tester"
ForEach {
$Email = $_.PrimarySmtpAddress
$sam = $_.Identity
########################################
$FullAccess = Get-MailboxPermission -Identity $Email | Where { ($_.AccessRights -eq 'FullAccess' -and $_.IsInherited -eq $False -and $_.User -notlike 'NT AUTHORITY\SELF') } | Select User
Foreach ($User in $FullAccess)
 {
  Write-Host $Email -ForegroundColor Magenta -NoNewLine; Write-Host ' - ' $User.User  -ForegroundColor Cyan
 }
}

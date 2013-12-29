Write-Host 'Mailbox___Full Access Granted' -ForegroundColor Magenta
Write-Host '-----------------------------' -ForegroundColor Green
Import-CSV 'c:\userList-ADID-SMTP.csv' |
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

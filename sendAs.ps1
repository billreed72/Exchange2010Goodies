Write-Host 'Mailbox			Send On Behalf Granted' -ForegroundColor Magenta
Write-Host '-------			----------------------' -ForegroundColor Green
Import-CSV 'c:\userList-ADID-SMTP.csv' |
foreach {
$Email = $_.PrimarySmtpAddress
$sam = $_.Identity
########################################
$SendAs = Get-ADPermission -Identity $sam | Where {$_.ExtendedRights -like 'Send-As' -and $_.IsInherited -eq $False -and $_.User -notlike 'NT AUTHORITY\SELF'} | Select User
foreach ($User in $SendAs)
 {
  Write-Host $Email -ForegroundColor Magenta -NoNewLine; Write-Host " - " $User.User  -ForegroundColor Cyan
 }
}

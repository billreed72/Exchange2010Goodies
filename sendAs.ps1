Write-Host 'Mailbox			Send As Granted' -ForegroundColor Magenta
Write-Host '-------			---------------' -ForegroundColor Green
Import-CSV 'c:\userList-ADID-SMTP.csv' |
# The CSV file contains 2 columns of data:
# "PrimarySmtpAddress","Identity"
# "btester@dex10.net","dex10.net/Staff/Bob Tester"
foreach {
$Email = $_.PrimarySmtpAddress
$sam = $_.Identity
$SendAs = Get-ADPermission -Identity $sam | Where {$_.ExtendedRights -like 'Send-As' -and $_.IsInherited -eq $False -and $_.User -notlike 'NT AUTHORITY\SELF'} | Select User
foreach ($User in $SendAs)
 {
  Write-Host $Email -ForegroundColor Magenta -NoNewLine; Write-Host " - " $User.User  -ForegroundColor Cyan
 }
}

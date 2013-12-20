$Password=Read-Host "Enter Password" -AsSecureString
Import-CSV C:\userList.csv |
foreach {
$domainName = '@dex10.net'
$userprincipalname = $_.Firstname + '.' +  $_.Lastname + $domainName
$firstandlastname = $_.Firstname + ' ' +  $_.Lastname
$OU = 'dex10.net/Staff'
$database = 'Mailbox Database 0155389800'
New-Mailbox -Name $firstandlastname -Alias $_.alias -OrganizationalUnit $OU -UserPrincipalName $userprincipalname -SamAccountName $_.alias -FirstName $_.Firstname -LastName $_.Lastname -Initials '' -Password $Password -ResetPasswordOnNextLogon $false -Database $database
}

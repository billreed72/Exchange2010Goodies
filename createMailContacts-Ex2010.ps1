Import-CSV C:\userList.csv |
foreach {
$domainAlias = '@galias.dex10.net'
$ExternalEmailAddress = $_.Firstname + '.' +  $_.Lastname + $domainAlias
$name = $_.Firstname + ' ' +  $_.Lastname + ' (SD)'
$OU = 'dex10.net/Special Delivery'
New-MailContact -ExternalEmailAddress $ExternalEmailAddress -Name $name -Alias $_.alias -FirstName $_.Firstname -LastName $_.Lastname -OrganizationalUnit $OU
}

Import-CSV C:\userList.csv |
foreach {
$userprincipalname = $_.Firstname + “.” +  $_.Lastname + “@dexlab.net”
$ExternalEmailAddress = $_.Firstname + “.” +  $_.Lastname + “@galias.dexlab.net”
$name = $_.Firstname + “ ” +  $_.Lastname + " (SD)"
$OU = "dexlab.net/Users"
New-MailContact -ExternalEmailAddress $ExternalEmailAddress -Name $name -Alias $_.alias -OrganizationalUnit $OU -FirstName $_.Firstname -LastName $_.Lastname
}

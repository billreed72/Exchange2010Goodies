Import-CSV C:\userList.csv |
foreach {
Set-MailContact $_.alias -HiddenFromAddressListsEnabled $True
}

Import-CSV C:\userList.csv |
foreach {
Get-MailContact $_.alias | select HiddenFromAddressListsEnabled,ExternalEmailAddress,DisplayName
}

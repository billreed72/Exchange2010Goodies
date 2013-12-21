Import-CSV C:\userList.csv |
foreach {
Get-Mailbox $_.alias | select Name, ForwardingAddress, DeliverToMailboxAndForward
}

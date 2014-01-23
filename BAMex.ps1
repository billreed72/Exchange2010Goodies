$xAppName = "BAM! (Bill's Access Manager) for Exchange 2010 – Version 0.1"
$unAss = "(tbd) *** [Unassigned]"
[BOOLEAN]$global:xExitSession=$false
function LoadMenuSystem(){
[INT]$xMenu1=0
[INT]$xMenu2=0
[BOOLEAN]$xValidSelection=$false
while ( $xMenu1 -lt 1 -or $xMenu1 -gt 4 ){
CLS
################################################################################
# MENU A: MAIN MENU
################################################################################
Write-Host “`n`t$xAppName`n” -ForegroundColor Magenta
Write-Host “`t`tPlease select an option`n” -Fore Cyan
# MENU OPTION TEXT A-1
Write-Host “`t`t`t1. Query Exchange Properties” -Fore Cyan
# MENU OPTION TEXT A-2
Write-Host “`t`t`t2. Query Multiple Mailbox Properties” -Fore Cyan
# MENU OPTION TEXT A-3
Write-Host “`t`t`t3. Other Misc Functions” -Fore Cyan
# MENU OPTION TEXT A-4
Write-Host “`t`t`t4. Quit and exit`n” -Fore Cyan
#… Retrieve the response from the user
[int]$xMenu1 = Read-Host “`t`tEnter Menu Option Number”
if( $xMenu1 -lt 1 -or $xMenu1 -gt 4 ){
Write-Host “`tPlease select one of the options available.`n” -Fore Red;start-Sleep -Seconds 
}
}
Switch ($xMenu1){    #… User has selected a valid entry.. load next menu
########################################################################
# MAIN MENU FUNCTION 1
########################################################################
1 {
while ( $xMenu2 -lt 1 -or $xMenu2 -gt 4 ){
CLS
################################################################
# SUB-MENU A-1: EXCHANGE PROPERTIES MENU
################################################################
Write-Host “`n`t$xAppName`n” -Fore Magenta
Write-Host “`t`tPlease select an option`n” -Fore Green
# MENU OPTION TEXT A-1-1
Write-Host “`t`t`t1. Exchange Schema Versions” -Fore Green
# MENU OPTION TEXT A-1-2
Write-Host “`t`t`t2. Exchange Server Names and Versions” -Fore Green
# MENU OPTION TEXT A-1-3
Write-Host “`t`t`t3. $unAss” -Fore DarkGreen
# MENU OPTION TEXT A-1-4
Write-Host “`t`t`t4. Go to Main Menu`n” -Fore Green
[int]$xMenu2 = Read-Host “`t`tEnter Menu Option Number”
if( $xMenu1 -lt 1 -or $xMenu1 -gt 4 ){
Write-Host “`tPlease select one of the options available.`n” -Fore Red;start-Sleep -Seconds 1
}
}
Switch ($xMenu2){
################
# MENU OPTION FUNCTION A-1-1
################
1{
Import-Module ActiveDirectory
$OutEXVdata = @()

$ExForestAndDomain = read-host 'Please enter your forest and domain (i.e. DC=dev10,DC=net)'
$ExOrg = read-host 'Please enter the Exchange Org Name (i.e. First Organization)'

$ExchangeSchemaVersion = get-ADObject "CN=ms-Exch-Schema-Version-pt,CN=Schema,CN=Configuration,$ExForestAndDomain" -Property rangeUpper | select rangeUpper
$ExchangeOrganizationForestVersion = get-ADObject "CN=$ExOrg,CN=Microsoft Exchange,CN=Services,CN=Configuration,$ExForestAndDomain" -Property objectVersion | select objectVersion
$ExchangeOrganizationDomainVersion = get-ADObject "CN=Microsoft Exchange System Objects,$ExForestAndDomain" -Property objectVersion | select objectVersion
$OutEXVer = "" | select ExSchmV,ExOrgForV,ExchOrgDomV
$OutEXVer.ExSchmV = $ExchangeSchemaVersion.rangeUpper
$OutEXVer.ExOrgForV = $ExchangeOrganizationForestVersion.objectVersion
$OutEXVer.ExchOrgDomV = $ExchangeOrganizationDomainVersion.objectVersion
$OutEXVdata += $OutEXVer

$SavePathEXVdata = ('ExchangeSchema-{1:yyyyMMddHHmmss}.csv' -f $env:COMPUTERNAME,(Get-Date))
$OutEXVdata | Export-csv  -Path $SavePathEXVdata
Write-host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
Write-host $SavePathEXVdata -Fore DarkRed -Back gray;start-sleep -seconds 3
}
################

################
# MENU OPTION FUNCTION A-1-2
################
2{
Import-Module ActiveDirectory
$ExchangeServerData = @()
$AdminDisplayVersion = get-exchangeServer | select *
$OutDVer = "" | select Name,ADV
$OutDVer.Name = $AdminDisplayVersion.Name
$OutDVer.ADV = $AdminDisplayVersion.AdminDisplayVersion
$ExchangeServerData += $OutDVer

$SavePathExServerData = ('ExchangeServers-{1:yyyyMMddHHmm}.csv' -f $env:COMPUTERNAME,(Get-Date))
$ExchangeServerData | Export-csv  -Path $SavePathExServerData
Write-host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
Write-host $SavePathExServerData -Fore DarkRed -Back gray;start-sleep -seconds 3
}
################
# MENU OPTION FUNCTION A-1-3
################
3{ Write-Host “`n`tYou Selected Option 3 – Put your Function or Action Here`n” -Fore Yellow;start-Sleep -Seconds 3 }
################

################
# MENU OPTION FUNCTION A-1-4
################
default { Write-Host “`n`tYou Selected Option 4 – Quit the Administration Tasks`n” -Fore Yellow; break}
################
}
}
################
########################################################################

########################################################################
# MAIN MENU FUNCTION 2
########################################################################
2 {
while ( $xMenu2 -lt 1 -or $xMenu2 -gt 4 ){
CLS
################################################################
# SUB-MENU A-2: MULTI-USER QUERY MENU
################################################################
Write-Host “`n`t$xAppName`n” -Fore Magenta
Write-Host “`t`tPlease select an option`n” -Fore Green
Write-Host “`t`t`t1. Who's got Full Access” -Fore Green
Write-Host “`t`t`t2. Who's got Send On Behalf Access” -Fore Green
Write-Host “`t`t`t3. Who's got Send-As Access” -Fore Green
Write-Host “`t`t`t4. Go to Main Menu`n” -Fore Green
[int]$xMenu2 = Read-Host “`t`tEnter Menu Option Number”
}
if( $xMenu1 -lt 1 -or $xMenu1 -gt 4 ){
Write-Host “`tPlease select one of the options available.`n” -Fore Red;start-Sleep -Seconds 1
}
Switch ($xMenu2){
################
# MENU OPTION FUNCTION A-2-1 FULL ACCESS
################
1{
Write-Host 'INPUT filename.' -ForegroundColor Cyan -BackgroundColor DarkBlue;
$UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
$OutFAData = @()
$UserList = Get-Content $UserListFile
$CurProcMbxFA = 1
####### Loop 1: Query user Mailbox Permissions for Full Access
Foreach ($UserID in $UserList) {
  write-host -NoNewLine $CurProcMbxFA -Fore Blue -Back White; write-host '.' -Fore Red -Back White -NoNewLine
  #write-progress "Mailboxes Queried for Full Access: " $CurProcMbxFA
  $GrantedFullAccessList = @()
  $FullAccessUserID = Get-MailboxPermission -Identity $UserID | Where { ($_.AccessRights -eq 'FullAccess' -and $_.IsInherited -eq $False -and $_.User -notlike 'NT AUTHORITY\SELF') } | Select User
  $GrantedFullAccessList += $FullAccessUserID
###### Loop 2: For each FullAccessUserID that's not NULL, list the grantor & grantee
Foreach ($FullAccessUserID in $GrantedFullAccessList) {
  If ($FullAccessUserID -ne $NULL) {
    $OutFAObject = "" | select Mailbox, FullAccess
    $OutFAObject.Mailbox = $UserID
    $OutFAObject.FullAccess = (Get-recipient $FullAccessUserID.User).PrimarySmtpAddress.ToString()
    $OutFAData += $OutFAObject
    $OutObject
  }
  $CurProcMbxFA++
}
}
$SavePathFAdata = ('FullAccess-{1:yyyyMMddHHmm}.csv' -f $env:COMPUTERNAME,(Get-Date))
$OutFAData | Export-csv  -Path $SavePathFAdata
Write-host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
Write-host $SavePathFAdata -Fore DarkRed -Back gray;start-sleep -seconds 3
}
################

################
# MENU OPTION FUNCTION A-2-2 SEND ON BEHALF
################
2{
Write-Host 'INPUT filename:' -ForegroundColor Cyan -BackgroundColor DarkBlue;
$UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
$OutSOBData = @()
$UserList = get-content $UserListFile
$CurProcMbxSOB = 1
###### Function 1: Need to gain understanding
Function RecExpand ($grpn) {
$grpfinal= @()
$grp = Get-DistributionGroupMember -Identity $grpn -ResultSize unlimited
###### Loop w/in Function: Expand group members into array
Foreach ($g in $grp) {
  if($g.RecipientType -like "*group*"){$grpfinal += RecExpand $g.Tostring()}
  else{$grpfinal += $g.Tostring()}
 }
 Return $grpfinal
}
###### Loop 1: Get the UserID, save to array
Foreach ($UserID in $UserList) {
  write-host -NoNewLine $CurProcMbxSOB -Fore Blue -Back White; write-host '.' -Fore Red -Back White -NoNewLine
  #write-progress "Mailboxes Queried for Send On Behalf: " $CurProcMbxSOB
  $FinalList = @()
  $User = Get-mailbox $UserID
  $InitialList = $User.GrantSendOnBehalfTo
###### Loop 2: Query for groups to expand, otherwise, save recipients to array
Foreach ($recipient in $InitialList) {
    $type = (Get-recipient $recipient.Name).RecipientType
    If ($type -like "*group*") {$FinalList += RecExpand ($recipient.Name)}
    Else {$FinalList += $recipient}
  }
###### Loop 3: 
  Foreach ($recipient in $FinalList) {
    If ($recipient -ne $NULL) {
      $OutSOBObject = "" | select Mailbox, SendOnBehalfAccess
      $OutSOBObject.Mailbox = $User.PrimarySmtpAddress
      $OutSOBObject.SendOnBehalfAccess = (Get-Recipient $recipient).PrimarySmtpAddress.ToString()
      $OutSOBData += $OutSOBObject
      #$OutSOBObject
    }
  }
  $CurProcMbxSOB++
}
$SavePathSOBdata = ('SendOnBehalf-{1:yyyyMMddHHmm}.csv' -f $env:COMPUTERNAME,(Get-Date))
$OutSOBData | Export-csv  -Path $SavePathSOBdata
Write-host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
Write-host $SavePathSOBdata -Fore DarkRed -Back gray;start-sleep -seconds 3
}
################
# MENU OPTION FUNCTION A-2-3 SEND-AS
################
3{
Write-Host 'INPUT filename.' -ForegroundColor Cyan -BackgroundColor DarkBlue;
$UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
$OutSAData = @()
$UserList = Get-Content $UserListFile
$CurProcMbxSA = 1
####### Loop 1: Convert smtp to ADIDs
Foreach ($UserID in $UserList) {
  write-host -NoNewLine $CurProcMbxSA -Fore Blue -Back White; write-host '.' -Fore Red -Back White -NoNewLine
  #write-progress "Processing Mailbox num: " $CurProcMbx
  $ADIDList = @()
  $UserADID = Get-mailbox $UserID | select PrimarySMTPAddress,Identity
  $ADIDList += $UserADID
####### Loop 2: Query user AD Permissions for Send-As
Foreach ($Identity in $ADIDList) {
  $GrantedSendAsList = @()
  $SendAsUserAD = Get-ADPermission $UserADID.Identity | Where {$_.ExtendedRights -like 'Send-As' -and $_.IsInherited -eq $False -and $_.User -notlike 'NT AUTHORITY\SELF'} | Select User
  $GrantedSendAsList += $SendAsUserAD
####### Loop 3: For each SendasUserAD, that's not NULL,, list the grantor & grantee
Foreach ($SendAsUserAD in $GrantedSendAsList) {
  If ($SendAsUserAD -ne $NULL) {
    $OutSAObject = "" | select Mailbox, SendAsAccess
    $OutSAObject.Mailbox = $UserID
    $OutSAObject.SendAsAccess = (Get-recipient $SendAsUserAD.User).PrimarySmtpAddress.ToString()
    $OutSAData += $OutSAObject
    #$OutSAObject
    }
  }
  $CurProcMbxSA++
}
}
$SavePathSAdata = ('SendAs-{1:yyyyMMddHHmm}.csv' -f $env:COMPUTERNAME,(Get-Date))
$OutSAData | Export-csv  -Path $SavePathSAdata
Write-host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
Write-host $SavePathSAdata -Fore DarkRed -Back gray;start-sleep -seconds 3
}
################
# MENU OPTION FUNCTION A-2-4
################
default { Write-Host “`n`tYou Selected Option 4 – Go to Main Menu`n” -Fore Yellow; break }
################
}
}
########################################################################
# MAIN MENU FUNCTION 3
########################################################################
3 {
while ( $xMenu2 -lt 1 -or $xMenu2 -gt 4 ){
CLS
################################################################
# SUB-MENU A-3: MISC
################################################################
Write-Host “`n`t$xAppName`n” -Fore Magenta
Write-Host “`t`tPlease select an option`n” -Fore Green
# MENU OPTION TEXT A-3-1
Write-Host “`t`t`t1. $unAss” -Fore DarkGreen
# MENU OPTION TEXT A-3-2
Write-Host “`t`t`t2. $unAss” -Fore DarkGreen
# MENU OPTION TEXT A-3-3
Write-Host “`t`t`t3. $unAss” -Fore DarkGreen
# MENU OPTION TEXT A-3-4
Write-Host “`t`t`t4. Go to Main Menu`n” -Fore Green
[int]$xMenu2 = Read-Host “`t`tEnter Menu Option Number”
if( $xMenu1 -lt 1 -or $xMenu1 -gt 4 ){
Write-Host “`tPlease select one of the options available.`n” -Fore Red;start-Sleep -Seconds 1
}
}
Switch ($xMenu2){
################
# MENU OPTION FUNCTION A-3-1
################
1{ Write-Host “`n`tYou Selected Option 1 – Put your Function or Action Here`n” -Fore Yellow;start-Sleep -Seconds 3 }
################

################
# MENU OPTION FUNCTION A-3-2
################
2{ Write-Host “`n`tYou Selected Option 2 – Put your Function or Action Here`n” -Fore Yellow;start-Sleep -Seconds 3 }
################

################
# MENU OPTION FUNCTION A-3-3
################
3{ Write-Host “`n`tYou Selected Option 3 – Put your Function or Action Here`n” -Fore Yellow;start-Sleep -Seconds 3 }
################

################
# MENU OPTION FUNCTION A-3-4
################
default { Write-Host “`n`tYou Selected Option 4 – Go to Main Menu`n” -Fore Yellow; break }
################
}
}
########################################################################
# MAIN MENU FUNCTION 4
########################################################################
default { $global:xExitSession=$true;break }
########################################################################
}
}
LoadMenuSystem
If ($xExitSession){
Exit-PSSession    #… User quit & Exit
} Else {
.\BAMex.ps1    #… Loop the function
$CurProcMbx = 1
}

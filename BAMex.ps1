$xAppName = "BAM! for Exchange 2010 – Version 0.1"
$unAss = "(tbd) *** [Unassigned]"
[BOOLEAN]$global:xExitSession=$false
#==============================================================================
# FUNCTION: Exchange Schema Versions
#==============================================================================
function GetExchangeSchemaVerions {
    Import-Module ActiveDirectory
    $OutEXVdata = @()
    $ExForestAndDomain = Read-Host 'Please enter your forest and domain (i.e. DC=dev10,DC=net)'
    $ExOrg = Read-Host 'Please enter the Exchange Org Name (i.e. First Organization)'
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
    Write-Host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
    Write-Host $SavePathEXVdata -Fore DarkRed -Back gray;start-sleep -seconds 1
}
#==============================================================================
# FUNCTION: Exchange Server Names and Versions
#==============================================================================
function GetExchangeServerNamesADV {
    Import-Module ActiveDirectory
    $ExchangeServerData = @()
    $AdminDisplayVersion = get-exchangeServer | select *
    $OutDVer = "" | select Name,ADV
    $OutDVer.Name = $AdminDisplayVersion.Name
    $OutDVer.ADV = $AdminDisplayVersion.AdminDisplayVersion
    $ExchangeServerData += $OutDVer
    $SavePathExServerData = ('ExchangeServers-{1:yyyyMMddHHmm}.csv' -f $env:COMPUTERNAME,(Get-Date))
    $ExchangeServerData | Export-csv  -Path $SavePathExServerData
    Write-Host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
    Write-Host $SavePathExServerData -Fore DarkRed -Back gray;start-sleep -seconds 1
}
#==============================================================================
# FUNCTION: Get Full Access Permissions
#==============================================================================
function GetFullAccess {
    Write-Host 'INPUT filename.' -Fore Cyan -Back DarkBlue;
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
    $OutFAData = @()
    $UserList = Get-Content $UserListFile
    $CurProcMbxFA = 1
    foreach ($UserID in $UserList) {
        Write-Host -NoNewLine $CurProcMbxFA -Fore Blue -Back White; write-host '.' -Fore Red -Back White -NoNewLine
        $GrantedFullAccessList = @()
        $FullAccessUserID = Get-MailboxPermission -Identity $UserID | Where { ($_.AccessRights -eq 'FullAccess' -and $_.IsInherited -eq $False -and $_.User -notlike 'NT AUTHORITY\SELF') } | Select User
        $GrantedFullAccessList += $FullAccessUserID
        foreach ($FullAccessUserID in $GrantedFullAccessList) {
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
    Write-Host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
    Write-Host $SavePathFAdata -Fore DarkRed -Back gray;start-sleep -seconds 1
}
#==============================================================================
# FUNCTION: Get Send On Behlaf Access Permissions
#==============================================================================
function GetSendOnBehalfAccess {
    Write-Host 'INPUT filename:' -Fore Cyan -Back DarkBlue;
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
    $OutSOBData = @()
    $UserList = Get-Content $UserListFile
    $CurProcMbxSOB = 1
    function RecExpand ($grpn) {
        $grpfinal= @()
        $grp = Get-DistributionGroupMember -Identity $grpn -ResultSize unlimited
        foreach ($g in $grp) {
            if($g.RecipientType -like "*group*"){$grpfinal += RecExpand $g.Tostring()}
            else{$grpfinal += $g.Tostring()
            }
        }
        return $grpfinal
    }
    foreach ($UserID in $UserList) {
        Write-Host -NoNewLine $CurProcMbxSOB -Fore Blue -Back White; Write-Host '.' -Fore Red -Back White -NoNewLine
        $FinalList = @()
        $User = Get-mailbox $UserID
        $InitialList = $User.GrantSendOnBehalfTo
        foreach ($recipient in $InitialList) {
            $type = (Get-recipient $recipient.Name).RecipientType
                if ($type -like "*group*") {$FinalList += RecExpand ($recipient.Name)}
                else {$FinalList += $recipient}
                }
        foreach ($recipient in $FinalList) {
            if ($recipient -ne $NULL) {
            $OutSOBObject = "" | select Mailbox, SendOnBehalfAccess
            $OutSOBObject.Mailbox = $User.PrimarySmtpAddress
            $OutSOBObject.SendOnBehalfAccess = (Get-Recipient $recipient).PrimarySmtpAddress.ToString()
            $OutSOBData += $OutSOBObject
            }
        }
    $CurProcMbxSOB++
    }
    $SavePathSOBdata = ('SendOnBehalf-{1:yyyyMMddHHmm}.csv' -f $env:COMPUTERNAME,(Get-Date))
    $OutSOBData | Export-csv  -Path $SavePathSOBdata
    Write-Host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
    Write-Host $SavePathSOBdata -Fore DarkRed -Back gray;start-sleep -seconds 1
}
#==============================================================================
# FUNCTION: Get Send As Access Permissions
#==============================================================================
function GetSendAsAccess {
    Write-Host 'INPUT filename.' -Fore Cyan -Back DarkBlue;
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
    $OutSAData = @()
    $UserList = Get-Content $UserListFile
    $CurProcMbxSA = 1
    foreach ($UserID in $UserList) {
        Write-Host -NoNewLine $CurProcMbxSA -Fore Blue -Back White; write-host '.' -Fore Red -Back White -NoNewLine
        $ADIDList = @()
        $UserADID = Get-mailbox $UserID | select PrimarySMTPAddress,Identity
        $ADIDList += $UserADID
        foreach ($Identity in $ADIDList) {
            $GrantedSendAsList = @()
            $SendAsUserAD = Get-ADPermission $UserADID.Identity | Where {$_.ExtendedRights -like 'Send-As' -and $_.IsInherited -eq $False -and $_.User -notlike 'NT AUTHORITY\SELF'} | Select User
            $GrantedSendAsList += $SendAsUserAD
            foreach ($SendAsUserAD in $GrantedSendAsList) {
                if ($SendAsUserAD -ne $NULL) {
                $OutSAObject = "" | select Mailbox, SendAsAccess
                $OutSAObject.Mailbox = $UserID
                $OutSAObject.SendAsAccess = (Get-recipient $SendAsUserAD.User).PrimarySmtpAddress.ToString()
                $OutSAData += $OutSAObject
                }
            }
        $CurProcMbxSA++
        }
    }
    $SavePathSAdata = ('SendAs-{1:yyyyMMddHHmm}.csv' -f $env:COMPUTERNAME,(Get-Date))
    $OutSAData | Export-csv  -Path $SavePathSAdata
    Write-Host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
    Write-Host $SavePathSAdata -Fore DarkRed -Back gray;start-sleep -seconds 1
}
#========================================================================
# FUNCTION: SETUP Dual Delivery (CREATES CONTACTS,SETS DELV OPTS, & HIDES CONTACTS)
#========================================================================
function SetupDualDelivery {
    Write-Host 'INPUT filename.' -Fore Cyan -Back DarkBlue
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)'
    $sdDomain = Read-Host 'Remote domain Special Delivery ( @galias.domain.com )'
    $sdOU = Read-Host 'OU for Special Delivery Contacts ( dev10.net/SpecialDelivery )'
    $UserList = Get-Content $UserListFile    
    foreach ($UserID in $UserList) {
        $F = (Get-recipient $USerID).firstName
        $L = (Get-recipient $UserID).lastName
        $D = (Get-recipient $UserID).displayName
        $A = (Get-recipient $UserID).alias
        $sdA = $A+"-DD"
        $sdSMTP = $A+$sdDomain
        $sdDName = $D+"(DD)"
        New-MailContact -ExternalEmailAddress $sdSMTP -Name $sdDName -Alias $sdA -FirstName $F -LastName $L -OrganizationalUnit $sdOU
        Set-Mailbox $UserID -DeliverToMailboxAndForward:$True -ForwardingAddress $sdSMTP
        Set-MailContact $sdA -HiddenFromAddressListsEnabled $True
    }
}
#========================================================================
# FUNCTION: SETUP Split Delivery (CREATES CONTACTS,SETS DELV OPTS, & HIDES CONTACTS)
#========================================================================
function SetupSplitDelivery {
    Write-Host 'INPUT filename.' -Fore Cyan -Back DarkBlue
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)'
    $sdDomain = Read-Host 'Special Delivery Domain ( @galias.domain.com )'
    $sdOU = Read-Host 'Special Delivery Contacts OU ( dev10.net/SpecialDelivery )'
    $UserList = Get-Content $UserListFile    
    foreach ($UserID in $UserList) {
        $F = (Get-recipient $USerID).firstName
        $L = (Get-recipient $UserID).lastName
        $D = (Get-recipient $UserID).displayName
        $A = (Get-recipient $UserID).alias
        $sdA = $A+"-SD"
        $sdSMTP = $A+$sdDomain
        $sdDName = $D+"(SD)"
        New-MailContact -ExternalEmailAddress $sdSMTP -Name $sdDName -Alias $sdA -FirstName $F -LastName $L -OrganizationalUnit $sdOU
        Set-Mailbox $UserID -DeliverToMailboxAndForward:$True -ForwardingAddress $sdSMTP
        Set-MailContact $sdA -HiddenFromAddressListsEnabled $True
    }
}
#========================================================================
# FUNCTION: VERIFY DUAL DELIVERY CONTACT DATA
#========================================================================
function ValidateDDContactData {
    Write-Host 'INPUT filename.' -Fore Cyan -Back DarkBlue
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)'
    $UserList = Get-Content $UserListFile  
    foreach ($UserID in $UserList) {
        $F = (Get-recipient $USerID).firstName
        $L = (Get-recipient $UserID).lastName
        $D = (Get-recipient $UserID).displayName
        $A = (Get-recipient $UserID).alias
        $sdA = $A+"-DD"
        Get-MailContact $sdA | select OrganizationalUnit,DisplayName,ExternalEmailAddress,HiddenFromAddressListsEnabled
    }
}
#========================================================================
# FUNCTION: VERIFY SPLIT DELIVERY CONTACT DATA
#========================================================================
function ValidateDDContactData {
    Write-Host 'INPUT filename.' -Fore Cyan -Back DarkBlue
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)'
    $UserList = Get-Content $UserListFile  
    foreach ($UserID in $UserList) {
        $F = (Get-recipient $USerID).firstName
        $L = (Get-recipient $UserID).lastName
        $D = (Get-recipient $UserID).displayName
        $A = (Get-recipient $UserID).alias
        $sdA = $A+"-SD"
        Get-MailContact $sdA | select OrganizationalUnit,DisplayName,ExternalEmailAddress,HiddenFromAddressListsEnabled
    }
}
#========================================================================
# FUNCTION: VERIFY DELIVERY OPTIONS
#========================================================================
function VerifyDeliveryOptions {
    Write-Host 'INPUT filename.' -Fore Cyan -Back DarkBlue
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)'
    $UserList = Get-Content $UserListFile 
    foreach ($UserID in $UserList) {
        $A = (Get-recipient $UserID).alias
        Get-Mailbox $A | select Name, ForwardingAddress, DeliverToMailboxAndForward
  }
}
#==============================================================================
# FUNCTION: Load Main Menu
#==============================================================================
function LoadMenuSystem () {
    [INT]$xMenu1=0
    [INT]$xMenu2=0
    [BOOLEAN]$xValidSelection=$false
    while ( $xMenu1 -lt 1 -or $xMenu1 -gt 4 ){
    CLS
    Write-Host "`n`t$xAppName`n" -Fore Magenta
    Write-Host "`t`tPlease select an option`n" -Fore Cyan
    Write-Host "`t`t`t1. Query Exchange Properties" -Fore Cyan
    Write-Host "`t`t`t2. Query Multiple Mailbox Properties" -Fore Cyan
    Write-Host "`t`t`t3. Special Delivery" -Fore Cyan
    Write-Host "`t`t`t4. Quit and exit`n" -Fore Cyan
    [int]$xMenu1 = Read-Host "`t`tEnter Menu Option Number"
        if( $xMenu1 -lt 1 -or $xMenu1 -gt 4 ){
        Write-Host "`tPlease select one of the options available.`n" -Fore Red;start-Sleep -Seconds 1
        }
    }
    switch ($xMenu1){    #… User has selected a valid entry.. load next menu
    #==============================================================================
    # MENU: Exchange Properties
    #==============================================================================
1 {
    while ( $xMenu2 -lt 1 -or $xMenu2 -gt 4 ){
    CLS
    Write-Host "`n`t$xAppName`n" -Fore Magenta
    Write-Host "`t`tPlease select an option`n" -Fore Green
    Write-Host "`t`t`t1. Exchange Schema Versions" -Fore Green
    Write-Host "`t`t`t2. Exchange Server Names and Versions" -Fore Green
    Write-Host "`t`t`t3. $unAss" -Fore DarkGreen
    Write-Host "`t`t`t4. Go to Main Menu`n" -Fore Green
    [int]$xMenu2 = Read-Host "`t`tEnter Menu Option Number"
        if( $xMenu1 -lt 1 -or $xMenu1 -gt 4 ){
        Write-Host "`tPlease select one of the options available.`n" -Fore Red;start-Sleep -Seconds 1
        }
    }
        switch ($xMenu2){
    1 { GetExchangeSchemaVerions }
    2 { GetExchangeServerNamesADV }
    3 { Write-Host "`n`t$unAss`n" -Fore Yellow;start-Sleep -Seconds 1 }
    default { Write-Host "`n`tYou Selected Option 4 – Quit the Administration Tasks`n" -Fore Yellow; break }
    }
}
    #==============================================================================
    # MENU: Query Multiple Mailbox Properties
    #==============================================================================
2 {
    while ( $xMenu2 -lt 1 -or $xMenu2 -gt 4 ){
    CLS
    Write-Host "`n`t$xAppName`n" -Fore Magenta
    Write-Host "`t`tPlease select an option`n" -Fore Green
    Write-Host "`t`t`t1. Who's got Full Access" -Fore Green
    Write-Host "`t`t`t2. Who's got Send On Behalf Access" -Fore Green
    Write-Host "`t`t`t3. Who's got Send-As Access" -Fore Green
    Write-Host "`t`t`t4. Go to Main Menu`n" -Fore Green
    [int]$xMenu2 = Read-Host "`t`tEnter Menu Option Number"
    }
        if( $xMenu1 -lt 1 -or $xMenu1 -gt 4 ){
        Write-Host "`tPlease select one of the options available.`n" -Fore Red;start-Sleep -Seconds 1
        }
    switch ($xMenu2){
    1 { GetFullAccess }
    2 { GetSendOnBehalfAccess }
    3 { GetSendAsAccess }
    default { Write-Host "`n`tYou Selected Option 4 – Go to Main Menu`n" -Fore Yellow; break }
    }
}
    #==============================================================================
    # MENU: Special Delivery Menu
    #==============================================================================
3 {
    while ( $xMenu2 -lt 1 -or $xMenu2 -gt 4 ){
    CLS
    Write-Host "`n`t$xAppName`n" -Fore Magenta
    Write-Host "`t`tPlease select an option`n" -Fore Green
    Write-Host "`t`t`t1. Set Dual-Delivery" -Fore Green
    Write-Host "`t`t`t2. Set Split-Delivery" -Fore Green
    Write-Host "`t`t`t3. Special Delivery Report" -Fore Green
    Write-Host "`t`t`t4. Go to Main Menu`n" -Fore Green
    [int]$xMenu2 = Read-Host "`t`tEnter Menu Option Number"
        if( $xMenu1 -lt 1 -or $xMenu1 -gt 4 ){
        Write-Host "`tPlease select one of the options available.`n" -Fore Red;start-Sleep -Seconds 1
        }
    }
    switch ($xMenu2){
    1 { SetupDualDelivery }
    2 { SetupSplitDelivery }
    3 { Write-Host "`n`t$unAss`n" -Fore Yellow;start-Sleep -Seconds 1 }
    default { Write-Host "`n`tYou Selected Option 4 – Go to Main Menu`n" -Fore Yellow; break }
    }
}
default { $global:xExitSession=$true;break }
}
}
LoadMenuSystem
if ($xExitSession){
Exit-PSSession
} else {
.\BAMex.ps1
$CurProcMbx = 1
}

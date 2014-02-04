$xAppName = "BAM! (Bill's Application Manager) – Version 0.2"
$unAss = "***[Unassigned]***"
$BamLogName = "BAMex"
$BamLogSource = "BAMSource"
If (!((Get-EventLog -List | Select-Object "Log") -match $BamLogName)) {new-EventLog -LogName $BamLogName -Source $BamLogSource}
#======================================
# Setup And Restrict Delivery
#======================================
function setupAndRestrictDelivery {
function createDistro {
    write-EventLog -LogName $BamLogName -EventID 666 -Message "Delivery Restriction setup started." -Source $BamLogSource -EntryType Information
    $distroName = @()
    $distroName=Read-Host "`tNew Distro Name (i.e. NoDelivery)"
    $distroOU=Read-Host "`tRestricted Delivery OU Name (i.e. dev10.net/Restricted Delivery)"
    new-DistributionGroup -Name $distroName -OrganizationalUnit $distroOU -SamAccountName $distroName -Alias $distroName | Out-Null
    set-Group -Identity $distroName -Notes "Created by BAMex!"
    write-EventLog -LogName $BamLogName -EventID 666 -Message "Distro [$distroName] created in Organizational Unit [$distroOU]." -Source $BamLogSource -EntryType Information
}
function createTransRule {
    $tRuleName=Read-Host "`tEnter a Transportation Rule Name"
    $tRuleFromMemberOf=(get-distributionGroup $distroName).primarySmtpAddress
    $tRuleRejectMessage="You are no longer authorized to send email from this system."
    $tRuleRejectMessageStatusCode="5.7.1"
    New-TransportRule -Name $tRuleName -Comments '' -Priority '0' -Enabled $true -FromMemberOf $tRuleFromMemberOf -RejectMessageReasonText $tRuleRejectMessage -RejectMessageEnhancedStatusCode $tRuleRejectMessageStatusCode | Out-Null
    write-EventLog -LogName $BamLogName -EventID 666 -Message "Transportation rule [$tRuleName] created." -Source $BamLogSource -EntryType Information
}    
function addMembersToDistro {
    Write-Host 'INPUT filename.' -Fore Cyan -Back DarkBlue;
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)'
    $UserList = Get-Content $UserListFile
    $groupName = get-group | Where { ($_.Notes -contains 'Created by BAMex!') } | select Identity
    $CurProcMbxARM = 1
    foreach ($UserID in $UserList) {
        Write-Host -NoNewLine $CurProcMbxARM -Fore Blue -Back White; write-host '.' -Fore Red -Back White -NoNewLine
        Add-DistributionGroupMember -Identity $groupName.Identity -Member $UserID
        $CurProcMbxARM++
    }
}
createDistro
createTransRule
measure-command { addMembersToDistro }; start-sleep -seconds 3
write-EventLog -LogName $BamLogName -EventID 666 -Message "Delivery Restriction setup completed." -Source $BamLogSource -EntryType Information
}
#======================================
# FUNCTION: Exchange Schema Versions
#======================================
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
#======================================
# FUNCTION: Exchange Server Names and Versions
#======================================
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
#======================================
# FUNCTION: Get Full Access Permissions
#======================================
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
#======================================
# FUNCTION: Get Send On Behlaf Access Permissions
#======================================
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
#======================================
# FUNCTION: Get Send As Access Permissions
#======================================
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
#======================================
# FUNCTION: SETUP Dual Delivery (CREATES CONTACTS,SETS DELV OPTS, & HIDES CONTACTS)
#======================================
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
#======================================
# FUNCTION: SETUP Split Delivery (CREATES CONTACTS,SETS DELV OPTS, & HIDES CONTACTS)
#======================================
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
        Set-Mailbox $UserID -DeliverToMailboxAndForward:$False -ForwardingAddress $sdSMTP
        Set-MailContact $sdA -HiddenFromAddressListsEnabled $True
    }
}
#======================================
# FUNCTION: VERIFY DUAL DELIVERY CONTACT DATA
#======================================
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
#======================================
# FUNCTION: VERIFY SPLIT DELIVERY CONTACT DATA
#======================================
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
#======================================
# FUNCTION: VERIFY DELIVERY OPTIONS
#======================================
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
#==============================================================================
#==============================================================================
# MENU: Special Delivery Menu
#==============================================================================
Function thinkSpecialDelivery {
Function showMenuSpecialDelivery {
    Param (
        [Parameter(Position=0,Mandatory=$TRUE,HelpMessage="Special Delivery Menu Help text")] [ValidateNotNullOrEmpty()] [string]$menuSpecialDelivery,
        [Parameter(Position=1)] [ValidateNotNullOrEmpty()] [string]$TitleSpecialDelivery="menuSpecialDelivery",
        [switch] $clearScreen
    )
    if ($clearScreen) {Clear-Host}
    Write-Host "`n`t$xAppName`n" -fore Magenta
    $menuSpecialDeliveryPrompt = $titleSpecialDelivery
    $menuSpecialDeliveryPrompt += "`n`t"
    $menuSpecialDeliveryPrompt += "="*$titleSpecialDelivery.Length
    $menuSpecialDeliveryPrompt += "`n"
    $menuSpecialDeliveryPrompt += $menuSpecialDelivery
    Read-Host -Prompt $menuSpecialDeliveryPrompt
}
$menuSpecialDelivery=@"
    1 Setup Dual Delivery
    2 Setup Split Delivery
    3 Setup & Restrict Delivery
    M Main Menu

    Select a task by number or M
"@
Do {
    Switch (showMenuSpecialDelivery $menuSpecialDelivery "`tSpecial Delivery Tasks" -clearScreen) {
        "1" { SetupDualDelivery }
        "2" { SetupSplitDelivery }
        "3" { setupAndRestrictDelivery }
        "M" { Return }
        Default { Write-Warning "Special Delivery MENU: Invalid Choice. Try again.";sleep -seconds 1 }
    }
} While ($TRUE) 
}
#==============================================================================
# MENU: MAILBOX PERMISSIONS
#==============================================================================
Function thinkMenuMailboxPermissions {
Function showMenuMailboxPermissions {
    Param(
    [Parameter(Position=0,Mandatory=$TRUE,HelpMessage="Mailbox Permissions Menu Help text")] [ValidateNotNullOrEmpty()] [string]$menuMailboxPermissions,
    [Parameter(Position=1)] [ValidateNotNullOrEmpty()] [string]$TitleMailboxPermissions="menuMailboxPermissions" ,
    [switch]$clearScreen
    )
    if ($clearScreen) {Clear-Host}
    Write-Host "`n`t$xAppName`n" -Fore Magenta
    $menuMailboxPermissionsPrompt=$titleMailboxPermissions
    $menuMailboxPermissionsPrompt+="`n`t"
    $menuMailboxPermissionsPrompt+="="*$titleMailboxPermissions.Length
    $menuMailboxPermissionsPrompt+="`n"
    $menuMailboxPermissionsPrompt+=$menuMailboxPermissions
    Read-Host -Prompt $menuMailboxPermissionsPrompt
}
$menuMailboxPermissions=@"
    1 Who has Full Access?
    2 Who has Send On Behalf Access?
    3 Who has Send As Access?
    M Main Menu

    Select a task by number or M
"@
Do {
    Switch (showMenuMailboxPermissions $menuMailboxPermissions "`tMailbox Permissions Tasks" -clearScreen) {
        "1" { GetFullAccess }
        "2" { GetSendOnBehalfAccess }
        "3" { GetSendAsAccess }
        "M" { Return }
        Default { Write-Warning "Mailbox Permissions MENU: Invalid Choice. Try again.";sleep -seconds 1 }
    }
} While ($TRUE) 
}
#==============================================================================
# MENU: EXCHANGE SYSTEM PROPERTIES
#==============================================================================
Function thinkMenuExchange {
Function showMenuExchange {
    Param(
    [Parameter(Position=0,Mandatory=$TRUE,HelpMessage="Exchange Menu Help Text")] [ValidateNotNullOrEmpty()] [string]$menuExchange,
    [Parameter(Position=1)] [ValidateNotNullOrEmpty()] [string]$TitleExchange="menuExchange" ,
    [switch]$clearScreen
    )
    if ($clearScreen) {Clear-Host}
    Write-Host "`n`t$xAppName`n" -Fore Magenta
    $menuExchangePrompt=$titleExchange
    $menuExchangePrompt+="`n`t"
    $menuExchangePrompt+="="*$titleExchange.Length
    $menuExchangePrompt+="`n"
    $menuExchangePrompt+=$menuExchange
    Read-Host -Prompt $menuExchangePrompt
}
$menuExchange=@"
    1 GetExchangeSchemaVerions
    2 GetExchangeServerNamesADV
    3 $unAss
    M Main Menu

    Select a task by number or M
"@
Do {
    Switch (showMenuExchange $menuExchange "`tExchange Tasks" -clearScreen) {
        "1" { GetExchangeSchemaVerions }
        "2" { GetExchangeServerNamesADV }
        "3" { Write-Host $unAss -fore Yellow -back blue; sleep -seconds 1 }
        "M" { Return }
        Default { Write-Warning "Exchange MENU: Invalid Choice. Try again.";sleep -seconds 1 }
    }
} While ($TRUE)
}
#==============================================================================
# MENU: MAIN APPLICATION MENU
#==============================================================================
Function thinkMenuMain {
    Function showMenuMain {
        Param(
        [Parameter(Position=0,Mandatory=$TRUE,HelpMessage="Main Menu Help text")] [ValidateNotNullOrEmpty()] [string]$menuMain,
        [Parameter(Position=1)] [ValidateNotNullOrEmpty()] [string]$TitleMain="menuMain" ,
        [switch]$clearScreen
        )
        if ($clearScreen) {Clear-Host}
        Write-Host "`n`t$xAppName`n" -Fore Magenta
        $menuMainPrompt = $titleMain
        $menuMainPrompt += "`n`t"
        $menuMainPrompt += "="*$titleMain.Length
        $menuMainPrompt += "`n"
        $menuMainPrompt += $menuMain
        Read-Host -Prompt $menuMainPrompt
    }
$menuMain=@"
    1 Query Exchange System
    2 Query Multiple Mailbox Properties
    3 Special Delivery
    Q Quit

    Select a task by number or Q to quit
"@
    Do {
        Switch (showMenuMain $menuMain "`tBAM! Operations" -clearScreen) {
            "1" { thinkMenuExchange }
            "2" { thinkMenuMailboxPermissions }
            "3" { thinkSpecialDelivery }
            "Q" { Write-Host "Have a nice day..." -fore Yellow -back darkBlue; sleep -milliseconds 300; 
                  Write-Host "Goodbye..." -fore cyan -back darkBlue;sleep -milliseconds 300; 
                  Return }
            Default { Write-Warning "MAIN MENU: Invalid Choice. Try again.";sleep -seconds 1 }
        }
    } While ($TRUE)
}
#==============================================================================
thinkMenuMain

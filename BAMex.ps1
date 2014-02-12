#-------------------------------------------------------------------------------
# GENERAL VARIABLES & CREATION OF WINDOWS APPLICATION EVENT LOG, IF NOT CREATED
#-------------------------------------------------------------------------------
$xAppName = "BAM! (Bill's Application Manager) – Version 0.2"
$createdOn = 'Feb 6, 2014 13:00:00 EST'
$createdBy = 'Bill Reed, wreed@appirio.com'
$company = 'Appirio, Inc.'
$unAss = "***[Unassigned]***"
$BamLogName = "BAMex"
$BamLogSource = "BAMSource"
If (!((Get-EventLog -List | Select-Object "Log") -match $BamLogName)) {new-EventLog -LogName $BamLogName -Source $BamLogSource}

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FUNCTION: Exchange Schema Versions
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FUNCTION: Exchange Server Names and Versions
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FUNCTION: Get Full Access Permissions
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FUNCTION: Get Send On Behalf Access Permissions
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FUNCTION: Get Send As Access Permissions
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FUNCTION: Get Mailbox Folder Total Message Counts & Sizes
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function GetMailboxFolderMsgCountsAndSize {
    Write-Host 'INPUT filename.' -ForegroundColor Cyan -BackgroundColor DarkBlue;
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
    $OutMFSData = @()
    $UserList = Get-Content $UserListFile
    $CurProcMbxMFS = 1
    write-EventLog -LogName $BamLogName -EventID 61 -Message "Results: Get Mailbox Folder Total Message Counts & Sizes saved: Started." -Source $BamLogSource -EntryType Information
    Foreach ($UserID in $UserList) {
        write-host -NoNewLine $CurProcMbxMFS -Fore Blue -Back White; write-host '.' -Fore Red -Back White -NoNewLine
        $FolderData = Get-MailboxFolderStatistics -Ide $UserID | Where {$_.Foldertype -ne "SyncIssues" -and $_.Foldertype -ne "Conflicts" -and $_.Foldertype -ne "LocalFailures" -and $_.Foldertype -ne "ServerFailures" -and $_.Foldertype -ne "RecoverableItemsRoot" -and $_.Foldertype -ne "RecoverableItemsDeletions" -and $_.Foldertype -ne "RecoverableItemsPurges" -and $_.Foldertype -ne "RecoverableItemsVersions" -and $_.Foldertype -ne "Root"} | select Identity,FolderPath,ItemsInFolder,FolderSize
        $OutMFSData += $FolderData
        $CurProcMbxMFS++
    }
    $SavePathFolderStatdata = ('MailboxFolderStats-{1:yyyyMMddHHmmss}.csv' -f $env:COMPUTERNAME,(Get-Date))
    $OutMFSdata | Export-csv  -Path $SavePathFolderStatdata
    Write-Host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
    Write-Host $SavePathFolderStatdata -Fore DarkRed -Back gray;start-sleep -seconds 1
    write-EventLog -LogName $BamLogName -EventID 62 -Message "Results: Get Mailbox Folder Total Message Counts & Sizes saved: [$SavePathFolderStatdata]." -Source $BamLogSource -EntryType Information
}
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FUNCTION: Get Mailbox Total Message Count & Size
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function GetMailboxMsgCountsAndSize {
    Write-Host 'INPUT filename.' -Fore Cyan -Back DarkBlue;
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
    $OutStatData = @()
    $UserList = Get-Content $UserListFile
    $CurProcMbxStat = 1
    write-EventLog -LogName $BamLogName -EventID 71 -Message "Results: Get Mailbox Total Message Count & Size saved: Started." -Source $BamLogSource -EntryType Information
    foreach ($UserID in $UserList) {
        Write-Host -NoNewLine $CurProcMbxStat -Fore Blue -Back White; write-host '.' -Fore Red -Back White -NoNewLine
        $StatList = @()
        $StatList += $UserID
        foreach ($USerID in $StatList) {
            If ($UserID -ne $NULL) {
            $OutStatObject = "" | select Mailbox,ItemCount,TotalItemSize,LastLogonTime,LastLoggedOnUserAccount,OriginatingServer
            $OutStatObject.Mailbox = $UserID
            $OutStatObject.ItemCount = (Get-MailboxStatistics -Identity $UserID).ItemCount
            $OutStatObject.TotalItemSize = (Get-MailboxStatistics -Identity $UserID).TotalItemSize
            $OutStatObject.LastLogonTime = (Get-MailboxStatistics -Identity $UserID).LastLogonTime
            $OutStatObject.LastLoggedOnUserAccount = (Get-MailboxStatistics -Identity $UserID).LastLoggedOnUserAccount
            $OutStatObject.OriginatingServer = (Get-MailboxStatistics -Identity $UserID).OriginatingServer
            $OutStatData += $OutStatObject
            }
        $CurProcMbxStat++
        }
    }
    $SavePathStatdata = ('MailboxStats-{1:yyyyMMddHHmmss}.csv' -f $env:COMPUTERNAME,(Get-Date))
    $OutStatData | Export-csv  -Path $SavePathStatdata
    Write-Host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
    Write-Host $SavePathStatdata -Fore DarkRed -Back gray;start-sleep -seconds 1
    write-EventLog -LogName $BamLogName -EventID 72 -Message "Results: Get Mailbox Total Message Count & Size saved: [$SavePathStatdata]." -Source $BamLogSource -EntryType Information
}
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FUNCTION: SETUP Dual Delivery (CREATES CONTACTS,SETS DELV OPTS, & HIDES CONTACTS)
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FUNCTION: SETUP Split Delivery (CREATES CONTACTS,SETS DELV OPTS, & HIDES CONTACTS)
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FUNCTION: Setup & Restrict Delivery for mailboxes
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FUNCTION: VERIFY DUAL DELIVERY CONTACT DATA
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FUNCTION: VERIFY SPLIT DELIVERY CONTACT DATA
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
# FUNCTION: VERIFY DELIVERY OPTIONS
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function VerifyDeliveryOptions {
    Write-Host 'INPUT filename.' -Fore Cyan -Back DarkBlue
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)'
    $UserList = Get-Content $UserListFile 
    foreach ($UserID in $UserList) {
        $A = (Get-recipient $UserID).alias
        Get-Mailbox $A | select Name, ForwardingAddress, DeliverToMailboxAndForward
  }
}
#=================================
# MENU: Special Delivery Menu
#=================================
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
  SPECIAL DELIVERY OPTIONS:
    1 Dual Delivery: Add Mailboxes
    2 Dual Delivery: Remove Mailboxes
    
    3 Split Delivery: Add Mailboxes
    4 Split Delivery: Remove Mailboxes
    
    5 Verify Delivery Options
 
  DENY DELIVERY OPTIONS:
    6 Deny Delivery: Setup
    7 Deny Delivery: Add Mailboxes
    8 Deny Delivery: Remove Mailboxes
    9 Deny Delivery: Tear-Down
    10 Deny Delivery: Report
    
    M Main Menu

    Select a task by number or M
"@
Do {
    Switch (showMenuSpecialDelivery $menuSpecialDelivery "`tSpecial Delivery Menu" -clearScreen) {
        "1" { SetupDualDelivery }
        "2" { write-host 'test2' -fore green;start-sleep -seconds 1 }
        "3" { SetupDualDelivery }
        "4" { write-host 'test4' -fore green;start-sleep -seconds 1 }
        "5" { write-host 'test5' -fore green;start-sleep -seconds 1 }
        "6" { setupAndRestrictDelivery }
        "7" { write-host 'test7' -fore green;start-sleep -seconds 1 }
        "8" { write-host 'test8' -fore green;start-sleep -seconds 1 }
        "9" { write-host 'test9' -fore green;start-sleep -seconds 1 }
        "10" { write-host 'test10' -fore green;start-sleep -seconds 1 }                
        "M" { Return }
        Default { Write-Warning "Special Delivery MENU: Invalid Choice. Try again.";sleep -seconds 1 }
    }
} While ($TRUE) 
}
#=================================
# MENU: Query Mailbox Statistics
#=================================
Function thinkMailboxStats {
Function showMenuMailboxStats {
    Param(
    [Parameter(Position=0,Mandatory=$TRUE,HelpMessage="Enter your Mailbox Statistics menu text")] [ValidateNotNullOrEmpty()] [string]$menuMailboxStats,
    [Parameter(Position=1)] [ValidateNotNullOrEmpty()] [string]$TitleMailboxStats="menuMailboxStats" ,
    [switch]$clearScreen
    )
    if ($clearScreen) {Clear-Host}
        Write-Host "`n`t$xAppName`n" -Fore Magenta
        $menuMailboxStatsPrompt=$titleMailboxStats
        $menuMailboxStatsPrompt+="`n`t"
        $menuMailboxStatsPrompt+="-"*$titleMailboxStats.Length
        $menuMailboxStatsPrompt+="`n"
        $menuMailboxStatsPrompt+=$menuMailboxStats
        Read-Host -Prompt $menuMailboxStatsPrompt
}
$menuMailboxStats=@"
    1 Message Counts & Sizes by Mailbox
    2 Message Counts & Sizes by Mailbox Folder
    3 $unAss
    M Main Menu

    Select a task by number or M
"@
Do {
    Switch (showMenuMailboxStats $menuMailboxStats "`tMailbox Statistics Menu" -clearscreen) {
        "1" {Measure-Command{GetMailboxMsgCountsAndSize};start-sleep 3}
        "2" {Measure-Command{GetMailboxFolderMsgCountsAndSize};start-sleep 3}
        "3" {Write-Host $unAss -fore Green; sleep -seconds 1 }
        "M" { Return }
        Default {Write-Warning "MailboxStats MENU: Invalid Choice. Try again.";sleep -milliseconds 750}
    }
} While ($TRUE) 
}
#=================================
# MENU: MAILBOX PERMISSIONS
#=================================
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
    Switch (showMenuMailboxPermissions $menuMailboxPermissions "`tMailbox Permissions Menu" -clearScreen) {
        "1" { GetFullAccess }
        "2" { GetSendOnBehalfAccess }
        "3" { GetSendAsAccess }
        "M" { Return }
        Default { Write-Warning "Mailbox Permissions MENU: Invalid Choice. Try again.";sleep -seconds 1 }
    }
} While ($TRUE) 
}
#=================================
# MENU: EXCHANGE SYSTEM PROPERTIES
#=================================
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
    1 Active Directory & Exchange Schema Versions
    2 Exchange Server(s) Admin Display Version(s)
    3 $unAss
    M Main Menu

    Select a task by number or M
"@
Do {
    Switch (showMenuExchange $menuExchange "`tExchange Properties Menu" -clearScreen) {
        "1" { GetExchangeSchemaVerions }
        "2" { GetExchangeServerNamesADV }
        "3" { Write-Host $unAss -fore Yellow -back blue; sleep -seconds 1 }
        "M" { Return }
        Default { Write-Warning "Exchange MENU: Invalid Choice. Try again.";sleep -seconds 1 }
    }
} While ($TRUE)
}
#=================================
# MENU: MAIN APPLICATION MENU
#=================================
Function thinkMenuMain {
    Function showMenuMain {
        Param(
        [Parameter(Position=0,Mandatory=$TRUE,HelpMessage="Main Menu Help text")] [ValidateNotNullOrEmpty()] [string]$menuMain,
        [Parameter(Position=1)] [ValidateNotNullOrEmpty()] [string]$TitleMain="menuMain" ,
        [switch]$clearScreen
        )
        if ($clearScreen) {Clear-Host}
        Write-Host "`n`t$xAppName`n" -Fore Cyan
        $menuMainPrompt = $titleMain
        $menuMainPrompt += "`n`t"
        $menuMainPrompt += "="*$titleMain.Length
        $menuMainPrompt += "`n"
        $menuMainPrompt += $menuMain
        Read-Host -Prompt $menuMainPrompt
    }
$menuMain=@"
    1 Exchange System Properties
    2 Mailbox Properties
    3 Mailbox Statistics
    4 Special Delivery
    Q Quit

    Select a task by number or Q to quit
"@
    Do {
        Switch (showMenuMain $menuMain "`tMain Menu" -clearScreen) {
            "1" { thinkMenuExchange }
            "2" { thinkMenuMailboxPermissions }
            "3" { thinkMailboxStats }
            "4" { thinkSpecialDelivery }
            "Q" { Write-Host "
╦ ╦┌─┐┬  ┬┌─┐  ┌─┐  ┌┐┌┬┌─┐┌─┐  ┌┬┐┌─┐┬ ┬
╠═╣├─┤└┐┌┘├┤   ├─┤  │││││  ├┤    ││├─┤└┬┘
╩ ╩┴ ┴ └┘ └─┘  ┴ ┴  ┘└┘┴└─┘└─┘  ─┴┘┴ ┴ ┴o
            " -fore Yellow; sleep -milliseconds 300; 
                  Write-Host "
                                                                                                                              
                                     _/_/_/                            _/  _/                                     
                                  _/          _/_/      _/_/      _/_/_/  _/_/_/    _/    _/    _/_/              
                                 _/  _/_/  _/    _/  _/    _/  _/    _/  _/    _/  _/    _/  _/_/_/_/             
                                _/    _/  _/    _/  _/    _/  _/    _/  _/    _/  _/    _/  _/                    
                                 _/_/_/    _/_/      _/_/      _/_/_/  _/_/_/      _/_/_/    _/_/_/  _/  _/  _/   
                                                                                      _/                          
                                                                                     _/_/                             

            " -fore cyan;sleep -milliseconds 300;
                  Return }
            Default { Write-Warning "MAIN MENU: Invalid Choice. Try again.";sleep -seconds 1 }
        }
    } While ($TRUE)
}
#The.End...Have.A.Nice.Day...The.End...Have.A.Nice.Day...The.End...Have.A.Nice.Day...
clear
write-host @"

                888888b.         d8888 888b     d888 888 
                888  "88b       d88888 8888b   d8888 888 
                888  .88P      d88P888 88888b.d88888 888 
                8888888K.     d88P 888 888Y88888P888 888 
                888  "Y88b   d88P  888 888 Y888P 888 888 
                888    888  d88P   888 888  Y8P  888 Y8P 
                888   d88P d8888888888 888   "   888  "  
                8888888P" d88P     888 888       888 888 

                $xAppName
                ==============================================
                 Created on:   $createdOn
                 Created by:   $createdBy
                 Organization: $company
                ==============================================
"@ -fore Green;start-sleep -seconds 3
thinkMenuMain

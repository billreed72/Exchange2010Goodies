#========================================================================
# Created on:   Feb 6, 2014 13:00:00 EST
# Created by:   William Reed, wreed@appirio.com
# Organization: Appirio, Inc.
# Usage:	./BAMforColumbia.ps1
#========================================================================
$xAppName = "BAM! Version 1.0 (Bill's Application Manager) `n`tColumbia University build"
$unAss = "***[Unassigned]***"
$BamLogName = "BAMex"
$BamLogSource = "BAMSource"
If (!((Get-EventLog -List | Select-Object "Log") -match $BamLogName)) {new-EventLog -LogName $BamLogName -Source $BamLogSource}
#======================================
# FUNCTION: Get Mailbox Folder Total Message Counts & Sizes
#======================================
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
#======================================
# FUNCTION: Get Mailbox Total Message Count & Size
#======================================
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
#======================================
# FUNCTION: Exchange Schema Versions
#======================================
function GetExchangeSchemaVerions {
    Import-Module ActiveDirectory
    $OutEXVdata = @()
    $ExForestAndDomain = Read-Host 'Please enter your forest and domain (i.e. DC=dev10,DC=net)'
    $ExOrg = Read-Host 'Please enter the Exchange Org Name (i.e. First Organization)'
    write-EventLog -LogName $BamLogName -EventID 11 -Message "Results: Exchange Schema Versions saved: Started." -Source $BamLogSource -EntryType Information
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
    write-EventLog -LogName $BamLogName -EventID 12 -Message "Results: Exchange Schema Versions saved: [$SavePathEXVdata]." -Source $BamLogSource -EntryType Information
}
#======================================
# FUNCTION: Exchange Server Names and Versions
#======================================
function GetExchangeServerNamesADV {
    Import-Module ActiveDirectory
    $ExchangeServerData = @()
    write-EventLog -LogName $BamLogName -EventID 21 -Message "Results: Exchange Server Names and Versions saved: Started." -Source $BamLogSource -EntryType Information
    $AdminDisplayVersion = get-exchangeServer | select *
    $OutDVer = "" | select Name,ADV
    $OutDVer.Name = $AdminDisplayVersion.Name
    $OutDVer.ADV = $AdminDisplayVersion.AdminDisplayVersion
    $ExchangeServerData += $OutDVer
    $SavePathExServerData = ('ExchangeServers-{1:yyyyMMddHHmm}.csv' -f $env:COMPUTERNAME,(Get-Date))
    $ExchangeServerData | Export-csv  -Path $SavePathExServerData
    Write-Host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
    Write-Host $SavePathExServerData -Fore DarkRed -Back gray;start-sleep -seconds 1
    write-EventLog -LogName $BamLogName -EventID 22 -Message "Results: Exchange Server Names and Versions saved: [$SavePathExServerData]." -Source $BamLogSource -EntryType Information
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
    write-EventLog -LogName $BamLogName -EventID 31 -Message "Results: Get Full Access Permission saved: Started." -Source $BamLogSource -EntryType Information
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
    write-EventLog -LogName $BamLogName -EventID 32 -Message "Results: Get Full Access Permission saved: [$SavePathFAdata]." -Source $BamLogSource -EntryType Information
}
#======================================
# FUNCTION: Get Send On Behalf Access Permissions
#======================================
function GetSendOnBehalfAccess {
    Write-Host 'INPUT filename:' -Fore Cyan -Back DarkBlue;
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
    $OutSOBData = @()
    $UserList = Get-Content $UserListFile
    $CurProcMbxSOB = 1
    write-EventLog -LogName $BamLogName -EventID 41 -Message "Results: Get Send On Behalf Access Permissions saved: Started." -Source $BamLogSource -EntryType Information
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
    write-EventLog -LogName $BamLogName -EventID 42 -Message "Results: Get Send On Behalf Access Permissions saved: [$SavePathSOBdata]." -Source $BamLogSource -EntryType Information
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
    write-EventLog -LogName $BamLogName -EventID 51 -Message "Results: Get Send As Access Permissions saved: Started." -Source $BamLogSource -EntryType Information
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
    write-EventLog -LogName $BamLogName -EventID 52 -Message "Results: Get Send As Access Permissions saved: [$SavePathSAdata]." -Source $BamLogSource -EntryType Information
}

    #==============================================================================
    # MENU: Query Mailbox Statistics
    #==============================================================================
Function thinkMailboxStats {
Function showMenuMailboxStats {
    Param(
    [Parameter(Position=0,Mandatory=$TRUE,HelpMessage="Enter your Special Delivery menu text")] [ValidateNotNullOrEmpty()] [string]$menuMailboxStats,
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
    #==============================================================================
$menuMailboxStats=@"
    1 Get Mailbox Total Message Counts & Sizes
    2 Get Mailbox Folder Total Message Counts & Sizes
    3 $unAss
    M Main Menu

    Select a task by number or M
"@
Do {
    Switch (showMenuMailboxStats $menuMailboxStats "`tMy Special Delivery Tasks" -clearscreen) {
        "1" {Measure-Command{GetMailboxMsgCountsAndSize};start-sleep 3}
        "2" {Measure-Command{GetMailboxFolderMsgCountsAndSize};start-sleep 3}
        "3" {Write-Host $unAss -fore Green; sleep -seconds 1 }
        "M" { Return }
        Default {Write-Warning "MailboxStats MENU: Invalid Choice. Try again.";sleep -milliseconds 750}
    }
} While ($TRUE) 
}
    #==============================================================================
    #==============================================================================

    #==============================================================================
    # MENU: MAILBOX PERMISSIONS
    #==============================================================================
Function thinkMenumailboxPermissions {
Function showMenuMailboxPermissions {
  
    Param(
    [Parameter(Position=0,Mandatory=$TRUE,HelpMessage="Enter your Mailbox Permissions menu text")] [ValidateNotNullOrEmpty()] [string]$menuMailboxPermissions,
    [Parameter(Position=1)] [ValidateNotNullOrEmpty()] [string]$TitleMailboxPermissions="menuMailboxPermissions" ,
    [switch]$clearScreen
    )
    if ($clearScreen) {Clear-Host}
        Write-Host "`n`t$xAppName`n" -Fore Magenta
        $menuMailboxPermissionsPrompt=$titleMailboxPermissions
        $menuMailboxPermissionsPrompt+="`n`t"
        $menuMailboxPermissionsPrompt+="-"*$titleMailboxPermissions.Length
        $menuMailboxPermissionsPrompt+="`n"
        $menuMailboxPermissionsPrompt+=$menuMailboxPermissions
        Read-Host -Prompt $menuMailboxPermissionsPrompt
}
    #==============================================================================
$menuMailboxPermissions=@"
    1 Query for Full Access
    2 Query for Send On Behlaf Access
    3 Query for Send As Access
    M Main Menu

    Select a task by number or M
"@
Do {
    Switch (showMenuMailboxPermissions $menuMailboxPermissions "`tMy Mailbox Permissions Tasks" -clearscreen) {
        "1" {Measure-Command{GetFullAccess};start-sleep 3}
        "2" {Measure-Command{GetSendOnBehalfAccess};start-sleep 3}
        "3" {Measure-Command{GetSendAsAccess};start-sleep 3}
        "M" { Return }
        Default {Write-Warning "Mailbox Permissions MENU: Invalid Choice. Try again.";sleep -milliseconds 750}
    }
} While ($TRUE) 
}
    #==============================================================================
    #==============================================================================

    #==============================================================================
    # MENU: EXCHANGE SYSTEM PROPERTIES
    #==============================================================================
Function thinkMenuExchange {
Function showMenuExchange {
  
    Param(
    [Parameter(Position=0,Mandatory=$TRUE,HelpMessage="Enter your Exchange menu text")] [ValidateNotNullOrEmpty()] [string]$menuExchange,
    [Parameter(Position=1)] [ValidateNotNullOrEmpty()] [string]$TitleExchange="menuExchange" ,
    [switch]$clearScreen
    )
    if ($clearScreen) {Clear-Host}
        Write-Host "`n`t$xAppName`n" -Fore Magenta
        $menuExchangePrompt=$titleExchange
        $menuExchangePrompt+="`n`t"
        $menuExchangePrompt+="-"*$titleExchange.Length
        $menuExchangePrompt+="`n"
        $menuExchangePrompt+=$menuExchange
        Read-Host -Prompt $menuExchangePrompt
}
    #==============================================================================
$menuExchange=@"
    1 GetExchangeSchemaVerions
    2 GetExchangeServerNamesADV
    3 $unAss
    M Main Menu

    Select a task by number or M
"@
Do {
    Switch (showMenuExchange $menuExchange "`tMy Exchange Tasks" -clearscreen) {
        "1" {Measure-Command{GetExchangeSchemaVerions};start-sleep 3}
        "2" {Measure-Command{GetExchangeServerNamesADV};start-sleep 3}
        "3" {Write-Host $unAss -fore Green; sleep -seconds 1 }
        "M" { Return }
        Default {Write-Warning "Exchange MENU: Invalid Choice. Try again.";sleep -milliseconds 750}
    }
} While ($TRUE) 
}
    #==============================================================================
    # MENU: MAIN APPLICATION MENU
    #==============================================================================
Function thinkMenuMain {
Function showMenuMain {
    Param(
    [Parameter(Position=0,Mandatory=$TRUE,HelpMessage="Enter your MAIN menu text")] [ValidateNotNullOrEmpty()] [string]$menuMain,
    [Parameter(Position=1)] [ValidateNotNullOrEmpty()] [string]$TitleMain="menuMain" ,
    [switch]$clearScreen
    )
    if ($clearScreen) {Clear-Host}
        Write-Host "`n`t$xAppName`n" -Fore Magenta
        $menuMainPrompt=$titleMain
        $menuMainPrompt+="`n`t"
        $menuMainPrompt+="-"*$titleMain.Length
        $menuMainPrompt+="`n"
        $menuMainPrompt+=$menuMain
        Read-Host -Prompt $menuMainPrompt
}
    #==============================================================================
$menuMain=@"
    1 Query Exchange Properties
    2 Query Mailbox Permissions
    3 Query Mailbox Statistics
    Q Quit

    Select a task by number or Q to quit
"@
Do {
    Switch (showMenuMain $menuMain "`tMy BAM! Tasks" -clearscreen) {
        "1" { thinkMenuExchange }
        "2" { thinkMenuMailboxPermissions }
        "3" { thinkMailboxStats }
        "Q" { Write-Host "Have a nice day..." -fore Yellow -back darkBlue; sleep -milliseconds 300; 
              Write-Host "Goodbye..." -fore cyan -back darkBlue;sleep -milliseconds 300; clear
              Return }
        Default {Write-Warning "MAIN MENU: Invalid Choice. Try again.";sleep -milliseconds 750}
    }
} While ($TRUE)
}
    #==============================================================================
    #==============================================================================
thinkMenuMain

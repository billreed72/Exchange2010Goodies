$xAppName = "BAM! (Bill's Application Manager) – Version 0.2"

    #==============================================================================
    # MENU: Special Delivery Menu
    #==============================================================================
Function thinkSpecialDelivery {
Function showMenuSpecialDelivery {
  
    Param(
    [Parameter(Position=0,Mandatory=$TRUE,HelpMessage="Enter your Special Delivery menu text")] [ValidateNotNullOrEmpty()] [string]$menuSpecialDelivery,
    [Parameter(Position=1)] [ValidateNotNullOrEmpty()] [string]$TitleSpecialDelivery="menuSpecialDelivery" ,
    [switch]$clearScreen
    )
    if ($clearScreen) {Clear-Host}
        Write-Host "`n`t$xAppName`n" -Fore Magenta
        $menuSpecialDeliveryPrompt=$titleSpecialDelivery
        $menuSpecialDeliveryPrompt+="`n`t"
        $menuSpecialDeliveryPrompt+="-"*$titleSpecialDelivery.Length
        $menuSpecialDeliveryPrompt+="`n"
        $menuSpecialDeliveryPrompt+=$menuSpecialDelivery
        Read-Host -Prompt $menuSpecialDeliveryPrompt
}
    #==============================================================================
$menuSpecialDelivery=@"
    1 SpecialDelivery 1
    2 SpecialDelivery 2
    3 SpecialDelivery 3
    M Main Menu

    Select a task by number or M
"@
Do {
    Switch (showMenuSpecialDelivery $menuSpecialDelivery "`tMy Special Delivery Tasks" -clearscreen) {
        "1" {Write-Host "SpecialDelivery 1" -fore Green; sleep -seconds 1 }
        "2" {Write-Host "SpecialDelivery 2" -fore Green; sleep -seconds 1 }
        "3" {Write-Host "SpecialDelivery 3" -fore Green; sleep -seconds 1 }
        "M" { Return }
        Default {Write-Warning "SpecialDelivery MENU: Invalid Choice. Try again.";sleep -milliseconds 750}
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
    1 MailboxPermissions 1
    2 MailboxPermissions 2
    3 MailboxPermissions 3
    M Main Menu

    Select a task by number or M
"@
Do {
    Switch (showMenuMailboxPermissions $menuMailboxPermissions "`tMy Mailbox Permissions Tasks" -clearscreen) {
        "1" {Write-Host "MailboxPermissions 1" -fore Green; sleep -seconds 1 }
        "2" {Write-Host "MailboxPermissions 2" -fore Green; sleep -seconds 1 }
        "3" {Write-Host "MailboxPermissions 3" -fore Green; sleep -seconds 1 }
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
    3 ***tbd***[unassigned]
    M Main Menu

    Select a task by number or M
"@
Do {
    Switch (showMenuExchange $menuExchange "`tMy Exchange Tasks" -clearscreen) {
        "1" {Write-Host "GetExchangeSchemaVerions" -fore Green; sleep -seconds 1 }
        "2" {Write-Host "GetExchangeServerNamesADV" -fore Green; sleep -seconds 1 }
        "3" {Write-Host "***tbd***[unassigned]" -fore Green; sleep -seconds 1 }
        "M" { Write-Host "Laterz..." -fore Cyan; Return }
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
    2 Query Multiple Mailbox Properties
    3 Special Delivery
    Q Quit

    Select a task by number or Q to quit
"@
Do {
    Switch (showMenuMain $menuMain "`tMy BAM! Tasks" -clearscreen) {
        "1" { thinkMenuExchange }
        "2" { thinkMenuMailboxPermissions }
        "3" { thinkSpecialDelivery }
        "Q" { Write-Host "Have a nice day..." -fore Yellow -back darkBlue; sleep -seconds 2; 
              Write-Host "Goodbye..." -fore cyan -back darkBlue;sleep -seconds 1; 
              Return }
        Default {Write-Warning "MAIN MENU: Invalid Choice. Try again.";sleep -milliseconds 750}
    }
} While ($TRUE)
}
    #==============================================================================
    #==============================================================================
thinkMenuMain

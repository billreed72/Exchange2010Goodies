new-Eventlog -LogName "BAMex" -Source "BAMSource"
function createDistro {
    write-EventLog -LogName "BAMex" -EventID 666 -Message "Delivery Restriction setup started." -Source "BAMSource" -EntryType Information
    $distroName = @()
    $distroName=Read-Host "New Distro Name (i.e. NoDelivery)"
    $distroOU=Read-Host "(i.e. dev10.net/Restricted Delivery)"
    new-DistributionGroup -Name $distroName -OrganizationalUnit $distroOU -SamAccountName $distroName -Alias $distroName | Out-Null
    set-Group -Identity $distroName -Notes "Created by BAMex!"
    write-EventLog -LogName "BAMex" -EventID 666 -Message "Distro [$distroName] created in Organizational Unit [$distroOU]." -Source "BAMSource" -EntryType Information
}
function createTransRule {
    $tRuleName=Read-Host "Enter a Transportation Rule Name"
    $tRuleFromMemberOf=(get-distributionGroup $distroName).primarySmtpAddress
    $tRuleRejectMessage="You are no longer authorized to send email from this system."
    $tRuleRejectMessageStatusCode="5.7.1"
    New-TransportRule -Name $tRuleName -Comments '' -Priority '0' -Enabled $true -FromMemberOf $tRuleFromMemberOf -RejectMessageReasonText $tRuleRejectMessage -RejectMessageEnhancedStatusCode $tRuleRejectMessageStatusCode | Out-Null
    write-EventLog -LogName "BAMex" -EventID 666 -Message "Transportation rule [$tRuleName] created." -Source "BAMSource" -EntryType Information
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
addMembersToDistro
write-host "BAM! I'M DONE!"
write-EventLog -LogName "BAMex" -EventID 666 -Message "Delivery Restriction setup completed." -Source "BAMSource" -EntryType Information

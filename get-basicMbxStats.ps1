function GetMailboxStats {
    Write-Host 'INPUT filename.' -Fore Cyan -Back DarkBlue;
    $UserListFile = Read-Host '(i.e. c:\userList.csv or userList.csv)';
    $OutStatData = @()
    $UserList = Get-Content $UserListFile
    $CurProcMbxStat = 1
    foreach ($UserID in $UserList) {
        Write-Host -NoNewLine $CurProcMbxStat -Fore Blue -Back White; write-host '.' -Fore Red -Back White -NoNewLine
        $StatList = @()
        $StatList += $UserID
        foreach ($USerID in $StatList) {
            If ($UserID -ne $NULL) {
#            $OutStatObject = "" | select DisplayName,Mailbox,ItemCount,TotalItemSize,LastLogonTime,Server,OriginatingServer,Database,ObjectClass
#            $OutStatObject = "" | select DisplayName,Mailbox,ItemCount,TotalItemSize
            $OutStatObject = "" | select Mailbox,ItemCount,TotalItemSize
#            $OutStatObject.DisplayName = (Get-MailboxStatistics -Identity $UserID).DisplayName
            $OutStatObject.Mailbox = $UserID
            $OutStatObject.ItemCount = (Get-MailboxStatistics -Identity $UserID).ItemCount
            $OutStatObject.TotalItemSize = (Get-MailboxStatistics -Identity $UserID).TotalItemSize
 #           $OutStatObject.LastLogonTime = (Get-MailboxStatistics -Identity $UserID).LastLogonTime
 #           $OutStatObject.Server = (Get-MailboxStatistics -Identity $UserID).Server
 #           $OutStatObject.OriginatingServer = (Get-MailboxStatistics -Identity $UserID).OriginatingServer
 #           $OutStatObject.Database = (Get-MailboxStatistics -Identity $UserID).Database
 #           $OutStatObject.ObjectClass = (Get-MailboxStatistics -Identity $UserID).ObjectClass
            $OutStatData += $OutStatObject
            $OutObject
            }
        $CurProcMbxStat++
        }
    }
    $SavePathStatdata = ('MailboxStats-{1:yyyyMMddHHmmss}.csv' -f $env:COMPUTERNAME,(Get-Date))
    $OutStatData | Export-csv  -Path $SavePathStatdata
    Write-Host 'Results saved: ' -Fore Yellow -Back Blue -NoNewLine;
    Write-Host $SavePathStatdata -Fore DarkRed -Back gray;start-sleep -seconds 1
}
Measure-Command { GetMailboxStats }

<#
Running with only Mailbox, ItemCount, TotalItemSize

Days              : 0
Hours             : 0
Minutes           : 0
Seconds           : 40
Milliseconds      : 858
Ticks             : 408584389
TotalDays         : 0.00047289859837963
TotalHours        : 0.0113495663611111
TotalMinutes      : 0.680973981666667
TotalSeconds      : 40.8584389
TotalMilliseconds : 40858.4389

Running with selected: DisplayName,Mailbox,ItemCount,TotalItemSize

Days              : 0
Hours             : 0
Minutes           : 0
Seconds           : 44
Milliseconds      : 125
Ticks             : 441255980
TotalDays         : 0.000510712939814815
TotalHours        : 0.0122571105555556
TotalMinutes      : 0.735426633333333
TotalSeconds      : 44.125598
TotalMilliseconds : 44125.598


Running with selected: DisplayName,Mailbox,ItemCount,TotalItemSize,LastLogonTime,Server,OriginatingServer,Database,ObjectClass

Days              : 0
Hours             : 0
Minutes           : 2
Seconds           : 7
Milliseconds      : 206
Ticks             : 1272068099
TotalDays         : 0.00147230104050926
TotalHours        : 0.0353352249722222
TotalMinutes      : 2.12011349833333
TotalSeconds      : 127.2068099
TotalMilliseconds : 127206.8099

#>

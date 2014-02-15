$End = Get-Date
$xStart = Read-Host "How many days back?"
$Start = $End.AddDays(-$xStart)
[Int64] $intSent = $intRec = 0
[Int64] $intSentSize = $intRecSize = 0
[String] $strEmails = $null
# Use this line for testing in production against only 2 Hts
#$hts = get-exchangeserver |? {$_.serverrole -match "hubtransport" -and $_.name -like "*Hub02"} |% {$_.name}
# Use this line for running against all production Hubs
$hts = get-exchangeserver |? {$_.serverrole -match "hubtransport"} |% {$_.name}
Write-Host "DayOfWeek,Date,Sent,Sent Size,Received,Received Size" -ForegroundColor Yellow -BackgroundColor DarkBlue

# Start building the variable that will hold the information for the day
Do {
    $strEmails = "$($Start.DayOfWeek),$($Start.ToShortDateString()),"
    $intSent = $intRec = 0
    $hts | Get-MessageTrackingLog -ResultSize Unlimited -Start $Start -End $End | ForEach {
    # Sent E-mails
    If ($_.EventId -eq "RECEIVE" -and $_.Source -eq "STOREDRIVER") {
        $intSent++
        $intSentSize += $_.TotalBytes
    }
    # Received E-mails
    If ($_.EventId -eq "DELIVER") {
        $intRec++
        $intRecSize += $_.TotalBytes
    }
}
$intSentSize = [Math]::Round($intSentSize/1MB, 0)
$intRecSize = [Math]::Round($intRecSize/1MB, 0)
# Add the numbers to the $strEmails variable and print the result for the day
$strEmails += "$intSent,$intSentSize,$intRec,$intRecSize"
$strEmails >> DailyStats.txt
# Increment the Start and End by one day
$Start = $Start.AddDays(1)
$End = $Start.AddDays(1)
}
While ($End -lt (Get-Date))

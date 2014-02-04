#
# create a local event log named emailstats
# Example: new-Eventlog -LogName "emailstats" -Source "emailstatsSource"
$hts = get-exchangeserver |? {$_.serverrole -match "hubtransport"} |% {$_.name}
$startdate = "01/27/2014 12:01:01 AM"
$enddate = "02/02/2014 11:59:50 PM"

function time_pipeline {
param ($increment  = 1000)
begin{$i=0;$timer = [diagnostics.stopwatch]::startnew()}
process {
    $i++
    if (!($i % $increment)){Write-host “`rProcessed $i in $($timer.elapsed.totalseconds) seconds” -nonewline}
    $_
    }
end {
	write-host “`rProcessed $i log records in $($timer.elapsed.totalseconds) seconds”
	Write-Host "   Average rate: $([int]($i/$timer.elapsed.totalseconds)) log recs/sec."
	}
}

function miniemailstats {
get-messagetrackinglog -Server $ht -EventID "DELIVER" -Start $startdate -End $enddate -resultsize unlimited | time_pipeline | %{$count++}
write-host "Server $ht Message count " $count
write-EventLog -LogName "emailstats" -EventID 666 -Message "Email Server $ht Stat Count $count number of 

messages between dates $startdate and $enddate " -Source "emailstats" -EntryType Information
}

foreach ($ht in $hts){
$count=0
miniemailstats
}

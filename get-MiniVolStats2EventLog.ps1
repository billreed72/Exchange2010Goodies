$hts = get-exchangeserver |? {$_.serverrole -match "hubtransport"} |% {$_.name}
$startdate = Read-Host "Start Date"
$enddate = Read-Host "End Date"
function time_pipeline {
param ($increment  = 1000)
begin{$i=0;$timer = [diagnostics.stopwatch]::startnew()}
process {
    $i++
    if (!($i % $increment)){Write-host “Processed $i in $($timer.elapsed.totalseconds) seconds” -nonewline}
    $_
    }
end {
	write-host “Processed $i log records in $($timer.elapsed.totalseconds) seconds”
	Write-Host "Average rate: $([int]($i/$timer.elapsed.totalseconds)) log recs/sec."
    write-EventLog -LogName "BAMex" -Message “`rProcessed $i log records in $($timer.elapsed.totalseconds) seconds
Average rate: $([int]($i/$timer.elapsed.totalseconds)) log recs/sec” -Source "BAMex" -EventID 666 -EntryType Information
	}
}
function miniemailstats {
get-messagetrackinglog -Server $ht -EventID "DELIVER" -Start $startdate -End $enddate -resultsize unlimited | time_pipeline | %{$count++}
write-host "Server:" $ht 
write-host "Messages Inbound:" $count
write-EventLog -LogName "BAMex" -EventID 666 -Message "Exchange Server: $ht
Messages: $count
Start Date: $startdate
End Date: $enddate " -Source "BAMex" -EntryType Information
}
foreach ($ht in $hts){
$count=0
miniemailstats
}

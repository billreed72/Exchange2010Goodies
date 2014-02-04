# create a local event log named emailstats
# Example: new-Eventlog -LogName "emailstats" -Source "emailstatsSource"
$hts = get-exchangeserver | ? {$_.serverrole -match "hubtransport"} | % {$_.name}
$startdate = "{0:yyyyMMdd}" -f (get-date).AddDays(-8)
$enddate = "{0:yyyyMMdd}" -f (get-date).AddDays(-1)

function time_pipeline {
	param ($increment  = 1000) begin { 
		$i = 0;
		$timer = [diagnostics.stopwatch]::startnew()
	}
	process {
    	$i++
    	if (!($i % $increment)) {
    		Write-Host “`rProcessed $i in $($timer.elapsed.totalseconds) seconds” -NoNewLine
    	}
    $_
    }
	end {
		Write-Host "`rProcessed $i log records in $($timer.elapsed.totalseconds) seconds"
		Write-Host "   Average rate: $([int]($i/$timer.elapsed.totalseconds)) log recs/sec."
		}
}

function miniemailstats {
	get-messagetrackinglog -Server $ht -EventID "DELIVER" -Start $startdate -End $enddate -resultsize unlimited | time_pipeline | %{$count++}
	Write-Host 'Server $ht Message count ' $count
	Write-EventLog -logName "emailstats" -eventID 666 -message "Email Server $ht Stat Count $count number of messages between dates $startdate and $enddate " -Source "emailstats" -EntryType Information
}

foreach ($ht in $hts) {
	$count=0
	miniemailstats
}

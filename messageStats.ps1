function msgstats {
    $startDate = read-host 'Start Date [mm/dd/yyyy]'
    $serverFQDN = read-host 'Server FQDN [dex.dev10.net]'
    get-messagetrackinglog -Server $serverFQDN -start $startDate -resultsize unlimited | where { ($_.eventID -eq 'DELIVER') -and ($_.source -eq 'STOREDRIVER') } | %{$count++}
    write-host '============================' -fore yellow -back darkBlue
    write-host 'Total Messages: ' -fore yellow -back darkBlue -NoNewLine; write-host "..." $count "..." -fore blue -back white
    write-host '============================' -fore yellow -back darkBlue
}
$count = 0
msgstats

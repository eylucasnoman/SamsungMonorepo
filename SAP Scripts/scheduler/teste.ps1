function getCountdownInSeconds($hour, $addDay) {
    $startDate = (GET-DATE)
    $endDate = (Get-Date).AddDays($addDay).Date.AddHours($hour)
    $seconds = [int](NEW-TIMESPAN –Start $startDate –End $endDate).TotalSeconds

    return $seconds
}

getCountdownInSeconds 5
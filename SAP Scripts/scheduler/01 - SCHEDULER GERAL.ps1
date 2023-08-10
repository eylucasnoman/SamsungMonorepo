function getCountdownInSeconds($hour, $addDay) {
    $startDate = (GET-DATE)
    $endDate = (Get-Date).AddDays($addDay).Date.AddHours($hour)
    $seconds = [int](NEW-TIMESPAN –Start $startDate –End $endDate).TotalSeconds

    return $seconds
}

function waitSeconds($seconds){
    while ($seconds -gt 0) {
        Start-Sleep -Seconds 1
        $seconds--
    }
}

$countDownToOutbound = getCountdownInSeconds 5 1
Write-Host "Contagem de $($countDownToOutbound) segundos para extração do Outbound."
waitSeconds $countDownToOutbound
Invoke-Item "C:\Users\seda.scm49\Documents\SAP Scripts\00 - SAP Login.vbs"

waitSeconds 30
Invoke-Item "C:\Users\seda.scm49\Documents\SAP Scripts\Sales\01 - SALES - OUTBOUND CE.vbs"

$countDownToOtherBases = getCountdownInSeconds 7 0
waitSeconds $countDownToOtherBases
Invoke-Item "C:\Users\seda.scm49\Documents\SAP Scripts\01 - Close SAP.ps1"

waitSeconds 30
Invoke-Item "C:\Users\seda.scm49\Documents\SAP Scripts\00 - SAP Login.vbs"

waitSeconds 30
Invoke-Item "C:\Users\seda.scm49\Documents\SAP Scripts\Sales\02 - SALES - DEVS, Estoque e SOL.vbs"

waitSeconds 300
Invoke-Item "C:\Users\seda.scm49\Documents\SAP Scripts\01 - Close SAP.ps1"
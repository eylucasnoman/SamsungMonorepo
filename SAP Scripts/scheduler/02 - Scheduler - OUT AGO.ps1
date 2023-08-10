$StartDate = (GET-DATE)
$EndDate = (Get-Date).Date.AddHours(13)
$CountDown = [int](NEW-TIMESPAN –Start $StartDate –End $EndDate).TotalSeconds

Write-Host "Contagem de $($CountDown) segundos iniciada."
while ($CountDown -gt 0) {
    # Write-Host $CountDown
    Start-Sleep -Seconds 1
    $CountDown--
}
Invoke-Item "C:\Users\seda.scm49\Documents\SAP Scripts\00 - SAP Login.vbs"

$CountDown = 30
# Write-Host "Starting countdown..."
while ($CountDown -gt 0) {
    # Write-Host $CountDown
    Start-Sleep -Seconds 1
    $CountDown--
}
Invoke-Item "C:\Users\seda.scm49\Documents\SAP Scripts\Sales\03 - SALES - OUTBOUND CE Ago.vbs"
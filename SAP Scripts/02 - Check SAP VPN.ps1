function waitSeconds($seconds){
   while ($seconds -gt 0) {
       Start-Sleep -Seconds 1
       $seconds--
   }
}

Invoke-Item "C:\Users\seda.scm49\Documents\SAP Scripts\00 - SAP Login.vbs"
waitSeconds 13
Invoke-Item "C:\Users\seda.scm49\Documents\SAP Scripts\Validation\VALIDACAO - Gerar SO LIST CE.vbs"
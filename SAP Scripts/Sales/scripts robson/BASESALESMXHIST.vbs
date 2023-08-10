If Not IsObject(application) Then
   Set WshShell = CreateObject("WScript.Shell")
   Set proc = WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
End If

WScript.Sleep 5000

Set WSHShell = Nothing
Set SapGui = GetObject("SAPGUI")
Set Applic = SapGui.GetScriptingEngine
Set connection = Applic.OpenConnection("01. [SEP] SET ERP(S/4HANA) PRD", True)
Set session = connection.Children(0)
Session.findById("wnd[0]").maximize

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "seda.scm31"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "areqhre634@"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 11
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "ZLLEJ50090"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").text = "HHP_PC"
session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.08.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.08.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 22 08 A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.08.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.08.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 22 08 B.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.09.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.09.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 22 09 A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.09.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.09.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 22 09 B.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.10.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.10.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 22 10 A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.10.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.10.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 22 10 B.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.11.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.11.10"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 22 11 A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.11.11"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.11.20"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 22 11 B.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.11.21"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.11.24"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 22 11 C.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.11.25"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.11.26"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 22 11 D.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.11.27"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.11.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 22 11 E.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.12.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.12.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 22 12 A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.12.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.12.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 22 12 B.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.01.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.01.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 01 A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.01.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.01.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 01 B.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.02.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.02.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 02 A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.02.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.02.28"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 02 B.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.03.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.03.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 03 A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.03.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.03.25"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 03 B.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.03.26"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.03.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 03 C.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.04.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.04.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 04 A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.04.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.04.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\BASES\OUT MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 04 B.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

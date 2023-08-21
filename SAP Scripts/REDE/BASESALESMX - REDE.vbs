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

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.07.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.07.06"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 07 A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.07.07"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.07.11"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 07 B.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.07.12"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.07.20"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 07 C.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.07.21"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.07.27"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 07 D.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.07.28"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.07.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 07 E.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.08.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.08.09"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 08 A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.08.10"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.08.14"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 08 B.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.08.15"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.08.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND MX 23 08 C.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 19
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/tbar[0]/okcd").text = "ZRMMJ310410"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").text = "HHP_PC"
session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtP_LAYOUT").text = "SALES_CE SAL"
session.findById("wnd[0]/usr/ctxtP_LAYOUT").setFocus
session.findById("wnd[0]/usr/ctxtP_LAYOUT").caretPosition = 12
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ESTOQUE IM.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/tbar[0]/okcd").text = "ZRSDD6A080"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").text = "HHP_PC"
session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtPA_VAR").text = "SALES_CE_SO"
session.findById("wnd[0]/usr/ctxtPA_VAR").setFocus
session.findById("wnd[0]/usr/ctxtPA_VAR").caretPosition = 11
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[0]/tbar[1]/btn[14]").press
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[0]/tbar[1]/btn[94]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\MX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "SO LIST IM.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

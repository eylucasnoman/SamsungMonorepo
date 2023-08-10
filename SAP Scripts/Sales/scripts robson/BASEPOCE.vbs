' If Not IsObject(application) Then
'    Set WshShell = CreateObject("WScript.Shell")
'    Set proc = WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
' End If

' WScript.Sleep 5000

' Set WSHShell = Nothing
' Set SapGui = GetObject("SAPGUI")
' Set Applic = SapGui.GetScriptingEngine
' Set connection = Applic.OpenConnection("01. [SEP] SET ERP(S/4HANA) PRD", True)
' Set session = connection.Children(0)
' Session.findById("wnd[0]").maximize

' If Not IsObject(application) Then
'    Set SapGuiAuto  = GetObject("SAPGUI")
'    Set application = SapGuiAuto.GetScriptingEngine
' End If
' If Not IsObject(connection) Then
'    Set connection = application.Children(0)
' End If
' If Not IsObject(session) Then
'    Set session    = connection.Children(0)
' End If
' If IsObject(WScript) Then
'    WScript.ConnectObject session,     "on"
'    WScript.ConnectObject application, "on"
' End If
' session.findById("wnd[0]").maximize
' session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "seda.scm31"
' session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "areqhre634@"
' session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
' session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 11
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[1]/tbar[0]/btn[0]").press

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
session.findById("wnd[1]/usr/txtV-LOW").text = "CE_DA_HME"
session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.09.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.09.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE PO"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_09.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.10.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.10.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE PO"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_10.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.11.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.11.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE PO"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_11_A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.11.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.11.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE PO"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_11_B.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.12.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.12.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE PO"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_12.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.01.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.01.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE PO"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_01.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.02.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.02.28"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE PO"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_02.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.03.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.03.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE PO"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_03.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.04.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.04.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE PO"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_04.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.05.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.05.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE PO"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_05 A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.05.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.05.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE PO"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_05 C.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.06.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.06.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE PO"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 9
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_06 A.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

' session.findById("wnd[0]/tbar[0]/okcd").text = "ZRMMJ310410"
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[0]/tbar[1]/btn[17]").press
' session.findById("wnd[1]/usr/txtV-LOW").text = "CE_DA_HME"
' session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
' session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
' session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
' session.findById("wnd[1]/tbar[0]/btn[8]").press
' session.findById("wnd[0]/usr/ctxtP_LAYOUT").text = "SALES_CE SAL"
' session.findById("wnd[0]/usr/ctxtP_LAYOUT").setFocus
' session.findById("wnd[0]/usr/ctxtP_LAYOUT").caretPosition = 12
' session.findById("wnd[0]/tbar[1]/btn[8]").press
' session.findById("wnd[0]/tbar[1]/btn[43]").press
' session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "STOCK.XLSX"
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
' session.findById("wnd[1]/tbar[0]/btn[11]").press
' session.findById("wnd[0]/tbar[0]/btn[3]").press
' session.findById("wnd[0]/tbar[0]/btn[3]").press

' session.findById("wnd[0]/tbar[0]/okcd").text = "ZRSDD6A080"
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
' session.findById("wnd[0]/tbar[1]/btn[17]").press
' session.findById("wnd[1]/usr/txtV-LOW").text = "CE_DA_HME"
' session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
' session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
' session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
' session.findById("wnd[1]/tbar[0]/btn[8]").press
' session.findById("wnd[0]/usr/ctxtPA_VAR").text = "SALES_CE"
' session.findById("wnd[0]/usr/ctxtPA_VAR").setFocus
' session.findById("wnd[0]/usr/ctxtPA_VAR").caretPosition = 8
' session.findById("wnd[0]/tbar[1]/btn[8]").press
' session.findById("wnd[0]/tbar[1]/btn[14]").press
' session.findById("wnd[0]/tbar[1]/btn[17]").press
' session.findById("wnd[0]/tbar[1]/btn[17]").press
' session.findById("wnd[0]/tbar[1]/btn[94]").press
' session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales"
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "SO LIST CE.XLSX"
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
' session.findById("wnd[1]/tbar[0]/btn[11]").press
' session.findById("wnd[0]/tbar[0]/btn[3]").press
' session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
' session.findById("wnd[0]/tbar[0]/btn[3]").press

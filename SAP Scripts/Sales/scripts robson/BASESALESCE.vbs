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
session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").setFocus
session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").caretPosition = 3
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").text = "CE_DA_HME"
session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.09.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.09.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_09 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.09.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.09.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_09 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.10.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.10.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_10 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.10.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.10.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_10 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.11.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.11.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_11 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.11.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.11.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_11 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.12.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.12.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_12 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.12.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.12.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_12 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.01.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.01.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_01 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.01.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.01.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_01 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.02.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.02.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_02 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.02.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.02.28"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_02 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.03.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.03.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_03 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.03.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.03.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_03 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.04.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.04.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_04 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.04.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.04.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_04 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.05.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.05.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_05 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.05.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.05.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_05 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.06.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.06.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "SCM CE/"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE SAP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_06 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3

' session.findById("wnd[0]/tbar[0]/okcd").text = "ZLLEJ50090"
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[0]/tbar[1]/btn[17]").press
' session.findById("wnd[1]/usr/txtENAME-LOW").text = "ODAIR.JR"
' session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
' session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
' session.findById("wnd[1]/tbar[0]/btn[8]").press
' session.findById("wnd[0]/usr/ctxtS_WAIST-LOW").text = "2023.06.01"
' session.findById("wnd[0]/usr/ctxtS_WAIST-HIGH").text = "2023.06.30"
' session.findById("wnd[0]/usr/ctxtP_VAR").text = "DEV MES"
' session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
' session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
' session.findById("wnd[0]/tbar[1]/btn[8]").press
' session.findById("wnd[0]").sendVKey 34
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "DEVOLUCAO SIT (CIMA).xlsx"
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 25
' session.findById("wnd[1]/tbar[0]/btn[11]").press
' session.findById("wnd[0]").sendVKey 3
' session.findById("wnd[0]").sendVKey 3

' session.findById("wnd[0]/tbar[0]/okcd").text = "ZLLEJ50090"
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[0]/tbar[1]/btn[17]").press
' session.findById("wnd[1]/usr/txtENAME-LOW").text = "ODAIR.JR"
' session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
' session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
' session.findById("wnd[1]/tbar[0]/btn[8]").press
' session.findById("wnd[0]/usr/ctxtS_WAIST-LOW").text = "2023.06.01"
' session.findById("wnd[0]/usr/ctxtS_WAIST-HIGH").text = "2023.06.30"
' session.findById("wnd[0]/usr/ctxtP_VAR").text = "SALES_CE"
' session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
' session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 8
' session.findById("wnd[0]/tbar[1]/btn[8]").press
' session.findById("wnd[0]").sendVKey 34
' session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE"
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "DEVOLUCAO WEEK (REAL + SIMBOLICA).xlsx"
' session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
' session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 43
' session.findById("wnd[1]/tbar[0]/btn[11]").press
' session.findById("wnd[0]").sendVKey 3
' session.findById("wnd[0]").sendVKey 3

' session.findById("wnd[0]/tbar[0]/okcd").text = "ZRLEJ56150"
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[0]/tbar[1]/btn[17]").press
' session.findById("wnd[1]/usr/txtV-LOW").text = "DEV_MES_SIT"
' session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
' session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
' session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
' session.findById("wnd[1]/tbar[0]/btn[8]").press
' session.findById("wnd[0]/usr/ctxtS_PODAT-LOW").text = "2023.06.01"
' session.findById("wnd[0]/usr/ctxtS_PODAT-HIGH").text = "2023.06.30"
' session.findById("wnd[0]/usr/ctxtS_PODAT-HIGH").setFocus
' session.findById("wnd[0]/usr/ctxtS_PODAT-HIGH").caretPosition = 10
' session.findById("wnd[0]/tbar[1]/btn[8]").press
' session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").pressToolbarButton "DOWN"
' session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE"
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "DEVOLUCAO SIT (BAIXO).xlsx"
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 26
' session.findById("wnd[1]/tbar[0]/btn[11]").press
' session.findById("wnd[0]").sendVKey 3
' session.findById("wnd[0]").sendVKey 3

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
' session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE"
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ESTOQUE CE.xlsx"
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
' session.findById("wnd[1]/tbar[0]/btn[11]").press
' session.findById("wnd[0]").sendVKey 3
' session.findById("wnd[0]").sendVKey 3

' session.findById("wnd[0]/tbar[0]/okcd").text = "ZRSDD6A080"
' session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
' session.findById("wnd[0]/tbar[1]/btn[17]").press
' session.findById("wnd[1]/usr/txtV-LOW").text = "CE_DA_HME"
' session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
' session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
' session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
' session.findById("wnd[1]/tbar[0]/btn[8]").press
' session.findById("wnd[0]/usr/ctxtPA_VAR").text = "SALES_CE_SO"
' session.findById("wnd[0]/usr/ctxtPA_VAR").setFocus
' session.findById("wnd[0]/usr/ctxtPA_VAR").caretPosition = 8
' session.findById("wnd[0]/tbar[1]/btn[8]").press
' session.findById("wnd[0]/tbar[1]/btn[14]").press
' session.findById("wnd[0]/tbar[1]/btn[17]").press
' session.findById("wnd[0]/tbar[1]/btn[17]").press
' session.findById("wnd[0]/tbar[1]/btn[94]").press
' session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm31\Desktop\Nova SALES\BASE"
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "SO LIST CE.xlsx"
' session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
' session.findById("wnd[1]/tbar[0]/btn[11]").press
' session.findById("wnd[0]/tbar[0]/btn[3]").press
' session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
' session.findById("wnd[0]/tbar[0]/btn[3]").press

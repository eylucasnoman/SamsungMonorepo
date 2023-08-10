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
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_09 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.09.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.09.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_09 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.10.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.10.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_10 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.10.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.10.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_10 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.11.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.11.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_11 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.11.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.11.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_11 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.12.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.12.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_12 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2022.12.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2022.12.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_22_12 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.01.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.01.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_01 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.01.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.01.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_01 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.02.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.02.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_02 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.02.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.02.28"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_02 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.03.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.03.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_03 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.03.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.03.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_03 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.04.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.04.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_04 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.04.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.04.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_04 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.05.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.05.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_05 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.05.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.05.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_05 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 24
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.06.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.06.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_06 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.06.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.06.30"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_06 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.07.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.07.15"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_07 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.07.16"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.07.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_07 B.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtS_BLDAT-LOW").text = "2023.08.01"
session.findById("wnd[0]/usr/ctxtS_BLDAT-HIGH").text = "2023.08.31"
session.findById("wnd[0]/usr/ctxtP_VAR").text = "/SCM CE"
session.findById("wnd[0]/usr/ctxtP_VAR").setFocus
session.findById("wnd[0]/usr/ctxtP_VAR").caretPosition = 7
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\OUTBOUNDS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "OUTBOUND_CE_23_08 A.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 22
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]").sendVKey 3
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
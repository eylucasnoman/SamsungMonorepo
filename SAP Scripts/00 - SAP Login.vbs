If Not IsObject(application) Then
   Set WshShell = CreateObject("WScript.Shell")
   Set proc = WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
End If
WScript.Sleep 5000
Set WSHShell = Nothing

Set SapGui = GetObject("SAPGUI")
Set Applic = SapGui.GetScriptingEngine
Set connection = Applic.OpenConnection("01. SET ERP PRD", True)
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

' Login no SAP
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "seda.scm49"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "958Y&8$Wk9Vspon$5^$MF"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 21
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
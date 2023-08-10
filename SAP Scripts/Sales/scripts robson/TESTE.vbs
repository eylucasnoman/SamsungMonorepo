If Not IsObject(application) Then
   Set WshShell = CreateObject("WScript.Shell")
   Set proc = WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
End If

WScript.Sleep 5000

Sub ConexaoSap()
    On Error Resume Next

    Dim SapGuiAuto, application, connection, session
    Dim strconexaosap
    
    Set SapGuiAuto = GetObject("SAPGUI")
    If Not TypeName(SapGuiAuto) = "CDispatch" Then
        Exit Sub
    End If
    
    strconexaosap = "01. [SEP] SET ERP(S/4HANA) PRD"
    MsgBox strconexaosap
    
    Set application = SapGuiAuto.GetScriptingEngine.OpenConnection(strconexaosap, True)
    If Not TypeName(application) = "CDispatch" Then
        Set SapGuiAuto = Nothing
        Exit Sub
    End If
    
    Set connection = application.Children(0)
    If Not TypeName(connection) = "CDispatch" Then
        Set application = Nothing
        Set SapGuiAuto = Nothing
        Exit Sub
    End If
    
    Set session = connection.Children(0)
    If Not TypeName(session) = "CDispatch" Then
        Set connection = Nothing
        Set application = Nothing
        Set SapGuiAuto = Nothing
        Exit Sub
    End If
    
    result = transacao(session, connection, application, SapGuiAuto)
    MsgBox "testes"
    
    On Error GoTo 0

    Set session = Nothing
    Set connection = Nothing
    Set application = Nothing
    Set SapGuiAuto = Nothing
End Sub

ConexaoSap()
transacao("","","", "")

Function transacao(session,connection,application, SapGuiAuto)
MsgBox "Passei"
End Function



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

' DEV BAIXO
session.findById("wnd[0]/tbar[0]/okcd").text = "ZRLEJ56150"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").text = "DEV_MES_SIT"
session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtS_PODAT-LOW").text = "2023.08.01"
session.findById("wnd[0]/usr/ctxtS_PODAT-HIGH").text = "2023.08.31"
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").pressToolbarButton "&MB_VARIANT"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").setCurrentCell 5,"TEXT"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "5"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").pressToolbarButton "DOWN"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\Bases"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "DEVOLUCAO SIT (BAIXO).xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 21
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press

' DEV CIMA
session.findById("wnd[0]/tbar[0]/okcd").text = "ZLLEJ50090"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "ODAIR.JR"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtS_WAIST-LOW").text = "2023.08.01"
session.findById("wnd[0]/usr/ctxtS_WAIST-HIGH").text = "2023.08.31"
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cmbG51_SCREEN-USPEC_LBOX").setFocus
wscript.sleep 500
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cmbG51_SCREEN-USPEC_LBOX").key = "X"
wscript.sleep 300
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\Bases"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "DEVOLUCAO SIT (CIMA).xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 14
session.findById("wnd[1]/tbar[0]/btn[11]").press

' DEV WEEK
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cmbG51_SCREEN-USPEC_LBOX").setFocus
wscript.sleep 1000
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cmbG51_SCREEN-USPEC_LBOX").key = "X"
wscript.sleep 1000
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").setCurrentCell 1,"TEXT"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]").sendVKey 34
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\Bases"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "DEVOLUCAO WEEK (REAL + SIMBOLICA).xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 33
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press

' Estoque
session.findById("wnd[0]/tbar[0]/okcd").text = "ZRMMJ310410"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").text = "CE_DA_HME"
session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press
WScript.Sleep 1000
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cmbG51_SCREEN-USPEC_LBOX").key = "X"
WScript.Sleep 1000
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\Bases"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ESTOQUE CE.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press

' SO List
session.findById("wnd[0]/tbar[0]/okcd").text = "ZRSDD6A080"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtV-LOW").text = "CE_DA_HME"
session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").text = "2023.01.01"
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").setFocus
session.findById("wnd[0]/usr/ctxtS_ERDAT-LOW").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cmbG51_SCREEN-USPEC_LBOX").setFocus
WScript.Sleep 500
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cmbG51_SCREEN-USPEC_LBOX").key = "X"
WScript.Sleep 300
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell").clickCurrentCell
On Error Resume Next
   For i = 0 To 10
      session.findById("wnd[0]/tbar[1]/btn[14]").press
   Next
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[0]/tbar[1]/btn[94]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\seda.scm49\Documents\Sales\Bases"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "SO LIST CE.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 15
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press

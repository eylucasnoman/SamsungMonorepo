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
session.findById("wnd[0]/tbar[0]/okcd").text = "ZLSDF50930"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/radPA_OPT2").setFocus
session.findById("wnd[0]/usr/radPA_OPT2").select

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "10"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "10"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 10 - TV ; AV; REF; WM; AC; SAC; VC"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D10_DC10.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "11"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "10"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 10 - TV ; AV; REF; WM; AC; SAC; VC"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D10_DC11.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "30"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "10"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 10 - TV ; AV; REF; WM; AC; SAC; VC"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D10_DC30.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "31"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "10"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 10 - TV ; AV; REF; WM; AC; SAC; VC"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D10_DC31.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3



session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "10"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "42"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 42 - MON"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D42_DC10.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "11"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "42"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 42 - MON"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D42_DC11.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "30"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "42"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 42 - MON"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D42_DC30.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "31"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "42"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 42 - MON"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D42_DC31.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3



session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "10"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "43"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 43 - NPC"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D43_DC10.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "11"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "43"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 43 - NPC"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D43_DC11.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "30"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "43"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 43 - NPC"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D43_DC30.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "31"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "43"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 43 - NPC"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D43_DC31.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3



session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "10"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "50"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 50 - HHP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D50_DC10.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "11"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "50"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 50 - HHP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D50_DC11.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "30"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "50"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 50 - HHP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D50_DC30.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "31"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "50"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 50 - HHP"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D50_DC31.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3



session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "10"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "61"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 61 - HME"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D61_DC10.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "11"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "61"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 61 - HME"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D61_DC11.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "30"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "61"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 61 - HME"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D61_DC30.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3

session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "C820"
session.findById("wnd[0]/usr/ctxtPA_VKORG").text = "8201"
session.findById("wnd[0]/usr/ctxtPA_VTWEG").text = "31"
session.findById("wnd[0]/usr/ctxtPA_SPART").text = "61"
session.findById("wnd[0]/usr/ctxtPA_VALID").setFocus
session.findById("wnd[0]/usr/ctxtPA_VALID").caretPosition = 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
session.findById("wnd[0]/usr/cntlG_CUSTOM_CONTAINER/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "W:\1_SALES_STATUS_HHP\BASES SAP\PUMI\Division 61 - HME"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Base_D61_DC31.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
session.findById("wnd[1]/tbar[0]/btn[11]").press
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3

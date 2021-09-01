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
session.findById("wnd[0]/tbar[0]/okcd").text = "lm01"
session.findById("wnd[0]").sendVKey 0

'this only shows up the first time you log into SAP.  Need to figure out how to check for that

session.findById("wnd[0]/usr/txtS2200_LGNUM_SCANNED").text = "491"
session.findById("wnd[0]/usr/txtS2200_LGNUM_SCANNED").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subEXIT0101:SAPLXLRF:2888/txtRLMOB-MENOPT").text = "3"
session.findById("wnd[0]/usr/subEXIT0101:SAPLXLRF:2888/txtRLMOB-MENOPT").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subEXIT0101:SAPLXLRF:2888/txtRLMOB-MENOPT").text = "2"
session.findById("wnd[0]/usr/subEXIT0101:SAPLXLRF:2888/txtRLMOB-MENOPT").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/txtS2100_MATNR_SCANNED").text = "5le23"
session.findById("wnd[0]/usr/txtS2100_SELECT").text = "1"
session.findById("wnd[0]/usr/txtS2100_SELECT").setFocus
session.findById("wnd[0]/usr/txtS2100_SELECT").caretPosition = 1
session.findById("wnd[0]/usr/btnRLMOB-PENTER").press

'this location is where the default should be

session.findById("wnd[0]/usr/txtS2200_LGPLA_SCANNED").text = "970701"
session.findById("wnd[0]/usr/txtS2200_LGPLA_SCANNED").caretPosition = 6
session.findById("wnd[0]/usr/btnRLMOB-PENTER").press

'first is max

session.findById("wnd[0]/usr/txtS2300_LPMAX_SCANNED").text = "6"

'second is min

session.findById("wnd[0]/usr/txtS2300_LPMIN_SCANNED").text = "6"
session.findById("wnd[0]/usr/txtS2300_LPMIN_SCANNED").setFocus
session.findById("wnd[0]/usr/txtS2300_LPMIN_SCANNED").caretPosition = 1
session.findById("wnd[0]/usr/btnRLMOB-PENTER").press

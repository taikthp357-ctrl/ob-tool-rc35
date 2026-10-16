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
session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB5").select
session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB5/ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1030/btn%_ST_EBELN_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB3").select
session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB3/ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1040/btn%_ST_MATNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/tbar[0]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[24]").press
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB3/ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1040/ctxtST_WERKS-LOW").text = "4007"
session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB3/ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1040/ctxtST_WERKS-LOW").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_ORDER_CRITERIA/tabpS0S_TAB3/ssub%_SUBSCREEN_ORDER_CRITERIA:RVV50R10C:1040/ctxtST_WERKS-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/chk[1,5]").selected = true
session.findById("wnd[0]/usr/txt[63,5]").text = "  100"
session.findById("wnd[0]/usr/txt[63,5]").setFocus
session.findById("wnd[0]/usr/txt[63,5]").caretPosition = 5
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/lbl[50,5]").setFocus
session.findById("wnd[0]/usr/lbl[50,5]").caretPosition = 5
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/lbl[50,2]").setFocus
session.findById("wnd[0]/usr/lbl[50,2]").caretPosition = 3
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/lbl[50,5]").setFocus
session.findById("wnd[0]/usr/lbl[50,5]").caretPosition = 5
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/lbl[50,7]").setFocus
session.findById("wnd[0]/usr/lbl[50,7]").caretPosition = 5
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/lbl[50,5]").setFocus
session.findById("wnd[0]/usr/lbl[50,5]").caretPosition = 5
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/lbl[63,2]").setFocus
session.findById("wnd[0]/usr/lbl[63,2]").caretPosition = 3
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/txt[63,5]").setFocus
session.findById("wnd[0]/usr/txt[63,5]").caretPosition = 5
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/chk[1,5]").selected = true
session.findById("wnd[0]/usr/txt[63,5]").text = " 100"
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/usr/lbl[50,5]").setFocus
session.findById("wnd[0]/usr/lbl[50,5]").caretPosition = 4
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/lbl[50,5]").setFocus
session.findById("wnd[0]/usr/lbl[50,5]").caretPosition = 3

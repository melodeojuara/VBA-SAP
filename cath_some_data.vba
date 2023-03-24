Sub teste()
       Set SapGui = GetObject("SAPGUI")
       Set Appl = SapGui.GetScriptingEngine

    If Not IsObject(Application) Then
       Set SapGui = GetObject("SAPGUI")
       Set Appl = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = Appl.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
        
    
    For Each instalacao In Range("A1:A126")
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "es32"
    
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 71
    session.findById("wnd[1]/usr/tabsSEARCHFIELDS/tabpTAB2").Select
    session.findById("wnd[1]/usr/tabsSEARCHFIELDS/tabpTAB2/ssubSUB2:SAPLEFND:0112/ctxtEFINDD-D_GERAET").Text = instalacao.Value
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]").sendVKey 12
    instalacao.Offset(0, 1).Value = session.findById("wnd[0]/usr/ctxtEANLD-ANLAGE").Text
    Next

End Sub

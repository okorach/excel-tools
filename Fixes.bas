Attribute VB_Name = "Fixes"
Public Sub FixAllSheets()
    Dim ws As Worksheet
    For Each ws In Worksheets
        Call FixWorksheet(ws)
        Call FixButtons(ws)
    Next ws
End Sub

Public Sub FixWorksheetActive()
    Call FixWorksheet(ActiveSheet)
End Sub
Public Sub FixWorksheet(ws As Worksheet)
    If Not IsAnAccount(ws) Then
        Exit Sub
    End If

    Dim i As Long
    Dim wsName As String
    wsName = Replace(Replace(Replace(LCase$(ws.name), " ", "_"), "é", "e"), "è", "e")
    For i = 1 To ws.ListObjects.Count
        Dim oName As String
        oName = LCase$(ws.ListObjects(i).name)
        If oName Like "*yield*" Or oName Like "*interest*" Then
            ws.ListObjects(i).name = wsName & "_" & INTEREST_TABLE_NAME
        ElseIf oName Like "*transaction*" Or oName Like "*balance*" Then
            ws.ListObjects(i).name = wsName & "_" & BALANCE_TABLE_NAME
        ElseIf oName Like "*deposit*" Or oName = wsName & "_" Then
            ws.ListObjects(i).name = wsName & "_" & DEPOSIT_TABLE_NAME
        End If
    Next i
End Sub

Public Sub FixButtonsActive()
    Call FixButtons(ActiveSheet)
End Sub
Public Sub FixButtons(ws As Worksheet)
    If Not IsAnAccount(ws) Then
        Exit Sub
    End If
    Dim s As Shape
    Dim sbw As Integer, lbw As Integer
    sbw = 40
    If ws.Shapes.Count <= 0 Then
        Exit Sub
    End If
    Dim sbw As Integer
    sbw = 40
    Dim i As Long
    i = 0
    Dim s As Shape

    For Each btnData In Array( _
        "BtnHome," & BTN_HOME_TEXT & ",Webdings,18,1,1,40" _
        , "BtnPrev5," & BTN_PREV_5_TEXT & ",Webdings,18,1,2,40" _
        , "BtnPrev," & BTN_PREV_TEXT & ",Webdings,18,1,3,40" _
        , "BtnNext," & BTN_NEXT_TEXT & ",Webdings,18,1,4,40" _
        , "BtnNext5," & BTN_NEXT_5_TEXT & ",Webdings,18,1,5,40" _
        , "BtnTop," & BTN_TOP_TEXT & ",Webdings,18,1,6,40" _
        , "BtnBottom," & BTN_BOTTOM_TEXT & ",Webdings,18,1,7,40" _
        , "BtnSort," & BTN_SORT_TEXT & ",Webdings,18,2,1,40" _
        , "BtnImport," & Chr$(71) & ",Webdings,18,2,2,40" _
        , "BtnAddEntry," & BTN_ADD_ROW_TEXT & ",Arial,14,2,3,40" _
        , "BtnInterest," & Chr$(143) & ",Webdings,18,2,4,40" _
        , "BtnFormat," & BTN_FORMAT_TEXT & ",Arial,18,2,5,80" _
        )
        values = Split(btnData, ",", -1, vbTextCompare)
        Set s = ShapeFind(ws, CStr(values(0)))
        If Not s Is Nothing Then
            Call BtnSetProperties(s, text:=CStr(values(1)), font:=CStr(values(2)), fontSize:=CInt(values(3)))
            Call ShapePlacement(s, BTN_HOME_X + (CInt(values(5)) - 1) * sbw, BTN_HOME_Y + (CInt(values(4)) - 1) * BTN_HEIGHT, CInt(values(6)) - 1, BTN_HEIGHT - 1)
        End If
    Next btnData
    ws.Range("A1").Select
End Sub


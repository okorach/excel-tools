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
    lbw = 100
    ws.Activate
    For Each s In ws.Shapes
        If s.name = "BtnHome" Then
            Call ShapePlacementXY(s, BTN_HOME_X, BTN_HOME_Y, BTN_HOME_X + sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="9", font:="Webdings", size:=18, action:="ThisWorkbook.GoToSolde")
        ElseIf s.name = "BtnPrev5" Then
            Call ShapePlacementXY(s, BTN_HOME_X + sbw, BTN_HOME_Y, BTN_HOME_X + 2 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="7", font:="Webdings", size:=18, action:="ThisWorkbook.GoBack5")
        ElseIf s.name = "BtnPrev" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 2 * sbw, BTN_HOME_Y, BTN_HOME_X + 3 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="3", font:="Webdings", size:=18, action:="ThisWorkbook.GoToPrev")
        ElseIf s.name = "BtnNext" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 3 * sbw, BTN_HOME_Y, BTN_HOME_X + 4 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="4", font:="Webdings", size:=18, action:="ThisWorkbook.GoToNext")
        ElseIf s.name = "BtnNext5" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 4 * sbw, BTN_HOME_Y, BTN_HOME_X + 5 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="8", font:="Webdings", size:=18, action:="ThisWorkbook.GoFwd5")
        ElseIf s.name = "BtnTop" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 5 * sbw, BTN_HOME_Y, BTN_HOME_X + 6 * sbw, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="5", font:="Webdings", size:=18, action:="scrollToTop")
        ElseIf s.name = "BtnBottom" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 6 * sbw, BTN_HOME_Y, BTN_HOME_X + 7 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="6", font:="Webdings", size:=18, action:="scrollToBottom")
        ElseIf s.name = "BtnSort" Then
            Call ShapePlacementXY(s, BTN_HOME_X, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="~", font:="Webdings", size:=18, action:="sortCurrentAccount")
        ElseIf s.name = "BtnImport" Then
            Call ShapePlacementXY(s, BTN_HOME_X + sbw, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 2 * sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:=Chr$(71), font:="Webdings", size:=18, action:="ImportAny")
        ElseIf s.name = "BtnAddEntry" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 2 * sbw, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 3 * sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="+1", font:="Arial", size:=14, action:="addSavingsRow")
        ElseIf s.name = "BtnInterests" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 3 * sbw, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 4 * sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:=Chr(143), font:="Webdings", size:=18, action:="btnAccountInterests")
        ElseIf s.name = "BtnFormat" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 4 * sbw, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 6 * sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="Format", font:="Arial", size:=12, action:="AccountFormatCurrent")
        End If
    Next s
End Sub

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
            ws.ListObjects(i).DisplayName = wsName & "_" & INTEREST_TABLE_NAME
        ElseIf oName Like "*transaction*" Or oName Like "*balance*" Then
            ws.ListObjects(i).name = wsName & "_" & BALANCE_TABLE_NAME
            ws.ListObjects(i).DisplayName = wsName & "_" & BALANCE_TABLE_NAME
        ElseIf oName Like "*deposit*" Or oName = wsName & "_" Then
            ws.ListObjects(i).name = wsName & "_" & DEPOSIT_TABLE_NAME
            ws.ListObjects(i).DisplayName = wsName & "_" & DEPOSIT_TABLE_NAME
        End If
    Next i
End Sub

Public Sub FixButtons(ws As Worksheet)
    ' For the future
End Sub

Private Sub setTechnicalSheetsVisibility(visibility As XlSheetVisibility)
    Dim isTechSheet As Boolean
    Dim ws As Worksheet
    Call FreezeDisplay
    For Each ws In Worksheets
        isTechSheet = True
        For Each protectedWsName In Array("Solde", "Solde par compte", "Interests", "Budget", "Comptes", "Paramètres")
            If ws.name = protectedWsName Then
                isTechSheet = False
            End If
            If Not isTechSheet Then Exit For
        Next protectedWsName
        If isTechSheet Then
            Dim oAccount As Account
            Set oAccount = LoadAccount(getAccountId(ws))
            If Not (oAccount Is Nothing) Then
                isTechSheet = False
            End If
        End If
        If isTechSheet Then
            ws.Visible = visibility
        End If
    Next ws
    Call UnfreezeDisplay
End Sub
Public Sub TechnicalSheetsShow()
    Call setTechnicalSheetsVisibility(xlSheetVisible)
End Sub

Public Sub TechnicalSheetsHide()
    Call setTechnicalSheetsVisibility(xlSheetHidden)
End Sub

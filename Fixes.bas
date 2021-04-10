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

End Sub



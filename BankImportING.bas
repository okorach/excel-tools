Attribute VB_Name = "BankImportING"
'------------------------------------------------------------------------------
' ING
'------------------------------------------------------------------------------

Public Sub ImportING(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True
    Dim modal As ProgressBar
    Set modal = NewProgressBar("Import ING in progress", GetLastNonEmptyRow() + 1)
    modal.Update
    
    Dim r As Long
    r = 1
    Do While LenB(Cells(r, 1).value) > 0
        oTable.ListRows.Add
        With oTable.ListRows(oTable.ListRows.Count)
            .Range(1, dateCol).value = Cells(r, 1).value
            .Range(1, amountCol).value = toAmount(Cells(r, 4).value)
            .Range(1, descCol).value = simplifyDescription(Cells(r, 2).value, subsTable)
        End With
        r = r + 1
        modal.Update
    Loop
    ActiveWorkbook.Close
    Set modal = Nothing
End Sub

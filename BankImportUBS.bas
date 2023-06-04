Attribute VB_Name = "BankImportUBS"
'------------------------------------------------------------------------------
' Import UBS
'------------------------------------------------------------------------------

Sub ImportUBS(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)
    Dim xlsFile As Variant
    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    If LCase$(Right$(fileToOpen, 4)) = ".csv" Then
        xlsFile = convertCsvToXls(fileToOpen)
        Workbooks.Open filename:=xlsFile, ReadOnly:=True
    Else
        Workbooks.Open filename:=fileToOpen, ReadOnly:=True
    End If
    Dim modal As ProgressBar
    Set modal = NewProgressBar("Import UBS in progress", GetLastNonEmptyRow())
    modal.Update
    
    Dim i As Long
    i = 2
    Do While LenB(Cells(i, 1).value) > 0
        oTable.ListRows.Add
        With oTable.ListRows(oTable.ListRows.Count)
            If Cells(i, 13) = "Solde prix prestations" Then
                .Range(1, amountCol).value = 0
            ElseIf LenB(Cells(i, 18).value) > 0 Then
                .Range(1, amountCol).value = toAmount(Cells(i, 18).value) ' Sous-montant column
            ElseIf LenB(Cells(i, 19).value) > 0 Then
                .Range(1, amountCol).value = -toAmount(Cells(i, 19).value) ' Debit column
            ElseIf LenB(Cells(i, 20).value) > 0 Then
                .Range(1, amountCol).value = toAmount(Cells(i, 20).value) ' Credit column
            Else
                .Range(1, amountCol).value = 0
            End If
            .Range(1, dateCol).value = CDate(DateValue(Replace(Cells(i, 12).value, ".", "/")))
            .Range(1, descCol).value = simplifyDescription(Cells(i, 13).value & " " & Cells(i, 14).value & " " & Cells(i, 15).value, subsTable)
        End With
        i = i + 1
        modal.Update
    Loop
    ActiveWorkbook.Close
    Set modal = Nothing
End Sub

Private Function convertCsvToXls(fileToOpen As Variant) As Variant
    ' Converts a CSV file in XLS to solve Unicode characters problems
    Dim xlsFile As Variant
    Workbooks.OpenText filename:=fileToOpen, _
        Origin:=65001, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
        , 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1)), _
        TrailingMinusNumbers:=True
    xlsFile = left$(fileToOpen, Len(fileToOpen) - 4) & Format$(Now(), "yyyy-MM-dd hh-mm-ss") & ".xls"
    ActiveWorkbook.SaveAs filename:=xlsFile, fileformat:=xlExcel8, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
    ActiveWorkbook.Close
    convertCsvToXls = xlsFile
End Function



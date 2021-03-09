Attribute VB_Name = "AccountImportExport"
Const MAX_IMPORT = 30000

Private Function toAmount(str) As Double
    If VarType(str) = vbString Then
        str = Replace(Replace(str, ",", "."), "'", "")
        toAmount = CDbl(str)
    Else
        toAmount = str
    End If
End Function


Private Function toMonth(str) As Long
    s = LCase$(Trim$(str))
    ' TODO handle accents in Fev and Dec
    If s Like "jan*" Then
        toMonth = 1
    ElseIf s Like "f?[bv]*" Then
        toMonth = 2
    ElseIf s Like "mar*" Then
        toMonth = 3
    ElseIf s Like "a*" Then
        toMonth = 4
    ElseIf s Like "mai*" Or s Like "may*" Then
        toMonth = 5
    ElseIf s Like "juin*" Or s Like "jun*" Then
        toMonth = 6
    ElseIf s Like "juil*" Or s Like "jul*" Then
        toMonth = 7
    ElseIf s Like "ao*" Or s Like "aug*" Then
        toMonth = 8
    ElseIf s Like "sep*" Then
        toMonth = 9
    ElseIf s Like "oct*" Then
        toMonth = 10
    ElseIf s Like "nov*" Then
        toMonth = 11
    ElseIf s Like "d?c*" Then
        toMonth = 12
    Else
        toMonth = 0
    End If
End Function

Private Function toDate(str) As Date
    A = Split(str, " ", -1, vbTextCompare)
    toDate = DateSerial(CInt(A(2)), toMonth(A(1)), CInt(A(0)))
End Function




'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportAny()

    Dim fileToOpen As Variant
    fileToOpen = Application.GetOpenFilename()
    If fileToOpen <> False Then
        Call FreezeDisplay
        Dim bank As String
        bank = Cells(3, 2).Value
        If (bank = "ING Direct") Then
            Call ImportING(ActiveSheet.ListObjects(1), fileToOpen)
        ElseIf (bank = "LCL") Then
            Call ImportLCL(ActiveSheet.ListObjects(1), fileToOpen)
        ElseIf (bank = "UBS") Then
            Call ImportUBS(fileToOpen)
        ElseIf (bank = "Revolut") Then
            Call ImportRevolut(fileToOpen)
        Else
            Call UnfreezeDisplay
            Call ErrorMessage("k.errorImportNotRecognized", "k.warningImportCancelled")
        End If
        Call UnfreezeDisplay
    Else
        Call ErrorMessage("k.warningImportCancelled")
    End If
End Sub


' PRLV SEPA CE URSSAF RHONE ALPES-CNTFS : FR55ZZZ143065 000828DC120181231145950A000136092 DE CE URSSAF RHONE ALPES-CNTFS : 000828DC120181231145950A000136092 FR55ZZZ143065
Private Function deleteDuplicateSepa(desc As String) As String
    Dim idstr As String
    idstr = "PRLV SEPA "
    deleteDuplicateSepa = desc
    If (InStr(desc, idstr) = 1) Then
        Dim i_end_emitter As Long
        Dim s_emitter As String
        Dim i_repeat_emitter As Long
        i_end_emitter = InStr(desc, ":")
        s_emitter = Mid$(desc, Len(idstr) + 1, i_end_emitter - Len(idstr) - 2)
        i_repeat_emitter = InStr(desc, " DE " & s_emitter)
        If i_repeat_emitter > 0 Then
            deleteDuplicateSepa = Left$(desc, i_repeat_emitter - 1)
        End If
    End If
End Function

Private Function strReplace(oldString, newString, targetString As String) As String
    strReplace = targetString
    i = InStr(targetString, oldString)
    If (i > 0) Then
        strReplace = Left$(targetString, i - 1) & newString & Right$(targetString, Len(targetString) - i - Len(oldString) + 1)
    End If
End Function

Private Function simplifyDescription(desc As String, subsTable As Variant) As String
    Dim s As String
    s = deleteDuplicateSepa(Trim$(desc))
    n = UBound(subsTable, 1)
    For i = 1 To n
        s = strReplace(subsTable(i, 1), subsTable(i, 2), s)
    Next i
    simplifyDescription = s
End Function

'------------------------------------------------------------------------------
' ING
'------------------------------------------------------------------------------

Public Sub ImportING(oTable As ListObject, fileToOpen As Variant)
    Dim iRow As Long, lastRow As Long
    Dim dateCol As Integer, amountCol As Integer, descCol As Integer
    dateCol = GetColumnNumberFromName(oTable, GetLabel(DATE_KEY))
    amountCol = GetColumnNumberFromName(oTable, GetLabel(AMOUNT_KEY))
    descCol = GetColumnNumberFromName(oTable, GetLabel(DESCRIPTION_KEY))
    
    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    iRow = 1
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True
    With oTable
        Do While LenB(Cells(iRow, 1).Value) > 0
            .ListRows.Add
            lastRow = .ListRows.Count
            .ListColumns(dateCol).DataBodyRange.Rows(lastRow).Value = Cells(iRow, 1).Value
            .ListColumns(amountCol).DataBodyRange.Rows(lastRow).Value = toAmount(Cells(iRow, 4).Value)
            .ListColumns(descCol).DataBodyRange.Rows(lastRow).Value = simplifyDescription(Cells(iRow, 2).Value, subsTable)
            iRow = iRow + 1
        Loop
    End With
    ActiveWorkbook.Close
    
    Call SortTable(oTable, GetLabel(DATE_KEY), xlAscending, GetLabel(AMOUNT_KEY), xlDescending)
    Range("A" & CStr(oTable.ListRows.Count)).Select
End Sub

'------------------------------------------------------------------------------
' LCL
'------------------------------------------------------------------------------

Public Sub ImportLCL(oTable As ListObject, fileToOpen As Variant)
    Dim iRow As Long, lastRow As Long
    Dim dateCol As Integer, amountCol As Integer, descCol As Integer
    dateCol = GetColumnNumberFromName(oTable, GetLabel(DATE_KEY))
    amountCol = GetColumnNumberFromName(oTable, GetLabel(AMOUNT_KEY))
    descCol = GetColumnNumberFromName(oTable, GetLabel(DESCRIPTION_KEY))

    Workbooks.Open filename:=fileToOpen, ReadOnly:=True
    Dim iRow As Long
    
    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    iRow = 1
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True
    With oTable
        Do While LenB(Cells(iRow + 1, 1).Value) > 0
            .ListRows.Add
            lastRow = .ListRows.Count
            .ListColumns(dateCol).DataBodyRange.Rows(lastRow).Value = DateValue(Cells(iRow, 1).Value)
            .ListColumns(amountCol).DataBodyRange.Rows(lastRow).Value = toAmount(Cells(iRow, 2).Value)
            If (Cells(iRow, 3).Value Like "Ch?que") Then
                .ListColumns(descCol).DataBodyRange.Rows(lastRow).Value = "Cheque " & simplifyDescription(CStr(Cells(iRow, 4).Value), subsTable)
            ElseIf (Cells(iRow, 3).Value = "Virement") Then
                .ListColumns(descCol).DataBodyRange.Rows(lastRow).Value = "Virement " & simplifyDescription(Cells(iRow, 5).Value, subsTable)
            Else
                .ListColumns(descCol).DataBodyRange.Rows(lastRow).Value = simplifyDescription(Cells(iRow, 3).Value & " " & Cells(iRow, 5).Value & " " & Cells(iRow, 6).Value, subsTable)
            End If
            iRow = iRow + 1
        Loop
    End With
    ActiveWorkbook.Close
    
    Call SortTable(oTable, GetLabel(DATE_KEY), xlAscending, GetLabel(AMOUNT_KEY), xlDescending)
    Range("A" & CStr(oTable.ListRows.Count)).Select
End Sub


Sub ImportRevolut(fileToOpen As Variant)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))

    Workbooks.Open filename:=fileToOpen, ReadOnly:=True
    Dim iRow As Long
    Dim tDates() As Variant
    Dim tDesc() As String
    Dim tAmounts() As Double
    
    
    iRow = 2
    Do While LenB(Cells(iRow, 1).Value) > 0 And iRow < MAX_IMPORT
        iRow = iRow + 1
    Loop
    nbRows = iRow - 2
    ReDim tDates(1 To nbRows)
    ReDim tDesc(1 To nbRows)
    ReDim tAmounts(1 To nbRows)
    iRow = 2
    Do While LenB(Cells(iRow, 1).Value) > 0
        tDates(iRow - 1) = toDate(Trim$(Cells(iRow, 1).Value))
        tDesc(iRow - 1) = ""
        If LenB(Trim$(Cells(iRow, 3).Value)) = 0 Then
            tAmounts(iRow - 1) = toAmount(Trim$(Cells(iRow, 4).Value))
            If LenB(Trim$(Cells(iRow, 6).Value)) > 0 Then
                tDesc(iRow - 1) = simplifyDescription(Trim$(Cells(iRow, 6).Value) & " : ", subsTable)
            End If
        Else
            tAmounts(iRow - 1) = -toAmount(Trim$(Cells(iRow, 3).Value))
            If LenB(Trim$(Cells(iRow, 5).Value)) > 0 Then
                tDesc(iRow - 1) = simplifyDescription(Trim$(Cells(iRow, 5).Value) & " : ", subsTable)
            End If
        End If
        tDesc(iRow - 1) = tDesc(iRow - 1) & Trim$(Cells(iRow, 2).Value)
        iRow = iRow + 1
    Loop
    ActiveWorkbook.Close
    
    Call addTransactionsSortAndSelect(ActiveSheet.ListObjects(1), tDates, tAmounts, tDesc)

End Sub


Sub ImportRevolutCSV(fileToOpen As Variant)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))

    Dim tbl As ListObject
    tbl = ActiveSheet.ListObjects(1)
    
    Dim dateCol As Long, amountCol As Long, descCol As Long
    dateCol = GetColumnNumberFromName(tbl, GetLabel(DATE_KEY))
    amountCol = GetColumnNumberFromName(tbl, GetLabel(AMOUNT_KEY))
    descCol = GetColumnNumberFromName(tbl, GetLabel(DESCRIPTION_KEY))
    
    Dim totalrows As Long
    totalrows = tbl.ListRows.Count
    
    Open fileToOpen For Input As #1
    Line Input #1, textline
    Do Until EOF(1)
        Line Input #1, textline
        A = Split(textline, ";", -1, vbTextCompare)
        tbl.ListRows.Add
        totalrows = totalrows + 1
        tbl.ListColumns(dateCol).DataBodyRange.Rows(totalrows).Value = toDate(Trim$(A(0)))
        If LenB(Trim$(A(2))) = 0 Then
            tbl.ListColumns(amountCol).DataBodyRange.Rows(totalrows).Value = CDbl(Trim$(A(3)))
            tbl.ListColumns(descCol).DataBodyRange.Rows(totalrows).Value = simplifyDescription(Trim$(A(1)) & " --> " & Trim$(A(5)), subsTable)
        Else
            tbl.ListColumns(amountCol).DataBodyRange.Rows(totalrows).Value = -CDbl(Trim$(A(2)))
            tbl.ListColumns(descCol).DataBodyRange.Rows(totalrows).Value = simplifyDescription(Trim$(A(1)) & " --> " & Trim$(A(4)), subsTable)
        End If
    Loop
    Close #1
    Call sortAccount(tbl)
    Range("A" & CStr(tbl.ListRows.Count)).Select

End Sub




Private Function CountUBSrows() As Long
    Dim i As Long
    i = 1
    Do While LenB(Cells(i, 1).Value) > 0
        i = i + 1
    Loop
    CountUBSrows = i - 1
End Function
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Private Sub readUBSdata(ByRef tDates As Variant, ByRef tDesc As Variant, ByRef tAmounts As Variant, nbRows As Long)
    Dim iRow As Long
    For iRow = 2 To nbRows
        If ws.Cells(iRow, 13) = "Solde prix prestations" Then
            tAmounts(iRow - 1) = 0
        ElseIf LenB(Cells(iRow, 18).Value) > 0 Then
            tAmounts(iRow - 1) = toAmount(Cells(iRow, 18).Value) ' Sous-montant column
        ElseIf LenB(Cells(iRow, 19).Value) > 0 Then
            tAmounts(iRow - 1) = -toAmount(Cells(iRow, 19).Value) ' Debit column
        ElseIf LenB(Cells(iRow, 20).Value) > 0 Then
            tAmounts(iRow - 1) = toAmount(Cells(iRow, 20).Value) ' Credit column
        Else
            tAmounts(iRow - 1) = 0
        End If
        tDates(iRow - 1) = CDate(DateValue(Replace(Cells(iRow, 12).Value, ".", "/")))
        tDesc(iRow - 1) = simplifyDescription(Cells(iRow, 13).Value & " " & Cells(iRow, 14).Value & " " & Cells(iRow, 15).Value, subsTable)
    Next iRow
End Sub
Sub ImportUBS(fileToOpen As Variant)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))

    Workbooks.Open filename:=fileToOpen, ReadOnly:=True

    Dim nbRows As Long
    nbRows = CountUBSrows()
    ReDim tDates(1 To nbRows - 1) As Variant
    ReDim tDesc(1 To nbRows - 1) As String
    ReDim tAmounts(1 To nbRows - 1) As Double
    Call readUBSdata(tDates, tDesc, tAmounts, nbRows)
    ActiveWorkbook.Close
    
    Call addTransactionsSortAndSelect(ActiveSheet.ListObjects(1), tDates, tAmounts, tDesc, "Montant CHF")

End Sub

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportUBScsv(fileToOpen As Variant)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))

    Workbooks.OpenText filename:="C:\Users\Olivier\Downloads\export.csv", Origin _
        :=65001, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
        Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
        , 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1)), _
        TrailingMinusNumbers:=True
        'ReadOnly:=True
    
    Dim nbRows As Long
    nbRows = CountUBSrows()
    ReDim tDates(1 To nbRows - 1) As Variant
    ReDim tDesc(1 To nbRows - 1) As String
    ReDim tAmounts(1 To nbRows - 1) As Double
    Call readUBSdata(tDates, tDesc, tAmounts, nbRows)
    ActiveWorkbook.Close
    
    Call addTransactionsSortAndSelect(ActiveSheet.ListObjects(1), tDates, tAmounts, tDesc, "Montant CHF")

End Sub

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportGeneric(fileToOpen As Variant)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))

    Workbooks.Open filename:=fileToOpen, ReadOnly:=True, local:=True
    'Workbooks.Open filename:="C:\Users\Olivier\Desktop\Test LCL.csv"
    Dim iRow As Long
    Dim tDates() As Variant
    Dim tDesc() As String
    Dim tSubCateg() As String
    Dim tBudgetSpread() As Variant
    Dim tAmounts() As Double
    
    iRow = 1
    
    ' Read Header part
    Do While LenB(Cells(iRow, 1).Value) > 0 And iRow < MAX_IMPORT
        iRow = iRow + 1
        If Cells(iRow, 1) = "Korach Exporter version" Then
            exporterVersion = Cells(iRow, 2).Value
        ElseIf Cells(iRow, 1) = "No Compte" Then
            accountNbr = Cells(iRow, 2).Value
        ElseIf Cells(iRow, 1) = "Nom Compte" Then
            accountName = Cells(iRow, 2).Value
        ElseIf Cells(iRow, 1) = "Banque" Then
            bank = Cells(iRow, 2).Value
        ElseIf Cells(iRow, 1) = "Status" Then
            accStatus = Cells(iRow, 2).Value
        ElseIf Cells(iRow, 1) Like "Disponibilit?" Then
            availability = Cells(iRow, 2).Value
        Else
            ' Do nothing
        End If
    Loop
    
    iRow = iRow + 1
    transactionStart = iRow
    ' Count nbr of transaction
    Do While LenB(Cells(iRow, 1).Value) > 0 And iRow < MAX_IMPORT
        iRow = iRow + 1
    Loop
    ' Read transaction part
    transactionStop = iRow - 1
    nbRows = transactionStop - transactionStart + 1
    ReDim tDates(1 To nbRows)
    ReDim tDesc(1 To nbRows)
    ReDim tSubCateg(1 To nbRows)
    ReDim tBudgetSpread(1 To nbRows)
    ReDim tAmounts(1 To nbRows)
    
    For iRow = transactionStart To transactionStop
        i = iRow - transactionStart + 1
        tDates(i) = Cells(iRow, 1).Value
        tDesc(i) = simplyDescription(Cells(iRow, 4).Value, subsTable)
        tAmounts(i) = toAmount(Cells(iRow, 3).Value)
        tSubCateg(i) = Cells(iRow, 5).Value
        tBudgetSpread(i) = Cells(iRow, 7).Value
    Next iRow
    ActiveWorkbook.Close
    
    ActiveSheet.Cells(1, 2).Value = accountName
    ActiveSheet.Cells(2, 2).Value = accountNbr
    ActiveSheet.Cells(3, 2).Value = bank
    ActiveSheet.Cells(4, 2).Value = accStatus
    ActiveSheet.Cells(5, 2).Value = availability
    
    Dim tbl As Variant
    Dim dateCol As Long
    Dim amountCol As Long
    Dim descCol As Long
    Dim subcatCol As Long
    Dim budgetCol As Long
    tbl = ActiveSheet.ListObjects(1)
    dateCol = GetColumnNumberFromName(tbl, GetLabel(DATE_KEY))
    amountCol = GetColumnNumberFromName(oTable, GetLabel(AMOUNT_KEY))
    descCol = GetColumnNumberFromName(oTable, GetLabel(DESCRIPTION_KEY))
    subcatCol = GetColumnNumberFromName(oTable, GetLabel(SUBCATEGORY_KEY))
    budgetCol = GetColumnNumberFromName(oTable, GetLabel(IN_BUDGET_KEY))
    With tbl
        totalrows = .ListRows.Count
        For iRow = 1 To nbRows
            .ListRows.Add
            totalrows = totalrows + 1
            .ListColumns(dateCol).DataBodyRange.Rows(totalrows).Value = tDates(iRow)
            .ListColumns(amountCol).DataBodyRange.Rows(totalrows).Value = tAmounts(iRow)
            .ListColumns(descCol).DataBodyRange.Rows(totalrows).Value = tDesc(iRow)
            .ListColumns(subcatCol).DataBodyRange.Rows(totalrows).Value = tSubCateg(iRow)
            .ListColumns(budgetCol).DataBodyRange.Rows(totalrows).Value = tBudgetSpread(iRow)
        Next iRow
    End With
    
    Call sortAccount(tbl)
    Range("A" & CStr(tbl.ListRows.Count)).Select

End Sub

Sub ExportGeneric(ws, Optional csvFile As String = "", Optional silent As Boolean = False)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))

    Dim sFolder As String
    exportFrom = ActiveWorkbook.name

    Sheets(ws).Select
    Range("A1:B8").Select
    Selection.Copy
    exportVersion = 1

    ' Create blank workbook and copy data on that workbook
    Workbooks.Add
    exportTo = ActiveWorkbook.name
    Range("A1").Select
    ActiveSheet.Paste
    Range("A9").Value = "Exporter version"
    Range("B9").Value = 1.2

    Workbooks(exportFrom).Activate
    Sheets(ws).ListObjects(1).DataBodyRange.Select
    Selection.Copy
    Workbooks(exportTo).Activate
    Range("A10").Select
    ActiveSheet.Paste
    Range("B:C").NumberFormat = "General"
    'Range("A:A").NumberFormat = Workbooks(exportFrom).Names("date_format").RefersToRange.Value
    Range("A:A").NumberFormat = "YYYY-mm-dd"

    ' Silently delete sheets in excess
    Call DeleteAllButSheetOne

    ' Get filename to save
    If LenB(csvFile) = 0 Then
        file = Application.GetSaveAsFilename
        If file Then
           csvFile = file & "csv"
        End If
    End If

    ' Save CSV file
    If LenB(csvFile) > 0 Then
        ActiveWorkbook.SaveAs filename:=csvFile, fileformat:=xlCSV, CreateBackup:=False, local:=True
        If (Not silent) Then
            MsgBox "File " & csvFile & " saved"
        End If
    Else
        If Not silent Then
            Call ErrorMessage("k.warningExportCancelled")
        End If
    End If
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True

End Sub
Sub ExportAll()
    Dim sFolder As String
    Dim filename As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With

    If LenB(sFolder) > 0 Then ' if a file was chosen
        Call FreezeDisplay
        For Each ws In Worksheets
            If ws.Cells(1, 1).Value = "Nom Compte" Then
                filename = sFolder & "\" & ws.name & ".csv"
                Call ExportGeneric(ws.name, filename, True)
            End If
        Next ws
        Call UnfreezeDisplay
    Else
        Call ErrorMessage("k.warningExportCancelled")
    End If
End Sub
Sub ExportLCL()
    Call ExportGeneric("LCL CC")
End Sub
Sub ExportING()
    Call ExportGeneric("ING CC")
End Sub


Private Sub addTransactions(oTable As ListObject, transDates As Variant, transAmounts As Variant, transDesc As Variant, _
                            Optional amountColName As String = "")
    Dim dateCol As Long
    Dim amountCol As Long
    Dim descCol As Long
    dateCol = GetColumnNumberFromName(oTable, GetLabel(DATE_KEY))
    If amountColName = "" Then
        amountColName = GetLabel(AMOUNT_KEY)
    End If
    amountCol = GetColumnNumberFromName(oTable, amountColName)
    descCol = GetColumnNumberFromName(oTable, GetLabel(DESCRIPTION_KEY))
    
    With oTable
        totalrows = .ListRows.Count
        For iRow = 1 To UBound(transDates)
            .ListRows.Add
            totalrows = totalrows + 1
            .ListColumns(dateCol).DataBodyRange.Rows(totalrows).Value = transDates(iRow)
            .ListColumns(amountCol).DataBodyRange.Rows(totalrows).Value = transAmounts(iRow)
            .ListColumns(descCol).DataBodyRange.Rows(totalrows).Value = transDesc(iRow)
        Next iRow
    End With
End Sub

Private Sub addTransactionsSortAndSelect(oTable As ListObject, transDates As Variant, transAmounts As Variant, transDesc As Variant, _
                            Optional amountColName As String = "")
    Call addTransactions(oTable, transDates, transAmounts, transDesc, amountColName)
    Call SortTable(oTable, GetLabel(DATE_KEY), xlAscending, GetLabel(AMOUNT_KEY), xlDescending)
    Range("A" & CStr(oTable.ListRows.Count)).Select

End Sub



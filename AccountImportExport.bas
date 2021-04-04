Attribute VB_Name = "AccountImportExport"
Const MAX_IMPORT = 30000

Private Const SUBSTITUTIONS_TABLE = "TblSubstitutions"

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
    ElseIf s Like "a[vp]r*" Then
        toMonth = 4
    ElseIf s Like "ma[iy]*" Then
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
        Dim oTable As ListObject
        Set oTable = ActiveSheet.ListObjects(1)
        Dim dateCol As Integer, amountCol As Integer, descCol As Integer
        dateCol = TableColNbrFromName(oTable, GetLabel(DATE_KEY))
        
        Dim defaultCurrency As String, accCurrency As String
        defaultCurrency = GetGlobalParam("DefaultCurrency")
        accCurrency = AccountCurrency(ActiveSheet.name)
        If accCurrency = defaultCurrency Then
            amountCol = TableColNbrFromName(oTable, GetLabel(AMOUNT_KEY))
        Else
            amountCol = TableColNbrFromName(oTable, GetLabel(AMOUNT_KEY) & " " & accCurrency)
        End If
        descCol = TableColNbrFromName(oTable, GetLabel(DESCRIPTION_KEY))
        Dim bank As String
        bank = Cells(3, 2).value
        If (bank = "ING Direct") Then
            Call ImportING(oTable, fileToOpen, dateCol, amountCol, descCol)
        ElseIf (bank = "LCL") Then
            Call ImportLCL(oTable, fileToOpen, dateCol, amountCol, descCol)
        ElseIf (bank = "UBS") Then
            Call ImportUBS(oTable, fileToOpen, dateCol, amountCol, descCol)
        ElseIf (bank = "Revolut") Then
            Call ImportRevolut(oTable, fileToOpen, dateCol, amountCol, descCol)
        Else
            Call UnfreezeDisplay
            Call ErrorMessage("k.errorImportNotRecognized", "k.warningImportCancelled")
        End If
        Call SortTable(oTable, GetLabel(DATE_KEY), xlAscending, GetLabel(AMOUNT_KEY), xlDescending)
        Range("A" & CStr(oTable.ListRows.Count)).Select
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

Public Sub ImportING(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True

    Dim iRow As Long
    iRow = 1
    Do While LenB(Cells(iRow, 1).value) > 0
        oTable.ListRows.Add
        With oTable.ListRows(oTable.ListRows.Count)
            .Range(1, dateCol).value = Cells(iRow, 1).value
            .Range(1, amountCol).value = toAmount(Cells(iRow, 4).value)
            .Range(1, descCol).value = simplifyDescription(Cells(iRow, 2).value, subsTable)
        End With
        iRow = iRow + 1
    Loop
    ActiveWorkbook.Close
End Sub

'------------------------------------------------------------------------------
' LCL
'------------------------------------------------------------------------------

Public Sub ImportLCL(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True

    Dim iRow As Long
    iRow = 1
    Do While LenB(Cells(iRow + 1, 1).value) > 0
        oTable.ListRows.Add
        With oTable.ListRows(oTable.ListRows.Count)
            .Range(1, dateCol).value = DateValue(Cells(iRow, 1).value)
            .Range(1, amountCol).value = toAmount(Cells(iRow, 2).value)
            If (Cells(iRow, 3).value Like "Ch?que") Then
                .Range(1, descCol).value = "Cheque " & simplifyDescription(CStr(Cells(iRow, 4).value), subsTable)
            ElseIf (Cells(iRow, 3).value = "Virement") Then
                .Range(1, descCol).value = "Virement " & simplifyDescription(Cells(iRow, 5).value, subsTable)
            Else
                .Range(1, descCol).value = simplifyDescription(Cells(iRow, 3).value & " " & Cells(iRow, 5).value & " " & Cells(iRow, 6).value, subsTable)
            End If
        End With
        iRow = iRow + 1
    Loop
    ActiveWorkbook.Close
End Sub

'------------------------------------------------------------------------------
' Revolut
'------------------------------------------------------------------------------

Sub ImportRevolut(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)
    If LCase$(Right(fileToOpen, 4)) = ".csv" Then
        Call importRevolutCsv(oTable, fileToOpen, dateCol, amountCol, descCol)
    Else
        Call importRevolutXls(oTable, fileToOpen, dateCol, amountCol, descCol)
    End If
End Sub

Private Sub importRevolutXls(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True
    
    Dim iRow As Long
    iRow = 2
    Do While LenB(Cells(iRow, 1).value) > 0
        oTable.ListRows.Add
        With oTable.ListRows(oTable.ListRows.Count)
            On Error GoTo ErrDate
            .Range(1, dateCol).value = DateValue(Trim$(Cells(iRow, 1).value))
            GoTo CheckAmount
ErrDate:
            .Range(1, dateCol).value = toDate(Trim$(Cells(iRow, 1).value))
CheckAmount:
            Dim desc As String
            desc = ""
            If LenB(Trim$(Cells(iRow, 3).value)) = 0 Then
                .Range(1, amountCol).value = toAmount(Trim$(Cells(iRow, 4).value))
                If LenB(Trim$(Cells(iRow, 6).value)) > 0 Then
                    desc = simplifyDescription(Trim$(Cells(iRow, 6).value) & " : ", subsTable)
                End If
            Else
                .Range(1, amountCol).value = -toAmount(Trim$(Cells(iRow, 3).value))
                If LenB(Trim$(Cells(iRow, 5).value)) > 0 Then
                    desc = simplifyDescription(Trim$(Cells(iRow, 5).value) & " : ", subsTable)
                End If
            End If
            .Range(1, descCol).value = desc & Trim$(Cells(iRow, 2).value)
        End With
        iRow = iRow + 1
    Loop
    ActiveWorkbook.Close
End Sub


Private Sub importRevolutCsv(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)
    ' Open file a first time to replace " , " and " ," by ";"

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))

    Workbooks.Add
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & fileToOpen, Destination:=Range("$A$1"))
        .name = "import"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 65001
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Cells.Replace What:=" , ", Replacement:=";", LookAt:=xlPart
    Cells.Replace What:=", ", Replacement:=";", LookAt:=xlPart

    Dim iRow As Long
    iRow = 2
    Do While LenB(Cells(iRow, 1).value) > 0
        A = Split(Cells(iRow, 1).value, ";", -1, vbTextCompare)
        oTable.ListRows.Add
        With oTable.ListRows(oTable.ListRows.Count)
            Dim desc As String, comment As String
            Dim amount As Double
            .Range(1, dateCol).value = toDate(Trim$(A(0)))
            desc = Trim$(A(1))
            If LenB(Trim$(A(2))) = 0 Then
                amount = CDbl(Trim$(A(3)))
                comment = Trim$(A(5))
            Else
                amount = -CDbl(Trim$(A(2)))
                comment = Trim$(A(4))
            End If
            .Range(1, amountCol).value = amount
            If comment <> "" Then
                .Range(1, descCol).value = simplifyDescription(desc & " --> " & comment, subsTable)
            Else
                .Range(1, descCol).value = simplifyDescription(desc, subsTable)
            End If
        End With
        iRow = iRow + 1
    Loop
    ActiveWorkbook.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' UBS
'------------------------------------------------------------------------------


Sub ImportUBS(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    If LCase$(Right(fileToOpen, 4)) = ".csv" Then
        xlsFile = convertCsvToXls(fileToOpen)
        Workbooks.Open filename:=xlsFile, ReadOnly:=True
    Else
        Workbooks.Open filename:=fileToOpen, ReadOnly:=True
    End If
    Dim iRow As Long
    iRow = 2
    Do While LenB(Cells(iRow, 1).value) > 0
        oTable.ListRows.Add
        With oTable.ListRows(oTable.ListRows.Count)
            If Cells(iRow, 13) = "Solde prix prestations" Then
                .Range(1, amountCol).value = 0
            ElseIf LenB(Cells(iRow, 18).value) > 0 Then
                .Range(1, amountCol).value = toAmount(Cells(iRow, 18).value) ' Sous-montant column
            ElseIf LenB(Cells(iRow, 19).value) > 0 Then
                .Range(1, amountCol).value = -toAmount(Cells(iRow, 19).value) ' Debit column
            ElseIf LenB(Cells(iRow, 20).value) > 0 Then
                .Range(1, amountCol).value = toAmount(Cells(iRow, 20).value) ' Credit column
            Else
                .Range(1, amountCol).value = 0
            End If
            .Range(1, dateCol).value = CDate(DateValue(Replace(Cells(iRow, 12).value, ".", "/")))
            .Range(1, descCol).value = simplifyDescription(Cells(iRow, 13).value & " " & Cells(iRow, 14).value & " " & Cells(iRow, 15).value, subsTable)
        End With
        iRow = iRow + 1
    Loop
    ActiveWorkbook.Close
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
    xlsFile = Left(fileToOpen, Len(fileToOpen) - 4) & format(Now(), "yyyy-MM-dd hh-mm-ss") & ".xls"
    ActiveWorkbook.SaveAs filename:=xlsFile, fileformat:=xlExcel8, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
    ActiveWorkbook.Close
    convertCsvToXls = xlsFile
End Function

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportGeneric(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True, local:=True

    Dim tDates() As Variant
    Dim tDesc() As String
    Dim tSubCateg() As String
    Dim tBudgetSpread() As Variant
    Dim tAmounts() As Double
    
    Dim iRow As Long
    iRow = 1
    ' Read Header part
    Do While LenB(Cells(iRow, 1).value) > 0 And iRow < MAX_IMPORT
        iRow = iRow + 1
        If Cells(iRow, 1) = "Korach Exporter version" Then
            exporterVersion = Cells(iRow, 2).value
        ElseIf Cells(iRow, 1) = "No Compte" Then
            accNumber = Cells(iRow, 2).value
        ElseIf Cells(iRow, 1) = "Nom Compte" Then
            accName = Cells(iRow, 2).value
        ElseIf Cells(iRow, 1) = "Banque" Then
            bank = Cells(iRow, 2).value
        ElseIf Cells(iRow, 1) = "Status" Then
            accStatus = Cells(iRow, 2).value
        ElseIf Cells(iRow, 1) Like "Disponibilit?" Then
            availability = Cells(iRow, 2).value
        Else
            ' Do nothing
        End If
    Loop
    
    iRow = iRow + 1
    transactionStart = iRow
    ' Count nbr of transaction
    Do While LenB(Cells(iRow, 1).value) > 0 And iRow < MAX_IMPORT
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
        tDates(i) = Cells(iRow, 1).value
        tDesc(i) = simplyDescription(Cells(iRow, 4).value, subsTable)
        tAmounts(i) = toAmount(Cells(iRow, 3).value)
        tSubCateg(i) = Cells(iRow, 5).value
        tBudgetSpread(i) = Cells(iRow, 7).value
    Next iRow
    ActiveWorkbook.Close
    
    ActiveSheet.Cells(1, 2).value = accName
    ActiveSheet.Cells(2, 2).value = accNumber
    ActiveSheet.Cells(3, 2).value = bank
    ActiveSheet.Cells(4, 2).value = accStatus
    ActiveSheet.Cells(5, 2).value = availability

    Dim subcatCol As Long, budgetCol As Long
    subcatCol = TableColNbrFromName(oTable, GetLabel(SUBCATEGORY_KEY))
    budgetCol = TableColNbrFromName(oTable, GetLabel(IN_BUDGET_KEY))
    For iRow = 1 To nbRows
        oTable.ListRows.Add
        With oTable.ListRows(oTable.ListRows.Count)
            .Range(1, dateCol).value = tDates(iRow)
            .Range(1, amountCol).value = tAmounts(iRow)
            .Range(1, descCol).value = tDesc(iRow)
            .Range(1, subcatCol).value = tSubCateg(iRow)
            .Range(1, budgetCol).value = tBudgetSpread(iRow)
        End With
    Next iRow
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
    Range("A9").value = "Exporter version"
    Range("B9").value = 1.2

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
            If ws.Cells(1, 1).value = "Nom Compte" Then
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



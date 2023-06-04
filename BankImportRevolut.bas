Attribute VB_Name = "BankImportRevolut"

'------------------------------------------------------------------------------
' Import Revolut
'------------------------------------------------------------------------------

Private Const REVOLUT_CSV_DATE_FIELD = 2
Private Const REVOLUT_CSV_DESC_FIELD1 = 0
Private Const REVOLUT_CSV_DESC_FIELD2 = 4
Private Const REVOLUT_CSV_AMOUNT_FIELD = 5
Private Const REVOLUT_CSV_FEE_FIELD = 6

Sub ImportRevolut(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)
    If LCase$(Right$(fileToOpen, 4)) = ".csv" Then
        Call importRevolutCsv(oTable, fileToOpen, dateCol, amountCol, descCol)
    Else
        Call importRevolutXls(oTable, fileToOpen, dateCol, amountCol, descCol)
    End If
End Sub

Private Sub importRevolutXls(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)
    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True
    Dim modal As ProgressBar
    Set modal = NewProgressBar("Import Revolut XLS in progress", GetLastNonEmptyRow())
    modal.Update

    Dim i As Long
    i = 2
    Do While LenB(Cells(i, 1).value) > 0
        oTable.ListRows.Add
        With oTable.ListRows(oTable.ListRows.Count)
            On Error GoTo ErrDate
            .Range(1, dateCol).value = DateValue(Trim$(Cells(i, 4).value))
            GoTo CheckAmount
ErrDate:
            d = Split(Trim$(Cells(i, 4).value), " ", -1, vbTextCompare)
            .Range(1, dateCol).value = ToDate(d(1))
CheckAmount:
            Dim desc As String
            desc = simplifyDescription(Trim$(Cells(i, 5).value), subsTable)
            .Range(1, amountCol).value = toAmount(Trim$(Cells(i, 6).value))
            If LenB(Trim$(Cells(i, 7).value)) > 0 Then
                .Range(1, amountCol).value = toAmount(Trim$(Cells(i, 6).value)) + toAmount(Trim$(Cells(i, 6).value))
                desc = desc & " - incl. fee " & toAmount(Trim$(Cells(i, 7).value))
            End If
            .Range(1, descCol).value = desc
        End With
        i = i + 1
        modal.Update
    Loop
    ActiveWorkbook.Close
    Set modal = Nothing
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
        .saveData = True
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
    Cells.Replace What:=",", Replacement:=";", LookAt:=xlPart
    Cells.Replace What:=";""", Replacement:=";", LookAt:=xlPart
    Cells.Replace What:="""", Replacement:="", LookAt:=xlPart

    Dim modal As ProgressBar
    Set modal = NewProgressBar("Import Revolut CSV in progress", GetLastNonEmptyRow())
    modal.Update

    Dim i As Long
    i = 2
    Do While LenB(Cells(i, 1).value) > 0
        oTable.ListRows.Add
        With oTable.ListRows(oTable.ListRows.Count)
            Dim desc As String, comment As String
            Dim amount As Double
            Dim a As Variant
            If Cells(i, 2).value = "" Then
                ' semicolon CSV cell separator did not work
                a = Split(Cells(i, 1).value, ";", -1, vbTextCompare)
            Else
                ' semicolon CSV cell separator did work
                ' 10 cells is arbitrary but OK for now (rows are 6 cells as per exports in Oct 2022)
                ReDim a(0 To 9) As Variant
                For j = 1 To 10
                    a(j - 1) = Cells(i, j).value
                Next j
            End If
            dateAndTime = Split(a(REVOLUT_CSV_DATE_FIELD), " ", -1, vbTextCompare)
            .Range(1, dateCol).value = ToDate(dateAndTime(0))
            desc = Trim$(a(REVOLUT_CSV_DESC_FIELD1))
            comment = Trim$(a(REVOLUT_CSV_DESC_FIELD2))
            amount = CDbl(Trim$(a(REVOLUT_CSV_AMOUNT_FIELD)))
            fee = CDbl(Trim$(a(REVOLUT_CSV_FEE_FIELD)))
            If fee <> 0 Then
                amount = amount + fee
                comment = comment & " (including fee of " & str(fee) & " ¤)"
            End If
            .Range(1, amountCol).value = amount
            If Len(comment) > 0 Then
                .Range(1, descCol).value = simplifyDescription(desc & " " & comment, subsTable)
            Else
                .Range(1, descCol).value = simplifyDescription(desc, subsTable)
            End If
        End With
        i = i + 1
        modal.Update
    Loop
    ActiveWorkbook.Close SaveChanges:=False
    Set modal = Nothing
End Sub



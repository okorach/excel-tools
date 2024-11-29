Attribute VB_Name = "BankImportYuh"

'------------------------------------------------------------------------------
' Import Yuh
'------------------------------------------------------------------------------

Private Const YUH_CSV_DATE_FIELD = 0
Private Const YUH_CSV_DESC_FIELD1 = 2
Private Const YUH_CSV_DESC_FIELD2 = 1
Private Const YUH_CSV_AMOUNT_DEBIT_FIELD = 3
Private Const YUH_CSV_CURRENCY_DEBIT_FIELD = 4
Private Const YUH_CSV_AMOUNT_CREDIT_FIELD = 5
Private Const YUH_CSV_CURRENCY_CREDIT_FIELD = 6
Private Const YUH_CSV_FEE_FIELD = 11

Sub ImportYuh(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer, accountCurrency As String)
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
    Cells.Replace What:="""", Replacement:="", LookAt:=xlPart

    Dim modal As ProgressBar
    Set modal = NewProgressBar("Import Yuh CSV in progress", GetLastNonEmptyRow())
    modal.Update

    Dim i As Long
    i = 2
    
    Do While LenB(Cells(i, 1).value) > 0
        Dim a As Variant
        Dim amount As Double
        If Cells(i, 2).value = "" Then
            ' semicolon CSV cell separator did not work
            a = Split(Cells(i, 1).value, ";", -1, vbTextCompare)
        Else
            ' semicolon CSV cell separator did work
            ' 12 cells is arbitrary but OK for now (rows are 11 cells as per exports in Oct 2022)
            ReDim a(0 To 16) As Variant
            For j = 1 To 16
                a(j - 1) = Cells(i, j).value
            Next j
        End If
        If a(YUH_CSV_AMOUNT_DEBIT_FIELD) <> "" Then
            transactionCurrency = Trim$(a(YUH_CSV_CURRENCY_DEBIT_FIELD))
        Else
            transactionCurrency = Trim$(a(YUH_CSV_CURRENCY_CREDIT_FIELD))
        End If
        If a(YUH_CSV_DESC_FIELD2) <> "REWARD_RECEIVED" And transactionCurrency = accountCurrency Then
            If a(YUH_CSV_AMOUNT_DEBIT_FIELD) <> "" Then
                amount = CDbl(Trim$(a(YUH_CSV_AMOUNT_DEBIT_FIELD)))
            Else
                amount = CDbl(Trim$(a(YUH_CSV_AMOUNT_CREDIT_FIELD)))
            End If
            oTable.ListRows.Add
            With oTable.ListRows(oTable.ListRows.Count)
                Dim desc As String, comment As String
                .Range(1, dateCol).value = ToDate(a(YUH_CSV_DATE_FIELD))
                desc = Trim$(a(YUH_CSV_DESC_FIELD1))
                ' comment = Trim$(a(YUH_CSV_DESC_FIELD2))
                .Range(1, amountCol).value = amount
                If a(YUH_CSV_FEE_FIELD) <> "" Then
                    fee = Abs(CDbl(Trim$(a(YUH_CSV_FEE_FIELD))))
                    If fee <> 0 Then
                        desc = desc & " (including fee of " & str(fee) & " " & accountCurrency & ")"
                    End If
                End If
                 .Range(1, descCol).value = simplifyDescription(desc, subsTable)
            End With
        End If
        i = i + 1
        modal.Update
    Loop
    ActiveWorkbook.Close SaveChanges:=False
    Set modal = Nothing
End Sub



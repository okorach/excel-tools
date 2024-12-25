Attribute VB_Name = "BankImportBoursoBank"
'------------------------------------------------------------------------------
' Import Boursorama
'------------------------------------------------------------------------------

Private Const BOURSORAMA_CSV_DATE_FIELD = 2
Private Const BOURSORAMA_CSV_AMOUNT_FIELD = 7
Private Const BOURSORAMA_CSV_DESC_FIELD = 3
Private Const BOURSORAMA_CSV_ACCOUNT_FIELD = 9

Sub ImportBoursorama(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer, accNbr As String)
    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    Workbooks.Add
    Dim Account As String
    
    Account = accNbr
    On Error Resume Next
        Account = CStr(CLng(accNbr))
    
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

    Dim modal As ProgressBar
    Set modal = NewProgressBar("Import Boursorama CSV in progress", GetLastNonEmptyRow())
    modal.Update

    Dim i As Long
    i = 2
    Dim desc As String
    Do While LenB(Cells(i, 1).value) > 0
        a = Split(Cells(i, 1).value, ";", -1, vbTextCompare)
        If (Cells(i, BOURSORAMA_CSV_ACCOUNT_FIELD) = Account) Then
            oTable.ListRows.Add

            With oTable.ListRows(oTable.ListRows.Count)
                .Range(1, dateCol).value = Cells(i, BOURSORAMA_CSV_DATE_FIELD)
                .Range(1, amountCol).value = toAmount(Cells(i, BOURSORAMA_CSV_AMOUNT_FIELD))
                desc = Trim$(Cells(i, BOURSORAMA_CSV_DESC_FIELD))
                .Range(1, descCol).value = simplifyDescription(desc, subsTable)
            End With
        ElseIf (Cells(i, BOURSORAMA_CSV_ACCOUNT_FIELD) = "Relevé différé Carte " + Account) Then
            ' Handle credit card statement
            oTable.ListRows.Add
            With oTable.ListRows(oTable.ListRows.Count)
                .Range(1, dateCol).value = Cells(i, BOURSORAMA_CSV_DATE_FIELD)
                .Range(1, amountCol).value = -toAmount(Cells(i, BOURSORAMA_CSV_AMOUNT_FIELD))
                desc = Trim$(Cells(i, BOURSORAMA_CSV_DESC_FIELD))
                .Range(1, descCol).value = simplifyDescription(desc, subsTable)
            End With
        End If
        i = i + 1
        modal.Update
    Loop
    ActiveWorkbook.Close SaveChanges:=False
    Set modal = Nothing
End Sub



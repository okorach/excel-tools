Attribute VB_Name = "AccountImportExport"
Private Const SUBSTITUTIONS_TABLE = "TblSubstitutions"

Private Const REVOLUT_CSV_DATE_FIELD = 2
Private Const REVOLUT_CSV_DESC_FIELD1 = 0
Private Const REVOLUT_CSV_DESC_FIELD2 = 4
Private Const REVOLUT_CSV_AMOUNT_FIELD = 5
Private Const REVOLUT_CSV_FEE_FIELD = 6

Private Const BOURSORAMA_CSV_DATE_FIELD = 2
Private Const BOURSORAMA_CSV_AMOUNT_FIELD = 6
Private Const BOURSORAMA_CSV_DESC_FIELD = 3

Private Const LCL_CSV_DATE_FIELD = 1
Private Const LCL_CSV_AMOUNT_FIELD = 2
Private Const LCL_CSV_TYPE_FIELD = 3
Private Const LCL_CSV_DESC1_FIELD = 4
Private Const LCL_CSV_DESC2_FIELD = 5
Private Const LCL_CSV_DESC3_FIELD = 6

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

Private Function ToDate(str) As Date
    If InStr(str, " ") Then
        a = Split(str, " ", -1, vbTextCompare)
        ToDate = DateSerial(CInt(a(2)), toMonth(a(1)), CInt(a(0)))
    ElseIf InStr(str, "/") Then
        a = Split(str, "/", -1, vbTextCompare)
        ' Assume DD/MM/YYYY
        ToDate = DateSerial(CInt(a(2)), CInt(a(1)), CInt(a(0)))
    ElseIf InStr(str, "-") Then
        ToDate = isoToDate(str)
    Else
        ToDate = DateSerial(0, 0, 0)
    End If
End Function

Private Function isoToDate(str) As Date
    a = Split(str, "-", -1, vbTextCompare)
    isoToDate = DateSerial(CInt(a(0)), CInt(a(1)), CInt(a(2)))
End Function




'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportAny()

    Dim fileToOpen As Variant
    fileToOpen = Application.GetOpenFilename()
    If fileToOpen <> False Then
        Call FreezeDisplay
        Dim oAccount As Account
        Set oAccount = LoadAccount(getAccountId(ActiveSheet))
        
        Dim defaultCurrency As String
        defaultCurrency = GetGlobalParam("DefaultCurrency")
        Dim oTable As ListObject
        Set oTable = oAccount.BalanceTable
        
        Dim dateCol As Integer, amountCol As Integer, descCol As Integer
        dateCol = TableColNbrFromName(oTable, GetLabel(DATE_KEY))
        
        If oAccount.MyCurrency = defaultCurrency Then
            amountCol = TableColNbrFromName(oTable, GetLabel(AMOUNT_KEY))
        Else
            amountCol = TableColNbrFromName(oTable, GetLabel(AMOUNT_KEY) & " " & oAccount.MyCurrency)
        End If
        descCol = TableColNbrFromName(oTable, GetLabel(DESCRIPTION_KEY))
        If (oAccount.Bank = "ING") Then
            Call ImportING(oTable, fileToOpen, dateCol, amountCol, descCol)
        ElseIf (oAccount.Bank = "LCL") Then
            Call ImportLCL(oTable, fileToOpen, dateCol, amountCol, descCol)
        ElseIf (oAccount.Bank = "UBS") Then
            Call ImportUBS(oTable, fileToOpen, dateCol, amountCol, descCol)
        ElseIf (oAccount.Bank = "Revolut") Then
            Call ImportRevolut(oTable, fileToOpen, dateCol, amountCol, descCol)
        ElseIf (oAccount.Bank = "Boursorama") Then
            Call ImportBoursorama(oTable, fileToOpen, dateCol, amountCol, descCol)
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



Public Sub AccountImport()
    Dim fileToOpen As Variant
    fileToOpen = Application.GetOpenFilename()
    If fileToOpen <> False Then
        Call ImportGeneric(ActiveSheet.name, fileToOpen)
    End If
End Sub

' PRLV SEPA CE URSSAF RHONE ALPES-CNTFS : FR55ZZZ143065 000828DC120181231145950A000136092 DE
' CE URSSAF RHONE ALPES-CNTFS : 000828DC120181231145950A000136092 FR55ZZZ143065
Private Function deleteDuplicateSepa(desc As String) As String
    Dim idstr As String
    idstr = "PRLV SEPA "
    deleteDuplicateSepa = desc
    If (InStr(desc, idstr) = 1) Then
        Dim i_end_emitter As Long
        Dim s_emitter As String
        Dim i_repeat_emitter As Long
        i_end_emitter = InStr(desc, ":")
        If i_end_emitter > 0 Then
            s_emitter = Mid$(desc, Len(idstr) + 1, i_end_emitter - Len(idstr) - 2)
            i_repeat_emitter = InStr(desc, " DE " & s_emitter)
            If i_repeat_emitter > 0 Then
                deleteDuplicateSepa = left$(desc, i_repeat_emitter - 1)
            End If
        End If
    End If
End Function

Private Function strReplace(oldString, newString, targetString As String) As String
    strReplace = targetString
    i = InStr(targetString, oldString)
    If (i > 0) Then
        strReplace = left$(targetString, i - 1) & newString & Right$(targetString, Len(targetString) - i - Len(oldString) + 1)
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

'------------------------------------------------------------------------------
' LCL
'------------------------------------------------------------------------------



Public Sub ImportLCL(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True
    Dim modal As ProgressBar
    Set modal = NewProgressBar("Import LCL in progress", GetLastNonEmptyRow() + 1)
    modal.Update
    
    Dim csvIsSplit As Boolean
    csvIsSplit = Not (Cells(1, 2).value = "")
    Dim r As Long
    r = 1
    Do While LenB(Cells(r + 1, 1).value) > 0
        oTable.ListRows.Add
        With oTable.ListRows(oTable.ListRows.Count)
            If csvIsSplit Then
                ' semicolon CSV cell separator did work
                rawDate = DateValue(Cells(r, LCL_CSV_DATE_FIELD).value)
                rawAmount = Cells(r, LCL_CSV_AMOUNT_FIELD).value
                If (Cells(r, LCL_CSV_TYPE_FIELD).value Like "Ch?que") Then
                    rawDesc = "Cheque " & simplifyDescription(CStr(Cells(r, LCL_CSV_DESC1_FIELD).value), subsTable)
                ElseIf (Cells(r, LCL_CSV_TYPE_FIELD).value = "Virement") Then
                    .Range(1, descCol).value = "Virement " & simplifyDescription(Cells(r, LCL_CSV_DESC2_FIELD).value, subsTable)
                Else
                    .Range(1, descCol).value = simplifyDescription(Cells(r, LCL_CSV_TYPE_FIELD).value & " " & Cells(r, LCL_CSV_DESC2_FIELD).value & " " & Cells(r, LCL_CSV_DESC3_FIELD).value, subsTable)
                End If
            Else
                ' semicolon CSV cell separator did not work
                Dim a() As String
                a = Split(Cells(r, 1).value, ";", -1, vbTextCompare)
                rawDate = ToDate(Trim$(a(LCL_CSV_DATE_FIELD - 1)))
                rawAmount = toAmount(Trim$(a(LCL_CSV_AMOUNT_FIELD - 1)))
                Dim des As String
                If (a(LCL_CSV_TYPE_FIELD - 1) Like "Ch?que") Then
                    rawDesc = "Cheque " & simplifyDescription(a(LCL_CSV_DESC1_FIELD - 1), subsTable)
                ElseIf (a(LCL_CSV_TYPE_FIELD - 1) = "Virement") Then
                    rawDesc = "Virement " & simplifyDescription(a(LCL_CSV_DESC2_FIELD - 1), subsTable)
                Else
                    rawDesc = simplifyDescription(a(LCL_CSV_TYPE_FIELD - 1) & " " & a(LCL_CSV_DESC2_FIELD - 1) & " " & a(LCL_CSV_DESC3_FIELD - 1), subsTable)
                End If

            End If
            .Range(1, dateCol).value = rawDate
            .Range(1, amountCol).value = toAmount(rawAmount)
            .Range(1, descCol).value = rawDesc
        End With
        r = r + 1
        modal.Update
    Loop
    ActiveWorkbook.Close
    Set modal = Nothing
End Sub

Sub ImportRevolut(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)
    If LCase$(Right$(fileToOpen, 4)) = ".csv" Then
        Call importRevolutCsv(oTable, fileToOpen, dateCol, amountCol, descCol)
    Else
        Call importRevolutXls(oTable, fileToOpen, dateCol, amountCol, descCol)
    End If
End Sub

'------------------------------------------------------------------------------
' Boursorama
'------------------------------------------------------------------------------

Sub ImportBoursorama(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)
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

    Dim modal As ProgressBar
    Set modal = NewProgressBar("Import Boursorama CSV in progress", GetLastNonEmptyRow())
    modal.Update

    Dim i As Long
    i = 2
    Do While LenB(Cells(i, 1).value) > 0
        a = Split(Cells(i, 1).value, ";", -1, vbTextCompare)
        oTable.ListRows.Add
        With oTable.ListRows(oTable.ListRows.Count)
            Dim desc As String, comment As String
            Dim amount As Double
            .Range(1, dateCol).value = Cells(i, BOURSORAMA_CSV_DATE_FIELD)
            .Range(1, amountCol).value = toAmount(Cells(i, BOURSORAMA_CSV_AMOUNT_FIELD))
            desc = Trim$(Cells(i, BOURSORAMA_CSV_DESC_FIELD))
            .Range(1, descCol).value = simplifyDescription(desc, subsTable)
        End With
        i = i + 1
        modal.Update
    Loop
    ActiveWorkbook.Close SaveChanges:=False
    Set modal = Nothing
End Sub

'------------------------------------------------------------------------------
' Revolut
'------------------------------------------------------------------------------

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

'------------------------------------------------------------------------------
' UBS
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

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Public Sub ImportGeneric(accountId As String, fileToOpen As Variant)

   ' Dim xlsFile As String
    Dim importFrom As String, importTo As String
    importTo = ActiveWorkbook.name

    subsTable = GetTableAsArray(Workbooks(importTo).Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    
    Dim accId As String, cur As String, typ As String, avail As Integer, exportDate As String
    Dim Bank As String, TaxRate As Double, accNbr As String
    Dim nbrTr As Long, nbrDep As Long

    Dim modal As ProgressBar
    Set modal = NewProgressBar("Import Generic CSV in progress", 9)
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True, local:=True
    importFrom = ActiveWorkbook.name
    Call AccountImportMetadata(ActiveSheet, accountId, exportDate, accNbr, Bank, avail, cur, typ, TaxRate, nbrTr, nbrDep)
    modal.Update

    Dim offset As Integer
    offset = 0
    If cur <> GetGlobalParam("DefaultCurrency", Workbooks(importTo)) Then
        offset = BL_FOREIGN_OFFSET
    End If

    Dim balanceTbl As ListObject, depositsTbl As ListObject
    Set balanceTbl = accountBalanceTable(accountId)
    Set depositsTbl = accountDepositTable(accountId)
    modal.Update
    
    Dim lastRow As String, firstRow As String
    lastRow = CStr(nbrTr + 1)

    Workbooks(importFrom).Activate
    Range("A2:A" & lastRow).Select
    Selection.Copy
    Workbooks(importTo).Activate
    balanceTbl.ListRows(1).Range(1, BL_DATE_COL).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    modal.Update
    
    Workbooks(importFrom).Activate
    Range("B2:B" & lastRow).Select
    Selection.Copy
    Workbooks(importTo).Activate
    balanceTbl.ListRows(1).Range(1, BL_AMOUNT_COL + offset).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    modal.Update

    Workbooks(importFrom).Activate
    Range("C2:C" & lastRow).Select
    Selection.Copy
    Workbooks(importTo).Activate
    balanceTbl.ListRows(1).Range(1, BL_BALANCE_COL + offset).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    modal.Update

    Workbooks(importFrom).Activate
    Range("D2:D" & lastRow).Select
    Selection.Copy
    Workbooks(importTo).Activate
    balanceTbl.ListRows(1).Range(1, BL_DESC_COL + offset).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    modal.Update

    Workbooks(importFrom).Activate
    Range("E2:E" & lastRow).Select
    Selection.Copy
    Workbooks(importTo).Activate
    balanceTbl.ListRows(1).Range(1, BL_SUBCATEG_COL + offset).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    modal.Update

    ' Read deposits part if any
    If nbrDep > 0 Then
        firstRow = CStr(nbrTr + 2)
        lastRow = CStr(nbrTr + nbrDep + 1)
        Workbooks(importFrom).Activate
        Range("A" & firstRow & ":A" & lastRow).Select
        Selection.Copy
        Workbooks(importTo).Activate
        depositsTbl.ListRows(1).Range(1, DP_DATE_COL).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        modal.Update

        Workbooks(importFrom).Activate
        Range("B" & firstRow & ":B" & lastRow).Select
        Selection.Copy
        Workbooks(importTo).Activate
        depositsTbl.ListRows(1).Range(1, DP_AMOUNT_COL).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        modal.Update
    Else
        modal.Update 2
    End If
    Application.DisplayAlerts = False
    Workbooks(importFrom).Close
    Application.DisplayAlerts = True
    Set modal = Nothing
End Sub


Public Sub AccountCreateFromCSV()
    Dim fileToOpen As Variant
    fileToOpen = Application.GetOpenFilename()
    If fileToOpen = False Then
        Exit Sub
    End If

    Dim importFrom As String, importTo As String
    importTo = ActiveWorkbook.name
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True, local:=True
    importFrom = ActiveWorkbook.name

    Dim accountId As String, cur As String, typ As String, avail As Integer, exportDate As String
    Dim Bank As String, TaxRate As Double, accNbr As String
    Dim nbrTr As Long, nbrDep As Long
    
    Call AccountImportMetadata(Workbooks(importFrom).ActiveSheet, accountId, exportDate, accNbr, Bank, avail, cur, _
        typ, TaxRate, nbrTr, nbrDep)
    
    Workbooks(importFrom).Close SaveChanges:=False
    Workbooks(importTo).Activate
    Call AccountCreate(accountId, cur, typ, avail, accNbr, Bank)
    Sheets(accountId).Activate
    Call ImportGeneric(accountId, fileToOpen)
End Sub


Public Sub AccountExportHere()
    Dim oAccount As Account
    Set oAccount = LoadAccount(getAccountId(ActiveSheet))
    If Not (oAccount Is Nothing) Then
        oAccount.Export
    End If
End Sub


Public Sub AccountExportAll()
    Dim sFolder As String
    Dim filename As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    If LenB(sFolder) = 0 Then ' if no directory was chosen
        Call ErrorMessage("k.warningExportCancelled")
        Exit Sub
    End If
    
    Dim modal As ProgressBar
    Dim ws As Worksheet, curWs As Worksheet
    Set modal = NewProgressBar("Export all accounts in progress", Worksheets.Count)
    Call FreezeDisplay
    Set curWs = ActiveSheet
    
    For Each ws In Worksheets
        Dim oAccount As Account
        Set oAccount = LoadAccount(getAccountId(ws))
        If Not (oAccount Is Nothing) Then
            filename = sFolder & "\" & ws.name & ".csv"
            Call oAccount.Export(csvFile:=filename, silent:=True)
        End If
        modal.Update
    Next ws
    curWs.Activate
    Set modal = Nothing
    Call UnfreezeDisplay
End Sub



Private Sub AccountExportMetadata(accountId As String, targetWs As Worksheet, nbrTransactions As Long, Optional nbrDeposits As Long = 0)
    ' Copy metadata on row 1
    targetWs.Range("A1") = "ExportDate=" & Format$(Now(), "YYYY-mm-dd HH:MM:SS")
    targetWs.Range("B1") = "AccountId=" & accountId
    targetWs.Range("C1") = "AccountNumber=" & AccountNumber(accountId)
    targetWs.Range("D1") = "Bank=" & AccountBank(accountId)
    avail = AccountAvailability(accountId)
    targetWs.Range("E1") = "Availability=" & avail
    targetWs.Range("F1") = "Currency=" & AccountCurrency(accountId)
    targetWs.Range("G1") = "Type=" & AccountType(accountId)
    targetWs.Range("H1") = "TaxRate=" & AccountTaxRate(accountId)
    targetWs.Range("I1") = "NbrTransactions=" & CStr(nbrTransactions)
    If nbrDeposits > 0 Then
        targetWs.Range("J1") = "NbrDeposits=" & CStr(nbrDeposits)
    End If
End Sub


Private Sub AccountImportMetadata(ws As Worksheet, accountId As String, exportDate As String, accNumber As String, _
    Bank As String, avail As Integer, accCurrency As String, accType As String, TaxRate As Double, _
    nbrTransactions As Long, nbrDeposits As Long)
    ' Copy metadata on row 1
    nbrTransactions = 0
    nbrDeposits = 0
    Dim i As Long
    i = 1
    Do While LenB(ws.Cells(1, i).value) > 0
        a = Split(ws.Cells(1, i).value, "=", -1, vbTextCompare)
        If a(0) = "AccountId" Then
            accountId = a(1)
        ElseIf a(0) = "ExportDate" Then
            exportDate = a(1)
        ElseIf a(0) = "AccountNumber" Then
            accNumber = a(1)
        ElseIf a(0) = "Bank" Then
            Bank = a(1)
        ElseIf a(0) = "Availability" Then
            avail = CInt(a(1))
        ElseIf a(0) = "Currency" Then
            accCurrency = a(1)
        ElseIf a(0) = "Type" Then
            accType = a(1)
        ElseIf a(0) = "TaxRate" Then
            TaxRate = CDbl(a(1))
        ElseIf a(0) = "NbrTransactions" Then
            nbrTransactions = CLng(a(1))
        ElseIf a(0) = "NbrDeposits" Then
            nbrDeposits = CLng(a(1))
        End If
        i = i + 1
    Loop
End Sub


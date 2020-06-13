Attribute VB_Name = "AccountMgr"

Public Const CHF_FORMAT = "#,###,##0.00"" CHF "";-#,###,##0.00"" CHF "";0.00"" CHF """
Public Const EUR_FORMAT = "#,###,##0.00"" € "";-#,###,##0.00"" € "";0.00"" € """
Public Const USD_FORMAT = "#,###,##0.00"" $ "";-#,###,##0.00"" $ "";0.00"" $ """

Public Const NOT_AN_ACCOUNT As Integer = 0
Public Const DOMESTIC_ACCOUNT As Integer = 1
Public Const FOREIGN_ACCOUNT As Integer = 2
Public Const DOMESTIC_SHARES_ACCOUNT As Integer = 3
Public Const FOREIGN_SHARES_ACCOUNT As Integer = 4

Public Const DATE_KEY As String = "k.date"
Public Const ACCOUNT_NAME_KEY As String = "k.accountName"
Public Const AMOUNT_KEY As String = "k.amount"
Public Const BALANCE_KEY As String = "k.accountBalance"
Public Const DESCRIPTION_KEY As String = "k.description"
Public Const SUBCATEGORY_KEY As String = "k.subcategory"
Public Const CATEGORY_KEY As String = "k.category"
Public Const IN_BUDGET_KEY As String = "k.inBudget"
Public Const SPREAD_KEY As String = "k.amountSpread"

Public Const PARAMS_SHEET As String = "Paramètres"
Public Const ACCOUNTS_SHEET As String = "Comptes"
Public Const MERGE_SHEET As String = "Comptes Merge"
Public Const BALANCE_SHEET As String = "Solde"

Public Const ACCOUNT_CLOSED As Integer = 0
Public Const ACCOUNT_OPEN As Integer = 1

Const ACCOUNT_NAME_LABEL = "A1"
Const ACCOUNT_NAME_VALUE = "B1"
Const ACCOUNT_NBR_LABEL = "A2"
Const ACCOUNT_NBR_VALUE = "B2"
Const ACCOUNT_BANK_LABEL = "A3"
Const ACCOUNT_BANK_VALUE = "B3"
Const ACCOUNT_STATUS_LABEL = "A4"
Const ACCOUNT_STATUS_VALUE = "B4"
Const ACCOUNT_AVAIL_LABEL = "A5"
Const ACCOUNT_AVAIL_VALUE = "B5"
Const ACCOUNT_CURRENCY_LABEL = "A6"
Const ACCOUNT_CURRENCY_VALUE = "B6"
Const ACCOUNT_TYPE_LABEL = "A7"
Const ACCOUNT_TYPE_VALUE = "B7"
Const IN_BUDGET_LABEL = "A8"
Const IN_BUDGET_VALUE = "B8"

Const DATE_COL = "A"
Const AMOUNT_COL = "B"
Const BALANCE_COL = "C"

Const OPEN_ACCOUNTS_TABLE = "tblOpenAccounts"
Const ACCOUNTS_TABLE = "tblAccounts"

Public Sub mergeAccounts()

    Call freezeDisplay

    For Each colKey In Array(DATE_KEY, ACCOUNT_NAME_KEY, AMOUNT_KEY, DESCRIPTION_KEY, SUBCATEGORY_KEY, IN_BUDGET_KEY)
        col = GetColName(colKey)
        firstAccount = True
        Dim array1d() As Variant
        For Each ws In Worksheets
           ' Make sure the sheet is not a template or anything else than an account
           If (isAnAccountSheet(ws)) Then
                'MsgBox (ws + " " + colNbr)
                ' Loop on all accounts of the sheet
                If (colKey = ACCOUNT_NAME_KEY) Then
                    arr1d = Create1DArray(ws.ListObjects(1).ListRows.Count, ws.Cells(1, 2).Value)
                ElseIf (colKey = IN_BUDGET_KEY And Not isAccountInBudget(ws.name)) Then
                    arr1d = Create1DArray(ws.ListObjects(1).ListRows.Count, 0)
                Else
                    arr1d = getTableColumn(ws.ListObjects(1), col, False)
                End If
                If (firstAccount) Then
                   totalColumn = arr1d
                   firstAccount = False
                Else
                   ret = ConcatenateArrays(totalColumn, arr1d)
                End If
            End If
        Next ws
        Call setTableColumn(Sheets(MERGE_SHEET).ListObjects(1), col, totalColumn, False)
        Erase totalColumn
    Next colKey
    ActiveWorkbook.Worksheets(MERGE_SHEET).ListObjects("AccountsMerge").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(MERGE_SHEET).ListObjects("AccountsMerge").Sort. _
        SortFields.Add key:=Range("AccountsMerge[[#Headers],[#Data],[Date]]"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(MERGE_SHEET).ListObjects("AccountsMerge"). _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets(MERGE_SHEET).PivotTables(1).PivotCache.Refresh

End Sub

Public Sub genBudget()

    Dim newSize As Integer
    'Sheets(MERGE_SHEET).ListObjects(1).Range.AutoFilter.ShowAllData
    nbRows = Sheets(MERGE_SHEET).ListObjects(1).ListRows.Count
    newSize = nbRows
    nbCols = Sheets(MERGE_SHEET).ListObjects(1).ListColumns.Count

    Dim oTable As Variant
    Dim dateCol() As Variant
    Dim accountCol() As Variant
    Dim amountCol() As Variant
    Dim descCol() As Variant
    Dim categCol() As Variant
    Dim spreadCol() As Variant

    With Sheets(MERGE_SHEET)
        dateCol = getTableColumn(.ListObjects(1), GetColName(DATE_KEY), False)
        accountCol = getTableColumn(.ListObjects(1), GetColName(ACCOUNT_NAME_KEY), False)
        amountCol = getTableColumn(.ListObjects(1), GetColName(AMOUNT_KEY), False)
        descCol = getTableColumn(.ListObjects(1), GetColName(DESCRIPTION_KEY), False)
        categCol = getTableColumn(.ListObjects(1), GetColName(SUBCATEGORY_KEY), False)
        spreadCol = getTableColumn(.ListObjects(1), GetColName(IN_BUDGET_KEY), False)
    End With

    Dim moreRows As Integer
    moreRows = 0
    For i = 1 To nbRows
        divider = spreadCol(i)
        If (IsNumeric(divider) And Int(divider) = divider And divider <> 1 And divider <> 0) Then
            moreRows = moreRows + divider - 1
        End If
    Next i
    ReDim Preserve dateCol(1 To nbRows + moreRows)
    ReDim Preserve accountCol(1 To nbRows + moreRows)
    ReDim Preserve amountCol(1 To nbRows + moreRows)
    ReDim Preserve descCol(1 To nbRows + moreRows)
    ReDim Preserve categCol(1 To nbRows + moreRows)
    ReDim Preserve spreadCol(1 To nbRows + moreRows)

    For i = 1 To nbRows
        divider = spreadCol(i)
        If LenB(divider) = 0 Then
            spreadCol(i) = -amountCol(i)
        End If
        If (IsNumeric(divider) And Int(divider) = divider And divider <> 1 And divider <> 0) Then
            newDate = dateCol(i)
            m = Month(newDate)
            y = Year(newDate)
            spreadCol(i) = -amountCol(i) / divider
            For k = 1 To divider - 1
                newSize = newSize + 1
                accountCol(newSize) = accountCol(i)
                spreadCol(newSize) = 1
                descCol(newSize) = descCol(i)
                categCol(newSize) = categCol(i)
                spreadCol(newSize) = -amountCol(i) / divider
                If (m >= 12) Then
                    m = 1
                    y = y + 1
                Else
                    m = m + 1
                End If
                dateCol(newSize) = DateSerial(y, m, 1)
            Next k
        End If
    Next i

    With Sheets(MERGE_SHEET)
        Call resizeTable(.ListObjects(1), nbRows + moreRows)
        Call setTableColumn(.ListObjects(1), GetColName(DATE_KEY), dateCol, False)
        Call setTableColumn(.ListObjects(1), GetColName(ACCOUNT_NAME_KEY), accountCol, False)
        Call setTableColumn(.ListObjects(1), GetColName(AMOUNT_KEY), amountCol, False)
        Call setTableColumn(.ListObjects(1), GetColName(DESCRIPTION_KEY), descCol, False)
        Call setTableColumn(.ListObjects(1), GetColName(SUBCATEGORY_KEY), categCol, False)
        Call setTableColumn(.ListObjects(1), GetColName(SPREAD_KEY), spreadCol, False)
        Call resizeTable(.ListObjects(1), newSize)
        .PivotTables(1).PivotCache.Refresh
    End With

End Sub

Public Sub refreshAllAccounts()
    Call freezeDisplay
    Call mergeAccounts
    Call genBudget
    Call unfreezeDisplay
End Sub

Sub CreateAccount()
    accountNbr = InputBox("Account number ?", "Account Number", "<accountNumber>")
    accountName = InputBox("Account name ?", "Account Name", "<accountName>")
    Sheets("Account Template").Visible = True
    Sheets("Account Template").Copy Before:=Sheets(1)
    Sheets("Account Template").Visible = False
    With Sheets(1)
        .name = accountName
        ' .Range("A1").Formula = "=VLOOKUP("k.account", TblKeys, LangId, FALSE)"
        .Range(ACCOUNT_NAME_VALUE).Value = accountName
        formulaRoot = "=VLOOKUP(B$1," & ACCOUNTS_TABLE
        .Range(ACCOUNT_NBR_VALUE).Formula = formulaRoot & ",2,FALSE)"
        .Range(ACCOUNT_BANK_VALUE).Formula = formulaRoot & ",4,FALSE)"
        .Range(ACCOUNT_STATUS_VALUE).Formula = formulaRoot & ",6,FALSE)"
        .Range(ACCOUNT_AVAIL_VALUE).Formula = formulaRoot & ",5,FALSE)"
    End With
End Sub


Public Sub doForAllAccounts()
'
' Applies a given macro to all account sheets
'
'
    Call showAllSheets
    For Each ws In Worksheets
       ' Make sure the sheet is not anything else than an account
        If (isAnAccountSheet(ws) Or isTemplate(ws)) Then
            ws.Select
            ' Call macro here
        End If
    Next ws
    Call hideClosedAccounts
    Call hideTemplateAccounts
End Sub
'-------------------------------------------------
Public Sub formatAccountSheets()
'
'  Reformat all account sheets
'
   For Each ws In Worksheets
       ' Make sure the sheet is not anything else than an account
       If (isAnAccountSheet(ws) Or isTemplate(ws)) Then
            Dim name As String
            Dim col As Integer
            name = ws.name
            col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(DATE_KEY))
            If col <> 0 Then
                Call SetColumnWidth(Chr(col + 64), 15, name)
                ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = "m/d/yyyy"
            End If
            col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(AMOUNT_KEY))
            If col <> 0 Then
                Call SetColumnWidth(Chr(col + 64), 15, name)
                ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = EUR_FORMAT
            End If
            col = GetColumnNumberFromName(ws.ListObjects(1), "Montant CHF")
            If col <> 0 Then
                Call SetColumnWidth(Chr(col + 64), 15, name)
                ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = CHF_FORMAT
            End If
            col = GetColumnNumberFromName(ws.ListObjects(1), "Montant USD")
            If col <> 0 Then
                Call SetColumnWidth(Chr(col + 64), 15, name)
                ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = USD_FORMAT
            End If
            col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(BALANCE_KEY))
            If col <> 0 Then
                Call SetColumnWidth(Chr(col + 64), 18, name)
                ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = EUR_FORMAT
            End If
            col = GetColumnNumberFromName(ws.ListObjects(1), "Solde CHF")
            If col <> 0 Then
                Call SetColumnWidth(Chr(col + 64), 18, name)
                ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = CHF_FORMAT
            End If
            col = GetColumnNumberFromName(ws.ListObjects(1), "Solde USD")
            If col <> 0 Then
                Call SetColumnWidth(Chr(col + 64), 18, name)
                ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = CHF_FORMAT
            End If
            col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(DESCRIPTION_KEY))
            If col <> 0 Then
                Call SetColumnWidth(Chr(col + 64), 70, name)
            End If
            col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(SUBCATEGORY_KEY))
            If col <> 0 Then
                Call SetColumnWidth(Chr(col + 64), 15, name)
            End If
            col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(CATEGORY_KEY))
            If col <> 0 Then
                Call SetColumnWidth(Chr(col + 64), 15, name)
            End If
            col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(IN_BUDGET_KEY))
            If col <> 0 Then
                Call SetColumnWidth(Chr(col + 64), 5, name)
                Call SetColumnWidth(Chr(col + 65), 5, name)
            End If
          ws.Cells.RowHeight = 13
          ws.Rows.Font.size = 10

          If (ws.Shapes.Count > 0) Then
            Dim i As Integer
            i = 0
            For Each Shape In ws.Shapes
                If (Shape.Type = msoFormControl) Then
                    ' This is a button, move it to right place
                    row = i Mod 4
                    col = i \ 4
                    Call ShapePlacementXY(Shape, 300 + col * 100, 5 + row * 22, 400 + col * 100, 25 + row * 22)
                    i = i + 1
                End If
            Next Shape
          End If
       End If
   Next ws
   Call hideClosedAccounts
   Call hideTemplateAccounts
End Sub

'-------------------------------------------------
Public Function isTemplate(ByVal ws As Worksheet) As Boolean
    isTemplate = (ws.Cells(1, 2).Value = "TEMPLATE")
End Function

'-------------------------------------------------
Private Sub setClosedAccountsVisibility(visibility)
    For Each ws In Worksheets
        If (isClosed(ws.name)) Then
            ws.Visible = visibility
        End If
    Next ws
End Sub

'-------------------------------------------------
Public Sub hideClosedAccounts()
    If (ThisWorkbook.Names("hideClosedAccounts").RefersToRange.Value = 1) Then
        Call setClosedAccountsVisibility(False)
    End If
End Sub

'-------------------------------------------------
Public Sub showClosedAccounts()
    Call setClosedAccountsVisibility(True)
End Sub

'-------------------------------------------------
Private Sub setTemplateAccountsVisibility(visibility)
    For Each ws In Worksheets
        If isTemplate(ws) Then
            ws.Visible = visibility
        End If
    Next ws
End Sub
'-------------------------------------------------
Public Sub hideTemplateAccounts()
    Call setTemplateAccountsVisibility(False)
End Sub
'-------------------------------------------------
Public Sub showTemplateAccounts()
    Call setTemplateAccountsVisibility(True)
End Sub
Public Sub refreshOpenAccountsList()
    Call freezeDisplay
    Call truncateTable(Sheets(PARAMS_SHEET).ListObjects(TABLE_OPEN_ACCOUNTS))
    With Sheets(PARAMS_SHEET).ListObjects(TABLE_OPEN_ACCOUNTS)
        For i = 1 To Sheets(ACCOUNTS_SHEET).ListObjects(TABLE_ACCOUNTS).ListRows.Count
            If (Sheets(ACCOUNTS_SHEET).ListObjects(TABLE_ACCOUNTS).ListRows(i).Range.Cells(1, 6).Value = "Open") Then
                .ListRows.Add ' Add 1 row at the end, then extend
                .ListRows(.ListRows.Count).Range.Cells(1, 1).Value = Sheets(ACCOUNTS_SHEET).ListObjects(TABLE_ACCOUNTS).ListRows(i).Range.Cells(1, 1).Value
            End If
        Next i
        nbrAccounts = .ListRows.Count + 1
    End With
    ActiveSheet.Shapes("Drop Down 2").Select
    With Selection
        .ListFillRange = PARAMS_SHEET & "!$L$2:$L$" & CStr(Sheets(PARAMS_SHEET).ListObjects(TABLE_OPEN_ACCOUNTS).ListRows.Count + 1)
        .LinkedCell = "$H$72"
        .DropDownLines = 8
        .Display3DShading = True
    End With
    Call unfreezeDisplay
End Sub

Public Sub sortCurrentAccount()
    Call sortAccount(ActiveSheet.ListObjects(1))
End Sub
Public Sub sortAccount(oTable)
    oTable.Sort.SortFields.Clear
    ' Sort table by date first, then by amount
    oTable.Sort.SortFields.Add key:=Range(oTable.name & "[" & GetLabel("k.date") & "]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    oTable.Sort.SortFields.Add key:=Range(oTable.name & "[" & GetLabel("k.amount") & "]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With oTable.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ' Reset date column format
    Call setTableColumnFormat(oTable, 1, "m/d/yyyy")
End Sub
'-------------------------------------------------
Public Function accountType(accountName As String) As String
    If (accountName = "Account Template") Then
        accountType = "Standard"
    ElseIf (Not accountExists(accountName)) Then
        accountType = "ERROR: Not an account"
    ElseIf (Sheets(accountName).Range("B6").Value = "EUR") Then
        accountType = Sheets(accountName).Range("B7").Value
    End If
End Function
'-------------------------------------------------
Public Function accountNumber(accountName As String) As String
    If (accountExists(accountName)) Then
        accountNumber = Sheets(accountName).Range(ACCOUNT_NBR_VALUE).Value
    Else
        accountNumber = ""
    End If
End Function
'-------------------------------------------------
Public Function accountBank(accountName As String) As String
    If (accountExists(accountName)) Then
        accountBank = Sheets(accountName).Range(ACCOUNT_BANK_VALUE).Value
    Else
        accountBank = ""
    End If
End Function

'-------------------------------------------------
Public Function accountStatus(accountName As String) As String
    If (accountExists(accountName)) Then
        accountStatus = Sheets(accountName).Range(ACCOUNT_STATUS_VALUE).Value
    Else
        accountStatus = ""
    End If
End Function
'-------------------------------------------------
Public Function accountAvailability(accountName As String) As String
    If (accountExists(accountName)) Then
        accountAvailability = Sheets(accountName).Range(ACCOUNT_AVAIL_VALUE).Value
    Else
        accountAvailability = ""
    End If
End Function
'-------------------------------------------------
Public Function accountCurrency(accountName As String) As String
    If (accountExists(accountName)) Then
        accountCurrency = Sheets(accountName).Range(ACCOUNT_CURRENCY_VALUE).Value
    Else
        accountCurrency = ""
    End If
End Function
'-------------------------------------------------
Public Function isAccountInBudget(accountName As String) As Boolean
    isAccountInBudget = (accountExists(accountName) And Sheets(accountName).Range(IN_BUDGET_VALUE).Value = "Yes")
End Function
'-------------------------------------------------
Public Function isOpen(accountName As String) As Boolean
    isOpen = (accountStatus(accountName) = "Open")
End Function

Public Function isClosed(accountName As String) As Boolean
    isClosed = Not isOpen(accountName)
End Function


Public Sub AddSavingsRow()
    Call AddInvestmentRow(ActiveSheet.ListObjects(1))
End Sub

Private Sub AddInvestmentRow(oTable)
    oTable.ListRows.Add
    nbRows = oTable.ListRows.Count
    
    col = GetColumnNumberFromName(oTable, GetLabel(DATE_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).FormulaR1C1 = Date
    
    col = GetColumnNumberFromName(oTable, GetLabel(BALANCE_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).Value = oTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).Value
    
    col = GetColumnNumberFromName(oTable, GetLabel(SUBCATEGORY_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).Value = oTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).Value
    
    col = GetColumnNumberFromName(oTable, GetLabel(AMOUNT_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).FormulaR1C1 = "=[Solde]-R[-1]C[1]"
    
    col = GetColumnNumberFromName(oTable, GetLabel(DESCRIPTION_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).FormulaR1C1 = "=CONCATENATE(""Delta solde "",TEXT(R[-1]C[-3],date_format))"
End Sub

'-------------------------------------------------
Public Function accountExists(accountName As String) As Boolean
    accountExists = (sheetExists(accountName) And Sheets(accountName).Range(ACCOUNT_NAME_LABEL) = GetLabel("k.accountName"))
End Function
'-------------------------------------------------
Public Function isAnAccountSheet(ByVal ws As Worksheet) As Boolean
    isAnAccountSheet = (ws.Cells(1, 1).Value = getNamedVariableValue("accountIdentifier") And Not isTemplate(ws))
End Function

'-------------------------------------------------
Public Sub showAllSheets()
    For Each ws In Worksheets
        ws.Visible = True
    Next ws
End Sub

Public Sub GoToSolde()
    Sheets(BALANCE_SHEET).Activate
End Sub

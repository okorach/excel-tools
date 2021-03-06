Attribute VB_Name = "AccountMgr"

Public Const CHF_FORMAT = "#,###,##0.00"" CHF "";-#,###,##0.00"" CHF "";0.00"" CHF """
Public Const EUR_FORMAT = "#,###,##0.00"" € "";-#,###,##0.00"" € "";0.00"" € """
Public Const USD_FORMAT = "#,###,##0.00"" $ "";-#,###,##0.00"" $ "";0.00"" $ """

Public Const NOT_AN_ACCOUNT As Long = 0
Public Const DOMESTIC_ACCOUNT As Long = 1
Public Const FOREIGN_ACCOUNT As Long = 2
Public Const DOMESTIC_SHARES_ACCOUNT As Long = 3
Public Const FOREIGN_SHARES_ACCOUNT As Long = 4

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

Public Const ACCOUNT_CLOSED As Long = 0
Public Const ACCOUNT_OPEN As Long = 1

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

Public Const OPEN_ACCOUNTS_TABLE = "tblOpenAccounts"
Public Const ACCOUNTS_TABLE = "tblAccounts"
Public Const SUBSTITUTIONS_TABLE = "TblSubstitutions"

Public Sub MergeAccounts()

    Dim firstAccount As Boolean
    Dim ws As Worksheet

    Call FreezeDisplay

    For Each colKey In Array(DATE_KEY, ACCOUNT_NAME_KEY, AMOUNT_KEY, DESCRIPTION_KEY, SUBCATEGORY_KEY, IN_BUDGET_KEY)
        Dim col As String
        col = GetColName(colKey)
        firstAccount = True
        Dim array1d() As Variant
        For Each ws In Worksheets
           ' Make sure the sheet is not a template or anything else than an account
           If (IsAnAccountSheet(ws)) Then
                'MsgBox (ws + " " + colNbr)
                ' Loop on all accounts of the sheet
                If (colKey = ACCOUNT_NAME_KEY) Then
                    arr1d = Create1DArray(ws.ListObjects(1).ListRows.Count, ws.Cells(1, 2).Value)
                ElseIf (colKey = IN_BUDGET_KEY And Not IsAccountInBudget(ws.name)) Then
                    arr1d = Create1DArray(ws.ListObjects(1).ListRows.Count, 0)
                Else
                    arr1d = GetTableColumn(ws.ListObjects(1), col)
                End If
                If (firstAccount) Then
                   totalColumn = arr1d
                   firstAccount = False
                Else
                   ret = ConcatenateArrays(totalColumn, arr1d)
                End If
            End If
        Next ws
        Call SetTableColumn(Sheets(MERGE_SHEET).ListObjects(1), col, totalColumn)
        Erase totalColumn
    Next colKey

    Call SortTable(Sheets(MERGE_SHEET).ListObjects("AccountsMerge"), GetLabel(DATE_KEY), xlAscending, GetLabel(AMOUNT_KEY), xlDescending)
    Sheets(MERGE_SHEET).PivotTables(1).PivotCache.Refresh

End Sub

Public Sub GenBudget()

    Dim i As Long
    Dim newSize As Long
    Dim nbRows As Long
    nbRows = Sheets(MERGE_SHEET).ListObjects(1).ListRows.Count
    newSize = nbRows

    Dim dateCol() As Variant
    Dim accountCol() As Variant
    Dim amountCol() As Variant
    Dim descCol() As Variant
    Dim categCol() As Variant
    Dim spreadCol() As Variant

    With Sheets(MERGE_SHEET)
        dateCol = GetTableColumn(.ListObjects(1), GetColName(DATE_KEY))
        accountCol = GetTableColumn(.ListObjects(1), GetColName(ACCOUNT_NAME_KEY))
        amountCol = GetTableColumn(.ListObjects(1), GetColName(AMOUNT_KEY))
        descCol = GetTableColumn(.ListObjects(1), GetColName(DESCRIPTION_KEY))
        categCol = GetTableColumn(.ListObjects(1), GetColName(SUBCATEGORY_KEY))
        spreadCol = GetTableColumn(.ListObjects(1), GetColName(IN_BUDGET_KEY))
    End With

    Dim moreRows As Long
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
            Dim k As Long
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
        Call ResizeTable(.ListObjects(1), nbRows + moreRows)
        Call SetTableColumn(.ListObjects(1), GetColName(DATE_KEY), dateCol)
        Call SetTableColumn(.ListObjects(1), GetColName(ACCOUNT_NAME_KEY), accountCol)
        Call SetTableColumn(.ListObjects(1), GetColName(AMOUNT_KEY), amountCol)
        Call SetTableColumn(.ListObjects(1), GetColName(DESCRIPTION_KEY), descCol)
        Call SetTableColumn(.ListObjects(1), GetColName(SUBCATEGORY_KEY), categCol)
        Call SetTableColumn(.ListObjects(1), GetColName(SPREAD_KEY), spreadCol)
        Call ResizeTable(.ListObjects(1), newSize)
        .PivotTables(1).PivotCache.Refresh
    End With

End Sub

Public Sub RefreshAllAccounts()
    Call FreezeDisplay
    Call MergeAccounts
    Call GenBudget
    Call UnfreezeDisplay
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
    Call ShowAllSheets
    For Each ws In Worksheets
       ' Make sure the sheet is not anything else than an account
        If (IsAnAccountSheet(ws) Or isTemplate(ws)) Then
            ws.Select
            ' Call macro here
        End If
    Next ws
    Call HideClosedAccounts
    Call hideTemplateAccounts
End Sub
'-------------------------------------------------
Public Sub FormatAccountSheet(ws As Worksheet)
' Make sure the sheet is not anything else than an account
    If (IsAnAccountSheet(ws) Or isTemplate(ws)) Then
        Dim name As String
        Dim col As Long
        name = ws.name
        col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(DATE_KEY))
        If col <> 0 Then
            Call SetColumnWidth(Chr$(col + 64), 15, name)
            ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = "m/d/yyyy"
        End If
        col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(AMOUNT_KEY))
        If col <> 0 Then
            Call SetColumnWidth(Chr$(col + 64), 15, name)
            ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = EUR_FORMAT
        End If
        col = GetColumnNumberFromName(ws.ListObjects(1), "Montant CHF")
        If col <> 0 Then
            Call SetColumnWidth(Chr$(col + 64), 17, name)
            ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = CHF_FORMAT
        End If
        col = GetColumnNumberFromName(ws.ListObjects(1), "Montant USD")
        If col <> 0 Then
            Call SetColumnWidth(Chr$(col + 64), 15, name)
            ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = USD_FORMAT
        End If
        col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(BALANCE_KEY))
        If col <> 0 Then
            Call SetColumnWidth(Chr$(col + 64), 18, name)
            ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = EUR_FORMAT
        End If
        col = GetColumnNumberFromName(ws.ListObjects(1), "Solde CHF")
        If col <> 0 Then
            Call SetColumnWidth(Chr$(col + 64), 18, name)
            ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = CHF_FORMAT
        End If
        col = GetColumnNumberFromName(ws.ListObjects(1), "Solde USD")
        If col <> 0 Then
            Call SetColumnWidth(Chr$(col + 64), 18, name)
            ws.ListObjects(1).ListColumns(col).DataBodyRange.NumberFormat = CHF_FORMAT
        End If
        col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(DESCRIPTION_KEY))
        If col <> 0 Then
            Call SetColumnWidth(Chr$(col + 64), 70, name)
        End If
        col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(SUBCATEGORY_KEY))
        If col <> 0 Then
            Call SetColumnWidth(Chr$(col + 64), 15, name)
        End If
        col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(CATEGORY_KEY))
        If col <> 0 Then
            Call SetColumnWidth(Chr$(col + 64), 15, name)
        End If
        col = GetColumnNumberFromName(ws.ListObjects(1), GetLabel(IN_BUDGET_KEY))
        If col <> 0 Then
            Call SetColumnWidth(Chr$(col + 64), 5, name)
            Call SetColumnWidth(Chr$(col + 65), 5, name)
        End If
        ws.Cells.RowHeight = 13
        ws.Rows.Font.size = 10
        
        Call formatAccountSheetButtons(ws)
    End If
End Sub

Private Sub formatAccountSheetButtons(ws As Worksheet)
    If (ws.Shapes.Count > 0) Then
        Dim home_x As Long
        Dim home_y As Long
        Dim btn_height As Long
        Dim i As Long
        home_x = 200
        home_y = 10
        btn_height = 22
        i = 0
        Dim s As Shape
        For Each s In ws.Shapes
            If s.name = "BtnPrev5" Then
                Call ShapePlacementXY(s, home_x, home_y, home_x + 29, home_y + btn_height - 1)
            ElseIf s.name = "BtnPrev" Then
                Call ShapePlacementXY(s, home_x + 30, home_y, home_x + 59, home_y + btn_height - 1)
            ElseIf s.name = "BtnHome" Then
                Call ShapePlacementXY(s, home_x + 60, home_y, home_x + 129, home_y + btn_height - 1)
            ElseIf s.name = "BtnNext" Then
                Call ShapePlacementXY(s, home_x + 130, home_y, home_x + 159, home_y + btn_height - 1)
            ElseIf s.name = "BtnNext5" Then
                Call ShapePlacementXY(s, home_x + 160, home_y, home_x + 189, home_y + btn_height - 1)
            ElseIf s.name = "BtnTop" Then
                Call ShapePlacementXY(s, home_x, home_y + btn_height, home_x + 99, home_y + 2 * btn_height - 1)
            ElseIf s.name = "BtnBottom" Then
                Call ShapePlacementXY(s, home_x, home_y + 2 * btn_height, home_x + 99, home_y + 3 * btn_height - 1)
            ElseIf s.name = "BtnSort" Then
                Call ShapePlacementXY(s, home_x, home_y + 3 * btn_height, home_x + 99, home_y + 4 * btn_height - 1)
            ElseIf s.name = "BtnInterests" Then
                Call ShapePlacementXY(s, home_x + 100, home_y + 3 * btn_height, home_x + 199, home_y + 4 * btn_height - 1)
            ElseIf s.name = "BtnImport" Then
                Call ShapePlacementXY(s, home_x + 100, home_y + btn_height, home_x + 199, home_y + 2 * btn_height - 1)
            ElseIf s.name = "BtnAddEntry" Then
                Call ShapePlacementXY(s, home_x + 100, home_y + 2 * btn_height, home_x + 199, home_y + 3 * btn_height - 1)

            ElseIf (s.Type = msoFormControl) Then
                ' This is a button, move it to right place
                row = i Mod 4
                col = i \ 4
                Call ShapePlacementXY(s, 300 + col * 100, 5 + row * 22, 400 + col * 100, 25 + row * 22)
                i = i + 1
            End If
        Next s
    End If
End Sub
Public Sub formatAllAccountSheets()
'
'  Reformat all account sheets
'
   For Each ws In Worksheets
       Call FormatAccountSheet(ws)
   Next ws
   Call HideClosedAccounts
   Call hideTemplateAccounts
End Sub

'-------------------------------------------------
Public Function isTemplate(ByVal ws As Worksheet) As Boolean
    isTemplate = (ws.Cells(1, 2).Value = "TEMPLATE")
End Function

'-------------------------------------------------
Private Sub setClosedAccountsVisibility(visibility)
    For Each ws In Worksheets
        If IsClosed(ws.name) Then
            ws.Visible = visibility
        End If
    Next ws
End Sub

'-------------------------------------------------
Public Sub HideClosedAccounts()
    If GetNamedVariableValue("hideClosedAccounts") = 1 Then
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
    Call FreezeDisplay
    Call TruncateTable(Sheets(PARAMS_SHEET).ListObjects(TABLE_OPEN_ACCOUNTS))
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
    Call UnfreezeDisplay
End Sub

Public Sub sortCurrentAccount()
    Call sortAccount(ActiveSheet.ListObjects(1))
    Call SortTable(ActiveSheet.ListObjects(1), GetLabel(DATE_KEY), xlAscending, GetLabel(AMOUNT_KEY), xlDescending)
    Call SetTableColumnFormat(ActiveSheet.ListObjects(1), 1, "m/d/yyyy")
End Sub

'-------------------------------------------------
Public Function AccountType(accountId As String) As String
    Dim ws As Worksheet
    Set ws = getAccountSheet(accountId)
    If (accountId = "Account Template") Then
        AccountType = "Standard"
    ElseIf (Not AccountExists(accountId)) Then
        AccountType = "ERROR: Not an account"
    ElseIf (wsRange("B6").Value = "EUR") Then
        AccountType = ws.Range("B7").Value
    End If
End Function
'-------------------------------------------------
Private Function AccountAttribute(accountId As String, attributeCell As String) As String
    AccountAttribute = ""
    If (AccountExists(accountId)) Then
        Dim ws As Worksheet
        Set ws = getAccountSheet(accountId)
        AccountAttribute = ws.Range(ACCOUNT_NBR_VALUE).Value
    End If
End Function
Public Function AccountNumber(accountId As String) As String
    AccountNumber = AccountAttribute(accountId, ACCOUNT_NBR_VALUE)
End Function

Public Function AccountBank(accountId As String) As String
    AccountBank = AccountAttribute(accountId, ACCOUNT_BANK_VALUE)
End Function

Public Function AccountStatus(accountId As String) As String
    AccountStatus = AccountAttribute(accountId, ACCOUNT_STATUS_VALUE)
End Function

Public Function AccountAvailability(accountId As String) As String
    AccountAvailability = AccountAttribute(accountId, ACCOUNT_AVAIL_VALUE)
End Function

Public Function AccountCurrency(accountId As String) As String
    AccountCurrency = AccountAttribute(accountId, ACCOUNT_CURRENCY_VALUE)
End Function

'-------------------------------------------------
Public Function IsAccountInBudget(accountId As String) As Boolean
    IsAccountInBudget = (AccountExists(accountId) And Sheets(accountId).Range(IN_BUDGET_VALUE).Value = "Yes")
End Function
'-------------------------------------------------
Public Function IsOpen(accountId As String) As Boolean
    IsOpen = (AccountStatus(accountId) = "Open")
End Function

Public Function IsClosed(accountId As String) As Boolean
    IsClosed = Not IsOpen(accountId)
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
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).FormulaR1C1 = oTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).FormulaR1C1
    
    col = GetColumnNumberFromName(oTable, GetLabel(DESCRIPTION_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).FormulaR1C1 = oTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).FormulaR1C1
End Sub

'-------------------------------------------------
Public Function AccountExists(accountId As String) As Boolean
    AccountExists = (SheetExists(accountId) And Sheets(accountId).Range(ACCOUNT_NAME_LABEL) = GetLabel(ACCOUNT_NAME_KEY))
End Function
'-------------------------------------------------
Public Function IsAnAccountSheet(ByVal ws As Worksheet) As Boolean
    IsAnAccountSheet = (ws.Cells(1, 1).Value = GetNamedVariableValue("accountIdentifier") And Not isTemplate(ws))
End Function

Public Function AccountDepositHistory(accountId As String) As Variant
    AccountDepositHistory = getDepositArray(accountId)
End Function

Public Function AccountBalanceHistory(accountId As String, Optional sampling As String = "Yearly") As Variant
    Dim histAll() As Variant
    Dim histSampled() As Variant
    Dim nbYears As Long
    Dim i As Long
    Dim j As Long
    Dim lastMonth As Long
    Dim lastYear As Long
    Dim lastBalance As Double
    Dim histSize As Long
    histAll = AccountBalanceArray(accountId)
    histSize = UBound(histAll, 1)
    nbYears = Year(histAll(histSize, 1)) - Year(histAll(1, 1)) + 2
    ReDim histSampled(1 To nbYears, 1 To 2)
    lastMonth = 0
    lastYear = Year(histAll(1, 1)) - 1
    lastBalance = 0
    j = 1
    For i = 1 To histSize
        m = Month(histAll(i, 1))
        y = Year(histAll(i, 1))
        If sampling = "Monthly" And m <> lastMonth Then
            d = Day(histAll(i, 1))
            histSampled(j, 1) = DateSerial(y, m, 1)
            If d <> 1 Then
                histSampled(j, 2) = lastBalance
            Else
                histSampled(j, 2) = histAll(i, 3)
            End If
            j = j + 1
        ElseIf sampling = "Yearly" And y <> lastYear Then
            While y <> lastYear
                histSampled(j, 1) = DateSerial(lastYear, 12, 31)
                histSampled(j, 2) = lastBalance
                j = j + 1
                lastYear = lastYear + 1
            Wend
        End If
        lastMonth = m
        lastBalance = histAll(i, 3)
        lastYear = y
    Next i
    If sampling = "Yearly" Then
        histSampled(j, 1) = histAll(histSize, 1)
        histSampled(j, 2) = lastBalance
    End If
    AccountBalanceHistory = histSampled
End Function


Public Sub CalcAccountInterests(accountId As String)
    Dim deposits As Variant
    Dim balances As Variant
    deposits = AccountDepositHistory(accountId)
    balances = AccountBalanceHistory(accountId, "Yearly")
    Call StoreAccountInterests(accountId, InterestsCalc(balances, deposits))
End Sub


Public Sub CalcInterestForAllAccounts()
    Dim accountId As String
    Dim ws As Worksheet
    FreezeDisplay
    For Each ws In Worksheets
        accountId = getAccountId(ws)
        If IsAnAccountSheet(ws) And IsOpen(accountId) Then
            Call CalcAccountInterests(accountId)
        End If
    Next i
    UnfreezeDisplay
End Sub

Public Sub StoreAccountInterests(accountId As String, yields As Variant)
    Dim nbrYields As Long
    Dim yieldIndex As Long
    Dim ws As Worksheet
    nbrYields = UBound(yields)
    yieldIndex = AccountYieldsTableIndex(accountId)
    Set ws = getAccountSheet(accountId)
    With ws
        .ListObjects(yieldIndex).ListColumns(2).DataBodyRange.Rows(1).Value = yields(nbrYields)
        For k = 2 To 5
         .ListObjects(yieldIndex).ListColumns(2).DataBodyRange.Rows(k).Value = "-"
        Next k
        If nbrYields >= 2 Then
            .ListObjects(yieldIndex).ListColumns(2).DataBodyRange.Rows(2).Value = yields(nbrYields - 1)
        End If
        If nbrYields >= 4 Then
            .ListObjects(yieldIndex).ListColumns(2).DataBodyRange.Rows(3).Value = ArrayAverage(yields, nbrYields - 3, nbrYields - 1)
        End If
        If nbrYields >= 6 Then
            .ListObjects(yieldIndex).ListColumns(2).DataBodyRange.Rows(4).Value = ArrayAverage(yields, nbrYields - 5, nbrYields - 1)
        End If
        If nbrYields >= 2 Then
            .ListObjects(yieldIndex).ListColumns(2).DataBodyRange.Rows(5).Value = ArrayAverage(yields, 1, nbrYields - 1)
        End If
    End With
End Sub


Private Function AccountTableIndexFromSuffix(accountId As String, suffix As String) As Long
    Dim ws As Worksheet
    Dim i As Long
    Set ws = getAccountSheet(accountId)
    For i = 1 To ws.ListObjects.Count
        If LCase$(ws.ListObjects(i).name) Like suffix & "_*" Then
            AccountTableIndexFromSuffix = i
            Exit For
        End If
    Next i
End Function

Public Function AccountTableArrayFromSuffix(accountId As String, suffix As String) As Variant
    Dim ws As Worksheet
    Dim i As Long
    Set ws = getAccountSheet(accountId)
    AccountTableArrayFromSuffix = Empty
    For i = 1 To ws.ListObjects.Count
        If LCase$(ws.ListObjects(i).name) Like suffix & "_*" Then
            AccountTableArrayFromSuffix = GetTableAsArray(ws.ListObjects(i))
            Exit For
        End If
    Next i
End Function

Public Function AccountBalanceTableIndex(accountId As String) As Long
    AccountBalanceTableIndex = AccountTableIndexFromSuffix(accountId, "balance")
End Function

Public Function AccountDepositsTableIndex(accountId As String) As Long
    AccountDepositsTableIndex = AccountTableIndexFromSuffix(accountId, "deposits")
End Function

Public Function AccountYieldsTableIndex(accountId As String) As Long
    AccountYieldsTableIndex = AccountTableIndexFromSuffix(accountId, "yields")
End Function

Public Function AccountBalanceArray(accountId As String) As Variant
    AccountBalanceArray = AccountTableArrayFromSuffix(accountId, "balance")
End Function

Public Function AccountDepositsArray(accountId As String) As Variant
    AccountDepositsArray = AccountTableArrayFromSuffix(accountId, "deposits")
End Function

Public Function AccountYieldsArray(accountId As String) As Variant
    AccountYieldsArray = AccountTableArrayFromSuffix(accountId, "yields")
End Function

Public Function getSelectedAccount() As String
    selectedNbr = GetNamedVariableValue("selectedAccount")
    getSelectedAccount = Sheets(PARAMS_SHEET).Range("L" & CStr(selectedNbr + 1))
End Function


'--------------------------------------------------------------------------
' Button methods
'--------------------------------------------------------------------------

Public Sub BtnAccountInterests()
    Call CalcAccountInterests(getAccountId(ActiveSheet))
End Sub

'--------------------------------------------------------------------------
' Private methods
'--------------------------------------------------------------------------

Private Function getAccountArray(accountId As String, tableType As String) As Variant
    Dim i As Long
    Dim ws As Worksheet
    getAccountArray = Empty
    Set ws = getAccountSheet(accountId)
    For i = 1 To ws.ListObjects.Count
        If ws.ListObjects(i).name Like tableType & "_*" Then
            getAccountArray = GetTableAsArray(ws.ListObjects(i))
            Exit For
        End If
    Next i
End Function

Private Function getDepositArray(accountId As String) As Variant
    getDepositArray = getAccountArray(accountId, "Deposits")
End Function

Private Function getTransactionArray(accountId As String) As Variant
    getTransactionArray = getAccountArray(accountId, "Transactions")
End Function

Private Function getAccountId(ws As Worksheet) As String
    getAccountId = ws.name
End Function

Private Function getAccountSheet(accountId As String) As Worksheet
    Set getAccountSheet = ThisWorkbook.Sheets(accountId)
End Function


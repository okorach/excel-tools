Attribute VB_Name = "AccountMgr"

Public Const CHF_FORMAT = "#,###,##0.00"" CHF "";-#,###,##0.00"" CHF "";0.00"" CHF """
Public Const EUR_FORMAT = "#,###,##0.00"" € "";-#,###,##0.00"" € "";0.00"" € """
Public Const USD_FORMAT = "#,###,##0.00"" $ "";-#,###,##0.00"" $ "";0.00"" $ """
Public Const DATE_FORMAT = "m/d/yyyy"

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
Public Const BALANCE_SHEET As String = "Solde"

Public Const ACCOUNT_TYPE_STANDARD As String = "Standard"
Public Const ACCOUNT_TYPE_SAVINGS As String = "Savings"
Public Const ACCOUNT_TYPE_INVESTMENT As String = "Investment"

Public Const ACCOUNT_CLOSED As Long = 0
Public Const ACCOUNT_OPEN As Long = 1

Private Const MAX_MERGE_SIZE As Long = 100000

Private Const ACCOUNTS_TABLE As String = "tblAccounts"
Private Const ACCOUNT_TYPES_TABLE As String = "tblAccountTypes"

Private Const ACCOUNT_NAME_LABEL As String = "A1"
Private Const ACCOUNT_NAME_VALUE As String = "B1"
Private Const ACCOUNT_NBR_LABEL As String = "A2"
Private Const ACCOUNT_NBR_VALUE As String = "B2"
Private Const ACCOUNT_BANK_LABEL As String = "A3"
Private Const ACCOUNT_BANK_VALUE As String = "B3"
Private Const ACCOUNT_STATUS_LABEL As String = "A4"
Private Const ACCOUNT_STATUS_VALUE As String = "B4"
Private Const ACCOUNT_AVAIL_LABEL As String = "A5"
Private Const ACCOUNT_AVAIL_VALUE As String = "B5"
Private Const ACCOUNT_CURRENCY_LABEL As String = "A6"
Private Const ACCOUNT_CURRENCY_VALUE As String = "B6"
Private Const ACCOUNT_TYPE_LABEL As String = "A7"
Private Const ACCOUNT_TYPE_VALUE As String = "B7"
Private Const IN_BUDGET_LABEL As String = "A8"
Private Const IN_BUDGET_VALUE As String = "B8"

Private Const MERGE_SHEET As String = "Comptes Merge"
Private Const ACCOUNT_MERGE_TABLE As String = "AccountsMerge"

Private Const BALANCE_TABLE_NAME As String = "balance"
Private Const OLD_BALANCE_TABLE_NAME As String = "transactions"
Private Const DEPOSIT_TABLE_NAME As String = "deposit"
Private Const INTEREST_TABLE_NAME As String = "interest"
Private Const OLD_INTEREST_TABLE_NAME As String = "yield"

Private Const BTN_HOME_X As Integer = 200
Private Const BTN_HOME_Y As Integer = 10
Private Const BTN_HEIGHT As Integer = 30

Public Sub MergeAccounts(columnKeys As Variant)

    Dim firstAccount As Boolean
    Dim balanceNdx As Integer
    Dim ws As Worksheet

    Call FreezeDisplay

'    For Each colKey In Array(DATE_KEY, ACCOUNT_NAME_KEY, AMOUNT_KEY, DESCRIPTION_KEY, SUBCATEGORY_KEY, IN_BUDGET_KEY)
    For Each colKey In columnKeys
        Dim col As String
        col = GetColName(colKey)
        firstAccount = True
        Dim array1d() As Variant
        For Each ws In Worksheets
           ' Make sure the sheet is not a template or anything else than an account
           If (IsAnAccountSheet(ws)) Then
                balanceNdx = accountBalanceTableIndex(ws.Name)
                If balanceNdx = 0 Then
                    balanceNdx = 1
                End If
                ' Loop on all accounts of the sheet
                If (colKey = ACCOUNT_NAME_KEY) Then
                    arr1d = Create1DArray(ws.ListObjects(balanceNdx).ListRows.Count, ws.Cells(1, 2).value)
                ElseIf (colKey = IN_BUDGET_KEY And Not IsAccountInBudget(ws.Name)) Then
                    arr1d = Create1DArray(ws.ListObjects(balanceNdx).ListRows.Count, 0)
                Else
                    arr1d = GetTableColumn(ws.ListObjects(balanceNdx), col)
                End If
                If (firstAccount) Then
                   totalColumn = arr1d
                   firstAccount = False
                Else
                   ret = ConcatenateArrays(totalColumn, arr1d)
                End If
            End If
        Next ws
        Call SetTableColumn(Sheets(MERGE_SHEET).ListObjects(ACCOUNT_MERGE_TABLE), col, totalColumn)
        Erase totalColumn
    Next colKey

    ' Call SortTable(Sheets(MERGE_SHEET).ListObjects(ACCOUNT_MERGE_TABLE), GetLabel(DATE_KEY), xlAscending, GetLabel(AMOUNT_KEY), xlDescending)
    
End Sub


Public Sub GenBudget()

    Dim i As Long
    Dim newSize As Long
    Dim nbRows As Long
    nbRows = Sheets(MERGE_SHEET).ListObjects(ACCOUNT_MERGE_TABLE).ListRows.Count
    newSize = nbRows

    Dim dateCol() As Variant
    Dim accountCol() As Variant
    Dim amountCol() As Variant
    Dim descCol() As Variant
    Dim categCol() As Variant
    Dim spreadCol() As Variant

    startTime = Now

    With Sheets(MERGE_SHEET)
        dateCol = GetTableColumn(.ListObjects(ACCOUNT_MERGE_TABLE), GetColName(DATE_KEY))
        accountCol = GetTableColumn(.ListObjects(ACCOUNT_MERGE_TABLE), GetColName(ACCOUNT_NAME_KEY))
        amountCol = GetTableColumn(.ListObjects(ACCOUNT_MERGE_TABLE), GetColName(AMOUNT_KEY))
        descCol = GetTableColumn(.ListObjects(ACCOUNT_MERGE_TABLE), GetColName(DESCRIPTION_KEY))
        categCol = GetTableColumn(.ListObjects(ACCOUNT_MERGE_TABLE), GetColName(SUBCATEGORY_KEY))
        spreadCol = GetTableColumn(.ListObjects(ACCOUNT_MERGE_TABLE), GetColName(IN_BUDGET_KEY))
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
        Call ResizeTable(.ListObjects(ACCOUNT_MERGE_TABLE), nbRows + moreRows)
        Call SetTableColumn(.ListObjects(ACCOUNT_MERGE_TABLE), GetColName(DATE_KEY), dateCol)
        Call SetTableColumn(.ListObjects(ACCOUNT_MERGE_TABLE), GetColName(ACCOUNT_NAME_KEY), accountCol)
        Call SetTableColumn(.ListObjects(ACCOUNT_MERGE_TABLE), GetColName(AMOUNT_KEY), amountCol)
        Call SetTableColumn(.ListObjects(ACCOUNT_MERGE_TABLE), GetColName(DESCRIPTION_KEY), descCol)
        Call SetTableColumn(.ListObjects(ACCOUNT_MERGE_TABLE), GetColName(SUBCATEGORY_KEY), categCol)
        Call SetTableColumn(.ListObjects(ACCOUNT_MERGE_TABLE), GetColName(SPREAD_KEY), spreadCol)
        .PivotTables(1).PivotCache.Refresh
    End With

End Sub

Public Sub AccountsQuickRefresh()
    ' startTime = Now
    Call FreezeDisplay
    Call ResizeTable(Sheets(MERGE_SHEET).ListObjects(ACCOUNT_MERGE_TABLE), 1)
    Call MergeAccounts(Array(DATE_KEY, ACCOUNT_NAME_KEY, AMOUNT_KEY, SUBCATEGORY_KEY, IN_BUDGET_KEY))
    Call GenBudget
    Call UnfreezeDisplay
    ' MsgBox ("Quick refresh duration = " & CStr(DateDiff("s", startTime, Now)))
End Sub

Public Sub AccountsFullRefresh()
    ' startTime = Now
    Call FreezeDisplay
    Call ResizeTable(Sheets(MERGE_SHEET).ListObjects(ACCOUNT_MERGE_TABLE), 1)
    Call MergeAccounts(Array(DATE_KEY, ACCOUNT_NAME_KEY, AMOUNT_KEY, DESCRIPTION_KEY, SUBCATEGORY_KEY, IN_BUDGET_KEY))
    Call GenBudget
    Call SortTable(Sheets(MERGE_SHEET).ListObjects(ACCOUNT_MERGE_TABLE), GetLabel(DATE_KEY), xlAscending, GetLabel(AMOUNT_KEY), xlDescending)
    Call UnfreezeDisplay
    ' MsgBox ("Full refresh duration = " & CStr(DateDiff("s", startTime, Now)))
End Sub

Sub CreateAccount()
    accountNbr = InputBox("Account number ?", "Account Number", "<accountNumber>")
    accountName = InputBox("Account name ?", "Account Name", "<accountName>")
    Sheets("Account Template").Visible = True
    Sheets("Account Template").Copy Before:=Sheets(1)
    Sheets("Account Template").Visible = False
    With Sheets(1)
        .Name = accountName
        ' .Range("A1").Formula = "=VLOOKUP("k.account", TblKeys, LangId, FALSE)"
        .Range(ACCOUNT_NAME_VALUE).value = accountName
        formulaRoot = "=VLOOKUP(B$1," & ACCOUNTS_TABLE
        .Range(ACCOUNT_NBR_VALUE).Formula = formulaRoot & ",2,FALSE)"
        .Range(ACCOUNT_BANK_VALUE).Formula = formulaRoot & ",4,FALSE)"
        .Range(ACCOUNT_STATUS_VALUE).Formula = formulaRoot & ",6,FALSE)"
        .Range(ACCOUNT_AVAIL_VALUE).Formula = formulaRoot & ",5,FALSE)"
    End With
End Sub

Public Sub FormatCurrentAccount()
    Call FormatAccount(ActiveSheet.Name)
End Sub
Public Sub FormatAccount(accountId As String)
    Dim ws As Worksheet
    Set ws = getAccountSheet(accountId)
    ws.Cells.RowHeight = 13
    ws.Rows.font.size = 10
    ws.Activate
    Call formatBalanceTable(accountId)
    Call formatDepositTable(accountId)
    Call formatInterestTable(accountId)
    Call formatAccountButtons(ws)
End Sub


Public Sub FormatAllAccountSheets()
'
'  Reformat all account sheets
'
    Dim ws As Worksheet
    Call ShowAllSheets
    For Each ws In Worksheets
       'Call FormatAccountSheet(ws)
       Call FormatAccount(ws.Name)
    Next ws
    Call HideClosedAccounts
    Call hideTemplateAccounts
End Sub

'-------------------------------------------------
Public Function isTemplate(ws As Worksheet) As Boolean
    isTemplate = (ws.Cells(1, 2).value = "TEMPLATE")
End Function

'-------------------------------------------------
Private Sub setClosedAccountsVisibility(visibility As XlSheetVisibility)
    Dim ws As Worksheet
    For Each ws In Worksheets
        If IsClosed(ws.Name) Then
            ws.Visible = visibility
        End If
    Next ws
End Sub

'-------------------------------------------------
Public Sub HideClosedAccounts()
    If GetNamedVariableValue("hideClosedAccounts") = 1 Then
        Call setClosedAccountsVisibility(xlSheetHidden)
    End If
End Sub

'-------------------------------------------------
Public Sub showClosedAccounts()
    Call setClosedAccountsVisibility(xlSheetVisible)
End Sub

'-------------------------------------------------
Private Sub setTemplateAccountsVisibility(visibility As XlSheetVisibility)
    Dim ws As Worksheet
    For Each ws In Worksheets
        If isTemplate(ws) Then
            ws.Visible = visibility
        End If
    Next ws
End Sub
'-------------------------------------------------
Public Sub hideTemplateAccounts()
    Call setTemplateAccountsVisibility(xlSheetHidden)
End Sub
'-------------------------------------------------
Public Sub showTemplateAccounts()
    Call setTemplateAccountsVisibility(xlSheetVisible)
End Sub
Public Sub refreshOpenAccountsList()
    Dim i As Long, nbrAccounts As Long
    Call FreezeDisplay
    Call TruncateTable(Sheets(PARAMS_SHEET).ListObjects(TABLE_OPEN_ACCOUNTS))
    With Sheets(PARAMS_SHEET).ListObjects(TABLE_OPEN_ACCOUNTS)
        For i = 1 To Sheets(ACCOUNTS_SHEET).ListObjects(TABLE_ACCOUNTS).ListRows.Count
            If (Sheets(ACCOUNTS_SHEET).ListObjects(TABLE_ACCOUNTS).ListRows(i).Range.Cells(1, 6).value = "Open") Then
                .ListRows.Add ' Add 1 row at the end, then extend
                .ListRows(.ListRows.Count).Range.Cells(1, 1).value = Sheets(ACCOUNTS_SHEET).ListObjects(TABLE_ACCOUNTS).ListRows(i).Range.Cells(1, 1).value
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

Public Sub SortCurrentAccount()
    Call SortTable(ActiveSheet.ListObjects(1), GetLabel(DATE_KEY), xlAscending, GetLabel(AMOUNT_KEY), xlDescending)
End Sub

'-------------------------------------------------
Public Function accountType(accountId As String) As String
    Dim ws As Worksheet
    Set ws = getAccountSheet(accountId)
    If (accountId = "Account Template") Then
        accountType = ACCOUNT_TYPE_STANDARD
    ElseIf (Not AccountExists(accountId)) Then
        accountType = "ERROR: Not an account"
    Else
        accountType = ws.Range(ACCOUNT_TYPE_VALUE).value
    End If
End Function
Public Function IsInterestAccount(accountId As String) As Boolean
    Dim accType As String
    accType = accountType(accountId)
    IsInterestAccount = Not (accType = "Courant" Or accType = "Autres")
End Function
'-------------------------------------------------
Private Function AccountAttribute(accountId As String, attributeCell As String) As String
    AccountAttribute = ""
    If (AccountExists(accountId)) Then
        Dim ws As Worksheet
        Set ws = getAccountSheet(accountId)
        AccountAttribute = ws.Range(attributeCell).value
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


Public Function AccountTax(accountId) As Double
    AccountTax = CDbl(Application.WorksheetFunction.VLookup(accountId, Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE).DataBodyRange, 10, False))
End Function


Public Function AccountInterestPeriod(accountType) As Integer
    AccountInterestPeriod = CInt(Application.WorksheetFunction.VLookup(accountType, Sheets(PARAMS_SHEET).ListObjects(ACCOUNT_TYPES_TABLE).DataBodyRange, 2, False))
End Function


'-------------------------------------------------
Public Function IsAccountInBudget(accountId As String) As Boolean
    IsAccountInBudget = (AccountExists(accountId) And Sheets(accountId).Range(IN_BUDGET_VALUE).value = "Yes")
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

Private Sub AddInvestmentRow(oTable As ListObject)
    oTable.ListRows.Add
    nbRows = oTable.ListRows.Count
    
    col = GetColumnNumberFromName(oTable, GetLabel(DATE_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).FormulaR1C1 = Date
    
    col = GetColumnNumberFromName(oTable, GetLabel(BALANCE_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).value = oTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).value
    
    col = GetColumnNumberFromName(oTable, GetLabel(SUBCATEGORY_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).value = oTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).value
    
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
    IsAnAccountSheet = (ws.Cells(1, 1).value = GetNamedVariableValue("accountIdentifier") And Not isTemplate(ws))
End Function

Public Function AccountDepositHistory(accountId As String) As Variant
    AccountDepositHistory = accountDepositArray(accountId)
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
        firstEntry = False
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
    Dim interestPeriod As Integer
    deposits = AccountDepositHistory(accountId)
    balances = AccountBalanceHistory(accountId, "Yearly")
    interestPeriod = AccountInterestPeriod(accountType(accountId))
    If interestPeriod > 0 Then
        Call StoreAccountInterests(accountId, InterestsCalc(balances, deposits, accountId, interestPeriod))
    End If
End Sub


Public Sub CalcInterestForAllAccounts()
    Dim accountId As String
    Dim ws As Worksheet
    FreezeDisplay
    For Each ws In Worksheets
        accountId = getAccountId(ws)
        If IsAnAccountSheet(ws) And IsOpen(accountId) And IsInterestAccount(accountId) Then
            Call CalcAccountInterests(accountId)
        End If
    Next ws
    UnfreezeDisplay
End Sub


Private Sub setInterest(r As Range, value As Variant, tax As Double)
    If value = "-" Then
        r.value = value
    Else
        r.value = minval(value, value * (1 - tax))
    End If
End Sub

Public Sub StoreAccountInterests(accountId As String, interestsArray As Variant)
    Dim nbrYears As Long
    Dim lastYear As Variant, last3years As Variant, last5year As Variant, allTime As Variant
    Dim interestsTable As ListObject
    Dim ws As Worksheet
    Dim tax As Double
    nbrYears = UBound(interestsArray)
    Set interestTable = accountInterestTable(accountId)
    thisYear = "-"
    lastYear = "-"
    last3years = "-"
    last5years = "-"
    allTime = "-"
    If interestTable.ListColumns.Count <= 2 Then
        interestTable.ListColumns.Add
        interestTable.ListColumns(3).Name = "Rend. Net"
    End If
    tax = AccountTax(accountId)
    
    For i = 1 To interestTable.ListRows.Count
        With interestTable.ListRows(i)
            interest = "-"
            If i = 1 Then
                interest = interestsArray(nbrYears)
            ElseIf i = 2 And nbrYears >= 3 Then
                interest = interestsArray(nbrYears - 1)
            ElseIf i = 3 And nbrYears >= 5 Then
                interest = ArrayAverage(interestsArray, nbrYears - 3, nbrYears - 1)
            ElseIf i = 4 And nbrYears >= 7 Then
                interest = ArrayAverage(interestsArray, nbrYears - 5, nbrYears - 1)
            ElseIf i = 5 And nbrYears >= 3 Then
                interest = ArrayAverage(interestsArray, 2, nbrYears - 1)
            End If
            .Range(1, 2).value = interest
            If interest = "-" Then
                .Range(1, 3) = interest
            Else
                .Range(1, 3).value = min(interest, interest * (1 - tax))
            End If
        End With
    Next i
    interestTable.ListColumns(2).DataBodyRange.NumberFormat = "0.0%"
    interestTable.ListColumns(3).DataBodyRange.NumberFormat = "0.0%"
    
    With interestTable.ListColumns(3).DataBodyRange
        Call InterestsStore(accountId, .Rows(1).value, .Rows(2).value, .Rows(3).value, .Rows(4).value, .Rows(5).value)
    End With
End Sub


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

Public Sub BtnAccountFormat()
    Call FormatAccount(getAccountId(ActiveSheet))
End Sub

'--------------------------------------------------------------------------
' Private methods
'--------------------------------------------------------------------------

'----------------------------------------------------------------------------
' Table as Tables
'----------------------------------------------------------------------------
Private Function accountTable(accountId As String, accountSection As String) As ListObject
    Dim ws As Worksheet
    Dim i As Long
    Set ws = getAccountSheet(accountId)
    For i = 1 To ws.ListObjects.Count
        If LCase$(ws.ListObjects(i).Name) Like accountSection & "*" Then
            Set accountTable = ws.ListObjects(i)
            Exit For
        End If
    Next i
End Function

Private Function accountDepositTable(accountId As String) As ListObject
    Set accountDepositTable = accountTable(accountId, "deposit")
End Function

Private Function accountBalanceTable(accountId As String) As ListObject
    Set accountBalanceTable = accountTable(accountId, "balance")
End Function

Private Function accountInterestTable(accountId As String) As ListObject
    Set accountInterestTable = accountTable(accountId, "interest")
    If accountInterestTable Is Nothing Then
        Set accountInterestTable = accountTable(accountId, "yield")
    End If
End Function

'----------------------------------------------------------------------------
' Table as Indexes
'----------------------------------------------------------------------------
Private Function accountTableIndex(accountId As String, accountSection As String) As Integer
    Dim ws As Worksheet
    Dim i As Long
    Set ws = getAccountSheet(accountId)
    accountTableIndex = 0
    For i = 1 To ws.ListObjects.Count
        If LCase$(ws.ListObjects(i).Name) Like accountSection & "*" Then
            accountTableIndex = i
            Exit For
        End If
    Next i
End Function

Private Function accountBalanceTableIndex(accountId As String) As Long
    accountBalanceTableIndex = accountTableIndex(accountId, BALANCE_TABLE_NAME)
    ' TODO: Remove old table names
    If accountBalanceTableIndex = 0 Then
        accountBalanceTableIndex = accountTableIndex(accountId, OLD_BALANCE_TABLE_NAME)
    End If
End Function

Private Function accountDepositTableIndex(accountId As String) As Long
    accountDepositTableIndex = accountTableIndex(accountId, DEPOSIT_TABLE_NAME)
End Function

Private Function accountInterestTableIndex(accountId As String) As Long
    accountInterestTableIndex = accountTableIndex(accountId, INTEREST_TABLE_NAME)
    ' TODO: Remove old table names
    If accountInterestTableIndex = 0 Then
        accountInterestTableIndex = accountTableIndex(accountId, INTEREST_TABLE_NAME)
    End If
End Function


'----------------------------------------------------------------------------
' Table as Arrays
'----------------------------------------------------------------------------
Private Function accountArray(accountId As String, accountSection As String) As Variant
    Dim i As Long
    Dim ws As Worksheet
    Set accountArray = Nothing
    Set ws = getAccountSheet(accountId)
    For i = 1 To ws.ListObjects.Count
        If LCase$(ws.ListObjects(i).Name) Like accountSection & "*" Then
            accountArray = GetTableAsArray(ws.ListObjects(i))
            Exit For
        End If
    Next i
End Function

Private Function accountDepositArray(accountId As String) As Variant
    accountDepositArray = accountArray(accountId, DEPOSIT_TABLE_NAME)
End Function

Private Function AccountBalanceArray(accountId As String) As Variant
    AccountBalanceArray = accountArray(accountId, BALANCE_TABLE_NAME)
End Function

Private Function getAccountId(ws As Worksheet) As String
    getAccountId = ws.Name
End Function

Private Function getAccountSheet(accountId As String) As Worksheet
    Set getAccountSheet = ThisWorkbook.Sheets(accountId)
End Function


'----------------------------------------------------------------------------
' Private formatting functions
'----------------------------------------------------------------------------

Private Sub setBtnAttributes(oBtn As Shape, Optional font As String = vbNullString, Optional text As String = vbNullString, _
                             Optional fontStyle As String = vbNullString, Optional size As Integer = 0)
    oBtn.Select
    If text <> vbNullString Then
        Selection.Characters.text = text
    End If
    If font <> vbNullString Then
        Selection.Characters.font.Name = font
    End If
    If size <> 0 Then
        Selection.Characters.font.size = size
    End If
    If fontStyle <> vbNullString Then
        Selection.Characters.font.fontStyle = style
    End If
    'With Selection.Characters().font
    '    .Name = font
    '    .fontStyle = "Normal"
    '    .size = 18
        '.Strikethrough = False
        '.Superscript = False
        '.Subscript = False
        '.OutlineFont = False
        '.Shadow = False
        '.Underline = xlUnderlineStyleNone
        '.ColorIndex = xlAutomatic
        '.TintAndShade = 0
        '.ThemeFont = xlThemeFontNone
    'End With
End Sub

Private Sub setBtnTextAndFont(oBtn, text As String, font As String)
    oBtn.Select
    Selection.Characters.text = text

End Sub

Private Sub formatAccountButtons(ws As Worksheet)
    If ws.Shapes.Count <= 0 Then
        Exit Sub
    End If
    Dim i As Long
    i = 0
    Dim s As Shape
    Dim sbw As Integer, lbw As Integer
    sbw = 40
    lbw = 100

        ' For Each btnName In Array("BtnPrev", "BtnNext5", "BtnBottom", "BtnHome", "BtnPrev5", "BtnTop", "BtnNext")

    For Each s In ws.Shapes
        If s.Name = "BtnHome" Then
            Call ShapePlacementXY(s, BTN_HOME_X, BTN_HOME_Y, BTN_HOME_X + sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call setBtnAttributes(s, text:="9", font:="Webdings", size:=18)
        ElseIf s.Name = "BtnPrev5" Then
            Call ShapePlacementXY(s, BTN_HOME_X + sbw, BTN_HOME_Y, BTN_HOME_X + 2 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call setBtnAttributes(s, text:="7", font:="Webdings", size:=18)
        ElseIf s.Name = "BtnPrev" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 2 * sbw, BTN_HOME_Y, BTN_HOME_X + 3 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call setBtnAttributes(s, text:="3", font:="Webdings", size:=18)
        ElseIf s.Name = "BtnNext" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 3 * sbw, BTN_HOME_Y, BTN_HOME_X + 4 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call setBtnAttributes(s, text:="4", font:="Webdings", size:=18)
        ElseIf s.Name = "BtnNext5" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 4 * sbw, BTN_HOME_Y, BTN_HOME_X + 5 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call setBtnAttributes(s, text:="8", font:="Webdings", size:=18)
        ElseIf s.Name = "BtnTop" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 5 * sbw, BTN_HOME_Y, BTN_HOME_X + 6 * sbw, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call setBtnAttributes(s, text:="5", font:="Webdings", size:=18)
        ElseIf s.Name = "BtnBottom" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 6 * sbw, BTN_HOME_Y, BTN_HOME_X + 7 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call setBtnAttributes(s, text:="6", font:="Webdings", size:=18)
        ElseIf s.Name = "BtnSort" Then
            Call ShapePlacementXY(s, BTN_HOME_X, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call setBtnAttributes(s, text:="~", font:="Webdings", size:=18)
        ElseIf s.Name = "BtnImport" Then
            Call ShapePlacementXY(s, BTN_HOME_X + sbw, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 2 * sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call setBtnAttributes(s, text:=Chr$(71), font:="Webdings", size:=18)
        ElseIf s.Name = "BtnAddEntry" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 2 * sbw, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 3 * sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call setBtnAttributes(s, text:="+1", font:="Arial", size:=14)
        ElseIf s.Name = "BtnInterests" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 3 * sbw, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 4 * sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call setBtnAttributes(s, text:=Chr(143), font:="Webdings", size:=18)
        ElseIf s.Name = "BtnFormat" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 4 * sbw, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 6 * sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call setBtnAttributes(s, text:="Format", font:="Arial", size:=12)

        ElseIf (s.Type = msoFormControl) Then
            ' This is a button, move it to right place
            row = i Mod 4
            col = i \ 4
            Call ShapePlacementXY(s, 400 + col * 100, 10 + row * BTN_HEIGHT, 400 + col * 100, 25 + row * BTN_HEIGHT - 1)
            i = i + 1
        End If
    Next s
End Sub

Private Sub formatBalanceTable(accountId As String)
    Dim oTable As ListObject
    Dim ws As Worksheet
    Set ws = getAccountSheet(accountId)
    Set oTable = accountBalanceTable(accountId)
    If IsEmpty(oTable) Then
        Exit Sub
    End If
    Call SetTableStyle(oTable, "TableStyleMedium2")
    Dim col As Long
    col = GetColumnNumberFromName(oTable, GetLabel(DATE_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
        Call SetTableColumnFormat(oTable, col, DATE_FORMAT)
    End If
    col = GetColumnNumberFromName(oTable, GetLabel(AMOUNT_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
        Call SetTableColumnFormat(oTable, col, EUR_FORMAT)
    End If
    col = GetColumnNumberFromName(oTable, "Montant CHF")
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 17, ws)
        Call SetTableColumnFormat(oTable, col, CHF_FORMAT)
    End If
    col = GetColumnNumberFromName(oTable, "Montant USD")
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
        Call SetTableColumnFormat(oTable, col, USD_FORMAT)
    End If
    col = GetColumnNumberFromName(oTable, GetLabel(BALANCE_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 18, ws)
        Call SetTableColumnFormat(oTable, col, EUR_FORMAT)
    End If
    col = GetColumnNumberFromName(oTable, "Solde CHF")
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 18, ws)
        Call SetTableColumnFormat(oTable, col, CHF_FORMAT)
    End If
    col = GetColumnNumberFromName(oTable, "Solde USD")
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 18, ws)
        Call SetTableColumnFormat(oTable, col, USD_FORMAT)
    End If
    col = GetColumnNumberFromName(oTable, GetLabel(DESCRIPTION_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 70, ws)
    End If
    col = GetColumnNumberFromName(oTable, GetLabel(SUBCATEGORY_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
    End If
    col = GetColumnNumberFromName(oTable, GetLabel(CATEGORY_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
    End If
    col = GetColumnNumberFromName(oTable, GetLabel(IN_BUDGET_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 5, ws)
        Call SetColumnWidth(Chr$(col + 65), 5, ws)
    End If
End Sub

Private Sub formatDepositTable(accountId As String)
    Dim oTable As ListObject
    Set oTable = accountDepositTable(accountId)
    Call SetTableStyle(oTable, "TableStyleMedium4")
    Call SetTableColumnFormat(oTable, 1, DATE_FORMAT)
    Call SetTableColumnFormat(oTable, 2, EUR_FORMAT)
End Sub

Private Sub formatInterestTable(accountId As String)
    Dim oTable As ListObject
    Set oTable = accountInterestTable(accountId)
    Call SetTableStyle(oTable, "TableStyleMedium5")
    Call SetTableColumnFormat(oTable, 2, "0.00%")
End Sub

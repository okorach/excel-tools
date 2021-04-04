Attribute VB_Name = "AccountMgr"

Public Const CHF_FORMAT = "#,###,##0.00"" CHF "";-#,###,##0.00"" CHF "";0.00"" CHF """
Public Const EUR_FORMAT = "#,###,##0.00"" � "";-#,###,##0.00"" � "";0.00"" � """
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

Public Const ACCOUNTS_SHEET As String = "Comptes"
Public Const BALANCE_SHEET As String = "Solde"

Public Const ACCOUNT_TYPE_STANDARD As String = "Standard"
Public Const ACCOUNT_TYPE_SAVINGS As String = "Savings"
Public Const ACCOUNT_TYPE_INVESTMENT As String = "Investment"

Public Const ACCOUNT_CLOSED As Long = 0
Public Const ACCOUNT_OPEN As Long = 1

Private Const MAX_MERGE_SIZE As Long = 100000

Private Const ACCOUNTS_TABLE As String = "tblAccounts"
Private Const ACCOUNT_TYPES_TABLE As String = "TblAccountTypes"

Private Const ACCOUNT_NAME_VALUE As String = "B1"
Private Const ACCOUNT_NBR_VALUE As String = "B2"
Private Const ACCOUNT_BANK_VALUE As String = "B3"
Private Const ACCOUNT_STATUS_VALUE As String = "B4"
Private Const ACCOUNT_AVAIL_VALUE As String = "B5"
Private Const ACCOUNT_CURRENCY_VALUE As String = "B6"

Private Const MERGE_SHEET As String = "Comptes Merge"
Private Const ACCOUNT_MERGE_TABLE As String = "AccountsMerge"

Public Const BALANCE_TABLE_NAME As String = "balance"
Public Const DEPOSIT_TABLE_NAME As String = "deposit"
Public Const INTEREST_TABLE_NAME As String = "interest"

Public Const BTN_HOME_X As Integer = 200
Public Const BTN_HOME_Y As Integer = 10
Public Const BTN_HEIGHT As Integer = 30

Public Sub MergeAccounts(columnKeys As Variant)

    Dim firstAccount As Boolean
    Dim balanceNdx As Integer
    Dim ws As Worksheet

    Call FreezeDisplay

    For Each colKey In columnKeys
        Dim col As String
        col = GetColName(colKey)
        firstAccount = True
        Dim array1d() As Variant
        For Each ws In Worksheets
           ' Make sure the sheet is not a template or anything else than an account
           If (IsAnAccount(ws)) Then
                balanceNdx = accountBalanceTableIndex(ws.name)
                If balanceNdx = 0 Then
                    balanceNdx = 1
                End If
                ' Loop on all accounts of the sheet
                If (colKey = ACCOUNT_NAME_KEY) Then
                    arr1d = Create1DArray(ws.ListObjects(balanceNdx).ListRows.Count, ws.name)
                ElseIf (colKey = IN_BUDGET_KEY And Not AccountIsInBudget(ws.name)) Then
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

Sub AccountCreate()
    Dim accNbr As String, accName As String
    accNbr = InputBox("Account number ?", "Account Number", "<accountNumber>")
    accName = InputBox("Account name ?", "Account Name", "<accountName>")
    Sheets("Account Template").Visible = True
    Sheets("Account Template").Copy Before:=Sheets(1)
    Sheets("Account Template").Visible = False
    With Sheets(1)
        .name = accName
        ' .Range("A1").Formula = "=VLOOKUP("k.account", TblKeys, LangId, FALSE)"
        .Range(ACCOUNT_NAME_VALUE).value = accName
        formulaRoot = "=VLOOKUP(B$1," & ACCOUNTS_TABLE
        .Range(ACCOUNT_NBR_VALUE).Formula = formulaRoot & ",2,FALSE)"
        .Range(ACCOUNT_BANK_VALUE).Formula = formulaRoot & ",4,FALSE)"
        .Range(ACCOUNT_STATUS_VALUE).Formula = formulaRoot & ",6,FALSE)"
        .Range(ACCOUNT_AVAIL_VALUE).Formula = formulaRoot & ",5,FALSE)"
    End With
End Sub

Public Sub AccountFormatCurrent()
    Call AccountFormat(ActiveSheet.name)
End Sub
Public Sub AccountFormat(accountId As String)
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


Public Sub AccountFormatAllSheets()
'
'  Reformat all account sheets
'
    Dim ws As Worksheet
    Call ShowAllSheets
    For Each ws In Worksheets
       If IsAnAccount(ws) Then
           Call AccountFormat(ws.name)
        End If
    Next ws
    Call AccountHideClosed
    Call AccountHideTemplates
End Sub

'-------------------------------------------------
Public Function isTemplate(ws As Worksheet) As Boolean
    isTemplate = (ws.name Like "*TEMPLATE*")
End Function

'-------------------------------------------------
Private Sub accountSetClosedVisibility(visibility As XlSheetVisibility)
    Dim ws As Worksheet
    For Each ws In Worksheets
        If AccountExists(ws.name) And AccountIsClosed(ws.name) Then
            ws.Visible = visibility
        End If
    Next ws
End Sub
Public Sub AccountHideClosed()
    If GetGlobalParam("hideClosedAccounts") = 1 Then
        Call accountSetClosedVisibility(xlSheetHidden)
    End If
End Sub

Public Sub AccountShowClosed()
    Call accountSetClosedVisibility(xlSheetVisible)
End Sub
'-------------------------------------------------
Private Sub accountSetTemplatesVisibility(visibility As XlSheetVisibility)
    Dim ws As Worksheet
    For Each ws In Worksheets
        If isTemplate(ws) Then
            ws.Visible = visibility
        End If
    Next ws
End Sub

Public Sub AccountHideTemplates()
    Call accountSetTemplatesVisibility(xlSheetHidden)
End Sub

Public Sub AccountShowTemplates()
    Call accountSetTemplatesVisibility(xlSheetVisible)
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
Public Function IsInterestAccount(accountId As String) As Boolean
    Dim accType As String
    accType = AccountType(accountId)
    IsInterestAccount = Not (accType = "Courant" Or accType = "Autres")
End Function
Public Function AccountInterestPeriod(AccountType) As Integer
    AccountInterestPeriod = CInt(KeyedTableValue(Sheets(PARAMS_SHEET).ListObjects(ACCOUNT_TYPES_TABLE), AccountType, 2))
End Function


'-------------------------------------------------
Public Function AccountExists(accountId As String) As Boolean
    AccountExists = (SheetExists(accountId) And KeyedTableValue(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE), accountId, 2) <> vbNull)
End Function
Public Function AccountNumber(accountId As String) As String
    AccountNumber = CStr(KeyedTableValue(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE), accountId, 2))
End Function
Public Function AccountName(accountId As String) As String
    AccountName = CStr(KeyedTableValue(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE), accountId, 3))
End Function
Public Function AccountBank(accountId As String) As String
    AccountBank = CStr(KeyedTableValue(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE), accountId, 4))
End Function
Public Function AccountAvailability(accountId As String) As String
    AccountAvailability = CStr(KeyedTableValue(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE), accountId, 5))
End Function
Public Function AccountStatus(accountId As String) As String
    AccountStatus = CStr(KeyedTableValue(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE), accountId, 6))
End Function
Public Function AccountIsOpen(accountId As String) As Boolean
    AccountIsOpen = (AccountStatus(accountId) = "Open")
End Function
Public Function AccountIsClosed(accountId As String) As Boolean
    AccountIsClosed = Not AccountIsOpen(accountId)
End Function
Public Function AccountCurrency(accountId As String) As String
    AccountCurrency = CStr(KeyedTableValue(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE), accountId, 7))
End Function
Public Function AccountType(accountId As String) As String
    AccountType = CStr(KeyedTableValue(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE), accountId, 8))
End Function
Public Function AccountIsInBudget(accountId As String) As Boolean
    AccountInBudget = (KeyedTableValue(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE), accountId, 9) = 1)
End Function
Public Function AccountTaxRate(accountId As String) As Double
    AccountTaxRate = CDbl(KeyedTableValue(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE), accountId, 10))
End Function

Public Sub AddSavingsRow()
    Call AddInvestmentRow(ActiveSheet.ListObjects(1))
End Sub

Private Sub AddInvestmentRow(oTable As ListObject)
    oTable.ListRows.Add
    nbRows = oTable.ListRows.Count
    
    col = TableColNbrFromName(oTable, GetLabel(DATE_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).FormulaR1C1 = Date
    
    col = TableColNbrFromName(oTable, GetLabel(BALANCE_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).value = oTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).value
    
    col = TableColNbrFromName(oTable, GetLabel(SUBCATEGORY_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).value = oTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).value
    
    col = TableColNbrFromName(oTable, GetLabel(AMOUNT_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).FormulaR1C1 = oTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).FormulaR1C1
    
    col = TableColNbrFromName(oTable, GetLabel(DESCRIPTION_KEY))
    oTable.ListColumns(col).DataBodyRange.Rows(nbRows).FormulaR1C1 = oTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).FormulaR1C1
End Sub


'-------------------------------------------------
Public Function IsAnAccount(accountIdOrWs As Variant) As Boolean
    IsAnAccount = True
    If VarType(accountIdOrWs) = vbString Then
        accountId = accountIdOrWs
    Else
        accountId = accountIdOrWs.name
    End If
    IsAnAccount = (KeyedTableValue(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE), accountId, 2) <> vbNull)
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
    interestPeriod = AccountInterestPeriod(AccountType(accountId))
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
        If IsAnAccount(ws) And AccountIsOpen(accountId) And IsInterestAccount(accountId) Then
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
    If interestTable.ListColumns.Count <= 2 Then
        interestTable.ListColumns.Add
        interestTable.ListColumns(3).name = "Rend. Net"
    End If
    tax = AccountTaxRate(accountId)
    
    For i = 1 To interestTable.ListRows.Count
        With interestTable.ListRows(i)
            Dim interest As Variant
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
    interestTable.ListColumns(2).DataBodyRange.NumberFormat = INTEREST_FORMAT
    interestTable.ListColumns(3).DataBodyRange.NumberFormat = INTEREST_FORMAT
    
    With interestTable.ListColumns(3).DataBodyRange
        Call InterestsStore(accountId, .Rows(1).value, .Rows(2).value, .Rows(3).value, .Rows(4).value, .Rows(5).value)
    End With
End Sub


Public Function getSelectedAccount() As String
    selectedNbr = GetNamedVariableValue("selectedAccount")
    getSelectedAccount = Sheets(PARAMS_SHEET).ListObjects("TblOpenAccounts").ListRows(selectedNbr).Range(1, 1)
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
        If LCase$(ws.ListObjects(i).name) Like "*_" & accountSection Then
            Set accountTable = ws.ListObjects(i)
            Exit For
        End If
    Next i
End Function

Private Function accountDepositTable(accountId As String) As ListObject
    Set accountDepositTable = accountTable(accountId, DEPOSIT_TABLE_NAME)
End Function

Private Function accountBalanceTable(accountId As String) As ListObject
    Set accountBalanceTable = accountTable(accountId, BALANCE_TABLE_NAME)
End Function

Private Function accountInterestTable(accountId As String) As ListObject
    Set accountInterestTable = accountTable(accountId, INTEREST_TABLE_NAME)
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
        If LCase$(ws.ListObjects(i).name) Like "*_" & accountSection Then
            accountTableIndex = i
            Exit For
        End If
    Next i
End Function

Private Function accountBalanceTableIndex(accountId As String) As Long
    accountBalanceTableIndex = accountTableIndex(accountId, BALANCE_TABLE_NAME)
End Function

Private Function accountDepositTableIndex(accountId As String) As Long
    accountDepositTableIndex = accountTableIndex(accountId, DEPOSIT_TABLE_NAME)
End Function

Private Function accountInterestTableIndex(accountId As String) As Long
    accountInterestTableIndex = accountTableIndex(accountId, INTEREST_TABLE_NAME)
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
        If LCase$(ws.ListObjects(i).name) Like "*_" & accountSection Then
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
    getAccountId = ws.name
End Function

Private Function getAccountSheet(accountId As String) As Worksheet
    Set getAccountSheet = ThisWorkbook.Sheets(accountId)
End Function


'----------------------------------------------------------------------------
' Private formatting functions
'----------------------------------------------------------------------------


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
        If s.name = "BtnHome" Then
            Call ShapePlacementXY(s, BTN_HOME_X, BTN_HOME_Y, BTN_HOME_X + sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="9", font:="Webdings", size:=18)
        ElseIf s.name = "BtnPrev5" Then
            Call ShapePlacementXY(s, BTN_HOME_X + sbw, BTN_HOME_Y, BTN_HOME_X + 2 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="7", font:="Webdings", size:=18)
        ElseIf s.name = "BtnPrev" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 2 * sbw, BTN_HOME_Y, BTN_HOME_X + 3 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="3", font:="Webdings", size:=18)
        ElseIf s.name = "BtnNext" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 3 * sbw, BTN_HOME_Y, BTN_HOME_X + 4 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="4", font:="Webdings", size:=18)
        ElseIf s.name = "BtnNext5" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 4 * sbw, BTN_HOME_Y, BTN_HOME_X + 5 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="8", font:="Webdings", size:=18)
        ElseIf s.name = "BtnTop" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 5 * sbw, BTN_HOME_Y, BTN_HOME_X + 6 * sbw, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="5", font:="Webdings", size:=18)
        ElseIf s.name = "BtnBottom" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 6 * sbw, BTN_HOME_Y, BTN_HOME_X + 7 * sbw - 1, BTN_HOME_Y + BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="6", font:="Webdings", size:=18)
        ElseIf s.name = "BtnSort" Then
            Call ShapePlacementXY(s, BTN_HOME_X, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="~", font:="Webdings", size:=18)
        ElseIf s.name = "BtnImport" Then
            Call ShapePlacementXY(s, BTN_HOME_X + sbw, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 2 * sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:=Chr$(71), font:="Webdings", size:=18)
        ElseIf s.name = "BtnAddEntry" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 2 * sbw, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 3 * sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="+1", font:="Arial", size:=14)
        ElseIf s.name = "BtnInterests" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 3 * sbw, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 4 * sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:=Chr(143), font:="Webdings", size:=18)
        ElseIf s.name = "BtnFormat" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 4 * sbw, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 6 * sbw - 1, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
            Call SetBtnAttributes(s, text:="Format", font:="Arial", size:=12)

        ElseIf (s.Type = msoFormControl) Then
            ' This is a button, move it to right place
            row = i Mod 4
            col = i \ 4
            Call ShapePlacementXY(s, 400 + col * 100, 10 + row * BTN_HEIGHT, 400 + col * 100, 25 + row * BTN_HEIGHT - 1)
            i = i + 1
        End If
    Next s
    ws.Range("A1").Select
End Sub

Private Sub formatBalanceTable(accountId As String)
    Dim oTable As ListObject
    Dim ws As Worksheet
    Set ws = getAccountSheet(accountId)
    Set oTable = accountBalanceTable(accountId)
    If oTable Is Nothing Then
        Exit Sub
    End If
    oTable.name = accountId & "_" & BALANCE_TABLE_NAME
    Call SetTableStyle(oTable, "TableStyleMedium2")
    Dim col As Long
    col = TableColNbrFromName(oTable, GetLabel(DATE_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
        Call SetTableColumnFormat(oTable, col, DATE_FORMAT)
    End If
    col = TableColNbrFromName(oTable, GetLabel(AMOUNT_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
        Call SetTableColumnFormat(oTable, col, EUR_FORMAT)
    End If
    col = TableColNbrFromName(oTable, "Montant CHF")
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 17, ws)
        Call SetTableColumnFormat(oTable, col, CHF_FORMAT)
    End If
    col = TableColNbrFromName(oTable, "Montant USD")
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
        Call SetTableColumnFormat(oTable, col, USD_FORMAT)
    End If
    col = TableColNbrFromName(oTable, GetLabel(BALANCE_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 18, ws)
        Call SetTableColumnFormat(oTable, col, EUR_FORMAT)
    End If
    col = TableColNbrFromName(oTable, "Solde CHF")
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 18, ws)
        Call SetTableColumnFormat(oTable, col, CHF_FORMAT)
    End If
    col = TableColNbrFromName(oTable, "Solde USD")
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 18, ws)
        Call SetTableColumnFormat(oTable, col, USD_FORMAT)
    End If
    col = TableColNbrFromName(oTable, GetLabel(DESCRIPTION_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 70, ws)
    End If
    col = TableColNbrFromName(oTable, GetLabel(SUBCATEGORY_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
    End If
    col = TableColNbrFromName(oTable, GetLabel(CATEGORY_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
    End If
    col = TableColNbrFromName(oTable, GetLabel(IN_BUDGET_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 5, ws)
        Call SetColumnWidth(Chr$(col + 65), 5, ws)
    End If
End Sub

Private Sub formatDepositTable(accountId As String)
    Dim oTable As ListObject
    Set oTable = accountDepositTable(accountId)
    If oTable Is Nothing Then
        Exit Sub
    End If
    oTable.name = accountId & "_" & DEPOSIT_TABLE_NAME
    Call SetTableStyle(oTable, "TableStyleMedium4")
    Call SetTableColumnFormat(oTable, 1, DATE_FORMAT)
    Call SetTableColumnFormat(oTable, 2, EUR_FORMAT)
End Sub

Private Sub formatInterestTable(accountId As String)
    Dim oTable As ListObject
    Set oTable = accountInterestTable(accountId)
    If oTable Is Nothing Then
        Exit Sub
    End If
    oTable.name = accountId & "_" & INTEREST_TABLE_NAME
    Call SetTableStyle(oTable, "TableStyleMedium5")
    Call SetTableColumnFormat(oTable, 2, INTEREST_FORMAT)
    Call SetTableColumnFormat(oTable, 3, INTEREST_FORMAT)
End Sub


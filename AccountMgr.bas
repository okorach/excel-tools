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

Public Const ACCOUNTS_SHEET As String = "Comptes"
Public Const BALANCE_SHEET As String = "Solde"

Public Const ACCOUNT_CLOSED As Long = 0
Public Const ACCOUNT_OPEN As Long = 1

Private Const MAX_MERGE_SIZE As Long = 100000

Private Const ACCOUNTS_TABLE As String = "tblAccounts"
Public Const ACCOUNT_TYPES_TABLE As String = "TblAccountTypes"

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


Private Const BTN_HOME_TEXT As String = "9"
Private Const BTN_PREV_5_TEXT As String = "7"
Private Const BTN_PREV_TEXT As String = "3"
Private Const BTN_NEXT_TEXT As String = "4"
Private Const BTN_NEXT_5_TEXT As String = "8"
Private Const BTN_BOTTOM_TEXT As String = "6"
Private Const BTN_TOP_TEXT As String = "5"
Private Const BTN_SORT_TEXT As String = "~"
Private Const BTN_IMPORT_TEXT As String = "G"
Private Const BTN_ADD_ROW_TEXT As String = "+1"
Private Const BTN_FORMAT_TEXT As String = "Format"



Public Sub MergeAccounts(columnKeys As Variant)

    Dim firstAccount As Boolean
    Dim balanceNdx As Integer
    Dim ws As Worksheet

    'Call FreezeDisplay
    Dim modal As ProgressBar
    Set modal = NewProgressBar("Refresh in progress", (UBound(columnKeys) + 1) * Worksheets.Count)
    
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
            modal.Update
        Next ws
        Call SetTableColumn(Sheets(MERGE_SHEET).ListObjects(ACCOUNT_MERGE_TABLE), col, totalColumn)
        Erase totalColumn
    Next colKey
    Set modal = Nothing

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

Public Sub AccountCreateBtn()
    CreateAccountUserForm.show
End Sub
Public Function AccountCreate(accountId As String, accCurrency As String, accType As String, _
    Optional avail As Integer = 0, Optional accNumber As String = vbNullString, _
    Optional bank As String = vbNullString, Optional inBudget As Boolean = True, Optional taxRate As Double = 0) As Boolean

    Dim res As Boolean
    Dim accTable As ListObject
    Set accTable = Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE)
    res = KeyedTableInsert(accTable, accountId, accountId, 2)
    If Not res Then
        MsgBox ("Account already exist, aborting")
        AccountCreate = False
        Exit Function
    End If
    AccountCreate = True
    
    Call FreezeDisplay
    
    Call KeyedTableUpdate(accTable, accountId, accNumber, 2)
    res = KeyedTableInsert(accTable, accountId, accountId, 3)
    Call KeyedTableUpdate(accTable, accountId, bank, 4)
    Call KeyedTableUpdate(accTable, accountId, avail, 5)
    Call KeyedTableUpdate(accTable, accountId, GetLabel("k.AccountOpen"), 6)
    Call KeyedTableUpdate(accTable, accountId, accCurrency, 7)
    Call KeyedTableUpdate(accTable, accountId, accType, 8)
    Call KeyedTableUpdate(accTable, accountId, inBudget, 9)
    Call KeyedTableUpdate(accTable, accountId, taxRate, 10)
    'tax = KeyedTableValue(Sheets(PARAMS_SHEET).ListObjects(ACCOUNT_TYPES_TABLE), accountId, 3)
    'Call KeyedTableUpdate(accTable, accountId, CDbl(tax), 10)

    Sheets.Add
    ActiveSheet.name = accountId
    Call accountAddBalanceTable(ActiveSheet, accountId, accCurrency, accType)
    Call accountAddStandardButtons(ActiveSheet)
    If accType = GetLabel("k.accountStandard") Then
        Call accountAddImportButton(ActiveSheet)
    Else
        Call accountAddDepositTable(ActiveSheet, accountId)
        Call accountAddInterestTable(ActiveSheet, accountId)
        Call accountAddInterestButtons(ActiveSheet)
    End If
    Call AccountFormat(accountId)
    Call UnfreezeDisplay
End Function

Private Sub accountAddStandardButtons(ws As Worksheet)
    For Each btnName In Array("BtnHome", "BtnPrev5", "BtnPrev", "BtnNext", "BtnNext5", "BtnTop", "BtnBottom", "BtnFormat", "BtnSort")
        If Not ShapeExist(ws, CStr(btnName)) Then
            Call BtnAdd(ws, name:=CStr(btnName), action:="ThisWorkbook.GoToSolde")
        End If
    Next btnName
End Sub
Private Sub accountAddInterestButtons(ws As Worksheet)
    For Each btnName In Array("BtnInterest", "BtnAddEntry")
        If Not ShapeExist(ws, CStr(btnName)) Then
            Call BtnAdd(ws, name:=CStr(btnName), action:="ThisWorkbook.GoToSolde")
        End If
    Next btnName
End Sub
Private Sub accountAddImportButton(ws As Worksheet)
    For Each btnName In Array("BtnImport")
        If Not ShapeExist(ws, CStr(btnName)) Then
            Call BtnAdd(ws, name:=CStr(btnName), action:="ThisWorkbook.GoToSolde")
        End If
    Next btnName
End Sub
Private Sub accountAddBalanceTable(ws As Worksheet, accountId As String, accCurrency As String, accType As String)
    Dim tblName As String
    tblName = accountId & "_" & BALANCE_TABLE_NAME
    
    ws.Range("A11:A12").Select
    ws.ListObjects.Add(xlSrcRange, Range("$A$10:$E$11"), , xlYes).name = tblName

    With ws.ListObjects(tblName)
        .TableStyle = "TableStyleMedium2"
        Dim lblAmount As String, lblBalance As String, lblDate As String, lblSubcat As String
        lblAmount = GetLabel("k.amount")
        lblBalance = GetLabel("k.accountBalance")
        lblDate = GetLabel("k.date")
        lblSubcat = GetLabel("k.subcategory")
        .ListColumns(1).name = lblDate
        .ListColumns(2).name = lblAmount
        .ListColumns(3).name = lblBalance
        .ListColumns(4).name = GetLabel("k.description")
        .ListColumns(5).name = lblSubcat
        If accType = GetLabel("k.accountStandard") Then
            .ListRows(1).Range(1, 3).FormulaR1C1 = "=[" & lblAmount & "]+IF(ISNUMBER(R[-1]C),R[-1]C,0)"
            .ListColumns.Add
            .ListColumns(6).name = GetLabel("k.category")
            .ListRows(1).Range(1, 6).FormulaR1C1 = "=VLOOKUP([" & lblSubcat & "],TableCategories,2,FALSE)"
            If AccountIsInBudget(accountId) Then
                .ListColumns.Add
                .ListColumns(7).name = GetLabel("k.inBudget")
            End If
        Else
            .ListRows(1).Range(1, 2).FormulaR1C1 = "=[" & lblBalance & "]-IF(ISNUMBER(R[-1]C[1]),R[-1]C[1],0)"
        End If
        If accCurrency <> GetGlobalParam("DefaultCurrency") Then
            lblAmount = lblAmount & " " & accCurrency
            lblBalance = lblBalance & " " & accCurrency
            .ListColumns.Add(4).name = lblAmount
            .ListColumns.Add(5).name = lblBalance
            If accType = GetLabel("k.accountStandard") Then
                .ListRows(1).Range(1, 5).FormulaR1C1 = "=[" & lblAmount & "]+IF(ISNUMBER(R[-1]C),R[-1]C,0)"
            Else
                .ListRows(1).Range(1, 4).FormulaR1C1 = "=[" & lblBalance & "]-IF(ISNUMBER(R[-1]C[1]),R[-1]C[1],0)"
            End If
            .ListRows(1).Range(1, 2).FormulaR1C1 = "=[" & lblAmount & "]/VLOOKUP([" & lblDate & "],CHFtoEUR,2,TRUE)"
            .ListRows(1).Range(1, 3).FormulaR1C1 = "=[" & lblBalance & "]/VLOOKUP([" & lblDate & "],CHFtoEUR,2,TRUE)"
        End If
    End With
End Sub

Private Sub accountAddDepositTable(ws As Worksheet, accountId As String, Optional accCurrency As String = "EUR")
    Dim tblName As String
    tblName = accountId & "_" & DEPOSIT_TABLE_NAME
    
    Dim r As String
    r = "$G$10:$H$11"
    If AccountCurrency(accountId) <> GetGlobalParam("DefaultCurrency") Then
        r = "$I$10:$J$11"
    End If
    ws.ListObjects.Add(xlSrcRange, Range(r), , xlYes).name = tblName
    With ws.ListObjects(tblName)
        .TableStyle = "TableStyleMedium4"
        .ListColumns(1).name = GetLabel("k.date")
        .ListColumns(2).name = GetLabel("k.amount")
    End With
End Sub
Private Sub accountAddInterestTable(ws As Worksheet, accountId As String)
    Dim tblName As String
    tblName = accountId & "_" & INTEREST_TABLE_NAME
    
    Dim r As String
    r = "$G$1:$I$6"
    If AccountCurrency(accountId) <> GetGlobalParam("DefaultCurrency") Then
        r = "$I$1:$K$6"
    End If
    ws.ListObjects.Add(xlSrcRange, Range(r), , xlYes).name = tblName
    With ws.ListObjects(tblName)
        .TableStyle = "TableStyleMedium5"
        .ListColumns(1).name = GetLabel("k.period")
        .ListColumns(2).name = GetLabel("k.grossYield")
        .ListColumns(3).name = GetLabel("k.netYield")
        .ListRows(1).Range(1, 1).value = GetLabel("k.thisYear")
        .ListRows(2).Range(1, 1).value = GetLabel("k.lastYear")
        .ListRows(3).Range(1, 1).value = GetLabel("k.last3Years")
        .ListRows(4).Range(1, 1).value = GetLabel("k.last5Years")
        .ListRows(5).Range(1, 1).value = GetLabel("k.allTime")
    End With
End Sub

Public Sub AccountFormatHere()
    Call FreezeDisplay
    Call AccountFormat(getAccountId(ActiveSheet))
    Call UnfreezeDisplay
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


Public Sub AccountFormatAll()
'
'  Reformat all account sheets
'
    Dim modal As ProgressBar
    Set modal = NewProgressBar("Formatting in progress", Worksheets.Count + 2)
    Call FreezeDisplay
    
    Dim ws As Worksheet, activeWs As Worksheet
    Set activeWs = ActiveSheet
    Call ShowAllSheets
    modal.Update
    
    For Each ws In Worksheets
       If IsAnAccount(ws) Then
           Call AccountFormat(getAccountId(ws))
        End If
        modal.Update
    Next ws
    Call AccountHideClosed
    modal.Update
    activeWs.Activate
    Set modal = Nothing
    Call UnfreezeDisplay
End Sub
Public Sub GoToSolde()
    Sheets(BALANCE_SHEET).Activate
End Sub

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

Public Sub refreshOpenAccountsList()
    Dim i As Long, nbrAccounts As Long
    Call FreezeDisplay
    Call TruncateTable(Sheets(PARAMS_SHEET).ListObjects(TABLE_OPEN_ACCOUNTS))
    With Sheets(PARAMS_SHEET).ListObjects(TABLE_OPEN_ACCOUNTS)
        For i = 1 To Sheets(ACCOUNTS_SHEET).ListObjects(TABLE_ACCOUNTS).ListRows.Count
            If Sheets(ACCOUNTS_SHEET).ListObjects(TABLE_ACCOUNTS).ListRows(i).Range.Cells(1, 6).value = GetLabel("k.accountOpen") Then
                .ListRows.Add ' Add 1 row at the end, then extend
                .ListRows(.ListRows.Count).Range.Cells(1, 1).value = _
                    Sheets(ACCOUNTS_SHEET).ListObjects(TABLE_ACCOUNTS).ListRows(i).Range.Cells(1, 1).value
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
    IsInterestAccount = Not (accType = GetLabel("k.accountStandard") Or accType = "Autres")
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
    AccountIsOpen = (AccountStatus(accountId) = GetLabel("k.accountOpen"))
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


Private Sub setInterest(r As Range, value As Variant, tax As Double)
    If value = "-" Then
        r.value = value
    Else
        r.value = minval(value, value * (1 - tax))
    End If
End Sub



Public Function getSelectedAccount() As String
    selectedNbr = GetNamedVariableValue("selectedAccount")
    getSelectedAccount = Sheets(PARAMS_SHEET).ListObjects("TblOpenAccounts").ListRows(selectedNbr).Range(1, 1)
End Function


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

Public Function accountDepositTable(accountId As String) As ListObject
    Set accountDepositTable = accountTable(accountId, DEPOSIT_TABLE_NAME)
End Function

Public Function accountBalanceTable(accountId As String) As ListObject
    Set accountBalanceTable = accountTable(accountId, BALANCE_TABLE_NAME)
End Function

Public Function accountInterestTable(accountId As String) As ListObject
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

Public Function getAccountId(ws As Worksheet) As String
    getAccountId = ws.name
End Function

Private Function getAccountSheet(accountId As String) As Worksheet
    Set getAccountSheet = ThisWorkbook.Sheets(accountId)
End Function


'----------------------------------------------------------------------------
' Private formatting functions
'----------------------------------------------------------------------------

Private Sub formatAccountButtons(ws As Worksheet)
    If ws.Shapes.Count <= 0 Then
        Exit Sub
    End If
    Dim sbw As Integer
    sbw = 40
    Dim i As Long
    i = 0
    Dim s As Shape

    For Each btnData In Array( _
        "BtnHome," & BTN_HOME_TEXT & ",GoToSolde,Webdings,18,1,1,40" _
        , "BtnPrev5," & BTN_PREV_5_TEXT & ",GoBack5,Webdings,18,1,2,40" _
        , "BtnPrev," & BTN_PREV_TEXT & ",GoToPrev,Webdings,18,1,3,40" _
        , "BtnNext," & BTN_NEXT_TEXT & ",GoToNext,Webdings,18,1,4,40" _
        , "BtnNext5," & BTN_NEXT_5_TEXT & ",GoFwd5,Webdings,18,1,5,40" _
        , "BtnTop," & BTN_TOP_TEXT & ",scrollToTop,Webdings,18,1,6,40" _
        , "BtnBottom," & BTN_BOTTOM_TEXT & ",scrollToBottom,Webdings,18,1,7,40" _
        , "BtnSort," & BTN_SORT_TEXT & ",SortCurrentAccount,Webdings,18,2,1,40" _
        , "BtnImport," & BTN_IMPORT_TEXT & ",ImportAny,Webdings,18,2,2,40" _
        , "BtnAddEntry," & BTN_ADD_ROW_TEXT & ",AddInvestmentRow,Arial,14,2,3,40" _
        , "BtnInterests," & Chr$(143) & ",InterestsCalcHere,Webdings,18,2,4,40" _
        , "BtnFormat," & BTN_FORMAT_TEXT & ",AccountFormatHere,Arial,18,2,5,80" _
        )
        values = Split(btnData, ",", -1, vbTextCompare)
        Set s = ShapeFind(ws, CStr(values(0)))
        If Not s Is Nothing Then
            Call BtnSetProperties(s, text:=CStr(values(1)), action:=CStr(values(2)), font:=CStr(values(3)), fontSize:=CInt(values(4)))
            Call ShapePlacement(s, BTN_HOME_X + (CInt(values(6)) - 1) * sbw, _
                BTN_HOME_Y + (CInt(values(5)) - 1) * BTN_HEIGHT, CInt(values(7)) - 1, BTN_HEIGHT - 1)
        End If
    Next btnData
    ws.Range("A1").Select
    
'    ElseIf (s.Type = msoFormControl) Then
'            ' Other buttons, move to row 3
'            r = 3 + (i Mod 6)
'            c = i \ 4
'            w = 2 * sbw
'            i = i + 1
'        End If
'        Call ShapePlacement(s, BTN_HOME_X + (c - 1) * sbw, BTN_HOME_Y + (r - 1) * BTN_HEIGHT, sbw - 1, BTN_HEIGHT - 1)
End Sub

Private Sub formatBalanceTable(accountId As String)
    Dim oTable As ListObject
    Dim ws As Worksheet
    Set ws = getAccountSheet(accountId)
    Set oTable = accountBalanceTable(accountId)
    If oTable Is Nothing Then
        Exit Sub
    End If
    oTable.name = Replace(Replace(Replace(LCase$(accountId), " ", "_"), "é", "e"), "è", "e") & "_" & BALANCE_TABLE_NAME
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
    For Each colName In Array(GetLabel(AMOUNT_KEY), GetLabel(BALANCE_KEY))
        If TableColumnNameExists(oTable, CStr(colName)) Then
            Call TableColumnFormatIcons(oTable, CStr(colName))
        End If
        For Each curr In Array("EUR", "CHF", "USD")
            Dim cname As String
            cname = CStr(colName) & " " & CStr(curr)
            If TableColumnNameExists(oTable, cname) Then
                Call TableColumnFormatIcons(oTable, cname)
            End If
        Next curr
    Next colName
End Sub

Private Sub formatDepositTable(accountId As String)
    Dim oTable As ListObject
    Set oTable = accountDepositTable(accountId)
    If oTable Is Nothing Then
        Exit Sub
    End If
    oTable.name = Replace(Replace(Replace(LCase$(accountId), " ", "_"), "é", "e"), "è", "e") & "_" & DEPOSIT_TABLE_NAME
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
    oTable.name = Replace(Replace(Replace(LCase$(accountId), " ", "_"), "é", "e"), "è", "e") & "_" & INTEREST_TABLE_NAME
    Call SetTableStyle(oTable, "TableStyleMedium5")
    Call SetTableColumnFormat(oTable, 2, INTEREST_FORMAT)
    Call SetTableColumnFormat(oTable, 3, INTEREST_FORMAT)
End Sub


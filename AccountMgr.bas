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

Private Const BALANCE_TABLE_NAME As String = "balance"
Private Const DEPOSIT_TABLE_NAME As String = "deposit"
Private Const INTEREST_TABLE_NAME As String = "interest"

Private Const BTN_HOME_X As Integer = 200
Private Const BTN_HOME_Y As Integer = 10
Private Const BTN_HEIGHT As Integer = 22

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


Public Sub MergeAccounts2()

    Dim firstAccount As Boolean
    Dim ws As Worksheet

    Call FreezeDisplay

    Call TruncateTable(Sheets(MERGE_SHEET).ListObjects("AccountsMerge"))
    For Each ws In Worksheets
        If (IsAnAccountSheet(ws)) Then
            Call MergeTables(Sheets(MERGE_SHEET).ListObjects("AccountsMerge"), ws.ListObjects(1))
        End If
    Next ws
    Call SortTable(Sheets(MERGE_SHEET).ListObjects("AccountsMerge"), GetLabel(DATE_KEY), xlAscending, GetLabel(AMOUNT_KEY), xlDescending)

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

Public Sub FormatAccount(accountId As String)
    Dim ws As Worksheet
    Set ws = getAccountSheet(accountId)
    ws.Cells.RowHeight = 13
    ws.Rows.Font.size = 10
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
       Call FormatAccount(ws.name)
    Next ws
    Call HideClosedAccounts
    Call hideTemplateAccounts
End Sub

'-------------------------------------------------
Public Function isTemplate(ws As Worksheet) As Boolean
    isTemplate = (ws.Cells(1, 2).Value = "TEMPLATE")
End Function

'-------------------------------------------------
Private Sub setClosedAccountsVisibility(visibility As XlSheetVisibility)
    Dim ws As Worksheet
    For Each ws In Worksheets
        If IsClosed(ws.name) Then
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

Public Sub SortCurrentAccount()
    Call SortTable(ActiveSheet.ListObjects(1), GetLabel(DATE_KEY), xlAscending, GetLabel(AMOUNT_KEY), xlDescending)
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
        AccountAttribute = ws.Range(attributeCell).Value
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

Private Sub AddInvestmentRow(oTable As ListObject)
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


Public Sub StoreAccountInterests(accountId As String, interestsArray As Variant)
    Dim nbrYears As Long
    Dim interestsTable As ListObject
    Dim ws As Worksheet
    nbrYears = UBound(interestsArray)
    interestTable = AccountInterestsTable(accountId)
    With interestsTable.ListColumns(2)
        .DataBodyRange.Rows(1).Value = interestsArray(nbrYears)
        For k = 2 To 5
            .DataBodyRange.Rows(k).Value = "-"
        Next k
        If nbrYears >= 2 Then
            .DataBodyRange.Rows(2).Value = interestsArray(nbrYears - 1)
        End If
        If nbrYears >= 4 Then
            .DataBodyRange.Rows(3).Value = ArrayAverage(interestsArray, nbrYears - 3, nbrYears - 1)
        End If
        If nbrYears >= 6 Then
            .DataBodyRange.Rows(4).Value = ArrayAverage(interestsArray, nbrYears - 5, nbrYears - 1)
        End If
        If nbrYears >= 2 Then
            .DataBodyRange.Rows(5).Value = ArrayAverage(interestsArray, 1, nbrYears - 1)
        End If
        .DataBodyRange.NumberFormat = "0.00%"
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
        If LCase$(ws.ListObjects(i).name) Like accountSection & "*" Then
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
        If LCase$(ws.ListObjects(i).name) Like accountSection & "*" Then
            accountTableIndex = i
            Exit For
        End If
    Next i
End Function

Private Function AccountBalanceTableIndex(accountId As String) As Long
    AccountBalanceTableIndex = accountTableIndex(accountId, BALANCE_TABLE_NAME)
End Function

Private Function AccountDepositTableIndex(accountId As String) As Long
    AccountDepositTableIndex = accountTableIndex(accountId, DEPOSIT_TABLE_NAME)
End Function

Private Function AccountInterestTableIndex(accountId As String) As Long
    AccountInterestTableIndex = accountTableIndex(accountId, INTEREST_TABLE_NAME)
    If AccountInterestTableIndex = 0 Then
        AccountInterestTableIndex = accountTableIndex(accountId, "yield")
    End If
End Function


'----------------------------------------------------------------------------
' Table as Arrays
'----------------------------------------------------------------------------
Private Function accountArray(accountId As String, accountSection As String) As Variant
    Dim i As Long
    Dim ws As Worksheet
    Set accountArray = Empty
    Set ws = getAccountSheet(accountId)
    For i = 1 To ws.ListObjects.Count
        If LCase$(ws.ListObjects(i).name) Like accountSection & "*" Then
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
    
Private Sub formatAccountButtons(ws As Worksheet)
    If ws.Shapes.Count <= 0 Then
        Exit Sub
    End If
    Dim i As Long
    i = 0
    Dim s As Shape
    For Each s In ws.Shapes
        If s.name = "BtnPrev5" Then
            Call ShapePlacementXY(s, BTN_HOME_X, BTN_HOME_Y, BTN_HOME_X + 29, BTN_HOME_Y + BTN_HEIGHT - 1)
        ElseIf s.name = "BtnPrev" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 30, BTN_HOME_Y, BTN_HOME_X + 59, BTN_HOME_Y + BTN_HEIGHT - 1)
        ElseIf s.name = "BtnHome" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 60, BTN_HOME_Y, BTN_HOME_X + 129, BTN_HOME_Y + BTN_HEIGHT - 1)
        ElseIf s.name = "BtnNext" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 130, BTN_HOME_Y, BTN_HOME_X + 159, BTN_HOME_Y + BTN_HEIGHT - 1)
        ElseIf s.name = "BtnNext5" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 160, BTN_HOME_Y, BTN_HOME_X + 189, BTN_HOME_Y + BTN_HEIGHT - 1)
        ElseIf s.name = "BtnTop" Then
            Call ShapePlacementXY(s, BTN_HOME_X, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 99, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
        ElseIf s.name = "BtnBottom" Then
            Call ShapePlacementXY(s, BTN_HOME_X, BTN_HOME_Y + 2 * BTN_HEIGHT, BTN_HOME_X + 99, BTN_HOME_Y + 3 * BTN_HEIGHT - 1)
        ElseIf s.name = "BtnSort" Then
            Call ShapePlacementXY(s, BTN_HOME_X, BTN_HOME_Y + 3 * BTN_HEIGHT, BTN_HOME_X + 99, BTN_HOME_Y + 4 * BTN_HEIGHT - 1)
        ElseIf s.name = "BtnInterests" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 100, BTN_HOME_Y + 3 * BTN_HEIGHT, BTN_HOME_X + 199, BTN_HOME_Y + 4 * BTN_HEIGHT - 1)
        ElseIf s.name = "BtnImport" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 100, BTN_HOME_Y + BTN_HEIGHT, BTN_HOME_X + 199, BTN_HOME_Y + 2 * BTN_HEIGHT - 1)
        ElseIf s.name = "BtnAddEntry" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 100, BTN_HOME_Y + 2 * BTN_HEIGHT, BTN_HOME_X + 199, BTN_HOME_Y + 3 * BTN_HEIGHT - 1)
        ElseIf s.name = "BtnFormat" Then
            Call ShapePlacementXY(s, BTN_HOME_X + 200, BTN_HOME_Y, BTN_HOME_X + 299, BTN_HOME_Y + BTN_HEIGHT - 1)

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
        Call SetTableColumnFormat(oTable, col, "m/d/yyyy")
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
    Call SetTableColumnFormat(oTable, 1, "m/d/yyyy")
    Call SetTableColumnFormat(oTable, 2, EUR_FORMAT)
End Sub

Private Sub formatInterestTable(accountId As String)
    Dim oTable As ListObject
    Set oTable = accountInterestTable(accountId)
    Call SetTableStyle(oTable, "TableStyleMedium5")
    Call SetTableColumnFormat(oTable, 2, "0.00%")
End Sub

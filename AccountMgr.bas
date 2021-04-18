Attribute VB_Name = "AccountMgr"
Public Const DATE_KEY As String = "k.date"
Public Const ACCOUNT_NAME_KEY As String = "k.accountName"
Public Const AMOUNT_KEY As String = "k.amount"
Public Const BALANCE_KEY As String = "k.accountBalance"
Public Const DESCRIPTION_KEY As String = "k.description"
Public Const SUBCATEGORY_KEY As String = "k.subcategory"
Public Const CATEGORY_KEY As String = "k.category"
Public Const IN_BUDGET_KEY As String = "k.inBudget"
Public Const SPREAD_KEY As String = "k.amountSpread"

Public Sub MergeAccounts(columnKeys As Variant)

    Dim firstAccount As Boolean
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
            Dim oAccount As Account
            Set oAccount = LoadAccount(getAccountId(ws))
            If Not (oAccount Is Nothing) Then
                Dim tbl As ListObject
                Set tbl = oAccount.BalanceTable()
                ' Loop on all accounts of the sheet
                If (colKey = ACCOUNT_NAME_KEY) Then
                    arr1d = Create1DArray(tbl.ListRows.Count, oAccount.name)
                ElseIf (colKey = IN_BUDGET_KEY And Not oAccount.IsInBudget()) Then
                    arr1d = Create1DArray(tbl.ListRows.Count, 0)
                Else
                    arr1d = GetTableColumn(tbl, col)
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

Public Sub AccountCreate()
    CreateAccountUserForm.show
End Sub


Public Function LoadAccount(accountId As String) As Account
    Set LoadAccount = New Account
    If Not LoadAccount.Load(accountId) Then
        Set LoadAccount = Nothing
    End If
End Function

Public Function NewAccount(aId As String, aNbr As String, aBank As String, Optional aCur As String = vbNullString, _
                           Optional aType As String = vbNullString, Optional aAvail As Integer = 0, _
                           Optional aInB As Boolean = False, Optional aTax As Double = 0) As Account
    Set NewAccount = New Account
    If Not NewAccount.Create(aId, aNbr, aBank, aCur, aType, aAvail, aInB, aTax) Then
        Set NewAccount = Nothing
    End If
End Function


Public Sub AccountFormatHere()
    Call FreezeDisplay
    Dim oAccount As Account
    Set oAccount = LoadAccount(getAccountId(ActiveSheet))
    oAccount.FormatMe
    Call UnfreezeDisplay
End Sub

Public Sub AccountSortHere()
    Call FreezeDisplay
    Dim oAccount As Account
    Set oAccount = LoadAccount(getAccountId(ActiveSheet))
    oAccount.Sort
    Call UnfreezeDisplay
End Sub


Public Sub AccountFormatAll()
    Dim modal As ProgressBar
    Set modal = NewProgressBar("Formatting in progress", Worksheets.Count + 2)
    Call FreezeDisplay
    
    Dim ws As Worksheet, activeWs As Worksheet
    Set activeWs = ActiveSheet
    Call ShowAllSheets
    modal.Update
    
    For Each ws In Worksheets
        If IsAnAccount(ws) Then
           Dim oAccount As Account
           Set oAccount = LoadAccount(getAccountId(ws))
           oAccount.FormatMe
        End If
        modal.Update
    Next ws
    Call AccountHideClosed
    modal.Update
    activeWs.Activate
    Set modal = Nothing
    Call UnfreezeDisplay
End Sub

Public Sub AccountHideClosed()
    If GetGlobalParam("hideClosedAccounts") = 1 Then
        Call accountSetClosedVisibility(xlSheetHidden)
    End If
End Sub

Public Sub AccountShowClosed()
    Call accountSetClosedVisibility(xlSheetVisible)
End Sub

Public Sub GoToSolde()
    Sheets(BALANCE_SHEET).Activate
End Sub

'-------------------------------------------------
Private Sub accountSetClosedVisibility(visibility As XlSheetVisibility)
    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim oAccount As Account
        Set oAccount = LoadAccount(getAccountId(ws))
        If Not (oAccount Is Nothing) Then
            If oAccount.IsClosed() Then
                ws.Visible = visibility
            End If
        End If
    Next ws
End Sub

Public Function getAccountId(ws As Worksheet) As String
    getAccountId = ws.name
End Function

Public Sub AccountRefreshOpenList()
    Call FreezeDisplay
    Dim modal As ProgressBar
    Set modal = NewProgressBar("Refresh open accounts list", Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE).ListRows.Count + 3)
    'startTime = Now
    Call TruncateTable(Sheets(PARAMS_SHEET).ListObjects(OPEN_ACCOUNTS_TABLE))
    'MsgBox ("Truncate duration = " & CStr(DateDiff("s", startTime, Now)))
    modal.Update
    With Sheets(PARAMS_SHEET).ListObjects(OPEN_ACCOUNTS_TABLE)
        For Each row In Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE).ListRows
            'startTime = Now
            Dim oAccount As Account
            Set oAccount = LoadAccount(row.Range(1, ACCOUNT_KEY_COL).value)
            If oAccount.IsOpen Then
                .ListRows.Add ' Add 1 row at the end, then extend
                .ListRows(.ListRows.Count).Range.Cells(1, 1).value = oAccount.Id
            End If
            modal.Update
            'MsgBox ("Add account " & oAccount.Id & " duration = " & CStr(DateDiff("s", startTime, Now)))
        Next row
    End With

    Dim refCell As String
    refCell = Names("selectedAccount").RefersTo
    refCell = Right$(refCell, LenB(refCell) - 1)
    For Each wsName In Array(BALANCE_SHEET, BALANCE_PER_ACCOUNT_SHEET)
        Sheets(wsName).Activate
        Sheets(wsName).Shapes(ACCOUNT_SELECTOR).Select
        With Selection
            .ListFillRange = PARAMS_SHEET & "!$K$2:$K$" & CStr(Sheets(PARAMS_SHEET).ListObjects(OPEN_ACCOUNTS_TABLE).ListRows.Count + 1)
            .LinkedCell = refCell
            .DropDownLines = 15
            .Display3DShading = True
        End With
        Sheets(wsName).Range("A1").Activate
        modal.Update
    Next wsName
    Set modal = Nothing
    Call UnfreezeDisplay
End Sub


Public Sub AddSavingsRow()
    Dim oAccount As Account
    Set oAccount = LoadAccount(getAccountId(ActiveSheet))
    oAccount.AddBalanceRow
End Sub


'-------------------------------------------------
Public Function IsAnAccount(accountIdOrWs As Variant) As Boolean
    IsAnAccount = True
    If VarType(accountIdOrWs) = vbString Then
        accountId = accountIdOrWs
    Else
        accountId = accountIdOrWs.name
    End If
    Dim accounts As KeyedTable
    Set accounts = NewKeyedTable(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE))
    IsAnAccount = accounts.KeyExists(accountId)
End Function


Public Function getSelectedAccount() As String
    selectedNbr = GetNamedVariableValue("selectedAccount")
    getSelectedAccount = Sheets(PARAMS_SHEET).ListObjects(OPEN_ACCOUNTS_TABLE).ListRows(selectedNbr).Range(1, 1)
End Function


'--------------------------------------------------------------------------
' Private methods
'--------------------------------------------------------------------------

Private Function getAccountSheet(accountId As String) As Worksheet
    Set getAccountSheet = ThisWorkbook.Sheets(accountId)
End Function


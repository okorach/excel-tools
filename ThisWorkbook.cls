VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub auto_open()
Application.Calculation = xlAutomatic
End Sub
Private Sub Workbook_Open()
Application.Calculation = xlAutomatic
End Sub

Sub mergeAccounts()

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

Public Sub GoToAccount()
    selectedNbr = Range("H72").Value
    Dim accountName As String
    accountName = Sheets(PARAMS_SHEET).Range("L" & CStr(selectedNbr + 1))
    If accountExists(accountName) Then
        Sheets(accountName).Activate
    End If
End Sub
Public Sub GoToSolde()
    Sheets(BALANCE_SHEET).Activate
End Sub

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

Public Function GetLabel(key)
    lang = getNamedVariableValue("Language")
    If lang = "English" Then
        col = 3
    Else
        col = 2
    End If

    GetLabel = Application.VLookup(key, Sheets("Language").ListObjects("TblKeys").DataBodyRange, col, False)
    If IsError(GetLabel) Then
        GetLabel = key & " not found"
    End If
End Function

Public Function GetColName(key)
    GetColName = GetLabel(key)
End Function

Public Sub ErrorInfo(key1 As String, Optional key2 As String = "")
    msg = GetLabel(key1)
    If key2 <> "" Then
        ms = msg & ", " & GetLabel(key2)
    End If
    MsgBox (msg)
End Sub

Sub mergeAccounts()

    Call freezeDisplay

    For Each colKey In Array("k.date", "k.accountName", "k.amount", "k.description", "k.subcategory", "k.inBudget")
        col = GetColName(colKey)
        firstAccount = True
        Dim array1d() As Variant
        For Each ws In Worksheets
           ' Make sure the sheet is not a template or anything else than an account
           If (isAnAccountSheet(ws)) Then
                'MsgBox (ws + " " + colNbr)
                ' Loop on all accounts of the sheet
                If (colKey = "k.accountName") Then
                    arr1d = Create1DArray(ws.ListObjects(1).ListRows.Count, ws.Cells(1, 2).Value)
                ElseIf (colKey = "k.inBudget" And Not isAccountInBudget(ws.name)) Then
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
        Call setTableColumn(Sheets("Comptes Merge").ListObjects(1), col, totalColumn, False)
        Erase totalColumn
    Next colKey
    ActiveWorkbook.Worksheets("Comptes Merge").ListObjects("AccountsMerge").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Comptes Merge").ListObjects("AccountsMerge").Sort. _
        SortFields.Add key:=Range("AccountsMerge[[#Headers],[#Data],[Date]]"), SortOn _
        :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Comptes Merge").ListObjects("AccountsMerge"). _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets("Comptes Merge").PivotTables(1).PivotCache.Refresh

End Sub

Public Sub genBudget()

    Dim newSize As Integer
    'Sheets("Comptes Merge").ListObjects(1).Range.AutoFilter.ShowAllData
    nbRows = Sheets("Comptes Merge").ListObjects(1).ListRows.Count
    newSize = nbRows
    nbCols = Sheets("Comptes Merge").ListObjects(1).ListColumns.Count

    Dim oTable As Variant
    Dim dateCol() As Variant
    Dim accountCol() As Variant
    Dim amountCol() As Variant
    Dim descCol() As Variant
    Dim categCol() As Variant
    Dim spreadCol() As Variant

    With Sheets("Comptes Merge")
        dateCol = getTableColumn(.ListObjects(1), GetColName("k.date"), False)
        accountCol = getTableColumn(.ListObjects(1), GetColName("k.accountName"), False)
        amountCol = getTableColumn(.ListObjects(1), GetColName("k.amount"), False)
        descCol = getTableColumn(.ListObjects(1), GetColName("k.description"), False)
        categCol = getTableColumn(.ListObjects(1), GetColName("k.subcategory"), False)
        spreadCol = getTableColumn(.ListObjects(1), GetColName("k.inBudget"), False)
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

    With Sheets("Comptes Merge")
        Call resizeTable(.ListObjects(1), nbRows + moreRows)
        Call setTableColumn(.ListObjects(1), GetColName("k.date"), dateCol, False)
        Call setTableColumn(.ListObjects(1), GetColName("k.accountName"), accountCol, False)
        Call setTableColumn(.ListObjects(1), GetColName("k.amount"), amountCol, False)
        Call setTableColumn(.ListObjects(1), GetColName("k.description"), descCol, False)
        Call setTableColumn(.ListObjects(1), GetColName("k.subcategory"), categCol, False)
        Call setTableColumn(.ListObjects(1), GetColName("k.amountSpread"), spreadCol, False)
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
    accountName = Sheets("Paramètres").Range("L" & CStr(selectedNbr + 1))
    If accountExists(accountName) Then
        Sheets(accountName).Activate
    End If
End Sub
Public Sub GoToSolde()
    Sheets("Solde").Activate
End Sub



Attribute VB_Name = "TablesMgr"
'==============================================================================
'
'  Table Tools
'
'==============================================================================

' resizeTable(oTable, targetSize as long)
' truncateTable(oTable)

' fillTableColumn(oTable, col As Variant, ByRef tValues)
' Function getTableAsArray(oTable, Optional colList As Variant = 0) As Variant
' setTableArray(oTable, ByRef tValues)
' Function getTableColumn(oTable, colNbr, Optional twoD As Boolean = True) As Variant
' setTableColumn(oTable, arr, colNbr, Optional twoD As Boolean = True)
' Function copyTable(oSrcTable, oTgtTable, Optional colList As Variant = 0) As Boolean
' Function appendTable(oSrcTable, oTgtTable, Optional colList As Variant = 0) As Boolean
' setTableColumnFormula(oTable, formula, colNbr)
' fillTableFormula(oTable, col As Variant, formula As String)
' clearTableColumn(oTable, colNbr)
' clearTableRow(oTable, rowNbr)
' clearTable(oTable)
' createArray(arraySize, elementValue) as Variant
' create2DArray(arraySize, elementValue) as Variant

'------------------------------------------------------------------------------
' Resize (extend or shrink) an Excel table object to a desired size
'------------------------------------------------------------------------------
Public Sub ResizeTable(oTable, targetSize As Long)
    ' TODO: Test TargetSize >0, oTable exists
    Dim i As Long
    Dim nbRows As Long
    nbRows = oTable.ListRows.Count
    If nbRows < targetSize Then
        oTable.ListRows.Add ' Add 1 row at the end, then extend
        If targetSize - nbRows > 1 Then
            oTable.ListRows(nbRows + 1).Range.Resize(targetSize - nbRows - 1).Insert shift:=xlDown
        End If
    ElseIf nbRows > targetSize Then
        oTable.ListRows(targetSize + 1).Range.Resize(nbRows - targetSize).Delete shift:=xlUp
        'For i = nbRows To targetSize + 1 Step -1
        '    oTable.ListRows(i).Delete
        'Next i
    End If

End Sub

Public Sub TruncateTable(oTable)
    Call ResizeTable(oTable, 0)
End Sub

'------------------------------------------------------------------------------
Public Sub ClearTable(oTable As ListObject)
    Dim tableSize As Long
    tableSize = oTable.ListRows.Count
    Call TruncateTable(oTable)
    Call ResizeTable(oTable, tableSize)
End Sub

Public Sub SortTable(oTable As ListObject, sortCol1 As String, Optional sortOrder1 As XlSortOrder = xlAscending, _
                     Optional sortCol2 As String = "", Optional sortOrder2 As XlSortOrder = xlDescending)
    oTable.Sort.SortFields.Clear
    ' Sort table by date first, then by amount
    oTable.Sort.SortFields.Add key:=Range(oTable.name & "[" & sortCol1 & "]"), SortOn:=xlSortOnValues, Order:=sortOrder1, _
        DataOption:=xlSortNormal
    If sortCol2 <> "" Then
        oTable.Sort.SortFields.Add key:=Range(oTable.name & "[" & sortCol2 & "]"), SortOn:=xlSortOnValues, Order:=sortOrder2, _
             DataOption:=xlSortNormal
    End If
    With oTable.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


'------------------------------------------------------------------------------
' Get the values of a table object cells in a 2 dimensions array
'------------------------------------------------------------------------------
Public Function GetTableAsArray(oTable As ListObject, Optional colList As Variant = 0) As Variant
    Dim nbrRows As Long
    Dim nbrCols As Long

    nbrCols = oTable.ListColumns.Count
    nbrRows = oTable.ListRows.Count
    Dim cList() As Variant
    cList = GetColList(oTable, colList)
    Dim arr() As Variant
    ReDim arr(1 To nbrRows, 1 To nbrCols)
    i = 0
    For Each C In cList
        i = i + 1
        For j = 1 To nbrRows
            arr(j, i) = oTable.ListColumns(C).DataBodyRange.Rows(j).Value
        Next j
    Next C
    GetTableAsArray = arr
End Function

'------------------------------------------------------------------------------
' Set the values of a 2D array into a table object
'------------------------------------------------------------------------------
Public Sub SetTableFromArray(oTable As ListObject, tValues As Variant)
    Call TruncateTable(oTable)
    nbrCols = oTable.ListColumns.Count
    nbrRows = UBound(tValues, 1)
    For i = 1 To nbrRows
        oTable.ListRows.Add
        For j = 1 To nbrCols
            oTable.ListColumns(j).DataBodyRange.Rows(i).Value = tValues(i, j)
        Next j
    Next i
End Sub

'------------------------------------------------------------------------------
' Copies a table in another, possibly only specific columns,
' assuming both have the same structure
' Returns true in case of success, false in case of any error
'------------------------------------------------------------------------------

Public Function CopyTable(oSrcTable As ListObject, oTgtTable As ListObject, Optional colList As Variant = 0) As Boolean
    Call ResizeTable(oTgtTable, oSrcTable.ListRows.Count)
    If (IsNumeric(colList)) Then
        For i = 1 To oSrcTable.ListColumns.Count
           oTgtTable.ListColumns(i).DataBodyRange.Value = oSrcTable.ListColumns(i).DataBodyRange.Value
        Next i
    Else
        For Each col In colList
           oTgtTable.ListColumns(col).DataBodyRange.Value = oSrcTable.ListColumns(col).DataBodyRange.Value
        Next col
    End If
    CopyTable = True
End Function

'------------------------------------------------------------------------------
' Appends oSrcTable at end of oTgtTable
'------------------------------------------------------------------------------

Public Sub AppendTableToTable(oSrcTable As ListObject, oTgtTable As ListObject, Optional colList As Variant = 0)
    Dim offset As Long
    offset = oTgtTable.ListRows.Count
    Call ResizeTable(oTgtTable, offset + oSrcTable.ListRows.Count)
    cList = GetColList(oSrcTable, colList)
    For Each col In cList
        For j = 1 To oSrcTable.ListColumns(col).DataBodyRange.Rows.Count
            oTgtTable.ListColumns(col).DataBodyRange.Rows(j + offset).Value = oSrcTable.ListColumns(col).DataBodyRange.Rows(j).Value
        Next j
    Next col
End Sub


Public Sub appendTableToTableFast(oSrcTable As ListObject, oTgtTable As ListObject, Optional colList As Variant = 0)
    sizeOffset = oTgtTable.ListRows.Count
    Call ResizeTable(oTgtTable, sizeOffset + oSrcTable.ListRows.Count)
    cList = GetColList(oSrcTable, colList)
    For Each col In cList
        Dim srcArr() As Variant
        Dim tgtArr() As Variant
        tgtArr = GetTableColumn(oTgtTable, col)
        srcArr = GetTableColumn(oSrcTable, col)
        sizeOffset = UBound(tgtArr)
        ReDim Arr1(UBound(srcArr) + sizeOffset)
        For i = 1 To UBound(srcArr)
            tgtArr(sizeOffset + i) = srcArr(i)
        Next i
    Next col
End Sub

'------------------------------------------------------------------------------
' Columns functions
'------------------------------------------------------------------------------
'------------------------------------------------------------------------------
' Get the values of a table column in an array
'------------------------------------------------------------------------------

Public Function GetTableColumn(oTable As ListObject, colNbrOrName, Optional twoD As Boolean = False) As Variant
    Dim nbrRows As Long
    nbrRows = oTable.ListRows.Count
    Dim arr() As Variant
    ReDim arr(1 To nbrRows)
    arr = oTable.ListColumns(colNbrOrName).DataBodyRange.Value
    If (twoD) Then
        GetTableColumn = arr
    Else
        GetTableColumn = TwoDtoOneD(arr)
    End If
End Function


'------------------------------------------------------------------------------
' Sets one column of a table to a given array
' col may be an long (Column Nbr) or a String (Column name)
'------------------------------------------------------------------------------

Public Sub SetTableColumn(oTable, colNbrOrName As Variant, tValues As Variant, Optional withResize As Boolean = True)
    Dim arr() As Variant
    If ArrayNbrDimensions(tValues) = 1 Then
        arr = OneDtoTwoD(tValues)
    Else
        arr = tValues
    End If
    If (withResize) Then
        Call ResizeTable(oTable, UBound(arr, 1))
    End If
    oTable.ListColumns(colNbrOrName).DataBodyRange.Value = arr
End Sub

'------------------------------------------------------------------------------
' Clears data in a table object
'------------------------------------------------------------------------------
Public Sub ClearTableColumn(oTable As ListObject, colNbrOrName As Variant)
    tableSize = oTable.ListRows.Count
    Dim emptyArr(1 To tableSize) As String
    For i = 1 To tableSize
        emptyArr(i) = ""
    Next i
    Call SetTableColumn(oTable, colNbrOrName, emptyArr)
End Sub
'------------------------------------------------------------------------------
' Sets the formula in one column of a table
' col must be an long (Column Nbr)
'------------------------------------------------------------------------------
Public Sub SetTableColumnFormula(oTable As ListObject, colNbr As Long, theFormula As String)
    oTable.ListRows(1).Range.Cells(1, colNbr).Formula = theFormula
End Sub
'------------------------------------------------------------------------------
' Sets the number format in one column of a table
' col must be an long (Column Nbr)
'------------------------------------------------------------------------------
Public Sub SetTableColumnFormat(oTable As ListObject, colNbr As Long, theFormat As String)
    oTable.ListColumns(colNbr).DataBodyRange.NumberFormat = theFormat
End Sub

Public Function GetColumnNumberFromName(oTable As ListObject, columnName As String) As Long
    On Error GoTo Except
    GetColumnNumberFromName = oTable.ListColumns(columnName).index
    GoTo ThisIsTheEnd
Except:
    GetColumnNumberFromName = 0
ThisIsTheEnd:
    
End Function

Private Function GetColList(oTable As ListObject, currentColList) As Variant
    If (IsNumeric(currentColList)) Then
        nbrCols = oTable.ListColumns.Count
        ReDim localColList(1 To nbrCols) As Variant
        For C = 1 To nbrCols
            localColList(C) = C
        Next C
        GetColList = localColList
    Else
        GetColList = currentColList
    End If
End Function

Public Function appendTableColToArray(oTable, colNbrOrName, oArray) As Variant
    oldSize = UBound(oArray)
    addSize = oTable.ListRows.Count
    newSize = oldSize + addSize
    Dim arr() As Variant
    ReDim arr(1 To newSize)
    For i = 1 To oldSize
        arr(i) = oArray(i)
    Next i
    aArray = GetTableColumn(oTable, colNbrOrName)
    For i = 1 To addSize
        arr(i + oldSize) = aArray(i)
    Next i
    appendTableColToArray = arr
End Function


'------------------------------------------------------------------------------
' Row functions
'------------------------------------------------------------------------------
Public Sub ClearTableRow(oTable As ListObject, rowNbr As Long)
    For j = 1 To oTable.ListColumns.Count
        oTable.ListRows(rowNbr).DataBodyRange.Columns(j).Value = ""
    Next j
End Sub






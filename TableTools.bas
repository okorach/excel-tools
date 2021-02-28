Attribute VB_Name = "TableTools"
'==============================================================================
'
'  Table Tools
'
'==============================================================================

' resizeTable(oTable, targetSize as Integer)
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
Public Sub resizeTable(oTable, targetSize As Integer)
    ' TODO: Test TargetSize >0, oTable exists
    Dim i As Integer
    Dim nbRows As Integer
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

Public Sub truncateTable(oTable)
    Call resizeTable(oTable, 0)
End Sub

Private Function getColList(oTable, currentColList) As Variant
    If (IsNumeric(currentColList)) Then
        nbrCols = oTable.ListColumns.Count
        Dim localColList() As Variant
        ReDim localColList(1 To nbrCols)
        For c = 1 To nbrCols
            localColList(c) = c
        Next c
        getColList = localColList
    Else
        getColList = currentColList
    End If
End Function

'------------------------------------------------------------------------------
' Get the values of a table object cells in a 2 dimensions array
'------------------------------------------------------------------------------
Public Function getTableAsArray(oTable, Optional colList As Variant = 0) As Variant
    Dim nbrRows As Integer
    Dim nbrCols As Integer

    nbrCols = oTable.ListColumns.Count
    nbrRows = oTable.ListRows.Count
    Dim cList() As Variant
    cList = getColList(oTable, colList)
    Dim arr() As Variant
    ReDim arr(1 To nbrRows, 1 To nbrCols)
    i = 0
    For Each c In cList
        i = i + 1
        For j = 1 To nbrRows
            arr(j, i) = oTable.ListColumns(c).DataBodyRange.Rows(j).Value
        Next j
    Next c
    getTableAsArray = arr
End Function

Public Function GetColumnNumberFromName(oTable, columnName As String) As Long
    On Error GoTo Except
    GetColumnNumberFromName = oTable.ListColumns(columnName).index
    GoTo ThisIsTheEnd
Except:
    GetColumnNumberFromName = 0
ThisIsTheEnd:
    
End Function
'------------------------------------------------------------------------------
' Set the values of a 2D array into a table object
'------------------------------------------------------------------------------
Public Sub setTableFromArray(oTable, ByRef tValues)
    Call truncateTable(oTable)
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
' Get the values of a table column in an array
'------------------------------------------------------------------------------

Public Function getTableColumn(oTable, colNbrOrName, Optional twoD As Boolean = True) As Variant
    Dim nbrRows As Integer
    nbrRows = oTable.ListRows.Count
    Dim arr() As Variant
    ReDim arr(1 To nbrRows)
    arr = oTable.ListColumns(colNbrOrName).DataBodyRange.Value
    If (twoD) Then
        getTableColumn = arr
    Else
        getTableColumn = TwoDtoOneD(arr)
    End If
End Function

Public Function getArrayColumn(matrix As Variant, colNbr As Integer, Optional twoD As Boolean = True) As Variant
    Dim nbrRows As Integer
    nbrRows = UBound(matrix, 1)
    Dim arr() As Variant
    ReDim arr(1 To nbrRows)
    n = UBound(arr)
    For i = 1 To n
        arr(i) = matrix(i, colNbr)
    Next i
    If (twoD) Then
        getArrayColumn = OneDtoTwoD(arr)
    Else
        getArrayColumn = arr
    End If
End Function


'------------------------------------------------------------------------------
' Sets one column of a table to a given array
' col may be an integer (Column Nbr) or a String (Column name)
'------------------------------------------------------------------------------

Public Sub setTableColumn(oTable, colNbrOrName As Variant, tValues As Variant, Optional twoD As Boolean = True, Optional withResize As Boolean = True)

    Dim arr() As Variant
    If twoD Then
        arr = OneDtoTwoD(tValues)
    Else
        arr = tValues
    End If
    If (withResize) Then
        Call resizeTable(oTable, UBound(arr, 1))
    End If
    oTable.ListColumns(colNbrOrName).DataBodyRange.Value = arr
End Sub


'------------------------------------------------------------------------------
' Copies a table in another, possibly only specific columns,
' assuming both have the same structure
' Returns true in case of success, false in case of any error
'------------------------------------------------------------------------------

Public Function copyTable(oSrcTable, oTgtTable, Optional colList As Variant = 0) As Boolean
    Call resizeTable(oTgtTable, oSrcTable.ListRows.Count)
    If (IsNumeric(colList)) Then
        For i = 1 To oSrcTable.ListColumns.Count
           oTgtTable.ListColumns(i).DataBodyRange.Value = oSrcTable.ListColumns(i).DataBodyRange.Value
        Next i
    Else
        For Each col In colList
           oTgtTable.ListColumns(col).DataBodyRange.Value = oSrcTable.ListColumns(col).DataBodyRange.Value
        Next col
    End If
    copyTable = True
End Function

'------------------------------------------------------------------------------
' Appends oSrcTable at end of oTgtTable
'------------------------------------------------------------------------------

Public Sub appendTableToTable(oSrcTable, oTgtTable, Optional colList As Variant = 0)
    Offset = oTgtTable.ListRows.Count
    Call resizeTable(oTgtTable, Offset + oSrcTable.ListRows.Count)
    cList = getColList(oSrcTable, colList)
    For Each col In cList
        For j = 1 To oSrcTable.ListColumns(col).DataBodyRange.Rows.Count
            oTgtTable.ListColumns(col).DataBodyRange.Rows(j + Offset).Value = oSrcTable.ListColumns(col).DataBodyRange.Rows(j).Value
        Next j
    Next col
End Sub


Public Sub appendTableToTableFast(oSrcTable, oTgtTable, Optional colList As Variant = 0)
    sizeOffset = oTgtTable.ListRows.Count
    Call resizeTable(oTgtTable, sizeOffset + oSrcTable.ListRows.Count)
    cList = getColList(oSrcTable, colList)
    For Each col In cList
        Dim srcArr() As Variant
        Dim tgtArr() As Variant
        tgtArr = getTableColumn(oTgtTable, col)
        srcArr = getTableColumn(oSrcTable, col)
        sizeOffset = UBound(tgtArr)
        ReDim Arr1(UBound(srcArr) + sizeOffset)
        For i = 1 To UBound(srcArr)
            tgtArr(sizeOffset + i) = srcArr(i)
        Next i
    Next col
End Sub

Public Function appendTableColToArray(oTable, colNbrOrName, oArray) As Variant
    oldSize = UBound(oArray)
    addSize = oTable.ListRows.Count
    newSize = oldSize + addSize
    Dim arr() As Variant
    ReDim arr(1 To newSize)
    For i = 1 To oldSize
        arr(i) = oArray(i)
    Next i
    aArray = getTableColumn(oTable, colNbrOrName)
    For i = 1 To addSize
        arr(i + oldSize) = aArray(i)
    Next i
    appendTableColToArray = arr
End Function

'------------------------------------------------------------------------------
' Clears data in a table object
'------------------------------------------------------------------------------
Public Sub clearTableColumn(oTable As Variant, colNbrOrName As Variant)
    tableSize = oTable.ListRows.Count
    Dim emptyArr() As String
    ReDim emptyArr(1 To tableSize)
    For i = 1 To tableSize
        emptyArr(i) = ""
    Next i
    Call setTableColumn(oTable, colNbrOrName, emptyArr)
End Sub
'------------------------------------------------------------------------------
Public Sub clearTableRow(oTable, rowNbr)
    For j = 1 To oTable.ListColumns.Count
        oTable.ListRows(rowNbr).DataBodyRange.Columns(j).Value = ""
    Next j
End Sub
'------------------------------------------------------------------------------
Public Sub clearTable(oTable)
    Dim tableSize As Integer
    tableSize = oTable.ListRows.Count
    Call truncateTable(oTable)
    Call resizeTable(oTable, tableSize)
End Sub

'------------------------------------------------------------------------------
' Sets the formula in one column of a table
' col must be an integer (Column Nbr)
'------------------------------------------------------------------------------
Public Sub setTableColumnFormula(oTable, colNbr, theFormula)
    oTable.ListRows(1).Range.Cells(1, colNbr).Formula = theFormula
End Sub

'------------------------------------------------------------------------------
' Sets the number format in one column of a table
' col must be an integer (Column Nbr)
'------------------------------------------------------------------------------
Public Sub setTableColumnFormat(oTable, colNbr, theFormat)
    oTable.ListColumns(colNbr).DataBodyRange.NumberFormat = theFormat
End Sub

'------------------------------------------------------------------------------
' Converts 1D to 2D arrays and vice versa
'------------------------------------------------------------------------------
Public Function OneDtoTwoD(arr As Variant) As Variant
    Dim lb As Integer
    Dim ub As Integer
    lb = LBound(arr)
    ub = UBound(arr)
    Dim arr2d() As Variant
    ReDim arr2d(lb To ub, 1 To 1)
    For i = lb To ub
        arr2d(i, 1) = arr(i)
    Next i
    OneDtoTwoD = arr2d
End Function
'------------------------------------------------------------------------------
Public Function TwoDtoOneD(arr2d As Variant) As Variant
    Dim lb As Integer
    Dim ub As Integer
    lb = LBound(arr2d, 1)
    ub = UBound(arr2d, 1)
    Dim arr1d() As Variant
    ReDim arr1d(lb To ub)
    For i = lb To ub
        arr1d(i) = arr2d(i, 1)
    Next i
    TwoDtoOneD = arr1d
End Function
'------------------------------------------------------------------------------
Public Function Create1DArray(arraySize As Integer, elementValue As Variant) As Variant
    Dim arr1d() As Variant
    ReDim arr1d(1 To arraySize)
    For i = 1 To arraySize
        arr1d(i) = elementValue
    Next i
    Create1DArray = arr1d
End Function

'------------------------------------------------------------------------------
Public Function Create2DArray(arraySize As Integer, elementValue As Variant) As Variant
    Create2DArray = OneDtoTwoD(Create1DArray(arraySize, elementValue))
End Function

'-------------------------------------------------------------------------------
Public Function ArraySum(arr As Variant, Optional lb As Integer = -1, Optional ub As Integer = -1) As Double
    ' Calculate the sum of an (1D) array
    If lb = -1 Then
        lb = LBound(arr)
    End If
    If ub = -1 Then
        ub = UBound(arr)
    End If
    Dim sum As Double
    sum = 0
    For i = lb To ub
        sum = sum + arr(i)
    Next i
    ArraySum = sum
End Function

'-------------------------------------------------------------------------------
Public Function ArrayAverage(arr As Variant, Optional lb As Integer = -1, Optional ub As Integer = -1) As Double
    ' Calculate the average of an (1D) array
    If lb = -1 Then
        lb = LBound(arr)
    End If
    If ub = -1 Then
        ub = UBound(arr)
    End If
    ArrayAverage = ArraySum(arr, lb, ub) / (ub - lb + 1)
End Function

'------------------------------------------------------------------------------
' Returns nbr of dimensions of array
'------------------------------------------------------------------------------
Public Function NumberOfDimensions(arr As Variant) As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' NumberOfArrayDimensions
' This function returns the number of dimensions of an array. An unallocated dynamic array
' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Ndx As Integer
    Dim Res As Integer
    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(arr, Ndx)
    Loop Until Err.Number <> 0

    NumberOfDimensions = Ndx - 1
End Function

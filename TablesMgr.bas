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
Public Sub ResizeTable(oTable As ListObject, targetSize As Long)
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
        oTable.AutoFilter.ShowAllData
        oTable.ListRows(targetSize + 1).Range.Resize(nbRows - targetSize).Delete shift:=xlUp
        'For i = nbRows To targetSize + 1 Step -1
        '    oTable.ListRows(i).Delete
        'Next i
    End If

End Sub
'------------------------------------------------------------------------------
Public Sub TruncateTable(oTable As ListObject)
    Call ResizeTable(oTable, 0)
End Sub

'------------------------------------------------------------------------------
Public Sub ClearTableOld(oTable As ListObject)
    Dim tableSize As Long
    tableSize = oTable.ListRows.Count
    Call TruncateTable(oTable)
    Call ResizeTable(oTable, tableSize)
End Sub

Public Sub ClearTable(oTable As ListObject)
    If Not oTable Is Nothing Then
        oTable.DataBodyRange.ClearContents
    End If
End Sub

Public Sub SortTable(oTable As ListObject, sortCol1 As String, Optional sortOrder1 As XlSortOrder = xlAscending, _
                     Optional sortCol2 As String = vbNullString, Optional sortOrder2 As XlSortOrder = xlDescending)
    oTable.Sort.SortFields.Clear
    oTable.Sort.SortFields.Add key:=Range(oTable.name & "[" & sortCol1 & "]"), SortOn:=xlSortOnValues, Order:=sortOrder1, _
        DataOption:=xlSortNormal
    If LenB(sortCol2) > 0 Then
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
    ReDim arr(1 To nbrRows, 1 To nbrCols) As Variant
    i = 0
    For Each C In cList
        i = i + 1
        For j = 1 To nbrRows
            arr(j, i) = oTable.ListColumns(C).DataBodyRange.Rows(j).value
        Next j
    Next C
    GetTableAsArray = arr
End Function

'------------------------------------------------------------------------------
' Set the values of a 2D array into a table object
'------------------------------------------------------------------------------
Public Sub SetTableFromArray(oTable As ListObject, oArray As Variant)
    Call TruncateTable(oTable)
    nbrCols = oTable.ListColumns.Count
    nbrRows = UBound(oArray, 1)
    For i = 1 To nbrRows
        oTable.ListRows.Add
        For j = 1 To nbrCols
            oTable.ListColumns(j).DataBodyRange.Rows(i).value = oArray(i, j)
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
           oTgtTable.ListColumns(i).DataBodyRange.value = oSrcTable.ListColumns(i).DataBodyRange.value
        Next i
    Else
        For Each col In colList
           oTgtTable.ListColumns(col).DataBodyRange.value = oSrcTable.ListColumns(col).DataBodyRange.value
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
            oTgtTable.ListColumns(col).DataBodyRange.Rows(j + offset).value = oSrcTable.ListColumns(col).DataBodyRange.Rows(j).value
        Next j
    Next col
End Sub


Public Sub MergeTables(firstTable As ListObject, table As ListObject)
    firstTable.Resize Union(firstTable.Range, table.Range)
End Sub

Public Sub MergeTablesLong(firstTable As ListObject, table As ListObject)

    Dim headerRow As Range, tableRange As Range
        
    'Merge other tables in active sheet with first table
    
    For Each table In ActiveSheet.ListObjects
        If table.Range.row <> firstTable.Range.row Then
            Set headerRow = table.HeaderRowRange
            Set tableRange = table.Range
            table.Unlist
            headerRow.Delete
            firstTable.Resize Union(firstTable.Range, tableRange)
        End If
    Next
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
    ReDim arr(1 To nbrRows, 1 To 1) As Variant
    If nbrRows = 1 Then
        arr(1, 1) = oTable.ListColumns(colNbrOrName).DataBodyRange.value
    Else
        arr = oTable.ListColumns(colNbrOrName).DataBodyRange.value
    End If
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

Public Sub SetTableColumn(oTable As ListObject, colNbrOrName As Variant, oArray As Variant, Optional withResize As Boolean = True)
    Dim arr() As Variant
    If ArrayNbrDimensions(oArray) = 1 Then
        arr = OneDtoTwoD(oArray)
    Else
        arr = oArray
    End If
    If (withResize) Then
        Call ResizeTable(oTable, UBound(arr, 1))
    End If
    oTable.ListColumns(colNbrOrName).DataBodyRange.value = arr
End Sub

'------------------------------------------------------------------------------
' Clears data in a table object
'------------------------------------------------------------------------------
Public Sub ClearTableColumn(oTable As ListObject, colNbrOrName As Variant, Optional includeHeader As Boolean = False)
    If Not oTable Is Nothing Then
        If includeHeader Then
            oTable.ListColumns(colNbrOrName).Range.ClearContents
        Else
            oTable.ListColumns(colNbrOrName).DataBodyRange.ClearContents
        End If
    End If
End Sub
'------------------------------------------------------------------------------
' Sets the formula in one column of a table
' col must be an long (Column Nbr)
'------------------------------------------------------------------------------
Public Sub SetTableColumnFormula(oTable As ListObject, colNbr As Long, theFormula As String)
    If Not oTable Is Nothing Then
        oTable.ListRows(1).Range.Cells(1, colNbr).Formula = theFormula
    End If
End Sub
'------------------------------------------------------------------------------
' Sets the number format in one column of a table
' col must be an long (Column Nbr)
'------------------------------------------------------------------------------
Public Sub SetTableColumnFormat(oTable As ListObject, colNbr As Long, theFormat As String, Optional includeHeader As Boolean = True)
    If Not oTable Is Nothing Then
        If includeHeader Then
            oTable.ListColumns(colNbr).Range.NumberFormat = theFormat
        Else
            oTable.ListColumns(colNbrOrName).DataBodyRange.NumberFormat = theFormat
        End If
    End If
End Sub
'------------------------------------------------------------------------------
Public Sub SetTableStyle(oTable As ListObject, style As String)
    If Not oTable Is Nothing Then
        oTable.Range.ClearFormats
        oTable.TableStyle = style
    End If
End Sub

Public Function TableColNbrFromName(oTable As ListObject, columnName As String) As Integer
    TableColNbrFromName = 0
    On Error Resume Next
    TableColNbrFromName = oTable.ListColumns(columnName).index
ThisIsTheEnd:
    
End Function

Public Function TableColNbr(oTable As ListObject, colNameOrNumber As Variant) As Integer
    If VarType(colNameOrNumber) = vbString Then
        TableColNbr = TableColNbrFromName(oTable, CStr(colNameOrNumber))
    Else
        TableColNbr = colNameOrNumber
    End If
End Function
Public Function TableColumnNameExists(oTable As ListObject, columnName As String) As Integer
    TableColumnNameExists = (TableColNbrFromName(oTable, columnName) > 0)
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



Public Function AppendTableColToArray(oTable As ListObject, colNbrOrName As Variant, oArray As Variant) As Variant
    Dim oldSize As Long, addSize As Long
    oldSize = UBound(oArray)
    addSize = oTable.ListRows.Count
    ReDim arr(1 To oldSize + addSize) As Variant
    Dim appendArray() As Variant
    For i = 1 To oldSize
        arr(i) = oArray(i)
    Next i
    appendArray = GetTableColumn(oTable, colNbrOrName)
    For i = 1 To addSize
        arr(i + oldSize) = appendArray(i)
    Next i
    AppendTableColToArray = arr
End Function


Public Sub TableColumnFormatIcons(oTable As ListObject, colNameOrNumber As Variant)
    Dim col As Integer
    col = TableColNbr(oTable, colNameOrNumber)
    With oTable.ListColumns(col).DataBodyRange
        .FormatConditions.AddIconSetCondition
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
    End With
    With oTable.ListColumns(col).DataBodyRange.FormatConditions(1)
        .IconCriteria(2).Type = xlConditionValueNumber
        .IconCriteria(2).value = 0
        .IconCriteria(2).Operator = xlGreater
        .IconCriteria(3).Type = xlConditionValueNumber
        .IconCriteria(3).value = 0
        .IconCriteria(3).Operator = xlGreaterEqual
    End With
    oTable.ListColumns(col).DataBodyRange.Rows(1).Select
    With Selection
        .FormatConditions.Add Type:=xlExpression, Formula1:="=AND($A10=$A11" + Application.International(xlListSeparator) + "$B10=$B11)"
    End With
    With Selection.FormatConditions(2).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Selection.FormatConditions(2).StopIfTrue = False
    Selection.Copy
    oTable.ListColumns(col).DataBodyRange.PasteSpecial Paste:=xlPasteFormats
End Sub
'------------------------------------------------------------------------------
' Row functions
'------------------------------------------------------------------------------
Public Sub ClearTableRow(oTable As ListObject, rowNbr As Long)
    For j = 1 To oTable.ListColumns.Count
        oTable.ListRows(rowNbr).DataBodyRange.Columns(j).value = vbNullString
    Next j
End Sub

Public Function TableIndex(likeString As String, Optional ws As Worksheet = Nothing, Optional wb As Workbook = Nothing) As Integer
    If wb Is Nothing Then
        wb = ActiveWorkbook
    End If
    If ws Is Nothing Then
        ws = wb.ActiveSheet
    End If
    Dim i As Long
    accountTableIndex = 0
    For i = 1 To ws.ListObjects.Count
        If LCase$(ws.ListObjects(i).name) Like "*_" & likeString Then
            TableIndex = i
            Exit For
        End If
    Next i
End Function

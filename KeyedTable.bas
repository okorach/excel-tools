Attribute VB_Name = "KeyedTable"


Public Function KeyedTableValue(oTable As ListObject, key As Variant, colNameOrNumber As Variant) As Variant
    Dim col As Integer
    col = TableColNbr(oTable, colNameOrNumber)
    On Error Resume Next
    KeyedTableValue = Application.WorksheetFunction.VLookup(key, oTable.DataBodyRange, col, False)
    If Err.Number <> 0 Then
        KeyedTableValue = vbNull
    End If
    On Error GoTo 0
End Function

Public Sub KeyedTableInsertOrReplace(oTable As ListObject, key As Variant, value As Variant, colNameOrNumber As Variant)
    Dim col As Integer
    Dim row As ListRow
    col = TableColNbr(oTable, colNameOrNumber)
    With oTable
        For Each row In .ListRows
            If row.Range.Cells(1, 1).value = key Then
                row.Range.Cells(1, col).value = value
                Exit Sub
            End If
        Next row
        .ListRows.Add
        .ListRows(.ListRows.Count).Range.Cells(1, 1).value = key
        .ListRows(.ListRows.Count).Range.Cells(1, col).value = value
    End With
End Sub

Public Function KeyedTableInsert(oTable As ListObject, key As Variant, value As Variant, colNameOrNumber As Variant) As Boolean
    Dim col As Integer
    Dim row As ListRow
    col = TableColNbr(oTable, colNameOrNumber)
    KeyedTableInsert = False
    If KeyedTableValue(oTable, key, col) <> vbNull Then
        Exit Function
    End If
    With oTable
        .ListRows.Add
        .ListRows(.ListRows.Count).Range.Cells(1, 1).value = key
        .ListRows(.ListRows.Count).Range.Cells(1, col).value = value
    End With
    KeyedTableInsert = True
End Function

Public Function KeyedTableReplace(oTable As ListObject, key As Variant, value As Variant, colNameOrNumber As Variant) As Boolean
    Dim col As Integer
    Dim row As ListRow
    col = TableColNbr(oTable, colNameOrNumber)
    KeyedTableReplace = False
    If KeyedTableValue(oTable, key, col) = vbNull Then
        Exit Function
    End If
    For Each row In oTable.ListRows
        If row.Range.Cells(1, 1).value = key Then
            row.Range.Cells(1, col).value = value
            KeyedTableReplace = True
            Exit Function
        End If
    Next row
End Function


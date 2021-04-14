VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyedTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private table As ListObject

Private Sub Class_Initialize()
    Set table = Nothing
End Sub

Public Sub Init(oTable As ListObject)
    Set table = oTable
End Sub


Public Function Lookup(key As Variant, colNameOrNumber As Variant) As Variant
    Dim col As Integer
    col = TableColNbr(table, colNameOrNumber)
    On Error Resume Next
    Lookup = Application.WorksheetFunction.vlookup(key, table.DataBodyRange, col, False)
    If Err.Number <> 0 Then
        Lookup = vbNull
    End If
    On Error GoTo 0
End Function

Public Function Update(key As Variant, value As Variant, colNameOrNumber As Variant) As Boolean
    Update = False

    Dim col As Integer
    Dim row As ListRow
    col = TableColNbr(table, colNameOrNumber)
    With table
        For Each row In .ListRows
            If row.Range.Cells(1, 1).value = key Then
                row.Range.Cells(1, col).value = value
                Update = True
                Exit Function
            End If
        Next row
    End With
End Function

Public Function Insert(key As Variant, Optional value As Variant = vbNull, Optional colNameOrNumber As Variant = vbNull) As Boolean
    Insert = False
    If Lookup(key, 2) <> vbNull Then
        Exit Function
    End If
    Dim row As ListRow
    With table
        .ListRows.Add
        .ListRows(.ListRows.Count).Range.Cells(1, 1).value = key
        If value <> vbNull And colNameOrNumber <> vbNull Then
            .ListRows(.ListRows.Count).Range.Cells(1, TableColNbr(table, colNameOrNumber)).value = value
        End If
    End With
    Insert = True
End Function


Public Function InsertOrUpdate(key As Variant, value As Variant, colNameOrNumber As Variant) As Boolean
    Dim b As Boolean
    b = Update(key, value, col)
    If Not b Then
        b = Insert(key, value, col)
    End If
    InsertOrUpdate = b
End Function


Public Function KeyExists(key As Variant) As Boolean
    KeyExists = (Lookup(key, 2) <> vbNull)
End Function

Public Function KeyRow(key As Variant) As Long
    KeyRow = 0
    With table
        For r = 1 To .ListRows.Count
            If .ListRows(r).Range.Cells(1, 1).value = key Then
                KeyRow = r
                Exit Function
            End If
        Next r
    End With
End Function


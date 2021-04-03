Attribute VB_Name = "Parameters"
Public Const PARAMS_SHEET As String = "Paramètres"
Public Const GLOBAL_PARAMS_TABLE As String = "TblGlobalParams"

Public Function GetGlobalParam(paramKey As String) As Variant
    GetGlobalParam = KeyedTableValue(Sheets(PARAMS_SHEET).ListObjects(GLOBAL_PARAMS_TABLE), paramKey, 2)
End Function

Public Sub SetGlobalParam(paramKey As String, paramValue As Variant)
    Call KeyedTableInsertOrReplace(Sheets(PARAMS_SHEET).ListObjects(GLOBAL_PARAMS_TABLE), paramKey, paramValue, 2)
End Sub


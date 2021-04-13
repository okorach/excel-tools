Attribute VB_Name = "Parameters"
Public Const PARAMS_SHEET As String = "Paramètres"
Public Const GLOBAL_PARAMS_TABLE As String = "TblGlobalParams"
Public Const CURRENCIES_TABLE As String = "TblCurrencies"

Public Function GetGlobalParam(paramKey As String, Optional wb As Workbook = Nothing) As Variant
    If wb Is Nothing Then
        GetGlobalParam = KeyedTableValue(Sheets(PARAMS_SHEET).ListObjects(GLOBAL_PARAMS_TABLE), paramKey, 2)
    Else
        GetGlobalParam = KeyedTableValue(wb.Sheets(PARAMS_SHEET).ListObjects(GLOBAL_PARAMS_TABLE), paramKey, 2)
    End If
End Function

Public Sub SetGlobalParam(paramKey As String, paramValue As Variant, Optional wb As Workbook = Nothing)
    If wb Is Nothing Then
        Call KeyedTableInsertOrReplace(Sheets(PARAMS_SHEET).ListObjects(GLOBAL_PARAMS_TABLE), paramKey, paramValue, 2)
    Else
        Call KeyedTableInsertOrReplace(wb.Sheets(PARAMS_SHEET).ListObjects(GLOBAL_PARAMS_TABLE), paramKey, paramValue, 2)
    End If
End Sub


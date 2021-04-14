Attribute VB_Name = "Parameters"
Public Const PARAMS_SHEET As String = "Paramètres"
Public Const GLOBAL_PARAMS_TABLE As String = "TblGlobalParams"
Public Const CURRENCIES_TABLE As String = "TblCurrencies"


Public Function GetGlobalParam(paramKey As String, Optional wb As Workbook = Nothing) As Variant
    If wb Is Nothing Then
        Set wb = ActiveWorkbook
    End If
    Dim paramsTable As KeyedTable
    Set paramsTable = NewKeyedTable(wb.Sheets(PARAMS_SHEET).ListObjects(GLOBAL_PARAMS_TABLE))
    GetGlobalParam = paramsTable.Lookup(paramKey, 2)
End Function


Public Function SetGlobalParam(paramKey As String, paramValue As Variant, Optional wb As Workbook = Nothing) As Boolean
    If wb Is Nothing Then
        Set wb = ActiveWorkbook
    End If
    Dim paramsTable As KeyedTable
    Set paramsTable = NewKeyedTable(wb.Sheets(PARAMS_SHEET).ListObjects(GLOBAL_PARAMS_TABLE))
    SetGlobalParam = paramsTable.InsertOrUpdate(paramKey, paramValue, 2)
End Function


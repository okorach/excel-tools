Attribute VB_Name = "Localization"
Public Function GetLabel(key)
    Dim col As Integer
    Dim lang As String
    lang = GetNamedVariableValue("Language")
    If lang = "English" Then
        col = 3
    Else
        col = 2
    End If
    GetLabel = Application.vlookup(key, Sheets("Language").ListObjects("TblKeys").DataBodyRange, col, False)
    If IsError(GetLabel) Then
        GetLabel = key & " not found"
    End If
End Function


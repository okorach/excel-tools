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

Public Sub GoToAccount()
    selectedNbr = Range("Solde!H72").Value
    Dim accountName As String
    accountName = Sheets(PARAMS_SHEET).Range("L" & CStr(selectedNbr + 1))
    If accountExists(accountName) Then
        Sheets(accountName).Activate
    End If
End Sub

Public Sub GoToNext()
    curr = ActiveSheet.Index
    If (curr < Sheets.Count) Then
        Sheets(curr + 1).Activate
    End If
End Sub

Public Sub GoToPrev()
    curr = ActiveSheet.Index
    If (curr > 1) Then
        Sheets(curr - 1).Activate
    End If
End Sub



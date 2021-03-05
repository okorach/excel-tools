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
    Dim accountName As String
    accountName = getSelectedAccount()
    If AccountExists(accountName) Then
        Sheets(accountName).Activate
    End If
End Sub

Public Sub GoToSheet(shift As Integer)
    Dim curr As Integer
    curr = ActiveSheet.index
    If (curr + shift) > 0 And (curr + shift) <= Sheets.Count Then
        Sheets(curr + shift).Activate
    ElseIf shift < 0 Then
        Sheets(1).Activate
    ElseIf shift > 0 Then
        Sheets(Sheets.Count).Activate
    End If
End Sub

Public Sub GoToNext()
    Call GoToSheet(1)
End Sub

Public Sub GoToPrev()
    Call GoToSheet(-1)
End Sub

Public Sub GoBack5()
    Call GoToSheet(-5)
End Sub

Public Sub GoFwd5()
    Call GoToSheet(5)
End Sub

Public Sub GoToSolde()
    Sheets(BALANCE_SHEET).Activate
End Sub

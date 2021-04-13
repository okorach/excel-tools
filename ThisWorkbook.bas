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
    Dim accName As String
    accName = getSelectedAccount()
    If AccountExists(accName) Then
        Sheets(accName).Activate
    End If
End Sub




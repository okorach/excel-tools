VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public msg As String
Public progress As Integer
Public goal As Integer
Public form As UserForm


Private Sub Class_Initialize()
    progress = 0
    Set form = ProgressForm
End Sub

Private Sub Class_Terminate()
    ProgressForm.Hide
End Sub

Public Sub Init(Optional pMsg As String = vbNullString, Optional pGoal As Integer = -1, Optional pProgress As Integer = -1)
    If pGoal >= 0 Then
        goal = pGoal
    End If
    If pProgress >= 0 Then
        progress = pProgress
    End If
    If pMsg <> vbNullString Then
        msg = pMsg
        form.MsgBox.Caption = CStr(msg & "..." & vbCrLf & vbCrLf & CStr((progress * 100) \ goal) & " %")
    End If
End Sub

Public Sub Update(Optional steps As Integer = 1)
    progress = progress + steps
    form.MsgBox.Caption = CStr(msg & "..." & vbCrLf & vbCrLf & CStr((progress * 100) \ goal) & " %")
    form.Repaint
End Sub

Public Sub CloseForm()
    ProgressForm.Hide
End Sub

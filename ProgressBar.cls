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
Public progress As Long
Public goal As Long
Public form As UserForm
Public startDate As Date


Private Sub Class_Initialize()
    progress = 0
    Set form = ProgressForm
    startDate = vbNull
End Sub

Private Sub Class_Terminate()
    ProgressForm.Hide
End Sub

Public Sub Init(Optional pMsg As String = vbNullString, Optional pGoal As Integer = -1, Optional pTimer As Boolean = False, Optional pProgress As Integer = -1)
    Dim startMsg As String
    startMsg = ""
    If pGoal >= 0 Then
        goal = pGoal
    End If
    If pProgress >= 0 Then
        progress = pProgress
    End If
    If pTimer Then
        startDate = Now()
        startMsg = "00:00 - "
    End If
    If pMsg <> vbNullString Then
        msg = pMsg
        form.MsgBox.Caption = CStr(msg & "..." & vbCrLf & vbCrLf & startMsg & CStr((progress * 100) \ goal) & " %")
    End If
End Sub

Public Sub Update(Optional steps As Integer = 1)
    progress = progress + steps
    Dim startMsg As String
    If startDate = vbNull Then
        startMsg = ""
    Else
        Dim seconds As Integer, minutes As Integer
        seconds = DateDiff("s", startDate, Now())
        minutes = seconds \ 60
        seconds = seconds - minutes * 60
        startMsg = Format$(minutes, "00") & ":" & Format$(seconds, "00") & " - "
    End If
    form.MsgBox.Caption = CStr(msg & "..." & vbCrLf & vbCrLf & startMsg & CStr((progress * 100) \ goal) & " %")
    form.Repaint
End Sub

Public Sub CloseForm()
    ProgressForm.Hide
End Sub

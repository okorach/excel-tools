Attribute VB_Name = "ProgressBar"
Public Sub ProgressBarStart(workInProgress As String, totalSteps As Integer)
    ProgressForm.Show False
    Call SetNamedVariableValue("progressBarMessage", workInProgress)
    Call SetNamedVariableValue("progressBarTotal", totalSteps)
    Call SetNamedVariableValue("progressBarCurrent", 0)
    Dim msg As String
    msg = workInProgress & "..." & vbCrLf & vbCrLf & "0 %"
    Call ProgressBarUpdate
End Sub
Public Sub ProgressBarUpdate()
    Dim msg As String
    Dim workInProgress As String, step As Long
    step = GetNamedVariableValue("progressBarCurrent")
    total = GetNamedVariableValue("progressBarTotal")
    step = step + 1
    Call SetNamedVariableValue("progressBarCurrent", step)
    msg = CStr(GetNamedVariableValue("progressBarMessage")) & "..." & vbCrLf & vbCrLf & CStr((step * 100) \ CLng(GetNamedVariableValue("progressBarTotal"))) & " %"
    ProgressForm.ProgressFormMessage.Caption = msg
    ProgressForm.Repaint
End Sub
Public Sub ProgressBarStop()
    ProgressForm.Hide
End Sub


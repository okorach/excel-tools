Attribute VB_Name = "ProgressBar"
Public Sub ProgressBarStart(initialMessage As String)
    ProgressForm.Show False
    Call ProgressBarUpdate(initialMessage)
End Sub
Public Sub ProgressBarUpdate(msg As String)
    ProgressForm.ProgressFormMessage.Caption = msg
    ProgressForm.Repaint
End Sub
Public Sub ProgressBarStop()
    ProgressForm.Hide
End Sub


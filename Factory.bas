Attribute VB_Name = "Factory"
Public Function NewProgressBar(msg As String, goal As Integer, Optional show As Boolean = True) As ProgressBar
    Set NewProgressBar = New ProgressBar
    NewProgressBar.Init msg, goal
    NewProgressBar.form.MsgBox.Caption = msg
    If show Then
        ProgressForm.show False
        'NewProgressBar.form.show False
    End If
End Function

Public Function NewInterest(accountId As String, Optional balTbl As ListObject = Nothing, Optional depTbl As ListObject = Nothing) As Interest
    Set NewInterest = New Interest
    NewInterest.Init accountId, balTbl, depTbl
End Function



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


Public Function NewInterest(accountId As String, Optional balancesArray As Variant = Nothing, Optional depositsArray As Variant = Nothing, _
                            Optional InterestPeriod As Integer = 1) As Interest
    Set NewInterest = New Interest
    NewInterest.Init accountId, balancesArray, depositsArray, InterestPeriod
End Function


Public Function NewKeyedTable(oTable As ListObject) As KeyedTable
    Set NewKeyedTable = New KeyedTable
    NewKeyedTable.Init oTable
End Function




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


Public Function LoadAccount(accountId As String) As Account
    Set LoadAccount = New Account
    If Not LoadAccount.Load(accountId) Then
        Set LoadAccount = Nothing
    End If
End Function

Public Function NewAccount(aId As String, aNbr As String, aBank As String, Optional aCur As String = vbNullString, _
                           Optional aType As String = vbNullString, Optional aAvail As Integer = 0, _
                           Optional aInB As Boolean = False, Optional aTax As Double = 0) As Account
    Set NewAccount = New Account
    If Not NewAccount.Create(aId, aNbr, aBank, aCur, aType, aAvail, aInB, aTax) Then
        Set NewAccount = Nothing
    End If
End Function


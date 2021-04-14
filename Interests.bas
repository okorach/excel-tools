Attribute VB_Name = "Interests"
Public Sub InterestsCalcHere()
    Call FreezeDisplay
    Call InterestsCalc(getAccountId(ActiveSheet))
    Call UnfreezeDisplay
End Sub


Public Sub InterestsCalc(accountId As String, Optional withModal As Boolean = True)
    Dim interestPeriod As Integer
    interestPeriod = AccountInterestPeriod(AccountType(accountId))
    If interestPeriod > 0 Then
        Dim deposits As Variant
        Dim balances As Variant
        deposits = AccountDepositHistory(accountId)
        balances = AccountBalanceHistory(accountId, "Yearly")

        Dim accountInterests As Interest
        Set accountInterests = NewInterest(accountId, balances, deposits, interestPeriod)
        accountInterests.Calc
        accountInterests.Store AccountTaxRate(accountId)
    End If
End Sub


Public Sub InterestsCalcAll()
    Dim modal As ProgressBar
    Set modal = NewProgressBar("Interests calculation in progress", Worksheets.Count)
    Call FreezeDisplay
    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim accountId As String
        accountId = getAccountId(ws)
        If IsAnAccount(ws) And AccountIsOpen(accountId) And IsInterestAccount(accountId) Then
            Call InterestsCalc(accountId, withModal:=False)
        End If
        modal.Update
    Next ws
    Call UnfreezeDisplay
    Set modal = Nothing
End Sub


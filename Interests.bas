Attribute VB_Name = "Interests"
Public Sub InterestsCalcHere()
    Call FreezeDisplay
    Dim oAccount As Account
    Set oAccount = LoadAccount(getAccountId(ActiveSheet))
    If Not oAccount Is Nothing Then
        Call oAccount.CalcInterests(force:=True)
    End If
    Call UnfreezeDisplay
End Sub


Public Sub InterestsCalc(accountId As String, Optional withModal As Boolean = True)
    Dim InterestPeriod As Integer
    InterestPeriod = AccountInterestPeriod(AccountType(accountId))
    If InterestPeriod > 0 Then
        Dim deposits As Variant
        Dim balances As Variant
        deposits = AccountDepositHistory(accountId)
        balances = AccountBalanceHistory(accountId, "Yearly")

        Dim accountInterests As Interest
        Set accountInterests = NewInterest(accountId, balances, deposits, InterestPeriod)
        accountInterests.Calc
        accountInterests.Store AccountTaxRate(accountId)
    End If
End Sub


Public Sub InterestsCalcAll()
    Dim modal As ProgressBar
    Set modal = NewProgressBar("Interests calculation in progress", AccountsCount(openOnly:=True, interestOnly:=True, noYearlyInterest:=True) * 6, True)
    Call FreezeDisplay
    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim oAccount As Account
        Set oAccount = LoadAccount(getAccountId(ws))
        If Not (oAccount Is Nothing) Then
            If oAccount.IsOpen() And oAccount.HasInterests() Then
                Call oAccount.CalcInterests(modal)
            End If
        End If
    Next ws
    Call UnfreezeDisplay
    Set modal = Nothing
End Sub


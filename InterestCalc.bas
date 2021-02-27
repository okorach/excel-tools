Attribute VB_Name = "InterestCalc"
Public Const INTEREST_CALC_SHEET As String = "Calculator"
Private Const BALANCE_HISTORY_TABLE As String = "TableBalanceHistory"

Sub CalcInterestForAllAccounts()

    freezeDisplay
    
    For i = 1 To Sheets.Count
        If (Sheets(i).name <> INTEREST_CALC_SHEET And Sheets(i).name <> "Params" And Sheets(i).name <> "Summary") Then
            Call CalcInterestForAccount(Sheets(i).name)
        End If
    Next i
    
    unfreezeDisplay
    
End Sub

Sub CalcInterestForAccount(accName As String)
    Call ImportToCalculator(accName)
    Call CalcAllInterests
    Sheets(INTEREST_CALC_SHEET).Activate
End Sub

Sub CalcAndStoreInterestForAccount(accName As String)
    Call CalcInterestForAccount(accName)
    Call ExportFromCalculator(accName)
End Sub

Sub ImportAccount()
    Call ImportToCalculator(getSelectedAccount())
End Sub

Sub ExportAccount()
    Call ExportFromCalculator(getSelectedAccount())
End Sub

Sub ExportFromCalculator(accName As String)
    Call ExportInterestResults(accName)
End Sub

Sub ImportToCalculator(accName As String)

    freezeDisplay
    Sheets(INTEREST_CALC_SHEET).Range("G1").Value = "Deposit history for " & accName
    Sheets(INTEREST_CALC_SHEET).Range("L1").Value = "Balance history for " & accName
    
    deposits = getDepositHistory(accName)
    balances = getBalanceHistory(accName, "Yearly")
    Call resizeTable(Sheets(INTEREST_CALC_SHEET).ListObjects(2), UBound(deposits, 1))
    Call resizeTable(Sheets(INTEREST_CALC_SHEET).ListObjects(1), UBound(balances, 1))
    
    ' Sheets(accName).ListObjects(1).name = "TableBalance" & Replace(accName, " ", "")
    ' Sheets(accName).ListObjects(2).name = "TableDeposit" & Replace(accName, " ", "")
    
    ' Copy 2 first columns of the 2 tables with history of deposits (date/amount) and history of balance (date/amount)
    Call setTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects(2), 1, getArrayColumn(deposits, 1, False))
    Call setTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects(2), 2, getArrayColumn(deposits, 2, False))
    Call setTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects(1), 1, getArrayColumn(balances, 1, False))
    Call setTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects(1), 2, getArrayColumn(balances, 2, False))
    'Sheets("Calculator").ListObjects(2).ListColumns(3).DataBodyRange.Cells(1).formula = "=IF(OR([Date]>target_date,[Date]<=start_date),0,FLOOR((target_date-[Date])/15.2,1))"
    'Sheets("Calculator").ListObjects(2).ListColumns(4).DataBodyRange.Cells(1).formula = "=IF([Nbr de périodes]<=0;IF(OR([Date]>=target_date;[Date]<=start_date);0;[Montant]);[Montant]*(1+$R$1)^[Nbr de périodes])"
    
    ' Clear old calculated interest rates
    Call clearTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects(1), 3)
    Call clearTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects(1), 4)
    
    unfreezeDisplay
End Sub

Sub ExportInterestResults(accName)
    Call setTableColumn(Sheets(accName).ListObjects(1), getTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects("TableBalanceHistory"), 3), 3)
    Call setTableColumn(Sheets(accName).ListObjects(1), getTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects("TableBalanceHistory"), 4), 4)
End Sub


Sub CalcAllInterests()
    Call CalcCompoundInterests
    Call CalcPeriodicInterests
End Sub

Sub CalcCompoundInterests()
    Sheets(INTEREST_CALC_SHEET).Range("B5").Value = 0

    For i = 2 To Sheets(INTEREST_CALC_SHEET).ListObjects("TableBalanceHistory").ListRows.Count
        Sheets(INTEREST_CALC_SHEET).Range("B2").Value = Sheets(INTEREST_CALC_SHEET).ListObjects(BALANCE_HISTORY_TABLE).ListColumns(1).DataBodyRange.Rows(1).Value
        Sheets(INTEREST_CALC_SHEET).Range("B3").Value = Sheets(INTEREST_CALC_SHEET).ListObjects(BALANCE_HISTORY_TABLE).ListColumns(1).DataBodyRange.Rows(i).Value
        Sheets(INTEREST_CALC_SHEET).Range("B4").GoalSeek Goal:=Sheets(INTEREST_CALC_SHEET).Range("C3").Value, ChangingCell:=Sheets(INTEREST_CALC_SHEET).Range("B5")
        Sheets(INTEREST_CALC_SHEET).ListObjects(BALANCE_HISTORY_TABLE).ListColumns(4).DataBodyRange.Rows(i).Value = Sheets(INTEREST_CALC_SHEET).Range("B5").Value
    Next i
End Sub

Sub CalcPeriodicInterests()
    Sheets(INTEREST_CALC_SHEET).Range("B5").Value = 0
    For i = 2 To Sheets(INTEREST_CALC_SHEET).ListObjects("TableBalanceHistory").ListRows.Count
        Sheets(INTEREST_CALC_SHEET).Range("B2").Value = Sheets(INTEREST_CALC_SHEET).ListObjects(BALANCE_HISTORY_TABLE).ListColumns(1).DataBodyRange.Rows(i - 1).Value
        Sheets(INTEREST_CALC_SHEET).Range("B3").Value = Sheets(INTEREST_CALC_SHEET).ListObjects(BALANCE_HISTORY_TABLE).ListColumns(1).DataBodyRange.Rows(i).Value
        Sheets(INTEREST_CALC_SHEET).Range("B5").Value = 0.1
        Sheets(INTEREST_CALC_SHEET).Range("B4").GoalSeek Goal:=Sheets(INTEREST_CALC_SHEET).Range("C3").Value, ChangingCell:=Sheets(INTEREST_CALC_SHEET).Range("B5")
        Sheets(INTEREST_CALC_SHEET).ListObjects(BALANCE_HISTORY_TABLE).ListColumns(3).DataBodyRange.Rows(i).Value = Sheets(INTEREST_CALC_SHEET).Range("B5").Value
    Next i
End Sub

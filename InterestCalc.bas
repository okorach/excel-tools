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
    Call CalcPeriodicInterests
    ' Call CalcCompoundInterests
    Call ExportInterestResults(accName)
End Sub

Sub CalcAndStoreInterestForAccount(accName As String)
    Call CalcInterestForAccount(accName)
    Call ExportFromCalculator(accName)
    Call ExportInterestResults(accName)
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
    Dim oCalcSheet, oDepositTbl, oBalanceTbl As Variant
    freezeDisplay
    With Sheets(INTEREST_CALC_SHEET)
        .Range("G1").Value = "Deposit history for " & accName
        .Range("L1").Value = "Balance history for " & accName
        deposits = getDepositHistory(accName)
        balances = getBalanceHistory(accName, "Yearly")
        Call resizeTable(.ListObjects(1), UBound(balances, 1))
        Call resizeTable(.ListObjects(2), UBound(deposits, 1))

    
        ' oBalanceTbl.name = "TableBalance" & Replace(accName, " ", "")
        ' oDepositTbl.name = "TableDeposit" & Replace(accName, " ", "")
        
        ' Copy 2 first columns of the 2 tables with history of deposits (date/amount) and history of balance (date/amount)
        Call setTableColumn(.ListObjects(2), 1, getArrayColumn(deposits, 1, False))
        Call setTableColumn(.ListObjects(2), 2, getArrayColumn(deposits, 2, False))
        Call setTableColumn(.ListObjects(1), 1, getArrayColumn(balances, 1, False))
        Call setTableColumn(.ListObjects(1), 2, getArrayColumn(balances, 2, False))
        '.ListObjects(2).ListColumns(3).DataBodyRange.Cells(1).formula = "=IF(OR([Date]>target_date,[Date]<=start_date),0,FLOOR((target_date-[Date])/15.2,1))"
        '.ListObjects(2).ListColumns(4).DataBodyRange.Cells(1).formula = "=IF([Nbr de périodes]<=0;IF(OR([Date]>=target_date;[Date]<=start_date);0;[Montant]);[Montant]*(1+$R$1)^[Nbr de périodes])"
        
        ' Clear old calculated interest rates
        Call clearTableColumn(.ListObjects(1), 3)
        Call clearTableColumn(.ListObjects(1), 4)
    End With
    unfreezeDisplay
End Sub

Sub ExportInterestResults(accName As String)
    Dim yields As Variant
    Dim colOffset As String
    Dim n, i, k As Integer
    yields = getTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects("TableBalanceHistory"), 3, False)
    nbrYields = UBound(yields)
    yieldIndex = AccountYieldsTableIndex(accName)
    With Sheets(accName)
        'colOffset = "I"
        'If .Range("B7") = "Shares" Then
        '    colOffset = "H"
        'End If
        .ListObjects(yieldIndex).ListColumns(2).DataBodyRange.Rows(1).Value = yields(nbrYields)
        For k = 2 To 5
         .ListObjects(yieldIndex).ListColumns(2).DataBodyRange.Rows(k).Value = "-"
        Next k
        If nbrYields >= 2 Then
            .ListObjects(yieldIndex).ListColumns(2).DataBodyRange.Rows(2).Value = yields(nbrYields - 1)
        End If
        If nbrYields >= 4 Then
            .ListObjects(yieldIndex).ListColumns(2).DataBodyRange.Rows(3).Value = ArrayAverage(yields, nbrYields - 3, nbrYields - 1)
        End If
        If nbrYields >= 6 Then
            .ListObjects(yieldIndex).ListColumns(2).DataBodyRange.Rows(4).Value = ArrayAverage(yields, nbrYields - 5, nbrYields - 1)
        End If
        If nbrYields >= 2 Then
            .ListObjects(yieldIndex).ListColumns(2).DataBodyRange.Rows(5).Value = ArrayAverage(yields, 1, nbrYields - 1)
        End If
    End With
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

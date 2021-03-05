Attribute VB_Name = "InterestsMgr"
Public Const INTEREST_CALC_SHEET As String = "Calculator"
Private Const BALANCE_HISTORY_TABLE As String = "TableBalanceHistory"


Public Function InterestsCalc(balanceArray As Variant, depositsArray As Variant, Optional account As String = "account", Optional calcPerPeriod As Boolean = True)
    FreezeDisplay
    Call InterestsLoadData(balanceArray, depositsArray, account)
    InterestsCalc = InterestsCalcFromData(calcPerPeriod)
    UnfreezeDisplay
End Function


Private Sub InterestsLoadData(balancesArray As Variant, depositsArray As Variant, Optional accName As String = "account")
    ' Loads data need for interests calculation in the calculation sheet
    With Sheets(INTEREST_CALC_SHEET)
        .Range("G1").Value = "Deposit history for " & accName
        .Range("L1").Value = "Balance history for " & accName

        Call resizeTable(.ListObjects(1), UBound(balancesArray, 1))
        Call resizeTable(.ListObjects(2), UBound(depositsArray, 1))
        
        ' Copy 2 first columns of the 2 tables with history of deposits (date/amount) and history of balance (date/amount)
        Call SetTableColumn(.ListObjects(2), 1, getArrayColumn(depositsArray, 1, False))
        Call SetTableColumn(.ListObjects(2), 2, getArrayColumn(depositsArray, 2, False))
        Call SetTableColumn(.ListObjects(1), 1, getArrayColumn(balancesArray, 1, False))
        Call SetTableColumn(.ListObjects(1), 2, getArrayColumn(balancesArray, 2, False))
        '.ListObjects(2).ListColumns(3).DataBodyRange.Cells(1).formula = "=IF(OR([Date]>target_date,[Date]<=start_date),0,FLOOR((target_date-[Date])/15.2,1))"
        '.ListObjects(2).ListColumns(4).DataBodyRange.Cells(1).formula = "=IF([Nbr de périodes]<=0;IF(OR([Date]>=target_date;[Date]<=start_date);0;[Montant]);[Montant]*(1+$R$1)^[Nbr de périodes])"
        
        ' Clear old calculated interest rates
        Call clearTableColumn(.ListObjects(1), 3)
    End With
End Sub


Private Function InterestsCalcFromData(Optional calcPerPeriod As Boolean = True)
    ' Calculate interest rates for each delta period or since the beginning of the balance history sheet
    With Sheets(INTEREST_CALC_SHEET)
        .Range("B5").Value = 0
        For i = 2 To .ListObjects(BALANCE_HISTORY_TABLE).ListRows.Count
            If calcPerPeriod Then
                .Range("B2").Value = .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(1).DataBodyRange.Rows(i - 1).Value
            Else
                .Range("B2").Value = .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(1).DataBodyRange.Rows(1).Value
            End If
            .Range("B3").Value = .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(1).DataBodyRange.Rows(i).Value
            .Range("B5").Value = 0.1
            .Range("B4").GoalSeek Goal:=.Range("C3").Value, ChangingCell:=.Range("B5")
            .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(3).DataBodyRange.Rows(i).Value = .Range("B5").Value
        Next i
        InterestsCalcFromData = getTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), 3, False)
    End With
End Function






Attribute VB_Name = "InterestsMgr"
Public Const INTEREST_CALC_SHEET As String = "Calculator"

Private Const BALANCE_HISTORY_TABLE As String = "TableBalanceHistory"
Private Const DEPOSITS_HISTORY_TABLE As String = "TableDepositHistory"

Private Const INTEREST_DATE_START_CELL As String = "B2"
Private Const INTEREST_DATE_STOP_CELL As String = "B3"
Private Const INTEREST_GOAL_SEEK_CELL As String = "B7"
Private Const INTEREST_RATE_CELL As String = "B8"
Private Const BALANCE_END_CELL As String = "C3"

Private Const DATE_COL = 1
Private Const BALANCE_COL = 2
Private Const INTEREST_COL = 3

Public Function InterestsCalc(balanceArray As Variant, depositsArray As Variant, Optional account As String = "account", Optional calcPerPeriod As Boolean = True)
    FreezeDisplay
    Call InterestsLoadData(balanceArray, depositsArray, account)
    InterestsCalc = InterestsCalcFromData(calcPerPeriod)
    UnfreezeDisplay
End Function


Private Sub InterestsLoadData(balancesArray As Variant, depositsArray As Variant, Optional accName As String = "account")
    ' Loads data need for interests calculation in the calculation sheet
    With Sheets(INTEREST_CALC_SHEET)
        .Range("B1").Value = accName
        .Range("A10").Value = "Deposit history for " & accName
        .Range("F10").Value = "Balance history for " & accName

        Call ResizeTable(.ListObjects(BALANCE_HISTORY_TABLE), UBound(balancesArray, 1))
        Call ResizeTable(.ListObjects(DEPOSITS_HISTORY_TABLE), UBound(depositsArray, 1))
        
        ' Copy 2 first columns of the 2 tables with history of deposits (date/amount) and history of balance (date/amount)
        Call SetTableColumn(.ListObjects(DEPOSITS_HISTORY_TABLE), 1, GetArrayColumn(depositsArray, 1, False))
        Call SetTableColumn(.ListObjects(DEPOSITS_HISTORY_TABLE), 2, GetArrayColumn(depositsArray, 2, False))
        Call SetTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), DATE_COL, GetArrayColumn(balancesArray, 1, False))
        Call SetTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), BALANCE_COL, GetArrayColumn(balancesArray, 2, False))
        '.ListObjects(2).ListColumns(3).DataBodyRange.Cells(1).formula = "=IF(OR([Date]>target_date,[Date]<=start_date),0,FLOOR((target_date-[Date])/15.2,1))"
        '.ListObjects(2).ListColumns(4).DataBodyRange.Cells(1).formula = "=IF([Nbr de périodes]<=0;IF(OR([Date]>=target_date;[Date]<=start_date);0;[Montant]);[Montant]*(1+$R$1)^[Nbr de périodes])"
        
        ' Clear old calculated interest rates
        Call ClearTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), INTEREST_COL)
    End With
End Sub


Private Function InterestsCalcFromData(Optional calcPerPeriod As Boolean = True)
    ' Calculate interest rates for each delta period or since the beginning of the balance history sheet
    With Sheets(INTEREST_CALC_SHEET)
        For i = 2 To .ListObjects(BALANCE_HISTORY_TABLE).ListRows.Count
            If calcPerPeriod Then
                .Range(INTEREST_DATE_START_CELL).Value = .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(DATE_COL).DataBodyRange.Rows(i - 1).Value
            Else
                .Range(INTEREST_DATE_START_CELL).Value = .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(DATE_COL).DataBodyRange.Rows(1).Value
            End If
            .Range(INTEREST_DATE_STOP_CELL).Value = .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(DATE_COL).DataBodyRange.Rows(i).Value
            .Range(INTEREST_RATE_CELL).Value = 0.1
            .Range(INTEREST_GOAL_SEEK_CELL).GoalSeek Goal:=.Range(BALANCE_END_CELL).Value, ChangingCell:=.Range(INTEREST_RATE_CELL)
            .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(INTEREST_COL).DataBodyRange.Rows(i).Value = .Range(INTEREST_RATE_CELL).Value
        Next i
        InterestsCalcFromData = GetTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), INTEREST_COL)
    End With
End Function






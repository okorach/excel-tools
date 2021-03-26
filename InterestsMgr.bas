Attribute VB_Name = "InterestsMgr"
Public Const INTEREST_CALC_SHEET As String = "Interests"

Private Const BALANCE_HISTORY_TABLE As String = "TableBalanceHistory"
Private Const DEPOSITS_HISTORY_TABLE As String = "TableDepositHistory"
Private Const INTEREST_TABLE As String = "AccountsInterests"

Private Const INTEREST_DATE_START_CELL As String = "I2"
Private Const INTEREST_DATE_STOP_CELL As String = "I3"
Private Const INTEREST_PERIOD_CELL As String = "I4"
Private Const INTEREST_GOAL_SEEK_CELL As String = "I8"
Private Const INTEREST_RATE_CELL As String = "I9"
Private Const BALANCE_END_CELL As String = "J3"

Private Const DATE_COL = 1
Private Const BALANCE_COL = 2
Private Const INTEREST_COL = 3

Public Function InterestsCalc(balanceArray As Variant, depositsArray As Variant, Optional account As String = "account", _
                              Optional interestPeriod As Integer = 1, Optional calcPerPeriod As Boolean = True)
    FreezeDisplay
    Call interestsLoadData(balanceArray, depositsArray, account, interestPeriod)
    InterestsCalc = interestsCalcFromData(calcPerPeriod)
    UnfreezeDisplay
End Function

Public Sub InterestsStore(ByVal accountId As String, ByVal thisYear As Variant, ByVal lastYear As Variant, ByVal last3years As Variant, ByVal last5years As Variant, ByVal allTime As Variant)
    Dim oTable As ListObject
    Set oTable = Sheets(INTEREST_CALC_SHEET).ListObjects(INTEREST_TABLE)
    Dim row As ListRow
    For Each row In oTable.ListRows
        If row.Range.Cells(1, 1).Value = accountId Then
            Call interestsRecord(row, thisYear, lastYear, last3years, last5years, allTime)
            Exit Sub
        End If
    Next row
    oTable.ListRows.Add
    oTable.ListRows(oTable.ListRows.Count).Range.Cells(1, 1).Value = accountId
    Call interestsRecord(oTable.ListRows(oTable.ListRows.Count), thisYear, lastYear, last3years, last5years, allTime)
End Sub

Private Sub interestsLoadData(balancesArray As Variant, depositsArray As Variant, Optional accName As String = "account", Optional interestPeriod As Integer = 1)
    ' Loads data need for interests calculation in the calculation sheet
    With Sheets(INTEREST_CALC_SHEET)
        .Range("I1").Value = accName
        .Range("H11").Value = "Deposit history for " & accName
        .Range("M11").Value = "Balance history for " & accName

        .Range(INTEREST_PERIOD_CELL).Value = interestPeriod
        
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


Private Function interestsCalcFromData(Optional calcPerPeriod As Boolean = True)
    ' Calculate interest rates for each delta period or since the beginning of the balance history sheet
    With Sheets(INTEREST_CALC_SHEET)
        For i = 2 To .ListObjects(BALANCE_HISTORY_TABLE).ListRows.Count
            If calcPerPeriod Then
                .Range(INTEREST_DATE_START_CELL).Value = .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(DATE_COL).DataBodyRange.Rows(i - 1).Value
            Else
                .Range(INTEREST_DATE_START_CELL).Value = .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(DATE_COL).DataBodyRange.Rows(1).Value
            End If
            .Range(INTEREST_DATE_STOP_CELL).Value = .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(DATE_COL).DataBodyRange.Rows(i).Value
            .Range(INTEREST_RATE_CELL).Value = 0
            .Range(INTEREST_GOAL_SEEK_CELL).GoalSeek Goal:=.Range(BALANCE_END_CELL).Value, ChangingCell:=.Range(INTEREST_RATE_CELL)
            .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(INTEREST_COL).DataBodyRange.Rows(i).Value = .Range(INTEREST_RATE_CELL).Value
        Next i
        interestsCalcFromData = GetTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), INTEREST_COL)
    End With
End Function

Private Sub interestsRecord(row As ListRow, thisYear As Variant, lastYear As Variant, last3years As Variant, last5years As Variant, allTime As Variant)
    row.Range.Cells(1, 2).Value = thisYear
    row.Range.Cells(1, 3).Value = lastYear
    row.Range.Cells(1, 4).Value = last3years
    row.Range.Cells(1, 5).Value = last5years
    row.Range.Cells(1, 6).Value = allTime
End Sub


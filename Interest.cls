VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Interest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' List of 3 tables of the unique "Interests" sheet
Private Const INTEREST_TABLE As String = "AccountsInterests"
Private Const BALANCE_HISTORY_TABLE As String = "TableBalanceHistory"
Private Const DEPOSITS_HISTORY_TABLE As String = "TableDepositHistory"

' Constants defining structure of global interests table
Private Const INTERESTS_ACCOUNT_COL As Integer = 1
Private Const INTERESTS_THIS_YEAR_COL As Integer = 2
Private Const INTERESTS_LAST_YEAR_COL As Integer = 3
Private Const INTERESTS_LAST_3_YEARS_COL As Integer = 4
Private Const INTERESTS_LAST_5_YEARS_COL As Integer = 5
Private Const INTERESTS_ALL_TIME_COL As Integer = 6

' constants defining structure of global interests and deposits table
Private Const I_DEPOSITS_DATE_COL = 1
Private Const I_DEPOSITS_AMOUNT_COL = 2
Private Const I_INT_DATE_COL = 1
Private Const I_INT_BALANCE_COL = 2
Private Const I_INT_INTERESTS_COL = 3


Private Const INTEREST_DATE_START_CELL As String = "J2"
Private Const INTEREST_DATE_STOP_CELL As String = "J3"
Private Const INTEREST_PERIOD_CELL As String = "J4"
Private Const INTEREST_GOAL_SEEK_CELL As String = "J8"
Private Const INTEREST_RATE_CELL As String = "J9"
Private Const BALANCE_END_CELL As String = "K3"
Private Const INTEREST_TITLE_CELL = "I1"
Private Const DEPOSITS_TITLE_CELL = "I11"
Private Const BALANCE_TITLE_CELL = "N11"

Private accountId As String
Private balanceArr As Variant
Private depositArr As Variant
Private calculatedInterests As Variant
Private keyInterests(1 To 5) As Variant
Private InterestPeriod As Integer


Private Sub Class_Initialize()
    accountId = vbNullString
End Sub

Public Sub Init(accId As String, Optional balancesArray As Variant = Nothing, Optional depositsArray As Variant = Nothing, _
                Optional period As Integer = 0)
    accountId = accId
    balanceArr = balancesArray
    depositArr = depositsArray
    calculatedInterests = vbNull
    For i = 1 To 5
        keyInterests(i) = vbNull
    Next i
    If period = 0 Then
        period = 1
    End If
    InterestPeriod = period
End Sub

Public Function Calc(Optional withModal As Boolean = True) As Variant
    Dim calcPerPeriod As Boolean
    calcPerPeriod = True
    Dim mymodal As ProgressBar
    
    With Sheets(INTEREST_CALC_SHEET).ListObjects(BALANCE_HISTORY_TABLE)
        If withModal Then
            Set mymodal = NewProgressBar("Interest calculation", .ListRows.Count)
        End If
        Call prepareForCalculation
        If withModal Then
            mymodal.Update
        End If
    
        ' Do calculations
        For i = 2 To .ListRows.Count
            Dim startDate As Date
            If calcPerPeriod Then
                startDate = .ListColumns(I_INT_DATE_COL).DataBodyRange.Rows(i - 1).value
            Else
                startDate = .ListColumns(I_INT_DATE_COL).DataBodyRange.Rows(1).value
            End If
            .ListColumns(I_INT_INTERESTS_COL).DataBodyRange.Rows(i).value = _
                CalcOnPeriod(startDate, .ListColumns(I_INT_DATE_COL).DataBodyRange.Rows(i).value)
            If withModal Then
                mymodal.Update
            End If
        Next i
    End With
    calculatedInterests = GetTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects(BALANCE_HISTORY_TABLE), I_INT_INTERESTS_COL)
    Calc = calculatedInterests
    If withModal Then
        Set mymodal = Nothing
    End If
End Function

Public Function CalcKeyInterests(Optional modal As ProgressBar = Nothing) As Variant
    If modal Is Nothing Then
        Set mymodal = NewProgressBar("Interest calculation", 6)
    Else
        Set mymodal = modal
    End If
    Call prepareForCalculation
    mymodal.UpdateMessage ("Interest calculation for " & accountId)
    mymodal.Update

    nbrYears = Sheets(INTEREST_CALC_SHEET).ListObjects(BALANCE_HISTORY_TABLE).ListRows.Count - 1
    ' Do calculations
    Dim thisYear As Integer
    thisYear = Year(Now())
    Dim lastBalanceDate As Date
    lastBalanceDate = CDate(balanceArr(UBound(balanceArr, 1), 1))
    keyInterests(1) = CalcOnPeriod(DateSerial(thisYear - 1, 12, 31), lastBalanceDate)
    mymodal.Update
    If nbrYears > 1 Then
        keyInterests(2) = CalcOnPeriod(DateSerial(thisYear - 2, 12, 31), DateSerial(thisYear - 1, 12, 31))
    Else
        keyInterests(2) = vbNull
    End If
    mymodal.Update
    If nbrYears > 3 Then
        keyInterests(3) = CalcOnPeriod(DateSerial(thisYear - 3, 12, 31), lastBalanceDate)
    Else
        keyInterests(3) = vbNull
    End If
    mymodal.Update
    If nbrYears > 5 Then
        keyInterests(4) = CalcOnPeriod(DateSerial(thisYear - 5, 12, 31), lastBalanceDate)
    Else
        keyInterests(4) = vbNull
    End If
    mymodal.Update
    keyInterests(5) = CalcOnPeriod(CDate(depositArr(1, 1) - 1), lastBalanceDate)
    mymodal.Update
    If modal Is Nothing Then
        Set mymodal = Nothing
    End If
    CalcKeyInterests = keyInterests
End Function

Public Function CalcOnPeriod(startDate As Date, endDate As Date) As Double
    If startDate >= endDate Then
        CalcOnPeriod = vbNull
        Exit Function
    End If
    
    With Sheets(INTEREST_CALC_SHEET)
        .Range(INTEREST_DATE_START_CELL).value = startDate
        .Range(INTEREST_DATE_STOP_CELL).value = endDate
        If .Range(BALANCE_END_CELL).value = vbNullString Then
            CalcOnPeriod = vbNull
            Exit Function
        End If
        .Range(INTEREST_RATE_CELL).value = 0
        .Range(INTEREST_GOAL_SEEK_CELL).GoalSeek goal:=.Range(BALANCE_END_CELL).value, ChangingCell:=.Range(INTEREST_RATE_CELL)
        CalcOnPeriod = .Range(INTEREST_RATE_CELL).value
    End With
End Function

Public Function Store(resultTable As ListObject, globalKey As String, Optional TaxRate As Double = 0) As Boolean
    Dim periodicInterests As Variant
    Dim row As ListRow
    periodicInterests = calcPerPeriod()
    With resultTable
        For i = 1 To 5
            .ListRows(i).Range(1, IT_GROSS_INTEREST_COL).value = periodicInterests(i)
            If VarType(periodicInterests(i)) = vbDouble And periodicInterests(i) <> vbNull Then
                .ListRows(i).Range(1, IT_NET_INTEREST_COL).value = periodicInterests(i) * (1 - TaxRate)
            End If
        Next i
    End With

    ' Compute net interest before storing globally
    For i = 1 To 5
        If VarType(periodicInterests(i)) = vbDouble And periodicInterests(i) <> vbNull Then
            periodicInterests(i) = periodicInterests(i) * (1 - TaxRate)
        Else
            periodicInterests(i) = "-"
        End If
    Next i
    With Sheets(INTEREST_CALC_SHEET).ListObjects(INTEREST_TABLE)
        For Each row In .ListRows
            If row.Range.Cells(1, INTERESTS_ACCOUNT_COL).value = globalKey Then
                Call saveData(row, periodicInterests)
                Exit Function
            End If
        Next row
        .ListRows.Add
        .ListRows(.ListRows.Count).Range.Cells(1, INTERESTS_ACCOUNT_COL).value = accountId
        Call saveData(.ListRows(oTable.ListRows.Count), periodicInterest)
    End With
    Store = True
End Function

Public Function StoreKeyInterests(resultTable As ListObject, globalKey As String, Optional TaxRate As Double = 0) As Boolean
    With resultTable
        For i = 1 To 5
            If VarType(keyInterests(i)) = vbDouble And keyInterests(i) <> vbNull Then
                .ListRows(i).Range(1, IT_GROSS_INTEREST_COL).value = keyInterests(i)
                .ListRows(i).Range(1, IT_NET_INTEREST_COL).value = keyInterests(i) * (1 - TaxRate)
                keyInterests(i) = keyInterests(i) * (1 - TaxRate)
            Else
                .ListRows(i).Range(1, IT_GROSS_INTEREST_COL).value = "-"
                .ListRows(i).Range(1, IT_NET_INTEREST_COL).value = "-"
                keyInterests(i) = "-"
            End If
        Next i
    End With
    
    ' Store net interests
    Dim row As ListRow
    With Sheets(INTEREST_CALC_SHEET).ListObjects(INTEREST_TABLE)
        For Each row In .ListRows
            If row.Range.Cells(1, INTERESTS_ACCOUNT_COL).value = globalKey Then
                Call saveData(row, keyInterests)
                Exit Function
            End If
        Next row
        .ListRows.Add
        .ListRows(.ListRows.Count).Range.Cells(1, INTERESTS_ACCOUNT_COL).value = accountId
        Call saveData(.ListRows(oTable.ListRows.Count), keyInterests)
    End With
    StoreKeyInterests = True
End Function

Private Function calcPerPeriod() As Variant
    Dim nbrYears As Long
    nbrYears = UBound(calculatedInterests)
    Dim periodicInterest(1 To 5) As Variant
    For i = 1 To 5
        periodicInterest(i) = "-"
    Next i
    periodicInterest(1) = calculatedInterests(nbrYears)
    If nbrYears >= 3 Then
        periodicInterest(2) = calculatedInterests(nbrYears - 1)
        periodicInterest(5) = ArrayAverage(calculatedInterests, 2, nbrYears - 1)
    End If
    If nbrYears >= 5 Then
        periodicInterest(3) = ArrayAverage(calculatedInterests, nbrYears - 2, nbrYears)
    End If
    If nbrYears >= 7 Then
        periodicInterest(4) = ArrayAverage(calculatedInterests, nbrYears - 4, nbrYears)
    End If
    calcPerPeriod = periodicInterest
End Function


Private Sub saveData(row As ListRow, netInterests As Variant)
    row.Range.Cells(1, 2).value = Now
    For i = 1 To 5
        row.Range.Cells(1, i + 2).value = netInterests(i)
    Next i
End Sub

Private Sub prepareForCalculation()
    ' Dumps the interest object data on the interest calc sheet to prepare interests calculation
    With Sheets(INTEREST_CALC_SHEET)
        ' Load data to prepare calculation
        .Range(INTEREST_TITLE_CELL).value = accountId
        .Range(DEPOSITS_TITLE_CELL).value = "Deposit history for " & accountId
        .Range(BALANCE_TITLE_CELL).value = "Balance history for " & accountId
        .Range(INTEREST_PERIOD_CELL).value = InterestPeriod
        
        Call ResizeTable(.ListObjects(BALANCE_HISTORY_TABLE), UBound(balanceArr, 1))
        Call ResizeTable(.ListObjects(DEPOSITS_HISTORY_TABLE), UBound(depositArr, 1))
        
        ' Copy 2 first columns of the 2 tables with history of deposits (date/amount) and history of balance (date/amount)
        Call SetTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), I_INT_DATE_COL, GetArrayColumn(balanceArr, 1, False))
        Call SetTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), I_INT_BALANCE_COL, GetArrayColumn(balanceArr, 2, False))
        Call SetTableColumn(.ListObjects(DEPOSITS_HISTORY_TABLE), I_DEPOSITS_DATE_COL, GetArrayColumn(depositArr, 1, False))
        Call SetTableColumn(.ListObjects(DEPOSITS_HISTORY_TABLE), I_DEPOSITS_AMOUNT_COL, GetArrayColumn(depositArr, 2, False))
        ' Call SetTableFromArray(.ListObjects(DEPOSITS_HISTORY_TABLE), depositArr)
        
        ' Clear old calculated interest rates
        Call ClearTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), I_INT_INTERESTS_COL)
    End With
End Sub

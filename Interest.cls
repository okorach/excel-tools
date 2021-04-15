VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Interest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const INTEREST_DATE_START_CELL As String = "I2"
Private Const INTEREST_DATE_STOP_CELL As String = "I3"
Private Const INTEREST_PERIOD_CELL As String = "I4"
Private Const INTEREST_GOAL_SEEK_CELL As String = "I8"
Private Const INTEREST_RATE_CELL As String = "I9"
Private Const BALANCE_END_CELL As String = "J3"

Private Const DATE_COL = 1
Private Const BALANCE_COL = 2
Private Const INTEREST_COL = 3

Private accountId As String
Private balanceArr As Variant
Private depositArr As Variant
Private calculatedInterests As Variant
Private InterestPeriod As Integer


Private Sub Class_Initialize()
    accountId = vbNullString
End Sub

Public Sub Init(accId As String, Optional balancesArray As Variant = Nothing, Optional depositsArray As Variant = Nothing, Optional period As Integer = 0)
    accountId = accId
    balanceArr = balancesArray
    depositArr = depositsArray
    calculatedInterests = vbNull
    If period = 0 Then
        period = 1
    End If
    InterestPeriod = period
End Sub

Public Function Calc(Optional modal As ProgressBar = Nothing) As Variant
    Dim calcPerPeriod As Boolean
    calcPerPeriod = True
    With Sheets(INTEREST_CALC_SHEET)
        .Range("I1").value = accountId
        .Range("H11").value = "Deposit history for " & accountId
        .Range("M11").value = "Balance history for " & accountId
        .Range(INTEREST_PERIOD_CELL).value = InterestPeriod
        
        Call ResizeTable(.ListObjects(BALANCE_HISTORY_TABLE), UBound(balanceArr, 1))
        Call ResizeTable(.ListObjects(DEPOSITS_HISTORY_TABLE), UBound(depositArr, 1))
        
        ' Copy 2 first columns of the 2 tables with history of deposits (date/amount) and history of balance (date/amount)
        Call SetTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), DATE_COL, GetArrayColumn(balanceArr, 1, False))
        Call SetTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), BALANCE_COL, GetArrayColumn(balanceArr, 2, False))
        Call SetTableColumn(.ListObjects(DEPOSITS_HISTORY_TABLE), 1, GetArrayColumn(depositArr, 1, False))
        Call SetTableColumn(.ListObjects(DEPOSITS_HISTORY_TABLE), 2, GetArrayColumn(depositArr, 2, False))

        ' Clear old calculated interest rates
        Call ClearTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), INTEREST_COL)
        loadData
        
        If modal Is Nothing Then
            Set mymodal = NewProgressBar("Interest calculation", .ListObjects(BALANCE_HISTORY_TABLE).ListRows.Count - 1)
        Else
            Set mymodal = modal
        End If
        For i = 2 To .ListObjects(BALANCE_HISTORY_TABLE).ListRows.Count
            If calcPerPeriod Then
                .Range(INTEREST_DATE_START_CELL).value = .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(DATE_COL).DataBodyRange.Rows(i - 1).value
            Else
                .Range(INTEREST_DATE_START_CELL).value = .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(DATE_COL).DataBodyRange.Rows(1).value
            End If
            .Range(INTEREST_DATE_STOP_CELL).value = .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(DATE_COL).DataBodyRange.Rows(i).value
            .Range(INTEREST_RATE_CELL).value = 0
            .Range(INTEREST_GOAL_SEEK_CELL).GoalSeek goal:=.Range(BALANCE_END_CELL).value, ChangingCell:=.Range(INTEREST_RATE_CELL)
            .ListObjects(BALANCE_HISTORY_TABLE).ListColumns(INTEREST_COL).DataBodyRange.Rows(i).value = .Range(INTEREST_RATE_CELL).value
            mymodal.Update
        Next i
        calculatedInterests = GetTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), INTEREST_COL)
        Calc = calculatedInterests
    End With
    If modal Is Nothing Then
        Set mymodal = Nothing
    End If
End Function

Public Function Store(resultTable As ListObject, globalKey As String, Optional TaxRate As Double = 0) As Boolean
    Dim periodicInterests As Variant
    Dim row As ListRow
    periodicInterests = calcPerPeriod()
    With resultTable
        For i = 1 To 5
            .ListRows(i).Range(1, 2).value = periodicInterests(i)
            If VarType(periodicInterests(i)) = vbDouble Then
                .ListRows(i).Range(1, 3).value = periodicInterests(i) * (1 - TaxRate)
            End If
        Next i
    End With

    ' Compute net interest before storing globally
    For i = 1 To 5
        If VarType(periodicInterests(i)) = vbDouble Then
            periodicInterests(i) = periodicInterests(i) * (1 - TaxRate)
        End If
    Next i
    With Sheets(INTEREST_CALC_SHEET).ListObjects(INTEREST_TABLE)
        For Each row In .ListRows
            If row.Range.Cells(1, 1).value = globalKey Then
                Call saveData(row, periodicInterests)
                Exit Function
            End If
        Next row
        .ListRows.Add
        .ListRows(.ListRows.Count).Range.Cells(1, 1).value = accountId
        Call saveData(.ListRows(oTable.ListRows.Count), periodicInterest)
    End With
    Store = True
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
        periodicInterest(3) = ArrayAverage(calculatedInterests, nbrYears - 3, nbrYears - 1)
    End If
    If nbrYears >= 7 Then
        periodicInterest(4) = ArrayAverage(calculatedInterests, nbrYears - 5, nbrYears - 1)
    End If
    calcPerPeriod = periodicInterest
End Function

Private Sub loadData()
    ' Loads data need for interests calculation in the calculation sheet
    With Sheets(INTEREST_CALC_SHEET)
        .Range("I1").value = accountId
        .Range("H11").value = "Deposit history for " & accountId
        .Range("M11").value = "Balance history for " & accountId
        .Range(INTEREST_PERIOD_CELL).value = InterestPeriod
        
        Call ResizeTable(.ListObjects(BALANCE_HISTORY_TABLE), UBound(balanceArr, 1))
        Call ResizeTable(.ListObjects(DEPOSITS_HISTORY_TABLE), UBound(depositArr, 1))
        
        ' Copy 2 first columns of the 2 tables with history of deposits (date/amount) and history of balance (date/amount)
        Call SetTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), DATE_COL, GetArrayColumn(balanceArr, 1, False))
        Call SetTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), BALANCE_COL, GetArrayColumn(balanceArr, 2, False))
        Call SetTableColumn(.ListObjects(DEPOSITS_HISTORY_TABLE), 1, GetArrayColumn(depositArr, 1, False))
        Call SetTableColumn(.ListObjects(DEPOSITS_HISTORY_TABLE), 2, GetArrayColumn(depositArr, 2, False))

        ' Clear old calculated interest rates
        Call ClearTableColumn(.ListObjects(BALANCE_HISTORY_TABLE), INTEREST_COL)
    End With
End Sub

Private Sub saveData(row As ListRow, netInterests As Variant)
    For i = 1 To 5
        row.Range.Cells(1, i + 2).value = netInterests(i)
    Next i
End Sub


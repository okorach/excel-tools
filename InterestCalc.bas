Attribute VB_Name = "InterestCalc"
Public Const INTEREST_CALC_SHEET As String = "Calculator"

Sub CalcInterestForAllAccounts()

    freezeDisplay
    
    For i = 1 To Sheets.Count
        If (Sheets(i).name <> INTEREST_CALC_SHEET And Sheets(i).name <> "Params" And Sheets(i).name <> "Summary") Then
            Call CalcInterestForAccount(Sheets(i).name)
        End If
    Next i
    
    unfreezeDisplay
    
End Sub

Sub CalcInterestForAccount(name As String)
    Call ImportAccountName(name)
    Call CalcAllInterests
    Call ExportAccountName(name)
End Sub

Sub ImportAccount()
    accNbr = Sheets(INTEREST_CALC_SHEET).Range("B1").Value
    Call ImportAccountName(Sheets("Params").Range("E" + CStr(accNbr)).Value)
End Sub

Sub ExportAccount()
    Call ExportAccountName(getSelectedAccount())
End Sub

Sub ExportAccountName(accName As String)
    Call ExportInterestResults(accName)
End Sub

Sub ImportAccountName(accName As String)

    freezeDisplay
    Sheets(INTEREST_CALC_SHEET).Range("G1").Value = "Deposit history for " & accName
    Sheets(INTEREST_CALC_SHEET).Range("L1").Value = "Balance history for " & accName
    
    Call resizeTable(Sheets(INTEREST_CALC_SHEET).ListObjects(1), Sheets(accName).ListObjects(1).ListRows.Count)
    Call resizeTable(Sheets(INTEREST_CALC_SHEET).ListObjects(2), Sheets(accName).ListObjects(2).ListRows.Count)
    
    Sheets(accName).ListObjects(1).name = "TableBalance" & Replace(accName, " ", "")
    Sheets(accName).ListObjects(2).name = "TableDeposit" & Replace(accName, " ", "")
    
    ' Copy 2 first colmuns of the 2 tables with history of deposits (date/amount) and history of balance (date/amount)
    Call setTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects(1), getTableColumn(Sheets(accName).ListObjects(1), 1), 1)
    Call setTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects(1), getTableColumn(Sheets(accName).ListObjects(1), 2), 2)
    Call setTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects(2), getTableColumn(Sheets(accName).ListObjects(2), 1), 1)
    Call setTableColumn(Sheets(INTEREST_CALC_SHEET).ListObjects(2), getTableColumn(Sheets(accName).ListObjects(2), 2), 2)
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

    For i = 2 To Sheets(INTEREST_CALC_SHEET).ListObjects("TableBalanceHistory").ListRows.Count
        Sheets(INTEREST_CALC_SHEET).Range("B2").Value = Sheets(INTEREST_CALC_SHEET).ListObjects("TableBalanceHistory").ListColumns(1).DataBodyRange.Rows(1).Value
        Sheets(INTEREST_CALC_SHEET).Range("B3").Value = Sheets(INTEREST_CALC_SHEET).ListObjects("TableBalanceHistory").ListColumns(1).DataBodyRange.Rows(i).Value
        Sheets(INTEREST_CALC_SHEET).Range("B4").GoalSeek Goal:=Sheets("Calculator").Range("C3").Value, ChangingCell:=Range("B5")
        Sheets(INTEREST_CALC_SHEET).ListObjects("TableBalanceHistory").ListColumns(4).DataBodyRange.Rows(i).Value = Sheets("Calculator").Range("B5").Value
    Next i
End Sub

Sub CalcPeriodicInterests()

    For i = 2 To Sheets(INTEREST_CALC_SHEET).ListObjects("TableBalanceHistory").ListRows.Count
        Sheets(INTEREST_CALC_SHEET).Range("B2").Value = Sheets(INTEREST_CALC_SHEET).ListObjects("TableBalanceHistory").ListColumns(1).DataBodyRange.Rows(i - 1).Value
        Sheets(INTEREST_CALC_SHEET).Range("B3").Value = Sheets(INTEREST_CALC_SHEET).ListObjects("TableBalanceHistory").ListColumns(1).DataBodyRange.Rows(i).Value
        Sheets(INTEREST_CALC_SHEET).Range("B5").Value = 0.1
        Sheets(INTEREST_CALC_SHEET).Range("B4").GoalSeek Goal:=Sheets(INTEREST_CALC_SHEET).Range("C3").Value, ChangingCell:=Range("B5")
        Sheets(INTEREST_CALC_SHEET).ListObjects("TableBalanceHistory").ListColumns(3).DataBodyRange.Rows(i).Value = Sheets(INTEREST_CALC_SHEET).Range("B5").Value
    Next i
End Sub

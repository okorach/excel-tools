Attribute VB_Name = "BankImportLCL"
'------------------------------------------------------------------------------
' Import LCL
'------------------------------------------------------------------------------
Private Const LCL_CSV_DATE_FIELD = 1
Private Const LCL_CSV_AMOUNT_FIELD = 2
Private Const LCL_CSV_TYPE_FIELD = 3
Private Const LCL_CSV_DESC1_FIELD = 4
Private Const LCL_CSV_DESC2_FIELD = 5
Private Const LCL_CSV_DESC3_FIELD = 6

Public Sub ImportLCL(oTable As ListObject, fileToOpen As Variant, dateCol As Integer, amountCol As Integer, descCol As Integer)

    subsTable = GetTableAsArray(Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True, Format:=6, delimiter:=";"
    Dim modal As ProgressBar
    Set modal = NewProgressBar("Import LCL in progress", GetLastNonEmptyRow() + 1)
    modal.Update
    
    Dim csvIsSplit As Boolean
    csvIsSplit = Not (Cells(1, 2).value = "")
    Dim r As Long
    r = 1
    Do While LenB(Cells(r + 1, 1).value) > 0
        oTable.ListRows.Add
        With oTable.ListRows(oTable.ListRows.Count)
            If csvIsSplit Then
                ' semicolon CSV cell separator did work
                rawDate = DateValue(Cells(r, LCL_CSV_DATE_FIELD).value)
                rawAmount = Cells(r, LCL_CSV_AMOUNT_FIELD).value
                If (Cells(r, LCL_CSV_TYPE_FIELD).value Like "Ch?que") Then
                    rawDesc = "Cheque " & simplifyDescription(CStr(Cells(r, LCL_CSV_DESC1_FIELD).value), subsTable)
                ElseIf (Cells(r, LCL_CSV_TYPE_FIELD).value = "Virement") Then
                    .Range(1, descCol).value = "Virement " & simplifyDescription(Cells(r, LCL_CSV_DESC2_FIELD).value, subsTable)
                Else
                    .Range(1, descCol).value = simplifyDescription(Cells(r, LCL_CSV_TYPE_FIELD).value & " " & Cells(r, LCL_CSV_DESC2_FIELD).value & " " & Cells(r, LCL_CSV_DESC3_FIELD).value, subsTable)
                End If
            Else
                ' semicolon CSV cell separator did not work
                Dim a() As String
                a = Split(Cells(r, 1).value, ";", -1, vbTextCompare)
                rawDate = ToDate(Trim$(a(LCL_CSV_DATE_FIELD - 1)))
                rawAmount = toAmount(Trim$(a(LCL_CSV_AMOUNT_FIELD - 1)))
                Dim des As String
                If (a(LCL_CSV_TYPE_FIELD - 1) Like "Ch?que") Then
                    rawDesc = "Cheque " & simplifyDescription(a(LCL_CSV_DESC1_FIELD - 1), subsTable)
                ElseIf (a(LCL_CSV_TYPE_FIELD - 1) = "Virement") Then
                    rawDesc = "Virement " & simplifyDescription(a(LCL_CSV_DESC2_FIELD - 1), subsTable)
                Else
                    rawDesc = simplifyDescription(a(LCL_CSV_TYPE_FIELD - 1) & " " & a(LCL_CSV_DESC2_FIELD - 1) & " " & a(LCL_CSV_DESC3_FIELD - 1), subsTable)
                End If

            End If
            .Range(1, dateCol).value = rawDate
            .Range(1, amountCol).value = toAmount(rawAmount)
            .Range(1, descCol).value = rawDesc
        End With
        r = r + 1
        modal.Update
    Loop
    ActiveWorkbook.Close
    Set modal = Nothing
End Sub

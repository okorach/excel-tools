Attribute VB_Name = "AccountImportExport"
Private Function toAmount(str) As Double
    If VarType(str) = vbString Then
        str = Replace(Replace(str, ",", "."), "'", "")
        toAmount = CDbl(str)
    Else
        toAmount = str
    End If
End Function

Private Function toMonth(str) As Integer
    s = LCase$(Trim$(str))
    If s Like "jan*" Then
        toMonth = 1
    ElseIf s Like "fe*" Or s Like "fé*" Then
        toMonth = 2
    ElseIf s Like "mar*" Then
        toMonth = 3
    ElseIf s Like "a*" Then
        toMonth = 4
    ElseIf s Like "mai*" Or s Like "may*" Then
        toMonth = 5
    ElseIf s Like "juin*" Or s Like "jun*" Then
        toMonth = 6
    ElseIf s Like "juil*" Or s Like "jul*" Then
        toMonth = 7
    ElseIf s Like "ao*" Or s Like "aug*" Then
        toMonth = 8
    ElseIf s Like "sep*" Then
        toMonth = 9
    ElseIf s Like "oct*" Then
        toMonth = 10
    ElseIf s Like "nov*" Then
        toMonth = 11
    ElseIf s Like "dec*" Or s Like "déc*" Then
        toMonth = 12
    Else
        toMonth = 0
    End If
End Function

Private Function toDate(str) As Date
    A = Split(str, " ", -1, vbTextCompare)
    toDate = DateSerial(CInt(A(2)), toMonth(A(1)), CInt(A(0)))
End Function

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportAny()

    Dim fileToOpen As Variant
    fileToOpen = Application.GetOpenFilename()
    If fileToOpen <> False Then
        Dim bank As String
        bank = Cells(3, 2).Value
        If (bank = "ING Direct") Then
            Call ImportING(fileToOpen)
        ElseIf (bank = "LCL") Then
            Call ImportLCL(fileToOpen)
        ElseIf (bank = "UBS") Then
            Call ImportUBS(fileToOpen)
        ElseIf (bank = "Revolut") Then
            Call ImportRevolut(fileToOpen)
        Else
            MsgBox ("Format d'import (banque) non identifiable, opération annulée")
        End If
    Else
        MsgBox ("Import annulé")
    End If
End Sub
Sub ImportING(fileToOpen As Variant)

Workbooks.Open filename:=fileToOpen, ReadOnly:=True
Dim iRow As Integer
Dim tDates() As Variant
Dim tDesc() As String
Dim tValues() As Double

ReDim tDates(1 To 30000)
ReDim tDesc(1 To 30000)
ReDim tValues(1 To 30000)
iRow = 1
Do While LenB(Cells(iRow, 1).Value) > 0 And iRow < 30000
    iRow = iRow + 1
Loop
nbRows = iRow - 1
iRow = 1
Do While LenB(Cells(iRow, 1).Value) > 0
    tDates(iRow) = Cells(iRow, 1).Value
    tDesc(iRow) = Cells(iRow, 2).Value
    tValues(iRow) = toAmount(Cells(iRow, 4).Value)
    iRow = iRow + 1
Loop
ActiveWorkbook.Close

With ActiveSheet.ListObjects(1)
    totalrows = .ListRows.Count
    For iRow = 1 To nbRows
        .ListRows.Add
        totalrows = totalrows + 1
        .ListColumns(1).DataBodyRange.Rows(totalrows).Value = tDates(iRow)
        .ListColumns(2).DataBodyRange.Rows(totalrows).Value = tValues(iRow)
        .ListColumns(4).DataBodyRange.Rows(totalrows).Value = tDesc(iRow)
    Next iRow
End With

Call sortAccount(ActiveSheet.ListObjects(1))

Range("A" & CStr(totalrows)).Select

End Sub

Sub ImportRevolut(fileToOpen As Variant)

Workbooks.Open filename:=fileToOpen, ReadOnly:=True
Dim iRow As Integer
Dim tDates() As Variant
Dim tDesc() As String
Dim tValues() As Double

ReDim tDates(1 To 30000)
ReDim tDesc(1 To 30000)
ReDim tValues(1 To 30000)
iRow = 2
Do While LenB(Cells(iRow, 1).Value) > 0 And iRow < 30000
    iRow = iRow + 1
Loop
nbRows = iRow - 2
iRow = 2
Do While LenB(Cells(iRow, 1).Value) > 0
    tDates(iRow - 1) = toDate(Trim$(Cells(iRow, 1).Value))
    tDesc(iRow - 1) = ""
    If LenB(Trim$(Cells(iRow, 3).Value)) = 0 Then
        tValues(iRow - 1) = toAmount(Trim$(Cells(iRow, 4).Value))
        If LenB(Trim$(Cells(iRow, 6).Value)) > 0 Then
            tDesc(iRow - 1) = Trim$(Cells(iRow, 6).Value) & " : "
        End If
    Else
        tValues(iRow - 1) = -toAmount(Trim$(Cells(iRow, 3).Value))
        If LenB(Trim$(Cells(iRow, 5).Value)) > 0 Then
            tDesc(iRow - 1) = Trim$(Cells(iRow, 5).Value) & " : "
        End If
    End If
    tDesc(iRow - 1) = tDesc(iRow - 1) & Trim$(Cells(iRow, 2).Value)
    iRow = iRow + 1
Loop
ActiveWorkbook.Close

With ActiveSheet.ListObjects(1)
    totalrows = .ListRows.Count
    For iRow = 1 To nbRows
        .ListRows.Add
        totalrows = totalrows + 1
        .ListColumns(1).DataBodyRange.Rows(totalrows).Value = tDates(iRow)
        .ListColumns(2).DataBodyRange.Rows(totalrows).Value = tValues(iRow)
        .ListColumns(4).DataBodyRange.Rows(totalrows).Value = tDesc(iRow)
    Next iRow
End With

Call sortAccount(ActiveSheet.ListObjects(1))

Range("A" & CStr(totalrows)).Select

End Sub
Sub ImportRevolutCSV(fileToOpen As Variant)

Open fileToOpen For Input As #1

Line Input #1, textline

With ActiveSheet.ListObjects(1)
totalrows = .ListRows.Count
Do Until EOF(1)
    Line Input #1, textline
    A = Split(textline, ";", -1, vbTextCompare)
    .ListRows.Add
    totalrows = totalrows + 1
    .ListColumns(1).DataBodyRange.Rows(totalrows).Value = toDate(Trim$(A(0)))
    If LenB(Trim$(A(2))) = 0 Then
        .ListColumns(2).DataBodyRange.Rows(totalrows).Value = CDbl(Trim$(A(3)))
        .ListColumns(4).DataBodyRange.Rows(totalrows).Value = Trim$(A(1)) & " --> " & Trim$(A(5))
    Else
        .ListColumns(2).DataBodyRange.Rows(totalrows).Value = -CDbl(Trim$(A(2)))
        .ListColumns(4).DataBodyRange.Rows(totalrows).Value = Trim$(A(1)) & " --> " & Trim$(A(4))
    End If
Loop
End With
Close #1
Call sortAccount(ActiveSheet.ListObjects(1))

Range("A" & CStr(totalrows)).Select

End Sub


'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportLCL(fileToOpen As Variant)


Workbooks.Open filename:=fileToOpen, ReadOnly:=True
Dim iRow As Integer
Dim tDates() As Variant
Dim tDesc() As String
Dim tValues() As Double

ReDim tDates(1 To 30000)
ReDim tDesc(1 To 30000)
ReDim tValues(1 To 30000)
iRow = 1
Do While LenB(Cells(iRow, 1).Value) > 0 And iRow < 30000
    iRow = iRow + 1
Loop
nbRows = iRow - 2 ' Last row is a total, don't import it
iRow = 1
Do While LenB(Cells(iRow, 1).Value) > 0
    tDates(iRow) = DateValue(Cells(iRow, 1).Value)
    tValues(iRow) = toAmount(Cells(iRow, 2).Value)
    If (Cells(iRow, 3).Value = "Chèque") Then
        tDesc(iRow) = "Chèque " & CStr(Cells(iRow, 4).Value)
    ElseIf (Cells(iRow, 3).Value = "Virement") Then
        tDesc(iRow) = "Virement " & Cells(iRow, 5).Value
    Else
        tDesc(iRow) = Cells(iRow, 3).Value & " " & Cells(iRow, 5).Value & " " & Cells(iRow, 6).Value
    End If
    iRow = iRow + 1
Loop
ActiveWorkbook.Close

With ActiveSheet.ListObjects(1)
    totalrows = .ListRows.Count
    For iRow = 1 To nbRows
        .ListRows.Add
        totalrows = totalrows + 1
        .ListColumns(1).DataBodyRange.Rows(totalrows).Value = tDates(iRow)
        .ListColumns(2).DataBodyRange.Rows(totalrows).Value = tValues(iRow)
        .ListColumns(4).DataBodyRange.Rows(totalrows).Value = tDesc(iRow)
    Next iRow
End With

Call sortAccount(ActiveSheet.ListObjects(1))
Range("A" & CStr(totalrows)).Select

End Sub

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportUBS(fileToOpen As Variant)

Workbooks.Open filename:=fileToOpen, ReadOnly:=True
Dim iRow As Integer
Dim tDates() As Variant
Dim tDesc() As String
Dim tValues() As Double

ReDim tDates(1 To 30000)
ReDim tDesc(1 To 30000)
ReDim tValues(1 To 30000)
iRow = 1
nbOps = 0
Do While LenB(Cells(iRow, 1).Value) > 0 And iRow < 30000
    iRow = iRow + 1
    If (LenB(Cells(iRow, 20).Value) > 0 Or LenB(Cells(iRow, 19).Value) > 0) Then
        nbOps = nbOps + 1
    End If
Loop
nbRows = iRow - 1
iRow = 2
nbOps = 0
Do While LenB(Cells(iRow, 1).Value) > 0
    If (LenB(Cells(iRow, 20).Value) > 0 Or LenB(Cells(iRow, 19).Value) > 0) Then
        nbOps = nbOps + 1
        If (LenB(Cells(iRow, 19).Value) > 0) Then
            tValues(nbOps) = -toAmount(Cells(iRow, 19).Value) ' Debit column
        Else
            tValues(nbOps) = toAmount(Cells(iRow, 20).Value) ' Credit column
        End If
        tDates(nbOps) = CDate(DateValue(Replace(Cells(iRow, 12).Value, ".", "/")))
        tDesc(nbOps) = Cells(iRow, 13).Value & " " & Cells(iRow, 14).Value & " " & Cells(iRow, 15).Value
    End If
    iRow = iRow + 1
Loop
ActiveWorkbook.Close

With ActiveSheet.ListObjects(1)
    n = .ListRows.Count
    For iRow = 1 To nbOps
        .ListRows.Add
        n = n + 1
        .ListColumns(1).DataBodyRange.Rows(n).Value = tDates(iRow)
        .ListColumns(3).DataBodyRange.Rows(n).Value = tValues(iRow)
        .ListColumns(4).DataBodyRange.Rows(n).Value = tDesc(iRow)
    Next iRow
End With

Call sortAccount(ActiveSheet.ListObjects(1))
Range("A" & CStr(n)).Select

End Sub

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportUBScsv(fileToOpen As Variant)

Workbooks.OpenText filename:="C:\Users\Olivier\Downloads\export.csv", Origin _
    :=65001, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
    xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=True, _
    Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
    Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), _
    Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15 _
    , 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1)), _
    TrailingMinusNumbers:=True
    'ReadOnly:=True

Dim iRow As Integer
Dim tDates() As Variant
Dim tDesc() As String
Dim tValues() As Double

ReDim tDates(1 To 30000)
ReDim tDesc(1 To 30000)
ReDim tValues(1 To 30000)
iRow = 1
nbOps = 0
Do While LenB(Cells(iRow, 1).Value) > 0 And iRow < 30000
    iRow = iRow + 1
    If LenB(Cells(iRow, 20).Value) > 0 Or LenB(Cells(iRow, 19).Value > 0) Then
        nbOps = nbOps + 1
    End If
Loop
nbRows = iRow - 1
iRow = 2
nbOps = 0
Do While LenB(Cells(iRow, 1).Value) > 0
    If LenB(Cells(iRow, 20).Value) > 0 Or LenB(Cells(iRow, 19).Value) > 0 Then
        nbOps = nbOps + 1
        If LenB(Cells(iRow, 19).Value) > 0 Then
            tValues(nbOps) = -toAmount(Cells(iRow, 19).Value) ' Debit column
        Else
            tValues(nbOps) = toAmount(Cells(iRow, 20).Value) ' Credit column
        End If
        tDates(nbOps) = CDate(DateValue(Replace(Cells(iRow, 12).Value, ".", "/")))
        tDesc(nbOps) = Cells(iRow, 13).Value & " " & Cells(iRow, 14).Value & " " & Cells(iRow, 15).Value
    End If
    iRow = iRow + 1
Loop
ActiveWorkbook.Close

With ActiveSheet.ListObjects(1)
    n = .ListRows.Count
    For iRow = 1 To nbOps
        .ListRows.Add
        n = n + 1
        .ListColumns(1).DataBodyRange.Rows(n).Value = tDates(iRow)
        .ListColumns(3).DataBodyRange.Rows(n).Value = tValues(iRow)
        .ListColumns(4).DataBodyRange.Rows(n).Value = tDesc(iRow)
    Next iRow
End With

Call sortAccount(ActiveSheet.ListObjects(1))
Range("A" & CStr(n)).Select

End Sub

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportGeneric(fileToOpen As Variant)

Workbooks.Open filename:=fileToOpen, ReadOnly:=True, local:=True
'Workbooks.Open filename:="C:\Users\Olivier\Desktop\Test LCL.csv"
Dim iRow As Integer
Dim tDates() As Variant
Dim tDesc() As String
Dim tSubCateg() As String
Dim tBudgetSpread() As Variant
Dim tValues() As Double

ReDim tDates(1 To 30000)
ReDim tDesc(1 To 30000)
ReDim tSubCateg(1 To 30000)
ReDim tBudgetSpread(1 To 30000)
ReDim tValues(1 To 30000)
iRow = 1

' Read Header part
Do While LenB(Cells(iRow, 1).Value) > 0 And iRow < 30000
    iRow = iRow + 1
    If Cells(iRow, 1) = "Korach Exporter version" Then
        exporterVersion = Cells(iRow, 2).Value
    ElseIf Cells(iRow, 1) = "No Compte" Then
        accountNbr = Cells(iRow, 2).Value
    ElseIf Cells(iRow, 1) = "Nom Compte" Then
        accountName = Cells(iRow, 2).Value
    ElseIf Cells(iRow, 1) = "Banque" Then
        bank = Cells(iRow, 2).Value
    ElseIf Cells(iRow, 1) = "Status" Then
        accStatus = Cells(iRow, 2).Value
    ElseIf Cells(iRow, 1) = "Disponibilité" Then
        availability = Cells(iRow, 2).Value
    Else
        ' Do nothing
    End If
Loop

iRow = iRow + 1
transactionStart = iRow
' Count nbr of transaction
Do While LenB(Cells(iRow, 1).Value) > 0 And iRow < 30000
    iRow = iRow + 1
Loop
' Read transaction part
nbRows = iRow - transactionStart
iRow = transactionStart
Do While LenB(Cells(iRow, 1).Value) > 0
    i = iRow - transactionStart + 1
    tDates(i) = Cells(iRow, 1).Value
    tDesc(i) = Cells(iRow, 4).Value
    tValues(i) = toAmount(Cells(iRow, 3).Value)
    tSubCateg(i) = Cells(iRow, 5).Value
    tBudgetSpread(i) = Cells(iRow, 7).Value
    iRow = iRow + 1
Loop
ActiveWorkbook.Close

ActiveSheet.Cells(1, 2).Value = accountName
ActiveSheet.Cells(2, 2).Value = accountNbr
ActiveSheet.Cells(3, 2).Value = bank
ActiveSheet.Cells(4, 2).Value = accStatus
ActiveSheet.Cells(5, 2).Value = availability

With ActiveSheet.ListObjects(1)
    totalrows = .ListRows.Count
    For iRow = 1 To nbRows
        .ListRows.Add
        totalrows = totalrows + 1
        .ListColumns(1).DataBodyRange.Rows(totalrows).Value = tDates(iRow)
        .ListColumns(2).DataBodyRange.Rows(totalrows).Value = tValues(iRow)
        .ListColumns(4).DataBodyRange.Rows(totalrows).Value = tDesc(iRow)
        .ListColumns(5).DataBodyRange.Rows(totalrows).Value = tSubCateg(iRow)
        .ListColumns(7).DataBodyRange.Rows(totalrows).Value = tBudgetSpread(iRow)
    Next iRow
End With

Call sortAccount(ActiveSheet.ListObjects(1))

Range("A" & CStr(totalrows)).Select

End Sub

Sub ExportGeneric(ws, Optional csvFile As String = "", Optional silent As Boolean = False)

    Dim sFolder As String
    exportFrom = ActiveWorkbook.name

    Sheets(ws).Select
    Range("A1:B8").Select
    Selection.Copy
    exportVersion = 1

    ' Create blank workbook and copy data on that workbook
    Workbooks.Add
    exportTo = ActiveWorkbook.name
    Range("A1").Select
    ActiveSheet.Paste
    Range("A9").Value = "Exporter version"
    Range("B9").Value = 1.2

    Workbooks(exportFrom).Activate
    Sheets(ws).ListObjects(1).DataBodyRange.Select
    Selection.Copy
    Workbooks(exportTo).Activate
    Range("A10").Select
    ActiveSheet.Paste
    Range("B:C").NumberFormat = "General"
    'Range("A:A").NumberFormat = Workbooks(exportFrom).Names("date_format").RefersToRange.Value
    Range("A:A").NumberFormat = "YYYY-mm-dd"

    ' Silently delete sheets in excess
    Call DeleteAllButSheetOne

    ' Get filename to save
    If LenB(csvFile) = 0 Then
        file = Application.GetSaveAsFilename
        If file Then
           csvFile = file & "csv"
        End If
    End If

    ' Save CSV file
    If LenB(csvFile) > 0 Then
        ActiveWorkbook.SaveAs filename:=csvFile, fileformat:=xlCSV, CreateBackup:=False, local:=True
        If (Not silent) Then
            MsgBox "File " & csvFile & " saved"
        End If
    Else
        If Not silent Then
            MsgBox "Export aborted"
        End If
    End If
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True

End Sub
Sub ExportAll()
    Dim sFolder As String
    Dim filename As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With

    If LenB(sFolder) > 0 Then ' if a file was chosen
        Call freezeDisplay
        For Each ws In Worksheets
            If ws.Cells(1, 1).Value = "Nom Compte" Then
                filename = sFolder & "\" & ws.name & ".csv"
                Call ExportGeneric(ws.name, filename, True)
            End If
        Next ws
        Call unfreezeDisplay
    Else
        MsgBox ("Export aborted")
    End If
End Sub
Sub ExportLCL()
    Call ExportGeneric("LCL CC")
End Sub
Sub ExportING()
    Call ExportGeneric("ING CC")
End Sub

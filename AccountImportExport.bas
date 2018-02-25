Attribute VB_Name = "AccountImportExport"


Private Function toAmount(str) As Double
    If VarType(str) = vbString Then
        toAmount = CDbl(Replace(Replace(str, ",", "."), "'", ""))
    Else
        toAmount = str
    End If
End Function
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportAny()

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
Dim tValues()

ReDim tDates(1 To 30000)
ReDim tDesc(1 To 30000)
ReDim tValues(1 To 30000)
iRow = 1
Do While Cells(iRow, 1).Value <> "" And iRow < 30000
    iRow = iRow + 1
Loop
nbRows = iRow - 1
iRow = 1
Do While Cells(iRow, 1).Value <> ""
    tDates(iRow) = Cells(iRow, 1).Value
    tDesc(iRow) = Cells(iRow, 2).Value
    tValues(iRow) = toAmount(Cells(iRow, 4).Value)
    iRow = iRow + 1
Loop
ActiveWorkbook.Close

With Sheets("ING CC").ListObjects(1)
    totalRows = .ListRows.Count
    For iRow = 1 To nbRows
        .ListRows.Add
        totalRows = totalRows + 1
        .ListColumns(1).DataBodyRange.Rows(totalRows).Value = tDates(iRow)
        .ListColumns(2).DataBodyRange.Rows(totalRows).Value = tValues(iRow)
        .ListColumns(4).DataBodyRange.Rows(totalRows).Value = tDesc(iRow)
    Next iRow
End With

Call sortAccount(Sheets("ING CC").ListObjects(1))

Range("A" + CStr(totalRows)).Select

End Sub


'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportLCL(fileToOpen As Variant)

Workbooks.Open filename:=fileToOpen, ReadOnly:=True
Dim iRow As Integer
Dim tDates() As Variant
Dim tDesc() As String
Dim tValues()

ReDim tDates(1 To 30000)
ReDim tDesc(1 To 30000)
ReDim tValues(1 To 30000)
iRow = 1
Do While Cells(iRow, 1).Value <> "" And iRow < 30000
    iRow = iRow + 1
Loop
nbRows = iRow - 2 ' Last row is a total, don't import it
iRow = 1
Do While Cells(iRow, 1).Value <> ""
    tDates(iRow) = DateValue(Cells(iRow, 1).Value)
    tValues(iRow) = toAmount(Cells(iRow, 2).Value)
    If (Cells(iRow, 3).Value = "Chèque") Then
        tDesc(iRow) = "Chèque " + CStr(Cells(iRow, 4).Value)
    ElseIf (Cells(iRow, 3).Value = "Virement") Then
        tDesc(iRow) = "Virement" + " " + Cells(iRow, 5).Value
    Else
        tDesc(iRow) = Cells(iRow, 3).Value + " " + Cells(iRow, 5).Value + " " + Cells(iRow, 6).Value
    End If
    iRow = iRow + 1
Loop
ActiveWorkbook.Close

With Sheets("LCL CC").ListObjects(1)
    totalRows = .ListRows.Count
    For iRow = 1 To nbRows
        .ListRows.Add
        totalRows = totalRows + 1
        .ListColumns(1).DataBodyRange.Rows(totalRows).Value = tDates(iRow)
        .ListColumns(2).DataBodyRange.Rows(totalRows).Value = tValues(iRow)
        .ListColumns(4).DataBodyRange.Rows(totalRows).Value = tDesc(iRow)
    Next iRow
End With

Call sortAccount(Sheets("LCL CC").ListObjects(1))
Range("A" + CStr(totalRows)).Select

End Sub

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportUBS(fileToOpen As Variant)

Workbooks.Open filename:=fileToOpen, ReadOnly:=True
Dim iRow As Integer
Dim tDates() As Variant
Dim tDesc() As String
Dim tValues()

ReDim tDates(1 To 30000)
ReDim tDesc(1 To 30000)
ReDim tValues(1 To 30000)
iRow = 1
nbOps = 0
Do While Cells(iRow, 1).Value <> "" And iRow < 30000
    iRow = iRow + 1
    If (Cells(iRow, 20).Value <> "" Or Cells(iRow, 19).Value <> "") Then
        nbOps = nbOps + 1
    End If
Loop
nbRows = iRow - 1
iRow = 2
nbOps = 0
Do While Cells(iRow, 1).Value <> ""
    If (Cells(iRow, 20).Value <> "" Or Cells(iRow, 19).Value <> "") Then
        nbOps = nbOps + 1
        If (Cells(iRow, 19).Value <> "") Then
            tValues(nbOps) = -toAmount(Cells(iRow, 19).Value) ' Debit column
        Else
            tValues(nbOps) = toAmount(Cells(iRow, 20).Value) ' Credit column
        End If
        tDates(nbOps) = CDate(DateValue(Replace(Cells(iRow, 12).Value, ".", "/")))
        tDesc(nbOps) = Cells(iRow, 13).Value + " " + Cells(iRow, 14).Value + " " + Cells(iRow, 15).Value
    End If
    iRow = iRow + 1
Loop
ActiveWorkbook.Close

With Sheets("UBS").ListObjects(1)
    n = .ListRows.Count
    For iRow = 1 To nbOps
        .ListRows.Add
        n = n + 1
        .ListColumns(1).DataBodyRange.Rows(n).Value = tDates(iRow)
        .ListColumns(3).DataBodyRange.Rows(n).Value = tValues(iRow)
        .ListColumns(4).DataBodyRange.Rows(n).Value = tDesc(iRow)
    Next iRow
End With

Call sortAccount(Sheets("UBS").ListObjects(1))
Range("A" + CStr(n)).Select

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
Dim tValues()

ReDim tDates(1 To 30000)
ReDim tDesc(1 To 30000)
ReDim tSubCateg(1 To 30000)
ReDim tBudgetSpread(1 To 30000)
ReDim tValues(1 To 30000)
iRow = 1

' Read Header part
Do While Cells(iRow, 1).Value <> "" And iRow < 30000
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
        accountStatus = Cells(iRow, 2).Value
    ElseIf Cells(iRow, 1) = "Disponibilité" Then
        availability = Cells(iRow, 2).Value
    Else
        ' Do nothing
    End If
Loop

iRow = iRow + 1
transactionStart = iRow
' Count nbr of transaction
Do While Cells(iRow, 1).Value <> "" And iRow < 30000
    iRow = iRow + 1
Loop
' Read transaction part
nbRows = iRow - transactionStart
iRow = transactionStart
Do While Cells(iRow, 1).Value <> ""
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
ActiveSheet.Cells(4, 2).Value = accountStatus
ActiveSheet.Cells(5, 2).Value = availability

With ActiveSheet.ListObjects(1)
    totalRows = .ListRows.Count
    For iRow = 1 To nbRows
        .ListRows.Add
        totalRows = totalRows + 1
        .ListColumns(1).DataBodyRange.Rows(totalRows).Value = tDates(iRow)
        .ListColumns(2).DataBodyRange.Rows(totalRows).Value = tValues(iRow)
        .ListColumns(4).DataBodyRange.Rows(totalRows).Value = tDesc(iRow)
        .ListColumns(5).DataBodyRange.Rows(totalRows).Value = tSubCateg(iRow)
        .ListColumns(7).DataBodyRange.Rows(totalRows).Value = tBudgetSpread(iRow)
    Next iRow
End With

Call sortAccount(ActiveSheet.ListObjects(1))

Range("A" + CStr(totalRows)).Select

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
    Range("A6").Value = "Korach Exporter version"
    Range("B6").Value = 1.1
    
    Workbooks(exportFrom).Activate
    Sheets(ws).ListObjects(1).DataBodyRange.Select
    Selection.Copy
    Workbooks(exportTo).Activate
    Range("A8").Select
    ActiveSheet.Paste
    Range("C:C").NumberFormat = "General"
    'Range("A:A").NumberFormat = Workbooks(exportFrom).Names("date_format").RefersToRange.Value
    Range("A:A").NumberFormat = "YYYY-mm-dd"
    
    ' Silently delete sheets in excess
    Call DeleteAllButSheetOne

    ' Get filename to save
    If (csvFile = "") Then
        file = Application.GetSaveAsFilename
        If file <> False Then
           csvFile = file + "csv"
        End If
    End If
    
    ' Save CSV file
    If csvFile <> "" Then
        ActiveWorkbook.SaveAs filename:=csvFile, fileformat:=xlCSV, CreateBackup:=False, local:=True
        If (Not silent) Then
            MsgBox "File " & csvFile & " saved"
        End If
    Else
        If (Not silent) Then
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
    
    If sFolder <> "" Then ' if a file was chosen
        Call freezeDisplay
        For Each ws In Worksheets
            If ws.Cells(1, 1).Value = "Nom Compte" Then
                filename = sFolder + "\" + ws.name + ".csv"
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

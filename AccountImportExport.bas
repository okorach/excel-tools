Attribute VB_Name = "AccountImportExport"

Const NOT_AN_ACCOUNT = 0
Const DOMESTIC_ACCOUNT = 1
Const FOREIGN_ACCOUNT = 2
Const DOMESTIC_SHARES_ACCOUNT = 3
Const FOREIGN_SHARES_ACCOUNT = 4

Const ACCOUNT_CLOSED = 0
Const ACCOUNT_OPEN = 1

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
    Range("A1:B5").Select
    Selection.Copy
    exportVersion = 1

    ' Create blank workbook and data on that workbook
    Workbooks.Add
    exportTo = ActiveWorkbook.name
    Range("A1").Select
    ActiveSheet.Paste
    Range("A6").Value = "Korach Exporter version"
    Range("B6").Value = 1
    
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



Sub getFolder()
Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then ' if a file was chosen
        ' *********************
        ' put your code in here
        ' *********************
    End If
End Sub

Sub CreateAccount()
    accountNbr = InputBox("Account number ?", "Account Number", "<accountNumber>")
    accountName = InputBox("Account name ?", "Account Name", "<accountName>")
    Sheets("Account Template").Visible = True
    Sheets("Account Template").Copy Before:=Sheets(1)
    Sheets("Account Template").Visible = False
    Sheets(1).name = accountName
    ' Sheets(1).Range("A1").Formula = "=VLOOKUP("k.account", TblKeys, LangId, FALSE)"
    Sheets(1).Range("B1").Value = accountName
    Sheets(1).Range("B2").Formula = "=VLOOKUP(B$1,TblAccounts,2,FALSE)"
    Sheets(1).Range("B3").Formula = "=VLOOKUP(B$1,TblAccounts,4,FALSE)"
    Sheets(1).Range("B4").Formula = "=VLOOKUP(B$1,TblAccounts,6,FALSE)"
    Sheets(1).Range("B5").Formula = "=VLOOKUP(B$1,TblAccounts,5,FALSE)"
End Sub

Public Sub refreshOpenAccountsList()
    Call freezeDisplay
    Call truncateTable(Sheets("Paramètres").ListObjects("tblOpenAccounts"))
    With Sheets("Paramètres").ListObjects("tblOpenAccounts")
        For i = 1 To Sheets("Comptes").ListObjects("tblAccounts").ListRows.Count
            If (Sheets("Comptes").ListObjects("tblAccounts").ListRows(i).Range.Cells(1, 6).Value = "Open") Then
                .ListRows.Add ' Add 1 row at the end, then extend
                .ListRows(.ListRows.Count).Range.Cells(1, 1).Value = Sheets("Comptes").ListObjects("tblAccounts").ListRows(i).Range.Cells(1, 1).Value
            End If
        Next i
        nbrAccounts = .ListRows.Count + 1
    End With
    ActiveSheet.Shapes("Drop Down 2").Select
    With Selection
        .ListFillRange = "Paramètres!$L$2:$L$" + CStr(Sheets("Paramètres").ListObjects("tblOpenAccounts").ListRows.Count + 1)
        .LinkedCell = "$H$72"
        .DropDownLines = 8
        .Display3DShading = True
    End With
    Call unfreezeDisplay
End Sub

Public Sub sortCurrentAccount()
    Call sortAccount(ActiveSheet.ListObjects(1))
End Sub
Public Sub sortAccount(oTable)
    oTable.Sort.SortFields.Clear
    ' Sort table by date first, then by amount
    oTable.Sort.SortFields.Add Key:=Range(oTable.name + "[Date]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    oTable.Sort.SortFields.Add Key:=Range(oTable.name + "[Montant]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With oTable.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ' Reset date column format
    Call setTableColumnFormat(oTable, 1, "m/d/yyyy")
End Sub
'-------------------------------------------------
Public Function accountType(accountName As String) As String
    If (accountName = "Account Template") Then
        accountType = "Standard"
    ElseIf (Not accountExists(accountName)) Then
        accountType = "ERROR: Not an account"
    ElseIf (Sheets(accountName).Range("B6").Value = "EUR") Then
        accountType = Sheets(accountName).Range("B7").Value
    End If
End Function
'-------------------------------------------------
Public Function accountNumber(accountName As String) As String
    If (accountExists(accountName)) Then
        accountNumber = Sheets(accountName).Range("B2").Value
    Else
        accountNumber = ""
    End If
End Function
'-------------------------------------------------
Public Function accountBank(accountName As String) As String
    If (accountExists(accountName)) Then
        accountBank = Sheets(accountName).Range("B3").Value
    Else
        accountBank = ""
    End If
End Function
'-------------------------------------------------
Public Function accountStatus(accountName As String) As String
    If (accountExists(accountName)) Then
        accountStatus = Sheets(accountName).Range("B4").Value
    Else
        accountStatus = ""
    End If
End Function
'-------------------------------------------------
Public Function isOpen(accountName As String) As Boolean
    isOpen = False
    If (accountStatus(accountName) = "Open") Then
        isOpen = True
    End If
End Function
Public Function isClosed(accountName As String) As Boolean
    isClosed = Not isOpen(accountName)
End Function
'-------------------------------------------------
Public Function accountAvailability(accountName As String) As String
    If (accountExists(accountName)) Then
        accountAvailability = Sheets(accountName).Range("B5").Value
    Else
        accountAvailability = ""
    End If
End Function
'-------------------------------------------------
Public Function accountCurrency(accountName As String) As String
    If (accountExists(accountName)) Then
        accountCurrency = Sheets(accountName).Range("B6").Value
    Else
        accountCurrency = ""
    End If
End Function

'-------------------------------------------------
Public Function accountExists(accountName As String) As Boolean
    If (sheetExists(accountName) And Sheets(accountName).Range("A1") = "Nom Compte") Then
        accountExists = True
    Else
        accountExists = False
    End If
End Function
'-------------------------------------------------
Public Function isAnAccountSheet(ByVal ws As Worksheet) As Boolean
    If (ws.Cells(1, 1).Value = getNamedVariableValue("accountIdentifier") And Not isTemplate(ws)) Then
        isAnAccountSheet = True
    Else
        isAnAccountSheet = False
    End If
End Function
'-------------------------------------------------
Public Function isTemplate(ByVal ws As Worksheet) As Boolean
    If (ws.Cells(1, 2).Value = "TEMPLATE") Then
        isTemplate = True
    Else
        isTemplate = False
    End If
End Function

'-------------------------------------------------
Public Sub hideClosedAccounts()
    If (ThisWorkbook.Names("hideClosedAccounts").RefersToRange.Value = 1) Then
    For Each ws In Worksheets
        If (accountStatus(ws.name) = "Closed") Then
            ws.Visible = False
        End If
    Next ws
    End If
End Sub
'-------------------------------------------------
Public Sub hideTemplateAccounts()
    For Each ws In Worksheets
        If (isTemplate(ws)) Then
            ws.Visible = False
        End If
    Next ws
End Sub
'-------------------------------------------------
Public Sub showAllSheets()
    For Each ws In Worksheets
        ws.Visible = True
    Next ws
End Sub
'-------------------------------------------------
Public Sub formatAccountSheets()

   For Each ws In Worksheets
       ' Make sure the sheet is not anything else than an account
       If (isAnAccountSheet(ws)) Then
            Dim name As String
            Dim acctype As String
            Dim acurrency As String
            name = ws.name
            acctype = accountType(name)
            acurrency = accountCurrency(name)
            If (acctype = "Standard") Then
                If (acurrency = "EUR") Then
                   Call SetColumnWidth("A", 15, name)
                   Call SetColumnWidth("B", 20, name)
                   Call SetColumnWidth("C", 20, name)
                   Call SetColumnWidth("D", 70, name)
                   Call SetColumnWidth("E", 15, name)
                   Call SetColumnWidth("F", 15, name)
                   Call SetColumnWidth("G", 5, name)
                   Call SetColumnWidth("H", 5, name)
                   Call SetColumnWidth("I", 15, name)
                Else
                   Call SetColumnWidth("A", 15, name)
                   Call SetColumnWidth("B", 20, name)
                   Call SetColumnWidth("C", 20, name)
                   Call SetColumnWidth("D", 70, name)
                   Call SetColumnWidth("E", 15, name)
                   Call SetColumnWidth("F", 15, name)
                   Call SetColumnWidth("G", 5, name)
                   Call SetColumnWidth("H", 5, name)
                   Call SetColumnWidth("I", 15, name)
                End If
            Else ' Shares accounts formatting
                If (acurrency = "EUR") Then
                   Call SetColumnWidth("A", 12, name)
                   Call SetColumnWidth("B", 20, name)
                   Call SetColumnWidth("C", 20, name)
                   Call SetColumnWidth("D", 70, name)
                   Call SetColumnWidth("E", 20, name)
                   Call SetColumnWidth("F", 5, name)
                   Call SetColumnWidth("G", 20, name)
                   Call SetColumnWidth("H", 20, name)
                Else
                   Call SetColumnWidth("A", 12, name)
                   Call SetColumnWidth("B", 20, name)
                   Call SetColumnWidth("C", 20, name)
                   Call SetColumnWidth("D", 70, name)
                   Call SetColumnWidth("E", 20, name)
                   Call SetColumnWidth("F", 15, name)
                   Call SetColumnWidth("G", 5, name)
                   Call SetColumnWidth("H", 15, name)
                End If
            End If
          ws.Cells.RowHeight = 13
          ws.Rows.Font.size = 10

          If (ws.Shapes.Count > 0) Then
            Dim i As Integer
            i = 0
            For Each Shape In ws.Shapes
                If (Shape.Type = msoFormControl) Then
                    ' This is a button, move it to right place
                    Call ShapePlacementXY(Shape, 300, 5 + i * 20, 400, 25 + i * 20)
                    i = i + 1
                End If
            Next Shape
          End If
       End If
   Next ws
   Call hideClosedAccounts
   Call hideTemplateAccounts
End Sub





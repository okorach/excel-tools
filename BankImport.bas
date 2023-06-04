Attribute VB_Name = "BankImport"
Private Const SUBSTITUTIONS_TABLE = "TblSubstitutions"

Function toAmount(str) As Double
    If VarType(str) = vbString Then
        str = Replace(Replace(Replace(str, ",", "."), "'", ""), " ", "")
        toAmount = CDbl(str)
    Else
        toAmount = str
    End If
End Function

Function toMonth(str) As Long
    s = LCase$(Trim$(str))
    If s Like "jan*" Then
        toMonth = 1
    ElseIf s Like "f?[bv]*" Then
        toMonth = 2
    ElseIf s Like "mar*" Then
        toMonth = 3
    ElseIf s Like "a[vp]r*" Then
        toMonth = 4
    ElseIf s Like "ma[iy]*" Then
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
    ElseIf s Like "d?c*" Then
        toMonth = 12
    Else
        toMonth = 0
    End If
End Function

Function ToDate(str) As Date
    If InStr(str, " ") Then
        a = Split(str, " ", -1, vbTextCompare)
        ToDate = DateSerial(CInt(a(2)), toMonth(a(1)), CInt(a(0)))
    ElseIf InStr(str, "/") Then
        a = Split(str, "/", -1, vbTextCompare)
        ' Assume DD/MM/YYYY
        ToDate = DateSerial(CInt(a(2)), CInt(a(1)), CInt(a(0)))
    ElseIf InStr(str, "-") Then
        ToDate = isoToDate(str)
    Else
        ToDate = DateSerial(0, 0, 0)
    End If
End Function

Private Function isoToDate(str) As Date
    a = Split(str, "-", -1, vbTextCompare)
    isoToDate = DateSerial(CInt(a(0)), CInt(a(1)), CInt(a(2)))
End Function

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Sub ImportAny()

    Dim fileToOpen As Variant
    fileToOpen = Application.GetOpenFilename()
    If fileToOpen <> False Then
        Call FreezeDisplay
        Dim oAccount As Account
        Set oAccount = LoadAccount(getAccountId(ActiveSheet))
        
        Dim defaultCurrency As String
        defaultCurrency = GetGlobalParam("DefaultCurrency")
        Dim oTable As ListObject
        Set oTable = oAccount.BalanceTable
        
        Dim dateCol As Integer, amountCol As Integer, descCol As Integer
        dateCol = TableColNbrFromName(oTable, GetLabel(DATE_KEY))
        
        If oAccount.MyCurrency = defaultCurrency Then
            amountCol = TableColNbrFromName(oTable, GetLabel(AMOUNT_KEY))
        Else
            amountCol = TableColNbrFromName(oTable, GetLabel(AMOUNT_KEY) & " " & oAccount.MyCurrency)
        End If
        descCol = TableColNbrFromName(oTable, GetLabel(DESCRIPTION_KEY))
        If (oAccount.Bank = "ING") Then
            Call ImportING(oTable, fileToOpen, dateCol, amountCol, descCol)
        ElseIf (oAccount.Bank = "LCL") Then
            Call ImportLCL(oTable, fileToOpen, dateCol, amountCol, descCol)
        ElseIf (oAccount.Bank = "UBS") Then
            Call ImportUBS(oTable, fileToOpen, dateCol, amountCol, descCol)
        ElseIf (oAccount.Bank = "Revolut") Then
            Call ImportRevolut(oTable, fileToOpen, dateCol, amountCol, descCol)
        ElseIf (oAccount.Bank = "Boursorama") Then
            Call ImportBoursorama(oTable, fileToOpen, dateCol, amountCol, descCol, oAccount.Number())
        Else
            Call UnfreezeDisplay
            Call ErrorMessage("k.errorImportNotRecognized", "k.warningImportCancelled")
        End If
        Call SortTable(oTable, GetLabel(DATE_KEY), xlAscending, GetLabel(AMOUNT_KEY), xlDescending)
        Range("A" & CStr(oTable.ListRows.Count)).Select
        Call UnfreezeDisplay
    Else
        Call ErrorMessage("k.warningImportCancelled")
    End If
End Sub


Public Sub AccountImport()
    Dim fileToOpen As Variant
    fileToOpen = Application.GetOpenFilename()
    If fileToOpen <> False Then
        Call ImportGeneric(ActiveSheet.name, fileToOpen)
    End If
End Sub

' PRLV SEPA CE URSSAF RHONE ALPES-CNTFS : FR55ZZZ143065 000828DC120181231145950A000136092 DE
' CE URSSAF RHONE ALPES-CNTFS : 000828DC120181231145950A000136092 FR55ZZZ143065
Function deleteDuplicateSepa(desc As String) As String
    Dim idstr As String
    idstr = "PRLV SEPA "
    deleteDuplicateSepa = desc
    If (InStr(desc, idstr) = 1) Then
        Dim i_end_emitter As Long
        Dim s_emitter As String
        Dim i_repeat_emitter As Long
        i_end_emitter = InStr(desc, ":")
        If i_end_emitter > 0 Then
            s_emitter = Mid$(desc, Len(idstr) + 1, i_end_emitter - Len(idstr) - 2)
            i_repeat_emitter = InStr(desc, " DE " & s_emitter)
            If i_repeat_emitter > 0 Then
                deleteDuplicateSepa = left$(desc, i_repeat_emitter - 1)
            End If
        End If
    End If
End Function

Function strReplace(oldString, newString, targetString As String) As String
    strReplace = targetString
    i = InStr(targetString, oldString)
    If (i > 0) Then
        strReplace = left$(targetString, i - 1) & newString & Right$(targetString, Len(targetString) - i - Len(oldString) + 1)
    End If
End Function

Function simplifyDescription(desc As String, subsTable As Variant) As String
    Dim s As String
    s = deleteDuplicateSepa(Trim$(desc))
    n = UBound(subsTable, 1)
    For i = 1 To n
        s = strReplace(subsTable(i, 1), subsTable(i, 2), s)
    Next i
    simplifyDescription = s
End Function

'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Public Sub ImportGeneric(accountId As String, fileToOpen As Variant)

   ' Dim xlsFile As String
    Dim importFrom As String, importTo As String
    importTo = ActiveWorkbook.name

    subsTable = GetTableAsArray(Workbooks(importTo).Sheets(PARAMS_SHEET).ListObjects(SUBSTITUTIONS_TABLE))
    
    Dim accId As String, cur As String, typ As String, avail As Integer, exportDate As String
    Dim Bank As String, TaxRate As Double, accNbr As String
    Dim nbrTr As Long, nbrDep As Long

    Dim modal As ProgressBar
    Set modal = NewProgressBar("Import Generic CSV in progress", 9)
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True, local:=True
    importFrom = ActiveWorkbook.name
    Call AccountImportMetadata(ActiveSheet, accountId, exportDate, accNbr, Bank, avail, cur, typ, TaxRate, nbrTr, nbrDep)
    modal.Update

    Dim offset As Integer
    offset = 0
    If cur <> GetGlobalParam("DefaultCurrency", Workbooks(importTo)) Then
        offset = BL_FOREIGN_OFFSET
    End If

    Dim balanceTbl As ListObject, depositsTbl As ListObject
    Set balanceTbl = accountBalanceTable(accountId)
    Set depositsTbl = accountDepositTable(accountId)
    modal.Update
    
    Dim lastRow As String, firstRow As String
    lastRow = CStr(nbrTr + 1)

    Workbooks(importFrom).Activate
    Range("A2:A" & lastRow).Select
    Selection.Copy
    Workbooks(importTo).Activate
    balanceTbl.ListRows(1).Range(1, BL_DATE_COL).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    modal.Update
    
    Workbooks(importFrom).Activate
    Range("B2:B" & lastRow).Select
    Selection.Copy
    Workbooks(importTo).Activate
    balanceTbl.ListRows(1).Range(1, BL_AMOUNT_COL + offset).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    modal.Update

    Workbooks(importFrom).Activate
    Range("C2:C" & lastRow).Select
    Selection.Copy
    Workbooks(importTo).Activate
    balanceTbl.ListRows(1).Range(1, BL_BALANCE_COL + offset).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    modal.Update

    Workbooks(importFrom).Activate
    Range("D2:D" & lastRow).Select
    Selection.Copy
    Workbooks(importTo).Activate
    balanceTbl.ListRows(1).Range(1, BL_DESC_COL + offset).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    modal.Update

    Workbooks(importFrom).Activate
    Range("E2:E" & lastRow).Select
    Selection.Copy
    Workbooks(importTo).Activate
    balanceTbl.ListRows(1).Range(1, BL_SUBCATEG_COL + offset).Select
    Selection.PasteSpecial Paste:=xlPasteValues
    modal.Update

    ' Read deposits part if any
    If nbrDep > 0 Then
        firstRow = CStr(nbrTr + 2)
        lastRow = CStr(nbrTr + nbrDep + 1)
        Workbooks(importFrom).Activate
        Range("A" & firstRow & ":A" & lastRow).Select
        Selection.Copy
        Workbooks(importTo).Activate
        depositsTbl.ListRows(1).Range(1, DP_DATE_COL).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        modal.Update

        Workbooks(importFrom).Activate
        Range("B" & firstRow & ":B" & lastRow).Select
        Selection.Copy
        Workbooks(importTo).Activate
        depositsTbl.ListRows(1).Range(1, DP_AMOUNT_COL).Select
        Selection.PasteSpecial Paste:=xlPasteValues
        modal.Update
    Else
        modal.Update 2
    End If
    Application.DisplayAlerts = False
    Workbooks(importFrom).Close
    Application.DisplayAlerts = True
    Set modal = Nothing
End Sub


Public Sub AccountCreateFromCSV()
    Dim fileToOpen As Variant
    fileToOpen = Application.GetOpenFilename()
    If fileToOpen = False Then
        Exit Sub
    End If

    Dim importFrom As String, importTo As String
    importTo = ActiveWorkbook.name
    Workbooks.Open filename:=fileToOpen, ReadOnly:=True, local:=True
    importFrom = ActiveWorkbook.name

    Dim accountId As String, cur As String, typ As String, avail As Integer, exportDate As String
    Dim Bank As String, TaxRate As Double, accNbr As String
    Dim nbrTr As Long, nbrDep As Long
    
    Call AccountImportMetadata(Workbooks(importFrom).ActiveSheet, accountId, exportDate, accNbr, Bank, avail, cur, _
        typ, TaxRate, nbrTr, nbrDep)
    
    Workbooks(importFrom).Close SaveChanges:=False
    Workbooks(importTo).Activate
    Call AccountCreate(accountId, cur, typ, avail, accNbr, Bank)
    Sheets(accountId).Activate
    Call ImportGeneric(accountId, fileToOpen)
End Sub


Public Sub AccountExportHere()
    Dim oAccount As Account
    Set oAccount = LoadAccount(getAccountId(ActiveSheet))
    If Not (oAccount Is Nothing) Then
        oAccount.Export
    End If
End Sub


Public Sub AccountExportAll()
    Dim sFolder As String
    Dim filename As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    If LenB(sFolder) = 0 Then ' if no directory was chosen
        Call ErrorMessage("k.warningExportCancelled")
        Exit Sub
    End If
    
    Dim modal As ProgressBar
    Dim ws As Worksheet, curWs As Worksheet
    Set modal = NewProgressBar("Export all accounts in progress", Worksheets.Count)
    Call FreezeDisplay
    Set curWs = ActiveSheet
    
    For Each ws In Worksheets
        Dim oAccount As Account
        Set oAccount = LoadAccount(getAccountId(ws))
        If Not (oAccount Is Nothing) Then
            filename = sFolder & "\" & ws.name & ".csv"
            Call oAccount.Export(csvFile:=filename, silent:=True)
        End If
        modal.Update
    Next ws
    curWs.Activate
    Set modal = Nothing
    Call UnfreezeDisplay
End Sub

Private Sub AccountExportMetadata(accountId As String, targetWs As Worksheet, nbrTransactions As Long, Optional nbrDeposits As Long = 0)
    ' Copy metadata on row 1
    targetWs.Range("A1") = "ExportDate=" & Format$(Now(), "YYYY-mm-dd HH:MM:SS")
    targetWs.Range("B1") = "AccountId=" & accountId
    targetWs.Range("C1") = "AccountNumber=" & AccountNumber(accountId)
    targetWs.Range("D1") = "Bank=" & AccountBank(accountId)
    avail = AccountAvailability(accountId)
    targetWs.Range("E1") = "Availability=" & avail
    targetWs.Range("F1") = "Currency=" & AccountCurrency(accountId)
    targetWs.Range("G1") = "Type=" & AccountType(accountId)
    targetWs.Range("H1") = "TaxRate=" & AccountTaxRate(accountId)
    targetWs.Range("I1") = "NbrTransactions=" & CStr(nbrTransactions)
    If nbrDeposits > 0 Then
        targetWs.Range("J1") = "NbrDeposits=" & CStr(nbrDeposits)
    End If
End Sub


Private Sub AccountImportMetadata(ws As Worksheet, accountId As String, exportDate As String, accNumber As String, _
    Bank As String, avail As Integer, accCurrency As String, accType As String, TaxRate As Double, _
    nbrTransactions As Long, nbrDeposits As Long)
    ' Copy metadata on row 1
    nbrTransactions = 0
    nbrDeposits = 0
    Dim i As Long
    i = 1
    Do While LenB(ws.Cells(1, i).value) > 0
        a = Split(ws.Cells(1, i).value, "=", -1, vbTextCompare)
        If a(0) = "AccountId" Then
            accountId = a(1)
        ElseIf a(0) = "ExportDate" Then
            exportDate = a(1)
        ElseIf a(0) = "AccountNumber" Then
            accNumber = a(1)
        ElseIf a(0) = "Bank" Then
            Bank = a(1)
        ElseIf a(0) = "Availability" Then
            avail = CInt(a(1))
        ElseIf a(0) = "Currency" Then
            accCurrency = a(1)
        ElseIf a(0) = "Type" Then
            accType = a(1)
        ElseIf a(0) = "TaxRate" Then
            TaxRate = CDbl(a(1))
        ElseIf a(0) = "NbrTransactions" Then
            nbrTransactions = CLng(a(1))
        ElseIf a(0) = "NbrDeposits" Then
            nbrDeposits = CLng(a(1))
        End If
        i = i + 1
    Loop
End Sub


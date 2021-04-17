VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private aId As String
Private aName As String
Private aNumber As String
Private aBank As String
Private aAvail As Integer
Private aStatus As String
Private aCurrency As String
Private aType As String
Private aInBudget As Boolean
Private aTaxRate As Double
Private aInterestPeriod As Integer
Private ws As Worksheet
Private oBalanceTable As ListObject
Private oDepositsTable As ListObject
Private oInterestTable As ListObject
Private oInterests As Interest

Private Sub Class_Initialize()
    aId = vbNullString
    aName = vbNullString
    aNumber = vbNullString
    aBank = vbNullString
    aAvail = 0
    aStatus = vbNullString
    aCurrency = vbNullString
    aType = vbNullString
    aInBudget = False
    aTaxRate = vbNull
    aInterestPeriod = 0
    Set oBalanceTable = Nothing
    Set oDepositsTable = Nothing
    Set oInterestTable = Nothing
    Set oInterests = Nothing
    Set ws = Nothing
End Sub

Public Function Load(accountId As String) As Boolean
    aId = accountId
    Set ws = Worksheets(aId)
    Set tbl = NewKeyedTable(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE))
    If Not Me.Exists Then
        Load = False
        Exit Function
    End If
    aNumber = CStr(tbl.Lookup(aId, 2))
    aName = CStr(tbl.Lookup(aId, 3))
    aBank = CStr(tbl.Lookup(aId, 4))
    aAvail = CInt(tbl.Lookup(aId, 5))
    aStatus = CStr(tbl.Lookup(aId, 6))
    aCurrency = CStr(tbl.Lookup(aId, 7))
    aType = CStr(tbl.Lookup(aId, 8))
    aInBudget = CBool(tbl.Lookup(aId, 9))
    aTaxRate = CDbl(tbl.Lookup(aId, 10))
    
    Dim accountTypes As KeyedTable
    Set accountTypes = NewKeyedTable(Sheets(PARAMS_SHEET).ListObjects(ACCOUNT_TYPES_TABLE))
    aInterestPeriod = CInt(accountTypes.Lookup(aType, 2))


    Set oBalanceTable = accountTable(BALANCE_TABLE_NAME)
    Set oDepositsTable = accountTable(DEPOSIT_TABLE_NAME)
    Set oInterestTable = accountTable(INTEREST_TABLE_NAME)

    Load = True
End Function

Public Function Create(accId As String, accNbr As String, accBank As String, _
                       Optional accCur As String = vbNullString, _
                       Optional accType As String = vbNullString, Optional accAvail As Integer = 0, _
                       Optional accInB As Boolean = False, Optional accTax As Double = 0) As Boolean
         
    Call FreezeDisplay
    
    Dim tbl As KeyedTable
    Set tbl = NewKeyedTable(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE))
    If tbl.KeyExists(accountId) Then
        MsgBox ("Account already exist, aborting")
        Create = False
        Exit Function
    End If

    
    If accCur = vbNullString Then
        accCur = GetGlobalParam("DefaultCurrency")
    End If
    If accType = vbNullString Then
        accType = "Courant"
    End If

    aId = accId
    aNumber = accNbr
    aType = accType
    aBank = accBank
    aCurrency = accCur
    aType = accType
    aAvail = accAvail
    aInB = accInB
    aTax = accTax
    
    tbl.Insert aId, aNumber, 2
    tbl.Update aId, aId, 3
    tbl.Update aId, aBank, 4
    tbl.Update aId, aAvail, 5
    tbl.Update aId, GetLabel("k.AccountOpen"), 6
    tbl.Update aId, aCurrency, 7
    tbl.Update aId, aType, 8
    tbl.Update aId, aInB, 9
    tbl.Update aId, aTax, 10

    Sheets.Add
    ActiveSheet.Name = aId
    Set ws = ActiveSheet
    addBalanceTable
    addStandardButtons
    If aType = GetLabel("k.accountStandard") Then
        addImportButton
    Else
        addDepositTable
        addInterestTable
        addInterestButtons
    End If
    FormatMe
    Call UnfreezeDisplay
    Create = True
End Function

Public Sub Export(Optional csvFile As String = "", Optional silent As Boolean = False)
    ' Get filename to save
    If LenB(csvFile) = 0 Then
        Dim file As Variant
        file = Application.GetSaveAsFilename(InitialFileName:=accountId & ".csv")
        If file = False Then
            Call ErrorMessage("k.warningExportCancelled")
            Exit Sub
        End If
        csvFile = CStr(file)
        If LCase$(Right$(csvFile, 3)) <> "csv" Then
            csvFile = csvFile & "csv"
        End If
    End If

    Call FreezeDisplay
    
    Dim exportFrom As String, exportTo As String
    exportFrom = ActiveWorkbook.Name
    
    Dim defaultCurrency As String
    defaultCurrency = GetGlobalParam("DefaultCurrency")

    ' Copy account transactions
    oBalanceTable.DataBodyRange.Select
    Selection.Copy
    
    ' Create blank workbook and copy data on that workbook to save as CSV
    Workbooks.Add (xlWBATWorksheet)
    exportTo = ActiveWorkbook.Name
    
    ' Paste account transactions starting from row 2
    Workbooks(exportTo).Activate
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues

    ' Delete useless category row
    If aCurrency = defaultCurrency Then
        For Each c In Array("F")
            ActiveSheet.Columns(c).EntireColumn.Delete shift:=xlToLeft
        Next c
    Else
        ' For non default currency account also remove columns with amounts converted to default currency
        For Each c In Array("H", "C", "B")
            ActiveSheet.Columns(c).EntireColumn.Delete shift:=xlToLeft
        Next c
    End If

    Workbooks(exportFrom).Activate
    ' If account has deposits (savings account), store them
    If Not oDepositsTable Is Nothing Then
        Dim rowNbr As Long
        oDepositsTable.DataBodyRange.Select
        Selection.Copy
        rowNbr = oBalanceTable.ListRows.Count + 1
        Workbooks(exportTo).Activate
        Range("A" & CStr(rowNbr)).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    End If

    ' Set universal format for dates and numbers
    Range("A:A").NumberFormat = "YYYY-mm-dd"
    Range("B:E").NumberFormat = "General"

    ' Copy metadata on row 1
    Workbooks(exportFrom).Activate
    If oDepositsTable Is Nothing Then
        Call ExportMetadata(Workbooks(exportTo).ActiveSheet, oBalanceTable.ListRows.Count)
    Else
        Call ExportMetadata(Workbooks(exportTo).ActiveSheet, oBalanceTable.ListRows.Count, oDepositsTable.ListRows.Count)
    End If

    ' Save CSV file
    Workbooks(exportTo).Activate
    ActiveWorkbook.SaveAs filename:=csvFile, fileformat:=xlCSV, CreateBackup:=False, local:=True
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
    Workbooks(exportFrom).ActiveSheet.Range("A1").Select
    Call UnfreezeDisplay
    If Not silent Then
        MsgBox "File " & csvFile & " saved"
    End If
End Sub

Public Sub CalcInterests()
    If Not Me.HasInterests Then
        Exit Sub
    End If
    Dim oInterests As Interest
    Set oInterests = NewInterest(aId, balanceHistory, GetTableAsArray(oDepositsTable), aInterestPeriod)
    oInterests.Calc
    Dim b As Boolean
    b = oInterests.Store(oInterestTable, aId, aTaxRate)
End Sub

Private Sub addStandardButtons()
    For Each btnName In Array("BtnHome", "BtnPrev5", "BtnPrev", "BtnNext", "BtnNext5", "BtnTop", "BtnBottom", "BtnFormat", "BtnSort")
        If Not ShapeExist(ws, CStr(btnName)) Then
            Call BtnAdd(ws, Name:=CStr(btnName), action:="ThisWorkbook.GoToSolde")
        End If
    Next btnName
End Sub

Private Sub addInterestButtons()
    For Each btnName In Array("BtnInterest", "BtnAddEntry")
        If Not ShapeExist(ws, CStr(btnName)) Then
            Call BtnAdd(ws, Name:=CStr(btnName), action:="ThisWorkbook.GoToSolde")
        End If
    Next btnName
End Sub

Private Sub addImportButton()
    For Each btnName In Array("BtnImport")
        If Not ShapeExist(ws, CStr(btnName)) Then
            Call BtnAdd(ws, Name:=CStr(btnName), action:="ThisWorkbook.GoToSolde")
        End If
    Next btnName
End Sub

Private Sub addBalanceTable()
    Dim tblName As String
    tblName = Replace(Replace(Replace(LCase$(aId), " ", "_"), "é", "e"), "è", "e") & "_" & BALANCE_TABLE_NAME
    ' TODO: Don't create table if already exists
    ws.Range("A11:A12").Select
    ws.ListObjects.Add(xlSrcRange, Range("$A$10:$E$11"), , xlYes).Name = tblName

    With ws.ListObjects(tblName)
        .TableStyle = "TableStyleMedium2"
        Dim lblAmount As String, lblBalance As String, lblDate As String, lblSubcat As String
        lblAmount = GetLabel("k.amount")
        lblBalance = GetLabel("k.accountBalance")
        lblDate = GetLabel("k.date")
        lblSubcat = GetLabel("k.subcategory")
        .ListColumns(BL_DATE_COL).Name = lblDate
        .ListColumns(BL_AMOUNT_COL).Name = lblAmount
        .ListColumns(BL_BALANCE_COL).Name = lblBalance
        .ListColumns(BL_DESC_COL).Name = GetLabel("k.description")
        .ListColumns(BL_SUBCATEG_COL).Name = lblSubcat
        If aType = GetLabel("k.accountStandard") Then
            .ListRows(1).Range(1, BL_BALANCE_COL).FormulaR1C1 = "=[" & lblAmount & "]+IF(ISNUMBER(R[-1]C),R[-1]C,0)"
            .ListColumns.Add
            .ListColumns(BL_CATEG_COL).Name = GetLabel("k.category")
            .ListRows(1).Range(1, BL_CATEG_COL).FormulaR1C1 = "=VLOOKUP([" & lblSubcat & "],TableCategories,2,FALSE)"
            If aInBudget Then
                .ListColumns.Add
                .ListColumns(BL_BUDGET_COL).Name = GetLabel("k.inBudget")
            End If
        Else
            .ListRows(1).Range(1, BL_AMOUNT_COL).FormulaR1C1 = "=[" & lblBalance & "]-IF(ISNUMBER(R[-1]C[1]),R[-1]C[1],0)"
        End If
        If aCurrency <> GetGlobalParam("DefaultCurrency") Then
            lblAmount = lblAmount & " " & aCurrency
            lblBalance = lblBalance & " " & aCurrency
            .ListColumns.Add(BL_AMOUNT_COL + BL_FOREIGN_OFFSET).Name = lblAmount
            .ListColumns.Add(BL_BALANCE_COL + BL_FOREIGN_OFFSET).Name = lblBalance
            If aType = GetLabel("k.accountStandard") Then
                .ListRows(1).Range(1, BL_BALANCE_COL + BL_FOREIGN_OFFSET).FormulaR1C1 = "=[" & lblAmount & "]+IF(ISNUMBER(R[-1]C),R[-1]C,0)"
            Else
                .ListRows(1).Range(1, BL_AMOUNT_COL + BL_FOREIGN_OFFSET).FormulaR1C1 = "=[" & lblBalance & "]-IF(ISNUMBER(R[-1]C[1]),R[-1]C[1],0)"
            End If
            ' TODO allow other currencies than CHF
            .ListRows(1).Range(1, BL_AMOUNT_COL).FormulaR1C1 = "=[" & lblAmount & "]/VLOOKUP([" & lblDate & "],CHFtoEUR,2,TRUE)"
            .ListRows(1).Range(1, BL_BALANCE_COL).FormulaR1C1 = "=[" & lblBalance & "]/VLOOKUP([" & lblDate & "],CHFtoEUR,2,TRUE)"
        End If
    End With
    Set oBalanceTable = ws.ListObjects(tblName)
End Sub


Private Sub addDepositTable()
    Dim tblName As String
    tblName = Replace(Replace(Replace(LCase$(aId), " ", "_"), "é", "e"), "è", "e") & "_" & DEPOSIT_TABLE_NAME
    ' TODO: Don't create table if already exists
    Dim r As String
    r = "$G$10:$H$11"
    If aCurrency <> GetGlobalParam("DefaultCurrency") Then
        r = "$I$10:$J$11"
    End If
    ws.ListObjects.Add(xlSrcRange, Range(r), , xlYes).Name = tblName
    With ws.ListObjects(tblName)
        .TableStyle = "TableStyleMedium4"
        .ListColumns(BL_DATE_COL).Name = GetLabel("k.date")
        .ListColumns(BL_AMOUNT_COL).Name = GetLabel("k.amount")
    End With
    Set oDepositsTable = ws.ListObjects(tblName)
End Sub


Private Sub addInterestTable()
    Dim tblName As String
    tblName = Replace(Replace(Replace(LCase$(aId), " ", "_"), "é", "e"), "è", "e") & "_" & INTEREST_TABLE_NAME
    ' TODO: Don't create table if already exists
    Dim r As String
    r = "$G$1:$I$6"
    If aCurrency <> GetGlobalParam("DefaultCurrency") Then
        r = "$I$1:$K$6"
    End If
    ws.ListObjects.Add(xlSrcRange, Range(r), , xlYes).Name = tblName
    With ws.ListObjects(tblName)
        .TableStyle = "TableStyleMedium5"
        .ListColumns(IT_PERIOD_COL).Name = GetLabel("k.period")
        .ListColumns(IT_GROSS_INTEREST_COL).Name = GetLabel("k.grossYield")
        .ListColumns(IT_NET_INTEREST_COL).Name = GetLabel("k.netYield")
        .ListRows(1).Range(1, IT_PERIOD_COL).value = GetLabel("k.thisYear")
        .ListRows(2).Range(1, IT_PERIOD_COL).value = GetLabel("k.lastYear")
        .ListRows(3).Range(1, IT_PERIOD_COL).value = GetLabel("k.last3Years")
        .ListRows(4).Range(1, IT_PERIOD_COL).value = GetLabel("k.last5Years")
        .ListRows(5).Range(1, IT_PERIOD_COL).value = GetLabel("k.allTime")
    End With
    Set oInterestTable = ws.ListObjects(tblName)
End Sub


Public Sub FormatMe()
    ws.Cells.RowHeight = 13
    ws.Rows.font.size = 10
    ws.Activate
    Call formatBalanceTable
    Call formatDepositTable
    Call formatInterestTable
    Call formatAccountButtons
End Sub


Public Sub Sort()
    Call SortTable(oBalanceTable, GetLabel(DATE_KEY), xlAscending, GetLabel(AMOUNT_KEY), xlDescending)
End Sub

Public Function HasInterests() As Boolean
    HasInterests = Not (aType = GetLabel("k.accountStandard") Or aType = "Autres")
End Function

Public Function InterestPeriod() As Integer
    InterestPeriod = aInterestPeriod
End Function


'-------------------------------------------------
Public Function Exists() As Boolean
    Dim accounts As KeyedTable
    Set accounts = NewKeyedTable(Sheets(ACCOUNTS_SHEET).ListObjects(ACCOUNTS_TABLE))
    Exists = (SheetExists(aId) And accounts.KeyExists(aId))
End Function

Property Get Id() As String
    Id = aId
End Property
Property Get Number() As String
    Number = aNumber
End Property
Property Get Name() As String
    Name = aName
End Property
Property Get Bank() As String
    Bank = aBank
End Property
Property Get Availability() As String
    Availability = aAvail
End Property
Property Get Status() As String
    Status = aStatus
End Property
Public Function IsOpen() As Boolean
    IsOpen = (aStatus = GetLabel("k.accountOpen"))
End Function
Public Function IsClosed() As Boolean
    IsClosed = (aStatus = GetLabel("k.accountClosed"))
End Function
Public Function AccountCurrency() As String
    AccountCurrency = aCurrency
End Function
Public Function AccountType() As String
    AccountType = aType
End Function
Property Get IsInBudget() As Boolean
    inBudget = aInBudget
End Property
Property Get TaxRate() As Double
    TaxRate = aTaxRate
End Property

Property Get MyCurrency() As String
    MyCurrency = aCurrency
End Property

Private Sub AddBalanceRow()
    If oBalanceTable Is Nothing Then
        Exit Sub
    End If
    
    oBalanceTable.ListRows.Add
    nbRows = oTable.ListRows.Count
    
    col = TableColNbrFromName(oBalanceTable, GetLabel(DATE_KEY))
    oBalanceTable.ListColumns(col).DataBodyRange.Rows(nbRows).FormulaR1C1 = Date
    
    col = TableColNbrFromName(oBalanceTable, GetLabel(BALANCE_KEY))
    oBalanceTable.ListColumns(col).DataBodyRange.Rows(nbRows).value = oBalanceTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).value
    
    col = TableColNbrFromName(oBalanceTable, GetLabel(SUBCATEGORY_KEY))
    oBalanceTable.ListColumns(col).DataBodyRange.Rows(nbRows).value = oBalanceTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).value
    
    col = TableColNbrFromName(oBalanceTable, GetLabel(AMOUNT_KEY))
    oBalanceTable.ListColumns(col).DataBodyRange.Rows(nbRows).FormulaR1C1 = oBalanceTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).FormulaR1C1
    
    col = TableColNbrFromName(oBalanceTable, GetLabel(DESCRIPTION_KEY))
    oBalanceTable.ListColumns(col).DataBodyRange.Rows(nbRows).FormulaR1C1 = oBalanceTable.ListColumns(col).DataBodyRange.Rows(nbRows - 1).FormulaR1C1
End Sub



Public Function DepositsHistory() As Variant
    DepositsHistory = depositsAsArray(aId)
End Function


Private Function balanceHistory(Optional sampling As String = "Yearly") As Variant
    Dim histAll() As Variant
    Dim histSampled() As Variant
    Dim nbYears As Long
    Dim i As Long
    Dim j As Long
    Dim lastMonth As Long
    Dim lastYear As Long
    Dim lastBalance As Double
    Dim histSize As Long
    histAll = balanceAsArray
    histSize = UBound(histAll, 1)
    nbYears = Year(histAll(histSize, 1)) - Year(histAll(1, 1)) + 2
    ReDim histSampled(1 To nbYears, 1 To 2)
    lastMonth = 0
    lastYear = Year(histAll(1, 1)) - 1
    lastBalance = 0
    j = 1
    For i = 1 To histSize
        m = Month(histAll(i, BL_DATE_COL))
        y = Year(histAll(i, BL_DATE_COL))
        If sampling = "Monthly" And m <> lastMonth Then
            d = Day(histAll(i, BL_DATE_COL))
            histSampled(j, 1) = DateSerial(y, m, 1)
            If d <> 1 Then
                histSampled(j, 2) = lastBalance
            Else
                histSampled(j, 2) = histAll(i, BL_BALANCE_COL)
            End If
            j = j + 1
        ElseIf sampling = "Yearly" And y <> lastYear Then
            While y <> lastYear
                histSampled(j, 1) = DateSerial(lastYear, 12, 31)
                histSampled(j, 2) = lastBalance
                j = j + 1
                lastYear = lastYear + 1
            Wend
        End If
        firstEntry = False
        lastMonth = m
        lastBalance = histAll(i, BL_BALANCE_COL)
        lastYear = y
    Next i
    If sampling = "Yearly" Then
        histSampled(j, 1) = histAll(histSize, 1)
        histSampled(j, 2) = lastBalance
    End If
    balanceHistory = histSampled
End Function


'--------------------------------------------------------------------------
' Private methods
'--------------------------------------------------------------------------

'----------------------------------------------------------------------------
' Table as Tables
'----------------------------------------------------------------------------
Private Function accountTable(accountSection As String) As ListObject
    Dim i As Long
    For i = 1 To ws.ListObjects.Count
        If LCase$(ws.ListObjects(i).Name) Like "*_" & accountSection Then
            Set accountTable = ws.ListObjects(i)
            Exit For
        End If
    Next i
End Function

Property Get DepositTable() As ListObject
    If oDepositsTable Is Nothing Then
        Set oDepositsTable = accountTable(DEPOSIT_TABLE_NAME)
    End If
    Set DepositTable = oDepositsTable
End Property

Property Get BalanceTable() As ListObject
    If oBalanceTable Is Nothing Then
        Set oBalanceTable = accountTable(BALANCE_TABLE_NAME)
    End If
    Set BalanceTable = oBalanceTable
End Property

Property Get InterestTable() As ListObject
    If oInterestTable Is Nothing Then
        Set oInterestTable = accountTable(INTEREST_TABLE_NAME)
    End If
    Set InterestTable = oInterestTable
End Property

'----------------------------------------------------------------------------
' Table as Indexes
'----------------------------------------------------------------------------

Public Function BalanceTableIndex() As Long
    BalanceTableIndex = TableIndex(BALANCE_TABLE_NAME, ws)
End Function

Public Function DepositTableIndex() As Long
    DepositTableIndex = TableIndex(DEPOSIT_TABLE_NAME, ws)
End Function

Public Function InterestTableIndex() As Long
    InterestTableIndex = TableIndex(INTEREST_TABLE_NAME, ws)
End Function


Private Sub ExportMetadata(targetWs As Worksheet, nbrTransactions As Long, Optional nbrDeposits As Long = 0)
    ' Copy metadata on row 1
    targetWs.Range("A1") = "ExportDate=" & Format(Now(), "YYYY-mm-dd HH:MM:SS")
    targetWs.Range("B1") = "AccountId=" & aId
    targetWs.Range("C1") = "AccountNumber=" & aNumber
    targetWs.Range("D1") = "Bank=" & aBank
    targetWs.Range("E1") = "Availability=" & CStr(aAvail)
    targetWs.Range("F1") = "Currency=" & aCurrency
    targetWs.Range("G1") = "Type=" & aType
    targetWs.Range("H1") = "TaxRate=" & CStr(aTaxRate)
    targetWs.Range("I1") = "NbrTransactions=" & CStr(nbrTransactions)
    If nbrDeposits > 0 Then
        targetWs.Range("J1") = "NbrDeposits=" & CStr(nbrDeposits)
    End If
End Sub

'----------------------------------------------------------------------------
' Table as Arrays
'----------------------------------------------------------------------------
Private Function asArray(accountSection As String) As Variant
    Set asArray = Nothing
    Dim o As ListObject
    For Each o In ws.ListObjects
        If LCase$(o.Name) Like "*_" & accountSection Then
            asArray = GetTableAsArray(o)
            Exit For
        End If
    Next o
End Function

Private Function depositsAsArray() As Variant
    depositsAsArray = asArray(DEPOSIT_TABLE_NAME)
End Function

Private Function balanceAsArray() As Variant
    balanceAsArray = asArray(BALANCE_TABLE_NAME)
End Function

'----------------------------------------------------------------------------
' Private formatting functions
'----------------------------------------------------------------------------

Private Sub formatAccountButtons()
    If ws.Shapes.Count <= 0 Then
        Exit Sub
    End If
    Dim sbw As Integer
    sbw = 40
    Dim i As Long
    i = 0
    Dim s As Shape

    For Each btnData In Array( _
        "BtnHome," & BTN_HOME_TEXT & ",GoToSolde,Webdings,18,1,1,40" _
        , "BtnPrev5," & BTN_PREV_5_TEXT & ",GoBack5,Webdings,18,1,2,40" _
        , "BtnPrev," & BTN_PREV_TEXT & ",GoToPrev,Webdings,18,1,3,40" _
        , "BtnNext," & BTN_NEXT_TEXT & ",GoToNext,Webdings,18,1,4,40" _
        , "BtnNext5," & BTN_NEXT_5_TEXT & ",GoFwd5,Webdings,18,1,5,40" _
        , "BtnTop," & BTN_TOP_TEXT & ",scrollToTop,Webdings,18,1,6,40" _
        , "BtnBottom," & BTN_BOTTOM_TEXT & ",scrollToBottom,Webdings,18,1,7,40" _
        , "BtnSort," & BTN_SORT_TEXT & ",AccountSortHere,Webdings,18,2,1,40" _
        , "BtnImport," & BTN_IMPORT_TEXT & ",ImportAny,Webdings,18,2,2,40" _
        , "BtnAddEntry," & BTN_ADD_ROW_TEXT & ",AddSavingsRow,Arial,14,2,3,40" _
        , "BtnInterests," & Chr$(143) & ",InterestsCalcHere,Webdings,18,2,4,40" _
        , "BtnFormat," & BTN_FORMAT_TEXT & ",AccountFormatHere,Arial,18,2,5,80" _
        )
        values = Split(btnData, ",", -1, vbTextCompare)
        Set s = ShapeFind(ws, CStr(values(0)))
        If Not s Is Nothing Then
            Call BtnSetProperties(s, text:=CStr(values(1)), action:=CStr(values(2)), font:=CStr(values(3)), fontSize:=CInt(values(4)))
            Call ShapePlacement(s, BTN_HOME_X + (CInt(values(6)) - 1) * sbw, _
                BTN_HOME_Y + (CInt(values(5)) - 1) * BTN_HEIGHT, CInt(values(7)) - 1, BTN_HEIGHT - 1)
        End If
    Next btnData
    ws.Range("A1").Select
    
'    ElseIf (s.Type = msoFormControl) Then
'            ' Other buttons, move to row 3
'            r = 3 + (i Mod 6)
'            c = i \ 4
'            w = 2 * sbw
'            i = i + 1
'        End If
'        Call ShapePlacement(s, BTN_HOME_X + (c - 1) * sbw, BTN_HOME_Y + (r - 1) * BTN_HEIGHT, sbw - 1, BTN_HEIGHT - 1)
End Sub

Private Sub formatBalanceTable()
    If oBalanceTable Is Nothing Then
        Exit Sub
    End If
    oBalanceTable.Name = Replace(Replace(Replace(LCase$(accountId), " ", "_"), "é", "e"), "è", "e") & "_" & BALANCE_TABLE_NAME
    Call SetTableStyle(oBalanceTable, "TableStyleMedium2")
    Dim col As Long
    col = TableColNbrFromName(oBalanceTable, GetLabel(DATE_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
        Call SetTableColumnFormat(oBalanceTable, col, DATE_FORMAT)
    End If
    col = TableColNbrFromName(oBalanceTable, GetLabel(AMOUNT_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
        Call SetTableColumnFormat(oBalanceTable, col, EUR_FORMAT)
    End If
    col = TableColNbrFromName(oBalanceTable, "Montant CHF")
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 17, ws)
        Call SetTableColumnFormat(oBalanceTable, col, CHF_FORMAT)
    End If
    col = TableColNbrFromName(oBalanceTable, "Montant USD")
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
        Call SetTableColumnFormat(oBalanceTable, col, USD_FORMAT)
    End If
    col = TableColNbrFromName(oBalanceTable, GetLabel(BALANCE_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 18, ws)
        Call SetTableColumnFormat(oBalanceTable, col, EUR_FORMAT)
    End If
    col = TableColNbrFromName(oBalanceTable, "Solde CHF")
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 18, ws)
        Call SetTableColumnFormat(oBalanceTable, col, CHF_FORMAT)
    End If
    col = TableColNbrFromName(oBalanceTable, "Solde USD")
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 18, ws)
        Call SetTableColumnFormat(oBalanceTable, col, USD_FORMAT)
    End If
    col = TableColNbrFromName(oBalanceTable, GetLabel(DESCRIPTION_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 70, ws)
    End If
    col = TableColNbrFromName(oBalanceTable, GetLabel(SUBCATEGORY_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
    End If
    col = TableColNbrFromName(oBalanceTable, GetLabel(CATEGORY_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 15, ws)
    End If
    col = TableColNbrFromName(oBalanceTable, GetLabel(IN_BUDGET_KEY))
    If col <> 0 Then
        Call SetColumnWidth(Chr$(col + 64), 5, ws)
        Call SetColumnWidth(Chr$(col + 65), 5, ws)
    End If
    For Each colName In Array(GetLabel(AMOUNT_KEY), GetLabel(BALANCE_KEY))
        If TableColumnNameExists(oBalanceTable, CStr(colName)) Then
            Call TableColumnFormatIcons(oBalanceTable, CStr(colName))
        End If
        For Each curr In Array("EUR", "CHF", "USD")
            Dim cname As String
            cname = CStr(colName) & " " & CStr(curr)
            If TableColumnNameExists(oBalanceTable, cname) Then
                Call TableColumnFormatIcons(oBalanceTable, cname)
            End If
        Next curr
    Next colName
End Sub

Private Sub formatDepositTable()
    If oDepositsTable Is Nothing Then
        Exit Sub
    End If
    oDepositsTable.Name = Replace(Replace(Replace(LCase$(accountId), " ", "_"), "é", "e"), "è", "e") & "_" & DEPOSIT_TABLE_NAME
    Call SetTableStyle(oDepositsTable, "TableStyleMedium4")
    Call SetTableColumnFormat(oDepositsTable, 1, DATE_FORMAT)
    Call SetTableColumnFormat(oDepositsTable, 2, EUR_FORMAT)
End Sub

Private Sub formatInterestTable()
    If oInterestTable Is Nothing Then
        Exit Sub
    End If
    oInterestTable.Name = Replace(Replace(Replace(LCase$(accountId), " ", "_"), "é", "e"), "è", "e") & "_" & INTEREST_TABLE_NAME
    Call SetTableStyle(oInterestTable, "TableStyleMedium5")
    Call SetTableColumnFormat(oInterestTable, 2, INTEREST_FORMAT)
    Call SetTableColumnFormat(oInterestTable, 3, INTEREST_FORMAT)
End Sub



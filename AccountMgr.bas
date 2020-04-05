Attribute VB_Name = "AccountMgr"
Const NOT_AN_ACCOUNT = 0
Const DOMESTIC_ACCOUNT = 1
Const FOREIGN_ACCOUNT = 2
Const DOMESTIC_SHARES_ACCOUNT = 3
Const FOREIGN_SHARES_ACCOUNT = 4

Const ACCOUNT_CLOSED = 0
Const ACCOUNT_OPEN = 1

Const ACCOUNT_NAME_LABEL = "A1"
Const ACCOUNT_NAME_VALUE = "B1"
Const ACCOUNT_NBR_LABEL = "A2"
Const ACCOUNT_NBR_VALUE = "B2"
Const ACCOUNT_BANK_LABEL = "A3"
Const ACCOUNT_BANK_VALUE = "B3"
Const ACCOUNT_STATUS_LABEL = "A4"
Const ACCOUNT_STATUS_VALUE = "B4"
Const ACCOUNT_AVAIL_LABEL = "A5"
Const ACCOUNT_AVAIL_VALUE = "B5"
Const ACCOUNT_CURRENCY_LABEL = "A6"
Const ACCOUNT_CURRENCY_VALUE = "B6"
Const ACCOUNT_TYPE_LABEL = "A7"
Const ACCOUNT_TYPE_VALUE = "B7"
Const IN_BUDGET_LABEL = "A8"
Const IN_BUDGET_VALUE = "B8"


Sub CreateAccount()
    accountNbr = InputBox("Account number ?", "Account Number", "<accountNumber>")
    accountName = InputBox("Account name ?", "Account Name", "<accountName>")
    Sheets("Account Template").Visible = True
    Sheets("Account Template").Copy Before:=Sheets(1)
    Sheets("Account Template").Visible = False
    Sheets(1).name = accountName
    ' Sheets(1).Range("A1").Formula = "=VLOOKUP("k.account", TblKeys, LangId, FALSE)"
    Sheets(1).Range(ACCOUNT_NAME_VALUE).Value = accountName
    Sheets(1).Range(ACCOUNT_NBR_VALUE).Formula = "=VLOOKUP(B$1,TblAccounts,2,FALSE)"
    Sheets(1).Range(ACCOUNT_BANK_VALUE).Formula = "=VLOOKUP(B$1,TblAccounts,4,FALSE)"
    Sheets(1).Range(ACCOUNT_STATUS_VALUE).Formula = "=VLOOKUP(B$1,TblAccounts,6,FALSE)"
    Sheets(1).Range(ACCOUNT_AVAIL_VALUE).Formula = "=VLOOKUP(B$1,TblAccounts,5,FALSE)"
End Sub


Public Sub doForAllAccounts()
'
' Applies a given macro to all account sheets
'
'
    Call showAllSheets
    For Each ws In Worksheets
       ' Make sure the sheet is not anything else than an account
        If (isAnAccountSheet(ws) Or isTemplate(ws)) Then
            ws.Select
            ' Call macro here
        End If
    Next ws
    Call hideClosedAccounts
    Call hideTemplateAccounts
End Sub
'-------------------------------------------------
Public Sub formatAccountSheets()
'
'  Reformat all account sheets
'
   For Each ws In Worksheets
       ' Make sure the sheet is not anything else than an account
       If (isAnAccountSheet(ws) Or isTemplate(ws)) Then
            Dim name As String
            Dim acctype As String
            Dim acurrency As String
            name = ws.name
            acctype = accountType(name)
            acurrency = accountCurrency(name)
            Call SetColumnWidth("A", 15, name)
            ws.ListObjects(1).ListColumns(1).DataBodyRange.NumberFormat = "m/d/yyyy"
            Call SetColumnWidth("B", 20, name)
            Call SetColumnWidth("C", 20, name)
            If (acctype = "Standard") Then
                If (acurrency = "EUR") Then
                    Call SetColumnWidth("D", 70, name)
                    Call SetColumnWidth("E", 15, name)
                    Call SetColumnWidth("F", 15, name)
                    Call SetColumnWidth("G", 5, name)
                    Call SetColumnWidth("H", 5, name)
                    Call SetColumnWidth("I", 15, name)
                Else:
                    Call SetColumnWidth("D", 20, name)
                    Call SetColumnWidth("E", 70, name)
                    Call SetColumnWidth("F", 15, name)
                    Call SetColumnWidth("G", 15, name)
                    Call SetColumnWidth("H", 5, name)
                    Call SetColumnWidth("I", 5, name)
                    Call SetColumnWidth("J", 15, name)
                End If
            Else ' Shares accounts formatting
                Call SetColumnWidth("D", 70, name)
                Call SetColumnWidth("E", 20, name)
                If (acurrency = "EUR") Then
                   Call SetColumnWidth("F", 5, name)
                   Call SetColumnWidth("G", 20, name)
                   Call SetColumnWidth("H", 20, name)
                Else
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
                    row = i Mod 4
                    col = i \ 4
                    Call ShapePlacementXY(Shape, 300 + col * 100, 5 + row * 22, 400 + col * 100, 25 + row * 22)
                    i = i + 1
                End If
            Next Shape
          End If
       End If
   Next ws
   Call hideClosedAccounts
   Call hideTemplateAccounts
End Sub

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
Public Sub showClosedAccounts()
    For Each ws In Worksheets
        If (accountStatus(ws.name) = "Closed") Then
            ws.Visible = True
        End If
    Next ws
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
Public Sub showTemplateAccounts()
    For Each ws In Worksheets
        If (isTemplate(ws)) Then
            ws.Visible = True
        End If
    Next ws
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
        .ListFillRange = "Paramètres!$L$2:$L$" & CStr(Sheets("Paramètres").ListObjects("tblOpenAccounts").ListRows.Count + 1)
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
    oTable.Sort.SortFields.Add key:=Range(oTable.name & "[Date]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    oTable.Sort.SortFields.Add key:=Range(oTable.name & "[Montant]"), SortOn:=xlSortOnValues, Order:= _
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
        accountNumber = Sheets(accountName).Range(ACCOUNT_NBR_VALUE).Value
    Else
        accountNumber = ""
    End If
End Function
'-------------------------------------------------
Public Function accountBank(accountName As String) As String
    If (accountExists(accountName)) Then
        accountBank = Sheets(accountName).Range(ACCOUNT_BANK_VALUE).Value
    Else
        accountBank = ""
    End If
End Function

'-------------------------------------------------
Public Function accountStatus(accountName As String) As String
    If (accountExists(accountName)) Then
        accountStatus = Sheets(accountName).Range(ACCOUNT_STATUS_VALUE).Value
    Else
        accountStatus = ""
    End If
End Function
'-------------------------------------------------
Public Function accountAvailability(accountName As String) As String
    If (accountExists(accountName)) Then
        accountAvailability = Sheets(accountName).Range(ACCOUNT_AVAIL_VALUE).Value
    Else
        accountAvailability = ""
    End If
End Function
'-------------------------------------------------
Public Function accountCurrency(accountName As String) As String
    If (accountExists(accountName)) Then
        accountCurrency = Sheets(accountName).Range(ACCOUNT_CURRENCY_VALUE).Value
    Else
        accountCurrency = ""
    End If
End Function
'-------------------------------------------------
Public Function isAccountInBudget(accountName As String) As Boolean
    isAccountInBudget = (accountExists(accountName) And Sheets(accountName).Range(IN_BUDGET_VALUE).Value = "Yes")
End Function
'-------------------------------------------------
Public Function isOpen(accountName As String) As Boolean
    isOpen = (accountStatus(accountName) = "Open")
End Function

Public Function isClosed(accountName As String) As Boolean
    isClosed = Not isOpen(accountName)
End Function


'-------------------------------------------------
Public Function accountExists(accountName As String) As Boolean
    accountExists = (sheetExists(accountName) And Sheets(accountName).Range(ACCOUNT_NAME_LABEL) = "Nom Compte")
End Function
'-------------------------------------------------
Public Function isAnAccountSheet(ByVal ws As Worksheet) As Boolean
    isAnAccountSheet = (ws.Cells(1, 1).Value = getNamedVariableValue("accountIdentifier") And Not isTemplate(ws))
End Function

'-------------------------------------------------
Public Sub showAllSheets()
    For Each ws In Worksheets
        ws.Visible = True
    Next ws
End Sub

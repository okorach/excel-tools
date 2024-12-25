Attribute VB_Name = "MiscTools"
'
' Module MiscTools
'
' Miscellaneous tools that may be useful for many Excel apps
'
' See change log at bottom of file
'
' setCellFormat(cellName, format)
' SwapCellsXY(row1, col1, row2, col2, optional wsName)
' SwapCells(cellName1, cellName2)
' SetColumnWidth(colString As String, width As Double, Optional ws As Variant = "")



Public Sub FreezeDisplay()
    ' Freeze the Excel display so that macros run faster
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
End Sub

Public Sub UnfreezeDisplay()
    ' Unfreeze the Excel display (to be used after executing macros, or when intermediate display is needed)
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
End Sub

Public Sub ScrollToBottom()
    ' Autoscroll to bottom of transactions table
    If ActiveSheet.ListObjects(1).ListRows.Count > 10 Then
        ActiveWindow.ScrollRow = ActiveSheet.ListObjects(1).ListRows.Count - 10
    End If
End Sub

Public Sub ScrollToTop()
    ' Autoscroll to top of page
    ActiveWindow.ScrollRow = 10
End Sub

Public Function max(ByVal val1 As Double, ByVal val2 As Double) As Double
    If val1 > val2 Then
        max = val1
    Else
        max = val2
    End If
End Function
Public Function min(ByVal val1 As Double, ByVal val2 As Double) As Double
    If val1 < val2 Then
        min = val1
    Else
        min = val2
    End If
End Function

Public Function GetLastNonEmptyRow(Optional col As Integer = 1)
    Dim i As Long
    i = 1
    Do While LenB(Cells(i, col).value) > 0
        i = i + 1
    Loop
    GetLastNonEmptyRow = i - 1
End Function

Public Function SheetExists(wsName As String) As Boolean
    ' Returns whether a sheet with the given name exists
    SheetExists = False
    For Each ws In Worksheets
        If wsName = ws.name Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

Public Sub ShowAllSheets()
    ' Makes all worksheets visible
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Visible = True
    Next ws
End Sub

Public Sub GoToSolde()
    ' Navigate to the Solde sheet
    Sheets(BALANCE_PER_ACCOUNT_SHEET).Activate
End Sub

Public Sub GoToSheet(shift As Long)
    ' Change of active worksheet shift positions right (left if shift is negative)
    Dim sheetToGo As Long
    sheetToGo = ActiveSheet.index + shift
    nbSheets = Sheets.Count
    While (sheetToGo < Sheets.Count And Sheets(sheetToGo).Visible = xlSheetHidden And sheetToGo > 0)
        sheetToGo = sheetToGo + (shift / Abs(shift))
    Wend
    If sheetToGo = 0 Then
        sheetToGo = 1
    ElseIf sheetToGo > Sheets.Count Then
        sheetToGo = nbSheets
    End If
    Sheets(sheetToGo).Activate
End Sub

Public Sub GoToNext()
    Call GoToSheet(1)
End Sub

Public Sub GoToPrev()
    Call GoToSheet(-1)
End Sub

Public Sub GoBack5()
    Call GoToSheet(-5)
End Sub

Public Sub GoFwd5()
    Call GoToSheet(5)
End Sub

Public Function GetNamedVariableValue(varName As String, Optional wb As Workbook = Nothing)
    If wb Is Nothing Then
        GetNamedVariableValue = Names(varName).RefersToRange.value
    Else
        GetNamedVariableValue = wb.Names(varName).RefersToRange.value
    End If
End Function

Public Sub SetNamedVariableValue(varName As String, varValue As Variant, Optional wb As Workbook = Nothing)
    If wb Is Nothing Then
        Names(varName).RefersToRange.value = varValue
    Else
        wb.Names(varName).RefersToRange.value = varValue
    End If
End Sub

Public Function GetColName(key) As String
    GetColName = GetLabel(key)
End Function

Public Sub ErrorMessage(key1 As String, Optional key2 As String = vbNullString)
    msg = GetLabel(key1)
    If LenB(key2) > 0 Then
        ms = msg & ", " & GetLabel(key2)
    End If
    MsgBox (msg)
End Sub

Public Sub FreezeCell(r As String, Optional wsName As String = "")
    ' Replaces a cell that may contain a formula by the result of this formula
    ' (Used in situation where the cell value should no longer be changed when the parameters of formula can still change)
    Dim ws As Worksheet

    If LenB(wsName) = 0 Then
        Set ws = ActiveSheet
    Else
        Set ws = Worksheets(wsName)
    End If
        If (IsError(ws.Range(r).value)) Then
            ws.Range(r).value = 0
        Else
            ws.Range(r).Formula = Range(r).value
        End If
End Sub

Public Sub FreezeCellXY(row As Long, col As Long, Optional wsName As String = "")
    Call FreezeCell(Cells(row, col).Address(False, False), wsName)
End Sub

Public Sub SwapCells(cell1 As String, cell2 As String)
    ' Swaps the value of 2 cells, referenced by cell name (ex: "A1", "K12")
    Dim Temp As Variant
    Temp = Range(cell1).value
    Range(cell1).value = Range(cell2).value
    Range(cell2).value = Temp
End Sub

Public Sub SwapCellsXY(ByVal Row1 As Long, ByVal Col1 As Long, ByVal Row2 As Long, ByVal Col2 As Long, Optional wsName As String = "")
    ' Swaps the value of 2 cells, referenced by row and col number
    Dim ws As Worksheet
    Dim Temp As Variant
    If LenB(wsName) = 0 Then
        Set ws = ActiveSheet
    Else
        Set ws = Sheets(wsName)
    End If
    Temp = ws.Cells(Row1, Col1).value
    ws.Cells(Row1, Col1).value = ws.Cells(Row2, Col2).value
    ws.Cells(Row2, Col2).value = Temp
End Sub

Public Sub FreezeRegion(r As String, Optional wsName As String = "")
    ' Replaces a cell region that may contain formulas by the result of these formulas
    ' (Used in situation where the cell value should no longer be changed when the parameters of formula can still change)
    Dim ws As Worksheet
    If LenB(wsName) = 0 Then
        Set ws = ActiveSheet
    Else
        Set ws = Sheets(wsName)
    End If
    ws.Range(r).Copy
    ws.Range(r).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

Public Sub FreezeRegionXY(x1 As Long, y1 As Long, x2 As Long, y2 As Long, Optional wsName As String = "")
    Dim r As String
    r = Cells(x1, y1).Address(False, False) & ":" & Cells(x2, y2).Address(False, False)
    Call FreezeRegion(r, wsName)
End Sub

Public Function IsInArray(str As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, str)) > -1)
End Function

Sub SetColumnValidationRule(colObject As ListColumn, validationList As String)
    ' Enforces a data validation rule on zones of an excel sheet
    ' On a column of an Excel table
    With colObject.DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=validationList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
    End With
End Sub

Sub SetRangeValidationRule(aRange As Range, validationList As String)
    ' On an arbitrary range like Range("A4") (one cell) or Range("B7:K20") (a complete region)
    With aRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=validationList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
    End With
End Sub

Private Sub ListAllItemObjects()
    For Each pvt In ActiveSheet.PivotTables
        For Each fld In pvt.PivotFields
            For Each itm In fld.PivotItems
                MsgBox itm
            Next itm
        Next fld
    Next pvt
End Sub

'==============================================================================
'
'  Cell formatting functions
'
'==============================================================================

Sub reformatAmount(colObject As ListColumn)
    With colObject.DataBodyRange
        .style = "Normal"
        .NumberFormat = CHF_FORMAT
    End With
End Sub

Sub setCellFormat(cellName, Format)
    Range(cellName).NumberFormat = Format
End Sub

Public Sub SetColumnWidth(colString As String, width As Double, Optional ws As Worksheet = Empty)
    If IsEmpty(ws) Then
        Columns(colString & ":" & colString).ColumnWidth = width
    Else
        ws.Columns(colString & ":" & colString).ColumnWidth = width
    End If
End Sub

Public Sub SetRowHeight(rowString As String, height As Double, Optional ws As Worksheet = Empty)
    If IsEmpty(ws) Then
        Rows(rowString & ":" & rowString).RowHeight = height
    Else
        ws.Rows(rowString & ":" & rowString).RowHeight = height
    End If
End Sub

Public Sub SetRowFontSize(rowString As String, size As Double, Optional ws As Worksheet = Empty)
    If IsEmpty(ws) Then
        Rows(rowString & ":" & rowString).font.size = size
    Else
        ws.Rows(rowString & ":" & rowString).font.size = size
    End If
End Sub

Public Sub SetRangeStyle(rangeString As String, aStyle As String, Optional ws As Worksheet = Empty)
    If IsEmpty(ws) = 0 Then
        Range(rangeString).style = aStyle
    Else
        ws.Range(rangeString).style = aStyle
    End If
End Sub

'==============================================================================
'
'  Shape Management Tools
'
'==============================================================================

Public Sub ShapePlacementOnCell(oShape, oCell)
    '------------------------------------------------------------------------------
    ' Places a shape (a button for instance) exactly above a cell
    '------------------------------------------------------------------------------
    Call ShapePlacementOnCells(oShape, oCell, oCell)
End Sub
Public Sub ShapePlacementOnCells(oShape, oCell1, oCell2)
    '------------------------------------------------------------------------------
    ' Places a shape (a button for instance) exactly above a range of cells
    ' defined by its top left and bottom right cells
    '------------------------------------------------------------------------------
    With oShape
        .top = oCell1.top
        .left = oCell1.left
        .width = oCell2.left - oCell1.left + oCell2.width
        .height = oCell2.top - oCell1.top + oCell2.height
    End With
End Sub


'==============================================================================
'
'  Misc Tools
'
'==============================================================================


Sub SaveAsNewFile(filename, fileformat)
    '------------------------------------------------------------------------------
    ' Records a file with enforcement of the file format and filename
    '------------------------------------------------------------------------------
    Dim fileSaveName As Variant

    If (fileformat = "xlsm") Then
        fmt = xlOpenXMLWorkbookMacroEnabled
        myFilter = "Excel Macro-Enabled workbook (*.xlsm), *.xlsm"
    ElseIf (fileformat = "xltm") Then
        fmt = xlOpenXMLTemplateMacroEnabled
        myFilter = "Excel Macro-Enabled template (*.xltm), *.xltm"
    ElseIf (fileformat = "xlsx") Then
        fmt = xlOpenXMLWorkbook
        myFilter = "Excel workbook (*.xlsx), *.xlsx"
    ElseIf (fileformat = "xltx") Then
        fmt = xlOpenXMLTemplate
        myFilter = "Excel template (*.xltx), *.xltx"
    End If
    'NewFileName = "CalcSheet " & [CustomerName] & " " & [ProjectName] & " vXXX" & ".xlsm" 'wb.Sheets("Sheet1").Range("B18").Value & ".xlsm"
    fileSaveName = Application.GetSaveAsFilename _
            (InitialFileName:=filename, filefilter:=myFilter, Title:="Select folder")
    If Not fileSaveName = False Then
        ActiveWorkbook.SaveAs filename:=fileSaveName, fileformat:=fmt
    Else
        MsgBox "File NOT Saved."
    End If
End Sub

Public Sub DeleteAllButSheetOne()
    '------------------------------------------------------------------------------
    ' Records a file with enforcement of the file format and filename
    '------------------------------------------------------------------------------
    Application.DisplayAlerts = False
    While ActiveWorkbook.Sheets.Count > 1
       ActiveWorkbook.Sheets(2).Delete
    Wend
    Application.DisplayAlerts = True
End Sub

Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String

  Dim StringLen As Long: StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim Result(StringLen) As String
    Dim i As Long, CharCode As Long
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          Result(i) = Char
        Case 32
          Result(i) = Space
        Case 0 To 15
          Result(i) = "%0" & Hex$(CharCode)
        Case Else
          Result(i) = "%" & Hex$(CharCode)
      End Select
    Next i
    URLEncode = Join(Result, "")
  End If
End Function


Private Sub AdoptPivotSourceFormatting()
    'Mike Alexander
    'www.datapigtechnologies'
    'Be sure you start with your cursor inside a pivot table.
    Dim oPivotTable As PivotTable
    Dim oPivotFields As PivotField
    Dim oSourceRange As Range
    Dim strLabel As String
    Dim strFormat As String
    Dim i As Long

    On Error GoTo MyErr
    
    'Identify PivotTable and capture source Range
    'ActiveCell.PivotTable.Name
    Set oPivotTable = ActiveSheet.PivotTables(ActiveCell.PivotTable.name)
    Set oSourceRange = Range(Application.ConvertFormula(oPivotTable.SourceData, xlR1C1, xlA1))

    'Refresh PivotTable to synch with latest data
    oPivotTable.PivotCache.Refresh
    
    'Start looping through the columns in source range
    For i = 1 To oSourceRange.Columns.Count

    'Trap the column name and number format for first row of the column
        strLabel = oSourceRange.Cells(1, i).value
        strFormat = oSourceRange.Cells(2, i).NumberFormat

        'Now loop through the fields PivotTable data area
        For Each oPivotFields In oPivotTable.DataFields

        'Check for match on SourceName then appply number format if there is a match
        'If oPivotFields.SourceName = strLabel Then
        'oPivotFields.NumberFormat = strFormat

        'Bonus: Change the name of field to Source Column Name
        'oPivotFields.Caption = strLabel & " "
        'End If

        Next oPivotFields
    Next i
    For Each oPivotFields In oPivotTable.DataFields
        oPivotFields.NumberFormat = strFormat
    Next oPivotFields
    Exit Sub
    'Error stuff
MyErr:
    If Err.Number = 1004 Then
        MsgBox "You must place your cursor inside of a pivot table."
    Else
        MsgBox Err.Number & vbCrLf & Err.Description
    End If
End Sub


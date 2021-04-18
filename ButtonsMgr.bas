Attribute VB_Name = "ButtonsMgr"

Public Const BTN_HOME_NAME As String = "BtnHome"
Public Const BTN_HOME_TEXT As String = "9"

Public Const BTN_PREV_5_NAME As String = "BtnPrev5"
Public Const BTN_PREV_5_TEXT As String = "7"

Public Const BTN_PREV_NAME As String = "BtnPrev"
Public Const BTN_PREV_TEXT As String = "3"

Public Const BTN_NEXT_NAME As String = "BtnNext"
Public Const BTN_NEXT_TEXT As String = "4"

Public Const BTN_NEXT_5_NAME As String = "BtnNext5"
Public Const BTN_NEXT_5_TEXT As String = "8"

Public Const BTN_BOTTOM_NAME As String = "BtnBottom"
Public Const BTN_BOTTOM_TEXT As String = "6"

Public Const BTN_TOP_NAME As String = "BtnTop"
Public Const BTN_TOP_TEXT As String = "5"

Public Const BTN_SORT_NAME As String = "BtnSort"
Public Const BTN_SORT_TEXT As String = "~"

Public Const BTN_IMPORT_NAME As String = "BtnImport"
Public Const BTN_IMPORT_TEXT As String = "G"

Public Const BTN_ADD_ROW_NAME As String = "BtnAddEntry"
Public Const BTN_ADD_ROW_TEXT As String = "+1"

Public Const BTN_FORMAT_NAME As String = "BtnFormat"
Public Const BTN_FORMAT_TEXT As String = "Format"

Public Const BTN_INTERESTS_NAME As String = "BtnInterests"
'Public Const BTN_INTERESTS_TEXT As String = Chr$(143)

Public Const BTN_HOME_X As Integer = 200
Public Const BTN_HOME_Y As Integer = 10
Public Const BTN_HEIGHT As Integer = 30


Sub SetBtnMacro()
    ActiveSheet.Shapes("BtnHome").Select
    Selection.OnAction = "ThisWorkbook.GoToSolde"
End Sub

Public Sub BtnSetProperties(oBtn As Shape, Optional font As String = vbNullString, Optional text As String = vbNullString, _
                             Optional fontStyle As String = vbNullString, Optional fontSize As Integer = 0, Optional action As String = vbNullString)
    oBtn.Select
    If text <> vbNullString Then
        Selection.Characters.text = text
    End If
    If font <> vbNullString Then
        Selection.Characters.font.name = font
    End If
    If fontSize <> 0 Then
        Selection.Characters.font.size = fontSize
    End If
    If fontStyle <> vbNullString Then
        Selection.Characters.font.fontStyle = style
    End If
    If action <> vbNullString Then
        Selection.OnAction = action
    End If
    'With Selection.Characters().font
    '   .Name = font
    '   .fontStyle = "Normal"
    '   .size = 18
    '   .Strikethrough = False
    '   .Superscript = False
    '   .Subscript = False
    '   .OutlineFont = False
    '   .Shadow = False
    '   .Underline = xlUnderlineStyleNone
    '   .ColorIndex = xlAutomatic
    '   .TintAndShade = 0
    '   .ThemeFont = xlThemeFontNone
    'End With
End Sub

Public Function BtnAdd(ws As Worksheet, name As String, action As String, Optional text As String = vbNullString, _
    Optional size As Integer = 0, Optional font As String = vbNullString, Optional fontSize As Integer = 18, _
    Optional X As Integer = 10, Optional y As Integer = 10, Optional w As Integer = 30, Optional h As Integer = 20) As Shape
    ws.Buttons.Add(X, y, w, h).Select
    Set BtnAdd = ws.Shapes(ws.Shapes.Count)
    BtnAdd.name = name
    If text = vbNullString Then
        text = name
    End If
    Call BtnSetProperties(BtnAdd, text:=text, font:=font, fontSize:=fontSize, action:=action)
End Function


Public Sub BtnAddByString(ws As Worksheet, stringData As String)
    values = Split(stringData, ",", -1, vbTextCompare)
    If Not ShapeExist(ws, CStr(values(0))) Then
        Dim s As Shape
        Set s = BtnAdd(ws:=ws, name:=CStr(values(0)), text:=CStr(values(1)), action:=CStr(values(2)), font:=CStr(values(3)), fontSize:=CInt(values(4)))
        Call ShapePlacement(s, BTN_HOME_X + (CInt(values(6)) - 1) * sbw, _
            BTN_HOME_Y + (CInt(values(5)) - 1) * BTN_HEIGHT, CInt(values(7)) - 1, BTN_HEIGHT - 1)
    End If
End Sub

Public Sub BtnAddByStringArray(ws As Worksheet, btnArr As Variant)
    For Each btnData In btnArr
        Call BtnAddByString(ws, CStr(btnData))
    Next btnData
End Sub


'------------------------------------------------------------------------------
' Places a shape (a button for instance) on given X, Y coordinates
'------------------------------------------------------------------------------
Public Sub ShapePlacement(oShape As Shape, Optional left As Integer = -1, Optional top As Integer = -1, _
                        Optional width As Integer = -1, Optional height As Integer = -1)
    With oShape
        If X >= 0 Then
            .left = left
        End If
        If top >= 0 Then
            .top = top
        End If
        If width >= 0 Then
            .width = width
        End If
        If height >= 0 Then
            .height = height
        End If
    End With
End Sub

Public Function ShapeFind(ws As Worksheet, name As String) As Shape
    For Each s In ws.Shapes
        If s.name = name Then
            Set ShapeFind = s
            Exit For
        End If
    Next s
    ' Return Nothing if not found
End Function

Public Function ShapeExist(ws As Worksheet, name As String) As Boolean
    ShapeExist = Not (ShapeFind(ws, name) Is Nothing)
End Function

Attribute VB_Name = "ButtonsMgr"
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
        Selection.Characters.font.Name = font
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

Public Sub BtnAdd(ws As Worksheet, Name As String, action As String, Optional text As String = vbNullString, _
    Optional size As Integer = 0, Optional font As String = vbNullString, Optional fontSize As Integer = 18, _
    Optional X As Integer = 10, Optional y As Integer = 10, Optional w As Integer = 30, Optional h As Integer = 20)
    Dim oBtn As Shape
    ws.Buttons.Add(X, y, w, h).Select
    Set oBtn = ws.Shapes(ws.Shapes.Count)
    oBtn.Name = Name
    If text = vbNullString Then
        text = Name
    End If
    Call BtnSetProperties(oBtn, text:=text, font:=font, fontSize:=fontSize, action:=action)
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

Public Function ShapeFind(ws As Worksheet, Name As String) As Shape
    For Each s In ws.Shapes
        If s.Name = Name Then
            Set ShapeFind = s
            Exit For
        End If
    Next s
    ' Return Nothing if not found
End Function

Public Function ShapeExist(ws As Worksheet, Name As String) As Boolean
    ShapeExist = Not (ShapeFind(ws, Name) Is Nothing)
End Function

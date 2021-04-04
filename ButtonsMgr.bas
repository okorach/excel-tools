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

Public Sub BtnAdd(ws As Worksheet, name As String, action As String, text As String, _
    Optional size As Integer = 0, Optional font As String = vbNullString, Optional fontSize As Integer = 18, _
    Optional x As Integer = 10, Optional y As Integer = 10, Optional w As Integer = 30, Optional h As Integer = 20)
    Dim oBtn As Shape
    ws.Buttons.Add(x, y, w, h).Select
    Set oBtn = ws.Shapes(ws.Shapes.Count)
    oBtn.name = name
    Call BtnSetProperties(oBtn, text:=text, font:=font, fontSize:=fontSize, action:=action)
End Sub



Attribute VB_Name = "ButtonsMgr"
Sub SetBtnMacro()
    ActiveSheet.Shapes("BtnHome").Select
    Selection.OnAction = "ThisWorkbook.GoToSolde"
End Sub

Public Sub SetBtnAttributes(oBtn As Shape, Optional font As String = vbNullString, Optional text As String = vbNullString, _
                             Optional fontStyle As String = vbNullString, Optional size As Integer = 0, Optional action As String = vbNullString)
    oBtn.Select
    If text <> vbNullString Then
        Selection.Characters.text = text
    End If
    If font <> vbNullString Then
        Selection.Characters.font.name = font
    End If
    If size <> 0 Then
        Selection.Characters.font.size = size
    End If
    If fontStyle <> vbNullString Then
        Selection.Characters.font.fontStyle = style
    End If
    If action <> vbNullString Then
        Selection.OnAction = action
    End If
    'With Selection.Characters().font
    '    .Name = font
    '    .fontStyle = "Normal"
    '    .size = 18
        '.Strikethrough = False
        '.Superscript = False
        '.Subscript = False
        '.OutlineFont = False
        '.Shadow = False
        '.Underline = xlUnderlineStyleNone
        '.ColorIndex = xlAutomatic
        '.TintAndShade = 0
        '.ThemeFont = xlThemeFontNone
    'End With
End Sub


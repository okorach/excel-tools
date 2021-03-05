Attribute VB_Name = "ButtonsMgr"
Sub SetBtnMacro()
Attribute SetBtnMacro.VB_ProcData.VB_Invoke_Func = "b\n14"
    ActiveSheet.Shapes("BtnInterests").Select
    Selection.OnAction = "BtnAccountInterests"
End Sub

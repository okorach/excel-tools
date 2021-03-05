Attribute VB_Name = "ButtonsMgr"
Sub SetBtnMacro()
Attribute SetBtnMacro.VB_ProcData.VB_Invoke_Func = "b\n14"
    ActiveSheet.Shapes("BtnHome").Select
    Selection.OnAction = "ThisWorkbook.GoToSolde"
End Sub

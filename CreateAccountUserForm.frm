VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateAccountUserForm 
   Caption         =   "Create Account"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5280
   OleObjectBlob   =   "CreateAccountUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateAccountUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CreateAccountFormCancelBtn_Click()
    CreateAccountUserForm.Hide
End Sub

Private Sub UserForm_Initialize()
    For Each row In Sheets(PARAMS_SHEET).ListObjects(ACCOUNT_TYPES_TABLE).ListRows
        FormItemAccountType.AddItem CStr(row.Range(1, 1))
    Next row
    For Each row In Sheets(PARAMS_SHEET).ListObjects(CURRENCIES_TABLE).ListRows
        CreateAccountFormCurrency.AddItem CStr(row.Range(1, 1))
    Next row
    CreateAccountFormCurrency.text = GetGlobalParam("DefaultCurrency")
    thisYear = Year(Now())
    For y = thisYear To 2050
        CreateAccountFormAvailability.AddItem y
    Next y
    CreateAccountFormAvailability.text = thisYear
End Sub


Private Sub ValidateBtn_Click()
    Dim tax As Double
    CreateAccountUserForm.Hide
    Dim taxStr As String
    taxStr = Trim$(CreateAccountFormTaxRate.value)
    If Len(taxStr) = 0 Then
        tax = 0
    Else
        taxStr = left$(taxStr, Len(taxStr) - 1)
        tax = CDbl(taxStr) / 100
    End If
    Dim oAccount As Account
    Set oAccount = NewAccount(aId:=CStr(FormItemAccountName.value), aNbr:=CStr(CreateAccountFormAccountNbr.value), _
        aBank:=CStr(CreateAccountFormBank.value), aCur:=CStr(CreateAccountFormCurrency.value), aType:=CStr(FormItemAccountType.value), _
        aAvail:=CInt(CreateAccountFormAvailability.value), aInB:=CreateAccountFormInBudget, aTax:=tax)
End Sub


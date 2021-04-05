VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateAccountUserForm 
   Caption         =   "Create Account"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5355
   OleObjectBlob   =   "CreateAccountUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateAccountUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    For Each row In Sheets(PARAMS_SHEET).ListObjects(ACCOUNT_TYPES_TABLE).ListRows
        FormItemAccountType.AddItem CStr(row.Range(1, 1))
    Next row
    For Each row In Sheets(PARAMS_SHEET).ListObjects(CURRENCIES_TABLE).ListRows
        CreateAccountFormCurrency.AddItem CStr(row.Range(1, 1))
    Next row
    CreateAccountFormCurrency.text = GetGlobalParam("DefaultCurrency")
    For y = 2020 To 2050
        CreateAccountFormAvailability.AddItem y
    Next y
    CreateAccountFormAvailability.text = "2021"
End Sub


Private Sub ValidateBtn_Click()
    Dim accName As String, accType As String
    accType = "Courant"
    CreateAccountUserForm.Hide
    Call AccountCreate(accountId:=CStr(FormItemAccountName.value), accCurrency:=CStr(CreateAccountFormCurrency.value), _
        accType:=CStr(FormItemAccountType.value), bank:=CStr(CreateAccountFormBank.value), accNumber:=CStr(CreateAccountFormAccountNbr.value), _
        avail:=CInt(CreateAccountFormAvailability.value))
End Sub

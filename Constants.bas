Attribute VB_Name = "Constants"
Public Const INTEREST_CALC_SHEET As String = "Interests"
Public Const MERGE_SHEET As String = "Comptes Merge"
Public Const ACCOUNTS_SHEET As String = "Comptes"
Public Const BALANCE_SHEET As String = "Solde"
Public Const ACCOUNTS_TABLE As String = "tblAccounts"
Public Const OPEN_ACCOUNTS_TABLE As String = "tblOpenAccounts"
Public Const ACCOUNT_TYPES_TABLE As String = "TblAccountTypes"
Public Const ACCOUNT_MERGE_TABLE As String = "AccountsMerge"
 
 
Public Const BALANCE_TABLE_NAME As String = "balance"
Public Const DEPOSIT_TABLE_NAME As String = "deposit"
Public Const INTEREST_TABLE_NAME As String = "interest"

Public Const INTEREST_FORMAT As String = "0.0%"
Public Const CHF_FORMAT = "#,###,##0.00"" CHF "";-#,###,##0.00"" CHF "";0.00"" CHF """
Public Const EUR_FORMAT = "#,###,##0.00"" € "";-#,###,##0.00"" € "";0.00"" € """
Public Const USD_FORMAT = "#,###,##0.00"" $ "";-#,###,##0.00"" $ "";0.00"" $ """
Public Const DATE_FORMAT = "m/d/yyyy"



' Constants defining structure of global accounts table
Public Const ACCOUNT_KEY_COL As Integer = 1
Public Const ACCOUNT_NBR_COL As Integer = 2
Public Const ACCOUNT_NAME_COL As Integer = 3
Public Const ACCOUNT_BANK_COL As Integer = 4
Public Const ACCOUNT_AVAIL_COL As Integer = 5
Public Const ACCOUNT_STATUS_COL As Integer = 6
Public Const ACCOUNT_CURRENCY_COL As Integer = 7
Public Const ACCOUNT_TYPE_COL As Integer = 8
Public Const ACCOUNT_BUDGET_COL As Integer = 9
Public Const ACCOUNT_TAX_COL As Integer = 10



' Constants defining structure of account balance table
Public Const BL_DATE_COL As Integer = 1
Public Const BL_AMOUNT_COL As Integer = 2
Public Const BL_BALANCE_COL As Integer = 3
Public Const BL_FOREIGN_AMOUNT_COL = 4
Public Const BL_FOREIGN_BALANCE_COL = 5
Public Const BL_DESC_COL = 4
Public Const BL_SUBCATEG_COL = 5
Public Const BL_CATEG_COL = 6
Public Const BL_BUDGET_COL = 7
Public Const BL_FOREIGN_OFFSET = 2

' Constants defining structure of account deposits table
Public Const DP_DATE_COL As Integer = 1
Public Const DP_AMOUNT_COL As Integer = 2

' Constants defining structure of account interests table
Public Const IT_PERIOD_COL As Integer = 1
Public Const IT_GROSS_INTEREST_COL As Integer = 2
Public Const IT_NET_INTEREST_COL As Integer = 3


Public Const AVAILABILITY_IMMEDIATE = 0

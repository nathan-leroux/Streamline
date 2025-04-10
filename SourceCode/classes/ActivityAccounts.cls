VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActivityAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ActivityAccounts
' class for accounts
' includes possibly multiple name, balance, prn

Option Explicit
Implements ActivityModule

Private p_accounts As Collection
Private p_lookup As Scripting.Dictionary
Private p_coll_lookup As Scripting.Dictionary

Private Const ACC_NAME As Integer = 1
Private Const ACC_BAL As Integer = 2
Private Const ACC_PRN As Integer = 3


Private Sub Class_Initialize()
    Set p_accounts = New Collection
    Set p_lookup = New Scripting.Dictionary
    Set p_coll_lookup = New Scripting.Dictionary
End Sub


Private Function ActivityModule_populate(inst_file As InstFile) As Variant
    Dim new_account As Collection
    Dim parser As New TemplateParser
    
    Do
        Set new_account = New Collection
        
        ' the order these are added are important, update helper functions if changed
        Call new_account.Add(parser.get_accname(inst_file))
    
        Call new_account.Add(parser.get_currency(inst_file))
    
        Call new_account.Add(inst_file.get_next_inst("arg").inst_value)
    
        Call p_accounts.Add(new_account)
        
    Loop While inst_file.next_inst_type() = "arg"
    
    Call fill_lookup
End Function


Private Sub fill_lookup()
    Dim account As Collection
    Dim iter_obj As Object
    Dim iter_var As Variant
    
    Dim name_full_coll As New Collection
    Dim name_short_coll As New Collection
    Dim name_abrev_coll As New Collection
    
    Dim format_bal As New Collection
    
    Call p_lookup.Add("total", Format(collection_sum(collection_columnise(p_accounts, ACC_BAL)), "Currency"))
    
    If p_accounts.Count() = 1 Then
        Set account = p_accounts(1)
        
        Call p_lookup.Add("is_multi", False)
        
        Call p_lookup.Add("full_name", account(ACC_NAME).full)
        Call p_lookup.Add("short_name", account(ACC_NAME).short)
        Call p_lookup.Add("abrev_name", account(ACC_NAME).abrev)
        Call p_lookup.Add("payment_number", account(ACC_PRN))
        Call p_lookup.Add("balance", Format(account(ACC_BAL), "Currency"))
        
    Else
        Call p_lookup.Add("is_multi", True)
        
        For Each iter_obj In collection_columnise(p_accounts, ACC_NAME)
            Call name_full_coll.Add(iter_obj.full)
            Call name_short_coll.Add(iter_obj.short)
            Call name_abrev_coll.Add(iter_obj.abrev)
        Next iter_obj
        
        For Each iter_var In collection_columnise(p_accounts, ACC_BAL)
            Call format_bal.Add(Format(iter_var, "Currency"))
        Next iter_var
        
        Call p_coll_lookup.Add("full_name", name_full_coll)
        Call p_coll_lookup.Add("short_name", name_short_coll)
        Call p_coll_lookup.Add("abrev_name", name_abrev_coll)
        
        Call p_coll_lookup.Add("payment_number", collection_columnise(p_accounts, ACC_PRN))
        
        Call p_coll_lookup.Add("balance", format_bal)
    End If
End Sub


Private Function ActivityModule_attach(act As Activity) As Variant
    Call act.add_module("acc", Me)
End Function


Private Function ActivityModule_search(query As String) As Variant
    ActivityModule_search = p_lookup(query)
End Function


Private Function ActivityModule_searchcoll(query As String) As Collection
    Set ActivityModule_searchcoll = p_coll_lookup(query)
End Function


Private Function ActivityModule_str() As String
    Dim result As String
    Dim account As Object
    
    result = "(accounts:" & Chr(TEXT_NEWLINE)
    
    For Each account In p_accounts
        result = result & Chr(TEXT_TAB) & "name=" & account(ACC_NAME).full & ", balance=" & account(ACC_BAL)
        result = result & ", prn=" & account(ACC_PRN) & Chr(TEXT_NEWLINE)
    Next account
    
    result = result & ")"

    ActivityModule_str = result
End Function

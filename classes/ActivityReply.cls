VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActivityReply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ActivityReply
' module to represent the associated reply activity
' contains id, activity reference and client date

Option Explicit
Implements ActivityModule

Private p_id As String
Private p_ref As String
Private p_date As TemplateDate
Private p_lookup As Scripting.Dictionary
Private p_coll_lookup As Scripting.Dictionary


Private Sub Class_Initialize()
    Set p_lookup = New Scripting.Dictionary
    Set p_coll_lookup = New Scripting.Dictionary
End Sub


Private Function ActivityModule_populate(inst_file As InstFile) As Variant
    Dim todays_date As New TemplateDate
    
    p_id = inst_file.get_next_inst("arg").inst_value
    
    p_ref = inst_file.get_next_inst("arg").inst_value
    
    Call todays_date.populate_from_date(Now())
    Set p_date = todays_date
    
    Call fill_lookup
End Function


Private Sub fill_lookup()
    Call p_lookup.Add("id", p_id)
    Call p_lookup.Add("ref", p_ref)
    Call p_lookup.Add("date_short", p_date.short)
    Call p_lookup.Add("date_full", p_date.full)
End Sub


Private Function ActivityModule_attach(act As Activity) As Variant
    Call act.add_module("rpy", Me)
End Function


Private Function ActivityModule_search(query As String) As Variant
    ActivityModule_search = p_lookup(query)
End Function


Private Function ActivityModule_searchcoll(query As String) As Collection
    Set ActivityModule_searchcoll = p_coll_lookup(query)
End Function


Private Function ActivityModule_str() As String
    Dim result As String
    
    result = "(reply: id=" & p_id & ", ref=" & p_ref & ")"
    
    ActivityModule_str = result
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActivityOutbound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ActivityOutbound
' outbound class
' contains outcome code and flags for response

Option Explicit
Implements ActivityModule

Private p_outcome_code As Long
Private p_outcome_flags As Collection
Private p_lookup As Scripting.Dictionary
Private p_coll_lookup As Scripting.Dictionary


Private Sub Class_Initialize()
    Set p_outcome_flags = New Collection
    Set p_lookup = New Scripting.Dictionary
    Set p_coll_lookup = New Scripting.Dictionary
End Sub


Private Function ActivityModule_populate(inst_file As InstFile) As Variant
    p_outcome_code = parse_outcome_code(inst_file)
    
    Do While inst_file.next_inst_type = "arg"
        Call p_outcome_flags.Add(parse_outcome_flag(inst_file))
    Loop
    
    Call fill_lookup
End Function


Private Sub fill_lookup()
    Dim is_approved As Long
    If p_outcome_code = OutcomeCode.yes Or p_outcome_code = OutcomeCode.YesApproval Then
        is_approved = True
    Else
        is_approved = False
    End If
    
    Call p_lookup.Add("is_details", collection_exists(p_outcome_flags, OutcomeFlag.Details))

    Call p_lookup.Add("is_approved", is_approved)
    Call p_lookup.Add("code", p_outcome_code)
End Sub


Private Function ActivityModule_attach(act As Activity) As Variant
    Call act.add_module("out", Me)
End Function


Private Function ActivityModule_search(query As String) As Variant
    ActivityModule_search = p_lookup(query)
End Function


Private Function ActivityModule_searchcoll(query As String) As Collection
    Set ActivityModule_searchcoll = p_coll_lookup(query)
End Function


Private Function ActivityModule_str() As String
    Dim result As String
    Dim line As Variant
    
    result = "(outcome: code=" & p_outcome_code & ", flags=" & Chr(TEXT_NEWLINE)
    
    For Each line In p_outcome_flags
        result = result & Chr(TEXT_TAB) & line & Chr(TEXT_NEWLINE)
    Next line
    
    result = result & ")"
    
    ActivityModule_str = result
End Function


Private Function parse_outcome_code(inst_file As InstFile) As Long
    Dim input_code As String
    Dim err_msg As String
    
    input_code = inst_file.get_next_inst("arg").inst_value
    
    Select Case input_code
        Case "y"
            parse_outcome_code = OutcomeCode.yes
        Case "ya"
            parse_outcome_code = OutcomeCode.YesApproval
        Case "n"
            parse_outcome_code = OutcomeCode.no
        Case "na"
            parse_outcome_code = OutcomeCode.NoApproval
        Case Else
            err_msg = "Invalid Outcome Code: '" & input_code & "'"
            Call throw_exception(inst_file.throw_last_inst(err_msg), ActivityError.BadInputError)
    End Select
    
End Function


Private Function parse_outcome_flag(inst_file As InstFile)
    Dim input_flag As String
    Dim err_msg As String
    
    input_flag = inst_file.get_next_inst("arg").inst_value
    
   Select Case input_flag
        Case "details"
            parse_outcome_flag = OutcomeFlag.Details
        Case Else
            err_msg = "Invalid Outcome Flag: '" & input_flag & "'"
            Call throw_exception(inst_file.throw_last_inst(err_msg), ActivityError.BadInputError)
    End Select
End Function




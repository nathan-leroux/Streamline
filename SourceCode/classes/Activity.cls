VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Activity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Main datastructure for activity
'Everything else connects to here

Private p_modules As Scripting.Dictionary
Private p_lookup As Scripting.Dictionary

Private p_note_path As String
Private p_letter_paths As Collection


Private Sub Class_Initialize()
    Set p_modules = New Scripting.Dictionary
    Set p_lookup = New Scripting.Dictionary
End Sub


Public Function populate(inst_file As InstFile)
    Dim module As ActivityModule
    Dim current_inst As Instruction
    Dim ex_msg As String

    Do While inst_file.next_inst_type() <> "eof"
        Set current_inst = inst_file.get_next_inst("cmd")
        Set module = parse_module(current_inst.inst_value)
        
        If module Is Nothing Then
            ex_msg = "'" & current_inst.inst_value & "' is not a valid command"
            Call throw_exception(inst_file.throw_last_inst(ex_msg), ActivityError.BadInputError)
        End If
        
        Call module.populate(inst_file)
        Call module.attach(Me)
    Loop
    
    Call fill_lookup
    Call fill_paths
End Function


Private Sub fill_lookup()
    Call p_lookup.Add("is_not", p_modules.Exists("not"))
    Call p_lookup.Add("is_acc", p_modules.Exists("acc"))
    Call p_lookup.Add("is_ovr", p_modules.Exists("ovr"))
End Sub


Private Function parse_module(cmd As String) As ActivityModule
    Dim result As ActivityModule
    
    Select Case cmd
        Case "in"
            Set result = New ActivityInbound
         
        Case "adr"
            Set result = New ActivityAddress
            
        Case "res"
            Set result = New ActivityReason
            
        Case "not"
            Set result = New ActivityNote
            
        Case "acc"
            Set result = New ActivityAccounts
            
        Case "out"
            Set result = New ActivityOutbound
            
        Case "rpy"
            Set result = New ActivityReply
            
        Case Else
            Set result = Nothing
    End Select
    
    Set parse_module = result
End Function


Public Function add_module(module_type As String, module As ActivityModule)
    Call p_modules.Add(module_type, module)
End Function


Public Function search(query As String) As Variant
    Dim query_array() As String
    Dim module As String
    Dim argument As String
    Dim result As String
    
    query_array = Split(query, ".")
    
    If array_len(query_array) <> 2 Then
        Call err.Raise(ActivityError.BadDocumentError, Description:="Invalid variable name: '" & query & "'")
    End If
    
    module = Split(query, ".")(0)
    argument = Split(query, ".")(1)
    result = ""
    
    If module = "act" Then
        result = p_lookup(argument)
    Else
        result = get_module(module).search(argument)
    End If
    
    search = result
End Function


Public Function searchcoll(query As String) As Collection
    Dim query_array() As String
    Dim module As String
    Dim argument As String
    Dim result As String
    
    query_array = Split(query, ".")
    
    If array_len(query_array) <> 2 Then
        Call err.Raise(ActivityError.BadDocumentError, Description:="Invalid var name: " & query & "'")
    End If
    
    module = Split(query, ".")(0)
    argument = Split(query, ".")(1)
    result = ""
    
    If get_module(module).searchcoll(argument) Is Nothing Then
        Call err.rasie(ActivityError.BadDocumentError, Description:="Invalid var name: " & query)
    Else
        Set searchcoll = get_module(module).searchcoll(argument)
    End If
End Function


Public Property Get note_path() As String
    note_path = p_note_path
End Property


Public Property Get letter_paths() As Collection
    Set letter_paths = p_letter_paths
End Property


Public Function str() As String
    Dim result As String
    Dim module As Variant
    Dim letter As Variant
    Dim dummy As ActivityModule

    result = ""
    
    result = "(filepaths:" & Chr(TEXT_NEWLINE)
    result = result & "note  " & p_note_path & Chr(TEXT_NEWLINE)
    result = result & "reply  " & p_reply_path & Chr(TEXT_NEWLINE)
    result = result & "letters:" & Chr(TEXT_NEWLINE)
    
    For Each letter In p_letter_paths
        result = result & letter & Chr(TEXT_NEWLINE)
    Next letter
    
    result = result & ")" & Chr(TEXT_NEWLINE)
    
    For Each module In p_modules.Items()
        Set dummy = module
        result = str_append(result, dummy)
    Next module
    
    str = result
End Function


Public Sub fill_paths()
    Dim outcome_code As String
    
    outcome_code = safe_search("out", "code")
    
    Set p_letter_paths = Nothing
    
    Select Case CLng(outcome_code)
        Case OutcomeCode.yes
            p_note_path = "example_note_approved"
            Set p_letter_paths = fill_letter_paths()
            
        Case OutcomeCode.no
            p_note_path = "example_note_denied"
            Set p_letter_paths = fill_letter_paths()
            
        Case OutcomeCode.YesApproval
            p_note_path = "example_note_approved_escalation"
            
        Case OutcomeCode.NoApproval
            p_note_path = "example_note_denied_escalation"
            
        Case Else
            Call throw_exception("Outcome code not implemented: " & outcome_code & "'", ActivityError.BadInputError)
    End Select
End Sub


Private Function fill_letter_paths() As Collection
    Dim outcome_string As String
    Dim pay_string As String
    Dim result_paths As New Collection
    
    If Not Me.search("act.is_acc") Then
        pay_string = "_nopay"
    ElseIf Not safe_search("acc", "is_multi") Then
        pay_string = "_singlepay"
    Else
        pay_string = "_multipay"
    End If
    
    Call result_paths.Add("example_letter" & pay_string)
    
    Set fill_letter_paths = result_paths
End Function


Private Function get_module(query As String) As ActivityModule
    If p_modules.Exists(query) Then
        Set get_module = p_modules(query)
    Else
        Call throw_exception("Input is missing '" & query & "' module, run it back.", ActivityError.BadInputError)
    End If
End Function


Private Function safe_search(module_name As String, query As String) As Variant
    Dim module As ActivityModule
    Dim result As String
    Dim err_msg As String
    
    Set module = get_module(module_name)
    
    result = module.search(query)
    If result = vbNullString Then
        err_msg = "tried to safe search '" & query & "', couldnt find anything, but it should be there"
        Call err.Raise(ActivityError.DumbassError, Description:=module.str() & Chr(TEXT_NEWLINE) & err_msg)
    End If
    
    safe_search = result
End Function


Private Function safe_searchcoll(module_name As String, query As String) As Collection
    Dim module As ActivityModule
    Dim result As Collection
    Dim err_msg As String
    
    Set module = get_module(module_name)
    
    Set result = module.searchcoll(query)
    If result Is Nothing Then
        err_msg = "tried to safe search(coll) '" & query & "', couldnt find anything, but it should be there"
        Call err.Raise(ActivityError.DumbassError, Description:=module.str() & Chr(TEXT_NEWLINE) & err_msg)
    End If
    
    Set safe_searchcoll = result
End Function


Private Function str_append(existing As String, new_mod As ActivityModule) As String
    str_append = existing & new_mod.str() & Chr(TEXT_NEWLINE)
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_instructions As Collection
Private p_index As Integer
Private p_line_index As Integer

Private Sub Class_Initialize()
    Set p_instructions = New Collection
    
    'VBA collection indexes start from 1 instead of 0 because microsoft are loose units.
    p_line_index = 1
End Sub


Private Function is_instruction(input_line As String) As Boolean
    Dim result As Boolean
    
    'is blank line
    If Len(input_line) = 0 Then
        result = False
    
    'is a comment line
    ElseIf Left(input_line, 1) = "#" Then
        result = False
    
    'is a valid instruction
    Else
        result = True
    End If
    
    is_instruction = result
End Function


Private Function within_range(coll_index As Integer) As Boolean
    Dim within_start As Boolean
    Dim within_end As Boolean
    
    within_start = coll_index >= 1
    within_end = coll_index <= p_instructions.Count()
    
    within_range = within_start And within_end
End Function


Private Function ex_msg(index As Integer) As String
    Dim diff As Integer
    Dim rel_index As Integer
    Dim result As String
    
    If Not within_range(index) Then
        Call err.Raise(ActivityError.DumbassError, Description:="ex_msg(): index '" & index & "' is not within inst range")
    End If
    
    result = ""
    For diff = -2 To 2
        rel_index = index + diff
        
        If diff = 0 Then
            result = result & ">>>" & Chr(TEXT_TAB) & p_instructions(rel_index).str() & Chr(TEXT_TAB) & "  <<<" & Chr(TEXT_NEWLINE)
        ElseIf within_range(rel_index) Then
            result = result & Chr(TEXT_TAB) & p_instructions(rel_index).str() & Chr(TEXT_NEWLINE)
        Else
            result = result & Chr(TEXT_TAB) & "-" & Chr(TEXT_NEWLINE)
        End If
    Next diff
    
    ex_msg = result
End Function


Public Function throw_last_inst(msg As String) As String
    throw_last_inst = ex_msg(p_line_index - 1) & Chr(TEXT_NEWLINE) & msg & Chr(TEXT_NEWLINE)
End Function


Public Sub read_instfile(FILE_PATH As String)
    Dim file_no As Variant
    Dim line As String
    Dim new_instruction As Instruction
    
    file_no = FreeFile()

    Open FILE_PATH For Input As #file_no
        Do While Not EOF(file_no)
            line = ""
            Set new_instruction = New Instruction
            
            Line Input #file_no, line
            
            If is_instruction(line) Then
                Call new_instruction.initialise(line)
                Call p_instructions.Add(new_instruction)
            End If
        Loop
    Close #file_no
End Sub


Public Property Get instructions() As Collection
    Set instructions = p_instructions
End Property


Public Function get_next_inst(expected As String) As Instruction
    Call Me.assert_next_inst(expected)
    
    Set get_next_inst = p_instructions.item(p_line_index)
    p_line_index = p_line_index + 1
End Function


Public Function assert_next_inst(expected As String)
    Dim next_inst As String
    next_inst = Me.next_inst_type()
    
    ' checks if this function was used incorrectly
    If expected <> "arg" And expected <> "cmd" Then
        Call err.Raise(ActivityError.DumbassError, Description:="assert_next_inst(): argument wasnt 'cmd' or 'arg', you typed something wierd")
    
    ' ideal case
    ElseIf next_inst = expected Then
        Exit Function
    
    'unexpected instruction types
    ElseIf expected = "arg" Then
        Call throw_exception(ex_msg(p_line_index), ActivityError.ArgExpectedError)
    
    ElseIf expected = "cmd" Then
        Call throw_exception(ex_msg(p_line_index), ActivityError.CmdExpectedError)
    
    'f none of the above cases occoured, oh no
    Else
        Call err.Raise(ActivityError.DumbassError, Description:="assert_next_inst(): weird default case reached, dizastah")
    End If
End Function


Public Function next_inst_type() As String
    Dim next_inst As Instruction
    
    'if there are no more instructions
    If p_line_index > p_instructions.Count() Then
        next_inst_type = "eof"
        Exit Function
    End If
    
    Set next_inst = instructions.item(p_line_index)
    next_inst_type = next_inst.inst_type
End Function


Public Function str() As String
    Dim result As String
    Dim inst As Instruction
    
    result = ""
    For Each inst In p_instructions
        result = result & inst.str() & Chr(TEXT_NEWLINE)
    Next inst
    
    str = result
End Function


Public Function debug_append_inst(inst_type As String, inst_value As String)
    Dim new_inst As New Instruction
    
    Call new_inst.debug_initialise(inst_type, inst_value)
    Call p_instructions.Add(new_inst)
End Function



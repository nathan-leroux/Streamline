VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Instruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Instruction Object

Option Explicit

Private p_inst_type As String
Private p_inst_value As String


Public Sub initialise(input_string As String)
    Dim first_char As String
    Dim is_cmd As Boolean
    
    first_char = Mid$(input_string, 1, 1)
    is_cmd = StrComp(first_char, "/", vbBinaryCompare) = 0
    
    If is_cmd Then
        input_string = Right(input_string, Len(input_string) - 1)
        p_inst_type = "cmd"
    Else
        p_inst_type = "arg"
    End If

    'trim() doesnt remove whitespace other than spaces
    p_inst_value = Trim(Replace(input_string, Chr(TEXT_TAB), ""))
End Sub


Public Sub debug_initialise(inst_type As String, inst_value As String)
    p_inst_type = inst_type
    p_inst_value = inst_value
End Sub


Public Property Get inst_type() As String
    inst_type = p_inst_type
End Property


Public Property Let inst_type(inst_type As String)
    p_inst_type = inst_type
End Property


Public Property Get inst_value() As String
    inst_value = p_inst_value
End Property

Public Property Let inst_value(inst_value As String)
    p_inst_value = inst_value
End Property


Public Function str() As String
    Dim shortened_value As String
    
    If Len(p_inst_value) > 20 Then
        shortened_value = Left(p_inst_value, 20) & "..."
    Else
        shortened_value = p_inst_value
    End If
    
    str = "(" & p_inst_type & ", " & shortened_value & ")"
End Function


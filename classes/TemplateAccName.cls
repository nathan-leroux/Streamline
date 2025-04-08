VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateAccName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'TemplateAccName
'class for dealing with account name abreviations

Option Explicit

Private p_full As String
Private p_abrev As String
Private p_number As String
Private p_initialised As Boolean


Private Sub Class_Initialize()
    p_initialised = False
End Sub


Public Function populate(input_str As String) As Integer
    Dim err_status As Integer
    
    If InStr(input_str, " ") = 0 Then
        err_status = populate_abrev(input_str)
    Else
        err_status = populate_literal(input_str)
    End If
    
    p_initialised = True
    populate = err_status
End Function


Public Property Get abrev() As String
    Call check_init(p_abrev)
    
    abrev = p_abrev
End Property


Public Property Get short() As String
    Dim account_index As Integer

    account_index = InStr(1, p_full, "account", vbTextCompare)
    
    Call check_init(p_full)
    
    If account_index = 0 Then
        short = p_full
    Else
        short = Left(p_full, account_index - 2)
    End If
End Property


Public Property Get full() As String
    Call check_init(p_full)
    Call check_init(p_number)
    
    full = p_full & " " & p_number
End Property


Private Function populate_abrev(input_str As String) As Integer
    Dim short_array() As String
    Dim unusual_num As String
    
    short_array = Split(input_str, ".")
    
    If array_len(short_array) = 1 Then
        populate_abrev = parse_abrev(short_array(0))
    
    ElseIf array_len(short_array) = 2 Then
        unusual_num = short_array(1)
        populate_abrev = parse_abrev(short_array(0), unusual_num)
        
    Else
        populate_abrev = 1
        Exit Function
    End If
End Function


Private Function populate_literal(input_str As String) As Integer
    Dim last_space_index As Integer
    Dim account_index As Integer
    Dim number_selection As String
    Dim account_selection As String
    
    last_space_index = InStrRev(input_str, " ")
    account_index = InStr(1, input_str, "account", vbTextCompare)
    
    number_selection = Right(input_str, Len(input_str) - last_space_index)
    If IsNumeric(number_selection) Then
        p_number = number_selection
    Else
        p_number = ""
    End If
    
    p_full = Left(input_str, last_space_index - 1)
    
    If account_index <> 0 Then
        p_abrev = Left(input_str, account_index - 1)
    Else
        p_abrev = Left(input_str, last_space_index - 1)
    End If
    
    populate_literal = 0
End Function


Private Function parse_abrev(input_str As String, Optional unusual_num) As Integer
    Dim case_num As String
    
    Select Case input_str
        Case "exp"
            p_full = "Expenses Account"
            p_abrev = "EXP"
            
            case_num = "1"
        
        Case "rec"
            p_full = "Recievables Account"
            p_abrev = "REC"
            
            case_num = "2"
        
        Case Else
            parse_abrev = 1
            Exit Function
    End Select
    
    If IsMissing(unusual_num) Then
        p_number = case_num
    Else
        p_number = unusual_num
    End If
    
    parse_abrev = 0
End Function


Private Function check_init(input_str As String)
    If Not p_initialised Then
        Call err.Raise(ActivityError.DumbassError, Description:="AccName used without .populatate() being called")
    End If
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActivityAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ActivityAddress
' stores client adress and name

Option Explicit
Implements ActivityModule


Private p_name As String
Private p_address As Collection
Private p_is_corp As Boolean
Private p_lookup As Scripting.Dictionary
Private p_coll_lookup As Scripting.Dictionary


Private Sub Class_Initialize()
    Set p_address = New Collection
    Set p_lookup = New Scripting.Dictionary
    Set p_coll_lookup = New Scripting.Dictionary
End Sub


Private Function ActivityModule_populate(inst_file As InstFile)
    Dim corp_flag As String
    
    p_name = inst_file.get_next_inst("arg").inst_value
    
    p_is_corp = parse_corp_flag(inst_file)
    
    'gotta be at least one line of address
    Call p_address.Add(inst_file.get_next_inst("arg").inst_value)
    
    Do While inst_file.next_inst_type() = "arg"
        p_address.Add (inst_file.get_next_inst("arg").inst_value)
    Loop
    
    Call fill_lookup
End Function


Private Sub fill_lookup()
    Call p_lookup.Add("post_block", format_post_block())
    Call p_lookup.Add("greeting", format_greeting())
End Sub


Private Function ActivityModule_attach(act As Activity)
    Call act.add_module("adr", Me)
End Function


Private Function ActivityModule_search(query As String) As Variant
    ActivityModule_search = p_lookup(query)
End Function


Private Function ActivityModule_searchcoll(query As String) As Collection
    Set ActivityModule_searchcoll = p_coll_lookup(query)
End Function


Private Function ActivityModule_str() As String
    Dim result As String
    Dim adr_line As Variant
    
    result = "(address: name='" & p_name & "', is_corp=" & p_is_corp & "," & Chr(TEXT_NEWLINE)
    For Each adr_line In p_address
        result = result & Chr(TEXT_TAB) & adr_line & Chr(TEXT_NEWLINE)
    Next adr_line
    result = result + ")"
    
    ActivityModule_str = result
End Function


Private Function parse_corp_flag(inst_file As InstFile) As Boolean
    Dim new_str As String
    Dim err_msg As String
    
    new_str = inst_file.get_next_inst("arg").inst_value
    
    If StrComp(new_str, "p", vbBinaryCompare) = 0 Then
        parse_corp_flag = False
    
    ElseIf StrComp(new_str, "c", vbBinaryCompare) = 0 Then
        parse_corp_flag = True
    
    Else
        err_msg = "Invalid Entity Flag: '" & new_str & "'"
        Call throw_exception(inst_file.throw_last_inst(err_msg), ActivityError.BadInputError)
    End If
End Function


Private Function format_post_block() As String
    Dim result As String
    Dim line As Variant
    
    result = p_name & Chr(TEXT_NEWLINE)
    
    For Each line In insert_state(p_address)
        result = result & line & Chr(TEXT_NEWLINE)
    Next line
    
    format_post_block = result
End Function


Private Function format_greeting() As String
    Dim first As String
    
    If p_is_corp Then
        format_greeting = "To Whom It May Concern,"
        Exit Function
    End If
    
    'gets clients first name and capitalises
    first = StrConv(Split(p_name)(0), vbProperCase)
    
    format_greeting = "Dear " & first & ","
End Function


Private Function insert_state(input_adr As Collection) As Collection
    Dim result As Collection
    Dim last_line As String
    Dim postcode As Integer
    Dim state As String
    Dim remainder As String
    
    Set result = input_adr
    
    last_line = result(p_address.Count())
    postcode = CInt(Right(last_line, 4))
    
    Select Case postcode
        Case 200 To 299
            state = "ACT"
        Case 800 To 999
            state = "NT"
        Case 1000 To 2599
            state = "NSW"
        Case 2600 To 2618
            state = "ACT"
        Case 2619 To 2899
            state = "NSW"
        Case 2900 To 2920
            state = "ACT"
        Case 2921 To 2999
            state = "NSW"
        Case 3000 To 3999
            state = "VIC"
        Case 4000 To 4999
            state = "QLD"
        Case 5000 To 5999
            state = "SA"
        Case 6000 To 6999
            state = "WA"
        Case 7000 To 7999
            state = "TAS"
        Case 8000 To 8999
            state = "VIC"
        Case 9000 To 9999
            state = "QLD"
        Case Else
            Call throw_exception("Invalid Postcode: " & postcode & "'", ActivityError.BadInputError)
    End Select
    
    ' everything except the postcode
    remainder = Left(last_line, Len(last_line) - 4)
    
    'collections dont do reassignment, this is jank
    Call result.Remove(p_address.Count())
    Call result.Add(remainder & " " & state & "  " & postcode)
    
    Set insert_state = result
End Function

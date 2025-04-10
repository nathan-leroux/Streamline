VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActivityNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ActivityNote
'contains extra text needed for completing activities
'either extra response reasoning or general notes

Option Explicit
Implements ActivityModule


Private p_notes As Collection
Private p_lookup As Scripting.Dictionary
Private p_coll_lookup As Scripting.Dictionary


Private Sub Class_Initialize()
    Set p_notes = New Collection
    Set p_lookup = New Scripting.Dictionary
    Set p_coll_lookup = New Scripting.Dictionary
End Sub


Private Function ActivityModule_populate(inst_file As InstFile) As Variant
    Do While inst_file.next_inst_type = "arg"
        Call p_notes.Add(inst_file.get_next_inst("arg").inst_value)
    Loop
    
    Call fill_lookup
End Function


Private Sub fill_lookup()
    Call p_lookup.Add("content_formatted", format_note(p_notes))
    Call p_lookup.Add("content_plain", collection_format(p_notes, Chr(TEXT_NEWLINE), Chr(TEXT_NEWLINE)))
End Sub


Private Function ActivityModule_attach(act As Activity) As Variant
    Call act.add_module("not", Me)
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
    
    result = "(notes:" & Chr(TEXT_NEWLINE)
    
    For Each line In p_notes
        result = result & Chr(TEXT_TAB) & line & Chr(TEXT_NEWLINE)
    Next line
    
    result = result & ")"
    
    ActivityModule_str = result
End Function


Private Function format_note(input_coll As Collection) As String
    Dim space_index As Integer
    Dim current_line As String
    Dim result As String
    Dim note_line As Variant
    
    result = ""
    
    For Each note_line In input_coll
        current_line = note_line
        
        space_index = InStr(current_line, " ")
        current_line = StrConv(Left(current_line, space_index), vbProperCase) & Right(current_line, Len(current_line) - space_index)

        If Not Right(current_line, 1) = "." Then
            current_line = current_line & "."
        End If
    
        result = result & "- " & current_line & Chr(TEXT_NEWLINE)
    Next note_line
    
    format_note = result
End Function

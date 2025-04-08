VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' class that represents a word document

Option Explicit

Private Const SEARCH_VAR As String = "\@*\@"
Private Const SEARCH_COND As String = "\{*\}"
Private Const SEARCH_TABLE As String = "\[*\]"

Private p_document As Document
Private p_activity As Activity
Private p_missing_vars As Collection


Private Sub Class_Initialize()
    Set p_missing_vars = New Collection
End Sub


Public Function open_doc(template_path As String)
    Dim fso As New Scripting.FileSystemObject
    Dim abs_path As String
    Dim err_msg As String
    
    abs_path = fso.BuildPath(ROOT_DIR, template_path)
    If fso.FileExists(abs_path) Then
        Set p_document = Documents.Open(FileName:=abs_path)
        
    Else
        err_msg = abs_path & " couldnt be found"
        Call err.Raise(ActivityError.BadDocumentError, Description:=err_msg)
        
    End If
End Function


Public Function use_current_doc()
    Set p_document = ActiveDocument
End Function


Public Function close_doc(Optional save As Boolean = True)
    Dim save_flag As WdSaveOptions
    
    If save Then
        save_flag = wdSaveChanges
    Else
        save_flag = wdDoNotSaveChanges
    End If
    
    Call p_document.Close(SaveChanges:=save_flag)
    Set p_document = Nothing
End Function


Public Function save_doc(Optional file_name As String = "")
    Dim fso As New Scripting.FileSystemObject
    Dim abs_path As String
    
    If file_name = vbNullString Then
        Call p_document.SaveAs2
        
    Else
        abs_path = fso.BuildPath(ROOT_DIR, file_name)
        Call p_document.SaveAs2(FileName:=abs_path)
    End If
End Function


Public Function populate_doc(act As Activity)
    Dim var As Variant
    Dim story As Object
    Dim err_msg As String
    
    Set p_activity = act
    
    Call fill_conditionals(wdMainTextStory)
    Call fill_tables(wdMainTextStory)
    Call fill_variables(wdMainTextStory)
    
    If story_exists(wdFirstPageHeaderStory) Then
        Call fill_conditionals(wdFirstPageHeaderStory)
        Call fill_variables(wdFirstPageHeaderStory)
    End If
    
    If story_exists(wdPrimaryHeaderStory) Then
        Call fill_conditionals(wdPrimaryHeaderStory)
        Call fill_variables(wdPrimaryHeaderStory)
    End If
    
    If p_missing_vars.Count > 0 Then
        err_msg = "Missing the following variables:" & Chr(TEXT_NEWLINE)
        err_msg = err_msg & collection_format(p_missing_vars, Chr(TEXT_NEWLINE), Chr(TEXT_NEWLINE))
        
        Call throw_exception(err_msg, ActivityError.BadDocumentError)
    End If
End Function


Public Function copy_note()
    Call p_document.StoryRanges(wdMainTextStory).Copy
End Function


Public Function copy_desc()
    Call p_document.StoryRanges(wdPrimaryHeaderStory).Copy
End Function


Public Function export_to_pdf(file_name As String)
    Dim fso As New Scripting.FileSystemObject
    Dim abs_path As String
    
    abs_path = fso.BuildPath(ROOT_DIR, file_name)
    
    If p_document.Saved Then
        Call p_document.ExportAsFixedFormat(abs_path, wdExportFormatPDF, openafterexport:=True)
    Else
        Call err.Raise(ActivityError.BadDocumentError, Description:=p_document.name & " has unsaved changes")
    End If
End Function


Private Function fill_variables(section As WdStoryType)
    Dim marker As range
    Dim replacement_text As String
    Dim variable As String
    
    Do While True
        Set marker = find_marker(p_document.StoryRanges(section), SEARCH_VAR)
    
        If marker Is Nothing Then
            'couldnt find anymore
            Exit Function
        End If
        
        variable = remove_delims(marker)
        replacement_text = p_activity.search(variable)
        
        If replacement_text = vbNullString Then
            Call collection_sorted_add(p_missing_vars, variable)
        End If
        
        marker.Text = replacement_text
    Loop
End Function


Private Function fill_conditionals(section As WdStoryType)
    Dim marker As range
    Dim conditional_array() As String
    Dim bool_str As String
    Dim true_str As String
    Dim false_str As String
    Dim replacement_text As String
    Dim bool_variable As Variant
    
    
    Do While True
        Set marker = find_marker(p_document.StoryRanges(section), SEARCH_COND)
        
        If marker Is Nothing Then
            Exit Function
        End If
        
        conditional_array = split_conditional(marker)
        bool_str = conditional_array(0)
        true_str = conditional_array(1)
        false_str = conditional_array(2)
        
        
        bool_variable = p_activity.search(bool_str)
        
        If bool_variable = vbNullString Then
            Call collection_sorted_add(p_missing_vars, bool_str)
            marker.Text = ""
            
        ElseIf bool_variable Then
            marker.Text = true_str
        Else
            marker.Text = false_str
        End If
    Loop
End Function


Private Function fill_tables(section As WdStoryType)
    Dim marker As range
    Dim target_table As Table
    Dim target_row As Row
    Dim lookup_colls As Scripting.Dictionary
    Dim returned_coll As Collection
    Dim key As Variant
    
    Do While True
        Set marker = find_marker(p_document.StoryRanges(section), SEARCH_TABLE)
        
        If marker Is Nothing Then
            Exit Function
        End If
        
        Set target_table = find_table(marker, p_document.Tables)
        Set target_row = find_row(marker, target_table.Rows)
    
        Set lookup_colls = New Scripting.Dictionary
        
        For Each key In find_iterable_args(target_row.range).Keys
            Set returned_coll = p_activity.searchcoll(CStr(key))
            
            If returned_coll.Count = 0 Then
                Call collection_sorted_add(p_missing_vars, CStr(key))
            End If
            
            Call lookup_colls.Add(key, returned_coll)
        Next key
    
        If p_missing_vars.Count = 0 Then
            Call fill_rows(target_table, target_row, lookup_colls)
        Else
            Exit Function
        End If
    Loop
End Function


Private Function find_marker(start_range As range, pattern As String) As range
    'find doesnt search within the range
    'it searches the entire doc and reasigns the range
    'cant change it now because it could mess up this whole class
    Dim search_success As Boolean
    Dim marker As range
    
    Set marker = start_range
    
    With marker.Find
        .MatchWildcards = True
        .Text = pattern
        search_success = .Execute()
    End With
    
    If search_success Then
        Set find_marker = marker
    Else
        Set find_marker = Nothing
    End If
End Function


Private Function check_marker_bounds(marker As range, start_bound As Long, end_bound As Long) As range
    'bit of spagette to make find_marker behave how its supposed to
    If marker Is Nothing Then
        Set check_marker_bounds = Nothing
        
    ElseIf marker.Start < start_bound Or marker.End > end_bound Then
        Set check_marker_bounds = Nothing
        
    Else
        Set check_marker_bounds = marker
    End If
End Function


Private Function split_conditional(input_range As range) As String()
    Dim input_string As String
    Dim result_array(3) As String
    Dim split_array() As String
    Dim split_array_len As Integer
    
    input_string = remove_delims(input_range)
    
    split_array = Split(input_string, "|")
    split_array_len = UBound(split_array) + 1
    
    ' conditional might have 2 or 3
    result_array(0) = split_array(0)
    result_array(1) = split_array(1)
    
    If split_array_len = 2 Then
        result_array(2) = ""
    ElseIf split_array_len = 3 Then
        result_array(2) = split_array(2)
    Else
        Call err.Raise(ActivityError.BadDocumentError)
    End If
    
    split_conditional = result_array
End Function


Private Function remove_delims(input_range As range)
    Dim result As String
    
    result = input_range.Text
    
    remove_delims = Mid(result, 2, Len(result) - 2)
End Function


Private Function find_table(input_range As range, table_coll As Tables) As Table
    Dim lil_table As Object
    Dim within_table As Boolean
    Dim result As Table
    Dim range_index As Long
    
    range_index = input_range.Start
    
    For Each lil_table In table_coll
        With lil_table.range
            within_table = .Start <= range_index And .End > range_index
        End With
        
        If within_table Then
            Set result = lil_table
        End If
    Next lil_table
    
    If result Is Nothing Then
        If table_coll.NestingLevel = 1 Then
            ' the variable is at root level and this function shouldnt be called
            Call err.Raise(ActivityError.DumbassError)
        End If
        Set result = table_coll.Parent
        
    ElseIf result.Tables.Count > 0 Then
        Set result = find_table(input_range, result.Tables)
    End If
    
    Set find_table = result
End Function


Private Function find_row(marker As range, rows_coll As Rows) As Row
    Dim range_index As Long
    Dim result As Row
    Dim lil_row As Object
    Dim within_row As Boolean
    
    range_index = marker.Start
    
    For Each lil_row In rows_coll
        With lil_row.range
            within_row = .Start <= range_index And .End > range_index
        End With
        
        If within_row Then
            Set result = lil_row
        End If
    Next lil_row
    
    If result Is Nothing Then
        ' if marker is within table, marker must be in these rows
        ' so result cant possibly be nothing
        Call err.Raise(ActivityError.DumbassError)
    End If
    
    Set find_row = result
End Function


Private Function find_iterable_args(input_range As range) As Scripting.Dictionary
    Dim result As New Scripting.Dictionary
    Dim found_key As String
    Dim marker As range
    Dim start_bound As Long
    Dim end_bound As Long
    
    start_bound = input_range.Start
    end_bound = input_range.End
    
    Call input_range.SetRange(start_bound, start_bound)
    
    Set marker = check_marker_bounds(find_marker(input_range, SEARCH_TABLE), start_bound, end_bound)
    
    Do While Not marker Is Nothing
        found_key = remove_delims(marker)
        If Not result.Exists(found_key) Then
            Call result.Add(found_key, New Collection)
        End If
        
        'so the same field isnt found again
        start_bound = marker.End
        Call input_range.SetRange(start_bound, start_bound)
        
        Set marker = check_marker_bounds(find_marker(input_range, SEARCH_TABLE), start_bound, end_bound)
    Loop
    
    Set find_iterable_args = result
End Function


Private Function find_iterable_len(input_colls As Scripting.Dictionary) As Integer
    Dim length As Integer
    Dim coll As Variant
    
    length = input_colls.Items(0).Count()
    
    For Each coll In input_colls.Items
        If coll.Count() <> length Then
            ' if this ever goes off then somehow the collection sizes differ
            ' and something is very broken
            Call err.Raise(ActivityError.DumbassError)
        End If
    Next coll
    
    find_iterable_len = length
End Function


Private Function fill_rows(input_table As Table, input_row As Row, input_colls As Scripting.Dictionary)
    Dim target_row As Row
    Dim marker As range
    Dim replacement_text As String
    Dim start_row_index As Integer
    Dim iterable_len As Integer
    Dim i As Integer
    
    start_row_index = input_row.index
    iterable_len = find_iterable_len(input_colls)
    Debug.Print iterable_len
    
    Call input_row.range.Copy
    ' one less as we use the original line
    For i = 1 To iterable_len - 1
        Call input_row.range.Paste
    Next i
    
    Set target_row = input_table.Rows(start_row_index)
    For i = 1 To iterable_len
        Set marker = find_marker(target_row.range, SEARCH_TABLE)
        
        Do While Not marker Is Nothing
            replacement_text = input_colls(remove_delims(marker))(i)
            marker.Text = replacement_text
            
            Set marker = find_marker(target_row.range, SEARCH_TABLE)
        Loop
        Set target_row = target_row.Next
    Next i
End Function


Private Function story_exists(input_story As WdStoryType) As Boolean
    Dim story As Object
    
    For Each story In p_document.StoryRanges
        If story.StoryType = input_story Then
            story_exists = True
            Exit Function
        End If
    Next story
    
    story_exists = False
End Function


Private Function print_range(input_range As range)
    Debug.Print "(" & input_range.Start & ", " & input_range.End & ")"
End Function

Attribute VB_Name = "TemplateUtil"
'wide scope functions that are useful

Option Explicit


Public Function array_len(ByRef input_array As Variant) As Integer
    array_len = UBound(input_array) - LBound(input_array) + 1
End Function


Public Function collection_print(coll As Collection)
    Dim result As String
    Dim item As Variant
    
    result = "("
    For Each item In coll
        result = result & item & ", "
    Next item
    
    If coll.Count > 0 Then
        result = Left(result, Len(result) - 2) & ")"
    Else
        result = result & ")"
    End If
    
    Debug.Print result
End Function


'adds stuff to collection alphabetically, no dups
Public Function collection_sorted_add(coll As Collection, coll_input As Variant)
    Dim item As Variant
    Dim index As Integer
    Dim comparison As Integer
    
    index = 1
    For Each item In coll
        comparison = StrComp(coll_input, item, vbTextCompare)
        
        'the same
        If comparison = 0 Then
            Exit Function
            
        'coll_input sorts before
        ElseIf comparison = -1 Then
            Call coll.Add(coll_input, Before:=index)
            Exit Function
        End If
        
        index = index + 1
    Next item
    
    'if no spot found, add to end
    Call coll.Add(coll_input)
End Function


Public Function collection_columnise(input_col As Collection, index As Integer) As Collection
    Dim inner_col As Object
    Dim result As New Collection
    
    For Each inner_col In input_col
        Call result.Add(inner_col(index))
    Next inner_col
    
    Set collection_columnise = result
End Function


Public Function collection_exists(input_col As Collection, expression As Variant) As Boolean
    Dim item As Variant
    
    For Each item In input_col
        If item = expression Then
            collection_exists = True
            Exit Function
        End If
    Next item
    
    collection_exists = False
End Function


Public Function collection_format(input_col As Collection, seperator As String, last_sep As String) As String
    Dim index As Integer
    Dim result As String
    
    If input_col.Count() = 0 Then
        collection_format = ""
        Exit Function
        
    ElseIf input_col.Count() = 1 Then
        collection_format = input_col(1)
        Exit Function
    End If
    
    result = input_col(1)
    index = 2
    
    Do While index < input_col.Count()
        result = result & seperator & input_col(index)
        index = index + 1
    Loop
    
    collection_format = result & last_sep & input_col(input_col.Count())
End Function


Public Function collection_filter(input_col As Collection) As Collection
    Dim item As Variant
    Dim result_col As New Collection
    
    For Each item In input_col
        Call collection_sorted_add(result_col, item)
    Next item
    
    Set collection_filter = result_col
End Function


Public Function collection_sum(input_col As Collection) As Currency
    Dim total As Currency
    Dim value As Variant
    
    total = 0
    
    For Each value In input_col
        total = total + value
    Next value
    
    collection_sum = total
End Function


Public Function dictionary_print(dict As Scripting.Dictionary)
    Dim key As Variant
    Dim result As String
    
    result = "(" & Chr(TEXT_NEWLINE)
    
    For Each key In dict.Keys()
        result = result & key & " : " & dict(key) & Chr(TEXT_NEWLINE)
    Next key
    
    result = result & ")"
    
    Debug.Print result
End Function


Public Function throw_exception(msg As String, errno As Integer)
    Dim ex_title As String
    
    Select Case errno
        Case ActivityError.ArgExpectedError
            ex_title = "Argument Expected"
            
        Case ActivityError.CmdExpectedError
            ex_title = "Command Expected"
            
        Case ActivityError.BadDocumentError
            ex_title = "Bad Document"
            
        Case ActivityError.BadInputError
            ex_title = "Bad Input File"
            
        Case Else
            Call err.Raise(errno, Description:="throw_exception(): tried to throw an unexpected errno.")
    End Select
    
    Call MsgBox(msg, vbCritical + vbOKOnly + vbDefaultButton1, ex_title)
    Call err.Raise(errno)
End Function


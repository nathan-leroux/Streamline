Attribute VB_Name = "TestCode"
Option Explicit

Type MyType
    arg1 As String
    arg2 As Integer
End Type


Sub test_class_prop()
    Dim test_class As New TestClass
    
    test_class.name = "bob"
    
    Debug.Print test_class.name & " is a goose"
End Sub


Sub test_class_func()
    Dim test_class As New TestClass
    test_class.name = "jef"
    Debug.Print test_class.hello("bob")
End Sub


Sub test_error()
    On Error GoTo ErrorHandler
    
    On Error GoTo SecondHandler
    Call Err.Raise(ActivityError.DumbassError, Source:="Dumbass Error", Description:="fat chance this works")
    
    On Error GoTo ErrorHandler
    
    Exit Sub
SecondHandler:
    Debug.Print "huzzah"
    Err.Clear
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error no: " & Err.Number
    Err.Clear
End Sub


Public Function class_recast(module As TestClass)
    module.hawyee
End Function


Sub test_interface()
    Dim test_class As New TestClass
    Dim test_inter As TestInterface
    Dim recast As TestClass
    
    Set test_inter = test_class
    test_inter.yeehaw
    
    class_recast test_inter
    
End Sub


Sub test_dictionary()
    Dim dict As New Scripting.Dictionary
    Dim dict_key As Variant
    
    Call dict.Add("one", 1)
    Call dict.Add("two", 2)
    Call dict.Add("three", 3)

    
    For Each dict_key In dict.Keys()
        Debug.Print Left(dict_key, 2)
    Next dict_key
End Sub


Sub test_collection()
    Dim test As New Collection
    Dim item As Variant
    
    Call test.Add("one")
    Call test.Add("two")
    Call test.Add("three")
    
    Call test.Remove(3)
    Call test.Add("four")
    
    For Each item In test
        Debug.Print item
    Next item
End Sub



Sub test_type()
    Dim test As MyType
    test.arg1 = "one"
    test.arg2 = 1
    
    Debug.Print test.arg1
End Sub


Sub test_array()
    Dim test(1) As String
    
    test(0) = "one"
    test(1) = "two"
    
    Debug.Print array_len(test)
End Sub


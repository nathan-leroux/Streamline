Attribute VB_Name = "TemplateTest"
Option Explicit

Private Const FILE_PATH As String = "C:\Users\UDIB5\Desktop\local_files\Macros\mock_current.txt"

Sub test_instruction()
    Dim instr1 As New Instruction
    Dim instr2 As New Instruction
    
    instr1.initialise ("/cli")
    instr2.initialise ("123456789")
    
    Debug.Print instr1.inst_type
    Debug.Print instr1.inst_value
    Debug.Print instr1.str()
    
    Debug.Print instr2.inst_type
    Debug.Print instr2.inst_value
    Debug.Print instr2.str()

End Sub


Sub test_inst_file()
    'On Error GoTo TestHandler
    
    Dim inst_file As New InstFile
    
    Call inst_file.read_instfile(FILE_PATH)
    
    Call inst_file.get_next_inst("cmd")
    Call inst_file.get_next_inst("arg")
    
    Debug.Print inst_file.str()
    
    Exit Sub
    
TestHandler:
    Debug.Print "raised number: " & err.Number
    err.Clear
    Exit Sub
End Sub


Sub test_activity_module()
    Dim module As ActivityModule
    Dim cli As New ActivityClientIdentifier
    
    Dim inst_file As New InstFile
    
    Set module = cli
    
    Call inst_file.read_instfile(FILE_PATH)
    Call inst_file.get_next_inst
    
    Call module.populate(inst_file)
    
    Debug.Print module.str()
End Sub


Sub test_get_module()
    Dim inst_file As New InstFile
    Dim module As ActivityModule
    
    Call inst_file.debug_append_inst("arg", "123456789")
    
    Set module = get_module("cli")
    
    Call module.populate(inst_file)
    
    Debug.Print module.str()
End Sub


Sub test_multi_module()
    Dim inst_file As New InstFile
    Dim mod1 As ActivityModule
    Dim mod2 As ActivityModule
    
    Set mod1 = New ActivityAccount
    Set mod2 = New ActivityAccount
    
    Call inst_file.debug_append_inst("arg", "ica")
    Call inst_file.debug_append_inst("arg", "10.00")
    Call inst_file.debug_append_inst("arg", "1111")
    Call inst_file.debug_append_inst("arg", "it")
    Call inst_file.debug_append_inst("arg", "20.00")
    Call inst_file.debug_append_inst("arg", "2222")
    
    Call mod1.populate(inst_file)
    Call mod2.populate(inst_file)
    
    Debug.Print mod1.str()
    
    Call mod1.append(mod2)
    'Debug.Print mod1.str()
End Sub


Sub test_activity_full()
    Dim act As New Activity
    Dim module As ActivityModule
    Dim inst_file As New InstFile
    Dim current_inst As Instruction
    Dim ex_msg As String
    
    Call inst_file.read_instfile(FILE_PATH_TEST)
    Debug.Print inst_file.str()

    Call act.populate(inst_file)
    Debug.Print act.str()
End Sub


Sub test_sorted_add()
    Dim test_coll As New Collection
    
    Call test_coll.Add("b")
    Call test_coll.Add("d")
    
    Call collection_print(test_coll)
    
    Call collection_sorted_add(test_coll, "c")
    
    Call collection_print(test_coll)
End Sub


Sub test_throw_ex()
    Call throw_exception("this is an example msg", activi)
End Sub

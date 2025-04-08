Attribute VB_Name = "TemplateParserTest"
'TemplateParserTest
'code for making sure the parser object works correctly

Option Explicit

Sub test_currency()
    Dim inst_file As New InstFile
    Dim parser As New TemplateParser
    
    Call inst_file.debug_append_inst("arg", "$123,456")
    
    Debug.Print parser.get_currency(inst_file)
End Sub


Sub test_lodgement_type()
    Dim inst_file As New InstFile
    Dim parser As New TemplateParser
    
    Call inst_file.debug_append_inst("arg", "as")
    
    Debug.Print parser.get_lodgement(inst_file)
End Sub


Sub test_remission_type()
    Dim inst_file As New InstFile
    Dim parser As New TemplateParser
    
    Call inst_file.debug_append_inst("arg", "gic")
    
    Debug.Print parser.get_remtype(inst_file)
End Sub


Sub test_accname()
    Dim inst_file As New InstFile
    Dim parser As New TemplateParser
    
    Call inst_file.debug_append_inst("arg", "it")
    
    Debug.Print parser.get_accname(inst_file).full
End Sub


Sub test_date()
    Dim inst_file As New InstFile
    Dim parser As New TemplateParser
    
    Call inst_file.debug_append_inst("arg", "21-05-2021")
    
    Debug.Print parser.get_date(inst_file).full
End Sub


Sub test_daterange()
    Dim inst_file As New InstFile
    Dim parser As New TemplateParser
    
    Call inst_file.debug_append_inst("arg", "mar.23")
    
    Debug.Print parser.get_daterange(inst_file).std_range
End Sub

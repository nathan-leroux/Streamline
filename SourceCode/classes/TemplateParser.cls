VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'TemplateParser
'collection of parsing functions to modularise input checks

Option Explicit


Public Function get_currency(inst_file As InstFile) As Currency
    Dim new_str As String
    Dim err_msg As String
    
    new_str = inst_file.get_next_inst("arg").inst_value
    
    new_str = Replace(new_str, "$", "")
    new_str = Replace(new_str, ",", "")
    
    If Not IsNumeric(new_str) Then
        err_msg = "Invalid Currency: '" & new_str & "'"
        Call throw_exception(inst_file.throw_last_inst(err_msg), ActivityError.BadInputError)
    End If
    
    get_currency = CCur(new_str)
End Function


Public Function get_lodgement(inst_file As InstFile) As String
    Dim new_str As String
    Dim err_msg As String
    
    new_str = inst_file.get_next_inst("arg").inst_value
    
    Select Case new_str
        Case "bs"
            get_lodgement = "Bank Statement"
        Case Else
            err_msg = "Invalid Lodgement Type: '" & new_str & "'"
            Call throw_exception(inst_file.throw_last_inst(err_msg), ActivityError.BadInputError)
    End Select
End Function


Public Function get_accname(inst_file As InstFile) As TemplateAccName
    Dim new_accname As New TemplateAccName
    Dim return_code As Integer
    Dim new_str As String
    Dim err_msg As String
    
    new_str = inst_file.get_next_inst("arg").inst_value
    
    return_code = new_accname.populate(new_str)
    
    If return_code <> 0 Then
        err_msg = "Invalid Account Name: '" & new_str & "'"
        Call throw_exception(inst_file.throw_last_inst(err_msg), ActivityError.BadInputError)
    End If
        
    Set get_accname = new_accname
End Function


Public Function get_date(inst_file As InstFile) As TemplateDate
    Dim new_date As New TemplateDate
    Dim return_code As Integer
    Dim new_str As String
    Dim err_msg As String
    
    new_str = inst_file.get_next_inst("arg").inst_value
    
    return_code = new_date.populate(new_str)
    
    If return_code <> 0 Then
        err_msg = "Invalid Date: '" & new_str & "'"
        Call throw_exception(inst_file.throw_last_inst(err_msg), ActivityError.BadInputError)
    End If
    
    Set get_date = new_date
End Function


Public Function get_daterange(inst_file As InstFile) As TemplateDateRange
    Dim new_daterange As New TemplateDateRange
    Dim return_code As Integer
    Dim new_str As String
    Dim err_msg As String
    
    new_str = inst_file.get_next_inst("arg").inst_value
    
    return_code = new_daterange.populate(new_str)
    
    If return_code <> 0 Then
        err_msg = "Invalid Date Range: '" & new_str & "'"
        Call throw_exception(inst_file.throw_last_inst(err_msg), ActivityError.BadInputError)
    End If
    
    Set get_daterange = new_daterange
End Function


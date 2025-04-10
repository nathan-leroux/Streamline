VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActivityInbound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ActivityInbound
' Class for client identification number
' Activity module

Option Explicit
Implements ActivityModule

Private p_cli As String
Private p_rec_date As TemplateDate
Private p_flags As Collection

Private p_lookup As Scripting.Dictionary
Private p_coll_lookup As Scripting.Dictionary


Private Sub Class_Initialize()
    Set p_lookup = New Scripting.Dictionary
    Set p_coll_lookup = New Scripting.Dictionary
End Sub


Public Function ActivityModule_populate(inst_file As InstFile)
    Dim parser As New TemplateParser
    
    p_cli = parse_cli(inst_file)
    
    Set p_rec_date = parser.get_date(inst_file)
    
    Do While inst_file.next_inst_type() = "arg"
        Call collection_sorted_add(p_flags, parse_flag(inst_file))
    Loop
    
    Call fill_lookup
End Function


Private Sub fill_lookup()
    Call p_lookup.Add("cli", format_cli(p_cli))
    Call p_lookup.Add("date_full", p_rec_date.full)
    Call p_lookup.Add("date_short", p_rec_date.short)
End Sub


Public Function ActivityModule_attach(act As Activity)
    Call act.add_module("in", Me)
End Function


Private Function ActivityModule_search(query As String) As Variant
    ActivityModule_search = p_lookup(query)
End Function


Private Function ActivityModule_searchcoll(query As String) As Collection
    Set ActivityModule_searchcoll = p_coll_lookup(query)
End Function


Public Function ActivityModule_str() As String
    Dim result As String
    
    result = "(Inbound: cli=" & p_cli & ", sender=" & p_sender & ", recieved=" & p_rec_date.short
    result = result & ", flags=" & p_flags.Count() & ")"
    
    ActivityModule_str = result
End Function


Private Function parse_cli(inst_file As InstFile) As String
    Dim length As Integer
    Dim new_str As String
    Dim err_msg As String
    
    new_str = inst_file.get_next_inst("arg").inst_value
    length = Len(new_str)
    
    If length = 8 Then
        parse_cli = new_str
    Else
        err_msg = "client no. should be 8 long, was " & length
        Call throw_exception(inst_file.throw_last_inst(err_msg), ActivityError.BadInputError)
    End If
End Function


Private Function parse_flag(inst_file As InstFile) As Long
    Dim new_str As String
    Dim err_msg As String
    
    new_str = inst_file.get_next_inst("arg").inst_value
    
    Select Case new_str
        Case "mail"
            parse_flag = InboundFlag.Mail
        Case "bp"
            parse_flag = InboundFlag.Bp
        Case "rich"
            parse_flag = InboundFlag.Rich
        Case Else
            err_msg = "Invalid Inbound Flag: '" & new_str & "'"
            Call throw_exception(inst_file.throw_last_inst(err_msg), ActivityError.BadInputError)
    End Select
End Function


Private Function format_cli(cli As String) As String
    Dim cli_remainder As String
    
    If Len(cli) <= 3 Then
        format_cli = cli
    
    Else
        'everything except the last 3
        cli_remainder = Left(cli, Len(cli) - 3)
        
        format_cli = format_cli(cli_remainder) & " " & Right(cli, 3)
    End If
End Function


Private Function format_sender(sender As String) As String
    If sender = "-" Then
        format_sender = "Tax Professional"
    Else
        format_sender = StrConv(p_sender, vbProperCase)
    End If
End Function

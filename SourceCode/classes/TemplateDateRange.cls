VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateDateRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'TemplateDateRange
'Stores start and end date and provides formating

Option Explicit

Private p_start_date As Date
Private p_end_date As Date
Private p_literal As String


Public Function populate(input_range As String) As Integer
    Dim is_shortcut As Boolean
    Dim return_code As Integer
    
    is_shortcut = InStr(input_range, " to ") = 0
    
    If is_shortcut Then
        return_code = populate_shortcut(input_range)
    Else
        return_code = populate_literal(input_range)
    End If
    
    populate = return_code
End Function


Public Property Get std_range() As String
    If p_literal = vbNullString Then
        Call err.Raise(ActivityError.DumbassError, Description:="DateRange used without .populatate() being called")
    End If
    std_range = p_literal
End Property


Private Function populate_shortcut(input_range As String) As Integer
    Dim short_array() As String
    Dim month As Integer
    Dim year As Integer
    
    short_array = Split(input_range, ".")
    
    If array_len(short_array) <> 2 Then
        populate_shortcut = 1
        Exit Function
    End If

    year = parse_year(short_array(1))
    If IsNull(year) Then
        populate_shortcut = 1
        Exit Function
    End If
    
    Select Case short_array(0)
        Case "jan"
            p_start_date = get_start_date(1, year)
            p_end_date = get_end_date(1, year)
        
        Case "feb"
            p_start_date = get_start_date(2, year)
            p_end_date = get_end_date(2, year)
            
        Case "mar"
            p_start_date = get_start_date(3, year)
            p_end_date = get_end_date(3, year)
            
        Case "apr"
            p_start_date = get_start_date(4, year)
            p_end_date = get_end_date(4, year)
            
        Case "may"
            p_start_date = get_start_date(5, year)
            p_end_date = get_end_date(5, year)
            
        Case "jun"
            p_start_date = get_start_date(6, year)
            p_end_date = get_end_date(6, year)
            
        Case "jul"
            p_start_date = get_start_date(7, year)
            p_end_date = get_end_date(7, year)
            
        Case "aug"
            p_start_date = get_start_date(8, year)
            p_end_date = get_end_date(8, year)
            
        Case "sep"
            p_start_date = get_start_date(9, year)
            p_end_date = get_end_date(9, year)
            
        Case "oct"
            p_start_date = get_start_date(10, year)
            p_end_date = get_end_date(10, year)
            
        Case "nov"
            p_start_date = get_start_date(11, year)
            p_end_date = get_end_date(11, year)
            
        Case "dec"
            p_start_date = get_start_date(12, year)
            p_end_date = get_end_date(12, year)
            
        Case "fy"
            p_start_date = get_start_date(7, year - 1)
            p_end_date = get_end_date(6, year)
            
        Case Else
            populate_shortcut = 1
            Exit Function
    End Select
    
    p_literal = to_datestring(p_start_date) & " -" & Chr(TEXT_NEWLINE) & to_datestring(p_end_date)
    populate_shortcut = 0
End Function


Private Function populate_literal(input_range As String) As Integer
    p_literal = Replace(input_range, " to ", " -" & Chr(TEXT_NEWLINE))
    populate_literal = 0
End Function


Private Function parse_year(keyword As String) As Integer
    If Not IsNumeric(keyword) And Len(keyword) <> 2 Then
        parse_year = Null
        Exit Function
    End If
    
    'sorry pre 21st century
    parse_year = CInt("20" & keyword)
End Function


Private Function format_day(input_day As String) As String
    If Len(input_day) = 1 Then
        input_day = "0" & input_day
    End If
    If Len(input_day) <> 2 Then
        Call err.Raise(ActivityError.DumbassError, Description:="format_day() called with invalid '" & input_day & "'")
    End If
    format_day = input_day
End Function


Private Function format_year(input_year As String) As String
    If Len(input_year) = 4 Then
        format_year = Right(input_year, 2)
    Else
        Call err.Raise(ActivityError.DumbassError, Description:="format_year() called with invalid '" & input_year & "'")
    End If
End Function


Private Function to_datestring(input_date As Date) As String
    to_datestring = format_day(day(input_date)) & " " & MonthName(month(input_date), True) & " " & format_year(year(input_date))
End Function


Private Function get_start_date(month As Integer, year As Integer) As Date
    get_start_date = DateSerial(year, month, 1)
End Function


Private Function get_end_date(month As Integer, year As Integer) As Date
    get_end_date = DateSerial(year, month + 1, 0)
End Function


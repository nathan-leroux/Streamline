VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'TemplateDate
'regular old date with a bit of formating options

Option Explicit

Private p_date As Date
Private p_initialised As Boolean


Private Sub Class_Initialize()
    p_initialised = False
End Sub


Public Function populate(str_date As String) As Integer
    Dim date_array() As String
    Dim date_part As Integer
    
    date_array = Split(str_date, "-")
    
    If array_len(date_array) <> 3 Then
        populate = 1
        Exit Function
    End If
    
    For date_part = 0 To 2
        If Not IsNumeric(date_array(date_part)) Then
            populate = 1
            Exit Function
        End If
    Next date_part
    
    p_date = DateSerial(CInt(date_array(2)), CInt(date_array(1)), CInt(date_array(0)))
    p_initialised = True
    populate = 0
End Function


Public Function populate_from_date(input_date As Date) As Integer
    p_date = input_date
    p_initialised = True
    
    populate_from_date = 0
End Function


Public Property Get short() As String
    Call check_init(p_date)
        
    short = day(p_date) & "/" & month(p_date) & "/" & year(p_date)
End Property


Public Property Get full() As String
    Call check_init(p_date)
    
    full = day(p_date) & day_exten(day(p_date)) & " " & MonthName(month(p_date)) & " " & year(p_date)
End Property


Private Function day_exten(input_day As Integer) As String
    Dim ones_digit As String
   
    If input_day > 10 And input_day < 20 Then
        day_exten = "th"
        Exit Function
    End If
    
    ones_digit = Right(CStr(input_day), 1)
    
    Select Case ones_digit
        Case "1"
            day_exten = "st"
        Case "2"
            day_exten = "nd"
        Case "3"
            day_exten = "rd"
        Case Else
            day_exten = "th"
    End Select
End Function

Private Function check_init(input_date As Date)
    If Not p_initialised Then
        Call err.Raise(ActivityError.DumbassError, Description:="Date used without .populatate() being called")
    End If
End Function



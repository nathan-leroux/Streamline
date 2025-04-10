Attribute VB_Name = "TemplateDatatypeTest"
' TemplateDatatypeTest
' for testing custom datatypes

Option Explicit

Sub test_date_range()
    Dim dr1 As New TemplateDateRange
    Dim dr2 As New TemplateDateRange
    Dim dr3 As New TemplateDateRange
    Dim dates_coll As New Collection
    Dim dr As Object
    Dim err As Integer
    
    err = err + dr1.populate("feb.23")
    err = err + dr2.populate("fy.23")
    err = err + dr3.populate("30 Jun 23 to 01 Jul 24")
    
    Call dates_coll.Add(dr1)
    Call dates_coll.Add(dr2)
    Call dates_coll.Add(dr3)
    
    For Each dr In dates_coll
        Debug.Print dr.std_range
    Next dr
End Sub


Sub test_date()
    Dim temp_date As New TemplateDate
    
    Call temp_date.populate("33-09-2020")
    
    Debug.Print temp_date.full
    Debug.Print temp_date.short
End Sub


Sub test_acc_name()
    Dim an1 As New TemplateAccName
    Dim an2 As New TemplateAccName
    Dim an3 As New TemplateAccName
    Dim accnames As New Collection
    Dim an As Object
    
    Call an1.populate("it")
    Call an2.populate("it.2")
    Call an3.populate("Test Account 2")
    
    Call accnames.Add(an1)
    Call accnames.Add(an2)
    Call accnames.Add(an3)
    
    For Each an In accnames
        Debug.Print an.abrev
        Debug.Print an.short
        Debug.Print an.full
    Next an
End Sub

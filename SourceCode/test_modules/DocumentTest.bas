Attribute VB_Name = "DocumentTest"
' test code for manipulating word docs

Option Explicit


Sub test_doc_export()
    Dim act As New Activity
    Dim test_doc As New TemplateDocument
    
    Call test_doc.use_current_doc
    Call test_doc.save_doc
    Call test_doc.export_to_pdf("result")
End Sub


Sub test_doc_populate()
    Dim test_doc As New TemplateDocument
    Dim act As New Activity
    
    Call test_doc.open_doc("Notes\balls.docx")
    Call test_doc.populate_doc(act)
End Sub

Sub test_doc_copy()
    Dim test_doc As New TemplateDocument
    Dim act As New Activity
    
    Call test_doc.use_current_doc
    Call test_doc.copy_desc
End Sub

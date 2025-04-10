Attribute VB_Name = "Template"
'Template
'program main entry points

Option Explicit


' Sub single_export()
'     Dim current_document As New TemplateDocument
'     Dim act As New Activity
'     Dim inst_file As New InstFile
'     Dim file_name As String
'
'     Call inst_file.read_instfile(INSTFILE_PATH)
'
'     Call act.populate(inst_file)
'
'     file_name = act.search("rpy.id")
'
'     Call current_document.use_current_doc
'     Call current_document.save_doc(file_name)
'     Call current_document.populate_doc(act)
'     Call current_document.save_doc
'     Call current_document.export_to_pdf(file_name)
'     'Call current_document.close_doc
' End Sub
'
'
' Sub single_copy()
'     Dim current_document As New TemplateDocument
'     Dim act As New Activity
'     Dim inst_file As New InstFile
'
'     Call inst_file.read_instfile(INSTFILE_PATH)
'
'     Call act.populate(inst_file)
'
'     Call current_document.use_current_doc
'     Call current_document.populate_doc(act)
'     Call current_document.copy_desc
'     Call MsgBox("Description copied.", vbApplicationModal)
'     Call current_document.copy_note
'     Call MsgBox("Note copied.", vbApplicationModal)
'     Call current_document.close_doc(save:=False)
' End Sub

Public SCRIPT_ROOT_DIR As String
Public SCRIPT_INSTFILE_PATH As String
Public SCRIPT_LETTER_DIR As String
Public SCRIPT_NOTE_DIR As String
Public SCRIPT_OUTPUT_DIR As String



Private Function get_path()
    SCRIPT_ROOT_DIR = ActiveDocument.AttachedTemplate.Path
    
    SCRIPT_INSTFILE_PATH = SCRIPT_ROOT_DIR & "\instruction_file.txt"
    SCRIPT_LETTER_DIR = SCRIPT_ROOT_DIR & "\Letters\"
    SCRIPT_NOTE_DIR = SCRIPT_ROOT_DIR & "\Notes\"
    SCRIPT_OUTPUT_DIR = SCRIPT_ROOT_DIR & "\Output\"
End Function

Sub Complete_Notes()
    Dim act As New Activity
    Dim inst_file As New InstFile
    
    Dim note_doc As New TemplateDocument
    Dim reply_doc As New TemplateDocument
    
    Call get_path
    
    Call inst_file.read_instfile(SCRIPT_INSTFILE_PATH)
    
    Call act.populate(inst_file)
    
    If act.note_path <> vbNullString Then
        Call note_doc.open_doc(SCRIPT_NOTE_DIR & act.note_path & ".docx")
        Call note_doc.populate_doc(act)
        
        Call execute_note(note_doc)
    End If
End Sub


Sub Complete_Letters()
    Dim act As New Activity
    Dim inst_file As New InstFile
    
    Dim letter_path As Variant
    Dim letter_name As String
    Dim letter_doc As TemplateDocument
    
    Call get_path
    
    Call inst_file.read_instfile(SCRIPT_INSTFILE_PATH)
    
    Call act.populate(inst_file)
    
    For Each letter_path In act.letter_paths
        Set letter_doc = New TemplateDocument
        
        letter_name = act.search("rpy.id")
        
        Call letter_doc.open_doc(SCRIPT_LETTER_DIR & CStr(letter_path) & ".docx")

        Call letter_doc.save_doc(SCRIPT_OUTPUT_DIR & letter_name)
        Call letter_doc.populate_doc(act)
    
        Call letter_doc.save_doc
        Call letter_doc.export_to_pdf(SCRIPT_OUTPUT_DIR & letter_name & ".pdf")
    Next letter_path
End Sub


Private Function execute_note(current_doc As TemplateDocument)
    Call current_doc.copy_desc
    Call MsgBox("Description copied.", vbApplicationModal)
    
    Call current_doc.copy_note
    Call MsgBox("Note copied.", vbApplicationModal)
    
    Call current_doc.close_doc(save:=False)
End Function


Private Function execute_reply(current_doc As TemplateDocument)
    Call current_doc.copy_note
    Call MsgBox("Reply Copied.", vbApplicationModal)
    
    Call current_doc.close_doc(save:=False)
End Function


Private Function execute_letter(current_doc As TemplateDocument, file_name As String)
    Call current_doc.save_doc(file_name)
    Call current_doc.populate_doc(act)
    
    Call current_doc.save_doc
    Call current_doc.export_to_pdf(file_name)
    
    Call current_doc.close_doc
End Function

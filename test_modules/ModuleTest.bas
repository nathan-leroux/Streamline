Attribute VB_Name = "ModuleTest"
' test file for activity modules

Option Explicit


Sub test_in()
    Dim module As ActivityModule
    Dim cli As New ActivityInbound
    Dim inst_file As New InstFile
    
    Set module = cli
    
    Call inst_file.debug_append_inst("arg", "123456789")
    Call inst_file.debug_append_inst("arg", "thibs")
    Call inst_file.debug_append_inst("arg", "01-01-2024")
    Call inst_file.debug_append_inst("arg", "mail")
    Call inst_file.debug_append_inst("arg", "rich")
    
    Call module.populate(inst_file)
    
    Debug.Print module.str()
End Sub


Sub test_adr()
    Dim module As ActivityModule
    Dim adr As New ActivityAddress
    Dim inst_file As New InstFile
    
    Set module = adr
    
    Call inst_file.debug_append_inst("arg", "JALEN BRUNSON")
    Call inst_file.debug_append_inst("arg", "p")
    Call inst_file.debug_append_inst("arg", "Court 1")
    Call inst_file.debug_append_inst("arg", "Maddison Square Garden")
    Call inst_file.debug_append_inst("arg", "Manhattan 6000")
    
    Call module.populate(inst_file)
    
    Debug.Print module.str()
End Sub


Sub test_rem()
    Dim module As ActivityModule
    Dim remish As New ActivityRemissions
    Dim inst_file As New InstFile
    
    Set module = remish
    
    Call inst_file.debug_append_inst("arg", "gic")
    Call inst_file.debug_append_inst("arg", "ica")
    Call inst_file.debug_append_inst("arg", "500.00")
    Call inst_file.debug_append_inst("arg", "gic")
    Call inst_file.debug_append_inst("arg", "it")
    Call inst_file.debug_append_inst("arg", "313.00")

    
    Call module.populate(inst_file)
    
    Debug.Print module.str()
End Sub


Sub test_acc()
    Dim module As ActivityModule
    Dim acc As New ActivityAccounts
    Dim inst_file As New InstFile
    
    Set module = acc
    
    Call inst_file.debug_append_inst("arg", "ica")
    Call inst_file.debug_append_inst("arg", "1000.00")
    Call inst_file.debug_append_inst("arg", "prn1")
    Call inst_file.debug_append_inst("arg", "it")
    Call inst_file.debug_append_inst("arg", "2000.00")
    Call inst_file.debug_append_inst("arg", "prn2")
    
    Call module.populate(inst_file)
    
    Debug.Print module.str()
End Sub


Sub test_ldg()
    Dim module As ActivityModule
    Dim ldg As New ActivityLodgements
    Dim inst_file As New InstFile
    
    Set module = ldg
    
    Call inst_file.debug_append_inst("arg", "as")
    Call inst_file.debug_append_inst("arg", "01 Apr 24 to 30 Apr 24")
    Call inst_file.debug_append_inst("arg", "626.00")
    Call inst_file.debug_append_inst("arg", "tr")
    Call inst_file.debug_append_inst("arg", "01 Sep 24 to 30 Sep 24")
    Call inst_file.debug_append_inst("arg", "313.00")
    
    Call module.populate(inst_file)
    
    Debug.Print module.str()
End Sub


Sub test_ovr()
    Dim module As ActivityModule
    Dim ovr As New ActivityOverdues
    Dim inst_file As New InstFile
    
    Set module = ovr
    
    Call inst_file.debug_append_inst("arg", "as")
    Call inst_file.debug_append_inst("arg", "dec.24")
    Call inst_file.debug_append_inst("arg", "tr")
    Call inst_file.debug_append_inst("arg", "sep.24")
    
    Call module.populate(inst_file)
    
    Debug.Print module.str()
End Sub


Sub test_out()
    Dim module As ActivityModule
    Dim out As New ActivityOutbound
    Dim inst_file As New InstFile
    
    Set module = out
    
    Call inst_file.debug_append_inst("arg", "low")
    Call inst_file.debug_append_inst("arg", "business")
    Call inst_file.debug_append_inst("arg", "covid")
    Call inst_file.debug_append_inst("arg", "exten")
    
    Call module.populate(inst_file)
    
    Debug.Print module.str()
End Sub


Sub test_res()
    Dim module As ActivityModule
    Dim res As New ActivityReason
    Dim inst_file As New InstFile
    
    Set module = res
    
    Call inst_file.debug_append_inst("arg", "my wife's boyfriend's cat died")
    Call inst_file.debug_append_inst("arg", "mercury was in retrograde")
    Call inst_file.debug_append_inst("arg", "muh cashflows")
    Call inst_file.debug_append_inst("arg", "boots &amp; cats &amp; rocknroll")
    
    Call module.populate(inst_file)
    
    Debug.Print module.str()
End Sub


Sub test_rpy()
    Dim module As ActivityModule
    Dim rpy As New ActivityReply
    Dim inst_file As New InstFile
    
    Set module = rpy
    
    Call inst_file.debug_append_inst("arg", "1-ABCDEF")
    Call inst_file.debug_append_inst("arg", "5510000000")
    
    Call module.populate(inst_file)
    
    Debug.Print module.str()
End Sub


Sub test_not()
    Dim module As ActivityModule
    Dim note As New ActivityNote
    Dim inst_file As New InstFile
    
    Set module = note
    
    Call inst_file.debug_append_inst("arg", "the client smells.")
    Call inst_file.debug_append_inst("arg", "the client is a gamer")
    
    Call module.populate(inst_file)
    
    Debug.Print "balls"
    Debug.Print module.str()
End Sub

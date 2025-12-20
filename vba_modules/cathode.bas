' Cathode material contribution (execute Module1 ~ Module6)

Sub cathode()

    ' Show UserForm asynchronously
    UserForm1.Show vbModeless
    ' Execute code while the message is displayed
    Call a ' Actual execution code
    ' Close UserForm after execution
    Unload UserForm1
    
    UserForm2.Show vbModeless
    Call b
    Unload UserForm2
    
    UserForm33.Show vbModeless
    Call cc
    Unload UserForm33
    
    UserForm44.Show vbModeless
    Call dd
    Unload UserForm44

    Call f
    
End Sub

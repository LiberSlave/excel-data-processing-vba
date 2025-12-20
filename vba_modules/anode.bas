' Anode material contribution (execute Module1 ~ Module6)

Sub anode()

    ' Show UserForm asynchronously
    UserForm1.Show vbModeless
    ' Execute code while the message is displayed
    Call a ' Actual execution code
    ' Close UserForm after execution
    Unload UserForm1
    
    UserForm2.Show vbModeless
    Call b
    Unload UserForm2
    
    UserForm3.Show vbModeless
    Call c
    Unload UserForm3
    
    UserForm4.Show vbModeless
    Call d
    Unload UserForm4
    
    Call f
    
End Sub

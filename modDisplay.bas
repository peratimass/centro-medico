Attribute VB_Name = "modDisplay"


Public Sub DisplayTextoDirect(dspText As String)
        
   On Error GoTo DisplayTexto_Error
    DoEvents
    
    Open "COM1:" For Output As #1
    Print #1, Chr$(12)
    Print #1, Chr$(24)
    Print #1, dspText
    Close #1
    
    DoEvents
   On Error GoTo 0
   Exit Sub

DisplayTexto_Error:
    
End Sub

Public Sub DisplayTextoCom(dspText As String, mscConector As MSComm)
   On Error GoTo DisplayTextoCom_Error
    DoEvents
    
    mscConector.PortOpen = True
    mscConector.Output = Chr$(12)
    mscConector.Output = Chr$(24)
    mscConector.Output = dspText
    mscConector.PortOpen = False

   On Error GoTo 0
   Exit Sub

DisplayTextoCom_Error:

End Sub



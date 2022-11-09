Attribute VB_Name = "ModVersion"
Public Sub put_version_update(ByVal in_ruc As String)
Dim strRuta_ini As String
MDIFrmPrincipal.Tiner_update.Enabled = True
strRuta_ini = App.Path & "\archivos\vitekey.ini"

FileName = Dir(strRuta_ini)
fnum = FreeFile
    Open strRuta_ini For Input As fnum
    i = 0
    Do While Not EOF(fnum)
        Select Case i
            Case 0  ' Servidor N01
                 Line Input #fnum, file_line
                  sys_Server = file_line
            Case 1 ' Servidor N02
                 Line Input #fnum, file_line
                  sys_Server2 = file_line
            Case 2 ' VERSION DE SOFT
                 Line Input #fnum, file_line
                 KEY_VERSION = file_line
            Case 3
                Line Input #fnum, file_line
                 KEY_RUC_VERSION = file_line
        End Select
        i = i + 1
    Loop
    
    
    Close #fnum
    
    If Len(in_ruc) > 0 Then
        Open strRuta_ini For Output As #1
        Print #1, sys_Server
        Print #1, sys_Server2
        Print #1, KEY_VERSION
        Print #1, in_ruc
        Close #1
    End If
    
    

End Sub

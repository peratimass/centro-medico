Attribute VB_Name = "ModExecute"
Public Sub Execute_Sql(ByVal Cadena As String)
On Error GoTo sin_conexion
reiniciar:
CnBd.Execute (Cadena)

Exit Sub
sin_conexion:
  
  Call reiniciar_conexion
  GoTo reiniciar
  Exit Sub


End Sub

Public Function get_error(ByVal in_error As Error) As String

    MsgBox "DISCULPE LAS MOLESTIAS A OCURRIDO UN FALLO DE RED." + Chr(13) + Err.Number & " " & Err.Description, vbInformation, MSGERROR
    
End Function


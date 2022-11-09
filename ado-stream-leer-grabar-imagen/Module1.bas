Attribute VB_Name = "Module1"
Public CnBd As ADODB.Connection
Public Rst As ADODB.Recordset
Public strCadena As String
Function GuardarFoto(ByVal sRutaFoto As String) As Byte()
    Dim b() As Byte
    Open sRutaFoto For Binary As #1
    ReDim b(FileLen(sRutaFoto))
    Get #1, , b
    Close #1
    GuardarFoto = b
    
End Function
Public Sub conexion(ByVal StrRuta As String)

   Set CnBd = New ADODB.Connection
   Set cnbd1 = New ADODB.Connection
   Dim sys_DataBase As String
   Dim sys_DataBase1 As String
   Dim sys_SUser As String
   Dim sys_SPassword As String
   Dim sys_ConString As String
   Dim sys_ConString1 As String
   Dim sys_Server As String
 ' CnBd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrRuta & "; "
' CnBd.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=lhorna_serv"
'sys_Server = "190.223.204.246"
 sys_Server = "localhost"
 sys_DataBase = "base1" 'ConfigRead("DataBase")
 'sys_DataBase1 = "factusoft_inventario" 'ConfigRead("DataBase")
 sys_SUser = "root" 'DecryptString(ConfigRead("SUser"))
 sys_SPassword = "" 'DecryptString(ConfigRead("SPassword"))
 sys_ConString = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server & ";" & _
            "Database=" & sys_DataBase & ";" & _
            "UID=" & sys_SUser & ";" & _
            "PWD=" & sys_SPassword & ";" & _
            "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
        
    'Set conConnection = New ADODB.Connection
  '  sys_ConString1 = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server & ";" & _
            "Database=" & sys_DataBase1 & ";" & _
            "UID=" & sys_SUser & ";" & _
            "PWD=" & sys_SPassword & ";" & _
            "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
    CnBd.ConnectionString = sys_ConString
    CnBd.Open
    
   ' cnbd1.ConnectionString = sys_ConString1
    'cnbd1.Open

'strCadena = "SELECT * FROM persona where cpersona='1000'"
'Call ConfiguraRst(strCadena)
'MsgBox Str(rst("cpersona"))
'CnBd.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=factusoft_inventario" ';Data Source=192.168.1.33"
  'CnBd.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=lhorna_tienda"
  'CnBd.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=lhorna_tienda;Data Source=PEP-SERV01"
  'CnBd.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=lhorna_serv;Data Source=PEP-SERV01"
  'CnBd.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=lhorna_tienda;Data Source=acc-serv01"
  'CnBd.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=lhorna_tienda;Data Source=5.205.119.170"
End Sub


' Función para almacenar el gráfico
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ConfiguraRst(ByVal strCadena As String)
  Set Rst = New ADODB.Recordset
  Rst.CursorLocation = adUseClient
  Rst.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
  Rst.ActiveConnection = Nothing
End Sub

Public Function Guardar_Imagen(ADO_Connection As ADODB.Connection, sql As String, Campo_Imagen As String, Optional Path_Imagen As String) As Boolean

Dim rs As Recordset
Dim Stream As ADODB.Stream
'Call conexion(x)
'On Error GoTo Error_Sub

    Set rs = New ADODB.Recordset
    ' Abre el recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, ADO_Connection, adOpenKeyset, adLockOptimistic
        
    'Nuevo objeto ADODB Stream
    Set Stream = New ADODB.Stream
    
    ' dato de tipo binario
    Stream.Type = adTypeBinary
    
    Stream.Open
        ' verifica que la ruta del gráfico no sea una cadena vacía
        If Len(Path_Imagen) <> 0 Then
            
            ' lee la imagen desde el path
            Stream.LoadFromFile Path_Imagen
            ' La guarda en campo
            'x = Stream.Read
          '  strCadena = "select Foto,nombre FROM table1  WHERE id='1'"
          '  Call ConfiguraRst(strCadena)
           
           If rs.RecordCount > 0 Then
               'rs(Campo_Imagen).GetChunk GuardarFoto(Path_Imagen)
             rs("foto").AppendChunk GuardarFoto(Path_Imagen)
               rs.Update
               'rs("nombre") = "-----"
               rs.Update
           End If
           'rs.Fields("Foto").AppendChunk GuardarFoto(Path_Imagen)
            'strCadena = "UPDATE table1 SET Foto='" & x & "' WHERE id='1'"
            'cn.Execute (strCadena)
            ' actualiza los cambios
          '  rs.Update
        End If
    ' cierra el recordset y elimina la referencia
    If rs.State = adStateOpen Then
        rs.Close
    End If
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    'Cierra el objeto ADODBStream y elimina la ref.
    If Stream.State = adStateOpen Then
        Stream.Close
    End If
    If Not Stream Is Nothing Then
        Set Stream = Nothing
    End If
    ' Retorno
    Guardar_Imagen = True
    
Exit Function

'Error_Sub:
 ' If Err.Number <> 0 Then
  '  MsgBox CStr(Err) & "  " & Error, vbExclamation
  'End If
  
End Function

'Función para Obtener la imagen de la bd

'''''''''''''''''''''''''''''''''''''''''''
Public Function Leer_Imagen(ADO_Connection As ADODB.Connection, _
                            sql As String, _
                            Campo_Imagen As String) As Picture

On Error GoTo error_Function

Dim rs As Recordset
Dim Stream As ADODB.Stream


    Set rs = New Recordset


    ' Llena el recordset
    rs.Open sql, ADO_Connection, adOpenKeyset, adLockOptimistic
    
    ' Si no hay registros sale de la función y retorna como _
     resultado un valor Nothing, es decir ninguna imagen

    If rs.RecordCount = 0 Then
       Set Leer_Imagen = Nothing
       rs.Close
       Set rs = Nothing
       Exit Function
    End If
    
    ' Nuevo objeto Stream para poder leer el campo de imagen
    Set Stream = New ADODB.Stream
    
    ' Especifica el tipo de datos ( binario )
    Stream.Type = adTypeBinary
    Stream.Open
    
    If IsNull(rs.Fields(Campo_Imagen).Value) Then
    GoTo error_Function
       Exit Function
    End If
    ' Graba los datos en el objeto stream
    Stream.Write rs.Fields(Campo_Imagen).Value
    
    ' este método graba un  archivo temporal  en disco _
     ( en el app.path que luego se elimina )
    Stream.SaveToFile App.Path & "\temp", adSaveCreateOverWrite
    
    ' Retorna la imagen a la función
    Set Leer_Imagen = LoadPicture(App.Path & "\temp")
    
    
    ' Elimina el archivo temporal
    Kill App.Path & "\temp"
    
    
    'Cierra el recordset y el objeto Stream
    If rs.State = adStateOpen Then
        rs.Close
    End If
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    
    If Stream.State = adStateOpen Then
        Stream.Close
    End If
    If Not Stream Is Nothing Then
        Set Stream = Nothing
    End If

Exit Function

error_Function:
 If Err.Number <> 0 Then
    MsgBox CStr(Err) & "  " & Error, vbExclamation
    ' elimina el temporal
    If Len(Dir(App.Path & "\temp")) Then
       Kill App.Path & "\temp"
    End If
 End If

End Function


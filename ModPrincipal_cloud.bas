Attribute VB_Name = "ModPrincipal_cloud"
Sub Main()
 Call conexion
 frmpanel.Show
 'frmexportar.Show
End Sub

Public Function cambio_venta(ByVal fecha As Date) As Single
nuevo:
If IsDate(fecha) Then

    strCadena = "SELECT valor_venta FROM tipo_cambio WHERE fecha<='" & Format(fecha, "YYYY-mm-dd") & "' AND id_creador='" & KEY_RUC & "' ORDer BY fecha Desc limit 1"
    Call ConfiguraRstLocal(strCadena)
    If rstLocal.RecordCount > 0 Then
        cambio_venta = rstLocal("valor_venta")
    Else
       
        GoTo nuevo
    End If
End If
End Function

Public Sub conexion_cloud(ByVal in_tipo_conexion As String)
Dim in_contador As Integer

   Set CnBd2 = New ADODB.Connection
   Dim sys_DataBase As String
   Dim sys_DataBase2 As String
   Dim sys_SUser2 As String
   Dim sys_SPassword2 As String
   Dim sys_ConString2 As String
   Dim sys_Server2 As String
   
 sys_Server2 = "localhost"
 sys_DataBase2 = "bd_vitekey_repos_ii" 'ConfigRead("DataBase")
 sys_SUser2 = "user_cord" 'DecryptString(ConfigRead("SUser"))
 sys_SPassword2 = "vitekey2018" 'DecryptString(ConfigRead("SPassword"))
 db_port = "3306"
 
' sys_Server2 = "localhost"
' sys_DataBase2 = "ginsacerp" 'ConfigRead("DataBase")
 'sys_SUser2 = "user_cord" 'DecryptString(ConfigRead("SUser"))
 'sys_SPassword2 = "123456" 'DecryptString(ConfigRead("SPassword"))
 'db_port = "3306"
 
 If in_tipo_conexion = "01" Then ' CONEXION MYSQL
            sys_ConString2 = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server2 & ";" & _
            "Database=" & sys_DataBase2 & ";" & _
            "UID=" & sys_SUser2 & ";" & _
            "PWD=" & sys_SPassword2 & ";" & _
            " PORT=" & db_port & ";OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
            CnBd2.ConnectionString = sys_ConString2
            CnBd2.Open
End If
        
If in_tipo_conexion = "02" Then ' CONEXION SQL 2000
    CnBd2.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" & sys_DataBase2
End If

If in_tipo_conexion = "03" Then ' CONEXION SQL 2005
'CnBd2.Open "Provider=SQLNCLI; " & _
             "Initial Catalog=bd; " & _
             "Data Source=192.168.1.35\SQLEXPRESS; " & _
             "integrated security=SSPI; persist security info=True;"
             
 CnBd2.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=bd;Data Source=HP"
             
             
End If

If in_tipo_conexion = "04" Then ' ACCESS
    CnBd2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strRuta & "; "
End If

If in_tipo_conexion = "05" Then
    'PostgreSQLGINSAC
   
     ' sys_ConString2 = "Provider=MSDASQL.1;Persist Security Info=False;User ID=postgres;Data Source=PostgreGinsac;Initial Catalog=vitekey_ginsac"
      sys_ConString2 = "Provider=MSDASQL.1;Persist Security Info=False;User ID=postgres;Data Source=PostgreSQLGINSAC;Initial Catalog=vitekey_ginsac"
   CnBd2.ConnectionString = sys_ConString2
   CnBd2.Open
End If
    


Exit Sub
   
End Sub
Public Sub conexion()

   Set CnBd = New ADODB.Connection

   Dim sys_DataBase As String
   
   Dim sys_SUser As String
   Dim sys_SPassword As String
   Dim sys_ConString As String
   
   
   Dim sys_Server As String
   


strRuta_ini = App.Path & "\archivos\vitekey_local.ini"
FileName = Dir(strRuta_ini)
If FileName = "" Then
    strServer_ini = "158.69.74.64"
    Open strRuta_ini For Output As #1
    Print #1, strServer_ini
    sys_Server = strServer_ini
    Close #1
Else
    Open strRuta_ini For Input As #1
    Line Input #1, sys_Server
    sys_Server = sys_Server
    Close #1
End If


If sys_Server = "158.69.74.64" Then
   KEY_CLOUD = "si"
Else
    KEY_CLOUD = "no"
End If


'sys_Server = "54.149.121.113"
'sys_Server = "localhost"

'sys_Server = "srv-ca.isn.bz"

'sys_Server = "7.75.65.209"

'sys_Server = "dbbasedatos.cyd9c2r3mxxp.sa-east-1.rds.amazonaws.com"
'sys_Server = "192.168.1.106"
'sys_Server = "10.10.6.208"
conectar_nuevamente:

'sys_DataBase = "bd_vitekey_demo" 'ConfigRead("DataBase")
sys_DataBase = "bd_vitekey_aurora" 'ConfigRead("DataBase")
'sys_DataBase = "bd_vitekey_aurora_contable" 'ConfigRead("DataBase")
'sys_DataBase = "bd_vitekey_repos_ii" 'ConfigRead("DataBase")'
'sys_DataBase = "bd_vitekey_conta" 'ConfigRead("DataBase")
 'sys_DataBase = "bd_vitekey_manager" 'ConfigRead("DataBase")
 'sys_DataBase1 = "gigane" 'ConfigRead("DataBase")
sys_SUser = "user_cord" 'DecryptString(ConfigRead("SUser"))
'sys_SPassword = "password" 'DecryptString(ConfigRead("SPassword"))
' sys_SUser = "user1" 'DecryptString(ConfigRead("SUser"))
 'sys_SPassword = "@02021974abc2016@123@cord" 'DecryptString(ConfigRead("SPassword"))
'sys_SPassword = "02021974abc2016" 'DecryptString(ConfigRead("SPassword"))
sys_SPassword = "vitekey2018" 'DecryptString(ConfigRead("SPassword"))
' sys_SPassword = "123456" 'DecryptString(ConfigRead("SPassword"))
 db_port = "3306"
' sys_SPassword = "u9ugejuza" 'DecryptString(ConfigRead("SPassword"))
  sys_ConString = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server & ";" & _
            "Database=" & sys_DataBase & ";" & _
            "UID=" & sys_SUser & ";" & _
            "PWD=" & sys_SPassword & ";" & _
            " PORT=" & db_port & ";OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
        
   'Set conConnection = New ADODB.Connection
   'sys_ConString1 = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server & ";" & _
            "Database=" & sys_DataBase1 & ";" & _
            "UID=" & sys_SUser & ";" & _
            "PWD=" & sys_SPassword & ";" & _
            "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
            
    
    
    CnBd.ConnectionString = sys_ConString
    CnBd.Open

    
   
 Exit Sub

End Sub
Public Sub CenterForm(ByRef ifrmFormulario As Form)
    ifrmFormulario.Move (Screen.Width - ifrmFormulario.Width) / 2, (Screen.Height - ifrmFormulario.Height) / 2
End Sub
Public Function LastRegistroRUC(ByVal Tabla As String, ByVal Campo As String, ByVal cnx As ADODB.Connection) As String
strCadena = "SELECT " & Campo & " FROM " & Tabla & " WHERE ruc='" & KEY_RUC & "'  ORDER BY  " & Campo & " DESC LIMIT 1"
Call ConfiguraRst_cnx(strCadena, cnx)
If rst_cnx.RecordCount > 0 Then
    LastRegistroRUC = rst_cnx(0)
Else
    LastRegistroRUC = 1
End If
End Function
Public Function validar_cadena(ByVal strSql As String, ByVal in_tabla As String, ByVal in_campo_primario As String, ByVal in_primario As Double, ByVal in_campo_secundario As String, ByVal in_secundario As Double) As Boolean
Dim in_codigo As Double
validar_cadena = False

If in_primario > 0 Then
            CnBd2.Execute (strSql)
            in_codigo = Val(LastRegistroRUC(in_tabla, in_campo_primario, CnBd2))
  
              strCadena = "SELECT * FROM entidad_acciones WHERE  id_secundario='" & in_primario & "' and actualizado='no' and  ruc='" & KEY_RUC & "' ORDER BY id ASC"
              Call ConfiguraRst(strCadena)
              If rst.RecordCount > 0 Then
                 rst.MoveFirst
                 For i = 0 To rst.RecordCount - 1
                    CnBd2.Execute (Replace(Replace(rst("cadena"), "´", "'"), in_primario, in_codigo))
                    strCadena = "UPDATE entidad_acciones SET actualizado='si' WHERE id='" & rst("id") & "'"
                    CnBd.Execute (strCadena)
                    rst.MoveNext
                Next i
              End If

              validar_cadena = True
              Exit Function
End If
 validar_cadena = False





End Function
Public Sub actualizar_cloud()
On Error GoTo salir
strCadena = "SELECT * FROM entidad_acciones WHERE id_secundario='0' and actualizado='no' and  (ruc='" & KEY_RUC & "' or ruc='20538975843') ORDER BY id ASC"
Call ConfiguraRstLocal(strCadena)
If rstLocal.RecordCount > 0 Then
   frmcloud.timer_cloud.Enabled = False
   rstLocal.MoveFirst
   For i = 0 To rstLocal.RecordCount - 1
       If validar_cadena(Replace(rstLocal("cadena"), "´", "'"), rstLocal("tabla"), rstLocal("in_campo_unico"), rstLocal("id_unico"), rstLocal("in_campo_secundario"), rstLocal("id_secundario")) = False Then
            CnBd2.Execute (Replace(rstLocal("cadena"), "´", "'"))
      End If
        strCadena = "UPDATE entidad_acciones SET actualizado='si' WHERE id='" & rstLocal("id") & "'"
        CnBd.Execute (strCadena)
        rstLocal.MoveNext
        DoEvents
   Next i
End If
frmcloud.timer_cloud.Enabled = True
Exit Sub

salir:

frmcloud.timer_cloud.Enabled = False
frmcloud.timer_local.Enabled = False
frmcloud.timer_arrancar.Enabled = True
CnBd.Close
CnBd2.Close
End Sub
Public Sub actualizar_local()
On Error GoTo salir_local
strCadena = "SELECT * FROM entidad_acciones WHERE id_secundario='0' and actualizado='no' and  (ruc='" & KEY_RUC & "' or ruc='20538975843') ORDER BY id ASC"
Call ConfiguraRstCloud(strCadena)
If rstCloud.RecordCount > 0 Then
   frmcloud.timer_local.Enabled = False
   rstCloud.MoveFirst
   For i = 0 To rstLocal.RecordCount - 1
       If validar_cadena(Replace(rstCloud("cadena"), "´", "'"), rstCloud("tabla"), rstCloud("in_campo_unico"), rstCloud("id_unico"), rstCloud("in_campo_secundario"), rstCloud("id_secundario")) = False Then
            CnBd.Execute (Replace(rstCloud("cadena"), "´", "'"))
      End If
        strCadena = "UPDATE entidad_acciones SET actualizado='si' WHERE id='" & rstCloud("id") & "'"
        CnBd2.Execute (strCadena)
        rstCloud.MoveNext
        DoEvents
   Next i
End If

frmcloud.timer_local = True
Exit Sub
salir_local:
frmcloud.timer_cloud.Enabled = False
frmcloud.timer_local.Enabled = False
frmcloud.timer_arrancar.Enabled = True
CnBd.Close
CnBd2.Close


End Sub
Public Function numero_registros_cloud(ByVal in_registros As Double) As Double
Dim num_registros As String
strRuta_reloj = App.Path & "\archivos\vitekeycloud.ini"
FileName = Dir(strRuta_reloj)
num_registros = Trim(Str(in_registros))
If FileName = "" Then
    
    Open strRuta_reloj For Output As #1
    Print #1, num_registros
    Close #1
    
Else
    Close #1
    Open strRuta_reloj For Output As #1
    'Line Input #1, num_registros
    Print #1, num_registros
    Close #1
End If

numero_registros_cloud = Val(num_registros)

End Function
Public Sub ConfiguraRstLocal(ByVal strCadena As String)
  Set rstLocal = New ADODB.Recordset
  rstLocal.CursorLocation = adUseClient
  rstLocal.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
  rstLocal.ActiveConnection = Nothing
End Sub
Public Sub ConfiguraRstCloud(ByVal strCadena As String)
  Set rstCloud = New ADODB.Recordset
  rstCloud.CursorLocation = adUseClient
  rstCloud.Open strCadena, CnBd2, adOpenKeyset, adLockOptimistic
  rstCloud.ActiveConnection = Nothing
End Sub
Public Sub ConfiguraRst(ByVal strCadena As String)
  Set rst = New ADODB.Recordset
  rst.CursorLocation = adUseClient
  rst.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
  rst.ActiveConnection = Nothing
End Sub
Public Sub ConfiguraRstK(ByVal strCadena As String)
  Set rstK = New ADODB.Recordset
  rstK.CursorLocation = adUseClient
  rstK.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
  rstK.ActiveConnection = Nothing
End Sub
Public Sub ConfiguraRstL(ByVal strCadena As String)
  Set rstL = New ADODB.Recordset
  rstL.CursorLocation = adUseClient
  rstL.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
  rstL.ActiveConnection = Nothing
End Sub
Public Sub ConfiguraRstMigrar(ByVal strCadena As String)
  Set rstMigrar = New ADODB.Recordset
  rstMigrar.CursorLocation = adUseClient
  rstMigrar.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
  rstMigrar.ActiveConnection = Nothing
End Sub
Public Sub ConfiguraRstP(ByVal strCadena As String)
  Set rstP = New ADODB.Recordset
  rstP.CursorLocation = adUseClient
  rstP.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
  rstP.ActiveConnection = Nothing
End Sub
Public Sub ConfiguraRst2(ByVal strCadena As String)
  Set rst2 = New ADODB.Recordset
  rst2.CursorLocation = adUseClient
  rst2.Open strCadena, CnBd2, adOpenKeyset, adLockOptimistic
  rst2.ActiveConnection = Nothing
End Sub
Public Sub ConfiguraRst3(ByVal strCadena As String)
  Set rst3 = New ADODB.Recordset
  rst3.CursorLocation = adUseClient
  rst3.Open strCadena, CnBd2, adOpenKeyset, adLockOptimistic
  rst3.ActiveConnection = Nothing
End Sub

Public Sub ConfiguraRst_cnx(ByVal strCadena As String, ByVal cnx As ADODB.Connection)
  Set rst_cnx = New ADODB.Recordset
  rst_cnx.CursorLocation = adUseClient
  rst_cnx.Open strCadena, cnx, adOpenKeyset, adLockOptimistic
  rst_cnx.ActiveConnection = Nothing
End Sub
Public Function get_persona(ByVal in_dni As String) As String
strCadena = "SELECT nombre_completo FROM persona where dni='" & in_dni & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    get_persona = rst("nombre_completo")
Else
    get_persona = "-"
End If
End Function
Public Function get_direccion(ByVal in_dni As String) As String
strCadena = "SELECT direccion FROM persona where dni='" & in_dni & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    get_direccion = rst("direccion")
Else
    get_direccion = "-"
End If
End Function

Public Function get_producto_vargas(ByVal in_codigo As String) As String
strCadena = "SELECT movnoa FROM qalmdet where movart='" & in_codigo & "' LIMIT 1"
Call ConfiguraRstMigrar(strCadena)
If rstMigrar.RecordCount > 0 Then
    get_producto_vargas = rstMigrar("movnoa")
Else
    get_producto_vargas = "-"
End If
End Function




Public Function ConsultaUltimoRegistro(ByVal Tabla As String, ByVal Campo As String, ByVal campoRuc As String, ByVal ruc As String) As String
    strCadena = "SELECT * FROM " & Tabla & " WHERE " & campoRuc & "='" & ruc & "' ORDER BY  " & Campo & " DESC LIMIT 1"
    Call ConfiguraRst(strCadena)
    
    If IsNull(rst(0)) = False And rst.RecordCount > 0 Then
        rst.MoveFirst
        ConsultaUltimoRegistro = Val(rst(0) + 1)
    Else
        ConsultaUltimoRegistro = 1
    End If
    
    
End Function

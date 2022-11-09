Attribute VB_Name = "modManejoArchivos"
Option Explicit

Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
  
' UDT para FindFirstFile
Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type
  
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    nFileSizeHigh    As Long
    nFileSizeLow     As Long
    dwReserved0      As Long
    dwReserved1      As Long
    cFileName        As String * MAX_PATH
    cAlternate       As String * 14
End Type
  
  
' Apis para buscar ficheros
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFilename As String, lpFindFileData As WIN32_FIND_DATA) As Long

' Apis para crear carpetas subniveles
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

' Apis para abrir archivos
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
  
' retorna Verdadero si el archivo existe
Public Function ExisteArchivo(ByVal strFile As String) As Boolean
  
    Dim lHandle As Long             ' Handle del archivo
    Dim wFD As WIN32_FIND_DATA      ' udt con los datos
      
    ' Comprobar la barra separadora de path y la longitud
    If ((Len(strFile) > 3) And (Right$(strFile, 1) = "\")) Then
        strFile = Left$(strFile, Len(strFile) - 1)
    End If
      
    lHandle = FindFirstFile(strFile, wFD) ' buscar
      
    ' si el código del handle es válido ...
    ExisteArchivo = lHandle <> INVALID_HANDLE_VALUE
      
    ' Liberar y cerrar el archivo con la función FindClose
    Call FindClose(lHandle)
  
End Function
  
' -----------------------------------------------------------------------------------------
' \\ -- Inicio -- Ejecutar la función para crear el directorio de múltiples niveles
' -----------------------------------------------------------------------------------------
      
    ' -- Crear un directorio de ejemplo -- Retorna True
    'bRet = Create_Directory("c:\carpeta 1\directorio\otro directorio mas")
      
' -----------------------------------------------------------------------------------------
' \\ -- Función de ajuste para ejecutar MakeSureDirectoryPathExists y poder crear los directorios+Subdirectorios
' -----------------------------------------------------------------------------------------
Public Function Create_Directory(ByVal sDirPath As String) As Boolean
      
    On Error GoTo error_Handler
      
    If Right(sDirPath, 1) <> "\" Then sDirPath = sDirPath & "\"
      
    ' -- Linea de código Opcional para  Comprobar previamente si el path ya existe
    If Len(Dir(sDirPath, vbDirectory)) = 0 Then
        Call MakeSureDirectoryPathExists(sDirPath)
        Create_Directory = CBool(Len(Dir(sDirPath, vbDirectory)))
    Else
        Create_Directory = True
    End If
      
    ' -- Errores
    Exit Function
error_Handler:
End Function

'Public Function DownFileFast(NomFile As String) As String
 '   Dim ObA As cbArchivo
  '  Set ObA = New cbArchivo
   ' On Error GoTo Salto
    'ObA.IdEmpresaUsuario = sys_Empresa
    'ObA.Nombre = NomFile
    'ObA.DescargarArchivo (aadEnDisco)
    'DownFileFast = ObA.RutaArchivo
'Salto:
 '   Set ObA = Nothing
'End Function

Public Function AbreArchivo(Archivo As String, Optional Aplicacion As String = vbNullString, Optional Modo As VbAppWinStyle = vbNormalFocus) As Boolean
    On Error GoTo Err_Abrir
    If ExisteArchivo(Archivo) Then
        If Aplicacion <> vbNullString Then
            ShellExecute 0&, vbNullString, ChrW$(34) & Aplicacion & ChrW$(34), ChrW$(34) & Archivo & ChrW$(34), vbNullString, vbNormalFocus
        Else
            ShellExecute 0&, vbNullString, ChrW$(34) & Archivo & ChrW$(34), vbNullString, vbNullString, vbNormalFocus
        End If
    End If
    Exit Function
Err_Abrir:
    MsgBox Err.Description, vbCritical
End Function
''Subir archivos personalizados de empresas
'
'Public Sub SubirArchivoEmpresa(AchName As String, AchPath As String, IdEmpresaUsuario As Long)
'
'    If Trim(AchName) = vbNullString Or Trim(AchPath) = vbNullString Then Exit Sub
'
'    Dim rsTab As ADODB.Recordset
'    Set rsTab = New ADODB.Recordset
'    Dim mystream As ADODB.Stream
'    Set mystream = New ADODB.Stream
'
'    Dim SqlStr As String
'
'    mystream.Type = adTypeBinary
'    rsTab.CursorLocation = adUseClient
'
'    On Error GoTo ErrorSubida
'
'    StrSql = "SELECT * FROM emp_res WHERE IdEmpresaUsuario = " & IdEmpresaUsuario & " AND ERNombre=" & Comillas(AchName)
'
'    rsTab.Open StrSql, conConnection, adOpenStatic, adLockOptimistic
'
'    If rsTab.EOF Then
'        rsTab.AddNew
'    End If
'
'    mystream.Open
'    mystream.LoadFromFile AchPath
'
'
'    rsTab!IdEmpresaUsuario = IdEmpresaUsuario
'    rsTab!ERNombre = AchName
'    rsTab!ERSize = mystream.Size
'    rsTab!ERFile = mystream.Read
'    rsTab.Update
'    mystream.Close
'    rsTab.Close
'    MsgBox "El archivo se subio correctamente"
'
'    Exit Sub
'
'ErrorSubida:
'    mystream.Close
'    rsTab.Close
'    MsgBox "Error: " & Err.Description
'End Sub
'
''Subir archivos generales del sistema
'
'Public Sub SubirArchivoGeneral(AchName As String, AchPath As String, RGTipo As Integer, RGFiltro As String, RGDescripcion As String, RGPersonalizable As Integer)
'
'    If Trim(AchName) = vbNullString Or Trim(AchPath) = vbNullString Then Exit Sub
'
'    Dim rsTab As ADODB.Recordset
'    Set rsTab = New ADODB.Recordset
'    Dim mystream As ADODB.Stream
'    Set mystream = New ADODB.Stream
'
'    Dim SqlStr As String
'
'    mystream.Type = adTypeBinary
'    rsTab.CursorLocation = adUseClient
'
'    On Error GoTo ErrorSubida
'
'    StrSql = "SELECT * FROM res_global WHERE RGNombre=" & Comillas(AchName)
'
'    rsTab.Open StrSql, conConnection, adOpenStatic, adLockOptimistic
'
'    If rsTab.EOF Then
'        rsTab.AddNew
'    End If
'
'    mystream.Open
'    mystream.LoadFromFile AchPath
'
'    rsTab!RINombre = AchName
'    rsTab!RISize = mystream.Size
'    rsTab!RIFile = mystream.Read
'
'    rsTab!RGTipo = RGTipo
'    rsTab!RGFiltro = RGFiltro
'    rsTab!RGDescripcion = RGDescripcion
'    rsTab!RGPersonalizable = RGPersonalizable
'
'    rsTab.Update
'    mystream.Close
'    rsTab.Close
'    MsgBox "El archivo se subio correctamente"
'
'    Exit Sub
'
'ErrorSubida:
'    mystream.Close
'    rsTab.Close
'    MsgBox "Error: " & Err.Description
'End Sub
'
'
'Public Sub DescargarArchivoEmpresa(AchName As String, AchPath As String, IdEmpresaUsuario As Long)
'
'    If Trim(AchName) = vbNullString Or Trim(AchPath) = vbNullString Then Exit Sub
'
'    Dim rsTab As ADODB.Recordset
'    Set rsTab = New ADODB.Recordset
'    Dim mystream As ADODB.Stream
'    Set mystream = New ADODB.Stream
'
'    Dim SqlStr As String
'
'    mystream.Type = adTypeBinary
'    rsTab.CursorLocation = adUseClient
'
'    On Error GoTo ErrorDescarga
'
'    StrSql = "SELECT * FROM emp_res WHERE IdEmpresaUsuario = " & IdEmpresaUsuario & " AND ERNombre=" & Comillas(AchName)
'    rsTab.Open StrSql, conConnection
'
'    If Not rsTab.EOF Then
'        mystream.Open
'        mystream.Write rsTab!ERFile
'        mystream.SaveToFile AchPath, adSaveCreateOverWrite
'        mystream.Close
'    End If
'
'    rsTab.Close
'    Exit Sub
'ErrorDescarga:
'    mystream.Close
'    rsTab.Close
'    MsgBox "Error: " & Err.Description
'End Sub
'
'
'Public Sub SubirImagen(ImgName As String, ImgPath As String)
'
'    If Trim(ImgName) = vbNullString Or Trim(ImgPath) = vbNullString Then Exit Sub
'
'    Dim rsTab As ADODB.Recordset
'    Set rsTab = New ADODB.Recordset
'    Dim mystream As ADODB.Stream
'    Set mystream = New ADODB.Stream
'
'    mystream.Type = adTypeBinary
'    rsTab.CursorLocation = adUseClient
'
'    On Error GoTo ErrorSubida
'    rsTab.Open "SELECT * FROM res_img WHERE RINombre=" & Comillas(ImgName), conConnection, adOpenStatic, adLockOptimistic
'
'    If rsTab.EOF Then
'        rsTab.AddNew
'    End If
'
'    mystream.Open
'    mystream.LoadFromFile ImgPath
'
'    rsTab!RINombre = ImgName
'    rsTab!RISize = mystream.Size
'    rsTab!RIFile = mystream.Read
'    rsTab.Update
'    mystream.Close
'    rsTab.Close
'    MsgBox "El archivo se subio correctamente"
'
'    Exit Sub
'
'ErrorSubida:
'    mystream.Close
'    rsTab.Close
'    MsgBox "Error: " & Err.Description
'End Sub
'
'Public Sub DescargarImagen(ImgName As String, ImgPath As String)
'
'    If Trim(ImgName) = vbNullString Or Trim(ImgPath) = vbNullString Then Exit Sub
'
'    Dim rsTab As ADODB.Recordset
'    Set rsTab = New ADODB.Recordset
'    Dim mystream As ADODB.Stream
'    Set mystream = New ADODB.Stream
'
'    mystream.Type = adTypeBinary
'    rsTab.CursorLocation = adUseClient
'
'    On Error GoTo ErrorDescarga
'
'    rsTab.Open "SELECT * FROM res_img WHERE RINombre=" & Comillas(ImgName), conConnection
'
'    If Not rsTab.EOF Then
'        mystream.Open
'        mystream.Write rsTab!RIFile
'        mystream.SaveToFile ImgPath, adSaveCreateOverWrite
'        mystream.Close
'    End If
'
'    rsTab.Close
'    Exit Sub
'ErrorDescarga:
'    mystream.Close
'    rsTab.Close
'    MsgBox "Error: " & Err.Description
'End Sub
'
'Public Sub SubirImagenEmpresa(ImgName As String, ImgPath As String, IdEmpresa As Long)
'
'    If Trim(ImgName) = vbNullString Or Trim(ImgPath) = vbNullString Then Exit Sub
'
'    Dim rsTab As ADODB.Recordset
'    Set rsTab = New ADODB.Recordset
'    Dim mystream As ADODB.Stream
'    Set mystream = New ADODB.Stream
'
'    mystream.Type = adTypeBinary
'    rsTab.CursorLocation = adUseClient
'
'    On Error GoTo ErrorSubida
'    rsTab.Open "SELECT * FROM res_img WHERE RINombre=" & Comillas(ImgName), conConnection, adOpenStatic, adLockOptimistic
'
'    If rsTab.EOF Then
'        rsTab.AddNew
'    End If
'
'    mystream.Open
'    mystream.LoadFromFile ImgPath
'
'    rsTab!RINombre = ImgName
'    rsTab!RISize = mystream.Size
'    rsTab!RIFile = mystream.Read
'    rsTab.Update
'    mystream.Close
'    rsTab.Close
'    MsgBox "El archivo se subio correctamente"
'
'    Exit Sub
'
'ErrorSubida:
'    mystream.Close
'    rsTab.Close
'    MsgBox "Error: " & Err.Description
'End Sub

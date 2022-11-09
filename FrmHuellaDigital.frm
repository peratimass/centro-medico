VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmHuellaDigital 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   18885
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn cmdactivar 
      Height          =   3375
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5953
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmHuellaDigital.frx":0000
      PICN            =   "FrmHuellaDigital.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   8160
      Top             =   5040
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   5400
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8640
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8760
      Top             =   6120
   End
   Begin VB.Label lblmensaje 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   570
      Left            =   4980
      TabIndex        =   6
      Top             =   6795
      Width           =   195
   End
   Begin VB.Image Image5 
      Height          =   1575
      Left            =   360
      Picture         =   "FrmHuellaDigital.frx":D039
      Top             =   120
      Width           =   5250
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmHuellaDigital.frx":14DE6
      Top             =   5640
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmHuellaDigital.frx":176C0
      Top             =   4680
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmHuellaDigital.frx":19F9A
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3840
      Picture         =   "FrmHuellaDigital.frx":1C874
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label lblcargo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   5640
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   480
      Top             =   -720
      Width           =   375
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Left            =   480
      Top             =   -360
      Width           =   1455
   End
   Begin VB.Label lblnombre 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   4440
      TabIndex        =   3
      Top             =   3720
      Width           =   5175
   End
   Begin VB.Label lblHora 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HORA"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1590
      Left            =   1920
      TabIndex        =   2
      Top             =   7440
      Width           =   6735
   End
   Begin VB.Label lbldni 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   4440
      TabIndex        =   1
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label lblempresa 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   4440
      TabIndex        =   0
      Top             =   4680
      Width           =   6615
   End
   Begin VB.Image imgfoto 
      Height          =   9000
      Left            =   11160
      Picture         =   "FrmHuellaDigital.frx":1F14E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7545
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9135
      Left            =   0
      Top             =   0
      Width           =   18885
   End
End
Attribute VB_Name = "FrmHuellaDigital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strRuta As String
Dim txtUrl As String
Dim limpiar As Boolean

Private Sub ChameleonBtn1_Click()


End Sub

Private Sub cmdcerrar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
End Sub

Private Sub cmdactivar_Click()
'If KEY_FINGERPRINT = "si" Then
 KEY_RUC = "20479779598"
 strCadena = "SELECT CURDATE()"
 Call ConfiguraRstZ(strCadena)
 KEY_FECHA = Format(rstZ(0), "YYYY-mm-dd")
    Call conectar
    Call compara_huella
    Call desconectar

   
'End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 250
KEY_RUC = "20479779598"
'KEY_FECHA = Format(Now(), "YYYY-mm-dd")
Me.lblhora.Caption = Time
KEY_EMPRESA = "KOREA MOTOS S.R.L"
'MDIFrmPrincipal.Toolbar1.Enabled = False
'Call ocultar_barra

End Sub
Private Sub ocultar_barra()
'Dim StartWindow As Long ' Lo primero que tenemos que hacer es localizar la barra de tareas con la instrucción ' de debajo y luego con el manejador pasarsela a la función que la oculta o la muestra
   ' StartWindow = FindWindow("Shell_TrayWnd", vbNullString)
   ' If MsgBox("¿Ocultar barra?", vbInformation + vbYesNo) = vbYes Then
    '   ShowWindow StartWindow, 0&  ' La ocultamos
    'Else
      ' ShowWindow StartWindow, 1& ' La mostramos
    'End If
End Sub
Private Sub Timer1_Timer()
Me.lblhora.Caption = Time
End Sub
Private Sub desconectar()
'==========================================================================='
    ' Uninit scanners
    '==========================================================================='
    Dim ufs_res As UFM_STATUS
    
    Screen.MousePointer = vbHourglass
    ufs_res = UFS_Uninit()
    Screen.MousePointer = vbDefault
    If (ufs_res = UFS_STATUS.OK) Then
        'AddMessage ("UFS_Uninit: OK" & vbCrLf)
    Else
        UFS_GetErrorString ufs_res, m_strError
        'AddMessage ("UFS_Uninit: " & m_strError & vbCrLf)
    End If
    '==========================================================================='
    
    '==========================================================================='
    ' Close database
    '==========================================================================='
    If (m_hDatabase <> 0) Then
        Dim ufd_res As UFD_STATUS
        
        ufd_res = UFD_Close(m_hDatabase)
        If (ufd_res = UFD_STATUS.OK) Then
            'AddMessage ("UFD_Close: OK" & vbCrLf)
        Else
            UFD_GetErrorString ufd_res, m_strError
            'AddMessage ("UFD_Close: " & m_strError & vbCrLf)
        End If
        m_hDatabase = 0
    End If
    
    'lvDatabaseList.ListItems.Clear
    '==========================================================================='
    
    '==========================================================================='
    ' Delete matcher
    '==========================================================================='
    If (m_hMatcher <> 0) Then
        Dim ufm_res As UFM_STATUS

        ufm_res = UFM_Delete(m_hMatcher)
        If (ufm_res = UFM_STATUS.OK) Then
          '  AddMessage ("UFM_Delete: OK" & vbCrLf)
        Else
            UFM_GetErrorString ufm_res, m_strError
           ' AddMessage ("UFM_Delete: " & m_strError & vbCrLf)
        End If
        m_hMatcher = 0
    End If
    '==========================================================================='
    
    
End Sub

Public Function ExtractTemplate(ByRef Template() As Byte, ByRef TemplateSize As Long) As Boolean
    Dim ufs_res As UFS_STATUS
    Dim EnrollQuality As Long

    UFS_ClearCaptureImageBuffer (m_hScanner)
    
    'AddMessage ("Place Finger" & vbCrLf)
    
    TemplateSize = 0
    Do
        ufs_res = UFS_CaptureSingleImage(m_hScanner)
         
        If (ufs_res <> UFS_STATUS.OK) Then
            UFS_GetErrorString ufs_res, m_strError
        '    AddMessage ("UFS_CaptureSingleImage: " & m_strError & vbCrLf)
            ExtractTemplate = False
            Exit Function
        End If

        ufs_res = UFS_Extract(m_hScanner, Template(0), TemplateSize, EnrollQuality)
        If (ufs_res = UFS_STATUS.OK) Then
         '   AddMessage ("UFS_Extract: OK" & vbCrLf)
            Exit Do
        Else
            UFS_GetErrorString ufs_res, m_strError
          '  AddMessage ("UFS_Extract: " & m_strError & vbCrLf)
        End If
    Loop

    ExtractTemplate = True
End Function

Private Sub compara_huella()
    Dim acceso As String * 2
    Dim ufd_res As UFD_STATUS
    Dim ufm_res As UFM_STATUS
    ' Input finger data
    Dim Template(MAX_TEMPLATE_SIZE - 1) As Byte
    Dim TemplateSize As Long
    ' DB data
    Dim DBTemplate() As Byte
    Dim DBTemplateSize() As Long
    Dim DBSerial() As Long
    Dim DBTemplatePtr() As Long
    Dim DBTemplateNum As Long
    '
    Dim MatchIndex As Long
    
    ufd_res = UFD_GetTemplateNumber(m_hDatabase, DBTemplateNum)
    If (ufd_res = UFD_STATUS.OK) Then
        'AddMessage ("UFD_GetTemplateNumber: " & DBTemplateNum & vbCrLf)
    Else
        UFD_GetErrorString ufd_res, m_strError
        'AddMessage ("UFD_GetTemplateNumber: " & m_strError & vbCrLf)
    End If
    
    ReDim DBTemplate(MAX_TEMPLATE_SIZE - 1, DBTemplateNum - 1) As Byte
    ReDim DBTemplateSize(DBTemplateNum - 1) As Long
    ReDim DBSerial(DBTemplateNum - 1) As Long
    
    ' Make template pointer array to pass two dimensional template data
    ReDim DBTemplatePtr(DBTemplateNum - 1) As Long
    Dim i As Long
    For i = 0 To DBTemplateNum - 1
        DBTemplatePtr(i) = VarPtr(DBTemplate(0, i))
    Next
        
    ufd_res = UFD_GetTemplateListWithSerial(m_hDatabase, DBTemplatePtr(0), DBTemplateSize(0), DBSerial(0))
    If (ufd_res <> UFD_STATUS.OK) Then
        UFD_GetErrorString ufd_res, m_strError
        'AddMessage ("UFD_GetTemplateListWithSerial: " & m_strError & vbCrLf)
        Exit Sub
    End If
    
    If (Not ExtractTemplate(Template, TemplateSize)) Then
        Exit Sub
    Else
        'pbImageFrame.Refresh
    End If
    
    Screen.MousePointer = vbHourglass
    ufm_res = UFM_Identify(m_hMatcher, Template(0), TemplateSize, DBTemplatePtr(0), DBTemplateSize(0), DBTemplateNum, 5000, MatchIndex)
    Screen.MousePointer = vbDefault
    If (ufm_res <> UFM_STATUS.OK) Then
        UFM_GetErrorString ufm_res, m_strError
        'AddMessage ("UFM_Identify: " & m_strError & vbCrLf)
        Exit Sub
    End If
    
    If (MatchIndex <> -1) Then
        codigo = DBSerial(MatchIndex)
        
       Dim veri As Boolean
       strCadena = "SELECT E.cod_unico,P.nombre_completo,E.id_cargo,id_empresa_rel,P.foto FROM Fingerprints F,entidad_empresa E,persona P WHERE F.UserID=E.cod_unico AND F.Serial='" & codigo & "' AND id_empresa='" & KEY_RUC & "' AND F.UserID=P.dni AND P.dni=E.cod_unico LIMIT 0,1"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
            Me.lbldni.Caption = rst("cod_unico")
            Me.lblNombre.Caption = rst("nombre_completo")
            If rst("id_empresa_rel") <> "0" Then
                strCadena = "SELECT nombre_completo FROM persona where dni='" & rst("id_empresa_rel") & "'"
                Call ConfiguraRstZ(strCadena)
                If rstZ.RecordCount > 0 Then
                   Me.LblEmpresa.Caption = UCase(rstZ("nombre_completo"))
                Else
                   Me.LblEmpresa.Caption = KEY_EMPRESA
                End If
            Else
                Me.LblEmpresa.Caption = KEY_EMPRESA
            End If
            
            strCadena = "SELECT descripcion FROM persona_cargos where id_cargo='" & rst("id_cargo") & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount > 0 Then
                Me.lblcargo.Caption = rstK("descripcion")
            Else
                Me.lblcargo.Caption = "CARGO NO ASIGNADO"
            End If
             Me.imgfoto.Picture = LoadPicture(App.Path + "\archivos\" + "sinfoto.jpg")
            strRuta = App.Path & "\archivos\" & rst("cod_unico")
            If VerificarFichero(App.Path + "\archivos\" + rst("cod_unico")) = False Then
                Call MkDir(App.Path & "\archivos\" & rst("cod_unico"))
                txtUrl = "http://localhost/public_html/usuarios/" & rst("cod_unico") & "/perfilFoto/" & rst("foto")
                Call descargar(App.Path & "\archivos\" & rst("cod_unico"))
            Else
                If VerificarFichero(strRuta) = False Then
                    Me.imgfoto.Picture = LoadPicture(App.Path + "\archivos\" + "sinfoto.jpg")
                Else
                    On Error GoTo hy
                    Me.imgfoto.Picture = LoadPicture(strRuta & "\" & rst("foto"))
hy:
                End If
            End If
            strCadena = "SELECT * FROM persona_asistencia WHERE dni='" & rst("cod_unico") & "' AND ruc='" & KEY_RUC & "' AND fecha='" & KEY_FECHA & "' ORDER BY id DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
               If rst("id_acceso") = "01" Then
                  acceso = "02"
                 
                  nhora = Format(TimeValue(Format(Time(), "HH:mm:ss")) - TimeValue(Format(rst("hora"), "HH:mm.ss")), "HH:mm:ss")
                  horas_trabajadas = hora_to_numero(nhora)
                  'nhora_inicio = Format(Time, "HH:mm:ss")
                  nhora_inicio = "00:00:00"
                  Me.lblmensaje.Caption = "MARCACION SALIDA"
                Else
                    acceso = "01"
                    horas_trabajadas = 0
                    nhora_inicio = "00:00:00"
                    Me.lblmensaje.Caption = "MARCACION INGRESO"
               End If
                
                
                
            Else
                horas_trabajadas = 0
                acceso = "01"
                'nhora_inicio = Format(Time, "HH:mm:ss")
                nhora_inicio = "00:00:00"
                Me.lblmensaje.Caption = "MARCACION INGRESO"
            End If
            strCadena = "INSERT INTO persona_asistencia(dni,ruc,fecha,hora,hora_inicio,horas_trabajo,id_acceso,dni_save)VALUES('" & Trim(Me.lbldni.Caption) & "','" & KEY_RUC & "','" & KEY_FECHA & "','" & Format(Time, "HH:mm:ss") & "','" & nhora_inicio & "','" & horas_trabajadas & "','" & acceso & "','" & KEY_USUARIO & "')"
            CnBd.Execute (strCadena)
            Call insertar_acciones(strCadena)
            limpiar = True
            Call blanquear
            Exit Sub
            
       End If
        
        
       ' AddMessage ("Identification succeed (Serial = " & DBSerial(MatchIndex) & ")" & vbCrLf)
    Else
        codigo = "Error"
        'MsgBox codigo
       ' AddMessage ("Identification failed" & vbCrLf)
    End If

End Sub
Private Function hora_to_numero(ByVal nhora As String) As Single
Dim horas As Single
Dim minutos As Single

horas = Val(Mid(nhora, 1, 2))
minutos = Val(Mid(nhora, 4, 2))

hora_to_numero = horas + (minutos / 60)

End Function
Private Sub descargar(ByVal ruta As String)
strRuta = ruta
With Inet1
    .AccessType = icUseDefault
    .URL = txtUrl
    .Execute , "GET"
End With
End Sub
Private Sub Inet1_StateChanged(ByVal State As Integer)
Dim vtData As Variant, nomArchivo As String
Dim bDone As Boolean, tempArray() As Byte
nomArchivo = Right(Inet1.URL, Len(Inet1.URL) - InStrRev(Inet1.URL, "/"))

Select Case State
    Case icResponseCompleted
         bDone = False
         filesize = Inet1.GetHeader("Content-length")
         contenttype = Inet1.GetHeader("Content-type")
         
         Open strRuta & "\" & nomArchivo For Binary As #1
         vtData = Inet1.GetChunk(1024, icByteArray)

         DoEvents

         If Len(vtData) = 0 Then
            bDone = True
         End If
            Shape2.Width = 0
    Do While Not bDone

       tempArray = vtData

       Put #1, , tempArray

       Shape2.Width = Shape2.Width + (Len(vtData) * 2) * Shape1.Width / filesize

       vtData = Inet1.GetChunk(1024, icByteArray)
       DoEvents

       If Len(vtData) = 0 Then
          bDone = True
       End If
    Loop

    Close #1

'Carga la imagen
Me.imgfoto.Picture = LoadPicture(strRuta & "\" & nomArchivo)

If Check1 Then Kill App.Path & "\" & nomArchivo

Shape2.Width = 0

End Select

End Sub

Private Sub conectar()
    '==========================================================================='
    ' Initilize scanners
    '==========================================================================='
    Dim ufs_res As UFS_STATUS
    Dim ScannerNumber As Long
     Dim ufd_res As UFD_STATUS
    
    Dim Connection As String
    Dim DataSource As String
    
    Screen.MousePointer = vbHourglass
    ufs_res = UFS_Init()
    Screen.MousePointer = vbDefault
    If (ufs_res = UFS_STATUS.OK) Then
        'AddMessage ("UFS_Init: OK" & vbCrLf)
    Else
        'UFS_GetErrorString ufs_res, m_strError
        'AddMessage ("UFS_Init: " & m_strError & vbCrLf)
        Exit Sub
    End If
    
    ufs_res = UFS_GetScannerNumber(ScannerNumber)
    If (ufs_res = UFS_STATUS.OK) Then
        'AddMessage ("UFS_GetScannerNumber: " & ScannerNumber & vbCrLf)
    Else
        'UFS_GetErrorString ufs_res, m_strError
        'AddMessage ("UFS_GetScannerNumber: " & m_strError & vbCrLf)
        Exit Sub
    End If
    
    If (ScannerNumber = 0) Then
        'AddMessage ("There's no available scanner" & vbCrLf)
        Exit Sub
    Else
        'AddMessage ("First scanner will be used" & vbCrLf)
    End If
    
    ufs_res = UFS_GetScannerHandle(0, m_hScanner)
    If (ufs_res <> UFS_STATUS.OK) Then
        UFS_GetErrorString ufs_res, m_strError
      '  AddMessage ("UFS_GetScannerHandle: " & m_strError & vbCrLf)
        Exit Sub
    End If
    '==========================================================================='
     
    '==========================================================================='
    ' Open database
    '==========================================================================='
   
    strRuta = App.Path & "\UFDatabase.mdb"
    'DataSource = "UFDatabase.mdb"
    '
    '---- cdFileDialog.Filter = "Database Files (*.mdb)|*.mdb"
    '----cdFileDialog.FileName = "UFDatabase.mdb"
    '----cdFileDialog.DefaultExt = "mdb"
    '----cdFileDialog.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
    '----cdFileDialog.ShowOpen
    '----DataSource = cdFileDialog.FileName
            
   ' Connection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrRuta & ";"
    'cnbd1.ConnectionString = Connection
    'cnbd1.Open
  sys_Server = "54.149.121.113"
'sys_Server = "localhost"
' sys_Server = "srv-ca.isn.bz"
'sys_Server = "7.75.65.209"
'sys_Server = "dbbasedatos.cyd9c2r3mxxp.sa-east-1.rds.amazonaws.com"
'sys_Server = "192.168.1.2"
'sys_Server = "10.10.6.208"
sys_DataBase = "bd_vitekey_repos" 'ConfigRead("DataBase")
 'sys_DataBase = "bd_vitekey_manager" 'ConfigRead("DataBase")
 'sys_DataBase1 = "gigane" 'ConfigRead("DataBase")
sys_SUser = "root" 'DecryptString(ConfigRead("SUser"))
sys_SPassword = "password" 'DecryptString(ConfigRead("SPassword"))
 'sys_DataBase1 = "factusoft_inventario" 'ConfigRead("DataBase")
 
 
 sys_ConString = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server & ";" & _
            "Database=" & sys_DataBase & ";" & _
            "UID=" & sys_SUser & ";" & _
            "PWD=" & sys_SPassword & ";" & _
            "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
    ufd_res = UFD_Open(sys_ConString, vbNullString, vbNullString, m_hDatabase)
    
    If (ufd_res = UFD_STATUS.OK) Then
       ' AddMessage ("UFD_Open: OK" & vbCrLf)
    Else
        UFD_GetErrorString ufd_res, m_strError
        'AddMessage ("UFD_Open: " & m_strError & vbCrLf)
    End If
    
    'UpdateDatabaseList
    '==========================================================================='
    
    '==========================================================================='
    ' Create matcher
    '==========================================================================='
    Dim ufm_res As UFM_STATUS

    ufm_res = UFM_Create(m_hMatcher)
    If (ufm_res = UFM_STATUS.OK) Then
        'AddMessage ("UFM_Create: OK" & vbCrLf)
    Else
        UFM_GetErrorString ufm_res, m_strError
        'AddMessage ("UFM_Create: " & m_strError & vbCrLf)
        Exit Sub
    End If
    '==========================================================================='
    
End Sub

Private Sub blanquear()
'Call Timer2_Timer
End Sub
Private Sub Timer2_Timer()
Static X As Integer
If limpiar = True Then

X = X + 1
If X = 5 Then
   Me.lbldni.Caption = ""
   Me.lblNombre.Caption = ""
   Me.LblEmpresa.Caption = ""
   Me.lblcargo.Caption = ""
   Me.imgfoto.Picture = Nothing
   X = 0
   limpiar = False
End If
End If
End Sub

Private Sub Timer3_Timer()
If KEY_FINGERPRINT = "si" Then
Call conectar
Call compara_huella
Call desconectar
Else
     'Me.lblestado.Caption = "PARAMETRO DESCATIVADO"
End If
End Sub

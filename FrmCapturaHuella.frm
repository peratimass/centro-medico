VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmCapturaHuella 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   7770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   16860
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn Command2 
      Height          =   375
      Left            =   12120
      TabIndex        =   17
      Top             =   6120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "         TOMAR INSTANTANEA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MICON           =   "FrmCapturaHuella.frx":0000
      PICN            =   "FrmCapturaHuella.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton cmdIndiceDerecho 
      Caption         =   "DEDO INDICE DERECHO"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   15
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton btnInit 
      Caption         =   "ACTIVAR ESCANER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   60
      Picture         =   "FrmCapturaHuella.frx":2D8D
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   180
      Width           =   3075
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   6960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton btnEnroll 
      Caption         =   "DEDO INDICE IZQUIERDO"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   240
      Width           =   4395
   End
   Begin VB.PictureBox pbImageFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   3240
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   4
      Top             =   720
      Width           =   4335
   End
   Begin VB.ComboBox cbType 
      Height          =   315
      ItemData        =   "FrmCapturaHuella.frx":5D4F
      Left            =   5520
      List            =   "FrmCapturaHuella.frx":5D5C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   16200
      Top             =   120
   End
   Begin VB.TextBox TxtDni 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      TabIndex        =   2
      Text            =   "42546269"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox pbImageFrame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   7680
      ScaleHeight     =   351
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.PictureBox picOutput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   12120
      ScaleHeight     =   5235
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   720
      Width           =   4695
   End
   Begin MSComctlLib.ListView lvDatabaseList 
      Height          =   3795
      Left            =   8280
      TabIndex        =   5
      Top             =   900
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   6694
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VitekeySoft.ChameleonBtn cmdRegistrar 
      Height          =   375
      Left            =   12120
      TabIndex        =   18
      Top             =   6600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "            REGISTRAR NUEVO"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MICON           =   "FrmCapturaHuella.frx":5D7D
      PICN            =   "FrmCapturaHuella.frx":5D99
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn Command1 
      Height          =   375
      Left            =   12120
      TabIndex        =   19
      Top             =   7080
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "         CERRAR SESION          "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      MICON           =   "FrmCapturaHuella.frx":7FF2
      PICN            =   "FrmCapturaHuella.frx":800E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   3720
      TabIndex        =   16
      Top             =   6960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblnombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   720
      TabIndex        =   14
      Top             =   2445
      Width           =   75
   End
   Begin VB.Label lbldni 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   1920
      Width           =   75
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   240
      Picture         =   "FrmCapturaHuella.frx":B145
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   240
      Picture         =   "FrmCapturaHuella.frx":DECB
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   600
      TabIndex        =   12
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATOS PACIENTE:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO DE DISPISITIVO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FOTO PUNTO REGISTRO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   13440
      TabIndex        =   9
      Top             =   360
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   7770
      Left            =   0
      Top             =   0
      Width           =   16860
   End
End
Attribute VB_Name = "FrmCapturaHuella"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub btnSelectionDelete_Click()

End Sub

Private Sub btnClear_Click()

End Sub

Private Sub btnSelectionUpdateTemplate_Click()

End Sub

Private Sub btnSelectionUpdateUserInfo_Click()

End Sub

Private Sub btnDeleteAll_Click()

End Sub

Private Sub btnIdentify_Click()

End Sub

Private Sub btnUninit_Click()

End Sub

Public Sub captura_imagen()
Dim strFoto As String, Numero As Integer
On Error GoTo SALIR
    strRuta = App.Path & "\archivos\" & Trim(Me.txtDni.Text)
    If VerificarFichero(strRuta) = False Then
       Call MkDir(App.Path & "\archivos\" & Trim(Me.txtDni.Text))
       strCadena = "SELECT count(*) FROM persona_foto WHERE dni='" & Trim(Me.txtDni.Text) & "'"
       Call ConfiguraRst(strCadena)
       If IsNull(rst(0)) = True Then
            strFoto = Trim(Me.txtDni.Text & "_" & Str(1) & ".jpg")
            Numero = 1
       Else
            strFoto = Trim(Me.txtDni.Text & "_" & Str(rst(0) + 1) & ".jpg")
            Numero = rst(0) + 1
       End If
       
       strCadena = "INSERT INTO persona_foto (dni,foto,detalle)VALUES('" & Trim(Me.txtDni.Text) & "','" & strFoto & "','" & KEY_FECHA & "')"
       CnBd.Execute (strCadena)
       Call insertar_acciones(strCadena)
       'FrmDetallePersona.Command4.Caption = "ALBUM" + Space(1) + Str(Numero)
       
       SavePicture CaptureWindow(picOutput.hwnd, True, 0, 0, (picOutput.Width / Screen.TwipsPerPixelX) - 4, (picOutput.Height / Screen.TwipsPerPixelY) - 4), strRuta & "\" & strFoto
       Capturar = DIWriteJpg(strRuta & "\", 20, 0)
       Me.picOutput = LoadPicture(strRuta & "\" & strFoto)
       strCadena = "UPDATE persona SET foto='" & strFoto & "' WHERE dni='" & Trim(Me.txtDni.Text) & "'"
       CnBd.Execute (strCadena)
       Call insertar_acciones(strCadena)
    Else
       strCadena = "SELECT count(*) FROM persona_foto WHERE dni='" & Trim(Me.txtDni.Text) & "'"
       Call ConfiguraRst(strCadena)
       If IsNull(rst(0)) = True Then
            strFoto = Trim(Me.txtDni.Text & "_" & Str(1) & ".jpg")
            Numero = 1
       Else
            strFoto = Trim(Me.txtDni.Text & "_" & Str(rst(0) + 1) & ".jpg")
            Numero = rst(0) + 1
       End If
       detalle = "Fecha:" & KEY_FECHA & Chr(13) & KEY_EMPRESA
       strCadena = "INSERT INTO persona_foto (dni,foto,detalle)VALUES('" & Trim(Me.txtDni.Text) & "','" & strFoto & "','" & Trim(detalle) & "')"
       CnBd.Execute (strCadena)
       Call insertar_acciones(strCadena)
       'FrmDetallePersona.Command4.Caption = "ALBUM" + Space(1) + Str(Numero)
       
       SavePicture CaptureWindow(picOutput.hwnd, True, 0, 0, (picOutput.Width / Screen.TwipsPerPixelX) - 4, (picOutput.Height / Screen.TwipsPerPixelY) - 4), strRuta & "\" & strFoto
       Capturar = DIWriteJpg(strRuta & "\", 20, 0)
      Me.picOutput = LoadPicture(strRuta & "\" & strFoto)
       strCadena = "UPDATE persona SET foto='" & strFoto & "' WHERE dni='" & Trim(Me.txtDni.Text) & "'"
       CnBd.Execute (strCadena)
       Call insertar_acciones(strCadena)
    End If
  '  FrmDetallePersona.Procedencia = Neutro
SALIR:
    Exit Sub






'If FrmDetalleEspecialista.Procedencia = Nuevo Then
 '   strRuta = App.Path & "\archivos\" & Trim(FrmDetalleEspecialista.TxtRuc.Text)
  '
   ' If VerificarFichero(strRuta) = False Then
    '   Call MkDir(App.Path & "\archivos\" & Trim(FrmDetalleEspecialista.TxtRuc.Text))
 '      strFoto = Trim(FrmDetalleEspecialista.TxtRuc.Text) & ".jpg"
     ''  SavePicture CaptureWindow(picOutput.hwnd, True, 0, 0, (picOutput.Width / Screen.TwipsPerPixelX) - 4, (picOutput.Height / Screen.TwipsPerPixelY) - 4), strRuta & "\" & strFoto
  '     Capturar = DIWriteJpg(strRuta & "\", 20, 0)
       'FrmDetalleEspecialista.Image1 = LoadPicture(strRuta & "\" & strFoto)
   '    strCadena = "SELECT * FROM persona WHERE dni='" & Trim(FrmDetalleEspecialista.TxtRuc.Text) & "'"
    '   Call ConfiguraRst(strCadena)
      ' If rst.RecordCount > 0 Then
     '       strCadena = "UPDATE persona SET foto='" & strFoto & "' WHERE dni='" & Trim(FrmDetalleEspecialista.TxtRuc.Text) & "'"
        '    CnBd.Execute (strCadena)  Call insertar_acciones(strCadena)
       'Else
         '   FrmDetalleEspecialista.img = strFoto
       'End If
   ' Else
    '   strFoto = Trim(FrmDetalleEspecialista.TxtRuc.Text) & ".jpg"
     '  SavePicture CaptureWindow(picOutput.hwnd, True, 0, 0, (picOutput.Width / Screen.TwipsPerPixelX) - 4, (picOutput.Height / Screen.TwipsPerPixelY) - 4), strRuta & "\" & strFoto
      ' Capturar = DIWriteJpg(strRuta & "\", 20, 0)
      ' FrmDetalleEspecialista.Image1 = LoadPicture(strRuta & "\" & strFoto)
      ' strCadena = "UPDATE persona SET foto='" & strFoto & "' WHERE dni='" & Trim(FrmDetalleEspecialista.TxtRuc.Text) & "'"
      ' CnBd.Execute (strCadena)  Call insertar_acciones(strCadena)
   ' End If
   ' FrmDetalleEspecialista.Procedencia = Neutro
    'Exit Sub
'End If
End Sub

Private Sub desconectar()
On Error GoTo SALIR
'==========================================================================='
    ' Uninit scanners
    '==========================================================================='
    Dim ufs_res As UFM_STATUS
    
    Screen.MousePointer = vbHourglass
    ufs_res = UFS_Uninit()
    Screen.MousePointer = vbDefault
    If (ufs_res = UFS_STATUS.OK) Then
        AddMessage ("UFS_Uninit: OK" & vbCrLf)
    Else
        UFS_GetErrorString ufs_res, m_strError
        AddMessage ("UFS_Uninit: " & m_strError & vbCrLf)
    End If
    '==========================================================================='
    
    '==========================================================================='
    ' Close database
    '==========================================================================='
    If (m_hDatabase <> 0) Then
        Dim ufd_res As UFD_STATUS
        
        ufd_res = UFD_Close(m_hDatabase)
        If (ufd_res = UFD_STATUS.OK) Then
            AddMessage ("UFD_Close: OK" & vbCrLf)
        Else
            UFD_GetErrorString ufd_res, m_strError
            AddMessage ("UFD_Close: " & m_strError & vbCrLf)
        End If
        m_hDatabase = 0
    End If
    
   ' lvDatabaseList.ListItems.Clear
    '==========================================================================='
    
    '==========================================================================='
    ' Delete matcher
    '==========================================================================='
    If (m_hMatcher <> 0) Then
        Dim ufm_res As UFM_STATUS

        ufm_res = UFM_Delete(m_hMatcher)
        If (ufm_res = UFM_STATUS.OK) Then
            AddMessage ("UFM_Delete: OK" & vbCrLf)
        Else
            UFM_GetErrorString ufm_res, m_strError
            AddMessage ("UFM_Delete: " & m_strError & vbCrLf)
        End If
        m_hMatcher = 0
    End If
    '==========================================================================='
     
SALIR:
Exit Sub
End Sub


Private Sub cmdIndiceDerecho_Click()
Static indexs As Integer
Dim ufd_res As UFD_STATUS
On Error GoTo errorHanlder
indexs = indexs + 1
   If (cbType.ListIndex = 0) Then
        UFS_SetTemplateType m_hScanner, UFS_TEMPLATE_TYPE.UFS_TEMPLATE_TYPE_SUPREMA
    ElseIf (cbType.ListIndex = 1) Then
        UFS_SetTemplateType m_hScanner, UFS_TEMPLATE_TYPE.UFS_TEMPLATE_TYPE_ISO19794_2
    ElseIf (cbType.ListIndex = 2) Then
        UFS_SetTemplateType m_hScanner, UFS_TEMPLATE_TYPE.UFS_TEMPLATE_TYPE_ANSI378
    End If
    
    If (Not ExtractTemplate(m_Template1, m_Template1Size)) Then
        Exit Sub
    End If
    
    
    
    AddMessage ("Input user data" & vbCrLf)
    'UserInfoForm.SetAdd
    'UserInfoForm.Show vbModal
    'If (Not UserInfoForm.DialogResult) Then
     '   AddMessage ("User data input is cancelled by user" & vbCrLf)
      '  Exit Sub
    'Else
        pbImageFrame5.Refresh
   ' End If
    
        ufd_res = UFD_AddData(m_hDatabase, Trim(Me.txtDni.Text), indexs, m_Template1(0), m_Template1Size, vbNull, 0, "Indice Izquierdo")
    
    
    If (ufd_res <> UFD_STATUS.OK) Then
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_AddData: " & m_strError & vbCrLf)
    Else
        cbType.Enabled = False
        cbType.Locked = True
    End If
    
    Dim codigo As Double
    Dim rst1 As New ADODB.Recordset
    Set rst1 = Nothing
    strCadena = "SELECT * FROM fingerprints ORDER BY Serial DESC "
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    If rst1.RecordCount > 0 Then
        rst1.MoveFirst
        codigo = Val(rst1("Serial"))
    End If
    Dim X As Integer
    X = codigo
    If X > 0 Then
    Set rst1 = Nothing
    strCadena = "SELECT * FROM fingerprints WHERE Serial=" + Str(codigo)
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    
    Dim rst2 As New ADODB.Recordset
    strCadena = "SELECT * FROM fingerprints"
    rst2.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
    rst2.AddNew
    rst2.Fields("UserID").Value = Trim(Me.txtDni.Text)
    'rst2.Fields("id_empresa").Value = KEY_RUC
    rst2.Fields("FingerIndex").Value = rst1.Fields("FingerIndex")
    rst2.Fields("Template1").Value = rst1.Fields("Template1")
    rst2.Fields("Template2").Value = rst1.Fields("Template2")
    rst2.Fields("Memo").Value = rst1.Fields("Memo")
    rst2.Update
    strCadena = "UPDATE persona SET finger='si' WHERE dni='" & Trim(Me.txtDni.Text) & "'"
    CnBd.Execute (strCadena)
    Call insertar_acciones(strCadena)
    'UpdateDatabaseList
    Set rst2 = Nothing
    strCadena = "DELETE FROM fingerprints WHERE Serial=" + Str(codigo)
    cnbd1.Execute (strCadena)
    'Call captura_imagen
    
    End If
errorHanlder:
Exit Sub

End Sub

Private Sub cmdRegistrar_Click()
Me.txtDni.Text = ""
Me.lblNombre.Caption = ""
Me.lbldni.Caption = ""

Procedencia = Selecionar
FrmPersonal.Show
Exit Sub

End Sub

Private Sub Command1_Click()
Call desconectar
Unload Me
End Sub

Private Sub Command2_Click()
Dim strFoto As String, Numero As Integer


    strRuta = App.Path & "\archivos\" & Trim(Me.txtDni.Text)
    '************* NOMBRE FOTO
    strCadena = "SELECT count(*) FROM persona_foto WHERE dni='" & Trim(Me.txtDni.Text) & "'"
    Call ConfiguraRst(strCadena)
       If IsNull(rst(0)) = True Then
            strFoto = Trim(Me.txtDni.Text & "_" & Str(1) & ".jpg")
            Numero = 1
       Else
            strFoto = Trim(Me.txtDni.Text & "_" & Str(rst(0) + 1) & ".jpg")
            Numero = rst(0) + 1
       End If
      '************************
    
    If VerificarFichero(strRuta) = False Then
       Call MkDir(App.Path & "\archivos\" & Trim(Me.txtDni.Text))
       SavePicture CaptureWindow(picOutput.hwnd, True, 0, 0, (picOutput.Width / Screen.TwipsPerPixelX) - 4, (picOutput.Height / Screen.TwipsPerPixelY) - 4), strRuta & "\" & strFoto
       Capturar = DIWriteJpg(strRuta & "\", 20, 0)
       Me.picOutput = LoadPicture(strRuta & "\" & strFoto)
       strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtDni.Text) & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
            strCadena = "UPDATE persona SET foto='" & strFoto & "' WHERE dni='" & Trim(Me.txtDni.Text) & "'"
            CnBd.Execute (strCadena)
            Call insertar_acciones(strCadena)
            detalle = KEY_FECHA + Space(2) + KEY_EMPRESA
            strCadena = "INSERT INTO persona_foto (dni,foto,detalle)VALUES('" & Trim(Me.txtDni.Text) & "','" & strFoto & "','" & detalle & "')"
            CnBd.Execute (strCadena)
            Call insertar_acciones(strCadena)
       End If
    Else
            
       SavePicture CaptureWindow(picOutput.hwnd, True, 0, 0, (picOutput.Width / Screen.TwipsPerPixelX) - 4, (picOutput.Height / Screen.TwipsPerPixelY) - 4), strRuta & "\" & strFoto
       Capturar = DIWriteJpg(strRuta & "\", 20, 0)
       Me.picOutput = LoadPicture(strRuta & "\" & strFoto)
       strCadena = "UPDATE persona SET foto='" & strFoto & "' WHERE dni='" & Trim(Me.txtDni.Text) & "'"
       CnBd.Execute (strCadena)
       Call insertar_acciones(strCadena)
       detalle = KEY_FECHA + Space(2) + KEY_EMPRESA
       strCadena = "INSERT INTO persona_foto (dni,foto,detalle)VALUES('" & Trim(Me.txtDni.Text) & "','" & strFoto & "','" & detalle & "')"
       CnBd.Execute (strCadena)
       Call insertar_acciones(strCadena)
    End If
    
    Exit Sub


            
            

End Sub

Private Sub pbImageFrame5_Paint()
Dim ufs_res As Long
On Error GoTo errorHanlder
ufs_res = UFS_DrawCaptureImageBuffer(m_hScanner, pbImageFrame5.hDC, pbImageFrame5.ScaleLeft, pbImageFrame5.ScaleTop, pbImageFrame5.ScaleLeft + pbImageFrame5.ScaleWidth, pbImageFrame5.ScaleTop + pbImageFrame5.ScaleHeight, 0)
errorHanlder:
Exit Sub
End Sub
Private Sub pbImageFrame_Paint()
Dim ufs_res As Long
On Error GoTo errorHanlder
ufs_res = UFS_DrawCaptureImageBuffer(m_hScanner, pbImageFrame.hDC, pbImageFrame.ScaleLeft, pbImageFrame.ScaleTop, pbImageFrame.ScaleLeft + pbImageFrame.ScaleWidth, pbImageFrame.ScaleTop + pbImageFrame.ScaleHeight, 0)
errorHanlder:
Exit Sub
End Sub
Public Function ExtractTemplate(ByRef Template() As Byte, ByRef TemplateSize As Long) As Boolean
    Dim ufs_res As UFS_STATUS
    Dim EnrollQuality As Long

    UFS_ClearCaptureImageBuffer (m_hScanner)
    
    AddMessage ("Place Finger" & vbCrLf)
    
    TemplateSize = 0
    Do
        ufs_res = UFS_CaptureSingleImage(m_hScanner)
         
        If (ufs_res <> UFS_STATUS.OK) Then
            UFS_GetErrorString ufs_res, m_strError
            AddMessage ("UFS_CaptureSingleImage: " & m_strError & vbCrLf)
            ExtractTemplate = False
            Exit Function
        End If

        ufs_res = UFS_Extract(m_hScanner, Template(0), TemplateSize, EnrollQuality)
        If (ufs_res = UFS_STATUS.OK) Then
            AddMessage ("UFS_Extract: OK" & vbCrLf)
            Exit Do
        Else
            UFS_GetErrorString ufs_res, m_strError
            AddMessage ("UFS_Extract: " & m_strError & vbCrLf)
        End If
    Loop

    ExtractTemplate = True
End Function
Public Sub AddMessage(ByVal Text As String)
    txtMessage.SelStart = Len(txtMessage.Text)
    txtMessage.SelText = Text
    txtMessage.Refresh
End Sub

Private Sub btnEnroll_Click()
Static indexs As Integer
Dim ufd_res As UFD_STATUS
On Error GoTo errorHanlder
indexs = indexs + 1

   If (cbType.ListIndex = 0) Then
        UFS_SetTemplateType m_hScanner, UFS_TEMPLATE_TYPE.UFS_TEMPLATE_TYPE_SUPREMA
    ElseIf (cbType.ListIndex = 1) Then
        UFS_SetTemplateType m_hScanner, UFS_TEMPLATE_TYPE.UFS_TEMPLATE_TYPE_ISO19794_2
    ElseIf (cbType.ListIndex = 2) Then
        UFS_SetTemplateType m_hScanner, UFS_TEMPLATE_TYPE.UFS_TEMPLATE_TYPE_ANSI378
    End If
    
    If (Not ExtractTemplate(m_Template1, m_Template1Size)) Then
        Exit Sub
    End If
    
    
    
    AddMessage ("Input user data" & vbCrLf)
    'UserInfoForm.SetAdd
    'UserInfoForm.Show vbModal
    'If (Not UserInfoForm.DialogResult) Then
     '   AddMessage ("User data input is cancelled by user" & vbCrLf)
      '  Exit Sub
    'Else
        pbImageFrame.Refresh
   ' End If
    
        ufd_res = UFD_AddData(m_hDatabase, Trim(Me.txtDni.Text), indexs, m_Template1(0), m_Template1Size, vbNull, 0, "Indice Izquierdo")
    
    
    If (ufd_res <> UFD_STATUS.OK) Then
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_AddData: " & m_strError & vbCrLf)
    Else
        cbType.Enabled = False
        cbType.Locked = True
    End If
    
    Dim codigo As Double
    Dim rst1 As New ADODB.Recordset
    Set rst1 = Nothing
    strCadena = "SELECT * FROM fingerprints ORDER BY Serial DESC "
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    If rst1.RecordCount > 0 Then
        rst1.MoveFirst
        codigo = Val(rst1("Serial"))
    End If
    Dim X As Integer
    X = codigo
    If X > 0 Then
    Set rst1 = Nothing
    strCadena = "SELECT * FROM fingerprints WHERE Serial=" + Str(codigo)
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    
    Dim rst2 As New ADODB.Recordset
    
    strCadena = "SELECT * FROM Fingerprints"
    rst2.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
    rst2.AddNew
    rst2.Fields("UserID").Value = Trim(Me.txtDni.Text)
    'rst2.Fields("id_empresa").Value = KEY_RUC
    rst2.Fields("FingerIndex").Value = rst1.Fields("FingerIndex")
    rst2.Fields("Template1").Value = rst1.Fields("Template1")
    rst2.Fields("Template2").Value = rst1.Fields("Template2")
    rst2.Fields("Memo").Value = rst1.Fields("Memo")
    rst2.Update
    strCadena = "UPDATE persona SET finger='si' WHERE dni='" & Trim(Me.txtDni.Text) & "'"
    CnBd.Execute (strCadena)
    Call insertar_acciones(strCadena)
    'UpdateDatabaseList
    Set rst2 = Nothing
    strCadena = "DELETE FROM fingerprints WHERE Serial=" + Str(codigo)
    cnbd1.Execute (strCadena)
    'Call captura_imagen
    
    End If
errorHanlder:
Exit Sub
End Sub

Private Sub btnInit_Click()
Call conectar
Me.cmdRegistrar.Enabled = True
Me.btnEnroll.Enabled = True
Me.cmdIndiceDerecho.Enabled = True
End Sub
Private Sub conectar()
On Error GoTo errorHanlder
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
        Me.lblEstado.Caption = "INICIO CORRECTO"
        'AddMessage ("UFS_Init: OK" & vbCrLf)
    Else
        UFS_GetErrorString ufs_res, m_strError
        'AddMessage ("UFS_Init: " & m_strError & vbCrLf)
        Me.lblEstado.Caption = "EQUIPO INACTIVO"
        Exit Sub
    End If
    
    ufs_res = UFS_GetScannerNumber(ScannerNumber)
    If (ufs_res = UFS_STATUS.OK) Then
        'AddMessage ("UFS_GetScannerNumber: " & ScannerNumber & vbCrLf)
        Me.lblEstado.Caption = "INICIO CORRECTO"
    Else
        UFS_GetErrorString ufs_res, m_strError
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
        'AddMessage ("UFS_GetScannerHandle: " & m_strError & vbCrLf)
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
            
    Connection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strRuta & ";"
    cnbd1.ConnectionString = Connection
    cnbd1.Open
    
    ufd_res = UFD_Open(cnbd1, vbNullString, vbNullString, m_hDatabase)
    
    If (ufd_res = UFD_STATUS.OK) Then
        AddMessage ("UFD_Open: OK" & vbCrLf)
    Else
        UFD_GetErrorString ufd_res, m_strError
        AddMessage ("UFD_Open: " & m_strError & vbCrLf)
    End If
    
    'UpdateDatabaseList
    '==========================================================================='
    
    '==========================================================================='
    ' Create matcher
    '==========================================================================='
    Dim ufm_res As UFM_STATUS

    ufm_res = UFM_Create(m_hMatcher)
    If (ufm_res = UFM_STATUS.OK) Then
        AddMessage ("UFM_Create: OK" & vbCrLf)
    Else
        UFM_GetErrorString ufm_res, m_strError
        AddMessage ("UFM_Create: " & m_strError & vbCrLf)
        Exit Sub
    End If
    '==========================================================================='
errorHanlder:
    Exit Sub
End Sub

Private Sub Form_Load()
CenterForm Me
    Me.Top = 150
    m_hScanner = 0
    m_hDatabase = 0
    m_hMatcher = 0
    
    'lvDatabaseList.ColumnHeaders.Add , , "Serial", 50, lvwColumnLeft
    'lvDatabaseList.ColumnHeaders.Add , , "UserID", 60, lvwColumnLeft
    'lvDatabaseList.ColumnHeaders.Add , , "FingerIndex", 80, lvwColumnLeft
    'lvDatabaseList.ColumnHeaders.Add , , "Template1", 80, lvwColumnLeft
    'lvDatabaseList.ColumnHeaders.Add , , "Template2", 80, lvwColumnLeft
    'lvDatabaseList.ColumnHeaders.Add , , "Memo", 60, lvwColumnLeft
    cbType.ListIndex = 0
    End Sub

Private Sub Timer1_Timer()
'SendMessage mCapHwnd, GET_FRAME, 0, 0

'Copy Current Frame to ClipBoard
'SendMessage mCapHwnd, COPY, 0, 0

'Put ClipBoard's Data to picOutput
'picOutput.Picture = Clipboard.GetData

'Clear ClipBoard
'Clipboard.Clear

End Sub

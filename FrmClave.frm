VERSION 5.00
Begin VB.Form FrmClave 
   BorderStyle     =   0  'None
   Caption         =   "Seguridad"
   ClientHeight    =   6195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn CmdCancelar 
      Height          =   495
      Left            =   10560
      TabIndex        =   3
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "SALIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
      MICON           =   "FrmClave.frx":0000
      PICN            =   "FrmClave.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtUsuario 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   12
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   9720
      TabIndex        =   1
      Top             =   2340
      Width           =   2565
   End
   Begin VB.Timer Timer3 
      Left            =   1320
      Top             =   4440
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   2160
      Tag             =   "Clave de Acceso "
      Top             =   2040
   End
   Begin VB.TextBox TxtClave 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   12
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   9720
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3120
      Width           =   2565
   End
   Begin VitekeySoft.ChameleonBtn CmdAceptar 
      Height          =   495
      Left            =   8640
      TabIndex        =   4
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "INGRESAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
      MICON           =   "FrmClave.frx":3031
      PICN            =   "FrmClave.frx":304D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblversion_barra 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V:03-02-08"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   1905
      TabIndex        =   6
      Top             =   5520
      Width           =   3255
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V:03-02-08"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2040
      TabIndex        =   5
      Top             =   5040
      Width           =   3465
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Height          =   6195
      Left            =   0
      Top             =   0
      Width           =   12630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   4680
      Width           =   75
   End
   Begin VB.Image Image3 
      Height          =   6195
      Left            =   0
      Picture         =   "FrmClave.frx":5632
      Top             =   0
      Width           =   12630
   End
End
Attribute VB_Name = "FrmClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim usuario As String
Dim CONTRASEÑA As String
Dim ALUMNO As String
Dim numreg As Integer
Dim rstusuario As New ADODB.Recordset
Dim empresa As String
Dim N As Double
Dim verifica As Integer
Public Procedencia As EnumProcede

Private Sub cmdAceptar_Click()
Call valida_ingreso

End Sub



Private Sub cmdCancelar_Click()
If MsgBox("¿REALMENTE DESEA SALIR DEL SISTEMA?", vbQuestion + vbYesNo, "¡ATENCIÓN!") = vbYes Then
     strCadena = "DELETE FROM gig_usuarios_online WHERE id_gigane='" & KEY_USUARIO & "'"
    CnBd.Execute (strCadena)
    
  Unload Me
  End
Else
 Me.TxtClave.SetFocus
End If
End Sub

Private Sub DtcUsuarios_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtClave)
End If
End Sub
Private Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub

Private Sub cmdUsuario_Change()

End Sub















Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()

'PlaySound App.Path & "\sonidos\bienvenido.wav"
Me.txtUsuario.SetFocus
 'Call valida_ingreso
End Sub


Private Sub Form_Load()
Dim ultimafecha As String
Dim dias As Integer

CenterForm Me

Me.Top = 1500


MDIFrmPrincipal.Toolbar1.Enabled = False
MDIFrmPrincipal.MnuMantenimientos.Enabled = False
MDIFrmPrincipal.MnuActualizacion.Enabled = False
MDIFrmPrincipal.MnuMovimientos.Enabled = False
MDIFrmPrincipal.MnuReportes.Enabled = False
MDIFrmPrincipal.MnuSeguridad.Enabled = False
'MDIFrmPrincipal.mnucaja.Enabled = False

Me.Timer3.Interval = 100
N = 1
verifica = 0

'strCadena = "SELECT codigo as Codigo,descripcion as Descripcion FROM persona_rubro WHERE codigo='00018' ORDER BY descripcion "

'Me.DtaModoAcceso.BoundText = "00018"
'Call listar_empresas(Me.DtaModoAcceso.BoundText)

'strCadena = "SELECT cod_unico as Codigo,nombre_completo as Descripcion FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni AND id_empresa='0' AND id_tipo_per<>'00009' AND id_tipo_per<>'' AND id_tipo_per<>'00022' AND id_tipo_per<>'00017'"
'Call ConfiguraRst(strCadena)
'm = rst.RecordCount
'Call LlenaDataCombo(Me.DtcEmpresa)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.cmdAceptar.SetFocus
End If
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub Image6_Click()

End Sub

Private Sub Timer3_Timer()
Dim var As String
If (N = Len(Trim(empresa))) Then
    Me.Label5.Caption = Empty
    N = 1
    Me.Timer3.Interval = 1000
Else
Me.Timer3.Interval = 100
var = Mid(empresa, N, 1)
Me.Label5.Caption = Me.Label5.Caption + var
N = N + 1
    If (N = Len(Trim(empresa))) Then
    Me.Timer3.Interval = 1000
    End If
End If
End Sub




Private Sub TxtClave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call valida_ingreso
End If
End Sub


Private Sub valida_ingreso()
usuario = Trim(Me.txtUsuario.Text)
password = Trim(Me.TxtClave.Text)
Dim fecha  As String
Dim cambio As Single
Dim rstB As New ADODB.Recordset
Dim codigo_cambio As String


strCadena = "SELECT dni,a_paterno,nombres FROM view_entidad WHERE dni='" & usuario & "' AND passwordaccesso='" & password & "' and id_personal='si' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    KEY_USUARIO = rst("dni")
    KEY_VENDEDOR = UCase(rst("a_paterno")) + Space(1) + UCase(rst("nombres"))
    KEY_PASSWORD = Trim(Me.TxtClave.Text)
    MDIFrmPrincipal.StatusBar1.Panels(7) = "USER:" + Space(1) + KEY_VENDEDOR
    FrmFechaTrabajo.Show
    Unload Me
    Exit Sub
    
    
   
    
   
Else
    MsgBox "PASSWORD INCORRECTO", vbCritical, "INTENTE NUEVAMENTE"
    Call Resalta(Me.TxtClave)
    Exit Sub
End If

       
End Sub

Private Sub TxtUsuario_KeyPress(KeyAscii As Integer)
If (Len(Me.txtUsuario.Text) > 0 And KeyAscii = 13) Then
    Call Resalta(Me.TxtClave)
End If
End Sub
Public Sub AddMessage(ByVal Text As String)
    txtMessage.SelStart = Len(txtMessage.Text)
    txtMessage.SelText = Text
    txtMessage.Refresh
End Sub

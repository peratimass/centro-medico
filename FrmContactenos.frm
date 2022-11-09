VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmContactenos 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Datos remotos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   7935
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1320
         TabIndex        =   22
         Text            =   "smtp.live.com"
         Top             =   360
         Width           =   2400
      End
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   4440
         TabIndex        =   21
         Text            =   "percy19_is@hotmail.com"
         Top             =   360
         Width           =   1170
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6600
         PasswordChar    =   "*"
         TabIndex        =   20
         Text            =   "200119828372000"
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label lblServer 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor SMTP"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   195
         Left            =   3840
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Password:"
         Height          =   195
         Left            =   5760
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "DATOS DE MAIL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9720
      TabIndex        =   12
      Top             =   1800
      Width           =   8415
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   315
         Left            =   7440
         TabIndex        =   16
         Top             =   360
         Width           =   315
      End
      Begin VB.TextBox txtAttach 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   4800
         TabIndex        =   15
         Top             =   360
         Width           =   2520
      End
      Begin VB.TextBox txtTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   600
         TabIndex        =   14
         Text            =   "percy19_is@hotmail.com"
         Top             =   360
         Width           =   3120
      End
      Begin VB.TextBox txtFrom 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   4800
         TabIndex        =   13
         Text            =   "percy19_is@hotmail.com"
         Top             =   0
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.Label lblAttach 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adjunto"
         Height          =   195
         Left            =   4080
         TabIndex        =   18
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Para"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8415
      Begin VB.ListBox lstStatus 
         BackColor       =   &H00808080&
         ForeColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   7800
      End
      Begin VB.Label lblProgress 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Progreso"
         Height          =   195
         Left            =   600
         TabIndex        =   11
         Top             =   240
         Width           =   630
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "SI CUENTA CON INTERNET ENVIE UN MENSAJE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   8415
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Salir"
         Height          =   435
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2280
         Width           =   1395
      End
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   1380
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   840
         Width           =   6960
      End
      Begin VB.TextBox txtSubject 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   6720
      End
      Begin VB.CommandButton cmdSend 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Enviar Email"
         Height          =   435
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2280
         Width           =   1635
      End
      Begin VB.Label lblMsg 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   600
      End
      Begin VB.Label lblSubject 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asunto:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.TextBox TxtPersona 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   840
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7200
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MAIL : percy19_is@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5760
      TabIndex        =   2
      Top             =   3135
      Width           =   2295
   End
   Begin VB.Image Image8 
      Height          =   255
      Left            =   5400
      Picture         =   "FrmContactenos.frx":0000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WEB  : www.vitekey.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5760
      TabIndex        =   33
      Top             =   2775
      Width           =   1815
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   5400
      Picture         =   "FrmContactenos.frx":058A
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CELL : 942867953"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5760
      TabIndex        =   32
      Top             =   2415
      Width           =   1305
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   5400
      Picture         =   "FrmContactenos.frx":0B14
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RPM : #156042"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5760
      TabIndex        =   31
      Top             =   2055
      Width           =   1125
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   5400
      Picture         =   "FrmContactenos.frx":109E
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   1695
      Left            =   5280
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENTA DE DOMINIOS Y ALQUILER DE HOSTING."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   720
      TabIndex        =   30
      Top             =   3255
      Width           =   3480
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   360
      Picture         =   "FrmContactenos.frx":1628
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CABLEADO ESTRUCUTURADO."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   720
      TabIndex        =   29
      Top             =   2895
      Width           =   2220
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   360
      Picture         =   "FrmContactenos.frx":1BB2
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESARROLLO DE DE SISTEMAS WEB CORPORATIVOS."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   720
      TabIndex        =   28
      Top             =   2535
      Width           =   3930
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   360
      Picture         =   "FrmContactenos.frx":213C
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESARROLLO DE SISTEMAS INFORMATICOS A MEDIDA."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   720
      TabIndex        =   27
      Top             =   2175
      Width           =   4065
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   360
      Picture         =   "FrmContactenos.frx":26C6
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VITEKEY SOFTWARE CORPORATION S.A.C"
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
      TabIndex        =   26
      Top             =   1800
      Width           =   3360
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   1935
      Left            =   120
      Top             =   1680
      Width           =   8415
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   6855
      Left            =   0
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "FrmContactenos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Importante: El ocx produce un pequeño Bug y es que no se debe _
establecer la propiedad Visible en False. Por lo tanto para que no se _
vea el ocx en tiempo de ejecución, se estableció el Left y el Top fuera del _
area del form en el Form_Load

'Si tenes idea porque puede ocurrir esto me podes enviar un mail a _
info@recursosvisualbasic.com.ar para ver si lo puedo arreglar y volver a _
compilar.



'Option Explicit
Option Compare Text


Private Sub cmdSend_Click()
Call enviar
End Sub
Private Sub enviar()
        Dim introduccion As String
        Dim reporte As String
        Dim FechaEnvio As String
        
    
'-------------------------------------
    cmdSend.Enabled = False
    lstStatus.Clear
    Screen.MousePointer = vbHourglass

    With sendmail1

        'Valida (opcional)
        .SMTPHostValidacion = VALIDATE_HOST_NONE
        'Valida la sintaxis de l cuenta (opcional)
        .ValidarEmail = VALIDATE_SYNTAX
        'Opcional
        .Delimitador = ";"
        'Texto  para visualizar en el campo De (opcional)
        .FromDisplayName = KEY_EMPRESA
        'Requerido (Nombre del servidor SMTP)
        .SMTPHost = txtServer.Text
        'Requerido
        .Remitente = txtFrom.Text
        'Requerido
        .Destinatario = txtTo.Text
        'Asunto del mensaje
        .Asunto = txtSubject.Text + Space(2) + "para" + Space(2) + Trim(KEY_FECHA) + Space(1) + str(Time)
        'Cuerpodel mensaje
        
        .Mensaje = Trim(Me.txtMsg.Text)
               
        'Adjunto (opcional)
        .Adjunto = Trim(txtAttach.Text)
        
        'Opcional (Prioridad del mensaje)
        .Prioridad = Alta
        'Opcional (si requiere autentificación)
        .UsarLoginSMTP = True
        'Requerido si requiere autentificación
        .usuario = txtUserName
        .password = TxtPassword
        
        txtServer.Text = .SMTPHost
       'Opcional (por defectoutiliza el Tipo MIME)
       .Codificacion = MIME_ENCODE
       
       'Envia el Mail
       .EnviarEmail
    
    End With
    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True
    FechaEnvio = Format$(Now(), "yyyy-mm-dd")
    strCadena = "INSERT INTO persona_mail(dni,fecha,motivo,detalle,ruc)VALUES('" & Val(Me.TxtPersona.Text) & "','" & FechaEnvio & "','" & Trim(Me.txtSubject.Text) & "','" & Trim(Me.txtMsg.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
     

Unload Me

End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub



Private Sub Command1_Click()
FrmMailEnviado.Show
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 800
sendmail1.Move -1000, -1000
End Sub

Private Sub sendmail1_SendSuccesful()

   ' MsgBox "Mensaje enviado correctamente", vbInformation, "Estado de envío"
    lblProgress = ""
End Sub

Private Sub sendmail1_Progress(lPercentCompete As Long)
    'Visualiza el porcentaje del progreso del envío en el Label
    lblProgress = lPercentCompete & "% completado"

End Sub

Private Sub sendmail1_SendFailed(Explanation As String)
    'En caso de fallar el envío se dispara este evento con la descripción del error
    MsgBox ("El envío del Email falló por las posibles razones:: " & vbCrLf & Explanation)
    lblProgress = ""
    Screen.MousePointer = vbDefault
    cmdSend.Enabled = True
    
End Sub



Private Sub sendmail1_Status(Status As String)
    'Visualiza el estado del envío
    lstStatus.AddItem Status
    lstStatus.ListIndex = lstStatus.ListCount - 1
    lstStatus.ListIndex = -1

End Sub



'Para los adjuntos

Private Sub cmdBrowse_Click()

    Dim ArchivosAdj()    As String
    Dim i               As Integer
    
    On Local Error GoTo ErrSub
    
    With CommonDialog1
        .FileName = ""
        .CancelError = True
        .Filter = "All Files (*.*)|*.*|HTML Files (*.htm;*.html;*.shtml)|*.htm;*.html;*.shtml|Images (*.bmp;*.jpg;*.gif)|*.bmp;*.jpg;*.gif"
        .FilterIndex = 1
        .DialogTitle = "Select File Attachment(s)"
        .MaxFileSize = &H7FFF
        .Flags = &H4 Or &H800 Or &H40000 Or &H200 Or &H80000
        .ShowOpen
        ArchivosAdj = Split(.FileName, vbNullChar)
    End With
    
    If UBound(ArchivosAdj) = 0 Then
        If txtAttach.Text = "" Then
            txtAttach.Text = ArchivosAdj(0)
        Else
            txtAttach.Text = txtAttach.Text & ";" & ArchivosAdj(0)
        End If
    ElseIf UBound(ArchivosAdj) > 0 Then
        If Right$(ArchivosAdj(0), 1) <> "\" Then ArchivosAdj(0) = ArchivosAdj(0) & "\"
        For i = 1 To UBound(ArchivosAdj)
            If txtAttach.Text = "" Then
                txtAttach.Text = ArchivosAdj(0) & ArchivosAdj(i)
            Else
                txtAttach.Text = txtAttach.Text & ";" & ArchivosAdj(0) & ArchivosAdj(i)
            End If
        Next
    Else
        Exit Sub
    End If
    
Exit Sub
ErrSub:
MsgBox Err.Description, vbCritical, "Error"

End Sub










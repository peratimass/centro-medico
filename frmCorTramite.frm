VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCorTramite 
   BorderStyle     =   0  'None
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   19305
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn cmdLegalizar 
      Height          =   495
      Left            =   16680
      TabIndex        =   87
      Top             =   1440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "LEGALIZAR DOCUMENTOS   "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmCorTramite.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtPlaca 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   84
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtTitulo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   82
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtMotor 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8640
      TabIndex        =   80
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtSerie 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6360
      TabIndex        =   78
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdLimpiarDNI 
      Caption         =   "x"
      Height          =   255
      Left            =   3360
      TabIndex        =   77
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtDNI 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   73
      Top             =   480
      Width           =   1695
   End
   Begin VB.Frame fraPlaCli 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2880
      TabIndex        =   57
      Top             =   1920
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command3 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   60
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   360
         TabIndex        =   59
         Top             =   1080
         Width           =   4095
      End
      Begin VB.CommandButton cmdConfRecojoPlaCli 
         Caption         =   "Confirmar"
         Height          =   255
         Left            =   1560
         TabIndex        =   58
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RECOJO DE PLACAS POR CLIENTE"
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
         Left            =   360
         TabIndex        =   62
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Height          =   165
         Left            =   375
         TabIndex        =   61
         Top             =   840
         Width           =   1245
      End
   End
   Begin VB.Frame fraRecojoPla 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2880
      TabIndex        =   51
      Top             =   2160
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdConfirmarRecojoPla 
         Caption         =   "Confirmar"
         Height          =   255
         Left            =   1560
         TabIndex        =   54
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   360
         TabIndex        =   53
         Top             =   1080
         Width           =   4095
      End
      Begin VB.CommandButton Command15 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   52
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Height          =   165
         Left            =   375
         TabIndex        =   56
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RECOJO DE PLACAS"
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
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame fraPlaLis 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2880
      TabIndex        =   45
      Top             =   1920
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command14 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   48
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   360
         TabIndex        =   47
         Top             =   1080
         Width           =   4095
      End
      Begin VB.CommandButton cmdConfirmarPlaLis 
         Caption         =   "Confirmar"
         Height          =   255
         Left            =   1560
         TabIndex        =   46
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PLACAS LISTAS"
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
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Height          =   165
         Left            =   375
         TabIndex        =   49
         Top             =   840
         Width           =   1245
      End
   End
   Begin VB.Frame fraPago 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2880
      TabIndex        =   39
      Top             =   2040
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdConfirmarPago 
         Caption         =   "Confirmar"
         Height          =   255
         Left            =   1560
         TabIndex        =   42
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   360
         TabIndex        =   41
         Top             =   1080
         Width           =   4095
      End
      Begin VB.CommandButton Command11 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   40
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Height          =   165
         Left            =   375
         TabIndex        =   44
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PAGO DE PLACAS"
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
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame fraRegPlaca 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2880
      TabIndex        =   33
      Top             =   1920
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command10 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   36
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   360
         TabIndex        =   35
         Top             =   1080
         Width           =   4095
      End
      Begin VB.CommandButton cmdConfirmarRegPlaca 
         Caption         =   "Confirmar"
         Height          =   255
         Left            =   1560
         TabIndex        =   34
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "REGISTRO DE PLACA EN LA WEB"
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
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Height          =   165
         Left            =   375
         TabIndex        =   37
         Top             =   840
         Width           =   1245
      End
   End
   Begin VB.Frame fraRecojoExpRP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2880
      TabIndex        =   27
      Top             =   1920
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtNroPlaca 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   72
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton Command8 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   30
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   360
         TabIndex        =   29
         Top             =   1560
         Width           =   4095
      End
      Begin VB.CommandButton cmdConfirmarRecojoExp 
         Caption         =   "Confirmar"
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NRO DE PLACA :"
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
         Height          =   165
         Left            =   390
         TabIndex        =   71
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label fraRecojoExp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RECOJO DE EXPEDIENTE DE REGISTRO PUBLICO"
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
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Height          =   165
         Left            =   375
         TabIndex        =   31
         Top             =   1320
         Width           =   1245
      End
   End
   Begin VB.Frame fraExpIns 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2880
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdConfirmarExpIns 
         Caption         =   "Confirmar"
         Height          =   255
         Left            =   1560
         TabIndex        =   24
         Top             =   2520
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   360
         TabIndex        =   23
         Top             =   1080
         Width           =   4095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   22
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Height          =   165
         Left            =   375
         TabIndex        =   26
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "EXPEDIENTE INSCRITO EN RRPP"
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
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame fraIngrExpRR 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2880
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command4 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   18
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   360
         TabIndex        =   17
         Top             =   1080
         Width           =   4095
      End
      Begin VB.CommandButton cmdConfirmarIngrExpRR 
         Caption         =   "Confirmar"
         Height          =   255
         Left            =   1560
         TabIndex        =   16
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "INGRESO DE EXPEDIENTE A RRPP"
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
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Height          =   165
         Left            =   375
         TabIndex        =   19
         Top             =   840
         Width           =   1245
      End
   End
   Begin VB.Frame fraExpediente 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   2880
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtMonto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   70
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtAnio 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   68
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtNroRecibo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   66
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtNroTitulo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   65
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton cmdConfirmarArmado 
         Caption         =   "Confirmar"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   360
         TabIndex        =   11
         Top             =   2280
         Width           =   4095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   10
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO :"
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
         Height          =   165
         Left            =   840
         TabIndex        =   69
         Top             =   1680
         Width           =   705
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AÑO :"
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
         Height          =   165
         Left            =   1080
         TabIndex        =   67
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NRO DE RECIBO :"
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
         Height          =   165
         Left            =   240
         TabIndex        =   64
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NRO DE TÍTULO :"
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
         Height          =   165
         Left            =   240
         TabIndex        =   63
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Height          =   165
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ARMADO DE EXPEDIENTE PARA RRPP"
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
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame fraLegal 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   2880
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdCerrarLeg 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   8
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtObservacionLeg 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   4095
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Confirmar"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblEtiqueta 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LEGALIZACIÓN EL FORMULARIO DE INMATRICULACION, MEDIO DE PAGO Y EL FORMULARIO DE PLACAS"
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
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Height          =   165
         Left            =   375
         TabIndex        =   5
         Top             =   840
         Width           =   1245
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridTramites 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   16245
      _ExtentX        =   28654
      _ExtentY        =   7435
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridDetalles 
      Height          =   3255
      Left            =   240
      TabIndex        =   1
      Top             =   5760
      Width           =   16245
      _ExtentX        =   28654
      _ExtentY        =   5741
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   9360
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorTramite.frx":001C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorTramite.frx":007A
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorTramite.frx":00D8
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorTramite.frx":0136
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorTramite.frx":0194
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorTramite.frx":01F2
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorTramite.frx":0250
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorTramite.frx":02AE
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorTramite.frx":030C
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorTramite.frx":036A
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorTramite.frx":03C8
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCorTramite.frx":0426
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin VitekeySoft.ChameleonBtn cmdSalir 
      Height          =   495
      Left            =   16680
      TabIndex        =   86
      Top             =   7440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "CERRAR PANTALLA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmCorTramite.frx":0484
      PICN            =   "frmCorTramite.frx":04A0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdIngresoRP 
      Height          =   495
      Left            =   16680
      TabIndex        =   89
      Top             =   2640
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "INGRESO A RRPP                 "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmCorTramite.frx":34B5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdExpedienteListo 
      Height          =   495
      Left            =   16680
      TabIndex        =   90
      Top             =   3240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "EXPEDIENTE INSCRITO EN RRPP   "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmCorTramite.frx":34D1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdRecojoRP 
      Height          =   495
      Left            =   16680
      TabIndex        =   91
      Top             =   3840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "   RECOJO EXPEDIENTE DE RRPP    "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmCorTramite.frx":34ED
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdRegPlacaWeb 
      Height          =   495
      Left            =   16680
      TabIndex        =   92
      Top             =   4440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "REGISTRO EXPEDIENTE EN LA WEB     "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmCorTramite.frx":3509
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdPagoPlacas 
      Height          =   495
      Left            =   16680
      TabIndex        =   93
      Top             =   5040
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "PAGO DE PLACAS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmCorTramite.frx":3525
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdPlacasListas 
      Height          =   495
      Left            =   16680
      TabIndex        =   94
      Top             =   5640
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "PAGO DE PLACAS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmCorTramite.frx":3541
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdRecojoPlacasInt 
      Height          =   495
      Left            =   16680
      TabIndex        =   95
      Top             =   6240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "RECOJO DE PLACAS INTERNO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmCorTramite.frx":355D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdRecojoPlacasCli 
      Height          =   495
      Left            =   16680
      TabIndex        =   96
      Top             =   6840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "RECOJO PLACAS POR CLIENTE"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmCorTramite.frx":3579
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdArmado 
      Height          =   495
      Left            =   16680
      TabIndex        =   88
      Top             =   2040
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "ARMAR DOCUMENTOS        "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmCorTramite.frx":3595
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLACA : "
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
      Height          =   165
      Left            =   3240
      TabIndex        =   85
      Top             =   960
      Width           =   705
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TITULO : "
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
      Height          =   165
      Left            =   480
      TabIndex        =   83
      Top             =   960
      Width           =   765
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOTOR : "
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
      Height          =   165
      Left            =   7815
      TabIndex        =   81
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SERIE : "
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
      Height          =   165
      Left            =   5655
      TabIndex        =   79
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblNombre 
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
      Left            =   4440
      TabIndex        =   76
      Top             =   480
      Width           =   2475
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE :"
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
      Left            =   5400
      TabIndex        =   75
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC / DNI :"
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
      Height          =   165
      Left            =   360
      TabIndex        =   74
      Top             =   480
      Width           =   885
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TRÁMITE DE DOCUMENTOS"
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
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   2235
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9120
      Left            =   0
      Top             =   0
      Width           =   19305
   End
End
Attribute VB_Name = "frmCorTramite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub cmdArmado_Click()
  Me.fraExpediente.Visible = True
End Sub

Private Sub cmdCerrarLeg_Click()
  Me.fraLegal.Visible = False
End Sub

Private Sub cmdConfirmar_Click()
  strCadena = " update imp_tramite set id_estado = '01', id_estado_detalle = '01' where id_tramite = " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0)
  CnBd.Execute strCadena
  
  strCadena = " insert into imp_tramide_detalle (id_tramite, id_movimiento, observacion , id_autor , fecha_entrada, hora_entrada) values ( " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0) & ",'01','" & Me.txtObservacionLeg.Text & "','" & KEY_USUARIO & "', CURDATE(), CURTIME())"
  CnBd.Execute strCadena
  
  limpiarFrames
  
  Call actualizaListaConFila(Me.gridTramites)
End Sub

Private Sub cmdConfirmarArmado_Click()
  strCadena = " update imp_tramite set id_estado_detalle = '02', anio ='" & Me.txtAnio.Text & "', nro_titulo = '" & txtNroTitulo.Text & _
  "', nro_recibo ='" & txtNroRecibo.Text & "', monto = " & txtMonto.Text & _
  " where id_tramite = " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0)
  
  CnBd.Execute strCadena
  
  strCadena = " insert into imp_tramide_detalle (id_tramite, id_movimiento, observacion , id_autor , fecha_entrada, hora_entrada) values ( " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0) & ",'02','" & Me.Text1.Text & "','" & KEY_USUARIO & "', CURDATE(), CURTIME())"
  
  CnBd.Execute strCadena
  
  limpiarFrames
  
  Call actualizaListaConFila(Me.gridTramites)
End Sub


Private Sub cmdConfirmarExpIns_Click()
  strCadena = " update imp_tramite set id_estado_detalle = '04' where id_tramite = " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0)
  CnBd.Execute strCadena
  
  strCadena = " insert into imp_tramide_detalle (id_tramite, id_movimiento, observacion , id_autor , fecha_entrada, hora_entrada) values ( " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0) & ",'04','" & Me.Text3.Text & "','" & KEY_USUARIO & "', CURDATE(), CURTIME())"
  CnBd.Execute strCadena
  
  limpiarFrames
  
  Call actualizaListaConFila(Me.gridTramites)
End Sub


Private Sub cmdConfirmarIngrExpRR_Click()
  strCadena = " update imp_tramite set id_estado_detalle = '03' where id_tramite = " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0)
  CnBd.Execute strCadena
  
  strCadena = " insert into imp_tramide_detalle (id_tramite, id_movimiento, observacion , id_autor , fecha_entrada, hora_entrada) values ( " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0) & ",'03','" & Me.Text2.Text & "','" & KEY_USUARIO & "', CURDATE(), CURTIME())"
  CnBd.Execute strCadena
  
  limpiarFrames
  
  Call actualizaListaConFila(Me.gridTramites)
End Sub


Private Sub cmdConfirmarPago_Click()
  strCadena = " update imp_tramite set id_estado_detalle = '07' where id_tramite = " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0)
  CnBd.Execute strCadena
  
  strCadena = " insert into imp_tramide_detalle (id_tramite, id_movimiento, observacion , id_autor , fecha_entrada, hora_entrada) values ( " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0) & ",'07','" & Me.Text6.Text & "','" & KEY_USUARIO & "', CURDATE(), CURTIME())"
  CnBd.Execute strCadena
  
  limpiarFrames
  
  Call actualizaListaConFila(Me.gridTramites)
End Sub


Private Sub cmdConfirmarPlaLis_Click()
    strCadena = " update imp_tramite set id_estado_detalle = '08' where id_tramite = " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0)
  CnBd.Execute strCadena
  
  strCadena = " insert into imp_tramide_detalle (id_tramite, id_movimiento, observacion , id_autor , fecha_entrada, hora_entrada) values ( " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0) & ",'08','" & Me.Text7.Text & "','" & KEY_USUARIO & "', CURDATE(), CURTIME())"
  CnBd.Execute strCadena
  
  limpiarFrames
  
  Call actualizaListaConFila(Me.gridTramites)
End Sub


Private Sub cmdConfirmarRecojoExp_Click()
  strCadena = " update imp_tramite set id_estado_detalle = '05', nro_placa = '" & Me.txtNroPlaca.Text & "' where id_tramite = " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0)
  CnBd.Execute strCadena
  
  strCadena = " insert into imp_tramide_detalle (id_tramite, id_movimiento, observacion , id_autor , fecha_entrada, hora_entrada) values ( " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0) & ",'05','" & Me.Text4.Text & "','" & KEY_USUARIO & "', CURDATE(), CURTIME())"
  CnBd.Execute strCadena
  
  limpiarFrames
  
  Call actualizaListaConFila(Me.gridTramites)
End Sub

Private Sub cmdConfirmarRecojoPla_Click()
  strCadena = " update imp_tramite set id_estado_detalle = '09' where id_tramite = " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0)
  CnBd.Execute strCadena
  
  strCadena = " insert into imp_tramide_detalle (id_tramite, id_movimiento, observacion , id_autor , fecha_entrada, hora_entrada) values ( " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0) & ",'09','" & Me.Text8.Text & "','" & KEY_USUARIO & "', CURDATE(), CURTIME())"
  CnBd.Execute strCadena
  
  limpiarFrames
  
  Call actualizaListaConFila(Me.gridTramites)
End Sub

Private Sub cmdConfirmarRegPlaca_Click()
  strCadena = " update imp_tramite set id_estado_detalle = '06' where id_tramite = " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0)
  CnBd.Execute strCadena
  
  strCadena = " insert into imp_tramide_detalle (id_tramite, id_movimiento, observacion , id_autor , fecha_entrada, hora_entrada) values ( " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0) & ",'06','" & Me.Text5.Text & "','" & KEY_USUARIO & "', CURDATE(), CURTIME())"
  CnBd.Execute strCadena
  
  limpiarFrames
  
  Call actualizaListaConFila(Me.gridTramites)
End Sub



Private Sub cmdConfRecojoPlaCli_Click()
  strCadena = " update imp_tramite set id_estado = '02', id_estado_detalle = '10' where id_tramite = " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0)
  CnBd.Execute strCadena
  
  strCadena = " insert into imp_tramide_detalle (id_tramite, id_movimiento, observacion , id_autor , fecha_entrada, hora_entrada) values ( " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0) & ",'10','" & Me.Text9.Text & "','" & KEY_USUARIO & "', CURDATE(), CURTIME())"
  CnBd.Execute strCadena
  
  limpiarFrames
  
  Call actualizaListaConFila(Me.gridTramites)
End Sub


Private Sub cmdExpedienteListo_Click()
  fraExpIns.Visible = True
End Sub

Private Sub cmdIngresoRP_Click()
  fraIngrExpRR.Visible = True
End Sub

Private Sub cmdLegalizar_Click()
  fraLegal.Visible = True
End Sub

Private Sub cmdLimpiarDNI_Click()
  limpiarDNI
End Sub

Private Sub cmdPagoPlacas_Click()
  fraPago.Visible = True
End Sub

Private Sub cmdPlacasListas_Click()
  fraPlaLis.Visible = True
End Sub

Private Sub cmdRecojoPlacasCli_Click()
  fraPlaCli.Visible = True
End Sub

Private Sub cmdRecojoPlacasInt_Click()
  fraRecojoPla.Visible = True
End Sub

Private Sub cmdRecojoRP_Click()
 fraRecojoExpRP.Visible = True
End Sub

Private Sub cmdRegPlacaWeb_Click()
  fraRegPlaca.Visible = True
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
 CenterForm Me
 Me.Top = 150
 Call limpiarBotones
 Call actualizaLista(Me.gridTramites, "00") '00 ningun filtro
End Sub


Private Sub limpiarFrames()
fraLegal.Visible = False

fraIngrExpRR.Visible = False
fraExpIns.Visible = False

fraRegPlaca.Visible = False
fraRecojoExpRP.Visible = False

fraPago.Visible = False
fraPlaLis.Visible = False

fraRecojoPla.Visible = False
fraPlaCli.Visible = False

Me.fraExpediente.Visible = False

End Sub


Public Sub actualizaLista(ByVal Grilla As MSHFlexGrid, tipo As String)

'Dim color As String, edad As Double
Call limpiarGrid(Me.gridDetalles)

strCadena = "  select i.`id_tramite`, i.id_estado_detalle, d.`id_detalle_venta`, p.`dni` " & _
 "  as dniCliente , pr.id_producto , p.dni, p.`nombre_completo` as cliente, " & _
 "  pr.`nombre_prod` as producto, v.`fecha_emision`, v.`hora` , e.descripcion as estado, mo.`descripcion` as estado_detalle , " & _
 "   d.`serie`, d.`nro_motor`, d.`nro_chasis` , i.`nro_placa`, i.`nro_titulo` " & _
 "  from movimiento_venta v , persona p, `movimiento_venta_detalle` d, " & _
 "  producto pr, imp_tramite i , imp_estado_documentos e, imp_tramite_tipo_mov mo where " & _
 "  p.`dni` = v.`id_cliente` and v.`id_venta` = d.`id_venta` and " & _
 "  i.id_venta = d.id_detalle_venta and e.id_estado = i.id_estado  and " & _
 "  d.`id_producto` = pr.`id_producto` and mo.`id_mov` = i.`id_estado_detalle` and pr.`ruc` = '" & KEY_RUC & "'"
 
 
 Select Case tipo
   
   Case "01"
     strCadena = strCadena & " and v.id_cliente like '%" & Me.txtDNI.Text & "%'"
     Me.txtDNI.Locked = True
     cmdLimpiarDNI.Visible = True
     
   Case "02"
     strCadena = strCadena & " and d.serie like '%" & Me.txtSerie.Text & "%'"
     
   Case "03"
     strCadena = strCadena & " and d.nro_motor like '%" & Me.txtMotor.Text & "%'"
       
   Case "04"
     strCadena = strCadena & " and i.nro_titulo like '%" & Me.txtTitulo.Text & "%'"
     
   Case "05"
     strCadena = strCadena & " and i.nro_placa like '%" & Me.txtPlaca.Text & "%'"
     
 End Select
                   
                  
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If


  N = 1
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 0
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 2500
           Grilla.ColWidth(4) = 1100
           Grilla.ColWidth(5) = 2800
           Grilla.ColWidth(6) = 1100
           Grilla.ColWidth(7) = 900
           Grilla.ColWidth(8) = 800
           Grilla.ColWidth(9) = 0
           Grilla.ColWidth(10) = 3000
           Grilla.ColWidth(11) = 800
           Grilla.ColWidth(12) = 800
           Grilla.ColWidth(13) = 800
           Grilla.ColWidth(14) = 800
           
        Next
         cabecera = "" & vbTab & "" & vbTab & "CODIGO" & vbTab & "PRODUCTO" & vbTab & "DNI CLIENTE" & vbTab & "CLIENTE" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "ESTADO" & vbTab & "" & vbTab & "ULT. MOV." & vbTab & "SERIE" & vbTab & "MOTOR" & vbTab & "TITULO" & vbTab & "PLACA"
         Grilla.AddItem cabecera
         For k = 0 To 14
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
        
             estado = Chr(168)
        
             Fila = rst("id_tramite") & vbTab & rst("id_detalle_venta") & vbTab & rst("id_producto") & vbTab & rst("producto") & _
              vbTab & rst("dni") & vbTab & rst("cliente") & vbTab & rst("fecha_emision") & vbTab & rst("hora") & vbTab & rst("estado") & vbTab & rst("id_estado_detalle") & vbTab & rst("estado_detalle") & vbTab & rst("serie") & vbTab & rst("nro_motor") & vbTab & rst("nro_titulo") & vbTab & rst("nro_placa")
             
             Grilla.AddItem Fila
             
             'With Grilla
                 '.row = i + 1 ' se posiciona en la fila
                 '.col = 1 '  .. en la columna
                 ' cambia la fuente para esta celda
                            
                 '.CellFontName = "Wingdings"
                 '.CellFontSize = 14
                 '.CellAlignment = flexAlignCenterCenter
    
              'End With
             
             
             Fila = ""
             rst.MoveNext
        Next i

End Sub



Public Sub actualizaListaConFila(ByVal Grilla As MSHFlexGrid)

'Dim color As String, edad As Double
Dim Ind As Integer
Ind = Grilla.Row

strCadena = "  select i.`id_tramite`, i.id_estado_detalle, d.`id_detalle_venta`, p.`dni` " & _
 "  as dniCliente , pr.id_producto , p.dni, p.`nombre_completo` as cliente, " & _
 "  pr.`nombre_prod` as producto, v.`fecha_emision`, v.`hora` , e.descripcion as estado, mo.`descripcion` as estado_detalle , " & _
 "   d.`serie`, d.`nro_motor`, d.`nro_chasis` , i.`nro_placa`, i.`nro_titulo` " & _
 "  from movimiento_venta v , persona p, `movimiento_venta_detalle` d, " & _
 "  producto pr, imp_tramite i , imp_estado_documentos e, imp_tramite_tipo_mov mo where " & _
 "  p.`dni` = v.`id_cliente` and v.`id_venta` = d.`id_venta` and " & _
 "  i.id_venta = d.id_detalle_venta and e.id_estado = i.id_estado  and " & _
 "  d.`id_producto` = pr.`id_producto` and mo.`id_mov` = i.`id_estado_detalle` and pr.`ruc` = '" & KEY_RUC & "'"
                  
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If


  N = 1
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 0
           Grilla.ColWidth(2) = 900
           Grilla.ColWidth(3) = 1800
           Grilla.ColWidth(4) = 800
           Grilla.ColWidth(5) = 2500
           Grilla.ColWidth(6) = 800
           Grilla.ColWidth(7) = 650
           Grilla.ColWidth(8) = 800
           Grilla.ColWidth(9) = 0
           Grilla.ColWidth(10) = 3000
           Grilla.ColWidth(11) = 800
           Grilla.ColWidth(12) = 800
           Grilla.ColWidth(13) = 800
           Grilla.ColWidth(14) = 800
           
        Next
         cabecera = "" & vbTab & "" & vbTab & "CODIGO" & vbTab & "PRODUCTO" & vbTab & "DNI CLIENTE" & vbTab & "CLIENTE" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "ESTADO" & vbTab & "" & vbTab & "ULT. MOV." & vbTab & "SERIE" & vbTab & "MOTOR" & vbTab & "TITULO" & vbTab & "PLACA"
         Grilla.AddItem cabecera
         For k = 0 To 14
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
        
             estado = Chr(168)
        
             Fila = rst("id_tramite") & vbTab & rst("id_detalle_venta") & vbTab & rst("id_producto") & vbTab & rst("producto") & _
              vbTab & rst("dni") & vbTab & rst("cliente") & vbTab & rst("fecha_emision") & vbTab & rst("hora") & vbTab & rst("estado") & vbTab & rst("id_estado_detalle") & vbTab & rst("estado_detalle") & vbTab & rst("serie") & vbTab & rst("nro_motor") & vbTab & rst("nro_titulo") & vbTab & rst("nro_placa")
             
             Grilla.AddItem Fila
             
             'With Grilla
                 '.row = i + 1 ' se posiciona en la fila
                 '.col = 1 '  .. en la columna
                 ' cambia la fuente para esta celda
                            
                 '.CellFontName = "Wingdings"
                 '.CellFontSize = 14
                 '.CellAlignment = flexAlignCenterCenter
    
              'End With
             
             
             Fila = ""
             rst.MoveNext
        Next i


 'Actualizo data
Grilla.Row = Ind
Call actualizaTramite

End Sub



Public Sub actualizaDetalles(ByVal Grilla As MSHFlexGrid)

'Dim color As String, edad As Double

 strCadena = " select  DATE_FORMAT(t.`hora_entrada`, '%H:%i') as hora_entrada , DATE_FORMAT(t.`fecha_entrada`,'%d/%m/%Y')  as `fecha_entrada`, m.`descripcion` as movimiento, p.`nombre_completo` as autor, t.`observacion` " & _
" from `imp_tramide_detalle` t , `imp_tramite_tipo_mov` m , persona p " & _
"where t.`id_movimiento` = m.`id_mov` and t.`id_autor` = p.`dni` and " & _
"t.`id_tramite` = " & Me.gridTramites.TextMatrix(Me.gridTramites.Row, 0)
                  
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If


  N = 1
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 800
           Grilla.ColWidth(2) = 4000
           Grilla.ColWidth(3) = 3500
           Grilla.ColWidth(4) = 2000
           
           
        Next
         cabecera = "FECHA" & vbTab & "HORA" & vbTab & "DESCRIPCION" & vbTab & "AUTOR" & vbTab & "OBSERVACION"
         Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
        
            
        
             Fila = rst("fecha_entrada") & vbTab & rst("hora_entrada") & vbTab & rst("movimiento") & _
             vbTab & rst("autor") & vbTab & rst("observacion")
             
             Grilla.AddItem Fila
             
             'With Grilla
                 '.row = i + 1 ' se posiciona en la fila
                 '.col = 1 '  .. en la columna
                 ' cambia la fuente para esta celda
                            
                 '.CellFontName = "Wingdings"
                 '.CellFontSize = 14
                 '.CellAlignment = flexAlignCenterCenter
    
              'End With
             
             
             Fila = ""
             rst.MoveNext
        Next i

End Sub


Private Sub gridTramites_DblClick()
   'If gridTramites.row > 0 Then
     'Call actualizaDetalles(Me.gridDetalles)
   'End If
End Sub

Private Sub gridTramites_Click()
   Call actualizaTramite
End Sub

Private Sub actualizaTramite()
   If gridTramites.Row = 0 Then
      Call limpiarBotones
     
     Else
     
     Call actualizaBotones
     Call actualizaDetalles(Me.gridDetalles)
   End If
End Sub


Private Sub actualizaBotones()
   limpiarBotones
   
   Select Case Me.gridTramites.TextMatrix(Me.gridTramites.Row, 9)
     Case "00"
       Call limpiarBotones
       cmdLegalizar.Visible = True
       
     Case "01"
       Call limpiarBotones
       cmdArmado.Visible = True
       
     Case "02"
       Call limpiarBotones
       cmdIngresoRP.Visible = True
       
     Case "03"
       Call limpiarBotones
       cmdExpedienteListo.Visible = True
       
       Case "04"
       Call limpiarBotones
       cmdRecojoRP.Visible = True
       
     Case "05"
       Call limpiarBotones
       cmdRegPlacaWeb.Visible = True
       
       Case "06"
       Call limpiarBotones
       cmdPagoPlacas.Visible = True
       
     Case "07"
       Call limpiarBotones
       cmdPlacasListas.Visible = True
     
     Case "08"
       Call limpiarBotones
       cmdRecojoPlacasInt.Visible = True
       
       Case "09"
       Call limpiarBotones
       cmdRecojoPlacasCli.Visible = True
       
     Case "10"
       Call limpiarBotones
       
   End Select
End Sub

Private Sub limpiarBotones()
    cmdLegalizar.Visible = False
    cmdArmado.Visible = False
    cmdIngresoRP.Visible = False
    cmdExpedienteListo.Visible = False
    cmdRecojoRP.Visible = False
    cmdRegPlacaWeb.Visible = False
    cmdPagoPlacas.Visible = False
    cmdPlacasListas.Visible = False
    cmdRecojoPlacasInt.Visible = False
    cmdRecojoPlacasCli.Visible = False
End Sub


Private Sub Command1_Click()
limpiarFrames
End Sub

Private Sub Command10_Click()
limpiarFrames
End Sub

Private Sub Command11_Click()
limpiarFrames
End Sub

Private Sub Command14_Click()
limpiarFrames
End Sub

Private Sub Command15_Click()
limpiarFrames
End Sub

Private Sub Command3_Click()
limpiarFrames
End Sub

Private Sub Command4_Click()
limpiarFrames
End Sub

Private Sub Command5_Click()
limpiarFrames
End Sub

Private Sub Command8_Click()
 limpiarFrames
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Key
     Case Is = "(Salir)"
      Unload Me
   End Select
End Sub


Private Sub txtDNI_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      
      strCadena = "select * from persona where dni ='" & txtDNI.Text & "'"
      'MsgBox strCadena
      Call ConfiguraRstL(strCadena)
      
      If rstL.RecordCount > 0 And Len(txtDNI.Text) > 0 And txtDNI.Text <> "" Then
         Me.lblNombre.Caption = rstL("nombres") & " " & rstL("a_paterno") & " " & rstL("a_materno")
         Call actualizaLista(Me.gridTramites, "01")
         
         Else
         
         Procedencia = Selecionar
         FrmPersona.Show
         Exit Sub
         'frmBuscarUsuario.Show
         
      End If

  End If
End Sub

Private Sub limpiarGrid(ByVal Grilla As MSHFlexGrid)
     Grilla.Rows = 0
     Grilla.Clear
End Sub

Private Sub limpiarDNI()
     Me.txtDNI.Text = ""
     Me.lblNombre.Caption = "-"
     Me.txtDNI.Locked = False
     cmdLimpiarDNI.Visible = False
     Call actualizaLista(Me.gridTramites, "00")
End Sub


Private Sub txtSerie_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call actualizaLista(Me.gridTramites, "02")
     
  End If
End Sub

Private Sub txtMotor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call actualizaLista(Me.gridTramites, "03")
  End If
End Sub


Private Sub txtTitulo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call actualizaLista(Me.gridTramites, "04")
  End If
End Sub


Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call actualizaLista(Me.gridTramites, "05")
  End If
End Sub


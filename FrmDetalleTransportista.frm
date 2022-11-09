VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmDetalleTransportistas 
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   12720
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2835
      MaxLength       =   8
      TabIndex        =   29
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox TxtEntidad 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2835
      MaxLength       =   50
      TabIndex        =   28
      Top             =   840
      Width           =   6855
   End
   Begin VB.TextBox TxtDireccion1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2835
      MaxLength       =   80
      TabIndex        =   27
      Top             =   1320
      Width           =   6855
   End
   Begin VB.TextBox TxtDireccion2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2835
      MaxLength       =   80
      TabIndex        =   26
      Top             =   1800
      Width           =   6855
   End
   Begin VB.TextBox TxtEmail 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2835
      MaxLength       =   50
      TabIndex        =   25
      Top             =   3360
      Width           =   6975
   End
   Begin VB.TextBox TxtTelefono1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2835
      MaxLength       =   9
      TabIndex        =   24
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox TxtObservacion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   675
      Left            =   2835
      MaxLength       =   100
      TabIndex        =   23
      Top             =   3840
      Width           =   6975
   End
   Begin VB.TextBox TxtDNI 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5715
      MaxLength       =   8
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Frame FrmCondPago 
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
      Begin VB.OptionButton OptJuridica 
         Caption         =   "Jurídica"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton OptNatural 
         Caption         =   "Natural"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtTelefono2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5715
      MaxLength       =   9
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox TxtFax 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   7995
      MaxLength       =   9
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton CmdFoto 
      Caption         =   "Seleccione su Foto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CheckBox ChkPercepcion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Afecto a Percepción"
      Height          =   330
      Left            =   2760
      TabIndex        =   1
      Top             =   4800
      Width           =   1770
   End
   Begin VB.CheckBox ChkRetencion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Afecto a Retencion."
      Height          =   330
      Left            =   5160
      TabIndex        =   0
      Top             =   4800
      Width           =   1770
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   1680
      Left            =   10320
      TabIndex        =   9
      Top             =   3600
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   2963
      BandCount       =   1
      ForeColor       =   -2147483635
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1755
      _CBHeight       =   1680
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   1620
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   1620
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2858
         ButtonWidth     =   1349
         ButtonHeight    =   953
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Grabar"
               Key             =   "(Grabar)"
               Object.ToolTipText     =   "Grabar Ctrl+G"
               ImageKey        =   "(Grabar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "(Cancelar)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Cancelar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LblCodTransportista 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2835
      TabIndex        =   30
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label LblApellido 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Razon Social:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   22
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label LblDireccion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección 1:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   21
      Top             =   1380
      Width           =   885
   End
   Begin VB.Label LblNDocumento 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4950
      TabIndex        =   20
      Top             =   2820
      Width           =   345
   End
   Begin VB.Label LblTipoDocumento 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruc:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   19
      Top             =   2940
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Persona:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   210
      TabIndex        =   18
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   17
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección 2:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   16
      Top             =   1860
      Width           =   885
   End
   Begin VB.Label LblFax 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7560
      TabIndex        =   15
      Top             =   2340
      Width           =   375
   End
   Begin VB.Label LblTelefono2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono 2 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4680
      TabIndex        =   14
      Top             =   2340
      Width           =   885
   End
   Begin VB.Label LblObservacion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones : "
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label LblTelefono1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono 1 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   12
      Top             =   2340
      Width           =   885
   End
   Begin VB.Label LblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E mail :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   11
      Top             =   3420
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2175
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   5115
      Left            =   1515
      Top             =   240
      Width           =   8535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   2460
      Left            =   10200
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "FrmDetalleTransportistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

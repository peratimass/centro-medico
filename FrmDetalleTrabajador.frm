VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmDetalleTrabajador 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4635
   ScaleWidth      =   11565
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1455
      MaxLength       =   8
      TabIndex        =   21
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox TxtDNI 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4395
      MaxLength       =   8
      TabIndex        =   18
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox TxtEntidad 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1455
      MaxLength       =   50
      TabIndex        =   7
      Top             =   720
      Width           =   6855
   End
   Begin VB.TextBox TxtDireccion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1455
      MaxLength       =   80
      TabIndex        =   6
      Top             =   1200
      Width           =   6855
   End
   Begin VB.TextBox TxtTelefono2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4335
      MaxLength       =   9
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox TxtEmail 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1455
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2760
      Width           =   6975
   End
   Begin VB.TextBox TxtTelefono1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1455
      MaxLength       =   9
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox TxtFax 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6720
      MaxLength       =   9
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox TxtObservacion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   675
      Left            =   1455
      MaxLength       =   100
      TabIndex        =   1
      Top             =   3240
      Width           =   6975
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
      Left            =   9000
      TabIndex        =   0
      Top             =   2760
      Width           =   2175
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2400
      Left            =   9000
      TabIndex        =   8
      Top             =   3480
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   4233
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1875
      _CBHeight       =   2400
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   2340
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1376
         ButtonWidth     =   1349
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
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
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   6360
      Top             =   5760
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
            Picture         =   "FrmDetalleTrabajador.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTrabajador.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTrabajador.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTrabajador.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTrabajador.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTrabajador.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTrabajador.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTrabajador.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTrabajador.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTrabajador.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTrabajador.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTrabajador.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label LblCodTrabajador 
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
      Left            =   1455
      TabIndex        =   22
      Top             =   240
      Width           =   1425
   End
   Begin VB.Label LblNDocumento 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3630
      TabIndex        =   20
      Top             =   2220
      Width           =   345
   End
   Begin VB.Label LblTipoDocumento 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruc:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   600
      TabIndex        =   19
      Top             =   2220
      Width           =   375
   End
   Begin VB.Label LblApellido 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   465
      TabIndex        =   17
      Top             =   780
      Width           =   615
   End
   Begin VB.Label LblDireccion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección 1:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   270
      TabIndex        =   16
      Top             =   1260
      Width           =   885
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   270
      TabIndex        =   15
      Top             =   360
      Width           =   615
   End
   Begin VB.Label LblFax 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6240
      TabIndex        =   14
      Top             =   1740
      Width           =   375
   End
   Begin VB.Label LblTelefono2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono 2 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3360
      TabIndex        =   13
      Top             =   1740
      Width           =   885
   End
   Begin VB.Label LblObservacion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones : "
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label LblTelefono1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono 1 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   1740
      Width           =   885
   End
   Begin VB.Label LblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E mail :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   2820
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2175
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   4155
      Left            =   195
      Top             =   120
      Width           =   8535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   2460
      Left            =   8880
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "FrmDetalleTrabajador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

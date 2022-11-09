VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmFichaIncripcion 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   14385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdComprobante 
      Caption         =   "GENERAR COMPROBANTE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12000
      TabIndex        =   44
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Frame Frame5 
      Caption         =   "COMPROBANTES RELACIONADOS"
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
      Height          =   2175
      Left            =   8400
      TabIndex        =   42
      Top             =   480
      Width           =   5895
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1815
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3201
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
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
   End
   Begin VB.TextBox TxtDni 
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1200
      MaxLength       =   80
      TabIndex        =   40
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox TxtDireccion 
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1200
      MaxLength       =   80
      TabIndex        =   39
      Top             =   1680
      Width           =   4815
   End
   Begin VB.TextBox TxtRazon 
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1200
      MaxLength       =   80
      TabIndex        =   38
      Top             =   1320
      Width           =   4815
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   300
      Left            =   1200
      TabIndex        =   34
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   21364737
      CurrentDate     =   41165
   End
   Begin VB.TextBox TxtIgv 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10800
      TabIndex        =   33
      Text            =   "si"
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox TxtCantidad 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1500
      MaxLength       =   80
      TabIndex        =   27
      Text            =   "1"
      Top             =   5595
      Width           =   495
   End
   Begin VB.TextBox TxtPrecio 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   26
      Top             =   5595
      Width           =   975
   End
   Begin VB.TextBox TxtDescripcionProducto 
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2070
      MaxLength       =   80
      TabIndex        =   25
      Top             =   5595
      Width           =   5175
   End
   Begin VB.TextBox TxtCodProducto 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   240
      MaxLength       =   80
      TabIndex        =   24
      Top             =   5595
      Width           =   1215
   End
   Begin VB.CommandButton CmdAgregar 
      Height          =   350
      Left            =   9600
      Picture         =   "FrmFichaIncripcion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5595
      Width           =   495
   End
   Begin VB.CommandButton CmdQuitar 
      Height          =   350
      Left            =   10110
      Picture         =   "FrmFichaIncripcion.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5595
      Width           =   495
   End
   Begin VB.Frame Frame4 
      Caption         =   "TIPO SERVICIO"
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
      Height          =   855
      Left            =   6240
      TabIndex        =   18
      Top             =   4440
      Width           =   8055
      Begin MSDataListLib.DataCombo DtcServicios 
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "SERVICIO :"
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
         Left            =   345
         TabIndex        =   20
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "DATOS PERSONAL"
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
      Height          =   1455
      Left            =   6240
      TabIndex        =   11
      Top             =   2880
      Width           =   8055
      Begin MSDataListLib.DataCombo DtcVendedor 
         Height          =   315
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcCobrador 
         Height          =   315
         Left            =   1680
         TabIndex        =   16
         Top             =   600
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcTecnico 
         Height          =   315
         Left            =   1680
         TabIndex        =   17
         Top             =   960
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "TECNICO :"
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
         Left            =   480
         TabIndex        =   14
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "COBRADOR :"
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
         Left            =   300
         TabIndex        =   13
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "VENDEDOR :"
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
         Left            =   345
         TabIndex        =   12
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "MEDIO DE CONTACTO"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   6015
      Begin VB.OptionButton OptOficina 
         Appearance      =   0  'Flat
         Caption         =   "AFILIACION EN OFICINA"
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
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton Opttelefono 
         Appearance      =   0  'Flat
         Caption         =   "AFILIACION VIA TELEFONICA"
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
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton OptVendedor 
         Appearance      =   0  'Flat
         Caption         =   "AFILIACION POR VENDEDOR"
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
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "TIPO TRABAJO"
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
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   6015
      Begin VB.OptionButton OptAnexo 
         Appearance      =   0  'Flat
         Caption         =   "ANEXO"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton ptTraslado 
         Appearance      =   0  'Flat
         Caption         =   "TRASLADO"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton OptReinstalacionTraslado 
         Appearance      =   0  'Flat
         Caption         =   "REINSTALACION TRASLADO"
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
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton OptInspeccion 
         Appearance      =   0  'Flat
         Caption         =   "INSPECCION"
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
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton OptReinstalacion 
         Appearance      =   0  'Flat
         Caption         =   "REINSTALACION"
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
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton OptInstalacion 
         Appearance      =   0  'Flat
         Caption         =   "INSTALACION"
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
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdProducto 
      Height          =   1815
      Left            =   120
      TabIndex        =   21
      Top             =   6000
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   3201
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
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
      Left            =   10080
      Top             =   6720
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
            Picture         =   "FrmFichaIncripcion.frx":0B14
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFichaIncripcion.frx":0E30
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFichaIncripcion.frx":1290
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFichaIncripcion.frx":16F0
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFichaIncripcion.frx":1A0C
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFichaIncripcion.frx":1E6C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFichaIncripcion.frx":2188
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFichaIncripcion.frx":25E8
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFichaIncripcion.frx":2A48
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFichaIncripcion.frx":3328
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFichaIncripcion.frx":3644
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFichaIncripcion.frx":3960
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   12285
      TabIndex        =   31
      Top             =   7290
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1482
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1875
      _CBHeight       =   840
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   780
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   32
         Top             =   30
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1376
         ButtonWidth     =   1296
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
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   1215
      TabIndex        =   45
      Top             =   150
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUCURSAL :"
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
      Left            =   120
      TabIndex        =   46
      Top             =   240
      Width           =   885
   End
   Begin VB.Image imgfoto 
      Height          =   2055
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "FECHA :"
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
      Left            =   480
      TabIndex        =   41
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "DIRECCION :"
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
      Left            =   120
      TabIndex        =   37
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "NOMBRE :"
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
      Left            =   345
      TabIndex        =   36
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "DNI/RUC :"
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
      Left            =   330
      TabIndex        =   35
      Top             =   960
      Width           =   750
   End
   Begin VB.Label lblNumero 
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   12360
      TabIndex        =   30
      Top             =   120
      Width           =   1635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "NUMERO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   11400
      TabIndex        =   29
      Top             =   120
      Width           =   795
   End
   Begin VB.Label LblTotalParcial 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Height          =   285
      Left            =   8310
      TabIndex        =   28
      Top             =   5595
      Width           =   1200
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00000000&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   120
      Top             =   5520
      Width           =   11775
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   420
      Left            =   120
      Top             =   90
      Width           =   6080
   End
End
Attribute VB_Name = "FrmFichaIncripcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub CmdAgregar_Click()
If Val(Me.TxtCantidad.text) > 0 And Trim(Me.TxtCodProducto.text) <> "" Then
    'strCadena = "SELECT COUNT(DISTINCT igv) INTO ncantidad FROM temporal_ventas WHERE ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "' AND id_alm='" & KEY_ALM & "'"
    
        strCadena = "INSERT INTO temporal_ventas(ruc,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,dni_save) VALUES " & _
        "('" & KEY_RUC & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.txtserie.text) & "','" & Me.TxtNumeroDoc.text & "','" & codigoP & "','" & Val(Me.TxtCantidad.text) & "'," & _
        "'" & Val(Me.TxtPrecio.text) & " ','" & Val(Me.TxtPrecio.text) * Val(Me.TxtCantidad.text) & "','" & Val(Me.TXtPeso.text) & "','" & Trim(Me.TxtIgv.text) & "','" & KEY_USUARIO & "')"
       CnBd.Execute (strCadena)
    
    Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.text, Me.txtserie.text, Me.DtcTipoDoc.BoundText)
    Call VerificaDocumento(Trim(Me.DtcTipoDoc.BoundText))
   
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
    Me.ChkPrecioAlterno.Value = 0
    Me.TxtCodProducto.text = "00000"
    Me.TxtCantidad.text = "0"
    Me.TxtDescripcionProducto.text = ""
    Me.TxtPrecio.text = ""
    Me.LblTotalParcial.Caption = ""
    Me.ChkPrecioAlterno.Enabled = False
    Call Resalta(Me.TxtCodProducto)
   
Else
    Call Resalta(Me.TxtCantidad)
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.DtpFecha.Value = KEY_FECHA
strCadena = "SELECT P.dni as Codigo,P.nombre_completo as Descripcion FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni AND E.id_empresa='" & KEY_RUC & "' AND E.id_personal='si' ORDER BY nombre_completo"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcVendedor)

strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM
Me.DtcAlmacen.Enabled = False
strCadena = "SELECT P.dni as Codigo,P.nombre_completo as Descripcion FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni AND E.id_empresa='" & KEY_RUC & "' AND E.id_personal='si' ORDER BY nombre_completo"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcCobrador)

strCadena = "SELECT P.dni as Codigo,P.nombre_completo as Descripcion FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni AND E.id_empresa='" & KEY_RUC & "' AND E.id_personal='si' ORDER BY nombre_completo"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcTecnico)

strCadena = "SELECT P.dni as Codigo,P.nombre_completo as Descripcion FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni AND E.id_empresa='" & KEY_RUC & "' AND E.id_personal='si' ORDER BY nombre_completo"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcVendedor)

strCadena = "SELECT id_producto as Codigo,nombre_prod as Descripcion FROM producto WHERE ruc='" & KEY_RUC & "' AND id_tipo='02' ORDER BY nombre_prod"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcServicios)



End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_CANCEL
         Unload Me
End Select
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
Dim TotalP As Single
If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 And KeyAscii <> 13 Then
        KeyAscii = 0
End If

If KeyAscii = 13 Then
    If KEY_AUTOMATICO = "si" Then
        Call Agregar_directo
    Else
    strCadena = "SELECT     Almacen_Productos.Stock, Producto.PrecioVenta " & _
    "FROM  Almacen_Productos INNER JOIN Producto ON Almacen_Productos.cProducto = Producto.cProducto WHERE (Almacen_Productos.cProducto='" & Trim(codigoP) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "')"
    Call ConfiguraRst(strCadena)
    Call Resalta(Me.TxtPrecio)
    If Val(Me.TxtPrecio.text) = 0 Then
        MsgBox "Este Producto no Cuenta con un Precio de Venta", vbExclamation
        Call Resalta(Me.TxtCodProducto)
       Exit Sub
    End If
    TotalP = Val(Me.TxtCantidad.text) * Val(Me.TxtPrecio.text)
    Me.LblTotalParcial.Caption = Format(TotalP, "#,##0.00")
    
    Me.ChkPrecioAlterno.Enabled = True
    If KEY_AUTOMATICO = "si" Then
        Call CmdAgregar_Click
    End If
    End If
    Set rst = Nothing
End If

End Sub

Private Sub TxtCodProducto_KeyPress(KeyAscii As Integer)
Dim Criterio As String
If KeyAscii = 13 Then
    
    If (Len(Me.TxtCodProducto.text) = 0) Or Val(Me.TxtCodProducto.text) = 0 Then
        
        Call Resalta(Me.TxtCodProducto)
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
    End If
    
     
    If KEY_BARRAS = "si" Then
        strCadena = "SELECT B.id_producto,P.nombre_prod,P.precio_venta,P.peso,P.id_igv FROM producto_barras B,producto P WHERE B.id_producto=P.id_producto AND B.ruc='" & KEY_RUC & "' " & _
        "AND P.ruc='" & KEY_RUC & "' AND B.cod_barra='" & Trim(Me.TxtCodProducto.text) & "'"
    Else
        Me.TxtCodProducto.text = FormatosCeros(Me.TxtCodProducto.text, 5)
        strCadena = "SELECT A.id_producto, P.nombre_prod,P.precio_venta,P.peso,P.id_igv FROM almacen_producto A,producto P WHERE A.id_producto=P.id_producto AND A.id_alm='" & KEY_ALM & "' AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND A.id_producto='" & Trim(Me.TxtCodProducto.text) & "'"
    End If
        
    
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        codigoP = rst("id_producto")
        Me.TxtDescripcionProducto.text = rst("nombre_prod")
        Me.TxtIgv.text = rst("id_igv")
        Me.TxtPrecio.text = rst("precio_venta")
        If Trim(Me.TxtCantidad.text) > 0 Then
            Me.TxtCantidad.text = Me.TxtCantidad.text
         Else
          Me.TxtCantidad.text = 1
        End If
        
        Call Resalta(Me.TxtCantidad)
        
        
        Set rst = Nothing
        
    Else
        
        Call Resalta(Me.TxtCodProducto)
        Procedencia = Selecionar
        FrmProducto.Show
    End If
End If

End Sub

Private Sub TxtDNI_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 If Trim(Me.TxtDni.text) = "" Then
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
 End If
  
If Len(Trim(Me.TxtDni.text)) = 8 Or Len(Trim(Me.TxtDni.text)) = 11 Then
    strCadena = "SELECT dni,nombre_completo,direccion,foto  FROM persona WHERE dni='" & Trim(Me.TxtDni.text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.txtRuc.text = Trim(Me.TxtDni.text)
        FrmDetallePersona.chkCliente.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.imgfoto.Visible = True
        
        Me.TxtRazon.text = UCase(rst("nombre_completo"))
        Me.TxtDireccion.text = UCase(rst("direccion"))
        If IsNull(rst("foto")) = False And Len(rst("foto")) > 5 Then
            If VerificarFichero(App.Path & "\archivos\" & rst("dni")) = True Then
                Me.imgfoto.Picture = LoadPicture(App.Path + "\archivos\" + rst("dni") + "\" + Trim(rst("foto")))
            Else
                Me.imgfoto.Picture = LoadPicture(App.Path + "\archivos\no_photo.jpg")
            End If
        Else
            Me.imgfoto.Picture = LoadPicture(App.Path + "\archivos\no_photo.jpg")
        End If
    End If
End If

     
      
      
      
      strCadena = "SELECT * FROM movimiento_venta V,movimiento_venta_cuotas C WHERE V.id_venta=C.id_venta AND V.ruc='" & KEY_RUC & "' AND C.ruc='" & KEY_RUC & "' AND C.saldo>0 AND V.id_cliente='" & Trim(Me.TxtDni.text) & "' AND V.anulado='no' AND C.vencimiento<='" & KEY_FECHA & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            
            'Me.CmdVisualizar.Visible = True
            'Me.CmdVisualizar.Caption = "TIENE" + Space(2) + Str(rst.RecordCount) + Space(2) + "PAGO PENDIENTE"
        Else
            
            strCadena = "SELECT sum(C.saldo) FROM movimiento_venta V,movimiento_venta_cuotas C WHERE V.id_venta=C.id_venta AND V.ruc='" & KEY_RUC & "' AND C.ruc='" & KEY_RUC & "' AND C.saldo>0 AND V.id_cliente='" & Trim(Me.TxtDni.text) & "' AND V.anulado='no'"
            Call ConfiguraRst(strCadena)
            If IsNull(rst(0)) = False Then
                'Me.CmdVisualizar.Visible = True
                'Me.CmdVisualizar.Caption = "TIENE" + Space(2) + Str(rst(0)) + Space(2) + "DE CREDITO"
            Else
            'Me.CmdVisualizar.Visible = False
            End If
            
            
        End If
      
End If
End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmAdelantoPersonal 
   BorderStyle     =   0  'None
   Caption         =   "Frm Salida de Dinero"
   ClientHeight    =   6780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmitf 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ITF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3765
      TabIndex        =   28
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox TxtItf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   600
         MaxLength       =   80
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OBSERVACIONES"
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
      Height          =   1215
      Left            =   6480
      TabIndex        =   26
      Top             =   4560
      Width           =   6135
      Begin VB.TextBox TxtObservacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   795
         Left            =   120
         MaxLength       =   80
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.TextBox TxtTc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   4365
      MaxLength       =   80
      TabIndex        =   25
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CUENTA DESTINO PROVEEDOR"
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
      Height          =   2775
      Left            =   6525
      TabIndex        =   14
      Top             =   1680
      Width           =   6135
      Begin VB.TextBox TxtOperacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   17
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   4200
         TabIndex        =   16
         Top             =   2085
         Width           =   375
      End
      Begin VB.TextBox TxtCuentaBancaria 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   15
         Top             =   2085
         Width           =   2415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshCuentasBancarias 
         Height          =   1095
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1931
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
      Begin MSDataListLib.DataCombo DtcBanco 
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         Top             =   1420
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo DtcMonedaCuenta 
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         Top             =   1760
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONEDA :"
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
         Left            =   825
         TabIndex        =   24
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BANCO   :"
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
         Left            =   855
         TabIndex        =   23
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N.OPERACION :"
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
         Left            =   405
         TabIndex        =   22
         Top             =   2460
         Width           =   1185
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUENTA BANCARIA:"
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
         Left            =   75
         TabIndex        =   21
         Top             =   2100
         Width           =   1515
      End
   End
   Begin VB.Frame FrmCheque 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PAGAR CON CHEQUE"
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
      Height          =   1215
      Left            =   1365
      TabIndex        =   9
      Top             =   4560
      Width           =   4695
      Begin VB.OptionButton OptChequeSi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SI"
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
         Left            =   840
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton OptChequeNO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NO"
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
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton cmdCargarCheque 
         Caption         =   "CARGAR CHEQUE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   3015
      End
      Begin MSDataListLib.DataCombo DtcCheque 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
   End
   Begin VB.TextBox TxtMontoPago 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1485
      MaxLength       =   80
      TabIndex        =   8
      Top             =   3060
      Width           =   1935
   End
   Begin VB.TextBox TxtNumeroDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   10365
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   7
      Text            =   "000000"
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox TxtSerie 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   8325
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   6
      Text            =   "000"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   5
      Top             =   1140
      Width           =   1575
   End
   Begin VB.TextBox TxtCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1485
      MaxLength       =   80
      TabIndex        =   4
      Top             =   1500
      Width           =   4695
   End
   Begin VB.TextBox TxtDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1485
      MaxLength       =   80
      TabIndex        =   3
      Top             =   1860
      Width           =   4695
   End
   Begin VB.TextBox TxtCcostos 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Left            =   1485
      MaxLength       =   80
      TabIndex        =   1
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox TxtCostosdet 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
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
      Left            =   2280
      MaxLength       =   80
      TabIndex        =   0
      Top             =   3480
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker DtpEmision 
      Height          =   300
      Left            =   6645
      TabIndex        =   30
      Top             =   1125
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   121176065
      CurrentDate     =   41130
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   120
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSalidaDinero.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSalidaDinero.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSalidaDinero.frx":077C
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSalidaDinero.frx":0809
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSalidaDinero.frx":0C69
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSalidaDinero.frx":0F85
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSalidaDinero.frx":13E5
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSalidaDinero.frx":1701
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSalidaDinero.frx":1B61
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSalidaDinero.frx":1FC1
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSalidaDinero.frx":28A1
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSalidaDinero.frx":2BBD
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmSalidaDinero.frx":2ED9
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   9585
      TabIndex        =   31
      Top             =   5820
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   2955
      _CBHeight       =   870
      _Version        =   "6.0.8169"
      Child1          =   "TlbGrabar"
      MinHeight1      =   810
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbGrabar 
         Height          =   810
         Left            =   30
         TabIndex        =   32
         Top             =   30
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   1429
         ButtonWidth     =   1429
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
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
               Caption         =   "Imprimir"
               Key             =   "(Imprimir)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Imprimir)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   &Salir  "
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   165
      TabIndex        =   33
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
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
   Begin MSDataListLib.DataCombo DtcCuentas 
      Height          =   315
      Left            =   1485
      TabIndex        =   34
      Top             =   3960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
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
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   315
      Left            =   8325
      TabIndex        =   35
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      ForeColor       =   8388608
      ListField       =   "0000º"
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
   Begin MSDataListLib.DataCombo DtcMoneda 
      Height          =   315
      Left            =   1485
      TabIndex        =   36
      Top             =   2295
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
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
   Begin MSComCtl2.DTPicker DtpValor 
      Height          =   300
      Left            =   8790
      TabIndex        =   37
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   121176065
      CurrentDate     =   41130
   End
   Begin VB.TextBox Txtid_solicitud 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label lblAnulado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   630
      Left            =   645
      TabIndex        =   41
      Top             =   360
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR:"
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
      Left            =   8160
      TabIndex        =   48
      Top             =   1155
      Width           =   585
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMISION:"
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
      Left            =   5940
      TabIndex        =   47
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T.C:"
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
      Left            =   3810
      TabIndex        =   46
      Top             =   2340
      Width           =   345
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESEMBOLSO :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   255
      TabIndex        =   45
      Top             =   3120
      Width           =   1155
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   465
      TabIndex        =   44
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAZON SOCIAL:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   225
      TabIndex        =   43
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUENTA ORIGEN:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   75
      TabIndex        =   42
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/DNI :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   615
      TabIndex        =   40
      Top             =   1140
      Width           =   795
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONEDA :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   585
      TabIndex        =   39
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.COSTOS :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   555
      TabIndex        =   38
      Top             =   3540
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   6780
      Left            =   0
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "FrmAdelantoPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mostrar As Boolean
Dim codigo_P As String
Public Procedencia As EnumProcede
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Dim tTotal As Double
tTotal = 0

strCadena = "SELECT      movimiento_caja.doc_cod,movimiento_caja.fecha_valor, Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, " & _
"movimiento_caja.descripcion_per , movimiento_caja.Monto,movimiento_caja.anulado " & _
"FROM         movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod " & _
" WHERE  movimiento_caja.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "'  ORDER BY  movimiento_caja.numero DESC"
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
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 600
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 1300
           Grilla.ColWidth(3) = 600
           Grilla.ColWidth(4) = 800
           Grilla.ColWidth(5) = 3000
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 0
           Grilla.ColWidth(8) = 0
          Next
         cabecera = "ITEM" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "SERIE" & vbTab & "NUMERO" & vbTab & "PERSONA" & vbTab & "MONTO" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 8
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = Fila & str(rst.RecordCount - i + 1) & vbTab & rst("fecha_valor") & vbTab & rst("doc_abrev") & vbTab & rst("serie") & vbTab & rst("numero") & vbTab & rst("descripcion_per") & vbTab & Format(rst("monto"), "#,##0.00")
            If (Fila = "") Then
                x = 1
            End If
            
          Grilla.AddItem Fila
            If (Trim(rst("anulado")) = "si") Then
                            For k = 0 To 8
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &H8080FF
                            Next k
        End If
        Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub ChkAdelantado_Click()

End Sub



Private Sub cmdCargarCheque_Click()
Dim glosa As String
Procedencia = 1
If Val(Me.TxtMontoPago.Text) > 0 Then
        strCadena = "DELETE FROM cheque_detalle WHERE id_cheque='" & Val(Me.DtcCheque.BoundText) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        strCadena = "DELETE FROM cheque_factura WHERE id_cheque='" & Val(Me.DtcCheque.BoundText) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        
        strCadena = "INSERT INTO cheque_detalle(id_cheque,detalle,monto,ruc)VALUES('" & Val(Me.DtcCheque.BoundText) & "','" & Trim(Me.TxtObservacion.Text) & "','" & Val(Me.TxtMontoPago.Text) & "','" & KEY_RUC & "')"
        Call CnBd.Execute(strCadena)
         
        
        strCadena = "SELECT * FROM solicitud_dinero WHERE dni='" & Me.TxtRuc.Text & "' AND ruc='" & KEY_RUC & "' AND saldo>0 AND anulado='no'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            For i = 0 To rst.RecordCount - 1
                strCadena = "INSERT INTO cheque_factura(id_cheque,id_compra,ruc)VALUES('" & Me.DtcCheque.BoundText & "','" & rst("id_solicitud") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                 
                 
                rst.MoveNext
            Next i
        End If
        FrmChequeNuevo.Txtcentrocosto.Text = "42121"
        FrmChequeNuevo.lblcostos.Text = "FACTURAS POR PAGAR"
        FrmChequeNuevo.txtMotivo.Text = ""
        FrmChequeNuevo.TxtMontoMotivo.Text = ""
        FrmChequeNuevo.TxtRuc.Text = Me.TxtRuc.Text
        FrmChequeNuevo.txtrazonsocial.Text = Me.txtcliente.Text
        FrmChequeNuevo.txtdireccion.Text = Me.txtdireccion.Text
        Call Resalta(FrmChequeNuevo.txtMotivo)
        
    End If

'FrmChequeNuevo.Show
'FrmChequeNuevo.TxtidCheque.text = Me.DtcCheque.BoundText
'Call FrmChequeNuevo.llenar_cheque(Me.DtcCheque.BoundText)

End Sub

Private Sub Command1_Click()
If Len(Me.TxtCuentaBancaria.Text) > 5 Then
    strCadena = "SELECT * FROM persona_cuentabancaria WHERE dni='" & Trim(Me.TxtRuc.Text) & "' AND cuenta='" & Trim(Me.TxtCuentaBancaria.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        strCadena = "INSERT INTO persona_cuentabancaria(dni,id_banco,id_moneda,cuenta)VALUES('" & Trim(Me.TxtRuc.Text) & "','" & Me.DtcBanco.BoundText & "','" & Me.DtcMonedaCuenta.BoundText & "','" & Trim(Me.TxtCuentaBancaria.Text) & "') "
        CnBd.Execute (strCadena)
         
         
        Call llenar_cuentas(Me.MshCuentasBancarias, Trim(Me.TxtRuc.Text))
    Else
    MsgBox "CUENTA YA REGISTRADA", vbInformation, KEY_EMPRESA
    End If
End If
End Sub



Private Sub DtcCuentas_Change()
Dim ssaldo As Double, residuo As Single, sitf As Single

strCadena = "SELECT * FROM mis_cuentas WHERE id_cuenta='" & Me.DtcCuentas.BoundText & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst("id_moneda") <> Me.DtcMoneda.BoundText Then
    If rst("id_moneda") = "00001" Then
        
        Me.TxtMontoPago.Text = Format(ssaldo * KEY_CAMBIO, "###0.00")
        Me.DtcMoneda.BoundText = rst("id_moneda")
    Else
        
        Me.TxtMontoPago.Text = Format(ssaldo / KEY_CAMBIO, "###0.00")
        Me.DtcMoneda.BoundText = rst("id_moneda")
    End If
End If
residuo = Val(Me.TxtMontoPago.Text) Mod 1000

If (Val(Me.TxtMontoPago.Text) - residuo) > 0 Then
    sitf = (Val(Me.TxtMontoPago.Text) - residuo) * 0.005 / 100
Else
    itf = 0#
End If
If rst("id_tipo") = "01" Then
    FrmCheque.Enabled = False
    Me.frmitf.Visible = False
Else
    FrmCheque.Enabled = True
    Me.frmitf.Visible = True
    Me.TxtItf.Text = Format(sitf, "#,##0.00")
End If
End Sub

Private Sub DtcMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtMontoPago)
End If
End Sub

Private Sub DtcTipoDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Me.DtcTipoDoc.BoundText) <> "0001" And Trim(Me.DtcTipoDoc.BoundText) <> "0003" Then
        Call Resalta(Me.TxtSerie)
    Else
        MsgBox "Srta:" + Space(1) + KEY_VENDEDOR + Space(1) + "Solo Salida de Dinero", vbInformation
        Me.DtcTipoDoc.SetFocus
    End If
    
    
End If
End Sub

Public Sub nuevo()
    
End Sub

Private Sub ActualizarAdelanto(ByVal TotalPedido As Double)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Dim id_solicitud As Double

Me.cmdCargarCheque.Visible = False
Me.DtpEmision.Value = KEY_FECHA
Me.DtpValor.Value = KEY_FECHA
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
      
  strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0097' AND ruc='" & KEY_RUC & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount < 1 Then
        strCadena = "INSERT INTO almacen_comprobante(ruc,id_alm,id_doc,serie,numero)VALUES ('" & KEY_RUC & "','" & Me.DtcAlmacen.BoundText & "','0097','001','000001')"
        CnBd.Execute (strCadena)
         
         
        strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0097' AND ruc='" & KEY_RUC & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "'"
        Call ConfiguraRst(strCadena)
        Me.TxtSerie.Text = rst("serie")
        Me.TxtNumeroDoc.Text = rst("numero")
  Else
      Me.TxtSerie.Text = rst("serie")
      Me.TxtNumeroDoc.Text = rst("numero")
  End If
  
  strCadena = "SELECT A.id_doc as Codigo,C.doc_abrev as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' ORDER BY C.doc_abrev  "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
  Me.DtcTipoDoc.BoundText = "0097"

  strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM moneda  ORDER BY id_moneda ASC"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtcMoneda)
  Me.DtcTipoDoc.Enabled = False
 
  'Call llenar_cuentas(Me.MshCuentasBancarias, rst("dni"))
  'Call llenar_facturas(Me.MshFacturas, Trim(Me.TxtRuc.text))
  strCadena = "SELECT id_cuenta as Codigo,CONCAT(C.descripcion,'-',M.descripcion,'  ',C.numero_cuenta) as Descripcion FROM mis_cuentas C,moneda M WHERE C.id_moneda=M.id_moneda AND ruc='" & KEY_RUC & "' "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCuentas)
  Me.TxtTc.Text = KEY_CAMBIO
  strCadena = "SELECT id_banco as Codigo,abreviatura as Descripcion FROM banco ORDER BY abreviatura"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtcBanco)
  strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMonedaCuenta)
   
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
    'Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
    'Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
End Sub
Public Sub llenar_cuentas(ByVal Grilla As MSHFlexGrid, ByVal dni As String)
On Error GoTo salir
strCadena = "SELECT B.id_banco,B.abreviatura,M.descripcion,PB.cuenta,M.id_moneda FROM persona_cuentabancaria PB,banco B,moneda M WHERE PB.id_banco=B.id_banco AND PB.id_moneda=M.id_moneda AND PB.dni='" & dni & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1200
            Grilla.ColWidth(2) = 1000
            Grilla.ColWidth(3) = 2000
            Grilla.ColWidth(4) = 0
        Next
        cabecera = "BCO" & vbTab & "MONEDA" & vbTab & "CUENTA"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_banco") & vbTab & rst("abreviatura") & vbTab & rst("descripcion") & vbTab & rst("cuenta") & vbTab & rst("id_moneda")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
        Next i
      
     
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Private Sub LlenarDatosCliente(ByVal Numero As String, ByVal Documento As String, ByVal serie As String, ByVal Almacen As String)

End Sub
Private Sub Save()
Dim monto_pago As Double, saldof As Double, comprobante As String, monto_pagado As Double, Saldo As Double, id_moneda As String, Documento As String
monto_pago = Val(Me.TxtMontoPago.Text)
saldof = 0
        If Me.OptChequeSi.Value = False Then
           
                Documento = Me.DtcTipoDoc.Text & ":" & Trim(Me.TxtSerie.Text) & "-" & Trim(Me.TxtNumeroDoc.Text)
                '------ VERIFICAR MONEDA
                id_moneda = Trim(BDBuscarCampo("mis_cuentas", "id_moneda", "id_cuenta", Me.DtcCuentas.BoundText))
                If Me.DtcMoneda.BoundText <> id_moneda Then
                        If rst("id_moneda") = "00001" Then
                               Saldo = rst("saldo") / KEY_CAMBIO
                        Else
                               Saldo = rst("saldo") * KEY_CAMBIO
                        End If
                  Else
                    monto_pagado = monto_pago
                End If
                '-------END
                
                
               
                    
                    
                    
                    strCadena = "INSERT INTO mis_cuentas_det(id_doc,serie,numero,documento,id_cuenta,fecha,fecha_sys,id_persona,glosa,monto,montoreal,tc,monto_letras,operacion,id_movimiento,dni_save,anulado,ccostos,ruc) " & _
                    " VALUES('" & Me.DtcTipoDoc.BoundText & "','" & Me.TxtSerie.Text & "','" & Me.TxtNumeroDoc.Text & "','" & Documento & "','" & Me.DtcCuentas.BoundText & "','" & Format(Me.DtpValor.Value, "YYYY-mm-dd") & "'," & _
                    "'" & KEY_FECHA & "','" & Me.TxtRuc.Text & "','" & Trim(Me.TxtObservacion.Text) & "','" & monto_pagado & "','" & monto_pagado * -1 & "'," & _
                    "'" & Val(Me.TxtTc.Text) & "','----','" & Me.txtOperacion.Text & "','0','" & KEY_USUARIO & "','no','" & Trim(Me.TxtCcostos.Text) & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                     
                     
                   
            End If
            
    
        
        
        
        nuevo_numero = formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6)
        strCadena = "UPDATE  almacen_comprobante SET numero='" & Trim(nuevo_numero) & "'  WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
        Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
        

        Exit Sub
    

End Sub




Private Sub MshCuentasBancarias_SelChange()
If Trim(Me.MshCuentasBancarias.TextMatrix(Me.MshCuentasBancarias.Row, 2) <> "") Then
    Me.DtcBanco.BoundText = Me.MshCuentasBancarias.TextMatrix(Me.MshCuentasBancarias.Row, 0)
    Me.DtcMonedaCuenta.BoundText = Me.MshCuentasBancarias.TextMatrix(Me.MshCuentasBancarias.Row, 4)
    Me.TxtCuentaBancaria.Text = Me.MshCuentasBancarias.TextMatrix(Me.MshCuentasBancarias.Row, 3)
    Call Resalta(Me.txtOperacion)
Else
    Me.TxtCuentaBancaria.Text = ""
End If
End Sub

Private Sub OptChequeNO_Click()
Me.DtcCheque.Visible = False
Me.cmdCargarCheque.Visible = False
End Sub

Private Sub OptChequeSi_Click()
Me.DtcCheque.Visible = True
strCadena = "SELECT C.id_cheque as Codigo,CONCAT('CHEQUE ',':',C.numero) as Descripcion FROM cheque C,mis_cuentas M,chequera CH  WHERE C.id_chequera=CH.id_chequera AND CH.id_cuenta=M.id_cuenta AND C.seleccionado='si' AND C.ruc='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' AND CH.ruc='" & KEY_RUC & "' AND M.id_cuenta='" & Me.DtcCuentas.BoundText & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
     Call LlenaDataCombo(Me.DtcCheque)
     Me.DtcCheque.SetFocus
     Me.cmdCargarCheque.Visible = True
Else
    Me.cmdCargarCheque.Visible = False
End If
End Sub



Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
      
    Case KEY_PRINT
    
        'strCadena = "SELECT * FROM mis_cuentas_det WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.text) & "' AND numero='" & Trim(Me.TxtNumeroDoc.text) & "' AND ruc='" & KEY_RUC & "'"
        'Call ConfiguraRst(strCadena)
            
    
        'strCadena = "SELECT     Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, movimiento_caja.cPersona, " & _
        "movimiento_caja.descripcion_per, Persona.sDireccionCliente1, Persona.Per_Ruc, movimiento_caja.fecha_valor," & _
        "movimiento_caja.cambio , movimiento_caja.glosa, centro_costos.descripcion, movimiento_caja.Monto,movimiento_caja.monto_letras " & _
        "FROM movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
        "centro_costos ON movimiento_caja.id_costo = centro_costos.id_costo INNER JOIN " & _
        "Persona ON movimiento_caja.cPersona = Persona.cPersona WHERE movimiento_caja.serie='" & Trim(Me.TxtSerie.text) & "' AND movimiento_caja.numero='" & Trim(Me.TxtNumeroDoc.text) & "' AND movimiento_caja.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND Ruc='" & KEY_RUC & "'"
        'Call ConfiguraRst(strCadena)
        'Ans = ShowMultiReport(rst, "RptReciboIngreso", , App.Path + "\Reportes\")
        'Set rst = Nothing
        
        'strCadena = "SELECT     Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, movimiento_caja.cPersona, " & _
        '"movimiento_caja.descripcion_per, Persona.sDireccionCliente1, Persona.Per_Ruc, movimiento_caja.fecha_valor," & _
        '"movimiento_caja.cambio , movimiento_caja.glosa, centro_costos.descripcion, movimiento_caja.Monto,movimiento_caja.monto_letras " & _
        '"FROM movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
        '"centro_costos ON movimiento_caja.id_costo = centro_costos.id_costo INNER JOIN " & _
        '"Persona ON movimiento_caja.cPersona = Persona.cPersona WHERE movimiento_caja.serie='" & Trim(Me.TxtSerie.text) & "' AND movimiento_caja.numero='" & Trim(Me.TxtNumeroDoc.text) & "'"
        
        'Call ConfiguraRst(strCadena)
        'Ans = ShowMultiReport(rst, "RptReciboIngreso", , App.Path + "\Reportes\")
        Call impresion_formato_1_rboingreso_tiket(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.TxtSerie.Text), Trim(Me.TxtNumeroDoc.Text))
Exit Sub
Error:
  
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  
  Exit Sub
  Case KEY_EXIT
    Unload Me
End Select
End Sub

Private Sub Imprimir(ByVal TipoDoc As String, ByVal CodAlm As String, ByVal serie As String, ByVal Numero As String)
Dim i As Integer, j As Integer
Dim totalletras As String

    Printer.ScaleMode = vbCharacters 'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = 10
       
If Me.DtcTipoDoc.BoundText = KEY_SALDINER Then
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.PaperSize = 1
    Printer.Print ""
    Printer.Print ""
    'Printer.Print Tab(15); (CVDate(Me.DtpActual.Value))
    Printer.Print ""
    Printer.Print ""
   'Printer.Print Tab(5); Mid(Me.TxtCliente.text + Space(80), 1, 65)
   ' Printer.Print Tab(5); Mid(Me.TxtDireccion.text + Space(80), 1, 65)
   ' Printer.Print Tab(5); Mid(Me.TxtRuc.Text + Space(50), 1, 40) & "SALDINER"; Space(1); Mid(Me.TxtSerie.Text + Space(50), 1, 4) & Space(1) & "-" & Me.TxtNumeroDoc.Text
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); Mid(Trim(KEY_VENDEDOR), 1, 10)
    Printer.Print ""
    Printer.Print ""
   ' Printer.Print Tab(15); "Monto Efectivo:" & "=============" & Space(20) & Me.TxtMontoIngresar.text
    Printer.CurrentY = Printer.CurrentY + 10
    'totalletras = UCase(EnLetras(Me.TxtMontoIngresar.text))
    Set rst = Nothing
    '---- fin totales
    'Printer.Print Tab(10); Mid(totalletras + Space(100), 1, 60)
    Printer.CurrentY = Printer.CurrentY + 0.2
    'Printer.Print Tab(60); Me.TxtMontoIngresar.text
    Printer.EndDoc
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
    Exit Sub
End If
End Sub


Private Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub




Private Sub TxtCcostos_Change()
strCadena = "SELECT plan_contable_det.plan_des FROM plan_contable_det  WHERE id_plancontable='0001' AND  pc_codigo LIKE '" & Trim(Me.TxtCcostos.Text) & "%' ORDER BY pc_codigo ASC LIMIT 1 "
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    
    Me.TxtCostosdet.Text = rstT("plan_des")
Else
    Me.TxtCostosdet.Text = ""
End If
End Sub

Private Sub TxtCcostos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Me.TxtCcostos.Text) = "" Then
        Procedencia = Selecionar
        FrmPlanContableCuentas.Show
        Exit Sub
    Else
        Me.DtcCuentas.SetFocus
    End If
End If
End Sub

Private Sub TxtMontoPago_Change()
Dim sitf As Single, residuo As Single
residuo = Val(Me.TxtMontoPago.Text) Mod 1000

If (Val(Me.TxtMontoPago.Text) - residuo) > 0 Then
    sitf = (Val(Me.TxtMontoPago.Text) - residuo) * 0.005 / 100
Else
    sitf = 0#
End If
Me.TxtItf.Text = Format(sitf, "#,##0.00")
End Sub

Private Sub TxtMontoPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtCcostos)
    
End If
End Sub

Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Call Resalta(Me.TxtMontoIngresar)
End If
End Sub



Private Sub txtruc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Me.TxtRuc.Text) <> "" Then
        strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.TxtRuc.Text) & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            Me.txtcliente.Text = UCase(rst("nombre_completo"))
            Me.txtdireccion.Text = UCase(rst("direccion"))
            Call llenar_cuentas(Me.MshCuentasBancarias, Trim(Me.TxtRuc.Text))
            Me.DtcMoneda.SetFocus
        Else
            Procedencia = Selecionar
            FrmPersona.Show
            Exit Sub
        End If
    Else
        Procedencia = Selecionar
        FrmPersona.Show
        Exit Sub
    End If
End If
End Sub

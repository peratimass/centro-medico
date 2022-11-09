VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmPedido 
   BorderStyle     =   0  'None
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   15975
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtcosto 
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
      Left            =   10200
      MaxLength       =   80
      TabIndex        =   31
      Top             =   7275
      Width           =   1095
   End
   Begin VitekeySoft.ChameleonBtn cmdprocesar 
      Height          =   855
      Left            =   240
      TabIndex        =   28
      Top             =   7800
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      btype           =   5
      tx              =   "PROCESAR"
      enab            =   -1  'True
      font            =   "FrmPedido1.frx":0000
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "FrmPedido1.frx":0028
      picn            =   "FrmPedido1.frx":0046
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.TextBox TxtUsuario 
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
      Left            =   14160
      MaxLength       =   80
      TabIndex        =   27
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtObservacion 
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
      Height          =   525
      Left            =   7920
      MaxLength       =   80
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   7800
      Width           =   4095
   End
   Begin VB.TextBox TxtCodigo 
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
      Left            =   14040
      MaxLength       =   80
      TabIndex        =   24
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtUnidad 
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
      Left            =   8835
      MaxLength       =   80
      TabIndex        =   23
      Top             =   7275
      Width           =   1215
   End
   Begin VB.TextBox TxtBusquedaRapido 
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
      Left            =   8565
      TabIndex        =   21
      Top             =   1080
      Width           =   2415
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
      Left            =   2565
      MaxLength       =   80
      TabIndex        =   6
      Top             =   7275
      Width           =   6015
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
      Left            =   1635
      MaxLength       =   80
      TabIndex        =   5
      Top             =   7275
      Width           =   855
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "ADD"
      Height          =   315
      Left            =   11415
      TabIndex        =   4
      Top             =   7260
      Width           =   615
   End
   Begin VB.CommandButton CmdQuitar 
      Caption         =   "DELL"
      Height          =   315
      Left            =   12135
      TabIndex        =   3
      Top             =   7260
      Width           =   720
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
      Left            =   8640
      MaxLength       =   80
      TabIndex        =   2
      Top             =   240
      Width           =   735
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
      Left            =   9480
      MaxLength       =   80
      TabIndex        =   1
      Top             =   240
      Width           =   1095
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
      Left            =   375
      MaxLength       =   80
      TabIndex        =   0
      Top             =   7275
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   345
      TabIndex        =   7
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
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
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   11775
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido1.frx":0922
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido1.frx":0D76
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido1.frx":1096
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido1.frx":14EA
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido1.frx":193E
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido1.frx":1C5E
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido1.frx":1F7E
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido1.frx":229E
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido1.frx":25BE
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones1 
      Height          =   6105
      Left            =   14655
      TabIndex        =   8
      Top             =   1560
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   10769
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   6105
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   5670
         Left            =   30
         TabIndex        =   9
         Top             =   375
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   10001
         ButtonWidth     =   1588
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Nuevo   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Anular"
               Key             =   "(Anular)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reporte"
               Key             =   "(Imprimir)"
               ImageKey        =   "(Buscar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13080
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido1.frx":28DE
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedido1.frx":31B8
            Key             =   "(Imprimir)"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DtpActual 
      Height          =   375
      Left            =   11040
      TabIndex        =   10
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
      CalendarBackColor=   0
      CalendarForeColor=   255
      Format          =   165216257
      CurrentDate     =   39535
   End
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   315
      Left            =   6000
      TabIndex        =   11
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSDataListLib.DataCombo DtcProveedor 
      Height          =   315
      Left            =   1680
      TabIndex        =   18
      Top             =   1080
      Width           =   5295
      _ExtentX        =   9340
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   5415
      Left            =   240
      TabIndex        =   20
      Top             =   1560
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   9551
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   855
      Left            =   1320
      TabIndex        =   29
      Top             =   7800
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      btype           =   5
      tx              =   "SALIR"
      enab            =   -1  'True
      font            =   "FrmPedido1.frx":3245
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      mptr            =   1
      micon           =   "FrmPedido1.frx":326D
      picn            =   "FrmPedido1.frx":328B
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.COSTO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   10245
      TabIndex        =   30
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNIDAD"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8820
      TabIndex        =   32
      Top             =   7080
      Width           =   585
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVACION:"
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
      Left            =   6705
      TabIndex        =   25
      Top             =   8040
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BUSQUEDA :"
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
      Left            =   7410
      TabIndex        =   22
      Top             =   1155
      Width           =   945
   End
   Begin VB.Label lblTotalpedido 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   19
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   795
      Left            =   6600
      Top             =   7680
      Width           =   5535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR:"
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
      Left            =   435
      TabIndex        =   17
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   540
      Left            =   240
      Top             =   960
      Width           =   13815
   End
   Begin VB.Label LblCantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   13020
      TabIndex        =   16
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "Calibri"
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
      TabIndex        =   15
      Top             =   7080
      Width           =   585
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2595
      TabIndex        =   14
      Top             =   7080
      Width           =   945
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1605
      TabIndex        =   13
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblAnulado 
      AutoSize        =   -1  'True
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   630
      Left            =   3840
      TabIndex        =   12
      Top             =   7680
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   540
      Left            =   255
      Top             =   7080
      Width           =   13815
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   705
      Left            =   255
      Top             =   120
      Width           =   13800
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   8775
      Left            =   0
      Top             =   0
      Width           =   15975
   End
End
Attribute VB_Name = "FrmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub CmdAgregar_Click()
If Val(Me.TxtCantidad.Text) > 0 And Len(Me.TxtCodProducto.Text) > 4 Then
    strCadena = "INSERT INTO temporal_pedido(id_producto,cantidad,precio_contado,total_contado,dni_save,ruc)VALUES " & _
    "('" & Trim(Me.TxtCodigo.Text) & "','" & Val(Me.TxtCantidad.Text) & "','" & Val(Me.txtcosto.Text) & "','" & Val(Me.txtcosto.Text) * Val(Me.TxtCantidad.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    Me.TxtCodigo.Text = ""
    Me.TxtCodProducto.Text = ""
    Me.TxtDescripcionProducto.Text = ""
    Me.TxtUnidad.Text = ""
    Me.TxtCantidad.Text = ""
    Call llenarGrid_grilla(Me.HfdDetalle)
    Call Resalta(Me.TxtCodProducto)
    
 End If
End Sub

Private Sub cmdProcesar_Click()
 Call Save
End Sub

Private Sub CmdQuitar_Click()
Me.HfdDetalle.col = 0
Call Quitar(Me.HfdDetalle.Text)
End Sub
Private Sub Quitar(ByVal codigo As String)
If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    strCadena = "DELETE FROM temporal_pedido WHERE id_temp='" & Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "'"
    CnBd.Execute (strCadena)
     
    Call llenarGrid_grilla(Me.HfdDetalle)
End If

End Sub

Private Sub llenar_producto(ByVal Grilla As MSHFlexGrid)
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 1400
           Grilla.ColWidth(1) = 5000
           Grilla.ColWidth(2) = 1400
           Grilla.ColWidth(3) = 1600
           Grilla.ColWidth(4) = 1600
       Next
        cabecera = "IDPRODUCTO" & vbTab & "PRODUCTO" & vbTab & "UNIDAD" & vbTab & "STOCK LOCAL" & vbTab & "STOCK TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
          Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & Format(rst("stock"), "#,##0.00") & vbTab & Format(BdAcumuladoCampo("almacen_producto", "id_producto", rst("id_producto"), "stock"), "#,##0.00")
          Grilla.AddItem Fila
          If rst("stock") < 5 Then
             For k = 0 To 4
                 Grilla.col = k
                 Grilla.Row = i + 1
                 Grilla.CellBackColor = &H8080FF
             Next k
          End If
        Fila = ""
        rst.MoveNext
             
        Next i
  End Sub



Private Sub Form_Load()
Dim nserie As String
 CenterForm Me
 Me.Top = 50
 strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  Me.DtcAlmacen.Enabled = False
  Me.TxtUsuario.Text = KEY_USUARIO
  
  
  strCadena = "SELECT A.id_doc,C.doc_des as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "'  AND A.id_doc='0103' AND A.id_alm='" & KEY_ALM & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount < 1 Then
    strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0103' AND ruc='" & KEY_RUC & "' ORDER BY serie DESC"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount < 1 Then
        strCadena = "INSERT INTO almacen_comprobante(ruc,id_alm,id_doc,serie,numero,id_formato_impresion)VALUES('" & KEY_RUC & "','" & KEY_ALM & "','0103','001','000001','1')"
    Else
        rstT.MoveFirst
        strCadena = "INSERT INTO almacen_comprobante(ruc,id_alm,id_doc,serie,numero,id_formato_impresion)VALUES('" & KEY_RUC & "','" & KEY_ALM & "','0103','" & formato_item(Val(rstT("serie")) + 1, 3) & "','000001','1')"
    End If
    CnBd.Execute (strCadena)
     
  Else
    
    
End If
 strCadena = "SELECT A.id_doc as Codigo,C.doc_des as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' AND A.id_alm='" & KEY_ALM & "' AND A.id_doc='0103'"
 Call ConfiguraRst(strCadena)
 Call LlenaDataCombo(Me.DtcTipoDoc)
  
  
  strCadena = "SELECT E.cod_unico as Codigo,P.nombre_completo as Descripcion FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni AND E.id_empresa='" & KEY_RUC & "' AND E.id_proveedor='si' "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcProveedor)
    Call nuevo
    Me.DtpActual.Value = KEY_FECHA
End Sub
Private Function CodigoDetalle() As String
    strCadena = "SELECT cDetalle FROM Detalle_Documento_Pedido ORDER BY cDetalle DESC"
    Call ConfiguraRst(strCadena)
    CodigoDetalle = GeneraCodigo(10)
    Set rst = Nothing
End Function
Private Sub Save()
Dim id_pedido As Double
If Me.DtcTipoDoc.BoundText <> "" And Me.TxtSerie.Text <> "" And Me.TxtNumeroDoc.Text <> "" Then
    strCadena = "P_insert_pedido('" & Me.DtcTipoDoc.BoundText & "','" & Me.TxtSerie.Text & "','" & Me.TxtNumeroDoc.Text & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_FECHA & "','" & Me.TxtObservacion.Text & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    
    id_pedido = LastRegistroRUC("movimiento_pedido", "id_pedido")
    StrNumero = FormatosCeros(Trim(str(Val(Me.TxtNumeroDoc.Text)) + 1), 6)
    strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "'  AND ruc='" & Trim(KEY_RUC) & "'"
    CnBd.Execute (strCadena)
     
    Call SaveDetalleDocumentoPedido(id_pedido)
   
End If
    
End Sub
Private Sub SaveDetalleDocumentoPedido(ByVal idPedido As Double)

   strCadena = "SELECT * FROM temporal_pedido WHERE  ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       For i = 0 To rstT.RecordCount - 1
           strCadena = "INSERT INTO movimiento_pedido_detalle(id_pedido,id_producto,cantidad,ruc) VALUES ('" & idPedido & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
            
           rstT.MoveNext
        Next i
       ' Call Nuevo
    End If
End Sub
Function GeneraCodTemporal() As Integer
Dim Codtemporal As Integer
Call ConfiguraRst(strCadena)
    If rst.EOF = True Then
        Codtemporal = 1
    Else
        Codtemporal = rst(0) + 1
    End If
  GeneraCodTemporal = Codtemporal
  Set rst = Nothing
End Function

Sub llenarGrid_det(ByVal Grilla As MSHFlexGrid)

On Error GoTo salir
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.Rows = 1
  Set Grilla.Recordset = rst
  Grilla.Rows = rst.RecordCount
  Grilla.ColWidth(0) = 800
  Grilla.ColWidth(1) = 6200
  
  
  
Me.LblCantidad.Caption = Trim(rst.RecordCount)
'Set rst = Nothing
  
  Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
  Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Sub llenarGrid_grilla(ByVal Grilla As MSHFlexGrid)
strCadena = "SELECT * FROM view_orden_pedido WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1400
           Grilla.ColWidth(2) = 5000
           Grilla.ColWidth(3) = 1500
           Grilla.ColWidth(4) = 1500
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 1500
    Next
        cabecera = "COD" & vbTab & "CODIGO" & vbTab & "PRODUCTO" & vbTab & "UNIDAD" & vbTab & "MARCA" & vbTab & "CANTIDAD" & vbTab & "P.COSTO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 7
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        in_total = 0
        For i = 0 To rst.RecordCount - 1
          in_total = in_total + rst("total_contado")
          Fila = rst("id_temp") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("unidad") & vbTab & rst("marca") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("precio_contado"), "#,##0.0000") & vbTab & Format(rst("total_contado"), "#,##0.0000")
          Grilla.AddItem Fila
          
        rst.MoveNext
             
        Next i
        cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(in_total, "###0.00")
        Grilla.AddItem cabecera
End Sub
Sub llenarGrid_detalle(ByVal Grilla As MSHFlexGrid, ByVal id_pedido As Double)
strCadena = "SELECT * FROM movimiento_pedido_detalle D,producto P,unidad U WHERE D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND  D.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND D.id_pedido='" & id_pedido & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1400
           Grilla.ColWidth(2) = 5000
           Grilla.ColWidth(3) = 1500
           Grilla.ColWidth(4) = 1500
           
    Next
        cabecera = "IDPEDIDO" & vbTab & "CODIGO" & vbTab & "PRODUCTO" & vbTab & "UNIDAD" & vbTab & "CANTIDAD"
        Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
          Fila = rst("id_detalle") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & Format(rst("cantidad"), "#,##0.00")
          Grilla.AddItem Fila
          Fila = ""
        rst.MoveNext
             
        Next i
   
End Sub

Private Sub HfProducto_Click()

End Sub



Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_NEW
        Call nuevo
    Case KEY_ANULAR
        If MsgBox(MSGANULAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
            Procedencia = anular
            FrmSeguridad.Show
        End If
    Case KEY_PRINT
            FrmPedidosListado.Show
    Case KEY_DELETE
        If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
           Procedencia = Eliminar
            FrmSeguridad.Show
       End If
    Case KEY_EXIT
      Unload Me
  End Select
End Sub
Public Sub nuevo()
    
    strCadena = "DELETE FROM temporal_pedido WHERE dni_save='" & Trim(KEY_USUARIO) & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    Call llenarGrid_grilla(Me.HfdDetalle)
    strCadena = "SELECT * FROM almacen_comprobante A WHERE  A.ruc='" & KEY_RUC & "' AND A.id_alm='" & KEY_ALM & "' AND A.id_doc='" & Me.DtcTipoDoc.BoundText & "'"
    Call ConfiguraRst(strCadena)
    Me.TxtSerie.Text = rst("serie")
    Me.TxtNumeroDoc.Text = rst("numero")
  
    Me.lblAnulado.Visible = False
    
    
    Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
    
    
End Sub
Public Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub

Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
     
      
    Case KEY_PRINT
        'Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Trim(Me.TxtSerie.Text), Trim(Me.TxtNumeroDoc.Text), "00001")
 End Select
  Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub TxtBusquedaRapido_Change()
  strCadena = "SELECT E.cod_unico as Codigo,P.nombre_completo as Descripcion FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni AND E.id_empresa='" & KEY_RUC & "' AND E.id_proveedor='si' AND P.nombre_completo LIKE '%" & Trim(Me.TxtBusquedaRapido.Text) & "%' "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcProveedor)
End Sub

Private Sub TxtBusquedaRapido_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcProveedor.Enabled = True Then
        Me.DtcProveedor.SetFocus
    End If
End If
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtcosto)
End If
End Sub



Private Sub TxtCodProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If KEY_BARRAS = "si" Then
    strCadena = "SELECT * FROM producto P,producto_barras B,unidad U WHERE P.id_unidad=U.id_und AND U.id_usu AND P.id_producto=B.id_producto AND B.cod_barra='" & Trim(Me.TxtCodProducto.Text) & "' AND P.ruc='" & KEY_RUC & "' AND B.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "'"
  Else
    Me.TxtCodProducto.Text = formato_item(Me.TxtCodProducto.Text, 5)
    strCadena = "SELECT * FROM producto P,unidad U WHERE P.id_unidad=U.id_und AND P.id_producto = '" & Trim(Me.TxtCodProducto.Text) & "' AND ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "'"
  End If
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
        Me.TxtCodigo.Text = rst("id_producto")
        Me.TxtDescripcionProducto.Text = rst("nombre_prod")
        Me.TxtUnidad.Text = rst("abreviatura")
        Call Resalta(Me.TxtCantidad)
 
 Else
        Procedencia = Selecionar
        FrmProducto.Show
 End If
End If
End Sub
Public Sub buscar_comprobante(Optional idPedido As Double)
    Me.TxtNumeroDoc.Text = FormatosCeros(Me.TxtNumeroDoc.Text, 6)
     strCadena = "SELECT * FROM movimiento_pedido WHERE (numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND ruc='" & KEY_RUC & "')"
    
    If FrmPedidosListado.Procedencia = buscar Then
        strCadena = "SELECT * FROM movimiento_pedido WHERE id_pedido='" & idPedido & "' AND ruc='" & KEY_RUC & "'"
        FrmPedidosListado.Procedencia = Neutro
    End If
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        strCadena = "SELECT * FROM movimiento_pedido_temporal WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
        Call llenarGrid_grilla(Me.HfdDetalle)
        Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
        End If
        
        
        
Else
        If idPedido = 0 Then
            idPedido = rst("id_pedido")
            
        End If
       
        Me.TxtUsuario.Text = rst("dni_save")
        Me.DtpActual.Value = rst("fecha")
        If rst("anulado") = "si" Then
            Me.lblAnulado.Visible = True
            Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
        Else
            Me.lblAnulado.Visible = False
            Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
        End If
        Call llenarGrid_detalle(Me.HfdDetalle, idPedido)
        
        Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
        
        
        Me.CmdAgregar.Enabled = False
        Me.CmdQuitar.Enabled = False
        Me.TxtCantidad.Enabled = False
        

    End If
End Sub






Private Sub TxtCosto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.CmdAgregar.SetFocus
End If
End Sub

Private Sub TxtNumeroDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscar_comprobante
End If
End Sub

Private Sub TxtProducto_Change()
strCadena = "SELECT * FROM almacen_producto A,producto P,unidad U WHERE A.id_producto=P.id_producto AND P.id_unidad=U.id_und AND A.id_alm='" & KEY_ALM & "' AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND P.id_proveedor='" & Me.DtcProveedor.BoundText & "' AND P.nombre_prod LIKE '%" & Trim(Me.TxtProducto.Text) & "%' "
Call llenar_producto(Me.HfdProducto)
End Sub

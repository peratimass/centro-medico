VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmDetalleDocumentoVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documento de Venta"
   ClientHeight    =   7260
   ClientLeft      =   2475
   ClientTop       =   2220
   ClientWidth     =   12900
   Icon            =   "FrmDetalleDocumentoVenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   12900
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdLimpiar 
      BackColor       =   &H00808080&
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      Picture         =   "FrmDetalleDocumentoVenta.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox TxtDescipcion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   7560
      TabIndex        =   36
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton CmdNuevoCliente 
      Caption         =   "Nuevo Cliente"
      Height          =   375
      Left            =   5595
      TabIndex        =   35
      ToolTipText     =   "Busca Cliente"
      Top             =   1560
      Width           =   1665
   End
   Begin VB.TextBox TxtProducto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1125
      MaxLength       =   80
      TabIndex        =   4
      Top             =   1545
      Width           =   3945
   End
   Begin VB.TextBox TxtEntidad 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1125
      MaxLength       =   50
      TabIndex        =   2
      Top             =   465
      Width           =   3945
   End
   Begin VB.CommandButton CmdEntidad 
      Caption         =   "..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      ToolTipText     =   "Busca Cliente"
      Top             =   825
      Width           =   495
   End
   Begin VB.CheckBox ChkCliente 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cliente"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   495
      Width           =   855
   End
   Begin VB.TextBox TxtNDocumento 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1125
      MaxLength       =   11
      TabIndex        =   3
      Top             =   900
      Width           =   1215
   End
   Begin VB.Frame FrmComprobante 
      Height          =   975
      Left            =   6045
      TabIndex        =   16
      Top             =   465
      Width           =   1215
      Begin VB.OptionButton OptFactura 
         Caption         =   "Factura"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   915
      End
      Begin VB.OptionButton OptBoleta 
         Caption         =   "Boleta"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox TxtCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1125
      TabIndex        =   6
      Top             =   2430
      Width           =   1215
   End
   Begin VB.TextBox TxtPrecio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1125
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1980
      Width           =   1215
   End
   Begin VB.TextBox TxtPrecioTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   3855
      MaxLength       =   6
      TabIndex        =   7
      Top             =   1980
      Width           =   1215
   End
   Begin VB.TextBox TxtSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   6045
      MaxLength       =   6
      TabIndex        =   15
      Top             =   4665
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtIgv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   6045
      MaxLength       =   6
      TabIndex        =   14
      Top             =   5145
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame FrmCondPago 
      Caption         =   "Pago"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   165
      TabIndex        =   13
      Top             =   4905
      Visible         =   0   'False
      Width           =   1335
      Begin VB.OptionButton OptCredito 
         Caption         =   "Crédito"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton OptContado 
         Caption         =   "Contado"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   450
      Left            =   6045
      MaxLength       =   6
      TabIndex        =   12
      Top             =   5565
      Width           =   1215
   End
   Begin VB.TextBox TxtFecha 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6045
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   1440
      Left            =   165
      TabIndex        =   19
      Top             =   3120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2540
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      ForeColorSel    =   16777215
      GridColor       =   -2147483635
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdProducto 
      Height          =   3360
      Left            =   7365
      TabIndex        =   5
      Top             =   1200
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5927
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   12582912
      ForeColorSel    =   16777215
      GridColor       =   -2147483635
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   2730
      TabIndex        =   20
      Top             =   5055
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1482
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1800
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
         TabIndex        =   21
         Top             =   30
         Width           =   1680
         _ExtentX        =   2963
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
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   2040
      Top             =   5280
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
            Picture         =   "FrmDetalleDocumentoVenta.frx":074C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoVenta.frx":0A68
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoVenta.frx":0EC8
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoVenta.frx":1328
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoVenta.frx":1644
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoVenta.frx":1AA4
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoVenta.frx":1DC0
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoVenta.frx":2220
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoVenta.frx":2680
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoVenta.frx":2F60
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoVenta.frx":327C
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoVenta.frx":3598
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbControl 
      Height          =   900
      Left            =   5595
      TabIndex        =   22
      Top             =   2152
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1588
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1665
      _CBHeight       =   900
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbControl 
         Height          =   780
         Left            =   30
         TabIndex        =   23
         Top             =   0
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1376
         ButtonWidth     =   1217
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar"
               Key             =   "(Agregar)"
               Object.ToolTipText     =   "Agregar"
               ImageKey        =   "(Agregar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Quitar"
               Key             =   "(Quitar)"
               Object.ToolTipText     =   "Quitar"
               ImageKey        =   "(Quitar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComCtl2.UpDown UdCantidad 
      Height          =   315
      Left            =   2340
      TabIndex        =   24
      Top             =   2430
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "TxtCantidad"
      BuddyDispid     =   196620
      OrigLeft        =   3900
      OrigTop         =   3075
      OrigRight       =   4140
      OrigBottom      =   3450
      Max             =   999
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   7440
      Picture         =   "FrmDetalleDocumentoVenta.frx":38B4
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   3915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Busqueda Rapida"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   7440
      TabIndex        =   37
      Top             =   0
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   7440
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label LblNDocumento 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Doc. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   240
      TabIndex        =   34
      Top             =   960
      Width           =   675
   End
   Begin VB.Label LblPrecioTotal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3255
      TabIndex        =   33
      Top             =   2040
      Width           =   465
   End
   Begin VB.Label LblCantidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   240
      TabIndex        =   32
      Top             =   2490
      Width           =   735
   End
   Begin VB.Label LblPrecio 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P. Unitario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   240
      TabIndex        =   31
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label LblProducto 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   240
      TabIndex        =   30
      Top             =   1605
      Width           =   765
   End
   Begin VB.Shape ShpEntidad 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   165
      Top             =   345
      Width           =   5055
   End
   Begin VB.Shape ShpProducto 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Left            =   165
      Top             =   1425
      Width           =   5055
   End
   Begin VB.Label LblSubTotal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubTotal :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5235
      TabIndex        =   29
      Top             =   4725
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LblIgv 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5235
      TabIndex        =   28
      Top             =   5205
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label LblTotal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5100
      TabIndex        =   27
      Top             =   5625
      Width           =   735
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5445
      TabIndex        =   26
      Top             =   172
      Width           =   555
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      Caption         =   "Nueva Venta"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   -15
      Width           =   5055
   End
End
Attribute VB_Name = "FrmDetalleDocumentoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EnumFrmDetalleVenta As EnumBuscarDocumento
Public StrCodEntidad As String

Dim StrCodDocumento As String, StrPersona As String
Dim StrCodEstado As String * 2, Tipo As String * 1
Dim DteFecha As Date, DtePago As Date, DteVencimiento As Date
Dim StrCodigo As String, Producto As String, Unidad As String
Dim Procedencia As String
Dim IntCantidad As Single, IntStock As Integer
Dim DblPrecio As Double
Dim DblTotal As Double, DblIGV As Double, DblSubTotal As Double
Dim Criterio As String


Private Sub Almacena()
  RstTemporal.MoveFirst
On Error GoTo Error
  CnBd.BeginTrans
    StrCodDocumento = Numero("DocumentoVenta", Tipo)
    StrCadena1 = "INSERT INTO DocumentoVenta(cDocumentoVenta,stipodocumento,ccliente,sPersona," & _
    " demisionventa, dvencimiento, dpago, nsubtotal, nigv, ntotalventa, cestado) VALUES " & _
    " ('" & StrCodDocumento & "','" & Tipo & "','" & StrCodEntidad & "','" & StrPersona & "'"

    
    Call EjecutaRST(StrCadena)
     If Me.EnumFrmDetalleVenta = BPedido Then
    StrCadena1 = "UPDATE Pedido SET cDocumentoVenta= '" & StrCodDocumento & "' " & _
    " "
      Call EjecutaRST(StrCadena)
    End If
    Do While Not RstTemporal.EOF
      StrCodigo = RstTemporal(0)
      IntCantidad = CDbl(RstTemporal(1))
      DblPrecio = CDbl(RstTemporal(4))
      '*** Registra Movimiento de Kárdex
      If Not Me.EnumFrmDetalleVenta = BPedido Then
        Call Kardex(StrCodigo, "S01", IntCantidad, DteFecha, StrCodDocumento, DblPrecio)
      End If
      StrCadena = "INSERT INTO detalleDocumentoVenta(cDocumentoVenta,cproducto,ncantidad," & _
      " nprecioventa)"
      Call EjecutaRST(StrCadena1)
      RstTemporal.MoveNext
    Loop
  CnBd.CommitTrans
  Set RstTemporal = Nothing
  MsgBox "Los Datos Fueron Almacenados Correctamente", vbOKOnly + vbInformation, "Mensaje para el Usuario"
  Exit Sub
Error:
  CnBd.RollbackTrans
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  MsgBox MSGREINGRESEDATOS, vbInformation + vbOKOnly, MSGGRABACION
  Exit Sub
End Sub

Private Sub ChkCliente_Click()
  TxtEntidad.Text = ""
  If ChkCliente.Value = 1 Then
    TxtEntidad.Enabled = False
    TxtNDocumento.Enabled = False
    CmdEntidad.Enabled = True
  Else
    TxtEntidad.Enabled = True
    TxtNDocumento.Enabled = True
    CmdEntidad.Enabled = False
  End If
End Sub

Private Sub ChkLinea_Click()
'If Me.ChkLinea Then
'Me.DtcLinea.Enabled = True
'Else
'Me.DtcLinea.Enabled = False
'End If
End Sub

Private Sub CmdBuscar_Click()
Dim Linea As String
'If ChkLinea.Value = 1 Then

        'Linea = Replace(DtcLinea.BoundText, "'", "''")
        
        'StrCadena = "SELECT cproducto as Código,sdescripcionproducto as Producto, " & _
  '" nprecioventa as Precio, nstockactual as Stock FROM " & _
 ' " Producto WHERE clinea LIKE '%" & Linea & "%'  ORDER BY cproducto"
  
 ' Call ConfiguraRst(StrCadena)
 ' If Not Rst.EOF Then
  
'Set HfdProducto.Recordset = Rst
'Me.HfdProducto.ColWidth(0) = 0
'Me.HfdProducto.ColWidth(1) = 2300
'Me.HfdProducto.ColWidth(2) = 700
'Me.HfdProducto.ColWidth(3) = 700
'Call DarFormato(HfdProducto, 2)
'Set Rst = Nothing
'End If
'Else
'Form_Activate
'End If

End Sub

Private Sub CmdEntidad_Click()
  FrmCliente.EnumFrmCliente = DocumentoVenta
 Me.Hide
 FrmCliente.Show
 
End Sub

Private Sub CmdLimpiar_Click()
Me.TxtDescipcion = Empty
StrCadena = "SELECT cproducto, sdescripcionproducto as Producto,nstockactual as Stock, " & _
  " nprecioventa as Precio FROM producto WHERE nstockactual > 0 ORDER BY cproducto"
  Call ConfiguraRst(StrCadena)
  Set HfdProducto.Recordset = Rst
  Set Rst = Nothing
  HfdProducto.ColWidth(0) = 0
  HfdProducto.ColWidth(1) = 2300
  HfdProducto.ColWidth(2) = 600
  HfdProducto.ColWidth(3) = 600
  Call DarFormato(HfdProducto, 3)
  

End Sub

Private Sub CmdNuevoCliente_Click()
  FrmCliente.Procedencia = Nuevo
  FrmDetalleCliente.Show
End Sub

Private Sub Form_Activate()
  StrCadena = "SELECT cproducto, sdescripcionproducto as Producto,nstockactual as Stock, " & _
  " nprecioventa as Precio FROM producto WHERE nstockactual > 0 ORDER BY cproducto"
  Call ConfiguraRst(StrCadena)
  Set HfdProducto.Recordset = Rst
  Set Rst = Nothing
  HfdProducto.ColWidth(0) = 0
  HfdProducto.ColWidth(1) = 2300
  HfdProducto.ColWidth(2) = 600
  HfdProducto.ColWidth(3) = 600
  Call DarFormato(HfdProducto, 3)
  
  Me.TxtDescipcion.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
CenterForm Me
StrCadena = "SELECT clinea as Codigo, sdescripcion as Descripcion FROM linea " & _
  "ORDER BY sdescripcion"
  Call ConfiguraRst(StrCadena)
  
  
  
  '*** configura un recordset vacio, tomando como referencia los campos de la tabla Detalle
  StrCadena = "SELECT texto as Cod_Prod, entero as Cant,texto as Unid,  texto as " & _
  " Descripcion, doble as Precio, doble as Total FROM tabla"
  Call ConfiguraTemporal(StrCadena)
  Set HfdDetalle.Recordset = RstTemporal
  HfdDetalle.ColWidth(0) = 0
  HfdDetalle.ColWidth(1) = 650
  HfdDetalle.ColWidth(2) = 800
  HfdDetalle.ColWidth(3) = 3700
  HfdDetalle.ColWidth(4) = 800
  HfdDetalle.ColWidth(5) = 800
  
  If Me.EnumFrmDetalleVenta = BPedido Then
    ChkCliente.Value = 1
    Call ChkCliente_Click
    HfdProducto.Enabled = False
    StrCadena = "SELECT pedido.ccliente, (snombrecliente & chr(32) & sapellidocliente), snumerodocumento " & _
    " FROM pedido INNER JOIN cliente ON cliente.ccliente = pedido.ccliente WHERE " & _
    " cPedido ='" & FrmPedido.StrCodDocumento & "'"
    Call EjecutaRST(StrCadena)
    StrCodEntidad = RstEjecuta(0)
    TxtEntidad.Text = RstEjecuta(1)
    TxtNDocumento.Text = RstEjecuta(2)
    StrCadena = "SELECT producto.cproducto,ncantidadPedido,sdescripcion,sdescripcionproducto," & _
    " nprecioPedido FROM detallePedido INNER JOIN (producto INNER JOIN unidad ON " & _
    " unidad.cunidad = producto.cunidad) ON producto.cproducto = detallepedido.cproducto WHERE " & _
    " cPedido ='" & FrmPedido.StrCodDocumento & "'"
    Call ConfiguraRst(StrCadena)
    Rst.MoveFirst
    Do While Not Rst.EOF
      RstTemporal.AddNew
      RstTemporal.Fields(0) = Trim(Rst(0))
      RstTemporal.Fields(1) = CInt(Rst(1))
      RstTemporal.Fields(2) = Trim(Rst(2))
      RstTemporal.Fields(3) = Trim(Rst(3))
      RstTemporal.Fields(4) = CDbl(Rst(4))
      RstTemporal.Fields(5) = CDbl(Rst(1) * Rst(4))
      Rst.MoveNext
    Loop
  End If
  
  TxtTotal.Text = Suma()
  Set HfdDetalle.Recordset = RstTemporal
  Call DarFormato(HfdDetalle, 4)
  Call DarFormato(HfdDetalle, 5)
  
  TxtFecha.Text = Date
  Call Limpia(False)
End Sub

Private Sub Limpia(ByVal Flag As Boolean)
  TlbAcciones.Buttons(KEY_CANCEL).Enabled = True
  If RstTemporal.RecordCount > 0 Then
    TlbAcciones.Buttons(KEY_SAVE).Enabled = True
  Else
    TlbAcciones.Buttons(KEY_SAVE).Enabled = False
  End If
  
  TlbControl.Buttons(KEY_AGREGAR).Enabled = Flag
  TlbControl.Buttons(KEY_QUITAR).Enabled = Flag
  
  TxtCantidad.Enabled = True
  TxtPrecio.Enabled = Flag
  TxtPrecioTotal.Enabled = Flag
  UdCantidad.Enabled = Flag
  
  If Flag = False Then
    TxtCantidad.Text = ""
    TxtPrecio.Text = ""
    TxtProducto.Text = ""
    TxtPrecioTotal.Text = ""
  End If
End Sub

Function Suma() As Double
  If RstTemporal.RecordCount > 0 Then
    RstTemporal.MoveFirst
    DblTotal = 0
    Do While Not RstTemporal.EOF
      DblTotal = CDbl(DblTotal + RstTemporal(5))
      RstTemporal.MoveNext
    Loop
    Suma = Format(DblTotal, "#,##0.00")
  End If
  StrCadena = "SELECT nigv FROM parametros"
  Call EjecutaRST(StrCadena)
  DblSubTotal = CDbl(DblTotal / (1 + RstEjecuta(0)))
  DblIGV = CDbl(DblTotal - DblSubTotal)
  TxtIgv.Text = Format(DblIGV, "#,##0.00")
  TxtSubTotal.Text = Format(DblSubTotal, "#,##0.00")
End Function

Private Sub Form_Unload(Cancel As Integer)
  FrmCliente.EnumFrmCliente = BuscarCliente
End Sub

Private Sub HfdDetalle_Click()
 If HfdDetalle.Row <> 0 Then
  HfdDetalle.Col = 0
  StrCodigo = HfdDetalle.Text
  RstTemporal.MoveFirst
  Do While Not RstTemporal.EOF
  If RstTemporal.Fields(0) = StrCodigo Then
    StrCadena = "SELECT nstockactual FROM producto WHERE cproducto = '" & StrCodigo & "' "
    Call EjecutaRST(StrCadena)
    IntStock = RstEjecuta(0)
    Set RstEjecuta = Nothing
    TxtCantidad.Text = RstTemporal.Fields(1).Value
    TxtProducto.Text = RstTemporal.Fields(3).Value
    TxtPrecio.Text = RstTemporal.Fields(4).Value
    TxtPrecioTotal.Text = RstTemporal.Fields(5).Value
    Exit Do
  End If
  RstTemporal.MoveNext
  Loop
  Set HfdDetalle.Recordset = RstTemporal
  Call DarFormato(HfdDetalle, 4)
  Call DarFormato(HfdDetalle, 5)
  Call Limpia(True)
  Procedencia = "M"
  End If
End Sub

Private Sub HfdProducto_Click()
  If HfdProducto.Row <> 0 Then
    HfdProducto.Col = 0
    StrCodigo = HfdProducto.Text
    If Not StrCodigo = "" Then
      StrCadena = "SELECT sdescripcionproducto,nprecioventa, sdescripcion, nstockactual " & _
      " FROM producto INNER JOIN unidad ON unidad.cunidad=producto.cunidad " & _
      " WHERE cproducto = '" & StrCodigo & "'"
      Call EjecutaRST(StrCadena)
      Producto = RstEjecuta(0)
      TxtProducto.Text = Producto
      TxtPrecio.Text = RstEjecuta(1)
      Unidad = RstEjecuta(2)
      IntStock = RstEjecuta(3)
      Set RstEjecuta = Nothing
      TxtCantidad.Text = 1
      Procedencia = ""
      TxtPrecioTotal.Text = ""
      Call Limpia(True)
      TlbControl.Buttons(KEY_QUITAR).Enabled = False
    End If
  End If
  Me.TxtCantidad.SetFocus
End Sub

Private Sub HfdProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If HfdProducto.Row <> 0 Then
    HfdProducto.Col = 0
    StrCodigo = HfdProducto.Text
    If Not StrCodigo = "" Then
      StrCadena = "SELECT sdescripcionproducto,nprecioventa, sdescripcion, nstockactual " & _
      " FROM producto INNER JOIN unidad ON unidad.cunidad=producto.cunidad " & _
      " WHERE cproducto = '" & StrCodigo & "'"
      Call EjecutaRST(StrCadena)
      Producto = RstEjecuta(0)
      TxtProducto.Text = Producto
      TxtPrecio.Text = RstEjecuta(1)
      Unidad = RstEjecuta(2)
      IntStock = RstEjecuta(3)
      Set RstEjecuta = Nothing
      TxtCantidad.Text = 1
      Procedencia = ""
      TxtPrecioTotal.Text = ""
      Call Limpia(True)
      TlbControl.Buttons(KEY_QUITAR).Enabled = False
    End If
  End If
End If
End Sub

Private Sub OptBoleta_Click()
  If OptBoleta.Value = True Then
    TxtSubTotal.Visible = False
    TxtIgv.Visible = False
    LblSubTotal.Visible = False
    LblIgv.Visible = False
  End If
End Sub

Private Sub OptCredito_Click()
  If ChkCliente.Value = 0 Then
    If StrCodEntidad = "" And Not TxtEntidad.Text = "" Then
      MsgBox MSGNOTCLIENTE, vbInformation, MSGVALIDACION
    End If
    ChkCliente.Value = 1
    CmdEntidad.Enabled = True
    TxtEntidad.Enabled = False
    TxtNDocumento.Enabled = False
  End If
End Sub

Private Sub OptFactura_Click()
  If OptFactura.Value = True Then
    If TxtEntidad.Text = "" Then
      ChkCliente.Value = 1
      CmdEntidad.Enabled = True
      TxtEntidad.Enabled = False
      TxtNDocumento.Enabled = False
    End If
    TxtSubTotal.Visible = True
    TxtIgv.Visible = True
    LblSubTotal.Visible = True
    LblIgv.Visible = True
  End If
End Sub

Private Sub Save()
  DteFecha = TxtFecha.Text
  If OptFactura.Value = True And TxtNDocumento.Text = "" Then
    MsgBox MSGFALTADATOS, vbInformation, MSGVALIDACION
  Else
    If ChkCliente.Value = 1 Then
      StrPersona = ""
      
    Else
      StrPersona = Left(Trim(TxtEntidad.Text), 50)
      StrCodEntidad = ""
    End If
    If OptBoleta.Value = True Then
      Tipo = "B"
    Else
      Tipo = "F"
    End If
    If OptContado.Value = True Then
      DtePago = DteFecha
      DteVencimiento = DteFecha
      StrCodEstado = "NN"
    Else
      If OptCredito.Value = True And StrCodEntidad = "" Then
          MsgBox MSGNOTCLIENTE, vbInformation, MSGVALIDACION
          Exit Sub
      Else
        DtePago = DTEMAXIMA
        DteVencimiento = InputBox(MSGFECHA, MSGVENCIMIENTO, Date)
        StrCodEstado = "PP"
      End If
    End If
    TxtTotal.Text = Suma()
    Call Almacena
    Unload Me
  End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.Key
    Case KEY_SAVE
      Call Save
    Case KEY_CANCEL
      If MsgBox(MSGCANCELAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Unload Me
      End If
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub TlbControl_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Bandera As Boolean
On Error Resume Next
  Bandera = False
  Select Case Button.Key
    Case KEY_AGREGAR
      If Not (TxtPrecioTotal.Text = "" Or TxtCantidad.Text = "") Then
        TxtPrecio.Text = Round(CDbl(TxtPrecioTotal.Text / TxtCantidad.Text), 2)
      Else
        If Not (TxtPrecioTotal.Text = "" Or TxtPrecio.Text = "") Then
          TxtCantidad.Text = CDbl(TxtPrecioTotal.Text / TxtPrecio.Text)
        End If
      End If
      If TxtCantidad.Text = "" Or TxtPrecio.Text = "" Then
        MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
      Else
        IntCantidad = (TxtCantidad.Text)
        If IntStock >= IntCantidad Then
          If TxtPrecioTotal.Text = "" Then
            DblPrecio = Format(CDbl(TxtPrecio.Text * TxtCantidad.Text), "#,##0.00")
          Else
            DblPrecio = CDbl(TxtPrecioTotal.Text)
          End If
          If Procedencia = "M" Then
            RstTemporal.MoveFirst
            Do While Not RstTemporal.EOF
              If RstTemporal(0) = StrCodigo Then
                RstTemporal.Update
                RstTemporal.Fields(1) = CInt(IntCantidad)
                RstTemporal.Fields(4) = CDbl(TxtPrecio.Text)
                RstTemporal.Fields(5) = CDbl(DblPrecio)
                Exit Do
              End If
              RstTemporal.MoveNext
            Loop
            Procedencia = ""
          Else
            If RstTemporal.RecordCount > 0 Then
              RstTemporal.MoveFirst
              Do While Not RstTemporal.EOF
                If RstTemporal(0) = StrCodigo Then
                  Bandera = True
                  Exit Do
                End If
                RstTemporal.MoveNext
              Loop
            End If
            If Bandera = False Then
              RstTemporal.AddNew
              RstTemporal.Fields(0) = StrCodigo
              RstTemporal.Fields(1) = IntCantidad
              RstTemporal.Fields(2) = Trim(Unidad)
              RstTemporal.Fields(3) = Trim(Producto)
              RstTemporal.Fields(4) = CDbl(TxtPrecio.Text)
              RstTemporal.Fields(5) = CDbl(DblPrecio)
              If ((IntStock - IntCantidad) <= 5) Then
                MsgBox "Su Stock se esta Agotando", vbInformation, "Mensaje Para el Usuario"
            End If
            Else
              MsgBox MSGDUPLICIDAD, vbInformation, MSGVALIDACION
            End If
          End If
            
          TxtTotal.Text = Format(Suma(), "#,##0.00")
          Set HfdDetalle.Recordset = RstTemporal
          Call DarFormato(HfdDetalle, 4)
          Call DarFormato(HfdDetalle, 5)
          Call Limpia(False)
        Else
          MsgBox MSGSTOCK, vbInformation, MSGVALIDACION
        End If
      End If
    Case KEY_QUITAR
        RstTemporal.MoveFirst
        Do While Not RstTemporal.EOF
          If RstTemporal.Fields(0) = StrCodigo Then
            RstTemporal.Delete
            Exit Do
          End If
          RstTemporal.MoveNext
        Loop
        TxtTotal.Text = Suma()
        Set HfdDetalle.Recordset = RstTemporal
        Call DarFormato(HfdDetalle, 4)
        Call DarFormato(HfdDetalle, 5)
        Call Limpia(False)
    End Select
End Sub

Private Sub TxtCantidad_Change()
Dim total As Double
total = Val(Me.TxtPrecio.Text) * Val(Me.TxtCantidad.Text)
Me.TxtPrecioTotal.Text = Format(total, "###0.00")
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
  'KeyAscii = ValidaNumero("I", KeyAscii)
End Sub

Private Sub TxtDescipcion_Change()

Criterio = Trim(Me.TxtDescipcion.Text)
If Trim(Me.TxtDescipcion.Text) <> "" Then
StrCadena = "SELECT cproducto,sdescripcionproducto as Descripcion, " & _
  " nprecioventa as Precio,nstockactual as Stock FROM Producto WHERE sdescripcionproducto " & _
  " LIKE '" & Criterio & "%' ORDER BY sdescripcionproducto "
    
  Call ConfiguraRst(StrCadena)
  
  If Not Rst.EOF Then
  
Set HfdProducto.Recordset = Rst
Me.HfdProducto.ColWidth(0) = 0
Me.HfdProducto.ColWidth(1) = 2300
Me.HfdProducto.ColWidth(2) = 700
Me.HfdProducto.ColWidth(3) = 700
Call DarFormato(HfdProducto, 2)
Set Rst = Nothing
End If
Else
Form_Activate
End If

End Sub

Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("D", KeyAscii)
End Sub

Private Sub TxtPrecioTotal_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("D", KeyAscii)
End Sub

Private Sub TxtProducto_Change()
Dim Criterio As String
  Criterio = Trim(TxtProducto.Text)
  StrCadena = "SELECT cproducto as Código,sdescripcionproducto as Descripción, " & _
  " nstockactual as Stock, nprecioventa AS Precio FROM Producto  WHERE  sdescripcionproducto  " & _
  " LIKE '%" & Criterio & "%' AND nstockactual > 0  ORDER BY cproducto"
  Call ConfiguraRst(StrCadena)
  Set HfdProducto.Recordset = Rst
  Call DarFormato(HfdProducto, 3)
End Sub

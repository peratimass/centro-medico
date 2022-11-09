VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDetalleDocumentoCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documento Compra"
   ClientHeight    =   6870
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   10920
   Icon            =   "FrmDetalleDocumentoCompra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   10920
   Begin VB.CommandButton CmdNuevoProducto 
      Caption         =   "Nuevo Producto"
      Height          =   375
      Left            =   5580
      TabIndex        =   7
      ToolTipText     =   "Busca Cliente"
      Top             =   1560
      Width           =   1665
   End
   Begin VB.TextBox TxtNDocumento 
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
      Left            =   1260
      TabIndex        =   28
      Top             =   1155
      Width           =   1215
   End
   Begin VB.TextBox TxtEntidad 
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
      Left            =   1260
      MaxLength       =   80
      TabIndex        =   27
      Top             =   720
      Width           =   3795
   End
   Begin VB.CommandButton CmdEntidad 
      Caption         =   "..."
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
      TabIndex        =   1
      ToolTipText     =   "Busca Cliente"
      Top             =   1080
      Width           =   495
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
      Left            =   6030
      TabIndex        =   15
      Top             =   495
      Width           =   1215
   End
   Begin VB.TextBox TxtTotal 
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
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   14
      Top             =   6420
      Width           =   1215
   End
   Begin VB.TextBox TxtCodigo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6030
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame FrmCondPago 
      Caption         =   "Pago"
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   150
      TabIndex        =   13
      Top             =   5520
      Width           =   1335
      Begin VB.OptionButton OptContado 
         Caption         =   "Contado"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptCredito 
         Caption         =   "Crédito"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
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
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   12
      Top             =   6000
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
      Left            =   6000
      MaxLength       =   6
      TabIndex        =   11
      Top             =   5520
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3840
      MaxLength       =   6
      TabIndex        =   5
      Top             =   2235
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
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   6
      Top             =   2235
      Width           =   1455
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
      Left            =   1260
      MaxLength       =   80
      TabIndex        =   2
      Top             =   1800
      Width           =   3795
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
      Left            =   1260
      TabIndex        =   4
      Top             =   2685
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   2160
      Left            =   150
      TabIndex        =   16
      Top             =   3240
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3810
      _Version        =   393216
      ForeColor       =   -2147483635
      FixedCols       =   0
      ForeColorFixed  =   -2147483635
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
      Height          =   6600
      Left            =   7380
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   11642
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
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   2715
      TabIndex        =   17
      Top             =   5910
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
         TabIndex        =   10
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
      Left            =   150
      Top             =   6240
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
            Picture         =   "FrmDetalleDocumentoCompra.frx":030A
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoCompra.frx":0626
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoCompra.frx":0A86
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoCompra.frx":0EE6
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoCompra.frx":1202
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoCompra.frx":1662
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoCompra.frx":197E
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoCompra.frx":1DDE
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoCompra.frx":223E
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoCompra.frx":2B1E
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoCompra.frx":2E3A
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleDocumentoCompra.frx":3156
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbControl 
      Height          =   900
      Left            =   5580
      TabIndex        =   18
      Top             =   2160
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
         TabIndex        =   19
         Top             =   30
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
      Left            =   2236
      TabIndex        =   20
      Top             =   2685
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "TxtCantidad"
      BuddyDispid     =   196624
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
   Begin VB.Label LblEntidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distribuidora :"
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
      Top             =   772
      Width           =   1005
   End
   Begin VB.Label LblPrecioTotal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   3120
      TabIndex        =   33
      Top             =   2287
      Width           =   495
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
      TabIndex        =   32
      Top             =   1852
      Width           =   765
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
      Top             =   2287
      Width           =   735
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
      TabIndex        =   30
      Top             =   2737
      Width           =   735
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
      TabIndex        =   29
      Top             =   1207
      Width           =   675
   End
   Begin VB.Shape ShpEntidad 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   150
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      Caption         =   "Nueva Compra"
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
      TabIndex        =   26
      Top             =   120
      Width           =   5055
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
      Left            =   5430
      TabIndex        =   25
      Top             =   547
      Width           =   555
   End
   Begin VB.Label LblTotalVenta 
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
      Left            =   5220
      TabIndex        =   24
      Top             =   6480
      Width           =   465
   End
   Begin VB.Label LblCodigo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº : "
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
      Left            =   5430
      TabIndex        =   23
      Top             =   157
      Width           =   345
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
      Left            =   5220
      TabIndex        =   22
      Top             =   6060
      Width           =   375
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
      Left            =   5220
      TabIndex        =   21
      Top             =   5580
      Width           =   735
   End
   Begin VB.Shape ShpProducto 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1455
      Left            =   150
      Top             =   1680
      Width           =   5055
   End
End
Attribute VB_Name = "FrmDetalleDocumentoCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StrCodEntidad As String

Dim StrCodDocumento As String
Dim StrCodEstado As String * 2
Dim DteFecha As Date, DtePago As Date, DteVencimiento As Date
Dim StrCodigo As String, Producto As String, Unidad As String
Dim Procedencia As String
Dim IntCantidad As Integer
Dim DblPrecio As Double
Dim DblTotal As Double, DblIGV As Double, DblSubTotal As Double

Private Sub Almacena()
  RstTemporal.MoveFirst
On Error GoTo Error
  CnBd.BeginTrans
    StrCodDocumento = Numero("GuiaRemision", "G")
    StrCadena = "INSERT INTO Facturacompra(cfactura,cdistribuidora,demisionfactura, " & _
    " dvencimiento, dpago, ntotalfactura, cestado) Values ('" & TxtCodigo.Text & "'," & _
    " '" & StrCodEntidad & "',cdate('" & DteFecha & "'),cdate('" & DteVencimiento & "')," & _
    " cdate('" & DtePago & "')," & DblTotal & ",'" & StrCodEstado & "')"
    Call EjecutaRST(StrCadena)
    StrCadena = "INSERT INTO GuiaRemision(cguiaremision,cfactura,cdistribuidora,dGuiaRemision, " & _
    " ntotalGuiaRemision) Values ('" & StrCodDocumento & "','" & TxtCodigo.Text & "'," & _
    " '" & StrCodEntidad & "',cdate('" & DteFecha & "')," & DblTotal & ")"
    Call EjecutaRST(StrCadena)
    Do While Not RstTemporal.EOF
      StrCodigo = RstTemporal(0)
      IntCantidad = CDbl(RstTemporal(1))
      DblPrecio = CDbl(RstTemporal(4))
      '*** Registra Movimiento de Kárdex
      Call Kardex(StrCodigo, "E01", IntCantidad, DteFecha, TxtCodigo.Text, DblPrecio)
      StrCadena = "INSERT INTO detallefacturacompra(cfactura,cdistribuidora,cproducto," & _
      " ncantidadfactura,npreciocompra) Values ('" & TxtCodigo.Text & "'," & _
      " '" & StrCodEntidad & "','" & StrCodigo & "'," & IntCantidad & "," & DblPrecio & ")"
      Call EjecutaRST(StrCadena)
      StrCadena = "INSERT INTO detalleguia(cguiaremision,cproducto,ncantidadguiaremision) " & _
      " VALUES ('" & StrCodDocumento & "','" & StrCodigo & "'," & IntCantidad & ")"
      Call EjecutaRST(StrCadena)
      RstTemporal.MoveNext
    Loop
  CnBd.CommitTrans
  Set RstTemporal = Nothing
  MsgBox "Los registros fueron grabados satisfactoriamente", vbOKOnly, "Grabar"
  Exit Sub
Error:
  CnBd.RollbackTrans
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  MsgBox MSGREINGRESEDATOS, vbInformation + vbOKOnly, MSGGRABACION
  Exit Sub
End Sub

Private Sub CmdEntidad_Click()
  FrmDistribuidora.EnumFrmDistribuidora = DocumentoCompra
  FrmDistribuidora.Show
End Sub

Private Sub CmdNuevoProducto_Click()
  FrmProducto.Procedencia = Nuevo
  FrmDetalleProducto.Show
End Sub

Private Sub Form_Activate()
  StrCadena = "SELECT cproducto, sdescripcionproducto as Producto,nstockactual " & _
  " as Stock, ncantidadreorden as Reorden FROM producto ORDER BY cproducto"
  Call ConfiguraRst(StrCadena)
  Set HfdProducto.Recordset = Rst
  Set Rst = Nothing
  HfdProducto.ColWidth(0) = 0
  HfdProducto.ColWidth(1) = 2000
  HfdProducto.ColWidth(2) = 450
  HfdProducto.ColWidth(3) = 550
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
  '*** configura un recordset vacio, tomando como referencia los campos de la tabla Detalle
  CenterForm Me
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
  Call DarFormato(HfdDetalle, 4)
  Call DarFormato(HfdDetalle, 5)
  
  TxtFecha.Text = Date
  TxtCodigo.Enabled = True
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
  
  TxtCantidad.Enabled = Flag
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
  FrmDistribuidora.EnumFrmDistribuidora = InicioDistribuidora
End Sub

Private Sub HfdDetalle_Click()
 If HfdDetalle.Row <> 0 Then
  HfdDetalle.Col = 0
  StrCodigo = HfdDetalle.Text
  RstTemporal.MoveFirst
  Do While Not RstTemporal.EOF
  If RstTemporal.Fields(0) = StrCodigo Then
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
      StrCadena = "SELECT sdescripcionproducto,ncantidadreorden, sdescripcion " & _
      " FROM producto INNER JOIN unidad ON unidad.cunidad=producto.cunidad " & _
      " WHERE cproducto = '" & StrCodigo & "'"
      Call EjecutaRST(StrCadena)
      Producto = RstEjecuta(0)
      TxtProducto.Text = Producto
      TxtCantidad.Text = RstEjecuta(1)
      Unidad = RstEjecuta(2)
      Procedencia = ""
      Set RstEjecuta = Nothing
      TxtPrecioTotal.Text = ""
      Call Limpia(True)
      TlbControl.Buttons(KEY_QUITAR).Enabled = False
    End If
  End If
End Sub

Private Sub HfdProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If HfdProducto.Row <> 0 Then
    HfdProducto.Col = 0
    StrCodigo = HfdProducto.Text
    If Not StrCodigo = "" Then
      StrCadena = "SELECT sdescripcionproducto,ncantidadreorden, sdescripcion " & _
      " FROM producto INNER JOIN unidad ON unidad.cunidad=producto.cunidad " & _
      " WHERE cproducto = '" & StrCodigo & "'"
      Call EjecutaRST(StrCadena)
      Producto = RstEjecuta(0)
      TxtProducto.Text = Producto
      TxtCantidad.Text = RstEjecuta(1)
      Unidad = RstEjecuta(2)
      Procedencia = ""
      Set RstEjecuta = Nothing
      TxtPrecioTotal.Text = ""
      Call Limpia(True)
      TlbControl.Buttons(KEY_QUITAR).Enabled = False
    End If
  End If
End If
End Sub

Private Sub Save()
  DteFecha = TxtFecha.Text
  If OptContado.Value = True Then
    DtePago = DteFecha
    DteVencimiento = DteFecha
    StrCodEstado = "NN"
  Else
    DtePago = DTEMAXIMA
    DteVencimiento = InputBox(MSGFECHA, MSGVENCIMIENTO, Date)
    StrCodEstado = "PP"
  End If
  TxtTotal.Text = Suma()
  Call Almacena
  Unload Me
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
      If TxtCodigo.Text = "" Or TxtEntidad.Text = "" Or TxtCantidad.Text = "" Or TxtPrecio.Text = "" Then
        MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
      Else
        TxtCodigo.Text = Format(TxtCodigo.Text, "0000000000")
        StrCadena = "SELECT cfactura FROM facturacompra WHERE cfactura = " & _
        " '" & TxtCodigo.Text & "' AND cdistribuidora = '" & StrCodEntidad & "'"
        Call EjecutaRST(StrCadena)
        If RstEjecuta.EOF Then
          IntCantidad = CInt(TxtCantidad.Text)
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
              RstTemporal.Fields(1) = CInt(IntCantidad)
              RstTemporal.Fields(2) = Trim(Unidad)
              RstTemporal.Fields(3) = Trim(Producto)
              RstTemporal.Fields(4) = CDbl(TxtPrecio.Text)
              RstTemporal.Fields(5) = CDbl(DblPrecio)
            Else
              MsgBox MSGDUPLICIDAD, vbInformation, MSGVALIDACION
            End If
          End If
          TxtTotal.Text = Suma()
          Set HfdDetalle.Recordset = RstTemporal
          Call DarFormato(HfdDetalle, 4)
          Call DarFormato(HfdDetalle, 5)
          TxtCodigo.Enabled = False
          Call Limpia(False)
        Else
          MsgBox MSGDUPLICIDADFACTURA, vbInformation, MSGVALIDACION
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

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("I", KeyAscii)
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
  StrCadena = "SELECT cproducto as Código,sdescripcionproducto as Descripción,nstockactual " & _
  " as Stock, ncantidadreorden AS Reorden FROM Producto  WHERE  sdescripcionproducto " & _
  " LIKE '%" & Criterio & "%' ORDER BY cproducto"
  Call ConfiguraRst(StrCadena)
  Set HfdProducto.Recordset = Rst
'   Set Rst = Nothing
End Sub

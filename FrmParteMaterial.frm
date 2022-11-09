VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmParteMaterial 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   14190
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtIdConductor 
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
      Left            =   9360
      MaxLength       =   80
      TabIndex        =   41
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Visualizar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12960
      TabIndex        =   53
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtid_venta 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      TabIndex        =   51
      Top             =   4320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtTiempoRuta 
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
      Left            =   9360
      MaxLength       =   80
      TabIndex        =   49
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox TxtHoraRetorno 
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
      Left            =   12840
      MaxLength       =   80
      TabIndex        =   48
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox TxtHoraInicio 
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
      Left            =   9360
      MaxLength       =   80
      TabIndex        =   46
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox TxtBusquedaPlaca 
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
      Height          =   315
      Left            =   12960
      MaxLength       =   80
      TabIndex        =   38
      Top             =   850
      Width           =   975
   End
   Begin VB.TextBox TxtSerie 
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
      Left            =   4080
      MaxLength       =   80
      TabIndex        =   19
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox TxtNumeroDoc 
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
      Left            =   4875
      MaxLength       =   80
      TabIndex        =   18
      Top             =   240
      Width           =   1050
   End
   Begin VB.TextBox TxtRucDestino 
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
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox TxtNombreDestino 
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
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   16
      Top             =   1875
      Width           =   5895
   End
   Begin VB.TextBox TxtDireccionDestino 
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
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   15
      Top             =   2190
      Width           =   5895
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
      Left            =   120
      MaxLength       =   80
      TabIndex        =   14
      Top             =   6285
      Width           =   1695
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
      Left            =   2700
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   13
      Top             =   6285
      Width           =   6135
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
      Left            =   1830
      MaxLength       =   80
      TabIndex        =   12
      Top             =   6285
      Width           =   855
   End
   Begin VB.CommandButton CmdAgregar 
      BackColor       =   &H0080FFFF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6300
      Width           =   495
   End
   Begin VB.CommandButton CmdQuitar 
      BackColor       =   &H0080FFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6300
      Width           =   495
   End
   Begin VB.TextBox TxtPeso 
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
      Left            =   8880
      MaxLength       =   80
      TabIndex        =   9
      Top             =   6285
      Width           =   735
   End
   Begin VB.TextBox TxtLugarDescarga 
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
      Left            =   9360
      MaxLength       =   80
      TabIndex        =   45
      Top             =   2205
      Width           =   4575
   End
   Begin VB.TextBox TxtUnidad 
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
      Left            =   9900
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   8
      Top             =   6285
      Width           =   1215
   End
   Begin VB.TextBox TxtObservacion 
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
      Height          =   645
      Left            =   8160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   6840
      Width           =   5055
   End
   Begin VB.CheckBox ChkExtraer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "EXTRAER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   915
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   800
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox TxtNumero_guia 
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
         Left            =   3525
         MaxLength       =   80
         TabIndex        =   4
         Top             =   130
         Width           =   1215
      End
      Begin VB.TextBox TxtSeri_guia 
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
         Left            =   2760
         MaxLength       =   80
         TabIndex        =   3
         Top             =   130
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DtcComprobanteGuia 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   130
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
   End
   Begin VB.TextBox TxtId_parte 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtLugarCarga 
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
      Left            =   9360
      MaxLength       =   80
      TabIndex        =   43
      Top             =   1875
      Width           =   4575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   6000
      TabIndex        =   17
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
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
      Format          =   121176065
      CurrentDate     =   41139
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12960
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteMaterial.frx":0000
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteMaterial.frx":08DA
            Key             =   "(GuiaRemision)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteMaterial.frx":0BF4
            Key             =   "(Imprimir)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   120
      TabIndex        =   20
      Top             =   6600
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   1482
      BandCount       =   1
      ForeColor       =   -2147483635
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   3630
      _CBHeight       =   840
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   780
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbGrabar 
         Height          =   780
         Left            =   120
         TabIndex        =   21
         Top             =   15
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   1376
         ButtonWidth     =   1376
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Grabar"
               Key             =   "(Grabar)"
               Object.ToolTipText     =   "Grabar Ctrl+G"
               ImageKey        =   "(Grabar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "+ Viajes"
               Key             =   "(Verificar)"
               ImageKey        =   "(GuiaRemision)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Imprimir"
               Key             =   "(Imprimir)"
               Object.ToolTipText     =   "Grabar Ctrl+I"
               ImageKey        =   "(Imprimir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDetalle 
      Height          =   2775
      Left            =   120
      TabIndex        =   22
      Top             =   3480
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   4895
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
   Begin MSDataListLib.DataCombo DtcAlmacenOrigen 
      Height          =   315
      Left            =   9360
      TabIndex        =   23
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
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
      Left            =   360
      TabIndex        =   24
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
      Left            =   5880
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteMaterial.frx":0C81
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteMaterial.frx":10D5
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteMaterial.frx":13F5
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteMaterial.frx":1849
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteMaterial.frx":1C9D
            Key             =   "(Atender)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteMaterial.frx":1FBD
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteMaterial.frx":22DD
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParteMaterial.frx":25FD
            Key             =   "(Declarar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   4065
      Left            =   13245
      TabIndex        =   25
      Top             =   3480
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   7170
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   4065
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   26
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1482
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
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
               Caption         =   "Busqueda"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Buscar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Anular"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Anular)"
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
   Begin MSDataListLib.DataCombo DtcTransporte 
      Height          =   315
      Left            =   9360
      TabIndex        =   39
      Top             =   850
      Width           =   3495
      _ExtentX        =   6165
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
   Begin MSDataListLib.DataCombo DtcParteMaquina 
      Height          =   315
      Left            =   9360
      TabIndex        =   40
      Top             =   1200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   ""
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
   Begin VB.Label lblItems 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   12240
      TabIndex        =   56
      Top             =   6360
      Width           =   885
   End
   Begin VB.Label lblRazonTransporte 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   10785
      TabIndex        =   42
      Top             =   1560
      Width           =   3120
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "CONDUCTOR :"
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
      TabIndex        =   55
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DATOS CLIENTE / OBRA"
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
      Left            =   1560
      TabIndex        =   54
      Top             =   1320
      Width           =   1860
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "PARTE MAQUINA :"
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
      Left            =   7935
      TabIndex        =   52
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "TIEMPO RUTA :"
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
      TabIndex        =   50
      Top             =   3000
      Width           =   1125
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "HORA RETORNO :"
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
      Left            =   11280
      TabIndex        =   47
      Top             =   2640
      Width           =   1320
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "HORA INICIO :"
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
      Left            =   8190
      TabIndex        =   44
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   7455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "LUGAR CARGA :"
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
      Left            =   7875
      TabIndex        =   37
      Top             =   240
      Width           =   1170
   End
   Begin VB.Label lblruc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   690
      TabIndex        =   36
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label lblrazon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RAZON SOCIAL :"
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
      Left            =   210
      TabIndex        =   35
      Top             =   1920
      Width           =   1230
   End
   Begin VB.Label lbldireccion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Left            =   480
      TabIndex        =   34
      Top             =   2280
      Width           =   960
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "VEHICULO/PLACA :"
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
      Left            =   7890
      TabIndex        =   32
      Top             =   915
      Width           =   1395
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "LUGAR DESCARGA :"
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
      Left            =   7830
      TabIndex        =   31
      Top             =   2205
      Width           =   1455
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "M3"
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
      Left            =   9645
      TabIndex        =   30
      Top             =   6360
      Width           =   210
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVACION :"
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
      TabIndex        =   29
      Top             =   6600
      Width           =   1185
   End
   Begin VB.Label lblanulado 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1560
      TabIndex        =   28
      Top             =   2520
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "LUGAR CARGA :"
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
      Left            =   8115
      TabIndex        =   27
      Top             =   1875
      Width           =   1170
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   2580
      Left            =   7680
      Top             =   795
      Width           =   6375
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   2100
      Left            =   120
      Top             =   780
      Width           =   7455
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   615
      Left            =   7680
      Top             =   120
      Width           =   6375
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
      Height          =   300
      Left            =   12165
      TabIndex        =   33
      Top             =   6300
      Width           =   975
   End
End
Attribute VB_Name = "FrmParteMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Public cprod As String
Dim strMotivo As Integer

Private Sub ChkExtraer_Click()
If Me.ChkExtraer.Value = 1 Then
    Me.Frame1.Visible = True
    Me.DtcComprobanteGuia.SetFocus
Else
    Me.Frame1.Visible = False
End If
End Sub

Private Sub CmdAgregar_Click()
If Val(Me.txtcantidad.Text) > 0 And Trim(Me.TxtCodProducto.Text) <> "" Then
    If Val(Me.TxtId_parte.Text) < 1 Then
    strCadena = "INSERT INTO movimiento_transferencia_temporal(id_doc,serie,numero,id_producto,cantidad,peso,total,dni_save,ruc) VALUES " & _
    "('" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.TxtSerie.Text) & "','" & Me.TxtNumeroDoc.Text & "','" & cprod & "','" & Val(Me.txtcantidad.Text) & "','" & Val(Me.txtpeso.Text) & "'," & _
    "'" & Val(Me.txtpeso.Text) * Val(Me.txtcantidad.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    Call Llenar_Temporal(Me.HfDetalle)
    Else
    strCadena = "INSERT INTO parte_material_detalle(id_parte,id_producto,cantidad,peso,total,ruc) VALUES ('" & Val(Me.TxtId_parte.Text) & "','" & Trim(Me.TxtCodProducto.Text) & "','" & Val(Me.txtcantidad.Text) & "','" & Val(Me.txtpeso.Text) & "','" & Val(Me.txtpeso.Text) * Val(Me.txtcantidad.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    Call llenar_detalle(Me.HfDetalle, Val(Me.TxtId_parte.Text))
    End If
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
    Me.TxtCodProducto.Text = ""
    Me.txtcantidad.Text = ""
    Me.TxtDescripcionProducto.Text = ""
    Me.txtunidad.Text = ""
    Me.txtpeso.Text = ""
    Call Resalta(Me.TxtCodProducto)
Else
    Call Resalta(Me.txtcantidad)
End If
End Sub

Private Sub CmdQuitar_Click()
If MsgBox("ESTA SEGURO DE ELIMINAR ESTE REGISTRO", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
    strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE id_temporal='" & Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    Call Llenar_Temporal(Me.HfDetalle)
End If
End Sub



Private Sub DtcComprobanterel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtSeri_guia)
End If
End Sub

Private Sub Command1_Click()
strCadena = "SELECT * FROM parte_maquinaria WHERE id_parte='" & Val(Me.DtcParteMaquina.BoundText) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
     'Call FrmParteDiaria.Nuevo
     FrmParteDiaria.DtcTipoDoc.BoundText = rstT("id_doc")
     FrmParteDiaria.TxtSerie.Text = rstT("serie")
     FrmParteDiaria.TxtNumeroDoc.Text = rstT("numero")
    
    Call FrmParteDiaria.buscar_comprobante(Me.DtcParteMaquina.BoundText)
End If
End Sub

Private Sub DtcAlmacenOrigen_Change()
strCadena = "SELECT * FROM persona WHERE dni='" & Me.DtcAlmacenOrigen.BoundText & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    Me.txtLugarCarga.Text = UCase(rstT("direccion"))
End If
End Sub


Private Sub DtcComprobanteGuia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtSeri_guia)
End If
End Sub

Private Sub DtcParteMaquina_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM persona P,parte_maquinaria M WHERE P.dni=M.id_operador AND M.ruc='" & KEY_RUC & "' AND M.id_parte='" & Me.DtcParteMaquina.BoundText & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    Me.TxtIdConductor.Text = rstT("dni")
    Me.lblRazonTransporte.Caption = UCase(rstT("nombre_completo"))
    Call Resalta(Me.txtHoraInicio)
End If
End If
End Sub

Private Sub DtcTransporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT P.id_parte as Codigo,CONCAT(C.doc_abrev,':',P.serie,':',P.numero,'  --->  ',(P.petroleo+P.gasolina+P.aceite+P.grasa)) as Descripcion FROM parte_maquinaria P,comprobantes C WHERE P.id_doc=C.id_doc AND  P.id_transporte='" & Trim(Me.DtcTransporte.BoundText) & "' AND P.ruc='" & KEY_RUC & "' AND P.anulado='no' ORDER BY id_parte DESC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcParteMaquina)
If Me.DtcParteMaquina.Enabled = True Then
    Me.DtcParteMaquina.SetFocus
End If
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 200
Me.DTPicker1.Value = KEY_FECHA

strCadena = "SELECT T.id_transporte as Codigo,CONCAT(TT.descripcion,'-',T.placa) as Descripcion FROM transporte T,transporte_tipo TT WHERE T.id_tipo_transporte=TT.id_tipo_transporte AND T.ruc='" & KEY_RUC & "' AND TT.ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTransporte)
 
strCadena = "SELECT DISTINCT A.id_doc as Codigo, C.doc_des as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' AND C.id_doc='0202' AND id_alm='" & KEY_ALM & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    MsgBox "CREE EL COMPROBANTE PARA ESTA SUCURSAL", vbInformation, KEY_EMPRESA
   
    Exit Sub
End If
Call LlenaDataCombo(Me.DtcTipoDoc)

strCadena = "SELECT DISTINCT A.id_doc as Codigo, C.doc_des as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' AND C.id_doc<>'0202' AND C.id_doc<>'0201' AND id_alm='" & KEY_ALM & "' AND A.venta='si' "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComprobanteGuia)

strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0202' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
Call ConfiguraRst(strCadena)
Me.TxtSerie.Text = rst("serie")
Me.TxtNumeroDoc.Text = rst("numero")
Call Llenar_Temporal(Me.HfDetalle)

strCadena = "SELECT E.cod_unico as Codigo,P.nombre_completo as Descripcion FROM persona P,entidad_empresa E WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' AND E.id_proveedor='si'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacenOrigen)
Me.DtcAlmacenOrigen.BoundText = KEY_ALM





Me.TlbGrabar.Buttons("(Verificar)").Enabled = False

End Sub

Private Sub BuscarResponsable(ByVal ruc As String)
If (Trim(ruc) = "") Then
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
End If
    strCadena = "SELECT *  FROM persona WHERE dni='" & ruc & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.TxtRuc.Text = ruc
        FrmDetallePersona.ChkPersonal.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.TxtRucDestino.Text = rst("dni")
        Me.TxtNombreDestino.Text = rst("nombre_completo")
        Me.TxtDireccionDestino.Text = rst("direccion")
        Me.TxtLugarDescarga.Text = rst("direccion")
        Me.DtcTransporte.SetFocus
        Exit Sub
       
    End If

End Sub
Private Sub BuscarTransporte(ByVal ruc As String)
If (Trim(ruc) = "") Then
    Procedencia = buscar
    FrmPersona.Show
    Exit Sub
End If
    strCadena = "SELECT *  FROM persona WHERE dni='" & ruc & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.TxtRuc.Text = ruc
        FrmDetallePersona.ChkTransporte.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.TxtIdConductor.Text = rst("dni")
        Me.lblRazonTransporte.Caption = rst("nombre_completo")
        Me.DtcTransporte.SetFocus
        Exit Sub
       
    End If

End Sub

Private Sub HfDetalle_DblClick()
If Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) > 0 Then
    If Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True Then
        FrmTransferencias_detalle.Show
    End If
End If
End Sub

Private Sub HfDetalle_SelChange()
If Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) > 0 Then
    Me.CmdQuitar.Visible = True
Else
    Me.CmdQuitar.Visible = False
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_NEW
        Call nuevo
    Case KEY_UPDATE
         Procedencia = buscar
         FrmPartematerialLista.Show
    Case KEY_DELETE
            Procedencia = anular
            FrmSeguridad.Show
            Exit Sub
    Case KEY_EXIT
        Unload Me
End Select
End Sub
Public Sub nuevo()
strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
 
strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0202' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
Call ConfiguraRst(strCadena)
Me.DtcTipoDoc.Enabled = True
Me.TxtSerie.Enabled = True
Me.TxtNumeroDoc.Enabled = True
Me.TxtSerie.Text = rst("serie")
Me.TxtNumeroDoc.Text = rst("numero")
Me.TxtRucDestino.Text = ""
Me.TxtNombreDestino.Text = ""
Me.TxtDireccionDestino.Text = ""
Me.txtHoraInicio.Text = ""
Me.TxtHoraRetorno.Text = ""
Me.txtTiempoRuta.Text = ""
Me.lblRazonTransporte.Caption = ""
Me.TxtLugarDescarga.Text = ""
Me.TxtIdConductor.Text = ""

Me.lblAnulado.Visible = False
Me.TxtId_parte.Text = ""

Me.DTPicker1.Value = KEY_FECHA
Call Resalta(Me.TxtRucDestino)
Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
Call Llenar_Temporal(Me.HfDetalle)
End Sub
Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_SAVE
        Call Save
    
    Case "(Verificar)"
         Procedencia = Modificar
         FrmSeguridad.Show
         Exit Sub
         
    Case KEY_PRINT
      
        'Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Me.TxtSerie.Text, Me.TxtNumeroDoc.Text, "00001")
        
End Select
End Sub
Private Sub Save()

If Me.DtcTipoDoc.BoundText = "" Or Me.TxtSerie.Text = "" Or Me.TxtNumeroDoc.Text = "" And Trim(Me.TxtIdConductor.Text) = "" Then
   MsgBox "LLENE TODOS LOS PARAMETROS", vbInformation, KEY_EMPRESA
   Exit Sub
Else
If Val(Me.TxtId_parte.Text) < 1 Then
        
        
        strCadena = "INSERT INTO parte_material(id_parte_maquinaria,fecha,id_doc,serie,numero,id_cliente,id_transporte,id_conductor,origen,destino,hora_inicio,hora_retorno,tiempo_ruta,observacion,dni_save,ruc) " & _
        "VALUES('" & Val(Me.DtcParteMaquina.BoundText) & "','" & Format(Me.DTPicker1.Value, "YYYY-mm-dd") & "','" & Me.DtcTipoDoc.BoundText & "','" & Me.TxtSerie.Text & "','" & Me.TxtNumeroDoc.Text & "','" & Trim(Me.TxtRucDestino.Text) & "','" & Trim(Me.DtcTransporte.BoundText) & "','" & Trim(Me.TxtIdConductor.Text) & "'," & _
        "'" & Trim(Me.txtLugarCarga.Text) & "','" & Trim(Me.TxtLugarDescarga.Text) & "','" & Trim(Me.txtHoraInicio.Text) & "','" & Trim(Me.TxtHoraRetorno.Text) & "','" & Trim(Me.txtTiempoRuta.Text) & "','" & Trim(Me.TxtObservacion.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        Me.TxtId_parte.Text = LastRegistro("parte_material", "id_material")
    Else
        
    
        
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
    StrNumero = FormatosCeros(Trim(str(Val(Me.TxtNumeroDoc.Text)) + 1), 6)
    strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE  id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND ruc='" & Trim(KEY_RUC) & "'"
    CnBd.Execute (strCadena)
     
    Call savedetalle(Val(Me.TxtId_parte.Text))
    End If
End If

End Sub
Private Sub savedetalle(ByVal id_transferencia As Double)
strCadena = "SELECT * FROM movimiento_transferencia_temporal WHERE (numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "')"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       For i = 0 To rstT.RecordCount - 1
           strCadena = "INSERT INTO parte_material_detalle(id_parte,id_producto,cantidad,peso,total,ruc) VALUES ('" & id_transferencia & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("peso") & "','" & rstT("total") & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
            
           rstT.MoveNext
        Next i
        strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "'"
        CnBd.Execute (strCadena)
         
    End If
End Sub


Private Sub TxtBusquedaPlaca_Change()
strCadena = "SELECT T.id_transporte as Codigo,CONCAT(TT.descripcion,'-',T.placa) as Descripcion FROM transporte T,transporte_tipo TT WHERE T.id_tipo_transporte=TT.id_tipo_transporte AND T.ruc='" & KEY_RUC & "' AND TT.ruc='" & KEY_RUC & "' AND T.placa LIKE '%" & Trim(Me.TxtBusquedaPlaca.Text) & "%'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTransporte)
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Me.txtcantidad.Text) > 0 Then
        Me.cmdagregar.SetFocus
    Else
        Call Resalta(Me.txtcantidad)
    End If
End If
End Sub

Private Sub TxtCodProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (Len(Me.TxtCodProducto.Text) = 0) Or Val(Me.TxtCodProducto.Text) = 0 Then
        Call Resalta(Me.TxtCodProducto)
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
    End If
    
 
    If Trim(Mid(Me.TxtCodProducto.Text, 1, 2)) = "00" And Len(Me.TxtCodProducto.Text) > 8 Then
       Me.txtcantidad.Text = Val(Mid(Trim(Me.TxtCodProducto.Text), 8, 4) / 1000)
       Me.TxtCodProducto.Text = Mid(Me.TxtCodProducto, 3, 5)
    End If
    
    If KEY_BARRAS = "si" Then
        strCadena = "SELECT B.id_producto,P.nombre_prod,P.precio_venta,P.peso,P.id_igv,U.abreviatura FROM producto_barras B,producto P,unidad U WHERE B.id_producto=P.id_producto AND B.ruc='" & KEY_RUC & "' " & _
        "AND P.ruc='" & KEY_RUC & "' AND B.cod_barra='" & Trim(Me.TxtCodProducto.Text) & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "'"
    Else
        Me.TxtCodProducto.Text = FormatosCeros(Me.TxtCodProducto.Text, 5)
        strCadena = "SELECT A.id_producto, P.nombre_prod,P.precio_venta,P.peso,U.abreviatura FROM almacen_producto A,producto P,unidad U WHERE  P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND A.id_producto=P.id_producto AND A.id_alm='" & KEY_ALM & "' AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND A.id_producto='" & Trim(Me.TxtCodProducto.Text) & "'"
    End If
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        cprod = rst("id_producto")
        Me.TxtDescripcionProducto.Text = rst("nombre_prod")
        Me.txtunidad.Text = rst("abreviatura")
        Me.txtpeso.Text = rst("peso")
        Call Resalta(Me.txtcantidad)
        Exit Sub
    End If
        
End If
End Sub

Private Sub TxtLicencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtCodProducto)
End If
End Sub

Private Sub TxtHoraInicio_LostFocus()
Me.txtHoraInicio.Text = Format(Me.txtHoraInicio.Text, "hh:mm:ss")
End Sub

Private Sub TxtHoraRetorno_KeyPress(KeyAscii As Integer)
Dim INICIO As Variant
Dim final As Variant
If KeyAscii = 13 Then
Me.TxtHoraRetorno.Text = Format(Me.TxtHoraRetorno.Text, "hh:mm:ss")
INICIO = Format(Me.txtHoraInicio.Text, "hh:mm:ss")
final = Format(Me.TxtHoraRetorno.Text, "hh:mm:ss")
Me.txtTiempoRuta.Text = Format(TimeValue(final) - TimeValue(INICIO), "hh:mm:ss")

End If
End Sub

Private Sub TxtIdConductor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Call BuscarTransporte(Me.TxtIdConductor.Text)
End If
End Sub



Private Sub TxtNumero_guia_KeyPress(KeyAscii As Integer)
Dim idVenta As Double
If KeyAscii = 13 Then
   ' Me.txtid_venta.text = 0
    Me.TxtNumero_guia.Text = FormatosCeros(Me.TxtNumero_guia.Text, 6)
    strCadena = "SELECT * FROM movimiento_venta WHERE (numero='" & Trim(Me.TxtNumero_guia.Text) & "' AND id_doc='" & Trim(Me.DtcComprobanteGuia.BoundText) & "' AND serie='" & Trim(Me.TxtSeri_guia.Text) & "'  AND ruc='" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Me.HfDetalle.Clear
         MsgBox "DOCUMENTO NO REGISTRADO ", vbInformation, KEY_EMPRESA
         Call Resalta(Me.TxtNumero_guia)
         Exit Sub
      
    Else
    idVenta = rst("id_venta")
    Me.txtid_venta.Text = rst("id_venta")
            If MsgBox("ESTA SEGURO DE REALIZAR ESTA OPERACION", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
                strCadena = "SELECT * FROM persona P,movimiento_venta V WHERE P.dni=V.id_cliente AND V.id_venta='" & idVenta & "' AND V.ruc='" & KEY_RUC & "'"
                Call ConfiguraRstT(strCadena)
                If rstT.RecordCount > 0 Then
                    Me.TxtRucDestino.Text = rstT("dni")
                    Me.TxtNombreDestino.Text = BDBuscarCampo("persona", "nombre_completo", "dni", rstT("dni"))
                    Me.TxtDireccionDestino.Text = rstT("direccion")
                    
                End If
               ' Call llenarGrid_Comprobante(Me.HfdDetalle, idVenta)
                Call Llenar_Temporal_transferencias(idVenta)
               ' Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.text, Me.TxtSerie.text, Me.DtcTipoDoc.BoundText)
                Me.TxtCodProducto.Enabled = False
                Me.TxtDescripcionProducto.Enabled = False
                
                Me.cmdagregar.Enabled = False
                Me.CmdQuitar.Enabled = False
                Call Resalta(TxtNumero_guia)
                Referencia = True
                Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
                Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
                
                Me.TxtCodProducto.Enabled = True
                
                Me.cmdagregar.Enabled = True
                Me.CmdQuitar.Enabled = True
                
                
            End If
    End If
End If
Set rst = Nothing

End Sub

Private Sub TxtNumeroDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscar_comprobante
End If
End Sub
Public Sub buscar_comprobante(Optional id_transferencia As Double)
      Me.TxtNumeroDoc.Text = FormatosCeros(Me.TxtNumeroDoc.Text, 6)
    strCadena = "SELECT * FROM parte_material WHERE (numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND ruc='" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    
        Me.TxtId_parte.Text = rst("id_material")
        Me.DtcTipoDoc.BoundText = rst("id_doc")
        Me.DTPicker1.Value = rst("fecha")
        Me.TxtSerie.Text = rst("serie")
        Me.TxtNumeroDoc.Text = rst("numero")
        
        If rst("anulado") = "si" Then
            Me.lblAnulado.Visible = True
            Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
        Else
            Me.lblAnulado.Visible = False
            Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True
        End If
        
        
        
        
        Me.TxtObservacion.Text = rst("observacion")
        If IsNull(rst("id_cliente")) = False And rst("id_cliente") <> "" Then
            Me.TxtRucDestino.Text = rst("id_cliente")
            Me.TxtNombreDestino.Text = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_cliente"))
            Me.TxtDireccionDestino.Text = rst("destino")
            
        Else
            Me.TxtNombreDestino.Text = ""
        End If
        
         If IsNull(rst("id_conductor")) = False And rst("id_conductor") <> "" Then
            Me.TxtIdConductor.Text = rst("id_conductor")
            Me.lblRazonTransporte.Caption = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_conductor"))
           
            
        Else
            Me.TxtNombreDestino.Text = ""
        End If
        
        If IsNull(rst("id_transporte")) = False And rst("id_transporte") <> "" Then
            Me.DtcTransporte.BoundText = rst("id_transporte")
            
          End If
            
            Me.txtHoraInicio.Text = rst("hora_inicio")
            Me.TxtHoraRetorno.Text = rst("hora_retorno")
            Me.txtTiempoRuta.Text = rst("tiempo_ruta")
            Me.TxtId_parte.Text = rst("id_material")
           
            Me.TxtObservacion.Text = rst("observacion")
       
         Call llenar_detalle(Me.HfDetalle, Val(Me.TxtId_parte.Text))
        Me.DtcTipoDoc.Enabled = False
        Me.TxtSerie.Enabled = False
        Me.TxtNumeroDoc.Enabled = False
        Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
        Me.TlbGrabar.Buttons("(Verificar)").Enabled = True
        Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
        Set rst = Nothing
    'Else
   ' Call Resalta(Me.TxtRuc)
    'End If
    End If
End Sub




Private Sub TxtRucDestino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscarResponsable(Trim(Me.TxtRucDestino.Text))
End If
End Sub


Private Sub Llenar_Temporal_transferencias(ByVal idVenta As Double)
Dim total_temp As Double
Dim rstTemporal As New ADODB.Recordset
Dim rstDetalle As New ADODB.Recordset
Dim i As Integer
strCadena = "SELECT * FROM movimiento_venta_detalle D WHERE id_venta='" & idVenta & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' "
CnBd.Execute (strCadena)
 
total_temp = 0
rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
    strCadena = "INSERT INTO movimiento_transferencia_temporal(id_doc,serie,numero,id_producto,cantidad,peso,total,dni_save,ruc) VALUES " & _
    "('" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.TxtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & rst("id_producto") & "','" & rst("cantidad") & "','" & rst("peso") & "'," & _
    "'" & Val(rst("peso")) * Val(rst("cantidad")) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    rst.MoveNext
    Next i
    Call Llenar_Temporal(Me.HfDetalle)
End If

 

End Sub

Public Sub Llenar_Temporal(ByVal Grilla As MSHFlexGrid)
'On Error GoTo salir
Dim tTotal As Double, ccostos As String
strCadena = "SELECT * FROM movimiento_transferencia_temporal T,producto P,unidad U WHERE T.id_producto=P.id_producto AND T.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND T.id_doc='" & Me.DtcTipoDoc.BoundText & "' AND T.serie='" & Me.TxtSerie.Text & "' AND T.numero='" & Me.TxtNumeroDoc.Text & "' AND T.dni_save='" & KEY_USUARIO & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
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
           Grilla.ColWidth(3) = 1300
           Grilla.ColWidth(4) = 1300
           Grilla.ColWidth(5) = 1300
           Grilla.ColWidth(6) = 1300
           Grilla.ColWidth(7) = 1300
       Next
        cabecera = "IDDETALLE" & vbTab & "COD PROD" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "VIAJES" & vbTab & "RECIBIDO" & vbTab & "UNIDAD" & vbTab & "PESO" & vbTab & " TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        Me.lblItems.Caption = str(rst.RecordCount) + Space(2) + "Viajes"
        For i = 0 To rst.RecordCount - 1
          tTotal = tTotal + rst("peso") * rst("cantidad")
          Fila = rst("id_temporal") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("recibido"), "#,##0.00") & vbTab & rst("abreviatura") & vbTab & Format(rst("peso"), "#,##0.00") & vbTab & Format(rst("peso") * rst("cantidad"), "#,##0.00")
          Grilla.AddItem Fila
          If rst("cantidad") <> rst("recibido") Then
          For k = 0 To 7
              Grilla.col = k
              Grilla.Row = i + 1
              Grilla.CellBackColor = &HC0C0FF
          Next k
          End If
          Fila = ""
          rst.MoveNext
      Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "**PESO TOTAL**" & vbTab & Format(tTotal, "###0.00")
        Grilla.AddItem Fila
       
      For k = 6 To 7
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0C0FF
      Next k
      Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
' Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Public Sub llenar_detalle(ByVal Grilla As MSHFlexGrid, ByVal id_transferencia As Double)
'On Error GoTo salir
Dim tTotal As Double, ccostos As String
strCadena = "SELECT * FROM parte_material_detalle T,producto P,unidad U WHERE T.id_producto=P.id_producto AND T.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND T.id_parte='" & id_transferencia & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
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
           Grilla.ColWidth(3) = 1300
           Grilla.ColWidth(4) = 1300
           Grilla.ColWidth(5) = 1300
           Grilla.ColWidth(6) = 1300
           Grilla.ColWidth(7) = 1300
       Next
        cabecera = "IDDETALLE" & vbTab & "COD PROD" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "ENVIADO" & vbTab & "RECIBIDO" & vbTab & "UNIDAD" & vbTab & "PESO" & vbTab & " TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 7
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        Me.lblItems.Caption = str(rst.RecordCount) + Space(2) + "Viajes"
        For i = 0 To rst.RecordCount - 1
          tTotal = tTotal + rst("total")
          Fila = rst("id_detalle") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & rst("abreviatura") & vbTab & Format(rst("peso"), "#,##0.00") & vbTab & Format(rst("total"), "#,##0.00")
          Grilla.AddItem Fila
          If rst("cantidad") <> rst("cantidad") Then
          For k = 0 To 7
              Grilla.col = k
              Grilla.Row = i + 1
              Grilla.CellBackColor = &HC0C0FF
          Next k
          End If
          Fila = ""
          rst.MoveNext
      Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "**PESO TOTAL**" & vbTab & Format(tTotal, "###0.00")
        Grilla.AddItem Fila
       
     
      'Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
      'Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
' Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub TxtSeri_guia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSeri_guia.Text = formato_item(Trim(Me.TxtSeri_guia.Text), 3)
    Call Resalta(Me.TxtNumero_guia)
End If

End Sub

Private Sub TxtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie.Text = formato_item(Me.TxtSerie.Text, 3)
    Call Resalta(Me.TxtNumeroDoc)
End If
End Sub



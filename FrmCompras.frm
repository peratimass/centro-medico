VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCompras_anterior 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Sección Compras"
   ClientHeight    =   9405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   13965
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DtpFechaReferencia 
      Height          =   375
      Left            =   10200
      TabIndex        =   67
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   68812801
      CurrentDate     =   40603
   End
   Begin VB.TextBox TxtPecepcion 
      Height          =   285
      Left            =   13200
      TabIndex        =   64
      Top             =   2520
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CheckBox ChkPercepcion 
      Caption         =   "Percep"
      Height          =   255
      Left            =   12360
      TabIndex        =   63
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CheckBox chkPorcenyaje 
      Caption         =   "Dsto(%)"
      Height          =   255
      Left            =   12360
      TabIndex        =   62
      Top             =   2280
      Width           =   975
   End
   Begin VB.CheckBox chkDstosoles 
      Caption         =   "Dsto(S/.)"
      Height          =   255
      Left            =   12360
      TabIndex        =   61
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox TxtImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   10680
      MaxLength       =   80
      TabIndex        =   57
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CheckBox chkIgv 
      BackColor       =   &H00DFDFE0&
      Caption         =   "IGV"
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
      Left            =   9960
      TabIndex        =   56
      Top             =   7680
      Width           =   735
   End
   Begin VB.TextBox TxtCosto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   7940
      MaxLength       =   80
      TabIndex        =   54
      Top             =   7800
      Width           =   855
   End
   Begin VB.TextBox TxtDescuento 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   8820
      MaxLength       =   80
      TabIndex        =   52
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox TxtUnidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   7070
      MaxLength       =   80
      TabIndex        =   48
      Top             =   7800
      Width           =   855
   End
   Begin VB.CheckBox ChkRef 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDFE0&
      Caption         =   "Ref:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7080
      TabIndex        =   46
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      MaxLength       =   80
      TabIndex        =   45
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox TxtDireccion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      MaxLength       =   80
      TabIndex        =   44
      Top             =   1860
      Width           =   4335
   End
   Begin VB.CommandButton CmdQuitar 
      BackColor       =   &H8000000D&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7785
      Width           =   375
   End
   Begin VB.CommandButton CmdAgregar 
      BackColor       =   &H8000000D&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7785
      Width           =   375
   End
   Begin VB.TextBox TxtNumero_Ref 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   10605
      MaxLength       =   80
      TabIndex        =   35
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox TxtSerie_Ref 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   9840
      MaxLength       =   80
      TabIndex        =   34
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   11880
      MaxLength       =   80
      TabIndex        =   33
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox TxtCodProducto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   240
      MaxLength       =   80
      TabIndex        =   31
      Top             =   7800
      Width           =   1455
   End
   Begin VB.TextBox TxtDescripcionProducto 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2720
      MaxLength       =   80
      TabIndex        =   30
      Top             =   7800
      Width           =   4335
   End
   Begin VB.TextBox TxtCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1720
      MaxLength       =   80
      TabIndex        =   29
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox TxtObservacion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   1800
      MaxLength       =   80
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   2640
      Width           =   4335
   End
   Begin VB.TextBox TxtSerie 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   7560
      MaxLength       =   80
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox TxtNumeroDoc 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   8685
      MaxLength       =   80
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox TxtCodProveedor 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   360
      MaxLength       =   80
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox TxtProveedor 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      MaxLength       =   80
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   420
      Left            =   330
      TabIndex        =   4
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   420
      Left            =   7530
      TabIndex        =   5
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   5040
      Top             =   -240
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
            Picture         =   "FrmCompras.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones1 
      Height          =   3945
      Left            =   12960
      TabIndex        =   6
      Top             =   3480
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   6959
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   3945
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
         TabIndex        =   7
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1482
         ButtonHeight    =   1376
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
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
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
      Left            =   240
      Top             =   5520
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
            Picture         =   "FrmCompras.frx":1FBC
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCompras.frx":2896
            Key             =   "(Imprimir)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   240
      TabIndex        =   8
      Top             =   8280
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
      Begin MSComctlLib.Toolbar TlbGrabar 
         Height          =   780
         Left            =   0
         TabIndex        =   9
         Top             =   30
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1376
         ButtonWidth     =   1191
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImageList1"
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
               Caption         =   "Imprimir"
               Key             =   "(Imprimir)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Imprimir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComCtl2.DTPicker DtpActual 
      Height          =   375
      Left            =   9405
      TabIndex        =   25
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   0
      CalendarForeColor=   255
      Format          =   68812801
      CurrentDate     =   39535
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   630
      Left            =   4680
      TabIndex        =   28
      Top             =   240
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      ImageList       =   "ImgIconos"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "(Grabar)"
            Object.ToolTipText     =   "Grabar Ctrl+G"
            ImageKey        =   "(Aceptar)"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSComCtl2.DTPicker DTPDetracion 
      Height          =   375
      Left            =   11880
      TabIndex        =   36
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   0
      CalendarForeColor=   255
      Format          =   68812801
      CurrentDate     =   39535
   End
   Begin MSDataListLib.DataCombo DtcTipoDoc_Ref 
      Height          =   315
      Left            =   8445
      TabIndex        =   37
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   3735
      Left            =   120
      TabIndex        =   39
      Top             =   3480
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   6588
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      ForeColorSel    =   16777215
      GridColor       =   -2147483635
      SelectionMode   =   1
      Appearance      =   0
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
   Begin MSDataListLib.DataCombo DtcFormaPago 
      Height          =   360
      Left            =   8400
      TabIndex        =   42
      Top             =   2640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblPercepcion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   9240
      TabIndex        =   66
      Top             =   8595
      Width           =   1365
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERCEPC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9435
      TabIndex        =   65
      Top             =   8400
      Width           =   885
   End
   Begin VB.Label lblIMPBruto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2760
      TabIndex        =   60
      Top             =   8595
      Width           =   1605
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IM.BRUTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3150
      TabIndex        =   59
      Top             =   8400
      Width           =   945
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   10800
      TabIndex        =   58
      Top             =   7560
      Width           =   885
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COSTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8040
      TabIndex        =   55
      Top             =   7560
      Width           =   675
   End
   Begin VB.Label lblPorcentaje 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCT(S/)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8850
      TabIndex        =   53
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7230
      TabIndex        =   51
      Top             =   7560
      Width           =   435
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3450
      TabIndex        =   50
      Top             =   7560
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1800
      TabIndex        =   49
      Top             =   7560
      Width           =   525
   End
   Begin VB.Label lblAnulado 
      BackStyle       =   0  'Transparent
      Caption         =   "---ANULADO---"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2160
      TabIndex        =   47
      Top             =   960
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape Shape7 
      FillStyle       =   5  'Downward Diagonal
      Height          =   555
      Left            =   8280
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Label LblComprobante_DR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forma Pago:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   7080
      TabIndex        =   43
      Top             =   2760
      Width           =   1155
   End
   Begin VB.Shape Shape6 
      FillStyle       =   5  'Downward Diagonal
      Height          =   495
      Left            =   8280
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detracciòn"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   12045
      TabIndex        =   38
      Top             =   300
      Width           =   1005
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
      Left            =   13125
      TabIndex        =   32
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   495
      TabIndex        =   27
      Top             =   7560
      Width           =   765
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Obs:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1050
      TabIndex        =   24
      Top             =   2760
      Width           =   435
   End
   Begin VB.Label LblDescripcion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serie:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   7755
      TabIndex        =   23
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   8940
      TabIndex        =   22
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   8640
      TabIndex        =   21
      Top             =   1320
      Width           =   645
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos Proveedor:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   345
      TabIndex        =   20
      Top             =   1140
      Width           =   1605
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   570
      TabIndex        =   19
      Top             =   1920
      Width           =   915
   End
   Begin VB.Label LblIdentificacion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruc:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1050
      TabIndex        =   18
      Top             =   2280
      Width           =   435
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCUENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4725
      TabIndex        =   17
      Top             =   8385
      Width           =   1185
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR VENTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6225
      TabIndex        =   16
      Top             =   8385
      Width           =   1335
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8370
      TabIndex        =   15
      Top             =   8385
      Width           =   345
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   11205
      TabIndex        =   14
      Top             =   8385
      Width           =   645
   End
   Begin VB.Label lblDescuento 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4455
      TabIndex        =   13
      Top             =   8595
      Width           =   1605
   End
   Begin VB.Label lblValorVenta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6135
      TabIndex        =   12
      Top             =   8595
      Width           =   1605
   End
   Begin VB.Label lblIgv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   7815
      TabIndex        =   11
      Top             =   8595
      Width           =   1365
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   10695
      TabIndex        =   10
      Top             =   8595
      Width           =   1725
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   900
      Left            =   120
      Top             =   120
      Width           =   6615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   2220
      Left            =   120
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   120
      Top             =   7440
      Width           =   13695
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   2640
      Top             =   8280
      Width           =   9975
   End
   Begin VB.Shape ShapeDR 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   1470
      Left            =   6960
      Top             =   1800
      Width           =   6855
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   1620
      Left            =   11280
      Top             =   120
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   1620
      Left            =   6960
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FrmCompras_anterior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doc_Tienda As String * 1
Dim cod_doc As String
Dim RstDetCompra As New ADODB.Recordset
Dim rstTemporal As New ADODB.Recordset
Dim StrCodDetCompra As String * 20
Dim StrCodReferencia As Double
Dim Referencia As Boolean
Public Procedencia As EnumProcede
Public ProcendenciaGuia As EnumGuia
Public ProcedenciaFactura As EnumFactura
Public codigoP As String

Private Sub AgregarGrilla()
If Val(Me.TxtCantidad.Text) > 0 And Trim(Me.TxtCodProducto.Text) <> "" Then
    Call AgregarTemporal
    Me.TxtCodProducto.Text = "0000"
    Me.TxtDescripcionProducto.Text = ""
    Me.TxtCantidad.Text = "0"
     
    Call Resalta(Me.TxtCodProducto)
Else
    Call Resalta(Me.TxtCantidad)
End If
End Sub
Private Sub AgregarTemporal()
    strCadena = "SELECT cProducto,Cantidad FROM Temporal_Compras WHERE cProducto='" & Trim(Me.TxtCodProducto.Text) & "' AND cDocumentoCompra='" & Trim(Me.TxtNumeroDoc.Text) & "'"
    Call ConfiguraRst(strCadena)
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
    
    'If Rst.RecordCount > 0 Then
     '   Call ModificarCantidad(Rst(1))
    'Else
        Call AgregarNuevo
    'End If
    
End Sub
Sub AgregarNuevo()
    Dim imp_uni As Single
    Dim imp_bruto As Double
    Dim desct As Double
    Dim valor_venta As Double
    Dim igv_c As Double
    Dim importe As Double
    Dim cantidad As Double
    
    If Me.chkDstosoles.Value = 1 Then
        cantidad = Me.TxtCantidad.Text
        importe = Me.TxtImporte.Text
        desct = Me.TxtDescuento.Text
        If Me.chkigv.Value = 1 Then
            igv_c = importe - (importe / 1.18)
            valor_venta = importe / 1.18
        Else
            igv_c = 0#
             valor_venta = importe
        End If
        
        imp_bruto = valor_venta + desct
        imp_uni = imp_bruto / cantidad
        GoTo llenar
    ElseIf Me.chkPorcenyaje.Value = 0 Then
        cantidad = Me.TxtCantidad.Text
        importe = Me.TxtImporte.Text
        desct = 0#
        If Me.chkigv.Value = 1 Then
            igv_c = importe - (importe / 1.18)
            valor_venta = importe / 1.18
        Else
            igv_c = 0#
             valor_venta = importe
        End If
        imp_bruto = valor_venta + desct
        imp_uni = imp_bruto / cantidad
        GoTo llenar
    End If
    
    If Me.chkPorcenyaje.Value = 1 Then
        cantidad = Me.TxtCantidad.Text
        importe = Me.TxtImporte.Text
        desct = Me.TxtDescuento.Text / 100
        If Me.chkigv.Value = 1 Then
            igv_c = importe - (importe / 1.18)
            valor_venta = importe / 1.18
        Else
            igv_c = 0#
             valor_venta = importe
        End If
        imp_bruto = valor_venta + desct
        imp_uni = imp_bruto / cantidad
        GoTo llenar
    Else
        cantidad = Me.TxtCantidad.Text
        importe = Me.TxtImporte.Text
        desct = 0#
        If Me.chkigv.Value = 1 Then
            igv_c = importe - (importe / 1.18)
            valor_venta = importe / 1.18
        Else
            igv_c = 0#
             valor_venta = importe
        End If
        imp_bruto = valor_venta + desct
        imp_uni = imp_bruto / cantidad
        GoTo llenar
    End If
       
llenar:
    strCadena = "SELECT cTemporal FROM Temporal_Compras WHERE id_usuario='" & KEY_USUARIO & "' ORDER BY cTemporal DESC "
    Set rst = Nothing
    
    strCadena = "INSERT INTO Temporal_Compras VALUES ('" & GeneraCodTemporalCompras & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.DtcTipoDoc.BoundText) & "'," & _
    "'" & Trim(Me.TxtSerie.Text) & "','" & Trim(codigoP) & "','" & Val(Me.TxtCantidad.Text) & "','" & imp_uni & "','" & imp_bruto & "','" & desct & "'," & _
    "'" & valor_venta & "','" & igv_c & "','" & importe & "','" & KEY_USUARIO & "')"
    
    
    '" & _
    "('" & GeneraCodTemporalCompras & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Trim(Me.TxtSerie.Text) & "','" & Trim(Me.TxtCodProducto.Text) & "','" & Val(Me.TxtCantidad.Text) & "'," & _
    "'" & Val(Me.TxtPrecio.Text) & "','" & Val(Me.LblTotalParcial.Text) & "')"
    Call EjecutaRST(strCadena)
    Set rst = Nothing
    Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcTipoDoc.BoundText, Me.TxtSerie.Text)
    Call VerificaDocumento(Trim(Me.DtcTipoDoc.BoundText))
End Sub
Private Sub VerificaDocumento(ByVal TipoDoc As String)
If Trim(Me.DtcTipoDoc.BoundText) = "0009" Then
    Me.TlbGrabar.Buttons(KEY_GUIAREMISION).Enabled = True
End If
End Sub
Sub ModificarCantidad(ByVal Can_Previa As Integer)
    Dim Can_actual As Integer
'    Can_actual = Can_Previa + Val(Me.TxtCantidad.Text)
 '   StrCadena = "UPDATE Temporal_Compras SET cantidad='" & Can_actual & "',Total='" & Can_actual * Val(Me.TxtPrecio.Text) & "' WHERE cProducto='" & Trim(Me.TxtCodProducto.Text) & "' "
  '  Call EjecutaRST(StrCadena)
   ' Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcTipoDoc.BoundText, Me.TxtSerie.Text)
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
Function GeneraCodTemporalCompras() As Integer
Dim Codtemporal As Integer
Call ConfiguraRst(strCadena)
    If rst.EOF = True Then
        Codtemporal = 1
    Else
        Codtemporal = rst(0) + 1
    End If
  GeneraCodTemporalCompras = Codtemporal
  Set rst = Nothing
End Function
Function GeneraCodReferencia() As Integer
Dim CodReferencia As Integer
strCadena = "SELECT IdReferencia FROM DocReferencia_Compra ORDER BY IdReferencia DESC "
Call ConfiguraRst(strCadena)
    If rst.EOF = True Then
        CodReferencia = 1
        
    Else
        CodReferencia = rst(0) + 1

    End If
  GeneraCodReferencia = CodReferencia
  
  
  Set rst = Nothing
End Function

Private Sub SaveReferencia(ByVal Codigo As String, ByVal TipoDoc As String, ByVal serie As String, ByVal Numero As String, ByVal fecha As Date, ByVal Almacen As String)
strCadena = "INSERT INTO DocReferencia_Compra(IdReferencia,doc_cod,sSerie,cDocumentoCompra,FechaProceso,Alm_Cod,des_doc) VALUES " & _
            "('" & Codigo & "','" & TipoDoc & "','" & serie & "','" & Numero & "','" & fecha & "','" & Almacen & "','" & Trim(Me.DtcTipoDoc_Ref.Text) & "')"
            Call EjecutaRST(strCadena)
            Set RstEjecuta = Nothing
End Sub
Sub llenarGrid_det(ByVal Grilla As MSHFlexGrid, ByVal Numero As String, ByVal TipoDoc As String, ByVal serie As String)
Dim total As Double
Dim SUBTOTAL As Double
Dim igv As Single

strCadena = "SELECT Temporal_Compras.cProducto as CODIGO, Producto.DescripcionProducto AS DESCRIPCION, Unidad.sAbreviatura AS UND, Temporal_Compras.Cantidad AS CANT, " & _
"Temporal_Compras.imp_uni AS IMP_UNI, Temporal_Compras.imp_bruto AS IMP_BRUTO, Temporal_Compras.desct AS DESCT, Temporal_Compras.valor_vta AS VALOR_VTA,Temporal_Compras.igv AS IGV ," & _
" Temporal_Compras.TOTAL AS TOTAL FROM Producto INNER JOIN Unidad ON Producto.cUnidad = Unidad.cUnidad INNER JOIN " & _
"Temporal_Compras ON Producto.cProducto = Temporal_Compras.cProducto WHERE (Temporal_Compras.cDocumentoCompra='" & Numero & "' AND Temporal_Compras.doc_cod='" & TipoDoc & "' AND Temporal_Compras.sSerie='" & serie & "')"
                      
On Error GoTo salir
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.Row = 0
  Grilla.Rows = 2
  Set Grilla.Recordset = rst
  Grilla.Rows = rst.RecordCount + 1
  Grilla.ColWidth(0) = 1500
  Grilla.ColWidth(1) = 3500
  Grilla.ColWidth(2) = 500
  Grilla.ColWidth(3) = 1000
  Grilla.ColWidth(4) = 1000
  Grilla.ColWidth(5) = 1000
  Grilla.ColWidth(6) = 1000
  Grilla.ColWidth(7) = 1000
  Grilla.ColWidth(8) = 1000
  Grilla.ColWidth(9) = 1000
Me.LblCantidad.Caption = Trim(rst.RecordCount)
Set rst = Nothing
Call Resalta(Me.TxtCodProducto)
Me.TxtCantidad.Text = 0
Me.TxtDescuento.Text = 0
Me.TxtImporte.Text = 0
Dim tpercepcion As Single
strCadena = "SELECT sum(imp_bruto),sum(desct),sum(valor_vta),sum(igv),sum(total) FROM Temporal_Compras WHERE (cDocumentoCompra='" & Trim(Me.TxtNumeroDoc.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND id_usuario='" & KEY_USUARIO & "')"
Call ConfiguraRst(strCadena)

If Me.ChkPercepcion.Value = 1 Then
    Me.LblPercepcion.Caption = Format(Me.TxtPecepcion.Text, "#,##0.00")
    tpercepcion = Me.LblPercepcion.Caption
Else
    tpercepcion = 0
    Me.LblPercepcion.Caption = "0.00"
End If

Me.lblIMPBruto.Caption = Format(rst(0), "#,##0.000")
Me.lblDescuento.Caption = Format(rst(1), "#,##0.000")
Me.lblValorVenta.Caption = Format(rst(2), "#,##0.000")
Me.lblIgv.Caption = Format(rst(3), "#,##0.000")
Me.lblTotal.Caption = Format(rst(4) + tpercepcion, "#,##0.000")


    
    Set rst = Nothing
  Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
  Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  Exit Sub

salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Sub llenarGrid_Comprobante(ByVal Grilla As MSHFlexGrid, ByVal NumeroCompra As String, ByVal TipoDocumento As String, ByVal serie As String)
Dim total As Double
Dim SUBTOTAL As Double
Dim igv As Single
Dim Items As Integer
strCadena = "SELECT Detalle_DocumentoCompra.cProducto, Producto.DescripcionProducto, Unidad.sAbreviatura as Und, Detalle_DocumentoCompra.cantidad, " & _
"Detalle_DocumentoCompra.imp_uni, Detalle_DocumentoCompra.imp_bruto, Detalle_DocumentoCompra.desct," & _
"Detalle_DocumentoCompra.valor_vta , Detalle_DocumentoCompra.igv, Detalle_DocumentoCompra.Total " & _
"FROM  DocumentoCompra INNER JOIN Detalle_DocumentoCompra ON DocumentoCompra.cDocumentoCompra = Detalle_DocumentoCompra.cDocumentoCompra AND " & _
"DocumentoCompra.Alm_cod = Detalle_DocumentoCompra.Alm_cod AND DocumentoCompra.doc_cod = Detalle_DocumentoCompra.doc_cod AND " & _
"DocumentoCompra.sSerie = Detalle_DocumentoCompra.sSerie INNER JOIN Producto ON Detalle_DocumentoCompra.cProducto = Producto.cProducto INNER JOIN " & _
"Unidad ON Producto.cUnidad = Unidad.cUnidad WHERE (Detalle_DocumentoCompra.cDocumentoCompra='" & Trim(NumeroCompra) & "' AND Detalle_DocumentoCompra.doc_cod='" & Trim(TipoDocumento) & "' AND Detalle_DocumentoCompra.sSerie='" & Trim(serie) & "')"
On Error GoTo salir
  Call ConfiguraRst(strCadena)
  Items = rst.RecordCount
  'Grilla.Clear
  'Grilla.Row = 0
  'Grilla.Rows = Rst.RecordCount
  Set Grilla.Recordset = rst
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 1500
  Grilla.ColWidth(1) = 3500
  Grilla.ColWidth(2) = 500
  Grilla.ColWidth(3) = 1000
  Grilla.ColWidth(4) = 1000
  Grilla.ColWidth(5) = 1000
  Grilla.ColWidth(6) = 1000
  Grilla.ColWidth(7) = 1000
  Grilla.ColWidth(8) = 1000
  Grilla.ColWidth(9) = 1000
Call DarFormato(Grilla, 3)
Call DarFormato(Grilla, 4)
Call DarFormato(Grilla, 5)
Call DarFormato(Grilla, 6)
Call DarFormato(Grilla, 7)
Call DarFormato(Grilla, 8)
Call DarFormato(Grilla, 9)
Me.LblCantidad.Caption = Trim(rst.RecordCount)
Set rst = Nothing
strCadena = "SELECT SUM(Total)FROM Detalle_DocumentoCompra WHERE (cDocumentoCompra='" & Trim(NumeroCompra) & "' AND doc_cod='" & Trim(TipoDocumento) & "' AND sSerie='" & Trim(serie) & "')"
Call ConfiguraRst(strCadena)
total = rst(0)
Me.lblTotal.Caption = Format(rst(0), "#,##0.00")
Set rst = Nothing
    strCadena = "SELECT Igv FROM parametros"
    Call ConfiguraRst(strCadena)
    igv = total * rst(0)
    SUBTOTAL = total - igv
    Me.lblIgv.Caption = Format(igv, "#,##0.00")
   ' Me.LblValorCompra.Caption = Format(SUBTOTAL, "#,##0.00")
   ' Me.LblPercepcion.Caption = "0.00"
    Set rst = Nothing
  Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
  Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub Command1_Click()

End Sub




Private Sub chkDstosoles_Click()
If Me.chkDstosoles.Value = 1 Then
    Me.lblPorcentaje.Caption = "Desct(S/)"
    Me.TxtDescuento.Enabled = True
    
      Else
       If Me.chkPorcenyaje.Value = 1 Then
            Me.LblPercepcion.Caption = "Desct(%)"
            Me.TxtDescuento.Text = "0.00"
            Me.TxtDescuento.Enabled = True
            Exit Sub
       End If
       Me.lblPorcentaje.Caption = "Desct"
       Me.TxtDescuento.Text = "0.00"
       Me.TxtDescuento.Enabled = False
     End If


End Sub

Private Sub ChkPercepcion_Click()
If Me.ChkPercepcion.Value = 1 Then
    Me.TxtPecepcion.Visible = True
    Call Resalta(Me.TxtPecepcion)
    
Else
    Me.TxtPecepcion.Visible = False
End If
End Sub

Private Sub chkPorcenyaje_Click()
If Me.chkPorcenyaje.Value = 1 Then
    Me.lblPorcentaje.Caption = "DESCT(%)"
    Me.TxtDescuento.Enabled = True
Else
        If Me.chkDstosoles.Value = 1 Then
            Me.lblPorcentaje.Caption = "Desct(S/)"
            Me.TxtDescuento.Text = "0.00"
            Me.TxtDescuento.Enabled = True
            Exit Sub
       End If
       
       Me.LblPercepcion.Caption = "Desct"
       Me.TxtDescuento.Text = "0.00"
       Me.TxtDescuento.Enabled = False
   
End If
End Sub

Private Sub ChkRef_Click()

If Me.ChkRef.Value = 1 Then
    
    Me.DtcTipoDoc_Ref.Enabled = True
    Me.TxtSerie_Ref.Enabled = True
    Me.TxtNumero_Ref.Enabled = True
Else
    Me.DtcTipoDoc_Ref.Enabled = False
    Me.TxtSerie_Ref.Enabled = False
    Me.TxtNumero_Ref.Enabled = False
    Referencia = False
End If
End Sub

Private Sub CmdAgregar_Click()
strCadena = "SELECT cDocumentoCompra FROM Temporal_Compras WHERE (cDocumentoCompra='" & Trim(Me.TxtNumeroDoc.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 100 Then
    Call AgregarGrilla
Else
    MsgBox "Cantidad de Articulos Excede al Comprobante", vbInformation
    Exit Sub
End If
End Sub

Private Sub CmdQuitar_Click()
Me.HfdDetalle.Col = 0
Call Quitar(Me.HfdDetalle.Text)
End Sub
Private Sub Quitar(ByVal Codigo As String)
strCadena = "DELETE FROM Temporal_Compras WHERE cProducto='" & Trim(Codigo) & "' AND cDocumentoCompra='" & Trim(Me.TxtNumeroDoc) & "'"
Call EjecutaRST(strCadena)
Set rst = Nothing
Call llenarGrid_det(Me.HfdDetalle, Trim(Me.TxtNumeroDoc.Text), Trim(Me.DtcTipoDoc.BoundText), Trim(Me.TxtSerie.Text))
End Sub
Private Sub DtcAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcTipoDoc.SetFocus
End If
End Sub
Private Sub DtcFormaPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie_Ref.SetFocus
End If
End Sub

Private Sub DtcTipoDoc_Click(Area As Integer)
doc_cod = Me.DtcTipoDoc.BoundText
Call VerificaIdentificacion(doc_cod)

End Sub
Private Sub VerificaIdentificacion(ByVal TipoDoc As String)
If TipoDoc = "0003" Then
    Me.LblIdentificacion.Caption = "DNI:"
Else
    Me.LblIdentificacion.Caption = "Ruc:"
End If
End Sub
Private Sub DtcTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Me.DtcAlmacen.SetFocus
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.TxtSerie)
End If
End Sub

Private Sub DtcTipoDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.TxtSerie.Text <> "" Then
        Call Resalta(Me.TxtSerie)
    Else
        Me.TxtSerie.SetFocus
    End If
End If
End Sub


Private Sub DtcTipoDoc_Ref_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtSerie_Ref)
End If
End Sub

Private Sub Form_Activate()
CenterForm Me
Me.Top = 100
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If Shift = 2 And KeyCode = Asc("G") Then
    If Int(Me.lblTotal.Caption) > 1 Then
        Call Save
        Exit Sub
    Else
        If MsgBox("Esta Intentado Grabar una Factura con Monto CERO" + Chr(13) + "Desea Continuar", vbQuestion + vbYesNo) = vbYes Then
            Call Save
            Exit Sub
        End If
  End If
End If
End Sub
Private Sub Form_Load()
Me.chkDstosoles.Value = 1
Me.chkigv.Value = 1
doc_Tienda = "V"
Me.DtpFechaReferencia.Value = Date
Me.DtpActual.Value = Date
 strCadena = "SELECT Alm_cod as Codigo, Alm_des as Descripcion FROM Almacen " & _
  " ORDER BY Alm_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = "0001"
  Me.DtcAlmacen.Enabled = False
  Set rst = Nothing
  
  strCadena = "SELECT doc_cod as Codigo, doc_abrev as Descripcion FROM Comprobantes " & _
  " WHERE doc_tienda='V' ORDER BY doc_abrev"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
  Me.DtcTipoDoc.BoundText = "0089"
  Set rst = Nothing
  
  strCadena = "SELECT doc_cod as Codigo, doc_abrev as Descripcion FROM Comprobantes " & _
  " WHERE doc_tienda='" & doc_Tienda & "' ORDER BY doc_abrev"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc_Ref)
  Set rst = Nothing
  
  
strCadena = "SELECT idFormaPago as Codigo,sFormapago as Descripcion FROM FormaPago ORDER BY sFormaPago "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFormaPago)
Me.DtcFormaPago.BoundText = "0001"
Set rst = Nothing

  
  Me.DtcTipoDoc.Enabled = False
  Me.TxtSerie.Enabled = False
  Me.TxtNumeroDoc.Enabled = False
  Me.ChkRef.Value = 1
  Referencia = True
  Me.TlbAcciones.Buttons(KEY_EXIT).Enabled = True
  
  Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
  Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
  
End Sub



Private Sub LblTotalParcial_KeyPress(KeyAscii As Integer)
Dim PrecioCompra As Double
Dim TotalParcial As Double
If KeyAscii = 13 Then
   ' TotalParcial = Me.LblTotalParcial.Text
    PrecioCompra = TotalParcial / Val(Me.TxtCantidad.Text)
    'Me.TxtPrecio.Text = Format(PrecioCompra, "#,##0.00")
    Me.cmdAgregar.SetFocus

End If
End Sub







Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_NEW
        Call Nuevo
    Case KEY_UPDATE
    Case KEY_ANULAR
         If MsgBox(MSGANULAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
            Procedencia = anular
            FrmSeguridad.Show
        End If
    Case KEY_DELETE
      If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Procedencia = Eliminar
        FrmSeguridad.Show
     End If
    Case KEY_EXIT
      Unload Me
  End Select
End Sub
Public Sub Nuevo()

     strCadena = "DELETE FROM Temporal_Compras"
     Call EjecutaRST(strCadena)
     Set RstEjecuta = Nothing
    strCadena = "SELECT cDocumentoCompra FROM DocumentoCompra WHERE (sSerie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "') ORDER BY intDocumentoCompra DESC"
    Call ConfiguraRst(strCadena)
    Me.TxtNumeroDoc.Text = GeneraCodigo(10)
    Me.TxtCodProveedor.Text = ""
    Me.TxtProveedor.Text = ""
    Me.TxtDireccion.Text = ""
    Me.TxtRuc.Text = ""
    Me.TxtObservacion.Text = ""
    Me.TxtDescripcionProducto.Text = ""
    Me.chkigv.Value = 1
    Me.chkDstosoles.Value = 1
    Me.chkPorcenyaje.Value = 0
    Me.ChkPercepcion.Value = 0
    Me.TxtCodProducto = "0000"
    Me.TxtDescuento.Enabled = True
    Me.TxtDescripcionProducto.Text = ""

    Me.TxtCantidad.Text = 0
    Me.TxtImporte.Text = 0#
    Me.LblCantidad.Caption = "0"
    Me.TxtCodProducto.Enabled = True
    Me.TxtDescripcionProducto.Enabled = True
    Me.TxtCantidad.Enabled = True
    Me.DtcFormaPago.BoundText = "0001"
    Me.cmdAgregar.Enabled = True
    Me.cmdQuitar.Enabled = True
    Me.DtcTipoDoc_Ref.Text = ""
    Me.TxtSerie_Ref.Text = ""
    Me.TxtNumero_Ref.Text = ""
    Me.lblAnulado.Visible = False
    
     Referencia = True
    
    Me.DtpFechaReferencia.Value = Date
    Me.DtcTipoDoc_Ref.BoundText = "0001"
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
    
    Me.HfdDetalle.Clear
    Me.DtcTipoDoc.SetFocus
    Me.DtpActual.Value = Date
    
    
    Set rst = Nothing
End Sub

Sub verifica(ByVal doc_deta As String)
    Select Case Val(doc_deta)
        Case 1
           ' Call Doc_Referencia(True, Val(doc_deta))
        Case 3
           ' Call Doc_Referencia(False, Val(doc_deta))
        Case 7
           ' Call Doc_Referencia(True, Val(doc_deta))
        Case 8
            'Call Doc_Referencia(True, Val(doc_deta))
        Case 9
            
            'Call Doc_Referencia(True, Val(doc_deta))
        Case 88
            'Call Doc_Referencia(False, Val(doc_deta))
        Case 89
            'Call Doc_Referencia(False, Val(doc_deta))
        Case 90
            'Call Doc_Referencia(False, Val(doc_deta))
        Case 95
            'Call Doc_Referencia(False, Val(doc_deta))
        Case 96
            'Call Doc_Referencia(True, Val(doc_deta))
    End Select
    
End Sub



Private Sub TlbAgregar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_AGREGAR
        Call AgregarGrilla
    Case KEY_QUITAR
        'Call Quitar
    
  End Select
End Sub

Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.Key
    Case KEY_SAVE
      Call Save
    Case KEY_PRINT
      Call Imprimir
    Case KEY_GUIAREMISION
      FrmDetalleGuia.Show
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
Exit Sub
End Sub
Private Function CodigoKardex() As String
strCadena = "SELECT int_Kardex FROM Kardex ORDER BY int_kardex DESC"
    Call ConfiguraRst(strCadena)
    CodigoKardex = GeneraCodigo(20)
    Set rst = Nothing
End Function
Private Sub Imprimir()
If Me.DtcTipoDoc.BoundText = KEY_INGALMA Then
Dim i As Integer, j As Integer
Dim laVenta, espacios
Dim MES As String
Dim Ans As Boolean
Dim cantidad As String, Und As String, descripcion As String, precio As String
Dim total As String, SUBTOTAL As String, igv As String
Dim totalPar As String
Dim Descuento As String
Dim GranTotal As String
Dim totalletras As String
Dim Peso As Double
Dim inc As Single
Dim Codigo As String, Unidad As String, PesoTotal As Double
Dim Toneladas As String

    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    Printer.Print Tab(20); "PROVEEDOR:" + Space(1) + Me.TxtProveedor.Text
    Printer.Print Tab(20); "DIRECCION:" + Space(1) + Me.TxtDireccion.Text
    Printer.Print Tab(20); "RUC      :" + Space(1) + Me.TxtRuc.Text
    Printer.Print Tab(20); "FACTURA"; Space(2); Mid(Me.TxtSerie_Ref.Text + Space(50), 1, 4) & Space(2) & "-" & Space(2) & Me.TxtNumero_Ref.Text + Space(6) + "INGALMA"; Space(2); Mid(Me.TxtSerie.Text + Space(50), 1, 4) & Space(2) & "-" & Space(2) & Me.TxtNumeroDoc.Text
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
strCadena = "SELECT Detalle_DocumentoCompra.cProducto, Detalle_DocumentoCompra.cantidad, Unidad.sAbreviatura, Producto.DescripcionProducto, " & _
            "Detalle_DocumentoCompra.precio , Detalle_DocumentoCompra.TOTAL, DocumentoCompra.nTotalCompra FROM DocumentoCompra INNER JOIN " & _
            "Detalle_DocumentoCompra ON DocumentoCompra.cDocumentoCompra = Detalle_DocumentoCompra.cDocumentoCompra AND " & _
            "DocumentoCompra.Alm_cod = Detalle_DocumentoCompra.Alm_cod AND " & _
            "DocumentoCompra.doc_cod = Detalle_DocumentoCompra.doc_cod AND " & _
            "DocumentoCompra.sSerie = Detalle_DocumentoCompra.sSerie INNER JOIN " & _
            "Producto ON Detalle_DocumentoCompra.cProducto = Producto.cProducto INNER JOIN " & _
            "Unidad ON Producto.cUnidad = Unidad.cUnidad " & _
            "WHERE (Detalle_DocumentoCompra.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND Detalle_DocumentoCompra.sSerie='" & Trim(Me.TxtSerie.Text) & "' " & _
            "AND Detalle_DocumentoCompra.cDocumentoCompra='" & Trim(Me.TxtNumeroDoc.Text) & "')"
    Call ConfiguraRst(strCadena)
    rst.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.2
            For j = 0 To rst.RecordCount - 1
                Codigo = Mid(rst(0) + Space(50), 1, 4)
                cantidad = Mid(Str(rst(1)) + Space(10), 1, 4)
                Und = rst(2)
                descripcion = Mid(rst(3) + Space(80), 1, 48)
                precio = Mid(Format(Str(rst(4)), "#,##0.00") + Space(4), 1, 8)
                totalPar = Mid(Format(Str(rst(5)), "#,##0.00") + Space(4), 1, 8)
                'Printer.Print Tab(2); Codigo & Space(7) & Cantidad & Space(1) & Und & Space(6) & descripcion & precio & Space(4) & totalPar
                Printer.Print Tab(5); Codigo & Space(2) & cantidad & Space(1) & Und & Space(2) & descripcion & precio & Space(4) & totalPar
                Printer.CurrentY = Printer.CurrentY + 0.2
                rst.MoveNext
            Next j
            inc = 0.5
            Do While (Val(Printer.CurrentY) <= 19)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    rst.MoveFirst
    total = Format(Str(rst(6)), "#,##0.00")
    Descuento = Format(Str(KEY_DSCTO), "#,##0.00")
    totalletras = UCase(EnLetras(total))
    Set rst = Nothing
    '---- fin totales
    Printer.CurrentY = Printer.CurrentY + 0.8
    Printer.Print Tab(10); Mid(totalletras + Space(100), 1, 100)
    Printer.Print Tab(55); Mid(total & Space(20), 1, 13) & Descuento & Space(15) & total
    Printer.EndDoc
    
    Exit Sub
End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Me.TlbAcciones.Enabled = True
Me.DtcAlmacen.Enabled = True
Me.DtcTipoDoc.Enabled = True
Me.TxtSerie.Enabled = True
Me.TxtSerie.Text = "0001"
Me.TxtNumeroDoc.Enabled = True
Call Nuevo
Me.DtcAlmacen.SetFocus
End Sub


Private Sub TxtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtCodProducto)
End If

End Sub
Private Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
Dim TotalP As Single
If KeyAscii = 13 And Val(Me.TxtCantidad.Text) > 0 Then
    If Me.TxtDescuento.Enabled = False Then
        Me.TxtDescuento.Text = 0#
        Call Resalta(Me.TxtImporte)
        Exit Sub
    End If
    Call Resalta(Me.TxtDescuento)

End If
End Sub

Private Sub TxtDescuento_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Val(Me.TxtDescuento.Text) = 0) Then
        Me.TxtDescuento.Text = 0
    End If
    Call Resalta(Me.TxtImporte)
End If

End Sub

Private Sub TxtImporte_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If Val(Me.TxtImporte.Text) < 1 Then
        If MsgBox("Esta Ingresando un Importe CERO" + Chr(13) + "Desea Coninuar", vbQuestion + vbYesNo) = vbYes Then
            Me.cmdAgregar.SetFocus
        Else
            Call Resalta(Me.TxtImporte)
        
        End If
     Else
    Me.cmdAgregar.SetFocus
    End If
End If
End Sub

Private Sub TxtNumero_Ref_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtNumero_Ref.Text = FormatosCeros(Me.TxtNumero_Ref.Text, 10)
    Call Resalta(Me.TxtCodProducto)
End If
End Sub



Private Sub TxtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtCodProveedor)
End If
If KeyCode = vbKeyRight Then
    Me.DtcTipoDoc_Ref.SetFocus
End If
End Sub

Private Sub TxtProveedor_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
        Call Resalta(Me.TxtDireccion)
                
End If
End Sub

Private Sub TxtCodProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtNumeroDoc)
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.TxtProveedor)
End If
End Sub

Private Sub TxtCodProveedor_KeyPress(KeyAscii As Integer)
On Error GoTo errohandler
 If KeyAscii = 13 Then
 If (Trim(Me.TxtCodProveedor.Text) = "") Then
    Procedencia = Selecionar
    FrmPersona.Show
    
    Exit Sub
End If



If (Len(Trim(Me.TxtCodProveedor.Text)) <= 5 And Trim(Me.TxtCodProveedor.Text) <> "") Then
    Me.TxtCodProveedor.Text = formato_item(Me.TxtCodProveedor.Text, 5)
    strCadena = "SELECT *  FROM Persona WHERE cPersona='" & Trim(Me.TxtCodProveedor.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.TxtProveedor.Text = rst("NombrePersona")
        Me.TxtDireccion.Text = rst("sDireccionCliente1")
        Me.TxtRuc.Text = rst("Per_Ruc")
        Call Me.DtcTipoDoc_Ref.SetFocus
        Exit Sub
    Else
        Call Resalta(Me.TxtCodProveedor)
        Exit Sub
    End If
    Set rst = Nothing
End If


 
If Len(Trim(Me.TxtCodProveedor.Text)) = 8 Then
    strCadena = "SELECT *  FROM Persona WHERE Per_Ruc='" & Trim(Me.TxtCodProveedor.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.TxtRuc.Text = Trim(Me.TxtCodProveedor.Text)
       ' FrmDetallePersona.OptJuridica.Value = True
        FrmDetallePersona.chkProveedor.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.TxtCodProveedor.Text = rst("cPersona")
        Me.TxtProveedor.Text = rst("NombrePersona")
        Me.TxtRuc.Text = rst("Per_Ruc")
        Me.TxtDireccion.Text = rst("sDireccionCliente1")
        Me.DtcTipoDoc_Ref.SetFocus
       
    End If
End If



If Len(Trim(Me.TxtCodProveedor.Text)) = 11 Then
    strCadena = "SELECT *  FROM Persona WHERE Per_Ruc='" & Trim(Me.TxtCodProveedor.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.TxtRuc.Text = Trim(Me.TxtCodProveedor.Text)
       ' FrmDetallePersona.OptJuridica.Value = True
        FrmDetallePersona.chkProveedor.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
         Me.TxtCodProveedor.Text = rst("cPersona")
        Me.TxtProveedor.Text = rst("NombrePersona")
         Me.TxtRuc.Text = rst("Per_Ruc")
        Me.TxtDireccion.Text = rst("sDireccionCliente1")
        Call Me.DtcTipoDoc_Ref.SetFocus
       
    End If
End If
End If
    

If (KeyAscii = 66 Or KeyAscii = 98) Then
    Procedencia = Selecionar
    FrmPersona.Show
End If
Exit Sub
errohandler: MsgBox "Hubo un Error Digite Nuevamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub TxtCodProducto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtCantidad)
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.TxtObservacion)
End If
End Sub

Private Sub TxtCodProducto_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then
    
     strCadena = "SELECT Producto_barras.cProducto, Producto.DescripcionProducto, Producto.PrecioCompra, Unidad.sAbreviatura " & _
     "FROM Producto INNER JOIN Producto_barras ON Producto.cProducto = Producto_barras.cProducto INNER JOIN Unidad ON " & _
     "Producto.cUnidad = Unidad.cUnidad WHERE Producto_barras.cod_barra='" & Trim(Me.TxtCodProducto.Text) & "'"
    
       Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.TxtDescripcionProducto.Text = rst(1)
        Me.txtCosto.Text = rst(2)
        Me.TxtUnidad.Text = rst(3)
        Me.TxtCantidad.Text = 0
        Call Resalta(Me.TxtCantidad)
        Set rst = Nothing
        
    Else
        Procedencia = Selecionar
        FrmProducto.Show
    End If
End If
End Sub

Private Sub TxtDireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtRuc)
End If
End Sub


Private Sub Llenar_Temporal()
'Dim RstTemporal As New ADODB.Recordset
'Dim rstDetalle As New ADODB.Recordset
'Dim i As Integer
'StrCadena = "SELECT * FROM Detalle_DocumentoCompra WHERE (cDocumentoCompra='" & Trim(Me.TxtNumero_guia.Text) & "' AND doc_cod='" & Trim(Me.DtcComprobanteGuia.BoundText) & "' AND " & _
'"sSerie='" & Trim(Me.TxtSeri_guia.Text) & "') "
'rstDetalle.Open StrCadena, CnBd, adOpenKeyset, adLockOptimistic
'StrCadena = "SELECT * FROM Temporal_Compras"
'RstTemporal.Open StrCadena, CnBd, adOpenKeyset, adLockOptimistic
'rstDetalle.MoveFirst

 'For i = 0 To rstDetalle.RecordCount - 1
  ' StrCadena = "SELECT cTemporal FROM Temporal_Compras ORDER BY cTemporal DESC "
   ' RstTemporal.AddNew
    'RstTemporal.Fields(0) = GeneraCodTemporal
    'RstTemporal.Fields(1) = Trim(Me.TxtNumeroDoc.Text)
    ''RstTemporal.Fields(2) = Trim(Me.DtcTipoDoc.BoundText)
    'RstTemporal.Fields(3) = Trim(Me.TxtSerie.Text)
    'RstTemporal.Fields(4) = rstDetalle.Fields(4)
    'RstTemporal.Fields(5) = rstDetalle.Fields(5)
    'RstTemporal.Fields(6) = rstDetalle.Fields(6)
    'RstTemporal.Fields(7) = rstDetalle.Fields(7)
    'RstTemporal.Update
    'rstDetalle.MoveNext
       
 'Next i
'Set RstTemporal = Nothing
'Set rstDetalle = Nothing
End Sub

Private Sub TxtNumeroDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtSerie)
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.TxtCodProveedor)
End If
End Sub

Private Sub TxtNumeroDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Me.TxtNumeroDoc.Text = FormatosCeros(Me.TxtNumeroDoc.Text, 10)
    strCadena = "SELECT cDocumentoCompra,sSerie FROM DocumentoCompra WHERE (cDocumentoCompra='" & Trim(Me.TxtNumeroDoc.Text) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Me.HfdDetalle.Clear
        If CVDate(Me.DtpActual.Value) <> Date Then
            If MsgBox("la Fecha no Coincide con la Fecha del Documento...Desea Continuar", vbInformation + vbYesNo, KEY_EMPRESA) = vbYes Then
                Me.TxtCodProveedor.Text = "00004"
                Call Resalta(Me.TxtCodProveedor)
            Else
                Me.DtpActual.SetFocus
            End If
        Else
        Me.TxtCodProveedor.Text = "00004"
        Call Resalta(Me.TxtCodProveedor)
        ProcendenciaGuia = NuevaGuia
    End If
    Else
        MsgBox "Documento ya Existe", vbInformation, KEY_EMPRESA
        Call LlenarDatosCliente(Trim(Me.TxtNumeroDoc.Text), Trim(Me.DtcTipoDoc.BoundText), Trim(Me.TxtSerie.Text))
        Call llenarGrid_Comprobante(Me.HfdDetalle, Trim(Me.TxtNumeroDoc.Text), Me.DtcTipoDoc.BoundText, Trim(Me.TxtSerie.Text))
       Call LLenarDatosReferencia(Trim(Me.TxtNumeroDoc.Text), Me.DtcTipoDoc.BoundText, Trim(Me.TxtSerie.Text)) 'extraer los datos Guardados del documento referncia
       Call VerificaAnulado(Trim(Me.TxtNumeroDoc.Text), Trim(Me.DtcTipoDoc.BoundText), Trim(Me.TxtSerie.Text), Trim(Me.DtcAlmacen.BoundText))
        Me.TxtCodProducto.Enabled = False
        Me.TxtDescripcionProducto.Enabled = False
        
        Me.cmdAgregar.Enabled = False
        Me.cmdQuitar.Enabled = False
        
        Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True
        Me.TxtCantidad.Enabled = False
        
        Call Resalta(Me.TxtNumeroDoc)
        
        
    End If
End If
Set rst = Nothing
End Sub
Private Sub LLenarDatosReferencia(ByVal Numero As String, ByVal Documento As String, ByVal serie As String)
Dim Numref As Double
strCadena = "Select IdReferencia FROM DocumentoCompra WHERE (cDocumentoCompra ='" & Numero & "' AND doc_cod='" & Documento & "' AND sSerie='" & serie & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 And rst(0) <> "" Then
    Numref = rst(0)
    
    Set rst = Nothing
    strCadena = "SELECT doc_cod,sSerie,cDocumentoCompra,FechaProceso FROM DocReferencia_Compra WHERE IdReferencia='" & Numref & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.DtcTipoDoc_Ref.BoundText = Trim(rst(0))
        Me.TxtSerie_Ref.Text = Trim(rst(1))
        Me.TxtNumero_Ref.Text = Trim(rst(2))
        Me.DtpFechaReferencia.Value = rst(3)
        Set rst = Nothing

    Else
        
        Set rst = Nothing
End If
Else
        Me.DtcTipoDoc_Ref.Text = ""
        Me.TxtSerie_Ref.Text = ""
        Me.TxtNumero_Ref.Text = ""
        Me.DtpFechaReferencia.Value = Date
End If
End Sub
Private Sub LlenarDatosCliente(ByVal Numero As String, ByVal Documento As String, ByVal serie As String)
Dim CodPersona As String
strCadena = "SELECT cPersona,dEmisionCompra,idFormaPago FROM DocumentoCompra WHERE (cDocumentoCompra ='" & Numero & "' AND doc_cod='" & Documento & "' AND sSerie='" & serie & "')"
Call ConfiguraRst(strCadena)
    CodPersona = Trim(rst(0))
    Me.DtpActual.Value = CVDate(rst(1))
    Me.DtcFormaPago.BoundText = rst(2)
Set rst = Nothing
strCadena = "SELECT cPersona,NombrePersona,sDireccionCliente1,Per_Ruc,Observacion,Per_Nat FROM Persona WHERE (cPersona ='" & CodPersona & "' )"
Call ConfiguraRst(strCadena)
    
        Me.TxtCodProveedor.Text = rst(0)
        Me.TxtProveedor.Text = rst(1)
        Me.TxtDireccion.Text = rst(2)
        Me.TxtRuc.Text = rst(3)
        Me.TxtObservacion.Text = rst(4)
        
    
    
Set rst = Nothing
Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
End Sub
Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    If Me.TxtCodProducto.Text = "" Then
        Me.TxtCodProducto.SetFocus
    Else
        Me.DtcTipoDoc_Ref.SetFocus
    End If
End If
End Sub



Private Sub TxtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtDescripcionProducto)
       
End If
If KeyCode = vbKeyRight Then
     Me.cmdAgregar.SetFocus
End If
End Sub

Private Sub TxtPrecio_KeyPress(KeyAscii As Integer)
Dim TotalP As Single
If KeyAscii = 13 Then
        
       ' TotalP = Val(Me.TxtCantidad.Text) * Val(Me.TxtPrecio.Text)
        'Me.LblTotalParcial.Text = Format(TotalP, "#,##0.000")
        Me.cmdAgregar.SetFocus
End If
End Sub
Private Sub Save()
Dim i As Integer
Dim anul As String * 1
Dim CodReferencia As String
Dim tSaldo As Double
Dim TotalFactura As Double
Dim ValorCompra As Double
Dim igv As Double
TotalFactura = Me.lblTotal.Caption
ValorCompra = Me.lblValorVenta.Caption
igv = Me.lblIgv.Caption
anul = "F"
If Trim(Me.DtcFormaPago.BoundText) = "0001" Then
    tSaldo = 0
Else
    tSaldo = TotalFactura
End If
'01----------------guardar en Documento Compra---------------------
If Referencia = True Then
CodReferencia = GeneraCodReferencia
strCadena = "INSERT INTO DocumentoCompra(cDocumentoCompra,doc_cod,Alm_cod,sSerie,cPersona,Persona,Observacion,idFormaPago, " & _
            "dEmisionCompra,dVencimiento,dPago,nSubTotal,nIgv,nTotalCompra,FechaProceso,IdReferencia," & _
            "IntDocumentoCompra,Anulado,IdUsuario,saldo)VALUES ('" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Trim(Me.DtcAlmacen.BoundText) & "','" & Trim(Me.TxtSerie.Text) & "'," & _
            "'" & Trim(Me.TxtCodProveedor.Text) & "','" & Trim(Me.TxtProveedor.Text) & "','" & Trim(Me.TxtObservacion.Text) & "','" & Trim(Me.DtcFormaPago.BoundText) & "','" & Me.DtpActual.Value & "','" & Me.DtpActual.Value & "','" & Me.DtpFechaReferencia.Value & "','" & Val(ValorCompra) & "'," & _
            "'" & Val(igv) & "','" & Val(TotalFactura) & "','" & Me.DtpActual.Value & "','" & CodReferencia & "','" & Val(Me.TxtNumeroDoc.Text) & "','" & anul & "','" & Trim(KEY_USUARIO) & "','" & tSaldo & "')"
            Call EjecutaRST(strCadena)
            Set RstEjecuta = Nothing
            Call SaveReferencia(CodReferencia, Trim(Me.DtcTipoDoc_Ref.BoundText), Trim(Me.TxtSerie_Ref.Text), Trim(Me.TxtNumero_Ref.Text), CVDate(Me.DtpActual.Value), Trim(Me.DtcAlmacen.BoundText))
            Referencia = False
Else
strCadena = "INSERT INTO DocumentoCompra(cDocumentoCompra,doc_cod,Alm_cod,sSerie,cPersona,Persona,Observacion,idFormaPago," & _
            "dEmisionCompra,dVencimiento,dPago,nSubTotal,nIgv,nTotalCompra,FechaProceso,intDocumentoCompra,Anulado,IdUsuario,saldo)" & _
            " VALUES ('" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Trim(Me.DtcAlmacen.BoundText) & "','" & Trim(Me.TxtSerie.Text) & "'," & _
            "'" & Trim(Me.TxtCodProveedor.Text) & "','" & Trim(Me.TxtProveedor.Text) & "','" & Trim(Me.TxtObservacion.Text) & "','" & Trim(Me.DtcFormaPago.BoundText) & "','" & Me.DtpActual.Value & "','" & Me.DtpActual.Value & "','" & Me.DtpFechaReferencia.Value & "','" & Val(ValorCompra) & "'," & _
            "'" & Val(igv) & "','" & Val(TotalFactura) & "','" & Me.DtpActual.Value & "','" & Val(Me.TxtNumeroDoc.Text) & "','" & anul & "','" & Trim(KEY_USUARIO) & "','" & tSaldo & "')"
            Call EjecutaRST(strCadena)
            Set rst = Nothing
End If

 '01-------------------------------------------------------------
 
 '02----------------guardar en detalle documento Compra-----------
 Call SaveDetalleDocumentoCompra
 '02-------------------------------------------------------------
                Me.TxtCodProducto.Enabled = False
                Me.TxtDescripcionProducto.Enabled = False
                Me.TxtCantidad.Enabled = False
                Me.txtCosto.Enabled = False
                 Me.cmdAgregar.Enabled = False
                Me.cmdQuitar.Enabled = False
                Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
                Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
                Exit Sub
End Sub
Private Sub SaveDetalleDocumentoCompra()
Dim precio_costo As Double
Dim precio_venta As Double
Dim rst_pp As New ADODB.Recordset
Dim rst_precio_compra As New ADODB.Recordset
Dim Codkardex As String
   strCadena = "SELECT * FROM Temporal_Compras WHERE (cDocumentoCompra='" & Trim(Me.TxtNumeroDoc.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND id_usuario='" & KEY_USUARIO & "')"
    rstTemporal.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
    
    If rstTemporal.RecordCount > 0 Then
        rstTemporal.MoveFirst
        For i = 0 To rstTemporal.RecordCount - 1
            strCadena = "SELECT * FROM Detalle_DocumentoCompra"
            RstDetCompra.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
            RstDetCompra.AddNew
            StrCodDetCompra = CodigoDetalleCompra
            Codkardex = CodigoKardex
            RstDetCompra("cDet_documentoCompra") = StrCodDetCompra
            RstDetCompra("cDocumentoCompra") = rstTemporal("cDocumentoCompra")
            RstDetCompra("doc_cod") = rstTemporal("doc_cod")
            RstDetCompra("Alm_cod") = Trim(Me.DtcAlmacen.BoundText)
            RstDetCompra("sSerie") = rstTemporal("sSerie")
            RstDetCompra("cProducto") = rstTemporal("cProducto")
            RstDetCompra("cantidad") = rstTemporal("Cantidad")
            RstDetCompra("imp_uni") = rstTemporal("imp_uni")
            RstDetCompra("imp_bruto") = rstTemporal("imp_bruto")
            RstDetCompra("desct") = rstTemporal("desct")
            RstDetCompra("valor_vta") = rstTemporal("valor_vta")
            RstDetCompra("igv") = rstTemporal("igv")
            RstDetCompra("total") = rstTemporal("total")
            RstDetCompra("Int_det_DocumentoCompra") = Val(StrCodDetCompra)
            RstDetCompra.Update
            Set RstDetCompra = Nothing
            
            precio_costo = rstTemporal("imp_uni")
            strCadena = "SELECT * FROM Producto WHERE cProducto='" & Trim(rstTemporal("cProducto")) & "'"
            rst_precio_compra.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
            rst_precio_compra("PrecioCompra") = precio_costo
            precio_venta = rst_precio_compra("PrecioVenta")
            rst_precio_compra.Update
            Set rst_precio_compra = Nothing
            
            strCadena = "SELECT * FROM Producto_precio  ORDER BY id_producto_precio DESC"
            Call ConfiguraRst(strCadena)
            StrNumero = GeneraCodigo(10)
            Set rst = Nothing
            strCadena = "SELECT * FROM Producto_precio"
            rst_pp.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
            rst_pp.AddNew
            rst_pp("id_producto_precio") = StrNumero
            rst_pp("cProducto") = rstTemporal("cProducto")
            rst_pp("precio_venta") = precio_venta
            rst_pp("precio_costo") = precio_costo
            rst_pp("fecha") = Me.DtpActual.Value
            rst_pp.Update
            Set rst_pp = Nothing
            
            
            
            'Call save_preciocompra(RstTemporal("cProducto"), precio_costo)
            
            Call Kardex(rstTemporal("cProducto"), rstTemporal("doc_cod"), Trim(Me.DtcAlmacen.BoundText), rstTemporal("cDocumentoCompra"), rstTemporal("sSerie"), KEY_ING, CVDate(Date), CVDate(Me.DtpActual.Value), rstTemporal("Cantidad"), rstTemporal("Cantidad"), , rstTemporal("Cantidad"), precio_costo, rstTemporal("total"), , rstTemporal("total"), Trim(Me.TxtProveedor.Text), Trim(Codkardex), dfactura)
            Call ActualizaStock_Almacenes(Trim(Me.DtcAlmacen.BoundText), Trim(rstTemporal("cProducto")), Trim(rstTemporal("Cantidad")), Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcTipoDoc_Ref.BoundText), Trim(Me.TxtSerie_Ref.Text), Trim(Me.TxtNumero_Ref.Text), dfactura)
            
            If rstTemporal.EOF = False Then
                rstTemporal.MoveNext
            End If
            
        Next i
    End If
                Set rstTemporal = Nothing
                Set RstDetCompra = Nothing
                strCadena = "DELETE FROM Temporal_Compras WHERE (cDocumentoCompra='" & Trim(Me.TxtNumeroDoc.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "')"
                Call EjecutaRST(strCadena)
                Set RstEjecuta = Nothing
                Set rstTemporal = Nothing
                Set RstDetCompra = Nothing
End Sub
Private Function CodigoDetalleCompra() As String
    strCadena = "SELECT int_det_documentoCompra FROM Detalle_DocumentoCompra ORDER BY int_det_documentoCompra DESC"
    Call ConfiguraRst(strCadena)
    CodigoDetalleCompra = GeneraCodigo(20)
    Set rst = Nothing
End Function
Private Sub save_preciocompra(ByVal Codigo As String, ByVal precio As Double)
Dim consecutivo As String
strCadena = "SELECT * FROM Producto_precio WHERE cProducto='" & Trim(Codigo) & "'"
Call ConfiguraRst(strCadena)
consecutivo = GeneraCodigo(10)



End Sub

Private Sub TxtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtObservacion)
End If
End Sub

Private Sub TxtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Me.DtcTipoDoc.SetFocus
End If
If KeyCode = vbKeyRight Then
        Call Resalta(Me.TxtNumeroDoc)
End If
End Sub
Private Sub TxtSerie_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Then
    Me.TxtSerie.Text = FormatosCeros(Me.TxtSerie.Text, 4)
    strCadena = "SELECT Alm_cod,doc_cod,serie FROM Det_alm_com WHERE (Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'" & _
        " AND  doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If Me.TxtNumeroDoc.Text = "" Then
            
            Me.TxtNumeroDoc.SetFocus
        Else
            Set rst = Nothing
            strCadena = "SELECT cDocumentoCompra FROM DocumentoCompra WHERE (doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "') ORDER BY intDocumentoCompra DESC"
            Call ConfiguraRst(strCadena)
            Me.TxtNumeroDoc.Text = GeneraCodigo(10)
            Call Resalta(Me.TxtNumeroDoc)
        End If
    Else
        MsgBox "Serie no Asiganda a a dicho Almacen", vbInformation, KEY_EMPRESA
        Call Resalta(Me.TxtSerie)
    End If
End If
End Sub

Private Sub VerificaAnulado(ByVal Numero As String, ByVal Documento As String, ByVal serie As String, ByVal Almacen As String)
strCadena = "Select Anulado FROM DocumentoCompra WHERE (cDocumentoCompra ='" & Numero & "' AND doc_cod='" & Documento & "' AND sSerie='" & serie & "'AND Alm_Cod='" & Almacen & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    If Trim(rst(0)) = "V" Then
        Me.lblAnulado.Visible = True
        
        Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
        Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
    Else
        Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
    End If
End If
Set rst = Nothing
End Sub




Private Sub TxtSerie_Ref_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie_Ref.Text = FormatosCeros(Me.TxtSerie_Ref.Text, 4)
    Call Resalta(Me.TxtNumero_Ref)
End If
End Sub

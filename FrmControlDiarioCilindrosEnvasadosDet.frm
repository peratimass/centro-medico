VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmControlDiarioCilindrosEnvasadosDet 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAgregarM 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   9990
      TabIndex        =   28
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox TxtCantidadM 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   8070
      MaxLength       =   80
      TabIndex        =   27
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtcodProductoMan 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1605
      MaxLength       =   80
      TabIndex        =   25
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox TxtDescripcionProductoM 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2790
      MaxLength       =   80
      TabIndex        =   24
      Top             =   4560
      Width           =   5175
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   10080
      TabIndex        =   22
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox TxtCantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   8040
      MaxLength       =   80
      TabIndex        =   21
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   8040
      MaxLength       =   80
      TabIndex        =   20
      Top             =   1140
      Width           =   1575
   End
   Begin VB.TextBox TxtCodProducto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1575
      MaxLength       =   80
      TabIndex        =   18
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox TxtDescripcionProducto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2760
      MaxLength       =   80
      TabIndex        =   17
      Top             =   1680
      Width           =   5175
   End
   Begin VB.TextBox TxtNumeroDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   4
      Text            =   "0000000000"
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox TxtSerie 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   3
      Text            =   "0000"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox TxtCodCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1575
      MaxLength       =   80
      TabIndex        =   2
      Top             =   1140
      Width           =   1095
   End
   Begin VB.TextBox TxtCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2760
      MaxLength       =   80
      TabIndex        =   1
      Top             =   1140
      Width           =   5175
   End
   Begin VB.TextBox TxtObservacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   1560
      MaxLength       =   80
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   6600
      Width           =   5055
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   8520
      Top             =   4800
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
            Picture         =   "FrmControlDiarioCilindrosEnvasadosDet.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioCilindrosEnvasadosDet.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioCilindrosEnvasadosDet.frx":077C
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioCilindrosEnvasadosDet.frx":0809
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioCilindrosEnvasadosDet.frx":0C69
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioCilindrosEnvasadosDet.frx":0F85
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioCilindrosEnvasadosDet.frx":13E5
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioCilindrosEnvasadosDet.frx":1701
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioCilindrosEnvasadosDet.frx":1B61
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioCilindrosEnvasadosDet.frx":1FC1
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioCilindrosEnvasadosDet.frx":28A1
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioCilindrosEnvasadosDet.frx":2BBD
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioCilindrosEnvasadosDet.frx":2ED9
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   10500
      TabIndex        =   5
      Top             =   6480
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1482
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1755
      _CBHeight       =   840
      _Version        =   "6.0.8169"
      Child1          =   "TlbGrabar"
      MinHeight1      =   780
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbGrabar 
         Height          =   780
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   1635
         _ExtentX        =   2884
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
      Height          =   315
      Left            =   10560
      TabIndex        =   7
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   -2147483635
      CalendarTitleBackColor=   16777215
      CalendarTitleForeColor=   -2147483635
      Format          =   61865985
      CurrentDate     =   37091
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   360
      Left            =   600
      TabIndex        =   8
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones1 
      Height          =   4185
      Left            =   11520
      TabIndex        =   9
      Top             =   960
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   7382
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   4185
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   5460
         Left            =   30
         TabIndex        =   10
         Top             =   345
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   9631
         ButtonWidth     =   1693
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "  &Nuevo   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Anular"
               Key             =   "(Anular)"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Eliminar"
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
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   360
      Left            =   3480
      TabIndex        =   11
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      ForeColor       =   8388608
      ListField       =   "0000º"
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfLlenos 
      Height          =   1935
      Left            =   1560
      TabIndex        =   12
      Top             =   2160
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1335
      Left            =   1560
      TabIndex        =   23
      Top             =   5040
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2355
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfVacios 
      Height          =   1935
      Left            =   6720
      TabIndex        =   29
      Top             =   2160
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
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
      Left            =   480
      TabIndex        =   26
      Top             =   4560
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
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
      Left            =   450
      TabIndex        =   19
      Top             =   1680
      Width           =   885
   End
   Begin VB.Label Label6 
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
      Left            =   9855
      TabIndex        =   16
      Top             =   240
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   12135
   End
   Begin VB.Label lblAnulado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   7440
      TabIndex        =   15
      Top             =   5760
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Responsable:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1140
      Width           =   1305
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle :"
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
      Left            =   285
      TabIndex        =   13
      Top             =   6720
      Width           =   765
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000000&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   7455
      Left            =   0
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "FrmControlDiarioCilindrosEnvasadosDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

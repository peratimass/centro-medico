VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmOrdenpago 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtCodCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   375
      MaxLength       =   80
      TabIndex        =   13
      Top             =   1500
      Width           =   1335
   End
   Begin VB.TextBox TxtCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      MaxLength       =   80
      TabIndex        =   12
      Top             =   1500
      Width           =   4695
   End
   Begin VB.TextBox TxtObservacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   1800
      MaxLength       =   80
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2580
      Width           =   4695
   End
   Begin VB.TextBox TxtDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      MaxLength       =   80
      TabIndex        =   10
      Top             =   1860
      Width           =   4695
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      MaxLength       =   80
      TabIndex        =   9
      Top             =   2220
      Width           =   1695
   End
   Begin VB.TextBox TxtMontoIngresar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox TxtSerie 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3870
      MaxLength       =   80
      TabIndex        =   7
      Text            =   "0000"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox TxtNumeroDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4875
      MaxLength       =   80
      TabIndex        =   6
      Text            =   "0000000000"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox ChkAdelantado 
      Caption         =   "Adelantado"
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo Cambio"
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   6840
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
      Begin VB.TextBox TxtTC 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         MaxLength       =   80
         TabIndex        =   4
         Top             =   200
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdListado 
      Caption         =   "Listado de Recibos"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   5760
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   7200
      Top             =   360
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
            Picture         =   "FrmOrdenpago.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenpago.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenpago.frx":077C
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenpago.frx":0809
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenpago.frx":0C69
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenpago.frx":0F85
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenpago.frx":13E5
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenpago.frx":1701
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenpago.frx":1B61
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenpago.frx":1FC1
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenpago.frx":28A1
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenpago.frx":2BBD
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenpago.frx":2ED9
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2490
      Left            =   4740
      TabIndex        =   14
      Top             =   3840
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   4392
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1755
      _CBHeight       =   2490
      _Version        =   "6.0.8169"
      Child1          =   "TlbGrabar"
      MinHeight1      =   2430
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbGrabar 
         Height          =   2430
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   4286
         ButtonWidth     =   1402
         ButtonHeight    =   1429
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
      Left            =   4800
      TabIndex        =   16
      Top             =   2205
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   -2147483635
      CalendarTitleBackColor=   16777215
      CalendarTitleForeColor=   -2147483635
      Format          =   135593985
      CurrentDate     =   37091
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   360
      Left            =   1080
      TabIndex        =   17
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
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
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   360
      Left            =   3840
      TabIndex        =   18
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      _Version        =   393216
      Locked          =   -1  'True
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
      Height          =   3225
      Left            =   8400
      TabIndex        =   19
      Top             =   360
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   5689
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   3225
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   4050
         Left            =   30
         TabIndex        =   20
         Top             =   375
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   7144
         ButtonWidth     =   1773
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
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
               Caption         =   "&Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo DtcCCostos 
      Height          =   315
      Left            =   1800
      TabIndex        =   21
      Top             =   3360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDetalle 
      Height          =   3495
      Left            =   240
      TabIndex        =   22
      Top             =   6480
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6165
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
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   315
      Left            =   360
      TabIndex        =   23
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   -2147483635
      CalendarTitleBackColor=   16777215
      CalendarTitleForeColor=   -2147483635
      Format          =   135593985
      CurrentDate     =   37091
   End
   Begin MSComCtl2.DTPicker DtpFinal 
      Height          =   315
      Left            =   2280
      TabIndex        =   24
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   -2147483635
      CalendarTitleBackColor=   16777215
      CalendarTitleForeColor=   -2147483635
      Format          =   135593985
      CurrentDate     =   37091
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos Persona:"
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
      Left            =   990
      TabIndex        =   38
      Top             =   1200
      Width           =   1425
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
      Left            =   585
      TabIndex        =   37
      Top             =   1860
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
      Left            =   1065
      TabIndex        =   36
      Top             =   2220
      Width           =   435
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Glosa:"
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
      Left            =   855
      TabIndex        =   35
      Top             =   2700
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Monto:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1005
      TabIndex        =   34
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label1 
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
      Left            =   3975
      TabIndex        =   33
      Top             =   2280
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   360
      Top             =   240
      Width           =   6135
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   5295
      Left            =   0
      Top             =   0
      Width           =   9495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M.Pago:"
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
      Left            =   675
      TabIndex        =   32
      Top             =   3360
      Width           =   765
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   240
      Top             =   5400
      Width           =   9135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "AL"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2040
      TabIndex        =   31
      Top             =   5760
      Width           =   195
   End
   Begin VB.Shape Shape4 
      Height          =   255
      Left            =   240
      Top             =   6240
      Width           =   9135
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   510
      TabIndex        =   30
      Top             =   6240
      Width           =   585
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Doc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1695
      TabIndex        =   29
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serie"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   2580
      TabIndex        =   28
      Top             =   6240
      Width           =   525
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3420
      TabIndex        =   27
      Top             =   6240
      Width           =   765
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Persona"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   5190
      TabIndex        =   26
      Top             =   6240
      Width           =   825
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   7920
      TabIndex        =   25
      Top             =   6240
      Width           =   645
   End
End
Attribute VB_Name = "FrmOrdenpago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Dim mostrar As Boolean
Dim codigo_P As String





Private Sub cmdBuscar_Click()
strCadena = "SELECT      movimiento_caja.doc_cod,movimiento_caja.fecha_valor, Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, " & _
"     movimiento_caja.descripcion_per , movimiento_caja.Monto " & _
"FROM         movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod " & _
" WHERE  movimiento_caja.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND movimiento_caja.fecha_valor>='" & CVDate(Me.dtpinicio.Value) & "' AND movimiento_caja.fecha_valor<='" & CVDate(Me.DtpFinal.Value) & "' ORDER BY  movimiento_caja.numero DESC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.cmdimprimir.Visible = True
Else
    Me.cmdimprimir.Visible = False
End If
Me.HfDetalle.Clear
Me.HfDetalle.rows = 1
Set Me.HfDetalle.Recordset = rst
Me.HfDetalle.rows = rst.RecordCount
Me.HfDetalle.ColWidth(0) = 0
Me.HfDetalle.ColWidth(1) = 1300
Me.HfDetalle.ColWidth(2) = 1300
Me.HfDetalle.ColWidth(3) = 600
Me.HfDetalle.ColWidth(4) = 1000
Me.HfDetalle.ColWidth(5) = 3000
Me.HfDetalle.ColWidth(6) = 1850
Call DarFormatoFecha(Me.HfDetalle, 1)
End Sub

Private Sub cmdimprimir_Click()
 strCadena = "SELECT     Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, movimiento_caja.cPersona, " & _
        "movimiento_caja.descripcion_per, Persona.sDireccionCliente1, Persona.Per_Ruc, movimiento_caja.fecha_valor," & _
        "movimiento_caja.cambio , movimiento_caja.glosa, centro_costos.descripcion, movimiento_caja.Monto,movimiento_caja.monto_letras " & _
        "FROM movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
        "centro_costos ON movimiento_caja.id_costo = centro_costos.id_costo INNER JOIN " & _
        "Persona ON movimiento_caja.cPersona = Persona.cPersona WHERE movimiento_caja.fecha_valor>='" & CVDate(Me.dtpinicio.Value) & "' AND movimiento_caja.fecha_valor<='" & CVDate(Me.DtpFinal.Value) & "' AND movimiento_caja.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "'"
    
     Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptRecibos", , App.Path + "\Reportes\")
        Set rst = Nothing

    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdListado_Click()
If mostrar = False Then
mostrar = True
Me.Height = 10290
Me.Shape3.Height = 10290
Else
mostrar = False
Me.Height = 5295
Me.Shape3.Height = 5295
End If
End Sub

Private Sub DtcAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcTipoDoc.SetFocus
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
    If Me.DtcAlmacen.Enabled = True Then
    
    strCadena = "SELECT numero FROM Det_alm_com WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "') ORDER BY numero DESC"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.TxtNumeroDoc.Text = rst(0)
    Else
    Me.TxtNumeroDoc.Text = GeneraCodigo(6)
    End If
    Set rst = Nothing
    Me.TxtCodCliente.Text = "00000"
    Me.TxtCliente.Text = ""
    Me.txtdireccion.Text = ""
    Me.txtobservacion.Text = ""
    Me.TxtMontoIngresar.Text = "0.00"
    Me.DtpActual.Value = Date
    
    Me.TxtRuc.Text = ""
    Me.DtcAlmacen.SetFocus
    Set rst = Nothing
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
Else
    MsgBox "Active el Almacen Correspondiente", vbInformation, KEY_EMPRESA
End If
End Sub

Private Sub Form_Activate()
CenterForm Me
Me.Top = 300
Call Resalta(Me.TxtCodCliente)
End Sub
Private Sub ActualizarAdelanto(ByVal TotalPedido As Double)
Dim MontoAnterior As Double
strCadena = "SELECT MontoAdelantado FROM Persona WHERE cPersona='" & Trim(Me.TxtCodCliente.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    MontoAnterior = rst(0)
    Set rst = Nothing
End If
strCadena = "UPDATE Persona SET MontoAdelantado='" & (MontoAnterior - TotalPedido) & "' WHERE cPersona='" & Trim(Me.TxtCodCliente.Text) & "'"
Call EjecutaRST(strCadena)
Set RstEjecuta = Nothing
End Sub
Private Sub Form_Load()
Me.DtpFinal.Value = Date
Me.dtpinicio.Value = Date
doc_Tienda = "V"
mostrar = False
 strCadena = "SELECT Alm_cod as Codigo, Alm_des as Descripcion FROM Almacen " & _
  " ORDER BY Alm_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = "0001"
   Set rst = Nothing
  
  strCadena = "SELECT doc_cod as Codigo, doc_abrev as Descripcion FROM Comprobantes " & _
  " WHERE doc_tienda='" & doc_Tienda & "' ORDER BY doc_abrev"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
  Me.DtcTipoDoc.BoundText = "0096"
  Me.TxtSerie.Text = "0003"
  
  Set rst = Nothing
 
  


    strCadena = "SELECT numero FROM Det_alm_com WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "') ORDER BY numero DESC"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.TxtNumeroDoc.Text = rst(0)
    Else
        Me.TxtNumeroDoc.Text = GeneraCodigo(6)
    End If
    Set rst = Nothing
    
  strCadena = "SELECT id_costo as Codigo, descripcion as Descripcion FROM centro_costos " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCCostos)
  Me.TxtTc.Text = cambio(Date)

End Sub
Private Sub LlenarDatosCliente(ByVal Numero As String, ByVal Documento As String, ByVal serie As String, ByVal Almacen As String)
Dim CodPersona As String
strCadena = "SELECT cPersona,Persona,dEmisionVenta,nTotalVenta,Observacion FROM DocumentoVenta WHERE (cDocumentoVenta ='" & Numero & "' AND doc_cod='" & Documento & "' AND sSerie='" & serie & "' AND Alm_Cod='" & Almacen & "')"
Call ConfiguraRst(strCadena)
    CodPersona = Trim(rst(0))
    Me.TxtCodCliente.Text = CodPersona
    Me.TxtCliente.Text = rst(1)
    Me.DtpActual.Value = CVDate(rst(2))
    Me.TxtMontoIngresar.Text = Format(rst(3) * -1, "#,##0.00")
    On Error GoTo SALIR
    Me.txtobservacion.Text = rst(4)
    'Me.DtpFechaReferencia.Value = CVDate(Rst(2))
    'Me.DtcFormaPago.BoundText = Rst(3)
SALIR:
Set rst = Nothing
strCadena = "SELECT sDireccionCliente1,Per_Ruc FROM Persona WHERE (cPersona ='" & CodPersona & "' )"
Call ConfiguraRst(strCadena)

        Me.txtdireccion.Text = rst(0)
        Me.TxtRuc.Text = rst(1)
        

Set rst = Nothing
Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True

End Sub
Private Sub Save()
Dim id_documento As String
Dim i As Integer
Dim Monto As Double
Dim codigo As String

Dim t_cambio As Single
Dim anul As String * 1
Dim Contado As String
Dim Adelantado As String
Dim moneda As String
Dim fecha_sis As Date
Dim Num_registros As Integer
Dim monto_letras As String
Monto = Me.TxtMontoIngresar.Text
    
    strCadena = "SELECT codigo FROM movimiento_caja WHERE id_cuenta='1011' ORDER BY codigo DESC"
     Call ConfiguraRst(strCadena)
     Num_registros = rst.RecordCount
     codigo = GeneraCodigo(14)
     moneda = "soles"
'01----------------guardar en Documento venta---------------------
      t_cambio = Me.TxtTc.Text
     
     Ingreso = 0
     Egreso = 0
     fecha_sis = Date
     strCadena = "SELECT sum(Ingreso-Egreso) FROM movimiento_caja WHERE id_cuenta='1011' "
     Call ConfiguraRst(strCadena)
     
     
     
        TM = "E"
        Egreso = Monto
        If Num_registros > 0 Then
            Saldo = rst(0) - Monto
        Else
            Saldo = -1 * Monto
        End If
     
     'Calculo del Saldo
     
     monto_letras = UCase(EnLetras(Monto))
     Set rst = Nothing
     'Fin Calculo del Saldo
        
     
        strCadena = "INSERT INTO movimiento_caja VALUES ('" & Trim(codigo) & "','" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Trim(Me.TxtSerie.Text) & "'," & _
        "'" & Trim(Me.TxtNumeroDoc.Text) & "','E','" & Trim(moneda) & "','" & Monto & "','" & Ingreso & "','" & Egreso & "','" & Saldo & "','','" & CVDate(Me.DtpActual.Value) & "','" & fecha_sis & "'," & _
        "'" & Trim(Me.DtcCCostos.BoundText) & "','" & Trim(Me.txtobservacion.Text) & "','" & t_cambio & "','" & Trim(Me.TxtRuc.Text) & "','" & Trim(Me.TxtCodCliente.Text) & "','" & Trim(Me.TxtCliente.Text) & "','" & Trim(monto_letras) & "','','1011')"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        
        Dim nuevo_numero As String
        
        nuevo_numero = formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6)
        strCadena = "UPDATE  Det_alm_com SET numero='" & Trim(nuevo_numero) & "'  WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "')"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        
        

'StrCadena = "INSERT INTO DocumentoVenta(id_documentoventa,cDocumentoVenta,doc_cod,Alm_cod,sSerie,cPersona,Persona,Observacion,idFormaPago," & _
            "dEmisionVenta,nTotalVenta,FechaProceso,intDocumentoVenta,Anulado,id_usuario)" & _
            "VALUES ('" & Trim(id_documento) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.DtcTipoDoc.BoundText) & "'," & _
            "'" & Trim(Me.DtcAlmacen.BoundText) & "','" & Trim(Me.TxtSerie.Text) & "'," & _
            "'" & Trim(Me.TxtCodCliente.Text) & "','" & Trim(Me.TxtCliente.Text) & "','" & Trim(Me.TxtObservacion.Text) & "','" & Trim(Contado) & "'," & _
            "'" & Me.DtpActual.Value & "','" & MontoIngresar * -1 & "','" & CVDate(Date) & "','" & Val(Me.TxtNumeroDoc.Text) & "','" & anul & "','" & KEY_USUARIO & "')"
            
            
            
            Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
            Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
            Exit Sub
    

End Sub

Private Sub HfDetalle_DblClick()
If Me.HfDetalle.rows > 0 Then
     strCadena = "SELECT     Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, movimiento_caja.cPersona, " & _
        "movimiento_caja.descripcion_per, Persona.sDireccionCliente1, Persona.Per_Ruc, movimiento_caja.fecha_valor," & _
        "movimiento_caja.cambio , movimiento_caja.glosa, centro_costos.descripcion, movimiento_caja.Monto,movimiento_caja.monto_letras " & _
        "FROM movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
        "centro_costos ON movimiento_caja.id_costo = centro_costos.id_costo INNER JOIN " & _
        "Persona ON movimiento_caja.cPersona = Persona.cPersona WHERE movimiento_caja.serie='" & Trim(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 3)) & "' AND movimiento_caja.numero='" & Trim(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 4)) & "' AND movimiento_caja.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "'"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptReciboCaja", , App.Path + "\Reportes\")
        Set rst = Nothing
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Error
  Select Case Button.key
    Case KEY_NEW
      Call nuevo
    Case KEY_DELETE
       If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Procedencia = Eliminar
        FrmSeguridad.Show
       End If
    Case KEY_EXIT
        Unload Me
'Error:
 ' MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  
  Exit Sub
End Select
End Sub

Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
      
    Case KEY_PRINT
        strCadena = "SELECT     Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, movimiento_caja.cPersona, " & _
        "movimiento_caja.descripcion_per, Persona.sDireccionCliente1, Persona.Per_Ruc, movimiento_caja.fecha_valor," & _
        "movimiento_caja.cambio , movimiento_caja.glosa, centro_costos.descripcion, movimiento_caja.Monto,movimiento_caja.monto_letras " & _
        "FROM movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
        "centro_costos ON movimiento_caja.id_costo = centro_costos.id_costo INNER JOIN " & _
        "Persona ON movimiento_caja.cPersona = Persona.cPersona WHERE movimiento_caja.serie='" & Trim(Me.TxtSerie.Text) & "' AND movimiento_caja.numero='" & Trim(Me.TxtNumeroDoc.Text) & "'"
        
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptReciboCaja", , App.Path + "\Reportes\")
    
  Exit Sub
Error:
  
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  
  Exit Sub
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
    Printer.Print Tab(15); (CVDate(Me.DtpActual.Value))
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); Mid(Me.TxtCliente.Text + Space(80), 1, 65)
    Printer.Print Tab(5); Mid(Me.txtdireccion.Text + Space(80), 1, 65)
    Printer.Print Tab(5); Mid(Me.TxtRuc.Text + Space(50), 1, 40) & "SALDINER"; Space(1); Mid(Me.TxtSerie.Text + Space(50), 1, 4) & Space(1) & "-" & Me.TxtNumeroDoc.Text
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); Mid(Trim(KEY_VENDEDOR), 1, 10)
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(15); "Monto Efectivo:" & "=============" & Space(20) & Me.TxtMontoIngresar.Text
    Printer.CurrentY = Printer.CurrentY + 10
    totalletras = UCase(EnLetras(Me.TxtMontoIngresar.Text))
    Set rst = Nothing
    '---- fin totales
    Printer.Print Tab(10); Mid(totalletras + Space(100), 1, 60)
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print Tab(60); Me.TxtMontoIngresar.Text
    Printer.EndDoc
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
    Exit Sub
End If
End Sub
Private Sub TxtCodCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If Len(Me.TxtCodCliente.Text) = 11 Or Len(Me.TxtCodCliente.Text) = 8 Then
    strCadena = "SELECT cPersona,NombrePersona,sDireccionCliente1,Per_Ruc,Observacion FROM " & _
    " Persona WHERE Per_Ruc='" & Trim(Me.TxtCodCliente.Text) & "'"
Else
    strCadena = "SELECT cPersona,NombrePersona,sDireccionCliente1,Per_Ruc,Observacion FROM " & _
    " Persona WHERE cPersona='" & Trim(Me.TxtCodCliente.Text) & "'"
End If
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        
        Me.TxtCodCliente.Text = rst(0)
        Me.TxtCliente.Text = rst(1)
        Me.txtdireccion.Text = rst(2)
        Me.TxtRuc.Text = rst(3)
        Me.txtobservacion.Text = rst(4)
        Call Resalta(Me.txtobservacion)
    Else
        Procedencia = Selecionar
        FrmPersona.Show
    End If
    
End If
End Sub
Private Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub

Private Sub TxtMontoIngresar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtMontoIngresar.Text = Format(Me.TxtMontoIngresar.Text, "#,##0.00")
End If
End Sub


Private Sub TxtNumeroDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Me.TxtNumeroDoc.Text = FormatosCeros(Me.TxtNumeroDoc.Text, 6)
    strCadena = "SELECT cDocumentoVenta,sSerie FROM DocumentoVenta WHERE (cDocumentoVenta='" & Trim(Me.TxtNumeroDoc.Text) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "')"
    Call ConfiguraRst(strCadena)
    
    If rst.RecordCount < 1 Then
       
        If CVDate(Me.DtpActual.Value) <> Date Then
            If MsgBox("la Fecha no Coincide con la Fecha del Documento...Desea Continuar", vbInformation + vbYesNo, KEY_EMPRESA) = vbYes Then
                Me.TxtCodCliente.Text = "00000"
                Call Resalta(Me.TxtCodCliente)
            Else
                Me.DtpActual.SetFocus
            End If
        Else
        Me.TxtCodCliente.Text = "00000"
        Call Resalta(Me.TxtCodCliente)
          End If
    Else
        MsgBox "Documento ya Existe", vbInformation, KEY_EMPRESA
       Call LlenarDatosCliente(Trim(Me.TxtNumeroDoc.Text), Trim(Me.DtcTipoDoc.BoundText), Trim(Me.TxtSerie.Text), Trim(Me.DtcAlmacen.BoundText))
       Call Resalta(Me.TxtNumeroDoc)
        
        
    End If
End If
Set rst = Nothing
End Sub

Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtMontoIngresar)
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
             strCadena = "SELECT numero FROM Det_alm_com WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "') ORDER BY numero DESC"
             Call ConfiguraRst(strCadena)
             Me.TxtNumeroDoc.Text = rst(0)
             Call Resalta(Me.TxtNumeroDoc)
             Set rst = Nothing
        End If
    Else
        MsgBox "Serie no Asiganda a a dicho Almacen", vbInformation, KEY_EMPRESA
    End If
End If

End Sub



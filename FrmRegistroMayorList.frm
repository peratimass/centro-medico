VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmRegistroMayorList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   16440
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRucEmpresa 
      Height          =   375
      Left            =   9240
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtFechaF 
      Height          =   495
      Left            =   1200
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox TxtDoccodNOta 
      Height          =   495
      Left            =   1200
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox TxtSerieNota 
      Height          =   495
      Left            =   1200
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox TxtNumeroNota 
      Height          =   495
      Left            =   1200
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox TxtTipoComprobante 
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
      Left            =   2160
      TabIndex        =   13
      Top             =   8520
      Width           =   495
   End
   Begin VB.TextBox TxtTipoCambio 
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
      Height          =   310
      Left            =   10120
      TabIndex        =   12
      Top             =   8520
      Width           =   450
   End
   Begin VB.TextBox TxtAnio 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   8520
      Width           =   535
   End
   Begin VB.TextBox txtdia 
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
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   8520
      Width           =   375
   End
   Begin VB.CommandButton cmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8400
      Width           =   375
   End
   Begin VB.TextBox txtafecto 
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
      Left            =   11820
      TabIndex        =   8
      Top             =   8520
      Width           =   1095
   End
   Begin VB.TextBox txtexonerado 
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
      Left            =   12940
      TabIndex        =   7
      Top             =   8520
      Width           =   900
   End
   Begin VB.TextBox txtIgv 
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
      Left            =   13880
      TabIndex        =   6
      Top             =   8520
      Width           =   900
   End
   Begin VB.TextBox txtTotal 
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
      Left            =   14810
      TabIndex        =   5
      Top             =   8520
      Width           =   1100
   End
   Begin VB.TextBox TxtCliente 
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
      Left            =   5685
      TabIndex        =   4
      Top             =   8520
      Width           =   3360
   End
   Begin VB.TextBox TxtRuc 
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
      Left            =   4425
      MaxLength       =   11
      TabIndex        =   3
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox TxtNumero 
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
      Left            =   3420
      TabIndex        =   2
      Top             =   8520
      Width           =   975
   End
   Begin VB.TextBox txtSerie 
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
      Left            =   2775
      TabIndex        =   1
      Top             =   8520
      Width           =   615
   End
   Begin VB.TextBox TxtAlmacen 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7800
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   12840
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroMayorList.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroMayorList.frx":0454
            Key             =   "(ImportarExcel)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroMayorList.frx":09EE
            Key             =   "(Cerrar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroMayorList.frx":0F88
            Key             =   "(Abrir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroMayorList.frx":1522
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroMayorList.frx":1842
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroMayorList.frx":1B5C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroMayorList.frx":1FB0
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroMayorList.frx":2404
            Key             =   "(RCompras)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroMayorList.frx":3B76
            Key             =   "(RVentas)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   7905
      Left            =   15360
      TabIndex        =   19
      Top             =   120
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   13944
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   7905
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
         TabIndex        =   20
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1561
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Nuevo"
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Anular"
               Key             =   "(Anular)"
               ImageKey        =   "(Anular)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Modificar"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Importar "
               Key             =   "(ImportarExcel)"
               ImageKey        =   "(ImportarExcel)"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Verificar"
               Key             =   "(Verificar)"
               ImageKey        =   "(Abrir)"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Actualizar"
               Key             =   "(Actualizar)"
               ImageKey        =   "(RCompras)"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo DtcMes 
      Height          =   315
      Left            =   525
      TabIndex        =   21
      Top             =   8520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
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
   Begin MSDataListLib.DataCombo DtcFormaPago 
      Height          =   315
      Left            =   10660
      TabIndex        =   22
      Top             =   8520
      Width           =   1120
      _ExtentX        =   1958
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   7335
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   12938
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
   Begin MSDataListLib.DataCombo DtcMoneda 
      Height          =   315
      Left            =   9075
      TabIndex        =   24
      Top             =   8520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
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
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "TC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   10120
      TabIndex        =   41
      Top             =   8280
      Width           =   195
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   9120
      TabIndex        =   40
      Top             =   8280
      Width           =   570
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "F.Pago"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   10660
      TabIndex        =   39
      Top             =   8280
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   1560
      TabIndex        =   38
      Top             =   8220
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Mes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   650
      TabIndex        =   37
      Top             =   8220
      Width           =   285
   End
   Begin VB.Label lblMes 
      BackStyle       =   0  'Transparent
      Caption         =   "LIBRO MAYOR:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   480
      TabIndex        =   36
      Top             =   240
      Width           =   4755
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Serie"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   2820
      TabIndex        =   35
      Top             =   8220
      Width           =   360
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3465
      TabIndex        =   34
      Top             =   8220
      Width           =   555
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ruc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4485
      TabIndex        =   33
      Top             =   8220
      Width           =   270
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   5685
      TabIndex        =   32
      Top             =   8220
      Width           =   495
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Afecto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   11820
      TabIndex        =   31
      Top             =   8220
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Exonerado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   12940
      TabIndex        =   30
      Top             =   8220
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "IGV"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   13880
      TabIndex        =   29
      Top             =   8220
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   14810
      TabIndex        =   28
      Top             =   8220
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "TD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   2160
      TabIndex        =   27
      Top             =   8220
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "Día"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   8220
      Width           =   225
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   9135
      Left            =   0
      Top             =   0
      Width           =   16440
   End
   Begin VB.Label lblEmpresa 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Ventas Mensual:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   5760
      TabIndex        =   25
      Top             =   240
      Width           =   9315
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   240
      Top             =   180
      Width           =   15015
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   795
      Left            =   0
      Top             =   8160
      Width           =   16395
   End
End
Attribute VB_Name = "FrmRegistroMayorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim igv As String, id_alm As String
Public Procedencia As EnumProcede
Dim correlativo As Double
Private Sub CmdAgregar_Click()
Dim fechaI As String, id_venta As Double
Dim Saldo As Single
Dim key_anulado2 As String
Static mostrar As Integer
key_anulado2 = KEY_ANULADO

If Me.DtcMoneda.BoundText = "00002" Then
    Me.txtafecto.Text = Val(Me.txtafecto.Text) * Val(Me.TxtTipoCambio.Text)
    Me.txtexonerado.Text = Val(Me.txtexonerado.Text) * Val(Me.TxtTipoCambio.Text)
    Me.TxtIgv.Text = Val(Me.TxtIgv.Text) * Val(Me.TxtTipoCambio.Text)
    Me.txttotal.Text = Val(Me.txttotal.Text) * Val(Me.TxtTipoCambio.Text)
End If

fechaI = Format(CVDate(Trim(Me.txtdia.Text) & "/" & Trim(Me.dtcmes.BoundText) & "/" & Trim(Me.txtanio.Text)), "YYYY-mm-dd")
If (Trim(Me.TxtTipoComprobante.Text) = "0001" And Len(Trim(Me.TxtRuc.Text)) <> 11 And KEY_ANULADO = "F") Then
    MsgBox "Ruc tiene que tener 11 Digitos", vbInformation, "Mensaje para el Usuario"
    Call Resalta(Me.TxtRuc)
    Exit Sub
End If

If (Val(Me.txtdia.Text) > 0 And Val(Me.txtdia.Text) < 32) Then
    If (Trim(Me.DtcFormaPago.BoundText) = "01") Then
        Saldo = 0
    Else
        Saldo = Val(Me.txttotal.Text)
    End If
    
If Trim(Me.cmdagregar.Caption) = "M" Then
    If Trim(Me.TxtTipoComprobante.Text) = "0003" Then
        Me.txtcliente.Text = ""
    End If
    fechaI = CVDate(Trim(Me.txtdia.Text) & "/" & Month(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "/" & Year(Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))))
    strCadena = "UPDATE RegistroVentasDetalle SET fecha='" & fechaI & "',doc_cod='" & Trim(Me.TxtTipoComprobante.Text) & "',serie='" & Trim(Me.TxtSerie.Text) & "', " & _
    "numero='" & Trim(Me.txtnumero.Text) & "',RucCliente='" & Trim(Me.TxtRuc.Text) & "',NombreCliente='" & Trim(Me.txtcliente.Text) & "',moneda='" & Trim(Me.DtcMoneda.BoundText) & "'," & _
    "afecto='" & Val(Me.txtafecto.Text) & "',exonerado='" & Val(txtexonerado.Text) & "',igv='" & Val(Me.TxtIgv.Text) & "',total='" & Val(Me.txttotal.Text) & "',anulado='F' WHERE codigounico='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 12)) & "'"
    CnBd.Execute (strCadena)
     
    If KEY_ANULADO = "V" Then
        strCadena = "UPDATE RegistroVentasDetalle SET anulado='V',NombreCliente='A N U L A D O',RucCliente='',afecto='0',exonerado='0',igv='0',total='0',saldo='0' WHERE codigounico='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 12)) & "'"
        CnBd.Execute (strCadena)
         
        KEY_ANULADO = "F"
        
        
   End If
   Procedencia = Neutro
    Call llenarGrid(Me.HfdPersona, Trim(Me.TxtRuc.Text))
    Me.txtnumero.Text = ""
    Me.TxtRuc.Text = ""
    Me.txtcliente.Text = ""
    Me.txtafecto.Text = 0#
    Me.txtexonerado.Text = 0#
    Me.TxtIgv.Text = 0#
    Me.txttotal.Text = 0#
    Me.cmdagregar.Caption = "+"
    Call Resalta(Me.TxtSerie)
    
   Exit Sub
 End If
  
   If Trim(Me.TxtTipoComprobante.Text) = "0020" Then
        strCadena = "P_insert_regventa('" & TxtTipoComprobante.Text & "','" & id_alm & "','" & Trim(Me.DtcFormaPago.BoundText) & "','" & Trim(Me.DtcMoneda.BoundText) & "','no'," & _
        "'" & Trim(Me.TxtSerie.Text) & "','" & Trim(Me.txtnumero.Text) & "','" & Me.TxtRuc.Text & "','" & Me.txtcliente.Text & "','0','0','0','0','" & Saldo & "'," & _
        "'" & Saldo & "','0','" & fechaI & "','" & fechaI & "','00001','" & KEY_USUARIO & "','" & Val(Me.TxtTipoCambio.Text) & "','no','" & formato_item(Month(fechaI), 2) & "','" & Year(fechaI) & "','','0','0','0','" & Val(Me.txttotal.Text) & "','" & Trim(Me.txtRucEmpresa.Text) & "')"
        CnBd.Execute (strCadena)
         
    ElseIf (Trim(Me.TxtTipoComprobante.Text) = "0007") Then
        strCadena = "P_insert_regventa('" & TxtTipoComprobante.Text & "','" & id_alm & "','" & Trim(Me.DtcFormaPago.BoundText) & "','" & Trim(Me.DtcMoneda.BoundText) & "','no'," & _
        "'" & Trim(Me.TxtSerie.Text) & "','" & Trim(Me.txtnumero.Text) & "','" & Me.TxtRuc.Text & "','" & Me.txtcliente.Text & "','" & Val(Me.txtafecto.Text) * -1 & "','" & Val(Me.TxtIgv.Text) * -1 & "','" & Val(Me.txtexonerado.Text) * -1 & "','" & Val(Me.txttotal.Text) * -1 & "','" & Saldo & "'," & _
        "'" & Saldo & "','0','" & fechaI & "','" & fechaI & "','00001','" & KEY_USUARIO & "','" & Val(Me.TxtTipoCambio.Text) & "','no','" & formato_item(Month(fechaI), 2) & "'," & _
        ",'" & Year(fechaI) & "','" & Format(CVDate(Me.TxtFechaF.Text), "YYYY-mm-dd") & "','" & Me.TxtDoccodNOta.Text & "','" & Trim(Me.TxtSerieNota.Text) & "','" & Trim(Me.TxtNumeroNota.Text) & "','" & Trim(Me.txtRucEmpresa.Text) & "')"
        CnBd.Execute (strCadena)
         
   Else
   strCadena = "P_insert_venta('" & TxtTipoComprobante.Text & "','" & id_alm & "','" & Trim(Me.DtcFormaPago.BoundText) & "','" & Trim(Me.DtcMoneda.BoundText) & "','no'," & _
        "'" & Trim(Me.TxtSerie.Text) & "','" & Trim(Me.txtnumero.Text) & "','" & Me.TxtRuc.Text & "','" & Me.txtcliente.Text & "','" & Val(Me.txtafecto.Text) & "','" & Val(Me.TxtIgv.Text) & "','" & Val(Me.txtexonerado.Text) & "','" & Val(Me.txttotal.Text) & "','" & Saldo & "'," & _
        "'" & Val(Me.txttotal.Text) & "','0','" & fechaI & "','" & fechaI & "','00001','" & KEY_USUARIO & "','" & Val(Me.TxtTipoCambio.Text) & "','no','" & formato_item(Month(fechaI), 2) & "','" & Year(fechaI) & "','" & Trim(Me.txtRucEmpresa.Text) & "')"
        CnBd.Execute (strCadena)
         
        
        id_venta = LastRegistroRUC("movimiento_venta", "id_venta")
        
        If Me.DtcFormaPago.BoundText = "01" Then
            forma_pago = "01"
        Else
            forma_pago = "08"
        End If
        strCadena = "P_insert_pagoventa('" & id_venta & "','" & forma_pago & "','" & Val(Me.txttotal.Text) & "','0','0','0','" & FrmRegistroMayor.TxtRuc.Text & "')"
        CnBd.Execute (strCadena)
         
   End If
   
   '-----ii
        
        
        StrNumero = FormatosCeros(Trim(str(Val(Me.txtnumero.Text)) + 1), 6)
        strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE id_alm='" & Trim(id_alm) & "' AND id_doc='" & Trim(TxtTipoComprobante.Text) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND ruc='" & Trim(FrmRegistroMayor.TxtRuc.Text) & "'"
        CnBd.Execute (strCadena)
         

   '------
  
   'MsgBox "CODIGO UNICO:" + Str(rst(0)), vbInformation, "Numero de Registro"
   
   If KEY_ANULADO = "si" Then
        strCadena = "UPDATE movimiento_venta SET anulado='si',ncliente='A N U L A D O',id_cliente='',afecto='0',exonerado='0',igv='0',total='0',saldo='0' WHERE id_venta='" & Val(rst(0)) & "'"
        CnBd.Execute (strCadena)
         
        KEY_ANULADO = "F"
        Me.txtcliente.Text = "A N U L A D O"
        Me.txtafecto.Text = 0#
        Me.txtexonerado.Text = 0#
        Me.TxtIgv.Text = 0#
        Me.txttotal.Text = 0#
   End If
   Set rst = Nothing
   
   '*******************
   If Trim(Me.TxtTipoComprobante.Text) = "0020" Then
        retencion = Val(Me.txttotal.Text)
    Else
        retencion = 0
   End If
      
      
 If correlativo < 1 Then
    Call formatearGrilla(Me.HfdPersona)
 End If
 If (Trim(Me.TxtTipoComprobante.Text) = "0007") Then
    Me.txtafecto.Text = Val(Me.txtafecto.Text) * -1
    Me.txtexonerado.Text = Val(Me.txtexonerado.Text) * -1
    Me.TxtIgv.Text = Val(Me.TxtIgv.Text) * -1
    Me.txttotal.Text = Val(Me.txttotal.Text) * -1
 End If
 Fila = str(correlativo + 1) & vbTab & fechaI & vbTab & Me.TxtTipoComprobante.Text & vbTab & Me.TxtSerie.Text & vbTab & Me.txtnumero.Text & vbTab & Me.TxtRuc.Text & vbTab & Me.txtcliente.Text & vbTab & Format(Val(Me.txtafecto.Text), "#,##0.00") & vbTab & Format(Val(Me.txtexonerado.Text), "#,##0.00") & vbTab & Format(Val(Me.TxtIgv.Text), "#,##0.00") & vbTab & Format(Val(Me.txttotal.Text), "#,##0.00") & vbTab & Format(Val(retencion), "#,##0.00") & vbTab & str(codigo_unico)
 correlativo = correlativo + 1
 Me.HfdPersona.AddItem Fila
  If key_anulado2 = "si" Then
        For k = 0 To 11
            HfdPersona.col = k
            HfdPersona.Row = correlativo
            HfdPersona.CellBackColor = &H8080FF
        Next k
 End If
 strCadena = "SELECT SUM(valor_venta), SUM(exonerado) , SUM(igv) , SUM(total) , SUM(retencion) FROM  movimiento_venta WHERE ruc='" & Trim(Me.txtRucEmpresa.Text) & "' AND id_mes='" & Trim(Me.dtcmes.BoundText) & "' AND id_anio='" & Trim(Me.txtanio.Text) & "' AND anulado='no' "
 Call ConfiguraRst(strCadena)
      
   
 Call nuevo
End If
End Sub
Private Sub formatearGrilla(ByVal Grilla As MSHFlexGrid)
         Grilla.Clear

   Grilla.Rows = 0

       ReDim arrColWidth(1 To rst.Fields.Count)
       For i = 0 To 0

           Grilla.ColWidth(0) = 600
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 600
           Grilla.ColWidth(3) = 600
           Grilla.ColWidth(4) = 1000
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 3800
           Grilla.ColWidth(7) = 1200
           Grilla.ColWidth(8) = 1100
           Grilla.ColWidth(9) = 1200
           Grilla.ColWidth(10) = 1450
           Grilla.ColWidth(11) = 1000
           Grilla.ColWidth(12) = 0
           
        Next i
         cabecera = "ITEM" & vbTab & "FECHA" & vbTab & "TD" & vbTab & "SERIE" & vbTab & "NUMERO" & vbTab & "RUC" & vbTab & "CLIENTE" & vbTab & "AFECTO" & vbTab & "EXONERADO" & vbTab & "IGV" & vbTab & "TOTAL" & vbTab & "RETENCION" & vbTab & "CODIGO"
        Grilla.AddItem cabecera
         For k = 0 To 12
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
End Sub
Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal ruc As String)
On Error GoTo salir
Dim debes As Double, habers As Double

strCadena = "SELECT * FROM registro_diario_detalle WHERE ruc='" & Trim(ruc) & "' AND id_mes='" & Me.dtcmes.BoundText & "' AND id_anio='" & Trim(Me.txtanio.Text) & "'  ORDER BY fecha,num_correlativo ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
 N = 1
   Grilla.Clear
   Grilla.Rows = 0

       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields

           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 1300
           Grilla.ColWidth(2) = 4500
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 4000
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1200
           
        Next
        cabecera = "CORRELATIVO" & vbTab & "FECHA OPERACION" & vbTab & "GLOSA O DESCRIPCION OPERACION" & vbTab & "CODIGO" & vbTab & "DENOMINACION" & vbTab & "DEBE" & vbTab & "HABER"
        Grilla.AddItem cabecera
         For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
        debes = 0
        habers = 0
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("num_correlativo") & vbTab & rst("fecha") & vbTab & rst("glosa") & vbTab & rst("id_cuenta") & vbTab & rst("denominacion") & vbTab & Format(rst("debe"), "#,##0.00") & vbTab & Format(rst("haber"), "#,##0.00")
            If rst("debe") > 0 Then
                debes = debes + rst("debe")
            End If
            If rst("haber") > 0 Then
                habers = habers + rst("haber")
            End If
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
             
        Next i
       Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & " ********* SALDOS ********* " & vbTab & Format(debes, "#,##0.00") & vbTab & Format(habers, "#,##0.00")
       Grilla.AddItem Fila
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Private Sub nuevo()
Me.cmdagregar.Caption = "+"
Me.txtafecto.Text = 0#
Me.txtexonerado.Text = 0#
Me.TxtIgv.Text = 0#
Me.txttotal.Text = 0#
Me.TxtRuc.Text = ""
Me.txtcliente.Text = ""
Me.TxtDoccodNOta.Text = ""
Me.TxtSerieNota.Text = ""
Me.TxtNumeroNota.Text = ""
'Me.txtSerie.Text = ""
Me.txtnumero.Text = formato_item(Val(Me.txtnumero.Text) + 1, 10)
Call Resalta(Me.txtnumero)
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdAgregar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
    Call Resalta(Me.txttotal)
End If
End Sub

Private Sub DtcFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
    Me.DtcMoneda.SetFocus
End If
End Sub

Private Sub DtcFormaPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (Trim(Me.TxtTipoComprobante) = "0001") Then
        Call Resalta(Me.txtafecto)
    Else
        Call Resalta(Me.txttotal)
    End If
End If
End Sub

Private Sub DtcMoneda_Change()
If Me.DtcMoneda.BoundText = "0002" Then
     Me.Label14.Visible = True
     Me.TxtTipoCambio.Visible = True
     strCadena = "SELECT * FROM tipo_cambio WHERE fecha='" & KEY_FECHA & "' and id_creador='" & KEY_RUC & "'"
     Call ConfiguraRst(strCadena)
     Me.TxtTipoCambio.Text = rst("valor")
     Set rst = Nothing
     Call Resalta(Me.TxtTipoCambio)
Else
    Me.Label14.Visible = False
    Me.TxtTipoCambio.Visible = False
    
End If
End Sub

Private Sub DtcMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Me.DtcMoneda.BoundText = "0002" Then
    Call Resalta(Me.TxtTipoCambio)
Else
    Me.DtcFormaPago.SetFocus
End If
End Sub

Private Sub DtComprobante_Change()
If (Trim(Me.TxtTipoComprobante) = "0003") Then
    Me.Label10.Caption = "DNI"
Else
    Me.Label10.Caption = "Ruc"
End If
End Sub

Private Sub DtComprobante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtSerie)
End If
End Sub
Public Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
correlativo = 0
strCadena = "SELECT * FROM entidad_parametros WHERE cod_unico='" & Trim(FrmRegistroMayor.TxtRuc.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    igv = rst("igv")
    id_alm = rst("id_alm")
End If
Set rst = Nothing
Me.LblEmpresa.Caption = FrmRegistroMayor.HfdPersona.TextMatrix(FrmRegistroMayor.HfdPersona.Row, 6) + "***[" + Space(1) + FrmRegistroMayor.HfdPersona.TextMatrix(FrmRegistroMayor.HfdPersona.Row, 0) + Space(1) + "]***"
Me.lblMes.Caption = FrmRegistroMayor.HfdPersona.TextMatrix(FrmRegistroMayor.HfdPersona.Row, 2) + Space(2) + FrmRegistroMayor.HfdPersona.TextMatrix(FrmRegistroMayor.HfdPersona.Row, 3)
Me.txtRucEmpresa.Text = FrmRegistroMayor.HfdPersona.TextMatrix(FrmRegistroMayor.HfdPersona.Row, 0)
  
strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM forma_pago ORDER BY id ASC "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFormaPago)
Me.DtcFormaPago.BoundText = "01"
  
  strCadena = "SELECT id_mes as Codigo, descripcion as Descripcion FROM meses " & _
  " ORDER BY id_mes ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.dtcmes)
  
    strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda ORDER BY id_moneda ASC "
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcMoneda)


  Me.dtcmes.BoundText = FrmRegistroMayor.HfdPersona.TextMatrix(FrmRegistroMayor.HfdPersona.Row, 1)
  Me.dtcmes.Enabled = False
  Me.txtanio.Text = FrmRegistroMayor.HfdPersona.TextMatrix(FrmRegistroMayor.HfdPersona.Row, 3)
  Call llenarGrid(Me.HfdPersona, Trim(txtRucEmpresa.Text))
  
  Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
  Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Me.TxtSerie.Text = formato_item(Me.TxtSerie.Text, 3)
    Call Resalta(Me.txtnumero)


End If
End Sub

Private Sub HfdPersona_Click()
If Me.HfdPersona.Rows > 0 Then
      Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
      Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True
End If
End Sub

Private Sub HfdPersona_GotFocus()
HookForm Me.HfdPersona
End Sub

Private Sub HfdPersona_LostFocus()
UnHookForm Me.HfdPersona
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case KEY_NEW
        Call nuevo
        Exit Sub
    Case KEY_ANULAR
       If MsgBox("Esta Seguro de Anular este Comprobante", vbQuestion + vbYesNo, "Informacion para el Usuario") = vbYes Then
        strCadena = "UPDATE movimiento_venta SET anulado='si',ncliente='A N U L A D O',valor_venta='0',igv='0',exonerado='0',total='0',saldo='0' WHERE id_venta='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 12)) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        Call llenarGrid(Me.HfdPersona, Trim(Me.txtRucEmpresa.Text))
       End If
    Case KEY_UPDATE
        If MsgBox("Esta seguro de Modificar este Registro", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
            strCadena = "SELECT * FROM RegistroVentasDetalle WHERE  codigounico='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 12)) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                Procedencia = Modificar
                Me.txtdia.Text = Day(rst("fecha"))
                Me.TxtTipoComprobante.Text = rst("doc_cod")
                Me.TxtSerie.Text = rst("serie")
                Me.txtnumero.Text = rst("numero")
                Me.TxtRuc.Text = rst("RucCliente")
                Me.txtcliente.Text = rst("NombreCliente")
                Me.DtcMoneda.BoundText = rst("moneda")
                Me.TxtTipoCambio.Text = rst("tc")
                Me.DtcFormaPago.BoundText = rst("idFormaPago")
                Me.txtafecto.Text = rst("afecto")
                Me.txtexonerado.Text = rst("exonerado")
                Me.TxtIgv.Text = rst("igv")
                Me.txttotal.Text = rst("total")
                Me.cmdagregar.Caption = "M"
                Exit Sub
            End If
        End If
    Case KEY_DELETE
      If MsgBox("Esta Seguro de eliminar este registro", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        strCadena = "DELETE FROM RegistroVentasDetalle WHERE  codigounico='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 12)) & "'"
        CnBd.Execute (strCadena)
         
        'Me.HfdPersona.RemoveItem (Me.HfdPersona.TabIndex)
        Call llenarGrid(Me.HfdPersona, Trim(Me.txtRucEmpresa.Text))
      End If
      
      Case KEY_IMPEXCEL
            Dim ruta As String
            Me.CommonDialog1.Filter = "*.xlsx OR *.xls"
            Me.CommonDialog1.ShowOpen
            ruta = Me.CommonDialog1.FileName
            'Inicio (ruta)
            'Caracteristicas
            'FormatearTabla
            'Call LlenadoDeTabla(FrmRegistromayor.TxtRuc.text, Me.DtcMes.BoundText, Me.TxtAnio.text)
    Case "(Verificar)"
        FrmRegistromayorDetalleBuscar.Show
    Case KEY_ACTUALIZAR
        Call llenarGrid(Me.HfdPersona, Trim(Me.txtRucEmpresa.Text))
    Case KEY_EXIT
        Call FrmRegistroMayor.actualizar
        Unload Me
  End Select
End Sub

Private Sub txtafecto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
    Me.DtcFormaPago.SetFocus
End If
End Sub

Private Sub txtafecto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Me.txtafecto.Text) > 0 Then
        
        Me.TxtIgv.Text = Format(Val(Me.txtafecto.Text) * (KEY_IGV), "###0.000")
    End If
    Me.txtafecto.Text = Format(Val(Me.txtafecto.Text), "###0.000")
    Call Resalta(Me.txtexonerado)
End If
End Sub

Private Sub TxtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
    Call Resalta(Me.TxtRuc)
End If
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Trim(Me.TxtTipoComprobante) = "0001" And Me.txtcliente.Text = "" Then
        MsgBox "Es Obligatorio el Ruc del Cliente"
    Call Resalta(Me.TxtRuc)
ElseIf (Trim(Me.txtcliente.Text) = "") Then
     If (Me.TxtTipoComprobante) = "0003" Then
        Me.DtcMoneda.SetFocus
        Exit Sub
    End If
    Call Resalta(Me.txtafecto)
Else
    If (Me.TxtTipoComprobante) = "0003" Then
        Me.DtcMoneda.SetFocus
        Exit Sub
    End If
    
   If Trim(Me.TxtTipoComprobante.Text) = "0001" And igv = "no" Then
        Call Resalta(Me.txtexonerado)
        Exit Sub
    ElseIf Trim(Me.TxtTipoComprobante.Text) = "0001" And igv = "si" Then
        Call Resalta(Me.txtafecto)
        Exit Sub
   End If
    
    
    Call Resalta(Me.txtafecto)
    Exit Sub
    'Me.DtcMoneda.SetFocus
End If
End If
End Sub

Private Sub txtdia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Me.txtdia.Text) > 31 Or Val(Me.txtdia.Text) < 1 Then
        MsgBox "Ingrese un Valor Valido", vbInformation, "Mensaje para el Usuario"
        Call Resalta(Me.txtdia)
        Exit Sub
    End If
    Me.txtdia.Text = formato_item(Me.txtdia.Text, 2)
    Call Resalta(TxtTipoComprobante)
End If
End Sub

Private Sub txtexonerado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
    Call Resalta(Me.txtafecto)
End If
End Sub

Private Sub txtexonerado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtexonerado.Text = Format(Val(Me.txtexonerado.Text), "###0.000")
    If Val(Me.txtexonerado.Text) = Val(Me.txttotal.Text) And Val(Me.txttotal.Text) <> 0 Then
        Me.txtafecto.Text = 0#
        Me.TxtIgv.Text = 0#
        Call Resalta(Me.txttotal)
        Exit Sub
    Else
        
       If (Trim(Me.TxtTipoComprobante.Text) = "0001" And igv = "no") Then
            Me.txttotal.Text = Format(Me.txtexonerado.Text, "###0.00")
            Call Resalta(Me.txttotal)
            Exit Sub
       End If
        
       ' Me.txtIgv.Text = Format(Val(Me.txtTotal.Text) - Val(Me.txtIgv.Text) - Val(Me.txtafecto.Text), "#,##0.00")
        Me.txttotal.Text = Format(Val(Me.txtafecto.Text) + Val(Me.txtexonerado.Text) + Val(Me.TxtIgv.Text), "#,##0.00")
    End If
    
    Call Resalta(Me.TxtIgv)
End If
End Sub

Private Sub txtIgv_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
    Call Resalta(Me.txtexonerado)
End If
End Sub

Private Sub TxtIGV_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtIgv.Text = Format(Round(Val(Me.TxtIgv.Text), 3), "###0.000")
    Call Resalta(Me.txttotal)
End If
End Sub

Private Sub TxtNumero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
    Call Resalta(Me.TxtSerie)
End If
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtnumero.Text = formato_item(Me.txtnumero.Text, 6)
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & Trim(Me.TxtTipoComprobante) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Trim(Me.txtnumero.Text) & "' AND ruc='" & Trim(FrmRegistroMayor.TxtRuc.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        MsgBox "Comprobante ya Registrado", vbInformation, "Duplicidad de Comprobante"
        Call Resalta(Me.txtnumero)
        Exit Sub
    End If
     FrmVerificarAnuladoVentas.Show
     
     Exit Sub
    'Call Resalta(Me.tXTrUC)
End If
End Sub

Private Sub TxtRuc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
    Call Resalta(Me.txtnumero)
End If
End Sub

Private Sub txtruc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.TxtRuc.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.txtcliente.Text = rst("nombre_completo")
    Else
                
        If Len(Trim(Me.TxtRuc.Text)) = 11 Then
            Procedencia = 1
            FrmDetallePersona.Show
            FrmDetallePersona.TxtRuc.Text = Trim(Me.TxtRuc.Text)
            FrmDetallePersona.ChkCliente.Value = 1
            Call FrmDetallePersona.precionar
            Exit Sub
         Else
             
                    Procedencia = buscar
                    FrmPersona.Show
                    Exit Sub
             
        End If
    End If
    Call Resalta(Me.txtcliente)
End If
End Sub
Public Sub foco_ruc()
Call Resalta(Me.TxtRuc)
End Sub

Private Sub TxtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
    Call Resalta(Me.TxtTipoComprobante)
End If
End Sub

Private Sub TxtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie.Text = formato_item(Me.TxtSerie.Text, 3)
    strCadena = "SELECT * FROM movimiento_venta WHERE ruc='" & Trim(FrmRegistroMayor.TxtRuc.Text) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND id_doc='" & Trim(Me.TxtTipoComprobante.Text) & "' ORDER BY numero DESC"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Me.txtnumero.Text = rst("numero") + 1
    Else
        Me.txtnumero.Text = 1
    End If
    Me.txtnumero.Text = formato_item(Me.txtnumero.Text, 6)
    Call Resalta(Me.txtnumero)
End If
End Sub
Public Sub foco_serie()
Call Resalta(Me.TxtSerie)
End Sub

Private Sub TxtTipoCambio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      Me.DtcFormaPago.SetFocus
End If
End Sub

Private Sub TxtTipoComprobante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
    Call Resalta(Me.txtdia)
End If
End Sub

Private Sub TxtTipoComprobante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.TxtTipoComprobante.Text = "" Then
        Procedencia = buscar
        FrmComprobantes.Show
    Else
   Me.TxtTipoComprobante.Text = formato_item(Me.TxtTipoComprobante.Text, 4)
   Call Resalta(Me.TxtSerie)
     If Trim(TxtTipoComprobante.Text) = "0007" Then
        FrmNotaCredito.Show
    End If
    End If
End If
End Sub

Private Sub txtTotal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 37 Then
    Call Resalta(Me.TxtIgv)
End If
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Me.txttotal.Text) = 0 Then
       If MsgBox("Desea Ingresar valor 0.00 Para este Comprobante", vbQuestion + vbYesNo, "Alerta") = vbNo Then
       Call Resalta(Me.txttotal)
       Exit Sub
       End If
    End If

        If Trim(Me.TxtTipoComprobante) = "0003" Then
            If igv = "si" And Val(Me.TxtIgv.Text) = 0 And Val(Me.txtexonerado.Text) = 0 Then
                Me.txttotal.Text = Format(Val(Me.txttotal.Text), "###0.000")
                Me.txtafecto.Text = Format(Round(Val(Me.txttotal.Text) / (KEY_IGV + 1), 3), "###0.000")
                Me.TxtIgv.Text = Format(Round(Val(txtafecto) * (KEY_IGV), 3), "###0.000")
                Me.cmdagregar.SetFocus
                Exit Sub
            ElseIf (igv = "si" And Val(Me.TxtIgv.Text) >= 0) Then
               
                Me.txttotal.Text = Format(Val(Me.txttotal.Text), "###0.000")
                Me.cmdagregar.SetFocus
                Exit Sub
            ElseIf (igv = "no") Then
                 Me.txtexonerado.Text = Format(Val(Me.txttotal.Text), "###0.000")
                 Me.txttotal.Text = Format(Val(Me.txttotal.Text), "###0.000")
                 Me.cmdagregar.SetFocus
                Exit Sub
            End If
        End If
       Me.cmdagregar.SetFocus
End If
End Sub









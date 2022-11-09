VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmReporteVentas 
   Caption         =   "Reporte de Ventas"
   ClientHeight    =   4485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   8010
   Begin TabDlg.SSTab SstKardex 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Venta"
      TabPicture(0)   =   "FrmReporteVentas1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblCantidad"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DtpHasta"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DtpDesde"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DtcLaboratorio"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DtcTipoProducto"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DtcLinea"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DtcProducto"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkLaboratorio"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ChkClinea"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ChkCtiprodu"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkProducto"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.CheckBox chkProducto 
         BackColor       =   &H80000004&
         Caption         =   "Por Producto"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   225
         TabIndex        =   4
         Top             =   1740
         Width           =   1335
      End
      Begin VB.CheckBox ChkCtiprodu 
         BackColor       =   &H80000004&
         Caption         =   "Marca"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   225
         TabIndex        =   3
         Top             =   3510
         Width           =   1335
      End
      Begin VB.CheckBox ChkClinea 
         BackColor       =   &H80000004&
         Caption         =   "Línea"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   225
         TabIndex        =   2
         Top             =   2940
         Width           =   1095
      End
      Begin VB.CheckBox chkLaboratorio 
         BackColor       =   &H80000004&
         Caption         =   "Proveedor"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   225
         TabIndex        =   1
         Top             =   2370
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DtcProducto 
         Height          =   315
         Left            =   1665
         TabIndex        =   5
         Top             =   1710
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcLinea 
         Height          =   315
         Left            =   1950
         TabIndex        =   6
         Top             =   2910
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcTipoProducto 
         Height          =   315
         Left            =   1950
         TabIndex        =   7
         Top             =   3480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcLaboratorio 
         Height          =   315
         Left            =   1950
         TabIndex        =   8
         Top             =   2340
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   840
         TabIndex        =   9
         Top             =   915
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   102236161
         CurrentDate     =   37091
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   315
         Left            =   3360
         TabIndex        =   10
         Top             =   915
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   102236161
         CurrentDate     =   37091
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         FillColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   600
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label LblCantidad 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al:"
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
         Left            =   2775
         TabIndex        =   12
         Top             =   960
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rango de Fechas:"
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
         Left            =   705
         TabIndex        =   11
         Top             =   400
         Width           =   1695
      End
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   1785
      Left            =   6840
      TabIndex        =   13
      Top             =   615
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   3149
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   1785
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   2340
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   4128
         ButtonWidth     =   1667
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Aceptar   "
               Key             =   "(Aceptar)"
               Object.ToolTipText     =   "Aceptar"
               ImageKey        =   "(Aceptar)"
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
      Left            =   6720
      Top             =   2655
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
            Picture         =   "FrmReporteVentas1.frx":001C
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas1.frx":0470
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas1.frx":0790
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas1.frx":0BE4
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas1.frx":1038
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas1.frx":1358
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas1.frx":17AC
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas1.frx":1908
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas1.frx":1D5C
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas1.frx":2078
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas1.frx":2954
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas1.frx":2C74
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas1.frx":2F94
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReporteVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call FormReport(Me)
End Sub

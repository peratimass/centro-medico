VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmReporteGanancia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ganancias"
   ClientHeight    =   3270
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7230
   Begin TabDlg.SSTab SstKardex 
      Height          =   2775
      Left            =   165
      TabIndex        =   0
      Top             =   225
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Producto"
      TabPicture(0)   =   "FrmReporteGanancia.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chkProducto"
      Tab(0).Control(1)=   "ChkCtiprodu"
      Tab(0).Control(2)=   "ChkClinea"
      Tab(0).Control(3)=   "chkLaboratorio"
      Tab(0).Control(4)=   "DtcProducto"
      Tab(0).Control(5)=   "DtcLinea"
      Tab(0).Control(6)=   "DtcTipoProducto"
      Tab(0).Control(7)=   "DtcLaboratorio"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Fechas"
      TabPicture(1)   =   "FrmReporteGanancia.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblDesde"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "DtpHasta"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DtpDesde"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ChkFecha"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CheckBox chkProducto 
         BackColor       =   &H80000004&
         Caption         =   "Por Producto"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74775
         TabIndex        =   5
         Top             =   660
         Width           =   1335
      End
      Begin VB.CheckBox ChkFecha 
         BackColor       =   &H80000004&
         Caption         =   "Por Rango de Fechas"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox ChkCtiprodu 
         BackColor       =   &H80000004&
         Caption         =   "Tipo producto"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74775
         TabIndex        =   3
         Top             =   2070
         Width           =   1335
      End
      Begin VB.CheckBox ChkClinea 
         BackColor       =   &H80000004&
         Caption         =   "Línea"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74775
         TabIndex        =   2
         Top             =   1620
         Width           =   1095
      End
      Begin VB.CheckBox chkLaboratorio 
         BackColor       =   &H80000004&
         Caption         =   "Proveedor"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74775
         TabIndex        =   1
         Top             =   1170
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DtcProducto 
         Height          =   315
         Left            =   -73335
         TabIndex        =   6
         Top             =   630
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   1275
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   62259201
         CurrentDate     =   37091
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   315
         Left            =   3480
         TabIndex        =   8
         Top             =   1275
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   62259201
         CurrentDate     =   37091
      End
      Begin MSDataListLib.DataCombo DtcLinea 
         Height          =   315
         Left            =   -73050
         TabIndex        =   9
         Top             =   1590
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcTipoProducto 
         Height          =   315
         Left            =   -73050
         TabIndex        =   10
         Top             =   2040
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcLaboratorio 
         Height          =   315
         Left            =   -73050
         TabIndex        =   11
         Top             =   1140
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Hasta : "
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2880
         TabIndex        =   13
         Top             =   1335
         Width           =   555
      End
      Begin VB.Label lblDesde 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Desde : "
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   1335
         Width           =   600
      End
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   1785
      Left            =   6165
      TabIndex        =   14
      Top             =   600
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
         TabIndex        =   15
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
      Left            =   6045
      Top             =   2640
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
            Picture         =   "FrmReporteGanancia.frx":0038
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteGanancia.frx":048C
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteGanancia.frx":07AC
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteGanancia.frx":0C00
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteGanancia.frx":1054
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteGanancia.frx":1374
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteGanancia.frx":17C8
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteGanancia.frx":1924
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteGanancia.frx":1D78
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteGanancia.frx":2094
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteGanancia.frx":2970
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteGanancia.frx":2C90
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteGanancia.frx":2FB0
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmReporteGanancia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Linea As String, Laboratorio As String, TipoProducto As String

Private Sub ChkClinea_Click()
  If ChkClinea.Value = 1 Then
    DtcLinea.Enabled = True
  Else
    DtcLinea.Enabled = False
    strCadena = "SELECT ctipoproducto as Codigo,sdescripcion as Descripcion FROM " & _
    " tipoproducto ORDER BY sdescripcion"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(DtcTipoProducto)
    DtcTipoProducto.Enabled = False
  End If
End Sub

Private Sub ChkCtiprodu_Click()
  If ChkCtiprodu.Value = 1 Then
    DtcTipoProducto.Enabled = True
  Else
    DtcTipoProducto.Enabled = False
  End If
End Sub

Private Sub ChkFecha_Click()
  If ChkFecha.Value = 1 Then
    DtpDesde.Enabled = True
    DtpHasta.Enabled = True
  Else
    DtpDesde.Enabled = False
    DtpHasta.Enabled = False
  End If
End Sub

Private Sub chkLaboratorio_Click()
  If chkLaboratorio.Value = 1 Then
    DtcLaboratorio.Enabled = True
  Else
    DtcLaboratorio.Enabled = False
  End If
End Sub

Private Sub ChkProducto_Click()
  If ChkProducto.Value = 1 Then
    DtcProducto.Enabled = True
  Else
    DtcProducto.Enabled = False
  End If
End Sub

Private Sub DtcLinea_Click(Area As Integer)
  Linea = Replace(DtcLinea.BoundText, "'", "''")
  strCadena = "SELECT ctipoproducto as Codigo,sdescripcion as Descripcion FROM " & _
  " tipoproducto WHERE clinea = '" & Linea & "' ORDER BY sdescripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(DtcTipoProducto)
  DtcTipoProducto.Enabled = False
End Sub

Private Sub Form_Load()
Call FormReport(Me)
  strCadena = "SELECT cproducto as Codigo, sdescripcionproducto as Descripcion " & _
  " FROM producto ORDER BY sdescripcionproducto"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(DtcProducto)
  
  strCadena = "SELECT clinea as Codigo, sdescripcion as Descripcion FROM " & _
  " linea ORDER BY sdescripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(DtcLinea)
  
  strCadena = "SELECT ctipoproducto as Codigo, sdescripcion as Descripcion FROM " & _
  " tipoproducto ORDER BY sdescripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(DtcTipoProducto)
  
  strCadena = "SELECT claboratorio as Codigo, srazonsocial as Descripcion FROM " & _
  " laboratorio ORDER BY srazonsocial"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(DtcLaboratorio)
    
  DtpDesde.Value = Date
  DtpHasta.Value = Date
  DtpDesde.Enabled = False
  DtpHasta.Enabled = False
  DtcProducto.Enabled = False
  DtcLinea.Enabled = False
  DtcTipoProducto.Enabled = False
  DtcLaboratorio.Enabled = False
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim DteInicio As Date, DteFin As Date
Dim StrProducto As String
Dim Ans As Boolean
Select Case Button.Key
  Case KEY_OK
    DteInicio = CDate(DTEMINIMA)
    DteFin = CDate(DTEMAXIMA)
    Linea = ""
    Laboratorio = ""
    TipoProducto = ""
    StrProducto = ""
    If ChkFecha.Value = 1 Then
      DteInicio = DtpDesde.Value
      DteFin = DtpHasta.Value
    End If
    If ChkProducto.Value = 1 Then
      StrProducto = Replace(DtcProducto.BoundText, "'", "''")
    End If
    If ChkClinea.Value = 1 Then
      Linea = Replace(DtcLinea.BoundText, "'", "''")
    End If
    If ChkCtiprodu.Value = 1 Then
      TipoProducto = Replace(DtcTipoProducto.BoundText, "'", "''")
    End If
    If chkLaboratorio.Value = 1 Then
      Laboratorio = Replace(DtcLaboratorio.BoundText, "'", "''")
    End If
    strCadena = "SELECT DV.demisionventa, P.sdescripcionproducto, V.ncantidad, " & _
    " V.nprecioventa, C.npreciocompra FROM DocumentoVenta as DV " & _
    " INNER JOIN ((Producto as P INNER JOIN DetalleDocumentoVenta as V ON P.cProducto = " & _
    " V.cProducto) INNER JOIN (FacturaCompra as FC INNER JOIN DetalleFacturaCompra as " & _
    " C ON (FC.cDistribuidora = C.cDistribuidora) AND (FC.cFactura = C.cFactura)) ON " & _
    " P.cProducto = C.cProducto) ON DV.cDocumentoVenta = V.cDocumentoVenta WHERE " & _
    " P.cproducto LIKE '" & StrProducto & "%' AND clinea LIKE '" & Linea & "%' AND " & _
    " ctipoproducto LIKE '" & TipoProducto & "%' AND claboratorio LIKE " & _
    " '" & Laboratorio & "%' AND demisionventa >= cdate('" & DteInicio & "') AND " & _
    " demisionventa <= cdate('" & DteFin & "') ORDER BY demisionventa, sdescripcionproducto"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptGanancia", , App.Path + "\Reportes\")
  Case KEY_CANCEL
    Unload Me
End Select
End Sub



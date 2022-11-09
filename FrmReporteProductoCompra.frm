VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmReporteProductoCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos por Comprar"
   ClientHeight    =   2220
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   7215
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   7215
   Begin TabDlg.SSTab SstRepProducto 
      Height          =   1800
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3175
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Clasificación"
      TabPicture(0)   =   "FrmReporteProductoCompra.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DtcTipoProducto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DtcLinea"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ChkCtiprodu"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ChkClinea"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Proveedor y Distribuidora"
      TabPicture(1)   =   "FrmReporteProductoCompra.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ChkDistribuidora"
      Tab(1).Control(1)=   "chkLaboratorio"
      Tab(1).Control(2)=   "DtcLaboratorio"
      Tab(1).Control(3)=   "DtcDistribuidora"
      Tab(1).ControlCount=   4
      Begin VB.CheckBox ChkDistribuidora 
         BackColor       =   &H80000004&
         Caption         =   "Distribuidora"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74640
         TabIndex        =   9
         Top             =   750
         Width           =   1335
      End
      Begin VB.CheckBox ChkClinea 
         BackColor       =   &H80000004&
         Caption         =   "Línea"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   765
         Width           =   1095
      End
      Begin VB.CheckBox chkLaboratorio 
         BackColor       =   &H80000004&
         Caption         =   "Proveedor"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74640
         TabIndex        =   2
         Top             =   1245
         Width           =   1335
      End
      Begin VB.CheckBox ChkCtiprodu 
         BackColor       =   &H80000004&
         Caption         =   "Tipo producto"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   1245
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DtcLinea 
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   735
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcLaboratorio 
         Height          =   315
         Left            =   -73065
         TabIndex        =   5
         Top             =   1215
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
         Left            =   2280
         TabIndex        =   6
         Top             =   1215
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcDistribuidora 
         Height          =   315
         Left            =   -73065
         TabIndex        =   10
         Top             =   720
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   1785
      Left            =   6105
      TabIndex        =   7
      Top             =   233
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
         TabIndex        =   8
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
      Left            =   5745
      Top             =   1545
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
            Picture         =   "FrmReporteProductoCompra.frx":0038
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteProductoCompra.frx":048C
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteProductoCompra.frx":07AC
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteProductoCompra.frx":0C00
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteProductoCompra.frx":1054
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteProductoCompra.frx":1374
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteProductoCompra.frx":17C8
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteProductoCompra.frx":1924
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteProductoCompra.frx":1D78
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteProductoCompra.frx":2094
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteProductoCompra.frx":2970
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteProductoCompra.frx":2C90
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteProductoCompra.frx":2FB0
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmReporteProductoCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Linea As String, Laboratorio As String, TiProdu As String, Distribuidora As String

Private Sub ChkClinea_Click()
  If ChkClinea.Value = 1 Then
    DtcLinea.Enabled = True
  Else
    DtcLinea.Enabled = False
    StrCadena = "SELECT ctipoproducto as Codigo,sdescripcion as Descripcion FROM " & _
    " tipoproducto ORDER BY sdescripcion"
    Call ConfiguraRst(StrCadena)
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

Private Sub ChkDistribuidora_Click()
  If ChkDistribuidora.Value = 1 Then
    DtcDistribuidora.Enabled = True
  Else
    DtcDistribuidora.Enabled = False
    StrCadena = "SELECT laboratorio.claboratorio as Codigo, snombrecorto as Descripcion FROM " & _
    " laboratorio ORDER BY snombrecorto"
    Call ConfiguraRst(StrCadena)
    Call LlenaDataCombo(DtcLaboratorio)
    DtcLaboratorio.Enabled = False
  End If
End Sub

Private Sub chkLaboratorio_Click()
  If chkLaboratorio.Value = 1 Then
    DtcLaboratorio.Enabled = True
  Else
    DtcLaboratorio.Enabled = False
  End If
End Sub

Private Sub DtcDistribuidora_Click(Area As Integer)
  Distribuidora = Replace(DtcDistribuidora.BoundText, "'", "''")
  StrCadena = "SELECT laboratorio.claboratorio as Codigo, snombrecorto as Descripcion FROM " & _
  " laboratorio INNER JOIN (distribuidoralaboratorio INNER JOIN distribuidora ON " & _
  " distribuidora.cdistribuidora = distribuidoralaboratorio.cdistribuidora) ON " & _
  " distribuidoralaboratorio.claboratorio = laboratorio.claboratorio WHERE " & _
  " distribuidora.cdistribuidora = '" & Distribuidora & "' ORDER BY snombrecorto"
  Call ConfiguraRst(StrCadena)
  Call LlenaDataCombo(DtcLaboratorio)
  DtcLaboratorio.Enabled = False
End Sub

Private Sub DtcLinea_Click(Area As Integer)
  Linea = Replace(DtcLinea.BoundText, "'", "''")
  StrCadena = "SELECT ctipoproducto as Codigo,sdescripcion as Descripcion FROM " & _
  " tipoproducto WHERE clinea = '" & Linea & "' ORDER BY sdescripcion"
  Call ConfiguraRst(StrCadena)
  Call LlenaDataCombo(DtcTipoProducto)
  DtcTipoProducto.Enabled = False
End Sub

Private Sub Form_Load()
CenterForm Me
  StrCadena = "SELECT clinea as Codigo, sdescripcion as Descripcion FROM " & _
  " linea ORDER BY sdescripcion"
  Call ConfiguraRst(StrCadena)
  Call LlenaDataCombo(DtcLinea)
  StrCadena = "SELECT ctipoproducto as Codigo,sdescripcion as Descripcion FROM " & _
  " tipoproducto ORDER BY sdescripcion"
  Call ConfiguraRst(StrCadena)
  Call LlenaDataCombo(DtcTipoProducto)
  StrCadena = "SELECT claboratorio as Codigo, sRazonSocial as Descripcion FROM " & _
  " laboratorio ORDER BY sRazonSocial"
  Call ConfiguraRst(StrCadena)
  Call LlenaDataCombo(DtcLaboratorio)
  StrCadena = "SELECT cDistribuidora as Codigo, srazonsocial as Descripcion FROM " & _
  " Distribuidora ORDER BY srazonsocial"
  Call ConfiguraRst(StrCadena)
  Call LlenaDataCombo(DtcDistribuidora)
  
  DtcLinea.Enabled = False
  DtcTipoProducto.Enabled = False
  DtcLaboratorio.Enabled = False
  DtcDistribuidora.Enabled = False
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Ans As Boolean
  Select Case Button.Key
    Case KEY_OK
      Linea = ""
      Laboratorio = ""
      TiProdu = ""
      Distribuidora = ""
      If ChkClinea.Value = 1 Then
        Linea = Replace(DtcLinea.BoundText, "'", "''")
      End If
      If ChkCtiprodu.Value = 1 Then
        TiProdu = Replace(DtcTipoProducto.BoundText, "'", "''")
      End If
      If chkLaboratorio.Value = 1 Then
        Laboratorio = Replace(DtcLaboratorio.BoundText, "'", "''")
      End If
      If ChkDistribuidora.Value = 1 Then
        Distribuidora = Replace(DtcDistribuidora.BoundText, "'", "''")
      End If
    Dim DteHoy As Date
      DteHoy = Date
      StrCadena = "SELECT DISTINCT Linea.sDescripcion as Linea, TipoProducto.sDescripcion as " & _
      " TipoProducto, Laboratorio.sRazonSocial as Laboratorio, Distribuidora.srazonsocial as " & _
      " Distribuidora, sDescripcionProducto, nStockActual, Unidad.sDescripcion as Unidad " & _
      " FROM Distribuidora INNER JOIN (Unidad INNER JOIN ((((Laboratorio INNER JOIN " & _
      " DistribuidoraLaboratorio ON Laboratorio.cLaboratorio = " & _
      " DistribuidoraLaboratorio.cLaboratorio) INNER JOIN (Linea INNER JOIN Producto ON " & _
      " Linea.cLinea = Producto.cLinea) ON Laboratorio.cLaboratorio = Producto.cLaboratorio) " & _
      " INNER JOIN TipoProducto ON (TipoProducto.cLinea = Producto.cLinea) AND " & _
      " (TipoProducto.cTipoProducto = Producto.cTipoProducto) AND (Linea.cLinea = " & _
      " TipoProducto.cLinea)) INNER JOIN LoteProducto ON Producto.cProducto = " & _
      " LoteProducto.cProducto) ON Unidad.cUnidad = Producto.cUnidad) ON " & _
      " Distribuidora.cDistribuidora = DistribuidoraLaboratorio.cDistribuidora WHERE " & _
      " (nstockactual<=nstockminimo OR dvencimiento <= cDate('" & DteHoy & "') OR " & _
      " cestado = 'LT' OR cestado = 'LV') AND Producto.cLinea LIKE '" & Linea & "%' AND " & _
      " Producto.cTipoProducto LIKE '" & TiProdu & "%' AND Producto.cLaboratorio " & _
      " LIKE '" & Laboratorio & "%' AND distribuidora.cdistribuidora LIKE '" & Distribuidora & "%' "
      Call ConfiguraRst(StrCadena)
      Ans = ShowMultiReport(rst, "RptProductoCompra", , App.Path + "\Reportes\")
      Set rst = Nothing
    Case KEY_CANCEL
      Unload Me
  End Select
End Sub



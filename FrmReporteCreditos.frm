VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmReporteCreditos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Créditos"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7305
   Begin TabDlg.SSTab SstKardex 
      Height          =   2775
      Left            =   206
      TabIndex        =   0
      Top             =   161
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Clientes y Créditos"
      TabPicture(0)   =   "FrmReporteCreditos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DtcEntidad"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ChkEntidad"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ChkVencido"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrmCredito"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Por Fechas"
      TabPicture(1)   =   "FrmReporteCreditos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblDesde"
      Tab(1).Control(1)=   "LblHasta"
      Tab(1).Control(2)=   "DtpHasta"
      Tab(1).Control(3)=   "DtpDesde"
      Tab(1).Control(4)=   "ChkFecha"
      Tab(1).ControlCount=   5
      Begin VB.Frame FrmCredito 
         Caption         =   "Pagos"
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   180
         TabIndex        =   11
         Top             =   1890
         Width           =   5370
         Begin VB.OptionButton OptPendiente 
            Caption         =   "Créditos Pendientes de Pago"
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   270
            TabIndex        =   13
            Top             =   247
            Width           =   2490
         End
         Begin VB.OptionButton OptCancelados 
            Caption         =   "Créditos Cancelados"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3105
            TabIndex        =   12
            Top             =   270
            Width           =   1950
         End
      End
      Begin VB.CheckBox ChkVencido 
         Caption         =   "Créditos Vencidos"
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   180
         TabIndex        =   10
         Top             =   1395
         Width           =   1770
      End
      Begin VB.CheckBox ChkEntidad 
         BackColor       =   &H80000004&
         Caption         =   "Entidad"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   840
         Width           =   1365
      End
      Begin VB.CheckBox ChkFecha 
         BackColor       =   &H80000004&
         Caption         =   "Vencimiento del Crédito:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74775
         TabIndex        =   1
         Top             =   720
         Width           =   2070
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   -73920
         TabIndex        =   2
         Top             =   1275
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   16711681
         CurrentDate     =   37091
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   315
         Left            =   -71520
         TabIndex        =   3
         Top             =   1275
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   16711681
         CurrentDate     =   37091
      End
      Begin MSDataListLib.DataCombo DtcEntidad 
         Height          =   315
         Left            =   1665
         TabIndex        =   9
         Top             =   810
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
      End
      Begin VB.Label LblHasta 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Hasta : "
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -72120
         TabIndex        =   5
         Top             =   1335
         Width           =   555
      End
      Begin VB.Label LblDesde 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Desde : "
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -74520
         TabIndex        =   4
         Top             =   1335
         Width           =   600
      End
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   1785
      Left            =   6198
      TabIndex        =   6
      Top             =   656
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
         Height          =   780
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
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
      Left            =   6142
      Top             =   2463
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
            Picture         =   "FrmReporteCreditos.frx":0038
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteCreditos.frx":048C
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteCreditos.frx":07AC
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteCreditos.frx":0C00
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteCreditos.frx":1054
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteCreditos.frx":1374
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteCreditos.frx":17C8
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteCreditos.frx":1924
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteCreditos.frx":1D78
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteCreditos.frx":2094
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteCreditos.frx":2970
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteCreditos.frx":2C90
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteCreditos.frx":2FB0
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmReporteCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EnumRptCredito As EnumCredito
Dim Entidad As String

Private Sub ChkEntidad_Click()
  If ChkEntidad.Value = 1 Then
    DtcEntidad.Enabled = True
  Else
    DtcEntidad.Enabled = False
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

Private Sub Form_Load()
Call FormReport(Me)
  If EnumRptCredito = CreditoCliente Then
    ChkEntidad.Caption = "Cliente"
    strCadena = "SELECT ccliente as Codigo, (snombrecliente & chr(32) & sapellidocliente) " & _
    " as Descripcion FROM cliente ORDER BY snombrecliente"
  ElseIf EnumRptCredito = CreditoDistribuidora Then
    ChkEntidad.Caption = "Distribuidora"
    strCadena = "SELECT cdistribuidora as Codigo, srazonsocial as Descripcion " & _
    " FROM distribuidora ORDER BY srazonsocial"
  End If
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(DtcEntidad)
  DtpDesde.Value = Date
  DtpHasta.Value = Date
  
  DtpDesde.Enabled = False
  DtpHasta.Enabled = False
  DtcEntidad.Enabled = False
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Ans As Boolean
Dim DteDesde As Date
Dim DteHasta As Date
Dim estado As String
  Select Case Button.Key
    Case KEY_OK
      Entidad = ""
      estado = ""
      DteDesde = DTEMINIMA
      DteHasta = DTEMAXIMA
      If ChkEntidad.Value = 1 Then
        Entidad = Replace(DtcEntidad.BoundText, "'", "''")
      End If
      If ChkFecha.Value = 1 Then
        DteDesde = DtpDesde.Value
        DteHasta = DtpHasta.Value
      End If
      If ChkVencido.Value = 1 Then
        estado = "PP"
        DteHasta = Date
      Else
        If OptPendiente.Value = True Then
          estado = "PP"
        ElseIf OptCancelados.Value = True Then
          estado = "NN"
        End If
      End If
      If EnumRptCredito = CreditoCliente Then
        strCadena = "SELECT DocumentoVenta.cDocumentoVenta as Documento, " & _
        " (snombrecliente & Chr(32) & sapellidocliente) AS Entidad, dVencimiento, " & _
        " dPago, nTotalVenta AS Total, Estado.sDescripcion AS Estado, " & _
        " CuotaVenta.cCuotaVenta AS Cuota, dPagoCuotaVenta AS dCuota, " & _
        " nTotalPagoCuota FROM Estado INNER JOIN ((Cliente INNER JOIN " & _
        " DocumentoVenta ON Cliente.cCliente = DocumentoVenta.cCliente) INNER JOIN " & _
        " CuotaVenta ON DocumentoVenta.cDocumentoVenta = CuotaVenta.cDocumentoVenta) " & _
        " ON Estado.cEstado = DocumentoVenta.cEstado WHERE DocumentoVenta.cCliente LIKE " & _
        " '" & Entidad & "%' AND DocumentoVenta.cEstado LIKE '" & estado & "%' AND " & _
        " dVencimiento >= cdate('" & DteDesde & "') AND dVencimiento <= " & _
        " cdate('" & DteHasta & "') ORDER BY DocumentoVenta.cCliente, " & _
        " DocumentoVenta.cDocumentoVenta,CuotaVenta.cCuotaVenta"
      ElseIf EnumRptCredito = CreditoDistribuidora Then
        strCadena = "SELECT FacturaCompra.cfactura as Documento, srazonsocial as " & _
        " Entidad, dVencimiento, dPago, nTotalFactura as Total, Estado.sdescripcion " & _
        " as Estado, ccuotacompra as Cuota,  dPagoCuotacompra as dCuota, " & _
        " nTotalPagoCuota FROM CuotaCompra INNER JOIN (Distribuidora INNER JOIN " & _
        " (facturacompra INNER JOIN Estado ON Estado.cestado = " & _
        " facturacompra.cestado) ON facturacompra.cdistribuidora=" & _
        " distribuidora.cdistribuidora) ON Facturacompra.cfactura= " & _
        " Cuotacompra.cfactura WHERE facturacompra.cdistribuidora LIKE " & _
        " '" & Entidad & "%' AND dvencimiento <= cdate('" & DteHasta & "') AND " & _
        " dvencimiento >= cdate('" & DteDesde & "') AND facturacompra.cestado LIKE " & _
        " '" & estado & "%'"
      End If
      Call ConfiguraRst(strCadena)
      Ans = ShowMultiReport(rst, "RptCredito", , App.Path + "\Reportes\")
      Set rst = Nothing
    Case KEY_CANCEL
      Unload Me
  End Select
End Sub



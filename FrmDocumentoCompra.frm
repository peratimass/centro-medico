VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmDocumentoCompra 
   Caption         =   "Documento Compra"
   ClientHeight    =   5190
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   5040
   Icon            =   "FrmDocumentoCompra.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   5040
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3240
      Top             =   4440
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
            Picture         =   "FrmDocumentoCompra.frx":08CA
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoCompra.frx":0D1E
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoCompra.frx":103E
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoCompra.frx":1492
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoCompra.frx":18E6
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoCompra.frx":1C06
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoCompra.frx":205A
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoCompra.frx":21B6
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoCompra.frx":260A
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoCompra.frx":2926
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoCompra.frx":3202
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoCompra.frx":3522
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoCompra.frx":3842
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   4860
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   8573
      BandCount       =   1
      ForeColor       =   8388608
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   4860
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   7020
         Left            =   30
         TabIndex        =   1
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   12383
         ButtonWidth     =   1640
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "    Nuevo    "
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
               Object.ToolTipText     =   "Anular"
               ImageKey        =   "(Anular)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "(Imprimir)"
               Object.ToolTipText     =   "Imprimir"
               ImageKey        =   "(Imprimir)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Pagar"
               Key             =   "(Pagar)"
               Object.ToolTipText     =   "Pagar"
               ImageKey        =   "(Pagar)"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Buscar"
               Key             =   "(Buscar)"
               Object.ToolTipText     =   "Buscar"
               ImageKey        =   "(Buscar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7435
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   -2147483635
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmDocumentoCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Procedencia As EnumProcede
Public StrCodDistribuidora As String, StrCodDocumento As String

Private Sub Form_Activate()
  StrCadena = "SELECT cfactura as Código,facturacompra.cdistribuidora, srazonsocial as Distribuidora," & _
  " ntotalFactura as Total,demisionFactura as FEmisión, dvencimiento as FVencimiento " & _
  " FROM FacturaCompra INNER JOIN distribuidora ON distribuidora.cdistribuidora=" & _
  " FacturaCompra.cdistribuidora WHERE srazonsocial LIKE " & _
  " '%" & FrmBuscarDocumentos.TxtEntidad.Text & "%' AND cfactura" & _
  " LIKE '%" & FrmBuscarDocumentos.TxtNDocumento.Text & "%' AND cestado " & _
  "LIKE '%" & FrmBuscarDocumentos.Estado & "%'  AND demisionFactura >= " & _
  " cdate('" & FrmBuscarDocumentos.DteInicio & "') AND demisionFactura <= " & _
  " cdate('" & FrmBuscarDocumentos.DteFin & "') AND ntotalFactura >= " & _
  " " & FrmBuscarDocumentos.MontoMinimo & " AND ntotalFactura <= " & _
  " " & FrmBuscarDocumentos.MontoMaximo & " ORDER BY cFactura"
  Call ConfiguraRst(StrCadena)
  Set HfdGrilla.Recordset = Rst
  HfdGrilla.ColWidth(0) = 1200
  HfdGrilla.ColWidth(1) = 0
  HfdGrilla.ColWidth(2) = 3500
  HfdGrilla.ColWidth(4) = 1200
  HfdGrilla.ColWidth(5) = 1300
  Call DarFormato(HfdGrilla, 3)
  TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
  TlbAcciones.Buttons(KEY_PRINT).Enabled = False
  TlbAcciones.Buttons(KEY_PAGAR).Enabled = False
  Set Rst = Nothing
End Sub

Private Sub Form_Load()
CenterForm Me
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  Dim intAuxTop As Integer
  intAuxTop = 100
  HfdGrilla.Move 100, intAuxTop, (Me.Width - ClbAcciones.Width - 400), _
  (Me.Height - intAuxTop - 540)
  ClbAcciones.Move HfdGrilla.Width + 200, intAuxTop, Me.Width, _
  (Me.Height - intAuxTop - 540)
End Sub

Private Sub HfdGrilla_Click()
  If HfdGrilla.Row > 0 Then
    TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
    TlbAcciones.Buttons(KEY_PRINT).Enabled = True
    TlbAcciones.Buttons(KEY_PAGAR).Enabled = True
    HfdGrilla.Col = 0
    StrCodDocumento = HfdGrilla.Text
    HfdGrilla.Col = 1
    StrCodDistribuidora = HfdGrilla.Text
  End If
End Sub

Private Sub HfdGrilla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If HfdGrilla.Row > 0 Then
    TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
    TlbAcciones.Buttons(KEY_PRINT).Enabled = True
    TlbAcciones.Buttons(KEY_PAGAR).Enabled = True
    HfdGrilla.Col = 0
    StrCodDocumento = HfdGrilla.Text
    HfdGrilla.Col = 1
    StrCodDistribuidora = HfdGrilla.Text
  End If
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case KEY_NEW
      Procedencia = Nuevo
      FrmDetalleDocumentoCompra.Show
    Case KEY_ANULAR
      StrCadena = "SELECT cfactura FROM facturacompra WHERE cfactura = " & _
      " '" & StrCodDocumento & "' AND cdistribuidora = '" & StrCodDistribuidora & "' AND " & _
      " cestado = 'NN' "
      Call EjecutaRST(StrCadena)
      If Not RstEjecuta.EOF Then
        If MsgBox(MSGANULAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
          CnBd.BeginTrans
            StrCadena = "UPDATE facturacompra SET cestado='AA' WHERE " & _
            " cfactura ='" & StrCodDocumento & "' AND cdistribuidora = '" & StrCodDistribuidora & "'"
            Call EjecutaRST(StrCadena)
            StrCadena = "SELECT cproducto, ncantidadfactura, npreciocompra FROM detallefacturacompra WHERE " & _
            " cfactura ='" & StrCodDocumento & "' AND cdistribuidora = '" & StrCodDistribuidora & "'"
            Call ConfiguraTemporal(StrCadena)
            RstTemporal.MoveFirst
            Do While Not RstTemporal.EOF
              Call Kardex(RstTemporal(0), "S02", RstTemporal(1), Date, StrCodDocumento, RstTemporal(2))
              RstTemporal.MoveNext
            Loop
          CnBd.CommitTrans
          Form_Activate
        End If
      Else
        MsgBox MSGNOTANULAR, vbInformation, MSGVALIDACION
      End If
    Case KEY_PRINT
      Dim Ans As Boolean
        StrCadena = "SELECT facturacompra.cfactura, sRazonSocial as Distribuidora, " & _
        " dEmisionFactura, dVencimiento, dPago, " & _
        " Estado.sDescripcion as Estado, nTotalFactura, " & _
        " nCantidadFactura, Unidad.sDescripcion as Unidad, " & _
        " sDescripcionProducto, nPrecioCompra, " & _
        " ([npreciocompra]*[ncantidadfactura]) AS PTotal FROM Estado INNER JOIN " & _
        " (Unidad INNER JOIN (Producto INNER JOIN (Distribuidora INNER JOIN (FacturaCompra  " & _
        " INNER JOIN DetalleFacturaCompra ON (FacturaCompra.cFactura =  " & _
        " DetalleFacturaCompra.cFactura) AND (FacturaCompra.cDistribuidora =  " & _
        " DetalleFacturaCompra.cDistribuidora)) ON (Distribuidora.cDistribuidora =  " & _
        " FacturaCompra.cDistribuidora) AND (DetalleFacturaCompra.cDistribuidora =  " & _
        " Distribuidora.cDistribuidora)) ON Producto.cProducto =  " & _
        " DetalleFacturaCompra.cProducto) ON Unidad.cUnidad = Producto.cUnidad) ON  " & _
        " Estado.cEstado = FacturaCompra.cEstado WHERE facturacompra.cfactura = " & _
        " '" & StrCodDocumento & "' AND facturacompra.cdistribuidora = '" & StrCodDistribuidora & "' "
        Call ConfiguraRst(StrCadena)
        Ans = ShowMultiReport(Rst, "RptFactura", , App.Path + "\Reportes\")
        Set Rst = Nothing
    Case KEY_PAGAR
      StrCadena = "SELECT cfactura FROM facturacompra WHERE cfactura = " & _
      " '" & StrCodDocumento & "' AND cdistribuidora = '" & StrCodDistribuidora & "' AND " & _
      " cestado = 'PP' "
      Call EjecutaRST(StrCadena)
      If Not RstEjecuta.EOF Then
        Set RstEjecuta = Nothing
        FrmPagoCuotaCredito.EnumFrmPago = BDocumentoCompra
        FrmPagoCuotaCredito.Show
      Else
        MsgBox MSGNOTPAGAR, vbInformation, MSGVALIDACION
      End If
    Case KEY_BROWSER
      Procedencia = Buscar
      Me.Hide
      FrmBuscarDocumentos.Show
    Case KEY_REPORTE
  End Select
End Sub

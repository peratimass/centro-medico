VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmDocumentoVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas"
   ClientHeight    =   5430
   ClientLeft      =   675
   ClientTop       =   810
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   9375
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3360
      Top             =   4680
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
            Picture         =   "FrmDocumentoVenta.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoVenta.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoVenta.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoVenta.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoVenta.frx":101C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoVenta.frx":133C
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoVenta.frx":1790
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoVenta.frx":18EC
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoVenta.frx":1D40
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoVenta.frx":205C
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoVenta.frx":2938
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoVenta.frx":2C58
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDocumentoVenta.frx":2F78
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   4860
      Left            =   8280
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
         ButtonWidth     =   1482
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
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
      Height          =   2175
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   2175
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3836
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmDocumentoVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Procedencia As EnumProcede
Public StrCodDocumento As String
Dim StrEntidad As String * 1

Private Sub Form_Activate()
  StrCadena = "SELECT cdocumentoventa as Código,(snombrecliente & chr(32) & sapellidocliente) as Cliente," & _
  " ntotalventa as Total,demisionventa as FEmisión " & _
  " FROM documentoventa INNER JOIN cliente ON cliente.ccliente=" & _
  " documentoventa.ccliente WHERE (snombrecliente LIKE " & _
  " '%" & FrmBuscarDocumentos.TxtEntidad.Text & "%' OR sapellidocliente LIKE " & _
  " '%" & FrmBuscarDocumentos.TxtEntidad.Text & "%') AND cdocumentoventa" & _
  " LIKE '%" & FrmBuscarDocumentos.TxtNDocumento.Text & "%' AND cestado " & _
  "LIKE '%" & FrmBuscarDocumentos.Estado & "%'  AND demisionventa >= " & _
  " cdate('" & FrmBuscarDocumentos.DteInicio & "') AND demisionventa <= " & _
  " cdate('" & FrmBuscarDocumentos.DteFin & "') AND ntotalventa >= " & _
  " " & FrmBuscarDocumentos.MontoMinimo & " AND ntotalventa <= " & _
  " " & FrmBuscarDocumentos.MontoMaximo & " ORDER BY cdocumentoventa"
  Call ConfiguraRst(StrCadena)
  Set HfdGrilla.Recordset = Rst
  HfdGrilla.ColWidth(0) = 1200
  HfdGrilla.ColWidth(1) = 3500
  HfdGrilla.ColWidth(3) = 1200
  HfdGrilla.ColWidth(4) = 1300
  Call DarFormato(HfdGrilla, 2)
  Set Rst = Nothing
  
   StrCadena = "SELECT cdocumentoventa as Código,spersona as Persona,ntotalventa as Total," & _
  " demisionventa as FEmisión " & _
  " FROM documentoventa WHERE spersona LIKE '%" & FrmBuscarDocumentos.TxtEntidad.Text & "%'  " & _
  " AND cdocumentoventa LIKE '%" & FrmBuscarDocumentos.TxtNDocumento.Text & "%'  " & _
  " AND cestado LIKE '%" & FrmBuscarDocumentos.Estado & "%'  AND demisionventa >= " & _
  " cdate('" & FrmBuscarDocumentos.DteInicio & "') AND demisionventa <= " & _
  " cdate('" & FrmBuscarDocumentos.DteFin & "') AND ntotalventa >= " & _
  " " & FrmBuscarDocumentos.MontoMinimo & " AND ntotalventa <= " & _
  " " & FrmBuscarDocumentos.MontoMaximo & "  AND ccliente = '' ORDER BY cdocumentoventa"
  Call ConfiguraRst(StrCadena)
  Set HfdPersona.Recordset = Rst
  HfdPersona.ColWidth(0) = 1200
  HfdPersona.ColWidth(1) = 3500
  HfdPersona.ColWidth(3) = 1200
  HfdPersona.ColWidth(4) = 1300
  Call DarFormato(HfdPersona, 2)
  Set Rst = Nothing
  
  TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
  TlbAcciones.Buttons(KEY_PRINT).Enabled = False
  TlbAcciones.Buttons(KEY_PAGAR).Enabled = False
End Sub

Private Sub Form_Load()
CenterForm Me
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  Dim intAuxTop As Integer
  intAuxTop = (Me.Height - 700) / 2
  HfdGrilla.Move 100, 100, (Me.Width - ClbAcciones.Width - 400), intAuxTop
  HfdPersona.Move 100, intAuxTop + 200, (Me.Width - ClbAcciones.Width - 400), intAuxTop
  ClbAcciones.Move HfdGrilla.Width + 200, 100, Me.Width, (Me.Height - 640)
End Sub

Private Sub HfdGrilla_Click()
  If HfdGrilla.Row > 0 Then
    TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
    TlbAcciones.Buttons(KEY_PRINT).Enabled = True
    TlbAcciones.Buttons(KEY_PAGAR).Enabled = True
    HfdGrilla.Col = 0
    StrCodDocumento = HfdGrilla.Text
    StrEntidad = "C"
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
      StrEntidad = "C"
    End If
  End If
End Sub

Private Sub HfdPersona_Click()
  If HfdPersona.Row > 0 Then
    TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
    TlbAcciones.Buttons(KEY_PRINT).Enabled = True
    HfdPersona.Col = 0
    StrCodDocumento = HfdPersona.Text
    StrEntidad = "P"
  End If
End Sub

Private Sub HfdPersona_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If HfdPersona.Row > 0 Then
      TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
      TlbAcciones.Buttons(KEY_PRINT).Enabled = True
      HfdPersona.Col = 0
      StrCodDocumento = HfdPersona.Text
      StrEntidad = "P"
    End If
  End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case KEY_NEW
      Procedencia = Nuevo
      FrmDetalleDocumentoVenta.Show
    Case KEY_ANULAR
      StrCadena = "SELECT cdocumentoventa FROM documentoventa WHERE " & _
      " cdocumentoventa = '" & StrCodDocumento & "' AND cestado = 'NN' "
      Call EjecutaRST(StrCadena)
      If Not RstEjecuta.EOF Then
        If MsgBox(MSGANULAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
          CnBd.BeginTrans
            StrCadena = "UPDATE documentoventa SET cestado='AA' WHERE " & _
            " cdocumentoventa='" & StrCodDocumento & "' "
            Call EjecutaRST(StrCadena)
            StrCadena = "SELECT cproducto, ncantidad, nprecioventa FROM detalledocumentoventa WHERE " & _
            " cdocumentoventa='" & StrCodDocumento & "' "
            Call ConfiguraTemporal(StrCadena)
            RstTemporal.MoveFirst
            Do While Not RstTemporal.EOF
              Call Kardex(RstTemporal(0), "E02", RstTemporal(1), Date, StrCodDocumento, RstTemporal(2))
              RstTemporal.MoveNext
            Loop
          CnBd.CommitTrans
          Set RstTemporal = Nothing
          Form_Activate
        End If
      Else
        MsgBox MSGNOTANULAR, vbInformation, MSGVALIDACION
      End If
    Case KEY_PRINT
      Dim Ans As Boolean
        If StrEntidad = "C" Then
          StrCadena = "SELECT documentoventa.cdocumentoventa, " & _
          " (sNombreCliente & chr(32) & sapellidocliente) as Cliente, demisionventa, " & _
          " dvencimiento, dpago, nsubtotal, nigv, ntotalventa, Estado.sDescripcion as Estado, " & _
          " nCantidad, Unidad.sDescripcion as Unidad, sDescripcionProducto, " & _
          " detalledocumentoventa.nPrecioventa, ([detalledocumentoventa.nprecioventa]*[ncantidad]) AS " & _
          " PTotal FROM Estado INNER JOIN (Cliente INNER JOIN (DocumentoVenta INNER JOIN " & _
          " ((Unidad INNER JOIN Producto ON Unidad.cUnidad = Producto.cUnidad) INNER JOIN " & _
          " DetalleDocumentoVenta ON Producto.cProducto = DetalleDocumentoVenta.cProducto) ON " & _
          " DocumentoVenta.cDocumentoVenta= DetalleDocumentoVenta.cDocumentoVenta) ON " & _
          " Cliente.cCliente = DocumentoVenta.cCliente) ON Estado.cEstado = " & _
          " DocumentoVenta.cEstado WHERE DocumentoVenta.cDocumentoVenta='" & StrCodDocumento & "'"
        End If
        If StrEntidad = "P" Then
          StrCadena = "SELECT documentoventa.cdocumentoventa, sPersona as Cliente, " & _
          " demisionventa, dvencimiento, dpago, nsubtotal, nigv, ntotalventa, " & _
          " Estado.sDescripcion as Estado, nCantidad, Unidad.sDescripcion as Unidad, " & _
          " sDescripcionProducto, detalledocumentoventa.nPrecioVenta, " & _
          " ([detalledocumentoventa.nPrecioVenta]*[ncantidad]) AS PTotal FROM " & _
          " Estado INNER JOIN (DocumentoVenta INNER JOIN ((Unidad " & _
          " INNER JOIN Producto ON Unidad.cUnidad = Producto.cUnidad) INNER JOIN " & _
          " DetalleDocumentoVenta ON Producto.cProducto = DetalleDocumentoVenta.cProducto) ON " & _
          " DocumentoVenta.cDocumentoVenta= DetalleDocumentoVenta.cDocumentoVenta) ON " & _
          " Estado.cEstado = DocumentoVenta.cEstado WHERE " & _
          " DocumentoVenta.cDocumentoVenta='" & StrCodDocumento & "'"
        End If
        Call ConfiguraRst(StrCadena)
        Ans = ShowMultiReport(Rst, "RptDocumentoVenta", , App.Path + "\Reportes\")
        Set Rst = Nothing
    Case KEY_PAGAR
      StrCadena = "SELECT cdocumentoventa FROM documentoventa WHERE " & _
      " cdocumentoventa= '" & StrCodDocumento & "' AND cestado = 'PP' "
      Call EjecutaRST(StrCadena)
      If Not RstEjecuta.EOF Then
        Set RstEjecuta = Nothing
        FrmPagoCuotaCredito.EnumFrmPago = BDocumentoVenta
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

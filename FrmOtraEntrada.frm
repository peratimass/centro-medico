VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOtraEntrada 
   Caption         =   "Otra Entrada"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   5010
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   4320
      Top             =   4080
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
            Picture         =   "FrmOtraEntrada.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraEntrada.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraEntrada.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraEntrada.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraEntrada.frx":101C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraEntrada.frx":133C
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraEntrada.frx":1790
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraEntrada.frx":18EC
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraEntrada.frx":1D40
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraEntrada.frx":205C
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraEntrada.frx":2938
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraEntrada.frx":2C58
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraEntrada.frx":2F78
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   3900
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   6879
      BandCount       =   1
      ForeColor       =   -2147483635
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   3900
      _Version        =   "6.7.8862"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   1
         Top             =   345
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1508
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
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
               Caption         =   "Buscar"
               Key             =   "(Buscar)"
               Object.ToolTipText     =   "Buscar"
               ImageKey        =   "(Buscar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2355
      _Version        =   393216
      ForeColor       =   -2147483635
      FixedCols       =   0
      ForeColorFixed  =   -2147483635
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdCliente 
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2355
      _Version        =   393216
      ForeColor       =   -2147483635
      FixedCols       =   0
      ForeColorFixed  =   -2147483635
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDistribuidora 
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2355
      _Version        =   393216
      ForeColor       =   -2147483635
      FixedCols       =   0
      ForeColorFixed  =   -2147483635
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmOtraEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Procedencia As EnumProcede
Public StrCodDocumento As String
Dim StrEntidad As String * 1

Private Sub Form_Activate()
  StrCadena = "SELECT cotraentrada as Código,sPersona as Persona,sdescripcionmovimiento " & _
  " as Tipo_Mov, dotraentrada as Fecha FROM otraentrada INNER JOIN tipomovimiento ON " & _
  " tipomovimiento.ctipomovimiento=otraentrada.ctipomovimiento WHERE spersona LIKE " & _
  " '%" & FrmBuscarDocumentos.TxtEntidad.Text & "%'  AND cotraentrada LIKE " & _
  " '%" & FrmBuscarDocumentos.TxtNDocumento.Text & "%' AND cestado LIKE " & _
  " '%" & FrmBuscarDocumentos.Estado & "%'  AND dotraentrada >= " & _
  " cdate('" & FrmBuscarDocumentos.DteInicio & "') AND dotraentrada <= " & _
  " cdate('" & FrmBuscarDocumentos.DteFin & "') AND stipoentidad = 'P' ORDER BY cotraentrada"
  Call ConfiguraRst(StrCadena)
  Set HfdPersona.Recordset = Rst
  Set Rst = Nothing
  HfdPersona.ColWidth(0) = 1200
  HfdPersona.ColWidth(1) = 3500
  HfdPersona.ColWidth(2) = 1200
  HfdPersona.ColWidth(3) = 1200
  
  StrCadena = "SELECT cotraentrada as Código,(snombrecliente & chr(32) & sapellidocliente) as Cliente," & _
  " sdescripcionmovimiento as Tipo_Mov, dotraentrada as Fecha FROM tipomovimiento INNER JOIN " & _
  " (otraentrada INNER JOIN cliente ON cliente.ccliente=otraentrada.centidadorigen) ON " & _
  " otraentrada.ctipomovimiento=tipomovimiento.ctipomovimiento WHERE (snombrecliente LIKE " & _
  " '%" & FrmBuscarDocumentos.TxtEntidad.Text & "%' OR sapellidocliente " & _
  " LIKE '%" & FrmBuscarDocumentos.TxtEntidad.Text & "%') AND cotraentrada LIKE " & _
  " '%" & FrmBuscarDocumentos.TxtNDocumento.Text & "%' AND cestado LIKE " & _
  " '%" & FrmBuscarDocumentos.Estado & "%'  AND dotraentrada >= " & _
  " cdate('" & FrmBuscarDocumentos.DteInicio & "') AND dotraentrada <= " & _
  " cdate('" & FrmBuscarDocumentos.DteFin & "') AND stipoentidad = 'C' ORDER BY cotraentrada"
  Call ConfiguraRst(StrCadena)
  Set HfdCliente.Recordset = Rst
  Set Rst = Nothing
  HfdCliente.ColWidth(0) = 1200
  HfdCliente.ColWidth(1) = 3500
  HfdCliente.ColWidth(2) = 1200
  HfdCliente.ColWidth(3) = 1200
  
  StrCadena = "SELECT cotraentrada as Código,srazonsocial as Distribuidora,sdescripcionmovimiento " & _
  " AS Tipo_Mov, dotraentrada as Fecha FROM tipomovimiento INNER JOIN " & _
  " (otraentrada INNER JOIN distribuidora ON distribuidora.cdistribuidora=otraentrada.centidadorigen) " & _
  " ON otraentrada.ctipomovimiento=tipomovimiento.ctipomovimiento WHERE srazonsocial " & _
  " LIKE '%" & FrmBuscarDocumentos.TxtEntidad.Text & "%' AND cotraentrada LIKE " & _
  " '%" & FrmBuscarDocumentos.TxtNDocumento.Text & "%' AND cestado LIKE " & _
  " '%" & FrmBuscarDocumentos.Estado & "%'  AND dotraentrada >= " & _
  " cdate('" & FrmBuscarDocumentos.DteInicio & "') AND dotraentrada <= " & _
  " cdate('" & FrmBuscarDocumentos.DteFin & "') AND stipoentidad = 'D' ORDER BY cotraentrada"
  Call ConfiguraRst(StrCadena)
  Set HfdDistribuidora.Recordset = Rst
  Set Rst = Nothing
  HfdDistribuidora.ColWidth(0) = 1200
  HfdDistribuidora.ColWidth(1) = 3500
  HfdDistribuidora.ColWidth(2) = 1200
  HfdDistribuidora.ColWidth(3) = 1200
  
  TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
  TlbAcciones.Buttons(KEY_PRINT).Enabled = False
  Call Form_Resize
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  Dim intAuxTop As Integer
  intAuxTop = (Me.Height - 810) / 3
  HfdPersona.Move 100, 100, (Me.Width - ClbAcciones.Width - 400), intAuxTop
  ClbAcciones.Move HfdPersona.Width + 200, 100, Me.ClbAcciones.Width, (Me.Height - 640)
  
  HfdCliente.Move 100, intAuxTop + 200, (Me.Width - ClbAcciones.Width - 400), intAuxTop
  
  HfdDistribuidora.Move 100, 2 * intAuxTop + 300, (Me.Width - ClbAcciones.Width - 400), intAuxTop
End Sub

Private Sub HfdCliente_Click()
  If HfdCliente.Row > 0 Then
    TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
    TlbAcciones.Buttons(KEY_PRINT).Enabled = True
    HfdCliente.Col = 0
    StrCodDocumento = HfdCliente.Text
    StrEntidad = "C"
  End If
End Sub

Private Sub HfdCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If HfdCliente.Row > 0 Then
    TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
    TlbAcciones.Buttons(KEY_PRINT).Enabled = True
    HfdCliente.Col = 0
    StrCodDocumento = HfdCliente.Text
    StrEntidad = "C"
  End If
End If
End Sub

Private Sub HfdDistribuidora_Click()
  If HfdDistribuidora.Row > 0 Then
    TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
    TlbAcciones.Buttons(KEY_PRINT).Enabled = True
    HfdDistribuidora.Col = 0
    StrCodDocumento = HfdDistribuidora.Text
    StrEntidad = "D"
  End If
End Sub

Private Sub HfdDistribuidora_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If HfdDistribuidora.Row > 0 Then
    TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
    TlbAcciones.Buttons(KEY_PRINT).Enabled = True
    HfdDistribuidora.Col = 0
    StrCodDocumento = HfdDistribuidora.Text
    StrEntidad = "D"
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
      FrmDetalleOtraEntrada.Show
    Case KEY_ANULAR
      StrCadena = "SELECT cotraentrada FROM otraentrada WHERE cotraentrada= '" & StrCodDocumento & "' " & _
      " AND (cestado = 'NN')"
      Call EjecutaRST(StrCadena)
      If Not RstEjecuta.EOF Then
        If MsgBox(MSGANULAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
          CnBd.BeginTrans
            StrCadena = "UPDATE otraentrada SET cestado='AA' WHERE cotraentrada='" & StrCodDocumento & "' "
            Call EjecutaRST(StrCadena)
            StrCadena = "SELECT cproducto, ncantidadotraentrada FROM detalleotraentrada WHERE " & _
            " cotraentrada='" & StrCodDocumento & "'"
            Call ConfiguraTemporal(StrCadena)
            RstTemporal.MoveFirst
            Do While Not RstTemporal.EOF
              Call Kardex(RstTemporal(0), "S02", RstTemporal(1), Date, StrCodDocumento, 0)
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
        If StrEntidad = "D" Then
          StrCadena = "SELECT OtraEntrada.cOtraEntrada, srazonsocial as Entidad, " & _
          " TipoMovimiento.sDescripcionMovimiento as TipoMovimiento, Estado.sDescripcion as " & _
          " Estado, sTipoEntidad, dOtraEntrada, nCantidadOtraEntrada, Unidad.sDescripcion " & _
          " as Unidad, sDescripcionProducto, nPrecioVenta, ([nPrecioVenta]*[ncantidadotraentrada]) " & _
          " AS PTotal FROM Estado INNER JOIN (Unidad INNER JOIN (TipoMovimiento INNER JOIN " & _
          " (Producto INNER JOIN ((Distribuidora INNER JOIN OtraEntrada ON Distribuidora.cDistribuidora= " & _
          " OtraEntrada.cEntidadOrigen) INNER JOIN DetalleOtraEntrada ON " & _
          " OtraEntrada.cOtraEntrada = DetalleOtraEntrada.cOtraEntrada) ON Producto.cProducto " & _
          " = DetalleOtraEntrada.cProducto) ON TipoMovimiento.cTipoMovimiento = " & _
          " OtraEntrada.cTipoMovimiento) ON Unidad.cUnidad = Producto.cUnidad) ON " & _
          " Estado.cEstado = OtraEntrada.cEstado WHERE OtraEntrada.cOtraEntrada= '" & StrCodDocumento & "'"
        End If
        If StrEntidad = "C" Then
          StrCadena = "SELECT OtraEntrada.cOtraEntrada, " & _
          " (sNombreCliente & chr(32) & sapellidocliente) as Entidad, " & _
          " TipoMovimiento.sDescripcionMovimiento as TipoMovimiento, Estado.sDescripcion as " & _
          " Estado, sTipoEntidad, dOtraEntrada, nCantidadOtraEntrada, Unidad.sDescripcion " & _
          " as Unidad, sDescripcionProducto, nPrecioVenta, ([nPrecioVenta]*[ncantidadotraentrada]) " & _
          " AS PTotal FROM Estado INNER JOIN (Unidad INNER JOIN (TipoMovimiento INNER JOIN " & _
          " (Producto INNER JOIN ((Cliente INNER JOIN OtraEntrada ON Cliente.cCliente = " & _
          " OtraEntrada.cEntidadOrigen) INNER JOIN DetalleOtraEntrada ON " & _
          " OtraEntrada.cOtraEntrada = DetalleOtraEntrada.cOtraEntrada) ON Producto.cProducto " & _
          " = DetalleOtraEntrada.cProducto) ON TipoMovimiento.cTipoMovimiento = " & _
          " OtraEntrada.cTipoMovimiento) ON Unidad.cUnidad = Producto.cUnidad) ON " & _
          " Estado.cEstado = OtraEntrada.cEstado WHERE OtraEntrada.cOtraEntrada= '" & StrCodDocumento & "'"
        End If
        If StrEntidad = "P" Then
          StrCadena = "SELECT OtraEntrada.cOtraEntrada, sPersona as Entidad, " & _
          " TipoMovimiento.sDescripcionMovimiento as TipoMovimiento, Estado.sDescripcion as " & _
          " Estado, sTipoEntidad, dOtraEntrada, nCantidadOtraEntrada, Unidad.sDescripcion " & _
          " as Unidad, sDescripcionProducto, nPrecioVenta, ([nPrecioVenta]*[ncantidadotraentrada]) " & _
          " AS PTotal FROM Estado INNER JOIN (Unidad INNER JOIN (TipoMovimiento INNER JOIN " & _
          " (Producto INNER JOIN (OtraEntrada INNER JOIN DetalleOtraEntrada ON " & _
          " OtraEntrada.cOtraEntrada = DetalleOtraEntrada.cOtraEntrada) ON Producto.cProducto " & _
          " = DetalleOtraEntrada.cProducto) ON TipoMovimiento.cTipoMovimiento = " & _
          " OtraEntrada.cTipoMovimiento) ON Unidad.cUnidad = Producto.cUnidad) ON " & _
          " Estado.cEstado = OtraEntrada.cEstado WHERE OtraEntrada.cOtraEntrada= '" & StrCodDocumento & "'"
        End If
        Call ConfiguraRst(StrCadena)
        Ans = ShowMultiReport(Rst, "RptOtraEntrada", , App.Path + "\Reportes\")
        Set Rst = Nothing
    Case KEY_BROWSER
      Procedencia = Buscar
      Me.Hide
      FrmBuscarDocumentos.Show
  End Select
End Sub

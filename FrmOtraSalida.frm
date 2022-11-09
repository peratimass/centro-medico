VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmOtraSalida 
   Caption         =   "Otra Salida"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5025
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
            Picture         =   "FrmOtraSalida.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraSalida.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraSalida.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraSalida.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraSalida.frx":101C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraSalida.frx":133C
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraSalida.frx":1790
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraSalida.frx":18EC
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraSalida.frx":1D40
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraSalida.frx":205C
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraSalida.frx":2938
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraSalida.frx":2C58
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOtraSalida.frx":2F78
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
         ButtonWidth     =   1429
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Nuevo  "
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
Attribute VB_Name = "FrmOtraSalida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Procedencia As EnumProcede
Public StrCodDocumento As String
Dim StrEntidad As String * 1

Private Sub Form_Activate()
  StrCadena = "SELECT cotrasalida as Código,sPersona as Persona,sdescripcionmovimiento " & _
  " as Tipo_Mov, dotrasalida as Fecha FROM otrasalida INNER JOIN tipomovimiento ON " & _
  " tipomovimiento.ctipomovimiento=otrasalida.ctipomovimiento WHERE spersona LIKE " & _
  " '%" & FrmBuscarDocumentos.TxtEntidad.Text & "%'  AND cotrasalida LIKE " & _
  " '%" & FrmBuscarDocumentos.TxtNDocumento.Text & "%' AND cestado LIKE " & _
  " '%" & FrmBuscarDocumentos.Estado & "%'  AND dotrasalida >= " & _
  " cdate('" & FrmBuscarDocumentos.DteInicio & "') AND dotrasalida <= " & _
  " cdate('" & FrmBuscarDocumentos.DteFin & "') AND stipoentidad = 'P' ORDER BY cotrasalida"
  Call ConfiguraRst(StrCadena)
  Set HfdPersona.Recordset = Rst
  Set Rst = Nothing
  HfdPersona.ColWidth(0) = 1200
  HfdPersona.ColWidth(1) = 3500
  HfdPersona.ColWidth(2) = 1200
  HfdPersona.ColWidth(3) = 1200
  
  StrCadena = "SELECT cotrasalida as Código,(snombrecliente & chr(32) & sapellidocliente) as Cliente," & _
  " sdescripcionmovimiento as Tipo_Mov, dotrasalida as Fecha FROM tipomovimiento INNER JOIN " & _
  " (otrasalida INNER JOIN cliente ON cliente.ccliente=otrasalida.centidadorigen) ON " & _
  " otrasalida.ctipomovimiento=tipomovimiento.ctipomovimiento WHERE (snombrecliente LIKE " & _
  " '%" & FrmBuscarDocumentos.TxtEntidad.Text & "%' OR sapellidocliente " & _
  " LIKE '%" & FrmBuscarDocumentos.TxtEntidad.Text & "%') AND cotrasalida LIKE " & _
  " '%" & FrmBuscarDocumentos.TxtNDocumento.Text & "%' AND cestado LIKE " & _
  " '%" & FrmBuscarDocumentos.Estado & "%'  AND dotrasalida >= " & _
  " cdate('" & FrmBuscarDocumentos.DteInicio & "') AND dotrasalida <= " & _
  " cdate('" & FrmBuscarDocumentos.DteFin & "') AND stipoentidad = 'C' ORDER BY cotrasalida"
  Call ConfiguraRst(StrCadena)
  Set HfdCliente.Recordset = Rst
  Set Rst = Nothing
  HfdCliente.ColWidth(0) = 1200
  HfdCliente.ColWidth(1) = 3500
  HfdCliente.ColWidth(2) = 1200
  HfdCliente.ColWidth(3) = 1200
  
  StrCadena = "SELECT cotrasalida as Código,srazonsocial as Distribuidora,sdescripcionmovimiento " & _
  " AS Tipo_Mov, dotrasalida as Fecha FROM tipomovimiento INNER JOIN " & _
  " (otrasalida INNER JOIN distribuidora ON distribuidora.cdistribuidora=otrasalida.centidadorigen) " & _
  " ON otrasalida.ctipomovimiento=tipomovimiento.ctipomovimiento WHERE srazonsocial " & _
  " LIKE '%" & FrmBuscarDocumentos.TxtEntidad.Text & "%' AND cotrasalida LIKE " & _
  " '%" & FrmBuscarDocumentos.TxtNDocumento.Text & "%' AND cestado LIKE " & _
  " '%" & FrmBuscarDocumentos.Estado & "%'  AND dotrasalida >= " & _
  " cdate('" & FrmBuscarDocumentos.DteInicio & "') AND dotrasalida <= " & _
  " cdate('" & FrmBuscarDocumentos.DteFin & "') AND stipoentidad = 'D' ORDER BY cotrasalida"
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
      FrmDetalleOtraSalida.Show
    Case KEY_ANULAR
      StrCadena = "SELECT cotrasalida FROM otrasalida WHERE cotrasalida= '" & StrCodDocumento & "' " & _
      " AND (cestado = 'NN')"
      Call EjecutaRST(StrCadena)
      If Not RstEjecuta.EOF Then
        If MsgBox(MSGANULAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
          CnBd.BeginTrans
            StrCadena = "UPDATE otrasalida SET cestado='AA' WHERE cotrasalida='" & StrCodDocumento & "' "
            Call EjecutaRST(StrCadena)
            StrCadena = "SELECT cproducto, ncantidadotrasalida FROM detalleotrasalida WHERE " & _
            " cotrasalida='" & StrCodDocumento & "'"
            Call ConfiguraTemporal(StrCadena)
            RstTemporal.MoveFirst
            Do While Not RstTemporal.EOF
              Call Kardex(RstTemporal(0), "E02", RstTemporal(1), Date, StrCodDocumento, 0)
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
          StrCadena = "SELECT otrasalida.cotrasalida, srazonsocial as Entidad, " & _
          " TipoMovimiento.sDescripcionMovimiento as TipoMovimiento, Estado.sDescripcion as " & _
          " Estado, sTipoEntidad, dotrasalida, nCantidadotrasalida, Unidad.sDescripcion " & _
          " as Unidad, sDescripcionProducto, nPrecioVenta, ([nPrecioVenta]*[ncantidadotrasalida]) " & _
          " AS PTotal FROM Estado INNER JOIN (Unidad INNER JOIN (TipoMovimiento INNER JOIN " & _
          " (Producto INNER JOIN ((Distribuidora INNER JOIN otrasalida ON Distribuidora.cDistribuidora= " & _
          " otrasalida.cEntidadOrigen) INNER JOIN Detalleotrasalida ON " & _
          " otrasalida.cotrasalida = Detalleotrasalida.cotrasalida) ON Producto.cProducto " & _
          " = Detalleotrasalida.cProducto) ON TipoMovimiento.cTipoMovimiento = " & _
          " otrasalida.cTipoMovimiento) ON Unidad.cUnidad = Producto.cUnidad) ON " & _
          " Estado.cEstado = otrasalida.cEstado WHERE otrasalida.cotrasalida= '" & StrCodDocumento & "'"
        End If
        If StrEntidad = "C" Then
          StrCadena = "SELECT otrasalida.cotrasalida, " & _
          " (sNombreCliente & chr(32) & sapellidocliente) as Entidad, " & _
          " TipoMovimiento.sDescripcionMovimiento as TipoMovimiento, Estado.sDescripcion as " & _
          " Estado, sTipoEntidad, dotrasalida, nCantidadotrasalida, Unidad.sDescripcion " & _
          " as Unidad, sDescripcionProducto, nPrecioVenta, ([nPrecioVenta]*[ncantidadotrasalida]) " & _
          " AS PTotal FROM Estado INNER JOIN (Unidad INNER JOIN (TipoMovimiento INNER JOIN " & _
          " (Producto INNER JOIN ((Cliente INNER JOIN otrasalida ON Cliente.cCliente = " & _
          " otrasalida.cEntidadOrigen) INNER JOIN Detalleotrasalida ON " & _
          " otrasalida.cotrasalida = Detalleotrasalida.cotrasalida) ON Producto.cProducto " & _
          " = Detalleotrasalida.cProducto) ON TipoMovimiento.cTipoMovimiento = " & _
          " otrasalida.cTipoMovimiento) ON Unidad.cUnidad = Producto.cUnidad) ON " & _
          " Estado.cEstado = otrasalida.cEstado WHERE otrasalida.cotrasalida= '" & StrCodDocumento & "'"
        End If
        If StrEntidad = "P" Then
          StrCadena = "SELECT otrasalida.cotrasalida, sPersona as Entidad, " & _
          " TipoMovimiento.sDescripcionMovimiento as TipoMovimiento, Estado.sDescripcion as " & _
          " Estado, sTipoEntidad, dotrasalida, nCantidadotrasalida, Unidad.sDescripcion " & _
          " as Unidad, sDescripcionProducto, nPrecioVenta, ([nPrecioVenta]*[ncantidadotrasalida]) " & _
          " AS PTotal FROM Estado INNER JOIN (Unidad INNER JOIN (TipoMovimiento INNER JOIN " & _
          " (Producto INNER JOIN (otrasalida INNER JOIN Detalleotrasalida ON " & _
          " otrasalida.cotrasalida = Detalleotrasalida.cotrasalida) ON Producto.cProducto " & _
          " = Detalleotrasalida.cProducto) ON TipoMovimiento.cTipoMovimiento = " & _
          " otrasalida.cTipoMovimiento) ON Unidad.cUnidad = Producto.cUnidad) ON " & _
          " Estado.cEstado = otrasalida.cEstado WHERE otrasalida.cotrasalida= '" & StrCodDocumento & "'"
        End If
        Call ConfiguraRst(StrCadena)
        Ans = ShowMultiReport(Rst, "RptOtraSalida", , App.Path + "\Reportes\")
        Set Rst = Nothing
    Case KEY_BROWSER
      Procedencia = Buscar
      Me.Hide
      FrmBuscarDocumentos.Show
  End Select
End Sub

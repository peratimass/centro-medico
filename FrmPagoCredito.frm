VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPagoCredito 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   13800
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdBuscar 
      Height          =   495
      Left            =   8760
      Picture         =   "FrmPagoCredito.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton CmdDeletePago 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Delete"
      Height          =   495
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   6360
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCredito.frx":030A
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCredito.frx":075E
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCredito.frx":0A7E
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCredito.frx":0ED2
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCredito.frx":1326
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCredito.frx":1646
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCredito.frx":1966
            Key             =   "(Monto)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   3945
      Left            =   12720
      TabIndex        =   2
      Top             =   480
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   6959
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   3945
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
         TabIndex        =   3
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1402
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Pagar   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Monto)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Modificar"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgFacturas 
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   5953
      _Version        =   393216
      BackColor       =   -2147483629
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorSel    =   255
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetallePagos 
      Height          =   2085
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   3678
      _Version        =   393216
      BackColor       =   -2147483629
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      ForeColorSel    =   8388608
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   360
      Left            =   6600
      TabIndex        =   6
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   0
      CalendarForeColor=   255
      Format          =   61538307
      CurrentDate     =   39535
   End
End
Attribute VB_Name = "FrmPagoCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ProcedenciaFacturas As EnumFactura
Public Procedencia As EnumProcede
Dim estado As String
Public Sub facturas()

Dim anul As String
anul = "F"


strCadena = "SELECT DocumentoVenta.cPersona, DocumentoVenta.Persona, DocumentoVenta.doc_cod, Comprobantes.doc_abrev, " & _
            "(DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta) as Numero, DocumentoVenta.dEmisionVenta as Emision, DocumentoVenta.dVencimiento as Vencimiento," & _
            "DocumentoVenta.nTotalVenta as Total, DocumentoVenta.Saldo, Estado.sDescripcion as Estado " & _
            "FROM DocumentoVenta INNER JOIN Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
            "Estado ON DocumentoVenta.cEstado = Estado.cEstado WHERE " & _
            "DocumentoVenta.idformapago='" & Trim(KEY_CREDITO) & "' AND DocumentoVenta.anulado='" & Trim(anul) & "'"
    
    Call llenarGrid(Me.HfgFacturas, Me)

End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal formulario As Form)
On Error GoTo salir
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  If rst.RecordCount < 1 Then
    MsgBox "No hay Datos con los Parametros Ingresado", vbInformation, "Mensaje para el Usuario"
    Me.HfgFacturas.Clear
    Set rst = Nothing
    Exit Sub
  End If
  Grilla.Rows = rst.RecordCount + 1
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 0
  Grilla.ColWidth(1) = 3000
  Grilla.ColWidth(2) = 0
  Grilla.ColWidth(3) = 1200
  Grilla.ColWidth(4) = 1700
  Grilla.ColWidth(5) = 1300
  Grilla.ColWidth(6) = 1300
  
Call DarFormatoFecha(Grilla, 5)
Call DarFormatoFecha(Grilla, 6)
Call DarFormato(Grilla, 7)
Call DarFormato(Grilla, 8)

Set rst = Nothing
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Sub llenarGrid_PagosLetras(ByVal Grilla As MSHFlexGrid, ByVal TipoDoc As String, ByVal NumeroVenta As String, ByVal serie As String, ByVal Persona As String)

strCadena = "SELECT DetallePagoCreditos.IdDetalleCredito AS NUM,DetallePagoCreditos.doc_cod,DetallePagoCreditos.cPersona, DetallePagoCreditos.FechaPago AS FECHA, Comprobantes.doc_abrev AS TIPODOC, (DetallePagoCreditos.Serie +'-'+ DetallePagoCreditos.Numero) AS NUMERO ," & _
"DetallePagoCreditos.Monto AS MONTO,DetallePagoCreditos.TipoDocVenta,DetallePagoCreditos.SerieVenta,DetallePagoCreditos.NUmeroVenta FROM DetallePagoCreditos INNER JOIN " & _
"Comprobantes ON DetallePagoCreditos.doc_cod = Comprobantes.doc_cod " & _
"WHERE (DetallePagoCreditos.NumeroVenta='" & NumeroVenta & "'  AND DetallePagoCreditos.SerieVenta='" & serie & "' AND DetallePagoCreditos.cPersona='" & Persona & "' AND DetallePagoCreditos.TipoDocVenta='" & Trim(TipoDoc) & "' )"
'On Error GoTo salir
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.Rows = rst.RecordCount
  If rst.RecordCount > 0 Then
      Set Grilla.Recordset = rst
      Grilla.ColWidth(0) = 600
      Grilla.ColWidth(1) = 0
      Grilla.ColWidth(2) = 0
      Grilla.ColWidth(3) = 2000
      Grilla.ColWidth(4) = 1300
      Grilla.ColWidth(5) = 1500
      Grilla.ColWidth(6) = 1500
      Grilla.ColWidth(7) = 0
      Grilla.ColWidth(8) = 0
      Grilla.ColWidth(9) = 0
      Call DarFormatoFecha(Grilla, 3)
      Call DarFormato(Grilla, 6)
    End If
      Set rst = Nothing

  'Me.TlbAcciones.Buttons(KEY_PRINT).Enabled = True
  'Me.TlbAcciones.Buttons(KEY_EXIT).Enabled = True
  Exit Sub
'salir: MsgBox "No hay Pagos de este comprobante", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub CmdBuscar_Click()
Call facturas
End Sub

Private Sub CmdDeletePago_Click()
Dim Monto As Double
Dim Deuda As Double
Dim Saldo As Double
Dim serie As String
Dim TipoDoc As String
    Dim Factura As String
    Dim Persona As String
If MsgBox("Esta Seguro de Eliminar este Pago", vbInformation + vbYesNo, KEY_EMPRESA) = vbYes Then
    Me.HfgDetallePagos.Col = 0
    strCadena = "SELECT Monto,cPersona,TipoDocVenta,SerieVenta,NumeroVenta FROM DetallePagoCREDITOS WHERE IdDetalleCredito='" & Val(Me.HfgDetallePagos.Text) & "'"
    Call ConfiguraRst(strCadena)
    Monto = rst(0)
    serie = rst(3)
    Factura = rst(4)
    TipoDoc = rst(2)
    Persona = rst(1)
    Set rst = Nothing
    strCadena = "SELECT Saldo FROM DocumentoVenta WHERE " & _
    "sSerie='" & Trim(serie) & "' AND cDocumentoVenta='" & Trim(Factura) & "' AND cPersona='" & Persona & "' AND doc_cod='" & Trim(TipoDoc) & "'"
    Call ConfiguraRst(strCadena)
    Saldo = rst(0)
    Set rst = Nothing
    
    strCadena = "UPDATE DocumentoVenta SET Saldo='" & (Saldo + Monto) & "',cEstado='" & KEY_PENDIENTE & "' WHERE " & _
    "sSerie='" & Trim(serie) & "' AND cDocumentoVenta='" & Trim(Factura) & "'AND cPersona='" & Persona & "' AND doc_cod='" & Trim(TipoDoc) & "'"
    Call EjecutaRST(strCadena)
    'Eliminar DocumentoVenta De Pago A credito
    Me.HfgDetallePagos.Col = 0
    strCadena = "SELECT cPersona,doc_cod,Serie,Numero FROM DetallePagoCREDITOS WHERE IdDetalleCredito='" & Val(Me.HfgDetallePagos.Text) & "'"
    Call ConfiguraRst(strCadena)
    Persona = rst(0)
    TipoDoc = rst(1)
    serie = rst(2)
    Factura = rst(3)
    Set rst = Nothing
    strCadena = "DELETE FROM DocumentoVenta WHERE doc_cod='" & Trim(TipoDoc) & "'AND sSerie='" & Trim(serie) & "' AND cDocumentoVenta='" & Trim(Factura) & "' AND cPersona='" & Trim(Persona) & "'"
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
    'End Eliminar
    
    strCadena = "DELETE FROM DetallePagoCreditos WHERE IdDetalleCredito='" & Val(Me.HfgDetallePagos.Text) & "'"
    Call EjecutaRST(strCadena)
    
    Call facturas
    Call HfgFacturas_Click
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 200
Me.DtpFecha.Value = CVDate(Date)
Call facturas
End Sub

Private Sub HfgFacturas_Click()
Dim serie As String
Dim Numero As String
Dim Persona As String
Dim TipoDoc As String
Me.HfgFacturas.Col = 0
Persona = Trim(Me.HfgFacturas.Text)
Me.HfgFacturas.Col = 2
TipoDoc = Me.HfgFacturas.Text
Me.HfgFacturas.Col = 4
serie = Left(Trim(Me.HfgFacturas.Text), 4)
Numero = Right(Trim(Me.HfgFacturas.Text), 10)
Call llenarGrid_PagosLetras(Me.HfgDetallePagos, TipoDoc, Numero, serie, Persona)

Exit Sub

End Sub
Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_NEW
      Procedencia = Nuevo
      FrmPagoCuotaDeuda.TxtMonto.SetFocus
      FrmPagoCuotaDeuda.Show
     
    Case KEY_UPDATE
      MsgBox "Para Modificar este Comprobante ingrese al Modulo de ventas", vbQuestion, MSGVERIFICACION
      Exit Sub
    Case KEY_DELETE
             
      MsgBox "Para Eliminar este Comprobante ingrese al Modulo de ventas", vbQuestion, MSGVERIFICACION
        Exit Sub
        
    Case KEY_EXIT
        Unload Me
    Case "(Salir)"
      Unload Me
  End Select
End Sub



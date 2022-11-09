VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmDetalleLoteGuia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Lote de Productos"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   10830
   Begin VB.CommandButton CmdBuscar 
      Height          =   615
      Left            =   4440
      Picture         =   "FrmDetalleLoteGuia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Buscar"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox TxtProducto 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   6960
      TabIndex        =   19
      Top             =   1200
      Width           =   3555
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDistribuidora 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3413
      _Version        =   393216
      ForeColor       =   -2147483635
      FixedCols       =   0
      ForeColorFixed  =   -2147483635
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox TxtObservacion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   6960
      MaxLength       =   100
      TabIndex        =   7
      Top             =   3600
      Width           =   3555
   End
   Begin VB.TextBox TxtPrecioCompra 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   6960
      TabIndex        =   9
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CheckBox ChkVencimiento 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vencimiento"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   5580
      TabIndex        =   8
      Top             =   1710
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox TxtPrecioSugerido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   6960
      TabIndex        =   6
      Top             =   3135
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DtpVencimiento 
      Height          =   315
      Left            =   6960
      TabIndex        =   4
      Top             =   1680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      CalendarForeColor=   -2147483635
      CalendarTitleForeColor=   -2147483635
      Format          =   61538305
      CurrentDate     =   37326
   End
   Begin MSDataListLib.DataCombo DtcEstado 
      Height          =   315
      Left            =   6960
      TabIndex        =   5
      Top             =   2160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ForeColor       =   -2147483635
      Text            =   "DataCombo1"
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   7200
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteGuia.frx":030A
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteGuia.frx":0626
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteGuia.frx":0A86
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteGuia.frx":0EE6
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteGuia.frx":1202
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteGuia.frx":1662
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteGuia.frx":197E
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteGuia.frx":1DDE
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteGuia.frx":223E
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteGuia.frx":2B1E
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteGuia.frx":2E3A
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteGuia.frx":3156
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   9120
      TabIndex        =   10
      Top             =   5130
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1482
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1500
      _CBHeight       =   840
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   780
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   1376
         ButtonWidth     =   1058
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Grabar"
               Key             =   "(Grabar)"
               Object.ToolTipText     =   "Grabar Ctrl+G"
               ImageKey        =   "(Grabar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Salir)"
               Object.ToolTipText     =   "Salir"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      CalendarForeColor=   -2147483635
      CalendarTitleForeColor=   -2147483635
      Format          =   61538305
      CurrentDate     =   37326
      MinDate         =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   3195
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4895
      _Version        =   393216
      ForeColor       =   -2147483635
      FixedCols       =   0
      ForeColorFixed  =   -2147483635
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label LblDetalle 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detalle :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   585
   End
   Begin VB.Label LblFecha 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   420
      Width           =   540
   End
   Begin VB.Label LblPrecioSugerido 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Precio Sugerido :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   5580
      TabIndex        =   16
      Top             =   3195
      Width           =   1215
   End
   Begin VB.Label LblProducto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Producto :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   5580
      TabIndex        =   15
      Top             =   1260
      Width           =   735
   End
   Begin VB.Label LblPrecioCompra 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Precio de compra :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   5580
      TabIndex        =   14
      Top             =   2700
      Width           =   1335
   End
   Begin VB.Label LblObservacion 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Observación :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   5580
      TabIndex        =   13
      Top             =   3660
      Width           =   990
   End
   Begin VB.Label LblEstado 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Estado :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   5580
      TabIndex        =   12
      Top             =   2220
      Width           =   585
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00A56E32&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   3780
      Left            =   5280
      Top             =   840
      Width           =   5415
   End
End
Attribute VB_Name = "FrmDetalleLoteGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Producto As String, estado As String

Private Sub Limpia()
  txtProducto.Text = ""
  TxtPrecioCompra.Text = ""
  TxtPrecioSugerido.Text = ""
  TxtObservacion.Text = ""
End Sub

Private Sub CmdBuscar_Click()
  strCadena = "SELECT cfactura as Código, srazonsocial as Distribuidora, " & _
  " ntotalfactura as Total, distribuidora.cdistribuidora FROM FacturaCompra " & _
  " INNER JOIN Distribuidora ON distribuidora.cdistribuidora= " & _
  " facturacompra.cdistribuidora WHERE demisionfactura = cdate('" & DtpFecha.Value & "') " & _
  " ORDER BY srazonsocial"
  Call ConfiguraRst(strCadena)
  Set HfdDistribuidora.Recordset = rst
  Call DarFormato(HfdDistribuidora, 2)
  HfdDistribuidora.ColWidth(0) = 1100
  HfdDistribuidora.ColWidth(1) = 2500
  HfdDistribuidora.ColWidth(3) = 0
  Set rst = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
  strCadena = "SELECT cestado as Codigo, sdescripcion as Descripcion FROM " & _
  " estado WHERE cestado LIKE 'L%' ORDER BY sdescripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(DtcEstado)
  DtpVencimiento.Value = Date
  DtpFecha.Value = Date
End Sub

Private Sub HfdDetalle_Click()
  If HfdDetalle.Row <> 0 Then
    HfdDetalle.Col = 0
    txtProducto.Text = HfdDetalle.Text
    HfdDetalle.Col = 2
    TxtPrecioCompra.Text = Format(HfdDetalle.Text, "#,##0.00")
    HfdDetalle.Col = 3
    Producto = HfdDetalle.Text
  End If
End Sub

Private Sub HfdDetalle_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If HfdDetalle.Row <> 0 Then
      HfdDetalle.Col = 0
      txtProducto.Text = HfdDetalle.Text
      HfdDetalle.Col = 2
      TxtPrecioCompra.Text = Format(HfdDetalle.Text, "#,##0.00")
      HfdDetalle.Col = 3
      Producto = HfdDetalle.Text
    End If
  End If
End Sub

Private Sub HfdDistribuidora_Click()
Dim StrCodDistribuidora As String
Dim StrCodFactura As String
    If HfdDistribuidora.Row <> 0 Then
        HfdDistribuidora.Col = 0
        StrCodFactura = HfdDistribuidora.Text
        HfdDistribuidora.Col = 3
        StrCodDistribuidora = HfdDistribuidora.Text
        strCadena = "SELECT sdescripcionproducto as Descripción, ncantidadfactura as " & _
        " Cant, npreciocompra as PCompra,Producto.cproducto FROM FacturaCompra " & _
        " INNER JOIN (DetalleFacturaCompra INNER JOIN Producto ON " & _
        " Producto.cproducto=DetalleFacturaCompra.cproducto) ON " & _
        " DetalleFacturaCompra.cfactura= FacturaCompra.cfactura WHERE " & _
        " DetalleFacturaCompra.cfactura= '" & StrCodFactura & "' AND " & _
        " DetalleFacturaCompra.cdistribuidora= '" & StrCodDistribuidora & "'  " & _
        " ORDER BY sdescripcionproducto"
        Call ConfiguraRst(strCadena)
        Set HfdDetalle.Recordset = rst
        Call DarFormato(HfdDetalle, 2)
        HfdDetalle.ColWidth(0) = 2500
        HfdDetalle.ColWidth(3) = 0
        Set rst = Nothing
    End If
End Sub

Private Sub HfdDistribuidora_KeyPress(KeyAscii As Integer)
Dim StrCodDistribuidora As String
Dim StrCodFactura As String
  If KeyAscii = 13 Then
    If HfdDistribuidora.Row <> 0 Then
      HfdDistribuidora.Col = 0
      StrCodFactura = HfdDistribuidora.Text
      HfdDistribuidora.Col = 3
      StrCodDistribuidora = HfdDistribuidora.Text
      strCadena = "SELECT sdescripcionproducto as Descripción, ncantidadfactura as " & _
      " Cant, npreciocompra as PCompra,Producto.cproducto FROM FacturaCompra " & _
      " INNER JOIN (DetalleFacturaCompra INNER JOIN Producto ON " & _
      " Producto.cproducto=DetalleFacturaCompra.cproducto) ON " & _
      " DetalleFacturaCompra.cfactura= FacturaCompra.cfactura WHERE " & _
      " DetalleFacturaCompra.cfactura= '" & StrCodFactura & "' AND " & _
      " DetalleFacturaCompra.cdistribuidora= '" & StrCodDistribuidora & "'  " & _
      " ORDER BY sdescripcionproducto"
      Call ConfiguraRst(strCadena)
      Set HfdDetalle.Recordset = rst
      Call DarFormato(HfdDetalle, 2)
      HfdDetalle.ColWidth(0) = 2500
      HfdDetalle.ColWidth(3) = 0
      Set rst = Nothing
    End If
  End If
End Sub

Private Sub Save()
  If txtProducto.Text = "" Or TxtPrecioSugerido.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    estado = Replace(DtcEstado.BoundText, "'", "''")
    TxtPrecioCompra.Text = CDbl(TxtPrecioCompra.Text)
    TxtPrecioSugerido.Text = CDbl(TxtPrecioSugerido.Text)
    If ChkVencimiento.Value = 0 Then
      DtpVencimiento.Value = DTEMAXIMA
      StrTipo = "S"
    Else
      StrTipo = "C"
    End If
    
  End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.Key
    Case KEY_SAVE
      Call Save
    Case KEY_EXIT
        Unload Me
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub TxtPrecioSugerido_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("D", KeyAscii)
End Sub

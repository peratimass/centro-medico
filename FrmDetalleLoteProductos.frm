VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDetalleLoteProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Lote de Productos"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7905
   Begin VB.TextBox TxtPrecioSugerido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   6029
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CheckBox ChkVencimiento 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vencimiento"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4124
      TabIndex        =   11
      Top             =   975
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DtpVencimiento 
      Height          =   315
      Left            =   5429
      TabIndex        =   2
      Top             =   945
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      CalendarForeColor=   -2147483635
      CalendarTitleForeColor=   -2147483635
      Format          =   24510465
      CurrentDate     =   37326
      MinDate         =   2
   End
   Begin MSDataListLib.DataCombo DtcEstado 
      Height          =   315
      Left            =   1889
      TabIndex        =   1
      Top             =   945
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ForeColor       =   -2147483635
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DtcProducto 
      Height          =   315
      Left            =   1889
      TabIndex        =   0
      Top             =   465
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ForeColor       =   -2147483635
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox TxtPrecioCompra 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   1889
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
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
      Left            =   1889
      MaxLength       =   100
      TabIndex        =   5
      Top             =   1905
      Width           =   5595
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   5152
      Top             =   3105
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
            Picture         =   "FrmDetalleLoteProductos.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteProductos.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteProductos.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteProductos.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteProductos.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteProductos.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteProductos.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteProductos.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteProductos.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteProductos.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteProductos.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLoteProductos.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   5932
      TabIndex        =   6
      Top             =   2835
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1482
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1755
      _CBHeight       =   840
      _Version        =   "6.7.8862"
      Child1          =   "TlbAcciones"
      MinHeight1      =   780
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1376
         ButtonWidth     =   1296
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
               Caption         =   "Cancelar"
               Key             =   "(Cancelar)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Cancelar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Label LblPrecioSugerido 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Precio Sugerido :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4124
      TabIndex        =   13
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label LblEstado 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Estado :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   404
      TabIndex        =   12
      Top             =   1005
      Width           =   585
   End
   Begin VB.Label LblObservacion 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Observación :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   405
      TabIndex        =   10
      Top             =   1965
      Width           =   990
   End
   Begin VB.Label LblPrecioCompra 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Precio de compra :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   404
      TabIndex        =   9
      Top             =   1500
      Width           =   1335
   End
   Begin VB.Label LblProducto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Producto :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   404
      TabIndex        =   8
      Top             =   525
      Width           =   735
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00A56E32&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   2460
      Left            =   217
      Top             =   225
      Width           =   7455
   End
End
Attribute VB_Name = "FrmDetalleLoteProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StrCodTabla As String
Dim Producto As String, Estado As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
  DtpVencimiento.Value = Date
  StrCadena = "SELECT cproducto as Codigo,sdescripcionproducto as Descripcion " & _
  " FROM producto ORDER BY sdescripcionproducto"
  Call ConfiguraRst(StrCadena)
  Call LlenaDataCombo(DtcProducto)
  StrCadena = "SELECT cestado as Codigo, sdescripcion as Descripcion FROM " & _
  " estado WHERE cestado LIKE 'L%' ORDER BY sdescripcion"
  Call ConfiguraRst(StrCadena)
  Call LlenaDataCombo(DtcEstado)
  Select Case FrmLoteProductos.Procedencia
    Case Modificar
      Call Llena
  End Select
End Sub

Private Sub Llena()
  FrmLoteProductos.HfdGrilla.Col = 0
  StrCadena = "SELECT clote, sdescripcionproducto, stipovencimiento, dvencimiento," & _
  " npreciocompra, npreciosugerido, loteproducto.sobservacion, estado.sdescripcion " & _
  " FROM producto INNER JOIN (loteproducto INNER JOIN Estado ON " & _
  " estado.cestado=loteproducto.cestado) ON loteproducto.cproducto= " & _
  " producto.cproducto WHERE clote = '" & FrmLoteProductos.HfdGrilla.Text & "'"
  Call EjecutaRST(StrCadena)
  StrCodTabla = RstEjecuta(0)
  DtcProducto.Text = RstEjecuta(1)
  DtpVencimiento.Value = RstEjecuta(3)
  TxtPrecioCompra.Text = RstEjecuta(4)
  TxtPrecioSugerido.Text = RstEjecuta(5)
  TxtObservacion.Text = RstEjecuta(6)
  DtcEstado.Text = RstEjecuta(7)
  If RstEjecuta!stipovencimiento = "S" Then
    ChkVencimiento.Value = 0
  Else
    ChkVencimiento.Value = 1
  End If
  Set RstEjecuta = Nothing
End Sub

Private Sub Save()
  If TxtPrecioCompra.Text = "" Or TxtPrecioSugerido.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    Estado = Replace(DtcEstado.BoundText, "'", "''")
    Producto = Replace(DtcProducto.BoundText, "'", "''")
    TxtPrecioCompra.Text = CDbl(TxtPrecioCompra.Text)
    TxtPrecioSugerido.Text = CDbl(TxtPrecioSugerido.Text)
    If ChkVencimiento.Value = 0 Then
      DtpVencimiento.Value = DTEMAXIMA
      StrTipo = "S"
    Else
      StrTipo = "C"
    End If
    Select Case FrmLoteProductos.Procedencia
      Case Nuevo
        StrCadena = "SELECT clote FROM loteproducto ORDER BY clote DESC"
        Call ConfiguraRst(StrCadena)
        StrCadena = "INSERT INTO loteproducto (clote, cproducto, stipovencimiento, " & _
        " dvencimiento, npreciocompra, npreciosugerido, sobservacion, cestado) " & _
        " VALUES ('" & GeneraCodigo(8) & "','" & Producto & "','" & StrTipo & "'," & _
        " cdate('" & DtpVencimiento.Value & "')," & TxtPrecioCompra.Text & ", " & _
        " " & TxtPrecioSugerido.Text & ",'" & TxtObservacion.Text & "','" & Estado & "')"
        Call EjecutaRST(StrCadena)
        Set RstEjecuta = Nothing
        Unload Me
      Case Modificar
        StrCadena = "UPDATE loteproducto SET cproducto = '" & Producto & "', " & _
        " cestado='" & Estado & "',stipovencimiento = '" & StrTipo & "',dvencimiento=" & _
        " cdate('" & DtpVencimiento.Value & "'),npreciocompra=" & _
        " " & TxtPrecioCompra.Text & ",npreciosugerido=" & TxtPrecioSugerido.Text & "," & _
        " sobservacion='" & TxtObservacion.Text & "' WHERE clote= '" & StrCodTabla & "'"
        Call EjecutaRST(StrCadena)
        Set RstEjecuta = Nothing
        Unload Me
    End Select
    StrCadena = "SELECT count(cestado) FROM loteproducto WHERE cestado = 'LS' " & _
    " AND cproducto = '" & Producto & "'"
    Call EjecutaRST(StrCadena)
    If RstEjecuta(0) > 1 Then
      MsgBox MSGSTAND, vbInformation, MSGVALIDACION
    End If
    Set RstEjecuta = Nothing
  End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.Key
    Case KEY_SAVE
      Call Save
    Case KEY_CANCEL
        Unload Me
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub TxtPrecioCompra_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("D", KeyAscii)
End Sub

Private Sub TxtPrecioSugerido_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("D", KeyAscii)
End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmDetalleOtraEntrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Otra Entrada"
   ClientHeight    =   7335
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   9945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9945
   Begin VB.TextBox TxtGlosa 
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
      Height          =   675
      Left            =   1020
      MaxLength       =   100
      TabIndex        =   6
      Top             =   5520
      Width           =   5235
   End
   Begin VB.TextBox TxtCantidad 
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
      Left            =   1260
      TabIndex        =   5
      Top             =   2700
      Width           =   960
   End
   Begin VB.TextBox TxtProducto 
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
      Left            =   1260
      MaxLength       =   80
      TabIndex        =   3
      Top             =   2295
      Width           =   3075
   End
   Begin VB.TextBox TxtFecha 
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
      Left            =   5025
      TabIndex        =   11
      Top             =   150
      Width           =   1215
   End
   Begin VB.CommandButton CmdEntidad 
      Caption         =   "..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "Busca Cliente"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox TxtEntidad 
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
      Left            =   1260
      TabIndex        =   0
      Top             =   975
      Width           =   3075
   End
   Begin VB.Frame FrmEntidad 
      Height          =   1335
      Left            =   4575
      TabIndex        =   9
      Top             =   720
      Width           =   1665
      Begin VB.OptionButton OptPersona 
         Caption         =   "Otra Entidad"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton OptCliente 
         Caption         =   "Cliente"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptDistribuidora 
         Caption         =   "Distribuidora"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1275
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   2160
      Left            =   150
      TabIndex        =   12
      Top             =   3255
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   3810
      _Version        =   393216
      ForeColor       =   -2147483635
      FixedCols       =   0
      ForeColorFixed  =   -2147483635
      ForeColorSel    =   16777215
      GridColor       =   -2147483635
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdProducto 
      Height          =   7080
      Left            =   6390
      TabIndex        =   4
      Top             =   135
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   12488
      _Version        =   393216
      ForeColor       =   -2147483635
      FixedCols       =   0
      ForeColorFixed  =   -2147483635
      ForeColorSel    =   16777215
      GridColor       =   -2147483635
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   2715
      TabIndex        =   13
      Top             =   6375
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1482
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1800
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
         TabIndex        =   14
         Top             =   30
         Width           =   1680
         _ExtentX        =   2963
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
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   315
      Top             =   6510
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
            Picture         =   "FrmDetalleOtraEntrada.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleOtraEntrada.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleOtraEntrada.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleOtraEntrada.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleOtraEntrada.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleOtraEntrada.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleOtraEntrada.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleOtraEntrada.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleOtraEntrada.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleOtraEntrada.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleOtraEntrada.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleOtraEntrada.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbControl 
      Height          =   900
      Left            =   4575
      TabIndex        =   15
      Top             =   2175
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1588
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1665
      _CBHeight       =   900
      _Version        =   "6.7.8862"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbControl 
         Height          =   780
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1376
         ButtonWidth     =   1217
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar"
               Key             =   "(Agregar)"
               Object.ToolTipText     =   "Agregar"
               ImageKey        =   "(Agregar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Quitar"
               Key             =   "(Quitar)"
               Object.ToolTipText     =   "Quitar"
               ImageKey        =   "(Quitar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComCtl2.UpDown UdCantidad 
      Height          =   315
      Left            =   2205
      TabIndex        =   17
      Top             =   2700
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "TxtCantidad"
      BuddyDispid     =   196610
      OrigLeft        =   3900
      OrigTop         =   3075
      OrigRight       =   4140
      OrigBottom      =   3450
      Max             =   999
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSDataListLib.DataCombo DtcTipoMovimiento 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   1440
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ForeColor       =   -2147483635
      Text            =   "DataCombo1"
   End
   Begin VB.Label LblGlosa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Glosa :"
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
      Height          =   210
      Left            =   240
      TabIndex        =   24
      Top             =   5625
      Width           =   525
   End
   Begin VB.Label LblCantidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad :"
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
      Height          =   210
      Left            =   360
      TabIndex        =   23
      Top             =   2752
      Width           =   735
   End
   Begin VB.Label LblProducto 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto :"
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
      Height          =   210
      Left            =   360
      TabIndex        =   22
      Top             =   2347
      Width           =   765
   End
   Begin VB.Label LblEntidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entidad :"
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
      Height          =   210
      Left            =   360
      TabIndex        =   21
      Top             =   1027
      Width           =   645
   End
   Begin VB.Label LblTipoMovimiento 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Mov. :"
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
      Height          =   225
      Left            =   315
      TabIndex        =   20
      Top             =   1395
      Width           =   885
   End
   Begin VB.Shape ShpProducto 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   150
      Top             =   2055
      Width           =   4335
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha :"
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
      Height          =   210
      Left            =   4470
      TabIndex        =   19
      Top             =   202
      Width           =   555
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      Caption         =   "Otra Entrada de Productos"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   4230
   End
   Begin VB.Shape ShpEntidad 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1080
      Left            =   150
      Top             =   855
      Width           =   4335
   End
End
Attribute VB_Name = "FrmDetalleOtraEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StrCodEntidad As String

Dim StrCodDocumento As String, StrMovimiento As String, TipoEntidad As String
Dim StrCodEstado As String * 2
Dim DteFecha As Date
Dim StrCodigo As String, Producto As String, Unidad As String, StrPersona As String
Dim Procedencia As String
Dim IntCantidad As Integer

Private Sub Almacena()
  RstTemporal.MoveFirst
On Error GoTo Error
  CnBd.BeginTrans
    StrCodDocumento = Numero("OtraEntrada", "E")
    StrCadena = "INSERT INTO OtraEntrada (cOtraEntrada,centidadorigen,sTipoEntidad, sPersona, " & _
    " ctipomovimiento, dOtraEntrada, cestado, sGlosa) VALUES ('" & StrCodDocumento & "', " & _
    " '" & StrCodEntidad & "','" & TipoEntidad & "','" & StrPersona & "','" & StrMovimiento & "'," & _
    " cdate('" & DteFecha & "'),'" & StrCodEstado & "','" & TxtGlosa.Text & "')"
    Call EjecutaRST(StrCadena)
    Do While Not RstTemporal.EOF
      StrCodigo = RstTemporal(0)
      IntCantidad = CDbl(RstTemporal(1))
      '*** Registra Movimiento de Kárdex
      Call Kardex(StrCodigo, StrMovimiento, IntCantidad, DteFecha, StrCodDocumento, 0)
      StrCadena = "INSERT INTO detalleOtraEntrada (cOtraEntrada,cproducto,ncantidadOtraEntrada) " & _
      " VALUES ('" & StrCodDocumento & "','" & StrCodigo & "'," & IntCantidad & ")"
      Call EjecutaRST(StrCadena)
      RstTemporal.MoveNext
    Loop
  CnBd.CommitTrans
  Set RstTemporal = Nothing
  MsgBox "Los registros fueron grabados satisfactoriamente", vbOKOnly, "Grabar"
  Exit Sub
Error:
  CnBd.RollbackTrans
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  MsgBox MSGREINGRESEDATOS, vbInformation + vbOKOnly, MSGGRABACION
  Exit Sub
End Sub

Private Sub CmdEntidad_Click()
  If OptCliente.Value = True Then
    FrmCliente.EnumFrmCliente = COtraEntrada
    FrmBuscarCliente.Show
  End If
  If OptDistribuidora.Value = True Then
    FrmDistribuidora.EnumFrmDistribuidora = DOtraEntrada
    FrmDistribuidora.Show
  End If
End Sub

Private Sub Form_Activate()
  StrCadena = "SELECT cproducto, sdescripcionproducto as Producto,nstockactual " & _
  " as Stock, nprecioventa as Precio FROM producto ORDER BY cproducto"
  Call ConfiguraRst(StrCadena)
  Set HfdProducto.Recordset = Rst
  Set Rst = Nothing
  HfdProducto.ColWidth(0) = 0
  HfdProducto.ColWidth(1) = 2000
  HfdProducto.ColWidth(2) = 450
  HfdProducto.ColWidth(3) = 550
  Call DarFormato(HfdProducto, 3)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
   StrCadena = "SELECT ctipomovimiento as Codigo,sdescripcionmovimiento as Descripcion " & _
  " FROM tipomovimiento WHERE ctipomovimiento LIKE 'E%' AND NOT " & _
  " (ctipomovimiento= 'E01' OR ctipomovimiento= 'E02') ORDER BY sdescripcionmovimiento"
  Call ConfiguraRst(StrCadena)
  Call LlenaDataCombo(DtcTipoMovimiento)
  
  '*** configura un recordset vacio, tomando como referencia los campos de la tabla Detalle
  StrCadena = "SELECT texto as Cod_Prod, entero as Cant,texto as Unid,  texto as " & _
  " Descripcion FROM tabla"
  Call ConfiguraTemporal(StrCadena)
  Set HfdDetalle.Recordset = RstTemporal
  HfdDetalle.ColWidth(0) = 0
  HfdDetalle.ColWidth(1) = 650
  HfdDetalle.ColWidth(2) = 900
  HfdDetalle.ColWidth(3) = 4200
  
  TxtFecha.Text = Date
  Call Limpia(False)
End Sub

Private Sub Limpia(ByVal Flag As Boolean)
  TlbAcciones.Buttons(KEY_CANCEL).Enabled = True
  If RstTemporal.RecordCount > 0 Then
    TlbAcciones.Buttons(KEY_SAVE).Enabled = True
  Else
    TlbAcciones.Buttons(KEY_SAVE).Enabled = False
  End If
  
  TlbControl.Buttons(KEY_AGREGAR).Enabled = Flag
  TlbControl.Buttons(KEY_QUITAR).Enabled = Flag
  
  TxtCantidad.Enabled = Flag
  UdCantidad.Enabled = Flag
  
  If Flag = False Then
    TxtCantidad.Text = ""
    TxtProducto.Text = ""
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FrmCliente.EnumFrmCliente = BuscarCliente
  FrmDistribuidora.EnumFrmDistribuidora = InicioDistribuidora
End Sub

Private Sub HfdDetalle_Click()
 If HfdDetalle.Row <> 0 Then
  HfdDetalle.Col = 0
  StrCodigo = HfdDetalle.Text
  RstTemporal.MoveFirst
  Do While Not RstTemporal.EOF
  If RstTemporal.Fields(0) = StrCodigo Then
    TxtCantidad.Text = RstTemporal.Fields(1).Value
    TxtProducto.Text = RstTemporal.Fields(3).Value
    Exit Do
  End If
  RstTemporal.MoveNext
  Loop
  Set HfdDetalle.Recordset = RstTemporal
  Call Limpia(True)
  Procedencia = "M"
  End If
End Sub

Private Sub HfdProducto_Click()
  If HfdProducto.Row <> 0 Then
    HfdProducto.Col = 0
    StrCodigo = HfdProducto.Text
    If Not StrCodigo = "" Then
      StrCadena = "SELECT sdescripcionproducto, sdescripcion FROM producto INNER JOIN " & _
      " unidad ON unidad.cunidad=producto.cunidad WHERE cproducto = '" & StrCodigo & "'"
      Call EjecutaRST(StrCadena)
      Producto = RstEjecuta(0)
      Unidad = RstEjecuta(1)
      Set RstEjecuta = Nothing
      Procedencia = ""
      TxtProducto.Text = Producto
      TxtCantidad.Text = 1
      Call Limpia(True)
      TlbControl.Buttons(KEY_QUITAR).Enabled = False
    End If
  End If
End Sub

Private Sub HfdProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If HfdProducto.Row <> 0 Then
    HfdProducto.Col = 0
    StrCodigo = HfdProducto.Text
    If Not StrCodigo = "" Then
      StrCadena = "SELECT sdescripcionproducto, sdescripcion FROM producto INNER JOIN " & _
      " unidad ON unidad.cunidad=producto.cunidad WHERE cproducto = '" & StrCodigo & "'"
      Call EjecutaRST(StrCadena)
      Producto = RstEjecuta(0)
      Unidad = RstEjecuta(1)
      Set RstEjecuta = Nothing
      Procedencia = ""
      TxtProducto.Text = Producto
      TxtCantidad.Text = 1
      Call Limpia(True)
      TlbControl.Buttons(KEY_QUITAR).Enabled = False
    End If
  End If
End If
End Sub

Private Sub OptCliente_Click()
  TxtEntidad.Enabled = False
  CmdEntidad.Enabled = True
  TxtEntidad.Text = ""
End Sub

Private Sub OptDistribuidora_Click()
  TxtEntidad.Enabled = False
  CmdEntidad.Enabled = True
  TxtEntidad.Text = ""
End Sub

Private Sub OptPersona_Click()
  TxtEntidad.Enabled = True
  CmdEntidad.Enabled = False
  TxtEntidad.Text = ""
End Sub

Private Sub Save()
  If OptCliente.Value = True Then
    StrPersona = ""
    TipoEntidad = "C"
  End If
  If OptDistribuidora.Value = True Then
    StrPersona = ""
    TipoEntidad = "D"
  End If
  If OptPersona.Value = True Then
    StrCodEntidad = ""
    StrPersona = Left(TxtEntidad.Text, 50)
    TipoEntidad = "P"
  End If
  If StrCodEntidad = "" And StrPersona = "" Then
    MsgBox MSGENTIDAD, vbInformation, MSGVALIDACION
  Else
    DteFecha = TxtFecha.Text
    StrCodEstado = "NN"
    StrMovimiento = Replace(DtcTipoMovimiento.BoundText, "'", "''")
    Call Almacena
    Unload Me
  End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.Key
    Case KEY_SAVE
      Call Save
    Case KEY_CANCEL
      If MsgBox(MSGCANCELAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Unload Me
      End If
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub TlbControl_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Bandera As Boolean
On Error Resume Next
  Bandera = False
  Select Case Button.Key
    Case KEY_AGREGAR
      If TxtEntidad.Text = "" Or TxtCantidad.Text = "" Then
        MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
      Else
        IntCantidad = CInt(TxtCantidad.Text)
        If Procedencia = "M" Then
          RstTemporal.MoveFirst
          Do While Not RstTemporal.EOF
            If RstTemporal(0) = StrCodigo Then
              RstTemporal.Update
              RstTemporal.Fields(1) = CInt(IntCantidad)
              Exit Do
            End If
            RstTemporal.MoveNext
          Loop
          Procedencia = ""
        Else
          If RstTemporal.RecordCount > 0 Then
            RstTemporal.MoveFirst
            Do While Not RstTemporal.EOF
              If RstTemporal(0) = StrCodigo Then
                Bandera = True
                Exit Do
              End If
              RstTemporal.MoveNext
            Loop
          End If
          If Bandera = False Then
            RstTemporal.AddNew
            RstTemporal.Fields(0) = StrCodigo
            RstTemporal.Fields(1) = CInt(IntCantidad)
            RstTemporal.Fields(2) = Trim(Unidad)
            RstTemporal.Fields(3) = Trim(Producto)
          Else
            MsgBox MSGDUPLICIDAD, vbInformation, MSGVALIDACION
          End If
        End If
        Set HfdDetalle.Recordset = RstTemporal
        Call Limpia(False)
      End If
    Case KEY_QUITAR
        RstTemporal.MoveFirst
        Do While Not RstTemporal.EOF
          If RstTemporal.Fields(0) = StrCodigo Then
            RstTemporal.Delete
            Exit Do
          End If
          RstTemporal.MoveNext
        Loop
        Set HfdDetalle.Recordset = RstTemporal
        Call Limpia(False)
    End Select
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("I", KeyAscii)
End Sub

Private Sub TxtProducto_Change()
Dim Criterio As String
  Criterio = Trim(TxtProducto.Text)
  StrCadena = "SELECT cproducto as Código,sdescripcionproducto as Descripción,nstockactual as " & _
  " Stock, nprecioventa AS Precio FROM Producto  WHERE  sdescripcionproducto LIKE " & _
  " '%" & Criterio & "%'  ORDER BY cproducto"
  Call ConfiguraRst(StrCadena)
  Set HfdProducto.Recordset = Rst
  Call DarFormato(HfdProducto, 3)
End Sub

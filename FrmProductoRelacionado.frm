VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmProductoRelacionado 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtBusquedaRapida 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6720
      MaxLength       =   80
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkProducto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PRODUCTO    :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox TxtCodigoBarra 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6255
      MaxLength       =   80
      TabIndex        =   23
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox TxtUtilidad 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3390
      MaxLength       =   80
      TabIndex        =   21
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "QUITAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   20
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "AGREGAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox TxtPrecioCompra 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1800
      MaxLength       =   80
      TabIndex        =   3
      Top             =   2460
      Width           =   975
   End
   Begin VB.TextBox TxtPrecioVenta 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1815
      MaxLength       =   80
      TabIndex        =   2
      Top             =   2130
      Width           =   975
   End
   Begin VB.TextBox TxtCantidadAfectaStock 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6255
      MaxLength       =   80
      TabIndex        =   1
      Top             =   1725
      Width           =   1815
   End
   Begin VB.TextBox TxtDescripcion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1800
      MaxLength       =   80
      TabIndex        =   6
      Top             =   120
      Width           =   6255
   End
   Begin VB.TextBox TxtNombrecomercial 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1800
      MaxLength       =   80
      TabIndex        =   5
      Top             =   525
      Width           =   6255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   3201
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin MSDataListLib.DataCombo DtcUnidad 
      Height          =   315
      Left            =   6240
      TabIndex        =   0
      Top             =   1335
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcLinea 
      Height          =   315
      Left            =   1800
      TabIndex        =   7
      Top             =   1725
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcMarca 
      Height          =   315
      Left            =   1800
      TabIndex        =   8
      Top             =   1335
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   7200
      Top             =   4725
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
            Picture         =   "FrmProductoRelacionado.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoRelacionado.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoRelacionado.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoRelacionado.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoRelacionado.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoRelacionado.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoRelacionado.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoRelacionado.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoRelacionado.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoRelacionado.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoRelacionado.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoRelacionado.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   8640
      TabIndex        =   17
      Top             =   5040
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1155
      _CBHeight       =   870
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   810
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   18
         Top             =   30
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1429
         ButtonWidth     =   1032
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Cancelar)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo DtcProducto 
      Height          =   315
      Left            =   1800
      TabIndex        =   26
      Top             =   960
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO BARRA :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4875
      TabIndex        =   24
      Top             =   2070
      Width           =   1305
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Util"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2925
      TabIndex        =   22
      Top             =   2175
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   5925
      Left            =   0
      Top             =   0
      Width           =   9990
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.COSTO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   930
      TabIndex        =   16
      Top             =   2475
      Width           =   795
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO VENTA :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   555
      TabIndex        =   15
      Top             =   2205
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANT AFECTA STOCK :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4485
      TabIndex        =   14
      Top             =   1755
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE COMERCIAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   75
      TabIndex        =   13
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE REAL :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   465
      TabIndex        =   12
      Top             =   240
      Width           =   1185
   End
   Begin VB.Label LblUnidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNIDAD :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5475
      TabIndex        =   11
      Top             =   1380
      Width           =   705
   End
   Begin VB.Label LblLinea 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLASIFICACION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   435
      TabIndex        =   10
      Top             =   1755
      Width           =   1215
   End
   Begin VB.Label LblLaboratorio 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MARCA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1095
      TabIndex        =   9
      Top             =   1380
      Width           =   555
   End
End
Attribute VB_Name = "FrmProductoRelacionado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strAfectoIGV As String * 2
Dim sub_producto As String
Dim StrPercepcion  As String
Public Procedencia As EnumProcede
Dim strCombo As String
Dim RstAlmProd As New ADODB.Recordset

Private Sub ChkProducto_Click()
If Me.chkProducto.Value = 1 Then
    Me.DtcProducto.Visible = True
    Me.TxtBusquedaRapida.Visible = True
    strCadena = "SELECT id_producto as Codigo,nombre_prod as Descripcion FROM producto WHERE ruc='" & KEY_RUC & "' ORDER BY nombre_prod"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcProducto)
Else
    Me.DtcProducto.Visible = False
    Me.TxtBusquedaRapida.Visible = False
    Me.DtcProducto.Visible = False
End If
End Sub

Private Sub cmdagregar_Click()
Dim cproducto As String
        strCadena = "SELECT * FROM producto WHERE id_relacionado='" & Trim(FrmDetalleProducto.LblCodigoProducto.Caption) & "' AND ruc='" & KEY_RUC & "' AND id_unidad='" & Me.DtcUnidad.BoundText & "' AND cantidad_afecta_stock='" & Val(Me.TxtCantidadAfectaStock.Text) & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            MsgBox "Producto ya registrado, con la Unidad Seleccionada", vbInformation, KEY_EMPRESA
            Exit Sub
        End If
        Call verifica
        cproducto = formato_item(ConsultaUltimoRegistro("producto", "id_producto", "ruc", KEY_RUC), 5)
         If Me.chkProducto.Value = 0 Then
           strCadena = "INSERT INTO producto (id_producto, id_unidad, id_linea,precio_venta,precio_compra, id_marca,nombre_prod,stock_total,stock_minimo,peso,id_percepcion,comentario,id_igv,id_relacionado,cantidad_afecta_stock, " & _
           "id_proveedor,id_auspiciador,id_combo,ruc,precio_delivery,imagen,id_tipo) VALUES ('" & cproducto & "','" & Me.DtcUnidad.BoundText & "','" & Me.DtcLinea.BoundText & "','" & Val(Me.TxtPrecioVenta.Text) & "','" & Val(Me.TxtPrecioCompra.Text) & "','" & Me.DtcMarca.BoundText & "'," & _
           "'" & Me.TxtDescripcion.Text & "','0','0','" & Val(FrmDetalleProducto.TxtPeso.Text) * Val(Me.TxtCantidadAfectaStock.Text) & "','" & StrPercepcion & "'," & _
           "'" & FrmDetalleProducto.TxtObservacion.Text & "','" & strAfectoIGV & "','" & Trim(FrmDetalleProducto.LblCodigoProducto.Caption) & "','" & Val(Me.TxtCantidadAfectaStock.Text) & "','" & Trim(FrmDetalleProducto.dtcProveedor.BoundText) & "','" & FrmDetalleProducto.DtcModelo.BoundText & "','no','" & KEY_RUC & "'," & _
           "'0','','" & FrmDetalleProducto.DtcTipoProducto.BoundText & "')"
        Else
            strCadena = "UPDATE producto SET id_relacionado='" & Trim(FrmDetalleProducto.LblCodigoProducto.Caption) & "',cantidad_afecta_stock='" & Val(Me.TxtCantidadAfectaStock.Text) & "' WHERE id_producto='" & Me.DtcProducto.BoundText & "' AND ruc='" & KEY_RUC & "'"
        End If
           CnBd.Execute (strCadena)
            
            
           strCadena = "SELECT * FROM almacen WHERE ruc='" & KEY_RUC & "'"
           RstAlmProd.CursorLocation = adUseClient
           RstAlmProd.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
           If RstAlmProd.RecordCount <= 0 Then
                MsgBox "No hay Ningun Almacen registrado", vbInformation
                MsgBox "Producto NO Grabado, Cree un Almacen", vbInformation
                Exit Sub
           End If
           RstAlmProd.MoveFirst
           If Me.chkProducto.Value = 0 Then
           For i = 0 To RstAlmProd.RecordCount - 1
             strCadena = "INSERT INTO almacen_producto(id_alm,id_producto,ruc) VALUES ('" & RstAlmProd("id_alm") & "','" & Trim(cproducto) & "','" & Trim(KEY_RUC) & "')"
             CnBd.Execute (strCadena)
              
              
             RstAlmProd.MoveNext
           Next i
           End If
           Set RstAlmProd = Nothing
           If Me.chkProducto.Value = 0 Then
           Call agrega_barra(cproducto)
           End If
           Call llenarGrid(Me.HfdGrilla)
           
End Sub
Private Sub agrega_barra(ByVal cproducto As String)
If Trim(Me.TxtCodigoBarra.Text) <> "" Then
    strCadena = "SELECT * FROM producto_barras WHERE cod_barra='" & Trim(Me.TxtCodigoBarra.Text) & "' AND id_producto='" & Trim(cproducto) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        MsgBox "Codigo de barras ya registrado", vbInformation, KEY_EMPRESA
    Else
        strCadena = "INSERT INTO producto_barras VALUES('" & Trim(cproducto) & "','" & Trim(Me.TxtCodigoBarra.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
         
    End If
        
    
End If
End Sub
Sub verifica()

If FrmDetalleProducto.ChkPercepcion.Value = 1 Then
    StrPercepcion = "si"
Else
    StrPercepcion = "no"
End If

If FrmDetalleProducto.chkSubproductos.Value = 1 Then
    sub_producto = "si"
Else
    sub_producto = "no"
End If
    
    If (FrmDetalleProducto.ChkIGV.Value = 1) Then
        strAfectoIGV = "si"
    Else
        strAfectoIGV = "no"
    End If

End Sub

Private Sub Command1_Click()
strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_producto='" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    MsgBox "Imposible Eliminar este Producto", vbInformation, KEY_EMPRESA
Else
    strCadena = "DELETE FROM producto WHERE id_producto='" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
     
     Call llenarGrid(Me.HfdGrilla)
End If
End Sub

Private Sub DtcUnidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtCantidadAfectaStock)
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 700
  
  Me.TxtDescripcion.Text = FrmDetalleProducto.TxtDescripcion.Text
  Me.TxtNombrecomercial.Text = FrmDetalleProducto.TxtNombrecomercial.Text
  strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLinea)
  Me.DtcLinea.BoundText = FrmDetalleProducto.DtcLinea.BoundText
  Me.DtcLinea.Locked = True
  
  strCadena = "SELECT id_marca as Codigo, descripcion as Descripcion FROM marca WHERE id_usu='" & KEY_RUC & "'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMarca)
  Me.DtcMarca.BoundText = FrmDetalleProducto.DtcMarca.BoundText
  Me.DtcMarca.Locked = True
  
  strCadena = "SELECT id_und as Codigo, abreviatura as Descripcion FROM unidad WHERE id_usu='" & KEY_RUC & "'  ORDER BY abreviatura"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcUnidad)
  Call llenarGrid(Me.HfdGrilla)
  
  
 
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_CANCEL
        Unload Me
        Set rst = Nothing
        Set RstAlmProd = Nothing
End Select
End Sub

Private Sub TxtBusquedaRapida_Change()
    strCadena = "SELECT id_producto as Codigo,nombre_prod as Descripcion FROM producto WHERE ruc='" & KEY_RUC & "' AND nombre_prod LIKE '%" & Trim(Me.TxtBusquedaRapida.Text) & "%' ORDER BY nombre_prod"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcProducto)
End Sub

Private Sub TxtCantidadAfectaStock_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Me.TxtCantidadAfectaStock.Text) > 0 Then
        Call Resalta(Me.TxtPrecioVenta)
    Else
        MsgBox "Ingrese un Valor mayor que 0", vbInformation, KEY_EMPRESA
        Call Resalta(Me.TxtCantidadAfectaStock)
    End If
End If
End Sub

Private Sub TxtCodigoBarra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.CmdAgregar.SetFocus
End If
End Sub

Private Sub TxtPrecioCompra_Change()
'If Val(Me.TxtPrecioCompra.text) > 0 Then
 '  Me.TxtUtilidad.text = Format((Val(Me.TxtPrecioVenta.text) - Val(Me.TxtPrecioCompra.text)) * 100 / Val(Me.TxtPrecioCompra.text), "#,##0.00")
'End If
End Sub

Private Sub TxtPrecioCompra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtPrecioCompra.Text = Format(Me.TxtPrecioCompra.Text, "#,##0.00")
    Call Resalta(Me.TxtCodigoBarra)
End If
End Sub

Private Sub TxtPrecioVenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtUtilidad.Text = "15"
    Me.TxtPrecioVenta.Text = Format(Me.TxtPrecioVenta.Text, "#,##0.00")
    Call Resalta(Me.TxtUtilidad)
End If
End Sub

Private Sub TxtUtilidad_Change()
If Val(Me.TxtPrecioVenta.Text) > 0 Then
    Me.TxtPrecioCompra.Text = Val(Me.TxtPrecioVenta.Text) - Val(Me.TxtPrecioVenta.Text) * Val(Me.TxtUtilidad.Text) / 100
    
End If
End Sub

Private Sub TxtUtilidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtPrecioCompra)
End If
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
strCadena = "SELECT * FROM producto P,unidad U WHERE P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND P.id_relacionado='" & Trim(FrmDetalleProducto.LblCodigoProducto.Caption) & "'"
Call ConfiguraRst(strCadena)
 
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
  

   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 4500
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 2000
           
          Next
         cabecera = "CODIGO" & vbTab & "PRODUCTO" & vbTab & "UNIDAD" & vbTab & "CANTIDAD" & vbTab & "BARRAS"
         Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & rst("cantidad_afecta_stock") & vbTab & BDBuscarCampoRuc("producto_barras", "cod_barra", "id_producto", rst("id_producto"))
                        
          Grilla.AddItem Fila
            
        Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub



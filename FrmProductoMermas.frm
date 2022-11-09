VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmProductoMermas 
   BorderStyle     =   0  'None
   Caption         =   "Detalle Mermas"
   ClientHeight    =   7860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtNumero 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   8565
      TabIndex        =   23
      Top             =   360
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   360
      Left            =   10560
      TabIndex        =   21
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   121176065
      CurrentDate     =   41219
   End
   Begin VB.TextBox TxtPVenta 
      Height          =   285
      Left            =   4320
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   11280
      Picture         =   "FrmProductoMermas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox txtCosto 
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
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitar 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11160
      Picture         =   "FrmProductoMermas.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtDefectuosos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   8280
      MaxLength       =   80
      TabIndex        =   10
      Top             =   2235
      Width           =   1935
   End
   Begin VB.TextBox TxtSerie 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox TxtProducto 
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
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox txtid_producto 
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
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox TxtDetalle 
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
      Height          =   615
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   6000
      Width           =   9135
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   1230
      TabIndex        =   3
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
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
      Left            =   3720
      Top             =   720
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
            Picture         =   "FrmProductoMermas.frx":06D4
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoMermas.frx":09F0
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoMermas.frx":0E50
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoMermas.frx":0EDD
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoMermas.frx":133D
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoMermas.frx":1659
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoMermas.frx":1AB9
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoMermas.frx":1DD5
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoMermas.frx":2235
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoMermas.frx":2695
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoMermas.frx":2F75
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoMermas.frx":3291
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoMermas.frx":35AD
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2490
      Left            =   6960
      TabIndex        =   7
      Top             =   6960
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4392
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   4995
      _CBHeight       =   2490
      _Version        =   "6.0.8169"
      Child1          =   "TlbGrabar"
      MinHeight1      =   2430
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbGrabar 
         Height          =   810
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   1429
         ButtonWidth     =   1667
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Nuevo"
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Grabar Ctrl+G"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Procesar"
               Key             =   "(Grabar)"
               ImageKey        =   "(Grabar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Imprimir"
               Key             =   "(Imprimir)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Imprimir)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Reporte"
               Key             =   "(Reporte)"
               ImageKey        =   "(Buscar)"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "& Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo DtcMerma 
      Height          =   315
      Left            =   8280
      TabIndex        =   13
      Top             =   1680
      Width           =   2895
      _ExtentX        =   5106
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   2895
      Left            =   240
      TabIndex        =   20
      Top             =   2760
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5106
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
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
   Begin MSDataListLib.DataCombo DtcTipodoc 
      Height          =   315
      Left            =   5640
      TabIndex        =   22
      Top             =   360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      Locked          =   -1  'True
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
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANT.MERMA:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6855
      TabIndex        =   19
      Top             =   2280
      Width           =   1185
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7290
      TabIndex        =   18
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.COSTO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6315
      TabIndex        =   17
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   435
      TabIndex        =   11
      Top             =   6240
      Width           =   705
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   3  'Dot
      DrawMode        =   5  'Not Copy Pen
      Height          =   1455
      Left            =   240
      Top             =   1200
      Width           =   11535
   End
   Begin VB.Label lblStock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   7320
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DFDFE0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   915
      Left            =   120
      Top             =   5880
      Width           =   11775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COD PRODUCTO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   495
      TabIndex        =   5
      Top             =   1320
      Width           =   1365
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Almacen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   300
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   5715
      Left            =   120
      Top             =   120
      Width           =   11775
   End
End
Attribute VB_Name = "FrmProductoMermas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Procedencia As EnumProcede
Public cod_producto As String
Dim strMerma As String
Dim costo_total As Single
Dim stock_total As Single


Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Call actualizar_conceptos
End Sub

Private Sub CmdQuitar_Click()
strCadena = "DELETE FROM merma_temporal WHERE id_detalle='" & Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "' AND id_usuario='" & KEY_USUARIO & "' "
CnBd.Execute (strCadena)
 
strCadena = "SELECT T.id_detalle,T.id_producto,P.nombre_prod,U.abreviatura,T.cantidad,T.costo,M.descripcion FROM merma_temporal T,producto P,merma_motivo M ,unidad U WHERE T.id_producto=P.id_producto AND T.id_motivo=M.id_merma AND T.id_usuario='" & KEY_USUARIO & "' AND T.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' "
          Call llenarGrid_prod(Me.HfdDetalle, Me)

End Sub

Private Sub DtcMerma_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtDefectuosos)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub
Public Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub
Private Sub Form_Load()
CenterForm Me
Me.Top = 500

  Call actualizar
 strCadena = "SELECT id_merma as Codigo, descripcion as Descripcion FROM merma_motivo WHERE ruc='" & KEY_RUC & "'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMerma)
  Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
  
  
  
  
End Sub
Sub actualizar_conceptos()
  strCadena = "SELECT id_merma as Codigo, descripcion as Descripcion FROM merma_motivo WHERE ruc='" & KEY_RUC & "'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMerma)
  Me.DtcMerma.SetFocus
  
End Sub
Sub actualizar()
   

  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'  ORDER BY id_alm ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
rebot:
  strCadena = "SELECT * FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' AND A.id_alm='" & KEY_ALM & "' AND A.id_doc='0105'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
    Me.TxtSerie.Text = rst("serie")
    Me.txtnumero.Text = rst("numero")
    strCadena = "SELECT A.id_doc as Codigo,C.doc_abrev as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' AND A.id_alm='" & KEY_ALM & "' AND A.id_doc='0105'"
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcTipoDoc)
  Else
    strCadena = "INSERT INTO almacen_comprobante(ruc,id_alm,id_doc,serie,numero,id_formato_impresion,id_usuario)VALUES('" & KEY_RUC & "','" & KEY_ALM & "','0105','001','000001','1','" & KEY_USUARIO & "')"
    CnBd.Execute (strCadena)
     
    GoTo rebot
  End If
  
  
  
  
  
  Me.DTPicker1.Value = KEY_FECHA
  Me.txtid_producto.Text = ""
  Me.txtProducto.Text = ""
  Me.lblStock.Caption = ""
          strCadena = "SELECT T.id_detalle,T.id_producto,P.nombre_prod,U.abreviatura,T.cantidad,T.costo,M.descripcion FROM merma_temporal T,producto P,merma_motivo M ,unidad U WHERE T.id_producto=P.id_producto AND T.id_motivo=M.id_merma AND T.id_usuario='" & KEY_USUARIO & "' AND T.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' "
          Call llenarGrid_prod(Me.HfdDetalle, Me)

  
End Sub

Private Sub HfdDetalle_Click()
If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    Me.CmdQuitar.Visible = True
Else
    Me.CmdQuitar.Visible = False
End If
End Sub

Private Sub HfdDetalle_SelChange()
If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    Me.CmdQuitar.Visible = True
Else
    Me.CmdQuitar.Visible = False
End If
End Sub

Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim id_merma As Double
Select Case Button.key
    Case KEY_NEW
        Call actualizar
        'Call llenar_grid
    Case KEY_SAVE
       If Me.DtcAlmacen.BoundText <> "" Then
        fecha = Format(Me.DTPicker1.Value, "yyyy-mm-dd")
        
        
        
        strCadena = "INSERT INTO merma(id_doc,serie,numero,id_alm,fecha,detalle,id_usuario,ruc) VALUES ('" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.TxtSerie.Text) & "','" & Trim(Me.txtnumero.Text) & "','" & KEY_ALM & "','" & fecha & "','" & Trim(Me.TxtDetalle.Text) & "','" & Trim(KEY_USUARIO) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        id_merma = LastRegistro("merma", "id_merma")
        strCadena = "SELECT * FROM merma_temporal WHERE ruc='" & KEY_RUC & "' AND id_usuario='" & KEY_USUARIO & "' AND id_alm='" & KEY_ALM & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            For i = 0 To rst.RecordCount - 1
                strCadena = "INSERT INTO merma_detalle(id_merma,id_producto,id_motivo,cantidad,costo,ruc)VALUES('" & id_merma & "','" & rst("id_producto") & "','" & rst("id_motivo") & "','" & rst("cantidad") & "','" & rst("costo") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                 
                rst.MoveNext
            Next i
            strCadena = "DELETE FROM merma_temporal WHERE id_usuario='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
            CnBd.Execute (strCadena)
             
        End If
       Else
       MsgBox "Ingrese Datos Obligatorios", vbInformation, KEY_EMPRESA
       Exit Sub
       End If
       strCadena = "UPDATE almacen_comprobante SET numero='" & formato_item(Val(Me.txtnumero.Text) + 1, 6) & "' WHERE serie='" & Trim(Me.TxtSerie.Text) & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' AND id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
       CnBd.Execute (strCadena)
        
       Me.CmdQuitar.Visible = False
        Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
        Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
        
    Case KEY_PRINT
       ' strCadena = "SELECT Producto_barras.cod_barra, Producto.DescripcionProducto, Unidad.sAbreviatura,Producto_mermas_detalle.costo ,Producto_mermas_detalle.cantidad,Producto_mermas_detalle.detalle, " & _
        "Producto_mermas.fecha , Producto_mermas.detalle, Producto_mermas.cMerma, Producto_mermas.usuario " & _
        "FROM Producto INNER JOIN Producto_barras ON Producto.cProducto = Producto_barras.cProducto INNER JOIN " & _
        "Unidad ON Producto.cUnidad = Unidad.cUnidad INNER JOIN " & _
        "Producto_mermas INNER JOIN Producto_mermas_detalle ON Producto_mermas.cMerma = Producto_mermas_detalle.cMerma ON " & _
        "Producto.cProducto = Producto_mermas_detalle.cProducto WHERE Producto_mermas.cMerma='" & Trim(Me.Txtcodmerma.text) & "'"
        'Call ConfiguraRst(strCadena)
        'Ans = ShowMultiReport(rst, "RptMermas3", , App.Path + "\Reportes\")
    Case "(Reporte)"
        FrmProducto_merma_lista.Show
    Case KEY_EXIT
        Unload Me
End Select
End Sub
Public Sub llenarMerma(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
strCadena = "SELECT * FROM merma M,merma_detalle D,producto WHERE M.id_merma=D.id_merma AND M.id_alm='" & KEY_ALM & "' AND M.ruc='" & KEY_RUC & "' AND "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
     Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
     Me.CmdQuitar.Visible = False
    Exit Sub
End If
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 5000
           Grilla.ColWidth(3) = 800
           Grilla.ColWidth(4) = 1100
           Grilla.ColWidth(5) = 1100
           Grilla.ColWidth(6) = 2500
         Next
        cabecera = "DETALLE" & vbTab & "CODIGO" & vbTab & "DESCRIPCION ARTICULO" & vbTab & "UND" & vbTab & "CANT" & vbTab & "COSTO" & vbTab & "MOTIVO"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_detalle") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("costo"), "#,##0.00") & vbTab & Format(rst("descripcion"), "#,##0.00")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
    Next i
        
  
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
  
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Private Sub TxtCodigo_KeyPress(KeyAscii As Integer)

End Sub
Public Sub llenarGrid_prod(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
     Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
     Me.CmdQuitar.Visible = False
    Exit Sub
End If
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 5000
           Grilla.ColWidth(3) = 800
           Grilla.ColWidth(4) = 1100
           Grilla.ColWidth(5) = 1100
           Grilla.ColWidth(6) = 2500
         Next
        cabecera = "DETALLE" & vbTab & "CODIGO" & vbTab & "DESCRIPCION ARTICULO" & vbTab & "UND" & vbTab & "CANT" & vbTab & "COSTO" & vbTab & "MOTIVO"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_detalle") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("costo"), "#,##0.00") & vbTab & Format(rst("descripcion"), "#,##0.00")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
    Next i
        
  
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
  
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub
Public Sub llenar_merma(ByVal Grilla As MSHFlexGrid, ByVal id_merma As Double)
On Error GoTo salir
  strCadena = "SELECT D.id_detalle_merma,D.id_producto,P.nombre_prod,U.abreviatura,D.cantidad,D.costo,M.descripcion FROM merma ME,merma_detalle D,producto P,merma_motivo M ,unidad U WHERE D.id_producto=P.id_producto AND D.id_motivo=M.id_merma AND ME.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND D.id_merma=ME.id_merma AND ME.id_merma='" & id_merma & "' AND ME.id_alm='" & KEY_ALM & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
     Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
    Exit Sub
End If
   Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
   Me.CmdQuitar.Visible = False
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 5000
           Grilla.ColWidth(3) = 800
           Grilla.ColWidth(4) = 1100
           Grilla.ColWidth(5) = 1100
           Grilla.ColWidth(6) = 2500
         Next
        cabecera = "DETALLE" & vbTab & "CODIGO" & vbTab & "DESCRIPCION ARTICULO" & vbTab & "UND" & vbTab & "CANT" & vbTab & "COSTO" & vbTab & "MOTIVO"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_detalle_merma") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("costo"), "#,##0.00") & vbTab & Format(rst("descripcion"), "#,##0.00")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
    Next i
        
  
  'Grilla.Row = 1
  'Grilla.col = 0
  'Grilla.ColSel = 1
  'Grilla.RowSel = 1
  
  
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub


Private Sub TxtDefectuosos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (Val(Me.TxtDefectuosos.Text) > 0) And Me.txtid_producto.Text <> "" And Val(Me.TxtDefectuosos.Text) > 0 Then
          strCadena = "INSERT INTO merma_temporal(id_producto,cantidad,costo,id_motivo,id_usuario,id_alm,ruc)VALUES('" & Trim(Me.txtid_producto.Text) & "'," & _
          "'" & Val(Me.TxtDefectuosos.Text) & "','" & Val(Me.txtCosto.Text) * Val(Me.TxtDefectuosos.Text) & "','" & Me.DtcMerma.BoundText & "','" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
          CnBd.Execute (strCadena)
           
           Me.txtid_producto.Text = ""
           Me.txtProducto.Text = ""
           Me.lblStock.Caption = ""
           Me.txtCosto.Text = ""
           Me.TxtDefectuosos.Text = ""
           Call Resalta(Me.txtid_producto)
          strCadena = "SELECT T.id_detalle,T.id_producto,P.nombre_prod,U.abreviatura,T.cantidad,T.costo,M.descripcion FROM merma_temporal T,producto P,merma_motivo M ,unidad U WHERE T.id_producto=P.id_producto AND T.id_motivo=M.id_merma AND T.id_usuario='" & KEY_USUARIO & "' AND T.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' "
          Call llenarGrid_prod(Me.HfdDetalle, Me)
        
        
        End If
End If
End Sub



Private Sub txtid_producto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
   If KEY_BARRAS = "si" Then
    strCadena = "SELECT B.id_producto,P.nombre_prod,U.abreviatura,A.stock,P.precio_compra,P.precio_venta FROM producto_barras B,producto P,almacen_producto A,unidad U " & _
    " WHERE B.cod_barra='" & Trim(Me.txtid_producto.Text) & "' AND A.id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND P.ruc='" & KEY_RUC & "' AND B.ruc='" & KEY_RUC & "' " & _
    "AND U.id_usu='" & KEY_RUC & "' AND B.id_producto=P.id_producto AND P.id_unidad=U.id_und"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        cod_producto = rst("id_producto")
        Me.txtProducto.Text = rst("nombre_prod")
        Me.lblStock.Caption = rst("stock")
        Me.txtCosto.Text = rst("precio_compra")
        Call Resalta(Me.TxtDefectuosos)
        Set rst = Nothing
        Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
        
    Else
         Procedencia = mermas
         FrmProducto.Show
    End If
    Else
      Me.txtid_producto.Text = formato_item(Me.txtid_producto.Text, 5)
      strCadena = "SELECT P.id_producto,P.nombre_prod, U.abreviatura,A.stock,P.precio_compra,P.precio_venta FROM producto P,almacen_producto A ,unidad U WHERE A.id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "'" & _
    " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND P.id_producto=A.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "' AND P.id_producto='" & Trim(Me.txtid_producto.Text) & "'"
     Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        cod_producto = rst("id_producto")
        Me.txtProducto.Text = rst("nombre_prod")
        Me.lblStock.Caption = rst("stock")
        Me.txtCosto.Text = rst("precio_compra")
        Me.DtcMerma.SetFocus
        Set rst = Nothing
        Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
        
    Else
         Procedencia = mermas
         FrmProducto.Show
    End If
         
    End If
End If
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtnumero.Text = formato_item(Val(Me.txtnumero.Text), 6)
   Call consulta(Trim(Me.txtnumero.Text))
End If
End Sub
Public Sub consulta(ByVal Numero As String)
strCadena = "SELECT * FROM merma WHERE numero='" & Numero & "' AND id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txtnumero.Text = rst("numero")
    Me.DTPicker1.Value = rst("fecha")
    Me.TxtDetalle.Text = rst("detalle")
    Call llenar_merma(Me.HfdDetalle, rst("id_merma"))
End If



End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDerivados 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Derivados"
   ClientHeight    =   7680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtUnidad 
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
      Height          =   315
      Left            =   10155
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Und"
      Top             =   5865
      Width           =   975
   End
   Begin VB.TextBox TxtStock 
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
      Height          =   315
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   5865
      Width           =   1095
   End
   Begin VB.TextBox TxtCombo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   1
      Top             =   360
      Width           =   2055
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
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5880
      Width           =   3495
   End
   Begin MSDataListLib.DataCombo DtcCombo 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   8880
      TabIndex        =   3
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      CurrentDate     =   39371
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   2070
      TabIndex        =   4
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
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
      Left            =   10320
      Top             =   3480
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
            Picture         =   "FrmDerivados.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDerivados.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDerivados.frx":077C
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDerivados.frx":0809
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDerivados.frx":0C69
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDerivados.frx":0F85
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDerivados.frx":13E5
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDerivados.frx":1701
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDerivados.frx":1B61
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDerivados.frx":1FC1
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDerivados.frx":28A1
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDerivados.frx":2BBD
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDerivados.frx":2ED9
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   4920
      TabIndex        =   5
      Top             =   6480
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   7635
      _CBHeight       =   870
      _Version        =   "6.0.8169"
      Child1          =   "TlbGrabar"
      MinHeight1      =   810
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbGrabar 
         Height          =   810
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   1429
         ButtonWidth     =   1693
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
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
               Caption         =   "  &Grabar  "
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
               Caption         =   "Anular"
               Key             =   "(Anular)"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Reporte"
               Key             =   "(Reporte)"
               ImageKey        =   "(Buscar)"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Control"
               Key             =   "(Control)"
               ImageKey        =   "(Aceptar)"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   4215
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   7435
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
   Begin VB.Label lblAnulado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   1080
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Actual:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   7335
      TabIndex        =   11
      Top             =   5880
      Width           =   1155
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   4920
      Top             =   5760
      Width           =   7575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Combo:"
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
      Left            =   1050
      TabIndex        =   10
      Top             =   720
      Width           =   735
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
      Left            =   1020
      TabIndex        =   9
      Top             =   240
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   3  'Dot
      Height          =   1095
      Left            =   7920
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   8295
      TabIndex        =   8
      Top             =   360
      Width           =   465
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   3  'Dot
      Height          =   1695
      Left            =   120
      Top             =   5760
      Width           =   4695
   End
   Begin VB.Label Label4 
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
      Left            =   285
      TabIndex        =   7
      Top             =   6120
      Width           =   705
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   120
      Top             =   120
      Width           =   12375
   End
End
Attribute VB_Name = "FrmDerivados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strCombo As String
Dim rst2 As New ADODB.Recordset
Public Procedencia As EnumProcede


Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub DtcCombo_Change()
Call LLENA
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
 strCadena = "SELECT id_producto as Codigo, nombre_prod as Descripcion FROM producto WHERE id_sub_producto='si' AND id_combo='no' " & _
  " ORDER BY nombre_prod"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCombo)
  Set rst = Nothing
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'  ORDER BY id_alm ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  
  Me.DTPicker1.Value = KEY_FECHA
  Call nuevo
  
End Sub
Private Sub nuevo()
  strCadena = "SELECT id_derivado FROM derivados ORDER BY id_derivado DESC"
  Call ConfiguraRst(strCadena)
  strCombo = GeneraCodigo(6)
  
  Me.TxtCombo.Text = strCombo
  Me.TxtDetalle.Text = ""
  Me.lblAnulado.Visible = False
  strCadena = "UPDATE producto_sub SET cantidad='0' WHERE id_producto_padre='" & Trim(Me.DtcCombo.BoundText) & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
  CnBd.Execute (strCadena)
   
  Call LLENA
  Set rst = Nothing
  
  Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
  Me.TlbGrabar.Buttons(KEY_ANULAR).Enabled = False
  Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
End Sub
Public Sub LLENA()
       
    strCadena = "SELECT * FROM producto_sub S,producto P,unidad U WHERE S.id_producto=P.id_producto AND P.id_unidad=U.id_und AND S.id_producto_padre='" & Trim(Me.DtcCombo.BoundText) & "' AND S.ruc='" & KEY_RUC & "' AND S.id_alm='" & KEY_ALM & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "'"
    Call llenarGrid_prod(Me.HfdDetalle, Me)
  
End Sub
Public Sub llenarGrid_prod(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 7000
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 1000
           
         Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION ARTICULO" & vbTab & "UND" & vbTab & "CANT"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & Format(rst("cantidad"), "#,##0.00")
            Grilla.AddItem Fila
            For k = 3 To 3
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &HC0FFC0
            Next k
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

Private Sub HfdDetalle_DblClick()
If Me.HfdDetalle.Rows > 0 Then
   frmDerivadosCantidad.Show
End If
End Sub

Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim costo_t As Single
Dim venta_t As Single
Select Case Button.key

    Case KEY_NEW
        Call nuevo
    Case KEY_SAVE
       Dim cantidad As Single
       Dim rst1 As New ADODB.Recordset
       Dim strDetalle As String
        
        fecha = Format(Me.DTPicker1.Value, "yyyy-mm-dd")
        strCombo = formato_item(ConsultaUltimoRegistro("derivados", "id_derivado", "ruc", KEY_RUC), 6)
        
        strCadena = "INSERT INTO derivados(id_derivado,id_doc,id_producto,fecha,detalle,id_usuario,id_alm,ruc) VALUES ('" & strCombo & "','0104','" & Me.DtcCombo.BoundText & "','" & fecha & "','" & Trim(Me.TxtDetalle.Text) & "','" & Trim(KEY_USUARIO) & "','" & KEY_ALM & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        strCadena = "SELECT * FROM producto_sub WHERE id_producto_padre='" & Me.DtcCombo.BoundText & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND cantidad>0"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            For i = 0 To rst.RecordCount - 1
                strCadena = "INSERT INTO derivado_detalle(id_derivado,id_producto,cantidad,id_alm,ruc)VALUES('" & strCombo & "','" & rst("id_producto") & "','" & rst("cantidad") & "','" & KEY_ALM & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                 
                rst.MoveNext
            Next i
        End If
      
      Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
      Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
      Me.TlbGrabar.Buttons(KEY_ANULAR).Enabled = True
   
   Case KEY_PRINT
        strCadena = "SELECT Derivados.cDerivado, Derivados.cProducto, Derivados.anulado, Derivados.fecha, Derivados.descripcion, Derivados.detalle," & _
        "Derivados.estado, Producto_barras.cod_barra, Producto.DescripcionProducto, Unidad.sAbreviatura, Derivado_Detalle.cantidad, " & _
        " Seguridad.Nombre FROM  Derivado_Detalle INNER JOIN Producto ON Derivado_Detalle.cProducto = Producto.cProducto INNER JOIN " & _
        " Producto_barras ON Producto.cProducto = Producto_barras.cProducto INNER JOIN Unidad ON Producto.cUnidad = Unidad.cUnidad INNER JOIN " & _
        " Derivados ON Derivado_Detalle.cDerivado = Derivados.cDerivado INNER JOIN Seguridad ON Derivados.id_usuario = Seguridad.IdUsuario " & _
        "WHERE Derivados.cDerivado='" & Trim(Me.TxtCombo.Text) & "'"
         Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptDerivados", , App.Path + "\Reportes\")
    Case KEY_ANULAR
        Procedencia = anular
        FrmSeguridad.Show
            
    Case "(Control)"
        FrmKardexCarne.Show
    Case "(Reporte)"
            FrmDesrivados_list.Show
    Case KEY_EXIT
        Unload Me
End Select
End Sub
Public Function GeneraCodigo1(ByVal longitud As Integer) As String
Dim x As Integer
Dim Formato As String
  Formato = ""
  For x = 1 To longitud
    Formato = Formato + "0"
  Next x
   
  If (rst2.BOF And rst2.EOF) Then
    StrNumero = Format(str(Val(Formato) + 1), Formato)
  Else
    StrNumero = Format(Trim(str(Val(Right(rst2(0), longitud + 1)) + 1)), Formato)
  End If
  Set rst2 = Nothing
  GeneraCodigo1 = Gencodigo + StrNumero
  Gencodigo = ""

End Function




Private Sub TxtCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
 Call consultar(FormatosCeros(Me.TxtCombo.Text, 6))
   
End If
End Sub
Public Sub consultar(ByVal combo As String)
    Me.TxtCombo.Text = combo
    strCadena = "SELECT * FROM derivados WHERE id_derivado='" & combo & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If rst("anulado") = "si" Then
            Me.TlbGrabar.Buttons(KEY_ANULAR).Enabled = False
            Me.lblAnulado.Visible = True
        Else
        Me.TlbGrabar.Buttons(KEY_ANULAR).Enabled = True
        Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
        Me.lblAnulado.Visible = False
    End If
    End If
    strCadena = "SELECT D.id_producto,P.nombre_prod,U.abreviatura,D.cantidad FROM derivados C,derivado_detalle D,producto P,unidad U WHERE C.id_derivado=D.id_derivado AND D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND C.id_alm='" & KEY_ALM & "' AND C.ruc='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_ALM & "' AND C.id_derivado='" & combo & "'"
    Call llenarGrid_prod(Me.HfdDetalle, Me)
    
End Sub


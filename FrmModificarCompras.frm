VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmModificarCompras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Ingreso de Mercaderia"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtNumero_Ref 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3405
      MaxLength       =   10
      TabIndex        =   29
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox TxtSerie_Ref 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2520
      MaxLength       =   80
      TabIndex        =   28
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CmdActualizar 
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   27
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox TxtCosto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   8295
      MaxLength       =   80
      TabIndex        =   21
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox TxtUnidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   7425
      MaxLength       =   80
      TabIndex        =   20
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox TxtCodProducto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   600
      MaxLength       =   80
      TabIndex        =   19
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtDescripcionProducto 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3075
      MaxLength       =   80
      TabIndex        =   18
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox TxtCantidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1965
      MaxLength       =   80
      TabIndex        =   17
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox TxtUnidades 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2520
      MaxLength       =   80
      TabIndex        =   16
      Text            =   "1"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox TxtUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   480
      MaxLength       =   80
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox TxtDctoSoles 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox TxtDstoporcentaje 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2640
      MaxLength       =   80
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox TxtValoerNeto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3720
      MaxLength       =   80
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox TxtISC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4800
      MaxLength       =   80
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox TxtIGV 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5880
      MaxLength       =   80
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox TxtIvap 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6960
      MaxLength       =   80
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox TxtPrecioVenta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   8040
      MaxLength       =   80
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   2280
      Width           =   975
   End
   Begin MSDataListLib.DataCombo DtcTipoDoc_Ref 
      Height          =   315
      Left            =   405
      TabIndex        =   30
      Top             =   240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape6 
      FillStyle       =   5  'Downward Diagonal
      Height          =   495
      Left            =   240
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COSTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8085
      TabIndex        =   26
      Top             =   960
      Width           =   675
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7275
      TabIndex        =   25
      Top             =   960
      Width           =   435
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3375
      TabIndex        =   24
      Top             =   960
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1740
      TabIndex        =   23
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   660
      TabIndex        =   22
      Top             =   960
      Width           =   765
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   855
      Left            =   240
      Top             =   840
      Width           =   9135
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Unitario"
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
      Left            =   600
      TabIndex        =   15
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dsto (S/.)"
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
      Left            =   1605
      TabIndex        =   14
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dsto (%)"
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
      Left            =   2715
      TabIndex        =   13
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.Neto Unit"
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
      Left            =   3795
      TabIndex        =   12
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISC"
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
      Left            =   5115
      TabIndex        =   11
      Top             =   1920
      Width           =   315
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV"
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
      Left            =   6195
      TabIndex        =   10
      Top             =   1920
      Width           =   315
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IVAP"
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
      Left            =   7215
      TabIndex        =   9
      Top             =   1920
      Width           =   435
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Venta"
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
      Left            =   8085
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   975
      Left            =   240
      Top             =   1800
      Width           =   9135
   End
End
Attribute VB_Name = "FrmModificarCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codigo_usua As String
Public codigoP As String
Dim codigo_A As String
Public Procedencia As EnumProcede
Private Sub CmdActualizar_Click()
Dim ncantidad As Single
ncantidad = 0
If Trim(codigo_usua) = Trim(KEY_USUARIO) Then

    
        strCadena = "SELECT * FROM DocumentoCompra WHERE cDocumentoCompra='" & Trim(FrmCompras.TxtNumeroDoc.text) & "' AND alm_cod='" & Trim(FrmCompras.DtcAlmacen.BoundText) & "' AND sSerie='" & Trim(FrmCompras.TxtSerie.text) & "'  AND doc_cod='" & Trim(FrmCompras.DtcTipoDoc.BoundText) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       strCadena = "UPDATE Docreferencia_Compra SET doc_cod='" & Trim(Me.DtcTipoDoc_Ref.BoundText) & "',sSerie='" & Trim(Me.TxtSerie_Ref.text) & "' " & _
    ",cDocumentoCompra='" & Trim(Me.TxtNumero_Ref.text) & "',des_doc='" & Trim(Me.DtcTipoDoc_Ref.text) & "' WHERE IdReferencia='" & rst("IdReferencia") & "' AND Alm_Cod='" & Trim(rst("Alm_cod")) & "'"
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
    End If
    Set rst = Nothing
    
    strCadena = "SELECT  * FROM  Detalle_DocumentoCompra  WHERE cDocumentoCompra='" & Trim(FrmCompras.TxtNumeroDoc.text) & "' " & _
    " AND alm_cod='" & Trim(FrmCompras.DtcAlmacen.BoundText) & "' AND sSerie='" & Trim(FrmCompras.TxtSerie.text) & "' " & _
    " AND doc_cod='" & Trim(FrmCompras.DtcTipoDoc.BoundText) & "' AND cProducto='" & Trim(FrmCompras.HfdDetalle.TextMatrix(FrmCompras.HfdDetalle.Row, 0)) & "'"
    Call ConfiguraRst(strCadena)
    ncantidad = Val(Me.TxtCantidad.text) * Val(Me.TxtUnidades.text)

    If rst.RecordCount > 0 Then
    strCadena = "UPDATE Detalle_DocumentoCompra SET cantidad='" & ncantidad & "',precio_venta='" & Val(Me.TxtPrecioVenta.text) & "',cProducto='" & Trim(codigoP) & "' WHERE cDocumentoCompra='" & Trim(FrmCompras.TxtNumeroDoc.text) & "' " & _
    " AND alm_cod='" & Trim(FrmCompras.DtcAlmacen.BoundText) & "' AND sSerie='" & Trim(FrmCompras.TxtSerie.text) & "' " & _
    " AND doc_cod='" & Trim(FrmCompras.DtcTipoDoc.BoundText) & "' AND cProducto='" & Trim(FrmCompras.HfdDetalle.TextMatrix(FrmCompras.HfdDetalle.Row, 0)) & "'"
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
    
    strCadena = "UPDATE Kardex SET Mov_cant='" & ncantidad & "',Ing_cant='" & ncantidad & "',Stk_Cant='" & ncantidad & "',cProducto='" & Trim(codigoP) & "' " & _
    " WHERE cTipoMovimiento='I01' AND NumeroDoc='" & Trim(FrmCompras.TxtNumeroDoc.text) & "' " & _
    " AND Alm_cod='" & Trim(FrmCompras.DtcAlmacen.BoundText) & "' AND sSerie='" & Trim(FrmCompras.TxtSerie.text) & "' " & _
    " AND doc_cod='" & Trim(FrmCompras.DtcTipoDoc.BoundText) & "' AND cProducto='" & Trim(FrmCompras.HfdDetalle.TextMatrix(FrmCompras.HfdDetalle.Row, 0)) & "'"
    Set rst = Nothing
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
    
    strCadena = "SELECT Stock FROM Almacen_Productos WHERE cProducto='" & Trim(FrmCompras.HfdDetalle.TextMatrix(FrmCompras.HfdDetalle.Row, 0)) & "' AND Alm_cod='" & Trim(FrmCompras.DtcAlmacen.BoundText) & "'"
    Call ConfiguraRst(strCadena)
    ncantidad = ncantidad + rst(0)
    strCadena = "UPDATE Almacen_Productos SET Stock='" & ncantidad & "' WHERE cProducto='" & Trim(FrmCompras.HfdDetalle.TextMatrix(FrmCompras.HfdDetalle.Row, 0)) & "' AND Alm_cod='" & Trim(FrmCompras.DtcAlmacen.BoundText) & "'"
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
    Unload Me
Else
       MsgBox "Usted no puede Modificar este Comprobante" + Chr(13) + "Consulte con la Persona que Ingreso dicho Comprobante" + Chr(13) + "  -----------   Gracias  ----------- ", vbInformation, "Mensaje para el Usuario"
       Exit Sub
    End If
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
If FrmCompras.Procedencia = Selecionar Then
codigo_usua = ""
strCadena = "SELECT  * FROM  Detalle_DocumentoCompra  WHERE cDocumentoCompra='" & Trim(FrmCompras.TxtNumeroDoc.text) & "' " & _
" AND alm_cod='" & Trim(FrmCompras.DtcAlmacen.BoundText) & "' AND sSerie='" & Trim(FrmCompras.TxtSerie.text) & "' " & _
" AND doc_cod='" & Trim(FrmCompras.DtcTipoDoc.BoundText) & "' AND cProducto='" & Trim(FrmCompras.HfdDetalle.TextMatrix(FrmCompras.HfdDetalle.Row, 0)) & "'"

Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.TxtCodProducto.text = rst("cProducto")
    codigo_A = ""
    codigo_A = rst("cProducto")
    codigoP = Me.TxtCodProducto.text
    Me.TxtCantidad.text = rst("cantidad")
    Me.TxtUnitario.text = rst("c_unitario")
    Me.txtIgv.text = rst("igv")

    Me.TxtDctoSoles.text = rst("Dsto_soles")
    Me.TxtDstoporcentaje.text = rst("Dsto_procentaje")
    Me.txtisc.text = rst("isc")
    Me.TxtIvap.text = rst("ivap")
    Me.TxtPrecioVenta.text = rst("precio_venta")
    codigo_usua = rst("id_usuario")
    Set rst = Nothing
    strCadena = "SELECT     Producto.DescripcionProducto, Producto.PrecioCompra, Unidad.sAbreviatura " & _
    " FROM  Producto INNER JOIN Unidad ON Producto.cUnidad = Unidad.cUnidad WHERE cProducto='" & Trim(Me.TxtCodProducto.text) & "'"
    Call ConfiguraRst(strCadena)
        Me.TxtDescripcionProducto.text = rst("DescripcionProducto")
        Me.TxtUnidad.text = rst("sAbreviatura")
        Me.txtCosto.text = rst("PrecioCompra")
        FrmCompras.Procedencia = Neutro
End If
 End If
End Sub

Private Sub TxtCodProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
     If Trim(Me.TxtCodProducto.text) = "" Then
     Procedencia = Selecionar
        FrmProducto.Show
     Exit Sub
     End If
     strCadena = "SELECT Producto_barras.cProducto, Producto.DescripcionProducto, Producto.PrecioCompra,Producto.PrecioVenta, Unidad.sAbreviatura " & _
     "FROM Producto INNER JOIN Producto_barras ON Producto.cProducto = Producto_barras.cProducto INNER JOIN Unidad ON " & _
     "Producto.cUnidad = Unidad.cUnidad WHERE Producto_barras.cod_barra='" & Trim(Me.TxtCodProducto.text) & "'"
    
       Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        codigoP = rst(0)
        Me.TxtCodProducto.text = Trim(Me.TxtCodProducto.text)
        Me.TxtDescripcionProducto.text = Trim(rst(1))
        Me.TxtUnidad.text = Trim(rst(4))
        Set rst = Nothing
        
    Else
        Procedencia = Selecionar
        FrmProducto.Show
    End If
End If
End Sub

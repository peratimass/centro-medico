VERSION 5.00
Begin VB.Form FrmBonificaciones 
   BorderStyle     =   0  'None
   Caption         =   "BONIFICACIONES"
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn Command1 
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   3960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "PROMEDIO MOVIL"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "FrmBonificaciones.frx":0000
      PICN            =   "FrmBonificaciones.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox TxtBonificacion 
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
      Left            =   2880
      TabIndex        =   14
      Top             =   2115
      Width           =   1215
   End
   Begin VB.TextBox TxtPrecioCompra_compra 
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
      Left            =   2880
      TabIndex        =   13
      Top             =   1545
      Width           =   1215
   End
   Begin VB.TextBox TxtStockCompra 
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
      Left            =   2880
      TabIndex        =   12
      Top             =   1245
      Width           =   1215
   End
   Begin VB.TextBox TxtPreciocompra_anterior 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   11
      Top             =   765
      Width           =   1215
   End
   Begin VB.TextBox TxtStock_anterior 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      TabIndex        =   10
      Top             =   465
      Width           =   1215
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
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
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
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VitekeySoft.ChameleonBtn cmdcerrar 
      Height          =   255
      Left            =   4680
      TabIndex        =   17
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmBonificaciones.frx":2601
      PICN            =   "FrmBonificaciones.frx":261D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PROMEDIO MOVIL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   1905
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "UTILIDAD :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   15
      Top             =   3105
      Width           =   2385
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO VENTA ACTUALIZADO:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   3525
      Width           =   2385
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO COSTO ACTUALIZADO:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   2385
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   1215
      Left            =   240
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRECIO COSTO :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   1905
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "BONIFICACION :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   1905
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "STOCK COMPRA:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1905
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRECIO COSTO     :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1905
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "STOCK ANTERIOR :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1905
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   750
      Left            =   240
      Top             =   360
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   705
      Left            =   240
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   420
      Left            =   240
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   4575
      Left            =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "FrmBonificaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcerrar_Click()
Unload Me
End Sub

Private Sub Command1_Click()

FrmCompras.TxtCostoHoy.Text = Format(Val(Me.TxtPrecioCompra.Text), "#,##0.00")
FrmCompras.txtUtilidadhoy.Text = Format(Val(Me.TxtUtilidad.Text), "#,##0.00")
FrmCompras.TxtventaHoy.Text = Format(Val(Me.TxtPrecioVenta.Text), "#,##0.00")
Unload Me
FrmCompras.CmdAgregar.Enabled = True
FrmCompras.CmdAgregar.SetFocus

End Sub

Private Sub Form_Load()
Dim cproducto As String, promedio As Single
CenterForm Me

cproducto = FrmCompras.TxtCodProducto.Text
strCadena = "SELECT * FROM producto P,almacen_producto A WHERE P.id_producto=A.id_producto AND A.id_alm='" & KEY_ALM & "' AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND P.id_producto='" & cproducto & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.TxtStock_anterior.Text = Format(rst("stock"), "#,##0.00")
    Me.TxtPreciocompra_anterior.Text = Format(rst("precio_compra"), "#,##0.00")
    Me.TxtStockCompra.Text = Format(Val(FrmCompras.TxtCantidad.Text) * Val(FrmCompras.TxtUnidades.Text), "#,##0.00")
    
    
    
    Me.TxtPrecioCompra_compra.Text = Format(Val(FrmCompras.TxtCostoHoy.Text), "#,##0.00")
    If Val(FrmCompras.TxtUnitario.Text) > 0 Then
        Me.TxtBonificacion.Text = Format("0", "#,##0.00")
    Else
        Me.TxtBonificacion.Text = Format(Val(FrmCompras.TxtCantidad.Text) * Val(FrmCompras.TxtUnidades.Text), "#,##0.00")
    End If
    
    strCadena = "SELECT * FROM movimiento_compra_temporal WHERE id_producto='" & cproducto & "' AND dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
         If Val(FrmCompras.TxtUnitario.Text) > 0 Then
            Me.TxtPrecioCompra_compra.Text = Format(rstT("p_costo"), "#,##0.00")
         Else
            Me.TxtPrecioCompra_compra.Text = Format(0, "#,##0.00")
         End If
         Me.TxtStockCompra.Text = Val(Me.TxtStockCompra.Text) ' rstT("cantidad")
    End If
    If (Val(Me.TxtStock_anterior.Text) + Val(Me.TxtStockCompra.Text)) <> 0 Then
    promedio = (Val(Me.TxtStock_anterior.Text) * Val(Me.TxtPreciocompra_anterior.Text) + Val(Me.TxtStockCompra.Text) * Val(Me.TxtPrecioCompra_compra.Text)) / (Val(Me.TxtStock_anterior.Text) + Val(Me.TxtStockCompra.Text))
    Else
    promedio = (Val(Me.TxtStock_anterior.Text) * Val(Me.TxtPreciocompra_anterior.Text) + Val(Me.TxtStockCompra.Text) * Val(Me.TxtPrecioCompra_compra.Text))
    End If
    Me.TxtPrecioCompra.Text = Format(promedio, "#,##0.00")
    If Val(Me.TxtPrecioCompra.Text) = 0 Then
         Me.TxtPrecioCompra.Text = Format(rstT("p_costo"), "#,##0.00")
    End If
    Me.TxtUtilidad.Text = Format(15, "#,##0.00")
    Me.TxtPrecioVenta.Text = Format(Val(Me.TxtPrecioCompra.Text) + Val(Me.TxtPrecioCompra.Text) * Val(Me.TxtUtilidad.Text) / 100, "#,##0.00")
End If

End Sub

Private Sub TxtPrecioVenta_KeyPress(KeyAscii As Integer)
Dim costo As Single
Dim utilidad As Single
If KeyAscii = 13 Then
costo = Val(Me.TxtPrecioCompra.Text)
venta = Val(Me.TxtPrecioVenta.Text)
If costo > 0 Then
utilidad = (venta - costo) * 100 / costo
Me.TxtUtilidad.Text = Format(utilidad, "###0.00")


End If
Me.Command1.SetFocus
End If
End Sub

Private Sub TxtUtilidad_Change()
Me.TxtPrecioVenta.Text = Format(Val(Me.TxtPrecioCompra.Text) + Val(Me.TxtPrecioCompra.Text) * Val(Me.TxtUtilidad.Text) / 100, "#,##0.00")
End Sub

Private Sub TxtUtilidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtUtilidad.Text = Format(Val(Me.TxtUtilidad.Text), "#,##0.00")
    Call Resalta(Me.TxtPrecioVenta)
End If
End Sub

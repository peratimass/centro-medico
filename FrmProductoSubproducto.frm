VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmProductoSubproducto 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtProducto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   240
      Width           =   5655
   End
   Begin VB.TextBox txtCodSubProducto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   240
      MaxLength       =   80
      TabIndex        =   7
      Top             =   4670
      Width           =   1335
   End
   Begin VB.TextBox txtDescripcionSubproducto 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1590
      MaxLength       =   80
      TabIndex        =   6
      Top             =   4670
      Width           =   4335
   End
   Begin VB.CommandButton CmdAgregar 
      BackColor       =   &H80000010&
      Caption         =   "+"
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
      Left            =   6840
      TabIndex        =   5
      Top             =   4665
      Width           =   375
   End
   Begin VB.CommandButton CmdQuitar 
      BackColor       =   &H80000010&
      Caption         =   "-"
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
      Left            =   7200
      TabIndex        =   4
      Top             =   4665
      Width           =   375
   End
   Begin VB.TextBox txtcodigo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H80000010&
      Caption         =   "Salir"
      Height          =   360
      Left            =   8280
      TabIndex        =   2
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox txtUnidad 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtCantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   6000
      MaxLength       =   80
      TabIndex        =   0
      Text            =   "1.00"
      Top             =   4680
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   3255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5741
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   8388608
      BackColorBkg    =   -2147483634
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   120
      Top             =   120
      Width           =   9135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00DFDFE0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   120
      Top             =   4560
      Width           =   9135
   End
   Begin VB.Label LblCantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7725
      TabIndex        =   10
      Top             =   4665
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   5415
      Left            =   0
      Top             =   0
      Width           =   9375
   End
End
Attribute VB_Name = "FrmProductoSubproducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Procedencia As EnumProcede

Private Sub CmdAgregar_Click()
strCadena = "SELECT * FROM producto WHERE id_producto='" & Trim(Me.txtCodSubProducto.Text) & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Call AgregarGrilla
Else
    MsgBox "No se puede Asignar a un Producto que tiene Sub-Productos", vbInformation, "Mensaje para el Administrador"
    Exit Sub
End If
End Sub
Private Sub AgregarGrilla()
Dim strcodigo As String
If Trim(Me.txtCodSubProducto.Text) <> "" And Trim(Me.txtDescripcionSubproducto.Text) <> "" Then
        
        
    strCadena = "INSERT INTO producto_sub(id_producto,id_producto_padre,cantidad,id_alm,ruc) VALUES ('" & Trim(Me.txtCodSubProducto.Text) & "','" & Trim(Me.txtcodigo.Text) & "','" & Val(Me.txtcantidad.Text) & "','" & KEY_ALM & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
     
    Call LLENA
    Me.txtCodSubProducto.Text = "00000"
    Me.txtDescripcionSubproducto.Text = ""
    Call Resalta(Me.txtCodSubProducto)
    
Else
    MsgBox " Ingrese un Produto de la Lista", vbInformation
End If
End Sub

Private Sub CmdQuitar_Click()
strCadena = "DELETE FROM producto_sub WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' AND id_producto_padre='" & Trim(Me.txtcodigo.Text) & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
CnBd.Execute (strCadena)
 
 
Call LLENA
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
CenterForm Me
Me.Top = 500

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = 27 Then
Unload Me
  End If

End Sub

Private Sub Form_Load()
CenterForm Me
  Select Case FrmDetalleProducto.Procedencia
    Case relacionar
      Call LLENA
      
  End Select
End Sub

Private Sub LLENA()
  Me.txtcodigo.Text = Trim(FrmDetalleProducto.LblCodigoProducto.Caption)
  Me.txtProducto.Text = Trim(FrmDetalleProducto.txtdescripcion.Text)
  Me.txtunidad.Text = Trim(FrmDetalleProducto.DtcUnidad.Text)
  strCadena = "SELECT S.id_producto,P.nombre_prod,U.abreviatura,S.cantidad FROM producto_sub S,producto P,unidad U WHERE S.id_producto=P.id_producto AND P.id_unidad=U.id_und AND S.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND S.id_producto_padre='" & Trim(Me.txtcodigo.Text) & "'"
  Call llenarGrid_prod(Me.HfdGrilla, Me)
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
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 5200
           Grilla.ColWidth(2) = 700
           Grilla.ColWidth(3) = 700
           
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


Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdagregar.SetFocus
End If
End Sub

Private Sub txtCodSubProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtCodSubProducto.Text = FormatosCeros(Me.txtCodSubProducto.Text, 5)
    strCadena = "SELECT id_producto,nombre_prod FROM producto WHERE id_producto='" & Trim(Me.txtCodSubProducto.Text) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.txtDescripcionSubproducto.Text = rst("nombre_prod")
        Me.cmdagregar.SetFocus
        Set rst = Nothing
    Else
         Procedencia = relacionar
        FrmProducto.Show
    End If
End If
End Sub
Private Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub



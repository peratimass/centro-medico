VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmProductosRelacionados 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn cmdCargarCombo 
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "CAGAR COMBOS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSubProductos.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chk_importar 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "IMPORTAR"
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
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   5760
      Width           =   1455
   End
   Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
      Height          =   375
      Left            =   10740
      TabIndex        =   10
      Top             =   50
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSubProductos.frx":001C
      PICN            =   "FrmSubProductos.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtCantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   315
      Left            =   8040
      MaxLength       =   80
      TabIndex        =   9
      Text            =   "1.00"
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox txtUnidad 
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
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtcodigo 
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   1215
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
      Left            =   9240
      TabIndex        =   5
      Top             =   4905
      Width           =   375
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
      Left            =   8880
      TabIndex        =   2
      Top             =   4905
      Width           =   375
   End
   Begin VB.TextBox txtDescripcionSubproducto 
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1950
      MaxLength       =   80
      TabIndex        =   1
      Top             =   4905
      Width           =   5895
   End
   Begin VB.TextBox txtCodSubProducto 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   480
      MaxLength       =   80
      TabIndex        =   0
      Top             =   4905
      Width           =   1335
   End
   Begin VB.TextBox TxtProducto 
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   7215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   3735
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6588
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   12632256
      ForeColorFixed  =   8388608
      BackColorBkg    =   -2147483634
      GridColor       =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VitekeySoft.ChameleonBtn cmdProcesarCombo 
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "PROCESAR INFORMACION"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSubProductos.frx":2EEC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Left            =   9765
      TabIndex        =   6
      Top             =   4905
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H8000000A&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   360
      Top             =   4800
      Width           =   10215
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   360
      Top             =   360
      Width           =   10215
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   0
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "FrmProductosRelacionados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Procedencia As EnumProcede

Private Sub ChameleonBtn1_Click()
Me.Visible = False

Call enabled_form(FrmDetalleProducto)
'Call enabled_form(FrmProducto)
Exit Sub
End Sub

Private Sub chk_importar_Click()
If Me.chk_importar.Value = 1 Then
    Me.cmdCargarCombo.Visible = True
End If
End Sub

Private Sub cmdagregar_Click()
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
        
    strCadena = "SELECT * FROM producto_combo_detalle WHERE id_productoc='" & Trim(Me.txtcodigo.Text) & "' AND id_producto='" & Trim(Me.txtCodSubProducto.Text) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
registrar:
        strCadena = "INSERT INTO producto_combo_detalle(id_productoc,id_producto,cantidad,ruc) VALUES ('" & Trim(Me.txtcodigo.Text) & "','" & Trim(Me.txtCodSubProducto.Text) & "','" & Val(Me.txtCantidad.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
    Else
        
        If MsgBox("ITEM YA REGISTRADO PARA ESTE COMBO" + Chr(13) + "DESEA AGREGARLO NUEVAMENTE ?", vbInformation + vbYesNo, KEY_EMPRESA) = vbYes Then
            GoTo registrar
        End If
    End If
    Call LLENA
    Me.txtCodSubProducto.Text = "00000"
    Me.txtDescripcionSubproducto.Text = ""
    Call Resalta(Me.txtCodSubProducto)
    
Else
    MsgBox " Ingrese un Produto de la Lista", vbInformation
End If
End Sub

Private Sub cmdCargarCombo_Click()
Dim Archivo As String
Archivo = Trim("Combo" & KEY_RUC) & ".xls"
      'Dim obj As New get_excel
      Set Me.HfdGrilla.DataSource = Leer_Excel(App.Path & "\comparar_percy\" & Archivo, "Sheet1")
      If Me.HfdGrilla.Rows > 0 Then
        Me.cmdProcesarCombo.Visible = True
      End If
      
End Sub

Private Sub cmdProcesarCombo_Click()
Dim in_combo As String
Dim in_producto As String
Dim in_cantidad As Single



If Me.HfdGrilla.Rows > 0 Then
    
   For i = 0 To Me.HfdGrilla.Rows - 3
    If Val(Me.HfdGrilla.TextMatrix(i, 0)) > 0 Then
        If Me.HfdGrilla.TextMatrix(i, 3) = "SI" Then
           in_combo = Format(Me.HfdGrilla.TextMatrix(i, 0), "00000")
           strCadena = "UPDATE producto SET id_combo='si' WHERE id_producto='" & in_combo & "' and ruc='" & KEY_RUC & "' LIMIT 1"
           CnBd.Execute (strCadena)
        Else
            in_producto = Format(Me.HfdGrilla.TextMatrix(i, 0), "00000")
            in_cantidad = Me.HfdGrilla.TextMatrix(i, 2)
            strCadena = "INSERT INTO producto_combo_detalle(id_productoc,id_producto,cantidad,ruc) VALUES ('" & Trim(in_combo) & "','" & Trim(in_producto) & "','" & in_cantidad & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        End If
        
        
        
        
    End If
   Next i
    
End If


End Sub

Private Sub CmdQuitar_Click()
strCadena = "DELETE FROM producto_combo_detalle WHERE id='" & Val(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' AND id_productoc='" & Trim(Me.txtcodigo.Text) & "' AND ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
 
Call LLENA
End Sub

Private Sub cmdSalir_Click()
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
  
End Sub

Public Sub LLENA()
  Me.txtcodigo.Text = Trim(FrmDetalleProducto.LblCodigoProducto.Caption)
  Me.txtproducto.Text = Trim(FrmDetalleProducto.txtDescripcion.Text)
  Me.TxtUnidad.Text = Trim(FrmDetalleProducto.DtcUnidad.Text)
  
  strCadena = "SELECT * FROM view_producto_combo WHERE id_productoc='" & Trim(Me.txtcodigo.Text) & "' and ruc='" & KEY_RUC & "'"
  Call llenarGrid_prod(Me.HfdGrilla, Me)
  
  
End Sub
Public Sub llenarGrid_prod(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
   ' Grilla.Clear
    Exit Sub
End If
  ' Grilla.Clear
   Me.LblCantidad.Caption = str(rst.RecordCount)
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 6000
           Grilla.ColWidth(3) = 1500
           Grilla.ColWidth(4) = 1500
           
         Next
        cabecera = "ID" & vbTab & "CODIGO" & vbTab & "DESCRIPCION ARTICULO" & vbTab & "UND" & vbTab & "CANT"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("descripcion") & vbTab & Format(rst("cantidad"), "#,##0.00000000")
            Grilla.AddItem Fila
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


Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdagregar.SetFocus
End If
End Sub

Private Sub txtCodSubProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
        
        Procedencia = relacionar
        FrmProducto.Show
        Exit Sub
End If
End Sub
Private Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub


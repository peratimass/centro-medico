VERSION 5.00
Begin VB.Form FrmEnvaseCargo 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3480
      TabIndex        =   9
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtCajas 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox TxtComentario 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1440
      Width           =   3735
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox TxtMonto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton OptNo 
      Appearance      =   0  'Flat
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton OptSi 
      Appearance      =   0  'Flat
      Caption         =   "SI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Comentario:"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   3  'Dot
      FillColor       =   &H000000C0&
      Height          =   2535
      Left            =   240
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DEJO ENVASE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1875
   End
End
Attribute VB_Name = "FrmEnvaseCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Dim envase As String
If Me.OptSi.Value = True Then
    
    KEY_DETALLE = Trim(Me.TxtComentario.Text)
    KEY_MONTOENVASE = 0#
    KEY_ENVASE = "si"
Else

    KEY_DETALLE = Trim(Me.TxtComentario.Text)
    KEY_MONTOENVASE = Val(Me.TxtMonto.Text)
    KEY_ENVASE = "no"
End If
Unload Me
Call FrmVentas.Save
End Sub

Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
  
   Set cnbd1 = New ADODB.Connection
   Dim sys_DataBase As String
   Dim sys_DataBase1 As String
   Dim sys_SUser As String
   Dim sys_SPassword As String
   Dim sys_ConString As String
   
   Dim sys_ConString1 As String
   Dim sys_Server As String
   Dim rstT As New ADODB.Recordset
   KEY_RUC = "20493915526"
 ' CnBd.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & StrRuta & "; "
 'cnbd1.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=lhorna_tienda;Data Source=25.21.178.0"
' cnbd1.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Tuco;"

strCadena = "SELECT * FROM entidad_empresa where id_empresa='20493915526'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        strCadena = "INSERT INTO entidad_empresa (cod_unico,)"
    Next i
End If
 'sys_Server = "190.223.204.246"
 'sys_Server = "localhost"
 'sys_DataBase = "bd_vitekey" 'ConfigRead("DataBase")
 'sys_DataBase1 = "factusoft_inventario" 'ConfigRead("DataBase")
 'sys_SUser = "user1" 'DecryptString(ConfigRead("SUser"))
 'sys_SPassword = "1020304050" 'DecryptString(ConfigRead("SPassword"))
 'sys_ConString = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server & ";" & _
            "Database=" & sys_DataBase & ";" & _
            "UID=" & sys_SUser & ";" & _
            "PWD=" & sys_SPassword & ";" & _
            "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
        
    'Set conConnection = New ADODB.Connection
  '  sys_ConString1 = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server & ";" & _
            "Database=" & sys_DataBase1 & ";" & _
            "UID=" & sys_SUser & ";" & _
            "PWD=" & sys_SPassword & ";" & _
            "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
   ' CnBd.ConnectionString = sys_ConString
    'cnbd1.Open
Dim rstL As New ADODB.Recordset

Dim precio_compra As Single

 
Dim tt As Integer

strCadena = "SELECT * FROM Producto "
  rstT.CursorLocation = adUseClient
  rstT.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
  If rstT.RecordCount > 0 Then
     rstT.MoveFirst
     For i = 0 To rstT.RecordCount - 1
        
      strCadena = "DELETE FROM producto_barras WHERE id_producto='" & formato_item(rstT("CodProducto"), 5) & "'" ' AND ruc='" & KEY_RUC & "'"
      CnBd.Execute (strCadena)
       
      rstT.MoveNext
    Next i
  End If
Set rstT = Nothing
strCadena = "SELECT * FROM Precios "
  rstT.CursorLocation = adUseClient
  rstT.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
  If rstT.RecordCount > 0 Then
     rstT.MoveFirst
     For i = 0 To rstT.RecordCount - 1
        
      strCadena = "DELETE FROM producto_barras WHERE id_producto='" & formato_item(rstT("CodProducto"), 5) & "'" ' AND ruc='" & KEY_RUC & "'"
      CnBd.Execute (strCadena)
       
      rstT.MoveNext
    Next i
  End If
  
  Set rstT = Nothing
  strCadena = "SELECT * FROM UnidadVenta "
  rstT.CursorLocation = adUseClient
  rstT.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
  If rstT.RecordCount > 0 Then
     rstT.MoveFirst
     For i = 0 To rstT.RecordCount - 1
        
     
   
       If IsNull(rstT("CodigoBarra")) = False And Trim(rstT("CodigoBarra")) <> "" Then
        strCadena = "SELECT * FROM producto_barras WHERE cod_barra='" & Trim(rstT("CodigoBarra")) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        
        If rst.RecordCount < 1 Then
            strCadena = "INSERT INTO producto_barras(id_producto,cod_barra,ruc)VALUES('" & formato_item(rstT("CodProducto"), 5) & "','" & Trim(rstT("CodigoBarra")) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
        Else
           strCadena = "DELETE FROM producto_barras WHERE id_producto='" & formato_item(rstT("CodProducto"), 5) & "' AND ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
            
            strCadena = "INSERT INTO producto_barras(id_producto,cod_barra,ruc)VALUES('" & formato_item(rstT("CodProducto"), 5) & "','" & Trim(rstT("CodigoBarra")) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
            
        End If
        End If
        rstT.MoveNext
  Next
  End If
   MsgBox "Proceso Completo" + Space(1) + str(tt)
   Exit Sub
  Dim RstAlmProd As New ADODB.Recordset
  Dim rstP As New ADODB.Recordset
  Dim cUnidad As String
  rstP.CursorLocation = adUseClient
  
  Exit Sub
  '
  'ingresar todos los productios
  strCadena = "SELECT * FROM Producto"
  rstT.CursorLocation = adUseClient
  rstT.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
  If rstT.RecordCount > 0 Then
  rstT.MoveFirst
  For i = 0 To rstT.RecordCount - 1
    codproducto = formato_item(Val(rstT("codProducto")), 5)
    strCadena = "SELECT * FROM producto WHERE id_producto='" & codproducto & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
         cUnidad = BDBuscarCampoUsu("unidad", "id_und", "abreviatura", rstT("CodUnidadMedidaVtaMinima"))
          strCadena = "INSERT INTO producto (id_producto, id_unidad, id_linea, id_marca,nombre_prod,id_igv, ruc) VALUES " & _
        "('" & codproducto & "','" & cUnidad & "','00001','00001','" & rstT("Descripcion") & "','no','" & KEY_RUC & "')"
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
           For j = 0 To RstAlmProd.RecordCount - 1
             strCadena = "INSERT INTO almacen_producto(id_alm,id_producto,ruc) VALUES ('" & RstAlmProd("id_alm") & "','" & Trim(codproducto) & "','" & Trim(KEY_RUC) & "')"
             CnBd.Execute (strCadena)
              
             RstAlmProd.MoveNext
           Next j
           Set RstAlmProd = Nothing
    End If
  rstT.MoveNext
  Next i
  End If
  
  
  Dim precio As Single
  
  'actualizar los precios y las barras
  Set rstT = Nothing
  strCadena = "DELETE  from producto_barras WHERE ruc='" & KEY_RUC & "'"
  CnBd.Execute (strCadena)
   
  strCadena = "SELECT * FROM Precios WHERE CantidadAfectarStock='1' ORDER BY CodProducto ASC"
  rstT.CursorLocation = adUseClient
  rstT.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
  If rstT.RecordCount > 0 Then
  rstT.MoveFirst
  For i = 0 To rstT.RecordCount - 1
    codproducto = formato_item(Val(rstT("codProducto")), 5)
    
    strCadena = "SELECT * FROM producto WHERE id_producto='" & codproducto & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If IsNull(rstT("PrecioPublico")) = True Then
                precio = 0
            Else
            precio = rstT("PrecioPublico")
        End If
        strCadena = "UPDATE producto SET precio_venta='" & Val(precio) & "' WHERE id_producto='" & codproducto & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        If IsNull(rstT("CodigoBarra")) = False Then
        Call agrega_barra(codproducto, rstT("CodigoBarra"))
        End If
    Else
     
     '-----
     'cproducto = codigo
    '   Dim cunidad As String
     '   strCadena = "SELECT * FROM unidad WHERE abreviatura='" & rstT("CodUnidadMedida") & "' AND id_usu='" & KEY_RUC & "'"
      '  Call ConfiguraRst(strCadena)
       ' If rst.RecordCount > 0 Then
        ''    cunidad = rst("id_und")
        'Else
       '     cunidad = "00001"
       ' End If
        
       ' strCadena = "SELECT * FROM linea WHERE abreviatura='" & rstT("CodUnidadMedida") & "' AND id_usu='" & KEY_RUC & "'"
       ' Call ConfiguraRst(strCadena)
       ' If rst.RecordCount > 0 Then
        '    cunidad = rst("id_und")
       ' Else
        '    cunidad = "00001"
        'End If
        
       ' strCadena = "INSERT INTO producto (id_producto, id_unidad, id_linea,precio_venta, id_marca,nombre_prod,id_igv, ruc) VALUES " & _
        "('" & cproducto & "','" & cunidad & "','" & Me.DtcLinea.BoundText & "','" & Val(Me.TxtPrecioVenta.text) & "','" & Me.DtcMarca.BoundText & "'," & _
           "'" & Me.TxtDescripcion.text & "','no','" & KEY_RUC & "')"
        '   CnBd.Execute (strCadena)
          ' strCadena = "SELECT * FROM almacen WHERE ruc='" & KEY_RUC & "'"
         '  RstAlmProd.CursorLocation = adUseClient
           'RstAlmProd.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
           'If RstAlmProd.RecordCount <= 0 Then
            '    MsgBox "No hay Ningun Almacen registrado", vbInformation
             '   MsgBox "Producto NO Grabado, Cree un Almacen", vbInformation
              '  Exit Sub
           'End If
           'RstAlmProd.MoveFirst
           'For i = 0 To RstAlmProd.RecordCount - 1
           '  strCadena = "INSERT INTO almacen_producto(id_alm,id_producto,ruc) VALUES ('" & RstAlmProd("id_alm") & "','" & Trim(cproducto) & "','" & Trim(KEY_RUC) & "')"
            ' CnBd.Execute (strCadena)
            ' RstAlmProd.MoveNext
           'Next i
           'Set RstAlmProd = Nothing
           'Call agrega_barra(cproducto, barra)
     '-----
     
    End If
    rstT.MoveNext
  Next i
  End If
    
End Sub
Private Sub configuraTT(ByVal strCadena As String)
  Set rstTT = New ADODB.Recordset
  rstTT.CursorLocation = adUseClient
  rstTT.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
  rstTT.ActiveConnection = Nothing
End Sub
Private Sub agrega_barra(ByVal cproducto As String, ByVal barra As String)
If Trim(barra) <> "" Then
    strCadena = "SELECT * FROM producto_barras WHERE cod_barra='" & Trim(barra) & "' AND id_producto='" & Trim(cproducto) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        MsgBox "Codigo de barras ya registrado", vbInformation, KEY_EMPRESA
    Else
        strCadena = "INSERT INTO producto_barras VALUES('" & Trim(cproducto) & "','" & Trim(barra) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
    End If
        
    
End If
End Sub

Private Sub Command2_Click()
strCadena = "SELECT * FROM producto WHERE ruc='" & KEY_RUC & "'"
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
If Me.OptSi.Value = True Then
    Me.TxtMonto.Visible = False
Else
    Me.TxtMonto.Visible = True
End If
End Sub

Private Sub OptNo_Click()
If Me.OptNo.Value = True Then
    Me.txtCajas.Visible = False
    Me.TxtMonto.Visible = True
    Call Resalta(Me.TxtMonto)
Else
    Me.txtCajas.Visible = False
End If
End Sub
Public Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub

Private Sub OptSi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdAceptar.SetFocus
End If
End Sub

Private Sub TxtComentario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdAceptar.SetFocus
End If
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtComentario)
End If
End Sub

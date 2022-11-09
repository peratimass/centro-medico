Attribute VB_Name = "ModDeclaraciones"
Option Explicit
Public Sub insertar_acciones(ByVal in_cadena As String, Optional in_tabla As String, Optional in_campo_unico As String, Optional in_unico As Double, Optional in_campo_secundario As String, Optional in_secundario As Double)

in_cadena = Replace(in_cadena, "'", "´")
If KEY_CLOUD = "no" Then  ' QUIERE DECIR TRABAJANDO EN LOCAL
  '  KEY_CADENA = "INSERT INTO entidad_acciones(cadena,tabla,fecha,id_unico,id_secundario,in_campo_unico,in_campo_secundario,ruc)VALUES('" & in_cadena & "','" & in_tabla & "',CURDATE(),'" & in_unico & "','" & in_secundario & "','" & in_campo_unico & "','" & in_campo_secundario & "','" & KEY_RUC & "')"
    
Else
 '   KEY_CADENA = "INSERT INTO entidad_acciones(cadena,tabla,fecha,id_unico,id_secundario,in_campo_unico,in_campo_secundario,ruc)VALUES('" & in_cadena & "','" & in_tabla & "',CURDATE(),'" & in_unico & "','" & in_secundario & "','" & in_campo_unico & "','" & in_campo_secundario & "','" & KEY_RUC & "')"
End If
'CnBd.Execute (KEY_CADENA)


End Sub

Public Sub DarFormato(ByVal HfdPrecio As MSHFlexGrid, ByVal longitud As Integer)
Dim Total As Double
If (HfdPrecio.Rows > 0) Then
  HfdPrecio.col = longitud
  HfdPrecio.Row = 0
  Do While Not HfdPrecio.Row = HfdPrecio.Rows - 1
    HfdPrecio.Row = HfdPrecio.Row + 1
    If HfdPrecio.Text <> "" Then
        HfdPrecio.Text = Format(HfdPrecio.Text, "#,##0.000")
    End If
  Loop
  If (HfdPrecio.Rows - 1) = 0 Then
       HfdPrecio.Text = Format(HfdPrecio.Text, "#,##0.000")
End If
  HfdPrecio.Refresh
  End If
  
End Sub
Public Function moneda(ByVal id_moneda As String)
strCadena = "SELECT * FROM moneda WHERE id_moneda='" & id_moneda & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    moneda = rstT("simbolo")
Else
    moneda = "S/."
End If
End Function
Public Sub Resalta(ByVal Texto As TextBox)
On Error GoTo herrorHandler
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
herrorHandler:
Exit Sub
End Sub
Public Sub DarFormato_t(ByVal HfdPrecio As MSHFlexGrid, ByVal longitud As Integer)
Dim tparcial As Double
  HfdPrecio.col = longitud
  HfdPrecio.Row = 0
  tparcial = 0
  Do While Not HfdPrecio.Row = HfdPrecio.Rows - 1
    HfdPrecio.Row = HfdPrecio.Row + 1
    HfdPrecio.Text = Format(HfdPrecio.Text, "#,##0.00")
    tparcial = HfdPrecio.Text + tparcial
    
  Loop
  FrmEstadistica.lblTotalVenta.Caption = Format(tparcial, "#,##0.00")
  HfdPrecio.Refresh
 
End Sub
Public Function cproducto(ByVal cod_barra As String) As String
strCadena = "SELECT cProducto FROM Producto_barras WHERE cod_barra='" & Trim(cod_barra) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   cproducto = rst(0)
End If
Set rst = Nothing
End Function
Public Sub actualizar_stock(ByVal cproducto As String, ByVal alm_cod As String)
On Error GoTo errorheandler
Dim stk As Single
stk = 0
    strCadena = "SELECT sum(Stk_Cant) FROM Kardex WHERE cProducto='" & Trim(cproducto) & "' AND Alm_cod='" & Trim(alm_cod) & "'"
    Call ConfiguraRst(strCadena)
    stk = rst(0)
    Set rst = Nothing
    strCadena = "UPDATE Almacen_Productos SET Stock='" & stk & "' WHERE cProducto='" & Trim(cproducto) & "'AND Alm_cod='" & Trim(alm_cod) & "'"
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
errorheandler:
    
 Exit Sub
End Sub
Public Function nombre_cuenta(ByVal strCod As String, ByVal StrPlan) As String
    Dim rst_c As New ADODB.Recordset
    strCadena = "SELECT * FROM plan_contable_det WHERE pc_codigo='" & Trim(strCod) & "' AND id_plancontable='" & Trim(StrPlan) & "'"
    rst_c.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
    If rst_c.RecordCount > 0 Then
        nombre_cuenta = rst_c("plan_des")
        Set rst_c = Nothing
    Else
    Set rst_c = Nothing
    End If
End Function

Public Sub DarFormato_u(ByVal HfdPrecio As MSHFlexGrid, ByVal longitud As Integer)
Dim uparcial As Double
  HfdPrecio.col = longitud
  HfdPrecio.Row = 0
  uparcial = 0
  Do While Not HfdPrecio.Row = HfdPrecio.Rows - 1
    HfdPrecio.Row = HfdPrecio.Row + 1
    HfdPrecio.Text = Format(HfdPrecio.Text, "#,##0.00")
    uparcial = HfdPrecio.Text + uparcial
    
  Loop
  FrmEstadistica.LblTotalUtilidad.Caption = Format(uparcial, "#,##0.00")
  HfdPrecio.Refresh
 
End Sub
Public Sub DarFormatoColor(ByVal HfdPrecio As MSHFlexGrid, ByVal longitud As Integer)
Dim j As Integer
Dim i As Integer
Dim Total As Double
If (HfdPrecio.Rows > 0) Then
  HfdPrecio.col = longitud
  HfdPrecio.Row = 0
  Do While Not HfdPrecio.Row = HfdPrecio.Rows - 1
    HfdPrecio.Row = HfdPrecio.Row + 1
    If Val(HfdPrecio.TextMatrix(HfdPrecio.Row, 3)) <= 1 Then
        For i = 3 To 3
           HfdPrecio.col = i
           HfdPrecio.CellForeColor = &HFF&
        Next i
    End If
  Loop
  If (HfdPrecio.Rows - 1) = 0 Then
       HfdPrecio.Text = Format(HfdPrecio.Text, "#,##0.00")
End If
  HfdPrecio.Refresh
  End If
  
End Sub


Public Sub des_und(ByVal HfdPrecio As MSHFlexGrid, ByVal longitud As Integer)
  HfdPrecio.col = longitud
  HfdPrecio.Row = 0
  Do While Not HfdPrecio.Row = HfdPrecio.Rows - 1
    
    HfdPrecio.Row = HfdPrecio.Row + 1
    HfdPrecio.Text = Format(HfdPrecio.Text, "#,##0.00")
  Loop
  HfdPrecio.Refresh
'If (x = 10) Then
 '   x = x + 1
'End If
  
End Sub
Public Sub DarFormatoFecha(ByVal HfdPrecio As MSHFlexGrid, ByVal longitud As Integer)
If (HfdPrecio.Rows > 0) Then
  HfdPrecio.col = longitud
  HfdPrecio.Row = 0
  Do While Not HfdPrecio.Row = HfdPrecio.Rows - 1
    HfdPrecio.Row = HfdPrecio.Row + 1
    If (HfdPrecio.Text <> "") Then
        HfdPrecio.Text = Format(HfdPrecio.Text, "dd-mm-YYYY")
    End If
  Loop
  If (HfdPrecio.Rows - 1) = 0 Then
   HfdPrecio.Text = Format(HfdPrecio.Text, "dd-mm-YYYY")
  End If
  End If
  HfdPrecio.Refresh
End Sub
Public Function formato_item(ByVal Campo As String, ByVal longitud As Integer) As String
  Dim X As Integer
  Dim Formato As String
  Formato = ""
  For X = 1 To longitud
    Formato = Formato + "0"
  Next X
  StrNumero = Format(Campo, Formato)
  formato_item = Gencodigo + StrNumero
  
End Function
Public Function cambio(ByVal fecha As Date) As Single
If IsDate(fecha) Then
    strCadena = "SELECT valor FROM tipo_cambio WHERE fecha<='" & Format(fecha, "YYYY-mm-dd") & "' AND id_creador='" & KEY_RUC & "' ORDer BY fecha Desc limit 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        cambio = rst("valor")
    End If
End If
End Function
Public Function cambio_venta(ByVal fecha As Date) As Single
nuevo:
If IsDate(fecha) Then

    strCadena = "SELECT valor_venta FROM tipo_cambio WHERE fecha<='" & Format(fecha, "YYYY-mm-dd") & "' AND id_creador='" & KEY_RUC & "' ORDer BY fecha Desc limit 1"
    Call ConfiguraRstPP(strCadena)
    If rstPP.RecordCount > 0 Then
        cambio_venta = rstPP("valor_venta")
    Else
        Call get_cambio_sbs(fecha)
        GoTo nuevo
    End If
End If
End Function

Public Function cambio_compra(ByVal fecha As Date) As Single
nuevo:
If IsDate(fecha) Then

    strCadena = "SELECT valor_compra FROM tipo_cambio WHERE fecha<='" & Format(fecha, "YYYY-mm-dd") & "' AND id_creador='" & KEY_RUC & "' ORDer BY fecha Desc limit 1"
    Call ConfiguraRstPP(strCadena)
    If rstPP.RecordCount > 0 Then
        cambio_compra = rstPP("valor_compra")
    Else
        Call get_cambio_sbs(fecha)
        GoTo nuevo
    End If
End If
End Function
Public Function GeneraCodigo(ByVal longitud As Integer) As String
Dim X As Integer
Dim Formato As String
  Formato = ""
  For X = 1 To longitud
    Formato = Formato + "0"
  Next X
   
  If (rst.BOF And rst.EOF) Then
    StrNumero = Format(str(Val(Formato) + 1), Formato)
  Else
    StrNumero = Format(Trim(str(Val(Right(rst(0), longitud + 1)) + 1)), Formato)
  End If
  Set rst = Nothing
  GeneraCodigo = Gencodigo + StrNumero
  Gencodigo = ""

End Function
Public Function GeneraCodigos() As String
If (rst.BOF And rst.EOF) Then
    StrNumero = str(1)
  Else
    StrNumero = Trim(str(Val(rst(0))) + 1)
  End If
  Set rst = Nothing
  GeneraCodigos = Gencodigo + StrNumero
  Gencodigo = ""

End Function
Public Function NombrePersona(ByVal dni As String) As String
strCadena = "SELECT nombre_completo FROM persona WHERE dni='" & dni & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    NombrePersona = UCase(rstT("nombre_completo"))
Else
    NombrePersona = "NO ESPECIFICADO"
End If
End Function

Public Sub LlenaDataCombo(ByVal DtcGuion As DataCombo)
  If Not rst.EOF Then
    Set DtcGuion.RowSource = rst
    DtcGuion.ListField = "Descripcion"
    DtcGuion.BoundColumn = "Codigo"
    DtcGuion.BoundText = rst("Codigo")
    DtcGuion.Enabled = True
  Else
      DtcGuion.Text = ""
      DtcGuion.Enabled = False
  End If
  Set rst = Nothing
End Sub

Public Sub LlenaDataComboL(ByVal DtcGuion As DataCombo)
  If Not rstL.EOF Then
    Set DtcGuion.RowSource = rstL
    DtcGuion.ListField = "Descripcion"
    DtcGuion.BoundColumn = "Codigo"
    DtcGuion.BoundText = rstL("Codigo")
    DtcGuion.Enabled = True
  Else
      DtcGuion.Text = ""
      DtcGuion.Enabled = False
  End If
  Set rstL = Nothing
End Sub

Public Sub LlenaDataComboT(ByVal DtcGuion As DataCombo)
  If Not rstT.EOF Then
    Set DtcGuion.RowSource = rstT
    DtcGuion.ListField = "Descripcion"
    DtcGuion.BoundColumn = "Codigo"
    DtcGuion.BoundText = rstT("Codigo")
    DtcGuion.Enabled = True
  Else
      DtcGuion.Text = ""
      DtcGuion.Enabled = False
  End If
  Set rstT = Nothing
End Sub


Public Function numero(ByVal Tabla As String, ByVal LTipo As String) As String
  strCadena = "SELECT sUltimoNumero, nLongitud FROM Numero WHERE " & _
  " sTabla = '" & Tabla & "' AND sTipo = '" & LTipo & "'"
  Call EjecutaRST(strCadena)
  Lnumdocu = CStr(Val(Right(RstEjecuta(0), 9)) + 1)
  If Len(Lnumdocu) <= RstEjecuta(1) Then
      Lnumdocu = LTipo + Format(Trim(Lnumdocu), "000000000")
      numero = Lnumdocu
      strCadena = "UPDATE Numero SET sUltimoNumero = '" & Lnumdocu & "' " & _
      " WHERE sTabla = '" & Tabla & "' AND sTipo = '" & LTipo & "'"
      Call EjecutaRST(strCadena)
      Set RstEjecuta = Nothing
  Else
      MsgBox MSGTAMAÑO, vbInformation + vbOKOnly, MSGVALIDACION
  End If
End Function

Public Function ValidaNumero(ByVal strCadena As String, ByVal KeyAsc As Integer) _
As Integer
  If KeyAsc = 8 Or KeyAsc = 46 Then
    If strCadena = "I" And KeyAsc = 46 Then
      ValidaNumero = 0
      Exit Function
    End If
  Else
    If KeyAsc < Asc("0") Or KeyAsc > Asc("9") Then
      ValidaNumero = 0
      Exit Function
    End If
  End If
  ValidaNumero = KeyAsc
End Function
Public Function get_comprobante_propio(ByVal in_ventanilla As String) As String
If Val(in_ventanilla) > 0 Then
    strCadena = "SELECT comprobantes_propios FROM almacen WHERE id_alm='" & Trim(in_ventanilla) & "' and  ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstP(strCadena)
    If rstP.RecordCount > 0 Then
       get_comprobante_propio = rstP("comprobantes_propios")
    End If
Else
    get_comprobante_propio = "no"
End If
End Function


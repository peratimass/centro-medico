Attribute VB_Name = "modUtiles"


'Public Function MantMultiple( _
 '           Tabla As String, _
  '          IdTabla As String, _
   '         Titulo As String, _
    '        Campo1 As String, _
     '       Optional Campo1Tipo As varTipoDato = vAlfanumerico, _
      '      Optional Campo1Letra As varTipoLetra = vMixto, _
       '     Optional Campo1Long As Integer = 0, _
        '    Optional Campo1Etiqueta As String = vbNullString, _
         '   Optional Campo1Tam As Long = 0, _
          '  Optional Campo2 As String = vbNullString, _
'            Optional Campo2Tipo As varTipoDato = vAlfanumerico, _
 '           Optional Campo2Letra As varTipoLetra = vMixto, _
  '          Optional Campo2Long As Integer = 0, _
   '         Optional Campo2Etiqueta As String = vbNullString, _
    '        Optional Campo2Tam As Long = 0, _
            Optional TextoBuscar As String = vbNullString, _
     ''       Optional CampoFiltro As String = vbNullString, _
            Optional FiltroValor As Integer = 0) As Long
       '
  '  Dim frmMMult As frmMantMultiple
'    Set frmMMult = New frmMantMultiple
 '
  '  frmMMult.Tabla = Tabla
   ' frmMMult.IdTabla = IdTabla
'    frmMMult.Titulo = Titulo
 '   frmMMult.Campo1 = Campo1
  '  frmMMult.Campo1Tipo = Campo1Tipo
'    frmMMult.Campo1Letra = Campo1Letra
'    frmMMult.Campo1Long = Campo1Long
'    frmMMult.Campo1Etiqueta = Campo1Etiqueta
'    frmMMult.Campo1Tam = Campo1Tam
    
 '   frmMMult.Campo2 = Campo2
 '   frmMMult.Campo2Tipo = Campo2Tipo
 '   frmMMult.Campo2Letra = Campo2Letra
  '  frmMMult.Campo2Long = Campo2Long
  '  frmMMult.Campo2Etiqueta = Campo2Etiqueta
  '  frmMMult.Campo2Tam = Campo2Tam
    
   ' frmMMult.BuscaText = TextoBuscar
   ' frmMMult.CampoFiltro = CampoFiltro
   ' frmMMult.FiltroValor = FiltroValor
   ' frmMMult.Show vbModal
   '
   ' MantMultiple = frmMMult.IdRetorna
    
'End Function

'Public Function CargaBuscador( _
 '               ByVal Tabla As String, _
  '              ByVal IdCampo As String, _
   '             ByVal Campo As String, _
    '            ByRef IdRetorna As Long, _
     '           Optional ByVal CtrlBotones As ControlBotones, _
      '          Optional ByVal NomVentana As String = vbNullString, _
       '         Optional TextoBuscar As String = vbNullString, _
        '        Optional ByVal AnchoVentana As Integer, _
         '       Optional ByVal AltoVentana As Integer, _
          '      Optional ByVal ConEnter As Boolean) As ControlOperaciones
           '
   ' Dim frmBusq As frmBuscador
   ' Set frmBusq = New frmBuscador
    
   ' frmBusq.Tabla = Tabla
   ' frmBusq.IdCampo = IdCampo
   ' frmBusq.Campo = Campo
   ' frmBusq.CtrlBotones = CtrlBotones
    'frmBusq.NomVentana = NomVentana
    'frmBusq.BuscaText = TextoBuscar
    'frmBusq.AnchoVentana = AnchoVentana
    'frmBusq.AltoVentana = AltoVentana
    'frmBusq.ConEnter = ConEnter
    'frmBusq.Show vbModal
    'IdRetorna = frmBusq.IdRetorna
    'TextoBuscar = frmBusq.BuscaText
    'CargaBuscador = frmBusq.Accion
    
   ' Set frmBusq = Nothing
    
'End Function

'Public Function CargaBuscadorSP( _
 '               ByVal SpSqlNombre As String, _
  '              ByRef IdRetorna As Long, _
   '             Optional ByVal CtrlBotones As ControlBotones, _
    '            Optional ByVal NomVentana As String = vbNullString, _
     '           Optional TextoBuscar As String = vbNullString, _
      '          Optional ByVal SpPar1 As Variant = vbNullString, _
       '         Optional ByVal SpPar2 As Variant = vbNullString, _
        '        Optional ByVal SpPar3 As Variant = vbNullString, _
         '       Optional ByVal AnchoVentana As Integer, _
          '      Optional ByVal AltoVentana As Integer, _
           '     Optional ByVal ConEnter As Boolean) As ControlOperaciones
                
   ' Dim frmBusq As frmBuscador
   ' Set frmBusq = New frmBuscador

    'frmBusq.SpSqlNombre = SpSqlNombre
    'frmBusq.SpPar1 = SpPar1
    'frmBusq.SpPar2 = SpPar2
    'frmBusq.SpPar3 = SpPar3
    'frmBusq.CtrlBotones = CtrlBotones
    'frmBusq.NomVentana = NomVentana
    'frmBusq.BuscaText = TextoBuscar
    'frmBusq.AnchoVentana = AnchoVentana
    'frmBusq.AltoVentana = AltoVentana
    'frmBusq.ConEnter = ConEnter
    'frmBusq.Show vbModal
    'IdRetorna = frmBusq.IdRetorna
    'TextoBuscar = frmBusq.BuscaText
    'CargaBuscadorSP = frmBusq.Accion
    
    'Set frmBusq = Nothing
    
'End Function


Public Function SearchIndex(cbo As Control, intKey) As Integer
    Dim i As Long
    
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = intKey Then Exit For
    Next i
    If i = cbo.ListCount Then
        SearchIndex = -1
    Else
        SearchIndex = i
    End If
End Function

Public Function SerchIndexLvw(Lst As ListView, intKey, Optional col As Integer = 0) As Integer
    Dim i As Long
    
    For i = 1 To Lst.ListItems.Count
        Select Case col
            Case 0
                If Lst.ListItems(i).Text = intKey Then Exit For
            Case Is > 0
                If Lst.ListItems(i).SubItems(col) = intKey Then Exit For
        End Select
    Next i
    If i > Lst.ListItems.Count Then
        SerchIndexLvw = -1
    Else
        SerchIndexLvw = i
    End If
End Function

Public Sub addELista(Lst As ListBox, cod As Long, Item1 As String)
    If cod > 0 And Item1 <> vbNullString Then
        Lst.AddItem Item1
        Lst.ItemData(Lst.NewIndex) = cod
    Else
         MsgBox "Falta dato", vbExclamation, "Alerta"
    End If
End Sub

Public Sub SubELista(Lst As ListBox, Idx As Integer)
    If Idx >= 0 And Idx < Lst.ListCount Then
        Lst.RemoveItem (Idx)
    End If
End Sub

Public Sub addItListaC0(lstvw As ListView, cod As Integer, Item1 As String)
    If cod > 0 And Item1 <> vbNullString Then
        With lstvw.ListItems.Add(, , cod)
            .SubItems(1) = Item1
        End With
    Else
        MsgBox "Falta dato", vbExclamation, "Alerta"
    End If
End Sub

Public Sub addItListaC1(lstvw As ListView, cod As Long, Item1 As String, Item2 As String)
    If cod > 0 And Item1 <> vbNullString And Item2 <> vbNullString Then
        With lstvw.ListItems.Add(, , cod)
            .SubItems(1) = Item1
            .SubItems(2) = Item2
        End With
    Else
        MsgBox "Falta dato", vbExclamation, "Alerta"
    End If
End Sub

Public Sub addItListaC2(lstvw As ListView, Cod1 As Long, Cod2 As Long, Item1 As String, Item2 As String)
    If Cod1 > 0 And Cod2 > 0 And Item1 <> vbNullString And Item2 <> vbNullString Then
        With lstvw.ListItems.Add(, , Cod1)
            .SubItems(1) = Cod2
            .SubItems(2) = Item1
            .SubItems(3) = Item2
        End With
    Else
        MsgBox "Falta dato", vbExclamation, "Alerta"
    End If
End Sub

Public Sub addItListaC3(lstvw As ListView, Cod1 As Long, Cod2 As Long, Item1 As String, Item2 As String, Item3 As String)
    If Cod1 > 0 And Cod2 > 0 And Item1 <> vbNullString And Item2 <> vbNullString And Item3 <> vbNullString Then
        With lstvw.ListItems.Add(, , Cod1)
            .SubItems(1) = Cod2
            .SubItems(2) = Item1
            .SubItems(3) = Item2
            .SubItems(4) = Item3
        End With
    Else
        MsgBox "Falta dato", vbExclamation, "Alerta"
    End If
End Sub
Public Sub SubItLista(lstvw As ListView, Idx As Integer)
    If Idx > 0 And Idx <= lstvw.ListItems.Count Then
        lstvw.ListItems.Remove (Idx)
    End If
End Sub

Public Sub AddRegGradilla(Grid As MSFlexGrid, ColCode As Integer)
    With Grid
        .Row = .Rows - 1
        If .TextMatrix(.Row, ColCode) <> vbNullString Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .RowData(.Row) = ROW_NEW
        End If
        .SetFocus
    End With
End Sub

Public Sub DelRegGradilla(Grid As MSFlexGrid, MsgConfirm As String, colNumItem As Integer, ArrCodes() As Long, Optional ItemCode As Long = -1)
    If Grid.Row = 0 Then Exit Sub
    Dim nResp As Integer
    nResp = MsgBox(MsgConfirm, vbYesNo + vbQuestion + vbDefaultButton2, "Advertencia")
    If nResp <> vbYes Then Exit Sub
    
    Dim i As Integer
    With Grid
        If ItemCode <> -1 Then
            If .RowData(.Row) = ROW_UPDATED Then
                ArrCodes(UBound(ArrCodes)) = Val(.TextMatrix(.Row, ItemCode))
                ReDim Preserve ArrCodes(UBound(ArrCodes) + 1)
            End If
        End If
                
        If .Rows = 2 Then
            .RowData(1) = ROW_NEW
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = vbNullString
            Next i
        Else
            For i = .Row + 1 To .Rows - 1
                If Val(.TextMatrix(i, colNumItem)) > 0 Then
                    .TextMatrix(i, colNumItem) = Val(.TextMatrix(i, colNumItem)) - 1
                End If
            Next i
            .RemoveItem (.Row)
        End If
        .SetFocus
    End With
End Sub

Public Sub LV_ColumnSort(ListViewControl As ListView, _
  Column As ColumnHeader)
 With ListViewControl
  If .SortKey <> Column.Index - 1 Then
   .SortKey = Column.Index - 1
   .SortOrder = lvwAscending
  Else
   If .SortOrder = lvwAscending Then
    .SortOrder = lvwDescending
   Else
    .SortOrder = lvwAscending
   End If
  End If
  .Sorted = -1
 End With
End Sub

Public Sub compras_producto(ByVal in_fecha As Date, ByVal in_producto As String)

strCadena = "SELECT id_compra,id_tipo,fecha_emision,fecha_kardex,id_doc,serie,numero,id_proveedor,id_producto,cantidad,c_unitario,id_alm,dni_save,ruc,id_orden_compra,id_recepcion,id_moneda,tc,incremento_neto,incremento_neto_gasto,valor_venta,igv,total,obsequio,factura_contable FROM vargas_kardex_compra where  id_producto='" & in_producto & "' and  fecha_kardex='" & Format(in_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_doc ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       If rst("id_doc") = "0089" Then
            in_tipo = "02"
       Else
            in_tipo = "02"
       End If
       
       If rst("factura_contable") = "si" Then
          in_tipo = "10"
       End If
       
       If rst("id_moneda") = "00002" Then
             If KEY_CON_IGV = "si" Then
                in_costo = (rst("valor_venta")) * rst("tc") / rst("cantidad") + (rst("incremento_neto") / rst("cantidad")) * rst("tc") + (rst("incremento_neto_gasto") / rst("cantidad")) * rst("tc")
             Else
                If KEY_PAIS = KEY_PERU Then
                    in_costo = rst("total") * rst("tc") / rst("cantidad") + (rst("incremento_neto") / rst("cantidad")) * rst("tc") + (rst("incremento_neto_gasto") / rst("cantidad")) * rst("tc")
                Else
                    in_costo = rst("total") / rst("cantidad") + (rst("incremento_neto") / rst("cantidad")) + (rst("incremento_neto_gasto") / rst("cantidad"))
                End If
             End If
        Else
            
            If KEY_CON_IGV = "si" Then
                in_costo = (rst("valor_venta")) / rst("cantidad") + rst("incremento_neto") / rst("cantidad") + rst("incremento_neto_gasto") / rst("cantidad")
            Else
                If rst("obsequio") = "si" Then
                       in_costo = 0
                Else
                       in_costo = rst("total") / rst("cantidad") + rst("incremento_neto") / rst("cantidad") + rst("incremento_neto_gasto") / rst("cantidad")
                End If
            End If
            
        End If
       
        strCadena = "call put_kardex_stock_vitekey_v1('" & in_tipo & "','" & Format(rst("fecha_kardex"), "YYYY-mm-dd") & "','" & rst("id_compra") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_proveedor") & "','" & rst("id_producto") & "','" & Val(Abs(rst("cantidad"))) & "','" & in_costo & "','" & rst("id_alm") & "','" & rst("dni_save") & "','" & rst("obsequio") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
       
       
       rst.MoveNext
   Next i
End If
End Sub
Public Sub compras_producto_vargas(ByVal in_fecha As Date, ByVal in_producto As String)

Call actualizar_kardex_recepcion(in_fecha, in_producto)


strCadena = "SELECT id_compra,id_tipo,fecha_kardex as fecha_emision,id_doc,serie,numero,id_proveedor,id_producto,cantidad,c_unitario,id_alm,dni_save,ruc,id_orden_compra,id_recepcion,id_moneda,tc,incremento_neto,incremento_neto_gasto,igv FROM vargas_kardex_compra where id_doc='0007' and  id_producto='" & in_producto & "' and  fecha_kardex='" & Format(in_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_kardex ASC,id_doc ASC"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    If rstL("id_moneda") = "00002" Then
            in_costo = rstL("c_unitario") * rstL("tc") + (rstL("incremento_neto")) * rstL("tc") + (rstL("incremento_neto_gasto")) * rstL("tc")
        Else
            in_costo = rstL("c_unitario") + rstL("incremento_neto") + rstL("incremento_neto_gasto")
        End If
    strCadena = "call put_kardex_stock_vitekey('02','" & Format(rstL("fecha_emision"), "YYYY-mm-dd") & "','" & rstL("id_compra") & "','" & rstL("id_doc") & "','" & rstL("serie") & "','" & rstL("numero") & "','" & rstL("id_proveedor") & "','" & rstL("id_producto") & "','" & Val(rstL("cantidad")) & "','" & in_costo & "','" & rstL("id_alm") & "','" & rstL("dni_save") & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
End If


strCadena = "SELECT id_compra,id_tipo,fecha_emision,id_doc,serie,numero,id_proveedor,id_producto,cantidad,c_unitario,id_alm,dni_save,ruc,id_orden_compra,id_recepcion,id_moneda,tc,incremento_neto,incremento_neto_gasto,igv FROM vargas_kardex_compra where id_doc='0089' and  id_producto='" & in_producto & "' and  fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_doc ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       If rst("id_doc") = "0089" Then
            in_tipo = "04"
       Else
            in_tipo = "02"
       End If
        
        If rst("id_moneda") = "00002" Then
            in_costo = rst("c_unitario") * rst("tc") + (rst("incremento_neto")) * rst("tc") + (rst("incremento_neto_gasto")) * rst("tc")
        Else
            in_costo = rst("c_unitario") + rst("incremento_neto") + rst("incremento_neto_gasto")
        End If
        
  
       strCadena = "call put_kardex_stock_vitekey('" & in_tipo & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & rst("id_compra") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_proveedor") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','" & in_costo & "','" & rst("id_alm") & "','" & rst("dni_save") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       
       rst.MoveNext
       
   Next i
End If
End Sub
Public Sub salidas_almacen(ByVal in_fecha As Date, ByVal in_producto As String)
strCadena = "SELECT id_compra,id_tipo,fecha_emision,id_doc,serie,numero,id_proveedor,id_producto,cantidad,c_unitario,id_alm,dni_save,ruc,id_orden_compra,id_recepcion,id_moneda,tc,incremento_neto,incremento_neto_gasto,igv FROM vargas_kardex_compra where id_doc IN('0090','0007') and  id_producto='" & in_producto & "' and  fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_doc ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       
            in_tipo = "02"
      
        
        If rst("id_moneda") = "00002" Then
            in_costo = rst("c_unitario") * rst("tc")
        Else
            in_costo = rst("c_unitario")
        End If
        
  
       strCadena = "call put_kardex_stock_vitekey('" & in_tipo & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & rst("id_compra") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_proveedor") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','" & in_costo & "','" & rst("id_alm") & "','" & rst("dni_save") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       
       rst.MoveNext
       
   Next i
End If

End Sub


Public Sub transferencia_ingreso_producto(ByVal in_fecha As Date, ByVal in_producto As String)
strCadena = "SELECT * FROM view_vargas_transferencia_ingreso where id_producto='" & in_producto & "' and  fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_doc ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       in_tipo = "03"
       
       strCadena = "SELECT * FROM almacen_producto WHERE ruc='" & KEY_RUC & "' and id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "'"
       Call ConfiguraRstL(strCadena)
       If rstL.RecordCount < 1 Then
             strCadena = "INSERT INTO almacen_producto(id_alm,precio_venta,precio_compra,id_producto,ruc) VALUES ('" & rst("id_alm") & "','" & get_precio_producto(rst("id_producto"), "00001") & "','" & get_costo_producto(rst("id_producto")) & "','" & rst("id_producto") & "','" & Trim(KEY_RUC) & "')"
             CnBd.Execute (strCadena)
       End If
       
       
       strCadena = "call put_kardex_stock_vitekey('03','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & rst("id_compra") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_remitente") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','" & get_costo_producto(rst("id_producto")) & "','" & rst("id_alm") & "','" & rst("dni_save") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       
       
       rst.MoveNext
   Next i
End If
End Sub

Public Sub transferencia_salida_producto(ByVal in_fecha As Date, ByVal in_producto As String)
strCadena = "SELECT * FROM view_vargas_transferencia_salida where id_producto='" & in_producto & "' and fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_doc ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       in_tipo = "03"
       
       
       If rst("id_motivo") = 1 Then
          strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "' and diferida='si' "
          Call ConfiguraRstPP(strCadena)
          If rstPP.RecordCount > 0 Then
            strCadena = "call put_kardex_stock_vitekey('03','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & rst("id_transferencia") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_destinatario") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "',0,'" & rst("id_alm") & "','" & rst("dni_save") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
          End If
        Else
            strCadena = "call put_kardex_stock_vitekey('03','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & rst("id_transferencia") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_destinatario") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','0','" & rst("id_alm") & "','" & rst("dni_save") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
       End If
       
      
       
       
       
        
                
       rst.MoveNext
   Next i
End If
End Sub

Public Sub ventas_producto(ByVal in_fecha As Date, ByVal in_producto As String)
Dim in_cantidad As Single





strCadena = "SELECT id_venta,id_tipo,fecha_emision,id_doc,serie,numero,id_cliente,id_producto,cantidad,id_alm,dni_save,ruc,id_orden_compra,id_recepcion,obsequio,agranel,id_unidad FROM vargas_kardex_ventas where diferida='no' and  id_producto='" & in_producto & "' and  fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        in_tipo = "01"
        
        If rst("agranel") = "si" Then
            strCadena = "SELECT ifnull(a.`cantidad`,1) as  in_cantidad FROM producto_unidad a WHERE a.`id_producto`='" & rst("id_producto") & "' and id_unidad='" & rst("id_unidad") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            Call ConfiguraRstAux(strCadena)
            If rstAux.RecordCount > 0 Then
               in_cantidad = rstAux("in_cantidad") * rst("cantidad")
            Else
                in_cantidad = rst("cantidad")
            End If
            
        Else
            in_cantidad = rst("cantidad")
        End If
    
      
       strCadena = "call put_kardex_stock_vitekey_v1('" & in_tipo & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & rst("id_venta") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_cliente") & "','" & rst("id_producto") & "','" & in_cantidad & "','0','" & rst("id_alm") & "','" & rst("dni_save") & "','" & rst("obsequio") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       
       On Error GoTo sit
       If KEY_CONTABILIDAD = "si" Then
        strCadena = "call update_costo_venta_vitekey('" & rst("id_venta") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
       End If
sit:
       rst.MoveNext
   Next i
End If
End Sub
Public Sub update_costo_venta(ByVal in_venta As String, ByVal in_ejercicio As Integer, ByVal in_mes As Integer)
strCadena = "UPDATE     con_asiento a     inner join con_periodo p on a.idperiodo = p.`Id` inner join con_asientomovimiento am on am.idasiento = a.id and am.activo = 1 and am.`IndExtorno` = 0 " & _
" inner join `con_cuentacontable` cc on am.idcuentacontable = cc.`Id` and cc.ejercicio = " & " inner join con_documento d on a.idreferencia = d.id  " & _
" inner join view_costo_venta cv on cv.`id_venta` = d.`IdReferencia` SET  " & _
"am.`DebeMN` = cv.`costo` " & _
" Where     a.`IdEmpresaSis` = '" & KEY_RUC & "' and a.activo = 1 and p.ejercicio = '" & in_ejercicio & "' and p.mes = '" & in_mes & "' " & _
" and a.idtipoasiento = '1CIX000000000137' and left(cc.nrocuenta,1) = '2'  and d.idtipodocumento in ('1CIX007')"

End Sub
Public Function get_diferida_venta(ByVal in_venta As String) As Boolean
strCadena = "SELECT * FROM movimiento_venta WHERE diferida='si' and  id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
    get_diferida_venta = True
Else
    get_diferida_venta = False
End If
End Function


Public Sub notas_producto(ByVal in_fecha As Date, ByVal in_producto As String)
strCadena = "SELECT id_venta,id_tipo,fecha_emision,id_doc,serie,numero,id_cliente,id_producto,cantidad,id_alm,dni_save,ruc,id_orden_compra,id_recepcion,id_comprobante,obsequio FROM vargas_kardex_notas where id_producto='" & in_producto & "' and  fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_doc ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
      in_tipo = "07"
      If get_diferida_venta(rst("id_comprobante")) = False Then
            
            strCadena = "call put_kardex_stock_vitekey_v1('" & in_tipo & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & rst("id_venta") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_cliente") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','0','" & rst("id_alm") & "','" & rst("dni_save") & "','" & rst("obsequio") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
       
            strCadena = "call update_costo_venta_vitekey_nota('" & rst("id_venta") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
       End If
    
      rst.MoveNext
   Next i
End If
End Sub
Public Function get_periodo_anio() As String

strCadena = "SELECT id_periodo FROM college_periodo WHERE id_anio='" & Year(KEY_FECHA) & "' and   ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_periodo_anio = rstL("id_periodo")
End If
End Function
Public Function get_familiar(ByVal in_dni As String, ByVal in_parentesco As String) As String


strCadena1 = "SELECT * FROM view_familiar_parentesco WHERE dni='" & in_dni & "' and id_parentesco='" & in_parentesco & "'"
Call ConfiguraRstT(strCadena1)
If rstT.RecordCount > 0 Then
   get_familiar = ",'" & rstT("dni_familia") & "','" & rstT("a_paterno") & "','" & rstT("a_materno") & "','" & rstT("nombres") & "','" & rstT("nivel") & "','" & rstT("descripcion") & "','" & rstT("direccion") & "','" & rstT("telefono") & "'"
Else
   get_familiar = ",'-','-','-','-','-','-','-','-'"
End If
End Function

Public Function put_cancelar_comprobante(ByVal in_serie As String, ByVal in_numero As String, ByVal in_doc As String, ByVal in_entidad As String, ByVal in_moneda As String, ByVal in_monto As Double, ByVal in_tc As Single, ByVal in_movimiento As String, ByVal in_referencia As String, ByVal emision As String, ByVal vencimiento As String, ByVal in_canje_anticipo As String)

                   
                    Documento = get_documento_abrev(in_doc) & ":" & in_serie & "-" & in_numero
                                        
                    strCadena = "call P_insert_venta_cancelacion_v13('" & in_doc & "','" & KEY_ALM & "','0','" & in_moneda & "','no'," & _
                    "'" & in_serie & "','" & in_numero & "','" & in_entidad & "','" & get_persona(in_entidad) & "','0','0','0','" & in_monto & "','0'," & _
                    "'" & in_monto & "','0','" & Format(emision, "YYYY-mm-dd") & "','" & Format(vencimiento, "YYYY-mm-dd") & "','01','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & Val(in_tc) & "','no','" & formato_item(Month(CVDate(emision)), 2) & "','" & Year(CVDate(emision)) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstlocal(strCadena)
                    id_venta = rstLocal(0)
                    
                    strCadena = "UPDATE movimiento_venta SET id_canje_anticipo='" & in_canje_anticipo & "'  WHERE id_venta='" & in_venta & "'"
                    CnBd.Execute (strCadena)
                                                          
                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & id_venta & "','00','" & in_referencia & "','-','1','" & in_monto & "','0','" & in_monto & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & id_venta & "','" & Val(in_movimiento) & "','" & in_monto & "','" & in_monto & "','" & in_moneda & "','" & in_moneda & "','" & in_tc & "')"
                    CnBd.Execute (strCadena)
                    
        End Function
        
Public Function get_turno(ByVal in_turno As String) As String
strCadena = "SELECT * FROM turno WHERE id_turno='" & in_turno & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   get_turno = rstA("descripcion")
Else
   get_turno = "-"
End If
End Function
        
        
Public Function get_tipo_cobertura(ByVal in_dni As String) As String

strCadena = "SELECT id_tipo_cliente FROM entidad_empresa WHERE cod_unico='" & in_dni & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
Call EjecutaRST(strCadena)
get_tipo_cobertura = RstEjecuta!id_tipo_cliente


End Function

Public Function get_periodo_cierre(ByVal in_periodo As String, in_area As String) As Boolean

On Error GoTo salir
strCadena = "SELECT * FROM view_cierre_periodo WHERE IndCierreContabilidad='1' and  id_periodo='" & in_periodo & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
    get_periodo_cierre = True
Else
    get_periodo_cierre = False
    
    Select Case in_area
        Case "ventas"
             strCadena = "SELECT * FROM view_cierre_periodo WHERE IndCierreVentas='1' and  id_periodo='" & in_periodo & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Case "compras"
             strCadena = "SELECT * FROM view_cierre_periodo WHERE IndCierreCompras='1' and  id_periodo='" & in_periodo & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Case "caja"
             strCadena = "SELECT * FROM view_cierre_periodo WHERE IndCierreCaja='1' and  id_periodo='" & in_periodo & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Case "almacen"
             strCadena = "SELECT * FROM view_cierre_periodo WHERE IndCierreAlmacen='1' and  id_periodo='" & in_periodo & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    End Select
    
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount > 0 Then
       get_periodo_cierre = True
    Else
       get_periodo_cierre = False
    End If
    
    
End If

Exit Function
salir:




End Function
Public Function get_periodo_cierre_fecha(ByVal in_fecha As Date) As Boolean
strCadena = "SELECT id FROM con_periodo where Ejercicio='" & Year(in_fecha) & "' and Mes='" & Month(in_fecha) & "' LIMit 1"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
   If get_periodo_cierre(rstT("Id"), "almacen") = True Then
      get_periodo_cierre_fecha = True
   Else
      get_periodo_cierre_fecha = False
   End If

End If




End Function

Public Sub quitar_bonificacion_linea(ByVal in_detalletem As Double)
Dim in_dni As String
Dim in_producto As String

 strCadena = "SELECT * FROM temporal_ventas WHERE id='" & in_detalletem & "' and ruc='" & KEY_RUC & "'"
 Call ConfiguraRst(strCadena)
 
 
 
 
 
 If rst.RecordCount > 0 Then
    
    in_dni = rst("id_dni")
    in_producto = rst("id_producto")
     
    strCadena = "DELETE FROM temporal_ventas WHERE id='" & in_detalletem & "' and ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
   
    strCadena = "CALL put_bonificacion_linea_dell('" & in_producto & "','" & in_dni & "','" & KEY_USUARIO & "','" & KEY_ALM & "','0','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
 End If

End Sub


Public Sub actualizar_kardex_movimiento(ByVal id_movimiento As String, ByVal in_tipo As String, ByVal in_progresbar As ProgressBar)
Dim in_fecha_kardex As Date

MsgBox "SE VA A PROCEDER A ACTUALIZAR KARDEX" + Chr(13) + Chr(13) + "PULSE ACEPTAR Y DEJE QUE TERMINE DE ACTUALIZAR.", vbInformation
    
    strCadena = "call ADM_kardex('" & id_movimiento & "','" & in_tipo & "','" & KEY_RUC & "')"
    Call ConfiguraRstIN(strCadena)
      If rstIN.RecordCount > 0 Then
        in_progresbar.Min = 1
        in_progresbar.Max = rstIN.RecordCount + 1
         rstIN.MoveFirst
         
         
         For i = 0 To rstIN.RecordCount - 1
            If KEY_RUC = "20128836251" Then
               Call update_kardex_Vargas_modulo_compra(rstIN("id_producto"), Format(rstIN("fecha_kardex"), "YYYY-mm-dd"))
            Else
               Call update_kardex_update(rstIN("id_producto"), Format(rstIN("fecha_kardex"), "YYYY-mm-dd"))
            End If
            
            
            rstIN.MoveNext
            in_progresbar.Value = i + 1
            DoEvents
         Next i
      
      
      
      End If
      
      MsgBox "Proceso Actualizacion Kardex Correcto.", vbInformation
      
      
      
End Sub

Public Function get_ubigeo_sunat(ByVal in_dni As String, ByVal text_codigo As TextBox, ByVal text_descripcion As TextBox)

strCadena = "SELECT codigo_ubigeo_sunat FROM persona WHERE dni='" & in_dni & "'"
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
    strCadena = "SELECT * FROM view_ubigeo_sunat WHERE cod_ubigeo_sunat='" & rstAux("codigo_ubigeo_sunat") & "'"
    Call ConfiguraRstAux(strCadena)
    If rstAux.RecordCount > 0 Then
        text_codigo = rstAux("cod_ubigeo_sunat")
        text_descripcion = rstAux("ubigeo")
    End If
End If



End Function
Public Function get_ubigeo_sunat_persona(ByVal in_dni As String) As String
strCadena = "SELECT codigo_ubigeo_sunat FROM persona WHERE dni='" & in_dni & "'"
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
   get_ubigeo_sunat_persona = rstAux("codigo_ubigeo_sunat")
End If
End Function

Public Function get_ubigeo_transferencia(ByVal in_transferencia As String, ByVal tipo As String) As String



strCadena = "SELECT ubigeo_origen,ubigeo_destino FROM movimiento_transferencia WHERE id_transferencia='" & in_transferencia & "' and ruc='" & KEY_RUC & "' LIMIT 1 "
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
   If tipo = "origen" Then
        get_ubigeo_transferencia = rstAux("ubigeo_origen")
   Else
        get_ubigeo_transferencia = rstAux("ubigeo_destino")
   End If
   
   
End If
End Function

Public Function get_ubigeo_sunat_v2(ByVal in_codigo As String, ByVal text_codigo As TextBox, ByVal text_descripcion As TextBox)


    strCadena = "SELECT * FROM view_ubigeo_sunat WHERE cod_ubigeo_sunat='" & in_codigo & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        text_codigo.Text = rstT("cod_ubigeo_sunat")
        text_descripcion.Text = rstT("ubigeo")
        get_ubigeo_sunat_v2 = rstT("cod_ubigeo_sunat")
    End If




End Function

Public Function get_ubigeo_sunat_descripcion(ByVal in_codigo As String) As String


    strCadena = "SELECT * FROM view_ubigeo_sunat WHERE cod_ubigeo_sunat='" & in_codigo & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then

       get_ubigeo_sunat_descripcion = rstT("ubigeo")
    End If




End Function

Public Function get_saldo_comprobante(ByVal in_venta As String) As Double
   On Error GoTo salir
    strCadena = "SELECT (total-function_pago_factura(id_venta,'" & KEY_FECHA & "',id_moneda,ruc)) as saldo FROM movimiento_venta WHERE id_venta='" & in_venta & "' LIMIT 1"
    Call ConfiguraRstAux(strCadena)
    get_saldo_comprobante = rstAux("saldo")
salir:
Exit Function
    get_saldo_comprobante = 0
End Function



Public Sub delete_asiento_venta_migracion(ByVal in_venta As String)
   On Error GoTo salir
    
    strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount > 0 Then
        in_doc = ""
        If rstA("id_doc") = "0001" Then
           in_doc = "FACTURA:"
        End If
         If rstA("id_doc") = "0003" Then
           in_doc = "BOLETA:"
        End If
         If rstA("id_doc") = "0007" Then
         in_doc = "NC:"
        End If
        in_glosa = Trim(in_doc & rstA("serie") & rstA("numero") & Space(1) & rstA("ncliente"))
        
        strCadena = "DELETE FROM movimiento_venta WHERE id_venta='" & in_venta & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        
        Call delete_asiento_masivo(rstA("id_venta"), in_glosa)
    End If
    
    
    
    
    
    
    
    
    
salir:
Exit Sub
    
End Sub




Private Sub delete_asiento_masivo(ByVal in_venta As String, ByVal in_numero As String)

strCadena = "SELECT id FROM con_documento WHERE Activo='1' and  idReferencia='" & in_venta & "'  and idEmpresaSis='" & KEY_RUC & "' "
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   rstK.MoveFirst
   For i = 0 To rstK.RecordCount - 1
       strCadena = "DELETE FROM  con_venta  where IdDocumento = '" & rstK("id") & "' AND Activo='1'"
       CnBd.Execute (strCadena)
       
       strCadena = "DELETE FROM  con_documento  where id = '" & rstK("id") & "' AND Activo='1'"
       CnBd.Execute (strCadena)
       
       
       
       strCadena = "SELECT * FROM con_asiento WHERE Glosa LIKE '%" & Trim(in_numero) & "%' and IdTipoAsiento IN('1CIX000000000137','1CIX000000000053','1CIX000000000055') and idEmpresaSis='" & KEY_RUC & "' "
       Call ConfiguraRstL(strCadena)
       If rstL.RecordCount > 0 Then
          rstL.MoveFirst
          For j = 0 To rstL.RecordCount - 1
                strCadena = "DELETE FROM  con_asiento  where id = '" & rstL("id") & "'"
                CnBd.Execute (strCadena)
                
                strCadena = "SELECT * FROM con_asientomovimiento where IdAsiento = '" & rstL("id") & "'"
                Call ConfiguraRstA(strCadena)
                If rstA.RecordCount > 0 Then
                   rstA.MoveFirst
                   For k = 0 To rstA.RecordCount - 1
                       strCadena = "DELETE FROM  con_asientomovimiento  where Id = '" & rstA("id") & "' and idEmpresaSis='" & KEY_RUC & "'"
                       CnBd.Execute (strCadena)
                       
                       strCadena = "DELETE FROM  con_asientomovimiento_documento  where IdAsientoMovimiento = '" & rstA("id") & "' and idEmpresaSis='" & KEY_RUC & "'"
                       CnBd.Execute (strCadena)
                       
                       strCadena = "DELETE FROM  CON_MovimientoCajaBanco  where IdAsientoMovimiento = '" & rstA("id") & "' and idEmpresaSis='" & KEY_RUC & "'"
                       CnBd.Execute (strCadena)
                       rstA.MoveNext
                   Next k
                End If
                
                
                
            rstL.MoveNext
                
          Next j
       End If
       rstK.MoveNext
   Next i
End If


End Sub






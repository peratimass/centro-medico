Attribute VB_Name = "ModKardex"

Public Function CodigoKardex() As String
strCadena = "SELECT cKardex FROM Kardex ORDER BY cKardex DESC"
    Call ConfiguraRst(strCadena)
    CodigoKardex = GeneraCodigos()
    Set rst = Nothing
End Function
Public Sub put_actualizar_kardex_inventario(ByVal in_producto As String, ByVal in_alm As String, ByVal in_stock_real As Double, ByVal in_stock_factura As Double, ByVal in_periodo As String)
Dim strInventario As String
Dim in_cantidad As Double
Dim in_cantidad_contable As Double
Dim in_numero As String
   

strCadena = "SELECT ifnull(sum(cantidad_real),0) FROM kardex  WHERE id_producto='" & in_producto & "' and  ruc='" & KEY_RUC & "'  AND id_alm='" & in_alm & "'"
Call ConfiguraRst(strCadena)
stock_actual = rst(0)


strCadena = "SELECT * FROM almacen_producto  WHERE id_producto='" & in_producto & "' and  ruc='" & KEY_RUC & "'  AND id_alm='" & in_alm & "' LIMIT 1"
Call ConfiguraRst(strCadena)


If rst.RecordCount > 0 Then
    cod_articulo = in_producto
 
    'strInventario = formato_item(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), 6)
    
   ' strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
    "('" & strInventario & "','" & cod_articulo & "','0106','001','" & strInventario & "','" & Val(rst("precio_compra")) & "','" & KEY_FECHA & "','" & in_alm & "','" & in_stock_real & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
    'CnBd.Execute (strCadena)
    in_cta_compra = KEY_CTA_COMPRA_SOLES
    
    If in_stock_real > Val(stock_actual) Then
       strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0089' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             
             strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "' and id_alm='" & in_alm & "' and id_doc='0089' LIMIT 1"
             Call ConfiguraRstlocal(strCadena)
             If rstLocal.RecordCount > 0 Then
               in_serie = rstLocal("serie")
               in_numero = formato_item(rstLocal("numero"), 8)
             Else
                MsgBox "CREE EL COMPROBANTE INGRESO A ALMACEN" + Chr(13) + "PARA ESTA SUCURSAL", vbInformation
                Exit Sub
             End If
                
             
       End If
            in_cantidad = in_stock_real - Val(stock_actual)
           
            strCadena = "call P_insert_compra_ultimate('0089','" & in_alm & "',CURDATE(),CURDATE(),'02'," & _
            "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
            "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
            "'0','0','0','0','0','0','0','0','0','0','0'," & _
            " '" & KEY_USUARIO & "','OBSERVACION','01','" & in_periodo & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
            Call ConfiguraRstP(strCadena)
           
           
           id_compra = rstP(0)
           
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & in_producto & "','" & in_cantidad & "','0'," & _
           "'0','0','0','" & in_cantidad * rst("precio_compra") & "','0','0', " & _
           "'0','0','0','" & in_cantidad * rst("precio_compra") & "','0','" & rst("precio_compra") * in_cantidad & "','" & Val(rst("precio_venta")) & "','" & Val(rst("precio_compra")) & "','" & in_alm & "','" & get_producto(in_producto) & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
           strCadena = "call put_kardex_stock_inventario_v2('04','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0089','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(in_producto) & "','" & in_cantidad & "','" & rst("precio_compra") & "','" & in_alm & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
        
    Else
        strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0090' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
        Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             
             strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "' and id_alm='" & in_alm & "' and id_doc='0090' LIMIT 1"
             Call ConfiguraRstlocal(strCadena)
             If rstLocal.RecordCount > 0 Then
               in_serie = rstLocal("serie")
               in_numero = formato_item(rstLocal("numero"), 8)
             Else
                MsgBox "CREE EL COMPROBANTE SALIDA A ALMACEN" + Chr(13) + "PARA ESTA SUCURSAL", vbInformation
                Exit Sub
             End If
             
        End If
            in_cantidad = Val(stock_actual) - in_stock_real
            If in_cantidad <> 0 Then
                strCadena = "call P_insert_compra_ultimate('0090','" & in_alm & "',CURDATE(),CURDATE(),'02'," & _
                "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
                "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
                "'0','0','0','0','0','0','0','0','0','0','0'," & _
                " '" & KEY_USUARIO & "','OBSERVACION','01','" & in_periodo & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
                Call ConfiguraRstP(strCadena)
                id_compra = rstP(0)
                strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
                "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & in_producto & "','" & in_cantidad & "','0'," & _
                "'0','0','0','" & in_cantidad * Val(rst("precio_compra")) & "','0','0', " & _
                "'0','0','0','" & in_cantidad * Val(rst("precio_compra")) & "','0','" & Val(rst("precio_compra")) * in_cantidad & "','" & rst("precio_venta") & "','" & Val(rst("precio_compra")) & "','" & in_alm & "','" & get_producto(in_producto) & "','0','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
           
                strCadena = "call put_kardex_stock_inventario_v2('01','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0090','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(in_producto) & "','" & in_cantidad & "','" & Val(rst("precio_compra")) & "','" & in_alm & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
            
            
            
        End If
        
        
  End If
   
    
    in_comentario = "INVENTARIO INICIAL:" & KEY_VENDEDOR + Chr(13) + "CONTEO FISICO :" + str(stock_actual) + Chr(13) + "AJUSTE :" + str(in_cantidad)
    strCadena = "UPDATE producto SET  inventario='si',comentario='" & in_comentario & "' WHERE id_producto='" & Trim(in_producto) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    
    

    strCadena = "UPDATE almacen_producto SET stock_factura='" & in_stock_factura & "', precio_venta='" & Val(rst("precio_venta")) & "',precio_compra='" & Val(rst("precio_compra")) & "' WHERE id_producto='" & Trim(in_producto) & "' AND id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    
 
    
    
  If KEY_STOCK_CONTABLE = "no" Then
    GoTo fin
  End If
  
  'If KEY_RUC <> "20495916830" Then
  '   GoTo fin
  'End If
    
    
    
    
strCadena = "SELECT * FROM almacen_producto  WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    cod_articulo = rst("id_producto")
    stock_actual = rst("stock_factura")
    
    If Val(in_stock_factura) <= 0 Then
        Exit Sub
    End If
    
    If Val(stock_actual) = Val(in_stock_factura) Then
        
        Exit Sub
    End If

    
    strInventario = formato_item(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), 6)
    
    strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
    "('" & strInventario & "','" & in_producto & "','0106','001','" & strInventario & "','" & Val(rst("precio_compra")) & "','" & KEY_FECHA & "','" & in_alm & "','" & Val(in_stock_factura) & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    in_cta_compra = KEY_CTA_COMPRA_SOLES
    
    If Val(in_stock_factura) > Val(stock_actual) Then
       strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0089' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             in_numero = Format(1, "00000000")
             in_serie = "0001"
       End If
            in_cantidad = in_stock_factura - Val(stock_actual)
           
            strCadena = "call P_insert_compra_ultimate('0089','" & in_alm & "',CURDATE(),CURDATE(),'02'," & _
            "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
            "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
            "'0','0','0','0','0','0','0','0','0','0','0'," & _
            " '" & KEY_USUARIO & "','OBSERVACION','01','" & in_periodo & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
            Call ConfiguraRstP(strCadena)
            id_compra = rstP(0)
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & in_producto & "','" & in_cantidad & "','0'," & _
           "'0','0','0','" & in_cantidad * Val(rst("precio_compra")) & "','0','0', " & _
           "'0','0','0','" & in_cantidad * Val(rst("precio_compra")) & "','0','" & Val(rst("precio_compra")) * in_cantidad & "','" & Val(rst("precio_venta")) & "','" & Val(rst("precio_compra")) & "','" & in_alm & "','" & get_producto(in_producto) & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
           strCadena = "call put_kardex_stock_inventario_v2('10','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0089','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & in_producto & "','" & in_cantidad & "','" & Val(rst("precio_compra")) & "','" & in_alm & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
        
    Else
        strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0090' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
        Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             in_numero = Format(1, "00000000")
             in_serie = "0001"
       End If
            in_cantidad = in_stock_factura - Val(stock_actual)
            If in_cantidad <> 0 Then
                strCadena = "call P_insert_compra_ultimate('0090','" & in_alm & "',CURDATE(),CURDATE(),'02'," & _
                "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
                "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
                "'0','0','0','0','0','0','0','0','0','0','0'," & _
                " '" & KEY_USUARIO & "','OBSERVACION','01','" & in_periodo & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
                Call ConfiguraRstP(strCadena)
                id_compra = rstP(0)
                strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
                "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(in_producto) & "','" & in_cantidad & "','0'," & _
                "'0','0','0','" & in_cantidad * Val(rst("precio_compra")) & "','0','0', " & _
                "'0','0','0','" & in_cantidad * Val(rst("precio_compra")) & "','0','" & Val(rst("precio_compra")) * in_cantidad & "','" & Val(rst("precio_venta")) & "','" & Val(rst("precio_compra")) & "','" & in_alm & "','" & get_producto(in_producto) & "','0','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
           
                strCadena = "call put_kardex_stock_inventario_v2('10','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0090','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(in_producto) & "','" & in_cantidad & "','" & Val(rst("precio_compra")) & "','" & in_alm & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
            
            
            
        End If
        
        
  End If
   
    
    
    

        strCadena = "UPDATE almacen_producto SET precio_venta='" & Val(rst("precio_venta")) & "',precio_compra='" & Val(rst("precio_compra")) & "' WHERE id_producto='" & Trim(in_producto) & "' AND id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
    
fin:
    
  

End Sub


Public Sub ActualizaKardex(ByVal StrProducto As String, ByVal codalmacen As String)
Dim TotalProd As Double
Dim StockFinal As Double
Dim Anterior As Double
Dim Actual As Double
Dim i As Double
Dim Registros As Integer
Dim Producto As String

Dim Total
'----------------------------

'----------------------------
strCadena = "SELECT SUM(Stk_Cant) FROM kardex WHERE cProducto='" & Trim(StrProducto) & "' AND Alm_Cod='" & codalmacen & "'"
Call ConfiguraRst(strCadena)
TotalProd = rst(0)
Set rst = Nothing

strCadena = "SELECT Stk_gen FROM kardex WHERE cProducto='" & Trim(StrProducto) & "' AND Alm_Cod='" & codalmacen & "' ORDER BY int_Kardex DESC"
Call ConfiguraRst(strCadena)
StockFinal = rst(0)
Set rst = Nothing

If TotalProd = StockFinal Then
    Exit Sub
End If
'If (Rst(0) = "V") Then
 '   StrCadena = "select * from Producto"
'End If
strCadena = "SELECT Stk_cant,Stk_Anterior,Stk_Gen FROM kardex WHERE cProducto='" & Trim(StrProducto) & "' AND Alm_Cod='" & codalmacen & "' ORDER BY FechaProceso ASC"
Call ConfiguraRst(strCadena)
rst.MoveFirst
For i = 0 To rst.RecordCount - 2
        Actual = rst(2)
        rst.MoveNext
        rst(1) = Actual
        rst(2) = rst(0) + Actual
        rst.Update
Next i
Set rst = Nothing

End Sub
Public Sub kardex(ByVal StrCodProducto As String, ByVal TipoDoc As String, _
ByVal Almacen As String, ByVal NumDoc As String, ByVal serie As String, ByVal TipoMov As String, ByVal FechaProceso As Date, _
ByVal FechaEmision As Date, ByVal IntMov As Double, Optional Ing_Cant As Single, _
Optional ByVal Sal_Cant As Single, Optional ByVal Stk_Cant As Double, Optional ByVal precio As Double, _
Optional ByVal Ing_Soles As Double, Optional ByVal Sal_Soles As Double, Optional ByVal Stk_Soles As Double, Optional ByVal strpersona As String, Optional ByVal CodKradex As String, Optional ByVal dfactura As Boolean)

Dim IntAnteriorStock As Single
Dim IntStockActual As Single
Dim cKardex As String * 20
        
    strCadena = "SELECT Stk_Cant FROM kardex WHERE (cProducto='" & StrCodProducto & "' AND Alm_Cod='" & Almacen & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Set rst = Nothing
        strCadena = "SELECT SUM(Stk_Cant)FROM kardex WHERE (cProducto='" & StrCodProducto & "' AND Alm_Cod='" & Almacen & "')"
        Call ConfiguraRst(strCadena)
        IntAnteriorStock = Val(rst(0))
        IntStockActual = Val(rst(0)) + Stk_Cant
        
        strCadena = "UPDATE Almacen_Productos SET Stock='" & IntStockActual & "' WHERE cProducto='" & Trim(StrCodProducto) & "' AND Alm_Cod='" & Trim(Almacen) & "'"
        CnBd.Execute (strCadena)
         
        
    Else
          IntStockActual = Stk_Cant
          IntAnteriorStock = Stk_Cant
          strCadena = "UPDATE Almacen_Productos SET Stock='" & IntStockActual & "' WHERE cProducto='" & Trim(StrCodProducto) & "' AND Alm_Cod='" & Trim(Almacen) & "'"
          Call EjecutaRST(strCadena)
          Set RstEjecuta = Nothing
          Set rst = Nothing
    End If
    
   
    
'*** Registra el movimiento del producto en la tabla Kardex ***

strCadena = "INSERT INTO Kardex(NumeroDoc,cProducto,doc_cod,Alm_cod,Persona,sSerie,cTipoMovimiento,FechaProceso,FechaEmision,Mov_Cant," & _
                             "Ing_Cant,Sal_Cant,Stk_Cant,Stk_Anterior,Stk_Gen,Precio,Ing_Soles,Sal_Soles,Stk_Soles,IdUsuario) VALUES " & _
                             "('" & NumDoc & "','" & StrCodProducto & "','" & TipoDoc & "','" & Almacen & "','" & strpersona & "','" & serie & "','" & TipoMov & "'," & _
                             "'" & FechaProceso & "','" & FechaEmision & "','" & IntMov & "','" & Ing_Cant & "','" & Val(Sal_Cant * -1) & "', " & _
                             "'" & Stk_Cant & "','" & IntAnteriorStock & "','" & IntStockActual & "','" & precio & "','" & Ing_Soles & "','" & Val(Sal_Soles * -1) & "','" & Stk_Soles & "','" & KEY_USUARIO & "')"
                            CnBd.Execute (strCadena)
                             
                             
       
End Sub
Public Sub ActualizaStock_Almacenes(ByVal strCodAlmacen As String, ByVal StrCodProducto As String, ByVal IntCantMov As Double, _
                                    ByVal TipoDoc As String, ByVal doc_ref As String, ByVal serie_ref As String, ByVal num_ref As String, ByVal dfactura As Boolean)
Dim StrTipoMovimiento As String * 3
Dim IntNuevoStock As Double
Dim RstStockActual As New ADODB.Recordset

strCadena = "SELECT Stock FROM Almacen_Productos WHERE (Alm_cod='" & strCodAlmacen & "' AND cProducto='" & StrCodProducto & "')"
RstStockActual.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic

        strCadena = "SELECT cTipoMovimiento FROM Comprobantes WHERE doc_cod='" & Trim(TipoDoc) & "'"
        Call ConfiguraRst(strCadena)
         StrTipoMovimiento = rst(0)
         Set rst = Nothing
 If Left(StrTipoMovimiento, 1) = "E" Then
    IntNuevoStock = RstStockActual(0) + IntCantMov
  Else
    IntNuevoStock = RstStockActual(0) - IntCantMov
    
  End If
  
  Set RstStockActual = Nothing
If (dfactura = False) Then
    strCadena = "UPDATE Almacen_Productos SET Stock='" & IntNuevoStock & "' WHERE (Alm_cod='" & strCodAlmacen & "' AND cProducto='" & StrCodProducto & "')"
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
    strCadena = "UPDATE Producto SET StockActual='" & IntNuevoStock & "' WHERE ( cProducto='" & StrCodProducto & "')"
    Call EjecutaRST(strCadena)
    
    Set RstEjecuta = Nothing
End If

  '---actualizar el stock general-----

  
 ' Call ActualizaStock(StrCodProducto, IntCantMov, StrTipoMovimiento, TipoDoc, Doc_ref, serie_ref, num_ref, dfactura)
  

End Sub
Public Sub ActualizaStock(ByVal StrCodProducto As String, ByVal IntCantMov As Double, ByVal TipoMovimiento As String, ByVal TipoDoc As String, ByVal doc_ref As String, ByVal serie_ref As String, ByVal num_ref As String, ByVal dfactura As Boolean)
Dim IntNuevoStock As Double
Dim IntNuevoStock_factura As Double
strCadena = "SELECT StockActual,stock_factura FROM Producto WHERE cProducto='" & StrCodProducto & "'"
Call ConfiguraRst(strCadena)
  If Left(TipoMovimiento, 1) = "E" Then
    IntNuevoStock = rst(0) + IntCantMov
    IntNuevoStock_factura = rst(1) + IntCantMov
  Else
    IntNuevoStock = rst(0) - IntCantMov
        
        If (doc_ref = "0001" Or doc_ref = "0003") Then
            IntNuevoStock_factura = rst(1) - IntCantMov
        Else
            IntNuevoStock_factura = rst(1)
        End If
    
  End If
  Set rst = Nothing
    
    
    If (dfactura = True) Then
          strCadena = "UPDATE Producto SET stock_factura='" & IntNuevoStock_factura & "' WHERE  cProducto='" & StrCodProducto & "'"
          Call EjecutaRST(strCadena)
          Set RstEjecuta = Nothing
          GoTo 50
    End If
    
    
    If ((doc_ref = "0001" Or doc_ref = "0003") And serie_ref <> "0000" And num_ref <> "0000000000") Then
          strCadena = "UPDATE Producto SET StockActual='" & IntNuevoStock & "',stock_factura='" & IntNuevoStock_factura & "' WHERE  cProducto='" & StrCodProducto & "'"
    Else
          strCadena = "UPDATE Producto SET StockActual='" & IntNuevoStock & "' WHERE  cProducto='" & StrCodProducto & "'"
          
    End If
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
50:
  End Sub

Public Sub insertar_kardex_producto(ByVal in_compra As String, ByVal in_venta As String, ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_proveedor As String, ByVal in_producto As String, ByVal in_alm As String, ByVal in_fecha As Date, ByVal in_cantidad As Single, ByVal in_costo As Single)
Dim in_stock As Double
Dim nprecio_compra_anterior As Single
Dim in_costo_promedio As Single

'*** CLACULO DE STOCK ACTUAL

strCadena = "SELECT sum(cantidad_real) FROM kardex WHERE  id_producto='" & in_producto & "' AND id_alm='" & in_alm & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRstZ(strCadena)
If IsNull(rstZ(0)) = True Then
    in_stock = 0
Else
    in_stock = rstZ(0)
End If

'*** PRECIO DE COSTO ACTUAL
strCadena = "SELECT costo_unitario FROM kardex  where id_producto='" & in_producto & "' AND id_alm='" & in_alm & "' AND ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC,id_kardex DESC LIMIT 1 "
Call ConfiguraRstZ(strCadena)
If IsNull(rstZ(0)) = True Then
     nprecio_compra_anterior = 0
Else
    nprecio_compra_anterior = rstZ(0)
End If

'*** COSTO PROMEDIO
in_costo_promedio = (in_cantidad * in_costo + in_stock * nprecio_compra_anterior) / (in_cantidad + in_stock)

strCadena = "INSERT INTO kardex(id_movimiento,fecha_emision,id_doc,id_serie,id_numero,id_alm,id_producto,cantidad,cantidad_ing,cantidad_real,costo_unitario, " & _
"costo_promedio,id_persona,ncliente,dni_save,ruc)VALUES " & _
" ('" & in_compra & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & in_doc & "','" & in_serie & "','" & in_numero & "','" & in_alm & "','" & in_producto & "','" & in_cantidad & "','" & in_cantidad & "','" & in_cantidad & "','" & in_costo & "','" & in_costo_promedio & "','" & in_proveedor & "','" & get_persona(in_proveedor) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
End Sub
Public Sub put_inventario(ByVal in_producto As String, ByVal in_alm As String, ByVal in_stock_nuevo As Single)
Dim cod_articulo As String
Dim in_periodo As String
in_periodo = "1CIX000000000031"

strCadena = "SELECT A.id_producto,U.abreviatura,A.stock,P.nombre_prod,A.precio_venta,A.precio_compra FROM almacen_producto A,producto P,unidad U WHERE A.id_producto=P.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND A.id_producto='" & Trim(in_producto) & "' AND A.id_alm='" & in_alm & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 And in_stock_nuevo >= 0 Then

    cod_articulo = in_producto
    stock_actual = rst("stock")
    strInventario = formato_item(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), 6)
    
    strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
    "('" & strInventario & "','" & cod_articulo & "','0106','001','" & strInventario & "','" & Val(rst("precio_compra")) & "','" & KEY_FECHA & "','" & in_alm & "','" & Val(in_stock_nuevo) & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    in_cta_compra = KEY_CTA_COMPRA_SOLES
    
    If in_stock_nuevo > Val(stock_actual) Then
       strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0089' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             in_numero = Format(1, "00000000")
             in_serie = "0001"
       End If
            in_cantidad = Val(in_stock_nuevo) - Val(stock_actual)
           
            strCadena = "call P_insert_compra_ultimate('0089','" & in_alm & "',CURDATE(),CURDATE(),'02'," & _
            "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
            "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
            "'0','0','0','0','0','0','0','0','0','0','0'," & _
            " '" & KEY_USUARIO & "','OBSERVACION','01','" & in_periodo & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
            Call ConfiguraRstP(strCadena)
            id_compra = rstP(0)
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(in_producto) & "','" & in_cantidad & "','0'," & _
           "'0','0','0','" & in_cantidad * Val(rst("precio_compra")) & "','0','0', " & _
           "'0','0','0','" & in_cantidad * Val(rst("precio_compra")) & "','0','" & rst("precio_compra") * in_cantidad & "','" & Val(rst("precio_venta")) & "','" & Val(rst("precio_compra")) & "','" & in_alm & "','" & rst("nombre_prod") & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
           strCadena = "call put_kardex_stock_vitekey('04','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0089','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(in_producto) & "','" & in_cantidad & "','" & Val(rst("precio_compra")) & "','" & in_alm & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
        
    Else
        strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0090' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
        Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             in_numero = Format(1, "00000000")
             in_serie = "0001"
       End If
            in_cantidad = Val(stock_actual) - Val(in_stock_nuevo)
        
            strCadena = "call P_insert_compra_ultimate('0090','" & in_alm & "',CURDATE(),CURDATE(),'02'," & _
            "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
            "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
            "'0','0','0','0','0','0','0','0','0','0','0'," & _
            " '" & KEY_USUARIO & "','OBSERVACION','01','" & in_periodo & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
            Call ConfiguraRstP(strCadena)
            id_compra = rstP(0)
             strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(in_producto) & "','" & in_cantidad & "','0'," & _
           "'0','0','0','" & in_cantidad * Val(rst("precio_compra")) & "','0','0', " & _
           "'0','0','0','" & in_cantidad * Val(rst("precio_compra")) & "','0','" & rst("precio_compra") * in_cantidad & "','" & Val(rst("precio_venta")) & "','" & Val(rst("precio_compra")) & "','" & in_alm & "','" & rst("nombre_prod") & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
            strCadena = "call put_kardex_stock_vitekey('01','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0090','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(in_producto) & "','" & in_cantidad & "','" & Val(rst("precio_compra")) & "','" & in_alm & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        End If
        
        
  End If
   
    
    in_comentario = "INVENTARIO:" & Space(2) & KEY_VENDEDOR + Chr(13) + "CONTEO FISICO :" + str(stock_actual) + Chr(13) + "AJUSTE :" + str(in_cantidad)
    
    strCadena = "UPDATE producto SET  inventario='si',comentario='" & in_comentario & "' WHERE id_producto='" & Trim(in_producto) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    


    



End Sub

Public Sub put_costo_promedio(ByVal in_kardex As Double, ByVal in_producto As String, ByVal in_costo_promedio As Double, ByVal in_cantidad_real As Double, ByVal in_alm As String)

strCadena = "UPDATE kardex SET costo_promedio='" & in_costo_promedio & "'  WHERE id_kardex='" & in_kardex & "' and id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
CnBd.Execute (strCadena)

'strCadena = "UPDATE almacen_producto SET precio_compra='" & in_costo_promedio & "' WHERE id_alm='" & in_alm & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
'CnBd.Execute (strCadena)


End Sub

Public Sub put_costo_promedio_venta(ByVal in_kardex As Double, ByVal in_producto As String, ByVal in_costo_promedio As Double, ByVal in_cantidad_real As Double, ByVal in_alm As String, ByVal in_movimiento As String)

strCadena = "UPDATE kardex SET costo_promedio='" & in_costo_promedio & "'  WHERE id_kardex='" & in_kardex & "' and id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
CnBd.Execute (strCadena)

strCadena = "UPDATE almacen_producto SET precio_compra='" & in_costo_promedio & "' WHERE id_alm='" & in_alm & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


End Sub

Public Function get_costo_unitario_fisico(ByVal in_movimiento As String, ByVal in_cantidad_real As Double, ByVal in_doc As String, ByVal in_alm As String, ByVal in_producto As String, Optional in_costo_anterior As Double)
Dim in_costo As Double
Dim in_costov1 As Double

If in_cantidad_real > 0 And in_doc = "0009" Then
    strCadena = "SELECT * FROM movimiento_transferencia_detalle WHERE id_transferencia='" & Val(in_movimiento) & "' and id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' limit 1"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
       get_costo_unitario_fisico = rstL("precio_costo") / 1.18
       If get_costo_unitario_fisico <> in_costo_anterior Then
          get_costo_unitario_fisico = in_costo_anterior
          strCadena = "UPDATE movimiento_transferencia_detalle SET precio_costo='" & in_costo_anterior & "' WHERE id_transferencia='" & Val(in_movimiento) & "' and id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' limit 1"
          CnBd.Execute (strCadena)
       End If
    End If
    Exit Function
End If

If in_cantidad_real > 0 And in_doc <> "0007" Then
    strCadena = "SELECT * FROM view_compra_detalle WHERE cantidad='" & Abs(in_cantidad_real) & "' and  id_producto='" & in_producto & "' and  id_compra='" & Val(in_movimiento) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
       If rstL("id_moneda") = "00002" Then ' DOLARES
          get_costo_unitario_fisico = (rstL("valor_venta") * rstL("tc") + rstL("incremento_neto_gasto")) / rstL("cantidad")
       Else
          get_costo_unitario_fisico = (rstL("valor_venta") + rstL("incremento_neto_gasto")) / rstL("cantidad")
       End If
    Else
        strCadena = "SELECT * FROM view_kardex WHERE id_tipo_movimiento<>'10' and  id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            in_costo = rstL("costo_promedio")
        Else
            in_costo = 0
        End If
                
        strCadena = "SELECT precio_compra,precio_venta FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            in_costov1 = rstL("precio_compra")
            If in_costo <= in_costov1 Then
                get_costo_unitario_fisico = in_costo
            Else
                If in_costov1 * 1.18 < rstL("precio_venta") Then
                   get_costo_unitario_fisico = in_costov1
                Else
                   get_costo_unitario_fisico = rstL("precio_venta") - rstL("precio_venta") * 10 / 100
                End If
            End If
            
            
        End If
        
    End If
Else
        
        
        strCadena = "SELECT * FROM view_kardex WHERE id_tipo_movimiento<>'10' and id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            in_costo = rstL("costo_promedio")
        Else
            in_costo = 0
        End If
                
        strCadena = "SELECT precio_compra,precio_venta FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            in_costov1 = rstL("precio_compra")
            If in_costo <= in_costov1 Then
                
                get_costo_unitario_fisico = in_costo
                 If in_costo * 1.18 < rstL("precio_venta") Then
                   get_costo_unitario_fisico = in_costo
                Else
                   get_costo_unitario_fisico = in_costo / 1.18
                End If
            
            Else
                If in_costov1 * 1.18 < rstL("precio_venta") Then
                   get_costo_unitario_fisico = in_costov1
                Else
                   get_costo_unitario_fisico = in_costov1 / 1.18
                End If
            End If
            
            
        End If
        
        
        
End If


End Function
Public Sub put_update_saldo_stock_contable(ByVal in_producto As String, ByVal in_alm As String)
 'STOCK FISICO
  Dim in_saldo_actual As Double
  Dim in_tag As Boolean
reiniciar_contable:
        in_tag = False
        strCadena = "call put_crear_kardex_id_producto('" & in_producto & "','" & in_alm & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        in_costo = 0
        in_saldo = 0
        in_cantidad_real = 0
        
        strCadena = "SELECT * FROM tmp_kardex_producto WHERE cantidad_factura<>0 and id_doc IN('0001','0003','0007','0009','0089','0090') and id_alm='" & in_alm & "' and id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            rstK.MoveFirst
            
            strCadena = "select stock_factura from  almacen_producto  WHERE id_producto='" & in_producto & "' and  id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            Call ConfiguraRstPP(strCadena)
            If rstPP.RecordCount > 0 Then
                in_saldo_actual = rstPP(0)
            Else
                in_saldo_actual = 0
            End If
            
            
            For j = 0 To rstK.RecordCount - 1
                   in_cantidad_real = rstK("cantidad_real")
                   
                   If rstK("id_tipo_movimiento") = "10" Then
                        in_cantidad_real = rstK("cantidad_factura")
                   End If
                   
                in_saldo = in_saldo + Val(in_cantidad_real)
                
                If j = rstK.RecordCount - 1 Then
                    If Val(in_saldo) <> Val(in_saldo_actual) Then
                        strCadena = "UPDATE almacen_producto SET stock_factura='" & Val(in_saldo) & "' WHERE id_producto='" & in_producto & "' and  id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                        CnBd.Execute (strCadena)
                    End If
                End If
                
                
                
                
                If Val(in_saldo) <> rstK("saldo_stock_contable") Then
                    strCadena = "UPDATE kardex SET saldo_stock_contable='" & Val(in_saldo) & "' WHERE id_kardex='" & rstK("id_kardex") & "'  LIMIT 1"
                    CnBd.Execute (strCadena)
                    in_tag = True
                End If
                rstK.MoveNext
                
            Next j
            
            If in_tag = True Then
                GoTo reiniciar_contable
            End If
            
        End If
        
        


        
  
End Sub
Public Sub put_update_saldo_stock_fisico(ByVal in_producto As String, ByVal in_alm As String)
Dim in_saldo_actual As Double
Dim in_tag As Boolean
'STOCK FISICO

reiniciar_fisico:
in_tag = False
        strCadena = "call put_crear_kardex_id_producto('" & in_producto & "','" & in_alm & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        in_costo = 0
        in_saldo = 0
        strCadena = "SELECT * FROM tmp_kardex_producto WHERE id_tipo_movimiento<>'10' and id_alm='" & in_alm & "' and id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            rstK.MoveFirst
            
            strCadena = "select stock from  almacen_producto  WHERE id_producto='" & in_producto & "' and  id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            Call ConfiguraRstPP(strCadena)
            If rstPP.RecordCount > 0 Then
                in_saldo_actual = rstPP(0)
            Else
                in_saldo_actual = 0
            End If
                
            For j = 0 To rstK.RecordCount - 1
                in_saldo = in_saldo + rstK("cantidad_real")
                
                If j = rstK.RecordCount - 1 Then
                    If Val(in_saldo) <> Val(in_saldo_actual) Then
                        strCadena = "UPDATE almacen_producto SET stock='" & Val(in_saldo) & "' WHERE id_producto='" & in_producto & "' and  id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                        CnBd.Execute (strCadena)
                    End If
                End If
                
                
                If Val(in_saldo) <> rstK("saldo_stock") Then
                    strCadena = "UPDATE kardex SET saldo_stock='" & Val(in_saldo) & "' WHERE id_kardex='" & rstK("id_kardex") & "' LIMIT 1"
                    CnBd.Execute (strCadena)
                    in_tag = True
                    
                End If
                rstK.MoveNext
                
            Next j
            If in_tag = True Then
                GoTo reiniciar_fisico
            End If
        End If
        
     

End Sub
Public Sub put_update_saldo_stock_all(ByVal in_producto As String, ByVal in_alm As String)
Dim in_saldo_actual As Double
Dim in_tag As Boolean
'STOCK FISICO

reiniciar_fisico:
in_tag = False
        strCadena = "call put_crear_kardex_id_producto('" & in_producto & "','" & in_alm & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        in_costo = 0
        in_saldo = 0
        strCadena = "SELECT * FROM tmp_kardex_producto WHERE  id_alm='" & in_alm & "' and id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            rstK.MoveFirst
            
            strCadena = "select stock from  almacen_producto  WHERE id_producto='" & in_producto & "' and  id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            Call ConfiguraRstPP(strCadena)
            If rstPP.RecordCount > 0 Then
                in_saldo_actual = rstPP(0)
            Else
                in_saldo_actual = 0
            End If
                
            
            For j = 0 To rstK.RecordCount - 1
                in_saldo = in_saldo + rstK("cantidad_real")
                
                If j = rstK.RecordCount - 1 Then
                    If Val(in_saldo) <> Val(in_saldo_actual) Then
                        strCadena = "UPDATE almacen_producto SET stock='" & Val(in_saldo) & "' WHERE id_producto='" & in_producto & "' and  id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                        CnBd.Execute (strCadena)
                    End If
                End If
                
                in_fecha = rstK("fecha_emision")
                If Val(in_saldo) <> rstK("saldo_stock") Then
                    strCadena = "UPDATE kardex SET saldo_stock='" & Val(in_saldo) & "' WHERE id_kardex='" & rstK("id_kardex") & "' LIMIT 1"
                    CnBd.Execute (strCadena)
                    GoTo reiniciar_fisico
                    
                End If
                
                
                rstK.MoveNext
                
            Next j
            If in_tag = True Then
                GoTo reiniciar_fisico
            End If
        End If
        
     

End Sub


Public Sub put_update_costo_promedio_contable(ByVal in_producto As String, ByVal in_alm As String)
reiniciar_contable:

strCadena = "SELECT * FROM kardex WHERE cantidad_factura<>0 and id_doc IN('0001','0003','0007','0009','0089','0090') and  id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC "
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   
   For i = 0 To rstA.RecordCount - 1
        
        in_saldo_cantidad = 0
        in_valorizado_anterior = 0
        in_valorizado_actual = 0
        
            'INGRESOS DE MERCADERIA
            If i = 0 Then
                in_costo_unitario = get_costo_unitario_fisico(rstA("id_movimiento"), rstA("cantidad_real"), rstA("id_doc"), in_alm, in_producto)
                in_costo_promedio = in_costo_unitario
                If Round(Val(in_costo_promedio), 2) <> Round(rstA("costo_promedio"), 2) Then
                    strCadena = "UPDATE kardex SET costo_unitario='" & Val(in_costo_promedio) & "' WHERE id_kardex='" & rstA("id_kardex") & "' and ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    Call put_costo_promedio(rstA("id_kardex"), in_producto, Val(in_costo_promedio), rstA("cantidad_real"), in_alm)
                    GoTo reiniciar_contable
                End If
            Else
                If rstA("cantidad_real") > 0 Then   ' *************Ingresos
                    
                    in_costo_unitario = get_costo_unitario_fisico(rstA("id_movimiento"), rstA("cantidad_real"), rstA("id_doc"), in_alm, in_producto)
                    rstA.MovePrevious
                    in_valorizado_anterior = rstA("saldo_stock") * rstA("costo_promedio")
                    in_costo_promedio = rstA("costo_promedio")
                    rstA.MoveNext
                    in_valorizado_actual = rstA("cantidad") * rstA("costo_unitario")
                    If rstA("saldo_stock") <> 0 Then
                        in_costo_promedio = (in_valorizado_anterior + in_valorizado_actual) / rstA("saldo_stock")
                    End If
                    
                    
                    If Round(Val(in_costo_promedio), 8) <> Round(rstA("costo_promedio"), 8) Then
                        Call put_costo_promedio(rstA("id_kardex"), in_producto, Val(in_costo_promedio), rstA("cantidad_real"), in_alm)
                        GoTo reiniciar_contable
                    End If
                    
                Else
                                                    '**************Salidas
                    rstA.MovePrevious
                    in_costo_unitario = rstA("costo_promedio")
                    in_valorizado_anterior = rstA("saldo_stock") * rstA("costo_promedio")
                    rstA.MoveNext
                    in_valorizado_actual = rstA("cantidad") * in_costo_unitario
                    in_costo_promedio = in_costo_unitario
                    
                    If Round(Val(in_costo_promedio), 2) <> Round(rstA("costo_promedio"), 2) Or Round(Val(in_costo_promedio), 2) <> Round(Val(rstA("costo_unitario")), 2) Then
                        Call put_costo_promedio(rstA("id_kardex"), in_producto, Val(in_costo_promedio), rstA("cantidad_real"), in_alm)
                        GoTo reiniciar_contable
                    End If
                End If
        End If
                
            
        If i = rstA.RecordCount - 1 Then
            If Round(Val(in_costo_promedio), 2) <> Round(get_costo_almacen(in_producto, in_alm), 2) Then
                strCadena = "UPDATE almacen_producto SET precio_compra='" & in_costo_promedio & "' WHERE id_alm='" & in_alm & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
            End If
        End If
        
        
        
        
        
        rstA.MoveNext
        DoEvents
   Next i
   
End If

End Sub
Public Sub put_update_costo_promedio_fisico(ByVal in_producto As String, ByVal in_alm As String)
Dim p As Double
Dim nanterior  As Double
reiniciar_fisico:
strCadena = "SELECT * FROM kardex WHERE id_tipo_movimiento<>'10' and  id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC "
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   
   For i = 0 To rstA.RecordCount - 1
        in_saldo_cantidad = 0
        in_valorizado_anterior = 0
        in_valorizado_actual = 0
        
            
            If i = 0 Then
                in_costo_unitario = get_costo_unitario_fisico(rstA("id_movimiento"), rstA("cantidad_real"), rstA("id_doc"), in_alm, in_producto)
                in_costo_promedio = in_costo_unitario
                If Round(Val(in_costo_promedio), 2) <> Round(rstA("costo_promedio"), 2) Then
                    strCadena = "UPDATE kardex SET costo_unitario='" & Val(in_costo_promedio) & "',costo_promedio='" & Val(in_costo_promedio) & "' WHERE id_kardex='" & rstA("id_kardex") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                    CnBd.Execute (strCadena)
                    'Call put_costo_promedio(rstA("id_kardex"), in_producto, Val(in_costo_promedio), rstA("cantidad_real"), in_alm)
                    GoTo reiniciar_fisico
                End If
            Else
                If rstA("cantidad_real") > 0 Then   ' *************Ingresos
                     rstA.MovePrevious
                    nanterior = rstA("costo_promedio")
                     rstA.MoveNext
                    in_costo_unitario = get_costo_unitario_fisico(rstA("id_movimiento"), rstA("cantidad_real"), rstA("id_doc"), in_alm, in_producto, nanterior)
                    rstA.MovePrevious
                    
                    in_valorizado_anterior = rstA("saldo_stock") * rstA("costo_promedio")
                    in_costo_promedio = rstA("costo_promedio")
                    rstA.MoveNext
                    
                    in_valorizado_actual = rstA("cantidad") * in_costo_unitario
                    If rstA("saldo_stock") <> 0 Then
                        in_costo_promedio = (in_valorizado_anterior + in_valorizado_actual) / rstA("saldo_stock")
                    End If
                    
                    
                    If Round(Val(in_costo_promedio), 2) <> Round(rstA("costo_promedio"), 2) Then
                        Call put_costo_promedio(rstA("id_kardex"), in_producto, Val(in_costo_promedio), rstA("cantidad_real"), in_alm)
                        GoTo reiniciar_fisico
                    End If
                    
                Else
                                                    '**************Salidas
                    rstA.MovePrevious
                    in_costo_unitario = rstA("costo_promedio")
                    'in_valorizado_anterior = rstA("saldo_stock") * rstA("costo_promedio")
                    rstA.MoveNext
                   ' in_valorizado_actual = rstA("cantidad") * in_costo_unitario
                    in_costo_promedio = in_costo_unitario
                    
                    If Round(Val(in_costo_promedio), 2) <> Round(rstA("costo_promedio"), 2) Then
                        strCadena = "UPDATE kardex SET costo_promedio='" & Val(in_costo_promedio) & "',costo_unitario='" & Val(in_costo_promedio) & "' WHERE id_kardex='" & rstA("id_kardex") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                        CnBd.Execute (strCadena)
                        GoTo reiniciar_fisico
                    End If
                End If
        End If
                
            
        If i = rstA.RecordCount - 1 Then
            If Round(Val(in_costo_promedio), 2) <> Round(get_costo_almacen(in_producto, in_alm), 2) Then
                strCadena = "UPDATE almacen_producto SET precio_compra='" & in_costo_promedio & "' WHERE id_alm='" & in_alm & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
            End If
        End If
        rstA.MoveNext
        DoEvents
   Next i
   
End If






                    
End Sub

Private Function get_costo_almacen(ByVal in_producto As String, ByVal in_alm As String)

strCadena = "SELECT precio_compra FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
get_costo_almacen = rstL("precio_compra")

End Function
Public Sub put_update_costo_promedio_saldo_stock(ByVal in_producto As String, ByVal in_alm As String)

reiniciar:
strCadena = "SELECT * FROM kardex WHERE id_tipo_movimiento<>'10' and  id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC "
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   
   For i = 0 To rstA.RecordCount - 1
        
        in_saldo_cantidad = 0
        in_valorizado_anterior = 0
        in_valorizado_actual = 0
        
            'INGRESOS DE MERCADERIA
            If i = 0 Then
                in_costo_unitario = get_costo_unitario_fisico(rstA("id_movimiento"), rstA("cantidad_real"), rstA("id_doc"), in_alm, in_producto)
                in_costo_promedio = in_costo_unitario
                If Val(in_costo_promedio) <> rstA("costo_promedio") Then
                    strCadena = "UPDATE kardex SET costo_unitario='" & Val(in_costo_promedio) & "' WHERE id_kardex='" & rstA("id_kardex") & "' and ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    Call put_costo_promedio(rstA("id_kardex"), in_producto, Val(in_costo_promedio), rstA("cantidad_real"), in_alm)
                    GoTo reiniciar
                End If
            Else
                If rstA("cantidad_real") > 0 Then   ' *************Ingresos
                    
                    in_costo_unitario = get_costo_unitario_fisico(rstA("id_movimiento"), rstA("cantidad_real"), rstA("id_doc"), in_alm, in_producto)
                    rstA.MovePrevious
                    in_valorizado_anterior = rstA("saldo_stock") * rstA("costo_promedio")
                    rstA.MoveNext
                    in_valorizado_actual = rstA("saldo_stock") * rstA("costo_unitario")
                    in_costo_promedio = (in_valorizado_anterior + in_valorizado_actual) / rstA("saldo_stock")
                    If Val(in_costo_promedio) <> rstA("costo_promedio") Then
                        Call put_costo_promedio(rstA("id_kardex"), in_producto, Val(in_costo_promedio), rstA("cantidad_real"), in_alm)
                        GoTo reiniciar
                    End If
                Else
                                                    '**************Salidas
                    
                    rstA.MovePrevious
                    in_costo_unitario = rstA("costo_promedio")
                    in_valorizado_anterior = rstA("saldo_stock") * rstA("costo_promedio")
                    rstA.MoveNext
                    in_valorizado_actual = rstA("saldo_stock") * in_costo_unitario
                    in_costo_promedio = in_costo_unitario
                    If Val(in_costo_promedio) <> rstA("costo_promedio") Then
                        Call put_costo_promedio(rstA("id_kardex"), in_producto, Val(in_costo_promedio), rstA("cantidad_real"), in_alm)
                        GoTo reiniciar
                    End If
                End If
            
                
        
        
        
        End If
                
            
        
        
        
        
        
        rstA.MoveNext
        DoEvents
   Next i
   
   
   
   
   
End If



                    
End Sub

Public Sub put_actualizar_kardex_update(ByVal in_producto As String, ByVal in_alm As String)

strCadena = "SELECT DISTINCT id_producto,id_alm FROM kardex where id_producto='" & Trim(in_producto) & "' and id_alm='" & in_alm & "' and  ruc='" & KEY_RUC & "' ORDER BY id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    in_costo = 0
    in_saldo = 0
    For i = 0 To rst.RecordCount - 1
    
    
    
ini:
      in_costo = 0
      in_saldo = 0
        strCadena = "call put_crear_kardex_id_producto('" & rst("id_producto") & "','" & rst("id_alm") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        strCadena = "SELECT * FROM tmp_kardex_producto WHERE id_alm='" & rst("id_alm") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then

            rstK.MoveFirst
            For j = 0 To rstK.RecordCount - 1
                in_saldo = in_saldo + rstK("cantidad_real")
                If Val(in_saldo) <> rstK("saldo_stock") Then
                    strCadena = "UPDATE kardex SET saldo_stock='" & Val(in_saldo) & "' WHERE id_producto='" & rst("id_producto") & "' and  id_kardex='" & rstK("id_kardex") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                    CnBd.Execute (strCadena)
                    GoTo ini
                End If
                
                If j = rstK.RecordCount - 1 Then
                    strCadena = "SELECT sum(cantidad_real) FROM tmp_kardex_producto WHERE id_alm='" & rst("id_alm") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstL(strCadena)
                    If rstL(0) <> Val(in_saldo) Then
                        MsgBox "CANTIDAD REAL:" & rstL(0) + Chr(13) + "CANTIDAD FINAL:" & str(in_saldo)
                    End If
                End If
                rstK.MoveNext
                
            Next j
        End If
        in_costo = 0
        in_saldo = 0
        

        
       
        
       
        rst.MoveNext
        DoEvents
        
    Next i
End If
End Sub

Public Function get_precio_venta_now(ByVal in_producto As String) As Single
strCadena = "SELECT precio_venta FROM almacen_producto where id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   get_precio_venta_now = rstP("precio_venta")
End If
End Function

Public Function get_cantidad_agranel(ByVal in_producto As String, ByVal in_unidad As String)
strCadena = "SELECT agranel FROM producto WHERE agranel='si' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
    strCadena = "SELECT cantidad FROM producto_unidad WHERE id_producto='" & in_producto & "' and id_unidad='" & in_unidad & "' and ruc='" & KEY_RUC & "' LIMIT 1 "
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount > 0 Then
        get_cantidad_agranel = rstA("cantidad")
    Else
        get_cantidad_agranel = 1
    End If
Else
    get_cantidad_agranel = 1
End If


    
End Function


Public Sub put_detalle_consumo_combo(ByVal in_combo As String, ByVal in_cantidad As Single)
Dim in_costo_acum As Double
strCadena = "SELECT * FROM view_combo_detalle WHERE id_productoc='" & in_combo & "' and   ruc='" & KEY_RUC & "' "
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   in_costo_acum = 0
   
   For i = 0 To rstA.RecordCount - 1
        
        strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0090' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
        Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             
             strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "' and id_alm='" & KEY_ALM & "' and id_doc='0090' LIMIT 1"
             Call ConfiguraRstlocal(strCadena)
             If rstLocal.RecordCount > 0 Then
               in_serie = rstLocal("serie")
               in_numero = formato_item(rstLocal("numero"), 8)
             Else
                MsgBox "CREE EL COMPROBANTE SALIDA A ALMACEN" + Chr(13) + "PARA ESTA SUCURSAL", vbInformation
                Exit Sub
             End If
             
        End If
        
        in_cta_compra = KEY_CTA_COMPRA_SOLES
        in_costo = get_costo_ultimo(rstA("id_producto"), KEY_FECHA)
        in_total = in_costo * rstA("cantidad")
        in_costo_acum = in_costo_acum + in_total
        
        strCadena = "call P_insert_compra_ultimate('0090','" & KEY_ALM & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','02'," & _
        "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
        "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
        "'0','" & in_valor_venta & "','" & in_igv & "','0','0','0','0','0','0','" & in_total & "','0'," & _
        " '" & KEY_USUARIO & "','OBSERVACION','01','" & get_periodo_actual(KEY_FECHA) & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
        Call ConfiguraRstP(strCadena)
        
        id_compra = rstP(0)
                
                
                in_total = in_costo * rstA("cantidad")
                If KEY_CON_IGV = "si" Then
                    in_valor_venta = in_total
                    in_igv = 0
                Else
                    in_valor_venta = in_total
                    in_igv = 0
                End If
                
                strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
                "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(rstA("id_producto")) & "','" & rstA("cantidad") & "','" & Val(in_costo) & "'," & _
                "'0','0','0','" & in_valor_venta & "','" & Val(in_igv) & "','0', " & _
                "'0','0','0','" & in_valor_venta & "','0','" & Val(in_costo) * rstA("cantidad") & "','" & Val(in_costo) & "','" & Val(in_costo) & "','" & KEY_ALM & "','" & rstA("nombre_prod") & "','0','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
               
                strCadena = "call put_kardex_stock_vitekey_v1('02','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0090','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(rstA("id_producto")) & "','" & rstA("cantidad") & "','" & Val(in_costo) & "','" & KEY_ALM & "','" & KEY_USUARIO & "','no','" & KEY_RUC & "')"
                
                CnBd.Execute (strCadena)
                
                rstA.MoveNext
                
                
   Next i
   
   
            
            in_total = in_costo_acum
                If KEY_CON_IGV = "si" Then
                    in_valor_venta = in_total / (1 + KEY_IGV)
                    in_igv = in_total - in_valor_venta
                Else
                    in_valor_venta = in_total
                    in_igv = 0
                End If
            
            
            strCadena = "call P_insert_compra_ultimate('0089','" & KEY_ALM & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','02'," & _
            "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
            "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
            "'0','" & Val(in_valor_venta) & "','" & Val(in_igv) & "','0','0','0','0','0','0','" & Val(in_total) & "','" & Val(in_total) & "'," & _
            " '" & KEY_USUARIO & "','OBSERVACION','01','" & get_periodo_actual(KEY_FECHA) & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
            Call ConfiguraRstP(strCadena)
            id_compra = rstP(0)
           
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(in_combo) & "','" & in_cantidad & "','" & Val(in_total) & "'," & _
           "'0','0','0','" & Val(in_total) & "','0','0', " & _
           "'0','0','0','" & Val(in_total) & "','0','" & Val(in_total) & "','" & Val(get_precio_venta_now(in_combo)) & "','" & Val(in_total) & "','" & KEY_ALM & "','" & get_producto(in_combo) & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
           
           
           
           strCadena = "call put_kardex_stock_vitekey_v1('02','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0089','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(in_combo) & "','" & in_cantidad & "','" & Val(in_total) & "','" & KEY_ALM & "','" & KEY_USUARIO & "','no','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
           
           
   
   
End If




End Sub
Public Function get_costo_ultimo(ByVal in_producto As String, ByVal in_fecha As String) As Double

strCadena = "SELECT funct_costo_final('" & in_producto & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & KEY_ALM & "','" & KEY_RUC & "')"
Call ConfiguraRstL(strCadena)
get_costo_ultimo = rstL(0)
End Function


Private Sub update_kardex_temp(ByVal in_producto As String, ByVal in_fecha_kardex As Date)
    Dim in_dias As Integer
    Dim in_flag As Boolean
    
    
    If verificacion_servicio(in_producto) = True Then
        Exit Sub
    End If
    
    strCadena = "DELETE FROM kardex WHERE fecha_emision>='" & Format(in_fecha_kardex, "YYYY-mm-dd") & "' and    id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    in_fechai = in_fecha_kardex
    in_dias = DateDiff("d", in_fechai, KEY_FECHA)
            
            
            
            For m = 0 To in_dias
                
                If Format(in_fechai, "YYYY-mm-dd") <= KEY_FECHA Then
                    
                    '--------------COMPRAS----------------------
                    strCadena = "select * from view_kardex_compra_existe WHERE fecha_kardex='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                    Call ConfiguraRstL(strCadena)
                    If rstL.RecordCount > 0 Then
                        Call compras_producto(in_fechai, in_producto)
                    End If
                    
                    '--------------TRANSFERENCIAS----------------------
                    strCadena = "select id_transferencia from movimiento_transferencia WHERE fecha='" & Format(in_fechai, "YYYY-mm-dd") & "'and ruc='" & KEY_RUC & "' LIMIT 1"
                    Call ConfiguraRstZ(strCadena)
                    If rstZ.RecordCount > 0 Then
                        strCadena = "select * from view_transferencia_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                        Call ConfiguraRstL(strCadena)
                        If rstL.RecordCount > 0 Then
                            Call transferencia_ingreso_producto(in_fechai, in_producto)
                            Call transferencia_salida_producto(in_fechai, in_producto)
                        End If
                    End If
                    
                    '--------------VENTAS------------------------------
                    strCadena = "select * from view_kardex_ventas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                    Call ConfiguraRstL(strCadena)
                    If rstL.RecordCount > 0 Then
                        Call ventas_producto(in_fechai, in_producto)
                    End If
                    
                    '--------------NOTAS DE SALIDAS----------------------
                    strCadena = "select * from view_kardex_notas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                    Call ConfiguraRstL(strCadena)
                    If rstL.RecordCount > 0 Then
                        Call notas_producto(in_fechai, in_producto)
                    End If
                End If
                DoEvents
            
                in_fechai = DateAdd("d", 1, in_fechai)
                
                
            Next m
    
                  

End Sub
Public Sub put_kardex_tipo(ByVal in_producto As String, ByVal in_fecha_ini As Date)

strCadena = "call ADM_kardex_update_temp('1','" & in_producto & "','" & Format(in_fecha_ini, "YYYY-mm-dd") & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM tmp_kardex_producto"
Call ConfiguraRst(strCadena)

If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
    
 
  
    
    Select Case rst("tipo_mov")
        Case "1"
                in_tipo = "02"
                If rst("afecta_factura") = "si" Then
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
                    If rst("igv") > 0 Then
                         If rst("obsequio") = "si" Then
                            in_costo = 0
                        Else
                            in_costo = Abs(rst("total")) / Abs(rst("cantidad")) + rst("incremento_neto") / Abs(rst("cantidad")) + rst("incremento_neto_gasto") / Abs(rst("cantidad"))
                        End If
                    Else
                        If rst("obsequio") = "si" Then
                            in_costo = 0
                        Else
                            in_costo = rst("total") / rst("cantidad") + rst("incremento_neto") / rst("cantidad") + rst("incremento_neto_gasto") / rst("cantidad")
                        End If
                    End If
                End If
       
                strCadena = "call put_kardex_stock_v16('" & in_tipo & "','" & Format(rst("fecha"), "YYYY-mm-dd") & "','" & rst("id_movimiento") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_entidad") & "','" & rst("entidad") & "','" & rst("id_producto") & "','" & Val(Abs(rst("cantidad"))) & "','" & in_costo & "','" & rst("id_alm") & "','" & rst("dni_save") & "','" & rst("obsequio") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
       Case "3"
                in_tipo = "03"
                strCadena = "call put_kardex_stock_v16('03','" & Format(rst("fecha"), "YYYY-mm-dd") & "','" & rst("id_movimiento") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_entidad") & "','" & rst("entidad") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','" & get_costo_sucursal(rst("id_producto"), rst("id_alm_origen"), rst("id_movimiento")) & "','" & rst("id_alm") & "','" & rst("dni_save") & "','" & rst("obsequio") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
       Case "4"
                in_tipo = "03"
                 strCadena = "call put_kardex_stock_v16('03','" & Format(rst("fecha"), "YYYY-mm-dd") & "','" & rst("id_movimiento") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_entidad") & "','" & rst("entidad") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','" & get_costo_sucursal(rst("id_producto"), rst("id_alm_origen"), rst("id_movimiento")) & "','" & rst("id_alm") & "','" & rst("dni_save") & "','" & rst("obsequio") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
       Case "5"
               in_tipo = "01"
               in_cantidad = rst("cantidad")
               
               strCadena = "call put_kardex_stock_v16('" & in_tipo & "','" & Format(rst("fecha"), "YYYY-mm-dd") & "','" & rst("id_movimiento") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_entidad") & "','" & rst("entidad") & "','" & rst("id_producto") & "','" & in_cantidad & "','0','" & rst("id_alm") & "','" & rst("dni_save") & "','" & rst("obsequio") & "','" & KEY_RUC & "')"
               CnBd.Execute (strCadena)
               
    Case "7"
                in_tipo = "07"
                If get_diferida_venta(rst("id_movimiento")) = False Then
                    strCadena = "call put_kardex_stock_v16('" & in_tipo & "','" & Format(rst("fecha"), "YYYY-mm-dd") & "','" & rst("id_movimiento") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_entidad") & "','" & rst("entidad") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','0','" & rst("id_alm") & "','" & rst("dni_save") & "','" & rst("obsequio") & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                End If
    
    End Select
    rst.MoveNext
Next i

End If



End Sub

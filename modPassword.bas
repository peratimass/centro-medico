Attribute VB_Name = "modPassword"
'cadena de la que tomamos los caracteres
Private Const c_caracteres = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ1234567890abcdefghijklmnñopqrstuvwxyz"

'parámetros opcionales'cuantos caracteres necesitas y la ubicación de la cadena de donde sacarlos
Public Function get_password_activation(Optional cuantos As Integer = 24, Optional Cadena As String = c_caracteres) As String
Dim i As Integer
Dim longitud As Integer
Dim espacio As String
longitud = Len(Cadena)
Randomize
espacio = ""
For i = 1 To cuantos
    
    get_password_activation = get_password_activation & espacio & UCase(Mid(Cadena, Int((longitud * Rnd) + 1), 1))
    If i Mod 4 = 0 Then
       espacio = "-"
    Else
       espacio = ""
    End If
Next i
strCadena = "UPDATE entidad_parametros SET secure='" & get_password_activation & "' WHERE cod_unico='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
End Function


Public Function verifica_existencia_persona(ByVal in_dni As String) As Boolean
On Error GoTo salir

 strCadena = "SELECT func_verifica_persona('" & Trim(in_dni) & "')"
 Call ConfiguraRstPP(strCadena)
 If rstPP(0) = "si" Then
    verifica_existencia_persona = True
 Else
    verifica_existencia_persona = False
 End If
 
 
 Exit Function
salir:
            
End Function

Public Sub put_impresion_a4(ByVal in_venta As String, ByVal in_total As Double, ByVal in_moneda As String)

strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & in_venta & "' LIMIT 1"
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount > 0 Then
                in_moneda = rstK("id_moneda")
                in_total = rstK("total")
                If rstK("id_tipo_factura") = "00002" Then
                   
                   strCadena = "SELECT `id_venta`,`fecha_emision`,`fecha_vencimiento`,doc_des,`documento`,`id_cliente`,`ncliente`,`direccion`,`exonerado`,`valor_venta`,`igv`,`total`,`sunat_key`,`sunat_hash`,`id_producto`,`detalle`,`observacion`,`cantidad`,`valor_neto`,`precio`,`todatl_detalle`,'---','" & UCase(EnLetras_moneda(in_total, in_moneda)) & "',resolucion,nro_chasis,nro_chasis,serie,modelo,color,marca,anio_fabricacion,nro_dua,nro_item,in_guia,`ruc` FROM view_factura_electronica_serial WHERE id_venta='" & Val(in_venta) & "'"
                   Call ConfiguraRst(strCadena)
                   strCadena = "SELECT id_venta,tipo_venta,descripcion,monto,cuotas FROM view_venta_pago WHERE id_venta='" & in_venta & "'"
                   Call ConfiguraRstK(strCadena)
                   Ans = ShowMultiReport(rst, "factura_elec_serial", , App.Path + "\Reportes\", , , , , rstK, "modalidad_pago")
                Else
                
                   Dim arr(0 To 2, 1 To 2) As String
                   Dim param As Variant
                   arr(0, 1) = "vendedor_proforma"
                   arr(1, 1) = "vendedor_telefono"
                   arr(2, 1) = "percepcion"
                   
                   arr(0, 2) = get_persona(rstK("id_vendedor"))
                   arr(1, 2) = get_telefono(rstK("id_vendedor"))
                   arr(2, 2) = rstK("percepcion")
                   param = arr()

                   
                    If rstK("id_doc") = "0099" Then
                        in_vencimiento = Format(DateAdd("d", 7, rstK("fecha_emision")), "dd-mm-YYYY")
                    Else
                        in_vencimiento = Format(rstK("fecha_vencimiento"), "dd-mm-YYYY")
                    End If
                    
                   strCadena = "SELECT `id_venta`,`fecha_emision`,'" & in_vencimiento & "', doc_des,`documento`,`id_cliente`,`ncliente`,`direccion`,`exonerado`,`valor_venta`,`igv`,`total`,`sunat_key`,`sunat_hash`,`id_producto`,`detalle`,`observacion`,`cantidad`,`valor_neto`,`precio`,`todatl_detalle`,unidad,'" & UCase(EnLetras_moneda(in_total, in_moneda)) & "',resolucion,in_guia,id_moneda,'-',tc,icbper,`ruc` FROM view_factura_electronica WHERE id_venta='" & Val(in_venta) & "'"
                   'strCadena = "SELECT `id_venta`,`fecha_emision`,'" & in_vencimiento & "', doc_des,`documento`,`id_cliente`,`ncliente`,`direccion`,`exonerado`,`valor_venta`,`igv`,`total`,`sunat_key`,`sunat_hash`,`id_producto`,`detalle`,`observacion`,`cantidad`,`valor_neto`,`precio`,`todatl_detalle`,unidad,'" & UCase(EnLetras_moneda(in_total, in_moneda)) & "',resolucion,in_guia,id_moneda,'-','-','-',`ruc` FROM view_factura_electronica WHERE id_venta='" & Val(in_venta) & "'"
                   Call ConfiguraRst(strCadena)
                   
                   strCadena = "SELECT id_venta,tipo_venta,descripcion,monto,cuotas FROM view_venta_pago WHERE id_venta='" & in_venta & "'"
                   Call ConfiguraRstK(strCadena)
                   Ans = ShowMultiReport(rst, "factura_elec", param, App.Path + "\Reportes\", , , , , rstK, "modalidad_pago")
                   
                   
                   
                   
                   
                
                
                
                
                
                
                
                
                
                   ' strCadena = "SELECT `id_venta`,`fecha_emision`,`fecha_vencimiento`,doc_des,`documento`,`id_cliente`,`ncliente`,`direccion`,`exonerado`,`valor_venta`,`igv`,`total`,`sunat_key`,`sunat_hash`,`id_producto`,`detalle`,`observacion`,`cantidad`,`valor_neto`,`precio`,`todatl_detalle`,'---','" & UCase(EnLetras_moneda(in_total, in_moneda)) & "',resolucion,in_guia,id_moneda,'-','-','-',`ruc` FROM view_factura_electronica WHERE id_venta='" & Val(in_venta) & "'"
                   ' Call ConfiguraRst(strCadena)
                   ' strCadena = "SELECT id_venta,tipo_venta,descripcion,monto,cuotas FROM view_venta_pago WHERE id_venta='" & in_venta & "'"
                   ' Call ConfiguraRstK(strCadena)
                   ' Ans = ShowMultiReport(rst, "factura_elec", , App.Path + "\Reportes\", , , , , rstK, "modalidad_pago")
                End If
            End If


Exit Sub

strCadena = "SELECT id_tipo_factura FROM movimiento_venta WHERE id_venta='" & in_venta & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    If rstK("id_tipo_factura") = "00002" Then
                   strCadena = "SELECT `id_venta`,`fecha_emision`,`fecha_vencimiento`,id_moneda,`documento`,`id_cliente`,`ncliente`,`direccion`,`exonerado`,`valor_venta`,`igv`,`total`,`sunat_key`,`sunat_hash`,`id_producto`,`detalle`,`observacion`,`cantidad`,`valor_neto`,`precio`,`todatl_detalle`,'---','" & UCase(EnLetras_moneda(in_total, in_moneda)) & "',resolucion,nro_chasis,nro_chasis,serie,modelo,color,marca,anio_fabricacion,nro_dua,nro_item,in_guia,`ruc` FROM view_factura_electronica_serial WHERE id_venta='" & Val(in_venta) & "'"
                   Call ConfiguraRst(strCadena)
                   strCadena = "SELECT id_venta,tipo_venta,descripcion,monto,cuotas FROM view_venta_pago WHERE id_venta='" & in_venta & "'"
                   Call ConfiguraRstK(strCadena)
                   Ans = ShowMultiReport(rst, "factura_elec_serial", , App.Path + "\Reportes\", , , , , rstK, "modalidad_pago")
                Else
                    strCadena = "SELECT `id_venta`,`fecha_emision`,`fecha_vencimiento`,id_moneda,`documento`,`id_cliente`,`ncliente`,`direccion`,`exonerado`,`valor_venta`,`igv`,`total`,`sunat_key`,`sunat_hash`,`id_producto`,`detalle`,`observacion`,`cantidad`,`valor_neto`,`precio`,`todatl_detalle`,'---','" & UCase(EnLetras_moneda(in_total, in_moneda)) & "',resolucion,in_guia,id_moneda,'-','-','-',`ruc` FROM view_factura_electronica WHERE id_venta='" & Val(in_venta) & "'"
                    Call ConfiguraRst(strCadena)
                    Ans = ShowMultiReport(rst, "factura_elec", , App.Path + "\Reportes\")
                End If
End If

                
            Exit Sub


strCadena = "SELECT `id_venta`,`fecha_emision`,`fecha_vencimiento`,`doc_des`,`documento`,`id_cliente`,`ncliente`,`direccion`,`exonerado`,`valor_venta`,`igv`,`total`,`sunat_key`,`sunat_hash`,`id_producto`,`detalle`,`referencia`,`cantidad`,`valor_neto`,`precio`,`todatl_detalle`,'---','" & UCase(EnLetras_moneda(in_total, in_moneda)) & "',resolucion,in_guia,id_moneda,'-','-','-',`ruc` FROM view_factura_electronica WHERE id_venta='" & Val(in_venta) & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "factura_elec", , App.Path + "\Reportes\")


End Sub

Public Sub get_serie_comprobante(ByVal in_combo As DataCombo, ByVal in_doc As String, ByVal in_serie As String)

If KEY_COMPROBANTES_PROPIOS = "si" Then
    strCadena = "SELECT serie as Codigo,serie as Descripcion FROM almacen_comprobante WHERE id_doc='" & in_doc & "' and id_alm='" & KEY_VENTANILLA & "'   and ruc='" & KEY_RUC & "' ORDER BY serie ASC"
Else
    strCadena = "SELECT serie as Codigo,serie as Descripcion FROM almacen_comprobante WHERE id_doc='" & in_doc & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' ORDER BY serie ASC"
End If
Call ConfiguraRstL(strCadena)
Call LlenaDataComboL(in_combo)

If in_serie <> "" Then
  in_combo.BoundText = in_serie
End If


End Sub
Public Sub get_serie_comprobante_all(ByVal in_combo As DataCombo, ByVal in_doc As String)
On Error GoTo salir
strCadena = "SELECT serie as Codigo,serie as Descripcion FROM almacen_comprobante WHERE id_doc='" & in_doc & "'  and ruc='" & KEY_RUC & "' ORDER BY serie ASC"
Call ConfiguraRstL(strCadena)
Call LlenaDataComboL(in_combo)

Exit Sub
salir:


End Sub

Public Sub get_serie_comprobante_alm(ByVal in_combo As DataCombo, ByVal in_doc As String, Optional in_serie As String)

If in_serie <> "" Then
    in_parametro = " and serie='" & in_serie & "'"
Else
    in_parametro = ""
End If


If KEY_COMPROBANTES_PROPIOS = "si" Then
    strCadena = "SELECT serie as Codigo,serie as Descripcion FROM almacen_comprobante WHERE id_doc='" & in_doc & "' and id_alm='" & KEY_VENTANILLA & "' and ruc='" & KEY_RUC & "' " & in_parametro & "  ORDER BY serie ASC"
Else
    
    strCadena = "SELECT serie as Codigo,serie as Descripcion FROM almacen_comprobante WHERE id_doc='" & in_doc & "' and  id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' " & in_parametro & "  ORDER BY serie ASC"
End If
Call ConfiguraRstL(strCadena)
Call LlenaDataComboL(in_combo)

End Sub

Public Function get_comprobante_numero(ByVal in_doc As String, ByVal in_serie As String) As String
strCadena = "SELECT numero FROM almacen_comprobante WHERE id_doc='" & in_doc & "' and serie='" & in_serie & "' and ruc='" & KEY_RUC & "' limit 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_comprobante_numero = rstL("numero")
Else
   get_comprobante_numero = ""
End If
End Function

Public Function get_registro_ventas(ByVal in_mes As String, ByVal in_anio As String) As Double
get_registro_ventas = 0
strCadena = "SELECT sum(total) as total FROM movimiento_venta WHERE ruc='" & KEY_RUC & "' and id_mes='" & in_mes & "' and id_anio='" & in_anio & "' and id_doc IN('0003','0001','0007')"
Call ConfiguraRstK(strCadena)
get_registro_ventas = rstK(0)
End Function

Public Function get_verifica_doc() As Boolean

End Function

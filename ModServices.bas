Attribute VB_Name = "ModServices"
Public Sub crear_producto_keyfacil(ByVal in_url As String, ByVal in_metodo As String, ByVal in_json As String, Optional in_headers As String)
Dim Headers As String
Dim PostData() As Byte
Dim p As Object


Set p = JSON.parse("{method: '" & in_metodo & "', url: '" & in_url & "', json: true , body: " & in_json & ", headers: " & in_headers & " }")
PostData = JSON.toString(p)
PostData = StrConv(PostData, vbFromUnicode)
Headers = "Content-Type: application/json" & vbCrLf

WebBrowser1.Navigate2 "http://api.vitekey.net:3001/intranet/utils/api_console", 0, "", PostData, Headers


End Sub


Public Function get_importar_keyfacil(ByVal in_fecha_ini As String, ByVal in_fecha_fin As String) As Boolean

    Dim strHtml As String
    Set DomDoc = New XMLHTTP
    
    
    urlstr = "https://api.vitekey.com/keyfact/utils/reporte-ventas?password=vitekey2018&ruc=" & KEY_RUC & "&date_start=" & Format(in_fecha_ini, "MM/dd/YYYY") & "&date_end=" & Format(in_fecha_fin, "MM/dd/YYYY")
    
    Set DomDoc = New XMLHTTP
     'Parámetros en formato URLEncode
     'Metodo a usar, url, y true en caso de manejar la respuesta en modo asíncrono
     DomDoc.Open "GET", urlstr, False
     'encabezados
     DomDoc.setRequestHeader "Connection", "close"
     DomDoc.send params
     'La respuesta, en caso de existir, está en responseBody.
    'También puedes especificar responseXml si tu aplicación devolviese XML
     strHtml = StrConv(DomDoc.responseBody, vbUnicode)
     
     
     Call procesar_ventas_keyfacil(strHtml)

End Function
Public Function get_importar_productos_keyfacil() As Boolean

Dim strHtml As String
Set DomDoc = New XMLHTTP
     'urlstr = "https://api.vitekey.com/keyfact/utils/reporte-ventas?password=vitekey2018&ruc=" & KEY_RUC & "&date_start=" & Format(in_fecha_ini, "MM/dd/YYYY") & "&date_end=" & Format(in_fecha_fin, "MM/dd/YYYY")
     
     urlstr = "https://api.vitekey.com/keyfact/erp/products/list?api_key=fd235235-e97a-4db6-8f50-fa84145c3f5d&company_id=" & KEY_TOKEN_CLOUD & "&offset=5000"
     
     Set DomDoc = New XMLHTTP
     'Parámetros en formato URLEncode
     'Metodo a usar, url, y true en caso de manejar la respuesta en modo asíncrono
     DomDoc.Open "GET", urlstr, False
     'encabezados
     DomDoc.setRequestHeader "Connection", "close"
     DomDoc.send params
     'La respuesta, en caso de existir, está en responseBody.
    'También puedes especificar responseXml si tu aplicación devolviese XML
     strHtml = StrConv(DomDoc.responseBody, vbUnicode)
     Call procesar_importacion_keyfacil(strHtml)

End Function

Public Function get_importar_MITIENDA(ByVal in_codigo As String) As Boolean

Dim strHtml As String
Set DomDoc = New XMLHTTP
     urlstr = "https://api.mitienda.pe/v1/mitienda/order?code=" & in_codigo & ""
     
     urlstr = "https://api.vitekey.com/keyfact/erp/products/list?api_key=fd235235-e97a-4db6-8f50-fa84145c3f5d&company_id=" & KEY_TOKEN_CLOUD & "&offset=5000"
     
     Set DomDoc = New XMLHTTP
     'Parámetros en formato URLEncode
     'Metodo a usar, url, y true en caso de manejar la respuesta en modo asíncrono
     DomDoc.Open "GET", urlstr, False
     'encabezados
     DomDoc.setRequestHeader "Connection", "close"
     DomDoc.send params
     'La respuesta, en caso de existir, está en responseBody.
    'También puedes especificar responseXml si tu aplicación devolviese XML
     strHtml = StrConv(DomDoc.responseBody, vbUnicode)
     'Call procesar_importacion_keyfacil(strHtml)

End Function

Public Sub procesar_importacion_keyfacil(ByVal strHtml As String)
Dim in_imagen As String
Dim json_r As Object
Set json_r = JSON.parse(strHtml)

If json_r.Count >= 1 Then
   
For i = 1 To json_r.Count  ' recorro la cantidad de comprobantes
    in_producto = Format(json_r(i).Item("code"), "00000")
    in_imagen = ""
    If json_r(i).Item("photos_urls").Count > 0 Then
       in_imagen = json_r(i).Item("photos_urls")(1)
       strCadena = "UPDATE producto SET imagen='" & Trim(in_imagen) & "' WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
       CnBd.Execute (strCadena)
    End If
    
Next i
End If


End Sub

Public Sub procesar_ventas_keyfacil(ByVal strHtml As String)
Dim in_error As Boolean
Dim in_hash As String
Dim in_total_comprobante As Double
Dim in_emision As Date
Dim in_fecha As String
Dim in_alm As String
Dim estatus As String
Dim statuss As Boolean
Dim in_fecha_actual As String
Dim in_fecha_comprobante As Date
Dim in_numero() As String
Dim json_r As Object
Set json_r = JSON.parse(strHtml)

If json_r.Count >= 1 Then
   FrmRegistroVentas.prog_indicador.Min = 0
   FrmRegistroVentas.prog_indicador.Max = json_r.Count + 1
For i = 1 To json_r.Count  ' recorro la cantidad de comprobantes
    in_doc = Format(json_r(i).Item("type"), "0000")
    in_serie = Format(json_r(i).Item("serie"), "0000")
    in_nume = Format(json_r(i).Item("number"), "000000")
    in_dni = json_r(i).Item("client_docid")
    in_cliente = json_r(i).Item("client_name")
    
       in_doc_origen = ""
       in_serie_origen = ""
       in_numero_origen = ""
       in_tipo_nota = ""
       in_motivo = ""
       in_comprobante_rel = 0
       
    If in_doc = "0007" Then
       in_doc_origen = json_r(i).Item("note_info").Item("invoice_modified_type")
       in_serie_origen = json_r(i).Item("note_info").Item("invoice_modified_serie")
       in_numero_origen = json_r(i).Item("note_info").Item("invoice_modified_number")
       in_tipo_nota = "01"
       in_motivo = json_r(i).Item("note_info").Item("reason")
       strCadena = "call ADM_servicios_generales('22','" & Format(in_numero_origen, "000000") & "','','','" & Format(in_serie_origen, "0000") & "','" & Format(in_doc_origen, "0000") & "','" & KEY_RUC & "')"
       Call ConfiguraRstA(strCadena)
       If rstA.RecordCount > 0 Then
          in_comprobante_rel = rstA(0)
       Else
          in_comprobante_rel = 0
       End If
    End If
    
    
    If IsNull(json_r(i).Item("voided_document_id")) = True Then
       estatus = "no"
    Else
       estatus = "si"
    End If
    
    If get_existe_comprobante_detallado(in_doc, in_serie, in_nume) = True Then
       
     
       
       GoTo nnn
    End If
    
    in_emision = Mid(json_r(i).Item("issue_date"), 1, 10)
    strCadena = "SELECT DATE_SUB('" & json_r(i).Item("issue_date") & "', INTERVAL 5 HOUR) as emision "
    Call ConfiguraRstlocal(strCadena)
    in_fecha = Format(rstLocal("emision"), "YYYY-mm-dd")
    in_fecha_comprobante = Mid(in_fecha, 1, 10)
    
    
    If json_r(i).Item("currency") = "PEN" Then
        in_moneda = "00001"
    Else
        in_moneda = "00002"
    End If
    
    
    
    in_hash = json_r(i).Item("id")
    in_key = "-"
    
     in_alm = get_alm_codigo(in_doc, in_serie)
     
     strCadena = "call p_nueva_venta_v11('" & KEY_USUARIO & "','" & in_alm & "','" & KEY_RUC & "')"
     CnBd.Execute (strCadena)
       
    If in_dni = "" Then
      in_dni = "00000000"
    End If
    
    strCadena = "SELECT * FROM persona WHERE dni='" & in_dni & "'"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount < 1 Then
        strCadena = " call p_insert_persona_iii('" & Trim(in_dni) & "','-','-','-','" & Replace(Trim(in_cliente), "'", "") & "','" & KEY_DIR_PUBLIC & "','','-','no','no','no','no','no','no','si','" & KEY_DEPARTAMENTO & "','" & KEY_PROVINCIA & "','" & KEY_DISTRITO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
    End If
    
    
    strCadena = "SELECT cod_unico FROM entidad_empresa WHERE cod_unico='" & in_dni & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount < 1 Then
        strCadena = "INSERT INTO entidad_empresa (cod_unico,id_empresa,id_cliente)VALUES('" & in_dni & "','" & KEY_RUC & "','si')"
        CnBd.Execute (strCadena)
    End If
    
    
    
    
    
    
    in_direccion = get_direccion(in_dni)
    in_vendedor = get_vendedor_erp(in_dni)
   
    
    
    If get_existe_comprobante(in_doc, in_serie, in_nume) = False Then
Iniciar:
    in_documento = get_comprobante_sunat(in_doc) & ":" & in_serie & "-" & in_nume
    
    in_observacion = "-"
    IN_TOTAL_TOTAL = json_r(i).Item("computed").Item("total")
    in_descuento_global = json_r(i).Item("computed").Item("total_discount")
    
    If IN_TOTAL_TOTAL = 0 And in_descuento_global > 0 Then
        in_obsequio = "si"
    Else
        in_obsequio = "no"
    End If
    
    
    If KEY_CON_IGV = "si" Then
        in_valor_venta = IN_TOTAL_TOTAL / (1 + KEY_IGV)
    Else
        in_valor_venta = IN_TOTAL_TOTAL
    End If
    
    
    in_total_exonerado = json_r(i).Item("computed").Item("total_exonerated")
    in_total_igv = json_r(i).Item("computed").Item("total_igv")
    in_total_comprobante = 0
    For j = 1 To json_r(i).Item("items").Count ' recorro la cantidad de productos
        in_producto = get_codigo_producto(Format(json_r(i).Item("items")(j).Item("code"), "00000"))
        
        If in_producto = "0" Then
            in_producto = get_codigo_producto_descripcion(Replace(json_r(i).Item("items")(j).Item("description"), "'", ""))
        End If
        
        in_cantidad = json_r(i).Item("items")(j).Item("quantity")
        in_precio = json_r(i).Item("items")(j).Item("unit_price")
        in_detalle = Replace(json_r(i).Item("items")(j).Item("description"), "'", "")
        in_detalle = Replace(in_detalle, "'", "")
        If in_obsequio = "si" Then
            in_total = 0
        Else
            in_total = Val(in_precio) * Val(in_cantidad)
        End If
        
        in_total_comprobante = in_total_comprobante + in_total
        
        strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,servicio,obsequio,costo) VALUES " & _
        "('" & KEY_RUC & "','" & get_unidad_producto(in_producto) & "','" & in_dni & "','" & in_alm & "','" & in_doc & "','" & in_serie & "','" & in_nume & "','" & in_producto & "','" & Val(in_cantidad) & "'," & _
        "'" & Val(in_precio) & " ','" & in_total & "','0','" & KEY_CON_IGV & "','" & in_detalle & "','" & KEY_USUARIO & "','no','" & in_obsequio & "','" & get_precio_costo(in_producto) & "')"
        CnBd.Execute (strCadena)
    Next j
              in_descuento_global = 0
              If (Val(in_total_comprobante) - Val(IN_TOTAL_TOTAL)) > 0.1 Then
                in_descuento_global = (Val(in_total_comprobante) - Val(IN_TOTAL_TOTAL))
              End If
              
               If KEY_RUC = "20538939618" Or KEY_RUC = "20604059136" Or KEY_RUC = "20602072887" Or KEY_RUC = "20605615288" Then
                   in_forma_pago = "02"
                   in_cta_cobrar = KEY_CTA_COBRAR_SERVICIO
                   in_cta_ingreso = KEY_CTA_INGRESO_SERVICIO
                   If KEY_RUC = "20538939618" Or KEY_RUC = "20604059136" Then
                      nservicio = "00002"
                   Else
                      nservicio = "00001"
                   End If
                Else
                    in_forma_pago = "01"
                    in_cta_cobrar = KEY_CTA_COBRAR_PRODUCTO
                    in_cta_ingreso = KEY_CTA_INGRESO_PRODUCTO
                    nservicio = "00001"
                End If
                
              If (Val(in_total_comprobante) - Val(IN_TOTAL_TOTAL + in_descuento_global)) < 0.1 Then
                
                Call get_pago_keyfacil(in_doc, in_serie, in_nume, IN_TOTAL_TOTAL, in_moneda, in_forma_pago, in_alm)
                
          
                
                in_cambio_fecha = get_tipo_cambio_dia(Format(in_fecha_comprobante, "YYYY-mm-dd"), "valor_venta")
                
          
               
            
                
                
                
                strCadena = "call p_insert_venta_cabecera_v15('" & in_doc & "','" & in_alm & "','01','" & in_moneda & "','no'," & _
                "'" & in_serie & "','" & in_nume & "','" & in_dni & "','" & Replace(in_cliente, "'", "") & "','" & in_valor_venta & "','" & in_total_igv & "','" & in_total_exonerado & "','" & IN_TOTAL_TOTAL & "','0'," & _
                "'" & Val(in_total_exonerado) & "','0','" & Format(in_fecha_comprobante, "YYYY-mm-dd") & "','" & Format(in_fecha_comprobante, "YYYY-mm-dd") & "','" & nservicio & "','" & in_vendedor & "','" & KEY_USUARIO & "','" & Val(in_cambio_fecha) & "','no','" & formato_item(Month(in_fecha_comprobante), 2) & "','" & Year(in_fecha_comprobante) & "'" & _
                ",'" & in_documento & "',CURTIME(),'T','" & in_direccion & "','no','" & in_hash & "','" & in_key & "','" & in_tipo_nota & "','" & in_motivo & "',' ',' ','" & KEY_VENTANILLA & "','01','" & in_seguro & "','" & in_observacion & "','no','" & KEY_CONTABILIDAD & "','" & in_cta_cobrar & "','" & in_cta_ingreso & "','" & in_descuento_global & "','0','0','" & in_comprobante_rel & "','no','0','" & KEY_SIN_EFECTO_CAJA & "','no','" & KEY_RUC & "')"
                Call ConfiguraRstPP(strCadena)
                id_venta = rstPP("in_venta")
                
                strCadena = "SELECT * FROM movimiento_venta_monto_temporal WHERE id_doc='" & in_doc & "' and serie='" & in_serie & "' and id_usuario='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstK(strCadena)
                If rstK.RecordCount > 0 Then
                    in_glosa = "COBRO:" & in_documento
                    in_fecha_actual = Format(Mid(in_fecha, 1, 10), "YYYY-mm-dd")
                    If in_forma_pago = "01" Then
                        in_mis_cuentas_det = procesar_transaccion_venta(rstK("id_forma_pago"), KEY_ALM, get_cuenta_pago(rstK("id_forma_pago")), Format(in_fecha_comprobante, "YYYY-mm-dd"), "00001", in_dni, in_cliente, in_glosa, IN_TOTAL_TOTAL, "0", id_venta, "0", in_documento, KEY_CAMBIO, rstK("id_tarjeta_operacion"), "1CIX000000000174", "1CIX000000000078", KEY_USUARIO, in_doc, KEY_RUC)
                        Call put_realizar_pago(id_venta, id_venta, Abs(IN_TOTAL_TOTAL), in_doc, KEY_CAMBIO, Val(in_mis_cuentas_det))
                    End If
                End If
                       
                
                
                If estatus = "si" Then
                    Call AnularVentas(in_doc, in_serie, in_nume, in_alm)
               
                End If
                Call put_correlativo_venta(in_doc, in_serie, in_nume)
                
                
                
             Else
                If MsgBox("Ha Ocurrido una Inconsistencia con este Pedido" + Space(2) + in_documento + Chr(13) + "Desea Pasarlo de Nuevo.", vbInformation + vbYesNo) = vbYes Then
                    strCadena = "call p_nueva_venta_v11('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    GoTo Iniciar
                End If
             End If
       End If
nnn:
        DoEvents
        
        FrmRegistroVentas.cmdImportarDesdeKeyfacil.Caption = str(i) & Space(10) & str(json_r.Count) + Space(1) + Day(in_fecha_comprobante)
        FrmRegistroVentas.prog_indicador.Value = i
        DoEvents
Next i
End If


strCadena = "call p_nueva_venta_v11('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)


End Sub

Public Function get_vendedor_erp(ByVal in_ruc As String) As String
    strCadena = "SELECT id_vendedor FROM entidad_empresa WHERE cod_unico='" & in_ruc & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       get_vendedor_erp = rstT("id_vendedor")
    Else
       get_vendedor_erp = "00000000"
    End If
    
End Function

Public Function get_alm_codigo(ByVal in_doc As String, ByVal in_serie As String) As String

strCadena = "SELECT id_alm FROM almacen_comprobante WHERE id_doc='" & in_doc & "' and serie='" & in_serie & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       get_alm_codigo = rstT("id_alm")
    Else
       get_alm_codigo = "00001"
    End If
    

End Function
Public Function get_codigo_producto(ByVal in_producto As String) As String
strCadena = "SELECT id_producto FROM producto WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount < 1 Then
       get_codigo_producto = 0
       
    End If
End Function

Public Function get_codigo_producto_descripcion(ByVal in_producto As String) As String
strCadena = "SELECT id_producto FROM producto WHERE nombre_prod LIKE '%" & in_producto & "%' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount < 1 Then
       get_codigo_producto_descripcion = "00001"
    Else
       get_codigo_producto_descripcion = rstT("id_producto")
    End If
End Function



Public Sub put_temporal_ventas(ByVal Grilla As MSHFlexGrid, ByVal in_venta As String)
       
        
End Sub
Private Sub validar_correlativos(ByVal in_ruc As String)
strCadena = "SELECT DISTINCT id_doc,serie FROM movimiento_venta_verificacion WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & in_ruc & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   rstK.MoveFirst
   For i = 0 To rstK.RecordCount - 1
       
       strCadena = "SELECT * FROM movimiento_venta_verificacion WHERE id_doc='" & rstK("id_doc") & "' and serie='" & rstK("serie") & "' and dni_save='" & KEY_USUARIO & "' and ruc='" & in_ruc & "' ORDER BY numero ASC"
       Call ConfiguraRstA(strCadena)
       If rstA.RecordCount > 0 Then
            rstA.MoveFirst
            For j = 0 To rstA.RecordCount - 1
                    If j = 0 Then
                        in_correlativo = Val(rstA("numero"))
                    Else
                        If Val(rstA("numero")) <> in_correlativo Then
                             MsgBox "NUMERO FALTANTE:" + rstA("id_doc") + ":" + rstA("serie") + "-" + str(in_correlativo), vbInformation
                             in_correlativo = in_correlativo + 1
                        End If
                    End If
                    in_correlativo = in_correlativo + 1
                    rstA.MoveNext
            Next j
       End If
       rstK.MoveNext
   Next i
End If

End Sub

Public Sub verificar_ventas_keyfacil(ByVal strHtml As String, ByVal in_ruc As String)
Dim in_error As Boolean
Dim in_hash As String
Dim in_total_comprobante As Double
Dim in_emision As Date
Dim in_fecha As String
Dim in_alm As String
Dim estatus As String
Dim statuss As Boolean
Dim in_fecha_actual As String
Dim in_fecha_comprobante As Date
Dim in_numero() As String
Dim json_r As Object
Set json_r = JSON.parse(strHtml)


If json_r.Count >= 1 Then
   FrmPersona.Prg_Contabilidad.Min = 0
   FrmPersona.Prg_Contabilidad.Max = json_r.Count + 1
   
   strCadena = "call ADM_comprobantes_verificar('0','','','','','','" & KEY_USUARIO & "','" & in_ruc & "')"
   CnBd.Execute (strCadena)
    
For i = 1 To json_r.Count  ' recorro la cantidad de comprobantes
    
    in_doc = Format(json_r(i).Item("type"), "0000")
    in_serie = Format(json_r(i).Item("serie"), "0000")
    in_nume = Format(json_r(i).Item("number"), "000000")
    in_dni = json_r(i).Item("client_docid")
    in_cliente = json_r(i).Item("client_name")
    in_hash = json_r(i).Item("id")
     
    If IsNull(json_r(i).Item("voided_document_id")) = True Then
       estatus = "no"
    Else
       estatus = "si"
    End If
    IN_TOTAL_TOTAL = json_r(i).Item("computed").Item("total")
    
    strCadena = "call ADM_comprobantes_verificar('1','" & in_doc & "','" & in_serie & "','" & in_nume & "','" & estatus & "','" & in_hash & "','" & KEY_USUARIO & "','" & in_ruc & "')"
    CnBd.Execute (strCadena)
     
    If get_existe_comprobante_monto(in_doc, in_serie, in_nume, IN_TOTAL_TOTAL, estatus, in_ruc) = False Then
        If MsgBox("DESEA CONTINUAR", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    
    FrmPersona.cmdGenerarContabilidad.Caption = str(i) & Space(2) & str(json_r.Count)
    FrmPersona.Prg_Contabilidad.Value = i
    DoEvents
Next i

Call validar_correlativos(in_ruc)

MsgBox "VERIFICACION TERMINADA", vbInformation




End If


   

End Sub


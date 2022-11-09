Attribute VB_Name = "CrearXml"
Public Sub CreateXmlVentasCompras(ByVal in_periodo As String)


Dim Ruta_XML As String
Dim fso As Object

Ruta_XML = App.Path & "\comparar_percy\archivo.xml"

Set fso = New Scripting.FileSystemObject
Set fXML = fso.CreateTextFile(App.Path & "\comparar_percy\archivo.xml", True)

strCadena = "SELECT * FROM con_periodo WHERE id='" & in_periodo & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   in_mes = rst("mes")
   in_anio = rst("ejercicio")
   in_fecha_inicio = rst("FechaInicio")
   in_fecha_fin = rst("FechaFin")
End If


strCadena = "SELECT * FROM view_registro_ventas WHERE mes='" & in_mes & "' and anio='" & in_anio & "' and  ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then

fXML.WriteLine "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no"" ?>"
fXML.WriteLine "<iva>"
If rst.RecordCount > 0 Then
fXML.WriteLine "<TipoIDInformante>" & "R" & "</TipoIDInformante>"
fXML.WriteLine "<IdInformante>" & KEY_RUC & "</IdInformante>"
fXML.WriteLine "<razonSocial>" & Replace(Replace(KEY_EMPRESA, ".", ""), ",", "") & "</razonSocial>"
fXML.WriteLine "<Anio>" & in_anio & "</Anio>"
fXML.WriteLine "<Mes>" & Format(in_mes, "00") & "</Mes>"
fXML.WriteLine "<numEstabRuc>" & Format(1, "000") & "</numEstabRuc>"
fXML.WriteLine "<totalVentas>" & Format(rst("acumulado_inter"), "###0.00") & "</totalVentas>"
fXML.WriteLine "<codigoOperativo>" & "IVA" & "</codigoOperativo>"

End If
End If

'***** COMPRAS SEGUN FORMATO
strCadena = "SELECT * FROM movimiento_compra WHERE id_tipo_compra<>'01' and  id_periodo='" & in_periodo & "' and  ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   fXML.WriteLine "<compras>"
   For i = 0 To rst.RecordCount - 1
        fXML.WriteLine "<detalleCompras>"
                fXML.WriteLine "<codSustento>" & "01" & "</codSustento>"
                fXML.WriteLine "<tpIdProv>" & "01" & "</tpIdProv>"
                fXML.WriteLine "<idProv>" & rst("id_proveedor") & "</idProv>"
                fXML.WriteLine "<tipoComprobante>" & Format(rst("id_doc"), "00") & "</tipoComprobante>"
                fXML.WriteLine "<parteRel>" & "NO" & "</parteRel>"
                fXML.WriteLine "<fechaRegistro>" & Format(rst("fecha_emision"), "dd/mm/YYYY") & "</fechaRegistro>"
                fXML.WriteLine "<establecimiento>" & rst("id_almacen") & "</establecimiento>"
                fXML.WriteLine "<puntoEmision>" & Format(rst("id_alm"), "000") & "</puntoEmision>"
                fXML.WriteLine "<secuencial>" & Format(rst("numero"), "000000000") & "</secuencial>"
                fXML.WriteLine "<fechaEmision>" & Format(rst("fecha_emision"), "dd/mm/YYYY") & "</fechaEmision>"
                fXML.WriteLine "<autorizacion>" & rst("autorizacion") & "</autorizacion>"
                If rst("igv") = 0 Then
                    in_base_imponible = 0
                Else
                    in_base_imponible = rst("valor_venta")
                End If
                
                fXML.WriteLine "<baseNoGraIva>" & Format(in_base_imponible, "#,##0.00") & "</baseNoGraIva>"
                fXML.WriteLine "<baseImponible>" & Format(in_base_imponible, "#,##0.00") & "</baseImponible>"
                fXML.WriteLine "<baseImpGrav>" & Format(in_base_imponible, "#,##0.00") & "</baseImpGrav>"
                fXML.WriteLine "<baseImpExe>" & Format(rst("exonerado"), "#,##0.00") & "</baseImpExe>"
                in_monto_ice = 0#
                fXML.WriteLine "<montoIce>" & Format(in_monto_ice, "#,##0.00") & "</montoIce>"
                fXML.WriteLine "<montoIva>" & Format(rst("igv"), "#,##0.00") & "</montoIva>"
                fXML.WriteLine "<valRetBien10>" & Format(0, "#,##0.00") & "</valRetBien10>"
                fXML.WriteLine "<valRetServ20>" & Format(0, "#,##0.00") & "</valRetServ20>"
                fXML.WriteLine "<valorRetBienes>" & Format(0, "#,##0.00") & "</valorRetBienes>"
                fXML.WriteLine "<valRetServ50>" & Format(0, "#,##0.00") & "</valRetServ50>"
                fXML.WriteLine "<valorRetServicios>" & Format(0, "#,##0.00") & "</valorRetServicios>"
                fXML.WriteLine "<valRetServ100>" & Format(0, "#,##0.00") & "</valRetServ100>"
                fXML.WriteLine "<totbasesImpReemb>" & Format(0, "#,##0.00") & "</totbasesImpReemb>"
                fXML.WriteLine "<pagoExterior>"
                    fXML.WriteLine "<pagoLocExt>" & "01" & "</pagoLocExt>"
                    fXML.WriteLine "<paisEfecPago>" & "NA" & "</paisEfecPago>"
                    fXML.WriteLine "<aplicConvDobTrib>" & "NA" & "</aplicConvDobTrib>"
                    fXML.WriteLine "<pagExtSujRetNorLeg>" & "NA" & "</pagExtSujRetNorLeg>"
                fXML.WriteLine "</pagoExterior>"
                fXML.WriteLine "<air>"
                     fXML.WriteLine "<detalleAir>"
                     fXML.WriteLine "<codRetAir>" & "332" & "</codRetAir>"
                     fXML.WriteLine "<baseImpAir>" & Format(in_base_imponible, "#,##0.00") & "</baseImpAir>"
                     fXML.WriteLine "<porcentajeAir>" & Format(0, "#,##0.00") & "</porcentajeAir>"
                     fXML.WriteLine "<valRetAir>" & Format(0, "#,##0.00") & "</valRetAir>"
                     fXML.WriteLine "</detalleAir>"
                fXML.WriteLine "</air>"
            fXML.WriteLine "</detalleCompras>"
            
            rst.MoveNext
   Next i
   fXML.WriteLine "</compras>"
End If
'******COMPRAS FIN



'***** VENTAS SEGUN FORMATO
in_total = 0
in_iva = 0
strCadena = "SELECT * FROM movimiento_venta WHERE anulado='no' and  fecha_emision>='" & Format(in_fecha_inicio, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(in_fecha_fin, "YYYY-mm-dd") & "' and id_doc In('0001') and  ruc='" & KEY_RUC & "' LIMIT 0"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst

   fXML.WriteLine "<ventas>"
   For i = 0 To rst.RecordCount - 1
        fXML.WriteLine "<detalleVentas>"
                If Right(rst("id_cliente"), 3) = "001" Then
                    in_idcliente = "04"
                Else
                    in_idcliente = "05"
                End If
                fXML.WriteLine "<tpIdCliente>" & in_idcliente & "</tpIdCliente>"
                fXML.WriteLine "<idCliente>" & rst("id_cliente") & "</idCliente>"
                fXML.WriteLine "<parteRelVtas>" & "NO" & "</parteRelVtas>"
                fXML.WriteLine "<tipoComprobante>" & Format(rst("id_doc"), "00") & "</tipoComprobante>"
                fXML.WriteLine "<tipoEmision>" & "F" & "</tipoEmision>"
                fXML.WriteLine "<numeroComprobantes>" & "1" & "</numeroComprobantes>"
                If rst("igv") = 0 Then
                    in_base_imponible = 0
                Else
                    in_base_imponible = rst("valor_venta")
                End If
                fXML.WriteLine "<baseNoGraIva>" & Format(in_base_imponible, "#,##0.00") & "</baseNoGraIva>"
                fXML.WriteLine "<baseImponible>" & Format(in_base_imponible, "#,##0.00") & "</baseImponible>"
                fXML.WriteLine "<baseImpGrav>" & Format(in_base_imponible, "#,##0.00") & "</baseImpGrav>"
                fXML.WriteLine "<montoIva>" & Format(rst("igv"), "#,##0.00") & "</montoIva>"
                fXML.WriteLine "<montoIce>" & Format(rst("igv"), "#,##0.00") & "</montoIce>"
                
                fXML.WriteLine "<valorRetIva>" & Format(0, "#,##0.00") & "</valorRetIva>"
                fXML.WriteLine "<valorRetRenta>" & Format(0, "#,##0.00") & "</valorRetRenta>"
                fXML.WriteLine "<formasDePago>"
                    fXML.WriteLine "<formaPago>" & "01" & "</formaPago>"
                fXML.WriteLine "</formasDePago>"
                
                
                
                
            fXML.WriteLine "</detalleVentas>"
            in_total = in_total + rst("total")
            in_iva = in_iva + rst("igv")
            
            rst.MoveNext
   Next i
   
   fXML.WriteLine "</ventas>"
   
   
   fXML.WriteLine "<ventasEstablecimiento>"
                    fXML.WriteLine "<ventaEst>"
                         fXML.WriteLine "<codEstab>"
                         fXML.WriteLine "<ventasEstab>" & Format(in_total, "#,##0.00") & "</ventasEstab>"
                         fXML.WriteLine "<ivaComp>" & Format(in_iva, "#,##0.00") & "</ivaComp>"
                    fXML.WriteLine "</ventaEst>"
   fXML.WriteLine "</ventasEstablecimiento>"
   
   
 
   
End If
'******COMPRAS FIN

fXML.WriteLine "</iva>"
  
    
   
    
End Sub

Attribute VB_Name = "modSeguridad"
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Function eliminar_informe_I(ByVal in_forme As Double) As Boolean
eliminar_informe_I = False
strCadena = "DELETE FROM proyecto_informe WHERE id_informe='" & in_forme & "'"
CnBd.Execute (strCadena)
eliminar_informe_I = True
End Function
Public Function get_ubigueo(ByVal in_departamento As String, ByVal in_provincia As String, ByVal in_distrito As String) As String
strCadena = "SELECT ubigueo FROM view_ubigeo WHERE id_depa='" & in_departamento & "' and id_provincia='" & in_provincia & "' and id_distrito='" & in_distrito & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    get_ubigueo = rstL("ubigueo")
Else
    get_ubigueo = ""

End If
End Function
Public Function get_pago_compra(ByVal in_compra As Double, ByVal fecha_ini As Date, fecha_fin As Date) As Single

Dim in_pago As Double

strCadena = "SELECT sum(monto_pagado) FROM view_comprobante_pago WHERE id_movimiento='" & Val(in_compra) & "' and fecha_emision>='" & Format(fecha_ini, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(fecha_fin, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRstZ(strCadena)
If IsNull(rstZ(0)) = True Then
    get_pago_compra = 0
Else
    get_pago_compra = rstZ(0)
End If





End Function

Public Function put_eliminar_plan(ByVal in_plan As String) As Boolean
On Error GoTo salir
strCadena = "DELETE FROM plan_servicio WHERE id_plan='" & Val(in_plan) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
put_eliminar_plan = True
Exit Function

salir:
put_eliminar_plan = False

End Function



Public Function anular_memorandum(ByVal in_memo As String)
On Error GoTo salir
strCadena = "UPDATE memorandun SET anulado='si' WHERE id_memo='" & Val(in_memo) & "'"
CnBd.Execute (strCadena)

strCadena = "call CON_InsertaAsiento_Memorandum_Extorno('" & Val(in_memo) & "')"
CnBd.Execute (strCadena)
MsgBox "Anulado Correctamente", vbInformation
salir:

End Function

Public Function get_forma_pago_contado() As Integer
strCadena = "SELECT * FROM forma_pago_detalle WHERE ruc='" & KEY_RUC & "' and id='01' and id_detalle='01' LIMIT 1"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   get_forma_pago_contado = rstP("id_registro")
Else
   get_forma_pago_contado = 0
End If
End Function

Public Function get_forma_pago_contado_keyfacil(ByVal in_moneda As String, ByVal in_alm As String, ByVal forma_pago As String) As Integer

If forma_pago = "01" Then
    strCadena = "SELECT * FROM forma_pago_detalle WHERE ruc='" & KEY_RUC & "' and id='01' and id_detalle='01' and id_moneda='" & in_moneda & "' and id_alm='" & in_alm & "' LIMIT 1"
Else
    strCadena = "SELECT * FROM forma_pago_detalle WHERE ruc='" & KEY_RUC & "' and id='02' and id_detalle='08' and id_moneda='" & in_moneda & "' and id_alm='" & in_alm & "' LIMIT 1"
End If


Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   get_forma_pago_contado_keyfacil = rstP("id_registro")
Else
   get_forma_pago_contado_keyfacil = 0
End If
End Function

Public Sub anular_manifiesto(ByVal in_manifiesto As String)



End Sub
Public Function anular_pedido(ByVal in_pedido As String) As Boolean
anular_pedido = False
strCadena = "UPDATE movimiento_pedido SET id_estado='00004' WHERE id_pedido='" & Val(in_pedido) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
anular_pedido = True
MsgBox "Pedido Anulado con Exito.", vbInformation

End Function

Public Function anular_orden_compra(ByVal in_orden As String) As Boolean

Dim in_guia As String
Dim in_flete As String
Dim in_factura As String




'****VERIFICACION SI HAY VARIAS RECEPCIONES
strCadena = "SELECT * FROM orden_compra WHERE id_orden='" & Val(in_orden) & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   in_orden_compra = rstL("id_recepcion")
    in_compra_flete = rstL("id_factura_flete")
    in_serie_guia = rstL("guia_serie")
    in_numero_guia = rstL("guia_numero")
    in_factura_unica = rstL("id_compra")
    
End If

strCadena = "SELECT * FROM orden_compra WHERE id_recepcion='" & Val(in_orden_compra) & "' and id_estado<>'3' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount = 1 Then

strCadena = "SELECT * FROM orden_compra WHERE id_orden='" & Val(in_orden) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    in_compra_flete = rstL("id_factura_flete")
    in_serie_guia = rstL("guia_serie")
    in_numero_guia = rstL("guia_numero")
    in_compra = rstL("id_compra")
            
    If MsgBox("Esta Anulacion tendra Efecto Contable. " + Chr(13) + "Desea Continuar ?", vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then
            in_compra_flete = rstL("id_factura_flete")
            in_serie_guia = rstL("guia_serie")
            in_numero_guia = rstL("guia_numero")
            strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & rstL("id_compra") & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRstZ(strCadena)
            If rstZ.RecordCount > 0 Then
               in_serie_compra = rstZ("serie")
               in_numero_compra = rstZ("numero")
               Call eliminar_compras_general(rstZ("id_doc"), rstZ("serie"), rstZ("numero"), rstZ("id_proveedor"))
            End If
            
            
            
            
            
            strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & in_compra_flete & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRstZ(strCadena)
            If rstZ.RecordCount > 0 Then
               Call eliminar_compras_general(rstZ("id_doc"), rstZ("serie"), rstZ("numero"), rstZ("id_proveedor"))
            End If
            
            strCadena = "call CON_InsertaAsiento_Recepcion_Extorno('" & Val(in_orden) & "')"
            CnBd.Execute (strCadena)
            
            
            
            If in_serie_guia = "" And in_numero_guia = "" Then
                Call put_delete_kardex_recepcion(in_orden, in_serie_compra, in_numero_compra)
            Else
                Call put_delete_kardex_recepcion(in_orden, in_serie_guia, in_numero_guia)
            End If
End If
Else
            strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & in_compra_flete & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRstZ(strCadena)
            If rstZ.RecordCount > 0 Then
               Call eliminar_compras_general(rstZ("id_doc"), rstZ("serie"), rstZ("numero"), rstZ("id_proveedor"))
            End If
            Call put_delete_kardex_recepcion(in_orden, in_serie_guia, in_numero_guia)

End If
    Else
           strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & in_compra_flete & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRstZ(strCadena)
            If rstZ.RecordCount > 0 Then
               Call eliminar_compras_general(rstZ("id_doc"), rstZ("serie"), rstZ("numero"), rstZ("id_proveedor"))
            End If
            Call put_delete_kardex_recepcion(in_orden, in_serie_guia, in_numero_guia)
            strCadena = "call CON_InsertaAsiento_Recepcion_Extorno('" & Val(in_orden) & "')"
            CnBd.Execute (strCadena)
    
    End If
    


strCadena = "UPDATE orden_compra SET id_estado='3' WHERE id_orden='" & Val(in_orden) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
anular_orden_compra = True








End Function


Public Function verificar_password(ByVal in_password As String) As Boolean
strCadena = "SELECT cod_unico FROM entidad_empresa WHERE cod_unico='" & KEY_USUARIO & "' and passwordaccesso='" & in_password & "' and id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    verificar_password = True
Else
    verificar_password = False
End If
End Function
Public Function get_licencia(ByVal in_dni As String) As String
strCadena = "SELECT * FROM persona WHERE dni='" & in_dni & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    get_licencia = rstK("licencia")
Else
    get_licencia = "-"
End If
End Function
Public Function verificar_password_admin(ByVal in_password As String) As String
strCadena = "SELECT id_cargo FROM entidad_empresa WHERE  passwordaccesso='" & in_password & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    verificar_password_admin = rst("id_cargo")
Else
    verificar_password_admin = "0"
End If
End Function

Public Sub put_biblio(ByVal in_dni As String, ByVal in_facultad As String, ByVal in_tipo_acceso As String, ByVal in_ciclo As String)
strCadena = "UPDATE entidad_empresa SET id_tipo_acceso='" & in_tipo_acceso & "',id_facultad='" & in_facultad & "',id_ciclo='" & in_ciclo & "' WHERE cod_unico='" & Trim(in_dni) & "' and id_empresa='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
End Sub

Public Function delete_formapago(ByVal in_forma_pago As String) As Boolean

strCadena = "SELECT count(*) FROM movimiento_venta_monto WHERE id_forma_pago='" & Trim(in_forma_pago) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst(0) = 0 Then
   strCadena = "DELETE FROM forma_pago_detalle WHERE id_registro='" & Trim(in_forma_pago) & "' and ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
   delete_formapago = True
Else
   delete_formapago = False
End If

End Function

Public Sub llenar_direccion(ByVal Grilla As MSHFlexGrid, ByVal in_dni As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT * FROM persona_direccion WHERE dni='" & in_dni & "'"
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 500
            Grilla.ColWidth(2) = 6000
            Grilla.ColWidth(3) = 500
        Next
        cabecera = "COD" & vbTab & "COD" & vbTab & "DIRECCION" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 1 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_direccion") & vbTab & Format(i + 1, "0000") & vbTab & rst("direccion") & vbTab & Chr(168)
            Grilla.AddItem Fila
           
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 3 '  .. en la columna
                            ' cambia la fuente para esta celda
                            
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            ' edita la celda
                            ' If rstT("estado") = "no" Then
                             '   estado = Chr(168)
                            'Else
                             '   estado = Chr(254)
                            'End If
                            
                        End With
      
            rst.MoveNext
            
        Next i
      
      
    
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Public Sub update_temporal(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_cliente As String, ByVal obj_serie As DataCombo, ByVal obj_numero As TextBox)
strCadena = "UPDATE temporal_ventas SET id_doc='" & in_doc & "',id_serie='" & in_serie & "',numero='" & Format(in_numero, "000000") & "' WHERE id_dni='" & in_cliente & "' and  dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' "
CnBd.Execute (strCadena)

strCadena = "UPDATE movimiento_venta_monto_temporal SET id_doc='" & in_doc & "',serie='" & in_serie & "',numero='" & Format(in_numero, "000000") & "' WHERE id_usuario='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

obj_serie.BoundText = in_serie
obj_numero.Text = Format(in_numero, "000000")
End Sub

Public Function insertar_transferencia_nula(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_almacen As String)

strCadena = "INSERT INTO movimiento_transferencia(id_doc,dni_atencion,atencion,id_tipo_guia,serie,numero,fecha,direccion,id_destinatario,destinatario,id_transporte,marca_placa,id_chofer,id_alm_origen,id_alm_destino,id_motivo,motivo_otros,observacion,id_venta,dni_save,ruc) " & _
        "VALUES('" & in_doc & "',' ',' ','01','" & in_serie & "','" & in_numero & "','" & KEY_FECHA & "','-','-','ANULADO'," & _
        "'-','-','-','" & in_almacen & "','0','01','-','-','0','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)

strCadena = "SELECT id_transferencia FROM movimiento_transferencia WHERE ruc='" & KEY_RUC & "'  ORDER BY id_transferencia DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   strCadena = "UPDATE movimiento_transferencia SET anulado='si' WHERE id_transferencia='" & rst("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
End If
End Function

Public Function anular_guia(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_alm As String)
        Dim in_almacen As String
        Dim in_almacen_destino As String
        Dim in_fecha As Date
        
        strCadena = "SELECT id_transferencia,id_alm_origen as id_alm,id_alm_destino,id_venta,id_motivo,serie,id_doc,fecha FROM movimiento_transferencia WHERE id_doc='" & in_doc & "' AND serie='" & in_serie & "' AND numero='" & in_numero & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
            anular_guia = True
            
            in_almacen = rstZ("id_alm")
            in_almacen_destino = rstZ("id_alm_destino")
            in_fecha = rstZ("fecha")
            
            strCadena = "UPDATE movimiento_transferencia SET anulado='si',id_estado='3' WHERE id_transferencia='" & rstZ("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
            If rstZ("id_motivo") = 3 Then
               strCadena = "DELETE FROM kardex WHERE id_alm='" & rstZ("id_alm") & "' and id_serie='" & rstZ("serie") & "' and id_doc='0009' and   id_movimiento='" & rstZ("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
               CnBd.Execute (strCadena)
               strCadena = "DELETE FROM kardex WHERE id_alm='" & rstZ("id_alm_destino") & "' and id_serie='" & rstZ("serie") & "' and id_doc='0009' and   id_movimiento='" & rstZ("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
               CnBd.Execute (strCadena)
            End If
            
            
            
            If get_diferida(rstZ("id_venta")) = "si" Then
                    strCadena = "DELETE FROM kardex WHERE id_tipo_movimiento='03' and  id_movimiento='" & rstZ("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "call CON_InsertaAsiento_GuiaDiferida_Extorno('" & rstZ("id_transferencia") & "')"
                    CnBd.Execute (strCadena)
            End If
            
            
            
            
            strCadena = "SELECT chasis,id_producto FROM movimiento_transferencia_series WHERE id_transferencia='" & rstZ("id_transferencia") & "'"
            Call ConfiguraRstL(strCadena)
            If rstL.RecordCount > 0 Then
                For p = 0 To rstL.RecordCount - 1
                    
                     ':::::::::::ACTUALIZA KARDEX
                     If KEY_RUC = "20128836251" Then
                            Call update_kardex_Vargas_modulo_compra(rstL("id_producto"), Format(in_fecha, "YYYY-mm-dd"))
                        Else
                            Call update_kardex_update(rstL("id_producto"), Format(in_fecha, "YYYY-mm-dd"))
                        End If
                     '--------------FIN KARDEX
                     
                    strCadena = "UPDATE imp_producto_detalle SET transferencia='no' WHERE id_producto='" & rstL("id_producto") & "' and  nro_chasis='" & rstL("chasis") & "' and id_alm='" & rstZ("id_alm") & "' and ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "UPDATE imp_producto_detalle SET ruc='0' WHERE id_producto='" & rstL("id_producto") & "' and  nro_chasis='" & rstL("chasis") & "' and id_alm='" & rstZ("id_alm_destino") & "' and ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "SELECT sum(cantidad_real),sum(cantidad_contable),sum(cantidad_pendiente) FROM kardex where id_alm='" & rstZ("id_alm") & "' and  id_producto='" & rstL("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstK(strCadena)
                    
                    strCadena = "UPDATE almacen_producto SET stock = '" & rstK(0) & "',stock_contable= '" & rstK(2) & "',stock_factura='" & rstK(1) & "'  WHERE ruc = '" & KEY_RUC & "' and id_alm ='" & rstZ("id_alm") & "' and id_producto = '" & rstL("id_producto") & "'"
                    CnBd.Execute (strCadena)
                                        
                    strCadena = "SELECT sum(cantidad_real),sum(cantidad_contable),sum(cantidad_pendiente) FROM kardex where id_alm='" & rstZ("id_alm_destino") & "' and  id_producto='" & rstL("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstK(strCadena)
                    
                    strCadena = "update almacen_producto set stock = '" & rstK(0) & "'  , `stock_contable` = '" & rstK(2) & "',stock_factura='" & rstK(1) & "'  where ruc = '" & KEY_RUC & "' and id_alm ='" & rstZ("id_alm_destino") & "' and id_producto = '" & rstL("id_producto") & "'"
                    CnBd.Execute (strCadena)
                    rstL.MoveNext
                Next p
            Else
                strCadena = "SELECT * FROM movimiento_transferencia_detalle WHERE id_transferencia='" & rstZ("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
                Call ConfiguraRstIN(strCadena)
                If rstIN.RecordCount > 0 Then
                   rstIN.MoveFirst
                   For i = 0 To rstIN.RecordCount - 1
                        in_producto = rstIN("id_producto")
                     ':::::::::::ACTUALIZA KARDEX
                     If KEY_RUC = "20128836251" Then
                            Call update_kardex_Vargas_modulo_compra(in_producto, Format(in_fecha, "YYYY-mm-dd"))
                        Else
                            Call update_kardex_update(in_producto, Format(in_fecha, "YYYY-mm-dd"))
                    End If
                     '--------------FIN KARDEX
                     
                    
                    
                    'ALMACEN ORIGEN ****
                    strCadena = "SELECT sum(cantidad_real),sum(cantidad_contable),sum(cantidad_pendiente) FROM kardex WHERE id_alm='" & in_almacen & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstK(strCadena)
                    
                    strCadena = "UPDATE almacen_producto SET stock = '" & rstK(0) & "',stock_contable= '" & rstK(2) & "',stock_factura='" & rstK(2) & "'  WHERE ruc = '" & KEY_RUC & "' and id_alm ='" & in_almacen & "' and id_producto = '" & in_producto & "'"
                    CnBd.Execute (strCadena)
                    
                    
                    strCadena = "SELECT saldo_stock FROM kardex WHERE id_alm='" & in_almacen & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC,id_kardex DESC LIMIT 1 "
                    Call ConfiguraRstK(strCadena)
                    If rstK.RecordCount > 0 Then
                    
                        in_saldo_stock = rstK("saldo_stock")
                    Else
                        in_saldo_stock = 0
                    End If
                    
                    
                    strCadena = "UPDATE almacen_producto SET stock = '" & in_saldo_stock & "'  WHERE ruc = '" & KEY_RUC & "' and id_alm ='" & in_almacen & "' and id_producto = '" & in_producto & "'"
                    CnBd.Execute (strCadena)
                    
                    
                    
                    
                                        
                    
                    
                    strCadena = "SELECT sum(cantidad_real),sum(cantidad_contable),sum(cantidad_pendiente) FROM kardex where id_alm='" & in_almacen_destino & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstK(strCadena)
                    
                    strCadena = "update almacen_producto set stock = '" & rstK(0) & "'  , `stock_contable` = '" & rstK(2) & "',stock_factura='" & rstK(2) & "'  where ruc = '" & KEY_RUC & "' and id_alm ='" & in_almacen_destino & "' and id_producto = '" & in_producto & "'"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "SELECT saldo_stock FROM kardex WHERE id_alm='" & in_almacen_destino & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC,id_kardex DESC LIMIT 1 "
                    Call ConfiguraRstK(strCadena)
                    If rstK.RecordCount > 0 Then
                    
                        in_saldo_stock = rstK("saldo_stock")
                    Else
                        in_saldo_stock = 0
                    End If
                    
                    strCadena = "UPDATE almacen_producto SET stock = '" & in_saldo_stock & "'  WHERE ruc = '" & KEY_RUC & "' and id_alm ='" & in_almacen_destino & "' and id_producto = '" & in_producto & "'"
                    CnBd.Execute (strCadena)
                    
                    
                    
                    
                    rstIN.MoveNext
                    
                   Next i
                End If
                
            End If
        Else
            If MsgBox("ESTA GUIA REMISION NO ESTA REGISTRADA" + Chr(13) + Chr(13) + "DESEA SEGUIR ANULANDO.", vbQuestion + vbYesNo) = vbYes Then
                
                Call insertar_transferencia_nula(in_doc, in_serie, in_numero, in_alm)
            End If
            
        End If
End Function

Public Function reservar_comprobante(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_alm As String) As String
    strCadena = "SELECT nombre_completo FROM view_movimeinto_venta_reserva WHERE in_doc='" & in_doc & "' and in_serie='" & in_serie & "' and in_numero='" & in_numero & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
           reservar_comprobante = True
           MsgBox "COMPROBANTE ESTA SIENDO UTILIZADO      " + Chr(13) + Chr(13) + "POR: " + UCase(rst("nombre_completo")) + Chr(13) + "USAMOS EL CORRELATIVO ?", vbInformation + vbYesNo, KEY_EMPRESA
           reservar_comprobante = Format(Val(in_numero) + 1, "000000")
           Exit Function
        
        
    Else
        strCadena = "INSERT INTO movimiento_venta_reserva(`in_doc`,`in_serie`,`in_numero`,`in_alm`,`dni_save`,`ruc`)VALUES('" & in_doc & "','" & in_serie & "','" & in_numero & "','" & in_alm & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        reservar_comprobante = False
    End If

End Function



Public Function eliminar_reservar(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String)
    strCadena = "DELETE FROM movimiento_venta_reserva WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
End Function
Public Function verificar_duplicado(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String) As Boolean
strCadena = "SELECT id_venta FROM movimiento_venta WHERE id_doc='" & in_doc & "' and serie='" & in_serie & "' and numero='" & in_numero & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
    verificar_duplicado = True
Else
    verificar_duplicado = False
End If
End Function
Public Function get_electronico(ByVal in_doc As String, ByVal in_serie As String) As String

strCadena = "SELECT electronico FROM almacen_comprobante WHERE  id_doc='" & in_doc & "' and serie='" & in_serie & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
   get_electronico = rstZ("electronico")
Else
   get_electronico = "no"
End If

End Function
Public Function get_ndocumento(ByVal in_doc As String) As String
get_ndocumento = "-"
Select Case id_doc
    Case "0003"
        get_ndocumento = "BOLETA:"
    Case "0001"
        
        get_ndocumento = "FACTURA:"
        
End Select


End Function

Public Function get_numero_comprobante(ByVal in_doc As String, ByVal in_serie As String) As String
strCadena = "SELECT numero FROM movimiento_venta WHERE id_doc='" & in_doc & "' and  serie='" & in_serie & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
    get_numero_comprobante = Format(Val(rstAux("numero")) + 1, "000000")
Else
    If KEY_FACTURACION_ELECTRONICA = "si" Then
        get_numero_comprobante = get_numero_doc(in_doc, in_serie)
    Else
        get_numero_comprobante = get_numero_doc(in_doc, in_serie)
    End If
End If
End Function
Public Function get_numero_doc(ByVal in_doc As String, ByVal in_serie As String) As String
strCadena = "SELECT numero FROM almacen_comprobante WHERE id_doc='" & in_doc & "' and serie='" & in_serie & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
   get_numero_doc = Format(rstAux("numero"), "000000")
Else
   get_numero_doc = "0"
End If
End Function

Public Function get_comprobante(ByVal in_venta As String) As String

strCadena = "SELECT fecha_emision,documento FROM movimiento_venta WHERE id_venta='" & in_venta & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
    get_comprobante = rstAux("documento") & ":" & Format(rstAux("fecha_emision"), "dd-mm-YYYY")
Else
    get_comprobante = "--:--:--"
End If

End Function

Public Function get_guia_numero(ByVal in_guia As String) As String
strCadena = "SELECT * FROM movimiento_transferencia WHERE id_transferencia='" & Val(in_guia) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
   get_guia_numero = "GUIA :" & rstIN("serie") & "-" & rstIN("numero") & Space(10) & Format(rstIN("fecha"), "dd-mm-YYYY")
Else
   get_guia_numero = "-"
End If
End Function
Public Function get_comprobante_orden_salida(ByVal in_venta As String, ByVal in_tipo As String) As String

If in_tipo = "03" Then
    strCadena = "SELECT id_transferencia FROM movimiento_transferencia WHERE id_transferencia='" & in_venta & "' and ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT fecha_emision,documento,id_guia FROM movimiento_venta WHERE id_venta='" & in_venta & "' and ruc='" & KEY_RUC & "'"
End If

Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
    If in_tipo = "03" Then
        get_comprobante_orden_salida = get_guia_numero(rstAux("id_transferencia"))
    Else
        get_comprobante_orden_salida = rstAux("documento") & ":" & Format(rstAux("fecha_emision"), "dd-mm-YYYY")
    End If
        
   
    
    
Else
    get_comprobante_orden_salida = "--:--:--"
End If

End Function
Public Function get_comprobante_venta(ByVal in_venta As String) As String
strCadena = "SELECT fecha_emision,documento FROM movimiento_venta WHERE id_venta='" & in_venta & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
    get_comprobante_venta = rstAux("documento")
Else
    get_comprobante_venta = "--:--:--"
End If
End Function
Public Sub disabled_form(ByVal Formulario As Form)
        Formulario.Enabled = False
       
End Sub
Public Sub enabled_form(ByVal Formulario As Form)
        Formulario.Enabled = True
End Sub
Public Sub save_boceto(ByVal in_imagen As String, ByVal in_ruta_web As String, ByVal in_ruta_local As String, ByVal in_backlog As Integer, ByVal in_observacion As String)
    strCadena = "INSERT INTO proyecto_backlog_img(`id_backlog`,`imagen`,in_ruta_web,in_ruta_local,observacion)values('" & in_backlog & "','" & in_imagen & "','" & in_ruta_web & "','" & in_ruta_local & "','" & in_observacion & "')"
    CnBd.Execute (strCadena)
End Sub
Public Function ExisteArchivo(ByVal Archivo As String) As Boolean
Dim Nombre As String ' Temporal para la búsqueda del archivo dado.
ExisteArchivo = False ' Supone que no existe.

    Nombre = Dir$(Archivo)
    If Len(Nombre) > 0 Then
        ExisteArchivo = True
    End If

End Function
Public Function DownloadFile(Url As String, LocalFilename As String) As Boolean

Dim lngRetVal As Long
lngRetVal = URLDownloadToFile(0, Url, LocalFilename, 0, 0)
If lngRetVal = 0 Then DownloadFile = True
End Function
Public Sub get_boceto(ByVal img_boceto As Image, ByVal ruta_web As String, ByVal in_ruta_local As String, ByVal in_imagen As String)
On Error GoTo salir
Dim str_ruta_img As String
Dim in_img_server As String
str_ruta_img = App.Path & "\archivos\" & KEY_RUC & "\download\" & Trim(in_imagen)
in_img_server = ruta_web
If Len(in_imagen) > 5 Then
      If ExisteArchivo(Trim(str_ruta_img)) = True Then
           img_boceto = LoadPicture(str_ruta_img)
      Else
           DownloadFile ruta_web, str_ruta_img
           img_boceto = LoadPicture(str_ruta_img)
      End If

    
End If
Exit Sub
salir:
End Sub
Public Sub get_bocetoO(ByVal img_boceto As Image, ByVal ruta_web As String, ByVal in_ruta_local As String, ByVal in_imagen As String)
On Error GoTo salir
Dim str_ruta_img As String
Dim in_img_server As String
str_ruta_img = App.Path & "\archivos\" & KEY_RUC & "\download\" & Trim(in_imagen)
in_img_server = ruta_web
If Len(in_imagen) > 5 Then
      If ExisteArchivo(Trim(str_ruta_img)) = True Then
           img_boceto = LoadPicture(str_ruta_img)
      Else
           DownloadFile ruta_web, str_ruta_img
           img_boceto = LoadPicture(str_ruta_img)
      End If

    
End If
Exit Sub
salir:
End Sub

Public Function get_pendientes() As Integer
Dim in_cantidad As Integer
strCadena = "SELECT  funct_backlog()"
Call ConfiguraRstP(strCadena)
get_pendientes = rstP(0)
in_cantidad = rstP(0)
If in_cantidad > 0 Then
    PlaySound App.Path & "\sonidos\dingding.wav"
End If

End Function
Public Function get_dni_reniec(ByVal in_dni As String) As Boolean

Dim Nombre As String
    On Error GoTo salir_cancel


    Dim strHtml As String
    'UrlStr = "https://soluciones.equifax.com.pe/e-commerce/flujo/validarPadron.htm?_HDIV_STATE_=33-1-43C78CD3FE0C2B14246C904C232C7E11"
     urlstr = "http://facturacion.vitekey.com/api/utiles/consultar_dni?dni=" & in_dni
     Set DomDoc = New XMLHTTP
     'Parámetros en formato URLEncode
    ' ventanilla = Split(rstaux("ventanilla"), "-")
     
     params = ""
     '
     'params = "tipoDocumento=" & 1 & "&nroDocumento=" & in_dni
     'Metodo a usar, url, y true en caso de manejar la respuesta en modo asíncrono
     DomDoc.Open "POST", urlstr, False
     'encabezados
     DomDoc.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
     DomDoc.setRequestHeader "Content-length", Len(params)
     DomDoc.setRequestHeader "Connection", "close"
     DomDoc.send params
     
    'La respuesta, en caso de existir, está en responseBody.
    'También puedes especificar responseXml si tu aplicación devolviese XML
    
     strHtml = StrConv(DomDoc.responseBody, vbUnicode)
     
     
     Dim p As Object
     Set p = JSON.parse(strHtml)
     Nombre = Trim(p.Item("apellidoPaterno")) & Space(1) & Trim(p.Item("apellidoMaterno")) & Space(1) & Trim(p.Item("nombres"))
     If Trim(Nombre) <> "" Then
        strCadena = " call P_insert_persona_ii('" & Trim(in_dni) & "','" & Trim(p.Item("apellidoPaterno")) & "','" & Trim(p.Item("apellidoMaterno")) & "','" & Trim(p.Item("nombres")) & "','" & Trim(Nombre) & "','-','','-','no','no','no','no','no','no','si','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        get_dni_reniec = True
    Else
        get_dni_reniec = False
    End If
    


     
     
     
     Exit Function
  get_dni_reniec = False
salir_cancel:

End Function

Public Function get_dni_reniec_ii(ByVal in_dni As String) As Boolean
Dim Nombre As String
Dim strHtml As String
Set DomDoc = New XMLHTTP
     
    On Error GoTo salir_cancel

     urlstr = "http://facturacion.vitekey.com/api/utiles/consultar_dni?dni=" & in_dni
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
     
     Dim p As Object
     Set p = JSON.parse(strHtml)
     Nombre = Replace(Trim(p.Item("nombre")), "‘", " ")
     Nombre = Replace(Nombre, "Ã ", "Ñ")
     Nombre = Replace(Nombre, "‘", " ")
     Nombre = Replace(Nombre, "'", " ")
     If Trim(Nombre) <> "" Then
        strCadena = " call P_insert_persona_ii('" & Trim(in_dni) & "','-','-','-','" & Trim(Nombre) & "','" & KEY_DIR_PUBLIC & "','','-','no','no','no','no','no','no','si','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        get_dni_reniec_ii = True
    Else
        get_dni_reniec_ii = False
    End If
    

     
     Exit Function
  get_dni_reniec_ii = False
salir_cancel:

End Function

Public Function get_dni_reniec_iii(ByVal in_dni As String, ByVal in_departamento As String, ByVal in_provincia As String, ByVal in_distrito As String) As Boolean
Dim Nombre As String
Dim strHtml As String
Set DomDoc = New XMLHTTP
     
    On Error GoTo salir_cancel
    
     If Len(Trim(in_dni)) = 8 Then
        urlstr = "http://facturacion.vitekey.com/api/utiles/consultar_dni?dni=" & in_dni
     Else
        urlstr = "https://api.vitekey.com/keyfact/erp/utils/search-ruc?ruc=" & in_dni & "&api_key=fd235235-e97a-4db6-8f50-fa84145c3f5d"
     End If
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
     
     Dim p As Object
     Set p = JSON.parse(strHtml)
     
      If Len(Trim(in_dni)) = 8 Then
         Nombre = Replace(Trim(p.Item("nombre")), "‘", " ")
         in_direccion = KEY_DIR_PUBLIC
      Else
         Nombre = Replace(Trim(p.Item("razon_social")), "‘", " ")
         in_direccion = Replace(Trim(p.Item("direccion")), "‘", " ")
         If in_direccion = "" Then
            in_direccion = KEY_DIR_PUBLIC
         End If
      End If
     
     Nombre = Replace(Nombre, "Ã ", "Ñ")
     Nombre = Replace(Nombre, "‘", " ")
     Nombre = Replace(Nombre, "'", " ")
     If Trim(Nombre) <> "" Then
        strCadena = " call p_insert_persona_iii('" & Trim(in_dni) & "','-','-','-','" & Trim(Nombre) & "','" & in_direccion & "','','-','no','no','no','no','no','no','si','" & in_departamento & "','" & in_provincia & "','" & in_distrito & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        get_dni_reniec_iii = True
    Else
        get_dni_reniec_iii = False
    End If
    

     
     
     Exit Function
  get_dni_reniec_iii = False
salir_cancel:

End Function


Public Sub put_insert_producto_alm(ByVal in_alm As String)
On Error GoTo salir
    strCadena = "SELECT * FROM almacen_producto WHERE id_alm='00001' and  ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
        strCadena = "INSERT INTO almacen_producto(id_alm,id_producto,stock,stock_factura,precio_venta,precio_compra,precio_mayor,habilitado,ruc) VALUES  " & _
        "('" & in_alm & "','" & rst("id_producto") & "','0','0','" & rst("precio_venta") & "','" & rst("precio_compra") & "','" & rst("precio_mayor") & "','" & rst("habilitado") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        rst.MoveNext
        Next i
       
    End If
    Exit Sub
salir:
End Sub

    

Public Sub delete_almacen(ByVal in_alm As String)
strCadena = "DELETE FROM almacen WHERE id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)



End Sub

Public Sub delete_producto(ByVal in_producto As String)
strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_producto='" & Trim(in_producto) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            MsgBox "ESTE PRODUCTO CUENTA CON DOCUMENTOS RELACIONADOS" + Chr(13) + "ELIMINE LOS COMPROBANTES ANTES DE ELIMINAR EL PRODUCTO", vbInformation, KEY_EMPRESA
            
            Exit Sub
        End If
            strCadena = "DELETE FROM producto WHERE id_producto='" & Trim(in_producto) & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
            strCadena = "DELETE FROM almacen_producto WHERE id_producto='" & Trim(in_producto) & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
End Sub

Public Function get_electronico_online(ByVal in_doc As String, ByVal in_serie As String) As String
strCadena = "SELECT electronico,online FROM almacen_comprobante WHERE electronico='si' and  id_doc='" & in_doc & "' and serie='" & in_serie & "' and id_alm='" & KEY_VENTANILLA & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
   get_electronico_online = rstZ("online")
Else
   get_electronico_online = "no"
End If

End Function
Public Function get_firma_online(ByVal in_doc As String, ByVal in_serie As String) As String

    strCadena = "SELECT firmado_online FROM almacen_comprobante WHERE electronico='si' and  id_doc='" & in_doc & "' and serie='" & in_serie & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        get_firma_online = rstZ("firmado_online")
    Else
        get_firma_online = "no"
        
    End If

End Function

Public Function get_comprobante_produccion(ByVal in_doc As String, ByVal in_serie As String) As String
strCadena = "SELECT produccion FROM almacen_comprobante WHERE electronico='si' and  id_doc='" & in_doc & "' and serie='" & in_serie & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
   get_comprobante_produccion = rstZ("produccion")
Else
   get_comprobante_produccion = "no"
End If

End Function

Public Function json_facturacion_electronica(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_fecha As String, ByVal in_cliente As String, ByVal in_cliente_nombre As String, ByVal in_tipo_cliente As String, ByVal in_descuento As Single)

Dim in_monto As Single
Dim in_detalle As String
Dim menor_edad As Boolean
Dim Item As String
          
      
strCadena = "SELECT * FROM temporal_ventas WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' ORDER BY id ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
 
   For i = 0 To rst.RecordCount - 1
        
        If i = 0 Then
           Item = "{ codigo: '" & rst("id_producto") & "', descripcion: '" & rst("detalle") & "', cantidad: '" & rst("cantidad") & "', precio_unitario:'" & rst("precio") & "',tipo_precio:'01', descuento:'0.00',tax_code:10 }"
        Else
           Item = Item & "," & "{ codigo: '" & rst("id_producto") & "', descripcion: '" & rst("detalle") & "', cantidad: '" & rst("cantidad") & "', precio_unitario:'" & rst("precio") & "',tipo_precio:'01', descuento:'0.00',tax_code:10 }"
        End If
        
        rst.MoveNext
   Next i
End If

If in_cliente = "00000000" Then
   in_cliente = "0"
End If
json_facturacion_electronica = "{ tipo_documento: '" & in_doc & "',serie: '" & in_serie & "',numero:'" & in_numero & "',ruc: '" & KEY_RUC & "',ubigeo: '140101', fecha: '" & Format(in_fecha, "YYYY-mm-dd") & "', id_cliente: '" & in_cliente & "', tipo_cliente:'" & in_tipo_cliente & "',nombre_cliente:'" & in_cliente_nombre & "',descuento_global:'" & Format(in_descuento, "#,##0.00") & "', items:[ " & Item & "  ]   }"


    

End Function
Public Function json_facturacion_electronica_firmar(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_fecha As String, ByVal in_cliente As String, ByVal in_cliente_nombre As String, ByVal in_cliente_direccion As String, ByVal in_tipo_cliente As String, ByVal in_descuento As Single, ByVal in_igv As Single, ByVal id_motivo_nota As String, ByVal motivo_nota As String, ByVal id_tipo_doc_afectado As String, ByVal in_serie_afectado As String, ByVal in_numero_afectado As String, ByVal in_moneda As String)
Dim in_monto As Single
Dim in_detalle As String
Dim menor_edad As Boolean
Dim Item As String
Dim in_referencia As String
Dim in_valor_venta As Single
     
strCadena = "SELECT * FROM temporal_ventas WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' ORDER BY id ASC"
'strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='324051'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
 
   For i = 0 To rst.RecordCount - 1
       If rst("igv") = "si" Then
          in_valor_venta = rst("precio") / (1 + KEY_IGV)
       Else
          in_valor_venta = rst("precio")
       End If
        
        '---- REFERENCIA
        in_referencia = ""
        If Len(rst("nro_chasis")) > 2 Then
            in_referencia = "MARCA:" & get_marca(rst("id_producto")) & "\nCHASIS:" & rst("nro_chasis") & "\nMOTOR:" & rst("serie") & "\nCOLOR:" & get_color(rst("id_producto")) & "\nAÑO FABRICACION:" & rst("anio_fabricacion") & "\nNRO DUA:" & rst("nro_dua") & "\nNRO ITEM:" & rst("nro_item")
        Else
            in_referencia = ""
        End If
        '--- END REFERENCIA
        
        If i = 0 Then
           Item = "{ codigo_producto: '" & rst("id_producto") & "', descripcion_producto: '" & rst("detalle") & "', cantidad: '" & rst("cantidad") & "', valor_unitario:'" & in_valor_venta & "',tipo_precio:'01',tipo_igv:'10',codigo_unidad_medida:'NIU',referencia_producto:'" & in_referencia & "', descuento:'0.00',tax_code:10 }"
        Else
           Item = Item & "," & "{ codigo_producto: '" & rst("id_producto") & "', descripcion_producto: '" & rst("detalle") & "', cantidad: '" & rst("cantidad") & "', valor_unitario:'" & in_valor_venta & "',tipo_precio:'01',tipo_igv:'10',codigo_unidad_medida:'NIU',referencia_producto:'" & in_referencia & "', descuento:'0.00',tax_code:10 }"
        End If
        
        rst.MoveNext
   Next i
End If

If in_cliente = "00000000" Then
   in_cliente = "00000000"
End If



json_facturacion_electronica_firmar = "{ enviar_sunat: false, tipo: '" & in_doc & "',serie: '" & in_serie & "',numero:'" & in_numero & "',tipo_operacion:'01',ruc: '" & KEY_RUC & "', fecha_emision: '" & Format(in_fecha, "YYYY-mm-dd") & "', cliente_numero_doc: '" & in_cliente & "', cliente_tipo_doc:'" & in_tipo_cliente & "',cliente_nombre:'" & Replace(in_cliente_nombre, "&", " ") & "',cliente_direccion:'" & in_cliente_direccion & "',cliente_email:'" & get_mail(in_cliente) & "',tipo_moneda:'" & in_moneda & "',igv_factor:'" & in_igv & "',descuento_global:'" & Format(in_descuento, "#,##0.00") & "',codigo_motivo:'" & id_motivo_nota & "',descripcion_motivo:'" & motivo_nota & "',tipo_documento_afectado:'" & id_tipo_doc_afectado & "',serie_documento_afectado:'" & in_serie_afectado & "',numero_documento_afectado:'" & in_numero_afectado & "', detalle:[ " & Item & "  ]   }"
    

End Function
Public Function json_facturacion_electronica_eliminar(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String)
Dim in_monto As Single
Dim in_detalle As String
Dim menor_edad As Boolean
Dim Item As String
Dim in_referencia As String
Dim in_valor_venta As Single
     

If KEY_SERVIDOR_KEYFACIL = "si" Then '{invoice_id: '"& in_key &"'}
    strCadena = "SELECT sunat_key FROM movimiento_venta WHERE id_doc='" & Format(in_doc, "0000") & "' and serie='" & in_serie & "' and numero='" & in_numero & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount > 0 Then
       json_facturacion_electronica_eliminar = "{invoice_id:'" & rstA("sunat_key") & "'}"
    End If
    
Else
    json_facturacion_electronica_eliminar = "{tipo: '" & in_doc & "',serie: '" & in_serie & "',numero:'" & in_numero & "',ruc: '" & KEY_RUC & "' }"
End If


    

End Function
Public Function json_crear_producto(ByVal in_codigo As String, ByVal in_descripcion As String, ByVal in_precio As Double)

json_crear_producto = "{type: 'NIU',unit: 'NIU',code:'" & in_codigo & "',description:'" & in_descripcion & "',currency:'PEN',price_unit:'" & Val(in_precio) & "',price_variants:[],igv_type:'10'}"

End Function

Public Function json_crear_get_producto()

json_crear_get_producto = "{company_id: '" & KEY_TOKEN_CLOUD & "',offset:0}"

End Function

Public Function json_facturacion_electronica_mail(ByVal ruc As String, ByVal in_key As String, IN_asunto As String, ByVal in_mail As String)
Dim in_monto As Single
Dim in_detalle As String
Dim menor_edad As Boolean
Dim Item As String
Dim in_referencia As String
Dim in_valor_venta As Single
Dim in_maill() As String

    If Len(in_key) > 36 Then
        json_facturacion_electronica_mail = "{key: '" & in_key & "',asunto:'" & IN_asunto & "',email: '" & in_mail & "',ruc: '" & KEY_RUC & "' }"
    Else
        in_maill = Split(Trim(in_mail), ";")
        If UBound(in_maill()) > 1 Then
            json_facturacion_electronica_mail = "{invoice_id: '" & in_key & "',asunto:'" & IN_asunto & "',email: ['" & in_maill(0) & "', '" & in_maill(1) & "'] }"
        Else
            'json_facturacion_electronica_mail = "{invoice_id: '" & in_key & "',asunto:'" & IN_asunto & "',email: '" & in_mail & "' }"
            json_facturacion_electronica_mail = "{invoice_id: '" & in_key & "',asunto:'" & IN_asunto & "',email: ['" & in_maill(0) & "'] }"
        End If
        
        
    End If
     
   
End Function

Public Function json_facturacion_electronica_firmar_id_venta(ByVal in_venta As String, ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_fecha As String, ByVal in_cliente As String, ByVal in_cliente_nombre As String, ByVal in_cliente_direccion As String, ByVal in_tipo_cliente As String, ByVal in_descuento As Single, ByVal in_igv As Single, ByVal id_motivo_nota As String, ByVal motivo_nota As String, ByVal id_tipo_doc_afectado As String, ByVal in_serie_afectado As String, ByVal in_numero_afectado As String, ByVal in_moneda As String, ByVal in_observacion As String)
Dim in_monto As Single
Dim in_detalle As String
Dim menor_edad As Boolean
Dim Item As String
Dim in_referencia As String
Dim in_valor_venta As Single
Dim in_tipo_igv As String

     
'strCadena = "SELECT * FROM temporal_ventas WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' ORDER BY id ASC"
strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
 
   For i = 0 To rst.RecordCount - 1
       If KEY_CON_IGV = "si" Then
          
          If KEY_APLICA_IGV = "no" Then
            in_valor_venta = (rst("total") / rst("cantidad"))
            in_valor_venta = in_valor_venta
            in_igv = 0
            in_tipo_igv = "20"
          Else
            in_valor_venta = (rst("total") / rst("cantidad"))
            in_valor_venta = in_valor_venta / (1 + KEY_IGV)
            If in_valor_venta = 0 Then
                in_tipo_igv = "21"
                in_valor_venta = rst("precio") / (1 + KEY_IGV)
            Else
                in_tipo_igv = "10"
            End If
            
            If in_tipo_cliente = "0" Then
                in_tipo_igv = "40"
            End If
            
            in_igv = KEY_IGV
          End If
       Else
          
          in_valor_venta = (rst("total") / rst("cantidad"))
          
          If in_valor_venta = 0 Then
                in_tipo_igv = "21"
                in_valor_venta = rst("precio")
          Else
                in_tipo_igv = "20"
                in_valor_venta = in_valor_venta
          End If
            
          
          in_igv = 0
        
          
          'RECORDAR QUE TIPO IGV=40 ES PARA EXPORTACION
       End If
        
        
        
        
        '---- REFERENCIA
        in_referencia = ""
        If KEY_REFERENCIA_COMPROBANTE = "si" Then
        If Len(rst("nro_chasis")) > 2 Then
            in_referencia = "MARCA:" & get_marca(rst("id_producto")) & "\n" & in_chasis & rst("nro_chasis") & "\n" & in_motor & rst("serie") & "\nCOLOR:" & get_color(rst("id_producto")) & "\nAÑO FABRICACION:" & rst("anio_fabricacion") & "\nNRO DUA:" & rst("nro_dua") & "\nNRO ITEM:" & rst("nro_item")
        Else
            in_referencia = ""
        End If
        End If
        '--- END REFERENCIA
        
        If i = 0 Then
           Item = "{ codigo_producto: '" & rst("id_producto") & "', descripcion_producto: '" & Replace(rst("detalle"), "&", " ") & "', cantidad: '" & rst("cantidad") & "', valor_unitario:'" & in_valor_venta & "',tipo_precio:'01',tipo_igv:'" & in_tipo_igv & "',codigo_unidad_medida:'NIU',referencia_producto:'" & in_referencia & "', descuento:'0.00',tax_code:10 }"
        Else
           Item = Item & "," & "{ codigo_producto: '" & rst("id_producto") & "', descripcion_producto: '" & Replace(rst("detalle"), "&", " ") & "', cantidad: '" & rst("cantidad") & "', valor_unitario:'" & in_valor_venta & "',tipo_precio:'01',tipo_igv:'" & in_tipo_igv & "',codigo_unidad_medida:'NIU',referencia_producto:'" & in_referencia & "', descuento:'0.00',tax_code:10 }"
        End If
        
        rst.MoveNext
   Next i
End If

If in_cliente = "00000000" Then
   in_cliente = "00000000"
End If



If KEY_CON_IGV = "si" Then
    json_facturacion_electronica_firmar_id_venta = "{ enviar_sunat: false, tipo: '" & in_doc & "',serie: '" & in_serie & "',numero:'" & in_numero & "',tipo_operacion:'01',ruc: '" & KEY_RUC & "', fecha_emision: '" & Format(in_fecha, "YYYY-mm-dd") & "', cliente_numero_doc: '" & in_cliente & "', cliente_tipo_doc:'" & in_tipo_cliente & "',cliente_nombre:'" & Replace(in_cliente_nombre, "&", " ") & "',cliente_direccion:'" & in_cliente_direccion & "',cliente_email:'" & get_mail(in_cliente) & "',tipo_moneda:'" & in_moneda & "',igv_factor:'" & in_igv & "',descuento_global_gravado:'" & Format(Round(in_descuento / (1 + KEY_IGV), 4), "###0.0000") & "',codigo_motivo:'" & id_motivo_nota & "',descripcion_motivo:'" & motivo_nota & "',tipo_documento_afectado:'" & id_tipo_doc_afectado & "',serie_documento_afectado:'" & in_serie_afectado & "',numero_documento_afectado:'" & in_numero_afectado & "',observacion:'" & in_observacion & "', detalle:[ " & Item & "  ]   }"
Else
    json_facturacion_electronica_firmar_id_venta = "{ enviar_sunat: false, tipo: '" & in_doc & "',serie: '" & in_serie & "',numero:'" & in_numero & "',tipo_operacion:'01',ruc: '" & KEY_RUC & "', fecha_emision: '" & Format(in_fecha, "YYYY-mm-dd") & "', cliente_numero_doc: '" & in_cliente & "', cliente_tipo_doc:'" & in_tipo_cliente & "',cliente_nombre:'" & Replace(in_cliente_nombre, "&", " ") & "',cliente_direccion:'" & in_cliente_direccion & "',cliente_email:'" & get_mail(in_cliente) & "',tipo_moneda:'" & in_moneda & "',igv_factor:'" & in_igv & "',descuento_global_exonerado:'" & Format(Round(in_descuento, 2), "###0.0000") & "',codigo_motivo:'" & id_motivo_nota & "',descripcion_motivo:'" & motivo_nota & "',tipo_documento_afectado:'" & id_tipo_doc_afectado & "',serie_documento_afectado:'" & in_serie_afectado & "',numero_documento_afectado:'" & in_numero_afectado & "',observacion:'" & in_observacion & "', detalle:[ " & Item & "  ]   }"
End If
    

End Function
Public Function json_facturacion_electronica_firmar_guia(ByVal in_transferencia As String)
Dim in_monto As Single
Dim in_detalle As String
Dim menor_edad As Boolean
Dim Item As String
Dim in_referencia As String
Dim in_traslado As Date
     
strCadena = "SELECT * FROM movimiento_transferencia WHERE id_transferencia='" & in_transferencia & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    
    
    'put_fecha = Format("2020-01-16", "YYYY/mm/dd HH:mm") & " GMT-0500"
    put_fecha = Format(rst("fecha"), "YYYY/mm/dd HH:mm") & " GMT-0500"
    put_traslado = Format(rst("fecha_traslado"), "YYYY/mm/dd HH:mm") & " GMT-0500"
    
   ' in_traslado = Format("2020-01-16", "YYYY-mm-dd")
    
    Select Case Len(rst("id_destinatario"))
            Case 8
                in_tipo_cliente = 1
            Case 11
                in_tipo_cliente = 6
            Case Break
                in_tipo_cliente = 4
    End Select
    
     Select Case Len(rst("id_transporte"))
            Case 8
                in_tipo_transporte = 1
            Case 11
                in_tipo_transporte = 6
            Case Break
                in_tipo_transporte = 4
    End Select
    
    If rst("id_venta") > 0 Then
        in_comprobante_rel = get_comprobante_venta(rst("id_venta"))
    Else
        in_comprobante_rel = ""
    End If
    origi = "{ ubigeo: '" & get_ubigeo_transferencia(in_transferencia, "origen") & "', address: '" & get_direccion(rst("id_remitente")) & "' }"
    deliver = "{ ubigeo: '" & get_ubigeo_transferencia(in_transferencia, "destino") & "', address: '" & rst("direccion_destino") & "' }"
    transportist = "{ type: '" & in_tipo_transporte & "', docid: '" & rst("id_transporte") & "',name: '" & get_persona(rst("id_transporte")) & "' }"
    in_tipo_chofer = 1
    
    
    
    If rst("id_chofer") = "" Then
        in_nchofer = "-"
        in_nchofer_des = "-"
        in_placa = "-"
    Else
        in_nchofer = rst("id_chofer")
        in_nchofer_des = get_persona(rst("id_chofer"))
        in_placa = rst("placa")
    End If
    
        driver = "{ type: '" & in_tipo_chofer & "', docid: '" & in_nchofer & "',name: '" & in_nchofer_des & "',plate: '" & in_placa & "' }"
   
    
    'If rst("tipo_transporte") = 1 Then
    '    transportist = "NULL"
    'Else
    '    driver = "NULL"
   ' End If
    
    
    shipmen = "{ transport_type: '" & Format(rst("tipo_transporte"), "00") & "', handling_code: '01', start_date: '" & Format(put_traslado, "YYYY-mm-dd") & "', arrival_port_code:'',split_consignment_indicator:'false',gross_weight_measure:'" & rst("peso_total") & "', gross_weight_unit:'KGM',total_unit_quantity:'" & rst("numero_bultos") & "',container_number:'',origin: " & origi & "  ,delivery: " & deliver & "  ,transportist: " & transportist & "  ,driver: " & driver & "  }"
    
    
    strCadena = "SELECT * FROM movimiento_transferencia_detalle WHERE id_transferencia='" & Val(in_transferencia) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstK(strCadena)
    If rstK.RecordCount > 0 Then
        rstK.MoveFirst
        
        For i = 0 To rstK.RecordCount - 1
            If i = 0 Then
                Item = "{ type: 'NIU',code: '" & rstK("id_producto") & "', description: '" & Replace(rstK("detalle"), "&", " ") & "', quantity: '" & rstK("cantidad") & "' }"
            Else
                Item = Item & "," & "{ type: 'NIU',code: '" & rstK("id_producto") & "', description: '" & Replace(rstK("detalle"), "&", " ") & "', quantity: '" & rstK("cantidad") & "' }"
            End If
            rstK.MoveNext
        Next i
    End If

    
    
    json_facturacion_electronica_firmar_guia = "{ office_id: '" & KEY_TOKEN_SUCURSAL & "',type: '" & Format(Val(rst("id_doc")), "00") & "',serie: '" & rst("serie") & "',number: '" & rst("numero") & "', issue_date: '" & put_fecha & "',receiver_type:'" & in_tipo_cliente & "',receiver_docid:'" & rst("id_destinatario") & "',receiver_name:'" & rst("destinatario") & "',receiver_email:null , receiver_address: '" & rst("direccion_destino") & "' ,observations:'" & rst("observacion") & "', shipment: " & shipmen & "   ,items:[ " & Item & "  ]   }"
End If
    
    

    

End Function


Public Function json_facturacion_electronica_firmar_id_venta_keyfacil(ByVal in_venta As String, ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_fecha As String, ByVal in_cliente As String, ByVal in_cliente_nombre As String, ByVal in_cliente_direccion As String, ByVal in_tipo_cliente As String, ByVal in_descuento As Single, ByVal in_igv As Single, ByVal id_motivo_nota As String, ByVal motivo_nota As String, ByVal id_tipo_doc_afectado As String, ByVal in_serie_afectado As String, ByVal in_numero_afectado As String, ByVal in_moneda As String, ByVal in_observacion As String, ByVal in_detraccion As String, ByVal in_ordencompra As String, Optional in_monto_percepcion As Single)
Dim in_monto As Single
Dim in_detalle As String
Dim menor_edad As Boolean
Dim Item As String
Dim in_referencia As String
Dim in_valor_venta As Single
Dim in_tipo_igv As String
Dim n_icbper As Boolean
Dim in_ccicper As String
Dim tipo_operacion As String

  Item = ""
'strCadena = "SELECT * FROM temporal_ventas WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' ORDER BY id ASC"
strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
 
   For i = 0 To rst.RecordCount - 1
       
       If KEY_CON_IGV = "si" Then
          If KEY_APLICA_IGV = "no" Then
            
            in_valor_venta = (rst("total") / rst("cantidad"))
            in_valor_venta = in_valor_venta
            in_igv = 0
            in_tipo_igv = "20"
          
          
          
          Else
        
            If KEY_SERVIDOR_KEYFACIL = "si" Then
                'in_valor_venta = (rst("total") / rst("cantidad"))
                in_valor_venta = rst("precio")
            Else
                in_valor_venta = in_valor_venta / (1 + KEY_IGV)
            End If
            
            If rst("total") = 0 Then
                in_tipo_igv = "31" 'Bonificacion
            Else
                in_tipo_igv = "10"
            End If
            
            in_igv = KEY_IGV
          
          End If
       Else
          'in_valor_venta = (rst("total") / rst("cantidad"))
          in_valor_venta = rst("precio")
          in_valor_venta = in_valor_venta
          in_igv = 0
            
            If rst("total") = 0 Then
                in_tipo_igv = "31" 'Bonificacion
            Else
                in_tipo_igv = "20"
            End If
            
          
          
          'RECORDAR QUE TIPO IGV=40 ES PARA EXPORTACION
       End If
        
        If in_tipo_cliente = "0" Then
                in_tipo_igv = "40"
            End If
        
        
        
        '---- REFERENCIA
        in_referencia = ""
        If KEY_REFERENCIA_COMPROBANTE = "si" Then
        If Len(rst("nro_chasis")) > 2 Then
            'in_referencia = "MARCA:" & get_marca(rst("id_producto")) & "\n" & in_chasis & rst("nro_chasis") & "\n" & in_motor & rst("serie") & "\nCOLOR:" & get_color(rst("id_producto")) & "\nAÑO FABRICACION:" & rst("anio_fabricacion") & "\nNRO DUA:" & rst("nro_dua") & "\nNRO ITEM:" & rst("nro_item")
            If KEY_RUC = "20479779598" Then
                in_referencia = "MARCA:" & get_marca(rst("id_producto")) & "\nCOLOR:" & get_color(rst("id_producto")) & "\nAÑO FABRICACION:" & rst("anio_fabricacion") & "\nNRO DUA:" & rst("nro_dua") & "\nNRO ITEM:" & rst("nro_item")
            Else
                in_referencia = "MARCA:" & get_marca(rst("id_producto")) & "\nAÑO FABRICACION:" & rst("anio_fabricacion") & "\nNRO DUA:" & rst("nro_dua") & "\nNRO ITEM:" & rst("nro_item")
                
                'in_referencia = "MARCA:" & get_marca(rst("id_producto")) & "\nCOLOR :" & get_color(rst("id_producto")) & "\nAÑO FABRICACION: 2017" & "\nNRO DUA:" & rst("nro_dua")
                
               ' in_referencia = "MARCA ..............: " & get_marca(rst("id_producto")) & "\nTIPO CARROCERIA........:    VOLQUETE" & "\nMODELO.............:    ZZ3257V3647P1" & "\nMOTOR..............:      180417016777" & "\nNRO VIN........:      LZZ5ELSD4JA369576" & "\nSERIE...........:      LZZ5ELSD4JA369576" & "\nAÑO FABRICACION.................:2018" & "\nAÑO MODELO...............:2018" & "\nCOLOR.............:" & get_color(rst("id_producto")) & "\nCATEGORIA.................:       N3" & "\nCOMBUSTIBLE..............:       DIESEL" & "\nCAPACIDAD TOLVA:        17 M3"
            End If
            
        Else
            in_referencia = ""
        End If
        End If
        '--- END REFERENCIA
        If rst("icbper") = "si" Then
            in_ccicper = "true"
        Else
            in_ccicper = "false"
        End If
        If i = 0 Then
            
           Item = "{ code: '" & rst("id_producto") & "', description: '" & Replace(rst("detalle"), "&", " ") & "', quantity: '" & rst("cantidad") & "', unit_price:'" & in_valor_venta & "',igv_type:'" & in_tipo_igv & "',detail:'" & in_referencia & "', discount_percent:'0.00',icbper:" & in_ccicper & " }"
           'Item = "{ code: '" & rst("id_producto") & "', description: '" & Replace(rst("detalle"), "&", " ") & "', quantity: '" & rst("cantidad") & "', unit_price:'" & in_valor_venta & "',igv_type:'" & in_tipo_igv & "',detail:'" & in_referencia & "', discount_percent:'0.00' }"
        Else
           Item = Item & "," & "{ code: '" & rst("id_producto") & "', description: '" & Replace(rst("detalle"), "&", " ") & "', quantity: '" & rst("cantidad") & "', unit_price:'" & in_valor_venta & "',igv_type:'" & in_tipo_igv & "',detail:'" & in_referencia & "', discount_percent:'0.00',icbper:" & in_ccicper & " }"
           'Item = Item & "," & "{ code: '" & rst("id_producto") & "', description: '" & Replace(rst("detalle"), "&", " ") & "', quantity: '" & rst("cantidad") & "', unit_price:'" & in_valor_venta & "',igv_type:'" & in_tipo_igv & "',detail:'" & in_referencia & "', discount_percent:'0.00' }"
        End If
        
        rst.MoveNext
   Next i
End If

If in_cliente = "00000000" Then
   in_cliente = "00000000"
End If

put_fecha = Format(in_fecha, "YYYY/mm/dd HH:mm") & " GMT-0500"



'******** percepcion
'operation_type:2001
'perception_percent:2




ref_ordencompra = "order_reference:'" & Trim(in_ordencompra) & "'"



If in_monto_percepcion > 0 Then
    tipo_operacion = "operation_type:'2001',perception_percent:'2'"
Else
    tipo_operacion = "operation_type:'0101'"
End If

If in_detraccion = "si" Then
    tipo_operacion = "operation_type:'1001', detraction_type: '024', detraction_percent:'" & KEY_PORCENTAJE_DETRACCION & "', detraction_account_number: '" & KEY_CTA_DETRACCION & "', detraction_payment_type: '001'"
End If


 
 


If in_tipo_cliente = "0" Then 'Exportacion
    tipo_operacion = "operation_type:'0200'"
End If


If KEY_CON_IGV = "si" Then
'note_info ": { "invoice_modified_type": "01", "invoice_modified_serie": "FF01","invoice_modified_number: 15, "reason": "Error", "type": "01"}
   
   
   
    If in_doc = "07" Or in_doc = "08" Then
        json_facturacion_electronica_firmar_id_venta_keyfacil = "{ office_id: '" & KEY_TOKEN_SUCURSAL & "',type: '" & in_doc & "',serie: '" & in_serie & "',number:'" & in_numero & "'," & ref_ordencompra & "," & tipo_operacion & ", issue_date: '" & put_fecha & "', client_docid: '" & in_cliente & "', client_type:'" & in_tipo_cliente & "',client_name:'" & Replace(in_cliente_nombre, "&", " ") & "',client_address:'" & in_cliente_direccion & "',client_email:'" & get_mail(in_cliente) & "',currency:'" & in_moneda & "',global_discount_percent:'" & Format(Round(in_descuento, 2), "###0.00") & "',note_info : { invoice_modified_type: '" & id_tipo_doc_afectado & "', invoice_modified_serie:'" & in_serie_afectado & "',invoice_modified_number:'" & in_numero_afectado & "', reason:'" & motivo_nota & "', type:'" & id_motivo_nota & "'} ,observations:'" & in_observacion & "', items:[ " & Item & "  ]   }"
    Else
        'json_facturacion_electronica_firmar_id_venta_keyfacil = "{ office_id: '" & KEY_TOKEN_SUCURSAL & "',type: '" & in_doc & "',serie: '" & in_serie & "',number:'" & in_numero & "',operation_type:'0101', issue_date: '" & put_fecha & "', client_docid: '" & in_cliente & "', client_type:'" & in_tipo_cliente & "',client_name:'" & Replace(in_cliente_nombre, "&", " ") & "',client_address:'" & in_cliente_direccion & "',client_email:'" & get_mail(in_cliente) & "',currency:'" & in_moneda & "',global_discount_percent:'" & Format(Round(in_descuento, 2), "###0.00") & "',observations:'" & in_observacion & "', items:[ " & Item & "  ]   }"
        json_facturacion_electronica_firmar_id_venta_keyfacil = "{ office_id: '" & KEY_TOKEN_SUCURSAL & "',type: '" & in_doc & "',serie: '" & in_serie & "',number:'" & in_numero & "'," & ref_ordencompra & "," & tipo_operacion & ", issue_date: '" & put_fecha & "', client_docid: '" & in_cliente & "', client_type:'" & in_tipo_cliente & "',client_name:'" & Replace(in_cliente_nombre, "&", " ") & "',client_address:'" & in_cliente_direccion & "',client_email:'" & get_mail(in_cliente) & "',currency:'" & in_moneda & "',global_discount_percent:'" & Format(Round(in_descuento, 2), "###0.00") & "',observations:'" & in_observacion & "', items:[ " & Item & "  ]   }"
    End If
Else
   
    If in_doc = "07" Or in_doc = "08" Then
        json_facturacion_electronica_firmar_id_venta_keyfacil = "{ office_id: '" & KEY_TOKEN_SUCURSAL & "',type: '" & in_doc & "',serie: '" & in_serie & "',number:'" & in_numero & "'," & ref_ordencompra & "," & tipo_operacion & ", issue_date: '" & put_fecha & "', client_docid: '" & in_cliente & "', client_type:'" & in_tipo_cliente & "',client_name:'" & Replace(in_cliente_nombre, "&", " ") & "',client_address:'" & in_cliente_direccion & "',client_email:'" & get_mail(in_cliente) & "',currency:'" & in_moneda & "',global_discount_percent:'" & Format(Round(in_descuento, 2), "###0.00") & "',note_info : { invoice_modified_type: '" & id_tipo_doc_afectado & "', invoice_modified_serie:'" & in_serie_afectado & "',invoice_modified_number:'" & in_numero_afectado & "', reason:'" & motivo_nota & "', type:'" & id_motivo_nota & "'} ,observations:'" & in_observacion & "', items:[ " & Item & "  ]   }"
    Else
        json_facturacion_electronica_firmar_id_venta_keyfacil = "{ office_id: '" & KEY_TOKEN_SUCURSAL & "',type: '" & in_doc & "',serie: '" & in_serie & "',number:'" & in_numero & "'," & ref_ordencompra & "," & tipo_operacion & ", issue_date: '" & put_fecha & "', client_docid: '" & in_cliente & "', client_type:'" & in_tipo_cliente & "',client_name:'" & Replace(in_cliente_nombre, "&", " ") & "',client_address:'" & in_cliente_direccion & "',client_email:'" & get_mail(in_cliente) & "',currency:'" & in_moneda & "',global_discount_percent:'" & Format(Round(in_descuento, 2), "###0.00") & "',observations:'" & in_observacion & "', items:[ " & Item & "  ]   }"
    End If
    
    
End If
    

End Function






Public Function get_ubigueo_persona(ByVal in_dni As String, ByVal in_direccion As String)
If in_direccion > 0 Then
            strCadena = "SELECT * FROM view_persona_direccion WHERE id_direccion='" & Val(in_direccion) & "'"
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount > 0 Then
               get_ubigueo_persona = UCase(rstK("distrito") & Space(5) & rstK("provincia") & Space(5) & rstK("departamento"))
            Else
              get_ubigueo_persona = ""
            End If
Else
    If in_dni <> "00000000" Then
        strCadena = "SELECT * FROM persona WHERE dni='" & in_dni & "'"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            get_ubigueo_persona = get_ubigueo(rstK("id_departamento"), rstK("id_provincia"), rstK("id_distrito"))
        End If
    Else
        get_ubigueo_persona = get_ubigueo(KEY_DEPARTAMENTO, KEY_PROVINCIA, KEY_DISTRITO)
    End If
End If
End Function

Public Function verificar_finalizado(ByVal in_orden_compra As String) As Boolean
verificar_finalizado = True
strCadena = "SELECT * FROM orden_compra_detalle WHERE id_orden='" & Val(in_orden_compra) & "' ORDER BY id_detalle ASC"
      Call ConfiguraRstL(strCadena)
      If rstL.RecordCount > 0 Then
         rstL.MoveFirst
         For i = 0 To rstL.RecordCount - 1
              in_pendiente = rstL("cantidad") - get_recepcionado(rstL("id_producto"), Val(in_orden_compra))
              If in_pendiente > 0 Then
                 verificar_finalizado = False
                 Exit For
              End If
              rstL.MoveNext

         Next i
      End If
End Function
Private Function get_recepcionado(ByVal in_producto As String, ByVal in_orden As String) As Single
strCadena = "SELECT sum(d.cantidad) FROM  orden_compra o,orden_compra_detalle d WHERE o.id_estado<>3 and  o.id_orden=d.id_orden and d.id_producto='" & in_producto & "' and o.id_recepcion='" & Val(in_orden) & "' and o.ruc='" & KEY_RUC & "'"
Call ConfiguraRstP(strCadena)
If IsNull(rstP(0)) = True Then
    get_recepcionado = 0
Else
    get_recepcionado = rstP(0)
End If


End Function

Public Function get_guia(ByVal in_guia As String) As String
strCadena = "SELECT func_get_guia('" & Val(in_guia) & "')"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   get_guia = rstK(0)
Else
   get_guia = "-"
End If
End Function

Public Function get_mail(ByVal in_dni As String) As String
strCadena = "SELECT mail FROM persona where mail LIKE '%@%' and  dni='" & in_dni & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   get_mail = rstK("mail")
   
Else
   get_mail = ""
End If
End Function

Public Function get_mail_guia(ByVal in_dni As String) As String
strCadena = "SELECT mail FROM persona where mail LIKE '%@%' and  dni='" & in_dni & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   get_mail_guia = rstK("mail")
   
Else
   get_mail_guia = ""
End If
End Function

Public Function get_telefonos(ByVal in_ruc As String) As String
get_telefonos = ""
strCadena = "SELECT * FROM persona_telefono WHERE dni='" & in_ruc & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    rstK.MoveFirst
    For i = 0 To rstK.RecordCount - 1
        If i = 0 Then
            get_telefonos = rstK("telefono")
        Else
            get_telefonos = get_telefonos & " / " & rstK("telefono")
        End If
        rstK.MoveNext
    Next i
End If
KEY_TELEFONO = get_telefonos


End Function

Public Function get_telefono_sucursal(ByVal in_alm As String) As String
get_telefono_sucursal = ""
strCadena = "SELECT telefonos FROM almacen WHERE ruc='" & KEY_RUC & "' and id_alm='" & KEY_ALM & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    KEY_TELEFONO = rstK("telefonos")
End If
get_telefono_sucursal = KEY_TELEFONO
End Function
Public Function get_telefono_proveedor() As String
get_telefono_proveedor = ""
strCadena = "SELECT telefonos FROM almacen WHERE ruc='" & KEY_PROVEEDOR & "' and id_alm='00001'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    get_telefono_proveedor = rstK("telefonos")
Else
    get_telefono_proveedor = "--"
End If

End Function
Public Function get_direccion_alm(ByVal in_alm As String) As String
strCadena = "SELECT direccion FROM almacen WHERE id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    get_direccion_alm = rstK("direccion")

Else
    get_direccion_alm = ""
End If
End Function
Public Function get_moneda_documento(ByVal in_doc As String, ByVal in_serie As String) As String
strCadena = "SELECT id_moneda FROM almacen_comprobante where id_doc='" & in_doc & "' and serie='" & in_serie & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstI(strCadena)
If rstI.RecordCount > 0 Then
   get_moneda_documento = rstI("id_moneda")
Else
   get_moneda_documento = "00001"
End If

End Function

Public Function get_moneda(ByVal in_moneda As String) As String
If in_moneda = "00001" Then
    get_moneda = "SOLES"
Else
    get_moneda = "DOLARES"
End If

End Function

Public Function get_direccion(ByVal in_dni As String) As String

If in_dni = "00000000" Then
    get_direccion = KEY_DIR_PUBLIC
Else

strCadena = "SELECT direccion FROM persona where dni='" & in_dni & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_direccion = rstL("direccion")
End If
End If
End Function
Public Function get_color(ByVal in_producto As String) As String
strCadena = "SELECT c.descripcion  FROM producto p,imp_color c WHERE p.id_color=c.id_color and p.id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1 "
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   get_color = rstK("descripcion")
Else
   get_color = "SIN COLOR"
End If
End Function

Public Function get_producto(ByVal in_producto As String) As String
strCadena = "SELECT nombre_prod FROM producto where id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraTemporal(strCadena)
If rstTemporal.RecordCount > 0 Then
    get_producto = rstTemporal("nombre_prod")
Else
    get_producto = "-"
End If
End Function


Public Function get_producto_comercial(ByVal in_producto As String) As String
strCadena = "SELECT nombre_comercial FROM producto where id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraTemporal(strCadena)
If rstTemporal.RecordCount > 0 Then
    get_producto_comercial = rstTemporal("nombre_comercial")
Else
    MsgBox "Producto NO REGISTRADO :" + in_producto + Chr(13) + Chr(13) + "Verifique...", vbInformation
    get_producto_comercial = "-"
End If
End Function



Public Sub put_tracking(ByVal in_venta As String, ByVal in_estado As String, ByVal in_observacion As String)

strCadena = "CALL p_insert_tracking('" & Val(in_venta) & "','" & KEY_ALM & "','" & in_estado & "','" & KEY_USUARIO & "','" & in_observacion & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
End Sub
Public Function proceso_persona(ByVal in_dni As String, ByVal in_nombre As String, ByVal in_mail As String, ByVal in_direccion As String)

strCadena = "SELECT * FROM persona WHERE dni='" & in_dni & "'"
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
    strCadena = "UPDATE persona SET direccion='" & in_direccion & "',nombre_completo='" & in_nombre & "',mail='" & in_mail & "' WHERE dni='" & in_dni & "'"
Else
    strCadena = " call P_insert_persona('" & in_dni & "','-','-','-','" & in_nombre & "','" & in_direccion & "','','" & in_mail & "','no','no','si','no','no','no','" & KEY_RUC & "')"
End If
CnBd.Execute (strCadena)
End Function

Public Function put_matricula(ByVal in_periodo As String, ByVal in_nivel As String, ByVal in_grado As Integer, ByVal in_dni As String, ByVal in_servicio As String)
Dim in_detalle As String
strCadena = "SELECT a.precio_venta,nombre_prod FROM almacen_producto a,producto p WHERE a.id_producto=p.id_producto and p.ruc=a.ruc and  a.id_producto='" & in_servicio & "' and  a.ruc='" & KEY_RUC & "' and id_alm='" & KEY_ALM & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    in_detalle = rstL("nombre_prod") & Space(1) & "PERIODO [2017]"
    
    strCadena = "call put_matricula('" & in_periodo & "','" & in_nivel & "','" & in_grado & "','" & in_dni & "','" & KEY_USUARIO & "','" & in_servicio & "','" & rstL("precio_venta") & "','" & in_detalle & "','" & KEY_ALM & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
End If
End Function
Public Function get_cuenta_contable(ByVal in_tarjeta As String, ByVal in_forma_pago As String) As String
If in_tarjeta = "00" Then
   strCadena = "SELECT * FROM forma_pago_detalle WHERE id_detalle='" & in_forma_pago & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRstK(strCadena)
   If rstK.RecordCount > 0 Then
      get_cuenta_contable = rstK("cuenta_contable")
   Else
      get_cuenta_contable = ""
   End If
   
Else
    strCadena = "SELECT * FROM targeta WHERE id='" & in_tarjeta & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstK(strCadena)
    If rstK.RecordCount > 0 Then
        get_cuenta_contable = rstK("numero_cuenta")
    Else
        get_cuenta_contable = " "
    End If
End If
End Function

Public Function get_cuenta_contable_caja(ByVal in_forma_pago_detalle As String) As String
   strCadena = "SELECT * FROM forma_pago_detalle WHERE id_registro='" & Val(in_forma_pago_detalle) & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRstK(strCadena)
   If rstK.RecordCount > 0 Then
      get_cuenta_contable_caja = rstK("cuenta_contable")
   Else
      get_cuenta_contable_caja = ""
   End If
   
End Function


Public Function get_last() As String
strCadena = "SELECT * FROM mis_cuentas_det ORDER BY id DESC LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    get_last = rstK("id")
End If
End Function

Public Function get_cuenta(ByVal in_cuenta As String) As String
strCadena = "SELECT Descripcion FROM con_cuentacontable WHERE NroCuenta='" & in_cuenta & "' and IdEmpresaSis='" & KEY_RUC & "'"
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
    get_cuenta = rstAux("Descripcion")
Else
    get_cuenta = ""
End If
End Function

Public Function put_verifica_cuenta_contable(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_cta_compra As String, ByVal in_tipo_producto As String) As Boolean

strCadena = "SELECT * FROM con_cuentacontable WHERE NroCuenta='" & in_cta_compra & "' and IdEmpresaSis='" & KEY_RUC & "' and Ejercicio='" & Year(KEY_FECHA) & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount < 1 Then
     MsgBox "CUENTA PAGO COMPRA NO EXISTE  : " + in_cta_compra + Chr(13) + "INGRESE UNA CONFIGURACION EN PANEL DE CONTROL", vbInformation
     put_verifica_cuenta_contable = False
     Exit Function
End If

'strCadena = "SELECT servicio FROM tipo_producto WHERE  id_tipoproducto='" & in_tipo_producto & "' and ruc='" & KEY_RUC & "'"
'Call ConfiguraRstK(strCadena)
'If rstK.RecordCount > 0 Then
'        If rstK("servicio") = "si" Then
'        strCadena = "SELECT id_linea, linea FROM view_temporal_compra_conta WHERE nro_cuenta='' and  numero='" & Trim(in_numero) & "' AND id_doc='" & Trim(in_doc) & "' AND serie='" & Trim(in_serie) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "' ORDER BY id_temporal ASC"
       ' Call ConfiguraRstK(strCadena)
       ' If rstK.RecordCount > 0 Then
        '   MsgBox "CONFIGURACION CONTABLE INCOMPLETA" + Chr(13) + str(rstK("id_linea")) & Space(2) & rstK("linea"), vbInformation
        '   put_verifica_cuenta_contable = False
        'Else
 '          put_verifica_cuenta_contable = True
        'End If
'        Else
            put_verifica_cuenta_contable = True
 '       End If
'End If
End Function

Public Function get_correlativo_table(ByVal in_tabla As String, ByVal in_campo As String) As String

strCadena = "SELECT " & in_campo & " FROM " & in_tabla & " WHERE ruc='" & KEY_RUC & "' ORDER BY  " & in_campo & " LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    get_correlativo_table = Format(Val(rstL(0)) + 1, "00000")
Else
    get_correlativo_table = "00001"
End If

End Function
Public Function get_tipo_cambio_dia(ByVal in_fecha As Date, ByVal in_valor As String) As Single
Dim flag As Integer
flag = 0
Dim in_tc As Single
seleccionar_nuevo:
strCadena = "SELECT   " & in_valor & " FROM tipo_cambio WHERE fecha='" & Format(in_fecha, "YYYY-mm-dd") & "' and  id_creador='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstPP(strCadena)
If rstPP.RecordCount > 0 Then
    If rstPP(0) = 0 Then
        GoTo siguientev
    End If
    get_tipo_cambio_dia = rstPP(0)
Else
siguientev:
        strCadena = "SELECT  " & in_valor & " FROM tipo_cambio WHERE fecha<='" & Format(in_fecha, "YYYY-mm-dd") & "' and valor_compra>0 and  id_creador='" & KEY_RUC & "' ORDER BY fecha DESC LIMIT 1"
        Call ConfiguraRstPP(strCadena)
        If rstPP.RecordCount > 0 Then
           get_tipo_cambio_dia = rstPP(0)
           Exit Function
        Else
           MsgBox "No existe Registros de TIPO CAMBIO para este dia." + Chr(13) + "El sistema ingresara un TC Manual.", vbInformation
           get_tipo_cambio_dia = 3.231
           Exit Function
        End If
  
    Call get_cambio_sbs(in_fecha)
    flag = 1
    GoTo seleccionar_nuevo
End If
    
    
End Function

Public Sub get_cambio_sbs(ByVal in_fecha As Date)
On Error GoTo salir
Dim in_fecha_consulta As Date
Dim fecha_parametro As String
Dim in_compra As Single
Dim in_venta As Single

in_fecha_consulta = DateAdd("d", -1, Format(in_fecha, "YYYY-mm-dd"))

If UCase(WeekdayName(Weekday(in_fecha_consulta))) = "SÁBADO" Then
       in_fecha_consulta = DateAdd("d", -1, Format(in_fecha_consulta, "YYYY-mm-dd"))
End If
    
 If UCase(WeekdayName(Weekday(in_fecha_consulta))) = "DOMINGO" Then
        in_fecha_consulta = DateAdd("d", -1, Format(in_fecha_consulta, "YYYY-mm-dd"))
        
        If UCase(WeekdayName(Weekday(in_fecha_consulta))) = "SÁBADO" Then
            in_fecha_consulta = DateAdd("d", -1, Format(in_fecha_consulta, "YYYY-mm-dd"))
        End If
End If
        
        
        fecha_parametro = Format(Day(in_fecha_consulta), "00") & Format(Month(in_fecha_consulta), "00") & Year(in_fecha_consulta)

        strCadena = "SELECT * FROM tipo_cambio WHERE sbs='si' and  valor_compra>0 and valor_venta>0 and  fecha='" & Format(in_fecha, "YYYY-mm-dd") & "'AND id_creador='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           If rst("valor_venta") > 0 Then
                KEY_CAMBIO_VENTA = rst("valor_venta")
                KEY_CAMBIO_COMPRA = rst("valor_compra")
                KEY_CAMBIO = KEY_CAMBIO_COMPRA
                KEY_CAMBIO_LOCAL = rst("valor_local")
                Exit Sub
           End If
       
       
       End If
        
        
     
     Dim strHtml As String
     urlstr = "https://intranet2.sbs.gob.pe/api/TipoCambio/" & fecha_parametro
     'strHtml = "http://www.apilayer.net/api/live?access_key=6ed44bd69881fb4393b3ac8cc827c066&format=1&currencies=PEN"
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
     
     Dim json_r As Object
     Set json_r = JSON.parse(strHtml)
    
     If Len(Trim(strHtml)) > 66 Then ' CON DATOS RECORDAR QUE SUNAT UN DIA ANTES QUE SBC
        in_venta = json_r("resultados").Item(1).Item("venta")
        in_compra = json_r("resultados").Item(1).Item("compra")
        KEY_CAMBIO_VENTA = in_venta
        KEY_CAMBIO_COMPRA = in_compra
        KEY_CAMBIO_LOCAL = in_compra
        
        strCadena = "SELECT * FROM tipo_cambio WHERE fecha='" & Format(in_fecha, "YYYY-mm-dd") & "'AND id_creador='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
                strCadena = "INSERT INTO tipo_cambio(sbs,descripcion,fecha,valor_venta,valor_compra,valor_local,id_creador) VALUES ('si','Compra Dolar','" & Format(in_fecha, "YYYY-mm-dd") & "','" & in_venta & "','" & in_compra & "','" & in_compra & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
        Else
                strCadena = "UPDATE tipo_cambio SET sbs='si', valor_venta='" & in_venta & "',valor_compra='" & in_compra & "' WHERE id_tipocambio='" & rst("id_tipocambio") & "' and id_creador='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
        End If
    Else
        strCadena = "SELECT * FROM tipo_cambio WHERE valor_compra>0 AND valor_venta>0 and  fecha<='" & KEY_FECHA & "'AND id_creador='" & KEY_RUC & "' ORDER BY fecha DESC LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            
            strCadena = "SELECT * FROM tipo_cambio WHERE valor_compra>0 AND valor_venta>0 and fecha='" & Format(in_fecha, "YYYY-mm-dd") & "'AND id_creador='" & KEY_RUC & "' LIMIT 1"
            Call ConfiguraRstA(strCadena)
            If rstA.RecordCount < 1 Then
                strCadena = "INSERT INTO tipo_cambio(descripcion,fecha,valor_venta,valor_compra,valor_local,id_creador) VALUES ('Compra Dolar','" & Format(in_fecha, "YYYY-mm-dd") & "','" & rst("valor_venta") & "','" & rst("valor_compra") & "','" & rst("valor_local") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            Else
                strCadena = "UPDATE tipo_cambio SET valor_venta='" & rst("valor_venta") & "',valor_compra='" & rst("valor_compra") & "',valor_local='" & rst("valor_local") & "' WHERE id_tipocambio='" & rst("id_tipocambio") & "' and id_creador='" & KEY_RUC & "' LIMIT 1"
                CnBd.Execute (strCadena)
            End If
                
           KEY_CAMBIO_VENTA = rst("valor_venta")
           KEY_CAMBIO_COMPRA = rst("valor_compra")
           KEY_CAMBIO_LOCAL = rst("valor_local")
            
        End If
    End If
      KEY_CAMBIO = KEY_CAMBIO_COMPRA

            

     
     Exit Sub
salir:
        KEY_CAMBIO_VENTA = in_venta
        KEY_CAMBIO_COMPRA = in_compra
        KEY_CAMBIO = KEY_CAMBIO_COMPRA
End Sub
Public Sub get_cambio()
On Error GoTo salir
    Dim strHtml As String
    strCadena = "SELECT * FROM tipo_cambio WHERE fecha='" & KEY_FECHA & "'AND id_creador='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If rst("valor") < 1 Then
            GoTo actualizar
        Else
            KEY_CAMBIO = rst("valor")
            Exit Sub
        End If
    End If
    
actualizar:
     urlstr = "http://www.apilayer.net/api/live?access_key=6ed44bd69881fb4393b3ac8cc827c066&format=1&currencies=PEN"
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
     
     Dim p As Object
     Set p = JSON.parse(strHtml)
     Valor = InStr(1, strHtml, "USDPEN")
     Valor = Mid(strHtml, Valor + 8, 7)
     If Val(Valor) > 0 Then
        strCadena = "SELECT * FROM tipo_cambio WHERE fecha='" & KEY_FECHA & "'AND id_creador='" & KEY_RUC & "' LIMIT 0,1 "
        Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                
                KEY_CAMBIO = Valor
                strCadena = "INSERT INTO tipo_cambio(descripcion,fecha,valor,id_creador) VALUES ('Compra Dolar','" & KEY_FECHA & "','" & Valor & "','" & KEY_RUC & "')"
                Call Execute_Sql(strCadena)
            Else
                 strCadena = "UPDATE tipo_cambio SET valor='" & Val(Valor) & "' WHERE fecha='" & KEY_FECHA & "' and id_creador='" & KEY_RUC & "'"
                  Call Execute_Sql(strCadena)
                   KEY_CAMBIO = Valor
                   
            End If
     End If
     
     Exit Sub
salir:

        strCadena = "SELECT * FROM tipo_cambio WHERE fecha='" & KEY_FECHA & "'AND id_creador='" & KEY_RUC & "' LIMIT 0,1 "
        Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                strCadena = "SELECT * FROM tipo_cambio WHERE id_creador='" & KEY_RUC & "' ORDER BY fecha DESC LIMIT 0,1"
                Call ConfiguraRst(strCadena)
                If rst.RecordCount > 0 Then
                    Valor = rst("valor")
                Else
                    Valor = 3.015
                End If
                KEY_CAMBIO = Valor
                strCadena = "INSERT INTO tipo_cambio(descripcion,fecha,valor,id_creador) VALUES ('Compra Dolar','" & KEY_FECHA & "','" & Valor & "','" & KEY_RUC & "')"
                Call Execute_Sql(strCadena)
            Else
                 strCadena = "UPDATE tipo_cambio SET valor='" & Val(Valor) & "' WHERE fecha='" & KEY_FECHA & "' and id_creador='" & KEY_RUC & "'"
                  Call Execute_Sql(strCadena)
                   KEY_CAMBIO = Valor
                   
            End If

     
     Exit Sub
End Sub
Public Sub put_ingreso(ByVal in_tipo As String)


strCadena = "call put_ingreso_salida_vitekey('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_VENTANILLA & "','" & in_tipo & "','" & KEY_FECHA & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
KEY_IGV = get_igv





strCadena = "SELECT direccion FROM persona_publico WHERE ruc='" & KEY_RUC & "' AND dni='00000000' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    KEY_DIR_PUBLIC = rst("direccion")
Else
    KEY_DIR_PUBLIC = "CHICLAYO"
End If


End Sub

Private Function get_igv() As Single

Dim in_aplica As String
strCadena = "SELECT igv FROM entidad_parametros WHERE cod_unico='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
in_aplica = rstK("igv")
    
    strCadena = "SELECT * FROM entidad_igv WHERE ruc='" & KEY_RUC & "' and fecha=CURDATE() LIMIT 1"
    Call ConfiguraRstP(strCadena)
    If rstP.RecordCount < 1 Then
        strCadena = "SELECT * FROM entidad_igv WHERE ruc='" & KEY_RUC & "' ORDER BY fecha DESC LIMIT 1"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            If rstK("igv") > 0 Then
                strCadena = "INSERT INTO entidad_igv(fecha,igv,ruc)VALUES(CURDATE(),'" & rstK("igv") & "','" & KEY_RUC & "')"
                get_igv = rstK("igv")
            Else
                
                strCadena = "INSERT INTO entidad_igv(fecha,igv,ruc)VALUES(CURDATE(),'0.18','" & KEY_RUC & "')"
                get_igv = 0.18
            End If
           
       Else
        strCadena = "INSERT INTO entidad_igv(fecha,igv,ruc)VALUES(CURDATE(),'0.18','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        get_igv = 0.18
        End If
    Else
        
        get_igv = rstP("igv")
    End If

 


End Function

Public Sub activacion_temporal(ByVal in_clave As String)
Dim in_fecha As String
strCadena = "SELECT CURDATE()"
Call ConfiguraRstK(strCadena)
in_fecha = rstK(0)
strCadena = "SELECT * FROM entidad_parametros WHERE secure='" & in_clave & "' and  cod_unico='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    strCadena = "UPDATE entidad_parametros SET activacion_permanente='no',caducidad='" & Format(in_fecha, "YYYY-mm-dd") & "' WHERE cod_unico='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Call FrmFechaTrabajo.verificacion_activacion
End If
End Sub


Public Sub EliminarVentas(ByVal TipoDoc As String, ByVal serie As String, ByVal numero As String, ByVal Almacen As String)
 Dim in_venta As String
 Dim in_fecha As Date
 strCadena = "SELECT * FROM movimiento_venta WHERE id_alm='" & Almacen & "' AND id_doc='" & TipoDoc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "' ORDER BY id_venta DESC LIMIT 1"
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    in_fecha = rst("fecha_emision")
    in_venta = rst("id_venta")
    If KEY_CONTABILIDAD = "si" Then
    strCadena = "SELECT funct_validar_periodo_cierre('" & rst("id_venta") & "','" & KEY_RUC & "')"
    Call ConfiguraRstK(strCadena)
    If rstK(0) = 0 Then ' PERIODO ABIERTO
        If Format(rst("fecha_emision"), "YYYY-mm-dd") < KEY_FECHA Then
            MsgBox "Este comprobante tiene fecha anterior" + Chr(13) + Chr(13) + "Puede Anular este Comprobante." + Chr(13) + "Puede Generar Nota Credito.", vbInformation, KEY_EMPRESA
            FrmVentas.Enabled = True
            Exit Sub
        End If
   
       strCadena = "call CON_EliminarVenta('" & rst("id_venta") & "','" & KEY_USUARIO & "') "
       CnBd.Execute (strCadena)
    Else      ' PERIODO CERRADO
        
    
       MsgBox "Este comprobante tiene fecha anterior" + Chr(13) + Chr(13) + "Puede Anular este Comprobante." + Chr(13) + "Puede Generar Nota Credito.", vbInformation, KEY_EMPRESA
       FrmVentas.Enabled = True
       Exit Sub
       Exit Sub
    End If
       
    End If
           
 
 
 
 
    If Format(rst("fecha_emision"), "YYYY-mm-dd") < KEY_FECHA Then
       MsgBox "Este comprobante tiene fecha anterior" + Chr(13) + Chr(13) + "Puede Anular este Comprobante." + Chr(13) + "Puede Generar Nota Credito.", vbInformation, KEY_EMPRESA
       FrmVentas.Enabled = True
       Exit Sub
    End If
    
    strCadena = "SELECT id_doc,serie,id_alm,numero FROM movimiento_venta WHERE id_venta='" & rst("id_recibo") & "' and  ruc='" & KEY_RUC & "'"
    Call ConfiguraRstK(strCadena)
    If rstK.RecordCount > 0 Then
            in_numeron = rstK("numero")
            strCadena = "DELETE FROM  movimiento_venta  WHERE id_venta='" & rst("id_recibo") & "' AND ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
            strCadena = "SELECT numero FROM movimiento_venta WHERE id_alm='" & rstK("id_alm") & "' AND id_doc='" & rstK("id_doc") & "' AND serie='" & rstK("serie") & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
            Call ConfiguraRstL(strCadena)
            If rstL.RecordCount > 0 Then
                If Val(in_numeron) >= Val(rstL("numero")) Then
                strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(rstL("numero")) + 1, "000000") & "' WHERE id_alm='" & rstK("id_alm") & "' AND id_doc='" & rstK("id_doc") & "' AND serie='" & rstK("serie") & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                End If
            End If
    End If
    
    
    

Call update_pago_comprobante(in_venta)

strCadena = "DELETE FROM movimiento_venta WHERE id_venta='" & in_venta & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
 
    
    strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstIN(strCadena)
    If rstIN.RecordCount > 0 Then
        rstIN.MoveFirst
        For m = 0 To rstIN.RecordCount - 1
            
            strCadena = "UPDATE imp_producto_detalle SET vendido='no' WHERE id_detalle='" & rstIN("id_detalle_serie") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
             ':::::::::::ACTUALIZA KARDEX
                     If KEY_RUC = "20128836251" Then
                            Call update_kardex_Vargas_modulo_compra(rstIN("id_producto"), Format(in_fecha, "YYYY-mm-dd"))
                        Else
                            Call update_kardex_update(rstIN("id_producto"), Format(in_fecha, "YYYY-mm-dd"))
                    End If
                     '--------------FIN KARDEX
            
            strCadena = "DELETE FROM movimiento_venta_detalle WHERE id_detalle_venta='" & rstIN("id_detalle_venta") & "' LIMIT 1"
            Call EjecutaRST(strCadena)
            
            rstIN.MoveNext
        Next m
    End If
 End If
  
 
 
 
 
 
 
 strCadena = "SELECT numero FROM movimiento_venta WHERE id_alm='" & Almacen & "' AND id_doc='" & TipoDoc & "' AND serie='" & serie & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
      If Val(numero) >= Val(rst("numero")) Then
         strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(rst("numero")) + 1, "000000") & "' WHERE id_alm='" & Almacen & "' AND id_doc='" & TipoDoc & "' AND serie='" & serie & "' and ruc='" & KEY_RUC & "'"
         CnBd.Execute (strCadena)
      End If
 End If
 
 
 FrmVentas.nuevo
 
 
End Sub
Public Sub EliminarVentas_error(ByVal TipoDoc As String, ByVal serie As String, ByVal numero As String, ByVal Almacen As String)
 Dim in_venta As String
 strCadena = "SELECT * FROM movimiento_venta WHERE id_alm='" & Almacen & "' AND id_doc='" & TipoDoc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "' ORDER BY id_venta DESC LIMIT 1"
 Call ConfiguraRstIN(strCadena)
 If rstIN.RecordCount > 0 Then
    in_venta = rstIN("id_venta")
    If KEY_CONTABILIDAD = "si" Then
    strCadena = "SELECT funct_validar_periodo_cierre('" & rstIN("id_venta") & "','" & KEY_RUC & "')"
    Call ConfiguraRstC(strCadena)
    If rstc(0) = 0 Then ' PERIODO ABIERTO
        If Format(rstIN("fecha_emision"), "YYYY-mm-dd") < KEY_FECHA Then
            MsgBox "Este comprobante tiene fecha anterior" + Chr(13) + Chr(13) + "Puede Anular este Comprobante." + Chr(13) + "Puede Generar Nota Credito.", vbInformation, KEY_EMPRESA
            FrmVentas.Enabled = True
            Exit Sub
        End If
   
       strCadena = "call CON_EliminarVenta('" & rstIN("id_venta") & "','" & KEY_USUARIO & "') "
       CnBd.Execute (strCadena)
    Else      ' PERIODO CERRADO
        
    
       MsgBox "Este comprobante tiene fecha anterior" + Chr(13) + Chr(13) + "Puede Anular este Comprobante." + Chr(13) + "Puede Generar Nota Credito.", vbInformation, KEY_EMPRESA
       FrmVentas.Enabled = True
       Exit Sub
       Exit Sub
    End If
       
    End If
    
        
 
 
 
 
    If Format(rstIN("fecha_emision"), "YYYY-mm-dd") < KEY_FECHA Then
       MsgBox "Este comprobante tiene fecha anterior" + Chr(13) + Chr(13) + "Puede Anular este Comprobante." + Chr(13) + "Puede Generar Nota Credito.", vbInformation, KEY_EMPRESA
       FrmVentas.Enabled = True
       Exit Sub
    End If
    
    strCadena = "SELECT id_doc,serie,id_alm,numero FROM movimiento_venta WHERE id_venta='" & rstIN("id_recibo") & "' and  ruc='" & KEY_RUC & "'"
    Call ConfiguraRstC(strCadena)
    If rstc.RecordCount > 0 Then
            in_numeron = rstc("numero")
            strCadena = "DELETE FROM  movimiento_venta  WHERE id_venta='" & rstIN("id_recibo") & "' AND ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
            strCadena = "SELECT numero FROM movimiento_venta WHERE id_alm='" & rstc("id_alm") & "' AND id_doc='" & rstc("id_doc") & "' AND serie='" & rstc("serie") & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
            Call ConfiguraRstL(strCadena)
            If rstL.RecordCount > 0 Then
                If Val(in_numeron) >= Val(rstL("numero")) Then
                strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(rstL("numero")) + 1, "000000") & "' WHERE id_alm='" & rstc("id_alm") & "' AND id_doc='" & rstc("id_doc") & "' AND serie='" & rstc("serie") & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                End If
            End If
    End If
    
    
    
    
    strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & rstIN("id_venta") & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        rstZ.MoveFirst
        For m = 0 To rstZ.RecordCount - 1
            strCadena = "UPDATE imp_producto_detalle SET vendido='no' WHERE id_detalle='" & rstZ("id_detalle_serie") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            Call update_kardex_producto(rstZ("id_producto"), rstIN("id_alm"), rstIN("id_venta"), rstIN("afecta_factura"), rstIN("id_doc"))
            
            
            rstZ.MoveNext
        Next m
    End If
 End If
  
 
 Call update_pago_comprobante(in_venta)
 strCadena = "DELETE FROM movimiento_venta WHERE id_venta='" & in_venta & "' and ruc='" & KEY_RUC & "'"
 CnBd.Execute (strCadena)
 strCadena = "SELECT numero FROM movimiento_venta WHERE id_alm='" & Almacen & "' AND id_doc='" & TipoDoc & "' AND serie='" & serie & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
 Call ConfiguraRstIN(strCadena)
 If rst.RecordCount > 0 Then
      If Val(numero) >= Val(rstIN("numero")) Then
         strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(rstIN("numero")) + 1, "000000") & "' WHERE id_alm='" & Almacen & "' AND id_doc='" & TipoDoc & "' AND serie='" & serie & "' and ruc='" & KEY_RUC & "'"
         CnBd.Execute (strCadena)
      End If
 End If
 
 
 
 
 
End Sub

Public Sub update_kardex_producto(ByVal in_producto As String, ByVal in_alm As String, ByVal in_venta As String, ByVal afecta_factura As String, ByVal in_doc As String)
Dim nstock As Single
Dim nstock_factura As Single

strCadena = "DELETE FROM kardex WHERE id_producto='" & in_producto & "' and  id_movimiento='" & in_venta & "' and   id_alm='" & in_alm & "' and  ruc='" & KEY_RUC & "' LIMIT 1"
CnBd.Execute (strCadena)

strCadena = "SELECT sum(cantidad_real),sum(cantidad_pendiente),sum(cantidad_contable) FROM kardex WHERE id_tipo_movimiento<>'10' and  id_alm='" & in_alm & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstAux(strCadena)
                    If IsNull(rstAux(0)) = True Then
                            in_real = 0
                        Else
                            in_real = rstAux(0)
                        End If
                        
                        If IsNull(rstAux(1)) = True Then
                            in_pendiente = 0
                        Else
                            in_pendiente = rstAux(1)
                        End If
                        
    strCadena = "UPDATE almacen_producto set stock = '" & in_real & "' , `stock_contable` = '" & in_pendiente & "'  WHERE ruc = '" & KEY_RUC & "' and id_alm ='" & in_alm & "' and id_producto = '" & in_producto & "'"
    CnBd.Execute (strCadena)
                    
    strCadena = "SELECT sum(cantidad_real),sum(cantidad_factura),sum(cantidad_contable) FROM kardex WHERE id_tipo_movimiento='10' and  id_alm='" & in_alm & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstAux(strCadena)
                        If IsNull(rstAux(1)) = True Then
                            in_contable = 0
                        Else
                            in_contable = rstAux(1)
                        End If
                    
                    strCadena = "UPDATE almacen_producto set  stock_factura='" & in_contable & "'   WHERE ruc = '" & KEY_RUC & "' and id_alm ='" & in_alm & "' and id_producto = '" & in_producto & "'"
                    CnBd.Execute (strCadena)



End Sub
Public Sub update_pago_comprobante(ByVal in_venta As String)

If Val(in_venta) > 0 Then
    strCadena = "UPDATE mis_cuentas_det SET id_cuenta='0' WHERE id_tipo_movimiento='00001' and id_venta='" & Val(in_venta) & "' and  ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
End If

End Sub

Public Sub load_usuario(ByVal in_combo As DataCombo, ByVal in_dni As String, ByVal in_criterio As String)

If Trim(in_criterio) = "" Then
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM persona where dni='" & in_dni & "'"
Else
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad where nombre_completo LIKE '%" & in_criterio & "%' and ruc='" & KEY_RUC & "' and id_personal='si' order by nombre_completo"
End If

Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(in_combo)

End Sub

Public Function get_telefono(ByVal in_dni As String) As String
strCadena = "SELECT celular FROM persona Where dni='" & in_dni & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    get_telefono = rstT("celular")
Else
    get_telefono = ""
End If

End Function
Public Sub delete_seguro(ByVal in_detalle As String)

On Error GoTo errorhandler
    
    strCadena = "DELETE  FROM seguro_medico_detalle WHERE id_detalle='" & Trim(in_detalle) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    
    
    Exit Sub


errorhandler:



End Sub
Public Function get_periodo_actual(ByVal in_fecha As Date)
On Error GoTo salir

strCadena = "SELECT id FROM con_periodo where Ejercicio='" & Year(in_fecha) & "' and Mes='" & Month(in_fecha) & "' LIMit 1"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
   get_periodo_actual = rstT("id")
Else
   get_periodo_actual = 0
End If


Exit Function
salir:

End Function



Public Function validar_periodo(ByVal in_mes As Integer, ByVal in_anio As Integer, ByVal in_periodo As String) As Boolean
strCadena = "SELECT Mes,Ejercicio FROM con_periodo where id='" & in_periodo & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
   If in_anio < rstT("Ejercicio") Then
      validar_periodo = True
        
  Else
    If in_mes <= rstT("Mes") And in_anio <= rstT("Ejercicio") Then
            validar_periodo = True
        Else
            validar_periodo = False
        End If
  
  End If
   
Else
   validar_periodo = False
End If
End Function
Public Function validar_periodo_nocontable(ByVal in_fecha As String) As Boolean

If Month(CVDate(in_fecha)) <> Month(KEY_FECHA) Then
   validar_periodo_nocontable = False
Else
    validar_periodo_nocontable = True
End If

End Function


Public Function put_delete_servicio(ByVal in_servicio As String) As Boolean
On Error GoTo Errorhanlder
strCadena = "DELETE FROM persona_plan_servicio_ii WHERE id='" & Val(in_servicio) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
put_delete_servicio = True
Exit Function
Errorhanlder:
put_delete_servicio = False

End Function

Public Function put_anular_cambio(ByVal in_cambio As String) As Boolean
    On Error GoTo Errorhanlder
    strCadena = "UPDATE cambio_aceite SET anulado='si' WHERE id_cambio='" & Val(in_cambio) & "'"
    CnBd.Execute (strCadena)
    put_anular_cambio = True
    Exit Function
Errorhanlder:
put_anular_cambio = False
End Function


Public Function get_dar_baja(ByVal in_venta As String) As Boolean
    
    strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & Val(in_venta) & "' "
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
         If get_firma_online(rstL("id_doc"), rstL("serie")) = "si" And Len(Trim(rstL("sunat_key"))) > 1 Then
                If DateDiff("d", KEY_FECHA, rst("fecha_emision")) >= 7 Then
                    get_dar_baja = True
                Else
                    get_dar_baja = False
                End If
         Else
            get_dar_baja = False
         End If
    End If
    
    
    
    
    
    

End Function


Public Function procesar_transaccion(ByVal in_alm As String, ByVal in_cta_origen As String, ByVal in_fecha As Date, ByVal in_tipo As String, ByVal in_proveedor As String, ByVal in_proveedor_des As String, ByVal in_observacion As String _
, ByVal in_monto As Single, ByVal in_cta_destino As String, ByVal in_venta As String, ByVal in_compra As String, ByVal in_documento As String, ByVal in_tc As Single, ByVal in_operacion As String, ByVal in_forma_pago As String, ByVal in_tipo_flujo As String, ByVal in_moneda As String, ByVal in_dni_save As String, ByVal in_ruc As String) As Double
Dim in_mis_cuentas_det As String
procesar_transaccion = False
 

If KEY_PAIS = KEY_PERU Then ' verifico que es de Peru
   'strCadena = "call sp_procesar_transaccion_caja_demo_ii('" & in_alm & "','" & in_cta_origen & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & in_tipo & "','" & in_proveedor & "','" & in_observacion & "','" & in_monto & "','" & in_cta_destino & "','" & in_tc & "','" & in_operacion & "','" & in_dni_save & "','" & in_venta & "','" & in_compra & "','" & in_proveedor_des & "','" & in_documento & "','" & in_forma_pago & "','" & in_moneda & "','" & in_tipo_flujo & "','" & in_ruc & "')"
                           strCadena = "call sp_insertar_transaccion_v2('" & in_alm & "','" & in_cta_origen & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & in_tipo & "','" & in_proveedor & "','" & in_observacion & "','" & in_monto & "','" & in_cta_destino & "','" & in_tc & "','" & in_operacion & "','" & in_dni_save & "','" & in_venta & "','" & in_compra & "','" & in_proveedor_des & "','" & in_documento & "','" & in_forma_pago & "','" & in_moneda & "','" & in_tipo_flujo & "','" & in_ruc & "')"
Else
   strCadena = "call sp_procesar_transaccion_caja_demo_internacional('" & in_alm & "','" & in_cta_origen & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & in_tipo & "','" & in_proveedor & "','" & in_observacion & "','" & in_monto & "','" & in_cta_destino & "','" & in_tc & "','" & in_operacion & "','" & in_dni_save & "','" & in_venta & "','" & in_compra & "','" & in_proveedor_des & "','" & in_documento & "','" & in_forma_pago & "','" & in_moneda & "','" & in_tipo_flujo & "','" & in_ruc & "')"
End If

Call ConfiguraRstPP(strCadena)






in_mis_cuentas_det = rstPP("in_origen")

procesar_transaccion = in_mis_cuentas_det


End Function
Public Function procesar_transaccion_retencion(ByVal in_venta As String, ByVal in_alm As String, ByVal in_fecha As Date, ByVal in_proveedor As String, ByVal in_serie_reten As String _
, ByVal in_numero_reten As String, ByVal in_monto As Single, ByVal in_tc As Single, ByVal in_moneda As String, ByVal in_dni_save As String, ByVal in_ruc As String) As Double

'***** Insertar Retencion
strCadena = "call insert_retencion('" & Format(in_fecha, "YYYY-mm-dd") & "','" & in_moneda & "','" & in_serie_reten & "','" & in_numero_reten & "','" & in_tc & "','" & in_monto & "','" & in_dni_save & "','" & in_proveedor & "','" & Val(in_venta) & "','" & KEY_RUC & "')"
Call ConfiguraRstK(strCadena)

'***** Insertar Detalle retencion
strCadena = "call CON_InsertaAsientoRetencion_vitekey('" & rstK("in_retencion") & "')"
CnBd.Execute (strCadena)


strCadena = "call sp_procesar_transaccion_retencion('" & in_alm & "','" & in_proveedor & "','" & get_persona(in_proveedor) & "','" & in_doc_retencion & "','" & in_tc & "','" & in_venta & "','" & in_moneda & "','" & in_monto & "','" & KEY_RUC & "')"
Call ConfiguraRstPP(strCadena)



End Function

Public Function procesar_transaccion_caja(ByVal in_alm As String, ByVal in_cta_origen As String, ByVal in_fecha As Date, ByVal in_tipo As String, ByVal in_proveedor As String, ByVal in_proveedor_des As String, ByVal in_observacion As String _
, ByVal in_monto As Single, ByVal in_cta_destino As String, ByVal in_venta As String, ByVal in_compra As String, ByVal in_documento As String, ByVal in_tc As Single, ByVal in_operacion As String, ByVal in_forma_pago As String, ByVal in_tipo_flujo As String, ByVal in_moneda As String, ByVal in_dni_save As String, ByVal in_ruc As String) As Double

If KEY_PAIS = KEY_PERU Then
    '          strCadena = "call sp_insertar_transaccion_v2('" & in_alm & "','" & in_cta_origen & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & in_tipo & "','" & in_proveedor & "','" & in_observacion & "','" & in_monto & "','" & in_cta_destino & "','" & in_tc & "','" & in_operacion & "','" & in_dni_save & "','" & in_venta & "','" & in_compra & "','" & in_proveedor_des & "','" & in_documento & "','" & in_forma_pago & "','" & in_moneda & "','" & in_tipo_flujo & "','" & in_ruc & "')"
    strCadena = "call sp_procesar_transaccion_caja_premiun('" & in_alm & "','" & in_cta_origen & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & in_tipo & "','" & in_proveedor & "','" & in_observacion & "','" & in_monto & "','" & in_cta_destino & "','" & in_tc & "','" & in_operacion & "','" & in_dni_save & "','" & in_venta & "','" & in_compra & "','" & in_proveedor_des & "','" & in_documento & "','" & in_forma_pago & "','" & in_moneda & "','" & in_tipo_flujo & "','" & in_ruc & "')"
Else
   strCadena = "call sp_procesar_transaccion_caja_premiun_internacional('" & in_alm & "','" & in_cta_origen & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & in_tipo & "','" & in_proveedor & "','" & in_observacion & "','" & in_monto & "','" & in_cta_destino & "','" & in_tc & "','" & in_operacion & "','" & in_dni_save & "','" & in_venta & "','" & in_compra & "','" & in_proveedor_des & "','" & in_documento & "','" & in_forma_pago & "','" & in_moneda & "','" & in_tipo_flujo & "','" & in_ruc & "')"
End If
Call ConfiguraRstP(strCadena)
procesar_transaccion_caja = rstP(0)


End Function
Private Function get_comprobante_relacionado(ByVal in_venta As String) As String
strCadena = "SELECT id_comprobante FROM movimiento_venta WHERE id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   get_comprobante_relacionado = rstLocal("id_comprobante")
Else
    get_comprobante_relacionado = 0
End If


End Function
Public Function procesar_transaccion_venta(ByVal in_detalle_forma_pago As String, ByVal in_alm As String, ByVal in_cta_origen As String, ByVal in_fecha As Date, ByVal in_tipo As String, ByVal in_proveedor As String, ByVal in_proveedor_des As String, ByVal in_observacion As String _
, ByVal in_monto As Single, ByVal in_cta_destino As String, ByVal in_venta As String, ByVal in_compra As String, ByVal in_documento As String, ByVal in_tc As Single, ByVal in_operacion As String, ByVal in_forma_pago As String, ByVal in_tipo_flujo As String, ByVal in_dni_save As String, ByVal in_doc As String, ByVal in_ruc As String) As Double
procesar_transaccion_venta = False
Dim in_relacionado As String

strCadena = "call sp_procesar_transaccion_caja_venta('" & in_alm & "','" & in_cta_origen & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & in_tipo & "','" & in_proveedor & "','" & in_observacion & "','" & in_monto & "','" & in_cta_destino & "','" & in_tc & "','" & in_operacion & "','" & in_dni_save & "','" & in_venta & "','" & in_compra & "','" & in_proveedor_des & "','" & in_documento & "','" & in_tipo_flujo & "','" & in_forma_pago & "','" & in_ruc & "')"
Call ConfiguraRstPP(strCadena)
procesar_transaccion_venta = rstPP(0)

End Function


Public Function procesar_transaccion_egreso(ByVal in_detalle_forma_pago As String, ByVal in_alm As String, ByVal in_cta_origen As String, ByVal in_fecha As Date, ByVal in_tipo As String, ByVal in_proveedor As String, ByVal in_proveedor_des As String, ByVal in_observacion As String _
, ByVal in_monto As Single, ByVal in_cta_destino As String, ByVal in_venta As String, ByVal in_compra As String, ByVal in_documento As String, ByVal in_tc As Single, ByVal in_operacion As String, ByVal in_forma_pago As String, ByVal in_tipo_flujo As String, ByVal in_dni_save As String, ByVal in_ruc As String) As Boolean
procesar_transaccion_egreso = False


strCadena = "call sp_procesar_transaccion_caja_venta('" & in_alm & "','" & in_cta_origen & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & in_tipo & "','" & in_proveedor & "','" & in_observacion & "','" & in_monto & "','" & in_cta_destino & "','" & in_tc & "','" & in_operacion & "','" & in_dni_save & "','" & in_compra & "','" & in_venta & "','" & in_proveedor_des & "','" & in_documento & "','" & in_tipo_flujo & "','" & in_forma_pago & "','" & in_ruc & "')"
Call ConfiguraRstPP(strCadena)


procesar_transaccion_egreso = True

End Function



Public Function get_cuenta_caja(ByVal in_cuenta As String) As String
strCadena = "SELECT id_cuenta_caja FROM mis_cuentas WHERE cuenta_ctble='" & in_cuenta & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_cuenta_caja = rstL("id_cuenta_caja")
Else
   get_cuenta_caja = 0
End If
End Function
Public Function get_cuenta_contable_cuenta(ByVal in_cuenta As String) As String
strCadena = "SELECT cuenta_ctble FROM mis_cuentas WHERE id_cuenta='" & in_cuenta & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_cuenta_contable_cuenta = rstL("cuenta_ctble")
Else
   get_cuenta_contable_cuenta = 0
End If
End Function
Public Function get_nombre_cuenta(ByVal in_cuenta As String) As String
strCadena = "SELECT * FROM view_cuenta_descripcion WHERE id_cuenta='" & in_cuenta & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_nombre_cuenta = rstL("descripcion")
Else
   get_nombre_cuenta = 0
End If
End Function
Public Function get_cuenta_pago(ByVal in_registro As String) As String
On Error GoTo salir
strCadena = "SELECT id_cuenta_caja FROM forma_pago_detalle WHERE id_registro='" & in_registro & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_cuenta_pago = rstL("id_cuenta_caja")
Else
   get_cuenta_pago = 0
End If
Exit Function
salir:

End Function


Public Function put_persona(ByVal in_dni As String)
                
                strCadena = "put_entidad_empresa('" & in_dni & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
End Function
Public Function anular_solicitud(ByVal in_solicitud As String) As Boolean
On Error GoTo salir
strCadena = "UPDATE solicitud_dinero SET anulado='si' where id_solicitud='" & Val(in_solicitud) & "'"
CnBd.Execute (strCadena)
anular_solicitud = True
Exit Function
salir:
anular_solicitud = False

End Function
Public Function eliminar_proyecto(ByVal in_proyecto As String) As Boolean
On Error GoTo salir
strCadena = "DELETE FROM  mis_proyectos WHERE id_proyecto='" & Val(in_proyecto) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
eliminar_proyecto = True
Exit Function
salir:
eliminar_proyecto = False

End Function
Public Function get_unidad_producto(ByVal in_producto As String) As String

strCadena = "SELECT id_unidad FROM producto where id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_unidad_producto = rstL("id_unidad")
Else
    get_unidad_producto = "00001"
End If


End Function

Public Function get_unidad_descripcion(ByVal in_unidad As String) As String
strCadena = "SELECT descripcion FROM unidad where id_und='" & in_unidad & "' and id_usu='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_unidad_descripcion = rstL("descripcion")
Else
    get_unidad_descripcion = "UNIDAD"
End If
End Function


Public Function get_unidad_abrev(ByVal in_unidad As String) As String
strCadena = "SELECT abreviatura FROM unidad where id_und='" & in_unidad & "' and id_usu='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_unidad_abrev = rstL("abreviatura")
Else
    get_unidad_abrev = "UND"
End If
End Function

Public Sub put_interes_comprobante(ByVal in_incremento As Single, ByVal in_doc As String, ByVal in_serie As String)
Dim in_detalle As String
If in_incremento > 0 Then
in_detalle = get_producto(KEY_PRODUCTO_INTERES)


strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_alm,id_doc,id_serie,id_producto,cantidad,precio,total,peso,detalle,dni_save,servicio,obsequio,costo) VALUES " & _
"('" & KEY_RUC & "','" & get_unidad_producto(KEY_PRODUCTO_INTERES) & "','" & KEY_ALM & "','" & in_doc & "','" & in_serie & "','" & KEY_PRODUCTO_INTERES & "','1'," & _
"'" & in_incremento & " ','" & in_incremento & "','0','" & in_detalle & "','" & KEY_USUARIO & "','si','no','" & get_precio_costo(codigoP) & "')"
CnBd.Execute (strCadena)
        End If
        
End Sub
Public Sub put_interes_venta_credito(ByVal in_monto As Single, ByVal in_total As Double, ByVal in_monto_pagado As Double, ByVal in_porcentaje As Single)
Dim in_precio As Single
Dim in_incremento As Single
Dim incremento_total  As Single
Dim in_doc As String
Dim in_serie As String
On Error GoTo errorhandler

strCadena = "call sp_interes_venta_credito('" & KEY_ALM & "','" & KEY_USUARIO & "','" & in_porcentaje & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)




strCadena = "SELECT * FROM temporal_ventas WHERE id_alm='" & KEY_ALM & "' and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    rstK.MoveFirst
    incremento_total = 0
    in_doc = rstK("id_doc")
    in_serie = rstK("id_serie")
    incremento_total = (Val(in_monto)) * (in_porcentaje / 100)
    strCadena = "DELETE  FROM temporal_ventas WHERE id_producto='" & KEY_PRODUCTO_INTERES & "' and  id_alm='" & KEY_ALM & "' and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    Call put_interes_comprobante(incremento_total, in_doc, in_serie)
    
End If



Exit Sub
errorhandler:
strCadena = "call sp_interes_venta_credito('" & KEY_ALM & "','" & KEY_USUARIO & "','0','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

End Sub



Public Function get_telefono_last(ByVal in_dni As String) As String

strCadena = "SELECT telefono FROM persona_telefono WHERE dni='" & in_dni & "' ORDER BY id_telefono DESC LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   get_telefono_last = rstK("telefono")
Else
   get_telefono_last = ""
End If

End Function

Public Function get_cuotas(ByVal in_venta As String) As Boolean

strCadena = "SELECT * FROM movimiento_venta WHERE cuotas>0 and  id_venta='" & Val(in_venta) & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   get_cuotas = True
Else
    get_cuotas = False
End If

End Function

Public Function get_comprobante_des(ByVal in_doc As String) As String
strCadena = "SELECT doc_des FROM comprobantes WHERE id_doc='" & in_doc & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   get_comprobante_des = rstK("doc_des")
Else
   get_comprobante_des = "-"
End If
End Function
Public Function get_id_registro_forma_pago(ByVal forma_pago As String, ByVal in_forma_pago As String) As String
strCadena = "SELECT id_registro FROM forma_pago_detalle WHERE id_detalle='" & forma_pago & "' and id='" & in_forma_pago & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    get_id_registro_forma_pago = rstK("id_registro")
Else
    get_id_registro_forma_pago = "01"
End If
End Function

Public Sub put_pagar_servicio_cobranza(ByVal in_venta As String, ByVal in_monto As Single, ByVal in_ruc As String)
    On Error GoTo salir
    strCadena = "put_pago_servicio_cobranza('" & Val(in_venta) & "','" & in_monto & "','" & in_ruc & "')"
    CnBd.Execute (strCadena)
    Exit Sub
salir:
End Sub

Public Function get_documento_referencia(ByVal in_venta As String) As String
strCadena = "SELECT documento FROM movimiento_venta WHERE id_venta='" & Val(in_venta) & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    get_documento_referencia = rstK("documento")
Else
    get_documento_referencia = "-"
End If


End Function



Public Function get_documento_venta(ByVal in_venta As String) As String
strCadena = "SELECT id_venta FROM movimiento_transferencia WHERE id_transferencia='" & Val(in_venta) & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
strCadena = "SELECT * FROM movimiento_venta where id_venta='" & Val(rstK("id_venta")) & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   get_documento_venta = rstK("documento")
Else
   get_documento_venta = "-"
End If

End If
End Function
Public Function get_marca_2(ByVal in_marca As String) As String
strCadena = "SELECT marca FROM persona_transporte WHERE id='" & Val(in_marca) & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_marca_2 = rstL("marca")
End If
End Function
Public Sub put_bloqueo()
  
        If DateDiff("d", KEY_FECHA, KEY_FECHA_CORTE) > 0 Then
                Call SlideForm(frmNotificacion, mostrar, 200, 5, Format(KEY_FECHA_CORTE, "dd-mm-YYYY"), "CONTACTE CON EL ADMIN DEL SISTEMA.")
                'MDIFrmPrincipal.timer_cobranza.Enabled = False
                MDIFrmPrincipal.Enabled = True
            Else
                Call SlideForm(frmNotificacion, mostrar, 200, 5, Format(KEY_FECHA_CORTE, "dd-mm-YYYY"), "CONTACTE CON EL ADMIN DEL SISTEMA.")
                MDIFrmPrincipal.Enabled = False
                Exit Sub
        End If
   
    

Exit Sub
End Sub

Public Sub get_cobranza()
On Error GoTo salir
Dim in_resultado() As String
Dim in_fecha_servidor As String
strCadena = "SELECT function_cobranza('" & KEY_RUC & "'),CURDATE())"
Call ConfiguraRstC(strCadena)

If Format(rstc(1), "YYYY-mm-dd") < "2017-12-01" Then
   Exit Sub
End If
If IsNull(rstc(0)) = False Then
    in_resultado = Split(rstc(0), ":")
    If UBound(in_resultado) > 0 Then
       If IsNull(in_resultado(0)) = False Then
            
            Call SlideForm(frmNotificacion, mostrar, 200, 5, Format(in_resultado(1), "YYYY-mm-dd"), "DEUDA A LA FECHA")
           ' MDIFrmPrincipal.timer_cobranza.Enabled = False
            If in_resultado(2) = "no" Then
                If DateDiff("d", KEY_FECHA, Format(in_resultado(1), "YYYY-mm-dd")) < 0 Then
                    ' MDIFrmPrincipal.timer_cobranza.Enabled = False
                     MDIFrmPrincipal.Enabled = False
                End If
            End If
       End If
    End If
 Else
    strCadena = "SELECT * FROM entidad_parametros WHERE cod_unico='" & KEY_RUC & "'"
    Call ConfiguraRstC(strCadena)
    If rstc.RecordCount > 0 Then
        If rstc("activacion_permanente") = "no" Then
            If Format(rstc("caducidad"), "YYYY-mm-dd") <= KEY_FECHA Then
                Call SlideForm(frmNotificacion, mostrar, 200, 5, Format(rstc("caducidad"), "YYYY-mm-dd"), "BLOQUEADO.")
               ' MDIFrmPrincipal.timer_cobranza.Enabled = False
                MDIFrmPrincipal.Enabled = True
            End If
        End If
    End If
    
End If
Exit Sub
salir:
End Sub
Public Function get_producto_habilitado(ByVal in_habilitado As String) As Boolean
    
    If in_habilitado = "si" Then
       get_producto_habilitado = True
    Else
       get_producto_habilitado = False
       MsgBox "Producto Desabilitado [NO ACTIVO]" + Chr(13) + Chr(13) + "Configure su CONFIGURACION", vbInformation, KEY_VENDEDOR
    End If

End Function

Public Function get_cuenta_contable_producto(ByVal in_producto As String) As String

strCadena = "SELECT cta_contable FROM producto WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   get_cuenta_contable_producto = rstK("cta_contable")
Else
   get_cuenta_contable_producto = ""
End If



End Function

Public Sub put_correlativo_venta(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String)

strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(in_numero) + 1, "000000") & "' WHERE  id_doc='" & Trim(in_doc) & "' AND serie='" & Trim(in_serie) & "'  AND ruc='" & KEY_RUC & "'"
Call ConfiguraRstlocal(strCadena)
End Sub



Public Function get_periodo_detalle(ByVal in_periodo As String, ByVal in_compra As String) As String
strCadena = "SELECT * FROM con_periodo WHERE id='" & in_periodo & "'"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   get_periodo_detalle = "[ " & rstP("Ejercicio") & "-" & Format(rstP("Mes"), "00") & " ] :" & in_compra
Else
   get_periodo_detalle = ""
End If

End Function



Public Function get_periodo_descripcion(ByVal in_periodo As String) As String
strCadena = "SELECT * FROM con_periodo WHERE id='" & in_periodo & "'"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   get_periodo_descripcion = "[ " & rstP("Ejercicio") & "-" & Format(rstP("Mes"), "00") & " ] "
Else
   get_periodo_descripcion = ""
End If

End Function

Public Sub cerrar_caja(ByVal in_cuenta As Double)
strCadena = "SELECT * FROM  view_mis_cuentas WHERE id_cuenta='" & in_cuenta & "' and Ejercicio='" & Year(KEY_FECHA) & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Call generar_recibo_ingreso(rst("saldo"), in_cuenta)
End If
End Sub


Public Sub generar_recibo_ingreso(ByVal in_monto As Double, ByVal in_cuenta As Double)
Dim in_numero As String
Dim in_serie As String

                    KEY_VENCIMIENTO = KEY_FECHA
                    id_tipo_factura = "00001"
                    igv = "si"
                    dfac = "no"
                    
                    
                    
                    horario = Format(Time, "hh:mm")
                    If horario >= "07:00" And horario <= "13:00" Then
                        turno = "M"
                    Else
                        turno = "T"
                    End If
                    
                    strCadena = "SELECT * FROM mis_cuentas WHERE id_cuenta='" & in_cuenta & "'"
                    Call ConfiguraRst(strCadena)
                    If rst.RecordCount > 0 Then
                        in_moneda = rst("id_moneda")
                    End If
                    strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0108' and id_alm='" & KEY_ALM & "' and id_moneda='" & in_moneda & "' and ruc='" & KEY_RUC & "'  limit 1"
                    Call ConfiguraRst(strCadena)
                    If rst.RecordCount > 0 Then
                            in_numero = rst("numero")
                            in_serie = rst("serie")
                            in_observacion = "SALDO INICIAL: " & KEY_FECHA
                            Documento = "RECIBO INGRESO" & ":" & rst("serie") & "-" & rst("numero")
                            strCadena = "P_insert_venta('0108','" & KEY_ALM & "','01','" & rst("id_moneda") & "','no'," & _
                            "'" & rst("serie") & "','" & rst("numero") & "','" & KEY_RUC & "','" & KEY_EMPRESA & "','0','0','0','" & in_monto & "','0'," & _
                            "'" & in_monto & "','0','" & KEY_FECHA & "','" & KEY_VENCIMIENTO & "','" & id_tipo_factura & "','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & KEY_CAMBIO_COMPRA & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                            Call ConfiguraRstP(strCadena)
                            id_venta = rstP(0)
                            
                            strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES " & _
                            "('" & id_venta & "','00000','" & in_observacion & "','-','1','" & in_monto & "','0','" & in_monto & "','" & KEY_RUC & "')"
                            CnBd.Execute (strCadena)
                            
                            strCadena = "INSERT INTO movimiento_venta_monto(id_venta,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,ruc)VALUES " & _
                            "('" & id_venta & "','01','" & get_id_registro_forma_pago("01", "01") & "','" & in_monto & "','" & in_monto & "','00','--','-','" & KEY_RUC & "')"
                            CnBd.Execute (strCadena)
                            
                            strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(in_numero) + 1, "000000") & "' WHERE id_doc='0108' AND serie='" & Trim(in_serie) & "' AND ruc='" & KEY_RUC & "'"
                            CnBd.Execute (strCadena)
                    End If
                    
End Sub
Public Function get_id_producto_grupo(ByVal in_empresaii As String) As String
Dim in_productoA As Double
Dim in_productoB As Double

strCadena = "SELECT id_producto FROM producto WHERE ruc='" & KEY_RUC & "' ORDER BY id_producto DESC LIMIT 1"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   in_productoA = Val(rstP("id_producto")) + 1
Else
   in_productoA = 1
End If

strCadena = "SELECT id_producto FROM producto WHERE ruc='" & in_empresaii & "' ORDER BY id_producto DESC LIMIT 1"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
    in_productoB = Val(rstP("id_producto")) + 1
Else
   in_productoB = 1
End If

If in_productoA <> in_productoB Then
   
   If in_productoA = in_productoB Then
      get_id_producto_grupo = Format(in_productoA, "00000")
   End If
   
   If in_productoA > in_productoB Then
      get_id_producto_grupo = Format(in_productoA, "00000")
   End If
   
   If in_productoB > in_productoA Then
      get_id_producto_grupo = Format(in_productoB, "00000")
   End If
Else
    get_id_producto_grupo = in_productoA
   
End If


End Function
Public Function put_revertir(ByVal in_origen As String) As Boolean
strCadena = "SELECT id_compra, id_venta,monto FROM mis_cuentas_det WHERE id='" & Val(in_origen) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    If rst("id_compra") > 0 Then
        strCadena = "UPDATE movimiento_compra SET saldo=saldo+'" & rst("monto") & "' WHERE id_compra='" & rst("id_compra") & "'"
        CnBd.Execute (strCadena)
    End If
    
    If rst("id_venta") > 0 Then
        strCadena = "UPDATE movimiento_venta SET saldo=saldo+'" & rst("monto") & "' WHERE id_venta='" & rst("id_venta") & "'"
        CnBd.Execute (strCadena)
    End If
    
    
End If

strCadena = "DELETE FROM mis_cuentas_det WHERE id='" & Val(in_origen) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

put_revertir = True

End Function
Public Function get_moneda_movimiento(ByVal in_venta As String, ByVal in_compra As String) As String
If Val(in_venta) > 0 Then
    strCadena = "SELECT id_moneda FROM movimiento_venta WHERE id_venta='" & in_venta & "'"
    Call ConfiguraRstP(strCadena)
    If rstP.RecordCount > 0 Then
        get_moneda_movimiento = rstP("id_moneda")
    End If
    Exit Function
End If

If Val(in_compra) > 0 Then
    strCadena = "SELECT id_moneda FROM movimiento_compra WHERE id_compra='" & in_compra & "'"
    Call ConfiguraRstP(strCadena)
    If rstP.RecordCount > 0 Then
        get_moneda_movimiento = rstP("id_moneda")
    End If
    Exit Function
End If

End Function

Public Function get_moneda_cuentas(ByVal in_cuenta As String) As String
strCadena = "SELECT id_moneda FROM mis_cuentas WHERE id_cuenta='" & in_cuenta & "'"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   get_moneda_cuentas = rstP("id_moneda")
End If
End Function

Public Function put_revertir_ultimate(ByVal in_origen As String) As Boolean
Dim in_moneda_factura As String
Dim in_moneda_cuenta As String
Dim in_compra As String
Dim in_movimiento As String
Dim in_destino As Double
Dim in_origen_cuenta As Double

in_movimiento = 0
in_destino = 0
in_origen_cuenta = 0

strCadena = "SELECT id_compra, id_venta,monto,id_cuenta,tc,id_origen,id_tipo_movimiento,id_destino FROM mis_cuentas_det WHERE id='" & Val(in_origen) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    in_destino = rst("id_destino")
    If rst("id_tipo_movimiento") = "00003" Then
        in_origen_cuenta = rst("id_origen")
    End If
    
    
    If rst("id_compra") > 0 Then
        in_movimiento = rst("id_compra")
        in_moneda_factura = get_moneda_movimiento(rst("id_venta"), rst("id_compra"))
        in_moneda_cuenta = get_moneda_cuentas(rst("id_cuenta"))
        
        If in_moneda_factura = in_moneda_cuenta Then
            in_monto = rst("monto")
        Else
            If in_moneda_factura = "00002" Then
               in_monto = rst("monto") / rst("tc")
            Else
               in_monto = rst("monto") * rst("tc")
            End If
        End If
    End If
    
    
    
    
    
    
    
    If rst("id_venta") > 0 Then
        
        in_movimiento = rst("id_venta")
        in_moneda_factura = get_moneda_movimiento(rst("id_venta"), rst("id_compra"))
        in_moneda_cuenta = get_moneda_cuentas(rst("id_cuenta"))
        
        If in_moneda_factura = in_moneda_cuenta Then
            in_monto = rst("monto")
        Else
            If in_moneda_factura = "00002" Then
               in_monto = rst("monto") '* rst("tc")
            Else
               in_monto = rst("monto") '/ rst("tc")
            End If
        End If
        
        strCadena = "UPDATE movimiento_venta SET tc_local='0',orden_compra='',id_orden_salida='',operacion='',pendiente='si',observacion='-' WHERE id_venta='" & rst("id_venta") & "' LIMIT 1"
        CnBd.Execute (strCadena)
    End If
    
    
    
    
    
End If

If Val(in_origen) > 0 Then
    strCadena = "DELETE FROM  mis_cuentas_det where id='" & Val(in_origen) & "' and ruc='" & KEY_RUC & "' LIMIT 1 "
    CnBd.Execute (strCadena)
    If in_destino > 0 Then
        strCadena = "DELETE FROM  mis_cuentas_det where id='" & Val(in_destino) & "' and ruc='" & KEY_RUC & "' LIMIT 1 "
        CnBd.Execute (strCadena)
    End If
    
    
    strCadena = "SELECT id_detalle FROM mis_cuentas_det_detalle WHERE id_cuenta_det='" & Val(in_origen) & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        strCadena = "UPDATE movimiento_venta SET anulado='si',total=0 WHERE id_doc IN ('0054','0097') and  id_venta='" & rst("id_detalle") & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
        strCadena = "SELECT id_comprobante FROM movimiento_venta WHERE id_doc IN ('0054','0097') and  id_venta='" & rst("id_detalle") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstA(strCadena)
        If rstA.RecordCount > 0 Then
            strCadena = "UPDATE movimiento_venta SET pendiente='si' WHERE   id_venta='" & rst("id_detalle") & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        End If
        
        
        
        
        '- update en mis cuentas det detalle
        strCadena = "DELETE FROM  mis_cuentas_det_detalle where id_cuenta_det='" & Val(in_origen) & "'"
        CnBd.Execute (strCadena)
    End If
End If

If in_origen_cuenta > 0 Then
   strCadena = "DELETE FROM  mis_cuentas_det where id_origen='" & Val(in_origen_cuenta) & "' and id_tipo_movimiento='00003' AND ruc='" & KEY_RUC & "' "
   CnBd.Execute (strCadena)
End If



   
put_revertir_ultimate = True

End Function
Public Function get_forma_pago_detalle(ByVal in_registro As String) As String
If Val(in_registro) > 0 Then
    strCadena = "SELECT funct_get_formapago('" & Val(in_registro) & "')"
    Call ConfiguraRstPP(strCadena)
    get_forma_pago_detalle = rstPP(0)
End If
End Function


Public Function get_forma_pago_detalle_contado() As String

    strCadena = "SELECT get_forma_pago_contado('" & KEY_RUC & "')"
    Call ConfiguraRstPP(strCadena)
    get_forma_pago_detalle_contado = rstPP(0)

End Function


Public Function get_moneda_cuenta(ByVal in_cuenta As String) As String

strCadena = "SELECT id_moneda FROM mis_cuentas WHERE id_cuenta='" & Val(in_cuenta) & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   get_moneda_cuenta = rst("id_moneda")
Else
   get_moneda_cuenta = "00001"
End If

End Function

Public Function get_nota_credito_admin() As Boolean

If KEY_NOTA_CREDITO_ADMIN = "si" Then
   If KEY_NOTA_CREDITO_USER <> "" Then
       If KEY_NOTA_CREDITO_USER = KEY_USUARIO Then
          get_nota_credito_admin = True
       Else
          get_nota_credito_admin = False
       End If
   Else
       get_nota_credito_admin = True
   End If
Else
    get_nota_credito_admin = True
End If

End Function

Public Function get_numero_letra(ByVal in_doc As String) As String
strCadena = "SELECT * FROM movimiento_venta WHERE ruc='" & KEY_RUC & "' and id_doc='0412' ORDER BY numero DESC LIMIT 1"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
    get_numero_letra = Format(Val(rstLocal("numero")) + 1, "000000")
Else
    get_numero_letra = Format(1, "000000")
End If
End Function
Public Function get_numero_credito() As String
strCadena = "SELECT count(*) FROM movimiento_venta WHERE id_forma_pago='02' and  id_doc IN('0001','0003','0007','0054') and ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If IsNull(rstL(0)) = False Then
    get_numero_credito = str(rstL(0))
Else
    get_numero_credito = "1"
End If
End Function

Public Function get_precio_costo(ByVal in_producto As String) As Double

strCadena = "SELECT precio_compra FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   get_precio_costo = rstLocal("precio_compra")
Else
   get_precio_costo = 0
End If

End Function

Public Function get_costo_producto(ByVal in_producto As String) As Single
strCadena = "SELECT precio_compra FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   get_costo_producto = rstLocal("precio_compra")
Else
   get_costo_producto = 0
End If

End Function
Public Function get_costo_sucursal(ByVal in_producto As String, ByVal in_alm As String, ByVal in_movimiento As String) As Double
strCadena = "SELECT ifnull(costo_promedio,0) as costo_promedio FROM kardex WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC , id_kardex DESC LIMIT 1"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   get_costo_sucursal = rstLocal("costo_promedio")
Else
   get_costo_sucursal = 0
End If
End Function

Public Function get_total_comprobante(ByVal in_venta As String) As Double
strCadena = "SELECT total FROM movimiento_venta WHERE id_venta='" & Val(in_venta) & "'"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   get_total_comprobante = rstP("total")
Else
   get_total_comprobante = 0
End If

End Function
Public Sub put_actualizar_kardex(ByVal in_producto As String, ByVal in_compra As String, ByVal in_costo As Double, ByVal in_alm As String)

strCadena = "UPDATE kardex SET costo_unitario='" & in_costo & "' WHERE id_movimiento='" & Val(in_compra) & "' and id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' and cantidad_real>0 LIMIT 1"
CnBd.Execute (strCadena)

strCadena = "UPDATE almacen_producto SET precio_compra='" & in_costo & "' WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' and id_alm='" & in_alm & "'"
CnBd.Execute (strCadena)



End Sub

Public Sub get_producto_agranel(ByVal in_producto As String, ByVal in_combo As DataCombo)

strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM view_unidad_producto WHERE id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataCombo(in_combo)
End Sub


Public Function get_precio_unidad(ByVal in_producto As String, ByVal in_unidad As String) As Single

strCadena = "SELECT precio FROM view_unidad_producto WHERE id_producto='" & in_producto & "' and id_unidad='" & in_unidad & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
   get_precio_unidad = rstT("precio")
Else
   get_precio_unidad = 0
End If





End Function

Public Function get_precio_producto(ByVal in_producto As String, ByVal in_alm As String)
strCadena = "SELECT precio_venta FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
   get_precio_producto = rstZ("precio_venta")
End If

End Function
Public Sub put_delete_kardex_recepcion(ByVal in_recepcion As String, ByVal in_serie_guia As String, ByVal in_numero_guia As String)
Dim in_doc As String
Dim in_serie As String
Dim in_numero As String

strCadena = "SELECT * FROM orden_compra_detalle WHERE id_orden='" & Val(in_recepcion) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
   rstZ.MoveFirst
   For i = 0 To rstZ.RecordCount - 1
       strCadena = "DELETE FROM kardex WHERE id_producto='" & rstZ("id_producto") & "' and id_movimiento='" & Val(in_recepcion) & "' and id_doc='0009' and id_serie='" & in_serie_guia & "' and id_numero='" & in_numero_guia & "' and ruc='" & KEY_RUC & "' and cantidad_real='" & rstZ("cantidad") & "' LIMIT 1"
       CnBd.Execute (strCadena)
       rstZ.MoveNext
   Next i
End If
   
   




End Sub

Public Sub eliminar_compras_general(ByVal id_doc As String, ByVal serie As String, ByVal numero As String, ByVal id_proveedor As String)
Dim idCompra As Double
strCadena = "SELECT * FROM movimiento_compra WHERE (numero='" & numero & "' AND serie='" & serie & "' AND id_doc='" & id_doc & "' AND id_proveedor='" & id_proveedor & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    idCompra = rst("id_compra")
    strCadena = "SELECT * FROM `imp_producto_detalle` WHERE id_compra='" & idCompra & "' and  id_alm='" & KEY_ALM & "' and  (vendido='si' or transferencia='si') and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
        MsgBox "SERIES VENDIDAS O TRANSFERIDAS" + Chr(13) + "IMPOSIBLE ELIMINAR", vbInformation, KEY_VENDEDOR
        Exit Sub
    End If
    strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & idCompra & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraTemporal(strCadena)
    If rstTemporal.RecordCount > 0 Then
        rstTemporal.MoveFirst
        For i = 0 To rstTemporal.RecordCount - 1
            strCadena = "DELETE FROM movimiento_compra_detalle WHERE id_compra='" & idCompra & "' AND id_detalle_compra='" & rstTemporal("id_detalle_compra") & "' "
            CnBd.Execute (strCadena)
            
            
            
            rstTemporal.MoveNext
        Next i
    End If
     strCadena = "DELETE FROM movimiento_compra WHERE id_compra='" & idCompra & "' AND ruc='" & KEY_RUC & "'"
     CnBd.Execute (strCadena)
     
     strCadena = "Call CON_Asiento_EliminarCompra('" & idCompra & "', '" & KEY_USUARIO & "', '" & KEY_RUC & "')"
     CnBd.Execute (strCadena)
    
End If

End Sub


Public Function verificar_cierre_caja(ByVal in_fecha As Date)
Dim in_mes As Integer
Dim in_anio As Integer

strCadena = "SELECT funct_caja_cerrada('" & Month(in_fecha) & "','" & Year(in_fecha) & "','" & KEY_RUC & "')"
Call ConfiguraRstP(strCadena)
verificar_cierre_caja = rstP(0)

End Function


Public Sub AnularVentas(ByVal TipoDoc As String, ByVal serie As String, ByVal numero As String, ByVal Almacen As String)
Dim nsaldo As Single
Dim in_venta As Double
Dim in_fecha As String
Dim in_observacion As String

 strCadena = "SELECT id_venta,fecha_emision FROM movimiento_venta WHERE id_doc='" & TipoDoc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "' ORDER BY id_venta DESC LIMIT 1"
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    in_venta = rst("id_venta")
    in_fecha = rst("fecha_emision")
 End If
 
 If KEY_CONTABILIDAD = "si" And (TipoDoc = "0001" Or TipoDoc = "0003" Or TipoDoc = "0007" Or TipoDoc = "0008") Then
 strCadena = "SELECT funct_validar_periodo_cierre('" & in_venta & "','" & KEY_RUC & "')"
 Call ConfiguraRstK(strCadena)
    If rstK(0) = 0 Then ' PERIODO ABIERTO
       ' strCadena = "call CON_EliminarVenta('" & in_venta & "','" & KEY_USUARIO & "') "
       'CnBd.Execute (strCadena)
        strCadena = "call CON_AnularVenta('" & in_venta & "','" & KEY_USUARIO & "')"
       CnBd.Execute (strCadena)
    Else      ' PERIODO CERRADO
       
        If get_firma_online(TipoDoc, serie) = "si" Then
            Call duplicar_comprobante(in_venta)
            Exit Sub
        Else
           MsgBox "IMPOSIBLE ANULAR PERIODO CERRADO", vbInformation, "PERIODO CERRADO"
           Exit Sub
        End If
       
    End If
 Else
        If get_firma_online(TipoDoc, serie) = "si" Then
            If validar_periodo_nocontable(in_fecha) = False Then
                MsgBox "Advertencia." + Chr(13) + "Comprobante puesto a CERO" + Chr(13) + "Tiene que Registrar un NUEVO DOCUMENTO" + Chr(13) + "En el periodo ACTUAL.", vbInformation, key_
            
            End If
           
        Else
           If Month(in_fecha) = Month(KEY_FECHA) And Year(in_fecha) = Year(KEY_FECHA) Then
              GoTo anularya
              Exit Sub
           End If
           
           If TipoDoc <> "0001" And TipoDoc <> "0003" And TipoDoc <> "0007" And TipoDoc <> "0008" Then
                GoTo anularya
           End If
           
           MsgBox "IMPOSIBLE ANULAR PERIODO CERRADO", vbInformation, "PERIODO CERRADO"
           Exit Sub
        End If
 End If
 
anularya:
 
 in_observacion = "ANULADO:" & KEY_FECHA & Space(2) & str(Time) & Space(2) & Space(2) & get_persona(KEY_USUARIO)
 
 strCadena = "UPDATE movimiento_venta SET observacion='" & in_observacion & "', anulado='si',igv=0,valor_venta=0,exonerado=0,total=0,saldo=0 WHERE id_venta='" & in_venta & "'"
 CnBd.Execute (strCadena)
 
 
 nsaldo = 0
 strCadena = "SELECT * FROM movimiento_venta WHERE id_alm='" & Almacen & "' AND id_doc='" & TipoDoc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "' AND anulado='si' LIMIT 1"
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    
    
    strCadena = "DELETE FROM imp_tramite WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    strCadena = "UPDATE movimiento_venta SET anulado='si' WHERE id_venta='" & rst("id_recibo") & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    
    nsaldo = rst("total")
    
    strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        rstZ.MoveFirst
        For m = 0 To rstZ.RecordCount - 1
            
            strCadena = "DELETE FROM kardex WHERE id_producto='" & rstZ("id_producto") & "' and id_doc='" & rst("id_doc") & "' and id_serie='" & rst("serie") & "' and id_movimiento='" & rstZ("id_venta") & "' and ruc='" & KEY_RUC & "' "
            CnBd.Execute (strCadena)
            
                    strCadena = "SELECT sum(cantidad_real),sum(cantidad_pendiente) FROM kardex WHERE id_tipo_movimiento<>'10' and  id_alm='" & rst("id_alm") & "' and  id_producto='" & rstZ("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstK(strCadena)
                    If IsNull(rstK(0)) = True Then
                            in_real = 0
                        Else
                            in_real = rstK(0)
                        End If
                        
                        If IsNull(rstK(1)) = True Then
                            in_pendiente = 0
                        Else
                            in_pendiente = rstK(1)
                        End If
                        
                    
                    strCadena = "SELECT sum(cantidad_factura) FROM kardex WHERE id_alm='" & rst("id_alm") & "' and  id_producto='" & rstZ("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstK(strCadena)
                    If IsNull(rstK(0)) = True Then
                            in_contable = 0
                    Else
                            in_contable = rstK(0)
                    End If
                        
                    
                    strCadena = "UPDATE almacen_producto set stock = '" & in_real & "' ,stock_factura='" & in_contable & "' , `stock_contable` = '" & in_pendiente & "'  WHERE ruc = '" & KEY_RUC & "' and id_alm ='" & rst("id_alm") & "' and id_producto = '" & rstZ("id_producto") & "'"
                    CnBd.Execute (strCadena)
                    
                    
            strCadena = "UPDATE imp_producto_detalle SET vendido='no' WHERE id_detalle='" & rstZ("id_detalle_serie") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
            
            
            rstZ.MoveNext
        Next m
    End If
    
    
    
    
    Call put_actualizar_anulacion(in_venta)
    
    
    
    
    
    
    
    
    
    
  
 End If
 
 

End Sub
Public Sub put_actualizar_anulacion(ByVal in_recibo As String)
strCadena = "SELECT * FROM mis_cuentas_det_detalle WHERE id_detalle='" & Val(in_recibo) & "' "
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   rstL.MoveFirst
   For i = 0 To rstL.RecordCount - 1
            
            If rstL("id_tipo") = "01" Then
                
                strCadena = "UPDATE mis_cuentas_det SET ruc='EXTORNADO'  WHERE ruc='" & KEY_RUC & "' and id_venta='" & rstL("id_movimiento") & "' and monto='" & rstL("monto_pagado") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                
            Else
                strCadena = "UPDATE mis_cuentas_det SET ruc='EXTORNADO'  WHERE ruc='" & KEY_RUC & "' and id_venta='" & rstL("id_detalle") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                
            End If
            CnBd.Execute (strCadena)
            rstL.MoveNext
        
   Next i
End If
End Sub



Public Sub duplicar_comprobante(ByVal in_venta As String)
Dim in_venta_nuevo As String
strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & in_venta & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
strCadena = "call P_insert_venta_conta_ii('" & rstK("id_doc") & "','" & rstK("id_alm") & "','" & rstK("id_forma_pago") & "','" & rstK("id_moneda") & "','" & rstK("id_delivery") & "'," & _
            "'" & rstK("serie") & "','" & rstK("numero") & "','" & rstK("id_cliente") & "','" & rstK("ncliente") & "','0','0','0','0','0'," & _
            "'" & rstK("monto_pago") & "','" & rstK("monto_vuelto") & "','" & KEY_FECHA & "','" & KEY_FECHA & "','" & rstK("id_tipo") & "','" & rstK("id_vendedor") & "','" & KEY_USUARIO & "','" & KEY_CAMBIO & "','no','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "'" & _
            ",'" & rstK("documento") & "',CURTIME(),'T','" & rstK("direccion") & "','no','" & rstK("sunat_hash") & "','" & rstK("sunat_key") & "','" & rstK("id_tipo_nota") & "','" & rstK("motivo_nota") & "','" & rstK("id_guia") & "','" & rstK("nguia") & "','" & KEY_VENTANILLA & "','" & rstK("id_tipo") & "','" & rstK("id_seguro") & "','" & KEY_RUC & "')"
            Call ConfiguraRstPP(strCadena)
            in_venta_nuevo = rstPP(0)
            
            If KEY_CONTABILIDAD = "si" Then
                strCadena = "call P_insert_venta_agenda('" & in_venta_nuevo & "')"
                CnBd.Execute (strCadena)
            End If

            
            
            
            strCadena = "SELECT * FROM  movimiento_venta_monto WHERE id_venta='" & in_venta & "'"
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount > 0 Then
               rstK.MoveFirst
               For i = 0 To rstK.RecordCount - 1
                   strCadena = "INSERT INTO movimiento_venta_monto(`id_venta`,`id_forma_pago`,`monto`,`monto_caja`,`id_tarjeta`,`id_tarjeta_numero`,`id_tarjeta_operacion`,`id_recibo`,`banco`,`cheque`,`cuenta_contable`,`ruc`)VALUES " & _
                   "('" & in_venta_nuevo & "','" & rstK("id_forma_pago") & "','" & rstK("monto") & "','" & rstK("monto_caja") & "','" & rstK("id_tarjeta") & "','" & rstK("id_tarjeta_numero") & "','" & rstK("id_tarjeta_operacion") & "','" & rstK("id_recibo") & "','" & rstK("banco") & "','" & rstK("cheque") & "','" & rstK("cuenta_contable") & "','" & KEY_RUC & "')"
                   CnBd.Execute (strCadena)
                   rstK.MoveNext
               Next i
            End If
           
            
End If
           
End Sub

Public Function get_forma_pago(ByVal in_cuenta As String) As String
strCadena = "SELECT * FROM mis_cuentas WHERE id_cuenta='" & in_cuenta & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    get_forma_pago = rstL("id_tipo")
Else
    get_forma_pago = "01"
End If
End Function


Public Function get_validar_cuenta_asociada(ByVal in_cuenta_contable As String) As Boolean


If Mid(in_cuenta_contable, 1, 1) <> "6" Then
    get_validar_cuenta_asociada = True
Else
strCadena = "SELECT * FROM con_cuentaasociada WHERE CuentaContable='" & in_cuenta_contable & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    get_validar_cuenta_asociada = True
Else
    get_validar_cuenta_asociada = False
End If
End If

End Function

Public Function get_documento_abrev(ByVal in_doc As String)
strCadena = "SELECT * FROM comprobantes WHERE id_doc='" & in_doc & "'"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
   get_documento_abrev = rstZ("doc_abrev")
Else
    get_documento_abrev = "OTROS"
End If

End Function

Public Sub put_cambio_precio(ByVal in_producto As String, ByVal in_precio As Double, in_observacion)

strCadena = "call put_variacion_precio('" & in_producto & "','" & in_precio & "','" & in_observacion & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
End Sub




Public Function Eliminar_gasto_viaticos(ByVal in_detalle As String, ByVal in_compra As String) As Boolean
     Eliminar_gasto_viaticos = False
     
     strCadena = "DELETE FROM solicitud_dinero_declarar WHERE id_detalle='" & Val(in_detalle) & "' and ruc='" & KEY_RUC & "'"
     CnBd.Execute (strCadena)
     
     strCadena = "DELETE FROM movimiento_compra WHERE id_compra='" & Val(in_compra) & "' AND ruc='" & KEY_RUC & "'"
     CnBd.Execute (strCadena)
      
     strCadena = "Call CON_Asiento_EliminarCompra('" & Val(in_compra) & "', '" & KEY_USUARIO & "', '" & KEY_RUC & "')"
     CnBd.Execute (strCadena)
    Eliminar_gasto_viaticos = True

End Function


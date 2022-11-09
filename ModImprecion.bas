Attribute VB_Name = "ModImprecion"
Public Sub Orden_Impresion(ByVal id_doc As String, ByVal serie As String, ByVal numero As String, ByVal id_tipo_factura As String, ByVal in_venta As String, Optional Direccion As String)
 '  Call AbreGaveta
    Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    '************ IMPRESORA X DEFECTO *********
                             If KEY_IMPRESORA = "si" Then
                                Call funSeleccionImpresora(KEY_PRINTER)            'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                           Else
                               Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                          End If
    '*******************************************
    'Printer.Font.name = "FontB11"Draft 17cpi
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If rst("impresion") = "no" Then
           MsgBox "NO ESTA ACTIVADO LA IMPRESION", vbInformation
        Else
           
          
                Call impresion_formato(rst("id_formato_impresion"), id_doc, serie, numero, id_tipo_factura, in_venta, Direccion)
              
          
           
        End If
    Else
        MsgBox "DOCUMENTO Y SERIE NO REGISTRADA EN LA SUCURSAL", vbInformation
    End If
    
    
End Sub
Public Sub generar_hc(ByVal in_dni As String)

strCadena = "SELECT dni,ruc_vinculado,a_paterno,a_materno,nombres,nombre_completo,nacimiento,direccion,'" & KEY_VENDEDOR & "', " & _
"ubigeo,edad,sexo,vinculado,titular,parentesco,estado,id_empresa,empresa,n_carne,seguro,expedicion,expiracion,celular,estado_civil FROM view_hc_ruc WHERE dni='" & Trim(in_dni) & "'"
Call ConfiguraRst(strCadena)

Ans = ShowMultiReport(rst, "RptHcP", , App.Path + "\Reportes\")
End Sub
Public Sub impresion_consolidado_arqueo_det(ByVal in_arqueo As String, ByVal in_total_sistema As Double)
Dim nombre_paciente As String
Dim id_producto As String
Dim ptotal As Double, direccion_destino As String, impresiones As Integer
Dim nn As String, id_agenda As Double, horaf As String, fechad As String
   'Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
  ' Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    

                            If KEY_IMPRESORA = "si" Then
                                Call funSeleccionImpresora(KEY_PRINTER)            'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            Else
                                Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            End If
    
    
    'Printer.DeviceName = Trim("EPSON TM-U220 Receipt")
    'If Prt.DeviceName = "EPSON TM-U220 Receipt" Then
     '       Set Printer = Prt
      '  End If
        
   ' If id_doc <> "0109" Then
    '    Printer.Font.name = "control"
     '   Printer.Print "A"
    'End If
    'Tahoma
    Printer.Font.name = "Tahoma"
    'Printer.Font.name = "Calibri"
    Printer.Font.Size = "8"
    'Printer.Font.name = "Draft 17cpi"
    'Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
    Printer.Print ""
    Printer.Font.Bold = True
    'Printer.Font.Size = "30"
    
    Printer.Print Tab(1); "::::::::::::::: ARQUEO DE CAJA :::::::::::::::::::"
    'Printer.Font.Size = "10"
    Printer.Font.Bold = False
    Printer.Print Tab(0); "-----------------------------------------------------------------"
    Printer.Print Tab(0); KEY_EMPRESA
    Printer.Print Tab(0); "RUC   :" & KEY_RUC
    Printer.Print Tab(0); "FECHA :" & KEY_FECHA
    Printer.Print Tab(0); "-----------------------------------------------------------------"
    Printer.Print Tab(0); "USUARIO :" + Space(2) + KEY_VENDEDOR
    Printer.Print Tab(0); ""
    Printer.Print Tab(1); "FIRMA   :" + Space(2) + "_ _ _ _ _ _ _ _ _ _"
    Printer.Print Tab(0); ""
    Printer.Print Tab(1); ":::::::::::::::  TURNO TRABAJO  ::::::::::::::::::"
    Printer.Print Tab(0); ""
    strCadena = "SELECT * FROM turno WHERE id_turno='" & turno & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); ":::::::::::: " & Space(1) & rst("descripcion") & Space(1) & " ::::::::::::"
    Else
        Printer.Print Tab(0); ":::::::::::: " & Space(1) & "NO SELECCIONADO" & Space(1) & " ::::::::::::"
    End If
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); "-----------------------------------------------------------------"
        strCadena = "SELECT * FROM view_arqueo WHERE id_arqueo='" & in_arqueo & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           in_total = 0
           Printer.Print Tab(0); Mid("   DIVISA" + Space(50), 1, 40) & Mid("CANT" + Space(10), 1, 10) & Mid("TOTAL" + Space(10), 1, 10)
           Printer.Print Tab(0); "-----------------------------------------------------------------"
           For i = 0 To rst.RecordCount - 1
               in_unidad = rst("cantidad") * rst("valor")
               Printer.Print Tab(0); Mid(rst("billete") + Space(50), 1, 35) & Mid(str(rst("cantidad")) + Space(10), 1, 10) & "=" & Mid(Format(in_unidad, "#,##0.00") + Space(10), 1, 10)
               in_total = in_total + in_unidad
               rst.MoveNext
           Next i
           Printer.Print Tab(0); "-----------------------------------------------------------------"
           Printer.Print Tab(0); "TOTAL FÍSICO :  " & Format(in_total, "#,##0.00")
        End If
        
        
        
        If in_total_sistema > in_total Then
              Printer.Print Tab(0); "-----------------------------------------------------------------"
              Printer.Print Tab(0); "FALTANTE :  " & Format(in_total_sistema - in_total, "#,##0.00")
           End If
           
           If in_total_sistema < in_total Then
              Printer.Print Tab(0); "-----------------------------------------------------------------"
              Printer.Print Tab(0); "SOBRANTE :  " & Format(in_total - in_total_sistema, "#,##0.00")
           End If
           
           
        
           
           Printer.Print Tab(0); "-----------------------------------------------------------------"
           Printer.Print Tab(0); "TOTAL SISTEMA :  " & Format(in_total_sistema, "#,##0.00")
           
           '***********
           
           
           
           
           
            Printer.Print Tab(0); " "
            Printer.Print Tab(0); " "
           Printer.Print Tab(0); "HORA IMPRESION:" + Space(1) + Format(Time, "hh:mm:ss am/pm")
           Printer.EndDoc
           
           
        Exit Sub
 
End Sub


Public Sub impresion_cierre_grifo(ByVal in_isla As String, ByVal in_turno As String, ByVal in_fecha As String, ByVal in_fecha_fin As String)
Dim nombre_paciente As String
Dim id_producto As String
Dim in_consumo As Double
Dim in_total As Double
Dim in_acumulado_galones As Double
Dim in_acumulado_soles As Double

   'Call CargaDefConfigEpsonTM
   Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    'Printer.Font.name = "FontB11"
    'Printer.Font.Size = "10"
    
   
  ' Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    

                            If KEY_IMPRESORA = "si" Then
                                Call funSeleccionImpresora(KEY_PRINTER)            'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            Else
                                Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            End If
    
    
    'Printer.DeviceName = Trim("EPSON TM-U220 Receipt")
    'If Prt.DeviceName = "EPSON TM-U220 Receipt" Then
     '       Set Printer = Prt
      '  End If
        
   ' If id_doc <> "0109" Then
     '   Printer.Font.name = "control"
     '   Printer.Print "A"
    'End If
    'Tahoma
    Printer.Font.name = "Tahoma"
    'Printer.Font.name = "Calibri"
    Printer.Font.Size = "7"
    'Printer.Font.name = "Draft 17cpi"
    'Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
    Printer.Print ""
    Printer.Font.Bold = True
    'Printer.Font.Size = "30"
    in_acumulado_galones = 0
    in_acumulado_soles = 0
    Printer.Print Tab(0); ":::::::::::::::      CIERRE DE TURNO :::::::::::::::::::"
    'Printer.Font.Size = "10"
    Printer.Font.Bold = False
    Printer.Print Tab(0); "--------------------------------------------------------------"
    Printer.Print Tab(0); KEY_EMPRESA
    Printer.Print Tab(0); "RUC         :" & KEY_RUC
    Printer.Print Tab(0); "DIRECCION   :" & KEY_DIRECCION
    Printer.Print Tab(0); "SUCURSAL    :" & KEY_DIRECCION_ALM
    Printer.Print Tab(0); "FECHA       :" & KEY_FECHA
    Printer.Print Tab(0); "---------------------------------------------------------------"
    Printer.Print Tab(0); "USUARIO :" + Space(2) + KEY_VENDEDOR
    Printer.Print Tab(0); ""
    Printer.Print Tab(1); "TURNO TRABAJO :" & get_turno(in_turno)
  
           
           '***********
    'Para recorrer todos los surtidores
    strCadena = "SELECT DISTINCT id_surtidor,almacen FROM view_lectura_surtidor WHERE id_isla='" & in_isla & "' and fecha>='" & Format(CVDate(in_fecha_ini), "YYYY-mm-dd") & "' and fecha<='" & Format(CVDate(in_fecha_fin), "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY id_surtidor ASC"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount > 0 Then
       rstA.MoveFirst
       Printer.Print Tab(0); "================================================================="
       Printer.Font.Bold = True
       Printer.Print Tab(1); "ISLA  :" + Space(2) + rstA("almacen")
       Printer.Print Tab(0); "================================================================="
       
   For j = 0 To rstA.RecordCount - 1
    
    strCadena = "SELECT * FROM view_lectura_surtidor WHERE id_surtidor='" & rstA("id_surtidor") & "' and  id_turno='" & in_turno & "' and  id_isla='" & in_isla & "' and fecha>='" & Format(CVDate(in_fecha_ini), "YYYY-mm-dd") & "' and fecha<='" & Format(CVDate(in_fecha_fin), "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY id_lectura ASC "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
           rst.MoveFirst
           
           in_consumo = 0
           in_total = 0
           Printer.Font.Bold = True
          
           Printer.Print Tab(0); rst("lado"); Space(5) & rst("descripcion")
           Printer.Print Tab(0); "================================================================="
           Printer.Font.Bold = False
           Printer.Print Tab(0); "PRODUCTO :" & get_producto(rst("id_producto"))
          ' Printer.Print Tab(0); "MARCACION [GALONES]  :"
           Printer.Print Tab(0); "HORA          |       M.INICIAL   |       M.FINAL     |        CANT"
           Printer.Print Tab(0); "================================================================="
       For i = 0 To rst.RecordCount - 1
           in_consumo = rst("lectura_fin") - rst("lectura_ini")
           in_total = in_total + in_consumo
           
           Printer.Print Tab(1); Mid(rst("hora_cadena") + Space(10), 1, 20) & Mid(Format(rst("lectura_ini"), "#,##0.000") + Space(10), 1, 20) & Mid(Format(rst("lectura_fin"), "#,##0.000") + Space(10), 1, 20) & Mid(Format(str(in_consumo), "#,##0.000") + Space(10), 1, 20)
           rst.MoveNext
       Next i
          in_acumulado_galones = in_acumulado_galones + in_total
           Printer.Print Tab(0); "================================================================="
           Printer.Print Tab(0); Mid("TOTAL GALONES :" + Space(70), 1, 60) & Mid(Format(in_total, "#,##0.000") + Space(10), 1, 20)
           Printer.Print Tab(0); "================================================================="
          
           rst.MoveFirst
           in_consumo = 0
           in_total = 0
          
           
           'Printer.Print Tab(0); "MARCACION [SOLES]  :"
           Printer.Print Tab(0); "HORA          |       M.INICIAL   |       M.FINAL     |        CANT"
           Printer.Print Tab(0); "================================================================="
         For i = 0 To rst.RecordCount - 1
           in_consumo = rst("lectura_fin_soles") - rst("lectura_ini_soles")
           in_total = in_total + in_consumo
           Printer.Print Tab(1); Mid(rst("hora_cadena") + Space(10), 1, 20) & Mid(Format(rst("lectura_ini_soles"), "#,##0.000") + Space(10), 1, 20) & Mid(Format(rst("lectura_fin_soles"), "#,##0.000") + Space(10), 1, 20) & Mid(Format(Format(in_consumo), "#,##0.000") + Space(10), 1, 20)
           rst.MoveNext
       Next i
           in_acumulado_soles = in_acumulado_soles + in_total
           Printer.Print Tab(0); "================================================================="
           Printer.Print Tab(0); Mid("TOTAL SOLES :" + Space(70), 1, 70) & Mid(Format(in_total, "#,##0.000") + Space(10), 1, 20)
           Printer.Print Tab(0); "================================================================="
          
          
          
           
           
          
           
          
           
       rstA.MoveNext
           
          
      

        
    End If
           
   Next j
           
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); ""
        
        Printer.Print Tab(0); "================================================================="
        Printer.Print Tab(0); Mid("ACUMULADO SOLES   :" + Space(70), 1, 50) & Mid(Format(in_acumulado_soles, "#,##0.000") + Space(10), 1, 20)
        Printer.Print Tab(0); Mid("ACUMULADO GALONES :" + Space(70), 1, 50) & Mid(Format(in_acumulado_galones, "#,##0.000") + Space(10), 1, 20)
        Printer.Print Tab(0); "================================================================="
           
           
        Printer.Print Tab(0); "      -------------------------------------------------------------    "
        Printer.Print Tab(10); "DNI:" & KEY_USUARIO & Space(2) & KEY_VENDEDOR
         End If
  Printer.EndDoc
  Exit Sub
 
End Sub



Public Sub impresion_consolidado_arqueo(ByVal in_fecha_ini As Date, ByVal in_fecha_fin As Date, ByVal in_ventanilla As String, ByVal in_operador As String, ByVal in_almacen As String)
Dim nombre_paciente As String
Dim id_producto As String
Dim ptotal As Double, direccion_destino As String, impresiones As Integer
Dim nn As String, id_agenda As Double, horaf As String, fechad As String
   'Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
  ' Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    

                            If KEY_IMPRESORA = "si" Then
                                Call funSeleccionImpresora(KEY_PRINTER)            'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            Else
                                Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            End If
    
    
    'Printer.DeviceName = Trim("EPSON TM-U220 Receipt")
    'If Prt.DeviceName = "EPSON TM-U220 Receipt" Then
     '       Set Printer = Prt
      '  End If
        
   ' If id_doc <> "0109" Then
    '    Printer.Font.name = "control"
     '   Printer.Print "A"
    'End If
    'Tahoma
    Printer.Font.name = "Tahoma"
    'Printer.Font.name = "Calibri"
    Printer.Font.Size = "8"
    'Printer.Font.name = "Draft 17cpi"
    'Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
    Printer.Print ""
    Printer.Font.Bold = True
    'Printer.Font.Size = "30"
    
    Printer.Print Tab(1); "::::::::::::::: ARQUEO DE CAJA :::::::::::::::::::"
    'Printer.Font.Size = "10"
    Printer.Font.Bold = False
    Printer.Print Tab(0); "-----------------------------------------------------------------"
    Printer.Print Tab(0); KEY_EMPRESA
    Printer.Print Tab(0); "RUC   :" & KEY_RUC
    Printer.Print Tab(0); "FECHA :" & KEY_FECHA
    Printer.Print Tab(0); "-----------------------------------------------------------------"
    Printer.Print Tab(0); "USUARIO :" + Space(2) + KEY_VENDEDOR
    Printer.Print Tab(0); ""
    
    Printer.Print Tab(0); ""
    
 strCadena = "SELECT nformapago,sum(monto_caja) " & _
"  FROM view_reporte_detallado_ultimate WHERE anulado='no' and  afecta_factura='no' and  afecta_caja='si' and id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and  id_vendedor LIKE '%" & Trim(in_operador) & "%' AND  id_alm LIKE  '%" & in_almacen & "%' and  fecha_emision>='" & Format(in_fecha_ini, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(in_fecha_fin, "YYYY-mm-dd") & "'  AND ruc='" & KEY_RUC & "' group by id_forma_pago"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   in_total = 0
   For i = 0 To rst.RecordCount - 1
        Printer.Print Tab(5); ":::" & Mid(rst("nformapago") & Space(100), 1, 30) & Space(4) & Format(rst(1), "#,##0.00")
        Printer.CurrentY = Printer.CurrentY + 1.5
        in_total = in_total + rst(1)
        rst.MoveNext
   Next i
   Printer.Print Tab(0); "-----------------------------------------------------------------"
        Printer.Print Tab(1); "::::::  TOTAL RECAUDADO   :::::::" & Format(in_total, "#,##0.00")
End If
     
    
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); ""
        
           
           
           Printer.Print Tab(0); "HORA IMPRESION:" + Space(1) + Format(Time, "hh:mm:ss am/pm")
           Printer.EndDoc
           
           
        Exit Sub
 
End Sub



Public Sub impresion_detallado_ticket(ByVal in_fecha_ini As Date, ByVal in_fecha_fin As Date, ByVal in_ventanilla As String, _
ByVal in_operador As String, ByVal in_almacen As String, ByVal in_turno As String)

    Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
  ' Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    

                            If KEY_IMPRESORA = "si" Then
                                Call funSeleccionImpresora(KEY_PRINTER)            'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            Else
                                Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            End If
    
    
    'Printer.DeviceName = Trim("EPSON TM-U220 Receipt")
    'If Prt.DeviceName = "EPSON TM-U220 Receipt" Then
     '       Set Printer = Prt
      '  End If
        
   ' If id_doc <> "0109" Then
    '    Printer.Font.name = "control"
     '   Printer.Print "A"
    'End If
    'Tahoma
    Printer.Font.name = "FontB11"
    Printer.Font.Size = "10"
    
    Printer.Print ""
    Printer.Font.Bold = True
    'Printer.Font.Size = "30"
    
    Printer.Print Tab(1); "::::::::::::::: ARQUEO DE CAJA :::::::::::::::::::"
    'Printer.Font.Size = "10"
    Printer.Font.Bold = False
    Printer.Print Tab(0); "-----------------------------------------------------------------"
    Printer.Print Tab(0); KEY_EMPRESA
    Printer.Print Tab(0); "RUC   :" & KEY_RUC
    Printer.Print Tab(0); "FECHA :" & KEY_FECHA
    Printer.Print Tab(0); "-----------------------------------------------------------------"
    Printer.Print Tab(0); "USUARIO :" + Space(2) + KEY_VENDEDOR
    Printer.Print Tab(0); ""
    Printer.Print Tab(0); ""
    
 
    Printer.EndDoc
           
           
        Exit Sub
 
End Sub


Public Sub impresion_formato(ByVal tipo As String, ByVal id_doc As String, ByVal serie As String, ByVal numero As String, ByVal id_tipo_factura As String, ByVal in_venta As String, Optional Direccion As String)
 Dim param As Variant
Select Case tipo
    Case 1 ' 1/4 A4.  PAPEL CONTINUO
            Select Case id_doc
                Case KEY_FACTURA
                     If KEY_FACTURACION_ELECTRONICA = "no" Then
                        Call impresion_formato_1_factura(id_doc, serie, numero)
                    Else
                        Call impresion_formato_1_factura_electronica(id_doc, serie, numero)
                    End If
                
                Case "0007"
                      Call impresion_formato_1_nota_electronica(id_doc, serie, numero)
                Case KEY_GUIA
                    
                    Call impresion_formato_1_guia(id_doc, serie, numero)
                Case KEY_BOLETA
                    If KEY_FACTURACION_ELECTRONICA = "si" Then
                        Call impresion_formato_1_boleta_electronica(id_doc, serie, numero)
                    Else
                        Call impresion_formato_1_boleta(id_doc, serie, numero)
                    End If
                Case "0054"
                    Call impresion_formato_1_cotizacion(id_doc, serie, numero)
                    
                Case "0099"
                    Call impresion_formato_1_cotizacion(id_doc, serie, numero)
                Case KEY_PEDIDO
                    
                    Call impresion_formato_1_pedido(id_doc, serie, numero)
                Case KEY_RBOINGRESO
                    Call impresion_formato_1_rboingreso(id_doc, serie, numero)
                Case Else
                    
                        If KEY_FACTURACION_ELECTRONICA = "si" Then
                            Call impresion_tiketera_electronica(id_doc, serie, numero)
                        Else
                            Call impresion_tiketera(id_doc, serie, numero)
                        End If
            
                    
                    
                    
                   
                    
            End Select
        
        
        
    Case 2 ' 1/2 A4.  PAPEL CONTINUO
        If id_doc = KEY_BOLETA Then
                   Call impresion_formato_2_boleta(id_doc, serie, numero)
        End If
        If id_doc = KEY_GUIA Then
            Call impresion_formato_2_guia(id_doc, serie, numero)
        End If
        If id_doc = KEY_FACTURA Then
             Call impresion_formato_2_factura(id_doc, serie, numero)
        End If
    
    Case 3
            If id_doc = "0009" Then
                
                Dim arrt(0 To 2, 1 To 2) As String
                
                arrt(0, 1) = "par_hora"
                
                
                strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0009' and serie='" & serie & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRst(strCadena)
                If rst.RecordCount > 0 Then
                    If rst("electronico") = "si" Then
                        strCadena = "call ADM_guia_remision('3','" & Val(in_venta) & "','" & KEY_RUC & "')"
                        Call ConfiguraRst(strCadena)
                        
                        strCadena = "call ADM_guia_remision('4','" & Val(in_venta) & "','" & KEY_RUC & "')"
                        Call ConfiguraRstK(strCadena)
                        
                        strCadena = "call ADM_guia_remision('5','" & Val(in_venta) & "','" & KEY_RUC & "')"
                        Call ConfiguraRstL(strCadena)
                        If rstL(0) > 0 Then
                            Ans = ShowMultiReport(rst, "rpt_guia_electronica_serv", param, App.Path + "\Reportes\", , , True, , rstK, "rptguia_detalle_electronico")
                        Else
                            Ans = ShowMultiReport(rst, "rpt_guia_electronica", param, App.Path + "\Reportes\", , , True, , rstK, "rptguia_detalle_electronico")
                        End If
                        
                        Exit Sub
                        
                    Else
                    
                    strCadena = "SELECT id_transferencia,fecha,fecha_traslado,id_doc,comprobante,id_remitente,remitente,id_destinatario,destinatario,direccion_origen,direccion_destino,ubigeo,dni_atencion,atencion,almacen_origen,id_alm_destino,id_transporte,transporte,marca_placa,placa,mtc,certificado,id_chofer,chofer,licencia,id_motivo,peso_total,valor_mercaderia,observacion2,observacion3,'" & get_documento_venta(in_venta) & "',ruc FROM view_tranferencia_cabecera WHERE id_transferencia='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRst(strCadena)
                    If rst.RecordCount > 0 Then
                        strCadena = "SELECT * FROM movimiento_transferencia WHERE id_transferencia='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
                        Call ConfiguraRstA(strCadena)
                        If rstA.RecordCount > 0 Then
                            arrt(0, 2) = rstA("observacion")
                        End If
                   
                        param = arrt()
                        
                        strCadena = "SELECT * FROM view_transferencia_detallado WHERE id_transferencia='" & Val(in_venta) & "'"
                        Call ConfiguraRstK(strCadena)
                        Ans = ShowMultiReport(rst, "rpt_guia_remision", param, App.Path + "\Reportes\", , , True, , rstK, "rptguia_detalle")
                End If
                End If
                End If
                Exit Sub
            End If
            
            
            
            strCadena = "SELECT id_tipo_factura,id_moneda,total,fecha_emision,fecha_vencimiento,id_doc,dni_save,id_vendedor,percepcion,id_tipo,afecto_detraccion FROM movimiento_venta WHERE id_venta='" & in_venta & "' LIMIT 1"
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount > 0 Then
                in_tipo_material_servicio = rstK("id_tipo")
                in_detraccion = rstK("afecto_detraccion")
                in_moneda = rstK("id_moneda")
                in_total = rstK("total")
                If rstK("id_tipo_factura") = "00002" Then
                   If KEY_CODIGO_UNIVERSAL_IMPRESION = "si" Then
                        strCadena = "SELECT `id_venta`,`fecha_emision`,`fecha_vencimiento`,doc_des,`documento`,`id_cliente`,`ncliente`,`direccion`,`exonerado`,`valor_venta`,`igv`,`total`,`sunat_key`,`sunat_hash`,`id_producto`,`detalle`,`observacion`,`cantidad`,`valor_neto`,`precio`,`todatl_detalle`,'---','" & UCase(EnLetras_moneda(in_total, in_moneda)) & "',resolucion,nro_chasis,nro_chasis,serie,modelo,color,marca,anio_fabricacion,nro_dua,nro_item,in_guia,poliza,ip,id_moneda,`ruc` FROM view_factura_electronica_serial WHERE id_venta='" & Val(in_venta) & "'"
                        Call ConfiguraRst(strCadena)
                        
                        strCadena = "SELECT id_venta,tipo_venta,descripcion,monto,cuotas FROM view_venta_pago WHERE id_venta='" & in_venta & "'"
                        Call ConfiguraRstK(strCadena)
                        
                        strCadena = "SELECT potencia_motor,num_cilindros,cilindraje,tipo_gasolina FROM producto_importaciones WHERE id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                        Call ConfiguraRstA(strCadena)
                        
                        Ans = ShowMultiReport(rst, "factura_elec_serial_universal", , App.Path + "\Reportes\", , , , , rstK, "modalidad_pago", rstA, "rpt_datosimportacion")
                        
                   Else
                        
                        strCadena = "SELECT `id_venta`,`fecha_emision`,`fecha_vencimiento`,doc_des,`documento`,`id_cliente`,`ncliente`,`direccion`,`exonerado`,`valor_venta`,`igv`,`total`,`sunat_key`,`sunat_hash`,`id_producto`,detalle,`observacion`,`cantidad`,`valor_neto`,`precio`,`todatl_detalle`,'---','" & UCase(EnLetras_moneda(in_total, in_moneda)) & "',resolucion,nro_chasis,nro_chasis,serie,modelo,color,marca,anio_fabricacion,nro_dua,nro_item,in_guia,id_moneda,`ruc` FROM view_factura_electronica_serial WHERE id_venta='" & Val(in_venta) & "'"
                        Call ConfiguraRst(strCadena)
                        strCadena = "SELECT id_venta,tipo_venta,descripcion,monto,cuotas FROM view_venta_pago WHERE id_venta='" & in_venta & "'"
                        Call ConfiguraRstK(strCadena)
                        Ans = ShowMultiReport(rst, "factura_elec_serial", , App.Path + "\Reportes\", , , , , rstK, "modalidad_pago")
                   End If
                   
                Else
                    
                   '*---------
                   Dim arry(0 To 2, 1 To 2) As String
                   Dim paramy As Variant
                   arry(0, 1) = "vendedor_proforma"
                   arry(1, 1) = "detraccion"
                   arry(2, 1) = "percepcion"
                   
                   arry(0, 2) = get_persona(rstK("id_vendedor"))
                   arry(1, 2) = in_detraccion
                   arry(2, 2) = rstK("percepcion")
                   paramy = arry()



                   '*---------
                    
                    If rstK("id_doc") = "0099" Then
                        in_vencimiento = Format(DateAdd("d", 7, rstK("fecha_emision")), "dd-mm-YYYY")
                    Else
                        in_vencimiento = Format(rstK("fecha_vencimiento"), "dd-mm-YYYY")
                    End If
                    
                    
                    If KEY_CODIGO_UNIVERSAL_IMPRESION = "si" Then
                        strCadena = "SELECT `id_venta`,`fecha_emision`,'" & in_vencimiento & "', doc_des,`documento`,`id_cliente`,`ncliente`,`direccion`,`exonerado`,`valor_venta`,`igv`,`total`,`sunat_key`,`sunat_hash`,`id_producto`,`detalle`,`observacion`,`cantidad`,`valor_neto`,`precio`,`todatl_detalle`,unidad,'" & UCase(EnLetras_moneda(in_total, in_moneda)) & "',resolucion,in_guia,id_moneda,orden_compra,tc,icbper,`ruc` FROM view_factura_electronica_universal WHERE id_venta='" & Val(in_venta) & "'"
                    Else
                        strCadena = "SELECT `id_venta`,`fecha_emision`,'" & in_vencimiento & "', doc_des,`documento`,`id_cliente`,`ncliente`,`direccion`,`exonerado`,`valor_venta`,`igv`,`total`,`sunat_key`,`sunat_hash`,`id_producto`,`detalle`,`observacion`,`cantidad`,`valor_neto`,`precio`,`todatl_detalle`,unidad,'" & UCase(EnLetras_moneda(in_total, in_moneda)) & "',resolucion,in_guia,id_moneda,orden_compra,tc,icbper,`ruc` FROM view_factura_electronica WHERE id_venta='" & Val(in_venta) & "'"
                    End If
                    
                    Call ConfiguraRst(strCadena)
                   strCadena = "SELECT id_venta,tipo_venta,descripcion,monto,cuotas FROM view_venta_pago WHERE id_venta='" & in_venta & "'"
                   Call ConfiguraRstK(strCadena)
                   If in_tipo_material_servicio = "01" Then 'MATERIAL
                      Ans = ShowMultiReport(rst, "factura_elec", paramy, App.Path + "\Reportes\", , , , , rstK, "modalidad_pago")
                   Else
                      Ans = ShowMultiReport(rst, "factura_elec_servicio", paramy, App.Path + "\Reportes\", , , , , rstK, "modalidad_pago")
                   End If
                   
                End If
            End If

                
            Exit Sub


strCadena = "SELECT `id_venta`,`fecha_emision`,`fecha_vencimiento`,`doc_des`,`documento`,`id_cliente`,`ncliente`,`direccion`,`exonerado`,`valor_venta`,`igv`,`total`,`sunat_key`,`sunat_hash`,`id_producto`,`detalle`,`referencia`,`cantidad`,`valor_neto`,`precio`,`todatl_detalle`,'---','" & UCase(EnLetras_moneda(in_total, in_moneda)) & "',resolucion,in_guia,id_moneda,'-','-','-',`ruc` FROM view_factura_electronica WHERE id_venta='" & Val(in_venta) & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "factura_elec", , App.Path + "\Reportes\")




            Exit Sub
        
    Case 4
        'If id_doc = KEY_RBOINGRESO Then
        '     Call impresion_formato_4_rboingreso(id_doc, serie, Numero)
        'Else
          If id_doc = "0009" Then
           strCadena = "SELECT * FROM movimiento_transferencia WHERE serie='" & serie & "' and numero='" & numero & "' and ruc='" & KEY_RUC & "' LIMIT 1"
           Call ConfiguraRstL(strCadena)
           If rstL.RecordCount > 0 Then
              Call guia_remision_tiketera(rstL("id_transferencia"))
              Exit Sub
           End If
       End If
            If KEY_FACTURACION_ELECTRONICA = "si" Then
                Call impresion_tiketera_electronica(id_doc, serie, numero)
            Else
                Call impresion_tiketera(id_doc, serie, numero)
            End If
    Case 0
        
        
        
        Select Case id_doc
            Case "0096" ' ORDEN DE PAGO
                   Call impresion_orden_pago(id_doc, serie, numero, Direccion)
            
            Case KEY_BOLETA               ' boleta personalizada
                
                If id_tipo_factura = "00003" Then
                    Call impresion_formato_boleta_suelta(id_doc, serie, numero, Direccion)
                    Exit Sub
                End If
                If id_tipo_factura = "00001" Then
                    'Call impresion_formato_boleta_suelta(id_doc, serie, Numero, direccion) ' KOREA MOTOS OK
                    'Call impresion_formato_1_boleta(id_doc, serie, Numero)
                    Call impresion_formato_boleta_daniel(id_doc, serie, numero)
                    Exit Sub
                Else
                    If KEY_FACTURACION_DETALLADA = "si" Then
                        Call impresion_formato_boleta_suelta_serie(id_doc, serie, numero, Direccion)
                        Exit Sub
                    Else
                        Call impresion_formato_boleta_suelta_serie_rep(id_doc, serie, numero, Direccion)
                        Exit Sub
                    End If
                End If
                
            Case "0054"    ' recibos personalizados
                    
                        Select Case id_tipo_factura
                            Case "00003"
                                Call impresion_formato_recibo_suelta_per(id_doc, serie, numero, Direccion)
                            Case "00002"
                                Call impresion_formato_recibo_suelta_serie(id_doc, serie, numero, Direccion)
                            Case "00001"
                                Call impresion_formato_recibo_suelta_serie(id_doc, serie, numero, Direccion)
                            
                        End Select
            
                    
                    
                    
                   

            Case "0099"    ' proform pedido
                
                If id_tipo_factura = "00001" Then
                    Call impresion_formato_boleta_suelta(id_doc, serie, numero, Direccion)
                Else
                     If KEY_FACTURACION_DETALLADA = "si" Then
                        Call impresion_formato_boleta_suelta_serie(id_doc, serie, numero, Direccion)
                     Else
                        Call impresion_formato_boleta_suelta_serie_rep(id_doc, serie, numero, Direccion)
                     End If
                End If
            
            Case "0104"
                  Call impresion_formato_boleta_suelta_serie_rep(id_doc, serie, numero, Direccion)
            
            Case KEY_GUIA
                
                
       
                
                
                
                
                If id_tipo_factura = "00001" Then
                   
                   'Call impresion_formato_guia_suelta2(id_doc, serie, Numero) ' LENIN OLIVOS
                   If KEY_RUC = "20219520281" Then
                        Call impresion_formato_guia_sr_montana(id_doc, serie, numero)
                        Exit Sub
                   End If
                   
                   
                   
                   If KEY_RUC = "20561193291" Then
                        Call impresion_formato_grupo_jm(id_doc, serie, numero)
                   Else
                        Call impresion_formato_guia_daniel(id_doc, serie, numero)
                   End If
                   
                   
                  '
                   
                   'Call guia_remision_tiketera
                   Exit Sub
                   
                   If KEY_ALM = "00003" Then
                        Call impresion_formato_guia_suelta_rep(id_doc, serie, numero) ' Call impresion_formato_guia_suelta_galvez(id_doc, serie, Numero)
                   Else
                        Call impresion_formato_guia_suelta(id_doc, serie, numero) ' Call impresion_formato_guia_suelta_galvez(id_doc, serie, Numero)
                   End If
                   
                   
                Else
                    If KEY_FACTURACION_DETALLADA = "si" Then
                        Call impresion_formato_3_guia_tienda(id_doc, serie, numero)
                    Else
                        Call impresion_formato_3_guia_tienda_rep(id_doc, serie, numero)
                    End If
                End If
                
                
                
            Case KEY_FACTURA
                  If id_tipo_factura = "00001" Then
                     ' Call impresion_formato_factura_suelta(id_doc, serie, Numero) ' KOREA MOTOSCORRECTO
                     'Call impresion_formato_2_factura(id_doc, serie, Numero) ' LENIN OLIVOS 13-09-2015
                    Call impresion_formato_grupo_jm(id_doc, serie, numero)
                     Call impresion_formato_factura_daniel(id_doc, serie, numero) ' LENIN OLIVOS 13-09-2015
                  Else
                    If KEY_FACTURACION_DETALLADA = "si" Then
                        Call impresion_formato_factura_suelta_serie(id_doc, serie, numero)
                    Else
                        Call impresion_formato_factura_suelta_serie_rep(id_doc, serie, numero)
                    End If
                  End If
            Case "0007"
                   If id_tipo_factura = "00003" Then
                      Call impresion_formato_nota_suelta(id_doc, serie, numero)
                      Exit Sub
                   End If
                  If id_tipo_factura = "00001" Then
                     Call impresion_formato_nota_suelta(id_doc, serie, numero)
                  End If
                  If id_tipo_factura = "00002" Then
                    Call impresion_formato_nota_suelta_serie_rep(id_doc, serie, numero)
                  End If
        End Select
End Select
End Sub

Public Sub impresion_consolidado_ticket(ByVal dni As String, ByVal turno As String, ByVal fecha_ini As String, ByVal fecha_fin As String, ByVal id_alm As String)
Dim nombre_paciente As String
Dim id_producto As String
Dim ptotal As Double, direccion_destino As String, impresiones As Integer
Dim nn As String, id_agenda As Double, horaf As String, fechad As String

Dim in_acumulado_maquina As Double
Dim in_acumulado_descuento As Double



   Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
  ' Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    

                            If KEY_IMPRESORA = "si" Then
                                Call funSeleccionImpresora(KEY_PRINTER)            'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            Else
                                Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            End If
    
    
    'Printer.DeviceName = Trim("EPSON TM-U220 Receipt")
    'If Prt.DeviceName = "EPSON TM-U220 Receipt" Then
     '       Set Printer = Prt
      '  End If
        
   ' If id_doc <> "0109" Then
    '    Printer.Font.name = "control"
     '   Printer.Print "A"
    'End If
    'Tahoma
    Printer.Font.name = "Tahoma"
    Printer.Font.Size = "7"
    'Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "7"
    
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
    Printer.Print ""
    'Printer.Font.Bold = True
    'Printer.Font.Size = "30"
    
    Printer.Print Tab(1); "*******     ARQUEO DE CAJA   *********"
    'Printer.Font.Size = "10"
    Printer.Print Tab(0); "-----------------------------------------"
    Printer.Print Tab(0); KEY_EMPRESA
    Printer.Print Tab(0); "RUC   :" & KEY_RUC
    Printer.Print Tab(0); "FECHA :" & KEY_FECHA
    Printer.Print Tab(0); "-----------------------------------------"
    Printer.Print Tab(0); "USUARIO :" + Space(2) + KEY_VENDEDOR
    Printer.Print Tab(0); ""
    Printer.Print Tab(0); "*************** TURNO *******************"
    Printer.Print Tab(0); ""
    strCadena = "SELECT * FROM turno WHERE id_turno='" & turno & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); "********* " & Space(1) & rst("descripcion") & Space(1) & " *********"
    Else
        Printer.Print Tab(0); "********* " & Space(1) & "NO SELECCIONADO" & Space(1) & " **********"
    End If
    
     
    
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); "------------------------------------------------------------------------------"
        Printer.Print Tab(0); "DESDE:" & Format(fecha_ini, "dd-mm-YYYY") & Space(2) & Format(fecha_fin, "dd-mm-YYYY")
        Printer.Print Tab(0); ""
        
        strCadena = "SELECT id_registro,descripcion FROM forma_pago_detalle WHERE id_alm LIKE '%" & id_alm & "%' and  ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           in_acumulado_descuento = 0
           in_acumulado_maquina = 0
           For i = 0 To rst.RecordCount - 1
                Printer.Print Tab(0); "------------------------------------------------------------------------------"
                Printer.Print Tab(0); "===== " & rst("descripcion")
                
                strCadena = "SELECT fecha_emision,documento,id_cliente,ncliente,id_forma_pago,id_tarjeta,tarjeta,nformapago,anulado,tipo_movimiento,monto_caja,id_venta,id_alm,total " & _
                "  FROM view_reporte_detallado_ultimate WHERE anulado='no' and  id_forma_pago='" & rst("id_registro") & "' and  afecta_factura='no' and  afecta_caja='si' and id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and  id_vendedor LIKE '%" & Trim(dni) & "%' AND  id_alm LIKE  '%" & id_alm & "%' and  fecha_emision>='" & Format(fecha_ini, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(fecha_fin, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' order by  fecha_emision asc,id_doc ASC,serie ASC,numero ASC"
                Call ConfiguraRstK(strCadena)
                If rstK.RecordCount > 0 Then
                   rstK.MoveFirst
                   in_monto = 0
                   in_original = 0
                   Printer.Print Tab(0); "                                                          P.MAQ    |   P.DES"
                   For j = 0 To rstK.RecordCount - 1
                       
                       in_orig = get_costo_comprobante(rstK("id_venta"), rstK("id_alm"))
                       
                       If rstK("monto_caja") <> rstK("total") Then
                         in_orig = in_orig * rstK("monto_caja") / rstK("total")
                       
                       End If
                      
                       
                       Printer.Print Tab(0); Mid(rstK("documento") & Space(30), 1, 40) & Mid(Format(in_orig, "#,##0.0000") & Space(15), 1, 15) & Mid(Format(rstK("monto_caja"), "#,##0.0000") & Space(15), 1, 15)
                       in_monto = in_monto + rstK("monto_caja")
                       in_original = in_original + in_orig
                       in_acumulado_descuento = in_acumulado_descuento + rstK("monto_caja")
                       in_acumulado_maquina = in_acumulado_maquina + in_orig
                       
                       rstK.MoveNext
                   Next j
                   Printer.Print Tab(0); "                                    -------------------------------------------------------"
                   Printer.Print Tab(0); "       =====  TOTAL  " & Mid(rst("descripcion") + "  :" & Space(30), 1, 20) & Mid(Format(in_original, "#,##0.0000"), 1, 20) & Space(5) & Mid(Format(in_monto, "#,##0.0000"), 1, 20)
                   Printer.Print Tab(0); "       =====  DIFERENCIA           :" & Space(8) & Mid(Format(in_original - in_monto, "#,##0.0000"), 1, 20)
                End If
                rst.MoveNext
           Next i
        End If
        
        
        
                   Printer.Print Tab(0); ""
                   Printer.Print Tab(0); ""
                   Printer.Print Tab(0); ""
                   Printer.Print Tab(0); ""
                   Printer.Print Tab(0); "                                    -------------------------------------------------------"
                   Printer.Print Tab(0); "       =====  ACUMULADO MAQUINA             :" & Space(8) & Mid(Format(in_acumulado_maquina, "#,##0.0000"), 1, 20)
                   Printer.Print Tab(0); "       =====  ACUMULADO DESCUENTO           :" & Space(8) & Mid(Format(in_acumulado_descuento, "#,##0.0000"), 1, 20)
                   
        
        
        Printer.EndDoc
        Exit Sub
 
End Sub

Private Function get_total_comprobante(ByVal in_venta As String) As Double
strCadena = "SELECT sum(cantidad*precio) FROM movimiento_venta_detalle WHERE obsequio='no' and  id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
    get_total_comprobante = rstIN(0)
End If
End Function
Public Sub impresion_comanda(ByVal in_comanda As String, ByVal in_reserva As String)
   Dim in_cliente As String
   Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
  ' Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    

                            If KEY_IMPRESORA = "si" Then
                                Call funSeleccionImpresora(KEY_PRINTER)            'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            Else
                                Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            End If
    
    
    'Printer.DeviceName = Trim("EPSON TM-U220 Receipt")
    'If Prt.DeviceName = "EPSON TM-U220 Receipt" Then
     '       Set Printer = Prt
      '  End If
        
   ' If id_doc <> "0109" Then
    '    Printer.Font.name = "control"
     '   Printer.Print "A"
    'End If
    'Tahoma
    Printer.Font.name = "Tahoma"
    Printer.Font.Size = "7"
    'Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "7"
    
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
    Printer.Print ""
    'Printer.Font.Bold = True
    'Printer.Font.Size = "30"
    
    strCadena = "SELECT * FROM restaurant_comanda WHERE id_comanda='" & Val(in_comanda) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    Printer.Print Tab(0); "========================================="
    Printer.Print Tab(0); KEY_EMPRESA
    Printer.Print Tab(0); "RUC   :" & KEY_RUC
    Printer.Print Tab(0); "========================================="
    Printer.Print Tab(0); "ORDEN COMANDA:" & Format(rst("id_comanda"), "000000")
    Printer.Print Tab(0); "FECHA        :" & Format(rst("fecha_hora"), "dd-mm-YYYY") + Space(3) & Format(rst("hora"), "hh:mm:ss")
    Printer.Print Tab(0); "MOZO         :" & get_persona(rst("dni_save"))
    Printer.Print Tab(0); "========================================="
    strCadena = "SELECT * FROM view_habitacion_piso WHERE id_reserva='" & Val(in_reserva) & "'"
    Call ConfiguraRstIN(strCadena)
    If rstIN.RecordCount > 0 Then
        Printer.Print Tab(0); "HABITACION   :" & rstIN("descripcion")
        Printer.Print Tab(0); "HUESPED      :" & rstIN("cliente")
        in_cliente = rstIN("cliente")
        Printer.Print Tab(0); "========================================="
    End If
      
        
        strCadena = "SELECT * FROM view_comanda_detalle WHERE id_comanda='" & Val(in_comanda) & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstIN(strCadena)
        If rstIN.RecordCount > 0 Then
           rstIN.MoveFirst
           in_monto = 0
           For i = 0 To rstIN.RecordCount - 1
                
                       Printer.Print Tab(0); rstIN("id_producto") & Space(2) & Mid(rstIN("nombre_prod") & Space(30), 1, 25) & Mid(Format(rstIN("cantidad"), "#,##0.00") & Space(15), 1, 5) & Mid(Format(rstIN("precio"), "#,##0.00") & Space(15), 1, 10) & Mid(Format(rstIN("total"), "#,##0.00") & Space(15), 1, 10)
                       in_monto = in_monto + rstIN("total")
                       
                       rstIN.MoveNext
                   Next i
                   Printer.Print Tab(0); "========================================="
                   Printer.Print Tab(20); "=====  TOTAL  " & Mid(Format(in_monto, "#,##0.00"), 1, 20)
                   
                End If
              
        End If
        
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        
        Printer.Print Tab(5); "-----------------------------------------"
        Printer.Print Tab(5); "HUESPED :" + Space(2) + in_cliente
    
        
        Printer.EndDoc
        Exit Sub
 
End Sub


Public Sub impresion_consolidado_ticket_normal(ByVal dni As String, ByVal turno As String, ByVal fecha_ini As String, ByVal fecha_fin As String, ByVal id_alm As String)
Dim nombre_paciente As String
Dim id_producto As String
Dim ptotal As Double, direccion_destino As String, impresiones As Integer
Dim nn As String, id_agenda As Double, horaf As String, fechad As String
 
 Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    '************ IMPRESORA X DEFECTO *********
                             If KEY_IMPRESORA = "si" Then
                                Call funSeleccionImpresora(KEY_PRINTER)            'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                           Else
                               Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                          End If
    '*******************************************
    Printer.Font.name = "FontB11"
    'Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "7"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
 
    'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    If id_doc <> "0109" Then
   '     Printer.Font.name = "control"
   '     Printer.Print "A"
    End If
    
   
 
    
    Printer.Print ""
    'Printer.Font.Bold = True
    'Printer.Font.Size = "30"
    
    Printer.Print Tab(1); "*******     ARQUEO DE CAJA   *********"
    'Printer.Font.Size = "10"
    Printer.Print Tab(0); "-----------------------------------------"
    Printer.Print Tab(0); KEY_EMPRESA
    Printer.Print Tab(0); "RUC   :" & KEY_RUC
    Printer.Print Tab(0); "FECHA :" & KEY_FECHA
    Printer.Print Tab(0); "-----------------------------------------"
    Printer.Print Tab(0); "USUARIO :" + Space(2) + KEY_VENDEDOR
    Printer.Print Tab(0); ""
    Printer.Print Tab(0); "*************** TURNO *******************"
    Printer.Print Tab(0); ""
    strCadena = "SELECT * FROM turno WHERE id_turno='" & turno & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); "********* " & Space(1) & rst("descripcion") & Space(1) & " *********"
    Else
        Printer.Print Tab(0); "********* " & Space(1) & "NO SELECCIONADO" & Space(1) & " **********"
    End If
    
     
    
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); "------------------------------------------------------------------------------"
        Printer.Print Tab(0); ""
        
        strCadena = "SELECT id_registro,descripcion FROM forma_pago_detalle WHERE   ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           
           For i = 0 To rst.RecordCount - 1
                Printer.Print Tab(0); "------------------------------------------------------------------------------"
                Printer.Print Tab(0); "===== " & rst("descripcion")
                
                strCadena = "SELECT fecha_emision,documento,id_cliente,ncliente,id_forma_pago,id_tarjeta,tarjeta,nformapago,anulado,tipo_movimiento,monto_caja,id_venta,id_alm " & _
                "  FROM view_reporte_detallado_ultimate WHERE anulado='no' and  id_forma_pago='" & rst("id_registro") & "' and  afecta_factura='no' and  afecta_caja='si' and id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and  id_vendedor LIKE '%" & Trim(dni) & "%' AND  id_alm LIKE  '%" & id_alm & "%' and  fecha_emision>='" & Format(fecha_ini, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(fecha_fin, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' order by  fecha_emision asc,id_doc ASC,serie ASC,numero ASC"
                Call ConfiguraRstK(strCadena)
                If rstK.RecordCount > 0 Then
                   rstK.MoveFirst
                   in_monto = 0
                   in_original = 0
                   Printer.Print Tab(0); "                                                          MONTO"
                   For j = 0 To rstK.RecordCount - 1
                      ' in_orig = get_costo_comprobante(rstK("id_venta"), rstK("id_alm"))
                       Printer.Print Tab(0); Format(rstK("fecha_emision"), "dd-mm-YYYY") & Space(1) & Mid(rstK("documento") & Space(30), 1, 24) & Mid(Format(rstK("monto_caja"), "#,##0.00") & Space(15), 1, 15)
                       in_monto = in_monto + rstK("monto_caja")
                       rstK.MoveNext
                   Next j
                   Printer.Print Tab(0); "                                    -------------------------------------------------------"
                   Printer.Print Tab(0); "       =====  TOTAL  " & Mid(rst("descripcion") + "  :" & Space(30), 1, 10) & Mid(Format(in_monto, "#,##0.0000"), 1, 20)
                   
                End If
                rst.MoveNext
           Next i
        End If
        
        
        Printer.EndDoc
        Exit Sub
 
End Sub


Private Function get_costo_comprobante(ByVal in_venta As String, ByVal in_alm As String) As Double

strCadena = "SELECT sum(total) FROM view_precio_original WHERE id_venta='" & Val(in_venta) & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    get_costo_comprobante = rstL(0)
Else
    get_costo_comprobante = 0
End If
End Function
Public Sub imprecion_kardex_cabecera()
    Printer.Print Tab(0); "FORMATO 13.1 REGISTRO DEL INVENTARIO PERMANENTE VALORIZADO-DETALLE DEL INVENTARIO VALORIZADO"
    Printer.Print Tab(0); "PERIODO                                          :2016"
    Printer.Print Tab(0); "RUC                                              :" & KEY_RUC
    Printer.Print Tab(0); "APELLIDOS Y NOMBRES, DENOMINACION O RAZON SOCIAL :" & KEY_EMPRESA
    Printer.Print Tab(0); "ESTABLECIMIENTO                                  :" & KEY_DIRECCION
    Printer.Print Tab(0); "TIPO                                             :01-MERCADERIA"
    Printer.Print Tab(0); "METODO DE EVALUACION                             :PROMEDIO"
    Printer.Print Tab(0); "EXPRESDO EN                                      :SOLES"
    Printer.Print Tab(0); "====================================================================================================================="
    
    
End Sub
Public Sub impresion_kardex_valorizado(ByVal fecha_ini As Date, ByVal fecha_fin As Date, ByVal in_alm As String)

    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    'Printer.CurrentX = 0
    'Printer.CurrentY = 0
 
    Printer.ScaleWidth = 10#
    Printer.ScaleHeight = 20#
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
   '------------------------------------------------------------------"
    
   'Call imprecion_kardex_cabecera
    Printer.Print Tab(0); "FORMATO 13.1 REGISTRO DEL INVENTARIO PERMANENTE VALORIZADO-DETALLE DEL INVENTARIO VALORIZADO"
    Printer.Print Tab(0); "PERIODO                                          :2016"
    Printer.Print Tab(0); "RUC                                              :" & KEY_RUC
    Printer.Print Tab(0); "APELLIDOS Y NOMBRES, DENOMINACION O RAZON SOCIAL :" & KEY_EMPRESA
    Printer.Print Tab(0); "ESTABLECIMIENTO                                  :" & KEY_DIRECCION
    Printer.Print Tab(0); "TIPO                                             :01-MERCADERIA"
    Printer.Print Tab(0); "METODO DE EVALUACION                             :PROMEDIO"
    Printer.Print Tab(0); "EXPRESDO EN                                      :SOLES"
    Printer.Print Tab(0); "====================================================================================================================="
   strCadena = "SELECT DISTINCT id_producto FROM kardex WHERE fecha_emision>='" & Format(fecha_ini, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(fecha_fin, "YYYY-mm-dd") & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY id_producto ASC limit 27 "
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      rst.MoveFirst
      For i = 0 To rst.RecordCount - 1
          in_producto = get_producto(rst("id_producto"))
          Printer.Print Tab(0); "-----------------------------------------"
          ' recorro producto x producto
                    strCadena = "SELECT * FROM kardex WHERE id_producto='" & rst("id_producto") & "' and fecha_emision>='" & Format(fecha_ini, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(fecha_fin, "YYYY-mm-dd") & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC "
                    Call ConfiguraRstK(strCadena)
                    If rstK.RecordCount > 0 Then
                       rstK.MoveFirst
                       For j = 0 To rstK.RecordCount - 1
                           in_codigo = rst("id_producto")
                           in_fecha = Format(rstK("fecha_emision"), "dd-mm-YYYY")
                           in_doc = Format(Val(rstK("id_doc")), "00")
                           in_serie = rstK("id_serie")
                           in_numero = rstK("id_numero")
                           in_tipo = rstK("id_tipo_movimiento")
                           'ingresos
                           If rstK("cantidad_real") > 0 Then
                              ing_cantidad = rstK("cantidad")
                              ing_costo = rstK("costo_unitario")
                              ing_total = rstK("cantidad") * rstK("costo_unitario")
                           Else
                              ing_cantidad = ""
                              ing_costo = ""
                              ing_total = ""
                           End If
                           'salidas
                           If rstK("cantidad_real") < 0 Then
                              sal_cantidad = rstK("cantidad")
                              sal_costo = rstK("costo_unitario")
                              sal_total = rstK("cantidad") * rstK("costo_unitario")
                           Else
                              ing_cantidad = ""
                              ing_costo = ""
                              in_total = ""
                           End If
                           'saldos
                           
                           Printer.Print Tab(0); rst("id_producto") & Space(2) & in_fecha & Space(2) & in_doc & Space(2) & in_serie & Space(2) & in_numero
                           
                           If Val(Printer.CurrentY) >= 19.85333 Then
                                Printer.NewPage
                                Printer.Print Tab(0); "FORMATO 13.1 REGISTRO DEL INVENTARIO PERMANENTE VALORIZADO-DETALLE DEL INVENTARIO VALORIZADO"
                                Printer.Print Tab(0); "PERIODO                                          :2016"
                                Printer.Print Tab(0); "RUC                                              :" & KEY_RUC
                                Printer.Print Tab(0); "APELLIDOS Y NOMBRES, DENOMINACION O RAZON SOCIAL :" & KEY_EMPRESA
                                Printer.Print Tab(0); "ESTABLECIMIENTO                                  :" & KEY_DIRECCION
                                Printer.Print Tab(0); "TIPO                                             :01-MERCADERIA"
                                Printer.Print Tab(0); "METODO DE EVALUACION                             :PROMEDIO"
                                Printer.Print Tab(0); "EXPRESDO EN                                      :SOLES"
                                Printer.Print Tab(0); "====================================================================================================================="
                            End If
                           rstK.MoveNext
                       Next j
                    End If
                    
                            
           rst.MoveNext
      Next i
   End If
   
   
    
    
        
        Printer.EndDoc
        Exit Sub
 
End Sub
Public Sub impresion_kardex_valorizado_demo(ByVal fecha_ini As Date, ByVal fecha_fin As Date, ByVal in_alm As String, ByVal in_producto As String)
    
    Dim sum_ing_cant As Double
    Dim sum_ing_tot As Double
    Dim sum_sali_cant As Double
    Dim sum_sali_tot As Double
    
    
    
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.CurrentX = 0
    Printer.CurrentY = 0
 
    Printer.ScaleWidth = 10#
    Printer.ScaleHeight = 28#
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "10"
    
    '***** COMPROBANTE *****
   '------------------------------------------------------------------"
    
   Call imprecion_kardex_cabecera
    
   'strCadena = "SELECT DISTINCT movart,movnoa FROM qalmdet WHERE  movfec>='" & Format(fecha_ini, "YYYY-mm-dd") & "' and movfec<='" & Format(fecha_fin, "YYYY-mm-dd") & "'  ORDER BY movart ASC limit 40 "
   

   If in_producto = "" Then
        strCadena = "SELECT DISTINCT id_producto FROM kardex WHERE id_producto not in('00','00000') and   ruc='" & KEY_RUC & "'   ORDER BY id_producto ASC "
   Else
    strCadena = "SELECT DISTINCT id_producto FROM kardex WHERE id_producto ='" & in_producto & "' and   ruc='" & KEY_RUC & "'   ORDER BY id_producto ASC "
   End If
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    rst.MoveFirst
  
    For i = 0 To rst.RecordCount - 1
              If Val(Printer.CurrentY) >= 27.29045 Then
                   
                    Printer.NewPage
                    Call imprecion_kardex_cabecera
              End If
                
        Printer.Print Tab(0); rst("id_producto") & Space(2) & "(7 UNIDADES)" & Space(2) & get_producto(rst("id_producto"))
        If Val(Printer.CurrentY) >= 27.29045 Then
                    
                    Printer.NewPage
                    Call imprecion_kardex_cabecera
        End If
                
        Printer.Print Tab(0); ""
        If Val(Printer.CurrentY) >= 27.29045 Then
                    
                    Printer.NewPage
                    Call imprecion_kardex_cabecera
                End If
        
        strCadena = "SELECT IFNULL(sum(cantidad_real),0) FROM kardex WHERE id_producto='" & rst("id_producto") & "' and fecha_emision<'" & Format(fecha_ini, "YYYY-mm-dd") & "'  and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstP(strCadena)
        in_cantidad_ini = rstP(0)
        
        
        
        
        strCadena = "SELECT costo_promedio,fecha_emision FROM kardex WHERE ruc='" & KEY_RUC & "' and  id_producto='" & rst("id_producto") & "' and fecha_emision<'" & Format(fecha_ini, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(fecha_fin, "YYYY-mm-dd") & "' ORDER BY fecha_emision ASC,id_kardex ASC LIMIT 1"
        Call ConfiguraRstP(strCadena)
        If rstP.RecordCount > 0 Then
                in_costo_ini = rstP("costo_promedio")
                ini_total_ini = in_cantidad_ini * in_costo_ini
                in_codigo = rst("id_producto")
                in_fecha = Format(fecha_ini, "dd-mm-YYYY")
                in_doc = Mid("" + Space(2), 1, 2)
                in_serie = Mid("" + Space(20), 4, 3)
                in_numero = Mid("" + Space(20), 8, 6)
                in_tipo = "16-SALDO INICIAL"
                
                ini_cantidad = Mid(Format(in_cantidad_ini, "#,##0.00") + Space(10), 1, 7)
                ini_costo = Mid(Format(in_costo_ini, "#,##0.00") + Space(10), 1, 6)
                ini_saldo = Mid(Format(in_cantidad_ini * in_costo_ini, "#,##0.00") + Space(10), 1, 10)
                
                ing_cantidad = Mid("" + Space(10), 1, 8)
                ing_costo = Mid(" " + Space(10), 1, 8)
                ing_total = Mid("" + Space(15), 1, 10)
                
                sal_cantidad = Mid(" " + Space(10), 1, 8)
                sal_costo = Mid(" " + Space(10), 1, 8)
                sal_total = Mid(" " + Space(15), 1, 10)
                
                sum_ing_cant = 0
                sum_sali_cant = 0
                
                sum_ing_tot = 0
                sum_sali_tot = 0
                
                Printer.Print Tab(0); rst("id_producto") & Space(2) & in_fecha & Space(3) & in_doc & Space(3) & in_serie & Space(2) & in_numero & Space(3) & Mid(in_tipo & Space(10), 1, 16) & ing_cantidad & ing_costo & ing_total & sal_cantidad & sal_costo & sal_total & sal_costo & Space(4) & Mid(Format(ini_cantidad, "#,##.00") + Space(15), 1, 10) & Mid(Format(ini_costo, "#,##0.00") + Space(15), 1, 10) & Mid(Format(ini_saldo, "#,##0.00") + Space(15), 1, 10)
                
        End If
        
        strCadena = "SELECT id_producto,fecha_emision,id_serie,id_numero,id_doc,id_tipo_movimiento,cantidad,costo_promedio,cantidad_real FROM kardex WHERE ruc='" & KEY_RUC & "' and  id_producto='" & rst("id_producto") & "' and fecha_emision>='" & Format(fecha_ini, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(fecha_fin, "YYYY-mm-dd") & "'  ORDER BY fecha_emision ASC,id_kardex  ASC"
        Call ConfiguraRstP(strCadena)
        If rstP.RecordCount > 0 Then
           rstP.MoveFirst
           For j = 0 To rstP.RecordCount - 1
                If Val(Printer.CurrentY) >= 27.29045 Then
                    
                    Printer.NewPage
                    Call imprecion_kardex_cabecera
                End If
                in_codigo = rst("id_producto")
                in_fecha = Format(rstP("fecha_emision"), "dd-mm-YYYY")
                in_doc = Mid(rstP("id_doc") + Space(2), 1, 2)
                in_serie = Mid(rstP("id_serie") + Space(20), 4, 3)
                in_numero = Mid(rstP("id_numero") + Space(20), 8, 6)
                If rstP("cantidad_real") < 0 Then
                    in_tipo = "01-VENTA"
                Else
                    in_tipo = "02-COMPRA"
                End If
                'ingresos
                If rstP("cantidad_real") > 0 Then
                              ing_cantidad = Mid(Format(rstP("cantidad"), "#,##0.00") + Space(10), 1, 7)
                              ing_costo = Mid(Format(rstP("costo_promedio"), "#,##0.00") + Space(10), 1, 6)
                              ing_total = Mid(Format(rstP("cantidad") * rstP("costo_promedio"), "#,##0.00") + Space(15), 1, 10)
                              sal_cantidad = Mid(" " + Space(10), 1, 8)
                              sal_costo = Mid(" " + Space(10), 1, 8)
                              sal_total = Mid(" " + Space(15), 1, 15)
                Else
                              ing_cantidad = Mid(" " + Space(10), 1, 7)
                              ing_costo = Mid(" " + Space(10), 1, 6)
                              ing_total = Mid(" " + Space(15), 1, 10)
                End If
                           'salidas
                If rstP("cantidad_real") < 0 Then
                              sal_cantidad = Mid(Format(rstP("cantidad"), "#,##0.00") + Space(10), 1, 8)
                              sal_costo = Mid(Format(rstP("costo_promedio"), "#,##0.00") + Space(10), 1, 8)
                              sal_total = Mid(Format(rstP("cantidad") * rstP("costo_promedio"), "#,##0.00") + Space(15), 1, 15)
                              ing_cantidad = Mid(" " + Space(10), 1, 7)
                              ing_costo = Mid(" " + Space(10), 1, 6)
                              ing_total = Mid(" " + Space(15), 1, 10)
                Else
                              sal_cantidad = Mid(" " + Space(10), 1, 8)
                              sal_costo = Mid(" " + Space(10), 1, 8)
                              sal_total = Mid(" " + Space(15), 1, 15)
                End If
                           'saldos
                
                If j = 0 Then
                    
                    If rstP("cantidad_real") < 0 Then
                        saldo_cantidad_saldo = Val(ini_cantidad) - rstP("cantidad")
                        If Val(saldo_cantidad_saldo) = 0 Then
                            saldo_prom_saldo = 0
                            saldo_total_saldo = 0
                        Else
                            saldo_prom_saldo = (Val(ini_saldo) + Val(sal_total)) / Val(saldo_cantidad_saldo)
                            saldo_total_saldo = Val(saldo_cantidad_saldo * saldo_prom_saldo)
                        End If
                        
                    Else
                        saldo_cantidad_saldo = Val(ini_cantidad) + rstP("cantidad")
                        If saldo_cantidad_saldo = 0 Then
                            saldo_prom_saldo = 0
                            saldo_total_saldo = 0
                        Else
                            saldo_prom_saldo = (ini_saldo + Val(ing_total)) / saldo_cantidad_saldo
                            saldo_total_saldo = Val(saldo_cantidad_saldo * saldo_prom_saldo)
                        End If
                        
                    End If
                    
                    
                    
                
                Else
                    If rstP("cantidad_real") < 0 Then
                        saldo_cantidad_saldo = Val(saldo_cantidad_saldo) - Val(rstP("cantidad"))
                        If Val(Val(saldo_cantidad_saldo)) = 0 Then
                            saldo_prom_saldo = 0
                            saldo_total_saldo = 0
                        Else
                            saldo_prom_saldo = saldo_total_saldo / Val(saldo_cantidad_saldo)
                            saldo_total_saldo = Val(saldo_cantidad_saldo * saldo_prom_saldo)
                        End If
                        
                    Else
                        saldo_cantidad_saldo = Val(saldo_cantidad_saldo) + Val(rstP("cantidad"))
                        If saldo_cantidad_saldo = 0 Then
                            saldo_prom_saldo = 0
                            saldo_total_saldo = 0
                        Else
                            saldo_prom_saldo = saldo_total_saldo / Val(saldo_cantidad_saldo)
                            saldo_total_saldo = Val(saldo_cantidad_saldo * saldo_prom_saldo)
                        End If
                        
                    End If
                    
                    
                    
                    
                End If
                
                Printer.Print Tab(0); rst("id_producto") & Space(2) & in_fecha & Space(4) & in_doc & Space(4) & in_serie & Space(4) & in_numero & Space(5) & Mid(in_tipo & Space(10), 1, 15) & Space(5) & ing_cantidad & ing_costo & ing_total & sal_cantidad & sal_costo & sal_total & Mid(Format(saldo_cantidad_saldo, "#,##.00") + Space(15), 1, 10) & Mid(Format(saldo_prom_saldo, "#,##0.00") + Space(15), 1, 10) & Mid(Format(saldo_total_saldo, "#,##0.00") + Space(15), 1, 10)
                
                sum_ing_cant = sum_ing_cant + Val(ing_cantidad)
                sum_sali_cant = sum_sali_cant + Val(sal_cantidad)
                
                sum_ing_tot = sum_ing_tot + Val(ing_total)
                sum_sali_tot = sum_sali_tot + Val(sal_total)
                
                sum_ing_cantn = Mid(Format(sum_ing_cant, "#,##0.00") + Space(10), 1, 8)
                sum_sali_cantn = Mid(Format(sum_sali_cant, "#,##0.00") + Space(15), 1, 10)
                              
                sum_ing_totn = Mid(Format(sum_ing_tot, "#,##0.00") + Space(10), 1, 8)
                sum_sali_totn = Mid(Format(sum_sali_tot, "#,##0.00") + Space(15), 1, 10)
                
                
                
                
                
                
                rstP.MoveNext
           Next j
           If Val(Printer.CurrentY) >= 27.29045 Then
                    
                    Printer.NewPage
                    Call imprecion_kardex_cabecera
                End If
           Printer.Print Tab(66); "-------" & Space(8) & "----------" & Space(8) & "-------" & Space(8) & "----------" & Space(10) & "----------"
           If Val(Printer.CurrentY) >= 27.29045 Then
                    
                    Printer.NewPage
                    Call imprecion_kardex_cabecera
                End If
           Printer.Print Tab(66); sum_ing_cantn & Space(6) & sum_ing_totn & sum_sali_cantn & Space(6) & sum_sali_totn
        End If
        
        
                           
        rst.MoveNext
    
    Next i
   End If
   
   
    
    
        
        Printer.EndDoc
        Exit Sub
 
 
 
End Sub



Public Sub impresion_detallado(ByVal strCadena As String)
Dim nombre_paciente As String
Dim id_producto As String
Dim ptotal As Double, direccion_destino As String, impresiones As Integer
Dim nn As String, id_agenda As Double, horaf As String, fechad As String
Dim Persona As String
   'Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
  ' Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    

                            If KEY_IMPRESORA = "si" Then
                                Call funSeleccionImpresora(KEY_PRINTER)            'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            Else
                                Printer.TrackDefault = True 'siempre apunta a la impresora predeter
                            End If
    
    
    'Printer.DeviceName = Trim("EPSON TM-U220 Receipt")
    'If Prt.DeviceName = "EPSON TM-U220 Receipt" Then
     '       Set Printer = Prt
      '  End If
        
   ' If id_doc <> "0109" Then
    '    Printer.Font.name = "control"
     '   Printer.Print "A"
    'End If
    'Tahoma
    'Printer.Font.name = "Tahoma"
    Printer.Font.name = "FontB11"
    Printer.Font.Size = "10"
    'Printer.Font.name = "Draft 17cpi"
    'Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
    Printer.Print ""
    'Printer.Font.Bold = True
    'Printer.Font.Size = "30"
    
    Printer.Print Tab(1); "******* ARQUEO DE CAJA *********"
    'Printer.Font.Size = "10"
    Printer.Print Tab(0); "-----------------------------------------"
    Printer.Print Tab(0); KEY_EMPRESA
    Printer.Print Tab(0); "RUC   :" & KEY_RUC
    Printer.Print Tab(0); "FECHA :" & KEY_FECHA
    Printer.Print Tab(0); "-----------------------------------------"
    Printer.Print Tab(0); "USUARIO :" + Space(2) + KEY_VENDEDOR
    Printer.Print Tab(0); "-----------------------------------------"
    
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        Printer.Print Tab(0); ""
        ptotal = 0
       For i = 0 To rstT.RecordCount - 1
            If rstT("anulado") = "si" Then
                Persona = "*** ANULADO ***"
            Else
                Persona = Mid(rstT("nombre_completo") + Space(50), 1, 15)
            End If
            Printer.Print Tab(0); Mid(rstT("comprobante") + Space(50), 1, 17) + "-" + Persona + Space(1) + Format(rstT("total"), "#,##0.00")
            ptotal = ptotal + rstT("total")
            rstT.MoveNext
        Next i
        Printer.Print Tab(0); "-----------------------------------------"
        Printer.Print Tab(0); "MONTO COBRANZA:" + Space(1) + "S/." & Format(ptotal, "#,##0.00")
        Printer.Print Tab(0); "-----------------------------------------"
        Printer.Print Tab(0); "HORA IMPRESION:" + Space(1) + str(Time)
        Printer.Print ""
        Printer.Print Tab(0); "   ****     VITEKEY SALUD      ****"
        Printer.Print ""
        Printer.EndDoc
        Exit Sub
 End If
End Sub

Private Sub impresion_formato_4_rboingreso(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
Dim MiArray() As String
Dim ctotal As Double
   Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "FontB11"
    Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    '
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM mis_cuentas_det WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
       
    Printer.Print ""
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(1); rst("documento")
    Printer.Print Tab(1); Mid("FECHA OPERACION" + Space(30), 1, 20) & ":" & formato_item(Day(rst("fecha")), 2) & Space(3) & formato_item(Month(rst("fecha")), 2) + Space(3) + str(Year(rst("fecha")))
    Printer.Print Tab(1); Mid("RUC/DNI" + Space(30), 1, 20) & ":" & rst("id_persona")
    Printer.Print Tab(1); Mid("NOMBRE RAZON" + Space(30), 1, 20) & ":" & UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_persona")))
    Printer.Print Tab(1); Mid("DIRECCION" + Space(30), 1, 20) & ":" & Mid(UCase(BDBuscarCampo("persona", "direccion", "dni", rst("id_persona"))) + Space(80), 1, 60)
    Printer.Print "" 'Tab(10); 'L 9
    
    Printer.Print Tab(1); "========================================="
    MiArray = Split(Trim(rst("glosa")), "/")
    For i = LBound(MiArray) To UBound(MiArray)
        Printer.Print Tab(1); Trim(MiArray(i))
    Next i
    Printer.Print Tab(1); "========================================="
   
    strCadena = "SELECT P.nombre_prod,U.abreviatura,D.cantidad,D.precio,D.total FROM movimiento_venta V,movimiento_venta_detalle D,unidad U,producto P   WHERE V.id_venta=D.id_venta AND V.ruc='" & KEY_RUC & "' " & _
    " AND D.ruc='" & KEY_RUC & "' AND V.id_cliente='" & rst("id_persona") & "' AND V.seleccion='si' AND D.id_producto=P.id_producto AND P.ruc='" & KEY_RUC & "' " & _
    " AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' ORDER BY D.id_detalle_venta"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        rstT.MoveFirst
        ctotal = 0
        For i = 0 To rstT.RecordCount - 1
            descripcion = Mid(rstT("nombre_prod") + Space(30), 1, 22)
            Und = Mid(rstT("abreviatura") + Space(10), 1, 4)
            CANT = Mid(Format(rstT("cantidad"), "#,##0.00") + Space(10), 1, 6)
            precio = Mid(Format(rstT("precio"), "#,##0.00") + Space(10), 1, 6)
            ttTotal = Mid(Format(rstT("total"), "#,##0.00") + Space(10), 1, 6)
          Printer.Print Tab(0); descripcion + Space(1); Und + Space(1) + CANT + Space(1) + ttTotal
          ctotal = ctotal + rstT("total")
          rstT.MoveNext
        Next i
    End If
    Printer.Print Tab(1); "========================================="
    Printer.Print Tab(1); "MONTO ACUMULADO DEUDA  :" + Space(2) + Mid(Format(ctotal, "#,##0.00"), 1, 10)
    Printer.Print Tab(1); "========================================="
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print Tab(1); "MONTO CANCELADO  :" + Space(2) + Mid(Format(rst("monto"), "#,##0.00"), 1, 10)
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print Tab(1); UCase(EnLetras(rst("monto")))
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print Tab(1); "ATENDIDO POR :"; KEY_VENDEDOR + Space(3) + str(Time)
    Printer.Print Tab(1); "========================================="
          
    Printer.EndDoc
    Exit Sub
End Sub

Private Sub impresion_tiketera(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String, impresiones As Integer
Dim in_detraccion As Boolean

Dim nn As String
Dim nserial As String
Dim in_venta As Double
Dim in_direccion  As String
Dim in_moneda As String
Dim in_simbolo As String
Dim in_referencia As String
Dim in_observacion As String
Dim in_filas As Integer
Dim in_dni As String
in_observacion = ""
  ' Call CargaDefConfigEpsonTM
    
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    If id_doc <> "0109" Then
   '     Printer.Font.name = "control"
   '     Printer.Print "A"
    End If
    
   
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    '
    strCadena = "SELECT serial,fuente FROM almacen_comprobante WHERE id_doc='" & id_doc & "' and serie='" & serie & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    nserial = rst("serial")
    'Printer.Font.name = "FontB11"
    'Printer.Font.Size = "10"
    Printer.Font.name = Trim(rst("fuente"))
    Printer.Font.Size = "7"
    
    
    
    ':::::::::: COMPROBANTE ::::::::::::::::::::::::::::::::::::::::::::::::
    strCadena = "SELECT v.id_venta,v.id_moneda,v.id_doc_fact,v.fecha_fact,v.observacion,v.serie_fact,v.numero_fact,v.id_doc,v.serie,v.numero,v.documento,v.exonerado,v.valor_venta,v.igv,v.id_vendedor,v.hora,v.fecha_emision,p.nombre_completo,v.id_cliente,v.ncliente,v.direccion as ndireccion,p.direccion,v.impresiones,v.id_tipo_factura,v.total,v.observacion,v.saldo,v.id_comprobante,v.id_copropietario,p.id_departamento,p.id_provincia,p.id_distrito,sunat_hash,motivo_nota,v.id_forma_pago,v.cuotas,v.fecha_vencimiento FROM movimiento_venta v LEFT JOIN persona p ON v.id_cliente=p.dni WHERE  v.id_doc='" & id_doc & "' AND v.serie='" & serie & "' AND v.numero='" & numero & "' AND v.ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    in_observacion = rst("observacion")
    in_venta = rst("id_venta")
    in_moneda = rst("id_moneda")
    in_dni = rst("id_cliente")
    
    If in_moneda = "00001" Then
       in_simbolo = "S/ "
    Else
       in_simbolo = "US$ "
       
    End If
    in_direccion = UCase(rst("direccion"))
    
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    
    Printer.FontBold = False
     
    'Printer.Print "" 'Tab(10); 'L 2
    
    Set FrmVentas.Picture1.Picture = LoadPicture(App.Path & "\Imagenes\logo11.jpg")
    Printer.PaintPicture FrmVentas.Picture1.Picture, 4, 0
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    'Printer.Print ""
    'Printer.Print "" 'Tab(10); 'L 2
    If rst("impresiones") > 1 Then
        If KEY_ALM <> "00001" Then
            Printer.Print Tab(1); "COPIA DE LA ORIGINAL" + Space(3) + rst("documento")
            Printer.Print Tab(1); "-------------------------------------------------------------"
        End If
    End If
    
    
    If Mid(KEY_RUC, 1, 2) = "10" Then
        Printer.Print Tab(0); KEY_NOMBRE_COMERCIAL
        'Printer.Print ""
        Printer.Print Tab(0); "DE:" & KEY_EMPRESA
        Printer.Print Tab(0); "-----------------------------------------------------------------"
    Else
        Printer.Print Tab(0); KEY_EMPRESA
    End If
    
    Printer.Print Tab(0); Mid(KEY_DIRECCION, 1, 36)
    Printer.Print Tab(0); Mid(KEY_DIRECCION, 37, 36)
    If KEY_ALM <> "00001" Then
       Printer.Print Tab(0); "-----------------------------------------------------------------"
       Printer.Print Tab(0); "SUC:" & KEY_DIRECCION_ALM
       Printer.Print Tab(0); "-----------------------------------------------------------------"
    End If
    If KEY_RUC = "20600133889" Then
        Printer.Print Tab(0); "-----------------------------------------------------------------"
       Printer.Print Tab(0); "SUC:" & "AV. AGRICULTURA KM.1 VIÑA DEL MAR."
       Printer.Print Tab(0); "-----------------------------------------------------------------"
    End If
    
    If KEY_RUBRO = "00025" Then
        Printer.Print Tab(0); "TELF:074-204985   CELL:#949458350"

    Else
        Printer.Print Tab(0); "TELF:" & KEY_TELEFONO

    End If
    
    Printer.Print Tab(0); "E-MAIL  :" & KEY_EMAIL
    Printer.Print Tab(1); ""
    Printer.Print Tab(0); "RUC          :" & KEY_RUC
    Printer.Print Tab(0); "FECHA EMISION:" & Format(rst("fecha_emision"), "dd-mm-YYYY")
    If rst("id_forma_pago") = "02" And rst("cuotas") < 1 Then
    Printer.Print Tab(0); "FECHA VENCIMIENTO:" & Format(rst("fecha_vencimiento"), "dd-mm-YYYY")
    End If
    
    
    Printer.Print Tab(0); ""
    Printer.Print Tab(0); "-----------------------------------------------------------------"
    
    
   ' Printer.Print Tab(0); "TICKET-" & rst("documento")
   ' Printer.Print Tab(0); "SERIE  :" + nserial
   ' Printer.Print Tab(0); "Nº AUT :" + "0071845087225"
    
    
    
    
    in_electronico = get_electronico(rst("id_doc"), rst("serie"))
    
            
    'Printer.Font.Bold = True
    If rst("id_doc") = "0002" Then
        Printer.Print Tab(0); BDBuscarCampo("comprobantes", "doc_abrev", "id_doc", rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
    Else
        If KEY_FACTURACION_ELECTRONICA = "si" And in_electronico = "si" Then
            
            Select Case rst("id_doc")
                   Case "0003"
                        Printer.Print Tab(7); "BOLETA DE VENTA ELECTRONICA"
                   Case "0001"
                        Printer.Print Tab(7); "FACTURA DE VENTA ELECTRONICA"
                   Case "0007"
                        Printer.Print Tab(7); "NOTA DE CREDITO ELECTRONICA"
            End Select
                        Printer.Print Tab(16); rst("serie") & "-" & rst("numero")
            
        Else
            
            If KEY_RUBRO = "00025" Then
                Printer.Print Tab(0); "TICKET " & BDBuscarCampo("comprobantes", "doc_abrev", "id_doc", rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
                Printer.Print Tab(0); "SERIE: FFFF018300"
                Printer.Print Tab(0); "Nro Aut: 0073845113276"
            Else
                'Printer.Print Tab(0); get_comprobante_des(rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
                Printer.Print Tab(0); "TICKET " & get_comprobante_des(rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
                If rst("id_doc") <> "0054" Then
                   ' Printer.Print Tab(0); "SERIE: FFCF214706"
                   ' Printer.Print Tab(0); "Nro Aut: 0073845123567"
                End If
            End If
            
        End If
    End If
            
            Printer.Font.Bold = False
            'Printer.Font.Size = "7"
       'Printer.Print Tab(0); "------------------------------------------------------------------------------------"
      Printer.Print Tab(0); "-----------------------------------------------------------------"
    
    
    
    
    If id_doc = "0001" Then
        nn = "RUC"
    Else
        nn = "DNI"
    End If
    
    Printer.Print Tab(0); Mid(nn + Space(20), 1, 8) & ":" & rst("id_cliente")
    If KEY_RUBRO = "00025" Then 'colegio
        Printer.Print Tab(0); Mid("ALUMNO " + Space(20), 1, 8) & ":" & Mid(UCase(rst("ncliente")), 1, 28)
    Else
        Printer.Print Tab(0); Mid("CLIENTE " + Space(20), 1, 8) & ":" & Mid(UCase(rst("ncliente")), 1, 30)
    End If
    Printer.Print Tab(9); Mid(UCase(rst("ncliente")), 31, 30)
    
    If IsNull(rst("direccion")) = True Then
        Printer.Print Tab(0); Mid("DIRECCION" + Space(20), 1, 10) & ":" & KEY_DIR_PUBLIC
    Else
        Printer.Print Tab(0); Mid("DIRECCION" + Space(20), 1, 10) & ":" & Mid(in_direccion, 1, 29)
        Printer.Print Tab(11); Mid(in_direccion, 30, 30)
        Printer.Print Tab(11); Mid(in_direccion, 60, 29)
       ' in_ubigeuo1 = Mid(UCase(get_ubigueo_persona(rst("id_cliente"), 0)) & Space(100), 1, 80)
        If KEY_RUC <> "20128836251" Then
        '    Printer.Print Tab(0); Mid("UBIGEO" + Space(20), 1, 11) & ":" & in_ubigeuo1
        End If
    End If
    
    If rst("id_doc") = "0007" Then
        strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & rst("id_comprobante") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
            Printer.Print Tab(0); Mid("DOC.REF" + Space(20), 1, 10) & ":" & rstZ("documento")
            Printer.Print Tab(0); Mid("FECHA DOC" + Space(20), 1, 10) & ":" & Format(rstZ("fecha_emision"), "dd-mm-YYYY")
            Printer.Print Tab(0); Mid("MOTIVO:" + Space(20), 1, 10) & ":"
            Printer.Print Tab(0); UCase(rst("motivo_nota"))
        End If
    End If
    
    If Len(rst("id_copropietario")) > 1 Then
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); "DATOS COPROPIETARIO."
         Printer.Print Tab(0); Mid("DNI " + Space(20), 1, 8) & ":" & Mid(UCase(rst("id_copropietario")), 1, 28)
         Printer.Print Tab(0); Mid("NOMBRE " + Space(20), 1, 8) & ":" & Mid(UCase(get_persona(rst("id_copropietario"))), 1, 28)
         Printer.Print Tab(9); Mid(UCase(get_persona(rst("id_copropietario"))), 29, 28)
    End If
    
'    Printer.Print Tab(1); "------------------------------------------------------------------------------------"
Printer.Print Tab(0); "-----------------------------------------------------------------"
   ' Printer.Print ""
    'Printer.FontBold = True
    in_detraccion = False
Select Case rst("id_tipo_factura")
       
       Case "00001" ' impresion normal
                   ' strCadena = "SELECT M.id_producto,M.detalle,U.abreviatura,mm.descripcion,M.precio,M.total,M.cantidad,mm.descripcion as marca,M.referencia FROM movimiento_venta_detalle M,producto P,unidad U,marca mm WHERE M.id_venta='" & rst("id_venta") & "' AND P.id_marca=mm.id_marca and P.ruc=mm.id_usu AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
                    strCadena = "SELECT * FROM view_detalle_venta where id_venta='" & rst("id_venta") & "'"
                    Call ConfiguraRstT(strCadena)
                    For j = 0 To rstT.RecordCount - 1
                            ptotal = ptotal + rstT("total")
                            codigo = "[" & Format(j + 1, "00") & "]:" & "[" & Mid((rstT("id_producto")) + Space(10), 1, 5) & "]  "
                            descripcion = Mid(rstT("detalle") + Space(80), 1, 40)
                            nunidad = Mid(" " & rstT("abreviatura"), 1, 5)
                            If rstT("id_tipo") = "02" Then
                                in_detraccion = True
                            End If
                            If Trim(rstT("detalle")) < 25 Then
                                descripcion = Mid(rstT("detalle") & " " & rst("marca") + Space(80), 1, 40)
                                'descripcion = Mid(rstT("detalle") & " " + Space(80), 1, 40)
                                Printer.Print Tab(0); codigo & descripcion
                            Else
                                Printer.Print Tab(0); codigo & descripcion
                                If Len(rstT("detalle")) > 40 Then
                                   Printer.Print Tab(10); Mid(rstT("detalle") + Space(80), 41, 40)
                                End If
                                
                                If KEY_RUBRO <> "00025" Then
                                    If rstT("id_producto") <> "00000" Then
                                     '   Printer.Print Tab(0); "MARCA   :" & rstT("marca")
                                    End If
                                End If
                               
                            End If
                            ' If rstT("id_producto") <> "00000" Then
                                Printer.Print Tab(0); Mid(Format(rstT("cantidad"), "#,##0.00") & nunidad + Space(15), 1, 12) & "P.UNIT:" & Mid(Format(rstT("precio"), "#,##0.00") & Space(15), 1, 6) & Space(1) & "TOTAL:" & Mid(Format(rstT("total"), "#,##0.00") + Space(15), 1, 8)
                            'End If
                            'Printer.Print Tab(1); "------------------------------------------------------------------------------------"
                            Printer.Print Tab(0); "-----------------------------------------------------------------"
                            rstT.MoveNext
                            in_filas = in_filas + 1
                    Next j
       Case "00002" ' impresion con serie chasis y motor
                    strCadena = "SELECT C.descripcion as color, U.descripcion as abreviatura,P.id_linea,P.id_producto,M.detalle as nombre_prod,M.cantidad,M.precio,M.total,R.descripcion as marca,M.serie,M.nro_chasis,M.anio_modelo,M.anio_fabricacion,M.nro_dua,M.nro_item,ss.descripcion as modelo,P.id_tipo FROM movimiento_venta_detalle M,producto P,unidad U,marca R,imp_color C,linea_sub ss WHERE  P.id_linea=ss.id_linea and  P.id_sublinea = ss.id_tipo and ss.id_usu = P.ruc and    P.id_color=C.id_color AND   P.id_marca=R.id_marca AND P.ruc=R.id_usu AND    M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstT(strCadena)
                    If rstT.RecordCount > 0 Then
                            in_linea = rstT("id_linea")
                            rstT.MoveFirst
                            For j = 0 To rstT.RecordCount - 1
                                codigo = Mid(rstT("id_producto") + Space(50), 1, 6)
                                cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 15)
                                nunidad = Mid(rstT("abreviatura"), 1, 5)
                                descripcion = Replace(rstT("nombre_prod"), "[", "")
                                descripcion = Mid(Replace(descripcion, "]", "") + Space(80), 1, 40)
                                precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 15)
                                totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 15)
                                Printer.Print Tab(0); codigo & descripcion '& Space(1) & nunidad
                                Printer.Print ""
                                Printer.Print Tab(0); cantidad & precio & Space(2) & totalPar
                                Printer.Print ""
                                If rstT("id_tipo") = "02" Then
                                    in_detraccion = True
                                End If
                                
                                If rstT("id_linea") <> "00047" Then
                                    Printer.Print Tab(2); "MARCA           :" & Space(2) & rstT("marca")
                                    Printer.Print Tab(2); "N°CHASIS        :" & Space(2) & rstT("nro_chasis")
                                    Printer.Print Tab(2); "MOTOR           :" & Space(2) & rstT("serie")
                                    Printer.Print Tab(2); "MODELO          :" & Space(2) & rstT("modelo")
                                    Printer.Print Tab(2); "COLOR           :" & Space(2) & rstT("color")
                                    Printer.Print Tab(2); "AÑO FABRICACION :" & Space(2) & rstT("anio_fabricacion")
                                    Printer.Print Tab(2); "NRO DUA         :" & Space(2) & rstT("nro_dua")
                                    Printer.Print Tab(2); "NRO ITEM        :" & Space(2) & rstT("nro_item")
                                    
                                Else
                                    Printer.Print Tab(2); "MARCA           :" & Space(2) & rstT("marca")
                                    Printer.Print Tab(2); "MOTOR           :" & Space(2) & rstT("serie")
                                    Printer.Print Tab(2); "MODELO          :" & Space(2) & rstT("modelo")
                                    Printer.Print Tab(2); "AÑO FABRICACION :" & Space(2) & rstT("anio_fabricacion")
                                    Printer.Print Tab(2); "NRO DUA         :" & Space(2) & rstT("nro_dua")
                                    Printer.Print Tab(2); "NRO ITEM        :" & Space(2) & rstT("nro_item")
                                End If
                                Call impresion_importacion(rstT("id_producto"))
                                rstT.MoveNext
                            Next j
                            
                            strCadena = "SELECT * FROM linea_mantenimiento WHERE id_linea='" & in_linea & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                            Call ConfiguraRstK(strCadena)
                            If rstK.RecordCount > 0 Then
                                Printer.Print ""
                                Printer.Print ""
                                Printer.Print Tab(0); "PERIODO GARANTIA **********************"
                                Printer.Print Tab(2); "KILOMETROS  :" & Space(2) & rstK("kilometros")
                                Printer.Print Tab(2); "MESES       :" & Space(2) & rstK("dias")
                                Printer.Print Tab(0); "***************************************"
                            End If
                    End If
                    in_filas = in_filas + 11

       Case "00003" ' impresion personalizada
                    in_referencia = ""
                    strCadena = "SELECT * FROM view_detalle_venta where id_venta='" & rst("id_venta") & "'"
                    Call ConfiguraRstT(strCadena)
                    For j = 0 To rstT.RecordCount - 1
                            in_referencia = rstT("referencia")
                            ptotal = ptotal + rstT("total")
                            codigo = "[" & Format(j + 1, "00") & "]::" & "[" & Mid((rstT("id_producto")) + Space(10), 1, 7) & "]  "
                            
                            
                            
                            descripcion = Mid(rstT("detalle") + Space(80), 1, 40)
                            nunidad = Mid("  " & rstT("abreviatura"), 1, 5)
                            If rstT("id_tipo") = "02" Then
                                in_detraccion = True
                            End If
                            If Len(Trim(rstT("detalle"))) < 25 Then
                                descripcion = Mid(rstT("detalle") & " " & rstT("marca") + Space(80), 1, 40)
                                Printer.Print Tab(0); codigo & descripcion
                            Else
                                Printer.Print Tab(0); codigo & descripcion
                                If Len(rstT("detalle")) > 40 Then
                                   Printer.Print Tab(18); Mid(rstT("detalle") + Space(80), 40, 40)
                                End If
                             End If
                            
                            
                            
                            
                            
                            If rstT("cantidad") = 0 Then
                                   ncantidad = " "
                            Else
                                   ncantidad = str(rstT("cantidad"))
                            End If
                            If rstT("precio") = 0 Then
                                   nprecio = " "
                            Else
                                   nprecio = str(Format(rstT("precio"), "#,##0.00"))
                            End If
                            
                            If rstT("total") = 0 Then
                                   nTotal = " "
                            Else
                                   nTotal = str(Format(rstT("total"), "#,##0.00"))
                            End If
                            
                            
                            Printer.Print Tab(0); Mid(Format(rstT("cantidad"), "#,##0.00") & Space(1) & nunidad + Space(15), 1, 15) & "P.UNIT : " & Mid(Format(rstT("precio"), "#,##0.00") & Space(15), 1, 15) & Space(2) & "TOTAL:" & Mid(Format(rstT("total"), "#,##0.00") + Space(15), 1, 15)
                            'Printer.Print Tab(0); Mid(ncantidad + Space(15), 1, 15) & Mid(nprecio + Space(10), 1, 15) & Space(2) & Mid(nTotal + Space(10), 1, 10)
                            Printer.Print Tab(1); "------------------------------------------------------------------------------------"
                            rstT.MoveNext
                            in_filas = in_filas + 1
                    Next j
                    
End Select
    
    
    
                    If Len(in_observacion) > 1 Then
                        Printer.Print Tab(0); in_observacion
                    End If

    
    
    
    
    
    
    
    
    
    
    
    Printer.Print "" 'Tab(10); 'L 9
    If id_doc = "0001" Then
                    Printer.Print Tab(0); "EXONERADO                                      :" & Space(2) & in_simbolo & Space(2) & Format(rst("exonerado"), "#,##0.00")
                    Printer.Print Tab(0); "VALOR VENTA                                    :" & Space(2) & in_simbolo & Space(2) & Format(rst("valor_venta"), "#,##0.00")
                    Printer.Print Tab(0); "IGV                                                  :" & Space(2) & in_simbolo & Space(2) & Format(rst("igv"), "#,##0.00")
    End If
    
    If id_doc = "0054" Then
                n_pago_anterior = 0
                
                strCadena = "SELECT sum(monto_caja),id_forma_pago FROM movimiento_venta_monto m WHERE id_venta='" & in_venta & "' "
                Call ConfiguraRstZ(strCadena)
                
                If rstZ.RecordCount > 0 Then
                    totalletras = UCase(EnLetras(rstZ(0)))
                    If Len(rst("observacion")) > 10 And rst("id_comprobante") > 0 Then
                        strCadena = "SELECT saldo FROM movimiento_venta WHERE id_venta='" & rst("id_comprobante") & "'"
                        Call ConfiguraRstK(strCadena)
                        If rstK.RecordCount > 0 Then
                                Printer.Print Tab(5); ""
                                Printer.Print Tab(0); rst("observacion")
                                Printer.Print Tab(5); ""
                        Printer.Print Tab(0); "MONTO SALDO                                :" & in_simbolo & Format(rstK("saldo"), "#,##0.00")
                        Printer.Print Tab(0); "MONTO AMORTIZADO                            :" & in_simbolo & Format(rstZ(0), "#,##0.00")
                        Printer.Print Tab(0); "SALDO PENDIENTE                           :" & in_simbolo & Format(rstK("saldo") - rstZ(0), "#,##0.00")
                        End If
                    Else
nuevon:
                        Printer.Print Tab(0); "MONTO TOTAL      :" & in_simbolo & Format(rst("total"), "#,##0.00")
                        If rstZ("id_forma_pago") = "12" Then
                        Printer.Print Tab(0); "MONTO CHEQUE     :" & in_simbolo & Format(rstZ(0), "#,##0.00")
                        Else
                        Printer.Print Tab(0); "MONTO EFECTIVO   :" & in_simbolo & Format(rstZ(0), "#,##0.00")
                        End If
                        
                        strCadena = "SELECT sum(monto_caja),id_forma_pago FROM movimiento_venta_monto m WHERE id_venta='" & in_venta & "' and ( id_forma_pago='10'  )  "
                        Call ConfiguraRstL(strCadena)
                        On Error GoTo mas
                        If rstL.RecordCount > 0 And IsNull(rstL(0)) = False Then
                        'Printer.Print Tab(5); Mid("SALDO PENDIENTE" + Space(20), 1, 15) & ":" & "S/." & Format(rst("total") - rstZ(0), "#,##0.00")
                                n_pago_anterior = rstZ(0)
                        Printer.Print Tab(0); "PAGO ANTERIOR                                :" & in_simbolo & Format(rstL(0), "#,##0.00")
                        Printer.Print Tab(0); "SALDO PENDIENTE                              :" & in_simbolo & Format(rst("total") - rstL(0) - rstZ(0), "#,##0.00")
                        Else
mas:
                        Printer.Print Tab(0); "SALDO PENDIENTE   :" & in_simbolo & Format(rst("total") - rstZ(0), "#,##0.00")
                        End If
                        
                    End If
                    
                    Printer.Print Tab(15); ""
                    Printer.Print Tab(0); totalletras & Space(1) & get_moneda(in_moneda)
        
                End If
    Else
        Printer.Print Tab(0); Mid("TOTAL" + Space(50), 1, 50) & ":" & Space(2) & in_simbolo & Space(2) & Format(rst("total"), "#,##0.00")
        totalletras = UCase(EnLetras(rst("total")))
        Printer.Print Tab(0); totalletras & Space(1) & get_moneda(in_moneda)
        'Printer.Print Tab(0); Mid(totalletras, 41, 40) & Space(1) & get_moneda(in_moneda)
    End If
    
    
    Printer.Print ""
    
    
strCadena = "SELECT * FROM movimiento_venta_monto M,forma_pago_detalle F WHERE M.id_forma_pago=F.id_registro AND id_venta='" & rst("id_venta") & "' AND M.ruc='" & KEY_RUC & "' AND F.ruc='" & KEY_RUC & "' "
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Printer.EndDoc
    Exit Sub
Else
    If rstT.RecordCount > 1 Then
        GoTo siguientel
     End If
End If
        tpago = 0
   If id_doc <> "0109" And id_doc <> "0099" Then
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); "FORMA DE PAGO *************************"
        For i = 0 To rstT.RecordCount - 1
            tpago = rstT("monto") + tpago
            If rstT("id_forma_pago") = "12" Then
                Printer.Print Tab(0); Mid(rstT("descripcion") & Space(1) & "[" & rstT("cheque") & "]" + Space(28), 1, 45) & ":" & Space(2) & in_simbolo & Space(1) & Format(rstT("monto"), "#,##0.00")
            Else
                Printer.Print Tab(0); Mid(rstT("descripcion") + Space(50), 1, 20) & ":" & Space(2) & in_simbolo & Space(1) & Format(rstT("monto"), "#,##0.00")
            End If
            
            rstT.MoveNext
        Next i
        Printer.Print Tab(0); Mid("VUELTO" + Space(50), 1, 20) & ":" & Space(2) & in_simbolo & Space(1) & Format(Val(tpago) - rst("total"), "#,##0.00")
  End If
siguientel:
    Printer.Print ""
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    If in_electronico = "si" And rst("id_doc") <> "0109" Then
      ' If firma = "si" Then
    ' Printer.PaintPicture FrmVentas.Pic_firma, 2, 30 + alt_seguro + al_vinculo + Val(FrmVentas.LblCantidad.Caption) * 2, FrmVentas.Pic_firma.Width * 1, _
                       FrmVentas.Pic_firma.Height * 1, 0, 0, FrmVentas.Pic_firma.Width, _
                       FrmVentas.Pic_firma.Height
                       
                       
    Printer.Print Tab(0); "ATENDIDO POR:" + Mid(get_persona(rst("id_vendedor")), 1, 15) + Space(2) + Format(rst("hora"), "HH:mm am/pm")
    Printer.Print Tab(0); "HORA DE IMPRESION:"; Format(Now(), "HH:mm:ss")
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print ""
    If Trim(in_referencia) <> "" And Trim(in_referencia) <> "-" Then
        Printer.Print "--------------------------------------------------"
        Printer.Print Tab(0); "REFERENCIA:" & in_referencia
        Printer.Print "--------------------------------------------------"
    End If
    
    
    If in_detraccion = True And rst("total") >= 750 Then
       Printer.Print Tab(0); "------------------------------------------------------------------------------------"
       Printer.Print Tab(0); "OPERACION SUJETA AL SISTEMA DE OBLIGACIONES"
       Printer.Print Tab(0); "TRIBUTARIAS DEL GOBIERNO CENTRAL"
       Printer.Print Tab(0); "TIPO SERVICIO                      :022"
       Printer.Print Tab(0); "PORCENTAJE DETRACCION       :" & KEY_PORCENTAJE_DETRACCION & "%"
       Printer.Print Tab(0); " CUENTA DETRACCIONES (BN) M/N    :" & "[**" & KEY_CTA_DETRACCION & "**]"
       
       Printer.Print Tab(15); "-----------------------------------------"
       Printer.Print Tab(15); "IMPORTE A DETRAER     :" & Format(rst("total"), "#,##0.00")
       in_monto_detraccion = rst("total") * KEY_PORCENTAJE_DETRACCION / 100
       Printer.Print Tab(15); "DETRACCION                :" & Format(in_monto_detraccion, "#,##0.00")
       Printer.Print Tab(15); "NETO A PAGAR              :" & Format(rst("total") - in_monto_detraccion, "#,##0.00")
       
       Printer.Print Tab(0); "------------------------------------------------------------------------------------"
       Printer.Print ""
    End If
    
    If in_dni = "00000000" Then
            Printer.Font.Bold = True
            Printer.Font.Size = "9"
            Printer.Print Tab(0); "-------------------------------------------"
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); "       AVISO IMPORTANTE !!!"
            Printer.Print Tab(0); " NO SE ACEPTAN DEVOLUCIONES "
            Printer.Print Tab(0); " POR NO ESTAR IDENTIFICADO [DNI/RUC] "
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); "-------------------------------------------"
            Printer.Print Tab(0); "          FIRMA DEL USUARIO  "
            Printer.Font.Bold = False
            Printer.Font.Size = "7"
            Printer.Print ""
        End If
        
        
    
    Printer.Print ""
    Printer.Print Tab(10); "Resumen:" & rst("sunat_hash")
    Printer.Print ""
    Printer.Print Tab(0); "Autorizado mediante resolucion N°: " & KEY_RESOLUCION & "/SUNAT"
    Printer.Print Tab(0); "     Representacion Impresa de la boleta de venta"
    Printer.Print Tab(0); "      Electronica consulte su documento en "
    Printer.Print "" 'Tab(10); 'L 9'
    If KEY_SERVIDOR_KEYFACIL = "si" Then
        Printer.Print Tab(0); "       https://keyfacil.com/consultar"
    Else
        Printer.Print Tab(0); "       http://facturacion.vitekey.com/consultar"
    End If
    
    
    Else
        If rst("id_doc") = "0099" Then
            Printer.Font.Bold = True
            Printer.Font.Size = "9"
            Printer.Print Tab(0); "       AVISO IMPORTANTE !!!"
            Printer.Print Tab(0); " ESTE DOCUMENTO NO ES VALIDO"
            Printer.Print Tab(0); " CANJE POR UNA [BOLETA][FACTURA]"
            Printer.Font.Bold = False
            Printer.Font.Size = "7"
            Printer.Print ""
        End If
        
        If in_dni = "00000000" Then
            Printer.Font.Bold = True
            Printer.Font.Size = "9"
            Printer.Print Tab(0); "-------------------------------------------"
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); "       AVISO IMPORTANTE !!!"
            Printer.Print Tab(0); " NO SE ACEPTAN DEVOLUCIONES "
            Printer.Print Tab(0); " POR NO ESTAR IDENTIFICADO [DNI/RUC] "
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); "-------------------------------------------"
            Printer.Print Tab(0); "          FIRMA DEL USUARIO  "
            Printer.Font.Bold = False
            Printer.Font.Size = "7"
            Printer.Print ""
        End If
        
        
        
        
        
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); "      EXCELENCIA A SU SERVICIO"
    End If
    
    'Call printer_barcode("Hola")
   ' Exit Sub
     'Printer.PaintPicture FrmVentas.Pic_firma, 0, 25 + in_filas, 850, _
                       850, 0, 0, FrmVentas.Pic_firma.Width, _
                       FrmVentas.Pic_firma.Height
                       
                       
                       
    
   ' Printer.Print ""
   ' Printer.Print ""
   ' Printer.Print Tab(0); "TODO RECLAMO DENTRO DE LOS 6 DIAS"
   ' Printer.Print Tab(0); "PRESENTANDO SU COMPROBANTE DE PAGO"
   ' Printer.Print Tab(0); ""
   ' Printer.Print Tab(0); ""
    Printer.Print Tab(0); "ATENDIDO POR:" + Mid(get_persona(rst("id_vendedor")), 1, 15) + Space(2) + Format(rst("hora"), "HH:mm am/pm")
    Printer.Print Tab(0); "HORA DE IMPRESION:"; Format(Now(), "HH:mm:ss")
    Printer.Print Tab(0); " "
    Printer.Print Tab(0); ""
    'Printer.Print Tab(0); ""
    'Printer.Print Tab(0); ""
    'Printer.Print Tab(0); ""
    'Printer.Print Tab(0); ""
    'Printer.Print Tab(0); ""
    'Printer.Print Tab(0); ""
    Printer.EndDoc
    
    Exit Sub
    If get_cuotas(in_venta) = True Then
        Call impresion_cuotas_credito(in_venta)
    End If
    
    
    'Call FrmVentas.Nuevo
    Exit Sub
End Sub

Private Sub impresion_tiketera_electronica(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String, impresiones As Integer
Dim in_detraccion As Boolean

Dim nn As String
Dim nserial As String
Dim in_venta As Double
Dim in_direccion  As String
Dim in_moneda As String
Dim in_simbolo As String
Dim in_referencia As String
Dim in_observacion As String
Dim in_filas As Integer
Dim in_dni As String
in_observacion = ""
  ' Call CargaDefConfigEpsonTM
    
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    If id_doc <> "0109" Then
   '     Printer.Font.name = "control"
   '     Printer.Print "A"
    End If
    'Printer.Font.name = "FontB11"
    'Printer.Font.Size = "10"
    
   
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    '
    strCadena = "SELECT serial,fuente FROM almacen_comprobante WHERE id_doc='" & id_doc & "' and serie='" & serie & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    nserial = rst("serial")
    
    Printer.Font.name = Trim(rst("fuente"))
    Printer.Font.Size = "7"
    
    
    ':::::::::: COMPROBANTE ::::::::::::::::::::::::::::::::::::::::::::::::
    strCadena = "SELECT v.id_venta,v.id_moneda,v.id_doc_fact,v.fecha_fact,v.observacion,v.serie_fact,v.numero_fact,v.id_doc,v.serie,v.numero,v.documento,v.exonerado,v.valor_venta,v.igv,v.id_vendedor,v.hora,v.fecha_emision,p.nombre_completo,v.id_cliente,v.ncliente,v.direccion as ndireccion,p.direccion,v.impresiones,v.id_tipo_factura,v.total,v.observacion,v.saldo,v.id_comprobante,v.id_copropietario,p.id_departamento,p.id_provincia,p.id_distrito,sunat_hash,motivo_nota,v.id_forma_pago,v.cuotas,v.fecha_vencimiento,v.id_alm,v.id_direccion,v.percepcion,v.descuento,v.icbper FROM movimiento_venta v LEFT JOIN persona p ON v.id_cliente=p.dni WHERE  v.id_doc='" & id_doc & "' AND v.serie='" & serie & "' AND v.numero='" & numero & "' AND v.ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    in_observacion = rst("observacion")
    in_venta = rst("id_venta")
    in_moneda = rst("id_moneda")
    in_dni = rst("id_cliente")
    
    If in_moneda = "00001" Then
       in_simbolo = "S/ "
    Else
       in_simbolo = "US$ "
       
    End If
    in_direccion = UCase(rst("ndireccion"))
    
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    
    Printer.FontBold = False
    
 
    
    'Printer.Print "" 'Tab(10); 'L 2
    
    Set FrmVentas.Picture1.Picture = LoadPicture(App.Path & "\Imagenes\logo11.jpg")
    Printer.PaintPicture FrmVentas.Picture1.Picture, 4, 0
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    
    
    'Printer.Print ""
    'Printer.Print "" 'Tab(10); 'L 2
    If rst("impresiones") >= 1 Then
       
            Printer.FontBold = True
            
            Printer.Font.Size = "12"
            Printer.Print Tab(1); "-------------------------------------------------------------"
            Printer.Print Tab(1); "CONTROL ADMINISTRATIVO"
            
            Printer.Print Tab(1); "-------------------------------------------------------------"
            Printer.Font.Bold = False
            Printer.Font.Size = "7"
    End If
    
    
    If Mid(KEY_RUC, 1, 2) = "10" Then
        
        Printer.Print Tab(0); KEY_NOMBRE_COMERCIAL
        'Printer.Print ""
        Printer.Print Tab(0); "DE:" & KEY_EMPRESA
        Printer.Print Tab(0); "-----------------------------------------------------------------"
    Else
        Printer.FontBold = True
        Printer.Font.Size = "10"
        Printer.Print Tab(0); KEY_EMPRESA
        Printer.Font.Bold = False
            Printer.Font.Size = "7"
    End If
    
    
  
    Printer.Print Tab(0); Mid(KEY_DIRECCION, 1, 36)
    Printer.Print Tab(0); Mid(KEY_DIRECCION, 37, 36)
    If KEY_DIRECCION <> KEY_DIRECCION_ALM Then
       Printer.Print Tab(1); "-------------------------------------------------------------"
       Printer.Print Tab(0); "SUC:" & KEY_DIRECCION_ALM
       Printer.Print Tab(1); "-------------------------------------------------------------"
    End If
    If KEY_RUC = "20600133889" Then
        Printer.Print Tab(1); "-------------------------------------------------------------"
       Printer.Print Tab(0); "SUC:" & "AV. AGRICULTURA KM.1 VIÑA DEL MAR."
       Printer.Print Tab(1); "-------------------------------------------------------------"
    End If
    
    If KEY_RUBRO = "00025" Then
        Printer.Print Tab(0); "TELF:074-204985   CELL:#949458350"

    Else
        Printer.Print Tab(0); "TELF:" & KEY_TELEFONO

    End If
    
    Printer.Print Tab(0); "E-MAIL  :" & KEY_EMAIL
    Printer.Print Tab(1); ""
    Printer.Print Tab(0); "RUC   :" & KEY_RUC
    Printer.Print Tab(0); "FECHA EMISION:" & Format(rst("fecha_emision"), "dd-mm-YYYY")
    If rst("id_forma_pago") = "02" And rst("cuotas") < 1 Then
    Printer.Print Tab(0); "FECHA VENCIMIENTO:" & Format(rst("fecha_vencimiento"), "dd-mm-YYYY")
    End If
    
    
    Printer.Print Tab(0); "------------------------------------------------------------------------------------"
 
    
    
   ' Printer.Print Tab(0); "TICKET-" & rst("documento")
   ' Printer.Print Tab(0); "SERIE  :" + nserial
   ' Printer.Print Tab(0); "Nº AUT :" + "0071845087225"
    
    
    
    
    in_electronico = get_electronico(rst("id_doc"), rst("serie"))
    
            
    Printer.Font.Bold = True
    If rst("id_doc") = "0002" Then
        Printer.Print Tab(0); BDBuscarCampo("comprobantes", "doc_abrev", "id_doc", rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
    Else
        If KEY_FACTURACION_ELECTRONICA = "si" And in_electronico = "si" Then
            
            Select Case rst("id_doc")
                   Case "0003"
                        Printer.Print Tab(7); "BOLETA DE VENTA ELECTRONICA"
                   Case "0001"
                        Printer.Print Tab(7); "FACTURA DE VENTA ELECTRONICA"
                   Case "0007"
                        Printer.Print Tab(7); "NOTA DE CREDITO ELECTRONICA"
            End Select
                        Printer.Print Tab(16); rst("serie") & "-" & rst("numero")
            
        Else
            
            If KEY_RUBRO = "00025" Then
                Printer.Print Tab(0); "TICKET " & BDBuscarCampo("comprobantes", "doc_abrev", "id_doc", rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
                Printer.Print Tab(0); "SERIE: FFFF018300"
                Printer.Print Tab(0); "Nro Aut: 0073845113276"
            Else
                'Printer.Print Tab(0); get_comprobante_des(rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
                Printer.Print Tab(0); "TICKET " & get_comprobante_des(rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
                'Printer.Print Tab(0); "SERIE: FFCF214706"
                'Printer.Print Tab(0); "Nro Aut: 0073845123567"
            End If
            
        End If
    End If
            
            Printer.Font.Bold = False
            Printer.Font.Size = "7"
            Printer.Print Tab(0); "------------------------------------------------------------------------------------"
       
    
    
    
    
    If id_doc = "0001" Then
        nn = "RUC"
    Else
        nn = "DNI"
    End If
    
    Printer.Print Tab(0); Mid(nn + Space(20), 1, 8) & ":" & rst("id_cliente")
    If KEY_RUBRO = "00025" Then 'colegio
        Printer.Print Tab(0); Mid("ALUMNO " + Space(20), 1, 8) & ":" & Mid(UCase(rst("ncliente")), 1, 40)
    Else
        Printer.Print Tab(0); Mid("CLIENTE " + Space(20), 1, 8) & ":" & Mid(UCase(rst("ncliente")), 1, 40)
    End If
    
    
    If IsNull(rst("direccion")) = True Then
        Printer.Print Tab(0); Mid("DIRECCION" + Space(20), 1, 10) & ":" & KEY_DIR_PUBLIC
    Else
        Printer.Print Tab(0); Mid("DIRECCION" + Space(20), 1, 10) & ":" & Mid(in_direccion, 1, 40)
        
        
        
        in_ubigeuo1 = Mid(UCase(get_ubigueo_persona(rst("id_cliente"), rst("id_direccion"))) & Space(100), 1, 80)
        If KEY_RUC <> "20128836251" Then  ' vargas
            Printer.Print Tab(0); Mid("UBIGEO" + Space(20), 1, 11) & ":" & in_ubigeuo1
        End If
    End If
    
    If rst("id_doc") = "0007" Then
        strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & rst("id_comprobante") & "'  and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
            Printer.Print Tab(0); Mid("DOC.REF" + Space(20), 1, 10) & ":" & rstZ("documento")
            Printer.Print Tab(0); Mid("FECHA DOC" + Space(20), 1, 10) & ":" & Format(rstZ("fecha_emision"), "dd-mm-YYYY")
            Printer.Print Tab(0); Mid("MOTIVO:" + Space(20), 1, 10) & ":"
            Printer.Print Tab(0); UCase(rst("motivo_nota"))
        End If
    End If
    
    If Len(rst("id_copropietario")) > 1 Then
        Printer.Print Tab(0); ""
        Printer.Print Tab(0); "DATOS COPROPIETARIO."
         Printer.Print Tab(0); Mid("DNI " + Space(20), 1, 8) & ":" & Mid(UCase(rst("id_copropietario")), 1, 28)
         Printer.Print Tab(0); Mid("NOMBRE " + Space(20), 1, 8) & ":" & Mid(UCase(get_persona(rst("id_copropietario"))), 1, 28)
         Printer.Print Tab(9); Mid(UCase(get_persona(rst("id_copropietario"))), 29, 28)
    End If
    
    Printer.Print Tab(1); "------------------------------------------------------------------------------------"
   ' Printer.Print ""
    'Printer.FontBold = True
    in_detraccion = False
   
    
Select Case rst("id_tipo_factura")
       
       Case "00001" ' impresion normal
                   ' strCadena = "SELECT M.id_producto,M.detalle,U.abreviatura,mm.descripcion,M.precio,M.total,M.cantidad,mm.descripcion as marca,M.referencia FROM movimiento_venta_detalle M,producto P,unidad U,marca mm WHERE M.id_venta='" & rst("id_venta") & "' AND P.id_marca=mm.id_marca and P.ruc=mm.id_usu AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
                   If KEY_AGRANEL = "si" Then
                        strCadena = "SELECT * FROM view_detalle_venta_agranel where id_venta='" & rst("id_venta") & "'"
                   Else
                        strCadena = "SELECT * FROM view_detalle_venta where id_venta='" & rst("id_venta") & "'"
                   End If
                    Call ConfiguraRstT(strCadena)
                    For j = 0 To rstT.RecordCount - 1
                            ptotal = ptotal + rstT("total")
                            codigo = "[" & Format(j + 1, "00") & "]::" & "[" & Mid((rstT("id_producto")) + Space(10), 1, 7) & "]  "
                            descripcion = Mid(rstT("detalle") + Space(80), 1, 60)
                            nunidad = Mid("  " & rstT("abreviatura"), 1, 5)
                            If rstT("id_tipo") = "02" Then
                                in_detraccion = True
                            End If
                            If Trim(rstT("detalle")) < 25 Then
                                descripcion = Mid(rstT("detalle") & " " & rst("marca") + Space(80), 1, 40)
                                'descripcion = Mid(rstT("detalle") & " " + Space(80), 1, 40)
                                Printer.Print Tab(0); codigo & descripcion
                            Else
                                Printer.Print Tab(0); codigo & descripcion
                                If Len(rstT("detalle")) > 40 Then
                                   Printer.Print Tab(10); Mid(rstT("detalle") + Space(80), 41, 40)
                                End If
                                
                                If KEY_RUBRO <> "00025" Then
                                    If rstT("id_producto") <> "00000" Then
                                        Printer.Print Tab(0); "MARCA   :" & rstT("marca")
                                    End If
                                End If
                               
                            End If
                            ' If rstT("id_producto") <> "00000" Then
                            Printer.Print Tab(0); Mid(Format(rstT("cantidad"), "#,##0.00") + Space(4) + nunidad + Space(15), 1, 15) & "P.UNIT : " & Mid(Format(rstT("precio"), "#,##0.00") & Space(15), 1, 8) & Space(2) & "TOTAL:" & Mid(Format(rstT("total"), "#,##0.00") + Space(15), 1, 12)
                            'End If
                            Printer.Print Tab(1); "------------------------------------------------------------------------------------"
                       
                            rstT.MoveNext
                            in_filas = in_filas + 1
                    Next j
       Case "00002" ' impresion con serie chasis y motor
                    'strCadena = "SELECT C.descripcion as color, U.descripcion as abreviatura,P.id_linea,P.id_producto,M.detalle as nombre_prod,M.cantidad,M.precio,M.total,R.descripcion as marca,M.serie,M.nro_chasis,M.anio_modelo,M.anio_fabricacion,M.nro_dua,M.nro_item,ss.descripcion as modelo,P.id_tipo FROM movimiento_venta_detalle M,producto P,unidad U,marca R,imp_color C,linea_sub ss WHERE  P.id_linea=ss.id_linea and  P.id_sublinea = ss.id_tipo and ss.id_usu = P.ruc and    P.id_color=C.id_color AND   P.id_marca=R.id_marca AND P.ruc=R.id_usu AND    M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
                    strCadena = "SELECT * FROM view_movimiento_venta_detalle WHERE id_venta='" & rst("id_venta") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstT(strCadena)
                    If rstT.RecordCount > 0 Then
                            in_linea = rstT("id_linea")
                            rstT.MoveFirst
                            For j = 0 To rstT.RecordCount - 1
                                
                                codigo = Mid(rstT("id_producto") + Space(50), 1, 6)
                                cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 15)
                                nunidad = Mid(rstT("abreviatura"), 1, 5)
                                descripcion = Replace(rstT("detalle"), "[", "")
                                descripcion = Mid(Replace(descripcion, "]", "") + Space(80), 1, 80)
                                precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 15)
                                totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 15)
                                Printer.Print Tab(0); codigo & descripcion '& Space(1) & nunidad
                                Printer.Print ""
                                Printer.Print Tab(0); cantidad & precio & Space(2) & totalPar
                                Printer.Print ""
                                If rstT("id_tipo") = "02" Then
                                    in_detraccion = True
                                End If
                                
                                If rstT("id_linea") <> "00047" Then
                                    Printer.Print Tab(2); "MARCA           :" & Space(2) & rstT("marca")
                                    
                                    If rstT("motor") = "no" Then
                                       strCadena = "SELECT * FROM parametros_produccion WHERE habilitado='si' and  codigo='chasis' and ruc='" & KEY_RUC & "' LIMIT 1"
                                       Call ConfiguraRstIN(strCadena)
                                       If rstIN.RecordCount > 0 Then
                                          Printer.Print Tab(2); rstIN("descripcion") & "      :" & Space(2) & rstT("nro_chasis")
                                       End If
                                    End If
                                    
                                    If rstT("serie") <> "" Then
                                       strCadena = "SELECT * FROM parametros_produccion WHERE habilitado='si' and  codigo='motor' and ruc='" & KEY_RUC & "' LIMIT 1"
                                       Call ConfiguraRstIN(strCadena)
                                       If rstIN.RecordCount > 0 Then
                                          Printer.Print Tab(2); rstIN("descripcion") & "      :" & Space(2) & rstT("serie")
                                       End If
                                       
                                    End If
                                    
                                    
                                    Printer.Print Tab(2); "MODELO          :" & Space(2) & rstT("modelo")
                                    Printer.Print Tab(2); "COLOR           :" & Space(2) & rstT("color")
                                    If rstT("anio_fabricacion") <> "" Then
                                    Printer.Print Tab(2); "AÑO FABRICACION :" & Space(2) & rstT("anio_fabricacion")
                                    End If
                                    If rstT("anio_modelo") <> "" Then
                                    Printer.Print Tab(2); "AÑO MODELO      :" & Space(2) & rstT("anio_modelo")
                                    End If
                                    If rstT("nro_dua") <> "" Then
                                    Printer.Print Tab(2); "NRO DUA         :" & Space(2) & rstT("nro_dua")
                                    End If
                                    If rstT("anio_dua") <> "" Then
                                    Printer.Print Tab(2); "AÑO DUA         :" & Space(2) & rstT("anio_dua")
                                    End If
                                    If rstT("nro_item") <> "" Then
                                    Printer.Print Tab(2); "NRO ITEM        :" & Space(2) & rstT("nro_item")
                                    End If
                                Else
                                    Printer.Print Tab(2); "MARCA           :" & Space(2) & rstT("marca")
                                    Printer.Print Tab(2); "MOTOR           :" & Space(2) & rstT("serie")
                                    Printer.Print Tab(2); "MODELO          :" & Space(2) & rstT("modelo")
                                    Printer.Print Tab(2); "AÑO FABRICACION :" & Space(2) & rstT("anio_fabricacion")
                                    Printer.Print Tab(2); "NRO DUA         :" & Space(2) & rstT("nro_dua")
                                    Printer.Print Tab(2); "NRO ITEM        :" & Space(2) & rstT("nro_item")
                                End If
                                Call impresion_importacion(rstT("id_producto"))
                                rstT.MoveNext
                            Next j
                            
                            strCadena = "SELECT * FROM linea_mantenimiento WHERE id_linea='" & in_linea & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                            Call ConfiguraRstK(strCadena)
                            If rstK.RecordCount > 0 Then
                                Printer.Print ""
                                Printer.Print ""
                                Printer.Print Tab(0); "PERIODO GARANTIA **********************"
                                Printer.Print Tab(2); "KILOMETROS  :" & Space(2) & rstK("kilometros")
                                Printer.Print Tab(2); "MESES       :" & Space(2) & rstK("dias")
                                Printer.Print Tab(0); "***************************************"
                            End If
                    End If
                    in_filas = in_filas + 11

       Case "00003" ' impresion personalizada
                    in_referencia = ""
                    strCadena = "SELECT * FROM view_detalle_venta where id_venta='" & rst("id_venta") & "'"
                    Call ConfiguraRstT(strCadena)
                    For j = 0 To rstT.RecordCount - 1
                            in_referencia = rstT("referencia")
                            ptotal = ptotal + rstT("total")
                            codigo = "[" & Format(j + 1, "00") & "]::" & "[" & Mid((rstT("id_producto")) + Space(10), 1, 7) & "]  "
                            
                            
                            
                            descripcion = Mid(rstT("detalle") + Space(80), 1, 40)
                            nunidad = Mid("  " & rstT("abreviatura"), 1, 5)
                            If rstT("id_tipo") = "02" Then
                                in_detraccion = True
                            End If
                            If Len(Trim(rstT("detalle"))) < 25 Then
                                descripcion = Mid(rstT("detalle") & " " & rstT("marca") + Space(80), 1, 40)
                                Printer.Print Tab(0); codigo & descripcion
                            Else
                                Printer.Print Tab(0); codigo & descripcion
                                If Len(rstT("detalle")) > 40 Then
                                   Printer.Print Tab(18); Mid(rstT("detalle") + Space(80), 40, 40)
                                End If
                             End If
                            
                            
                            
                            
                            
                            If rstT("cantidad") = 0 Then
                                   ncantidad = " "
                            Else
                                   ncantidad = str(rstT("cantidad"))
                            End If
                            If rstT("precio") = 0 Then
                                   nprecio = " "
                            Else
                                   nprecio = str(Format(rstT("precio"), "#,##0.00"))
                            End If
                            
                            If rstT("total") = 0 Then
                                   nTotal = " "
                            Else
                                   nTotal = str(Format(rstT("total"), "#,##0.00"))
                            End If
                            
                            
                            Printer.Print Tab(0); Mid(Format(rstT("cantidad"), "#,##0.00") & Space(1) & nunidad + Space(15), 1, 15) & "P.UNIT : " & Mid(Format(rstT("precio"), "#,##0.00") & Space(15), 1, 15) & Space(2) & "TOTAL:" & Mid(Format(rstT("total"), "#,##0.00") + Space(15), 1, 15)
                            'Printer.Print Tab(0); Mid(ncantidad + Space(15), 1, 15) & Mid(nprecio + Space(10), 1, 15) & Space(2) & Mid(nTotal + Space(10), 1, 10)
                            Printer.Print Tab(1); "------------------------------------------------------------------------------------"
                            rstT.MoveNext
                            in_filas = in_filas + 1
                    Next j
                    
End Select
    
    
    
                    If Len(in_observacion) > 1 Then
                        Printer.Print Tab(0); in_observacion
                    End If

    
    
    
    
    
    
    
    
    
    
    
    Printer.Print "" 'Tab(10); 'L 9
    If id_doc = "0001" Then
                    If KEY_RUC = "20525999115" Then
                        Printer.Print Tab(0); "DESCUENTO     :" & Space(2) & in_simbolo & Space(2) & Format(rst("descuento"), "#,##0.00")
                        Printer.Print Tab(0); "EXONERADO     :" & Space(2) & in_simbolo & Space(2) & Format(rst("exonerado"), "#,##0.00")
                        Printer.Print Tab(0); "GRAVADO       :" & Space(2) & in_simbolo & Space(2) & Format(rst("valor_venta"), "#,##0.00")
                        Printer.Print Tab(0); "IGV [" & KEY_IGV & "% ]   " & Space(2) & in_simbolo & Space(2) & Format(rst("igv"), "#,##0.00")
                        
                        If KEY_IMPUESTO_BOLSAS = "si" Then
                        Printer.Print Tab(0); "ICBPER        :" & Space(2) & in_simbolo & Space(2) & Format(rst("icbper"), "#,##0.00")
                        End If
                        
                        Printer.Print Tab(0); "GRAVADO       :" & Space(2) & in_simbolo & Space(2) & Format(rst("valor_venta"), "#,##0.00")
                        Printer.Print Tab(0); "TOTAL         :" & Space(2) & in_simbolo & Space(2) & Format(rst("total"), "#,##0.00")
                    Else
                        Printer.Print Tab(0); "DESCUENTO                                 :" & Space(2) & in_simbolo & Space(2) & Format(rst("descuento"), "#,##0.00")
                        
                        If KEY_CON_IGV = "si" Then
                        Printer.Print Tab(0); "GRAVADO                                  :" & Space(2) & in_simbolo & Space(2) & Format(rst("valor_venta"), "#,##0.00")
                        Else
                        Printer.Print Tab(0); "EXONERADO                                :" & Space(2) & in_simbolo & Space(2) & Format(rst("exonerado"), "#,##0.00")
                        End If
                        Printer.Print Tab(0); "IGV                                               :" & Space(2) & in_simbolo & Space(2) & Format(rst("igv"), "#,##0.00")
                        If KEY_IMPUESTO_BOLSAS = "si" Then
                        Printer.Print Tab(0); "ICBPER                                        :" & Space(2) & in_simbolo & Space(2) & Format(rst("icbper"), "#,##0.00")
                        End If
                        Printer.Print Tab(0); "TOTAL                                          :" & Space(2) & in_simbolo & Space(2) & Format(rst("total"), "#,##0.00")
                    End If
    End If
    
    If id_doc = "0003" Then
                    If KEY_RUC = "20525999115" Then
                        Printer.Print Tab(0); "DESCUENTO     :" & Space(2) & in_simbolo & Space(2) & Format(rst("descuento"), "#,##0.00")
                        Printer.Print Tab(0); "EXONERADO     :" & Space(2) & in_simbolo & Space(2) & Format(rst("exonerado"), "#,##0.00")
                        Printer.Print Tab(0); "GRAVADO       :" & Space(2) & in_simbolo & Space(2) & Format(rst("valor_venta"), "#,##0.00")
                        Printer.Print Tab(0); "IGV           :" & Space(2) & in_simbolo & Space(2) & Format(rst("igv"), "#,##0.00")
                        If KEY_IMPUESTO_BOLSAS = "si" Then
                        Printer.Print Tab(0); "ICBPER        :" & Space(2) & in_simbolo & Space(2) & Format(rst("icbper"), "#,##0.00")
                        End If
                        Printer.Print Tab(0); "TOTAL         :" & Space(2) & in_simbolo & Space(2) & Format(rst("total"), "#,##0.00")
                    Else
                        Printer.Print Tab(0); "DESCUENTO                                 :" & Space(2) & in_simbolo & Space(2) & Format(rst("descuento"), "#,##0.00")
                        If KEY_CON_IGV = "si" Then
                        Printer.Print Tab(0); "GRAVADO                                    :" & Space(2) & in_simbolo & Space(2) & Format(rst("valor_venta"), "#,##0.00")
                        Else
                        Printer.Print Tab(0); "EXONERADO                                 :" & Space(2) & in_simbolo & Space(2) & Format(rst("exonerado"), "#,##0.00")
                        End If
                        
                        
                        Printer.Print Tab(0); "IGV                                              :" & Space(2) & in_simbolo & Space(2) & Format(rst("igv"), "#,##0.00")
                        If KEY_IMPUESTO_BOLSAS = "si" Then
                        Printer.Print Tab(0); "ICBPER                                         :" & Space(2) & in_simbolo & Space(2) & Format(rst("icbper"), "#,##0.00")
                        End If
                        Printer.Print Tab(0); "TOTAL                                          :" & Space(2) & in_simbolo & Space(2) & Format(rst("total"), "#,##0.00")
                    End If
    End If
    
    If id_doc = "0007" Then
                    If KEY_RUC = "20525999115" Then
                        Printer.Print Tab(0); "DESCUENTO     :" & Space(2) & in_simbolo & Space(2) & Format(rst("descuento"), "#,##0.00")
                        Printer.Print Tab(0); "EXONERADO     :" & Space(2) & in_simbolo & Space(2) & Format(rst("exonerado"), "#,##0.00")
                        Printer.Print Tab(0); "GRAVADO       :" & Space(2) & in_simbolo & Space(2) & Format(rst("valor_venta"), "#,##0.00")
                        Printer.Print Tab(0); "IGV           :" & Space(2) & in_simbolo & Space(2) & Format(rst("igv"), "#,##0.00")
                        If KEY_IMPUESTO_BOLSAS = "si" Then
                        Printer.Print Tab(0); "ICBPER        :" & Space(2) & in_simbolo & Space(2) & Format(rst("icbper"), "#,##0.00")
                        End If
                        Printer.Print Tab(0); "TOTAL         :" & Space(2) & in_simbolo & Space(2) & Format(rst("total"), "#,##0.00")
                    Else
                        Printer.Print Tab(0); "DESCUENTO                                 :" & Space(2) & in_simbolo & Space(2) & Format(rst("descuento"), "#,##0.00")
                        If KEY_CON_IGV = "si" Then
                        Printer.Print Tab(0); "GRAVADO                                  :" & Space(2) & in_simbolo & Space(2) & Format(rst("valor_venta"), "#,##0.00")
                        Else
                        Printer.Print Tab(0); "EXONERADO                                 :" & Space(2) & in_simbolo & Space(2) & Format(rst("exonerado"), "#,##0.00")
                        End If
                        
                        
                        Printer.Print Tab(0); "IGV                                              :" & Space(2) & in_simbolo & Space(2) & Format(rst("igv"), "#,##0.00")
                        If KEY_IMPUESTO_BOLSAS = "si" Then
                        Printer.Print Tab(0); "ICBPER                                         :" & Space(2) & in_simbolo & Space(2) & Format(rst("icbper"), "#,##0.00")
                        End If
                        Printer.Print Tab(0); "TOTAL                                          :" & Space(2) & in_simbolo & Space(2) & Format(rst("total"), "#,##0.00")
                    End If
    End If
    
    If id_doc = "0054" Then
                n_pago_anterior = 0
                
                strCadena = "SELECT sum(monto_caja),id_forma_pago FROM movimiento_venta_monto m WHERE id_venta='" & in_venta & "' "
                Call ConfiguraRstZ(strCadena)
                
                If rstZ.RecordCount > 0 Then
                    totalletras = UCase(EnLetras(rstZ(0)))
                    If Len(rst("observacion")) > 10 And rst("id_comprobante") > 0 Then
                        strCadena = "SELECT (total-function_pago_factura('" & rst("id_comprobante") & "','" & KEY_FECHA & "',id_moneda,ruc)) FROM movimiento_venta WHERE id_venta='" & rst("id_comprobante") & "'"
                        Call ConfiguraRstK(strCadena)
                        If rstK.RecordCount > 0 Then
                                Printer.Print Tab(5); ""
                                Printer.Print Tab(0); rst("observacion")
                                Printer.Print Tab(5); ""
                        Printer.Print Tab(0); "MONTO TOTAL                                :" & in_simbolo & Format(rstK(0) + rstZ(0), "#,##0.00")
                        Printer.Print Tab(0); "MONTO AMORTIZADO                            :" & in_simbolo & Format(rstZ(0), "#,##0.00")
                        Printer.Print Tab(0); "SALDO PENDIENTE                           :" & in_simbolo & Format(rstK(0), "#,##0.00")
                        End If
                    Else
nuevon:
                        Printer.Print Tab(0); "MONTO TOTAL                                    :" & in_simbolo & Format(rst("total"), "#,##0.00")
                        If rstZ("id_forma_pago") = "12" Then
                        Printer.Print Tab(0); "MONTO CHEQUE                                  :" & in_simbolo & Format(rstZ(0), "#,##0.00")
                        Else
                        Printer.Print Tab(0); "MONTO EFECTIVO                                :" & in_simbolo & Format(rstZ(0), "#,##0.00")
                        End If
                        
                        strCadena = "SELECT sum(monto_caja),id_forma_pago FROM movimiento_venta_monto m WHERE id_venta='" & in_venta & "' and ( id_forma_pago='10'  )  "
                        Call ConfiguraRstL(strCadena)
                        On Error GoTo mas
                        If rstL.RecordCount > 0 And IsNull(rstL(0)) = False Then
                        'Printer.Print Tab(5); Mid("SALDO PENDIENTE" + Space(20), 1, 15) & ":" & "S/." & Format(rst("total") - rstZ(0), "#,##0.00")
                                n_pago_anterior = rstZ(0)
                        Printer.Print Tab(0); "PAGO ANTERIOR                                :" & in_simbolo & Format(rstL(0), "#,##0.00")
                        Printer.Print Tab(0); "SALDO PENDIENTE                              :" & in_simbolo & Format(rst("total") - rstL(0) - rstZ(0), "#,##0.00")
                        Else
mas:
                        Printer.Print Tab(0); "SALDO PENDIENTE                              :" & in_simbolo & Format(rst("total") - rstZ(0), "#,##0.00")
                        End If
                        
                    End If
                    
                    Printer.Print Tab(15); ""
                    Printer.Print Tab(0); totalletras & Space(1) & get_moneda(in_moneda)
        
                End If
    Else
        Printer.Print Tab(0); "------------------------------------------------------------------------------------"
        'Printer.Print Tab(0); Mid("TOTAL" + Space(50), 1, 50) & ":" & Space(2) & in_simbolo & Space(2) & Format(rst("total"), "#,##0.00")
        totalletras = UCase(EnLetras(rst("total")))
        Printer.Print Tab(0); totalletras & Space(1) & get_moneda(in_moneda)
        
         If KEY_RUC = "20561358550" Then
                 If rst("percepcion") > 0 Then
                 Printer.Print Tab(0); Mid("PERCEPCION" + Space(50), 1, 50) & ":" & Space(2) & in_simbolo & Space(2) & Format(rst("percepcion"), "#,##0.00")
                 End If
            End If
        'Printer.Print Tab(0); Mid(totalletras, 41, 40) & Space(1) & get_moneda(in_moneda)
    End If
    
    
    Printer.Print ""
    
If KEY_RUC = "20566449383" Then
    If rst("id_doc") <> "0054" Then
    
        GoTo siguientel
    End If
End If
strCadena = "SELECT * FROM movimiento_venta_monto M,forma_pago_detalle F WHERE M.id_forma_pago=F.id_registro AND id_venta='" & rst("id_venta") & "' AND M.ruc='" & KEY_RUC & "' AND F.ruc='" & KEY_RUC & "' "
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Printer.EndDoc
    Exit Sub
Else
    'If rstT.RecordCount > 1 Then
    '    GoTo siguientel
    ' End If
End If
        tpago = 0
   If id_doc <> "0109" And id_doc <> "0099" Then
        Printer.Print Tab(0); ""
        Printer.FontBold = True
        
            Printer.Font.Size = "9"
        Printer.Print Tab(0); "FORMA DE PAGO ************************"
        For i = 0 To rstT.RecordCount - 1
            tpago = rstT("monto") + tpago
            
            If rstT("id_forma_pago") = "12" Then
                Printer.Print Tab(0); Mid(rstT("descripcion") & Space(1) & "[" & rstT("cheque") & "]" + Space(28), 1, 25) & ":" & Space(2) & in_simbolo & Space(1) & Format(rstT("monto"), "#,##0.00")
            Else
                Printer.Print Tab(0); Mid(rstT("descripcion") + Space(50), 1, 25) & ":" & Space(2) & in_simbolo & Space(1) & Format(rstT("monto"), "#,##0.00")
            End If
            
            rstT.MoveNext
        Next i
        Printer.Print Tab(0); Mid("VUELTO" + Space(50), 1, 25) & ":" & Space(4) & in_simbolo & Space(1) & Format(Val(tpago) - rst("total"), "#,##0.00")
  End If


Printer.FontBold = False
Printer.Font.Size = "7"
siguientel:
Printer.Print ""
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    If in_electronico = "si" And rst("id_doc") <> "0109" Then
      ' If firma = "si" Then
    ' Printer.PaintPicture FrmVentas.Pic_firma, 2, 30 + alt_seguro + al_vinculo + Val(FrmVentas.LblCantidad.Caption) * 2, FrmVentas.Pic_firma.Width * 1, _
                       FrmVentas.Pic_firma.Height * 1, 0, 0, FrmVentas.Pic_firma.Width, _
                       FrmVentas.Pic_firma.Height
                       
                       
    Printer.Print Tab(0); "ATENDIDO POR:" + Mid(get_persona(rst("id_vendedor")), 1, 15) + Space(2) + Format(rst("hora"), "HH:mm am/pm")
    Printer.Print Tab(0); "HORA DE IMPRESION:"; Format(Now(), "HH:mm:ss")
    Printer.Print "" 'Tab(10); 'L 9
    
    If Trim(in_referencia) <> "" And Trim(in_referencia) <> "-" Then
        Printer.Print "--------------------------------------------------"
        Printer.Print Tab(0); "REFERENCIA:" & in_referencia
        Printer.Print "--------------------------------------------------"
    End If
    
    
    If in_detraccion = True And rst("total") >= 750 Then
       Printer.Print Tab(0); "------------------------------------------------------------------------------------"
       Printer.Print Tab(0); "OPERACION SUJETA AL SISTEMA DE OBLIGACIONES"
       Printer.Print Tab(0); "TRIBUTARIAS DEL GOBIERNO CENTRAL"
       Printer.Print Tab(0); "TIPO SERVICIO                      :022"
       Printer.Print Tab(0); "PORCENTAJE DETRACCION       :" & KEY_PORCENTAJE_DETRACCION & "%"
       Printer.Print Tab(0); " CUENTA DETRACCIONES (BN) M/N    :" & "[**" & KEY_CTA_DETRACCION & "**]"
       
       Printer.Print Tab(15); "-----------------------------------------"
       Printer.Print Tab(15); "IMPORTE A DETRAER     :" & Format(rst("total"), "#,##0.00")
       in_monto_detraccion = rst("total") * KEY_PORCENTAJE_DETRACCION / 100
       Printer.Print Tab(15); "DETRACCION                :" & Format(in_monto_detraccion, "#,##0.00")
       Printer.Print Tab(15); "NETO A PAGAR              :" & Format(rst("total") - in_monto_detraccion, "#,##0.00")
       
       Printer.Print Tab(0); "------------------------------------------------------------------------------------"
       Printer.Print ""
    End If
    
    If in_dni = "00000000" Then
            Printer.Font.Bold = True
            Printer.Font.Size = "9"
            Printer.Print Tab(0); "-------------------------------------------"
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); "       AVISO IMPORTANTE !!!"
            Printer.Print Tab(0); " NO SE ACEPTAN DEVOLUCIONES "
            Printer.Print Tab(0); " POR NO ESTAR IDENTIFICADO [DNI/RUC] "
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); "-------------------------------------------"
            Printer.Print Tab(0); "          FIRMA DEL USUARIO  "
            Printer.Font.Bold = False
            Printer.Font.Size = "7"
            Printer.Print ""
        End If
        
        
  
    Printer.Print ""
    Printer.Print Tab(10); "Resumen:" & rst("sunat_hash")
    Printer.Print ""
    Printer.Print Tab(0); "Autorizado mediante resolucion N°: " & KEY_RESOLUCION & "/SUNAT"
    Printer.Print Tab(0); "     Representacion Impresa de la boleta de venta"
    Printer.Print Tab(0); "      Electronica consulte su documento en "
    Printer.Print "" 'Tab(10); 'L 9'
    If KEY_SERVIDOR_KEYFACIL = "si" Then
        Printer.Print Tab(0); "       https://keyfacil.com/consultar"
    Else
        Printer.Print Tab(0); "       http://facturacion.vitekey.com/consultar"
    End If
    
    
    
    Else
        If rst("id_doc") = "0099" Then
            Printer.Font.Bold = True
            Printer.Font.Size = "9"
            Printer.Print Tab(0); "       AVISO IMPORTANTE !!!"
            Printer.Print Tab(0); " ESTE DOCUMENTO NO ES VALIDO"
            Printer.Print Tab(0); " CANJE POR UNA [BOLETA][FACTURA]"
            Printer.Font.Bold = False
            Printer.Font.Size = "7"
            Printer.Print ""
        End If
        
        If in_dni = "00000000" Then
            Printer.Font.Bold = True
            Printer.Font.Size = "9"
            Printer.Print Tab(0); "-------------------------------------------"
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); "       AVISO IMPORTANTE !!!"
            Printer.Print Tab(0); " NO SE ACEPTAN DEVOLUCIONES "
            Printer.Print Tab(0); " POR NO ESTAR IDENTIFICADO [DNI/RUC] "
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); "-------------------------------------------"
            Printer.Print Tab(0); "          FIRMA DEL USUARIO  "
            Printer.Font.Bold = False
            Printer.Font.Size = "7"
            Printer.Print ""
        End If
        
        
        
        
      '  If KEY_RUC = "20566449383" And rst("id_doc") = "0054" Then
      '      Printer.Print Tab(0); ""
      ' '      Printer.Print Tab(0); ""
      '      Printer.Print Tab(0); "---------------------------------------     ---------------------------------------"
      '      Printer.Print Tab(0); "          FIRMA DEL CLIENTE                              COLABORADOR "
      '      Printer.Print Tab(0); ""
      '      Printer.Print Tab(0); "DNI :     ___________________"
      '
      '      Printer.Print Tab(0); ""
      '      Printer.Print Tab(0); ""
      '      Printer.Print Tab(0); ""
            
      '  End If
        
       
        
        
        
        
        Printer.Print Tab(0); "           EXCELENCIA A SU SERVICIO"
    End If
    
    
     If KEY_RUC = "20566449383" And rst("id_forma_pago") = "02" And rst("id_doc") = "0054" Then
            Printer.Print Tab(0); ""
             Printer.Print Tab(0); ""
            Printer.Print Tab(0); "---------------------------------------     ---------------------------------------"
            Printer.Print Tab(0); "          FIRMA DEL CLIENTE                              COLABORADOR "
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); "DNI :     ___________________"
            
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); ""
            Printer.Print Tab(0); ""
            
        End If
    'Call printer_barcode("Hola")
   ' Exit Sub
     'Printer.PaintPicture FrmVentas.Pic_firma, 0, 25 + in_filas, 850, _
                       850, 0, 0, FrmVentas.Pic_firma.Width, _
                       FrmVentas.Pic_firma.Height
                       
                       
    '----
    
   ' Printer.Print ""
    Printer.Print Tab(0); ""
    pie_pagina = get_pie_pagina
    
    Printer.Print Tab(0); "      " + Mid(pie_pagina, 1, 40)
    Printer.Print Tab(0); "      " + Mid(pie_pagina, 41, 40)
    Printer.Print Tab(0); "      " + Mid(pie_pagina, 81, 40)
    
    
   ' Printer.Print Tab(0); ""
   ' Printer.Print Tab(0); ""
   ' Printer.Print Tab(0); "ATENDIDO POR:" + Mid(get_persona(rst("id_vendedor")), 1, 15) + Space(2) + Format(rst("hora"), "HH:mm am/pm")
   ' Printer.Print Tab(0); "HORA DE IMPRESION:"; Format(Now(), "HH:mm:ss")
    'Printer.Print Tab(0); ""
    
    
    
    
    
    Printer.Print Tab(0); ""
    
   If rst("id_doc") = "0054" Then
    Printer.Print "" '
    Printer.Print Tab(0); "ESTE COMPROBANTE NO TIENE VALIDEZ FISCAL"
    Printer.Print Tab(0); "   DEBE SER CANJEADO POR UNA BOLETA / FACTURA"
     Printer.Print "" '
  End If

  If KEY_APLICA_IGV = "no" Then
        Printer.Print Tab(0); " BIENES TRANSFERIDOS EN LA AMAZONIA"
        Printer.Print Tab(0); "  PARA SER CONSUMIDOS EN LA MISMA"
        Printer.Print "" '
    End If
    Printer.EndDoc
    
    
    If get_cuotas(in_venta) = True Then
        Call impresion_cuotas_credito(in_venta)
    End If
    
    
    'Call FrmVentas.Nuevo
    Exit Sub
End Sub

Private Function get_pie_pagina() As String

strCadena = "SELECT pie_pagina FROM almacen WHERE id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_pie_pagina = rstL("pie_pagina")
Else
   get_pie_pagina = ""
End If

End Function


Private Sub impresion_importacion(ByVal in_producto As String)
strCadena = "SELECT * FROM producto_importaciones WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
                                    Printer.Print Tab(2); "--------------------------------------------------"
                                    Printer.Print Tab(2); "CILINDRAJE       :" & Space(2) & UCase(rstP("cilindraje"))
                                    Printer.Print Tab(2); "N°CILINDROS      :" & Space(2) & rstP("num_cilindros")
                                    Printer.Print Tab(2); "POTENCIA MOTOR   :" & Space(2) & rstP("potencia_motor")
                                    Printer.Print Tab(2); "TIPO COMBUSTIBLE :" & Space(2) & UCase(rstP("tipo_gasolina"))
 End If

End Sub


Public Sub impresion_despacho(ByVal id_venta As String)
Dim ptotal As Double, direccion_destino As String, impresiones As Integer
Dim nn As String
Dim nserial As String
Dim in_venta As Double
Dim in_direccion  As String
    ' Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "Tahoma"
    Printer.Font.Size = "8"
   
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
    Printer.Print Tab(0); KEY_EMPRESA
    Printer.Print Tab(0); Mid(KEY_DIRECCION, 1, 36)
    Printer.Print Tab(0); Mid(KEY_DIRECCION, 37, 36)
    Printer.Print Tab(0); ""
    
    Printer.Print Tab(1); "-------------------------------------------------------------"
    strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & Val(id_venta) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Printer.Print Tab(0); "ORDEN DESPACHO :" & Format(id_venta, "00000000")
       Printer.Print Tab(1); "-------------------------------------------------------------"
       Printer.Print Tab(0); "NOMBRE CLIENTE :" & rst("ncliente")
       Printer.Print Tab(1); "-------------------------------------------------------------"
       Printer.Print Tab(0); "FECHA VENTA    :" & Format(rst("fecha_emision"), "dd-mm-YYYY")
       Printer.Print Tab(1); "-------------------------------------------------------------"
       Printer.Print Tab(0); "COMPROBANTE    :" & rst("documento")
       Printer.Print Tab(1); "-------------------------------------------------------------"
       Printer.Print Tab(0); "HORA ENTREGA   :" & Format(KEY_FECHA, "dd-mm-YYYY")
       Printer.Print Tab(0); ""
       Printer.Print Tab(0); ""
       Printer.Print Tab(0); "  ----  DETALLE PEDIDO -----  "
       Printer.Print Tab(1); "-------------------------------------------------------------"
       strCadena = "SELECT M.id_producto,M.detalle,U.abreviatura,mm.descripcion,M.precio,M.total,M.cantidad,mm.descripcion as marca FROM movimiento_venta_detalle M,producto P,unidad U,marca mm WHERE M.id_venta='" & Val(id_venta) & "' AND P.id_marca=mm.id_marca and P.ruc=mm.id_usu AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstK(strCadena)
                    For j = 0 To rstK.RecordCount - 1
                            
                            codigo = Mid((rstK("id_producto")) + Space(10), 1, 7)
                            descripcion = Mid(rstK("detalle") + Space(80), 1, 40)
                            nunidad = Mid("  " & rstK("abreviatura"), 1, 5)
                            If Trim(rstK("detalle")) < 25 Then
                                descripcion = Mid(rstK("detalle") & " " & rstK("marca") + Space(80), 1, 40)
                                Printer.Print Tab(0); codigo & descripcion
                            Else
                                Printer.Print Tab(0); "[ " & codigo & " ] " & descripcion
                                Printer.Print Tab(0); "MARCA   :" & rstK("marca")
                            End If
                            
                            Printer.Print Tab(0); Mid(Format(rstK("cantidad"), "#,##0.00") & nunidad + Space(15), 1, 15) & "P.UNIT : " & Mid(Format(rstK("precio"), "#,##0.00") & Space(15), 1, 15) & Space(2) & Mid(Format(rstK("total"), "#,##0.00") + Space(15), 1, 15)
                            Printer.Print Tab(1); "-------------------------------------------------------------"
                            rstK.MoveNext
                    Next j
       Printer.Print Tab(0); ""
       
       
       
       
       
    End If
    
    
    
    
    
    
    
    Printer.Print Tab(0); "DESPACHADO POR:" + Mid(KEY_VENDEDOR, 1, 15)
    Printer.Print Tab(0); "HORA DE IMPRESION:"; Format(Now(), "HH:mm:ss")
    Printer.Print "" 'Tab(10); 'L 9
    
                       
    
    

    
    
    Printer.EndDoc
    'Call FrmVentas.Nuevo
    Exit Sub
End Sub

Public Sub impresion_barras(ByVal id_producto As String, ByVal numero As Integer)
Dim ptotal As Double, direccion_destino As String, impresiones As Integer
Dim nn As String
   'Call CargaDefConfigEpsonTM
    
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    
    'Printer.Font.name = "FontB11"
    'Printer.Font.Size = "10"
   ' m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
    strCadena = "SELECT * FROM producto_barras B,producto P WHERE  B.id_producto=P.id_producto AND B.id_producto='" & id_producto & "' AND B.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        For i = 1 To numero
            Printer.Font.name = "Arial"
            Printer.Font.Size = "6"
            'Printer.Print Tab(5); "0123456789-0123456789-0123456789-0123456789-0123456789-0123456789-0123456789-01234567"
           ' Printer.Print ""
            'Printer.CurrentY = Printer.CurrentY + 0.5
            Printer.Print ""
             Printer.CurrentY = Printer.CurrentY + 0.5
            Printer.Print Tab(5); Mid(rst("nombre_prod") + Space(10), 1, 23) + Space(7) + Mid(rst("nombre_prod") + Space(10), 1, 23) + Space(6) + Mid(rst("nombre_prod") + Space(10), 1, 23)
            'Printer.Print ""
            Printer.Print Tab(12); Mid("PRECIO  :" + "S/." + Space(2) + Format(rst("precio_venta"), "#,##0.00") + Space(20), 1, 35) + Mid("PRECIO  :" + "S/." + Space(2) + Format(rst("precio_venta"), "#,##0.00") + Space(20), 1, 40) + Mid("PRECIO  :" + "S/." + Space(2) + Format(rst("precio_venta"), "#,##0.00"), 1, 27)
            'Printer.Print ""
            Printer.Font.name = "3 of 9 Barcode"
            Printer.Font.Size = "15"
            Printer.Print Tab(2); Chr$(1) & rst("cod_barra") & Chr$(2) + Space(3) + Chr$(1) & rst("cod_barra") & Chr$(2) + Space(2) + Chr$(1) & rst("cod_barra") & Chr$(2)
            Printer.Print ""
            Printer.Print ""
           
            
        Next i
    End If
    
    Printer.EndDoc
    
    Exit Sub
End Sub

Private Sub impresion_formato_1_rboingreso(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser

    Printer.ScaleWidth = 11.8
    Printer.ScaleHeight = 15.3
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM mis_cuentas_det WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
       
    Printer.Print ""
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(50); rst("documento")
    Printer.Print Tab(10); Mid("FECHA OPERACION" + Space(30), 1, 20) & ":" & formato_item(Day(rst("fecha")), 2) & Space(3) & formato_item(Month(rst("fecha")), 2) + Space(3) + str(Year(rst("fecha")))
    Printer.Print Tab(10); Mid("RUC/DNI" + Space(30), 1, 20) & ":" & rst("id_persona")
    Printer.Print Tab(10); Mid("NOMBRE RAZON" + Space(30), 1, 20) & ":" & UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_persona")))
    Printer.Print Tab(10); Mid("DIRECCION" + Space(30), 1, 20) & ":" & Mid(UCase(BDBuscarCampo("persona", "direccion", "dni", rst("id_persona"))) + Space(80), 1, 60)
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print Tab(1); "============================================================================"
    Printer.Print Tab(1); Trim(rst("glosa"))
    Printer.Print Tab(1); "============================================================================"
          
    Printer.EndDoc
    Exit Sub
End Sub
Public Sub impresion_formato_1_rboingreso_tiket(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
    Dim ptotal As Double, direccion_destino As String
   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser

    Printer.Font.name = "FontB11"
    Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM mis_cuentas_det WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
       
    Printer.Print Tab(1); KEY_EMPRESA
    Printer.Print Tab(1); KEY_DIRECCION
    Printer.Print Tab(1); "TELF:529174-#988933943 - #979409952"
    Printer.Print Tab(1); "RUC   :" & KEY_RUC
    Printer.Print Tab(1); "FECHA :" & str(rst("fecha"))
    Printer.Print Tab(1); "-----------------------------------------"
    Printer.Print Tab(1); rst("documento")
    Printer.Print Tab(1); "-----------------------------------------"
    Printer.Print Tab(1); Mid("FECHA OPERACION" + Space(30), 1, 20) & ":" & formato_item(Day(rst("fecha")), 2) & Space(3) & formato_item(Month(rst("fecha")), 2) + Space(3) + str(Year(rst("fecha")))
    Printer.Print Tab(1); Mid("RUC/DNI" + Space(30), 1, 20) & ":" & rst("id_persona")
    Printer.Print Tab(1); Mid("NOMBRE RAZON" + Space(30), 1, 20) & ":" & UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_persona")))
    Printer.Print Tab(1); Mid("DIRECCION" + Space(30), 1, 20) & ":" & Mid(UCase(BDBuscarCampo("persona", "direccion", "dni", rst("id_persona"))) + Space(80), 1, 60)
    Printer.Print Tab(1); "-----------------------------------------"
    Printer.Print Tab(1); "MONTO :" & Format(rst("monto"), "#,##0.00")
    Printer.Print Tab(1); UCase(EnLetras(rst("monto")))
    Printer.Print Tab(1); "-----------------------------------------"
    Printer.Print Tab(1); Trim(rst("glosa"))
    
          
    Printer.EndDoc
    Exit Sub
End Sub


Private Sub impresion_formato_1_pedido(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser

    Printer.ScaleWidth = 11.8
    Printer.ScaleHeight = 15.3
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM movimiento_pedido WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM movimiento_pedido_detalle M,producto P,unidad U WHERE M.id_pedido='" & rst("id_pedido") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
    Printer.Print ""
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(50); BDBuscarCampo("comprobantes", "doc_abrev", "id_doc", rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
    Printer.Print Tab(10); Mid("FECHA PEDIDO" + Space(30), 1, 20) & ":" & formato_item(Day(rst("fecha")), 2) & Space(3) & formato_item(Month(rst("fecha")), 2) + Space(3) + str(Year(rst("fecha")))
    Printer.Print Tab(10); Mid("RUC/DNI" + Space(30), 1, 20) & ":" & rst("dni_save")
    Printer.Print Tab(10); Mid("NOMBRE RAZON" + Space(30), 1, 20) & ":" & UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rst("dni_save")))
    Printer.Print Tab(10); Mid("DIRECCION" + Space(30), 1, 20) & ":" & Mid(UCase(BDBuscarCampo("persona", "direccion", "dni", rst("dni_save"))) + Space(80), 1, 60)
    Printer.Print Tab(10); Mid("SUCURSAL" + Space(30), 1, 20) & ":" & Mid(UCase(BDBuscarCampo("almacen", "descripcion", "id_alm", KEY_ALM)) + Space(80), 1, 60)
    Printer.Print "" 'Tab(10); 'L 9
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(1); Mid("CODIGO" + Space(30), 1, 10) & Mid("PRODUCTO" + Space(30), 1, 40) & Mid("UNIDAD" + Space(30), 1, 11) & Mid("CANTIDAD" + Space(30), 1, 10)
    Printer.Print Tab(1); "============================================================================"
    For j = 0 To rstT.RecordCount - 1
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 10)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 11)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 40)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 10)
         Printer.Print Tab(1); codigo & descripcion & Und & cantidad
         rstT.MoveNext
    Next j
          inc = 0.5
     Printer.Print Tab(1); "============================================================================"
          Do While (Val(Printer.CurrentY) <= 7)
             Printer.CurrentY = Printer.CurrentY + inc
          Loop
    Printer.FontBold = True
    Printer.EndDoc
    Exit Sub
End Sub
Private Sub impresion_formato_1_partemaquina(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser

    Printer.ScaleWidth = 11.8
    Printer.ScaleHeight = 15.3
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM parte_maquianria WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    
    
    Printer.Print ""
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(50); BDBuscarCampo("comprobantes", "doc_abrev", "id_doc", rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
    Printer.Print Tab(10); Mid("FECHA PEDIDO" + Space(30), 1, 20) & ":" & formato_item(Day(rst("fecha")), 2) & Space(3) & formato_item(Month(rst("fecha")), 2) + Space(3) + str(Year(rst("fecha")))
    Printer.Print Tab(10); Mid("RUC/DNI" + Space(30), 1, 20) & ":" & rst("dni_save")
    Printer.Print Tab(10); Mid("NOMBRE RAZON" + Space(30), 1, 20) & ":" & UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rst("dni_save")))
    Printer.Print Tab(10); Mid("DIRECCION" + Space(30), 1, 20) & ":" & Mid(UCase(BDBuscarCampo("persona", "direccion", "dni", rst("dni_save"))) + Space(80), 1, 60)
    Printer.Print Tab(10); Mid("SUCURSAL" + Space(30), 1, 20) & ":" & Mid(UCase(BDBuscarCampo("almacen", "descripcion", "id_alm", KEY_ALM)) + Space(80), 1, 60)
    Printer.Print "" 'Tab(10); 'L 9
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    Printer.EndDoc
    Exit Sub
End Sub

Private Sub impresion_formato_2_factura(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String, impresiones As Integer

   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
     Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM movimiento_venta_detalle M,producto P,unidad U WHERE M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    Printer.Print ""
   ' Printer.Print "" 'Tab(10); 'L 2
   ' If rst("impresiones") > 1 Then
       ' Printer.Print Tab(20); "COPIA DE LA ORIGINAL" + Space(20) + BDBuscarCampo("comprobantes", "doc_abrev", "id_doc", rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
  '  Else
        Printer.Print Tab(114); rst("documento")
  '  End If
    Printer.Print Tab(15); UCase(rst("ncliente"))
    Printer.Print ""
    Printer.Print Tab(15); Mid(UCase(BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente"))) + Space(50), 1, 50) + Space(45) & formato_item(Day(rst("fecha_emision")), 2) & Space(3) & formato_item(Month(rst("fecha_emision")), 2) + Space(3) + str(Year(rst("fecha_emision")))
     Printer.Print ""
    Printer.Print Tab(15); rst("id_cliente")
       
    Printer.Print "" 'Tab(10); 'L 9
     Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("total")
         'Codigo = Mid((rstT("id_producto")) + Space(10), 1, 7)
         'Und = Mid((rstT("abreviatura")) + Space(10), 1, 10)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 85)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 13)
         precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(10), 1, 17)
         Printer.Print Tab(3); cantidad & descripcion & Und & Space(3) & precio & Format(rstT("total"), "#,##0.00")
         rstT.MoveNext
    Next j
          inc = 0.5
          
          Do While (Val(Printer.CurrentY) <= 11)
             Printer.CurrentY = Printer.CurrentY + inc
          Loop
    Printer.FontBold = True
    Printer.Print Tab(7); Mid(UCase(EnLetras(ptotal)) & Space(120), 1, 113) & Format(rst("valor_venta"), "#,##0.00")
    'Printer.Print "" 'Tab(10); 'L 9
   ' Printer.Print Tab(45); Mid("EXONERADO" + Space(50), 1, 15) & ":" & Space(2) & Format(rst("exonerado"), "#,##0.00")
    'Printer.Print Tab(85); Mid("VALOR VENTA" + Space(50), 1, 15) & ":" & Space(2) & Format(rst("valor_venta"), "#,##0.00")
    'Printer.Print Tab(120); Format(rst("valor_venta"), "#,##0.00")
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print Tab(120); Format(rst("igv"), "#,##0.00")
     Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print Tab(120); Format(rst("total"), "#,##0.00")
    'Printer.Print Tab(7); "ATENDIDO POR:" + Space(1) + KEY_VENDEDOR + Space(5) + Str(Time)
    Printer.EndDoc
    Exit Sub
End Sub
Private Sub impresion_formato_factura_daniel(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String, impresiones As Integer

   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
     Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM movimiento_venta_detalle M,producto P,unidad U WHERE M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    Printer.Print ""
   ' Printer.Print "" 'Tab(10); 'L 2
   ' If rst("impresiones") > 1 Then
       ' Printer.Print Tab(20); "COPIA DE LA ORIGINAL" + Space(20) + BDBuscarCampo("comprobantes", "doc_abrev", "id_doc", rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
  '  Else
        Printer.Print Tab(114); rst("documento")
  '  End If
    Printer.Print Tab(15); UCase(rst("ncliente"))
    Printer.Print ""
    Printer.Print Tab(15); Mid(UCase(BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente"))) + Space(50), 1, 50) + Space(35) & formato_item(Day(rst("fecha_emision")), 2) & Space(3) & formato_item(Month(rst("fecha_emision")), 2) + Space(3) + str(Year(rst("fecha_emision")))
     Printer.Print ""
    Printer.Print Tab(15); rst("id_cliente")
       
    Printer.Print "" 'Tab(10); 'L 9
     Printer.Print ""
     Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("total")
         'Codigo = Mid((rstT("id_producto")) + Space(10), 1, 7)
         'Und = Mid((rstT("abreviatura")) + Space(10), 1, 10)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 85)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 13)
         precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(10), 1, 17)
         Printer.Print Tab(3); cantidad & descripcion & Und & Space(3) & precio & Format(rstT("total"), "#,##0.00")
         rstT.MoveNext
    Next j
          inc = 0.5
          
          Do While (Val(Printer.CurrentY) <= 19.5)
             Printer.CurrentY = Printer.CurrentY + inc
          Loop
    Printer.FontBold = False
    Printer.Print Tab(7); Mid(UCase(EnLetras(ptotal)) & Space(120), 1, 113) & Format(rst("valor_venta"), "#,##0.00")
    'Printer.Print "" 'Tab(10); 'L 9
   ' Printer.Print Tab(45); Mid("EXONERADO" + Space(50), 1, 15) & ":" & Space(2) & Format(rst("exonerado"), "#,##0.00")
    'Printer.Print Tab(85); Mid("VALOR VENTA" + Space(50), 1, 15) & ":" & Space(2) & Format(rst("valor_venta"), "#,##0.00")
    'Printer.Print Tab(120); Format(rst("valor_venta"), "#,##0.00")
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print Tab(120); Format(rst("igv"), "#,##0.00")
     
    Printer.Print Tab(120); Format(rst("total"), "#,##0.00")
    'Printer.Print Tab(7); "ATENDIDO POR:" + Space(1) + KEY_VENDEDOR + Space(5) + Str(Time)
    Printer.EndDoc
    Exit Sub
End Sub

Private Sub impresion_formato_1_factura(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String, impresiones As Integer

   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    Printer.ScaleWidth = 11.8
    Printer.ScaleHeight = 15.3
    Printer.CurrentX = 0
    Printer.CurrentY = 0
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM movimiento_venta_detalle M,producto P,unidad U WHERE M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    Printer.Print ""
    Printer.Print "" 'Tab(10); 'L 2
    If rst("impresiones") > 1 Then
        Printer.Print Tab(20); "COPIA DE LA ORIGINAL" + Space(10) + BDBuscarCampo("comprobantes", "doc_abrev", "id_doc", rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
    Else
        Printer.Print Tab(50); BDBuscarCampo("comprobantes", "doc_abrev", "id_doc", rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
    End If
    Printer.Print Tab(20); UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_cliente")))
    Printer.Print Tab(20); UCase(BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente")))
    Printer.Print Tab(20); rst("id_cliente") + Space(25) & formato_item(Day(rst("fecha_emision")), 2) & Space(3) & formato_item(Month(rst("fecha_emision")), 2) + Space(3) + str(Year(rst("fecha_emision")))
       
    Printer.Print "" 'Tab(10); 'L 9
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("total")
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 7)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 6)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 40)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 7)
         precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(10), 1, 8)
         Printer.Print Tab(1); codigo & descripcion & Und & Space(1) & cantidad & precio & Format(rstT("total"), "#,##0.00")
         rstT.MoveNext
    Next j
          inc = 0.5
          
          Do While (Val(Printer.CurrentY) <= 7)
             Printer.CurrentY = Printer.CurrentY + inc
          Loop
    Printer.FontBold = True
    Printer.Print Tab(7); UCase(EnLetras(ptotal))
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print Tab(45); Mid("EXONERADO" + Space(50), 1, 15) & ":" & Space(2) & Format(rst("exonerado"), "#,##0.00")
    Printer.Print Tab(45); Mid("VALOR VENTA" + Space(50), 1, 15) & ":" & Space(2) & Format(rst("valor_venta"), "#,##0.00")
    Printer.Print Tab(45); Mid("IGV" + Space(50), 1, 15) & ":" & Space(2) & Format(rst("igv"), "#,##0.00")
    Printer.Print Tab(45); Mid("PRECIO VENTA" + Space(50), 1, 15) & ":" & Space(2) & Format(rst("total"), "#,##0.00")
    Printer.Print Tab(7); "ATENDIDO POR:" + Space(1) + KEY_VENDEDOR + Space(5) + str(Time)
    Printer.EndDoc
    Exit Sub
End Sub

Private Sub impresion_formato_2_boleta(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
 'Printer.PaperSize = 1
  
   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
   Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    Printer.ScaleWidth = 22.9
    Printer.ScaleHeight = 14
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM movimiento_venta_detalle M,producto P,unidad U WHERE M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
    Printer.Print ""
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(50); BDBuscarCampo("comprobantes", "doc_abrev", "id_doc", rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
    Printer.Print Tab(20); formato_item(Day(rst("fecha_emision")), 2) & Space(3) & formato_item(Month(rst("fecha_emision")), 2) + Space(3) + str(Year(rst("fecha_emision")))
    Printer.Print Tab(20); UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_cliente")))
    Printer.Print Tab(20); Mid(UCase(BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente"))) + Space(80), 1, 60) + rst("id_cliente")
    Printer.Print "" 'Tab(10); 'L 9
    Printer.CurrentY = Printer.CurrentY + 0.5
    For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("total")
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 10)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 8)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 60)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 10)
         precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(10), 1, 12)
         Printer.Print Tab(1); codigo & descripcion & Und & Space(1) & cantidad & precio & Format(rstT("total"), "#,##0.00")
         rstT.MoveNext
    Next j
          inc = 0.5
          
          Do While (Val(Printer.CurrentY) <= 7)
             Printer.CurrentY = Printer.CurrentY + inc
          Loop
    Printer.FontBold = True
    Printer.Print Tab(35); "****** MONTO TOTAL *******"; Space(10) & Format(ptotal, "#,##0.00")
    Printer.Print Tab(11); UCase(EnLetras(ptotal))
    Printer.Print Tab(11); "ATENDIDO POR:" + Space(1) + KEY_VENDEDOR
    Printer.EndDoc
    Exit Sub
End Sub
Private Sub impresion_formato_1_cotizacion(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String

   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "10"
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    'Printer.ScaleWidth = 11.8
   ' Printer.ScaleHeight = 15.22
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM movimiento_venta_detalle M,producto P,unidad U WHERE M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
    Printer.Print ""
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(50); BDBuscarCampo("comprobantes", "doc_abrev", "id_doc", rst("id_doc")) + ":" + rst("serie") + "-" + rst("numero")
    Printer.Print Tab(7); formato_item(Day(rst("fecha_emision")), 2) & Space(2) + "-" + formato_item(Month(rst("fecha_emision")), 2) + Space(2) + "-" + str(Year(rst("fecha_emision")))
    Printer.Print Tab(7); "DNI:" + Space(2) + rst("id_cliente")
    Printer.Print Tab(7); "CLIENTE :" + Space(2) + UCase(rst("ncliente"))
    
    If Val(rst("id_cliente")) > 0 Then
        Printer.Print Tab(7); "DIRECCION :" + Mid(UCase(BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente"))) + Space(80), 1, 60)
    Else
        Printer.Print Tab(7); "DIRECCION :"; Mid(UCase(BDBuscarCampoRuc("persona_publico", "direccion", "dni", rst("id_cliente"))) + Space(80), 1, 60) + rst("id_cliente")
    End If
    
    
    Printer.Print "" 'Tab(10); 'L 9
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(1); "COD" & Space(5) & "CANTIDAD" & Space(2) & "UND" & Space(2) & "DESCRIPCION" & Space(28) & "PRECIO" & Space(2) & "TOTAL"
    Printer.Print Tab(1); "-------------------------------------------------------------------------------------------"
    For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("total")
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 9)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 5)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 40)
         cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
         precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(10), 1, 8)
         'Printer.Print Tab(1); Codigo & descripcion & Und & Space(1) & cantidad & precio & Format(rstT("total"), "#,##0.00")
         Printer.Print Tab(1); codigo & cantidad & Und & descripcion & Space(1) & precio & Format(rstT("total"), "#,##0.00")
         Printer.CurrentY = Printer.CurrentY + 0.1
         rstT.MoveNext
    Next j
          inc = 0.5
        Printer.Print "" 'Tab(10); 'L 9
         Printer.Print "" 'Tab(10); 'L 9
          Printer.Print "" 'Tab(10); 'L 9
           Printer.Print "" 'Tab(10); 'L 9
           Printer.Print "" 'Tab(10); 'L 9
           Printer.Print "" 'Tab(10); 'L 9
           Printer.Print "" 'Tab(10); 'L 9
         '  Printer.Print "" 'Tab(10); 'L 9
         ' Do While (Val(Printer.CurrentY) <= 13)
            ' Printer.CurrentY = Printer.CurrentY + inc
        '  Loop
      
    Printer.Print Tab(35); "****** MONTO TOTAL *******"; Space(5) & Format(ptotal, "#,##0.00")
    Printer.Print Tab(9); UCase(EnLetras(ptotal))
    Printer.Print Tab(9); "ATENDIDO POR:" + Space(1) + KEY_VENDEDOR + Space(10) + str(Time)
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print "" '
    Printer.Print Tab(9); "ESTE COMPROBANTE NO TIENE VALIDEZ FISCAL"
    Printer.Print Tab(9); "   DEBE SER CANJEADO POR UNA BOLETA / FACTURA"
     Printer.Print "" '
    
    If KEY_DIRECCION <> KEY_DIRECCION_ALM Then
        Printer.Print Tab(9); KEY_DIRECCION_ALM
    End If
    Printer.EndDoc
    Call FrmVentas.nuevo
    Exit Sub
End Sub

Private Sub impresion_formato_1_boleta(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM movimiento_venta_detalle M,producto P,unidad U WHERE M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    
    Printer.Print Tab(70); rst("documento")
    Printer.Print Tab(20); Mid(rst("ncliente") + Space(50), 1, 60) & Space(1) & Format(rst("fecha_emision"), "dd-mm-YYYY")
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print ""
    If Val(rst("id_cliente")) > 0 Then
        Printer.Print Tab(20); Mid(UCase(BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente"))) + Space(80), 1, 59) + rst("id_cliente")
    Else
        Printer.Print Tab(20); Mid(UCase(BDBuscarCampoRuc("persona_publico", "direccion", "dni", rst("id_cliente"))) + Space(80), 1, 59) + rst("id_cliente")
    End If
    Printer.Print "" 'Tab(10); 'L 9
    Printer.CurrentY = Printer.CurrentY + 0.5
    If rstT.RecordCount < 1 Then
        strCadena = "SELECT * FROM movimiento_venta_detalle M WHERE M.id_venta='" & rst("id_venta") & "'  AND M.ruc='" & KEY_RUC & "' ORDER BY id_detalle_venta ASC"
        Call ConfiguraRstT(strCadena)
            For j = 0 To rstT.RecordCount - 1
            ptotal = ptotal + rstT("total")
            If rstT("id_producto") = "0" Then
                codigo = Mid("" + Space(10), 1, 7)
                Und = Mid("" + Space(10), 1, 6)
            Else
                codigo = Mid((rstT("id_producto")) + Space(10), 1, 7)
                strCadena = "SELECT abreviatura FROM unidad U,producto P WHERE P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND P.id_producto='" & rstT("id_producto") & "'"
                Call ConfiguraTemporal(strCadena)
                Und = Mid(rstTemporal("abreviatura") + Space(10), 1, 6)
            End If
            
            
            descripcion = Mid(rstT("detalle") + Space(80), 1, 40)
            If rstT("cantidad") = 0 Then
                cantidad = ""
            Else
            cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 7)
            End If
            If rstT("precio") = 0 And rstT("cantidad") = 0 Then
                precio = Mid("" + Space(10), 1, 8)
                rtotal = ""
            Else
                precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(10), 1, 8)
                rtotal = Format(rstT("total"), "#,##0.00")
            End If
            
            Printer.Print Tab(1); codigo & descripcion & Und & Space(1) & cantidad & precio & rtotal
            rstT.MoveNext
    Next j
    Else
    For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("total")
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 7)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 10)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 58)
         cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 12)
         precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(10), 1, 10)
         
         Printer.Print Tab(5); cantidad & descripcion & precio & Format(rstT("total"), "#,##0.00")
         rstT.MoveNext
    Next j
    End If
          inc = 0.5
          
          Do While (Val(Printer.CurrentY) <= 7.5)
             Printer.CurrentY = Printer.CurrentY + inc
          Loop
      
    
    Printer.Print Tab(12); UCase(EnLetras(ptotal))
    Printer.Print ""
    Printer.Print Tab(87); "S/" & Format(ptotal, "#,##0.00")
    
    Printer.EndDoc
    Call FrmVentas.nuevo
    Exit Sub
End Sub
Private Sub impresion_formato_1_boleta_electronica(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    in_moneda = rst("id_moneda")
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT M.id_producto,M.detalle,U.abreviatura,mm.descripcion,M.precio,M.total,M.cantidad,mm.descripcion as marca FROM movimiento_venta_detalle M,producto P,unidad U,marca mm WHERE M.id_venta='" & rst("id_venta") & "' AND P.id_marca=mm.id_marca and P.ruc=mm.id_usu AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    
    'strCadena = "SELECT * FROM view_venta_detalle WHERE id_venta='" & rst("id_venta") & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); "RAZON SOCIAL     :" & KEY_EMPRESA
    Printer.Print Tab(5); "DIRECCION FISCAL :" & Mid(KEY_DIRECCION, 1, 40)
    Printer.Print Tab(29); Mid(KEY_DIRECCION, 41, 40)
  
  
       strCadena = "SELECT * FROM almacen WHERE  ruc='" & KEY_RUC & "'"
       Call ConfiguraRstI(strCadena)
       If rstI.RecordCount > 0 Then
          rstI.MoveFirst
               Printer.Print Tab(5); "-----------------------------------------------------------------------"
          For k = 0 To rstI.RecordCount - 1
                    If KEY_DIRECCION <> Trim(rstI("direccion")) Then
                        Printer.Print Tab(5); rstI("descripcion") & Space(2) & rstI("direccion")
                    End If
                    rstI.MoveNext
          Next k
           Printer.Print Tab(5); "-----------------------------------------------------------------------"
       End If
       
            
       
       
    
    
    Printer.Print Tab(5); "RUC              :" & KEY_RUC
    Printer.Print Tab(5); "FECHA            :" & KEY_FECHA
    Printer.Print Tab(5); "-----------------------------------------------------------------------"
    Printer.Print Tab(50); "BOLETA DE VENTA ELECTRONICA"
    Printer.Print Tab(55); rst("documento")
    Printer.Print Tab(5); "DNI CLIENTE    :" & rst("id_cliente")
    Printer.Print Tab(5); "CLIENTE        :" & rst("ncliente")
    Printer.Print Tab(5); "DIRECCION      :" & rst("direccion")
    Printer.Print Tab(5); "FECHA EMISION  :" & Format(rst("fecha_emision"), "dd-mm-YYYY")
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       Printer.Print Tab(5); "CANT" & Space(1) & "DESCRIPCION" & Space(35) & "UND" & Space(5) & "V.UNIT" & Space(4) & "V.TOTAL"
       Printer.Print ""
        For j = 0 To rstT.RecordCount - 1
             ptotal = ptotal + rstT("total")
             codigo = Mid((rstT("id_producto")) + Space(10), 1, 6)
             Und = Mid((rstT("abreviatura")) + Space(10), 1, 5)
             descripcion = Mid(rstT("detalle") + Space(80), 1, 45)
             cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 5)
             precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(10), 1, 10)
             Printer.Print Tab(5); cantidad & descripcion & Und & precio & Format(rstT("total"), "#,##0.00")
             rstT.MoveNext
        Next j
    End If
    
    inc = 0.5
    Do While (Val(Printer.CurrentY) <= 7.5)
             Printer.CurrentY = Printer.CurrentY + inc
    Loop
      
   in_total = rst("total")
   in_valor_venta = rst("total") / (1 + KEY_IGV)
   in_igv = in_total - in_valor_venta
   Printer.Print Tab(52); "MONTO GRAVADO    :"; Format(in_valor_venta, "#,##0.00")
   Printer.Print Tab(52); "IGV (" & KEY_IGV * 100 & ")          :"; Format(in_igv, "#,##0.00")
   Printer.Print Tab(52); "MONTO TOTAL      :"; Format(in_total, "#,##0.00")
   Printer.Print Tab(5); UCase(EnLetras(in_total)) & Space(1) & get_moneda(in_moneda)
   Printer.Print Tab(5); "ATENDIDO POR:" & get_persona(rst("id_vendedor")) & Space(2) & Format(rst("hora"), "HH:mm:ss")
   Printer.Print ""
   Printer.Print Tab(10); "REPRESENTACION IMPRESA DE LA BOLETA DE VENTA ELECTRONICA"
   Printer.Print Tab(10); "            PARA CONSULTAR EL DOCUMENTO VISITE :"
   Printer.Print Tab(10); "       "
   
   If KEY_SERVIDOR_KEYFACIL = "si" Then
        Printer.Print Tab(0); "       https://keyfacil.com/consultar"
   Else
        Printer.Print Tab(0); "       http://facturacion.vitekey.com/consultar"
   End If
   
   Printer.Print Tab(10); "AUTORIZADO MEDIANTE RESOLUCION  N°" & KEY_RESOLUCION
   Printer.Print Tab(10); "       "
   Printer.Print Tab(10); "Resumen:" & rst("sunat_hash")
    
    
    Printer.EndDoc
    Call FrmVentas.nuevo
    Exit Sub
End Sub

Private Sub impresion_formato_1_factura_electronica(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
Dim in_moneda As String
Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    in_moneda = rst("id_moneda")
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT M.id_producto,M.detalle,U.abreviatura,mm.descripcion,M.precio,M.total,M.cantidad,mm.descripcion as marca FROM movimiento_venta_detalle M,producto P,unidad U,marca mm WHERE M.id_venta='" & rst("id_venta") & "' AND P.id_marca=mm.id_marca and P.ruc=mm.id_usu AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    'strCadena = "SELECT * FROM view_venta_detalle WHERE id_venta='" & rst("id_venta") & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); "RAZON SOCIAL     :" & KEY_EMPRESA
    Printer.Print Tab(5); "DIRECCION FISCAL :" & KEY_DIRECCION
    Printer.Print Tab(5); "RUC              :" & KEY_RUC
    Printer.Print Tab(5); "FECHA            :" & KEY_FECHA
    Printer.Print Tab(5); "------------------------------------------------------------------------"
    Printer.Print Tab(50); "FACTURA DE VENTA ELECTRONICA"
    Printer.Print Tab(55); rst("documento")
    Printer.Print Tab(5); "RUC CLIENTE    :" & rst("id_cliente")
    Printer.Print Tab(5); "RAZON SOCIAL   :" & rst("ncliente")
    Printer.Print Tab(5); "DIRECCION      :" & rst("direccion")
    Printer.Print Tab(5); "FECHA EMISION  :" & Format(rst("fecha_emision"), "dd-mm-YYYY")
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       Printer.Print Tab(5); "CANT" & Space(1) & "DESCRIPCION" & Space(35) & "UND" & Space(5) & "V.UNIT" & Space(4) & "V.TOTAL"
       Printer.Print ""
        For j = 0 To rstT.RecordCount - 1
             ptotal = ptotal + rstT("total")
             codigo = Mid((rstT("id_producto")) + Space(10), 1, 6)
             Und = Mid((rstT("abreviatura")) + Space(10), 1, 5)
             descripcion = Mid(rstT("detalle") + Space(80), 1, 45)
             cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 5)
             precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(10), 1, 10)
             Printer.Print Tab(5); cantidad & descripcion & Und & precio & Format(rstT("total"), "#,##0.00")
             rstT.MoveNext
        Next j
    End If
    
    'inc = 0.5
    'Do While (Val(Printer.CurrentY) <= 7.5)
    '         Printer.CurrentY = Printer.CurrentY + inc
    'Loop
   Printer.Print ""
   Printer.Print ""
   in_total = rst("total")
   in_valor_venta = rst("valor_venta")
   in_igv = rst("igv")
   Printer.Print Tab(52); "MONTO GRAVADO    :"; Format(in_valor_venta, "#,##0.00")
   Printer.Print Tab(52); "IGV (" & KEY_IGV * 100 & ")        :"; Format(in_igv, "#,##0.00")
   Printer.Print Tab(52); "MONTO TOTAL      :"; Format(in_total, "#,##0.00")
   
   Printer.Print Tab(5); UCase(EnLetras(in_total)) & Space(1) & get_moneda(in_moneda)
   Printer.Print Tab(5); "ATENDIDO POR:" & get_persona(rst("id_vendedor")) & Space(2) & Format(rst("hora"), "HH:mm:ss")
   Printer.Print ""
   Printer.Print Tab(10); "REPRESENTACION IMPRESA DE LA FACTURA DE VENTA ELECTRONICA"
   Printer.Print Tab(10); "            PARA CONSULTAR EL DOCUMENTO VISITE :"
   Printer.Print Tab(10); "       "
  If KEY_SERVIDOR_KEYFACIL = "si" Then
        Printer.Print Tab(0); "       https://keyfacil.com/consultar"
    Else
        Printer.Print Tab(0); "       http://facturacion.vitekey.com/consultar"
    End If
   Printer.Print Tab(10); "AUTORIZADO MEDIANTE RESOLUCION  N°" & KEY_RESOLUCION
   Printer.Print Tab(10); "       "
   Printer.Print Tab(10); "Resumen:" & rst("sunat_hash")
    
    
    Printer.EndDoc
    Call FrmVentas.nuevo
    Exit Sub
End Sub

Private Sub impresion_formato_1_nota_electronica(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT M.id_producto,M.detalle,U.abreviatura,mm.descripcion,M.precio,M.total,M.cantidad,mm.descripcion as marca FROM movimiento_venta_detalle M,producto P,unidad U,marca mm WHERE M.id_venta='" & rst("id_venta") & "' AND P.id_marca=mm.id_marca and P.ruc=mm.id_usu AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    'strCadena = "SELECT * FROM view_venta_detalle WHERE id_venta='" & rst("id_venta") & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); "RAZON SOCIAL     :" & KEY_EMPRESA
    Printer.Print Tab(5); "DIRECCION FISCAL :" & KEY_DIRECCION
    Printer.Print Tab(5); "RUC              :" & KEY_RUC
    Printer.Print Tab(5); "FECHA            :" & KEY_FECHA
    Printer.Print Tab(5); "------------------------------------------------------------------------"
    Printer.Print Tab(50); "NOTA DE CREDITO ELECTRONICA"
    Printer.Print Tab(55); rst("documento")
    Printer.Print Tab(5); "RUC CLIENTE    :" & rst("id_cliente")
    Printer.Print Tab(5); "RAZON SOCIAL   :" & rst("ncliente")
    Printer.Print Tab(5); "DIRECCION      :" & rst("direccion")
    Printer.Print Tab(5); "FECHA EMISION  :" & Format(rst("fecha_emision"), "dd-mm-YYYY")
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       Printer.Print Tab(5); "CANT" & Space(1) & "DESCRIPCION" & Space(35) & "UND" & Space(5) & "V.UNIT" & Space(4) & "V.TOTAL"
       Printer.Print ""
        For j = 0 To rstT.RecordCount - 1
             ptotal = ptotal + rstT("total")
             codigo = Mid((rstT("id_producto")) + Space(10), 1, 6)
             Und = Mid((rstT("abreviatura")) + Space(10), 1, 5)
             descripcion = Mid(rstT("detalle") + Space(80), 1, 45)
             cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 5)
             precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(10), 1, 10)
             Printer.Print Tab(5); cantidad & descripcion & Und & precio & Format(rstT("total"), "#,##0.00")
             rstT.MoveNext
        Next j
    End If
    
    'inc = 0.5
    'Do While (Val(Printer.CurrentY) <= 7.5)
    '         Printer.CurrentY = Printer.CurrentY + inc
    'Loop
   Printer.Print ""
   Printer.Print ""
   in_total = rst("total")
   in_valor_venta = rst("valor_venta")
   in_igv = rst("igv")
   Printer.Print Tab(52); "MONTO GRAVADO    :"; Format(in_valor_venta, "#,##0.00")
   Printer.Print Tab(52); "IGV (" & KEY_IGV * 100 & ")        :"; Format(in_igv, "#,##0.00")
   Printer.Print Tab(52); "MONTO TOTAL      :"; Format(in_total, "#,##0.00")
   
   Printer.Print Tab(5); UCase(EnLetras(in_total))
   Printer.Print Tab(5); "ATENDIDO POR:" & get_persona(rst("id_vendedor")) & Space(2) & Format(rst("hora"), "HH:mm:ss")
   Printer.Print ""
   Printer.Print Tab(10); "REPRESENTACION IMPRESA DE LA FACTURA DE VENTA ELECTRONICA"
   Printer.Print Tab(10); "            PARA CONSULTAR EL DOCUMENTO VISITE :"
   Printer.Print Tab(10); "       "
  If KEY_SERVIDOR_KEYFACIL = "si" Then
        Printer.Print Tab(0); "       https://keyfacil.com/consultar"
    Else
        Printer.Print Tab(0); "       http://facturacion.vitekey.com/consultar"
    End If
   Printer.Print Tab(10); "AUTORIZADO MEDIANTE RESOLUCION  N°" & KEY_RESOLUCION
   Printer.Print Tab(10); "       "
   Printer.Print Tab(10); "Resumen:" & rst("sunat_hash")
    
    
    Printer.EndDoc
    Call FrmVentas.nuevo
    Exit Sub
End Sub

Private Sub impresion_formato_boleta_daniel(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
      Printer.Font.name = "FontB11"
    Printer.Font.Size = "10"
    
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM movimiento_venta_detalle M,producto P,unidad U WHERE M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(70); rst("documento")
    Printer.Print Tab(16); Mid(rst("ncliente") + Space(50), 1, 60) & Space(1) & Format(rst("fecha_emision"), "dd-mm-YYYY")
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print ""
    'Printer.Print ""
    If Val(rst("id_cliente")) > 0 Then
        Printer.Print Tab(16); Mid(UCase(BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente"))) + Space(80), 1, 59) + rst("id_cliente")
    Else
        Printer.Print Tab(16); Mid(UCase(BDBuscarCampoRuc("persona_publico", "direccion", "dni", rst("id_cliente"))) + Space(80), 1, 59) + rst("id_cliente")
    End If
    Printer.Print "" 'Tab(10); 'L 9
    Printer.CurrentY = Printer.CurrentY + 0.5
    If rstT.RecordCount < 1 Then
        strCadena = "SELECT * FROM movimiento_venta_detalle M WHERE M.id_venta='" & rst("id_venta") & "'  AND M.ruc='" & KEY_RUC & "' ORDER BY id_detalle_venta ASC"
        Call ConfiguraRstT(strCadena)
            For j = 0 To rstT.RecordCount - 1
            ptotal = ptotal + rstT("total")
            If rstT("id_producto") = "0" Then
                codigo = Mid("" + Space(10), 1, 7)
                Und = Mid("" + Space(10), 1, 6)
            Else
                codigo = Mid((rstT("id_producto")) + Space(10), 1, 7)
                strCadena = "SELECT abreviatura FROM unidad U,producto P WHERE P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND P.id_producto='" & rstT("id_producto") & "'"
                Call ConfiguraTemporal(strCadena)
                Und = Mid(rstTemporal("abreviatura") + Space(10), 1, 6)
            End If
            
            
            descripcion = Mid(rstT("detalle") + Space(80), 1, 40)
            If rstT("cantidad") = 0 Then
                cantidad = ""
            Else
            cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 7)
            End If
            If rstT("precio") = 0 And rstT("cantidad") = 0 Then
                precio = Mid("" + Space(10), 1, 8)
                rtotal = ""
            Else
                precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(10), 1, 8)
                rtotal = Format(rstT("total"), "#,##0.00")
            End If
            
            Printer.Print Tab(1); codigo & descripcion & Und & Space(1) & cantidad & precio & rtotal
            rstT.MoveNext
    Next j
    Else
    For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("total")
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 7)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 10)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 62)
         cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 10)
         precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(10), 1, 10)
         
         Printer.Print Tab(0); cantidad & descripcion & precio & Format(rstT("total"), "#,##0.00")
         rstT.MoveNext
    Next j
    End If
          inc = 0.5
          
          Do While (Val(Printer.CurrentY) <= 11.5)
             Printer.CurrentY = Printer.CurrentY + inc
          Loop
      
    
    Printer.Print Tab(5); UCase(EnLetras(ptotal))
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print ""
    Printer.Print Tab(82); "S/" & Format(ptotal, "#,##0.00")
    
    Printer.EndDoc
    Call FrmVentas.nuevo
    Exit Sub
End Sub

Public Sub impresion_formato_orden(ByVal id_numero As Double, ByVal strCadena As String)
Dim ptotal As Double, Direccion As String, sector As String
  Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    'Printer.Font.name = "FontB11"Draft 17cpi
    'Printer.Font.name = "Draft 17cpi"
    'Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    Printer.Font.name = "Draft 17cpi"
    'Printer.Font.Size = "10"
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    Printer.ScaleWidth = 15.3
    Printer.ScaleHeight = 10.2
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = strCadena
    Call ConfiguraRst(strCadena)
   ' Strcadena = "SELECT * FROM tv_servicio_tecnico WHERE id_servicio='" & id_numero & "' AND ruc='" & KEY_RUC & "'"
   ' Call ConfiguraRst(Strcadena)
    '***** END COMPROBANTE***
    'Strcadena = "SELECT * FROM persona_direccion D WHERE dni='" & rst("dni") & "' AND id_direccion='" & rst("id_direccion") & "' "
    'Call ConfiguraRstT(Strcadena)
    'If rstT.RecordCount > 0 Then
     '   direcion = rstT("direccion")
      '  Strcadena = "SELECT * FROM urbanizacion WHERE id_urbanizacion='" & rstT("id_urbanizacion") & "'"
       ' Call ConfiguraRstT(Strcadena)
        'If rstT.RecordCount > 0 Then
         '   sector = rstT("descripcion")
        'End If
    'End If
    'Strcadena = "SELECT * FROM tv_servicio_tecnico_materiales M,producto P,unidad U WHERE M.id_servicio='" & id_numero & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
   ' Call ConfiguraRstT(Strcadena)
    '**** DETALLE COMPROBANTE ***
    'Printer.Print ""
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
    Printer.CurrentY = Printer.CurrentY - 0.5
      Printer.Print ""
        Printer.Print ""
    Printer.Print Tab(75); formato_item(id_numero, 6)
    Printer.Print ""
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(45); rst("tipo_servicio")
    Printer.Print Tab(10); Mid(BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_tecnico1")) & Space(50), 1, 85) + BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_tecnico2"))
    Printer.Print Tab(15); formato_item(Day(rst("fecha_solicitud")), 2) & Space(3) & formato_item(Month(rst("fecha_solicitud")), 2) + Space(3) + str(Year(rst("fecha_solicitud"))) + Space(32) + rst("hora_solicitud")
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(5); Mid(rst("nombre_completo") + Space(80), 1, 65) + rst("dni")
    Printer.Print Tab(5); Mid(rst("direccion") + Space(80), 1, 65) + sector
    'Printer.Print Tab(10); 'Mid(rst("descripcion_problema"), 1, 60)
   ' Printer.Print "" 'Tab(10); 'L 9
'    Printer.CurrentY = Printer.CurrentY + 0.5
   ' If rstT.RecordCount > 0 Then
   ' For j = 0 To rstT.RecordCount - 1
    '     ptotal = ptotal + rstT("total")
     '    Codigo = Mid((rstT("id_producto")) + Space(10), 1, 7)
      '   Und = Mid((rstT("abreviatura")) + Space(10), 1, 6)
       '  descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 40)
        ' cantidad = Mid(Format(Str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 7)
        ' precio = Mid(Format(Str(rstT("precio")), "#,##0.00") + Space(10), 1, 8)
        ' Printer.Print Tab(1); Codigo & descripcion & Und & Space(1) & cantidad & precio & Format(rstT("total"), "#,##0.00")
        ' rstT.MoveNext
'    Next j
 '   End If
  '        inc = 0.5
          
          'Do While (Val(Printer.CurrentY) <= 7)
          '   Printer.CurrentY = Printer.CurrentY + inc
          'Loop
    
    Printer.EndDoc
    Next i
    Exit Sub
End Sub


Private Sub impresion_formato_1_guia(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String, transporte As String
 'Printer.PaperSize = 1
  
   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    Printer.ScaleWidth = 11.8
    Printer.ScaleHeight = 15.3
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT T.id_transferencia,T.fecha,AO.direccion as origen,AD.direccion as destino,marca_placa,id_transporte,id_chofer,id_destinatario,id_motivo,serie,numero FROM movimiento_transferencia T,almacen AO,almacen AD WHERE T.id_alm_origen=AO.id_alm AND T.id_alm_destino=AD.id_alm AND T.ruc='" & KEY_RUC & "' AND AO.ruc='" & KEY_RUC & "' AND AD.ruc='" & KEY_RUC & "' AND T.id_doc='" & id_doc & "' AND T.serie='" & serie & "' AND T.numero='" & numero & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM movimiento_transferencia_detalle D,producto P,unidad U WHERE D.id_transferencia='" & rst("id_transferencia") & "' AND D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
    Printer.Print "" 'Tab(10); 'L 1
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(60); rst("serie") + "-" + rst("numero")
    
    Printer.Print Tab(30); rst("fecha") 'L 3
    Printer.Print Tab(30); UCase(rst("origen")) 'L 4
    If rst("id_motivo") = 3 Then
        direccion_destino = UCase(rst("destino"))
    Else
        direccion_destino = UCase(BDBuscarCampo("persona", "direccion", "dni", rst("id_destinatario")))
    End If
    Printer.Print Tab(1); direccion_destino + Space(40) + rst("id_destinatario") 'L 5
    If rst("id_transporte") <> "" Then
        transporte = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_transporte"))
    Else
        transporte = ""
    End If
    Printer.Print Tab(1); rst("marca_placa") + Space(40) + transporte 'L 6
    If rst("id_chofer") <> "" Then
        Printer.Print Tab(1); BDBuscarCampo("persona", "licencia", "dni", rst("id_chofer")) + Space(40) + rst("id_transporte") 'L 7
    Else
        Printer.Print Tab(1); ""
    End If
    Printer.Print "" 'Tab(10); 'L 8
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print "" 'Tab(10); 'L 10
     For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("total")
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 6)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 10)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 45)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 8)
         Printer.Print Tab(1); codigo & Space(2) & descripcion & Space(2) & Und & Space(2) & cantidad
         rstT.MoveNext
    Next j
          inc = 0.5
          
          Do While (Val(Printer.CurrentY) <= 7)
             Printer.CurrentY = Printer.CurrentY + inc
          Loop
    Printer.FontBold = True
    Printer.Print Tab(40); "***** PESO TOTAL *******"; Space(5) & Format(ptotal, "#,##0.00") + Space(2) + "Kg."
    Printer.EndDoc
    Exit Sub
End Sub
Private Sub impresion_formato_guia_suelta(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
   ' Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.CurrentX = 0
    Printer.CurrentY = 0

    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "10"
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    


   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    'Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    'Printer.ScaleWidth = 22.9
    'Printer.ScaleHeight = 14
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT T.id_transferencia,T.dni_atencion,T.direccion,T.atencion,T.fecha,AO.direccion as origen,AD.direccion as destino,marca_placa,id_transporte,id_chofer,id_destinatario,id_motivo,serie,numero,id_venta FROM movimiento_transferencia T,almacen AO,almacen AD WHERE T.id_alm_origen=AO.id_alm AND T.id_alm_destino=AD.id_alm AND T.ruc='" & KEY_RUC & "' AND AO.ruc='" & KEY_RUC & "' AND AD.ruc='" & KEY_RUC & "' AND T.id_doc='" & id_doc & "' AND T.serie='" & serie & "' AND T.numero='" & numero & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    
  
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(90); rst("serie") + "-" + rst("numero")
    'Printer.Print "" 'Tab(10); 'L 2
    Printer.CurrentY = Printer.CurrentY + 0.5
   ' Printer.Print "" 'Tab(10); 'L 2
   If rst("id_venta") > 0 Then
   strCadena = "SELECT C.doc_abrev,serie,numero FROM movimiento_venta M,comprobantes C WHERE M.id_doc=C.id_doc and   id_venta='" & rst("id_venta") & "' "
   Call ConfiguraRstZ(strCadena)
   
   If rst.RecordCount > 0 Then
        nboleta = Mid(rstZ("doc_abrev"), 1, 3) & ":" & rstZ("serie") & "-" & rstZ("numero")
   Else
        nboleta = ""
   End If
   Else
    nboleta = ""
   End If
    If rst("id_motivo") = 3 Then
        direccion_destino = UCase(rst("destino"))
        Destinatario = KEY_EMPRESA
    Else
        strCadena = "SELECT nombre_completo,direccion,id_departamento,id_distrito,id_provincia FROM persona WHERE dni='" & rst("id_destinatario") & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
            If rstZ("id_departamento") <> "0" Then
                strCadena = "select * from departamentos WHERE id_depa='" & rstZ("id_departamento") & "'"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    ndepartamento = "- " & UCase(rstL("descripcion"))
                Else
                    ndepartamento = " "
                End If
                
                If rstZ("id_provincia") <> "0" Then
                strCadena = "select * from provincia WHERE id_provincia='" & rstZ("id_provincia") & "'"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    nprovincia = "- " & UCase(rstL("descripcion"))
                Else
                    nprovincia = " "
                End If
                End If
                
                
            If rstZ("id_distrito") <> "0" Then
                strCadena = "select * from distrito WHERE id_distrito='" & rstZ("id_distrito") & "'"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    ndistrito = "- " & UCase(rstL("descripcion"))
                Else
                    ndistrito = " "
                End If
                End If
            End If
            If Len(Trim(rst("direccion"))) > 1 Then
                direccion_destino = rst("direccion")
            Else
                direccion_destino = rstZ("direccion") & Space(1) & ndistrito & Space(1) & nprovincia & Space(1) & ndepartamento
           
            End If
            
            Destinatario = rstZ("nombre_completo")
        Else
            direccion_destino = ""
            Destinatario = ""
        End If
        direccion_destino = UCase(direccion_destino)
    End If
    
    Printer.Print Tab(20); Mid(Destinatario + Space(80), 1, 90)
    Printer.Print Tab(20); rst("fecha")
    Printer.Print "" 'Tab(10); 'L 2
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(20); Mid(UCase(rst("origen")) + Space(80), 1, 85) & Space(2) & rst("id_destinatario")
    Printer.Print Tab(20); Mid(direccion_destino + Space(80), 1, 83) & Space(2) & nboleta
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 8
    strCadena = "SELECT D.id_producto,D.total,U.abreviatura,P.nombre_prod,D.cantidad,ma.descripcion as marca FROM movimiento_transferencia_detalle D,producto P,unidad U,marca ma WHERE  ma.id_marca=P.id_marca and ma.id_usu=P.ruc and  D.id_transferencia='" & rst("id_transferencia") & "' AND D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
     For j = 0 To rstT.RecordCount - 1
      
         
         
         
                codigo = Mid(rstT("id_producto") + Space(50), 1, 5)
                cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                Und = Mid(rstT("abreviatura") + Space(10), 1, 4)
                strmarca = Mid(rstT("marca") + Space(20), 1, 23)
                descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 48)
                
               
                Printer.Print Tab(15); cantidad & Space(2) & Und & Space(1) & descripcion & Space(1) & strmarca
         
         
         
         
         
         rstT.MoveNext
    Next j
          inc = 0.5
    Do While (Val(Printer.CurrentY) <= 40)
       Printer.CurrentY = Printer.CurrentY + inc
    Loop
    If Len(Trim(rst("atencion"))) > 1 Then
        Printer.Print Tab(20); "ATENCION  A:" & Space(2) & rst("atencion") & Space(2); "(" & rst("dni_atencion") & ")"
        
    End If
    'Printer.FontBold = True
   ' Printer.Print Tab(48); "***** PESO TOTAL *******"; Space(50) & Format(ptotal, "#,##0.00") + Space(2) + "Kg."
   ' Printer.Print ""
   ' Printer.Print ""
   
    
    
    'If rst("id_transporte") <> "" Then
       ' strCadena = "SELECT nombre_completo FROM persona WHERE dni='" & rst("id_transporte") & "'"
      '  Call ConfiguraRstL(strCadena)
        
     '   transporte = rstL("nombre_completo")
    'Else
    '    transporte = ""
   ' End If
    'Printer.Print Tab(9); Mid(transporte + Space(80), 1, 65)
    'Printer.CurrentY = Printer.CurrentY + 0.2
   ' Printer.Print Tab(9); Mid(rst("id_transporte") + Space(80), 1, 65)
    
    
  '  Printer.Print ""
 '   Printer.Print ""
'Printer.CurrentY = Printer.CurrentY + 0.2
    'If rst("id_venta") <> 0 Then
    'strCadena = "select serie,numero,doc_des from movimiento_venta M,comprobantes C WHERE M.id_doc=C.id_doc AND  id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
    'Call ConfiguraRstT(strCadena)
    
    
    'Printer.Print Tab(9); Trim(rstT("serie")) + "-" + rstT("numero") + Space(5) + rstT("doc_des")
    'End If
    Printer.EndDoc
    Exit Sub
End Sub

Private Sub impresion_formato_guia_suelta_rep(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
   
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "10"
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    'Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    'Printer.ScaleWidth = 22.9
    'Printer.ScaleHeight = 14
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT T.id_transferencia,T.fecha,AO.direccion as origen,AD.direccion as destino,marca_placa,id_transporte,id_chofer,id_destinatario,id_motivo,serie,numero,id_venta,dni_atencion,atencion FROM movimiento_transferencia T,almacen AO,almacen AD WHERE T.id_alm_origen=AO.id_alm AND T.id_alm_destino=AD.id_alm AND T.ruc='" & KEY_RUC & "' AND AO.ruc='" & KEY_RUC & "' AND AD.ruc='" & KEY_RUC & "' AND T.id_doc='" & id_doc & "' AND T.serie='" & serie & "' AND T.numero='" & numero & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT D.id_producto,D.total,U.abreviatura,P.nombre_prod,D.cantidad,C.descripcion as color,S.descripcion as modelo FROM movimiento_transferencia_detalle D,producto P,unidad U,imp_color C,linea_sub S WHERE  P.id_color=C.id_color and P.id_sublinea=S.id_tipo and P.ruc=S.id_usu and   D.id_transferencia='" & rst("id_transferencia") & "' AND D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
    Printer.Print "" 'Tab(10); 'L 1
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(90); rst("serie") + "-" + rst("numero")
    'Printer.Print "" 'Tab(10); 'L 2
    Printer.CurrentY = Printer.CurrentY + 0.5
   ' Printer.Print "" 'Tab(10); 'L 2
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
      If rst("id_venta") > 0 Then
        strCadena = "SELECT documento FROM movimiento_venta WHERE id_venta='" & rst("id_venta") & "' "
        Call ConfiguraRstZ(strCadena)
    If rst.RecordCount > 0 Then
        nboleta = rstZ("documento")
    Else
        nboleta = ""
    End If
   Else
    nboleta = ""
   End If
    Printer.Print Tab(35); rst("fecha") & Space(70) & nboleta
    'Printer.Print "" 'Tab(10); 'L 2
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(45); Mid(UCase(rst("origen")) + Space(80), 1, 70)
    
    
    
    If rst("id_motivo") = 3 Then
        direccion_destino = UCase(rst("destino"))
        Destinatario = KEY_EMPRESA
    Else
        
        
        strCadena = "SELECT nombre_completo,direccion,id_departamento,id_distrito,id_provincia FROM persona WHERE dni='" & rst("id_destinatario") & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
            If rstZ("id_departamento") <> "0" Then
                strCadena = "select * from departamentos WHERE id_depa='" & rstZ("id_departamento") & "'"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    ndepartamento = "- " & UCase(rstL("descripcion"))
                Else
                    ndepartamento = " "
                End If
                
                If rstZ("id_provincia") <> "0" Then
                strCadena = "select * from provincia WHERE id_provincia='" & rstZ("id_provincia") & "'"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    nprovincia = "- " & UCase(rstL("descripcion"))
                Else
                    nprovincia = " "
                End If
                End If
                
                
            If rstZ("id_distrito") <> "0" Then
                strCadena = "select * from distrito WHERE id_distrito='" & rstZ("id_distrito") & "'"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    ndistrito = "- " & UCase(rstL("descripcion"))
                Else
                    ndistrito = " "
                End If
                End If
            End If
            direccion_destino = rstZ("direccion") & Space(1) & ndistrito & Space(1) & nprovincia & Space(1) & ndepartamento
            Destinatario = rstZ("nombre_completo")
        Else
            direccion_destino = ""
            Destinatario = ""
        End If
        direccion_destino = UCase(direccion_destino)
    End If
    
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(35); Mid(Destinatario + Space(80), 1, 90) + rst("id_destinatario")  'L 5
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(35); Mid(direccion_destino + Space(80), 1, 105)
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 8
    strCadena = "SELECT D.id_producto,D.total,U.abreviatura,P.nombre_prod,D.cantidad,ma.descripcion as marca FROM movimiento_transferencia_detalle D,producto P,unidad U,marca ma WHERE  ma.id_marca=P.id_marca and ma.id_usu=P.ruc and  D.id_transferencia='" & rst("id_transferencia") & "' AND D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
     For j = 0 To rstT.RecordCount - 1
                codigo = Mid(rstT("id_producto") + Space(50), 1, 5)
                cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                Und = Mid(rstT("abreviatura") + Space(10), 1, 4)
                strmarca = Mid(rstT("marca") + Space(20), 1, 23)
                descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 48)
                
               
                Printer.Print Tab(15); cantidad & Space(2) & Und & Space(1) & descripcion & Space(1) & strmarca
         
         
         
         
         
         rstT.MoveNext
    Next j
          inc = 0.5
    Do While (Val(Printer.CurrentY) <= 40)
       Printer.CurrentY = Printer.CurrentY + inc
    Loop
    If Len(Trim(rst("atencion"))) > 1 Then
        Printer.Print Tab(20); "ATENCION  A:" & Space(2) & rst("atencion") & Space(2); "(" & rst("dni_atencion") & ")"
        
    End If
    'Printer.FontBold = True
   ' Printer.Print Tab(48); "***** PESO TOTAL *******"; Space(50) & Format(ptotal, "#,##0.00") + Space(2) + "Kg."
   ' Printer.Print ""
   ' Printer.Print ""
   
    
    
    'If rst("id_transporte") <> "" Then
       ' strCadena = "SELECT nombre_completo FROM persona WHERE dni='" & rst("id_transporte") & "'"
      '  Call ConfiguraRstL(strCadena)
        
     '   transporte = rstL("nombre_completo")
    'Else
    '    transporte = ""
   ' End If
    'Printer.Print Tab(9); Mid(transporte + Space(80), 1, 65)
    'Printer.CurrentY = Printer.CurrentY + 0.2
   ' Printer.Print Tab(9); Mid(rst("id_transporte") + Space(80), 1, 65)
    
    
  '  Printer.Print ""
 '   Printer.Print ""
'Printer.CurrentY = Printer.CurrentY + 0.2
    'If rst("id_venta") <> 0 Then
    'strCadena = "select serie,numero,doc_des from movimiento_venta M,comprobantes C WHERE M.id_doc=C.id_doc AND  id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
    'Call ConfiguraRstT(strCadena)
    
    
    'Printer.Print Tab(9); Trim(rstT("serie")) + "-" + rstT("numero") + Space(5) + rstT("doc_des")
    'End If
    Printer.EndDoc
    Exit Sub
End Sub


Private Sub impresion_formato_guia_suelta_galvez(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String, transporte As String
' Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    '
    '***** COMPROBANTE *****
   '------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "5 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "6 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "7 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "8 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "9 ---------------------------------------------------------------------"
   ' Printer.Print "" 'Tab(10); "10---------------------------------------------------------------------"
   ' Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT T.id_transferencia,T.fecha,AO.direccion as origen,AD.direccion as destino,marca_placa,id_transporte,id_chofer,id_destinatario,id_motivo,serie,numero,id_venta ,mtc FROM movimiento_transferencia T,almacen AO,almacen AD WHERE T.id_alm_origen=AO.id_alm AND T.id_alm_destino=AD.id_alm AND T.ruc='" & KEY_RUC & "' AND AO.ruc='" & KEY_RUC & "' AND AD.ruc='" & KEY_RUC & "' AND T.id_doc='" & id_doc & "' AND T.serie='" & serie & "' AND T.numero='" & numero & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM movimiento_transferencia_detalle D,producto P,unidad U WHERE D.id_transferencia='" & rst("id_transferencia") & "' AND D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    Printer.Print Tab(85); rst("serie") + "-" + rst("numero")
    
    'Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(12); str(rst("fecha")) + Space(35) + str(rst("fecha"))
    strCadena = "SELECT direccion,nombre_completo FROM persona WHERE dni='" & rst("id_destinatario") & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
         strdestino = rstL("direccion")
         strrazon = rstL("nombre_completo")
    Else
        strdestino = "NO ESPECIFICADO"
        strrazon = ""
    End If
    Printer.Print Tab(12); Mid(UCase(strrazon) + Space(80), 1, 70)
    
    Printer.Print Tab(12); Mid(UCase(rst("origen")) + Space(80), 1, 75) & Space(5) & rst("id_destinatario")
    strCadena = "SELECT documento FROM movimiento_venta where id_venta ='" & rst("id_venta") & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRstK(strCadena)
    If rstK.RecordCount > 0 Then
        Printer.Print Tab(12); Mid(UCase(strdestino) + Space(80), 1, 75) & Space(5) & rstK("documento")
    Else
        Printer.Print Tab(3); Mid(UCase(strdestino) + Space(80), 1, 75) & Space(5) & "--"
    End If
    
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print Tab(81); rst("marca_placa")
    Printer.CurrentY = Printer.CurrentY + 0.2
    'Printer.Print Tab(95); rst("mtc")
    
    If rst("id_chofer") <> "" Then
        strCadena = "SELECT * FROM persona WHERE dni='" & rst("id_chofer") & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            Licencia = rstL("licencia")
            Chofer = rstL("nombre_completo")
        End If
        
    Else
        Licencia = ""
        Chofer = ""
    End If
    
    'Printer.Print Tab(6); Mid(rst("id_destinatario") + Space(80), 1, 20); Space(63) + Mid(Licencia + Space(80), 1, 20)
    'Printer.Print Tab(75); Mid(Chofer + Space(80), 1, 50)
    
    
 '   Printer.Print Tab(1); direccion_destino + Space(40) + rst("id_destinatario") 'L 5
    
    'If rst("id_chofer") <> "" Then
     '   Printer.Print Tab(1); BDBuscarCampo("persona", "licencia", "dni", rst("id_chofer")) + Space(40) + rst("id_transporte") 'L 7
  '  Else
   '     Printer.Print Tab(1); ""
   ' End If
    
    Printer.Print ""
    Printer.Print ""
    ptotal = 0
     For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("peso") * rstT("cantidad")
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 6)
         Und = Mid((rstT("descripcion")) + Space(10), 1, 10)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 50)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 8)
         Printer.Print Tab(4); cantidad & Space(4) & codigo & "-" & descripcion & Space(10) & Und & Space(4) & Mid(Format(rstT("peso"), "#,##0.00") + Space(10), 1, 6) & Space(3) & Format(cantidad * rstT("peso"), "#,##0.00")
         rstT.MoveNext
    Next j
          inc = 0.5
    Do While (Val(Printer.CurrentY) <= 22)
       Printer.CurrentY = Printer.CurrentY + inc
    Loop
    
    'Printer.FontBold = True
    Printer.Print Tab(48); "***** PESO TOTAL *******"; Space(35) & Format(ptotal, "#,##0.00") + Space(2) + "Kg."
    Printer.Print ""
    Printer.Print ""
   
    
    
    If rst("id_transporte") <> "" Then
        strCadena = "SELECT nombre_completo FROM persona WHERE dni='" & rst("id_transporte") & "'"
        Call ConfiguraRstL(strCadena)
        
        transporte = rstL("nombre_completo")
    Else
        transporte = ""
    End If
    Printer.Print Tab(9); Mid(transporte + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print Tab(9); Mid(rst("id_transporte") + Space(80), 1, 65)
    
    
    Printer.Print ""
    Printer.Print ""
Printer.CurrentY = Printer.CurrentY + 0.2

    Printer.EndDoc
    Exit Sub
End Sub

Public Sub impresion_formato_guia_suelta2(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String, transporte As String
' Printer.PaperSize = 1
     Printer.TrackDefault = True 'siempre apunta a la impresora predeter
     Printer.CurrentX = 0
    Printer.CurrentY = 0
 
   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    'Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    'Printer.ScaleWidth = 11.8
    'Printer.ScaleHeight = 15.3
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
   '------------------------------------------------------------------"
    
    'Printer.Print "" 'Tab(10); "6 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "7 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "8 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "9 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "10---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    strCadena = "SELECT T.id_transferencia,T.fecha,AO.direccion as origen,T.direccion as destino,marca_placa,id_transporte,id_chofer,id_destinatario,id_motivo,serie,numero,T.mtc,T.id_venta,T.licencia FROM movimiento_transferencia T,almacen AO,almacen AD WHERE T.id_alm_origen=AO.id_alm AND T.id_alm_destino=AD.id_alm AND T.ruc='" & KEY_RUC & "' AND AO.ruc='" & KEY_RUC & "' AND AD.ruc='" & KEY_RUC & "' AND T.id_doc='" & id_doc & "' AND T.serie='" & serie & "' AND T.numero='" & numero & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM movimiento_transferencia_detalle D,producto P,unidad U WHERE D.id_transferencia='" & rst("id_transferencia") & "' AND D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    Printer.Print ""
    Printer.Print Tab(90); rst("serie") + "-" + rst("numero")
    Printer.Print Tab(20); str(rst("fecha")) + Space(35) + str(rst("fecha"))
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(10); Mid(UCase(rst("origen")) + Space(80), 1, 60) & Space(10) & Mid(UCase(rst("destino")) + Space(80), 1, 70)
    Printer.Print ""
     Printer.CurrentY = Printer.CurrentY + 0.2
     'Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
    Printer.Print Tab(88); Mid(rst("marca_placa") + Space(80), 1, 50)
     Printer.Print Tab(10); Mid(UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_destinatario"))) + Space(80), 1, 90) & rst("mtc")
    
    'Printer.CurrentY = Printer.CurrentY + 0.2
   
    If rst("id_chofer") <> "" Then
        Licencia = BDBuscarCampo("persona", "licencia", "dni", rst("id_chofer"))
        Chofer = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_chofer"))
    Else
        Licencia = ""
        Chofer = ""
    End If
    Licencia = rst("licencia")
    Printer.Print Tab(10); Mid(rst("id_destinatario") + Space(80), 1, 20) + Space(75) + Mid(Licencia + Space(80), 1, 20)
    Printer.Print Tab(72); Mid(Chofer + Space(80), 1, 20)
    
    
 '   Printer.Print Tab(1); direccion_destino + Space(40) + rst("id_destinatario") 'L 5
    
    'If rst("id_chofer") <> "" Then
     '   Printer.Print Tab(1); BDBuscarCampo("persona", "licencia", "dni", rst("id_chofer")) + Space(40) + rst("id_transporte") 'L 7
  '  Else
   '     Printer.Print Tab(1); ""
   ' End If
    
    
    Printer.Print ""
    Printer.Print ""
    'Printer.Print ""
    ptotal = 0
     For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("cantidad") * rstT("peso")
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 6)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 10)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 65)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 8)
         Printer.Print Tab(10); codigo & Space(4) & descripcion & Space(17) & Und & Space(4) & Mid(Format(rstT("peso"), "#,##0.00") + Space(10), 1, 6) & Space(3) & cantidad
         rstT.MoveNext
    Next j
          inc = 0.5
    Do While (Val(Printer.CurrentY) <= 22)
       Printer.CurrentY = Printer.CurrentY + inc
    Loop
    
    'Printer.FontBold = True
    Printer.Print Tab(48); "***** PESO TOTAL *******"; Space(40) & Format(ptotal, "#,##0.00") + Space(2) + "Kg."
    
    Printer.Print ""
    If rst("id_transporte") <> "" Then
        transporte = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_transporte"))
    Else
        transporte = ""
    End If
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(15); Mid(transporte + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(15); Mid(rst("id_transporte") + Space(80), 1, 65)
    
    If rst("id_transporte") <> "" Then
        dirtransporte = BDBuscarCampo("persona", "direccion", "dni", rst("id_transporte"))
    Else
        transporte = ""
    End If
    
    Printer.Print Tab(15); Mid(dirtransporte + Space(80), 1, 36)
    Printer.Print ""
    Printer.Print ""
    
    Printer.Print Tab(10); Right(Trim(serie), 3) + "-" + Right(numero, 7) + Space(8) + BDBuscarCampo("comprobantes", "doc_abrev", "id_doc", BDBuscarCampoRuc("movimiento_venta", "id_doc", "id_venta", rst("id_venta")))
    Printer.EndDoc
    Exit Sub
End Sub
Public Sub impresion_formato_guia_daniel(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
    Dim ptotal As Double, direccion_destino As String, transporte As String
    
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "10"
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    
    'Printer.ScaleMode = 7
    'Printer.Width = 11
    
   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    'Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    'Printer.ScaleWidth = 11.8
    'Printer.ScaleHeight = 15.3
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
   '------------------------------------------------------------------"
    
    'Printer.Print "" 'Tab(10); "6 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "7 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "8 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "9 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "10---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   'Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   'Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    strCadena = "SELECT T.id_transferencia,T.fecha,AO.direccion as origen,T.direccion as destino,marca_placa,placa,id_transporte,id_chofer,id_destinatario,id_motivo,serie,numero,T.mtc,T.id_venta,T.licencia FROM movimiento_transferencia T,almacen AO,almacen AD WHERE T.id_alm_origen=AO.id_alm AND T.id_alm_destino=AD.id_alm AND T.ruc='" & KEY_RUC & "' AND AO.ruc='" & KEY_RUC & "' AND AD.ruc='" & KEY_RUC & "' AND T.id_doc='" & id_doc & "' AND T.serie='" & serie & "' AND T.numero='" & numero & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM movimiento_transferencia_detalle D,producto P,unidad U WHERE D.id_transferencia='" & rst("id_transferencia") & "' AND D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    Printer.Print ""
    Printer.Print Tab(100); rst("serie") + "-" + rst("numero")
    Printer.Print ""
    Printer.Print Tab(12); str(rst("fecha")) + Space(35) + str(rst("fecha"))
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(0); Mid(UCase(rst("origen")) + Space(80), 1, 60) & Space(10) & Mid(UCase(rst("destino")) + Space(80), 1, 70)
    Printer.Print ""
    Printer.Print ""
     'Printer.CurrentY = Printer.CurrentY + 0.2
     Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
         Printer.Print Tab(85); Mid(rst("marca_placa") + Space(80), 1, 20) & rst("placa")
         'Printer.Print ""
     Printer.Print Tab(5); Mid(UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_destinatario"))) + Space(80), 1, 90) & rst("mtc")
    
    'Printer.CurrentY = Printer.CurrentY + 0.2
   
    If rst("id_chofer") <> "" Then
        Licencia = BDBuscarCampo("persona", "licencia", "dni", rst("id_chofer"))
        Chofer = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_chofer"))
    Else
        Licencia = ""
        Chofer = ""
    End If
    Licencia = rst("licencia")
    Printer.Print Tab(8); Mid(rst("id_destinatario") + Space(80), 1, 20) + Space(80) + Mid(Licencia + Space(80), 1, 20)
    Printer.Print Tab(80); Mid(Chofer + Space(80), 1, 20)
    
    
 '   Printer.Print Tab(1); direccion_destino + Space(40) + rst("id_destinatario") 'L 5
    
    'If rst("id_chofer") <> "" Then
     '   Printer.Print Tab(1); BDBuscarCampo("persona", "licencia", "dni", rst("id_chofer")) + Space(40) + rst("id_transporte") 'L 7
  '  Else
   '     Printer.Print Tab(1); ""
   ' End If
    
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    'Printer.Print ""
    ptotal = 0
     For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("cantidad") * rstT("peso")
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 6)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 10)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 65)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 8)
         Printer.Print Tab(0); codigo & Space(4) & descripcion & Space(20) & Und & Space(4) & Mid(Format(rstT("peso"), "#,##0.00") + Space(10), 1, 6) & Space(3) & cantidad
         rstT.MoveNext
    Next j
          inc = 0.5
    
    Do While (Val(Printer.CurrentY) <= 50)
       Printer.CurrentY = Printer.CurrentY + inc
    Loop
    
    'Printer.FontBold = True
    Printer.Print Tab(48); "***** PESO TOTAL *******"; Space(40) & Format(ptotal, "#,##0.00") + Space(2) + "Kg."
    
    Printer.Print ""
    Printer.Print ""
    If rst("id_transporte") <> "" Then
        transporte = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_transporte"))
    Else
        transporte = ""
    End If
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(10); Mid(transporte + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(10); Mid(rst("id_transporte") + Space(80), 1, 65)
    
    If rst("id_transporte") <> "" Then
        dirtransporte = BDBuscarCampo("persona", "direccion", "dni", rst("id_transporte"))
    Else
        transporte = ""
    End If
    
    Printer.Print Tab(10); Mid(dirtransporte + Space(80), 1, 36)
    Printer.Print ""
    Printer.Print ""
    If rst("id_venta") > 0 Then
        Dim strfactura() As String
        
        strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            strfactura = Split(rstL("documento"), ":")
            Printer.Print Tab(10); Right(Trim(rstL("serie")), 3) + "-" + Right(rstL("numero"), 7) + Space(8) + strfactura(0)
        End If
        
    End If
    
    Printer.EndDoc
    Exit Sub
End Sub
Public Sub impresion_formato_grupo_jm(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
    Dim ptotal As Double, direccion_destino As String, transporte As String
    
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    'Printer.Font.name = "MS Gothic"
    'Printer.Font.Size = "8"
     
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "10"
   
    
    Printer.ScaleMode = 7 '  1:TWIST   6:Milimetros   7:Centrimetros
    'Printer.Width = 11
    
   ' Printer.CurrentX = 0
   ' Printer.CurrentY = 0
    'Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    'Printer.ScaleWidth = 11.8
    'Printer.ScaleHeight = 15.3
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
   '------------------------------------------------------------------"
    
    'Printer.Print "" 'Tab(10); "6 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "7 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "8 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "9 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "10---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
  ' Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
  ' Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   'Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   'Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   'Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   'Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   'Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   'Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    strCadena = "SELECT T.id_transferencia,T.fecha,AO.direccion as origen,T.direccion as destino,marca_placa,placa,id_transporte,id_chofer,id_destinatario,id_motivo,serie,numero,T.mtc,T.id_venta,T.licencia FROM movimiento_transferencia T,almacen AO,almacen AD WHERE T.id_alm_origen=AO.id_alm AND T.id_alm_destino=AD.id_alm AND T.ruc='" & KEY_RUC & "' AND AO.ruc='" & KEY_RUC & "' AND AD.ruc='" & KEY_RUC & "' AND T.id_doc='" & id_doc & "' AND T.serie='" & serie & "' AND T.numero='" & numero & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM movimiento_transferencia_detalle D,producto P,unidad U WHERE D.id_transferencia='" & rst("id_transferencia") & "' AND D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(110); rst("serie") + "-" + rst("numero")
    Printer.Print ""
    Printer.Print Tab(20); str(rst("fecha"))
    '-- TRANSPORTE
     If rst("id_transporte") <> "" Then
        transporte = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_transporte"))
    Else
        transporte = ""
    End If
    
    Printer.Print Tab(20); str(rst("fecha")) & Space(58) & transporte
    
   ' Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(10); Mid(UCase(rst("origen")) + Space(80), 1, 60) & Space(10) & rst("id_transporte")
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(10); Mid(UCase(rst("destino")) + Space(80), 1, 60)
   ' Printer.Print ""
     'Printer.CurrentY = Printer.CurrentY + 0.2
     Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
         
   ' Printer.Print Tab(8); Mid(rst("id_destinatario") + Space(80), 1, 20) + Space(80) + Mid(Licencia + Space(80), 1, 20)
   ' Printer.Print Tab(80); Mid(Chofer + Space(80), 1, 20)
    
    
 '   Printer.Print Tab(1); direccion_destino + Space(40) + rst("id_destinatario") 'L 5
    
    'If rst("id_chofer") <> "" Then
     '   Printer.Print Tab(1); BDBuscarCampo("persona", "licencia", "dni", rst("id_chofer")) + Space(40) + rst("id_transporte") 'L 7
  '  Else
   '     Printer.Print Tab(1); ""
   ' End If
    
    
   ' Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    Printer.Print ""
    ptotal = 0
     For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("cantidad") * rstT("peso")
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 6)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 10)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 65)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 8)
         Printer.Print Tab(0); codigo & Space(4) & descripcion & Space(8) & Und & Space(4) & cantidad & Space(4) & Mid(Format(rstT("peso"), "#,##0.00") + Space(10), 1, 6)
         rstT.MoveNext
    Next j
          
  
        inc = 0.01
         Do While (Val(Printer.CurrentY) <= 8.85)
          Printer.CurrentY = Printer.CurrentY + inc
        Loop
   
    
   
    'Printer.Print Tab(5); "*****************************************************************"
    'Printer.FontBold = True
    Printer.Print Tab(48); "***** PESO TOTAL *******"; Space(40) & Format(ptotal, "#,##0.00") + Space(2) + "Kg."
    
    'Printer.Print ""
    Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    If rst("id_chofer") <> "" Then
        Licencia = BDBuscarCampo("persona", "licencia", "dni", rst("id_chofer"))
        Chofer = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_chofer"))
    Else
        Licencia = ""
        Chofer = ""
    End If
    Licencia = rst("licencia")
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.Print Tab(16); Mid(UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_destinatario"))) + Space(80), 1, 65) & Chofer
   ' Printer.CurrentY = Printer.CurrentY + 0.5
   
         'Printer.Print ""
     
    
    'Printer.CurrentY = Printer.CurrentY + 0.2
   
    
    
    If rst("id_venta") > 0 Then
        Dim strfactura() As String
        
        strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            strfactura = Split(rstL("documento"), ":")
            in_tipo_comprobante = strfactura(0)
            in_numero = strfactura(1)
            
        End If
   Else
        in_tipo_comprobante = Mid("-" + Space(40), 1, 10)
        in_numero = "-"
    End If
    
 
     
    Printer.Print Tab(10); Mid(rst("id_destinatario") + Space(80), 1, 25) & in_tipo_comprobante & Space(45) & Licencia
    Printer.Print Tab(40); in_numero & Space(45) & Mid(rst("marca_placa") + Space(80), 1, 12) & rst("placa") & Space(10) & rst("mtc")
   
     If rst("id_motivo") = 3 Then
        Printer.Print ""
        Printer.CurrentY = Printer.CurrentY + 0.3
        Printer.Print Tab(76); "X"
     Else
         If rst("id_venta") > 0 Then
            'Printer.Print ""
            Printer.Print Tab(37); "X"
            End If
    End If
    
   
    
   
    Printer.Print ""
    Printer.Print ""
   
    
    Printer.EndDoc
    Exit Sub
End Sub

Public Sub impresion_formato_guia_sr_montana(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
    Dim ptotal As Double, direccion_destino As String, transporte As String
    Dim in_ubigeuo1 As String
    Dim in_ubigueo2 As String
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "10"
    Printer.CurrentX = 0
    Printer.CurrentY = 0
  
    Printer.Print "" 'Tab(10); "7 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "8 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "9 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "10---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   
   'Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
   'Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    strCadena = "SELECT T.remitente,T.destinatario,T.placa,T.id_transferencia,T.certificado,T.fecha,AO.direccion as origen,T.direccion as destino,marca_placa,id_transporte,id_direccion,id_chofer,T.observacion,id_remitente,id_destinatario,id_motivo,serie,numero,T.mtc,T.id_venta,T.licencia,T.peso_total,T.fecha_traslado FROM movimiento_transferencia T,almacen AO,almacen AD WHERE T.id_alm_origen=AO.id_alm AND T.id_alm_destino=AD.id_alm AND T.ruc='" & KEY_RUC & "' AND AO.ruc='" & KEY_RUC & "' AND AD.ruc='" & KEY_RUC & "' AND T.id_doc='" & id_doc & "' AND T.serie='" & serie & "' AND T.numero='" & numero & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM view_transferencia_detalle WHERE id_transferencia='" & rst("id_transferencia") & "' and ruc='" & KEY_RUC & "' "
    Call ConfiguraRstT(strCadena)
    Printer.Print ""
    Printer.Print Tab(100); rst("serie") + "-" + rst("numero")
    Printer.Print ""
    Printer.Print Tab(25); str(rst("fecha")) + Space(70) + Format(rst("fecha_traslado"), "dd-mm-YYYY")
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(0); Mid(UCase(rst("origen")) + Space(100), 1, 60) & Space(10) & Mid(UCase(rst("destino")) + Space(100), 1, 80)
    'Printer.Print ""
    
    in_ubigeuo1 = Mid(UCase(get_ubigueo_persona(rst("id_remitente"), 0)) & Space(100), 1, 80)
    in_ubigeuo2 = Mid(UCase(get_ubigueo_persona(rst("id_destinatario"), rst("id_direccion"))) & Space(100), 1, 60)
    
    Printer.Print Tab(10); in_ubigeuo1 & Space(10) & in_ubigeuo2
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(25); Mid(UCase(rst("remitente")) + Space(80), 1, 60) & Space(20) & Mid(UCase(rst("destinatario")) + Space(80), 1, 70)
    Printer.Print Tab(40); Mid(rst("id_remitente") + Space(80), 1, 60) & Space(10) & Mid(UCase(rst("id_destinatario")) + Space(80), 1, 70)
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    ptotal = 0
    For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("total")
         codigo = Mid(Format(j + 1, "000") + Space(10), 1, 6)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 10)
         descripcion = Mid(rstT("detalle") + Space(80), 1, 75)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 8)
         Printer.Print Tab(5); codigo & Space(4) & cantidad & Space(4) & descripcion & Space(4) & Mid(Format(rstT("total"), "#,##0.00") + Space(10), 1, 6) & Space(3) & Und
         rstT.MoveNext
    Next j
          inc = 0.5
    
    Do While (Val(Printer.CurrentY) <= 30)
       Printer.CurrentY = Printer.CurrentY + inc
    Loop
    Printer.Print Tab(30); ":::::::: PESO TOTAL :::::::::" & Space(40) & Format(rst("peso_total"), "#,##0.00") + Space(2) + "Kg."
    Printer.Print ""
    Printer.Print Tab(30); ":::::::: REFERENCIA ::::::::: " & Space(2) & rst("observacion")
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
     Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(30); Mid(rst("marca_placa") + Space(80), 1, 20) & Space(5) & rst("placa")
     Printer.Print Tab(40); Mid(rst("mtc") + Space(80), 1, 20)
    
    If rst("id_chofer") <> "" Then
        Licencia = BDBuscarCampo("persona", "licencia", "dni", rst("id_chofer"))
        Chofer = get_persona(rst("id_chofer"))
    Else
        Licencia = ""
        Chofer = ""
    End If
    Printer.Print ""
    If rst("id_transporte") = KEY_RUC Then
        in_contratado = ""
        in_transporte = ""
        Chofer = ""
    Else
        in_contratado = Mid(get_persona(rst("id_transporte")) + Space(80), 1, 60)
        in_transporte = rst("id_transporte")
        Chofer = get_persona(rst("id_chofer"))
    End If
    Printer.Print Tab(40); Mid(rst("certificado") + Space(80), 1, 60) & Space(10) & in_contratado
    Licencia = rst("licencia")
    Printer.Print Tab(40); Mid(Licencia + Space(80), 1, 60) & Space(10) & in_transporte
    Printer.Print Tab(110); Chofer
       
    
    
    
    
    
    Printer.EndDoc
    Exit Sub
End Sub





Private Sub impresion_formato_2_guia(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
' Printer.PaperSize = 1
 Printer.TrackDefault = True 'siempre apunta a la impresora predeter
   Printer.CurrentX = 0
   Printer.CurrentY = 0
    'Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    'Printer.ScaleWidth = 22.9
    'Printer.ScaleHeight = 14
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    
    strCadena = "SELECT T.id_transferencia,T.fecha,AO.direccion as origen,AD.direccion as destino,marca_placa,id_transporte,id_chofer,id_destinatario,id_motivo,serie,numero FROM movimiento_transferencia T,almacen AO,almacen AD WHERE T.id_alm_origen=AO.id_alm AND T.id_alm_destino=AD.id_alm AND T.ruc='" & KEY_RUC & "' AND AO.ruc='" & KEY_RUC & "' AND AD.ruc='" & KEY_RUC & "' AND T.id_doc='" & id_doc & "' AND T.serie='" & serie & "' AND T.numero='" & numero & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT * FROM view_transferencia_detalle WHERE id_transferencia='" & Val(rst("id_transferencia")) & "'"
    Call ConfiguraRstT(strCadena)
    
    Printer.Print "" 'Tab(10); 'L 1
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(100); rst("serie") + "-" + rst("numero")
    
    Printer.Print Tab(18); rst("fecha") 'L 3
    Printer.Print "" 'Tab(10); 'L 2
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(5); Mid(UCase(rst("origen")) + Space(80), 1, 70)
    If rst("id_motivo") = 3 Then
        direccion_destino = UCase(rst("destino"))
        Destinatario = KEY_EMPRESA
    Else
        strCadena = "SELECT nombre_completo,direccion FROM persona WHERE dni='" & rst("id_destinatario") & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
            direccion_destino = rstZ("direccion")
            Destinatario = rstZ("nombre_completo")
        Else
            direccion_destino = ""
            Destinatario = ""
        End If
        direccion_destino = UCase(direccion_destino)
    End If
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(5); Mid(Destinatario + Space(80), 1, 105) + rst("id_destinatario")  'L 5
    Printer.Print Tab(5); Mid(direccion_destino + Space(80), 1, 105) + rst("id_destinatario")  'L 5
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 8
    Printer.CurrentY = Printer.CurrentY + 0.5
    For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("total")
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 8)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 10)
         descripcion = Mid(rstT("detalle") + Space(80), 1, 85)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 13)
         Printer.Print Tab(5); cantidad & Space(2) & Und & Space(2); descripcion & Space(2) & codigo & Format(rstT("total"), "#,##0.00")
         strCadena = "SELECT * FROM movimiento_transferencia_series   WHERE  id_producto='" & rstT("id_producto") & "' and  id_transferencia='" & rst("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
         Call ConfiguraRstZ(strCadena)
         If rstZ.RecordCount > 0 Then
            rstZ.MoveFirst
            Printer.Print Tab(25); "**************************"
            Printer.Print Tab(25); "COLOR           :" & Space(2) & rstT("color")
            Printer.Print Tab(25); "MODELO          :" & Space(2) & rstT("modelo")
            Printer.Print "" 'Tab(10); 'L 8
                
            For m = 0 To rstZ.RecordCount - 1
                Printer.Print Tab(25); "ITEM   :" & Space(2) & str(m + 1)
                Printer.Print Tab(30); "Nº MOTOR         :" & Space(2) & rstZ("motor")
                Printer.Print Tab(30); "N°CHASIS         :" & Space(2) & rstZ("chasis")
                Printer.Print Tab(30); "AÑO              :" & Space(2) & rstZ("anio_fabricacion")
                rstZ.MoveNext
            Next m
         End If

         rstT.MoveNext
    Next j
          inc = 0.5
          
          Do While (Val(Printer.CurrentY) <= 17)
             Printer.CurrentY = Printer.CurrentY + inc
          Loop
    Printer.FontBold = True
    Printer.Print Tab(40); "***** PESO TOTAL *******"; Space(5) & Format(ptotal, "#,##0.00") + Space(2) + "Kg."
    
    
    strCadena = "SELECT nombre_completo FROM persona where dni='" & rst("id_transporte") & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        transporte = rstZ("nombre_completo")
    Else
        transporte = ""
    End If
    Printer.Print Tab(20); Mid(rst("marca_placa") + Space(80), 1, 70) + transporte
    strCadena = "SELECT licencia FROM persona where dni='" & rst("id_chofer") & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        Chofer = rstZ("nombre_completo")
    Else
        Chofer = ""
    End If
    
    
    Printer.Print Tab(23); Mid(Chofer + Space(80), 1, 95) + rst("id_transporte") 'L 7
    Printer.EndDoc
    Exit Sub
   
End Sub
Private Sub impresion_formato_3_guia_tienda(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
     
     
    'Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "10"
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    'Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    'Printer.ScaleWidth = 22.9
    'Printer.ScaleHeight = 14
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT T.id_transferencia,T.fecha,AO.direccion as origen,AD.direccion as destino,marca_placa,id_transporte,id_chofer,id_destinatario,id_motivo,serie,numero,id_venta FROM movimiento_transferencia T,almacen AO,almacen AD WHERE T.id_alm_origen=AO.id_alm AND T.id_alm_destino=AD.id_alm AND T.ruc='" & KEY_RUC & "' AND AO.ruc='" & KEY_RUC & "' AND AD.ruc='" & KEY_RUC & "' AND T.id_doc='" & id_doc & "' AND T.serie='" & serie & "' AND T.numero='" & numero & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    
    strCadena = "SELECT D.id_producto,D.total,U.abreviatura,P.nombre_prod,D.cantidad,C.descripcion as color,S.descripcion as modelo FROM movimiento_transferencia_detalle D,producto P,unidad U,imp_color C,linea_sub S WHERE  P.id_color=C.id_color and P.id_sublinea=S.id_tipo and P.ruc=S.id_usu and   D.id_transferencia='" & rst("id_transferencia") & "' AND D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
    Printer.Print "" 'Tab(10); 'L 1
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(90); rst("serie") + "-" + rst("numero")
    'Printer.Print "" 'Tab(10); 'L 2
    Printer.CurrentY = Printer.CurrentY + 0.5
   ' Printer.Print "" 'Tab(10); 'L 2
   If rst("id_venta") > 0 Then
   strCadena = "SELECT documento FROM movimiento_venta WHERE id_venta='" & rst("id_venta") & "' "
   Call ConfiguraRstZ(strCadena)
   If rst.RecordCount > 0 Then
        nboleta = rstZ("documento")
    Else
        nboleta = ""
   End If
   Else
    nboleta = ""
   End If
    Printer.Print Tab(35); rst("fecha") & Space(70) & nboleta
    'Printer.Print "" 'Tab(10); 'L 2
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(45); Mid(UCase(rst("origen")) + Space(80), 1, 70)
    
    
    
    If rst("id_motivo") = 3 Then
        direccion_destino = UCase(rst("destino"))
        Destinatario = KEY_EMPRESA
    Else
        
        
        strCadena = "SELECT nombre_completo,direccion,id_departamento,id_distrito,id_provincia FROM persona WHERE dni='" & rst("id_destinatario") & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
            If rstZ("id_departamento") <> "0" Then
                strCadena = "select * from departamentos WHERE id_depa='" & rstZ("id_departamento") & "'"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    ndepartamento = "- " & UCase(rstL("descripcion"))
                Else
                    ndepartamento = " "
                End If
                
                If rstZ("id_provincia") <> "0" Then
                strCadena = "select * from provincia WHERE id_provincia='" & rstZ("id_provincia") & "'"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    nprovincia = "- " & UCase(rstL("descripcion"))
                Else
                    nprovincia = " "
                End If
                End If
                
                
            If rstZ("id_distrito") <> "0" Then
                strCadena = "select * from distrito WHERE id_distrito='" & rstZ("id_distrito") & "'"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    ndistrito = "- " & UCase(rstL("descripcion"))
                Else
                    ndistrito = " "
                End If
                End If
            End If
            direccion_destino = rstZ("direccion") & Space(1) & ndistrito & Space(1) & nprovincia & Space(1) & ndepartamento
            Destinatario = rstZ("nombre_completo")
        Else
            direccion_destino = ""
            Destinatario = ""
        End If
        direccion_destino = UCase(direccion_destino)
    End If
    
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(35); Mid(Destinatario + Space(80), 1, 90) + rst("id_destinatario")  'L 5
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(35); Mid(direccion_destino + Space(80), 1, 105)
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 8
    'Printer.CurrentY = Printer.CurrentY + 0.5
    For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("total")
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 12)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 10)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 70)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 13)
         Printer.Print Tab(14); cantidad & Space(2) & Und & Space(2); descripcion & Space(2) & codigo & Format(rstT("total"), "#,##0.00")
         strCadena = "SELECT * FROM movimiento_transferencia_series   WHERE  id_producto='" & rstT("id_producto") & "' and  id_transferencia='" & rst("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
         Call ConfiguraRstZ(strCadena)
         If rstZ.RecordCount > 0 Then
            rstZ.MoveFirst
            Printer.Print Tab(40); "**************************"
            Printer.Print Tab(40); "COLOR           :" & Space(2) & rstT("color")
            Printer.Print Tab(40); "MODELO          :" & Space(2) & rstT("modelo")
            Printer.Print "" 'Tab(10); 'L 8
                
            For m = 0 To rstZ.RecordCount - 1
                Printer.Print Tab(40); "ITEM   :" & Space(2) & str(m + 1)
                Printer.Print Tab(45); "MARCA            :" & Space(2) & get_marca(rstZ("id_producto"))
                Printer.Print Tab(45); "Nº MOTOR         :" & Space(2) & rstZ("motor")
                Printer.Print Tab(45); "N°CHASIS         :" & Space(2) & rstZ("chasis")
                Printer.Print Tab(45); "AÑO              :" & Space(2) & rstZ("anio_fabricacion")
                Printer.Print Tab(45); "NRO DUA          :" & Space(2) & rstZ("nro_dua")
                Printer.Print Tab(45); "NRO ITEM         :" & Space(2) & rstZ("nro_item")
                rstZ.MoveNext
            Next m
         End If

         rstT.MoveNext
    Next j
          inc = 0.5
          
          Do While (Val(Printer.CurrentY) <= 17)
             Printer.CurrentY = Printer.CurrentY + inc
          Loop
    'Printer.FontBold = True
  '  Printer.Print Tab(40); "***** PESO TOTAL *******"; Space(5) & Format(ptotal, "#,##0.00") + Space(2) + "Kg."
    
    
    strCadena = "SELECT nombre_completo FROM persona where dni='" & rst("id_transporte") & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        transporte = rstZ("nombre_completo")
    Else
        transporte = ""
    End If
    Printer.Print Tab(20); Mid(rst("marca_placa") + Space(80), 1, 70) + transporte
    strCadena = "SELECT licencia FROM persona where dni='" & rst("id_chofer") & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        Chofer = rstZ("nombre_completo")
    Else
        Chofer = ""
    End If
    
    
    'Printer.Print Tab(23); Mid(Chofer + Space(80), 1, 95) + rst("id_transporte") 'L 7
    Printer.EndDoc
    Exit Sub
   
End Sub
Public Sub impresion_manifiesto(ByVal in_manifiesto As String)

     
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "10"
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM `transferencia_manifiesto` m,view_manifiesto_comprobantes c,persona p Where m.id_manifiesto='" & Val(in_manifiesto) & "' and m.dni_chofer=p.dni and  m.`id_manifiesto`=c.`id_manifiesto` and m.`ruc` =c.`ruc` LIMIT 1 "
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    Printer.Print "" 'Tab(10); 'L 1
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(40); "MANIFIESTO DE CARGA"
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(1); "MANIFIESTO DE CARGA Nª :" & rst("id_anio") & "-" & rst("id_numero") & Space(20) & "FECHA :" & Format(rst("fecha"), "dd-mm-YYYY")
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print Tab(1); "CHOFER  :" & Mid(rst("nombre_completo") & Space(50), 1, 50) & Space(5) & "PROPIETARIO  :" & get_persona(rst("ruc_propietario"))
    Printer.Print Tab(1); "BREVETE :" & Mid(rst("licencia") & Space(50), 1, 50) & Space(5) & "DIRECCION    :" & rst("direccion")
    Printer.Print Tab(1); "PLACA   :" & Mid(rst("placa") & Space(50), 1, 50) & Space(5) & "R.U.C        :" & rst("ruc_propietario")
    Printer.Print Tab(1); "------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print Tab(1); "DESTINO       FACTURA/BOLETA     GUIA.R    DOC.CLI   REMITENTE              DESTINATARIO           BULTO            PESO (KG)        FLETE"
    Printer.Print Tab(1); "------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    rst.MoveFirst
    t_bultos = 0
    t_peso = 0
    t_flete = 0
    For j = 0 To rst.RecordCount - 1
        in_distrito = Mid(UCase(rst("distrito")) & Space(20), 1, 14)
        in_comprobante = Mid(rst("documento") & Space(20), 1, 20)
        in_guia = Mid(rst("guia") & Space(20), 1, 18)
        in_doc_cli = ""
        in_remitente = Mid(rst("remitente") & Space(30), 1, 25)
        in_destinatario = Mid(rst("destinatario") & Space(30), 1, 25)
        in_bulto = Mid(rst("cantidad") & Space(20), 1, 6)
        in_peso = Mid(Format(rst("peso_total"), "#,##0.00") & Space(10), 1, 10)
        in_flete = Mid(Format(rst("total"), "#,##0.00") & Space(10), 1, 10)
        t_bultos = t_bultos + rst("cantidad")
        t_peso = t_peso + rst("peso_total")
        If IsNull(rst("total")) = False Then
            t_flete = t_flete + rst("total")
        End If
        Printer.Print Tab(1); in_distrito & in_comprobante & in_guia & in_doc_cli & in_remitente & Space(2) & in_destinatario & Space(2) & in_bulto & in_peso & in_flete
        rst.MoveNext
    Next j
    Printer.Print Tab(1); "------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print Tab(1); "                                                                                                     " & Mid(Format(t_bultos, "#,##0.00") & Space(10), 1, 10) & Mid(Format(t_peso, "#,##0.00") & Space(10), 1, 15) & Mid(Format(t_flete, "#,##0.00") & Space(10), 1, 10)
    Printer.Print Tab(1); "------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(1); "RESUMEN"
    Printer.Print ""
    Printer.Print Tab(20); "-----------------------------------------------            -----------------------------------------------------------"
    Printer.Print Tab(20); "   NªGUIAS    Nª BULTOS    PESO KG                          CONT.OFICINA  VUELTA GUIA  POR COBRAR   TOTAL "
    Printer.Print Tab(20); "-----------------------------------------------            -----------------------------------------------------------"
    strCadena = "SELECT distrito,COUNT(*) as cantidad,sum(c.`cantidad`) as bultos,sum(c.`peso_total`)as peso,c.total FROM `transferencia_manifiesto` m,view_manifiesto_comprobantes c WHERE m.id_manifiesto='" & in_manifiesto & "'  and m.`id_manifiesto`=c.`id_manifiesto` and m.`ruc` =c.`ruc` GROUP by c.`distrito`"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       
        ibultos = 0
        iguias = 0
        ipeso = 0
        itotal = 0
       For i = 0 To rst.RecordCount - 1
            ibultos = ibultos + rst("bultos")
            iguias = iguias + rst("cantidad")
            ipeso = ipeso + rst("peso")
            itotal = itotal + rst("total")
            in_distrito = Mid(UCase(rst("distrito")) & Space(20), 1, 20)
            in_guias = Mid(rst("cantidad") & Space(15), 1, 20)
            in_bultos = Mid(rst("bultos") & Space(15), 1, 15)
            in_peso = Mid(Format(rst("peso"), "#,##0.00") & Space(15), 1, 15)
            in_monto1 = Mid("0.00" & Space(20), 1, 12)
            in_total = Mid(Format(rst("total"), "#,##0.00") & Space(10), 1, 10)
            Printer.Print Tab(1); in_distrito & in_guias & in_bultos & in_peso & Space(20) & in_monto1 & in_monto1 & in_total & in_total
            rst.MoveNext
       Next i
           Printer.Print Tab(20); "-----------------------------------------------            -----------------------------------------------------------"
       Printer.Print ""
       Printer.Print Tab(20); Mid(iguias & Space(10), 1, 20) & Mid(str(ibultos) & Space(10), 1, 15) & Mid(Format(ipeso, "#,##0.00") & Space(40), 1, 35) & in_monto1 & in_monto1 & Mid(Format(itotal, "#,##0.00") & Space(10), 1, 12) & Mid(Format(itotal, "#,##0.00") & Space(10), 1, 12)
    End If
    Printer.Print ""
    Printer.Print Tab(50); "COMISION 15%                   " & in_monto1 & in_monto1 & in_monto1 & Format(itotal * 15 / 100, "#,##0.00")
    Printer.Print ""
    
    Printer.EndDoc
    Exit Sub
   
End Sub

Private Sub impresion_formato_3_guia_tienda_rep(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Dim ptotal As Double, direccion_destino As String
Dim direccion_per As String
Dim dni_atencion As String
Dim atencion As String
 'Printer.PaperSize = 1
  Printer.TrackDefault = True 'siempre apunta a la impresora predeter
     Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "10"

  
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    'Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    'Printer.ScaleWidth = 22.9
    'Printer.ScaleHeight = 14
     'MsgBox Printer.Height
     'MsgBox Printer.Width
    '
    '***** COMPROBANTE *****
    strCadena = "SELECT T.id_transferencia,T.fecha,AO.direccion as origen,T.direccion,AD.direccion as destino,marca_placa,id_transporte,id_chofer,id_destinatario,id_motivo,serie,numero,id_venta,T.dni_atencion,T.atencion FROM movimiento_transferencia T,almacen AO,almacen AD WHERE T.id_alm_origen=AO.id_alm AND T.id_alm_destino=AD.id_alm AND T.ruc='" & KEY_RUC & "' AND AO.ruc='" & KEY_RUC & "' AND AD.ruc='" & KEY_RUC & "' AND T.id_doc='" & id_doc & "' AND T.serie='" & serie & "' AND T.numero='" & numero & "'"
    Call ConfiguraRst(strCadena)
    direccion_per = rst("direccion")
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    strCadena = "SELECT D.id_producto,D.total,U.abreviatura,P.nombre_prod,D.cantidad,C.descripcion as color,S.descripcion as modelo FROM movimiento_transferencia_detalle D,producto P,unidad U,imp_color C,linea_sub S WHERE  P.id_color=C.id_color and P.id_sublinea=S.id_tipo and P.ruc=S.id_usu and   D.id_transferencia='" & rst("id_transferencia") & "' AND D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    
  
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 2
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(90); rst("serie") + "-" + rst("numero")
    'Printer.Print "" 'Tab(10); 'L 2
    Printer.CurrentY = Printer.CurrentY + 0.5
   ' Printer.Print "" 'Tab(10); 'L 2
   If rst("id_venta") > 0 Then
   strCadena = "SELECT C.doc_abrev,serie,numero FROM movimiento_venta M,comprobantes C WHERE M.id_doc=C.id_doc and   id_venta='" & rst("id_venta") & "' "
   Call ConfiguraRstZ(strCadena)
   
   If rst.RecordCount > 0 Then
        nboleta = Mid(rstZ("doc_abrev"), 1, 3) & ":" & rstZ("serie") & "-" & rstZ("numero")
   Else
        nboleta = ""
   End If
   Else
    nboleta = ""
   End If
    If rst("id_motivo") = 3 Then
        direccion_destino = UCase(rst("destino"))
        Destinatario = KEY_EMPRESA
    Else
        strCadena = "SELECT nombre_completo,direccion,id_departamento,id_distrito,id_provincia FROM persona WHERE dni='" & rst("id_destinatario") & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
            If rstZ("id_departamento") <> "0" Then
                strCadena = "select * from departamentos WHERE id_depa='" & rstZ("id_departamento") & "'"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    ndepartamento = "- " & UCase(rstL("descripcion"))
                Else
                    ndepartamento = " "
                End If
                
                If rstZ("id_provincia") <> "0" Then
                strCadena = "select * from provincia WHERE id_provincia='" & rstZ("id_provincia") & "'"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    nprovincia = "- " & UCase(rstL("descripcion"))
                Else
                    nprovincia = " "
                End If
                End If
                
                
            If rstZ("id_distrito") <> "0" Then
                strCadena = "select * from distrito WHERE id_distrito='" & rstZ("id_distrito") & "'"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    ndistrito = "- " & UCase(rstL("descripcion"))
                Else
                    ndistrito = " "
                End If
                End If
            End If
            If direccion_per = "-" Then
            
                direccion_destino = rstZ("direccion") & Space(1) & ndistrito & Space(1) & nprovincia & Space(1) & ndepartamento
            Else
                direccion_destino = direccion_per
            End If
            Destinatario = rstZ("nombre_completo")
        Else
            direccion_destino = ""
            Destinatario = ""
        End If
        direccion_destino = UCase(direccion_destino)
    End If
    
    Printer.Print Tab(20); Mid(Destinatario + Space(80), 1, 90)
    Printer.Print Tab(20); rst("fecha")
    Printer.Print "" 'Tab(10); 'L 2
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(20); Mid(UCase(rst("origen")) + Space(80), 1, 85) & Space(2) & rst("id_destinatario")
    Printer.Print Tab(20); Mid(direccion_destino + Space(80), 1, 83) & Space(2) & nboleta
    Printer.Print "" 'Tab(10); 'L 2
    Printer.Print "" 'Tab(10); 'L 8
    'Printer.CurrentY = Printer.CurrentY + 0.5
    For j = 0 To rstT.RecordCount - 1
         ptotal = ptotal + rstT("total")
         codigo = Mid((rstT("id_producto")) + Space(10), 1, 12)
         Und = Mid((rstT("abreviatura")) + Space(10), 1, 10)
         descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 70)
         cantidad = Mid(Format(str(rstT("cantidad")), "#,##0.00") + Space(10), 1, 8)
         Printer.Print Tab(12); cantidad & Space(2) & Und & Space(2); descripcion & Space(2) & "" & ""
         strCadena = "SELECT * FROM movimiento_transferencia_series   WHERE  id_producto='" & rstT("id_producto") & "' and  id_transferencia='" & rst("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
         Call ConfiguraRstZ(strCadena)
         If rstZ.RecordCount > 0 Then
            rstZ.MoveFirst
            Printer.Print Tab(40); "**************************"
            Printer.Print Tab(40); "MODELO          :" & Space(2) & rstT("modelo")
            Printer.Print "" 'Tab(10); 'L 8
                
            For m = 0 To rstZ.RecordCount - 1
                Printer.Print Tab(40); "ITEM   :" & Space(2) & str(m + 1)
                Printer.Print Tab(45); "MARCA            :" & Space(2) & get_marca(rstZ("id_producto"))
                Printer.Print Tab(45); "Nº MOTOR         :" & Space(2) & rstZ("motor")
                Printer.Print Tab(45); "AÑO              :" & Space(2) & rstZ("anio_fabricacion")
                Printer.Print Tab(45); "NRO DUA          :" & Space(2) & rstZ("nro_dua")
                Printer.Print Tab(45); "NRO ITEM         :" & Space(2) & rstZ("nro_item")
                rstZ.MoveNext
            Next m
         End If

         rstT.MoveNext
    Next j
          inc = 0.5
          
          Do While (Val(Printer.CurrentY) <= 15)
             Printer.CurrentY = Printer.CurrentY + inc
          Loop
    'Printer.FontBold = True
  '  Printer.Print Tab(40); "***** PESO TOTAL *******"; Space(5) & Format(ptotal, "#,##0.00") + Space(2) + "Kg."
    If Len(Trim(rst("atencion"))) > 1 Then
        Printer.Print Tab(20); "ATENCION  A:" & Space(2) & rst("atencion") & Space(2); "(" & rst("dni_atencion") & ")"
        
    End If
    strCadena = "SELECT nombre_completo FROM persona where dni='" & rst("id_transporte") & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        transporte = rstZ("nombre_completo")
    Else
        transporte = ""
    End If
    Printer.Print Tab(20); Mid(rst("marca_placa") + Space(80), 1, 70) + transporte
    strCadena = "SELECT licencia FROM persona where dni='" & rst("id_chofer") & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        Chofer = rstZ("nombre_completo")
    Else
        Chofer = ""
    End If
    
    
    'Printer.Print Tab(23); Mid(Chofer + Space(80), 1, 95) + rst("id_transporte") 'L 7
    Printer.EndDoc
    Exit Sub
   
End Sub

Private Sub impresion_formato_boleta_suelta(ByVal id_doc As String, ByVal serie As String, ByVal numero As String, Optional Direccion As String)
Dim descomp As String
Dim nvendedor As String
Dim in_moneda As String
Dim in_tipo_factura As String
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT * from comprobantes where id_doc='" & id_doc & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        descomp = Mid(rst("doc_des"), 1, 8)
    End If
    Printer.Print Tab(88); descomp; Space(2); Mid(serie + Space(50), 1, 4) & Space(1) & "-" & Space(1) & Trim(numero)
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    'Printer.Print ""
    in_moneda = rst("id_moneda")
    in_tipo_factura = rst("id_tipo_factura")
    nvendedor = rst("id_vendedor")
    Printer.CurrentY = Printer.CurrentY + 0.5
     Printer.Print Tab(18); Mid(Trim(rst("ncliente")) + Space(80), 1, 40) + Space(40) & rst("id_cliente")
     Printer.Print ""
    
    
    'Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT direccion FROM persona WHERE dni='" & Trim(rst("id_cliente")) & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
        Printer.Print Tab(18); Mid(rstL("direccion") + Space(80), 1, 40) + Space(40) + str(rst("fecha_emision"))
    Else
        Printer.Print Tab(18); Mid(KEY_DIR_PUBLIC + Space(80), 1, 40) + Space(40) + str(rst("fecha_emision"))
    End If
    
    'If Len(Trim(direccion)) > 0 Then
     '   Printer.Print Tab(18); Mid(Trim(direccion) + Space(80), 1, 40) + Space(30) + rst("id_cliente")
    'Else
      '  Printer.Print Tab(18); Mid(BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente")) + Space(80), 1, 40) + Space(38) + rst("id_cliente")
    'e 'nd If
   ' Printer.CurrentY = Printer.CurrentY + 1#
    Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT M.id_producto,M.cantidad,U.abreviatura,M.detalle,M.precio,M.total,ma.descripcion as marca FROM movimiento_venta_detalle M,producto P,unidad U,marca ma WHERE P.id_marca=ma.id_marca and ma.id_usu=P.ruc and   M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' ORDER BY M.id_detalle_venta ASC"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.5
            For j = 0 To rstT.RecordCount - 1
                codigo = Mid(rstT("id_producto") + Space(50), 1, 5)
               
               If rstT("id_producto") = KEY_COD_PER Then
                       If rstT("cantidad") = 0 Then
                            Und = ""
                            marca = ""
                            cantidad = Mid("" + Space(10), 1, 7)
                        Else
                            cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                            Und = rstT("abreviatura")
                            marca = Mid(rstT("marca") + Space(10), 1, 12)
                        End If
                        If rstT("precio") = 0 Then
                           precio = Mid(" " + Space(4), 1, 6)
                           totalPar = Mid("          " + Space(4), 1, 8)
                       Else
                            precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 6)
                            totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
                        End If
               Else
                            cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                            Und = rstT("abreviatura")
                            marca = Mid(rstT("marca") + Space(10), 1, 12)
                            precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 6)
                            totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
               End If
               
               
                
               
                descripcion = Mid(rstT("detalle") + Space(80), 1, 50)
                Printer.Print Tab(11); cantidad & Space(1) & Und & Space(2) & descripcion & Space(3) & marca & Space(7) & precio & Space(5) & totalPar
                rstT.MoveNext
            Next j
    End If
            inc = 0.5
            Do While (Val(Printer.CurrentY) <= 11)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
            Printer.Print ""
    totalletras = UCase(EnLetras(rst("total")))
    
    '---- fin totales
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(20); Mid(Trim(totalletras & Space(1) & get_moneda(in_moneda)) + Space(100), 1, 100)
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
   ' Printer.Print Tab(10); "N.DIRECCION FISCAL :" + KEY_DIRECCION
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print Tab(19); Mid(get_persona(nvendedor), 1, 15) & Space(1) & "a las:" & str(Now()) & Space(40) & "S/" & Format(rst("total"), "#,##0.00")
    
   ' Printer.Print ""
    Printer.EndDoc
    
    Exit Sub
End Sub

Private Sub impresion_orden_pago(ByVal id_doc As String, ByVal serie As String, ByVal numero As String, Optional Direccion As String)
Dim descomp As String
Dim nvendedor As String
Dim in_tipo_factura As String
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT * from comprobantes where id_doc='" & id_doc & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        descomp = Mid(rst("doc_des"), 1, 8)
    End If
    Printer.Font.name = "Tahoma"
    Printer.Font.Size = "15"
    Printer.Print Tab(30); "ORDEN DE PAGO :" & Mid(serie + Space(50), 1, 4) & Space(1) & "-" & Space(1) & Trim(numero)
        Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = "10"
    Printer.Print Tab(88); ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    'Printer.Print ""
    Printer.Print Tab(70); KEY_DIR_PUBLIC & ":" & Format(rst("fecha_emision"), "dd-mm-YYYY")
    in_tipo_factura = rst("id_tipo_factura")
    nvendedor = rst("id_vendedor")
     Printer.CurrentY = Printer.CurrentY + 0.5
     strCadena = "SELECT direccion FROM persona WHERE dni='" & Trim(rst("id_cliente")) & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
        ndireccion = Mid(rstL("direccion") + Space(80), 1, 40)
    Else
        ndireccion = KEY_DIR_PUBLIC
    End If
     Printer.Print Tab(10); "CANCELAR AL SR."; Trim(rst("ncliente")) & Space(1) & "IDENTIFICADO CON DNI: " & rst("id_cliente")
     Printer.Print Tab(10); "DOMICILIADO EN ." & ndireccion
     Printer.Print Tab(10); "LA CANTIDAD DE :" & "S/" & Format(rst("total"), "#,##0.00") & Space(2) & "NUEVOS SOLES."
     Printer.Print ""
     Printer.Print Tab(10); "POR CONCEPTO DE :"
    
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    'If Len(Trim(direccion)) > 0 Then
     '   Printer.Print Tab(18); Mid(Trim(direccion) + Space(80), 1, 40) + Space(30) + rst("id_cliente")
    'Else
      '  Printer.Print Tab(18); Mid(BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente")) + Space(80), 1, 40) + Space(38) + rst("id_cliente")
    'e 'nd If
   ' Printer.CurrentY = Printer.CurrentY + 1#
    Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT M.id_producto,M.cantidad,U.abreviatura,M.detalle,M.precio,M.total,ma.descripcion as marca FROM movimiento_venta_detalle M,producto P,unidad U,marca ma WHERE P.id_marca=ma.id_marca and ma.id_usu=P.ruc and   M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' ORDER BY M.id_detalle_venta ASC"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.5
            For j = 0 To rstT.RecordCount - 1
                codigo = Mid(rstT("id_producto") + Space(50), 1, 5)
               
               If rstT("id_producto") = KEY_COD_PER Then
                       If rstT("cantidad") = 0 Then
                            Und = ""
                            marca = ""
                            cantidad = Mid("" + Space(10), 1, 7)
                        Else
                            cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                            Und = rstT("abreviatura")
                            marca = Mid(rstT("marca") + Space(10), 1, 12)
                        End If
                        If rstT("precio") = 0 Then
                           precio = Mid(" " + Space(4), 1, 6)
                           totalPar = Mid("          " + Space(4), 1, 8)
                       Else
                            precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 6)
                            totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
                        End If
               Else
                            cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                            Und = rstT("abreviatura")
                            marca = Mid(rstT("marca") + Space(10), 1, 12)
                            precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 6)
                            totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
               End If
               
               
                
               
                descripcion = Mid(rstT("detalle") + Space(80), 1, 50)
                Printer.Print Tab(11); cantidad & Space(1) & Und & Space(2) & descripcion & Space(3) & marca & Space(7) & precio & Space(5) & totalPar
                rstT.MoveNext
            Next j
    End If
            inc = 0.5
            Do While (Val(Printer.CurrentY) <= 8)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
            Printer.Print ""
    totalletras = UCase(EnLetras(rst("total")))
    
    '---- fin totales
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(20); Mid(Trim(totalletras) + Space(100), 1, 100) & "S/" & Format(rst("total"), "#,##0.00")
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
   ' Printer.Print Tab(10); "N.DIRECCION FISCAL :" + KEY_DIRECCION
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print Tab(19); Mid("--------------------------------------------------------", 1, 50) & Space(20) & Mid("--------------------------------------------------------", 1, 50)
    Printer.Print Tab(19); Mid(get_persona(nvendedor) + Space(50), 1, 50) & Space(20) & Mid(Trim(rst("ncliente")) & Space(50), 1, 50)
    
   ' Printer.Print ""
    Printer.EndDoc
    
    Exit Sub
End Sub






Private Sub impresion_formato_recibo_suelta_per(ByVal id_doc As String, ByVal serie As String, ByVal numero As String, Optional Direccion As String)
Dim descomp As String
Dim nvendedor As String
Dim in_venta As Double
Dim in_tipo_factura As String
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    'Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT * from comprobantes where id_doc='" & id_doc & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        descomp = Mid(rst("doc_des"), 1, 8)
    End If
    Printer.Print Tab(88); descomp; Space(2); Mid(serie + Space(50), 1, 4) & Space(1) & "-" & Space(1) & Trim(numero)
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    'Printer.Print ""
    in_venta = rst("id_venta")
    in_tipo_factura = rst("id_tipo_factura")
    nvendedor = rst("id_vendedor")
    Printer.CurrentY = Printer.CurrentY + 0.5
     Printer.Print Tab(18); Mid(Trim(rst("ncliente")) + Space(80), 1, 40) + Space(40) & rst("id_cliente")
     Printer.Print ""
    
    
    'Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT direccion FROM persona WHERE dni='" & Trim(rst("id_cliente")) & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
        Printer.Print Tab(18); Mid(rstL("direccion") + Space(80), 1, 40) + Space(40) + str(rst("fecha_emision"))
    Else
        Printer.Print Tab(18); Mid(KEY_DIR_PUBLIC + Space(80), 1, 40) + Space(40) + str(rst("fecha_emision"))
    End If
    
    'If Len(Trim(direccion)) > 0 Then
     '   Printer.Print Tab(18); Mid(Trim(direccion) + Space(80), 1, 40) + Space(30) + rst("id_cliente")
    'Else
      '  Printer.Print Tab(18); Mid(BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente")) + Space(80), 1, 40) + Space(38) + rst("id_cliente")
    'e 'nd If
   ' Printer.CurrentY = Printer.CurrentY + 1#
    Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT M.id_producto,M.cantidad,U.abreviatura,M.detalle,M.precio,M.total FROM movimiento_venta_detalle M,producto P,unidad U WHERE   M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' ORDER BY M.id_detalle_venta ASC"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.5
            For j = 0 To rstT.RecordCount - 1
                codigo = Mid(rstT("id_producto") + Space(50), 1, 5)
               
               If rstT("id_producto") = KEY_COD_PER Then
                       If rstT("cantidad") = 0 Then
                            Und = ""
                            marca = ""
                            cantidad = Mid("" + Space(10), 1, 7)
                        Else
                            cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                            Und = rstT("abreviatura")
                            marca = Mid(" " + Space(12), 1, 12)
                        End If
                        If rstT("precio") = 0 Then
                           precio = Mid(" " + Space(4), 1, 6)
                           totalPar = Mid("          " + Space(4), 1, 8)
                       Else
                            precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 6)
                            totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
                        End If
               Else
                            cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                            Und = rstT("abreviatura")
                            marca = Mid("" + Space(12), 1, 12)
                            precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 6)
                            totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
               End If
               
               
                
               
                descripcion = Mid(rstT("detalle") + Space(80), 1, 50)
                Printer.Print Tab(11); cantidad & Space(1) & Und & Space(2) & descripcion & Space(3) & marca & Space(7) & precio & Space(5) & totalPar
                rstT.MoveNext
            Next j
    End If
            inc = 0.5
            Do While (Val(Printer.CurrentY) <= 11)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
            Printer.Print ""
            
            
    strCadena = "SELECT sum(monto_caja) FROM movimiento_venta_monto m    WHERE id_venta='" & in_venta & "' and `id_forma_pago`='01' "
    Call ConfiguraRstZ(strCadena)
    If rst.RecordCount > 0 Then
        
            totalletras = UCase(EnLetras(rstZ(0)))
            Printer.Print Tab(15); Mid(Trim(totalletras) + Space(100), 1, 100); Space(8) & "S/" & Format(rst("total"), "#,##0.00")
            Printer.CurrentY = Printer.CurrentY + 0.5
            Printer.Print Tab(123); "S/" & Format(rstZ(0), "#,##0.00")
            Printer.CurrentY = Printer.CurrentY + 0.5
            Printer.Print Tab(123); "S/" & Format(rst("total") - rstZ(0), "#,##0.00")
        
    End If
    
    'Printer.CurrentY = Printer.CurrentY + 0.2
    'Printer.Print Tab(19); Mid(get_persona(nvendedor), 1, 15) & Space(1) & "a las:" & Str(Now()) & Space(40) & "S/." & Format(rst("total"), "#,##0.00")
    
   ' Printer.Print ""
    Printer.EndDoc
    
    Exit Sub
End Sub



Private Sub impresion_formato_boleta_suelta_serie(ByVal id_doc As String, ByVal serie As String, ByVal numero As String, Optional Direccion As String)
Dim in_tipo_factura As String
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
   ' Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0
    Printer.Print Tab(90); "BOLETA"; Space(2); Mid(serie + Space(50), 1, 4) & Space(2) & "-" & Space(2) & Trim(numero)
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT * FROM movimiento_venta V,forma_pago_detalle F WHERE V.id_forma_pago=F.id_detalle AND V.ruc=F.ruc AND  id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND V.ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    in_tipo_factura = rst("id_tipo_factura")
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    If rst("conyugue") = "si" Then
         strCadena = "SELECT p.nombre_completo,a.dni_familia FROM persona_accidentes a,persona p WHERE a.dni_familia=p.dni and  a.dni='" & rst("id_cliente") & "' LIMIT 1 "
         Call ConfiguraRstZ(strCadena)
         If rstZ.RecordCount > 0 Then
             Printer.CurrentY = Printer.CurrentY + 0.5
             Printer.Print Tab(22); Mid(Trim(rstZ("nombre_completo")) + Space(80), 1, 40) + Space(58) & str(rstZ("dni_familia"))
         End If
    Else
        Printer.Print ""
        Printer.CurrentY = Printer.CurrentY + 0.5
    End If
    
    
    Printer.Print Tab(22); Mid(Trim(rst("ncliente")) + Space(80), 1, 40) + Space(58) & Trim(rst("id_cliente"))
    Printer.Print ""
    
    
    'Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT direccion,id_departamento,id_provincia,id_distrito FROM  persona WHERE dni='" & Trim(rst("id_cliente")) & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
                    If rstL("id_departamento") <> "0" Then
                            strCadena = "select * from departamentos WHERE id_depa='" & rstL("id_departamento") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndepartamento = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndepartamento = " "
                            End If
                            
                            If rstL("id_provincia") <> "0" Then
                            strCadena = "select * from provincia WHERE id_provincia='" & rstL("id_provincia") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                nprovincia = "- " & UCase(rstZ("descripcion"))
                            Else
                                nprovincia = " "
                            End If
                            End If
                            
                            
                        If rstL("id_distrito") <> "0" Then
                            strCadena = "select * from distrito WHERE id_distrito='" & rstL("id_distrito") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndistrito = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndistrito = " "
                            End If
                    End If
                    End If
          ndireccion = rstL("direccion") & ndistrito & nprovincia & ndepartamento
                    
        Printer.Print Tab(22); Mid(ndireccion + Space(80), 1, 90) + Space(8) + str(rst("fecha_emision"))
    Else
        Printer.Print Tab(22); Mid("---" + Space(80), 1, 40) + Space(58) + Trim(rst("id_cliente"))
    End If
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(22); Mid(rst("descripcion") + Space(80), 1, 40)
        
    
    'strcadena="SELECT * FROM movimiento_venta V,forma_pago_detalle F WHERE  V.id_forma_pago=F.id_forma_pago WHERE V.id_venta='"&  &"' "
    
    
    'If Len(Trim(direccion)) > 0 Then
     '   Printer.Print Tab(18); Mid(Trim(direccion) + Space(80), 1, 40) + Space(30) + rst("id_cliente")
    'Else
      '  Printer.Print Tab(18); Mid(BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente")) + Space(80), 1, 40) + Space(38) + rst("id_cliente")
    'e 'nd If
   ' Printer.CurrentY = Printer.CurrentY + 1#
    Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT C.descripcion as color, U.descripcion as abreviatura,P.id_linea,P.id_producto,P.nombre_prod,M.cantidad,M.precio,M.total,R.descripcion as marca,M.serie,M.nro_chasis,M.anio_modelo,M.anio_fabricacion,M.nro_dua,M.nro_item,ss.descripcion as modelo FROM movimiento_venta_detalle M,producto P,unidad U,marca R,imp_color C,linea_sub ss WHERE P.id_sublinea = ss.id_tipo and ss.id_usu = P.ruc and    P.id_color=C.id_color AND   P.id_marca=R.id_marca AND P.ruc=R.id_usu AND    M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.5
            For j = 0 To rstT.RecordCount - 1
                codigo = Mid(rstT("id_producto") + Space(50), 1, 5)
                cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                Und = rstT("abreviatura")
                descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 85)
                precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 8)
                totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
                Printer.Print Tab(12); cantidad & Space(4) & descripcion & precio & Space(6) & totalPar
                Printer.Print ""
                Printer.Print ""
                Printer.FontBold = True
                
                
                If rstT("id_linea") <> "00047" Then
                    Printer.Print Tab(22); "MARCA           :" & Space(2) & rstT("marca")
                    Printer.Print Tab(22); "N°CHASIS        :" & Space(2) & rstT("nro_chasis")
                    Printer.Print Tab(22); "MOTOR           :" & Space(2) & rstT("serie")
                    Printer.Print Tab(22); "MODELO          :" & Space(2) & rstT("modelo")
                    Printer.Print Tab(22); "COLOR           :" & Space(2) & rstT("color")
                    Printer.Print Tab(22); "AÑO FABRICACION :" & Space(2) & rstT("anio_fabricacion")
                    Printer.Print Tab(22); "NRO DUA         :" & Space(2) & rstT("nro_dua")
                    Printer.Print Tab(22); "NRO ITEM        :" & Space(2) & rstT("nro_item")
                    Printer.FontBold = True
                Else
                    Printer.Print Tab(22); "MARCA           :" & Space(2) & rstT("marca")
                    Printer.Print Tab(22); "MOTOR           :" & Space(2) & rstT("serie")
                    Printer.Print Tab(22); "MODELO          :" & Space(2) & rstT("anio_modelo")
                    Printer.Print Tab(22); "AÑO FABRICACION :" & Space(2) & rstT("anio_fabricacion")
                    Printer.Print Tab(22); "NRO DUA         :" & Space(2) & rstT("nro_dua")
                    Printer.Print Tab(22); "NRO ITEM        :" & Space(2) & rstT("nro_item")
                    Printer.FontBold = True
                End If
                rstT.MoveNext
                'Printer.CurrentY = Printer.CurrentY + 0.2
                
            Next j
    End If
            inc = 0.5
            Do While (Val(Printer.CurrentY) <= 13)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    totalletras = UCase(EnLetras(rst("total")))
    
    '---- fin totales
    'Printer.CurrentY = Printer.CurrentY + 0.8
    Printer.Print Tab(15); Mid(Trim(totalletras) + Space(100), 1, 100)
    Printer.Print ""
   ' Printer.Print Tab(10); "N.DIRECCION FISCAL :" + KEY_DIRECCION
    Printer.CurrentY = Printer.CurrentY + 0.5
   ' Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(118); "S/" & Format(rst("total"), "#,##0.00")
    
    Printer.Print ""
    Printer.EndDoc
    
    Exit Sub
End Sub
Private Sub impresion_formato_boleta_suelta_serie_rep(ByVal id_doc As String, ByVal serie As String, ByVal numero As String, Optional Direccion As String)
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
   ' Printer.Print ""
    Printer.Print ""
    
   ' Printer.CurrentY = Printer.CurrentY + 0
    Printer.Print Tab(90); "BOLETA"; Space(2); Mid(serie + Space(50), 1, 4) & Space(2) & "-" & Space(2) & Trim(numero)
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT * FROM movimiento_venta V,forma_pago_detalle F WHERE V.id_forma_pago=F.id_detalle AND V.ruc=F.ruc AND  id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND V.ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    If rst("conyugue") = "si" Then
         strCadena = "SELECT p.nombre_completo,a.dni_familia FROM persona_accidentes a,persona p WHERE a.dni_familia=p.dni and  a.dni='" & rst("id_cliente") & "' LIMIT 1 "
         Call ConfiguraRstZ(strCadena)
         If rstZ.RecordCount > 0 Then
         '    Printer.CurrentY = Printer.CurrentY + 0.5
             Printer.Print Tab(22); Mid(Trim(rstZ("nombre_completo")) + Space(80), 1, 40) + Space(58) & str(rstZ("dni_familia"))
         End If
    Else
    '    Printer.Print ""
        Printer.CurrentY = Printer.CurrentY + 0.5
    End If
    
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(22); Mid(Trim(rst("ncliente")) + Space(80), 1, 40) + Space(40) & Trim(rst("id_cliente"))
    'Printer.Print ""
    
    
    'Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT direccion,id_departamento,id_provincia,id_distrito FROM  persona WHERE dni='" & Trim(rst("id_cliente")) & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
                    If rstL("id_departamento") <> "0" Then
                            strCadena = "select * from departamentos WHERE id_depa='" & rstL("id_departamento") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndepartamento = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndepartamento = " "
                            End If
                            
                            If rstL("id_provincia") <> "0" Then
                            strCadena = "select * from provincia WHERE id_provincia='" & rstL("id_provincia") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                nprovincia = "- " & UCase(rstZ("descripcion"))
                            Else
                                nprovincia = " "
                            End If
                            End If
                            
                            
                        If rstL("id_distrito") <> "0" Then
                            strCadena = "select * from distrito WHERE id_distrito='" & rstL("id_distrito") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndistrito = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndistrito = " "
                            End If
                    End If
                    End If
          ndireccion = rstL("direccion") & ndistrito & nprovincia & ndepartamento
        Printer.CurrentY = Printer.CurrentY + 0.1
        Printer.Print Tab(22); Mid(ndireccion + Space(80), 1, 70) + Space(10) + str(rst("fecha_emision"))
    Else
        Printer.Print Tab(22); Mid(KEY_DIR_PUBLIC + Space(80), 1, 70) + Space(10) + str(rst("fecha_emision"))
    End If
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.Print Tab(22); Mid(rst("descripcion") + Space(80), 1, 40)
        
    
    'strcadena="SELECT * FROM movimiento_venta V,forma_pago_detalle F WHERE  V.id_forma_pago=F.id_forma_pago WHERE V.id_venta='"&  &"' "
    
    
    'If Len(Trim(direccion)) > 0 Then
     '   Printer.Print Tab(18); Mid(Trim(direccion) + Space(80), 1, 40) + Space(30) + rst("id_cliente")
    'Else
      '  Printer.Print Tab(18); Mid(BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente")) + Space(80), 1, 40) + Space(38) + rst("id_cliente")
    'e 'nd If
   ' Printer.CurrentY = Printer.CurrentY + 1#
    'Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT C.descripcion as color, U.descripcion as abreviatura,P.id_linea,P.id_producto,P.nombre_prod,M.cantidad,M.precio,M.total,R.descripcion as marca,M.serie,M.nro_chasis,M.anio_modelo,M.anio_fabricacion,M.nro_dua,M.nro_item,cc.descripcion as modelo FROM movimiento_venta_detalle M,producto P,unidad U,marca R,imp_color C,linea_sub cc WHERE cc.id_tipo=P.id_sublinea and cc.id_usu=P.ruc and   P.id_color=C.id_color AND   P.id_marca=R.id_marca AND P.ruc=R.id_usu AND    M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.5
            For j = 0 To rstT.RecordCount - 1
                codigo = Mid(rstT("id_producto") + Space(50), 1, 5)
                cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                Und = rstT("abreviatura")
                descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 75)
                precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 8)
                totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
                Printer.Print Tab(12); cantidad & Space(4) & descripcion & precio & Space(4) & totalPar
                Printer.Print ""
                Printer.Print ""
                Printer.FontBold = True
                If rstT("id_linea") <> "00047" Then
                    Printer.Print Tab(22); "MARCA           :" & Space(2) & rstT("marca")
                    Printer.Print Tab(22); "N°CHASIS        :" & Space(2) & rstT("nro_chasis")
                    Printer.Print Tab(22); "MOTOR           :" & Space(2) & rstT("serie")
                    Printer.Print Tab(22); "MODELO          :" & Space(2) & rstT("modelo")
                    Printer.Print Tab(22); "COLOR           :" & Space(2) & rstT("color")
                    Printer.Print Tab(22); "AÑO FABRICACION :" & Space(2) & rstT("anio_fabricacion")
                    Printer.Print Tab(22); "NRO DUA         :" & Space(2) & rstT("nro_dua")
                    Printer.Print Tab(22); "NRO ITEM        :" & Space(2) & rstT("nro_item")
                    Printer.FontBold = True
                Else
                    Printer.Print Tab(22); "MARCA           :" & Space(2) & rstT("marca")
                    Printer.Print Tab(22); "MOTOR           :" & Space(2) & rstT("serie")
                    Printer.Print Tab(22); "MODELO          :" & Space(2) & rstT("modelo")
                    Printer.Print Tab(22); "AÑO FABRICACION :" & Space(2) & rstT("anio_fabricacion")
                    Printer.Print Tab(22); "NRO DUA         :" & Space(2) & rstT("nro_dua")
                    Printer.Print Tab(22); "NRO ITEM        :" & Space(2) & rstT("nro_item")
                    Printer.FontBold = True
                End If
                rstT.MoveNext
                'Printer.CurrentY = Printer.CurrentY + 0.2
                
            Next j
    End If
            inc = 0.5
            Do While (Val(Printer.CurrentY) <= 12)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    totalletras = UCase(EnLetras(rst("total")))
    
    '---- fin totales
   
    'Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.8
    Printer.Print Tab(18); Mid(Trim(totalletras) + Space(100), 1, 100)
   ' Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.2
   ' Printer.Print Tab(10); "N.DIRECCION FISCAL :" + KEY_DIRECCION
    'Printer.CurrentY = Printer.CurrentY + 0.5
   ' Printer.CurrentY = Printer.CurrentY + 0.5
    'Printer.Print Tab(105); "S/." & Format(rst("total"), "#,##0.00")
    Printer.Print Tab(19); Mid(KEY_VENDEDOR, 1, 15) & Space(1) & "a las:" & str(Now()) & Space(41) & "S/" & Format(rst("total"), "#,##0.00")
    
    
    
    
    
    
    
    
    
    
    Printer.Print ""
    Printer.EndDoc
    
    Exit Sub
End Sub

Private Sub impresion_formato_recibo_suelta_serie(ByVal id_doc As String, ByVal serie As String, ByVal numero As String, Optional Direccion As String)
Dim n_venta As Double
Dim nvendedor As String
Dim ncomprobante As Double
    Printer.ScaleMode = 7
    'Printer.Width = 11
    'Printer.Height = 10
    'Printer.PaperSize = vbPRPSUser
   
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
   ' Printer.Print ""
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0
    Printer.Print Tab(95); "RECIBO"; Space(2); Mid(serie + Space(50), 1, 4) & Space(2) & "-" & Space(2) & Trim(numero)
    'Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT * FROM movimiento_venta V,forma_pago_detalle F WHERE V.id_forma_pago=F.id_detalle AND V.ruc=F.ruc AND  id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND V.ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    
    n_venta = rst("id_venta")
    nvendedor = rst("id_vendedor")
    
    ncomprobante = rst("id_comprobante")
    '***** END COMPROBANTE***
    '**** DETALLE COMPROBANTE ***
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(25); Mid(Trim(rst("ncliente")) + Space(80), 1, 40) + Space(58) & Trim(rst("id_cliente"))
    Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT direccion,id_departamento,id_provincia,id_distrito FROM persona WHERE dni='" & Trim(rst("id_cliente")) & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
            If rstL("id_departamento") <> "0" Then
                            strCadena = "select * from departamentos WHERE id_depa='" & rstL("id_departamento") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndepartamento = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndepartamento = " "
                            End If
                            
                            If rstL("id_provincia") <> "0" Then
                            strCadena = "select * from provincia WHERE id_provincia='" & rstL("id_provincia") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                nprovincia = "- " & UCase(rstZ("descripcion"))
                            Else
                                nprovincia = " "
                            End If
                            End If
                            
                            
                        If rstL("id_distrito") <> "0" Then
                            strCadena = "select * from distrito WHERE id_distrito='" & rstL("id_distrito") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndistrito = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndistrito = " "
                            End If
                    End If
                    End If
          ndireccion = rstL("direccion") & ndistrito & nprovincia & ndepartamento
        Printer.Print Tab(25); Mid(ndireccion + Space(80), 1, 90) + Space(8) + str(rst("fecha_emision"))
    Else
        Printer.Print Tab(25); Mid("---" + Space(80), 1, 40) + Space(58) + rst("id_cliente")
    End If
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT telefono FROM persona_telefono WHERE dni='" & rst("id_cliente") & "' ORDER BY id_telefono DESC LIMIT 0,1"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        Printer.Print Tab(40); "TELEFONO:" + rstZ("telefono")
    Else
        Printer.Print Tab(40); ""
    End If
    
    'Printer.Print Tab(25); Mid(rst("descripcion") + Space(80), 1, 40)
        
    
    'strcadena="SELECT * FROM movimiento_venta V,forma_pago_detalle F WHERE  V.id_forma_pago=F.id_forma_pago WHERE V.id_venta='"&  &"' "
    
    
    'If Len(Trim(direccion)) > 0 Then
     '   Printer.Print Tab(18); Mid(Trim(direccion) + Space(80), 1, 40) + Space(30) + rst("id_cliente")
    'Else
      '  Printer.Print Tab(18); Mid(BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente")) + Space(80), 1, 40) + Space(38) + rst("id_cliente")
    'e 'nd If
   ' Printer.CurrentY = Printer.CurrentY + 1#
    Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT C.descripcion as color, U.descripcion as abreviatura,P.id_producto,P.nombre_prod,M.cantidad,M.precio,M.total,R.descripcion as marca,M.serie,M.nro_chasis,M.anio_modelo,M.anio_fabricacion,M.nro_dua,M.nro_item,cc.descripcion as modelo FROM movimiento_venta_detalle M,producto P,unidad U,marca R,imp_color C,linea_sub cc WHERE P.id_sublinea=cc.id_tipo and P.ruc=cc.id_usu and   P.id_color=C.id_color AND   P.id_marca=R.id_marca AND P.ruc=R.id_usu AND    M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.5
            For j = 0 To rstT.RecordCount - 1
                codigo = Mid(rstT("id_producto") + Space(50), 1, 5)
                cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                Und = rstT("abreviatura")
                descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 85)
                precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 8)
                totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
                Printer.Print Tab(18); cantidad & Space(4) & descripcion & precio & Space(6) & totalPar
                Printer.Print ""
                Printer.Print ""
                nchasis = rstT("nro_chasis")
                Printer.Print Tab(28); "MARCA           :" & Space(2) & rstT("marca")
                Printer.Print Tab(28); "N°CHASIS        :" & Space(2) & rstT("nro_chasis")
                Printer.Print Tab(28); "MOTOR           :" & Space(2) & rstT("serie")
                Printer.Print Tab(28); "MODELO          :" & Space(2) & rstT("modelo")
                Printer.Print Tab(28); "COLOR           :" & Space(2) & rstT("color")
                Printer.Print Tab(28); "AÑO FABRICACION :" & Space(2) & rstT("anio_fabricacion")
                Printer.Print Tab(28); "NRO DUA         :" & Space(2) & rstT("nro_dua")
                Printer.Print Tab(28); "NRO ITEM        :" & Space(2) & rstT("nro_item")
                rstT.MoveNext
                'Printer.CurrentY = Printer.CurrentY + 0.2
                
            Next j
            
           strCadena = "SELECT P.nombre_prod FROM imp_producto_detalle I,imp_producto_insumo S,producto P WHERE I.id_detalle=S.id_producto_detalle AND  S.id_producto=P.id_producto AND I.nro_chasis='" & Trim(nchasis) & "' AND S.id_linea='05' and  I.ruc=P.ruc AND I.ruc='" & KEY_RUC & "'"
           Call ConfiguraRstZ(strCadena)
           If rstZ.RecordCount > 0 Then
                Printer.Print ""
                nincluye = ""
                rstZ.MoveFirst
              For m = 0 To rstZ.RecordCount - 1
                  nincluye = rstZ("nombre_prod") & Space(1) & "-" & Space(1) & nincluye
                  rstZ.MoveNext
              Next m
              Printer.Print Tab(28); "INCLUYE :" & Space(2) & nincluye
           End If
         flag = "normal"
         
         
    Else
       'SI EL CONTENIDO ES UNA AMORTIZACION
        flag = "amortizacion"
        nsaldo = 0
        nsaldo_final = 0
        strCadena = "SELECT v.fecha_emision,v.documento,v.total,v.monto_vehiculo FROM movimiento_venta v WHERE  v.id_venta='" & ncomprobante & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
              nsaldo = rstZ("total")
              Printer.Print ""
              Printer.Print Tab(28); "COMPROBANTE RELACIONADO  :" & Space(2) & rstZ("documento") & Space(3) & Format(rstZ("fecha_emision"), "dd-mm-YYYY") & Space(5) & Format(rstZ("total"), "#,##0.00")
              Printer.Print Tab(28); ""
              Printer.Print Tab(28); "-----------------------------------------------------------------"
              Printer.Print Tab(28); "HISTORIAL DE PAGOS REALIZADOS"
              Printer.Print Tab(28); ""
              strCadena = "SELECT v.fecha_emision,v.documento,m.monto FROM movimiento_venta v,movimiento_venta_monto m WHERE v.id_venta=m.id_venta and  v.id_comprobante='" & ncomprobante & "' and m.id_forma_pago='01' ORDER BY v.fecha_emision ASC"
              Call ConfiguraRstZ(strCadena)
              If rstZ.RecordCount > 0 Then
                 rstZ.MoveFirst
                 
                 For m = 0 To rstZ.RecordCount - 1
                        Printer.Print Tab(28); "                           " & Space(2) & rstZ("documento") & Space(3) & Format(rstZ("fecha_emision"), "dd-mm-YYYY") & Space(5) & " (" & Format(rstZ("monto"), "#,##0.00") & ")"
                        nsaldo = nsaldo - rstZ("monto")
                        nsaldo_final = rstZ("monto")
                        rstZ.MoveNext
                 Next m
                 Printer.Print Tab(28); ""
                 Printer.Print Tab(28); "SALDO RESTANTE             :" & Space(36) & Space(5) & Format(nsaldo, "#,##0.00")
              End If
        Else
              flag = ""
              Printer.Print Tab(28); "MONTO ADELANTADO "
              Printer.Print Tab(28); ""
              Printer.Print Tab(28); "-----------------------------------------------------------------"
              Printer.Print Tab(28); rst("observacion")
              Printer.Print Tab(28); ""
        End If
        
        
    End If
            inc = 0.5
            Do While (Val(Printer.CurrentY) <= 11.5)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    
    
    '---- fin totales
    strCadena = "SELECT * FROM persona where dni='" & nvendedor & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
        Printer.Print Tab(15); "VENDEDOR :" & Space(2) & rstL("nombre_completo")
    Else
        Printer.Print Tab(15); ""
    End If
    Printer.Print Tab(15); ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    If flag = "amortizacion" Then
        totalletras = UCase(EnLetras(nsaldo + nsaldo_final))
        Printer.Print Tab(15); Mid(Trim(totalletras) + Space(100), 1, 100); Space(8) & "S/" & Format(nsaldo + nsaldo_final, "#,##0.00")
    Else
        If flag = "normal" Then
          totalletras = UCase(EnLetras(rst("total") - rst("saldo")))
          Printer.Print Tab(15); Mid(Trim(totalletras) + Space(100), 1, 100); Space(8) & "S/" & Format(rst("total"), "#,##0.00")
        Else
            totalletras = UCase(EnLetras(rst("total") - rst("saldo")))
        Printer.Print Tab(15); Mid(Trim(totalletras) + Space(100), 1, 100); Space(8) & "S/" & Format(rst("monto_vehiculo"), "#,##0.00")
        End If
        
    End If
    
    
    
     Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    
   
                            strCadena = "SELECT monto FROM movimiento_venta_monto WHERE id_venta='" & n_venta & "' and ruc='" & KEY_RUC & "' ORDER BY id_forma_pago ASC"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                rstZ.MoveFirst
                                For j = 0 To rstZ.RecordCount - 1
                                    If rstZ.RecordCount = 1 Then
                                        If rst("total") = rstZ("monto") Then
                                        
                                            If rst("monto_vehiculo") > 0 Then
                                                'Printer.Print Tab(123); "S/." & Format(rst("monto_vehiculo"), "#,##0.00")
                                                'Printer.Print ""
                                                Printer.Print Tab(123); "S/" & Format(rstZ("monto"), "#,##0.00")
                                                Printer.Print ""
                                                Printer.Print Tab(123); "S/" & Format(rst("monto_vehiculo") - rstZ("monto"), "#,##0.00")
                                                GoTo cerrarrecibo
                                            Else
                                                Printer.Print Tab(123); "S/" & Format(rstZ("monto"), "#,##0.00")
                                                Printer.Print ""
                                            End If
                                           ' Printer.Print Tab(123); "S/" & Format(rstZ("monto"), "#,##0.00")
                                           ' Printer.Print ""
                                            
                                            
                                            
                                            If flag = "amortizacion" Then
                                                Printer.Print Tab(123); "S/" & Format(nsaldo, "#,##0.00")
                                            Else
                                                Printer.Print Tab(123); "S/" & Format(rstZ("monto") - rstZ("monto"), "#,##0.00")
                                            End If
                                            
                                        Else
                                            Printer.Print Tab(123); "S/" & Format(rstZ("monto"), "#,##0.00")
                                            Printer.Print ""
                                            Printer.Print Tab(123); "S/" & Format(nsaldo, "#,##0.00")
                                          End If
                                        
                                    Else
                                        Printer.Print Tab(123); "S/" & Format(rstZ("monto"), "#,##0.00")
                                        Printer.CurrentY = Printer.CurrentY + 0.5
                                    End If
                                    rstZ.MoveNext
                                Next j
                                
                            End If
 
    
cerrarrecibo:
    Printer.Print ""
    Printer.EndDoc
    
    Exit Sub
End Sub

Private Sub impresion_formato_factura_suelta_serie(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
    
    Printer.CurrentX = 0
    Printer.CurrentY = 0
 'Printer.PaperSize = vbPRPSLetter
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    'Printer.Print Tab(70); (CVDate(Me.DtpActual.Value))
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
     Printer.Print Tab(99); "FACTURA"; Space(2); Mid(serie + Space(50), 1, 4) & Space(2) & "-" & Space(2) & Trim(numero)
    'Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Printer.CurrentY = Printer.CurrentY + 0.7
    
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(23); Mid(Trim(rst("ncliente")) + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(23); Mid(rst("id_cliente") + Space(100), 1, 75) + Space(25) & str(rst("fecha_emision"))
    Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT direccion,id_departamento,id_provincia,id_distrito FROM  persona WHERE dni='" & Trim(rst("id_cliente")) & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
                    If rstL("id_departamento") <> "0" Then
                            strCadena = "select * from departamentos WHERE id_depa='" & rstL("id_departamento") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndepartamento = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndepartamento = " "
                            End If
                            
                            If rstL("id_provincia") <> "0" Then
                            strCadena = "select * from provincia WHERE id_provincia='" & rstL("id_provincia") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                nprovincia = "- " & UCase(rstZ("descripcion"))
                            Else
                                nprovincia = " "
                            End If
                            End If
                            
                            
                        If rstL("id_distrito") <> "0" Then
                            strCadena = "select * from distrito WHERE id_distrito='" & rstL("id_distrito") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndistrito = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndistrito = " "
                            End If
                    End If
                    End If
          ndireccion = rstL("direccion") & ndistrito & nprovincia & ndepartamento
                    
        Printer.Print Tab(23); Mid(ndireccion + Space(80), 1, 90) + Space(10) + str(rst("fecha_emision"))
    Else
        Printer.Print Tab(23); Mid("---" + Space(80), 1, 40) + Space(58) + rst("id_cliente")
    End If
    Printer.CurrentY = Printer.CurrentY + 0.5
    'Printer.Print Tab(22); Mid(rst("descripcion") + Space(80), 1, 40)
    
    Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    'Printer.Print ""
    
   ' Printer.Print ""
    
    
    strCadena = "SELECT P.id_producto,M.cantidad,U.abreviatura,P.nombre_prod,M.precio,M.total,M.serie,M.nro_chasis,M.anio_modelo,M.anio_fabricacion,M.nro_dua,M.nro_item,mm.descripcion as marca,ss.descripcion as modelo,C.descripcion as color FROM movimiento_venta_detalle M,producto P,unidad U,marca mm,linea_sub ss,imp_color C WHERE  P.id_color=C.id_color and P.id_sublinea = ss.id_tipo and ss.id_usu = P.ruc AND   P.id_marca=mm.id_marca and P.ruc=mm.id_usu and   M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.5
            For j = 0 To rstT.RecordCount - 1
                codigo = Mid(rstT("id_producto") + Space(50), 1, 5)
                cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                Und = rstT("abreviatura")
                descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 84)
                precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 10)
                totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
                
                 Printer.Print Tab(16); cantidad & Space(5) & descripcion & precio & Space(2) & totalPar
                'Printer.CurrentY = Printer.CurrentY + 0.2
                
                Printer.Print ""
                Printer.Print ""
                Printer.Print Tab(25); "MARCA           :" & Space(2) & rstT("marca")
                Printer.Print Tab(25); "N°MOTOR         :" & Space(2) & rstT("serie")
                Printer.Print Tab(25); "N°CHASIS        :" & Space(2) & rstT("nro_chasis")
                Printer.Print Tab(25); "MODELO          :" & Space(2) & rstT("modelo")
                Printer.Print Tab(25); "COLOR           :" & Space(2) & rstT("color")
                Printer.Print Tab(25); "AÑO FABRICACION :" & Space(2) & rstT("anio_fabricacion")
                Printer.Print Tab(25); "NRO DUA         :" & Space(2) & rstT("nro_dua")
                Printer.Print Tab(25); "NRO ITEM        :" & Space(2) & rstT("nro_item")
                rstT.MoveNext
            Next j
            inc = 0.5
           ' Printer.FontBold = True
            Printer.Print ""
            'Printer.FontBold = False
            
            
            
            Do While (Val(Printer.CurrentY) <= 29)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    End If
    
    
    
   
    rst.MoveFirst
    tTotal = rst("total")
    SUBTOTAL = Format(str(rst("valor_venta")), "#,##0.00")
    igv = rst("igv") 'Format((tTotal - SUBTOTAL) / 1.18, "#,##0.00")
    totalletras = UCase(EnLetras(str(tTotal)))
    descuento_i = Mid(Format(str(KEY_DSCTO), "#,##0.00") + Space(4), 1, 6)
    Set rst = Nothing
    '---- fin totales
    Printer.CurrentY = Printer.CurrentY + 0.75
    Printer.Print ""
    If (tTotal = 51922.5) Then
       totalletras = "CINCUENTA Y UN MIL NOVECIENTOS VEINTIDOS CON 50/100 NUEVOS SOLES"
    End If
    Printer.Print Tab(23); Mid(totalletras + Space(100), 1, 80) & Space(22) & "S/" & SUBTOTAL
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    'Printer.Print ""
    Printer.Print Tab(125); " S/" & igv
    Printer.CurrentY = Printer.CurrentY + 1
    'Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(125); "S/" & Format(tTotal, "#,##0.00")
    'Printer.Print Tab(10); "ATENDIDO POR:" & Mid(Trim(KEY_VENDEDOR), 1, 10)
    Printer.EndDoc
    Exit Sub
    
    
    
        
End Sub
Private Sub impresion_formato_factura_suelta_serie_rep(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)

    Printer.CurrentX = 0
    Printer.CurrentY = 0
 'Printer.PaperSize = vbPRPSLetter
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    'Printer.Print Tab(70); (CVDate(Me.DtpActual.Value))
    
     Printer.Print Tab(99); "FACTURA"; Space(2); Mid(serie + Space(50), 1, 4) & Space(2) & "-" & Space(2) & Trim(numero)
    'Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Printer.CurrentY = Printer.CurrentY + 0.7
    in_vendedor = rst("id_vendedor")
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(23); Mid(Trim(rst("ncliente")) + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(23); Mid(rst("id_cliente") + Space(100), 1, 75) + Space(10) & str(rst("fecha_emision"))
    Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT direccion,id_departamento,id_provincia,id_distrito FROM  persona WHERE dni='" & Trim(rst("id_cliente")) & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
                    If rstL("id_departamento") <> "0" Then
                            strCadena = "select * from departamentos WHERE id_depa='" & rstL("id_departamento") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndepartamento = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndepartamento = " "
                            End If
                            
                            If rstL("id_provincia") <> "0" Then
                            strCadena = "select * from provincia WHERE id_provincia='" & rstL("id_provincia") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                nprovincia = "- " & UCase(rstZ("descripcion"))
                            Else
                                nprovincia = " "
                            End If
                            End If
                            
                            
                        If rstL("id_distrito") <> "0" Then
                            strCadena = "select * from distrito WHERE id_distrito='" & rstL("id_distrito") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndistrito = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndistrito = " "
                            End If
                    End If
                    End If
          ndireccion = rstL("direccion") & ndistrito & nprovincia & ndepartamento
                    
        Printer.Print Tab(23); Mid(ndireccion + Space(80), 1, 90) + Space(1) + str(rst("fecha_emision"))
    Else
        Printer.Print Tab(23); Mid("---" + Space(80), 1, 40) + Space(58) + rst("id_cliente")
    End If
    Printer.CurrentY = Printer.CurrentY + 0.5
    'Printer.Print Tab(22); Mid(rst("descripcion") + Space(80), 1, 40)
    
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT P.id_producto,M.cantidad,U.abreviatura,P.nombre_prod,M.precio,M.total,M.serie,M.nro_chasis,M.anio_modelo,M.anio_fabricacion,M.nro_dua,M.nro_item,mm.descripcion as marca,ss.descripcion as modelo FROM movimiento_venta_detalle M,producto P,unidad U,marca mm,linea_sub ss WHERE P.id_sublinea = ss.id_tipo and ss.id_usu = P.ruc   and P.id_marca=mm.id_marca and P.ruc=mm.id_usu and   M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.5
            For j = 0 To rstT.RecordCount - 1
                codigo = Mid(rstT("id_producto") + Space(50), 1, 5)
                cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                Und = rstT("abreviatura")
                descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 75)
                precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 8)
                totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
                Printer.Print Tab(16); cantidad & Space(5) & descripcion & precio & Space(2) & totalPar
                'Printer.CurrentY = Printer.CurrentY + 0.2
                
                Printer.Print ""
                Printer.Print ""
                Printer.Print Tab(25); "MARCA           :" & Space(2) & rstT("marca")
                Printer.Print Tab(25); "N°MOTOR         :" & Space(2) & rstT("serie")
                Printer.Print Tab(25); "MODELO          :" & Space(2) & rstT("modelo")
                Printer.Print Tab(25); "AÑO FABRICACION :" & Space(2) & rstT("anio_fabricacion")
                Printer.Print Tab(25); "NRO DUA         :" & Space(2) & rstT("nro_dua")
                Printer.Print Tab(25); "NRO ITEM        :" & Space(2) & rstT("nro_item")
                rstT.MoveNext
            Next j
            inc = 0.5
           ' Printer.FontBold = True
            Printer.Print ""
            'Printer.FontBold = False
            
            
            
            Do While (Val(Printer.CurrentY) <= 42)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    End If
    
    rst.MoveFirst
    tTotal = rst("total")
    SUBTOTAL = Format(str(rst("valor_venta")), "#,##0.00")
    igv = rst("igv") 'Format((tTotal - SUBTOTAL) / 1.18, "#,##0.00")
    totalletras = UCase(EnLetras(str(tTotal)))
    descuento_i = Mid(Format(str(KEY_DSCTO), "#,##0.00") + Space(4), 1, 6)
    Set rst = Nothing
    '---- fin totales
    Printer.CurrentY = Printer.CurrentY + 2.5
    Printer.Print ""
    If (tTotal = 51922.5) Then
       totalletras = "CINCUENTA Y UN MIL NOVECIENTOS VEINTIDOS CON 50/100 NUEVOS SOLES"
    End If
    Printer.Print Tab(23); Mid(totalletras + Space(100), 1, 80)
    
    Printer.CurrentY = Printer.CurrentY + 1
    'Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.2
    
    Printer.Print Tab(23); Mid(Mid(get_persona(in_vendedor), 1, 15) & Space(1) & "a las:" & str(Now()) + Space(100), 1, 57) & " S/" & SUBTOTAL & Space(5) & "S/" & Format(igv, "#,##0.00") & Space(7) & "S/" & Format(tTotal, "#,##0.00")
    
    'Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    'Printer.Print Tab(10); "ATENDIDO POR:" & Mid(Trim(KEY_VENDEDOR), 1, 10)
    Printer.EndDoc
    Exit Sub
    
    
    
        
End Sub
Private Sub impresion_formato_nota_suelta_serie_rep(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.CurrentX = 0
    Printer.CurrentY = 0
 'Printer.PaperSize = vbPRPSLetter
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    'Printer.Print Tab(70); (CVDate(Me.DtpActual.Value))
    
     Printer.Print Tab(99); "NOTA CREDITO"; Space(2); Mid(serie + Space(50), 1, 4) & Space(2) & "-" & Space(2) & Trim(numero)
    'Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Printer.CurrentY = Printer.CurrentY + 0.7
    in_vendedor = rst("id_vendedor")
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(23); Mid(Trim(rst("ncliente")) + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(23); Mid(rst("id_cliente") + Space(100), 1, 75) + Space(10) & str(rst("fecha_emision"))
    Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT direccion,id_departamento,id_provincia,id_distrito FROM  persona WHERE dni='" & Trim(rst("id_cliente")) & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
                    If rstL("id_departamento") <> "0" Then
                            strCadena = "select * from departamentos WHERE id_depa='" & rstL("id_departamento") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndepartamento = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndepartamento = " "
                            End If
                            
                            If rstL("id_provincia") <> "0" Then
                            strCadena = "select * from provincia WHERE id_provincia='" & rstL("id_provincia") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                nprovincia = "- " & UCase(rstZ("descripcion"))
                            Else
                                nprovincia = " "
                            End If
                            End If
                            
                            
                        If rstL("id_distrito") <> "0" Then
                            strCadena = "select * from distrito WHERE id_distrito='" & rstL("id_distrito") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndistrito = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndistrito = " "
                            End If
                    End If
                    End If
          ndireccion = rstL("direccion") & ndistrito & nprovincia & ndepartamento
                    
        Printer.Print Tab(23); Mid(ndireccion + Space(80), 1, 90) + Space(1) + str(rst("fecha_emision"))
    Else
        Printer.Print Tab(23); Mid("---" + Space(80), 1, 40) + Space(58) + rst("id_cliente")
    End If
    Printer.CurrentY = Printer.CurrentY + 0.5
    'Printer.Print Tab(22); Mid(rst("descripcion") + Space(80), 1, 40)
    
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
      Printer.Print ""
    
    
    strCadena = "SELECT P.id_producto,M.cantidad,U.abreviatura,P.nombre_prod,M.precio,M.total,M.serie,M.nro_chasis,M.anio_modelo,M.anio_fabricacion,M.nro_dua,M.nro_item,mm.descripcion as marca FROM movimiento_venta_detalle M,producto P,unidad U,marca mm WHERE P.id_marca=mm.id_marca and P.ruc=mm.id_usu and   M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.5
            For j = 0 To rstT.RecordCount - 1
                codigo = Mid(rstT("id_producto") + Space(50), 1, 5)
                cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                Und = rstT("abreviatura")
                descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 75)
                precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 8)
                totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
                Printer.Print Tab(16); cantidad & Space(5) & descripcion & precio & Space(2) & totalPar
                'Printer.CurrentY = Printer.CurrentY + 0.2
                
                Printer.Print ""
                Printer.Print ""
                Printer.Print Tab(25); "MARCA           :" & Space(2) & rstT("marca")
                Printer.Print Tab(25); "N°MOTOR         :" & Space(2) & rstT("serie")
                Printer.Print Tab(25); "MODELO          :" & Space(2) & rstT("anio_modelo")
                Printer.Print Tab(25); "AÑO FABRICACION :" & Space(2) & rstT("anio_fabricacion")
                Printer.Print Tab(25); "NRO DUA         :" & Space(2) & rstT("nro_dua")
                Printer.Print Tab(25); "NRO ITEM        :" & Space(2) & rstT("nro_item")
                rstT.MoveNext
            Next j
            inc = 0.5
           ' Printer.FontBold = True
            Printer.Print ""
            'Printer.FontBold = False
            
            
            
            Do While (Val(Printer.CurrentY) <= 42)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    End If
    
    rst.MoveFirst
    tTotal = rst("total")
    
    SUBTOTAL = Format(str(rst("valor_venta")), "#,##0.00")
    igv = rst("igv") 'Format((tTotal - SUBTOTAL) / 1.18, "#,##0.00")
    totalletras = UCase(EnLetras(str(tTotal)))
    descuento_i = Mid(Format(str(KEY_DSCTO), "#,##0.00") + Space(4), 1, 6)
    Set rst = Nothing
    '---- fin totales
    Printer.CurrentY = Printer.CurrentY + 2.5
    Printer.Print ""
    If (tTotal = 51922.5) Then
       totalletras = "CINCUENTA Y UN MIL NOVECIENTOS VEINTIDOS CON 50/100 NUEVOS SOLES"
    End If
    Printer.Print Tab(23); Mid(totalletras + Space(100), 1, 80)
    
    Printer.CurrentY = Printer.CurrentY + 1
    'Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.2
    
    Printer.Print Tab(23); Mid(Mid(get_persona(in_vendedor), 1, 15) & Space(1) & "a las:" & str(Now()) + Space(100), 1, 57) & " S/" & SUBTOTAL & Space(5) & "S/" & Format(igv, "#,##0.00") & Space(7) & "S/" & Format(tTotal, "#,##0.00")
    
    'Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    'Printer.Print Tab(10); "ATENDIDO POR:" & Mid(Trim(KEY_VENDEDOR), 1, 10)
    Printer.EndDoc
    Exit Sub
    
    
    
        
End Sub

Private Sub impresion_formato_factura_suelta(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
 Printer.CurrentX = 0
    Printer.CurrentY = 0
 'Printer.PaperSize = vbPRPSLetter
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    'Printer.Print Tab(70); (CVDate(Me.DtpActual.Value))
    
     Printer.Print Tab(99); "FACTURA"; Space(2); Mid(serie + Space(50), 1, 4) & Space(2) & "-" & Space(2) & Trim(numero)
    'Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Printer.CurrentY = Printer.CurrentY + 0.7
    in_vendedor = rst("id_vendedor")
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(23); Mid(Trim(rst("ncliente")) + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(23); Mid(rst("id_cliente") + Space(100), 1, 75) + Space(13) & str(rst("fecha_emision"))
    Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT direccion,id_departamento,id_provincia,id_distrito FROM  persona WHERE dni='" & Trim(rst("id_cliente")) & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
                    If rstL("id_departamento") <> "0" Then
                            strCadena = "select * from departamentos WHERE id_depa='" & rstL("id_departamento") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndepartamento = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndepartamento = " "
                            End If
                            
                            If rstL("id_provincia") <> "0" Then
                            strCadena = "select * from provincia WHERE id_provincia='" & rstL("id_provincia") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                nprovincia = "- " & UCase(rstZ("descripcion"))
                            Else
                                nprovincia = " "
                            End If
                            End If
                            
                            
                        If rstL("id_distrito") <> "0" Then
                            strCadena = "select * from distrito WHERE id_distrito='" & rstL("id_distrito") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndistrito = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndistrito = " "
                            End If
                    End If
                    End If
          ndireccion = rstL("direccion") & ndistrito & nprovincia & ndepartamento
                    
        Printer.Print Tab(23); Mid(ndireccion + Space(80), 1, 90) + Space(4) + str(rst("fecha_emision"))
    Else
        Printer.Print Tab(23); Mid("---" + Space(80), 1, 40) + Space(58) + rst("id_cliente")
    End If
    Printer.CurrentY = Printer.CurrentY + 0.5
    'Printer.Print Tab(22); Mid(rst("descripcion") + Space(80), 1, 40)
    
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    
    strCadena = "SELECT M.id_producto,M.cantidad,U.abreviatura,P.nombre_prod,M.precio,M.total,ma.descripcion as marca FROM movimiento_venta_detalle M,producto P,unidad U,marca ma WHERE P.id_marca=ma.id_marca and ma.id_usu=P.ruc and   M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.5
            For j = 0 To rstT.RecordCount - 1
                codigo = Mid(rstT("id_producto") + Space(50), 1, 5)
                cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                Und = Mid(rstT("abreviatura") + Space(10), 1, 4)
                strmarca = Mid(rstT("marca") + Space(20), 1, 23)
                descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 48)
                precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 10)
                totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
                Printer.Print Tab(15); cantidad & Space(2) & Und & Space(1) & descripcion & Space(1) & strmarca & precio & Space(4) & totalPar
                'Printer.CurrentY = Printer.CurrentY + 0.2
                
               rstT.MoveNext
            Next j
            inc = 0.5
           ' Printer.FontBold = True
            Printer.Print ""
            'Printer.FontBold = False
            
            
            Do While (Val(Printer.CurrentY) <= 42)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    End If
    
    rst.MoveFirst
    tTotal = rst("total")
    SUBTOTAL = Format(str(rst("valor_venta")), "#,##0.00")
    igv = rst("igv") 'Format((tTotal - SUBTOTAL) / 1.18, "#,##0.00")
    totalletras = UCase(EnLetras(str(tTotal)))
    descuento_i = Mid(Format(str(KEY_DSCTO), "#,##0.00") + Space(4), 1, 6)
    Set rst = Nothing
    '---- fin totales
    Printer.CurrentY = Printer.CurrentY + 2.5
    Printer.Print ""
    If (tTotal = 51922.5) Then
       totalletras = "CINCUENTA Y UN MIL NOVECIENTOS VEINTIDOS CON 50/100 NUEVOS SOLES"
    End If
    Printer.Print Tab(23); Mid(totalletras + Space(100), 1, 80)
    
    Printer.CurrentY = Printer.CurrentY + 1
    'Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.2
    
    Printer.Print Tab(23); Mid(Mid(get_persona(in_vendedor), 1, 15) & Space(1) & "a las:" & str(Now()) + Space(100), 1, 57) & " S/" & SUBTOTAL & Space(5) & "S/" & Format(igv, "#,##0.00") & Space(7) & "S/" & Format(tTotal, "#,##0.00")
    
    'Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    'Printer.Print Tab(10); "ATENDIDO POR:" & Mid(Trim(KEY_VENDEDOR), 1, 10)
    Printer.EndDoc
    Exit Sub
    
    
    
        
End Sub
Private Sub impresion_formato_nota_suelta(ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
  Dim ncomprobante() As String
  Printer.TrackDefault = True 'siempre apunta a la impresora predeter
  Printer.Font.name = "Draft 17cpi"
 Printer.CurrentX = 0
    Printer.CurrentY = 0
 'Printer.PaperSize = vbPRPSLetter
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    'Printer.Print Tab(70); (CVDate(Me.DtpActual.Value))
    
     Printer.Print Tab(99); "NOTA CREDITO"; Space(2); Mid(serie + Space(50), 1, 4) & Space(2) & "-" & Space(2) & Trim(numero)
    'Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & id_doc & "' AND serie='" & serie & "' AND numero='" & numero & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Printer.CurrentY = Printer.CurrentY + 0.7
    in_vendedor = rst("id_vendedor")
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(23); Mid(Trim(rst("ncliente")) + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT direccion,id_departamento,id_provincia,id_distrito FROM  persona WHERE dni='" & Trim(rst("id_cliente")) & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
                    If rstL("id_departamento") <> "0" Then
                            strCadena = "select * from departamentos WHERE id_depa='" & rstL("id_departamento") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndepartamento = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndepartamento = " "
                            End If
                            
                            If rstL("id_provincia") <> "0" Then
                            strCadena = "select * from provincia WHERE id_provincia='" & rstL("id_provincia") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                nprovincia = "- " & UCase(rstZ("descripcion"))
                            Else
                                nprovincia = " "
                            End If
                            End If
                            
                            
                        If rstL("id_distrito") <> "0" Then
                            strCadena = "select * from distrito WHERE id_distrito='" & rstL("id_distrito") & "'"
                            Call ConfiguraRstZ(strCadena)
                            If rstZ.RecordCount > 0 Then
                                ndistrito = "- " & UCase(rstZ("descripcion"))
                            Else
                                ndistrito = " "
                            End If
                    End If
                    End If
          ndireccion = rstL("direccion") & ndistrito & nprovincia & ndepartamento
                    
        Printer.Print Tab(23); Mid(ndireccion + Space(80), 1, 90)
    Else
        Printer.Print Tab(23); Mid("---" + Space(80), 1, 40) + Space(58) + rst("id_cliente")
    End If
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(23); Mid(rst("id_cliente") + Space(100), 1, 75) + Space(8) & str(rst("fecha_emision"))
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    ncomprobante = Split(get_comprobante(rst("id_comprobante")), ":")
    
    Printer.Print Tab(35); ncomprobante(0) + Space(20) & ncomprobante(1) & Space(21) & ncomprobante(2)
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print ""
    Printer.Print ""
    
    strCadena = "SELECT M.id_producto,M.cantidad,U.abreviatura,M.detalle,M.precio,M.total,ma.descripcion as marca FROM movimiento_venta_detalle M,producto P,unidad U,marca ma WHERE P.id_marca=ma.id_marca and ma.id_usu=P.ruc and   M.id_venta='" & rst("id_venta") & "' AND M.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.5
            For j = 0 To rstT.RecordCount - 1
                codigo = Mid(rstT("id_producto") + Space(50), 1, 5)
               
                descripcion = Mid(rstT("detalle") + Space(80), 1, 48)
                precio = Mid(Format(str(rstT("precio")), "#,##0.00") + Space(4), 1, 10)
                totalPar = Mid(Format(str(rstT("total")), "#,##0.00") + Space(4), 1, 8)
                
                If rstT("id_producto") = KEY_COD_PER Then
                    Und = Mid("  " + Space(10), 1, 4)
                    strmarca = Mid("  " + Space(20), 1, 23)
                    cantidad = Mid("  " + Space(10), 1, 7)
                Else
                    Und = Mid(rstT("abreviatura") + Space(10), 1, 4)
                    strmarca = Mid(rstT("marca") + Space(20), 1, 23)
                    cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 7)
                End If
                Printer.Print Tab(15); cantidad & Space(2) & Und & Space(1) & descripcion & Space(1) & strmarca & precio & Space(4) & totalPar
                'Printer.CurrentY = Printer.CurrentY + 0.2
                
               rstT.MoveNext
            Next j
            inc = 0.5
           ' Printer.FontBold = True
            Printer.Print ""
            'Printer.FontBold = False
            
            
            Do While (Val(Printer.CurrentY) <= 42)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    End If
    
    rst.MoveFirst
    If rst("total") < 1 Then
        tTotal = rst("total") * -1
        SUBTOTAL = Format(str(rst("valor_venta") * -1), "#,##0.00")
        igv = rst("igv") * -1 'Format((tTotal - SUBTOTAL) / 1.18, "#,##0.00")
    Else
        tTotal = rst("total")
        SUBTOTAL = Format(str(rst("valor_venta")), "#,##0.00")
        igv = rst("igv")
    End If
    
    totalletras = UCase(EnLetras(str(tTotal)))
    descuento_i = Mid(Format(str(KEY_DSCTO), "#,##0.00") + Space(4), 1, 6)
    Set rst = Nothing
    '---- fin totales
    Printer.CurrentY = Printer.CurrentY + 2.5
    Printer.Print ""
    If (tTotal = 51922.5) Then
       totalletras = "CINCUENTA Y UN MIL NOVECIENTOS VEINTIDOS CON 50/100 NUEVOS SOLES"
    End If
    Printer.Print Tab(23); Mid(totalletras + Space(100), 1, 80)
    
    Printer.CurrentY = Printer.CurrentY + 1
    'Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.2
    
    Printer.Print Tab(23); Mid(Mid(get_persona(in_vendedor), 1, 15) & Space(1) & "a las:" & str(Now()) + Space(100), 1, 57) & " S/" & SUBTOTAL & Space(5) & "S/" & Format(igv, "#,##0.00") & Space(7) & "S/" & Format(tTotal, "#,##0.00")
    
    'Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    
    'Printer.Print Tab(10); "ATENDIDO POR:" & Mid(Trim(KEY_VENDEDOR), 1, 10)
    Printer.EndDoc
    Exit Sub
    
    
    
        
End Sub

Public Sub Imprimir(ByVal idVenta As Double)
Dim RstDoc As New ADODB.Recordset
Dim i As Integer, j As Integer
Dim laVenta, espacios
Dim MES As String
Dim Ans As Boolean
Dim cantidad As String, Und As String, descripcion As String, precio As String
Dim Total As String, SUBTOTAL As String, igv As String
Dim totalPar As String, Descuento As String, GranTotal As String, totalletras As String, Peso As Double
Dim inc As Single, codigo As String, Unidad As String, PesoTotal As Double, Toneladas As String
Dim doc_identidad As String, tTotal As Double, tdescuento As Double, tpago As Double, tvuelto As Double
Dim cod_unico As String, id_cliente As String, fecha_doc As Date, nimpresiones As Integer
Dim per_ruc As String
Dim nombre_persona As String
Dim DireccionCliente As String

  If FrmVentas.cmdimprimir.Enabled = False Then
    Exit Sub
  End If
    
    Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "FontB11"
    Printer.Font.Size = "10"
    
    
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    strCadena = "SELECT * FROM DocumentoVenta WHERE idVenta='" & idVenta & "'"
    Call ConfiguraRst(strCadena)
    cod_unico = rst(0)
    nombre_persona = rst("Persona")
    fecha_doc = rst("dEmisionVenta")
    nimpresiones = rst("impresiones") + 1
    fecha_doc = rst("dEmisionVenta")
    If rst("estado") = "Pendiente" Then
        rst("estado") = "Cancelado"
    End If
    id_cliente = rst("cPersona")
    Set rst = Nothing
    strCadena = "SELECT * FROM Persona WHERE cPersona='" & Val(id_cliente) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        per_ruc = rst("Per_Ruc")
        DireccionCliente = rst("sDireccionCliente1")
    End If

    strCadena = "UPDATE DocumentoVenta SET estado='Cancelado',impresiones='" & nimpresiones & "' WHERE idventa='" & idVenta & "'"
    CnBd.Execute (strCadena)
     
       
    
    If KEY_RUC = "20480460538" Then  ' LIRIO DE LOS VALLES
        'Call Impresion_Lirio(KEY_RUC, doc_cod, serie, Numero, nimpresiones, id_cliente, rst("NombrePersona"), rst("sDireccionCliente1"), rst("Per_Ruc"), fecha_doc)
    End If
    
        
    If KEY_RUC = "20362013802" Then  ' UNIGAS
            strCadena = "SELECT * FROM DocumentoVenta WHERE idVenta='" & idVenta & "' AND Ruc='" & KEY_RUC & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                Call Impresion_Unigas(idVenta, rst("doc_cod"), rst("sSerie"), rst("cDocumentoVenta"), nimpresiones, id_cliente, rst("Persona"), DireccionCliente, rst("dEmisionVenta"), rst("alm_cod"), per_ruc)
            End If
              
    End If
End Sub
Public Sub Impresion_Unigas(ByVal idVenta As Double, ByVal doc_cod As String, ByVal serie As String, ByVal numero As String, ByVal nimpresiones As Integer, ByVal cPersona As String, ByVal Persona As String, ByVal Direccion As String, ByVal fecha As Date, ByVal Alm As String, ByVal per_ruc As String)
Dim Total As Single
Dim Descuento As Single
Dim transporte As Single
strCadena = "SELECT * FROM DocumentoVenta WHERE idVenta='" & idVenta & "' AND Ruc='" & KEY_RUC & "'"
Call ConfiguraTemporal(strCadena)
 If doc_cod = KEY_FACTURA Then
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    If nimpresiones > 1 Then
        Printer.Print Tab(1); "==================================="
        Printer.Print Tab(5); "Copia de Original:" + Space(1) + str(nimpresiones) + Space(1) + "Impresiones"
        Printer.Print Tab(1); "==================================="
    End If
    Printer.Print Tab(1); KEY_EMPRESA
    Printer.Print Tab(1); Mid(KEY_DIRECCION, 1, 37)
    Printer.Print Tab(1); "TEL:(042)522727-(042)521623-RPM:#913647"
    Printer.Print Tab(1); "TARAPOTO - SAN MARTIN - SAN MARTIN"
    Printer.Print Tab(1); "RUC:" + KEY_RUC
    Printer.Print Tab(1); "-----------------------------------"
    Printer.Print Tab(1); "TICKET FACT:"; Mid(serie + Space(50), 1, 4) & "-" & numero & Space(1) & Trim(fecha)
    Printer.Print Tab(1); "CLIENTE  :"; Mid(Persona + Space(80), 1, 30)
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.Print Tab(1); "DIRECCION:"; Mid(Direccion + Space(80), 1, 30)
    Printer.Print Tab(1); "RUC      :"; Mid(per_ruc + Space(80), 1, 30)
    Printer.Print Tab(1); "===================================="
    Printer.Print Tab(1); "CANT" + Space(2) + "DESCRIPCION" + Space(10) + "PV" + Space(5) + "Total"
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.Print Tab(1); "===================================="
    strCadena = "SELECT Detalle_DocumentoVenta.cProducto as Codigo,Detalle_DocumentoVenta.Cantidad as Cantidad,Unidad.sAbreviatura as Unidad,Producto.DescripcionProducto as Producto,Detalle_DocumentoVenta.Precio as Precio,Detalle_DocumentoVenta.Total as Total,transporte " & _
   "FROM Detalle_DocumentoVenta INNER JOIN (Producto INNER JOIN Unidad ON Producto.cunidad=Unidad.cunidad) ON Detalle_DocumentoVenta.cProducto=Producto.cProducto WHERE (Detalle_DocumentoVenta.idVenta='" & idVenta & "')"
    Call ConfiguraRst(strCadena)
    rst.MoveFirst
               For j = 0 To rst.RecordCount - 1
                codigo = rst(0)
                transporte = rst("transporte")
                cantidad = str(rst("Cantidad"))
                Und = rst("Unidad")
                descripcion = Mid(rst("Producto") + Space(80), 1, 40)
                precio = Mid(Format(str(rst("Precio")), "#,##0.00") + Space(4), 1, 6)
                totalPar = Mid(Format(str(rst("Total")), "#,##0.00") + Space(4), 1, 7)
                Printer.Print Tab(0); descripcion
                Printer.Print Tab(2); Format(cantidad, "#,##0.00") & Space(2) + Und + Space(12) & precio & Space(6) & totalPar
                Printer.CurrentY = Printer.CurrentY + 0.1
                rst.MoveNext
            Next j
            rst.MoveFirst
            
    Set rst = Nothing
    tTotal = rstTemporal("nTotalVenta")
    tdescuento = rstTemporal("nDescuento")
    tpago = rstTemporal("monto_pagado")
    tvuelto = rstTemporal("monto_vuelto")
    tsubtotal = tTotal
    tigv = tTotal - tsubtotal
    Dim sin_IGV As Single
    sin_IGV = 0
    Printer.Print Tab(1); "==================================="
    Printer.Print Tab(1); "SUBTOTAL      S/." + Space(1) + Format(tTotal, "#,##0.00")
    Printer.Print Tab(1); "DESCUENTO     S/." + Space(1) + Format(tdescuento, "#,##0.00")
    Printer.Print Tab(1); "IGV(18%)      S/." + Space(1) + Format(sin_IGV, "#,##0.00")
    Printer.Print Tab(1); "TOTAL         S/." + Space(1) + Format(tTotal, "#,##0.00")
    strCadena = "SELECT * FROM DocumentoVenta_montos WHERE cDocumentoVenta='" & Trim(numero) & "' AND id_usuario='" & Trim(KEY_USUARIO) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
        If rst("id_formapago") = "0001" Then
        If rstTemporal("delivery") = "F" Then
    Printer.Print Tab(1); "EFECTIVO         S/." + Space(1) + Format(rst("monto"), "#,##0.00")
        End If
            
        End If
        If rst("id_formapago") = "0002" Then
    Printer.Print Tab(1); "TARJETA CREDITO  S/." + Space(1) + Format(rst("monto"), "#,##0.00")
            
        End If
        If rst("id_formapago") = "0003" Then
    Printer.Print Tab(1); "TARJETA DEBITO   S/." + Space(1) + Format(rst("monto"), "#,##0.00")
        
        End If
    If rst("id_formapago") = "0004" Then
    Printer.Print Tab(1); "CREDITO          S/." + Space(1) + Format(rst("monto"), "#,##0.00")
        End If
        rst.MoveNext
    Next i
    Else
     strCadena = "SELECT idFormaPago,nTotalVenta FROM DocumentoVenta WHERE idVenta='" & idVenta & "'"
     Call ConfiguraRst(strCadena)
     If rst.RecordCount > 1 Then
     rst.MoveFirst
     For i = 0 To rst.RecordCount - 1
        If rst("idFormaPago") = "0001" Then
    Printer.Print Tab(1); "EFECTIVO         S/." + Space(1) + Format(rst("nTotalVenta"), "#,##0.00")
            
        End If
        If rst("idFormaPago") = "0002" Then
    Printer.Print Tab(1); "TARJETA CREDITO  S/." + Space(1) + Format(rst("nTotalVenta"), "#,##0.00")
            
        End If
        If rst("idFormaPago") = "0003" Then
    Printer.Print Tab(1); "TARJETA DEBITO   S/." + Space(1) + Format(rst("nTotalVenta"), "#,##0.00")
        
        End If
    If rst("idFormaPago") = "0004" Then
    Printer.Print Tab(1); "CREDITO          S/." + Space(1) + Format(rst("nTotalVenta"), "#,##0.00")
        End If
        rst.MoveNext
    Next i
     Else
    Printer.Print Tab(1); "EFECTIVO         S/." + Space(1) + Format(tpago, "#,##0.00")
     End If
    
    End If
    If rstTemporal("delivery") = "F" Then
    Printer.Print Tab(1); "VUELTO           S/." + Space(1) + Format(tvuelto, "#,##0.00")
    End If
    Printer.Print Tab(1); "==================================="
    Printer.Print Tab(0); "LO ATENDIO:" + Space(1) + KEY_VENDEDOR + Space(1) + "A LAS:" + str(Time)
    Printer.Print Tab(20); "ID:" & idVenta
    
    'StrCadena = "SELECT puntos FROM Persona WHERE cPersona='" & Trim(id_cliente) & "'"
    'Call ConfiguraRst(StrCadena)
    'Printer.Print Tab(0); "Ud, ha Ganado:" & Space(2) & Str(CInt(tTotal)) & "Punto(s)"
    'Printer.Print Tab(0); "Ud, tiene :" & Str(rst(0)) & Space(2) & "Puntos Acumulados"
    Printer.Print Tab(3); " BIENES TRANSFERIDOS EN LA AMAZONIA"
    Printer.Print Tab(3); "  PARA SER CONSUMIDOS EN LA MISMA"
    Printer.Print Tab(3); "  GRACIAS POR SU COMPRA"
    Printer.Print Tab(3); "      REGRESE PRONTO  "
    Call AbreGaveta
    Printer.EndDoc
    If rstTemporal("delivery") = "V" Then
     Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "FontB11"
    Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print Tab(1); "==================================="
    Printer.Print Tab(1); "COBRANZA"
    Printer.Print Tab(1); "==================================="
    Printer.Print Tab(1); "TICKET:"; Space(1); Mid(serie + Space(50), 2, 4) & "-" & Mid(numero, 4, 10) & Space(10) & "S/." + Space(1) + Format(rstTemporal("nTotalVenta"), "#,##0.00")
    Printer.Print Tab(1); Mid("Por Compra y Delivery" + Space(50), 1, 20) & Space(10) & "S/." + Space(1) + Format(transporte, "#,##0.00")
    Printer.Print Tab(1); Mid("TOTAL" + Space(50), 1, 20) & Space(10) & "S/." + Space(1) + Format(transporte + rstTemporal("nTotalVenta"), "#,##0.00")
    Printer.Print Tab(1); Mid("EFECTIVO " + Space(50), 1, 20) & Space(10) & "S/." + Space(1) + Format(rstTemporal("monto_pagado"), "#,##0.00")
    Printer.Print Tab(1); Mid("VUELTO" + Space(50), 1, 20) + Space(10) & "S/." + Format(rstTemporal("monto_vuelto") - transporte, "#,##0.00")
    Printer.Print Tab(1); "==================================="
    Printer.EndDoc
    End If
    Call FrmVentas.nuevo
    Exit Sub
End If
    
If doc_cod = KEY_BOLETA Then
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    If nimpresiones > 1 Then
    Printer.Print Tab(1); "==================================="
    Printer.Print Tab(1); "Copia de Original:" + Space(1) + str(nimpresiones) + Space(1) + "Impresiones"
    Printer.Print Tab(1); "==================================="
    End If
    Printer.Print Tab(1); KEY_EMPRESA
    Printer.Print Tab(1); Mid(KEY_DIRECCION, 1, 37)
    Printer.Print Tab(1); "TEL:(042)522727-(042)521623-RPM:#913647"
    Printer.Print Tab(1); "TARAPOTO - SAN MARTIN - SAN MARTIN"
    Printer.Print Tab(1); "RUC:" + KEY_RUC
    Printer.Print Tab(1); "-----------------------------------"
    Printer.Print Tab(1); "TICKET BOLT:"; Space(1); Mid(serie + Space(50), 2, 4) & "-" & Mid(numero, 4, 10) & Space(2) & Trim(fecha)
    Printer.Print Tab(1); "CLIENTE  :" + Space(2); Mid(Persona + Space(80), 1, 35)
    Printer.Print Tab(1); "DIRECCION:" + Space(2); Mid(Direccion + Space(80), 1, 35)
    Printer.Print Tab(1); "TELEFONO :" + Space(2); Mid(Telefono + Space(80), 1, 35)
    Printer.Print Tab(1); "===================================="
    Printer.Print Tab(1); "CANT" + Space(2) + "DESCRIPCION" + Space(7) + "PV" + Space(5) + "Total"
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.Print Tab(1); "===================================="
    
    strCadena = "SELECT Detalle_DocumentoVenta.cProducto as Codigo,Detalle_DocumentoVenta.Cantidad as Cantidad,Unidad.sAbreviatura as Unidad,Producto.DescripcionProducto as Producto,Detalle_DocumentoVenta.Precio as Precio,Detalle_DocumentoVenta.Total as Total,transporte " & _
   "FROM Detalle_DocumentoVenta INNER JOIN (Producto INNER JOIN Unidad ON Producto.cunidad=Unidad.cunidad) ON Detalle_DocumentoVenta.cProducto=Producto.cProducto WHERE (Detalle_DocumentoVenta.idventa='" & idVenta & "')"
    Call ConfiguraRst(strCadena)
    rst.MoveFirst
               For j = 0 To rst.RecordCount - 1
                codigo = rst("Codigo")
                transporte = rst("transporte")
                cantidad = str(rst("Cantidad"))
                Und = rst("Unidad")
                descripcion = Mid(rst("Producto") + Space(80), 1, 40)
                precio = Mid(Format(str(rst("Precio")), "#,##0.00") + Space(4), 1, 6)
                totalPar = Mid(Format(str(rst("Total")), "#,##0.00") + Space(4), 1, 7)
                Printer.Print Tab(0); descripcion
                Printer.Print Tab(2); Format(cantidad, "#,##0.00") & Space(2) + Und + Space(11) & precio & Space(3) & totalPar
                Printer.CurrentY = Printer.CurrentY + 0.1
                rst.MoveNext
            Next j
            rst.MoveFirst
    
    tdescuento = rstTemporal("nDescuento")
    tpago = rstTemporal("monto_pagado")
    tvuelto = rstTemporal("monto_vuelto")
    Printer.Print Tab(1); "==================================="
    Dim TTventa As Double
    TTventa = rstTemporal("nTotalVenta")
    If (tdescuento > 0) Then
    Printer.Print Tab(1); "DESCUENTO        S/." + Space(1) + Format(tdescuento, "#,##0.00")
    End If
    Printer.Print Tab(1); Mid("TOTAL" + Space(50), 1, 25) & "S/." + Space(1) + Format(TTventa, "#,##0.00")
    strCadena = "SELECT * FROM DocumentoVenta_montos WHERE cDocumentoVenta='" & Trim(numero) & "' AND id_usuario='" & Trim(KEY_USUARIO) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
        If rst("id_formapago") = "0001" Then
        If rstTemporal("delivery") = "F" Then
    Printer.Print Tab(1); Mid("EFECTIVO" + Space(50), 1, 25) & "S/." + Space(1) + Format(rst("monto"), "#,##0.00")
        End If
        End If
        If rst("id_formapago") = "0002" Then
    Printer.Print Tab(1); Mid("TARJETA CREDITO" + Space(50), 1, 25) & "S/." + Space(1) + Format(rst("monto"), "#,##0.00")
            
        End If
        If rst("id_formapago") = "0003" Then
    Printer.Print Tab(1); Mid("TARJETA DEBITO" + Space(50), 1, 25) & "S/." + Space(1) + Format(rst("monto"), "#,##0.00")
        
        End If
    If rst("id_formapago") = "0004" Then
    Printer.Print Tab(1); Mid("CREDITO" + Space(50), 1, 25) & "S/." + Space(1) + Format(rst("monto"), "#,##0.00")
        End If
        rst.MoveNext
    Next i
    Else
     strCadena = "SELECT idFormaPago,nTotalVenta FROM DocumentoVenta WHERE  idVenta='" & idVenta & "'"
     Call ConfiguraRst(strCadena)
     If rst.RecordCount > 1 Then
     rst.MoveFirst
     For i = 0 To rst.RecordCount - 1
        If rst("idFormaPago") = "0001" Then
    Printer.Print Tab(1); "EFECTIVO         S/." + Space(1) + Format(rst("nTotalVenta"), "#,##0.00")
            
        End If
        If rst("idFormaPago") = "0002" Then
    Printer.Print Tab(1); "TARJETA CREDITO  S/." + Space(1) + Format(rst("nTotalVenta"), "#,##0.00")
            
        End If
        If rst("idFormaPago") = "0003" Then
    Printer.Print Tab(1); "TARJETA DEBITO   S/." + Space(1) + Format(rst("nTotalVenta"), "#,##0.00")
        
        End If
    If rst("idFormaPago") = "0004" Then
    Printer.Print Tab(1); "CREDITO          S/." + Space(1) + Format(rst("nTotalVenta"), "#,##0.00")
        End If
        rst.MoveNext
    Next i
     Else
    Printer.Print Tab(1); "EFECTIVO         S/." + Space(1) + Format(tpago, "#,##0.00")
     End If
    
    End If
    If rstTemporal("delivery") = "F" Then
    Printer.Print Tab(1); Mid("VUELTO" + Space(50), 1, 25) & "S/." + Space(1) + Format(tvuelto, "#,##0.00")
    End If
    Printer.Print Tab(1); "==================================="
    Printer.Print Tab(0); "LO ATENDIO:" + Space(1) + KEY_VENDEDOR + Space(1) + "A LAS:" + str(Time)
    Printer.Print Tab(20); "ID:" & idVenta
    
    Printer.Print Tab(3); " BIENES TRANSFERIDOS EN LA AMAZONIA"
    Printer.Print Tab(3); "  PARA SER CONSUMIDOS EN LA MISMA"
    Printer.Print Tab(3); "  GRACIAS POR SU COMPRA"
    Printer.Print Tab(3); "     REGRESE PRONTO  "
    Call AbreGaveta
    Printer.EndDoc
    
    If rstTemporal("delivery") = "V" Then
     Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "FontB11"
    Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print Tab(1); "==================================="
    Printer.Print Tab(1); "COBRANZA"
    Printer.Print Tab(1); "==================================="
    Printer.Print Tab(1); "TICKET:"; Space(1); Mid(serie + Space(50), 2, 4) & "-" & Mid(numero, 4, 10) & Space(10) & "S/." + Space(1) + Format(rstTemporal("nTotalVenta"), "#,##0.00")
    Printer.Print Tab(1); Mid("Por Compra y Delivery" + Space(50), 1, 21) & Space(9) & "S/." + Space(1) + Format(transporte, "#,##0.00")
    Printer.Print Tab(1); Mid("TOTAL" + Space(50), 1, 20) & Space(10) & "S/." + Space(1) + Format(transporte + rstTemporal("nTotalVenta"), "#,##0.00")
    Printer.Print Tab(1); Mid("EFECTIVO " + Space(50), 1, 20) & Space(10) & "S/." + Space(1) + Format(rstTemporal("monto_pagado"), "#,##0.00")
    Printer.Print Tab(1); Mid("VUELTO" + Space(50), 1, 20) + Space(10) & "S/." & Format(rstTemporal("monto_vuelto") - transporte, "#,##0.00")
    Printer.Print Tab(1); "==================================="
    Printer.EndDoc
    End If
    Call FrmVentas.nuevo
    Exit Sub
End If


If Trim(doc_cod) = KEY_COTIZA Then
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.3
    Printer.Print Tab(4); "COD      :" + Space(1); cPersona
    Printer.Print Tab(4); "CLIENTE  :" + Space(1); Mid(Persona + Space(80), 1, 40) & Space(1) & CVDate(fecha)
    Printer.Print ""
    Printer.Print Tab(4); "DIRECCIÓN:" + Space(1); Mid(Direccion + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0
    Printer.Print Tab(40); "COTIZA:"; Space(2); Mid(serie + Space(50), 1, 4) & Space(2) & "-" & Space(2) & numero
    Printer.Print Tab(1); "LO ATENDIO:" + Space(1) + KEY_VENDEDOR + Space(2) + "A LAS:" + Space(1) + str(Time)
    Printer.Print "----------------------------------------------------------------------"
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT Detalle_DocumentoVenta.cProducto as Codigo,Detalle_DocumentoVenta.Cantidad as Cantidad,Unidad.sAbreviatura as Unidad,Producto.DescripcionProducto as Producto,Detalle_DocumentoVenta.Precio as Precio,Detalle_DocumentoVenta.Total as Total " & _
   "FROM Detalle_DocumentoVenta INNER JOIN (Producto INNER JOIN Unidad ON Producto.cunidad=Unidad.cunidad) ON Detalle_DocumentoVenta.cProducto=Producto.cProducto WHERE (Detalle_DocumentoVenta.cDocumentoVenta='" & numero & "' AND Detalle_DocumentoVenta.doc_cod='" & doc_cod & "' AND Detalle_DocumentoVenta.sSerie='" & serie & "' AND Detalle_DocumentoVenta.Alm_Cod='" & Alm & "')"
    Call ConfiguraRst(strCadena)
    
    rst.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print "======================================================================"
    Printer.Print Tab(2); "CODIGO" & Space(1) & "CANT" & Space(1) & "UND" & Space(2) & "DESCRIPCION" & Space(28) & "PRECIO" & Space(5) & "TOTAL"
    Printer.Print "======================================================================"
    Total = 0
            For j = 0 To rst.RecordCount - 1
                codigo = Mid(FormatosCeros(rst(0), 4) + Space(50), 1, 4)
                cantidad = Mid(str(rst(1)) + Space(10), 1, 5)
                Und = Mid(rst(2) + Space(10), 1, 5)
                descripcion = Mid(rst(3) + Space(80), 1, 34)
                precio = Mid(Format(str(rst(4)), "#,##0.00") + Space(4), 1, 7)
                totalPar = Mid(Space(3) + Format(str(rst(5)), "#,##0.00"), 1, 10)
                Total = Total + rst(5)
                Printer.Print Tab(3); codigo & Space(1) & cantidad & Space(1) & Und & Space(2) & descripcion & Space(3) & precio & totalPar
                Printer.CurrentY = Printer.CurrentY + 0.2
                rst.MoveNext
            Next j
            inc = 0.5
            Do While (Val(Printer.CurrentY) <= 19)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    
    Total = Total
    Descuento = Format(str(KEY_DSCTO), "#,##0.00")
    totalletras = UCase(EnLetras(str(Total)))
    Set rst = Nothing
    Printer.CurrentY = Printer.CurrentY + 0.4
    Printer.Print Tab(10); Mid(totalletras + Space(100), 1, 100)
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Print Tab(45); "TOTAL COTIZACION:" + Space(2) + Format(Total, "#,##0.00")
    Printer.Print "======================================================================"
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); KEY_DIRECCION
    Printer.Print ""
    Printer.Print Tab(5); "CANJEE ESTE DOCUMENTO POR : BOLETA / FACTURA   GRACIAS !!!"
    Printer.EndDoc
    Call FrmVentas.nuevo
    Exit Sub
End If

'---------------------------------------------
If Trim(doc_cod) = KEY_GUIA Then
    Dim RucEmpTrans As String * 11
    Dim RazonSocialTrans As String
    Dim DomicilioTrans As String
    Dim strMTC As String
    Dim marca As String, Placa As String
    Dim Licencia As String, Chofer As String
    Dim PesoFormato As String
    Dim PesoTotalForm As String
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.PaperSize = 1
    Printer.Print "" 'Tab(10); "1 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "2 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "3 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "4 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "5 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "6 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "7 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "8 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "9 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "10---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "12---------------------------------------------------------------------"
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT sOrigen, sDestino, sRazonDestinatario, sRucDestinatario, sRucTransporte, sEmpresaTransporte," & _
    "sDireccionTransporte, MTC, marca, placa, slicencia,Chofer FROM DetalleGuia WHERE DetalleGuia.sSerieGuia='" & Trim(serie) & "' " & _
            "AND DetalleGuia.sNumeroGuia='" & Trim(numero) & "'AND DetalleGuia.doc_cod='" & Trim(doc_cod) & "' " & _
            "AND DetalleGuia.Alm_cod='" & Trim(Alm) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount <= 0 Then
                MsgBox "Guia Imcompleta favor de llenar Bien los Datos", vbInformation, KEY_EMPRESA
                Exit Sub
            End If
    ruc_transporte = rst("sRucTransporte")
    razon_transporte = rst("sEmpresaTransporte")
    dir_transporte = rst("sDireccionTransporte")
    marca = rst("marca")
    'MTC = rst("MTC")
    Placa = rst("placa")
    Licencia = rst("slicencia")
    Chofer = rst("Chofer")
    Printer.Print Tab(8); str(fecha) & Space(35) & str(doc_cod) & Space(20) & "GUIAREM" & Space(1) & Trim(serie) & "-" & numero
    Printer.Print "" 'Tab(10); "16 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "17 ---------------------------------------------------------------------"
    Printer.Print Tab(0); Mid(rst(0) + Space(80), 1, 67) & Space(2) & Mid(rst(1) + Space(80), 1, 70)
    Printer.Print "" 'Tab(10); "19 ---------------------------------------------------------------------"
    'Printer.Print "" 'Tab(10); "20 ---------------------------------------------------------------------"
    Printer.Print "" ' Tab(10); "21 ---------------------------------------------------------------------"
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(75); Mid(marca + Space(80), 1, 20) & Space(5) & Mid(Placa + Space(80), 1, 20)
   Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(6); Mid(rst(2) + Space(80), 1, 65) & Space(20) & Mid(MTC + Space(80), 1, 20)
    Printer.Print Tab(6); Mid(rst(3) + Space(80), 1, 20)
    Printer.Print Tab(80); Mid(Licencia + Space(80), 1, 20)
    Printer.Print Tab(70); Mid(Chofer + Space(80), 1, 20)
    Printer.Print "" ' Tab(10); "27---------------------------------------------------------------------"
    Printer.Print Tab(10); 'Tab(10); "28---------------------------------------------------------------------"
    strCadena = "SELECT Detalle_DocumentoVenta.cProducto as Codigo,Detalle_DocumentoVenta.Cantidad as Cantidad,Unidad.sAbreviatura as Unidad,Producto.DescripcionProducto as Producto,Producto.prod_peso as Peso " & _
   "FROM Detalle_DocumentoVenta INNER JOIN (Producto INNER JOIN Unidad ON Producto.cunidad=Unidad.cunidad) ON Detalle_DocumentoVenta.cProducto=Producto.cProducto WHERE (Detalle_DocumentoVenta.cDocumentoVenta='" & numero & "' AND Detalle_DocumentoVenta.doc_cod='" & doc_cod & "' AND Detalle_DocumentoVenta.sSerie='" & serie & "' AND Detalle_DocumentoVenta.Alm_Cod='" & Alm & "')"
    Call ConfiguraRst(strCadena)
    
        If rst.RecordCount <= 0 Then
            MsgBox "No Hay Productos Registrados", vbInformation, KEY_EMPRESA
            Exit Sub
        End If
        rst.MoveFirst
        Peso = 0
            For j = 0 To rst.RecordCount - 1
                codigo = Mid((rst(0)) + Space(50), 1, 4)
                Und = Mid((rst("Unidad")) + Space(50), 1, 4)
                descripcion = Mid(rst("Producto") + Space(80), 1, 70)
                cantidad = Mid(Format(str(rst("Cantidad")), "###,##0.00") + Space(10), 1, 4)
                Toneladas = rst("Peso")
                'Format(Str((Rst(3) * Val(Cantidad) / 1000)), "###,##0.00")
                Peso = Peso + Toneladas
                PesoFormato = Format(Toneladas, "#,##0.00")
                Printer.Print Tab(-10); codigo & Space(4) & descripcion & Space(16) & Und & Space(4) & PesoFormato & Space(9) & cantidad
                Printer.CurrentY = Printer.CurrentY + 0.2
                rst.MoveNext
            Next j
            inc = 0.5
            ' Printer.FontBold = True
            Printer.Print Tab(14); "NUEVA DF:" & KEY_DIRECCION
            Do While (Val(Printer.CurrentY) <= 35)
                Printer.CurrentY = Printer.CurrentY + inc
                            Loop
    rst.MoveFirst
    PesoTotalForm = Format(str(Peso), "#,##0.00")
    Printer.Print Tab(48); "PESO TOTAL ->"; Space(55) & PesoTotalForm + Space(2) + "Kg."
    Printer.Print "" '29---------------------------------------------------------------------"
    Printer.Print "" '30---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "31---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "32---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "33---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "33 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "34---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "35---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "36---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "37---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "38---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "39---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "40---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "41---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "42---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "43---------------------------------------------------------------------"
    Printer.Print Tab(10); Mid(razon_transporte + Space(80), 1, 65); '49-------
    Printer.Print Tab(10); Mid(ruc_transporte + Space(80), 1, 65);   '50-------
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print Tab(2); Mid(dir_transporte + Space(80), 1, 40);   '51-------
    Printer.Print Tab(2); Mid(dir_transporte + Space(80), 41, 50);   '51-------
    Printer.Print "" 'Tab(10); "52---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "53---------------------------------------------------------------------"
    'Printer.Print "" 'Tab(10); "54---------------------------------------------------------------------"
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Print Tab(-10); Right(Trim(FrmVentas.TxtSeri_guia.Text), 3) + "-" + Right(FrmVentas.TxtNumero_guia.Text, 7) & Space(10) & FrmVentas.DtcComprobanteGuia.Text
    Printer.EndDoc
    Exit Sub
End If

End Sub

Public Sub Impresion_Olivos(ByVal ruc As String, ByVal doc_cod As String, ByVal serie As String, ByVal numero As String, ByVal impresiones As Integer, ByVal cPersona As String, ByVal Persona As String, ByVal Direccion As String, ByVal fecha As Date, ByVal Alm As String, ByVal per_ruc As String)
Dim Total As Single
Dim Descuento As Single
 If Trim(doc_cod) = KEY_FACTURA Then
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    If nimpresiones > 1 Then
        Printer.Print Tab(1); "==================================="
        Printer.Print Tab(5); "Copia de Original:" + Space(1) + str(nimpresiones) + Space(1) + "Impresiones"
        Printer.Print Tab(1); "==================================="
    End If
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(99); "FACTURA"; Space(1); Mid(serie + Space(50), 1, 4) & Space(1) & "-" & numero
    Printer.CurrentY = Printer.CurrentY + 0.7
    Printer.Print ""
    Printer.Print Tab(14); Mid(Persona + Space(80), 1, 65)
  '  Printer.CurrentY = Printer.CurrentY + 1
    Printer.Print ""
    Printer.Print Tab(14); Mid(Direccion + Space(80), 1, 75)
   ' Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(14); Mid(per_ruc + Space(100), 1, 85) & Space(10) & (CVDate(fecha))
    Printer.Print ""
    Printer.Print ""
    
    strCadena = "SELECT Detalle_DocumentoVenta.cProducto as Codigo,Detalle_DocumentoVenta.Cantidad as Cantidad,Unidad.sAbreviatura as Unidad,Producto.DescripcionProducto as Producto,Detalle_DocumentoVenta.Precio as Precio,Detalle_DocumentoVenta.Total as Total " & _
   "FROM Detalle_DocumentoVenta INNER JOIN (Producto INNER JOIN Unidad ON Producto.cunidad=Unidad.cunidad) ON Detalle_DocumentoVenta.cProducto=Producto.cProducto WHERE (Detalle_DocumentoVenta.cDocumentoVenta='" & numero & "' AND Detalle_DocumentoVenta.doc_cod='" & doc_cod & "' AND Detalle_DocumentoVenta.sSerie='" & serie & "' AND Detalle_DocumentoVenta.Alm_Cod='" & Alm & "')"
    Call ConfiguraRst(strCadena)
    rst.MoveFirst
    Total = 0
               For j = 0 To rst.RecordCount - 1
                codigo = Mid(FormatosCeros(rst(0), 4) + Space(50), 1, 4)
                cantidad = Mid(str(rst(1)) + Space(10), 1, 6)
                Und = rst(2)
                descripcion = Mid(rst(3) + Space(80), 1, 85)
                precio = Mid(Format(str(rst(4)), "#,##0.00") + Space(4), 1, 8)
                totalPar = Mid(Format(str(rst(5)), "#,##0.00") + Space(4), 1, 8)
                Total = Total + rst(5)
                Printer.Print Tab(4); cantidad & Space(4) & descripcion & "S/." & precio & Space(3) & "S/." & totalPar
                Printer.CurrentY = Printer.CurrentY + 0.4
                rst.MoveNext
            Next j
            rst.MoveFirst
            inc = 0.5
           ' Printer.Print Tab(14); "NUEVA DF:" & KEY_DIRECCION
            Do While (Val(Printer.CurrentY) <= 28.5)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    
    
     If KEY_CON_IGV = "si" Then
        SUBTOTAL = Format(Total / (1 + KEY_IGV), "#,##0.00")
        igv = Format(SUBTOTAL * (KEY_IGV), "#,##0.00")
    Else
       SUBTOTAL = Format(Total, "#,##0.00")
        igv = Format(0, "#,##0.00")
    End If
      
    totalletras = UCase(EnLetras(str(Total)))
    Descuento = Mid(Format(str(KEY_DSCTO), "#,##0.00") + Space(4), 1, 6)
    Set rst = Nothing
   '  Printer.CurrentY = Printer.CurrentY + 1.5
    Printer.Print ""
    Printer.Print Tab(4); Mid(totalletras + Space(100), 1, 66) & Space(40) & "S/." & SUBTOTAL
     Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 1
    Printer.Print Tab(110); " S/." & igv
     Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 1
    Printer.Print Tab(110); "S/." & Format(Total, "#,##0.00")
   ' Call AbreGaveta
    Printer.EndDoc
    Call FrmVentas.nuevo
    Exit Sub
End If



If Trim(doc_cod) = KEY_BOLETA Then
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    If nimpresiones > 1 Then
        Printer.Print Tab(1); "==================================="
        Printer.Print Tab(5); "Copia de Original:" + Space(1) + str(nimpresiones) + Space(1) + "Impresiones"
        Printer.Print Tab(1); "==================================="
    End If
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0
    Printer.Print Tab(68); "BOLETA"; Space(2); Mid(serie + Space(50), 1, 4) & Space(2) & "-" & Space(2) & numero
    Printer.CurrentY = Printer.CurrentY + 0.5
    If Trim(ruc) <> "" Then
        Printer.Print Tab(15); Mid(Persona + Space(80), 1, 40) + Space(20) & str(CVDate(fecha))
    Else
    Printer.Print Tab(15); Mid(Persona + Space(80), 1, 40) + Space(20) & str(CVDate(fecha))
    End If
    Printer.CurrentY = Printer.CurrentY + 0.5
    If (Len(ruc) = 0) Then
        ruc = "."
    End If
    Printer.Print Tab(15); Mid(Direccion + Space(80), 1, 40) + Space(20) + per_ruc
    Printer.CurrentY = Printer.CurrentY + 1#
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT Detalle_DocumentoVenta.cProducto as Codigo,Detalle_DocumentoVenta.Cantidad as Cantidad,Unidad.sAbreviatura as Unidad,Producto.DescripcionProducto as Producto,Detalle_DocumentoVenta.Precio as Precio,Detalle_DocumentoVenta.Total as Total " & _
   "FROM Detalle_DocumentoVenta INNER JOIN (Producto INNER JOIN Unidad ON Producto.cunidad=Unidad.cunidad) ON Detalle_DocumentoVenta.cProducto=Producto.cProducto WHERE (Detalle_DocumentoVenta.cDocumentoVenta='" & numero & "' AND Detalle_DocumentoVenta.doc_cod='" & doc_cod & "' AND Detalle_DocumentoVenta.sSerie='" & serie & "' AND Detalle_DocumentoVenta.Alm_Cod='" & Alm & "')"
    Call ConfiguraRst(strCadena)
    rst.MoveFirst
                Total = 0
               For j = 0 To rst.RecordCount - 1
                codigo = Mid(FormatosCeros(rst(0), 4) + Space(50), 1, 5)
                cantidad = Mid(str(rst(1)) + Space(10), 1, 5)
                Und = rst(2)
                descripcion = Mid(rst(3) + Space(80), 1, 63)
                precio = Mid(Format(str(rst(4)), "#,##0.00") + Space(4), 1, 8)
                totalPar = Mid(Format(str(rst(5)), "#,##0.00") + Space(4), 1, 9)
                Total = Total + rst(5)
                Printer.Print Tab(5); cantidad & Space(4) & descripcion & precio & Space(4) & totalPar
                Printer.CurrentY = Printer.CurrentY + 0.2
                rst.MoveNext
            Next j
            inc = 0.5
            Do While (Val(Printer.CurrentY) <= 18.5)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    totalletras = UCase(EnLetras(str(Total)))
    Set rst = Nothing
    Descuento = Format(str(KEY_DSCTO), "#,##0.00")
    Printer.Print Tab(10); Mid(totalletras + Space(100), 1, 100)
     Printer.CurrentY = Printer.CurrentY + 1.5
    Printer.Print Tab(54); Mid(Total & Space(20), 1, 10) & Descuento & Space(20) & "S/." & Format(Total, "#,##0.00")
    Printer.Print ""
    Printer.FontBold = True
   ' Printer.Print Tab(20); "NUEVA DF:" + Space(3) + KEY_DIRECCION
    Call AbreGaveta
    Printer.EndDoc
    Call FrmVentas.nuevo
    Exit Sub
End If


If Trim(doc_cod) = KEY_COTIZA Then
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.3
    Printer.Print Tab(4); "COD      :" + Space(1); cPersona
    Printer.Print Tab(4); "CLIENTE  :" + Space(1); Mid(Persona + Space(80), 1, 40) & Space(1) & CVDate(fecha)
    Printer.Print ""
    Printer.Print Tab(4); "DIRECCIÓN:" + Space(1); Mid(Direccion + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0
    Printer.Print Tab(40); "COTIZA:"; Space(2); Mid(serie + Space(50), 1, 4) & Space(2) & "-" & Space(2) & numero
    Printer.Print Tab(1); "LO ATENDIO:" + Space(1) + KEY_VENDEDOR + Space(2) + "A LAS:" + Space(1) + str(Time)
    Printer.Print "----------------------------------------------------------------------"
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT Detalle_DocumentoVenta.cProducto as Codigo,Detalle_DocumentoVenta.Cantidad as Cantidad,Unidad.sAbreviatura as Unidad,Producto.DescripcionProducto as Producto,Detalle_DocumentoVenta.Precio as Precio,Detalle_DocumentoVenta.Total as Total " & _
   "FROM Detalle_DocumentoVenta INNER JOIN (Producto INNER JOIN Unidad ON Producto.cunidad=Unidad.cunidad) ON Detalle_DocumentoVenta.cProducto=Producto.cProducto WHERE (Detalle_DocumentoVenta.cDocumentoVenta='" & numero & "' AND Detalle_DocumentoVenta.doc_cod='" & doc_cod & "' AND Detalle_DocumentoVenta.sSerie='" & serie & "' AND Detalle_DocumentoVenta.Alm_Cod='" & Alm & "')"
    Call ConfiguraRst(strCadena)
    
    rst.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print "======================================================================"
    Printer.Print Tab(2); "CODIGO" & Space(1) & "CANT" & Space(1) & "UND" & Space(2) & "DESCRIPCION" & Space(28) & "PRECIO" & Space(5) & "TOTAL"
    Printer.Print "======================================================================"
    Total = 0
            For j = 0 To rst.RecordCount - 1
                codigo = Mid(FormatosCeros(rst(0), 4) + Space(50), 1, 4)
                cantidad = Mid(str(rst(1)) + Space(10), 1, 5)
                Und = Mid(rst(2) + Space(10), 1, 5)
                descripcion = Mid(rst(3) + Space(80), 1, 34)
                precio = Mid(Format(str(rst(4)), "#,##0.00") + Space(4), 1, 7)
                totalPar = Mid(Space(3) + Format(str(rst(5)), "#,##0.00"), 1, 10)
                Total = Total + rst(5)
                Printer.Print Tab(3); codigo & Space(1) & cantidad & Space(1) & Und & Space(2) & descripcion & Space(3) & precio & totalPar
                Printer.CurrentY = Printer.CurrentY + 0.2
                rst.MoveNext
            Next j
            inc = 0.5
            Do While (Val(Printer.CurrentY) <= 19)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    
    Total = Total
    Descuento = Format(str(KEY_DSCTO), "#,##0.00")
    totalletras = UCase(EnLetras(str(Total)))
    Set rst = Nothing
    Printer.CurrentY = Printer.CurrentY + 0.4
    Printer.Print Tab(10); Mid(totalletras + Space(100), 1, 100)
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Print Tab(45); "TOTAL COTIZACION:" + Space(2) + Format(Total, "#,##0.00")
    Printer.Print "======================================================================"
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); KEY_DIRECCION
    Printer.Print ""
    Printer.Print Tab(5); "CANJEE ESTE DOCUMENTO POR : BOLETA / FACTURA   GRACIAS !!!"
    Printer.EndDoc
    Call FrmVentas.nuevo
    Exit Sub
End If

'---------------------------------------------
If Trim(doc_cod) = KEY_GUIA Then
    Dim RucEmpTrans As String * 11
    Dim RazonSocialTrans As String
    Dim DomicilioTrans As String
    Dim strMTC As String
    Dim marca As String, Placa As String
    Dim Licencia As String, Chofer As String
    Dim PesoFormato As String
    Dim PesoTotalForm As String
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.PaperSize = 1
    Printer.Print "" 'Tab(10); "1 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "2 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "3 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "4 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "5 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "6 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "7 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "8 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "9 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "10---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "12---------------------------------------------------------------------"
    
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    strCadena = "SELECT sOrigen, sDestino, sRazonDestinatario, sRucDestinatario, sRucTransporte, sEmpresaTransporte," & _
    "sDireccionTransporte, MTC, marca, placa, slicencia,Chofer FROM DetalleGuia WHERE DetalleGuia.sSerieGuia='" & Trim(serie) & "' " & _
            "AND DetalleGuia.sNumeroGuia='" & Trim(numero) & "'AND DetalleGuia.doc_cod='" & Trim(doc_cod) & "' " & _
            "AND DetalleGuia.Alm_cod='" & Trim(Alm) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount <= 0 Then
                MsgBox "Guia Imcompleta favor de llenar Bien los Datos", vbInformation, KEY_EMPRESA
                Exit Sub
            End If
    ruc_transporte = rst("sRucTransporte")
    razon_transporte = rst("sEmpresaTransporte")
    dir_transporte = rst("sDireccionTransporte")
    marca = rst("marca")
    'MTC = rst("MTC")
    Placa = rst("placa")
    Licencia = rst("slicencia")
    Chofer = rst("Chofer")
    Printer.Print Tab(8); str(fecha) & Space(35) & str(doc_cod) & Space(20) & "GUIAREM" & Space(1) & Trim(serie) & "-" & numero
    Printer.Print "" 'Tab(10); "16 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "17 ---------------------------------------------------------------------"
    Printer.Print Tab(0); Mid(rst(0) + Space(80), 1, 67) & Space(2) & Mid(rst(1) + Space(80), 1, 70)
    Printer.Print "" 'Tab(10); "19 ---------------------------------------------------------------------"
    'Printer.Print "" 'Tab(10); "20 ---------------------------------------------------------------------"
    Printer.Print "" ' Tab(10); "21 ---------------------------------------------------------------------"
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(75); Mid(marca + Space(80), 1, 20) & Space(5) & Mid(Placa + Space(80), 1, 20)
   Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(6); Mid(rst(2) + Space(80), 1, 65) & Space(20) & Mid(MTC + Space(80), 1, 20)
    Printer.Print Tab(6); Mid(rst(3) + Space(80), 1, 20)
    Printer.Print Tab(80); Mid(Licencia + Space(80), 1, 20)
    Printer.Print Tab(70); Mid(Chofer + Space(80), 1, 20)
    Printer.Print "" ' Tab(10); "27---------------------------------------------------------------------"
    Printer.Print Tab(10); 'Tab(10); "28---------------------------------------------------------------------"
    strCadena = "SELECT Detalle_DocumentoVenta.cProducto as Codigo,Detalle_DocumentoVenta.Cantidad as Cantidad,Unidad.sAbreviatura as Unidad,Producto.DescripcionProducto as Producto,Producto.prod_peso as Peso " & _
   "FROM Detalle_DocumentoVenta INNER JOIN (Producto INNER JOIN Unidad ON Producto.cunidad=Unidad.cunidad) ON Detalle_DocumentoVenta.cProducto=Producto.cProducto WHERE (Detalle_DocumentoVenta.cDocumentoVenta='" & numero & "' AND Detalle_DocumentoVenta.doc_cod='" & doc_cod & "' AND Detalle_DocumentoVenta.sSerie='" & serie & "' AND Detalle_DocumentoVenta.Alm_Cod='" & Alm & "')"
    Call ConfiguraRst(strCadena)
    
        If rst.RecordCount <= 0 Then
            MsgBox "No Hay Productos Registrados", vbInformation, KEY_EMPRESA
            Exit Sub
        End If
        rst.MoveFirst
        Peso = 0
            For j = 0 To rst.RecordCount - 1
                codigo = Mid((rst(0)) + Space(50), 1, 4)
                Und = Mid((rst("Unidad")) + Space(50), 1, 4)
                descripcion = Mid(rst("Producto") + Space(80), 1, 70)
                cantidad = Mid(Format(str(rst("Cantidad")), "###,##0.00") + Space(10), 1, 4)
                Toneladas = rst("Peso")
                'Format(Str((Rst(3) * Val(Cantidad) / 1000)), "###,##0.00")
                Peso = Peso + Toneladas
                PesoFormato = Format(Toneladas, "#,##0.00")
                Printer.Print Tab(-10); codigo & Space(4) & descripcion & Space(16) & Und & Space(4) & PesoFormato & Space(9) & cantidad
                Printer.CurrentY = Printer.CurrentY + 0.2
                rst.MoveNext
            Next j
            inc = 0.5
            ' Printer.FontBold = True
            Printer.Print Tab(14); "NUEVA DF:" & KEY_DIRECCION
            Do While (Val(Printer.CurrentY) <= 35)
                Printer.CurrentY = Printer.CurrentY + inc
                            Loop
    rst.MoveFirst
    PesoTotalForm = Format(str(Peso), "#,##0.00")
    Printer.Print Tab(48); "PESO TOTAL ->"; Space(55) & PesoTotalForm + Space(2) + "Kg."
    Printer.Print "" '29---------------------------------------------------------------------"
    Printer.Print "" '30---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "31---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "32---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "33---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "33 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "34---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "35---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "36---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "37---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "38---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "39---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "40---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "41---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "42---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "43---------------------------------------------------------------------"
    Printer.Print Tab(10); Mid(razon_transporte + Space(80), 1, 65); '49-------
    Printer.Print Tab(10); Mid(ruc_transporte + Space(80), 1, 65);   '50-------
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print Tab(2); Mid(dir_transporte + Space(80), 1, 40);   '51-------
    Printer.Print Tab(2); Mid(dir_transporte + Space(80), 41, 50);   '51-------
    Printer.Print "" 'Tab(10); "52---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "53---------------------------------------------------------------------"
    'Printer.Print "" 'Tab(10); "54---------------------------------------------------------------------"
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Print Tab(-10); Right(Trim(FrmVentas.TxtSeri_guia.Text), 3) + "-" + Right(FrmVentas.TxtNumero_guia.Text, 7) & Space(10) & FrmVentas.DtcComprobanteGuia.Text
    Printer.EndDoc
    Exit Sub
End If
End Sub
Public Function funSeleccionImpresora(ByVal sNombreImpresora As String) As Boolean
Dim bEncontrada As Boolean
Dim i As Integer
'Selecciona la impresora para imprimir, si no puede seleccionarla devuelve false

bEncontrada = False

For i = 0 To Printers.Count - 1
    If UCase(Trim(Printers(i).DeviceName)) = UCase(Trim(sNombreImpresora)) Then
        KEY_PRINTER = Printers(i).DeviceName
        bEncontrada = True
        Exit For
    End If
Next i

If bEncontrada = False Then
    For i = 0 To Printers.Count - 1
    
    If UCase(Trim(Printers(i).DeviceName)) = Trim(Mid(UCase(Trim(sNombreImpresora)), 16, 50)) Then
        KEY_PRINTER = Printers(i).DeviceName
        bEncontrada = True
        Exit For
    End If
Next i
End If

If bEncontrada Then
    Set Printer = Printers(i)
    funSeleccionImpresora = True
Else
    funSeleccionImpresora = False
End If

End Function

Public Sub impresion_pedido(ByVal id_movimiento As Double)
Dim ptotal As Double, direccion_destino As String, impresiones As Integer
Dim nn As String, id_agenda As Double, horaf As String, fechad As String
Dim strvendedor As String
   'Call CargaDefConfigEpsonTM

    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
  ' Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    If KEY_IMPRESORA = "si" Then
       Call funSeleccionImpresora(KEY_PRINTER)            'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Else
       Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    End If
    'Printer.DeviceName = Trim("EPSON TM-U220 Receipt")
    'If Prt.DeviceName = "EPSON TM-U220 Receipt" Then
     '       Set Printer = Prt
      '  End If
        
   ' If id_doc <> "0109" Then
    '    Printer.Font.name = "control"
     '   Printer.Print "A"
    'End If
    'Tahoma
    Printer.Font.name = "Tahoma"
    Printer.Font.Size = "15"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    Printer.Print ""
    'Printer.Font.Bold = True
    'Printer.Font.Size = "30"
    Printer.Print "" 'Tab(10); 'L 2
    'Printer.Font.Size = "10"
    Printer.Print Tab(1); "---------------------------------------------"
    Printer.Print Tab(1); KEY_EMPRESA
    'Printer.Print Tab(1); Mid(KEY_DIRECCION, 1, 40)
    'Printer.Print Tab(1); "TELF  :  074-437529"
  
    Printer.Print Tab(1); "---------------------------------------------"
    
    Printer.Font.name = "3 of 9 Barcode"
    Printer.Font.Size = "30"
    Printer.Font.name = "Tahoma"
    Printer.Font.Size = "12"
    Printer.Print Tab(1); "ORDEN PEDIDO Nº:" & formato_item(id_movimiento, 10)
    Printer.Font.name = "Tahoma"
    Printer.Font.Size = "8"
    strCadena = "SELECT * FROM movimiento_venta v,movimiento_venta_detalle d,producto WHERE v.id_venta=d.id_venta and "
    strCadena = "SELECT * FROM view_movimientos where id_venta='" & id_movimiento & "' and ruc='" & KEY_RUC & "' "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        strvendedor = rst("vendedor")
    End If
    Printer.Print Tab(1); Mid("PROFORMA" + Space(20), 1, 20) & ":" & Space(2) & rst("documento")
     Printer.Print ""
    Printer.Print Tab(1); Mid("FECHA EMISION" + Space(20), 1, 20) & ":" & Space(2) & KEY_FECHA
    Printer.Print Tab(1); Mid("HORA IMPRESION" + Space(20), 1, 19) & ":" & Space(2) & Format(Time, "hh:mm:ss am/pm")

    Printer.Print Tab(1); "----------------------------------------------------"
    Printer.Print Tab(1); Mid("DNI    " + Space(20), 1, 20) & "   :" & Space(2) & rst("id_cliente")
    Printer.Print Tab(1); Mid("NOMBRE" + Space(20), 1, 19) & ":" & Space(2) & Mid(rst("ncliente"), 1, 50)
    
    
    strCadena = "SELECT descripcion FROM almacen WHERE id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstZ(strCadena)
    Printer.Print Tab(1); "----------------------------------------------------"
    Printer.Print Tab(1); Mid("TIENDA" + Space(20), 1, 20) & ":" & Space(2) & rstZ("descripcion")
    Printer.Print Tab(1); "----------------------------------------------------"
   
    rst.MoveFirst
    nntotal = 0
    Printer.Print Tab(0); Mid("CANT" + Space(20), 1, 6) & Mid("DESCRIPCION" + Space(20), 1, 40) & Space(2) & Mid("MARCA" + Space(20), 1, 10)
    For i = 0 To rst.RecordCount - 1
       Printer.Print Tab(0); Mid(str(rst("cantidad")) + Space(20), 1, 6) & Mid(rst("nombre_prod") + Space(20), 1, 38) & Space(1) & Mid(rst("marca") + Space(20), 1, 10)
       Printer.Print Tab(6); Mid("P.UNITARIO" + Space(20), 1, 5) & ":" & Mid(Format(rst("precio"), "#,##0.00") + Space(20), 1, 10) & Space(2) & "P.TOTAL :" & Mid(Format(rst("total"), "#,##0.00") + Space(20), 1, 8)
       nntotal = nntotal + rst("total")
       rst.MoveNext
    Next i
    Printer.Print Tab(0); "-----------------------------------------"
    Printer.Print Tab(6); "MONTO A PAGAR       S/.  :" & Mid(Format(nntotal, "#,##0.00") + Space(20), 1, 8)
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print Tab(0); "-----------------------------------------"
    
    Printer.Print Tab(1); "ATENDIDO POR:" + Space(1) + strvendedor + Space(3) + str(Time)
    Printer.Print "" 'Tab(10); 'L 9
    
    Printer.Print Tab(0); " ESPERE A SER LLAMADO PARA SER ATENDIDO "
   
    Printer.EndDoc
    Exit Sub
   ' GoTo ss
    
End Sub

Public Sub impresion_cuotas_credito(ByVal in_venta As Double)
   'Call CargaDefConfigEpsonTM

    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
  ' Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    If KEY_IMPRESORA = "si" Then
       Call funSeleccionImpresora(KEY_PRINTER)            'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Else
       Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    End If
    'Printer.DeviceName = Trim("EPSON TM-U220 Receipt")
    'If Prt.DeviceName = "EPSON TM-U220 Receipt" Then
     '       Set Printer = Prt
      '  End If
        
   ' If id_doc <> "0109" Then
    '    Printer.Font.name = "control"
     '   Printer.Print "A"
    'End If
    'Tahoma
    Printer.Font.name = "Tahoma"
    Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    Printer.Print ""
    'Printer.Font.Bold = True
    'Printer.Font.Size = "30"
    Printer.Print "" 'Tab(10); 'L 2
    'Printer.Font.Size = "10"
    Printer.Print Tab(1); "---------------------------------------------"
    Printer.Print Tab(1); KEY_EMPRESA
    'Printer.Print Tab(1); Mid(KEY_DIRECCION, 1, 40)
    'Printer.Print Tab(1); "TELF  :  074-437529"
  
    Printer.Print Tab(1); "---------------------------------------------"
    
    Printer.Font.name = "3 of 9 Barcode"
    Printer.Font.Size = "30"

    Printer.Print Tab(1); in_venta
    Printer.Font.name = "Tahoma"
    Printer.Font.Size = "8"
    Printer.Print Tab(1); "----------------------------------------------------"
    
    Printer.Print Tab(1); ""
    strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & in_venta & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    Printer.Print Tab(0); Mid(KEY_DIRECCION, 1, 36)
    Printer.Print Tab(0); Mid(KEY_DIRECCION, 37, 36)
    'If KEY_ALM <> "00001" Then
       Printer.Print Tab(1); "-------------------------------------------------------------"
       Printer.Print Tab(0); "SUC:" & KEY_DIRECCION_ALM
       Printer.Print Tab(1); "-------------------------------------------------------------"
   ' End If
    
    Printer.Print Tab(0); "TELF:" & KEY_TELEFONO
    Printer.Print Tab(0); "E-MAIL  :" & KEY_EMAIL
        Printer.Print Tab(0); "RUC   :" & KEY_RUC
    Printer.Print Tab(0); "FECHA EMISION :" & Format(rst("fecha_emision"), "dd-mm-YYYY")
    Printer.Print Tab(1); "-------------------------------------------------------------"
    Printer.Print Tab(0); Mid("REFERENCIA   :" + Space(40), 1, 20) + rst("documento")
    Printer.Print Tab(0); Mid("DNI/RUC      :" + Space(40), 1, 20) + rst("id_cliente")
    Printer.Print Tab(0); Mid("NOMBRE/RAZON :" + Space(40), 1, 20) + rst("ncliente")
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "-------------------------------------------------------------"
    Printer.Print Tab(0); Mid("TAZA INTERES :" + Space(40), 1, 20) + str(rst("interes"))
    Printer.Print Tab(0); Mid("NRO CUOTAS   :" + Space(40), 1, 20) + str(rst("cuotas"))
    Printer.Print Tab(1); ""
    End If
    strCadena = "SELECT * FROM movimiento_venta WHERE id_referencia='" & in_venta & "' ORDER BY fecha_vencimiento ASC "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        in_saldo = 0
        Printer.Print Tab(1); "===================================="
        For i = 0 To rst.RecordCount - 1
            Printer.Print Tab(1); "LETRA :" & rst("numero") & Space(2) & Format(rst("fecha_vencimiento"), "dd-mm-YYYY") & Space(2) & Format(rst("total"), "###0.00") & Space(2) & Format(rst("saldo"), "###0.00")
            Printer.Print Tab(1); "----------------------------------------------------"
            in_saldo = in_saldo + rst("saldo")
            rst.MoveNext
        Next i
        Printer.Print ""
        Printer.Print ""
        Printer.Print Tab(1); "----------------------------------------------------"
        Printer.Print Tab(1); "TOTAL ADEUDADO                          :" & Format(in_saldo, "###0.00")
        Printer.Print Tab(1); "----------------------------------------------------"
      
    End If
    Printer.Print ""
    Printer.EndDoc
    Exit Sub
   ' GoTo ss
    
End Sub

Public Sub guia_remision_tiketera(ByVal in_transferencia As String)
Dim ptotal As Double, direccion_destino As String, impresiones As Integer
Dim nn As String
Dim in_venta As String
   'Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    
    If KEY_IMPRESORA = "si" Then
        Call funSeleccionImpresora(KEY_PRINTER)            'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Else
         Printer.TrackDefault = True 'siempre apunta a la impresora predeter
         
         
         
    End If
    Printer.Font.name = "FontB11"
    Printer.Font.Size = "8"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM movimiento_transferencia WHERE id_transferencia='" & in_transferencia & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    
    '***** END COMPROBANTE***
    
    '**** DETALLE COMPROBANTE ***
    
    Printer.Print ""
    Printer.Print "" 'Tab(10); 'L 2
    
    
    Printer.Print Tab(1); KEY_EMPRESA
    Printer.Print Tab(1); KEY_DIRECCION
    Printer.Print Tab(1); "TELF  :  074-437529"
    Printer.Print Tab(1); "RUC   :" & KEY_RUC
    Printer.Print Tab(1); "FECHA :" & str(rst("fecha"))
    Printer.Print Tab(1); "-----------------------------------------------------------------------------"
    Printer.Print Tab(1); Mid("GUIA DE REMISION" + Space(20), 1, 20) & ": " & rst("serie") + "-" + rst("numero")
    Printer.Print Tab(1); "-----------------------------------------------------------------------------"
    
    
    Printer.Print Tab(0); "DATOS DEL EMISOR    ::::::::::::::::::::::::::: "
    Printer.Print Tab(1); Mid("RUC" + Space(20), 1, 12) & ":" & KEY_RUC
    Printer.Print Tab(0); Mid("NOMBRE RAZON" + Space(20), 1, 12) & ":" & KEY_EMPRESA
    Printer.Print Tab(0); Mid("PUNTO DE PARTIDA" + Space(20), 1, 12) & ":" & KEY_DIRECCION
    Printer.Print Tab(1); "---------------------------------------------------------------"
    
    
    Printer.Print Tab(0); "DATOS DEL DESTINATARIO    ::::::::::::::::::::::::::: "
    Printer.Print Tab(1); Mid("RUC" + Space(20), 1, 12) & ":" & rst("id_destinatario")
    Printer.Print Tab(0); Mid("NOMBRE RAZON" + Space(20), 1, 12) & ":" & rst("destinatario")
    Printer.Print Tab(0); Mid("PUNTO DE PARTIDA" + Space(20), 1, 12) & ":" & in_ubigeuo2 = Mid(UCase(get_ubigueo_persona(rst("id_destinatario"), rst("id_direccion"))) & Space(100), 1, 60)
    Printer.Print Tab(1); "---------------------------------------------------------------"
    
    
    Printer.Print Tab(0); "DOCUMENTO RELACIONADO ::::::::::::::::::::::::::: "
    
    Printer.Print Tab(1); Mid("RUC" + Space(20), 1, 12) & ":" & rst("id_destinatario")
    Printer.Print Tab(0); Mid("NOMBRE RAZON" + Space(20), 1, 12) & ":" & rst("destinatario")
    Printer.Print Tab(0); Mid("PUNTO DE PARTIDA" + Space(20), 1, 12) & ":" & in_ubigeuo2 = Mid(UCase(get_ubigueo_persona(rst("id_destinatario"), rst("id_direccion"))) & Space(100), 1, 60)
    Printer.Print Tab(1); "---------------------------------------------------------------"
    
    
             
    strCadena = "SELECT P.id_producto,P.nombre_prod,P.concentracion,M.cantidad,P.presentacion,P.presentacion_und,F.abreviatura FROM movimiento_transferencia_detalle M,producto P,unidad F  WHERE P.id_unidad=F.id_und AND   M.id_producto=P.id_producto and M.ruc=P.ruc and P.`ruc`=F.id_usu  and   M.id_transferencia='" & in_transferencia & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       For j = 0 To rstT.RecordCount - 1
                codigo = Mid((rstT("id_producto")) + Space(10), 1, 7)
                descripcion = Mid(rstT("nombre_prod") + Space(80), 1, 50)
                concentracion = Mid(rstT("concentracion") + Space(80), 1, 15)
                cantidad = Mid(str(rstT("cantidad")) + Space(80), 1, 10)
                Printer.Print Tab(0); "CODIGO :" & codigo
                Printer.Print Tab(0); descripcion
                strCadena = "SELECT s.motor,s.chasis,s.anio_fabricacion,c.descripcion as color,l.descripcion,s.nro_dua,s.nro_item FROM movimiento_transferencia_series s,producto p,linea_sub l,`imp_color` c   WHERE s.`id_producto`=p.`id_producto` and p.`id_sublinea`=l.`id_tipo` and  s.`ruc`=p.`ruc` and p.`ruc`=l.`id_usu` and p.`id_color`=c.`id_color` and  s.id_producto='" & rstT("id_producto") & "' and  id_transferencia='" & rst("id_transferencia") & "' and s.ruc='" & KEY_RUC & "'"
                Call ConfiguraRstZ(strCadena)
         If rstZ.RecordCount > 0 Then
            rstZ.MoveFirst
            Printer.Print Tab(25); "**************************"
            'Printer.Print Tab(25); "COLOR           :" & Space(2) & rstT("color")
            'Printer.Print Tab(25); "MODELO          :" & Space(2) & rstT("modelo")
            Printer.Print "" 'Tab(10); 'L 8
                
            For m = 0 To rstZ.RecordCount - 1
                
                Printer.Print Tab(5); "ITEM   :" & Space(2) & str(m + 1)
                Printer.Print Tab(10); "COLOR           :" & Space(2) & rstZ("color")
                Printer.Print Tab(10); "MODELO          :" & Space(2) & rstZ("descripcion")
                Printer.Print Tab(10); "Nº MOTOR         :" & Space(2) & rstZ("motor")
                Printer.Print Tab(10); "N°CHASIS         :" & Space(2) & rstZ("chasis")
                Printer.Print Tab(10); "N°DUA         :" & Space(2) & rstZ("nro_dua")
                Printer.Print Tab(10); "N° ITEM         :" & Space(2) & rstZ("nro_item")
                Printer.Print Tab(10); "AÑO              :" & Space(2) & rstZ("anio_fabricacion")
                rstZ.MoveNext
            Next m
         End If
                
                Printer.Print Tab(1); "-------------------------------------------------------------"
                rstT.MoveNext
            Next j
           
        End If
    
          
         
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print ""
    
strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   If rstL("igv") > 0 Then
      Printer.Print Tab(0); Mid("AFECTO" + Space(20), 1, 20) & ":" & rstL("valor_venta")
      'Printer.Print Tab(0); Mid("DSCTO" + Space(20), 1, 20) & ":" & rstL("descuento")
      'Printer.Print Tab(0); Mid("V.VENTA" + Space(20), 1, 20) & ":" & rstL("descuento")
   End If
   Printer.Print ""
   Printer.Print Tab(0); Mid("TOTAL COMPROBANTE" + Space(20), 1, 20) & ":" & rstL("total")
   Printer.Print ""
End If

    
      Printer.Print Tab(0); "TIPO TRASLADO :****** VENTA"
    
'    Printer.Print Tab(1); "ATENDIDO POR:" + Space(1) + KEY_VENDEDOR + Space(3) + Format(rst("hora"), "HH:mm")
    'Printer.Print "" 'Tab(10); 'L 9
    'Printer.Print Tab(1); "CAJA        :" + Space(1) + get_descripcion_alm(rst("id_alm_origen"))
    Printer.Print "" 'Tab(10); 'L 9'
    Printer.Print Tab(3); " EXCELENCIA A SU SERVICIO"
    'Printer.Print Tab(3); "  PARA SER CONSUMIDOS EN LA MISMA"
    'Printer.Print Tab(3); "  GRACIAS POR SU COMPRA"
    Printer.EndDoc
    
    'Call FrmVentas.cerrar
    Exit Sub
End Sub



Public Sub orden_salida_tiketera(ByVal in_orden As String)

    Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    
    If KEY_IMPRESORA = "si" Then
        Call funSeleccionImpresora(KEY_PRINTER)            'Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Else
         Printer.TrackDefault = True 'siempre apunta a la impresora predeter
         
         
         
    End If
    Printer.Font.name = "FontB11"
    Printer.Font.Size = "8"
    Printer.Font.Bold = False
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
    '***** COMPROBANTE *****
    strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & in_orden & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    n_venta = rst("id_comprobante")
    Printer.Print ""
    Printer.Print "" 'Tab(10); 'L 2
    
    
    Printer.Print Tab(1); KEY_EMPRESA
    Printer.Print Tab(1); KEY_DIRECCION
    Printer.Print Tab(1); "RUC   :" & KEY_RUC
    Printer.Print ""
    Printer.Print Tab(1); "-----------------------------------------------------------------------------"
    Printer.Print Tab(1); rst("documento")
    Printer.Print Tab(1); "FECHA :" & Format(rst("fecha_emision"), "dd-mm-YYYY")
    Printer.Print Tab(1); "-----------------------------------------------------------------------------"
    Printer.Print ""
    
    Printer.Print Tab(0); "DATOS CLIENTE    ::::::::::::::::::::::::::: "
    Printer.Print Tab(1); "---------------------------------------------------------------"
    Printer.Print Tab(1); Mid("DNI/RUC" + Space(20), 1, 12) & ":" & rst("id_cliente")
    Printer.Print Tab(0); Mid("NOMBRE RAZON" + Space(20), 1, 12) & ":" & rst("ncliente")
    Printer.Print Tab(0); Mid("DIRECCION" + Space(20), 1, 12) & ":" & rst("direccion")
    Printer.Print Tab(1); "---------------------------------------------------------------"
    Printer.Print ""
    Printer.Print Tab(0); "DOCUMENTO RELACIONADO ::::::::::::::::::::::::::: "
    Printer.Print ""
    If rst("id_guia") > 0 Then
        in_tipo = "03"
    Else
        in_tipo = "01"
    End If
    Printer.Print Tab(1); Mid("REF" + Space(20), 1, 12) & ":" & get_comprobante_orden_salida(rst("id_comprobante"), in_tipo)
    Printer.Print Tab(1); "---------------------------------------------------------------"
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(0); "LISTADO DE PRODUCTOS  ::::::::::::::::::::::::::: "
    Printer.Print Tab(1); "---------------------------------------------------------------"
    strCadena = "SELECT * FROM view_movimiento_venta_detalle WHERE id_alm='" & rst("id_alm") & "' and id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       For j = 0 To rstT.RecordCount - 1
                codigo = "COD:" & Mid((rstT("id_producto")) + Space(7), 1, 7)
                descripcion = Mid(rstT("detalle") + Space(50), 1, 50)
                cantidad = Mid(str(rstT("cantidad")) + Space(10), 1, 10)
                Printer.Print Tab(0); codigo & descripcion
                Printer.Print Tab(0); "CANT:" & cantidad & rstT("unidad")
                Printer.Print Tab(0); "-------------------------------------------------------------"
                rstT.MoveNext
            Next j
           
        End If
    
          
         
    Printer.Print "" 'Tab(10); 'L 9
    Printer.Print ""
    
     
    
    Printer.Print Tab(0); "        RECIBI CONFORME        "
    Printer.Print Tab(0); ""
    Printer.Print Tab(0); ""
    Printer.Print Tab(0); "-------------------------------------------------------------"
    Printer.Print Tab(0); rst("ncliente")
    Printer.Print Tab(0); rst("id_cliente")
    Printer.Print Tab(0); ""
    Printer.Print Tab(0); ""
    Printer.Print Tab(0); ""
    Printer.Print Tab(0); "-------------------------------------------------------------"
    Printer.Print Tab(0); KEY_VENDEDOR
    Printer.Print Tab(0); "        OPERADOR DESPACHO        "
    Printer.Print Tab(0); ""
    Printer.Print Tab(0); ""

    Printer.Print Tab(0); "      GRACIAS POR SU COMPRA        "
    Printer.Print Tab(0); ""
    Printer.EndDoc
    End If
    'Call FrmVentas.cerrar
    Exit Sub
End Sub


Public Function Establecer_Impresora_predeterminada(ByVal NamePrinter As String) As Boolean
On Error GoTo ErrSub
      
    'Variable de referencia
    Dim obj_Impresora As Object
      
    'Creamos la referencia
    Set obj_Impresora = CreateObject("WScript.Network")
        obj_Impresora.setdefaultprinter NamePrinter
        
    Set obj_Impresora = Nothing
          
        'La función devuelve true y se cambió con éxito
        Establecer_Impresora = True
        'MsgBox "La impresora se cambió correctamente", vbInformation
    Exit Function
      
      
'Error al cambiar la impresora
ErrSub:
If Err.Number = 0 Then Exit Function
   Establecer_Impresora = False
   MsgBox "error: " & Err.Number & Chr(13) & "Description: " & Err.Description
   On Error GoTo 0
End Function


Public Sub printer_orden_salida(ByVal in_orden As String, ByVal in_venta As String, ByVal in_almacenero As String)
Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant


strCadena = "SELECT id_venta,fecha_emision,hora,documento,id_cliente,ncliente,direccion,'" & get_comprobante_venta(in_venta) & "',id_producto,detalle,cantidad,unidad,'" & get_persona(in_almacenero) & "' FROM view_orden_salida_detalle WHERE id_venta='" & Val(in_orden) & "'"
                   Call ConfiguraRst(strCadena)
                   If rst.RecordCount > 0 Then
                       arr(0, 1) = "vendedor_proforma"
                       arr(1, 1) = "vendedor_telefono"
                       arr(0, 2) = "--"
                       arr(1, 2) = "--"
                       
                       param = arr()
                       
                       Ans = ShowMultiReport(rst, "rptOrdenSalida", param, App.Path + "\Reportes\")

                   End If
                   
End Sub

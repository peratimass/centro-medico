Attribute VB_Name = "ModAjustes"
Public Sub verificar_versionsoft()
On Error GoTo salir
Call alerta_stock

strCadena = "SELECT version FROM version_empresa WHERE version<>'" & KEY_VERSION & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstVersion(strCadena)
If rstVersion.RecordCount > 0 Then
    If MsgBox("HAY UNA VERSION MAS RECIENTE DEL SOFTWARE." + Chr(13) + Chr(13) + "VERSION INSTALADA  :" + KEY_VERSION + Chr(13) + "VERSION ACTUAL       :" + str(rstVersion("version")) + Chr(13) + Chr(13) + "DESEA ACTUALIZAR AHORA ?", vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then
        Dim in_ruta_aplicacion As String
        in_ruta_aplicacion = App.Path & "\Vitekey Business.exe"
        Shell in_ruta_aplicacion, vbNormalFocus
        DoEvents
        End
    End If
End If
  
Exit Sub
salir:
End Sub

Public Function get_existe_comprobante_monto(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_monto As Double, ByVal in_anulado As String, ByVal in_ruc As String) As Boolean

Dim in_redondeo As Single

strCadena = "SELECT id_venta,total,documento,anulado,fecha_emision,tc,id_tipo FROM movimiento_venta WHERE id_doc='" & in_doc & "' and  serie='" & in_serie & "' and numero='" & in_numero & "' and ruc='" & in_ruc & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount < 1 Then
    MsgBox "NO EXISTE ESTE COMPROBANTE" + Chr(13) + in_doc + ":" + in_serie + "-" + in_numero + Chr(13) + "TOTAL COMPROBANTE: " + str(in_monto), vbInformation
    get_existe_comprobante_monto = False
Else
    
    If rstA("total") = 0 And in_anulado = "si" Then
        get_existe_comprobante_monto = True
    Else
        
        in_redondeo = Abs(rstA("total") - Round(in_monto, 2))
        If in_redondeo > 0.01 Then
            MsgBox "COMPROBANTE LOCAL :" + rstA("documento") + Space(2) + Format(rstA("total"), "#,##0.00") + Chr(13) + "COMPROBANTE SUNAT: " + rstA("documento") + Space(2) + Format(in_monto, "#,##0.00"), vbInformation
            
            
            If rstA("anulado") = "si" Or rstA("total") <> Round(in_monto, 2) Then
                    If MsgBox("DESEA ACTUALIZAR EL COMPROBANTE", vbYesNo + vbQuestion) = vbYes Then
                       MsgBox "RECUERDE VERIFICAR EL COMPROBANTE: " + rstA("documento"), vbInformation
                       in_total = in_monto
                       in_valor_venta = in_monto / (1 + KEY_IGV)
                       in_igv = in_total - in_valor_venta
                       
                       strCadena = "UPDATE movimiento_venta SET anulado='no',total='" & in_monto & "',valor_venta='" & in_valor_venta & "',igv='" & in_igv & "' WHERE id_venta='" & rstA("id_venta") & "' and ruc='" & in_ruc & "' LIMIT 1"
                       CnBd.Execute (strCadena)
                       
                       
                       n_fecha = rstA("fecha_emision")
                       id_venta = rstA("id_venta")
                       
                   
                   
                    strCadena = "SELECT * FROM con_documento WHERE Activo=1 and  IdEmpresaSis='" & in_ruc & "' and IdReferencia='" & id_venta & "' LIMIT 1"
                    Call ConfiguraRstAux(strCadena)
                    If rstAux.RecordCount > 0 Then
                         strCadena = "call CON_EliminarVenta('" & id_venta & "','" & KEY_USUARIO & "') "
                         CnBd.Execute (strCadena)
                         
                    End If
                    
                    
                    strCadena = "call ADM_servicios_generales('11','" & in_ruc & "','','','','','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "call P_insert_venta_agenda_test('" & id_venta & "')"
                    CnBd.Execute (strCadena)

                    strCadena = "call ADM_servicios_generales('12','" & in_ruc & "','','','','','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)

                    
                    
            End If
            End If
            get_existe_comprobante_monto = False
        Else
            get_existe_comprobante_monto = True
        End If
    End If
     
    
    
    
End If
End Function

Public Sub put_fecha_corte(ByVal in_ruc As String)

strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & in_ruc & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   in_fecha_fin = DateSerial(Year(KEY_FECHA), Val(Month(KEY_FECHA)) + 1, 1 - 1)
   in_fecha_corte = DateAdd("d", rstL("dias_prorroga"), in_fecha_fin)
   
   strCadena = "UPDATE entidad_empresa SET fecha_corte='" & Format(in_fecha_corte, "YYYY-mm-dd") & "' WHERE cod_unico='" & in_ruc & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
   CnBd.Execute (strCadena)
   
   strCadena = "UPDATE entidad_parametros SET caducidad='" & Format(in_fecha_corte, "YYYY-mm-dd") & "' WHERE cod_unico='" & in_ruc & "' LIMIT 1"
   CnBd.Execute (strCadena)
   
   
End If
End Sub
Public Sub alerta_stock()
On Error GoTo salir
If KEY_ALARMA_STOCK = "si" Then
    strCadena = "SELECT id_producto FROM view_producto WHERE stock<=stock_minimo and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraAlarma(strCadena)
    If rstAlarma.RecordCount > 0 Then
        MDIFrmPrincipal.StatusBar1.Panels(6) = ""
        MDIFrmPrincipal.StatusBar1.Panels(6).Picture = LoadPicture(App.Path & "\Imagenes\alarma_stock.jpg")
    Else
        MDIFrmPrincipal.StatusBar1.Panels(6) = ""
        MDIFrmPrincipal.StatusBar1.Panels(6).Picture = LoadPicture(App.Path & "\Imagenes\alarma_stock_on.jpg")
    End If
Else
End If
Exit Sub
salir:

End Sub

Public Sub actualizar(ByVal in_ruta As String, ByVal in_version_nueva As String, ByVal in_ruc As String)
Dim imagen As String
Dim str_ruta_img As String
Dim Archivo As String
Dim datos(3) As String

Archivo = App.Path & "\main.exe"
On Error GoTo sit

If VerificarArchivo(Archivo) = True Then
    Kill Archivo
End If
sit:
str_ruta_img = App.Path & "\main.exe"
DownloadFile in_ruta, str_ruta_img

strRuta_ini = App.Path & "\archivos\vitekey.ini"
fnum = FreeFile
    Open strRuta_ini For Input As fnum
    i = 0
    Do While Not EOF(fnum)
        
        Select Case i
            Case 0
                 Line Input #fnum, file_line
                  datos(0) = file_line
            Case 1
                 Line Input #fnum, file_line
                  datos(1) = file_line
            Case 2
                Line Input #fnum, file_line
                datos(2) = in_version_nueva
            Case 3
                Line Input #fnum, file_line
                datos(3) = in_ruc
        End Select
        i = i + 1
    Loop
    Close #fnum


    Open strRuta_ini For Output As fnum
    For i = LBound(datos) To UBound(datos)
        Print #1, datos(i)
    Next i
  Close #fnum
  
  

End Sub

Public Function DownloadFile(Url As String, LocalFilename As String) As Boolean

Dim lngRetVal As Long
lngRetVal = URLDownloadToFile(0, Url, LocalFilename, 0, 0)
If lngRetVal = 0 Then DownloadFile = True
End Function







Public Sub put_linea_cuenta(ByVal in_tipo As String, ByVal in_monto As Double, in_venta As String)




End Sub
Public Function get_existe_comprobante(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String) As Boolean
strCadena = "SELECT id_venta FROM movimiento_venta WHERE id_doc='" & in_doc & "' and  serie='" & in_serie & "' and numero='" & in_numero & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount < 1 Then
    get_existe_comprobante = False
Else
    
    get_existe_comprobante = True
End If
End Function


Public Function get_existe_comprobante_detallado(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String) As Boolean
Dim in_venta As String

strCadena = "SELECT id_venta,documento,fecha_emision FROM movimiento_venta WHERE id_doc='" & in_doc & "' and  serie='" & in_serie & "' and numero='" & in_numero & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount < 1 Then
    get_existe_comprobante_detallado = False
Else
    in_venta = rstA("id_venta")
    strCadena = "SELECT id_venta FROM movimiento_venta_detalle WHERE id_venta='" & rstA("id_venta") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount < 1 Then
        Call delete_asiento_venta_migracion(in_venta)
        get_existe_comprobante_detallado = False
    Else
        
        in_referencia = ""
        strCadena = "SELECT * FROM con_documento where idReferencia='" & in_venta & "' and IdEmpresasis='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount < 1 Then
           ' strCadena = "call P_insert_venta_agenda('" & in_venta & "')"
           ' CnBd.Execute (strCadena)
        End If
        'in_referencia = rstL("Id")
        
        
           ' strCadena = "SELECT * FROM con_asiento where Activo=1 and  idReferencia='" & in_referencia & "' and IdEmpresasis='" & KEY_RUC & "'"
           ' Call ConfiguraRstL(strCadena)
           ' If rstL.RecordCount < 1 Then
            '   Call delete_asiento_venta_migracion(in_venta)
            '    strCadena = "call P_insert_venta_agenda('" & in_venta & "')"
            '    CnBd.Execute (strCadena)
             '   Exit Function
           ' End If
        
        
        get_existe_comprobante_detallado = True
    End If
End If
End Function



Public Function get_dni_keyfacil(ByVal in_keyfacil As String) As String
If Len(in_keyfacil) < 1 Then
    get_dni_keyfacil = "00000000"
Else
    strCadena = "SELECT dni FROM persona WHERE id_keyfacil='" & Trim(in_keyfacil) & "' LIMIT 1"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount > 0 Then
        get_dni_keyfacil = rstA("dni")
    Else
        get_dni_keyfacil = "00000000"
    End If
End If
End Function

Public Function get_dni(ByVal in_nombre As String)

Dim nombres() As String
            
            nombres = Split(in_nombre, " ")
           ' nombre = nombres(0)
           ' paterno = nombres(1)
          '  materno = nombres(2)
          '  nombre_completo = paterno & Space(1) & materno & Space(1) & nombre
            
            
strCadena = "SELECT dni FROM persona WHERE nombre_completo LIKE '%" & UCase(in_nombre) & "%' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
    get_dni = rstA("dni")
Else
    get_dni = "00000000"
End If

End Function


Public Sub get_auto_pago_main(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_total As Double)
Dim in_cuenta_caja As String
Dim in_forma_pago_detalle As Integer
If in_doc = "0099" Then

strCadena = "SELECT * FROM movimiento_venta_monto_temporal WHERE id_doc='" & in_doc & "' and serie='" & in_serie & "' and id_usuario='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount < 1 Then

in_forma_pago_detalle = get_forma_pago_contado
in_cuenta_caja = get_cuenta_contable_caja(in_forma_pago_detalle)


strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,cuenta_contable,cuotas,id_usuario,id_recibo,detalle,banco,cheque,fecha,id_alm,ruc) VALUES " & _
       " ('" & in_doc & "','" & in_serie & "','" & in_numero & "','01','" & in_forma_pago_detalle & "','00001','" & in_total & "','" & in_total & "','00','-','-','" & in_cuenta_caja & "','0','" & KEY_USUARIO & "','0','-','-','-','" & KEY_FECHA & "','" & KEY_ALM & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
End If

End If
End Sub
Public Sub get_pago_keyfacil(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_total As Double, ByVal in_moneda As String, ByVal in_forma_pago As String, ByVal in_alm As String)
Dim in_cuenta_caja As String
Dim in_forma_pago_detalle As Integer


strCadena = "SELECT * FROM movimiento_venta_monto_temporal WHERE id_doc='" & in_doc & "' and serie='" & in_serie & "' and id_usuario='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount < 1 Then

in_forma_pago_detalle = get_forma_pago_contado_keyfacil(in_moneda, in_alm, in_forma_pago)
in_cuenta_caja = get_cuenta_contable_caja(in_forma_pago_detalle)


strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,cuenta_contable,cuotas,id_usuario,id_recibo,detalle,banco,cheque,fecha,id_alm,ruc) VALUES " & _
       " ('" & in_doc & "','" & in_serie & "','" & in_numero & "','01','" & in_forma_pago_detalle & "','00001','" & in_total & "','" & in_total & "','00','-','-','" & in_cuenta_caja & "','0','" & KEY_USUARIO & "','0','-','-','-','" & KEY_FECHA & "','" & KEY_ALM & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
End If


End Sub
Public Function get_precio_segmentacion(ByVal in_producto As String, ByVal in_dni As String) As Double
Dim in_tipo_cliente As String

strCadena = "SELECT id_tipo_cliente FROM entidad_empresa WHERE cod_unico='" & in_dni & "' and id_empresa='" & KEY_RUC & "'  LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
    in_tipo_cliente = rstA("id_tipo_cliente")
    strCadena = "SELECT * FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount > 0 Then
         Select Case in_tipo_cliente
                Case "01"
                    get_precio_segmentacion = rstA("precio_venta")
                Case "02"
                    get_precio_segmentacion = rstA("precio_alterno_a")
                Case "03"
                    get_precio_segmentacion = rstA("precio_mayor")
         End Select
    End If
    
End If


End Function


Public Function get_precio_propio(ByVal in_producto As String, ByVal in_dni As String, ByVal in_precio As Double) As Double
strCadena = "SELECT precio as monto FROM view_plan_servicio_persona WHERE id_producto='" & in_producto & "' and  dni='" & in_dni & "' and ruc='" & KEY_RUC & "'  LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   get_precio_propio = rstA("monto")
Else
    get_precio_propio = in_precio
  
End If
End Function

Public Function get_last_buy(ByVal in_producto As String) As String
strCadena = "SELECT * FROM view_last_date_buy WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
    get_last_buy = Format(rstA("fecha_emision"), "dd-mm-YYYY")
Else
    get_last_buy = Format("********", "dd-mm-YYYY")
End If


End Function



Public Sub put_estado_cuenta(ByVal in_dni As String)

strCadena = "DELETE FROM persona_estado_cuenta WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM movimiento_venta WHERE in_doc IN('0001','0003','0007','0054','0412','0000') id_cliente='" & in_dni & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_venta ASC "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       Select Case rst("id_doc")
              Case "0001" Or "0003"
                        
       End Select
   Next i
End If



End Sub




Public Sub Mayusculas(KeyAscii As Integer)
If KeyAscii >= 97 And KeyAscii <= 122 Then
    KeyAscii = KeyAscii - 32
End If
End Sub


 Function FormatosCeros(ByVal cod As String, ByVal longitud As Integer) As String
Dim X As Integer
Dim Formato As String
  Formato = ""
  For X = 1 To longitud
    Formato = Formato + "0"
  Next X
    cod = Format(Trim(str(Val(Right(cod, longitud)))), Formato)

FormatosCeros = cod



End Function
Public Sub ingresar_tramite(ByVal in_venta As Double)

End Sub

Public Function correlativo_comprobante(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String) As Boolean

strCadena = "SELECT numero FROM movimiento_venta where id_doc='" & in_doc & "' and  serie='" & in_serie & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    If Val(in_numero) <= Val(rstK("numero")) Then
        'MsgBox "Numero de Comprobante inferior al numero Actual", vbInformation, KEY_EMPRESA
        correlativo_comprobante = True
        Exit Function
    End If
Else
    correlativo_comprobante = False
End If

End Function
Public Sub generar_mantenimientos(ByVal in_venta As Double, ByVal in_dni As String, ByVal fecha_venta As Date)
Dim in_mantenimiento As Double
Dim nfecha As Date
strCadena = "SELECT * FROM movimiento_venta_mantenimiento WHERE id_venta='" & in_venta & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
     strCadena = "SELECT * FROM view_movimiento_venta WHERE id_venta='" & in_venta & "' and afecto_garantia='si' and mantenimientos>0 and ruc='" & KEY_RUC & "' "
     Call ConfiguraRstK(strCadena)
     If rstK.RecordCount > 0 Then
                 strCadena = "INSERT INTO movimiento_venta_mantenimiento(`id_venta`,`dni_cliente`,`id_producto`,`fecha_venta`,`dni_save`,`imp_detalle`,`ruc`)VALUES " & _
                "('" & in_venta & "','" & in_dni & "','" & rstK("id_producto") & "','" & Format(fecha_venta, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & rstK("id_detalle_serie") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                 
                strCadena = "SELECT * FROM movimiento_venta_mantenimiento ORDER BY id_mantenimiento DESC LIMIT 1"
                Call ConfiguraRst(strCadena)
                If rst.RecordCount > 0 Then
                    in_mantenimiento = rst("id_mantenimiento") ' ID MANTENIMIENTO
                End If
                
                strCadena = "SELECT * FROM linea_mantenimiento WHERE id_linea='" & rstK("id_linea") & "' and ruc='" & KEY_RUC & "' ORDER BY id_mantenimiento ASC"
                Call ConfiguraRst(strCadena)
                If rst.RecordCount > 0 Then
                        nfecha = fecha_venta
                        For i = 1 To rst.RecordCount - 1
                            nfecha = DateAdd("d", rst("dias"), nfecha)
                            strCadena = "INSERT INTO movimiento_venta_mantenimiento_listado(`id_mantenimiento`,`fecha_aproximada`,`id_venta`,`dni_save`,`ruc`)VALUES('" & in_mantenimiento & "','" & Format(nfecha, "YYYY-mm-dd") & "','" & in_venta & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                            CnBd.Execute (strCadena)
                             
                            strCadena = "SELECT * FROM movimiento_venta_mantenimiento_listado ORDER BY id DESC LIMIT 1"
                            Call ConfiguraRstL(strCadena)
                            Call generar_insumos(rst("id_mantenimiento"), rstL("id"))
                            rst.MoveNext
                        Next i
               End If
     End If
     
End If
Call reporte_matenimiento(in_venta)



End Sub
Private Sub reporte_matenimiento(ByVal in_venta As Double)
strCadena = "SELECT id_cliente,ncliente,documento,'-',nombre_prod,placa,nro_chasis,nro_motor,fecha_emision,fecha_mantenimiento,id,fecha_aproximada,nombre_completo,estado FROM  view_repor_mantenimiento WHERE id_venta='" & in_venta & "'"

Call ConfiguraRstK(strCadena)
Ans = ShowMultiReport(rstK, "RptMantenimientos", , App.Path + "\Reportes\")
End Sub
Private Sub generar_insumos(ByVal in_detalle As Double, ByVal in_listado As Double)
strCadena = "SELECT * FROM linea_mantenimiento_detalle WHERE id_mantenimiento='" & in_detalle & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    For i = 0 To rstL.RecordCount - 1
        strCadena = "INSERT INTO movimiento_venta_mantenimiento_insumos(`id_listado`,`id_producto`,`cantidad`,`pagado`,`ruc`) VALUES " & _
        " ('" & in_listado & "','" & rstL("id_producto") & "','" & rstL("cantidad") & "','" & rstL("pagado") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        rstL.MoveNext
    Next i
End If
End Sub
Public Sub generar_mantenimiento(ByVal in_venta As Double)
strCadena = "SELECT  FROM movimiento_venta_detalle d,linea l,producto p WHERE id_venta='" & in_venta & "' and d.id_producto=p.id_producto and p.id_linea=l.id_linea and p.ruc=l.id_usu and p.ruc='" & KEY_RUC & "'  "

End Sub
Public Sub manejador_error()
MsgBox "A ocurrido un Fallo en la Red" + Chr(13) + Chr(13) + "Disuclpe las molestias." + Chr(13) + KEY_VENDEDOR, vbInformation
Exit Sub
End Sub
Public Sub actualizar_credito(ByVal in_dni As String, ByVal monto_pagado As Single)

strCadena = "UPDATE entidad_empresa SET monto_credito=monto_credito+'" & monto_pagado & "' WHERE cod_unico='" & in_dni & "' and id_empresa='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
'
End Sub
Function DigitoVerificadorRUC(m_sRUC As String) As String
Dim m_sRUCC As Double
     If Not IsNumeric(m_sRUC) Then
        m_sErrValRUC = "El valor no es numérico"
     Else
        If Len(m_sRUC) <> 10 Then
            m_sErrValRUC = "Número de dígitos inválido"
        Else
        
            Dim dig01 As Integer
            Dim dig02 As Integer
            Dim dig03 As Integer
            Dim dig04 As Integer
            Dim dig05 As Integer
            Dim dig06 As Integer
            Dim dig07 As Integer
            Dim dig08 As Integer
            Dim dig09 As Integer
            Dim dig10 As Integer
            Dim dig11 As Integer
            Dim suma As Integer
            Dim residuo As Integer
            Dim resta As Integer
            Dim digChk As Integer
    
            dig01 = CInt(Mid$(m_sRUC, 1, 1)) * 5
            dig02 = CInt(Mid$(m_sRUC, 2, 1)) * 4
            dig03 = CInt(Mid$(m_sRUC, 3, 1)) * 3
            dig04 = CInt(Mid$(m_sRUC, 4, 1)) * 2
            dig05 = CInt(Mid$(m_sRUC, 5, 1)) * 7
            dig06 = CInt(Mid$(m_sRUC, 6, 1)) * 6
            dig07 = CInt(Mid$(m_sRUC, 7, 1)) * 5
            dig08 = CInt(Mid$(m_sRUC, 8, 1)) * 4
            dig09 = CInt(Mid$(m_sRUC, 9, 1)) * 3
            dig10 = CInt(Mid$(m_sRUC, 10, 1)) * 2
            'dig11 = CInt(Mid$(m_sRUC, 11, 1))
            
            suma = dig01 + dig02 + dig03 + dig04 + dig05 + dig06 + dig07 + dig08 + dig09 + dig10
            residuo = suma Mod 11
            resta = 11 - residuo
            
            If resta = 11 Then
                digChk = 1
            ElseIf resta = 10 Then
                digChk = 0
            Else
                digChk = resta
            End If
            dig11 = digChk
            
            If dig11 = digChk Then
                RUC_EsValido = True
            Else
                m_sErrValRUC = "El número de RUC no es válido"
            End If
          
        End If
    End If
DigitoVerificadorRUC = Trim(m_sRUC) & Trim(str(dig11))
End Function
Public Function RUC_valido(ByVal m_sRUC As String) As Boolean

    If Not IsNumeric(m_sRUC) Then
        m_sErrValRUC = "El valor no es numérico"
    Else
        If Len(m_sRUC) <> 11 Then
            m_sErrValRUC = "Número de dígitos inválido"
        Else
        
            Dim dig01 As Integer
            Dim dig02 As Integer
            Dim dig03 As Integer
            Dim dig04 As Integer
            Dim dig05 As Integer
            Dim dig06 As Integer
            Dim dig07 As Integer
            Dim dig08 As Integer
            Dim dig09 As Integer
            Dim dig10 As Integer
            Dim dig11 As Integer
            
            Dim suma As Integer
            Dim residuo As Integer
            Dim resta As Integer
            
            Dim digChk As Integer
    
            dig01 = CInt(Mid$(m_sRUC, 1, 1)) * 5
            dig02 = CInt(Mid$(m_sRUC, 2, 1)) * 4
            dig03 = CInt(Mid$(m_sRUC, 3, 1)) * 3
            dig04 = CInt(Mid$(m_sRUC, 4, 1)) * 2
            dig05 = CInt(Mid$(m_sRUC, 5, 1)) * 7
            dig06 = CInt(Mid$(m_sRUC, 6, 1)) * 6
            dig07 = CInt(Mid$(m_sRUC, 7, 1)) * 5
            dig08 = CInt(Mid$(m_sRUC, 8, 1)) * 4
            dig09 = CInt(Mid$(m_sRUC, 9, 1)) * 3
            dig10 = CInt(Mid$(m_sRUC, 10, 1)) * 2
            dig11 = CInt(Mid$(m_sRUC, 11, 1))
            
            suma = dig01 + dig02 + dig03 + dig04 + dig05 + dig06 + dig07 + dig08 + dig09 + dig10
            residuo = suma Mod 11
            resta = 11 - residuo
            
            If resta = 11 Then
                digChk = 1
            ElseIf resta = 10 Then
                digChk = 0
            Else
                digChk = resta
            End If
            
            
            If dig11 = digChk Then
                RUC_valido = True
            Else
                RUC_valido = False
            End If
          
        End If
    End If
End Function
Public Function formato_fecha(ByVal strfecha As String) As String
formato_fecha = Format(strfecha, "YYYY-mm-dd")
End Function
Public Function get_marca(ByVal id_producto As String)
strCadena = "SELECT m.descripcion as marca FROM producto p,marca m WHERE p.id_producto='" & id_producto & "' and  p.id_marca=m.id_marca and p.ruc=m.id_usu and p.ruc='" & KEY_RUC & "'"
Call ConfiguraRstF(strCadena)
If rstF.RecordCount > 0 Then
    get_marca = rstF("marca")
Else
    get_marca = "----"
End If
End Function
Public Sub actualizar_cloud()
Call conexion_cloud
strCadena = "SELECT * FROM entidad_acciones WHERE actualizado='no' and  ruc='" & KEY_RUC & "' ORDER BY id ASC"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   rstLocal.MoveFirst
   For i = 0 To rstLocal.RecordCount - 1
       CnBd2.Execute (Replace(rstLocal("cadena"), "´", "'"))
       DoEvents
       strCadena = "UPDATE entidad_acciones SET actualizado='si' WHERE id='" & rstLocal("id") & "'"
       CnBd.Execute (strCadena)
       
       rstLocal.MoveNext
       DoEvents
   Next i
End If
'MDIFrmPrincipal.timer_cloud.Enabled = True
CnBd2.close
End Sub
Public Function numero_registros_cloud(ByVal in_registros As Double) As Double
Dim Num_registros As String
strRuta_reloj = App.Path & "\archivos\vitekeycloud.ini"
FileName = Dir(strRuta_reloj)
If FileName = "" Then
    Num_registros = Trim(str(in_registros))
    Open strRuta_reloj For Output As #1
    Print #1, Num_registros
    Close #1
    
Else
    Open strRuta_reloj For Input As #1
    Line Input #1, Num_registros
    Close #1
End If

numero_registros_cloud = Val(Num_registros)

End Function
Public Sub listar_telefono(ByVal Grilla As MSHFlexGrid, ByVal in_dni As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT * from view_telefono WHERE dni='" & in_dni & "'"
Call ConfiguraRstT(strCadena)

If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1500
            Grilla.ColWidth(2) = 1250
            Grilla.ColWidth(3) = 1400
        Next
        cabecera = "REFERENCIA" & vbTab & "REFERENCIA" & vbTab & "TELEFONO"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstT.MoveFirst
        
        For i = 0 To rstT.RecordCount - 1
            Fila = rstT("id_telefono") & vbTab & rstT("referencia") & vbTab & rstT("descripcion") & vbTab & rstT("telefono")
            Grilla.AddItem Fila
            rstT.MoveNext
        Next i
    Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"

End Sub

Public Sub update_saldo_inicial(ByVal in_producto As String)
strCadena = "SELECT * FROM ginsac WHERE id_producto='" & in_producto & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
            in_alm1 = "00001"
            in_alm2 = "00002"
            strInventario = Format(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), "000000")
            in_fecha = "2018-01-01"
            strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
            "('" & strInventario & "','" & in_producto & "','0106','001','" & strInventario & "','" & rstZ("precio_costo") & "','2018-01-01','" & in_alm1 & "','" & rstZ("stock_alm1") & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        
            strCadena = "call put_kardex_stock_vitekey('06','" & Format(in_fecha, "YYYY-mm-dd") & "','" & Val(strInventario) & "','0106','001','" & strInventario & "','" & KEY_RUC & "','" & in_producto & "','" & rstZ("stock_alm1") & "','" & rstZ("precio_costo") & "','" & in_alm1 & "','" & KEY_RUC & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        
        'Almacen 2
        
            strInventario = Format(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), "000000")
            in_fecha = "2018-01-01"
            
            strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
            "('" & strInventario & "','" & in_producto & "','0106','001','" & strInventario & "','" & rstZ("precio_costo") & "','2018-01-01','" & in_alm2 & "','" & rstZ("stock_alm2") & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        
            strCadena = "call put_kardex_stock_vitekey('06','" & Format(in_fecha, "YYYY-mm-dd") & "','" & Val(strInventario) & "','0106','001','" & strInventario & "','" & KEY_RUC & "','" & in_producto & "','" & rstZ("stock_alm2") & "','" & rstZ("precio_costo") & "','" & in_alm2 & "','" & KEY_RUC & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
End If
End Sub
Public Sub update_saldo_inicial_vargas(ByVal in_producto As String)
strCadena = "SELECT * FROM inventario_vargas_31122017 WHERE id_producto='" & Format(in_producto, "000000") & "'"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
            in_alm1 = "00001"
            
            strInventario = Format(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), "000000")
            in_fecha = "2017-12-31"
            strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
            "('" & strInventario & "','" & in_producto & "','0106','001','" & strInventario & "','" & rstZ("costo") & "','2018-01-01','" & in_alm1 & "','" & rstZ("cantidad") & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        
            strCadena = "call put_kardex_stock_vitekey('06','" & Format(in_fecha, "YYYY-mm-dd") & "','" & Val(strInventario) & "','0106','001','" & strInventario & "','" & KEY_RUC & "','" & in_producto & "','" & rstZ("cantidad") & "','" & rstZ("costo") & "','" & in_alm1 & "','" & KEY_RUC & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        
        
End If
End Sub
Public Sub update_kardex_all(ByVal in_producto As String)
    Dim in_dias As Integer
    
    If verificacion_servicio(in_producto) = True Then
        Exit Sub
    End If
    
    

    strCadena = "DELETE FROM kardex WHERE fecha_emision>='" & Format(get_fecha_periodo_abierto, "YYYY-mm-dd") & "' and    id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    If Format(get_fecha_periodo_abierto, "YYYY-mm-dd") = "2018-01-01" Then
        Call update_saldo_inicial(in_producto)
        in_fechai = "2018-01-01"
    Else
        in_fechai = Format(get_fecha_periodo_abierto, "YYYY-mm-dd")
    End If
    
  
            in_dias = DateDiff("d", in_fechai, KEY_FECHA)
            For m = 0 To in_dias
                If Format(in_fechai, "YYYY-mm-dd") <= KEY_FECHA Then
                
                'COMPRAS
                strCadena = "select * from view_kardex_compra_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    If KEY_RUC = "20128836251" Then
                        Call compras_producto_vargas(in_fechai, in_producto)
                    Else
                        Call compras_producto(in_fechai, in_producto)
                    End If
                    
                End If
               
               
                
                ' transferencias ingreso
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
               ' DoEvents
               
                'ventas salida
       
                
                strCadena = "select * from view_kardex_ventas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call ventas_producto(in_fechai, in_producto)
                End If
                
              
                'DoEvents
                'notas salida
                
                strCadena = "select * from view_kardex_notas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call notas_producto(in_fechai, in_producto)
                End If
                
                End If
                

                in_fechai = DateAdd("d", 1, in_fechai)
                
                
            Next m
    
       
       
        


End Sub
Public Sub update_kardex(ByVal in_producto As String)
    Dim in_dias As Integer
    
    If verificacion_servicio(in_producto) = True Then
        Exit Sub
    End If
    
    

    strCadena = "DELETE FROM kardex WHERE fecha_emision>='" & Format(get_fecha_periodo_abierto, "YYYY-mm-dd") & "' and    id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    If Format(get_fecha_periodo_abierto, "YYYY-mm-dd") = "2018-01-01" Then
        Call update_saldo_inicial(in_producto)
        in_fechai = "2018-01-01"
    Else
        in_fechai = Format(get_fecha_periodo_abierto, "YYYY-mm-dd")
    End If
    
            in_dias = DateDiff("d", in_fechai, KEY_FECHA)
            in_progress.Min = 0
            in_progress.Max = in_dias
            
            For m = 0 To in_dias
                If Format(in_fechai, "YYYY-mm-dd") <= KEY_FECHA Then
                
                'COMPRAS
                strCadena = "select * from view_kardex_compra_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    If KEY_RUC = "20128836251" Then
                        Call compras_producto_vargas(in_fechai, in_producto)
                    Else
                        Call compras_producto(in_fechai, in_producto)
                    End If
                    
                End If
               ' transferencias ingreso
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
               ' DoEvents
               
                'ventas salida
       
                
                strCadena = "select * from view_kardex_ventas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call ventas_producto(in_fechai, in_producto)
                End If
                
              
                'DoEvents
                'notas salida
                
                strCadena = "select * from view_kardex_notas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call notas_producto(in_fechai, in_producto)
                End If
                
                End If
                DoEvents
                in_progress.Value = m
                in_fechai = DateAdd("d", 1, in_fechai)
                
                
            Next m
    
       
       MsgBox "PROCESADO.....", vbInformation, KEY_VENDEDOR
        


End Sub
Public Sub update_kardex_update(ByVal in_producto As String, ByVal in_fecha As String)

Dim in_dias As Integer



strCadena = "DELETE FROM kardex WHERE fecha_emision>='" & Format(in_fecha, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)



   
        in_fechai = Format(CVDate(in_fecha), "YYYY-mm-dd")
  
    
            
            
            in_dias = DateDiff("d", in_fechai, KEY_FECHA)
            For m = 0 To in_dias
                If Format(in_fechai, "YYYY-mm-dd") <= KEY_FECHA Then
                
                'COMPRAS
                strCadena = "select id_compra from view_kardex_compra_existe WHERE fecha_kardex='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    If KEY_RUC = "20128836251" Then
                        Call compras_producto_vargas(in_fechai, in_producto)
                    Else
                        Call compras_producto(in_fechai, in_producto)
                    End If
                    
                End If
               
               
                
                ' transferencias ingreso
                strCadena = "select id_transferencia from movimiento_transferencia WHERE fecha='" & Format(in_fechai, "YYYY-mm-dd") & "'and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstC(strCadena)
                If rstc.RecordCount > 0 Then
                
                strCadena = "select * from view_transferencia_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                  Call transferencia_ingreso_producto(in_fechai, in_producto)
                  Call transferencia_salida_producto(in_fechai, in_producto)
                End If
               End If
               ' DoEvents
               
                'ventas salida
       
                
                strCadena = "select * from view_kardex_ventas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call ventas_producto(in_fechai, in_producto)
                End If
                
              
                'DoEvents
                'notas salida
                
                strCadena = "select * from view_kardex_notas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call notas_producto(in_fechai, in_producto)
                End If
                
                End If
                

                in_fechai = DateAdd("d", 1, in_fechai)
                
                
            Next m
    
       
       
        


End Sub


Public Sub update_kardex_internacional(ByVal in_producto As String, ByVal in_fecha As String)

Dim in_dias As Integer



strCadena = "DELETE FROM kardex WHERE fecha_emision>='" & Format(in_fecha, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)



   
        in_fechai = Format(CVDate(in_fecha), "YYYY-mm-dd")
  
    
            
            
            in_dias = DateDiff("d", in_fechai, KEY_FECHA)
            FrmCompras.progresbar_kardex.Min = 0
            FrmCompras.progresbar_kardex.Max = in_dias
            
            For m = 0 To in_dias
                If Format(in_fechai, "YYYY-mm-dd") <= KEY_FECHA Then
                
                'COMPRAS
                strCadena = "CALL ADM_verificar_kardex_inter('2','" & in_producto & "','" & Format(in_fechai, "YYYY-mm-dd") & "','" & KEY_RUC & "')"
                Call ConfiguraRstL(strCadena)
                If rstL(0) > 0 Then
                    Call compras_producto(in_fechai, in_producto)
                End If
               
               
               'ventas salida
       
                
                strCadena = "CALL ADM_verificar_kardex_inter('1','" & in_producto & "','" & Format(in_fechai, "YYYY-mm-dd") & "','" & KEY_RUC & "')"
                Call ConfiguraRstL(strCadena)
                If rstL(0) > 0 Then
                    Call ventas_producto(in_fechai, in_producto)
                End If
                
              
              
                

                in_fechai = DateAdd("d", 1, in_fechai)
                DoEvents
                FrmCompras.progresbar_kardex.Value = m
                FrmCompras.lblidCompra.Caption = str(in_fechai)
                
           End If
            Next m
    
       
       
        


End Sub






Public Sub update_kardex_bebe(ByVal in_producto As String)
    Dim in_dias As Integer
    
    strCadena = "DELETE FROM kardex WHERE   id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    strCadena = "UPDATE almacen_producto SET stock=0 WHERE ruc='" & KEY_RUC & "' and id_producto='" & in_producto & "'"
    CnBd.Execute (strCadena)
    
    strCadena = "SELECT fecha_emision FROM kardex WHERE id_producto='" & in_producto & "' and  ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC LIMIT 1"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        in_fechai = Format(rstZ("fecha_emision"), "dd-mm-YYYY")
    Else
        in_fechai = "2018-06-20"
    End If
    
            
            
            in_dias = DateDiff("d", in_fechai, KEY_FECHA)
            For m = 0 To in_dias
                If Format(in_fechai, "YYYY-mm-dd") <= KEY_FECHA Then
                
                'COMPRAS
                strCadena = "select * from view_kardex_compra_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    If KEY_RUC = "20128836251" Then
                        Call compras_producto_vargas(in_fechai, in_producto)
                    Else
                        Call compras_producto(in_fechai, in_producto)
                    End If
                    
                End If
               
               
                
                ' transferencias ingreso
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
               ' DoEvents
               
                'ventas salida
       
                
                strCadena = "select * from view_kardex_ventas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call ventas_producto(in_fechai, in_producto)
                End If
                
              
                'DoEvents
                'notas salida
                
                strCadena = "select * from view_kardex_notas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call notas_producto(in_fechai, in_producto)
                End If
                
                End If
                

                in_fechai = DateAdd("d", 1, in_fechai)
                
                
            Next m
    
       
       
        


End Sub

Public Function verificacion_servicio(ByVal in_producto As String) As Boolean

strCadena = "SELECT * FROM view_producto_tipo WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    If rst("servicio") = "si" Then
            strCadena = "DELETE FROM kardex WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
            strCadena = "UPDATE almacen_producto SET stock=0 WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
            strCadena = "UPDATE movimiento_venta_detalle SET precio_costo=0 WHERE id_producto='" & Trim(in_producto) & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            MsgBox "PROCESO EXITOSO..." + Chr(13) + "ESTE ITEM ES UN SERVICIO..", vbInformation
            verificacion_servicio = True
            Exit Function
            
    Else
        verificacion_servicio = False
    End If
        
End If

End Function


Public Sub update_kardex_VARGAS(ByVal in_producto As String)
    Dim in_dias As Integer
    
    If verificacion_servicio(in_producto) = True Then
        Exit Sub
    End If
    
   
    
    strCadena = "DELETE FROM kardex WHERE fecha_emision>='" & Format(get_fecha_periodo_abierto, "YYYY-mm-dd") & "' and    id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    If Format(get_fecha_periodo_abierto, "YYYY-mm-dd") = "2017-12-31" Then
        Call update_saldo_inicial_vargas(in_producto)
        in_fechai = "2018-01-01"
    Else
        in_fechai = Format(get_fecha_periodo_abierto, "YYYY-mm-dd")
    End If
    
        
        in_dias = DateDiff("d", in_fechai, KEY_FECHA)
            For m = 0 To in_dias
                If Format(in_fechai, "YYYY-mm-dd") <= KEY_FECHA Then
                
                'COMPRAS
                
                Call compras_producto_vargas(in_fechai, in_producto)
                
                'transferencias ingreso
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
              
       
                
                strCadena = "select * from view_kardex_ventas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call ventas_producto(in_fechai, in_producto)
                End If
                
              
                strCadena = "select * from view_kardex_notas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and id_tipo_nota NOT IN('04','05') and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call notas_producto(in_fechai, in_producto)
                End If
                
                
                Call salidas_almacen(in_fechai, in_producto)
                End If
                

                in_fechai = DateAdd("d", 1, in_fechai)
                
                
            Next m
    
       
       
       
       
      ' MsgBox "PROCESADO.....", vbInformation, KEY_VENDEDOR
        


End Sub
Public Sub update_kardex_Vargas_modulo_compra(ByVal in_producto As String, ByVal in_fecha_kardex As String)
    Dim in_dias As Integer
    
    If verificacion_servicio(in_producto) = True Then
        Exit Sub
    End If
    
   
    
    strCadena = "DELETE FROM kardex WHERE fecha_emision>='" & Format(in_fecha_kardex, "YYYY-mm-dd") & "' and    id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
  
        in_fechai = Format(in_fecha_kardex, "YYYY-mm-dd")
  
        FrmCompras.progresbar_kardex.Min = 0
        
        in_dias = DateDiff("d", in_fechai, KEY_FECHA)
        If in_dias = 0 Then
           in_dias = 1
        End If
        FrmCompras.progresbar_kardex.Max = in_dias
            For m = 0 To in_dias
                If Format(in_fechai, "YYYY-mm-dd") <= KEY_FECHA Then
                
                'COMPRAS
                
                Call compras_producto_vargas(in_fechai, in_producto)
                
                'transferencias ingreso
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
              
       
                
                strCadena = "select * from view_kardex_ventas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call ventas_producto(in_fechai, in_producto)
                End If
                
              
                strCadena = "select * from view_kardex_notas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and id_tipo_nota NOT IN('04','05') and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call notas_producto(in_fechai, in_producto)
                End If
                
                
                Call salidas_almacen(in_fechai, in_producto)
                End If
                

                in_fechai = DateAdd("d", 1, in_fechai)
                DoEvents
                FrmCompras.progresbar_kardex.Value = m
                DoEvents
            Next m
    
       
       
       
       
      ' MsgBox "PROCESADO.....", vbInformation, KEY_VENDEDOR
        


End Sub


Public Function get_comprobante_sunat(ByVal in_doc As String)

strCadena = "SELECT doc_abrev FROM comprobantes WHERE id_doc='" & in_doc & "'"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
   get_comprobante_sunat = rstZ("doc_abrev")
Else
   get_comprobante_sunat = ""
End If
End Function


Public Sub put_vincular_pagos(ByVal in_origen As String)

strCadena = "SELECT * FROm comprobante_asociado WHERE  id_venta='" & Val(in_origen) & "' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   
   'strCadena = "DELETE FROM mis_cuentas_det_detalle WHERE id_movimiento='" & Val(in_origen) & "'"
   'CnBd.Execute (strCadena)
   
   
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM movimiento_venta where id_venta='" & rst("id_asociado") & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
            strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & rst("id_asociado") & "','" & Val(in_origen) & "','" & rst("monto") & "','" & rst("monto") & "','" & rstZ("id_moneda") & "','" & rstZ("id_moneda") & "','" & rstZ("tc") & "')"
            CnBd.Execute (strCadena)
            If rstZ("id_doc") = "0054" Then
                strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & Val(in_origen) & "','" & rst("id_asociado") & "','" & rst("monto") & "','" & rst("monto") & "','" & rstZ("id_moneda") & "','" & rstZ("id_moneda") & "','" & rstZ("tc") & "')"
                CnBd.Execute (strCadena)
            End If
            
            If rstZ("id_doc") = "0412" Then
                strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & Val(in_origen) & "','" & rst("id_asociado") & "','" & rst("monto") & "','" & rst("monto") & "','" & rstZ("id_moneda") & "','" & rstZ("id_moneda") & "','" & rstZ("tc") & "')"
                CnBd.Execute (strCadena)
            End If
            
            If rstZ("id_doc") = "0008" Then
                strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & Val(in_origen) & "','" & rst("id_asociado") & "','" & rst("monto") & "','" & rst("monto") & "','" & rstZ("id_moneda") & "','" & rstZ("id_moneda") & "','" & rstZ("tc") & "')"
                CnBd.Execute (strCadena)
            End If
            If rstZ("id_doc") = "0001" Or rstZ("id_doc") = "0003" Then
                strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & Val(in_origen) & "','" & rst("id_asociado") & "','" & rst("monto") & "','" & rst("monto") & "','" & rstZ("id_moneda") & "','" & rstZ("id_moneda") & "','" & rstZ("tc") & "')"
                CnBd.Execute (strCadena)
            End If
            
            
        End If
        rst.MoveNext
   Next i
End If
End Sub



Public Function generar_recibo_egreso(ByVal emision As Date, ByVal in_proveedor As String, ByVal in_monto As Double, ByVal in_tc As Single, ByVal in_moneda As String, ByVal in_operacion As String, ByVal in_cuenta_caja As String, ByVal in_compra As String, ByVal in_nota As String) As Double
                    
                    Dim in_numero As String
                    KEY_VENCIMIENTO = KEY_FECHA
                    id_tipo_factura = "0002"
                    igv = "si"
                    dfac = "no"
                    
                    strCadena = "SELECT numero FROM  movimiento_venta WHERE id_doc='0097' and serie='001'  and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
                    Call ConfiguraRstZ(strCadena)
                    If rstZ.RecordCount > 0 Then
                        in_numero = Format(Val(rstZ("numero")) + 1, "000000")
                    Else
                        in_numero = Format(1, "000000")
                    End If
                    horario = Format(Time, "hh:mm")
                    If horario >= "07:00" And horario <= "13:00" Then
                        turno = "M"
                    Else
                        turno = "T"
                    End If
                    
                    Documento = get_comprobante_des("0097") & ":" & "001" & "-" & in_numero
                    strCadena = "P_insert_venta('0097','" & KEY_ALM & "','01','" & in_moneda & "','" & delivery & "'," & _
                    "'001','" & in_numero & "','" & in_proveedor & "','" & get_persona(in_proveedor) & "','0','0','0','" & in_monto & "','0'," & _
                    "'" & in_monto & "','0','" & Format(emision, "YYYY-mm-dd") & "','" & Format(emision, "YYYY-mm-dd") & "','" & id_tipo_factura & "','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & in_tc & "','no','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstP(strCadena)
                    
                    id_venta = rstP(0)
                    
                    in_detalle = get_documento_compra(in_compra)
                                       
                    
                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES " & _
                    "('" & id_venta & "','00','" & in_detalle & "','-','1','" & in_monto & "','0','" & in_monto & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                                    
                    strCadena = "INSERT INTO movimiento_venta_monto(id_venta,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,banco,cheque,cuenta_contable,forma_pago_contable,flujo_caja,id_cuenta_origen,ruc)VALUES " & _
                    "('" & id_venta & "','01','" & get_forma_pago_anterior(in_moneda) & "','" & in_monto & "','" & in_monto * -1 & "','00','-','" & in_operacion & "','-','-','" & get_cuenta_contable_cuenta(in_cuenta_caja) & "','-','-','" & in_cuenta_caja & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(in_numero + 1), "000000") & "' WHERE id_doc='0097' AND serie='001' AND ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    
                    strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & id_venta & "','" & in_compra & "','" & in_monto & "','" & in_monto & "','" & in_moneda & "','" & in_moneda & "','" & in_tc & "')"
                    CnBd.Execute (strCadena)
                                    
                    strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & in_compra & "','" & id_venta & "','" & in_monto & "','" & in_monto & "','" & in_moneda & "','" & in_moneda & "','" & in_tc & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & id_venta & "','" & in_nota & "','" & in_monto & "','" & in_monto & "','" & in_moneda & "','" & in_moneda & "','" & in_tc & "')"
                    CnBd.Execute (strCadena)
            
            
                    
                    generar_recibo_egreso = id_venta
                    
                    
                    
                    
                    
End Function


Public Function generar_recibo_retencion(ByVal emision As Date, ByVal in_proveedor As String, ByVal in_monto As Double, ByVal in_tc As Single, ByVal in_moneda As String, ByVal in_operacion As String, ByVal in_cuenta_caja As String, ByVal in_compra As String, ByVal in_nota As String) As Double
                    
                    Dim in_numero As String
                    KEY_VENCIMIENTO = KEY_FECHA
                    id_tipo_factura = "0002"
                    igv = "si"
                    dfac = "no"
                    
                    strCadena = "SELECT numero FROM  movimiento_venta WHERE id_doc='0097' and serie='001'  and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
                    Call ConfiguraRstZ(strCadena)
                    If rstZ.RecordCount > 0 Then
                        in_numero = Format(Val(rstZ("numero")) + 1, "000000")
                    Else
                        in_numero = Format(1, "000000")
                    End If
                    horario = Format(Time, "hh:mm")
                    If horario >= "07:00" And horario <= "13:00" Then
                        turno = "M"
                    Else
                        turno = "T"
                    End If
                    
                    Documento = get_comprobante_des("0097") & ":" & "001" & "-" & in_numero
                    strCadena = "P_insert_venta('0097','" & KEY_ALM & "','01','" & in_moneda & "','" & delivery & "'," & _
                    "'001','" & in_numero & "','" & in_proveedor & "','" & get_persona(in_proveedor) & "','0','0','0','" & in_monto & "','0'," & _
                    "'" & in_monto & "','0','" & Format(emision, "YYYY-mm-dd") & "','" & Format(emision, "YYYY-mm-dd") & "','" & id_tipo_factura & "','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & in_tc & "','no','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstP(strCadena)
                    
                    id_venta = rstP(0)
                    
                    in_detalle = get_documento_compra(in_compra)
                                       
                    
                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES " & _
                    "('" & id_venta & "','00','" & in_detalle & "','-','1','" & in_monto & "','0','" & in_monto & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                                    
                    strCadena = "INSERT INTO movimiento_venta_monto(id_venta,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,banco,cheque,cuenta_contable,forma_pago_contable,flujo_caja,id_cuenta_origen,ruc)VALUES " & _
                    "('" & id_venta & "','01','" & get_forma_pago_anterior(in_moneda) & "','" & in_monto & "','" & in_monto * -1 & "','00','-','" & in_operacion & "','-','-','" & in_cuenta_caja & "','-','-','" & in_cuenta_caja & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(in_numero + 1), "000000") & "' WHERE id_doc='0097' AND serie='001' AND ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    
                    strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & id_venta & "','" & in_compra & "','" & in_monto & "','" & in_monto & "','" & in_moneda & "','" & in_moneda & "','" & in_tc & "')"
                    CnBd.Execute (strCadena)
                                    
                    strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & in_compra & "','" & id_venta & "','" & in_monto & "','" & in_monto & "','" & in_moneda & "','" & in_moneda & "','" & in_tc & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & id_venta & "','" & in_nota & "','" & in_monto & "','" & in_monto & "','" & in_moneda & "','" & in_moneda & "','" & in_tc & "')"
                    CnBd.Execute (strCadena)
            
            
                    
                    generar_recibo_retencion = id_venta
                    
                    
                    
                    
                    
End Function

Public Function get_documento_compra(ByVal in_compra As String) As String


strCadena = "SELECT comprobante FROM view_cuentas_cobrar WHERE id_compra='" & Val(in_compra) & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    get_documento_compra = rstL("comprobante")
Else
    get_documento_compra = "-"
End If


End Function
Public Sub cancelar_comprobante_nota(ByVal in_venta As String, ByVal in_nota As String, ByVal in_monto As Double)
Dim in_monto_deposito As Single

strCadena = "SELECT (total-function_pago_factura(id_venta,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc)) as pago FROM movimiento_venta WHERE id_venta='" & in_venta & "'"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   If rstLocal(0) > 0 Then  ' tiene saldo pendiente
        If rstLocal(0) > in_monto Then
          in_monto_deposito = in_monto
        Else
          in_monto_deposito = rstLocal(0)
        End If
        
        strCadena = "call p_insert_cancelacion_venta('" & Val(in_nota) & "','" & Val(in_venta) & "','" & in_monto_deposito & "','" & in_monto_deposito & "','" & in_id_mis_cuentas_det & "','01')"
        CnBd.Execute (strCadena)
   
   End If


    
End If
End Sub
Public Sub put_realizar_pago(ByVal in_detalle As String, ByVal in_movimiento As String, ByVal in_monto As Double, ByVal in_doc As String, ByVal in_tc As Single, ByVal in_id_mis_cuentas_det As String, Optional in_tipo As String)
            
            
            
    If in_tipo = "" Then
        in_tipo = "01"
      End If
      
    If in_doc = "0007" Then 'MATAR CON NOTA DE CREDITO
       Call cancelar_comprobante_nota(in_movimiento, in_detalle, in_monto)
    Else
       strCadena = "CALL p_insert_cancelacion_venta('" & Val(in_detalle) & "','" & Val(in_movimiento) & "','" & in_monto & "','" & in_monto & "','" & in_id_mis_cuentas_det & "','" & in_tipo & "')"
       CnBd.Execute (strCadena)
    End If
    
    
          
           
          
           
    
    
    
           
         
End Sub




Public Function get_nota_venta(ByVal in_venta As String, ByVal in_referencia As String) As String

strCadena = "SELECT documento FROM movimiento_venta WHERE id_doc='0007' and  id_comprobante='" & Val(in_referencia) & "' and anulado='no' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   get_nota_venta = rstLocal("documento")
Else
   strCadena = "SELECT documento FROM movimiento_venta WHERE id_venta='" & in_referencia & "' and ruc='" & KEY_RUC & "' LIMIT 1"
   Call ConfiguraRstlocal(strCadena)
   If rstLocal.RecordCount > 0 Then
      get_nota_venta = rstLocal("documento")
   Else
      get_nota_venta = "  -   "
   End If
End If



End Function

Public Function get_diferida(ByVal in_venta) As String

strCadena = "SELECT diferida FROM movimiento_venta WHERE id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
    get_diferida = rstLocal("diferida")
Else
    get_diferida = "no"
End If

End Function



Public Sub actualizar_kardex_recepcion(ByVal in_fecha As Date, ByVal in_producto As String)
    Dim in_costo_igv As Single
    Dim in_afecto_igv As String
    Dim in_moneda As String
    Dim in_factor As Single
    
   'OBTENGO LA ORDEN DE COMPRA
   strCadena = "SELECT * FROM view_orden_compra_kardex WHERE fecha_solicitud='" & Format(in_fecha, "YYYY-mm-dd") & "' and  id_producto='" & Trim(in_producto) & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRstL(strCadena)
   If rstL.RecordCount > 0 Then
               rstL.MoveFirst
               
               For i = 0 To rstL.RecordCount - 1
                    in_recepcion = rstL("id_orden")
                    in_moneda = rstL("id_moneda")
                    in_afecto_igv = rstL("afecto_igv")
                    in_tc = rstL("tc")
                    in_guia_serie = rstL("guia_serie")
                    in_guia_numero = rstL("guia_numero")
                    If in_moneda = "00001" Then
                        in_factor = 1
                    Else
                        in_factor = in_tc
                    End If
                    in_monto_neto = (rstL("precio") * in_factor + rstL("incremento_neto"))
                    If in_afecto_igv = "si" Then
                       
                        in_costo_igv = rstL("precio") * in_factor + rstL("precio") * KEY_IGV * in_factor + rstL("incremento_neto")
                    Else
                        in_costo_igv = in_monto_neto
                    End If
                    
                    If Trim(in_guia_serie) = "" And Trim(in_guia_numero) = "" Then
                        strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & rstL("id_compra") & "' and ruc='" & KEY_RUC & "'"
                        Call ConfiguraRstK(strCadena)
                        If rstK.RecordCount > 0 Then
                            strCadena = "call put_kardex_stock_vitekey('04','" & Format(rstL("fecha_solicitud"), "YYYY-mm-dd") & "','" & Val(in_recepcion) & "','0001','" & rstK("serie") & "','" & rstK("numero") & "','" & Trim(rstL("id_proveedor")) & "','" & rstL("id_producto") & "','" & rstL("cantidad") & "','" & in_costo_igv & "','" & rstL("id_alm") & "','" & rstL("dni_save") & "','" & KEY_RUC & "')"
                        End If
                    Else
                        strCadena = "call put_kardex_stock_vitekey('04','" & Format(rstL("fecha_solicitud"), "YYYY-mm-dd") & "','" & Val(in_recepcion) & "','0009','" & Trim(in_guia_serie) & "','" & Trim(in_guia_numero) & "','" & Trim(rstL("id_proveedor")) & "','" & rstL("id_producto") & "','" & rstL("cantidad") & "','" & in_costo_igv & "','" & rstL("id_alm") & "','" & rstL("dni_save") & "','" & KEY_RUC & "')"
                    End If
                    
                    CnBd.Execute (strCadena)
                    rstL.MoveNext
               Next i
          
        End If
   
End Sub



Public Function get_id_nota(ByVal in_serie As String, ByVal in_numero As String) As String
strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='0007' and serie='" & in_serie & "' and numero='" & in_numero & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
    get_id_nota = rstZ("id_venta")
Else
    get_id_nota = 0
End If
End Function

Public Function get_servicio(ByVal in_producto As String) As String

strCadena = "SELECT * FROM view_producto_servicio WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
    get_servicio = rstLocal("servicio")
Else
    get_servicio = "no"
End If


End Function

Public Function get_fecha_periodo_abierto() As Date
strCadena = "SELECT FechaInicio FROM view_cierre_periodo WHERE ruc='" & KEY_RUC & "' and IndCierreAlmacen='0' and FechaInicio>='2018-01-01' ORDER BY FechaInicio ASC LIMIT 1 "
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
    If Format(rstP("FechaInicio"), "YYYY-mm-dd") = "2018-01-01" Then
        If KEY_PAIS = KEY_PERU Then
            get_fecha_periodo_abierto = "2017-12-31"
        Else
            get_fecha_periodo_abierto = "2018-04-30"
        End If
        
    Else
        get_fecha_periodo_abierto = rstP("FechaInicio")
    End If
Else
    get_fecha_periodo_abierto = "2017-12-31"
End If
End Function

Public Function get_periodo_fecha_ini(ByVal in_periodo As String) As Date
strCadena = "SELECT FechaInicio FROM con_periodo WHERE id='" & in_periodo & "'   ORDER BY FechaInicio ASC LIMIT 1 "
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
    get_periodo_fecha_ini = rstP("FechaInicio")
Else
    get_periodo_fecha_ini = rstP("FechaInicio")
End If
End Function

Public Function get_periodo_cerrado(ByVal in_fecha As Date) As Boolean
strCadena = "SELECT FechaInicio FROM view_cierre_periodo WHERE ruc='" & KEY_RUC & "' and IndCierreAlmacen='1' and FechaInicio>='" & Format(in_fecha, "YYYY-mm-dd") & "' ORDER BY FechaInicio ASC LIMIT 1 "
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
    get_periodo_cerrado = True
Else
    get_periodo_cerrado = False
End If
End Function


Public Function get_fecha_periodo_abierto_compras(ByVal in_fecha As Date) As Date

strCadena = "SELECT FechaInicio FROM view_cierre_periodo WHERE ruc='" & KEY_RUC & "' and IndCierreAlmacen='0' and month(FechaInicio)='" & Month(Format(in_fecha, "YYYY-mm-dd")) & "'  and year(FechaInicio)='" & Year(Format(in_fecha, "YYYY-mm-dd")) & "'   ORDER BY FechaInicio ASC LIMIT 1 "
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
    If Format(rstP("FechaInicio"), "YYYY-mm-dd") = "2018-01-01" Then
        get_fecha_periodo_abierto_compras = "2017-12-31"
    Else
        get_fecha_periodo_abierto_compras = rstP("FechaInicio")
    End If
    
    
Else
    get_fecha_periodo_abierto_compras = "2017-12-31"
End If



End Function


Public Function get_extranjero(ByVal in_dni As String) As String

strCadena = "SELECT * FROM persona WHERE dni='" & in_dni & "'"
Call ConfiguraRstF(strCadena)
If rstF.RecordCount > 0 Then
    get_extranjero = rstF("extranjero")
Else
    get_extranjero = "no"
End If

End Function





Public Function get_persona_existe(ByVal in_dni As String) As Boolean

strCadena = "SELECT * FROM persona WHERE dni='" & in_dni & "' LIMIT 1"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount < 1 Then
    get_persona_existe = False
Else
    get_persona_existe = True
End If

End Function

Public Function get_almacen(ByVal in_alm As String) As String

strCadena = "SELECT * FROM almacen WHERE id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
    get_almacen = rstZ("descripcion")
Else
    get_almacen = "-"
End If


End Function

Public Function control_stock_general(ByVal in_producto As String, ByVal in_cantidad As Double, ByVal in_doc As String) As Boolean

If in_servicio = "si" Then
    control_stock_general = True
Else
strCadena = "SELECT stock FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then

    If rstIN("stock") < in_cantidad And KEY_MOVIMIENTO_SIN_STOCK = "no" Then
                
                MsgBox "Producto no cuenta con STOCK" + Chr(13) + Chr(13) + get_producto(in_producto) + Space(2) + Chr(13) + Chr(13) + "STOCK ACTUAL : " + str(rstIN("stock")) + Chr(13) + "TOTAL PEDIDO :" + str(in_cantidad) + Chr(13) + Chr(13) + "Consulte con el Area de Almacen.", vbInformation, KEY_EMPRESA
                
                If in_doc = "0099" Then
                    control_stock_general = True
                Else
                    
                    
                    control_stock_general = False
                End If
                
    Else
        If KEY_MOVIMIENTO_SIN_STOCK = "si" Then
            control_stock_general = True
        Else
            control_stock_general = True
        End If
        
    End If
End If
End If
End Function


Public Function control_stock_pedido(ByVal in_producto As String, ByVal in_cantidad As Double, ByVal in_comprobante As String) As Boolean

If in_servicio = "si" Then
    control_stock_pedido = True
Else
    
    strCadena = "SELECT stock,habilitado FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstIN(strCadena)
    If rstIN.RecordCount > 0 Then
    
    
    
    If rstIN("stock") < in_cantidad And KEY_MOVIMIENTO_SIN_STOCK = "no" Then
                
              If MsgBox("Producto no cuenta con STOCK" + Chr(13) + Chr(13) + "PEDIDO:" + in_comprobante + Chr(13) + Chr(13) + get_producto(in_producto) + Space(2) + Chr(13) + Chr(13) + "STOCK ACTUAL : " + str(rstIN("stock")) + Chr(13) + "TOTAL PEDIDO :" + str(in_cantidad) + Chr(13) + Chr(13) + "Desea agregar al pedido de todas Formas:", vbInformation + vbYesNo, KEY_EMPRESA) = vbYes Then
                   control_stock_pedido = True
              Else
                  control_stock_pedido = False
              End If
    Else
        
        If rstIN("habilitado") = "no" Then
            MsgBox "Producto no esta Habilitado para la VENTA.", vbInformation
            control_stock_pedido = False
        Else
            control_stock_pedido = True
        End If
        
        
    End If
    
    
End If
End If
End Function

Public Function put_verificar_cobertura(ByVal in_cobertura As String, ByVal in_cliente As String) As Boolean
strCadena = "SELECT cod_unico FROM entidad_empresa WHERE id_tipo_cliente='" & in_cobertura & "' and  cod_unico='" & in_cliente & "' and id_empresa='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
   put_verificar_cobertura = True
Else
   put_verificar_cobertura = False
End If

End Function

Public Function put_verificar_bonificacion(ByVal in_producto As String, ByVal in_cantidad As String, ByVal in_cliente As String, ByVal in_doc As String, ByVal in_serie As String) As Boolean



Dim in_obsequio As String
Dim in_cant As Double
Dim in_precio As Double
Dim in_precios As Single
Dim in_multiplo As Integer





strCadena = "SELECT id_linea,id_sublinea,id_modelo,id_unidad FROM producto WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
        
        
        
     'BONIFICACION GENERAL
     strCadena = "SELECT * FROM view_acumulado_bonificacion WHERE dni_save='" & KEY_USUARIO & "' and id_linea='" & rstL("id_linea") & "' and id_sublinea='" & rstL("id_sublinea") & "' and id_modelo='" & rstL("id_modelo") & "' and ruc='" & KEY_RUC & "'"
     Call ConfiguraRstA(strCadena)
     If rstA.RecordCount > 0 Then
        in_cant = rstA("cantidad")
        
        
        
        strCadena = "SELECT * FROM bonificacion WHERE fecha_fin>='" & KEY_FECHA & "' and  cruzada='no' and por_monto='no' and  anulado='no' and cantidad<='" & Val(in_cant) & "' and  id_linea='" & rstL("id_linea") & "'  and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstA(strCadena)
        If rstA.RecordCount > 0 Then
           
           
           
           
           If put_verificar_cobertura(rstA("id_cobertura"), in_cliente) = False Then
              GoTo nsalir
           End If
        
        'Eliminar Bonificacion del mismo Tipo
        strCadena = "DELETE FROM temporal_ventas WHERE id_producto='" & rstA("id_producto") & "' and obsequio='si' and tipo_bonificacion='01' and dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        'verificacion multiplos
        
        
        in_multiplo = Int(Val(in_cant) / rstA("cantidad"))
        
        in_cantidad = in_multiplo * rstA("cantidad_bonificacion")
        
        in_precios = get_precio_venta_now(rstA("id_producto"))
        
      If control_stock_general(rstA("id_producto"), in_cantidad, in_doc) = True Then
        strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,servicio,obsequio,costo,tipo_bonificacion) VALUES " & _
        "('" & KEY_RUC & "','" & rstL("id_unidad") & "','" & in_cliente & "','" & KEY_ALM & "','" & in_doc & "','" & in_serie & "','-','" & rstA("id_producto") & "','" & Val(in_cantidad) & "'," & _
        "'" & in_precios & " ','0','0','no','" & get_producto(rstA("id_producto")) & "','" & KEY_USUARIO & "','no','si','" & get_precio_costo(rstA("id_producto")) & "','01')"
        CnBd.Execute (strCadena)
        put_verificar_bonificacion = True
       End If
        
        
    Else
nsalir:
        put_verificar_bonificacion = False
    End If
    End If
End If




End Function


Public Function put_verificar_bonificacion_monto(ByVal in_producto As String, ByVal in_cantidad As String, ByVal in_cliente As String, ByVal in_doc As String, ByVal in_serie As String) As Boolean
Dim in_obsequio As String
Dim in_cant As Double
Dim in_precio As Double
Dim in_precios As Single
Dim in_multiplo As Integer
strCadena = "SELECT id_linea,id_unidad FROM producto WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
     strCadena = "SELECT * FROM view_acumulado_bonificacion_monto WHERE dni_save='" & KEY_USUARIO & "' and id_linea='" & rstL("id_linea") & "' and ruc='" & KEY_RUC & "'"
     Call ConfiguraRstA(strCadena)
     If rstA.RecordCount > 0 Then
        in_cant = rstA("total")
        
        
        strCadena = "SELECT * FROM bonificacion WHERE id_cobertura='" & get_tipo_cobertura(in_cliente) & "' and fecha_fin>='" & KEY_FECHA & "' and  cruzada='no' and por_monto='si' and  anulado='no' and cantidad<='" & Val(in_cant) & "' and  id_linea='" & rstL("id_linea") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstA(strCadena)
        If rstA.RecordCount > 0 Then
        
        
        'Eliminar Bonificacion del mismo Tipo
        strCadena = "DELETE FROM temporal_ventas WHERE id_producto='" & rstA("id_producto") & "' and obsequio='si' and tipo_bonificacion='03' and dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        'verificacion multiplos
        
        
        in_multiplo = Int(Val(in_cant) / rstA("cantidad"))
        
        in_cantidad = in_multiplo * rstA("cantidad_bonificacion")
        
        in_precios = get_precio_venta_now(rstA("id_producto"))
        If control_stock_general(rstA("id_producto"), in_cantidad, in_doc) = True Then
        strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,servicio,obsequio,costo,tipo_bonificacion) VALUES " & _
        "('" & KEY_RUC & "','" & rstL("id_unidad") & "','" & in_cliente & "','" & KEY_ALM & "','" & in_doc & "','" & in_serie & "','-','" & rstA("id_producto") & "','" & Val(in_cantidad) & "'," & _
        "'" & in_precios & " ','0','0','no','" & get_producto(rstA("id_producto")) & "','" & KEY_USUARIO & "','no','si','" & get_precio_costo(rstA("id_producto")) & "','03')"
        CnBd.Execute (strCadena)
        End If
        put_verificar_bonificacion_monto = True
        
        
        
    Else
        put_verificar_bonificacion_monto = False
    End If
    End If
End If




End Function

Public Function put_descuento_categoria(ByVal in_producto As String, ByVal in_precio As Single) As Single
Dim in_obsequio As String
Dim in_precios As Single
Dim in_multiplo As Integer
Dim in_descuento As Single
strCadena = "SELECT id_linea FROM producto WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
     
        
        strCadena = "SELECT porcentaje_descuento FROM bonificacion WHERE fecha_fin>='" & KEY_FECHA & "' and  descuento='si'  and  anulado='no' and  id_linea='" & rstL("id_linea") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstA(strCadena)
        If rstA.RecordCount > 0 Then
            put_descuento_categoria = rstA("porcentaje_descuento")
        
        Else
            put_descuento_categoria = 0
        End If
End If


End Function

Public Sub put_verificar_bonificacion_cruzada(ByVal in_producto As String, ByVal in_cantidad As String, ByVal in_cliente As String, ByVal in_doc As String, ByVal in_serie As String)
Dim in_obsequio As String
Dim in_cant As Double
Dim in_precio As Double
Dim in_precios As Single
Dim in_multiplo As Integer
Dim in_multiplo_ant As Integer
Dim in_bonificcion As Boolean
Dim multiplo() As Integer
Dim codigo_bonificacion As Double

    in_bonificcion = False
    



    strCadena = "SELECT DISTINCT id_bonificacion FROM view_bonificacion_cruzada_venta WHERE  dni_save='" & KEY_USUARIO & "' and id_dni='" & in_cliente & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstC(strCadena)
    If rstc.RecordCount > 0 Then
    rstc.MoveFirst
     
     For R = 0 To rstc.RecordCount - 1
        
     
     
        in_bonificcion = False
        
        strCadena = "SELECT * FROM view_bonificacion_cruzada_venta WHERE id_bonificacion='" & rstc("id_bonificacion") & "'and  dni_save='" & KEY_USUARIO & "' and id_dni='" & in_cliente & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstA(strCadena)
        
        strCadena = "SELECT * FROM bonificacion_detalle WHERE   id_bonificacion='" & rstc("id_bonificacion") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        
        If rstA.RecordCount = rstL.RecordCount Then
            in_multiplo = 1
            rstA.MoveFirst
           
            codigo_bonificacion = rstA("id_bonificacion")
            ReDim multiplo(rstA.RecordCount - 1)
            in_multiplo_ant = 0
            For i = 0 To rstA.RecordCount - 1
                 rstL.MoveFirst
                For j = 0 To rstL.RecordCount - 1
                        
                     If rstA("id_producto") = rstL("id_producto") Then
                        
                        in_multiplo = Int(rstA("cantidad_temp") / rstL("cantidad"))
                        multiplo(i) = in_multiplo
                        
                        
                        in_bonificcion = True
                        Exit For
                     Else
                        in_bonificcion = False
                     End If
                     
                     rstL.MoveNext
                Next j
                rstA.MoveNext
            Next i
      
            
        'verificacion multiplos
        
        
        strCadena = "SELECT * FROM bonificacion WHERE anulado='no' and id_bonificacion='" & codigo_bonificacion & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstA(strCadena)
        If rstA.RecordCount > 0 Then
        For i = 0 To UBound(multiplo)
            If i = 0 Then
                in_multiplo = multiplo(i)
            Else
                If in_multiplo <= multiplo(i) Then
                   in_multiplo = in_multiplo
                Else
                   in_multiplo = multiplo(i)
                End If
            End If
        Next i
        in_cantidad = in_multiplo * rstA("cantidad_bonificacion")
        strCadena = "SELECT id_linea,id_sublinea,id_unidad FROM producto WHERE id_producto='" & rstA("id_producto") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
        If control_stock_general(rstA("id_producto"), in_cantidad, in_doc) = True Then
        strCadena = "DELETE FROM temporal_ventas WHERE ruc='" & KEY_RUC & "' and id_alm='" & KEY_ALM & "' and dni_save='" & KEY_USUARIO & "' and obsequio='si' and  id_producto='" & rstA("id_producto") & "' "
        CnBd.Execute (strCadena)
                
        in_precios = get_precio_venta_now(rstA("id_producto"))
        
        strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,servicio,obsequio,costo) VALUES " & _
        "('" & KEY_RUC & "','" & rstL("id_unidad") & "','" & in_cliente & "','" & KEY_ALM & "','" & in_doc & "','" & in_serie & "','-','" & rstA("id_producto") & "','" & Val(in_cantidad) & "'," & _
        "'" & in_precios & " ','0','0','no','" & get_producto(rstA("id_producto")) & "','" & KEY_USUARIO & "','no','si','" & get_precio_costo(rstA("id_producto")) & "')"
        CnBd.Execute (strCadena)
        End If
        End If
        End If
        End If
        
      rstc.MoveNext
     Next R


End If


End Sub



Public Sub put_verificar_bonificacion_cruzada_v2(ByVal in_producto As String, ByVal in_cantidad As String, ByVal in_cliente As String, ByVal in_doc As String, ByVal in_serie As String)
Dim in_obsequio As String
Dim in_cant As Double
Dim in_precio As Double
Dim in_precios As Single
Dim in_multiplo As Integer
Dim in_multiplo_ant As Integer
Dim in_pedido_anterior As Double
Dim in_bonificcion As Boolean
Dim multiplo() As Integer
Dim codigo_bonificacion As Double

    in_bonificcion = False
    strCadena = "SELECT DISTINCT id_bonificacion,all_canal,id_cobertura FROM view_bonificacion_cruzada_venta WHERE  fecha_fin>='" & KEY_FECHA & "' and  id_producto='" & in_producto & "' and  dni_save='" & KEY_USUARIO & "' and id_dni='" & in_cliente & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstC(strCadena)
    If rstc.RecordCount > 0 Then
        rstc.MoveFirst
     
        For R = 0 To rstc.RecordCount - 1
        
            If rstc("all_canal") = "si" Then
        
            Else
                If rstc("id_cobertura") <> get_tipo_cobertura(in_cliente) Then
                    GoTo siguiente_bonificacion
                End If
            End If
        
        
        
     
     
        in_bonificcion = False
        
        strCadena = "SELECT DISTINCT id_producto,sum(cantidad_temp) as cantidad_temp,id_bonificacion FROM view_bonificacion_cruzada_venta WHERE  id_bonificacion='" & rstc("id_bonificacion") & "'and  dni_save='" & KEY_USUARIO & "' and id_dni='" & in_cliente & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' GROUP BY id_bonificacion,id_producto"
        Call ConfiguraRstA(strCadena)
        
        strCadena = "SELECT * FROM bonificacion_detalle WHERE  id_bonificacion='" & rstc("id_bonificacion") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        
        If rstA.RecordCount = rstL.RecordCount And rstA.RecordCount > 0 Then
            in_multiplo = 1
            rstA.MoveFirst
            in_pedido_anterior = rstA("cantidad_temp")
            codigo_bonificacion = rstA("id_bonificacion")
            ReDim multiplo(rstA.RecordCount - 1)
            in_multiplo_ant = 0
            For i = 0 To rstA.RecordCount - 1
                 rstL.MoveFirst
                For j = 0 To rstL.RecordCount - 1
                    If rstA("id_producto") = rstL("id_producto") Then
                        in_multiplo = Int(rstA("cantidad_temp") / rstL("cantidad"))
                        multiplo(i) = in_multiplo
                        in_bonificcion = True
                        Exit For
                     Else
                        in_bonificcion = False
                     End If
                     
                     rstL.MoveNext
                Next j
                rstA.MoveNext
            Next i
      
       End If
       
        'verificacion multiplos
        
        strCadena = "SELECT * FROM bonificacion_cruzada_detalle WHERE  id_bonificacion='" & codigo_bonificacion & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstA(strCadena)
        If rstA.RecordCount > 0 Then
            rstA.MoveFirst
            For m = 0 To rstA.RecordCount - 1
                in_unidad = rstA("id_unidad")
                For i = 0 To UBound(multiplo)
                    If i = 0 Then
                        in_multiplo = multiplo(i)
                    Else
                        If in_multiplo <= multiplo(i) Then
                            in_multiplo = in_multiplo
                        Else
                            in_multiplo = multiplo(i)
                        End If
                    End If
                Next i
                in_cantidad = in_multiplo * rstA("cantidad")
                strCadena = "SELECT id_linea,id_sublinea,id_unidad,agranel FROM producto WHERE id_producto='" & rstA("id_producto") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    If control_stock_general(rstA("id_producto"), in_pedido_anterior + in_cantidad, in_doc) = True Then
                        If rstL("agranel") = "si" Then
                            in_precios = get_precio_unidad(rstA("id_producto"), in_unidad)
                        Else
                            in_precios = get_precio_venta_now(rstA("id_producto"))
                        End If
        
                        strCadena = "CALL get_idTemporalventasv2('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
                        Call ConfiguraRstlocal(strCadena)
                        in_idVenta = rstLocal(0)
        
                        strCadena = "CALL put_bonificacion('2','" & codigo_bonificacion & "','','" & in_cliente & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                        Call ConfiguraRstlocal(strCadena)
                        
        
                        If rstLocal("in_mensaje") = "NORMAL" Then
                            strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,servicio,obsequio,costo,tipo_bonificacion,id_persona_analisis,id_bonificacion,agranel) VALUES " & _
                            "('" & KEY_RUC & "','" & in_unidad & "','" & in_cliente & "','" & KEY_ALM & "','" & in_doc & "','" & in_serie & "','-','" & rstA("id_producto") & "','" & Val(in_cantidad) & "'," & _
                            "'" & in_precios & " ','0','0','no','" & get_producto(rstA("id_producto")) & "','" & KEY_USUARIO & "','no','si','" & get_precio_costo(rstA("id_producto")) & "','02','" & in_idVenta & "','" & codigo_bonificacion & "','" & rstL("agranel") & "')"
                            CnBd.Execute (strCadena)
                        Else
                            If rstLocal("in_mensaje") = 1 Then
                                strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,servicio,obsequio,costo,tipo_bonificacion,id_persona_analisis,id_bonificacion,agranel) VALUES " & _
                                "('" & KEY_RUC & "','" & in_unidad & "','" & in_cliente & "','" & KEY_ALM & "','" & in_doc & "','" & in_serie & "','-','" & rstA("id_producto") & "','" & Val(in_cantidad) / in_multiplo & "'," & _
                                "'" & in_precios & " ','0','0','no','" & get_producto(rstA("id_producto")) & "','" & KEY_USUARIO & "','no','si','" & get_precio_costo(rstA("id_producto")) & "','02','" & in_idVenta & "','" & codigo_bonificacion & "','" & rstL("agranel") & "')"
                                CnBd.Execute (strCadena)
                            End If
                        End If
                    End If
                End If
            rstA.MoveNext
        Next m
     End If
       

siguiente_bonificacion:
      rstc.MoveNext
     Next R


End If


End Sub


Public Function get_credito_disponible_persona(ByVal in_dni As String) As Double

        strCadena = "SELECT func_credito_persona('" & in_dni & "','" & KEY_RUC & "')"
        Call ConfiguraRstT(strCadena)
        in_credito_persona = rstT(0)
        
        If Val(in_credito_persona) > 0 Then
             strCadena = "SELECT funct_total_saldo('" & Trim(in_dni) & "','" & KEY_RUC & "')"
             Call ConfiguraRstT(strCadena)
             in_consumo_persona = rstT(0)
        Else
             in_consumo_persona = 0
        End If

        get_credito_disponible_persona = Val(in_credito_persona) - Val(in_consumo_persona)


End Function

Public Function get_hora_actual() As String

strCadena = "SELECT DATE_SUB(NOW(), INTERVAL 5 HOUR) "
Call ConfiguraRstIN(strCadena)

get_hora_actual = rstIN(0)


End Function
Public Function get_repetido_transferencia(ByVal in_producto As String) As Boolean
strCadena = "SELECT id_producto FROM movimiento_transferencia_temporal WHERE id_producto='" & in_producto & "'  and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   MsgBox "ESTE PRODUCTO YA ESTA EN LA GUIA.... !!!" + Chr(13) + "Modifique la Cantidad del Item Seleccionado.", vbInformation, "Modulo de ITEM'S DUPLICADOS [ACTIVADO]"
   get_repetido_transferencia = True
Else
   get_repetido_transferencia = False
End If
End Function


Public Function get_monto_pago_comprobante(ByVal in_venta As String, ByVal in_fecha As String, ByVal in_moneda As String) As Double
strCadena = "SELECT function_pago_factura('" & Val(in_venta) & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & in_moneda & "','" & KEY_RUC & "')"
Call ConfiguraRstIN(strCadena)
get_monto_pago = rstIN(0)
End Function

Public Function get_deuda_fecha(ByVal in_ruc As String) As Double

strCadena = "SELECT ifnull(sum(saldo),0) FROM cobranza_servicio_persona WHERE id_venta=0 and  dni='" & in_ruc & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If IsNull(0) = True Then
    get_deuda_fecha = 0
Else
    get_deuda_fecha = rstT(0)
End If
End Function


Public Function get_ultimo_dia_periodo(ByVal in_periodo) As Date

    
    strCadena = "SELECT FechaFin FROM con_periodo WHERE id='" & in_periodo & "' LIMIT 1"
    Call ConfiguraRstAux(strCadena)
    If rstAux.RecordCount > 0 Then
       get_ultimo_dia_periodo = rstAux("FechaFin")
    End If
    
    
    
End Function

Public Function get_comprobante_electronico(ByVal in_doc As String, ByVal in_serie As String) As Boolean
strCadena = "SELECT electronico FROM almacen_comprobante WHERE id_doc='" & in_doc & "' and serie='" & Trim(in_serie) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
    If rstAux("electronico") = "si" Then
        get_comprobante_electronico = True
    Else
        get_comprobante_electronico = False
    End If
End If

End Function

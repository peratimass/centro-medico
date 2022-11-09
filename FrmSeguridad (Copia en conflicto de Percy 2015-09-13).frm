VERSION 5.00
Begin VB.Form FrmSeguridad 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMotivo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox TxtClave 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   0
      Picture         =   "FrmSeguridad.frx":0000
      Top             =   0
      Width           =   3555
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "FrmSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fecha As Date
Dim KEY_CARGO_OPE As String
Public Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 2000
ClaveCorrecta = False
If FrmVentas.Procedencia = anular Or FrmVentas.Procedencia = Eliminar Or FrmCompras.Procedencia = revertir Or frmCombo.Procedencia = anular Then
    Me.Width = 3015
    Me.Height = 1815
    Me.Shape1.Height = 1815
    Me.txtMotivo.Text = "Motivo del Proceso"
   
Else
Me.Width = 3015
Me.Height = 990
Me.Shape1.Height = 990
End If
End Sub



Private Sub TxtClave_KeyPress(KeyAscii As Integer)
Dim id_unico As String
Dim Numero As String
Dim serie As String
Dim doc_cod As String
Dim cPersona As String
If KeyAscii = 27 Then
    Unload Me
    Exit Sub
End If
If KeyAscii = 13 Then

If FrmParametrosEmpresa.Procedencia = Nuevo Or FrmParametrosEmpresa.Procedencia = Modificar Then
    strCadena = "SELECT * FROM entidad_empresa WHERE password='" & Trim(Me.TxtClave.Text) & "' AND id_empresa='" & KEY_RUC & "' AND cod_unico='" & KEY_USUARIO & "' and id_cargo='00004' "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
            FrmDetallesParametros.Show
            Unload Me
            
            Exit Sub
        
    End If
End If
       strCadena = "SELECT * FROM entidad_empresa E WHERE password='" & Trim(Me.TxtClave.Text) & "' AND id_empresa='" & KEY_RUC & "' LIMIT 0,1"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount < 1 Then
        MsgBox "PASSWORD DE OPERACIONES INCORRECTA", vbInformation, KEY_EMPRESA
        Unload Me
        Set rst = Nothing
        Exit Sub
       Else
        rst.MoveFirst
        KEY_CARGO_OPE = rst("id_cargo")
        KEY_VENDEDOR_AMBULANTE = rst("cod_unico")
       End If
        
      
    
    
    If FrmTipocambio.Procedencia = Modificar Then
        FrmDetalleTipocaambio.Show
        Unload Me
        FrmTipocambio.Procedencia = Neutro
        Exit Sub
    End If
    
    If FrmVentas.Procedencia = modificar_credito Then
        FrmVentas.frmcredito.Visible = True
        Call Resalta(FrmVentas.txtmontocredito)
        FrmVentas.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    
    If FrmFechaTrabajo.Procendencia = Selecionar Then
        strCadena = "SELECT * FROM entidad_empresa WHERE password='" & Trim(Me.TxtClave.Text) & "' AND id_empresa='" & KEY_RUC & "' AND (id_cargo='00004' OR id_cargo='00009') LIMIT 1" ' Administrador or Cordinador
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            strCadena = "UPDATE almacen SET dni_save='0' WHERE id_alm='" & FrmFechaTrabajo.DtcVentanilla.BoundText & "' and id_sucursal='" & FrmFechaTrabajo.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            strAlm = FrmFechaTrabajo.DtcVentanilla.BoundText
            strCadena = "INSERT INTO almacen_sucesos(id_alm,dni,fecha,hora,ruc)VALUES('" & FrmFechaTrabajo.DtcVentanilla.BoundText & "','" & rst("cod_unico") & "',CURDATE(),CURTIME(),'" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            FrmFechaTrabajo.Procendencia = Neutro
            Call FrmFechaTrabajo.llenar_ventanilla(FrmFechaTrabajo.DtcAlmacen.BoundText)
            FrmFechaTrabajo.DtcVentanilla.BoundText = strAlm
            FrmFechaTrabajo.DTPicker1.Enabled = True
            Unload Me
            Exit Sub
        Else
            MsgBox "PASSWORD DE OPERACIONES INCORRECTA", vbInformation, KEY_EMPRESA
            FrmFechaTrabajo.Procendencia = Neutro
            Unload Me
            Exit Sub
        End If
    End If
    
    
    If FrmDetalleLinea.Procedencia = Eliminar Then
       strCadena = "DELETE FROM linea_mantenimiento WHERE id_mantenimiento='" & Val(FrmDetalleLinea.hfmantenimientos.TextMatrix(FrmDetalleLinea.hfmantenimientos.Row, 0)) & "'"
       CnBd.Execute (strCadena)
       Call FrmDetalleLinea.llenar_mantenimientos(FrmLinea.HfgLinea.TextMatrix(FrmLinea.HfgLinea.Row, 0), FrmDetalleLinea.hfmantenimientos)
       FrmDetalleLinea.Procedencia = Neutro
       Unload Me
       Exit Sub
    End If
    

If FrmDetalleLinea.Procedencia = eliminar_insumo Then
       strCadena = "DELETE FROM linea_mantenimiento_detalle WHERE id_detalle='" & Val(FrmDetalleLinea.HfInsumos.TextMatrix(FrmDetalleLinea.HfInsumos.Row, 0)) & "'"
       CnBd.Execute (strCadena)
       Call FrmDetalleLinea.llenar_insumos(FrmDetalleLinea.hfmantenimientos.TextMatrix(FrmDetalleLinea.hfmantenimientos.Row, 0), FrmDetalleLinea.HfInsumos)
       FrmDetalleLinea.Procedencia = Neutro
       Unload Me
       Exit Sub
    End If

    
    If FrmVentas.Procedencia = seleccionar_vendedor Then
       FrmVentas.DtcVendedor.BoundText = KEY_VENDEDOR_AMBULANTE
       If Val(FrmVentas.TxtIdVenta.Text) < 1 Then
            Call FrmVentas.get_auto_pago(FrmVentas.DtcTipoDoc.BoundText)
            Call FrmVentas.Save
            Call impresion_pedido(Val(FrmVentas.TxtIdVenta.Text))
        Else
            Call AnularVentas(Trim(FrmVentas.DtcTipoDoc.BoundText), Trim(FrmVentas.TxtSerie.Text), Trim(FrmVentas.TxtNumeroDoc.Text), Trim(FrmVentas.DtcAlmacen.BoundText))
            Call EliminarVentas(Trim(FrmVentas.DtcTipoDoc.BoundText), Trim(FrmVentas.TxtSerie.Text), Trim(FrmVentas.TxtNumeroDoc.Text), Trim(FrmVentas.DtcAlmacen.BoundText))
            Call FrmVentas.get_auto_pago(FrmVentas.DtcTipoDoc.BoundText)
            Call FrmVentas.Save
            Call impresion_pedido(Val(FrmVentas.TxtIdVenta.Text))
        End If
        
        Call FrmVentas.Nuevo
        FrmVentas.Procedencia = Neutro
        Unload Me
       Exit Sub
    End If
    
    If FrmVentas.Procedencia = Modificar Then
       strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & FrmVentas.TxtIdVenta.Text & "' and ruc='" & KEY_RUC & "'"
       Call ConfiguraRstZ(strCadena)
       If rstZ.RecordCount > 0 Then
          rstZ.MoveFirst
          strCadena = "DELETE FROM temporal_ventas WHERE id_doc='" & FrmVentas.DtcTipoDoc.BoundText & "' and numero='" & FrmVentas.TxtNumeroDoc.Text & "' and id_serie='" & FrmVentas.TxtSerie.Text & "' and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
          CnBd.Execute (strCadena)
          
          For i = 0 To rstZ.RecordCount - 1
            strCadena = "INSERT INTO temporal_ventas(ruc,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,dni_save) VALUES " & _
            "('" & KEY_RUC & "','" & KEY_ALM & "','" & FrmVentas.DtcTipoDoc.BoundText & "','" & Trim(FrmVentas.TxtSerie.Text) & "','" & FrmVentas.TxtNumeroDoc.Text & "','" & rstZ("id_producto") & "','" & rstZ("cantidad") & "'," & _
            "'" & rstZ("precio") & " ','" & rstZ("total") & "','0','si','" & KEY_USUARIO & "')"
            CnBd.Execute (strCadena)
            rstZ.MoveNext
          Next i
          Call FrmVentas.llenarGrid_det(FrmVentas.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.TxtSerie.Text, FrmVentas.DtcTipoDoc.BoundText)
          FrmVentas.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
       End If
       FrmVentas.Procedencia = Neutro
       Unload Me
       Exit Sub
    End If
    
    
    If FrmFechaTrabajo.Procendencia = buscar Then
         
         If KEY_CARGO = KEY_ADMINISTRADOR Or KEY_CARGO = KEY_SUPERVISOR Or KEY_CARGO = KEY_AUTOR Then
            FrmFechaTrabajo.DTPicker1.Enabled = True
         Else
            MsgBox "PASSWORD INCORRECTO", vbQuestion, KEY_EMPRESA
         End If
         Unload Me
         FrmFechaTrabajo.Procendencia = Neutro
         Exit Sub
    End If
    
    
    
    
    If FrmFechaTrabajo.Procendencia = Eliminar Then
         
         If KEY_CARGO = KEY_ADMINISTRADOR Or KEY_CARGO = KEY_SUPERVISOR Or KEY_CARGO = KEY_AUTOR Then
            strCadena = "DELETE FROM impresora where id_impresora='" & Val(FrmFechaTrabajo.HfPrinter.TextMatrix(FrmFechaTrabajo.HfPrinter.Row, 0)) & "' and id_alm='" & KEY_ALM & "' ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            Call FrmFechaTrabajo.llenar_printer(FrmFechaTrabajo.HfPrinter)
         Else
            MsgBox "USTED NO CUENTA CON EL PERMISO PARA REALIZAR ESTE PROCESO", vbQuestion, KEY_EMPRESA
         End If
         Unload Me
         FrmFechaTrabajo.Procendencia = Neutro
         Exit Sub
    End If
    
    
    
    If FrmDetallesParametros.Procedencia = buscar Then
         strCadena = "SELECT * FROM Seguridad WHERE Clave='" & Trim(Me.TxtClave.Text) & "' AND id_cargo='" & KEY_AUTOR & "' "
         Call ConfiguraRst(strCadena)
         If rst.RecordCount > 0 Then
            
            strCadena = "DELETE FROM Producto WHERE Ruc='" & Trim(FrmParametrosEmpresa.HfgMarcas.TextMatrix(FrmParametrosEmpresa.HfgMarcas.Row, 0)) & "'"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Persona "
            CnBd.Execute (strCadena)
             strCadena = "INSERT INTO Persona(Telefono1,Telefono2,NombrePersona,sDireccionCliente1,sDireccionCliente2,Per_Ruc,Per_fax,Per_Percepcion,Per_Retencion," & _
             "sEmailCliente,Cliente,proveedor,contable,transportista,personal,Observacion, MontoAdelantado,Monto,puntos,registrado_por,descuento_por) VALUES " & _
                "('00000000','00000000','CLIENTE','CHICLAYO','CHICLAYO','00000000000','00000000','F','F','percy19_is@hotmail.com','V','V','V','V','V','CLIENTE GENERAL', " & _
                "'0','0','0','" & Trim(KEY_USUARIO) & "','0')"
            CnBd.Execute (strCadena)
           ' strCadena = "DELETE FROM Marcas "
            'CnBd.Execute (strCadena)
            'strCadena = "DELETE FROM Linea"
            'CnBd.Execute (strCadena)
             strCadena = "DELETE FROM Detalle_DocumentoCompra"
            CnBd.Execute (strCadena)
             strCadena = "DELETE FROM Detalle_Documentoventa"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM movimiento_venta_cuotas"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Almacen_Productos"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Chequera"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Cheques"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Combo"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Combo_detalle"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM CuentasCorrientes"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Derivado_Detalle"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Detalle_Documento_pedido"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Detalle_DocumentoCompra"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Detalle_DocumentoTransferencia"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM DetalleAdelantos"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM DetalleGuia"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM DetallePagoCreditos"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM DetallePagos"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM DetallePedido"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Documento_Pedido"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM DocumentoCompra"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM DocumentoTransferencia"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM DocumentoVenta"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM DocumentoVenta_montos"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM DocumentoVenta_Targeta"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Kardex"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM mis_cuentas"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM mis_cuentas_det"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM movimiento_caja"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM orden_pago"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM OrdenPago"
            CnBd.Execute (strCadena)
           
            strCadena = "DELETE FROM Producto_barras"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Producto_mermas"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Producto_mermas_detalle"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Producto_precio"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Producto_Proveedor"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Producto_Recomendado"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Producto_sub"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Temporal_Compras"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Temporal_Pedido"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Temporal_Transferencias"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM Temporal_Ventas"
            CnBd.Execute (strCadena)
            FrmDetallesParametros.Procedencia = Neutro
            MsgBox "Proceso Completado con Exito !!!"
            Unload Me
         End If
         Exit Sub
    End If
    If FrmDetallesParametros.Procedencia = Eliminar Then
                  
            strCadena = "DELETE FROM movimiento_venta WHERE ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM movimiento_venta_detalle WHERE ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM movimiento_compra WHERE ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM movimiento_compra_detalle WHERE ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM movimiento_venta_cuotas"
            CnBd.Execute (strCadena)
            strCadena = "DELETE FROM movimiento_transferencia WHERE ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
            strCadena = "DELETE FROM movimiento_transferencia_detalle WHERE ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
            strCadena = "DELETE FROM kardex WHERE ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            strCadena = "UPDATE almacen_comprobante SET numero='000001' WHERE ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            strCadena = "UPDATE almacen_producto SET stock_contable=0 WHERE ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
           ' strCadena = "DELETE FROM movimiento_venta_targeta WHERE ruc='" & KEY_RUC & "'"
            'CnBd.Execute (strCadena)
            'strCadena = "DELETE FROM movimiento_caja WHERE ruc='" & KEY_RUC & "'"
            'CnBd.Execute (strCadena)
            strCadena = "DELETE FROM mis_cuentas_det WHERE ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            FrmDetallesParametros.Procedencia = Neutro
            MsgBox "Proceso Completado con Exito !!!", vbInformation, KEY_EMPRESA
            Unload Me
           Exit Sub
    End If
    
    
    If FrmVentas.Procedencia = imprimir_s Then
        Dim impresiones As Integer
         'If KEY_CARGO = KEY_ADMIN Or KEY_CARGO = KEY_SUPER Or KEY_CARGO = KEY_AUTOR or key_cargo Then
            strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & FrmVentas.DtcTipoDoc.BoundText & "' AND serie='" & FrmVentas.TxtSerie.Text & "' AND numero='" & FrmVentas.TxtNumeroDoc.Text & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRst(strCadena)
            impresiones = rst("impresiones") + 1
            strCadena = "UPDATE movimiento_venta SET impresiones='" & impresiones & "' WHERE id_venta='" & rst("id_venta") & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            Call Orden_Impresion(FrmVentas.DtcTipoDoc.BoundText, FrmVentas.TxtSerie.Text, FrmVentas.TxtNumeroDoc.Text, rst("id_tipo_factura"))
            Unload Me
            FrmVentas.Procedencia = Neutro
            Exit Sub
    End If
    
   If FrmVentasPersonalizada.Procedencia = imprimir_s Then
        
         If KEY_CARGO = KEY_ADMIN Or KEY_CARGO = KEY_SUPER Or KEY_CARGO = KEY_AUTOR Then
            strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & FrmVentasPersonalizada.DtcTipoDoc.BoundText & "' AND serie='" & FrmVentasPersonalizada.TxtSerie.Text & "' AND numero='" & FrmVentasPersonalizada.TxtNumeroDoc.Text & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRst(strCadena)
            impresiones = rst("impresiones") + 1
            strCadena = "UPDATE movimiento_venta SET impresiones='" & impresiones & "' WHERE id_venta='" & rst("id_venta") & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
            Call Orden_Impresion(FrmVentas.DtcTipoDoc.BoundText, FrmVentasPersonalizada.TxtSerie.Text, FrmVentasPersonalizada.TxtNumeroDoc.Text, rst("id_tipo_factura"))
            
         Else
            MsgBox "PASSWORD INCORRECTO", vbInformation, KEY_EMPRESA
            
         End If
         Unload Me
         FrmVentas.Procedencia = Neutro
         Exit Sub
    End If
    
    If frmCorProcesos.Procedencia = Modificar Then
       Select Case frmCorProcesos.Txtid_estado
            Case "01"
                
                strCadena = "DELETE FROM imp_producto_movimiento WHERE id_detalle='" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "' and id_proceso='01' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                 strCadena = "UPDATE imp_producto_detalle set id_estado='01' WHERE id_detalle='" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "'"
                CnBd.Execute (strCadena)
                
                'strCadena = "UPDATE imp_producto_movimiento set estado='0'  WHERE id_detalle='" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "' and id_proceso='01' and ruc='" & KEY_RUC & "'"
                'CnBd.Execute (strCadena)Call seleccion
                'frmCorProcesos.txtreiniciar.Text = "si"
                'Call frmCorProcesos.evaluarBotonesReiniciar("01", "0")
                Call frmCorProcesos.seleccion
                frmCorProcesos.cmdreiniciarSoldadura.Visible = False
                
            Case "02"
                strCadena = "DELETE FROM imp_producto_movimiento WHERE id_detalle='" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "' and id_proceso='02' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                strCadena = "UPDATE imp_producto_detalle set id_estado='01' WHERE id_detalle='" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "'"
                CnBd.Execute (strCadena)
                Call frmCorProcesos.actualizaGrid(frmCorProcesos.gridDetalle, "01", frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0))
                'strCadena = "UPDATE imp_producto_movimiento set estado='0'  WHERE id_detalle='" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "' and id_proceso='02' and ruc='" & KEY_RUC & "'"
                ''CnBd.Execute (strCadena)
                'frmCorProcesos.txtreiniciar.Text = "si"
                'Call frmCorProcesos.evaluarBotonesReiniciar("02", "0")
                Call frmCorProcesos.actualizaGrid(frmCorProcesos.gridDetalle, "01", frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0))
                Call frmCorProcesos.seleccion
                frmCorProcesos.cmdreiniciarSoldadura.Visible = False
            Case "03"
                strCadena = "DELETE FROM imp_producto_movimiento WHERE id_detalle='" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "' and id_proceso='03' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                 strCadena = "UPDATE imp_producto_detalle set id_estado='01' WHERE id_detalle='" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "'"
                CnBd.Execute (strCadena)
                Call frmCorProcesos.actualizaGrid(frmCorProcesos.gridDetalle, "01", frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0))
                Call frmCorProcesos.seleccion
                frmCorProcesos.cmdreiniciartapiz.Visible = False
       End Select
       frmCorProcesos.Procedencia = Neutro
       Unload Me
       Exit Sub
    End If
    
    
    
    If FrmTransferencias.Procedencia = Modificar Then
        strCadena = "SELECT * FROM movimiento_transferencia WHERE id_doc='" & FrmTransferencias.DtcTipoDoc.BoundText & "' AND serie='" & Trim(FrmTransferencias.TxtSerie.Text) & "' AND numero='" & FrmTransferencias.TxtNumeroDoc.Text & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
               FrmTransferencias.HfChasis.Enabled = False
               FrmTransferencias.TxtId_transferencia.Text = rst("id_transferencia")
               FrmTransferencias.txtverificado.Text = "si"
               
               strCadena = "SELECT * FROM movimiento_transferencia_detalle WHERE id_transferencia='" & rst("id_transferencia") & "' AND ruc='" & KEY_RUC & "'"
               Call ConfiguraRst(strCadena)
               If rst.RecordCount > 0 Then
                  rst.MoveFirst
                  strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
                  CnBd.Execute (strCadena)
                  For i = 0 To rst.RecordCount - 1
                    strCadena = "INSERT INTO movimiento_transferencia_temporal(id_doc,serie,numero,id_producto,cantidad,recibido,peso,total,dni_save,ruc)VALUES " & _
                    "('" & FrmTransferencias.DtcTipoDoc.BoundText & "','" & FrmTransferencias.TxtSerie.Text & "','" & FrmTransferencias.TxtNumeroDoc.Text & "','" & rst("id_producto") & "'," & _
                    "'" & rst("cantidad") & "','" & rst("cantidad") & "','" & rst("peso") & "','" & rst("total") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    rst.MoveNext
                  Next i
               End If
               
               Call FrmTransferencias.Llenar_Temporal(FrmTransferencias.HfDetalle)
               FrmTransferencias.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
               FrmTransferencias.TlbGrabar.Buttons("(Verificar)").Enabled = False
               FrmTransferencias.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
               'FrmTransferencias.HfChasis.Enabled = True
               FrmTransferencias.HfSeries.Enabled = True
            End If
        FrmTransferencias.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    If FrmTransferencias.Procedencia = anular Then
        
        strCadena = "SELECT id_transferencia,id_alm_origen as id_alm,id_alm_destino FROM movimiento_transferencia WHERE id_doc='" & FrmTransferencias.DtcTipoDoc.BoundText & "' AND serie='" & Trim(FrmTransferencias.TxtSerie.Text) & "' AND numero='" & FrmTransferencias.TxtNumeroDoc.Text & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
            strCadena = "UPDATE movimiento_transferencia SET anulado='si' WHERE id_transferencia='" & rstZ("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            strCadena = "SELECT chasis,id_producto FROM movimiento_transferencia_series WHERE id_transferencia='" & rstZ("id_transferencia") & "'"
            Call ConfiguraRstL(strCadena)
            If rstL.RecordCount > 0 Then
                For p = 0 To rstL.RecordCount - 1
                    strCadena = "DELETE FROM imp_producto_detalle WHERE nro_chasis='" & rstL("chasis") & "' and id_alm='" & rstZ("id_alm_destino") & "' and ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "SELECT sum(cantidad_real),sum(cantidad_contable) FROM kardex where id_alm='" & rstZ("id_alm") & "' and  id_producto='" & rstL("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstK(strCadena)
                    
                    strCadena = "UPDATE almacen_producto set stock = '" & rstK(0) & "'  , `stock_contable` = '" & rstK(1) & "'  where ruc = '" & KEY_RUC & "' and id_alm ='" & rstZ("id_alm") & "' and id_producto = '" & rstL("id_producto") & "'"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "SELECT sum(cantidad_real),sum(cantidad_contable) FROM kardex where id_alm='" & rstZ("id_alm_destino") & "' and  id_producto='" & rstL("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstK(strCadena)
                    
                    strCadena = "update almacen_producto set stock = '" & rstK(0) & "'  , `stock_contable` = '" & rstK(1) & "'  where ruc = '" & KEY_RUC & "' and id_alm ='" & rstZ("id_alm_destino") & "' and id_producto = '" & rstL("id_producto") & "'"
                    CnBd.Execute (strCadena)
    
                    rstL.MoveNext
                Next p
            End If
            
            
        End If
        
        
        
        FrmTransferencias.lblAnulado.Visible = True
        FrmTransferencias.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    
    If FrmTransferencias.Procedencia = Eliminar Then
        
        strCadena = "SELECT id_transferencia,id_alm_origen as id_alm,id_alm_destino FROM movimiento_transferencia WHERE id_doc='" & FrmTransferencias.DtcTipoDoc.BoundText & "' AND serie='" & Trim(FrmTransferencias.TxtSerie.Text) & "' AND numero='" & FrmTransferencias.TxtNumeroDoc.Text & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
           
            strCadena = "SELECT chasis,id_producto FROM movimiento_transferencia_series WHERE id_transferencia='" & rstZ("id_transferencia") & "'"
            Call ConfiguraRstL(strCadena)
            If rstL.RecordCount > 0 Then
                For p = 0 To rstL.RecordCount - 1
                    strCadena = "DELETE FROM imp_producto_detalle WHERE nro_chasis='" & rstL("chasis") & "' and id_alm='" & rstZ("id_alm_destino") & "' and ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "SELECT sum(cantidad_real),sum(cantidad_contable) FROM kardex where id_alm='" & rstZ("id_alm") & "' and  id_producto='" & rstL("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstK(strCadena)
                    
                    strCadena = "UPDATE almacen_producto set stock = '" & rstK(0) & "'  , `stock_contable` = '" & rstK(1) & "'  where ruc = '" & KEY_RUC & "' and id_alm ='" & rstZ("id_alm") & "' and id_producto = '" & rstL("id_producto") & "'"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "SELECT sum(cantidad_real),sum(cantidad_contable) FROM kardex where id_alm='" & rstZ("id_alm_destino") & "' and  id_producto='" & rstL("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstK(strCadena)
                    
                    strCadena = "update almacen_producto set stock = '" & rstK(0) & "'  , `stock_contable` = '" & rstK(1) & "'  where ruc = '" & KEY_RUC & "' and id_alm ='" & rstZ("id_alm_destino") & "' and id_producto = '" & rstL("id_producto") & "'"
                    CnBd.Execute (strCadena)
    
                    rstL.MoveNext
                Next p
                strCadena = "DELETE FROM movimiento_transferencia WHERE id_transferencia='" & rstZ("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
            End If
            
            
        End If
        FrmTransferencias.Procedencia = Neutro
        Call FrmTransferencias.Nuevo
        Unload Me
        Exit Sub
    End If
    
    
    If FrmParteDiaria.Procedencia = anular Then
        strCadena = "UPDATE parte_maquinaria SET anulado='si' WHERE id_parte='" & Val(FrmParteDiaria.TxtId_parte.Text) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        FrmParteDiaria.lblAnulado.Visible = True
        FrmParteDiaria.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    If FrmParteMaterial.Procedencia = anular Then
        strCadena = "UPDATE parte_material SET anulado='si' WHERE id_material='" & Val(FrmParteMaterial.TxtId_parte.Text) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        FrmParteMaterial.lblAnulado.Visible = True
        FrmParteMaterial.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    
    If FrmParteMaterial.Procedencia = Modificar Then
       FrmParteMaterial.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
       FrmParteMaterial.Procedencia = Neutro
       Unload Me
       Exit Sub
    End If
    
       If frmCombo.Procedencia = anular Then
          Call anular_combo(frmCombo.TxtCombo.Text)
          frmCombo.Procedencia = Neutro
          Unload Me
          Exit Sub
       End If
       
       If FrmDerivados.Procedencia = anular Then
          Call anular_derivado(FrmDerivados.TxtCombo.Text)
          FrmDerivados.Procedencia = Neutro
          Unload Me
          Exit Sub
       End If
        
        
        If FrmTipocambio.Procedencia = Modificar Then
            FrmTipocambio.HfgCambio.col = 0
            strCadena = "SELECT * FROM tipo_cambio WHERE fecha='" & CVDate(FrmTipocambio.HfgCambio.Text) & "'"
            Call ConfiguraRst(strCadena)
            FrmDetalleTipocaambio.Show
            fecha = CVDate(FrmTipocambio.HfgCambio.Text)
            FrmDetalleTipocaambio.txtCambio.Text = rst("valor")
            FrmTipocambio.Procedencia = Neutro
            Unload Me
            Exit Sub
        End If
        
        If FrmCompras.Procedencia = revertir Then
            FrmCompras.rever = True
            FrmCompras.TlbGrabar.Buttons(KEY_REVERTIR).Enabled = False
            FrmCompras.DtcTipoDoc.Enabled = True
            
            FrmCompras.txtAñoFabricacion.Locked = False
            FrmCompras.txtnumero_dua.Locked = False
            FrmCompras.TxtAnioDua.Locked = False
            FrmCompras.txtañomodelo.Locked = False
            
            
            FrmCompras.TxtSerie.Enabled = True
            FrmCompras.TxtNumeroDoc.Enabled = True
            FrmCompras.CmdActualizar.Visible = True
            FrmCompras.txtcantidad.Enabled = True
            FrmCompras.TxtDescripcionProducto.Enabled = True
            FrmCompras.cmdRevertir.Visible = False
            FrmCompras.CmdAgregar.Enabled = True
            FrmCompras.CmdQuitar.Enabled = True
            FrmCompras.cmdGastos.Visible = True
            FrmCompras.cmdRecalcular.Visible = True
            c_serie = FrmCompras.TxtSerie.Text
            c_numero = FrmCompras.TxtNumeroDoc.Text
            c_doc_cod = FrmCompras.DtcTipoDoc.BoundText
            cPersona = FrmCompras.txtruc.Text
            
            
            
            Dim idCompra As Double
            strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Trim(c_doc_cod) & "' AND serie='" & Trim(c_serie) & "' AND numero='" & Trim(c_numero) & "' AND id_proveedor='" & Trim(cPersona) & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                idCompra = rst("id_compra")
            Else
                strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Trim(c_doc_cod) & "' AND serie='" & Trim(c_serie) & "' AND numero='" & formato_item(Trim(c_numero), 6) & "' AND id_proveedor='" & Trim(cPersona) & "' AND ruc='" & KEY_RUC & "'"
                Call ConfiguraRst(strCadena)
                If rst.RecordCount > 0 Then
                    idCompra = rst("id_compra")
                Else
                    MsgBox "OCURRIO UN ERROR INESPERADO", vbInformation, KEY_EMPRESA
                    Exit Sub
                End If
                
            End If
            
            strCadena = "DELETE FROM movimiento_compra_temporal  WHERE id_doc='" & Trim(c_doc_cod) & "' AND serie='" & Trim(c_serie) & "' AND numero='" & Trim(c_numero) & "' and dni_save='" & Trim(KEY_USUARIO) & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & idCompra & "' AND ruc='" & KEY_RUC & "' "
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                
                For i = 0 To rst.RecordCount - 1
                strCadena = "INSERT INTO movimiento_compra_temporal(id_doc,serie,numero,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento," & _
                "valor_neto,isc,igv,ivap,otros,percepcion,exonerado,valor_venta,precio_venta,p_venta,p_costo,dni_save,id_alm,ruc) VALUES " & _
                "('" & Trim(c_doc_cod) & "','" & Trim(c_serie) & "','" & Trim(c_numero) & "','" & rst("id_producto") & "','" & rst("cantidad") & "','" & rst("c_unitario") & "'," & _
                "'" & rst("dsto_soles") & "', '" & rst("dsto_procentaje") & "','" & rst("total_descuento") & "','" & rst("valor_neto") & "','" & rst("isc") & "','" & rst("igv") & "'," & _
                "'" & rst("ivap") & "','" & rst("otros") & "','" & rst("percepcion") & "','" & rst("exonerado") & "','" & rst("valor_venta") & "','" & rst("total") & "','" & rst("p_venta") & "'," & _
                "'" & rst("p_costo") & "','" & KEY_USUARIO & "','" & rst("id_alm") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                
                
                rst.MoveNext
            Next i
            End If
            
            strCadena = "DELETE FROM imp_producto_detalle_temp where id_compra='" & idCompra & "' "
            CnBd.Execute (strCadena)
            strCadena = "SELECT * FROM imp_producto_detalle WHERE id_compra='" & idCompra & "'"
                Call ConfiguraRstZ(strCadena)
                If rstZ.RecordCount > 0 Then
                    rstZ.MoveFirst
                    For m = 0 To rstZ.RecordCount - 1
                        strCadena = "INSERT INTO imp_producto_detalle_temp(id_compra,id_producto,serie,id_estado,id_estado_detalle,id_alm,anio_fabricacion,anio_contenedor,nro_contenedor,nro_chasis,nro_motor,anio_modelo,item,serie_asignada,id_orden,vendido)VALUES  " & _
                        "('" & rstZ("id_compra") & "','" & rstZ("id_producto") & "','" & rstZ("serie") & "','" & rstZ("id_estado") & "','" & rstZ("id_estado_detalle") & "','" & rstZ("id_alm") & "','" & rstZ("anio_fabricacion") & "','" & rstZ("anio_contenedor") & "','" & rstZ("nro_contenedor") & "','" & rstZ("nro_chasis") & "','" & rstZ("nro_motor") & "','" & rstZ("anio_modelo") & "','" & rstZ("item") & "','" & rstZ("serie_asignada") & "','" & rstZ("id_orden") & "','" & rstZ("vendido") & "') "
                        CnBd.Execute (strCadena)
                        rstZ.MoveNext
                    Next m
                End If
            
                        
            FrmCompras.Procedencia = Neutro
            FrmCompras.TxtCodProducto.Locked = False
            FrmCompras.TxtCodProducto.Enabled = True
           
            Call FrmCompras.llenarGrid_det(FrmCompras.HfdDetalle, c_numero, c_doc_cod, c_serie)
            FrmCompras.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
            FrmCompras.DtcDocumento.Enabled = True
            FrmCompras.TxtSerieG.Enabled = True
            FrmCompras.TxtNumeroG.Enabled = True
            FrmCompras.DtcDocumento.SetFocus
            Unload Me
            Exit Sub
        End If
       
         
         
         
         If FrmDelivery.Procedencia = Modificar Then
            
            strCadena = "UPDATE movimiento_venta SET id_delivery='no' WHERE id_venta='" & Trim(FrmDelivery.HfgLinea.TextMatrix(FrmDelivery.HfgLinea.Row, 0)) & "' AND ruc='" & KEY_RUC & "' "
            CnBd.Execute (strCadena)
            
            Call FrmDelivery.LlenarDelivery(FrmDelivery.HfgLinea)
            FrmDelivery.Procedencia = Neutro
            Unload Me
            Exit Sub
        End If
        If FrmPedido.Procedencia = anular Then
            If KEY_CARGO = "00004" Or KEY_USUARIO = FrmPedido.TxtUsuario.Text Then
            strCadena = "UPDATE movimiento_pedido SET anulado='si' WHERE numero='" & Trim(FrmPedido.TxtNumeroDoc.Text) & "' AND serie='" & Trim(FrmPedido.TxtSerie.Text) & "' AND id_doc='" & Trim(FrmPedido.DtcTipoDoc.BoundText) & "' AND ruc='" & KEY_RUC & "' "
            CnBd.Execute (strCadena)
            FrmPedido.Procedencia = Neutro
            FrmPedido.lblAnulado.Visible = True
        Else
            MsgBox "NO TIENE LOS PERMISOS PARA ANULAR", vbInformation, KEY_EMPRESA
        End If
            Unload Me
            Exit Sub
        End If
        If FrmPedido.Procedencia = Eliminar Then
            strCadena = "DELETE from Documento_Pedido  WHERE numero='" & Trim(FrmPedido.TxtNumeroDoc.Text) & "' AND serie='" & Trim(FrmPedido.TxtSerie.Text) & "' AND doc_cod='" & Trim(FrmPedido.DtcTipoDoc.BoundText) & "' "
            Call EjecutaRST(strCadena)
            Set RstEjecuta = Nothing
            FrmPedido.Nuevo
            Unload Me
            Exit Sub
        End If
        If FrmPendientes.Procedencia = Modificar Then
            fecha = Date
            FrmPendientes.HfGuardado.col = 0
            id_unico = FrmPendientes.HfGuardado.Text
            FrmPendientes.HfGuardado.col = 1
            Numero = FrmPendientes.HfGuardado.Text
            FrmPendientes.HfGuardado.col = 2
            serie = FrmPendientes.HfGuardado.Text
            FrmPendientes.HfGuardado.col = 3
            doc_cod = FrmPendientes.HfGuardado.Text
            
            
            strCadena = "UPDATE DocumentoVenta SET dEmisionVenta='" & CVDate(fecha) & "',envase='si', id_usuario='" & Trim(KEY_USUARIO) & "' WHERE cDocumentoVenta='" & Trim(Numero) & "' AND id_documentoventa='" & Trim(id_unico) & "' AND sSerie='" & Trim(serie) & "' AND  doc_cod='" & Trim(doc_cod) & "' "
            CnBd.Execute (strCadena)
            If KEY_CERVECERIA = "si" Then
                 Dim nuevo_doc As String
                 Dim doc_numero As String
                 Dim doc_nuevo_n As String
                 Codkardex = CodigoKardex
                 Dim rst_fac As String
                 Dim tipo_doc As String
                 Dim cantidad As Single
                 Dim pprecio As Single
                 tipo_doc = "0000"
                 serie = "0001"
                 pprecio = 0#
                 Dim dfac As Boolean
                 strCadena = "SELECT numero FROM Det_alm_com WHERE serie='" & Trim(serie) & "' AND doc_cod='" & Trim(tipo_doc) & "' AND Alm_cod='0001'"
                 Call ConfiguraRst(strCadena)
                 doc_numero = rst(0)
                 Set rst = Nothing
                 strCadena = "SELECT     Producto_sub.cProducto AS Expr1,Detalle_DocumentoVenta.cantidad " & _
                 "FROM         Detalle_DocumentoVenta INNER JOIN " & _
                 "Producto ON Detalle_DocumentoVenta.cProducto = Producto.cProducto INNER JOIN " & _
                 "Producto_sub ON Producto.cProducto = Producto_sub.cProductoPadre WHERE cDocumentoVenta='" & Trim(Numero) & "' AND id_documentoventa='" & Trim(id_unico) & "' AND sSerie='" & Trim(serie) & "' AND  doc_cod='" & Trim(doc_cod) & "' "
                 
                Call ConfiguraRst(strCadena)
                If rst.RecordCount > 0 Then
                    rst.MoveFirst
                End If
                For i = 0 To rst.RecordCount - 1
                
                dfac = False
                Call Kardex(rst(0), tipo_doc, "0001", doc_numero, serie, KEY_ING, Date, Date, Val(rst(1)), Val(rst(1)), , _
                rst(1), pprecio, pprecio, , pprecio, "CLIENTE", _
                Trim(Codkardex), dfac)
                nuevo_doc = FormatosCeros(doc_numero + 1, 7)
                strCadena = "UPDATE Det_alm_com SET numero='" & Trim(nuevo_doc) & "' WHERE serie='" & Trim(serie) & "' AND doc_cod='" & Trim(tipo_doc) & "' AND Alm_cod='0001'"
                CnBd.Execute (strCadena)
                 rst.MoveNext
                Next i
                End If
            Unload Me
            FrmPendientes.HfgDetalle.Clear
            FrmPendientes.llenar
            
            Exit Sub
        End If
         If FrmPendientes.Procedencia = anular Then
            FrmPendientes.HfGuardado.col = 0
            id_unico = FrmPendientes.HfGuardado.Text
            FrmPendientes.HfGuardado.col = 1
            Numero = FrmPendientes.HfGuardado.Text
            FrmPendientes.HfGuardado.col = 2
            serie = FrmPendientes.HfGuardado.Text
             FrmPendientes.HfGuardado.col = 3
            doc_cod = FrmPendientes.HfGuardado.Text
            
            Call AnularVentas(Trim(doc_cod), Trim(serie), Trim(Numero), Trim(FrmVentas.DtcAlmacen.BoundText))
            FrmPendientes.Procedencia = Neutro
            FrmPendientes.HfgDetalle.Clear
            FrmPendientes.llenar
            Exit Sub
        End If
        
        If FrmPendientes.Procedencia = anular Then
            fecha = Date
            FrmPendientes.HfGuardado.col = 0
            id_unico = FrmPendientes.HfGuardado.Text
            FrmPendientes.HfGuardado.col = 1
            Numero = FrmPendientes.HfGuardado.Text
            FrmPendientes.HfGuardado.col = 2
            serie = FrmPendientes.HfGuardado.Text
            strCadena = "UPDATE DocumentoVenta SET dEmisionVenta='" & CVDate(fecha) & "',estado='Cancelado' WHERE cDocumentoVenta='" & Trim(Numero) & "' AND id_documentoventa='" & Trim(id_unico) & "' AND sSerie='" & Trim(serie) & "' AND id_usuario='" & Trim(KEY_USUARIO) & "' "
            Call EjecutaRST(strCadena)
            Set RstEjecuta = Nothing
            Unload Me
            FrmPendientes.llenar
            Exit Sub
        End If
        
        
        If FrmVentas.Procedencia = anular Then
            If KEY_CARGO_OPE = KEY_ADMIN Or KEY_CARGO_OPE = KEY_SUPER Or KEY_CARGO = "00001" Or KEY_CARGO = "00004" Or KEY_CARGO = "00006" Or KEY_CARGO = "00008" Then
                Call AnularVentas(Trim(FrmVentas.DtcTipoDoc.BoundText), Trim(FrmVentas.TxtSerie.Text), Trim(FrmVentas.TxtNumeroDoc.Text), Trim(FrmVentas.DtcAlmacen.BoundText))
                FrmVentas.Procedencia = Neutro
                Unload Me
                Exit Sub
            Else
                MsgBox "PASSWORD DE OPERACIONES INCORECCTO", vbInformation, KEY_EMPRESA
                FrmVentas.Procedencia = Neutro
                Unload Me
                Exit Sub
            End If
        End If
        
        If FrmVentas.Procedencia = Eliminar Then
            If KEY_CARGO_OPE = KEY_ADMIN Or KEY_CARGO_OPE = KEY_SUPER Then
                Call AnularVentas(Trim(FrmVentas.DtcTipoDoc.BoundText), Trim(FrmVentas.TxtSerie.Text), Trim(FrmVentas.TxtNumeroDoc.Text), Trim(FrmVentas.DtcAlmacen.BoundText))
                Call EliminarVentas(Trim(FrmVentas.DtcTipoDoc.BoundText), Trim(FrmVentas.TxtSerie.Text), Trim(FrmVentas.TxtNumeroDoc.Text), Trim(FrmVentas.DtcAlmacen.BoundText))
                FrmVentas.Procedencia = Neutro
                Unload Me
                Exit Sub
            Else
                MsgBox "PASSWORD DE OPERACIONES INCORECCTO", vbInformation, KEY_EMPRESA
                FrmVentas.Procedencia = Neutro
                Unload Me
                Exit Sub
            End If
        End If
        If FrmCompras.Procedencia = 5 Then
            Call AnularCompras(Trim(FrmCompras.DtcTipoDoc.BoundText), Trim(FrmCompras.TxtSerie.Text), Trim(FrmCompras.TxtNumeroDoc.Text), Trim(FrmCompras.DtcAlmacen.BoundText), Trim(FrmCompras.txtruc.Text))
            FrmCompras.Procedencia = Neutro
            Unload Me
            Exit Sub
        End If
        
        If FrmCompras.Procedencia = 3 Then
            Call EliminarCompras(Trim(FrmCompras.DtcTipoDoc.BoundText), Trim(FrmCompras.TxtSerie.Text), Trim(FrmCompras.TxtNumeroDoc.Text), Trim(FrmCompras.txtruc.Text))
            FrmCompras.Procedencia = Neutro
            Exit Sub
        End If
        
        If FrmDetalleGuia.Procedencia = 3 Then
            Call EliminarGuia(Trim(FrmDetalleGuia.TxtSerie_Guia.Text), Trim(FrmDetalleGuia.TxtNumero_guia.Text))
            FrmDetalleGuia.Procedencia = Neutro
            Exit Sub
        End If
        If FrmTransferencias.Procedencia = 5 Then
            'Call AnularTransferencia(Trim(FrmTransferencias.DtcTipoDoc.BoundText), Trim(FrmTransferencias.TxtSerie.text), Trim(FrmTransferencias.TxtNumeroDoc.text), Trim(FrmTransferencias.DtcAlmacenOrigen.BoundText), Trim(FrmTransferencias.DtcDestino.BoundText))
            FrmTransferencias.Procedencia = Neutro
            Exit Sub
        End If
        
        If FrmTransferencias.Procedencia = 3 Then
            'Call EliminarTransferencia(Trim(FrmTransferencias.DtcTipoDoc.BoundText), Trim(FrmTransferencias.TxtSerie.text), Trim(FrmTransferencias.TxtNumeroDoc.text), Trim(FrmTransferencias.DtcAlmacenOrigen.BoundText), Trim(FrmTransferencias.DtcDestino.BoundText))
            FrmTransferencias.Procedencia = Neutro
            Exit Sub
        End If
        If FrmAdelantoPersonal.Procedencia = 3 Then
            Call EliminarAdelanto(Trim(FrmAdelantoPersonal.DtcTipoDoc.BoundText), Trim(FrmAdelantoPersonal.TxtSerie.Text), Trim(FrmAdelantoPersonal.TxtNumeroDoc.Text), Trim(FrmAdelantoPersonal.txtruc.Text))
            FrmAdelantoPersonal.Procedencia = Neutro
            Exit Sub
        End If
        If FrmIngresoDinero.Procedencia = 3 Then
            Call EliminarIngresoDinero(Trim(FrmIngresoDinero.DtcTipoDoc.BoundText), Trim(FrmIngresoDinero.TxtSerie.Text), Trim(FrmIngresoDinero.TxtNumeroDoc.Text), Trim(FrmIngresoDinero.TxtCodCliente.Text))
               FrmIngresoDinero.Procedencia = Neutro
            Exit Sub
        End If
        If FrmPrecios.Procedencia = 2 Then
            Call FrmPrecios.LLenaDatos
            Call FrmPrecios.Resize
            FrmPrecios.Procedencia = Neutro
            Unload Me
            Exit Sub
            Unload Me
        End If
        If frmCajaEgreso.Procedencia = anular Then
            strCadena = "UPDATE mis_cuentas_det SET anulado='si' WHERE id='" & Val(frmCajaEgreso.HfDetalle.TextMatrix(frmCajaEgreso.HfDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            frmCajaEgreso.Procedencia = Neutro
            Unload Me
            Exit Sub
        End If
       If frmmistareas.Procedencia = Eliminar Then
       If KEY_USUARIO = "42546269" Then
          strCadena = "DELETE FROM proyecto_backlog WHERE id_codigo='" & frmmistareas.HfgLinea.TextMatrix(frmmistareas.HfgLinea.Row, 0) & "'"
          CnBd.Execute (strCadena)
          Call frmmistareas.backlog(frmmistareas.HfgLinea)
       End If
       frmmistareas.Procedencia = Neutro
       Unload Me
       Exit Sub
    End If

If frmmistareas.Procedencia = anular Then
        strCadena = "DELETE FROM proyecto_tareas WHERE id_tarea='" & Val(frmmistareas.HfTarea.TextMatrix(frmmistareas.HfTarea.Row, 0)) & "'"
        CnBd.Execute (strCadena)
        Call frmmistareas.tareas(frmmistareas.HfTarea, frmmistareas.txtid_backlog.Text)
        frmmistareas.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
       If frmmistareas.Procedencia = anular_asignacion Then ' eliminar un criterio de aceptacion
        strCadena = "DELETE FROM proyecto_testing WHERE id_testing='" & frmmistareas.HfCriterios.TextMatrix(frmmistareas.HfCriterios.Row, 0) & "'"
        CnBd.Execute (strCadena)
        Call frmmistareas.Criterio(frmmistareas.HfCriterios, frmmistareas.txtidtarea.Text)
        frmmistareas.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
 End If


End Sub
Sub anular_combo(ByVal id_combo As String)
strCadena = "UPDATE combo SET anulado='si' WHERE id_combo='" & id_combo & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
CnBd.Execute (strCadena)
strCadena = "SELECT * FROM combo WHERE id_combo='" & id_combo & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    frmCombo.TlbGrabar.Buttons(KEY_ANULAR).Enabled = False
    frmCombo.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
    frmCombo.lblAnulado.Visible = True
Else
    frmCombo.TlbGrabar.Buttons(KEY_ANULAR).Enabled = True
    frmCombo.lblAnulado.Visible = False
End If
End Sub
Sub anular_derivado(ByVal id_combo As String)
strCadena = "UPDATE derivados SET anulado='si' WHERE id_derivado='" & id_combo & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
CnBd.Execute (strCadena)
strCadena = "SELECT * FROM derivados WHERE id_derivado='" & id_combo & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    FrmDerivados.TlbGrabar.Buttons(KEY_ANULAR).Enabled = False
    FrmDerivados.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
    FrmDerivados.lblAnulado.Visible = True
Else
    FrmDerivados.TlbGrabar.Buttons(KEY_ANULAR).Enabled = True
    FrmDerivados.lblAnulado.Visible = False
End If
End Sub

Sub AnularVentas(ByVal TipoDoc As String, ByVal serie As String, ByVal Numero As String, ByVal Almacen As String)
 strCadena = "UPDATE movimiento_venta SET anulado='si',ncliente='A N U L A D O' WHERE id_alm='" & Almacen & "' AND id_doc='" & TipoDoc & "' AND serie='" & serie & "' AND numero='" & Numero & "' AND ruc='" & KEY_RUC & "'"
 CnBd.Execute (strCadena)
 
 
 
 
 
 strCadena = "SELECT * FROM movimiento_venta WHERE id_alm='" & Almacen & "' AND id_doc='" & TipoDoc & "' AND serie='" & serie & "' AND numero='" & Numero & "' AND ruc='" & KEY_RUC & "' AND anulado='si'"
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        rstZ.MoveFirst
        For m = 0 To rstZ.RecordCount - 1
            strCadena = "UPDATE imp_producto_detalle SET vendido='no' WHERE nro_chasis='" & rstZ("nro_chasis") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            rstZ.MoveNext
        Next m
    End If
 
    FrmVentas.lblAnulado.Visible = True
    FrmVentas.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
  Else
  FrmVentas.lblAnulado.Visible = True
  FrmVentas.lblAnulado.Caption = "ERROR AL ANULAR"
 End If
 
 Set rst = Nothing

End Sub
Sub Imprimir_motivo(ByVal TipoDoc As String, ByVal Almacen As String, ByVal serie As String, ByVal Numero As String)
Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "FontB11"
    Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    

    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print Tab(5); "PEPE'S AUTOSERVICIOS S.A.C"
    Printer.Print Tab(2); "-----------------------------------"
    Printer.Print Tab(5); "ANULACION DE COMPROBANTES"
    If TipoDoc = "0001" Then
        des_comp = "TIKET FACT:"
    End If
    If TipoDoc = "0003" Then
        des_comp = "TIKET BOL:"
    End If
    Printer.Print Tab(0); des_comp; Space(1); Mid(serie + Space(50), 1, 4) & "-" & Numero
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.Print Tab(1); "===================================="
    Printer.Print Tab(3); " FECHA/ HORA   :" & Space(3) & Str(Date) & Space(3) & Trim(Time)
    Printer.Print Tab(3); " MOTIVO DE ANULACION:"
    Printer.Print Tab(3); Trim(Me.txtMotivo.Text)
    Printer.Print Tab(1); "===================================="
    Printer.EndDoc
End Sub
Sub AnularCompras(ByVal TipoDoc As String, ByVal serie As String, ByVal Numero As String, ByVal Almacen As String, ByVal id_proveedor As String)
strCadena = "UPDATE movimiento_compra SET anulado='si' WHERE id_doc='" & TipoDoc & "' AND serie='" & serie & "' AND numero='" & Numero & "' AND id_alm='" & Almacen & "' AND id_proveedor='" & id_proveedor & "' AND ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
FrmCompras.lblAnulado.Visible = True
FrmCompras.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False

End Sub
Sub AnularTransferencia(ByVal TipoDoc As String, ByVal serie As String, ByVal Numero As String, ByVal AlmacenOrigen As String, ByVal AlmacenDestino As String)
Dim Anulado As String * 1
Anulado = "V"
'---------------
    Dim RstDetTransferencia As New ADODB.Recordset
    strCadena = "SELECT cProducto,cantidad FROM Detalle_DocumentoTransferencia WHERE (doc_cod='" & TipoDoc & "'AND sSerie ='" & serie & "'AND cDocumentoTransferencia='" & Numero & "' AND Alm_Origen='" & AlmacenOrigen & "' AND Alm_Destino='" & AlmacenDestino & "' ) ORDER BY cProducto ASC"
    RstDetTransferencia.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
    
    strCadena = "SELECT Almacen_Productos.cProducto,Stock FROM Almacen_Productos  WHERE Almacen_Productos.Alm_cod='" & AlmacenOrigen & "' ORDER BY Almacen_Productos.cProducto ASC"
    Call ConfiguraRst(strCadena)
    RstDetTransferencia.MoveFirst
    rst.MoveFirst
    For i = 0 To RstDetTransferencia.RecordCount - 1
        For j = 0 To rst.RecordCount - 1
            If RstDetTransferencia(0) = rst(0) Then
                rst(1) = rst(1) + RstDetTransferencia(1)
                rst.Update
                Exit For
            Else
                rst.MoveNext
            End If
         Next j
         RstDetTransferencia.MoveNext
         rst.MoveFirst
    Next i
'Almacen Destino---
Set rst = Nothing
strCadena = "SELECT Almacen_Productos.cProducto,Stock FROM Almacen_Productos  WHERE Almacen_Productos.Alm_cod='" & AlmacenDestino & "' ORDER BY Almacen_Productos.cProducto ASC"
    Call ConfiguraRst(strCadena)
    RstDetTransferencia.MoveFirst
    rst.MoveFirst
    For i = 0 To RstDetTransferencia.RecordCount - 1
        For j = 0 To rst.RecordCount - 1
            If RstDetTransferencia(0) = rst(0) Then
                rst(1) = rst(1) - RstDetTransferencia(1)
                rst.Update
                Exit For
            Else
                rst.MoveNext
            End If
         Next j
         RstDetTransferencia.MoveNext
         rst.MoveFirst
    Next i
    Set RstDetTransferencia = Nothing
    Set rst = Nothing
'---------------
strCadena = "UPDATE DocumentoTransferencia SET Anulado='" & Anulado & "' WHERE (cDocumentoTransferencia='" & Numero & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "'AND Alm_Origen='" & AlmacenOrigen & "' AND Alm_Destino='" & AlmacenDestino & "')"
Call EjecutaRST(strCadena)
Set rst = Nothing
strCadena = "DELETE Kardex  WHERE (NumeroDoc='" & Numero & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "' AND Alm_cod='" & Almacen & "')"
Call EjecutaRST(strCadena)
Set rst = Nothing
Unload Me
'FrmTransferencias.lblAnulado.Visible = True
'FrmTransferencias.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
End Sub
Sub EliminarTransferencia(ByVal TipoDoc As String, ByVal serie As String, ByVal Numero As String, ByVal AlmacenOrigen As String, ByVal AlmacenDestino As String)
'---------------------------
    Dim RstDetTransferencia As New ADODB.Recordset
    strCadena = "SELECT cProducto,cantidad FROM Detalle_DocumentoTransferencia WHERE (doc_cod='" & TipoDoc & "'AND sSerie ='" & serie & "'AND cDocumentoTransferencia='" & Numero & "' AND Alm_Origen='" & AlmacenOrigen & "' AND Alm_Destino='" & AlmacenDestino & "') ORDER BY cProducto ASC"
    RstDetTransferencia.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
    'Almacen Origen
    strCadena = "SELECT Almacen_Productos.cProducto,Stock FROM Almacen_Productos  WHERE Almacen_Productos.Alm_cod='" & AlmacenOrigen & "' ORDER BY Almacen_Productos.cProducto ASC"
    Call ConfiguraRst(strCadena)
    RstDetTransferencia.MoveFirst
    rst.MoveFirst
    For i = 0 To RstDetTransferencia.RecordCount - 1
        For j = 0 To rst.RecordCount - 1
            If RstDetTransferencia(0) = rst(0) Then
                rst(1) = rst(1) + RstDetTransferencia(1)
                rst.Update
                Exit For
            Else
                rst.MoveNext
            End If
         Next j
         RstDetTransferencia.MoveNext
         rst.MoveFirst
    Next i
    Set rst = Nothing
    'Almacen Destino
    strCadena = "SELECT Almacen_Productos.cProducto,Stock FROM Almacen_Productos  WHERE Almacen_Productos.Alm_cod='" & AlmacenDestino & "' ORDER BY Almacen_Productos.cProducto ASC"
    Call ConfiguraRst(strCadena)
    RstDetTransferencia.MoveFirst
    rst.MoveFirst
    For i = 0 To RstDetTransferencia.RecordCount - 1
        For j = 0 To rst.RecordCount - 1
            If RstDetTransferencia(0) = rst(0) Then
                rst(1) = rst(1) - RstDetTransferencia(1)
                rst.Update
                Exit For
            Else
                rst.MoveNext
            End If
         Next j
         RstDetTransferencia.MoveNext
         rst.MoveFirst
    Next i
    Set rst = Nothing
    Set RstDetTransferencia = Nothing
'----------------------
strCadena = "DELETE DocumentoTransferencia  WHERE (cDocumentoTransferencia='" & Numero & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "' AND Alm_Origen='" & AlmacenOrigen & "' AND Alm_Destino='" & AlmacenDestino & "')"
Call EjecutaRST(strCadena)
Set rst = Nothing
strCadena = "DELETE Kardex  WHERE (NumeroDoc='" & Numero & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "' AND Alm_cod='" & Almacen & "')"
Call EjecutaRST(strCadena)
Set rst = Nothing
Unload Me
'FrmTransferencias.Nuevo
'FrmTransferencias.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
End Sub
Sub EliminarVentas(ByVal TipoDoc As String, ByVal serie As String, ByVal Numero As String, ByVal Almacen As String)
 strCadena = "SELECT * FROM movimiento_venta WHERE id_alm='" & Almacen & "' AND id_doc='" & TipoDoc & "' AND serie='" & serie & "' AND numero='" & Numero & "' AND ruc='" & KEY_RUC & "' AND anulado='si'"
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        rstZ.MoveFirst
        For m = 0 To rstZ.RecordCount - 1
            strCadena = "UPDATE imp_producto_detalle SET vendido='no' WHERE nro_chasis='" & rstZ("nro_chasis") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            rstZ.MoveNext
        Next m
    End If
 End If
 
 
 strCadena = "DELETE FROM movimiento_venta WHERE id_alm='" & Almacen & "' AND id_doc='" & TipoDoc & "' AND serie='" & serie & "' AND numero='" & Numero & "' AND ruc='" & KEY_RUC & "'"
 CnBd.Execute (strCadena)
 FrmVentas.Nuevo
End Sub
Sub EliminarAdelanto(ByVal TipoDoc As String, ByVal serie As String, ByVal Numero As String, ByVal Persona As String)
                 
strCadena = "DELETE DocumentoVenta  WHERE (cDocumentoVenta='" & Numero & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "' AND cpersona='" & Persona & "')"
Call EjecutaRST(strCadena)
Set rst = Nothing
Unload Me
FrmAdelantoPersonal.Nuevo
'FrmAdelantoPersonal.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
End Sub
Sub EliminarIngresoDinero(ByVal TipoDoc As String, ByVal serie As String, ByVal Numero As String, ByVal Persona As String)
              
strCadena = "DELETE DocumentoVenta  WHERE (cDocumentoVenta='" & Numero & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "' AND cpersona='" & Persona & "')"
Call EjecutaRST(strCadena)
Set rst = Nothing
Unload Me
FrmIngresoDinero.Nuevo
FrmIngresoDinero.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
End Sub

Sub EliminarCompras(ByVal id_doc As String, ByVal serie As String, ByVal Numero As String, ByVal id_proveedor As String)
Dim idCompra As Double
strCadena = "SELECT * FROM movimiento_compra WHERE (numero='" & Numero & "' AND serie='" & serie & "' AND id_doc='" & id_doc & "' AND id_proveedor='" & id_proveedor & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    idCompra = rst("id_compra")
    strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & idCompra & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraTemporal(strCadena)
    If rstTemporal.RecordCount > 0 Then
        rstTemporal.MoveFirst
        For i = 0 To rstTemporal.RecordCount - 1
            strCadena = "DELETE FROM movimiento_compra_detalle WHERE id_compra='" & idCompra & "' AND id_detalle_compra='" & rstTemporal("id_detalle_compra") & "'"
            CnBd.Execute (strCadena)
            rstTemporal.MoveNext
        Next i

    End If
     strCadena = "DELETE FROM movimiento_compra WHERE id_compra='" & idCompra & "' AND ruc='" & KEY_RUC & "'"
      CnBd.Execute (strCadena)
     'strCadena = "DELETE FROM imp_producto_detalle WHERE id_compra='" & idCompra & "' and ruc='" & KEY_RUC & "'"
     'CnBd.Execute (strCadena)
    
    
End If

FrmCompras.Nuevo
FrmCompras.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
Unload Me
End Sub

Sub EliminarGuia(ByVal serie As String, ByVal Numero As String)
strCadena = "DELETE Detalleguia  WHERE (sNumeroGuia='" & Numero & "' AND sSerieGuia='" & serie & "')"
Call EjecutaRST(strCadena)
Set rst = Nothing
Unload Me
FrmDetalleGuia.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
End Sub

Private Sub TxtMotivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
    Exit Sub
End If
If KeyAscii = 13 Then
    Call Resalta(Me.TxtClave)
End If
End Sub

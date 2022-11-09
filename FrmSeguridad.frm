VERSION 5.00
Begin VB.Form FrmSeguridad 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMotivo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1080
      Width           =   3250
   End
   Begin VB.TextBox TxtClave 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   1815
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
Public Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If frmCajaEgreso.Procedencia = cerrarcaja Then
       Call enabled_form(frmCajaEgreso)
       Call enabled_form(FrmMiscuentas)
       Unload Me
       Exit Sub
    End If
    
    If FrmVentas.Procedencia = modificar_credito Then
       FrmVentas.Procedencia = Neutro
       Call enabled_form(FrmVentas)
       Unload Me
       Exit Sub
    End If
    
    
    If FrmVentas.Procedencia = imprimir_s Then
       FrmVentas.Procedencia = Neutro
       Call enabled_form(FrmVentas)
       Unload Me
       Exit Sub
    End If
    
    
    If FrmTransferencias.Procedencia = modificar Then
       FrmTransferencias.Procedencia = Neutro
       Call enabled_form(FrmTransferencias)
       Unload Me
       Exit Sub
    End If
    
    If FrmTipocambio.Procedencia = modificar Then
       FrmTipocambio.Procedencia = Neutro
       Call enabled_form(FrmTipocambio)
       Unload Me
       Exit Sub
    End If
    
    
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    If FrmVentas.Procedencia = anular Or FrmVentas.Procedencia = Eliminar Or FrmVentas.Procedencia = seleccionar_vendedor Or FrmVentas.Procedencia = modificar_credito Then
        
        FrmVentas.Procedencia = Neutro
         Call enabled_form(FrmVentas)
        Unload Me
        Exit Sub
    End If
    
    
    
    If FrmTipocambio.Procedencia = nuevo Then
       Call enabled_form(FrmTipocambio)
       Unload Me
       Exit Sub
    End If
    
     
    
    If frmCajaEgreso.Procedencia = cerrarcaja Then
       Call enabled_form(frmCajaEgreso)
       Call enabled_form(FrmMiscuentas)
       Unload Me
       Exit Sub
    End If
    
    
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
Dim numero As String
Dim serie As String
Dim doc_cod As String
Dim cPersona As String
If KeyAscii = 27 Then
        If FrmVentas.Procedencia = anular Or FrmVentas.Procedencia = Eliminar Or FrmVentas.Procedencia = seleccionar_vendedor Then
            FrmVentas.Enabled = True
            FrmVentas.Procedencia = Neutro
        End If
        Unload Me
        Exit Sub
End If


If KeyAscii = 13 Then

If FrmParametrosEmpresa.Procedencia = nuevo Or FrmParametrosEmpresa.Procedencia = modificar Then
    strCadena = "SELECT * FROM entidad_empresa WHERE password='" & Trim(Me.TxtClave.Text) & "' AND id_empresa='" & KEY_RUC & "' AND cod_unico='" & KEY_USUARIO & "' and id_cargo='00004' "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
            FrmDetallesParametros.Show
            Unload Me
            FrmParametrosEmpresa.Procedencia = Neutro
            Exit Sub
        
    End If
End If
       
       
       strCadena = "SELECT * FROM entidad_empresa E WHERE password='" & Trim(Me.TxtClave.Text) & "'  AND id_empresa='" & KEY_RUC & "' LIMIT 1"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount < 1 Then
            MsgBox "PASSWORD DE OPERACIONES INCORRECTA", vbInformation, KEY_EMPRESA
               ' Unload Me
               ' Exit Sub
            
            If FrmVentas.Procedencia = seleccionar_vendedor Then
                Call enabled_form(FrmVentas)
                Unload Me
                Exit Sub
            End If
                'Call enabled_form(FrmVentas)
                Unload Me
                Exit Sub
            Exit Sub
       Else
            rst.MoveFirst
            KEY_CARGO_OPE = rst("id_cargo")
            KEY_VENDEDOR_AMBULANTE = rst("cod_unico")
       End If
        
      If FrmTipocambio.Procedencia = modificar Then
        
        
        
        FrmDetalleTipocaambio.Show
        FrmTipocambio.Procedencia = Neutro
        FrmDetalleTipocaambio.Dtp_fecha.Enabled = False
        FrmDetalleTipocaambio.TxtId_codigo.Text = Val(FrmTipocambio.HfgCambio.TextMatrix(FrmTipocambio.HfgCambio.Row, 0))
        Call FrmDetalleTipocaambio.LLENA(FrmTipocambio.HfgCambio.TextMatrix(FrmTipocambio.HfgCambio.Row, 0))
        Call enabled_form(FrmTipocambio)
        
        
        Unload Me
        
        Exit Sub
    End If
    
    
    If FrmTipocambio.Procedencia = nuevo Then
        FrmDetalleTipocaambio.Show
        FrmDetalleTipocaambio.cmdSbs.Visible = True
        FrmDetalleTipocaambio.Dtp_fecha.Value = KEY_FECHA
         
        FrmDetalleTipocaambio.Dtp_fecha.Enabled = True
        FrmDetalleTipocaambio.TxtId_codigo.Text = 0
        Unload Me
        
        FrmTipocambio.Procedencia = Neutro
        Exit Sub
    End If
    
    
    
    
    
    If frmCajaEgreso.Procedencia = cerrarcaja Then
       frmCajaEgreso.Procedencia = Neutro
       Call cerrar_caja(Val(frmCajaEgreso.TxtidCuenta.Text))
       Call enabled_form(frmCajaEgreso)
       Call enabled_form(FrmMiscuentas)
       Unload Me
       Exit Sub
    End If
    
    
    
    If FrmVentas.Procedencia = modificar_credito Then
         FrmVentas.Procedencia = Neutro
        in_pass_admin = verificar_password_admin(Me.TxtClave.Text)
        If in_pass_admin = "00004" Or in_pass_admin = "00003" Or in_pass_admin = "00009" Then
           
           If KEY_RUC = "20128836251" Then ' vargas
           If KEY_USUARIO = "10002108" Or KEY_USUARIO = "43321337" Or KEY_USUARIO = "42546269" Or KEY_USUARIO = "00122875" Then
                FrmVentas.frmcredito.Visible = True
                Unload Me
                Call enabled_form(FrmVentas)
                Exit Sub
           Else
                MsgBox "USTED NO ESTA AUTORIZADO." + Chr(13) + "PARA DAR CREDITOS..", vbInformation, KEY_VENDEDOR
                Call enabled_form(FrmVentas)
                Unload Me
                Exit Sub
           End If
           Else
                FrmVentas.frmcredito.Visible = True
                Unload Me
                Call enabled_form(FrmVentas)
                Exit Sub
           End If
           
           
           
           
        Else
           MsgBox "INGRESE UNA PASSWORD SUPERIOR !!!", vbInformation, KEY_VENDEDOR
           Call enabled_form(FrmVentas)
           Unload Me
           Exit Sub
        End If
       FrmVentas.frmcredito.Visible = True
next_sig:
        strCadena = "SELECT monto_credito FROM entidad_empresa WHERE cod_unico='" & FrmVentas.TxtCodCliente.Text & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            FrmVentas.txtmontocredito.Text = rst("monto_credito")
            Call Resalta(FrmVentas.txtmontocredito)
            FrmVentas.Procedencia = Neutro
            Call enabled_form(FrmVentas)
            Unload Me
            Exit Sub
        Else
            Call put_persona(Trim(FrmVentas.TxtCodCliente.Text))
            GoTo next_sig
        End If
    End If
    
    
    If FrmVentas.Procedencia = modificar Then
       FrmVentas.Procedencia = Neutro
       If FrmVentas.DtcTipoDoc.BoundText = "0099" Then
        FrmVentas.cmdProcesar.Enabled = True
        Call FrmVentas.get_update_proform
       End If
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
       If KEY_CARGO <> "00004" Then
            FrmVentas.DtcVendedor.BoundText = KEY_VENDEDOR_AMBULANTE
       
        End If
       If Val(FrmVentas.txtidVenta.Text) < 1 Then
            Call FrmVentas.get_auto_pago(FrmVentas.DtcTipoDoc.BoundText)
            Call FrmVentas.Save
            If KEY_CARGO = "00008" Then
                strCadena = "UPDATE movimiento_venta SET pendiente='no' WHERE id_venta='" & Val(FrmVentas.txtidVenta.Text) & "'"
                CnBd.Execute (strCadena)
                 
                Call impresion_pedido(Val(FrmVentas.txtidVenta.Text))
            End If
        Else
            Call AnularVentas(Trim(FrmVentas.DtcTipoDoc.BoundText), Trim(FrmVentas.DtcSerieDoc.BoundText), Trim(FrmVentas.TxtNumeroDoc.Text), Trim(FrmVentas.DtcAlmacen.BoundText))
            Call EliminarVentas(Trim(FrmVentas.DtcTipoDoc.BoundText), Trim(FrmVentas.DtcSerieDoc.BoundText), Trim(FrmVentas.TxtNumeroDoc.Text), Trim(FrmVentas.DtcAlmacen.BoundText))
            Call FrmVentas.get_auto_pago(FrmVentas.DtcTipoDoc.BoundText)
            Call FrmVentas.Save
            Call impresion_pedido(Val(FrmVentas.txtidVenta.Text))
        End If
        If FrmVentas.DtcTipoDoc.BoundText = "0099" And KEY_RUC = "20561193291" Then
            FrmVentas.Procedencia = Neutro
            Unload Me
            FrmVentas.Enabled = True
            FrmVentas.cmdProcesar.Enabled = False
            FrmVentas.cmdImprimir.Enabled = True
            Exit Sub
        Else
        Call FrmVentas.nuevo
        FrmVentas.Procedencia = Neutro
        Unload Me
        FrmVentas.Enabled = True
       Exit Sub
        End If
        
        
    End If
    
  '  If FrmVentas.Procedencia = Modificar Then
  '     strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & FrmVentas.TxtIdVenta.Text & "' and ruc='" & KEY_RUC & "'"
  '     Call ConfiguraRstZ(strCadena)
  '     If rstZ.RecordCount > 0 Then
  '        rstZ.MoveFirst
  '        strCadena = "DELETE FROM temporal_ventas WHERE id_doc='" & FrmVentas.DtcTipoDoc.BoundText & "' and numero='" & FrmVentas.TxtNumeroDoc.Text & "' and id_serie='" & FrmVentas.DtcSerieDoc.BoundText & "' and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
  '        CnBd.Execute (strCadena)
  '
          
  '        For i = 0 To rstZ.RecordCount - 1
  '          strCadena = "INSERT INTO temporal_ventas(ruc,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,dni_save) VALUES " & _
  '          "('" & KEY_RUC & "','" & KEY_ALM & "','" & FrmVentas.DtcTipoDoc.BoundText & "','" & Trim(FrmVentas.DtcSerieDoc.BoundText) & "','" & FrmVentas.TxtNumeroDoc.Text & "','" & rstZ("id_producto") & "','" & rstZ("cantidad") & "'," & _
  '          "'" & rstZ("precio") & " ','" & rstZ("total") & "','0','si','" & KEY_USUARIO & "')"
  '          CnBd.Execute (strCadena)
             
  '          rstZ.MoveNext
  '        Next i
  '        Call FrmVentas.llenarGrid_det(FrmVentas.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText, FrmVentas.txtformato_impresion.Text)
          
 '         FrmVentas.cmdProcesar.Enabled = True
  '     End If
   '    FrmVentas.Procedencia = Neutro
    '   Unload Me
   '    Exit Sub
   ' End If
    
    
    If FrmFechaTrabajo.Procendencia = buscar Then
         
         If KEY_CARGO = KEY_ADMINISTRADOR Or KEY_CARGO = KEY_SUPERVISOR Or KEY_CARGO = KEY_AUTOR Then
            FrmFechaTrabajo.DTPicker1.Enabled = True
            FrmFechaTrabajo.DtcTurno.Locked = False
         Else
            MsgBox "PASSWORD INCORRECTO", vbQuestion, KEY_EMPRESA
         End If
         Unload Me
         FrmFechaTrabajo.Procendencia = Neutro
         Exit Sub
    End If
    
    
    
    
    If FrmFechaTrabajo.Procendencia = Eliminar Then
         
         If KEY_CARGO = KEY_ADMINISTRADOR Or KEY_CARGO = KEY_SUPERVISOR Or KEY_CARGO = KEY_AUTOR Then
            strCadena = "DELETE FROM impresora where id_impresora='" & Val(FrmFechaTrabajo.HfPrinter.TextMatrix(FrmFechaTrabajo.HfPrinter.Row, 0)) & "' and id_alm='" & FrmFechaTrabajo.DtcAlmacen.BoundText & "' and  ruc='" & FrmFechaTrabajo.DtcEmpresa.BoundText & "'"
            CnBd.Execute (strCadena)
             
            Call FrmFechaTrabajo.llenar_printer(FrmFechaTrabajo.HfPrinter, KEY_RUC)
         Else
            MsgBox "USTED NO CUENTA CON EL PERMISO PARA REALIZAR ESTE PROCESO", vbQuestion, KEY_EMPRESA
         End If
         Unload Me
         FrmFechaTrabajo.Procendencia = Neutro
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
            strCadena = "SELECT id_tipo_factura FROM movimiento_venta WHERE id_venta='" & Val(FrmVentas.txtidVenta.Text) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                strCadena = "call p_update_impresiones('" & Val(FrmVentas.txtidVenta.Text) & "')"
                CnBd.Execute (strCadena)
                Call Orden_Impresion(FrmVentas.DtcTipoDoc.BoundText, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.TxtNumeroDoc.Text, rst("id_tipo_factura"), FrmVentas.txtidVenta.Text)
                FrmVentas.Procedencia = Neutro
                Call enabled_form(FrmVentas)
                Unload Me
       
                
                Exit Sub
            End If
    End If
    
   
   
   If FrmVentasPersonalizada.Procedencia = imprimir_s Then
            If KEY_CARGO = KEY_ADMIN Or KEY_CARGO = KEY_SUPER Or KEY_CARGO = KEY_AUTOR Then
            strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & FrmVentasPersonalizada.DtcTipoDoc.BoundText & "' AND serie='" & FrmVentasPersonalizada.txtserie.Text & "' AND numero='" & FrmVentasPersonalizada.TxtNumeroDoc.Text & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRst(strCadena)
            impresiones = rst("impresiones") + 1
            strCadena = "UPDATE movimiento_venta SET impresiones='" & impresiones & "' WHERE id_venta='" & rst("id_venta") & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
             
            
            Call Orden_Impresion(FrmVentas.DtcTipoDoc.BoundText, FrmVentasPersonalizada.txtserie.Text, FrmVentasPersonalizada.TxtNumeroDoc.Text, rst("id_tipo_factura"), FrmVentas.txtidVenta.Text)
            
         Else
            MsgBox "PASSWORD INCORRECTO", vbInformation, KEY_EMPRESA
            
         End If
         Unload Me
         FrmVentas.Procedencia = Neutro
         Exit Sub
    End If
    
    If frmCorProcesos.Procedencia = modificar Then
       Select Case frmCorProcesos.Txtid_estado
            Case "01"
                
                strCadena = "DELETE FROM imp_producto_movimiento WHERE id_detalle='" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "' and id_proceso='01' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                 
                 strCadena = "UPDATE imp_producto_detalle set id_estado='01' WHERE id_detalle='" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "'"
                CnBd.Execute (strCadena)
                 
                
                'strCadena = "UPDATE imp_producto_movimiento set estado='0'  WHERE id_detalle='" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "' and id_proceso='01' and ruc='" & KEY_RUC & "'"
                'CnBd.Execute (strCadena)   Call seleccion
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
    
    
    
    If FrmTransferencias.Procedencia = modificar Then
        FrmTransferencias.Procedencia = Neutro
        strCadena = "SELECT * FROM movimiento_transferencia WHERE id_transferencia='" & Val(FrmTransferencias.TxtId_transferencia.Text) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            If rst("id_motivo") = 3 Then  ' grinter
               If rst("id_alm_destino") <> KEY_ALM Then
                  MsgBox "ALMACEN DIFERENTE AL DESTINO" + Chr(13) + "Ingrese al Almacen Correcto para Recepcionar", vbInformation, KEY_VENDEDOR
                  Exit Sub
               End If
            End If
            
            
            FrmTransferencias.HfSeries.Enabled = True
            FrmTransferencias.cmdProcesar.Enabled = True
            FrmTransferencias.cmdverificar.Enabled = False
        End If
        Call enabled_form(FrmTransferencias)
        Unload Me
        Exit Sub
    End If
    
    
    
    If FrmTransferencias.Procedencia = Eliminar Then
        
         If get_periodo_cierre_fecha(FrmTransferencias.DtpFechaEmision.Value) = True Then
                 MsgBox "PERIODO CONTABLE.!!!" + Chr(13) + "YA ESTA CERRADO.", vbInformation, KEY_VENDEDOR
                 Exit Sub
          End If
             
        
        strCadena = "SELECT id_transferencia,id_alm_origen as id_alm,id_alm_destino,id_venta,id_motivo,fecha FROM movimiento_transferencia WHERE id_doc='" & FrmTransferencias.DtcTipoDoc.BoundText & "' AND serie='" & Trim(FrmTransferencias.DtcSerieGuia.BoundText) & "' AND numero='" & FrmTransferencias.TxtNumeroDoc.Text & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
        in_motivo = rstZ("id_motivo")
            If get_diferida(rstZ("id_venta")) = "si" Then
                    strCadena = "call CON_InsertaAsiento_GuiaDiferida_Extorno('" & rstZ("id_transferencia") & "')"
                    CnBd.Execute (strCadena)
            End If
            
            
            
            
            strCadena = "SELECT chasis,id_producto FROM movimiento_transferencia_series WHERE id_transferencia='" & rstZ("id_transferencia") & "'"
            Call ConfiguraRstIN(strCadena)
            If rstIN.RecordCount > 0 Then
                For p = 0 To rstIN.RecordCount - 1
                    
                    strCadena = "DELETE FROM kardex WHERE  id_doc='0009' and id_serie='" & Trim(FrmTransferencias.DtcSerieGuia.Text) & "' and id_movimiento='" & rstZ("id_transferencia") & "' and   id_alm='" & rstZ("id_alm") & "' and  id_producto='" & rstIN("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    If in_motivo = "3" Then ' transferencias
                            strCadena = "DELETE FROM kardex WHERE  id_doc='0009' and id_serie='" & Trim(FrmTransferencias.DtcSerieGuia.Text) & "' and id_movimiento='" & rstZ("id_transferencia") & "' and   id_alm='" & rstZ("id_alm_destino") & "' and  id_producto='" & rstIN("id_producto") & "' and ruc='" & KEY_RUC & "'"
                            CnBd.Execute (strCadena)
                            
                            strCadena = "UPDATE imp_producto_detalle SET transferencia='no',ruc='0'  WHERE id_producto='" & rstIN("id_producto") & "' and  nro_chasis='" & rstIN("chasis") & "' and id_alm='" & rstZ("id_alm_destino") & "' and ruc='" & KEY_RUC & "'"
                            CnBd.Execute (strCadena)
                    End If
                    
                    ':::::::::::ACTUALIZA KARDEX
                     If KEY_RUC = "20128836251" Then
                            Call update_kardex_Vargas_modulo_compra(rstIN("id_producto"), Format(rstZ("fecha"), "YYYY-mm-dd"))
                     Else
                            Call update_kardex_update(rstIN("id_producto"), Format(rstZ("fecha"), "YYYY-mm-dd"))
                     End If
                     '--------------FIN KARDEX
                        
                        
                        
                        
                    
                    strCadena = "UPDATE imp_producto_detalle SET transferencia='no'  WHERE nro_chasis='" & rstIN("chasis") & "' and id_alm='" & rstZ("id_alm") & "' and ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                     
                    strCadena = "SELECT sum(cantidad_real),sum(cantidad_pendiente),sum(cantidad_contable) FROM kardex where id_alm='" & rstZ("id_alm") & "' and  id_producto='" & rstIN("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstK(strCadena)
                    
                    strCadena = "UPDATE almacen_producto set stock = '" & rstK(0) & "'  , `stock_contable` = '" & rstK(1) & "',stock_factura='" & rstK(2) & "'  where ruc = '" & KEY_RUC & "' and id_alm ='" & rstZ("id_alm") & "' and id_producto = '" & rstIN("id_producto") & "'"
                    CnBd.Execute (strCadena)
                     
                    
                    strCadena = "SELECT sum(cantidad_real),sum(cantidad_pendiente),sum(cantidad_contable) FROM kardex where id_alm='" & rstZ("id_alm_destino") & "' and  id_producto='" & rstIN("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstK(strCadena)
                    
                    strCadena = "update almacen_producto set stock = '" & rstK(0) & "'  , `stock_contable` = '" & rstK(1) & "',stock_factura='" & rstK(2) & "'  where ruc = '" & KEY_RUC & "' and id_alm ='" & rstZ("id_alm_destino") & "' and id_producto = '" & rstIN("id_producto") & "'"
                    CnBd.Execute (strCadena)
                     
    
                    rstIN.MoveNext
                Next p
                strCadena = "DELETE FROM movimiento_transferencia WHERE id_transferencia='" & rstZ("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                 
            Else
                 
                 
                 strCadena = "SELECT * FROM movimiento_transferencia_detalle WHERE id_transferencia='" & rstZ("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
                 Call ConfiguraRstA(strCadena)
                 If rstA.RecordCount > 0 Then
                    rstA.MoveFirst
                    For i = 0 To rstA.RecordCount - 1
                        
                        strCadena = "DELETE FROM kardex WHERE  id_doc='0009' and id_serie='" & Trim(FrmTransferencias.DtcSerieGuia.Text) & "' and id_movimiento='" & rstZ("id_transferencia") & "' and   id_alm='" & rstZ("id_alm") & "' and  id_producto='" & rstA("id_producto") & "' and ruc='" & KEY_RUC & "'"
                        CnBd.Execute (strCadena)
                        
                        If KEY_RUC = "20128836251" Then
                            Call update_kardex_Vargas_modulo_compra(rstA("id_producto"), Format(rstZ("fecha"), "YYYY-mm-dd"))
                        Else
                            Call update_kardex_update(rstA("id_producto"), Format(rstZ("fecha"), "YYYY-mm-dd"))
                        End If
            
                        
                        
                        If in_motivo = "3" Then ' transferencias
                            strCadena = "DELETE FROM kardex WHERE  id_doc='0009' and id_serie='" & Trim(FrmTransferencias.DtcSerieGuia.Text) & "' and id_movimiento='" & rstZ("id_transferencia") & "' and   id_alm='" & rstZ("id_alm_destino") & "' and  id_producto='" & rstA("id_producto") & "' and ruc='" & KEY_RUC & "'"
                            CnBd.Execute (strCadena)
                        End If
                        
                        
                        strCadena = "SELECT sum(cantidad_real),sum(cantidad_pendiente),sum(cantidad_contable) FROM kardex where id_alm='" & rstZ("id_alm") & "' and  id_producto='" & rstA("id_producto") & "' and ruc='" & KEY_RUC & "'"
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
                        If IsNull(rstK(2)) = True Then
                            in_contable = 0
                        Else
                            in_contable = rstK(1)
                        End If
                        
                    
                    
                    strCadena = "UPDATE almacen_producto set stock = '" & in_real & "'  , `stock_contable` = '" & in_pendiente & "',stock_factura='" & in_contable & "'   where ruc = '" & KEY_RUC & "' and id_alm ='" & rstZ("id_alm") & "' and id_producto = '" & rstA("id_producto") & "'"
                    CnBd.Execute (strCadena)
                     
                    
                    strCadena = "SELECT sum(cantidad_real),sum(cantidad_pendiente),sum(cantidad_contable) FROM kardex where id_alm='" & rstZ("id_alm_destino") & "' and  id_producto='" & rstA("id_producto") & "' and ruc='" & KEY_RUC & "'"
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
                        If IsNull(rstK(1)) = True Then
                            in_contable = 0
                        Else
                            in_contable = rstK(1)
                        End If
                    
                    strCadena = "UPDATE almacen_producto set stock = '" & in_real & "' ,stock_factura='" & in_contable & "' , `stock_contable` = '" & in_pendiente & "'  where ruc = '" & KEY_RUC & "' and id_alm ='" & rstZ("id_alm_destino") & "' and id_producto = '" & rstA("id_producto") & "'"
                    CnBd.Execute (strCadena)
                    rstA.MoveNext
                    
                    Next i
                 End If
                 
                 
                 strCadena = "DELETE FROM movimiento_transferencia WHERE id_transferencia='" & rstZ("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
                 CnBd.Execute (strCadena)
                 
            End If
            
            
        End If
        FrmTransferencias.Procedencia = Neutro
        Call FrmTransferencias.nuevo
        Unload Me
        Exit Sub
    End If
    
    
   
  
    
      
       
       If FrmDerivados.Procedencia = anular Then
          Call anular_derivado(FrmDerivados.TxtCombo.Text)
          FrmDerivados.Procedencia = Neutro
          Unload Me
          Exit Sub
       End If
        
        

        
        If FrmCompras.Procedencia = revertir Then
            
            If get_periodo_cierre(FrmCompras.DtcPeriodo.BoundText, "compras") = True Then
                MsgBox "PERIODO DE COMPRAS CERRARDO.!!!", vbInformation, KEY_VENDEDOR
                Exit Sub
            End If
            
            
            FrmCompras.rever = True
            FrmCompras.cmdModificar.Enabled = False
            FrmCompras.DtcTipoDoc.Enabled = True
            FrmCompras.txtAñoFabricacion.Locked = False
            FrmCompras.txtnumero_dua.Locked = False
            FrmCompras.TxtAnioDua.Locked = False
            FrmCompras.txtañomodelo.Locked = False
            FrmCompras.DtTipoCompra.Locked = False
            
            FrmCompras.txtserie.Enabled = True
            FrmCompras.TxtNumeroDoc.Enabled = True
            FrmCompras.cmdactualizar.Visible = True
            FrmCompras.txtCantidad.Enabled = True
            FrmCompras.TxtDescripcionProducto.Enabled = True
            FrmCompras.cmdRevertir.Visible = False
            FrmCompras.cmdAgregar.Enabled = True
            FrmCompras.CmdQuitar.Enabled = True
            FrmCompras.cmdgastos.Visible = True
            
            
            FrmCompras.lblgastos.Visible = True
            c_serie = FrmCompras.txtserie.Text
            c_numero = FrmCompras.TxtNumeroDoc.Text
            c_doc_cod = FrmCompras.DtcTipoDoc.BoundText
            cPersona = FrmCompras.txtRuc.Text
            
            Dim idCompra As Double
            strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Trim(c_doc_cod) & "' AND serie='" & Trim(c_serie) & "' AND numero='" & Trim(c_numero) & "' AND id_proveedor='" & Trim(cPersona) & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                idCompra = rst("id_compra")
                
                If rst("retencion") > 0 Then
                    FrmCompras.chk_suspencion_retencion.Value = 0
                Else
                    FrmCompras.chk_suspencion_retencion.Value = 1
                End If
                
                FrmCompras.TxtTotalRetencion.Text = Format(rst("retencion"), "#,##00.00")
            Else
                strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Trim(c_doc_cod) & "' AND serie='" & Trim(c_serie) & "' AND numero='" & formato_item(Trim(c_numero), 8) & "' AND id_proveedor='" & Trim(cPersona) & "' AND ruc='" & KEY_RUC & "'"
                Call ConfiguraRst(strCadena)
                If rst.RecordCount > 0 Then
                    idCompra = rst("id_compra")
                    FrmCompras.DtTipoCompra.BoundText = rst("id_tipo_compra")
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
                strCadena = "INSERT INTO movimiento_compra_temporal(id_doc,serie,numero,id_producto,detalle,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento," & _
                "valor_neto,isc,igv,ivap,otros,percepcion,exonerado,valor_venta,precio_venta,p_venta,p_costo,incremento,dni_save,id_alm,retencion,obsequio,ruc) VALUES " & _
                "('" & Trim(c_doc_cod) & "','" & Trim(c_serie) & "','" & Trim(c_numero) & "','" & rst("id_producto") & "','" & rst("detalle") & "','" & rst("cantidad") & "','" & rst("c_unitario") & "'," & _
                "'" & rst("dsto_soles") & "', '" & rst("dsto_procentaje") & "','" & rst("total_descuento") & "','" & rst("valor_neto") & "','" & rst("isc") & "','" & rst("igv") & "'," & _
                "'" & rst("ivap") & "','" & rst("otros") & "','" & rst("percepcion") & "','" & rst("exonerado") & "','" & rst("valor_venta") & "','" & rst("total") & "','" & rst("p_venta") & "'," & _
                "'" & rst("p_costo") & "','" & rst("incremento") & "','" & KEY_USUARIO & "','" & rst("id_alm") & "','" & rst("retencion") & "','" & rst("obsequio") & "','" & KEY_RUC & "')"
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
            FrmCompras.cmdsave.Enabled = True
            'FrmCompras.DtcDocumento.Enabled = True
            'FrmCompras.TxtSerieG.Enabled = True
            'FrmCompras.TxtNumeroG.Enabled = True
            'FrmCompras.DtcDocumento.SetFocus
            Unload Me
            Exit Sub
        End If
       
         
         
         
         If FrmDelivery.Procedencia = modificar Then
            
            strCadena = "UPDATE movimiento_venta SET id_delivery='no' WHERE id_venta='" & Trim(FrmDelivery.HfgLinea.TextMatrix(FrmDelivery.HfgLinea.Row, 0)) & "' AND ruc='" & KEY_RUC & "' "
            CnBd.Execute (strCadena)
             
            
            Call FrmDelivery.LlenarDelivery(FrmDelivery.HfgLinea)
            FrmDelivery.Procedencia = Neutro
            Unload Me
            Exit Sub
        End If
        
        
     
 
        
        
        
        If FrmVentas.Procedencia = anular Then
            FrmVentas.Procedencia = Neutro
            If KEY_CARGO_OPE = KEY_ADMIN Or KEY_CARGO_OPE = KEY_SUPER Or KEY_CARGO = "00001" Or KEY_CARGO = "00004" Or KEY_CARGO = "00006" Or KEY_CARGO = "00008" Then
                
             If get_periodo_cierre_fecha(FrmVentas.DtpActual.Value) = True Then
                 MsgBox "PERIODO CONTABLE.!!!" + Chr(13) + "YA ESTA CERRADO.", vbInformation, KEY_VENDEDOR
                 Exit Sub
             End If
                
                
                Call AnularVentas(Trim(FrmVentas.DtcTipoDoc.BoundText), Trim(FrmVentas.DtcSerieDoc.BoundText), Trim(FrmVentas.TxtNumeroDoc.Text), Trim(FrmVentas.DtcAlmacen.BoundText))
               
                
                If KEY_CONTABILIDAD = "si" Then
                    If KEY_ASIENTO_GLOBAL_CTA_PAGAR = "si" Then
                        strCadena = "call CON_InsertaAsiento_PagoGlobal_Extorno('" & Val(FrmVentas.txtidVenta.Text) & "')"
                        CnBd.Execute (strCadena)
                        
                        strCadena = "SELECT * FROM mis_cuentas_det_detalle mc WHERE mc.id_detalle='" & Val(FrmVentas.txtidVenta.Text) & "' LIMIT 1 "
                        Call ConfiguraRstK(strCadena)
                        If rstK.RecordCount > 0 Then
                            strCadena = "SELECT id_detalle,id_movimiento FROM mis_cuentas_det_detalle WHERE id_movimiento='" & rstK("id_movimiento") & "' and id_tipo='03' and monto_pagado='" & rstK("monto_pagado") & "' LIMIT 1"
                            Call ConfiguraRstK(strCadena)
                            If rstK.RecordCount > 0 Then
                            strCadena = "call CON_InsertaAsiento_Memorandum_Extorno('" & rstK("id_detalle") & "')"
                            CnBd.Execute (strCadena)
                            
                            strCadena = "DELETE FROM mis_cuentas_det_detalle WHERE id_detalle='" & Val(FrmVentas.txtidVenta.Text) & "' and id_movimiento='" & rstK("id_movimiento") & "'"
                            CnBd.Execute (strCadena)
                           
                            
                            End If
                        End If
                    End If
                End If
                
                
                
                FrmVentas.lblAnulado.Visible = True
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
                
               If get_periodo_cierre_fecha(FrmVentas.DtpActual.Value) = True Then
                 MsgBox "PERIODO CONTABLE.!!!" + Chr(13) + "YA ESTA CERRADO.", vbInformation, KEY_VENDEDOR
                 Exit Sub
              End If
            
                
                If get_firma_online(FrmVentas.DtcTipoDoc.BoundText, Trim(FrmVentas.DtcSerieDoc.BoundText)) = "si" Then
                    Call FrmVentas.firma_electronica_eliminar
                    Unload Me
                    'Call EliminarVentas(Trim(FrmVentas.DtcTipoDoc.BoundText), Trim(FrmVentas.DtcSerieDoc.BoundText), Trim(FrmVentas.TxtNumeroDoc.Text), Trim(FrmVentas.DtcAlmacen.BoundText))
                    'FrmVentas.Procedencia = Neutro
                    'Unload Me
                    Exit Sub
                Else
                    Call EliminarVentas(Trim(FrmVentas.DtcTipoDoc.BoundText), Trim(FrmVentas.DtcSerieDoc.BoundText), Trim(FrmVentas.TxtNumeroDoc.Text), Trim(FrmVentas.DtcAlmacen.BoundText))
                    FrmVentas.Procedencia = Neutro
                    Unload Me
                    Exit Sub
                End If
                
                
                
            Else
                MsgBox "PASSWORD DE OPERACIONES INCORECCTO", vbInformation, KEY_EMPRESA
                FrmVentas.Procedencia = Neutro
                Unload Me
                Exit Sub
            End If
        End If
        
        If FrmCompras.Procedencia = 5 Then
            
            If get_periodo_cierre(FrmCompras.DtcPeriodo.BoundText, "compras") = True Then
                MsgBox "PERIODO DE COMPRAS CERRARDO.!!!", vbInformation, KEY_VENDEDOR
                Exit Sub
            End If
            
            Call AnularCompras(Trim(FrmCompras.txtIdCompra.Text))
            FrmCompras.Procedencia = Neutro
            Unload Me
            Exit Sub
        End If
        
        
        If FrmCompras.Procedencia = 3 Then
            FrmCompras.Procedencia = Neutro
             
            If get_periodo_cierre(FrmCompras.DtcPeriodo.BoundText, "compras") = True Then
                MsgBox "PERIODO DE COMPRAS CERRARDO.!!!", vbInformation, KEY_VENDEDOR
                Exit Sub
            End If
     
            Call EliminarCompras(Trim(FrmCompras.DtcTipoDoc.BoundText), Trim(FrmCompras.txtserie.Text), Trim(FrmCompras.TxtNumeroDoc.Text), Trim(FrmCompras.txtRuc.Text))
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
        
        
        
        
        If FrmPrecios.Procedencia = 2 Then
            FrmPrecios.Procedencia = Neutro
            If KEY_CARGO_OPE = "00004" Then
                Call FrmPrecios.LLenaDatos
                Call FrmPrecios.Resize
            Else
                MsgBox "Esta siendo Notificado..." + Chr(13) + "Sera Reportado si lo Intenta Nuevamente" + Chr(13) + Chr(13) + "RANGACHO....", vbInformation, KEY_VENDEDOR
            End If
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
       frmmistareas.Procedencia = Neutro
       If KEY_USUARIO = "42546269" Then
          strCadena = "DELETE FROM proyecto_backlog WHERE id_codigo='" & frmmistareas.HfgLinea.TextMatrix(frmmistareas.HfgLinea.Row, 0) & "'"
          CnBd.Execute (strCadena)
           
          Call frmmistareas.backlog(frmmistareas.HfgLinea)
       End If
       
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
         
        'Call frmmistareas.Criterio(frmmistareas.HfCriterios, frmmistareas.txtidtarea.Text)
        frmmistareas.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
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
Sub Imprimir_motivo(ByVal TipoDoc As String, ByVal Almacen As String, ByVal serie As String, ByVal numero As String)
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
    Printer.Print Tab(0); des_comp; Space(1); Mid(serie + Space(50), 1, 4) & "-" & numero
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.Print Tab(1); "===================================="
    Printer.Print Tab(3); " FECHA/ HORA   :" & Space(3) & str(Date) & Space(3) & Trim(Time)
    Printer.Print Tab(3); " MOTIVO DE ANULACION:"
    Printer.Print Tab(3); Trim(Me.txtMotivo.Text)
    Printer.Print Tab(1); "===================================="
    Printer.EndDoc
End Sub
Sub AnularCompras(ByVal in_compra As String)
On Error GoTo salir

strCadena = "UPDATE movimiento_compra SET anulado='si' WHERE id_compra='" & Val(in_compra) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


strCadena = "SELECT * FROM movimiento_compra_detalle  WHERE id_compra='" & Val(in_compra) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   
   strCadena = "UPDATE imp_producto_detalle SET ruc='0' WHERE id_compra='" & Val(in_compra) & "'  and ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
   
   For i = 0 To rst.RecordCount - 1
        
        strCadena = "DELETE FROM kardex where id_producto='" & rst("id_producto") & "' and cantidad_real>0 and  id_movimiento='" & rst("id_compra") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        
        strCadena = "call put_actualizar_stock_producto('" & rst("id_producto") & "','" & rst("id_alm") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)

        rst.MoveNext
        
   Next i
    
End If











FrmCompras.lblAnulado.Visible = True
FrmCompras.cmdAnular.Enabled = False

Exit Sub
salir:
MsgBox "Ocurrio un Error en la Conexion", vbInformation



End Sub

Sub AnularTransferencia(ByVal TipoDoc As String, ByVal serie As String, ByVal numero As String, ByVal AlmacenOrigen As String, ByVal AlmacenDestino As String)
Dim Anulado As String * 1
Anulado = "V"
'---------------
    Dim RstDetTransferencia As New ADODB.Recordset
    strCadena = "SELECT cProducto,cantidad FROM Detalle_DocumentoTransferencia WHERE (doc_cod='" & TipoDoc & "'AND sSerie ='" & serie & "'AND cDocumentoTransferencia='" & numero & "' AND Alm_Origen='" & AlmacenOrigen & "' AND Alm_Destino='" & AlmacenDestino & "' ) ORDER BY cProducto ASC"
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
strCadena = "UPDATE DocumentoTransferencia SET Anulado='" & Anulado & "' WHERE (cDocumentoTransferencia='" & numero & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "'AND Alm_Origen='" & AlmacenOrigen & "' AND Alm_Destino='" & AlmacenDestino & "')"
Call EjecutaRST(strCadena)
Set rst = Nothing
strCadena = "DELETE Kardex  WHERE (NumeroDoc='" & numero & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "' AND Alm_cod='" & Almacen & "')"
Call EjecutaRST(strCadena)
Set rst = Nothing
Unload Me
'FrmTransferencias.lblAnulado.Visible = True
'FrmTransferencias.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
End Sub
Sub EliminarTransferencia(ByVal TipoDoc As String, ByVal serie As String, ByVal numero As String, ByVal AlmacenOrigen As String, ByVal AlmacenDestino As String)
'---------------------------
    Dim RstDetTransferencia As New ADODB.Recordset
    strCadena = "SELECT cProducto,cantidad FROM Detalle_DocumentoTransferencia WHERE (doc_cod='" & TipoDoc & "'AND sSerie ='" & serie & "'AND cDocumentoTransferencia='" & numero & "' AND Alm_Origen='" & AlmacenOrigen & "' AND Alm_Destino='" & AlmacenDestino & "') ORDER BY cProducto ASC"
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
strCadena = "DELETE DocumentoTransferencia  WHERE (cDocumentoTransferencia='" & numero & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "' AND Alm_Origen='" & AlmacenOrigen & "' AND Alm_Destino='" & AlmacenDestino & "')"
Call EjecutaRST(strCadena)
Set rst = Nothing
strCadena = "DELETE Kardex  WHERE (NumeroDoc='" & numero & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "' AND Alm_cod='" & Almacen & "')"
Call EjecutaRST(strCadena)
Set rst = Nothing
Unload Me
'FrmTransferencias.Nuevo
'FrmTransferencias.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
End Sub

Sub EliminarAdelanto(ByVal TipoDoc As String, ByVal serie As String, ByVal numero As String, ByVal Persona As String)
                 
strCadena = "DELETE DocumentoVenta  WHERE (cDocumentoVenta='" & numero & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "' AND cpersona='" & Persona & "')"
Call EjecutaRST(strCadena)
Set rst = Nothing
Unload Me
FrmAdelantoPersonal.nuevo
'FrmAdelantoPersonal.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
End Sub
Sub EliminarIngresoDinero(ByVal TipoDoc As String, ByVal serie As String, ByVal numero As String, ByVal Persona As String)
              
strCadena = "DELETE DocumentoVenta  WHERE (cDocumentoVenta='" & numero & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "' AND cpersona='" & Persona & "')"
Call EjecutaRST(strCadena)
Set rst = Nothing
Unload Me
FrmIngresoDinero.nuevo
FrmIngresoDinero.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
End Sub

Public Sub EliminarCompras(ByVal id_doc As String, ByVal serie As String, ByVal numero As String, ByVal id_proveedor As String)
Dim idCompra As Double

strCadena = "SELECT * FROM movimiento_compra WHERE numero='" & numero & "' AND serie='" & serie & "' AND id_doc='" & id_doc & "' AND id_proveedor='" & id_proveedor & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    idCompra = rst("id_compra")
    strCadena = "SELECT * FROM `imp_producto_detalle` WHERE id_compra='" & idCompra & "' and  id_alm='" & KEY_ALM & "' and  (vendido='si' or transferencia='si') and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
        MsgBox "SERIES VENDIDAS O TRANSFERIDAS" + Chr(13) + "IMPOSIBLE ELIMINAR", vbInformation, KEY_VENDEDOR
        Exit Sub
    End If
    
     strCadena = "DELETE FROM movimiento_compra WHERE id_compra='" & idCompra & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
     CnBd.Execute (strCadena)
     
     strCadena = "Call CON_Asiento_EliminarCompra('" & idCompra & "', '" & KEY_USUARIO & "', '" & KEY_RUC & "')"
     CnBd.Execute (strCadena)
     
   
    
    
    
    strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & idCompra & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraTemporal(strCadena)
    If rstTemporal.RecordCount > 0 Then
        rstTemporal.MoveFirst
        FrmCompras.progresbar_kardex.Min = 0
        FrmCompras.progresbar_kardex.Max = rstTemporal.RecordCount
        MsgBox "SE VA A PROCEDER A ACTUALIZAR KARDEX" + Chr(13) + Chr(13) + "PULSE ACEPTAR Y DEJE QUE TERMINE DE ACTUALIZAR.", vbInformation
        For i = 0 To rstTemporal.RecordCount - 1
            
            
            
            strCadena = "DELETE FROM movimiento_compra_detalle WHERE id_compra='" & idCompra & "' AND id_detalle_compra='" & rstTemporal("id_detalle_compra") & "' "
            CnBd.Execute (strCadena)
            
            strCadena = "DELETE FROM kardex WHERE  id_producto='" & rstTemporal("id_producto") & "' and  id_movimiento='" & idCompra & "' and id_doc='" & id_doc & "' and id_serie='" & serie & "' and id_numero='" & numero & "' and id_persona='" & id_proveedor & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            CnBd.Execute (strCadena)
            
            strCadena = "DELETE FROM imp_producto_detalle WHERE id_detalle_compra='" & rstTemporal("id_detalle_compra") & "' and id_producto='" & rstTemporal("id_producto") & "' and  id_compra='" & idCompra & "'  and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
            
            
            If FrmCompras.DtcTipo.BoundText = "01" Then ' material
        
                If KEY_RUC = "20128836251" Then
                    Call update_kardex_Vargas_modulo_compra(rstTemporal("id_producto"), Format(FrmCompras.DtpKardex.Value, "YYYY-mm-dd"))
                Else
                    Call update_kardex_update(rstTemporal("id_producto"), Format(FrmCompras.DtpKardex.Value, "YYYY-mm-dd"))
                End If
            End If
           
            FrmCompras.progresbar_kardex.Value = i
       
            
            
            
            DoEvents
            rstTemporal.MoveNext
            DoEvents
        Next i
 MsgBox "Proceso Actualizacion Kardex Correcto.", vbInformation
    End If
     
     
     
     
     
                
                
    
     
     
     
     
    
                
    
            '---
    
End If

FrmCompras.nuevo
'FrmCompras.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
FrmCompras.cmdEliminar.Enabled = False
Unload Me
End Sub

Sub EliminarGuia(ByVal serie As String, ByVal numero As String)
strCadena = "DELETE Detalleguia  WHERE (sNumeroGuia='" & numero & "' AND sSerieGuia='" & serie & "')"
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

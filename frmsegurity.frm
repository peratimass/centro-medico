VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsegurity 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMotivo 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox TxtClave 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   1830
   End
   Begin VB.Timer timer_camara 
      Interval        =   1000
      Left            =   1680
      Top             =   1080
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   2160
      TabIndex        =   2
      Top             =   1100
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image img_of 
      Height          =   1530
      Left            =   0
      Picture         =   "frmsegurity.frx":0000
      Top             =   -120
      Width           =   1785
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   1785
      Picture         =   "frmsegurity.frx":50A7
      Top             =   0
      Width           =   2715
   End
   Begin VB.Image img_on 
      Height          =   1530
      Left            =   0
      Picture         =   "frmsegurity.frx":9CC0
      Top             =   -120
      Visible         =   0   'False
      Width           =   1785
   End
End
Attribute VB_Name = "frmsegurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Call cerrar
End If
End Sub
Public Sub cerrar()
If FrmAlmacenes.Procedencia = Eliminar Then
    FrmAlmacenes.Procedencia = Neutro
    Call enabled_form(FrmAlmacenes)
    Unload Me
    Exit Sub
End If

If frmMemorandun.Procedencia = anular Then
   frmMemorandun.Procedencia = Neutro
   Call enabled_form(frmMemorandun)
   Unload Me
   Exit Sub
End If


If FrmDetalleventa.Procedencia = modificar Then
   FrmDetalleventa.Procedencia = Neutro
   Call enabled_form(FrmDetalleventa)
   Unload Me
   Exit Sub
End If

If FrmDetallePersona.Procedencia = Eliminar Then
   FrmDetallePersona.Procedencia = Neutro
   Call enabled_form(FrmPersona)
   Call enabled_form(FrmDetallePersona)
   Unload Me
   Exit Sub
End If

If FrmProducto.Procedencia = Eliminar Then
   FrmProducto.Procedencia = Neutro
   Call enabled_form(FrmProducto)
   Unload Me
   Exit Sub
End If

If FrmSolicitudViaticos.Procedencia = anular Then
    FrmSolicitudViaticos.Procedencia = Neutro
    Call enabled_form(FrmSolicitudViaticos)
    Unload Me
    Exit Sub
End If

If frmmisproyectos.Procedencia = Eliminar Then
   frmmisproyectos.Procedencia = Neutro
   Call enabled_form(frmmisproyectos)
   Unload Me
   Exit Sub
End If

If FrmVentas.Procedencia = modificar_precio Then
    
   FrmVentas.Procedencia = Neutro
   Call enabled_form(FrmVentas)
   Unload Me
   Exit Sub
End If




If FrmVentaCantidad.Procedencia = modificar_precio Then
   FrmVentaCantidad.Procedencia = Neutro
   Call enabled_form(FrmVentas)
   Call enabled_form(FrmVentaCantidad)
   Unload Me
   Exit Sub
End If

If FrmVentaCantidad.Procedencia = modificar_precio_unitario Then
   FrmVentaCantidad.Procedencia = Neutro
   FrmVentaCantidad.txtprecio.Text = Format(Val(FrmVentaCantidad.lblprecio_original.Caption), "###0.00")
   Call enabled_form(FrmVentas)
   Call enabled_form(FrmVentaCantidad)
   Unload Me
   Exit Sub
End If

If FrmVentas.Procedencia = diferida Then
   FrmVentas.Procedencia = Neutro
   FrmVentas.chk_venta_diferida.Value = 0
   Call enabled_form(FrmVentas)
   Unload Me
    Exit Sub
End If


If frmCajaEgreso.Procedencia = extornar Then
   frmCajaEgreso.Procedencia = Neutro
   Call enabled_form(frmCajaEgreso)
   Unload Me
   Exit Sub
End If

If frmCajaEgreso.Procedencia = modificar Then
   frmCajaEgreso.Procedencia = Neutro
   'Call frmCajaEgreso.get_movimiento(Val(frmCajaEgreso.HfDetalle.TextMatrix(frmCajaEgreso.HfDetalle.Row, 0)))
   'frmCajaEgreso.cmdProcesar.Enabled = True
   Call enabled_form(frmCajaEgreso)
   Call enabled_form(FrmMiscuentas)
   Unload Me
   Exit Sub
End If


If FrmCompras.Procedencia = modificar_precio Then
   FrmCompras.Procedencia = Neutro
   Call enabled_form(FrmCompras)
   Unload Me
   Exit Sub
End If


If FrmOrdenCompra.Procedencia = anular Then
   FrmOrdenCompra.Procedencia = Neutro
   Call enabled_form(FrmOrdenCompra)
   Unload Me
   Exit Sub
End If

If FrmPedido.Procedencia = anular Then
   FrmPedido.Procedencia = Neutro
   Call enabled_form(FrmPedido)
   Unload Me
   Exit Sub
End If


If FrmCambioAceite.Procedencia = anular Then
    FrmCambioAceite.Procedencia = Neutro
    Call enabled_form(FrmCambioAceite)
    Unload Me
    Exit Sub
End If


If frmmanifiesto.Procedencia = anular Then
   frmmanifiesto.Procedencia = Neutro
   Call enabled_form(frmmanifiesto)
   Unload Me
   Exit Sub
End If

If frmPlanesServicio.Procedencia = Eliminar Then
   frmPlanesServicio.Procedencia = Neutro
   Call enabled_form(frmPlanesServicio)
   Unload Me
   Exit Sub
End If

If FrmSolicitudViaticosDeclarar.Procedencia = Eliminar Then
   FrmSolicitudViaticosDeclarar.Procedencia = Neutro
   Call enabled_form(FrmSolicitudViaticosDeclarar)
   Unload Me
   Exit Sub
End If


If frmHotelInfraestructura.Procedencia = Eliminar Then
   frmHotelInfraestructura.Procedencia = Neutro
   Call enabled_form(frmHotelInfraestructura)
   Unload Me
   Exit Sub
End If


If frmHotelInfraestructura.Procedencia = eliminar_informe Then
   frmHotelInfraestructura.Procedencia = Neutro
   Call enabled_form(frmHotelInfraestructura)
   Unload Me
   Exit Sub
End If



If MDIFrmPrincipal.Procedencia = seleccionar_per Then
   MDIFrmPrincipal.Procedencia = Neutro
    Unload Me
    Exit Sub
End If


If FrmModelo.Procedencia = Eliminar Then
    Call enabled_form(FrmModelo)
    Unload Me
    Exit Sub
End If



Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 3000
End Sub

Private Sub TxtClave_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim in_pass_admin As String
    
    
    If FrmVentas.Procedencia = modificar_precio Then
       FrmVentas.Procedencia = Neutro
       in_pass_admin = verificar_password_admin(Me.TxtClave.Text)
       If in_pass_admin = "00004" Or in_pass_admin = "00003" Or in_pass_admin = "00009" Or KEY_USUARIO = "74915823" Then
          
          If FrmVentas.control_stock(Trim(FrmVentas.TxtCodProducto.Text), Val(FrmVentas.txtCantidad.Text)) = True Then
                Call FrmVentas.insertar_item
          End If
    
         
       End If
       
       Call enabled_form(FrmVentas)
       Unload Me
       Exit Sub
    End If
    
    
    
      
      If FrmDetalleventa.Procedencia = modificar Then
         FrmDetalleventa.Procedencia = Neutro
         in_pass_admin = verificar_password_admin(Me.TxtClave.Text)
          If in_pass_admin = "00004" Then
            FrmDetalleventa.lblid_registro.Caption = FrmDetalleventa.HfPagos.TextMatrix(FrmDetalleventa.HfPagos.Row, 0)
            FrmDetalleventa.put_formapago (FrmDetalleventa.lblid_registro.Caption)
         End If
         Call enabled_form(FrmDetalleventa)
         Unload Me
         Exit Sub
      End If
      
      
      If FrmDetalleventa.Procedencia = modificar_credito Then
         FrmDetalleventa.Procedencia = Neutro
         in_pass_admin = verificar_password_admin(Me.TxtClave.Text)
          If in_pass_admin = "00004" Then
            FrmDetalleventa.cmdprocesarVendedor.Visible = True
            
          Else
            FrmDetalleventa.cmdprocesarVendedor.Visible = False
         End If
         Call enabled_form(FrmDetalleventa)
         Unload Me
         Exit Sub
      End If
      
      
      If FrmVentaCantidad.Procedencia = modificar_precio Then
       FrmVentaCantidad.Procedencia = Neutro
       in_pass_admin = verificar_password_admin(Me.TxtClave.Text)
       If in_pass_admin = "00004" Or in_pass_admin = "00003" Or in_pass_admin = "00009" Or KEY_USUARIO = "74915823" Then
          FrmVentaCantidad.update_precio
       Else
          Call enabled_form(FrmVentaCantidad)
          Call Resalta(FrmVentaCantidad.TxtTotal)
       End If
       'FrmVentaCantidad.txtPrecio.Text = Format(Val(FrmVentaCantidad.lblprecio_original.Caption), "###0.00")
       Call enabled_form(FrmVentas)
       Unload Me
       Exit Sub
    End If
    
    
    
    
    If frmagenda.Procedencia = anular Then
        frmagenda.Procedencia = Neutro
        On Error GoTo salir_agenda
        Call frmagenda.put_anular(frmagenda.HfdDetalle.TextMatrix(frmagenda.HfdDetalle.Row, 0))
        Call enabled_form(frmagenda)
salir_agenda:
        Unload Me
        Exit Sub
    End If
    
    
    
    If FrmVentaCantidad.Procedencia = modificar_precio_unitario Then
       FrmVentaCantidad.Procedencia = Neutro
       in_pass_admin = verificar_password_admin(Me.TxtClave.Text)
       If in_pass_admin = "00004" Or in_pass_admin = "00003" Or in_pass_admin = "00009" Or KEY_USUARIO = "74915823" Then
          FrmVentaCantidad.update_precio_unitario
       Else
          FrmVentaCantidad.txtprecio.Text = Format(Val(FrmVentaCantidad.lblprecio_original.Caption), "###0.00")
          Call enabled_form(FrmVentaCantidad)
          Call Resalta(FrmVentaCantidad.TxtTotal)
       End If
       Call enabled_form(FrmVentas)
       Unload Me
       Exit Sub
    End If
    
    
    
   
    
    If FrmVentas.Procedencia = diferida Then
         FrmVentas.Procedencia = Neutro
         in_pass_admin = verificar_password_admin(Me.TxtClave.Text)
            If in_pass_admin = "00057" Then
                
                
            Else
                 FrmVentas.chk_venta_diferida.Value = 0
                 MsgBox "Usted No esta Autorizado para ESTA FUNCION", vbInformation, KEY_VENDEDOR
            End If
            
            Call enabled_form(FrmVentas)
            Unload Me
            Exit Sub
        End If
        
        
        
    If frmPlanesServicio.Procedencia = Eliminar Then
       frmPlanesServicio.Procedencia = Neutro
       in_pass_admin = verificar_password_admin(Me.TxtClave.Text)
       If in_pass_admin = "00004" Or in_pass_admin = "00003" Or in_pass_admin = "00009" Then
          Call put_eliminar_plan(frmPlanesServicio.HfgLinea.TextMatrix(frmPlanesServicio.HfgLinea.Row, 0))
          Call frmPlanesServicio.actualizar(frmPlanesServicio.HfgLinea)
       Else
         If in_pass_admin = "0" Then
            MsgBox "PASSWORD DE OPERACIONES INCORRECTA" + Chr(13) + "INTENTE NUEVAMENTE", vbInformation
         Else
            MsgBox "USTED NO CUENTA CON LOS PERMISOS" + Chr(13) + "PARA ELIMINAR EL REGISTRO", vbInformation
         End If
         
         
       End If
       Call enabled_form(frmPlanesServicio)
       Unload Me
       Exit Sub
    End If
        
    
    
    
    If verificar_password(Trim(Me.TxtClave.Text)) = True Then
       
       If FrmTransferencias.Procedencia = anular Then
          FrmTransferencias.Procedencia = Neutro
          Call anular_guia(FrmTransferencias.DtcTipoDoc.BoundText, FrmTransferencias.DtcSerieGuia.Text, FrmTransferencias.TxtNumeroDoc.Text, FrmTransferencias.DtcAlmacenOrigen.BoundText)
          FrmTransferencias.frmanulado.Visible = True
          FrmTransferencias.lblAnulado.Visible = True
          'FrmTransferencias.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
          FrmTransferencias.cmdProcesar.Enabled = False
          Unload Me
          Exit Sub
       End If
       
        If FrmSolicitudViaticos.Procedencia = Eliminar Then
           FrmSolicitudViaticos.Procedencia = Neutro
           If anular_solicitud(FrmSolicitudViaticos.HfgDetalle.TextMatrix(FrmSolicitudViaticos.HfgDetalle.Row, 0)) = True Then
              Call FrmSolicitudViaticos.actualizar
           End If
           Call enabled_form(FrmSolicitudViaticos)
           
           Unload Me
           Exit Sub
        End If
    
        
        
        
        
        
        If frmHotelInfraestructura.Procedencia = Eliminar Then
           frmHotelInfraestructura.Procedencia = Neutro
           Call enabled_form(frmHotelInfraestructura)
           Call frmHotelInfraestructura.put_delete_piso
           Unload Me
           Exit Sub
        End If
        
        
           
        If frmHotelInfraestructura.Procedencia = eliminar_informe Then
           frmHotelInfraestructura.Procedencia = Neutro
           Call enabled_form(frmHotelInfraestructura)
           Call frmHotelInfraestructura.put_delete_habitacion
           Unload Me
           Exit Sub
        End If
        
        
        
        If MDIFrmPrincipal.Procedencia = seleccionar_per Then
           MDIFrmPrincipal.Procedencia = Neutro
           FrmParametrosEmpresa.Show
           Unload Me
           Exit Sub
           
        End If
        
        If FrmFormaPago.Procedencia = Eliminar Then
           FrmFormaPago.Procedencia = Neutro
           Call delete_formapago(FrmFormaPago.HfdPersona.TextMatrix(FrmFormaPago.HfdPersona.Row, 0))
           FrmFormaPago.actualizar
           Call enabled_form(FrmFormaPago)
           Unload Me
           Exit Sub
        End If
        
        
        If FrmCambioAceite.Procedencia = anular Then
            FrmCambioAceite.Procedencia = Neutro
            If put_anular_cambio(FrmCambioAceite.hfmensualidad.TextMatrix(FrmCambioAceite.hfmensualidad.Row, 0)) = True Then
                For k = 0 To 7
                                FrmCambioAceite.hfmensualidad.col = k
                                FrmCambioAceite.hfmensualidad.Row = FrmCambioAceite.hfmensualidad.Row
                                FrmCambioAceite.hfmensualidad.CellBackColor = &H8080FF
                Next k
            End If
            Call enabled_form(FrmCambioAceite)
            Unload Me
            Exit Sub
        End If
        
        If frmCajaEgreso.Procedencia = modificar Then
            frmCajaEgreso.Procedencia = Neutro
            Call frmCajaEgreso.get_movimiento(Val(frmCajaEgreso.HfDetalle.TextMatrix(frmCajaEgreso.HfDetalle.Row, 0)))
            frmCajaEgreso.cmdEliminarPago.Enabled = True
            frmCajaEgreso.cmdModificar_pago.Enabled = True
            frmCajaEgreso.cmdProcesar.Enabled = True
            Unload Me
            Exit Sub
        End If

        If FrmDetallePersona.Procedencia = Eliminar Then
           FrmDetallePersona.Procedencia = Neutro
           Call put_delete_servicio(Val(FrmDetallePersona.Hfplanservicio.TextMatrix(FrmDetallePersona.Hfplanservicio.Row, 0)))
           Call enabled_form(FrmPersona)
           Call enabled_form(FrmDetallePersona)
           Call FrmDetallePersona.llenar_plan_servicio(FrmDetallePersona.Hfplanservicio, FrmDetallePersona.txtRuc.Text)
           Unload Me
           Exit Sub
        End If
       
       If frmMemorandun.Procedencia = anular Then
          frmMemorandun.Procedencia = Neutro
          Call anular_memorandum(frmMemorandun.HfMemorandum.TextMatrix(frmMemorandun.HfMemorandum.Row, 0))
          Call enabled_form(frmMemorandun)
          Unload Me
          Exit Sub
       End If
       
       
       If FrmOrdenCompra.Procedencia = anular Then
            FrmOrdenCompra.Procedencia = Neutro
            Call anular_orden_compra(FrmOrdenCompra.HfgDetalle.TextMatrix(FrmOrdenCompra.HfgDetalle.Row, 0))
            Call enabled_form(FrmOrdenCompra)
            Call FrmOrdenCompra.actualizar
            Unload Me
            Exit Sub
       End If
       
       
       If FrmPedido.Procedencia = anular Then
          FrmPedido.Procedencia = Neutro
          Call anular_pedido(FrmPedido.HfgDetalle.TextMatrix(FrmPedido.HfgDetalle.Row, 0))
          Call enabled_form(FrmPedido)
          Unload Me
          Exit Sub
       End If
       
       
       If FrmCompras.Procedencia = modificar_precio Then
          FrmCompras.Procedencia = Neutro
          
          
          If KEY_CARGO <> "00004" Then
             FrmCompras.TxtventaHoy.Text = Format(FrmCompras.txtPrecioVentaAnt.Text, "###0.00")
             MsgBox "Usted No tiene el Permiso para Hacer la Modificacion de Precios", vbInformation
          End If
          
          Call enabled_form(FrmCompras)
          Call FrmCompras.AgregarNuevo
          Unload Me
          Exit Sub
       End If
       
       
       
       
       If frmmanifiesto.Procedencia = anular Then
          frmmanifiesto.Procedencia = Neutro
          
          
          Call enabled_form(frmmanifiesto)
          Unload Me
          Exit Sub
       End If
       
       
       
       
       If frmCajaEgreso.Procedencia = extornar Then
          frmCajaEgreso.Procedencia = Neutro
          
          'strCadena = "SELECT id_venta,id_compra FROM mis_cuentas_det WHERE id='" & Val(frmCajaEgreso.HfDetalle.TextMatrix(frmCajaEgreso.HfDetalle.Row, 0)) & "' "
          'Call ConfiguraRstP(strCadena)
          'If rstP.RecordCount > 0 Then
          '  If rstP("id_venta") > 0 Then
          '      strCadena = "SELECT * FROM con_documento WHERE idreferencia ='" & rstP("id_venta") & "'"
          '      Call ConfiguraRstPP(strCadena)
          '      If rstPP.RecordCount > 0 Then
          '          strCadena = "call sp_extornar_transaccion_caja('" & Val(frmCajaEgreso.HfDetalle.TextMatrix(frmCajaEgreso.HfDetalle.Row, 0)) & "')"
          '          CnBd.Execute (strCadena)
          '      End If
          '  End If
            
          '   If rstP("id_compra") > 0 Then
          '      strCadena = "SELECT * FROM con_documento WHERE idreferencia ='" & rstP("id_compra") & "'"
          '      Call ConfiguraRstPP(strCadena)
          '      If rstPP.RecordCount > 0 Then
          '          strCadena = "call sp_extornar_transaccion_caja('" & Val(frmCajaEgreso.HfDetalle.TextMatrix(frmCajaEgreso.HfDetalle.Row, 0)) & "')"
          '          CnBd.Execute (strCadena)
          '      End If
          '  End If
            
           
          'End If
          
          
          
          strCadena = "call ADM_revertir_CajaBanco_v2('" & Val(frmCajaEgreso.HfDetalle.TextMatrix(frmCajaEgreso.HfDetalle.Row, 0)) & "','" & KEY_RUC & "')"
          CnBd.Execute (strCadena)
          
         ' strCadena = "call ADM_revertir_pago('" & Val(frmCajaEgreso.HfDetalle.TextMatrix(frmCajaEgreso.HfDetalle.Row, 0)) & "','" & KEY_RUC & "')"
         ' CnBd.Execute (strCadena)
          
          
          Call put_revertir_ultimate(Val(frmCajaEgreso.HfDetalle.TextMatrix(frmCajaEgreso.HfDetalle.Row, 0)))
          
          MsgBox "Extorno Correcto", vbInformation, KEY_VENDEDOR
          Call enabled_form(frmCajaEgreso)
           Unload Me
          
          Call frmCajaEgreso.recientes(Val(frmCajaEgreso.TxtidCuenta.Text), frmCajaEgreso.DtpInicio.Value, frmCajaEgreso.DtpFin.Value, get_moneda_cuenta(frmCajaEgreso.TxtidCuenta.Text))
         
          Exit Sub
       End If
       
       
       
       If frmmistareas.Procedencia = eliminar_informe Then
          frmmistareas.Procedencia = Neutro
          If eliminar_informe_I(frmmistareas.HfInforme01.TextMatrix(frmmistareas.HfInforme01.Row, 0)) = True Then
            Call frmmistareas.llenar_informe(frmmistareas.HfInforme01, frmmistareas.MonthInforme.Value, KEY_USUARIO)
          End If
          Unload Me
          Exit Sub
       End If
       
       
       If FrmSolicitudViaticos.Procedencia = anular Then
          FrmSolicitudViaticos.Procedencia = Neutro
          If anular_solicitud(Val(FrmSolicitudViaticos.HfgDetalle.TextMatrix(FrmSolicitudViaticos.HfgDetalle.Row, 0))) = True Then
             Call FrmSolicitudViaticos.actualizar
          End If
          Call enabled_form(FrmSolicitudViaticos)
          Unload Me
          Exit Sub
       End If
       
       If frmmisproyectos.Procedencia = Eliminar Then
        frmmisproyectos.Procedencia = Neutro
        If eliminar_proyecto(frmmisproyectos.HfdPersona.TextMatrix(frmmisproyectos.HfdPersona.Row, 0)) = True Then
           Call frmmisproyectos.llenarGrid(frmmisproyectos.HfdPersona)
        End If
        Unload Me
        Exit Sub
       End If
       
       If FrmAlmacenes.Procedencia = Eliminar Then
          FrmAlmacenes.Procedencia = Neutro
          Call delete_almacen(FrmAlmacenes.HfgAlmacen.TextMatrix(FrmAlmacenes.HfgAlmacen.Row, 0))
          FrmAlmacenes.Actualizar_Alm
          Call enabled_form(FrmAlmacenes)
          Unload Me
          Exit Sub
        End If
       
       
       If FrmSeguros.Procedencia = Eliminar Then
          FrmSeguros.Procedencia = Neutro
          Call delete_seguro(FrmSeguros.HfSeguros.TextMatrix(FrmSeguros.HfSeguros.Row, 0))
          Call FrmSeguros.actualizar
          Call enabled_form(FrmSeguros)
          
          Unload Me
          
          Exit Sub
       End If
       
       If FrmVentaCantidad.Procedencia = modificar_precio Then
          FrmVentaCantidad.Procedencia = Neutro
          Call enabled_form(FrmVentas)
          Call enabled_form(FrmVentaCantidad)
          Call FrmVentaCantidad.update_precio
          Unload Me
          Exit Sub
       End If
       
       
       
       
       If FrmSolicitudViaticosDeclarar.Procedencia = Eliminar Then
          FrmSolicitudViaticosDeclarar.Procedencia = Neutro
          Call eliminar_gasto_viatico(FrmSolicitudViaticosDeclarar.HfGastos.TextMatrix(FrmSolicitudViaticosDeclarar.HfGastos.Row, 0), FrmSolicitudViaticosDeclarar.HfGastos.TextMatrix(FrmSolicitudViaticosDeclarar.HfGastos.Row, 1))
          Call enabled_form(FrmSolicitudViaticosDeclarar)
          Unload Me
          Exit Sub
       End If
       
       
       If FrmProducto.Procedencia = Eliminar Then
          FrmProducto.Procedencia = Neutro
          Call delete_producto(FrmProducto.HfdGrilla.TextMatrix(FrmProducto.HfdGrilla.Row, 0))
          Call FrmProducto.ActualizarProd
          Call enabled_form(FrmProducto)
          Unload Me
          Exit Sub
       End If
       
       
       If FrmModelo.Procedencia = Eliminar Then
          FrmModelo.Procedencia = Neutro
          strCadena = "call ADM_Sublinea_Modelo('E','" & Val(FrmModelo.HfModelo.TextMatrix(FrmModelo.HfModelo.Row, 0)) & "','','','','" & KEY_RUC & "')"
          Call ConfiguraRst(strCadena)
          Call enabled_form(FrmModelo)
          Call FrmModelo.actualizar(FrmModelo.HfModelo)
          Unload Me
          Exit Sub
       End If
       
       
    
    
    Else
        
        MsgBox "PASSWORD DE ACCIONES INCORRECTA." + Chr(13) + Chr(13) + "INTENTE NUEVAMENTE.", vbInformation, "SR(A)." & KEY_VENDEDOR
        Unload Me
        
        If FrmVentas.Procedencia = seleccionar_vendedor Then
           FrmVentas.Procedencia = Neutro
           Call enabled_form(FrmVentas)
        End If
        
        If FrmVentas.Procedencia = diferida Then
           FrmVentas.Procedencia = Neutro
           Call enabled_form(FrmVentas)
        End If
        
        Exit Sub
    End If
End If
End Sub

Private Sub eliminar_gasto_viatico(ByVal in_detalle As String, ByVal in_compra As String)
Call Eliminar_gasto_viaticos(in_detalle, in_compra)
Call FrmSolicitudViaticosDeclarar.llenar_gastos(FrmSolicitudViaticosDeclarar.HfGastos, FrmSolicitudViaticosDeclarar.Txtid_solicitud.Text)

End Sub
Private Sub TxtMotivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtClave)
    Exit Sub
End If
End Sub

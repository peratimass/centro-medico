VERSION 5.00
Begin VB.Form FrmVentaCantidad 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modificar Cantidad / Precio"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txttotal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2640
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtPrecio 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox TxtCantidad 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lbl_obsequio 
      Height          =   135
      Left            =   3480
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblprecio_original 
      Height          =   135
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbltotal 
      Caption         =   "0"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "FrmVentaCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub Form_Activate()
Call Resalta(Me.txtCantidad)
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 4000
If FrmVentas.DtcTipoDoc.BoundText = "0099" And KEY_UPDATE_PROFORM = "si" And Val(FrmVentas.txtidVenta.Text) > 0 And FrmVentas.cmdModificar.Enabled = False Then
    strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_detalle_venta='" & Val(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT * FROM temporal_ventas WHERE id='" & Val(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0)) & "'"
End If

Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtCantidad.Text = rst("cantidad")
   Me.txtprecio.Text = Format(rst("precio"), "###0.00")
   Me.TxtTotal.Text = Format(rst("total"), "###0.00")
   Me.lblTotal.Caption = Val(Me.TxtTotal.Text)
End If




End Sub
Public Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub
Private Sub put_interes_nota()
Dim in_interes As Single



End Sub


Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
Dim cantidad As Single
Dim precio As Single
Dim Total As Double
If KeyAscii = 13 Then
  If FrmVentas.DtcTipoDoc.BoundText = "0007" Then
    
    
     If Me.lbl_obsequio.Caption = "si" Then
        strCadena = "UPDATE temporal_ventas SET cantidad='" & Val(Me.txtCantidad.Text) & "',precio='" & Val(Me.txtprecio.Text) & "',total='0' WHERE id='" & Val(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0)) & "' AND dni_save='" & KEY_USUARIO & "'"
     Else
        strCadena = "UPDATE temporal_ventas SET cantidad='" & Val(Me.txtCantidad.Text) & "',precio='" & Val(Me.txtprecio.Text) & "',total='" & Val(Me.txtCantidad.Text) * Val(Me.txtprecio.Text) & "' WHERE id='" & Val(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0)) & "' AND dni_save='" & KEY_USUARIO & "'"
     End If
     CnBd.Execute (strCadena)
    
     'Updateo de Interes
      strCadena = "CALL put_interes_nota_credito('" & Val(FrmVentas.txtid_venta_ref.Text) & "','" & KEY_USUARIO & "','" & KEY_PRODUCTO_INTERES & "')"
      CnBd.Execute (strCadena)
      
      strCadena = "SELECT ifnull(sum(total),0) FROM  temporal_ventas  WHERE id_producto='" & KEY_PRODUCTO_INTERES & "' and  id_alm='" & KEY_ALM & "' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
      Call ConfiguraRstlocal(strCadena)
      
      
    
    Call FrmVentas.llenarGrid_det(FrmVentas.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText, FrmVentas.txtformato_impresion.Text)
    Call FrmVentas.Resalta(FrmVentas.TxtCodProducto)
    Unload Me
    Exit Sub
        
  Else
   
     If Me.lbl_obsequio.Caption = "si" Then
        strCadena = "UPDATE temporal_ventas SET cantidad='" & Val(Me.txtCantidad.Text) & "',precio='" & Val(Me.txtprecio.Text) & "',total='0' WHERE id='" & Val(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0)) & "' AND dni_save='" & KEY_USUARIO & "'"
     Else
        strCadena = "UPDATE temporal_ventas SET cantidad='" & Val(Me.txtCantidad.Text) & "',precio='" & Val(Me.txtprecio.Text) & "',total='" & Val(Me.txtCantidad.Text) * Val(Me.txtprecio.Text) & "' WHERE id='" & Val(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0)) & "' AND dni_save='" & KEY_USUARIO & "'"
     End If
     CnBd.Execute (strCadena)
   
   If Val(Me.txtprecio.Text) < Val(Me.lblprecio_original.Caption) And KEY_CAMBIO_PRECIO_PASS = "si" Then
        Procedencia = modificar_precio_unitario
        Call disabled_form(FrmVentas)
        Call disabled_form(Me)
        frmsegurity.Show
        Exit Sub
    Else
        Call put_actualizar_interes_nota(FrmVentas.DtcTipoDoc.BoundText, FrmVentas.DtcComprobanteGuia.BoundText, FrmVentas.TxtSeri_guia.Text, FrmVentas.TxtNumero_guia.Text, FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0), Val(Me.txtCantidad.Text))
        Call modificar_precios
    End If
   End If
End If
End Sub
Private Sub put_actualizar_interes_nota(ByVal in_doc As String, ByVal in_doc_origen As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_detalle_temporal As Double, ByVal in_cantidad)
Dim in_venta As String
Dim in_interes As Single
Dim in_total_interes As Single
Dim in_nuevo_interes As Single
If in_doc = "0007" Then 'devolucion
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & in_doc_origen & "' and serie='" & in_serie & "' and numero='" & in_numero & "' and ruc='" & KEY_RUC & "'  LIMIT 1"
    Call ConfiguraRstL(strCadena)
    If rstL.RecordCount > 0 Then
        in_venta = rstL("id_venta")
        in_interes = rstL("interes")
        strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & in_venta & "' and  id_producto='" & KEY_PRODUCTO_INTERES & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            in_total_interes = rstL("total")
            ' prorratear interes
            'ACTUALIZACION DEL NUEVO TOTAL
            strCadena = "SELECT * FROM  temporal_ventas WHERE id='" & in_detalle_temporal & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRstP(strCadena)
            If rstP.RecordCount > 0 Then
                in_total_parcial = in_cantidad * rstP("precio")
                strCadena = "UPDATE temporal_ventas SET cantidad='" & in_cantidad & "',total='" & in_total_parcial & "' WHERE id='" & in_detalle_temporal & "' AND ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
            End If
            ' PRORRATEO GENERAL
            strCadena = "SELECT * FROM temporal_ventas WHERE id_producto<>'" & KEY_PRODUCTO_INTERES & "' and  id_doc='" & in_doc & "' and dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and   ruc='" & KEY_RUC & "' ORDER BY id ASC"
            Call ConfiguraRstZ(strCadena)
            If rstZ.RecordCount > 0 Then
               rstZ.MoveFirst
               in_nuevo_interes = 0
               For i = 0 To rstZ.RecordCount - 1
                    in_nuevo_interes = in_nuevo_interes + rstZ("total") * in_interes / 100
                    rstZ.MoveNext
               Next i
               strCadena = "UPDATE temporal_ventas SET precio='" & in_nuevo_interes & "',total='" & in_nuevo_interes & "' WHERE id_producto='" & KEY_PRODUCTO_INTERES & "' AND dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and   ruc='" & KEY_RUC & "'"
               CnBd.Execute (strCadena)
            End If
            
            
        End If
    End If
End If


End Sub
Private Sub modificar_precios()
Dim nTotal As Single
    FrmVentas.TxtCodProducto.Enabled = True
    cantidad = Val(Me.txtCantidad.Text)
    If cantidad > 0 Then
        If FrmVentas.DtcTipoDoc.BoundText = "0099" And KEY_UPDATE_PROFORM = "si" And Val(FrmVentas.txtidVenta.Text) > 0 And FrmVentas.cmdModificar.Enabled = False Then
                nTotal = Val(Me.txtprecio.Text) * cantidad
                strCadena = "UPDATE movimiento_venta_detalle SET cantidad='" & cantidad & "',precio='" & Val(Me.txtprecio.Text) & "',total='" & nTotal & "' WHERE id_detalle_venta='" & Val(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0)) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                CnBd.Execute (strCadena)
                
                Call FrmVentas.llenarGrid_Comprobante_edit(FrmVentas.HfdDetalle, Val(FrmVentas.txtidVenta.Text))
                Call FrmVentas.Resalta(FrmVentas.TxtCodProducto)
                FrmVentas.cmdProcesar.Enabled = True
        Else
                
                
                nTotal = Val(Me.txtprecio.Text) * cantidad
                                    
               If control_stock_general(Trim(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 1)), cantidad, FrmVentas.DtcTipoDoc.BoundText) = True Then
                    strCadena = "UPDATE temporal_ventas SET cantidad='" & cantidad & "',precio='" & Val(Me.txtprecio.Text) & "',total='" & nTotal & "' WHERE id='" & Val(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0)) & "' AND dni_save='" & KEY_USUARIO & "'"
                    CnBd.Execute (strCadena)
               End If
                
                If KEY_BONIFICACIONES = "si" Then
                    
                    
                    'Call put_verificar_bonificacion(Trim(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 1)), Val(Me.txtcantidad.Text), Trim(FrmVentas.TxtCodCliente.Text), FrmVentas.DtcTipoDoc.BoundText, FrmVentas.DtcSerieDoc.BoundText)
                    
                    
                   ' Call quitar_bonificacion_linea(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0))
                    
                    strCadena = "CALL get_idTemporalventas('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
                    Call ConfiguraRst(strCadena)
                    in_idVenta = rst(0)
                   
                    strCadena = "CALL put_bonificacion_linea('" & Trim(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 1)) & "','" & Trim(FrmVentas.TxtCodCliente.Text) & "','" & KEY_USUARIO & "','" & KEY_ALM & "','" & in_idVenta & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    
                    
                    Call put_verificar_bonificacion_monto(Trim(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 1)), Val(Me.txtCantidad.Text), Trim(FrmVentas.TxtCodCliente.Text), FrmVentas.DtcTipoDoc.BoundText, FrmVentas.DtcSerieDoc.BoundText)
                    
                    Call FrmVentas.quitar_bonificacion_cruzada(Trim(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 1)), FrmVentas.TxtCodCliente.Text, Trim(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0)))
                    Call put_verificar_bonificacion_cruzada_v2(Trim(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 1)), Val(Me.txtCantidad.Text), Trim(FrmVentas.TxtCodCliente.Text), FrmVentas.DtcTipoDoc.BoundText, FrmVentas.DtcSerieDoc.BoundText)
                
              
                
                
                
                End If
                
                
                Call FrmVentas.llenarGrid_det(FrmVentas.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText, FrmVentas.txtformato_impresion.Text)
                Call FrmVentas.Resalta(FrmVentas.TxtCodProducto)
        End If
               
        Unload Me
        Exit Sub
End If
End Sub




Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Call actualizar_precio_cantidad
        
    
End If
End Sub
Public Sub actualizar_precio_cantidad()
If Val(Me.txtprecio.Text) < Val(Me.lblprecio_original.Caption) And KEY_CAMBIO_PRECIO_PASS = "si" Then
        Procedencia = modificar_precio_unitario
        Call disabled_form(FrmVentas)
        Call disabled_form(Me)
        frmsegurity.Show
        Exit Sub
    Else
        Call modificar_precios
    End If
    
    Exit Sub
    If FrmVentas.chkPrecios.Value = 1 And FrmVentas.HfPrecios.Rows > 0 Then
           'If FrmVentas.validar_precio(Format(FrmVentas.HfPrecios.TextMatrix(FrmVentas.HfPrecios.Row, 1), "###00.00"), Val(Me.txtprecio.Text)) = True Then
           '     Exit Sub
           ' End If
    Else
          ' If FrmVentas.validar_precio(Val(FrmVentas.txtpreciooriginal.Text), Val(Me.txtprecio.Text)) = True Then
          '      Exit Sub
          ' End If
    End If
    Call modificar_precios
End Sub

Public Sub update_precio()
    
    If KEY_GRIFO = "si" Then
         Me.txtCantidad = Val(Me.TxtTotal.Text) / Val(Me.lblprecio_original.Caption)
         Me.txtprecio.Text = Val(Me.lblprecio_original.Caption)
    Else
        Me.txtprecio.Text = Format(Val(Me.TxtTotal.Text) / Val(Me.txtCantidad.Text), "###0.00")
    End If
    
    
    
    If Val(Me.txtprecio.Text) > 0 Then
    strCadena = "UPDATE temporal_ventas SET cantidad='" & Val(Me.txtCantidad.Text) & "',precio='" & Val(Me.txtprecio.Text) & "',total='" & Val(Me.TxtTotal.Text) & "' WHERE id='" & Val(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0)) & "' AND dni_save='" & KEY_USUARIO & "'"
    CnBd.Execute (strCadena)
    
    Call FrmVentas.llenarGrid_det(FrmVentas.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText, FrmVentas.txtformato_impresion.Text)
    Call FrmVentas.Resalta(FrmVentas.TxtCodProducto)
    End If
    Unload Me
    Exit Sub
End Sub
Public Sub update_precio_unitario()
    Me.txtprecio.Text = Format(Val(Me.txtprecio), "###0.00")
    If Val(Me.txtprecio.Text) > 0 Then
        Me.TxtTotal.Text = Val(Me.txtprecio.Text) * Val(Me.txtCantidad.Text)
        If KEY_UPDATE_PROFORM = "si" And FrmVentas.DtcTipoDoc.BoundText = "0099" And Val(FrmVentas.txtidVenta.Text) > 0 Then
            strCadena = "UPDATE movimiento_venta_detalle  SET cantidad='" & Val(Me.txtCantidad.Text) & "',precio='" & Val(Me.txtprecio.Text) & "',total='" & Val(Me.TxtTotal.Text) & "' WHERE id_detalle_venta='" & Val(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "' "
            CnBd.Execute (strCadena)
            Call FrmVentas.llenarGrid_Comprobante_edit(FrmVentas.HfdDetalle, Val(FrmVentas.txtidVenta.Text))
        Else
            strCadena = "UPDATE temporal_ventas SET cantidad='" & Val(Me.txtCantidad.Text) & "',precio='" & Val(Me.txtprecio.Text) & "',total='" & Val(Me.TxtTotal.Text) & "' WHERE id='" & Val(FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 0)) & "' AND dni_save='" & KEY_USUARIO & "'"
            CnBd.Execute (strCadena)
            Call FrmVentas.llenarGrid_det(FrmVentas.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText, FrmVentas.txtformato_impresion.Text)
        End If
        Call FrmVentas.Resalta(FrmVentas.TxtCodProducto)
    End If
        Unload Me
        Exit Sub
End Sub


Private Sub txtTotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Val(Me.txtCantidad.Text) > 0 Then
    If Val(Me.TxtTotal.Text) < Val(Me.lblTotal.Caption) And KEY_CAMBIO_PRECIO_PASS = "si" Then
        Procedencia = modificar_precio
        Call disabled_form(FrmVentas)
        Call disabled_form(Me)
        frmsegurity.Show
        Exit Sub
    Else
        Call update_precio
    End If
End If
End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmOrdeneSalida 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   18000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdsalir 
      Height          =   255
      Left            =   17520
      Picture         =   "frmOrdeneSalida.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VitekeySoft.ChameleonBtn cmdconsultar 
      Height          =   735
      Left            =   6120
      TabIndex        =   9
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "CONSULTAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrdeneSalida.frx":2EA4
      PICN            =   "frmOrdeneSalida.frx":2EC0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chk_salidas_automaticas 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "PARAR SALIDAS AUTOMATICAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   14280
      TabIndex        =   8
      Top             =   480
      Width           =   3615
   End
   Begin VB.Timer timmer_generar_orden 
      Interval        =   4000
      Left            =   10560
      Top             =   480
   End
   Begin VB.TextBox txtid_ordensalida 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9840
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtcliente 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   4335
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "VENDEDOR :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DtcVendedor 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   8388608
      CalendarTitleForeColor=   8388608
      Format          =   162922497
      CurrentDate     =   42236
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPendientes 
      Height          =   6975
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   15495
      _ExtentX        =   27331
      _ExtentY        =   12303
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VitekeySoft.ChameleonBtn cmdgenerarOrden 
      Height          =   735
      Left            =   15840
      TabIndex        =   11
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "GENERAR SALIDA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrdeneSalida.frx":31DA
      PICN            =   "frmOrdeneSalida.frx":31F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdImprimirOrden 
      Height          =   735
      Left            =   15840
      TabIndex        =   12
      Top             =   2160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "PRINT SALIDA    "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrdeneSalida.frx":5AE0
      PICN            =   "frmOrdeneSalida.frx":5AFC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FECHA     :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CLIENTE  :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   8505
      Left            =   0
      Top             =   0
      Width           =   18000
   End
End
Attribute VB_Name = "frmOrdeneSalida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private obj_Word As Object



Private Sub ChameleonBtn2_Click()

End Sub

Private Sub chk_salidas_automaticas_Click()
If Me.chk_salidas_automaticas.Value = 1 Then
    Me.chk_salidas_automaticas.Caption = "PARAR SALIDAS AUTOMATICAS"
    Me.timmer_generar_orden.Enabled = True
Else
    Me.chk_salidas_automaticas.Caption = "INICIAR SALIDAS AUTOMATICAS"
    Me.timmer_generar_orden.Enabled = False
End If
End Sub

Private Sub cmdconsultar_Click()

Dim strpersona As String
strpersona = ""
strpersona = Trim(Me.txtCliente.Text)

strCadena = "SELECT * FROM view_orden_salida WHERE ncliente LIKE '%" & Trim(Me.txtCliente.Text) & "%' and   fecha_emision='" & Format(Me.DTPicker1.Value, "YYYY-mm-dd") & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
Call llenar_pendientes(Me.HfPendientes)



End Sub


Private Sub cmdgenerarOrden_Click()
If Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) > 0 Then
    in_orden_salida = put_orden_salida(Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)), Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 8))
    Call orden_salida_tiketera(in_orden_salida)
    
    strCadena = "SELECT * FROM view_orden_salida WHERE  fecha_emision='" & KEY_FECHA & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
    Call llenar_pendientes(Me.HfPendientes)
End If

'Call put_orden_salida(Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)))
'Call printer_orden_salida(Val(Me.txtid_ordensalida.Text), Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)), KEY_USUARIO)
End Sub

Private Sub buscar_pendiente_orden()
Dim in_orden_salida As String
strCadena = "SELECT id_venta,tipo FROM view_salida_pendiente WHERE  fecha_emision='" & KEY_FECHA & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1 "
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    Me.timmer_generar_orden.Enabled = False
    in_orden_salida = put_orden_salida(rstK("id_venta"), rstK("tipo"))
    
    Call orden_salida_tiketera(in_orden_salida)
    
    strCadena = "SELECT * FROM view_orden_salida WHERE  fecha_emision='" & KEY_FECHA & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
    Call llenar_pendientes(Me.HfPendientes)
    
    
    Me.timmer_generar_orden.Enabled = True
End If

End Sub






Private Function put_orden_salida(ByVal in_venta As String, ByVal in_tipo As String) As Double

    strCadena = "SELECT numero,serie FROM movimiento_venta WHERE id_doc='0500' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        in_serie = rst("serie")
        in_numero = Format(Val(rst("numero") + 1), "000000")
    Else
        strCadena = "SELECT serie,numero FROM almacen_comprobante WHERE id_doc='0500' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           in_serie = rst("serie")
           in_numero = rst("numero")
        Else
            MsgBox "ORDEN DE SALIDA NO CONFIGIGURADO", vbInformation, "CORDINE CON EL AREA DE SISTEMAS"
            Exit Function
        End If
        
    End If
    
    ':::::VERIFICACION TIPO
    If in_tipo = "01" Then '::: ventas
        strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & Val(in_venta) & "'  and ruc='" & KEY_RUC & "'"
        in_guia = 0
    Else
                           '::: guias
        strCadena = "SELECT * FROM view_guia_diferida WHERE id_venta='" & Val(in_venta) & "'  and ruc='" & KEY_RUC & "'"
        in_guia = in_venta
    End If
    
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Call insertar_item_venta(rst("id_venta"), rst("id_cliente"), rst("id_alm"), "0500", in_serie, in_numero, rst("dni_save"), in_tipo)
            
            Documento = "ORDEN SALIDA:" & in_serie & "-" & in_numero
            in_hora = str(Time)
            
            strCadena = "call p_insert_venta_orden_salida('0500','" & rst("id_alm") & "','" & rst("id_forma_pago") & "','" & rst("id_moneda") & "','" & rst("id_delivery") & "'," & _
            "'" & in_serie & "','" & in_numero & "','" & rst("id_cliente") & "','" & rst("ncliente") & "','" & rst("valor_venta") & "','" & rst("igv") & "','" & rst("exonerado") & "','" & rst("total") & "','" & rst("saldo") & "', " & _
            "'" & rst("monto_pago") & "','" & rst("monto_vuelto") & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & rst("id_tipo_factura") & "','" & rst("id_vendedor") & "','" & KEY_USUARIO & "','" & KEY_CAMBIO & "','no','" & Format(Month(rst("fecha_emision")), "00") & "','" & Year(rst("fecha_emision")) & "'" & _
            ",'" & Documento & "','" & rst("hora") & "','" & rst("turno") & "','" & rst("direccion") & "','0','-','-','" & rst("id_tipo_nota") & "','" & rst("motivo_nota") & "','" & in_guia & "','" & rst("nguia") & "','" & rst("id_ventanilla") & "','01', " & _
            "'0','" & rst("observacion") & "','no','si','1212','70111','0','0','0','" & in_venta & "','no','" & in_tipo & "','" & KEY_RUC & "')"
            Call ConfiguraRstPP(strCadena)
            id_venta = rstPP("in_venta")
            Me.txtid_ordensalida.Text = id_venta
            StrNumero = Format(Trim(str(Val(in_numero)) + 1), "000000")
            strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE  id_doc='0500' AND serie='" & in_serie & "'  AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
            strCadena = "UPDATE movimiento_venta SET id_almacenero='" & KEY_USUARIO & "' WHERE id_venta='" & in_venta & "'"
            CnBd.Execute (strCadena)
            put_orden_salida = Val(Me.txtid_ordensalida.Text)
End If

End Function
Public Sub insertar_item_venta(ByVal in_venta As String, ByVal in_cliente As String, ByVal in_alm As String, ByVal in_tipo_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_save As String, ByVal in_tipo As String)
        
        
        strCadena = "DELETE from temporal_ventas where ruc='" & KEY_RUC & "' and dni_save='" & KEY_USUARIO & "'"
        CnBd.Execute (strCadena)
        
   If in_tipo = "01" Then
        strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & in_venta & "' "
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
           rstK.MoveFirst
           For i = 0 To rstK.RecordCount - 1
                strCadena = "INSERT INTO temporal_ventas(ruc,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save) VALUES " & _
                "('" & KEY_RUC & "','" & in_cliente & "','" & Format(in_alm, "00000") & "','" & in_tipo_doc & "','" & in_serie & "','" & in_numero & "','" & rstK("id_producto") & "','" & rstK("cantidad") & "'," & _
                "'" & rstK("precio") & " ','" & rstK("total") & "','" & rstK("peso") & "','si','" & rstK("detalle") & "','" & KEY_USUARIO & "')"
                CnBd.Execute (strCadena)
            rstK.MoveNext
           Next i
        End If
   Else
        strCadena = "SELECT * FROM movimiento_transferencia_detalle WHERE id_transferencia='" & in_venta & "' and ruc='" & KEY_RUC & "' "
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
           rstK.MoveFirst
           For i = 0 To rstK.RecordCount - 1
                strCadena = "INSERT INTO temporal_ventas(ruc,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save) VALUES " & _
                "('" & KEY_RUC & "','" & in_cliente & "','" & Format(in_alm, "00000") & "','" & in_tipo_doc & "','" & in_serie & "','" & in_numero & "','" & rstK("id_producto") & "','" & rstK("cantidad") & "'," & _
                "'0 ','0','" & rstK("peso") & "','si','" & rstK("detalle") & "','" & KEY_USUARIO & "')"
                CnBd.Execute (strCadena)
            rstK.MoveNext
           Next i
        End If
   
   End If
        
      
    

End Sub

Private Sub cmdImprimirOrden_Click()
Dim in_orden_salida As String
If Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) > 0 Then
    If Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 8) = "03" Then
        strCadena = "SELECT id_orden_salida FROM movimiento_transferencia where id_transferencia='" & Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Else
        strCadena = "SELECT id_orden_salida FROM movimiento_venta where id_venta='" & Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    End If
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       
            Call orden_salida_tiketera(rst("id_orden_salida"))
        
    End If
End If

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Function get_orden_salida_guia(ByVal in_guia As String) As String
strCadena = "SELECT * FROM movimiento_transferencia WHERE id_transferencia='" & Val(in_guia) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
    get_orden_salida_guia = rstIN("id_orden_salida")
Else
    get_orden_salida_guia = 0
End If
End Function


Private Sub DtcVendedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT id_venta,ncliente,documento,total,referencia,vendedor FROM view_listado_pendientes_ref WHERE id_vendedor='" & Me.DtcVendedor.BoundText & "' and  fecha_emision='" & KEY_FECHA & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' ORDER BY id_venta DESC"
    Call llenar_pendientes(Me.HfPendientes)
End If
End Sub
Private Sub Form_Load()
Dim pantalla As Single
CenterForm Me
'pantalla = Screen.Width
'pantalla1 = (pantalla - FrmVentas.Width) / 2
Me.DTPicker1.Value = KEY_FECHA
'pantalla3 = FrmVentas.Width - pantalla1
'Me.Left = pantalla3 - 500
Me.Top = 50

strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad  WHERE  id_personal='si' and habilitado='si' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcVendedor)
Me.DtcVendedor.BoundText = 0
  
  

Me.timmer_generar_orden.Enabled = False
strCadena = "SELECT * FROM view_orden_salida WHERE  fecha_emision='" & KEY_FECHA & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
Call llenar_pendientes(Me.HfPendientes)

Me.chk_salidas_automaticas.Caption = "INICIAR SALIDAS AUTOMATICAS"

End Sub
Public Sub llenar_pendientes(ByVal Grilla As MSHFlexGrid)


Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 2200
           Grilla.ColWidth(3) = 1500
           Grilla.ColWidth(4) = 4000
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 2300
           Grilla.ColWidth(7) = 2500
           Grilla.ColWidth(8) = 0
           
        Next
        cabecera = "IDVENTA" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "DNI/RUC" & vbTab & "CLIENTE" & vbTab & "TOTAL" & vbTab & "ORDEN" & vbTab & "ALMACENERO" & vbTab & "tipo"
        Grilla.AddItem cabecera
         For k = 1 To 7
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
      
            
        For i = 0 To rstI.RecordCount - 1
          
            
          Fila = rstI("id_venta") & vbTab & Format(rstI("fecha_emision"), "dd-mm-YYYY") & vbTab & rstI("documento") & vbTab & rstI("id_cliente") & vbTab & rstI("ncliente") & vbTab & Format(rstI("total"), "#,##0.00") & vbTab & rstI("orden") & vbTab & rstI("nombre_completo") & vbTab & rstI("tipo_salida")
          Grilla.AddItem Fila
        
        
          
          rstI.MoveNext
      Next i
    
End Sub

Private Sub HfPendientes_DblClick()

If Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) > 0 Then
    FrmVentas.txt_id_pendiente.Text = Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0))
    Call FrmVentas.get_comprobante(Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)))
    FrmVentas.timer_pendientes.Enabled = False
    Unload Me

End If

End Sub



Private Sub timmer_generar_orden_Timer()
Call buscar_pendiente_orden
End Sub

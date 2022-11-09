VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmgestion_college 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IMPORTAR COMPROBANTES"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   4080
      TabIndex        =   7
      Top             =   1440
      Width           =   15135
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   720
         Top             =   4680
      End
      Begin VB.TextBox TxtIdVenta 
         Height          =   285
         Left            =   600
         TabIndex        =   14
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_sunat_key 
         Height          =   285
         Left            =   600
         TabIndex        =   13
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_hash 
         Height          =   285
         Left            =   600
         TabIndex        =   12
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DtcMes 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   4194304
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
      Begin VitekeySoft.ChameleonBtn cmdLoadExcel 
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "IMPORTAR EXCEL                      "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmgestion_college.frx":0000
         PICN            =   "frmgestion_college.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdGenerarComprobante 
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "GENERAR COMPROBANTES"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmgestion_college.frx":2906
         PICN            =   "frmgestion_college.frx":2922
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfproductos 
         Height          =   6255
         Left            =   2640
         TabIndex        =   11
         Top             =   480
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   11033
         _Version        =   393216
         ForeColor       =   8388608
         Rows            =   0
         FixedRows       =   0
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
   End
   Begin VitekeySoft.ChameleonBtn cmdexit 
      Height          =   375
      Left            =   19320
      TabIndex        =   6
      Top             =   8040
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmgestion_college.frx":520C
      PICN            =   "frmgestion_college.frx":5228
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdnivelgrado 
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   1395
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "NIVELES Y VACANTES                     "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmgestion_college.frx":5618
      PICN            =   "frmgestion_college.frx":5634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmddocentes 
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   5280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "REGISTRO DE DOCENTES               "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmgestion_college.frx":7F1E
      PICN            =   "frmgestion_college.frx":7F3A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdcursos 
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   3315
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "REGISTRO DE CURSOS                     "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmgestion_college.frx":A824
      PICN            =   "frmgestion_college.frx":A840
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdmatricula 
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   2355
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "REGISTRO  MATRICULADOS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmgestion_college.frx":D12A
      PICN            =   "frmgestion_college.frx":D146
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdgenerarmensualidad 
      Height          =   735
      Left            =   600
      TabIndex        =   4
      Top             =   4275
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "GENERACION MENSUALIDAD"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmgestion_college.frx":FA30
      PICN            =   "frmgestion_college.frx":FA4C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdperiodocollege 
      Height          =   735
      Left            =   600
      TabIndex        =   5
      Top             =   435
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "PERIODO ESCOLAR                    "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmgestion_college.frx":12336
      PICN            =   "frmgestion_college.frx":12352
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image2 
      Height          =   1230
      Left            =   5640
      Picture         =   "frmgestion_college.frx":14C3C
      Top             =   240
      Width           =   5250
   End
   Begin VB.Shape Shape8 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8655
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "frmgestion_college"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmddocentes_Click()
frmpersonal.Show
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdGenerarComprobante_Click()

Me.Timer1.Enabled = True

End Sub
Public Sub put_empezar()
Me.Timer1.Enabled = False
strCadena = "SELECT * FROM movimiento_venta_excel  WHERE comprobante='-' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' ORDER BY id ASC LIMIT 1"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   rstLocal.MoveFirst
  
    
    Call put_factura(rstLocal("dni"), rstLocal("alumno"), rstLocal("direccion"), rstLocal("monto"), rstLocal("operacion"), "262", rstLocal("descripcion"))
   
  
Else
    MsgBox "TODOS LOS COMPROBANTES SE HAN GENERADO", vbInformation
    
    Me.Timer1.Enabled = False
    Call llenar_productos(Me.hfproductos)
    Exit Sub
End If
End Sub
Public Sub put_factura(ByVal in_dni As String, ByVal in_nombre As String, ByVal in_direccion As String, ByVal in_monto As Single, ByVal in_operacion As String, ByVal in_forma_pago As String, ByVal in_detalle As String)
Dim in_doc As String
Dim in_serie As String
Dim in_numero As String

Dim in_subtotal As Single
Dim in_igv As Single
Dim in_exonerado As Single
Dim in_total As Single
strCadena = "call P_nueva_venta('" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

If Len(in_dni) = 8 Then
    in_doc = "0003"
Else
    in_doc = "0001"
End If

strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='" & in_doc & "' and electronico='si' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   in_serie = rst("serie")
   strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & in_doc & "' and serie='" & rst("serie") & "' and id_alm='" & KEY_ALM & "' and  ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      in_numero = Format(Val(rst("numero")) + 1, "000000")
   Else
      in_numero = "000001"
   End If
      'inserta temporal
      strCadena = "INSERT INTO temporal_ventas(ruc,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,costo,servicio) VALUES " & _
      "('" & KEY_RUC & "','" & in_dni & "','" & KEY_ALM & "','" & in_doc & "','" & in_serie & "','" & in_numero & "','00','1'," & _
      "'" & in_monto & "','" & in_monto & "','0','no','" & in_detalle & "','" & KEY_USUARIO & "','1','si')"
      CnBd.Execute (strCadena)

     

    strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,cuenta_contable,cuotas,id_usuario,id_recibo,detalle,banco,cheque,fecha,id_alm,serie_nota,numero_nota,ruc) VALUES " & _
    " ('" & in_doc & "','" & in_serie & "','" & in_numero & "','01','" & in_forma_pago & "','00001','" & in_monto & "','" & in_monto & "','00','','" & in_operacion & "','" & get_cuenta_contable_caja(in_forma_pago) & "','0','" & KEY_USUARIO & "','0','','-','-','" & KEY_FECHA & "','" & KEY_ALM & "','0','0','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    in_cta_cobrar = KEY_CTA_COBRAR_SERVICIO
    in_cta_ingreso = KEY_CTA_INGRESO_SERVICIO

in_valorventa = in_monto
in_igv = 0
in_exonerado = in_monto
in_total = in_monto
in_documento = "BOLETA:" & in_serie & "-" & in_numero
    
    strCadena = "call p_insert_venta_cabecera_premiun_demo('" & in_doc & "','" & KEY_ALM & "','01','00001','no'," & _
    "'" & Trim(in_serie) & "','" & Trim(in_numero) & "','" & in_dni & "','" & in_nombre & "','" & in_valorventa & "','" & in_igv & "','" & in_exonerado & "','" & in_total & "','0'," & _
    "'" & in_total & "','0','" & KEY_FECHA & "','" & KEY_FECHA & "','00001','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & Val(KEY_CAMBIO_COMPRA) & "','no','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "'" & _
    ",'" & in_documento & "',CURDATE(),'T','" & Trim(in_direccion) & "','no','" & Trim(Me.txt_hash.Text) & "','" & Trim(Me.txt_sunat_key.Text) & "',' ',' ',' ',' ','" & KEY_VENTANILLA & "','01','0','-','no','" & KEY_CONTABILIDAD & "','" & in_cta_cobrar & "','" & in_cta_ingreso & "','0','0','0','0','no','" & KEY_RUC & "')"
    
    
    
    Call ConfiguraRstPP(strCadena)
    id_venta = rstPP("in_venta")
    Me.txtidVenta.Text = id_venta
    
    Call update_excel(in_dni, in_operacion, in_documento)
    
    
    Call put_correlativo_venta(in_doc, in_serie, in_numero)
    
        If KEY_FACTURACION_ELECTRONICA = "si" Then
                If get_firma_online(in_doc, in_serie) = "si" Then
                   Call firma_electronica(in_doc, "no", " ", id_venta, in_numero, in_serie, in_dni, in_nombre, in_direccion)
                   Exit Sub
                End If
           End If
     End If
      
End Sub
Private Sub update_excel(ByVal in_dni As String, ByVal in_operacion As String, ByVal in_comprobante As String)

strCadena = "UPDATE movimiento_venta_excel SET comprobante= '" & in_comprobante & "' WHERE dni='" & in_dni & "' and operacion='" & in_operacion & "'"
CnBd.Execute (strCadena)

End Sub


Private Function firma_electronica(ByVal in_doc As String, ByVal in_extranjero As String, ByVal in_observacion As String, ByVal in_venta As String, ByVal numero As String, ByVal in_serie As String, ByVal in_dni As String, ByVal in_alumno As String, ByVal in_direccion As String)
Dim in_moneda As String
Call disabled_form(Me)
FrmLoad_web_service.Show
FrmLoad_web_service.nom_prcedimiento = "procesar_firma_electronica_local"
Set FrmLoad_web_service.FormPadre = Me

Select Case in_doc
    Case "0003"
         in_tipo_doc = "1"
         If Trim(in_extranjero) = "si" Then
            in_tipo_doc = "4"
         End If
    Case "0001"
        in_tipo_doc = "6"
         If Trim(in_extranjero) = "si" Then
            in_tipo_doc = "4"
         End If
         
    Case "0002"
        in_tipo_doc = "1"
End Select

    id_motivo_nota = ""
    motivo_nota = ""
    in_serie_afectado = ""
    in_numero_afectado = ""
    in_observacion = Replace(Trim(in_observacion), "'", " ")
    
    id_motivo_nota = ""
    motivo_nota = ""
    in_serie_afectado = ""
    in_numero_afectado = ""
    in_moneda = "PEN"



If get_comprobante_produccion(in_doc, in_serie) = "si" Then
    in_numero = Trim(numero)
    If KEY_SERVIDOR_KEYFACIL = "si" Then
        Call FrmLoad_web_service.crear_json_facturacion_electronica("https://api.vitekey.com/keyfact/utils/erp-invoice?password=vitekey2018&company_id=" & KEY_TOKEN_CLOUD & "", "POST", json_facturacion_electronica_firmar_id_venta_keyfacil(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, "", in_serie_afectado, in_numero_afectado, in_moneda, in_observacion, "no", ""), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
    Else
        If KEY_SERVIDOR_CLOUD = "si" Then
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, "", in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
        Else
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_LOCAL & "'}")
        End If
    End If


Else
    in_numero = Trim(numero)
    
    
    
    If KEY_SERVIDOR_KEYFACIL = "si" Then
        Call FrmLoad_web_service.crear_json_facturacion_electronica("https://api.vitekey.com/keyfact/utils/erp-invoice?password=vitekey2018&company_id=" & KEY_TOKEN_CLOUD & "", "POST", json_facturacion_electronica_firmar_id_venta_keyfacil(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, "", in_serie_afectado, in_numero_afectado, in_moneda, in_observacion, "no", ""), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
    Else
        If KEY_SERVIDOR_CLOUD = "si" Then
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, "", in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
        Else
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_LOCAL & "'}")
        End If
    End If
    
    
    
    
    
    
    
    
    
End If



End Function

Public Sub procesar_firma_electronica_local(ByVal strHtml As String)
On Error GoTo procesar_nuevamente
Dim in_error As Boolean
Dim in_hash As String
Dim in_numero() As String
Dim json_r As Object
Me.txt_hash.Text = ""
in_hash = ""
Set json_r = JSON.parse(strHtml)
in_error = json_r.Item("error")
If in_error = True Then

Else
     
     If KEY_SERVIDOR_KEYFACIL = "si" Then
        in_hash = Trim(json_r.Item("response").Item("id"))
        in_key = Trim(json_r.Item("response").Item("id"))
        'get_numero = Trim(json_r.Item("response").Item("numero"))
     Else
        in_hash = Trim(json_r.Item("response").Item("digest_value"))
        in_key = Trim(json_r.Item("response").Item("key"))
        get_numero = Trim(json_r.Item("response").Item("numero"))
     End If
     
     Me.txt_hash.Text = Trim(in_hash)
     Me.txt_sunat_key.Text = Trim(in_key)
     
   
     
     strCadena = "UPDATE movimiento_venta SET sunat_key='" & Trim(in_key) & "',sunat_hash='" & Trim(in_hash) & "' WHERE id_venta='" & Val(Me.txtidVenta.Text) & "'"
     CnBd.Execute (strCadena)
     
     Me.txt_sunat_key.Text = ""
     Me.txt_hash.Text = ""
     
     Me.Timer1.Enabled = True
     Me.Enabled = True
     Exit Sub
     
     'Call procesar_comprobante
     
End If
Exit Sub
procesar_nuevamente:
MsgBox "SE PRESENTO UN PROBLEMA CON EL INTERNET" + Chr(13) + Chr(13) + "INTENTENTALO NUEVAMENTE.", vbInformation, KEY_USUAURIO
Me.Enabled = True

End Sub





Private Sub cmdgenerarmensualidad_Click()
frmgeneracionmensualidad.Show
End Sub

Private Sub cmdLoadExcel_Click()






Dim Archivo As String
Archivo = Trim("bbva" & KEY_RUC) & ".xls"
      'Dim obj As New get_excel
      Set Me.hfproductos.DataSource = Leer_Excel(App.Path & "\comparar_percy\" & Archivo, "Sheet1")
      'Me.frm_importacion.Visible = True
      'Set obj = Nothing
      
      
      strCadena = "delete from movimiento_venta_excel WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
      CnBd.Execute (strCadena)
      
      

'On Error GoTo salir
For i = 0 To Me.hfproductos.Rows - 1
      
        in_dni = ""
        in_descripcion = ""
        in_alumno = ""
        in_direcion = ""
        If Val(Me.hfproductos.TextMatrix(i, 1)) > 0 And Val(Me.hfproductos.TextMatrix(i, 4)) > 0 Then
            in_dni = Format(Trim(Me.hfproductos.TextMatrix(i, 1)), "00000000")
            strCadena = "SELECT * FROM view_estudiante WHERE  dni='" & in_dni & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
            Call ConfiguraRstK(strCadena)
            
            If rstK.RecordCount > 0 Then
               in_descripcion = Trim(Me.hfproductos.TextMatrix(i, 2)) & Space(2) & "[" & rstK("nivel") & Space(1) & rstK("grado") & "]"
               in_alumno = rstK("nombre_completo")
               in_direcion = rstK("direccion")
            Else
                MsgBox "Este Alumno no esta registrado correctamente Verifique el Grado, y el Nivel" + Chr(13) + in_dni & Space(2) & get_persona(in_dni) + Chr(13) + "Registrelo antes de Realizar los comprobantes.", vbInformation, KEY_VENDEDOR
                strCadena = "delete from movimiento_venta_excel WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                Me.hfproductos.Rows = 0
                Exit Sub
            End If
            
            in_operacion = Trim(Me.hfproductos.TextMatrix(i, 6))
            
            in_monto = Val(Me.hfproductos.TextMatrix(i, 4))
        
            strCadena = "INSERT INTO movimiento_venta_excel(`dni`,alumno,direccion,`descripcion`,`monto`,`operacion`,`dni_save`,`ruc`)VALUES" & _
            "('" & in_dni & "','" & in_alumno & "','" & in_direcion & "','" & in_descripcion & "','" & Val(in_monto) & "','" & in_operacion & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
    End If
        
Next i
'salir:
Call llenar_productos(Me.hfproductos)


End Sub
Public Sub llenar_productos(ByVal Grilla As MSHFlexGrid)

Dim in_acumulado As Double
in_acumulado = 0
strCadena = "SELECT * FROM movimiento_venta_excel WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' ORDER BY id ASC"
Call ConfiguraRstT(strCadena)


If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
    Me.cmdGenerarComprobante.Visible = False
    Exit Sub
End If
   
   Me.cmdGenerarComprobante.Visible = True
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
        For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 500 'codigo
           Grilla.ColWidth(1) = 1100 'codigo
           Grilla.ColWidth(2) = 2400 'producto
           Grilla.ColWidth(3) = 5000 'linea
           Grilla.ColWidth(4) = 1000 'modelo
           Grilla.ColWidth(5) = 1000 'unidad
           Grilla.ColWidth(6) = 1500 'unidad
           Grilla.ColWidth(7) = 0 'unidad
           Grilla.ColWidth(8) = 0 'unidad
           Grilla.ColWidth(9) = 0 'unidad
           Grilla.ColWidth(10) = 0 'unidad
        Next
        cabecera = "ITEM" & vbTab & "DNI" & vbTab & "ALUMNO" & vbTab & "DESCRIPCION" & vbTab & "MONTO" & vbTab & "OPERACION" & vbTab & "COMPROBANTE"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
       
            in_acumulado = in_acumulado + rstT("monto")
            Fila = Format(i + 1, "000") & vbTab & rstT("dni") & vbTab & rstT("alumno") & vbTab & rstT("descripcion") & vbTab & Format(rstT("monto"), "#,##0.00") & vbTab & rstT("operacion") & vbTab & rstT("comprobante")
            Grilla.AddItem Fila
            rstT.MoveNext
      Next i
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "ACUMULADO :" & vbTab & Format(in_acumulado, "#,##0.00") & vbTab & ""
            Grilla.AddItem Fila
      
    
     
End Sub

      
      


Private Sub cmdmatricula_Click()
FrmMatricula.Show
End Sub

Private Sub cmdnivelgrado_Click()
frmgradonivel.Show
End Sub

Private Sub cmdperiodocollege_Click()
frmPeriodo.Show
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50

strCadena = "SELECT id_mes as Codigo,descripcion as Descripcion FROM meses ORDER BY id_mes"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMes)

End Sub

Private Sub Timer1_Timer()

If Me.txt_sunat_key.Text = "" Then
    Call put_empezar
End If

End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmpersonaDeuda 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16110
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   16110
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmDetalle 
      BackColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   14535
      Begin VB.TextBox txtObservacion 
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
         Height          =   915
         Left            =   1320
         TabIndex        =   23
         Top             =   2640
         Width           =   6975
      End
      Begin VB.TextBox txtMonto 
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
         Height          =   320
         Left            =   1320
         TabIndex        =   17
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtServicio 
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
         Height          =   320
         Left            =   2520
         TabIndex        =   16
         Top             =   1680
         Width           =   5775
      End
      Begin VB.TextBox txtId_producto 
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
         Height          =   320
         Left            =   1320
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DtcMes 
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Top             =   1080
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
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
      Begin VitekeySoft.ChameleonBtn cmdAgregar 
         Height          =   855
         Left            =   6240
         TabIndex        =   18
         Top             =   3720
         Width           =   975
         _extentx        =   1720
         _extenty        =   1508
         btype           =   5
         tx              =   "AGREGAR"
         enab            =   -1  'True
         font            =   "frmpersonaDeuda.frx":0000
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   16777215
         bcolo           =   16777215
         fcol            =   12582912
         fcolo           =   12582912
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmpersonaDeuda.frx":0028
         picn            =   "frmpersonaDeuda.frx":0046
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   2
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdCerrar 
         Height          =   855
         Left            =   7320
         TabIndex        =   19
         Top             =   3720
         Width           =   975
         _extentx        =   1720
         _extenty        =   1508
         btype           =   5
         tx              =   "SALIR"
         enab            =   -1  'True
         font            =   "frmpersonaDeuda.frx":368E
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   16777215
         bcolo           =   16777215
         fcol            =   12582912
         fcolo           =   12582912
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmpersonaDeuda.frx":36B6
         picn            =   "frmpersonaDeuda.frx":36D4
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   2
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin VB.Label lblidCobranza 
         Height          =   855
         Left            =   9000
         TabIndex        =   25
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACION :"
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
         Left            =   120
         TabIndex        =   22
         Top             =   3000
         Width           =   1035
      End
      Begin VB.Label lblCliente 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   21
         Top             =   600
         Width           =   6975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE :"
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
         TabIndex        =   20
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO :"
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
         TabIndex        =   13
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SERVICIO :"
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
         TabIndex        =   12
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIODO :"
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
         TabIndex        =   11
         Top             =   1080
         Width           =   690
      End
   End
   Begin VB.Frame frmpago 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   785
      Left            =   240
      TabIndex        =   4
      Top             =   6330
      Visible         =   0   'False
      Width           =   7455
      Begin MSDataListLib.DataCombo DtcComprobante 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   120
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
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
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   75
         Width           =   1335
         _extentx        =   2355
         _extenty        =   661
         btype           =   3
         tx              =   "PROCESAR"
         enab            =   -1  'True
         font            =   "frmpersonaDeuda.frx":66FC
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   16777215
         bcolo           =   16777215
         fcol            =   12582912
         fcolo           =   12582912
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmpersonaDeuda.frx":6724
         picn            =   "frmpersonaDeuda.frx":6742
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin VB.Label lblComprobante 
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   480
         Width           =   4350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMPROBANTE :"
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
         Left            =   240
         TabIndex        =   5
         Top             =   165
         Width           =   1125
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   10398
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   8388608
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
   Begin VitekeySoft.ChameleonBtn cmddelete 
      Height          =   855
      Left            =   14880
      TabIndex        =   1
      Top             =   1320
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      btype           =   5
      tx              =   "ELIMINAR"
      enab            =   -1  'True
      font            =   "frmpersonaDeuda.frx":8D2A
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   12582912
      fcolo           =   12582912
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmpersonaDeuda.frx":8D52
      picn            =   "frmpersonaDeuda.frx":8D70
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdSalir 
      Height          =   855
      Left            =   14880
      TabIndex        =   2
      Top             =   5400
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      btype           =   5
      tx              =   "SALIR"
      enab            =   -1  'True
      font            =   "frmpersonaDeuda.frx":B1BC
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   12582912
      fcolo           =   12582912
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmpersonaDeuda.frx":B1E4
      picn            =   "frmpersonaDeuda.frx":B202
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdVincularPago 
      Height          =   855
      Left            =   14880
      TabIndex        =   3
      Top             =   3240
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      btype           =   5
      tx              =   "VIN PAGO"
      enab            =   -1  'True
      font            =   "frmpersonaDeuda.frx":B5F2
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   12582912
      fcolo           =   12582912
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmpersonaDeuda.frx":B61A
      picn            =   "frmpersonaDeuda.frx":B638
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   855
      Left            =   14880
      TabIndex        =   9
      Top             =   360
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      btype           =   5
      tx              =   "NUEVO"
      enab            =   -1  'True
      font            =   "frmpersonaDeuda.frx":DF24
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   12582912
      fcolo           =   12582912
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmpersonaDeuda.frx":DF4C
      picn            =   "frmpersonaDeuda.frx":DF6A
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdDetalle 
      Height          =   855
      Left            =   14880
      TabIndex        =   24
      Top             =   2280
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      btype           =   5
      tx              =   "DETALLE"
      enab            =   0   'False
      font            =   "frmpersonaDeuda.frx":E3BE
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   12582912
      fcolo           =   12582912
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmpersonaDeuda.frx":E3E6
      picn            =   "frmpersonaDeuda.frx":E404
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdgenerarComprobante 
      Height          =   855
      Left            =   14880
      TabIndex        =   26
      Top             =   4200
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      btype           =   5
      tx              =   "GEN CPE"
      enab            =   -1  'True
      font            =   "frmpersonaDeuda.frx":E720
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   12582912
      fcolo           =   12582912
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmpersonaDeuda.frx":E748
      picn            =   "frmpersonaDeuda.frx":E766
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   7155
      Left            =   0
      Top             =   0
      Width           =   16110
   End
End
Attribute VB_Name = "frmpersonaDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub ChameleonBtn1_Click()

End Sub

Private Sub ChameleonBtn2_Click()

End Sub

Private Sub cmdagregar_Click()
Call put_agregar
End Sub
Private Sub put_agregar()
If Trim(Me.txtId_producto.Text) = "" Then
   MsgBox "Ingrese un SERVICIO VALIDO", vbInformation
   Exit Sub
End If

If Trim(Me.txtServicio.Text) = "" Then
   MsgBox "Ingrese un SERVICIO VALIDO", vbInformation
   Exit Sub
End If

If Val(Me.txtMonto.Text) = 0 Then
   MsgBox "Ingrese un SERVICIO VALIDO", vbInformation
   Exit Sub
End If

   strCadena = "SELECT * FROM cobranza_periodo WHERE id_periodo='" & Val(Me.DtcMes.BoundText) & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
        in_mes = rst("id_mes")
        in_fecha_fin = DateSerial(rst("id_anio"), Val(rst("id_mes")) + 1, 1 - 1)
        in_fecha_ini = "01-" & Format(Val(rst("id_mes")), "00") & "-" & Trim(rst("id_anio"))
        in_dias = DateDiff("d", in_fecha_ini, in_fecha_fin)
        in_producto = Trim(Me.txtId_producto.Text)
        in_producto_des = Trim(Me.txtServicio.Text) & Space(1) & Trim(Me.DtcMes.Text)
        in_precio = Format(Me.txtMonto.Text, "###0.00")
        strCadena = "call sp_insert_servicio_mensual('" & Me.DtcMes.BoundText & "','" & Me.lblCliente.Tag & "','" & in_producto & "','" & in_producto_des & "','" & in_precio & "','" & KEY_USUARIO & "','" & KEY_ALM & "','" & Format(in_fecha_ini, "YYYY-mm-dd") & "','" & Format(in_fecha_fin, "YYYY-mm-dd") & "','" & UCase(Me.txtObservacion.Text) & "','" & KEY_RUC & "')"
               CnBd.Execute (strCadena)
               
               Call put_fecha_corte(Me.lblCliente.Tag)
               
               
    End If
    Me.frmDetalle.Visible = False
    Call llenarGrid_deuda(Me.HfdPersona, Me.lblCliente.Tag)



End Sub
Private Sub cmdCerrar_Click()
Me.frmDetalle.Visible = False
End Sub

Private Sub cmddelete_Click()

If MsgBox("Esta seguro de Eliminar esta Deuda", vbQuestion + vbYesNo) = vbYes Then
    
    strCadena = "DELETE FROM cobranza_servicio_persona WHERE id_detalle='" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    Call Me.llenarGrid_deuda(Me.HfdPersona, Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
    
 End If

End Sub

Private Sub cmdDetalle_Click()
Call load_cobranza(Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)))
End Sub

Private Sub load_cobranza(ByVal in_servicio As String)

strCadena = "SELECT * FROM cobranza_servicio_persona WHERE id_detalle='" & Val(in_servicio) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.DtcMes.BoundText = rst("id_periodo")
   Me.txtId_producto.Text = rst("id_servicio")
   Me.txtServicio.Text = rst("detalle")
   Me.txtMonto.Text = rst("monto")
   Me.txtObservacion.Text = rst("observacion")
   Me.frmDetalle.Visible = True
End If


End Sub

Private Sub cmdgenerarComprobante_Click()

FrmVentas.Show
Call FrmVentas.activar
FrmVentas.TxtCodCliente.Text = Me.lblCliente.Tag
Unload Me
Call FrmVentas.precionar_cliente
Call FrmVentas.put_mensualidad
Exit Sub

End Sub

Private Sub cmdNuevo_Click()

Me.frmDetalle.Visible = True
Me.DtcMes.SetFocus


End Sub

Private Sub cmdProcesar_Click()

strCadena = "UPDATE cobranza_servicio_persona SET id_venta='" & Val(Me.DtcComprobante.BoundText) & "' WHERE id_detalle='" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "'"
CnBd.Execute (strCadena)

MsgBox "Vinculo Correcto.", vbInformation


End Sub

Private Sub cmdSalir_Click()
Unload Me
Call enabled_form(FrmPersona)
End Sub

Private Sub cmdVincularPago_Click()

   
    Me.DtcComprobante.BoundText = "0001"
    Call put_serie_numero(Me.lblCliente.Tag)
    Me.frmpago.Visible = True
    

End Sub

Private Sub put_serie_numero(ByVal in_cliente As String)

strCadena = "SELECT id_venta as Codigo,CONCAT(fecha_emision,' -- ',documento) as Descripcion FROM movimiento_venta WHERE id_cliente='" & in_cliente & "' and   ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC LIMIT 5"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComprobante)
End Sub



Private Sub Form_Load()
CenterForm Me
Me.Top = 200


strCadena = "SELECT id_doc as Codigo,doc_des as Descripcion FROM comprobantes ORDER BY doc_des"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComprobante)

strCadena = "SELECT id_periodo as Codigo,descripcion as Descripcion from cobranza_periodo WHERE ruc='" & KEY_RUC & "' order by id_periodo DESC "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMes)

'

End Sub
Public Sub llenarGrid_deuda(ByVal Grilla As MSHFlexGrid, ByVal in_ruc As String)
Dim in_monto_acumulado As Double
Dim in_monto_deuda As Double
On Error GoTo salir
Me.lblCliente.Tag = in_ruc
Me.lblCliente.Caption = get_persona(in_ruc)

strCadena = "SELECT * FROM cobranza_servicio_persona WHERE dni='" & in_ruc & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
  
   Grilla.Rows = 0
            ReDim arrColWidth(1 To rst.Fields.Count)
            For Each Campo In rst.Fields
                Grilla.ColWidth(0) = 800
                Grilla.ColWidth(1) = 1200
                Grilla.ColWidth(2) = 8500
                Grilla.ColWidth(3) = 1200
                Grilla.ColWidth(4) = 2500
            Next
            cabecera = "CODIGO" & vbTab & "PERIODO" & vbTab & "PLAN CONTRATADO" & vbTab & "MONTO A FACTURAR" & vbTab & "REFERENCIA"
            Grilla.AddItem cabecera
         
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        in_monto_acumulado = 0
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If rst("id_venta") > 0 Then
               in_referencia = get_comprobante(rst("id_venta"))
            Else
               in_referencia = "PENDIENTE"
            End If
            
            Fila = Format(rst("id_detalle"), "000000") & vbTab & Trim(rst("dni")) & vbTab & rst("detalle") & vbTab & Format(rst("monto"), "#,##0.00") & vbTab & in_referencia
            Grilla.AddItem Fila
            If rst("id_venta") <= 0 Then
            in_monto_acumulado = in_monto_acumulado + rst("monto")
            End If
        rst.MoveNext
        Next i
        
         Fila = "" & vbTab & "" & vbTab & "" & vbTab & Format(in_monto_acumulado, "#,##0.00")
         Grilla.AddItem Fila
             
                For k = 3 To 3
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
               Next k
  
  Grilla.ColAlignment(0) = 1
  Grilla.ColAlignment(1) = 1
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
         
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub HfdPersona_SelChange()

If Val(Me.HfdPersona.Rows) > 0 Then
    If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
       Me.cmdDetalle.Enabled = True
    Else
       Me.cmdDetalle.Enabled = False
    End If
End If


End Sub

Private Sub txtId_producto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Procedencia = Selecionar
   FrmProducto.Show
   Exit Sub
End If
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)



End Sub

Private Sub txtserie_Change()

End Sub

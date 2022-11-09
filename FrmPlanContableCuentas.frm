VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmPlanContableCuentas 
   BorderStyle     =   0  'None
   Caption         =   "Cuentas Contables"
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtDescripcion 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   6630
      TabIndex        =   8
      Top             =   8310
      Width           =   3615
   End
   Begin VB.TextBox TxtPlanContable 
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   1560
      TabIndex        =   0
      Top             =   8310
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3360
      Top             =   2670
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanContableCuentas.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanContableCuentas.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanContableCuentas.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanContableCuentas.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanContableCuentas.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanContableCuentas.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanContableCuentas.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanContableCuentas.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPlanContableCuentas.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgPlanContable 
      Height          =   6975
      Left            =   240
      TabIndex        =   1
      Top             =   1110
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   12303
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
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   6945
      Left            =   10485
      TabIndex        =   2
      Top             =   1080
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   12250
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   6945
      _Version        =   "6.7.9782"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   5670
         Left            =   30
         TabIndex        =   3
         Top             =   375
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   10001
         ButtonWidth     =   1588
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Nuevo   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Modificar"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo DtcPlanCOntable 
      Height          =   315
      Left            =   2520
      TabIndex        =   4
      Top             =   600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION CUENTA :"
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
      Left            =   4800
      TabIndex        =   9
      Top             =   8310
      Width           =   1575
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUENTAS PLAN CONTABLE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   285
      TabIndex        =   7
      Top             =   0
      Width           =   2445
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº CUENTA :"
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
      Left            =   570
      TabIndex        =   6
      Top             =   8310
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO PLAN CONTABLE :"
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
      Left            =   495
      TabIndex        =   5
      Top             =   600
      Width           =   1725
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   240
      Top             =   8190
      Width           =   10095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00DFDFE0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   240
      Top             =   360
      Width           =   10095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9015
      Left            =   0
      Top             =   0
      Width           =   11520
   End
End
Attribute VB_Name = "FrmPlanContableCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub DtcPlanCOntable_Change()
'strCadena = "SELECT     plan_contable_det.id_plancontable_det, plan_contable_det.pc_codigo, plan_contable_det.plan_des, Nivel_Cuenta.descripcion as NIVEL_CUENTA, " & _
"                 Tipo_Cuenta.descripcion AS TIPO_CUENTA FROM         plan_contable_det INNER JOIN   Nivel_Cuenta ON plan_contable_det.nivel_cuenta = Nivel_Cuenta.nivel_cuenta INNER JOIN " & _
"                 Tipo_Cuenta ON plan_contable_det.tipo_cuenta = Tipo_Cuenta.tipo_cuenta WHERE id_plancontable='" & Trim(Me.DtcPlanCOntable.BoundText) & "'"
'Call llenarGridME(Me.HfgPlanContable, Me)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()




CenterForm Me
Me.Top = 50
strCadena = "SELECT id_plancontable as Codigo, pc_descripcion as Descripcion FROM plan_contable "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.dtcPlancontable)
  Me.dtcPlancontable.BoundText = "0001"
  Call actualizar
  
 
End Sub
Public Sub actualizar()

strCadena = "SELECT * FROM con_cuentacontable WHERE IdEmpresaSis='" & KEY_RUC & "' and Ejercicio='" & Year(KEY_FECHA) & "' LIMIT 30"
Call llenarGrid(Me.HfgPlanContable, Me)


End Sub
Private Sub HfgPlanContable_Click()
If Me.HfgPlanContable.Rows > 0 Then
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    TlbAcciones.Buttons(KEY_DELETE).Enabled = True
  End If
End Sub
Private Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub

Private Sub HfgPlanContable_KeyPress(KeyAscii As Integer)
On Error GoTo error
If KeyAscii = 13 And FrmCentroCostosDetalle.Procedencia = Selecionar And FrmCentroCostosDetalle.debe_haber = debe Then
    Me.HfgPlanContable.col = 0
    strCadena = "SELECT * FROM plan_contable_det WHERE pc_codigo='" & Trim(Me.HfgPlanContable.Text) & "' AND id_plancontable='" & Trim(FrmCentroCostosDetalle.dtcPlancontable.BoundText) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmCentroCostosDetalle.txtCuentaDebe.Text = rst("pc_codigo")
        FrmCentroCostosDetalle.lblDebe.Caption = rst("plan_des")
        Unload Me
        Set rst = Nothing
    End If
    Exit Sub
End If

If FrmDetalleProducto.Procedencia = Selecionar Then
   FrmDetalleProducto.Procedencia = Neutro
   FrmDetalleProducto.txtcuenta_contable.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
   FrmDetalleProducto.lblcuentacontable.Caption = UCase(Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 2))
   Unload Me
   Exit Sub
End If


If FrmListadoFacturasCompra.Procedencia = Selecionar Then
   FrmListadoFacturasCompra.Procedencia = Neutro
   FrmListadoFacturasCompra.txtCuenta_redondeo.Text = Trim(Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1))
   Unload Me
   Exit Sub
End If

If FrmListadoFacturasCompra.Procedencia = buscar Then
   FrmListadoFacturasCompra.Procedencia = Neutro
   FrmListadoFacturasCompra.txtCuentaPrincipal.Text = Trim(Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1))
   Unload Me
   Exit Sub
End If

If FrmListadoFacturasCompra.Procedencia = seleccionar_per Then
   FrmListadoFacturasCompra.Procedencia = Neutro
   FrmListadoFacturasCompra.txtCuentaPerdida.Text = Trim(Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1))
   Unload Me
   Exit Sub
End If


If FrmListadoFacturasCompra.Procedencia = seleccionar_soldadura Then
   FrmListadoFacturasCompra.Procedencia = Neutro
   FrmListadoFacturasCompra.txtCuentaGanancia.Text = Trim(Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1))
   Unload Me
   Exit Sub
End If


If FrmReporteRegistroVentas.Procedencia = Selecionar Then
   FrmReporteRegistroVentas.Procedencia = Neutro
   FrmReporteRegistroVentas.txtCuentaPrincipal.Text = Trim(Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1))
   Unload Me
   Exit Sub
End If
If FrmReporteRegistroVentas.Procedencia = seleccionar_otro Then
   FrmReporteRegistroVentas.Procedencia = Neutro
   FrmReporteRegistroVentas.txtCuentaPerdida.Text = Trim(Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1))
   Unload Me
   Exit Sub
End If

If FrmReporteRegistroVentas.Procedencia = seleccionar_per Then
   FrmReporteRegistroVentas.Procedencia = Neutro
   FrmReporteRegistroVentas.txtCuentaGanancia.Text = Trim(Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1))
   Unload Me
   Exit Sub
End If



If FrmMisCuentasDet.Procedencia = Selecionar Then
   FrmMisCuentasDet.txtcuentacontable.Text = Trim(Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1))
   
   FrmMisCuentasDet.Procedencia = Neutro
   Unload Me
   Exit Sub
End If


If FrmListadoFacturasCompra.Procedencia = seleccionar_otro Then
   FrmListadoFacturasCompra.Procedencia = Neutro
   FrmListadoFacturasCompra.txtcuenta_redondeo_anticipo.Text = Trim(Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1))
   Unload Me
   Exit Sub
End If


If FrmReporteProducto.Procedencia = Selecionar Then
   FrmReporteProducto.Procedencia = Neutro
   FrmReporteProducto.txtcuentacontable.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
   Unload Me
   Exit Sub
End If


If FrmComprasPagos.Procedencia = Selecionar Then
   FrmComprasPagos.Procedencia = Neutro
   FrmComprasPagos.txtCuenta_redondeo.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
   Unload Me
   Exit Sub
End If




If FrmComprasPagos.Procedencia = seleccionar_otro Then
   FrmComprasPagos.Procedencia = Neutro
   FrmComprasPagos.txtCuenta_anticipo.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
   Unload Me
   Exit Sub
End If



If FrmComprasPagos.Procedencia = seleccionar_insumo Then
   FrmComprasPagos.Procedencia = Neutro
   FrmComprasPagos.txtCtaRetencion.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
   Unload Me
   Exit Sub
End If



If frmformapago.Procedencia = Selecionar Then
   frmformapago.Procedencia = Neutro
   frmformapago.txtcuentacontable.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
   frmformapago.lblcuentacontable.Caption = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 2)
   Unload Me
   Exit Sub
End If



'If frmretencion.Procedencia = debe1 Then
'   frmretencion.Procedencia = Neutro
'   frmretencion.txtcuentadebe1.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
'   Unload Me
'   Exit Sub
'End If

'If frmretencion.Procedencia = debe2 Then
'   frmretencion.Procedencia = Neutro
'   frmretencion.txtcuentadebe2.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
'   Unload Me
'   Exit Sub
'End If
'If frmretencion.Procedencia = haber1 Then
'   frmretencion.Procedencia = Neutro
'   frmretencion.TxtCuentahaber.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
'   Unload Me
'   Exit Sub
'End If

If FrmDetalleLinea.Procedencia = Selecionar Then
   FrmDetalleLinea.Procedencia = Neutro
   FrmDetalleLinea.txtcodigo_contable.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
   FrmDetalleLinea.lblcuenta_contable.Caption = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 2)
   Unload Me
   Exit Sub
End If

If FrmDetalleLinea.Procedencia = seleccionar_otro Then
   FrmDetalleLinea.Procedencia = Neutro
   FrmDetalleLinea.txtCuentaContableImportacion.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
   FrmDetalleLinea.lblcuenta_contable_import.Caption = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 2)
   Unload Me
   Exit Sub
End If

If FrmDetalleLinea.Procedencia = seleccionar_soldadura Then
   FrmDetalleLinea.Procedencia = Neutro
   FrmDetalleLinea.txtCostoDebe.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
   FrmDetalleLinea.lblcuentadebecosto.Caption = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 2)
   Unload Me
   Exit Sub
End If


If FrmDetalleLinea.Procedencia = seleccionar_tapiz Then
   FrmDetalleLinea.Procedencia = Neutro
   FrmDetalleLinea.txtCostoHaber.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
   FrmDetalleLinea.lblcuentahabercosto.Caption = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 2)
   Unload Me
   Exit Sub
End If




If frmAnalisisporCuenta.Procedencia = Selecionar Then
   frmAnalisisporCuenta.Procedencia = Neutro
   frmAnalisisporCuenta.txtCuentaIni.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
   frmAnalisisporCuenta.txtCuentaFin.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
   Unload Me
   Exit Sub
End If

If frmAnalisisporCuenta.Procedencia = buscar Then
   frmAnalisisporCuenta.Procedencia = Neutro
   frmAnalisisporCuenta.txtCuentaFin.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
   Unload Me
   Exit Sub
End If



If KeyAscii = 13 And frmNuevoComprobante.busqueda = Buscar1 Then
    frmNuevoComprobante.txtOrigen.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    frmNuevoComprobante.busqueda = NeutroCostos
    Unload Me
    Exit Sub
End If
If KeyAscii = 13 And FrmChequeNuevo.Procedencia = buscar Then
    FrmChequeNuevo.Txtcentrocosto.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    FrmChequeNuevo.lblcostos.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 2)
    FrmChequeNuevo.Procedencia = Neutro
    Call Resalta(FrmChequeNuevo.txtMotivo)
    Unload Me
    Exit Sub
End If
If KeyAscii = 13 And FrmSolicitudViaticosDet.Procedencia = buscar Then
    FrmSolicitudViaticosDet.TxtCcostos.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    FrmSolicitudViaticosDet.lblccostos.Caption = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 2)
    FrmSolicitudViaticosDet.Procedencia = Neutro
    
    Unload Me
    Exit Sub
End If

If KeyAscii = 13 And FrmMisCuentasDet.Procedencia = Selecionar Then
    FrmMisCuentasDet.txtcuentacontable.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    FrmMisCuentasDet.DtcEntidadBancaria.SetFocus
    FrmMisCuentasDet.Procedencia = Neutro
    Unload Me
    Exit Sub
End If
If KeyAscii = 13 And frmNuevoComprobante.busqueda = Buscar2 Then
    frmNuevoComprobante.TxtNaturaleza.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    frmNuevoComprobante.busqueda = NeutroCostos
    frmNuevoComprobante.TxtCostos1.SetFocus
    Unload Me
    
    Exit Sub
End If
If KeyAscii = 13 And frmNuevoComprobante.busqueda = Buscar3 Then
    frmNuevoComprobante.TxtCostos1.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    frmNuevoComprobante.busqueda = NeutroCostos
    diferencia = Val(frmNuevoComprobante.TxtMontoFactura.Text)
    frmNuevoComprobante.Monto1.Text = Format(diferencia, "###0.00")
    Call Resalta(frmNuevoComprobante.Monto1)
    Unload Me
    Exit Sub
End If
If KeyAscii = 13 And frmNuevoComprobante.busqueda = Buscar4 Then
    frmNuevoComprobante.txtCostos2.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    frmNuevoComprobante.busqueda = NeutroCostos
    diferencia = Val(frmNuevoComprobante.TxtMontoFactura.Text) - Val(frmNuevoComprobante.Monto1.Text)
    frmNuevoComprobante.Monto2.Text = Format(diferencia, "###0.00")
   Call Resalta(frmNuevoComprobante.Monto2)
    Unload Me
    Exit Sub
End If
If KeyAscii = 13 And frmNuevoComprobante.busqueda = Buscar5 Then
    frmNuevoComprobante.txtCostos3.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    frmNuevoComprobante.busqueda = NeutroCostos
    diferencia = Val(frmNuevoComprobante.TxtMontoFactura.Text) - Val(frmNuevoComprobante.Monto1.Text) - Val(frmNuevoComprobante.Monto2.Text)
    frmNuevoComprobante.Monto3.Text = Format(diferencia, "###0.00")
   Call Resalta(frmNuevoComprobante.Monto3)
    Unload Me
    Exit Sub
End If
If KeyAscii = 13 And frmNuevoComprobante.busqueda = Buscar6 Then
    frmNuevoComprobante.txtCostos4.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    frmNuevoComprobante.busqueda = NeutroCostos
    diferencia = Val(frmNuevoComprobante.TxtMontoFactura.Text) - Val(frmNuevoComprobante.Monto1.Text) - Val(frmNuevoComprobante.Monto2.Text) - Val(frmNuevoComprobante.Monto3.Text)
    frmNuevoComprobante.Monto4.Text = Format(diferencia, "###0.00")
    Call Resalta(frmNuevoComprobante.Monto4)
    Unload Me
    Exit Sub
End If

If KeyAscii = 13 And FrmCCostos.busqueda = Buscar2 Then
    FrmCCostos.TxtNaturaleza.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    FrmCCostos.busqueda = NeutroCostos
    FrmCCostos.TxtCostos1.SetFocus
    Unload Me
    
    Exit Sub
End If

If KeyAscii = 13 And FrmCCostos.busqueda = Buscar3 Then
    FrmCCostos.TxtCostos1.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    FrmCCostos.busqueda = NeutroCostos
    If Val(FrmRegistroComprasList.TxtValorcompra.Text) = 0 Then
        diferencia = Val(FrmRegistroComprasList.TxtValorCompraNoAfecta.Text)
    Else
    diferencia = Val(FrmRegistroComprasList.TxtValorcompra.Text)
    End If
    
    
    FrmCCostos.Monto1.Text = Format(diferencia, "###0.00")
    Call Resalta(FrmCCostos.Monto1)
    Unload Me
    Exit Sub
End If



If KeyAscii = 13 And FrmCCostos.busqueda = Buscar4 Then
    FrmCCostos.txtCostos2.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    FrmCCostos.busqueda = NeutroCostos
    diferencia = Val(FrmRegistroComprasList.TxtValorcompra.Text) - Val(FrmCCostos.Monto1.Text)
    FrmCCostos.Monto2.Text = Format(diferencia, "###0.00")
    Call Resalta(FrmCCostos.Monto2)
    Unload Me
    Exit Sub
End If

If KeyAscii = 13 And FrmCCostos.busqueda = Buscar5 Then
    FrmCCostos.txtCostos3.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    FrmCCostos.busqueda = NeutroCostos
    diferencia = Val(FrmRegistroComprasList.TxtValorcompra.Text) - Val(FrmCCostos.Monto2.Text) - Val(FrmCCostos.Monto1.Text)
    FrmCCostos.Monto3.Text = Format(diferencia, "###0.00")
    Call Resalta(FrmCCostos.Monto3)
    Unload Me
    Exit Sub
End If
If KeyAscii = 13 And FrmCCostos.busqueda = Buscar6 Then
    FrmCCostos.txtCostos4.Text = Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 1)
    FrmCCostos.busqueda = NeutroCostos
    diferencia = Val(FrmRegistroComprasList.TxtValorcompra.Text) - Val(FrmCCostos.Monto2.Text) - Val(FrmCCostos.Monto1.Text) - Val(FrmCCostos.Monto3.Text)
    FrmCCostos.Monto4.Text = Format(diferencia, "###0.00")
    Call Resalta(FrmCCostos.Monto4)
    Unload Me
    Exit Sub
End If



If KeyAscii = 13 And FrmCentroCostosDetalle.Procedencia = Selecionar And FrmCentroCostosDetalle.debe_haber = HABER Then
    Me.HfgPlanContable.col = 0
    strCadena = "SELECT * FROM plan_contable_det WHERE pc_codigo='" & Trim(Me.HfgPlanContable.Text) & "' AND id_plancontable='" & Trim(FrmCentroCostosDetalle.dtcPlancontable.BoundText) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmCentroCostosDetalle.TxtCuentahaber.Text = rst("pc_codigo")
        FrmCentroCostosDetalle.lelehaber.Caption = rst("plan_des")
        Unload Me
        Set rst = Nothing
    End If
    Exit Sub
End If
Exit Sub
error: MsgBox "Intente Nuevamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case KEY_NEW
      Procedencia = nuevo
      FrmDetallePlanContable.Show
    Case KEY_UPDATE
      Procedencia = modificar
     FrmDetallePlanContable.Show
    Case KEY_DELETE
      If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        strCadena = "DELETE FROM plan_contable_det WHERE id_plancontable_det='" & Val(Me.HfgPlanContable.TextMatrix(Me.HfgPlanContable.Row, 0)) & "'"
        Call Execute_Sql(strCadena)
        Call actualizar
      End If
    Case "(Salir)"
      Unload Me
  End Select
End Sub

Private Sub TxtDescripcion_Change()
'strCadena = "SELECT     plan_contable_det.id_plancontable_det, plan_contable_det.pc_codigo, plan_contable_det.plan_des, nivel_cuenta.descripcion , " & _
"                 tipo_cuenta.descripcion FROM plan_contable_det INNER JOIN   nivel_cuenta ON plan_contable_det.nivel_cuenta = nivel_cuenta.nivel_cuenta INNER JOIN " & _
"                 tipo_cuenta ON plan_contable_det.tipo_cuenta = tipo_cuenta.tipo_cuenta WHERE id_plancontable='" & Trim(Me.DtcPlanCOntable.BoundText) & "' AND  plan_des LIKE '" & Trim(Me.TxtDescripcion.Text) & "%' and ruc='" & KEY_RUC & "' ORDER BY pc_codigo ASC LIMIT 0,50 "
'Call llenarGrid(Me.HfgPlanContable, Me)

strCadena = "SELECT * FROM con_cuentacontable WHERE Descripcion LIKE '%" & Trim(Me.TxtDescripcion.Text) & "%' and IdEmpresaSis='" & KEY_RUC & "' and Ejercicio='" & Year(KEY_FECHA) & "' ORDER BY NroCuenta ASC LIMIT 0,200 "
Call llenarGrid(Me.HfgPlanContable, Me)


End Sub

Private Sub TxtPlanContable_Change()
strCadena = "SELECT * FROM con_cuentacontable WHERE NroCuenta LIKE '" & Trim(Me.TxtPlanContable.Text) & "%' AND IndMovimiento='1' and activo=1 and IdEmpresaSis='" & KEY_RUC & "' and  Ejercicio='" & Year(KEY_FECHA) & "' ORDER BY NroCuenta ASC LIMIT 0,100 "
Call llenarGrid(Me.HfgPlanContable, Me)
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
'On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
   
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 6000
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1000
         
         Next
        cabecera = "ID" & vbTab & "CUENTA" & vbTab & "PLAN CONTABLE" & vbTab & "NIVEL CUENTA" & vbTab & "TIPO CUENTA"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("id") & vbTab & rst("NroCuenta") & vbTab & UCase(rst("Descripcion")) & vbTab & "" & vbTab & ""
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
    Next i
    Grilla.ColAlignment(1) = 1
  
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  Exit Sub
'salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub






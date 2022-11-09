VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmPagoCuotaCredito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago Credito."
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   8325
   Begin VB.TextBox TxtOperacion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3960
      TabIndex        =   23
      Top             =   3360
      Width           =   1515
   End
   Begin VB.TextBox TxtNumeroRef 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4065
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtSerieRef 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3225
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox TxtSaldo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1800
      TabIndex        =   15
      Top             =   3360
      Width           =   1515
   End
   Begin VB.TextBox TxtSerie 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox TxtFactura 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox TxtCodEntidad 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   877
      Width           =   735
   End
   Begin VB.TextBox TxtDeuda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Top             =   2400
      Width           =   1515
   End
   Begin VB.TextBox TxtFechaEmision 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox TxtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   2880
      Width           =   1515
   End
   Begin VB.TextBox TxtEntidad 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2550
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   877
      Width           =   4335
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   7440
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaCredito.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaCredito.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaCredito.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaCredito.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaCredito.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaCredito.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaCredito.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaCredito.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaCredito.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaCredito.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaCredito.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPagoCuotaCredito.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   6150
      TabIndex        =   2
      Top             =   2880
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1875
      _CBHeight       =   870
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   810
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1429
         ButtonWidth     =   1402
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Grabar"
               Key             =   "(Grabar)"
               Object.ToolTipText     =   "Grabar Ctrl+G"
               ImageKey        =   "(Grabar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "(Cancelar)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Cancelar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComCtl2.DTPicker DtpFechaPago 
      Height          =   360
      Left            =   5115
      TabIndex        =   17
      Top             =   315
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   0
      CalendarForeColor=   255
      Format          =   135593987
      CurrentDate     =   39535
   End
   Begin MSDataListLib.DataCombo DtcInterno 
      Height          =   315
      Left            =   1800
      TabIndex        =   18
      Top             =   1320
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcExterno 
      Height          =   315
      Left            =   1800
      TabIndex        =   22
      Top             =   1800
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "OPERACION:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4200
      TabIndex        =   24
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nº EXTERNO:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   585
      TabIndex        =   21
      Top             =   1830
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "SALDO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   960
      TabIndex        =   16
      Top             =   3360
      Width           =   600
   End
   Begin VB.Label LblMonto 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "MONTO A PAGAR:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   1320
   End
   Begin VB.Label LblFechaPago 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "FECHA PAGO:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3960
      TabIndex        =   9
      Top             =   405
      Width           =   1020
   End
   Begin VB.Label LblNDocumento 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nº INTERNO:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   1350
      Width           =   960
   End
   Begin VB.Label LblFechaEmision 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "FECHA EMISION:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   315
      TabIndex        =   6
      Top             =   405
      Width           =   1245
   End
   Begin VB.Label LblDeuda 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "SALDO:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1005
      TabIndex        =   5
      Top             =   2400
      Width           =   555
   End
   Begin VB.Label LblEntidad 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "ENTIDAD:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   840
      TabIndex        =   4
      Top             =   915
      Width           =   720
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00A56E32&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   3660
      Left            =   120
      Top             =   150
      Width           =   8055
   End
End
Attribute VB_Name = "FrmPagoCuotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EnumFrmPago As EnumBuscarDocumento
Dim DblTotal As Double




Private Sub Form_Activate()
Me.TxtDeuda.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Save()
Dim idDetalle  As String
Dim monto_pago As Double
Dim saldo_factura As Double
monto_pago = Me.txtMonto.Text
saldo_factura = Me.TxtSaldo.Text

strCadena = "SELECT * FROM DetallePagos ORDER BY id_detalle DESC "
Call ConfiguraRst(strCadena)
 idDetalle = GeneraCodigo(8)
 Set rst = Nothing
 If monto_pago > 0 Then
    
    strCadena = "INSERT INTO DetallePagos(id_detalle,serie,numero,cPersona,doc_cod,FechaPago,Monto,Operacion)VALUES ('" & idDetalle & "','" & Trim(Me.TxtSerie.Text) & "','" & Trim(Me.TxtFactura.Text) & "'," & _
    "'" & Trim(Me.TxtCodEntidad.Text) & "','" & Trim(Me.DtcInterno.BoundText) & "','" & CVDate(Date) & "','" & monto_pago & "','" & Trim(Me.txtOperacion.Text) & "')"
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
    
    strCadena = "UPDATE DocumentoCompra SET saldo='" & Trim(saldo_factura) & "' WHERE cDocumentoCompra='" & Trim(Me.TxtFactura.Text) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcInterno.BoundText) & "' AND cPersona='" & Trim(Me.TxtCodEntidad.Text) & "'"
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
    Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = False
    FrmListadoFacturasCompra.facturas
 End If
Exit Sub

End Sub
Private Sub Form_Load()
CenterForm Me
Me.Top = 500
If FrmListadoFacturasCompra.Procedencia = nuevo Then
        Dim serie As String
        Dim Numero As String
        Dim Persona As String
        Dim doc_cod As String
        Dim doc_ref As String
        Dim cod_referencia As String
        FrmListadoFacturasCompra.HfgFacturas.col = 0
        serie = Trim(FrmListadoFacturasCompra.HfgFacturas.Text)
        FrmListadoFacturasCompra.HfgFacturas.col = 1
        Numero = Trim(FrmListadoFacturasCompra.HfgFacturas.Text)
        FrmListadoFacturasCompra.HfgFacturas.col = 2
        Persona = Trim(FrmListadoFacturasCompra.HfgFacturas.Text)
        FrmListadoFacturasCompra.HfgFacturas.col = 3
        doc_cod = Trim(FrmListadoFacturasCompra.HfgFacturas.Text)
        
        strCadena = "SELECT * FROM DocumentoCompra WHERE (sSerie='" & serie & "' AND cDocumentoCompra='" & Numero & "' AND cPersona='" & Persona & "' AND doc_cod='" & Trim(doc_cod) & "')"
        Call ConfiguraRst(strCadena)
        Me.TxtSerie.Text = rst("sSerie")
        Me.TxtFactura.Text = rst("cDocumentoCompra")
        Me.TxtFechaEmision.Text = str(CVDate(rst("dEmisionCompra")))
        Me.DtpFechaPago.Value = CVDate(rst("dVencimiento"))
        Me.TxtCodEntidad.Text = rst("cPersona")
        Me.TxtEntidad.Text = rst("Persona")
        'Me.DtcInterno.BoundText = doc_cod
        cod_referencia = rst("IdReferencia")
        Me.TxtDeuda.Text = Format(rst("saldo"), "#,##0.00")
        Me.TxtSaldo.Text = Format(rst("saldo"), "#,##0.00")
        Me.TxtDeuda.Locked = True
        Set rst = Nothing
        
        strCadena = "SELECT * FROM Docreferencia_Compra WHERE IdReferencia='" & Trim(cod_referencia) & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            doc_ref = rst("doc_cod")
            Me.txtSerieRef.Text = rst("sSerie")
            Me.TxtNumeroRef.Text = rst("cDocumentoCompra")
        End If
End If

    strCadena = "SELECT doc_cod as Codigo, doc_abrev as Descripcion FROM Comprobantes " & _
  " WHERE doc_tienda='V' ORDER BY doc_abrev"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcInterno)
  Me.DtcInterno.BoundText = doc_cod
  Set rst = Nothing
  
  strCadena = "SELECT doc_cod as Codigo, doc_abrev as Descripcion FROM Comprobantes " & _
  " WHERE doc_tienda='V' ORDER BY doc_abrev"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcExterno)
  Me.DtcExterno.BoundText = doc_ref
  Set rst = Nothing
End Sub
Private Sub LLENA()

If FrmListadoFacturasCompra.Procedencia = nuevo Then
        Dim serie As String
        Dim Numero As String
        Dim Persona As String
        Dim doc_cod As String
        Dim cod_referencia As String
        FrmListadoFacturasCompra.HfgFacturas.col = 0
        serie = Trim(FrmListadoFacturasCompra.HfgFacturas.Text)
        FrmListadoFacturasCompra.HfgFacturas.col = 1
        Numero = Trim(FrmListadoFacturasCompra.HfgFacturas.Text)
        FrmListadoFacturasCompra.HfgFacturas.col = 2
        Persona = Trim(FrmListadoFacturasCompra.HfgFacturas.Text)
        FrmListadoFacturasCompra.HfgFacturas.col = 3
        doc_cod = Trim(FrmListadoFacturasCompra.HfgFacturas.Text)
        
        strCadena = "SELECT * FROM DocumentoCompra WHERE (sSerie='" & serie & "' AND cDocumentoCompra='" & Numero & "' AND cPersona='" & Persona & "' AND doc_cod='" & Trim(doc_cod) & "')"
        Call ConfiguraRst(strCadena)
        Me.TxtSerie.Text = rst("sSerie")
        Me.TxtFactura.Text = rst("cDocumentoCompra")
        Me.TxtFechaEmision.Text = str(CVDate(rst("dEmisionCompra")))
        Me.DtpFechaPago.Value = CVDate(rst("dVencimiento"))
        Me.TxtCodEntidad.Text = rst("cPersona")
        Me.TxtEntidad.Text = rst("Persona")
        'Me.DtcInterno.BoundText = doc_cod
        cod_referencia = rst("IdReferencia")
        Me.TxtDeuda.Text = Format(rst("saldo"), "#,##0.00")
        Me.TxtSaldo.Text = Format(rst("saldo"), "#,##0.00")
        Me.TxtDeuda.Locked = True
        Set rst = Nothing
        
        strCadena = "SELECT * FROM Docreferencia_Compra WHERE IdReferencia='" & Trim(cod_referencia) & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            Me.DtcExterno.BoundText = rst("doc_cod")
            Me.txtSerieRef.Text = rst("sSerie")
            Me.TxtNumeroRef.Text = rst("cDocumentoCompra")
        End If
End If
End Sub
Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case KEY_SAVE
      Call Save
    'FrmListadoFacturasCompra.Facturas
    Case KEY_CANCEL
        Unload Me
  End Select
End Sub

Private Sub TxtDeuda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Me.TxtDeuda.Text = Format(Me.TxtDeuda.Text, "#,##0.00")
        Me.txtMonto.SetFocus
End If
End Sub

Private Sub TxtMonto_Change()
Dim Saldo As Double
Dim Monto As Double
Dim resto As Double
Saldo = Me.TxtDeuda.Text
 If Val(Me.txtMonto.Text) > 0 Then
    
    Monto = Me.txtMonto.Text
    resto = Saldo - Monto
    Me.TxtSaldo.Text = Format(resto, "#,##0.00")
Else
    Me.TxtSaldo.Text = Format(Saldo, "#,##0.00")
 End If
End Sub


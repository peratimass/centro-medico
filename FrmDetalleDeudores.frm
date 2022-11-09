VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmDetalleDeudores 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13230
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13230
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Imprimir Detalle"
      Height          =   375
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Deudas Corporativas"
      Height          =   1095
      Left            =   11760
      TabIndex        =   31
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cancelar  Deuda"
      Height          =   375
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      MaxLength       =   80
      TabIndex        =   6
      Top             =   1995
      Width           =   1575
   End
   Begin VB.TextBox TxtDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      MaxLength       =   80
      TabIndex        =   5
      Top             =   1620
      Width           =   3855
   End
   Begin VB.TextBox TxtCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      MaxLength       =   80
      TabIndex        =   4
      Top             =   1260
      Width           =   3855
   End
   Begin VB.TextBox TxtCodCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   840
      MaxLength       =   80
      TabIndex        =   3
      Top             =   1260
      Width           =   735
   End
   Begin VB.CheckBox ChkAlmacen 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Almacen"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   510
      Width           =   1095
   End
   Begin VB.TextBox TxtSaldo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4080
      MaxLength       =   80
      TabIndex        =   1
      Top             =   1995
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   1095
      Left            =   11760
      TabIndex        =   0
      Top             =   7200
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   2325
      TabIndex        =   7
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   -2147483635
      Text            =   "DataCombo1"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgSalidas 
      Height          =   1335
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   2355
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MsPagos 
      Height          =   2175
      Left            =   120
      TabIndex        =   18
      Top             =   6240
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDeudas 
      Height          =   2895
      Left            =   5760
      TabIndex        =   19
      Top             =   480
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5106
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin VB.Label lbltotaldeuda 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   9945
      TabIndex        =   29
      Top             =   3480
      Width           =   75
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Deuda:"
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
      Left            =   8085
      TabIndex        =   28
      Top             =   3480
      Width           =   1245
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Pago"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   300
      TabIndex        =   27
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comprobante Pago"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2205
      TabIndex        =   26
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monto Pagado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5535
      TabIndex        =   25
      Top             =   6000
      Width           =   1395
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   9705
      TabIndex        =   24
      Top             =   3960
      Width           =   525
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   7410
      TabIndex        =   23
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comprobante"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4500
      TabIndex        =   22
      Top             =   3960
      Width           =   1275
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   1920
      TabIndex        =   21
      Top             =   3960
      Width           =   1155
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emision"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   345
      TabIndex        =   20
      Top             =   3960
      Width           =   705
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Top             =   6000
      Width           =   11175
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   120
      Top             =   3960
      Width           =   11175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   11835
      TabIndex        =   17
      Top             =   240
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8250
      TabIndex        =   16
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5820
      TabIndex        =   15
      Top             =   240
      Width           =   675
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Left            =   5760
      Top             =   240
      Width           =   7200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cuotas Canceladas:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   4005
      TabIndex        =   13
      Top             =   5760
      Width           =   2205
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Documentos Pendientes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   3735
      TabIndex        =   12
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      FillColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   840
      Top             =   315
      Width           =   4695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos Persona:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   855
      TabIndex        =   11
      Top             =   960
      Width           =   1425
   End
   Begin VB.Label LblIdentificacion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruc:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   885
      TabIndex        =   10
      Top             =   1980
      Width           =   555
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   720
      TabIndex        =   9
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   3435
      TabIndex        =   8
      Top             =   2040
      Width           =   645
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   3255
      Left            =   120
      Top             =   120
      Width           =   12975
   End
End
Attribute VB_Name = "FrmDetalleDeudores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Dim RstSalidas As New ADODB.Recordset
Dim RstDeudores As New ADODB.Recordset
Dim TotalAdelanto As Double
Dim TotalSalidas As Double




Private Sub ChkSalidaDinero_Click()
MostrarDocumentos
End Sub

Sub MostrarDocumentos()
If Me.HfgDeudas.Rows > 0 Then
strCadena = "SELECT DocumentoVenta.dEmisionVenta as EMISION, DocumentoVenta.dVencimiento AS VENCIMIENTO, (Comprobantes.doc_abrev + ':' + DocumentoVenta.sSerie +'-'+ " & _
"DocumentoVenta.cDocumentoVenta) AS COMPROBANTE , DocumentoVenta.nTotalVenta AS TOTAL, DocumentoVenta.Saldo AS SALDO,DocumentoVenta.doc_cod,DocumentoVenta.sSerie,DocumentoVenta.cDocumentoVenta,DocumentoVenta.Alm_cod,DocumentoVenta.id_documentoventa FROM DocumentoVenta INNER JOIN " & _
"Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod  WHERE " & _
            "DocumentoVenta.idFormaPago='" & Trim(KEY_CREDITO) & "' AND DocumentoVenta.anulado='F' AND DocumentoVenta.cPersona='" & Trim(Me.HfgDeudas.TextMatrix(Me.HfgDeudas.Row, 0)) & "' AND DocumentoVenta.saldo>0 ORDER BY 1 ASC"
            Call LlenarGrillaSalida(strCadena)
End If

End Sub


Private Sub LlenarGrillaSalida(ByVal Cadena As String)
    Set RstSalidas = Nothing
    RstSalidas.Open Cadena, CnBd, adOpenKeyset, adLockOptimistic
         If RstSalidas.RecordCount > 0 Then
        
        
        Me.HfgSalidas.Clear
        Me.HfgSalidas.Rows = 1
        Set Me.HfgSalidas.Recordset = RstSalidas
        Me.HfgSalidas.Rows = RstSalidas.RecordCount
        HfgSalidas.ColWidth(0) = 1500
        HfgSalidas.ColWidth(1) = 1500
        HfgSalidas.ColWidth(2) = 3500
        HfgSalidas.ColWidth(3) = 2000
        HfgSalidas.ColWidth(4) = 2000
        HfgSalidas.ColWidth(5) = 0
        HfgSalidas.ColWidth(6) = 0
        HfgSalidas.ColWidth(7) = 0
        HfgSalidas.ColWidth(8) = 0
        HfgSalidas.ColWidth(9) = 0
        Call DarFormatoFecha(HfgSalidas, 0)
        Call DarFormatoFecha(HfgSalidas, 1)
        
        RstSalidas.MoveFirst
        
        
     End If
Set RstSalidas = Nothing
End Sub
Private Sub LlenarGrillaDeudores(ByVal Cadena As String)
    Set RstDeudores = Nothing
    Dim registros As Integer
    strCadena = "SELECT * FROM DocumentoVenta where (DocumentoVenta.idFormaPago ='" & KEY_CREDITO & "') AND (DocumentoVenta.Anulado = 'F')  AND estado='Pendiente' AND (DocumentoVenta.Saldo <> 0)"
    Call ConfiguraRst(strCadena)
    registros = rst.RecordCount
    Set rst = Nothing
        
        RstDeudores.Open Cadena, CnBd, adOpenKeyset, adLockOptimistic
        Me.HfgDeudas.Clear
        Me.HfgDeudas.Rows = 1
        Set Me.HfgDeudas.Recordset = RstDeudores
       ' Me.HfgDeudas.Rows = RstDeudores.RecordCount + 1
        HfgDeudas.ColWidth(0) = 1000
        HfgDeudas.ColWidth(1) = 4000
        HfgDeudas.ColWidth(2) = 1800
       
        Set RstDeudores = Nothing
End Sub



Private Sub CmdImprimir_Click()
strCadena = "SELECT     Detalle_DocumentoVenta.cProducto, Producto.DescripcionProducto, Unidad.sAbreviatura, Detalle_DocumentoVenta.cantidad," & _
"                      Detalle_DocumentoVenta.Precio, Detalle_DocumentoVenta.Total, DocumentoVenta.Saldo, DocumentoVenta.cPersona," & _
"                      DocumentoVenta.Persona, Detalle_DocumentoVenta.doc_cod, Detalle_DocumentoVenta.sSerie," & _
"                      Detalle_DocumentoVenta.cDocumentoVenta , Comprobantes.doc_abrev " & _
"FROM         DocumentoVenta INNER JOIN " & _
"                      Detalle_DocumentoVenta ON DocumentoVenta.id_documentoventa = Detalle_DocumentoVenta.id_documentoventa AND " & _
"                      DocumentoVenta.cDocumentoVenta = Detalle_DocumentoVenta.cDocumentoVenta AND " & _
"                      DocumentoVenta.doc_cod = Detalle_DocumentoVenta.doc_cod AND " & _
"                      DocumentoVenta.Alm_cod = Detalle_DocumentoVenta.Alm_Cod AND " & _
"                      DocumentoVenta.sSerie = Detalle_DocumentoVenta.sSerie INNER JOIN " & _
"                      Producto ON Detalle_DocumentoVenta.cProducto = Producto.cProducto INNER JOIN " & _
"                      Unidad ON Producto.cUnidad = Unidad.cUnidad INNER JOIN " & _
"                      Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod WHERE DocumentoVenta.cPersona='" & Trim(Me.HfgDeudas.TextMatrix(Me.HfgDeudas.Row, 0)) & "' AND DocumentoVenta.saldo>0 AND anulado='F'"

Call ConfiguraRst(strCadena)
     Ans = ShowMultiReport(rst, "RptDetalleDeuda", , App.Path + "\Reportes\")
      Set rst = Nothing

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Command1_Click()
 Procedencia = Nuevo
       FrmPagoCuotaDeuda.Show
       FrmPagoCuotaDeuda.TxtMonto.SetFocus
End Sub

Private Sub Command2_Click()
If Val(Me.HfgDeudas.Rows) > 0 Then
      Procedencia = Nuevo
       FrmPagoCuotaDeuda.Show
       FrmPagoCuotaDeuda.TxtMonto.SetFocus
End If
End Sub

Private Sub Command3_Click()
FrmDetalleDeudoresEmpresa.Show
End Sub

Private Sub DtcAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtCodCliente)
End If
End Sub



Private Sub Form_Activate()
CenterForm Me
Me.Top = 200
Me.DtcAlmacen.SetFocus

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub
Sub cargar_deudores()
strCadena = "SELECT SUM(DocumentoVenta.Saldo) FROM DocumentoVenta WHERE (DocumentoVenta.idFormaPago ='" & KEY_CREDITO & "') AND (DocumentoVenta.Anulado = 'F')  AND (DocumentoVenta.Saldo <> 0)"
Call ConfiguraRst(strCadena)
    If IsNull(rst(0)) = True Then
        Me.lbltotaldeuda.Caption = 0
    Else
        Me.lbltotaldeuda.Caption = "S/." & Str(rst(0))
    End If
Set rst = Nothing
strCadena = "SELECT DocumentoVenta.cPersona as Codigo, DocumentoVenta.Persona as Cliente, SUM(DocumentoVenta.Saldo) AS Saldo " & _
            "FROM DocumentoVenta INNER JOIN Persona ON DocumentoVenta.cPersona = Persona.cPersona " & _
            "WHERE (DocumentoVenta.idFormaPago ='" & KEY_CREDITO & "') AND (DocumentoVenta.Anulado = 'F')  AND (DocumentoVenta.Saldo <> 0)" & _
            "GROUP BY DocumentoVenta.cPersona,DocumentoVenta.Persona"
            Call LlenarGrillaDeudores(strCadena)
End Sub
Private Sub Form_Load()
CenterForm Me
Me.Top = 200
doc_Tienda = "V"
 strCadena = "SELECT Alm_cod as Codigo, Alm_des as Descripcion FROM Almacen " & _
  " ORDER BY Alm_cod ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.ChkAlmacen.Value = 1
  Set rst = Nothing
  Me.HfgSalidas.Rows = 0
  Me.HfgDeudas.Rows = 0
  
  Dim Anulado As String
Call cargar_deudores
    
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Error
  Select Case Button.Key
    Case KEY_NEW
      'Call Nuevo
    Case KEY_DELETE
       If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Procedencia = Eliminar
        FrmSeguridad.Show
       End If
    Case KEY_EXIT
        Unload Me
'Error:
 ' MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  
  Exit Sub
End Select
End Sub


Private Sub Frame1_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub HfgDeudas_Click()

  If Me.HfgDeudas.TextMatrix(Me.HfgDeudas.Row, 1) = "" Then
    Set rst = Nothing
    Exit Sub
  End If
  Set rst = Nothing
  
  strCadena = "SELECT * FROM Persona WHERE cPersona='" & Trim(Me.HfgDeudas.TextMatrix(Me.HfgDeudas.Row, 0)) & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
    Me.TxtCodCliente.Text = rst("cPersona")
    Me.TxtCliente.Text = rst("NombrePersona")
    Me.TxtDireccion.Text = rst("sDireccionCliente1")
    Me.TxtRuc.Text = rst("Per_Ruc")
    Me.TxtSaldo.Text = Format(Me.HfgDeudas.TextMatrix(Me.HfgDeudas.Row, 2), "#,##0.00")
    
    Set rst = Nothing
    Call MostrarDocumentos
    Call llenar_pagos
End If
    
  
End Sub

Private Sub HfgSalidas_Click()
If Me.HfgSalidas.Rows > 0 Then
strCadena = "SELECT     DetallePagoCreditos.FechaPago, (Comprobantes.doc_abrev+':'+ DetallePagoCreditos.Serie+'-'+ DetallePagoCreditos.Numero) as COMPROBANTE, " & _
" DetallePagoCreditos.Monto , Seguridad.Usuario as Recibio FROM         DetallePagoCreditos INNER JOIN " & _
" Comprobantes ON DetallePagoCreditos.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
" Seguridad ON DetallePagoCreditos.id_usuario = Seguridad.IdUsuario " & _
"WHERE cPersona='" & Trim(Me.TxtCodCliente.Text) & "' AND SerieVenta='" & Trim(Me.HfgSalidas.TextMatrix(Me.HfgSalidas.Row, 6)) & "' AND NumeroVenta='" & Trim(Me.HfgSalidas.TextMatrix(Me.HfgSalidas.Row, 7)) & "' AND alm_cod='" & Trim(Me.HfgSalidas.TextMatrix(Me.HfgSalidas.Row, 8)) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   
    Me.MsPagos.Clear
    Me.MsPagos.Rows = 1
    Set Me.MsPagos.Recordset = rst
    Me.MsPagos.Rows = rst.RecordCount
    MsPagos.ColWidth(0) = 1500
    MsPagos.ColWidth(1) = 3000
    MsPagos.ColWidth(2) = 2000
    
     Call DarFormatoFecha(MsPagos, 0)
      Call DarFormatoFecha(MsPagos, 1)
     Set rst = Nothing
Else
    Me.MsPagos.Rows = 0
    Me.MsPagos.Clear
End If

End If

End Sub
Sub llenar_pagos()
If Val(Me.HfgSalidas.Rows) > 0 Then
strCadena = "SELECT DetallePagoCreditos.FechaPago AS FECHA, (Comprobantes.doc_abrev +':'+ DetallePagoCreditos.Serie +'-'+ DetallePagoCreditos.Numero) AS COMPROBANTE, " & _
"DetallePagoCreditos.Monto AS MONTO FROM DetallePagoCreditos INNER JOIN Comprobantes ON DetallePagoCreditos.doc_cod = Comprobantes.doc_cod " & _
"WHERE cPersona='" & Trim(Me.TxtCodCliente.Text) & "'  AND alm_cod='" & Trim(Me.HfgSalidas.TextMatrix(Me.HfgSalidas.Row, 8)) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   
    Me.MsPagos.Clear
    Me.MsPagos.Rows = 1
    Set Me.MsPagos.Recordset = rst
    Me.MsPagos.Rows = rst.RecordCount
    MsPagos.ColWidth(0) = 1500
    MsPagos.ColWidth(1) = 3000
    MsPagos.ColWidth(2) = 2000
    
     Call DarFormatoFecha(MsPagos, 0)
      Call DarFormatoFecha(MsPagos, 1)
     Set rst = Nothing
Else
    Me.MsPagos.Rows = 0
    Me.MsPagos.Clear
End If

End If
End Sub

Private Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub





Private Sub HfgSalidas_DblClick()
If Val(Me.HfgSalidas.Rows) > 0 Then
    Procedencia = buscar
    FrmCOmprobantecredito.Show
End If
End Sub


VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmCOmprobantecredito 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Detalle Credito"
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cerrar"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5400
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   3645
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6429
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorSel    =   16777215
      ForeColorSel    =   8388608
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridLines       =   2
      GridLinesFixed  =   1
      GridLinesUnpopulated=   1
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLineWidthBand=   1
   End
   Begin VB.Label lblNumero 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "RUC:20493899229"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarapoto - San Martin - San Martin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DELIVERY:521559  RPM:#913647"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Jiron Ramon Castilla Nº 155"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Pepe's Autoservicios SAC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label LblVuelto 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label lblMontoPagado 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label LblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VUELTO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   4755
      TabIndex        =   3
      Top             =   6120
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONTO PAGADO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   3990
      TabIndex        =   2
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONTO FACTURA:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Left            =   3960
      TabIndex        =   1
      Top             =   5400
      Width           =   1635
   End
End
Attribute VB_Name = "FrmCOmprobantecredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
Dim X As String
'DocumentoVenta.doc_cod,DocumentoVenta.sSerie,DocumentoVenta.cDocumentoVenta,DocumentoVenta.Alm_cod

If FrmDetalleDeudores.Procedencia = buscar Then
id_tipo_doc = FrmDetalleDeudores.HfgSalidas.TextMatrix(FrmDetalleDeudores.HfgSalidas.Row, 5)
id_serie = FrmDetalleDeudores.HfgSalidas.TextMatrix(FrmDetalleDeudores.HfgSalidas.Row, 6)
id_numero = FrmDetalleDeudores.HfgSalidas.TextMatrix(FrmDetalleDeudores.HfgSalidas.Row, 7)
id_codigo = FrmDetalleDeudores.HfgSalidas.TextMatrix(FrmDetalleDeudores.HfgSalidas.Row, 9)

Me.lblNumero.Caption = "TIKET:" + Trim(id_serie) + "-" + Trim(id_numero)
strCadena = "SELECT nTotalVenta, monto_pagado, monto_vuelto FROM DocumentoVenta WHERE " & _
" id_documentoventa='" & Trim(id_codigo) & "' AND sSerie='" & Trim(id_serie) & "' AND cDocumentoVenta='" & Trim(id_numero) & "'"

Call ConfiguraRst(strCadena)
Me.lblTotal.Caption = Format(rst(0), "#,##0.00")
Me.lblMontoPagado.Caption = Format(rst("monto_pagado"), "#,##0.00")
Me.LblVuelto.Caption = Format(rst("monto_vuelto"), "#,##0.00")
Set rst = Nothing
strCadena = "SELECT   Detalle_DocumentoVenta.cProducto as Codigo, Producto.DescripcionProducto as Descripcion, Unidad.sAbreviatura as UND, Detalle_DocumentoVenta.Precio, " & _
"Detalle_DocumentoVenta.cantidad as Cantidad , Detalle_DocumentoVenta.total as Total FROM DocumentoVenta INNER JOIN Detalle_DocumentoVenta ON " & _
"DocumentoVenta.id_documentoventa = Detalle_DocumentoVenta.id_documentoventa AND DocumentoVenta.cDocumentoVenta = Detalle_DocumentoVenta.cDocumentoVenta " & _
"AND DocumentoVenta.doc_cod = Detalle_DocumentoVenta.doc_cod AND    DocumentoVenta.Alm_cod = Detalle_DocumentoVenta.Alm_Cod AND " & _
"DocumentoVenta.sSerie = Detalle_DocumentoVenta.sSerie INNER JOIN Producto ON Detalle_DocumentoVenta.cProducto = Producto.cProducto INNER JOIN " & _
"Unidad ON Producto.cUnidad = Unidad.cUnidad " & _
" WHERE (DocumentoVenta.id_documentoventa='" & Trim(id_codigo) & "' " & _
" AND DocumentoVenta.cDocumentoVenta='" & Trim(id_numero) & "' " & _
" AND DocumentoVenta.doc_cod='" & Trim(id_tipo_doc) & "' AND DocumentoVenta.sSerie='" & Trim(id_serie) & "')"

Call llenarGridDetalle(Me.HfgDetalle, Me)
FrmDetalleDeudores.Procedencia = Neutro
End If
End Sub
Sub llenarGridDetalle(ByVal Grilla As MSHFlexGrid, ByVal formulario As Form)
On Error GoTo SALIR
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 0
  Grilla.ColWidth(1) = 2500
  Grilla.ColWidth(2) = 700
  Grilla.Refresh
   'Me.HfgDetalle.SetFocus
  Exit Sub
SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub



VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmDetalleAnuladas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   9975
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   2925
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5159
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorSel    =   16777215
      ForeColorSel    =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VUELTO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   6450
      TabIndex        =   6
      Top             =   4110
      Width           =   1245
   End
   Begin VB.Label lblVuelto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   7905
      TabIndex        =   5
      Top             =   4020
      Width           =   1965
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAGO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   6750
      TabIndex        =   4
      Top             =   3630
      Width           =   945
   End
   Begin VB.Label lblPago 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   7905
      TabIndex        =   3
      Top             =   3570
      Width           =   1965
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL A PAGAR:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   330
      Left            =   5280
      TabIndex        =   2
      Top             =   3195
      Width           =   2415
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   7905
      TabIndex        =   1
      Top             =   3120
      Width           =   1965
   End
End
Attribute VB_Name = "FrmDetalleAnuladas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub Form_Load()
CenterForm Me
Dim x As String

'If FrmReporteRecaudacionDiaria.Procedencia = buscar Then
'id_tipo_doc = FrmReporteRecaudacionDiaria.HfgAnuladas.TextMatrix(FrmReporteRecaudacionDiaria.HfgAnuladas.Row, 4)
'id_codigo = FrmReporteRecaudacionDiaria.HfgAnuladas.TextMatrix(FrmReporteRecaudacionDiaria.HfgAnuladas.Row, 5)
'id_serie = FrmReporteRecaudacionDiaria.HfgAnuladas.TextMatrix(FrmReporteRecaudacionDiaria.HfgAnuladas.Row, 6)
'id_numero = FrmReporteRecaudacionDiaria.HfgAnuladas.TextMatrix(FrmReporteRecaudacionDiaria.HfgAnuladas.Row, 7)

'strCadena = "SELECT nTotalVenta, monto_pagado, monto_vuelto FROM DocumentoVenta WHERE " & _
" id_documentoventa='" & Trim(id_codigo) & "' AND sSerie='" & Trim(id_serie) & "' AND cDocumentoVenta='" & Trim(id_numero) & "'"

'Call ConfiguraRst(strCadena)
'Me.lblTotal.Caption = Format(rst(0), "#,##0.00")
'Me.lblPago.Caption = Format(rst(1), "#,##0.00")
'Me.lblVuelto.Caption = Format(rst(2), "#,##0.00")
'Set rst = Nothing
'strCadena = "SELECT   Detalle_DocumentoVenta.cProducto as Codigo, Producto.DescripcionProducto as Descripcion, Unidad.sAbreviatura as UND, Detalle_DocumentoVenta.Precio, " & _
"Detalle_DocumentoVenta.cantidad as Cantidad , Detalle_DocumentoVenta.total as Total FROM DocumentoVenta INNER JOIN Detalle_DocumentoVenta ON " & _
"DocumentoVenta.id_documentoventa = Detalle_DocumentoVenta.id_documentoventa AND DocumentoVenta.cDocumentoVenta = Detalle_DocumentoVenta.cDocumentoVenta " & _
"AND DocumentoVenta.doc_cod = Detalle_DocumentoVenta.doc_cod AND    DocumentoVenta.Alm_cod = Detalle_DocumentoVenta.Alm_Cod AND " & _
"DocumentoVenta.sSerie = Detalle_DocumentoVenta.sSerie INNER JOIN Producto ON Detalle_DocumentoVenta.cProducto = Producto.cProducto INNER JOIN " & _
"Unidad ON Producto.cUnidad = Unidad.cUnidad " & _
" WHERE (DocumentoVenta.id_documentoventa='" & Trim(id_codigo) & "' " & _
" AND DocumentoVenta.cDocumentoVenta='" & Trim(id_numero) & "' " & _
" AND DocumentoVenta.doc_cod='" & Trim(id_tipo_doc) & "' AND DocumentoVenta.sSerie='" & Trim(id_serie) & "')"

'Call llenarGridDetalle(Me.HfgDetalle, Me)
'End If
End Sub
Sub llenarGridDetalle(ByVal Grilla As MSHFlexGrid, ByVal formulario As Form)
On Error GoTo SALIR
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.Rows = rst.RecordCount
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 700
  Grilla.ColWidth(1) = 5400
  Grilla.ColWidth(2) = 700
  Grilla.Refresh
   'Me.HfgDetalle.SetFocus
  Exit Sub
SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub



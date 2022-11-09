VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmDesrivados_list 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte Derivados"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   61538305
      CurrentDate     =   40615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDetalle 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8916
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
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   61538305
      CurrentDate     =   40615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4200
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   1080
      Width           =   570
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Left            =   120
      Top             =   1080
      Width           =   11175
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8640
      TabIndex        =   5
      Top             =   1080
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   11175
   End
End
Attribute VB_Name = "FrmDesrivados_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdBuscar_Click()
strCadena = "SELECT     cDerivado, fecha, descripcion,  detalle " & _
"FROM  Derivados WHERE Derivados.fecha>='" & CVDate(Me.DtPInicio.Value) & "' AND Derivados.fecha<='" & Trim(Me.DtpFin.Value) & "' ORDER BY Derivados.cDerivado DESC"
Call llenar_grid
End Sub
Sub llenar_grid()
Me.HfDetalle.Clear
Me.HfDetalle.Rows = 1
Call ConfiguraRst(strCadena)
Set Me.HfDetalle.Recordset = rst
    Me.HfDetalle.Rows = rst.RecordCount
    Me.HfDetalle.ColWidth(0) = 1200
    Me.HfDetalle.ColWidth(1) = 1100
    Me.HfDetalle.ColWidth(2) = 4500
    Me.HfDetalle.ColWidth(3) = 4500
    
End Sub

Private Sub cmSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
Me.DtpFin.Value = Date
Me.DtPInicio.Value = Date
End Sub


Private Sub HfDetalle_DblClick()
If Me.HfDetalle.Rows > 0 Then
   strCadena = "SELECT Derivados.cDerivado, Derivados.cProducto, Derivados.anulado, Derivados.fecha, Derivados.descripcion, Derivados.detalle," & _
        "Derivados.estado, Producto_barras.cod_barra, Producto.DescripcionProducto, Unidad.sAbreviatura, Derivado_Detalle.cantidad, " & _
        " Seguridad.Nombre FROM  Derivado_Detalle INNER JOIN Producto ON Derivado_Detalle.cProducto = Producto.cProducto INNER JOIN " & _
        " Producto_barras ON Producto.cProducto = Producto_barras.cProducto INNER JOIN Unidad ON Producto.cUnidad = Unidad.cUnidad INNER JOIN " & _
        " Derivados ON Derivado_Detalle.cDerivado = Derivados.cDerivado INNER JOIN Seguridad ON Derivados.id_usuario = Seguridad.IdUsuario " & _
        "WHERE Derivados.cDerivado='" & Trim(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) & "'"
         Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptDerivados", , App.Path + "\Reportes\")
End If
End Sub





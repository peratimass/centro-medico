VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmCombo_list 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Listado Combos"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   300
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   148570113
      CurrentDate     =   40615
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   300
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   148570113
      CurrentDate     =   40615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDetalle 
      Height          =   6375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11245
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   11775
   End
End
Attribute VB_Name = "FrmCombo_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdbuscar_Click()
Dim fecha_ini As String
Dim fecha_fin As String
fecha_ini = Format(Me.DtpInicio.Value, Format("yyyy-mm-dd"))
fecha_fin = Format(Me.DtpFin.Value, Format("yyyy-mm-dd"))
strCadena = "SELECT C.id_combo,P.nombre_prod,U.abreviatura,C.cantidad,PE.nombre_completo,C.fecha FROM combo C,producto P,persona PE,unidad U WHERE C.id_producto=P.id_producto AND P.id_unidad=U.id_und AND C.id_usuario=PE.dni AND C.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND  C.fecha>='" & fecha_ini & "' AND C.fecha<='" & fecha_fin & "' ORDER BY C.fecha DESC"
Call llenar_grid(Me.HfDetalle)
End Sub
Public Sub llenar_grid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 4000
           Grilla.ColWidth(3) = 1100
           Grilla.ColWidth(4) = 1100
           Grilla.ColWidth(5) = 3000
      Next
        
        cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "DESCRIPCION COMBO" & vbTab & "UND" & vbTab & "CANT" & vbTab & "RESPONSABLE"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_combo") & vbTab & rst("fecha") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & rst("nombre_completo")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
    Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
  
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub


Private Sub cmSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
Me.DtpFin.Value = KEY_FECHA
Me.DtpInicio.Value = KEY_FECHA
Call actualizar
End Sub
Public Sub actualizar()
strCadena = "SELECT C.id_combo,P.nombre_prod,U.abreviatura,C.cantidad,PE.nombre_completo,C.fecha FROM combo C,producto P,persona PE,unidad U WHERE C.id_producto=P.id_producto AND P.id_unidad=U.id_und AND C.id_usuario=PE.dni AND C.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' "
Call llenar_grid(Me.HfDetalle)
End Sub

Private Sub HfDetalle_DblClick()
If Me.HfDetalle.Rows > 0 Then
   strCadena = "SELECT     Combo.cCombo, Combo.fecha, Combo.descripcion, Combo.cantidad, Combo.detalle, Combo.anulado, " & _
        "Producto_barras.cod_barra, Producto.DescripcionProducto, Unidad.sAbreviatura, Combo_detalle.cantidad AS Expr1," & _
        " Seguridad.Nombre FROM         Combo INNER JOIN " & _
        " Combo_detalle ON Combo.cCombo = Combo_detalle.cCombo INNER JOIN " & _
        " Producto ON Combo_detalle.cProducto = Producto.cProducto INNER JOIN " & _
        " Producto_barras ON Producto.cProducto = Producto_barras.cProducto INNER JOIN " & _
        " Unidad ON Producto.cUnidad = Unidad.cUnidad INNER JOIN " & _
        " Seguridad ON Combo.id_usuario = Seguridad.IdUsuario WHERE Combo.cCombo='" & Trim(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) & "'"
         Call ConfiguraRst(strCadena)
         
    strCadena = "SELECT C.id_combo,C.fecha,P.nombre_prod,C.cantidad,C.detalle,C.anulado FROM combo C,producto  WHERE P.id_producto=C.id_producto AND C.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND C.id_combo='" & Trim(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) & "'"
        Ans = ShowMultiReport(rst, "RptCombo", , App.Path + "\Reportes\")
End If
End Sub




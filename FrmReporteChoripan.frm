VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmReporteMargenBruto 
   BorderStyle     =   0  'None
   ClientHeight    =   8685
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   18315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   18315
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn cmdGenerarReporte 
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   5880
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "MARGEN BRUTO GENERAL  "
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteChoripan.frx":0000
      PICN            =   "FrmReporteChoripan.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   8415
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   14843
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
   Begin MSComCtl2.DTPicker DtpIni 
      Height          =   405
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   146931713
      CurrentDate     =   43227
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   405
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   146931713
      CurrentDate     =   43227
   End
   Begin MSDataListLib.DataCombo DtcLinea 
      Height          =   330
      Left            =   1560
      TabIndex        =   4
      Top             =   2280
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VitekeySoft.ChameleonBtn cmdgenerarPantalla 
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   3240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "GENERAR EN PANTALLA"
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
      BCOL            =   33023
      BCOLO           =   33023
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteChoripan.frx":25ED
      PICN            =   "FrmReporteChoripan.frx":2609
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcSublinea 
      Height          =   330
      Left            =   1560
      TabIndex        =   9
      Top             =   2760
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VitekeySoft.ChameleonBtn cmdImprimir 
      Height          =   615
      Left            =   1560
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "GENERAR LINEA SUBLINEA"
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteChoripan.frx":4BEE
      PICN            =   "FrmReporteChoripan.frx":4C0A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar progress_costo 
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   3960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUB-LINEA :"
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
      Left            =   600
      TabIndex        =   11
      Top             =   2760
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LINEA :"
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
      Left            =   915
      TabIndex        =   10
      Top             =   2400
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   18000
      Picture         =   "FrmReporteChoripan.frx":7E0E
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RESUMEN POR LINEA DE PRODUCTOS Y MARGEN BRUTO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA INICIO :"
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
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA FINAL:"
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
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8415
      Left            =   120
      Top             =   120
      Width           =   5535
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8685
      Left            =   0
      Top             =   0
      Width           =   18315
   End
End
Attribute VB_Name = "frmReporteMargenBruto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdgenerarPantalla_Click()


strCadena = "SELECT funct_get_margen('" & Format(Me.DtpIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & KEY_ALM & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
in_acumuldo = rst(0)

strCadena = "SELECT '" & in_acumuldo & "' ,fecha_emision,id_doc,id_producto,nombre_prod,id_linea,linea,id_sublinea,sublinea,sum(cantidad) as cantidad,sum(costo) as costo,sum(venta) as venta,id_alm,ruc " & _
" From view_resumen_margen_bruto where fecha_emision>='" & Format(Me.DtpIni.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and id_linea='" & Me.DtcLinea.BoundText & "' and id_sublinea='" & Me.DtcSubLinea.BoundText & "'  and ruc='" & KEY_RUC & "' GROUP BY id_producto"
Call llenarGrid(Me.HfdGrilla)

End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
'On Error GoTo salir
Dim in_cantidad As Double
Dim in_venta As Double
Dim in_costo As Double
Dim in_rentabilidad As Single
Dim in_margen As Single
Dim in_participacionn As Single
Dim in_margenn As Single
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 1500
           Grilla.ColWidth(3) = 1500
           Grilla.ColWidth(4) = 1500
           Grilla.ColWidth(5) = 1500
           Grilla.ColWidth(6) = 1500
        
          Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "CANTIDAD" & vbTab & "VALOR VENTA" & vbTab & "PARTICIPACION" & vbTab & "VALOR COSTO" & vbTab & "MARGEN"
         Grilla.AddItem cabecera
         For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        Me.cmdImprimir.Visible = True
        in_margen = 0
        in_participacionn = 0
        For i = 0 To rst.RecordCount - 1
             in_division_v = rst("venta")
             in_division_c = rst("costo")
             If rst("venta") <= 0 Then
                in_division_v = 1
             End If
             If rst("costo") <= 0 Then
                in_division_c = 1
             End If
             
             in_participacion = Format(rst("venta") * 100 / rst(0), "#,##0.00")
             in_margen = (1 - (in_division_c / in_division_v)) * 100
             Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & Format(rst("cantidad"), "#,##.00") & vbTab & Format(rst("venta"), "#,##0.00") & vbTab & Format(in_participacion, "#,##0.0000") & vbTab & Format(rst("costo"), "#,##0.00") & vbTab & Format(in_margen, "#,##0.00")
             Grilla.AddItem Fila
             in_cantidad = in_cantidad + rst("cantidad")
             in_venta = in_venta + rst("venta")
             in_costo = in_costo + rst("costo")
             in_participacionn = in_participacionn + in_participacion
             in_margenn = in_margenn + in_margen
             
        rst.MoveNext
        Next i
        
        Fila = "" & vbTab & "" & vbTab & Format(in_cantidad, "#,##0.00") & vbTab & Format(in_venta, "#,##0.00") & vbTab & Format(in_participacionn, "#,##0.00") & vbTab & Format(in_costo, "#,##0.00") & vbTab & Format(in_margenn, "#,##0.00")
        Grilla.AddItem Fila
        For k = 2 To 6
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
        
        
        
Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub cmdgenerarreporte_Click()
Dim in_acumuldo As Double
Dim param As Variant
Dim arr(0 To 1, 1 To 2) As String
arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"
arr(0, 2) = Format(Me.DtpIni.Value, "dd-mm-YYYY")
arr(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
param = arr()

'strCadena = "SELECT funct_get_margen('" & Format(Me.DtpIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & KEY_ALM & "','" & KEY_RUC & "')"
'Call ConfiguraRst(strCadena)
'in_acumuldo = rst(0)


strCadena = "SELECT funct_get_margen('" & Format(Me.DtpIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & KEY_ALM & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
in_acumuldo = rst(0)



strCadena = "call ADM_MargenBrutoResumen('2','" & Me.DtcLinea.BoundText & "','" & Me.DtcSubLinea.BoundText & "','" & Format(Me.DtpIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & KEY_ALM & "','" & in_acumuldo & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)


Ans = ShowMultiReport(rst, "rpt_resumen_margen_bruto", param, App.Path + "\Reportes\")


End Sub

Private Sub cmdImprimir_Click()
Dim in_acumuldo As Double
Dim param As Variant
Dim arr(0 To 1, 1 To 2) As String
arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"
arr(0, 2) = Format(Me.DtpIni.Value, "dd-mm-YYYY")
arr(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
param = arr()

'strCadena = "SELECT funct_get_margen('" & Format(Me.DtpIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & KEY_ALM & "','" & KEY_RUC & "')"
'Call ConfiguraRst(strCadena)
'in_acumuldo = rst(0)


strCadena = "SELECT funct_get_margen('" & Format(Me.DtpIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & KEY_ALM & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
in_acumuldo = rst(0)

'strCadena = "SELECT '" & in_acumuldo & "' ,fecha_emision,id_doc,id_producto,nombre_prod,id_linea,linea,id_sublinea,sublinea,sum(cantidad) as cantidad,sum(costo) as costo,sum(venta) as venta,id_alm,ruc " & _
" From view_resumen_margen_bruto where fecha_emision>='" & Format(Me.DtpIni.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and id_linea='" & Me.DtcLinea.BoundText & "' and id_sublinea='" & Me.DtcSubLinea.BoundText & "'  and ruc='" & KEY_RUC & "' GROUP BY id_producto"
'Call llenarGrid(Me.HfdGrilla)



'Call ConfiguraRst(strCadena)

strCadena = "call ADM_MargenBrutoResumen('1','" & Me.DtcLinea.BoundText & "','" & Me.DtcSubLinea.BoundText & "','" & Format(Me.DtpIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & KEY_ALM & "','" & in_acumuldo & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)


Ans = ShowMultiReport(rst, "rpt_resumen_margen_bruto_det", param, App.Path + "\Reportes\")
End Sub

Private Sub DtcLinea_Change()
strCadena = "SELECT id_tipo as Codigo, descripcion as Descripcion FROM linea_sub WHERE id_usu='" & KEY_RUC & "' AND id_linea='" & Me.DtcLinea.BoundText & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtcSubLinea)
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
Me.DtpIni.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA

strCadena = "SELECT id_linea as Codigo,descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcLinea)



End Sub

Private Sub Image1_Click()
Unload Me
End Sub

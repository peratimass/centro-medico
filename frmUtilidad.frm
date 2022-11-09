VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmUtilidad 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DtcPeriodo 
      Height          =   405
      Left            =   5040
      TabIndex        =   8
      Top             =   840
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   12
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VitekeySoft.ChameleonBtn cmdInforme 
      Height          =   495
      Left            =   15120
      TabIndex        =   5
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "GENERAR INFORME"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   9.75
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUtilidad.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdLoad 
      Height          =   600
      Left            =   4800
      TabIndex        =   6
      Top             =   1680
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1058
      BTYPE           =   3
      TX              =   "CARGAR INFORMACION"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   11.25
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUtilidad.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfIngresos 
      Height          =   5415
      Left            =   210
      TabIndex        =   10
      Top             =   3480
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   9551
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorSel    =   8388608
      BackColorBkg    =   16777215
      GridColor       =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfEgresos 
      Height          =   5415
      Left            =   5175
      TabIndex        =   11
      Top             =   3480
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   9551
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorSel    =   8388608
      BackColorBkg    =   16777215
      GridColor       =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfImpuestos 
      Height          =   5415
      Left            =   10185
      TabIndex        =   12
      Top             =   3480
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   9551
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorSel    =   8388608
      BackColorBkg    =   16777215
      GridColor       =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCuentas 
      Height          =   5415
      Left            =   15195
      TabIndex        =   13
      Top             =   3480
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   9551
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorSel    =   8388608
      BackColorBkg    =   16777215
      GridColor       =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VitekeySoft.ChameleonBtn cmdMiscuentas 
      Height          =   600
      Left            =   19180
      TabIndex        =   18
      Top             =   2640
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1058
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   11.25
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUtilidad.frx":0038
      PICN            =   "frmUtilidad.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTA: NO SE CONSIDERA LAS CUENTAS POR COBRAR"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   9720
      TabIndex        =   17
      Top             =   1320
      Width           =   3705
   End
   Begin VB.Label lblSaldoActual 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   12120
      TabIndex        =   16
      Top             =   840
      Width           =   1755
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SALDO ACTUAL :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   9780
      TabIndex        =   15
      Top             =   720
      Width           =   1785
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   9480
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUENTAS BANCARIAS "
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   16170
      TabIndex        =   14
      Top             =   2760
      Width           =   2445
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   6495
      Left            =   15105
      Top             =   2520
      Width           =   4845
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   1455
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODO ACTUAL"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   5175
      TabIndex        =   9
      Top             =   240
      Width           =   1875
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SALDO ANTERIOR"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   19680
      Picture         =   "frmUtilidad.frx":036E
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IMPUESTOS / COMISIONES"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   11085
      TabIndex        =   4
      Top             =   2760
      Width           =   2865
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   6495
      Left            =   10095
      Top             =   2520
      Width           =   4845
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EGRESOS MENSUALES"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   6165
      TabIndex        =   3
      Top             =   2760
      Width           =   2475
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   6495
      Left            =   5100
      Top             =   2520
      Width           =   4845
   End
   Begin VB.Label lblsaldo_anterior 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   2655
      TabIndex        =   2
      Top             =   960
      Width           =   1755
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SALDO ANTERIOR"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   450
      TabIndex        =   1
      Top             =   840
      Width           =   1905
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INGRESOS MENSUALES"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   1200
      TabIndex        =   0
      Top             =   2760
      Width           =   2565
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   6495
      Left            =   120
      Top             =   2520
      Width           =   4845
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "frmUtilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChameleonBtn1_Click()

End Sub

Private Sub cmdInforme_Click()
Dim cam3(0 To 5, 1 To 5)  As String
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(3, 1) = "empresa"
    cam3(4, 1) = "direccion"
    cam3(5, 1) = "titulo"
    
    strCadena = "SELECT CURDATE()"
    Call ConfiguraRstA(strCadena)
    
    strCadena = "SELECT * FROM con_periodo WHERE id='" & Me.DtcPeriodo.BoundText & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    
    cam3(0, 2) = Format(rst("FechaInicio"), "dd-mm-YYYY")
    cam3(1, 2) = Format(rst("FechaFin"), "dd-mm-YYYY")
    cam3(2, 2) = Format(rstA(0), "dd-mm-YYYY")
    cam3(3, 2) = KEY_EMPRESA
    cam3(4, 2) = KEY_DIRECCION_ALM
    cam3(5, 2) = "SALDO X PERIODO"
    param = cam3()
    
strCadena = "CALL ADM_reportes_generales('54','','','" & KEY_USUARIO & "','0','','','','','" & Me.DtcPeriodo.BoundText & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)

strCadena = "CALL ADM_reportes_generales('52','','','','0','','','','','" & Me.DtcPeriodo.BoundText & "','" & KEY_RUC & "')"
Call ConfiguraRstA(strCadena)

Ans = ShowMultiReport(rst, "RptInformeCuentas", param, App.Path + "\Reportes\", , , , , rstA, "RptInformeCuentas_det")



Exit Sub
End Sub

Private Sub cmdLoad_Click()


strCadena = "SELECT * FROM con_periodo WHERE id='" & Me.DtcPeriodo.BoundText & "' LIMIT 1"
Call ConfiguraRst(strCadena)

Call load_periodo(rst("Ejercicio"), rst("Mes"))

'************** INGRESOS*********************
'--------------- load ventas
strCadena = "CALL ADM_reportes_generales('50','" & Format(rst("FechaInicio"), "YYYY-mm-dd") & "','" & Format(rst("FechaFin"), "YYYY-mm-dd") & "','','0','','','','','" & Me.DtcPeriodo.BoundText & "','" & KEY_RUC & "')"
Call ConfiguraRstA(strCadena)
Me.lblSaldoActual.Caption = Format(Val(Format(Me.lblsaldo_anterior.Caption, "###0.00")) + rstA("_ingresos"), "#,##0.00")

strCadena = "CALL ADM_reportes_generales('51','" & Format(rst("FechaInicio"), "YYYY-mm-dd") & "','" & Format(rst("FechaFin"), "YYYY-mm-dd") & "','','0','','','','02','','" & KEY_RUC & "')"
Call llenarGrid(Me.HfIngresos)

strCadena = "CALL ADM_reportes_generales('51','','','','0','','','','03','','" & KEY_RUC & "')"
Call llenarGrid(Me.HfEgresos)

strCadena = "CALL ADM_reportes_generales('51','','','','0','','','','04','','" & KEY_RUC & "')"
Call llenarGrid(Me.HfImpuestos)

strCadena = "CALL ADM_reportes_generales('52','','','','0','','','','','" & Me.DtcPeriodo.BoundText & "','" & KEY_RUC & "')"
Call llenarCuentas(Me.HfCuentas)






End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
       Grilla.Rows = 0
      
       ReDim arrColWidth(1 To rst.Fields.Count)
       
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 600
           Grilla.ColWidth(1) = 2500
           Grilla.ColWidth(2) = 1500
           
       
        cabecera = "ITEM" & vbTab & "DESCRIPCION" & vbTab & "MONTO"
        Grilla.AddItem cabecera
        For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
                            
        rst.MoveFirst
        in_acumulado = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id") & vbTab & rst("descripcion") & vbTab & Format(rst("valor"), "#,##0.00")
            Grilla.AddItem Fila
            in_acumulado = in_acumulado + rst("valor")
            rst.MoveNext
             
        Next i
        
        cabecera = "" & vbTab & "TOTAL   :" & vbTab & Format(in_acumulado, "#,##0.00")
        Grilla.AddItem cabecera
        For k = 1 To 2
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H80FF80
        Next k
        
    Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub
Private Sub llenarCuentas(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
       Grilla.Rows = 0
      
       ReDim arrColWidth(1 To rst.Fields.Count)
       
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 500
           Grilla.ColWidth(3) = 1100
           
       
        cabecera = "ID" & vbTab & "DESCRIPCION" & vbTab & "TC" & vbTab & "MONTO"
        Grilla.AddItem cabecera
        For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
                            
        rst.MoveFirst
        in_acumulado = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id") & vbTab & rst("descripcion") & vbTab & rst("cambio") & vbTab & Format(rst("valor"), "#,##0.00")
            Grilla.AddItem Fila
            in_acumulado = in_acumulado + rst("valor")
            rst.MoveNext
             
        Next i
        
        cabecera = "" & vbTab & "" & vbTab & "TOTAL   :" & vbTab & Format(in_acumulado, "#,##0.00")
        Grilla.AddItem cabecera
        For k = 1 To 3
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H80FF80
        Next k
        
    Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub

Private Sub cmdMiscuentas_Click()
FrmMiscuentas.Show
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50

strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPeriodo)



End Sub

Public Sub load_periodo(ByVal in_anio As String, ByVal in_mes As String)
in_fecha = Format("01-" + in_mes + "-" + in_anio, "YYYY-mm-dd")

strCadena = "SELECT id, FechaInicio FROM con_periodo WHERE FechaInicio='" & in_fecha & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
    Me.DtcPeriodo.BoundText = rstA("id")
    
    strCadena = "CALL ADM_reportes_generales('39','','" & Format(DateAdd("d", -1, rstA("FechaInicio")), "YYYY-mm-dd") & "','','0','','','','','','" & KEY_RUC & "')"
    Call ConfiguraRstK(strCadena)
    Me.lblsaldo_anterior.Caption = Format(rstK(0), "#,##0.00")
End If

End Sub



Private Sub Image1_Click()
    
    Unload Me
    
End Sub

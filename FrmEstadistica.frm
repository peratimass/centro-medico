VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmEstadistica 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   13365
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn cmdprocesar 
      Height          =   975
      Left            =   11880
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "PROCESAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmEstadistica.frx":0000
      PICN            =   "FrmEstadistica.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox ChkAlmacen 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "SUCURSAL"
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
      Left            =   480
      TabIndex        =   5
      Top             =   945
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   6855
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12091
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
   Begin MSComCtl2.DTPicker DtpDesde 
      Height          =   315
      Left            =   480
      TabIndex        =   6
      Top             =   510
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   -2147483635
      CalendarTitleForeColor=   -2147483635
      Format          =   43057153
      CurrentDate     =   37091
   End
   Begin MSComCtl2.DTPicker DtpHasta 
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Top             =   510
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   -2147483635
      CalendarTitleForeColor=   -2147483635
      Format          =   43057153
      CurrentDate     =   37091
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   330
      Left            =   2085
      TabIndex        =   8
      Top             =   915
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
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
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   975
      Left            =   11880
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "SALIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmEstadistica.frx":2906
      PICN            =   "FrmEstadistica.frx":2922
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RANGO DE FECHAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   420
      TabIndex        =   9
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ACUMULADO VENTAS:"
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
      Height          =   195
      Left            =   8325
      TabIndex        =   3
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL UTILIDAD :"
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
      Height          =   195
      Left            =   8685
      TabIndex        =   2
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label lblTotalVenta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   10080
      TabIndex        =   1
      Top             =   8400
      Width           =   1545
   End
   Begin VB.Label LblTotalUtilidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   10080
      TabIndex        =   0
      Top             =   8760
      Width           =   1545
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   13365
   End
End
Attribute VB_Name = "FrmEstadistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TVentas As Double
Public TUtilidad As Double
Private Sub chkAlmacen_Click()
If Me.ChkAlmacen.Value = 1 Then
    Me.DtcAlmacen.Enabled = True
Else
    Me.DtcAlmacen.Enabled = False
End If
End Sub

Private Sub cmdProcesar_Click()
 Dim in_almacen As String
 in_almacen = ""
 If Me.ChkAlmacen.Value = 1 Then
    in_almacen = Me.DtcAlmacen.BoundText
 End If
    
   
    strCadena = "SELECT * FROM view_utilidad_v2 WHERE afecta_caja='si' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_alm LIKE '%" & in_almacen & "%' and ruc='" & KEY_RUC & "'"
    Call llenarGrid(Me.HfdDetalle, Me)
    
    
    
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100

strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM
Me.DtcAlmacen.Enabled = False
Me.DtpDesde.Value = KEY_FECHA
Me.DtpHasta.Value = KEY_FECHA
TVentas = 0
TUtilidad = 0
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Dim tTotal As Double, utilidad As Double
tTotal = 0
utilidad = 0
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
  

   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 2200
           Grilla.ColWidth(1) = 4400
           Grilla.ColWidth(2) = 1100
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1000
           Grilla.ColWidth(5) = 1200
        Next
         cabecera = "COMPROBANTE" & vbTab & "PRODUCTO" & vbTab & "CANTIDAD" & vbTab & "P.VENTA" & vbTab & "P.COSTO" & vbTab & "UTILIDAD"
         Grilla.AddItem cabecera
         For k = 0 To 5
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            If rst("tipo_movimiento") = "02" Then
                in_utilidad = rst("utilidad") * -1
            Else
                in_utilidad = rst("utilidad")
            End If
             Fila = rst("documento") & vbTab & rst("detalle") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("precio"), "#,##0.00") & vbTab & Format(rst("precio_compra"), "#,##0.00") & vbTab & Format(in_utilidad, "#,##0.00")
             
             If rst("utilidad") <> 0 Then
                utilidad = utilidad + in_utilidad
                If rst("tipo_movimiento") = "02" Then
                   tTotal = tTotal - rst("total")
                Else
                 
                    tTotal = tTotal + rst("total")
            End If
             
             End If
            Grilla.AddItem Fila
            
        rst.MoveNext
        Next i
        Me.lblTotalVenta.Caption = Format(tTotal, "#,##0.00")
        Me.LblTotalUtilidad.Caption = Format(utilidad, "#,##0.00")
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub llenar_grilla()
        
        Call ConfiguraRst(strCadena)
       
        Set Me.HfdDetalle.Recordset = rst
        Me.HfdDetalle.ColWidth(0) = 2300
        Me.HfdDetalle.ColWidth(1) = 4300
        Me.HfdDetalle.ColWidth(2) = 600
        Me.HfdDetalle.ColWidth(3) = 800
        Me.HfdDetalle.ColWidth(4) = 800
        Me.HfdDetalle.ColWidth(5) = 800
       Call DarFormato_t(HfdDetalle, 3)
       Call DarFormato_u(HfdDetalle, 5)
       
        'Call DarFormato(Me.HfdDetalle, 4)
End Sub



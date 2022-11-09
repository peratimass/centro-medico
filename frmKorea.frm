VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmKorea 
   BorderStyle     =   0  'None
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtcomprobante 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   350
      Left            =   4440
      TabIndex        =   16
      Top             =   2670
      Width           =   1695
   End
   Begin VitekeySoft.ChameleonBtn btnRepDosis 
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   2280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "VISUALIZAR REPORTE"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmKorea.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkTrabajador 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1920
      Width           =   255
   End
   Begin VB.CheckBox chkProveedor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox chkProceso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtFechaFin 
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   148766721
      CurrentDate     =   41769
   End
   Begin MSComCtl2.DTPicker dtFechaIni 
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   148766721
      CurrentDate     =   41769
   End
   Begin MSDataListLib.DataCombo cboProceso 
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo cboProveedor 
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo cboTrabajador 
      Height          =   285
      Left            =   2040
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VitekeySoft.ChameleonBtn cmdcerrar 
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   3055
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "CERRAR PANTALLA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmKorea.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   2670
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "GENERAR ORDEN PAGO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmKorea.frx":0038
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
      Caption         =   "TRABAJADOR :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   720
      TabIndex        =   6
      Top             =   1920
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   720
      TabIndex        =   5
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROCESO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   720
      TabIndex        =   4
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA INICIO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA FIN :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   930
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   3495
      Left            =   0
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmKorea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnRepDosis_Click()
   
   
   Dim arr(0 To 5, 1 To 2) As String
   arr(0, 1) = "p_fecha_ini"
   arr(0, 2) = Me.dtFechaIni.Value
   arr(1, 1) = "p_fecha_fin"
   arr(1, 2) = Me.dtFechaFin.Value
   
   
   
  
  Dim clausula As String
  Dim p_trab As String
  Dim p_prov As String
  Dim p_proc As String
  
  
  clausula = ""
  p_trab = ""
  p_prov = ""
  p_proc = ""
  
  If Me.chkProveedor.Value = 1 Then
     clausula = " and pro.dni like '" & Me.cboProveedor.BoundText & "'"
     p_prov = "PROVEEDOR : " & Me.cboProveedor.Text
  End If
  
  If Me.chkTrabajador.Value = 1 Then
     clausula = clausula & " and tra.dni like '" & Me.cboTrabajador.BoundText & "'"
     p_trab = "TRABAJADOR : " & Me.cboTrabajador.Text
  End If
  
  If Me.chkProceso.Value = 1 Then
     clausula = clausula & " and de.id_detalle_estado like '" & Me.cboProceso.BoundText & "'"
     p_proc = "PROCESO : " & Me.cboProceso.Text
  End If
  
    arr(2, 1) = "p_trabajador"
     arr(2, 2) = p_trab
     arr(3, 1) = "p_proceso"
     arr(3, 2) = p_proc
     arr(4, 1) = "p_proveedor"
     arr(4, 2) = p_prov

strCadena = " select pr.`id_producto` as id_prod,  CONCAT (pr.`nombre_prod`, ' ', co.descripcion) as prod, ins.`id_producto`, " & _
" ins.`nombre_prod`, i.`cantidad` , d.`id_detalle`, d.`id_detalle_compra` , " & _
" d.`anio_modelo`, d.`anio_fabricacion`, d.`nro_chasis`, d.`nro_contenedor`, " & _
" d.`nro_motor`, d.`serie`, (i.precio * i.cantidad) as subtotal,i.precio, " & _
" pro.dni as rucPro, pro.`nombre_completo` as proveedor, tra.dni as dniTra, " & _
" tra.`nombre_completo` as trabajador, de.`descripcion` as proceso " & _
" from `imp_producto_detalle` d, " & _
" imp_producto_insumo i, producto ins , producto pr, " & _
"   imp_producto_movimiento m, persona pro , persona tra, " & _
"   imp_estado_detalle de, imp_color co  where " & _
"   de.id_detalle_estado = m.`id_proceso` and " & _
"   ((m.`fecha_entrada` between STR_TO_DATE('" & dtFechaIni.Value & "','%d/%m/%Y') AND STR_TO_DATE('" & dtFechaFin.Value & "','%d/%m/%Y') ) or ( m.fecha_salida between  STR_TO_DATE('" & dtFechaIni.Value & "','%d/%m/%Y') AND STR_TO_DATE('" & dtFechaFin.Value & "','%d/%m/%Y')) ) " & _
"   and m.`id_detalle` = i.`id_producto_detalle` and d.`id_detalle` = i.`id_producto_detalle` " & _
"   and d.`id_producto` = pr.`id_producto` and d.`ruc` = pr.`ruc` and i.`id_producto` = ins.`id_producto` " & _
"   and i.`ruc` = ins.`ruc` and co.id_color = pr.id_color and m.`id_autor` = tra.`dni` and m.`ruc_empresa` = pro.`dni` " & clausula
   
   '" and r.`idestado` = 'PEN' " & _

   Call ConfiguraRst(strCadena)

  
   Dim param As Variant
   param = arr()
  
   Ans = ShowMultiReport(rst, "RptKorOPValorizado", param, App.Path + "\Reportes\")
End Sub

Private Sub ChameleonBtn1_Click()
    
   Dim arr(0 To 5, 1 To 2) As String
   arr(0, 1) = "p_fecha_ini"
   arr(0, 2) = Me.dtFechaIni.Value
   arr(1, 1) = "p_fecha_fin"
   arr(1, 2) = Me.dtFechaFin.Value
  
  If Me.chkProveedor.Value = 1 Then
     clausula = " and pro.dni like '" & Me.cboProveedor.BoundText & "'"
     p_prov = "PROVEEDOR : " & Me.cboProveedor.Text
     
     Else
     
     Exit Sub
  End If
  
  
  
     arr(2, 1) = "p_trabajador"
     arr(2, 2) = p_trab
     arr(3, 1) = "p_proceso"
     arr(3, 2) = p_proc
     arr(4, 1) = "p_proveedor"
     arr(4, 2) = p_prov

    strCadena = " select p.dni, p.nombre_completo from persona p where p.dni = '" & Me.cboProveedor.BoundText & "'"
   
   '" and r.`idestado` = 'PEN' " & _

   Call ConfiguraRst(strCadena)
  
   Dim param As Variant
   param = arr()
  
   Ans = ShowMultiReport(rst, "RptKorOPFirma", param, App.Path + "\Reportes\")

End Sub

Private Sub chkProveedor_Click()
   If chkProveedor.Value = 1 Then
     Me.cboProveedor.Visible = True
     
    Else
     Me.cboProveedor.Visible = False
     
   End If
   
   
End Sub

Private Sub chkProceso_Click()
   If chkProceso.Value = 1 Then
     Me.cboProceso.Visible = True
     
    Else
     Me.cboProceso.Visible = False
     
   End If
   
   
End Sub

Private Sub chkTrabajador_Click()
   If chkTrabajador.Value = 1 Then
     Me.cboTrabajador.Visible = True
     
    Else
     Me.cboTrabajador.Visible = False
     
   End If
End Sub

Private Sub cmdOrden_Click()
  KEY_RUC = "20479779598"
  FrmPedido.Show
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 400
  Me.dtFechaIni.Value = KEY_FECHA
  Me.dtFechaFin.Value = KEY_FECHA
  
  Call llenarProveedores
  Call llenarProcesos
  Call llenarTrabajador
End Sub


Private Sub llenarProveedores()
  
  strCadena = "SELECT E.cod_unico as Codigo,P.nombre_completo as Descripcion   FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni and E.id_empresa='" & KEY_RUC & "' and E.id_almacen='si'"
  'strCadena = " select dni as Codigo, nombre_completo as Descripcion from persona where dni = '20524402387'"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(cboProveedor)
End Sub


Private Sub llenarProcesos()
  strCadena = " select e.`id_detalle_estado` as Codigo, e.`descripcion` as Descripcion from imp_estado_detalle e "
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(cboProceso)
End Sub


Private Sub llenarTrabajador()
  strCadena = " select e.`dni` as Codigo, e.`nombre_completo` as Descripcion from persona e where e.dni in ('00022884') "
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(cboTrabajador)
End Sub


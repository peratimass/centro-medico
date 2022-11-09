VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCorReportes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Reportes"
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "AREA DE PRODUCCION"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   570
      Left            =   600
      Picture         =   "frmCorReportes.frx":0000
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PROCESO :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   600
      Picture         =   "frmCorReportes.frx":25D5
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin MSComCtl2.MonthView dtIni 
      Height          =   2460
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   4339
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   16777215
      BorderStyle     =   1
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
      StartOfWeek     =   127729665
      CurrentDate     =   42318
   End
   Begin MSDataListLib.DataCombo DtcProceso 
      Height          =   330
      Left            =   1920
      TabIndex        =   2
      Top             =   3120
      Width           =   4455
      _ExtentX        =   7858
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
   Begin VitekeySoft.ChameleonBtn cmdReporteEmpleados 
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   3360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "GENERAR REPORTE"
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
      MICON           =   "frmCorReportes.frx":4BAA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdcerrar 
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   3840
      Width           =   2175
      _ExtentX        =   3836
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
      MICON           =   "frmCorReportes.frx":4BC6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.MonthView dtFin 
      Height          =   2460
      Left            =   3480
      TabIndex        =   5
      Top             =   360
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   4339
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   16777215
      BorderStyle     =   1
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
      StartOfWeek     =   127729665
      CurrentDate     =   42318
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   330
      Left            =   1920
      TabIndex        =   7
      Top             =   3600
      Width           =   4455
      _ExtentX        =   7858
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
   Begin VB.Image Image2 
      Height          =   300
      Left            =   120
      Picture         =   "frmCorReportes.frx":4BE2
      Top             =   3600
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   120
      Picture         =   "frmCorReportes.frx":71B7
      Top             =   3120
      Width           =   300
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA FIN."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA INICIO."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmCorReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdReporteEmpleados_Click()
 '  strCadena = "select p.`nombre_completo` as trabajador, pr.`nombre_completo` as empresa, " & _
 "e.`descripcion` as proceso, i.`fecha_salida` , i.`hora_salida`, pr.dni as ruc_service, " & _
 "p.dni as dni_trabajador, o.`nombre_prod` , o.`id_producto`,  " & _
 "de.`serie`, de.`nro_motor`, de.`nro_chasis`, d.`id_compra`, e.precio " & _
 " from `imp_producto_movimiento` i , persona p , imp_estado_detalle e, persona pr, " & _
 " movimiento_compra_detalle d , producto o, `imp_producto_detalle` de " & _
 " where i.`ruc_empresa` = pr.`dni` and i.`id_autor` = p.`dni` " & _
 " and e.`id_detalle_estado` = i.`id_proceso` " & _
 " and de.`id_detalle` = i.`id_detalle` " & _
 " and d.`id_detalle_compra` = de.`id_detalle_compra`  and i.fecha_salida between STR_TO_DATE('" & Me.dtIni.Value & "','%d/%m/%Y') AND STR_TO_DATE('" & Me.dtFin.Value & "','%d/%m/%Y')   " & _
 " and d.`id_producto` = o.`id_producto` " & _
 " and o.ruc = '" & KEY_RUC & "' order by empresa asc, e.id_detalle_estado asc"
 
      
  strCadena = "select p.`id_producto`,  CONCAT(p.`nombre_prod`,'-',cc.descripcion) as producto, d.`serie`, d.`anio_fabricacion`, " & _
  "  d.`nro_motor`,d.`nro_chasis`, d.`nro_contenedor`, e.`descripcion` as estado,l.id_linea,l.descripcion from `imp_producto_detalle` d, movimiento_compra_detalle c,  producto p , imp_estado e,imp_color cc,linea l  " & _
  "where p.id_linea=l.id_linea and p.ruc=l.id_usu and  p.id_color=cc.id_color and   d.`id_detalle_compra` = c.`id_detalle_compra` and  c.`id_producto` = p.`id_producto` and  " & _
  "p.`ruc` = '" & KEY_RUC & "' and  e.`id_estado` = d.`id_estado` and d.fecha_mod>='" & Format(Me.dtIni.Value, "YYYY-mm-dd") & "' and d.fecha_mod<='" & Format(Me.dtFin.Value, "YYYY-mm-dd") & "'  and d.id_estado in ('" & Me.DtcProceso.BoundText & "')"
      

  
  Dim arr(0 To 1, 1 To 2) As String
  arr(0, 1) = "p_fecha_ini"
  arr(0, 2) = Me.dtIni.Value
  arr(1, 1) = "p_fecha_fin"
  arr(1, 2) = Me.dtFin.Value
 
  
  Dim param As Variant
  param = arr()
  
  Call ConfiguraRstK(strCadena)
      
  Ans = ShowMultiReport(rstK, "RptProduccioncor", param, App.Path + "\Reportes\")
  
      
      
End Sub
 Private Sub llenarEstados(ByVal cbo As DataCombo)
   
    strCadena = "select e.id_estado as Codigo, e.`descripcion` as Descripcion from imp_estado e"
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcProceso)
    
End Sub
Private Sub Form_Load()
  CenterForm Me
  Me.Top = 400
  Me.dtIni.Value = KEY_FECHA
  Me.dtFin.Value = KEY_FECHA
  Call llenarEstados(Me.DtcProceso)
 

End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmcierre 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn cmdReporte 
      Height          =   615
      Left            =   6600
      TabIndex        =   7
      Top             =   6600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "       GENERAR REPORTE"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcierre.frx":0000
      PICN            =   "frmcierre.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   10920
      OleObjectBlob   =   "frmcierre.frx":25ED
      Top             =   720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PARAMETROS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1935
      Left            =   6600
      TabIndex        =   3
      Top             =   2280
      Width           =   6255
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   330
         Left            =   1680
         TabIndex        =   10
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   4194304
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
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
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
         Format          =   136183809
         CurrentDate     =   43507
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   300
         Left            =   4200
         TabIndex        =   5
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
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
         Format          =   136183809
         CurrentDate     =   43507
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ALMACEN :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHAS :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   390
         Width           =   735
      End
   End
   Begin VB.OptionButton opt_cierre_almacen 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "CIERRE DE ALMACEN"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   350
      Left            =   480
      TabIndex        =   2
      Top             =   3360
      Width           =   3855
   End
   Begin VB.OptionButton opt_cierre_ventas 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "CIERRE DE VENTAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   350
      Left            =   480
      TabIndex        =   1
      Top             =   2880
      Width           =   3855
   End
   Begin VB.OptionButton opt_cierre_caja 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "CIERRE DE CAJA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   350
      Left            =   480
      TabIndex        =   0
      Top             =   2400
      Value           =   -1  'True
      Width           =   3855
   End
   Begin VitekeySoft.ChameleonBtn cmdSalir 
      Height          =   615
      Left            =   9720
      TabIndex        =   8
      Top             =   6600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "SALIR"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmcierre.frx":2821
      PICN            =   "frmcierre.frx":283D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar prog_indicador 
      Height          =   195
      Left            =   6600
      TabIndex        =   11
      Top             =   6360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image Image1 
      Height          =   1605
      Left            =   0
      Picture         =   "frmcierre.frx":2C2D
      Top             =   120
      Width           =   11250
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   7485
      Left            =   0
      Top             =   0
      Width           =   14400
   End
End
Attribute VB_Name = "frmcierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReporte_Click()
Dim param As Variant
Dim cam3(0 To 2, 1 To 2)  As String

If Me.opt_cierre_caja.Value = True Then
   cam3(0, 1) = "inicial"
                    cam3(1, 1) = "final"
                    cam3(2, 1) = "almacen"
                    cam3(0, 2) = Format(DtpDesde.Value, "dd-mm-YYYY")
                    cam3(1, 2) = Format(DtpHasta.Value, "dd-mm-YYYY")
                    cam3(2, 2) = Me.DtcAlmacen.Text
                    param = cam3()

          turno = ""
          StrAlmacen = ""
          operador = ""
          in_ventanilla = ""
         
          
          
          
          

    
    strCadena = "SELECT fecha_emision,documento,id_cliente,ncliente,id_forma_pago,id_tarjeta,tarjeta,nformapago,anulado,tipo_movimiento,monto_caja " & _
   "  FROM view_reporte_detallado_ultimate WHERE  afecta_factura='no' and  afecta_caja='si' and id_ventanilla LIKE '%" & Trim(in_ventanilla) & "%' and  id_vendedor LIKE '%" & Trim(operador) & "%' AND  id_alm LIKE  '%" & StrAlmacen & "%' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_recibo=0 AND ruc='" & KEY_RUC & "' order by id_forma_pago, fecha_emision asc,id_doc ASC,serie ASC,numero ASC"
   

Call ConfiguraRst(strCadena)

Ans = ShowMultiReport(rst, "Rpt_detallado_caja_ii", param, App.Path + "\Reportes\")
Exit Sub
End If




If Me.opt_cierre_ventas.Value = True Then
    cam3(0, 1) = "inicial"
    cam3(1, 1) = "final"
    cam3(2, 1) = "almacen"
    cam3(0, 2) = Format(DtpDesde.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(DtpHasta.Value, "dd-mm-YYYY")
    cam3(2, 2) = Me.DtcAlmacen.Text
    param = cam3()
    
    strCadena = "SELECT fecha_emision,id_doc,doc_des,documento,id_cliente,ncliente,forma_pago,forma_pago_detalle,monto_caja FROM view_cierre_venta WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' "
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "Rpt_cierre_venta", param, App.Path + "\Reportes\")
    
   Exit Sub
End If




If Me.opt_cierre_almacen.Value = True Then

    Dim in_saldo_inicial As Double
    cam3(0, 1) = "inicial"
    cam3(1, 1) = "final"
    cam3(2, 1) = "almacen"
    cam3(0, 2) = Format(DtpDesde.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(DtpHasta.Value, "dd-mm-YYYY")
    cam3(2, 2) = Me.DtcAlmacen.Text
    param = cam3()


'saldo incial

    strCadena = "DELETE FROM kardex_valorizado_sunat WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    
    strCadena = "SELECT sum(k.`cantidad_real`*k.`costo_unitario`) as inicial FROM  kardex k WHERE k.`fecha_emision`<'" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and k.`ruc`='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    in_saldo_inicial = rst(0)
    
  








strCadena = "SELECT DISTINCT k.id_producto,p.nombre_prod FROM kardex k,producto p WHERE k.id_producto<>'00' and  k.id_producto=p.id_producto and k.ruc=p.ruc   and k.fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and  k.fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and  k.ruc='" & KEY_RUC & "' ORDER BY k.id_producto"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.prog_indicador.Min = 0
   Me.prog_indicador.Max = rst.RecordCount
   
   strCadena = "call put_crear_kardex_temporal('" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & KEY_RUC & "');"
   CnBd.Execute (strCadena)
   
   
   For i = 0 To rst.RecordCount - 1
        strCadena = "CALL procedure_kardex_cierre('" & rst("id_producto") & "','" & rst("nombre_prod") & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        rst.MoveNext
        Me.prog_indicador.Value = i
        Me.cmdReporte.Caption = (i + 1) & Space(2) & "-" & Space(1) & rst.RecordCount
        DoEvents
   Next i

End If



strCadena = "SELECT id_producto,producto,unidad,linea,cantidad_inicial,cantidad_ingreso,cantidad_salida,cantidad_final,saldo_final,valorizado,'" & in_saldo_inicial & "' FROM view_cierre_almacen WHERE dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "Rpt_cierre_almacen", param, App.Path + "\Reportes\")


    Exit Sub
End If



End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()

CenterForm Me
Me.Top = 500

Me.DtpDesde.Value = KEY_FECHA
Me.DtpHasta.Value = KEY_FECHA


strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE id_sucursal='0' and id_tipoentidad='0' and  ruc='" & KEY_RUC & "'  ORDER BY descripcion "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM

End Sub


VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmMatricula 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk_grado 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "GRADO :"
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
      Height          =   280
      Left            =   12000
      TabIndex        =   16
      Top             =   360
      Width           =   1000
   End
   Begin VB.CheckBox chk_nivel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "NIVEL :"
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
      Height          =   280
      Left            =   4440
      TabIndex        =   13
      Top             =   360
      Width           =   855
   End
   Begin VB.CheckBox chk_periodo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "PERIODO :"
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
      Height          =   280
      Left            =   8040
      TabIndex        =   12
      Top             =   360
      Width           =   1000
   End
   Begin MSDataListLib.DataCombo DtcPeriodo 
      Height          =   315
      Left            =   9045
      TabIndex        =   10
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtDni 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3135
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TxtApellido 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   975
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   7935
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   13996
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
   Begin VitekeySoft.ChameleonBtn cmdexit 
      Height          =   855
      Left            =   18960
      TabIndex        =   3
      Top             =   8110
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmatricula.frx":0000
      PICN            =   "frmmatricula.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmddelete 
      Height          =   855
      Left            =   18960
      TabIndex        =   4
      Top             =   2850
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "ELIMINAR"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmatricula.frx":040C
      PICN            =   "frmmatricula.frx":0428
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdupdate 
      Height          =   855
      Left            =   18960
      TabIndex        =   5
      Top             =   1965
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MODIFICAR"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmatricula.frx":2872
      PICN            =   "frmmatricula.frx":288E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   18960
      TabIndex        =   6
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "NUEVO"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmatricula.frx":2BA8
      PICN            =   "frmmatricula.frx":2BC4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdmatricula 
      Height          =   855
      Left            =   18960
      TabIndex        =   7
      Top             =   4605
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "PAGOS"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmatricula.frx":3016
      PICN            =   "frmmatricula.frx":3032
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdficha 
      Height          =   855
      Left            =   18960
      TabIndex        =   8
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MATRICULA"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmatricula.frx":592A
      PICN            =   "frmmatricula.frx":5946
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdContrato 
      Height          =   855
      Left            =   18960
      TabIndex        =   9
      Top             =   5480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "CONTRATO"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmatricula.frx":7B9F
      PICN            =   "frmmatricula.frx":7BBB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdReserva 
      Height          =   855
      Left            =   18960
      TabIndex        =   11
      Top             =   6340
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "CONSTANCIA"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmatricula.frx":A18C
      PICN            =   "frmmatricula.frx":A1A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcNivel 
      Height          =   315
      Left            =   5295
      TabIndex        =   14
      Top             =   360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcGrado 
      Height          =   315
      Left            =   13005
      TabIndex        =   15
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VitekeySoft.ChameleonBtn cmdReporte 
      Height          =   855
      Left            =   18960
      TabIndex        =   18
      Top             =   7220
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "REPORTE"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmatricula.frx":A4C2
      PICN            =   "frmmatricula.frx":A4DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcEstado 
      Height          =   315
      Left            =   16800
      TabIndex        =   21
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI :"
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
      Left            =   2760
      TabIndex        =   20
      Top             =   360
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "APELLIDO :"
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
      Left            =   240
      TabIndex        =   19
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblcount 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   18720
      TabIndex        =   17
      Top             =   360
      Width           =   1335
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   120
      Top             =   180
      Width           =   19935
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9135
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmMatricula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub cmdContrato_Click()

strCadena = "SELECT * FROM view_contrato_colegio  WHERE dni='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' and id_parentesco='4' LIMIT 1 "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Ans = ShowMultiReport(rst, "rpt_contrato", , App.Path + "\Reportes\")
End If


End Sub

Private Sub cmddelete_Click()
If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        strCadena = "SELECT * FROM movimiento_venta WHERE id_cliente='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' and ruc='" & KEY_RUC & "' "
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
         strCadena = "call p_delete_persona('" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "','" & KEY_RUC & "')"
         CnBd.Execute (strCadena)
          
         'Call actualizar
         Else
            MsgBox "Imposible Eliminar a este Usuario, esta Vinculado a Movimientos"
         End If
      End If
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdficha_Click()
Dim in_dni As String

in_dni = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
strCadena = "SELECT `dni`,`id_anio`,`fecha`,`a_paterno`,`a_materno`,`nombres`,`nacimiento`,`departamento`,`provincia`,`distrito`,`vive_papa`,`vive_mama`,`estatura`,`peso`,`tipo_nacimiento`, " & _
"`enfermedades`,`vacunas`,`alergias`" & get_familiar(in_dni, "1") & get_familiar(in_dni, "2") & get_familiar(in_dni, "4") & ", `procedencia`,`promovido`,`requiere_recuperacion`,`tercio_estudiantil`,`habilidad`,`grado`,`nivel` " & _
" FROM view_ficha_matricula_v2 WHERE id_matricula='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_ficha_matricula", param, App.Path + "\Reportes\")

End Sub

Private Sub cmdmatricula_Click()
Dim in_dni As String
'Call FrmVentas.Show
Call FrmVentas.activar
strCadena = "P_nueva_venta('" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
in_dni = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
FrmVentas.TxtCodCliente.Text = in_dni
Call FrmVentas.precionar_cliente
strCadena = "SELECT * FROM college_servicio_persona WHERE dni='" & in_dni & "' and saldo>0 and ruc='" & KEY_RUC & "' ORDER BY id_detalle ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "INSERT INTO temporal_ventas(ruc,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save) VALUES " & _
        "('" & KEY_RUC & "','" & in_dni & "','" & KEY_ALM & "','" & FrmVentas.DtcTipoDoc.BoundText & "','" & Trim(FrmVentas.DtcSerieDoc.BoundText) & "','" & Trim(FrmVentas.TxtNumeroDoc.Text) & "','" & rst("id_servicio") & "','1'," & _
        "'" & Val(rst("saldo")) & " ','" & Val(rst("saldo")) & "','0','si','" & rst("detalle") & "','" & KEY_USUARIO & "')"
        CnBd.Execute (strCadena)
        rst.MoveNext
   Next i
End If
Call FrmVentas.llenarGrid_det(FrmVentas.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText, Trim(FrmVentas.txtformato_impresion.Text))
       

End Sub

Private Sub cmdNuevo_Click()
 Procedencia = nuevo
      FrmDetallePersona.Show
      'Call Resalta(FrmDetallePersona.TxtRuc)
      Exit Sub
End Sub

Private Sub cmdReporte_Click()
If Me.chk_periodo.Value = 1 And Me.chk_grado.Value = 1 And Me.chk_nivel.Value = 1 Then
    strCadena = "SELECT dni,nombre_completo,periodo,nivel,grado,dni_apoderado,apoderado,direccion,telefono,estado FROM view_matricula_reporte WHERE id_periodo='" & Me.DtcPeriodo.BoundText & "' and id_grado='" & Me.DtcGrado.BoundText & "' and id_nivel='" & Me.DtcNivel.BoundText & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "Rptmatricula_reporte", , App.Path + "\Reportes\")
    Exit Sub
End If


If Me.chk_periodo.Value = 1 And Me.chk_grado.Value = 1 Then
    strCadena = "SELECT dni,nombre_completo,periodo,nivel,grado,dni_apoderado,apoderado,direccion,telefono,estado FROM view_matricula_reporte WHERE id_periodo='" & Me.DtcPeriodo.BoundText & "' and id_grado='" & Me.DtcGrado.BoundText & "'  and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "Rptmatricula_reporte", , App.Path + "\Reportes\")
    Exit Sub
End If

If Me.chk_periodo.Value = 1 And Me.chk_nivel.Value = 1 Then
    strCadena = "SELECT dni,nombre_completo,periodo,nivel,grado,dni_apoderado,apoderado,direccion,telefono,estado FROM view_matricula_reporte WHERE id_periodo='" & Me.DtcPeriodo.BoundText & "' and id_nivel='" & Me.DtcNivel.BoundText & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "Rptmatricula_reporte", , App.Path + "\Reportes\")
    Exit Sub
End If

If Me.chk_grado.Value = 1 And Me.chk_nivel.Value = 1 Then
    strCadena = "SELECT dni,nombre_completo,periodo,nivel,grado,dni_apoderado,apoderado,direccion,telefono,estado FROM view_matricula_reporte WHERE id_grado='" & Me.DtcGrado.BoundText & "' and  id_periodo='" & Me.DtcPeriodo.BoundText & "' and id_nivel='" & Me.DtcNivel.BoundText & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "Rptmatricula_reporte", , App.Path + "\Reportes\")
    Exit Sub
End If


If Me.chk_periodo.Value = 1 Then
    strCadena = "SELECT dni,nombre_completo,periodo,nivel,grado,dni_apoderado,apoderado,direccion,telefono,estado FROM view_matricula_reporte WHERE id_periodo='" & Me.DtcPeriodo.BoundText & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "Rptmatricula_reporte", , App.Path + "\Reportes\")
    Exit Sub
End If

If Me.chk_grado.Value = 1 Then
    strCadena = "SELECT dni,nombre_completo,periodo,nivel,grado,dni_apoderado,apoderado,direccion,telefono,estado FROM view_matricula_reporte WHERE id_periodo='" & Me.DtcPeriodo.BoundText & "' and  id_grado='" & Me.DtcGrado.BoundText & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "Rptmatricula_reporte", , App.Path + "\Reportes\")
    Exit Sub
End If

If Me.chk_nivel.Value = 1 Then
    strCadena = "SELECT dni,nombre_completo,periodo,nivel,grado,dni_apoderado,apoderado,direccion,telefono,estado FROM view_matricula_reporte WHERE id_periodo='" & Me.DtcPeriodo.BoundText & "' and id_nivel='" & Me.DtcNivel.BoundText & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "Rptmatricula_reporte", , App.Path + "\Reportes\")
    Exit Sub
End If




End Sub

Private Sub cmdReserva_Click()
Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant

arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"

strCadena = "SELECT CONCAT(nivel,':',codigo_modular) as modular FROM view_ficha_matricula_v2 WHERE id_matricula='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' "
Call ConfiguraRst(strCadena)



arr(0, 2) = rst("modular")
arr(1, 2) = rst("modular")


param = arr()





strCadena = "SELECT `dni`,`nombre_completo`,`periodo`,`nivel`,`grado`  FROM view_ficha_matricula_v2 WHERE id_matricula='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "Rptconstancia_vacante", param, App.Path + "\Reportes\")
End Sub

Private Sub cmdupdate_Click()
 Procedencia = modificar
 FrmDetallePersona.Show
End Sub

Private Sub DtcEstado_KeyPress(KeyAscii As Integer)
Dim in_nivel As String
Dim in_periodo As String
Dim in_grado As String

If KeyAscii = 13 Then
    
    If Me.chk_nivel.Value = 1 Then
       in_nivel = Me.DtcNivel.BoundText
    Else
        in_nivel = ""
    End If
    If Me.chk_grado.Value = 1 Then
        in_grado = Me.DtcGrado.BoundText
    Else
        in_grado = ""
    End If
    If Me.chk_periodo.Value = 1 Then
        in_periodo = Me.DtcPeriodo.BoundText
    Else
        in_periodo = ""
    End If
    
    
    strCadena = "SELECT * FROM view_matricula WHERE id_nivel LIKE '%" & in_nivel & "%' and id_periodo LIKE '%" & in_periodo & "%' and id_grado LIKE  '%" & in_grado & "%' and  id_estado = '" & Me.DtcEstado.BoundText & "' and  ruc='" & KEY_RUC & "'"
    Call Me.llenarGrid(Me.HfdPersona)
End If
End Sub

Private Sub DtcGrado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_matricula WHERE id_grado = '" & Me.DtcGrado.BoundText & "' and  ruc='" & KEY_RUC & "'and id_periodo='" & Me.DtcPeriodo.BoundText & "'  "
    Call Me.llenarGrid(Me.HfdPersona)
End If
End Sub

Private Sub DtcNivel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_matricula WHERE id_nivel = '" & Me.DtcNivel.BoundText & "' and  ruc='" & KEY_RUC & "'and id_periodo='" & Me.DtcPeriodo.BoundText & "'  "
    Call Me.llenarGrid(Me.HfdPersona)
End If
End Sub

Private Sub DtcPeriodo_Change()
strCadena = "SELECT id_grado as Codigo, CONCAT(descripcion,' [ ',nivel,' ] ') as Descripcion FROM view_nivel_periodo WHERE id_periodo='" & Me.DtcPeriodo.BoundText & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcGrado)
End Sub

Private Sub DtcPeriodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_matricula WHERE   ruc='" & KEY_RUC & "'and id_periodo='" & Me.DtcPeriodo.BoundText & "'  "
    Call Me.llenarGrid(Me.HfdPersona)
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
strCadena = "SELECT id_periodo as Codigo, descripcion as Descripcion FROM college_periodo WHERE ruc='" & KEY_RUC & "' ORDER BY id_periodo DESC "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPeriodo)
Me.DtcPeriodo.BoundText = get_periodo_anio


strCadena = "SELECT id_nivel as Codigo, descripcion as Descripcion FROM nivel_educativo "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcNivel)

strCadena = "SELECT id_grado as Codigo, CONCAT(descripcion,' [ ',nivel,' ] ') as Descripcion FROM view_nivel_periodo WHERE id_periodo='" & Me.DtcPeriodo.BoundText & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcGrado)

strCadena = "SELECT id_estado as Codigo,descripcion as Descripcion FROM college_estado ORDER BY id_estado DESC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstado)
     

strCadena = "SELECT * FROM view_matricula WHERE id_alm='" & KEY_ALM & "' and  ruc='" & KEY_RUC & "'and id_periodo='" & get_periodo_anio & "'  LIMIT 29"
Call Me.llenarGrid(Me.HfdPersona)


End Sub
Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
  
   Grilla.Rows = 0
   Me.lblcount.Caption = str(rst.RecordCount) & Space(2) & "ALUMNOS"
         
 
           ReDim arrColWidth(1 To rst.Fields.Count)
           For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1200
            Grilla.ColWidth(2) = 3500
            Grilla.ColWidth(3) = 3200
            Grilla.ColWidth(4) = 1400
            Grilla.ColWidth(5) = 1400
            Grilla.ColWidth(6) = 1700
            Grilla.ColWidth(7) = 1300
            Grilla.ColWidth(8) = 1300
            Grilla.ColWidth(9) = 1200
            Grilla.ColWidth(10) = 1200
            Grilla.ColWidth(11) = 1000
          Next
            cabecera = "MATRICULA" & vbTab & "DNI" & vbTab & "ALUMNO" & vbTab & "DIRECCION" & vbTab & "GRADO" & vbTab & "NIVEL" & vbTab & "PERIODO" & vbTab & "MATRICULA" & vbTab & "PENSION" & vbTab & "BECA" & vbTab & "1/2 BECA" & vbTab & "ESTADO"
            Grilla.AddItem cabecera
            For k = 1 To 11
                Grilla.col = k
                Grilla.Row = 0
                Grilla.CellBackColor = &HDFDFE0
            Next k
            rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             If rst("beca") = "si" Then
                in_beca = "[  X  ]"
              Else
                in_beca = "[     ]"
             End If
             If rst("media_beca") = "si" Then
                in_media_beca = "[  X  ]"
             Else
                in_media_beca = "[     ]"
             End If
             Fila = rst("id_matricula") & vbTab & rst("dni") & vbTab & UCase(rst("nombre_completo")) & vbTab & Mid(UCase(rst("direccion")), 1, 35) & vbTab & rst("grado") & vbTab & rst("nivel") & vbTab & rst("periodo") & vbTab & rst("matricula") & vbTab & rst("pension") & vbTab & in_beca & vbTab & in_media_beca & vbTab & rst("estado")
             Grilla.AddItem Fila
            
        rst.MoveNext
        Next i
     
     
  Grilla.ColAlignment(0) = 1
  Grilla.ColAlignment(1) = 1
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
         
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub TxtApellido_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_matricula WHERE nombre_completo LIKE '%" & Trim(Me.TxtApellido.Text) & "%' and  ruc='" & KEY_RUC & "'and id_periodo='" & Me.DtcPeriodo.BoundText & "'"
    Call Me.llenarGrid(Me.HfdPersona)
End If
End Sub

Private Sub txtDni_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_matricula WHERE dni LIKE '%" & Trim(Me.txtDni.Text) & "%' and  ruc='" & KEY_RUC & "'and id_periodo='" & Me.DtcPeriodo.BoundText & "'"
    Call Me.llenarGrid(Me.HfdPersona)
End If




End Sub

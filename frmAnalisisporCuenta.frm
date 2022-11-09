VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAnalisisporCuenta 
   BorderStyle     =   0  'None
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   17250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Command3"
      Height          =   495
      Left            =   12000
      TabIndex        =   32
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   855
      Left            =   12000
      TabIndex        =   31
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox TxtRucVinculo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   8640
      TabIndex        =   28
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtIdContable 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   7080
      TabIndex        =   26
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CheckBox chk_periodoContable 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "SALDO REFERENCIA :"
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
      Left            =   5280
      TabIndex        =   24
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdateFecha 
      Caption         =   "FECHAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15240
      TabIndex        =   23
      ToolTipText     =   "ACTUALIZA FECHAS CORRECTAS"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox chk_rango 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "RANGO PERIODOS"
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
      Height          =   675
      Left            =   5280
      TabIndex        =   21
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SALDO INICIAL DE FACTURAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15240
      TabIndex        =   18
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "VINCULAR ASIENTOS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15240
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo DtcPeriodo 
      Height          =   330
      Left            =   2520
      TabIndex        =   12
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin VB.CheckBox chk_periodo 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "SALDO REFERENCIA :"
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
      Left            =   720
      TabIndex        =   11
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtCuentaFin 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   13800
      TabIndex        =   9
      Top             =   300
      Width           =   1215
   End
   Begin VB.TextBox txtCuentaIni 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   12480
      TabIndex        =   8
      Top             =   300
      Width           =   1215
   End
   Begin VB.OptionButton optDolares 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "DOLARES"
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
      Left            =   13800
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton optSoles 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "SOLES"
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
      Left            =   12480
      TabIndex        =   6
      Top             =   720
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VitekeySoft.ChameleonBtn cmdBuscar 
      Height          =   435
      Left            =   15240
      TabIndex        =   0
      Top             =   2400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BTYPE           =   5
      TX              =   "POR DOCUMENTO"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAnalisisporCuenta.frx":0000
      PICN            =   "frmAnalisisporCuenta.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpDesde 
      Height          =   315
      Left            =   7080
      TabIndex        =   2
      Top             =   300
      Width           =   1380
      _ExtentX        =   2434
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
      Format          =   146997249
      CurrentDate     =   41251
   End
   Begin MSComCtl2.DTPicker DtpHasta 
      Height          =   315
      Left            =   8760
      TabIndex        =   3
      Top             =   300
      Width           =   1380
      _ExtentX        =   2434
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
      Format          =   146997249
      CurrentDate     =   41251
   End
   Begin VitekeySoft.ChameleonBtn cmdAnalisisDetallado 
      Height          =   435
      Left            =   15240
      TabIndex        =   10
      Top             =   2880
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BTYPE           =   5
      TX              =   "POR CUENTA           "
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAnalisisporCuenta.frx":2601
      PICN            =   "frmAnalisisporCuenta.frx":261D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdClose 
      Height          =   315
      Left            =   16800
      TabIndex        =   13
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "    "
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAnalisisporCuenta.frx":4C02
      PICN            =   "frmAnalisisporCuenta.frx":4C1E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcPeriodoUnico 
      Height          =   330
      Left            =   7080
      TabIndex        =   19
      Top             =   720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin MSDataListLib.DataCombo DtcPeriodoIni 
      Height          =   330
      Left            =   7080
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin MSDataListLib.DataCombo DtcPeriodoFin 
      Height          =   330
      Left            =   7080
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
      Height          =   315
      Left            =   10320
      TabIndex        =   30
      Top             =   2400
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   ""
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAnalisisporCuenta.frx":7AD2
      PICN            =   "frmAnalisisporCuenta.frx":7AEE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI/RUC"
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
      Left            =   8640
      TabIndex        =   29
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID CONTABLE"
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
      Left            =   7335
      TabIndex        =   27
      Top             =   2160
      Width           =   915
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "VINCULAR:"
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
      Height          =   315
      Left            =   5280
      TabIndex        =   25
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lblcliente 
      BackColor       =   &H00C0C0C0&
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
      Height          =   315
      Left            =   720
      TabIndex        =   17
      Top             =   1080
      Width           =   3150
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RANGO DE FECHAS:"
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
      Left            =   5700
      TabIndex        =   16
      Top             =   360
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INGRESE N° CUENTA :"
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
      Left            =   10950
      TabIndex        =   15
      Top             =   360
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   120
      Top             =   3120
      Width           =   15015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONEDA:"
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
      Left            =   11685
      TabIndex        =   5
      Top             =   720
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "DNI / RUC :"
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
      Height          =   220
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2970
      Left            =   120
      Top             =   90
      Width           =   15015
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   8370
      Left            =   0
      Top             =   0
      Width           =   17250
   End
End
Attribute VB_Name = "frmAnalisisporCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub ChameleonBtn1_Click()

strCadena = "UPDATE con_asientomovimiento SET TipoRef='1',IdReferencia='" & Trim(Me.TxtRucVinculo.Text) & "',Referencia='" & get_persona(Trim(Me.TxtRucVinculo.Text)) & "' WHERE idAsiento='" & Trim(Me.txtIdContable.Text) & "'"
CnBd.Execute (strCadena)



MsgBox "Proceso Realizado", vbInformation





End Sub

Private Sub Check1_Click()

End Sub

Private Sub chk_periodo_Click()
If Me.chk_periodo.Value = 1 Then
   Me.DtcPeriodo.Enabled = True
Else
   Me.DtcPeriodo.Enabled = False
End If
End Sub

Private Sub chk_rango_Click()

If Me.chk_rango.Value = 1 Then
   
   Me.DtcPeriodoIni.Visible = True
   Me.DtcPeriodoFin.Visible = True
Else
    Me.DtcPeriodoIni.Visible = False
    Me.DtcPeriodoFin.Visible = False
End If

End Sub

Private Sub cmdAnalisisDetallado_Click()
Dim cam3(0 To 2, 1 To 2)  As String
Dim in_origen As String
If Me.optSoles.Value = True Then
   in_moneda = "00001"
Else
   in_moneda = "00002"
End If

                    cam3(0, 1) = "in_fecha_ini"
                    cam3(1, 1) = "in_fecha_fin"
                    cam3(2, 1) = "in_vendedor"
                    
                   
    
                    cam3(0, 2) = Format(Me.DtpDesde.Value, "dd-mm-YYYY")
                    cam3(1, 2) = Format(Me.DtpHasta.Value, "dd-mm-YYYY")
                    cam3(2, 2) = KEY_VENDEDOR
                    param = cam3()
                  
                  
               ' If Me.opt_origen_si.Value = True Then
                   
                '   strCadena = "call CON_Analisis42_LST('2','" & Me.DtcPeriodo.BoundText & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtCuentaIni.Text) & "','" & Trim(Me.txtCuentaFin.Text) & "','" & in_moneda & "','" & KEY_RUC & "')"
               ' Else
              
                If Me.chk_periodo.Value = 1 Then
                   in_periodo = Me.DtcPeriodo.BoundText
                Else
                   in_periodo = ""
                End If
                
                
                
                strCadena = "call CON_Analisis42_LST('8','" & in_periodo & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtCuentaIni.Text) & "','" & Trim(Me.txtCuentaFin.Text) & "','" & in_moneda & "','" & Trim(Me.txtRuc.Text) & "','" & KEY_RUC & "')"
                Call ConfiguraRst(strCadena)
                Ans = ShowMultiReport(rst, "RptAnalisisCuenta_detalle", param, App.Path + "\Reportes\")
                
                'strCadena = "call CON_AnalisisCuenta_LST('3','" & Me.DtcPeriodoUnico.BoundText & "','" & in_moneda & "','" & Trim(Me.txtRuc.Text) & "','" & KEY_RUC & "')"
                'Call ConfiguraRst(strCadena)
                'Ans = ShowMultiReport(rst, "RptAnalisisCuenta_detalle", param, App.Path + "\Reportes\")
                
                
                
                
                       
                       
                       

End Sub

Private Sub cmdBuscar_Click()
Dim cam3(0 To 2, 1 To 2)  As String

If Me.optSoles.Value = True Then
   in_moneda = "00001"
Else
   in_moneda = "00002"
End If

                    cam3(0, 1) = "in_fecha_ini"
                    cam3(1, 1) = "in_fecha_fin"
                    cam3(2, 1) = "in_vendedor"
                    
                   
    
                    cam3(0, 2) = Format(Me.DtpDesde.Value, "dd-mm-YYYY")
                    cam3(1, 2) = Format(Me.DtpHasta.Value, "dd-mm-YYYY")
                    cam3(2, 2) = KEY_VENDEDOR
                    param = cam3()
                  
                  
                   '    If Me.opt_origen_si.Value = True Then
                    '        strCadena = "call CON_Analisis42_LST('1','" & Me.DtcPeriodo.BoundText & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtCuentaIni.Text) & "','" & Trim(Me.txtCuentaFin.Text) & "','" & in_moneda & "','" & KEY_RUC & "')"
                    '   Else
                          
                        strCadena = "call CON_Analisis42_LST('9','" & in_periodo & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtCuentaIni.Text) & "','" & Trim(Me.txtCuentaFin.Text) & "','" & in_moneda & "','" & Trim(Me.txtRuc.Text) & "','" & KEY_RUC & "')"
                         
                    '   End If
                
                  
                
                       Call ConfiguraRst(strCadena)
                       Ans = ShowMultiReport(rst, "RptAnalisisCuenta", param, App.Path + "\Reportes\")








End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdUpdateFecha_Click()

strCadena = "SELECT Id,FechaInicio FROM con_periodo  ORDER BY Id DESC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        
        strCadena = "SELECT Id FROM con_asiento a WHERE fecha_busqueda is null and  a.IdPeriodo='" & rst("id") & "' and  a.idEmpresasis='" & KEY_RUC & "' ORDER BY Id ASC"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            'rstK.MoveFirst
         '   For j = 0 To rstK.RecordCount - 1
                strCadena = "UPDATE con_asiento SET fecha_busqueda='" & Format(rst("FechaInicio"), "YYYY-mm-dd") & "' WHERE fecha_busqueda is null and  IdPeriodo='" & rst("id") & "' and idEmpresasis='" & KEY_RUC & "' "
                CnBd.Execute (strCadena)
           '     rstK.MoveNext
               
          '  Next j
        End If
        
        rst.MoveNext
        DoEvents
   Next i
End If





End Sub

Private Sub Command1_Click()






Dim in_referencia As String
strCadena = "SELECT Id,Glosa FROM con_asiento WHERE IdTipoAsiento IN ('1CIX000000000137') and Activo='1' and  fecha>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and IdEmpresaSis='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT  IdReferencia,Referencia FROM con_asientomovimiento WHERE TipoRef='1' and  idAsiento='" & rst("id") & "' AND Activo='1' and idEmpresasis='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
           in_referencia = rstL("IdReferencia")
           in_referencia = rstL("Referencia")
        
        Else
            GoTo siguiente:
        End If
         
         strCadena = "UPDATE con_asientomovimiento SET IdReferencia='" & rstL("IdReferencia") & "',Referencia='" & in_referencia & "' WHERE TipoRef='0' and IdReferencia='' and  idAsiento='" & rst("id") & "' AND Activo='1' and idEmpresasis='" & KEY_RUC & "' ORDER BY IdReferencia DESC "
         CnBd.Execute (strCadena)
         
                
         
                
                        
                
        
siguiente:
        DoEvents
        rst.MoveNext
        Me.Command1.Caption = str(i) & Space(5) & str(rst.RecordCount)
   Next i
End If


End Sub

Private Sub Command2_Click()









Dim in_asiento As String
Dim in_cambio As Single

in_cambio = 3.379

strCadena = "SELECT Id FROM con_asiento WHERE IdMoneda='1CIX000000000005' and  asiento_inicial='si' and  IdEmpresaSis='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
    in_asiento = rstIN("Id")
    
    strCadena = "DELETE FROM con_asiento WHERE IdMoneda='1CIX000000000005' and  asiento_inicial='si' and  IdEmpresaSis='" & KEY_RUC & "' LIMIT 1"
    CnBd.Execute (strCadena)
    
    strCadena = "DELETE FROM con_asientomovimiento WHERE IdAsiento='" & in_asiento & "'   and  IdEmpresaSis='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
End If


'8601000000121517

in_asiento = get_asiento_id("con_asiento")

strCadena = "INSERT INTO con_asiento(Id,`IdEmpresaSis`,`Fecha`,`Glosa`,`IdMoneda`,`TipoCambio`,`TotalDebe`,`TotalHaber`,`UsuarioCrea`,`FechaCrea`,`Activo`,`asiento_inicial`,fecha_busqueda) VALUES " & _
"('" & in_asiento & "','" & KEY_RUC & "','2019-01-01','SALDO INICIAL','1CIX000000000005','" & in_cambio & "',0,0,'" & KEY_USUARIO & "',CURDATE(),'1','si','2019-01-01')"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM migracion_asientomovimiento WHERE id IN (1685,1686) ORDER BY id ASC "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   in_saldo_debe = 0
   in_saldo_haber = 0
   
   For i = 0 To rst.RecordCount - 1
       in_glosa = "SALDO INICIAL :" & rst("comprobante")
       in_asiento = "8601000000121518"
       
       
       strCadena = "INSERT INTO con_asientomovimiento(`Id`,`IdEmpresaSis`,`IdAsiento`,`IdCuentaContable`,`Glosa`,`DebeMN`,`HaberMN`,`DebeME`,`HaberME`, " & _
       "`UsuarioCrea`,`FechaCrea`,`Activo`,`TipoRef`,`IdReferencia`,`Referencia`,`TipoDoc`,`SerieDoc`,`NroDoc`,`RazonSocial`,`IdMoneda`)VALUES " & _
       "('" & get_asiento_id("con_asientomovimiento") & "','" & KEY_RUC & "','" & in_asiento & "','" & get_id_cuenta(rst("nrocuenta")) & "'," & _
       "'" & in_glosa & "','" & rst("debe") & "','" & rst("haber") & "','" & rst("debe") / in_cambio & "','" & rst("haber") / in_cambio & "','" & KEY_USUARIO & "'," & _
       "CURDATE(),'1','1','" & rst("dni") & "','" & rst("razonsocial") & "','" & Format(rst("id_doc"), "0000") & "','" & Format(rst("serie"), "000") & "','" & Format(rst("numero"), "000000") & "','" & rst("razonsocial") & "','1CIX000000000005') "
       CnBd.Execute (strCadena)
       in_saldo_debe = in_saldo_debe + rst("debe")
       in_saldo_haber = in_saldo_haber + rst("haber")
       
       rst.MoveNext
       
   Next i
   
   strCadena = "UPDATE con_asiento SET TotalDebe='" & in_saldo_debe & "',TotalHaber='" & in_saldo_haber & "' WHERE Id='" & in_asiento & "' LIMIT 1"
   CnBd.Execute (strCadena)
End If




strCadena = "SELECT Id FROM con_asiento WHERE IdMoneda='1CIX000000000008' and  asiento_inicial='si' and  IdEmpresaSis='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
     in_asiento = rstIN("Id")
    
    strCadena = "DELETE FROM con_asiento WHERE IdMoneda='1CIX000000000008' and  asiento_inicial='si' and  IdEmpresaSis='" & KEY_RUC & "' LIMIT 1"
    CnBd.Execute (strCadena)
    
    strCadena = "DELETE FROM con_asientomovimiento WHERE IdAsiento='" & in_asiento & "' and   and  IdEmpresaSis='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
End If






in_asiento = get_asiento_id("con_asiento")

strCadena = "INSERT INTO con_asiento(Id,`IdEmpresaSis`,`Fecha`,`Glosa`,`IdMoneda`,`TipoCambio`,`TotalDebe`,`TotalHaber`,`UsuarioCrea`,`FechaCrea`,`Activo`,`asiento_inicial`) VALUES " & _
"('" & in_asiento & "','" & KEY_RUC & "','2019-01-01','SALDO INICIAL','1CIX000000000008','" & in_cambio & "',0,0,'" & KEY_USUARIO & "',CURDATE(),'1','si')"
CnBd.Execute (strCadena)



strCadena = "SELECT * FROM migracion_asientomovimiento WHERE id_moneda='00002' and id='439'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   in_saldo_debe = 0
   in_saldo_haber = 0
   
   For i = 0 To rst.RecordCount - 1
       in_glosa = "SALDO INICIAL :" & rst("comprobante")
       
       
       strCadena = "INSERT INTO con_asientomovimiento(`Id`,`IdEmpresaSis`,`IdAsiento`,`IdCuentaContable`,`Glosa`,`DebeMN`,`HaberMN`,`DebeME`,`HaberME`, " & _
       "`UsuarioCrea`,`FechaCrea`,`Activo`,`TipoRef`,`IdReferencia`,`Referencia`,`TipoDoc`,`SerieDoc`,`NroDoc`,`RazonSocial`,`IdMoneda`)VALUES " & _
       "('" & get_asiento_id("con_asientomovimiento") & "','" & KEY_RUC & "','" & in_asiento & "','" & get_id_cuenta(rst("nrocuenta")) & "'," & _
       "'" & in_glosa & "','" & rst("debe") * in_cambio & "','" & rst("haber") * in_cambio & "','" & rst("debe") & "','" & rst("haber") & "','" & KEY_USUARIO & "'," & _
       "CURDATE(),'1','1','" & rst("dni") & "','" & rst("razonsocial") & "','" & Format(rst("id_doc"), "0000") & "','" & Format(rst("serie"), "000") & "','" & rst("numero") & "','" & rst("razonsocial") & "','1CIX000000000005') "
       CnBd.Execute (strCadena)
       in_saldo_debe = in_saldo_debe + rst("debe") * in_cambio
       in_saldo_haber = in_saldo_haber + rst("haber") * in_cambio
       
      
       
       rst.MoveNext
       
   Next i
   strCadena = "UPDATE con_asiento SET TotalDebe='" & in_saldo_debe & "',TotalHaber='" & in_saldo_haber & "' WHERE Id='" & in_asiento & "' LIMIT 1"
   CnBd.Execute (strCadena)
   
End If

MsgBox "Migracion Exitosa", vbInformation


End Sub
Private Function get_id_cuenta(ByVal in_cuenta As String) As String

strCadena = "SELECT * FROM con_cuentacontable WHERE IdEmpresaSis='" & KEY_RUC & "' and  NroCuenta='" & in_cuenta & "' and Ejercicio='2019' LIMIT 1"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
   get_id_cuenta = rstIN("Id")
End If

End Function
Private Function get_asiento_id(ByVal in_tabla As String) As String
Dim in_asiento As String
Dim in_numero  As Integer
strCadena = "SELECT Id FROM " & in_tabla & "  ORDER BY Id DESC LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   in_asiento = Mid(rstL("Id"), 1, 10)
   in_correlativo = Mid(rstL("Id"), 11, 6)
   get_asiento_id = in_asiento & Format(Val(in_correlativo) + 1, "00")
End If
End Function

Private Sub Command3_Click()

'cc.`NroCuenta`='" & Trim(Me.txtCuentaIni.Text) & "'

strCadena = "SELECT * FROM con_asiento a INNER JOIN con_asientomovimiento am ON a.`Id`=am.`IdAsiento` INNER JOIN " & _
" `con_cuentacontable` cc ON am.`IdCuentaContable`=cc.`Id`  Where  a.`Activo`=am.`Activo` AND  a.`Activo`=1 and  " & _
"  a.`IdPeriodo`='" & Me.DtcPeriodo.BoundText & "' and cc.`NroCuenta`='9511101' and am.CuentaAsociada='-' and " & _
"  a.`IdEmpresaSis`='" & KEY_RUC & "'"



Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        X = rst("idAsiento") & Space(10) & rst("IdTipoAsiento") & Space(10) & rst("nroCuenta") & Space(10) & rst("DebeMN")
        
     '   strCadena = "UPDATE con_asientomovimiento SET IdCuentaContable='2501000000007508' WHERE IdCuentaContable='2501000000002497' and idAsiento='" & rst("idAsiento") & "' LIMIT 1"
      '  CnBd.Execute (strCadena)
         
         
      '  strCadena = "UPDATE con_asientomovimiento SET IdCuentaContable='2501000000007014' WHERE IdCuentaContable='2501000000002004' and idAsiento='" & rst("idAsiento") & "' LIMIT 1"
       ' CnBd.Execute (strCadena)
        
       ' strCadena = "UPDATE con_asientomovimiento SET IdCuentaContable='2501000000006149' WHERE IdCuentaContable='2501000000001141' and idAsiento='" & rst("idAsiento") & "' LIMIT 1"
       ' CnBd.Execute (strCadena)
        
        
       
       rst.MoveNext
       DoEvents
       Me.Command3.Caption = str(i) & Space(3) & str(rst.RecordCount)
   Next i
End If


End Sub
Private Sub recorrer_asiento(ByVal in_asiento As String)

Dim in_monto1 As Double
Dim in_monto2 As Double
Dim in_monto3 As Double
strCadena = "SELECT NroCuenta,a.DebeMN,a.HaberMN,a.Id FROM con_asientomovimiento a INNER JOIN con_cuentacontable c ON  a.IdCuentaContable=c.Id WHERE " & _
" idAsiento='" & in_asiento & "' and a.Activo='1' AND Left(c.NroCuenta,1)='6'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   rstK.MoveFirst
   For i = 0 To rstK.RecordCount - 1
        
        
        strCadena = "SELECT CuentaAsociada1,Porcentaje1,CuentaAsociada2,Porcentaje3,CuentaAsociada3 FROM con_cuentaasociada WHERE   CuentaContable='" & rstK("NroCuenta") & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            
            strCadena = "SELECT NroCuenta,a.DebeMN ,a.HaberMN,a.Id FROM con_asientomovimiento a INNER JOIN con_cuentacontable c ON  a.IdCuentaContable=c.Id WHERE " & _
            " idAsiento='" & in_asiento & "' and a.Activo='1' and NroCuenta='" & rstL("CuentaAsociada1") & "' and CuentaAsociada='-' LIMIT 1 "
            Call ConfiguraRstA(strCadena)
            If rstA.RecordCount > 0 Then
                id_mov1 = rstA("Id")
                in_monto1 = rstA("DebeMN")
            Else
                in_monto1 = 0
                id_mov1 = 0
            End If
            
            
            strCadena = "SELECT NroCuenta,a.DebeMN,a.HaberMN,a.Id FROM con_asientomovimiento a INNER JOIN con_cuentacontable c ON  a.IdCuentaContable=c.Id WHERE " & _
            " idAsiento='" & in_asiento & "' and a.Activo='1' and NroCuenta='" & rstL("CuentaAsociada3") & "' and  CuentaAsociada='-' LIMIT 1  "
            Call ConfiguraRstA(strCadena)
            If rstA.RecordCount > 0 Then
                id_mov2 = rstA("Id")
                in_monto2 = rstA("DebeMN")
            Else
               in_monto2 = 0
               id_mov2 = 0
            End If
            If in_monto1 = 0 And in_monto2 = 0 Then
                GoTo a
            End If
            
            If Round(rstK("DebeMN"), 2) = Round(in_monto1 + in_monto2, 2) Then
                strCadena = "UPDATE con_asientomovimiento SET CuentaAsociada='" & rstK("NroCuenta") & "' WHERE Id='" & id_mov1 & "' LIMIT 1"
                CnBd.Execute (strCadena)
                
                 strCadena = "UPDATE con_asientomovimiento SET CuentaAsociada='" & rstK("NroCuenta") & "' WHERE Id='" & id_mov2 & "' LIMIT 1"
                CnBd.Execute (strCadena)
            Else
                X = 0
            End If
a:
            
        End If
        
        DoEvents
        rstK.MoveNext
   Next i
  End If
End Sub


Private Sub Command4_Click()

'---- recorremos los asientos 6910101
strCadena = "SELECT * FROM con_asiento a INNER JOIN con_asientomovimiento am ON a.`Id`=am.`IdAsiento` INNER JOIN " & _
" `con_cuentacontable` cc ON am.`IdCuentaContable`=cc.`Id` INNER JOIN con_cuentaasociada ca ON (cc.NroCuenta=ca.CuentaContable and cc.IdEmpresaSis=ca.IdEmpresaSis )  Where  a.`Activo`=am.`Activo` AND  a.`Activo`=1 and  " & _
"  a.`IdPeriodo`='" & Me.DtcPeriodo.BoundText & "' and  Left(cc.NroCuenta,1)='6' and  cc.NroCuenta<>'6910101' and  " & _
"  a.`IdEmpresaSis`='" & KEY_RUC & "' "

'strCadena = "SELECT * FROM con_asiento a WHERE a.idTipoAsiento<>'1CIX000000000137' and  a.Activo='1' and  a.IdPeriodo='" & Me.DtcPeriodo.BoundText & "' and idEmpresasis='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       
       Call recorrer_asiento(rst("IdAsiento"))
        X = rst("idAsiento") & Space(10) & rst("IdTipoAsiento") & Space(10) & rst("nroCuenta") & Space(10) & rst("DebeMN")
       
       rst.MoveNext
       DoEvents
       Me.Command3.Caption = str(i) & Space(3) & str(rst.RecordCount)
   Next i
End If

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100

Me.DtpDesde.Value = KEY_FECHA
Me.DtpHasta.Value = KEY_FECHA

strCadena = "SELECT Id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion  FROM con_periodo order by codigo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodo)
  Me.DtcPeriodo.BoundText = get_periodo_actual(KEY_FECHA)
  
  
  strCadena = "SELECT Id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion  FROM con_periodo order by codigo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodoIni)
  Me.DtcPeriodo.BoundText = get_periodo_actual(KEY_FECHA)
  
  strCadena = "SELECT Id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion  FROM con_periodo order by codigo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodoFin)
  Me.DtcPeriodo.BoundText = get_periodo_actual(KEY_FECHA)
  

  strCadena = "SELECT Id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion  FROM con_periodo order by codigo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodoUnico)
  Me.DtcPeriodo.BoundText = get_periodo_actual(KEY_FECHA)
  
  

End Sub

Private Sub txtCuentaFin_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Procedencia = buscar
   FrmPlanContableCuentas.Show
   Exit Sub
End If


End Sub

Private Sub txtCuentaIni_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   
   Procedencia = Selecionar
   FrmPlanContableCuentas.Show
   Exit Sub
End If


End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Me.txtRuc.Text) = "" Then
        Procedencia = Selecionar
        FrmPersona.Show
        Exit Sub
    Else
        strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           Me.lblcliente.Caption = rst("nombre_completo")
        Else
           Me.lblcliente.Caption = ""
        End If
    End If
    
End If
End Sub

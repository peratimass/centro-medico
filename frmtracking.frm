VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmtracking 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   18855
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   6705
      TabIndex        =   21
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame frmtracking 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "TRACKING COMPROBANTE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   4320
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   12855
      Begin VB.CommandButton cmdcerrar 
         Height          =   255
         Left            =   12000
         Picture         =   "frmtracking.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Hftracking 
         Height          =   4695
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   8281
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
      Begin VB.Label lbltracking 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRACKING  "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame frmestado 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
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
      Height          =   1455
      Left            =   12480
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox txtobservacion 
         Appearance      =   0  'Flat
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
         Height          =   700
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   600
         Width           =   3015
      End
      Begin MSDataListLib.DataCombo DtcEstadoproceso 
         Height          =   330
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
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
      Begin VitekeySoft.ChameleonBtn cmdprocesarestado 
         Height          =   330
         Left            =   4080
         TabIndex        =   14
         Top             =   240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         BTYPE           =   5
         TX              =   ""
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtracking.frx":2EA4
         PICN            =   "frmtracking.frx":2EC0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERV:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   585
      End
   End
   Begin VB.CheckBox chk_estado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ESTADO :"
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
      Left            =   8040
      TabIndex        =   11
      Top             =   380
      Width           =   975
   End
   Begin MSDataListLib.DataCombo DtcEstado 
      Height          =   330
      Left            =   9120
      TabIndex        =   10
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
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
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   345
      Left            =   14040
      TabIndex        =   7
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
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
      Format          =   248905729
      CurrentDate     =   42757
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   4155
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox TxtApellido 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1395
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   7935
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   17655
      _ExtentX        =   31141
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
      Left            =   17880
      TabIndex        =   3
      Top             =   3700
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmtracking.frx":5214
      PICN            =   "frmtracking.frx":5230
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdtracking 
      Height          =   855
      Left            =   17880
      TabIndex        =   4
      Top             =   1050
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "TRACKING"
      ENAB            =   0   'False
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmtracking.frx":5620
      PICN            =   "frmtracking.frx":563C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker Dtpfin 
      Height          =   345
      Left            =   15360
      TabIndex        =   8
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
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
      Format          =   247726081
      CurrentDate     =   42757
   End
   Begin VitekeySoft.ChameleonBtn cmdbuscar 
      Height          =   345
      Left            =   16800
      TabIndex        =   9
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "BUSCAR"
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
      MICON           =   "frmtracking.frx":7C75
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdEnvios 
      Height          =   855
      Left            =   17880
      TabIndex        =   23
      Top             =   1940
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "ENTREGAS"
      ENAB            =   0   'False
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmtracking.frx":7C91
      PICN            =   "frmtracking.frx":7CAD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdVisualizar 
      Height          =   855
      Left            =   17880
      TabIndex        =   24
      Top             =   2820
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "VISUALIZAR"
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
      MICON           =   "frmtracking.frx":7FC7
      PICN            =   "frmtracking.frx":7FE3
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
      Caption         =   "DOCUMENTO:"
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
      Left            =   5625
      TabIndex        =   22
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/DNI :"
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
      Left            =   3330
      TabIndex        =   6
      Top             =   360
      Width           =   705
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAZON SOCIAL:"
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
      Left            =   225
      TabIndex        =   5
      Top             =   360
      Width           =   1065
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Left            =   120
      Top             =   60
      Width           =   17655
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9135
      Left            =   0
      Top             =   0
      Width           =   18855
   End
End
Attribute VB_Name = "frmtracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede



Private Sub cmdBuscar_Click()

If Me.chk_estado.Value = 0 Then
    strCadena = "SELECT * FROM view_listado_tracking where ruc='" & KEY_RUC & "' and fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and ncliente Like '%" & Trim(Me.TxtApellido.Text) & "%' and id_cliente Like '%" & Trim(Me.txtRuc.Text) & "%' "
Else
    strCadena = "SELECT * FROM view_listado_tracking where ruc='" & KEY_RUC & "' and fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and ncliente Like '%" & Trim(Me.TxtApellido.Text) & "%' and id_cliente Like '%" & Trim(Me.txtRuc.Text) & "%' and  estado='" & Trim(Me.DtcEstado.Text) & "' "
End If

Call Me.llenarGrid(Me.HfdPersona)

End Sub

Private Sub cmdCerrar_Click()
Me.frmtracking.Visible = False

End Sub

Private Sub cmddelete_Click()

End Sub

Private Sub cmdEnvios_Click()

strCadena = "SELECT id_transferencia,id_venta,fecha,guia,id_producto,detalle,cantidad,id_cliente,ncliente,documento,fecha_emision,det_venta,cant_venta,ruc FROM view_entrega_producto WHERE id_venta='" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_entrega_producto", , App.Path + "\Reportes\")
'Me.lbltracking.Caption = "HISTORIAL DE SALIDAS"
'Call llenarGrid_envios(Me.Hftracking, Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)))
'Me.frmtracking.Visible = True
End Sub

Private Sub cmdexit_Click()

Unload Me
Exit Sub


End Sub

Private Sub cmdprocesarestado_Click()
 Call put_tracking(Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)), Me.DtcEstadoproceso.BoundText, Trim(UCase(Me.txtObservacion.Text)))
 Me.frmestado.Visible = False
 in_color = &H8080FF
            Select Case Me.DtcEstadoproceso.BoundText
                Case "05"
                    in_color = &H80FF80
                Case "04"
                    in_color = &H80FF&
            End Select
            Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 8) = Me.DtcEstadoproceso.Text
            For k = 8 To 8
                HfdPersona.col = k
                HfdPersona.Row = Me.HfdPersona.Row
                HfdPersona.CellBackColor = in_color
            Next k
 
End Sub

Private Sub cmdtracking_Click()

Me.lbltracking.Caption = "HISTORIAL"
Call llenarGrid_tracking(Me.Hftracking, Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)))
Me.frmtracking.Visible = True


End Sub


Private Sub cmdVisualizar_Click()
If Me.HfdPersona.Rows > 0 Then
    
    If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
        Procedencia = buscar
        FrmDetalleventa.Show
    End If
    
End If

End Sub

Private Sub DtcEstadoproceso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtObservacion)
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA

strCadena = "SELECT id_estado as Codigo,descripcion as Descripcion FROM tracking_estado "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstado)

strCadena = "SELECT id_estado as Codigo,descripcion as Descripcion FROM tracking_estado "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstadoproceso)


strCadena = "SELECT * FROM view_listado_tracking where ruc='" & KEY_RUC & "' LIMIT 29 "
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
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 2100
           Grilla.ColWidth(4) = 3500
           Grilla.ColWidth(5) = 3500
           Grilla.ColWidth(6) = 2200
           Grilla.ColWidth(7) = 1200
           Grilla.ColWidth(8) = 2800
       Next
        cabecera = "ID_DETALLE" & vbTab & "FECHA" & vbTab & "HORA " & vbTab & "COMPROBANTE " & vbTab & "CLIENTE " & vbTab & "DIRECCION" & vbTab & "SUCURSAL" & vbTab & "TOTAL" & vbTab & "ESTADO"
        Grilla.AddItem cabecera
         For k = 1 To 8
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
      For i = 0 To rst.RecordCount - 1
            
            Fila = rst("id_venta") & vbTab & rst("fecha_emision") & vbTab & rst("hora") & vbTab & rst("documento") & vbTab & rst("ncliente") & vbTab & rst("direccion") & vbTab & rst("descripcion") & vbTab & Format(rst("total"), "###0.00") & vbTab & rst("estado")
            Grilla.AddItem Fila
            in_color = &H8080FF
            Select Case rst("estado")
                Case "ENTREGADA COMPLETO"
                    in_color = &HC000&
                Case "ENTREGA PARCIAL"
                    in_color = &H80FF&
                 Case "ENVIADO"
                    in_color = &HFFFF80
                 Case "FACTURATADO LISTO A ENVIAR"
                    in_color = &HFFFF&
                Case "PENDIENTE DE ENTREHA"
                    in_color = &HC0&
                
                    
            End Select
            
            For k = 8 To 8
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = in_color
        Next k
        
            
            rst.MoveNext
    Next i
  
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub
Public Sub llenarGrid_tracking(ByVal Grilla As MSHFlexGrid, ByVal in_venta As String)
On Error GoTo salir
strCadena = "SELECT * FROM view_tracking where id_venta='" & in_venta & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    
    Grilla.Rows = 0
    
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 2800
           Grilla.ColWidth(4) = 3500
           Grilla.ColWidth(5) = 3000
           
       
       Next
        cabecera = "ID_DETALLE" & vbTab & "FECHA" & vbTab & "HORA " & vbTab & "DESCRIPCION ESTADO " & vbTab & "OBSERVACION" & vbTab & "OPERADOR"
        Grilla.AddItem cabecera
         For k = 1 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
      For i = 0 To rst.RecordCount - 1
            
            Fila = rst("id_venta") & vbTab & Format(rst("fecha_registro"), "dd-mm-YYYY") & vbTab & Format(rst("hora_registro"), "HH:mm am/pm") & vbTab & rst("descripcion") & vbTab & rst("observacion") & vbTab & rst("nombre_completo")
            Grilla.AddItem Fila
            in_color = &H8080FF
            Select Case rst("id_estado")
                Case "05"
                    in_color = &H80FF80
                Case "04"
                    in_color = &H80FF&
            End Select
            
        For k = 3 To 5
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = in_color
        Next k
        
            
            rst.MoveNext
    Next i
  
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub
Public Sub llenarGrid_envios(ByVal Grilla As MSHFlexGrid, ByVal in_venta As String)
On Error GoTo salir
strCadena = "SELECT * FROM view_tracking where id_venta='" & in_venta & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    
    Grilla.Rows = 0
    
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 2800
           Grilla.ColWidth(4) = 3500
           Grilla.ColWidth(5) = 3000
           
       
       Next
        cabecera = "ID_DETALLE" & vbTab & "FECHA" & vbTab & "HORA " & vbTab & "DESCRIPCION ESTADO " & vbTab & "OBSERVACION" & vbTab & "OPERADOR"
        Grilla.AddItem cabecera
         For k = 1 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
      For i = 0 To rst.RecordCount - 1
            
            Fila = rst("id_venta") & vbTab & Format(rst("fecha_registro"), "dd-mm-YYYY") & vbTab & Format(rst("hora_registro"), "HH:mm am/pm") & vbTab & rst("descripcion") & vbTab & rst("observacion") & vbTab & rst("nombre_completo")
            Grilla.AddItem Fila
            in_color = &H8080FF
            Select Case rst("id_estado")
                Case "05"
                    in_color = &H80FF80
                Case "04"
                    in_color = &H80FF&
            End Select
            
        For k = 3 To 5
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = in_color
        Next k
        
            
            rst.MoveNext
    Next i
  
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Private Sub HfdPersona_DblClick()
If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
    If Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 8)) <> "ENTREGADA COMPLETO" Then
    
    Me.frmestado.Visible = True
    Me.DtcEstadoproceso.SetFocus
    End If
End If
End Sub

Private Sub HfdPersona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtObservacion)
End If
End Sub

Private Sub HfdPersona_SelChange()
If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
   Me.cmdEnvios.Enabled = True
   Me.cmdtracking.Enabled = True
Else
    Me.cmdEnvios.Enabled = False
   Me.cmdtracking.Enabled = False
End If
End Sub

Private Sub TxtApellido_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_listado_tracking where ruc='" & KEY_RUC & "'  and ncliente Like '%" & Trim(Me.TxtApellido.Text) & "%' "
    Call Me.llenarGrid(Me.HfdPersona)
End If
End Sub

Private Sub TxtDocumento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_listado_tracking where ruc='" & KEY_RUC & "'  and documento Like '%" & Trim(Me.TxtDocumento.Text) & "%' "
    Call Me.llenarGrid(Me.HfdPersona)
End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_listado_tracking where ruc='" & KEY_RUC & "'  and id_cliente Like '%" & Trim(Me.txtRuc.Text) & "%' "
    Call Me.llenarGrid(Me.HfdPersona)
End If

End Sub

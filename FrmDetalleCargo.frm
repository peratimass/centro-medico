VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleCargo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmturno 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   5775
      Begin VitekeySoft.ChameleonBtn cmdsave 
         Height          =   495
         Left            =   1920
         TabIndex        =   13
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "PROCESAR"
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
         MICON           =   "FrmDetalleCargo.frx":0000
         PICN            =   "FrmDetalleCargo.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txthora_inicio 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1920
         MaxLength       =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1365
      End
      Begin VB.TextBox txthora_final 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1920
         MaxLength       =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1365
      End
      Begin VB.TextBox txtid_detalle 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Left            =   4560
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSDataListLib.DataCombo DtcTurno 
         Height          =   330
         Left            =   1920
         TabIndex        =   9
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   ""
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
      Begin VitekeySoft.ChameleonBtn cmdcerrar 
         Height          =   255
         Left            =   5400
         TabIndex        =   14
         Top             =   120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "FrmDetalleCargo.frx":3664
         PICN            =   "FrmDetalleCargo.frx":3680
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA ENTRADA :"
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
         Height          =   210
         Left            =   345
         TabIndex        =   12
         Top             =   960
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA SALIDA  :"
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
         Height          =   210
         Left            =   480
         TabIndex        =   11
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TURNO :"
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
         Height          =   210
         Left            =   960
         TabIndex        =   10
         Top             =   480
         Width           =   645
      End
   End
   Begin VB.TextBox TxtDescripcion 
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   1455
      MaxLength       =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4725
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgLinea 
      Height          =   1575
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2778
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
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
   Begin MSDataListLib.DataCombo dtcArea 
      Height          =   330
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   ""
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
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   120
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleCargo.frx":6534
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleCargo.frx":6850
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleCargo.frx":6CB0
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleCargo.frx":7110
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleCargo.frx":742C
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleCargo.frx":788C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleCargo.frx":7BA8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleCargo.frx":8008
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleCargo.frx":8468
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleCargo.frx":8D48
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleCargo.frx":9064
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleCargo.frx":9380
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   4260
      TabIndex        =   15
      Top             =   5490
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1875
      _CBHeight       =   870
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   810
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1429
         ButtonWidth     =   1402
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Grabar"
               Key             =   "(Grabar)"
               Object.ToolTipText     =   "Grabar Ctrl+G"
               ImageKey        =   "(Grabar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "(Cancelar)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Cancelar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   495
      Left            =   6240
      TabIndex        =   17
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetalleCargo.frx":969C
      PICN            =   "FrmDetalleCargo.frx":96B8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdupdate 
      Height          =   495
      Left            =   6240
      TabIndex        =   18
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "UPDATE"
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
      MICON           =   "FrmDetalleCargo.frx":C98E
      PICN            =   "FrmDetalleCargo.frx":C9AA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmddelete 
      Height          =   550
      Left            =   6240
      TabIndex        =   19
      Top             =   4200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   "DELETE"
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
      MICON           =   "FrmDetalleCargo.frx":FC80
      PICN            =   "FrmDetalleCargo.frx":FC9C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblDescripcion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION :"
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
      Left            =   300
      TabIndex        =   4
      Top             =   165
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AREA :"
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
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   6735
      Left            =   0
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "FrmDetalleCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim StrCodTabla As String
Dim strCodLinea As String

Private Sub cmdcerrar_Click()
Me.frmturno.Visible = False

End Sub

Private Sub cmddelete_Click()
in_cargo = FrmCargosPersonal.HfgLinea.TextMatrix(FrmCargosPersonal.HfgLinea.Row, 0)
If MsgBox("Esta seguro de ELIMINAR este turno", vbYesNo + vbQuestion) = vbYes Then
    strCadena = "DELETE FROM persona_cargo_turno WHERE id_detalle='" & Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) & "'"
    CnBd.Execute (strCadena)

Call llenar_turnos(Me.HfgLinea, in_cargo)
End If
End Sub

Private Sub cmdnuevo_Click()
Me.frmturno.Visible = True
Me.Txtid_detalle.Text = ""
Me.txthora_inicio.Text = ""
Me.txthora_final.Text = ""
Call Resalta(Me.txthora_inicio)

End Sub

Private Sub cmdsave_Click()
Dim in_cargo As String
    in_cargo = FrmCargosPersonal.HfgLinea.TextMatrix(FrmCargosPersonal.HfgLinea.Row, 0)
    
    INICIO = Format(Me.txthora_inicio.Text, "hh:mm:ss")
    final = Format(Me.txthora_final.Text, "hh:mm:ss")
    
    
    If Round((TimeValue(final) - TimeValue(INICIO)) * 24, 3) > 0 And Val(Me.Txtid_detalle.Text) < 1 Then
       strCadena = "INSERT INTO persona_cargo_turno(`id_cargo`,`id_turno`,`hora_inicio`,`hora_final`,`ruc`)VALUES('" & in_cargo & "','" & Me.DtcTurno.BoundText & "','" & INICIO & "','" & final & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
    Else
       strCadena = "UPDATE persona_cargo_turno SET hora_inicio='" & INICIO & "',hora_final='" & final & "' WHERE id_cargo='" & in_cargo & "' and id_turno='" & Me.DtcTurno.BoundText & "' and ruc='" & KEY_RUC & "'"
       CnBd.Execute (strCadena)
    End If
    Me.frmturno.Visible = False
    
    Call llenar_turnos(Me.HfgLinea, in_cargo)
    
End Sub

Private Sub cmdupdate_Click()
If Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) > 0 Then
   Call load_turno(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0))
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub
Private Sub load_turno(ByVal in_detalle As String)
 Me.Txtid_detalle.Text = Val(in_detalle)
 Me.txthora_inicio.Text = Format(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 2), "HH:mm")
 Me.txthora_final.Text = Format(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 3), "HH:mm")
 Me.frmturno.Visible = True
End Sub
Private Sub Form_Load()
CenterForm Me
Me.Top = 1200

strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' and id_sucursal='0'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcArea)
Me.DtcArea.BoundText = 0

strCadena = "SELECT id_turno as Codigo,descripcion as Descripcion FROM turno WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTurno)









  Select Case FrmCargosPersonal.Procedencia
    Case modificar
      Call LLENA(FrmCargosPersonal.HfgLinea.TextMatrix(FrmCargosPersonal.HfgLinea.Row, 0))
  End Select
End Sub

Private Sub LLENA(ByVal in_cargo As String)
  strCadena = "SELECT * FROM persona_cargos WHERE id_cargo='" & in_cargo & "' and  id_empresa='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
     Me.DtcArea.BoundText = rst("id_alm")
     Me.TxtDescripcion.Text = rst("descripcion")
     Call llenar_turnos(Me.HfgLinea, in_cargo)
  End If
  
End Sub
Private Sub funcion_cargo(ByVal id_cargo As String)
strCadena = "SELECT * FROM funciones_empresa WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
  
        strCadena = "SELECT count(*) FROM cargo_funcion WHERE id_cargo='" & id_cargo & "' AND id_funcion='" & rst("id_funcion") & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstT(strCadena)
        If rstT(0) < 1 Then
            strCadena = "INSERT INTO cargo_funcion(id_cargo,id_funcion,ruc)VALUES('" & id_cargo & "','" & rst("id_funcion") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        End If
        rst.MoveNext
    Next i
End If
End Sub

Private Sub Save()
Dim strcodigo As String
  If TxtDescripcion.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    Select Case FrmCargosPersonal.Procedencia
      Case nuevo
            strCadena = "SELECT * FROM persona_cargos WHERE id_empresa ='" & KEY_RUC & "' ORDER BY id_cargo DESC LIMIT 0,1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                strcodigo = formato_item(Val(rst("id_cargo")) + 1, 5)
            Else
                strcodigo = formato_item(4, 5)
            End If
            strCadena = "INSERT INTO persona_cargos (id_cargo,id_alm,descripcion,ruc,id_empresa) VALUES('" & strcodigo & "','" & Me.DtcArea.BoundText & "','" & Trim(Me.TxtDescripcion.Text) & "','si','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            Call funcion_cargo(strcodigo)
            Call FrmCargosPersonal.actualizar
            Unload Me
            Exit Sub
      Case modificar
        strCadena = "UPDATE persona_cargos SET id_alm='" & Me.DtcArea.BoundText & "', descripcion='" & TxtDescripcion.Text & "' WHERE id_cargo = '" & FrmCargosPersonal.HfgLinea.TextMatrix(FrmCargosPersonal.HfgLinea.Row, 0) & "' AND id_empresa='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        Call funcion_cargo(FrmCargosPersonal.HfgLinea.TextMatrix(FrmCargosPersonal.HfgLinea.Row, 0))
        Call FrmCargosPersonal.actualizar
        Unload Me
        Exit Sub
    End Select
  End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
    Case KEY_CANCEL
        Unload Me
  End Select
  Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub


Private Sub TxtDescripcion_Change()
strCadena = "SELECT id_cargo as Codigo,descripcion as Descripcion FROM persona_cargos WHERE ruc='si' and id_empresa='" & KEY_RUC & "' AND descripcion LIKE '%" & Trim(Me.TxtDescripcion.Text) & "%' ORDER BY descripcion"
Call llenar_grid(Me.HfgLinea, strCadena)
End Sub
Private Sub llenar_grid(ByVal Grilla As MSHFlexGrid, ByVal Cadena As String)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Grilla.Clear
    Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = True
    Exit Sub

End If
  
  

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 3500
        Next
        cabecera = "ID CARGO" & vbTab & "DESCRIPCION"
        Grilla.AddItem cabecera
         For k = 0 To 1
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("Codigo") & vbTab & rst("descripcion")
            Grilla.AddItem Fila
        
            rst.MoveNext
        Next i
        
  If FrmCargosPersonal.Procedencia = modificar Then
        Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = True
  End If
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub
Private Sub llenar_turnos(ByVal Grilla As MSHFlexGrid, ByVal in_cargo As String)
On Error GoTo salir
strCadena = "SELECT * FROM view_turno_cargo WHERE id_cargo='" & in_cargo & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    
    Exit Sub

End If
  
  
  
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2500
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
        Next
        cabecera = "CODIGO" & vbTab & "TURNO" & vbTab & "HORA INICIO" & vbTab & "HORA FIN"
        Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_detalle") & vbTab & rst("turno") & vbTab & Format(rst("hora_inicio"), "HH:mm") & vbTab & Format(rst("hora_final"), "HH:mm")
            Grilla.AddItem Fila
            
            rst.MoveNext
        Next i
        
  
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub


Private Sub txtdescripcion_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)



End Sub






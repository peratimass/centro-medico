VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form FrmCentroCostosDetalle 
   Caption         =   "MANTENIMIENTO CENTRO COSTOS"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   8385
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Plan Contable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   480
      TabIndex        =   11
      Top             =   1200
      Width           =   7575
      Begin MSDataListLib.DataCombo DtcPlanCOntable 
         Height          =   315
         Left            =   2400
         TabIndex        =   13
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Plan:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1125
         TabIndex        =   12
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.TextBox txtDetalle 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2760
      MaxLength       =   100
      TabIndex        =   4
      Top             =   720
      Width           =   4965
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Amarres"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   7575
      Begin VB.TextBox TxtCuentahaber 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   8
         Top             =   840
         Width           =   1005
      End
      Begin VB.TextBox txtCuentaDebe 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   7
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label lelehaber 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   10
         Top             =   900
         Width           =   4425
      End
      Begin VB.Label lblDebe 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   9
         Top             =   360
         Width           =   4425
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amarre al Haber:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   105
         TabIndex        =   6
         Top             =   840
         Width           =   1545
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amarre al Debe:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   135
         TabIndex        =   5
         Top             =   360
         Width           =   1485
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Centro de Costo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      Begin VB.TextBox txtCodcentrocosto 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   3
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   105
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   5400
      Top             =   3720
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
            Picture         =   "FrmCentroCostosDetalle.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostosDetalle.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostosDetalle.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostosDetalle.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostosDetalle.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostosDetalle.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostosDetalle.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostosDetalle.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostosDetalle.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostosDetalle.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostosDetalle.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCentroCostosDetalle.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   6000
      TabIndex        =   14
      Top             =   3720
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1995
      _CBHeight       =   870
      _Version        =   "6.7.9782"
      Child1          =   "TlbAcciones"
      MinHeight1      =   810
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   1875
         _ExtentX        =   3307
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
   Begin VB.Shape ShpProducto 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   4575
      Left            =   240
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "FrmCentroCostosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrCodTabla As String
Public Procedencia As EnumProcede
Public debe_haber As debe_haber
Private Sub Form_Load()
CenterForm Me
Me.Width = 8505
Me.Top = 500
Me.Height = 5340
strCadena = "SELECT id_plancontable as Codigo, pc_descripcion as Descripcion FROM plan_contable "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPlanCOntable)
  


 Select Case FrmCentroCostos.Procedencia
    Case nuevo
        strCadena = "SELECT id_costo FROM centro_costos ORDER BY id_costo DESC"
        Call ConfiguraRst(strCadena)
        StrCodTabla = GeneraCodigo(5)
        Me.txtCodcentrocosto.Text = StrCodTabla
    Case Modificar
      Call LLENA
  End Select

End Sub

Private Sub LLENA()
  FrmCentroCostos.HfgAgrupacionCuenta.col = 0
  StrCodTabla = FrmCentroCostos.HfgAgrupacionCuenta.Text
 strCadena = "SELECT * FROM  centro_costos WHERE id_costo='" & Trim(StrCodTabla) & "'"
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    Me.txtCodcentrocosto.Text = StrCodTabla
    Me.txtdetalle.Text = rst("descripcion")
    Me.txtCuentaDebe.Text = rst("id_debe")
    Me.TxtCuentahaber.Text = rst("id_haber")
    Me.lblDebe.Caption = nombre_cuenta(rst("id_debe"), Trim(rst("id_plan")))
    Me.lelehaber.Caption = nombre_cuenta(rst("id_haber"), Trim(rst("id_plan")))
    Set rst = Nothing
    
 Else
    Exit Sub
 End If
End Sub
Private Sub Save()
  If Me.txtdetalle.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    Select Case FrmCentroCostos.Procedencia
      Case nuevo
         strCadena = "INSERT INTO centro_costos VALUES ('" & Trim(Me.txtCodcentrocosto.Text) & "','" & Trim(Me.txtdetalle.Text) & "','" & Trim(Me.txtCuentaDebe.Text) & "', " & _
         "'" & Trim(Me.TxtCuentahaber.Text) & "','" & Trim(Me.DtcPlanCOntable.BoundText) & "')"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        FrmCentroCostos.actualizar
        Unload Me
      Case Modificar
        strCadena = "UPDATE centro_costos SET descripcion='" & Trim(Me.txtdetalle.Text) & "',id_debe='" & Trim(Me.txtCuentaDebe.Text) & "', " & _
        "id_haber='" & Trim(Me.TxtCuentahaber.Text) & "',id_plan='" & Trim(Me.DtcPlanCOntable.BoundText) & "' WHERE clinea = '" & StrCodTabla & "'"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Unload Me
    End Select
  End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
    Case KEY_CANCEL
        Unload Me
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub txtCuentaDebe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM plan_contable_det WHERE pc_codigo='" & Trim(Me.txtCuentaDebe.Text) & "' AND id_plancontable='" & Trim(Me.DtcPlanCOntable.BoundText) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.lblDebe.Caption = rst("plan_des")
        Set rst = Nothing
    Else
        Set rst = Nothing
        Procedencia = Selecionar
        debe_haber = debe
        FrmPlanContableCuentas.Show
    End If
End If
End Sub

Private Sub TxtCuentahaber_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM plan_contable_det WHERE pc_codigo='" & Trim(Me.TxtCuentahaber.Text) & "' AND id_plancontable='" & Trim(Me.DtcPlanCOntable.BoundText) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.lelehaber.Caption = rst("plan_des")
        Set rst = Nothing
    Else
        Set rst = Nothing
        Procedencia = Selecionar
        debe_haber = HABER
        FrmPlanContableCuentas.Show
    End If
End If

End Sub

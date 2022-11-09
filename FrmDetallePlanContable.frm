VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmDetallePlanContable 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detalles"
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
      Height          =   1455
      Left            =   4560
      TabIndex        =   31
      Top             =   4320
      Width           =   3135
      Begin VB.CheckBox ChkCtaAnalizada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Debe ser Analaizada"
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
         Height          =   330
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   2850
      End
      Begin VB.CheckBox ChkCtaretencion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cuenta de Retencion"
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
         Height          =   330
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   2730
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5520
      MaxLength       =   20
      TabIndex        =   29
      Top             =   3240
      Width           =   1365
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cuentas de Amarre"
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
      Height          =   1455
      Left            =   4560
      TabIndex        =   26
      Top             =   2760
      Width           =   3135
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   960
         MaxLength       =   20
         TabIndex        =   30
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Haber:"
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
         Left            =   150
         TabIndex        =   28
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debe:"
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
         Left            =   180
         TabIndex        =   27
         Top             =   480
         Width           =   585
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo de Analisis"
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
      Height          =   3015
      Left            =   2160
      TabIndex        =   21
      Top             =   2760
      Width           =   2295
      Begin VB.OptionButton OptSoloDetalle 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solo Detalle."
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
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton OptuentasBanco 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cuenta de Banco"
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
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton OptPorDocumentos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Por Documentos"
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
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton OptSinAnalisis 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sin Analisis"
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
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.OptionButton Optayor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mayor"
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
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   5280
      Width           =   1215
   End
   Begin VB.OptionButton OptOrden 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Orden"
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
      Height          =   375
      Left            =   480
      TabIndex        =   19
      Top             =   4920
      Width           =   1215
   End
   Begin VB.OptionButton OptFuncion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funcion"
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
      Height          =   375
      Left            =   480
      TabIndex        =   18
      Top             =   4560
      Width           =   1215
   End
   Begin VB.OptionButton OptNaturaleza 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Naturaleza"
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
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo Cuenta"
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
      Height          =   3015
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   1815
      Begin VB.OptionButton OptActivo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Activo"
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
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptPasivo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pasivo"
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
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton OptResultado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resultado"
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
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NIVEL DE CUENTA"
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
      Height          =   855
      Left            =   480
      TabIndex        =   9
      Top             =   1440
      Width           =   5175
      Begin VB.OptionButton OptRegistro 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Registro"
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
         Height          =   375
         Left            =   3000
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OptSubcuenta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sub-Cuenta"
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
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptBalance 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Balance"
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
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtDescripcion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2040
      MaxLength       =   100
      TabIndex        =   6
      Top             =   600
      Width           =   5085
   End
   Begin VB.TextBox TxtCuenta 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   5
      Top             =   240
      Width           =   2325
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   4560
      Top             =   6600
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
            Picture         =   "FrmDetallePlanContable.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePlanContable.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePlanContable.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePlanContable.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePlanContable.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePlanContable.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePlanContable.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePlanContable.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePlanContable.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePlanContable.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePlanContable.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetallePlanContable.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   5685
      TabIndex        =   3
      Top             =   6120
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   2955
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
         TabIndex        =   4
         Top             =   30
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   1429
         ButtonWidth     =   1482
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
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
               Caption         =   "&Enlaces"
               Key             =   "(Enlaces)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Red)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo dtcPlancontable 
      Height          =   315
      Left            =   2040
      TabIndex        =   8
      Top             =   960
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLAN CONTABLE :"
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
      Left            =   525
      TabIndex        =   7
      Top             =   960
      Width           =   1185
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   3495
      Left            =   120
      Top             =   2520
      Width           =   8535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta para Cierre:"
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
      Height          =   480
      Left            =   4680
      TabIndex        =   2
      Top             =   5760
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION :"
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
      Left            =   675
      TabIndex        =   1
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NRO CUENTA :"
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
      Left            =   915
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape ShpProducto 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   2295
      Left            =   120
      Top             =   120
      Width           =   8535
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   7095
      Left            =   0
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "FrmDetallePlanContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NBalance As Integer, NSubcuenta As Integer, NRegistro As Integer
Dim Tactivo As Integer, Tpasivo As Integer, Tnaturaleza As Integer, TResultado As Integer, Tfuncion As Integer, TOrden As Integer, Tmayor As Integer
Private Sub Form_Activate()
CenterForm Me
Me.Top = 800
End Sub

Private Sub Form_Load()

 strCadena = "SELECT id_plancontable AS Codigo,pc_descripcion as Descripcion FROM plan_contable ORDER BY id_plancontable ASC"
 Call ConfiguraRst(strCadena)
 Call LlenaDataCombo(Me.dtcPlancontable)
 Set rst = Nothing
 
 Select Case FrmPlanContableCuentas.Procedencia
    
    Case modificar
        Call LLENA
 End Select
End Sub


Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
    Call Save
    Case KEY_EXIT
      Unload Me
  End Select
  Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub
Private Sub Save()
Dim retencion As String
Dim analiza As String
Dim agrupa As String
Dim idCuenta As String

  If TxtDescripcion.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    Select Case FrmPlanContableCuentas.Procedencia
      Case nuevo
        strCadena = "SELECT * FROM con_cuentacontable ORDER BY Id DESC LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            idCuenta = Trim(Mid(rst("Id"), 1, Len(rst("Id")) - 3)) & Trim(str(Val(Right(rst("Id"), 3)) + 1))
            strCadena = "INSERT INTO con_cuentacontable(`Id`,`IdEmpresaSis`,`IdSucursal`,`Ejercicio`,`NroCuenta`,`Descripcion`,`IndMovimiento`,`Activo`,`UsuarioCrea`,`FechaCrea`)VALUES " & _
            "('" & idCuenta & "','" & KEY_RUC & "','','" & Year(KEY_FECHA) & "','" & Trim(Me.TxtCuenta.Text) & "','" & Trim(Me.TxtDescripcion.Text) & "','1','1','" & KEY_USUARIO & "','" & KEY_FECHA & "')"
            CnBd.Execute (strCadena)
        End If
        
        
         
        Call FrmPlanContableCuentas.actualizar
        Unload Me
        Exit Sub
             
      Case modificar
        strCadena = "UPDATE con_cuentacontable SET NroCuenta='" & Trim(Me.TxtCuenta.Text) & "',Descripcion='" & Trim(Me.TxtDescripcion.Text) & "', " & _
        "UsuarioModifica='" & KEY_USUARIO & "',FechaModifica='" & KEY_FECHA & "' WHERE id_plancontable_det = '" & Val(FrmPlanContableCuentas.HfgPlanContable.TextMatrix(FrmPlanContableCuentas.HfgPlanContable.Row, 0)) & "'"
        CnBd.Execute (strCadena)
         
        Call FrmPlanContableCuentas.actualizar
        Unload Me
    End Select
  End If
End Sub

Private Sub LLENA()
Dim cod_plan As String
Dim cod_cuenta As String
  cod_cuenta = FrmPlanContableCuentas.HfgPlanContable.TextMatrix(FrmPlanContableCuentas.HfgPlanContable.Row, 0)
  cod_plan = FrmPlanContableCuentas.dtcPlancontable.BoundText
  strCadena = "SELECT * FROM plan_contable_det WHERE id_plancontable_det = '" & Trim(cod_cuenta) & "' AND id_plancontable='" & Trim(cod_plan) & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
    StrCodTabla = rst("id_plancontable_det")
    Me.TxtCuenta.Text = rst("pc_codigo")
    Me.TxtDescripcion.Text = rst("plan_des")
    
    
  End If

End Sub

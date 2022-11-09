VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmAgrupacionCuentas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agrupacion Cuentas"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtPlanContable 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   6990
      Width           =   4935
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3240
      Top             =   2550
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAgrupacionCuentas.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAgrupacionCuentas.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAgrupacionCuentas.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAgrupacionCuentas.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAgrupacionCuentas.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAgrupacionCuentas.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAgrupacionCuentas.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAgrupacionCuentas.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAgrupacionCuentas.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgAgrupacionCuenta 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   11668
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   3945
      Left            =   6885
      TabIndex        =   2
      Top             =   120
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   6959
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   3945
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   5460
         Left            =   30
         TabIndex        =   3
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   9631
         ButtonWidth     =   1482
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Nuevo   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Modificar"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   405
      TabIndex        =   4
      Top             =   7110
      Width           =   705
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   120
      Top             =   6870
      Width           =   6735
   End
End
Attribute VB_Name = "FrmAgrupacionCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub Form_Activate()
strCadena = "SELECT agruCuenta_cod as Código,agruCuenta_des as Descripcion, agruCeunta_abrev " & _
  " as Resumen FROM AgrupaCuentas ORDER BY agruCuenta_des ASC"
  Call llenarGridME(Me.HfgAgrupacionCuenta, Me)
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
End Sub
Private Sub Reziseform(ByVal formulario As Form)
'CenterForm formulario
'formulario.Width = 8160
'formulario.Height = 8385
'formulario.Top = 400
End Sub

Private Sub HfgAgrupacionCuenta_Click()
 If Me.HfgAgrupacionCuenta.Rows > 0 Then
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    TlbAcciones.Buttons(KEY_DELETE).Enabled = True
    
   End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_NEW
      Procedencia = Nuevo
      FrmDetalleAgrupacionCuentas.Show
    Case KEY_UPDATE
      Procedencia = Modificar
      FrmDetalleAgrupacionCuentas.Show
    Case KEY_DELETE
      If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Me.HfgAgrupacionCuenta.col = 0
        strCadena = "DELETE FROM AgrupaCuentas WHERE agruCuenta_cod='" & Me.HfgAgrupacionCuenta.text & "'"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Form_Activate
      End If
    Case "(Salir)"
      Unload Me
  End Select
End Sub
Sub llenarGridME(ByVal Grilla As MSHFlexGrid, ByVal formulario As Form)
On Error GoTo SALIR
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.Rows = rst.RecordCount
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 700
  Grilla.ColWidth(1) = 3500
  Grilla.ColWidth(2) = 2000
'Call DarFormato1(Grilla, 0)
Grilla.Refresh
  formulario.TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
  formulario.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  Exit Sub
SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

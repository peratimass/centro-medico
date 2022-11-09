VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleTipoProducto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Tipo de Productos"
   ClientHeight    =   2745
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   4260
   ControlBox      =   0   'False
   Icon            =   "FrmDetalleTipoProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4260
   Begin MSDataListLib.DataCombo DtcLinea 
      Height          =   315
      Left            =   1523
      TabIndex        =   0
      Top             =   465
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox TxtDescripcion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1523
      MaxLength       =   20
      TabIndex        =   1
      Top             =   945
      Width           =   2325
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   1560
      Top             =   1950
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
            Picture         =   "FrmDetalleTipoProducto.frx":0442
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoProducto.frx":075E
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoProducto.frx":0BBE
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoProducto.frx":101E
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoProducto.frx":133A
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoProducto.frx":179A
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoProducto.frx":1AB6
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoProducto.frx":1F16
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoProducto.frx":2376
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoProducto.frx":2C56
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoProducto.frx":2F72
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoProducto.frx":328E
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   2190
      TabIndex        =   4
      Top             =   1680
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
         TabIndex        =   5
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
   Begin VB.Label LblLinea 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Línea :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   413
      TabIndex        =   3
      Top             =   525
      Width           =   525
   End
   Begin VB.Label LblDescripcion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción :"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   413
      TabIndex        =   2
      Top             =   997
      Width           =   975
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   203
      Top             =   225
      Width           =   3855
   End
End
Attribute VB_Name = "FrmDetalleTipoProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StrCodTabla As String
Dim Linea As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
CenterForm Me
  strCadena = "SELECT clinea as Codigo, sdescripcion as Descripcion FROM  " & _
  " linea ORDER BY sdescripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(DtcLinea)
  Select Case FrmTipoProducto.Procedencia
    Case Modificar
      Call LLENA
  End Select
End Sub

Private Sub LLENA()
  FrmTipoProducto.HfdGrilla.col = 0
  StrCodTabla = FrmTipoProducto.HfdGrilla.Text
  FrmTipoProducto.HfdGrilla.col = 1
  TxtDescripcion.Text = FrmTipoProducto.HfdGrilla.Text
  FrmTipoProducto.HfdGrilla.col = 2
  DtcLinea.Text = FrmTipoProducto.HfdGrilla.Text
End Sub

Private Sub Save()
  If TxtDescripcion.Text = "" Or DtcLinea.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    Linea = Replace(DtcLinea.BoundText, "'", "''")
    Select Case FrmTipoProducto.Procedencia
      Case nuevo
        strCadena = "SELECT ctipoProducto FROM tipoProducto ORDER BY ctipoProducto DESC"
        Call ConfiguraRst(strCadena)
        strCadena = "INSERT INTO tipoProducto (ctipoProducto,sdescripcion,clinea)  " & _
        " VALUES ('" & GeneraCodigo(4) & "','" & TxtDescripcion.Text & "','" & Linea & "')"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Unload Me
      Case Modificar
        strCadena = "UPDATE tipoProducto SET sdescripcion= " & _
        " '" & TxtDescripcion.Text & "', clinea='" & Linea & "' WHERE ctipoproducto " & _
        " = '" & StrCodTabla & "'"
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

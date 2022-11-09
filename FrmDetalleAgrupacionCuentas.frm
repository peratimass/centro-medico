VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleAgrupacionCuentas 
   Caption         =   "Form2"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4260
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   2445
   ScaleWidth      =   4260
   Begin VB.TextBox TxtDescripcion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1455
      MaxLength       =   20
      TabIndex        =   1
      Top             =   240
      Width           =   2325
   End
   Begin VB.TextBox TxtResumen 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1455
      MaxLength       =   8
      TabIndex        =   0
      Top             =   720
      Width           =   1125
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   1320
      Top             =   1605
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
            Picture         =   "FrmDetalleAgrupacionCuentas.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleAgrupacionCuentas.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleAgrupacionCuentas.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleAgrupacionCuentas.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleAgrupacionCuentas.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleAgrupacionCuentas.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleAgrupacionCuentas.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleAgrupacionCuentas.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleAgrupacionCuentas.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleAgrupacionCuentas.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleAgrupacionCuentas.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleAgrupacionCuentas.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
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
   Begin VB.Label LblDescripcion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción :"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   285
      Width           =   975
   End
   Begin VB.Label LblResumen 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resumen :"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   765
      Width           =   795
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   1035
      Left            =   120
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "FrmDetalleAgrupacionCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrCodTabla As String
Dim StrCodAgrupacion As String
Private Sub Form_Load()
Reziseform Me
Select Case FrmAgrupacionCuentas.Procedencia
           
    Case Modificar
      Call LLENA
  End Select
End Sub
Private Sub Reziseform(ByVal formulario As Form)
CenterForm formulario
formulario.Width = 4695
formulario.Height = 2660
formulario.Top = 400
End Sub

Private Sub LLENA()
  FrmAgrupacionCuentas.HfgAgrupacionCuenta.col = 0
  StrCodTabla = FrmAgrupacionCuentas.HfgAgrupacionCuenta.Text
  FrmAgrupacionCuentas.HfgAgrupacionCuenta.col = 1
  txtdescripcion.Text = FrmAgrupacionCuentas.HfgAgrupacionCuenta.Text
  FrmAgrupacionCuentas.HfgAgrupacionCuenta.col = 2
  TxtResumen.Text = FrmAgrupacionCuentas.HfgAgrupacionCuenta.Text
End Sub

Private Sub Save()
  If txtdescripcion.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    If TxtResumen.Text = "" Then
      If Len(txtdescripcion.Text) > 8 Then
        TxtResumen.Text = Left(txtdescripcion.Text, 7) & "."
      Else
        TxtResumen.Text = txtdescripcion.Text
      End If
    End If
    Select Case FrmAgrupacionCuentas.Procedencia
      Case nuevo
        strCadena = "SELECT int_agruCuenta FROM AgrupaCuentas ORDER BY int_agruCuenta DESC"
        Call ConfiguraRst(strCadena)
        StrCodAgrupacion = GeneraCodigo(2)
        strCadena = "INSERT INTO AgrupaCuentas (agruCuenta_cod,agruCuenta_des,agruCeunta_abrev,int_agruCuenta)  " & _
        " VALUES ('" & StrCodAgrupacion & "','" & txtdescripcion.Text & "', " & _
        " '" & TxtResumen.Text & "','" & Val(StrCodAgrupacion) & "')"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Unload Me
      Case Modificar
        strCadena = "UPDATE AgrupaCuentas SET agruCuenta_des='" & txtdescripcion.Text & "',   " & _
        " agruCeunta_abrev='" & TxtResumen.Text & "' WHERE agruCuenta_cod = '" & StrCodTabla & "'"
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

Private Sub txtdescripcion_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub


Private Sub TxtResumen_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub


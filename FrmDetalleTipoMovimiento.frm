VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleTipoMovimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Tipo de Movimiento"
   ClientHeight    =   2760
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "FrmDetalleTipoMovimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   Begin VB.OptionButton OptSalida 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salida"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3353
      TabIndex        =   1
      Top             =   473
      Width           =   855
   End
   Begin VB.OptionButton OptEntrada 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Entrada"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   353
      TabIndex        =   0
      Top             =   473
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox TxtDescripcion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1553
      MaxLength       =   100
      TabIndex        =   2
      Top             =   953
      Width           =   2685
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   1560
      Top             =   1920
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
            Picture         =   "FrmDetalleTipoMovimiento.frx":0442
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoMovimiento.frx":075E
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoMovimiento.frx":0BBE
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoMovimiento.frx":101E
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoMovimiento.frx":133A
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoMovimiento.frx":179A
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoMovimiento.frx":1AB6
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoMovimiento.frx":1F16
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoMovimiento.frx":2376
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoMovimiento.frx":2C56
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoMovimiento.frx":2F72
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleTipoMovimiento.frx":328E
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2490
      Left            =   2700
      TabIndex        =   4
      Top             =   1695
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   4392
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1755
      _CBHeight       =   2490
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   2430
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   2430
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   4286
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
      Left            =   353
      TabIndex        =   3
      Top             =   1005
      Width           =   975
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   233
      Top             =   233
      Width           =   4215
   End
End
Attribute VB_Name = "FrmDetalleTipoMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrCodTabla As String
Dim strCodTipoMovimiento As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
CenterForm Me
  Select Case FrmTipoMovimiento.Procedencia
    Case modificar
      OptEntrada.Enabled = False
      OptSalida.Enabled = False
      Call LLENA
  End Select
End Sub

Private Sub LLENA()
  FrmTipoMovimiento.HfdGrilla.col = 0
  StrCodTabla = FrmTipoMovimiento.HfdGrilla.Text
  FrmTipoMovimiento.HfdGrilla.col = 1
  txtdescripcion.Text = FrmTipoMovimiento.HfdGrilla.Text
  If Left(StrCodTabla, 1) = "E" Then
    OptEntrada.Value = True
  Else
    OptSalida.Value = True
  End If
End Sub

Private Sub Save()
  If txtdescripcion.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    Select Case FrmTipoMovimiento.Procedencia
      Case nuevo
        If OptEntrada.Value = True Then
          Gencodigo = "E"
        Else
          Gencodigo = "S"
        End If
        
        strCadena = "SELECT int_tipomovimiento FROM Tipomovimiento ORDER BY int_tipomovimiento DESC"
        Call ConfiguraRst(strCadena)
        strCodTipoMovimiento = GeneraCodigo(2)
        
        strCadena = "INSERT INTO Tipomovimiento (cTipomovimiento, " & _
        " sdescripcionmovimiento,int_tipomovimiento) VALUES ('" & strCodTipoMovimiento & "', " & _
        " '" & txtdescripcion.Text & "','" & Val(Mid(strCodTipoMovimiento, 2, 2)) & "')"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Unload Me
      Case modificar
        strCadena = "UPDATE Tipomovimiento SET sdescripcionmovimiento " & _
        " ='" & txtdescripcion.Text & "' WHERE ctipomovimiento= '" & StrCodTabla & "'"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Unload Me
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

Private Sub txtdescripcion_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub

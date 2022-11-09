VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form FrmAcceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dar Acceso"
   ClientHeight    =   6915
   ClientLeft      =   8130
   ClientTop       =   6615
   ClientWidth     =   6570
   Icon            =   "FrmAcceso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   6570
   Begin VB.Frame FraUsuario 
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtIdUsuario 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Tag             =   "-1"
         Text            =   "0"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtUsuario 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Tag             =   "-1"
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Código :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   5295
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9340
      _Version        =   393216
      ForeColor       =   12582912
      FixedCols       =   0
      ForeColorFixed  =   12582912
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   360
      Top             =   960
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
            Picture         =   "FrmAcceso.frx":0442
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAcceso.frx":075E
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAcceso.frx":0BBE
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAcceso.frx":101E
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAcceso.frx":133A
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAcceso.frx":179A
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAcceso.frx":1AB6
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAcceso.frx":1F16
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAcceso.frx":2376
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAcceso.frx":2C56
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAcceso.frx":2F72
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAcceso.frx":328E
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2490
      Left            =   4710
      TabIndex        =   6
      Top             =   360
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
         TabIndex        =   7
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
End
Attribute VB_Name = "FrmAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrCodTabla As String
Public intIdUsuario As String
Private Sub Form_Load()
  Select Case FrmControlAccesos.Procedencia
    Case Nuevo
      Call LLENA
  End Select
  CenterForm Me
End Sub
Private Sub Form_Activate()

 ' FrmControlAccesos.HfdGrilla.Col = 0
  'Me.txtIdUsuario = FrmControlAccesos.HfdGrilla.Text
  strCadena = "SELECT AccesoUsuario.IdUsuario, AccesoMenu.Descripcion AS Menu, " & _
             "AccesoUsuario.Habilitado FROM AccesoMenu INNER JOIN " & _
             "AccesoUsuario ON AccesoMenu.IdOpcion = AccesoUsuario.IdOpcion WHERE AccesoUsuario.IdUsuario = '" & intIdUsuario & "'"
  
  Call ConfiguraRst(strCadena)
  Set HfdGrilla.Recordset = rst
     
  HfdGrilla.ColWidth(1) = 2500
  HfdGrilla.ColWidth(2) = 1100
  
  Set rst = Nothing
End Sub

Private Sub LLENA()
  FrmControlAccesos.HfdGrilla.col = 0
  StrCodTabla = FrmControlAccesos.HfdGrilla.Text
  FrmControlAccesos.HfdGrilla.col = 1
  Me.txtUsuario = FrmControlAccesos.HfdGrilla.Text
  FrmControlAccesos.HfdGrilla.col = 0
  Me.txtIdUsuario = FrmControlAccesos.HfdGrilla.Text
  Me.txtIdUsuario.Enabled = False
  Me.txtUsuario.Enabled = False
End Sub

Private Sub HfdGrilla_Click()
Dim reg As Integer
Dim rsthabilitar As New ADODB.Recordset
'rsthabilitar.Open "AccesoUsuario", CnBd, adOpenKeyset, adLockOptimistic

strCadena = "SELECT AccesoUsuario.Habilitado FROM AccesoMenu INNER JOIN " & _
             "AccesoUsuario ON AccesoMenu.IdOpcion = AccesoUsuario.IdOpcion WHERE AccesoUsuario.IdUsuario = '" & intIdUsuario & "'"
Call ConfiguraRst(strCadena)

'reg = Me.HfdGrilla.Index
Me.HfdGrilla.col = 2
If Me.HfdGrilla.col = 2 Then
    If rst!habilitado Then
        rst!habilitado = False
    Else
        rst!habilitado = True
    End If
    rst.Update
    Form_Activate
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_SAVE
      
    Case KEY_CANCEL
      Unload Me
  End Select
End Sub

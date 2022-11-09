VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDetalleLibrosAuxiliares 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Libros Auxiliares"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtDescripcion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1575
      MaxLength       =   20
      TabIndex        =   0
      Top             =   360
      Width           =   2325
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   1335
      Top             =   1200
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
            Picture         =   "FrmDetalleLibrosAuxiliares.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLibrosAuxiliares.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLibrosAuxiliares.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLibrosAuxiliares.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLibrosAuxiliares.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLibrosAuxiliares.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLibrosAuxiliares.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLibrosAuxiliares.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLibrosAuxiliares.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLibrosAuxiliares.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLibrosAuxiliares.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLibrosAuxiliares.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   2100
      TabIndex        =   1
      Top             =   930
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1482
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1875
      _CBHeight       =   840
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   780
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1376
         ButtonWidth     =   1349
         ButtonHeight    =   1376
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
      Left            =   375
      TabIndex        =   3
      Top             =   405
      Width           =   975
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   240
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "FrmDetalleLibrosAuxiliares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrCodLibroAuxilar As String * 2

Private Sub Form_Activate()
CenterForm Me
End Sub

Private Sub Reziseform(ByVal formulario As Form)
CenterForm formulario
formulario.Width = 4695
formulario.Height = 2625
formulario.Top = 400
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_SAVE
        Call Save
    Case KEY_CANCEL
        Unload Me
End Select
End Sub
Private Sub Save()
  If Me.TxtDescripcion.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    
    Select Case FrmLibrosAuxiliares.Procedencia
      Case Nuevo
        strCadena = "SELECT intLibrosauxiliares FROM LibrosAuxiliares ORDER BY intLibrosauxiliares DESC"
        Call ConfiguraRst(strCadena)
        
        StrCodLibroAuxilar = GeneraCodigo(2)
        strCadena = "INSERT INTO LibrosAuxiliares(cLibrosAuxiliares,sLibrosAuxiliares,intLibrosauxiliares)VALUES " & _
        "('" & StrCodLibroAuxilar & "','" & TxtDescripcion.Text & "','" & Val(StrCodLibroAuxilar) & "')"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Unload Me
      Case Modificar
        strCadena = "UPDATE LibrosAuxiliares SET sLibrosAuxiliares='" & TxtDescripcion.Text & "' WHERE cLibrosAuxiliares = '" & StrCodTabla & "'"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Unload Me
                
    End Select
  End If
End Sub

VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleUnidad 
   BorderStyle     =   0  'None
   Caption         =   "Detalle de Unidad"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   5175
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDetalleUnidad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtResumen 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1515
      MaxLength       =   8
      TabIndex        =   1
      Top             =   795
      Width           =   1125
   End
   Begin VB.TextBox TxtDescripcion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1515
      MaxLength       =   20
      TabIndex        =   0
      Top             =   315
      Width           =   2325
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   1380
      Top             =   1680
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
            Picture         =   "FrmDetalleUnidad.frx":030A
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUnidad.frx":0626
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUnidad.frx":0A86
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUnidad.frx":0EE6
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUnidad.frx":1202
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUnidad.frx":1662
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUnidad.frx":197E
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUnidad.frx":1DDE
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUnidad.frx":223E
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUnidad.frx":2B1E
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUnidad.frx":2E3A
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUnidad.frx":3156
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   3000
      TabIndex        =   4
      Top             =   1650
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
   Begin VB.Label LblResumen 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ABREVIATURA :"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   90
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label LblDescripcion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION :"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   210
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2715
      Left            =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "FrmDetalleUnidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrCodTabla As String
Dim StrCodUnidad As String
Private Sub Form_Load()
CenterForm Me
Me.Top = 1000
  Select Case FrmUnidad.Procedencia
    Case Modificar
      Call LLENA
  End Select
End Sub

Private Sub LLENA()
  StrCodTabla = FrmUnidad.HfgUnidad.TextMatrix(FrmUnidad.HfgUnidad.Row, 0)
  txtdescripcion.Text = FrmUnidad.HfgUnidad.TextMatrix(FrmUnidad.HfgUnidad.Row, 1)
  txtResumen.Text = FrmUnidad.HfgUnidad.TextMatrix(FrmUnidad.HfgUnidad.Row, 2)
End Sub

Private Sub Save()
  If txtdescripcion.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    If txtResumen.Text = "" Then
      If Len(txtdescripcion.Text) > 8 Then
        txtResumen.Text = Left(txtdescripcion.Text, 7) & "."
      Else
        txtResumen.Text = txtdescripcion.Text
      End If
    End If
    Select Case FrmUnidad.Procedencia
      Case nuevo
        strCadena = "SELECT * FROM unidad WHERE id_usu='" & KEY_RUC & "' ORDER BY id_und DESC"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            StrCodUnidad = formato_item(Val(rst("id_und") + 1), 5)
        Else
            StrCodUnidad = formato_item(1, 5)
        End If
        
        strCadena = "INSERT INTO unidad (id_und,descripcion,abreviatura,id_usu)  " & _
        " VALUES ('" & StrCodUnidad & "','" & txtdescripcion.Text & "', " & _
        " '" & txtResumen.Text & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
         
        
        Call FrmUnidad.actualizar
        Unload Me
      Case Modificar
        strCadena = "UPDATE unidad SET descripcion='" & txtdescripcion.Text & "',   " & _
        " abreviatura='" & txtResumen.Text & "' WHERE id_und = '" & StrCodTabla & "' AND id_usu='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        Call FrmUnidad.actualizar
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

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmProductosGastosDet 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1320
      Width           =   5535
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3000
      Top             =   2520
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
            Picture         =   "FrmProductoGastosDet.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoGastosDet.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoGastosDet.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoGastosDet.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoGastosDet.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoGastosDet.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoGastosDet.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoGastosDet.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoGastosDet.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoGastosDet.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoGastosDet.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoGastosDet.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   5340
      TabIndex        =   2
      Top             =   2880
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
         TabIndex        =   3
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
   Begin MSDataListLib.DataCombo DtcUnidad 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
      Text            =   "DataCombo1"
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label LblRuc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   585
      TabIndex        =   6
      Top             =   900
      Width           =   675
   End
   Begin VB.Label LblRazonSocial 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle Gastos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   1545
      TabIndex        =   5
      Top             =   180
      Width           =   1665
   End
   Begin VB.Label LblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   195
      TabIndex        =   4
      Top             =   1380
      Width           =   1065
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   2235
      Left            =   120
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "FrmProductosGastosDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrCodTabla As String
Dim strCodLinea As String

Private Sub DtcUnidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtdescripcion)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
strCadena = "SELECT cUnidad as Codigo, sAbreviatura as Descripcion FROM Unidad " & _
 " ORDER BY sAbreviatura"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcUnidad)
  
  Select Case FrmProductosGastos.Procedencia
    Case Modificar
      Call LLENA
  End Select
End Sub

Private Sub LLENA()
   StrCodTabla = FrmProductosGastos.HfgLinea.TextMatrix(FrmProductosGastos.HfgLinea.Row, 0)
   strCadena = "SELECT* FROM ProductoGastos WHERE cGastos='" & StrCodTabla & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
        Me.DtcUnidad.BoundText = rst("cUnidad")
        Me.txtdescripcion.Text = rst("Descripcion")
   End If
   Set rst = Nothing
End Sub
Private Sub Save()
  If txtdescripcion.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    Select Case FrmProductosGastos.Procedencia
      Case nuevo
         strCadena = "INSERT INTO ProductoGastos(cUnidad,Descripcion,Ruc) VALUES " & _
        " ('" & Me.DtcUnidad.BoundText & "','" & Me.txtdescripcion.Text & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        Call FrmProductosGastos.actualizar
        Unload Me
      Case Modificar
        strCadena = "UPDATE ProductoGastos SET Descripcion='" & txtdescripcion.Text & "',cUnidad='" & Me.DtcUnidad.BoundText & "' WHERE cGastos = '" & StrCodTabla & "'"
        CnBd.Execute (strCadena)
         
        Call FrmProductosGastos.actualizar
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



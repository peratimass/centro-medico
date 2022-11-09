VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleIGV 
   BorderStyle     =   0  'None
   Caption         =   "Parametros IGV"
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtDescripcion 
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   1575
      MaxLength       =   20
      TabIndex        =   0
      Top             =   240
      Width           =   2325
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   1215
      Top             =   1080
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
            Picture         =   "FrmDetalleIGV.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleIGV.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleIGV.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleIGV.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleIGV.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleIGV.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleIGV.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleIGV.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleIGV.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleIGV.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleIGV.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleIGV.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   1980
      TabIndex        =   1
      Top             =   810
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
         TabIndex        =   2
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
   Begin VB.Label LblDescripcion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   180
      TabIndex        =   3
      Top             =   285
      Width           =   1125
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2160
      Left            =   0
      Top             =   0
      Width           =   4305
   End
End
Attribute VB_Name = "FrmDetalleIGV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrCodTabla As String
Dim strCodLinea As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 2000
  Select Case FrmParametros.Procedencia
    Case Modificar
      Call LLENA
  End Select
End Sub

Private Sub LLENA()
  StrCodTabla = Format(FrmParametros.HfgLinea.TextMatrix(FrmParametros.HfgLinea.Row, 0), "yyyy-mm-dd")
  TxtDescripcion.Text = FrmParametros.HfgLinea.TextMatrix(FrmParametros.HfgLinea.Row, 1)
End Sub
Private Sub Save()
  If TxtDescripcion.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    Select Case FrmParametros.Procedencia
      Case nuevo
        strCadena = "SELECT int_linea FROM linea ORDER BY int_linea DESC"
        Call ConfiguraRst(strCadena)
        strCodLinea = GeneraCodigo(4)
        strCadena = "INSERT INTO linea (clinea,sdescripcion,int_linea) VALUES " & _
        " ('" & strCodLinea & "','" & TxtDescripcion.Text & "','" & Val(strCodLinea) & "')"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Unload Me
      Case Modificar
       strCadena = "UPDATE entidad_igv SET igv='" & Val(TxtDescripcion.Text) & "'  WHERE fecha = '" & StrCodTabla & "' and ruc='" & KEY_RUC & "'"
       CnBd.Execute (strCadena)
        
       strCadena = " SELECT * FROM entidad_igv WHERE fecha='" & CVDate(KEY_FECHA) & "' and ruc='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
            If CVDate(StrCodTabla) = CVDate(rst("fecha")) Then
                KEY_IGV = rst("igv")
                MDIFrmPrincipal.StatusBar1.Panels(2) = "IGV:" + Space(2) + str(KEY_IGV * 100) + "%"
                MDIFrmPrincipal.StatusBar1.Panels(3) = "FECHA:" + Space(2) + str(KEY_FECHA)
            End If
       End If
       Set rst = Nothing
       Call FrmParametros.actualizar(FrmParametros.HfgLinea, FrmParametros)
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



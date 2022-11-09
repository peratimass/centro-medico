VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmUnidadesTransporteDet 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtPrecio 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   50
      TabIndex        =   17
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox TxtcapacidadTanque 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtConsumoHora 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox TxtMtc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   7
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox TxtMarca 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1560
      MaxLength       =   200
      TabIndex        =   0
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox txtPlaca 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1560
      MaxLength       =   200
      TabIndex        =   2
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox TxtLlantas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1680
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   6000
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
            Picture         =   "FrmUnidadesTransporteDet.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporteDet.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporteDet.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporteDet.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporteDet.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporteDet.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporteDet.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporteDet.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporteDet.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporteDet.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporteDet.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmUnidadesTransporteDet.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   3600
      TabIndex        =   1
      Top             =   3720
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
         TabIndex        =   3
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
   Begin MSDataListLib.DataCombo DtcTipoTransporte 
      Height          =   330
      Left            =   1560
      TabIndex        =   13
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gl"
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
      Left            =   3960
      TabIndex        =   20
      Top             =   2400
      Width           =   195
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gl"
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
      Left            =   3960
      TabIndex        =   19
      Top             =   2100
      Width           =   195
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Hora :"
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
      Left            =   225
      TabIndex        =   18
      Top             =   2820
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cap Tanque:"
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
      Left            =   300
      TabIndex        =   16
      Top             =   2460
      Width           =   1125
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Consu x Hora:"
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
      Left            =   225
      TabIndex        =   15
      Top             =   2100
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mtc  :"
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
      Left            =   945
      TabIndex        =   14
      Top             =   3240
      Width           =   525
   End
   Begin VB.Label LblTelefono2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marca :"
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
      Left            =   795
      TabIndex        =   12
      Top             =   1020
      Width           =   675
   End
   Begin VB.Label LblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Llantas :"
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
      Left            =   705
      TabIndex        =   11
      Top             =   1740
      Width           =   765
   End
   Begin VB.Label LblRazonSocial 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DETALLE TRANSPORTE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1800
      TabIndex        =   10
      Top             =   300
      Width           =   1875
   End
   Begin VB.Label LblRuc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo :"
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
      Left            =   945
      TabIndex        =   9
      Top             =   660
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Placa :"
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
      Left            =   855
      TabIndex        =   8
      Top             =   1380
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "FrmUnidadesTransporteDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrCodTabla As String
Dim strCodLinea As String

Private Sub DtcTipoTransporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
CenterForm Me
 Me.Top = 1200
 strCadena = "SELECT id_tipo_transporte as Codigo, descripcion as Descripcion FROM transporte_tipo where ruc='" & KEY_RUC & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoTransporte)
  
  Select Case FrmUnidadesTransporte.Procedencia
    Case Modificar
      Call LLENA
  End Select
End Sub

Private Sub LLENA()
  strCadena = "SELECT * FROM transporte WHERE id_transporte='" & FrmUnidadesTransporte.HfgDetalle.TextMatrix(FrmUnidadesTransporte.HfgDetalle.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
    Me.DtcTipoTransporte.BoundText = rst("id_tipo_transporte")
    Me.txtmarca.Text = rst("marca")
    Me.txtplaca.Text = rst("placa")
    Me.TxtLlantas.Text = rst("llantas")
    Me.txtmtc.Text = rst("mtc")
    Me.txtConsumoHora.Text = Format(rst("consumo_hora"), "#,##0.00")
    Me.TxtcapacidadTanque.Text = Format(rst("capacidad_tanque"), "#,##0.00")
    Me.txtprecio.Text = Format(rst("precio_hora"), "#,##0.00")
  End If
  Set rst = Nothing
End Sub
Private Sub Save()
  If Me.txtmarca.Text = "" Or Me.txtplaca.Text = "" Or Me.TxtLlantas.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    Select Case FrmUnidadesTransporte.Procedencia
      Case nuevo
        StrCodTabla = formato_item(ConsultaUltimoRegistro("transporte", "id_transporte", "ruc", KEY_RUC), 5)
        strCadena = "INSERT INTO transporte (id_transporte,id_tipo_transporte,marca,placa,llantas,mtc,consumo_hora,capacidad_tanque,precio_hora,ruc) VALUES " & _
        " ('" & StrCodTabla & "','" & Me.DtcTipoTransporte.BoundText & "','" & Me.txtmarca.Text & "','" & Me.txtplaca.Text & "','" & Val(Me.TxtLlantas.Text) & "','" & Me.txtmtc.Text & "','" & Val(Me.txtConsumoHora.Text) & "','" & Val(Me.TxtcapacidadTanque.Text) & "','" & Val(Me.txtprecio.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        Call FrmUnidadesTransporte.actualizar
         Unload Me
      Case Modificar
        strCadena = "UPDATE transporte SET id_tipo_transporte='" & Me.DtcTipoTransporte.BoundText & "',marca='" & Me.txtmarca.Text & "', " & _
        " placa='" & Me.txtplaca.Text & "',consumo_hora='" & Val(Me.txtConsumoHora.Text) & "',precio_hora='" & Val(Me.txtprecio.Text) & "',llantas='" & Val(Me.TxtLlantas.Text) & "',mtc='" & Me.txtmtc.Text & "',capacidad_tanque='" & Val(Me.TxtcapacidadTanque.Text) & "' WHERE id_transporte = '" & FrmUnidadesTransporte.HfgDetalle.TextMatrix(FrmUnidadesTransporte.HfgDetalle.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        Call FrmUnidadesTransporte.actualizar
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




Private Sub TxtLlantas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtmtc)
End If
End Sub

Private Sub TxtMarca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtplaca)
End If
End Sub

Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtLlantas)
End If
End Sub

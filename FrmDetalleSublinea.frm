VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmDetalleSublinea 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFoto 
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
      Left            =   2520
      MaxLength       =   200
      TabIndex        =   7
      Top             =   1320
      Width           =   2565
   End
   Begin VB.TextBox txtbuscar 
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
      Left            =   5160
      MaxLength       =   20
      TabIndex        =   6
      Top             =   840
      Width           =   885
   End
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
      Left            =   2535
      MaxLength       =   200
      TabIndex        =   0
      Top             =   360
      Width           =   2445
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   1080
      Top             =   2280
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
            Picture         =   "FrmDetalleSublinea.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleSublinea.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleSublinea.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleSublinea.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleSublinea.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleSublinea.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleSublinea.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleSublinea.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleSublinea.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleSublinea.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleSublinea.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleSublinea.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   4260
      TabIndex        =   1
      Top             =   1890
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
      _Version        =   "6.7.9782"
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
   Begin MSDataListLib.DataCombo DtcLinea 
      Height          =   330
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IMAGEN :"
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
      Height          =   210
      Left            =   1605
      TabIndex        =   8
      Top             =   1440
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LINEA - CLASIFICACION :"
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
      Height          =   210
      Left            =   585
      TabIndex        =   4
      Top             =   960
      Width           =   1875
   End
   Begin VB.Label LblDescripcion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION MODELO:"
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
      Height          =   210
      Left            =   570
      TabIndex        =   3
      Top             =   405
      Width           =   1815
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1635
      Left            =   360
      Top             =   240
      Width           =   5790
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   2805
      Left            =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "FrmDetalleSublinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrCodTabla As String
Dim strCodLinea As String

Private Sub Form_Activate()
Select Case FrmSublineas.Procedencia
    Case modificar
      Call LLENA
  End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 1000

 strCadena = "SELECT  id_linea as Codigo,descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLinea)
  
  
End Sub

Private Sub LLENA()
  
  strCadena = "SELECT * FROM linea_sub WHERE id_tipo='" & FrmSublineas.HfgLinea.TextMatrix(FrmSublineas.HfgLinea.Row, 0) & "' AND id_usu='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
    Me.txtFoto.Text = rst("foto")
    StrCodTabla = FrmSublineas.HfgLinea.TextMatrix(FrmSublineas.HfgLinea.Row, 0)
    TxtDescripcion.Text = FrmSublineas.HfgLinea.TextMatrix(FrmSublineas.HfgLinea.Row, 2)
    
    Me.DtcLinea.BoundText = rst("id_linea")
  End If
  
End Sub
Private Sub Save()
  If TxtDescripcion.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    Select Case FrmSublineas.Procedencia
      Case nuevo
        strCadena = "SELECT * FROM linea_sub WHERE id_usu='" & KEY_RUC & "' ORDER BY id_tipo DESC"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            strCodLinea = formato_item(Val(rst("id_tipo")) + 1, 5)
        Else
            strCodLinea = formato_item(1, 5)
        End If
        
        strCadena = "INSERT INTO linea_sub (id_tipo,id_linea,descripcion,id_usu) VALUES " & _
        " ('" & strCodLinea & "','" & Me.DtcLinea.BoundText & "','" & TxtDescripcion.Text & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        strCadena = "SELECT S.id_tipo,L.descripcion as linea,S.descripcion as sublinea  FROM linea L,linea_sub S WHERE L.id_linea=S.id_linea AND L.id_usu='" & KEY_RUC & "' AND S.id_usu='" & KEY_RUC & "' and L.id_linea='" & Me.DtcLinea.BoundText & "' LIMIT 50"
        Call FrmSublineas.actualizar
        FrmSublineas.Procedencia = Neutro
        Unload Me
      Case modificar
        strCadena = "UPDATE linea_sub SET foto='" & Trim(Me.txtFoto.Text) & "',descripcion='" & TxtDescripcion.Text & "',id_linea='" & Me.DtcLinea.BoundText & "' WHERE id_tipo = '" & StrCodTabla & "' and id_linea='" & Me.DtcLinea.BoundText & "' AND id_usu='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        strCadena = "SELECT S.id_tipo,L.descripcion as linea,S.descripcion as sublinea  FROM linea L,linea_sub S WHERE L.id_linea=S.id_linea AND L.id_linea='" & Me.DtcLinea.BoundText & "' and L.id_usu=S.id_usu and L.id_usu='" & KEY_RUC & "' LIMIT 50"
        Call FrmSublineas.actualizar
        FrmSublineas.Procedencia = Neutro
        Unload Me
    End Select
    
  End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
      Call enabled_form(FrmSublineas)
      Exit Sub
    Case KEY_CANCEL
        Call enabled_form(FrmSublineas)
        Unload Me
  End Select
  Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 strCadena = "SELECT  id_linea as Codigo,descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' and descripcion like '%" & Trim(Me.txtbuscar.Text) & "%' ORDER BY descripcion "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLinea)
End If
End Sub

Private Sub txtdescripcion_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub



VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Proveedor"
   ClientHeight    =   5535
   ClientLeft      =   855
   ClientTop       =   960
   ClientWidth     =   12795
   ControlBox      =   0   'False
   Icon            =   "FrmDetalleLaboratorio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   12795
   Begin VB.TextBox TxtEntidad 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2835
      MaxLength       =   50
      TabIndex        =   21
      Top             =   720
      Width           =   6855
   End
   Begin VB.TextBox TxtDireccion1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2835
      MaxLength       =   80
      TabIndex        =   20
      Top             =   1200
      Width           =   6855
   End
   Begin VB.Frame FrmCondPago 
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   1335
      Begin VB.OptionButton OptJuridica 
         Caption         =   "Jurídica"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton OptNatural 
         Caption         =   "Natural"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtDireccion2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2835
      MaxLength       =   80
      TabIndex        =   16
      Top             =   1680
      Width           =   6855
   End
   Begin VB.TextBox TxtTelefono2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5655
      MaxLength       =   10
      TabIndex        =   15
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox TxtEmail 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   14
      Top             =   3240
      Width           =   6975
   End
   Begin VB.TextBox Txttelefono1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2835
      MaxLength       =   10
      TabIndex        =   13
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox TxtFax 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   8040
      MaxLength       =   10
      TabIndex        =   12
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox TxtObservacion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   675
      Left            =   2775
      MaxLength       =   100
      TabIndex        =   11
      Top             =   3720
      Width           =   6975
   End
   Begin VB.CommandButton CmdFoto 
      Caption         =   "Seleccione su Foto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   10
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CheckBox ChkPercepcion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Afecto a Percepción"
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   2880
      TabIndex        =   9
      Top             =   4680
      Width           =   1770
   End
   Begin VB.CheckBox ChkRetencion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Afecto a Retencion."
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   5160
      TabIndex        =   8
      Top             =   4680
      Width           =   1770
   End
   Begin VB.TextBox TxtDNI 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5655
      MaxLength       =   8
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2835
      MaxLength       =   11
      TabIndex        =   6
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CheckBox ChkCliente 
      Caption         =   "Cliente."
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CheckBox ChkProveedor 
      Caption         =   "Proveedor."
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CheckBox ChkContable 
      Caption         =   "Contable."
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CheckBox ChkTransporte 
      Caption         =   "Transporte"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CheckBox ChkPersonal 
      Caption         =   "Personal."
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   10320
      TabIndex        =   22
      Top             =   3480
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
         TabIndex        =   23
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
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   10560
      Top             =   4560
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
            Picture         =   "FrmDetalleLaboratorio.frx":0442
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLaboratorio.frx":075E
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLaboratorio.frx":0BBE
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLaboratorio.frx":101E
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLaboratorio.frx":133A
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLaboratorio.frx":179A
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLaboratorio.frx":1AB6
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLaboratorio.frx":1F16
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLaboratorio.frx":2376
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLaboratorio.frx":2C56
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLaboratorio.frx":2F72
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLaboratorio.frx":328E
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00800000&
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label LblEntidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1755
      TabIndex        =   37
      Top             =   780
      Width           =   615
   End
   Begin VB.Label LblDireccion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección 1:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1590
      TabIndex        =   36
      Top             =   1260
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Persona:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   210
      TabIndex        =   35
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1590
      TabIndex        =   34
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección 2:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1590
      TabIndex        =   33
      Top             =   1740
      Width           =   885
   End
   Begin VB.Label LblFax 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7560
      TabIndex        =   32
      Top             =   2220
      Width           =   375
   End
   Begin VB.Label LblTelefono2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono 2 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4680
      TabIndex        =   31
      Top             =   2220
      Width           =   885
   End
   Begin VB.Label LblObservacion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones : "
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   30
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label LblTelefono1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono 1 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1680
      TabIndex        =   29
      Top             =   2220
      Width           =   885
   End
   Begin VB.Label LblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E mail :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1680
      TabIndex        =   28
      Top             =   3300
      Width           =   525
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label LblTipoDocumento 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruc:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1800
      TabIndex        =   27
      Top             =   2700
      Width           =   375
   End
   Begin VB.Label LblNDocumento 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4950
      TabIndex        =   26
      Top             =   2700
      Width           =   345
   End
   Begin VB.Label LblCodPersona 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2835
      TabIndex        =   25
      Top             =   240
      Width           =   945
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accesado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   120
      TabIndex        =   24
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   2760
      Top             =   4560
      Width           =   4455
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   5115
      Left            =   1515
      Top             =   120
      Width           =   8535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   2460
      Left            =   10200
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "FrmDetalleProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrCodTabla As String
Dim TipoDocumento As String
Dim strCodProveedor As String
Dim cli As String, prov As String, Per_N As String, Per As String, Ret As String
Dim img As String



Private Sub Save()
  If Me.TxtEntidad.Text = "" Or Me.TxtDireccion1.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
      Call verificaTipo
     Select Case FrmProveedor.Procedencia
     
      Case nuevo
            strCadena = "SELECT int_persona FROM Persona ORDER BY int_persona DESC"
            Call ConfiguraRst(strCadena)
            strCodProveedor = GeneraCodigo(5)
            
         
            If Me.OptJuridica.Value = True Then
                  strCadena = "INSERT INTO Persona(cPersona,Telefono1,Telefono2,Cliente,sRazonSocial,sDireccionCliente1,sDireccionCliente2,Per_Nat " & _
                ",Per_Ruc,Per_fax,Per_Percepcion,Per_Retencion,sEmailCliente,Int_Persona,proveedor,DNI,Observacion) VALUES " & _
                "('" & strCodProveedor & "','" & Trim(Me.Txttelefono1.Text) & "','" & Trim(Me.TxtTelefono2.Text) & "'," & _
                " '" & cli & "','" & Me.TxtEntidad.Text & "','" & Me.TxtDireccion1.Text & "','" & Me.TxtDireccion2.Text & "'," & _
                " '" & Per_N & "','" & Me.TxtRuc.Text & "','" & Me.TxtFax.Text & "','" & Per & "','" & Ret & "'," & _
                "'" & Me.TxtEmail.Text & "','" & Val(strCodProveedor) & "','" & prov & "','" & Trim(Me.txtDNI.Text) & "','" & Me.txtobservacion.Text & "')"
            Else
                strCadena = "INSERT INTO Persona(cPersona,NombrePersona,Telefono1,Telefono2,Cliente,sDireccionCliente1,sDireccionCliente2,Per_Nat " & _
                ",Per_Ruc,Per_fax,Per_Percepcion,Per_Retencion,sEmailCliente,Int_Persona,proveedor,DNI,Observacion) VALUES " & _
                "('" & strCodProveedor & "','" & Me.TxtEntidad.Text & "','" & Trim(Me.Txttelefono1.Text) & "','" & Trim(Me.TxtTelefono2.Text) & "'," & _
                " '" & cli & "','" & Me.TxtDireccion1.Text & "','" & Me.TxtDireccion2.Text & "'," & _
                " '" & Per_N & "','" & Me.TxtRuc.Text & "','" & Me.TxtFax.Text & "','" & Per & "','" & Ret & "'," & _
                "'" & Me.TxtEmail.Text & "','" & Val(strCodProveedor) & "','" & prov & "','" & Trim(Me.txtDNI.Text) & "','" & Me.txtobservacion.Text & "')"
            End If
            
            Call EjecutaRST(strCadena)
            Set RstEjecuta = Nothing
            Unload Me
            
      Case Modificar
            If Me.OptJuridica.Value = True Then
                strCadena = "UPDATE Persona SET Telefono1='" & Me.Txttelefono1.Text & "'," & _
                "Telefono2='" & Me.TxtTelefono2.Text & "',Cliente='" & cli & "'," & _
                "sRazonSocial='" & Me.TxtEntidad.Text & "', sDireccionCliente1=" & _
                " '" & Me.TxtDireccion1.Text & "',sDireccionCliente2='" & Me.TxtDireccion2.Text & "' ," & _
                "Per_Nat='" & Per_N & "', Per_Ruc='" & Me.TxtRuc.Text & "',Per_fax='" & Me.TxtFax.Text & "' , " & _
                "Per_Percepcion='" & Per & "', Per_Retencion='" & Ret & "',sEmailCliente='" & Me.TxtEmail.Text & "' , " & _
                "proveedor='" & prov & "',Observacion='" & Me.txtobservacion.Text & "' WHERE cPersona= '" & StrCodTabla & "'"
            Else
                    strCadena = "UPDATE Persona SET Telefono1='" & Me.Txttelefono1.Text & "'," & _
                "Telefono2='" & Me.TxtTelefono2.Text & "',Cliente='" & cli & "'," & _
                "NombrePersona='" & Me.TxtEntidad.Text & "', sDireccionCliente1=" & _
                " '" & Me.TxtDireccion1.Text & "',sDireccionCliente2='" & Me.TxtDireccion2.Text & "'," & _
                "Per_Nat='" & Per_N & "', Per_Ruc='" & Me.TxtRuc.Text & "',Per_fax='" & Me.TxtFax.Text & "', " & _
                "Per_Percepcion='" & Per & "', Per_Retencion='" & Ret & "',sEmailCliente='" & Me.TxtEmail.Text & "'," & _
                "proveedor='" & prov & "',DNI='" & Trim(Me.txtDNI.Text) & "',Observacion='" & Me.txtobservacion.Text & "' WHERE cPersona= '" & StrCodTabla & "'"
            End If
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Unload Me
    End Select
  End If

End Sub

Private Sub CmdFoto_Click()
Me.CommonDialog1.Filter = "*.Jpg"
Me.CommonDialog1.ShowOpen
Me.Image1.Picture = LoadPicture(Me.CommonDialog1.FileName)
img = Me.CommonDialog1.FileName
End Sub

Private Sub Form_Activate()
CenterForm Me
Me.Width = 13230
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
  Select Case FrmProveedor.Procedencia
        Case nuevo
         strCadena = "SELECT Int_persona FROM Persona ORDER BY Int_persona DESC"
        Call ConfiguraRst(strCadena)
         strCodProveedor = GeneraCodigo(5)
         Me.LblCodPersona.Caption = strCodProveedor
         Me.chkProveedor.Value = 1
         Set rst = Nothing
    Case Modificar
      Call LLENA
  End Select
End Sub
 
Private Sub LLENA()
  FrmProveedor.HfdPersona.col = 0
  strCadena = "SELECT * FROM Persona WHERE cPersona = '" & FrmProveedor.HfdPersona.Text & "'"
  Call EjecutaRST(strCadena)
  StrCodTabla = RstEjecuta!cPersona
  Me.LblCodPersona.Caption = StrCodTabla
  If Trim(RstEjecuta!NombrePersona) <> "" Then
        Me.LblEntidad.Caption = "Persona:"
        Me.TxtEntidad.Text = RstEjecuta!NombrePersona
  Else
        Me.TxtEntidad.Text = RstEjecuta!sRazonSocial
  End If
  
  Me.Txttelefono1.Text = RstEjecuta!Telefono1
  Me.TxtTelefono2.Text = RstEjecuta!Telefono2
  Me.TxtDireccion1.Text = RstEjecuta!sDireccionCliente1
  Me.TxtDireccion2.Text = RstEjecuta!sDireccionCliente2
  If Trim(RstEjecuta!Per_Nat) = "N" Then
    Me.OptNatural.Value = True
   Else
    Me.OptJuridica.Value = True
  End If
  Me.TxtRuc.Text = RstEjecuta!per_ruc
  On Error GoTo sindnI
  Me.txtDNI.Text = RstEjecuta!dni
sindnI:
  On Error GoTo sinfax
  Me.TxtFax.Text = RstEjecuta!Per_fax
sinfax:
  If Trim(RstEjecuta!Per_Percepcion) = "V" Then
    Me.ChkPercepcion.Value = 1
    Per = "V"
   Else
    Me.ChkPercepcion.Value = 0
    Per = "F"
  End If
  
  If Trim(RstEjecuta!Per_Retencion) = "V" Then
    Me.ChkRetencion.Value = 1
    Ret = "V"
   Else
    Me.ChkRetencion.Value = 0
    Ret = "F"
  End If
  If Trim(RstEjecuta!Cliente) = "V" Then
    Me.ChkCliente.Value = 1
    cli = "V"
   Else
    Me.ChkCliente.Value = 0
    cli = "F"
  End If
  If Trim(RstEjecuta!Proveedor) = "V" Then
    Me.chkProveedor.Value = 1
    prov = "V"
   Else
    Me.chkProveedor.Value = 0
    prov = "F"
  End If
  Me.TxtEmail.Text = RstEjecuta!sEmailCliente
  'On Error GoTo sinObs
  'Me.TxtObservacion.Text = RstEjecuta!Observacion
'sinObs:
  Set RstEjecuta = Nothing
End Sub

Private Sub OptJuridica_Click()
    Me.LblEntidad.Caption = "Razon Social:"
    
End Sub

Private Sub OptNatural_Click()
    Me.LblEntidad.Caption = "Nombre:"
    
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

Private Sub TxtNDocumento_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("I", KeyAscii)
End Sub

Private Sub TxtTelefono_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("I", KeyAscii)
End Sub

Sub verificaTipo()
If Me.ChkCliente.Value = 1 Then
    cli = "V"
Else
    cli = "F"
End If
If Me.chkProveedor.Value = 1 Then
    prov = "V"
Else
    prov = "F"
End If
If OptJuridica.Value = True Then
    Per_N = "J"
Else
    Per_N = "N"
End If

If Me.ChkPercepcion.Value = 1 Then
    Per = "V"
Else
    Per = "F"
End If
If Me.ChkRetencion.Value = 1 Then
    Ret = "V"
Else
    Ret = "F"
End If
End Sub


Private Sub TxtDireccion1_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    Me.TxtDireccion2.SetFocus
End If
End Sub

Private Sub TxtDireccion2_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    Me.Txttelefono1.SetFocus
End If
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT NombrePersona FROM Persona WHERE Per_Ruc='" & Trim(Me.TxtRuc.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        MsgBox "DNI Usado por:" + Chr(32) + rst(0) + Chr(32), vbInformation, "Mensaje para el operador"
    End If
    Me.TxtEmail.SetFocus
End If
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtobservacion.SetFocus
End If
End Sub

Private Sub TxtEntidad_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    Me.TxtDireccion1.SetFocus
End If
End Sub

Private Sub TxtFax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtRuc.SetFocus
End If
End Sub

Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub

Private Sub txtruc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT NombrePersona,sRazonSocial FROM Persona WHERE Per_Ruc='" & Trim(Me.TxtRuc.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If rst(0) = "" Then
            MsgBox "Ruc Usado por:" + Chr(32) + rst(1) + Chr(32), vbInformation, "Mensaje para el operador"
        Else
            MsgBox "Ruc Usado por:" + Chr(32) + rst(0) + Chr(32), vbInformation, "Mensaje para el operador"
        End If
    End If
    Me.txtDNI.SetFocus
End If
End Sub

Private Sub TxtTelefono1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.TxtTelefono2.SetFocus
End If
End Sub

Private Sub TxtTelefono2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtFax.SetFocus
End If
End Sub



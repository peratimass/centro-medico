VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDetalleUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Usuario"
   ClientHeight    =   3465
   ClientLeft      =   705
   ClientTop       =   840
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5340
   Begin VB.TextBox TxtNombrecompleto 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      MaxLength       =   80
      TabIndex        =   8
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox TxtClave 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1680
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Txtnombre 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1680
      MaxLength       =   80
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   2160
      Top             =   2760
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
            Picture         =   "FrmDetalleUsuario.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUsuario.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUsuario.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUsuario.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUsuario.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUsuario.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUsuario.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUsuario.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUsuario.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUsuario.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUsuario.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleUsuario.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   2940
      TabIndex        =   2
      Top             =   2520
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
         TabIndex        =   3
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
   Begin MSDataListLib.DataCombo DtcCargo 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   1800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   675
      TabIndex        =   9
      Top             =   420
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   675
      TabIndex        =   6
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label LblDireccion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clave :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   1380
      Width           =   645
   End
   Begin VB.Label LblRazonSocial 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login :"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   690
      TabIndex        =   4
      Top             =   900
      Width           =   495
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   2235
      Left            =   240
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "FrmDetalleUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrCodTabla As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub

Private Sub Form_Load()
CenterForm Me
strCadena = "SELECT id_cargo as Codigo, descripcion as Descripcion FROM cargo " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCargo)
  Set rst = Nothing
  
  Select Case FrmUsuarios.Procedencia
    Case modificar
      Call LLENA
  End Select
End Sub

Private Sub LLENA()
  FrmUsuarios.HfdGrilla.col = 0
  strCadena = "SELECT IdUsuario,Usuario,Clave,id_cargo,Nombre FROM Seguridad WHERE IdUsuario= '" & FrmUsuarios.HfdGrilla.Text & "'"
  Call EjecutaRST(strCadena)
  StrCodTabla = RstEjecuta(0)
  txtNombre.Text = RstEjecuta(1)
  TxtClave.Text = RstEjecuta(2)
  Me.DtcCargo.BoundText = RstEjecuta(3)
  Me.TxtNombrecompleto.Text = RstEjecuta(4)
    Set RstEjecuta = Nothing
End Sub

Private Sub Save()
Dim strCod As String
Dim j As Integer
Dim i As Integer
Dim rstU As New ADODB.Recordset
Dim rstP As New ADODB.Recordset
  If txtNombre.Text = "" Or TxtClave.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    Select Case FrmUsuarios.Procedencia
      Case nuevo
        If Trim(Me.TxtNombrecompleto.Text) <> "" Or Trim(Me.txtNombre.Text) <> "" Then
        
        strCadena = "SELECT IdUsuario FROM Seguridad ORDER BY IdUsuario DESC"
        Call ConfiguraRst(strCadena)
        strCod = GeneraCodigo(4)
        strCadena = "INSERT INTO Seguridad (IdUsuario, Usuario,Nombre,Clave,id_cargo) VALUES ('" & Trim(strCod) & "'," & _
        " '" & Trim(txtNombre.Text) & "','" & Trim(Me.TxtNombrecompleto.Text) & "','" & TxtClave.Text & "','" & Trim(Me.DtcCargo.BoundText) & "')"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        
    strCadena = "SELECT  id_menu, nombre FROM   Menu ORDER BY id_menu ASC"
    Call ConfiguraRst(strCadena)
    rst.MoveFirst
    
    strCadena = "SELECT  IdUsuario FROM   Seguridad WHERE idUsuario='" & Trim(strCod) & "' "
    rstU.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
    rstU.MoveFirst
 For j = 0 To rstU.RecordCount - 1
    For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * from Usuario_permisos"
        rstP.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
        rstP.AddNew
        rstP(0) = rst(0)
        rstP(1) = rstU(0)
        rstP(2) = "no"
        rstP.Update
        Set rstP = Nothing
        rst.MoveNext
    Next i
    rstU.MoveNext
    rst.MoveFirst
Next j
        
        
        
        
        Unload Me
        Else
            MsgBox "Complete los Datos"
        End If
      Case modificar
        strCadena = "UPDATE Seguridad SET Usuario='" & txtNombre.Text & "', " & _
        " Clave='" & TxtClave.Text & "' ,id_cargo='" & Trim(Me.DtcCargo.BoundText) & "',Nombre='" & Trim(Me.TxtNombrecompleto.Text) & "' WHERE IdUsuario= '" & StrCodTabla & "'"
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

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("I", KeyAscii)
End Sub

Private Sub TxtTelefono1_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("I", KeyAscii)
End Sub

Private Sub TxtTelefono2_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("I", KeyAscii)
End Sub

Private Sub TxtFax_KeyPress(KeyAscii As Integer)
  KeyAscii = ValidaNumero("I", KeyAscii)
End Sub



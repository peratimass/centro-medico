VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmTransportista 
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   14235
   Begin VB.TextBox TxtCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   5475
      TabIndex        =   3
      Top             =   6840
      Width           =   2175
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   5475
      TabIndex        =   2
      Top             =   6120
      Width           =   2175
   End
   Begin VB.TextBox TxtRazonSocial 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   6840
      Width           =   2655
   End
   Begin VB.TextBox TxtApellido 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1395
      TabIndex        =   0
      Top             =   6120
      Width           =   2655
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   12240
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransportistas.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransportistas.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransportistas.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransportistas.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransportistas.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransportistas.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransportistas.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransportistas.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransportistas.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   3945
      Left            =   12720
      TabIndex        =   4
      Top             =   720
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   6959
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   3945
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   5
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1482
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Nuevo   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Modificar"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Salir)"
               Object.ToolTipText     =   "Buscar"
               ImageKey        =   "(Cancelar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   5655
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   9975
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   4665
      TabIndex        =   10
      Top             =   6840
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      Height          =   615
      Left            =   4515
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Width           =   3345
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00000000&
      Height          =   615
      Left            =   4515
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   3345
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruc:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   4785
      TabIndex        =   9
      Top             =   6120
      Width           =   435
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00000000&
      Height          =   615
      Left            =   315
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Width           =   3825
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Razon:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   600
      TabIndex        =   8
      Top             =   6840
      Width           =   645
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000000&
      Height          =   615
      Left            =   315
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   3825
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellidos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   420
      TabIndex        =   7
      Top             =   6120
      Width           =   915
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   1515
      Left            =   240
      Top             =   5940
      Width           =   7695
   End
End
Attribute VB_Name = "FrmTransportista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EnumFrmCliente As EnumCliente
Public Procedencia As EnumProcede

Private Sub Form_Activate()
  
  StrCadena = "SELECT cpersona as Código,NombrePersona  " & _
  "as Nombre, sRazonSocial as Razon_Social, sDireccionCliente1 as Direccion,Per_Ruc as RUC " & _
  "FROM Persona ORDER BY NombrePersona ASC"
    
  Call llenarGrid(Me.HfdPersona, Me)
End Sub
Private Sub Form_Load()
 CenterForm Me
End Sub


Private Sub OptApellido_Click()
End Sub





Private Sub HfdPersona_Click()
If HfdPersona.Row > 0 Then
  TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
  TlbAcciones.Buttons(KEY_DELETE).Enabled = True
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case KEY_NEW
      Procedencia = Nuevo
      FrmDetallePersona.Show
    Case KEY_UPDATE
      Procedencia = Modificar
      FrmDetallePersona.Show
    Case KEY_DELETE
      If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Me.HfdPersona.Col = 0
        StrCadena = "DELETE FROM Persona WHERE cPersona='" & Me.HfdPersona.Text & "'"
        Call EjecutaRST(StrCadena)
        Set RstEjecuta = Nothing
        Form_Activate
      End If
    Case "(Salir)"
      Unload Me
  End Select
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal formulario As Form)
On Error GoTo salir
   
  Call ConfiguraRst(StrCadena)
  Grilla.Clear
  Grilla.Rows = rst.RecordCount + 1
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 650
  Grilla.ColWidth(1) = 3750
  Grilla.ColWidth(2) = 3500
  Grilla.ColWidth(3) = 4100
  Grilla.ColWidth(4) = 1200

Grilla.Refresh
  formulario.TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
  formulario.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
    Me.Width = 13860
    Me.Height = 8445
    Me.Top = 1000
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub TxtApellido_Change()
StrCadena = "SELECT cpersona as Código,NombrePersona  " & _
  "as Nombre, sRazonSocial as Razon_Social, sDireccionCliente1 as Direccion,Per_Ruc as RUC " & _
  "FROM Persona WHERE NombrePersona LIKE '%" & Trim(Me.TxtApellido.Text) & "%'ORDER BY NombrePersona ASC"
Call llenarGrid(Me.HfdPersona, Me)
End Sub

Private Sub TxtApellido_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub

Private Sub TxtCodigo_Change()
StrCadena = "SELECT cpersona as Código,NombrePersona  " & _
  "as Nombre, sRazonSocial as Razon_Social, sDireccionCliente1 as Direccion,Per_Ruc as RUC " & _
  "FROM Persona WHERE int_persona = '" & Val(Me.TxtCodigo.Text) & "'ORDER BY sRazonSocial ASC"
Call llenarGrid(Me.HfdPersona, Me)
End Sub

Private Sub TxtRazonSocial_Change()
StrCadena = "SELECT cpersona as Código,NombrePersona  " & _
  "as Nombre, sRazonSocial as Razon_Social, sDireccionCliente1 as Direccion,Per_Ruc as RUC " & _
  "FROM Persona WHERE sRazonSocial LIKE '%" & Trim(Me.TxtRazonSocial.Text) & "%'ORDER BY sRazonSocial ASC"
Call llenarGrid(Me.HfdPersona, Me)
End Sub

Private Sub TxtRazonSocial_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub

Private Sub TxtRuc_Change()
StrCadena = "SELECT cpersona as Código,NombrePersona  " & _
  "as Nombre, sRazonSocial as Razon_Social, sDireccionCliente1 as Direccion,Per_Ruc as RUC " & _
  "FROM Persona WHERE Per_Ruc LIKE '%" & Trim(Me.TxtRuc.Text) & "%'ORDER BY sRazonSocial ASC"
Call llenarGrid(Me.HfdPersona, Me)
End Sub



VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmDistribuidoraProveedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distribuidora por Proveedor"
   ClientHeight    =   5970
   ClientLeft      =   450
   ClientTop       =   645
   ClientWidth     =   10950
   Icon            =   "FrmDistribuidoraLaboratorio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   10950
   Begin VB.TextBox Txtlaboratorio 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1515
      MaxLength       =   80
      TabIndex        =   0
      Top             =   2880
      Width           =   4260
   End
   Begin VB.CheckBox Chkexclusivo 
      Caption         =   "Exclusivo"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   9480
      TabIndex        =   4
      Top             =   3840
      Value           =   1  'Checked
      Width           =   1125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdLab 
      Height          =   2115
      Left            =   210
      TabIndex        =   1
      Top             =   480
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   3731
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   -2147483635
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDisLab 
      Height          =   1875
      Left            =   240
      TabIndex        =   2
      Top             =   3840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3307
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   -2147483635
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDis 
      Height          =   2115
      Left            =   5640
      TabIndex        =   3
      Top             =   480
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   3731
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   -2147483635
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   2490
      Left            =   8955
      TabIndex        =   9
      Top             =   4875
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   4392
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1800
      _CBHeight       =   2490
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   2430
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1429
         ButtonWidth     =   1402
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "(Cancelar)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Cancelar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Salir)"
               Object.ToolTipText     =   "Salir"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   7440
      Top             =   5160
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
            Picture         =   "FrmDistribuidoraLaboratorio.frx":030A
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDistribuidoraLaboratorio.frx":0626
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDistribuidoraLaboratorio.frx":0A86
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDistribuidoraLaboratorio.frx":0EE6
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDistribuidoraLaboratorio.frx":1202
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDistribuidoraLaboratorio.frx":1662
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDistribuidoraLaboratorio.frx":197E
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDistribuidoraLaboratorio.frx":1DDE
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDistribuidoraLaboratorio.frx":223E
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDistribuidoraLaboratorio.frx":2B1E
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDistribuidoraLaboratorio.frx":2E3A
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDistribuidoraLaboratorio.frx":3156
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbControl 
      Height          =   900
      Left            =   7440
      TabIndex        =   11
      Top             =   2760
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1588
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1665
      _CBHeight       =   900
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbControl 
         Height          =   810
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1429
         ButtonWidth     =   1296
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar"
               Key             =   "(Agregar)"
               Object.ToolTipText     =   "Agregar"
               ImageKey        =   "(Agregar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Quitar"
               Key             =   "(Quitar)"
               Object.ToolTipText     =   "Quitar"
               ImageKey        =   "(Quitar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Distribuidoras : "
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5640
      TabIndex        =   8
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Proveedores : "
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label lblDistLab 
      AutoSize        =   -1  'True
      Caption         =   "Distribuidora por Proveedor: "
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   2010
   End
   Begin VB.Label LblIdLaboratorio 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   405
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   500
      Left            =   240
      Top             =   2760
      Width           =   5650
   End
End
Attribute VB_Name = "FrmDistribuidoraProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Laboratorio As String
Dim Distribuidora As String
Dim Criterio As String

Private Sub Form_Load()
CenterForm Me
  strCadena = "SELECT claboratorio as Código, srazonsocial as Razón_Social,  " & _
  " stelefonolaboratorio1 as Teléfono1 FROM laboratorio ORDER BY claboratorio"
  Call ConfiguraRst(strCadena)
  Set HfdLab.Recordset = rst
  Set rst = Nothing
  HfdLab.ColWidth(1) = 3000
  Txtlaboratorio.Enabled = True
  strCadena = "SELECT cDistribuidora as Código, srazonsocial as Razón_Social,  " & _
  " stelefonodistribuidora1 as Teléfono1 FROM distribuidora ORDER BY cdistribuidora"
  Call ConfiguraRst(strCadena)
  Set HfdDis.Recordset = rst
  Set rst = Nothing
  HfdDis.ColWidth(1) = 3000
End Sub

Private Sub HfdDis_Click()
  HfdDis.col = 0
  Distribuidora = HfdDis.Text
  strCadena = "SELECT cDistribuidora as Código, srazonsocial as Razón_Social,  " & _
  " stelefonodistribuidora1 as Teléfono FROM distribuidora WHERE cdistribuidora =  '" & Distribuidora & "'"
  Call ConfiguraRst(strCadena)
  Set HfdDis.Recordset = rst
  Set rst = Nothing
End Sub

Private Sub HfdLab_Click()
  HfdLab.col = 0
  Laboratorio = HfdLab.Text
  strCadena = "SELECT claboratorio as Código, srazonsocial as Razón_Social,  " & _
  " stelefonolaboratorio1 as Teléfono FROM laboratorio WHERE claboratorio =  '" & Laboratorio & "'"
  Call ConfiguraRst(strCadena)
  Set HfdLab.Recordset = rst
  Txtlaboratorio.Text = rst(1)
  Set rst = Nothing
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
  Case KEY_CANCEL
    strCadena = "SELECT claboratorio as Código, srazonsocial as Razón_Social,  " & _
    " stelefonolaboratorio1 as Teléfono1 FROM laboratorio ORDER BY claboratorio"
    Call ConfiguraRst(strCadena)
    Set HfdLab.Recordset = rst
    Set rst = Nothing
    strCadena = "SELECT cDistribuidora as Código, srazonsocial as Razón_Social,  " & _
    " stelefonodistribuidora1 as Teléfono1 FROM distribuidora ORDER BY cdistribuidora"
    Call ConfiguraRst(strCadena)
    Set HfdDis.Recordset = rst
    Set rst = Nothing
    Txtlaboratorio.Text = ""
  Case KEY_EXIT
    Unload Me
End Select
End Sub

Private Sub TlbControl_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
  Case KEY_AGREGAR
    HfdLab.col = 0
    Laboratorio = HfdLab.Text
    HfdDis.col = 0
    Distribuidora = HfdDis.Text
    strCadena = "SELECT * FROM DistribuidoraLaboratorio WHERE claboratorio " & _
    " ='" & Laboratorio & "' AND cdistribuidora='" & Distribuidora & "' "
    Call EjecutaRST(strCadena)
    If RstEjecuta.EOF Then
        strCadena = "INSERT INTO DistribuidoraLaboratorio (claboratorio, cdistribuidora,  " & _
        " lexclusivo) VALUES('" & Laboratorio & "','" & Distribuidora & "','" & Chkexclusivo.Value & "')"
        Call EjecutaRST(strCadena)
    Else
        MsgBox "Relación Duplicada", vbCritical, "Validación"
    End If
    Set RstEjecuta = Nothing
    strCadena = "SELECT laboratorio.claboratorio, laboratorio.srazonsocial as Laboratorio,  " & _
    " distribuidora.cdistribuidora, distribuidora.srazonsocial as  Distribuidora,  " & _
    " distribuidoralaboratorio.lexclusivo as Exclusividad FROM laboratorio  " & _
    " INNER JOIN (Distribuidoralaboratorio INNER JOIN Distribuidora ON  " & _
    " distribuidora.cdistribuidora = distribuidoralaboratorio.cdistribuidora) ON  " & _
    " distribuidoralaboratorio.claboratorio = laboratorio.claboratorio  WHERE  " & _
    " laboratorio.srazonsocial LIKE '" & Criterio & "%'  ORDER BY laboratorio.claboratorio"
    Call ConfiguraRst(strCadena)
    Set HfdDisLab.Recordset = rst
    Set rst = Nothing
    strCadena = "SELECT cDistribuidora as Código, srazonsocial as Razón_Social, " & _
    "stelefonodistribuidora1 as Teléfono1 FROM distribuidora ORDER BY cdistribuidora"
    Call ConfiguraRst(strCadena)
    Set HfdDis.Recordset = rst
    Set rst = Nothing
  Case KEY_QUITAR
    HfdDisLab.col = 0
    Laboratorio = HfdDisLab.Text
    HfdDisLab.col = 2
    Distribuidora = HfdDisLab.Text
    strCadena = "DELETE FROM distribuidoralaboratorio WHERE " & _
    " claboratorio = '" & Laboratorio & "' AND cdistribuidora= '" & Distribuidora & "'"
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
    strCadena = "SELECT laboratorio.claboratorio, laboratorio.srazonsocial " & _
    " as Laboratorio, distribuidora.cdistribuidora, distribuidora.srazonsocial as  " & _
    " Distribuidora, distribuidoralaboratorio.lexclusivo as Exclusividad FROM laboratorio " & _
    " INNER JOIN (Distribuidoralaboratorio INNER JOIN Distribuidora ON " & _
    " distribuidora.cdistribuidora = distribuidoralaboratorio.cdistribuidora) ON " & _
    " distribuidoralaboratorio.claboratorio = laboratorio.claboratorio  WHERE " & _
    " laboratorio.srazonsocial LIKE '" & Criterio & "%'  ORDER BY laboratorio.claboratorio"
    Call ConfiguraRst(strCadena)
    Set HfdDisLab.Recordset = rst
    Set rst = Nothing
    strCadena = "SELECT cDistribuidora as Código, srazonsocial as Razón_Social," & _
    " stelefonodistribuidora1 as Teléfono1 FROM distribuidora ORDER BY cdistribuidora"
    Call ConfiguraRst(strCadena)
    Set HfdDis.Recordset = rst
    Set rst = Nothing
End Select
End Sub

Private Sub Txtlaboratorio_Change()
  Criterio = Trim(Txtlaboratorio.Text)
  strCadena = "SELECT claboratorio as Código, srazonsocial as Razón_Social, " & _
  " stelefonolaboratorio1 as Teléfono1 FROM laboratorio WHERE srazonsocial " & _
  " LIKE '" & Criterio & "%' ORDER BY claboratorio "
  Call ConfiguraRst(strCadena)
  Set HfdLab.Recordset = rst
  Set rst = Nothing
  strCadena = "SELECT laboratorio.claboratorio, laboratorio.srazonsocial as Laboratorio, " & _
  " distribuidora.cdistribuidora, distribuidora.srazonsocial as  Distribuidora, " & _
  " distribuidoralaboratorio.lexclusivo as Exclusividad FROM laboratorio " & _
  " INNER JOIN (Distribuidoralaboratorio INNER JOIN Distribuidora ON " & _
  " distribuidora.cdistribuidora = distribuidoralaboratorio.cdistribuidora) ON " & _
  " distribuidoralaboratorio.claboratorio = laboratorio.claboratorio  WHERE " & _
  " laboratorio.srazonsocial LIKE '" & Criterio & "%'  ORDER BY laboratorio.claboratorio"
  Call ConfiguraRst(strCadena)
  Set HfdDisLab.Recordset = rst
  Set rst = Nothing
  HfdDisLab.ColWidth(0) = 0
  HfdDisLab.ColWidth(1) = 3000
  HfdDisLab.ColWidth(2) = 0
  HfdDisLab.ColWidth(3) = 3000
End Sub

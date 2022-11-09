VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmTipoMovimiento 
   BorderStyle     =   0  'None
   Caption         =   "Tipo de Movimiento"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   -15
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3360
      Top             =   2640
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
            Picture         =   "FrmTipoMovimiento.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoMovimiento.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoMovimiento.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoMovimiento.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoMovimiento.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoMovimiento.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoMovimiento.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoMovimiento.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTipoMovimiento.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   3900
      Left            =   6000
      TabIndex        =   0
      Top             =   360
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   6879
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   3900
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   5670
         Left            =   30
         TabIndex        =   1
         Top             =   375
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   10001
         ButtonWidth     =   1588
         ButtonHeight    =   1429
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
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   5535
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9763
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Shape Shape1 
      Height          =   6495
      Left            =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "FrmTipoMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Procedencia As EnumProcede
Private Sub Form_Activate()
CenterForm Me
Me.Top = 500
  strCadena = "SELECT ctipoMovimiento as Código,sdescripcionmovimiento as " & _
  " Descripción_Movimiento FROM tipomovimiento ORDER BY ctipomovimiento"
  Call ConfiguraRst(strCadena)
  Me.HfdGrilla.Clear
  HfdGrilla.Rows = rst.RecordCount + 1
  Set HfdGrilla.Recordset = rst
  HfdGrilla.ColWidth(0) = 700
  HfdGrilla.ColWidth(1) = 4500
  
  TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
  TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  Set rst = Nothing
End Sub

Private Sub HfdGrilla_Click()
  If HfdGrilla.Row > 0 Then
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    TlbAcciones.Buttons(KEY_DELETE).Enabled = True
  End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case KEY_NEW
      Procedencia = nuevo
      FrmDetalleTipoMovimiento.Show
    Case KEY_UPDATE
      Procedencia = modificar
      FrmDetalleTipoMovimiento.Show
    Case KEY_DELETE
      If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        HfdGrilla.col = 0
        strCadena = "DELETE FROM tipomovimiento WHERE ctipomovimiento='" & HfdGrilla.Text & "'"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Form_Activate
      End If
    Case KEY_EXIT
      Unload Me
  End Select
End Sub

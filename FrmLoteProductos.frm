VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLoteProductos 
   Caption         =   "Lote de Productos"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5145
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
            Picture         =   "FrmLoteProductos.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLoteProductos.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLoteProductos.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLoteProductos.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLoteProductos.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLoteProductos.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLoteProductos.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLoteProductos.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmLoteProductos.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4260
      _Version        =   393216
      ForeColor       =   -2147483635
      FixedCols       =   0
      ForeColorFixed  =   -2147483635
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   3825
      Left            =   4140
      TabIndex        =   1
      Top             =   0
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   6747
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   3825
      _Version        =   "6.7.8862"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   2
         Top             =   345
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1508
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
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
         EndProperty
         OLEDropMode     =   1
      End
   End
End
Attribute VB_Name = "FrmLoteProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Procedencia As EnumProcede

Private Sub Form_Activate()
  StrCadena = "SELECT clote as Código,sdescripcionproducto as Producto, " & _
  " stipovencimiento as TpoV, dvencimiento as Vencimiento, npreciocompra as " & _
  " PCompra, sdescripcion as Estado FROM Producto INNER JOIN (loteProducto " & _
  " INNER JOIN Estado ON estado.cestado=loteproducto.cestado) ON " & _
  " loteproducto.cproducto=producto.cproducto ORDER BY clote"
  Call ConfiguraRst(StrCadena)
  Set HfdGrilla.Recordset = Rst
  HfdGrilla.ColWidth(0) = 900
  HfdGrilla.ColWidth(1) = 3500
  HfdGrilla.ColWidth(2) = 500
  HfdGrilla.ColWidth(3) = 1200
  HfdGrilla.ColWidth(4) = 700
  HfdGrilla.ColWidth(5) = 1300
  Call DarFormato(HfdGrilla, 4)
  TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
  TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  Set Rst = Nothing
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  Dim intAuxTop As Integer
  intAuxTop = 100
  HfdGrilla.Move 100, intAuxTop, (Me.Width - ClbAcciones.Width - 400), _
  (Me.Height - intAuxTop - 540)
  ClbAcciones.Move HfdGrilla.Width + 200, intAuxTop, Me.Width, _
  (Me.Height - intAuxTop - 540)
End Sub

Private Sub HfdGrilla_Click()
  If HfdGrilla.Row > 0 Then
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    TlbAcciones.Buttons(KEY_DELETE).Enabled = True
  End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case KEY_NEW
      Procedencia = Nuevo
      If MsgBox(MSGCONGUIA, vbQuestion + vbYesNo, MSGVERIFICACION) = 6 Then
        FrmDetalleLoteGuia.Show
      Else
        FrmDetalleLoteProducto.Show
      End If
    Case KEY_UPDATE
      Procedencia = Modificar
      FrmDetalleLoteProducto.Show
    Case KEY_DELETE
      If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        HfdGrilla.Col = 0
        StrCadena = "DELETE FROM loteproducto WHERE clote='" & HfdGrilla.Text & "'"
        Call EjecutaRST(StrCadena)
        Set RstEjecuta = Nothing
        Form_Activate
      End If
  End Select
End Sub

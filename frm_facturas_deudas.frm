VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_facturas_deudas 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtCod 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5220
      TabIndex        =   1
      Top             =   6240
      Width           =   975
   End
   Begin VB.TextBox TxtProducto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   6240
      Width           =   2895
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   9840
      Top             =   5400
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
            Picture         =   "frm_facturas_deudas.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_facturas_deudas.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_facturas_deudas.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_facturas_deudas.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_facturas_deudas.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_facturas_deudas.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_facturas_deudas.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_facturas_deudas.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_facturas_deudas.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
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
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   4665
      Left            =   9600
      TabIndex        =   3
      Top             =   480
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   8229
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   4665
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
         TabIndex        =   4
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1296
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Ver"
               Key             =   "(Imprimir)"
               Object.ToolTipText     =   "Buscar"
               ImageKey        =   "(Aceptar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas Pendientes de Pago"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   225
      TabIndex        =   7
      Top             =   120
      Width           =   3555
   End
   Begin VB.Label Label1 
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
      Left            =   4410
      TabIndex        =   6
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      TabIndex        =   5
      Top             =   6240
      Width           =   675
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   120
      Top             =   6120
      Width           =   6615
   End
End
Attribute VB_Name = "frm_facturas_deudas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call CenterForm(Me)

Me.Top = 300
strCadena = "SELECT DocumentoVenta.dVencimiento as Vencimiento, Comprobantes.doc_abrev as Tipo, DocumentoVenta.sSerie as Serie, DocumentoVenta.cDocumentoVenta as Numero, " & _
"DocumentoVenta.Persona as Cliente , DocumentoVenta.nTotalVenta as Total,DocumentoVenta.saldo as Saldo FROM DocumentoVenta INNER JOIN " & _
"Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod WHERE (DocumentoVenta.estado = 'Pendiente')"
  Call llenarGrid_prod(Me.HfdGrilla, Me)
End Sub
Sub llenarGrid_prod(ByVal Grilla As MSHFlexGrid, ByVal formulario As Form)
On Error GoTo salir
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.Rows = 100
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 950
  Grilla.ColWidth(1) = 1000
  Grilla.ColWidth(2) = 600
  Grilla.ColWidth(3) = 1100
  Grilla.ColWidth(4) = 3400
  Grilla.ColWidth(5) = 1000
  Grilla.ColWidth(6) = 1000
'Call DarFormato(Grilla, 3)
Call DarFormato(Grilla, 5)
Call DarFormato(Grilla, 6)

 
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    
    Case KEY_EXIT
        Unload Me
  End Select
End Sub

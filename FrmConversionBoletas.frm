VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmConversionBoletas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Frm Conversion Boletas"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdBuscar 
      Height          =   495
      Left            =   9600
      Picture         =   "FrmConversionBoletas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   6360
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConversionBoletas.frx":030A
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConversionBoletas.frx":075E
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConversionBoletas.frx":0A7E
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConversionBoletas.frx":0ED2
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConversionBoletas.frx":1326
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConversionBoletas.frx":1646
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConversionBoletas.frx":1966
            Key             =   "(Monto)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   5025
      Left            =   10680
      TabIndex        =   1
      Top             =   720
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   8864
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   5025
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   5460
         Left            =   30
         TabIndex        =   2
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   9631
         ButtonWidth     =   1402
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Pagar   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Monto)"
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
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgNotasCredito 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5953
      _Version        =   393216
      BackColor       =   -2147483629
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorSel    =   255
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgBoletas 
      Height          =   2085
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   3678
      _Version        =   393216
      BackColor       =   -2147483629
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      ForeColorSel    =   8388608
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   360
      Left            =   7440
      TabIndex        =   5
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   0
      CalendarForeColor=   255
      Format          =   71172099
      CurrentDate     =   39535
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   360
      Left            =   3960
      TabIndex        =   6
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmConversionBoletas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
strCadena = "SELECT Alm_cod as Codigo, Alm_des as Descripcion FROM Almacen " & _
  " ORDER BY Alm_des"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
    Beep
  End If
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = "0001"
  Me.DtcAlmacen.Enabled = True
  Me.DtpFecha.Value = CVDate(Date)
  Set rst = Nothing
    Call LLENA
End Sub

Private Sub LLENA()
strCadena = "SELECT Comprobantes.doc_cod,Comprobantes.doc_abrev as Doc, (DocumentoVenta.sSerie + '-' + DocumentoVenta.cDocumentoVenta) as Numero, DocumentoVenta.dEmisionVenta as Emision," & _
            "DocumentoVenta.Persona as Cliente , DocumentoVenta.FechaProceso as Proceso FROM DocumentoVenta INNER JOIN " & _
            "Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod " & _
            "WHERE (DocumentoVenta.doc_cod='" & KEY_NOTACRED & "' AND DocumentoVenta.Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'AND DocumentoVenta.dEmisionVenta>='" & CVDate(Me.DtpFecha.Value) & "') ORDER BY 2 ASC "
            Call llenarGrid(Me.HfgNotasCredito, Me)
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal formulario As Form)
On Error GoTo SALIR
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  If rst.RecordCount < 1 Then
    MsgBox "No hay Datos con los Parametros Ingresado", vbInformation, "Mensaje para el Usuario"
    Set rst = Nothing
    Exit Sub
  End If
  Grilla.Rows = rst.RecordCount + 1
  Set Grilla.Recordset = rst
'  Grilla.ColWidth(0) = 0
 ' Grilla.ColWidth(1) = 1100
 ' Grilla.ColWidth(2) = 2000
  'Grilla.ColWidth(3) = 1400
  'Grilla.ColWidth(4) = 4300
  'Grilla.ColWidth(5) = 2200
  'Grilla.ColWidth(6) = 1400
    
Call DarFormatoFecha(Grilla, 3)
Call DarFormatoFecha(Grilla, 5)

Set rst = Nothing

  
  Exit Sub
SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub


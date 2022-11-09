VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmReporteSalidas 
   Caption         =   "Resumen Analitico de Salidas de Productos"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   7770
   Begin TabDlg.SSTab SstKardex 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Salidas"
      TabPicture(0)   =   "FrmReporteVentas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblCantidad"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DtpHasta"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DtpDesde"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DtcAlmacen"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DtcProducto"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkProducto"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkAlmacen"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.CheckBox chkAlmacen 
         BackColor       =   &H80000004&
         Caption         =   "Por Almacen:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   225
         TabIndex        =   2
         Top             =   2370
         Width           =   1335
      End
      Begin VB.CheckBox chkProducto 
         BackColor       =   &H80000004&
         Caption         =   "Por Producto:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   225
         TabIndex        =   1
         Top             =   1740
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DtcProducto 
         Height          =   315
         Left            =   1905
         TabIndex        =   3
         Top             =   1710
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   315
         Left            =   1950
         TabIndex        =   4
         Top             =   2340
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   915
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   62128129
         CurrentDate     =   37091
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   315
         Left            =   3360
         TabIndex        =   6
         Top             =   915
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   62128129
         CurrentDate     =   37091
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rango de Fechas:"
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
         Left            =   705
         TabIndex        =   8
         Top             =   400
         Width           =   1695
      End
      Begin VB.Label LblCantidad 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Al:"
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
         Left            =   2775
         TabIndex        =   7
         Top             =   960
         Width           =   225
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         FillColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   600
         Top             =   720
         Width           =   4815
      End
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   4065
      Left            =   6480
      TabIndex        =   9
      Top             =   240
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   7170
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   4065
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1667
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Aceptar   "
               Key             =   "(Aceptar)"
               Object.ToolTipText     =   "Aceptar"
               ImageKey        =   "(Aceptar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Imprimir"
               Key             =   "(Imprimir)"
               ImageKey        =   "(Imprimir)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exportar"
               Key             =   "(Exportar)"
               ImageKey        =   "(Excel)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Left            =   6720
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":001C
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":0470
            Key             =   "(Excel)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":084A
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":0B6A
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":0FBE
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":1412
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":1732
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":1B86
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":1CE2
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":2136
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":2452
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":2D2E
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":304E
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteVentas.frx":336E
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CmdlExcel 
      Left            =   1320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgExcel 
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmReporteSalidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Producto As String
Dim Linea As String
Dim Proveedor As String
Dim marca As String

Private Sub chkAlmacen_Click()
If Me.chkAlmacen.Value = 1 Then
    Me.DtcAlmacen.Enabled = True
    
Else
    Me.DtcAlmacen.Enabled = False
    Almacen = ""
End If
End Sub

Private Sub ChkProducto_Click()
If Me.chkProducto.Value = 1 Then
    Me.DtcProducto.Enabled = True
    
Else
    Me.DtcProducto.Enabled = False
    Producto = ""
End If
End Sub

Private Sub Form_Load()
Call FormReport(Me)
Me.DtpDesde.Value = CVDate(Date)
Me.DtpHasta.Value = CVDate(Date)
'---------Llenar----- Producto
  strCadena = "SELECT cProducto as Codigo, DescripcionProducto as Descripcion FROM " & _
  " Producto ORDER BY DescripcionProducto"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcProducto)
  Set rst = Nothing
 '----------------------------
  strCadena = "SELECT Alm_cod as Codigo, Alm_des as Descripcion FROM " & _
  " Almacen ORDER BY Alm_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Set rst = Nothing
  
  Me.DtcAlmacen.Enabled = False
  Me.DtcProducto.Enabled = False

End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim DteInicio As Date, DteFin As Date
Dim StrProducto As String
Dim Ans As Boolean
Dim Anulado As String
Select Case Button.Key
  Case KEY_OK
    DteInicio = CDate(Me.DtpDesde.Value)
    DteFin = CDate(Me.DtpHasta.Value)
    Producto = ""
    Linea = ""
    Producto = ""
    Anulado = "F"
    If chkProducto.Value = 1 Then
      Producto = Replace(DtcProducto.BoundText, "'", "''")
    End If
    If Me.chkAlmacen.Value = 1 Then
      Almacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
    End If
        
    
    strCadena = "SELECT     Detalle_DocumentoVenta.cProducto, DocumentoVenta.dEmisionVenta,Comprobantes.doc_abrev ,(Detalle_DocumentoVenta.sSerie + '-'+" & _
                    "Detalle_DocumentoVenta.cDocumentoVenta) as Numero, Producto.DescripcionProducto, Detalle_DocumentoVenta.Precio," & _
                    "Detalle_DocumentoVenta.Cantidad , Detalle_DocumentoVenta.TOTAL, DocumentoVenta.Persona " & _
                    "FROM         DocumentoVenta INNER JOIN " & _
                    "Detalle_DocumentoVenta ON DocumentoVenta.cDocumentoVenta = Detalle_DocumentoVenta.cDocumentoVenta AND " & _
                    "DocumentoVenta.doc_cod = Detalle_DocumentoVenta.doc_cod AND " & _
                    "DocumentoVenta.Alm_cod = Detalle_DocumentoVenta.Alm_Cod AND " & _
                    "DocumentoVenta.sSerie = Detalle_DocumentoVenta.sSerie INNER JOIN " & _
                    "Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
                    "Producto ON Detalle_DocumentoVenta.cProducto = Producto.cProducto " & _
                    "WHERE (DocumentoVenta.Alm_cod LIKE '" & Almacen & "%'AND Detalle_DocumentoVenta.cProducto LIKE '" & Producto & "%' AND " & _
                    "DocumentoVenta.dEmisionVenta>='" & DteInicio & "' AND DocumentoVenta.dEmisionVenta<='" & DteFin & "' AND DocumentoVenta.Anulado='" & Trim(Anulado) & "')ORDER BY 2 ASC"
    
    Call ConfiguraRst(strCadena)
    'Set DataReport2.DataSource = Rst
    '    DataReport2.Show
     Ans = ShowMultiReport(rst, "RptSalida", , App.Path + "\Reportes\")
     Set Me.HfgExcel.Recordset = rst
   Case KEY_EXCEL
     '   Call ReporteExcel(HfgExcel, Me.CmdlExcel)
        Me.TlbAcciones.Buttons(KEY_EXCEL).Enabled = False
    Case KEY_CANCEL
        Unload Me
End Select
End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmReporteKardex 
   Caption         =   "Reporte de Kardex"
   ClientHeight    =   10530
   ClientLeft      =   195
   ClientTop       =   450
   ClientWidth     =   15255
   Icon            =   "FrmReporteKardex.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10530
   ScaleWidth      =   15255
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   8520
      TabIndex        =   19
      Top             =   2160
      Width           =   2655
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6495
      Left            =   240
      TabIndex        =   18
      Top             =   3480
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   11456
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab SstKardex 
      Height          =   2775
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4895
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Producto"
      TabPicture(0)   =   "FrmReporteKardex.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DtcAlmacen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DtcProveedor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DtcLinea"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DtcProducto"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkProducto"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ChkClinea"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkProveedor"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ChkAlmacen"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Fechas"
      TabPicture(1)   =   "FrmReporteKardex.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "lblDesde"
      Tab(1).Control(2)=   "DtpHasta"
      Tab(1).Control(3)=   "DtpDesde"
      Tab(1).Control(4)=   "ChkFecha"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Tipo Movimiento"
      TabPicture(2)   =   "FrmReporteKardex.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DtcTipoMovimiento"
      Tab(2).Control(1)=   "ChkTipoMovimiento"
      Tab(2).ControlCount=   2
      Begin VB.CheckBox ChkAlmacen 
         BackColor       =   &H80000004&
         Caption         =   "Almacen"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   2070
         Width           =   1095
      End
      Begin VB.CheckBox ChkTipoMovimiento 
         BackColor       =   &H80000004&
         Caption         =   "Por Tipo Movimiento:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74775
         TabIndex        =   14
         Top             =   795
         Width           =   1830
      End
      Begin VB.CheckBox chkProveedor 
         BackColor       =   &H80000004&
         Caption         =   "Proveedor"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1170
         Width           =   1335
      End
      Begin VB.CheckBox ChkClinea 
         BackColor       =   &H80000004&
         Caption         =   "Línea"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1620
         Width           =   1095
      End
      Begin VB.CheckBox ChkFecha 
         BackColor       =   &H80000004&
         Caption         =   "Por Rango de Fechas"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74760
         TabIndex        =   7
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox chkProducto 
         BackColor       =   &H80000004&
         Caption         =   "Por Producto"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   660
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DtcProducto 
         Height          =   315
         Left            =   1965
         TabIndex        =   0
         Top             =   630
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   -73920
         TabIndex        =   1
         Top             =   1275
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   153878529
         CurrentDate     =   37091
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   315
         Left            =   -71520
         TabIndex        =   2
         Top             =   1275
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   153878529
         CurrentDate     =   37091
      End
      Begin MSDataListLib.DataCombo DtcLinea 
         Height          =   315
         Left            =   1965
         TabIndex        =   11
         Top             =   1590
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcProveedor 
         Height          =   315
         Left            =   1965
         TabIndex        =   13
         Top             =   1140
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcTipoMovimiento 
         Height          =   315
         Left            =   -72795
         TabIndex        =   15
         Top             =   765
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   315
         Left            =   1965
         TabIndex        =   17
         Top             =   2040
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin VB.Label lblDesde 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Desde : "
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -74520
         TabIndex        =   9
         Top             =   1335
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Hasta : "
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -72120
         TabIndex        =   8
         Top             =   1335
         Width           =   555
      End
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   1785
      Left            =   6240
      TabIndex        =   3
      Top             =   600
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   3149
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   1785
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   4
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
            NumButtons      =   3
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
      Left            =   6120
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteKardex.frx":035E
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteKardex.frx":07B2
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteKardex.frx":0AD2
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteKardex.frx":0F26
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteKardex.frx":137A
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteKardex.frx":169A
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteKardex.frx":1AEE
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteKardex.frx":1C4A
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteKardex.frx":209E
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteKardex.frx":23BA
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteKardex.frx":2C96
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteKardex.frx":2FB6
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReporteKardex.frx":32D6
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmReporteKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim Linea As String, strProveedor As String
Dim TipoProducto As String, TipoMovimiento As String
Dim Proveedor As String
Dim Almacen As String

Private Sub chkAlmacen_Click()
If Me.ChkAlmacen.Value = 1 Then
    Me.DtcAlmacen.Enabled = True
Else
    Me.DtcAlmacen.Enabled = False
End If
End Sub

Private Sub ChkClinea_Click()
  If ChkClinea.Value = 1 Then
    DtcLinea.Enabled = True
  Else
    DtcLinea.Enabled = False
   End If
End Sub



Private Sub ChkFecha_Click()
  If ChkFecha.Value = 1 Then
    DtpDesde.Enabled = True
    DtpHasta.Enabled = True
  Else
    DtpDesde.Enabled = False
    DtpHasta.Enabled = False
  End If
End Sub



Private Sub ChkProducto_Click()
  If chkProducto.Value = 1 Then
    DtcProducto.Enabled = True
  Else
    DtcProducto.Enabled = False
  End If
End Sub

Private Sub chkProveedor_Click()
If chkProveedor.Value = 1 Then
    Me.DtcProveedor.Enabled = True
  Else
    Me.DtcProveedor.Enabled = False
  End If
End Sub

Private Sub ChkTipoMovimiento_Click()
  If ChkTipoMovimiento.Value = 1 Then
    DtcTipomovimiento.Enabled = True
  Else
    DtcTipomovimiento.Enabled = False
  End If
End Sub

Private Sub Command1_Click()
strCadena = "SELECT     TOP 100 PERCENT Producto_barras.cod_barra, Producto.DescripcionProducto, Unidad.sAbreviatura, " & _
  "                      Almacen_Productos.Stock -                          (SELECT     SUM(Stk_cant) " & _
  "                           From Kardex " & _
  "                            WHERE      Producto.cProducto = Kardex.cProducto AND FechaProceso >= '03-05-2011') AS Inicial, " & _
  "                          (SELECT     SUM(Ing_Cant)                            From Kardex " & _
  "                            WHERE      Producto.cProducto = Kardex.cProducto AND FechaProceso >= '03-05-2011') AS Ingresos, " & _
  "                          (SELECT     SUM(Sal_Cant)                             From Kardex " & _
  "                            WHERE      Producto.cProducto = Kardex.cProducto AND FechaProceso >= '03-05-2011') AS Salidas, Almacen_Productos.Stock " & _
  "FROM         Almacen_Productos INNER JOIN " & _
  "                      Producto ON Almacen_Productos.cProducto = Producto.cProducto INNER JOIN " & _
  "                      Unidad ON Producto.cUnidad = Unidad.cUnidad INNER JOIN " & _
  "                      Producto_barras ON Producto.cProducto = Producto_barras.cProducto " & _
  "ORDER BY Producto.cod_barra"
  Call ConfiguraRst(strCadena)
  Ans = ShowMultiReport(rst, "RptValorizado_mes", , App.Path + "\Reportes\")
End Sub

Private Sub Form_Load()
Call FormReport(Me)
  strCadena = "SELECT cProducto as Codigo, DescripcionProducto as Descripcion " & _
  " FROM Producto ORDER BY DescripcionProducto"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(DtcProducto)
  
  
  
 
  
  
  strCadena = "SELECT cLinea as Codigo, sDescripcion as Descripcion FROM " & _
  " Linea ORDER BY sDescripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(DtcLinea)
  
  Proveedor = "V"
  strCadena = "SELECT cPersona as Codigo, NombrePersona as Descripcion FROM " & _
  " Persona WHERE (proveedor='" & Proveedor & "' AND NombrePersona IS NOT NULL) ORDER BY NombrePersona ASC "
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
  Call LlenaDataCombo(Me.DtcProveedor)
  End If
  strCadena = "SELECT cTipoMovimiento as Codigo, sDescripcionMovimiento as Descripcion FROM " & _
  " TipoMovimiento ORDER BY sDescripcionMovimiento"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(DtcTipomovimiento)
  
  strCadena = "SELECT Alm_cod as Codigo, Alm_Des as Descripcion FROM " & _
  " Almacen ORDER BY Alm_Des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  
  DtpDesde.Value = Date
  DtpHasta.Value = Date
  DtpDesde.Enabled = False
  DtpHasta.Enabled = False
  DtcProducto.Enabled = False
  DtcLinea.Enabled = False
  Me.DtcProveedor.Enabled = False
  DtcTipomovimiento.Enabled = False
  
  
  Set Me.MSHFlexGrid1.Recordset = rst
  
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim DteInicio As Date, DteFin As Date
Dim StrProducto As String
Dim Ans As Boolean
Select Case Button.key
  Case KEY_OK
    DteInicio = CDate(DTEMINIMA)
    DteFin = CDate(DTEMAXIMA)
    Linea = ""
    strProveedor = ""
    TipoMovimiento = ""
    StrProducto = ""
    If ChkFecha.Value = 1 Then
      DteInicio = CVDate(DtpDesde.Value)
      DteFin = CVDate(DtpHasta.Value)
    End If
    If chkProducto.Value = 1 Then
      StrProducto = Replace(DtcProducto.BoundText, "'", "''")
    End If
    If ChkClinea.Value = 1 Then
      Linea = Replace(DtcLinea.BoundText, "'", "''")
    End If
    If Me.chkProveedor.Value = 1 Then
      strProveedor = Replace(Me.DtcProveedor.Text, "'", "''")
    End If
    If ChkTipoMovimiento.Value = 1 Then
      TipoMovimiento = Replace(DtcTipomovimiento.BoundText, "'", "''")
    End If
    If Me.ChkAlmacen.Value = 1 Then
        Almacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
    End If
    
    strCadena = "SELECT Kardex.cKardex, Kardex.cTipoMovimiento, TipoMovimiento.sDescripcionMovimiento, Kardex.doc_cod, Comprobantes.doc_abrev, " & _
    "(Kardex.sSerie +'-'+ Kardex.NumeroDoc), Kardex.cProducto, Producto.DescripcionProducto, Kardex.Mov_Cant, Kardex.Precio," & _
    "Kardex.Stk_Soles , Kardex.FechaEmision, Linea.cLinea, Linea.sDescripcion, Kardex.Alm_cod, almacen.Alm_des " & _
    "FROM Kardex INNER JOIN TipoMovimiento ON Kardex.cTipoMovimiento = TipoMovimiento.cTipoMovimiento INNER JOIN " & _
    "Comprobantes ON Kardex.doc_cod = Comprobantes.doc_cod INNER JOIN Producto ON Kardex.cProducto = Producto.cProducto INNER JOIN " & _
    "Linea ON Producto.cLinea = Linea.cLinea INNER JOIN Almacen ON Kardex.Alm_cod = Almacen.Alm_cod " & _
    "WHERE Kardex.cTipoMovimiento LIKE '" & TipoMovimiento & "%' " & _
    " AND kardex.cproducto LIKE '" & StrProducto & "%' AND Linea.cLinea LIKE '" & Linea & "%' AND Kardex.Persona LIKE " & _
    "'" & strProveedor & "%' AND Kardex.FechaEmision >= '" & CVDate(DteInicio) & "' AND Kardex.FechaEmision " & _
    " <= '" & CVDate(DteFin) & "' AND Kardex.Alm_cod LIKE '" & Almacen & "%' ORDER BY Kardex.FechaEmision, Producto.DescripcionProducto"
    Call ConfiguraRst(strCadena)
    'Set DataReport2.DataSource = Rst
     '   DataReport2.Show
    Ans = ShowMultiReport(rst, "RptKardex", , App.Path + "\Reportes\")
  Case KEY_CANCEL
    Unload Me
End Select
End Sub

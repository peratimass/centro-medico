VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmReporteValorizado 
   Caption         =   "Movimiento Valorizado de Productos"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   7950
   Begin VB.OptionButton OptDescripcion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Descripcion."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   3480
      Width           =   1455
   End
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
      TabCaption(0)   =   "Mov. Valorizado"
      TabPicture(0)   =   "FrmValorizado.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblCantidad"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Shape2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DtpHasta"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DtpDesde"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DtcAlmacen"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "DtcProducto"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkProducto"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkAlmacen"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "OptCodigo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.OptionButton OptCodigo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Codigo."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CheckBox chkAlmacen 
         Caption         =   "Almacen"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   225
         TabIndex        =   2
         Top             =   2370
         Width           =   1335
      End
      Begin VB.CheckBox chkProducto 
         Caption         =   "Por Producto"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   225
         TabIndex        =   1
         Top             =   1740
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DtcProducto 
         Height          =   315
         Left            =   1665
         TabIndex        =   3
         Top             =   1710
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   -2147483635
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   315
         Left            =   1710
         TabIndex        =   4
         Top             =   2340
         Width           =   3735
         _ExtentX        =   6588
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
         Format          =   176685057
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
         Format          =   176685057
         CurrentDate     =   37091
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenado Por:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2040
         TabIndex        =   10
         Top             =   2880
         Width           =   1320
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         FillColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   600
         Top             =   2880
         Width           =   4815
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
      TabIndex        =   12
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
         Height          =   5670
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   10001
         ButtonWidth     =   1746
         ButtonHeight    =   1429
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
            Picture         =   "FrmValorizado.frx":001C
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmValorizado.frx":0470
            Key             =   "(Excel)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmValorizado.frx":084A
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmValorizado.frx":0B6A
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmValorizado.frx":0FBE
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmValorizado.frx":1412
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmValorizado.frx":1732
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmValorizado.frx":1B86
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmValorizado.frx":1CE2
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmValorizado.frx":2136
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmValorizado.frx":2452
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmValorizado.frx":2D2E
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmValorizado.frx":304E
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmValorizado.frx":336E
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CmdlExcel 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmReporteValorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Producto As String
Dim Almacen As String

Private Sub ChkProducto_Click()
If Me.chkProducto.Value = 1 Then
    Me.DtcProducto.Enabled = True
Else
    Me.DtcProducto.Enabled = False
End If
End Sub

Private Sub Form_Load()

Call FormReport(Me)
CenterForm Me
Me.Top = 200
strCadena = "SELECT cProducto as Codigo,DescripcionProducto as Descripcion FROM Producto ORDER BY DescripcionProducto"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcProducto)
Me.DtcProducto.Enabled = False

strCadena = "SELECT Alm_cod as Codigo,Alm_des as Descripcion FROM Almacen ORDER BY Alm_des"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.Enabled = True
Me.OptCodigo.Value = True
Me.DtpDesde.Value = Date
Me.DtpHasta.Value = Date
Me.ChkAlmacen.Value = 1
Me.TlbAcciones.Buttons(KEY_EXCEL).Enabled = False
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Ans As Boolean
        
Select Case Button.key
    Case KEY_OK
        
        DteInicio = CVDate(Me.DtpDesde.Value)
        DteFin = CVDate(Me.DtpHasta.Value)
        
    If chkProducto.Value = 1 Then
      Producto = Replace(DtcProducto.BoundText, "'", "''")
    End If
    
    If Me.ChkAlmacen.Value = 1 Then
      Almacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
    End If
       
     If CVDate(Me.DtpDesde.Value) < CVDate(Date) Then
     
     strCadena = "SELECT     TOP 100 PERCENT Producto_barras.cod_barra, Producto.DescripcionProducto, Unidad.sAbreviatura, " & _
     " Almacen_Productos.Stock -(SELECT     SUM(Stk_cant) From Kardex WHERE      Producto.cProducto = Kardex.cProducto AND FechaProceso >= '" & CVDate(DteInicio) & "' AND FechaProceso <= '" & CVDate(DteFin) & "') AS Inicial, " & _
     "(SELECT     SUM(Ing_Cant) From Kardex WHERE      Producto.cProducto = Kardex.cProducto AND FechaProceso >= '" & CVDate(DteInicio) & "' AND FechaProceso <= '" & DteFin & "') AS Ingresos, " & _
     "(SELECT     SUM(Sal_Cant) From Kardex WHERE      Producto.cProducto = Kardex.cProducto AND FechaProceso >= '" & CVDate(DteInicio) & "' AND FechaProceso <= '" & CVDate(DteFin) & "') AS Salidas, Almacen_Productos.Stock, " & _
     "Linea.sDescripcion , Producto.PrecioCompra FROM         Almacen_Productos INNER JOIN " & _
     "Producto ON Almacen_Productos.cProducto = Producto.cProducto INNER JOIN " & _
     "Unidad ON Producto.cUnidad = Unidad.cUnidad INNER JOIN " & _
     "Producto_barras ON Producto.cProducto = Producto_barras.cProducto INNER JOIN " & _
     "Linea ON Producto.cLinea = Linea.cLinea WHERE     (Producto_barras.cod_barra <> '123456789') ORDER BY Linea.sDescripcion"
      
      Call ConfiguraRst(strCadena)
      Ans = ShowMultiReport(rst, "RptValorizado_mes", , App.Path + "\Reportes\")
      Set rst = Nothing
     Else
     strCadena = "SELECT     Almacen_Productos.cProducto, Producto_barras.cod_barra, Producto.DescripcionProducto, Unidad.sAbreviatura," & _
     "Linea.cLinea, Linea.sDescripcion , Almacen_Productos.Stock, Producto.PrecioCompra FROM Almacen_Productos INNER JOIN " & _
     "Producto ON Almacen_Productos.cProducto = Producto.cProducto INNER JOIN Unidad ON Producto.cUnidad = Unidad.cUnidad INNER JOIN " & _
     "Producto_barras ON Producto.cProducto = Producto_barras.cProducto INNER JOIN Linea ON Producto.cLinea = Linea.cLinea"
     Call ConfiguraRst(strCadena)
     Ans = ShowMultiReport(rst, "RptValorizado01", , App.Path + "\Reportes\")
      Set rst = Nothing
     End If
    
    
    
    
    Case KEY_PRINT
       DteInicio = Me.DtpDesde.Value
       DteFin = Me.DtpHasta.Value
If Me.chkProducto.Value = 0 Then
        
        
        
        
Else
        
End If
        
       Call ConfiguraRst(strCadena)
     
      Set rst = Nothing
    Case KEY_EXCEL
        
    Case KEY_CANCEL
        Unload Me
End Select
End Sub



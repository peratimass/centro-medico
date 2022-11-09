VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmReporteConsolidado 
   BorderStyle     =   0  'None
   Caption         =   "Consolidado de Mov Fisico de Alamacen"
   ClientHeight    =   4455
   ClientLeft      =   -60
   ClientTop       =   -15
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SstKardex 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Mov. Valorizado"
      TabPicture(0)   =   "FrmConsolidado.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblCantidad"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Shape4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DtpHasta"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DtpDesde"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "DtcAlmacen"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkAlmacen"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ChkValorizado"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ChkStock"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ORDENADO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   11
         Top             =   1560
         Width           =   5775
         Begin VB.OptionButton OptDescripcion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "DESCRIPCION"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3000
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton OptCodigo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "CODIGO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   720
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CheckBox ChkStock 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "STOCK"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CheckBox ChkValorizado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "VALORIZADO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CheckBox chkAlmacen 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ALMACEN"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   825
         TabIndex        =   1
         Top             =   2970
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   315
         Left            =   2550
         TabIndex        =   2
         Top             =   2940
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtpDesde 
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   915
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   166133761
         CurrentDate     =   37091
      End
      Begin MSComCtl2.DTPicker DtpHasta 
         Height          =   315
         Left            =   3960
         TabIndex        =   4
         Top             =   915
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   -2147483635
         CalendarTitleForeColor=   -2147483635
         Format          =   166133761
         CurrentDate     =   37091
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         FillColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   600
         Top             =   2880
         Width           =   5775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RANGO DE FECHAS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   810
         TabIndex        =   6
         Top             =   405
         Width           =   1485
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
         TabIndex        =   5
         Top             =   960
         Width           =   225
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   600
         Top             =   720
         Width           =   5775
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   600
         Top             =   2280
         Width           =   5775
      End
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   3705
      Left            =   8160
      TabIndex        =   7
      Top             =   120
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   6535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   3705
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1429
         ButtonWidth     =   1720
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Fisico"
               Key             =   "(Aceptar)"
               Object.ToolTipText     =   "Fisico"
               ImageKey        =   "(Aceptar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Contable"
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
      Left            =   10320
      Top             =   3360
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
            Picture         =   "FrmConsolidado.frx":001C
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsolidado.frx":0470
            Key             =   "(Excel)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsolidado.frx":084A
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsolidado.frx":0B6A
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsolidado.frx":0FBE
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsolidado.frx":1412
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsolidado.frx":1732
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsolidado.frx":1B86
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsolidado.frx":1CE2
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsolidado.frx":2136
            Key             =   "(Reporte)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsolidado.frx":2452
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsolidado.frx":2D2E
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsolidado.frx":304E
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsolidado.frx":336E
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CmdlExcel 
      Left            =   1080
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   4455
      Left            =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "FrmReporteConsolidado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strcodigo As String
Dim strDescripcion As String
Dim OrderCod As Boolean
Private Sub chkAlmacen_Click()
If Me.chkAlmacen.Value = 1 Then
    Me.DtcAlmacen.Enabled = True
Else
    Me.DtcAlmacen.Enabled = False
End If
End Sub

Private Sub Form_Load()
Call FormReport(Me)
Me.Height = 4440
Me.Width = 9300
Me.Top = 800
strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen where ruc='" & KEY_RUC & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM
Me.chkAlmacen.Value = 1
Me.OptCodigo.Value = True
Me.DtpDesde.Value = Date
Me.DtpHasta.Value = Date
Me.TlbAcciones.Buttons(KEY_EXCEL).Enabled = False
End Sub

Private Sub OptCodigo_Click()
If Me.OptCodigo.Value = 1 Then
    OrderCod = True
End If

End Sub

Private Sub OptDescripcion_Click()
If Me.OptDescripcion.Value = 1 Then
    OrderCod = False
End If
End Sub
Private Sub Impresion(ByVal FechaIni As Date, ByVal FechaFin As Date)
Dim rst_ventas As New ADODB.Recordset
Dim rst_compras As New ADODB.Recordset
Printer.ScaleMode = vbCharacters 'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = 14
    Anulado = "F"
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    
    'StrCadena = "SELECT Almacen.Alm_des, Almacen_Productos.cProducto, Producto.DescripcionProducto, Unidad.sAbreviatura, " & _
    "Almacen_Productos.Stock , Producto.PrecioVenta FROM Almacen_Productos INNER JOIN " & _
    "Producto ON Almacen_Productos.cProducto = Producto.cProducto INNER JOIN " & _
    "Unidad ON Producto.cUnidad = Unidad.cUnidad INNER JOIN " & _
    "Almacen ON Almacen_Productos.Alm_cod = Almacen.Alm_cod " & _
            "WHERE (Almacen_Productos.Alm_cod = '" & Me.DtcAlmacen.BoundText & "') ORDER BY Producto.DescripcionProducto"
    strCadena = "SELECT Almacen_Productos.cProducto, Producto.DescripcionProducto, Unidad.sAbreviatura, Almacen_Productos.Stock, " & _
    "Producto.PrecioVenta FROM Almacen_Productos INNER JOIN Producto ON Almacen_Productos.cProducto = Producto.cProducto INNER JOIN " & _
    "Unidad ON Producto.cUnidad = Unidad.cUnidad"
    Call ConfiguraRst(strCadena)
    rst.MoveFirst
    
    
     
    Printer.CurrentY = Printer.CurrentY + 0.3
    Printer.Print Tab(4); "CONSOLIDADO DE ALMACEN      :" + "LISTADO GENERAL" + Space(10); "LIRIO DE LOS VALLES SAC"
    Printer.Print Tab(4); "DEL :" + Space(1); str(FechaIni) + Space(1), "Al" + Space(1) + str(FechaFin)
    Printer.Print ""
    Printer.Print Tab(4); "ALMACEN:" + Space(1) + rst(0) + Space(10) + "USUARIO:" + Space(2) + KEY_VENDEDOR;
    Printer.Print Tab(4); "--------------------------------------------------------------------------------------"
    Printer.Print Tab(2); "CODIGO:" + Space(15) + "DESCRIPCION" + Space(27) + "UND" + Space(10) + "STOCK" + Space(10) + "PRECIO" + Space(2) + "TOTAL";
             
    rst.MoveFirst
    
    Printer.CurrentY = Printer.CurrentY + 0.2
    Dim precio As String
    Dim valor_parcial As Double
    Dim valor_total As Double
    valor_total = 0
            For j = 0 To rst.RecordCount - 1
                strCadena = "SELECT Ing_Cant,Sal_Cant FROM Kardex   WHERE cProducto='" & Trim(rst(0)) & "' AND FechaEmision<='" & FechaFin & "'"
                rst_ventas.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
                If (rst_ventas.RecordCount > 0) Then
                    Set rst_ventas = Nothing
                strCadena = "SELECT cProducto, FechaEmision, Ing_Cant, Sal_Cant, Stk_Gen FROM Kardex   WHERE cProducto='" & Trim(rst(0)) & "' AND FechaEmision<='" & FechaFin & "' ORDER BY 1 ASC"
                rst_ventas.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
                rst_ventas.MoveLast
                
                 stock = rst_ventas(4)
                Else
                    stock = 0
                End If
                
                precio = Format(rst(4), "#,##0.00")
                valor_parcial = Format(stock * rst(4), "#,##0.00")
                valor_total = valor_total + valor_parcial
                
                
                Printer.Print Tab(4); rst(0) & Space(5) & Mid(rst(1) + Space(50), 1, 50) & Space(2) & Mid(rst(2) + Space(10), 1, 10) + Space(2) + Mid(str(stock) + Space(10), 1, 10) + Space(2) + precio + Space(3) + str(valor_parcial);
                Printer.CurrentY = Printer.CurrentY + 0.4
                Set rst_ventas = Nothing
               
                rst.MoveNext
            Next j
            
    Set rst = Nothing
    Printer.Print Tab(10); "MONTO ACUMULADO TOTAL" + Space(3); Format(valor_total, "#,##0.00")
       
    Printer.EndDoc
End Sub
Private Sub Impresion_contable(ByVal FechaIni As Date, ByVal FechaFin As Date)
Dim rst_ini As New ADODB.Recordset
Dim rst_sal As New ADODB.Recordset
Dim rst_art As New ADODB.Recordset
Dim rst_fecha As New ADODB.Recordset
Dim rst_fac As New ADODB.Recordset
Dim rst_ingresos As New ADODB.Recordset
Dim rst_salidas As New ADODB.Recordset
Dim total_mes As Double
Printer.ScaleMode = vbCharacters 'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = 14
    Anulado = "F"
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.3
    Printer.Print Tab(4); "STOCK CONTABLE      :" + "LISTADO GENERAL" + Space(10); "LIRIO DE LOS VALLES SAC"
    Printer.Print Tab(4); "DEL :" + Space(1); str(FechaIni) + Space(1), "Al" + Space(1) + str(FechaFin)
    Printer.Print ""
    Printer.Print Tab(4); "ALMACEN:" + Space(1) + "Bolivar 460" + Space(10) + "USUARIO:" + Space(2) + KEY_VENDEDOR;
    Printer.Print Tab(4); "--------------------------------------------------------------------------------------"
    Printer.Print Tab(2); "CODIGO:" + Space(15) + "DESCRIPCION" + Space(25) + "UND" + Space(4) + "INICIAL" + Space(3) + "INGRESOS" + Space(3) + "SALIDAS" + Space(5) + "STOCK" + Space(3) + "PRECIO CPRA" + Space(3) + "MONTO";
    total_mes = 0
    
    strCadena = "SELECT * FROM "
    strCadena = "SELECT Producto.cProducto, Producto.DescripcionProducto, Unidad.sAbreviatura,Producto.Stock_factura,Producto.PrecioCompra " & _
    "FROM Producto INNER JOIN Unidad ON Producto.cUnidad = Unidad.cUnidad ORDER BY Producto.cProducto"
    Call ConfiguraRst(strCadena)
    rst.MoveFirst
    
    
    
    Printer.CurrentY = Printer.CurrentY + 0.2
    
            For j = 0 To rst.RecordCount - 1
                'PRECIO = Format(Rst(5), "#,##0.00")
                
                strCadena = "SELECT * FROM Kardex WHERE cProducto='" & Trim(rst(0)) & "' AND Kardex.FechaEmision>='" & FechaIni & "' AND Kardex.FechaEmision<='" & FechaFin & "' AND (Kardex.Persona!='PUBLICO EN GENERAL' OR doc_cod!='0088')"
                rst_art.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
                
                If rst_art.RecordCount > 0 Then
                
                 strCadena = "SELECT * FROM Kardex WHERE cProducto='" & Trim(rst(0)) & "' AND Kardex.FechaEmision<='" & FechaIni & "' AND (Kardex.Persona!='PUBLICO EN GENERAL' OR doc_cod!='0088') "
                 rst_fecha.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
                 
                 
                 strCadena = "SELECT * FROM DocumentoVenta INNER JOIN " & _
                 "Detalle_DocumentoVenta ON DocumentoVenta.cDocumentoVenta = Detalle_DocumentoVenta.cDocumentoVenta AND " & _
                 "DocumentoVenta.doc_cod = Detalle_DocumentoVenta.doc_cod AND DocumentoVenta.Alm_cod = Detalle_DocumentoVenta.Alm_Cod AND DocumentoVenta.sSerie = Detalle_DocumentoVenta.sSerie " & _
                 " WHERE Detalle_DocumentoVenta.cProducto='" & Trim(rst(0)) & "' AND DocumentoVenta.dEmisionVenta>'" & FechaFin & "' AND DocumentoVenta.anulado='F' AND DocumentoVenta.dfactura='si'"
                 rst_fac.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
                 
                 If rst_fac.RecordCount > 0 Then
                    Set rst_fac = Nothing
                    strCadena = "SELECT SUM(Detalle_DocumentoVenta.cantidad) AS Expr1 FROM DocumentoVenta INNER JOIN " & _
                 "Detalle_DocumentoVenta ON DocumentoVenta.cDocumentoVenta = Detalle_DocumentoVenta.cDocumentoVenta AND " & _
                 "DocumentoVenta.doc_cod = Detalle_DocumentoVenta.doc_cod AND DocumentoVenta.Alm_cod = Detalle_DocumentoVenta.Alm_Cod AND DocumentoVenta.sSerie = Detalle_DocumentoVenta.sSerie " & _
                 " WHERE Detalle_DocumentoVenta.cProducto='" & Trim(rst(0)) & "' AND DocumentoVenta.dEmisionVenta>'" & FechaFin & "' AND DocumentoVenta.anulado='F' AND DocumentoVenta.dfactura='si'"
                 rst_fac.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
                    cant_des = rst_fac(0)
                 Else
                    cant_des = 0
                 End If
                 
                 'If (rst_fecha.RecordCount > 0) Then
                  '   StrCadena = "SELECT     SUM(Ing_Cant-Sal_Cant) From Kardex WHERE     cProducto = '" & Trim(Rst(0)) & "' AND Kardex.FechaEmision<'" & FechaIni & "' AND (doc_cod!='0088' OR Persona!='PUBLICO EN GENERAL')"
                   ' rst_ini.Open StrCadena, CnBd, adOpenKeyset, adLockOptimistic
                    'saldo_ini = rst_ini(0)
               ' Else
                '   saldo_ini = 0
                'End If
                
                Dim rst_ingreso As Double
                Dim rst_salida As Double
                
                strCadena = "SELECT * FROM DocumentoCompra INNER JOIN " & _
                "Detalle_DocumentoCompra ON DocumentoCompra.cDocumentoCompra = Detalle_DocumentoCompra.cDocumentoCompra AND " & _
                "DocumentoCompra.Alm_cod = Detalle_DocumentoCompra.Alm_cod AND DocumentoCompra.doc_cod = Detalle_DocumentoCompra.doc_cod AND " & _
                "DocumentoCompra.sSerie = Detalle_DocumentoCompra.sSerie WHERE (Detalle_DocumentoCompra.cProducto = '" & Trim(rst(0)) & "') AND (DocumentoCompra.dEmisionCompra >= '" & FechaIni & "') AND " & _
                "(DocumentoCompra.dEmisionCompra <= '" & FechaFin & "') AND (DocumentoCompra.cPersona <> '00004') AND " & _
                "(DocumentoCompra.Anulado = 'F')"
                rst_ingresos.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
                If (rst_ingresos.RecordCount > 0) Then
                    Set rst_ingresos = Nothing
                    
                strCadena = "SELECT sum(Detalle_DocumentoCompra.cantidad) FROM DocumentoCompra INNER JOIN " & _
                "Detalle_DocumentoCompra ON DocumentoCompra.cDocumentoCompra = Detalle_DocumentoCompra.cDocumentoCompra AND " & _
                "DocumentoCompra.Alm_cod = Detalle_DocumentoCompra.Alm_cod AND DocumentoCompra.doc_cod = Detalle_DocumentoCompra.doc_cod AND " & _
                "DocumentoCompra.sSerie = Detalle_DocumentoCompra.sSerie WHERE (Detalle_DocumentoCompra.cProducto = '" & Trim(rst(0)) & "') AND (DocumentoCompra.dEmisionCompra >= '" & FechaIni & "') AND " & _
                "(DocumentoCompra.dEmisionCompra <= '" & FechaFin & "') AND (DocumentoCompra.cPersona <> '00004') AND " & _
                "(DocumentoCompra.Anulado = 'F')"
                rst_ingresos.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
                rst_ingreso = rst_ingresos(0)
                Else
                 
                  rst_ingreso = 0
                End If
                
                
                strCadena = "SELECT Detalle_DocumentoVenta.cantidad FROM DocumentoVenta INNER JOIN Detalle_DocumentoVenta ON DocumentoVenta.cDocumentoVenta = Detalle_DocumentoVenta.cDocumentoVenta AND " & _
                "DocumentoVenta.doc_cod = Detalle_DocumentoVenta.doc_cod AND DocumentoVenta.Alm_cod = Detalle_DocumentoVenta.Alm_Cod AND DocumentoVenta.sSerie = Detalle_DocumentoVenta.sSerie AND " & _
                "DocumentoVenta.dfactura = 'si' AND DocumentoVenta.dEmisionVenta >= '" & FechaIni & "' AND " & _
                "DocumentoVenta.dEmisionVenta <= '" & FechaFin & "' AND DocumentoVenta.Anulado = 'F' AND Detalle_DocumentoVenta.cProducto = '" & Trim(rst(0)) & "'"
                rst_salidas.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
                
                If (rst_salidas.RecordCount > 0) Then
                    Set rst_salidas = Nothing
                    strCadena = "SELECT sum(Detalle_DocumentoVenta.cantidad) FROM DocumentoVenta INNER JOIN Detalle_DocumentoVenta ON DocumentoVenta.cDocumentoVenta = Detalle_DocumentoVenta.cDocumentoVenta AND " & _
                "DocumentoVenta.doc_cod = Detalle_DocumentoVenta.doc_cod AND DocumentoVenta.Alm_cod = Detalle_DocumentoVenta.Alm_Cod AND DocumentoVenta.sSerie = Detalle_DocumentoVenta.sSerie AND " & _
                "DocumentoVenta.dfactura = 'si' AND DocumentoVenta.dEmisionVenta >= '" & FechaIni & "' AND " & _
                "DocumentoVenta.dEmisionVenta <= '" & FechaFin & "' AND DocumentoVenta.Anulado = 'F' AND Detalle_DocumentoVenta.cProducto = '" & Trim(rst(0)) & "'"
                rst_salidas.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
                rst_salida = rst_salidas(0)
                Else
                    rst_salida = 0
                End If
                If ((rst(3) + cant_des) > 0) Then
                
                
                Printer.Print Tab(4); rst(0) & Space(3) & Mid(rst(1) + Space(50), 1, 45) & Space(5) & Mid(rst(2) + Space(10), 1, 4) + Space(5) + Mid(str(rst(3) + cant_des + rst_salida - rst_ingreso) + Space(5), 1, 5) + Space(5) + Mid(str(rst_ingreso) + Space(5), 1, 5) + Space(5) + Mid(str(rst_salida) + Space(5), 1, 5) + Space(5) + Mid(str(rst(3) + cant_des) + Space(5), 1, 6) + Space(2) + Mid(Format(rst(4), "#,##0.00"), 1, 6) + Space(5) + Mid(Format(Format(rst(4), "#,##0.00") * (rst(3) + cant_des), "#,##0.00"), 1, 11)
                total_mes = total_mes + Format(Format(rst(4), "#,##0.00") * (rst(3) - cant_des), "#,##0.00")
                Printer.CurrentY = Printer.CurrentY + 0.4
                End If
                Set rst_ini = Nothing
                Set rst_salidas = Nothing
                Set rst_ingresos = Nothing
                Set rst_fecha = Nothing
                Set rst_fac = Nothing
                rst.MoveNext
                Set rst_art = Nothing
                Else
                    Set rst_art = Nothing
                    Set rst_fecha = Nothing
                     rst.MoveNext
                End If
            Next j
    Printer.Print Tab(20); "TOTAL MES ===================" + Space(10) + Format(total_mes, "#,##0.00")
    
    Set rst = Nothing
    
       
   Printer.EndDoc
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Resp As Double

        
        DteInicio = Format(Me.DtpDesde.Value, "YYYY-mm-dd")
        DteFin = Format(Me.DtpHasta.Value, "YYYY-mm-dd")
                    
               

               
Select Case Button.key
    Case KEY_OK

        Dim Ans As Boolean
        Dim i As Integer
        
        
        strCadena = "SELECT id_producto,nombre_prod,unidad,linea,stock,precio_venta,precio_compra FROM view_producto WHERE habilitado='si' and  ruc='" & KEY_RUC & "' and  id_alm='" & Me.DtcAlmacen.BoundText & "' ORDER BY  id_linea,id_sublinea,nombre_prod "
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptConsolidado_tuco", , App.Path + "\Reportes\")
        
       
        
    Case KEY_PRINT
            If MsgBox("Desea Imprimir el Reporte Contable", vbQuestion + vbYesNo, "MENSAJE PARA EL USER") = vbYes Then
                Call Impresion_contable(DteInicio, DteFin)
            End If
    Case KEY_EXCEL
        
        Me.TlbAcciones.Buttons(KEY_EXCEL).Enabled = False
        
    Case KEY_CANCEL
        Unload Me
End Select

End Sub


 Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
  Set Grilla.Recordset = rst
  If rst.RecordCount > 0 Then
        Me.TlbAcciones.Buttons(KEY_EXCEL).Enabled = True
  Else
        Me.TlbAcciones.Buttons(KEY_EXCEL).Enabled = False
  End If
  Grilla.ColWidth(0) = 650
  Grilla.ColWidth(1) = 4500
  Grilla.ColWidth(2) = 700
  Grilla.ColWidth(3) = 900
  Call DarFormato(Grilla, 3)

Set rst = Nothing
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub







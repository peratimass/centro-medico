VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmPedidosListado 
   BorderStyle     =   0  'None
   Caption         =   "Pedidos"
   ClientHeight    =   8580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOrdenCompra 
      Caption         =   "GENERAR ORDENES DE COMPRA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   11
      Top             =   1560
      Width           =   3735
   End
   Begin VB.CommandButton cmdProveedor 
      Caption         =   "MOSTRAR PROVEEDOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   10
      Top             =   2080
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MOSTRAR TODOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   9
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox TxtcodProveedor 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   225
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10680
      Top             =   7320
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
            Picture         =   "FrmPedidosListado.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedidosListado.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedidosListado.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedidosListado.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedidosListado.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedidosListado.frx":133C
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPedidosListado.frx":165C
            Key             =   "(Reporte)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones1 
      Height          =   4665
      Left            =   14280
      TabIndex        =   2
      Top             =   3240
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   8229
      BandCount       =   1
      ForeColor       =   8388608
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
         Height          =   5670
         Left            =   30
         TabIndex        =   3
         Top             =   375
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   10001
         ButtonWidth     =   1455
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Pagar   "
               Key             =   "(Pagar)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Aceptar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Historial"
               Key             =   "(Anular)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Buscar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reporte"
               Key             =   "(Imprimir)"
               ImageKey        =   "(Reporte)"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgFacturas 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4260
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   5295
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9340
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgAlmacen 
      Height          =   5295
      Left            =   10440
      TabIndex        =   8
      Top             =   3240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   9340
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ALMACEN:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   255
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   270
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   555
      Left            =   120
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "FrmPedidosListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ProcedenciaFacturas As EnumFactura
Public Procedencia As EnumProcede
Dim serie As String
Dim numero As String
Dim Persona As String
Public Sub pedidos()
strCadena = "SELECT PE.id_pedido,PE.fecha,CONCAT(C.doc_abrev,':',PE.serie,'-',PE.numero) as comprobante,P.nombre_completo,A.descripcion,atendido FROM movimiento_pedido PE,comprobantes C,persona P,almacen A WHERE PE.id_alm=A.id_alm AND A.ruc='" & KEY_RUC & "' AND PE.id_doc=C.id_doc AND PE.dni_save=P.dni AND PE.ruc='" & KEY_RUC & "' AND PE.atendido='no' AND PE.anulado='no' ORDER BY PE.fecha,PE.serie,PE.numero "
Call llenarGrid(Me.HfgFacturas, Me)
Call llenar_productos(Me.HfdDetalle)
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Dim tTotal As Double, tSaldo As Double, nsaldo As Double
tTotal = 0
tSaldo = 0
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
  Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 2500
           Grilla.ColWidth(3) = 3000
           Grilla.ColWidth(4) = 3000
      Next
        cabecera = "IDPEDIDO" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "PERSONAL" & vbTab & "ALMACEN"
        Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_pedido") & vbTab & rst("fecha") & vbTab & rst("comprobante") & vbTab & UCase(rst("nombre_completo")) & vbTab & rst("descripcion")
            Grilla.AddItem Fila
           If rst("atendido") = "no" Then
            For k = 0 To 4
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
            Next k
           End If
            Fila = ""
            rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub
Private Sub llenar_proveedor(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim Proveedor As String
strCadena = "SELECT DISTINCT PD.id_producto,P.nombre_prod,U.abreviatura,P.id_proveedor FROM movimiento_pedido PE,movimiento_pedido_detalle PD,producto P,unidad U WHERE PE.id_pedido=PD.id_pedido AND PE.atendido='no' AND PE.anulado='no' AND PE.ruc='" & KEY_RUC & "' AND PD.ruc='" & KEY_RUC & "' AND PD.id_producto=P.id_producto AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' ORDER BY id_proveedor"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 3500
           Grilla.ColWidth(2) = 3500
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1200
        Next
        cabecera = "CODIGO" & vbTab & "PRODUCTO" & vbTab & "PROVEEDOR" & vbTab & "UNIDAD" & vbTab & "CANTIDAD"
        Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            strCadena = "SELECT sum(cantidad) FROM movimiento_pedido_detalle PD,movimiento_pedido PE WHERE PD.id_pedido=PE.id_pedido AND PE.atendido='no' AND PE.ruc='" & KEY_RUC & "' AND PD.ruc='" & KEY_RUC & "' AND PD.id_producto='" & rst("id_producto") & "'"
            Call ConfiguraRstT(strCadena)
            If rst("id_proveedor") <> 0 Then
                Proveedor = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_proveedor"))
            Else
                Proveedor = "*******************"
            End If
            Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & Proveedor & vbTab & rst("abreviatura") & vbTab & Format(rstT(0), "#,##0.00")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub llenar_productos(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
strCadena = "SELECT DISTINCT PD.id_producto,P.nombre_prod,U.abreviatura FROM movimiento_pedido PE,movimiento_pedido_detalle PD,producto P,unidad U WHERE PE.id_pedido=PD.id_pedido AND PE.atendido='no' AND PE.anulado='no' AND PE.ruc='" & KEY_RUC & "' AND PD.ruc='" & KEY_RUC & "' AND PD.id_producto=P.id_producto AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' ORDER BY P.id_proveedor"

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Cols = 0
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 5000
           Grilla.ColWidth(2) = 1500
           Grilla.ColWidth(3) = 2000
        Next
        cabecera = "IDPRODUCTO" & vbTab & "PRODUCTO" & vbTab & "UNIDAD" & vbTab & "CANTIDAD"
        Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            strCadena = "SELECT sum(cantidad) FROM movimiento_pedido_detalle PD,movimiento_pedido PE WHERE PD.id_pedido=PE.id_pedido AND PE.atendido='no' AND PE.ruc='" & KEY_RUC & "' AND PD.ruc='" & KEY_RUC & "' AND PD.id_producto='" & rst("id_producto") & "'"
            Call ConfiguraRstT(strCadena)
            Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & Format(rstT(0), "#,##0.00")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub llenar_pedido(ByVal Grilla As MSHFlexGrid, ByVal id_pedido As Double)
On Error GoTo salir
strCadena = "SELECT PD.id_producto,P.nombre_prod,U.abreviatura,PD.cantidad FROM movimiento_pedido PE,movimiento_pedido_detalle PD,producto P,unidad U WHERE PE.id_pedido=PD.id_pedido AND PE.atendido='no' AND PE.ruc='" & KEY_RUC & "' AND PD.ruc='" & KEY_RUC & "' AND PD.id_producto=P.id_producto AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND PE.id_pedido='" & Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0) & "'"

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 5000
           Grilla.ColWidth(2) = 1500
           Grilla.ColWidth(3) = 2000
        Next
        cabecera = "IDPRODUCTO" & vbTab & "PRODUCTO" & vbTab & "UNIDAD" & vbTab & "CANTIDAD"
        Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & Format(rst("cantidad"), "#,##0.00")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub

Public Sub llenarGrid_Proveedor()
On Error GoTo salir
Dim tTotal As Double
tTotal = 0

strCadena = "SELECT  DocumentoCompra.idCompra, DocumentoCompra.Alm_cod, DocumentoCompra.dEmisionCompra, DocumentoCompra.dVencimiento," & _
" (Comprobantes.doc_abrev + ':'+ DocumentoCompra.sSerie +'-'+ DocumentoCompra.cDocumentoCompra) as DOCUMENTO,DocumentoCompra.Persona, " & _
"DocumentoCompra.moneda , DocumentoCompra.tc, DocumentoCompra.nTotalCompra, DocumentoCompra.Saldo,Anulado,DocumentoCompra.cPersona,seleccion " & _
"FROM DocumentoCompra INNER JOIN Comprobantes ON DocumentoCompra.doc_cod = Comprobantes.doc_cod WHERE (DocumentoCompra.doc_cod='0001' OR DocumentoCompra.doc_cod='0003') AND saldo>0 AND DocumentoCompra.Persona LIKE '%" & Trim(Me.TxtcodProveedor.Text) & "%' ORDER BY DocumentoCompra.dVencimiento ASC "

Call ConfiguraRst(strCadena)
 
If rst.RecordCount < 1 Then
    Me.HfgFacturas.Rows = 1
    HfgFacturas.Clear
    Exit Sub

End If
  
  N = 1
  
   If Me.HfgFacturas.Rows > 0 Then
   HfgFacturas.Clear
   HfgFacturas.Refresh
   HfgFacturas.Rows = 0
   End If
   
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           HfgFacturas.ColWidth(0) = 600
           HfgFacturas.ColWidth(1) = 1000
           HfgFacturas.ColWidth(2) = 1000
           HfgFacturas.ColWidth(3) = 2300
           HfgFacturas.ColWidth(4) = 4000
           HfgFacturas.ColWidth(5) = 500
           HfgFacturas.ColWidth(6) = 500
           HfgFacturas.ColWidth(7) = 1500
           HfgFacturas.ColWidth(8) = 1500
           HfgFacturas.ColWidth(9) = 0
           HfgFacturas.ColWidth(10) = 0
           
           
        Next
         cabecera = "ITEM" & vbTab & "EMISION" & vbTab & "VENCIMIENTO" & vbTab & "COMPROBANTE" & vbTab & "PROVEEDOR" & vbTab & "M" & vbTab & "TC" & vbTab & "TOTAL" & vbTab & "SALDO" & vbTab & "codigounico" & vbTab & "cPersona"
        HfgFacturas.AddItem cabecera
         For k = 0 To 9
                                HfgFacturas.col = k
                                HfgFacturas.Row = 0
                                HfgFacturas.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If (rst("moneda") = "0001") Then
               ' moneda = "S/."
            Else
               ' moneda = "US$."
            End If
            'fila = fila & Str(i + 1) & vbTab & rst("dEmisionCompra") & vbTab & rst("dVencimiento") & vbTab & rst(4) & vbTab & rst("Persona") & vbTab & moneda & vbTab & rst("tc") & vbTab & Format(rst("nTotalCompra"), "#,##0.00") & vbTab & Format(rst("Saldo"), "#,##0.00") & vbTab & rst("idCompra") & vbTab & rst("cPersona")
            If (Fila = "") Then
                x = 1
            End If
          HfgFacturas.AddItem Fila
               
                    
                    
                       If (Trim(rst("Anulado")) = "V") Then
                            For k = 0 To 9
                                HfgFacturas.col = k
                                HfgFacturas.Row = i
                                HfgFacturas.CellBackColor = &H8080FF
                            Next k
                        Else
                        
                        tTotal = tTotal + rst("saldo")
                        End If
                        
                        If (Trim(rst("seleccion")) = "si") Then
                            For k = 0 To 9
                                HfgFacturas.col = k
                                HfgFacturas.Row = i + 1
                                HfgFacturas.CellBackColor = &H80FF80
                            Next k
                       End If
                   
            
            
            Fila = ""
            rst.MoveNext
             
        Next i
    
    
    
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub

Sub llenarGrid_PagosLetras(ByVal Grilla As MSHFlexGrid, ByVal NumeroVenta As String, ByVal serie As String, ByVal Persona As String)

strCadena = "SELECT id_detalle as Codigo, FechaPago, Monto, Operacion FROM  DetallePagos WHERE (DetallePagos.numero='" & NumeroVenta & "'  AND DetallePagos.serie='" & serie & "' AND DetallePagos.cPersona='" & Persona & "')"
On Error GoTo salir
  Call ConfiguraRst(strCadena)

  Grilla.Clear
  Set Grilla.Recordset = rst
  Grilla.Rows = rst.RecordCount
  Grilla.ColWidth(0) = 1500
  Grilla.ColWidth(1) = 2000
  Grilla.ColWidth(2) = 1300
  Grilla.ColWidth(3) = 1300
  Grilla.ColWidth(4) = 0
Call DarFormatoFecha(Grilla, 1)
Call DarFormato(Grilla, 2)


Set rst = Nothing

  'Me.TlbAcciones.Buttons(KEY_PRINT).Enabled = True
  'Me.TlbAcciones.Buttons(KEY_EXIT).Enabled = True
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub cmdbuscar_Click()
Procedencia = buscar
FrmPersona.Show
End Sub

Private Sub cmdOrdenCompra_Click()
Dim id_orden As Double, precio As Double, serie As String, numero As String
 strCadena = "SELECT A.id_doc,C.doc_des as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "'  AND A.id_doc='0110' AND A.id_alm='" & KEY_ALM & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount < 1 Then
    strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0110' AND ruc='" & KEY_RUC & "' ORDER BY serie DESC"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount < 1 Then
        strCadena = "INSERT INTO almacen_comprobante(ruc,id_alm,id_doc,serie,numero,id_formato_impresion)VALUES('" & KEY_RUC & "','" & KEY_ALM & "','0110','001','000001','1')"
    Else
        rstT.MoveFirst
        strCadena = "INSERT INTO almacen_comprobante(ruc,id_alm,id_doc,serie,numero,id_formato_impresion)VALUES('" & KEY_RUC & "','" & KEY_ALM & "','0110','" & formato_item(Val(rstT("serie")) + 1, 3) & "','000001','1')"
    End If
    CnBd.Execute (strCadena)
     
  Else
 End If
strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0110' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
Call ConfiguraRstT(strCadena)
serie = rstT("serie")
numero = rstT("numero")
strCadena = "SELECT DISTINCT P.id_proveedor FROM movimiento_pedido PE,movimiento_pedido_detalle PD,producto P,unidad U WHERE PE.id_pedido=PD.id_pedido AND PE.atendido='no' AND PE.anulado='no' AND PE.ruc='" & KEY_RUC & "' AND PD.ruc='" & KEY_RUC & "' AND PD.id_producto=P.id_producto AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' ORDER BY id_proveedor"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
          
          strCadena = "INSERT INTO orden_compra(id_doc,serie,numero,fecha,id_proveedor,dni_save,ruc)VALUES('0110','" & serie & "','" & numero & "','" & KEY_FECHA & "','" & rst("id_proveedor") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
          CnBd.Execute (strCadena)
           
          id_pedido = LastRegistro("orden_compra", "id_orden")
          Call save_detalleorden(id_pedido, rst("id_proveedor"))
          numero = formato_item(Val(numero) + 1, 6)
          strCadena = "UPDATE almacen_comprobante SET numero='" & numero & "' WHERE id_doc='0110' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
          CnBd.Execute (strCadena)
           
          rst.MoveNext
    Next i
        strCadena = "UPDATE movimiento_pedido SET atendido='si' WHERE atendido='no' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        Unload Me
        FrmOrdenCompra.Show
End If
End Sub

Private Sub save_detalleorden(ByVal id_orden As Double, ByVal id_proveedor As String)
    Dim cantidad As Double
    strCadena = "SELECT DISTINCT PD.id_producto,P.nombre_prod,U.abreviatura,P.id_proveedor FROM movimiento_pedido PE,movimiento_pedido_detalle PD,producto P,unidad U WHERE PE.id_pedido=PD.id_pedido AND PE.atendido='no' AND PE.anulado='no' AND PE.ruc='" & KEY_RUC & "' AND PD.ruc='" & KEY_RUC & "' AND PD.id_producto=P.id_producto AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND P.id_proveedor='" & id_proveedor & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        rstT.MoveFirst
        For j = 0 To rstT.RecordCount - 1
        strCadena = "SELECT sum(cantidad) FROM movimiento_pedido_detalle PD,movimiento_pedido PE WHERE PD.id_pedido=PE.id_pedido AND PE.atendido='no' AND PE.ruc='" & KEY_RUC & "' AND PD.ruc='" & KEY_RUC & "' AND PD.id_producto='" & rstT("id_producto") & "'"
        Call ConfiguraTemporal(strCadena)
        If IsNull(rstTemporal()) = True Then
            cantidad = 0
        Else
            cantidad = rstTemporal(0)
        End If
        '*** PRECIO PROVEEDOR
            precio = BDBuscarCampoEmpresa("producto", "precio_compra", "id_producto", rstT("id_producto"), id_proveedor)
            If IsNull(precio) = True Then
                precio = 0
            End If
        '********************
        strCadena = "INSERT INTO orden_compra_detalle(id_orden,id_producto,cantidad,precio,total,ruc)VALUES('" & id_orden & "','" & rstT("id_producto") & "','" & cantidad & "','" & Val(precio) & "','" & Val(precio) * cantidad & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        rstT.MoveNext
        Next j
        
    End If
End Sub
Private Sub cmdProveedor_Click()
Call llenar_proveedor(Me.HfdDetalle)
End Sub

Private Sub Command1_Click()
Call pedidos
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  Me.DtcAlmacen.Enabled = False
  
Call pedidos

  
End Sub

Private Sub HfdDetalle_SelChange()
If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    strCadena = "SELECT id_alm,descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'"
    Call llenar_almacenes(Me.HfgAlmacen)
End If
End Sub

Private Sub HfgFacturas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) > 0 Then
       ' Call FrmPedido.nuevo
        strCadena = "SELECT * FROM movimiento_pedido WHERE id_pedido='" & Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        FrmPedido.DtcAlmacen.Enabled = True
        FrmPedido.DtcTipoDoc.Enabled = True
        FrmPedido.TxtSerie.Enabled = True
        FrmPedido.TxtNumeroDoc.Enabled = True
        
        FrmPedido.DtcTipoDoc.BoundText = rst("id_doc")
        FrmPedido.DtcTipoDoc.BoundText = rst("id_doc")
        FrmPedido.DtpActual.Value = rst("fecha")
        FrmPedido.TxtSerie.Text = rst("serie")
        FrmPedido.TxtNumeroDoc.Text = rst("numero")
        Procedencia = buscar
        Call FrmPedido.buscar_comprobante(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0))
        Set rst = Nothing
        Unload Me
    End If
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_PAGAR
    
      strCadena = "UPDATE movimiento_compra SET seleccion='si' WHERE id_compra='" & Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
      CnBd.Execute (strCadena)
       
      strCadena = "UPDATE movimiento_compra SET seleccion='no' WHERE id_compra<>'" & Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0) & "' AND ruc='" & KEY_RUC & "' AND id_proveedor='" & Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 10) & "' AND saldo>0"
      CnBd.Execute (strCadena)
       
      FrmComprasPagos.Show
            
    Case KEY_EXIT
        Unload Me
    Case "(Salir)"
      Unload Me
  End Select
End Sub




Private Sub HfgFacturas_SelChange()
If Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) > 0 Then
    Call llenar_pedido(Me.HfdDetalle, Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0))
End If


End Sub
Public Sub llenar_almacenes(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim tTotal As Double, Suma As Double
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 2500
            Grilla.ColWidth(2) = 1100
            
        Next
        cabecera = "IDALM" & vbTab & "ALMACEN" & vbTab & "CANTIDAD"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            strCadena = "SELECT sum(PD.cantidad) FROM movimiento_pedido_detalle PD,movimiento_pedido PE WHERE PD.id_pedido=PE.id_pedido AND PE.id_alm='" & rst("id_alm") & "' AND PD.id_producto='" & Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0) & "' AND PE.ruc='" & KEY_RUC & "' AND PE.atendido='no' AND  PD.ruc='" & KEY_RUC & "'"
            Call ConfiguraRstT(strCadena)
            If IsNull(rstT(0)) = True Then
                Suma = 0
            Else
                Suma = rstT(0)
            End If
                
            tTotal = tTotal + Suma
            Fila = rst("id_alm") & vbTab & rst("descripcion") & vbTab & Format(Suma, "#,##0.00")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
       For k = 0 To 2
            Grilla.col = 2
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
    
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub




Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_NEW
    Case KEY_PRINT
          strCadena = "SELECT * FROM DocumentoCompra WHERE idCompra='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 7)) & "'"
          Call ConfiguraRst(strCadena)
          cod_doc = rst("doc_cod")
          serie_doc = rst("sSerie")
          numero_doc = rst("cDocumentoCompra")
          FrmCompras.DtcAlmacen.Enabled = True
          FrmCompras.DtcTipoDoc.Enabled = True
          FrmCompras.TxtSerie.Enabled = True
          FrmCompras.TxtNumeroDoc.Enabled = True
          FrmCompras.DtcTipoDoc.BoundText = cod_doc
          FrmCompras.Txtdoc_cod.Text = cod_doc
          FrmCompras.TxtSerie.Text = serie_doc
          FrmCompras.TxtNumeroDoc.Text = numero_doc
          Procedencia = buscar
          Call FrmCompras.buscar_comprobante(Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 7)))
          FrmCompras.Top = 50
          Exit Sub
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.key
    Case KEY_DELETE
          
    Case KEY_PRINT
         ' If Val(Me.Hfpagos.TextMatrix(Me.Hfpagos.Row, 4)) > 0 Then
        'strCadena = "SELECT     Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, movimiento_caja.cPersona, " & _
        "movimiento_caja.descripcion_per, Persona.sDireccionCliente1, Persona.Per_Ruc, movimiento_caja.fecha_valor," & _
        "movimiento_caja.cambio , movimiento_caja.glosa, centro_costos.descripcion, movimiento_caja.Monto,movimiento_caja.monto_letras " & _
        "FROM movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
        "centro_costos ON movimiento_caja.id_costo = centro_costos.id_costo INNER JOIN " & _
        "Persona ON movimiento_caja.cPersona = Persona.cPersona WHERE movimiento_caja.codigo='" & Val(Me.Hfpagos.TextMatrix(Me.Hfpagos.Row, 4)) & "'"
        'Call ConfiguraRst(strCadena)
        'Ans = ShowMultiReport(rst, "RptReciboIngreso", , App.Path + "\Reportes\")
        'Set rst = Nothing
         ' End If
    Case KEY_EXIT
        Unload Me
     
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub


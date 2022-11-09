VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmListadoFacturasVarios 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtcodProveedor 
      Appearance      =   0  'Flat
      Height          =   325
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Height          =   325
      Left            =   12000
      Picture         =   "FrmFacturasVariasList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox tXTrUC 
      Appearance      =   0  'Flat
      Height          =   325
      Left            =   10320
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   6360
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFacturasVariasList.frx":014A
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFacturasVariasList.frx":059E
            Key             =   "(Visualizar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFacturasVariasList.frx":0B38
            Key             =   "(Pagos)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFacturasVariasList.frx":10D2
            Key             =   "(Monto)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   5745
      Left            =   14040
      TabIndex        =   3
      Top             =   720
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   10134
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   5745
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   5685
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      BandStyle1      =   1
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   4
         Top             =   345
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1429
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
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
               Caption         =   "Detalle"
               Key             =   "(Modificar)"
               ImageKey        =   "(Visualizar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Pagos"
               Key             =   "(Eliminar)"
               ImageKey        =   "(Pagos)"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgFacturas 
      Height          =   7575
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   13361
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   360
      Left            =   1320
      TabIndex        =   6
      Top             =   225
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   13695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ALMACEN:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   250
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   8685
      Left            =   0
      Top             =   0
      Width           =   15270
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   270
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "RUC:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   9840
      TabIndex        =   7
      Top             =   270
      Width           =   375
   End
End
Attribute VB_Name = "FrmListadoFacturasVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ProcedenciaFacturas As EnumFactura
Public Procedencia As EnumProcede
Dim serie As String
Dim Numero As String
Dim Persona As String
Public Sub facturas()

    

    Call llenarGrid(Me.HfgFacturas, Me)

End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Dim tTotal As Double
tTotal = 0



strCadena = "SELECT  DocumentoCompra.idCompra, DocumentoCompra.Alm_cod, DocumentoCompra.dEmisionCompra, DocumentoCompra.dVencimiento," & _
" (Comprobantes.doc_abrev + ':'+ DocumentoCompra.sSerie +'-'+ DocumentoCompra.cDocumentoCompra) as DOCUMENTO,DocumentoCompra.Persona, " & _
"DocumentoCompra.moneda , DocumentoCompra.tc, DocumentoCompra.nTotalCompra, DocumentoCompra.Saldo,Anulado,DocumentoCompra.cPersona " & _
"FROM DocumentoCompra INNER JOIN Comprobantes ON DocumentoCompra.doc_cod = Comprobantes.doc_cod WHERE (DocumentoCompra.doc_cod<>'0077' ) AND saldo>0 AND tipo_factura='egreso' ORDER BY DocumentoCompra.dVencimiento ASC "
'frmNuevoComprobante.Procedencia = Neutro
    


Call ConfiguraRst(strCadena)
 
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
  
  N = 1
  
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 600
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 2300
           Grilla.ColWidth(4) = 4000
           Grilla.ColWidth(5) = 500
           Grilla.ColWidth(6) = 500
           Grilla.ColWidth(7) = 1500
           Grilla.ColWidth(8) = 1500
           Grilla.ColWidth(9) = 0
           
           
           
        Next
         cabecera = "ITEM" & vbTab & "EMISION" & vbTab & "VENCIMIENTO" & vbTab & "COMPROBANTE" & vbTab & "PROVEEDOR" & vbTab & "M" & vbTab & "TC" & vbTab & "TOTAL" & vbTab & "SALDO" & vbTab & "codigounico" & vbTab & "cpersona"
        Grilla.AddItem cabecera
         For k = 0 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
           ' If (rst("moneda") = "0001") Then
                'moneda = "S/."
            'Else
                'moneda = "US$."
           ' End If
           ' fila = fila & Str(i + 1) & vbTab & rst("dEmisionCompra") & vbTab & rst("dVencimiento") & vbTab & rst(4) & vbTab & rst("Persona") & vbTab & moneda & vbTab & rst("tc") & vbTab & Format(rst("nTotalCompra"), "#,##0.00") & vbTab & Format(rst("Saldo"), "#,##0.00") & vbTab & rst("idCompra") & vbTab & rst("cPersona")
            If (Fila = "") Then
                x = 1
            End If
          Grilla.AddItem Fila
               
                    
                    
            '           If (Trim(rst("Anulado")) = "V") Then
                            For k = 0 To 9
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &H8080FF
                            Next k
              '          Else
                        
                        tTotal = tTotal + rst("saldo")
              '          End If
                        
                   
            
            
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
"FROM DocumentoCompra INNER JOIN Comprobantes ON DocumentoCompra.doc_cod = Comprobantes.doc_cod WHERE (DocumentoCompra.doc_cod='0001' OR DocumentoCompra.doc_cod='0003') AND saldo>0 AND tipo_factura='egreso' AND DocumentoCompra.Persona LIKE '%" & Trim(Me.TxtcodProveedor.Text) & "%' ORDER BY DocumentoCompra.dVencimiento ASC "

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
           ' fila = fila & Str(i + 1) & vbTab & rst("dEmisionCompra") & vbTab & rst("dVencimiento") & vbTab & rst(4) & vbTab & rst("Persona") & vbTab & moneda & vbTab & rst("tc") & vbTab & Format(rst("nTotalCompra"), "#,##0.00") & vbTab & Format(rst("Saldo"), "#,##0.00") & vbTab & rst("idCompra") & vbTab & rst("cPersona")
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
Public Sub llenarGrid_ProveedorRUC()
On Error GoTo salir
Dim tTotal As Double
tTotal = 0

strCadena = "SELECT     DocumentoCompra.idCompra, DocumentoCompra.Alm_cod, DocumentoCompra.dEmisionCompra, DocumentoCompra.dVencimiento, " & _
"                      Comprobantes.doc_abrev, DocumentoCompra.sSerie, DocumentoCompra.cDocumentoCompra, DocumentoCompra.Persona, " & _
"                      DocumentoCompra.moneda, DocumentoCompra.tc, DocumentoCompra.nTotalCompra, DocumentoCompra.saldo, " & _
"                      DocumentoCompra.Anulado ,   DocumentoCompra.cPersona,seleccion FROM         DocumentoCompra INNER JOIN " & _
"                      Comprobantes ON DocumentoCompra.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
"                      Persona ON DocumentoCompra.cPersona = Persona.cPersona WHERE (DocumentoCompra.doc_cod='0001' OR DocumentoCompra.doc_cod='0003') AND saldo>0 AND tipo_factura='egreso' AND Persona.Per_Ruc LIKE '%" & Trim(Me.TxtRuc.Text) & "%' ORDER BY DocumentoCompra.dVencimiento ASC "


Call ConfiguraRst(strCadena)
 
If rst.RecordCount < 1 Then
    Me.HfgFacturas.Rows = 1
    HfgFacturas.Clear
    Exit Sub

End If
  
  N = 1
  
   HfgFacturas.Clear
   HfgFacturas.Refresh
   HfgFacturas.Rows = 0
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
           
           
           
        Next
         cabecera = "ITEM" & vbTab & "EMISION" & vbTab & "VENCIMIENTO" & vbTab & "COMPROBANTE" & vbTab & "PROVEEDOR" & vbTab & "M" & vbTab & "TC" & vbTab & "TOTAL" & vbTab & "SALDO" & vbTab & "codigounico" & vbTab & "cpersona"
        HfgFacturas.AddItem cabecera
         For k = 0 To 9
                                HfgFacturas.col = k
                                HfgFacturas.Row = 0
                                HfgFacturas.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If (rst("moneda") = "0001") Then
              '  moneda = "S/."
            Else
               ' moneda = "US$."
            End If
            If IsNull(rst("dVencimiento")) = True Then
                vencimiento = rst("dEmisionCompra")
            End If
            'fila = fila & Str(i + 1) & vbTab & rst("dEmisionCompra") & vbTab & vencimiento & vbTab & rst(4) & vbTab & rst("Persona") & vbTab & moneda & vbTab & rst("tc") & vbTab & Format(rst("nTotalCompra"), "#,##0.00") & vbTab & Format(rst("Saldo"), "#,##0.00") & vbTab & rst("idCompra") & vbTab & rst("cPersona")
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
                                HfgFacturas.Row = i - 1
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

Private Sub Command1_Click()
Call facturas
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
strCadena = "SELECT Alm_cod as Codigo, Alm_des as Descripcion FROM Almacen " & _
  " ORDER BY Alm_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = "0001"
  Me.DtcAlmacen.Enabled = False
  Set rst = Nothing
Call facturas
Me.TlbAcciones.Buttons(KEY_NEW).Enabled = False
  
End Sub
Private Sub HfgFacturas_Click()
If Me.HfgFacturas.Rows > 0 Then
    
    
    Me.TlbAcciones.Buttons(KEY_NEW).Enabled = True
Else
    Me.TlbAcciones.Buttons(KEY_NEW).Enabled = False
End If
End Sub

Private Sub HfgFacturas_DblClick()
Dim estado As String
If Me.HfgFacturas.Rows > 0 Then
    strCadena = "SELECT * FROM DocumentoCompra WHERE idCompra='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 9)) & "'"
    Call ConfiguraRst(strCadena)
    estado = rst("seleccion")
    Set rst = Nothing
    If estado = "si" Then
        estado = "no"
    Else
        estado = "si"
    End If

    If Trim(Me.TxtcodProveedor.Text) <> "" Then
    
        strCadena = "UPDATE DocumentoCompra SET seleccion='" & Trim(estado) & "',IdUsuario='" & Trim(KEY_USUARIO) & "' WHERE idCompra='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 9)) & "'"
        CnBd.Execute (strCadena)
         
        Call llenarGrid_Proveedor
        Exit Sub
       
   End If
    If Trim(Me.TxtRuc.Text) <> "" Then
        strCadena = "UPDATE DocumentoCompra SET seleccion='" & Trim(estado) & "',IdUsuario='" & Trim(KEY_USUARIO) & "' WHERE idCompra='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 9)) & "'"
        CnBd.Execute (strCadena)
         
        Call llenarGrid_ProveedorRUC
        Exit Sub
       
   End If
   
End If
End Sub

Private Sub HfgFacturas_KeyPress(KeyAscii As Integer)
Dim facturas As String
Dim total_facturas As Single
If Me.HfgFacturas.Rows > 0 Then
    If KeyAscii = 13 Then
        If frmNuevoComprobante.Procedencia = buscar Then
            
            facturas = ""
            total_facturas = 0
            strCadena = "SELECT Comprobantes.doc_abrev, DocumentoCompra.sSerie, DocumentoCompra.cDocumentoCompra, DocumentoCompra.saldo, " & _
            " DocumentoCompra.cPersona FROM DocumentoCompra INNER JOIN Comprobantes ON DocumentoCompra.doc_cod = Comprobantes.doc_cod WHERE seleccion='si' AND IdUsuario='" & Trim(KEY_USUARIO) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 1 Then
                frmNuevoComprobante.FrameComprobante.Visible = False
                frmNuevoComprobante.FrameFacturas.Visible = True
                frmNuevoComprobante.FrameCCostos.Visible = False
                frmNuevoComprobante.Procedencia = Neutro
                rst.MoveFirst
                For i = 0 To rst.RecordCount - 1
                   
                    total_facturas = total_facturas + rst("saldo")
                    rst.MoveNext
                Next i
                Set rst = Nothing
                strCadena = "SELECT Per_Ruc,NombrePersona FROM Persona WHERE  cPersona='" & Trim(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 10)) & "'"
                Call ConfiguraRst(strCadena)
                frmNuevoComprobante.TxtRuc.Text = rst("Per_Ruc")
                Call frmNuevoComprobante.buscarPersona
                frmNuevoComprobante.TxtTotalFacturas.Text = Format(total_facturas, "###0.00")
                Call frmNuevoComprobante.llenarFacturas
                frmNuevoComprobante.TxtGlosa.Text = facturas
                Unload Me
                frmNuevoComprobante.TxtMonto1.SetFocus
                Exit Sub
            End If
            
            
            strCadena = "SELECT * FROM DocumentoCompra WHERE idCompra='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 9)) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                frmNuevoComprobante.FrameComprobante.Visible = True
                frmNuevoComprobante.FrameFacturas.Visible = False
                frmNuevoComprobante.TxtTD.Text = rst("doc_cod")
                frmNuevoComprobante.TxtSerie.Text = rst("sSerie")
                frmNuevoComprobante.txtnumero.Text = rst("cDocumentoCompra")
                frmNuevoComprobante.TxtMontoFactura.Text = Format(rst("nTotalCompra"), "###0.00")
                Unload Me
                frmNuevoComprobante.Procedencia = Neutro
                Call frmNuevoComprobante.precionar
                Exit Sub
            End If
        End If
    End If
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_NEW
      FrmMovimientoCaja.Show
     
    Case KEY_UPDATE
         
         strCadena = "SELECT * FROM DocumentoCompra WHERE idCompra='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 9)) & "'"
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
         
    Call FrmCompras.buscar_comprobante(Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 9)))
         FrmCompras.Top = 50
    Case KEY_DELETE
        
    Case KEY_EXIT
        Unload Me
    Case "(Salir)"
      Unload Me
  End Select
End Sub




Private Sub TxtcodProveedor_Change()
If Trim(Me.TxtcodProveedor.Text) <> "" Then
    Me.TxtRuc.Text = ""
    Call llenarGrid_Proveedor
End If
End Sub

Private Sub TxtRuc_Change()
If Trim(Me.TxtRuc.Text) <> "" Then
    Me.TxtcodProveedor.Text = ""
    Call llenarGrid_ProveedorRUC
End If
End Sub



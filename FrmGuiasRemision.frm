VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmGuiasRemision 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   17970
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn cmdbuscar 
      Height          =   300
      Left            =   15600
      TabIndex        =   11
      Top             =   240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      BTYPE           =   4
      TX              =   "BUSCAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmGuiasRemision.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   300
      Left            =   12480
      TabIndex        =   9
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
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
      Format          =   179503105
      CurrentDate     =   42381
   End
   Begin VB.TextBox TxtFactura 
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
      Left            =   11040
      TabIndex        =   8
      Top             =   240
      Width           =   975
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
      Left            =   5880
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox tXTrUC 
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
      Left            =   8640
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      Top             =   225
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgFacturas 
      Height          =   7455
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   13150
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   300
      Left            =   14160
      TabIndex        =   10
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
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
      Format          =   179503105
      CurrentDate     =   42381
   End
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   855
      Left            =   16680
      TabIndex        =   13
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "CERRAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmGuiasRemision.frx":001C
      PICN            =   "FrmGuiasRemision.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdReporte 
      Height          =   855
      Left            =   16680
      TabIndex        =   14
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "REPORTE"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmGuiasRemision.frx":305F
      PICN            =   "FrmGuiasRemision.frx":307B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AL"
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
      Left            =   13920
      TabIndex        =   12
      Top             =   280
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NRO GUIA :"
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
      Left            =   10080
      TabIndex        =   7
      Top             =   270
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ALMACEN:"
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
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   255
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESTINATARIO:"
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
      Left            =   4560
      TabIndex        =   5
      Top             =   270
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DNI / RUC :"
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
      Height          =   255
      Left            =   7680
      TabIndex        =   4
      Top             =   270
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Inside Solid
      Height          =   645
      Left            =   120
      Top             =   120
      Width           =   16455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   8475
      Left            =   0
      Top             =   0
      Width           =   17970
   End
End
Attribute VB_Name = "FrmGuiasRemision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ProcedenciaFacturas As EnumFactura
Public Procedencia As EnumProcede
Dim serie As String
Dim numero As String
Dim Persona As String
Public Sub facturas()

strCadena = "SELECT * FROM view_guia_remision WHERE fecha='" & KEY_FECHA & "' and  ruc='" & KEY_RUC & "'   "

Call llenarGrid(Me.HfgFacturas, Me)

End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub

End If
   N = 1
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 1800
           Grilla.ColWidth(4) = 1800
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 3000
           Grilla.ColWidth(8) = 2000
            Grilla.ColWidth(9) = 2200
        Next
        cabecera = "IDTRANSFERENCIA" & vbTab & "EMISION" & vbTab & "GUIA REMISION " & vbTab & "GUIA VINCULADA" & vbTab & "COMPROBANTE VENTA" & vbTab & "VALORIZADO" & vbTab & "RUC/DNI" & vbTab & "DESTINATARIO" & vbTab & "UBIGEO" & vbTab & "ESTADO"
        Grilla.AddItem cabecera
         For k = 1 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
         Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If IsNull(rst("vinculada")) = True Then
                in_vinculada = "     [     ]"
                If rst("id_doc") = "0031" Then
                    in_numero = "     [     ]"
                    in_vinculada = rst("numero")
                Else
                    in_numero = rst("numero")
                End If
            Else
                in_numero = rst("numero")
                in_vinculada = rst("vinculada")
            End If
            
            
            
           
            
            Fila = rst("id_transferencia") & vbTab & Format(rst("fecha"), "dd-mm-YYYY") & vbTab & in_numero & vbTab & in_vinculada & vbTab & rst("venta") & vbTab & Format(rst("valor_mercaderia"), "#,##0.00") & vbTab & rst("id_destinatario") & vbTab & rst("destinatario") & vbTab & rst("ubigeo") & vbTab & rst("estado")
            Grilla.AddItem Fila
            If rst("anulado") = "si" Then
            For k = 1 To 9
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &HC0C0FF
            Next k
            End If
          
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
                X = 1
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
"                      Persona ON DocumentoCompra.cPersona = Persona.cPersona WHERE (DocumentoCompra.doc_cod='0001' OR DocumentoCompra.doc_cod='0003') AND saldo>0 AND Persona.Per_Ruc LIKE '%" & Trim(Me.txtRuc.Text) & "%' ORDER BY DocumentoCompra.dVencimiento ASC "


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
               ' moneda = "S/."
            Else
               ' moneda = "US$."
            End If
            If IsNull(rst("dVencimiento")) = True Then
                vencimiento = rst("dEmisionCompra")
            End If
            'fila = fila & Str(i + 1) & vbTab & rst("dEmisionCompra") & vbTab & vencimiento & vbTab & rst(4) & vbTab & rst("Persona") & vbTab & moneda & vbTab & rst("tc") & vbTab & Format(rst("nTotalCompra"), "#,##0.00") & vbTab & Format(rst("Saldo"), "#,##0.00") & vbTab & rst("idCompra") & vbTab & rst("cPersona")
            If (Fila = "") Then
                X = 1
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

Private Sub cmdBuscar_Click()

strCadena = "SELECT * FROM view_guia_remision WHERE fecha>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' "
Call llenarGrid(Me.HfgFacturas, Me)



End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdhistorial_Click()

End Sub

Private Sub cmdReporte_Click()
Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant

arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"



arr(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
arr(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")


param = arr()


    strCadena = "SELECT id_transferencia,fecha,numero,id_destinatario,destinatario,direccion,ruc  FROM view_guia_remision WHERE fecha>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptGuiasremision", param, App.Path + "\Reportes\")

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub DtcAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_guia_remision WHERE fecha>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and id_alm='" & Me.DtcAlmacen.BoundText & "' and  ruc='" & KEY_RUC & "' "
    Call llenarGrid(Me.HfgFacturas, Me)
End If
End Sub

Private Sub Form_Load()

CenterForm Me
Me.Top = 500
strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM

Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA
Me.DtcAlmacen.Enabled = False



  'Me.Toolbar1.Buttons(KEY_PRINT).Enabled = False
  'Me.Toolbar1.Buttons(KEY_DELETE).Enabled = False
  'Me.TlbGrabar.Buttons(KEY_PAGAR).Enabled = False
  'Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False

Call facturas

  
End Sub

Private Sub HfgFacturas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) > 0 Then
        
        
        
        strCadena = "SELECT * FROM movimiento_transferencia WHERE id_transferencia='" & Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstIN(strCadena)
        If rstIN.RecordCount > 0 Then
        FrmTransferencias.chk_consultar.Value = 1
       ' Call FrmTransferencias.llenar_serie_guia("0009")
        
        FrmTransferencias.DtcTipoDoc.BoundText = rstIN("id_doc")
        FrmTransferencias.DtcSerieGuia.BoundText = rstIN("serie")
        Call FrmTransferencias.nuevo
        
        FrmTransferencias.TxtNumeroDoc.Text = rstIN("numero")
        FrmTransferencias.TxtMarcayPlaca.Text = rstIN("marca_placa")
        FrmTransferencias.TxtPlaca.Text = rstIN("placa")
        FrmTransferencias.TxtLicencia.Text = rstIN("licencia")
        FrmTransferencias.txtcertificado.Text = rstIN("certificado")
        
        
        Call FrmTransferencias.buscar_comprobante(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0))
        
        Unload Me
        Exit Sub
       End If
    End If
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub




Private Sub HfgFacturas_SelChange()
'If Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 7)) > 0 Then
 '   Call LlenarPagos(Me.Hfpagos, Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 7))
  '  Me.TlbGrabar.Buttons(KEY_PAGAR).Enabled = True
   ' Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
    'Me.Toolbar1.Buttons(KEY_PRINT).Enabled = False
    'Me.Toolbar1.Buttons(KEY_DELETE).Enabled = False
'Else
  '  Me.TlbGrabar.Buttons(KEY_PAGAR).Enabled = False
 '   Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
'End If



End Sub
Public Sub LlenarPagos(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT     id,fecha, monto,documento ,anulado FROM  mis_cuentas_det WHERE IdMovimiento='" & idVenta & "' AND Ruc='" & KEY_RUC & "' AND tipo_trans='E'"
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
            Grilla.ColWidth(1) = 700
            Grilla.ColWidth(2) = 1300
            Grilla.ColWidth(3) = 2300
            Grilla.ColWidth(4) = 0
        Next
        cabecera = "IDCAJA" & vbTab & "ITEM" & vbTab & "FECHA" & vbTab & "MONTO" & vbTab & "OPERACION"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("id") & vbTab & (i + 1) & vbTab & rst("fecha") & vbTab & Format(rst("monto"), "#,##0.00") & vbTab & rst("documento")
            Grilla.AddItem Fila
            
            Fila = ""
            If rst("anulado") = "V" Then
                For k = 0 To 3
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
            Else
            tTotal = tTotal + rst("monto")
            End If
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
       For k = 0 To 3
            Grilla.col = 3
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
          FrmCompras.txtserie.Enabled = True
          FrmCompras.TxtNumeroDoc.Enabled = True
          FrmCompras.DtcTipoDoc.BoundText = cod_doc
          FrmCompras.Txtdoc_cod.Text = cod_doc
          FrmCompras.txtserie.Text = serie_doc
          FrmCompras.TxtNumeroDoc.Text = numero_doc
          Procedencia = buscar
          Call FrmCompras.buscar_comprobante(Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 7)))
          FrmCompras.Top = 50
          Exit Sub
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
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
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub TxtcodProveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_guia_remision WHERE destinatario like '%" & Trim(Me.TxtcodProveedor.Text) & "%' and  ruc='" & KEY_RUC & "' "
    Call llenarGrid(Me.HfgFacturas, Me)
End If
End Sub

Private Sub TxtFactura_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_guia_remision WHERE (numero LIKE '%" & Trim(Me.TxtFactura.Text) & "%' OR vinculada LIKE '%" & Trim(Me.TxtFactura.Text) & "%' ) and  ruc='" & KEY_RUC & "' "
    Call llenarGrid(Me.HfgFacturas, Me)


End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   strCadena = "SELECT * FROM view_guia_remision WHERE id_destinatario='" & Trim(Me.txtRuc.Text) & "' and  ruc='" & KEY_RUC & "' "
    Call llenarGrid(Me.HfgFacturas, Me)
End If
End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmPartematerialLista 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16065
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   16065
   ShowInTaskbar   =   0   'False
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
      Left            =   5280
      TabIndex        =   3
      Top             =   240
      Width           =   1695
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
      Left            =   8160
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Height          =   325
      Left            =   13200
      Picture         =   "FrmPartematerialLista.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   495
   End
   Begin MSComCtl2.DTPicker DtpFechaIni 
      Height          =   300
      Left            =   10560
      TabIndex        =   1
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   170786817
      CurrentDate     =   41303
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   225
      Width           =   3375
      _ExtentX        =   5953
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
      Left            =   12600
      Top             =   2880
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
            Picture         =   "FrmPartematerialLista.frx":014A
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPartematerialLista.frx":059E
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPartematerialLista.frx":08BE
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPartematerialLista.frx":0D12
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPartematerialLista.frx":1166
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPartematerialLista.frx":1486
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPartematerialLista.frx":17A6
            Key             =   "(Reporte)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones1 
      Height          =   6585
      Left            =   15120
      TabIndex        =   5
      Top             =   840
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   11615
      BandCount       =   1
      ForeColor       =   8388608
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   6585
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   2430
         Left            =   30
         TabIndex        =   6
         Top             =   375
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   4286
         ButtonWidth     =   1508
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Reporte"
               Key             =   "(reporte)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Aceptar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgFacturas 
      Height          =   6615
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   11668
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
   Begin MSComCtl2.DTPicker DtpFechaFin 
      Height          =   300
      Left            =   11880
      TabIndex        =   12
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   170786817
      CurrentDate     =   41303
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA :"
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
      Height          =   195
      Left            =   9960
      TabIndex        =   11
      Top             =   270
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUCURSAL :"
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
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   255
      Width           =   885
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "OPERADOR :"
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
      Left            =   7200
      TabIndex        =   9
      Top             =   270
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLACA :"
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
      Height          =   195
      Left            =   4680
      TabIndex        =   8
      Top             =   270
      Width           =   585
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   645
      Left            =   120
      Top             =   120
      Width           =   14895
   End
End
Attribute VB_Name = "FrmPartematerialLista"
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
strCadena = "SELECT P.id_material,P.fecha,CONCAT(C.doc_abrev,':',P.serie,'-',P.numero) as comprobante,P.id_cliente,CONCAT(TT.descripcion,'-',T.placa) as transporte,anulado,dni_save,P.tiempo_ruta FROM parte_material P,comprobantes C,transporte T,transporte_tipo TT WHERE T.id_tipo_transporte=TT.id_tipo_transporte AND P.id_transporte=T.id_transporte AND P.id_doc=C.id_doc AND P.ruc='" & KEY_RUC & "' AND T.ruc='" & KEY_RUC & "' AND TT.ruc='" & KEY_RUC & "'   ORDER BY P.fecha ASC  "
Call llenarGrid(Me.HfgFacturas, Me)
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Dim Factura As String
Dim tViajes As Double, tHoras As Variant


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
           Grilla.ColWidth(2) = 2300
           Grilla.ColWidth(3) = 4000
           Grilla.ColWidth(4) = 2300
           Grilla.ColWidth(5) = 3000
           Grilla.ColWidth(6) = 1500
           
        Next
        cabecera = "IDPARTE" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "OPERADOR" & vbTab & "TRANSPORTE" & vbTab & "VIAJES" & vbTab & "HORAS"
        Grilla.AddItem cabecera
         For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
         Next k
        tViajes = 0
        tHoras = 0
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If Len(rst("id_cliente")) > 1 Then
                beneficiario = BDBuscarCampo("persona", "nombre_completo", "dni", rst("id_cliente"))
            Else
                beneficiario = ""
            End If
            viajes = Val(BdAcumuladoCampo("parte_material_detalle", "id_parte", rst("id_material"), "cantidad"))
            
            Fila = rst("id_material") & vbTab & rst("fecha") & vbTab & rst("comprobante") & vbTab & UCase(beneficiario) & vbTab & rst("transporte") & vbTab & str(viajes) & vbTab & rst("tiempo_ruta")
            Grilla.AddItem Fila
            tViajes = tViajes + viajes

            
            tHoras = Format(tHoras, "hh:mm:ss")
            final = Format(rst("tiempo_ruta"), "hh:mm:ss")
            tHoras = Format(TimeValue(final) + TimeValue(tHoras), "hh:mm:ss")
            
            If rst("anulado") = "si" Then
            For k = 0 To 6
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &HC0C0FF
            Next k
            End If
            Fila = ""
            rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & tViajes & vbTab & tHoras
        Grilla.AddItem Fila
         For k = 4 To 6
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HDFDFE0
         Next k
    
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
Public Sub llenarGrid_ProveedorRUC()
On Error GoTo salir
Dim tTotal As Double
tTotal = 0

strCadena = "SELECT     DocumentoCompra.idCompra, DocumentoCompra.Alm_cod, DocumentoCompra.dEmisionCompra, DocumentoCompra.dVencimiento, " & _
"                      Comprobantes.doc_abrev, DocumentoCompra.sSerie, DocumentoCompra.cDocumentoCompra, DocumentoCompra.Persona, " & _
"                      DocumentoCompra.moneda, DocumentoCompra.tc, DocumentoCompra.nTotalCompra, DocumentoCompra.saldo, " & _
"                      DocumentoCompra.Anulado ,   DocumentoCompra.cPersona,seleccion FROM         DocumentoCompra INNER JOIN " & _
"                      Comprobantes ON DocumentoCompra.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
"                      Persona ON DocumentoCompra.cPersona = Persona.cPersona WHERE (DocumentoCompra.doc_cod='0001' OR DocumentoCompra.doc_cod='0003') AND saldo>0 AND Persona.Per_Ruc LIKE '%" & Trim(Me.txtruc.Text) & "%' ORDER BY DocumentoCompra.dVencimiento ASC "


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

Private Sub Command2_Click()
strCadena = "SELECT P.id_material,P.fecha,CONCAT(C.doc_abrev,':',P.serie,'-',P.numero) as comprobante,P.id_cliente,CONCAT(TT.descripcion,'-',T.placa) as transporte,anulado,dni_save,P.tiempo_ruta FROM parte_material P,comprobantes C,transporte T,transporte_tipo TT WHERE T.id_tipo_transporte=TT.id_tipo_transporte AND P.id_transporte=T.id_transporte AND P.id_doc=C.id_doc AND P.ruc='" & KEY_RUC & "' AND T.ruc='" & KEY_RUC & "' AND TT.ruc='" & KEY_RUC & "' AND T.placa LIKE '%" & Trim(Me.TxtcodProveedor.Text) & "%' AND P.fecha >= '" & Format(Me.DtpFechaIni.Value, "YYYY-mm-dd") & "' AND P.fecha<='" & Format(Me.DtpFechaFin.Value, "YYYY-mm-dd") & "'  ORDER BY P.fecha ASC  "
Call llenarGrid(Me.HfgFacturas, Me)


End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 200
Me.DtpFechaFin.Value = KEY_FECHA
Me.DtpFechaIni.Value = KEY_FECHA
strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM
Me.DtcAlmacen.Enabled = False
  'Me.Toolbar1.Buttons(KEY_PRINT).Enabled = False
  'Me.Toolbar1.Buttons(KEY_DELETE).Enabled = False
  'Me.TlbGrabar.Buttons(KEY_PAGAR).Enabled = False
  'Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
Call facturas

  
End Sub
Private Sub HfgFacturas_DblClick()
Dim estado As String
End Sub

Private Sub HfgFacturas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) > 0 Then
        Call FrmParteMaterial.nuevo
        strCadena = "SELECT * FROM parte_material WHERE id_material='" & Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        FrmParteMaterial.DtcTipoDoc.BoundText = rst("id_doc")
        FrmParteMaterial.txtSerie.Text = rst("serie")
        FrmParteMaterial.TxtNumeroDoc.Text = rst("numero")
        Call FrmParteMaterial.buscar_comprobante(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0))
        Unload Me
        Exit Sub
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
          FrmCompras.txtSerie.Enabled = True
          FrmCompras.TxtNumeroDoc.Enabled = True
          FrmCompras.DtcTipoDoc.BoundText = cod_doc
          FrmCompras.Txtdoc_cod.Text = cod_doc
          FrmCompras.txtSerie.Text = serie_doc
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

Private Sub TxtcodProveedor_Change()
strCadena = "SELECT P.id_material,P.fecha,CONCAT(C.doc_abrev,':',P.serie,'-',P.numero) as comprobante,P.id_cliente,CONCAT(TT.descripcion,'-',T.placa) as transporte,anulado,dni_save,P.tiempo_ruta FROM parte_material P,comprobantes C,transporte T,transporte_tipo TT WHERE T.id_tipo_transporte=TT.id_tipo_transporte AND P.id_transporte=T.id_transporte AND P.id_doc=C.id_doc AND P.ruc='" & KEY_RUC & "' AND T.ruc='" & KEY_RUC & "' AND TT.ruc='" & KEY_RUC & "' AND T.placa LIKE '%" & Trim(Me.TxtcodProveedor.Text) & "%'   ORDER BY P.fecha ASC  "
Call llenarGrid(Me.HfgFacturas, Me)


End Sub

Private Sub TxtRuc_Change()
strCadena = "SELECT P.id_material,P.fecha,CONCAT(C.doc_abrev,':',P.serie,'-',P.numero) as comprobante,P.id_cliente,CONCAT(TT.descripcion,'-',T.placa) as transporte,anulado,dni_save,P.tiempo_ruta FROM parte_material P,comprobantes C,transporte T,transporte_tipo TT,persona PP WHERE P.id_cliente=PP.dni AND T.id_tipo_transporte=TT.id_tipo_transporte AND P.id_transporte=T.id_transporte AND P.id_doc=C.id_doc AND P.ruc='" & KEY_RUC & "' AND T.ruc='" & KEY_RUC & "' AND TT.ruc='" & KEY_RUC & "' AND  PP.nombre_completo LIKE '%" & Trim(Me.txtruc.Text) & "%'   ORDER BY P.fecha ASC  "
Call llenarGrid(Me.HfgFacturas, Me)



End Sub







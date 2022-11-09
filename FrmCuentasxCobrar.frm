VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmCuentasxCobrar 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   15555
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtNombreRazon 
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
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   2775
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
      Left            =   6720
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   6480
      Top             =   6600
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
            Picture         =   "FrmCuentasxCobrar.frx":0000
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasxCobrar.frx":0454
            Key             =   "(Visualizar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasxCobrar.frx":04E1
            Key             =   "(Pagar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCuentasxCobrar.frx":07FB
            Key             =   "(Eliminar)"
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   9600
      TabIndex        =   2
      Top             =   225
      Width           =   3855
      _ExtentX        =   6800
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
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   120
      TabIndex        =   5
      Top             =   7800
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   1482
      BandCount       =   1
      ForeColor       =   -2147483635
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   2430
      _CBHeight       =   840
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   780
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbGrabar 
         Height          =   810
         Left            =   120
         TabIndex        =   6
         Top             =   15
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   1429
         ButtonWidth     =   1746
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Pagar"
               Key             =   "(Pagar)"
               Object.ToolTipText     =   "Grabar Ctrl+G"
               ImageKey        =   "(Pagar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Visualizar"
               Key             =   "(Imprimir)"
               Object.ToolTipText     =   "Grabar Ctrl+I"
               ImageKey        =   "(Visualizar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   840
      Left            =   10200
      TabIndex        =   7
      Top             =   7800
      Width           =   3270
      _ExtentX        =   5768
      _ExtentY        =   1482
      BandCount       =   1
      ForeColor       =   -2147483635
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   3270
      _CBHeight       =   840
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   780
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   810
         Left            =   120
         TabIndex        =   8
         Top             =   15
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   1429
         ButtonWidth     =   1667
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Grabar Ctrl+G"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Imprimir"
               Key             =   "(Imprimir)"
               Object.ToolTipText     =   "Grabar Ctrl+I"
               ImageKey        =   "(Visualizar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir "
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfcomprobantes 
      Height          =   6975
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   12303
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfCuotasPendientes 
      Height          =   3015
      Left            =   10200
      TabIndex        =   11
      Top             =   1080
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5318
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfPagos 
      Height          =   3255
      Left            =   10200
      TabIndex        =   12
      Top             =   4440
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5741
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
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
   Begin VB.Image Image2 
      Height          =   240
      Left            =   10320
      Picture         =   "FrmCuentasxCobrar.frx":0D95
      Top             =   4160
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAGOS REALIZADOS"
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
      Left            =   10635
      TabIndex        =   14
      Top             =   4185
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   10320
      Picture         =   "FrmCuentasxCobrar.frx":131F
      Top             =   750
      Width           =   240
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COUTAS PENDIENTES DE PAGO"
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
      Left            =   10635
      TabIndex        =   13
      Top             =   795
      Width           =   2280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUCURSAL:"
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
      Left            =   8640
      TabIndex        =   9
      Top             =   270
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE/RAZON SOCIAL:"
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
      Left            =   360
      TabIndex        =   4
      Top             =   300
      Width           =   1875
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/DNI:"
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
      Left            =   5880
      TabIndex        =   3
      Top             =   270
      Width           =   705
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   500
      Left            =   120
      Top             =   120
      Width           =   15255
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   10200
      Top             =   720
      Width           =   5175
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   10200
      Top             =   4120
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   8685
      Left            =   0
      Top             =   0
      Width           =   15555
   End
End
Attribute VB_Name = "FrmCuentasxCobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede



Private Sub Form_Load()
  CenterForm Me
  Me.Top = 50
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  Me.DtcAlmacen.Enabled = False
  Me.Toolbar1.Buttons(KEY_PRINT).Enabled = False
  Me.Toolbar1.Buttons(KEY_DELETE).Enabled = False
  Me.TlbGrabar.Buttons(KEY_PAGAR).Enabled = False
  Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
  Call LlenarComprobantes(Me.hfcomprobantes)
End Sub
Public Sub LlenarComprobantes(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim tTotal As Double
Dim tSaldo As Double
strCadena = "SELECT M.id_venta,M.fecha_emision,M.fecha_vencimiento,C.doc_abrev,M.serie,M.numero,P.nombre_completo,M.total,M.saldo FROM movimiento_venta M,comprobantes C,persona P WHERE M.id_doc=C.id_doc AND M.id_cliente=P.dni AND M.ruc='" & KEY_RUC & "' AND anulado='no' AND M.saldo>0 AND P.nombre_completo LIKE '%" & Trim(Me.TxtNombreRazon.Text) & "%' AND P.dni LIKE '%" & Trim(Me.TxtRuc.Text) & "%'   ORDER BY id_venta ASC "
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
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 1150
           Grilla.ColWidth(2) = 1800
           Grilla.ColWidth(3) = 3500
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1200
           
         Next
        cabecera = "CODIGO" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "CLIENTE" & vbTab & "TOTAL" & vbTab & "SALDO"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        tSaldo = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_venta") & vbTab & rst("fecha_emision") & vbTab & rst("doc_abrev") + ":" + rst("serie") + "-" + rst("numero") & vbTab & UCase(rst("nombre_completo")) & vbTab & Format(rst("total"), "#,##0.00") & vbTab & Format(rst("saldo"), "#,##0.00")
            Grilla.AddItem Fila
            tTotal = tTotal + rst("total")
            tSaldo = tSaldo + rst("saldo")
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(tTotal, "#,##0.00") & vbTab & Format(tSaldo, "#,##0.00")
      Grilla.AddItem Fila
       For k = 4 To 5
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
  
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub
Public Sub LlenarPagos(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT id,fecha,CONCAT(C.doc_abrev,':',M.serie,'-',M.numero) as comprobante,monto,anulado FROM mis_cuentas_det M,comprobantes C WHERE M.id_doc=C.id_doc AND id_movimiento='" & idVenta & "' AND ruc='" & KEY_RUC & "' AND montoreal>=0"
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 500
            Grilla.ColWidth(2) = 1100
            Grilla.ColWidth(3) = 2200
            Grilla.ColWidth(4) = 1200
        Next
        cabecera = "IDCAJA" & vbTab & "ITEM" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "MONTO"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id") & vbTab & formato_item((i + 1), 2) & vbTab & rst("fecha") & vbTab & rst("comprobante") & vbTab & Format(rst("monto"), "#,##0.00")
            Grilla.AddItem Fila
            
            Fila = ""
            If rst("anulado") = "si" Then
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
     
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
       For k = 0 To 4
            Grilla.col = 4
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
    
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub




Private Sub Hfcomprobantes_SelChange()
If Val(Me.hfcomprobantes.TextMatrix(Me.hfcomprobantes.Row, 0)) > 0 Then
    Call LlenarLetras(Me.hfCuotasPendientes, Me.hfcomprobantes.TextMatrix(Me.hfcomprobantes.Row, 0))
    Call LlenarPagos(Me.HfPagos, Me.hfcomprobantes.TextMatrix(Me.hfcomprobantes.Row, 0))
    Me.TlbGrabar.Buttons(KEY_PAGAR).Enabled = True
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
    Me.Toolbar1.Buttons(KEY_PRINT).Enabled = False
    Me.Toolbar1.Buttons(KEY_DELETE).Enabled = False
Else
    Me.TlbGrabar.Buttons(KEY_PAGAR).Enabled = False
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
End If
End Sub
Public Sub LlenarLetras(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT id,id_cuota,vencimiento,monto,saldo FROM movimiento_venta_cuotas WHERE id_venta='" & idVenta & "' AND ruc='" & KEY_RUC & "' ORDER BY id DESC"
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
            Grilla.ColWidth(3) = 1300
            Grilla.ColWidth(4) = 1300
        Next
        cabecera = "IDUNICO" & vbTab & "CUOTA" & vbTab & "VENCIMIENTO" & vbTab & "MONTO" & vbTab & vbTab & "SALDO"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id") & vbTab & rst("id_cuota") & vbTab & rst("vencimiento") & vbTab & Format(rst("monto"), "###0.00") & vbTab & Format(rst("saldo"), "###0.00")
            Grilla.AddItem Fila
            Fila = ""
               If Format(rst("vencimiento"), "YYYY-mm-dd") <= KEY_FECHA Then
                    For k = 0 To 4
                        Grilla.col = k
                        Grilla.Row = i + 1
                        Grilla.CellBackColor = &H8080FF
                    Next k
              End If
            tTotal = tTotal + rst("saldo")
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
       For k = 0 To 4
            Grilla.col = 4
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
    
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Private Sub Hfpagos_SelChange()
If Val(Me.HfPagos.TextMatrix(Me.HfPagos.Row, 0)) > 0 Then
     
     Me.Toolbar1.Buttons(KEY_PRINT).Enabled = True
     Me.Toolbar1.Buttons(KEY_DELETE).Enabled = True
Else
     Me.Toolbar1.Buttons(KEY_PRINT).Enabled = False
      Me.Toolbar1.Buttons(KEY_DELETE).Enabled = False
End If
End Sub

Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.key
    Case KEY_PAGAR
        Procedencia = nuevo
        frmVentasPagos.Show
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error
  Select Case Button.key
    Case KEY_DELETE
          If Val(Me.HfPagos.TextMatrix(Me.HfPagos.Row, 0)) > 0 Then
            If MsgBox("Esta Seguro de Anular este Pago", vbQuestion + vbYesNo, "Mensaje para el Usuario") = vbYes Then
                strCadena = "UPDATE mis_cuentas_det SET anulado='si' WHERE  id='" & Val(Me.HfPagos.TextMatrix(Me.HfPagos.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                 
                
                saldoa = Val(hfcomprobantes.TextMatrix(Me.hfcomprobantes.Row, 6)) + Val(Me.HfPagos.TextMatrix(Me.HfPagos.Row, 4))
                strCadena = "UPDATE movimiento_venta SET saldo='" & saldoa & "' WHERE id_venta='" & Val(Me.hfcomprobantes.TextMatrix(Me.hfcomprobantes.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                 
         
         
         '***** ACTUALIZAR CUOTAS
         saldoa = Val(Me.HfPagos.TextMatrix(Me.HfPagos.Row, 4))
         strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & Val(Me.hfcomprobantes.TextMatrix(Me.hfcomprobantes.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
         Call ConfiguraRst(strCadena)
         If rst.RecordCount > 0 Then
                strCadena = "SELECT * FROM movimiento_venta_cuotas WHERE id_venta='" & Val(Me.hfcomprobantes.TextMatrix(Me.hfcomprobantes.Row, 0)) & "' AND ruc='" & KEY_RUC & "' AND monto<>saldo ORDER BY id DESC"
                Call ConfiguraRstT(strCadena)
                If rstT.RecordCount > 0 Then
                    rstT.MoveFirst
                    For k = 0 To rstT.RecordCount - 1
                      If saldoa > 0 Then
                            If saldoa >= (rstT("monto") - rstT("saldo")) Then
                                strCadena = "UPDATE movimiento_venta_cuotas SET saldo='" & rstT("monto") & "' WHERE id='" & rstT("id") & "' AND ruc='" & KEY_RUC & "'"
                                CnBd.Execute (strCadena)
                                 
                                saldoa = saldoa - (rstT("monto") - rstT("saldo"))
                            Else
                                strCadena = "UPDATE movimiento_venta_cuotas SET saldo='" & Val(rstT("saldo") + saldoa) & "' WHERE id='" & rstT("id") & "' AND ruc='" & KEY_RUC & "'"
                                CnBd.Execute (strCadena)
                                 
                                saldoa = 0
                            End If
                      
                      End If
                      rstT.MoveFirst
                    Next k
                    
                End If
                
         End If
         'FIN ACTUALIZAR CUOTAS
         
         
                
                
                Call LlenarPagos(Me.HfPagos, Me.hfcomprobantes.TextMatrix(Me.hfcomprobantes.Row, 0))
                Call LlenarComprobantes(Me.hfcomprobantes)
                
                Exit Sub
            End If
          End If
    Case KEY_PRINT
         MsgBox "COMUNIQUESE CON EL MAGISTER DE SOFTWARE, PARA ASIGNAR UN FORMATO", vbInformation, KEY_EMPRESA
    Case KEY_EXIT
        Unload Me
     
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub TxtNombreRazon_Change()

    Me.TxtRuc.Text = ""
    Call LlenarComprobantes(Me.hfcomprobantes)

End Sub

Private Sub TxtRuc_Change()
If Trim(Me.TxtRuc.Text) <> "" Then
    Me.TxtNombreRazon.Text = ""
    Call LlenarComprobantes(Me.hfcomprobantes)
End If
End Sub

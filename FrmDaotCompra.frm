VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDaotCompra 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdConsultar 
      Caption         =   "CONSULTAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox ChkAlmacen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ALMACEN :"
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
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   8760
      TabIndex        =   7
      Top             =   240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   225
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
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
      Format          =   62390273
      CurrentDate     =   41250
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   6495
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   11456
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
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   12360
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDaotCompra.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDaotCompra.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDaotCompra.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDaotCompra.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDaotCompra.frx":101C
            Key             =   "(RCompras)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDaotCompra.frx":278E
            Key             =   "(RVentas)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   6465
      Left            =   14040
      TabIndex        =   1
      Top             =   840
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   11404
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   6465
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   2
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1482
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Exportar"
               Key             =   "(Exportar)"
               ImageKey        =   "(RCompras)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Detalle"
               Key             =   "(Detalle)"
               ImageKey        =   "(RVentas)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComCtl2.DTPicker DtpFinal 
      Height          =   315
      Left            =   3960
      TabIndex        =   6
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
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
      Format          =   62390273
      CurrentDate     =   41250
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPRA"
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
      Height          =   195
      Left            =   14160
      TabIndex        =   11
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DAOT"
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
      Height          =   195
      Left            =   14280
      TabIndex        =   10
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HASTA:"
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
      Height          =   195
      Left            =   3120
      TabIndex        =   5
      Top             =   255
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESDE:"
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
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   570
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   7470
      Left            =   0
      Top             =   0
      Width           =   15120
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   240
      Top             =   60
      Width           =   13575
   End
End
Attribute VB_Name = "FrmDaotCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub chkAlmacen_Click()
If Me.ChkAlmacen.Value = 1 Then
    Me.DtcAlmacen.Enabled = True
Else
    Me.DtcAlmacen.Enabled = False
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub CmdConsultar_Click()
Dim StrAlmacen As String
StrAlmacen = ""
If Me.ChkAlmacen.Value = 1 Then
    StrAlmacen = Replace(Me.DtcAlmacen.BoundText, "'", "''")
End If

strCadena = "SELECT C.id_proveedor,P.nombre_completo,SUM(exonerado) as exonerado,SUM(valor_venta) as valorventa,SUM(igv) as igv,SUM(total)as total FROM movimiento_compra C,persona P WHERE C.fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' AND C.fecha_emision<='" & Format(Me.DtpFinal.Value, "YYYY-mm-dd") & "' AND C.id_proveedor=P.dni AND C.ruc='" & KEY_RUC & "' AND C.anulado='no' AND C.id_alm LIKE '%" & Trim(StrAlmacen) & "%' AND (C.id_doc='0003' OR C.id_doc='0001') GROUP BY C.id_proveedor HAVING SUM(total)>0 ORDER BY SUM(total) DESC"
Call llenarGrid(Me.HfdDetalle)

End Sub
Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo SALIR
Dim tTotal As Double, tTexonerado As Double, tTvalorventa As Double, tTigv As Double
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Me.TlbAcciones.Buttons(KEY_EXPORTAR).Enabled = False
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
   Me.TlbAcciones.Buttons(KEY_EXPORTAR).Enabled = True
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 900
           Grilla.ColWidth(1) = 1500
           Grilla.ColWidth(2) = 4500
           Grilla.ColWidth(3) = 1500
           Grilla.ColWidth(4) = 1500
           Grilla.ColWidth(5) = 1500
           Grilla.ColWidth(6) = 1500
        Next
        cabecera = "Nº ORDEN" & vbTab & "RUC" & vbTab & "RAZON SOCIAL " & vbTab & "EXONERADO" & vbTab & "VALOR VENTA" & vbTab & "IGV" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = formato_item(i + 1, 5) & vbTab & rst("id_proveedor") & vbTab & UCase(rst("nombre_completo")) & vbTab & Format(rst("exonerado"), "#,##0.00") & vbTab & Format(rst("valorventa"), "#,##0.00") & vbTab & Format(rst("igv"), "#,##0.00") & vbTab & Format(rst("total"), "#,##0.00")
            Grilla.AddItem Fila
            tTotal = tTotal + rst("total")
            tTexonerado = rst("exonerado") + tTexonerado
            tTvalorventa = tTvalorventa + rst("valorventa")
            tTigv = tTigv + rst("igv")
            rst.MoveNext
    Next i
    Fila = "" & vbTab & "" & vbTab & "************** ACUMULADO PERIODO **************" & vbTab & Format(tTexonerado, "#,##0.00") & vbTab & Format(tTvalorventa, "#,##0.00") & vbTab & Format(tTigv, "#,##0.00") & vbTab & Format(tTotal, "#,##0.00")
    Grilla.AddItem Fila
     For k = 0 To 6
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
    Next k
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1

Exit Sub
SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFinal.Value = KEY_FECHA
strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM
Me.DtcAlmacen.Enabled = False
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_EXPORTAR
          strCadena = "SELECT C.id_proveedor,P.nombre_completo,SUM(exonerado) as exonerado,SUM(valor_venta),SUM(igv),SUM(total) FROM movimiento_compra C,persona P WHERE C.fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' AND C.fecha_emision<='" & Format(Me.DtpFinal.Value, "YYYY-mm-dd") & "' AND C.id_proveedor=P.dni AND C.ruc='" & KEY_RUC & "' AND C.anulado='no' AND C.id_alm LIKE '%" & Trim(StrAlmacen) & "%' AND (C.id_doc='0003' OR C.id_doc='0001') GROUP BY C.id_proveedor HAVING SUM(total)>0 ORDER BY SUM(valor_venta) DESC"
          Call ConfiguraRst(strCadena)
          Ans = ShowMultiReport(rst, "RptDaotCompraSum", , App.Path + "\Reportes\")

    Case "(Detalle)"
           strCadena = "SELECT C.id_proveedor,P.nombre_completo,C.fecha_emision,CONCAT(D.doc_abrev,':',C.serie,'-',C.numero) as comprobante,C.exonerado,C.valor_venta,C.igv,C.total FROM movimiento_compra C,persona P,comprobantes D WHERE C.id_doc=D.id_doc AND  C.fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' AND C.fecha_emision<='" & Format(Me.DtpFinal.Value, "YYYY-mm-dd") & "' AND C.id_proveedor=P.dni AND C.ruc='" & KEY_RUC & "' AND C.anulado='no' AND C.id_alm LIKE '%" & Trim(StrAlmacen) & "%' AND (C.id_doc='0003' OR C.id_doc='0001') ORDER BY C.id_proveedor"
           Call ConfiguraRst(strCadena)
           Ans = ShowMultiReport(rst, "RptDaotCompra", , App.Path + "\Reportes\")
    Case KEY_EXIT
        Unload Me
End Select
End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmProducto_merma_lista 
   BorderStyle     =   0  'None
   Caption         =   "Detalle Mermas"
   ClientHeight    =   7860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   14265
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Todo"
      Height          =   300
      Left            =   11400
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton CmdBuscarMotivo 
      Caption         =   "Buscar"
      Height          =   300
      Left            =   10560
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   12840
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      Format          =   66387969
      CurrentDate     =   40615
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      Format          =   66387969
      CurrentDate     =   40615
   End
   Begin MSDataListLib.DataCombo DtcMerma 
      Height          =   315
      Left            =   5640
      TabIndex        =   4
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   6855
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   12091
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   14055
   End
End
Attribute VB_Name = "FrmProducto_merma_lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdBuscar_Click()
Strcadena = "SELECT * FROM merma M,persona P,comprobantes C WHERE M.id_doc=C.id_doc AND  M.id_usuario=P.dni AND M.ruc='" & KEY_RUC & "' AND M.fecha>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' AND M.fecha<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "'"
Call llenar_grid(Me.HfdDetalle)

End Sub
Public Sub llenar_grid(ByVal Grilla As MSHFlexGrid)
On Error GoTo SALIR
Dim tTotal As Double
Call ConfiguraRst(Strcadena)
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
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 5000
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 4000
           
         Next
        cabecera = "CODIGO" & vbTab & "COMPROBANTE" & vbTab & "FECHA" & vbTab & "DETALLE" & vbTab & "MONTO" & vbTab & "RESPONSABLE"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_merma") & vbTab & rst("doc_abrev") & ":" & rst("serie") & "-" & rst("numero") & vbTab & rst("fecha") & vbTab & rst("detalle") & vbTab & Format(rst("costo"), "#,##0.00") & vbTab & rst("nombre_completo")
            Grilla.AddItem Fila
            tTotal = tTotal + rst("costo")
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "" & "" & vbTab & " ACUMULADO TOTAL " & vbTab & Format(tTotal, "#,##0.00") & vbTab & ""
      Grilla.AddItem Fila
       For k = 3 To 5
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


Private Sub CmdBuscarMotivo_Click()
Strcadena = "SELECT DISTINCT D.id_merma,serie,numero,detalle,M.costo,C.doc_abrev,P.nombre_completo,M.fecha FROM merma M,persona P,comprobantes C,merma_detalle D WHERE M.id_merma=D.id_merma AND M.id_doc=C.id_doc AND  M.id_usuario=P.dni AND M.ruc='" & KEY_RUC & "' AND M.fecha>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' AND M.fecha<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' AND D.id_motivo='" & Me.DtcMerma.BoundText & "' AND M.id_alm='" & KEY_ALM & "'"
Call llenar_grid(Me.HfdDetalle)

End Sub

Private Sub cmSalir_Click()
Unload Me
End Sub



Private Sub Command1_Click()
Call actualizar
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
Me.DtpFin.Value = Date
Me.DtpInicio.Value = Date
 Strcadena = "SELECT id_merma as Codigo, descripcion as Descripcion FROM merma_motivo  ORDER BY descripcion"
  Call ConfiguraRst(Strcadena)
  Call LlenaDataCombo(Me.DtcMerma)
  Set rst = Nothing
Call actualizar
  
End Sub
Public Sub actualizar()
Strcadena = "SELECT * FROM merma M,persona P,comprobantes C WHERE M.id_doc=C.id_doc AND  M.id_usuario=P.dni AND M.ruc='" & KEY_RUC & "'"
Call llenar_grid(Me.HfdDetalle)

End Sub





Private Sub HfdDetalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Strcadena = "SELECT * FROM merma WHERE id_merma='" & Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(Strcadena)
    If rst.RecordCount > 0 Then
        FrmProductoMermas.consulta (rst("numero"))
        Unload Me
    End If
End If
End Sub

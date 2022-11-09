VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmVentasCuotas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CUOTAS FECHA PAGO CREDITO"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn Command2 
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   3840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "CERRAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "FrmVentasCuotas.frx":0000
      PICN            =   "FrmVentasCuotas.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame FrmPanel 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   200
         Width           =   495
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         CalendarTitleBackColor=   33023
         Format          =   166985729
         CurrentDate     =   41078
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FECHA:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   200
         Width           =   615
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6165
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
End
Attribute VB_Name = "FrmVentasCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
Dim strfecha As String
strfecha = Format(Me.DTPicker1.Value, "YYYY-mm-dd")
strCadena = "UPDATE movimiento_venta_cuotas_temporal SET vencimiento='" & strfecha & "' WHERE id='" & Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0) & "'"
CnBd.Execute (strCadena)
 
Call llenarGrid_det(Me.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText)
Me.FrmPanel.Visible = False
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Me.Command2.SetFocus
End Sub

Private Sub Form_Load()
Me.Top = FrmVentas.TxtMontoPagado.Top + FrmVentas.TxtMontoPagado.Height
Me.Left = FrmVentas.TxtMontoPagado.Left + FrmVentas.TxtMontoPagado.Width
Call llenarGrid_det(Me.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText)
End Sub
Public Sub llenarGrid_det(ByVal Grilla As MSHFlexGrid, ByVal id_numero As String, ByVal id_serie As String, ByVal id_doc As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT id,id_cuota,vencimiento,monto,saldo FROM movimiento_venta_cuotas_temporal T  WHERE  T.id_doc='" & id_doc & "' AND serie='" & id_serie & "' AND numero='" & id_numero & "' AND id_usuario='" & KEY_USUARIO & "' ANd ruc='" & KEY_RUC & "' ORDER BY id ASC "
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
           Grilla.ColWidth(1) = 900
           Grilla.ColWidth(2) = 1300
           Grilla.ColWidth(3) = 1300
           Grilla.ColWidth(4) = 1300
           
        Next
        cabecera = "IDCUOTA" & vbTab & "N-CUOTA" & vbTab & "VENCIMIENTO " & vbTab & "MONTO " & vbTab & "SALDO"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id") & vbTab & rst("id_cuota") & vbTab & Format(rst("vencimiento"), "dd-mm-YYYY") & vbTab & Format(rst("monto"), "###0.0000") & vbTab & Format(rst("saldo"), "#,##0.0000")
            Grilla.AddItem Fila
            tTotal = tTotal + rst("saldo")
            rst.MoveNext
    Next i
    Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(tTotal, "###0.00")
    Grilla.AddItem Fila
     For k = 0 To 4
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

Public Sub llenarGrid_cuotas(ByVal in_venta As String, ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,documento,total,saldo FROM movimiento_venta  WHERE id_doc='0412' and  id_referencia='" & Val(in_venta) & "' and  ruc='" & KEY_RUC & "' ORDER BY id_venta ASC "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 1300
           Grilla.ColWidth(4) = 1300
           
        Next
        cabecera = "IDCUOTA" & vbTab & "FECHA" & vbTab & "LETRA " & vbTab & "TOTAL " & vbTab & "SALDO"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_venta") & vbTab & Format(rst("fecha_vencimiento"), "dd-mm-YYYY") & vbTab & rst("documento") & vbTab & Format(rst("total"), "###0.00") & vbTab & Format(rst("saldo"), "#,##0.00")
            Grilla.AddItem Fila
            tTotal = tTotal + rst("saldo")
            rst.MoveNext
    Next i
    Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(tTotal, "###0.00")
    Grilla.AddItem Fila
     For k = 0 To 4
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


Private Sub HfdDetalle_DblClick()
If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    Me.FrmPanel.Visible = True
    Me.DTPicker1.Value = CVDate(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 2))
    Me.cmdOk.SetFocus
End If
End Sub

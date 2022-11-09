VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmBuscardoc 
   BorderStyle     =   0  'None
   Caption         =   "Buscar"
   ClientHeight    =   6570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfGuardado 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   11535
      _ExtentX        =   20346
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDetalle 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   5106
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
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOCUMENTOS GUARDADOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Height          =   6570
      Left            =   0
      Top             =   0
      Width           =   11835
   End
End
Attribute VB_Name = "frmBuscardoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()

  Me.HfGuardado.SetFocus
End Sub


Sub llenarGridME(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.Rows = rst.RecordCount
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 1000
  Grilla.ColWidth(1) = 4000
  Grilla.ColWidth(2) = 0
  Grilla.ColWidth(3) = 3000
  
'Call DarFormato1(Grilla, 0)
Grilla.Refresh
 
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Sub llenarGridDetalle(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.Rows = rst.RecordCount
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 700
  Grilla.ColWidth(1) = 4800
  Grilla.ColWidth(2) = 700
  'Grilla.ColWidth(3) = 3000
'Call DarFormato1(Grilla, 0)
Grilla.Refresh
 Me.HfGuardado.SetFocus
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 27) Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
Call llenarGridGuardado(Me.HfGuardado)
End Sub
Private Sub llenarGridGuardado(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT * FROM temporal_venta_guardado WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
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
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1500
           Grilla.ColWidth(3) = 4500
           Grilla.ColWidth(4) = 1200
          Next
         cabecera = "CODIGO" & vbTab & "EMISION" & vbTab & "HORA" & vbTab & "CLIENTE" & vbTab & "MONTO"
         Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = rst("id_codigo") & vbTab & rst("fecha") & vbTab & rst("hora") & vbTab & rst("cliente") & vbTab & Format(rst("monto_guardado"), "#,##0.00")
        Grilla.AddItem Fila
        Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Private Sub HfGuardado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And FrmVentas.Procedencia = buscar Then
    Me.HfGuardado.col = 0
    strCadena = "UPDATE temporal_ventas SET save='no',id_medic='0',id_serie='" & FrmVentas.DtcSerieDoc.BoundText & "',id_doc='" & FrmVentas.DtcTipoDoc.BoundText & "',numero='" & Trim(FrmVentas.TxtNumeroDoc.Text) & "' WHERE id_medic='" & Val(Me.HfGuardado.TextMatrix(Me.HfGuardado.Row, 0)) & "' AND dni_save='" & Trim(KEY_USUARIO) & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    strCadena = "DELETE FROM temporal_venta_guardado WHERE id_codigo='" & Val(Me.HfGuardado.TextMatrix(Me.HfGuardado.Row, 0)) & "' and ruc='" & Trim(KEY_RUC) & "'"
    CnBd.Execute (strCadena)
     
    
     FrmVentas.TxtCodCliente.Text = "00000000"
     FrmVentas.txtcliente.Text = Me.HfGuardado.TextMatrix(Me.HfGuardado.Row, 3)
     FrmVentas.txtdireccion.Text = KEY_DIR_PUBLIC
     
    Call FrmVentas.llenarGrid_det(FrmVentas.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText, Trim(FrmVentas.txtformato_impresion.Text))
    'FrmVentas.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
    FrmVentas.cmdprocesar.Enabled = True
    Unload Me
    Exit Sub
End If
End Sub



Private Sub llenarGridDetalleGuardado(ByVal Grilla As MSHFlexGrid, ByVal id_codigo As Double)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT * FROM temporal_ventas T,producto P,unidad U WHERE T.id_producto=P.id_producto AND T.id_medic='" & id_codigo & "' AND T.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "'"
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
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 3500
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1000
           Grilla.ColWidth(5) = 1200
          Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "UNIDAD" & vbTab & "CANTIDAD" & vbTab & "PRECIO" & vbTab & "TOTAL"
         Grilla.AddItem cabecera
         For k = 0 To 5
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & rst("cantidad") & vbTab & Format(rst("precio"), "#,##0.00") & vbTab & Format(rst("total"), "#,##0.00")
        Grilla.AddItem Fila
        Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub HfGuardado_SelChange()
If Val(Me.HfGuardado.TextMatrix(Me.HfGuardado.Row, 0)) > 0 Then
    Call llenarGridDetalleGuardado(Me.HfDetalle, Val(Me.HfGuardado.TextMatrix(Me.HfGuardado.Row, 0)))
End If
End Sub

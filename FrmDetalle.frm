VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmDetalle 
   BorderStyle     =   0  'None
   Caption         =   "Detalle venta"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   60
   ClientWidth     =   13095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   13095
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdhash 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      TabIndex        =   13
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox txtHash 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      TabIndex        =   12
      Top             =   5280
      Width           =   6375
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrarPantalla 
      Height          =   375
      Left            =   12600
      TabIndex        =   9
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      MICON           =   "FrmDetalle.frx":0000
      PICN            =   "FrmDetalle.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   2700
      Left            =   420
      TabIndex        =   0
      Top             =   2475
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   4763
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   12582912
      BackColorSel    =   16777215
      ForeColorSel    =   8388608
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfpago 
      Height          =   1140
      Left            =   8100
      TabIndex        =   10
      Top             =   5400
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   2011
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   0
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblempresa 
      BackStyle       =   0  'Transparent
      Caption         =   "PERCY ANTICONA "
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   600
      TabIndex        =   11
      Top             =   120
      Width           =   11655
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10800
      TabIndex        =   8
      Top             =   5520
      Width           =   1365
   End
   Begin VB.Label lblPago 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10800
      TabIndex        =   7
      Top             =   5880
      Width           =   1365
   End
   Begin VB.Label lblVuelto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   10800
      TabIndex        =   6
      Top             =   6240
      Width           =   1365
   End
   Begin VB.Label lblnumero 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   8520
      TabIndex        =   5
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label lblruc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   8520
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lbldni 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   150
   End
   Begin VB.Label lblnombre 
      BackColor       =   &H00FFFFFF&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   2040
      Width           =   5415
   End
   Begin VB.Label lblfecha 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   8640
      TabIndex        =   1
      Top             =   2040
      Width           =   150
   End
   Begin VB.Image ImgCancelado 
      Height          =   1095
      Left            =   2640
      Picture         =   "FrmDetalle.frx":2ED0
      Top             =   5520
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.Image imganulado 
      Height          =   1170
      Left            =   2640
      Picture         =   "FrmDetalle.frx":DD09
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   4275
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   6735
      Left            =   0
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "FrmDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub cmdCerrar_Click()

Call cerrar

End Sub

Public Sub cerrar()




If FrmReporteRecaudacionDiaria.Procedencia = buscar Then
    Call enabled_form(FrmReporteRecaudacionDiaria)
    Unload Me
    FrmReporteRecaudacionDiaria.Procedencia = Neutro
    Exit Sub
End If




If FrmVentas.Procedencia = buscar Then
   Call enabled_form(FrmVentas)
   Unload Me
   FrmVentas.Procedencia = Neutro
   Exit Sub
End If



End Sub



Private Sub cmdCerrarpantalla_Click()
Call cerrar
End Sub

Private Sub cmdhash_Click()

strCadena = "UPDATE movimiento_venta SET sunat_key='" & Trim(Me.txtHash.Text) & "',sunat_hash='" & Trim(Me.txtHash.Text) & "' WHERE id_venta='" & Val(FrmVentas.HfFacturas.TextMatrix(FrmVentas.HfFacturas.Row, 0)) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
CnBd.Execute (strCadena)
MsgBox "Listo"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Call cerrar
End If
End Sub
Public Sub llena_pagosVenta(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo salir
Dim tpago As Double
strCadena = "SELECT * FROM movimiento_venta_monto M,forma_pago_detalle F WHERE M.id_forma_pago=F.id_registro AND id_venta='" & idVenta & "' AND M.ruc='" & KEY_RUC & "' AND F.ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
       
    Exit Sub
    
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2000
           Grilla.ColWidth(2) = 1500
       Next
        cabecera = "CODIGO" & vbTab & "FORMA PAGO" & vbTab & "MONTO"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tpago = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_detalle") & vbTab & rst("descripcion") & vbTab & Format(rst("monto"), "###0.00")
            Grilla.AddItem Fila
            tpago = rst("monto") + tpago
            rst.MoveNext
    Next i
    Dim tventa As Double
    tventa = Val(Format(Me.lblTotal.Caption, "###0.000"))
    Me.lblTotal.Caption = Format(tventa, "###0.000")
    Me.lblPago.Caption = Format(tpago, "###0.000")
    Me.lblVuelto.Caption = Format(tpago - tventa, "#,##0.000")
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 1500
Me.lblempresa.Caption = KEY_EMPRESA
Dim id_vent As Double
Me.lblruc.Caption = "RUC :" & KEY_RUC
If FrmReporteRecaudacionDiaria.Procedencia = buscar Then
    strCadena = "SELECT total, monto_pago, monto_vuelto,M.ncliente as nombre_completo,M.fecha_emision,M.id_cliente,anulado,M.documento,sunat_key FROM movimiento_venta M  WHERE  id_venta='" & FrmReporteRecaudacionDiaria.HfgEfectivo.TextMatrix(FrmReporteRecaudacionDiaria.HfgEfectivo.Row, 0) & "' AND ruc='" & KEY_RUC & "' "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.lbldni.Caption = "DNI :" & rst("id_cliente")
        Me.lblfecha.Caption = Format(rst("fecha_emision"), "dd-mm-YYYY")
        Me.lblnombre.Caption = UCase(rst("nombre_completo"))
        Me.txtHash.Text = rst("sunat_key")
        If rst("anulado") = "si" Then
            Me.imganulado.Visible = True
            Me.ImgCancelado.Visible = False
        Else
            Me.imganulado.Visible = False
            Me.ImgCancelado.Visible = True
        End If
    
    Me.lblnumero.Caption = rst("documento")
    Me.lblTotal.Caption = Format(rst(0), "#,##0.00")
    Me.lblPago.Caption = Format(rst(1), "#,##0.00")
    Me.lblVuelto.Caption = Format(rst(2), "#,##0.00")
    id_vent = FrmReporteRecaudacionDiaria.HfgEfectivo.TextMatrix(FrmReporteRecaudacionDiaria.HfgEfectivo.Row, 0)
    strCadena = "SELECT D.id_producto as codigo,D.detalle as producto,U.abreviatura AS unidad,D.cantidad AS cantidad,D.precio AS precio,D.total AS total FROM movimiento_venta_detalle D,producto P,unidad U WHERE D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND D.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND D.id_venta='" & id_vent & "'"
    Call llenarGridDetalle(Me.HfgDetalle)
    Call llena_pagosVenta(Me.hfpago, Val(FrmReporteRecaudacionDiaria.HfgEfectivo.TextMatrix(FrmReporteRecaudacionDiaria.HfgEfectivo.Row, 0)))
   
    Exit Sub
Else
   
    Exit Sub
End If
End If




If FrmVentas.Procedencia = buscar Then
    strCadena = "SELECT total, monto_pago, monto_vuelto,P.nombre_completo,M.fecha_emision,M.id_cliente,anulado,sunat_key FROM movimiento_venta M,persona P WHERE M.id_cliente=P.dni AND id_venta='" & FrmVentas.HfFacturas.TextMatrix(FrmVentas.HfFacturas.Row, 0) & "' AND ruc='" & KEY_RUC & "' "
    Call ConfiguraRst(strCadena)
       Me.lbldni.Caption = "DNI           :" & rst("id_cliente")
       Me.lblfecha.Caption = Format(rst("fecha_emision"), "dd-mm-YYYY")
    Me.lblnombre.Caption = "PACIENTE :" & UCase(rst("nombre_completo"))
     Me.txtHash.Text = rst("sunat_key")
    If rst("anulado") = "si" Then
        Me.imganulado.Visible = True
        Me.ImgCancelado.Visible = False
    Else
        Me.imganulado.Visible = False
        Me.ImgCancelado.Visible = True
    End If
    
    id_vent = FrmVentas.HfFacturas.TextMatrix(FrmVentas.HfFacturas.Row, 0)
    
    Me.lblnumero.Caption = "[" & str(id_vent) & "]" & FrmVentas.HfFacturas.TextMatrix(FrmVentas.HfFacturas.Row, 2)
    Me.lblTotal.Caption = Format(rst(0), "#,##0.00")
    Me.lblPago.Caption = Format(rst(1), "#,##0.00")
    Me.lblVuelto.Caption = Format(rst(2), "#,##0.00")
    
    
    
End If




strCadena = "SELECT D.id_producto as codigo,D.detalle as producto,U.abreviatura AS unidad,D.cantidad AS cantidad,D.precio AS precio,D.total AS total FROM movimiento_venta_detalle D,producto P,unidad U WHERE D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND D.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND D.id_venta='" & id_vent & "'"
Call llenarGridDetalle(Me.HfgDetalle)
Call llena_pagosVenta(Me.hfpago, Val(FrmVentas.HfFacturas.TextMatrix(FrmVentas.HfFacturas.Row, 0)))

End Sub

Public Sub llenarGridDetalle(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
        Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 7500
           Grilla.ColWidth(3) = 1100
           Grilla.ColWidth(4) = 1100
           
           
       Next
        
        rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
            Fila = rst("cantidad") & vbTab & rst("unidad") & vbTab & rst("codigo") & Space(4) & rst("producto") & vbTab & Format(rst("precio"), "#,##00.00") & vbTab & Format(rst("precio") * rst("cantidad"), "#,##0.00")
            Grilla.AddItem Fila
            rst.MoveNext
    Next i
                Grilla.ColAlignment(1) = 0
                Grilla.ColAlignment(2) = 0
                Grilla.ColAlignment(3) = 0
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub




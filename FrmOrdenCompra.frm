VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmOrdenCompra 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   18645
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn cmdbuscar 
      Height          =   340
      Left            =   12960
      TabIndex        =   12
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "BUSCAR"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmOrdenCompra.frx":0000
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
      Height          =   340
      Left            =   9480
      TabIndex        =   10
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
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
      Format          =   55836673
      CurrentDate     =   43141
   End
   Begin VB.TextBox TxtProveedor 
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   4920
      TabIndex        =   9
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox txtNumero 
      Appearance      =   0  'Flat
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
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   600
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   1230
      Width           =   17295
      _ExtentX        =   30506
      _ExtentY        =   13785
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
   Begin VitekeySoft.ChameleonBtn cmdAnularOrden 
      Height          =   1020
      Left            =   17520
      TabIndex        =   1
      Top             =   3330
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1799
      BTYPE           =   5
      TX              =   "ANULAR "
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
      MICON           =   "FrmOrdenCompra.frx":001C
      PICN            =   "FrmOrdenCompra.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   1020
      Left            =   17520
      TabIndex        =   2
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1799
      BTYPE           =   5
      TX              =   "NUEVO"
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
      MICON           =   "FrmOrdenCompra.frx":2482
      PICN            =   "FrmOrdenCompra.frx":249E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdEditable 
      Height          =   1020
      Left            =   17520
      TabIndex        =   3
      Top             =   2265
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1799
      BTYPE           =   5
      TX              =   "VISUALIZAR"
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
      MICON           =   "FrmOrdenCompra.frx":28F0
      PICN            =   "FrmOrdenCompra.frx":290C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrarpantalla 
      Height          =   1020
      Left            =   17520
      TabIndex        =   4
      Top             =   4395
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1799
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
      MICON           =   "FrmOrdenCompra.frx":2C26
      PICN            =   "FrmOrdenCompra.frx":2C42
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   345
      Left            =   11160
      TabIndex        =   11
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
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
      Format          =   55836673
      CurrentDate     =   43141
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR :"
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
      Height          =   195
      Left            =   3960
      TabIndex        =   8
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NUMERO :"
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
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORDEN COMPRA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   50
      Width           =   1665
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   17295
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   18645
   End
End
Attribute VB_Name = "FrmOrdenCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Public Sub actualizar()

strCadena = "SELECT * FROM view_orden_compra WHERE ruc='" & KEY_RUC & "' LIMIT 29 "
Call llenarGrid(Me.HfgDetalle, Me)



End Sub

Private Sub HfgMarcas_Click()
If HfgMarcas.Row > 0 Then
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    TlbAcciones.Buttons(KEY_DELETE).Enabled = True
  End If
End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub cmdAnularOrden_Click()
Procedencia = anular
Call disabled_form(Me)
frmsegurity.Show
Exit Sub
End Sub

Private Sub cmdBuscar_Click()
strCadena = "SELECT * FROM view_orden_compra WHERE fecha_solicitud>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_solicitud<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and   ruc='" & KEY_RUC & "' "
Call llenarGrid(Me.HfgDetalle, Me)
End Sub

Private Sub cmdCerrarpantalla_Click()
Unload Me
End Sub

Private Sub cmdEditable_Click()
       Procedencia = Selecionar
       FrmOrdenCompraDet.Show
       Call FrmOrdenCompraDet.get_orden(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0))
       FrmOrdenCompraDet.DtcComrpobante.Enabled = False
End Sub

Private Sub cmdEliminar_Click()

End Sub

Private Sub cmdNuevo_Click()
Procedencia = nuevo
FrmOrdenCompraDet.Show
Call FrmOrdenCompraDet.nuevo_registro
   
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA



Call actualizar
End Sub

Private Sub HfgDetalle_SelChange()
If Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) > 0 Then
     'TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    'TlbAcciones.Buttons(KEY_DELETE).Enabled = True
Else
     'TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
    'TlbAcciones.Buttons(KEY_DELETE).Enabled = False
End If
End Sub


Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Dim tTotal As Double, tSaldo As Double, nsaldo As Double
tTotal = 0
tSaldo = 0
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub

End If
  
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1900
           Grilla.ColWidth(2) = 1100
           Grilla.ColWidth(3) = 1100
           Grilla.ColWidth(4) = 1600
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 3200
           Grilla.ColWidth(7) = 1200
           Grilla.ColWidth(8) = 1000
           Grilla.ColWidth(9) = 700
           Grilla.ColWidth(10) = 2200
           Grilla.ColWidth(11) = 1800
         
           
      Next
        cabecera = "ORDEN" & vbTab & "COMPROBANTE" & vbTab & "F. EMISION" & vbTab & "F.PAGO" & vbTab & "RECEPCION " & vbTab & "RUC " & vbTab & "PROVEEDOR " & vbTab & "MONEDA " & vbTab & "TOTAL " & vbTab & "TC " & vbTab & "ESTADO " & vbTab & "RESPONSABLE "
        Grilla.AddItem cabecera
         For k = 0 To 11
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If rst("id_moneda") = "00001" Then
                in_moneda = "SOLES"
            Else
                in_moneda = "DOLARES"
            End If
            
            If rst("id_doc") = "0110" Then
               in_doc = "O-COMPRA :"
            Else
               in_doc = "RECEPCION :"
            End If
            
            Fila = Format(rst("id_orden"), "000000") & vbTab & rst("comprobante") & vbTab & Format(rst("fecha_solicitud"), "dd-mm-YYYY") & vbTab & Format(rst("fecha_pago"), "dd-mm-YYYY") & vbTab & rst("recepcion") & vbTab & rst("id_proveedor") & vbTab & rst("nombre_completo") & vbTab & in_moneda & vbTab & Format(rst("total"), "#,##0.00") & vbTab & rst("tc") & vbTab & rst("estado") & vbTab & Mid(rst("operador"), 1, 20)
            Grilla.AddItem Fila
            If rst("id_estado") = 1 Then
                For k = 8 To 10
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                Next k
            End If
           If rst("id_estado") = 2 Then
                For k = 8 To 10
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H80FF&
                Next k
            End If
            rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub







Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_orden_compra WHERE comprobante LIKE '%" & Trim(Me.txtNumero.Text) & "%' and   ruc='" & KEY_RUC & "' "
    Call llenarGrid(Me.HfgDetalle, Me)

End If
End Sub

VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmPedido 
   BorderStyle     =   0  'None
   Caption         =   "Pedidos"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   -30
   ClientWidth     =   16920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   16920
   ShowInTaskbar   =   0   'False
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
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VitekeySoft.ChameleonBtn cmdbuscar 
      Height          =   345
      Left            =   13920
      TabIndex        =   0
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
      MICON           =   "FrmPedido.frx":0000
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
      Height          =   345
      Left            =   10560
      TabIndex        =   1
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
      Format          =   160759809
      CurrentDate     =   43141
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   7815
      Left            =   120
      TabIndex        =   3
      Top             =   1230
      Width           =   15495
      _ExtentX        =   27331
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
      Left            =   15720
      TabIndex        =   4
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
      MICON           =   "FrmPedido.frx":001C
      PICN            =   "FrmPedido.frx":0038
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
      Left            =   15720
      TabIndex        =   5
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
      MICON           =   "FrmPedido.frx":2482
      PICN            =   "FrmPedido.frx":249E
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
      Left            =   15720
      TabIndex        =   6
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
      MICON           =   "FrmPedido.frx":28F0
      PICN            =   "FrmPedido.frx":290C
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
      Left            =   15720
      TabIndex        =   7
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
      MICON           =   "FrmPedido.frx":2C26
      PICN            =   "FrmPedido.frx":2C42
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
      Left            =   12240
      TabIndex        =   8
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
      Format          =   160759809
      CurrentDate     =   43141
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORDEN DE PEDIDOS"
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
      TabIndex        =   10
      Top             =   45
      Width           =   1980
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
      TabIndex        =   9
      Top             =   600
      Width           =   690
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   15495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   16920
   End
End
Attribute VB_Name = "FrmPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Public Sub actualizar()

strCadena = "SELECT * FROM view_pedido WHERE ruc='" & KEY_RUC & "' LIMIT 29 "
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

Private Sub cmdbuscar_Click()
strCadena = "SELECT * FROM view_pedido WHERE fecha>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and   ruc='" & KEY_RUC & "' "
Call llenarGrid(Me.HfgDetalle, Me)





End Sub

Private Sub cmdCerrarpantalla_Click()
Unload Me
End Sub

Private Sub cmdEditable_Click()
       Procedencia = Selecionar
       FrmDetallePedido.Show
       Call FrmDetallePedido.get_orden(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0))
       FrmDetallePedido.DtcComrpobante.Enabled = False
End Sub

Private Sub cmdeliminar_Click()

End Sub

Private Sub cmdNuevo_Click()
Call nuevo_pedido
Procedencia = nuevo
FrmDetallePedido.Show

   
End Sub

Private Sub nuevo_pedido()
    strCadena = "DELETE FROM movimiento_pedido_detalle_temp WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
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

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub

End If
  
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 1500
           Grilla.ColWidth(3) = 5000
           Grilla.ColWidth(4) = 3000
           Grilla.ColWidth(5) = 2400
           
         
           
      Next
        cabecera = "N° PEDIDO" & vbTab & "COMPROBANTE" & vbTab & "F. EMISION" & vbTab & "SUCURSAL" & vbTab & "OPERADOR " & vbTab & "ESTADO "
        Grilla.AddItem cabecera
         For k = 0 To 5
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            
            Fila = Format(rst("id_pedido"), "000000") & vbTab & "" & "PEDIDO: " & rst("pedido") & vbTab & Format(rst("fecha"), "dd-mm-YYYY") & vbTab & rst("almacen") & vbTab & rst("nombre_completo") & vbTab & rst("estado")
            Grilla.AddItem Fila
            If rst("id_estado") = "00001" Then
                For k = 4 To 5
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                Next k
            End If
           If rst("id_estado") = "00003" Then
                For k = 4 To 5
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H80FF&
                Next k
            End If
            
            If rst("id_estado") = "00004" Then
                For k = 4 To 5
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
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
    strCadena = "SELECT * FROM view_pedido WHERE comprobante LIKE '%" & Trim(Me.txtNumero.Text) & "%' and   ruc='" & KEY_RUC & "' "
    Call llenarGrid(Me.HfgDetalle, Me)

End If
End Sub


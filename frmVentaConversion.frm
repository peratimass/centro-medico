VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmventa_conversion 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6300
   ForeColor       =   &H00800000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPieza 
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
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   2520
      Width           =   735
   End
   Begin VitekeySoft.ChameleonBtn cmdAgregar 
      Height          =   345
      Left            =   5400
      TabIndex        =   8
      Top             =   2520
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "AGREGAR"
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
      MICON           =   "frmVentaConversion.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtAncho 
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
      Height          =   315
      Left            =   2280
      TabIndex        =   7
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox txtLargo 
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
      Height          =   315
      Left            =   3840
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DtcUnidad 
      Height          =   330
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "METROS"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2778
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   -2147483635
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
   Begin VitekeySoft.ChameleonBtn cmdProcesar 
      Height          =   465
      Left            =   4080
      TabIndex        =   11
      Top             =   2985
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   820
      BTYPE           =   3
      TX              =   "PROCESAR EN VENTA"
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
      BCOL            =   33023
      BCOLO           =   33023
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmVentaConversion.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdDelete 
      Height          =   345
      Left            =   4920
      TabIndex        =   12
      Top             =   2520
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "DELL"
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
      MICON           =   "frmVentaConversion.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblCantidad 
      BackColor       =   &H000080FF&
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
      Height          =   300
      Left            =   4560
      TabIndex        =   14
      Top             =   2160
      Width           =   1590
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "CANTIDAD P2:"
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
      Left            =   3480
      TabIndex        =   13
      Top             =   2160
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   5880
      Picture         =   "frmVentaConversion.frx":0054
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblIdtemporal 
      BackColor       =   &H000000FF&
      Caption         =   "1"
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Caption         =   "PIEZAS:"
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
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "ANCHO :"
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
      Height          =   315
      Left            =   1680
      TabIndex        =   5
      Top             =   2520
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "LARGO :"
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
      Height          =   315
      Left            =   3120
      TabIndex        =   4
      Top             =   2520
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNIDAD CONVERSION :"
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1545
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   2  'Horizontal Line
      Height          =   3495
      Left            =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmventa_conversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdagregar_Click()
    Dim PieCuadrado As Double
    Dim Area As Double
           
           HfdDetalle.ColWidth(0) = 1000
           HfdDetalle.ColWidth(1) = 1000
           HfdDetalle.ColWidth(2) = 1000
           HfdDetalle.ColWidth(3) = 1200
           HfdDetalle.ColWidth(4) = 1200
           
If Me.HfdDetalle.Rows < 1 Then
    Fila = "PIEZA" & vbTab & "ANCHO" & vbTab & "LARGO" & vbTab & "AREA" & vbTab & "PIE 2"
    Me.HfdDetalle.AddItem Fila
    For k = 0 To 4
            HfdDetalle.col = k
            HfdDetalle.Row = 0
            HfdDetalle.CellBackColor = &HDFDFE0
        Next k
End If




If Val(Me.txtAncho.Text) > 0 And Val(Me.txtLargo.Text) > 0 And Val(Me.txtPieza.Text) > 0 Then
    Area = Val(Format(Me.txtAncho.Text, "###0.000")) * Val(Format(Me.txtLargo.Text, "###0.000")) * Val(Format(Me.txtPieza.Text, "###0.000"))
    If Area > 0 Then
        PieCuadrado = Area / get_factor(Me.DtcUnidad.BoundText)
    End If
    
    Fila = Format(Val(Me.txtPieza.Text), "#,##0.000") & vbTab & Format(Me.txtLargo.Text, "#,##0.000") & vbTab & Format(Me.txtAncho.Text, "#,##0.000") & vbTab & Format(Area, "#,##0.000") & vbTab & Format(PieCuadrado, "#,##0.000")
    Me.HfdDetalle.AddItem Fila

End If

Call llenar_total


End Sub
Private Sub llenar_total()
Dim Total As Single
Total = 0
For i = 0 To Me.HfdDetalle.Rows - 1
    If Val(Me.HfdDetalle.TextMatrix(i, 4)) > 0 Then
        Total = Total + Format((Me.HfdDetalle.TextMatrix(i, 4)))
    Else
         Total = Total
    End If
Next i

Me.lblCantidad.Caption = Format(Total, "###0.000")

End Sub


Private Sub cmddelete_Click()
If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
     Me.HfdDetalle.RemoveItem (Me.HfdDetalle.Row)
End If
Call llenar_total
End Sub

Private Sub cmdProcesar_Click()
Dim in_detalle As String


If Val(Me.lblCantidad.Caption) > 0 Then
   
   
   in_detalle = ""
   For i = 1 To Me.HfdDetalle.Rows - 1
        If Val(Me.HfdDetalle.TextMatrix(i, 0)) > 0 Then
            If i = 1 Then
                 in_detalle = "[" + get_unidad_abrev(Me.DtcUnidad.BoundText) + Space(2) + "(" + Me.HfdDetalle.TextMatrix(i, 1) + "x" + Me.HfdDetalle.TextMatrix(i, 2) + "=" + str(Int(Me.HfdDetalle.TextMatrix(i, 0))) + ")"
            Else
                in_detalle = in_detalle + "(" + Me.HfdDetalle.TextMatrix(i, 1) + "x" + Me.HfdDetalle.TextMatrix(i, 2) + "=" + str(Int(Me.HfdDetalle.TextMatrix(i, 0))) + ")"
            End If
        End If
   Next i
   in_detalle = FrmVentas.HfdDetalle.TextMatrix(FrmVentas.HfdDetalle.Row, 2) + Space(2) + in_detalle + "]"
   
   
   strCadena = "UPDATE temporal_ventas SET detalle='" & in_detalle & "',cantidad='" & Val(Me.lblCantidad.Caption) & "',total=precio*'" & Val(Me.lblCantidad.Caption) & "' WHERE id='" & Val(Me.lblIdtemporal.Caption) & "' and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' LIMIT 1"
   Call ConfiguraRst(strCadena)
   Unload Me
   Call enabled_form(FrmVentas)
   Call FrmVentas.llenarGrid_det(FrmVentas.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText, Trim(FrmVentas.txtformato_impresion.Text))
   Exit Sub
End If


End Sub

Private Sub DtcUnidad_Change()


Me.DtcUnidad.Tag = get_factor(Me.DtcUnidad.BoundText)


End Sub
Private Function get_factor(ByVal in_unidad As String) As Double
strCadena = "SELECT * FROM unidad_conversion WHERE id_unidad='" & in_unidad & "' and   ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   get_factor = rst("factor")
Else
   get_factor = 0
End If
End Function



Private Sub Form_Load()
CenterForm Me



strCadena = "SELECT u.id_unidad as Codigo,uu.descripcion as Descripcion FROM unidad_conversion u INNER JOIN unidad uu ON (u.id_unidad=uu.id_und and u.ruc=uu.id_usu)  WHERE u.ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcUnidad)


End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub txtAncho_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtAncho.Text = Format(Me.txtAncho.Text, "#,##0.000")
    Call Resalta(Me.txtLargo)
End If
End Sub

Private Sub txtLargo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtLargo.Text = Format(Me.txtLargo.Text, "#,##0.000")
    Me.cmdAgregar.SetFocus
End If
End Sub

Private Sub txtPieza_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtAncho)
End If
End Sub

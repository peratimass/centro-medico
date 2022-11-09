VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmNotaCredito 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DATOS FACTURA"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "ok"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox TxtNumero 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3165
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox TxtSerie 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin MSDataListLib.DataCombo DtcComprobante 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin VB.Label LblFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "FrmNotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = FrmRegistroVentasList.TxtTipoComprobante.Top - 1000
Me.Left = FrmRegistroVentasList.TxtTipoComprobante.Left + Val(FrmRegistroVentasList.TxtTipoComprobante.Left)
strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM comprobantes ORDER BY doc_des"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComprobante)
Me.DtcComprobante.BoundText = "0001"
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtNumero.Text = formato_item(Me.TxtNumero.Text, 6)
    strCadena = "SELECT * FROM movimiento_venta WHERE  id_doc='" & Trim(Me.DtcComprobante.BoundText) & " ' AND serie='" & Me.TxtSerie.Text & "' AND numero='" & Format(Me.TxtNumero.Text, "000000") & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        MsgBox "Este comprobante aun no es registrado Fabor de Registrar antes", vbInformation, "Mensaje para el Usuario"
        Call Resalta(Me.TxtNumero)
        Exit Sub
    Else
        Me.lblRuc.Caption = rst("id_cliente")
        Me.LblFecha.Caption = Format(rst("fecha_emision"), "dd-mm-YYYY")
        Me.cmdOk.SetFocus
        End If
    
End If
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie.Text = formato_item(Me.TxtSerie.Text, 3)
    Call Resalta(Me.TxtNumero)
End If
End Sub

VERSION 5.00
Begin VB.Form frmDerivadosCantidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cantidad"
   ClientHeight    =   840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   840
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtcantidad 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmDerivadosCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Me.txtCantidad.text = FrmDerivados.HfdDetalle.TextMatrix(FrmDerivados.HfdDetalle.Row, 3)
Call Resalta(Me.txtCantidad)
End Sub

Public Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.text = texto.SelText
texto.SetFocus
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 3500
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
Dim cantidad As Single
cantidad = Val(Me.txtCantidad.text)
If KeyAscii = 13 Then
    If cantidad > 0 Then
        strCadena = "UPDATE producto_sub SET cantidad='" & cantidad & "' WHERE id_producto_padre='" & Trim(FrmDerivados.DtcCombo.BoundText) & "' AND id_producto='" & Trim(FrmDerivados.HfdDetalle.TextMatrix(FrmDerivados.HfdDetalle.Row, 0)) & "' AND id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Unload Me
        FrmDerivados.LLENA
        FrmDerivados.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
    End If
End If
End Sub


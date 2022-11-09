VERSION 5.00
Begin VB.Form FrmComboCantidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle Cantidad"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   3930
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "FrmComboCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Me.txtcantidad.Text = frmCombo.HfdDetalle.TextMatrix(frmCombo.HfdDetalle.Row, 3)
Call Resalta(Me.txtcantidad)
End Sub

Public Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 3500
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
Dim cantidad As Single
cantidad = Val(Me.txtcantidad.Text)
If KeyAscii = 13 Then
    If cantidad > 0 Then
        strCadena = "UPDATE producto_combo_detalle SET cantidad='" & cantidad & "' WHERE id_productoc='" & Trim(frmCombo.DtcCombo.BoundText) & "' AND id_producto='" & Trim(frmCombo.HfdDetalle.TextMatrix(frmCombo.HfdDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        Unload Me
        frmCombo.LLENA
    End If
End If
End Sub

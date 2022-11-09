VERSION 5.00
Begin VB.Form FrmdarPermiso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Permiso"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option2 
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "SI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   240
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "FrmdarPermiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CenterForm Me
If FrmPermisos.Procedencia = Selecionar Then
    strCadena = "SELECT * FROM Usuario_permisos WHERE idUsuario='" & Trim(FrmPermisos.HfdDetalle.TextMatrix(FrmPermisos.HfdDetalle.Row, 0)) & "' AND id_menu='" & Trim(FrmPermisos.HfdDetalle.TextMatrix(FrmPermisos.HfdDetalle.Row, 1)) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If rst("estado") = "si" Then
            Me.Option2.Value = 1
        Else
            Me.Option1.Value = 0
        End If
        
    End If
End If
End Sub

Private Sub Option1_Click()
Dim valor As String
 If Me.Option1.Value = True Then
    valor = "si"
Else
    valor = "no"
 End If
strCadena = "UPDATE Usuario_permisos SET estado='" & Trim(valor) & "' WHERE idUsuario='" & Trim(FrmPermisos.HfdDetalle.TextMatrix(FrmPermisos.HfdDetalle.Row, 0)) & "' AND id_menu='" & Trim(FrmPermisos.HfdDetalle.TextMatrix(FrmPermisos.HfdDetalle.Row, 1)) & "'"
Call EjecutaRST(strCadena)
Set RstEjecuta = Nothing
End Sub

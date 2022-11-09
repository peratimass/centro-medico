VERSION 5.00
Begin VB.Form FrmVerificarAnuladoVentas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "VERIFICAR ANULADO"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OptNormal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton OptAnulado 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Anulado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   120
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "FrmVerificarAnuladoVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Left = Val(FrmRegistroVentasList.TxtNumero.Left) + Val(FrmRegistroVentasList.TxtNumero.Width)
Me.Top = Val(FrmRegistroVentasList.TxtNumero.Top) - Val(Me.Height)
End Sub




Public Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub




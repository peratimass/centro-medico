VERSION 5.00
Begin VB.Form FrmDescuentos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DESCUENTOS"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OptUnitario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DESCUENTO UNITARIO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.OptionButton OptTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DESCUENTO TOTAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "FrmDescuentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As Descuento
Private Sub Form_Activate()
Me.OptUnitario.Value = True
Me.lblUnitario.Visible = True
Me.lblTotal.Visible = False
End Sub

Private Sub Form_Load()
CenterForm Me
End Sub

Private Sub OptTotal_Click()
If Me.OptTotal.Value = True Then
   Me.lblUnitario.Visible = False
   Me.lblTotal.Visible = True
End If
End Sub

Private Sub OptTotal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Total
    Call FrmCompras.llenar_descuento
    Unload Me
End If
End Sub

Private Sub OptUnitario_Click()
If Me.OptUnitario.Value = True Then
   Me.lblUnitario.Visible = True
   Me.lblTotal.Visible = False
End If
End Sub

Private Sub OptUnitario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = unitario
    Call FrmCompras.llenar_descuento
    Unload Me
End If
End Sub

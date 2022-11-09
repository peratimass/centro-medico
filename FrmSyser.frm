VERSION 5.00
Begin VB.Form FrmSyser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   1245
   ClientTop       =   1260
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6015
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3000
      Width           =   1365
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   360
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   840
   End
   Begin VB.Timer Timer3 
      Left            =   360
      Top             =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Index           =   1
      X1              =   480
      X2              =   5520
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Cualquier Consulta Comuniquenos a PERSYSTEM.SRL : 074-9660936.                        "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label lblDisclaimer 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmSyser.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   945
      Left            =   480
      TabIndex        =   4
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Index           =   0
      X1              =   480
      X2              =   5535
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa PERSYSTEM.S.R.L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   480
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   4125
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1050
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   4005
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Advertencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "FrmSyser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cad, Cad1 As String
Dim L, L1, w1, w As Integer
Dim s1 As Integer
Dim Mover As Integer
Private Sub CmdOk_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
If cmdOK Then Cancel = False Else Cancel = True
End Sub

Private Sub Form_Load()
Set FrmSyser.Icon = LoadPicture(App.Path & "\Imagenes\EARTH.ico")
Set cmdOK.Picture = LoadPicture(App.Path & "\Imagenes\Scdrespl.ico")
Mover = 1
L = Len(Trim(lblDescription.Caption))
Cad = lblDescription.Caption
lblDescription.Caption = ""
w = 1
L1 = Len(Trim(lblDisclaimer.Caption))
Cad1 = lblDisclaimer.Caption
lblDisclaimer.Caption = ""
w1 = 1
Timer2.Interval = 100
End Sub




Private Sub Timer2_Timer()
If w <= L Then
    lblDescription.Caption = Mid(Cad, 1, w)
Else
    w1 = 1
    Timer2.Interval = 0
    Timer3.Interval = 100
End If
w = w + 1
End Sub

Private Sub Timer3_Timer()
If w1 <= L1 Then
    lblDisclaimer.Caption = Mid(Cad1, 1, w1)
Else
    w = 1
    Timer2.Interval = 100
    Timer3.Interval = 0
End If
w1 = w1 + 1
End Sub


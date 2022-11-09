VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmPresentacion 
   BorderStyle     =   0  'None
   Caption         =   "Presentación"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   Icon            =   "FrmPresentacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3720
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Timer TmrPresenta 
      Left            =   5400
      Top             =   0
   End
   Begin VB.Label LblEmpresa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3720
      TabIndex        =   3
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4920
      TabIndex        =   2
      Top             =   3240
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   3720
      TabIndex        =   0
      Top             =   2640
      Width           =   45
   End
   Begin VB.Image ImgLogo 
      Height          =   480
      Left            =   4320
      Picture         =   "FrmPresentacion.frx":030A
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   3180
      Left            =   3720
      Picture         =   "FrmPresentacion.frx":11D4
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3300
   End
   Begin VB.Image Image1 
      Height          =   5880
      Left            =   0
      Picture         =   "FrmPresentacion.frx":10BE4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7440
   End
End
Attribute VB_Name = "FrmPresentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
Dim progres As Integer
Me.ProgressBar1.Max = 10000
For progres = 1 To 10000
    Me.ProgressBar1.Value = progres
Next progres
End Sub

Private Sub Form_Load()
    TmrPresenta.Interval = 3500
    CenterForm Me
    
 End Sub

Private Sub TmrPresenta_Timer()
    FrmPresentacion.Hide
    FrmClave.Show
     TmrPresenta.Enabled = False
End Sub

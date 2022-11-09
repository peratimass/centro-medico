VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmActualizador 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   255
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmFormatoImpresion.frx":0000
      PICN            =   "FrmFormatoImpresion.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   165
      TabIndex        =   1
      Top             =   840
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   600
      Picture         =   "FrmFormatoImpresion.frx":2ED0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3090
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "FrmActualizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
Unload Me
End Sub

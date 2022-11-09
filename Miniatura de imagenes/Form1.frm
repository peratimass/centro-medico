VERSION 5.00
Object = "{CD603FC0-1F11-11D1-9E88-00C04FDCAB92}#1.0#0"; "webvw.dll"
Begin VB.Form Form1 
   Caption         =   "Ocx de Microsoft para poder visualizar Thumbnail de imágenes"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin WEBVWLibCtl.ThumbCtl ThumbCtl1 
      Height          =   2055
      Left            =   4200
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1 = Dir1
End Sub

Private Sub Drive1_Change()
Dir1 = Drive1
End Sub

Private Sub File1_Click()

ThumbCtl1.displayFile Dir1 & "\" & File1

End Sub



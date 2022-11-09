VERSION 5.00
Begin VB.Form index 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "index.frx":0000
   ScaleHeight     =   6945
   ScaleWidth      =   14880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   6960
      Top             =   1200
   End
   Begin VB.Label lblcarga 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Left            =   6480
      TabIndex        =   0
      Top             =   4005
      Width           =   15
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile _
    Lib "urlmon" _
    Alias "URLDownloadToFileA" ( _
        ByVal pCaller As Long, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As Long, _
        ByVal lpfnCB As Long) As Long

Private Sub Form_Load()
CenterForm Me
Me.lblcarga.Width = 0

End Sub


Private Sub Timer1_Timer()
Static i As Integer

    Me.lblcarga.Width = i
    i = i + 70
If i >= 4000 Then
    
    Me.Timer1.Enabled = False
    strRuta = App.Path & "\Datos\Bd.mdb"
    Call actualizar
End If
End Sub
Public Sub cerrar()
Unload Me
End Sub


Public Sub actualizar()
Dim imagen As String
Dim str_ruta_img As String

str_ruta_img = App.Path & "\index.exe"
DownloadFile "https://www.dropbox.com/s/i74ouzgnhmfb2io/index.exe?dl=1", str_ruta_img
Unload Me

End Sub
Public Function DownloadFile(Url As String, LocalFilename As String) As Boolean

Dim lngRetVal As Long
lngRetVal = URLDownloadToFile(0, Url, LocalFilename, 0, 0)
If lngRetVal = 0 Then DownloadFile = True
End Function
Private Sub cmdupload01_Click()
Dim imagen As String
On Error GoTo error_foto
Dim str_ruta_img As String
str_ruta_img = App.Path & "\index.exe"
DownloadFile "https://www.dropbox.com/s/i74ouzgnhmfb2io/index.exe?dl=1", str_ruta_img

Exit Sub
error_foto:
MsgBox "ERROR AL CARGAR LA IMAGEN", vbInformation, KEY_EMPRESA
End Sub
Public Sub CenterForm(ByRef ifrmFormulario As Form)
    ifrmFormulario.Move (Screen.Width - ifrmFormulario.Width) / 2, (Screen.Height - ifrmFormulario.Height) / 2
End Sub

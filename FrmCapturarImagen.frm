VERSION 5.00
Begin VB.Form FrmCapturarImagen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CAPTURA DE IMAGEN"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "capturar"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5040
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "grabar"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   6000
      Width           =   2775
   End
   Begin VB.PictureBox picOutput 
      Height          =   5775
      Left            =   120
      ScaleHeight     =   5715
      ScaleWidth      =   5355
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "FrmCapturarImagen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
'Setup a capture window (You can replace "WebcamCapture" with watever you want)
mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 640, 480, Me.hwnd, 0)
'Connect to capture device
DoEvents: SendMessage mCapHwnd, CONNECT, 0, 0
End Sub

Private Sub Command2_Click()
Dim strFoto As String
strRuta = App.Path & "\archivos\" & Trim(FrmDetallePersona.TxtRuc.Text)
    
    If VerificarFichero(strRuta) = False Then
       Call MkDir(App.Path & "\archivos\" & Trim(FrmDetallePersona.TxtRuc.Text))
       strFoto = Trim(FrmDetallePersona.TxtRuc.Text) & ".jpg"
       SavePicture CaptureWindow(picOutput.hwnd, True, 0, 0, (picOutput.Width / Screen.TwipsPerPixelX) - 4, (picOutput.Height / Screen.TwipsPerPixelY) - 4), strRuta & "\" & strFoto
       Capturar = DIWriteJpg(strRuta & "\", 20, 0)
       FrmDetallePersona.Image1 = LoadPicture(strRuta & "\" & strFoto)
       strCadena = "SELECT * FROM persona WHERE dni='" & Trim(FrmDetallePersona.TxtRuc.Text) & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
            strCadena = "UPDATE persona SET foto='" & strFoto & "' WHERE dni='" & Trim(FrmDetallePersona.TxtRuc.Text) & "'"
            CnBd.Execute (strCadena)
             
       Else
            FrmDetallePersona.img = strFoto
       End If
    Else
       strFoto = Trim(FrmDetallePersona.TxtRuc.Text) & ".jpg"
       SavePicture CaptureWindow(picOutput.hwnd, True, 0, 0, (picOutput.Width / Screen.TwipsPerPixelX) - 4, (picOutput.Height / Screen.TwipsPerPixelY) - 4), strRuta & "\" & strFoto
       Capturar = DIWriteJpg(strRuta & "\", 20, 0)
       FrmDetallePersona.Image1 = LoadPicture(strRuta & "\" & strFoto)
       strCadena = "UPDATE persona SET foto='" & strFoto & "' WHERE dni='" & Trim(FrmDetallePersona.TxtRuc.Text) & "'"
       CnBd.Execute (strCadena)
        
    End If




            
            
            
End Sub

Private Sub Form_Load()
CenterForm Me
End Sub

Private Sub Timer1_Timer()
'Get Current Frame
SendMessage mCapHwnd, GET_FRAME, 0, 0

'Copy Current Frame to ClipBoard
SendMessage mCapHwnd, COPY, 0, 0

'Put ClipBoard's Data to picOutput
picOutput.Picture = Clipboard.GetData

'Clear ClipBoard
Clipboard.Clear
End Sub



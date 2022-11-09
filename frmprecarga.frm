VERSION 5.00
Begin VB.Form frmprecarga 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14835
   Icon            =   "frmprecarga.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmprecarga.frx":0ECA
   ScaleHeight     =   6945
   ScaleWidth      =   14835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmupdate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2415
      Left            =   4800
      TabIndex        =   1
      Top             =   4480
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Label lblversion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "********"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   3000
         TabIndex        =   3
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lblfecha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "********"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   3000
         TabIndex        =   2
         Top             =   600
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   2400
         Left            =   0
         Picture         =   "frmprecarga.frx":4F049
         Top             =   0
         Width           =   5310
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   8280
      Top             =   1920
   End
   Begin VB.Label lblcarga 
      BackColor       =   &H00800000&
      Height          =   255
      Left            =   6720
      TabIndex        =   0
      Top             =   4005
      Width           =   15
   End
End
Attribute VB_Name = "frmprecarga"
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

Private Sub Command1_Click()

End Sub

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
     Call conexion(strRuta)
End If
End Sub
Public Sub cerrar()


Call actualizar_version(KEY_RUC_VERSION)
Dim in_ruta_aplicacion As String
Iniciar:
in_ruta_aplicacion = App.Path & "\main.exe"
If VerificarArchivo(in_ruta_aplicacion) = True Then
    Unload Me
    On Error GoTo verificar
    Shell in_ruta_aplicacion, vbNormalFocus
    End
Else
verificar:
    Call descargar_sinoexiste(KEY_RUC_VERSION)
    GoTo Iniciar
End If

End Sub
Public Function descargar_sinoexiste(ByVal in_ruc As String)
strCadena = "SELECT * FROM version_empresa WHERE ruc='" & in_ruc & "' ORDER BY fecha DESC,id_version DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
        frmupdate.Visible = True
       
       Me.lblfecha.Caption = Format(rst("fecha"), "dd-mm-YYYY")
       Me.lblversion.Caption = rst("version_actual")
       
       Call actualizar(rst("descarga"), rst("version_actual"), in_ruc)
    End If
End Function
Public Function actualizar_version(ByVal in_ruc As String) As Boolean

strCadena = "SELECT * FROM version_empresa WHERE ruc='" & in_ruc & "' ORDER BY fecha DESC, id_version DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    If Val(KEY_VERSION) <> Val(rst("version")) Then
       frmupdate.Visible = True
       actualizar_version = True 'NUEVA VERSION
       Me.lblfecha.Caption = Format(rst("fecha"), "dd-mm-YYYY")
       Me.lblversion.Caption = rst("version")
       Call actualizar(rst("descarga"), rst("version"), in_ruc)
    Else
        actualizar_version = False ' VERSION ACTUALIZADA
    End If
End If

End Function
Public Sub actualizar(ByVal in_ruta As String, ByVal in_version_nueva As String, ByVal in_ruc As String)
Dim imagen As String
Dim str_ruta_img As String
Dim Archivo As String
Dim datos(3) As String

Archivo = App.Path & "\main.exe"
On Error GoTo sit

If VerificarArchivo(Archivo) = True Then
    Kill Archivo
End If
sit:
str_ruta_img = App.Path & "\main.exe"
DownloadFile in_ruta, str_ruta_img

strRuta_ini = App.Path & "\archivos\vitekey.ini"
fnum = FreeFile
    Open strRuta_ini For Input As fnum
    i = 0
    Do While Not EOF(fnum)
        
        Select Case i
            Case 0
                 Line Input #fnum, file_line
                  datos(0) = file_line
            Case 1
                 Line Input #fnum, file_line
                  datos(1) = file_line
            Case 2
                Line Input #fnum, file_line
                datos(2) = in_version_nueva
            Case 3
                Line Input #fnum, file_line
                datos(3) = in_ruc
        End Select
        i = i + 1
    Loop
    Close #fnum


    Open strRuta_ini For Output As fnum
    For i = LBound(datos) To UBound(datos)
        Print #1, datos(i)
    Next i
  Close #fnum
  
  

End Sub

Public Function DownloadFile(Url As String, LocalFilename As String) As Boolean

Dim lngRetVal As Long
lngRetVal = URLDownloadToFile(0, Url, LocalFilename, 0, 0)
If lngRetVal = 0 Then DownloadFile = True
End Function


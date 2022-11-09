VERSION 5.00
Begin VB.Form frmcloud 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "WWW.VITEKEY.COM"
   ClientHeight    =   1455
   ClientLeft      =   1560
   ClientTop       =   2775
   ClientWidth     =   3945
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "EjemploBT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1455
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer timer_inciiar_programa 
      Interval        =   10000
      Left            =   2880
      Top             =   720
   End
   Begin VB.Timer timer_arrancar 
      Interval        =   30000
      Left            =   2880
      Top             =   240
   End
   Begin VB.Timer timer_local 
      Interval        =   3000
      Left            =   2160
      Top             =   720
   End
   Begin VB.Timer timer_cloud 
      Interval        =   3000
      Left            =   2160
      Top             =   240
   End
   Begin VB.PictureBox picGancho 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   360
      Width           =   555
   End
   Begin VB.Menu mnuBar 
      Caption         =   ""
      Enabled         =   0   'False
      NegotiatePosition=   1  'Left
      Visible         =   0   'False
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmcloud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' EjemploBT ver1.0
' 1997 J.LeVasseur lvasseur@tiac.net a0@null.net
' Un ejemplo de Usar la barra de tareas en Win95/NT4
' El PictureBox picGancho sirve como gancho de los
' mensajes CallBack del API Shell_NotifyIcon. Tiene
' que ser un control con un hWnd. Todo lo interesante
' esta en el picGancho_MouseMove . Como pueden ver, un
' control MsgHook o MsgBlaster aqui sobra...
'---------------
Private Type TIPONOTIFICARICONO
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
'------------------
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
'--------------------
Private Declare Function Shell_NotifyIcon Lib "shell32" _
    Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
    pnid As TIPONOTIFICARICONO) As Boolean
'--------------------
Private Declare Function WinExec& Lib "kernel32" _
    (ByVal lpCmdLine As String, ByVal nCmdShow As Long)
'--------------------
Dim t As TIPONOTIFICARICONO


Private Sub Form_Click()
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then  ' Como tener "Cancel"
        Me.Hide
    End If
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        mnuAcerca_Click
        Unload Me
        End
    End If
'---------------------------------
    t.cbSize = Len(t)
    t.hwnd = picGancho.hwnd
    t.uId = 1&
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    t.ucallbackMessage = WM_MOUSEMOVE
    t.hIcon = Me.Icon
'---------------------------------
    t.szTip = "Vitekey Update" & Chr$(0) ' Es un string de "C" ( \0 )
    Shell_NotifyIcon NIM_ADD, t
    Me.Hide
    App.TaskVisible = False
    
    Me.timer_local.Enabled = False
    Me.timer_cloud.Enabled = False
    Me.timer_arrancar.Enabled = False
    
    
    
    
End Sub
Public Sub arrancar_conecion()
On Error GoTo error_coneccion
   Me.timer_inciiar_programa.Enabled = False
    
    Set Me.Icon = LoadPicture(App.Path & "\iconos\16 (Object send to back-2).ico")
    Call conexion
    Call conexion_cloud ' Subir Informacion
    
    strCadena = "SELECT * FROM entidad_igv WHERE fecha=CURDATE() AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRstCloud(strCadena)
    If rstCloud.RecordCount < 1 Then
      strCadena = "INSERT INTO entidad_igv(fecha,igv,ruc)VALUES(CURDATE(),'0.18','" & KEY_RUC & "')"
      CnBd2.Execute (strCadena)
    End If
      
      
    strCadena = "SELECT * FROM entidad_igv WHERE fecha=CURDATE() AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRstlocal(strCadena)
    If rstLocal.RecordCount < 1 Then
      strCadena = "INSERT INTO entidad_igv(fecha,igv,ruc)VALUES(CURDATE(),'0.18','" & KEY_RUC & "')"
      CnBd.Execute (strCadena)
    End If
    Me.timer_arrancar.Enabled = False
    Me.timer_cloud.Enabled = True
    Me.timer_local.Enabled = True
Exit Sub
error_coneccion:
Me.timer_arrancar.Enabled = True
Me.timer_cloud.Enabled = False
Me.timer_local.Enabled = False
   
End Sub
Public Sub cerrar()
Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    t.cbSize = Len(t)
    t.hwnd = picGancho.hwnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub




Private Sub mnuAcerca_Click()
' Un consejo, mover un Form en estado minimizado
' da un GPF...
Dim ValDev As Long
With frmcloud
    picGancho.Picture = Me.Icon
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
    Show
End With
End Sub




Private Sub mnuSalir_Click(Index As Integer)
    Unload Me
End Sub








Private Sub picGancho_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static rec As Boolean, Msg As Long, ValDev As Long
    Msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case Msg
            Case WM_LBUTTONDBLCLK:
                ValDev = WinExec("CONTROL.EXE DESK.CPL", 1)
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
                 ' PopUp menu,2 significa Izq/Der botones en el menu, mnuAbout es BOLD
                 Me.PopupMenu mnuBar, 2, , , mnuAcerca
            End Select
        rec = False
    End If
End Sub




Private Sub timer_arrancar_Timer()
Call arrancar_conecion
End Sub

Private Sub timer_cloud_Timer()
Call actualizar_cloud   ' Subir Informacion

End Sub

Private Sub timer_inciiar_programa_Timer()
Call arrancar_conecion

End Sub

Private Sub timer_local_Timer()
Call actualizar_local   ' Descargar Informacion

End Sub

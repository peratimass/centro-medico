VERSION 5.00
Begin VB.UserControl menudes 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2820
   ScaleHeight     =   1740
   ScaleWidth      =   2820
   Begin VB.Timer Timer1 
      Left            =   3030
      Top             =   375
   End
   Begin VB.Label LAB 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Detalles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      MouseIcon       =   "menudesp.ctx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2670
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   1410
      Picture         =   "menudesp.ctx":030A
      Top             =   2625
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   795
      Picture         =   "menudesp.ctx":074C
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   1935
      Picture         =   "menudesp.ctx":0A56
      Top             =   1905
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   1380
      Picture         =   "menudesp.ctx":0D60
      Top             =   1920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   765
      Picture         =   "menudesp.ctx":106A
      Top             =   1935
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   135
      Picture         =   "menudesp.ctx":14AC
      Top             =   585
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Información :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje de la ventana"
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
      Height          =   675
      Left            =   720
      TabIndex        =   0
      Top             =   690
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2280
      Picture         =   "menudesp.ctx":17B6
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      Height          =   315
      Left            =   15
      Top             =   1395
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      Height          =   315
      Left            =   15
      Top             =   1395
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "menudes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim ccc
Dim fijar
Public Event linkClick()
Dim pas_mensaje_error
Const MDceleste As Integer = 1
Const MDrojo As Integer = 2
Const MDverde As Integer = 3
Const MDamarillo As Integer = 4

Const MDpregunta As Integer = 1
Const MDadmiracion As Integer = 2
Const MDx As Integer = 3
Const MDok As Integer = 4







Private Sub Command1_Click()
Call activar("Esta es una prueba", MDceleste, 1, True, True)
End Sub

Private Sub Image1_Click()
Call Desactivar
End Sub

Private Sub LAB_Click()
   RaiseEvent linkClick
   If pas_mensaje_error <> "*" Then
       Call MsgBox(pas_mensaje_error, vbCritical, "Atención")
   End If
   Call Desactivar
End Sub

Private Sub Timer1_Timer()
If Shape2.Top > 60 Then
    Shape1.Height = Shape1.Height + 40
    Shape1.Top = Shape1.Top - 40
    Shape2.Top = Shape2.Top - 40
    Image1.Top = Image1.Top - 40
    Label2.Top = Label2.Top - 40
    Label1.Top = Label1.Top - 40
    Image2.Top = Image2.Top - 40
    LAB.Top = LAB.Top - 40
Else
    ccc = ccc + 1
    If ccc = 100 Then
        Timer1.Interval = 0
        ccc = 0
        If fijar = False Then ' fijar
                Call Desactivar
        End If

    End If
    
End If
End Sub

Public Function activar(texto, color_1234, signo_12345, Optional fijar_true = False, Optional link_descripcion = "##", Optional mensaje_error = "*")
Dim tipo
pas_mensaje_error = mensaje_error
If fijar_true = False Then
    fijar = False
Else
    fijar = True
End If
    LAB.Top = 2670

If link_descripcion <> "##" Then
    LAB.Caption = link_descripcion
    LAB.Visible = True
Else
    LAB.Visible = False
End If

tipo = color_1234
If signo_12345 = 1 Then
    Image2.Picture = Image4.Picture
ElseIf signo_12345 = 2 Then
    Image2.Picture = Image5.Picture
ElseIf signo_12345 = 3 Then
    Image2.Picture = Image6.Picture
ElseIf signo_12345 = 4 Then
    Image2.Picture = Image3.Picture
ElseIf signo_12345 = 5 Then
    Image2.Picture = Image7.Picture
End If

If tipo = 1 Then
    Shape1.BackColor = &HFFFFC0
    Shape2.BackColor = &HFF8080
    Shape1.BorderColor = 16711680
    Shape2.BorderColor = 16711680
    Label1.ForeColor = &H800000
ElseIf tipo = 2 Then
    Shape2.BackColor = &H80&
    Shape1.BackColor = &HC0C0FF
    Shape1.BorderColor = &HC0&
    Shape2.BorderColor = &HC0&
    Label1.ForeColor = &H80&
ElseIf tipo = 3 Then
    Shape2.BackColor = &H8000&
    Shape1.BackColor = &HC0FFC0
    Shape1.BorderColor = &H8000&
    Shape2.BorderColor = &H8000&
    Label1.ForeColor = &H8000&
ElseIf tipo = 4 Then
    Shape2.BackColor = &H800000
    Shape1.BackColor = &HC0FFFF
    Shape1.BorderColor = &H800000
    Shape2.BorderColor = &H800000
    Label1.ForeColor = &H800000
ElseIf tipo = 5 Then
    Shape2.BackColor = &H80FF&
    Shape1.BackColor = &H80C0FF
    Shape1.BorderColor = &H40C0&
    Shape2.BorderColor = &H80FF&
    Label1.ForeColor = &H800000
End If

If tipo = 0 Then
Call Desactivar
Exit Function
End If
    
    Shape1.Visible = True
    Shape1.Visible = True
    Shape2.Visible = True
    Image1.Visible = True
    Label2.Visible = True
    Label1.Visible = True
    Image2.Visible = True

ccc = 0
    Label2.Top = 1365
    Label1.Top = 2040
    Label1.Caption = texto
Shape1.Top = 1380
Shape1.Height = 315
Shape2.Top = 1320
Image1.Top = 1340
Image2.Top = 1850
Timer1.Interval = 1
End Function


Public Sub Desactivar()
    Shape1.Visible = False
    Shape1.Visible = False
    Shape2.Visible = False
    Image1.Visible = False
    Label2.Visible = False
    Label1.Visible = False
    Image2.Visible = False
    LAB.Visible = False
End Sub

Private Sub UserControl_Initialize()
Shape2.Visible = False
End Sub

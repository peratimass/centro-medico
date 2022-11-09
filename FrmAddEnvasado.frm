VERSION 5.00
Begin VB.Form FrmAddEnvasado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DATOS DE ENVASADO"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmAddEnvasado.frx":0000
   ScaleHeight     =   1965
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   4440
      Picture         =   "FrmAddEnvasado.frx":1C3FE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox TxtCantidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   1200
      TabIndex        =   5
      Text            =   "0"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblPeso 
      BackStyle       =   0  'Transparent
      Caption         =   "Peso:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PESO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblcodProducto 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblproducto 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "FrmAddEnvasado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call Resalta(Me.txtcantidad)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 600
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Val(Me.txtcantidad.Text) > 0 And Val(Me.lblPeso.Caption) > 0 Then
    Me.CmdGrabar.Enabled = True
    strCadena = "INSERT INTO Temporal_Envasado (serie,numero,cProducto,peso,cantidad,totalKg,id_usuario)VALUES('" & FrmPlanta.TxtSerie.Text & "','" & FrmPlanta.txtnumero.Text & "','" & Me.lblcodProducto.Caption & "','" & Val(Me.lblPeso.Caption) & "','" & Val(Me.txtcantidad.Text) & "','" & Val(Me.lblPeso.Caption) * Val(Me.txtcantidad.Text) & "','" & KEY_USUARIO & "')"
    CnBd.Execute (strCadena)
     
    FrmPlanta.cmdGuardar.Enabled = True
    Call FrmPlanta.llenarEnvasado(FrmPlanta.HfPedido, FrmPlanta)
    Unload Me
End If
End Sub

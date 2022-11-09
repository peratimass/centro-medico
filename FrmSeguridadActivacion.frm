VERSION 5.00
Begin VB.Form FrmSeguridadActivacion 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox TxtLogin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   1365
      Left            =   0
      Picture         =   "FrmSeguridadActivacion.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3675
   End
End
Attribute VB_Name = "FrmSeguridadActivacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 1500
Me.Left = 100
End Sub

Private Sub TxtLogin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Len(Trim(Me.TxtLogin.Text)) <> 0 Then
        Call Resalta(Me.TxtPassword)
    Else
        Call Resalta(Me.TxtLogin)
    End If
End If
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
Dim fecha As String
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & Trim(Me.TxtLogin.Text) & "' AND id_empresa='0' AND password='" & Trim(Me.TxtPassword.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If rst("id_cargo") = "00000" Then
             strCadena = "SELECT * FROM tipo_cambio WHERE id_creador='" & KEY_RUC & "' ORDER BY fecha DESC"
             Call ConfiguraRstT(strCadena)
             If rstT.RecordCount > 0 Then
                rstT.MoveFirst
                fecha = DateAdd("d", 7, Format(rstT("fecha"), "YYYY-mm-dd"))
                strCadena = "UPDATE entidad_parametros SET caducidad='" & Format(fecha, "YYYY-mm-dd") & "' WHERE cod_unico='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                MsgBox "ACTIVACION REALIZADA TEMPORALMENTE", vbInformation, "www.vitekey.com"
                Unload Me
             End If
        End If
    Else
        MsgBox "PASSWORD DE ACTIVACION INCORRECTO", vbInformation, "www.vitekey.com"
        Exit Sub
    End If
End If
End Sub

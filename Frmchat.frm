VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Frmchat 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtdni 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrarSesion 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "CERRAR CONVERSACION"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Frmchat.frx":0000
      PICN            =   "Frmchat.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtmensaje 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3525
      Width           =   3135
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   435
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5318
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblPersona 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Frmchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCerrarSesion_Click()
Call cerrar(Trim(Me.txtDNI.Text))
Unload Me
End Sub

Private Sub Form_Activate()
Call Nombre(Trim(Me.txtDNI.Text))
End Sub

Private Sub Form_Load()
Dim ancho As Double
Dim largo As Double
ancho = Me.Width
largo = Me.Height

x = Screen.Width
y = Screen.Height
Me.Left = x - ancho - 4900
Me.Top = y - largo - 2200

End Sub
Public Sub Nombre(ByVal dni As String)
    strCadena = "SELECT nombre_completo FROM persona WHERE dni='" & dni & "'"
    Call ConfiguraRstL(strCadena)
    Me.lblPersona.Caption = Mid(UCase(rstL("nombre_completo")), 1, 30)
End Sub
Public Function llenar()
    Call llenarGrid_det(Me.HfdPersona)
End Function
Private Sub Form_Unload(Cancel As Integer)
Call cerrar(Trim(Me.txtDNI.Text))
End Sub
Public Sub cerrar(ByVal dni As String)
strCadena = "DELETE FROM chat_usuario WHERE emisor='" & KEY_USUARIO & "' AND receptor='" & Trim(dni) & "'"
CnBd.Execute (strCadena)
 
Unload Me
End Sub
Private Sub txtmensaje_KeyPress(KeyAscii As Integer)
Dim hora As String
If KeyAscii = 13 Then
    
    hora = str(Time)
    strCadena = "INSERT INTO chat(envia,recibe,message,hora,sent)VALUES('" & KEY_USUARIO & "','" & Trim(Me.txtDNI.Text) & "','" & Trim(Me.txtmensaje.Text) & "','" & hora & "',now())"
    CnBd.Execute (strCadena)
     
     
    strCadena = "UPDATE  chat SET recd='1' WHERE envia='" & KEY_USUARIO & "' AND recibe='" & Trim(Me.txtDNI.Text) & "' and hora<>'" & hora & "'  "
    CnBd.Execute (strCadena)
     
     
    Call llenarGrid_det(Me.HfdPersona)
    Me.txtmensaje.Text = ""
    Me.txtmensaje.SetFocus
    
End If
End Sub
Public Sub llenarGrid_det(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
strCadena = "SELECT * FROM chat WHERE ((envia='" & KEY_USUARIO & "' AND recibe='" & Trim(Me.txtDNI.Text) & "') OR (envia='" & Trim(Me.txtDNI.Text) & "' AND recibe='" & KEY_USUARIO & "'))  AND  sent =  '" & KEY_FECHA & "' ORDER BY sent ASC "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2800
        Next
        cabecera = "CODIGO" & vbTab & "MENSAJE"
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Dim subjeto As String
            subjeto = rst("envia")
            If subjeto = KEY_USUARIO Then
              subjeto = "YO"
            End If
            Fila = subjeto & vbTab & rst("message")
            Grilla.AddItem Fila
            If rst("envia") = KEY_USUARIO Then
                  '  For k = 0 To 5
                   ' Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H80FF80
                'Next k
            End If
            'Grilla.RowHeight(i + 1) = 450
            
            
            rst.MoveNext
       Next i
       

  
Grilla.Row = Grilla.Rows - 2
Grilla.TopRow = Grilla.Row
Grilla.RowSel = Grilla.Row
'Grilla.col = 0
'Grilla.ColSel = Grilla.Cols - 1
Exit Sub

salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub



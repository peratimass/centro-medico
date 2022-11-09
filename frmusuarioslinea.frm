VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmusuarioslinea 
   BorderStyle     =   0  'None
   Caption         =   "USUARIOS EN LINEA"
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "CERRAR PANTALLA"
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
      MICON           =   "frmusuarioslinea.frx":0000
      PICN            =   "frmusuarioslinea.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtBuscar 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   90
      Width           =   3495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   3615
      Left            =   105
      TabIndex        =   1
      Top             =   420
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   6376
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BUSCAR :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   240
      TabIndex        =   2
      Top             =   90
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   4530
      Left            =   0
      Top             =   0
      Width           =   4710
   End
End
Attribute VB_Name = "frmusuarioslinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Formulario() As String
Dim chat() As String

Private Sub cmdCerrarSesion_Click()
Unload Me
End Sub

Private Sub ChameleonBtn1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Ancho As Double
Dim largo As Double
Ancho = Me.Width
largo = Me.Height

X = Screen.Width
Y = Screen.Height
Me.Left = X - Ancho - 150
Me.Top = Y - largo - 2200
strCadena = "select o.`id_gigane`,p.`nombre_completo`,'PRINCIPAL' as descripcion from gig_usuarios_online o inner join persona p ON (o.`id_gigane`=p.`dni`) WHERE o.ruc='" & KEY_RUC & "' and id_gigane<>'" & KEY_USUARIO & "'"
'strCadena = "SELECT O.id_gigane,P.nombre_completo,A.descripcion FROM gig_usuarios_online O,persona P LEFT JOIN almacen A ON P.dni=A.dni_save WHERE A.ruc=O.ruc AND O.id_gigane=P.dni AND  O.ruc='" & KEY_RUC & "' AND id_gigane<>'" & KEY_USUARIO & "' ORDER BY P.nombre_completo"
Call llenar_online(Me.HfdPersona)

End Sub
Public Sub llenar_online(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim tTotal As Double

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim Formulario(1 To rst.RecordCount)
       ReDim chat(1 To rst.RecordCount)
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 2500
            Grilla.ColWidth(2) = 1700
        Next
        
        
        
        cabecera = "DNI" & vbTab & "USUARIO" & vbTab & "AREA"
        Grilla.AddItem cabecera
         For k = 1 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_gigane") & vbTab & Mid(UCase(rst("nombre_completo")), 1, 25) & vbTab & Mid(UCase(rst("descripcion")), 1, 35)
            Grilla.AddItem Fila
            Fila = ""
            Formulario(i + 1) = rst("id_gigane")
            chat(i + 1) = "0"
            rst.MoveNext
        Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
  Exit Sub
salir:    MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Private Sub HfdPersona_DblClick()
    Dim Nuevo_Form As Form
    
    If Me.HfdPersona.Row > 0 Then
    strCadena = "SELECT * FROM chat_usuario WHERE emisor='" & KEY_USUARIO & "' AND receptor='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    strCadena = "DELETE FROM chat_usuario WHERE emisor='" & KEY_USUARIO & "' AND receptor='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
         
    If rstT.RecordCount < 1 Then
        
        strCadena = "INSERT INTO chat_usuario(emisor,receptor,ruc) VALUES ('" & KEY_USUARIO & "','" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
         
        
        If Not IsFormLoaded(Frmchat, Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))) Then
        
            Set Nuevo_Form = New Frmchat
            Nuevo_Form.txtDni.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
            Nuevo_Form.Caption = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
            Nuevo_Form.Show
            Frmchat.llenar
            
            'Frmchat.Show
            'Frmchat.txtDni.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
            'Frmchat.Caption = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
            'Frmchat.llenar
        End If
   
    End If
    End If
End Sub



Public Function IsFormLoaded(fForm As Form, pDni As String) As Boolean
'On Error GoTo Err_Proc

Dim X As Integer

For X = 0 To Forms.Count - 1

If (Forms(X) Is fForm) Then
  Dim frmaux As Frmchat
  Set frmaux = Forms(X)
  
  
  If frmaux.txtDni.Text = pDni Then
  
    IsFormLoaded = True

   
        
    frmaux.Show
          
    
                
    
    
    
    Exit Function
  End If
  
End If

Next X

 IsFormLoaded = False




End Function


Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'strCadena = "SELECT O.id_gigane,P.nombre_completo FROM gig_usuarios_online O,persona P WHERE O.id_gigane=P.dni AND  O.ruc='" & KEY_RUC & "' AND id_gigane<>'" & KEY_USUARIO & "' AND P.nombre_completo LIKE '%" & Trim(Me.TxtBuscar.Text) & "%'"
    strCadena = "SELECT O.id_gigane,P.nombre_completo,A.descripcion FROM gig_usuarios_online O,persona P LEFT JOIN almacen A ON P.dni=A.dni_save WHERE A.ruc=O.ruc AND O.id_gigane=P.dni AND  O.ruc='" & KEY_RUC & "' AND id_gigane<>'" & KEY_USUARIO & "'AND P.nombre_completo LIKE '%" & Trim(Me.txtBuscar.Text) & "%' ORDER BY P.nombre_completo"
    Call llenar_online(Me.HfdPersona)

End If
End Sub


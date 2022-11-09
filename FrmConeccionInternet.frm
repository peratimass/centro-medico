VERSION 5.00
Begin VB.Form FrmConeccionInternet 
   Caption         =   "Form1"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frame_coneccion 
      Caption         =   "Conexión"
      Height          =   1695
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   8055
      Begin VB.TextBox servidor 
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Text            =   "www.gigane.com"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox base_de_datos 
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Text            =   "molirey"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox usuario 
         Height          =   285
         Left            =   4080
         TabIndex        =   9
         Text            =   "giganec"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox password 
         Height          =   285
         Left            =   6000
         TabIndex        =   8
         Text            =   "ganec3654"
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton boton_conectar 
         Caption         =   "Conectar"
         Default         =   -1  'True
         Height          =   375
         Left            =   6480
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton boton_desconectar 
         Caption         =   "Desconectar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Servidor MySQL"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Index           =   1
         Left            =   4080
         TabIndex        =   14
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Base de Datos"
         Height          =   195
         Index           =   2
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña"
         Height          =   195
         Index           =   3
         Left            =   6000
         TabIndex        =   12
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.Frame frame_base_de_datos 
      Caption         =   "Base de Datos"
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   8055
      Begin VB.ListBox tablas 
         Height          =   3180
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.ListBox contenido 
         Height          =   3180
         Left            =   2160
         TabIndex        =   1
         Top             =   600
         Width           =   5655
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Tablas"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Etiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Contenido"
         Height          =   195
         Index           =   5
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmConeccionInternet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' WWW.ELGURUPROGRAMADOR.COM.AR

Option Explicit

Public WithEvents db As rdoConnection
Attribute db.VB_VarHelpID = -1

Private Sub boton_conectar_Click()
    Dim cadena_conexion
    
    cadena_conexion = "DRIVER={MySQL ODBC 3.51 Driver}; SERVER=" & servidor & "; DATABASE=" & base_de_datos & " ;PWD=" & password & "; UID=" & usuario & ";OPTION=3"
    
    Set db = New rdoConnection
    
    db.CONNECT = cadena_conexion
    db.CursorDriver = rdUseServer
    db.EstablishConnection
End Sub

Private Sub boton_desconectar_Click()
    db.Close
End Sub

Private Sub db_Connect(ByVal ErrorOccurred As Boolean)
    Dim Tabla As rdoTable
    Dim hay_tablas As Boolean
    
    hay_tablas = False
    cambiar_botones True
    
    For Each Tabla In db.rdoTables
        tablas.AddItem Tabla.name
        hay_tablas = True
    Next
    
    If Not hay_tablas Then
        MsgBox "La base de datos esta vacia"
        boton_desconectar_Click
    End If
End Sub

Private Sub db_Disconnect()
    cambiar_botones False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If boton_conectar.Enabled = False Then
        db.Close
    End If
End Sub

Private Sub cambiar_botones(conectado As Boolean)

    contenido.Clear
    tablas.Clear
    
    boton_conectar.Enabled = Not conectado
    boton_desconectar.Enabled = conectado
    tablas.Enabled = conectado
    contenido.Enabled = conectado
    
End Sub

Private Sub tablas_Click()
    Dim Tabla As String
    Dim consulta As New rdoQuery
    Dim resultados As rdoResultset
    Dim contenido_row As String
    Dim columna As rdoColumn
    
    contenido.Clear
    
    Tabla = tablas.List(tablas.ListIndex)
    
    Set consulta.ActiveConnection = db
    
    consulta.sql = "SELECT * FROM " & Tabla & " WHERE 1"
    consulta.Execute
    
    Set resultados = consulta.OpenResultset
    
    While Not resultados.EOF
        
        contenido_row = ""
        
        For Each columna In resultados.rdoColumns
            contenido_row = contenido_row & columna.name & "=" & resultados(columna.name) & "; "
        Next
        
        contenido.AddItem contenido_row
        resultados.MoveNext
    Wend
        
    resultados.Close
    Set resultados = Nothing
    
End Sub


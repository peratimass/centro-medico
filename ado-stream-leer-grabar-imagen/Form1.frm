VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Leer y grabar imagenes con ADO Stream"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6495
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2566
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2 - Para cambiarla doble clic"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "1 - Para ver la imagen Clic en una fila"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2625
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   6495
      Begin VB.PictureBox Picture1 
         Height          =   3735
         Left            =   120
         ScaleHeight     =   3675
         ScaleWidth      =   6195
         TabIndex        =   1
         Top             =   240
         Width           =   6255
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public cn As Connection
Dim Rst As ADODB.Recordset











Private Sub Form_Load()

Dim Path_BD As String
    
    ' Nuevo objeto Connection
    Set cn = New Connection
        
        ' Ruta de la bd
      '  Path_BD = App.Path & "\bd1.mdb"
        
        ' Cadena de conexión
        'cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                              "Data Source=" & Path_BD & ";" & _
                              "Mode=ReadWrite|Share Deny None;" & _
                              "Persist Security Info=False;Jet OLEDB"
'cn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=base1"
 sys_Server = "localhost"
 sys_DataBase = "base1" 'ConfigRead("DataBase")
 sys_SUser = "user1" 'DecryptString(ConfigRead("SUser"))
 sys_SPassword = "1020304050" 'DecryptString(ConfigRead("SPassword"))
 sys_ConString = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server & ";" & _
            "Database=" & sys_DataBase & ";" & _
            "UID=" & sys_SUser & ";" & _
            "PWD=" & sys_SPassword & ";" & _
            "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
        
    
    cn.ConnectionString = sys_ConString
    cn.Open
        
    
        
        Set Rst = New ADODB.Recordset
        
        Rst.Open "select Id,Nombre from table1", cn, adOpenStatic, adLockOptimistic
        
        Set MSHFlexGrid1.DataSource = Rst
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    ' Cierra la conexión ado
    If cn.State = adStateOpen Then
       cn.Close
    End If
    
    If cn Is Nothing Then
       Set cn = Nothing
    End If

End Sub

Private Sub MSHFlexGrid1_Click()

Dim sql As String

sql = "select Foto From table1 Where id=" & CInt(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 0))

Picture1.Picture = Leer_Imagen(cn, sql, "Foto")
End Sub

Private Sub MSHFlexGrid1_DblClick()
Dim ret As Boolean
Dim id As Variant
Dim Path_Imagen As String
    
    ' pide la ruta del gráfico
   ' Path_Imagen = InputBox("Ruta del gráfico a guardar en el campo " & _
                            "FOTO en la base de datos bd1.mdb")
                            
    Path_Imagen = "C:\Documents and Settings\Percy\Mis documentos\Mis imágenes\Invierno.jpg"
    ' pide el ID
    id = InputBox("Id del registro para cambiar la imagen " & _
                            " en el campo FOTO en la base de datos bd1.mdb")
    
    
    
    If Path_Imagen = vbNullString Or id = vbNullString Then Exit Sub
    
    ' Le pasa la conexión , el comando sql, el nombre del campo _
      y el path del archivo
      strCadena = "SELECT  foto FROM table1 WHERE id='" & id & "'"
    ret = Guardar_Imagen(cn, strCadena, "foto", Path_Imagen)
    
    
    
    If ret Then
        MsgBox "Imagen guardada", vbInformation
    End If
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmTemporal 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTemporal.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ARREGLAR OBSERVACION PRODUCTOS"
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   2880
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "INGRESAR PRODUCTOS POR ALMACEN"
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   2400
      Width           =   3495
   End
   Begin VB.TextBox TxtCampoInt 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txttabla 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox TxtCampo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ARREGLAR CODIGO"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Campo:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   1260
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tabla:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   300
      Width           =   465
   End
   Begin VB.Label LblDescripcion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Campo:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   435
      TabIndex        =   2
      Top             =   780
      Width           =   555
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   1500
      Left            =   120
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "FrmTemporal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Cadena As String
Dim x As Integer
Dim CodFor As New ADODB.Recordset

Cadena = "SELECT cunidad FROM Unidad ORDER BY intUnidad ASC"
CodFor.Open Cadena, CnBd, adOpenKeyset, adLockOptimistic

StrCadena = "SELECT cunidad FROM Unidad ORDER BY intUnidad ASC"
Call ConfiguraRst(StrCadena)
x = Rst.RecordCount
Rst.MoveFirst
CodFor.MoveFirst
For x = 0 To Rst.RecordCount - 1
    Rst.Fields(0) = Formatos(Str(CodFor.Fields(0)))
     
    Rst.Update
    If Rst.EOF = False Then
        Rst.MoveNext
        CodFor.MoveNext
    Else
        MsgBox "Se Ingresaron todos los datos"
    End If
Next x
End Sub

Function Formatos(ByVal cod As String) As String
Dim x As Integer
Dim Formato As String
  Formato = ""
  For x = 1 To 4
    Formato = Formato + "0"
  Next x
    cod = Format(Trim(Str(Val(Right(cod, 4)))), Formato)

Formatos = cod

End Function

Private Sub Command2_Click()
Dim rstProducto As New ADODB.Recordset
Dim rstAlmacen As New ADODB.Recordset
Dim rstAlmacen_Producto As New ADODB.Recordset
Dim habilitar As String * 1
Dim CadenaSQL As String
Dim i As Integer, j As Integer

CadenaSQL = "SELECT Alm_cod FROM Almacen ORDER BY int_almacen ASC"
rstAlmacen.Open CadenaSQL, CnBd, adOpenKeyset, adLockOptimistic
rstAlmacen.MoveFirst

CadenaSQL = "SELECT cProducto FROM Producto ORDER BY intProducto ASC"
rstProducto.Open CadenaSQL, CnBd, adOpenKeyset, adLockOptimistic
rstProducto.MoveFirst

CadenaSQL = "SELECT *FROM Almacen_Productos "
rstAlmacen_Producto.Open CadenaSQL, CnBd, adOpenKeyset, adLockOptimistic

'---------Agregar un Nuevo registro----------------------

For i = 0 To rstAlmacen.RecordCount - 1
    For j = 0 To rstProducto.RecordCount - 1
        rstAlmacen_Producto.AddNew
        rstAlmacen_Producto(0) = rstAlmacen(0)
        rstAlmacen_Producto(1) = rstProducto(0)
        rstAlmacen_Producto(2) = 0
        rstAlmacen_Producto(3) = "V"
        rstAlmacen_Producto.Update
        If rstProducto.EOF = False Then
            rstProducto.MoveNext
        Else
            GoTo 10
        End If
    Next j
10:
        If rstAlmacen.EOF = False Then
            rstAlmacen.MoveNext
            rstProducto.MoveFirst
        Else
            Exit Sub
        End If
Next i
End Sub

Private Sub Command3_Click()
Dim i As Integer
Dim rstObservacion As New ADODB.Recordset
StrCadena = "SELECT sObservacion FROM Producto"
Call ConfiguraRst(StrCadena)
Rst.MoveFirst
For i = 0 To Rst.RecordCount - 1
    Rst(0) = ""
    Rst.Update
    Rst.MoveNext
Next i
End Sub

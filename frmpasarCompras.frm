VERSION 5.00
Begin VB.Form frmpasarCompras 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   13515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "compras"
      Height          =   1215
      Left            =   2880
      TabIndex        =   0
      Top             =   840
      Width           =   5415
   End
End
Attribute VB_Name = "frmpasarCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim rstP As New ADODB.Recordset
StrCadena = "SELECT * FROM Producto"
rstP.Open StrCadena, CnBd, adOpenKeyset
rstP.MoveFirst
For i = 0 To rstP.RecordCount - 1
    StrCadena = "SELECT * FROM Almacen_Productos WHERE cProducto='" & rstP("cProducto") & "' AND Alm_cod='0001'"
    Call ConfiguraRst(StrCadena)
    If rst.RecordCount < 1 Then
        StrCadena = "INSERT Almacen_Productos VALUES('0001','" & rstP("cProducto") & "','0','" & CVDate(Date) & "','V')"
        CnBd.Execute (StrCadena)
    End If
    Set rst = Nothing
    rstP.MoveNext
Next i
End Sub

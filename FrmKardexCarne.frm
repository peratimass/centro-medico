VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmKardexCarne 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kardex carne"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   12726
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " * * * * S A L D O S * * * *"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   10965
      TabIndex        =   3
      Top             =   120
      Width           =   2565
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " * * * *E G R E S O S * * * *"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   7455
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* * * * *I N G R E S O S * * * *"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4035
      TabIndex        =   1
      Top             =   120
      Width           =   2925
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   14895
   End
End
Attribute VB_Name = "FrmKardexCarne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub Form_Load()
CenterForm Me
Me.Top = 500

    stcodigop = Trim(FrmDerivados.DtcCombo.BoundText)

strCadena = "SELECT     Kardex.FechaEmision,Kardex.FechaEmision as Emision, (Comprobantes.doc_abrev + ':' + Kardex.sSerie + '-' + Kardex.NumeroDoc) AS Comprobante, Kardex.Ing_Cant as Cant, " & _
"                      Kardex.Precio, Kardex.Ing_Cant * Kardex.Precio AS Total, Kardex.Sal_Cant as Cant, Kardex.Precio AS Precio," & _
"                      Kardex.Sal_Cant * Kardex.Precio AS Total, Kardex.Stk_Gen as Cant, Kardex.Precio,Kardex.Stk_Gen* Kardex.Precio as Total,Kardex.cProducto " & _
"FROM         Kardex INNER JOIN Comprobantes ON Kardex.doc_cod = Comprobantes.doc_cod " & _
"WHERE (Kardex.cProducto='" & Trim(stcodigop) & "' AND Kardex.alm_Cod ='" & Trim(FrmDerivados.DtcAlmacen.BoundText) & "')   ORDER BY cKardex DESC "
Call ConfiguraRst(strCadena)
        
        Call llenarGrid(Me.HfdDetalle, Me)
        
        
End Sub


Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal formulario As Form)
On Error GoTo SALIR
'  Grilla.Clear
  
 '  Grilla.Rows = Rst.RecordCount
   
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 0
  Grilla.ColWidth(1) = 1000
  Grilla.ColAlignment(1) = 7
  Grilla.ColWidth(2) = 2400
  Grilla.ColAlignment(2) = 7
  Grilla.ColWidth(3) = 1200
  Grilla.ColAlignment(3) = 7
  Grilla.ColWidth(4) = 1200
  Grilla.ColAlignment(4) = 7
  Grilla.ColWidth(5) = 1200
  Grilla.ColAlignment(5) = 7
  Grilla.ColWidth(6) = 1200
  Grilla.ColWidth(7) = 1200
  Grilla.ColWidth(8) = 1200
  Grilla.ColWidth(9) = 1200
  Grilla.ColWidth(10) = 1200
  Grilla.ColWidth(11) = 1200
  Grilla.ColAlignment(6) = 7
  Grilla.ColAlignment(7) = 7
  Grilla.ColAlignment(8) = 7
  Grilla.ColAlignment(9) = 7
  Grilla.ColAlignment(10) = 7
  Grilla.ColAlignment(11) = 7
  Grilla.ColWidth(12) = 0
  
  
Call DarFormatoFecha(Grilla, 1)
'Call DarFormato(Grilla, 3)
'Call DarFormato(Grilla, 4)
'Call DarFormato(Grilla, 5)
'Call DarFormato(Grilla, 6)
'Call DarFormato(Grilla, 7)
'Call DarFormato(Grilla, 8)
'Call DarFormato(Grilla, 9)
'Call DarFormato(Grilla, 10)
Call DarFormato(Grilla, 11)

Set rst = Nothing
For i = 3 To 5
    For j = 1 To Grilla.Rows - 1
        Grilla.col = i
        Grilla.Row = j
        Grilla.CellBackColor = &HC0FFC0
    Next j
  Next i
  
  For i = 6 To 8
    For j = 1 To Grilla.Rows - 1
        Grilla.col = i
        Grilla.Row = j
        Grilla.CellBackColor = &H80FFFF
    Next j
  Next i
  For i = 9 To 11
    For j = 1 To Grilla.Rows - 1
        Grilla.col = i
        Grilla.Row = j
        Grilla.CellBackColor = &H8080FF
    Next j
  Next i
  
  Exit Sub
SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub








Private Sub HfdDetalle_DblClick()
Procedencia = buscar
FrmProductoMermas.Show
Unload Me
End Sub

VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmPermisos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Permisos"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   5985
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
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
End
Attribute VB_Name = "FrmPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub Form_Load()
CenterForm Me
Me.Top = 200
Dim rstP As New ADODB.Recordset
Dim rstU As New ADODB.Recordset
If FrmUsuarios.Procedencia = Selecionar Then
    strCadena = "SELECT    Usuario_permisos.idUsuario, Usuario_permisos.id_menu, Menu.nombre, Usuario_permisos.estado " & _
    "FROM         Usuario_permisos INNER JOIN Menu ON Usuario_permisos.id_menu = Menu.id_menu " & _
    " WHERE idUsuario='" & Trim(FrmUsuarios.HfdGrilla.TextMatrix(FrmUsuarios.HfdGrilla.Row, 0)) & "'"
    Call ConfiguraRst(strCadena)
    Me.HfdDetalle.Clear
    Me.HfdDetalle.Rows = 1
    Set Me.HfdDetalle.Recordset = rst
    Me.HfdDetalle.Rows = rst.RecordCount + 1
    Me.HfdDetalle.ColWidth(0) = 0
    Me.HfdDetalle.ColWidth(1) = 500
    Me.HfdDetalle.ColWidth(2) = 2500
    Me.HfdDetalle.ColWidth(3) = 400
    Set rst = Nothing
End If
End Sub

Private Sub HfdDetalle_DblClick()
Procedencia = Selecionar
FrmdarPermiso.Show
End Sub

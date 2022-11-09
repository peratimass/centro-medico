VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmSublineas 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo DtcLinea 
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   6960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtLinea 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   7440
      Width           =   1935
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgLinea 
      Height          =   6255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   11033
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VitekeySoft.ChameleonBtn cmdEliminar 
      Height          =   945
      Left            =   7320
      TabIndex        =   6
      Top             =   2490
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1667
      BTYPE           =   5
      TX              =   "ELIMINAR"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSublineas.frx":0000
      PICN            =   "FrmSublineas.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   945
      Left            =   7320
      TabIndex        =   7
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1667
      BTYPE           =   5
      TX              =   "NUEVO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSublineas.frx":2466
      PICN            =   "FrmSublineas.frx":2482
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdModificar 
      Height          =   1020
      Left            =   7320
      TabIndex        =   8
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1799
      BTYPE           =   5
      TX              =   "MODIFICAR"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSublineas.frx":28D4
      PICN            =   "FrmSublineas.frx":28F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdSalir 
      Height          =   945
      Left            =   7320
      TabIndex        =   9
      Top             =   3480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1667
      BTYPE           =   5
      TX              =   "SALIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSublineas.frx":4F29
      PICN            =   "FrmSublineas.frx":4F45
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUB-LINEA :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   7440
      Width           =   795
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LINEA :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   330
      TabIndex        =   3
      Top             =   7080
      Width           =   495
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUB FAMILIAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   285
      TabIndex        =   2
      Top             =   120
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8145
      Left            =   0
      Top             =   0
      Width           =   8460
   End
End
Attribute VB_Name = "FrmSublineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub cmdEliminar_Click()
If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        strCadena = "SELECT * FROM producto WHERE id_sublinea='" & Trim(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            MsgBox "Imposible Eliminar Este Item, esta Asociado a un producto Registrado", vbInformation, "No Puede ELiminar este Registro"
            Exit Sub
        Else
        strCadena = "DELETE FROM linea_sub WHERE id_tipo='" & Trim(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) & "' and id_usu='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        strCadena = "SELECT S.id_tipo,L.descripcion as linea,S.descripcion as sublinea  FROM linea L,linea_sub S WHERE L.id_linea=S.id_linea AND L.id_usu='" & KEY_RUC & "' AND S.id_usu='" & KEY_RUC & "' LIMIT 50"
        Call actualizar
        End If
    End If
End Sub

Private Sub cmdmodificar_Click()
 
 Procedencia = modificar
 FrmDetalleSublinea.Show
 
End Sub

Private Sub cmdNuevo_Click()
 
 Procedencia = nuevo
 Call disabled_form(Me)
 Call enabled_form(FrmDetalleSublinea)
 FrmDetalleSublinea.Show
 
End Sub

Private Sub cmdSalir_Click()



Unload Me

End Sub

Private Sub DtcLinea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT S.id_tipo,L.descripcion as linea,S.descripcion as sublinea  FROM linea L,linea_sub S WHERE L.id_linea=S.id_linea AND L.id_usu=S.id_usu AND L.id_linea='" & Me.DtcLinea.BoundText & "' and L.id_usu='" & KEY_RUC & "'"
    Call actualizar
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
strCadena = "SELECT id_linea as Codigo,descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcLinea)
strCadena = "SELECT S.id_tipo,L.descripcion as linea,S.descripcion as sublinea  FROM linea L,linea_sub S WHERE L.id_linea=S.id_linea AND L.id_usu='" & KEY_RUC & "' AND S.id_usu='" & KEY_RUC & "' LIMIT 50"
Call actualizar

End Sub
Public Sub actualizar()

Call llenarGrid(Me.HfgLinea)
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Call ConfiguraRst(strCadena)

 
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
  
  N = 1
  
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 2800
           Grilla.ColWidth(2) = 2800
                      
       Next
         cabecera = "CODIGO" & vbTab & "LINEA" & vbTab & "SUBLINEA"
         Grilla.AddItem cabecera
         For k = 0 To 2
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = Fila & rst("id_tipo") & vbTab & UCase(rst("linea")) & vbTab & UCase(rst("sublinea"))
             Grilla.AddItem Fila
             Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub HfgLinea_SelChange()

If Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) > 0 Then
    Me.cmdEliminar.Enabled = True
    Me.cmdModificar.Enabled = True
Else
    Me.cmdEliminar.Enabled = False
    Me.cmdModificar.Enabled = False
End If

End Sub

Private Sub TxtLinea_Change()
    strCadena = "SELECT S.id_tipo,L.descripcion as linea,S.descripcion as sublinea  FROM linea L,linea_sub S WHERE L.id_linea=S.id_linea AND L.id_usu='" & KEY_RUC & "' and S.descripcion LIKE '%" & Trim(Me.TxtLinea.Text) & "%' AND S.id_usu='" & KEY_RUC & "' LIMIT 50"
    Call actualizar

End Sub








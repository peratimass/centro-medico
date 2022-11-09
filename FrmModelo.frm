VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmModelo 
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
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox txtIdModelo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   2040
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Top             =   1680
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo DtcLinea 
         Height          =   330
         Left            =   2040
         TabIndex        =   10
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcSubLinea 
         Height          =   330
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VitekeySoft.ChameleonBtn cmdProcesar 
         Height          =   825
         Left            =   2400
         TabIndex        =   13
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1455
         BTYPE           =   5
         TX              =   "PROCESAR"
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
         MICON           =   "FrmModelo.frx":0000
         PICN            =   "FrmModelo.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdCancelar 
         Height          =   825
         Left            =   3960
         TabIndex        =   14
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1455
         BTYPE           =   5
         TX              =   "CANCELAR"
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
         MICON           =   "FrmModelo.frx":3664
         PICN            =   "FrmModelo.frx":3680
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCION :"
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
         Left            =   780
         TabIndex        =   9
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUB-CLASIFICACION :"
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
         TabIndex        =   8
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLASIFICACION :"
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
         Left            =   660
         TabIndex        =   7
         Top             =   720
         Width           =   1125
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfModelo 
      Height          =   6975
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   12303
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
      TabIndex        =   2
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
      MICON           =   "FrmModelo.frx":399A
      PICN            =   "FrmModelo.frx":39B6
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
      TabIndex        =   3
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
      MICON           =   "FrmModelo.frx":5E00
      PICN            =   "FrmModelo.frx":5E1C
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
      TabIndex        =   4
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
      MICON           =   "FrmModelo.frx":626E
      PICN            =   "FrmModelo.frx":628A
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
      TabIndex        =   5
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
      MICON           =   "FrmModelo.frx":88C3
      PICN            =   "FrmModelo.frx":88DF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MODELOS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   195
      TabIndex        =   1
      Top             =   120
      Width           =   1065
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
Attribute VB_Name = "FrmModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub ChameleonBtn1_Click()

End Sub

Private Sub cmdCancelar_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub cmdEliminar_Click()

    Procedencia = Eliminar
    Call disabled_form(Me)
    frmsegurity.Show
    Exit Sub

End Sub

Private Sub cmdModificar_Click()


strCadena = "call ADM_Sublinea_Modelo('C','" & Val(Me.HfModelo.TextMatrix(Me.HfModelo.Row, 0)) & "','" & Me.DtcSubLinea.BoundText & "','" & Me.DtcLinea.BoundText & "','" & Replace(UCase(Me.txtDescripcion.Text), "'", "") & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtIdModelo.Text = Me.HfModelo.TextMatrix(Me.HfModelo.Row, 0)
   Me.DtcLinea.BoundText = rst("id_linea")
   Me.DtcSubLinea.BoundText = rst("id_sublinea")
   Me.txtDescripcion.Text = rst("descripcion")
   Me.frmdetalle.Visible = True
End If



End Sub

Private Sub cmdNuevo_Click()
    
    
    Me.txtDescripcion.Text = ""
    Me.frmdetalle.Visible = True
    
    
    
End Sub

Private Sub cmdProcesar_Click()

If validar = True Then
    strCadena = "call ADM_Sublinea_Modelo('I','" & Val(Me.txtIdModelo.Text) & "','" & Me.DtcSubLinea.BoundText & "','" & Me.DtcLinea.BoundText & "','" & Replace(UCase(Me.txtDescripcion.Text), "'", "") & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Me.frmdetalle.Visible = False
    Call actualizar(Me.HfModelo)
End If


End Sub
Private Function validar() As Boolean
If Me.txtDescripcion.Text = "" Then
    MsgBox "Ingrese una Descripcion Valida", vbInformation
    validar = False
    Exit Function
End If
If Me.DtcLinea.BoundText = "" Then
   MsgBox "Ingrese una Linea Valida", vbInformation
   validar = False
   Exit Function
End If

If Me.DtcSubLinea.BoundText = "" Then
   MsgBox "Ingrese una SubLinea Valida", vbInformation
   validar = False
   Exit Function
End If
validar = True

End Function
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub DtcLinea_Change()
strCadena = "SELECT id_tipo as Codigo,descripcion as Descripcion FROM linea_sub WHERE id_linea='" & Me.DtcLinea.BoundText & "' and  id_usu='" & KEY_RUC & "' ORDER BY descripcion"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcSubLinea)
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100

strCadena = "SELECT id_linea as Codigo,descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcLinea)

strCadena = "SELECT id_tipo as Codigo,descripcion as Descripcion FROM linea_sub WHERE id_linea='" & Me.DtcLinea.BoundText & "' and  id_usu='" & KEY_RUC & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcSubLinea)





Call actualizar(Me.HfModelo)
End Sub

Private Sub HfModelo_SelChange()
If Val(Me.HfModelo.TextMatrix(Me.HfModelo.Row, 0)) > 0 Then
   Me.cmdEliminar.Enabled = True
   Me.cmdModificar.Enabled = True
Else
    Me.cmdEliminar.Enabled = False
    Me.cmdModificar.Enabled = False
End If
End Sub

Public Sub actualizar(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
strCadena = "call ADM_Sublinea_Modelo('L','','','','','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 1900
           Grilla.ColWidth(2) = 1900
           Grilla.ColWidth(3) = 1900
           
         Next
        cabecera = "CODIGO" & vbTab & "LINEA" & vbTab & "SUBLINEA" & vbTab & "MODELO"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = Format(rst("Id"), "00000") & vbTab & rst("linea") & vbTab & UCase(rst("sublinea")) & vbTab & rst("modelo")
            Grilla.AddItem Fila
            rst.MoveNext
    Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub



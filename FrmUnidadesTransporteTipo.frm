VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmUnidadesTransporteTipo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   10335
      Begin VB.TextBox txtid_tipo_veiculo 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6120
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtprecio 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtdescripcion 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   480
         Width           =   4095
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   495
         Left            =   1800
         TabIndex        =   13
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUnidadesTransporteTipo.frx":0000
         PICN            =   "FrmUnidadesTransporteTipo.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdsalir2 
         Height          =   495
         Left            =   3360
         TabIndex        =   14
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "CERRAR"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmUnidadesTransporteTipo.frx":3664
         PICN            =   "FrmUnidadesTransporteTipo.frx":3680
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO :"
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
         Left            =   810
         TabIndex        =   10
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label Label1 
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
         Left            =   600
         TabIndex        =   9
         Top             =   600
         Width           =   1005
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   975
      Left            =   10680
      TabIndex        =   4
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmUnidadesTransporteTipo.frx":66A7
      PICN            =   "FrmUnidadesTransporteTipo.frx":66C3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   6120
      Width           =   3015
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgLinea 
      Height          =   5295
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9340
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   8388608
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
   Begin VitekeySoft.ChameleonBtn cmdmodificar 
      Height          =   975
      Left            =   10680
      TabIndex        =   5
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmUnidadesTransporteTipo.frx":6B15
      PICN            =   "FrmUnidadesTransporteTipo.frx":6B31
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdeliminar 
      Height          =   975
      Left            =   10680
      TabIndex        =   6
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmUnidadesTransporteTipo.frx":6E4B
      PICN            =   "FrmUnidadesTransporteTipo.frx":6E67
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   975
      Left            =   10680
      TabIndex        =   7
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmUnidadesTransporteTipo.frx":92B1
      PICN            =   "FrmUnidadesTransporteTipo.frx":92CD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblFecha 
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
      Left            =   375
      TabIndex        =   3
      Top             =   6120
      Width           =   1005
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLASIFICACION DE TRANSPORTE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   120
      Width           =   2205
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   240
      Top             =   5880
      Width           =   10335
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   6735
      Left            =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "FrmUnidadesTransporteTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub cmdEliminar_Click()
 If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        
        strCadena = "DELETE FROM transporte_tipo WHERE id_tipo_transporte='" & Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        Call actualizar
End If
End Sub

Private Sub cmdmodificar_Click()
Me.txtid_tipo_veiculo.Text = Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)
Me.TxtDescripcion.Text = Trim(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 1))
Me.txtprecio.Text = Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 2))
Me.frmdetalle.Visible = True
End Sub

Private Sub cmdnuevo_Click()

Me.frmdetalle.Visible = True
Me.TxtDescripcion.Text = ""
Me.txtprecio.Text = 0
Me.txtid_tipo_veiculo.Text = 0


Call Resalta(Me.txtprecio)
      
      
End Sub

Private Sub cmdprocesar_Click()
If Trim(Me.TxtDescripcion.Text) <> "" Then
    If Val(Me.txtid_tipo_veiculo.Text) < 1 Then
        Me.txtid_tipo_veiculo.Text = get_correlativo_table("transporte_tipo", "id_tipo_transporte")
    End If
    strCadena = "call p_put_unidad_transporte('" & Trim(Me.txtid_tipo_veiculo.Text) & "','" & Trim(Me.TxtDescripcion.Text) & "','" & Val(Me.txtprecio.Text) & "','" & KEY_RUC & "')"
    Call Execute_Sql(strCadena)
    Me.frmdetalle.Visible = False
    Call actualizar
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub cmdsalir2_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
Call actualizar

End Sub
Public Sub actualizar()

strCadena = "SELECT * FROM transporte_tipo WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion"
Call llenarGrid(Me.HfgLinea, Me)
End Sub

Private Sub HfgLinea_SelChange()
If Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) > 0 Then
    Me.cmdmodificar.Enabled = True
    Me.cmdeliminar.Enabled = True
Else
    Me.cmdmodificar.Enabled = False
    Me.cmdeliminar.Enabled = False
End If
End Sub

Private Sub TxtLinea_Change()
strCadena = "SELECT id_tipo_transporte ,descripcion FROM transporte_tipo WHERE descripcion LIKE '%" & Trim(Me.TxtLinea.Text) & "%' AND ruc='" & KEY_RUC & "' ORDER BY Descripcion ASC"
Call llenarGrid(Me.HfgLinea, Me)
End Sub


Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.rows = 0
   
    Exit Sub

End If
 
 
  
   Grilla.rows = 0
        ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 1300
            Grilla.ColWidth(1) = 4500
            Grilla.ColWidth(2) = 2500
          Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "PRECIO"
         Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id_tipo_transporte") & vbTab & rst("descripcion") & vbTab & Format(rst("precio"), "###0.00")
             Grilla.AddItem Fila
          
            rst.MoveNext
        Next i
        
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
        Me.cmdeliminar.Enabled = True
        Me.cmdmodificar.Enabled = True
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub





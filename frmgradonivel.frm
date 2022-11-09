VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmgradonivel 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      Height          =   7455
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   18735
      Begin VB.TextBox TxtDocente 
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
         Height          =   315
         Left            =   6360
         TabIndex        =   27
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox txtvacantes 
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
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtgrado 
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
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Top             =   1560
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo DtcNivel 
         Height          =   330
         Left            =   1800
         TabIndex        =   11
         Top             =   960
         Width           =   4575
         _ExtentX        =   8070
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
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   615
         Left            =   1800
         TabIndex        =   13
         Top             =   5160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "PROCESAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmgradonivel.frx":0000
         PICN            =   "frmgradonivel.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdsalir 
         Height          =   615
         Left            =   4200
         TabIndex        =   14
         Top             =   5160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "SALIR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmgradonivel.frx":3664
         PICN            =   "frmgradonivel.frx":3680
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcPeriodo 
         Height          =   330
         Left            =   1800
         TabIndex        =   15
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
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
      Begin MSDataListLib.DataCombo DtcMatricula 
         Height          =   330
         Left            =   1800
         TabIndex        =   21
         Top             =   2880
         Width           =   4455
         _ExtentX        =   7858
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
      Begin MSDataListLib.DataCombo DtcPension 
         Height          =   330
         Left            =   1800
         TabIndex        =   22
         Top             =   3480
         Width           =   4455
         _ExtentX        =   7858
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
      Begin MSDataListLib.DataCombo DtcDocente 
         Height          =   330
         Left            =   1800
         TabIndex        =   25
         Top             =   4080
         Width           =   4455
         _ExtentX        =   7858
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
      Begin VB.Label lblid_nivel 
         Height          =   375
         Left            =   9360
         TabIndex        =   28
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOCENTE ENCARGADO:"
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
         Left            =   135
         TabIndex        =   26
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PENSION MENSUAL:"
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
         Left            =   240
         TabIndex        =   24
         Top             =   3600
         Width           =   1365
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MATRICULA :"
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
         Left            =   720
         TabIndex        =   23
         Top             =   3000
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VACANTES :"
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
         TabIndex        =   18
         Top             =   2160
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIODO :"
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
         Left            =   900
         TabIndex        =   16
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRADO :"
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
         Left            =   1020
         TabIndex        =   10
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL :"
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
         Left            =   1110
         TabIndex        =   9
         Top             =   1080
         Width           =   495
      End
   End
   Begin MSDataListLib.DataCombo DtcPeriodoBusqueda 
      Height          =   330
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgLinea 
      Height          =   7455
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   13150
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
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
   Begin VitekeySoft.ChameleonBtn cmdexit 
      Height          =   855
      Left            =   19080
      TabIndex        =   4
      Top             =   3255
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmgradonivel.frx":3A70
      PICN            =   "frmgradonivel.frx":3A8C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmddelete 
      Height          =   855
      Left            =   19080
      TabIndex        =   5
      Top             =   2370
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "ELIMINAR"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmgradonivel.frx":3E7C
      PICN            =   "frmgradonivel.frx":3E98
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdupdate 
      Height          =   855
      Left            =   19080
      TabIndex        =   6
      Top             =   1485
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MODIFICAR"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmgradonivel.frx":62E2
      PICN            =   "frmgradonivel.frx":62FE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   19080
      TabIndex        =   7
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmgradonivel.frx":6618
      PICN            =   "frmgradonivel.frx":6634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcNivelAcademico 
      Height          =   330
      Left            =   1860
      TabIndex        =   19
      Top             =   8160
      Width           =   3975
      _ExtentX        =   7011
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NIVEL ACADEMICO :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   240
      TabIndex        =   20
      Top             =   8160
      Width           =   1515
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VACANTES"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   405
      TabIndex        =   3
      Top             =   120
      Width           =   885
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODO :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   1980
      TabIndex        =   2
      Top             =   120
      Width           =   795
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8655
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "frmgradonivel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdnuevo_Click()
Me.txtgrado.Text = 0
Call Resalta(Me.txtgrado)
Me.frmdetalle.Visible = True
Me.DtcMatricula.BoundText = 0
Me.DtcPension.BoundText = 0

End Sub

Private Sub cmdprocesar_Click()

strCadena = "CALL p_put_nivel_grado_ii('" & Val(Me.lblid_nivel.Caption) & "','" & Me.DtcNivel.BoundText & "','" & Val(Me.DtcPeriodo.BoundText) & "','" & Val(Me.txtvacantes.Text) & "','" & Trim(Me.txtgrado.Text) & "','" & Me.DtcMatricula.BoundText & "','" & Me.DtcPension.BoundText & "','" & Me.DtcDocente.BoundText & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

Me.frmdetalle.Visible = False
strCadena = "SELECT * FROM  view_nivel_grado_ii  WHERE id_periodo='" & Me.DtcPeriodoBusqueda.BoundText & "' and id_nivel='" & Me.DtcNivelAcademico.BoundText & "' and  ruc='" & KEY_RUC & "'"
Call llenarGrid(Me.HfgLinea)

End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1500
           Grilla.ColWidth(2) = 2500
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 3200
           Grilla.ColWidth(5) = 3200
           Grilla.ColWidth(6) = 1300
           Grilla.ColWidth(7) = 1300
           Grilla.ColWidth(8) = 1300
           Grilla.ColWidth(9) = 3000
        Next
        cabecera = "CODIGO" & vbTab & "NIVEL" & vbTab & "PERIODO" & vbTab & "GRADO" & vbTab & "MATRICULA" & vbTab & "PENSION" & vbTab & "VACANTES" & vbTab & "MATRICULADOS" & vbTab & "RESTANTES" & vbTab & "DOCENTE"
        Grilla.AddItem cabecera
         For k = 1 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
         Next k
                            
        rst.MoveFirst
        in_vacantes = 0
        in_matriculados = 0
        For i = 0 To rst.RecordCount - 1
            Fila = Format(rst("id_grado"), "0000") & vbTab & rst("nivel") & vbTab & rst("periodo") & vbTab & rst("descripcion") & vbTab & rst("matricula") & vbTab & rst("pension") & vbTab & rst("vacantes") & vbTab & rst("matriculados") & vbTab & rst("vacantes") - rst("matriculados") & vbTab & rst("docente")
            Grilla.AddItem Fila
            in_matriculados = in_matriculados + rst("matriculados")
            in_vacantes = in_vacantes + rst("vacantes")
            
            rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "TOTALES ::" & vbTab & in_vacantes & vbTab & in_matriculados & vbTab & in_vacantes - in_matriculados & vbTab & ""
        Grilla.AddItem Fila
        For k = 1 To 9
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HDFDFE0
         Next k
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub

Private Sub cmdsalir_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub cmdupdate_Click()
If Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) > 0 Then
    strCadena = "SELECT * FROM view_nivel_grado_ii WHERE id_grado='" & Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.lblid_nivel.Caption = rst("id_grado")
       Me.DtcPeriodo.BoundText = rst("id_periodo")
       Me.DtcNivel.BoundText = rst("id_nivel")
       Me.txtgrado.Text = rst("descripcion")
       Me.txtvacantes.Text = rst("vacantes")
       Me.DtcMatricula.BoundText = rst("id_matricula")
       Me.DtcPension.BoundText = rst("id_pension")
       Me.DtcDocente.BoundText = rst("dni_docente")
       Me.frmdetalle.Visible = True
    Else
       Me.lblid_nivel.Caption = 0
    End If
End If
End Sub



Private Sub DtcNivelAcademico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM  view_nivel_grado_ii WHERE id_nivel='" & Me.DtcNivelAcademico.BoundText & "' and  id_periodo='" & Me.DtcPeriodoBusqueda.BoundText & "' and  ruc='" & KEY_RUC & "'"
    Call llenarGrid(Me.HfgLinea)
End If
End Sub
Private Function get_periodo_anio() As String

strCadena = "SELECT id_periodo FROM college_periodo WHERE id_anio='" & Year(KEY_FECHA) & "' and   ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_periodo_anio = rstL("id_periodo")
End If
End Function

Private Sub Form_Load()
CenterForm Me
Me.Top = 100

strCadena = "SELECT id_periodo as Codigo,descripcion as Descripcion FROM college_periodo where ruc='" & KEY_RUC & "' ORDER By id_periodo"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPeriodo)



strCadena = "SELECT id_periodo as Codigo,descripcion as Descripcion FROM college_periodo where ruc='" & KEY_RUC & "' ORDER By id_periodo"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPeriodoBusqueda)
Me.DtcPeriodoBusqueda.BoundText = get_periodo_anio


  

strCadena = "SELECT id_nivel as Codigo, descripcion as Descripcion FROM nivel_educativo "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcNivel)

strCadena = "SELECT id_nivel as Codigo, descripcion as Descripcion FROM nivel_educativo "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcNivelAcademico)


strCadena = "SELECT id_producto as Codigo,CONCAT(nombre_prod,' [ S/. ',precio_venta,' ]') as Descripcion FROM view_producto WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMatricula)

strCadena = "SELECT id_producto as Codigo,CONCAT(nombre_prod,' [ S/. ',precio_venta,' ]') as Descripcion FROM view_producto WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPension)

strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_personal='si'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcDocente)



strCadena = "SELECT * FROM  view_nivel_grado_ii WHERE  id_periodo='" & Me.DtcPeriodoBusqueda.BoundText & "' and  ruc='" & KEY_RUC & "'"
Call llenarGrid(Me.HfgLinea)
End Sub

Private Sub TxtDocente_Change()
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_personal='si' and nombre_completo LIKE '%" & Trim(Me.TxtDocente.Text) & "%'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcDocente)
End Sub

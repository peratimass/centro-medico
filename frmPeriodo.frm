VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmPeriodo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   8535
      Begin VB.TextBox txtid_periodo 
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Text            =   "0"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
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
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   1560
         Width           =   3855
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   615
         Left            =   1680
         TabIndex        =   2
         Top             =   2160
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
         MICON           =   "frmPeriodo.frx":0000
         PICN            =   "frmPeriodo.frx":001C
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
         Left            =   3840
         TabIndex        =   3
         Top             =   2160
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
         MICON           =   "frmPeriodo.frx":3664
         PICN            =   "frmPeriodo.frx":3680
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcMes 
         Height          =   330
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
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
      Begin MSDataListLib.DataCombo DtcAnio 
         Height          =   330
         Left            =   1680
         TabIndex        =   13
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AÑO :"
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
         Left            =   1155
         TabIndex        =   7
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCION :"
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
         Left            =   495
         TabIndex        =   6
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MES :"
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
         Left            =   1185
         TabIndex        =   5
         Top             =   360
         Width           =   435
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPeriodo 
      Height          =   6255
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   11033
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
      Left            =   9000
      TabIndex        =   9
      Top             =   2295
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
         Weight          =   700
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
      MICON           =   "frmPeriodo.frx":3A70
      PICN            =   "frmPeriodo.frx":3A8C
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
      Left            =   9000
      TabIndex        =   10
      Top             =   1365
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
         Weight          =   700
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
      MICON           =   "frmPeriodo.frx":3E7C
      PICN            =   "frmPeriodo.frx":3E98
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
      Left            =   9000
      TabIndex        =   11
      Top             =   480
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
         Weight          =   700
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
      MICON           =   "frmPeriodo.frx":41B2
      PICN            =   "frmPeriodo.frx":41CE
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
      Caption         =   "PERIODO ESCOLAR"
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
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   7665
      Left            =   0
      Top             =   0
      Width           =   10140
   End
End
Attribute VB_Name = "frmPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
Dim in_precio As String
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
  
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 4500
           Grilla.ColWidth(2) = 2000
           
    Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "OPERADOR"
         Grilla.AddItem cabecera
         For k = 0 To 2
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
           
              
             
       
            
            Fila = rst("id_periodo") & vbTab & rst("descripcion") & vbTab & get_persona(rst("dni_save"))
            Grilla.AddItem Fila
            
        
        rst.MoveNext
        Next i
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Public Sub actualizar()
strCadena = "SELECT * FROM college_periodo WHERE ruc='" & KEY_RUC & "'"
Call Me.llenarGrid(Me.HfPeriodo)
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub
Sub nuevo()
Me.txtid_periodo.Text = ""
Me.DtcMes.BoundText = Month(KEY_FECHA)
Me.DtcAnio.BoundText = Year(KEY_FECHA)
Me.txtdescripcion.Text = "AÑO ESCOLAR" & Space(1) & Year(KEY_FECHA)
Me.frmdetalle.Visible = True

End Sub

Private Sub cmdnuevo_Click()
Call nuevo
End Sub

Private Sub cmdprocesar_Click()
If Trim(Me.txtdescripcion.Text) <> "" Then
    
    strCadena = "call put_periodo_college('" & Val(Me.txtid_periodo.Text) & "','" & Me.DtcMes.BoundText & "','" & Me.DtcAnio.BoundText & "','" & UCase(Trim(Me.txtdescripcion.Text)) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Me.frmdetalle.Visible = False
    Call actualizar
End If
End Sub

Private Sub cmdsalir_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100

strCadena = "SELECT id_mes as Codigo,descripcion as Descripcion FROM mes"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMes)
Me.DtcMes.BoundText = Month(KEY_FECHA)

strCadena = "SELECT anio as Codigo,anio as Descripcion FROM anio ORDER BY anio ASC "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAnio)
Me.DtcAnio.BoundText = Year(KEY_FECHA)


Call actualizar
End Sub

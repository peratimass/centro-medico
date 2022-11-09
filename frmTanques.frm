VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmTanques 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   960
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox txtMinimo 
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
         Left            =   2400
         TabIndex        =   14
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtMaxima 
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
         Left            =   2400
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
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
         Left            =   2400
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtid_tanque 
         Height          =   375
         Left            =   4440
         TabIndex        =   1
         Text            =   "0"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   615
         Left            =   2400
         TabIndex        =   3
         Top             =   2280
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
         MICON           =   "frmTanques.frx":0000
         PICN            =   "frmTanques.frx":001C
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
         TabIndex        =   4
         Top             =   2280
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
         MICON           =   "frmTanques.frx":3664
         PICN            =   "frmTanques.frx":3680
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
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
         Caption         =   "LECTURA MINIMA :"
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
         Left            =   615
         TabIndex        =   13
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LECTURA MÁXIMA :"
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
         Left            =   660
         TabIndex        =   6
         Top             =   1200
         Width           =   1485
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
         Left            =   1095
         TabIndex        =   5
         Top             =   480
         Width           =   1125
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfTanques 
      Height          =   6255
      Left            =   240
      TabIndex        =   7
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
      TabIndex        =   8
      Top             =   2280
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
      MICON           =   "frmTanques.frx":3A70
      PICN            =   "frmTanques.frx":3A8C
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
      TabIndex        =   9
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
      MICON           =   "frmTanques.frx":3E7C
      PICN            =   "frmTanques.frx":3E98
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
      TabIndex        =   10
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
      MICON           =   "frmTanques.frx":41B2
      PICN            =   "frmTanques.frx":41CE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   7665
      Left            =   0
      Top             =   0
      Width           =   10140
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
      TabIndex        =   11
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmTanques"
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
           Grilla.ColWidth(1) = 3500
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
    Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "MINIMA" & vbTab & "MAXIMA"
         Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
            Fila = rst("id_tanque") & vbTab & rst("descripcion") & vbTab & rst("minimo") & vbTab & rst("maxima")
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Public Sub actualizar()
strCadena = "SELECT * FROM tanque WHERE ruc='" & KEY_RUC & "'"
Call Me.llenarGrid(Me.HfTanques)
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub
Sub nuevo()
Me.txtid_tanque.Text = ""
Me.txtdescripcion.Text = ""
Me.frmdetalle.Visible = True

End Sub

Private Sub cmdNuevo_Click()
Call nuevo
End Sub

Private Sub cmdprocesar_Click()
If Trim(Me.txtdescripcion.Text) <> "" Then
    
    strCadena = "call put_tanque('" & Val(Me.txtid_tanque.Text) & "','" & Trim(Me.txtdescripcion.Text) & "','" & Val(Me.txtMinimo.Text) & "','" & Val(Me.txtMaxima.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Me.frmdetalle.Visible = False
    Call actualizar
End If
End Sub

Private Sub modificar(ByVal in_tanque As String)
strCadena = "SELECT * FROM tanque WHERE id_tanque ='" & Val(in_tanque) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtdescripcion.Text = rst("descripcion")
   Me.txtMaxima.Text = rst("maxima")
   Me.txtMinimo.Text = rst("minimo")
   Me.txtid_tanque.Text = rst("id_tanque")
   Me.frmdetalle.Visible = True
End If


End Sub


Private Sub cmdsalir_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub cmdupdate_Click()
If Val(Me.HfTanques.TextMatrix(Me.HfTanques.Row, 0)) > 0 Then
   Call modificar(Me.HfTanques.TextMatrix(Me.HfTanques.Row, 0))
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100




Call actualizar
End Sub


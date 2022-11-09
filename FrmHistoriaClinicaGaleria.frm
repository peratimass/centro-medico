VERSION 5.00
Begin VB.Form FrmHistoriaClinicaGaleria 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   19845
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "ANTERIOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   8760
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "DESCRIPCION RESULTADOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8535
      Left            =   13200
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      Begin VB.TextBox TxtResultado 
         Appearance      =   0  'Flat
         Height          =   3615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   4680
         Width           =   6255
      End
      Begin VB.Label LblBarra 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "3 of 9 Barcode"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   675
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   4245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO BARRA :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label lbldireccion 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CODIGO EXAMEN :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   675
         Left            =   2040
         TabIndex        =   15
         Top             =   3480
         Width           =   3885
      End
      Begin VB.Image Image2 
         Height          =   1575
         Left            =   240
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lbllaboratorio 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CODIGO EXAMEN :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2040
         TabIndex        =   14
         Top             =   3120
         Width           =   3885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LABORATORIO"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   13
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         Height          =   1935
         Left            =   120
         Top             =   2640
         Width           =   6015
      End
      Begin VB.Label lbltratante 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CODIGO EXAMEN :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Top             =   2040
         Width           =   4245
      End
      Begin VB.Label lblfecha 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CODIGO EXAMEN :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Top             =   1680
         Width           =   4245
      End
      Begin VB.Label LblDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CODIGO EXAMEN :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         TabIndex        =   10
         Top             =   1320
         Width           =   4245
      End
      Begin VB.Label lblCodExamen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CODIGO EXAMEN :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   960
         Width           =   4245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DR. TRATANTE :"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "FECHA :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   960
         TabIndex        =   7
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPCION :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   420
         TabIndex        =   6
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO EXAMEN :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   960
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17400
      TabIndex        =   1
      Top             =   8760
      Width           =   2295
   End
   Begin VB.CommandButton cmdsiguiente 
      Caption         =   "SIGUIENTE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   8760
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   8535
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   12975
   End
End
Attribute VB_Name = "FrmHistoriaClinicaGaleria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAnterior_Click()
If rst.EOF = True Or rst.BOF = True Then
    rst.MoveFirst
Else
    rst.MovePrevious
    If rst.EOF = True Or rst.BOF = True Then
        rst.MoveFirst
    End If
End If
Call llenar_resultado
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub cmdsiguiente_Click()
If rst.EOF = True Or rst.BOF = True Then
    rst.MoveFirst
Else
    rst.MoveNext
    If rst.EOF = True Or rst.BOF = True Then
        rst.MoveFirst
    End If
End If
Call llenar_resultado
End Sub


Private Sub llenar_resultado()
If IsNull(rst("imagen")) = False And Len(rst("imagen")) > 2 Then
    If VerificarFichero(App.Path & "\archivos\" & rst("dni")) = True Then
         Me.Image1.Picture = LoadPicture(App.Path + "\archivos\" + rst("dni") + "\" + Trim(rst("imagen")))
         strCadena = "SELECT PA.resultado,PA.estado FROM persona_analisis PA WHERE PA.id='" & FrmHistoriaClinicaConsulta.HfgExamenes.TextMatrix(FrmHistoriaClinicaConsulta.HfgExamenes.Row, 0) & "'"
         Call ConfiguraRstT(strCadena)
         Me.TxtResultado.text = rstT("resultado")
     End If
End If
End Sub



Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Dim fotol As String
strCadena = "SELECT C.dni_save,A.id_analisis,A.ruc_empresa,A.fecha FROM persona_consultas C,persona_analisis A WHERE C.dni_save=A.dni_save AND A.id='" & FrmHistoriaClinicaConsulta.HfgExamenes.TextMatrix(FrmHistoriaClinicaConsulta.HfgExamenes.Row, 0) & "'"
Call ConfiguraRstT(strCadena)
Me.LblBarra.Caption = rstT("id_analisis")
Me.lblCodExamen.Caption = rstT("id_analisis")
Me.LblDescripcion.Caption = BDBuscarCampo("analisis_clinico_listado", "descripcion", "id_analisis", rstT("id_analisis"))
Me.lblfecha.Caption = rstT("fecha")
Me.lbltratante.Caption = BDBuscarCampo("persona", "nombre_completo", "dni", rstT("dni_save"))
Me.lbllaboratorio.Caption = UCase(BDBuscarCampo("persona", "nombre_completo", "dni", rstT("ruc_empresa")))
Me.lbldireccion.Caption = UCase(BDBuscarCampo("persona", "direccion", "dni", rstT("ruc_empresa")))
fotol = BDBuscarCampo("persona", "foto", "dni", rstT("ruc_empresa"))
fotol = "tmb_" & fotol
If IsNull(fotol) = False And Len(fotol) > 2 Then
    If VerificarFichero(App.Path & "\archivos\" & rstT("ruc_empresa")) = True Then
      Me.Image2.Picture = LoadPicture(App.Path + "\archivos\" + rstT("ruc_empresa") + "\" + Trim(fotol))
         End If
Else
    Me.Image2 = Nothing
End If

strCadena = "SELECT * FROM persona_fotos_examen WHERE id_examen='" & FrmHistoriaClinicaConsulta.HfgExamenes.TextMatrix(FrmHistoriaClinicaConsulta.HfgExamenes.Row, 0) & "'"
Call ConfiguraRst(strCadena)
rst.MoveFirst
'--------- foto--------
Call llenar_resultado
'--------- foto--------

End Sub

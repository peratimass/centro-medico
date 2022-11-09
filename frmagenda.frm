VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmAgenda 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmagenda 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   6360
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   11295
      Begin VB.TextBox txtObservacion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Left            =   2280
         TabIndex        =   23
         Top             =   2880
         Width           =   8775
      End
      Begin VB.TextBox txtMinutos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2280
         TabIndex        =   17
         Text            =   "30"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txthora_ini 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2280
         TabIndex        =   16
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtBusqueda 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   9580
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo dtcEspecialista 
         Height          =   330
         Left            =   2280
         TabIndex        =   12
         Top             =   1080
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtdni 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2280
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VitekeySoft.ChameleonBtn cmdAgendar 
         Height          =   855
         Left            =   4320
         TabIndex        =   18
         Top             =   3840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "AGENDAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmagenda.frx":0000
         PICN            =   "frmagenda.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdFacturar 
         Height          =   855
         Left            =   6360
         TabIndex        =   19
         Top             =   3840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "FACTURAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmagenda.frx":0336
         PICN            =   "frmagenda.frx":0352
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00E0E0E0&
         BorderWidth     =   2
         Height          =   4935
         Left            =   0
         Top             =   0
         Width           =   11295
      End
      Begin VB.Label lblidAgenda 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OPERADOR :"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         TabIndex        =   26
         Top             =   3720
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lbloperador 
         BackColor       =   &H00E0E0E0&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7200
         TabIndex        =   25
         Top             =   2160
         Width           =   3840
      End
      Begin VB.Label lblcomprobante 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7200
         TabIndex        =   24
         Top             =   1680
         Width           =   3840
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DETALLE :"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   555
         Left            =   600
         TabIndex        =   22
         Top             =   2880
         Width           =   1560
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OPERADOR :"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5280
         TabIndex        =   21
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "COMPROBANTE :"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   5280
         TabIndex        =   20
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   10920
         Picture         =   "frmagenda.frx":2C3C
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "MINUTOS  :"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   600
         TabIndex        =   15
         Top             =   2280
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "HORA INICIO  :"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   480
         TabIndex        =   14
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ESPECIALISTA :"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   480
         TabIndex        =   11
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label lblcliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3840
         TabIndex        =   10
         Top             =   480
         Width           =   7200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "DNI CLIENTE :"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   960
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         Height          =   1095
         Left            =   5160
         Shape           =   4  'Rounded Rectangle
         Top             =   1560
         Width           =   6015
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   1095
      Left            =   18960
      TabIndex        =   3
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      BTYPE           =   5
      TX              =   "NUEVO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmagenda.frx":5AE0
      PICN            =   "frmagenda.frx":5AFC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   4770
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   8414
      _Version        =   393216
      ForeColor       =   8388608
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   18
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   171966465
      CurrentDate     =   44858
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   8175
      Left            =   5280
      TabIndex        =   0
      Top             =   840
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   14420
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   -2147483635
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VitekeySoft.ChameleonBtn cmdVisualizar 
      Height          =   1095
      Left            =   18960
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      BTYPE           =   5
      TX              =   "VISUALIZAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmagenda.frx":5F4E
      PICN            =   "frmagenda.frx":5F6A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdAnular 
      Height          =   1095
      Left            =   18960
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      BTYPE           =   5
      TX              =   "ANULAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmagenda.frx":9240
      PICN            =   "frmagenda.frx":925C
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
      Height          =   1095
      Left            =   18960
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      BTYPE           =   5
      TX              =   "SALIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmagenda.frx":B6A6
      PICN            =   "frmagenda.frx":B6C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "LISTADO DE AGENDADOS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub nuevo()
Me.txtdni.Text = ""
Me.lblcliente.Caption = ""
Me.txtBusqueda.Text = ""

strCadena = "CALL procedure_agenda('1','','','','','','','','','','','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcEspecialista)

strCadena = "CALL procedure_agenda('6','2','3','" & Me.dtcEspecialista.BoundText & "','5','" & Format(Me.MonthView1.Value, "YYYY-mm-dd") & "','7','8','9','','','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txthora_ini.Text = Format(rst("hora_fin"), "HH:mm")
Else
    Me.txthora_ini.Text = Format("08:00", "HH:mm")
End If
Me.txtMinutos.Text = 30

Me.cmdAgendar.Enabled = True
Me.frmagenda.Visible = True
Call Resalta(Me.txtdni)

End Sub
Public Sub put_anular(ByVal in_agenda As String)

strCadena = "CALL procedure_agenda('8','" & Val(in_agenda) & "','','','','','','','','','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Call Me.llenar_agenda(Me.HfdDetalle)


End Sub
Public Sub get_agenda(ByVal in_agenda As String)
strCadena = "CALL procedure_agenda('5','" & Val(in_agenda) & "','','','','','','','','','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.lblidAgenda.Caption = rst("id")
    Me.txtdni.Text = rst("dni_cliente")
    Me.txtObservacion.Text = rst(9)
    Me.lblcliente.Caption = get_persona(rst("dni_cliente"))
    Me.dtcEspecialista.BoundText = rst("dni_medico")
    Me.txthora_ini.Text = Format(rst("hora_ini"), "HH:mm")
    Me.txtMinutos.Text = Format(rst("hora_fin"), "HH:mm")
   
    Me.lblcomprobante.Caption = rst("comprobante")
    Me.lbloperador.Caption = get_persona(rst("dni_save"))
    Me.cmdAgendar.Enabled = False
    If rst("id_venta") > 0 Then
        Me.cmdFacturar.Enabled = False
    Else
        Me.cmdFacturar.Enabled = True
    End If
    Me.frmagenda.Visible = True
End If


End Sub
Public Sub agendar()
Dim in_minutos As String
in_minutos = "00:" + Format(Me.txtMinutos.Text, "00")
If Trim(Me.txtdni.Text) <> "" And Trim(Me.txthora_ini.Text) <> "" And Val(Me.txtMinutos.Text) > 0 Then
    strCadena = "CALL procedure_agenda('4','','" & Trim(Me.txtdni.Text) & "','" & Me.dtcEspecialista.BoundText & "','','" & Format(Me.MonthView1.Value, "YYYY-mm-dd") & "','" & Format(Trim(Me.txthora_ini.Text), "HH:ss") & "','','','','','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        strCadena = "CALL procedure_agenda('2','','" & Trim(Me.txtdni.Text) & "','" & Me.dtcEspecialista.BoundText & "','" & rst("id_especialidad") & "','" & Format(Me.MonthView1.Value, "YYYY-mm-dd") & "','" & Format(Trim(Me.txthora_ini.Text), "HH:mm") & "','" & Format(in_minutos, "HH:mm") & "','','" & Trim(Me.txtObservacion.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        Me.lblidAgenda.Caption = rst(0)
        Call Me.llenar_agenda(Me.HfdDetalle)
        
        
    End If
    
    
    
Else
    MsgBox "INGRESE DATOS CORRECTOS", vbInformation, KEY_VENDEDOR
    Call Resalta(Me.txtdni)
End If
End Sub

Public Sub llenar_agenda(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim color As String
strCadena = "CALL procedure_agenda('3','','" & Trim(Me.txtdni.Text) & "','','','" & Format(Me.MonthView1.Value, "YYYY-mm-dd") & "','" & Format(Trim(Me.txthora_ini.Text), "HH:ss") & "','" & Format(in_minutos, "HH:ss") & "','','','','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
           Grilla.Rows = 0
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 800
           Grilla.ColWidth(4) = 1000
           Grilla.ColWidth(5) = 2800
           Grilla.ColWidth(6) = 2000
           Grilla.ColWidth(7) = 2400
           Grilla.ColWidth(8) = 2400
      
         cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "HORA INI" & vbTab & "HORA FIN" & vbTab & "DNI CLIENTE" & vbTab & "PACIENTE" & vbTab & "ESPECIALIDAD" & vbTab & "MEDICO" & vbTab & "COMPROBANTE"
         Grilla.AddItem cabecera
         For k = 1 To 8
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            
            
             Fila = rst("id") & vbTab & Format(rst("fecha"), "dd-mm-YYYY") & vbTab & Format(rst("hora_ini"), "HH:mm") & vbTab & Format(rst("hora_fin"), "HH:mm") & vbTab & rst("dni_cliente") & vbTab & rst("nombre_completo") & vbTab & rst("especialidad") & vbTab & rst("medico") & vbTab & rst("comprobante")
             Grilla.AddItem Fila
             
            
             
             If Val(rst("id_estado")) = 3 Then
                For k = 1 To 8
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
             End If
             
            
         Grilla.RowHeight(i + 1) = 400
        rst.MoveNext
        Next i
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub



Private Sub ChameleonBtn2_Click()

End Sub

Private Sub cmdAgendar_Click()

Call agendar
Me.cmdAgendar.Enabled = False

End Sub

Private Sub cmdAnular_Click()

Procedencia = anular
Call disabled_form(Me)
frmsegurity.Show
Exit Sub
End Sub

Private Sub cmdFacturar_Click()

Call FrmVentas.facturar_agenda(Me.lblidAgenda.Caption)
Me.frmagenda.Visible = False
End Sub



Private Sub cmdNuevo_Click()
Call nuevo
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdVisualizar_Click()
Call get_agenda(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
strCadena = "CALL procedure_agenda('1','','','','','','','','','','','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcEspecialista)
Me.MonthView1.Value = KEY_FECHA
Call llenar_agenda(Me.HfdDetalle)




End Sub

Private Sub Image1_Click()
Me.frmagenda.Visible = False
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Call llenar_agenda(Me.HfdDetalle)
End Sub

Private Sub txtdni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
buscar_nuevamente:
    strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtdni.Text) & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.txtdni.Text = rst("dni")
        Me.lblcliente.Caption = rst("nombre_completo")
    Else
        If get_dni_reniec_iii(Trim(Me.txtdni.Text), KEY_DEPARTAMENTO, KEY_PROVINCIA, KEY_DISTRITO) = True Then
             GoTo buscar_nuevamente
        End If
    End If
    
End If
End Sub

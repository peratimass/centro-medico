VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FrmHistoriaClinicaVideo 
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
      Height          =   9135
      Left            =   13200
      TabIndex        =   2
      Top             =   0
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
         TabIndex        =   16
         Top             =   960
         Width           =   1365
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
         TabIndex        =   15
         Top             =   1440
         Width           =   1140
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
         TabIndex        =   14
         Top             =   1800
         Width           =   600
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
         TabIndex        =   13
         Top             =   2160
         Width           =   1200
      End
      Begin VB.Label lblCodExamen 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0107"
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
         Top             =   960
         Width           =   4245
      End
      Begin VB.Label LblDescripcion 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ECOCARDIOGRAMA"
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
         Top             =   1320
         Width           =   4245
      End
      Begin VB.Label lblfecha 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "03/08/2012"
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
         Top             =   1680
         Width           =   4245
      End
      Begin VB.Label lbltratante 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "------"
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
         Top             =   2040
         Width           =   4245
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
         TabIndex        =   8
         Top             =   2640
         Width           =   1095
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
         TabIndex        =   7
         Top             =   3120
         Width           =   3885
      End
      Begin VB.Image Image2 
         Height          =   1575
         Left            =   240
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   1455
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
         TabIndex        =   6
         Top             =   3480
         Width           =   3885
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
         TabIndex        =   5
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label LblBarra 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0203040"
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
         TabIndex        =   4
         Top             =   240
         Width           =   4245
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         Height          =   1935
         Left            =   120
         Top             =   2640
         Width           =   6015
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   9015
      Left            =   120
      ScaleHeight     =   8955
      ScaleWidth      =   12795
      TabIndex        =   1
      Top             =   120
      Width           =   12855
   End
   Begin MCI.MMControl MMControl1 
      Height          =   1695
      Left            =   14280
      TabIndex        =   0
      Top             =   9360
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   2990
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
End
Attribute VB_Name = "FrmHistoriaClinicaVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.top = 50
MMControl1.Visible = False
MMControl1.Notify = False
MMControl1.Wait = True
MMControl1.Shareable = False
MMControl1.DeviceType = "AVIVideo"
MMControl1.hWndDisplay = Picture1.hWnd
MMControl1.Filename = "c:\DAEWON.avi"

MMControl1.Command = "Open"
MMControl1.Command = "Play"
Picture1.AutoSize = True
If MMControl1.Error <> 0 Then MsgBox MMControl1.ErrorMessage
End Sub

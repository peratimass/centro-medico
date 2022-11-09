VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmconsultadua 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   16575
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8415
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   16335
      ExtentX         =   28813
      ExtentY         =   14843
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VitekeySoft.ChameleonBtn cmdSalir 
      Height          =   300
      Left            =   14280
      TabIndex        =   0
      Top             =   8640
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   529
      BTYPE           =   5
      TX              =   "CERRAR PANTALLA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmconsultadua.frx":0000
      PICN            =   "frmconsultadua.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   9015
      Left            =   0
      Top             =   0
      Width           =   16575
   End
End
Attribute VB_Name = "frmconsultadua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsalir_Click()
Unload Me

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 150
Me.WebBrowser1.Navigate ("http://www.sunat.gob.pe/servlet/SgDetSerie?codaduana=118&numecorre=494534&anoprese=2014&n=10&regimen=10&fini=2014&fechingsi=16/12/2014&ordemb=%20&tipodocdecla=4&docdecla=20479779598&codubigeo=%20&ndcl=%20&mcaduregpre=&mfanoregpre=&mcodiregpre=&mndclregpre=&mnserregpre=%20&numeorden=001641&tipodespacho=%3Cb%3E%20NORMAL%3C/b%3E&tipoaforo=D&legajada=&mod=1&serie=1ACTIVO")
End Sub

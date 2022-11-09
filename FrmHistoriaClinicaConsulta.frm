VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmHistoriaClinicaConsulta 
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
   Begin VB.Frame Frame9 
      Caption         =   "DATOS DEL PACIENTE"
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
      Height          =   4095
      Left            =   15240
      TabIndex        =   40
      Top             =   3840
      Width           =   4455
      Begin VB.Image imgpaciente 
         Height          =   3495
         Left            =   600
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3255
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   120
      TabIndex        =   38
      Top             =   6120
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "DIAGNOSTICO APRIORI"
      TabPicture(0)   =   "FrmHistoriaClinicaConsulta.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image2"
      Tab(0).Control(1)=   "TxtDiagnosticoApriori"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "DIAGNOSTICO FINAL"
      TabPicture(1)   =   "FrmHistoriaClinicaConsulta.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtDiagnostico"
      Tab(1).Control(1)=   "Image1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "MEDICACION"
      TabPicture(2)   =   "FrmHistoriaClinicaConsulta.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "HfgReceta"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "TRATAMIENTO"
      TabPicture(3)   =   "FrmHistoriaClinicaConsulta.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Image3"
      Tab(3).Control(1)=   "TxtTratamiento"
      Tab(3).ControlCount=   2
      Begin VB.TextBox TxtDiagnostico 
         Appearance      =   0  'Flat
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
         Height          =   1965
         Left            =   -73200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   900
         Width           =   12735
      End
      Begin VB.TextBox TxtTratamiento 
         Appearance      =   0  'Flat
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
         Height          =   1965
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   900
         Width           =   12735
      End
      Begin VB.TextBox TxtDiagnosticoApriori 
         Appearance      =   0  'Flat
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
         Height          =   1965
         Left            =   -73200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   900
         Width           =   12735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgReceta 
         Height          =   2175
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   3836
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Image Image1 
         Height          =   2370
         Left            =   -74880
         Picture         =   "FrmHistoriaClinicaConsulta.frx":0070
         Stretch         =   -1  'True
         Top             =   540
         Width           =   1110
      End
      Begin VB.Image Image3 
         Height          =   2370
         Left            =   -61800
         Picture         =   "FrmHistoriaClinicaConsulta.frx":D637
         Stretch         =   -1  'True
         Top             =   540
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   2370
         Left            =   -74880
         Picture         =   "FrmHistoriaClinicaConsulta.frx":1969E
         Stretch         =   -1  'True
         Top             =   540
         Width           =   1110
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "RESULTADO DEL EXAMEN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   8880
      TabIndex        =   36
      Top             =   3720
      Width           =   6255
      Begin VB.TextBox TXtResultado 
         Appearance      =   0  'Flat
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
         Height          =   1965
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "EXAMENES A REALIZAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   120
      TabIndex        =   30
      Top             =   3720
      Width           =   8655
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgExamenes 
         Height          =   1935
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3413
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   1905
         Left            =   7605
         TabIndex        =   34
         Top             =   240
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   3360
         BandCount       =   1
         ForeColor       =   8388608
         ImageList       =   "ImgIconos"
         FixedOrder      =   -1  'True
         VariantHeight   =   0   'False
         Orientation     =   1
         EmbossPicture   =   -1  'True
         _CBWidth        =   900
         _CBHeight       =   1905
         _Version        =   "6.0.8169"
         Child1          =   "TlbAcciones"
         MinHeight1      =   840
         Width1          =   3180
         FixedBackground1=   0   'False
         NewRow1         =   0   'False
         Begin MSComctlLib.Toolbar TlbAcciones 
            Height          =   2340
            Left            =   30
            TabIndex        =   35
            Top             =   30
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   4128
            ButtonWidth     =   1429
            ButtonHeight    =   1376
            Style           =   1
            ImageList       =   "ImgIconos"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Imagenes"
                  Key             =   "(Nuevo)"
                  Object.ToolTipText     =   "Nuevo"
                  ImageKey        =   "(Fotos)"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   " Videos"
                  Key             =   "(Modificar)"
                  Object.ToolTipText     =   "Modificar"
                  ImageKey        =   "(Videos)"
               EndProperty
            EndProperty
            OLEDropMode     =   1
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "ANAMESIS O INTERROGATORIO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   8880
      TabIndex        =   28
      Top             =   1920
      Width           =   6255
      Begin VB.TextBox TXtAnamesis 
         Appearance      =   0  'Flat
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
         Height          =   1245
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "TOMA DE DATOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   15240
      TabIndex        =   19
      Top             =   120
      Width           =   4455
      Begin VB.TextBox TxtGlucosa 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1680
         TabIndex        =   27
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox TxtTemperatura 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox TxtPulso 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1680
         TabIndex        =   25
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox TXtPeso 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "NIVEL GLUCOSA :"
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
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "TEMPERATURA :"
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
         Left            =   315
         TabIndex        =   22
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "PESO :"
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
         Left            =   1020
         TabIndex        =   21
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "PULSO :"
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
         Left            =   930
         TabIndex        =   20
         Top             =   720
         Width           =   585
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "PRESION ARTERIAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   15240
      TabIndex        =   16
      Top             =   2040
      Width           =   4455
      Begin VB.TextBox TxtDiastolia 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1680
         TabIndex        =   33
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox TXtSistolica 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1680
         TabIndex        =   32
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DIASTOLICA :"
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
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SISTOLICA :"
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
         TabIndex        =   17
         Top             =   480
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "MOTIVO CONSULTA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   8655
      Begin VB.TextBox TxtMotivo 
         Appearance      =   0  'Flat
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
         Height          =   1245
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   360
         Width           =   8175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "DATOS CONSULTA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7095
      Begin VB.TextBox TxtConsulta 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtFecha 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox TxtClinica 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nº CONSULTA :"
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
         TabIndex        =   13
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label6 
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
         Left            =   720
         TabIndex        =   12
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "LUGAR :"
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
         Left            =   720
         TabIndex        =   11
         Top             =   1080
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DATOS TRATANTE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1575
      Left            =   7320
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.TextBox TxtDoctor 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox TxtDni 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox txtCip 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "NºCOLEGIADO"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "NOMBRE :"
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
         Left            =   720
         TabIndex        =   3
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DNI :"
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
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CODIGO MEDICO:"
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
         Left            =   135
         TabIndex        =   1
         Top             =   360
         Width           =   1320
      End
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   7200
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHistoriaClinicaConsulta.frx":26C65
            Key             =   "(Fotos)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmHistoriaClinicaConsulta.frx":2A0C5
            Key             =   "(Videos)"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image4 
      Height          =   915
      Left            =   15480
      Picture         =   "FrmHistoriaClinicaConsulta.frx":2D349
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   3900
   End
End
Attribute VB_Name = "FrmHistoriaClinicaConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Select Case FrmHistoriaClinica.Procedencia
    Case Selecionar
        Call llenar_historia(FrmHistoriaClinica.HfgConsultas.TextMatrix(FrmHistoriaClinica.HfgConsultas.Row, 0), FrmHistoriaClinica.TxtRuc.text)
End Select

End Sub
Private Sub llenar_historia(ByVal id_historia As String, ByVal dni As String)
Dim strFoto As String
strCadena = "SELECT * FROM persona_consultas WHERE id_consulta='" & id_historia & "' AND dni='" & dni & "'"
Call ConfiguraRst(strCadena)
Me.TxtConsulta.text = rst("id_consulta")
Me.txtFecha.text = rst("fecha")
Me.TxtClinica.text = BDBuscarCampo("persona", "nombre_completo", "dni", rst("ruc_clinica"))
Me.TxtDni.text = rst("dni_save")
Me.TxtDoctor.text = BDBuscarCampo("persona", "nombre_completo", "dni", rst("dni_save"))
Me.TxtMotivo.text = UCase(rst("motivo"))
Me.TXtSistolica.text = rst("sistole")
Me.TxtDiastolia.text = rst("diastole")
Me.TXtPeso.text = rst("peso")
Me.TxtPulso.text = rst("pulso")
Me.TxtTemperatura.text = rst("temperatura")
Me.TxtGlucosa.text = rst("glucosa")
Me.TXtAnamesis.text = UCase(rst("interrogatorio"))
Me.TxtDiagnosticoApriori.text = UCase(rst("diagnostico"))
Me.TxtDiagnostico.text = UCase(rst("diagnostico_final"))
Me.TxtTratamiento.text = UCase(rst("tratamiento"))
strCadena = "SELECT PA.id,AC.descripcion,PA.estado FROM persona_analisis PA,analisis_clinico_listado AC WHERE PA.dni='" & dni & "' AND PA.id_consulta='" & id_historia & "' AND PA.id_analisis=AC.id_analisis"
Call llenar_examenes(Me.HfgExamenes)
strCadena = "SELECT R.id,R.id_medic,G.descripcion AS generico,M.descripcion,M.concentracion,F.descripcion as fdes,L.descripcion as ldes,saldo,R.estado FROM persona_receta R,genericos G,medicamento M,forma_farmaceutica F,laboratorio L WHERE R.id_medic=M.id_medic AND M.id_forma=F.id AND M.id_laboratorio=L.id_laboratorio AND R.id_consulta='" & id_historia & "' AND R.dni='" & dni & "' AND M.id_generico=G.id"
Call llenar_receta(Me.HfgReceta)
Me.Frame9.Caption = UCase(FrmHistoriaClinica.txtRazonSocial.text)
    '--------- foto--------
    strFoto = BDBuscarCampo("persona", "foto", "dni", dni)
If IsNull(strFoto) = False And Len(strFoto) > 5 Then
    If VerificarFichero(App.Path & "\archivos\" & dni) = True Then
       Me.imgpaciente.Visible = True
        Me.imgpaciente.Picture = LoadPicture(App.Path + "\archivos\" + dni + "\" + Trim(strFoto))
        img = Trim(strFoto)
    Else
        Me.imgpaciente.Visible = False
    End If
End If
End Sub
Private Sub llenar_receta(ByVal Grilla As MSHFlexGrid)
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 500
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 3500
           Grilla.ColWidth(3) = 2000
           Grilla.ColWidth(4) = 4000
           Grilla.ColWidth(5) = 1100
        Next
        cabecera = "ID" & vbTab & "NOM. GENERICO" & vbTab & "NOM. COMERCIAL" & vbTab & "PRESENTACION" & vbTab & "LABORATORIO" & vbTab & "CANTIDAD"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rstT.RecordCount - 1
          Fila = rstT("id_medic") & vbTab & UCase(rstT("generico")) & vbTab & UCase(rstT("descripcion")) & vbTab & UCase(rstT("fdes")) & vbTab & UCase(rstT("ldes")) & vbTab & Format(UCase(rstT("saldo")), "#,##0.00")
          Grilla.AddItem Fila
          If rstT("estado") = "V" Then
            For k = 0 To 5
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
         Next k
          End If
          Fila = ""
          rstT.MoveNext
      Next i
   
End Sub

Private Sub llenar_examenes(ByVal Grilla As MSHFlexGrid)
Dim color As String
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 7000
       Next
        cabecera = "IDEXAMEN" & vbTab & "NOMBRE EXAMEN"
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rstT.RecordCount - 1
          Fila = rstT("id") & vbTab & UCase(rstT("descripcion"))
          Grilla.AddItem Fila
          If rstT("estado") = "Pendiente" Then
             color = &H8080FF
          Else
             color = &HC0FFC0
          End If
          
          
          For k = 0 To 1
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = color
         Next k
          
          Fila = ""
          rstT.MoveNext
      Next i
   
End Sub

Private Sub HfgExamenes_SelChange()
If Val(Me.HfgExamenes.TextMatrix(Me.HfgExamenes.Row, 0)) > 0 Then
    strCadena = "SELECT PA.resultado,PA.estado FROM persona_analisis PA WHERE PA.id='" & Me.HfgExamenes.TextMatrix(Me.HfgExamenes.Row, 0) & "'"
    Call ConfiguraRst(strCadena)
    Me.TxtResultado.text = rst("resultado")
    strCadena = "SELECT * from persona_fotos_examen WHERE id_examen='" & Me.HfgExamenes.TextMatrix(Me.HfgExamenes.Row, 0) & "' AND dni='" & FrmHistoriaClinica.TxtRuc.text & "'"
    Call ConfiguraRstT(strCadena)
    If rst("estado") = "Pendiente" Then
        Me.TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
    Else
        Me.TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    End If
    If rstT.RecordCount > 0 Then
        Me.TlbAcciones.Buttons(KEY_NEW).Enabled = True
    Else
        Me.TlbAcciones.Buttons(KEY_NEW).Enabled = False
    End If
Else
    Me.TxtResultado.text = ""
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_NEW
        FrmHistoriaClinicaGaleria.Show
    Case KEY_UPDATE
        'Form1.TxtDni.text = Trim(FrmHistoriaClinica.TxtRuc.text)
        Form1.Show
        'FrmHistoriaClinicaVideo.Show
End Select
End Sub

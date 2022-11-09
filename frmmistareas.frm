VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmistareas 
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
   Begin VB.CheckBox chkestado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   16560
      TabIndex        =   125
      Top             =   6840
      Width           =   255
   End
   Begin MSDataListLib.DataCombo DtcEstadobusqueda 
      Height          =   330
      Left            =   16920
      TabIndex        =   124
      Top             =   6840
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
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
   Begin VB.Timer timmer_update 
      Interval        =   10000
      Left            =   16680
      Top             =   7920
   End
   Begin VitekeySoft.ChameleonBtn cmdbacklog 
      Height          =   495
      Left            =   16560
      TabIndex        =   45
      Top             =   3840
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "NUEVO PRODUCT BACKLOG"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmmistareas.frx":0000
      PICN            =   "frmmistareas.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frmbacklog 
      BackColor       =   &H00FFFFFF&
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
      Height          =   9120
      Left            =   120
      TabIndex        =   8
      Top             =   45
      Visible         =   0   'False
      Width           =   16335
      Begin VB.Frame frmboceto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9100
         Left            =   0
         TabIndex        =   100
         Top             =   8760
         Visible         =   0   'False
         Width           =   16320
         Begin VB.CommandButton cmdclose 
            Height          =   255
            Left            =   15960
            Picture         =   "frmmistareas.frx":2601
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   120
            Width           =   255
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
            Height          =   495
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   101
            Top             =   8520
            Width           =   15735
         End
         Begin VB.Image boceto_mater 
            Height          =   8295
            Left            =   240
            Stretch         =   -1  'True
            Top             =   120
            Width           =   15735
         End
      End
      Begin VB.TextBox txtid_backlog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   31
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComCtl2.MonthView calendario1 
         Height          =   2460
         Left            =   9240
         TabIndex        =   20
         Top             =   165
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   4339
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   14737632
         BorderStyle     =   1
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
         StartOfWeek     =   136118273
         TitleBackColor  =   33023
         CurrentDate     =   42221
      End
      Begin MSDataListLib.DataCombo dtcprioridad 
         Height          =   330
         Left            =   1680
         TabIndex        =   14
         Top             =   840
         Width           =   3735
         _ExtentX        =   6588
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
      Begin VB.TextBox txttitulo 
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
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   6975
      End
      Begin MSDataListLib.DataCombo dtcasignado 
         Height          =   330
         Left            =   1680
         TabIndex        =   15
         Top             =   1320
         Width           =   3735
         _ExtentX        =   6588
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
      Begin MSDataListLib.DataCombo dtcmodulo 
         Height          =   330
         Left            =   1680
         TabIndex        =   16
         Top             =   1800
         Width           =   3735
         _ExtentX        =   6588
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   6255
         Left            =   240
         TabIndex        =   17
         Top             =   2760
         Width           =   15735
         _ExtentX        =   27755
         _ExtentY        =   11033
         _Version        =   393216
         Tabs            =   4
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "DESCRIPCION"
         TabPicture(0)   =   "frmmistareas.frx":54A5
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label6"
         Tab(0).Control(1)=   "Label7"
         Tab(0).Control(2)=   "HfObservaciones"
         Tab(0).Control(3)=   "cmdcerrarpantalla"
         Tab(0).Control(4)=   "cmdprocesar"
         Tab(0).Control(5)=   "txtdetalle"
         Tab(0).Control(6)=   "cmdnuevaobservacion"
         Tab(0).Control(7)=   "frmobservacion"
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "BOCETOS"
         TabPicture(1)   =   "frmmistareas.frx":54C1
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "boceto(0)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "boceto(1)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "boceto(2)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "boceto(3)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cmdmaximizar1(3)"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "cmdmaximizar1(2)"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "cmdmaximizar1(1)"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "cmdupload04"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "cmdupload03"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "cmdupload02"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "cmdmaximizar1(0)"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "cmdupload01"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "txtboceto1(0)"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "txtobservacion1(0)"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "txtobservacion1(1)"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).Control(15)=   "txtobservacion1(2)"
         Tab(1).Control(15).Enabled=   0   'False
         Tab(1).Control(16)=   "txtobservacion1(3)"
         Tab(1).Control(16).Enabled=   0   'False
         Tab(1).Control(17)=   "ProgressBar1"
         Tab(1).Control(17).Enabled=   0   'False
         Tab(1).Control(18)=   "txtboceto1(1)"
         Tab(1).Control(18).Enabled=   0   'False
         Tab(1).Control(19)=   "txtboceto1(2)"
         Tab(1).Control(19).Enabled=   0   'False
         Tab(1).Control(20)=   "txtboceto1(3)"
         Tab(1).Control(20).Enabled=   0   'False
         Tab(1).Control(21)=   "txtimagen(0)"
         Tab(1).Control(21).Enabled=   0   'False
         Tab(1).Control(22)=   "txtimagen(1)"
         Tab(1).Control(22).Enabled=   0   'False
         Tab(1).Control(23)=   "txtimagen(2)"
         Tab(1).Control(23).Enabled=   0   'False
         Tab(1).Control(24)=   "txtimagen(3)"
         Tab(1).Control(24).Enabled=   0   'False
         Tab(1).ControlCount=   25
         TabCaption(2)   =   "TAREAS"
         TabPicture(2)   =   "frmmistareas.frx":54DD
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmddeletetarea"
         Tab(2).Control(1)=   "cmdnuevatarea"
         Tab(2).Control(2)=   "HfTarea"
         Tab(2).Control(3)=   "frmtarea"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "INFORME DIARIO"
         TabPicture(3)   =   "frmmistareas.frx":54F9
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "frminforme"
         Tab(3).Control(1)=   "dtpFechaInformeconsulta"
         Tab(3).Control(2)=   "cmdconsultar"
         Tab(3).Control(3)=   "HfInforme01"
         Tab(3).Control(4)=   "cmdinformenuevo"
         Tab(3).Control(5)=   "cmdeliminarinforme"
         Tab(3).Control(6)=   "DtcColaboradorInforme"
         Tab(3).Control(7)=   "cmdconsultarcolaborador"
         Tab(3).Control(8)=   "cmdvisualizar"
         Tab(3).Control(9)=   "Label22"
         Tab(3).Control(10)=   "Label16"
         Tab(3).ControlCount=   11
         Begin VB.TextBox txtimagen 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   12000
            TabIndex        =   122
            Top             =   4440
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtimagen 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   8280
            TabIndex        =   121
            Top             =   4440
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtimagen 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   3960
            TabIndex        =   120
            Top             =   4440
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtimagen 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   119
            Top             =   4440
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtboceto1 
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
            Index           =   3
            Left            =   12000
            TabIndex        =   115
            Top             =   4440
            Width           =   3615
         End
         Begin VB.TextBox txtboceto1 
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
            Index           =   2
            Left            =   8280
            TabIndex        =   114
            Top             =   4440
            Width           =   3615
         End
         Begin VB.TextBox txtboceto1 
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
            Index           =   1
            Left            =   3960
            TabIndex        =   113
            Top             =   4440
            Width           =   3615
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   5880
            TabIndex        =   112
            Top             =   1920
            Visible         =   0   'False
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.TextBox txtobservacion1 
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
            Height          =   765
            Index           =   3
            Left            =   12000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   111
            Top             =   4800
            Width           =   3615
         End
         Begin VB.TextBox txtobservacion1 
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
            Height          =   765
            Index           =   2
            Left            =   8280
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   110
            Top             =   4800
            Width           =   3615
         End
         Begin VB.TextBox txtobservacion1 
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
            Height          =   765
            Index           =   1
            Left            =   3960
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   109
            Top             =   4800
            Width           =   3615
         End
         Begin VB.TextBox txtobservacion1 
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
            Height          =   765
            Index           =   0
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   108
            Top             =   4800
            Width           =   3615
         End
         Begin VB.TextBox txtboceto1 
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
            Index           =   0
            Left            =   240
            TabIndex        =   103
            Top             =   4440
            Width           =   3615
         End
         Begin VB.Frame frminforme 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5415
            Left            =   -74520
            TabIndex        =   81
            Top             =   600
            Visible         =   0   'False
            Width           =   15135
            Begin VB.TextBox txtidinforme 
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
               Left            =   8160
               TabIndex        =   99
               Top             =   480
               Width           =   1455
            End
            Begin VB.CommandButton cmdcerrarinforme 
               Height          =   300
               Left            =   14400
               Picture         =   "frmmistareas.frx":5515
               Style           =   1  'Graphical
               TabIndex        =   94
               Top             =   120
               Width           =   300
            End
            Begin VB.TextBox txtcelularresponsable 
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
               Left            =   6000
               TabIndex        =   92
               Top             =   4200
               Width           =   3255
            End
            Begin VB.TextBox txtresponsableinforme 
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
               Left            =   6000
               TabIndex        =   91
               Top             =   3720
               Width           =   3255
            End
            Begin VB.TextBox txtdescripcion_informe 
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
               Height          =   2055
               Left            =   6000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   87
               Top             =   1560
               Width           =   8655
            End
            Begin VB.TextBox txthora_inicio_informe 
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
               Left            =   6000
               TabIndex        =   86
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox txthora_fin_informe 
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
               Left            =   6000
               TabIndex        =   85
               Top             =   1080
               Width           =   1455
            End
            Begin MSComCtl2.MonthView MonthInforme 
               Height          =   3210
               Left            =   360
               TabIndex        =   84
               Top             =   480
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   5662
               _Version        =   393216
               ForeColor       =   -2147483630
               BackColor       =   14737632
               BorderStyle     =   1
               Appearance      =   0
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               StartOfWeek     =   136118273
               TitleBackColor  =   33023
               CurrentDate     =   42221
            End
            Begin VitekeySoft.ChameleonBtn cmdguardarinforme 
               Height          =   645
               Left            =   12480
               TabIndex        =   93
               Top             =   4080
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   1138
               BTYPE           =   4
               TX              =   "GUARDAR  ITEM"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
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
               MICON           =   "frmmistareas.frx":83B9
               PICN            =   "frmmistareas.frx":83D5
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NUM. CELULAR :"
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
               Left            =   4590
               TabIndex        =   90
               Top             =   4320
               Width           =   1245
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "RESPONSABLE :"
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
               Left            =   4650
               TabIndex        =   89
               Top             =   3840
               Width           =   1185
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DESCRIPCION INFORME :"
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
               Left            =   3960
               TabIndex        =   88
               Top             =   1920
               Width           =   1875
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HORA FINALIZACION :"
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
               Left            =   4170
               TabIndex        =   83
               Top             =   1080
               Width           =   1665
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HORA INICIO :"
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
               Left            =   4770
               TabIndex        =   82
               Top             =   480
               Width           =   1065
            End
         End
         Begin MSComCtl2.DTPicker dtpFechaInformeconsulta 
            Height          =   345
            Left            =   -72960
            TabIndex        =   78
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   136118273
            CurrentDate     =   42386
         End
         Begin VB.Frame frmobservacion 
            BackColor       =   &H00FFFFFF&
            Caption         =   "DETALLE DE OBSERVACION."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   -70080
            TabIndex        =   70
            Top             =   2880
            Visible         =   0   'False
            Width           =   10335
            Begin VB.TextBox txtobservacionbacklog 
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
               Height          =   720
               Left            =   480
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   71
               Top             =   360
               Width           =   7575
            End
            Begin VitekeySoft.ChameleonBtn cmdsaveobservacion 
               Height          =   615
               Left            =   8160
               TabIndex        =   72
               Top             =   360
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   1085
               BTYPE           =   4
               TX              =   "AGREGAR OBSERVACION"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
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
               MICON           =   "frmmistareas.frx":896F
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
         End
         Begin VitekeySoft.ChameleonBtn cmdnuevaobservacion 
            Height          =   375
            Left            =   -72600
            TabIndex        =   69
            Top             =   3000
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            BTYPE           =   4
            TX              =   "AGREGAR OBSERVACION"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
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
            MICON           =   "frmmistareas.frx":898B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Frame frmtarea 
            Height          =   5655
            Left            =   -74640
            TabIndex        =   22
            Top             =   420
            Visible         =   0   'False
            Width           =   15255
            Begin VB.Frame Frame1 
               BackColor       =   &H00FFFFFF&
               Height          =   2175
               Left            =   1440
               TabIndex        =   36
               Top             =   2895
               Visible         =   0   'False
               Width           =   12015
               Begin VB.TextBox TxtHoraFin 
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
                  Left            =   4320
                  TabIndex        =   44
                  Top             =   1320
                  Width           =   1455
               End
               Begin VB.TextBox txtHoraInicio 
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
                  Left            =   1560
                  TabIndex        =   42
                  Top             =   1320
                  Width           =   1455
               End
               Begin VB.TextBox txtpersonaincidencia 
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
                  Left            =   1560
                  TabIndex        =   40
                  Top             =   960
                  Width           =   8535
               End
               Begin VB.TextBox txtincidencia 
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
                  Height          =   645
                  Left            =   1560
                  TabIndex        =   37
                  Top             =   240
                  Width           =   8535
               End
               Begin VitekeySoft.ChameleonBtn cmdagregarIncidencia 
                  Height          =   405
                  Left            =   10320
                  TabIndex        =   53
                  Top             =   240
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   714
                  BTYPE           =   4
                  TX              =   "AGREGAR"
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
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
                  MICON           =   "frmmistareas.frx":89A7
                  PICN            =   "frmmistareas.frx":89C3
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VitekeySoft.ChameleonBtn cmdcerrarincidencia 
                  Height          =   405
                  Left            =   10320
                  TabIndex        =   54
                  Top             =   720
                  Width           =   1485
                  _ExtentX        =   2619
                  _ExtentY        =   714
                  BTYPE           =   4
                  TX              =   "CERRAR "
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
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
                  MICON           =   "frmmistareas.frx":8F5D
                  PICN            =   "frmmistareas.frx":8F79
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label14 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "HORA FIN:"
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
                  Height          =   210
                  Left            =   3465
                  TabIndex        =   43
                  Top             =   1440
                  Width           =   810
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "HORA  INICIO:"
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
                  Height          =   210
                  Left            =   180
                  TabIndex        =   41
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "PERSONAL INCIDENCIA:"
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
                  Height          =   525
                  Left            =   -360
                  TabIndex        =   39
                  Top             =   720
                  Width           =   1635
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "DETALLE:"
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
                  Height          =   210
                  Left            =   585
                  TabIndex        =   38
                  Top             =   360
                  Width           =   690
               End
            End
            Begin VB.Frame frmestado 
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   495
               Left            =   9600
               TabIndex        =   32
               Top             =   1800
               Visible         =   0   'False
               Width           =   3735
               Begin MSDataListLib.DataCombo DtcEstado 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   33
                  Top             =   120
                  Width           =   2895
                  _ExtentX        =   5106
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
               Begin VitekeySoft.ChameleonBtn cmdestadocriterio 
                  Height          =   405
                  Left            =   3240
                  TabIndex        =   50
                  Top             =   40
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   714
                  BTYPE           =   3
                  TX              =   ""
                  ENAB            =   -1  'True
                  BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9
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
                  MICON           =   "frmmistareas.frx":BF8E
                  PICN            =   "frmmistareas.frx":BFAA
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
            End
            Begin VB.TextBox txtidtarea 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   10200
               TabIndex        =   30
               Top             =   720
               Visible         =   0   'False
               Width           =   1215
            End
            Begin MSComCtl2.DTPicker DtpInicio 
               Height          =   350
               Left            =   8400
               TabIndex        =   28
               Top             =   240
               Width           =   1400
               _ExtentX        =   2461
               _ExtentY        =   609
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   136118273
               CurrentDate     =   42224
            End
            Begin VB.TextBox TxtCriterio 
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
               Height          =   435
               Left            =   1440
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   27
               Top             =   735
               Width           =   6855
            End
            Begin VB.TextBox txtnombretarea 
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
               Height          =   525
               Left            =   1440
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   24
               Top             =   200
               Width           =   6855
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCriterios 
               Height          =   2295
               Left            =   1440
               TabIndex        =   25
               Top             =   1200
               Width           =   12015
               _ExtentX        =   21193
               _ExtentY        =   4048
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
            Begin MSComCtl2.DTPicker DtpFin 
               Height          =   350
               Left            =   10200
               TabIndex        =   29
               Top             =   240
               Width           =   1400
               _ExtentX        =   2461
               _ExtentY        =   609
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   136118273
               CurrentDate     =   42224
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfIncidencia 
               Height          =   1455
               Left            =   1440
               TabIndex        =   34
               Top             =   3600
               Width           =   12015
               _ExtentX        =   21193
               _ExtentY        =   2566
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
            Begin VitekeySoft.ChameleonBtn cmdagregarcriterio 
               Height          =   405
               Left            =   8400
               TabIndex        =   51
               Top             =   720
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   714
               BTYPE           =   3
               TX              =   ""
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
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
               MICON           =   "frmmistareas.frx":E58F
               PICN            =   "frmmistareas.frx":E5AB
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VitekeySoft.ChameleonBtn cmdeliminarcriterio 
               Height          =   405
               Left            =   8880
               TabIndex        =   52
               Top             =   720
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   714
               BTYPE           =   3
               TX              =   ""
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
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
               MICON           =   "frmmistareas.frx":EB45
               PICN            =   "frmmistareas.frx":EB61
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VitekeySoft.ChameleonBtn cmdguardartarea 
               Height          =   405
               Left            =   11520
               TabIndex        =   57
               Top             =   5100
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   714
               BTYPE           =   4
               TX              =   "GUARDAR TAREA"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
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
               MICON           =   "frmmistareas.frx":F0FB
               PICN            =   "frmmistareas.frx":F117
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VitekeySoft.ChameleonBtn ChameleonBtn5 
               Height          =   405
               Left            =   9480
               TabIndex        =   58
               Top             =   5100
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   714
               BTYPE           =   4
               TX              =   "CERRAR TAREAS"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
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
               MICON           =   "frmmistareas.frx":F6B1
               PICN            =   "frmmistareas.frx":F6CD
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VitekeySoft.ChameleonBtn cmdagergarincidencia 
               Height          =   1005
               Left            =   240
               TabIndex        =   73
               Top             =   4050
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   1773
               BTYPE           =   4
               TX              =   "AGREGAR INCIDENCIA"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
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
               MICON           =   "frmmistareas.frx":126E2
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "INCIDENCIAS / RECOMENDACIONES"
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
               Height          =   405
               Left            =   0
               TabIndex        =   35
               Top             =   3600
               Width           =   1365
            End
            Begin VB.Image Image5 
               Height          =   240
               Left            =   9840
               Picture         =   "frmmistareas.frx":126FE
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "CRITERIOS DE ACEPTACION :"
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
               Height          =   405
               Left            =   30
               TabIndex        =   26
               Top             =   840
               Width           =   1365
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DESCRIPCION :"
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
               Height          =   210
               Left            =   135
               TabIndex        =   23
               Top             =   360
               Width           =   1155
            End
         End
         Begin VB.TextBox txtdetalle 
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
            Height          =   1845
            Left            =   -72600
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   660
            Width           =   13095
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfTarea 
            Height          =   5295
            Left            =   -74640
            TabIndex        =   19
            Top             =   780
            Width           =   13695
            _ExtentX        =   24156
            _ExtentY        =   9340
            _Version        =   393216
            ForeColor       =   8388608
            FixedCols       =   0
            ForeColorFixed  =   8388608
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VitekeySoft.ChameleonBtn cmdupload01 
            Height          =   405
            Left            =   240
            TabIndex        =   49
            Top             =   5640
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            BTYPE           =   5
            TX              =   "SUBIR PROTIPO"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmistareas.frx":12C88
            PICN            =   "frmmistareas.frx":12CA4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdnuevatarea 
            Height          =   345
            Left            =   -60795
            TabIndex        =   55
            Top             =   780
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            BTYPE           =   4
            TX              =   "NUEVA     "
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
            MICON           =   "frmmistareas.frx":15289
            PICN            =   "frmmistareas.frx":152A5
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmddeletetarea 
            Height          =   345
            Left            =   -60795
            TabIndex        =   56
            Top             =   1140
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            BTYPE           =   4
            TX              =   "ELIMINAR"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
            MICON           =   "frmmistareas.frx":1583F
            PICN            =   "frmmistareas.frx":1585B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdprocesar 
            Height          =   405
            Left            =   -64200
            TabIndex        =   64
            Top             =   5685
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   714
            BTYPE           =   4
            TX              =   "PROCESAR"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmistareas.frx":15DF5
            PICN            =   "frmmistareas.frx":15E11
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdcerrarpantalla 
            Height          =   405
            Left            =   -61920
            TabIndex        =   65
            Top             =   5685
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   714
            BTYPE           =   4
            TX              =   "CERRAR PANTALLA"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmistareas.frx":163AB
            PICN            =   "frmmistareas.frx":163C7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfObservaciones 
            Height          =   2055
            Left            =   -72600
            TabIndex        =   68
            Top             =   3480
            Width           =   12855
            _ExtentX        =   22675
            _ExtentY        =   3625
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
         Begin VitekeySoft.ChameleonBtn cmdconsultar 
            Height          =   345
            Left            =   -71400
            TabIndex        =   75
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   609
            BTYPE           =   4
            TX              =   "CONSULTAR"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmistareas.frx":193DC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfInforme01 
            Height          =   4935
            Left            =   -74520
            TabIndex        =   76
            Top             =   1080
            Width           =   13215
            _ExtentX        =   23310
            _ExtentY        =   8705
            _Version        =   393216
            ForeColor       =   8388608
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
         Begin VitekeySoft.ChameleonBtn cmdinformenuevo 
            Height          =   585
            Left            =   -61200
            TabIndex        =   79
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1032
            BTYPE           =   4
            TX              =   "NUEVA     "
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
            MICON           =   "frmmistareas.frx":193F8
            PICN            =   "frmmistareas.frx":19414
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdeliminarinforme 
            Height          =   585
            Left            =   -61200
            TabIndex        =   80
            Top             =   2520
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1032
            BTYPE           =   4
            TX              =   "ELIMINAR"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
            MICON           =   "frmmistareas.frx":199AE
            PICN            =   "frmmistareas.frx":199CA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DtcColaboradorInforme 
            Height          =   330
            Left            =   -67680
            TabIndex        =   96
            Top             =   600
            Width           =   4695
            _ExtentX        =   8281
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
         Begin VitekeySoft.ChameleonBtn cmdconsultarcolaborador 
            Height          =   345
            Left            =   -62640
            TabIndex        =   97
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            BTYPE           =   4
            TX              =   "CONSULTAR"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmistareas.frx":19F64
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdvisualizar 
            Height          =   585
            Left            =   -61200
            TabIndex        =   98
            Top             =   1800
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1032
            BTYPE           =   4
            TX              =   "VISUALIZAR"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
            MICON           =   "frmmistareas.frx":19F80
            PICN            =   "frmmistareas.frx":19F9C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdmaximizar1 
            Height          =   405
            Index           =   0
            Left            =   2160
            TabIndex        =   104
            Top             =   5640
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            BTYPE           =   5
            TX              =   "MAXIMIZAR"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmistareas.frx":1A2B6
            PICN            =   "frmmistareas.frx":1A2D2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdupload02 
            Height          =   405
            Left            =   3960
            TabIndex        =   105
            Top             =   5640
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            BTYPE           =   5
            TX              =   "SUBIR PROTIPO"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmistareas.frx":1C8B7
            PICN            =   "frmmistareas.frx":1C8D3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdupload03 
            Height          =   405
            Left            =   8280
            TabIndex        =   106
            Top             =   5640
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            BTYPE           =   5
            TX              =   "SUBIR PROTIPO"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmistareas.frx":1EEB8
            PICN            =   "frmmistareas.frx":1EED4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdupload04 
            Height          =   405
            Left            =   12000
            TabIndex        =   107
            Top             =   5640
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            BTYPE           =   5
            TX              =   "SUBIR PROTIPO"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmistareas.frx":214B9
            PICN            =   "frmmistareas.frx":214D5
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdmaximizar1 
            Height          =   405
            Index           =   1
            Left            =   5880
            TabIndex        =   116
            Top             =   5640
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            BTYPE           =   5
            TX              =   "MAXIMIZAR"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmistareas.frx":23ABA
            PICN            =   "frmmistareas.frx":23AD6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdmaximizar1 
            Height          =   405
            Index           =   2
            Left            =   10200
            TabIndex        =   117
            Top             =   5640
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            BTYPE           =   5
            TX              =   "MAXIMIZAR"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmistareas.frx":260BB
            PICN            =   "frmmistareas.frx":260D7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdmaximizar1 
            Height          =   405
            Index           =   3
            Left            =   13920
            TabIndex        =   118
            Top             =   5640
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            BTYPE           =   5
            TX              =   "MAXIMIZAR"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmistareas.frx":286BC
            PICN            =   "frmmistareas.frx":286D8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Image boceto 
            Height          =   3855
            Index           =   3
            Left            =   12000
            Stretch         =   -1  'True
            Top             =   480
            Width           =   3615
         End
         Begin VB.Image boceto 
            Height          =   3855
            Index           =   2
            Left            =   8280
            Stretch         =   -1  'True
            Top             =   480
            Width           =   3615
         End
         Begin VB.Image boceto 
            Height          =   3855
            Index           =   1
            Left            =   3960
            Stretch         =   -1  'True
            Top             =   480
            Width           =   3615
         End
         Begin VB.Image boceto 
            Height          =   3855
            Index           =   0
            Left            =   240
            Stretch         =   -1  'True
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COLABORADOR :"
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
            Left            =   -69075
            TabIndex        =   95
            Top             =   720
            Width           =   1275
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA INFORME :"
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
            Left            =   -74505
            TabIndex        =   77
            Top             =   600
            Width           =   1365
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "OBSERVACIONES Y/O INCONVENIENTES."
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
            Height          =   1290
            Left            =   -74655
            TabIndex        =   67
            Top             =   4080
            Width           =   1785
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPCION GENERAL:"
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
            Height          =   210
            Left            =   -74745
            TabIndex        =   66
            Top             =   1320
            Width           =   1845
         End
      End
      Begin MSComCtl2.MonthView calendario2 
         Height          =   2460
         Left            =   13080
         TabIndex        =   21
         Top             =   165
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   4339
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
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
         StartOfWeek     =   136118273
         TitleBackColor  =   33023
         CurrentDate     =   42221
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   12480
         Picture         =   "frmmistareas.frx":2ACBD
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROYECTO :"
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
         Height          =   210
         Left            =   360
         TabIndex        =   13
         Top             =   1920
         Width           =   915
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ASIGNADO A :"
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
         Height          =   210
         Left            =   165
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRIORIDAD :"
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
         Height          =   210
         Left            =   255
         TabIndex        =   11
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TITULO :"
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
         Height          =   210
         Left            =   585
         TabIndex        =   9
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.TextBox txtcolaborador 
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
      Height          =   330
      Left            =   9480
      TabIndex        =   7
      Top             =   8640
      Width           =   975
   End
   Begin MSDataListLib.DataCombo dtccolaborador 
      Height          =   330
      Left            =   5160
      TabIndex        =   5
      Top             =   8640
      Width           =   4215
      _ExtentX        =   7435
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
   Begin MSComCtl2.MonthView calendario 
      Height          =   3210
      Left            =   16560
      TabIndex        =   4
      Top             =   480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5662
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   136118273
      TitleBackColor  =   33023
      CurrentDate     =   42219
   End
   Begin VB.TextBox txtbuscar 
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
      Height          =   330
      Left            =   1800
      TabIndex        =   0
      Top             =   8640
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3360
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmistareas.frx":2B247
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmistareas.frx":2B69B
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmistareas.frx":2B9BB
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmistareas.frx":2BE0F
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmistareas.frx":2C263
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmistareas.frx":2C583
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmistareas.frx":2C8A3
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmistareas.frx":2CBC3
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmistareas.frx":2CEE3
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin VitekeySoft.ChameleonBtn cmdingresar 
      Height          =   495
      Left            =   16560
      TabIndex        =   46
      Top             =   4440
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "INGRESAR BACKLOG            "
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmmistareas.frx":2D203
      PICN            =   "frmmistareas.frx":2D21F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdeliminar 
      Height          =   495
      Left            =   16560
      TabIndex        =   47
      Top             =   5040
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "ELIMINAR BACKLOG            "
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmmistareas.frx":2F804
      PICN            =   "frmmistareas.frx":2F820
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdcerrar 
      Height          =   495
      Left            =   16560
      TabIndex        =   48
      Top             =   8520
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "CERRAR PANTALLA              "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmmistareas.frx":31E05
      PICN            =   "frmmistareas.frx":31E21
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frmEstadoBacklog 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   795
      Left            =   11040
      TabIndex        =   59
      Top             =   2400
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdcerrarEstado 
         Height          =   220
         Left            =   4920
         Picture         =   "frmmistareas.frx":34406
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   50
         Width           =   220
      End
      Begin MSDataListLib.DataCombo DtcEstadoBacklog 
         Height          =   330
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VitekeySoft.ChameleonBtn cmdestadoBacklog 
         Height          =   375
         Left            =   4080
         TabIndex        =   61
         Top             =   360
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
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
         MICON           =   "frmmistareas.frx":372AA
         PICN            =   "frmmistareas.frx":372C6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO BACKLOG"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Width           =   1380
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgLinea 
      Height          =   7815
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   13785
      _Version        =   393216
      ForeColor       =   8388608
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
   Begin VitekeySoft.ChameleonBtn cmdinformes 
      Height          =   495
      Left            =   16560
      TabIndex        =   74
      Top             =   5640
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "INFORME DE TRABAJO       "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmmistareas.frx":398AB
      PICN            =   "frmmistareas.frx":398C7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VitekeySoft.ChameleonBtn cmdactualizar 
      Height          =   495
      Left            =   16560
      TabIndex        =   123
      Top             =   6240
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "ACTUALIZAR                        "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmistareas.frx":3BEAC
      PICN            =   "frmmistareas.frx":3BEC8
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
      Caption         =   "COLABORADOR "
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
      Height          =   210
      Left            =   3930
      TabIndex        =   6
      Top             =   8640
      Width           =   1245
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "METODOLOGIA SCRUM."
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
      Left            =   225
      TabIndex        =   3
      Top             =   120
      Width           =   2445
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION :"
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
      Height          =   210
      Left            =   420
      TabIndex        =   2
      Top             =   8640
      Width           =   1155
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   240
      Top             =   8400
      Width           =   16095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "frmmistareas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public str_ruta_img As String
Public Procedencia As EnumProcede

Private Sub ChameleonBtn3_Click()

End Sub


Public Sub cargar_bocetos()
Call disabled_form(Me)
If Val(Me.txtid_backlog.Text) < 1 Then
    GoTo SALIR_1
End If
strCadena = "SELECT * FROM proyecto_backlog_img WHERE id_backlog='" & Val(Me.txtid_backlog.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.ProgressBar1.Min = 0
   Me.ProgressBar1.Max = rst.RecordCount
   Me.ProgressBar1.Visible = True
   
   For i = 0 To rst.RecordCount - 1
       If i < 4 Then
           Call get_boceto(Me.boceto(i), rst("in_ruta_web"), rst("in_ruta_local"), rst("imagen"))
           Me.txtobservacion1(i).Text = UCase(rst("observacion"))
           Me.txtboceto1(i).Text = rst("in_ruta_web")
           Me.txtimagen(i).Text = rst("imagen")
           DoEvents
           Me.ProgressBar1.Value = i
       End If
       rst.MoveNext
   Next i
Else
SALIR_1:
    str_ruta_img = App.Path & "\imagenes\no_disponible.jpg"
    For i = 0 To 3
        Me.boceto(i) = LoadPicture(str_ruta_img)
        Me.txtobservacion1(i).Text = ""
        Me.txtboceto1(i).Text = ""
        Me.txtimagen(i).Text = ""
    Next i
End If
Call enabled_form(Me)
Me.ProgressBar1.Visible = False
End Sub
Private Sub ChameleonBtn5_Click()
frmtarea.Visible = False
End Sub

Private Sub cmd_Click()

End Sub

Private Sub CmdActualizar_Click()
If KEY_USUARIO <> "42546269" Then
    If Me.chkestado.Value = 1 Then
        strCadena = "SELECT * FROM view_backlog WHERE (dni_colaborador='" & KEY_USUARIO & "' or dni_save='" & KEY_USUARIO & "') AND  id_estado='" & Me.DtcEstadobusqueda.BoundText & "' and ruc='" & KEY_RUC & "' ORDER BY id_codigo ASC "
    Else
        strCadena = "SELECT * FROM view_backlog WHERE (dni_colaborador='" & KEY_USUARIO & "' or dni_save='" & KEY_USUARIO & "') AND ruc='" & KEY_RUC & "' ORDER BY id_codigo ASC "
    End If
    
Else
    If Me.chkestado.Value = 1 Then
        strCadena = "SELECT * FROM view_backlog WHERE  id_estado='" & Me.DtcEstadobusqueda.BoundText & "'  ORDER BY id_codigo DESC LIMIT 100 "
    Else
        strCadena = "SELECT * FROM view_backlog   ORDER BY id_codigo ASC   "
    End If
    
End If

Call backlog(Me.HfgLinea)
End Sub

Private Sub cmdagergarincidencia_Click()
Me.Frame1.Visible = True
Call Resalta(Me.txtincidencia)
End Sub

Private Sub cmdagregarcriterio_Click()

If Trim(Me.TxtCriterio.Text) <> "" And Val(Me.txtidtarea.Text) > 0 Then
strCadena = "INSERT INTO proyecto_testing(id_tarea,descripcion)VALUES('" & Val(Me.txtidtarea.Text) & "','" & Trim(Me.TxtCriterio.Text) & "')"
CnBd.Execute (strCadena)
 
Call Criterio(Me.HfCriterios, Val(Me.txtidtarea.Text))
Me.TxtCriterio.Text = ""
Call Resalta(Me.TxtCriterio)
Else
    MsgBox "Primero debe GRABAR la tarea para luego ingresar" + Chr(13) + "los criterios.", vbInformation
    Call Resalta(Me.TxtCriterio)
    Exit Sub
End If

End Sub

Private Sub cmdagregarIncidencia_Click()
If Trim(Me.txtincidencia.Text) <> "" And Trim(Me.txtpersonaincidencia.Text) <> "" Then
    strCadena = "INSERT INTO proyecto_incidencia (`id_tarea`,`descripcion`,`hora_inicio`,`hora_fin`,`operador`,`fecha_registro`) " & _
    "VALUES('" & Val(Me.txtidtarea.Text) & "','" & Trim(Me.txtincidencia.Text) & "','" & Trim(Me.txtHoraInicio.Text) & "','" & Trim(Me.TxtHoraFin.Text) & "','" & Trim(Me.txtpersonaincidencia.Text) & "',CURDATE())"
    CnBd.Execute (strCadena)
     
    Me.Frame1.Visible = False
    Call Incidencias(Me.HfIncidencia, Val(Me.txtidtarea.Text))
    
Else
    MsgBox "Ingrese los campos obligatorios", vbInformation
    Call Resalta(Me.txtincidencia)
    Exit Sub
End If
End Sub

Private Sub cmdbacklog_Click()
Me.calendario1.Value = KEY_FECHA
Me.calendario2.Value = KEY_FECHA


Call llenar_materia
Me.frmbacklog.Visible = True
Me.SSTab1.Tab = 0
Me.txttitulo.Text = ""
Me.TxtDetalle.Text = ""
Me.txtid_backlog.Text = ""
frmtarea.Visible = False
Me.HfTarea.Rows = 0
Me.HfObservaciones.Rows = 0

 str_ruta_img = App.Path & "\imagenes\no_disponible.jpg"
    For i = 0 To 3
        Me.boceto(i) = LoadPicture(str_ruta_img)
        Me.txtobservacion1(i).Text = ""
        Me.txtboceto1(i).Text = ""
        Me.txtimagen(i).Text = ""
    Next i
    
    
Call Resalta(Me.txttitulo)
End Sub
Public Sub llenar_materia()
    strCadena = "SELECT id_proyecto as Codigo,descripcion as Descripcion FROM proyecto"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.dtcmodulo)

strCadena = "SELECT id_prioridad as Codigo,descripcion as Descripcion FROM proyecto_prioridad order by id_prioridad asc"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcprioridad)

If KEY_USUARIO = "42546269" Then
    strCadena = "SELECT id_estado as Codigo,descripcion as Descripcion FROM proyecto_estado order by id_estado asc"
Else
    strCadena = "SELECT id_estado as Codigo,descripcion as Descripcion FROM proyecto_estado WHERE id_estado<>'04' order by id_estado asc"
End If
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstado)


strCadena = "SELECT dni as  Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcasignado)
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdcerrarEstado_Click()
frmEstadoBacklog.Visible = False
End Sub

Private Sub cmdcerrarincidencia_Click()
Me.Frame1.Visible = False
End Sub

Private Sub cmdcerrarinforme_Click()
Me.frminforme.Visible = False
End Sub

Private Sub cmdCerrarpantalla_Click()
Me.frmbacklog.Visible = False
End Sub

Private Sub cmdcerrartareas_Click()
Me.frmtarea.Visible = False
End Sub

Private Sub cmdcerrartareas_KeyPress(KeyAscii As Integer)

End Sub

Private Sub cmdClose_Click()
Me.frmboceto.Visible = False
End Sub

Private Sub cmdconsultarcolaborador_Click()
Call Me.llenar_informe(Me.HfInforme01, Me.dtpFechaInformeconsulta.Value, Me.DtcColaboradorInforme.BoundText)
End Sub

Private Sub cmddeletetarea_Click()
Procedencia = anular
FrmSeguridad.Show
Exit Sub
End Sub

Private Sub cmdEliminar_Click()
Procedencia = Eliminar
FrmSeguridad.Show
Exit Sub
End Sub

Private Sub cmdeliminarcriterio_Click()
Procedencia = anular_asignacion
FrmSeguridad.Show
Exit Sub
End Sub

Private Sub cmdeliminarinforme_Click()
Procedencia = eliminar_informe
frmsegurity.Show
Exit Sub
End Sub

Private Sub cmdestadoBacklog_Click()

strCadena = "UPDATE proyecto_backlog SET id_estado='" & Me.DtcEstadoBacklog.BoundText & "' WHERE id_codigo='" & Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) & "'"
CnBd.Execute (strCadena)
Me.frmbacklog.Visible = False
    
Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 7) = Me.DtcEstadoBacklog.Text
    For k = 0 To 5
            HfgLinea.col = k
            HfgLinea.Row = Me.HfgLinea.Row
            HfgLinea.CellBackColor = &HFFFF80
     Next k

Me.frmEstadoBacklog.Visible = False
End Sub

Private Sub cmdestadocriterio_Click()
strCadena = "UPDATE proyecto_testing SET id_estado='" & Me.DtcEstado.BoundText & "' WHERE id_testing='" & Val(Me.HfCriterios.TextMatrix(Me.HfCriterios.Row, 0)) & "'"
CnBd.Execute (strCadena)
 
Me.frmestado.Visible = False

If Me.DtcEstado.BoundText <> "1" Then  ' en proceso
   strCadena = "UPDATE proyecto_tareas SET id_estado='" & Me.DtcEstado.BoundText & "' WHERE  id_tarea='" & Val(Me.txtidtarea.Text) & "'"
   CnBd.Execute (strCadena)
   strCadena = "UPDATE proyecto_backlog SET id_estado='2' WHERE id_codigo='" & Trim(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) & "'"
   CnBd.Execute (strCadena)
    
End If

If Me.DtcEstado.BoundText = "1" Then ' en proceso
   strCadena = "SELECT * FROM proyecto_testing WHERE id_estado<>'" & Me.DtcEstado.BoundText & "' AND id_tarea='" & Val(Me.txtidtarea.Text) & "'"
   CnBd.Execute (strCadena)
    
   If rst.RecordCount < 1 Then
       strCadena = "UPDATE proyecto_tareas SET id_estado='" & Me.DtcEstado.BoundText & "' WHERE  id_tarea='" & Val(Me.txtidtarea.Text) & "'"
       CnBd.Execute (strCadena)
        
    Else
        strCadena = "UPDATE proyecto_tareas SET id_estado='1' WHERE  id_tarea='" & Val(Me.txtidtarea.Text) & "'"
       CnBd.Execute (strCadena)
        
   End If
End If

If Me.DtcEstado.BoundText = "5" Then ' en proceso
   strCadena = "UPDATE proyecto_tareas SET id_estado='" & Me.DtcEstado.BoundText & "' WHERE  id_tarea='" & Val(Me.txtidtarea.Text) & "'"
   CnBd.Execute (strCadena)
    
   
   strCadena = "SELECT * FROM proyecto_testing WHERE id_estado<>'" & Me.DtcEstado.BoundText & "' AND id_tarea='" & Val(Me.txtidtarea.Text) & "'"
   CnBd.Execute (strCadena)
    
   If rst.RecordCount < 1 Then
       strCadena = "UPDATE proyecto_tareas SET id_estado='" & Me.DtcEstado.BoundText & "' WHERE  id_tarea='" & Val(Me.txtidtarea.Text) & "'"
       CnBd.Execute (strCadena)
        
    Else
       strCadena = "UPDATE proyecto_tareas SET id_estado='2' WHERE  id_tarea='" & Val(Me.txtidtarea.Text) & "'"
       CnBd.Execute (strCadena)
        
   End If
   strCadena = "SELECT * FROM proyecto_tareas WHERE id_estado='2' and id_backlog='" & Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount < 1 Then
      strCadena = "UPDATE proyecto_backlog SET id_estado='5' WHERE id_codigo='" & Trim(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) & "'"
      CnBd.Execute (strCadena)
   End If
   
End If

If Me.DtcEstado.BoundText = "4" Then ' aprobado
   strCadena = "SELECT * FROM proyecto_testing WHERE id_estado<>'" & Me.DtcEstado.BoundText & "' AND id_tarea='" & Val(Me.txtidtarea.Text) & "'"
   CnBd.Execute (strCadena)
    
   If rst.RecordCount < 1 Then
       strCadena = "UPDATE proyecto_tareas SET id_estado='" & Me.DtcEstado.BoundText & "' WHERE  id_tarea='" & Val(Me.txtidtarea.Text) & "'"
       CnBd.Execute (strCadena)
        
    Else
        strCadena = "UPDATE proyecto_tareas SET id_estado='2' WHERE  id_tarea='" & Val(Me.txtidtarea.Text) & "'"
       CnBd.Execute (strCadena)
        
   End If
End If
Call Me.tareas(Me.HfTarea, Val(Me.txtid_backlog.Text))
Call Me.Criterio(Me.HfCriterios, Val(Me.txtidtarea.Text))
End Sub

Private Sub cmdguardarinforme_Click()
If Trim(Me.txthora_inicio_informe.Text) <> "" And Trim(Me.txthora_fin_informe.Text) <> "" And Trim(Me.txtdescripcion_informe.Text) <> "" Then
    strCadena = "INSERT INTO proyecto_informe(`fecha`,`hora`,hora_final,`descripcion`,`responsable`,`celular`,`dni_save`,`ruc`)VALUES " & _
    "('" & Format(Me.MonthInforme.Value, "YYYY-mm-dd") & "','" & Trim(Me.txthora_inicio_informe.Text) & "','" & Trim(Me.txthora_fin_informe.Text) & "','" & Trim(Me.txtdescripcion_informe.Text) & "','" & Trim(Me.txtresponsableinforme.Text) & "','" & Trim(Me.txtcelularresponsable.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Call Me.llenar_informe(Me.HfInforme01, Me.MonthInforme.Value, KEY_USUARIO)
    Me.frminforme.Visible = False
Else
    MsgBox "INGRESE LOS CAMPOS OBLIGATORIOS.", vbInformation, KEY_EMPRESA
End If
End Sub

Private Sub cmdguardartarea_Click()
If Val(Me.txtidtarea.Text) > 0 Then
    strCadena = "UPDATE proyecto_tareas SET descripcion='" & Trim(Me.txtnombretarea.Text) & "',fecha_ini='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "',fecha_fin='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' WHERE id_tarea='" & Val(Me.txtidtarea.Text) & "'"
Else
    If Trim(Me.txtnombretarea.Text) <> "" Then
        strCadena = "INSERT INTO proyecto_tareas(id_backlog,descripcion,fecha_ini,fecha_fin)VALUES('" & Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) & "','" & Trim(Me.txtnombretarea.Text) & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "')"
    Else
        MsgBox "Debe de ingresar un Descripcion", vbQuestion
        Call Resalta(Me.txtnombretarea)
        Exit Sub
    End If
End If
    CnBd.Execute (strCadena)
     
    Me.frmtarea.Visible = False
    Call Me.tareas(Me.HfTarea, Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)))
End Sub

Public Sub llena_tarea(ByVal in_tarea As Integer)
strCadena = "SELECT * FROM proyecto_tareas WHERE id_tarea='" & in_tarea & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.frmtarea.Visible = True
    Me.txtnombretarea.Text = rst("descripcion")
    Me.DtpInicio.Value = rst("fecha_ini")
    Me.DtpFin.Value = rst("fecha_fin")
    Call Criterio(Me.HfCriterios, Val(Me.txtidtarea.Text))
End If
End Sub

Private Sub cmdinformenuevo_Click()
Me.MonthInforme.Value = KEY_FECHA
Me.txthora_inicio_informe.Text = ""
Me.txthora_fin_informe.Text = ""
Me.txtdescripcion_informe.Text = ""
Me.txtresponsableinforme.Text = ""
Me.txtcelularresponsable.Text = ""
Me.frminforme.Visible = True
Me.cmdguardarinforme.Enabled = True
Call Resalta(Me.txthora_inicio_informe)
 
End Sub

Private Sub cmdinformes_Click()
 strCadena = "SELECT dni as  Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and  ruc='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcColaboradorInforme)
  Me.dtpFechaInformeconsulta.Value = KEY_FECHA
  Me.frmbacklog.Visible = True
  Me.SSTab1.Tab = 3


End Sub

Private Sub cmdingresar_Click()
Call Me.llenar_materia
Call LLENA(Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)))
Call backlog_observacion(Me.HfObservaciones, Val(Me.txtid_backlog.Text))
End Sub
Private Sub LLENA(ByVal in_codigo As Integer)
strCadena = "SELECT * FROM view_backlog WHERE id_codigo='" & in_codigo & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtid_backlog.Text = rst("id_codigo")
   Me.txttitulo.Text = rst("backlog")
   Me.dtcasignado.BoundText = rst("dni_colaborador")
   Me.dtcprioridad.BoundText = rst("id_prioridad")
   Me.dtcmodulo.BoundText = rst("id_proyecto")
      Me.TxtDetalle.Text = rst("detalle")
   Me.calendario1.Value = rst("fecha_inicio")
   Me.calendario2.Value = rst("fecha_final")
   
   Call Me.tareas(Me.HfTarea, txtid_backlog.Text)
   Me.Frame1.Visible = False
   Me.frmbacklog.Visible = True
Else
   Me.txtid_backlog.Text = 0
   Me.txttitulo.Text = ""
   Me.TxtDetalle.Text = ""
   Me.dtcasignado.BoundText = 0

   Me.frmbacklog.Visible = True
   
End If
End Sub

Public Function DownloadFile(Url As String, LocalFilename As String) As Boolean

Dim lngRetVal As Long
lngRetVal = URLDownloadToFile(0, Url, LocalFilename, 0, 0)
If lngRetVal = 0 Then DownloadFile = True
End Function



Private Sub cmdmaximizar2_Click()

End Sub

Private Sub cmdmaximizar1_Click(Index As Integer)
Me.frmboceto.Visible = True
Me.frmboceto.Top = 50
 Call get_boceto(Me.boceto_mater, Trim(Me.txtboceto1(Index)), "", Trim(Me.txtimagen(Index)))
 Me.txtDescripcion.Text = Me.txtobservacion1(Index)
End Sub

Private Sub cmdnuevaobservacion_Click()
Me.txtobservacionbacklog.Text = ""
Me.frmobservacion.Visible = True
Call Resalta(Me.txtobservacionbacklog)
End Sub

Private Sub cmdnuevatarea_Click()
Me.txtnombretarea.Text = ""
Me.txtidtarea.Text = ""
Me.frmtarea.Visible = True
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA
Call Resalta(Me.txtnombretarea)
Call Me.tareas(Me.HfTarea, Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)))
Call Me.Criterio(Me.HfCriterios, Val(Me.txtidtarea.Text))
End Sub

Private Sub cmdProcesar_Click()
If Trim(Me.txttitulo.Text) <> "" Then


If Trim(Me.txttitulo.Text) <> "" And Val(Me.txtid_backlog.Text) < 1 Then
    strCadena = "INSERT INTO proyecto_backlog(descripcion,dni_colaborador,id_proyecto,fecha_inicio,fecha_final,id_prioridad,detalle,dni_save,ruc) VALUES  " & _
    "('" & UCase(Trim(Me.txttitulo.Text)) & "','" & Me.dtcasignado.BoundText & "','" & Me.dtcmodulo.BoundText & "','" & Format(Me.calendario1.Value, "YYYY-mm-dd") & "','" & Format(Me.calendario2.Value, "YYYY-mm-dd") & "','" & Me.dtcprioridad.BoundText & "','" & Trim(Me.TxtDetalle.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
    If KEY_USUARIO <> "42546269" Then
        strCadena = "SELECT * FROM view_backlog WHERE (dni_colaborador='" & KEY_USUARIO & "' or dni_save='" & KEY_USUARIO & "') AND  id_estado<>'4' and ruc='" & KEY_RUC & "' ORDER BY id_codigo ASC limit 10"
    Else
        strCadena = "SELECT * FROM view_backlog   ORDER BY id_codigo ASC limit 10"
    End If
    Call backlog(Me.HfgLinea)
    Me.frmbacklog.Visible = False
Else
    strCadena = "UPDATE  proyecto_backlog SET detalle='" & Trim(Me.TxtDetalle.Text) & "', descripcion='" & UCase(Trim(Me.txttitulo.Text)) & "',dni_colaborador='" & Me.dtcasignado.BoundText & "',id_proyecto='" & Me.dtcmodulo.BoundText & "',fecha_inicio='" & Format(Me.calendario1.Value, "YYYY-mm-dd") & "',id_prioridad='" & Me.dtcprioridad.BoundText & "' WHERE id_codigo='" & Val(Me.txtid_backlog.Text) & "'"
    CnBd.Execute (strCadena)
     
    If KEY_USUARIO <> "42546269" Then
        strCadena = "SELECT * FROM view_backlog WHERE (dni_colaborador='" & KEY_USUARIO & "' or dni_save='" & KEY_USUARIO & "') AND  id_estado<>'4' and ruc='" & KEY_RUC & "' ORDER BY id_codigo ASC "
    Else
        strCadena = "SELECT * FROM view_backlog WHERE  id_estado<>'4'  ORDER BY id_codigo ASC "
    End If
    Call backlog(Me.HfgLinea)
    Me.frmbacklog.Visible = False
    Call calcular_dias(Me.txtid_backlog.Text)
End If
Else
    MsgBox "INGRESE UN TITULO", vbInformation, KEY_EMPRESA
End If
End Sub
Private Sub calcular_dias(ByVal in_backlog As Integer)

strCadena = "SELECT fecha_fin FROM proyecto_tareas WHERE id_backlog='" & in_backlog & "' ORDER BY fecha_ini DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   strCadena = "UPDATE proyecto_backlog SET fecha_final='" & Format(rst("fecha_fin"), "YYYY-mm-dd") & "' WHERE id_codigo='" & in_backlog & "'"
   CnBd.Execute (strCadena)
    
   Call Me.backlog(Me.HfgLinea)
End If
End Sub








Private Sub cmdsaveobservacion_Click()
strCadena = "INSERT INTO proyecto_backlog_observacion(fecha,hora,id_codigo,detalle)VALUES(CURDATE(),CURTIME(),'" & Val(Me.txtid_backlog.Text) & "','" & UCase(Trim(Me.txtobservacionbacklog.Text)) & "')"
CnBd.Execute (strCadena)
Me.txtobservacionbacklog.Text = ""
Me.frmobservacion.Visible = False
Call backlog_observacion(Me.HfObservaciones, Val(Me.txtid_backlog.Text))
End Sub



Private Sub cmdupload01_Click()
Dim imagen As String
On Error GoTo error_foto
imagen = Replace(Replace(Replace(KEY_FECHA & str(TimeValue(Time)), ".", ""), ":", ""), " ", "") & ".jpg"
str_ruta_img = App.Path & "\archivos\" & KEY_RUC & "\download\" & imagen
DownloadFile Trim(Me.txtboceto1(0).Text), str_ruta_img
Me.boceto(0) = LoadPicture(str_ruta_img)
Call save_boceto(imagen, Trim(Me.txtboceto1(0).Text), str_ruta_img, Val(Me.txtid_backlog.Text), Trim(Me.txtobservacion1(0).Text))
Exit Sub
error_foto:
MsgBox "ERROR AL CARGAR LA IMAGEN", vbInformation, KEY_EMPRESA
End Sub

Private Sub cmdupload02_Click()
Dim imagen As String
On Error GoTo error_foto
imagen = Replace(Replace(Replace(KEY_FECHA & str(TimeValue(Time)), ".", ""), ":", ""), " ", "") & ".jpg"
str_ruta_img = App.Path & "\archivos\" & KEY_RUC & "\download\" & imagen
DownloadFile Trim(Me.txtboceto1(1).Text), str_ruta_img

Me.boceto(1) = LoadPicture(str_ruta_img)
Call save_boceto(imagen, Trim(Me.txtboceto1(1).Text), str_ruta_img, Val(Me.txtid_backlog.Text), Trim(Me.txtobservacion1(1).Text))
Exit Sub
error_foto:
MsgBox "ERROR AL CARGAR LA IMAGEN", vbInformation, KEY_EMPRESA
End Sub

Private Sub cmdupload03_Click()
Dim imagen As String
On Error GoTo error_foto
imagen = Replace(Replace(Replace(KEY_FECHA & str(TimeValue(Time)), ".", ""), ":", ""), " ", "") & ".jpg"
str_ruta_img = App.Path & "\archivos\" & KEY_RUC & "\download\" & imagen
DownloadFile Trim(Me.txtboceto1(2).Text), str_ruta_img
Me.boceto(2) = LoadPicture(str_ruta_img)
Call save_boceto(imagen, Trim(Me.txtboceto1(2).Text), str_ruta_img, Val(Me.txtid_backlog.Text), Trim(Me.txtobservacion1(2).Text))
Exit Sub
error_foto:
MsgBox "ERROR AL CARGAR LA IMAGEN", vbInformation, KEY_EMPRESA
End Sub

Private Sub cmdupload04_Click()
Dim imagen As String
On Error GoTo error_foto
imagen = Replace(Replace(Replace(KEY_FECHA & str(TimeValue(Time)), ".", ""), ":", ""), " ", "") & ".jpg"
str_ruta_img = App.Path & "\archivos\" & KEY_RUC & "\download\" & imagen
DownloadFile Trim(Me.txtboceto1(3).Text), str_ruta_img
Me.boceto(3) = LoadPicture(str_ruta_img)
Call save_boceto(imagen, Trim(Me.txtboceto1(3).Text), str_ruta_img, Val(Me.txtid_backlog.Text), Trim(Me.txtobservacion1(3).Text))
Exit Sub
error_foto:
MsgBox "ERROR AL CARGAR LA IMAGEN", vbInformation, KEY_EMPRESA
End Sub

Private Sub cmdVisualizar_Click()
strCadena = "SELECT * FROM proyecto_informe WHERE id_informe='" & Val(Me.txtidinforme.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txtdescripcion_informe.Text = rst("descripcion")
    Me.txthora_inicio_informe.Text = rst("hora")
    Me.txthora_fin_informe.Text = rst("hora_final")
    Me.txtresponsableinforme.Text = rst("responsable")
    Me.txtcelularresponsable.Text = rst("celular")
    Me.MonthInforme.Value = rst("fecha")
    Me.cmdguardarinforme.Enabled = False
    Me.frminforme.Visible = True
Else
    Me.frminforme.Visible = False
End If
End Sub



Private Sub dtccolaborador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   strCadena = "SELECT * FROM view_backlog WHERE dni_colaborador='" & Me.dtccolaborador.BoundText & "'  AND    ruc='" & KEY_RUC & "' ORDER BY id_codigo ASC "
   Call backlog(Me.HfgLinea)
    
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.calendario1.Value = KEY_FECHA
Me.calendario2.Value = KEY_FECHA
Me.dtpFechaInformeconsulta.Value = KEY_FECHA
Me.calendario.Value = KEY_FECHA
Me.MonthInforme.Value = KEY_FECHA

strRuta = App.Path & "\archivos\" & KEY_RUC
If VerificarFichero(strRuta) = False Then
       Call MkDir(strRuta)
End If

strRuta = App.Path & "\archivos\" & KEY_RUC & "\download"
If VerificarFichero(strRuta) = False Then
       Call MkDir(strRuta)
End If



    strCadena = "SELECT id_estado as Codigo,descripcion as Descripcion FROM proyecto_estado order by id_estado asc"

Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstadobusqueda)



       
If KEY_USUARIO = "42546269" Then
    strCadena = "SELECT id_estado as Codigo,descripcion as Descripcion FROM proyecto_estado order by id_estado asc"
Else
    strCadena = "SELECT id_estado as Codigo,descripcion as Descripcion FROM proyecto_estado WHERE id_estado<>'04' order by id_estado asc"
End If
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstadoBacklog)
  strCadena = "SELECT dni as  Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and  ruc='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.dtccolaborador)

If KEY_USUARIO <> "42546269" Then
    strCadena = "SELECT * FROM view_backlog WHERE (dni_colaborador='" & KEY_USUARIO & "' or dni_save='" & KEY_USUARIO & "') AND  id_estado<>'4' and ruc='" & KEY_RUC & "' ORDER BY id_codigo ASC "
Else
    strCadena = "SELECT * FROM view_backlog WHERE  id_estado<>'4'  ORDER BY id_codigo ASC "
End If

'Call backlog(Me.HfgLinea)
   
   
   

End Sub
Public Sub backlog_observacion(ByVal Grilla As MSHFlexGrid, ByVal in_backlog As Double)
Dim Orden As Integer
On Error GoTo salir
strCadena = "SELECT * FROM proyecto_backlog_observacion WHERE  id_codigo='" & in_backlog & "' ORDER BY fecha,hora ASC "
Call ConfiguraRstL(strCadena)
If rstL.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rstL.Fields.Count)
       For Each Campo In rstL.Fields
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 10000
           
           
          Next
         cabecera = "FECHA" & vbTab & "HORA" & vbTab & "DETALLE"
         Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
         
         
        rstL.MoveFirst
       
        For i = 0 To rstL.RecordCount - 1
            Fila = Format(rstL("fecha"), "dd-mm-YYYY") & vbTab & Format(rstL("hora"), "HH:mm am/pm") & vbTab & rstL("detalle")
            Grilla.AddItem Fila
            rstL.MoveNext
        Next i
  
         
        
         
Exit Sub
salir:
Call manejador_error
End Sub
Public Sub llenar_informe(ByVal Grilla As MSHFlexGrid, ByVal in_fecha As String, ByVal in_usuario As String)
Dim Orden As Integer
On Error GoTo salir
strCadena = "SELECT * FROM proyecto_informe WHERE  fecha='" & Format(in_fecha, "YYYY-mm-dd") & "' and dni_save='" & in_usuario & "' and ruc='" & KEY_RUC & "' ORDER BY id_informe ASC "
Call ConfiguraRstL(strCadena)
If rstL.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rstL.Fields.Count)
       For Each Campo In rstL.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 8500
           Grilla.ColWidth(4) = 2000
          Next
         cabecera = "CODIGO" & vbTab & "HORA INICIO" & vbTab & "HORA FIN" & vbTab & "DESCRIPCION" & vbTab & "RESPONSABLE"
         Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
         
         
        rstL.MoveFirst
       
        For i = 0 To rstL.RecordCount - 1
            Fila = rstL("id_informe") & vbTab & rstL("hora") & vbTab & rstL("hora_final") & vbTab & rstL("descripcion") & vbTab & rstL("responsable")
            Grilla.AddItem Fila
            rstL.MoveNext
        Next i
  
         
        
         
Exit Sub
salir:
Call manejador_error
End Sub
Public Sub backlog(ByVal Grilla As MSHFlexGrid)
Dim Orden As Integer
Dim espacio As Single
On Error GoTo salir

Call ConfiguraRstL(strCadena)
If rstL.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rstL.Fields.Count)
       For Each Campo In rstL.Fields
           Grilla.ColWidth(0) = 600
           Grilla.ColWidth(1) = 4300
           Grilla.ColWidth(2) = 2500
           Grilla.ColWidth(3) = 2500
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 2000
           Grilla.ColWidth(6) = 1500
           Grilla.ColWidth(7) = 1100
           
Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "PROYECTO" & vbTab & "COLABORADOR" & vbTab & "PRIORIDAD" & vbTab & "FECHA TRABAJO" & vbTab & "COLABORADOR" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 7
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
         
         
        rstL.MoveFirst
       
        For i = 0 To rstL.RecordCount - 1
            Fila = Format(rstL("id_codigo"), "00000") & vbTab & rstL("backlog") & vbTab & rstL("proyecto") & vbTab & rstL("colaborador") & vbTab & rstL("prioridad") & vbTab & Format(rstL("fecha_inicio"), "dd-mm-YYYY") & " - " & Format(rstL("fecha_final"), "dd-mm-YYYY") & vbTab & get_persona(rstL("dni_save")) & vbTab & rstL("estado")
            Grilla.AddItem Fila
            Fila = ""
            For k = 0 To 7
                Grilla.col = k
                Grilla.Row = i + 1
                Select Case rstL("estado")
                    Case "EN PROCESO"
                            Grilla.CellBackColor = &H80C0FF       ' proceso
                    Case "RECHAZADO"
                            Grilla.CellBackColor = &H8080FF       ' desaprobado
                    Case "APROBADO"
                            Grilla.CellBackColor = &H80FF80    ' aprobado
                    Case "TERMINADO"
                            Grilla.CellBackColor = &HFFFF80       ' terminado
                End Select
            Next k
            rstL.MoveNext
        Next i
  
         
        
         
Exit Sub
salir:
Call manejador_error
End Sub
Public Sub tareas(ByVal Grilla As MSHFlexGrid, ByVal in_backlog As Integer)
Dim Orden As Integer
Dim espacio As Single

On Error GoTo salir
strCadena = "SELECT t.id_tarea,t.descripcion as tarea,t.fecha_ini,t.fecha_fin,e.descripcion as estado FROM proyecto_tareas t,proyecto_estado e WHERE t.id_estado=e.id_estado and t.id_backlog='" & Val(Me.txtid_backlog.Text) & "' ORDER BY id_tarea ASC "
Call ConfiguraRstL(strCadena)
If rstL.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rstL.Fields.Count)
       For Each Campo In rstL.Fields
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 9000
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 1200
      
          
           
          Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "F.INICIO" & vbTab & "F.FIN" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
         
         
        rstL.MoveFirst
       
        For i = 0 To rstL.RecordCount - 1
            Fila = Format(rstL("id_tarea"), "000") & vbTab & rstL("tarea") & vbTab & Format(rstL("fecha_ini"), "dd-mm-YYYY") & vbTab & Format(rstL("fecha_fin"), "dd-mm-YYYY") & vbTab & rstL("estado")
            Grilla.AddItem Fila
            Fila = ""
            For k = 0 To 4
                Grilla.col = k
                Grilla.Row = i + 1
                Select Case rstL("estado")
                    Case "EN PROCESO"
                            Grilla.CellBackColor = &H80C0FF       ' proceso
                    Case "RECHAZADO"
                            Grilla.CellBackColor = &H8080FF       ' desaprobado
                    Case "APROBADO"
                            Grilla.CellBackColor = &H80FF80    ' aprobado
                    Case "TERMINADO"
                            Grilla.CellBackColor = &HFFFF80       ' terminado
                End Select
            Next k
            rstL.MoveNext
        Next i
  
         
        
         
Exit Sub
salir:
Call manejador_error
End Sub
Public Sub Criterio(ByVal Grilla As MSHFlexGrid, ByVal in_tarea As Integer)
Dim Orden As Integer
Dim espacio As Single

On Error GoTo salir
strCadena = "SELECT t.id_testing,t.descripcion,e.descripcion as estado FROM proyecto_testing t,proyecto_estado e WHERE t.id_estado=e.id_estado and  t.id_tarea='" & in_tarea & "' ORDER BY id_testing ASC "
Call ConfiguraRstL(strCadena)
If rstL.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rstL.Fields.Count)
       For Each Campo In rstL.Fields
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 8500
           Grilla.ColWidth(2) = 1200
           
          
           
          Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
         
         
        rstL.MoveFirst
       
        For i = 0 To rstL.RecordCount - 1
            Fila = Format(rstL("id_testing"), "000") & vbTab & rstL("descripcion") & vbTab & rstL("estado")
            Grilla.AddItem Fila
            Fila = ""
            For k = 0 To 2
                Grilla.col = k
                Grilla.Row = i + 1
                Select Case rstL("estado")
                    Case "EN PROCESO"
                            Grilla.CellBackColor = &H80C0FF       ' proceso
                    Case "RECHAZADO"
                            Grilla.CellBackColor = &H8080FF       ' desaprobado
                    Case "APROBADO"
                            Grilla.CellBackColor = &H80FF80       ' aprobado
                    Case "TERMINADO"
                            Grilla.CellBackColor = &HFFFF80       ' terminado
                End Select
            Next k
            rstL.MoveNext
        Next i
  
         
        
         
Exit Sub
salir:
Call manejador_error
End Sub
Public Sub Incidencias(ByVal Grilla As MSHFlexGrid, ByVal in_tarea As Integer)
Dim Orden As Integer
Dim espacio As Single

On Error GoTo salir
strCadena = "SELECT * FROM proyecto_incidencia WHERE id_tarea='" & Val(Me.txtidtarea.Text) & "' ORDER BY hora_inicio asc"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rstL.Fields.Count)
       For Each Campo In rstL.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 5500
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 750
           Grilla.ColWidth(4) = 750
           
          Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "DESTINO" & vbTab & "H.INICIO" & vbTab & "H.FIN"
         Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
         
         
        rstL.MoveFirst
       
        For i = 0 To rstL.RecordCount - 1
            Fila = rstL("id_incidencia") & vbTab & rstL("descripcion") & vbTab & rstL("operador") & vbTab & rstL("hora_inicio") & vbTab & rstL("hora_fin")
            Grilla.AddItem Fila
            Fila = ""
            
            rstL.MoveNext
        Next i
  
         
        
         
Exit Sub
salir:
Call manejador_error
End Sub

Private Sub HfCriterios_Click()
Me.TxtCriterio.Text = Me.HfCriterios.TextMatrix(Me.HfCriterios.Row, 1)
End Sub

Private Sub HfCriterios_DblClick()
If Val(Me.HfCriterios.TextMatrix(Me.HfCriterios.Row, 0)) > 0 Then
    Me.frmestado.Visible = True
    
End If
End Sub

Private Sub HfgLinea_DblClick()

If Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) > 0 Then
   Me.frmEstadoBacklog.Visible = True
   Me.DtcEstadoBacklog.SetFocus
End If

End Sub

Private Sub HfgLinea_SelChange()
If Val(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0)) > 0 Then
   Me.cmdingresar.Enabled = True
   Me.cmdEliminar.Enabled = True
Else
    Me.cmdingresar.Enabled = False
   Me.cmdEliminar.Enabled = False
End If
End Sub

Private Sub HfInforme01_SelChange()
If Val(Me.HfInforme01.TextMatrix(Me.HfInforme01.Row, 0)) > 0 Then
    Me.txtidinforme.Text = Val(Me.HfInforme01.TextMatrix(Me.HfInforme01.Row, 0))
    Me.cmdeliminarinforme.Enabled = True
    Me.cmdVisualizar.Enabled = True
Else
    Me.txtidinforme.Text = 0
    Me.cmdeliminarinforme.Enabled = False
    Me.cmdVisualizar.Enabled = False
    
End If
End Sub

Private Sub HfTarea_DblClick()
Me.txtidtarea.Text = Val(Me.HfTarea.TextMatrix(Me.HfTarea.Row, 0))
Call llena_tarea(Val(Me.txtidtarea.Text))
Call Me.Incidencias(Me.HfIncidencia, Val(Me.txtidtarea.Text))
End Sub

Private Sub HfTarea_SelChange()
If Val(Me.HfTarea.TextMatrix(Me.HfTarea.Row, 0)) > 0 Then
    Me.cmddeletetarea.Enabled = True
Else
    Me.cmddeletetarea.Enabled = False
End If
End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
If Me.SSTab1.Tab = 1 Then
    Call cargar_bocetos
End If
End Sub

Private Sub timmer_update_Timer()
'Me.cmdactualizar.Caption = "ACTUALIZAR" & Space(1) & get_pendientes & Space(1) & "PENDIENTES"
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    strCadena = "SELECT * FROM view_backlog WHERE backlog LIKE '%" & Trim(Me.txtBuscar.Text) & "%'  AND    ruc='" & KEY_RUC & "' ORDER BY id_codigo ASC "
    Call backlog(Me.HfgLinea)
End If
End Sub

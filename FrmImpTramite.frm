VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmImpTramite 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   19200
      Top             =   5760
   End
   Begin VB.Frame frmmail 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAIL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   14760
      TabIndex        =   98
      Top             =   7680
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox txtmail 
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
         Left            =   120
         TabIndex        =   99
         Text            =   "comprobantesdepago_chicl@sunarp.gob.pe;v-2808@hotmail.com"
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.Frame frmweb 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8840
      Left            =   45
      TabIndex        =   93
      Top             =   60
      Visible         =   0   'False
      Width           =   18880
      Begin SHDocVwCtl.WebBrowser wbrInfo 
         Height          =   8445
         Left            =   0
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   18735
         ExtentX         =   33046
         ExtentY         =   14896
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
         Location        =   "http:///"
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesarapp 
         Height          =   345
         Left            =   14640
         TabIndex        =   95
         Top             =   8490
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmImpTramite.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdcerrar2 
         Height          =   345
         Left            =   16680
         TabIndex        =   96
         Top             =   8490
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         BTYPE           =   5
         TX              =   "CERRAR"
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmImpTramite.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesarSunnarp 
         Height          =   345
         Left            =   14640
         TabIndex        =   100
         Top             =   8490
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmImpTramite.frx":0038
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
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   18820
      Begin VB.TextBox txtdni_tarjeta 
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
         Height          =   285
         Left            =   2040
         TabIndex        =   108
         Top             =   7560
         Width           =   975
      End
      Begin VB.TextBox txtentrega_tarjeta 
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
         Height          =   285
         Left            =   3040
         TabIndex        =   107
         Top             =   7560
         Width           =   2800
      End
      Begin VB.TextBox txtobservacion_calificacion 
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
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   101
         Top             =   8280
         Width           =   5175
      End
      Begin VB.TextBox Text1 
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
         Left            =   6280
         TabIndex        =   84
         Top             =   4080
         Width           =   900
      End
      Begin MSMask.MaskEdBox txtfecharegistro 
         Height          =   285
         Left            =   2040
         TabIndex        =   75
         Top             =   5400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dtcparentesco 
         Height          =   330
         Left            =   11280
         TabIndex        =   73
         Top             =   6600
         Width           =   2175
         _ExtentX        =   3836
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
      Begin VitekeySoft.ChameleonBtn cmdcerrardos 
         Height          =   255
         Left            =   18240
         TabIndex        =   72
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MICON           =   "FrmImpTramite.frx":0054
         PICN            =   "FrmImpTramite.frx":0070
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
         Height          =   885
         Left            =   17280
         TabIndex        =   71
         Top             =   7800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1561
         BTYPE           =   7
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
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmImpTramite.frx":2F24
         PICN            =   "FrmImpTramite.frx":2F40
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtobservacion 
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
         Height          =   1005
         Left            =   11280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   69
         Top             =   7560
         Width           =   5175
      End
      Begin VB.TextBox txtentregadoa 
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
         Height          =   285
         Left            =   12600
         TabIndex        =   68
         Top             =   6240
         Width           =   3735
      End
      Begin VB.TextBox txtnumeroplaca 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11280
         TabIndex        =   67
         Top             =   5280
         Width           =   2175
      End
      Begin VB.TextBox txtdnientrega 
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
         Height          =   285
         Left            =   11280
         TabIndex        =   64
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox txtnumerotarjeta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   59
         Top             =   7200
         Width           =   2175
      End
      Begin VB.TextBox txtzonaregistral 
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
         Height          =   285
         Left            =   2040
         TabIndex        =   57
         Top             =   6120
         Width           =   5175
      End
      Begin VB.TextBox txttitulo 
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
         Height          =   285
         Left            =   2040
         TabIndex        =   55
         Top             =   5760
         Width           =   2175
      End
      Begin VB.TextBox txtmontoventa 
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
         Height          =   285
         Left            =   15120
         TabIndex        =   49
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtmontosaldo 
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
         Height          =   285
         Left            =   15120
         TabIndex        =   47
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtdocumento 
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
         Height          =   285
         Left            =   11400
         TabIndex        =   44
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox txtcolor 
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
         Left            =   2040
         TabIndex        =   40
         Top             =   3720
         Width           =   2535
      End
      Begin VB.TextBox txtmarca 
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
         Left            =   2040
         TabIndex        =   37
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox txtnumeromotor 
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
         Left            =   2040
         TabIndex        =   36
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox txtnumerochasis 
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
         Left            =   2040
         TabIndex        =   35
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox txtvehiculo 
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
         Left            =   2040
         TabIndex        =   34
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox txtcelularcliente 
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
         Left            =   10320
         TabIndex        =   27
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtubigueo 
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
         Left            =   10320
         TabIndex        =   25
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox txtdireccioncliente 
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
         Left            =   10320
         TabIndex        =   24
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txtnombrecliente 
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
         Left            =   2040
         TabIndex        =   23
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox txtdnicliente 
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
         Left            =   2040
         TabIndex        =   22
         Top             =   720
         Width           =   2175
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfFacturas 
         Height          =   1155
         Left            =   10080
         TabIndex        =   42
         Top             =   3240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   2037
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
      Begin MSDataListLib.DataCombo DtcEstado 
         Height          =   315
         Left            =   2040
         TabIndex        =   74
         Top             =   5010
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtfechaventa 
         Height          =   285
         Left            =   11400
         TabIndex        =   76
         Top             =   2400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfechaentrega 
         Height          =   285
         Left            =   11280
         TabIndex        =   77
         Top             =   7080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfecharecepcion 
         Height          =   285
         Left            =   11280
         TabIndex        =   78
         Top             =   5760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo DtcStatus 
         Height          =   330
         Left            =   11280
         TabIndex        =   80
         Top             =   4800
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8421631
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
      Begin MSDataListLib.DataCombo Dtcvendedor 
         Height          =   330
         Left            =   2040
         TabIndex        =   83
         Top             =   4080
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   8421631
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
      Begin MSDataListLib.DataCombo DtcResultado 
         Height          =   315
         Left            =   2040
         TabIndex        =   87
         Top             =   6480
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtfecha_entrega_tarjeta 
         Height          =   285
         Left            =   2040
         TabIndex        =   102
         Top             =   7920
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfecha_recepcion_tarjeta 
         Height          =   285
         Left            =   2040
         TabIndex        =   103
         Top             =   6840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo DtcParentescoTarjeta 
         Height          =   330
         Left            =   5880
         TabIndex        =   110
         Top             =   7560
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARENTESCO :"
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
         Left            =   5880
         TabIndex        =   111
         Top             =   7320
         Width           =   1080
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI (ENTREGA):"
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
         Left            =   720
         TabIndex        =   109
         Top             =   7560
         Width           =   1215
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACION :"
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
         Left            =   750
         TabIndex        =   106
         Top             =   8400
         Width           =   1185
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.RECEPCION :"
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
         Left            =   870
         TabIndex        =   105
         Top             =   6840
         Width           =   1065
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.ENTREGA :"
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
         Left            =   990
         TabIndex        =   104
         Top             =   7920
         Width           =   945
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACION :"
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
         Left            =   9960
         TabIndex        =   70
         Top             =   7920
         Width           =   1185
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N� PLACA :"
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
         Left            =   10350
         TabIndex        =   66
         Top             =   5280
         Width           =   795
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARENTESCO :"
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
         Left            =   10065
         TabIndex        =   65
         Top             =   6720
         Width           =   1080
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI (ENTREGA)"
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
         Left            =   9975
         TabIndex        =   63
         Top             =   6240
         Width           =   1170
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.ENTREGA :"
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
         Left            =   10200
         TabIndex        =   62
         Top             =   7200
         Width           =   945
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.RECEPCION :"
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
         Left            =   10080
         TabIndex        =   61
         Top             =   5760
         Width           =   1065
      End
      Begin VB.Image Image5 
         Height          =   300
         Left            =   9720
         Picture         =   "FrmImpTramite.frx":6588
         Top             =   4800
         Width           =   300
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "PLACAS"
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
         Left            =   10080
         TabIndex        =   60
         Top             =   4860
         Width           =   555
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00E0E0E0&
         Height          =   4060
         Left            =   9600
         Top             =   4680
         Width           =   7575
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N� TARJETA :"
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
         Left            =   960
         TabIndex        =   58
         Top             =   7200
         Width           =   975
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ZONA REGISTRAL :"
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
         Left            =   540
         TabIndex        =   56
         Top             =   6120
         Width           =   1395
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RES. CALIFICACION :"
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
         Left            =   420
         TabIndex        =   54
         Top             =   6480
         Width           =   1485
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N� TITULO :"
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
         Left            =   1065
         TabIndex        =   53
         Top             =   5760
         Width           =   870
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. REGISTRO :"
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
         Left            =   960
         TabIndex        =   52
         Top             =   5400
         Width           =   1035
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO  :"
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
         Left            =   1185
         TabIndex        =   51
         Top             =   5040
         Width           =   750
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   480
         Picture         =   "FrmImpTramite.frx":8B5D
         Top             =   4720
         Width           =   300
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "TRAMITE TARJETA DE PROPIEDAD"
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
         Left            =   840
         TabIndex        =   50
         Top             =   4740
         Width           =   3390
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00E0E0E0&
         Height          =   4060
         Left            =   360
         Top             =   4680
         Width           =   6975
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO VENTA:"
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
         Left            =   13800
         TabIndex        =   48
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SALDO :"
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
         Left            =   14400
         TabIndex        =   46
         Top             =   2760
         Width           =   600
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA VENTA :"
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
         Left            =   10080
         TabIndex        =   45
         Top             =   2400
         Width           =   1140
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMPROBANTE :"
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
         Left            =   10080
         TabIndex        =   43
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Image Image3 
         Height          =   300
         Left            =   9720
         Picture         =   "FrmImpTramite.frx":B132
         Top             =   1920
         Width           =   300
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "MONTOS Y SALDOS "
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
         Left            =   10080
         TabIndex        =   41
         Top             =   1980
         Width           =   3315
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00E0E0E0&
         Height          =   2775
         Left            =   9600
         Top             =   1800
         Width           =   7575
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENDEDOR :"
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
         Left            =   480
         TabIndex        =   39
         Top             =   4080
         Width           =   960
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COLOR :"
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
         Left            =   480
         TabIndex        =   38
         Top             =   3720
         Width           =   600
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MARCA :"
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
         Left            =   480
         TabIndex        =   33
         Top             =   3360
         Width           =   645
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N� MOTOR :"
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
         Left            =   480
         TabIndex        =   32
         Top             =   3000
         Width           =   915
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N� CHASIS :"
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
         Left            =   480
         TabIndex        =   31
         Top             =   2640
         Width           =   870
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCTO :"
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
         Left            =   480
         TabIndex        =   30
         Top             =   2280
         Width           =   945
      End
      Begin VB.Image Image2 
         Height          =   300
         Left            =   480
         Picture         =   "FrmImpTramite.frx":D707
         Top             =   1920
         Width           =   300
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "DATOS DE VEHICULO."
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
         Left            =   840
         TabIndex        =   29
         Top             =   1980
         Width           =   3435
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00E0E0E0&
         Height          =   2775
         Left            =   360
         Top             =   1800
         Width           =   6975
      End
      Begin VB.Image Image1 
         Height          =   300
         Left            =   480
         Picture         =   "FrmImpTramite.frx":FCDC
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "DATOS DE CLIENTE."
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
         Left            =   840
         TabIndex        =   28
         Top             =   420
         Width           =   3390
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00E0E0E0&
         Height          =   1335
         Left            =   360
         Top             =   240
         Width           =   16815
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CELULAR :"
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
         Left            =   9360
         TabIndex        =   26
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UBIGUEO :"
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
         Left            =   9360
         TabIndex        =   21
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCION :"
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
         Left            =   9240
         TabIndex        =   20
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE CLIENTE :"
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
         Left            =   480
         TabIndex        =   19
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI CLIENTE :"
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
         Left            =   480
         TabIndex        =   18
         Top             =   720
         Width           =   1020
      End
   End
   Begin VB.CheckBox chkmail 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "CONF MAIL"
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
      Height          =   250
      Left            =   18960
      TabIndex        =   97
      Top             =   8520
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   320
      Left            =   12960
      TabIndex        =   88
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   175112193
      CurrentDate     =   42873
   End
   Begin VB.TextBox txtplaca 
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
      Left            =   8880
      TabIndex        =   10
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtmotor 
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
      Left            =   5520
      TabIndex        =   8
      Top             =   525
      Width           =   2295
   End
   Begin VB.TextBox txtchasis 
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
      Left            =   5520
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtnombre_cliente 
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
      Left            =   1560
      TabIndex        =   4
      Top             =   520
      Width           =   2535
   End
   Begin VB.TextBox txtdni_cliente 
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
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Hflistado 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   18810
      _ExtentX        =   33179
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
   Begin VitekeySoft.ChameleonBtn cmdentregar 
      Height          =   840
      Left            =   18975
      TabIndex        =   11
      Top             =   1110
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1482
      BTYPE           =   5
      TX              =   "ACTUALIZAR"
      ENAB            =   0   'False
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmImpTramite.frx":122B1
      PICN            =   "FrmImpTramite.frx":122CD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdsunarp 
      Height          =   840
      Left            =   18975
      TabIndex        =   12
      Top             =   1980
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1482
      BTYPE           =   5
      TX              =   ""
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
      MICON           =   "FrmImpTramite.frx":14BC5
      PICN            =   "FrmImpTramite.frx":14BE1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdcerrar 
      Height          =   840
      Left            =   18975
      TabIndex        =   13
      Top             =   4600
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1482
      BTYPE           =   5
      TX              =   "CERRAR"
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmImpTramite.frx":17D59
      PICN            =   "FrmImpTramite.frx":17D75
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdapp 
      Height          =   840
      Left            =   18975
      TabIndex        =   79
      Top             =   2850
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1482
      BTYPE           =   5
      TX              =   ""
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
      MICON           =   "FrmImpTramite.frx":18165
      PICN            =   "FrmImpTramite.frx":18181
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcStatusBuscar 
      Height          =   330
      Left            =   8880
      TabIndex        =   81
      Top             =   525
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   8421631
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
   Begin VitekeySoft.ChameleonBtn cmdmail 
      Height          =   840
      Left            =   18975
      TabIndex        =   85
      Top             =   3720
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   1482
      BTYPE           =   5
      TX              =   "E-MAIL"
      ENAB            =   0   'False
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmImpTramite.frx":1AD50
      PICN            =   "FrmImpTramite.frx":1AD6C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   315
      Left            =   12960
      TabIndex        =   89
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   175112193
      CurrentDate     =   42873
   End
   Begin VitekeySoft.ChameleonBtn cmdbuscar 
      Height          =   645
      Left            =   14640
      TabIndex        =   92
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1138
      BTYPE           =   5
      TX              =   ""
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmImpTramite.frx":1F28F
      PICN            =   "FrmImpTramite.frx":1F2AB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F. INICIO:"
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
      Left            =   12120
      TabIndex        =   91
      Top             =   120
      Width           =   630
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F. FIN:"
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
      Left            =   12240
      TabIndex        =   90
      Top             =   525
      Width           =   420
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "ENVIADO A REGISTROS"
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
      Height          =   225
      Left            =   15720
      TabIndex        =   86
      Top             =   825
      Width           =   2640
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESTADO :"
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
      Left            =   8040
      TabIndex        =   82
      Top             =   525
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "PENDIENTE DE REGISTRO"
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
      Height          =   225
      Left            =   15720
      TabIndex        =   16
      Top             =   570
      Width           =   2640
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "LISTO PARA ENTREGAR"
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
      Height          =   225
      Left            =   15720
      TabIndex        =   15
      Top             =   315
      Width           =   2640
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "TRAJETA Y PLACA ENTREGADO"
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
      Height          =   225
      Left            =   15720
      TabIndex        =   14
      Top             =   60
      Width           =   2640
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N� PLACA :"
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
      Left            =   7920
      TabIndex        =   9
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N� MOTOR  :"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   525
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N� CHASIS  :"
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
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOM. CLIENTE :"
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
      Left            =   240
      TabIndex        =   3
      Top             =   525
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI CLIENTE    :"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   990
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8925
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmImpTramite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Dim in_envio As Integer

Private Sub ChameleonBtn1_Click()

End Sub



Private Sub chkmail_Click()
If Me.chkmail.Value = 1 Then
   Me.frmmail.Visible = True
Else
   Me.frmmail.Visible = False
End If
End Sub

Private Sub cmdapp_Click()
Me.wbrInfo.Visible = True
wbrInfo.Navigate "https://www.placas.pe/Public/CheckPlateStatus.aspx"
Me.frmweb.Visible = True
Me.cmdprocesarSunnarp.Visible = False
Me.cmdprocesarapp.Visible = True
End Sub

Private Sub cmdBuscar_Click()
strCadena = "SELECT * FROM view_tramite WHERE fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'  "
Call actualizar(Me.HfListado)
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdcerrar2_Click()
Me.frmweb.Visible = False
End Sub

Private Sub cmdcerrardos_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub cmdcerrarnavegador_Click()

End Sub

Private Sub cmdentregar_Click()
Call llenar_datos(Me.HfListado.TextMatrix(Me.HfListado.Row, 0))
Me.frmdetalle.Visible = True
End Sub
Public Sub llenar_datos(ByVal in_tramite As Double)

strCadena = "SELECT * FROM view_tramite WHERE id_tramite='" & in_tramite & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtdnicliente.Text = rst("id_cliente")
   Me.txtnombrecliente.Text = rst("ncliente")
   Me.txtdireccioncliente.Text = rst("direccion")
   Me.txtcelularcliente.Text = ""
   Me.txtubigueo.Text = get_ubigueo_persona(rst("id_cliente"), 0)
   Me.txtvehiculo.Text = rst("detalle")
   Me.txtnumerochasis.Text = rst("nro_chasis")
   Me.txtnumeromotor.Text = rst("serie")
   Me.TxtMarca.Text = rst("marca")
   Me.txtcolor.Text = rst("color")
   Me.txtobservacion_calificacion.Text = rst("calificacion_observacion")
   Me.txtcelularcliente.Text = get_telefono(rst("id_cliente"))
   Call load_usuario(Me.DtcVendedor, rst("id_vendedor"), "")
  
  ' Me.txtvendedor.Text = rst("vendedor")
   Me.txtfechaventa.Text = Format(rst("fecha_emision"), "dd-mm-YYYY")
   Me.txtmontoventa.Text = rst("total")
   Me.txtmontosaldo.Text = rst("saldo")
   Me.TxtDocumento.Text = rst("documento")
   Me.DtcStatus.BoundText = rst("id_status")
   
   
   If IsNull(rst("fecha_recepcion_tarjeta")) = True Then
       Me.txtfecha_recepcion_tarjeta.Mask = ""
       Me.txtfecha_recepcion_tarjeta.Text = ""
       Me.txtfecha_recepcion_tarjeta.Mask = "##-##-####"
   Else
       txtfecha_recepcion_tarjeta.Text = Format(rst("fecha_recepcion_tarjeta"), "dd-mm-YYYY")
   End If
   
   If IsNull(rst("fecha_entrega_tarjeta")) = True Then
       Me.txtfecha_entrega_tarjeta.Mask = ""
       Me.txtfecha_entrega_tarjeta.Text = ""
       Me.txtfecha_entrega_tarjeta.Mask = "##-##-####"
   Else
       txtfecha_entrega_tarjeta.Text = Format(rst("fecha_entrega_tarjeta"), "dd-mm-YYYY")
   End If
   
   
   
   If IsNull(rst("fecha_ingreso_registro")) = True Then
       Me.txtfecharegistro.Mask = ""
       Me.txtfecharegistro.Text = ""
       Me.txtfecharegistro.Mask = "##-##-####"
   Else
       txtfecharegistro.Text = Format(rst("fecha_ingreso_registro"), "dd-mm-YYYY")
   End If
   
   
   
   Me.DtcEstado.BoundText = rst("id_estado")
   Me.DtcResultado.BoundText = rst("id_resultado")
   
   Me.txttitulo.Text = rst("titulo")
   Me.txtzonaregistral.Text = rst("zona_registral")
   Me.txtnumerotarjeta.Text = rst("numero_tarjeta")
   
   
   If IsNull(rst("fecha_recepcion_placa")) = True Then
       Me.txtfecharecepcion.Mask = ""
       Me.txtfecharecepcion.Text = ""
       Me.txtfecharecepcion.Mask = "##-##-####"
   Else
       Me.txtfecharecepcion.Text = Format(rst("fecha_recepcion_placa"), "dd-mm-YYYY")
   End If
   
   If IsNull(rst("fecha_entrega")) = True Then
       Me.txtfechaentrega.Mask = ""
       Me.txtfechaentrega.Text = ""
       Me.txtfechaentrega.Mask = "##-##-####"
   Else
       Me.txtfechaentrega.Text = Format(rst("fecha_entrega"), "dd-mm-YYYY")
   End If
   
   If IsNull(rst("dni_entrega_tarjeta")) = False Then
        Me.txtdni_tarjeta.Text = rst("dni_entrega_tarjeta")
        Me.txtentrega_tarjeta.Text = rst("nombre_entrega_tarjeta")
        Me.DtcParentescoTarjeta.BoundText = rst("parentesco_tarjeta")
    Else
        Me.txtdni_tarjeta.Text = " "
        Me.txtentrega_tarjeta.Text = rst("nombre_entrega_tarjeta")
        Me.DtcParentescoTarjeta.BoundText = rst("parentesco_tarjeta")
   End If
   
   
   
   Me.txtdnientrega.Text = rst("dni_entrega")
   Me.txtentregadoa.Text = rst("nombre_entrega")
   Me.dtcparentesco.BoundText = rst("parentesco")
   
   
   
   Me.txtnumeroplaca.Text = rst("numero_placa")
   Me.txtObservacion.Text = rst("observacion")
   
   Call llenar_pagos(Me.HfFacturas, rst("id_venta"))
   
   
   
End If


End Sub
Public Sub enviar_mail_xml(ByVal strHtml As String)


5 strCadena = "UPDATE imp_tramite SET id_status='4' WHERE id_tramite='" & Val(Me.HfListado.TextMatrix(Me.HfListado.Row, 0)) & "'"
CnBd.Execute (strCadena)

For k = 9 To 14
                HfListado.col = k
                HfListado.Row = Me.HfListado.Row
                HfListado.CellBackColor = &HFFFF00
Next k

Call enabled_form(Me)
                    
End Sub

Private Sub cmdestado_Click()
wbrInfo.Navigate "http://www.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias"
End Sub

Public Function enviar_mail(ByVal in_key As String, ByVal IN_asunto As String, ByVal in_mail As String)
Call disabled_form(Me)
Procedencia = mailenviar
FrmLoad_web_service.Show
FrmLoad_web_service.nom_prcedimiento = "enviar_mail_xml"
Set FrmLoad_web_service.FormPadre = Me
If KEY_SERVIDOR_CLOUD = "si" Then
     
     
     If Len(in_key) > 36 Then
        in_abc = "5e8836bb8b3445ac15f0c0c5a22815d31cd7e84df87b5b60f40ff612cff7d41c"
        Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/enviarxml", "POST", json_facturacion_electronica_mail(KEY_RUC, in_key, IN_asunto, in_mail), "{x-api-token: '" & in_abc & "', x-api-produccion: 'yes'}")
     Else
        Call FrmLoad_web_service.crear_json_facturacion_electronica("https://api.vitekey.com/keyfact/erp/invoice-send-email?password=vitekey2018", "POST", json_facturacion_electronica_mail(KEY_RUC, in_key, IN_asunto, in_mail), "{x-api-token: '" & KEY_TOKEN_CLOUD & "', x-api-produccion: 'yes'}")
     
     End If
     
    
Else
    Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/enviarxml", "POST", json_facturacion_electronica_mail(KEY_RUC, in_key, IN_asunto, in_mail), "{x-api-token: '" & KEY_TOKEN_LOCAL & "', x-api-produccion: 'yes'}")
End If
End Function
Private Function get_titulo(ByVal in_tramite As String) As String

strCadena = "SELECT titulo FROM imp_tramite WHERE id_tramite='" & Val(in_tramite) & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
   get_titulo = rstT("titulo")

End If


End Function

Private Sub cmdmail_Click()

Me.txtmail.Tag = 0
Me.Timer1.Enabled = True








Exit Sub
 
 
 
 
 
 
Dim in_titulo As String
Dim in_mail() As String

strCadena = "SELECT sunat_key FROM movimiento_venta where id_venta='" & Val(Me.HfListado.TextMatrix(Me.HfListado.Row, 14)) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   If Len(rst("sunat_key")) > 5 Then
      in_titulo = Trim(get_titulo(Val(Me.HfListado.TextMatrix(Me.HfListado.Row, 0))))
        If Len(in_titulo) < 2 Then
            MsgBox "Este Expediente aun no cuenta con un TITULO", vbInformation
            Exit Sub
        End If
       in_titulo = "TITULO:" & in_titulo
       
       'in_mail = Split(Trim(Me.txtmail.Text), ";")
       
       'If UBound(in_mail) > 0 Then
          'For i = 0 To UBound(in_mail)
            
              Call enviar_mail(rst("sunat_key"), in_titulo, Trim(Me.txtmail.Text))
              
              
              'Call enviar_mail(rst("sunat_key"), in_titulo, Trim(in_mail(i)))
          'Next i
       'End If
    
      
       
       
   Else
        MsgBox "REGULARICE ESTE COMPROBANTE" + Chr(13) + "INCONSISTENCIA DE XML", vbInformation
   End If
End If

End Sub

Private Sub enviar_mail_individual(ByVal in_venta As String, ByVal in_mail As String)
strCadena = "SELECT sunat_key FROM movimiento_venta where id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   If Len(rst("sunat_key")) > 5 Then
      in_titulo = Trim(get_titulo(Val(Me.HfListado.TextMatrix(Me.HfListado.Row, 0))))
        If Len(in_titulo) < 2 Then
            MsgBox "Este Expediente aun no cuenta con un TITULO", vbInformation
            Exit Sub
        End If
       in_titulo = "TITULO:" & in_titulo
       Call enviar_mail(rst("sunat_key"), in_titulo, Trim(in_mail))
              
           
      
       
       
   Else
        MsgBox "REGULARICE ESTE COMPROBANTE" + Chr(13) + "INCONSISTENCIA DE XML", vbInformation
   End If
End If
End Sub




Private Sub cmdProcesar_Click()
Dim fecha_registro As String
Dim fecha_entrega As String
Dim fecha_recepcion As String
Dim StrCriterio As String

StrCriterio = ""
If IsDate(txtfecharegistro.Text) = False Then
    fecha_registro = ""
Else
  
    StrCriterio = ",fecha_ingreso_registro='" & Format(Me.txtfecharegistro.Text, "YYYY-mm-dd") & "'"
End If





If IsDate(Me.txtfechaentrega.Text) = False Then
   fecha_entrega = ""
Else
  
    If Len(StrCriterio) > 0 Then
        StrCriterio = StrCriterio & ",fecha_entrega='" & Format(Me.txtfechaentrega.Text, "YYYY-mm-dd") & "'"
    Else
        StrCriterio = ",fecha_entrega='" & Format(Me.txtfechaentrega.Text, "YYYY-mm-dd") & "'"
    End If
End If

'Fecha recepcion tarjeta
If IsDate(Me.txtfecha_recepcion_tarjeta.Text) = False Then
   fecha_recepcion_tarjeta = ""
Else
    If Len(StrCriterio) > 0 Then
        StrCriterio = StrCriterio & ",fecha_recepcion_tarjeta='" & Format(Me.txtfecha_recepcion_tarjeta.Text, "YYYY-mm-dd") & "'"
    Else
        StrCriterio = ",fecha_recepcion_tarjeta='" & Format(Me.txtfecha_recepcion_tarjeta.Text, "YYYY-mm-dd") & "'"
    End If
End If


'Fecha entrega tarjeta
If IsDate(Me.txtfecha_entrega_tarjeta.Text) = False Then
   fecha_entrega_tarjeta = ""
Else
    If Len(StrCriterio) > 0 Then
        StrCriterio = StrCriterio & ",fecha_entrega_tarjeta='" & Format(Me.txtfecha_entrega_tarjeta.Text, "YYYY-mm-dd") & "'"
    Else
        StrCriterio = ",fecha_entrega_tarjeta='" & Format(Me.txtfecha_entrega_tarjeta.Text, "YYYY-mm-dd") & "'"
    End If
End If





If IsDate(Me.txtfecharecepcion.Text) = False Then
   fecha_recepcion = ""
Else
    If Len(StrCriterio) > 0 Then
        StrCriterio = StrCriterio & ",fecha_recepcion_placa='" & Format(Me.txtfecharecepcion.Text, "YYYY-mm-dd") & "'"
    Else
        StrCriterio = ",fecha_recepcion_placa='" & Format(Me.txtfecharecepcion.Text, "YYYY-mm-dd") & "'"
    End If
    
End If


strCadena = "UPDATE imp_tramite SET calificacion_observacion='" & Trim(Me.txtobservacion_calificacion.Text) & "',id_resultado='" & Me.DtcResultado.BoundText & "', id_status='" & Me.DtcStatus.BoundText & "', id_estado='" & Me.DtcEstado.BoundText & "',titulo='" & Trim(Me.txttitulo.Text) & "',zona_registral='" & Trim(Me.txtzonaregistral.Text) & "',numero_tarjeta='" & Trim(Me.txtnumerotarjeta.Text) & "',numero_placa='" & Trim(Me.txtnumeroplaca.Text) & "', " & _
" dni_entrega_tarjeta='" & Trim(Me.txtdni_tarjeta.Text) & "',nombre_entrega_tarjeta='" & Trim(Me.txtentrega_tarjeta.Text) & "',parentesco_tarjeta='" & Me.DtcParentescoTarjeta.BoundText & "',dni_entrega='" & Trim(Me.txtdnientrega.Text) & "',nombre_entrega='" & Trim(Me.txtentregadoa.Text) & "',parentesco='" & Me.dtcparentesco.BoundText & "',observacion='" & Trim(Me.txtObservacion.Text) & "'" & StrCriterio & " WHERE id_tramite='" & Val(Me.HfListado.TextMatrix(Me.HfListado.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

Me.frmdetalle.Visible = False

strCadena = "SELECT * FROM view_tramite WHERE id_tramite='" & Val(Me.HfListado.TextMatrix(Me.HfListado.Row, 0)) & "' and   ruc='" & KEY_RUC & "'"
Call actualizar(Me.HfListado)



End Sub

Private Sub cmdprocesarapp_Click()
    contenido = Me.wbrInfo.Document.documentElement.innerHTML
    PosNomCom = InStr(1, wbrInfo.Document.documentElement.innerHTML, "dgConsulta_ctl02_lblNroContrato") + 1
    PosNomCom = InStr(PosNomCom, wbrInfo.Document.documentElement.innerHTML, ">") + 1
    PosNomCFin = InStr(PosNomCom, wbrInfo.Document.documentElement.innerHTML, "</")
    
    If Len(Mid$(wbrInfo.Document.documentElement.innerHTML, PosNomCom, PosNomCFin - PosNomCom)) > 12 Then
        Print
        'Me.lblActivo.Caption = "*****"
    Else
        
        PosNomCom = InStr(1, wbrInfo.Document.documentElement.innerHTML, "dgConsulta_ctl02_Label1") + 1
        PosNomCom = InStr(PosNomCom, wbrInfo.Document.documentElement.innerHTML, ">") + 1
        PosNomCFin = InStr(PosNomCom, wbrInfo.Document.documentElement.innerHTML, "</")
        'Me.lblActivo.Caption = Mid$(wbrInfo.Document.documentElement.innerHTML, PosNomCom, PosNomCFin - PosNomCom)
    
    End If
    
End Sub

Private Sub cmdprocesarSunnarp_Click()
contenido = Me.wbrInfo.Document.documentElement.innerHTML
     
    
    PosNomCom = InStr(1, wbrInfo.Document.documentElement.innerHTML, "Resultado de la Calificaci�n") + 1
    PosNomCom = InStr(PosNomCom, wbrInfo.Document.documentElement.innerHTML, ">") + 1
    PosNomCFin = InStr(PosNomCom, wbrInfo.Document.documentElement.innerHTML, "</")
    
    If Len(Mid$(wbrInfo.Document.documentElement.innerHTML, PosNomCom, PosNomCFin - PosNomCom)) > 12 Then
        Print
        'Me.lblActivo.Caption = "*****"
    Else
        
        PosNomCom = InStr(1, wbrInfo.Document.documentElement.innerHTML, "dgConsulta_ctl02_Label1") + 1
        PosNomCom = InStr(PosNomCom, wbrInfo.Document.documentElement.innerHTML, ">") + 1
        PosNomCFin = InStr(PosNomCom, wbrInfo.Document.documentElement.innerHTML, "</")
        'Me.lblActivo.Caption = Mid$(wbrInfo.Document.documentElement.innerHTML, PosNomCom, PosNomCFin - PosNomCom)
    
    End If
End Sub

Private Sub cmdsunarp_Click()
Me.wbrInfo.Visible = True
wbrInfo.Navigate "https://enlinea.sunarp.gob.pe/sunarpweb/pages/acceso/frmTitulos.faces"
Me.frmweb.Visible = True
Me.cmdprocesarSunnarp.Visible = True
Me.cmdprocesarapp.Visible = False
End Sub

Private Sub DtcStatusBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_tramite WHERE id_status = '" & Me.DtcStatusBuscar.BoundText & "' and  ruc='" & KEY_RUC & "'"
    Call actualizar(Me.HfListado)
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50

Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA

strCadena = "SELECT id_estado as Codigo,descripcion as Descripcion FROM imp_tramite_estado WHERE ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstado)

strCadena = "SELECT id_resultado as Codigo, descripcion as Descripcion FROM imp_tramite_resultado "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcResultado)
Me.DtcResultado.BoundText = 1

strCadena = "SELECT id_status as Codigo,descripcion as Descripcion FROM imp_tramite_status "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcStatus)
strCadena = "SELECT id_status as Codigo,descripcion as Descripcion FROM imp_tramite_status "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(DtcStatusBuscar)
 
 strCadena = "SELECT id_parentesco as Codigo,descripcion as Descripcion FROM parentesco ORDER BY descripcion"
 Call ConfiguraRst(strCadena)
 Call LlenaDataCombo(Me.dtcparentesco)
 Me.dtcparentesco.BoundText = 0
 
 strCadena = "SELECT id_parentesco as Codigo,descripcion as Descripcion FROM parentesco ORDER BY descripcion"
 Call ConfiguraRst(strCadena)
 Call LlenaDataCombo(Me.DtcParentescoTarjeta)
 Me.DtcParentescoTarjeta.BoundText = 0
 

strCadena = "SELECT * FROM view_tramite WHERE ruc='" & KEY_RUC & "' LIMIT 28 "

Call actualizar(Me.HfListado)



End Sub
Public Sub actualizar(ByVal Grilla As MSHFlexGrid)
Dim color As String

Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
                           
       ' edita la celda
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 1800
           Grilla.ColWidth(3) = 900
           Grilla.ColWidth(4) = 900
           Grilla.ColWidth(5) = 1100
           Grilla.ColWidth(6) = 2500
           Grilla.ColWidth(7) = 2500
           Grilla.ColWidth(8) = 1000
           Grilla.ColWidth(9) = 1700
           Grilla.ColWidth(10) = 1700
           Grilla.ColWidth(11) = 1200
           Grilla.ColWidth(12) = 1200
           Grilla.ColWidth(13) = 1200
           Grilla.ColWidth(14) = 0
        Next
        
        cabecera = "ID" & vbTab & "FECHA VENTA" & vbTab & "COMPROBANTE" & vbTab & "TOTAL" & vbTab & "SALDO" & vbTab & "DNI/RUC" & vbTab & "CLIENTE" & vbTab & "DETALLE VEHICULO" & vbTab & "COLOR" & vbTab & "N� CHASIS" & vbTab & "N�SERIE" & vbTab & "FECHA REG" & vbTab & "N� TARJETA" & vbTab & "N� PLACA" & vbTab & "IDVENTA"
        Grilla.AddItem cabecera
       
         For k = 1 To 13
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
         rstT.MoveFirst
        
        For i = 0 To rstT.RecordCount - 1
            
        
        
          Fila = rstT("id_tramite") & vbTab & Format(rstT("fecha_emision"), "dd-mm-YYYY") & vbTab & rstT("documento") & vbTab & Format(rstT("total"), "#,##0.00") & vbTab & Format(rstT("saldo"), "#,##0.00") & vbTab & rstT("id_cliente") & vbTab & rstT("ncliente") & vbTab & rstT("detalle") & vbTab & rstT("color") & vbTab & rstT("nro_chasis") & vbTab & rstT("serie") & vbTab & rstT("fecha_ingreso_registro") & vbTab & rstT("numero_tarjeta") & vbTab & rstT("numero_placa") & vbTab & rstT("id_venta")
          Grilla.AddItem Fila
          Select Case rstT("id_status")
            Case "1"
                color = &H8080FF ' rojo
            Case "2"
                color = &H80C0FF     ' verde
            Case "3"
                color = &H80FF80      ' amarillo
            Case "4"
                color = &HFFFF00      ' amarillo
            End Select
          
          For k = 9 To 14
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = color
          Next k
          Grilla.ColAlignment(9) = 0
          Grilla.ColAlignment(10) = 0
          rstT.MoveNext
      Next i
     
    
End Sub


Public Sub llenar_pagos(ByVal Grilla As MSHFlexGrid, ByVal id_venta As Double)
Dim color As String
strCadena = "SELECT * FROM movimiento_venta WHERE id_comprobante='" & id_venta & "' and id_recibo='0' and anulado='no' and ruc='" & KEY_RUC & "'"
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
                           
       ' edita la celda
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1300
           Grilla.ColWidth(2) = 2200
           Grilla.ColWidth(3) = 1200
           
        Next
        cabecera = "CODIGO" & vbTab & "FECHA PAGO" & vbTab & "COMPROBANTE" & vbTab & "MONTO"
        Grilla.AddItem cabecera
       
         For k = 1 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
            
        
        
          Fila = rstT("id_venta") & vbTab & Format(rstT("fecha_emision"), "dd-mm-YYYY") & vbTab & rstT("documento") & vbTab & rstT("total")
          Grilla.AddItem Fila
         
          
          rstT.MoveNext
      Next i
     
    
End Sub

Private Sub Hflistado_SelChange()
If Val(Me.HfListado.TextMatrix(Me.HfListado.Row, 0)) > 0 Then
    Me.cmdentregar.Enabled = True
    Me.cmdmail.Enabled = True
    Me.cmdentregar.Enabled = True
Else
    Me.cmdentregar.Enabled = False
    Me.cmdmail.Enabled = False
    Me.cmdentregar.Enabled = False
End If
End Sub

Private Sub Timer1_Timer()
Dim in_mail() As String
    in_mail = Split(Trim(Me.txtmail.Text), ";")
       
       If UBound(in_mail) >= Val(Me.txtmail.Tag) Then
          'For i = 0 To UBound(in_mail)
            
             ' Call enviar_mail(rst("sunat_key"), in_titulo, Trim(Me.txtmail.Text))
              
              
            Call enviar_mail_individual(Val(Me.HfListado.TextMatrix(Me.HfListado.Row, 14)), in_mail(Me.txtmail.Tag))
            Me.txtmail.Tag = Val(Me.txtmail.Tag) + 1
          'Next i
        Else
            Me.Timer1.Enabled = False
       End If
       
       
       
       
End Sub

Private Sub txtChasis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_tramite WHERE nro_chasis LIKE '%" & Trim(Me.txtchasis.Text) & "%' and  ruc='" & KEY_RUC & "'"
    Call actualizar(Me.HfListado)
End If
End Sub

Private Sub txtdni_cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_tramite WHERE id_cliente LIKE '%" & Trim(Me.txtdni_cliente.Text) & "%' and  ruc='" & KEY_RUC & "'"
    Call actualizar(Me.HfListado)
End If
End Sub

Private Sub txtdni_tarjeta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   strCadena = "SELECT nombre_completo FROM persona WHERE dni='" & Trim(Me.txtdni_tarjeta.Text) & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      Me.txtentrega_tarjeta.Text = rst("nombre_completo")
   End If
End If
End Sub

Private Sub txtdnientrega_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   strCadena = "SELECT nombre_completo FROM persona WHERE dni='" & Trim(Me.txtdnientrega.Text) & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      Me.txtentregadoa.Text = rst("nombre_completo")
   End If
End If
End Sub

Private Sub txtMotor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_tramite WHERE serie LIKE '%" & Trim(Me.TxtMotor.Text) & "%' and  ruc='" & KEY_RUC & "'"
    Call actualizar(Me.HfListado)
End If
End Sub

Private Sub txtnombre_cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_tramite WHERE ncliente LIKE '%" & Trim(Me.txtnombre_cliente.Text) & "%' and  ruc='" & KEY_RUC & "'"
    Call actualizar(Me.HfListado)
End If

End Sub

Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_tramite WHERE numero_placa LIKE '%" & Trim(Me.TxtPlaca.Text) & "%' and  ruc='" & KEY_RUC & "'"
    Call actualizar(Me.HfListado)
End If
End Sub


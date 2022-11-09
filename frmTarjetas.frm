VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTarjetas 
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
   Begin VB.Frame frm_agenda 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6735
      Left            =   -4560
      TabIndex        =   77
      Top             =   2040
      Visible         =   0   'False
      Width           =   9015
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2460
         Left            =   5880
         TabIndex        =   81
         Top             =   240
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   4339
         _Version        =   393216
         ForeColor       =   8388608
         BackColor       =   -2147483633
         BorderStyle     =   1
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
         StartOfWeek     =   186974209
         CurrentDate     =   44611
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDias 
         Height          =   2175
         Left            =   120
         TabIndex        =   78
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   3836
         _Version        =   393216
         ForeColor       =   8388608
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         GridColor       =   12582912
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfAgenda 
         Height          =   4095
         Left            =   120
         TabIndex        =   80
         Top             =   2520
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   7223
         _Version        =   393216
         ForeColor       =   8388608
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         GridColor       =   12582912
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
      Begin VitekeySoft.ChameleonBtn cmdGenerarAgenda 
         Height          =   495
         Left            =   5880
         TabIndex        =   82
         Top             =   3360
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "GENERAR AGENDA"
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
         MICON           =   "frmTarjetas.frx":0000
         PICN            =   "frmTarjetas.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcVendedorAgenda 
         Height          =   315
         Left            =   5880
         TabIndex        =   83
         Top             =   2880
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VitekeySoft.ChameleonBtn cmdAgendar 
         Height          =   520
         Left            =   6120
         TabIndex        =   84
         Top             =   5160
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "AGENDAR CLIENTE"
         ENAB            =   0   'False
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
         MICON           =   "frmTarjetas.frx":32F2
         PICN            =   "frmTarjetas.frx":330E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdAnularAgenda 
         Height          =   520
         Left            =   6120
         TabIndex        =   85
         Top             =   5880
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "LIMPIAR AGENDA"
         ENAB            =   0   'False
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
         MICON           =   "frmTarjetas.frx":5A95
         PICN            =   "frmTarjetas.frx":5AB1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdVerAgenda 
         Height          =   315
         Left            =   8560
         TabIndex        =   86
         Top             =   2880
         Width           =   340
         _ExtentX        =   609
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   ""
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
         MICON           =   "frmTarjetas.frx":5DCB
         PICN            =   "frmTarjetas.frx":5DE7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   1575
         Left            =   5880
         Shape           =   4  'Rounded Rectangle
         Top             =   5040
         Width           =   2775
      End
      Begin VB.Shape Shape12 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Height          =   6735
         Left            =   0
         Top             =   0
         Width           =   8775
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AGENDA PERSONAL."
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         TabIndex        =   79
         Top             =   0
         Width           =   1515
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   8640
         Picture         =   "frmTarjetas.frx":8312
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.Frame frmAsignacion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Left            =   14040
      TabIndex        =   71
      Top             =   1680
      Visible         =   0   'False
      Width           =   3960
      Begin VitekeySoft.ChameleonBtn cmdAsignacion_clientes 
         Height          =   645
         Left            =   120
         TabIndex        =   72
         Top             =   360
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   1138
         BTYPE           =   5
         TX              =   "ASIGNACION AUTOMATICA"
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
         MICON           =   "frmTarjetas.frx":B1B6
         PICN            =   "frmTarjetas.frx":B1D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar prg_avance 
         Height          =   195
         Left            =   120
         TabIndex        =   73
         Top             =   120
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfVendedores 
         Height          =   3015
         Left            =   120
         TabIndex        =   74
         Top             =   1080
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   5318
         _Version        =   393216
         ForeColor       =   8388608
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         GridColor       =   12582912
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
      Begin VB.Shape Shape13 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   4215
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   3960
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   3600
         Picture         =   "frmTarjetas.frx":EF13
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.Frame frm_impresion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   19800
      TabIndex        =   66
      Top             =   5040
      Visible         =   0   'False
      Width           =   4935
      Begin VitekeySoft.ChameleonBtn cmdReporteEstado 
         Height          =   780
         Left            =   240
         TabIndex        =   67
         Top             =   195
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1376
         BTYPE           =   5
         TX              =   "TARJETA ESTADO"
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
         MICON           =   "frmTarjetas.frx":11DB7
         PICN            =   "frmTarjetas.frx":11DD3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdReporteVendedor 
         Height          =   780
         Left            =   1680
         TabIndex        =   68
         Top             =   195
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1376
         BTYPE           =   5
         TX              =   "TARJETA VENDEDOR"
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
         MICON           =   "frmTarjetas.frx":14FD7
         PICN            =   "frmTarjetas.frx":14FF3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdRendimiento 
         Height          =   780
         Left            =   3120
         TabIndex        =   69
         Top             =   195
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1376
         BTYPE           =   5
         TX              =   "RENDIMIENTO"
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
         MICON           =   "frmTarjetas.frx":181F7
         PICN            =   "frmTarjetas.frx":18213
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   4560
         Picture         =   "frmTarjetas.frx":1B417
         Top             =   120
         Width           =   240
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   1215
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   4575
      End
   End
   Begin VB.Timer timmer_notificacion 
      Interval        =   5000
      Left            =   19440
      Top             =   7560
   End
   Begin VB.CheckBox chk_vendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      Caption         =   "VENDEDOR :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3960
      TabIndex        =   55
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtBusquedaCelular 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   5160
      TabIndex        =   50
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtBusquedaApoderado 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   1560
      TabIndex        =   48
      Top             =   600
      Width           =   2055
   End
   Begin VB.Frame frmtarjeta 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DETALLE TARJETA"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8000
      Left            =   4920
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   13575
      Begin TabDlg.SSTab SSTab1 
         Height          =   7695
         Left            =   240
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   13573
         _Version        =   393216
         TabHeight       =   520
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "DATOS DE CLIENTE"
         TabPicture(0)   =   "frmTarjetas.frx":1E2BB
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Shape3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label6"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label4"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label5"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label3"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label2"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label11"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Shape5"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label8"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label9"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label10"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label12"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "lblcomprobante"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label15"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "lblhora"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "lblNotificacion"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label21"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Shape15"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label22"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "cmdagenda_persona"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "cmdregistrar_sinruc"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "DtcPlan"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "DtcEstado"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "DtcMedio"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "DtcVendedor"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "DtpFecha"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "cmdTransferir"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "cmdRegistrar"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "DtcRubro"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "txtRazonSocialApoderado"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "txtDniApoderado"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "TxtRazonSocial"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "txtCelular"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "txtRucEmpresa"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "chk_asignado"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "frmtransferir"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "txtEmail"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "txtBuscarProducto"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "opt_interesado"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "opt_demostracion"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "opt_poratender"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).ControlCount=   41
         TabCaption(1)   =   "ACTIVIDADES"
         TabPicture(1)   =   "frmTarjetas.frx":1E2D7
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "chk_servicio_tecnico"
         Tab(1).Control(1)=   "chk_notificacion"
         Tab(1).Control(2)=   "frmactividad_estado"
         Tab(1).Control(3)=   "chk_cambio_estado"
         Tab(1).Control(4)=   "txtActividad"
         Tab(1).Control(5)=   "HfActividades"
         Tab(1).Control(6)=   "cmdAgregarActividad"
         Tab(1).Control(7)=   "cmdNotificar"
         Tab(1).Control(8)=   "DtcSupervidor"
         Tab(1).Control(9)=   "cmdCerrarNotificaciones"
         Tab(1).Control(10)=   "Label7"
         Tab(1).Control(11)=   "Shape4"
         Tab(1).ControlCount=   12
         TabCaption(2)   =   "SOPORTE TECNICO"
         TabPicture(2)   =   "frmTarjetas.frx":1E2F3
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "HfSoporteTecnico"
         Tab(2).ControlCount=   1
         Begin VB.CheckBox chk_servicio_tecnico 
            Appearance      =   0  'Flat
            BackColor       =   &H00800000&
            Caption         =   "SOPORTE TECNICO"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiCondensed"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   350
            Left            =   -64680
            TabIndex        =   98
            Top             =   480
            Width           =   1935
         End
         Begin VB.OptionButton opt_poratender 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "POR ATENDER"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3720
            TabIndex        =   95
            Top             =   7190
            Width           =   3255
         End
         Begin VB.OptionButton opt_demostracion 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "DEMOSTRACION"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3720
            TabIndex        =   93
            Top             =   6840
            Width           =   3255
         End
         Begin VB.OptionButton opt_interesado 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            Caption         =   "INTERESADO"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   3720
            TabIndex        =   92
            Top             =   6500
            Width           =   3255
         End
         Begin VB.CheckBox chk_notificacion 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "NOTIFICAR"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiCondensed"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   350
            Left            =   -73440
            TabIndex        =   65
            Top             =   2100
            Width           =   2055
         End
         Begin VB.TextBox txtBuscarProducto 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   7860
            TabIndex        =   58
            Top             =   5460
            Width           =   750
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   2040
            TabIndex        =   16
            Top             =   2700
            Width           =   5760
         End
         Begin VB.Frame frmtransferir 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1575
            Left            =   3960
            TabIndex        =   43
            Top             =   3300
            Visible         =   0   'False
            Width           =   6015
            Begin MSDataListLib.DataCombo DtcVendedorDestino 
               Height          =   360
               Left            =   1200
               TabIndex        =   44
               Top             =   480
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   635
               _Version        =   393216
               Appearance      =   0
               ForeColor       =   8388608
               Text            =   "DataCombo1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Bahnschrift SemiLight SemiConde"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   350
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VitekeySoft.ChameleonBtn cmdTransferirCliente 
               Height          =   375
               Left            =   3840
               TabIndex        =   46
               Top             =   960
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "PROCESAR"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Bahnschrift SemiLight Condensed"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   350
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
               MICON           =   "frmTarjetas.frx":1E30F
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Image img_cerrar 
               Height          =   240
               Left            =   5520
               Picture         =   "frmTarjetas.frx":1E32B
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "VENDEDOR :"
               BeginProperty Font 
                  Name            =   "Bahnschrift SemiLight SemiConde"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   350
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   240
               TabIndex        =   45
               Top             =   480
               Width           =   885
            End
            Begin VB.Shape Shape7 
               BorderColor     =   &H00C0C0C0&
               BorderWidth     =   3
               FillColor       =   &H00FFFFFF&
               FillStyle       =   0  'Solid
               Height          =   1335
               Left            =   120
               Shape           =   4  'Rounded Rectangle
               Top             =   120
               Width           =   5655
            End
         End
         Begin VB.Frame frmactividad_estado 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   600
            Left            =   -71400
            TabIndex        =   40
            Top             =   1455
            Visible         =   0   'False
            Width           =   6135
            Begin VitekeySoft.ChameleonBtn cmdCambioEstado 
               Height          =   375
               Left            =   4320
               TabIndex        =   42
               Top             =   100
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   661
               BTYPE           =   3
               TX              =   "PROCESAR"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Bahnschrift SemiLight Condensed"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   350
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
               MICON           =   "frmTarjetas.frx":211CF
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSDataListLib.DataCombo DtcEstadoActividad 
               Height          =   360
               Left            =   360
               TabIndex        =   41
               Top             =   100
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   635
               _Version        =   393216
               Appearance      =   0
               ForeColor       =   8388608
               Text            =   "DataCombo1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Bahnschrift SemiLight SemiConde"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   350
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Shape Shape6 
               BorderColor     =   &H00C0C0C0&
               FillColor       =   &H00C0C0C0&
               FillStyle       =   0  'Solid
               Height          =   555
               Left            =   240
               Shape           =   4  'Rounded Rectangle
               Top             =   0
               Width           =   5775
            End
         End
         Begin VB.CheckBox chk_cambio_estado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "CAMBIAR DE ESTADO :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiCondensed"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   450
            Left            =   -73440
            TabIndex        =   39
            Top             =   1500
            Width           =   2055
         End
         Begin VB.TextBox txtActividad 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   840
            Left            =   -73440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   555
            Width           =   8040
         End
         Begin VB.CheckBox chk_asignado 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Caption         =   "ASIGNADO A :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   480
            TabIndex        =   27
            Top             =   4380
            Width           =   1455
         End
         Begin VB.TextBox txtRucEmpresa 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   2040
            TabIndex        =   9
            Top             =   1275
            Width           =   1440
         End
         Begin VB.TextBox txtCelular 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   2040
            TabIndex        =   15
            Top             =   2220
            Width           =   5760
         End
         Begin VB.TextBox TxtRazonSocial 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   3495
            TabIndex        =   10
            Top             =   1275
            Width           =   4320
         End
         Begin VB.TextBox txtDniApoderado 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   2040
            TabIndex        =   11
            Top             =   1740
            Width           =   1440
         End
         Begin VB.TextBox txtRazonSocialApoderado 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   360
            Left            =   3495
            TabIndex        =   13
            Top             =   1740
            Width           =   4320
         End
         Begin MSDataListLib.DataCombo DtcRubro 
            Height          =   360
            Left            =   2040
            TabIndex        =   19
            Top             =   3780
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VitekeySoft.ChameleonBtn cmdRegistrar 
            Height          =   780
            Left            =   9665
            TabIndex        =   21
            Top             =   6780
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1376
            BTYPE           =   5
            TX              =   "PROCESAR"
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
            MICON           =   "frmTarjetas.frx":211EB
            PICN            =   "frmTarjetas.frx":21207
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdTransferir 
            Height          =   780
            Left            =   10530
            TabIndex        =   12
            Top             =   6780
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1376
            BTYPE           =   5
            TX              =   "TRANSFERIR"
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
            MICON           =   "frmTarjetas.frx":2484F
            PICN            =   "frmTarjetas.frx":2486B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker DtpFecha 
            Height          =   375
            Left            =   2040
            TabIndex        =   14
            Top             =   780
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   11.25
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   186974209
            CurrentDate     =   44584
         End
         Begin MSDataListLib.DataCombo DtcVendedor 
            Height          =   360
            Left            =   2040
            TabIndex        =   18
            Top             =   4380
            Visible         =   0   'False
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcMedio 
            Height          =   360
            Left            =   2040
            TabIndex        =   17
            Top             =   3300
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfActividades 
            Height          =   4575
            Left            =   -74640
            TabIndex        =   29
            Top             =   2580
            Width           =   12255
            _ExtentX        =   21616
            _ExtentY        =   8070
            _Version        =   393216
            ForeColor       =   8388608
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   8388608
            BackColorBkg    =   16777215
            GridColor       =   12582912
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VitekeySoft.ChameleonBtn cmdAgregarActividad 
            Height          =   795
            Left            =   -64680
            TabIndex        =   31
            Top             =   900
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   1402
            BTYPE           =   5
            TX              =   "GRABAR"
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
            MICON           =   "frmTarjetas.frx":27155
            PICN            =   "frmTarjetas.frx":27171
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DtcEstado 
            Height          =   360
            Left            =   3600
            TabIndex        =   36
            Top             =   4980
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DtcPlan 
            Height          =   360
            Left            =   3600
            TabIndex        =   37
            Top             =   5460
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VitekeySoft.ChameleonBtn cmdregistrar_sinruc 
            Height          =   375
            Left            =   7920
            TabIndex        =   87
            Top             =   850
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "REGISTRAR SIN RUC"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight Condensed"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
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
            MICON           =   "frmTarjetas.frx":29A5B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdNotificar 
            Height          =   375
            Left            =   -67080
            TabIndex        =   88
            Top             =   2100
            Visible         =   0   'False
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "PROCESAR"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight Condensed"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
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
            MICON           =   "frmTarjetas.frx":29A77
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DtcSupervidor 
            Height          =   360
            Left            =   -71040
            TabIndex        =   89
            Top             =   2100
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VitekeySoft.ChameleonBtn cmdagenda_persona 
            Height          =   780
            Left            =   11520
            TabIndex        =   90
            Top             =   6780
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   1376
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
            MICON           =   "frmTarjetas.frx":29A93
            PICN            =   "frmTarjetas.frx":29AAF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdCerrarNotificaciones 
            Height          =   375
            Left            =   -66120
            TabIndex        =   96
            Top             =   2100
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "CERRAR"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight Condensed"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
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
            MICON           =   "frmTarjetas.frx":29DC9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfSoporteTecnico 
            Height          =   6255
            Left            =   -74760
            TabIndex        =   97
            Top             =   840
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   11033
            _Version        =   393216
            ForeColor       =   8388608
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   8388608
            BackColorBkg    =   16777215
            GridColor       =   12582912
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRIORIDAD :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiCondensed"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   2520
            TabIndex        =   94
            Top             =   6780
            Width           =   915
         End
         Begin VB.Shape Shape15 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   1140
            Left            =   2040
            Shape           =   4  'Rounded Rectangle
            Top             =   6420
            Width           =   6855
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COMPROBANTE :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiCondensed"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   2280
            TabIndex        =   91
            Top             =   6660
            Width           =   1230
         End
         Begin VB.Label lblNotificacion 
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            Caption         =   "CLIENTE YA REGISTRADO."
            BeginProperty Font 
               Name            =   "Bahnschrift SemiCondensed"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   520
            Left            =   7920
            TabIndex        =   75
            Top             =   1260
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label lblhora 
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   3600
            TabIndex        =   57
            Top             =   780
            Width           =   2775
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-MAIL :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1080
            TabIndex        =   47
            Top             =   2700
            Width           =   630
         End
         Begin VB.Label lblcomprobante 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Bahnschrift SemiCondensed"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   3720
            TabIndex        =   38
            Top             =   5940
            Width           =   3870
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ESTADO TARJETA :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiCondensed"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   2145
            TabIndex        =   35
            Top             =   4980
            Width           =   1365
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COMPROBANTE :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiCondensed"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   2280
            TabIndex        =   34
            Top             =   5940
            Width           =   1230
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PLAN CONTRADO :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiCondensed"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   2160
            TabIndex        =   33
            Top             =   5460
            Width           =   1350
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ESTADO TARJETA :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   480
            TabIndex        =   32
            Top             =   5460
            Width           =   1365
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00800000&
            FillStyle       =   0  'Solid
            Height          =   1500
            Left            =   2040
            Shape           =   4  'Rounded Rectangle
            Top             =   4905
            Width           =   6855
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ACTIVIDAD :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   -74640
            TabIndex        =   28
            Top             =   780
            Width           =   870
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   6915
            Left            =   -74880
            Top             =   400
            Width           =   12735
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA REGISTRO :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   435
            TabIndex        =   26
            Top             =   780
            Width           =   1365
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RUC EMPRESA :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   615
            TabIndex        =   25
            Top             =   1275
            Width           =   1185
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DNI APODERADO :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   480
            TabIndex        =   24
            Top             =   1740
            Width           =   1320
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CELULAR :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   1020
            TabIndex        =   23
            Top             =   2220
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RUBRO NEGOCIO :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   480
            TabIndex        =   22
            Top             =   3780
            Width           =   1320
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MEDIO DE INGRESO :"
            BeginProperty Font 
               Name            =   "Bahnschrift SemiLight SemiConde"
               Size            =   9.75
               Charset         =   0
               Weight          =   350
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   285
            TabIndex        =   20
            Top             =   3300
            Width           =   1515
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   6930
            Left            =   240
            Top             =   705
            Width           =   12495
         End
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   13245
         Picture         =   "frmTarjetas.frx":29DE5
         Top             =   165
         Width           =   240
      End
   End
   Begin VB.TextBox txtBusquedaEmpresa 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   18720
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "NUEVA "
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
      MICON           =   "frmTarjetas.frx":2CC89
      PICN            =   "frmTarjetas.frx":2CCA5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   7935
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   18255
      _ExtentX        =   32200
      _ExtentY        =   13996
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   12582912
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
   Begin VitekeySoft.ChameleonBtn cmdReporte 
      Height          =   855
      Left            =   18720
      TabIndex        =   3
      Top             =   3920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "REPORTE"
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
      MICON           =   "frmTarjetas.frx":2D0F7
      PICN            =   "frmTarjetas.frx":2D113
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrar 
      Height          =   975
      Left            =   18720
      TabIndex        =   4
      Top             =   8160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "CERRAR"
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
      MICON           =   "frmTarjetas.frx":2F55D
      PICN            =   "frmTarjetas.frx":2F579
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdVisualizar 
      Height          =   855
      Left            =   18720
      TabIndex        =   5
      Top             =   2115
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTarjetas.frx":325A0
      PICN            =   "frmTarjetas.frx":325BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcVendedorBusqueda 
      Height          =   360
      Left            =   5160
      TabIndex        =   52
      Top             =   600
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   345
      Left            =   15760
      TabIndex        =   53
      Top             =   375
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   11.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   186974209
      CurrentDate     =   44584
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   345
      Left            =   17280
      TabIndex        =   54
      Top             =   380
      Width           =   1420
      _ExtentX        =   2514
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   11.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   186974209
      CurrentDate     =   44584
   End
   Begin VitekeySoft.ChameleonBtn cmdBuscar 
      Height          =   390
      Left            =   18840
      TabIndex        =   56
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   688
      BTYPE           =   3
      TX              =   "BUSCAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTarjetas.frx":328D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcBusquedaEstado 
      Height          =   360
      Left            =   10440
      TabIndex        =   60
      Top             =   210
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcBusquedaMedio 
      Height          =   360
      Left            =   10440
      TabIndex        =   62
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VitekeySoft.ChameleonBtn cmdNotificaciones 
      Height          =   375
      Left            =   13680
      TabIndex        =   64
      Top             =   525
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight Condensed"
         Size            =   9.75
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
      MICON           =   "frmTarjetas.frx":328F2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdAsignacionAutomatica 
      Height          =   855
      Left            =   18720
      TabIndex        =   70
      Top             =   3020
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "ASIGNACION"
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
      MICON           =   "frmTarjetas.frx":3290E
      PICN            =   "frmTarjetas.frx":3292A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdAgenda 
      Height          =   855
      Left            =   18720
      TabIndex        =   76
      Top             =   7260
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "AGENDA ERP"
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
      MICON           =   "frmTarjetas.frx":3666B
      PICN            =   "frmTarjetas.frx":36687
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTIFICACIONES"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   13920
      TabIndex        =   63
      Top             =   240
      Width           =   1230
   End
   Begin VB.Shape Shape10 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   795
      Left            =   13620
      Shape           =   4  'Rounded Rectangle
      Top             =   180
      Width           =   1935
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MEDIO  :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   9720
      TabIndex        =   61
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESTADO :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   9660
      TabIndex        =   59
      Top             =   240
      Width           =   675
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   795
      Left            =   9600
      Shape           =   4  'Rounded Rectangle
      Top             =   180
      Width           =   3975
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00808080&
      Height          =   615
      Left            =   15600
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CELULAR :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4320
      TabIndex        =   51
      Top             =   240
      Width           =   780
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "APODERADO :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   49
      Top             =   600
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   555
      TabIndex        =   6
      Top             =   240
      Width           =   825
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   900
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   19695
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "frmTarjetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub chk_asignado_Click()

If Me.chk_asignado.Value = 1 Then
    Call load_vendedores
    Me.DtcVendedor.Visible = True
Else
    Me.DtcVendedor.Visible = False
End If

End Sub

Private Sub chk_cambio_estado_Click()

If Me.chk_cambio_estado.Value = 1 Then
    Me.frmactividad_estado.Visible = True
Else
    Me.frmactividad_estado.Visible = False
End If

End Sub

Private Sub chk_notificacion_Click()

If Me.chk_notificacion.Value = 1 Then
    
    
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_cargo='00004' and  id_personal='si' and  ruc='" & KEY_RUC & "' ORDER BY nombre_completo"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcSupervidor)
    Me.DtcSupervidor.Visible = True
    Me.cmdNotificar.Visible = True
    Me.cmdCerrarNotificaciones.Visible = True
 Else
    Me.DtcSupervidor.Visible = False
    Me.cmdNotificar.Visible = False
    Me.cmdCerrarNotificaciones.Visible = False
    
End If

End Sub

Private Sub cmdAgendar_Click()

 strCadena = "CALL ADM_tarjeta_cliente('31','" & Val(Me.HfAgenda.TextMatrix(Me.HfAgenda.Row, 0)) & "','" & Trim(Me.txtRucEmpresa.Text) & "','3','4','5','6','7','8','9','" & Me.DtcVendedorAgenda.BoundText & "','11','12','13','14','15','16','" & Format(Me.MonthView1.Value, "YYYY-mm-dd") & "','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
 Call ConfiguraRst(strCadena)
 in_mensaje = "AGENDA :" + Me.HfAgenda.TextMatrix(Me.HfAgenda.Row, 1)
 Call put_actividad(Me.frmtarjeta.Tag, in_mensaje)
   
 Call llenarAgenda_detalle(Me.HfAgenda, Me.HfDias.TextMatrix(Me.HfDias.Row, 1), Me.HfDias.TextMatrix(Me.HfDias.Row, 0))




End Sub

Private Sub cmdAgenda_Click()



If KEY_CARGO = "00004" Then
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and  ruc='" & KEY_RUC & "' ORDER BY nombre_completo"
    in_vendedor = KEY_USUARIO
Else
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE dni='" & KEY_USUARIO & "' and  id_personal='si' and  ruc='" & KEY_RUC & "'"
    in_vendedor = KEY_USUARIO
End If
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcVendedorAgenda)

Me.DtcVendedorAgenda.BoundText = in_vendedor
Call Me.llenarAgenda(Me.HfDias, Me.DtcVendedorAgenda.BoundText)
Me.frm_agenda.Visible = True

End Sub

Private Sub cmdagenda_persona_Click()

If KEY_CARGO = "00004" Then
    
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and  ruc='" & KEY_RUC & "' ORDER BY nombre_completo"
    
Else
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE dni='" & KEY_USUARIO & "' and  id_personal='si' and  ruc='" & KEY_RUC & "'"
End If
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcVendedorAgenda)

Me.DtcVendedorAgenda.BoundText = Me.DtcVendedor.BoundText
Call Me.llenarAgenda(Me.HfDias, Me.DtcVendedorAgenda.BoundText)
Me.frm_agenda.Visible = True



    
    


End Sub

Private Sub cmdAgregarActividad_Click()

If Me.txtActividad.Text <> "" Then
    Call put_actividad(Me.frmtarjeta.Tag, Trim(Me.txtActividad.Text))
    Me.txtActividad.Text = ""
    Call Resalta(Me.txtActividad)
End If

End Sub

Private Sub cmdAnularAgenda_Click()

 strCadena = "CALL ADM_tarjeta_cliente('32','" & Val(Me.HfAgenda.TextMatrix(Me.HfAgenda.Row, 0)) & "','" & Trim(Me.txtRucEmpresa.Text) & "','3','4','5','6','7','8','9','" & Me.DtcVendedorAgenda.BoundText & "','11','12','13','14','15','16','" & Format(Me.MonthView1.Value, "YYYY-mm-dd") & "','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
 Call ConfiguraRst(strCadena)
 in_mensaje = "ANULACION AGENDA  :" + Me.HfAgenda.TextMatrix(Me.HfAgenda.Row, 1)
 Call put_actividad(Me.frmtarjeta.Tag, in_mensaje)
 
 Call llenarAgenda_detalle(Me.HfAgenda, Me.HfDias.TextMatrix(Me.HfDias.Row, 1), Me.HfDias.TextMatrix(Me.HfDias.Row, 0))

End Sub

Private Sub cmdAsignacion_clientes_Click()

strCadena = "CALL ADM_tarjeta_cliente('24','','','','','','','','','','','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
Call llenarVendedor(Me.HfVendedores)
  

strCadena = "CALL ADM_tarjeta_cliente('23','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   Call llenarVendedor(Me.HfVendedores)
   rstA.MoveFirst
   Me.prg_avance.Min = 0
   Me.prg_avance.Max = rstA.RecordCount
   For i = 0 To rstA.RecordCount - 1
        strCadena = "CALL ADM_tarjeta_cliente('25','" & rstA("id") & "','','','','','','','','','','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
    
 
        Call llenarVendedor(Me.HfVendedores)
        rstA.MoveNext
        DoEvents
        Me.prg_avance.Value = i + 1
        DoEvents
   Next i
End If
MsgBox "ASIGNACION ALEATORIA CULMINADA.", vbInformation
Call llenarVendedor(Me.HfVendedores)


End Sub

Private Sub cmdAsignacionAutomatica_Click()
Me.frm_impresion.Visible = False
Me.frmAsignacion.Visible = True
Call llenarVendedor(Me.HfVendedores)

End Sub

Private Sub cmdBuscar_Click()
If KEY_CARGO = "00052" Then
    in_operacion = "27"
Else
    in_operacion = "7"
End If

strCadena = "CALL ADM_tarjeta_cliente('" & in_operacion & "','','" & Trim(Me.txtRucEmpresa.Text) & "','" & Trim(Me.TxtRazonSocial.Text) & "','" & Trim(Me.txtDniApoderado.Text) & "'," & _
"'" & Trim(Me.txtRazonSocialApoderado.Text) & "','" & Trim(Me.txtCelular.Text) & "','" & Trim(Me.txtEmail.Text) & "','" & Me.DtcMedio.BoundText & "'," & _
"'" & Me.DtcRubro.BoundText & "','" & KEY_USUARIO & "','11','12','13','14','15','16','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call Me.llenarGrid(Me.HfgDetalle)

End Sub

Private Sub cmdCambioEstado_Click()

  
strCadena = "CALL ADM_tarjeta_cliente('6','" & Val(Me.frmtarjeta.Tag) & "','2','3','4'," & _
"'5','6','7','8','9','10','" & Me.DtcEstadoActividad.BoundText & "','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)

in_actividad = "CAMBIO DE ESTADO :" + Me.DtcEstadoActividad.Text
Call put_actividad(Me.frmtarjeta.Tag, in_actividad)


 




End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdCerrarTarjeta_Click()


End Sub

Private Sub cmdCerrarNotificaciones_Click()
 strCadena = "CALL ADM_tarjeta_cliente('38','" & Val(Me.frmtarjeta.Tag) & "','2','3','4'," & _
    "'5','6','7','8','9','" & Me.DtcSupervidor.BoundText & "','11','12','13','si','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)

    in_actividad = "CERRO NOTIFICACIONES "
    Call put_actividad(Me.frmtarjeta.Tag, in_actividad)
    
End Sub

Private Sub cmdGenerarAgenda_Click()
 
 
 strCadena = "CALL ADM_tarjeta_cliente('30','1','" & Trim(Me.txtRucEmpresa.Text) & "','3','4','5','6','7','8','9','" & Me.DtcVendedorAgenda.BoundText & "','11','12','13','14','15','16','" & Format(Me.MonthView1.Value, "YYYY-mm-dd") & "','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
 Call ConfiguraRst(strCadena)
 Call Me.llenarAgenda(Me.HfDias, Me.DtcVendedorAgenda.BoundText)
 
 
End Sub

Private Sub cmdNotificaciones_Click()

strCadena = "CALL ADM_tarjeta_cliente('14','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call Me.llenarGrid(Me.HfgDetalle)


End Sub

Private Sub cmdNotificar_Click()
    
    strCadena = "CALL ADM_tarjeta_cliente('15','" & Val(Me.frmtarjeta.Tag) & "','2','3','4'," & _
    "'5','6','7','8','9','" & Me.DtcSupervidor.BoundText & "','11','12','13','si','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)

    in_actividad = "NOTIFICACION A  :" + Me.DtcSupervidor.Text
    Call put_actividad(Me.frmtarjeta.Tag, in_actividad)

End Sub

Private Sub cmdnuevo_Click()
Call nuevo

End Sub

Private Sub load_vendedores()
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and  ruc='" & KEY_RUC & "' ORDER BY nombre_completo"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcVendedor)
End Sub

Private Sub nuevo()

Me.txtRucEmpresa.Text = ""
Me.TxtRazonSocial.Text = ""
Me.txtDniApoderado.Text = ""
Me.txtRazonSocialApoderado.Text = ""
Me.txtCelular.Text = ""
Me.txtEmail.Text = ""
Me.DtcRubro.BoundText = "00034"
Me.DtcEstado.BoundText = "01"
Me.DtcEstado.Enabled = False
Me.chk_asignado.Value = 0
Me.DtcVendedor.Visible = False
Me.DtcPlan.Visible = False
Me.frmtarjeta.Visible = True
Me.SSTab1.Tab = 0
Me.frmtarjeta.Tag = 0
Me.txtBuscarProducto.Visible = False
Me.lblNotificacion.Visible = False
Me.lblcomprobante.Caption = ""
Me.cmdregistrar_sinruc.Visible = True
Call Resalta(Me.txtRucEmpresa)

End Sub
Private Sub put_actividad(ByVal in_tarjeta As String, ByVal in_actividad As String)




strCadena = "CALL ADM_tarjeta_cliente('4','" & Val(in_tarjeta) & "','2','3','4'," & _
"'5','6','7','8','9','10','" & Me.DtcEstado.BoundText & "','12','13','14','15','" & Trim(in_actividad) & "','17','18','19','" & in_servicio & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)

Call llenarActividad(Me.HfActividades, in_tarjeta)

End Sub
Public Sub llenarActividad(ByVal Grilla As MSHFlexGrid, ByVal in_tarjeta As String)
On Error GoTo salir

strCadena = "CALL ADM_tarjeta_cliente('5','" & Val(in_tarjeta) & "','2','3','4'," & _
"'5','6','7','8','9','10','11','12','13','14','15','" & Trim(Me.txtActividad.Text) & "','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
   
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 900
           Grilla.ColWidth(3) = 7000
           Grilla.ColWidth(4) = 2500
           
           
        cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "ACTIVIDAD" & vbTab & "OPERADOR "
        Grilla.AddItem cabecera
         For k = 1 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            
            Fila = rst("id") & vbTab & Format(rst("fecha"), "dd-mm-YYYY") & vbTab & Format(rst("hora"), "HH:mm:ss") & vbTab & rst("actividad") & vbTab & rst("operador")
            Grilla.AddItem Fila
            rst.MoveNext
    Next i
    
    
   
    Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Public Sub llenarVendedor(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir

strCadena = "CALL ADM_tarjeta_cliente('21','" & Val(in_tarjeta) & "','2','3','4'," & _
"'5','6','7','8','9','10','11','12','13','14','15','" & Trim(Me.txtActividad.Text) & "','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
   
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
           Grilla.ColWidth(0) = 400
           Grilla.ColWidth(1) = 2500
           Grilla.ColWidth(2) = 500
           
        cabecera = "ID" & vbTab & "NOMBRE VENDEDOR" & vbTab & "ITEM"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            
            Fila = Format(i + 1, "00") & vbTab & rst("nombre_completo") & vbTab & rst("registros")
            Grilla.AddItem Fila
            rst.MoveNext
    Next i
    
    
   
    Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Private Sub cmdRegistrar_Click()



If Val(Me.frmtarjeta.Tag) < 1 Then
    strCadena = "CALL ADM_tarjeta_cliente('12','1','" & Trim(Me.txtRucEmpresa.Text) & "','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    If rst(0) > 0 Then
        MsgBox "RUC :" + Trim(Me.txtRucEmpresa.Text) + Space(2) + "YA REGISTRADA" + Chr(13) + "REALICE UNA BUSQUEDA", vbInformation
        Me.txtBusquedaEmpresa.Text = Trim(Me.txtRucEmpresa.Text)
        Exit Sub
    End If

End If



'REGISTRO
If Me.chk_asignado.Value = 0 Then
   in_vendedor = "0"
Else
   in_vendedor = Me.DtcVendedor.BoundText
End If

in_prioridad = 3
If Me.opt_interesado.Value = True Then
   in_prioridad = 1
End If
If Me.opt_demostracion.Value = True Then
   in_prioridad = 2
End If

strCadena = "CALL ADM_tarjeta_cliente('1','" & Val(Me.frmtarjeta.Tag) & "','" & Trim(Me.txtRucEmpresa.Text) & "','" & Trim(Me.TxtRazonSocial.Text) & "','" & Trim(Me.txtDniApoderado.Text) & "'," & _
"'" & Trim(Me.txtRazonSocialApoderado.Text) & "','" & Trim(Me.txtCelular.Text) & "','" & Trim(Me.txtEmail.Text) & "','" & Me.DtcMedio.BoundText & "'," & _
"'" & Me.DtcRubro.BoundText & "','" & in_vendedor & "','" & Me.DtcEstado.BoundText & "','" & Me.DtcPlan.BoundText & "','13','14','15','16','17','18','" & in_prioridad & "','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)

'REGISTRO ACTIVIDAD
in_actividad = "REGISTRO DE TARJETA :"
Call put_actividad(rst(0), in_actividad)


'LISTADO
If KEY_CARGO = "00052" Then
    in_operacion = "26"
Else
    in_operacion = "2"
End If

strCadena = "CALL ADM_tarjeta_cliente('" & in_operacion & "','','" & Trim(Me.txtRucEmpresa.Text) & "','" & Trim(Me.TxtRazonSocial.Text) & "','" & Trim(Me.txtDniApoderado.Text) & "'," & _
"'" & Trim(Me.txtRazonSocialApoderado.Text) & "','" & Trim(Me.txtCelular.Text) & "','" & Trim(Me.txtEmail.Text) & "','" & Me.DtcMedio.BoundText & "'," & _
"'" & Me.DtcRubro.BoundText & "','" & Me.DtcVendedor.BoundText & "','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"

Me.frmtarjeta.Visible = False
Me.HfgDetalle.Enabled = True
Call Me.llenarGrid(Me.HfgDetalle)




End Sub
Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir


Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
   
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 550
           Grilla.ColWidth(2) = 1700
           Grilla.ColWidth(3) = 1100
           Grilla.ColWidth(4) = 3200
           Grilla.ColWidth(5) = 2600
           Grilla.ColWidth(6) = 1000
           Grilla.ColWidth(7) = 2000
           Grilla.ColWidth(8) = 1800
           Grilla.ColWidth(9) = 2500
           Grilla.ColWidth(10) = 1500
           
           
        cabecera = "CODIGO" & vbTab & "ITEM" & vbTab & "FECHA" & vbTab & "RUC EMPRESA " & vbTab & "EMPRESA " & vbTab & "CONTACTO" & vbTab & "CELULAR" & vbTab & "RUBRO" & vbTab & "MEDIO" & vbTab & "VENDEDOR" & vbTab & "ESTADO"
        Grilla.AddItem cabecera
         For k = 1 To 10
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            
            Fila = rst("id") & vbTab & Format(i + 1, "000") & vbTab & Format(rst("fecha_registro"), "dd-mm-YYYY") + Space(2) + Format(rst("hora_registro"), "HH:mm:ss") & vbTab & rst("ruc_empresa") & vbTab & rst("razon_social") & vbTab & rst("apoderado") & vbTab & rst("celular") & vbTab & rst("rubro") & vbTab & rst("medio") & vbTab & rst("vendedor") & vbTab & rst("estado")
            Grilla.AddItem Fila
            
            Grilla.col = 10
            Grilla.Row = i + 1
            Select Case rst("prioridad")
                Case 1
                    Grilla.CellBackColor = &H80FF80
                Case 2
                    Grilla.CellBackColor = &H80FF&
            End Select
            
            
            
            
            rst.MoveNext
    Next i
    
    Grilla.ColAlignment(2) = 1
  
   
    Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Private Sub cmdregistrar_sinruc_Click()
strCadena = "CALL ADM_tarjeta_cliente('35','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtRucEmpresa.Text = rst("dni")
   Me.TxtRazonSocial.Text = rst("nombre_completo")
End If

End Sub

Private Sub cmdReporte_Click()

Me.frm_impresion.Visible = True
        
 

End Sub

Private Sub cmdReporteEstado_Click()
Dim cam3(0 To 5, 1 To 5)  As String
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(3, 1) = "empresa"
    cam3(4, 1) = "direccion"
    cam3(5, 1) = "titulo"
    
    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    cam3(2, 2) = in_almacen
    cam3(3, 2) = KEY_EMPRESA
    cam3(4, 2) = KEY_DIRECCION_ALM
    cam3(5, 2) = "LISTADO DE TARJETAS"
    param = cam3()
    

strCadena = "CALL ADM_tarjeta_cliente('16','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "Rpttarjetas", param, App.Path + "\Reportes\")
Exit Sub
    
End Sub

Private Sub cmdReporteVendedor_Click()
Dim cam3(0 To 5, 1 To 5)  As String
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(3, 1) = "empresa"
    cam3(4, 1) = "direccion"
    cam3(5, 1) = "titulo"
    
    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    cam3(2, 2) = in_almacen
    cam3(3, 2) = KEY_EMPRESA
    cam3(4, 2) = KEY_DIRECCION_ALM
    cam3(5, 2) = "LISTADO DE TARJETAS"
    param = cam3()
    
 If Me.chk_vendedor.Value = 1 Then
    in_vendedor = Me.DtcVendedorBusqueda.BoundText
 Else
    in_vendedor = ""
 End If

strCadena = "CALL ADM_tarjeta_cliente('19','1','2','3','4','5','6','7','8','9','" & in_vendedor & "','11','12','13','14','15','16','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RpttarjetasVendedor", param, App.Path + "\Reportes\")
Exit Sub

End Sub

Private Sub cmdTransferir_Click()
Me.frmtransferir.Visible = True


End Sub

Private Sub cmdTransferirCliente_Click()

'REGISTRO

in_vendedor = DtcVendedorDestino.BoundText
in_actividad = "CLIENTE TRANSFERIDO -> " + Me.DtcVendedorDestino.Text

strCadena = "CALL ADM_tarjeta_cliente('1','" & Val(Me.frmtarjeta.Tag) & "','" & Trim(Me.txtRucEmpresa.Text) & "','" & Trim(Me.TxtRazonSocial.Text) & "','" & Trim(Me.txtDniApoderado.Text) & "'," & _
"'" & Trim(Me.txtRazonSocialApoderado.Text) & "','" & Trim(Me.txtCelular.Text) & "','" & Trim(Me.txtEmail.Text) & "','" & Me.DtcMedio.BoundText & "'," & _
"'" & Me.DtcRubro.BoundText & "','" & in_vendedor & "','" & Me.DtcEstado.BoundText & "','" & Me.DtcPlan.BoundText & "','13','14','15','" & in_actividad & "','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)

in_actividad = "TRANSFERIDO A :" + Me.DtcVendedorDestino.Text

Call put_actividad(Me.frmtarjeta.Tag, in_actividad)


'LISTADO

strCadena = "CALL ADM_tarjeta_cliente('2','','" & Trim(Me.txtRucEmpresa.Text) & "','" & Trim(Me.TxtRazonSocial.Text) & "','" & Trim(Me.txtDniApoderado.Text) & "'," & _
"'" & Trim(Me.txtRazonSocialApoderado.Text) & "','" & Trim(Me.txtCelular.Text) & "','" & Trim(Me.txtEmail.Text) & "','" & Me.DtcMedio.BoundText & "'," & _
"'" & Me.DtcRubro.BoundText & "','" & Me.DtcVendedor.BoundText & "','10','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"

Me.frmtarjeta.Visible = False
Me.HfgDetalle.Enabled = True
Call Me.llenarGrid(Me.HfgDetalle)



End Sub

Private Sub cmdVerAgenda_Click()
 Call Me.llenarAgenda(Me.HfDias, Me.DtcVendedorAgenda.BoundText)
End Sub

Private Sub cmdVisualizar_Click()

Me.frmtarjeta.Tag = Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0))
Me.HfgDetalle.Enabled = False

strCadena = "CALL ADM_tarjeta_cliente('3','" & Val(Me.frmtarjeta.Tag) & "','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.DtpFecha.Value = rst("fecha_registro")
    Me.lblhora.Caption = rst("hora_registro")
    Me.txtRucEmpresa.Text = rst("ruc_empresa")
    Me.TxtRazonSocial.Text = rst("razon_social")
    Me.txtDniApoderado.Text = rst("dni_apoderado")
    Me.txtRazonSocialApoderado.Text = rst("apoderado")
    Me.txtCelular.Text = rst("celular")
    Me.txtEmail.Text = rst("mail")
    Me.DtcMedio.BoundText = rst("id_medio")
    Me.DtcRubro.BoundText = rst("id_rubro")
    If rst("id_vendedor") = "0" Then
       Me.chk_asignado.Value = 0
    Else
        Me.chk_asignado.Value = 1
        Me.DtcVendedor.BoundText = rst("id_vendedor")
    End If
    Me.DtcPlan.Visible = True
    Me.DtcEstado.Enabled = True
    Me.DtcEstado.BoundText = rst("id_estado")
    Me.DtcEstadoActividad.BoundText = rst("id_estado")
    Me.DtcPlan.BoundText = rst("id_producto")
    Me.lblcomprobante.Caption = rst("comprobante")
    Call get_prioridad(rst("prioridad"))
    Me.txtBuscarProducto.Visible = True
    Me.SSTab1.Tab = 0
    Me.cmdregistrar_sinruc.Visible = False
    Me.frmtarjeta.Visible = True
End If

End Sub

Private Sub get_prioridad(ByVal in_prioridad As Integer)

Select Case in_prioridad
    Case 1
        Me.opt_interesado.Value = True
    Case 2
        Me.opt_demostracion.Value = True
    Case 3
        Me.opt_poratender.Value = True
End Select


End Sub

Private Sub DtcBusquedaEstado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "CALL ADM_tarjeta_cliente('17','','','',''," & _
    "'','','','','','','" & Me.DtcBusquedaEstado.BoundText & "','12','13','14','15','16','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call Me.llenarGrid(Me.HfgDetalle)
End If


End Sub

Private Sub DtcBusquedaMedio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "CALL ADM_tarjeta_cliente('18','','','',''," & _
    "'','','','" & Me.DtcBusquedaMedio.BoundText & "','','','','12','13','14','15','16','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call Me.llenarGrid(Me.HfgDetalle)
End If
End Sub

Private Sub DtcVendedorBusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "CALL ADM_tarjeta_cliente('11','','" & Trim(Me.txtBusquedaEmpresa.Text) & "','" & Trim(Me.txtBusquedaEmpresa.Text) & "',''," & _
    "'" & Trim(Me.txtBusquedaApoderado.Text) & "','" & Trim(Me.txtBusquedaCelular.Text) & "','','','','" & Me.DtcVendedorBusqueda.BoundText & "','11','12','13','14','15','16','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call Me.llenarGrid(Me.HfgDetalle)
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA


strCadena = "SELECT id_medio as Codigo,descripcion as Descripcion FROM empresa_medio_contacto ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMedio)

'-------------------------------
strCadena = "SELECT codigo as Codigo,descripcion as Descripcion FROM persona_rubro WHERE activo='si' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcRubro)

'-------------------------------
strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM estado_tarjeta ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstado)
'-------------------------------
strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM estado_tarjeta ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcBusquedaEstado)

'-------------------------------
strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM estado_tarjeta ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstadoActividad)
'-------------------------------
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and  ruc='" & KEY_RUC & "' ORDER BY nombre_completo"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcVendedorBusqueda)
'---------------------------------------
strCadena = "SELECT id_medio as Codigo,descripcion as Descripcion FROM empresa_medio_contacto ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcBusquedaMedio)


If KEY_CARGO = "00052" Then
    in_operacion = "26"
    Me.cmdAsignacionAutomatica.Enabled = False
Else
    in_operacion = "2"
    Me.cmdAsignacionAutomatica.Enabled = True
End If




strCadena = "CALL ADM_tarjeta_cliente('" & in_operacion & "','','" & Trim(Me.txtRucEmpresa.Text) & "','" & Trim(Me.TxtRazonSocial.Text) & "','" & Trim(Me.txtDniApoderado.Text) & "'," & _
"'" & Trim(Me.txtRazonSocialApoderado.Text) & "','" & Trim(Me.txtCelular.Text) & "','" & Trim(Me.txtEmail.Text) & "','" & Me.DtcMedio.BoundText & "'," & _
"'" & Me.DtcRubro.BoundText & "','" & KEY_USUARIO & "','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call Me.llenarGrid(Me.HfgDetalle)




End Sub

Private Sub HfAgenda_SelChange()

If Val(Me.HfAgenda.TextMatrix(Me.HfAgenda.Row, 0)) > 0 Then
    Me.cmdAnularAgenda.Enabled = True
    If Len((Me.HfAgenda.TextMatrix(Me.HfAgenda.Row, 2))) > 2 Then
        Me.cmdAgendar.Enabled = False
    Else
        Me.cmdAgendar.Enabled = True
    End If
    
Else
    Me.cmdAnularAgenda.Enabled = False
    Me.cmdAgendar.Enabled = False
End If

End Sub

Private Sub HfDias_SelChange()

Call llenarAgenda_detalle(Me.HfAgenda, Me.HfDias.TextMatrix(Me.HfDias.Row, 1), Me.HfDias.TextMatrix(Me.HfDias.Row, 0))

End Sub

Private Sub Image1_Click()
Me.frmtarjeta.Tag = 0
Me.frmtarjeta.Visible = False
Me.HfgDetalle.Enabled = True
End Sub

Private Sub Image2_Click()
Me.frm_impresion.Visible = False

End Sub

Private Sub Image3_Click()
Me.frmAsignacion.Visible = False
End Sub

Private Sub Image4_Click()
Me.frm_agenda.Visible = False
End Sub

Private Sub img_cerrar_Click()
Me.frmtransferir.Visible = False
End Sub

Private Sub lblNotificacion_Click()

If Me.lblNotificacion.Visible = True Then
    
    strCadena = "CALL ADM_tarjeta_cliente('8','','" & Trim(Me.txtRucEmpresa.Text) & "','',''," & _
    "'','','','','','','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call Me.llenarGrid(Me.HfgDetalle)
    
End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If Me.SSTab1.Tab = 1 Then
    Call Me.llenarActividad(Me.HfActividades, Me.frmtarjeta.Tag)
End If
End Sub

Private Sub timmer_notificacion_Timer()

Call get_notificaciones
Call get_notificacion_agenda


End Sub

Private Sub get_notificaciones()
strCadena = "CALL ADM_tarjeta_cliente('13','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRstP(strCadena)
If rstP(0) > 0 Then
    Me.cmdNotificaciones.Visible = True
    Me.cmdNotificaciones.Caption = "[ " + str(rstP(0)) + "  ]" + Space(2) + "NOTIFICACIONES"
Else
    Me.cmdNotificaciones.Visible = False
End If

End Sub

Private Sub get_notificacion_agenda()
If KEY_CARGO = "00004" Then
    strCadena = "CALL ADM_tarjeta_cliente('33','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Else
    strCadena = "CALL ADM_tarjeta_cliente('34','1','2','3','4','5','6','7','8','9','" & KEY_USUARIO & "','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
End If

Call ConfiguraRstP(strCadena)
If rstP(0) > 0 Then
    Me.cmdAgenda.ForeColor = &HFF&
    Me.cmdAgenda.Caption = "[" + str(rstP(0)) + " ]" + Space(2) + "AGENDAS"
Else
    Me.cmdAgenda.ForeColor = &HC00000
    Me.cmdAgenda.Caption = " AGENDA"
End If

End Sub



Private Sub txtBusquedaApoderado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     strCadena = "CALL ADM_tarjeta_cliente('9','','" & Trim(Me.txtBusquedaEmpresa.Text) & "','" & Trim(Me.txtBusquedaEmpresa.Text) & "',''," & _
    "'" & Trim(Me.txtBusquedaApoderado.Text) & "','" & Trim(Me.txtCelular.Text) & "','','','','','11','12','13','14','15','16','17','18','19','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call Me.llenarGrid(Me.HfgDetalle)
End If

End Sub

Private Sub txtBusquedaCelular_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    strCadena = "CALL ADM_tarjeta_cliente('10','','" & Trim(Me.txtBusquedaEmpresa.Text) & "','" & Trim(Me.txtBusquedaEmpresa.Text) & "',''," & _
    "'" & Trim(Me.txtBusquedaApoderado.Text) & "','" & Trim(Me.txtBusquedaCelular.Text) & "','','','','','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call Me.llenarGrid(Me.HfgDetalle)
End If

End Sub

Private Sub txtBusquedaEmpresa_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    strCadena = "CALL ADM_tarjeta_cliente('8','','" & Trim(Me.txtBusquedaEmpresa.Text) & "','',''," & _
    "'','','','','','','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call Me.llenarGrid(Me.HfgDetalle)
End If

End Sub

Private Sub txtCelular_Change()
Me.lblNotificacion.Visible = False

If Val(Me.frmtarjeta.Tag) < 1 Then
    Call verifica_registro(Trim(Me.txtRucEmpresa.Text), 2)
Else
    Me.cmdRegistrar.Enabled = True
End If
End Sub

Private Sub txtDniApoderado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
buscar_nuevamente:
    strCadena = "SELECT * FROM  persona WHERE  dni='" & Trim(Me.txtDniApoderado.Text) & "' LIMIT 1 "
    Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            If get_dni_reniec_iii(Trim(Me.txtDniApoderado.Text), KEY_DEPARTAMENTO, KEY_PROVINCIA, KEY_DISTRITO) = True Then
                GoTo buscar_nuevamente
            End If
        Else
            Me.txtDniApoderado.Text = rst("dni")
            Me.txtRazonSocialApoderado.Text = UCase(rst("nombre_completo"))
            Call Resalta(Me.txtCelular)
        End If
End If
End Sub

Private Sub txtRucEmpresa_Change()
Me.lblNotificacion.Visible = False

If Val(Me.frmtarjeta.Tag) < 1 Then
    Call verifica_registro(Trim(Me.txtRucEmpresa.Text), 1)
Else
    Me.cmdRegistrar.Enabled = True
End If


End Sub
Private Sub verifica_registro(ByVal in_ruc As String, ByVal opcion As Integer)
If opcion = 1 Then
If Len(in_ruc) > 10 Then
    strCadena = "CALL ADM_tarjeta_cliente('22','','" & Trim(Me.txtRucEmpresa.Text) & "','',''," & _
    "'','','','','','','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount > 0 Then
       Me.lblNotificacion.Height = 710
       Me.lblNotificacion.Caption = "CLIENTE YA REGISTRADO" + Chr(13) + "FECHA: " + Format(rstA("fecha_registro"), "dd-mm-YYYY") + Chr(13) + "VENDEDOR: " + get_persona(rstA("id_vendedor"))
       Me.lblNotificacion.Visible = True
       If Val(Me.frmtarjeta.Tag) < 1 Then
          Me.cmdRegistrar.Enabled = False
       Else
         Me.cmdRegistrar.Enabled = True
       End If
    Else
         Me.cmdRegistrar.Enabled = True
    End If
End If
End If


If opcion = 2 Then

    strCadena = "CALL ADM_tarjeta_cliente('37','','" & Trim(Me.txtRucEmpresa.Text) & "','',''," & _
    "'','','','','','','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount > 0 Then
       Me.lblNotificacion.Height = 710
       Me.lblNotificacion.Caption = "CLIENTE YA REGISTRADO" + Chr(13) + "FECHA: " + Format(rstA("fecha_registro"), "dd-mm-YYYY") + Chr(13) + "VENDEDOR: " + get_persona(rstA("id_vendedor"))
       Me.lblNotificacion.Visible = True
       If Val(Me.frmtarjeta.Tag) < 1 Then
          Me.cmdRegistrar.Enabled = False
       Else
         Me.cmdRegistrar.Enabled = True
       End If
    Else
         Me.cmdRegistrar.Enabled = True
    End If
End If

End Sub
Private Sub txtRucEmpresa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
buscar_nuevamente:
    strCadena = "SELECT * FROM  persona WHERE  dni='" & Trim(Me.txtRucEmpresa.Text) & "' LIMIT 1 "
    Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            If get_dni_reniec_iii(Trim(Me.txtRucEmpresa.Text), KEY_DEPARTAMENTO, KEY_PROVINCIA, KEY_DISTRITO) = True Then
                GoTo buscar_nuevamente
            End If
        Else
            Me.txtRucEmpresa.Text = rst("dni")
            Me.TxtRazonSocial.Text = UCase(rst("nombre_completo"))
            Call Resalta(Me.txtDniApoderado)
        End If
End If
End Sub

Public Sub llenarAgenda(ByVal Grilla As MSHFlexGrid, ByVal in_dni As String)
On Error GoTo salir

Me.HfAgenda.Rows = 0

If KEY_CARGO = "00004" Then
    strCadena = "CALL ADM_tarjeta_cliente('36','','2','3','4'," & _
"'5','6','7','8','9','" & in_dni & "','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Else
    strCadena = "CALL ADM_tarjeta_cliente('28','','2','3','4'," & _
    "'5','6','7','8','9','" & in_dni & "','11','12','13','14','15','16','17','18','19','20','" & KEY_USUARIO & "','" & KEY_RUC & "')"
End If


Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
   
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
           Grilla.ColWidth(0) = 1100
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 2800
           
        cabecera = "DNI" & vbTab & "FECHA AGENDA" & vbTab & "PERSONAL"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 1 To rst.RecordCount
            
            Fila = rst("dni") & vbTab & Format(rst("fecha"), "dd-mm-YYYY") & vbTab & rst("nombre_completo")
            Grilla.AddItem Fila
            Grilla.RowHeight(i) = 320
            rst.MoveNext
    Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
  Call llenarAgenda_detalle(Me.HfAgenda, Me.HfDias.TextMatrix(Me.HfDias.Row, 1), Me.HfDias.TextMatrix(Me.HfDias.Row, 0))
  
    
   
    Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Public Sub llenarAgenda_detalle(ByVal Grilla As MSHFlexGrid, ByVal in_fecha As String, ByVal in_dni As String)
On Error GoTo salir

strCadena = "CALL ADM_tarjeta_cliente('29','1','2','3','4'," & _
"'5','6','7','8','9','10','11','12','13','14','15','16','" & Format(in_fecha, "YYYY-mm-dd") & "','" & Format(in_fecha, "YYYY-mm-dd") & "','19','20','" & in_dni & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
   
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 3000
           
        cabecera = "ID" & vbTab & "HORARIO" & vbTab & "DNI/RUC" & vbTab & "NOMBRE CLIENTE"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 1 To rst.RecordCount
            
            Fila = rst("id") & vbTab & rst("hora") & vbTab & rst("dni_cliente") & vbTab & rst("nombre_completo")
            Grilla.AddItem Fila
            Grilla.RowHeight(i) = 320
            rst.MoveNext
    Next i
    
    
   
    Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

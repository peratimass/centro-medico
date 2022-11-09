VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReporteRegistroVentas 
   BorderStyle     =   0  'None
   Caption         =   "Registro Ventas"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk_Manifiesto 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "MANIFES"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   12720
      TabIndex        =   99
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox chk_periodo 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "HASTA        :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5760
      TabIndex        =   94
      Top             =   600
      Width           =   1150
   End
   Begin VB.CheckBox chk_sucursal 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "SUCURSAL  :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5760
      TabIndex        =   93
      Top             =   960
      Width           =   1150
   End
   Begin VB.Frame frmEstadoCuenta 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   14520
      TabIndex        =   88
      Top             =   2040
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chk_todos 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TODOS."
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
         Left            =   120
         TabIndex        =   91
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CheckBox chk_credito 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "COMP. CREDITO"
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
         Left            =   120
         TabIndex        =   90
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chk_pendiente_pago 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "PENDIENTE PAGO."
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
         Left            =   120
         TabIndex        =   89
         Top             =   240
         Width           =   1935
      End
      Begin VitekeySoft.ChameleonBtn cmdEstadoCuenta 
         Height          =   825
         Left            =   2280
         TabIndex        =   92
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1455
         BTYPE           =   5
         TX              =   "GENERAR ESTADO DE CUENTA"
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRegVentas.frx":0000
         PICN            =   "FrmReporteRegVentas.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   180
         Left            =   2280
         TabIndex        =   98
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Image cmdcerrarEstado 
         Height          =   240
         Left            =   3780
         Picture         =   "FrmReporteRegVentas.frx":32F2
         Top             =   50
         Width           =   240
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00808080&
         Height          =   1575
         Left            =   0
         Top             =   0
         Width           =   4095
      End
   End
   Begin VB.Frame frmajusteCobrar 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "AJUSTE TIPO CAMBIO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   3255
      Left            =   12600
      TabIndex        =   73
      Top             =   4440
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtAjusteporCuenta 
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
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1200
         TabIndex        =   97
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chk_cta_contable 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "CTA CONT:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtCuentaPrincipal 
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
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1200
         TabIndex        =   87
         Top             =   1150
         Width           =   1815
      End
      Begin VB.TextBox txtCliente 
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
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1200
         TabIndex        =   84
         Top             =   1605
         Width           =   1815
      End
      Begin VB.TextBox txtCuentaGanancia 
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
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   4440
         TabIndex        =   83
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtCuentaPerdida 
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
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   4440
         TabIndex        =   82
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox txtTipocambio 
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
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1200
         TabIndex        =   74
         Top             =   720
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DtcPeriodo 
         Height          =   330
         Left            =   1200
         TabIndex        =   75
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   4194304
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
      Begin VitekeySoft.ChameleonBtn cmdProcesar 
         Height          =   435
         Left            =   1200
         TabIndex        =   76
         Top             =   2640
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   767
         BTYPE           =   5
         TX              =   "GENERAR AJUSTE"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRegVentas.frx":6196
         PICN            =   "FrmReporteRegVentas.frx":61B2
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
         Height          =   255
         Left            =   1200
         TabIndex        =   79
         Top             =   1950
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CUENTA REL:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CLIENTE :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   360
         TabIndex        =   85
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CTA GANACIA:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3390
         TabIndex        =   81
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CTA PERDIDA:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   3405
         TabIndex        =   80
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "PERIODO :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   405
         TabIndex        =   78
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "T.CAMBIO :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   345
         TabIndex        =   77
         Top             =   720
         Width           =   750
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         Height          =   3255
         Left            =   0
         Top             =   0
         Width           =   5895
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   5520
         Picture         =   "FrmReporteRegVentas.frx":8797
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.CheckBox chk_serie 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "SERIE:"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   280
      Left            =   9960
      TabIndex        =   71
      Top             =   900
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame frmblanquear 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   12840
      TabIndex        =   29
      Top             =   3840
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtMonto_pago 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4320
         TabIndex        =   43
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1200
         TabIndex        =   41
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox TxtNumeroAsociada 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3120
         TabIndex        =   37
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox TxtSerieAsociada 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2520
         TabIndex        =   36
         Top             =   3600
         Width           =   495
      End
      Begin VitekeySoft.ChameleonBtn cmdBlanquear 
         Height          =   525
         Left            =   360
         TabIndex        =   31
         Top             =   840
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   926
         BTYPE           =   3
         TX              =   "REESTABLECER"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRegVentas.frx":B63B
         PICN            =   "FrmReporteRegVentas.frx":B657
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn CMDASOCIAR 
         Height          =   645
         Left            =   360
         TabIndex        =   34
         Top             =   4485
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1138
         BTYPE           =   3
         TX              =   "VINCULAR"
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
         BCOL            =   16576
         BCOLO           =   16576
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRegVentas.frx":DC3C
         PICN            =   "FrmReporteRegVentas.frx":DC58
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcComprobanteAsociado 
         Height          =   330
         Left            =   360
         TabIndex        =   35
         Top             =   3600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DtcCobrador"
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
      Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
         Height          =   2085
         Left            =   5280
         TabIndex        =   39
         Top             =   1440
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   3678
         BTYPE           =   3
         TX              =   "DELL"
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
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRegVentas.frx":1023D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgVinculados 
         Height          =   2055
         Left            =   360
         TabIndex        =   40
         Top             =   1440
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   3625
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BUSCAR:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   330
         TabIndex        =   42
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label lblid_asociado 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3240
         TabIndex        =   38
         Top             =   4080
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   4680
         Picture         =   "FrmReporteRegVentas.frx":10259
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lblid_venta 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   375
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Label lblDocumento 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "BLANQUEAR DOCUMENTO:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   390
         TabIndex        =   30
         Top             =   360
         Width           =   3675
      End
   End
   Begin VB.TextBox txtTipo_cambio 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   1440
      TabIndex        =   61
      Top             =   675
      Width           =   1335
   End
   Begin VB.CheckBox chk_fechas 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "FECHAS      :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5760
      TabIndex        =   60
      Top             =   240
      Width           =   1150
   End
   Begin VB.Frame frm_dudosa 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7215
      Left            =   4440
      TabIndex        =   50
      Top             =   1320
      Visible         =   0   'False
      Width           =   14175
      Begin VB.TextBox txtDudosaRuc 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6030
         TabIndex        =   67
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtDudosaRazonSocial 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3930
         TabIndex        =   65
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtBuscarDudosa 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1320
         TabIndex        =   64
         Top             =   840
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDudosa 
         Height          =   5775
         Left            =   120
         TabIndex        =   51
         Top             =   1320
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   10186
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
      Begin VitekeySoft.ChameleonBtn cmdlanzardudosa 
         Height          =   645
         Left            =   12360
         TabIndex        =   52
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1138
         BTYPE           =   5
         TX              =   "LANZAR DUDOSA"
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
         BCOL            =   16576
         BCOLO           =   16576
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRegVentas.frx":130FD
         PICN            =   "FrmReporteRegVentas.frx":13119
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn mdRevertirDudosa 
         Height          =   645
         Left            =   12360
         TabIndex        =   53
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1138
         BTYPE           =   5
         TX              =   "ANULAR DUDOSA"
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
         BCOL            =   16576
         BCOLO           =   16576
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRegVentas.frx":156FE
         PICN            =   "FrmReporteRegVentas.frx":1571A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpInicio_dudosa 
         Height          =   315
         Left            =   1320
         TabIndex        =   54
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   137035777
         CurrentDate     =   41251
      End
      Begin MSComCtl2.DTPicker DtpFin_dudosa 
         Height          =   315
         Left            =   2760
         TabIndex        =   55
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   137035777
         CurrentDate     =   41251
      End
      Begin VitekeySoft.ChameleonBtn cmdConsultar 
         Height          =   375
         Left            =   4200
         TabIndex        =   57
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "CONSULTAR"
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
         BCOL            =   16576
         BCOLO           =   16576
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRegVentas.frx":17CFF
         PICN            =   "FrmReporteRegVentas.frx":17D1B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpFechaContable 
         Height          =   315
         Left            =   10440
         TabIndex        =   58
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   137035777
         CurrentDate     =   41251
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI/RUC:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5355
         TabIndex        =   68
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RAZON SOCIAL:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2820
         TabIndex        =   66
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMPROBANTE :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   30
         TabIndex        =   63
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA CONTABLE:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   9165
         TabIndex        =   59
         Top             =   285
         Width           =   1245
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHAS :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   600
         TabIndex        =   56
         Top             =   285
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   13080
         Picture         =   "FrmReporteRegVentas.frx":1A300
         Top             =   120
         Width           =   240
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         Height          =   5055
         Left            =   12300
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame frmIntereses 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   4920
      TabIndex        =   44
      Top             =   2640
      Visible         =   0   'False
      Width           =   13575
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfIntereses 
         Height          =   4095
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   7223
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
      Begin VitekeySoft.ChameleonBtn cmdGenerarInteres 
         Height          =   645
         Left            =   12000
         TabIndex        =   47
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1138
         BTYPE           =   5
         TX              =   "GENERAR INTERES"
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
         BCOL            =   16576
         BCOLO           =   16576
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmReporteRegVentas.frx":1D1A4
         PICN            =   "FrmReporteRegVentas.frx":1D1C0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         Height          =   4095
         Left            =   11940
         Top             =   480
         Width           =   1450
      End
      Begin VB.Image cmdcerrarFrame 
         Height          =   240
         Left            =   13200
         Picture         =   "FrmReporteRegVentas.frx":1F7A5
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INTERESES DIFERIDOS."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   135
         TabIndex        =   45
         Top             =   120
         Width           =   1785
      End
   End
   Begin VB.Frame frmcobrador 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   11040
      TabIndex        =   25
      Top             =   6120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdQuitarCobro 
         Caption         =   "QUITAR COBRO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   28
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton cmdcobradorAsignar 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3480
         TabIndex        =   27
         Top             =   360
         Width           =   495
      End
      Begin MSDataListLib.DataCombo DtcCobrador_update 
         Height          =   330
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DtcCobrador"
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
      Begin VB.Image cmdclose 
         Height          =   240
         Left            =   3840
         Picture         =   "FrmReporteRegVentas.frx":22649
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Timer timier_despacho 
      Enabled         =   0   'False
      Left            =   240
      Top             =   2280
   End
   Begin VB.TextBox txtserial 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   13800
      TabIndex        =   19
      Top             =   165
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkserial 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "CHASIS"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   12720
      TabIndex        =   18
      Top             =   200
      Width           =   975
   End
   Begin VitekeySoft.ChameleonBtn cmdamortizar 
      Height          =   705
      Left            =   18720
      TabIndex        =   14
      Top             =   1335
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1244
      BTYPE           =   5
      TX              =   "AMORTIZAR"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":254ED
      PICN            =   "FrmReporteRegVentas.frx":25509
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkTipoComprobante 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "TIPO COMPROBANTE"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   250
      Left            =   9960
      TabIndex        =   12
      Top             =   200
      Width           =   2535
   End
   Begin MSDataListLib.DataCombo dtcalmacen 
      Height          =   315
      Left            =   6960
      TabIndex        =   11
      Top             =   900
      Width           =   2565
      _ExtentX        =   4524
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
   Begin VitekeySoft.ChameleonBtn cmdBuscar 
      Height          =   315
      Left            =   15720
      TabIndex        =   9
      Top             =   120
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "BUSCAR                     "
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
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":27DF3
      PICN            =   "FrmReporteRegVentas.frx":27E0F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtrazonsocial 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4065
      TabIndex        =   7
      Top             =   200
      Width           =   1575
   End
   Begin VB.TextBox txtruc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4065
      TabIndex        =   2
      Top             =   675
      Width           =   1575
   End
   Begin VB.TextBox TxtNumero 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   200
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DtpDesde 
      Height          =   315
      Left            =   6960
      TabIndex        =   0
      Top             =   200
      Width           =   1260
      _ExtentX        =   2223
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
      Format          =   137035777
      CurrentDate     =   41251
   End
   Begin MSComCtl2.DTPicker DtpHasta 
      Height          =   315
      Left            =   8280
      TabIndex        =   6
      Top             =   200
      Width           =   1260
      _ExtentX        =   2223
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
      Format          =   137035777
      CurrentDate     =   41251
   End
   Begin VitekeySoft.ChameleonBtn cmdcuentasporcobrar 
      Height          =   315
      Left            =   15720
      TabIndex        =   10
      Top             =   480
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   " CUENTAS POR COBRAR"
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
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":283A9
      PICN            =   "FrmReporteRegVentas.frx":283C5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo dtcComprobante 
      Height          =   315
      Left            =   9960
      TabIndex        =   13
      Top             =   525
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   7815
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   18495
      _ExtentX        =   32623
      _ExtentY        =   13785
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
   Begin VitekeySoft.ChameleonBtn cmdhistorial 
      Height          =   795
      Left            =   18720
      TabIndex        =   15
      Top             =   6990
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1402
      BTYPE           =   5
      TX              =   "HISTORIAL"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":2895F
      PICN            =   "FrmReporteRegVentas.frx":2897B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdcerrarpantalla 
      Height          =   555
      Left            =   18720
      TabIndex        =   16
      Top             =   8565
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   ""
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":2BF84
      PICN            =   "FrmReporteRegVentas.frx":2BFA0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdreporte 
      Height          =   705
      Left            =   18720
      TabIndex        =   17
      Top             =   2115
      Width           =   650
      _ExtentX        =   1138
      _ExtentY        =   1244
      BTYPE           =   5
      TX              =   "E.CUENTA"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":2C390
      PICN            =   "FrmReporteRegVentas.frx":2C3AC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdproduccion 
      Height          =   780
      Left            =   18720
      TabIndex        =   20
      Top             =   6180
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1376
      BTYPE           =   5
      TX              =   "AJUSTE TC"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":2ECA4
      PICN            =   "FrmReporteRegVentas.frx":2ECC0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdReportedetallado 
      Height          =   705
      Left            =   18720
      TabIndex        =   21
      Top             =   5355
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1244
      BTYPE           =   5
      TX              =   "REPORT-DET"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":2F136
      PICN            =   "FrmReporteRegVentas.frx":2F152
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCredito 
      Height          =   315
      Left            =   15720
      TabIndex        =   22
      Top             =   840
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "VENTAS AL CREDITO     "
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
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":31A4A
      PICN            =   "FrmReporteRegVentas.frx":31A66
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdGeneral 
      Height          =   705
      Left            =   18720
      TabIndex        =   23
      Top             =   3735
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1244
      BTYPE           =   5
      TX              =   "CTAS COBRAR"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":32000
      PICN            =   "FrmReporteRegVentas.frx":3201C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCobrador 
      Height          =   765
      Left            =   18720
      TabIndex        =   24
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1349
      BTYPE           =   5
      TX              =   "COBRADOR"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":34914
      PICN            =   "FrmReporteRegVentas.frx":34930
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdseleccionar 
      Height          =   315
      Left            =   14640
      TabIndex        =   33
      Top             =   960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   ""
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
      BCOL            =   8454016
      BCOLO           =   8454016
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":34C4A
      PICN            =   "FrmReporteRegVentas.frx":34C66
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdIntereses 
      Height          =   465
      Left            =   17880
      TabIndex        =   48
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   820
      BTYPE           =   5
      TX              =   "INTERESES VENCIDOS"
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
      BCOL            =   33023
      BCOLO           =   33023
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":34F80
      PICN            =   "FrmReporteRegVentas.frx":34F9C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmddudosa 
      Height          =   495
      Left            =   17880
      TabIndex        =   49
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "COBRANZA DUDOSA"
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
      BCOL            =   33023
      BCOLO           =   33023
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":38CDD
      PICN            =   "FrmReporteRegVentas.frx":38CF9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdResumen 
      Height          =   705
      Left            =   18720
      TabIndex        =   69
      Top             =   2925
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1244
      BTYPE           =   5
      TX              =   "CTA RESUMEN"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":395D3
      PICN            =   "FrmReporteRegVentas.frx":395EF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcSerie 
      Height          =   315
      Left            =   11040
      TabIndex        =   70
      Top             =   900
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VitekeySoft.ChameleonBtn cmdEstadoCuentadetallado 
      Height          =   705
      Left            =   19400
      TabIndex        =   72
      Top             =   2115
      Width           =   650
      _ExtentX        =   1138
      _ExtentY        =   1244
      BTYPE           =   5
      TX              =   "DETALLE"
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
      BCOL            =   33023
      BCOLO           =   33023
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmReporteRegVentas.frx":3C8C5
      PICN            =   "FrmReporteRegVentas.frx":3C8E1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcPeriodoBusqueda 
      Height          =   315
      Left            =   6960
      TabIndex        =   95
      Top             =   550
      Width           =   2565
      _ExtentX        =   4524
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
   Begin MSDataListLib.DataCombo DtcManifiesto 
      Height          =   315
      Left            =   13800
      TabIndex        =   100
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO CAMBIO :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   62
      Top             =   720
      Width           =   930
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAZON SOCIAL :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2895
      TabIndex        =   8
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI/ RUC :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3255
      TabIndex        =   5
      Top             =   720
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N° DOCUMENTO :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   240
      Width           =   1065
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   120
      Top             =   60
      Width           =   19935
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmReporteRegistroVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Dim strCadenaII As String

Public Sub pendientes_despacho()
On Error GoTo salir
strCadena = "SELECT funct_pendiente_despacho('" & KEY_ALM & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & KEY_RUC & "')"
Call ConfiguraRstI(strCadena)
If rstI(0) > 0 Then
    PlaySound App.Path & "\sonidos\dingding.wav"
    PlaySound App.Path & "\sonidos\dingding.wav"
    strCadena = "SELECT id_venta FROM movimiento_venta WHERE fecha_emision ='" & KEY_FECHA & "' and pendiente='si' and  id_doc IN('0001','0003') and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 10"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       Me.timier_despacho.Enabled = False
       rstT.MoveFirst
       For i = 0 To rstT.RecordCount - 1
           Call impresion_despacho(rstT("id_venta"))
           strCadena = "p_update_despacho('" & rstT("id_venta") & "')"
           CnBd.Execute (strCadena)
           rstT.MoveNext
       Next i
       Me.timier_despacho.Enabled = True
    End If
   Exit Sub
End If
salir:
Me.timier_despacho.Enabled = True
End Sub

Private Sub ChameleonBtn1_Click()
If Me.hgVinculados.Rows > 0 Then
    strCadena = "DELETE FROM comprobante_asociado WHERE id_detalle='" & Val(Me.hgVinculados.TextMatrix(Me.hgVinculados.Row, 0)) & "'"
    CnBd.Execute (strCadena)
    Call llenar_vinculados(Me.hgVinculados)
End If
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chk_credito_Click()
If Me.chk_credito.Value = 1 Then
    Me.chk_todos.Value = 0
   Me.chk_pendiente_pago.Value = 0
End If

End Sub

Private Sub chk_cta_contable_Click()
If Me.chk_cta_contable.Value = 1 Then
    Me.txtAjusteporCuenta.Visible = True
Else
    Me.txtAjusteporCuenta.Visible = False
End If
End Sub

Private Sub chk_manifiesto_Click()
If Me.chk_Manifiesto.Value = 1 Then
    strCadena = "SELECT id_manifiesto as Codigo,manifiesto as Descripcion FROM view_manifiesto_numero WHERE ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcManifiesto)
    Me.DtcManifiesto.Visible = True
Else
    Me.DtcManifiesto.Visible = False
End If
End Sub

Private Sub chk_pendiente_pago_Click()

If Me.chk_pendiente_pago.Value = 1 Then
   Me.chk_credito.Value = 0
   Me.chk_todos.Value = 0
End If




End Sub

Private Sub chk_serie_Click()
If Me.chk_serie.Value = 1 Then
   
   strCadena = "SELECT serie as Codigo,serie as Descripcion FROM almacen_comprobante WHERE id_doc='" & Me.DtcComprobante.BoundText & "' and ruc='" & KEY_RUC & "' ORDER BY serie ASC"
   Call ConfiguraRst(strCadena)
   Call LlenaDataCombo(Me.DtcSerie)
   Me.DtcSerie.Visible = True
Else
  Me.DtcSerie.Visible = False
End If
End Sub

Private Sub chk_todos_Click()
If Me.chk_todos.Value = 1 Then
   Me.chk_credito.Value = 0
   Me.chk_pendiente_pago.Value = 0
End If

End Sub

Private Sub chkserial_Click()
If Me.chkserial.Value = 1 Then
   Me.txtserial.Visible = True
   Call Resalta(Me.txtserial)
Else
   Me.txtserial.Visible = False
End If
End Sub

Private Sub chkTipoComprobante_Click()
If Me.chkTipoComprobante.Value = 1 Then
    Me.DtcComprobante.Visible = True
    Me.chk_serie.Visible = True
Else
    Me.DtcComprobante.Visible = False
    Me.chk_serie.Visible = False
End If
End Sub

Private Sub actualizar_recibo(ByVal in_recibo As String)

strCadena = "DELETE FROM mis_cuentas_det_detalle WHERE id_movimiento='" & Val(in_recibo) & "'"
CnBd.Execute (strCadena)


strCadena = "SELECT * FROM mis_cuentas_det_detalle WHERE id_detalle='" & Val(in_recibo) & "' "
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   rstL.MoveFirst
   For i = 0 To rstL.RecordCount - 1
        strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & rstL("id_movimiento") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
           strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & rstL("id_movimiento") & "','" & Val(in_recibo) & "','" & rstL("monto_pagado") & "','" & rstL("monto_pagado") & "','" & rstZ("id_moneda") & "','" & rstZ("id_moneda") & "','" & rstZ("tc") & "')"
           CnBd.Execute (strCadena)
        End If
        rstL.MoveNext
   Next i
   MsgBox "Blanqueamiento Exitoso", vbInformation
End If

End Sub

Private Sub cmdamortizar_Click()
Procedencia = nuevo
frmVentasPagos.Show
End Sub
Private Sub blanquear_venta(ByVal in_venta As String)
strCadena = "select * from movimiento_venta where id_venta='" & Val(in_venta) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    
    If rst("id_doc") = "0054" Then  ' ACTUALIZAR RECIBO
    
       Call actualizar_recibo(rst("id_venta"))
       MsgBox "Blanqueamiento Correcto", vbInformation
       Exit Sub
    End If
    
    
    
    
    
    
    
    strCadena = "SELECT * FROM mis_cuentas_det_detalle WHERE id_movimiento='" & Val(in_venta) & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
            strCadena = "DELETE FROM mis_cuentas_det WHERE id='" & Val(rstT("id_cuenta_det")) & "'"
            CnBd.Execute (strCadena)
            rstT.MoveNext
        Next i
        strCadena = "DELETE FROM  mis_cuentas_det_detalle WHERE id_movimiento='" & Val(in_venta) & "'"
        CnBd.Execute (strCadena)
    End If
    
siguiente:
    On Error GoTo siguiente
    strCadena = "call CON_EliminarVenta('" & Val(in_venta) & "','" & KEY_USUARIO & "') "
    CnBd.Execute (strCadena)
     
     
    strCadena = "call P_insert_venta_agenda_test('" & Val(in_venta) & "')"
    CnBd.Execute (strCadena)
    
    
    If rst("id_forma_pago") = "01" Then
        strCadena = "SELECT * FROM movimiento_venta_monto WHERE id_venta='" & Val(in_venta) & "'"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            rstK.MoveFirst
            For i = 0 To rstK.RecordCount - 1
                 in_flujo = "1CIX000000000078"
                 in_glosa = "COBRO:" & rst("documento")
                 Call procesar_transaccion_venta(rstK("id_forma_pago"), KEY_ALM, get_cuenta_pago(rstK("id_forma_pago")), rst("fecha_emision"), "00001", rst("id_cliente"), rst("ncliente"), in_glosa, rstK("monto_caja"), "0", rst("id_venta"), "0", rst("documento"), rst("tc"), rstK("id_tarjeta_operacion"), "1CIX000000000174", in_flujo, KEY_USUARIO, rst("id_doc"), KEY_RUC)
                 Call put_realizar_pago(Val(in_venta), Val(in_venta), rstK("monto_caja"), rst("id_doc"), rst("tc"), Val(in_mis_cuentas_det), "01")
                 rstK.MoveNext
            Next i
        End If
        
       
       
       
       
       
       
       
       If rst("id_doc") = "0007" Then  ' ACTUALIZAR nota
               Call actualizar_nota(rst("id_venta"), rst("id_comprobante"))
               MsgBox "Blanqueamiento Exitoso", vbInformation
       Exit Sub
    End If
Else
    If rst("id_doc") = "0007" Then  ' ACTUALIZAR nota
               Call actualizar_nota(rst("id_venta"), rst("id_comprobante"))
               MsgBox "Blanqueamiento Exitoso", vbInformation
       Exit Sub
    End If
    End If
    
    
    
    MsgBox "Blanqueamiento Exitoso", vbInformation
End If
End Sub


Private Sub blanquear_venta_solo_notas(ByVal in_venta As String)
strCadena = "select * from movimiento_venta where id_venta='" & Val(in_venta) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    
    strCadena = "SELECT * FROM mis_cuentas_det_detalle WHERE id_movimiento='" & Val(in_venta) & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        rstT.MoveFirst
        
        For i = 0 To rstT.RecordCount - 1
            strCadena = "DELETE FROM mis_cuentas_det WHERE id='" & Val(rstT("id_cuenta_det")) & "'"
            CnBd.Execute (strCadena)
            rstT.MoveNext
        Next i
        strCadena = "DELETE FROM  mis_cuentas_det_detalle WHERE id_movimiento='" & Val(in_venta) & "'"
        CnBd.Execute (strCadena)
    End If
   
    strCadena = "call CON_EliminarVenta('" & Val(in_venta) & "','" & KEY_USUARIO & "') "
    CnBd.Execute (strCadena)
     
    strCadena = "call P_insert_venta_agenda_test('" & Val(in_venta) & "')"
    CnBd.Execute (strCadena)
    
   
   
    Call actualizar_nota(rst("id_venta"), rst("id_comprobante"))
   
End If
    
    
    

End Sub

Private Function get_saldo_factura(ByVal in_venta As String) As Double

strCadena = "SELECT (total-function_pago_factura(id_venta,'" & KEY_FECHA & "',id_moneda,ruc)) as saldo FROM movimiento_venta WHERE id_venta='" & Val(in_venta) & "' and  ruc='" & KEY_RUC & "' "
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
    get_saldo_factura = rstAux("saldo")
Else
    get_saldo_factura = 0
End If

End Function

Private Function get_saldo_nota(ByVal in_nota As String) As Double
strCadena = "SELECT total-function_pago_factura(id_venta,'" & KEY_FECHA & "',id_moneda,ruc) as pago FROM view_listado_comprobante_vargas WHERE id_venta='" & Val(in_nota) & "' and  ruc='" & KEY_RUC & "' "
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    get_monto_nota = rstL("pago")
Else
    get_monto_nota = 0
End If

End Function


Private Sub actualizar_nota(ByVal in_nota As String, ByVal in_venta As String)
Dim in_monto_afecta As Double
Dim in_monto_nota As Double

If in_venta > 0 Then
strCadena = "DELETE FROM mis_cuentas_det_detalle WHERE id_movimiento='" & Val(in_venta) & "'"
Call CnBd.Execute(strCadena)

strCadena = "DELETE FROM mis_cuentas_det_detalle WHERE id_movimiento='" & Val(in_nota) & "'"
Call CnBd.Execute(strCadena)

strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & Val(in_nota) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
           in_monto_afecta = get_saldo_factura(in_venta)
           in_monto_nota = get_saldo_nota(in_nota)
           
           If in_monto_nota >= in_monto_afecta Then
              in_monto_efectivo = in_monto_afecta
           Else
            in_monto_efectivo = in_monto_nota
           End If
           
           If in_monto_efectivo > 0 Then
                strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & in_nota & "','" & in_venta & "','" & in_monto_efectivo & "','" & in_monto_efectivo & "','" & rst("id_moneda") & "','" & rst("id_moneda") & "','" & rst("tc") & "')"
                CnBd.Execute (strCadena)
           
                strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & in_venta & "','" & in_nota & "','" & in_monto_efectivo & "','" & in_monto_efectivo & "','" & rst("id_moneda") & "','" & rst("id_moneda") & "','" & rst("tc") & "')"
                CnBd.Execute (strCadena)
            End If
   
End If
        
        
        
        End If



End Sub


Private Sub CMDASOCIAR_Click()
If Val(Me.lblid_venta.Caption) > 0 And Me.hgVinculados.Rows > 0 Then
    Call put_vincular_pagos(Val(Me.lblid_venta.Caption))
    MsgBox "Vinculacion Correcta", vbInformation, KEY_VENDEDOR
End If
End Sub

Private Sub cmdBlanquear_Click()
    
    
    Call blanquear_venta(Val(Me.lblid_venta.Caption))
    



End Sub

Private Sub cmdBuscar_Click()

Dim in_operacion As String
    
If Me.chkserial.Value = 1 Then
    strCadena = "SELECT * FROM view_venta_chasis WHERE ruc='" & KEY_RUC & "' and nro_chasis LIKE '%" & Trim(Me.txtserial.Text) & "%'"
    Call llenar_grid(Me.HfdPersona)
    Exit Sub
End If


in_operacion = "7"

If Me.chk_fechas.Value = 1 Then
    in_operacion = "7"
    
    If Trim(Me.txtRuc.Text) <> "" Or Trim(Me.txtrazonsocial.Text) <> "" Or Trim(Me.TxtNumero.Text) <> "" Then
        in_operacion = "8"
    End If
    
    If Trim(Me.txtrazonsocial.Text) <> "" Then
         in_operacion = "18"
    End If
    
    
    If Trim(Me.txtRuc.Text) <> "" Then
        in_operacion = "16"
    End If
    
    
    
    If Me.chkTipoComprobante.Value = 1 Then
       in_operacion = "9"
       If Me.chk_serie.Value = 1 Then
           in_operacion = "10"
       End If
    End If

Else
    If Me.chkTipoComprobante.Value = 1 Then
        in_operacion = "11"
        If Me.chk_serie.Value = 1 Then
           in_operacion = "12"
       End If
    End If
    
    If Trim(Me.txtRuc.Text) <> "" Then
    in_operacion = "17"
    End If
    
End If

strCadena = "CALL CON_CuentaCobrar_LST('" & in_operacion & "','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.txtrazonsocial.Text) & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Me.DtcComprobante.BoundText & "','" & Trim(Me.TxtNumero.Text) & "','" & Me.DtcSerie.BoundText & "','" & KEY_RUC & "')"
Call llenar_grid(Me.HfdPersona)



Me.cmdReporte.Enabled = True
Me.cmdReporteDetallado.Enabled = True
End Sub



Private Sub cmdCliente_Click()

End Sub

Private Sub cmdfechas_Click()
 
End Sub

Private Sub cmdcerrarEstado_Click()
Me.frmEstadoCuenta.Visible = False
End Sub

Private Sub cmdcerrarframe_Click()
Me.frmIntereses.Visible = False
End Sub

Private Sub cmdClose_Click()
Me.frmcobrador.Visible = False
End Sub

Private Sub cmdCobrador_Click()
If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
   Me.frmcobrador.Visible = True
End If
End Sub

Private Sub cmdcobradorAsignar_Click()

strCadena = "UPDATE movimiento_venta SET id_almacenero='" & Me.DtcCobrador_update.BoundText & "' WHERE id_venta='" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
Me.frmcobrador.Visible = False

End Sub



Private Sub cmdCobrar_Click()

End Sub

Private Sub cmdConsultar_Click()

Me.DtpFechaContable.Value = KEY_FECHA
strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,dias,ncliente,comprobante,fecha_cobranza_dudosa,moneda,(total-function_pago_factura(id_venta,'" & KEY_FECHA & "',id_moneda,ruc)) as saldo,dudosa_contable FROM view_dudosa WHERE " & _
" fecha_emision>='" & Format(Me.DtpInicio_dudosa.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_dudosa.Value, "YYYY-mm-dd") & "' and  dias>=365  and   ruc='" & KEY_RUC & "' "
Call llenar_dudosa(HfDudosa)
End Sub

Private Sub cmddudosa_Click()

Me.DtpFechaContable.Value = KEY_FECHA

strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,dias,ncliente,comprobante,fecha_cobranza_dudosa,moneda,(total-function_pago_factura(id_venta,'" & KEY_FECHA & "',id_moneda,ruc)) as saldo,dudosa_contable FROM view_dudosa WHERE dudosa_contable='no'  and   dias>=365 and ruc='" & KEY_RUC & "' LIMIT 20"
Call llenar_dudosa(HfDudosa)
Me.frm_dudosa.Visible = True


End Sub
Private Sub llenar_dudosa(ByVal Grilla As MSHFlexGrid)
Dim nsaldo As Double
On Error GoTo salir

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
   

   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 600
           Grilla.ColWidth(4) = 2200
           Grilla.ColWidth(5) = 3000
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 1500
           
           
        Next
         cabecera = "IDVENTA" & vbTab & "F.EMISION" & vbTab & "F.VENCIMIENTO" & vbTab & "DIAS" & vbTab & "COMPROBANTE" & vbTab & "DATOS CLIENTE" & vbTab & "SALDO" & vbTab & "FECHA DUDOSA"
         Grilla.AddItem cabecera
         For k = 1 To 7
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        nsaldo = 0
        For i = 0 To rst.RecordCount - 1
             in_saldo = rst("saldo")
             Fila = rst("id_venta") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & Format(rst("fecha_vencimiento"), "dd-mm-YYYY") & vbTab & rst("dias") & vbTab & rst("comprobante") & vbTab & rst("ncliente") & vbTab & Format(in_saldo, "#,##0.00") & vbTab & Format(rst("fecha_cobranza_dudosa"), "dd-mm-YYYY")
             Grilla.AddItem Fila
             nsaldo = nsaldo + in_saldo
             '&H000080FF&
             If rst("dudosa_contable") = "si" Then
                For k = 4 To 7
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H80FF&
                Next k
             End If
             
             rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL CANCELADO:" & vbTab & Format(nsaldo, "#,##0.00")
        Grilla.AddItem Fila
        For k = 1 To 6
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
            Next k
 ' Grilla.Row = 1
 ' Grilla.col = 0
 ' Grilla.ColSel = 1
 ' Grilla.RowSel = 1
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub cmdEstadoCuenta_Click()
Dim arr(0 To 2, 1 To 2) As String
Dim in_doc As String
Dim param As Variant

arr(0, 1) = "fecha_ini"
arr(1, 1) = "fecha_fin"
arr(2, 1) = "telefono"

arr(0, 2) = get_direccion(Trim(Me.txtRuc.Text))
arr(1, 2) = Format(Me.DtpHasta.Value, "dd-mm-YYYY")
arr(2, 2) = get_telefono(Trim(Me.txtRuc.Text))
param = arr()
 'TODOS LOS PAGOS
    strCadena = "DELETE FROM adm_estado_cuenta WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    
If Me.chk_pendiente_pago.Value = 1 Then
    strCadena = "call ADM_EstadoCuentav2('5','" & Trim(Me.txtRuc.Text) & "','" & Val(Me.txtTipo_cambio.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       Me.ProgressBar1.Min = 0
       Me.ProgressBar1.Max = rst.RecordCount
       
       For i = 1 To rst.RecordCount
           strCadena = "call put_estado_cuenta_saldo('" & rst("id_venta") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           rst.MoveNext
            Me.ProgressBar1.Value = i
           DoEvents
       Next i
       
    End If
    
    
    
    strCadena = "call ADM_EstadoCuentav2('6','" & Trim(Me.txtRuc.Text) & "','" & Val(Me.txtTipo_cambio.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptEstadodeCuenta", param, App.Path + "\Reportes\")
    Exit Sub
End If












If Me.chk_credito.Value = 1 Then
    strCadena = "call ADM_EstadoCuentav2('7','" & Trim(Me.txtRuc.Text) & "','" & Val(Me.txtTipo_cambio.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       Me.ProgressBar1.Min = 0
       Me.ProgressBar1.Max = rst.RecordCount
       
       For i = 1 To rst.RecordCount
           strCadena = "call put_estado_cuenta_saldo('" & rst("id_venta") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           rst.MoveNext
            Me.ProgressBar1.Value = i
           DoEvents
       Next i
       
    End If
    
    
    
    strCadena = "call ADM_EstadoCuentav2('8','" & Trim(Me.txtRuc.Text) & "','" & Val(Me.txtTipo_cambio.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptEstadodeCuenta", param, App.Path + "\Reportes\")
    Exit Sub
End If



If Me.chk_todos.Value = 1 Then
    strCadena = "call ADM_EstadoCuentav2('7','" & Trim(Me.txtRuc.Text) & "','" & Val(Me.txtTipo_cambio.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       Me.ProgressBar1.Min = 0
       Me.ProgressBar1.Max = rst.RecordCount
       
       For i = 1 To rst.RecordCount
           strCadena = "call put_estado_cuenta_saldo('" & rst("id_venta") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           rst.MoveNext
            Me.ProgressBar1.Value = i
           DoEvents
       Next i
       
    End If
    
    
    
    strCadena = "call ADM_EstadoCuentav2('8','" & Trim(Me.txtRuc.Text) & "','" & Val(Me.txtTipo_cambio.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptEstadodeCuenta", param, App.Path + "\Reportes\")
    Exit Sub
End If





End Sub

Private Sub cmdEstadoCuentadetallado_Click()
If Trim(Me.txtRuc.Text) = "" Then
   Me.txtRuc.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 5)
End If



Me.frmEstadoCuenta.Visible = True

End Sub

Private Sub cmdexportar_Click()

End Sub

Private Sub cmdProcesar_Click()

If MsgBox("Esta seguro de Realizar el ajuste [TC] para:" + Chr(13) + Chr(13) + "TIPO CAMBIO:" & str(Me.TxtTipoCambio.Text) + Chr(13) + "PERIODO AJUSTE:" + Me.DtcPeriodo.Text, vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then
If KEY_RUC = "20128836251" Then
    If Trim(Me.TxtCliente.Text) <> "" Then
       strCadena = "SELECT id_cliente,id_doc,serie,numero,comprobante FROM view_comprobante WHERE id_cliente='" & Trim(Me.TxtCliente.Text) & "' and id_moneda='00002' and  anulado='no' and  ruc='" & KEY_RUC & "' and id_forma_pago='02'"
    Else
       strCadena = "SELECT id_cliente,id_doc,serie,numero,comprobante FROM view_comprobante WHERE id_moneda='00002' and  anulado='no' and  ruc='" & KEY_RUC & "' and id_forma_pago='02' "
    End If
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Me.prg_avance.Min = 0
        Me.prg_avance.Max = rst.RecordCount - 1
        For i = 0 To rst.RecordCount - 1
            
            strCadena = "call CON_InsertaAsiento_AjusteTC_Persona_comprobante('" & KEY_RUC & "','" & Me.DtcPeriodo.BoundText & "','" & Trim(Me.txtCuentaPrincipal.Text) & "','" & Trim(rst("id_cliente")) & "','" & Val(Me.TxtTipoCambio.Text) & "','" & Val(rst("id_doc")) & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("comprobante") & "','" & Trim(Me.txtCuentaPerdida.Text) & "','" & Trim(Me.txtCuentaGanancia.Text) & "','" & KEY_USUARIO & "')"
            CnBd.Execute (strCadena)
            rst.MoveNext
            prg_avance.Value = i
            DoEvents
        Next i
   
    End If


Else
    
    
    
    If Me.chk_cta_contable.Value = 1 Then
        If Trim(Me.txtAjusteporCuenta.Text) <> "" And Trim(TxtCliente.Text) <> "" Then
            strCadena = "call CON_InsertaAsiento_AjusteTC_Empresa('" & KEY_RUC & "','" & Me.DtcPeriodo.BoundText & "','" & Trim(Me.txtAjusteporCuenta.Text) & "','" & Trim(TxtCliente.Text) & "','" & Val(Me.TxtTipoCambio.Text) & "','" & KEY_USUARIO & "')"
            ConfiguraRstK (strCadena)
        End If
    Else
    
    If Trim(Me.TxtCliente.Text) <> "" Then
        strCadena = "SELECT DISTINCT id_cliente FROM view_listado_comprobante_vitekey WHERE id_cliente='" & Trim(Me.TxtCliente.Text) & "' and  id_moneda='00002' and  anulado='no' and  ruc='" & KEY_RUC & "' and id_forma_pago='02' and (total-function_pago_factura(id_venta,'" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "',id_moneda,ruc))>0"
    Else
        strCadena = "SELECT DISTINCT id_cliente FROM view_listado_comprobante_vitekey WHERE id_moneda='00002' and  anulado='no' and  ruc='" & KEY_RUC & "' and id_forma_pago='02' and (total-function_pago_factura(id_venta,'" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "',id_moneda,ruc))>0"
    End If
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Me.prg_avance.Min = 0
        Me.prg_avance.Max = rst.RecordCount - 1
        For i = 0 To rst.RecordCount - 1
           strCadena = "call CON_InsertaAsiento_AjusteTC_Cliente('" & KEY_RUC & "','" & Me.DtcPeriodo.BoundText & "','1212','" & Trim(rst("id_cliente")) & "','" & Val(Me.TxtTipoCambio.Text) & "','" & KEY_USUARIO & "')"
           ConfiguraRstK (strCadena)
           rst.MoveNext
           prg_avance.Value = i
           DoEvents
        Next i
    End If
    End If
End If





MsgBox "AJUSTE DE TIPO CAMBIO REALIZADO", vbInformation, KEY_VENDEDOR
End If





End Sub

Private Sub cmdGenerarInteres_Click()

Dim in_venta As String
in_venta = Val(Me.HfIntereses.TextMatrix(Me.HfIntereses.Row, 0))

If Me.HfIntereses.Rows > 0 Then
    If Val(Me.HfIntereses.TextMatrix(Me.HfIntereses.Row, 0)) > 0 Then
        strCadena = "call CON_InsertaAsiento_InteresDevengado('" & in_venta & "')"
        CnBd.Execute (strCadena)
        strCadena = "UPDATE movimiento_venta SET interes_diferido='si',interes_revertido='0' WHERE id_venta='" & in_venta & "'"
        CnBd.Execute (strCadena)
        
    End If
    MsgBox "DEVENGADO Correctamente.....", vbInformation
    Call llenar_letras_vencidas(Me.HfIntereses)
End If
End Sub

Private Sub cmdIntereses_Click()
Call llenar_letras_vencidas(Me.HfIntereses)
Me.frmIntereses.Visible = True

End Sub

Private Sub cmdlanzardudosa_Click()

strCadena = "UPDATE movimiento_venta SET cobranza_dudosa='si',dudosa_contable='si',fecha_cobranza_dudosa='" & Format(Me.DtpFechaContable.Value, "YYYY-mm-dd") & "' WHERE id_venta='" & Val(Me.HfDudosa.TextMatrix(Me.HfDudosa.Row, 0)) & "'"
CnBd.Execute (strCadena)

strCadena = "call CON_InsertaAsiento_CobranzaDudosa('" & Val(Me.HfDudosa.TextMatrix(Me.HfDudosa.Row, 0)) & "')"
CnBd.Execute (strCadena)

MsgBox "Ingresado a Cobranza Dudosa", vbInformation



End Sub
Private Function get_id_doc(ByVal in_descripcion As String) As String
get_id_doc = ""
Select Case in_descripcion
    Case "FACTURA"
          get_id_doc = "0001"
    Case "BOLETA"
          get_id_doc = "0003"
    Case "RECIBO VENTA"
        get_id_doc = "0054"
    Case "RBO VENTA"
        get_id_doc = "0054"
    Case "N"
        
    
End Select


End Function
Private Sub cmdpagos_Click()
End Sub
Private Function get_id_venta(ByVal in_doc, ByVal in_serie As String, ByVal in_numero As String) As Double

strCadena = "SELECT * FROM movimiento_venta where id_doc='" & in_doc & "' and serie='" & in_serie & "' and numero='" & in_numero & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   get_id_venta = rstK("id_venta")
End If

End Function


Private Function get_id_compra(ByVal in_doc, ByVal in_serie As String, ByVal in_numero As String, ByVal in_proveedor As String) As Double

strCadena = "SELECT id_compra FROM movimiento_compra where id_proveedor='" & in_proveedor & "' and  id_doc='" & in_doc & "' and serie='" & in_serie & "' and numero='" & in_numero & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   get_id_compra = rstK("id_compra")
Else
    get_id_compra = 0
End If

End Function


Private Sub cmdQuitarCobro_Click()
strCadena = "UPDATE movimiento_venta SET id_almacenero='0' WHERE id_venta='" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
MsgBox "Se a quitado la Cobrazan de este comprobante", vbInformation, KEY_VENDEDOR
End Sub





Private Sub cmdCerrar_Click()

End Sub

Private Sub cmdCerrarpantalla_Click()
Unload Me
End Sub

Private Sub cmdCredito_Click()
    
    
    
    
    
    
    
    strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,hora,numero,comprobante,id_cliente,ncliente,total,saldo,anulado,id_moneda,tc,id_alm,id_doc," & _
    " id_proyecto,vendedor as nombre_completo,descripcion,simbolo,id_forma_pago,referencia,function_pago_factura(id_venta,'" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "',id_moneda,ruc) as pago " & _
    " FROM view_listado_comprobante_vitekey WHERE ruc='" & KEY_RUC & "'  and id_forma_pago='02'and fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "'"
        Call llenar_grid(Me.HfdPersona)
        
    

End Sub

Private Sub cmdcuentasporcobrar_Click()
Dim in_operacion As String
    If Me.chk_fechas.Value = 0 Then
       
        in_operacion = "1"
       
       If Trim(Me.txtrazonsocial.Text) <> "" Then
          in_operacion = "6"
       End If
       If Trim(Me.txtRuc.Text) <> "" Then
          in_operacion = "5"
       End If
       
       If Me.chk_periodo.Value = 1 Then
        
       End If
    
    
    
    Else
        in_operacion = "2"
       
       If Trim(Me.txtrazonsocial.Text) <> "" Then
          in_operacion = "4"
       End If
       If Trim(Me.txtRuc.Text) <> "" Then
          in_operacion = "3"
       End If
    End If
       
    
    
    
    
    
    If Me.chk_serie.Value = 0 Then
        in_serie = ""
    Else
        in_serie = Me.DtcSerie.BoundText
    End If
       
    
    
    
    strCadena = "CALL CON_CuentaCobrar_LST('" & in_operacion & "','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.txtrazonsocial.Text) & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Me.DtcComprobante.BoundText & "','" & Trim(Me.TxtNumero.Text) & "','" & in_serie & "','" & KEY_RUC & "')"
    Call llenar_grid(Me.HfdPersona)
       






End Sub



Private Sub cmdGeneral_Click()
Dim cam3(0 To 2, 1 To 2)  As String


Dim param As Variant


                    cam3(0, 1) = "fecha_ini"
                    cam3(1, 1) = "fecha_fin"
                    cam3(2, 1) = "cambio"
                    
                   If Me.chk_fechas.Value = 1 Then
    
                    cam3(0, 2) = Format(Me.DtpDesde.Value, "dd-mm-YYYY")
                   Else
                    cam3(0, 2) = "INICIO"
                   End If
                    cam3(1, 2) = Format(Me.DtpHasta.Value, "dd-mm-YYYY")
                    cam3(2, 2) = KEY_VENDEDOR
                    param = cam3()
                  
                  

 
 'If MsgBox("Desea generar con fecha de Ultimo Pago", vbQuestion + vbYesNo, KEY_EMPRESA) = vbNo Then
 
 If Me.chk_fechas.Value = 1 Then
    strCadena = "SELECT fecha_emision,fecha_vencimiento,comprobante,id_cliente,ncliente,total,total-function_pago_factura(id_venta,'" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,id_moneda,tc " & _
    " FROM view_listado_comprobante_vitekey WHERE (total-function_pago_factura(id_venta,'" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "',id_moneda,ruc))<>0 and  " & _
    " anulado='no' and  ruc='" & KEY_RUC & "' and id_doc not in('0099','0097') and id_forma_pago='02' and fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "'"
 Else
    strCadena = "SELECT fecha_emision,fecha_vencimiento,comprobante,id_cliente,ncliente,total,total-function_pago_factura(id_venta,'" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,id_moneda,tc " & _
    " FROM view_listado_comprobante_vitekey WHERE (total-function_pago_factura(id_venta,'" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "',id_moneda,ruc)) <>0 and " & _
    " anulado='no' and  ruc='" & KEY_RUC & "' and id_doc not in('0099','0097') and id_forma_pago='02' and (total-function_pago_factura(id_venta,'" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "',id_moneda,ruc))<>0"
 End If

If Me.chk_fechas.Value = 1 Then
   in_operacion = "19"
Else
   in_operacion = "18"
End If



strCadena = "CALL CON_CuentaCobrar_LST2('" & in_operacion & "','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.txtrazonsocial.Text) & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Me.DtcComprobante.BoundText & "','" & Trim(Me.TxtNumero.Text) & "','" & in_serie & "','" & Val(Me.txtTipo_cambio.Text) & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_cliente_detalle_cobranza", param, App.Path + "\Reportes\")
 



End Sub
Private Function get_id_moneda(ByVal in_venta As String) As String
strCadena = "SELECT id_moneda FROM movimiento_venta WHERE id_venta='" & in_venta & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    get_id_moneda = rst("id_moneda")
End If
End Function
Private Sub cmdHistorial_Click()



Dim arr(0 To 2, 1 To 2) As String
Dim param As Variant

arr(0, 1) = "in_cliente"
arr(1, 1) = "id_moneda"
arr(0, 2) = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 6))
arr(1, 2) = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 7)

param = arr()
  
strCadena = "call ADM_historial_pago('1','" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "','" & KEY_RUC & "')"
  
 'strCadena = "SELECT fecha_emision,id_cliente,documento,fecha_origen,recibo,id_moneda,monto_pagado,tc,nombre_completo,forma_pago,id_tarjeta_operacion,total FROM view_historial_pago_v2 WHERE id_venta='" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
          Call ConfiguraRst(strCadena)
          Ans = ShowMultiReport(rst, "RptHistorial_venta", param, App.Path + "\Reportes\")
         
          
End Sub



Private Sub cmdproduccion_Click()
strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodo)
  

Me.frmajusteCobrar.Visible = True

Me.frmajusteCobrar.Visible = True


Exit Sub
If MsgBox("Esta seguro de Realizar el ajuste [TC] para:" + Chr(13) + Chr(13) + "TIPO CAMBIO:" & str(Me.txtTipo_cambio.Text) + Chr(13) + "PERIODO AJUSTE:" + get_periodo_descripcion(get_periodo_actual(Me.DtpDesde.Value)), vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then

strCadena = "SELECT DISTINCT id_cliente FROM view_listado_comprobante_vitekey WHERE id_moneda='00002' and  anulado='no' and  ruc='" & KEY_RUC & "' and id_forma_pago='02' and (total-function_pago_factura(id_venta,'" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "',id_moneda,ruc))>0"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "call CON_InsertaAsiento_AjusteTC_Empresa('" & KEY_RUC & "','" & get_periodo_actual(Me.DtpDesde.Value) & "','1212','" & Trim(rst("id_cliente")) & "','" & Val(Me.txtTipo_cambio.Text) & "','" & KEY_USUARIO & "')"
        CnBd.Execute (strCadena)
        rst.MoveNext
        DoEvents
   Next i
   
End If


MsgBox "AJUSTE DE TIPO CAMBIO REALIZADO", vbInformation, KEY_VENDEDOR
End If

'strCadena = "SELECT id_venta,documento,fecha_emision,id_cliente,ncliente,total,id_vendedor,nombre_completo,id_linea,descripcion,nro_chasis,serie " & _
"  FROM view_venta_produccion WHERE nro_chasis<>'-' and   id_alm =  '" & Me.DtcAlmacen.BoundText & "' and  fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "'"
'Call ConfiguraRst(strCadena)
'Ans = ShowMultiReport(rst, "rpt_ventas_produccion", param, App.Path + "\Reportes\")

End Sub



Private Sub cmdReporte_Click()
Dim arr(0 To 2, 1 To 2) As String
Dim param As Variant
Dim in_fecha As Date
Dim in_ruc As String



   Me.txtRuc.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 5)


If Len(Trim(Me.txtRuc.Text)) = 8 Then
   in_ruc = "10" & Trim(Me.txtRuc.Text)
   in_ruc = DigitoVerificadorRUC(Trim(in_ruc))

End If

arr(0, 1) = "in_ruc"
arr(0, 2) = in_ruc
param = arr()


If Me.chk_fechas.Value = 1 Then
   in_fecha = Format(Me.DtpHasta.Value, "YYYY-mm-dd")
   in_parametro = " and fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "'"
Else
   in_fecha = KEY_FECHA
   in_parametro = ""
End If


If Trim(Me.txtRuc.Text) <> "" Then
    strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,dias,hora,numero,comprobante,id_cliente,ncliente,direccion,celular,total, if (id_doc='0007',(total-function_pago_factura(id_venta,'" & Format(in_fecha, "YYYY-mm-dd") & "',id_moneda,ruc))*-1,(total-function_pago_factura(id_venta,'" & Format(in_fecha, "YYYY-mm-dd") & "',id_moneda,ruc)))   as saldo  ,anulado,id_moneda,ruc,'" & Val(Me.txtTipo_cambio.Text) & "',id_alm,id_doc,referencia,descripcion,simbolo FROM view_listado_comprobante_vitekey WHERE (total-function_pago_factura(id_venta,'" & Format(in_fecha, "YYYY-mm-dd") & "',id_moneda,ruc))<>0 and id_forma_pago='02' and  anulado='no'  and id_cliente LIKE '%" & Trim(Me.txtRuc.Text) & "%' AND ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,dias,hora,numero,comprobante,id_cliente,ncliente,direccion,celular,total, if (id_doc='0007',(total-function_pago_factura(id_venta,'" & Format(in_fecha, "YYYY-mm-dd") & "',id_moneda,ruc))*-1,(total-function_pago_factura(id_venta,'" & Format(in_fecha, "YYYY-mm-dd") & "',id_moneda,ruc)))   as saldo  ,anulado,id_moneda,ruc,tc,id_alm,id_doc,referencia,descripcion,simbolo FROM view_listado_comprobante_vitekey WHERE (total-function_pago_factura(id_venta,'" & Format(in_fecha, "YYYY-mm-dd") & "',id_moneda,ruc))<>0 and id_forma_pago='02' and  anulado='no'  and ncliente LIKE '%" & Trim(Me.txtrazonsocial.Text) & "%' AND ruc='" & KEY_RUC & "'"
End If

Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_cliente_detalle_cobranza_letras", param, App.Path + "\Reportes\")



End Sub
Private Sub cmdReporteDetallado_Click()

Dim arr(0 To 2, 1 To 2)  As String
Dim param As Variant
arr(0, 1) = "emision"
arr(1, 1) = "vencimiento"
arr(2, 1) = "almacen"

arr(0, 2) = Format(Me.DtpDesde.Value, "dd/mm/YYYY")
arr(1, 2) = Format(Me.DtpHasta.Value, "dd/mm/YYYY")
arr(2, 2) = Trim(Me.DtcAlmacen.Text)
param = arr()



If Me.chkTipoComprobante.Value = 1 Then
    
    If Me.chk_serie.Value = 1 Then
        If Me.chk_fechas.Value = 1 Then
        '-
            strCadena = "SELECT id_venta,fecha_emision,documento,id_cliente,ncliente,nombre_completo,total,saldo,anulado,id_producto,detalle,cantidad,precio,ttotal,id_alm,ruc " & _
            " FROM view_listado_comprobante_iii WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and   id_doc='" & Me.DtcComprobante.BoundText & "' and serie='" & Me.DtcSerie.BoundText & "' and ruc='" & KEY_RUC & "'"
            
            strCadena = "SELECT  `id_producto`,LEFT(nombre_prod, 70),linea,`cantidad` " & _
            " FROM view_venta_detalle_vitekey WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and   id_doc='" & Me.DtcComprobante.BoundText & "' and serie='" & Me.DtcSerie.BoundText & "' and ruc='" & KEY_RUC & "'"
            
            
        Else
            strCadena = "SELECT id_venta,fecha_emision,documento,id_cliente,ncliente,nombre_completo,total,saldo,anulado,id_producto,detalle,cantidad,precio,ttotal,id_alm,ruc " & _
            " FROM view_venta_detalle_ii WHERE id_doc='" & Me.DtcComprobante.BoundText & "' and serie='" & Me.DtcSerie.BoundText & "' and ruc='" & KEY_RUC & "'"
            
            strSQL = "SELECT `id_producto`,`nombre_prod`,linea,`cantidad` " & _
            " FROM view_venta_detalle_vitekey WHERE   id_doc='" & Me.DtcComprobante.BoundText & "' and serie='" & Me.DtcSerie.BoundText & "' and ruc='" & KEY_RUC & "'"
            
        End If
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "rpt_ventas_consolidado", param, App.Path + "\Reportes\")
        Exit Sub
    Else
        If Me.chk_fechas.Value = 1 Then
            strCadena = "SELECT id_venta,fecha_emision,documento,id_cliente,ncliente,nombre_completo,total,saldo,anulado,id_producto,detalle,cantidad,precio,ttotal,id_alm,ruc " & _
            " FROM view_listado_comprobante_iii WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and   id_doc='" & Me.DtcComprobante.BoundText & "'  and ruc='" & KEY_RUC & "'"
        Else
            strCadena = "SELECT id_venta,fecha_emision,documento,id_cliente,ncliente,nombre_completo,total,saldo,anulado,id_producto,detalle,cantidad,precio,ttotal,id_alm,ruc " & _
            " FROM view_listado_comprobante_iii WHERE id_cliente LIKE '%" & Trim(Me.txtRuc.Text) & "%' and  id_doc='" & Me.DtcComprobante.BoundText & "'  and ruc='" & KEY_RUC & "'"
        End If
    End If
Else
    
    
    
        If Me.chk_fechas.Value = 1 Then
            If Trim(Me.txtRuc.Text) <> "" Then
                strCadena = "SELECT id_venta,fecha_emision,documento,id_cliente,ncliente,nombre_completo,total,saldo,anulado,id_producto,detalle,cantidad,precio,ttotal,id_alm,ruc " & _
            " FROM view_listado_comprobante_iii WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and   id_cliente  LIKE '%" & Trim(Me.txtRuc.Text) & "%'  and ruc='" & KEY_RUC & "'"
            Else
                strCadena = "SELECT id_venta,fecha_emision,documento,id_cliente,ncliente,nombre_completo,total,saldo,anulado,id_producto,detalle,cantidad,precio,ttotal,id_alm,ruc " & _
                " FROM view_listado_comprobante_iii WHERE fecha_emision>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and   nombre_completo  LIKE '%" & Trim(Me.txtrazonsocial.Text) & "%'  and ruc='" & KEY_RUC & "'"
            End If
            
        Else
            
            If Trim(Me.txtRuc.Text) <> "" Then
                strCadena = "SELECT id_venta,fecha_emision,documento,id_cliente,ncliente,nombre_completo,total,saldo,anulado,id_producto,detalle,cantidad,precio,ttotal,id_alm,ruc " & _
            " FROM view_listado_comprobante_iii WHERE id_cliente  LIKE '%" & Trim(Me.txtRuc.Text) & "%'  and ruc='" & KEY_RUC & "'"
            Else
                strCadena = "SELECT id_venta,fecha_emision,documento,id_cliente,ncliente,nombre_completo,total,saldo,anulado,id_producto,detalle,cantidad,precio,ttotal,id_alm,ruc " & _
            " FROM view_listado_comprobante_iii WHERE nombre_completo  LIKE '%" & Trim(Me.txtrazonsocial.Text) & "%'  and ruc='" & KEY_RUC & "'"
            End If
            
        End If
    End If
    
    
    
    


Call ConfiguraRst(strCadena)


Ans = ShowMultiReport(rst, "rpt_cliente_detalle_ii", "", App.Path + "\Reportes\")

End Sub

Private Sub cmdResumen_Click()
Dim arr(0 To 1, 1 To 2)  As String
Dim param As Variant
arr(0, 1) = "emision"
arr(1, 1) = "vencimiento"


arr(0, 2) = Format(Me.DtpDesde.Value, "dd/mm/YYYY")
arr(1, 2) = Format(Me.DtpHasta.Value, "dd/mm/YYYY")

param = arr()

strCadena = "SELECT v.`id_cliente`,v.`ncliente`,v.`direccion`,if (v.`id_moneda`='00002',sum(v.`total`*v.`tc`),sum(v.`total`)) as total,sum((v.`total`- function_pago_factura(v.`id_venta`,CURDATE(),v.id_moneda,v.ruc))) as saldo,'" & Val(Me.txtTipo_cambio.Text) & "' " & _
" From movimiento_venta v where v.ruc='" & KEY_RUC & "' and v.id_doc In('0001','0003','0007','0054') AND  v.`id_forma_pago`='02' and (v.`total`-function_pago_factura(v.`id_venta`,CURDATE(),v.id_moneda,v.ruc))>0 Group By v.`id_cliente`"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptCuentasCobrarResumen", param, App.Path + "\Reportes\")
End Sub

Private Sub cmdseleccionar_Click()


'strCadena = "SELECT id_venta,documento,id_doc,tc,total,(total-function_pago_factura(a.id_venta,'" & KEY_FECHA & "',a.id_moneda,a.ruc)) as saldo FROM movimiento_venta a WHERE    (total-function_pago_factura(a.id_venta,'" & KEY_FECHA & "',a.id_moneda,a.ruc))>0 and  id_forma_pago='01' and id_doc IN('0001','0003') and  ruc='" & KEY_RUC & "'"
'Call ConfiguraRst(strCadena)
'If rst.RecordCount > 0 Then
'   rst.MoveFirst
'   For i = 0 To rst.RecordCount - 1
'A:
   
'       strCadena = "SELECT * FROM movimiento_venta_monto WHERE  id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "' "
'       Call ConfiguraRstT(strCadena)
'       If rstT.RecordCount = 1 Then
       
'          If rstT("forma_pago") = "02" Then
'              strCadena = "UPDATE movimiento_venta_monto SET forma_pago='01' WHERE id_detalle='" & rstT("id_detalle") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
'              CnBd.Execute (strCadena)
'          End If

 '          strCadena = "SELECT * FROM mis_cuentas_det_detalle d,movimiento_venta v WHERE   d.id_detalle=v.id_venta and v.anulado='no' and v.ruc='" & KEY_RUC & "' and   d.id_movimiento='" & rst("id_venta") & "'"
  '         Call ConfiguraRstL(strCadena)
   '        If rstL.RecordCount > 0 Then
    '          GoTo N:
     '      End If

      '     Call put_realizar_pago(rst("id_venta"), rst("id_venta"), Abs(rstT("monto_caja")), rst("id_doc"), rst("tc"), 0)
      '  Else
      '      strCadena = "SELECT * FROM movimiento_venta_monto WHERE monto_caja='" & rst("total") & "' and  id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "' "
      '      Call ConfiguraRstIN(strCadena)
      '      If rstIN.RecordCount = 1 Then
      '      strCadena = "DELETE FROM movimiento_venta_monto WHERE monto_caja<>'" & rst("total") & "' and  id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "' LIMIT 3 "
       '     CnBd.Execute (strCadena)
       '     GoTo A
       '     End If
       'End If
'N:
'       rst.MoveNext
 '      DoEvents
 '  Next i
'End If

'Exit Sub


If Me.HfdPersona.Rows > 0 Then
    
   
   'TODOS CONTADOS
   
   
   Me.lblid_venta.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
   Me.lblDocumento.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 4)
   
   strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM comprobantes "
   Call ConfiguraRst(strCadena)
   Call LlenaDataCombo(Me.DtcComprobanteAsociado)
   
   strCadena = "DELETE FROM comprobante_asociado WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
   Call llenar_vinculados(Me.hgVinculados)
   
   Me.frmblanquear.Visible = True
 End If

End Sub



Private Sub DtcComprobanteAsociado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtSerieAsociada)
End If
End Sub

Private Sub DtcManifiesto_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    strCadena = "CALL CON_CuentaCobrar_LST('19','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.txtrazonsocial.Text) & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Me.DtcComprobante.BoundText & "','" & Me.DtcManifiesto.BoundText & "','" & Me.DtcSerie.BoundText & "','" & KEY_RUC & "')"
    Call llenar_grid(Me.HfdPersona)
    Me.cmdReporte.Enabled = True
    Me.cmdReporteDetallado.Enabled = True

End If

End Sub

Private Sub DtcPeriodo_Change()
 Dim in_fecha As Date
strCadena = "SELECT * FROM con_periodo WHERE id='" & Me.DtcPeriodo.BoundText & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    in_fecha = DateSerial(rst("Ejercicio"), rst("mes") + 1, 0)
    Me.TxtTipoCambio.Text = get_tipo_cambio_dia(in_fecha, "valor_venta")
End If

End Sub

Private Sub DtpHasta_Change()
Me.txtTipo_cambio.Text = get_tipo_cambio_dia(CVDate(Me.DtpHasta.Value), "valor_compra")
End Sub

Private Sub Form_Load()
On Error GoTo salir

CenterForm Me
Me.Top = 50
Me.DtpDesde.Value = KEY_FECHA
Me.DtpHasta.Value = KEY_FECHA

Me.DtpInicio_dudosa.Value = KEY_FECHA
Me.DtpFin_dudosa.Value = KEY_FECHA


strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE id_tipoentidad<>'00012' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM

strCadena = "SELECT  DISTINCT a.id_doc as Codigo,c.doc_des as Descripcion FROM almacen_comprobante a, comprobantes c WHERE  a.id_doc=c.id_doc and   ruc='" & KEY_RUC & "' AND a.venta='si'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComprobante)

strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPeriodoBusqueda)
Me.DtcPeriodoBusqueda.BoundText = get_periodo_actual(KEY_FECHA)
  

strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and   ruc='" & KEY_RUC & "' ORDER BY nombre_completo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCobrador_update)
  Me.txtTipo_cambio.Text = get_tipo_cambio_dia(CVDate(Me.DtpHasta.Value), "valor_compra")

Exit Sub
salir:



End Sub

Public Sub llenar_grid(ByVal Grilla As MSHFlexGrid)
Dim nsaldo As Double
Dim in_facturado As Double
Dim in_acumulado_total As Double
Dim in_operador As String
Dim in_monto_pagado  As Double
'On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.HfdPersona.Rows = 0
    Me.cmdReporte.Enabled = False
    Me.cmdReporteDetallado.Enabled = False
    Me.cmdamortizar.Enabled = False
    Me.cmdGeneral.Enabled = False
    Exit Sub
End If
   Me.cmdReporte.Enabled = True
   Me.cmdReporteDetallado.Enabled = True
   Me.cmdGeneral.Enabled = True
   Me.HfdPersona.Rows = 0
   
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 2500
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 3000
           Grilla.ColWidth(7) = 900
           Grilla.ColWidth(8) = 700
           Grilla.ColWidth(9) = 1300
           Grilla.ColWidth(10) = 1300
           Grilla.ColWidth(11) = 1300
           Grilla.ColWidth(12) = 2800
          ' Grilla.ColWidth(13) = 1700
           
        Next
         
         cabecera = "IDVENTA" & vbTab & "EMISION" & vbTab & "VENCIMIENTO" & vbTab & "FORM .PAGO" & vbTab & "COMPROBANTE" & vbTab & "DNI CLIENTE" & vbTab & "DATOS CLIENTE" & vbTab & "MONEDA" & vbTab & "TC" & vbTab & "  TOTAL   " & vbTab & "SALDO [DOLAR]" & vbTab & "SALDO [SOLES] " & vbTab & "OPERADOR" & vbTab & "REFERENCIA" & vbTab & "%SEGURO"
         Grilla.AddItem cabecera
         
         For k = 1 To 12
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
         
        rst.MoveFirst
        nfactor = 0
        in_acumulado = 0
        in_acumulado_total = 0
        For i = 0 To rst.RecordCount - 1
             
            ' If IsNull(rst("guia")) = True Then
            '    in_operador = Mid(UCase(rst("nombre_completo")), 1, 20)
            ' Else
            '    in_operador = rst("guia") & Space(4) & "[ S/." & rst("tseguro") & " ]"
            ' End If
             
                If KEY_PAIS = KEY_PERU Then
                
                If rst("id_moneda") = "00002" Then
                    
                    in_facturado = rst("total") * Val(Me.txtTipo_cambio.Text)
                    n_saldo_dolar = (rst("total") - rst("pago")) 'rst("pago"))
                    n_saldo = (rst("total") - rst("pago")) * Val(Me.txtTipo_cambio.Text)
                   
                Else
                
                    n_saldo = (rst("total") - rst("pago"))
                    If Val(Me.txtTipo_cambio.Text) <= 0 Then
                        n_saldo_dolar = (rst("total") - rst("pago")) / KEY_CAMBIO
                        in_facturado = rst("total")
                    Else
                        n_saldo_dolar = (rst("total") - rst("pago")) / Val(Me.txtTipo_cambio.Text)
                        in_facturado = rst("total")
                    End If
                    
                
                End If
                in_acumulado_total = in_acumulado_total + in_facturado
                
                
                Else
                     in_facturado = rst("total")
                     in_acumulado_total = in_acumulado_total + in_facturado
                     n_saldo_dolar = (rst("total") - rst("pago"))
                     n_saldo = (rst("total") - rst("pago")) * Val(Me.txtTipo_cambio.Text)
                End If
                
                
                
                
          
             If rst("id_forma_pago") = "01" Then
                in_forma_pago = "CONTADO"
             Else
                in_forma_pago = "CREDITO"
             End If
             
             If rst("id_doc") = "0007" Then
                n_saldo = n_saldo * -1
                n_saldo_dolar = n_saldo_dolar * -1
             End If
             
             Fila = rst("id_venta") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & Format(rst("fecha_vencimiento"), "dd-mm-YYYY") & vbTab & in_forma_pago & vbTab & rst("comprobante") & vbTab & rst("id_cliente") & vbTab & Mid(rst("ncliente"), 1, 40) & vbTab & rst("descripcion") & vbTab & rst("tc") & vbTab & Format(rst("total"), "#,##0.00") & vbTab & Format(n_saldo_dolar, "#,##0.00") & vbTab & Format(n_saldo, "#,##0.00") & vbTab & rst("vendedor")
             Grilla.AddItem Fila
             
                
              
                  in_saldo = in_saldo + n_saldo
                  in_saldo_dolar = in_saldo_dolar + n_saldo_dolar
                  
            If Val(rst("total") - rst("pago")) <> 0 Then
            For k = 10 To 11
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H80C0FF
            Next k
            End If
            
            If rst("anulado") = "si" Then
           
            For k = 8 To 11
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
            Next k
            End If
            rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(in_acumulado_total, "#,##0.00") & vbTab & Format(in_saldo_dolar, "#,##0.00") & vbTab & Format(in_saldo, "#,##0.00") & vbTab & "" & vbTab & "" & vbTab & Format(nfactor, "#,##0.00")
        Grilla.AddItem Fila
        
        
        For k = 9 To 11
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
            Next k
            
   Grilla.ColAlignment(3) = 1
  Grilla.ColAlignment(5) = 1
  Grilla.ColAlignment(7) = 7
  Grilla.ColAlignment(8) = 7
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
         
Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub llenar_recibos(ByVal Grilla As MSHFlexGrid)
Dim nsaldo As Double
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
   

   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 2200
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 4000
           Grilla.ColWidth(6) = 2000
           
           
           
        Next
         cabecera = "IDVENTA" & vbTab & "F.EMISION" & vbTab & "H.REGISTRO" & vbTab & "COMPROBANTE" & vbTab & "DNI CLIENTE" & vbTab & "DATOS CLIENTE" & vbTab & "TOTAL CANCELADO"
         Grilla.AddItem cabecera
         For k = 1 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        nsaldo = 0
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id_movimiento") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & Format(rst("hora"), "HH:mm:ss am/pm") & vbTab & rst("documento") & vbTab & rst("id_cliente") & vbTab & rst("ncliente") & vbTab & Format(rst("total"), "#,##0.00")
             Grilla.AddItem Fila
             nsaldo = nsaldo + rst("total")
             rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL CANCELADO:" & vbTab & Format(nsaldo, "#,##0.00")
        Grilla.AddItem Fila
        For k = 1 To 6
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
            Next k
 ' Grilla.Row = 1
 ' Grilla.col = 0
 ' Grilla.ColSel = 1
 ' Grilla.RowSel = 1
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub llenar_vinculados(ByVal Grilla As MSHFlexGrid)
Dim nsaldo As Double
On Error GoTo salir
strCadena = "SELECT * FROM comprobante_asociado where dni_save='" & KEY_USUARIO & "' and id_venta='" & Val(Me.lblid_venta.Caption) & "'"

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    
    Exit Sub
End If
    Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 500
           Grilla.ColWidth(1) = 2500
           Grilla.ColWidth(2) = 1200
        Next
         cabecera = "ID DETALLE" & vbTab & "COMPROBANTE " & vbTab & "MONTO "
         Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        nsaldo = 0
        in_total = 0
        For i = 0 To rst.RecordCount - 1
        in_total = in_total + rst("monto")
             Fila = rst("id_detalle") & vbTab & get_comprobante_venta(rst("id_asociado")) & vbTab & rst("monto")
             Grilla.AddItem Fila
             rst.MoveNext
        Next i
        
            
            Fila = "" & vbTab & "" & vbTab & Format(in_total, "#,##0.00")
             Grilla.AddItem Fila
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub



Private Sub llenar_letras_vencidas(ByVal Grilla As MSHFlexGrid)
Dim nsaldo As Double
On Error GoTo salir
strCadena = "SELECT id_venta,interes_diferido,monto_interes,interes_diferido,interes_revertido,fecha_emision,fecha_vencimiento,comprobante,id_cliente,ncliente,total,id_referencia,nombre_completo,function_pago_factura(id_venta,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc) as pago " & _
" FROM view_listado_comprobante_vargas WHERE  monto_interes>0 and id_doc='0412' and fecha_vencimiento>='" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
     Exit Sub
End If
   

   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 1100
           Grilla.ColWidth(3) = 1800
           Grilla.ColWidth(4) = 2000
           Grilla.ColWidth(5) = 1000
           Grilla.ColWidth(6) = 1000
           Grilla.ColWidth(7) = 2500
           Grilla.ColWidth(8) = 1200
        Next
        
               
         cabecera = "IDVENTA" & vbTab & "F.EMISION " & vbTab & "F.VENCIMIENTO " & vbTab & "COMPROBANTE " & vbTab & "CLIENTE" & vbTab & "INTERES" & vbTab & "SALDO" & vbTab & "REFERENCIA" & vbTab & "DEVENGADO"
         Grilla.AddItem cabecera
         For k = 0 To 8
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        nsaldo = 0
        nTotal = 0
        For i = 0 To rst.RecordCount - 1
        in_total = in_total + rst("monto_interes")
            If rst("interes_diferido") = "si" Then
               IN_DEVENGADO = "REALIZADO"
            Else
               IN_DEVENGADO = "PENDIENTE"
            End If
             
             Fila = rst("id_venta") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & Format(rst("fecha_vencimiento"), "dd-mm-YYYY") & vbTab & rst("comprobante") & vbTab & rst("ncliente") & vbTab & Format(rst("monto_interes"), "#,##0.00") & vbTab & Format(rst("interes_revertido"), "#,##0.00") & vbTab & get_nota_venta(rst("id_venta"), rst("id_referencia")) & vbTab & IN_DEVENGADO
             Grilla.AddItem Fila
             
             
             
             rst.MoveNext
        Next i
            
'            Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(in_total, "#,##0.00")
 '           Grilla.AddItem Fila
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub



Private Sub HfdPersona_DblClick()
If Me.HfdPersona.Rows > 0 Then
If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
    Procedencia = Selecionar
    
    FrmDetalleventa.Show
End If
End If
End Sub

Private Sub Image1_Click()

Me.frmblanquear.Visible = False

End Sub

Private Sub Image2_Click()
Me.frm_dudosa.Visible = Tr
End Sub

Private Sub Image3_Click()

    frmajusteCobrar.Visible = False

End Sub

Private Sub mdRevertirDudosa_Click()
If Me.HfDudosa.Rows > 0 Then
   If Val(Me.HfDudosa.TextMatrix(Me.HfDudosa.Row, 0)) > 0 Then
     
     If MsgBox("Recuerde Eliminar el Registro Contable", vbInformation + vbYesNo) = vbYes Then
        strCadena = "UPDATE movimiento_venta SET cobranza_dudosa='no',dudosa_contable='no',fecha_cobranza_dudosa='" & Format(Me.DtpFechaContable.Value, "YYYY-mm-dd") & "' WHERE id_venta='" & Val(Me.HfDudosa.TextMatrix(Me.HfDudosa.Row, 0)) & "'"
        CnBd.Execute (strCadena)
        MsgBox "Proceso Exitoso ...", vbInformation
     End If
     
   End If
End If

End Sub

Private Sub Text1_Change()
strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM comprobantes where doc_des LIKE '%" & Trim(Me.Text1.Text) & "%' "
   Call ConfiguraRst(strCadena)
   Call LlenaDataCombo(Me.DtcComprobanteAsociado)
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub




Private Sub txtEmpresa_Change()

End Sub



Private Sub HfdPersona_SelChange()
If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
    Call get_saldo(Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 11)), Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)))
    
 Else
    Me.cmdamortizar.Enabled = False
    Me.cmdHistorial.Enabled = False
End If
End Sub

Private Function get_saldo(ByVal in_saldo As Double, ByVal in_venta As String) As Double
strCadena = "SELECT total,function_pago_factura(id_venta,'" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "',id_moneda,ruc) as saldo FROM view_listado_comprobante_vargas WHERE id_venta='" & in_venta & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    If (rst("total") - rst("saldo")) > 0 Then
    Me.cmdamortizar.Enabled = True
    Me.cmdHistorial.Enabled = True
    Me.cmdReporte.Enabled = True
    Me.cmdReporteDetallado.Enabled = True
    Else
    Me.cmdamortizar.Enabled = False
    Me.cmdHistorial.Enabled = True
    Me.cmdReporte.Enabled = True
    Me.cmdReporteDetallado.Enabled = True
    End If
Else
    Me.cmdamortizar.Enabled = False
    Me.cmdHistorial.Enabled = True
    Me.cmdReporte.Enabled = True
    Me.cmdReporteDetallado.Enabled = True
End If
End Function

Private Sub timier_despacho_Timer()
'Call pendientes_despacho
End Sub

Private Sub txtBuscarDudosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,dias,ncliente,comprobante,fecha_cobranza_dudosa,moneda,(total-function_pago_factura(id_venta,'" & KEY_FECHA & "',id_moneda,ruc)) as saldo,dudosa_contable FROM view_dudosa WHERE " & _
        " comprobante LIKE '%" & Trim(Me.txtBuscarDudosa.Text) & "%' and  dias>=365  and   ruc='" & KEY_RUC & "' "
        Call llenar_dudosa(HfDudosa)
End If
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Procedencia = Selecionar
   FrmPersona.Show
   Exit Sub
End If
End Sub

Private Sub txtCuentaGanancia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = seleccionar_per
    FrmPlanContableCuentas.Show
    Exit Sub
End If
End Sub

Private Sub txtCuentaPerdida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = seleccionar_otro
    FrmPlanContableCuentas.Show
    Exit Sub
End If
End Sub

Private Sub txtCuentaPrincipal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmPlanContableCuentas.Show
    Exit Sub
End If
End Sub

Private Sub txtDudosaRazonSocial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,dias,ncliente,comprobante,fecha_cobranza_dudosa,moneda,(total-function_pago_factura(id_venta,'" & KEY_FECHA & "',id_moneda,ruc)) as saldo,dudosa_contable FROM view_dudosa WHERE " & _
        " ncliente LIKE  '%" & Trim(Me.txtDudosaRazonSocial.Text) & "%' and  dias>=365  and   ruc='" & KEY_RUC & "' "
        Call llenar_dudosa(HfDudosa)
End If
End Sub

Private Sub txtDudosaRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,dias,ncliente,comprobante,fecha_cobranza_dudosa,moneda,(total-function_pago_factura(id_venta,'" & KEY_FECHA & "',id_moneda,ruc)) as saldo,dudosa_contable FROM view_dudosa WHERE " & _
        " id_cliente = '" & Trim(Me.txtDudosaRuc.Text) & "' and  dias>=365  and   ruc='" & KEY_RUC & "' "
        Call llenar_dudosa(HfDudosa)
End If

End Sub

Private Sub txtMonto_pago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "INSERT INTO comprobante_asociado(`dni_save`,`id_asociado`,`id_venta`,monto,`ruc`)values " & _
    "('" & KEY_USUARIO & "','" & Val(Me.lblid_asociado.Caption) & "','" & Val(Me.lblid_venta.Caption) & "','" & Val(Me.txtMonto_pago.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        Call llenar_vinculados(Me.hgVinculados)
End If
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call llenar_numero
End If
End Sub
Private Sub llenar_numero()
  
strCadena = "CALL CON_CuentaCobrar_LST('13','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.txtrazonsocial.Text) & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Me.DtcComprobante.BoundText & "','" & Trim(Me.TxtNumero.Text) & "','" & Me.DtcSerie.BoundText & "','" & KEY_RUC & "')"
Call llenar_grid(Me.HfdPersona)
  
  Me.cmdReporte.Enabled = True
  Me.cmdReporteDetallado.Enabled = True
End Sub

Private Sub TxtNumeroAsociada_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtNumeroAsociada.Text = Format(Val(Me.TxtNumeroAsociada.Text), "000000")
    strCadena = "SELECT * FROM movimiento_venta where id_doc='" & Me.DtcComprobanteAsociado.BoundText & "' and  serie='" & Trim(Me.TxtSerieAsociada.Text) & "' and numero='" & Trim(Me.TxtNumeroAsociada.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.lblid_asociado.Caption = rst("id_venta")
        Me.txtMonto_pago.Text = rst("total")
        Call Resalta(Me.txtMonto_pago)
        Exit Sub
        
    End If
    
    
    
End If
End Sub

Private Sub TxtRazonSocial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
strCadena = "CALL CON_CuentaCobrar_LST('15','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.txtrazonsocial.Text) & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Me.DtcComprobante.BoundText & "','" & Trim(Me.TxtNumero.Text) & "','" & Me.DtcSerie.BoundText & "','" & KEY_RUC & "')"
Call llenar_grid(Me.HfdPersona)
    Me.cmdReporte.Enabled = True
    Me.cmdReporteDetallado.Enabled = True
End If

End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
Dim in_ruc As String
If KeyAscii = 13 Then

strCadena = "CALL CON_CuentaCobrar_LST('14','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.txtrazonsocial.Text) & "','" & Format(Me.DtpDesde.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "','" & Me.DtcComprobante.BoundText & "','" & Trim(Me.TxtNumero.Text) & "','" & Me.DtcSerie.BoundText & "','" & KEY_RUC & "')"
Call llenar_grid(Me.HfdPersona)


Me.cmdReporte.Enabled = True
Me.cmdReporteDetallado.Enabled = True

End If
End Sub

Public Sub get_estado_cuenta(ByVal in_cliente As String)

End Sub




Private Sub TxtSerieAsociada_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    
End If
End Sub

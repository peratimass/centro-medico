VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmVentasPagos 
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
   Begin VB.Frame frmrecibos_unificados 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LISTADO DE RECIBOS"
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
      Height          =   4335
      Left            =   16440
      TabIndex        =   95
      Top             =   360
      Width           =   3615
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfRecibos 
         Height          =   3135
         Left            =   120
         TabIndex        =   96
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5530
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
      Begin VitekeySoft.ChameleonBtn cmdPagarRecibos 
         Height          =   555
         Left            =   120
         TabIndex        =   97
         Top             =   3600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   979
         BTYPE           =   5
         TX              =   "GENERAR COMPROBANTE [SUNAT]"
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
         MICON           =   "frmVentasPagos.frx":0000
         PICN            =   "frmVentasPagos.frx":001C
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
   Begin VB.Frame frmretencion 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   7440
      TabIndex        =   88
      Top             =   4680
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox txtserie_retencion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   91
         Text            =   "000"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtNumero_retencion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   90
         Text            =   "000000"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtid_retencion 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   720
         MaxLength       =   80
         TabIndex        =   89
         Top             =   960
         Visible         =   0   'False
         Width           =   300
      End
      Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
         Height          =   180
         Left            =   3960
         TabIndex        =   92
         Top             =   120
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   318
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmVentasPagos.frx":3664
         PICN            =   "frmVentasPagos.frx":3680
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "COMPROBANTE DE RETENCION"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1050
         TabIndex        =   94
         Top             =   120
         Width           =   2355
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "NUMERO :"
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
         Left            =   360
         TabIndex        =   93
         Top             =   600
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PROYECTOS"
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
      Height          =   7455
      Left            =   16440
      TabIndex        =   84
      Top             =   360
      Width           =   3615
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfProyectos 
         Height          =   3135
         Left            =   120
         TabIndex        =   85
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5530
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
      Begin VitekeySoft.ChameleonBtn cmdprocesarSunat 
         Height          =   555
         Left            =   120
         TabIndex        =   86
         Top             =   6720
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   979
         BTYPE           =   5
         TX              =   "GENERAR COMPROBANTE [SUNAT]"
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
         MICON           =   "frmVentasPagos.frx":6534
         PICN            =   "frmVentasPagos.frx":6550
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfServicios 
         Height          =   3135
         Left            =   120
         TabIndex        =   87
         Top             =   3480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   5530
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
   End
   Begin VB.TextBox txtBuscarForma 
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
      Left            =   6675
      TabIndex        =   80
      Top             =   4680
      Width           =   615
   End
   Begin VB.Frame frmnota_credito 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2775
      Left            =   7440
      TabIndex        =   72
      Top             =   4680
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox txtidMemorandum 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         MaxLength       =   80
         TabIndex        =   81
         Top             =   2040
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtNumero_nota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   75
         Text            =   "000000"
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtSerie_nota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   74
         Text            =   "000"
         Top             =   840
         Width           =   495
      End
      Begin VitekeySoft.ChameleonBtn cmdCerrar_frame_nota 
         Height          =   180
         Left            =   3960
         TabIndex        =   73
         Top             =   120
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   318
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmVentasPagos.frx":9B98
         PICN            =   "frmVentasPagos.frx":9BB4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcTipoMemorandum 
         Height          =   315
         Left            =   360
         TabIndex        =   82
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
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
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "NUMERO :"
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
         Left            =   360
         TabIndex        =   83
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblsaldo_nota 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1200
         TabIndex        =   71
         Top             =   2160
         Width           =   2205
      End
      Begin VB.Label lblReferencia_nota 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   885
         Left            =   1185
         TabIndex        =   79
         Top             =   1200
         Width           =   2205
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "SALDO :"
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
         Left            =   525
         TabIndex        =   78
         Top             =   2160
         Width           =   525
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCIA:"
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
         Left            =   195
         TabIndex        =   77
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lbldetalle_nota 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "INGRESE NOTA DE CREDITO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   315
         TabIndex        =   76
         Top             =   120
         Width           =   2130
      End
   End
   Begin VB.Frame frmtarjeta 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   7440
      TabIndex        =   66
      Top             =   4680
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtOperacionTarjeta 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   70
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox TxtNumeroTargeta 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         MaxLength       =   80
         TabIndex        =   69
         Top             =   840
         Width           =   1215
      End
      Begin VitekeySoft.ChameleonBtn cmdcerrarTarjeta 
         Height          =   180
         Left            =   3960
         TabIndex        =   67
         Top             =   120
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   318
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmVentasPagos.frx":CA68
         PICN            =   "frmVentasPagos.frx":CA84
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcTarjeta 
         Height          =   330
         Left            =   240
         TabIndex        =   68
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
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
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      Left            =   6000
      MaxLength       =   80
      TabIndex        =   65
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame frmbanco 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2745
      Left            =   7440
      TabIndex        =   54
      Top             =   4680
      Visible         =   0   'False
      Width           =   4260
      Begin VB.TextBox txtBanco 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   57
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox txtCheque 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
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
         Left            =   1440
         TabIndex        =   56
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txtbuscarbanco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   55
         Top             =   1680
         Width           =   2415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfBancos 
         Height          =   1575
         Left            =   165
         TabIndex        =   58
         Top             =   45
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2778
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
      Begin VitekeySoft.ChameleonBtn cmdcerrarbanco 
         Height          =   180
         Left            =   3960
         TabIndex        =   59
         Top             =   120
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   318
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmVentasPagos.frx":F938
         PICN            =   "frmVentasPagos.frx":F954
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BANCO :"
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
         Left            =   285
         TabIndex        =   62
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N� CHEQUE:"
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
         Left            =   0
         TabIndex        =   61
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BUSCAR :"
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
         Left            =   210
         TabIndex        =   60
         Top             =   1680
         Width           =   705
      End
   End
   Begin VB.TextBox txtTipoFlujo 
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
      Left            =   6555
      TabIndex        =   47
      Top             =   5925
      Width           =   855
   End
   Begin VB.TextBox txtFormaPago 
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
      Left            =   6675
      TabIndex        =   46
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txtId_recibo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5520
      MaxLength       =   80
      TabIndex        =   45
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   855
      Left            =   16560
      TabIndex        =   39
      Top             =   8280
      Width           =   855
      _ExtentX        =   1508
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmVentasPagos.frx":12808
      PICN            =   "frmVentasPagos.frx":12824
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame framevehiculo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   50
      TabIndex        =   34
      Top             =   7560
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox TxtMontoReal 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Height          =   285
         Left            =   3720
         MaxLength       =   80
         TabIndex        =   37
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtmontovehiculo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
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
         Height          =   285
         Left            =   1710
         MaxLength       =   80
         TabIndex        =   35
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label valorvehiculo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR VEHICULO :"
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
         TabIndex        =   36
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.TextBox txtid_producto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5520
      MaxLength       =   80
      TabIndex        =   33
      Top             =   8040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkElegir 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ELEGIR VEHICULO / REPUESTO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1680
      TabIndex        =   32
      Top             =   8040
      Width           =   2775
   End
   Begin VB.TextBox txtOperacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      Left            =   1785
      MaxLength       =   80
      TabIndex        =   31
      Top             =   5160
      Width           =   4695
   End
   Begin VB.TextBox TxtObservacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   795
      Left            =   1785
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   6720
      Width           =   5655
   End
   Begin VB.TextBox TxtDireccion 
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1785
      MaxLength       =   80
      TabIndex        =   14
      Top             =   2100
      Width           =   5415
   End
   Begin VB.TextBox TxtCliente 
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1785
      MaxLength       =   80
      TabIndex        =   13
      Top             =   1740
      Width           =   5415
   End
   Begin VB.TextBox TxtRuc 
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1785
      MaxLength       =   80
      TabIndex        =   12
      Top             =   1380
      Width           =   1215
   End
   Begin VB.TextBox TxtSaldo 
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1785
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   11
      Top             =   2895
      Width           =   1935
   End
   Begin VB.TextBox TxtMontoPago 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   1785
      MaxLength       =   80
      TabIndex        =   0
      Top             =   3300
      Width           =   1935
   End
   Begin VB.TextBox TxtTc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      Left            =   5970
      MaxLength       =   80
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox tXTIdVenta 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5520
      MaxLength       =   80
      TabIndex        =   6
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox TxtSerie 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   2
      Text            =   "000"
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox TxtNumeroDoc 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4920
      MaxLength       =   80
      TabIndex        =   1
      Text            =   "000000"
      Top             =   720
      Width           =   2175
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   330
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
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
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   330
      Left            =   3720
      TabIndex        =   4
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      ForeColor       =   8388608
      ListField       =   "0000�"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCuotas 
      Height          =   2670
      Left            =   7800
      TabIndex        =   7
      Top             =   6480
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4710
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfVinculados 
      Height          =   5535
      Left            =   7800
      TabIndex        =   8
      Top             =   360
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9763
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
   Begin MSComCtl2.DTPicker DtpEmision 
      Height          =   300
      Left            =   3885
      TabIndex        =   9
      Top             =   1365
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
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
      Format          =   182321153
      CurrentDate     =   41130
   End
   Begin MSDataListLib.DataCombo DtcCuentas 
      Height          =   315
      Left            =   1785
      TabIndex        =   15
      Top             =   3720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
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
   Begin MSDataListLib.DataCombo DtcMoneda 
      Height          =   330
      Left            =   1785
      TabIndex        =   16
      Top             =   2510
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
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
   Begin MSComCtl2.DTPicker DtpValor 
      Height          =   300
      Left            =   5910
      TabIndex        =   17
      Top             =   1365
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
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
      Format          =   182321153
      CurrentDate     =   41130
   End
   Begin VitekeySoft.ChameleonBtn cmdSave 
      Height          =   855
      Left            =   17445
      TabIndex        =   40
      Top             =   8280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "PROCESAR"
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
      MICON           =   "frmVentasPagos.frx":12C76
      PICN            =   "frmVentasPagos.frx":12C92
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdPrinter 
      Height          =   855
      Left            =   18330
      TabIndex        =   41
      Top             =   8280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "IMPRIMIR"
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
      MICON           =   "frmVentasPagos.frx":162DA
      PICN            =   "frmVentasPagos.frx":162F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdExit 
      Height          =   855
      Left            =   19200
      TabIndex        =   42
      Top             =   8280
      Width           =   855
      _ExtentX        =   1508
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmVentasPagos.frx":188C7
      PICN            =   "frmVentasPagos.frx":188E3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcFormaPago 
      Height          =   315
      Left            =   1785
      TabIndex        =   48
      Top             =   4200
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
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
   Begin MSDataListLib.DataCombo DtcFlujo 
      Height          =   315
      Left            =   1785
      TabIndex        =   49
      Top             =   5685
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
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
   Begin MSDataListLib.DataCombo DtcForma_pago_detalle 
      Height          =   315
      Left            =   1785
      TabIndex        =   52
      Top             =   4680
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
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
   Begin MSDataListLib.DataCombo DtcVentanilla 
      Height          =   315
      Left            =   1785
      TabIndex        =   63
      Top             =   6120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
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
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008080FF&
      Caption         =   "VENTANILLA :"
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
      Height          =   195
      Left            =   300
      TabIndex        =   64
      Top             =   6240
      Width           =   1410
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008080FF&
      Caption         =   "FORMA PAGO:"
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
      Height          =   195
      Left            =   240
      TabIndex        =   53
      Top             =   4800
      Width           =   1485
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008080FF&
      Caption         =   "TIPO DE FLUJO :"
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
      Height          =   195
      Left            =   300
      TabIndex        =   51
      Top             =   5760
      Width           =   1410
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008080FF&
      Caption         =   "CONCEPTO PAGO:"
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
      Height          =   195
      Left            =   225
      TabIndex        =   50
      Top             =   4320
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "CUENTA DESTINO :"
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
      Left            =   270
      TabIndex        =   44
      Top             =   3795
      Width           =   1440
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "LISTADO DE COMPROBANTE RELACIONADOS"
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
      Left            =   15120
      TabIndex        =   43
      Top             =   120
      Width           =   4935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   7800
      X2              =   14760
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   7680
      X2              =   7680
      Y1              =   240
      Y2              =   8160
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "PAGOS RELACIONADOS"
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
      Left            =   7800
      TabIndex        =   38
      Top             =   6120
      Width           =   7335
   End
   Begin VB.Label lblOperacion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "N� OPERACION:"
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
      TabIndex        =   30
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "LISTADO DE COMPROBANTE RELACIONADOS"
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
      Left            =   7800
      TabIndex        =   29
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "GLOSA :"
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
      Left            =   240
      TabIndex        =   28
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "MONEDA :"
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
      TabIndex        =   26
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "RUC/DNI :"
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
      TabIndex        =   25
      Top             =   1380
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "RAZON SOCIAL:"
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
      TabIndex        =   24
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "DIRECCION :"
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
      TabIndex        =   23
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "SALDO :"
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
      TabIndex        =   22
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "A CUENTA  :"
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
      TabIndex        =   21
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T.C:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   5505
      TabIndex        =   20
      Top             =   2580
      Width           =   315
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMISION:"
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
      Left            =   3045
      TabIndex        =   19
      Top             =   1440
      Width           =   765
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR:"
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
      Left            =   5295
      TabIndex        =   18
      Top             =   1395
      Width           =   555
   End
   Begin VB.Label lblAnulado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   885
      Left            =   1680
      TabIndex        =   5
      Top             =   8280
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   120
      Top             =   120
      Width           =   7095
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "frmVentasPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mostrar As Boolean
Dim codigo_P As String
Public Procedencia As EnumProcede
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Dim tTotal As Double
tTotal = 0

strCadena = "SELECT      movimiento_caja.doc_cod,movimiento_caja.fecha_valor, Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, " & _
"movimiento_caja.descripcion_per , movimiento_caja.Monto,movimiento_caja.anulado " & _
"FROM         movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod " & _
" WHERE  movimiento_caja.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "'  ORDER BY  movimiento_caja.numero DESC"
Call ConfiguraRst(strCadena)
 
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
  
  N = 1
  
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 600
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 1300
           Grilla.ColWidth(3) = 600
           Grilla.ColWidth(4) = 800
           Grilla.ColWidth(5) = 3000
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 0
           Grilla.ColWidth(8) = 0
          Next
         cabecera = "ITEM" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "SERIE" & vbTab & "NUMERO" & vbTab & "PERSONA" & vbTab & "MONTO" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 8
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = Fila & str(rst.RecordCount - i + 1) & vbTab & rst("fecha_valor") & vbTab & rst("doc_abrev") & vbTab & rst("serie") & vbTab & rst("numero") & vbTab & rst("descripcion_per") & vbTab & Format(rst("monto"), "#,##0.00")
            If (Fila = "") Then
                X = 1
            End If
            
          Grilla.AddItem Fila
            If (Trim(rst("anulado")) = "si") Then
                            For k = 0 To 8
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &H8080FF
                            Next k
        End If
        Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub ChkAdelantado_Click()

End Sub





Private Sub ChameleonBtn1_Click()
Me.frmretencion.Visible = False
End Sub

Private Sub ChameleonBtn2_Click()

End Sub

Private Sub chkElegir_Click()
If Me.chkElegir.Value = 1 Then
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
Else
    Me.TxtObservacion.Text = ""
End If
End Sub


Private Sub cmdCerrar_frame_nota_Click()
Me.frmnota_credito.Visible = False
End Sub

Private Sub cmdcerrarbanco_Click()
Me.frmbanco.Visible = False
End Sub

Private Sub cmdCerrarTarjeta_Click()
Me.frmtarjeta.Visible = False
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdnuevo_Click()
Call nuevo
End Sub

Private Sub cmdPagarRecibos_Click()
FrmVentas.Show
Call FrmVentas.put_load_recibos(Trim(Me.TxtRuc.Text))
End Sub

Private Sub cmdPrinter_Click()
 Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Trim(Me.TxtSerie.Text), Trim(Me.TxtNumeroDoc.Text), "00001", Val(Me.txtId_recibo.Text))
 
        Exit Sub
End Sub

Private Sub put_comprobante(ByVal in_recibo As String)

'strCadena = "SELECT id_movimiento FROM mis_cuentas_det_detalle WHERE id_detalle='" & Val(Me.txtid_recibo.Text) & "'"
'Call ConfiguraRst(strCadena)
'If rst.RecordCount > 0 Then
 '  rst.MoveFirst
  ' For i = 0 To rst.RecordCount - 1
       '  strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & rst("id_movimiento") & "' and ruc='" & KEY_RUC & "'"
     '    Call ConfiguraRstK(strCadena)
     '    If rstK.RecordCount > 0 Then
         '   rstK.MoveFirst
      '      For j = 0 To rstK.RecordCount - 1
          '      strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,servicio,obsequio,costo) VALUES " & _
                "('" & KEY_RUC & "','" & rstK("id_unidad") & "','" & Trim(Me.TxtRuc.Text) & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.DtcSerieDoc.BoundText) & "','" & Me.TxtNumeroDoc.Text & "','" & codigoP & "','" & Val(Me.txtCantidad.Text) & "'," & _
                "'" & Val(Me.txtPrecio.Text) & " ','" & in_total_parcial & "','" & Val(Me.txtpeso.Text) & "','" & Trim(Me.TxtIgv.Text) & "','" & Trim(Me.TxtDescripcionProducto.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtServicio.Text) & "','" & in_obsequio & "','" & get_precio_costo(codigoP) & "')"
           '     CnBd.Execute (strCadena)
            '    rstK.MoveNext
     '       Next j
    '     End If
         
        
        
        'Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcSerieDoc.BoundText, Me.DtcTipoDoc.BoundText, Trim(Me.txtformato_impresion.Text))
  ' Next i
'End If


End Sub


Private Sub cmdprocesarSunat_Click()
FrmVentas.Show
Call FrmVentas.put_load_proyecto(Me.txtId_recibo.Text, Trim(Me.TxtRuc.Text))
End Sub

Private Sub cmdsave_Click()

If verificar_cierre_caja(Format(Me.DtpEmision.Value, "dd-mm-YYYY")) = 1 Then
    MsgBox "AVISO IMPORTANTE..." + Chr(13) + Chr(13) + "CAJA CONTABLE YA CERRADA.", vbInformation, KEY_VENDEDOR
    Exit Sub
End If
If Val(Me.TxtTc.Text) < 1 And Me.DtcMoneda.BoundText = "00002" Then
   MsgBox "Es Necesario INGRESAR UN TIPO CAMBIO " + Chr(13) + "Ingrese al Modulo de Tipo de Cambio en el Menu.", vbInformation, KEY_VENDEDOR
   Exit Sub
End If


If Val(Me.TxtTc.Text) < 1 And Me.DtcMoneda.BoundText = "00001" Then
   
   Me.TxtTc.Text = 1
   
End If



strCadena = "SELECT id_moneda FROM mis_cuentas WHERE id_cuenta='" & Me.DtcCuentas.BoundText & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   If Me.DtcMoneda.BoundText <> rst("id_moneda") Then
       If MsgBox("El Tipo de moneda es Diferente" + Chr(13) + "Desea realizar la conversion al TC: " + Trim(Me.TxtTc.Text), vbQuestion + vbYesNo) = vbYes Then
            
          If Me.DtcMoneda.BoundText = "00001" Then
             Me.DtcMoneda.BoundText = rst("id_moneda")
             If Val(Me.TxtTc.Text) > 0 Then
                Me.TxtMontoPago.Text = Format(Val(Me.TxtMontoPago.Text) / Val(Me.TxtTc.Text), "###00.00000000")
                Exit Sub
            Else
                MsgBox "Ingrese el Tipo Cambio", vbInformation
                Exit Sub
            End If
            
          Else
                    Me.DtcMoneda.BoundText = rst("id_moneda")
                    If Val(Me.TxtTc.Text) > 0 Then
                        Me.TxtMontoPago.Text = Format(Val(Me.TxtMontoPago.Text) * Val(Me.TxtTc.Text), "###00.00")
                        Exit Sub
                    Else
                        MsgBox "Ingrese el Tipo Cambio", vbInformation
                        Exit Sub
                    End If
          End If
       
       Else
        Exit Sub
       End If
   End If
End If


 Call Save_v2
 
 If KEY_PROYECTO = "si" Then
    Me.cmdprocesarSunat.Visible = True
 End If
End Sub

Public Sub agregar_producto()

End Sub

Private Sub DtcCuentas_Change()
Call load_saldo_comprobante
End Sub

Private Sub DtcCuentas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.txtOperacion.Visible = True Then
        Call Resalta(Me.txtOperacion)
    Else
        Call Resalta(Me.TxtObservacion)
    End If
End If
End Sub
Private Sub load_saldo_comprobante()
Dim ssaldo As Double, residuo As Single, sitf As Single
ssaldo = Val(Me.TxtMontoPago.Text)
strCadena = "SELECT * FROM mis_cuentas WHERE id_cuenta='" & Me.DtcCuentas.BoundText & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst("id_moneda") <> Me.DtcMoneda.BoundText Then
    If rst("id_moneda") = "00001" Then
        Me.TxtSaldo.Text = Format(ssaldo * Val(Me.TxtTc.Text), "###0.00")
        Me.TxtMontoPago.Text = Format(ssaldo * Val(Me.TxtTc.Text), "###0.00")
        Me.DtcMoneda.BoundText = rst("id_moneda")
    Else
        Me.TxtSaldo.Text = Format(ssaldo / Val(Me.TxtTc.Text), "###0.00")
        Me.TxtMontoPago.Text = Format(ssaldo / Val(Me.TxtTc.Text), "###0.00")
        Me.DtcMoneda.BoundText = rst("id_moneda")
    End If
Else
    
End If
residuo = Val(Me.TxtMontoPago.Text) Mod 1000


strCadena = "SELECT id_registro as Codigo, CONCAT(descripcion,'-',observacion) as Descripcion FROM forma_pago_detalle  WHERE   id_cuenta_caja='" & Me.DtcCuentas.BoundText & "' AND ruc='" & KEY_RUC & "' AND estado='si' ORDER BY id_detalle"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcForma_pago_detalle)

End Sub
Private Sub DtcForma_pago_Click(Area As Integer)

End Sub

Private Sub DtcForma_pago_KeyPress(KeyAscii As Integer)

End Sub

Private Sub DtcForma_pago_KeyUp(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub DtcForma_pago_detalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim in_forma_pago As String
in_forma_pago = get_forma_pago_detalle(Me.DtcForma_pago_detalle.BoundText)
Me.frmtarjeta.Visible = False
Me.frmbanco.Visible = False
Me.DtcTipoMemorandum.Visible = False

If in_forma_pago = "12" Then
       strCadena = "SELECT * FROM entidadfinanciera   ORDER BY descripcion"
        Call llenar_bancos(Me.HfBancos)
        Me.frmbanco.Visible = True
        Exit Sub
End If

If in_forma_pago = "03" Or in_forma_pago = "04" Then
       Me.frmtarjeta.Visible = True
       Me.DtcTarjeta.SetFocus
        Exit Sub
End If

If in_forma_pago = "13" Then
       Me.frmnota_credito.Visible = True
       Me.lbldetalle_nota.Caption = "INGRESE NOTA DE CREDITO"
       Me.txtSerie_nota.Text = ""
       Me.txtNumero_nota.Text = ""
       Me.lblReferencia_nota.Caption = ""
       Me.lblsaldo_nota.Caption = ""
       Call Resalta(Me.txtSerie_nota)
       Exit Sub
End If
If in_forma_pago = "20" Then
    Me.frmretencion.Visible = True
    Me.txtserie_retencion.Text = ""
    Me.txtNumero_retencion.Text = ""
    Call Resalta(Me.txtserie_retencion)
    Exit Sub
End If

If in_forma_pago = "14" Then
       Me.frmnota_credito.Visible = True
        Me.DtcTipoMemorandum.Visible = True
        
        strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM view_almacen_comprobante_ultimate WHERE id_doc in ('0415','0416') and ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcTipoMemorandum)
        
        Me.lbldetalle_nota.Caption = "INGRESE NUMERO MEMO"
        Me.txtSerie_nota.Text = ""
        Me.txtNumero_nota.Text = ""
        Me.lblReferencia_nota.Caption = ""
        Me.lblsaldo_nota.Caption = ""
        Call Resalta(Me.txtSerie_nota)
        Exit Sub
End If


End If
End Sub
Public Sub llenar_bancos(ByVal Grilla As MSHFlexGrid)
Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
   Grilla.Rows = 0
   Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 400
        Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
        c = 2
        NumeroCampo = 2
            
        For i = 0 To rstI.RecordCount - 1
          estado = Chr(168)
          descripcion = ""
            
            
          Fila = rstI("Codigo") & vbTab & rstI("Descripcion") & vbTab & estado
          Grilla.AddItem Fila
        
        If c = NumeroCampo Then
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = c '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
                            
                        End With
        End If
        Fila = ""
          
          rstI.MoveNext
      Next i
    
End Sub

Private Sub DtcTarjeta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call Resalta(Me.TxtNumeroTargeta)
End If
End Sub

Private Sub DtcTipoDoc_Change()
Call load_comprobante(Me.DtcTipoDoc.BoundText)
End Sub

Private Sub DtcTipoDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Me.DtcTipoDoc.BoundText) <> "0001" And Trim(Me.DtcTipoDoc.BoundText) <> "0003" Then
        Call Resalta(Me.TxtSerie)
    Else
        MsgBox "Srta:" + Space(1) + KEY_VENDEDOR + Space(1) + "Solo Salida de Dinero", vbInformation
        Me.DtcTipoDoc.SetFocus
    End If
    
    
End If
End Sub

Public Sub nuevo()

strCadena = "SELECT A.id_doc as Codigo,C.doc_abrev as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc='0054' and  A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' ORDER BY C.doc_abrev  "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
    
  
  strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0054' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
    
    Me.TxtSerie.Text = rst("serie")
    Me.TxtNumeroDoc.Text = rst("numero")
  Else
    MsgBox "CREE UN RECIBO DE INGRESO", vbInformation, KEY_EMPRESA
    
    Me.cmdSave.Enabled = False
    Exit Sub
  End If
  
    Me.frmnota_credito.Visible = False
    Me.TxtRuc.Text = ""
    Me.TxtCliente.Text = ""
    Me.TxtDireccion.Text = ""
    Me.TxtObservacion.Text = ""
    Me.TxtMontoPago.Text = "0.00"
    Me.DtpEmision.Value = KEY_FECHA
    Me.DtpValor.Value = KEY_FECHA
    Me.tXTIdVenta.Text = ""
    Me.TxtMontoPago.Text = ""
    Me.chkElegir.Value = 0
    Call LlenarVinculados(Me.HfVinculados, "")
    'Call LlenarCuotas(Me.HfCuotas, -1)
    Call Resalta(Me.TxtRuc)
    
    Me.cmdPrinter.Enabled = False
    Me.cmdSave.Enabled = False
    
    
    
    Me.lblAnulado.Visible = False

End Sub



Private Sub DtpEmision_Change()
Me.TxtTc.Text = get_tipo_cambio_dia(CVDate(Me.DtpEmision.Value), "valor_compra")
Me.DtpValor.Value = Me.DtpEmision.Value


End Sub

Private Sub Form_Activate()
Call Resalta(Me.TxtMontoPago)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub put_seleccion_null(ByVal in_cliente As String, ByVal in_venta As String)
 strCadena = "SELECT * FROM movimiento_venta WHERE seleccion='si'  and ruc='" & KEY_RUC & "'"
 Call ConfiguraRstlocal(strCadena)
 If rstLocal.RecordCount > 0 Then
    rstLocal.MoveFirst
    For i = 0 To rstLocal.RecordCount - 1
        strCadena = "UPDATE movimiento_venta SET seleccion='no',dni_use='0' WHERE id_venta='" & rstLocal("id_venta") & "' LIMIT 1 "
        CnBd.Execute (strCadena)
        rstLocal.MoveNext
    Next i
 End If
 tXTIdVenta.Text = in_venta
 
 strCadena = "UPDATE movimiento_venta SET seleccion='si',dni_use='" & KEY_USUARIO & "' WHERE id_venta='" & Val(in_venta) & "' LIMIT 1"
 CnBd.Execute (strCadena)

End Sub


Private Sub Form_Load()
  CenterForm Me
  Me.Top = 100
  Me.TxtTc.Text = KEY_CAMBIO_COMPRA
 Me.DtpEmision.Value = KEY_FECHA
 Me.DtpValor.Value = KEY_FECHA
 strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  Me.DtcAlmacen.Enabled = False
  
  strCadena = "SELECT id_doc as Codigo,doc_des Descripcion FROM view_comprobante_almacen WHERE id_doc IN ('0054','0415') and  ruc='" & KEY_RUC & "' ORDER BY id_doc ASC  "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
    
    strCadena = "SELECT id as Codigo, descripcion as Descripcion FROM targeta ORDER BY id ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTarjeta)
  
  strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcVentanilla)
  If KEY_VENTANILLA <> "" Then
     Me.DtcVentanilla.BoundText = KEY_VENTANILLA
  Else
     Me.DtcVentanilla.BoundText = KEY_ALM
  End If
  
  
strCadena = "SELECT id as Codigo,Descripcion  as Descripcion FROM vw_mediopago_nombre  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFormaPago)

strCadena = "SELECT id as Codigo,Nombre  as Descripcion FROM adm_flujocaja ORDER BY Nombre  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFlujo)
  
  strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM moneda  ORDER BY id_moneda ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMoneda)
  
  
  strCadena = "SELECT id_cuenta as Codigo,cuenta as Descripcion FROM view_mis_cuentas_contable WHERE ruc='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCuentas)
  
  
  
  
    If FrmReporteRegistroVentas.Procedencia = 1 Then
        FrmReporteRegistroVentas.Procedencia = Neutro
        strCadena = "SELECT * FROM movimiento_venta  WHERE id_venta='" & Val(FrmReporteRegistroVentas.HfdPersona.TextMatrix(FrmReporteRegistroVentas.HfdPersona.Row, 0)) & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            in_cliente = rst("id_cliente")
            Call put_seleccion_null(rst("id_cliente"), FrmReporteRegistroVentas.HfdPersona.TextMatrix(FrmReporteRegistroVentas.HfdPersona.Row, 0))
            Me.TxtRuc.Text = rst("id_cliente")
            Me.TxtCliente.Text = UCase(rst("ncliente"))
            Me.TxtDireccion.Text = UCase(rst("direccion"))
            Me.TxtObservacion.Text = "COBRO:  " + FrmReporteRegistroVentas.HfdPersona.TextMatrix(FrmReporteRegistroVentas.HfdPersona.Row, 4)
            Me.DtcMoneda.BoundText = rst("id_moneda")
            Call LlenarVinculados(Me.HfVinculados, in_cliente)
            'Call llenar_recibos(Me.HfRecibos, Trim(Me.txtruc.Text))
            Me.TxtSaldo.Locked = True
            
        End If
    End If
    
    
   ' If KEY_CONTABILIDAD = "si" Then
       strCadena = "SELECT id_registro as Codigo, cuenta as Descripcion FROM view_forma_pago_conta  WHERE id_cuenta='" & Me.DtcCuentas.BoundText & "' and   ruc='" & KEY_RUC & "'"
   ' Else
   '    strCadena = "SELECT id_registro as Codigo, CONCAT(descripcion,'-',observacion) as Descripcion FROM forma_pago_detalle  WHERE   id_moneda='" & Me.DtcMoneda.BoundText & "' and  id='01' AND ruc='" & KEY_RUC & "' AND estado='si' ORDER BY id_detalle"
   ' End If
    
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcForma_pago_detalle)
    
'DtcForma_pago_detalle
    
    
    
If KEY_PROYECTO = "si" Then
    Call llenar_proyecto(Me.HfProyectos, Trim(Me.TxtRuc.Text))
End If
    
    Me.cmdSave.Enabled = True
    Me.cmdPrinter.Enabled = False
    
End Sub
Private Sub get_monto_pagar()
Dim in_pagar As Double
in_pagar = 0
For m = 0 To Me.HfVinculados.Rows - 1
    If Val(Me.HfVinculados.TextMatrix(m, 0)) > 0 Then
    If Me.HfVinculados.TextMatrix(m, 6) = Chr(254) Then
       in_pagar = in_pagar + Format(Me.HfVinculados.TextMatrix(m, 5), "###0.00")
    
    
    End If
    End If
Next m

Me.TxtMontoPago.Text = Format(in_pagar, "###0.00")

End Sub
Private Sub load_comprobante(ByVal in_doc As String)
strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='" & in_doc & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
  Call ConfiguraRstT(strCadena)
  If rstT.RecordCount > 0 Then
    
    Me.TxtSerie.Text = rstT("serie")
    Me.TxtNumeroDoc.Text = rstT("numero")
  Else
    MsgBox "NO TIENE CREADO UN RECIBO DE INGRESO [0054]", vbInformation, KEY_EMPRESA
    Me.cmdSave.Enabled = False
    Exit Sub
  End If
End Sub
Public Sub LlenarCuotas(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo salir
Dim tmes As Double, tTotal As Double
strCadena = "SELECT  * FROM view_movimiento_venta_pagos WHERE  id_comprobante='" & idVenta & "'"
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
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 2000
           Grilla.ColWidth(4) = 1000
           Grilla.ColWidth(5) = 1400
            
        Next
        cabecera = "IDUNICO" & vbTab & "FECHA PAGO" & vbTab & "HORA PAGO" & vbTab & "COMPROBANTE" & vbTab & "MONTO" & vbTab & "OPERADOR"
        Grilla.AddItem cabecera
         For k = 1 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
      
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_venta") & vbTab & Format(rst("fecha_emision"), "YYYY-mm-dd") & vbTab & Format(rst("hora"), "HH:mm:ss") & vbTab & rst("documento") & vbTab & Format(rst("monto"), "###0.00") & vbTab & rst("nombre_completo")
            Grilla.AddItem Fila
            Fila = ""
      
            tTotal = tTotal + rst("monto")
           
            rst.MoveNext
        Next i
      
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL PAGADO :::>" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
      For k = 1 To 3
            Grilla.col = 3
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H80FF&
      Next k
    
     
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub
Public Sub LlenarLetras(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo salir
Dim tmes As Double, tTotal As Double
strCadena = "UPDATE movimiento_venta_cuotas SET seleccion='no' WHERE id_venta='" & idVenta & "' "
CnBd.Execute (strCadena)
strCadena = "SELECT * FROM movimiento_venta_cuotas WHERE id_venta='" & idVenta & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 800
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 600
        Next
        cabecera = "CODIGO" & vbTab & "LETRA" & vbTab & "F.VENCIMIENTO" & vbTab & "MONTO " & vbTab & "SALDO" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 1 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
      
        tTotal = 0
        
        For i = 0 To rst.RecordCount - 1
            If rst("seleccion") = "no" Then
                If rst("saldo") = 0 Then
                    estado = ""
                Else
                    estado = Chr(168)
                End If
                
            Else
                estado = Chr(254)
            End If
        
            Fila = rst("id") & vbTab & rst("id_cuota") & vbTab & Format(rst("vencimiento"), "dd-mm-YYYY") & vbTab & Format(rst("monto"), "##0.00") & vbTab & Format(rst("saldo"), "###0.00") & vbTab & estado
            Grilla.AddItem Fila
            tTotal = tTotal + rst("saldo")
            With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 5 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
            End With
      
        
        
        
        
        
            
            
            rst.MoveNext
        Next i
      
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "SALDO :::>" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
      For k = 1 To 3
            Grilla.col = 3
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
    
     
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub
Public Sub llenar_proyecto(ByVal Grilla As MSHFlexGrid, ByVal in_cliente As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT * FROM view_proyectos WHERE finalizado='no' and  id_cliente='" & in_cliente & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_inicio"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1800
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 400
           
         Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "MONTO" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 1 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
      
        in_total_pago = 0
        For i = 0 To rst.RecordCount - 1
           
            estado = Chr(168)
            Fila = rst("id_proyecto") & vbTab & rst("descripcion") & vbTab & Format(rst("monto_cobrado"), "#,##0.00") & vbTab & estado
            Grilla.AddItem Fila
            in_total_pago = in_total_pago + rst("monto_cobrado")
            With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 3 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
            End With
        rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "SALDO TOTAL :" & vbTab & Format(in_total_pago, "#,##0.00")
      Grilla.AddItem Fila
      
      For k = 1 To 3
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H80FF&
        Next k
            
  Exit Sub
salir:    MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Public Sub llenar_linea(ByVal Grilla As MSHFlexGrid, ByVal in_dni As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT id_venta,id_proyecto,id_linea,linea,sum(total) as total,id_cliente,ruc FROM view_proyecto_linea WHERE id_cliente='" & in_dni & "'  and ruc='" & KEY_RUC & "' GROUP BY id_linea "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1800
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 400
           
         Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "MONTO" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 1 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
      
        in_total_pago = 0
        For i = 0 To rst.RecordCount - 1
           
            estado = Chr(168)
            in_acumulado = rst("total")
            Fila = rst("id_linea") & vbTab & rst("linea") & vbTab & Format(in_acumulado, "#,##0.00") & vbTab & estado
            Grilla.AddItem Fila
            in_total_pago = in_total_pago + Val(in_acumulado)
            With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 3 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
            End With
        rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "SALDO TOTAL :" & vbTab & Format(in_total_pago, "#,##0.00")
      Grilla.AddItem Fila
      
      For k = 1 To 3
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H80FF&
        Next k
            
  Exit Sub
salir:    MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Public Sub llenar_recibos(ByVal Grilla As MSHFlexGrid, ByVal in_dni As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT id_venta,documento,(total-pago) as saldo FROM view_recibo_pago WHERE seleccion='si' and dni_use='" & KEY_USUARIO & "' and  id_doc='0054' and id_cliente='" & in_dni & "'  and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1800
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 400
           
         Next
        cabecera = "CODIGO" & vbTab & "COMPROBANTE" & vbTab & "MONTO" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 1 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
      
        in_total_pago = 0
        For i = 0 To rst.RecordCount - 1
           
            estado = Chr(168)
            in_acumulado = rst("saldo")
            Fila = rst("id_venta") & vbTab & rst("documento") & vbTab & Format(in_acumulado, "#,##0.00") & vbTab & estado
            Grilla.AddItem Fila
            in_total_pago = in_total_pago + Val(in_acumulado)
            With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 3 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
            End With
        rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "SALDO TOTAL :" & vbTab & Format(in_total_pago, "#,##0.00")
      Grilla.AddItem Fila
      
      For k = 1 To 3
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H80FF&
        Next k
            
  Exit Sub
salir:    MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub



Public Sub LlenarVinculados(ByVal Grilla As MSHFlexGrid, ByVal cPersona As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,comprobante,descripcion,total,pago,seleccion,id_doc FROM view_listado_comprobante_optimizada WHERE total>pago and id_cliente='" & cPersona & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2000
           Grilla.ColWidth(2) = 2200
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 400
         Next
        cabecera = "CODIGO" & vbTab & "EMISION  - VENCE" & vbTab & "COMPROBANTE" & vbTab & "MONEDA" & vbTab & "MONTO TOTAL" & vbTab & "SALDO" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 1 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        in_total_pago = 0
        For i = 0 To rst.RecordCount - 1
            If rst("seleccion") = "si" Then
                estado = Chr(254)
            Else
                estado = Chr(168)
            End If
            in_saldo = rst("total") - rst("pago")
             If rst("id_doc") = "0007" Then
                in_saldo = in_saldo * -1
             End If
             
            Fila = rst("id_venta") & vbTab & "[" & Format(rst("fecha_emision"), "dd-mm-YYYY") & " - " & Format(rst("fecha_vencimiento"), "dd-mm-YYYY") & "]" & vbTab & rst("comprobante") & vbTab & rst("descripcion") & vbTab & Format(rst("total"), "#,##0.00") & vbTab & Format(in_saldo, "#,##0.00") & vbTab & estado
            Grilla.AddItem Fila
            in_total_pago = in_total_pago + in_saldo
            With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 6 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
            End With
      
            
            
            If rst("seleccion") = "si" Then
                tTotal = tTotal + in_saldo
                For k = 1 To 5
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
                Next k
            End If
            
            Fila = ""
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "SALDO TOTAL :::>" & vbTab & Format(in_total_pago, "#,##0.00")
      Grilla.AddItem Fila
      
            Grilla.col = 5
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      
            Me.TxtSaldo.Text = Format(tTotal, "###0.00")
            Me.TxtSaldo.Locked = True
            Me.TxtMontoPago.Text = Format(tTotal, "###0.00")
  Exit Sub
salir:    MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub


Private Sub Save()
Dim id_documento As String
Dim i As Integer
Dim Monto As Double
Dim codigo As String
Dim t_cambio As Single
Dim anul As String * 2
Dim Contado As String
Dim Adelantado As String
Dim Num_registros As Integer
Dim monto_letras As String
Dim nuevo_numero As String, Documento As String
Monto = Val(Me.TxtMontoPago.Text)
 
     If Monto < 1 Then
        MsgBox "Ingrese un Monto Mayor que cero", vbInformation, "Mensaje para el Usuario"
        Exit Sub
     End If
    If Me.DtcCuentas.BoundText = "" Then
        MsgBox "Indique una Cuenta Destino", vbInformation, "Mensaje para el Usuario"
        Me.DtcCuentas.SetFocus
        Exit Sub
    End If
       strCadena = "SELECT * FROM movimiento_venta WHERE id_cliente='" & Trim(Me.TxtRuc.Text) & "' AND seleccion='si' AND saldo>0 AND anulado='no' AND ruc='" & KEY_RUC & "' ORDER BY saldo DESC"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
       rst.MoveFirst
         
         'monto_letras = UCase(EnLetras(Monto))
         Documento = Trim(Me.DtcTipoDoc.Text) + ":" + Trim(Me.TxtSerie.Text) + "-" + Trim(Me.TxtNumeroDoc.Text)
         'strCadena = "INSERT INTO mis_cuentas_det(id_alm,id_doc,serie,numero,documento,id_cuenta,fecha,fecha_sys,id_persona,glosa,monto,montoreal,tc,monto_letras,operacion,id_movimiento,ccostos," & _
         "dni_save,ruc) VALUES ('" & Trim(Me.DtcAlmacen.BoundText) & "' ,'" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Me.txtSerie.Text & "','" & Me.TxtNumeroDoc.Text & "','" & Documento & "','" & Me.DtcCuentas.BoundText & "','" & KEY_FECHA & "'," & _
         "'" & Format(Now(), "YYYY-mm-dd") & "','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.txtObservacion.Text) & "','" & Monto & "','" & Monto & "','" & Me.TxtTc.Text & "','" & monto_letras & "'," & _
         "'" & Trim(Me.TxtOperacion.Text) & "','" & Me.tXTIdVenta.Text & "','16','" & KEY_USUARIO & "','" & KEY_RUC & "')"
         
         
         
         strCadena = "SELECT * FROM movimiento_venta WHERE id_cliente='" & Trim(Me.TxtRuc.Text) & "' AND seleccion='si' AND saldo>0 AND anulado='no' AND ruc='" & KEY_RUC & "' ORDER BY saldo DESC"
         Call ConfiguraRst(strCadena)
         If rst.RecordCount > 0 Then
            rst.MoveFirst
            
            For i = 0 To rst.RecordCount - 1
                strCadena = "SELECT * FROM movimiento_venta_cuotas WHERE id_venta='" & rst("id_venta") & "' AND ruc='" & KEY_RUC & "' AND saldo>0 ORDER BY id DESC"
                Call ConfiguraRstT(strCadena)
                If rstT.RecordCount > 0 Then
                    rstT.MoveFirst
                    For k = 0 To rstT.RecordCount - 1
                        If Monto >= rstT("saldo") And Monto > 0 Then
                            strCadena = "UPDATE movimiento_venta_cuotas SET saldo='0' WHERE id='" & rstT("id") & "' AND ruc='" & KEY_RUC & "'"
                            CnBd.Execute (strCadena)
                            
                            Monto = Monto - rstT("saldo")
                            strCadena = "UPDATE movimiento_venta SET saldo='" & rst("saldo") - rstT("saldo") & "' WHERE id_venta='" & rst("id_venta") & "' AND ruc='" & KEY_RUC & "'"
                            CnBd.Execute (strCadena)
                            
                        Else
                            If Monto > 0 Then
                            strCadena = "UPDATE movimiento_venta_cuotas SET saldo='" & Val(rstT("saldo") - Monto) & "' WHERE id='" & rstT("id") & "' AND ruc='" & KEY_RUC & "'"
                            CnBd.Execute (strCadena)
                            
                            
                            strCadena = "UPDATE movimiento_venta SET saldo='" & Val(rst("saldo") - Monto) & "' WHERE id_venta='" & rst("id_venta") & "' AND ruc='" & KEY_RUC & "'"
                            CnBd.Execute (strCadena)
                            
                            Monto = 0
                            End If
                        End If
                        
                        rstT.MoveNext
                    Next k
                    
                End If
                rst.MoveNext
            Next i
         End If
         
         
         
         
        End If
        nuevo_numero = formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6)
        strCadena = "UPDATE  almacen_comprobante SET numero='" & Trim(nuevo_numero) & "'  WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        
        Me.cmdPrinter.Enabled = True
        Me.cmdSave.Enabled = False

        Call FrmCuentasxCobrar.LlenarComprobantes(FrmCuentasxCobrar.HfComprobantes)
        Exit Sub
    

End Sub

Private Sub put_verificacion(ByVal in_venta As String)
in_observacion = "[" + KEY_FECHA + Space(2) + str(Time) + "]  " + Mid(KEY_VENDEDOR, 1, 30) + " : " + UCase(Me.TxtObservacion.Text)

in_validador = True
in_pendiente = "no"


strCadena = "UPDATE movimiento_venta SET orden_compra='" & Me.DtcFlujo.BoundText & "',id_orden_salida='" & Me.DtcCuentas.BoundText & "',operacion='" & Trim(Me.txtOperacion.Text) & "',pendiente='" & in_pendiente & "',observacion='" & in_observacion & "' WHERE id_venta='" & in_venta & "'"
CnBd.Execute (strCadena)



End Sub

Private Sub Save_v2()
Dim id_documento As String
Dim i As Integer
Dim Monto As Double
Dim codigo As String
Dim t_cambio As Single
Dim anul As String * 2
Dim Contado As String
Dim Adelantado As String
Dim Num_registros As Integer
Dim monto_letras As String
Dim nuevo_numero As String, Documento As String
Monto = Val(Me.TxtMontoPago.Text)
 
     If Monto <= 0 Then
        MsgBox "Ingrese un Monto Mayor que cero", vbInformation, "Mensaje para el Usuario"
        Exit Sub
     End If
     
    If Me.DtcCuentas.BoundText = "" Then
        MsgBox "Indique una Cuenta Destino", vbInformation, "Mensaje para el Usuario"
        Me.DtcCuentas.SetFocus
        Exit Sub
    End If
       
       

        
    Call generar_recibo ' graba el recibo
    '----
    If Val(Me.tXTIdVenta.Text) < 1 Then
        Me.cmdPrinter.Enabled = True
        Me.cmdSave.Enabled = True
        Exit Sub
    End If
                
     strCadena = "SELECT id_venta,id_doc,id_tipo,tc,comprobante,id_moneda,nombre_completo,descripcion,grupo_empresarial,cta_cobrar,(total-function_pago_factura(id_venta,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc)) as saldo " & _
    " FROM view_listado_comprobante_vargas WHERE seleccion='si' and dni_use='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' "
    Call ConfiguraRst(strCadena)
         If rst.RecordCount > 0 Then
            rst.MoveFirst
             in_tipo = rst("id_tipo")
             
             If rst("grupo_empresarial") = "si" Then
                Call put_empresa_vinculada(rst("grupo_empresarial"), rst("cta_cobrar"), rst("id_venta"))
             End If
             
                        If rst("id_moneda") <> Me.DtcMoneda.BoundText Then
                           If rst("id_moneda") = "00002" Then
                              Monto = Monto / Val(Me.TxtTc.Text)
                           Else
                               Monto = Monto * Val(Me.TxtTc.Text)
                           End If
                        End If
             
            For i = 0 To rst.RecordCount - 1
                       
                
                        If Monto >= rst("saldo") And Monto > 0 Then
                           Monto = Monto - rst("saldo")
                                                      
                           
                           in_glosa_item = "COBRO :" & rst("comprobante")
                           
                           If rst("id_moneda") <> Me.DtcMoneda.BoundText Then
                                If Me.DtcMoneda.BoundText = "00001" Then
                                   in_monto_deposito = rst("saldo") * Val(Me.TxtTc.Text)
                                Else
                                    in_monto_deposito = rst("saldo") / Val(Me.TxtTc.Text)
                                End If
                            Else
                                in_monto_deposito = rst("saldo")
                            End If
                            
                           If get_forma_pago_detalle(Me.DtcForma_pago_detalle.BoundText) = "20" Then  ' RETENCION
                                Call procesar_transaccion_retencion(rst("id_venta"), KEY_ALM, Format(Me.DtpEmision.Value, "YYYY-mm-dd"), Trim(Me.TxtRuc.Text), Trim(Me.txtserie_retencion.Text), Trim(Me.txtNumero_retencion.Text), rst("saldo"), Val(Me.TxtTc.Text), Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                                in_mis_cuentas_det = 0
                           Else
                                
                                
                                
                                
                                If KEY_ASIENTO_GLOBAL_CTA_PAGAR = "no" Then
                                    If rst("id_doc") = "0007" Then
                                        in_mis_cuentas_det = procesar_transaccion(KEY_ALM, Me.DtcCuentas.BoundText, Format(Me.DtpEmision.Value, "YYYY-mm-dd"), "00002", Trim(Me.TxtRuc.Text), get_persona(Trim(Me.TxtRuc.Text)), in_glosa, in_monto_deposito, "0", rst("id_venta"), "0", rst("comprobante"), Val(Me.TxtTc.Text), Trim(Me.txtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                                    Else
                                        in_mis_cuentas_det = procesar_transaccion(KEY_ALM, Me.DtcCuentas.BoundText, Format(Me.DtpEmision.Value, "YYYY-mm-dd"), "00001", Trim(Me.TxtRuc.Text), get_persona(Trim(Me.TxtRuc.Text)), in_glosa, in_monto_deposito, "0", rst("id_venta"), "0", rst("comprobante"), Val(Me.TxtTc.Text), Trim(Me.txtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                                    End If
                                End If
                                
                           End If
                           
                           
                           If get_forma_pago_detalle(Me.DtcForma_pago_detalle.BoundText) = "14" Then   ' PAGAR CON MEMORANDUM
                                    
                                    
                                    Call put_realizar_pago(Val(Me.txtidMemorandum.Text), rst("id_venta"), in_monto_deposito, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det), "03")
                                    Call put_realizar_pago(Val(Me.txtId_recibo.Text), rst("id_venta"), in_monto_deposito, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det))
                                    Call put_realizar_pago(rst("id_venta"), Val(Me.txtId_recibo.Text), in_monto_deposito, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det))
                            Else
                                    
                                    
                            If get_forma_pago_detalle(Me.DtcForma_pago_detalle.BoundText) = "13" Then   ' pagar con nota de credito
                                in_id_nota = get_id_nota(Trim(Me.txtSerie_nota.Text), Trim(Me.txtNumero_nota.Text))
                                Call put_realizar_pago(rst("id_venta"), in_id_nota, in_monto_deposito, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det))
                                Call put_realizar_pago(in_id_nota, rst("id_venta"), in_monto_deposito, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det))
                                Call put_realizar_pago(Val(Me.txtId_recibo.Text), Val(Me.txtId_recibo.Text), in_monto_deposito, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det))
                                
                             Else
                                    Call put_realizar_pago(Val(Me.txtId_recibo.Text), rst("id_venta"), in_monto_deposito, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det))
                                    Call put_realizar_pago(rst("id_venta"), Val(Me.txtId_recibo.Text), in_monto_deposito, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det))
                            End If
                           End If
                           
                           
                           
                                                      
                           
                         

                           strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & Val(Me.txtId_recibo.Text) & "','00','" & rst("comprobante") & "','-','1','" & rst("saldo") & "','0','" & rst("saldo") & "','" & KEY_RUC & "')"
                           CnBd.Execute (strCadena)
                               
                        Else
                            If Monto > 0 Then
                            
                            If rst("id_moneda") <> Me.DtcMoneda.BoundText Then
                                If Me.DtcMoneda.BoundText = "00001" Then
                                   in_monto_deposito = Monto * Val(Me.TxtTc.Text)
                                Else
                                    in_monto_deposito = Monto / Val(Me.TxtTc.Text)
                                End If
                            Else
                                in_monto_deposito = Monto
                            End If
                            
                            
                            in_glosa = "COBRO :" & rst("comprobante") & "[" & Trim(Me.TxtObservacion.Text) & "]"
                            
                            If get_forma_pago_detalle(Me.DtcForma_pago_detalle.BoundText) = "20" Then  ' RETENCION
                                Call procesar_transaccion_retencion(rst("id_venta"), KEY_ALM, Format(Me.DtpEmision.Value, "YYYY-mm-dd"), Trim(Me.TxtRuc.Text), Trim(Me.txtserie_retencion.Text), Trim(Me.txtNumero_retencion.Text), Val(Me.TxtMontoPago.Text), Val(Me.TxtTc.Text), Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                                in_mis_cuentas_det = 0
                            Else
                                If KEY_ASIENTO_GLOBAL_CTA_PAGAR = "no" Then
                                    in_mis_cuentas_det = procesar_transaccion(KEY_ALM, Me.DtcCuentas.BoundText, Format(Me.DtpEmision.Value, "YYYY-mm-dd"), "00001", Trim(Me.TxtRuc.Text), get_persona(Trim(Me.TxtRuc.Text)), in_glosa, in_monto_deposito, "0", rst("id_venta"), "0", rst("comprobante"), Val(Me.TxtTc.Text), Trim(Me.txtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                                End If
                                
                                
                            End If
                            
                            
                           
                            
                            If get_forma_pago_detalle(Me.DtcForma_pago_detalle.BoundText) = "13" Then   ' pagar con nota de credito
                                'Call put_actualizar_saldo_nota(id_venta, Monto, Trim(Me.txtserie_nota.Text), Trim(Me.txtNumero_nota.Text), Trim(Me.txtruc.Text))
                                in_id_nota = get_id_nota(Trim(Me.txtSerie_nota.Text), Trim(Me.txtNumero_nota.Text))
                                Call put_realizar_pago(id_venta, in_id_nota, Monto, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det))
                                Call put_realizar_pago(in_id_nota, id_venta, Monto, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det))
                            Else
                                
                                If get_forma_pago_detalle(Me.DtcForma_pago_detalle.BoundText) = "14" Then   ' PAGAR CON MEMORANDUM
                                    Call put_realizar_pago(Val(Me.txtidMemorandum.Text), rst("id_venta"), in_monto_deposito, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det), "03")
                                    Call put_realizar_pago(Val(Me.txtId_recibo.Text), rst("id_venta"), in_monto_deposito, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det))
                                    Call put_realizar_pago(rst("id_venta"), Val(Me.txtId_recibo.Text), in_monto_deposito, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det))
                                Else
                                    Call put_realizar_pago(Val(Me.txtId_recibo.Text), rst("id_venta"), in_monto_deposito, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det))
                                    Call put_realizar_pago(rst("id_venta"), Val(Me.txtId_recibo.Text), in_monto_deposito, Me.DtcTipoDoc.BoundText, Val(Me.TxtTc.Text), Val(in_mis_cuentas_det))
                                End If
                            End If
                            in_glosa_item = "COBRO :" & rst("comprobante")
                            strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & Val(Me.txtId_recibo.Text) & "','00','" & rst("comprobante") & "','-','1','" & Monto & "','0','" & Monto & "','" & KEY_RUC & "')"
                            CnBd.Execute (strCadena)
                            Monto = 0
                            End If
                        End If
                        
                        Call put_verificacion(rst("id_venta"))
                        
                        rst.MoveNext
                Next i
                
                
                
                If KEY_ASIENTO_GLOBAL_CTA_PAGAR = "si" Then
                 
                   If Val(Me.txtidMemorandum.Text) > 0 Then
                        strCadena = "call CON_InsertaAsiento_Memorandum('" & Val(Me.txtidMemorandum.Text) & "')"
                        CnBd.Execute (strCadena)
                   Else
                        
                       If get_forma_pago_detalle(Me.DtcForma_pago_detalle.BoundText) <> "13" Then
                        strCadena = "call CON_InsertaAsiento_CobroGlobal('" & Val(Me.txtId_recibo.Text) & "')"
                        CnBd.Execute (strCadena)
                       End If
                        
                        
                   End If
                   
                   If Me.DtcMoneda.BoundText = "00002" Then
                      strCadena = "call CON_AjusteTC_Global('" & Val(Me.txtId_recibo.Text) & "')"
                      CnBd.Execute (strCadena)
                    End If
                          
                   Call procesar_transaccion_venta(Me.DtcForma_pago_detalle.BoundText, KEY_ALM, Me.DtcCuentas.BoundText, Format(DtpEmision.Value, "YYYY-mm-dd"), "00001", Trim(Me.TxtRuc.Text), Trim(Me.TxtCliente.Text), Trim(Me.TxtObservacion.Text), Val(Me.TxtMontoPago.Text), "0", Val(txtId_recibo.Text), "0", in_recibo, Val(Me.TxtTc.Text), Trim(Me.txtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, KEY_USUARIO, Me.DtcTipoDoc.BoundText, KEY_RUC)
                End If
                 
         End If
         
         
         
         
        
        Me.cmdPrinter.Enabled = True
        Me.cmdSave.Enabled = False
        
       ' FrmReporteRegistroVentas.HfdPersona.TextMatrix(FrmReporteRegistroVentas.HfdPersona.Row, 7) = Format(Format(FrmReporteRegistroVentas.HfdPersona.TextMatrix(FrmReporteRegistroVentas.HfdPersona.Row, 7), "###0.00") - Format(Me.TxtMontoPago.Text, "###0.00"), "#,##0.00")
        Call actualizar_credito(Trim(TxtRuc.Text), Val(Me.TxtMontoPago.Text))
        If KEY_PROYECTO = "si" Then
            Call put_proyecto_estado(Me.txtId_recibo.Text)
        End If
        
        Me.txtidMemorandum.Text = 0
        Exit Sub
    

End Sub




Private Sub put_proyecto_estado(ByVal in_recibo As String)

strCadena = "SELECT DISTINCT v.id_proyecto FROM mis_cuentas_det_detalle d,movimiento_venta v WHERE d.id_movimiento=v.id_venta and d.id_detalle='" & Val(in_recibo) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
     strCadena = "SELECT SUM(total-function_pago_factura(id_venta,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc)) as saldo " & _
    " FROM view_listado_comprobante_vitekey WHERE anulado='no' and id_proyecto='" & rst("id_proyecto") & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRstK(strCadena)
    If rstK("saldo") = 0 Then
        strCadena = "UPDATE mis_proyectos SET finalizado='si' WHERE id_proyecto='" & rst("id_proyecto") & "'"
        CnBd.Execute (strCadena)
    End If
    rst.MoveNext
    Next i
End If




End Sub

Private Sub put_empresa_vinculada(ByVal in_grupo As String, ByVal in_cta_cobrar As String, ByVal in_venta As String)

If in_grupo = "si" And Mid(in_cta_cobrar, 1, 2) = "12" Then
    strCadena = "UPDATE movimiento_venta SET cta_cobrar='1310301',cta_ingreso='7010102' WHERE id_venta='" & Val(in_venta) & "'"
    CnBd.Execute (strCadena)
End If

End Sub


Private Sub insert_pago_letra(ByVal in_venta As String, ByVal in_monto As Single)
strCadena = "SELECT * FROM movimiento_venta_cuotas WHERE id_venta='" & in_venta & "' and ruc='" & KEY_RUC & "' ORDER BY seleccion ASC"
         Call ConfiguraRstK(strCadena)
         If rstK.RecordCount > 0 Then
            rstK.MoveFirst
            
            For i = 0 To rstK.RecordCount - 1
                
                        If in_monto >= rstK("saldo") And in_monto > 0 Then
                           in_monto = in_monto - rstK("saldo")
                           strCadena = "UPDATE movimiento_venta_cuotas SET saldo='0' WHERE id='" & rstK("id") & "'"
                           CnBd.Execute (strCadena)
                        Else
                            If in_monto > 0 Then
                                strCadena = "UPDATE movimiento_venta_cuotas SET saldo=saldo - '" & in_monto & "' WHERE id='" & rstK("id") & "'"
                                CnBd.Execute (strCadena)
                                in_monto = 0
                            End If
                        End If
                        rstK.MoveNext
            Next i
         End If

End Sub

Private Sub generar_recibo()

                    KEY_VENCIMIENTO = KEY_FECHA
                    id_tipo_factura = "00001"
                    igv = "si"
                    dfac = "no"
                    
                    strCadena = "SELECT count(*) FROM  movimiento_venta WHERE id_doc='0054' and serie='" & Trim(Me.TxtSerie.Text) & "' and numero='" & Trim(Me.TxtNumeroDoc.Text) & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstZ(strCadena)
                    If rstZ(0) > 0 Then
                        MsgBox "Recibo ya generado verifique su correlativo", vbInformation, KEY_EMPRESA
                        Call Resalta(Me.TxtNumeroDoc)
                        Exit Sub
                    End If
                    
                    
                    horario = Format(Time, "hh:mm")
                    If horario >= "07:00" And horario <= "13:00" Then
                        turno = "M"
                    Else
                        turno = "T"
                    End If
                    
                    
                    Documento = Trim(Me.DtcTipoDoc.Text) & ":" & Trim(Me.TxtSerie.Text) & "-" & Trim(Me.TxtNumeroDoc.Text)
                    strCadena = "P_insert_venta('" & Me.DtcTipoDoc.BoundText & "','" & KEY_ALM & "','01','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
                    "'" & Trim(Me.TxtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Me.TxtRuc.Text & "','" & Me.TxtCliente.Text & "','0','0','0','" & Val(Me.TxtMontoPago.Text) & "','0'," & _
                    "'" & Val(Me.TxtMontoPago.Text) & "','0','" & Format(Me.DtpEmision.Value, "YYYY-mm-dd") & "','" & KEY_VENCIMIENTO & "','" & id_tipo_factura & "','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & Val(Me.TxtTc.Text) & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstP(strCadena)
                    id_venta = rstP(0)
                    Me.txtId_recibo.Text = id_venta
                    If Me.frmtarjeta.Visible = True Then
                        in_tarjeta = Me.DtcTarjeta.BoundText
                    Else
                        in_tarjeta = "00"
                    End If
                    
                    'strCadena = "INSERT INTO movimiento_venta_monto(id_venta,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,banco,cheque,cuenta_contable,forma_pago_contable,flujo_caja,id_cuenta_origen,cta_redondeo,cta_anticipo,monto_redondeo,monto_anticipo,ruc)VALUES " & _
                    '"('" & id_venta & "','01','" & get_forma_pago_anterior(Me.DtcMoneda.BoundText) & "','" & Val(Me.TxtMontoPago.Text) & "','" & Val(Me.TxtMontoPago.Text) * -1 & "','00','-','" & Trim(Me.txtOperacion.Text) & "','-','" & Me.DtcCheque.BoundText & "','" & get_cuenta_contable_cuenta(Me.DtcCuentas.BoundText) & "','" & DtcFormaPago.BoundText & "','" & Me.DtcFlujo.BoundText & "','" & Me.DtcCuentas.BoundText & "','" & Trim(Me.txtCuenta_redondeo.Text) & "','" & Trim(Me.txtCuenta_anticipo.Text) & "','" & Val(Me.txtMontoRedondeo.Text) & "','" & Val(Me.txtMontoAnticipo.Text) & "','" & KEY_RUC & "')"
                    'CnBd.Execute (strCadena)
                    
                    
                    strCadena = "INSERT INTO movimiento_venta_monto(id_venta,id_forma_pago,forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,id_recibo,banco,cheque,cuenta_contable,forma_pago_contable,flujo_caja,ruc) VALUES " & _
                    "('" & Val(Me.txtId_recibo.Text) & "','" & Me.DtcForma_pago_detalle.BoundText & "','01','" & Val(Me.TxtMontoPago.Text) & "','" & Val(Me.TxtMontoPago.Text) & "','" & in_tarjeta & "','" & Trim(Me.TxtNumeroTargeta.Text) & "','" & Trim(Me.txtOperacion.Text) & "','0','" & Me.txtBanco.Text & "','" & Trim(Me.txtCheque.Text) & "','" & get_cuenta_contable_cuenta(Me.DtcCuentas.BoundText) & "','" & Me.DtcFormaPago.BoundText & "','" & Me.DtcFlujo.BoundText & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    If Me.chkElegir.Value = 1 And Val(Me.TxtMontoReal.Text) > 0 Then
                        strCadena = "UPDATE movimiento_venta SET monto_vehiculo='" & Val(Me.txtmontovehiculo.Text) & "', id_comprobante='" & Val(Me.tXTIdVenta.Text) & "',observacion='" & Trim(Me.TxtObservacion.Text) & "' WHERE id_venta='" & id_venta & "'"
                    Else
                        strCadena = "UPDATE movimiento_venta SET id_ventanilla='" & Me.DtcVentanilla.BoundText & "',id_comprobante='" & Val(Me.tXTIdVenta.Text) & "',observacion='" & Trim(Me.TxtObservacion.Text) & "',id_referencia='" & id_venta & "' WHERE id_venta='" & id_venta & "'"
                    End If
                    CnBd.Execute (strCadena)
                    
                    
                    strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(Me.TxtNumeroDoc.Text + 1), "000000") & "' WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                 '   Call llenar_montos(id_venta)
End Sub
Private Sub llenar_montos(ByVal id_venta As Double)


       strCadena = "INSERT INTO movimiento_venta_monto(id_venta,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,ruc)VALUES('" & id_venta & "','01','" & get_id_registro_forma_pago("01", "01") & "','" & Val(Me.TxtMontoPago.Text) & "','" & Val(Me.TxtMontoPago.Text) & "','00','--','" & Trim(Me.txtOperacion.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
     





End Sub
Private Sub SaveDetalleDocumentoVentaRecibo(ByVal idVenta As Double)


End Sub

Private Sub HfBancos_Click()
Me.txtBanco.Text = Me.HfBancos.TextMatrix(Me.HfBancos.Row, 1)
Call Resalta(Me.txtCheque)
    Exit Sub
End Sub





Private Sub HfProyectos_DblClick()
Dim in_status As String

If Me.HfProyectos.Rows > 0 Then
   If Val(Me.HfProyectos.TextMatrix(Me.HfProyectos.Row, 0)) > 0 Then
        
      If Trim(Me.HfProyectos.TextMatrix(Me.HfProyectos.Row, 3)) = Chr(168) Then
         Me.HfProyectos.TextMatrix(Me.HfProyectos.Row, 3) = Chr(254)
         in_status = "si"
      Else
         Me.HfProyectos.TextMatrix(Me.HfProyectos.Row, 3) = Chr(168)
         in_status = "no"
      End If
        
      
      
      strCadena = "CALL cursor_put_proyecto('" & Val(Me.HfProyectos.TextMatrix(Me.HfProyectos.Row, 0)) & "','" & in_status & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
      CnBd.Execute (strCadena)
      
      Call LlenarVinculados(Me.HfVinculados, Trim(Me.TxtRuc.Text))
      Call llenar_linea(Me.HfServicios, Trim(Me.TxtRuc.Text))
   End If
End If
End Sub

Private Sub HfVinculados_Click()
If Me.HfVinculados.Rows > 0 Then
    If Val(Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 0)) > 0 Then
       Call put_seleccionar
    End If
End If
End Sub
Private Sub put_seleccionar()
Dim cPersona  As String
Dim in_seleccion As Boolean
    strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & Val(Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 0)) & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 And Val(Me.tXTIdVenta.Text) <> Val(Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 0)) Then
        If rst("saldo") < 0 Then
           MsgBox "DOCUMENTO DEREFERENCIA NO PUEDE SELECCIONARSE", vbInformation, KEY_VENDEDOR
           Exit Sub
        End If
        
        If rst("seleccion") = "no" Then
            in_seleccion = False
            strCadena = "UPDATE movimiento_venta SET seleccion='si',dni_use='" & KEY_USUARIO & "' WHERE id_venta='" & rst("id_venta") & "' AND ruc='" & KEY_RUC & "'"
        Else
            in_seleccion = True
            strCadena = "UPDATE movimiento_venta SET seleccion='no',dni_use='0' WHERE id_venta='" & rst("id_venta") & "'AND ruc='" & KEY_RUC & "'"
        End If
        
        cPersona = rst("id_cliente")
        CnBd.Execute (strCadena)
        
        
        If in_seleccion = True Then
            in_color = &HFFFFFF
            Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 6) = Chr(168)
        Else
            in_color = &H8080FF
            Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 6) = Chr(254)
        End If
     
            For k = 1 To 5
                HfVinculados.col = k
                HfVinculados.Row = Me.HfVinculados.Row
                HfVinculados.CellBackColor = in_color
            Next k
           
        
        
        
        Me.TxtObservacion.Text = "COBRO:" + comprobante
        Call Resalta(Me.TxtMontoPago)
   End If
       
         
               
         
        
        
        
        'Call LlenarVinculados(Me.HfVinculados, cPersona)
        'Call llenar_recibos(Me.HfRecibos, Trim(Me.txtruc.Text))
       
       
   Call get_monto_pagar
            
            
            
       

End Sub
Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Error
  Select Case Button.key
    Case KEY_NEW
      Call nuevo
    Case KEY_DELETE
        'If MsgBox("Esta Seguro de Eliminar Este Comprobante", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        'strCadena = "DELETE FROM movimiento_caja  WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "')"
        'CnBd.Execute (strCadena)
        MsgBox "Para eliminar ingrese al modulo de ventas", vbInformation
        
       'End If
    Case KEY_ANULAR
    If MsgBox("Esta Seguro de Anular Este Comprobante", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        'strCadena = "UPDATE movimiento_caja SET anulado='si' WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "')"
        'CnBd.Execute (strCadena)
         'doc = Me.DtcTipoDoc.Text + ":" + Me.TxtSerie.Text + "-" + Me.TxtNumeroDoc.Text
         's'trCadena = "DELETE mis_cuentas_det WHERE documento='" & Trim(doc) & "'"
         'CnBd.Execute (strCadena)
        'Me.lblAnulado.Visible = True
        'Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
        MsgBox "Para anular ingrese al modulo de ventas", vbInformation
       End If
    Case KEY_EXIT
        Unload Me
'Error:
 ' MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  
  Exit Sub
End Select
End Sub

Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
     
      
    Case KEY_PRINT
       

End Select
  Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR

End Sub

Private Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub




Private Sub HfVinculados_SelChange()
'If Val(Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 0)) > 0 Then
    'Call LlenarLetras(Me.HfLetras, Val(Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 0)))
    'Call LlenarCuotas(Me.HfCuotas, Val(Me.HfVinculados.TextMatrix(Me.HfVinculados.Row, 0)))
'End If
End Sub

Private Sub txtbuscarbanco_Change()
 strCadena = "SELECT * FROM entidadfinanciera  WHERE descripcion LIKE '%" & Trim(Me.txtbuscarbanco.Text) & "%'  ORDER BY descripcion"
     Call llenar_bancos(Me.HfBancos)
End Sub

Private Sub txtBuscarForma_Change()
If KEY_CONTABILIDAD = "si" Then
       strCadena = "SELECT id_registro as Codigo, cuenta as Descripcion FROM view_forma_pago_conta  WHERE  cuenta LIKE '%" & Trim(Me.txtBuscarForma.Text) & "%' and  ruc='" & KEY_RUC & "'"
    Else
       strCadena = "SELECT id_registro as Codigo, CONCAT(descripcion,'-',observacion) as Descripcion FROM forma_pago_detalle  WHERE observacion LIKE '%" & Trim(Me.txtBuscarForma.Text) & "%' and  id_moneda='" & Me.DtcMoneda.BoundText & "' and  id='01' AND ruc='" & KEY_RUC & "' AND estado='si' ORDER BY id_detalle"
    End If
    
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcForma_pago_detalle)
    End Sub

Private Sub txtCheque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtOperacion.Text = Trim(Me.txtOperacion.Text) & Space(1) & "CHEQUE: " & Trim(Me.txtCheque.Text)
End If
End Sub

Private Sub TxtMontoPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtMontoPago.Text = Format(Val(Me.TxtMontoPago.Text), "###0.00")
    If Me.chkElegir.Value = 1 Then
        Me.TxtSaldo.Text = Val(Me.txtmontovehiculo.Text) - Val(Me.TxtMontoPago.Text)
    End If
    Me.DtcCuentas.SetFocus
End If
End Sub

Private Sub txtmontovehiculo_LostFocus()
If Val(Me.txtmontovehiculo.Text) < Val(Me.TxtMontoReal.Text) Then
    MsgBox "Monto Ingresado Es menor al Precio Original del Vehiculo", vbInformation
    Me.txtmontovehiculo.Text = Val(Me.TxtMontoReal.Text)
    Call Resalta(Me.txtmontovehiculo)
End If
End Sub

Private Sub txtNumero_nota_KeyPress(KeyAscii As Integer)
Dim in_tipo_comprobante As String
If KeyAscii = 13 Then
    in_tipo_comprobante = get_forma_pago_detalle(Me.DtcForma_pago_detalle.BoundText)
    
    If in_tipo_comprobante = "13" Then
        Call load_nota
    End If
    
    If in_tipo_comprobante = "14" Then
        Call load_memo
    End If
    
    
End If
    
End Sub
Private Sub load_nota()
    Me.txtNumero_nota.Text = FormatosCeros(txtNumero_nota, 6)
    Me.TxtMontoPago.Text = get_saldo_nota_credito(Trim(Me.txtSerie_nota.Text), Trim(Me.txtNumero_nota.Text), Trim(Me.TxtRuc.Text))
End Sub
Private Sub load_memo()
    Me.txtNumero_nota.Text = FormatosCeros(txtNumero_nota, 6)
    Me.TxtMontoPago.Text = get_monto_memorandum(Me.DtcTipoMemorandum.BoundText, Trim(Me.txtSerie_nota.Text), Trim(Me.txtNumero_nota.Text), Trim(Me.TxtRuc.Text))
End Sub
Private Function get_monto_memorandum(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_dni As String) As Single
strCadena = "SELECT * FROM memorandun WHERE id_doc='" & in_doc & "' and  in_dni_cliente='" & in_dni & "' and  serie='" & Trim(Me.txtSerie_nota.Text) & "' and numero='" & Trim(Me.txtNumero_nota.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    Me.txtidMemorandum.Text = rstL("id_memo")
    Me.lblReferencia_nota.Caption = rstL("asunto")
    Me.lblsaldo_nota.Caption = Format(rstL("monto"), "###0.00") & Space(2) & get_moneda(rstL("id_moneda"))
    get_monto_memorandum = Format(rstL("monto"), "###0.00")
    If rstL("id_moneda") = "00002" Then
       Me.DtcMoneda.BoundText = "00001"
       get_monto_memorandum = rstL("monto") * Val(Me.TxtTc.Text)
    End If
Else
    MsgBox "NOTA DE CREDITO NO REGISTRADA", vbInformation, KEY_VENDEDOR
End If

End Function

Private Sub TxtNumeroDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Call Resalta(Me.TxtRuc)
   
    End If


End Sub
Private Function get_saldo_nota_credito(ByVal in_serie As String, ByVal in_numero As String, ByVal in_dni As String) As Single
strCadena = "SELECT (total-function_pago_factura(id_venta,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc)) as saldo,fecha_emision,id_venta FROM movimiento_venta WHERE id_cliente='" & in_dni & "' and  id_doc='0007' and  serie='" & Trim(Me.txtSerie_nota.Text) & "' and numero='" & Trim(Me.txtNumero_nota.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    get_saldo_nota_credito = True
   
      Me.lblsaldo_nota.Caption = Format(rstL("saldo"), "###0.00")
      get_saldo_nota_credito = rstL("saldo")
      
  
   strCadena = "SELECT  get_referencia_nota('" & rstL("id_venta") & "','" & KEY_RUC & "')"
   Call ConfiguraRstL(strCadena)
   Me.lblReferencia_nota.Caption = rstL(0)
   
Else
    MsgBox "NOTA DE CREDITO NO REGISTRADA", vbInformation, KEY_VENDEDOR
End If

End Function



Private Sub TxtNumeroTargeta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtOperacionTarjeta)
End If
End Sub

Private Sub txtOperacionTarjeta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtOperacion.Text = Trim(Me.txtOperacion.Text) & Space(1) & "TAR:" & Trim(Me.TxtNumeroTargeta.Text) & "OP:" & Trim(Me.txtOperacionTarjeta.Text)
End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM persona where dni='" & Trim(Me.TxtRuc.Text) & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        Me.TxtCliente.Text = rstZ("nombre_completo")
        Me.TxtDireccion.Text = rstZ("direccion")
        Call Resalta(Me.TxtMontoPago)
    Else
       Procedencia = Selecionar
       FrmPersona.Show
       Exit Sub
    End If
End If
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie.Text = FormatosCeros(Me.TxtSerie.Text, 4)
    strCadena = "SELECT Alm_cod,doc_cod,serie FROM Det_alm_com WHERE (Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'" & _
        " AND  doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If Me.TxtNumeroDoc.Text = "" Then
            
            Me.TxtNumeroDoc.SetFocus
        Else
            Set rst = Nothing
             strCadena = "SELECT numero FROM Det_alm_com WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "') ORDER BY numero DESC"
             Call ConfiguraRst(strCadena)
             Me.TxtNumeroDoc.Text = rst(0)
             Call Resalta(Me.TxtNumeroDoc)
             Set rst = Nothing
        End If
    Else
        MsgBox "Serie no Asiganda a a dicho Almacen", vbInformation, KEY_EMPRESA
    End If
End If

End Sub




Private Sub txtSerie_nota_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtSerie_nota.Text = UCase(Me.txtSerie_nota.Text)
    Call Resalta(Me.txtNumero_nota)
End If
End Sub

Private Sub txtserie_retencion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Call Resalta(Me.txtNumero_retencion)
End If
End Sub

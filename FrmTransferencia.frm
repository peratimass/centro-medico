VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmTransferencias 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   17970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmtransporte 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   8400
      TabIndex        =   94
      Top             =   1095
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton cmdguia_cerrar 
         Height          =   255
         Left            =   7680
         Picture         =   "FrmTransferencia.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   120
         Width           =   255
      End
      Begin VB.TextBox TxtPlaca1 
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
         Left            =   1440
         MaxLength       =   80
         TabIndex        =   109
         Top             =   960
         Width           =   3300
      End
      Begin VB.TextBox TxtBuscarMarcaPlaca 
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
         Height          =   330
         Left            =   6240
         MaxLength       =   80
         TabIndex        =   108
         Top             =   600
         Width           =   1380
      End
      Begin VB.TextBox txtBuscarTransporte 
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
         Height          =   330
         Left            =   6240
         MaxLength       =   80
         TabIndex        =   107
         Top             =   200
         Width           =   1380
      End
      Begin VB.TextBox txtLicencia1 
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
         Left            =   1440
         MaxLength       =   80
         TabIndex        =   105
         Top             =   2520
         Width           =   3300
      End
      Begin VB.TextBox txtCertficado1 
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
         Left            =   1440
         MaxLength       =   80
         TabIndex        =   101
         Top             =   1680
         Width           =   3300
      End
      Begin VB.TextBox txtMtc1 
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
         Left            =   1440
         MaxLength       =   80
         TabIndex        =   99
         Top             =   1320
         Width           =   3300
      End
      Begin MSDataListLib.DataCombo DtcTransporte 
         Height          =   330
         Left            =   1440
         TabIndex        =   95
         Top             =   200
         Width           =   4695
         _ExtentX        =   8281
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
      Begin MSDataListLib.DataCombo DtcMarcaPlaca 
         Height          =   330
         Left            =   1440
         TabIndex        =   97
         Top             =   600
         Width           =   4695
         _ExtentX        =   8281
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
      Begin MSDataListLib.DataCombo dtcChofer 
         Height          =   330
         Left            =   1440
         TabIndex        =   103
         Top             =   2040
         Width           =   4695
         _ExtentX        =   8281
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
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "PLACA :"
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
         Left            =   720
         TabIndex        =   110
         Top             =   990
         Width           =   555
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "LICENCIA :"
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
         Left            =   480
         TabIndex        =   106
         Top             =   2565
         Width           =   780
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "CHOFER :"
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
         Left            =   555
         TabIndex        =   104
         Top             =   2100
         Width           =   705
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "CERTIFICADO :"
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
         Left            =   150
         TabIndex        =   102
         Top             =   1725
         Width           =   1110
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "MTC :"
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
         Left            =   840
         TabIndex        =   100
         Top             =   1350
         Width           =   420
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "MARCA-PLACA :"
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
         Left            =   90
         TabIndex        =   98
         Top             =   660
         Width           =   1170
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "TRANSPORTE :"
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
         TabIndex        =   96
         Top             =   200
         Width           =   1095
      End
   End
   Begin VB.Frame frmanulado 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   11400
      TabIndex        =   63
      Top             =   7560
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Label lblanulado 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ANULADO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   630
         Left            =   240
         TabIndex        =   64
         Top             =   120
         Visible         =   0   'False
         Width           =   2235
      End
   End
   Begin VB.TextBox txtdocumentoReferencia 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Left            =   15240
      MaxLength       =   80
      TabIndex        =   145
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox txtDireccionLlegada 
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
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   143
      Top             =   3630
      Width           =   5655
   End
   Begin VB.Frame frmimpresoras 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3015
      Left            =   2520
      TabIndex        =   59
      Top             =   5355
      Visible         =   0   'False
      Width           =   4740
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfImpresoras 
         Height          =   2055
         Left            =   45
         TabIndex        =   60
         Top             =   120
         Width           =   4335
         _ExtentX        =   7646
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
      Begin VitekeySoft.ChameleonBtn cmdcerrarimpresora 
         Height          =   180
         Left            =   4440
         TabIndex        =   61
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
         MICON           =   "FrmTransferencia.frx":2EA4
         PICN            =   "FrmTransferencia.frx":2EC0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdenviarimpresion 
         Height          =   705
         Left            =   45
         TabIndex        =   62
         Top             =   2235
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1244
         BTYPE           =   3
         TX              =   "  ENVIAR A IMPRESION                            "
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmTransferencia.frx":5D74
         PICN            =   "FrmTransferencia.frx":5D90
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.TextBox txtOtros 
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
      Height          =   350
      Left            =   5475
      MultiLine       =   -1  'True
      TabIndex        =   142
      Top             =   7970
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtContenedor 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Left            =   13560
      MaxLength       =   80
      TabIndex        =   140
      Top             =   8040
      Width           =   975
   End
   Begin VB.Frame frmUbigeo 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2325
      Left            =   7320
      TabIndex        =   135
      Top             =   480
      Visible         =   0   'False
      Width           =   8175
      Begin VB.TextBox txtBuscaUbigeo 
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
         Height          =   315
         Left            =   1440
         MaxLength       =   80
         TabIndex        =   136
         Top             =   120
         Width           =   2535
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfUbigeo 
         Height          =   1815
         Left            =   120
         TabIndex        =   137
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3201
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         GridColor       =   8421376
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
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
      Begin VB.Label lblDepartamento 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BUSCAR UBIGEO:"
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
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   138
         Top             =   120
         Width           =   1155
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   7800
         Picture         =   "FrmTransferencia.frx":8361
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.TextBox txtDireccionPartida 
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
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   133
      Top             =   2160
      Width           =   5655
   End
   Begin VB.Frame frmdireccion 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2175
      Left            =   7425
      TabIndex        =   65
      Top             =   3450
      Visible         =   0   'False
      Width           =   7815
      Begin VB.CommandButton cmdcerrardireccion 
         Height          =   255
         Left            =   7515
         Picture         =   "FrmTransferencia.frx":B205
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   120
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfdireccion 
         Height          =   1935
         Left            =   120
         TabIndex        =   67
         Top             =   120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   3413
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
   Begin VB.TextBox TxtBultos 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Left            =   10860
      MaxLength       =   80
      TabIndex        =   131
      Top             =   8070
      Width           =   1215
   End
   Begin VB.TextBox txt_hash 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   16920
      TabIndex        =   129
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_sunat_key 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   17040
      TabIndex        =   128
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chk_consultar 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "CONSULTAR"
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
      Left            =   6240
      TabIndex        =   127
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frameseries 
      BackColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   9240
      TabIndex        =   43
      Top             =   720
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox txtEstado 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   118
         Top             =   5040
         Width           =   735
      End
      Begin VitekeySoft.ChameleonBtn cmdCerrarpantalla 
         Height          =   300
         Left            =   5280
         TabIndex        =   48
         Top             =   5040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         BTYPE           =   5
         TX              =   "  CERRAR SERIES"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmTransferencia.frx":E0A9
         PICN            =   "FrmTransferencia.frx":E0C5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtchasis 
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
         Left            =   2760
         MaxLength       =   80
         TabIndex        =   45
         Top             =   240
         Width           =   2895
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfChasis 
         Height          =   2415
         Left            =   240
         TabIndex        =   44
         Top             =   600
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4260
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfSeries 
         Height          =   1695
         Left            =   240
         TabIndex        =   47
         Top             =   3240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2990
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
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "BUSCAR CHASIS:"
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
         Height          =   195
         Left            =   1320
         TabIndex        =   46
         Top             =   210
         Width           =   1095
      End
   End
   Begin VB.OptionButton Opt_transporte_privado 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "TRANSPORTE PRIVADO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   12480
      TabIndex        =   125
      Top             =   1320
      Width           =   1815
   End
   Begin VB.OptionButton Opt_transporte_publico 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "TRANSPORTE PUBLICO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   10560
      TabIndex        =   124
      Top             =   1320
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.TextBox txtIdUbigeoDestino 
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
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   122
      Top             =   3950
      Width           =   1335
   End
   Begin VB.TextBox txtUbigeoOrigen 
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
      Left            =   2925
      MaxLength       =   80
      TabIndex        =   121
      Top             =   2490
      Width           =   4305
   End
   Begin VB.TextBox txtIdUbigeoOrigen 
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
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   120
      Top             =   2490
      Width           =   1335
   End
   Begin VB.TextBox txtdiferida 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   117
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Txtmoneda 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   115
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtTc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   114
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtid_vinculado 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6480
      TabIndex        =   93
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtvalor_mercaderia 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Left            =   10860
      MaxLength       =   80
      TabIndex        =   92
      Top             =   7770
      Width           =   1215
   End
   Begin VB.CheckBox chk_valor_mercaderia 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "VALOR MERCADERIA."
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
      Height          =   250
      Left            =   8880
      TabIndex        =   91
      Top             =   7770
      Width           =   1815
   End
   Begin VB.TextBox txtcantidad_total 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Left            =   16515
      MaxLength       =   80
      TabIndex        =   90
      Top             =   6660
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtcertificado 
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
      Left            =   10560
      MaxLength       =   80
      TabIndex        =   88
      Top             =   2835
      Width           =   3300
   End
   Begin VB.TextBox txtpesototal 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Left            =   10860
      MaxLength       =   80
      TabIndex        =   83
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CheckBox chk_pesoglobal 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "PESO TOTAL GUIA:"
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
      Height          =   250
      Left            =   8880
      TabIndex        =   82
      Top             =   7485
      Width           =   1815
   End
   Begin VB.TextBox txt_idremitente 
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
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   81
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtremitente 
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
      Left            =   2925
      MaxLength       =   80
      TabIndex        =   80
      Top             =   1800
      Width           =   4305
   End
   Begin VB.TextBox txtplaca 
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
      Left            =   11925
      MaxLength       =   80
      TabIndex        =   78
      Top             =   2145
      Width           =   1935
   End
   Begin VitekeySoft.ChameleonBtn cmdverificar 
      Height          =   795
      Left            =   1320
      TabIndex        =   77
      Top             =   7560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1402
      BTYPE           =   5
      TX              =   "VERIFICAR"
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
      MICON           =   "FrmTransferencia.frx":110DA
      PICN            =   "FrmTransferencia.frx":110F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdprocesar 
      Height          =   795
      Left            =   120
      TabIndex        =   75
      Top             =   7560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1402
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
      MICON           =   "FrmTransferencia.frx":1372F
      PICN            =   "FrmTransferencia.frx":1374B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdimprimir 
      Height          =   795
      Left            =   2520
      TabIndex        =   76
      Top             =   7560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1402
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
      MICON           =   "FrmTransferencia.frx":16D93
      PICN            =   "FrmTransferencia.frx":16DAF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chk_manifiesto 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "MANIFIESTO"
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
      Left            =   240
      TabIndex        =   73
      Top             =   4275
      Width           =   1300
   End
   Begin VB.TextBox txtUbigeoDestino 
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
      Left            =   2920
      MaxLength       =   80
      TabIndex        =   71
      Top             =   3950
      Width           =   4305
   End
   Begin VB.TextBox txtid_direccion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   69
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chk_direccion 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7320
      TabIndex        =   68
      Top             =   3000
      Width           =   255
   End
   Begin VB.TextBox txt_dni_atencion 
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
      Left            =   4920
      MaxLength       =   80
      TabIndex        =   58
      Top             =   4275
      Width           =   735
   End
   Begin VB.TextBox txt_atencion 
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
      Left            =   5685
      MaxLength       =   80
      TabIndex        =   57
      Top             =   4275
      Width           =   1575
   End
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   825
      Left            =   16800
      TabIndex        =   51
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
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
      MICON           =   "FrmTransferencia.frx":19380
      PICN            =   "FrmTransferencia.frx":1939C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txt_tipo_transferencia 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   50
      Top             =   1800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtverificado 
      Height          =   285
      Left            =   7920
      TabIndex        =   49
      Text            =   "no"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TXTmTC 
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
      Left            =   10560
      MaxLength       =   80
      TabIndex        =   40
      Top             =   2490
      Width           =   3300
   End
   Begin VB.TextBox txtid_venta 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   39
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   550
      Left            =   1560
      TabIndex        =   35
      Top             =   800
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox TxtSeri_guia 
         Alignment       =   1  'Right Justify
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
         Left            =   3960
         MaxLength       =   80
         TabIndex        =   37
         Top             =   160
         Width           =   615
      End
      Begin VB.TextBox TxtNumero_guia 
         Alignment       =   1  'Right Justify
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
         Left            =   4605
         MaxLength       =   80
         TabIndex        =   36
         Top             =   160
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DtcComprobanteGuia 
         Height          =   315
         Left            =   1680
         TabIndex        =   38
         Top             =   165
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSDataListLib.DataCombo DtcTipoMovimiento 
         Height          =   315
         Left            =   120
         TabIndex        =   113
         Top             =   165
         Width           =   1455
         _ExtentX        =   2566
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
   End
   Begin VB.CheckBox ChkExtraer 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "EXTRAER"
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
      Left            =   480
      TabIndex        =   34
      Top             =   1035
      Width           =   1215
   End
   Begin VB.TextBox TxtId_transferencia 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7800
      TabIndex        =   33
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtObservacion 
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
      Height          =   405
      Left            =   13560
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   7560
      Width           =   2895
   End
   Begin VB.TextBox TxtUnidad 
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
      Left            =   9540
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   30
      Top             =   7150
      Width           =   1215
   End
   Begin VB.TextBox txtLicencia 
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
      Left            =   10560
      MaxLength       =   80
      TabIndex        =   29
      Top             =   3675
      Width           =   3255
   End
   Begin VB.TextBox TxtRucChofer 
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
      Left            =   10560
      MaxLength       =   80
      TabIndex        =   26
      Top             =   3300
      Width           =   1335
   End
   Begin VB.TextBox TxtMarcayPlaca 
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
      Left            =   10560
      MaxLength       =   80
      TabIndex        =   24
      Top             =   2145
      Width           =   1335
   End
   Begin VB.TextBox TxtRucTransporte 
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
      Left            =   10560
      MaxLength       =   80
      TabIndex        =   21
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox TxtPeso 
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
      Left            =   8400
      MaxLength       =   80
      TabIndex        =   19
      Top             =   7150
      Width           =   855
   End
   Begin VB.CommandButton CmdQuitar 
      BackColor       =   &H008080FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7140
      Width           =   375
   End
   Begin VB.CommandButton CmdAgregar 
      BackColor       =   &H008080FF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7140
      Width           =   375
   End
   Begin VB.TextBox TxtCantidad 
      Alignment       =   1  'Right Justify
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
      Left            =   1230
      MaxLength       =   80
      TabIndex        =   15
      Top             =   7150
      Width           =   855
   End
   Begin VB.TextBox TxtDescripcionProducto 
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
      Left            =   2100
      MaxLength       =   1200
      TabIndex        =   14
      Top             =   7150
      Width           =   6255
   End
   Begin VB.TextBox TxtCodProducto 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      MaxLength       =   80
      TabIndex        =   13
      Top             =   7150
      Width           =   1095
   End
   Begin VB.TextBox TxtDireccionFiscal 
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
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   10
      Top             =   3315
      Width           =   5655
   End
   Begin VB.TextBox TxtNombreDestino 
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
      Left            =   2920
      MaxLength       =   80
      TabIndex        =   9
      Top             =   3000
      Width           =   4305
   End
   Begin VB.TextBox TxtRucDestino 
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
      Height          =   285
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DtpFechaEmision 
      Height          =   320
      Left            =   1560
      TabIndex        =   6
      Top             =   1440
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
      Format          =   55705601
      CurrentDate     =   41139
   End
   Begin VB.TextBox TxtNumeroDoc 
      Alignment       =   1  'Right Justify
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
      Left            =   5115
      MaxLength       =   80
      TabIndex        =   1
      Top             =   240
      Width           =   1050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDetalle 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   4048
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
   Begin MSDataListLib.DataCombo DtcAlmacenOrigen 
      Height          =   330
      Left            =   10680
      TabIndex        =   2
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
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
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   330
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
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
   Begin MSDataListLib.DataCombo DtcAlmacenDestino 
      Height          =   330
      Left            =   10680
      TabIndex        =   7
      Top             =   675
      Width           =   4575
      _ExtentX        =   8070
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
   Begin VitekeySoft.ChameleonBtn cmdSeriales 
      Height          =   345
      Left            =   14400
      TabIndex        =   42
      Top             =   7140
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "SERIALES"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTransferencia.frx":197EE
      PICN            =   "FrmTransferencia.frx":1980A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdBuscar 
      Height          =   825
      Left            =   16800
      TabIndex        =   52
      Top             =   4095
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
      BTYPE           =   5
      TX              =   "LIST.GUIA"
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
      MICON           =   "FrmTransferencia.frx":19DA4
      PICN            =   "FrmTransferencia.frx":19DC0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdEliminar 
      Height          =   825
      Left            =   16800
      TabIndex        =   53
      Top             =   5820
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
      BTYPE           =   5
      TX              =   "ELIMINAR"
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
      MICON           =   "FrmTransferencia.frx":1A0DA
      PICN            =   "FrmTransferencia.frx":1A0F6
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
      Height          =   825
      Left            =   16800
      TabIndex        =   54
      Top             =   6675
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
      BTYPE           =   5
      TX              =   "ANULAR"
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
      MICON           =   "FrmTransferencia.frx":1A548
      PICN            =   "FrmTransferencia.frx":1A564
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
      Height          =   825
      Left            =   16800
      TabIndex        =   55
      Top             =   7500
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
      BTYPE           =   5
      TX              =   "CERRAR"
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
      MICON           =   "FrmTransferencia.frx":1A87E
      PICN            =   "FrmTransferencia.frx":1A89A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcManifiesto 
      Height          =   315
      Left            =   1560
      TabIndex        =   72
      Top             =   4275
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VitekeySoft.ChameleonBtn cmdmanifiesto 
      Height          =   825
      Left            =   16800
      TabIndex        =   74
      Top             =   4965
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
      BTYPE           =   5
      TX              =   "MANIFIESTO"
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
      MICON           =   "FrmTransferencia.frx":1D8C1
      PICN            =   "FrmTransferencia.frx":1D8DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpTraslado 
      Height          =   315
      Left            =   4320
      TabIndex        =   86
      Top             =   1440
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
      Format          =   55705601
      CurrentDate     =   41139
   End
   Begin VitekeySoft.ChameleonBtn cmdAutomatico 
      Height          =   375
      Left            =   15240
      TabIndex        =   112
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "AUTOMATICO"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmTransferencia.frx":1FE21
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcTipoConsumo 
      Height          =   315
      Left            =   10830
      TabIndex        =   116
      Top             =   7140
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
   Begin MSDataListLib.DataCombo DtcSerieGuia 
      Height          =   330
      Left            =   4080
      TabIndex        =   126
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
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
   Begin MSDataListLib.DataCombo DtcMotivo 
      Height          =   315
      Left            =   5475
      TabIndex        =   132
      Top             =   7635
      Width           =   3135
      _ExtentX        =   5530
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
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "REF:"
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
      Height          =   255
      Left            =   14760
      TabIndex        =   146
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label lbldireccion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECC LLEGADA:"
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
      Height          =   195
      Index           =   2
      Left            =   280
      TabIndex        =   144
      Top             =   3600
      Width           =   1185
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "MOTIVO TRASLADO :"
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
      Left            =   3960
      TabIndex        =   141
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "N� CONTENEDOR :"
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
      Height          =   255
      Left            =   12240
      TabIndex        =   139
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label lbldireccion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION :"
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
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   134
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "N� BULTOS:"
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
      Height          =   255
      Left            =   8880
      TabIndex        =   130
      Top             =   8040
      Width           =   1815
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "MODALIDAD TRASLADO:"
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
      Left            =   8520
      TabIndex        =   123
      Top             =   1320
      Width           =   1875
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "UBIGEO :"
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
      Height          =   195
      Left            =   720
      TabIndex        =   119
      Top             =   2520
      Width           =   600
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "CERTIFICADO :"
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
      Left            =   9240
      TabIndex        =   89
      Top             =   2880
      Width           =   1110
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "F.TRASLADO :"
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
      Height          =   195
      Left            =   3240
      TabIndex        =   87
      Top             =   1515
      Width           =   900
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EMISION :"
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
      Height          =   195
      Left            =   705
      TabIndex        =   85
      Top             =   1440
      Width           =   660
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "OBSERVACION :"
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
      Height          =   375
      Left            =   12240
      TabIndex        =   84
      Top             =   7575
      Width           =   1215
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "REMITENTE :"
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
      Height          =   195
      Left            =   510
      TabIndex        =   79
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "UBIGUEO :"
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
      Height          =   195
      Left            =   660
      TabIndex        =   70
      Top             =   3975
      Width           =   705
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ATENCION :"
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
      Height          =   195
      Left            =   4080
      TabIndex        =   56
      Top             =   4320
      Width           =   765
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "MTC :"
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
      Left            =   9960
      TabIndex        =   41
      Top             =   2520
      Width           =   420
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   795
      Left            =   3840
      Top             =   7560
      Width           =   4935
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Kg"
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
      Left            =   9285
      TabIndex        =   31
      Top             =   7200
      Width           =   180
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "LICENCIA :"
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
      Left            =   9600
      TabIndex        =   28
      Top             =   3720
      Width           =   780
   End
   Begin VB.Label lblRazonChofer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   11925
      TabIndex        =   27
      Top             =   3300
      Width           =   4410
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "CHOFER :"
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
      Left            =   9690
      TabIndex        =   25
      Top             =   3375
      Width           =   705
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "MARCA Y PLACA :"
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
      Left            =   9105
      TabIndex        =   23
      Top             =   2205
      Width           =   1290
   End
   Begin VB.Label lblRazonTransporte 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   11925
      TabIndex        =   22
      Top             =   1800
      Width           =   4410
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "EMP. TRANSPORTE :"
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
      Left            =   8910
      TabIndex        =   20
      Top             =   1800
      Width           =   1515
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   3060
      Left            =   8400
      Top             =   1155
      Width           =   8055
   End
   Begin VB.Label LblCantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   15495
      TabIndex        =   18
      Top             =   7140
      Width           =   975
   End
   Begin VB.Label lbldireccion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECC FISCAL     :"
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
      Height          =   195
      Index           =   0
      Left            =   285
      TabIndex        =   12
      Top             =   3360
      Width           =   1170
   End
   Begin VB.Label lblruc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DESTINATARIO :"
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
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   1050
   End
   Begin VB.Label lblalmacendestino 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "ALMACEN DESTINO  :"
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
      Left            =   8880
      TabIndex        =   5
      Top             =   720
      Width           =   1605
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "ALMACEN ORIGEN  :"
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
      Left            =   8955
      TabIndex        =   4
      Top             =   240
      Width           =   1530
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   975
      Left            =   8400
      Top             =   120
      Width           =   8055
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   7575
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   225
      Top             =   1380
      Width           =   7095
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1740
      Left            =   225
      Top             =   2880
      Width           =   7095
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   3900
      Left            =   120
      Top             =   780
      Width           =   7575
   End
   Begin VB.Shape Shape6 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8415
      Left            =   0
      Top             =   0
      Width           =   17970
   End
End
Attribute VB_Name = "FrmTransferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Public cprod As String
Dim strMotivo As Integer

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub chk_consultar_Click()
Call llenar_serie_guia(Me.DtcTipoDoc.BoundText)
End Sub

Private Sub chk_direccion_Click()
If Me.chk_direccion.Value = 1 Then
   Call llenar_direccion(Me.hfdireccion, Trim(Me.TxtRucDestino.Text))
   Me.frmdireccion.Visible = True
Else
    Me.frmdireccion.Visible = False
End If

End Sub

Public Sub llenar_serie_guia(ByVal in_doc As String)
If Me.chk_consultar.Value = 1 Then
    strCadena = "SELECT serie as Codigo,serie as Descripcion FROM almacen_comprobante WHERE  id_doc='" & in_doc & "' and ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT serie as Codigo,serie as Descripcion FROM almacen_comprobante WHERE id_alm='" & KEY_ALM & "' and id_doc='" & in_doc & "' and ruc='" & KEY_RUC & "'"
End If

Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcSerieGuia)



End Sub


Private Sub chk_manifiesto_Click()
If Me.chk_Manifiesto.Value = 1 Then
   Call load_manifiesto(0)
   Me.DtcManifiesto.Visible = True
   Call load_datos_manifiesto(Me.DtcManifiesto.BoundText)
Else
   Me.DtcManifiesto.Visible = False
End If
End Sub
Public Function control_stock(ByVal in_producto As String, ByVal in_cantidad As Double) As Single
strCadena = "SELECT stock FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then

    If rstK("stock") < in_cantidad And KEY_MOVIMIENTO_SIN_STOCK = "no" Then
     If MsgBox("PRODUCTO NO CUENTA CON STOCK." + Chr(13) + Chr(13) + get_producto(in_producto) + Space(2) + Chr(13) + Chr(13) + "STOCK ACTUAL : " + str(rstK("stock")) + Chr(13) + "TOTAL PEDIDO :" + str(in_cantidad) + Chr(13) + Chr(13) + "Desea Entregar lo Disponible ", vbInformation + vbYesNo, KEY_EMPRESA) = vbYes Then
        control_stock = rstK("stock")
     Else
        control_stock = 0
     End If
     Else
        control_stock = in_cantidad
End If
End If
End Function

Private Sub load_datos_manifiesto(ByVal in_manifiesto As String)
strCadena = "SELECT * FROM transferencia_manifiesto WHERE id_manifiesto='" & Val(in_manifiesto) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   Call BuscarTransporte(rstK("ruc_propietario"))
   Me.TxtMarcayPlaca.Text = rstK("marca")
   Me.TxtPlaca.Text = rstK("placa")
   Me.txtmtc.Text = rstK("placa2")
   Me.txtcertificado.Text = rstK("certificado")
   Call BuscarChofer(Trim(rstK("dni_chofer")))
Else
   Me.TxtMarcayPlaca.Text = ""
   Me.TxtPlaca.Text = ""
   Me.txtmtc.Text = ""
   TxtRucChofer.Text = ""
   Me.TxtLicencia.Text = ""
End If


End Sub
Private Sub load_manifiesto(ByVal in_manifiesto As String)
If Val(in_manifiesto) > 0 Then
    strCadena = "SELECT id_manifiesto as Codigo,CONCAT('N� MANIFIESTO:',id_anio,'-',id_numero) as Descripcion FROM transferencia_manifiesto WHERE id_manifiesto='" & in_manifiesto & "' and  ruc='" & KEY_RUC & "' ORDER BY id_manifiesto "
Else
    strCadena = "SELECT id_manifiesto as Codigo,CONCAT('N� MANIFIESTO:',id_anio,'-',id_numero) as Descripcion FROM transferencia_manifiesto WHERE ruc='" & KEY_RUC & "' ORDER BY id_manifiesto DESC LIMIT 10"
End If
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcManifiesto)
End Sub

Private Sub ChkExtraer_Click()
If Me.ChkExtraer.Value = 1 Then
    Me.Frame1.Visible = True
    Me.DtcTipomovimiento.SetFocus
Else
    Me.Frame1.Visible = False
End If
End Sub

Private Sub cmdAutomatico_Click()
Me.frmtransporte.Visible = True
End Sub

Private Sub cmdcerrardireccion_Click()
Me.frmdireccion.Visible = False
End Sub

Private Sub cmdguia_cerrar_Click()
Me.frmtransporte.Visible = False
End Sub

Private Sub cmdImprimir_Click()


If KEY_GUIA_FRACCIONADA = "si" Then
    Call impresion_formato_grupo_jm(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieGuia.Text), Trim(Me.TxtNumeroDoc.Text))
Else
    Call llenar_impresoras(Me.HfImpresoras)
End If

        
  '      Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Me.TxtSerie.Text, Me.TxtNumeroDoc.Text, Trim(Me.txt_tipo_transferencia.Text), Val(Me.TxtId_transferencia.Text))
End Sub

Private Sub cmdManifiesto_Click()
frmmanifiesto.Show
End Sub
Private Function validar_seriales() As Boolean

validar_seriales = True


strCadena = "select * from view_verificacion_produccion WHERE produccion='si' and id_doc='" & Me.DtcTipoDoc.BoundText & "' and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM movimiento_transferencia_series WHERE id_producto='" & rst("id_producto") & "' and id_doc='" & Me.DtcTipoDoc.BoundText & "' and serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' and numero='" & Trim(Me.TxtNumeroDoc.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1 "
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount < 1 Then
            MsgBox "INGRESE LA SERIE AL PRODUCTO:" + get_producto(rst("id_producto")), vbInformation
            validar_seriales = False
        End If
        rst.MoveNext
   Next i
End If

End Function



Private Sub cmdProcesar_Click()

If Me.chk_consultar.Value = 1 And Val(Me.TxtId_transferencia.Text) < 1 Then
   MsgBox "DESACTIVE LA OPCION CONSULTAR" + Chr(13) + "PARA PORDER GUARDAR", vbInformation
   Exit Sub
End If



If Me.DtcMotivo.BoundText = 3 Then
    If validar_seriales = False Then
        Exit Sub
    End If
End If

If get_periodo_cierre(get_periodo_actual(Me.DtpFechaEmision.Value), "almacen") = True Then
    MsgBox "PERIODO CONTABLE CERRADO ...." + Chr(13) + Chr(13) + "CONSULTE CON CONTABILIDAD", vbInformation, KEY_VENDEDOR
    
    Exit Sub
End If














If Val(Me.TxtId_transferencia.Text) > 0 And Me.cmdProcesar.Enabled = True And Me.cmdverificar.Enabled = True Then
    Call firma_electronica
    Exit Sub
End If



If Val(Me.TxtId_transferencia.Text) > 0 Then
     Call Save
     Exit Sub
End If


If KEY_GUIA_FRACCIONADA = "no" Then
    Call Save
    strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0009' and serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRst(strCadena)
                If rst.RecordCount > 0 Then
                    If rst("electronico") = "si" Then
                        Call firma_electronica
                    End If
                End If
    
 Else
    
    
    Call put_dividir_contenido
    strCadena = "SELECT DISTINCT id_guia_fraccionada FROM movimiento_transferencia_temporal WHERE  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' and id_doc='" & Me.DtcTipoDoc.BoundText & "' and serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' ORDER BY id_guia_fraccionada ASC"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount > 0 Then
       rstA.MoveFirst
       For i = 0 To rstA.RecordCount - 1
            
            
            Call Save_fraccionada(rstA("id_guia_fraccionada"), Trim(Me.TxtNumeroDoc.Text), strMotivo)
            
            rstA.MoveNext
       Next i
    End If
    
    
 End If
End Sub

Private Sub Generar_Electronico(ByVal in_doc As String, ByVal in_serie As String)

    
    If KEY_FACTURACION_ELECTRONICA = "si" Then
                If get_firma_online(Me.DtcTipoDoc.BoundText, Me.DtcSerieGuia.BoundText) = "si" Then
                   Call firma_electronica
                   Exit Sub
                End If
     End If
        
    

End Sub

Private Function firma_electronica()

Dim in_moneda As String
Call disabled_form(Me)
Procedencia = mailenviar
FrmLoad_web_service.Show
FrmLoad_web_service.nom_prcedimiento = "procesar_guia_electronica"

Set FrmLoad_web_service.FormPadre = Me


    

If get_comprobante_produccion(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieGuia.BoundText)) = "si" Then
    
        in_numero = Trim(Me.TxtNumeroDoc.Text)
End If
    
    
If KEY_SERVIDOR_KEYFACIL = "si" Then
        Call FrmLoad_web_service.crear_json_facturacion_electronica("http://api.vitekey.com/keyfact/erp/despatch-documents?password=vitekey2018&company_id=" & KEY_TOKEN_CLOUD & "&ruc=" & KEY_RUC, "POST", json_facturacion_electronica_firmar_guia(Val(Me.TxtId_transferencia.Text)), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
        
   
End If




End Function



Public Sub procesar_guia_electronica(ByVal strHtml As String)
On Error GoTo procesar_nuevamente
Dim in_error As Boolean
Dim in_hash As String
Dim in_numero() As String
Dim json_r As Object
Me.txt_hash.Text = ""
in_hash = ""
Set json_r = JSON.parse(strHtml)
in_error = json_r.Item("error")
If in_error = True Then

Else
     If KEY_SERVIDOR_KEYFACIL = "si" Then
        in_hash = Trim(json_r.Item("response").Item("xml_hash"))
        in_key = Trim(json_r.Item("response").Item("id"))
        'get_numero = Trim(json_r.Item("response").Item("numero"))
     Else
        in_hash = Trim(json_r.Item("response").Item("digest_value"))
        in_key = Trim(json_r.Item("response").Item("key"))
        get_numero = Trim(json_r.Item("response").Item("numero"))
     End If
     
     Me.txt_hash.Text = Trim(in_hash)
     Me.txt_sunat_key.Text = Trim(in_key)
     'Me.TxtNumeroDoc.Text = Trim(get_numero)
     
     strCadena = "UPDATE movimiento_transferencia SET sunat_key='" & Trim(in_key) & "',sunat_hash='" & Trim(in_hash) & "' WHERE id_transferencia='" & Val(Me.TxtId_transferencia.Text) & "'"
     CnBd.Execute (strCadena)
     'Call OrdenImpresion(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcSerieGuia.BoundText), Trim(Me.TxtNumeroDoc.Text))
     'Call next_save
     
                    
     Me.Enabled = True
     Exit Sub
     
     'Call procesar_comprobante
     
End If
Exit Sub
procesar_nuevamente:
MsgBox "SE PRESENTO UN PROBLEMA CON EL INTERNET" + Chr(13) + Chr(13) + "INTENTENTALO NUEVAMENTE.", vbInformation, KEY_USUAURIO
Me.Enabled = True
Me.cmdProcesar.Enabled = True
End Sub



Private Sub put_dividir_contenido()
Dim in_guia As Integer
strCadena = "SELECT * FROM movimiento_transferencia_temporal WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' and id_doc='" & Me.DtcTipoDoc.BoundText & "' and serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' ORDER BY id_temporal ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    in_guia = 1
    For i = 1 To rst.RecordCount
         
         If i > 8 And i <= 16 Then
            in_guia = 2
         End If
         If i > 16 And i <= 24 Then
            in_guia = 3
         End If
         If i > 24 And i <= 32 Then
            in_guia = 4
         End If
         If i > 32 And i <= 40 Then
            in_guia = 5
         End If
         If i > 40 And i <= 48 Then
            in_guia = 6
         End If
         If i > 48 And i <= 56 Then
            in_guia = 7
         End If
         If i > 56 And i <= 64 Then
            in_guia = 8
         End If
         If i > 64 And i <= 72 Then
            in_guia = 9
         End If
         If i > 72 And i <= 80 Then
            in_guia = 10
         End If
         
         If i > 80 And i <= 88 Then
            in_guia = 11
         End If
         
         If i > 88 And i <= 96 Then
            in_guia = 12
         End If
         If i > 96 And i <= 104 Then
            in_guia = 13
         End If
         If i > 104 And i <= 112 Then
            in_guia = 14
         End If
          If i > 112 And i <= 120 Then
            in_guia = 15
         End If
         
         
         strCadena = "UPDATE movimiento_transferencia_temporal SET id_guia_fraccionada='" & in_guia & "' WHERE id_temporal='" & rst("id_temporal") & "' and ruc='" & KEY_RUC & "' LIMIT 1 "
         CnBd.Execute (strCadena)
         rst.MoveNext
    Next i
End If


End Sub

Private Sub cmdverificar_Click()
 
 'Call save_detalle_finalizado(Me.TxtId_transferencia.Text)
 
         Procedencia = modificar
         Call disabled_form(Me)
         
         FrmSeguridad.Show
         Exit Sub
End Sub



Private Sub Command1_Click()


End Sub

Private Sub dtcChofer_Change()
Me.txtLicencia1.Text = get_licencia(Me.dtcChofer.BoundText)
End Sub

Private Sub DtcMarcaPlaca_Change()
strCadena = "SELECT * FROM persona_transporte WHERE id='" & Val(Me.DtcMarcaPlaca.BoundText) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   Me.TxtPlaca1.Text = rstK("placa")
   Me.txtCertficado1.Text = rstK("certificado")
Else
   Me.txtCertficado1.Text = ""
End If
End Sub

Private Sub motivo_traslado(ByVal in_motivo As String)

    Select Case in_motivo
           
           Case 1
                '****VENTA
                strMotivo = 1
                Me.DtcAlmacenDestino.Visible = False
                Me.lblalmacendestino.Visible = False
                Me.DtcTipoConsumo.BoundText = 0
                Me.DtcTipoConsumo.Visible = False
                Me.txtOtros.Text = ""
                Me.txtOtros.Visible = False
           Case 2
                '****DEVOLUCION
                strMotivo = 2
                Me.DtcAlmacenDestino.Visible = False
                Me.lblalmacendestino.Visible = False
                Me.DtcTipoConsumo.BoundText = 0
                Me.DtcTipoConsumo.Visible = False
                Me.txtOtros.Text = ""
                Me.txtOtros.Visible = False
            Case 3
                '****TRANSFERENCIAS
                strMotivo = 3
                Me.txt_idremitente.Text = KEY_RUC
                Me.txtremitente.Text = KEY_EMPRESA
                Me.TxtRucDestino.Text = KEY_RUC
                Me.TxtNombreDestino.Text = KEY_EMPRESA
                Me.txtdireccionfiscal.Text = UCase(BDBuscarCampoRuc("almacen", "direccion", "id_alm", Me.DtcAlmacenDestino.BoundText))
                Me.txtDireccionLlegada.Text = UCase(BDBuscarCampoRuc("almacen", "direccion", "id_alm", Me.DtcAlmacenDestino.BoundText))
                Me.DtcAlmacenDestino.Visible = True
                Me.lblalmacendestino.Visible = True
                Me.DtcTipoConsumo.BoundText = 0
                Me.DtcTipoConsumo.Visible = False
                Me.txtOtros.Visible = False
            Case 4
                strMotivo = 4
                Me.DtcTipoConsumo.Visible = True
                
                strCadena = "SELECT id_tipo_consumo as Codigo,CONCAT('[',id_tipo_consumo,'] ',descripcion) as Descripcion FROM  movimiento_transferencia_consumo WHERE ruc='" & KEY_RUC & "' ORDER BY id_tipo_consumo"
                Call ConfiguraRstT(strCadena)
                Call LlenaDataComboT(Me.DtcTipoConsumo)
                Me.DtcTipoConsumo.Visible = True
                Me.DtcAlmacenDestino.Visible = False
                Me.lblalmacendestino.Visible = False
                Me.txtOtros.Visible = True
    End Select
    

    strMotivo = 1
    
End Sub

Private Sub DtcMotivoOtro_Change()

End Sub

Private Sub DtcMotivoOtro_Click(Area As Integer)

End Sub

Private Sub DtcMotivo_Change()
Call motivo_traslado(Me.DtcMotivo.BoundText)
End Sub

Private Sub DtcSerieGuia_Change()

strCadena = "SELECT numero FROM almacen_comprobante WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' and serie='" & Me.DtcSerieGuia.BoundText & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
   Me.TxtNumeroDoc.Text = rstT("numero")
   strCadena = "UPDATE movimiento_transferencia_temporal SET serie='" & Me.DtcSerieGuia.BoundText & "',numero='" & Trim(Me.TxtNumeroDoc.Text) & "' WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRst(strCadena)
End If
End Sub

Private Sub DtcTipoConsumo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.cmdAgregar.SetFocus
End If
End Sub

Private Sub DtcTipoDoc_Change()
'strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' and  ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"

End Sub


Private Sub select_direccion()
Me.txtid_direccion.Text = Me.hfdireccion.TextMatrix(Me.hfdireccion.Row, 0)
Me.hfdireccion.TextMatrix(Me.hfdireccion.Row, 3) = Chr(254)
strCadena = "SELECT * FROM persona_direccion WHERE id_direccion='" & Val(Me.txtid_direccion.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtdireccionfiscal.Text = Me.hfdireccion.TextMatrix(Me.hfdireccion.Row, 2)
   Me.txtDireccionLlegada.Text = Me.hfdireccion.TextMatrix(Me.hfdireccion.Row, 2)
   Me.txtIdUbigeoDestino.Text = get_ubigeo_sunat_v2(rst("codigo_ubigeo_sunat"), Me.txtIdUbigeoDestino, Me.txtUbigeoDestino)
Else
   Me.txtIdUbigeoOrigen.Text = ""
   Me.txtUbigeoOrigen.Text = ""
End If


'Me.frmdireccion.Visible = False
End Sub
Private Sub cmdagregar_Click()
Dim in_peso As Single
If Val(Me.txtCantidad.Text) > 0 And Trim(Me.TxtCodProducto.Text) <> "" Then
    
    If Me.chk_pesoglobal.Value = 1 Then
        in_peso = Val(Me.txtpeso.Text)
    Else
        in_peso = Val(Me.txtpeso.Text) * Val(Me.txtCantidad.Text)
    End If
    
    
 If get_repetido_transferencia(Trim(Me.TxtCodProducto.Text)) = True Then
    Exit Sub
 End If
    
   
   
   
    
    
If Me.DtcMotivo.BoundText = 4 Then
    GoTo siguiente
End If
    
    
    
If control_stock_general(cprod, Val(Me.txtCantidad.Text), Me.DtcTipoDoc.BoundText) = False Then
    Exit Sub
End If
    
siguiente:
     
    strCadena = "INSERT INTO movimiento_transferencia_temporal(id_doc,serie,numero,id_producto,detalle,cantidad,peso,total,precio_costo,id_tipo_consumo,dni_save,ruc) VALUES " & _
    "('" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Me.TxtNumeroDoc.Text & "','" & cprod & "','" & Trim(Me.TxtDescripcionProducto.Text) & "','" & Val(Me.txtCantidad.Text) & "','" & Val(Me.txtpeso.Text) & "'," & _
    "'" & in_peso & "','" & get_precio_costo(cprod) & "','" & Me.DtcTipoConsumo.BoundText & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
   
    
    strCadena = "SELECT L.produccion FROM producto P,linea L WHERE P.id_linea=L.id_linea AND P.ruc=L.id_usu AND P.id_producto='" & Trim(Me.TxtCodProducto.Text) & "' AND P.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstL(strCadena)
    If rstL("produccion") = "si" Then
        Me.txt_tipo_transferencia.Text = "00002"
    Else
        Me.txt_tipo_transferencia.Text = "00001"
    End If
    
    Call Llenar_Temporal(Me.HfDetalle)
    Me.cmdProcesar.Enabled = True
    
    Me.TxtCodProducto.Text = ""
    Me.txtCantidad.Text = ""
    Me.TxtDescripcionProducto.Text = ""
    Me.TxtUnidad.Text = ""
    Me.txtpeso.Text = ""
    Call Resalta(Me.TxtCodProducto)
Else
    Call Resalta(Me.txtCantidad)
End If
End Sub

Private Sub cmdAnular_Click()
            
            Procedencia = anular
            frmsegurity.Show
            
            
            Exit Sub
End Sub

Private Sub cmdBuscar_Click()
Procedencia = buscar
         FrmGuiasRemision.Show
End Sub

Private Sub cmdcerrarimpresora_Click()
Me.frmimpresoras.Visible = False
End Sub

Private Sub cmdCerrarpantalla_Click()
Me.Frameseries.Visible = False
End Sub

Private Sub cmdEliminar_Click()
            Procedencia = Eliminar
            FrmSeguridad.Show
            Exit Sub
End Sub

Private Sub cmdenviarimpresion_Click()
Dim in_impresora As String

in_impresora = ""
Printer.TrackDefault = True
in_impresora = Trim(Printer.DeviceName)

Call Establecer_Impresora_predeterminada(Trim(Me.HfImpresoras.TextMatrix(Me.HfImpresoras.Row, 1)))

Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Me.DtcSerieGuia.BoundText, Me.TxtNumeroDoc.Text, Trim(Me.txt_tipo_transferencia.Text), Val(Me.TxtId_transferencia.Text))


Call Establecer_Impresora_predeterminada(in_impresora)
Printer.TrackDefault = True
Me.frmimpresoras.Visible = False


Exit Sub

  
  
  
  strCadena = "SELECT id_formato_impresion FROM almacen_comprobante WHERE id_doc='0009' AND serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
    Call ConfiguraRst(strCadena)
    Set Printer = Printers(Val(Me.HfImpresoras.TextMatrix(Me.HfImpresoras.Row, 0)) - 1)
    Call impresion_formato(rst("id_formato_impresion"), "0009", Trim(Me.DtcSerieGuia.BoundText), Trim(Me.TxtNumeroDoc.Text), Trim(Me.txt_tipo_transferencia.Text), Trim(Me.txtdireccionfiscal.Text))
    Me.frmimpresoras.Visible = False
End Sub

Private Sub printer_fraccion()

Printer.TrackDefault = True
Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Me.DtcSerieGuia.BoundText, Me.TxtNumeroDoc.Text, Trim(Me.txt_tipo_transferencia.Text), Val(Me.TxtId_transferencia.Text))


End Sub



Private Sub cmdNuevo_Click()
 Call nuevo
End Sub

Private Sub CmdQuitar_Click()
If MsgBox("ESTA SEGURO DE ELIMINAR ESTE REGISTRO", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
    strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE id_temporal='" & Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    Call Llenar_Temporal(Me.HfDetalle)
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSeriales_Click()
On Error GoTo salir



If (Val(Me.txtEstado.Text) <> 2 And Val(Me.txtEstado.Text) <> 4) Or Val(Me.txtEstado.Text) = 0 Then
    strCadena = "SELECT codigo,nro_chasis,motor,id_estado FROM view_producto_serie WHERE vendido='no' and  id_producto='" & Trim(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 1)) & "' and ruc='" & KEY_RUC & "' and id_alm='" & KEY_ALM & "'"
    Call busca_serial_caja(Me.HfChasis, Trim(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 1)))
Else
    Me.HfChasis.Rows = 0
End If

Call llenar_series(Me.HfSeries, Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcSerieGuia.BoundText), Trim(Me.TxtNumeroDoc.Text))
Me.Frameseries.Visible = True
Call Resalta(Me.txtchasis)

Exit Sub
salir:
End Sub

Private Sub DtcAlmacenDestino_Change()
If Me.DtcAlmacenDestino.Visible = True Then
    
    Me.txtdireccionfiscal.Text = UCase(BDBuscarCampoRuc("almacen", "direccion", "id_alm", Me.DtcAlmacenDestino.BoundText))
    Me.txtDireccionLlegada.Text = Me.txtdireccionfiscal.Text
    
End If
End Sub

Private Sub DtcComprobanterel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtSeri_guia)
End If
End Sub

Private Sub DtcComprobanteGuia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtSeri_guia)
End If
End Sub

Private Sub DtcTipomovimiento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcComprobanteGuia.SetFocus
End If
End Sub

Private Sub DtcTransporte_Change()

Call load_marca(Me.DtcTransporte.BoundText)
Call load_chofer(Me.DtcTransporte.BoundText)

End Sub
Private Sub load_marca(ByVal in_empresa As String)
strCadena = "SELECT id as Codigo,CONCAT(marca,'      -     ',placa) as Descripcion FROM persona_transporte  WHERE    id_persona='" & Trim(in_empresa) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMarcaPlaca)
End Sub

Private Sub load_chofer(ByVal in_empresa As String)
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_chofer_empresa WHERE id_persona='" & Trim(in_empresa) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcChofer)
End Sub


Private Sub Form_Load()
On Error GoTo salir

CenterForm Me
Me.Top = 500
Me.DtpFechaEmision.Value = KEY_FECHA
Me.DtpTraslado.Value = KEY_FECHA

 
If KEY_TRANSPORTE_MIGRA = "si" Then
   Me.cmdAutomatico.Visible = True
Else
   Me.cmdAutomatico.Visible = False
End If
 
strCadena = "SELECT DISTINCT A.id_doc as Codigo, C.doc_des as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' AND C.id_doc IN('0009','0031') AND id_alm='" & KEY_ALM & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    MsgBox "CREE EL COMPROBANTE PARA ESTA SUCURSAL", vbInformation, KEY_EMPRESA
   
    Exit Sub
End If
Call LlenaDataCombo(Me.DtcTipoDoc)

strCadena = "SELECT DISTINCT A.id_doc as Codigo, C.doc_des as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComprobanteGuia)

strCadena = "SELECT id_tipomov as Codigo,descripcion as Descripcion FROM tipo_movimiento ORDER By id_tipomov"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcTipomovimiento)



strCadena = "SELECT id_motivo as Codigo,descripcion as Descripcion FROM movimiento_transferencia_motivo ORDER By id_motivo"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcMotivo)





Call llenar_serie_guia(Me.DtcTipoDoc.BoundText)
Call nuevo







strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacenOrigen)
Me.DtcAlmacenOrigen.BoundText = KEY_ALM
Me.DtcAlmacenOrigen.Locked = True

strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE id_tipoentidad='0' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacenDestino)

Me.cmdverificar.Enabled = False

For i = 0 To Printers.Count - 1
   If UCase(Trim(Printers(i).DeviceName)) = UCase(Trim(sNombreImpresora)) Then
        KEY_PRINTER = Printers(i).DeviceName
        bEncontrada = True
        Exit For
    End If
Next i


Exit Sub
salir:



End Sub

Private Sub BuscarResponsable(ByVal ruc As String)
If (Trim(ruc) = "") Then
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
End If
    strCadena = "SELECT *  FROM persona WHERE dni='" & ruc & "'"
    Call ConfiguraRstAux(strCadena)
    If rstAux.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.txtRuc.Text = ruc
        FrmDetallePersona.ChkPersonal.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.TxtRucDestino.Text = rstAux("dni")
        Me.TxtNombreDestino.Text = rstAux("nombre_completo")
        Me.txtdireccionfiscal.Text = rstAux("direccion")
        Me.txtDireccionLlegada.Text = rstAux("direccion")
        Call get_ubigeo_sunat((Me.TxtRucDestino.Text), Me.txtIdUbigeoDestino, Me.txtUbigeoDestino)
        Call Resalta(Me.TxtRucTransporte)
        Exit Sub
       
    End If

End Sub
Private Sub BuscarRemitente(ByVal ruc As String)
If (Trim(ruc) = "") Then
    Procedencia = seleccionar_otro
    FrmPersona.Show
    Exit Sub
End If
    strCadena = "SELECT *  FROM persona WHERE dni='" & ruc & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.txtRuc.Text = ruc
        FrmDetallePersona.ChkPersonal.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.txt_idremitente.Text = rst("dni")
        Me.txtremitente.Text = rst("nombre_completo")
        
        
        Exit Sub
       
    End If

End Sub

Private Sub BuscarChofer(ByVal ruc As String)
If (Trim(ruc) = "") Then
    Procedencia = modificar
    FrmPersona.Show
    Exit Sub
End If
    strCadena = "SELECT *  FROM persona WHERE dni='" & ruc & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.txtRuc.Text = ruc
        FrmDetallePersona.ChkTransporte.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.TxtRucChofer.Text = rst("dni")
        Me.lblRazonChofer.Caption = rst("nombre_completo")
        If IsNull(rst("licencia")) = True Then
            Me.TxtLicencia.Text = ""
        Else
            Me.TxtLicencia.Text = rst("licencia")
        End If
        
         Call Resalta(Me.TxtLicencia)
        Exit Sub
       
    End If

End Sub
Private Sub BuscarTransporte(ByVal ruc As String)
If (Trim(ruc) = "") Then
    Procedencia = buscar
    FrmPersona.Show
    Exit Sub
End If
    strCadena = "SELECT dni,nombre_completo FROM persona WHERE dni='" & ruc & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.txtRuc.Text = ruc
        FrmDetallePersona.ChkTransporte.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.TxtRucTransporte.Text = rst("dni")
        Me.lblRazonTransporte.Caption = rst("nombre_completo")
        Call Resalta(Me.TxtMarcayPlaca)
        Exit Sub
       
    End If

End Sub

Private Sub HfChasis_Click()
If Val(Me.HfChasis.TextMatrix(Me.HfChasis.Row, 0)) > 0 Then
    If Val(Me.HfSeries.Rows) > 0 Then
            If Val(Me.HfSeries.Rows) <= Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 3)) Then
                Call ActualizarImagen(Me.HfChasis.TextMatrix(Me.HfChasis.Row, 0), Me.HfChasis)
            Else
                MsgBox "INGRESE SOLO LA CANTIDAD DE SERIES A TRASLADAR", vbInformation
            End If
    Else
               Call ActualizarImagen(Me.HfChasis.TextMatrix(Me.HfChasis.Row, 0), Me.HfChasis)
    End If
End If
End Sub
Private Sub ActualizarImagen(ByVal id_detalle As Double, ByVal Grilla As MSHFlexGrid)
     Dim estado As String
      strCadena = "SELECT t.serie,t.numero FROM movimiento_transferencia_series s,movimiento_transferencia t WHERE t.anulado='no' and s.anulado='no' and  s.id_transferencia=t.id_transferencia  and s.chasis='" & Me.HfChasis.TextMatrix(Me.HfChasis.Row, 1) & "' and id_producto='" & Trim(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 1)) & "' and t.id_alm_origen='" & KEY_ALM & "' and t.ruc='" & KEY_RUC & "'"
      Call ConfiguraRst(strCadena)
      If rst.RecordCount < 1 Then
        strCadena = "DELETE FROM movimiento_transferencia_series WHERE chasis='" & Me.HfChasis.TextMatrix(Me.HfChasis.Row, 1) & "' and id_producto='" & Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 1) & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        strCadena = "SELECT * FROM imp_producto_detalle WHERE nro_chasis='" & Me.HfChasis.TextMatrix(Me.HfChasis.Row, 1) & "' and ruc='" & KEY_RUC & "' LIMIT 0,1"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
            strCadena = "INSERT INTO movimiento_transferencia_series(id_doc,serie,numero,id_producto,chasis,motor,nro_dua,anio_fabricacion,nro_item,ruc)VALUES('" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 1) & "','" & Me.HfChasis.TextMatrix(Me.HfChasis.Row, 1) & "','" & Me.HfChasis.TextMatrix(Me.HfChasis.Row, 2) & "','" & rstZ("nro_contenedor") & "','" & rstZ("anio_fabricacion") & "','" & rstZ("item") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        End If
           
           
           Me.HfChasis.TextMatrix(Me.HfChasis.Row, 4) = Chr(254)
           Call llenar_series(Me.HfSeries, Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieGuia.BoundText), Trim(Me.TxtNumeroDoc.Text))
           Exit Sub
        Else
            MsgBox "Esta Serie ya esta asignada" + Chr(13) + Chr(13) + "GUIA R:" + rst("serie") + "-" + rst("numero"), vbInformation
            Exit Sub
        End If
        
      
      


      
      
      Call Resalta(Me.txtchasis)
      
    End Sub

Private Sub HfDetalle_DblClick()
On Error GoTo salir
If Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) > 0 Then
    Me.cmdProcesar.Enabled = True
    
        FrmTransferencias_detalle.Show
        FrmTransferencias_detalle.txtid_producto.Text = Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 1)
    End If
    Exit Sub
salir:
End Sub

Private Sub HfDetalle_SelChange()
If Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) > 0 Then
    Me.CmdQuitar.Visible = True
Else
    Me.CmdQuitar.Visible = False
End If
End Sub

Private Sub hfdireccion_Click()
    Call select_direccion
End Sub


Private Sub HfSeries_Click()
If Val(Me.HfSeries.TextMatrix(Me.HfSeries.Row, 0)) > 0 Then
    Call Actualizar_serie(Me.HfSeries.TextMatrix(Me.HfSeries.Row, 0), Me.HfSeries)
End If
End Sub
Private Sub Actualizar_serie(ByVal id_detalle As Double, ByVal Grilla As MSHFlexGrid)
     Dim estado As String
      
     If Val(Me.TxtId_transferencia.Text) < 1 And Me.lblAnulado.Visible = False Then
      
      If MsgBox("Esta seguro de quitar este Item", vbQuestion + vbYesNo) = vbYes Then
        strCadena = "DELETE FROM movimiento_transferencia_series WHERE id_detalle='" & id_detalle & "'"
        CnBd.Execute (strCadena)
         
        Call llenar_series(Me.HfSeries, Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieGuia.BoundText), Trim(Me.TxtNumeroDoc.Text))
        
      End If
    Else
    If Trim(Me.HfSeries.TextMatrix(Me.HfSeries.Row, 3)) = Chr(168) Then
        strCadena = "SELECT * from imp_producto_detalle WHERE nro_chasis='" & Trim(Me.HfSeries.TextMatrix(Me.HfSeries.Row, 1)) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
            strCadena = "INSERT INTO imp_producto_detalle (`id_compra`,`id_detalle_compra`,`id_producto`,`serie`,`id_estado`,`id_estado_detalle`,`id_alm`, " & _
            "`anio_fabricacion`,`anio_contenedor`,`nro_contenedor`,`nro_chasis`,`nro_motor`,`anio_modelo`,`item`,`dni_save`,`nsave`,`fecha_mod`,`hora_mod`,`serie_asignada`,`vendido`,`id_orden`,`ruc`) " & _
            "VALUES('" & rstZ("id_compra") & "','" & rstZ("id_detalle_compra") & "','" & rstZ("id_producto") & "','" & rstZ("serie") & "','" & rstZ("id_estado") & "','" & rstZ("id_estado_detalle") & "','" & KEY_ALM & "' " & _
            ",'" & rstZ("anio_fabricacion") & "','" & rstZ("anio_contenedor") & "','" & rstZ("nro_contenedor") & "','" & rstZ("nro_chasis") & "', " & _
            " '" & rstZ("nro_motor") & "','" & rstZ("anio_modelo") & "','" & rstZ("item") & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_FECHA & "',CURTIME(),'" & rstZ("serie_asignada") & "','" & rstZ("vendido") & "','" & rstZ("id_orden") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
            strCadena = "UPDATE movimiento_transferencia_series SET recibido='si' WHERE id_detalle='" & Val(Me.HfSeries.TextMatrix(Me.HfSeries.Row, 0)) & "'  and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
             
            Me.HfSeries.TextMatrix(Me.HfSeries.Row, 3) = Chr(254)
       
        End If
         Else
            If Me.lblAnulado.Visible = True Then
                MsgBox "Documento Anulado, valores de referencia", vbInformation
            Else
                MsgBox "Este producto fue recibido CORRECTAMENTE", vbInformation
            End If
            
    End If
      
 End If
      
      


      
      
      Call Resalta(Me.txtchasis)
      
    End Sub

Private Sub HfUbigeo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.HfUbigeo.Rows > 0 Then
       
       If Procedencia = buscar Then
            Me.txtIdUbigeoOrigen.Text = Me.HfUbigeo.TextMatrix(Me.HfUbigeo.Row, 0)
            txtUbigeoOrigen.Text = get_ubigeo_sunat_descripcion(Me.txtIdUbigeoOrigen.Text)
            Procedencia = Neutro
            Me.frmUbigeo.Visible = False
            Exit Sub
       End If
       
       If Procedencia = Selecionar Then
            Me.txtIdUbigeoDestino.Text = Me.HfUbigeo.TextMatrix(Me.HfUbigeo.Row, 0)
            Me.txtUbigeoDestino.Text = get_ubigeo_sunat_descripcion(Me.txtIdUbigeoDestino.Text)
            Procedencia = Neutro
            Me.frmUbigeo.Visible = False
            Exit Sub
       End If
       
       
        
    End If
End If
End Sub

Private Sub Opt_transporte_privado_Click()
If Me.Opt_transporte_privado.Value = True Then
   Me.TxtRucTransporte.Text = KEY_RUC
   lblRazonTransporte.Caption = get_persona(KEY_RUC)
End If
End Sub

Private Sub OptDevolucion_Click()

End Sub

Private Sub OptGrinter_Click()
End Sub

Private Sub Option2_Click()

End Sub

Private Sub OptOtros_Click()
End Sub

Private Sub OpVenta_Click()
End Sub

Public Sub load_transporte()
If KEY_TRANSPORTE_MIGRA = "si" Then
   strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_transporte='si' and  ruc='" & KEY_RUC & "'"
   Call ConfiguraRst(strCadena)
   Call LlenaDataCombo(Me.DtcTransporte)
   Me.frmtransporte.Visible = True
Else
   Me.frmtransporte.Visible = False
End If

End Sub
Public Sub nuevo()


strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0009' and serie='" & Me.DtcSerieGuia.BoundText & "'  AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Me.DtcTipoDoc.Enabled = True
    Me.txt_idremitente.Text = ""
    Me.txtremitente.Text = ""
    Me.TxtNumeroDoc.Enabled = True
    Me.TxtDescripcionProducto.Enabled = True

    Me.TxtNumeroDoc.Text = rst("numero")
    Me.TxtRucDestino.Text = ""
    Me.txtid_direccion.Text = ""
    Me.frmdireccion.Visible = False
    Me.txtid_vinculado.Text = ""
    Me.TxtMarcayPlaca.Text = ""
    Me.TxtPlaca.Text = ""
    Me.txtmtc.Text = ""
    Me.txtcontenedor.Text = ""
    Me.txtdocumentoReferencia.Text = ""
    Me.TxtBultos.Text = ""
    Me.TxtLicencia.Text = ""
    Me.TxtRucChofer.Text = ""
    Me.lblRazonChofer.Caption = ""
    Me.txtpesototal.Text = 0
    
Me.Txtmoneda.Text = "00001"
Me.txtpesototal.Text = ""
Me.txtTc.Text = KEY_CAMBIO
Me.txtObservacion.Text = ""
Me.txtIdUbigeoOrigen.Text = ""
Me.txtUbigeoOrigen.Text = ""
Me.txtIdUbigeoDestino.Text = ""
Me.txtUbigeoDestino.Text = ""
Me.chk_direccion.Value = 0
Me.TxtNombreDestino.Text = ""
Me.txtdireccionfiscal.Text = ""
Me.txtDireccionLlegada.Text = ""
Me.TxtRucTransporte.Text = ""
Me.lblRazonTransporte.Caption = ""
Me.txt_idremitente.Text = KEY_RUC
Me.txtremitente.Text = KEY_EMPRESA
Me.txtDireccionPartida.Text = KEY_DIRECCION_ALM
Call get_ubigeo_sunat(KEY_RUC, Me.txtIdUbigeoOrigen, Me.txtUbigeoOrigen)



Me.TxtCodProducto.Enabled = True
Me.lblAnulado.Visible = False
Me.txt_atencion.Text = ""
Me.txt_dni_atencion.Text = ""

Me.TxtId_transferencia.Text = ""
Me.frmanulado.Visible = False
Me.DtpFechaEmision.Value = KEY_FECHA
Me.cmdverificar.Enabled = False
Call Resalta(Me.TxtRucDestino)

Me.cmdImprimir.Enabled = False
Me.cmdAgregar.Enabled = True

Me.ChkExtraer.Value = 0
Me.TxtSeri_guia.Text = ""
Me.TxtNumero_guia.Text = ""
Me.txtvalor_mercaderia.Text = ""
Me.chk_pesoglobal.Value = 0
Me.chk_valor_mercaderia.Value = 0
Me.txtdiferida.Text = "no"
Me.txtEstado.Text = ""
Me.Opt_transporte_privado.Value = True

If KEY_USUARIO = "42546269" Or KEY_USUARIO = "900001" Or KEY_USUARIO = "46947665" Or KEY_USUARIO = "71574340" Then
    Me.cmdEliminar.Enabled = True
Else
    Me.cmdEliminar.Enabled = False
End If



Call Me.load_transporte

End If

Call Llenar_Temporal(Me.HfDetalle)

End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub


Public Sub llenar_impresoras(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
   Grilla.Clear
   Grilla.Rows = 0
   If Printers.Count > 0 Then
      Me.frmimpresoras.Visible = True
      Me.cmdenviarimpresion.Visible = True
    Else
        Me.frmimpresoras.Visible = False
      Me.cmdenviarimpresion.Visible = False
      Exit Sub
   End If
  

       ReDim arrColWidth(1 To Printers.Count)
       'For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 400
           Grilla.ColWidth(1) = 3100
           
       'Next
        cabecera = "N�" & vbTab & "IMPRESORA"
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        
            For i = 0 To Printers.Count - 1
                Fila = Format(i + 1, "00") & vbTab & UCase(Trim(Printers(i).DeviceName))
                Grilla.AddItem Fila
            Next i
salir:
Exit Sub
End Sub

Private Function verficar_series_recibidas(ByVal id_transferencia As Double) As Boolean
    
    strCadena = "SELECT * FROM movimiento_transferencia_series WHERE id_transferencia='" & id_transferencia & "' and  recibido='no' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        If MsgBox("Hay series pendientes por RECIBIR" + Chr(13) + "Tiene que seleccionar por lo menos una serie.", vbInformation + vbYesNo) = vbYes Then
            verficar_series_recibidas = False
         Else
            verficar_series_recibidas = False
        End If
        
Else
    verficar_series_recibidas = True
    End If
    
    
End Function
Private Function verifica_existe(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_remitente As String) As Boolean
If in_doc = "0009" Then
    strCadena = "SELECT * FROM movimiento_transferencia WHERE id_doc='" & in_doc & "' and serie='" & in_serie & "' and numero='" & in_numero & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Else
    strCadena = "SELECT * FROM movimiento_transferencia WHERE id_doc='" & in_doc & "' and serie='" & in_serie & "' and numero='" & in_numero & "' and id_remitente='" & in_remitente & "' and ruc='" & KEY_RUC & "' LIMIT 1"
End If
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   verifica_existe = True
Else
   verifica_existe = False
End If

End Function
Private Sub Save()
Dim id_transferencia As Double

If Me.DtcAlmacenOrigen.BoundText = Me.DtcAlmacenDestino.BoundText And strMotivo = 3 Then
    MsgBox "SELECCIONE ALMACENES DISTINTOS", vbInformation, KEY_EMPRESA
    Me.DtcAlmacenDestino.SetFocus
    Exit Sub
End If

If Me.DtcTipoDoc.BoundText = "" Or Me.DtcSerieGuia.BoundText = "" Or Me.TxtNumeroDoc.Text = "" Then
   MsgBox "LLENE TODOS LOS PARAMETROS", vbInformation, KEY_EMPRESA
   Exit Sub
Else
        
    If Val(Me.TxtId_transferencia.Text) < 1 Then
        If verifica_existe(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieGuia.BoundText), Trim(Me.TxtNumeroDoc.Text), Trim(Me.txt_idremitente.Text)) = True Then
                MsgBox "Documento ya Ingresado", vbInformation
           Exit Sub
        End If
        
        
        If Val(Me.txtpesototal.Text) <= 0 Then
            MsgBox "Debe ingresar un Peso Referencial", vbInformation
            Exit Sub
        End If
        
        If Val(Me.TxtBultos.Text) <= 0 Then
            MsgBox "Debe ingresar Numero de Bultos Referencial", vbInformation
            Exit Sub
        End If
        
        If Me.frmtransporte.Visible = True Then
            Me.Opt_transporte_publico.Value = True
        End If
        
        If Me.Opt_transporte_privado.Value = True Then
            
            If Me.TxtLicencia.Text = "" Then
                 MsgBox "Debe ingresar una Licencia Valida", vbInformation
                Exit Sub
            End If
            
            If Me.TxtMarcayPlaca.Text = "" Or Me.TxtPlaca.Text = "" Then
                 MsgBox "Debe ingresar una Marca y Placa Valida", vbInformation
                 Exit Sub
            End If
         Else
            If Trim(Me.TxtRucTransporte.Text) = "" Then
               MsgBox "Debe ingresar un Ruc de Transportista", vbInformation
               Exit Sub
            End If
        End If
        
        
        
        
        If Me.frmtransporte.Visible = True Then
            strCadena = "UPDATE persona_transporte SET certificado='" & Trim(Me.txtCertficado1.Text) & "' WHERE id='" & Me.DtcMarcaPlaca.BoundText & "'"
            CnBd.Execute (strCadena)
        End If
        If Me.chk_Manifiesto.Value = 1 Then
           in_manifiesto = Me.DtcManifiesto.BoundText
        Else
           in_manifiesto = 0
        End If
        
        If Me.frmtransporte.Visible = True Then
            in_transporte = Me.DtcTransporte.BoundText
            in_marca = get_marca_2(Me.DtcMarcaPlaca.BoundText)
            in_placa = Me.TxtPlaca1.Text
            in_chofer = Me.dtcChofer.BoundText
            in_mtc = Trim(Me.txtMtc1.Text)
            in_licencia = Me.txtLicencia1.Text
            in_certificado = Trim(Me.txtCertficado1.Text)
            strCadena = "call sp_marca_placa('" & Val(Me.DtcMarcaPlaca.BoundText) & "','" & Trim(Me.DtcTransporte.BoundText) & "','" & Trim(in_marca) & "','" & Trim(in_placa) & "','" & Trim(Me.txtCertficado1.Text) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
            in_transporte = Trim(Me.TxtRucTransporte.Text)
            in_marca = Trim(Me.TxtMarcayPlaca.Text)
            in_placa = Trim(Me.TxtPlaca.Text)
            in_chofer = Trim(Me.TxtRucChofer.Text)
            in_mtc = Trim(Me.txtmtc.Text)
            in_licencia = Trim(Me.TxtLicencia.Text)
            in_certificado = Trim(Me.txtcertificado.Text)
        End If
        
        If Me.DtcMotivo.BoundText = 3 Then
            in_estado = 4
        Else
            in_estado = 1
        End If
        
         If Me.DtcMotivo.BoundText = 4 Then
           in_tipo_consumo = Me.DtcTipoConsumo.BoundText
           
        Else
           in_tipo_consumo = 0
        End If
        
        Me.txtObservacion.Text = Trim(Me.txtObservacion.Text)
        
        
        
        
        If Me.Opt_transporte_privado.Value = True Then
           in_tipo_transporte = 2
        End If
        If Me.Opt_transporte_publico.Value = True Then
           in_tipo_transporte = 1
        End If
        
        
        
        
        
        in_doc_origen = ""
        in_doc_relacionado = ""
        in_doc_relacionado = Me.txtdocumentoReferencia.Text
        
        If Me.ChkExtraer.Value = 1 Then
           in_doc_origen = Me.DtcComprobanteGuia.BoundText
           in_doc_relacionado = DtcComprobanteGuia.Text & ":" & Me.TxtSeri_guia.Text & "-" & Trim(Me.TxtNumero_guia.Text)
       
           
        End If
        
        
        strCadena = "INSERT INTO movimiento_transferencia(id_doc,dni_atencion,atencion,id_tipo_guia,serie,numero,fecha,fecha_traslado,id_direccion," & _
        "id_manifiesto,direccion,id_remitente,remitente,id_destinatario,destinatario,direccion_destino,direccion_fiscal,id_transporte,marca_placa,placa," & _
        "id_chofer,id_alm_origen,id_alm_destino,id_motivo,motivo_otros,observacion,id_venta,dni_save,mtc,licencia,cantidad_total,peso_total," & _
        "certificado,valor_mercaderia,id_estado,id_moneda,tc,id_tipo_consumo,diferida,tipo_transporte,numero_bultos,ubigeo_origen,ubigeo_destino,contenedor,id_doc_origen,comprobante_relacionado,ruc) " & _
        "VALUES('" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.txt_dni_atencion.Text) & "','" & Trim(Me.txt_atencion.Text) & "'," & _
        "'" & Trim(Me.txt_tipo_transferencia.Text) & "','" & Me.DtcSerieGuia.BoundText & "','" & Me.TxtNumeroDoc.Text & "','" & Format(Me.DtpFechaEmision.Value, "YYYY-mm-dd") & "'," & _
        "'" & Format(Me.DtpTraslado.Value, "YYYY-mm-dd") & "','" & Val(Me.txtid_direccion.Text) & "','" & in_manifiesto & "','" & Trim(Me.txtDireccionPartida.Text) & "'," & _
        "'" & Trim(Me.txt_idremitente.Text) & "','" & Trim(Me.txtremitente.Text) & "','" & Me.TxtRucDestino.Text & "','" & Trim(Me.TxtNombreDestino.Text) & "'," & _
        "'" & Replace(Me.txtDireccionLlegada.Text, "'", "") & "','" & Replace(Me.txtdireccionfiscal.Text, "'", "") & "','" & in_transporte & "','" & in_marca & "','" & in_placa & "','" & in_chofer & "','" & Me.DtcAlmacenOrigen.BoundText & "'," & _
        "'" & Me.DtcAlmacenDestino.BoundText & "','" & Me.DtcMotivo.BoundText & "','" & Me.txtOtros.Text & "','" & Trim(Me.txtObservacion.Text) & "'," & _
        "'" & Val(Me.txtid_venta.Text) & "','" & KEY_USUARIO & "','" & in_mtc & "','" & in_licencia & "','" & Val(Me.txtcantidad_total.Text) & "'," & _
        "'" & Val(Me.txtpesototal.Text) & "','" & in_certificado & "','" & Val(Me.txtvalor_mercaderia.Text) & "','" & in_estado & "',  " & _
        "'" & Trim(Me.Txtmoneda.Text) & "'  ,'" & Val(Me.txtTc.Text) & "'," & _
        "'" & in_tipo_consumo & "','" & Trim(Me.txtdiferida.Text) & "','" & in_tipo_transporte & "','" & Val(Me.TxtBultos.Text) & "'," & _
        "'" & Trim(Me.txtIdUbigeoOrigen.Text) & "','" & Trim(Me.txtIdUbigeoDestino.Text) & "','" & Trim(Me.txtcontenedor.Text) & "','" & in_doc_origen & "','" & in_doc_relacionado & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        id_transferencia = LastRegistroRUC("movimiento_transferencia", "id_transferencia")
        
        If Val(Me.txtid_vinculado.Text) > 0 Then
            strCadena = "UPDATE movimiento_transferencia SET id_guia_remitente='" & Val(id_transferencia) & "' WHERE id_transferencia='" & Val(Me.txtid_vinculado.Text) & "'"
            CnBd.Execute (strCadena)
        End If
        
        If Val(Me.txtid_venta.Text) > 0 Then
            strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & Val(Me.txtid_venta.Text) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
               rst.MoveFirst
               For i = 0 To rst.RecordCount - 1
                    strCadena = "INSERT INTO movimiento_transferencia_series(id_transferencia,id_doc,serie,numero,id_producto,chasis,motor,nro_dua,anio_fabricacion,nro_item,ruc)VALUES('" & id_transferencia & "','" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & rst("id_producto") & "','" & rst("nro_chasis") & "','" & rst("serie") & "','" & rst("nro_dua") & "','" & rst("anio_fabricacion") & "','" & rst("nro_item") & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    rst.MoveNext
               Next i
            End If
        Else
        
        strCadena = "UPDATE movimiento_transferencia_series SET id_transferencia='" & id_transferencia & "' WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' and serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' and numero='" & Trim(Me.TxtNumeroDoc.Text) & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
        If Me.DtcMotivo.BoundText = 3 Then
            Call put_transferencia_serie(id_transferencia)
        End If
        End If
        
        If Me.TxtLicencia.Text <> "" Then
            strCadena = "UPDATE persona SET licencia='" & Me.TxtLicencia.Text & "' WHERE dni='" & Me.TxtRucChofer.Text & "'"
            CnBd.Execute (strCadena)
        End If
    Else
         If verficar_series_recibidas(Val(Me.TxtId_transferencia.Text)) = False Then
             Exit Sub
         End If
         strCadena = "UPDATE movimiento_transferencia SET peso_total='" & Val(Me.txtpesototal.Text) & "',valor_mercaderia='" & Val(Me.txtvalor_mercaderia.Text) & "',observacion='" & UCase(Trim(Me.txtObservacion.Text)) & "',finalizado='si',id_recibio='" & KEY_USUARIO & "',nsave='" & KEY_VENDEDOR & "',id_estado='2' WHERE id_transferencia='" & Val(Me.TxtId_transferencia.Text) & "' ANd ruc='" & KEY_RUC & "'"
         CnBd.Execute (strCadena)
         'Call savedetalle(Val(Me.TxtId_transferencia.Text), "si")
         Call save_detalle_finalizado(Val(Me.TxtId_transferencia.Text))
         Me.cmdProcesar.Enabled = False
         Me.cmdImprimir.Enabled = True
         
         Exit Sub
    End If
    
    Me.cmdProcesar.Enabled = False
    Me.cmdImprimir.Enabled = True
    Me.TxtId_transferencia.Text = id_transferencia
    StrNumero = FormatosCeros(Trim(str(Val(Me.TxtNumeroDoc.Text)) + 1), 6)
    strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE id_alm='" & Trim(Me.DtcAlmacenOrigen.BoundText) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' AND ruc='" & Trim(KEY_RUC) & "'"
    CnBd.Execute (strCadena)
    
    Call savedetalle(id_transferencia, "no")
    If KEY_CONTABILIDAD = "si" Then
            
           ' If Me.OptDevolucion.Value = True Then
                    'strCadena = "call CON_InsertaAsiento_Devolucion('" & id_transferencia & "')"
                    'CnBd.Execute (strCadena)
           ' End If
            
           ' If Me.OptOtros.Value = True Then
           '      strCadena = "call CON_InsertaAsiento_ConsumoInterno('" & id_transferencia & "')"
           '      CnBd.Execute (strCadena)
           ' End If
            
            
            If Me.DtcMotivo.BoundText = 1 Then
                If get_diferida(Val(Me.txtid_venta.Text)) = "si" Then
                    strCadena = "call CON_InsertaAsiento_GuiaDiferida('" & Val(id_transferencia) & "')"
                    CnBd.Execute (strCadena)
                End If
            End If
            
            
            End If
    End If

End Sub



Private Sub Save_fraccionada(ByVal in_guia As String, ByVal in_numero As String, ByVal in_motivo As String)
Dim id_transferencia As Double

If Me.DtcAlmacenOrigen.BoundText = Me.DtcAlmacenDestino.BoundText And in_motivo = 3 Then
    MsgBox "SELECCIONE ALMACENES DISTINTOS", vbInformation, KEY_EMPRESA
    Me.DtcAlmacenDestino.SetFocus
    Exit Sub
End If

If Me.DtcTipoDoc.BoundText = "" Or Me.DtcSerieGuia.BoundText = "" Or Me.TxtNumeroDoc.Text = "" Then
   MsgBox "LLENE TODOS LOS PARAMETROS", vbInformation, KEY_EMPRESA
   Exit Sub
Else
        
    If Val(Me.TxtId_transferencia.Text) < 1 Then
        If verifica_existe(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieGuia.BoundText), Trim(Me.TxtNumeroDoc.Text), Trim(Me.txt_idremitente.Text)) = True Then
                MsgBox "Documento ya Ingresado", vbInformation
           Exit Sub
        End If
        If Me.frmtransporte.Visible = True Then
            strCadena = "UPDATE persona_transporte SET certificado='" & Trim(Me.txtCertficado1.Text) & "' WHERE id='" & Me.DtcMarcaPlaca.BoundText & "'"
            CnBd.Execute (strCadena)
        End If
        If Me.chk_Manifiesto.Value = 1 Then
           in_manifiesto = Me.DtcManifiesto.BoundText
        Else
           in_manifiesto = 0
        End If
        
        If Me.frmtransporte.Visible = True Then
            in_transporte = Me.DtcTransporte.BoundText
            in_marca = get_marca_2(Me.DtcMarcaPlaca.BoundText)
            in_placa = Me.TxtPlaca1.Text
            in_chofer = Me.dtcChofer.BoundText
            in_mtc = Trim(Me.txtMtc1.Text)
            in_licencia = Me.txtLicencia1.Text
            in_certificado = Trim(Me.txtCertficado1.Text)
             strCadena = "call sp_marca_placa('" & Val(Me.DtcMarcaPlaca.BoundText) & "','" & Trim(Me.DtcTransporte.BoundText) & "','" & Trim(in_marca) & "','" & Trim(in_placa) & "','" & Trim(Me.txtCertficado1.Text) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
            in_transporte = Trim(Me.TxtRucTransporte.Text)
            in_marca = Trim(Me.TxtMarcayPlaca.Text)
            in_placa = Trim(Me.TxtPlaca.Text)
            in_chofer = Trim(Me.TxtRucChofer.Text)
            in_mtc = Trim(Me.txtmtc.Text)
            in_licencia = Trim(Me.TxtLicencia.Text)
            in_certificado = Trim(Me.txtcertificado.Text)
        End If
        
        If Me.DtcMotivo.BoundText = 3 Then
        
            in_estado = 4
        Else
            in_estado = 1
        End If
        
        If Me.DtcMotivo.BoundText = 4 Then
           in_tipo_consumo = Me.DtcTipoConsumo.BoundText
           
        Else
           in_tipo_consumo = 0
        End If
        
        
        
        
        
        strCadena = "INSERT INTO movimiento_transferencia(id_doc,dni_atencion,atencion,id_tipo_guia,serie,numero,fecha,fecha_traslado,id_direccion,id_manifiesto,direccion,id_remitente,remitente,id_destinatario,destinatario,id_transporte,marca_placa,placa,id_chofer,id_alm_origen,id_alm_destino,id_motivo,motivo_otros,observacion,id_venta,dni_save,mtc,licencia,cantidad_total,peso_total,certificado,valor_mercaderia,id_estado,id_moneda,tc,id_tipo_consumo,ruc) " & _
        "VALUES('" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.txt_dni_atencion.Text) & "','" & Trim(Me.txt_atencion.Text) & "','" & Trim(Me.txt_tipo_transferencia.Text) & "','" & Me.DtcSerieGuia.BoundText & "','" & in_numero & "','" & KEY_FECHA & "','" & Format(Me.DtpTraslado.Value, "YYYY-mm-dd") & "','" & Val(Me.txtid_direccion.Text) & "','" & in_manifiesto & "','" & Trim(Me.txtdireccionfiscal.Text) & "','" & Trim(Me.txt_idremitente.Text) & "','" & Trim(Me.txtremitente.Text) & "','" & Me.TxtRucDestino.Text & "','" & Trim(Me.TxtNombreDestino.Text) & "'," & _
        "'" & in_transporte & "','" & in_marca & "','" & in_placa & "','" & in_chofer & "','" & Me.DtcAlmacenOrigen.BoundText & "','" & Me.DtcAlmacenDestino.BoundText & "','" & in_motivo & "','" & Trim(Me.txtOtros.Text) & "','" & Trim(Me.txtObservacion.Text) & "','" & Val(Me.txtid_venta.Text) & "','" & KEY_USUARIO & "','" & in_mtc & "','" & in_licencia & "','" & Val(Me.txtcantidad_total.Text) & "','" & Val(Me.txtpesototal.Text) & "','" & in_certificado & "','" & Val(Me.txtvalor_mercaderia.Text) & "','" & in_estado & "','" & Trim(Me.Txtmoneda.Text) & "'  ,'" & Val(Me.txtTc.Text) & "','" & in_tipo_consumo & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        id_transferencia = LastRegistroRUC("movimiento_transferencia", "id_transferencia")
        If Val(Me.txtid_vinculado.Text) > 0 Then
            strCadena = "UPDATE movimiento_transferencia SET id_guia_remitente='" & Val(id_transferencia) & "' WHERE id_transferencia='" & Val(Me.txtid_vinculado.Text) & "'"
            CnBd.Execute (strCadena)
        End If
        If Val(Me.txtid_venta.Text) > 0 Then
            strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & Val(Me.txtid_venta.Text) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
               rst.MoveFirst
               For i = 0 To rst.RecordCount - 1
                    strCadena = "INSERT INTO movimiento_transferencia_series(id_transferencia,id_doc,serie,numero,id_producto,chasis,motor,nro_dua,anio_fabricacion,nro_item,ruc)VALUES('" & id_transferencia & "','" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & rst("id_producto") & "','" & rst("nro_chasis") & "','" & rst("serie") & "','" & rst("nro_dua") & "','" & rst("anio_fabricacion") & "','" & rst("nro_item") & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    rst.MoveNext
               Next i
            End If
        Else
        
        strCadena = "UPDATE movimiento_transferencia_series SET id_transferencia='" & id_transferencia & "' WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' and serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' and numero='" & Trim(Me.TxtNumeroDoc.Text) & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
        Call put_transferencia_serie(id_transferencia)
        End If
        
        If Me.TxtLicencia.Text <> "" Then
            strCadena = "UPDATE persona SET licencia='" & Me.TxtLicencia.Text & "' WHERE dni='" & Me.TxtRucChofer.Text & "'"
            CnBd.Execute (strCadena)
        End If
   
    End If
    
    
    
    
    Me.cmdProcesar.Enabled = False
    Me.cmdImprimir.Enabled = True
    Me.TxtId_transferencia.Text = id_transferencia
    
    Call savedetalle_fraccion(id_transferencia, in_guia)
    
    StrNumero = FormatosCeros(Trim(str(Val(Me.TxtNumeroDoc.Text)) + 1), 6)
    strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE id_alm='" & Trim(Me.DtcAlmacenOrigen.BoundText) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' AND ruc='" & Trim(KEY_RUC) & "'"
    CnBd.Execute (strCadena)
    Me.TxtNumeroDoc.Text = StrNumero
    
    
    If KEY_CONTABILIDAD = "si" Then
            
            If Me.DtcMotivo.BoundText = 2 Then
                  '  strCadena = "call CON_InsertaAsiento_Devolucion('" & id_transferencia & "')"
                  '  CnBd.Execute (strCadena)
            End If
            
            If Me.DtcMotivo.BoundText = 4 Then
                 'strCadena = "call CON_InsertaAsiento_ConsumoInterno('" & id_transferencia & "')"
                 'CnBd.Execute (strCadena)
            End If
            
            
            If Me.DtcMotivo.BoundText = 1 Then
                If get_diferida(Val(Me.txtid_venta.Text)) = "si" Then
                    strCadena = "call CON_InsertaAsiento_GuiaDiferida('" & Val(id_transferencia) & "')"
                    CnBd.Execute (strCadena)
                End If
            End If
            
            
            End If
    End If
    
    Call impresion_formato_grupo_jm(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieGuia.BoundText), in_numero)
    
TxtId_transferencia.Text = 0
End Sub


Private Sub save_detalle_finalizado(ByVal in_transferencia As Double)
strCadena = "SELECT * FROM movimiento_transferencia_detalle WHERE id_transferencia='" & in_transferencia & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        strCadena = "UPDATE movimiento_transferencia_detalle SET recibido=recibido WHERE id_detalle='" & rst("id_detalle") & "'"
        CnBd.Execute (strCadena)
        
        strCadena = "UPDATE kardex SET ruc='0' WHERE id_movimiento='" & Val(in_transferencia) & "' and id_producto='" & rst("id_producto") & "' and  cantidad_pendiente>0 and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        
        strCadena = "call put_kardex_stock_premiun('03','" & Format(Me.DtpFechaEmision.Value, "YYYY-mm-dd") & "','" & Val(in_transferencia) & "','0009','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.TxtRucDestino.Text) & "','" & rst("id_producto") & "','" & rst("cantidad") & "','" & rst("precio_costo") & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        rst.MoveNext
    Next i
End If
End Sub
Private Sub put_transferencia_serie(ByVal in_transferencia As String)
strCadena = "SELECT * FROM movimiento_transferencia_series WHERE id_transferencia='" & Val(in_transferencia) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   rstK.MoveFirst
   For i = 0 To rstK.RecordCount - 1
        strCadena = "put_transferencia_serie('" & rstK("chasis") & "','" & rstK("id_producto") & "','" & Me.DtcAlmacenOrigen.BoundText & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        rstK.MoveNext
   Next
End If

End Sub
Private Sub put_verificar_estado()
 If Val(Me.txtid_venta.Text) > 0 Then
 strCadena = "SELECT id_producto FROM movimiento_venta_detalle WHERE id_venta='" & Val(Me.txtid_venta.Text) & "' and ruc='" & KEY_RUC & "'"
 Call ConfiguraRstT(strCadena)
 If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    For i = 0 To rstT.RecordCount - 1
     
        strCadena = "SELECT funct_entrega_producto_guia('" & Val(Me.txtid_venta.Text) & "','" & rstT("id_producto") & "')"
        Call ConfiguraRstK(strCadena)
        If rstK(0) <> 0 Then
            in_acumulado = in_acumulado + 1
        End If
     
        rstT.MoveNext
    Next i
 End If
 
            If in_acumulado = 0 Then
                in_estado = "05"
            Else
                in_estado = "06"
            End If
            Call put_tracking(Val(Me.txtid_venta.Text), in_estado, "-")
            End If
 
End Sub
Private Sub savedetalle(ByVal id_transferencia As Double, ByVal nfinalizado As String)
Dim in_acumulado As Integer
strCadena = "SELECT * FROM movimiento_transferencia_temporal WHERE (numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "')"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       in_acumulado = 0
       For i = 0 To rstT.RecordCount - 1
            If nfinalizado = "si" Then
                strCadena = "DELETE FROM movimiento_transferencia_detalle WHERE id_producto='" & rstT("id_producto") & "' and id_transferencia='" & id_transferencia & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
            End If
            
            strCadena = "INSERT INTO movimiento_transferencia_detalle(id_transferencia,id_producto,detalle,cantidad,recibido,peso,precio_venta,precio_costo,id_tipo_consumo,total,ruc) VALUES ('" & id_transferencia & "','" & rstT("id_producto") & "','" & rstT("detalle") & "','" & rstT("cantidad") & "','" & rstT("recibido") & "','" & rstT("peso") & "','" & rstT("precio_venta") & "','" & rstT("precio_costo") & "','" & rstT("id_tipo_consumo") & "','" & rstT("total") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
            
            If Me.DtcMotivo.BoundText = 2 Then
            
         
                strCadena = "call put_kardex_stock_vitekey('03','" & Format(Me.DtpFechaEmision.Value, "YYYY-mm-dd") & "','" & Val(id_transferencia) & "','0009','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.TxtRucDestino.Text) & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("precio_costo") & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
            
             If Me.DtcMotivo.BoundText = 4 Then
                strCadena = "call put_kardex_stock_vitekey('03','" & Format(Me.DtpFechaEmision.Value, "YYYY-mm-dd") & "','" & Val(id_transferencia) & "','0009','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.TxtRucDestino.Text) & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("precio_costo") & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
            
            If Me.DtcMotivo.BoundText = 3 Then
                strCadena = "call put_kardex_stock_transferencia('03','" & Format(Me.DtpFechaEmision.Value, "YYYY-mm-dd") & "','" & Val(id_transferencia) & "','0009','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.TxtRucDestino.Text) & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("precio_costo") & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
            
            
            If Me.DtcMotivo.BoundText = 1 Then
                If get_diferida(Val(Me.txtid_venta.Text)) = "si" Then
                    strCadena = "call put_kardex_stock_vitekey('03','" & Format(Me.DtpFechaEmision.Value, "YYYY-mm-dd") & "','" & Val(id_transferencia) & "','0009','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.TxtRucDestino.Text) & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "',0,'" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                End If
            End If
            
            
            
            
            rstT.MoveNext
        Next i
        
        
        If KEY_CONTROL_MERCADERIA = "si" Then
            Call put_verificar_estado
        End If
        strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' and serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' and  numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
       
    End If
End Sub
Private Sub savedetalle_fraccion(ByVal id_transferencia As Double, ByVal in_guia As String)
Dim in_acumulado As Integer
strCadena = "SELECT * FROM movimiento_transferencia_temporal WHERE  id_guia_fraccionada='" & Val(in_guia) & "' and   id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "' ORDER BY id_temporal ASC"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       in_acumulado = 0
       For i = 0 To rstT.RecordCount - 1
            If nfinalizado = "si" Then
                strCadena = "DELETE FROM movimiento_transferencia_detalle WHERE id_producto='" & rstT("id_producto") & "' and id_transferencia='" & id_transferencia & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
            End If
            strCadena = "INSERT INTO movimiento_transferencia_detalle(id_transferencia,id_producto,detalle,cantidad,recibido,peso,precio_venta,precio_costo,id_tipo_consumo,total,ruc) VALUES ('" & id_transferencia & "','" & rstT("id_producto") & "','" & rstT("detalle") & "','" & rstT("cantidad") & "','" & rstT("recibido") & "','" & rstT("peso") & "','" & rstT("precio_venta") & "','" & rstT("precio_costo") & "','" & rstT("id_tipo_consumo") & "','" & rstT("total") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
            
            
            If Me.DtcMotivo.BoundText = 2 Then
                strCadena = "call put_kardex_stock_vitekey('03','" & Format(Me.DtpFechaEmision.Value, "YYYY-mm-dd") & "','" & Val(id_transferencia) & "','0009','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.TxtRucDestino.Text) & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("precio_costo") & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
            
            If Me.DtcMotivo.BoundText = 4 Then
                strCadena = "call put_kardex_stock_vitekey('03','" & Format(Me.DtpFechaEmision.Value, "YYYY-mm-dd") & "','" & Val(id_transferencia) & "','0009','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.TxtRucDestino.Text) & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("precio_costo") & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
            
            If Me.DtcMotivo.BoundText = 3 Then
                strCadena = "call put_kardex_stock_premiun('03','" & Format(Me.DtpFechaEmision.Value, "YYYY-mm-dd") & "','" & Val(id_transferencia) & "','0009','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.TxtRucDestino.Text) & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("precio_costo") & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
            
            
           If Me.DtcMotivo.BoundText = 1 Then
                If get_diferida(Val(Me.txtid_venta.Text)) = "si" Then
                    strCadena = "call put_kardex_stock_vitekey('03','" & Format(Me.DtpFechaEmision.Value, "YYYY-mm-dd") & "','" & Val(id_transferencia) & "','0009','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.TxtRucDestino.Text) & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "',0,'" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                End If
            End If
            
            
            
            
            rstT.MoveNext
        Next i
        
        
        If KEY_CONTROL_MERCADERIA = "si" Then
            Call put_verificar_estado
        End If
        strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE id_guia_fraccionada='" & Val(in_guia) & "' and  id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' and serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' and  ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
       
    End If
End Sub


Private Sub put_stock_contable(ByVal in_producto As String)
strCadena = "SELECT * FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 And Me.DtcMotivo.BoundText = 3 Then
  '  strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
    "('" & strInventario & "','" & cod_articulo & "','" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcSerieGuia.BoundText   & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & rst("precio_compra") & "','" & KEY_FECHA & "','" & Me.DtcAlmacen.BoundText & "','" & Val(Me.TxtStock_nuevo.Text) & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
   ' CnBd.Execute (strCadena)
    
End If
End Sub

Private Sub txt_dni_atencion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = seleccionar_atencion
    FrmPersona.Show
    Exit Sub
End If
End Sub

Private Sub txt_idremitente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscarRemitente(Trim(Me.txt_idremitente.Text))
End If
End Sub

Private Sub TxtBuscaChofer_Change()

End Sub

Private Sub txtBuscarTransporte_Change()
   
   strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE nombre_completo LIKE '%" & Trim(Me.txtBuscarTransporte.Text) & "%' and id_transporte='si' and  ruc='" & KEY_RUC & "'"
   Call ConfiguraRst(strCadena)
   Call LlenaDataCombo(Me.DtcTransporte)
End Sub

Private Sub txtBuscaUbigeo_Change()
If Len(Trim(Me.txtBuscaUbigeo.Text)) >= 3 Then
    Call llenar_ubigeo(Me.HfUbigeo, Trim(Me.txtBuscaUbigeo.Text))
End If

End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Me.txtCantidad.Text) > 0 Then
        Call Resalta(Me.TxtDescripcionProducto)
    Else
        Call Resalta(Me.txtCantidad)
    End If
End If
End Sub
Public Sub llenar_ubigeo(ByVal Grilla As MSHFlexGrid, ByVal in_busqueda As String)
On Error GoTo salir
strCadena = "SELECT * FROM view_ubigeo_sunat WHERE ubigeo LIKE '%" & Replace(in_busqueda, "'", "") & "%'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
            Grilla.ColWidth(0) = 800
            Grilla.ColWidth(1) = 6700
            
        Next
        cabecera = "CODIGO" & vbTab & "UBIGEO"
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstT.MoveFirst
          For i = 0 To rstT.RecordCount - 1
            Fila = rstT("cod_ubigeo_sunat") & vbTab & rstT("ubigeo")
            Grilla.AddItem Fila
            rstT.MoveNext
        Next i
 Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rstT = Nothing
End Sub

Private Sub txtChasis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT codigo,nro_chasis,motor,id_estado FROM view_producto_serie WHERE nro_chasis LIKE '%" & Trim(Me.txtchasis.Text) & "%' and  vendido='no' and id_producto='" & Trim(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 1)) & "' and ruc='" & KEY_RUC & "'"
    Call busca_serial_caja(Me.HfChasis, Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0))
End If
End Sub

Private Sub TxtCodProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (Len(Me.TxtCodProducto.Text) = 0) Or Val(Me.TxtCodProducto.Text) = 0 Then
        Call Resalta(Me.TxtCodProducto)
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
    End If
    
 
    If Trim(Mid(Me.TxtCodProducto.Text, 1, 2)) = "00" And Len(Me.TxtCodProducto.Text) > 8 Then
       Me.txtCantidad.Text = Val(Mid(Trim(Me.TxtCodProducto.Text), 8, 4) / 1000)
       Me.TxtCodProducto.Text = Mid(Me.TxtCodProducto, 3, 5)
    End If
    
    If KEY_BARRAS = "si" Then
        strCadena = "SELECT B.id_producto,P.nombre_prod,P.precio_venta,P.peso,P.id_igv,U.abreviatura FROM producto_barras B,producto P,unidad U WHERE B.id_producto=P.id_producto AND B.ruc='" & KEY_RUC & "' " & _
        "AND P.ruc='" & KEY_RUC & "' AND B.cod_barra='" & Trim(Me.TxtCodProducto.Text) & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "'"
    Else
        Me.TxtCodProducto.Text = FormatosCeros(Me.TxtCodProducto.Text, 5)
        strCadena = "SELECT A.id_producto, P.nombre_prod,P.precio_venta,P.peso,U.abreviatura FROM almacen_producto A,producto P,unidad U WHERE  P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND A.id_producto=P.id_producto AND A.id_alm='" & KEY_ALM & "' AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND A.id_producto='" & Trim(Me.TxtCodProducto.Text) & "'"
    End If
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        cprod = rst("id_producto")
        Me.TxtDescripcionProducto.Text = rst("nombre_prod")
        Me.TxtUnidad.Text = rst("abreviatura")
        Me.txtpeso.Text = rst("peso")
        Call Resalta(Me.txtCantidad)
        Exit Sub
    End If
        
End If
End Sub

Private Sub TxtDescripcionProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtpeso)
End If
End Sub

Private Sub txtIdUbigeoDestino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    Me.frmUbigeo.Visible = True
    Me.frmUbigeo.Top = Me.txtIdUbigeoDestino.Top
   
    Call Resalta(Me.txtBuscaUbigeo)
End If
End Sub

Private Sub txtIdUbigeoOrigen_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
     Procedencia = buscar
    Me.frmUbigeo.Visible = True
    Me.frmUbigeo.Top = Me.txtIdUbigeoOrigen.Top
   
    Call Resalta(Me.txtBuscaUbigeo)
End If

End Sub

Private Sub TxtLicencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtCodProducto)
End If
End Sub

Private Sub TxtMarcayPlaca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtRucChofer)
End If
End Sub

Private Sub TxtNumero_guia_KeyPress(KeyAscii As Integer)
Dim idVenta As Double
If KeyAscii = 13 Then
    Me.txtid_venta.Text = 0
    
    
    Select Case Me.DtcTipomovimiento.BoundText
        Case "00001"
            Me.TxtNumero_guia.Text = FormatosCeros(Me.TxtNumero_guia.Text, 6)
            strCadena = "SELECT id_venta,documento,id_moneda,tc FROM movimiento_venta WHERE (numero='" & Trim(Me.TxtNumero_guia.Text) & "' AND id_doc='" & Trim(Me.DtcComprobanteGuia.BoundText) & "' AND serie='" & Trim(Me.TxtSeri_guia.Text) & "'  AND ruc='" & KEY_RUC & "') LIMIT 1"
        Case "00002"
            Me.TxtNumero_guia.Text = FormatosCeros(Me.TxtNumero_guia.Text, 8)
    
            strCadena = "SELECT id_compra as id_venta,CONCAT(serie,'-',numero) as documento,id_moneda,tc FROM movimiento_compra WHERE id_proveedor='" & Trim(Me.TxtRucDestino.Text) & "' and  numero='" & Trim(Me.TxtNumero_guia.Text) & "' AND id_doc='" & Trim(Me.DtcComprobanteGuia.BoundText) & "' AND serie='" & Trim(Me.TxtSeri_guia.Text) & "'  AND ruc='" & KEY_RUC & "' LIMIT 1"
        Case "00003"
            Me.TxtNumero_guia.Text = FormatosCeros(Me.TxtNumero_guia.Text, 6)
            strCadena = "SELECT id_transferencia as id_venta,observacion as documento,id_moneda,tc FROM movimiento_transferencia WHERE (numero='" & Trim(Me.TxtNumero_guia.Text) & "' AND id_doc='" & Trim(Me.DtcComprobanteGuia.BoundText) & "' AND serie='" & Trim(Me.TxtSeri_guia.Text) & "'  AND ruc='" & KEY_RUC & "')"
    End Select
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
         Me.HfDetalle.Rows = 0
         MsgBox "DOCUMENTO NO REGISTRADO ", vbInformation, KEY_EMPRESA
         Call Resalta(Me.TxtNumero_guia)
         Exit Sub
    Else
            idVenta = rst("id_venta")
            txtid_vinculado = ""
            Me.txtObservacion.Text = rst("documento")
            Me.txtid_venta.Text = rst("id_venta")
            If MsgBox("ESTA SEGURO DE REALIZAR ESTA OPERACION", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
                
                
                Select Case Me.DtcTipomovimiento.BoundText
                    Case "00001"
                        strCadena = "SELECT id_cliente as dni,ncliente,direccion,id_tipo_factura,'0.00' as peso_total,total,id_moneda, tc FROM movimiento_venta  WHERE id_venta='" & idVenta & "' AND ruc='" & KEY_RUC & "'"
                    Case "00002"
                        strCadena = "SELECT id_compra as id_venta,'00001' as id_tipo_factura,id_proveedor as dni,'" & Trim(Me.TxtNombreDestino.Text) & "',CONCAT(serie,'-',numero) as documento ,'0.00' as peso_total,total,id_moneda,tc FROM movimiento_compra WHERE id_proveedor='" & Trim(TxtRucDestino.Text) & "' and  numero='" & Trim(Me.TxtNumero_guia.Text) & "' AND id_doc='" & Trim(Me.DtcComprobanteGuia.BoundText) & "' AND serie='" & Trim(Me.TxtSeri_guia.Text) & "'  AND ruc='" & KEY_RUC & "' LIMIT 1"
                    Case "00003"
                         If Me.DtcComprobanteGuia.BoundText = "0031" Or Me.DtcComprobanteGuia.BoundText = "0009" Then
                         Me.txtid_vinculado.Text = rst("id_venta")
                         strCadena = "SELECT t.`id_destinatario` as dni,t.`destinatario` as nombre_completo,t.`direccion`,'00001' as id_tipo_factura,peso_total,valor_mercaderia as total,id_moneda,tc FROM `movimiento_transferencia` t  where id_transferencia='" & Val(Me.txtid_venta.Text) & "'"
                         End If
                End Select
    
                
               
                
                
                Call ConfiguraRstT(strCadena)
                If rstT.RecordCount > 0 Then
                    Call BuscarResponsable(rstT("dni"))
                    Me.txt_tipo_transferencia.Text = rstT("id_tipo_factura")
                    Me.chk_pesoglobal.Value = 1
                    Me.Txtmoneda.Text = rstT("id_moneda")
                    Me.txtTc.Text = rstT("tc")
                    Me.chk_valor_mercaderia.Value = 1
                    Me.txtvalor_mercaderia.Text = Format(rstT("total"), "###0.00")
                End If
                
                Call Llenar_Temporal_transferencias(idVenta)
                If Val(Me.txtpesototal.Text) = 0 Then
                    Me.txtpesototal.Text = rstT("peso_total")
                End If
                
                Me.TxtCodProducto.Enabled = False
                Me.TxtDescripcionProducto.Enabled = False
                
                Me.cmdAgregar.Enabled = False
                Me.CmdQuitar.Enabled = False
                Call Resalta(TxtNumero_guia)
                Referencia = True
                Me.cmdProcesar.Enabled = True
                Me.cmdImprimir.Enabled = False
                
                
                Me.TxtCodProducto.Enabled = True
                
                Me.cmdAgregar.Enabled = True
                Me.CmdQuitar.Enabled = True
                
                
            End If
    End If
End If
Set rst = Nothing

End Sub

Private Sub TxtNumeroDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscar_comprobante
End If
End Sub
Public Sub buscar_comprobante(Optional id_transferencia As Double)
    
    Dim in_manifiesto As String
    Dim in_hash As String
    in_hash = ""
    Me.TxtNumeroDoc.Text = FormatosCeros(Me.TxtNumeroDoc.Text, 6)
    If Val(id_transferencia) > 0 Then
        strCadena = "SELECT * FROM movimiento_transferencia WHERE id_transferencia='" & Val(id_transferencia) & "'"
    Else
        strCadena = "SELECT * FROM movimiento_transferencia WHERE (numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND ruc='" & KEY_RUC & "')"
    End If
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        in_manifiesto = 0
        strCadena = "SELECT * FROM movimiento_transferencia_temporal WHERE (serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
        Call Llenar_Temporal(Me.HfDetalle)
        Me.cmdProcesar.Enabled = True
        
        End If
        
    Else
        Me.frmtransporte.Visible = False
        Me.TxtId_transferencia.Text = rst("id_transferencia")
        If id_transferencia = 0 Then
            id_transferencia = rst("id_transferencia")
        End If
        Me.txtEstado.Text = rst("id_estado")
        
        in_hash = rst("sunat_key")
        
        in_manifiesto = rst("id_manifiesto")
        Me.txt_dni_atencion.Text = rst("dni_atencion")
        Me.txt_atencion.Text = rst("atencion")
        Me.DtcTipoDoc.BoundText = rst("id_doc")
        
        Me.DtpFechaEmision.Value = rst("fecha")
        Me.DtpTraslado.Value = rst("fecha_traslado")
        
        Me.DtcSerieGuia.BoundText = rst("serie")
        Me.TxtNumeroDoc.Text = rst("numero")
        Me.txt_dni_atencion.Text = rst("dni_atencion")
        Me.txt_atencion.Text = rst("atencion")
        Me.chk_pesoglobal.Value = 1
        Me.chk_valor_mercaderia.Value = 1
        Me.txtpesototal.Text = rst("peso_total")
        Me.txtvalor_mercaderia.Text = rst("valor_mercaderia")
        
        
        If rst("id_remitente") = "0" Then
            Me.txt_idremitente.Text = KEY_RUC
            Me.txtremitente.Text = KEY_EMPRESA
        Else
            Me.txt_idremitente.Text = rst("id_remitente")
            Me.txtremitente.Text = rst("remitente")
        End If
        Me.txtpesototal.Text = Format(rst("peso_total"), "###0.00")
        Me.TxtRucDestino.Text = rst("id_destinatario")
        Me.TxtNombreDestino.Text = rst("destinatario")
        
        Me.txtdireccionfiscal.Text = rst("direccion_fiscal")
        Me.txtDireccionLlegada.Text = rst("direccion_destino")
        Me.txtDireccionPartida.Text = rst("direccion")
        
        'Me.TxtDireccionDestino.Text = rst("direccion")
        Me.txt_tipo_transferencia.Text = rst("id_tipo_guia")
        
        Me.txtmtc.Text = rst("mtc")
        If rst("anulado") = "si" Then
            Me.frmanulado.Visible = True
            Me.lblAnulado.Visible = True
            Me.cmdverificar.Enabled = False
            
        Else
            Me.frmanulado.Visible = False
            Me.lblAnulado.Visible = False
           
            
        End If
        If rst("id_alm_origen") = KEY_ALM Then
            Me.cmdverificar.Enabled = True
            
        Else
           
            Me.cmdImprimir.Enabled = False
            
            Me.HfChasis.Enabled = False
            Me.HfSeries.Enabled = False
        End If
        Me.DtcAlmacenOrigen.BoundText = rst("id_alm_origen")
        Me.DtcAlmacenDestino.BoundText = rst("id_alm_destino")
        
        Me.txtObservacion.Text = rst("observacion")
        
        
       

        
        Me.txtdireccionfiscal.Text = rst("direccion_fiscal")
        Me.txtDireccionLlegada.Text = rst("direccion_destino")
        
        Me.txtIdUbigeoOrigen.Text = rst("ubigeo_origen")
        Me.txtUbigeoOrigen.Text = get_ubigeo_sunat_descripcion(Me.txtIdUbigeoOrigen.Text)
        Me.txtIdUbigeoDestino.Text = rst("ubigeo_destino")
        Me.txtUbigeoDestino.Text = get_ubigeo_sunat_descripcion(Me.txtIdUbigeoDestino.Text)
        
        
        If IsNull(rst("id_transporte")) = False And rst("id_transporte") <> "" Then
            Me.TxtRucTransporte.Text = rst("id_transporte")
            Me.lblRazonTransporte.Caption = get_persona(rst("id_transporte"))
            Me.TxtMarcayPlaca.Text = rst("marca_placa")
            Me.TxtPlaca.Text = rst("placa")
        Else
            Me.lblRazonTransporte.Caption = ""
        End If
        
        If rst("tipo_transporte") = 2 Then
           Me.Opt_transporte_privado.Value = True
        Else
           Me.Opt_transporte_publico.Value = True
        End If
        
        Me.txtmtc.Text = rst("mtc")
        Me.txtcontenedor.Text = rst("contenedor")
        Me.txtdocumentoReferencia.Text = rst("comprobante_relacionado")
        Me.TxtBultos.Text = rst("numero_bultos")
        
        
        
        If IsNull(rst("id_chofer")) = False And rst("id_chofer") <> "" Then
            Me.TxtRucChofer.Text = rst("id_chofer")
            Me.lblRazonChofer.Caption = get_persona(rst("id_chofer"))
            Me.TxtLicencia.Text = rst("licencia")
        Else
            Me.lblRazonChofer.Caption = ""
        End If
        
        Me.DtcMotivo.BoundText = rst("id_motivo")
        
        
        If rst("id_motivo") = 4 Then
           Me.txtOtros.Text = rst("motivo_otros")
        End If
        If rst("id_motivo") = 3 And rst("finalizado") = "no" Then
           Me.cmdverificar.Enabled = True
        End If
                
        
        If rst("id_manifiesto") > 0 Then
           Me.chk_Manifiesto.Value = 1
           Call load_manifiesto(in_manifiesto)
        End If
        
        Call llenar_detalle(Me.HfDetalle, id_transferencia)
        Me.DtcTipoDoc.Enabled = False
       
        Me.TxtNumeroDoc.Enabled = False
        Me.TxtCodProducto.Enabled = False
        Me.TxtDescripcionProducto.Enabled = False
        Me.cmdAgregar.Enabled = False
        Me.CmdQuitar.Enabled = False
        Me.cmdImprimir.Enabled = True
        
        Me.txtCantidad.Enabled = False
        Set rst = Nothing
        
        
        If KEY_USUARIO = "42546269" Or KEY_USUARIO = "900001" Or KEY_USUARIO = "46947665" Or KEY_USUARIO = "71574340" Then
            Me.cmdEliminar.Enabled = True
        Else
            Me.cmdEliminar.Enabled = False
        End If
        
        If get_comprobante_electronico(Me.DtcTipoDoc.BoundText, Me.DtcSerieGuia.Text) = True Then
            If Len(in_hash) < 2 Then
                Me.cmdProcesar.Enabled = True
            End If
        End If

    'Else
   ' Call Resalta(Me.TxtRuc)
    'End If
    End If
End Sub

Private Sub TxtPeso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
        Me.cmdAgregar.SetFocus
    
End If
End Sub

Private Sub TxtRucChofer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscarChofer(Me.TxtRucChofer)
End If
End Sub



Private Sub TxtRucDestino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscarResponsable(Trim(Me.TxtRucDestino.Text))
End If
End Sub

Private Sub TxtRucTransporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscarTransporte(Me.TxtRucTransporte)
End If
End Sub
Private Function get_descontar(ByVal in_producto As String, ByVal in_venta As String)
'032697
strCadena = "SELECT sum(cantidad) FROM view_entrega_producto WHERE anulado='no' and  ruc='" & KEY_RUC & "' and id_venta='" & Val(in_venta) & "' and id_producto='" & in_producto & "' "
Call ConfiguraRstP(strCadena)
If IsNull(rstP(0)) = True Then
   get_descontar = 0
Else
    get_descontar = rstP(0)
End If

End Function
Private Function get_venta_diferida(ByVal in_venta As String) As String
strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    get_venta_diferida = rstL("diferida")
Else
    get_venta_diferida = "no"
End If
End Function
Private Sub Llenar_Temporal_transferencias(ByVal idVenta As Double)

Dim total_temp As Double
Dim rstTemporal As New ADODB.Recordset
Dim rstDetalle As New ADODB.Recordset
Dim i As Integer

Select Case Me.DtcTipomovimiento.BoundText
                    Case "00001"
                        strCadena = "SELECT * FROM movimiento_venta_detalle D WHERE id_venta='" & idVenta & "' AND ruc='" & KEY_RUC & "'"
                    Case "00002"
                        strCadena = "SELECT * FROM movimiento_compra_detalle D WHERE id_compra='" & idVenta & "' AND ruc='" & KEY_RUC & "'"
                    Case "00003"
                         If Me.DtcComprobanteGuia.BoundText = "0031" Or Me.DtcComprobanteGuia.BoundText = "0009" Then
                            strCadena = "SELECT * FROM movimiento_transferencia_detalle D WHERE id_transferencia='" & idVenta & "' AND ruc='" & KEY_RUC & "'"
                         End If
                End Select





Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' "
CnBd.Execute (strCadena)

total_temp = 0
rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
    
    If KEY_CONTROL_MERCADERIA = "si" Then
        in_cantidad_exacta = rst("cantidad") - get_descontar(rst("id_producto"), idVenta)
       
    Else
        in_cantidad_exacta = rst("cantidad")
    End If
    If in_cantidad_exacta > 0 Then
        Select Case Me.DtcTipomovimiento.BoundText
                    Case "00001"
                            If get_venta_diferida(Me.txtid_venta.Text) = "si" Then
                             Me.txtdiferida.Text = "si"
                            in_cantidad_exacta = control_stock(rst("id_producto"), Val(in_cantidad_exacta))
                            If in_cantidad_exacta > 0 Then
                                strCadena = "INSERT INTO movimiento_transferencia_temporal(id_doc,serie,numero,id_producto,detalle,cantidad,peso,total,precio_venta,precio_costo,dni_save,ruc) VALUES " & _
                                "('" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & rst("id_producto") & "','" & rst("detalle") & "','" & in_cantidad_exacta & "','" & rst("peso") & "'," & _
                                "'" & Val(rst("peso")) * Val(in_cantidad_exacta) & "','" & rst("precio") & "','" & get_precio_costo(rst("id_producto")) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                            End If
                            Else
                                strCadena = "INSERT INTO movimiento_transferencia_temporal(id_doc,serie,numero,id_producto,detalle,cantidad,peso,total,precio_venta,precio_costo,dni_save,ruc) VALUES " & _
                                "('" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & rst("id_producto") & "','" & rst("detalle") & "','" & in_cantidad_exacta & "','" & rst("peso") & "'," & _
                                "'" & Val(rst("peso")) * Val(in_cantidad_exacta) & "','" & rst("precio") & "','" & get_precio_costo(rst("id_producto")) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                            End If
                    Case "00002"
                        strCadena = "INSERT INTO movimiento_transferencia_temporal(id_doc,serie,numero,id_producto,detalle,cantidad,peso,total,precio_costo,dni_save,ruc) VALUES " & _
                        "('" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & rst("id_producto") & "','" & rst("detalle") & "','" & in_cantidad_exacta & "','" & rst("peso") & "'," & _
                        "'" & Val(rst("peso")) * Val(in_cantidad_exacta) & "','" & rst("p_costo") + rst("incremento_neto_gasto") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                    Case "00003"
                         strCadena = "INSERT INTO movimiento_transferencia_temporal(id_doc,serie,numero,id_producto,detalle,cantidad,peso,total,precio_costo,dni_save,ruc) VALUES " & _
                        "('" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.DtcSerieGuia.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & rst("id_producto") & "','" & rst("detalle") & "','" & in_cantidad_exacta & "','" & rst("peso") & "'," & _
                        "'" & Val(rst("peso")) * Val(in_cantidad_exacta) & "','" & rst("precio_costo") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        End Select
                        CnBd.Execute (strCadena)
    End If
    
    rst.MoveNext
    Next i
    Call Llenar_Temporal(Me.HfDetalle)
End If

 

End Sub

Public Sub Llenar_Temporal(ByVal Grilla As MSHFlexGrid)
'On Error GoTo salir
Dim tTotal As Double, ccostos As String
Dim in_cantidad_total As Single

strCadena = "SELECT * FROM view_transferencia_temporal WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Me.DtcSerieGuia.BoundText & "' AND numero='" & Me.TxtNumeroDoc.Text & "' AND dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.cmdProcesar.Enabled = False
    Me.cmdImprimir.Enabled = False
    
    Grilla.Rows = 0
    
    Exit Sub

End If
   Me.lblCantidad.Caption = rst.RecordCount
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 7500
           Grilla.ColWidth(3) = 1100
           Grilla.ColWidth(4) = 1100
           Grilla.ColWidth(5) = 1100
           Grilla.ColWidth(6) = 1100
           Grilla.ColWidth(7) = 850
           Grilla.ColWidth(8) = 1200
           Grilla.ColWidth(9) = 1200
       Next
        cabecera = "IDDETALLE" & vbTab & "COD PROD" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "ENVIADO" & vbTab & "RECIBIDO" & vbTab & "UNIDAD" & vbTab & "MARCA" & vbTab & "PESO" & vbTab & "TIPO MOVIMIENTO" & vbTab & " TOTAL"
        Grilla.AddItem cabecera
         For k = 1 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        in_cantidad_total = 0
        For i = 0 To rst.RecordCount - 1
          tTotal = tTotal + rst("total")
          Fila = rst("id_temporal") & vbTab & rst("id_producto") & vbTab & rst("detalle") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("recibido"), "#,##0.00") & vbTab & rst("abreviatura") & vbTab & rst("descripcion") & vbTab & Format(rst("peso"), "#,##0.00") & vbTab & rst("tipo_consumo") & vbTab & Format(rst("total"), "#,##0.00")
          Grilla.AddItem Fila
          
          in_cantidad_total = in_cantidad_total + rst("cantidad")
          
          If rst("cantidad") <> rst("recibido") Then
          For k = 1 To 9
              Grilla.col = k
              Grilla.Row = i + 1
              Grilla.CellBackColor = &HC0C0FF
          Next k
          End If
          Fila = ""
          rst.MoveNext
      Next i
        
        Me.txtcantidad_total.Text = in_cantidad_total
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "**PESO TOTAL**" & vbTab & Format(tTotal, "###0.00")
        Grilla.AddItem Fila
        If Val(Me.txtpesototal.Text) = 0 Then
            Me.txtpesototal.Text = Format(tTotal, "###0.00")
        End If
        
      For k = 6 To 9
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0C0FF
      Next k
    Me.cmdProcesar.Enabled = True
    Me.cmdImprimir.Enabled = False
    
' Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Public Sub llenar_detalle(ByVal Grilla As MSHFlexGrid, ByVal id_transferencia As Double)
'On Error GoTo salir
Dim tTotal As Double, ccostos As String
strCadena = "SELECT  * FROM view_transferencia_detalle WHERE id_transferencia='" & id_transferencia & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.cmdProcesar.Enabled = False
    Me.cmdImprimir.Enabled = False
    Grilla.Rows = 0
    
    Exit Sub

End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 800
           Grilla.ColWidth(2) = 7000
           Grilla.ColWidth(3) = 1300
           Grilla.ColWidth(4) = 1300
           Grilla.ColWidth(5) = 1300
           Grilla.ColWidth(6) = 1300
           Grilla.ColWidth(7) = 1300
           Grilla.ColWidth(8) = 1300
       Next
        cabecera = "IDDETALLE" & vbTab & "COD PROD" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "ENVIADO" & vbTab & "RECIBIDO" & vbTab & "UNIDAD" & vbTab & "MARCA" & vbTab & "PESO" & vbTab & " TOTAL"
        Grilla.AddItem cabecera
         For k = 1 To 8
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
          tTotal = tTotal + rst("total")
          Fila = rst("id_detalle") & vbTab & rst("id_producto") & vbTab & rst("detalle") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("recibido"), "#,##0.00") & vbTab & rst("abreviatura") & vbTab & rst("marca") & vbTab & Format(rst("peso"), "#,##0.00") & vbTab & Format(rst("total"), "#,##0.00")
          Grilla.AddItem Fila
          If rst("cantidad") <> rst("recibido") Then
          For k = 1 To 8
              Grilla.col = k
              Grilla.Row = i + 1
              Grilla.CellBackColor = &HC0C0FF
          Next k
          End If
          Fila = ""
          rst.MoveNext
      Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "**PESO TOTAL**" & vbTab & Format(tTotal, "###0.00")
        Grilla.AddItem Fila
       
     
      'Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
      'Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
' Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Public Sub busca_seriales(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String, ByVal id_transferencia As Double)
On Error GoTo salir

If id_transferencia > 0 Then
    strCadena = "SELECT id_detalle,chasis,motor FROM movimiento_transferencia_series WHERE id_transferencia='" & id_transferencia & "' and ruc='" & KEY_RUC & "' "
Else
    strCadena = "SELECT id_detalle,chasis_motor FROM movimiento_transferencia_series WHERE id_producto='" & id_producto & "' and id_doc='" & Me.DtcTipoDoc.BoundText & "' and serie='" & Trim(Me.DtcSerieGuia.BoundText) & "' and numero='" & Trim(Me.TxtNumeroDoc.Text) & "' and ruc='" & KEY_RUC & "'"
End If


Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.cmdProcesar.Enabled = False
    Me.cmdImprimir.Enabled = False
    Grilla.Rows = 0
    
    Exit Sub

End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2500
           Grilla.ColWidth(2) = 2500
       Next
        cabecera = "CODIGO" & vbTab & "CHASIS" & vbTab & "MOTOR"
        Grilla.AddItem cabecera
         For k = 0 To 2
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
          
          Fila = rst("id_detalle") & vbTab & rst("chasis") & vbTab & rst("motor")
          Grilla.AddItem Fila
          If rst("cantidad") <> rst("recibido") Then
          For k = 0 To 2
              Grilla.col = k
              Grilla.Row = i + 1
              Grilla.CellBackColor = &HC0C0FF
          Next k
          End If
          Fila = ""
          rst.MoveNext
      Next i
        
       
     
      'Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
      'Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
 Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Public Sub llenar_series(ByVal Grilla As MSHFlexGrid, ByVal id_doc As String, ByVal serie As String, ByVal numero As String)
strCadena = "SELECT * FROM movimiento_transferencia_series WHERE id_producto='" & Trim(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 1)) & "' and  id_doc='" & id_doc & "' and serie='" & serie & "' and numero='" & numero & "' and ruc='" & KEY_RUC & "'"
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
           Grilla.ColWidth(1) = 2400
           Grilla.ColWidth(2) = 2400
           Grilla.ColWidth(3) = 400
          
        Next
        If KEY_RUC = "20480516771" Then
            cabecera = "ID" & vbTab & "SERIE" & vbTab & "CONTADOR" & vbTab & ""
        Else
            cabecera = "ID" & vbTab & "CHASIS" & vbTab & "MOTOR" & vbTab & ""
        End If
        Grilla.AddItem cabecera
         For k = 1 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        
        c = 3
        NumeroCampo = 3
            
        For i = 0 To rstT.RecordCount - 1
          
          If rstT("recibido") = "si" Then
            estado = Chr(254)
          Else
            estado = Chr(168)
          End If
          
          
          
          
          Fila = rstT("id_detalle") & vbTab & rstT("chasis") & vbTab & rstT("motor") & vbTab & estado
          Grilla.AddItem Fila
           For k = 1 To 3
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HDFDFE0
         Next k
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
          
          rstT.MoveNext
      Next i
      


End Sub

Public Sub busca_serial_caja(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
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
           Grilla.ColWidth(1) = 2200
           Grilla.ColWidth(2) = 2200
           Grilla.ColWidth(3) = 1500
           Grilla.ColWidth(4) = 400
        Next
        If KEY_RUC = "20480516771" Then
            cabecera = "ID" & vbTab & "SERIE" & vbTab & "CONTADOR" & vbTab & "ESTADO" & vbTab & ""
        Else
            cabecera = "ID" & vbTab & "CHASIS" & vbTab & "MOTOR" & vbTab & "ESTADO" & vbTab & ""
        End If
        Grilla.AddItem cabecera
         For k = 1 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        
        
            c = 4
            NumeroCampo = 4
            
        For i = 0 To rstT.RecordCount - 1
          
              estado = Chr(168)
          Select Case rstT("id_estado")
            Case "01"
                nestado = "SIN PROCESAR"
            Case "02"
                nestado = "EN PROCESO"
            Case "03"
                nestado = "TERMINADO"
          End Select
          
          Fila = Format(i + 1, "00") & vbTab & rstT("nro_chasis") & vbTab & rstT("motor") & vbTab & nestado & vbTab & estado
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
          
          rstT.MoveNext
      Next i
      


End Sub

Private Sub TxtSeri_guia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSeri_guia.Text = UCase(formato_item(Trim(Me.TxtSeri_guia.Text), 3))
    Call Resalta(Me.TxtNumero_guia)
End If

End Sub



VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmOrdenCompraDet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   16920
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtObservacion 
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
      Height          =   645
      Left            =   1920
      MaxLength       =   80
      TabIndex        =   129
      Top             =   7420
      Width           =   4575
   End
   Begin VB.TextBox TxtDescuentoParcial 
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
      Left            =   9360
      MaxLength       =   80
      TabIndex        =   124
      ToolTipText     =   "Descuento"
      Top             =   7035
      Width           =   1215
   End
   Begin VB.Frame frmGastos 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "FLETE VINCULADO"
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
      Height          =   6495
      Left            =   0
      TabIndex        =   76
      Top             =   480
      Visible         =   0   'False
      Width           =   12615
      Begin VitekeySoft.ChameleonBtn cmdDeleteGasto 
         Height          =   975
         Left            =   11640
         TabIndex        =   125
         Top             =   5280
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1720
         BTYPE           =   3
         TX              =   "DELETE"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   4194304
         FCOLO           =   4194304
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmOrdenPagoDet.frx":0000
         PICN            =   "FrmOrdenPagoDet.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame frm_domiciliada 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NO DOMICILIADA"
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
         Height          =   1815
         Left            =   8880
         TabIndex        =   77
         Top             =   840
         Visible         =   0   'False
         Width           =   2775
         Begin VB.TextBox txtSubtotal 
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
            TabIndex        =   80
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtRetencion 
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
            TabIndex        =   79
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text1 
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
            TabIndex        =   78
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SUB TOTAL :"
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
            Left            =   180
            TabIndex        =   83
            Top             =   480
            Width           =   795
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RETENCION :"
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
            Left            =   135
            TabIndex        =   82
            Top             =   840
            Width           =   840
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL :"
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
            Left            =   480
            TabIndex        =   81
            Top             =   1320
            Width           =   495
         End
      End
      Begin VB.TextBox TxtFleteNumero 
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
         Left            =   6960
         TabIndex        =   98
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtdescripcion 
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
         Height          =   525
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   97
         Top             =   4080
         Width           =   6735
      End
      Begin VB.TextBox TxtfleteSerie 
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
         Left            =   5880
         TabIndex        =   96
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox txtmonto 
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
         Left            =   1800
         TabIndex        =   94
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Txtdni 
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
         Left            =   1800
         TabIndex        =   93
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtTipoCambio 
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
         Left            =   3960
         TabIndex        =   92
         Top             =   1440
         Width           =   1000
      End
      Begin VB.TextBox txtcodigoprod 
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
         Left            =   1800
         TabIndex        =   91
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txtproducto 
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
         Left            =   3480
         TabIndex        =   90
         Top             =   3360
         Width           =   5055
      End
      Begin VB.Frame frmRetencion 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   280
         Left            =   8640
         TabIndex        =   87
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
         Begin VB.CheckBox chk_suspencion_retencion 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "SUSPENCION RETENCION"
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
            Height          =   255
            Left            =   0
            TabIndex        =   88
            Top             =   10
            Width           =   2415
         End
      End
      Begin VB.CheckBox chkresponsable 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "RESPONSABLE :"
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
         Height          =   320
         Left            =   3600
         TabIndex        =   86
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtBuscarresponsable 
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
         MaxLength       =   80
         TabIndex        =   85
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox chk_nodomiciliada 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "NO DOMICILIADA"
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
         Height          =   255
         Left            =   8640
         TabIndex        =   84
         Top             =   540
         Width           =   2655
      End
      Begin VitekeySoft.ChameleonBtn cmdProcesarFlete 
         Height          =   465
         Left            =   1800
         TabIndex        =   89
         Top             =   4680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "AGREGAR FLETE"
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
         COLTYPE         =   3
         FOCUSR          =   -1  'True
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmOrdenPagoDet.frx":2466
         PICN            =   "FrmOrdenPagoDet.frx":2482
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpfecha 
         Height          =   325
         Left            =   1800
         TabIndex        =   95
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53477377
         CurrentDate     =   41126
      End
      Begin MSDataListLib.DataCombo DtcComprobante 
         Height          =   315
         Left            =   1800
         TabIndex        =   99
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
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
      Begin MSComCtl2.DTPicker dtpvencimientoFlete 
         Height          =   330
         Left            =   4200
         TabIndex        =   100
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53477377
         CurrentDate     =   41126
      End
      Begin MSDataListLib.DataCombo DtcAfectoIgv 
         Height          =   315
         Left            =   1800
         TabIndex        =   101
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         ListField       =   ""
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
      Begin MSDataListLib.DataCombo DtcMonedaFlete 
         Height          =   315
         Left            =   1800
         TabIndex        =   102
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSDataListLib.DataCombo DtcTipoCompra 
         Height          =   315
         Left            =   1800
         TabIndex        =   103
         Top             =   2895
         Width           =   2535
         _ExtentX        =   4471
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
      Begin MSDataListLib.DataCombo DtcPeriodo 
         Height          =   315
         Left            =   1800
         TabIndex        =   104
         Top             =   2520
         Width           =   2535
         _ExtentX        =   4471
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
      Begin MSDataListLib.DataCombo DtcResponsable 
         Height          =   330
         Left            =   4980
         TabIndex        =   105
         Top             =   1080
         Width           =   3135
         _ExtentX        =   5530
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshGastos 
         Height          =   975
         Left            =   240
         TabIndex        =   123
         Top             =   5280
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   1720
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
      Begin VB.Image cerrar_flete 
         Height          =   240
         Left            =   12120
         Picture         =   "FrmOrdenPagoDet.frx":5ACA
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INGRESO DE FLETE OBLIGATORIO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   6000
         TabIndex        =   122
         Top             =   1680
         Width           =   4335
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOCUMENTO :"
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
         Left            =   255
         TabIndex        =   121
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONEDA :"
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
         Left            =   540
         TabIndex        =   120
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO  :"
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
         Left            =   600
         TabIndex        =   119
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUC/DNI  :"
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
         Left            =   525
         TabIndex        =   118
         Top             =   720
         Width           =   705
      End
      Begin VB.Label lblcliente 
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
         Left            =   3600
         TabIndex        =   117
         Top             =   720
         Width           =   4875
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA  :"
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
         Left            =   690
         TabIndex        =   116
         Top             =   2200
         Width           =   540
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AFECTO IGV :"
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
         Left            =   375
         TabIndex        =   115
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "T.C :"
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
         Left            =   3600
         TabIndex        =   114
         Top             =   1440
         Width           =   285
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "VENCIMI:"
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
         Left            =   3480
         TabIndex        =   113
         Top             =   2200
         Width           =   630
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   540
         TabIndex        =   112
         Top             =   2640
         Width           =   690
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCION :"
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
         TabIndex        =   111
         Top             =   4200
         Width           =   990
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO GASTO :"
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
         Left            =   360
         TabIndex        =   110
         Top             =   3000
         Width           =   870
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SERVICIO :"
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
         Left            =   540
         TabIndex        =   109
         Top             =   3345
         Width           =   690
      End
      Begin VB.Label lblcuenta_contable 
         BackColor       =   &H008080FF&
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
         Height          =   255
         Left            =   3480
         TabIndex        =   108
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lblcuenta_detalle 
         BackColor       =   &H008080FF&
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
         Height          =   255
         Left            =   4560
         TabIndex        =   107
         Top             =   3720
         Width           =   3975
      End
      Begin VB.Label lblidCompra 
         BackColor       =   &H008080FF&
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
         Height          =   495
         Left            =   3720
         TabIndex        =   106
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   6255
         Left            =   120
         Top             =   120
         Width           =   12375
      End
   End
   Begin VB.TextBox txtMontoFlete 
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
      Left            =   12120
      TabIndex        =   67
      Top             =   8040
      Width           =   975
   End
   Begin VB.TextBox txtId_recepcion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8160
      TabIndex        =   63
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtRecibido 
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
      Height          =   300
      Left            =   6520
      TabIndex        =   62
      Top             =   8880
      Width           =   975
   End
   Begin VB.TextBox txtTc 
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
      Left            =   6600
      TabIndex        =   60
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame frmrecepcion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1905
      Left            =   9960
      TabIndex        =   50
      Top             =   1060
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox txtIdfactura_flete 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   71
         Top             =   1245
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox Txtfactura_serie 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1920
         TabIndex        =   70
         Top             =   760
         Width           =   615
      End
      Begin VB.TextBox txtGuia_serie 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1920
         TabIndex        =   69
         Top             =   400
         Width           =   615
      End
      Begin VB.TextBox txtFacturaFlete 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1920
         TabIndex        =   65
         Top             =   1485
         Width           =   1740
      End
      Begin VB.TextBox TxtOrden_serie 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   1920
         TabIndex        =   59
         Top             =   40
         Width           =   615
      End
      Begin VitekeySoft.ChameleonBtn cmdExtraerOrden 
         Height          =   345
         Left            =   3720
         TabIndex        =   58
         Top             =   45
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "EXTRAER"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmOrdenPagoDet.frx":896E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtFactura_numero 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2570
         TabIndex        =   55
         Top             =   760
         Width           =   1095
      End
      Begin VB.TextBox TxtGuia_numero 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2570
         TabIndex        =   54
         Top             =   400
         Width           =   1095
      End
      Begin VB.TextBox TxtOrden_numero 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2570
         TabIndex        =   53
         Top             =   40
         Width           =   1095
      End
      Begin VitekeySoft.ChameleonBtn cmdIngresarFactura 
         Height          =   345
         Left            =   4680
         TabIndex        =   66
         Top             =   1455
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "INGRESAR FLETE"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmOrdenPagoDet.frx":898A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpEmision 
         Height          =   315
         Left            =   5160
         TabIndex        =   73
         Top             =   720
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   53477377
         CurrentDate     =   40974
      End
      Begin MSComCtl2.DTPicker DtpVencimiento 
         Height          =   315
         Left            =   5160
         TabIndex        =   75
         Top             =   1080
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   53477377
         CurrentDate     =   40974
      End
      Begin MSDataListLib.DataCombo DtcPeriodoCompra 
         Height          =   345
         Left            =   1920
         TabIndex        =   127
         Top             =   1110
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1095
         TabIndex        =   126
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENCE :"
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
         Left            =   4635
         TabIndex        =   74
         Top             =   1200
         Width           =   525
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Left            =   4500
         TabIndex        =   72
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FACTURA FLETE :"
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
         Left            =   675
         TabIndex        =   64
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FACTURA :"
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
         Left            =   1095
         TabIndex        =   61
         Top             =   825
         Width           =   705
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GUIA REMISION :"
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
         Left            =   645
         TabIndex        =   52
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ORDEN DE COMPRA :"
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
         Left            =   375
         TabIndex        =   51
         Top             =   105
         Width           =   1425
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   1905
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   6735
      End
   End
   Begin VB.TextBox TxtTotal 
      Appearance      =   0  'Flat
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
      Left            =   9360
      TabIndex        =   48
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox txtImporteBruto 
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
      Left            =   9360
      TabIndex        =   46
      Top             =   7605
      Width           =   1455
   End
   Begin VB.TextBox TxtIgv 
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
      Left            =   9360
      TabIndex        =   45
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox txtValorVenta 
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
      Left            =   9360
      TabIndex        =   44
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox txtDescuento 
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
      Left            =   12120
      TabIndex        =   43
      Top             =   7680
      Width           =   975
   End
   Begin VB.CheckBox chk_igv 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "AFECTO A IGV"
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
      Height          =   280
      Left            =   7560
      TabIndex        =   34
      Top             =   2175
      Width           =   1335
   End
   Begin VB.TextBox txtAutorizado 
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
      Height          =   300
      Left            =   6520
      TabIndex        =   33
      Top             =   8475
      Width           =   975
   End
   Begin VitekeySoft.ChameleonBtn CmdAgregar 
      Height          =   315
      Left            =   13320
      TabIndex        =   29
      Top             =   7035
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "ADD"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmOrdenPagoDet.frx":89A6
      PICN            =   "FrmOrdenPagoDet.frx":89C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Left            =   240
      MaxLength       =   80
      TabIndex        =   26
      Top             =   7035
      Width           =   1215
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
      Left            =   1500
      MaxLength       =   80
      TabIndex        =   25
      Top             =   7035
      Width           =   975
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
      Left            =   2550
      MaxLength       =   80
      TabIndex        =   24
      Top             =   7035
      Width           =   6735
   End
   Begin VB.TextBox TxtUnidad 
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
      Left            =   10740
      MaxLength       =   80
      TabIndex        =   23
      Top             =   7035
      Width           =   1215
   End
   Begin VB.TextBox txtcosto 
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
      Left            =   12105
      MaxLength       =   80
      TabIndex        =   22
      Top             =   7035
      Width           =   975
   End
   Begin VB.TextBox TxtId_orden 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8160
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxtProveedor 
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
      Height          =   285
      Left            =   2160
      TabIndex        =   16
      Top             =   1440
      Width           =   5295
   End
   Begin MSDataListLib.DataCombo DtcTerminosEntrega 
      Height          =   330
      Left            =   2160
      TabIndex        =   15
      Top             =   2175
      Width           =   5295
      _ExtentX        =   9340
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
   Begin VB.TextBox TxtRuc 
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
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   1080
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DtpPedido 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
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
      Format          =   53477377
      CurrentDate     =   40974
   End
   Begin VB.TextBox TxtNumero 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   14760
      MaxLength       =   50
      TabIndex        =   1
      Top             =   520
      Width           =   1455
   End
   Begin VB.TextBox TxtSerie 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   13800
      MaxLength       =   50
      TabIndex        =   0
      Top             =   520
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   3975
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   7011
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
   Begin MSComCtl2.DTPicker DtpPago 
      Height          =   315
      Left            =   5880
      TabIndex        =   14
      Top             =   1800
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
      Format          =   53477377
      CurrentDate     =   40974
   End
   Begin VitekeySoft.ChameleonBtn cmdImprimir 
      Height          =   900
      Left            =   15080
      TabIndex        =   19
      Top             =   8265
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1588
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
      MICON           =   "FrmOrdenPagoDet.frx":AEED
      PICN            =   "FrmOrdenPagoDet.frx":AF09
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdProcesar 
      Height          =   900
      Left            =   14200
      TabIndex        =   20
      Top             =   8280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1588
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
      MICON           =   "FrmOrdenPagoDet.frx":D4DA
      PICN            =   "FrmOrdenPagoDet.frx":D4F6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrarpantalla 
      Height          =   900
      Left            =   15960
      TabIndex        =   21
      Top             =   8250
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1588
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
      MICON           =   "FrmOrdenPagoDet.frx":10B3E
      PICN            =   "FrmOrdenPagoDet.frx":10B5A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn CmdQuitar 
      Height          =   315
      Left            =   15000
      TabIndex        =   28
      Top             =   7035
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "DELL"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmOrdenPagoDet.frx":13B81
      PICN            =   "FrmOrdenPagoDet.frx":13B9D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcCreador 
      Height          =   330
      Left            =   1920
      TabIndex        =   30
      Top             =   8100
      Width           =   4575
      _ExtentX        =   8070
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
   Begin MSDataListLib.DataCombo DtcAutorizado 
      Height          =   330
      Left            =   1920
      TabIndex        =   31
      Top             =   8475
      Width           =   4575
      _ExtentX        =   8070
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
   Begin MSDataListLib.DataCombo DtcRecibido 
      Height          =   330
      Left            =   1920
      TabIndex        =   32
      Top             =   8850
      Width           =   4575
      _ExtentX        =   8070
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
   Begin VitekeySoft.ChameleonBtn cmdUpdate 
      Height          =   315
      Left            =   14160
      TabIndex        =   35
      Top             =   7035
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "UPD"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmOrdenPagoDet.frx":14137
      PICN            =   "FrmOrdenPagoDet.frx":14153
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   900
      Left            =   13320
      TabIndex        =   37
      Top             =   8280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1588
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
      MICON           =   "FrmOrdenPagoDet.frx":164A7
      PICN            =   "FrmOrdenPagoDet.frx":164C3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcComrpobante 
      Height          =   345
      Left            =   10320
      TabIndex        =   38
      Top             =   520
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   345
      Left            =   10320
      TabIndex        =   49
      Top             =   140
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcMoneda 
      Height          =   330
      Left            =   4680
      TabIndex        =   56
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVACION :"
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
      Left            =   570
      TabIndex        =   128
      Top             =   7680
      Width           =   1065
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FLETE :"
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
      Left            =   11580
      TabIndex        =   68
      Top             =   8040
      Width           =   465
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONEDA :"
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
      Left            =   3945
      TabIndex        =   57
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL :"
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
      Left            =   8790
      TabIndex        =   47
      Top             =   8760
      Width           =   525
   End
   Begin VB.Label lblruc 
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA EMISION :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   960
      TabIndex        =   17
      Top             =   500
      Width           =   7665
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR VENTA :"
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
      Left            =   8265
      TabIndex        =   42
      Top             =   8040
      Width           =   1005
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV :"
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
      Left            =   8925
      TabIndex        =   41
      Top             =   8400
      Width           =   345
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCUENTO :"
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
      Left            =   11115
      TabIndex        =   40
      Top             =   7680
      Width           =   915
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTE BRUTO :"
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
      Left            =   8055
      TabIndex        =   39
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   1770
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   7395
      Width           =   5415
   End
   Begin VB.Label lblid_detalle 
      Height          =   255
      Left            =   14160
      TabIndex        =   36
      Top             =   7680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LblCantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   15765
      TabIndex        =   27
      Top             =   7035
      Width           =   945
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR :"
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
      Left            =   360
      TabIndex        =   13
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RECIBIDO POR :"
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
      Left            =   615
      TabIndex        =   10
      Top             =   8850
      Width           =   1065
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AUTORIZODO POR :"
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
      Left            =   345
      TabIndex        =   9
      Top             =   8475
      Width           =   1305
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ELABORADO POR :"
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
      Left            =   435
      TabIndex        =   8
      Top             =   8100
      Width           =   1245
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TERMINOS DE ENTREGA:"
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
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA PAGO :"
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
      Left            =   4155
      TabIndex        =   6
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label lblempresa 
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA EMISION :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   960
      TabIndex        =   5
      Top             =   200
      Width           =   7665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/PROVEEDOR:"
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
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA PEDIDO:"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   915
      Left            =   9960
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   6735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   920
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   80
      Width           =   8775
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   16920
   End
End
Attribute VB_Name = "FrmOrdenCompraDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Function validar_datos() As Boolean
If Trim(Me.Txtfactura_serie.Text) = "" Or Trim(Me.txtFactura_numero.Text) = "" Then
    MsgBox "Ingrese una factura de REFERENCIA.", vbInformation, KEY_VENDEDOR
    validar_datos = False
End If

If Trim(Me.txtGuia_serie.Text) = "" Or Trim(Me.TxtGuia_numero.Text) = "" Then
    MsgBox "Ingrese una factura GUIA DE RECEPCION.", vbInformation, KEY_VENDEDOR
    validar_datos = False
End If

If Trim(Me.TxtOrden_serie.Text) = "" Or Trim(Me.TxtOrden_numero.Text) = "" Then
    MsgBox "Ingrese una ORDEN DE COMPRA VINCULADA.", vbInformation, KEY_VENDEDOR
    validar_datos = False
End If


End Function
Private Sub put_asiento_contable(ByVal in_compra As Double)
strCadena = "SELECT * FROM orden_compra WHERE id_recepcion='" & Val(Me.txtId_recepcion.Text) & "' AND id_estado<>'3' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount = 1 Then
If KEY_CONTABILIDAD = "si" Then
             strCadena = "call p_insert_compra_emitido_ii('" & in_compra & "')"
             CnBd.Execute (strCadena)
             MsgBox "NUMERO DE VOUCHER GENERADO  : " & Trim(in_compra), vbInformation, KEY_VENDEDOR
End If
End If
End Sub

Private Function validar_periodo_recepcion() As Boolean
validar_periodo_recepcion = True

If get_cierre_periodo(Me.DtcPeriodoCompra.BoundText) = False Then
            MsgBox "EL PERIODO DE EMISION NO ESTA CREADO" + Chr(13) + "COORDINE CON EL AREA CONTABLE", vbInformation, KEY_VENDEDOR
            
             validar_periodo_recepcion = False
                        Me.cmdProcesar.Enabled = True
           Exit Function
       End If
       '******************************
       
       in_periodo_factura = Me.DtcPeriodoCompra.BoundText   'get_periodo_actual(Me.DtpEmision.Value)
       
       in_periodo_recepcion = get_periodo_actual(Me.DtpPedido.Value)
       
       
       
       
       
       If get_cierre_periodo(in_periodo_recepcion) = False Then
        
           MsgBox "EL PERIODO DE RECEPCION ESTA CERRADO" + Chr(13) + "COORDINE CON EL AREA CONTABLE", vbInformation, KEY_VENDEDOR
          
           validar_periodo_recepcion = False
                       Me.cmdProcesar.Enabled = True
           Exit Function

       End If
End Function
Private Sub Save()

Dim in_orden As Double
Dim in_finalizado As String
Dim in_finalizado_confirmar As String

If Trim(Me.txtRuc.Text) = "" Then
   MsgBox "INGRESE UN PROVEEDOR PARA SU ORDEN", vbInformation, KEY_VENDEDOR
   Exit Sub
End If
Me.cmdProcesar.Enabled = False


'***** INGRESAR CABECERA Y DETALLE DE LA COMPRA KARDEX NO*********

If Me.DtcComrpobante.BoundText = "0414" Then
    
    If validar_periodo_recepcion = False Then
        Exit Sub
    End If
    
    
    in_compra = put_compra

'  If in_compra > 0 Then
'  strCadena = "select * from kardex where id_serie='" & Trim(Me.Txtfactura_serie.Text) & "' and id_numero='" & Trim(Me.txtFactura_numero.Text) & "' and cantidad_real>0 and  id_movimiento='" & Val(in_compra) & "' and ruc='" & KEY_RUC & "'"
'  Call ConfiguraRstL(strCadena)
'   If rstL.RecordCount > 0 Then
'       rstL.MoveFirst
'       For i = 0 To rstL.RecordCount - 1
'          strCadena = "DELETE FROM kardex where id_kardex='" & rstL("id_kardex") & "'"
'            CnBd.Execute (strCadena)
'            rstL.MoveNext
'        Next i
'    End If
'   End If
Else
     in_compra = 0
End If



If Trim(Me.txtRuc.Text) = "" Then
   MsgBox "INGRESE UN PROVEEDOR PARA SU ORDEN", vbInformation, KEY_VENDEDOR
   Exit Sub
End If
If Me.chk_igv.Value = 1 Then
    in_igv = "si"
Else
    in_igv = "no"
End If



strCadena = "call put_orden_compra_vitekey('" & Me.DtcComrpobante.BoundText & "','" & Trim(Me.txtSerie.Text) & "','" & get_nueva_orden(DtcComrpobante.BoundText, "") & "','" & Trim(Me.txtRuc.Text) & "','" & KEY_FECHA & "', " & _
" '" & Format(Me.DtpPedido.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpPago.Value, "YYYY-mm-dd") & "','" & Me.DtcTerminosEntrega.BoundText & "','" & KEY_USUARIO & "'," & _
" '" & Me.DtcAutorizado.BoundText & "','" & in_igv & "','" & KEY_USUARIO & "','" & KEY_ALM & "','" & Val(Me.TxtDescuento.Text) & "','" & Val(Me.TxtTotal.Text) & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtTc.Text) & "','" & in_compra & "','" & Trim(Me.txtObservacion.Text) & "','" & KEY_RUC & "')"
Call ConfiguraRstP(strCadena)
in_orden = rstP("in_orden")
Me.TxtId_orden.Text = in_orden


strCadena = "UPDATE orden_compra SET  id_recepcion='" & Val(Me.txtId_recepcion.Text) & "' WHERE id_orden='" & Val(Me.TxtId_orden.Text) & "'"
CnBd.Execute (strCadena)


strCadena = "SELECT * FROM orden_compra_detalle_temp WHERE id_doc='" & Me.DtcComrpobante.BoundText & "'and id_alm='" & Me.DtcAlmacen.BoundText & "' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   in_finalizado = "si"
   For i = 0 To rst.RecordCount - 1
        strCadena = "INSERT INTO orden_compra_detalle(`id_orden`,`id_producto`,`cantidad`,`precio`,incremento_neto,`total`,`ruc`)VALUES('" & in_orden & "','" & rst("id_producto") & "','" & rst("cantidad") & "','" & rst("precio") & "','" & rst("incremento_neto") & "','" & rst("total") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        If rst("cantidad") <> rst("cantidad_pendiente") Then
           in_finalizado = "no"
        End If
        rst.MoveNext
   Next i
End If


If Me.frmrecepcion.Visible = True Then
    
    
    If in_finalizado = "si" Then
       If verificar_finalizado(Me.txtId_recepcion.Text) = True Then
            in_estado = 2
       Else
            in_estado = 4
       End If
    Else
       in_estado = 4
    End If
    
    in_factura_compra = Trim(Me.Txtfactura_serie.Text) & "-" & Trim(Me.txtFactura_numero.Text)
    
    strCadena = "UPDATE orden_compra SET id_factura_flete='" & Val(Me.txtIdfactura_flete.Text) & "', id_recepcion='" & Val(Me.txtId_recepcion.Text) & "',guia_serie='" & Trim(Me.txtGuia_serie.Text) & "',guia_numero='" & Trim(Me.TxtGuia_numero.Text) & "', guia_remision='" & Trim(Me.TxtGuia_numero.Text) & "',id_compra='" & in_compra & "',factura_compra='" & in_factura_compra & "',dni_recepcion='" & Me.DtcRecibido.BoundText & "',id_estado='" & in_estado & "' WHERE id_orden='" & Val(Me.TxtId_orden.Text) & "'"
    CnBd.Execute (strCadena)
    
    strCadena = "UPDATE movimiento_compra_gasto SET id_recepcion='" & Val(Me.TxtId_orden.Text) & "' WHERE id_compra_gasto='" & Val(Me.txtIdfactura_flete.Text) & "' and ruc='" & KEY_RUC & "' "
    CnBd.Execute (strCadena)
    Call put_factura_flete(Me.TxtId_orden.Text)
    
    Call actualizar_pendiente_orden(in_orden, Val(Me.txtId_recepcion.Text))
    Call put_asiento_contable(in_compra)
    Call verifica_unica_recepcion(Me.txtId_recepcion.Text, Val(Me.TxtId_orden.Text), in_estado)
    
    If Val(Me.txtIdfactura_flete.Text) > 0 Then
        strCadena = "call p_insert_compra_emitido_ii('" & Val(Me.txtIdfactura_flete.Text) & "')"
        CnBd.Execute (strCadena)
    End If
    
 
    
End If
     


strCadena = "DELETE FROM orden_compra_detalle_temp WHERE id_doc='" & Me.DtcComrpobante.BoundText & "' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

If Me.frmrecepcion.Visible = True Then
    Call actualizar_kardex(in_orden, in_estado)
End If
Me.cmdProcesar.Enabled = False
Me.cmdImprimir.Enabled = True

End Sub
Private Sub verifica_unica_recepcion(ByVal in_orden_compra As String, ByVal in_recepcion As String, ByVal in_estado As String)

If in_estado = "2" Then
strCadena = "SELECT * FROM orden_compra WHERE id_recepcion='" & Val(in_orden_compra) & "' and id_estado='2' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstP(strCadena)
    If rstP.RecordCount = 1 Then
        If KEY_CONTABILIDAD = "si" Then
        strCadena = "call CON_InsertaAsiento_Recepcion('" & in_recepcion & "')"
        CnBd.Execute (strCadena)
        End If
    End If
Else
    If in_estado = "4" And verificar_periodo_recepcion(in_recepcion) = True Then
        strCadena = "call CON_InsertaAsiento_Recepcion('" & in_recepcion & "')"
        CnBd.Execute (strCadena)
    End If
End If
End Sub
Private Function verificar_periodo_recepcion(ByVal in_recepcion As String) As Boolean
Dim in_fecha_recepcion As Date
Dim in_fecha_compra As Date
strCadena = "SELECT * FROM orden_compra WHERE id_orden='" & in_recepcion & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
    in_fecha_recepcion = rstP("fecha_solicitud")
   If rstP("id_compra") > 0 Then
        strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & rstP("id_compra") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
           in_fecha_compra = rstL("fecha_emision")
        End If
        
        If Month(in_fecha_recepcion) <> Month(in_fecha_compra) Then
            verificar_periodo_recepcion = True
        Else
            verificar_periodo_recepcion = False
        End If
        
        
   End If
End If


End Function




Private Sub put_factura_flete(ByVal in_recepcion As String)
Dim in_total_flete As Single

strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & Val(Me.txtIdfactura_flete.Text) & "'"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   in_total_flete = rstP("total")
Else
   in_total_flete = 0
End If


strCadena = "UPDATE orden_compra SET monto_flete='" & in_total_flete & "' WHERE id_orden='" & in_recepcion & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "UPDATE movimiento_compra_gasto SET id_recepcion='" & Val(in_recepcion) & "' WHERE id_orden_compra='" & Val(Me.txtId_recepcion.Text) & "' and  id_recepcion='0' and  id_compra='" & Val(Me.txtIdfactura_flete.Text) & "'"
CnBd.Execute (strCadena)




End Sub

Private Function actualizar_pendiente_orden(ByVal in_recepcion As String, ByVal in_orden_compra As String)

strCadena = "SELECT * FROM orden_compra WHERE id_recepcion='" & Val(in_orden_compra) & "' and id_estado<>'03' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount = 1 Then
      strCadena = "SELECT * FROM orden_compra_detalle WHERE id_orden='" & Val(in_orden_compra) & "' ORDER BY id_detalle ASC"
      Call ConfiguraRstL(strCadena)
      If rstL.RecordCount > 0 Then
         rstL.MoveFirst
         For i = 0 To rstL.RecordCount - 1
              in_pendiente = rstL("cantidad") - get_recepcionado(rstL("id_producto"), Val(in_orden_compra))
              strCadena = "UPDATE orden_compra_detalle SET cantidad_pendiente='" & in_pendiente & "' WHERE id_detalle='" & rstL("id_detalle") & "'"
              CnBd.Execute (strCadena)
              rstL.MoveNext

         Next i
      End If
Else
   ' strCadena = "call CON_InsertaAsiento_Recepcion('" & in_recepcion & "')"
   ' CnBd.Execute (strCadena)
  End If
End Function
Private Sub actualizar_kardex(ByVal in_recepcion As String, ByVal in_estado As String)
    Dim in_costo_igv As Single
    Dim in_afecto_igv As String
    Dim in_moneda As String
    Dim in_factor As Single
    
   'OBTENGO LA ORDEN DE COMPRA
   strCadena = "SELECT * FROM orden_compra WHERE id_recepcion='" & Val(Me.txtId_recepcion.Text) & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRstL(strCadena)
   If rstL.RecordCount > 0 Then
        'ACTUALIZAR KARDEX CON GUIA
        in_moneda = rstL("id_moneda")
        in_afecto_igv = rstL("afecto_igv")
        
        'If rstL.RecordCount = 1 And in_estado = 2 Then
            ' aqui ya se agrego con la factura
        'Else
            strCadena = "SELECT * FROM orden_compra_detalle WHERE id_orden='" & Val(in_recepcion) & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRstL(strCadena)
            If rstL.RecordCount > 0 Then
               rstL.MoveFirst
               If in_moneda = "00001" Then
                   in_factor = 1
               Else
                   in_factor = Val(Me.txtTc.Text)
               End If
               For i = 0 To rstL.RecordCount - 1
                    
                    in_monto_neto = (rstL("precio") * in_factor + rstL("incremento_neto"))
                    
                    
                    If in_afecto_igv = "si" Then
                        in_costo_igv = rstL("precio") * in_factor + rstL("precio") * KEY_IGV * in_factor + rstL("incremento_neto")
                    Else
                        in_costo_igv = in_monto_neto
                    End If
                    
                    
                    If Trim(Me.txtGuia_serie.Text) = "" And Trim(Me.TxtGuia_numero.Text) = "" Then
                        strCadena = "call put_kardex_stock_16('04','" & Format(Me.DtpPedido.Value, "YYYY-mm-dd") & "','" & Val(in_recepcion) & "','0001','" & Trim(Me.Txtfactura_serie.Text) & "','" & Trim(Me.txtFactura_numero.Text) & "','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.TxtProveedor.Text) & "','" & rstL("id_producto") & "','" & rstL("cantidad") & "','" & in_costo_igv & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','no','" & KEY_RUC & "')"
                    Else
                        strCadena = "call put_kardex_stock_v16('04','" & Format(Me.DtpPedido.Value, "YYYY-mm-dd") & "','" & Val(in_recepcion) & "','0009','" & Trim(Me.txtGuia_serie.Text) & "','" & Trim(Me.TxtGuia_numero.Text) & "','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.TxtProveedor.Text) & "','" & rstL("id_producto") & "','" & rstL("cantidad") & "','" & in_costo_igv & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','no','" & KEY_RUC & "')"
                    End If
                    
                   
                    
                    CnBd.Execute (strCadena)
                    rstL.MoveNext
               Next i
            End If
        End If
   ' End If
End Sub



Private Sub save_detalle_factura(ByVal in_compra As String)
Dim in_afecto_igv As String
strCadena = "SELECT * FROM orden_compra_detalle WHERE id_producto<>'' and  id_orden='" & Val(Me.txtId_recepcion.Text) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
   rstT.MoveFirst
   If Me.chk_igv.Value = 1 Then
      in_afecto_igv = "si"
   Else
      in_afecto_igv = "no"
   End If
   
   For i = 0 To rstT.RecordCount - 1
   
        in_precio_costo = 0
        If in_afecto_igv = "si" Then
            in_igv = rstT("total") * (KEY_IGV + 1) - rstT("total")
            in_exonerado = 0
            in_valor_venta = rstT("total")
            in_total_parcial = in_igv + in_valor_venta
        Else
             in_igv = 0
             in_exonerado = rstT("total")
             in_valor_venta = 0
             in_total_parcial = in_exonerado
        End If
        
        If rstT("cantidad") <> 0 Then
            in_precio_costo = in_total_parcial / rstT("cantidad")
        Else
            in_precio_costo = rst("precio")
        End If
        
        strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,valor_neto,isc,igv,valor_venta,exonerado,total," & _
        "p_venta,p_costo,id_alm,detalle,id_detalle_orden,incremento_neto_gasto,ruc) VALUES " & _
        "('" & in_compra & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & in_precio_costo & "'," & _
        "'" & rstT("total") & "','0','" & in_igv & "', '" & in_valor_venta & "','" & in_exonerado & "','" & in_total_parcial & "', " & _
        "'" & get_precio_producto(rstT("id_producto"), Me.DtcAlmacen.BoundText) & "','" & in_precio_costo & "','" & Me.DtcAlmacen.BoundText & "'," & _
        "'" & get_producto(rstT("id_producto")) & "','" & rstT("id_detalle") & "','" & rstT("incremento_neto") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        rstT.MoveNext
   Next i
End If



End Sub
Private Function put_compra() As Double

put_compra = get_id_compra("0001", Trim(Me.Txtfactura_serie.Text), Trim(Me.txtFactura_numero.Text), Trim(Me.txtRuc.Text))

If put_compra > 0 Then
    
    Exit Function
End If

        If Me.DtcMoneda.BoundText = "00001" Then
                in_cta_compra = KEY_CTA_COMPRA_SOLES
         Else
                in_cta_compra = KEY_CTA_COMPRA_DOLARES
         End If
           
        If Len(Trim(Me.txtRuc.Text)) = 8 Then
            cod_identidad = 1
        End If
        If Len(Trim(Me.txtRuc.Text)) = 11 Then
            cod_identidad = 6
        End If
        If Len(Trim(Me.txtRuc.Text)) <> 8 And Len(Trim(Me.txtRuc.Text)) <> 11 Then
            cod_identidad = 0
        End If
        
        If KEY_CONTABILIDAD = "si" Then
           If put_verifica_cuenta_contable("0001", Trim(Me.Txtfactura_serie.Text), Trim(Me.txtFactura_numero.Text), in_cta_compra, "00002") = False Then
              Exit Function
           End If
           
        End If
       in_responsable = "0"
       
       ' VALIDACION DE PERIODO CERRADO
       If get_cierre_periodo(Me.DtcPeriodoCompra.BoundText) = False Then
            MsgBox "EL PERIODO DE EMISION NO ESTA CREADO" + Chr(13) + "COORDINE CON EL AREA CONTABLE", vbInformation, KEY_VENDEDOR
           put_compra = 0
           Exit Function
       End If
       '******************************
       
       in_periodo_factura = Me.DtcPeriodoCompra.BoundText   'get_periodo_actual(Me.DtpEmision.Value)
       
       in_periodo_recepcion = get_periodo_actual(Me.DtpPedido.Value)
       
       
       
       
       
       If get_cierre_periodo(in_periodo_recepcion) = False Then
           MsgBox "EL PERIODO DE RECEPCION ESTA CERRADO" + Chr(13) + "COORDINE CON EL AREA CONTABLE", vbInformation, KEY_VENDEDOR
           put_compra = 0
           Exit Function
       End If
       
       
       
       If in_periodo_factura = 0 Then
           MsgBox "EL PERIODO DE EMISION NO ESTA CREADO" + Chr(13) + "COORDINE CON EL AREA CONTABLE", vbInformation, KEY_VENDEDOR
           put_compra = 0
           Exit Function
       End If
       
    'SEGURO DE PERIODOS
      
      ' If get_periodo_actual(Me.DtpPedido.Value) <> in_periodo_factura Then
      '      MsgBox "PERIODOS CONTABLES DIFERENTES  ES NECESARIO " + Chr(13) + "COORDINAR CON EL AREA CONTABLE", vbInformation, KEY_VENDEDOR
      '      put_compra = 0
      '      Exit Function
            
     '  End If
        
        in_exonerado = 0
        in_valor_venta = 0
        in_total = 0
        strCadena = "SELECT * FROM orden_compra WHERE id_orden='" & Val(Me.txtId_recepcion.Text) & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
        
       
        
            If rstL("afecto_igv") = "si" Then
                in_exonerado = 0
                in_valor_venta = rstL("total") / (1 + KEY_IGV)
                in_total = rstL("total")
                in_igv = in_total - in_valor_venta
            Else
                in_exonerado = rstL("total")
                in_valor_venta = 0
                in_total = rstL("total")
                in_igv = 0
            End If
       
        
        
        strCadena = "call P_insert_compra_ultimate('0001','" & Me.DtcAlmacen.BoundText & "','" & Format(Me.DtpEmision.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpVencimiento.Value, "YYYY-mm-dd") & "','02'," & _
        "'00002','--','" & Me.DtcMoneda.BoundText & "','" & formato_item(Month(Me.DtpEmision.Value), 2) & "','" & Year(Me.DtpEmision.Value) & "','" & Trim(Me.Txtfactura_serie.Text) & "'," & _
        "'" & Format(Trim(Me.txtFactura_numero.Text), "00000000") & "','" & cod_identidad & "','" & Trim(Me.txtRuc.Text) & "','" & UCase(Me.TxtProveedor.Text) & "','" & Trim(Me.txtTc.Text) & "'," & _
        "'0','" & in_valor_venta & "','" & in_igv & "','0','0','0','0','" & in_exonerado & "','0','" & in_total & "','" & in_total & "'," & _
        " '" & KEY_USUARIO & "','--','01','" & Me.DtcPeriodoCompra.BoundText & "','" & in_cta_compra & "','" & in_responsable & "','0','0','0','0','" & KEY_RUC & "')"
        Call ConfiguraRstP(strCadena)
        id_compra = rstP(0)
        
        strCadena = "UPDATE movimiento_compra SET id_tipo_compra='02',id_orden_compra='" & Val(Me.txtId_recepcion.Text) & "' WHERE id_compra='" & id_compra & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
        Call save_detalle_factura(id_compra)
        put_compra = id_compra
        strCadena = "p_update_proveedor('" & Trim(Me.txtRuc.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        End If
End Function

Private Sub cerrar_flete_Click()
Me.frmgastos.Visible = False
End Sub

Private Sub ChameleonBtn1_Click()


End Sub
Private Sub llenar_gastos(ByVal Grilla As MSHFlexGrid, ByVal id_compra As Double)
Dim Total As Double
strCadena = "SELECT * FROM view_factura_vinculada_gasto WHERE id_recepcion='" & id_compra & "' AND ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.txtMontoFlete.Text = Format(0, "###0.00")
    Grilla.Rows = 0
    
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 3000
           Grilla.ColWidth(4) = 800
           Grilla.ColWidth(5) = 500
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 1200
           Grilla.ColWidth(8) = 1200
        Next
         cabecera = "IDGASTO" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "CLIENTE/PROVEEDOR" & vbTab & "MONEDA" & vbTab & "TC" & vbTab & "MONTO" & vbTab & "VALOR VENTA [S/.]" & vbTab & "TOTAL [S/.]"
         Grilla.AddItem cabecera
         For k = 0 To 8
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        
        
             in_valor_venta = 0
             in_parcial = 0
             in_total = 0
             in_parcial_venta = 0
             in_total = 0
        For i = 1 To rst.RecordCount
             
            If rst("id_moneda") = "00002" Then
                in_parcial = rst("monto") * rst("tc")
            Else
               in_parcial = rst("monto")
            End If
             
            If KEY_CON_IGV = "si" Then
                If rst("afecto_igv") = "si" Then
                    in_valor_venta = in_valor_venta + in_parcial / (1 + KEY_IGV)
                    in_parcial_venta = in_parcial / (1 + KEY_IGV)
                Else
                    in_valor_venta = in_valor_venta + in_parcial
                    in_parcial_venta = in_parcial
                End If
            Else
                    in_valor_venta = in_valor_venta + in_parcial
                    in_parcial_venta = in_parcial
            End If
             in_total = in_total + in_parcial
             
             Fila = rst("id_gasto") & vbTab & rst("fecha") & vbTab & rst("comprobante") & vbTab & rst("nombre_completo") & vbTab & rst("moneda") & vbTab & Format(rst("tc"), "#,##0.0000") & vbTab & Format(rst("monto"), "#,##0.0000") & vbTab & Format(in_parcial_venta, "###0.000") & vbTab & Format(in_parcial, "###0.000")
             Grilla.AddItem Fila
             
        rst.MoveNext
        Next i
         cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "=========================" & vbTab & Format(in_valor_venta, "#,##0.0000") & vbTab & Format(in_total, "#,##0.0000")
         Grilla.AddItem cabecera
          For k = 0 To 8
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &HC0C0FF
                            Next k
                            Me.txtMontoFlete.Text = Format(in_valor_venta, "###0.00")
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub cmdagregar_Click()
If Me.DtcComrpobante.BoundText = "0110" Then
    strCadena = "call orden_compra_temporal('" & Trim(Me.TxtCodProducto.Text) & "','" & Trim(Me.txtCantidad.Text) & "','" & Val(Me.txtcosto.Text) & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & Val(Me.lblid_detalle.Caption) & "','" & Me.DtcComrpobante.BoundText & "','" & KEY_RUC & "')"
Else
    strCadena = "call up_orden_recepcion_temporal('" & Trim(Me.TxtCodProducto.Text) & "','" & Val(Me.txtCantidad.Text) & "','" & Val(Me.txtCantidad.Text) & "','" & Val(Me.txtcosto.Text) & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & Val(Me.lblid_detalle.Caption) & "','" & Me.DtcComrpobante.BoundText & "','" & KEY_RUC & "')"
End If
CnBd.Execute (strCadena)
Me.TxtCodProducto.Text = ""
Me.TxtUnidad.Text = ""
Me.txtcosto.Text = ""
Me.txtCantidad.Text = "1"
Me.TxtDescripcionProducto.Text = ""
Me.lblid_detalle.Caption = 0
If Me.DtcComrpobante.BoundText = "0110" Then
    Call Me.llenar_orden(Me.HfdDetalle, Val(Me.TxtId_orden.Text))
Else
    Call Me.llenar_recepcion(Me.HfdDetalle, Val(Me.TxtId_orden.Text))
    Call prorratear_flete
    Call Me.llenar_recepcion(Me.HfdDetalle, Val(Me.TxtId_orden.Text))
    
End If

Call Resalta(Me.TxtCodProducto)
End Sub

Private Sub cmdCerrarpantalla_Click()

Unload Me
Exit Sub
End Sub

Private Sub cmdEditable_Click()

End Sub

Private Sub cmddelete_Click()

End Sub

Private Sub cmdDeleteGasto_Click()
If MsgBox("Esta Seguro de eliminar este Registro", vbYesNo + vbQuestion, KEY_EMPRESA) = vbYes Then
            
            strCadena = "SELECT * FROM movimiento_compra_gasto WHERE id_gasto='" & Val(Me.mshGastos.TextMatrix(Me.mshGastos.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
                strCadena = "DELETE FROM movimiento_compra_gasto WHERE id_gasto='" & Val(Me.mshGastos.TextMatrix(Me.mshGastos.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                strCadena = "Call CON_Asiento_EliminarCompra('" & rstT("id_compra_gasto") & "', '" & KEY_USUARIO & "', '" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                strCadena = "DELETE FROM movimiento_compra WHERE id_compra='" & rstT("id_compra_gasto") & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                
            End If
            
             
        End If
         Call llenar_gastos(Me.mshGastos, Val(Me.TxtId_orden.Text))

End Sub

Private Sub cmdExtraerOrden_Click()
Call load_detalle_orden
End Sub
Private Sub load_detalle_orden()
strCadena = "DELETE FROM orden_compra_detalle_temp WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
 Dim in_orden As String
strCadena = "SELECT * FROM orden_compra WHERE id_doc='0110' and  serie='" & Trim(Me.TxtOrden_serie.Text) & "' and numero='" & Trim(Me.TxtOrden_numero.Text) & "' and ruc='" & KEY_RUC & "' ORDER BY id_orden DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   If MsgBox("Desea Realizar esta Transaccion", vbInformation + vbYesNo, KEY_VENDEDOR) = vbYes Then
      Me.txtId_recepcion.Text = rst("id_orden")
      Me.TxtId_orden.Text = 0
      Me.txtRuc.Text = rst("id_proveedor")
      Me.DtcMoneda.BoundText = rst("id_moneda")
      Me.TxtProveedor.Text = get_persona(rst("id_proveedor"))
      
      If DtcComrpobante.BoundText = "0414" Then
          Me.DtpPago.Value = rst("fecha_pago")
          Me.DtpPedido.Value = KEY_FECHA
      Else
        Me.DtpPago.Value = rst("fecha_pago")
         Me.DtpPedido.Value = rst("fecha_solicitud")
      End If
      
      
      
      
        strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and dni='" & rst("id_autorizado") & "'"
        Call ConfiguraRstT(strCadena)
        Call LlenaDataComboT(Me.DtcAutorizado)
        
        strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and dni='" & KEY_USUARIO & "'"
        Call ConfiguraRstT(strCadena)
        Call LlenaDataComboT(Me.DtcRecibido)
      If rst("afecto_igv") = "si" Then
         Me.chk_igv.Value = 1
      Else
        Me.chk_igv.Value = 0
      End If
      strCadena = "SELECT * FROM orden_compra_detalle WHERE id_orden='" & Val(Me.txtId_recepcion.Text) & "' ORDER BY id_detalle ASC"
      Call ConfiguraRstK(strCadena)
      If rstK.RecordCount > 0 Then
         rstK.MoveFirst
         For i = 0 To rstK.RecordCount - 1
              in_pendiente = rstK("cantidad") - get_recepcionado(rstK("id_producto"), Val(Me.txtId_recepcion.Text))
              strCadena = "call orden_compra_temporal_recepcion('" & rstK("id_producto") & "','" & in_pendiente & "','" & in_pendiente & "','" & rstK("precio") & "','" & KEY_ALM & "','" & KEY_USUARIO & "','0','" & Me.DtcComrpobante.BoundText & "','" & KEY_RUC & "')"
              CnBd.Execute (strCadena)
              
              rstK.MoveNext

         Next i
      End If
   End If
   
   strCadena = "SELECT * FROM movimiento_compra WHERE id_orden_compra='" & Val(Me.txtId_recepcion.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
   Call ConfiguraRstK(strCadena)
   If rstK.RecordCount > 0 Then
      Me.Txtfactura_serie.Text = rstK("serie")
      Me.txtFactura_numero.Text = rstK("numero")
   End If
   
   Call Me.llenar_recepcion(Me.HfdDetalle, 0)
End If

End Sub
Private Function get_recepcionado(ByVal in_producto As String, ByVal in_orden As String) As Single
strCadena = "SELECT sum(d.cantidad) FROM  orden_compra o,orden_compra_detalle d WHERE o.id_estado<>3 and  o.id_orden=d.id_orden and d.id_producto='" & in_producto & "' and o.id_recepcion='" & Val(in_orden) & "' and o.ruc='" & KEY_RUC & "'"
Call ConfiguraRstP(strCadena)
If IsNull(rstP(0)) = True Then
    get_recepcionado = 0
Else
    get_recepcionado = rstP(0)
End If


End Function
Private Sub cmdImprimir_Click()

Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant
Dim in_total As String

arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"



arr(0, 2) = Me.DtcMoneda.BoundText
If Val(Me.TxtTotal.Text) = 0 Then
    arr(1, 2) = "CERO CON 00/100 SOLES"
Else
    arr(1, 2) = UCase(EnLetras(Val(Me.TxtTotal.Text))) & Space(1) & Me.DtcMoneda.Text
End If



param = arr()

If Me.DtcComrpobante.BoundText = "0110" Then
    strCadena = "SELECT id_orden,comprobante,afecto_igv,fecha_registro,fecha_solicitud,fecha_pago,id_proveedor,nombre_completo,direccion,'AQUI',id_producto,nombre_prod,unidad,cantidad,precio,descuento,total,operador FROM view_orden_compra_print WHERE id_orden='" & Val(Me.TxtId_orden.Text) & "'"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptOrdenCompra", param, App.Path + "\Reportes\")
Else
    strCadena = "SELECT * FROM view_orden_recepcion WHERE id_orden='" & Val(Me.TxtId_orden.Text) & "'"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptOrdenRecepcion", param, App.Path + "\Reportes\")
End If


End Sub

Private Sub put_delete(ByVal in_detalle As String)
strCadena = "DELETE FROM orden_compra_detalle_temp WHERE id_detalle='" & Val(in_detalle) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

Call Me.llenar_orden(Me.HfdDetalle, Val(Me.TxtId_orden.Text))
End Sub


Private Sub cmdIngresarFactura_Click()

Dim nombreperiodo As String
Me.DtpFecha.Value = KEY_FECHA
Me.dtpvencimientoFlete.Value = KEY_FECHA
Me.txtTc.Text = KEY_CAMBIO
strCadena = "SELECT id_doc as Codigo,doc_des as Descripcion FROM comprobantes ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComprobante)
Me.DtcComprobante.BoundText = "0001"

strCadena = "SELECT tipo_compra as Codigo,descripcion as Descripcion FROM tipo_compra WHERE tipo_compra<>'01'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoCompra)

strCadena = "SELECT igv as Codigo,igv as Descripcion FROM afecto_igv ORDER BY igv"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAfectoIgv)

strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda ORDER BY id_moneda"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMonedaFlete)

strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPeriodo)
Me.DtcPeriodo.BoundText = get_periodo_actual(Me.DtpPedido.Value)
'Me.DtcPeriodo.Locked = True

 strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_personal='si'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcResponsable)
  
Me.TxtTipoCambio.Text = Format(KEY_CAMBIO_VENTA, "#,##0.0000")
Call llenar_gastos(Me.mshGastos, Val(Me.TxtId_orden.Text))
Me.frmgastos.Visible = True
Call Resalta(Me.TxtfleteSerie)
End Sub

Private Sub cmdnuevo_Click()

If MsgBox("Desea Limpiar esta Orden TEMPORAL", vbYesNo + vbQuestion, KEY_VENDEDOR) = vbYes Then
    strCadena = "DELETE FROM orden_compra_detalle_temp WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Call Me.llenar_orden(Me.HfdDetalle, Val(Me.TxtId_orden.Text))
    Me.TxtId_orden.Text = ""
    Me.txtId_recepcion.Text = ""
    Me.txtValorVenta.Text = ""
    Me.TxtIgv.Text = ""
    Me.TxtTotal.Text = ""
    Me.txtFacturaFlete.Text = ""
    Me.cmdProcesar.Enabled = True
End If


End Sub

Private Sub cmdProcesar_Click()

If Format(DtpEmision.Value, "YYYY-mm-dd") >= Format(get_fecha_periodo_abierto, "YYYY-mm-dd") Then
    Call Save
    Call FrmOrdenCompra.actualizar
Else
    
    If get_cierre_periodo(Me.DtcPeriodoCompra.BoundText) = False Then
        MsgBox "PERIODO CERRADO COORDINE CON EL AREA CONTABLE", vbInformation
    Else
        Call Save
        Call FrmOrdenCompra.actualizar
    End If
    Exit Sub
End If




End Sub
Private Function get_cierre_periodo(ByVal in_periodo As String) As Boolean

strCadena = "SELECT * FROM view_cierre_periodo WHERE IndCierreAlmacen='0' and  id_periodo='" & in_periodo & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
    get_cierre_periodo = True
Else
    get_cierre_periodo = False
End If


End Function





Private Sub cmdProcesarFlete_Click()
Dim cod_identidad As String * 1
Dim valor_venta As Double
Dim igv As Double
Dim Total As Double
in_nodomiciliada = "no"
in_retencion = 0
If Me.chk_nodomiciliada.Value = 1 Then
   in_nodomiciliada = "si"
End If

If Me.cmdAgregar.Caption = "Modificar" Then
    strCadena = "UPDATE movimiento_compra_gasto SET no_domiciliada='" & in_nodomiciliada & "',id_doc='" & Me.DtcComprobante.BoundText & "',serie='" & Me.txtSerie.Text & "',numero='" & Me.TxtNumero.Text & "', id_persona='" & Me.txtDni.Text & "' WHERE id_gasto='" & Me.mshGastos.TextMatrix(Me.mshGastos.Row, 0) & "' ANd ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    Call llenar_gastos(Me.mshGastos, Val(Me.TxtId_orden.Text))
    Me.cmdAgregar.Caption = "Agregar"
    Me.TxtfleteSerie.Text = ""
    Me.TxtFleteNumero.Text = ""
    Me.TxtNumero.Text = ""
    Me.txtDni.Text = ""
    Me.lblcliente.Caption = ""
    Me.txtMonto.Text = 0#
    Me.DtpFecha.Value = Me.DtpFecha.Value
    Me.txtDescripcion.Text = ""
    Me.DtcComprobante.SetFocus
    Exit Sub
End If


strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Me.DtcComprobante.BoundText & "' AND serie='" & Me.TxtfleteSerie.Text & "' AND numero='" & Me.TxtFleteNumero.Text & "' AND id_proveedor='" & Me.txtDni.Text & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
If Len(Trim(Me.txtDni.Text)) = 8 Then
    cod_identidad = 1
End If
If Len(Trim(Me.txtDni.Text)) = 11 Then
    cod_identidad = 6
End If
If Len(Trim(Me.txtDni.Text)) <> 8 And Len(Trim(Me.txtDni.Text)) <> 11 Then
    cod_identidad = 0
End If

If Me.DtcAfectoIgv.BoundText = "SI" Then
    exonerado = 0
    valor_venta = Val(Me.txtMonto.Text) / (KEY_IGV + 1)
    igv = Val(Me.txtMonto.Text) - valor_venta
Else
    
    igv = 0
    exonerado = Val(Me.txtMonto.Text)
    valor_venta = 0
    
End If


        
        If Me.DtcMoneda.BoundText = "00001" Then
                in_cta_compra = KEY_CTA_COMPRA_SOLES
           Else
                in_cta_compra = KEY_CTA_COMPRA_SOLES
           End If
           
            If Me.DtcComprobante.BoundText = "0002" Then
             in_cta_compra = KEY_CTA_COMPRA_RH
        End If
        
        
        If KEY_CONTABILIDAD = "si" Then
           If put_verifica_cuenta_contable(Me.DtcComprobante.BoundText, Trim(Me.TxtfleteSerie.Text), Trim(Me.TxtFleteNumero.Text), in_cta_compra, Me.DtcTipoCompra.BoundText) = False Then
              Exit Sub
           End If
        End If
       
      
        If Me.chk_nodomiciliada.Value = 1 Then
            exonerado = 0
            valor_venta = Val(Me.txtSubtotal.Text)
            igv = 0
            Me.txtMonto.Text = Val(Me.TxtTotal.Text)
            in_retencion = Val(Me.TxtRetencion.Text)
        
        End If
       
        
        
        strCadena = "call P_insert_compra_ultimate('" & Me.DtcComprobante.BoundText & "','" & KEY_ALM & "','" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "','" & Format(Me.dtpvencimientoFlete.Value, "YYYY-mm-dd") & "','02'," & _
        "'" & Me.DtcTipoCompra.BoundText & "','--','" & Me.DtcMonedaFlete.BoundText & "','" & formato_item(Month(Me.DtpFecha.Value), 2) & "','" & Year(Me.DtpFecha.Value) & "','" & Trim(Me.TxtfleteSerie.Text) & "'," & _
        "'" & Format(Trim(Me.TxtFleteNumero.Text), "00000000") & "','" & cod_identidad & "','" & Trim(Me.txtDni.Text) & "','" & UCase(Me.lblcliente.Caption) & "','" & Trim(Me.TxtTipoCambio.Text) & "'," & _
        "'0','" & valor_venta & "','" & igv & "','0','0','0','" & in_retencion & "','" & exonerado & "','0','" & Val(Me.txtMonto.Text) & "','" & Val(Me.txtMonto.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtDescripcion.Text) & "','02','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & Me.DtcResponsable.BoundText & "','0','0','0','0','" & KEY_RUC & "')"
        Call ConfiguraRstP(strCadena)
        id_compra = rstP(0)
        
        
        
        If Me.chk_nodomiciliada.Value = 1 Then
            strCadena = "UPDATE movimiento_compra SET no_domiciliada='si' WHERE id_compra='" & id_compra & "'"
            CnBd.Execute (strCadena)
        End If
        
        
        
        If Me.DtcAfectoIgv.BoundText = "SI" Then
            in_afecto = "si"
            in_exonerado = 0
            valor_venta = Val(Me.txtMonto.Text) / (KEY_IGV + 1)
            igv = Val(Me.txtMonto.Text) - valor_venta
        Else
            in_afecto = "no"
            igv = 0
            in_exonerado = Val(Me.txtMonto.Text)
            valor_venta = 0
    
        End If
        
        
        If Me.chk_nodomiciliada.Value = 1 Then
            exonerado = 0
            valor_venta = Val(Me.txtSubtotal.Text)
            igv = 0
            Me.txtMonto.Text = valor_venta
            in_retencion = Val(Me.TxtRetencion.Text)
        End If
        
        strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,ivap,otros,percepcion, " & _
        "valor_venta,exonerado,total,p_venta,p_costo,id_alm,retencion,ruc) VALUES ('" & id_compra & "','" & Trim(Me.TxtcodigoProd.Text) & "','1','" & Val(Me.txtMonto.Text) & "'," & _
        "'0','0','0','" & valor_venta & "','0','" & igv & "', " & _
        "'0','0','0','" & valor_venta & "','" & in_exonerado & "','" & Val(Me.txtMonto.Text) & "','" & Val(Me.txtMonto.Text) & "','" & get_precio_costo(Trim(Me.TxtcodigoProd.Text)) & "','" & KEY_ALM & "','" & in_retencion & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        

        strCadena = "INSERT INTO movimiento_compra_gasto(id_compra,id_persona,id_doc,serie,numero,monto,fecha,descripcion,tc,id_moneda,id_compra_gasto,afecto_igv,id_orden_compra,ruc)VALUES " & _
        " ('" & Val(Me.TxtId_orden.Text) & "','" & Me.txtDni.Text & "','" & DtcComprobante.BoundText & "','" & Trim(Me.TxtfleteSerie.Text) & "','" & Trim(Me.TxtFleteNumero.Text) & "','" & Val(Me.txtMonto.Text) & "','" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "','" & Me.txtDescripcion.Text & "','" & Val(Me.TxtTipoCambio.Text) & "','" & Me.DtcMonedaFlete.BoundText & "','" & id_compra & "','" & in_afecto & "','" & Val(Me.txtId_recepcion.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)


        
'02----------------guardar en detalle documento Compra-----------
 
 
' If KEY_CONTABILIDAD = "si" And Me.DtcComprobante.BoundText <> "0089" Then
'    strCadena = "p_insert_compra_emitido_ii('" & id_compra & "')"
'    Call Execute_Sql(strCadena)
' End If
 
 
 Me.txtIdfactura_flete.Text = id_compra
 Call llenar_gastos(Me.mshGastos, Val(Me.TxtId_orden.Text))
 Call prorratear_flete
        
Me.lblidCompra.Caption = get_periodo_detalle(Me.DtcPeriodo.BoundText, id_compra)
MsgBox "NUMERO DE VOUCHER GENERADO  : " & Trim(Me.lblidCompra.Caption), vbInformation, KEY_VENDEDOR
Me.txtFacturaFlete.Text = Trim(DtcComprobante.Text) & ":" & Trim(Me.TxtfleteSerie.Text) & "-" & Trim(Me.TxtFleteNumero.Text)


    
Call Me.llenar_recepcion(Me.HfdDetalle, Val(Me.TxtId_orden.Text))


Me.txtDni.Text = ""
Me.lblcliente.Caption = ""
Me.txtMonto.Text = 0#
Me.DtpFecha.Value = KEY_FECHA
Me.txtDescripcion.Text = ""
Me.DtcComprobante.SetFocus
Me.lblcuenta_contable.Caption = ""
Me.lblcuenta_detalle.Caption = ""
Me.frmgastos.Visible = False


Exit Sub
Else
    MsgBox "COMPROBANTE YA REGISTRADO, IMPOSIBLE GUARDAR ", vbInformation, KEY_EMPRESA
End If
Set rst = Nothing
End Sub
Public Sub prorratear_flete()
Dim in_monto_gasto As Single
strCadena = "SELECT * FROM orden_compra_detalle_temp WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   in_monto_gasto = Val(Me.txtMontoFlete.Text)
   
   If Me.DtcMonedaFlete.BoundText = "00002" Then
      in_monto_gasto = in_monto_gasto / Val(Me.txtTc.Text)
   End If
   
   For i = 0 To rst.RecordCount - 1
        If in_monto_gasto > 0 Then
            in_monto_parcial = rst("precio") / (Val(Me.txtValorVenta.Text)) * in_monto_gasto
            in_monto_porcentaje = in_monto_parcial * 100 / in_monto_gasto
       Else
           in_monto_parcial = 0
           in_monto_porcentaje = 0
        
        End If
        
        
        strCadena = "UPDATE orden_compra_detalle_temp SET incremento_neto='" & in_monto_parcial & "'  WHERE id_detalle='" & rst("id_detalle") & "'"
        CnBd.Execute (strCadena)
        rst.MoveNext
   Next i
   
  
  
End If


End Sub
Private Sub CmdQuitar_Click()
If Val(Me.HfdDetalle.Rows) > 0 Then
Call put_delete(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))

If Me.DtcComrpobante.BoundText = "0110" Then
    Call Me.llenar_orden(Me.HfdDetalle, Val(Me.TxtId_orden.Text))
Else
    Call Me.llenar_recepcion(Me.HfdDetalle, Val(Me.TxtId_orden.Text))
End If



End If
End Sub

Private Sub cmdupdate_Click()

strCadena = "SELECT * FROM view_orden_compra_detalle_temp WHERE ruc='" & KEY_RUC & "' and  id_detalle='" & Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    Me.TxtCodProducto.Text = rstK("id_producto")
    Me.txtCantidad.Text = rstK("cantidad")
    Me.TxtDescripcionProducto.Text = rstK("nombre_prod")
    Me.TxtUnidad.Text = rstK("unidad")
    Me.txtcosto.Text = rstK("precio")
    Me.lblid_detalle.Caption = Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))
    Me.cmdupdate.Enabled = False
End If



End Sub

Private Sub Command1_Click()
End Sub

Private Sub DtcAfectoIgv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcMonedaFlete.SetFocus
End If
End Sub

Private Sub DtcComrpobante_Change()
Call get_comprobante(Me.DtcComrpobante.BoundText)
If Me.DtcComrpobante.BoundText = "0414" Then
   Me.frmrecepcion.Visible = True
Else
   Me.frmrecepcion.Visible = False
End If
End Sub

Private Sub DtcMonedaFlete_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Me.DtcPeriodo.SetFocus
End If
End Sub

Private Sub DtcPeriodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtcodigoProd)
End If
End Sub

Private Sub DtpPedido_Change()
Me.txtTc.Text = cambio_venta(CVDate(Me.DtpPedido.Value))
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
strCadena = "SELECT id_termino as Codigo,descripcion as Descripcion FROM terminos_entrega ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTerminosEntrega)


strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)
Me.DtcMoneda.BoundText = "00001"




strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' and id_tipoentidad='0' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM

strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM view_almacen_comprobante_ultimate WHERE id_doc IN('0110','0414') and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComrpobante)
Me.DtcComrpobante.BoundText = "0110"

Me.lblruc.Caption = KEY_RUC
Me.LblEmpresa.Caption = KEY_EMPRESA


strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPeriodoCompra)
Me.DtcPeriodoCompra.BoundText = get_periodo_actual(KEY_FECHA)


strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_personal='si'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAutorizado)


Me.DtpPedido.Value = KEY_FECHA
Me.DtpPago.Value = KEY_FECHA
Me.DtpEmision.Value = KEY_FECHA
Me.DtpVencimiento.Value = KEY_FECHA
Me.txtTc.Text = KEY_CAMBIO_VENTA



End Sub
Private Sub get_comprobante(ByVal in_doc As String)
strCadena = "SELECT * FROM orden_compra WHERE id_doc='" & in_doc & "' and   ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
     Me.txtSerie.Text = rst("serie")
     Me.TxtNumero.Text = Format(Val(rst("numero")) + 1, "000000")
Else
    Me.txtSerie.Text = "001"
    Me.TxtNumero.Text = "000001"
End If
End Sub
Private Function get_nueva_orden(ByVal in_doc As String, ByVal in_serie As String) As String
strCadena = "SELECT * FROM orden_compra WHERE id_doc='" & in_doc & "' and serie='" & in_serie & "' and   ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
     get_nueva_orden = Format(Val(rstL("numero")) + 1, "000000")
Else
     get_nueva_orden = "000001"
End If

End Function
Private Function get_factura_flete(ByVal in_flete As String)
strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & Val(in_flete) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
      get_factura_flete = rstZ("id_doc") & "-" & rstZ("serie") & "-" & rstZ("numero")
     If KEY_CON_IGV = "si" Then
        Me.txtMontoFlete.Text = rstZ("total") / (1 + KEY_IGV)
     Else
        Me.txtMontoFlete.Text = rstZ("total")
     End If
        
      
      
      
Else
    get_factura_flete = ""
End If
End Function
Public Sub get_orden(ByVal in_orden As String)
Dim in_afecto As String
Dim in_factura_flete As String

strCadena = "SELECT * FROM orden_compra WHERE id_orden='" & Val(in_orden) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   in_factura_flete = rst("id_factura_flete")
   in_afecto = rst("afecto_igv")
   Me.TxtId_orden.Text = rst("id_orden")
   Me.txtObservacion.Text = rst("observacion")
   Me.txtSerie.Text = rst("serie")
   Me.TxtNumero.Text = rst("numero")
   Me.txtRuc.Text = rst("id_proveedor")
   Me.TxtProveedor.Text = get_persona(rst("id_proveedor"))
   Me.DtpPedido.Value = rst("fecha_solicitud")
   Me.DtpPago.Value = rst("fecha_pago")
   Me.DtcMoneda.BoundText = rst("id_moneda")
   Me.txtDni.Text = rst("id_proveedor")
   Me.txtFacturaFlete.Text = get_factura_flete(in_factura_flete)
   Me.txtGuia_serie.Text = rst("guia_serie")
   Me.TxtGuia_numero.Text = rst("guia_numero")
   
   lblcliente.Caption = get_persona(rst("id_proveedor"))
   
   strCadena = "SELECT serie,numero FROM movimiento_compra WHERE id_compra='" & rst("id_compra") & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRstT(strCadena)
   If rstT.RecordCount > 0 Then
      Me.Txtfactura_serie.Text = rstT("serie")
      Me.txtFactura_numero.Text = rstT("numero")
   End If
   
   
   strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and dni='" & rst("dni_save") & "'"
   Call ConfiguraRstT(strCadena)
   Call LlenaDataComboT(Me.DtcCreador)
   
   
   strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and dni='" & rst("id_autorizado") & "'"
   Call ConfiguraRstT(strCadena)
   Call LlenaDataComboT(Me.DtcAutorizado)
    
    
   
   strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and dni='" & rst("dni_recepcion") & "'"
   Call ConfiguraRstT(strCadena)
   Call LlenaDataComboT(Me.DtcRecibido)
   
    Me.DtcComrpobante.BoundText = rst("id_doc")
   
    If in_afecto = "si" Then
      Me.chk_igv.Value = 1
    Else
       Me.chk_igv.Value = 0
   End If
    
   Call llenar_orden(Me.HfdDetalle, Val(Me.TxtId_orden.Text))
   Me.cmdProcesar.Enabled = False
End If

End Sub
Public Sub llenar_orden(ByVal Grilla As MSHFlexGrid, ByVal id_orden As Double)
'On Error GoTo salir
Dim tTotal As Double
If Val(id_orden) > 0 Then
    strCadena = "SELECT * FROM view_orden_compra_detalle WHERE id_orden='" & id_orden & "' and ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT * FROM view_orden_compra_detalle_temp WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
End If

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub

End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 5500
           Grilla.ColWidth(3) = 2000
           Grilla.ColWidth(4) = 2200
           Grilla.ColWidth(5) = 1300
           Grilla.ColWidth(6) = 1800
           Grilla.ColWidth(7) = 2000
        Next
        cabecera = "IDDETALLE" & vbTab & "CODIGO" & vbTab & "PRODUCTO" & vbTab & "UNIDAD" & vbTab & "CLASIFICACION" & vbTab & "CANTIDAD" & vbTab & "PRECIO COSTO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 7
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
          Next k
        tTotal = 0
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            tTotal = tTotal + rst("total")
            Fila = rst("id_detalle") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("unidad") & vbTab & rst("linea") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("precio"), "#,##0.0000") & vbTab & Format(rst("total"), "#,##0.00")
            Grilla.AddItem Fila
            
            rst.MoveNext
        Next i
        
        cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "===============" & vbTab & Format(tTotal, "#,##0.00")
        Grilla.AddItem cabecera
        For k = 6 To 7
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
          Next k
          Me.lblCantidad.Caption = rst.RecordCount
          
          in_bruto = Val(tTotal)
          in_descuento = Val(Me.TxtDescuento.Text)
          
          Me.txtImporteBruto.Text = Format(tTotal, "###0.00")
          Me.TxtDescuento.Text = Format(Val(Me.TxtDescuento.Text), "###0.00")
          If Me.chk_igv.Value = 1 Then
             
             
             Me.txtValorVenta.Text = Format(Val(tTotal), "###0.00")
             Me.TxtIgv.Text = Format(Val(Me.txtValorVenta.Text) * KEY_IGV, "###0.00")
             Me.TxtTotal.Text = Format(Val(Me.txtValorVenta.Text) + Val(Me.TxtIgv.Text), "###0.00")
          Else
             Me.TxtTotal.Text = Format((tTotal - in_descuento), "###0.00")
             Me.txtValorVenta.Text = Format(Val(Me.TxtTotal.Text), "###0.00")
             Me.TxtIgv.Text = Format(Val(Me.TxtTotal.Text) - Val(Me.txtValorVenta.Text), "###0.00")
          End If
          
'Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub
Public Sub llenar_recepcion(ByVal Grilla As MSHFlexGrid, ByVal id_orden As Double)
'On Error GoTo salir
Dim tTotal As Double
If Val(id_orden) > 0 Then
    strCadena = "SELECT * FROM view_orden_compra_detalle WHERE id_orden='" & id_orden & "' and ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT * FROM view_orden_compra_detalle_temp WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
End If

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub

End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 4500
           Grilla.ColWidth(3) = 1500
           Grilla.ColWidth(4) = 2200
           Grilla.ColWidth(5) = 1400
           Grilla.ColWidth(6) = 1400
           Grilla.ColWidth(7) = 1400
           Grilla.ColWidth(8) = 1400
           Grilla.ColWidth(9) = 1400
        Next
        cabecera = "IDDETALLE" & vbTab & "CODIGO" & vbTab & "PRODUCTO" & vbTab & "UNIDAD" & vbTab & "CLASIFICACION" & vbTab & "CANTIDAD" & vbTab & "CANT PENDIENTE" & vbTab & "PRECIO COSTO" & vbTab & "INC.FLETE" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
          Next k
        tTotal = 0
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            tTotal = tTotal + rst("total")
            Fila = rst("id_detalle") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("unidad") & vbTab & rst("linea") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("cantidad_pendiente"), "#,##0.00") & vbTab & Format(rst("precio"), "#,##0.0000") & vbTab & Format(rst("incremento_neto"), "#,##0.0000") & vbTab & Format(rst("total"), "#,##0.00")
            Grilla.AddItem Fila
            
            rst.MoveNext
        Next i
        
        cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "===============" & vbTab & Format(tTotal, "#,##0.00")
        Grilla.AddItem cabecera
        For k = 6 To 9
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
          Next k
          Me.lblCantidad.Caption = rst.RecordCount
          
          in_bruto = Val(tTotal)
          in_descuento = Val(Me.TxtDescuento.Text)
          
          Me.txtImporteBruto.Text = Format(tTotal, "###0.00")
          Me.TxtDescuento.Text = Format(Val(Me.TxtDescuento.Text), "###0.00")
          If Me.chk_igv.Value = 1 Then
             
             
             Me.txtValorVenta.Text = Format(Val(tTotal), "###0.00")
             Me.TxtIgv.Text = Format(Val(Me.txtValorVenta.Text) * KEY_IGV, "###0.00")
             Me.TxtTotal.Text = Format(Val(Me.txtValorVenta.Text) + Val(Me.TxtIgv.Text), "###0.00")
          Else
             Me.TxtTotal.Text = Format((tTotal - in_descuento), "###0.00")
             Me.txtValorVenta.Text = Format(Val(Me.TxtTotal.Text), "###0.00")
             Me.TxtIgv.Text = Format(Val(Me.TxtTotal.Text) - Val(Me.txtValorVenta.Text), "###0.00")
          End If
          
'Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub

Public Sub nuevo_registro()
Me.TxtId_orden.Text = 0
strCadena = "SELECT * FROM orden_compra WHERE id_doc='" & DtcComrpobante.BoundText & "' and  ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtSerie.Text = rst("serie")
   Me.TxtNumero.Text = Format(Val(rst("numero")) + 1, "000000")
   
   
   
Else
   Me.txtSerie.Text = "001"
   Me.TxtNumero.Text = "000001"
End If
Me.DtcComrpobante.Enabled = True
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and dni='" & KEY_USUARIO & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcCreador)
Me.DtcRecibido.BoundText = ""
Call Me.llenar_orden(Me.HfdDetalle, Val(Me.TxtId_orden.Text))
End Sub



Private Sub HfdDetalle_SelChange()
If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    Me.cmdupdate.Enabled = True
Else
    Me.cmdupdate.Enabled = False
End If

End Sub

Private Sub txtAutorizado_Change()
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_personal='si' and nombre_completo LIKE '%" & Trim(Me.txtAutorizado.Text) & "%' ORDER BY nombre_completo"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAutorizado)
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtDescuentoParcial)
End If
End Sub

Private Sub TxtcodigoProd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM producto where id_producto='" & Trim(Me.TxtcodigoProd.Text) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.TxtcodigoProd.Text = Trim(Me.TxtcodigoProd.Text)
       Me.Txtproducto.Text = rst("nombre_prod")
    Else
        Procedencia = buscar
        FrmProducto.Show
        Exit Sub
    End If
End If
End Sub

Private Sub TxtCodProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtCodProducto.Text = Format(Trim(Me.TxtCodProducto.Text), "00000")
    strCadena = "SELECT * FROM view_producto WHERE id_producto = '" & Trim(Me.TxtCodProducto.Text) & "' AND ruc='" & KEY_RUC & "' and id_alm='" & KEY_ALM & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.TxtCodProducto.Text = rst("id_producto")
        Me.TxtDescripcionProducto.Text = rst("nombre_prod")
        Me.TxtUnidad.Text = rst("unidad")
        Call Resalta(Me.txtCantidad)
 
 Else
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
 End If
End If
End Sub

Private Sub txtcosto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.cmdAgregar.SetFocus
End If
End Sub

Private Sub TxtDescuentoParcial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtcosto)
End If
End Sub

Private Sub txtDni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtDni.Text) & "'"
   Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.lblcliente.Caption = UCase(rst("nombre_completo"))
        Call Resalta(Me.txtMonto)
    Else
        If Len(Me.txtDni.Text) > 7 And Len(Me.txtDni.Text) < 12 Then
        Procedencia = nuevo
        FrmDetallePersona.Show
        FrmDetallePersona.txtRuc.Text = Trim(Me.txtDni)
        FrmDetallePersona.chkProveedor.Value = 1
        FrmDetallePersona.ChkCliente.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
        Else
            Procedencia = buscar
            FrmPersona.Show
            Exit Sub
        End If
End If

End If
End Sub

Private Sub txtFactura_numero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtFactura_numero.Text = Format(Me.txtFactura_numero.Text, "00000000")
End If
End Sub





Private Sub Txtfactura_serie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Txtfactura_serie.Text = Format(Me.Txtfactura_serie.Text, "0000")
    Call Resalta(Me.txtFactura_numero)
End If

End Sub

Private Sub TxtFleteNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtFleteNumero.Text = Format(Trim(Me.TxtFleteNumero.Text), "00000000")
    Call Resalta(Me.txtDni)
End If
End Sub

Private Sub TxtfleteSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtfleteSerie.Text = formato_item(Me.TxtfleteSerie.Text, 3)
    Call Resalta(Me.TxtFleteNumero)
End If
End Sub

Private Sub TxtGuia_numero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.TxtGuia_numero.Text = Format(Trim(Me.TxtGuia_numero.Text), "000000")
   Call Resalta(Me.Txtfactura_serie)
   
End If
End Sub

Private Sub txtGuia_serie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtGuia_serie.Text = Format(Trim(Me.txtGuia_serie.Text), "0000")
    Call Resalta(Me.TxtGuia_numero)
End If
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcAfectoIgv.SetFocus
End If
End Sub

Private Sub TxtOrden_numero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtOrden_numero.Text = Format(Me.TxtOrden_numero.Text, "000000")
    Call load_detalle_orden
    Call Resalta(Me.txtGuia_serie)
End If
End Sub

Private Sub txtOrden_serie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtOrden_serie.Text = Format(Me.TxtOrden_serie.Text, "000")
    Call Resalta(Me.TxtOrden_numero)
End If

End Sub

Private Sub txtRecibido_Change()
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_personal='si' and nombre_completo LIKE '%" & Trim(Me.txtAutorizado.Text) & "%' ORDER BY nombre_completo"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcRecibido)

End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
            Me.TxtProveedor.Text = (rst("nombre_completo"))
            Call Resalta(Me.TxtCodProducto)
   Else
        Procedencia = Selecionar
        FrmPersona.Show
        Exit Sub
   End If
End If
End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmComprasPagos 
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
   Begin VB.Frame frmretencion 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   520
      Left            =   3120
      TabIndex        =   105
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txtMontoRetencion 
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
         Left            =   3600
         MaxLength       =   80
         TabIndex        =   109
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txtCtaRetencion 
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
         Left            =   1320
         MaxLength       =   80
         TabIndex        =   107
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO :"
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
         Left            =   2805
         TabIndex        =   108
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CTA RETENCION :"
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
         Left            =   135
         TabIndex        =   106
         Top             =   120
         Width           =   1155
      End
   End
   Begin VB.Frame frm_trabajores 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   7605
      TabIndex        =   94
      Top             =   2220
      Visible         =   0   'False
      Width           =   10575
      Begin VB.TextBox txtMonto_trabajador 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6720
         MaxLength       =   80
         TabIndex        =   100
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox txtBusqueda_dni 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   720
         MaxLength       =   80
         TabIndex        =   97
         Top             =   3480
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfTrabador 
         Height          =   2895
         Left            =   240
         TabIndex        =   96
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5106
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
      Begin VitekeySoft.ChameleonBtn cmdagregar 
         Height          =   330
         Left            =   7785
         TabIndex        =   98
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "AGREGAR"
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
         MICON           =   "FrmComprasPagos.frx":0000
         PICN            =   "FrmComprasPagos.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdDelete 
         Height          =   330
         Left            =   9120
         TabIndex        =   102
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "QUITAR"
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
         MICON           =   "FrmComprasPagos.frx":2547
         PICN            =   "FrmComprasPagos.frx":2563
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   10320
         Picture         =   "FrmComprasPagos.frx":2AFD
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lbltrabajador 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   103
         Top             =   3480
         Width           =   3855
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO :"
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
         Left            =   6105
         TabIndex        =   101
         Top             =   3555
         Width           =   615
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI :"
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
         Left            =   300
         TabIndex        =   99
         Top             =   3555
         Width           =   345
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LISTADO DE GASTOS"
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
         Left            =   255
         TabIndex        =   95
         Top             =   120
         Width           =   1365
      End
   End
   Begin VB.TextBox txtMonto_dolares 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   92
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   6960
      TabIndex        =   89
      Top             =   3840
      Width           =   615
   End
   Begin VB.Frame frmmonto_pagar 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   15000
      TabIndex        =   84
      Top             =   960
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox txtMontoPagar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Left            =   1320
         MaxLength       =   80
         TabIndex        =   86
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOTA: ELIJA EL TIPO CAMBIO"
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
         Left            =   150
         TabIndex        =   110
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblid_compra 
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   3045
         Picture         =   "FrmComprasPagos.frx":59A1
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO PAGO :"
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
         Left            =   75
         TabIndex        =   85
         Top             =   600
         Width           =   1035
      End
   End
   Begin VB.CheckBox chk_ajustes_contables 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "AJUSTES CONTABLES"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   1560
      TabIndex        =   83
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Frame frmajustes_contables 
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
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   1560
      TabIndex        =   74
      Top             =   5100
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CheckBox chk_trabajador 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "TRABAJADOR"
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
         Height          =   250
         Left            =   3960
         TabIndex        =   93
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtMontoAnticipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2880
         MaxLength       =   80
         TabIndex        =   80
         Top             =   690
         Width           =   975
      End
      Begin VB.TextBox txtMontoRedondeo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2880
         MaxLength       =   80
         TabIndex        =   79
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtCuenta_anticipo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   990
         MaxLength       =   80
         TabIndex        =   76
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox txtCuenta_redondeo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   990
         MaxLength       =   80
         TabIndex        =   75
         Top             =   360
         Width           =   1095
      End
      Begin VitekeySoft.ChameleonBtn cmdAgregarGasto 
         Height          =   330
         Left            =   3960
         TabIndex        =   104
         Top             =   680
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "AGREGAR ..."
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
         MICON           =   "FrmComprasPagos.frx":8845
         PICN            =   "FrmComprasPagos.frx":8861
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO :"
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
         Left            =   2280
         TabIndex        =   82
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO :"
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
         Left            =   2280
         TabIndex        =   81
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ANTICIPO :"
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
         TabIndex        =   78
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REDONDEO :"
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
         Left            =   120
         TabIndex        =   77
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.TextBox txtMonedaOrigen 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   330
      Left            =   3600
      MaxLength       =   80
      TabIndex        =   73
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxtAcumulado_pagar 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   330
      Left            =   3600
      MaxLength       =   80
      TabIndex        =   72
      Top             =   2700
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtid_recibo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   7680
      MaxLength       =   80
      TabIndex        =   71
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtBuscar_proveedor 
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
      Left            =   16395
      MaxLength       =   80
      TabIndex        =   70
      Top             =   4995
      Width           =   2895
   End
   Begin VB.TextBox txtbuscar_ruc 
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
      Left            =   11955
      MaxLength       =   80
      TabIndex        =   68
      Top             =   4995
      Width           =   1815
   End
   Begin VB.TextBox txtbuscar_comprobante 
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
      Left            =   9480
      MaxLength       =   80
      TabIndex        =   66
      Top             =   4995
      Width           =   1575
   End
   Begin VB.TextBox TxtObservacion 
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
      Height          =   850
      Left            =   8040
      MaxLength       =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   63
      Top             =   8280
      Width           =   8535
   End
   Begin VitekeySoft.ChameleonBtn cmdProcesar 
      Height          =   855
      Left            =   16680
      TabIndex        =   59
      Top             =   8340
      Width           =   975
      _ExtentX        =   1720
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmComprasPagos.frx":AD8C
      PICN            =   "FrmComprasPagos.frx":ADA8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtOperacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   80
      TabIndex        =   57
      Top             =   4680
      Width           =   1695
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
      Left            =   6960
      TabIndex        =   52
      Top             =   3480
      Width           =   615
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
      Left            =   6960
      TabIndex        =   51
      Top             =   4245
      Width           =   615
   End
   Begin VB.TextBox TxtCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      MaxLength       =   80
      TabIndex        =   30
      Top             =   1500
      Width           =   6015
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1575
      MaxLength       =   80
      TabIndex        =   29
      Top             =   1140
      Width           =   1455
   End
   Begin VB.TextBox TxtSerie 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   5160
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   28
      Text            =   "000"
      Top             =   135
      Width           =   615
   End
   Begin VB.TextBox TxtNumeroDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   330
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   27
      Text            =   "000000"
      Top             =   135
      Width           =   1695
   End
   Begin VB.TextBox TxtSaldo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   26
      Top             =   2295
      Width           =   1935
   End
   Begin VB.TextBox TxtMontoPago 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   330
      Left            =   1560
      MaxLength       =   80
      TabIndex        =   25
      Top             =   2700
      Width           =   1935
   End
   Begin VB.Frame FrmCheque 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PAGAR CON CHEQUE"
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
      Height          =   855
      Left            =   1560
      TabIndex        =   21
      Top             =   8280
      Width           =   6015
      Begin VitekeySoft.ChameleonBtn cmdCargarCheque 
         Height          =   525
         Left            =   4440
         TabIndex        =   50
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   926
         BTYPE           =   5
         TX              =   "CHEQUE"
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
         MICON           =   "FrmComprasPagos.frx":E3F0
         PICN            =   "FrmComprasPagos.frx":E40C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.OptionButton OptChequeNO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NO"
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
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton OptChequeSi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SI"
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
         Height          =   375
         Left            =   720
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin MSDataListLib.DataCombo DtcCheque 
         Height          =   330
         Left            =   1320
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
   End
   Begin VB.TextBox TxtTc 
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
      Left            =   5880
      MaxLength       =   80
      TabIndex        =   20
      Top             =   2420
      Width           =   1695
   End
   Begin VB.CheckBox chkdetraccion 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "AFECTO A DETRACCION"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   1680
      TabIndex        =   4
      Top             =   6240
      Width           =   2175
   End
   Begin VB.TextBox txtmontototal 
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
      Left            =   5865
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   3
      Top             =   2055
      Width           =   1695
   End
   Begin VB.TextBox TxtItf 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   330
      Left            =   5880
      MaxLength       =   80
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox txtid_compra 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   7680
      MaxLength       =   80
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker DtpEmision 
      Height          =   315
      Left            =   3840
      TabIndex        =   0
      Top             =   1125
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   211025921
      CurrentDate     =   41130
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   345
      Left            =   240
      TabIndex        =   31
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
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
   Begin MSDataListLib.DataCombo DtcCuentas 
      Height          =   315
      Left            =   1560
      TabIndex        =   32
      Top             =   3120
      Width           =   6015
      _ExtentX        =   10610
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
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   345
      Left            =   2520
      TabIndex        =   33
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   609
      _Version        =   393216
      Locked          =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      ForeColor       =   8388608
      ListField       =   "0000�"
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
      Height          =   315
      Left            =   1560
      TabIndex        =   34
      Top             =   1935
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshFacturas 
      Height          =   4455
      Left            =   8040
      TabIndex        =   35
      Top             =   360
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7858
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
   Begin MSComCtl2.DTPicker DtpValor 
      Height          =   320
      Left            =   6105
      TabIndex        =   36
      Top             =   1125
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   211025921
      CurrentDate     =   41130
   End
   Begin VB.Frame frmdetraccion 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Height          =   1965
      Left            =   1560
      TabIndex        =   5
      Top             =   6240
      Width           =   6015
      Begin VB.TextBox txtprocentajedetraccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   80
         TabIndex        =   12
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtmontodetraccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   80
         TabIndex        =   11
         Top             =   930
         Width           =   1455
      End
      Begin VB.TextBox txtnumerodetraccion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   80
         TabIndex        =   10
         Top             =   1605
         Width           =   1455
      End
      Begin VB.TextBox txtconstancia 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   80
         TabIndex        =   8
         Top             =   1260
         Width           =   1455
      End
      Begin VB.TextBox txttasa 
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
         Left            =   4560
         MaxLength       =   80
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker Dtpfechadetraccion 
         Height          =   345
         Left            =   4080
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
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
         Format          =   211025921
         CurrentDate     =   42752
      End
      Begin MSDataListLib.DataCombo DtcCompra 
         Height          =   315
         Left            =   3240
         TabIndex        =   9
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSDataListLib.DataCombo DtcTiposervicio 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
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
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PORCENTAJE :"
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
         TabIndex        =   19
         Top             =   600
         Width           =   930
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO DETRACC:"
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
         Left            =   210
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N� DE OPERACION :"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N� CONSTANCIA :"
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
         Left            =   285
         TabIndex        =   16
         Top             =   1260
         Width           =   1140
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T.SERVICIO :"
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
         TabIndex        =   15
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:"
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
         Left            =   3360
         TabIndex        =   14
         Top             =   1680
         Width           =   540
      End
   End
   Begin MSDataListLib.DataCombo DtcFormaPago 
      Height          =   315
      Left            =   1560
      TabIndex        =   53
      Top             =   3480
      Width           =   5295
      _ExtentX        =   9340
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
      Left            =   1560
      TabIndex        =   54
      Top             =   4245
      Width           =   5295
      _ExtentX        =   9340
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
   Begin VitekeySoft.ChameleonBtn cmdImprimir 
      Height          =   855
      Left            =   17760
      TabIndex        =   60
      Top             =   8340
      Width           =   975
      _ExtentX        =   1720
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmComprasPagos.frx":11582
      PICN            =   "FrmComprasPagos.frx":1159E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   855
      Left            =   18840
      TabIndex        =   61
      Top             =   8340
      Width           =   975
      _ExtentX        =   1720
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmComprasPagos.frx":13B6F
      PICN            =   "FrmComprasPagos.frx":13B8B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfLista 
      Height          =   2655
      Left            =   8040
      TabIndex        =   62
      Top             =   5400
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4683
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
   Begin VitekeySoft.ChameleonBtn cmdQuitar_pago 
      Height          =   255
      Left            =   18960
      TabIndex        =   87
      Top             =   75
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BTYPE           =   5
      TX              =   "QUITAR"
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
      MICON           =   "FrmComprasPagos.frx":13F7B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcForma_pago_detalle 
      Height          =   315
      Left            =   1560
      TabIndex        =   90
      Top             =   3840
      Width           =   5295
      _ExtentX        =   9340
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
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FORMA DE PAGO:"
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
      TabIndex        =   91
      Top             =   3960
      Width           =   1170
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0C0C0&
      Height          =   465
      Left            =   8040
      Top             =   4905
      Width           =   11895
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR :"
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
      Left            =   15360
      TabIndex        =   69
      Top             =   5040
      Width           =   945
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC:"
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
      Left            =   11460
      TabIndex        =   67
      Top             =   5040
      Width           =   345
   End
   Begin VB.Label Label14 
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8205
      TabIndex        =   65
      Top             =   5040
      Width           =   1185
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GLOSA :"
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
      Left            =   8055
      TabIndex        =   64
      Top             =   8040
      Width           =   555
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N.OPERACION :"
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
      Left            =   315
      TabIndex        =   58
      Top             =   4740
      Width           =   1035
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONCEPTO PAGO:"
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
      Left            =   150
      TabIndex        =   56
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO DE FLUJO :"
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
      Left            =   300
      TabIndex        =   55
      Top             =   4320
      Width           =   1050
   End
   Begin VB.Label Label4 
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
      Left            =   645
      TabIndex        =   49
      Top             =   1920
      Width           =   705
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/DNI :"
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
      TabIndex        =   48
      Top             =   1140
      Width           =   705
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
      Left            =   720
      TabIndex        =   47
      Top             =   340
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUENTA ORIGEN:"
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
      Left            =   195
      TabIndex        =   46
      Top             =   3240
      Width           =   1155
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FACTURAS A PAGAR."
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
      Left            =   8280
      TabIndex        =   45
      Top             =   135
      Width           =   1425
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAZON SOCIAL:"
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
      Left            =   315
      TabIndex        =   44
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   795
      TabIndex        =   43
      Top             =   2400
      Width           =   555
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONTO A PAGAR :"
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
      Left            =   105
      TabIndex        =   42
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T.C:"
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
      Left            =   5310
      TabIndex        =   41
      Top             =   2460
      Width           =   285
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMISION:"
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
      Left            =   3180
      TabIndex        =   40
      Top             =   1200
      Width           =   645
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5505
      TabIndex        =   39
      Top             =   1155
      Width           =   525
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONTO TOTAL:"
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
      Left            =   4545
      TabIndex        =   38
      Top             =   2160
      Width           =   1035
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ITF :"
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
      Left            =   5310
      TabIndex        =   37
      Top             =   2880
      Width           =   285
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      BorderStyle     =   6  'Inside Solid
      Height          =   260
      Left            =   8040
      Top             =   80
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmComprasPagos"
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
                x = 1
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

End Sub

Private Sub chk_ajustes_contables_Click()

If Me.chk_ajustes_contables.Value = 1 Then
   Me.frmajustes_contables.Visible = True
Else
   Me.frmajustes_contables.Visible = False
End If

End Sub

Private Sub chk_trabajador_Click()
If Me.chk_trabajador.Value = 1 Then
   Me.frm_trabajores.Visible = True
   Me.cmdAgregarGasto.Visible = False
   
   Call Me.llenar_trabajadores(Me.HfTrabador)
Else
   Me.frm_trabajores.Visible = True
   Me.cmdAgregarGasto.Visible = True
End If
End Sub
Public Sub llenar_trabajadores(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
If Val(Me.txtid_recibo.Text) > 0 Then
    strCadena = "SELECT * FROM view_persona_gasto_cuenta WHERE id_venta='" & Val(Me.txtid_recibo.Text) & "' and ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT * FROM view_persona_gasto_cuenta WHERE dni_save='" & KEY_USUARIO & "' and id_venta='0' and ruc='" & KEY_RUC & "'"
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
            Grilla.ColWidth(1) = 1200
            Grilla.ColWidth(2) = 2200
            Grilla.ColWidth(3) = 1200
            Grilla.ColWidth(4) = 1200
            Grilla.ColWidth(5) = 3500
            
        Next
        cabecera = "ID" & vbTab & "CTA CONTA" & vbTab & "DESCRIPCION" & vbTab & "MONTO" & vbTab & "DNI " & vbTab & "TRABAJADOR "
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        in_monto = 0
        For i = 0 To rst.RecordCount - 1
            in_monto = in_monto + rst("monto")
            Fila = rst("id") & vbTab & rst("cuenta_contable") & vbTab & rst("Descripcion") & vbTab & Format(rst("monto"), "#,##0.00") & vbTab & rst("dni") & vbTab & rst("nombre_completo")
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
            
            Fila = "" & vbTab & "" & vbTab & "ACUMULADO ::::" & vbTab & Format(in_monto, "#,##0.00")
            Grilla.AddItem Fila
            
            Me.txtMontoRedondeo.Text = in_monto
     For k = 3 To 5
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H80FF&
        Next k
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub



Private Sub chkdetraccion_Click()
If Me.chkdetraccion.Value = 1 Then
   Me.frmdetraccion.Enabled = True
   Me.txtprocentajedetraccion.Text = 10
   Me.txtmontodetraccion.Text = Format(Val(Me.txtmontototal.Text) * Val(Me.txtprocentajedetraccion.Text) / 100, "###0.00")
   
   
   
   
   
   strCadena = "SELECT id_compra as Codigo,CONCAT(C.doc_abrev,':',MC.serie,'-',MC.numero) as Descripcion FROM movimiento_compra MC,comprobantes C WHERE MC.seleccion='si' and  MC.id_doc=C.id_doc AND MC.ruc='" & KEY_RUC & "' AND MC.saldo>0 AND MC.anulado='no' AND MC.id_proveedor='" & Trim(Me.txtRuc.Text) & "' ORDER BY MC.fecha_emision ASC  "
   Call ConfiguraRst(strCadena)
   Call LlenaDataCombo(Me.DtcCompra)
   
   
Else
   Me.frmdetraccion.Enabled = False
End If

End Sub

Private Sub ClbAcciones1_HeightChanged(ByVal NewHeight As Single)

End Sub

Private Sub cmdAgregar_Click()
Call agregar_gasto_trabajador
End Sub
Private Sub agregar_gasto_trabajador()
If get_persona(Trim(Me.txtBusqueda_dni.Text)) <> "-" And Val(Me.txtMonto_trabajador.Text) > 0 Then
    
   ' strCadena = "call put_gasto_trabajador('0','" & Val(Me.txtid_recibo.Text) & "','" & Trim(Me.txtBusqueda_dni.Text) & "','" & get_persona(Me.txtBusqueda_dni.Text) & "','" & Val(Me.txtMonto_trabajador.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
   ' CnBd.Execute (strCadena)
    
    strCadena = "call put_gasto_trabajador_gasto('0','" & Val(Me.txtid_recibo.Text) & "','" & Trim(Me.txtCuenta_redondeo.Text) & "','" & Trim(Me.txtBusqueda_dni.Text) & "','" & get_persona(Me.txtBusqueda_dni.Text) & "','" & Val(Me.txtMonto_trabajador.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)

    
    
    Call Me.llenar_trabajadores(Me.HfTrabador)
    Me.txtMonto_trabajador.Text = ""
    Me.lbltrabajador.Caption = ""
    Me.txtBusqueda_dni.Text = ""
    Me.txtMontoRedondeo.Text = ""
    Me.txtCuenta_redondeo.Text = ""
    Call Resalta(Me.txtBusqueda_dni)
    Exit Sub
End If

End Sub


Private Sub cmdAgregarGasto_Click()


strCadena = "call put_gasto_trabajador_gasto('0','" & Val(Me.txtid_recibo.Text) & "','" & Trim(Me.txtCuenta_redondeo.Text) & "','0','-','" & Val(Me.txtMontoRedondeo.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
    
Call Me.llenar_trabajadores(Me.HfTrabador)
Me.txtMontoRedondeo.Text = ""
Me.txtCuenta_redondeo.Text = ""
Me.frm_trabajores.Visible = True



End Sub

Private Sub cmdCargarCheque_Click()
Dim glosa As String
Procedencia = 1
If Val(Me.TxtMontoPago.Text) > 0 Then
        strCadena = "DELETE FROM cheque_detalle WHERE id_cheque='" & Val(Me.DtcCheque.BoundText) & "' AND ruc='" & KEY_RUC & "'"
        Call Execute_Sql(strCadena)
        strCadena = "DELETE FROM cheque_factura WHERE id_cheque='" & Val(Me.DtcCheque.BoundText) & "' AND ruc='" & KEY_RUC & "'"
        Call Execute_Sql(strCadena)
        
        strCadena = "INSERT INTO cheque_detalle(id_cheque,detalle,monto,ruc)VALUES('" & Val(Me.DtcCheque.BoundText) & "','" & Mid(Trim(Me.txtObservacion.Text), 1, Len(Me.txtObservacion.Text) - 1) & "','" & Val(Me.TxtMontoPago.Text) & "','" & KEY_RUC & "')"
        Call CnBd.Execute(strCadena)
        strCadena = "SELECT * FROM movimiento_compra WHERE id_proveedor='" & Me.txtRuc.Text & "' AND ruc='" & KEY_RUC & "' AND saldo>0 AND anulado='no' AND seleccion='si'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            For i = 0 To rst.RecordCount - 1
                strCadena = "INSERT INTO cheque_factura(id_cheque,id_compra,ruc)VALUES('" & Me.DtcCheque.BoundText & "','" & rst("id_compra") & "','" & KEY_RUC & "')"
                Call Execute_Sql(strCadena)
                rst.MoveNext
            Next i
        End If
        FrmChequeNuevo.Txtcentrocosto.Text = "42121"
        FrmChequeNuevo.lblcostos.Text = "FACTURAS POR PAGAR"
        FrmChequeNuevo.txtMotivo.Text = ""
        FrmChequeNuevo.TxtMontoMotivo.Text = ""
        FrmChequeNuevo.txtRuc.Text = Me.txtRuc.Text
        FrmChequeNuevo.txtrazonsocial.Text = Me.TxtCliente.Text
        
        Call Resalta(FrmChequeNuevo.txtMotivo)
        
    End If

'FrmChequeNuevo.Show
'FrmChequeNuevo.TxtidCheque.text = Me.DtcCheque.BoundText
'Call FrmChequeNuevo.llenar_cheque(Me.DtcCheque.BoundText)

End Sub

Private Sub cmdDelete_Click()
If Me.HfTrabador.Rows > 0 Then
    If Val(Me.txtid_recibo.Text) < 1 Then
    strCadena = "call put_gasto_trabajador('" & Val(Me.HfTrabador.TextMatrix(Me.HfTrabador.Row, 0)) & "','" & Val(Me.txtid_recibo.Text) & "','-','-','-','','')"
    CnBd.Execute (strCadena)
Else
    MsgBox "ESTE MOVIMIENTO YA FUE PROCESADO", vbInformation, KEY_USUARIO
    End If
End If
    Call Me.llenar_trabajadores(Me.HfTrabador)
End Sub

Private Sub cmdImprimir_Click()
        

strCadena = "SELECT id_venta,fecha_emision,hora,documento,id_cliente,ncliente,total,forma_pago,flujo,cuenta_origen,observacion,nombre_completo,tc,'" & UCase(EnLetras(Val(Me.TxtMontoPago.Text))) & "',operacion,monto_redondeo,monto_anticipo,cta_redondeo,cta_anticipo,ruc FROM view_compra_pago WHERE id_venta='" & Val(Me.txtid_recibo.Text) & "'"
Call ConfiguraRst(strCadena)


strCadena = "SELECT fecha_emision,id_proveedor,nproveedor,comprobante,tc,id_moneda,monto_inicial,monto_pagado FROM view_compra_pago_detalle WHERE id_venta='" & Val(Me.txtid_recibo.Text) & "'"
Call ConfiguraRstK(strCadena)
Ans = ShowMultiReport(rst, "RptVoucher_pago", , App.Path + "\Reportes\", , , , , rstK, "RptVoucher_pago_detalle")
                        
                        
End Sub

Private Sub cmdProcesar_Click()

If verificar_cierre_caja(Format(Me.DtpEmision.Value, "dd-mm-YYYY")) = 1 Then
    MsgBox "AVISO IMPORTANTE..." + Chr(13) + Chr(13) + "CAJA CONTABLE YA CERRADA.", vbInformation, KEY_VENDEDOR
    Exit Sub
End If


If Val(Me.txtTc.Text) < 1 And Me.DtcMoneda.BoundText = "00002" Then
   MsgBox "Es Necesario INGRESAR UN TIPO CAMBIO " + Chr(13) + "Ingrese al Modulo de Tipo de Cambio en el Menu.", vbInformation, KEY_VENDEDOR
   Exit Sub
End If


If Val(Me.txtTc.Text) < 1 And Me.DtcMoneda.BoundText = "00001" Then
   Me.txtTc.Text = 1
   
End If




Call Save
      
End Sub

Private Sub cmdQuitar_pago_Click()
If Val(Me.MshFacturas.TextMatrix(Me.MshFacturas.Row, 0)) > 0 Then
    strCadena = "UPDATE movimiento_compra SET seleccion='no',dni_save_pago='0' WHERE id_compra='" & Val(Me.MshFacturas.TextMatrix(Me.MshFacturas.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Call llenar_facturas_pagar(Me.MshFacturas, Me.txtRuc.Text)
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub




Private Sub DtcCuentas_Change()
Call load_saldo_comprobante
End Sub

Private Sub load_saldo_comprobante()
Dim ssaldo As Double, residuo As Single, sitf As Single
Dim in_cambio As Single

'TxtMontoPago
in_cambio = Val(Me.txtTc.Text)
If in_cambio = 0 Then
    in_cambio = KEY_CAMBIO
End If
If in_cambio = 0 Then
    in_cambio = 1
End If

strCadena = "SELECT * FROM mis_cuentas WHERE id_cuenta='" & Me.DtcCuentas.BoundText & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)



If KEY_PAIS = KEY_PERU Then
If rst("id_moneda") = "00002" Then
      ssaldo = Val(Me.txtmontototal.Text)
      Me.TxtMontoPago.Text = Format(ssaldo / in_cambio, "###0.0000")
      
Else
     
     If Me.txtMonedaOrigen.Text = "00002" Then
         ssaldo = Val(Me.txtmontototal.Text)
     Else
        ssaldo = Val(Me.txtmontototal.Text)
     End If
     Me.TxtMontoPago.Text = Format(ssaldo, "###0.0000")
End If

Else
     ssaldo = Val(Me.txtmontototal.Text)
     Me.TxtMontoPago.Text = Format(ssaldo, "###0.0000")
End If

residuo = Val(Me.TxtMontoPago.Text) Mod 1000



If (Val(Me.TxtMontoPago.Text) - residuo) > 0 Then
    sitf = (Val(Me.TxtMontoPago.Text) - residuo) * 0.005 / 100
Else
    itf = 0#
End If
If rst("id_tipo") = "01" Then
    FrmCheque.Enabled = False
Else
    FrmCheque.Enabled = True
    Me.TxtItf.Text = Format(sitf, "#,##0.00")
End If
Me.DtcMoneda.BoundText = rst("id_moneda")
 
 
 
 
 strCadena = "SELECT id_registro as Codigo, CONCAT(descripcion,'-',observacion) as Descripcion FROM forma_pago_detalle  WHERE   id_cuenta_caja='" & Me.DtcCuentas.BoundText & "' AND ruc='" & KEY_RUC & "' AND estado='si' ORDER BY id_detalle"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcForma_pago_detalle)

 
 

End Sub

Private Sub DtcCuentas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcFormaPago.SetFocus
End If
End Sub

Private Sub DtcFlujo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtOperacion)
End If
End Sub

Private Sub DtcFormaPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcFlujo.SetFocus
End If
End Sub

Private Sub DtcTipoDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Me.DtcTipoDoc.BoundText) <> "0001" And Trim(Me.DtcTipoDoc.BoundText) <> "0003" Then
        Call Resalta(Me.txtserie)
    Else
        MsgBox "Srta:" + Space(1) + KEY_VENDEDOR + Space(1) + "Solo Salida de Dinero", vbInformation
        Me.DtcTipoDoc.SetFocus
    End If
    
    
End If
End Sub

Public Sub nuevo()
    
End Sub

Private Sub ActualizarAdelanto(ByVal TotalPedido As Double)

End Sub

Private Sub DtcTipoServicio_Change()
Me.txttasa.Text = get_tasa(Me.DtcTiposervicio.BoundText)
End Sub

Private Function get_tasa(ByVal in_codigo As String)

strCadena = "SELECT * FROM tipo_detraccion WHERE codigo='" & in_codigo & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   get_tasa = rstK("tasa")
Else
   get_tasa = 0
End If

End Function

Private Sub DtpEmision_Change()
Me.txtTc.Text = get_tipo_cambio_dia(CVDate(Me.DtpEmision.Value), "valor_venta")


Call put_retencion_visualizar(Me.txtid_compra.Text)

Call put_update_tipo_cambio

DtpValor.Value = Me.DtpEmision.Value


End Sub
Private Sub put_update_tipo_cambio()
    strCadena = "SELECT  (total-function_pago_factura(id_compra,'" & KEY_FECHA & "',id_moneda,ruc)) as saldon,id_compra FROM view_cuentas_cobrar WHERE id_moneda='00002' and  seleccion='si' and dni_save_pago='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       For i = 0 To rst.RecordCount - 1
           If Val(Me.txtTc.Text) <= 0 Then
                in_tc = KEY_CAMBIO
            Else
                in_tc = Val(Me.txtTc.Text)
            End If
            If KEY_PAIS = KEY_PERU Then
                saldon = rst("saldon") * in_tc
            Else
                saldon = rst("saldon")
            End If
            
            strCadena = "UPDATE movimiento_compra SET monto_pagar='" & Val(saldon) & "' WHERE id_compra='" & rst("id_compra") & "' "
            CnBd.Execute (strCadena)
           
           rst.MoveNext
       Next i
       
    End If
    Call Me.llenar_facturas_pagar(Me.MshFacturas, Trim(Me.txtRuc.Text))
        
End Sub

Private Sub put_retencion_visualizar(ByVal in_compra As String)
strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='0002' and  id_compra='" & Val(in_compra) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   If rstLocal("retencion") > 0 And Month(rstLocal("fecha_emision")) <> Month(Me.DtpEmision.Value) Then
              Me.frmretencion.Visible = True
              Me.txtMontoRetencion.Text = Format(rstLocal("retencion"), "###0.00")
   Else
              Me.frmretencion.Visible = False
              Me.txtMontoRetencion.Text = 0#
    End If
   
End If


End Sub


Private Sub Form_Activate()
Call Resalta(Me.TxtMontoPago)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50

Me.cmdCargarCheque.Visible = False
Me.DtpEmision.Value = KEY_FECHA
Me.DtpValor.Value = KEY_FECHA
Me.txtTc.Text = KEY_CAMBIO_VENTA
 

  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  Me.DtcAlmacen.Enabled = False
  
  
  strCadena = "SELECT codigo as Codigo,CONCAT(codigo,'-',descripcion) as Descripcion FROM tipo_detraccion order by codigo ASC"
   Call ConfiguraRst(strCadena)
   Call LlenaDataCombo(Me.DtcTiposervicio)
   
   Me.Dtpfechadetraccion.Value = KEY_FECHA
   
  
  
  strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0097' AND ruc='" & KEY_RUC & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' LIMIT 1"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount < 1 Then
        strCadena = "INSERT INTO almacen_comprobante(ruc,id_alm,id_doc,serie,numero)VALUES ('" & KEY_RUC & "','" & Me.DtcAlmacen.BoundText & "','0097','001','000001')"
        Call Execute_Sql(strCadena)
        strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0097' AND ruc='" & KEY_RUC & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "'"
        Call ConfiguraRst(strCadena)
        Me.txtserie.Text = rst("serie")
        Me.TxtNumeroDoc.Text = rst("numero")
  Else
      Me.txtserie.Text = rst("serie")
      Me.TxtNumeroDoc.Text = rst("numero")
  End If
  
  strCadena = "SELECT A.id_doc as Codigo,C.doc_abrev as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' ORDER BY C.doc_abrev  "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
  Me.DtcTipoDoc.BoundText = "0097"
   
  Me.txtid_compra.Text = FrmListadoFacturasCompra.HfgFacturas.TextMatrix(FrmListadoFacturasCompra.HfgFacturas.Row, 0)
  
 
  
  
  strCadena = "SELECT * FROM movimiento_compra MC,persona P WHERE id_compra='" & Val(Me.txtid_compra.Text) & "' AND ruc='" & KEY_RUC & "' AND id_proveedor=P.dni"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
        Me.txtRuc.Text = rst("id_proveedor")
        Me.TxtCliente.Text = UCase(rst("nproveedor"))
        Me.txtMonedaOrigen.Text = rst("id_moneda")
        If rst("id_doc") = "0002" Then '**** RECIBO X HONORARIOS *****
           If rst("retencion") > 0 And Month(rst("fecha_emision")) <> Month(KEY_FECHA) Then
              Me.frmretencion.Visible = True
              Me.txtMontoRetencion.Text = Format(rst("retencion"), "###0.00")
           End If
        End If
  End If

  
  strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM moneda  ORDER BY id_moneda ASC"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtcMoneda)
  Me.DtcMoneda.BoundText = rst("id_moneda")
  
  
  Me.DtcTipoDoc.Enabled = False
  Me.txtmontototal.Text = Format(rst("total"), "###0.00")
  Me.txtsaldo.Text = Format(rst("saldo"), "###0.00")
  Me.TxtMontoPago.Text = Format(rst("saldo"), "###0.00")
  
  
  
  strCadena = "SELECT id_cuenta as Codigo,cuenta as Descripcion FROM view_mis_cuentas_contable WHERE ruc='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCuentas)
  
  
  
  
  strCadena = "SELECT id_moneda FROM mis_cuentas WHERE id_cuenta='" & DtcCuentas.BoundText & "'"
  Call ConfiguraRstT(strCadena)
  If rstT.RecordCount > 0 Then
     Me.DtcMoneda.BoundText = Trim(Me.txtMonedaOrigen.Text)
  End If
  
  Me.txtTc.Text = KEY_CAMBIO_VENTA
  
   
  
strCadena = "SELECT id as Codigo,Descripcion  as Descripcion FROM vw_mediopago_nombre  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFormaPago)
Me.DtcFormaPago.BoundText = "1CIX000000000174"

strCadena = "SELECT id as Codigo,Nombre  as Descripcion FROM adm_flujocaja ORDER BY Nombre  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFlujo)
Me.DtcFlujo.BoundText = "1CIX000000000017"

'If KEY_CONTABILIDAD = "si" Then
       strCadena = "SELECT id_registro as Codigo, cuenta as Descripcion FROM view_forma_pago_conta  WHERE ruc='" & KEY_RUC & "'"
'    Else
'       strCadena = "SELECT id_registro as Codigo, CONCAT(descripcion,'-',observacion) as Descripcion FROM forma_pago_detalle  WHERE   id_moneda='" & Me.DtcMoneda.BoundText & "' and  id='01' AND ruc='" & KEY_RUC & "' AND estado='si' ORDER BY id_detalle"
'    End If
    
    
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcForma_pago_detalle)
Call Me.llenar_facturas_pagar(Me.MshFacturas, Trim(Me.txtRuc.Text))
Me.txtObservacion.Text = "PAGO:" + FrmListadoFacturasCompra.HfgFacturas.TextMatrix(FrmListadoFacturasCompra.HfgFacturas.Row, 3)

Call load_saldo_comprobante

End Sub
Public Sub llenar_cuentas(ByVal Grilla As MSHFlexGrid, ByVal dni As String)
On Error GoTo salir
strCadena = "SELECT B.id_banco,B.abreviatura,M.descripcion,PB.cuenta,M.id_moneda FROM persona_cuentabancaria PB,banco B,moneda M WHERE PB.id_banco=B.id_banco AND PB.id_moneda=M.id_moneda AND PB.dni='" & dni & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1200
            Grilla.ColWidth(2) = 1000
            Grilla.ColWidth(3) = 2000
            Grilla.ColWidth(4) = 0
        Next
        cabecera = "BCO" & vbTab & "MONEDA" & vbTab & "CUENTA"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_banco") & vbTab & rst("abreviatura") & vbTab & rst("descripcion") & vbTab & rst("cuenta") & vbTab & rst("id_moneda")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
        Next i
      
     
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Public Sub llenar_facturas_pagar(ByVal Grilla As MSHFlexGrid, ByVal ruc As String)
On Error GoTo salir
Dim tSaldoD As Double, tSaldoS As Double, saldoD As Double, saldoS As Double, glosa As String
Dim sldo_soles As Double, sldo_dolar As Double
strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,seleccion,monto_pagar,id_doc FROM view_cuentas_cobrar WHERE  (total-function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc))<>0 and dni_save_pago='" & KEY_USUARIO & "' and seleccion='si' and ruc='" & KEY_RUC & "'"

'strCadena = "SELECT * FROM view_compra_lista WHERE  dni_save_pago='" & KEY_USUARIO & "' and seleccion='si' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Me.cmdProcesar.Enabled = False
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 1100
           Grilla.ColWidth(3) = 2900
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1400
           Grilla.ColWidth(6) = 1400
           Grilla.ColWidth(7) = 1400
           Grilla.ColWidth(8) = 600
           Grilla.ColWidth(9) = 500
        Next
        cabecera = "IDCOMPRA" & vbTab & "EMISION" & vbTab & "VENCIMIENTO" & vbTab & "COMPROBANTE" & vbTab & "MONEDA" & vbTab & "SALDO (US$)" & vbTab & "SALDO (S/.)" & vbTab & "PAGAR [SOLES]" & vbTab & "TC"
        Grilla.AddItem cabecera
         For k = 0 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
        glosa = ""
        rst.MoveFirst
        
        in_total_soles = 0
        in_total_dolar = 0
        in_total_pagar = 0
        
        For i = 0 To rst.RecordCount - 1
            in_tc = Val(Me.txtTc.Text)
            If in_tc = 0 Then
                in_tc = KEY_CAMBIO
            End If
            
            in_pago = rst("pago")
            
            If rst("id_moneda") = "00002" Then
               in_saldo_soles = (rst("total") - in_pago) * in_tc
               in_saldo_dolar = (rst("total") - in_pago)
            Else
               in_saldo_soles = (rst("total") - in_pago)
               in_saldo_dolar = (rst("total") - in_pago) / in_tc
            End If
            in_monto_pagar = rst("monto_pagar")
            
           
            
            
            
            Fila = rst("id_compra") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & Format(rst("fecha_cancelacion"), "dd-mm-YYYY") & vbTab & rst("comprobante") & vbTab & rst("moneda") & vbTab & Format(in_saldo_dolar, "#,##0.00") & vbTab & Format(in_saldo_soles, "#,##0.00") & vbTab & Format(in_monto_pagar, "#,##0.00") & vbTab & Format(rst("tc"), "#,##0.000") & vbTab & Chr(254)
            Grilla.AddItem Fila
            
             
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 9 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
                            
                        End With
            
        
                in_total_soles = in_saldo_soles + in_total_soles
                in_total_dolar = in_saldo_dolar + in_total_dolar
                in_total_pagar = in_monto_pagar + in_total_pagar
                
                
                glosa = rst("comprobante") & Space(1) & glosa
                
                For k = 4 To 9
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
                Next k
            
            Fila = ""
            rst.MoveNext
        Next i
        
        cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "[ SALDO :::]" & vbTab & Format(in_total_dolar, "#,##0.00") & vbTab & Format(in_total_soles, "#,##0.00") & vbTab & Format(in_total_pagar, "#,##0.00")
        Grilla.AddItem cabecera
        
        Me.txtObservacion.Text = "PAGO:" + Space(2) + glosa
        Me.TxtAcumulado_pagar.Text = Format(in_total_pagar, "###0.00000")
        Me.TxtMontoPago.Text = Format(in_total_pagar, "###0.00000")
        
        
        Me.txtmontototal.Text = Format(in_total_pagar, "###0.00000")
        Me.txtMonto_dolares.Text = Format(in_total_dolar, "###0.00000")
        Me.txtsaldo.Text = Format(in_total_pagar, "###0.00000")
          For k = 4 To 9
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H80C0FF
                            Next k
                            
                            Me.cmdProcesar.Enabled = True
    
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub llenarGrid_busqueda(ByVal Grilla As MSHFlexGrid)
'On Error GoTo salir
Dim tTotal As Double, tSaldo As Double, nsaldo As Double
tTotal = 0
tSaldo = 0
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
           Grilla.ColWidth(3) = 2200
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1100
           Grilla.ColWidth(6) = 800
           Grilla.ColWidth(7) = 1200
           Grilla.ColWidth(8) = 1200
           Grilla.ColWidth(9) = 1200
           Grilla.ColWidth(10) = 500
        Next
        cabecera = "IDCOMPRA" & vbTab & "EMISION" & vbTab & "VENCIMIENTO" & vbTab & "COMPROBANTE" & vbTab & "RUC/DNI" & vbTab & "MONEDA" & vbTab & "  TC " & vbTab & "FACTURADO" & vbTab & "SALDO [DOLAR]" & vbTab & "SALDO [SOLES]" & vbTab & "[ ]"
        Grilla.AddItem cabecera
        For k = 0 To 10
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
       Next k
                            
        rst.MoveFirst
        j = 0
        For i = 0 To rst.RecordCount - 1
            in_pago = rst("pago")
            If rst("id_moneda") = "00002" Then
                nsaldo_soles = (rst("total") - in_pago) * rst("tc")
                nsaldo_dolar = rst("total") - in_pago
            Else
               nsaldo_soles = (rst("total") - in_pago)
               nsaldo_dolar = (rst("total") - in_pago) / rst("tc")
                
            End If
            
            
            tSaldo_soles = tSaldo_soles + nsaldo_soles
            tSaldo_dolares = tSaldo_dolares + nsaldo_dolar
            If rst("seleccion") = "si" Then
                in_estado = Chr(254)
            Else
                in_estado = Chr(168)
            End If
            
            
            
            
            Fila = rst("id_compra") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & Format(rst("fecha_cancelacion"), "dd-mm-YYYY") & vbTab & rst("comprobante") & vbTab & rst("id_proveedor") & vbTab & rst("moneda") & vbTab & Format(rst("tc"), "#,##0.0000") & vbTab & Format(rst("total"), "#,##0.00") & vbTab & Format(nsaldo_dolar, "#,##0.00") & vbTab & Format(nsaldo_soles, "#,##0.00") & vbTab & in_estado
            Grilla.AddItem Fila
           
            For k = 8 To 10
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H80C0FF
            Next k
            
            With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 10 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
                            
                        End With
         
            rst.MoveNext
        Next i
        cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(tSaldo_dolares, "#,##0.00") & vbTab & Format(tSaldo_soles, "#,##0.00")
        Grilla.AddItem cabecera
                            For k = 8 To 9
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
                            Next k
    
    
    
 ' Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub

Public Sub LlenarVinculados(ByVal Grilla As MSHFlexGrid, ByVal cPersona As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT id_venta,fecha_vencimiento,CONCAT(C.doc_abrev,':',serie,'-',numero)as comprobante,saldo,seleccion FROM movimiento_venta M,comprobantes C WHERE M.id_doc=C.id_doc AND M.id_cliente='" & cPersona & "' AND M.ruc='" & KEY_RUC & "' AND M.saldo>0 AND M.id_forma_pago='02' AND M.anulado='no'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 2100
           Grilla.ColWidth(3) = 1300
         Next
        cabecera = "IdVenta" & vbTab & "Vencimiento" & vbTab & "Documento" & vbTab & "Saldo"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("id_venta") & vbTab & rst("fecha_vencimiento") & vbTab & rst("comprobante") & vbTab & Format(rst("saldo"), "#,##0.00")
            Grilla.AddItem Fila
            If rst("seleccion") = "si" Then
                For k = 0 To 3
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
                Next k
            End If
            tTotal = tTotal + rst("saldo")
            Fila = ""
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "TOTAL DEUDA:" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
      
            Grilla.col = 3
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  Exit Sub
salir:    MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub
Private Function generar_recibo() As String

                    KEY_VENCIMIENTO = KEY_FECHA
                    id_tipo_factura = "0002"
                    igv = "si"
                    dfac = "no"
                    
                    strCadena = "SELECT numero FROM  movimiento_venta WHERE id_doc='0097' and serie='" & Trim(Me.txtserie.Text) & "'  and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
                    Call ConfiguraRstZ(strCadena)
                    If rstZ.RecordCount > 0 Then
                        Me.TxtNumeroDoc.Text = Format(Val(rstZ("numero")) + 1, "000000")
                    Else
                        Me.TxtNumeroDoc.Text = Format(1, "000000")
                    End If
                    
                    horario = Format(Time, "hh:mm")
                    If horario >= "07:00" And horario <= "13:00" Then
                        turno = "M"
                    Else
                        turno = "T"
                    End If
                    
                    Documento = Trim(Me.DtcTipoDoc.Text) & ":" & Trim(Me.txtserie.Text) & "-" & Trim(Me.TxtNumeroDoc.Text)
                    strCadena = "P_insert_venta('" & Me.DtcTipoDoc.BoundText & "','" & KEY_ALM & "','" & get_forma_pago(Me.DtcCuentas.BoundText) & "','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
                    "'" & Trim(Me.txtserie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Me.txtRuc.Text & "','" & Me.TxtCliente.Text & "','0','0','0','" & Val(Me.TxtMontoPago.Text) & "','0'," & _
                    "'" & Val(Me.TxtMontoPago.Text) & "','0','" & Format(Me.DtpEmision.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpValor.Value, "YYYY-mm-dd") & "','" & id_tipo_factura & "','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & Val(Me.txtTc.Text) & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstP(strCadena)
                    
                    id_venta = rstP(0)
                    Me.txtid_recibo.Text = id_venta
                    If Val(Me.TxtMontoPago.Text) > Val(Me.TxtAcumulado_pagar.Text) Then
                       in_redondeo = 0
                    Else
                        in_redondeo = 0
                    End If
                    
                    strCadena = "UPDATE movimiento_venta SET redondeo='" & Val(Me.txtMontoRedondeo.Text) & "', observacion='" & Trim(Me.txtObservacion.Text) & "',operacion='" & Trim(Me.txtOperacion.Text) & "' WHERE id_venta='" & id_venta & "'"
                    CnBd.Execute (strCadena)
                    
                    
                   ' strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & id_venta & "','00','" & Trim(Me.TxtObservacion.Text) & "','-','1','" & Val(Me.TxtMontoPago.Text) & "','0','" & Val(Me.TxtMontoPago.Text) & "','" & KEY_RUC & "')"
                   ' CnBd.Execute (strCadena)
                               
                   
                    in_redondeo = Val(Me.txtMontoRedondeo.Text)
                   
                    
                    strCadena = "INSERT INTO movimiento_venta_monto(id_venta,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,banco,cheque,cuenta_contable,forma_pago_contable,flujo_caja,id_cuenta_origen,cta_redondeo,cta_anticipo,monto_redondeo,monto_anticipo,ruc)VALUES " & _
                    "('" & id_venta & "','01','" & get_forma_pago_anterior(Me.DtcMoneda.BoundText) & "','" & Val(Me.TxtMontoPago.Text) & "','" & Val(Me.TxtMontoPago.Text) * -1 & "','00','-','" & Trim(Me.txtOperacion.Text) & "','-','" & Me.DtcCheque.BoundText & "','" & get_cuenta_contable_cuenta(Me.DtcCuentas.BoundText) & "','" & DtcFormaPago.BoundText & "','" & Me.DtcFlujo.BoundText & "','" & Me.DtcCuentas.BoundText & "','" & Trim(Me.txtCuenta_redondeo.Text) & "','" & Trim(Me.txtCuenta_anticipo.Text) & "','" & in_redondeo & "','" & Val(Me.txtMontoAnticipo.Text) & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    
                    
                    strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(Me.TxtNumeroDoc.Text + 1), "000000") & "' WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.txtserie.Text) & "' AND ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    
                    generar_recibo = "RECIBO EGRESO:" & Trim(Me.txtserie.Text) & "-" & Trim(Me.TxtNumeroDoc.Text)
                    
End Function
Private Sub put_gasto_trabajador()
If Me.chk_ajustes_contables.Value = 1 And Me.chk_trabajador.Value = 1 Then
    strCadena = "UPDATE persona_gasto SET id_venta='" & Val(Me.txtid_recibo.Text) & "' WHERE id_venta='0' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
End If
End Sub
Private Function get_acumulado_nota() As Double
strCadena = "SELECT sum((total-function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc))) as saldo FROM view_cuentas_cobrar WHERE id_doc='0007' and dni_save_pago='" & KEY_USUARIO & "' and seleccion='si' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstlocal(strCadena)
If IsNull(rstLocal(0)) = True Then
    get_acumulado_nota = 0
Else
    get_acumulado_nota = Abs(rstLocal(0))
End If


End Function
Private Sub Save()
Dim in_recibo As String
Dim monto_pago As Double, saldof As Double, comprobante As String, monto_pagado As Double, Saldo As Double, id_moneda As String
Dim in_registro As Double
Dim in_tipo As String
Dim monto_nota As Double
Dim in_monto_redondeo As Single
Dim in_monto_retencion As Single

If KEY_CONTABILIDAD = "si" Then
    If Trim(Me.txtCuenta_redondeo.Text) <> "" Then
        If get_validar_cuenta_asociada(Trim(Me.txtCuenta_redondeo.Text)) = False Then
            MsgBox "ATENCION." + Chr(13) + "LA CUENTA :" & Trim(Me.txtCuenta_redondeo.Text) & Space(2) & "::: NO TIENE CUENTA ASOCIADA", vbInformation, KEY_VENDEDOR
            Exit Sub
        End If
    End If
End If

If Me.frmretencion.Visible = True And Trim(Me.txtCtaRetencion.Text) <> "" Then
            in_monto_retencion = Val(Me.txtMontoRetencion.Text)
        Else
            in_monto_retencion = 0
        End If
        
Me.TxtMontoPago.Text = Val(Me.TxtMontoPago.Text) - in_monto_retencion

monto_pago = Val(Me.TxtMontoPago.Text)
saldof = 0

        
        in_recibo = generar_recibo
        
        strCadena = "call put_vincular_persona_gasto('" & Val(Me.txtid_recibo.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
       ' Call put_gasto_trabajador ' actualizar gasto a trabajador
        
        
        
        monto_pago = Val(Me.TxtMontoPago.Text) + get_acumulado_nota
        in_monto_redondeo = Val(Me.txtMontoRedondeo.Text)
        
        If in_monto_redondeo > 0 Then
            If Mid(Trim(Me.txtCuenta_redondeo.Text), 1, 1) = "6" Then
               monto_pago = monto_pago
            Else
               monto_pago = monto_pago + in_monto_redondeo
            End If
        End If
        
        
        If Me.OptChequeSi.Value = False Then
                    Call put_detraccion(Me.txtid_compra.Text, Trim(Me.txtconstancia.Text), Me.DtcTiposervicio.BoundText, Format(Me.Dtpfechadetraccion.Value, "YYYY-mm-dd"))
                    
                    
                    
                    strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,seleccion,monto_pagar FROM view_cuentas_cobrar WHERE id_doc<>'0007' and  dni_save_pago='" & KEY_USUARIO & "' and seleccion='si' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRst(strCadena)
                    If rst.RecordCount > 0 Then
                        rst.MoveFirst
                        
                        For i = 0 To rst.RecordCount - 1
                          '**** Saldo Inicial
                          
                          
                          If KEY_PAIS = KEY_PERU Then
                          If rst("id_moneda") = "00002" Then
                              If rst("id_moneda") = Me.DtcMoneda.BoundText Then
                                 in_saldo = rst("monto_pagar") / Val(Me.txtTc.Text)
                              Else
                                
                                    in_saldo = rst("monto_pagar") / Val(Me.txtTc.Text)
                                    in_saldo = in_saldo * Val(Me.txtTc.Text)
                               End If
                         Else
                            in_saldo = rst("monto_pagar")
                         End If
                         Else
                            in_saldo = rst("monto_pagar")
                         End If
                           
                           
                           
                           '*****************
                           
                           in_saldo_inicial = in_saldo
                           saldof = in_saldo
                           
                          
                           If Val(monto_pago) > Val(saldof) Then
                               
                               If KEY_ASIENTO_GLOBAL_CTA_PAGAR = "no" Then
                                    in_codigo_det = procesar_transaccion(KEY_ALM, Me.DtcCuentas.BoundText, Me.DtpEmision.Value, "00002", rst("id_proveedor"), rst("nproveedor"), in_glosa, saldof, "0", "0", rst("id_compra"), rst("comprobante"), Val(Me.txtTc.Text), Trim(Me.txtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                               End If
                                    
                                strCadena = "CALL p_insert_pago_factura_vitekey('" & Val(Me.txtid_recibo.Text) & "','" & rst("id_compra") & "','" & in_saldo_inicial & "','" & saldof & "','" & rst("id_moneda") & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtTc.Text) & "','" & Val(in_codigo_det) & "')"
                                CnBd.Execute (strCadena)
                                
                                strCadena = "CALL p_insert_pago_factura_vitekey('" & rst("id_compra") & "','" & Val(Me.txtid_recibo.Text) & "','" & in_saldo_inicial & "','" & saldof & "','" & rst("id_moneda") & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtTc.Text) & "','" & Val(in_codigo_det) & "')"
                                CnBd.Execute (strCadena)
                                
                                
                                    
                                    
                                    monto_pago = monto_pago - saldof
                                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & Val(Me.txtid_recibo.Text) & "','00','" & Trim(rst("comprobante")) & "','-','1','" & saldof & "','0','" & saldof & "','" & KEY_RUC & "')"
                                    CnBd.Execute (strCadena)
                                    
                            Else
                                
                                in_glosa = "[" & UCase(Trim(Me.txtObservacion.Text)) & "]"
                                
                                If KEY_ASIENTO_GLOBAL_CTA_PAGAR = "no" Then
                                    in_codigo_det = procesar_transaccion(KEY_ALM, Me.DtcCuentas.BoundText, Me.DtpEmision.Value, "00002", rst("id_proveedor"), rst("nproveedor"), in_glosa, monto_pago, "0", "0", rst("id_compra"), rst("comprobante"), Val(Me.txtTc.Text), Trim(Me.txtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                                End If
                                in_monto_nota = 0
                                
                                
                                If Mid(Trim(Me.txtCuenta_redondeo.Text), 1, 1) = "6" Then
                                    in_monto_redondeo = 0
                                Else
                                    monto_pago = monto_pago - in_monto_redondeo
                                    in_monto_redondeo = in_monto_redondeo
                                End If
                                
                                strCadena = "CALL p_insert_pago_factura_vitekey('" & Val(Me.txtid_recibo.Text) & "','" & rst("id_compra") & "','" & in_saldo_inicial & "','" & monto_pago & "','" & rst("id_moneda") & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtTc.Text) & "','" & Val(in_codigo_det) & "')"
                                CnBd.Execute (strCadena)
                                
                                strCadena = "CALL p_insert_pago_factura_vitekey('" & rst("id_compra") & "','" & Val(Me.txtid_recibo.Text) & "','" & in_saldo_inicial & "','" & monto_pago & "','" & rst("id_moneda") & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtTc.Text) & "','" & Val(in_codigo_det) & "')"
                                CnBd.Execute (strCadena)
                                
                                
                                If in_monto_redondeo <> 0 Then
                                    strCadena = "CALL p_insert_pago_factura_ultimate_tipo('" & Val(Me.txtid_recibo.Text) & "','" & rst("id_compra") & "','" & Abs(in_monto_redondeo) & "','" & Abs(in_monto_redondeo) & "','00')"
                                    CnBd.Execute (strCadena)
                                End If
                                
                                strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & Val(Me.txtid_recibo.Text) & "','00','" & Trim(rst("comprobante")) & "','-','1','" & monto_pago & "','0','" & monto_pago & "','" & KEY_RUC & "')"
                                CnBd.Execute (strCadena)
                                monto_pago = 0
                               
                            End If
                          
                        rst.MoveNext
                        Next i
                        
                        'CANCELACION CON NOTAS DE CREDITO
                        
                        Call put_cancelar_con_notas(Me.txtid_recibo.Text)
                        
siguiente:
                          
                          If Trim(Me.txtCuenta_redondeo.Text) <> "" And Val(Me.txtMontoRedondeo.Text) <> 0 Then
                           If Mid(Trim(Me.txtCuenta_redondeo.Text), 1, 1) = "6" Then
                            strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & Val(Me.txtid_recibo.Text) & "','00','REDONDEO','-','1','" & Val(Me.txtMontoRedondeo.Text) * -1 & "','0','" & Val(Me.txtMontoRedondeo.Text) * -1 & "','" & KEY_RUC & "')"
                            CnBd.Execute (strCadena)
                            in_total = Val(Me.TxtMontoPago.Text)
                           Else
                            strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & Val(Me.txtid_recibo.Text) & "','00','REDONDEO','-','1','" & Val(Me.txtMontoRedondeo.Text) & "','0','" & Val(Me.txtMontoRedondeo.Text) & "','" & KEY_RUC & "')"
                            CnBd.Execute (strCadena)
                            in_total = Val(Me.TxtMontoPago.Text) + Val(Me.txtMontoRedondeo.Text)
                           End If
                           
                           
                        End If
                        If Trim(Me.txtCuenta_anticipo.Text) <> "" And Val(Me.txtMontoAnticipo.Text) <> 0 Then
                            strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & Val(Me.txtid_recibo.Text) & "','00','ANTICIPO','-','1','" & Val(Me.txtMontoAnticipo.Text) & "','0','" & Val(Me.txtMontoAnticipo.Text) & "','" & KEY_RUC & "')"
                            CnBd.Execute (strCadena)
                            in_total = Val(Me.TxtMontoPago.Text) + Val(Me.txtMontoAnticipo.Text) + Val(Me.txtMontoRedondeo.Text)
                        End If
                        
                          If Val(Me.txtMontoAnticipo.Text) <> 0 Or Val(Me.txtMontoRedondeo.Text) <> 0 Then
                          strCadena = "UPDATE movimiento_venta SET total='" & Val(in_total) & "' WHERE id_venta='" & Val(Me.txtid_recibo.Text) & "'"
                          CnBd.Execute (strCadena)
                          
                           
                          
                          strCadena = "UPDATE movimiento_venta_monto SET monto='" & Val(in_total) & "',monto_caja='" & Val(in_total) * -1 & "' WHERE id_venta='" & Val(Me.txtid_recibo.Text) & "'"
                          CnBd.Execute (strCadena)
                          End If
                          
                          If KEY_ASIENTO_GLOBAL_CTA_PAGAR = "si" Then
                            
                            strCadena = "call CON_InsertaAsiento_PagoGlobal('" & Val(Me.txtid_recibo.Text) & "')"
                            CnBd.Execute (strCadena)
                            
                            If Me.frmretencion.Visible = True And Val(Me.txtMontoRetencion.Text) > 0 Then
                                in_retencion = generar_recibo_retencion(Me.DtpEmision.Value, Trim(Me.txtRuc.Text), Val(Me.txtMontoRetencion.Text), Val(Me.txtTc.Text), Me.DtcMoneda.BoundText, Trim(Me.txtOperacion.Text), Trim(Me.txtCtaRetencion.Text), Val(Me.txtid_compra.Text), 0)
                                strCadena = "call CON_InsertaAsiento_PagoRetencion('" & Val(in_retencion) & "')"
                                CnBd.Execute (strCadena)
                            End If
                            
                            If Me.DtcMoneda.BoundText = "00002" Then
                                strCadena = "call CON_AjusteTC_Global('" & Val(Me.txtid_recibo.Text) & "')"
                                CnBd.Execute (strCadena)
                            End If
                          in_monto_banco = Val(Me.TxtMontoPago.Text) + Val(Me.txtMontoAnticipo.Text) + Val(Me.txtMontoRedondeo.Text)
                          Call procesar_transaccion_egreso(Me.DtcForma_pago_detalle.BoundText, Me.DtcAlmacen.BoundText, Me.DtcCuentas.BoundText, Me.DtpEmision.Value, "00002", Trim(Me.txtRuc.Text), Trim(Me.TxtCliente.Text), Trim(Me.txtObservacion.Text), in_monto_banco, "", 0, Val(txtid_recibo.Text), in_recibo, Val(Me.txtTc.Text), Trim(Me.txtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, KEY_USUARIO, KEY_RUC)
                        End If
                    End If
        End If
        nuevo_numero = formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6)
        strCadena = "UPDATE  almacen_comprobante SET numero='" & Trim(nuevo_numero) & "'  WHERE (serie='" & Trim(Me.txtserie.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "')"
        Call Execute_Sql(strCadena)
        Me.cmdImprimir.Enabled = True
        Me.cmdProcesar.Enabled = False
        
        

        Exit Sub
    

End Sub
Private Sub put_cancelar_con_notas(ByVal in_recibo As String)
Dim in_diferencia As Double
strCadena = "SELECT (total-function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc)) as saldo,id_compra as id_nota,tc,id_moneda,comprobante FROM view_cuentas_cobrar WHERE id_doc='0007' and dni_save_pago='" & KEY_USUARIO & "' and seleccion='si' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   rstLocal.MoveFirst
   For j = 0 To rstLocal.RecordCount - 1
        If IsNull(rstLocal("saldo")) = True Then
            monto_pago = 0
        Else
            monto_pago = Abs(rstLocal("saldo"))
        End If
        
                   
                   strCadena = "CALL p_insert_pago_factura_ultimate_premiun('" & in_recibo & "','" & rstLocal("id_nota") & "','" & rstLocal("saldo") & "','0','" & monto_pago * -1 & "','" & rstLocal("id_moneda") & "','" & rstLocal("id_moneda") & "','" & rstLocal("tc") & "')"
                   CnBd.Execute (strCadena)
                   strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & Val(Me.txtid_recibo.Text) & "','00','" & Trim(rstLocal("comprobante")) & "','-','1','" & monto_pago * -1 & "','0','" & monto_pago * -1 & "','" & KEY_RUC & "')"
                                CnBd.Execute (strCadena)
                   
                   rstLocal.MoveNext
    Next j
End If


Exit Sub


strCadena = "SELECT (total-function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc)) as saldo,id_compra as id_nota,tc,id_moneda FROM view_cuentas_cobrar WHERE id_doc='0007' and dni_save_pago='" & KEY_USUARIO & "' and seleccion='si' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   rstLocal.MoveFirst
   For j = 0 To rstLocal.RecordCount - 1
        If IsNull(rstLocal("saldo")) = True Then
            monto_pago = 0
        Else
            monto_pago = Abs(rstLocal("saldo"))
        End If
        strCadena = "SELECT (monto_pagado+monto_nota) as acumulado,monto_inicial,id_movimiento,id FROM mis_cuentas_det_detalle WHERE monto_inicial>(monto_pagado+monto_nota) and  id_detalle='" & Val(in_recibo) & "' and id_tipo='02' ORDER BY monto_nota ASC"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
           rstK.MoveFirst
           For i = 0 To rstK.RecordCount - 1
                If monto_pago > 0 Then
                   in_diferencia = rstK("monto_inicial") - rstK("acumulado")
                   
                   If monto_pago > in_diferencia Then
                      monto_nota = in_diferencia
                   Else
                      monto_nota = monto_pago
                   End If
                   
                   strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & in_recibo & "','" & rstLocal("id_nota") & "','" & rstLocal("saldo") & "','" & monto_nota * -1 & "','" & rstLocal("id_moneda") & "','" & rstLocal("id_moneda") & "','" & rstLocal("tc") & "')"
                   CnBd.Execute (strCadena)
                   
                   strCadena = "CALL p_insert_pago_factura_ultimate_tipo('" & in_recibo & "','" & rstK("id_movimiento") & "','" & rstLocal("saldo") & "','" & monto_nota & "','00')"
                   CnBd.Execute (strCadena)
                   
                   strCadena = "UPDATE mis_cuentas_det_detalle SET monto_nota=monto_nota+'" & monto_nota & "' WHERE id='" & rstK("id") & "'"
                   CnBd.Execute (strCadena)
                   monto_pago = monto_pago - monto_nota
                   If monto_pago = 0 Then
                      Exit For
                   End If
                   
               
             
                 
                End If
                  rstK.MoveNext
           Next i
        End If
        rstLocal.MoveNext
    Next j
End If


Exit Sub


strCadena = "SELECT (total-function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc)) as saldo,id_compra as id_nota FROM view_cuentas_cobrar WHERE id_doc='0007' and dni_save_pago='" & KEY_USUARIO & "' and seleccion='si' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   rstLocal.MoveFirst
   For j = 0 To rstLocal.RecordCount - 1
   If IsNull(rstLocal("saldo")) = True Then
        monto_pago = 0
   Else
        monto_pago = Abs(rstLocal("saldo"))
        strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`nombre_completo`,`id_alm`,`ruc`, (total-function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc)) as saldo,seleccion,monto_pagar FROM view_cuentas_cobrar WHERE id_doc<>'0007' and   (total-function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc))>0 and  dni_save_pago='" & KEY_USUARIO & "' and seleccion='si' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRst(strCadena)
                    If rst.RecordCount > 0 Then
                        rst.MoveFirst
                        
                        For i = 0 To rst.RecordCount - 1
                          '**** Saldo Inicial
                          If rst("id_moneda") = "00002" Then
                              If rst("id_moneda") = Me.DtcMoneda.BoundText Then
                                 in_saldo = rst("saldo") / rst("tc")
                              Else
                                
                                    in_saldo = rst("saldo") / rst("tc")
                                    in_saldo = in_saldo * Val(Me.txtTc.Text)
                               End If
                         Else
                            in_saldo = rst("saldo")
                         End If
                           '*****************
                           
                           in_saldo_inicial = in_saldo
                           saldof = in_saldo
                           
                               
                           If Val(monto_pago) > Val(saldof) Then
                               
                               If KEY_ASIENTO_GLOBAL_CTA_PAGAR = "no" Then
                                    Call procesar_transaccion(KEY_ALM, Me.DtcCuentas.BoundText, Me.DtpEmision.Value, "00002", rst("id_proveedor"), rst("nproveedor"), in_glosa, saldof, "0", "0", rst("id_compra"), rst("comprobante"), Val(Me.txtTc.Text), Trim(Me.txtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                               End If
                                    
                                    strCadena = "CALL p_insert_pago_factura_ultimate_premiun('" & Val(Me.txtid_recibo.Text) & "','" & rst("id_compra") & "','" & in_saldo_inicial & "','" & saldof & "','" & in_monto_nota & "','" & rst("id_moneda") & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtTc.Text) & "')"
                                    CnBd.Execute (strCadena)
                                    
                                    strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & rst("id_compra") & "','" & Val(Me.txtid_recibo.Text) & "','" & in_saldo_inicial & "','" & saldof & "','" & rst("id_moneda") & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtTc.Text) & "')"
                                    CnBd.Execute (strCadena)
                                    
                                    monto_pago = monto_pago - saldof
                                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & Val(Me.txtid_recibo.Text) & "','00','" & Trim(rst("comprobante")) & "','-','1','" & saldof & "','0','" & saldof & "','" & KEY_RUC & "')"
                                    CnBd.Execute (strCadena)
                                    
                            Else
                                
                                in_glosa = "[" & UCase(Trim(Me.txtObservacion.Text)) & "]"
                                
                                If KEY_ASIENTO_GLOBAL_CTA_PAGAR = "no" Then
                                    Call procesar_transaccion(KEY_ALM, Me.DtcCuentas.BoundText, Me.DtpEmision.Value, "00002", rst("id_proveedor"), rst("nproveedor"), in_glosa, monto_pago, "0", "0", rst("id_compra"), rst("comprobante"), Val(Me.txtTc.Text), Trim(Me.txtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                                End If
                                
                                
                                strCadena = "SELECT * FROM mis_cuentas_det_detalle WHERE id_movimiento='" & rst("id_compra") & "' and id_detalle='" & Val(Me.txtid_recibo.Text) & "'"
                                Call ConfiguraRstPP(strCadena)
                                If rstPP.RecordCount > 0 Then
                                    strCadena = "UPDATE mis_cuentas_det_detalle SET monto_nota='" & monto_pago & "' WHERE id='" & rstPP("id") & "'"
                                    CnBd.Execute (strCadena)
                                End If
                                
                                in_recibo = generar_recibo_egreso(Me.DtpEmision.Value, Me.txtRuc.Text, monto_pago, Val(Me.txtTc.Text), Me.DtcMoneda.BoundText, Trim(Me.txtOperacion.Text), Me.DtcForma_pago_detalle.BoundText, rst("id_compra"), rstLocal("id_nota"))
                                
                               ' strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & rstLocal("id_nota") & "','" & rst("id_compra") & "','" & in_saldo_inicial & "','" & monto_pago & "','" & rst("id_moneda") & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.TxtTc.Text) & "')"
                               ' CnBd.Execute (strCadena)
                                
                                
                                
                               ' strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & rst("id_compra") & "','" & rstLocal("id_nota") & "','" & in_saldo_inicial & "','" & monto_pago & "','" & rst("id_moneda") & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.TxtTc.Text) & "')"
                               ' CnBd.Execute (strCadena)
                                
                                strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & Val(Me.txtid_recibo.Text) & "','00','" & Trim(rst("comprobante")) & "','-','1','" & in_saldo_inicial & "','0','" & in_saldo_inicial & "','" & KEY_RUC & "')"
                                CnBd.Execute (strCadena)
                                monto_pago = 0
                                
                            End If
                        rst.MoveNext
                        Next i
End If
        
        
        
        
   End If
   rstLocal.MoveNext
   Next j

End If
                    
                    
                    
                    
                    
                    
                    
End Sub

Private Sub put_detraccion(ByVal in_compra As String, ByVal in_constancia As String, ByVal in_tipo_servicio As String, ByVal in_fecha As String)
If Me.chkdetraccion.Value = 1 Then
            strCadena = "SELECT * FROM movimiento_compra_detraccion WHERE id_compra='" & Val(in_compra) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                strCadena = "p_insert_detraccion('" & in_compra & "','" & Trim(Me.txtnumerodetraccion.Text) & "','" & Val(Me.txtmontodetraccion.Text) & "','--','" & in_tipo_servicio & "','" & in_constancia & "','" & s & "')"
                Call Execute_Sql(strCadena)
            End If
        End If
End Sub

Private Sub HfLista_DblClick()
If Me.HfLista.TextMatrix(Me.HfLista.Row, 0) > 0 Then
    strCadena = "SELECT  (total-function_pago_factura(id_compra,'" & KEY_FECHA & "',id_moneda,ruc)) as saldon,id_moneda,tc,id_doc FROM view_cuentas_cobrar WHERE id_compra='" & Val(Me.HfLista.TextMatrix(Me.HfLista.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If rst("tc") <= 0 Then
            in_tc = KEY_CAMBIO
        Else
            in_tc = Val(Me.txtTc.Text)
        End If
    
        If rst("id_moneda") = "00002" Then
            If KEY_PAIS = KEY_PERU Then
                saldon = rst("saldon") * in_tc
            Else
                saldon = rst("saldon")
            End If
            
        Else
            saldon = rst("saldon")
        End If
      
      strCadena = "UPDATE movimiento_compra SET seleccion='si',dni_save_pago='" & KEY_USUARIO & "',monto_pagar='" & Val(saldon) & "' WHERE id_compra='" & Val(Me.HfLista.TextMatrix(Me.HfLista.Row, 0)) & "' "
      CnBd.Execute (strCadena)
     
     
     If Me.HfLista.TextMatrix(Me.HfLista.Row, 10) = Chr(254) Then
        Me.HfLista.TextMatrix(Me.HfLista.Row, 10) = Chr(168)
     Else
        Me.HfLista.TextMatrix(Me.HfLista.Row, 10) = Chr(254)
     End If
 
 
   End If


    
    
    Call llenar_facturas_pagar(Me.MshFacturas, Me.txtRuc.Text)
End If
End Sub





Private Sub Image1_Click()
Me.frmmonto_pagar.Visible = False
End Sub

Private Sub Image2_Click()
Me.frm_trabajores.Visible = False
End Sub

Private Sub MshFacturas_DblClick()

If Val(Me.MshFacturas.TextMatrix(Me.MshFacturas.Row, 0)) > 0 Then
   strCadena = "SELECT  (total-function_pago_factura(id_compra,'" & KEY_FECHA & "',id_moneda,ruc)) as saldon FROM view_cuentas_cobrar WHERE id_compra='" & Val(Me.MshFacturas.TextMatrix(Me.MshFacturas.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      Me.frmmonto_pagar.Visible = True
      Me.lblid_compra.Caption = Val(Me.MshFacturas.TextMatrix(Me.MshFacturas.Row, 0))
      Me.txtMontoPagar.Text = rst("saldon")
      Call Resalta(Me.txtMontoPagar)
      Exit Sub
   End If
End If

End Sub

Private Sub OptChequeNO_Click()
Me.DtcCheque.Visible = False
Me.cmdCargarCheque.Visible = False
End Sub

Private Sub OptChequeSi_Click()
Me.DtcCheque.Visible = True
strCadena = "SELECT C.id_cheque as Codigo,CONCAT('CHEQUE ',':',C.numero) as Descripcion FROM cheque C,mis_cuentas M,chequera CH  WHERE C.id_chequera=CH.id_chequera AND CH.id_cuenta=M.id_cuenta AND C.seleccionado='si' AND C.ruc='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' AND CH.ruc='" & KEY_RUC & "' AND M.id_cuenta='" & Me.DtcCuentas.BoundText & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
     Call LlenaDataCombo(Me.DtcCheque)
     Me.DtcCheque.SetFocus
     Me.cmdCargarCheque.Visible = True
Else
    Me.cmdCargarCheque.Visible = False
End If
End Sub





Private Sub Imprimir(ByVal TipoDoc As String, ByVal CodAlm As String, ByVal serie As String, ByVal numero As String)
Dim i As Integer, j As Integer
Dim totalletras As String

    Printer.ScaleMode = vbCharacters 'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = 10
       
If Me.DtcTipoDoc.BoundText = KEY_SALDINER Then
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.PaperSize = 1
    Printer.Print ""
    Printer.Print ""
 '   Printer.Print Tab(15); (CVDate(Me.DtpActual.Value))
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); Mid(Me.TxtCliente.Text + Space(80), 1, 65)
    ' Printer.Print Tab(5); Mid(Me.TxtRuc.Text + Space(50), 1, 40) & "SALDINER"; Space(1); Mid(Me.TxtSerie.Text + Space(50), 1, 4) & Space(1) & "-" & Me.TxtNumeroDoc.Text
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); Mid(Trim(KEY_VENDEDOR), 1, 10)
    Printer.Print ""
    Printer.Print ""
   ' Printer.Print Tab(15); "Monto Efectivo:" & "=============" & Space(20) & Me.TxtMontoIngresar.text
    Printer.CurrentY = Printer.CurrentY + 10
   ' totalletras = UCase(EnLetras(Me.TxtMontoIngresar.text))
    Set rst = Nothing
    '---- fin totales
    Printer.Print Tab(10); Mid(totalletras + Space(100), 1, 60)
    Printer.CurrentY = Printer.CurrentY + 0.2
   ' Printer.Print Tab(60); Me.TxtMontoIngresar.text
    Printer.EndDoc
    
    Exit Sub
End If
End Sub



Private Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub




Private Sub txtBuscarForma_Change()






If KEY_CONTABILIDAD = "si" Then
       strCadena = "SELECT id_registro as Codigo, cuenta as Descripcion FROM view_forma_pago_conta  WHERE   cuenta LIKE '%" & Trim(Me.txtBuscarForma.Text) & "%' and  ruc='" & KEY_RUC & "'"
    Else
       strCadena = "SELECT id_registro as Codigo, CONCAT(descripcion,'-',observacion) as Descripcion FROM forma_pago_detalle  WHERE observacion LIKE '%" & Trim(Me.txtBuscarForma.Text) & "%' and  id_moneda='" & Me.DtcMoneda.BoundText & "' and  id='01' AND ruc='" & KEY_RUC & "' AND estado='si' ORDER BY id_detalle"
    End If
    
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcForma_pago_detalle)

End Sub

Private Sub txtbuscar_comprobante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,seleccion,id_doc FROM view_cuentas_cobrar WHERE comprobante LIKE '%" & Trim(Me.txtbuscar_comprobante.Text) & "%' and dni_save_pago<>'" & KEY_USUARIO & "'  and ruc='" & KEY_RUC & "'"
    'strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,seleccion FROM view_cuentas_pagar_fecha WHERE saldo>0 and  comprobante LIKE '%" & Trim(Me.txtbuscar_comprobante.Text) & "%' and dni_save_pago<>'" & KEY_USUARIO & "'  and ruc='" & KEY_RUC & "'"
    Call llenarGrid_busqueda(Me.HfLista)
End If
End Sub

Private Sub txtBuscar_proveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`,  function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,seleccion,id_doc FROM view_cuentas_cobrar WHERE (total-function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc))>0 and    nproveedor LIKE '%" & Trim(Me.txtBuscar_proveedor.Text) & "%' and dni_save_pago<>'" & KEY_USUARIO & "'  and ruc='" & KEY_RUC & "'"
    'strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,seleccion FROM view_cuentas_pagar_fecha WHERE saldo>0 and  nproveedor LIKE '%" & Trim(Me.txtBuscar_proveedor.Text) & "%' and dni_save_pago<>'" & KEY_USUARIO & "'  and ruc='" & KEY_RUC & "'"
    Call llenarGrid_busqueda(Me.HfLista)
    
    
    
    
    
End If
End Sub

Private Sub txtbuscar_ruc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,seleccion,id_doc FROM view_cuentas_cobrar WHERE  (total-function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc))>0 and  id_proveedor LIKE '%" & Trim(Me.txtbuscar_ruc.Text) & "%' and dni_save_pago<>'" & KEY_USUARIO & "'  and ruc='" & KEY_RUC & "'"
    'strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,seleccion FROM view_cuentas_pagar_fecha WHERE saldo>0 and  id_proveedor LIKE '%" & Trim(Me.txtbuscar_ruc.Text) & "%' and dni_save_pago<>'" & KEY_USUARIO & "'  and ruc='" & KEY_RUC & "'"
    Call llenarGrid_busqueda(Me.HfLista)

End If
End Sub



Private Sub txtBusqueda_dni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Procedencia = seleccionar_per
   frmpersonal.Show
   Exit Sub
End If
End Sub

Private Sub txtCtaRetencion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Procedencia = seleccionar_insumo
   FrmPlanContableCuentas.Show
   Exit Sub
   
End If
End Sub

Private Sub txtCuenta_anticipo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   Procedencia = seleccionar_otro
   FrmPlanContableCuentas.Show
   Exit Sub
End If
End Sub

Private Sub txtCuenta_redondeo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Procedencia = Selecionar
   FrmPlanContableCuentas.Show
   Exit Sub
   
End If
End Sub

Private Sub txtFormaPago_Change()
strCadena = "SELECT id as Codigo,Descripcion  as Descripcion FROM vw_mediopago_nombre WHERE Descripcion LIKE '%" & Trim(Me.txtFormaPago.Text) & "%'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFormaPago)

End Sub

Private Sub txtMonto_trabajador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call agregar_gasto_trabajador
End If
End Sub

Private Sub txtMontoPagar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim in_monton As Double
    strCadena = "SELECT * FROM movimiento_compra WHERE  id_compra='" & Val(Me.lblid_compra.Caption) & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
        If rstZ("id_moneda") = "00002" Then
          in_monton = Val(Me.txtMontoPagar.Text) * Val(Me.txtTc.Text)
        Else
            in_monton = Val(Me.txtMontoPagar.Text)
        End If
        
        strCadena = "UPDATE movimiento_compra SET monto_pagar='" & Val(in_monton) & "' WHERE id_compra='" & Val(Me.lblid_compra.Caption) & "'"
        CnBd.Execute (strCadena)
        Me.frmmonto_pagar.Visible = False
        Call llenar_facturas_pagar(Me.MshFacturas, Me.txtRuc.Text)
    End If
End If
End Sub

Private Sub TxtMontoPago_Change()
Dim sitf As Single, residuo As Single
residuo = Val(Me.TxtMontoPago.Text) Mod 1000

If (Val(Me.TxtMontoPago.Text) - residuo) > 0 Then
    sitf = (Val(Me.TxtMontoPago.Text) - residuo) * 0.005 / 100
Else
    sitf = 0#
End If
Me.TxtItf.Text = Format(sitf, "#,##0.00")
End Sub

Private Sub TxtMontoPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Me.DtcCuentas.SetFocus
End If
End Sub

Private Sub TxtNumeroDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
     
    Me.TxtNumeroDoc.Text = FormatosCeros(Me.TxtNumeroDoc.Text, 6)
    
        
        
    End If

End Sub

Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Call Resalta(Me.TxtMontoIngresar)
End If
End Sub

Private Sub TxtOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtObservacion)
End If
End Sub

Private Sub txtprocentajedetraccion_Change()
Me.txtmontodetraccion.Text = Format(Val(Me.txtmontototal.Text) * Val(Me.txtprocentajedetraccion.Text) / 100, "###0.00")
End Sub


Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      Me.TxtCliente.Text = rst("nombre_completo")
    Else
        Me.TxtCliente.Text = ""
        MsgBox "REGISTRE A SU PROVEEDOR", vbInformation
        Call Resalta(Me.txtRuc)
        Exit Sub
   End If
End If
End Sub

Private Sub TxtTc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call load_saldo_comprobante
End If
End Sub

Private Sub txtTipoFlujo_Change()
strCadena = "SELECT id as Codigo,Nombre  as Descripcion FROM adm_flujocaja WHERE Nombre LIKE '%" & Trim(Me.txtTipoFlujo.Text) & "%' ORDER BY Nombre  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFlujo)

End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form FrmListadoFacturasCompra 
   BorderStyle     =   0  'None
   Caption         =   "FrmListadoFacturas por Pagar"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   240
      TabIndex        =   118
      Top             =   480
      Width           =   615
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
      Height          =   3495
      Left            =   12720
      TabIndex        =   92
      Top             =   3600
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
         TabIndex        =   108
         Top             =   2520
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
         TabIndex        =   107
         Top             =   2520
         Width           =   975
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
         TabIndex        =   97
         Top             =   720
         Width           =   1815
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
         TabIndex        =   96
         Top             =   1260
         Width           =   1095
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
         TabIndex        =   95
         Top             =   1680
         Width           =   1095
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
         TabIndex        =   94
         Top             =   1725
         Width           =   1815
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
         TabIndex        =   93
         Top             =   1200
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DtcPeriodoAjuste 
         Height          =   330
         Left            =   1200
         TabIndex        =   98
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
      Begin VitekeySoft.ChameleonBtn cmdProcesarAjuste 
         Height          =   435
         Left            =   1200
         TabIndex        =   99
         Top             =   3000
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
         MICON           =   "FrmListadoFacturas.frx":0000
         PICN            =   "FrmListadoFacturas.frx":001C
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
         TabIndex        =   100
         Top             =   2115
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   5520
         Picture         =   "FrmListadoFacturas.frx":2601
         Top             =   240
         Width           =   240
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         Height          =   3495
         Left            =   0
         Top             =   0
         Width           =   5895
      End
      Begin VB.Label Label29 
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
         TabIndex        =   106
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label27 
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
         TabIndex        =   105
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label26 
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
         TabIndex        =   104
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label25 
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
         TabIndex        =   103
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   102
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label23 
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
         TabIndex        =   101
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame frmcanjeLetras 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7335
      Left            =   6840
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   11775
      Begin VB.TextBox txtid_canje 
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
         Left            =   0
         TabIndex        =   60
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chk_ajustecontable 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "AJUSTE CONTABLE"
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
         Height          =   240
         Left            =   3240
         TabIndex        =   59
         Top             =   6340
         Width           =   2055
      End
      Begin VB.Frame frmajustescontables 
         BackColor       =   &H00FFFFFF&
         Height          =   640
         Left            =   3240
         TabIndex        =   54
         Top             =   6480
         Visible         =   0   'False
         Width           =   4935
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
            Left            =   1470
            MaxLength       =   80
            TabIndex        =   56
            Top             =   240
            Width           =   1095
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
            Left            =   3600
            MaxLength       =   80
            TabIndex        =   55
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA REDONDEO:"
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
            TabIndex        =   58
            Top             =   285
            Width           =   1095
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
            Left            =   2880
            TabIndex        =   57
            Top             =   240
            Width           =   600
         End
      End
      Begin VB.TextBox txtBusqueda_proveedor 
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
         Left            =   8760
         TabIndex        =   53
         Top             =   5520
         Width           =   735
      End
      Begin VB.TextBox txtid_compranueva 
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
         Left            =   9000
         TabIndex        =   50
         Top             =   2085
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtnumero_factura 
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
         Left            =   3135
         TabIndex        =   45
         Top             =   2080
         Width           =   1335
      End
      Begin VB.TextBox txtserie_factura 
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
         Left            =   2400
         TabIndex        =   44
         Top             =   2080
         Width           =   735
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
         Left            =   1080
         TabIndex        =   38
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox txtserie_letra 
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
         Left            =   5280
         TabIndex        =   36
         Top             =   5940
         Width           =   615
      End
      Begin VB.TextBox TxtNumeroLetra 
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
         Left            =   6000
         TabIndex        =   29
         Top             =   5940
         Width           =   1095
      End
      Begin VB.TextBox txtMontoLetra 
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
         Left            =   7800
         TabIndex        =   26
         Top             =   5940
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DtpFechaLetra 
         Height          =   345
         Left            =   1080
         TabIndex        =   25
         Top             =   5460
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
         Format          =   170655745
         CurrentDate     =   43222
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfLetras 
         Height          =   2055
         Left            =   360
         TabIndex        =   21
         Top             =   2595
         Width           =   10695
         _ExtentX        =   18865
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
      Begin VitekeySoft.ChameleonBtn cmdDelete 
         Height          =   525
         Left            =   11160
         TabIndex        =   22
         Top             =   3240
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   926
         BTYPE           =   7
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
         MICON           =   "FrmListadoFacturas.frx":54A5
         PICN            =   "FrmListadoFacturas.frx":54C1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdupdate 
         Height          =   405
         Left            =   8760
         TabIndex        =   27
         Top             =   5925
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   714
         BTYPE           =   3
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
         MICON           =   "FrmListadoFacturas.frx":790B
         PICN            =   "FrmListadoFacturas.frx":7927
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdProcesar 
         Height          =   1005
         Left            =   9720
         TabIndex        =   31
         Top             =   4725
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1773
         BTYPE           =   5
         TX              =   "SAVE"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmListadoFacturas.frx":ABFD
         PICN            =   "FrmListadoFacturas.frx":AC19
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcMoneda 
         Height          =   315
         Left            =   3840
         TabIndex        =   32
         Top             =   4800
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
      Begin MSComCtl2.DTPicker DtpVence 
         Height          =   345
         Left            =   1080
         TabIndex        =   33
         Top             =   5880
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
         Format          =   170655745
         CurrentDate     =   43222
      End
      Begin MSDataListLib.DataCombo DtcPeriodo 
         Height          =   315
         Left            =   3840
         TabIndex        =   39
         Top             =   5160
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
      Begin VitekeySoft.ChameleonBtn cmdAnular 
         Height          =   1005
         Left            =   9720
         TabIndex        =   41
         Top             =   4800
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1773
         BTYPE           =   5
         TX              =   "ANULAR "
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
         MICON           =   "FrmListadoFacturas.frx":E261
         PICN            =   "FrmListadoFacturas.frx":E27D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfListadoFacturas 
         Height          =   1815
         Left            =   360
         TabIndex        =   42
         Top             =   120
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   3201
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
      Begin MSDataListLib.DataCombo Dtccomprobante 
         Height          =   315
         Left            =   360
         TabIndex        =   43
         Top             =   2085
         Width           =   2010
         _ExtentX        =   3545
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
      Begin MSDataListLib.DataCombo dtcProveedor 
         Height          =   315
         Left            =   5760
         TabIndex        =   46
         Top             =   2085
         Width           =   3255
         _ExtentX        =   5741
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
      Begin VitekeySoft.ChameleonBtn cmdagregar_factura 
         Height          =   405
         Left            =   10200
         TabIndex        =   48
         Top             =   2040
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   714
         BTYPE           =   3
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
         MICON           =   "FrmListadoFacturas.frx":E597
         PICN            =   "FrmListadoFacturas.frx":E5B3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdDEletefactura 
         Height          =   525
         Left            =   11160
         TabIndex        =   49
         Top             =   600
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   926
         BTYPE           =   8
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
         MICON           =   "FrmListadoFacturas.frx":11889
         PICN            =   "FrmListadoFacturas.frx":118A5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcProveedorLetra 
         Height          =   315
         Left            =   3840
         TabIndex        =   51
         Top             =   5520
         Width           =   4815
         _ExtentX        =   8493
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
      Begin MSDataListLib.DataCombo DtcTipoCanje 
         Height          =   315
         Left            =   2880
         TabIndex        =   87
         Top             =   5880
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
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROVEEDOR:"
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
         Left            =   2880
         TabIndex        =   52
         Top             =   5595
         Width           =   870
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROVEEDOR:"
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
         Left            =   4680
         TabIndex        =   47
         Top             =   2160
         Width           =   870
      End
      Begin VB.Label Label13 
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
         Left            =   2940
         TabIndex        =   40
         Top             =   5235
         Width           =   690
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TC :"
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
         TabIndex        =   37
         Top             =   5100
         Width           =   240
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONEDA:"
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
         Left            =   2970
         TabIndex        =   35
         Top             =   4875
         Width           =   660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENCE:"
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
         TabIndex        =   34
         Top             =   6000
         Width           =   480
      End
      Begin VB.Label Label9 
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
         Left            =   7200
         TabIndex        =   30
         Top             =   6000
         Width           =   600
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4560
         TabIndex        =   28
         Top             =   6000
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:"
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
         TabIndex        =   24
         Top             =   5565
         Width           =   480
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   2505
         Left            =   360
         Top             =   4680
         Width           =   10695
      End
      Begin VB.Label lbl_idcompra 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
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
         Left            =   360
         TabIndex        =   23
         Top             =   4440
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Image img_cerrar 
         Height          =   240
         Left            =   11475
         Picture         =   "FrmListadoFacturas.frx":13CEF
         Top             =   60
         Width           =   240
      End
      Begin VB.Label lblcomprobante 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMPROBANTE:"
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
         Left            =   360
         TabIndex        =   20
         Top             =   120
         Width           =   1125
      End
   End
   Begin VB.Frame frmEstadoCuenta 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   14520
      TabIndex        =   113
      Top             =   2520
      Visible         =   0   'False
      Width           =   4095
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
         TabIndex        =   116
         Top             =   240
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
         TabIndex        =   115
         Top             =   720
         Width           =   1935
      End
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
         TabIndex        =   114
         Top             =   1200
         Width           =   1935
      End
      Begin VitekeySoft.ChameleonBtn CmdEstadoCuentaProveedor 
         Height          =   1065
         Left            =   2280
         TabIndex        =   117
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1879
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
         MICON           =   "FrmListadoFacturas.frx":16B93
         PICN            =   "FrmListadoFacturas.frx":16BAF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00808080&
         Height          =   1575
         Left            =   0
         Top             =   0
         Width           =   4095
      End
      Begin VB.Image cmdcerrarEstado 
         Height          =   240
         Left            =   3780
         Picture         =   "FrmListadoFacturas.frx":19E85
         Top             =   50
         Width           =   240
      End
   End
   Begin VB.CheckBox chk_emision 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Left            =   9960
      TabIndex        =   112
      Top             =   480
      Width           =   1000
   End
   Begin VB.CheckBox chk_periodo 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Left            =   9960
      TabIndex        =   111
      Top             =   120
      Width           =   1000
   End
   Begin VB.Frame frm_canje_anticipos 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "ANTICIPOS"
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
      Height          =   3735
      Left            =   7320
      TabIndex        =   62
      Top             =   4800
      Visible         =   0   'False
      Width           =   11295
      Begin VB.TextBox txtId_canje_anticipo 
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
         Left            =   6000
         MaxLength       =   80
         TabIndex        =   81
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txttotal_anticipo_reversion 
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
         Left            =   6000
         MaxLength       =   80
         TabIndex        =   79
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtTotal_factura_reversion 
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
         Left            =   4200
         MaxLength       =   80
         TabIndex        =   78
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chk_ajuste_canje_anticipo 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "AJUSTE CONTABLE"
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
         Height          =   240
         Left            =   120
         TabIndex        =   75
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Frame frmajuste_anticipo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   120
         TabIndex        =   70
         Top             =   3015
         Visible         =   0   'False
         Width           =   5055
         Begin VB.TextBox txtmonto_anticipo 
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
            Left            =   3600
            MaxLength       =   80
            TabIndex        =   72
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtcuenta_redondeo_anticipo 
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
            Left            =   1470
            MaxLength       =   80
            TabIndex        =   71
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label20 
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
            Left            =   2880
            TabIndex        =   74
            Top             =   240
            Width           =   600
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA REDONDEO:"
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
            TabIndex        =   73
            Top             =   285
            Width           =   1095
         End
      End
      Begin VitekeySoft.ChameleonBtn cmdrevertir_anticipo 
         Height          =   885
         Left            =   8640
         TabIndex        =   63
         Top             =   2520
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1561
         BTYPE           =   3
         TX              =   "REALIZAR REVERSION"
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
         MICON           =   "FrmListadoFacturas.frx":1CD29
         PICN            =   "FrmListadoFacturas.frx":1CD45
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCompras 
         Height          =   1695
         Left            =   120
         TabIndex        =   64
         Top             =   720
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2990
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfAnticipos 
         Height          =   1695
         Left            =   6000
         TabIndex        =   65
         Top             =   720
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2990
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
      Begin VitekeySoft.ChameleonBtn cmd_agregar_factura 
         Height          =   555
         Left            =   3795
         TabIndex        =   66
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "AGREGAR"
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
         MICON           =   "FrmListadoFacturas.frx":1F32A
         PICN            =   "FrmListadoFacturas.frx":1F346
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmd_agregar_anticipo 
         Height          =   555
         Left            =   9480
         TabIndex        =   69
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "AGREGAR"
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
         MICON           =   "FrmListadoFacturas.frx":2192B
         PICN            =   "FrmListadoFacturas.frx":21947
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdEliminar_factura 
         Height          =   555
         Left            =   3120
         TabIndex        =   76
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   979
         BTYPE           =   3
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
         MICON           =   "FrmListadoFacturas.frx":23F2C
         PICN            =   "FrmListadoFacturas.frx":23F48
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmd_eliminar_anticipo 
         Height          =   555
         Left            =   8760
         TabIndex        =   77
         Top             =   120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   979
         BTYPE           =   3
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
         MICON           =   "FrmListadoFacturas.frx":26392
         PICN            =   "FrmListadoFacturas.frx":263AE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdAnularReversion 
         Height          =   885
         Left            =   8640
         TabIndex        =   80
         Top             =   2640
         Visible         =   0   'False
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1561
         BTYPE           =   3
         TX              =   "ANULAR REVERSION"
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
         MICON           =   "FrmListadoFacturas.frx":287F8
         PICN            =   "FrmListadoFacturas.frx":28814
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   11040
         Picture         =   "FrmListadoFacturas.frx":28B2E
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LISTADO DE COMPROBANTES"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   68
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LISTADO DE ANTICIPOS"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   6000
         TabIndex        =   67
         Top             =   360
         Width           =   2130
      End
   End
   Begin VB.CheckBox chk_tipo_doc 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "TIPO DOC:"
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
      Left            =   6720
      TabIndex        =   86
      Top             =   120
      Width           =   1000
   End
   Begin VB.TextBox txtTipo_cambio 
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
      Left            =   16080
      TabIndex        =   85
      Top             =   360
      Width           =   735
   End
   Begin VB.CheckBox chk_porvencerse 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "POR VENCER"
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
      Left            =   14520
      TabIndex        =   17
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox chkVencidas 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "VENCIDAS"
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
      Left            =   14520
      TabIndex        =   16
      Top             =   405
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   300
      Left            =   11160
      TabIndex        =   9
      Top             =   480
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
      Format          =   169803777
      CurrentDate     =   42817
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "ALMACEN :"
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
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtnumero 
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
      Height          =   285
      Left            =   7800
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox TxtcodProveedor 
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
      Height          =   285
      Left            =   4725
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox tXTrUC 
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
      Left            =   4725
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   330
      Left            =   1320
      TabIndex        =   2
      Top             =   105
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgFacturas 
      Height          =   7935
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   18375
      _ExtentX        =   32411
      _ExtentY        =   13996
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
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   300
      Left            =   12600
      TabIndex        =   10
      Top             =   480
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
      Format          =   169803777
      CurrentDate     =   42817
   End
   Begin VitekeySoft.ChameleonBtn cmdbuscar 
      Height          =   420
      Left            =   16920
      TabIndex        =   11
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   741
      BTYPE           =   5
      TX              =   "CUENTAS POR PAGAR"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmListadoFacturas.frx":2B9D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdamortizar 
      Height          =   855
      Left            =   18840
      TabIndex        =   12
      Top             =   840
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "AMORTIZAR"
      ENAB            =   0   'False
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
      MICON           =   "FrmListadoFacturas.frx":2B9EE
      PICN            =   "FrmListadoFacturas.frx":2BA0A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdhistorial 
      Height          =   855
      Left            =   18840
      TabIndex        =   13
      Top             =   4440
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "HISTORIAL"
      ENAB            =   0   'False
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
      MICON           =   "FrmListadoFacturas.frx":2E2F4
      PICN            =   "FrmListadoFacturas.frx":2E310
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
      TabIndex        =   14
      Top             =   8040
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "FrmListadoFacturas.frx":31919
      PICN            =   "FrmListadoFacturas.frx":31935
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCanjeLetra 
      Height          =   855
      Left            =   18840
      TabIndex        =   18
      Top             =   6280
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "CANJE LETRA"
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
      MICON           =   "FrmListadoFacturas.frx":31D25
      PICN            =   "FrmListadoFacturas.frx":31D41
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdVincularAnticipo 
      Height          =   855
      Left            =   18840
      TabIndex        =   61
      Top             =   7150
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "REVERTIR ANTICIPO"
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
      MICON           =   "FrmListadoFacturas.frx":3205B
      PICN            =   "FrmListadoFacturas.frx":32077
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcTipoComprobante 
      Height          =   330
      Left            =   7800
      TabIndex        =   82
      Top             =   75
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
   Begin MSDataListLib.DataCombo DtcBusquedaPeriodo 
      Height          =   330
      Left            =   11160
      TabIndex        =   83
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VitekeySoft.ChameleonBtn cmdDetallado 
      Height          =   855
      Left            =   18840
      TabIndex        =   88
      Top             =   2600
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "DETALLADO"
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
      MICON           =   "FrmListadoFacturas.frx":342D0
      PICN            =   "FrmListadoFacturas.frx":342EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdAjuste 
      Height          =   915
      Left            =   18840
      TabIndex        =   89
      Top             =   5340
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1614
      BTYPE           =   5
      TX              =   "AJUSTE TC"
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
      MICON           =   "FrmListadoFacturas.frx":34606
      PICN            =   "FrmListadoFacturas.frx":34622
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdRegenerar 
      Height          =   350
      Left            =   1320
      TabIndex        =   91
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "RE-GENERAR ASIENTO"
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
      MICON           =   "FrmListadoFacturas.frx":34A98
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdEstadoCuenta 
      Height          =   855
      Left            =   18840
      TabIndex        =   109
      Top             =   3480
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "ESTADO CUENTA"
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
      MICON           =   "FrmListadoFacturas.frx":34AB4
      PICN            =   "FrmListadoFacturas.frx":34AD0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdGeneral 
      Height          =   855
      Left            =   18840
      TabIndex        =   110
      Top             =   1720
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "CUENTAS PAGAR"
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
      MICON           =   "FrmListadoFacturas.frx":34DEA
      PICN            =   "FrmListadoFacturas.frx":34E06
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CUENTAS X PAGAR"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18840
      TabIndex        =   90
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T.CAMBIO"
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
      Left            =   16080
      TabIndex        =   84
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUENTAS POR PAGAR."
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
      Left            =   240
      TabIndex        =   15
      Top             =   8880
      Width           =   1800
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   7080
      TabIndex        =   7
      Top             =   510
      Width           =   690
   End
   Begin VB.Label Label2 
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
      Left            =   3600
      TabIndex        =   5
      Top             =   150
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC:"
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
      Left            =   4320
      TabIndex        =   4
      Top             =   510
      Width           =   330
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmListadoFacturasCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ProcedenciaFacturas As EnumFactura
Public Procedencia As EnumProcede
Dim serie As String
Dim numero As String
Dim Persona As String
Public Sub facturas()
Dim in_inicio_mes As String

in_inicio_mes = Trim("01-" & Format(Month(KEY_FECHA), "0") & "-" & Year(KEY_FECHA))

strCadena = "SELECT `id_compra`,id_doc,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_facturaii(id_compra,'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,reversion  FROM view_cuentas_cobrar WHERE fecha_emision>='" & Format(in_inicio_mes, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
Call llenarGrid(Me.HfgFacturas)

End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
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
           Grilla.ColWidth(3) = 3000
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 3800
           Grilla.ColWidth(6) = 1100
           Grilla.ColWidth(7) = 800
           Grilla.ColWidth(8) = 1200
           Grilla.ColWidth(9) = 1200
           Grilla.ColWidth(10) = 1200
           Grilla.ColWidth(11) = 2400
        Next
        cabecera = "IDCOMPRA" & vbTab & "EMISION" & vbTab & "VENCIMIENTO" & vbTab & "COMPROBANTE" & vbTab & "RUC/DNI" & vbTab & "PROVEEDOR" & vbTab & "MONEDA" & vbTab & "  TC " & vbTab & "FACTURADO" & vbTab & "SALDO [DOLAR]" & vbTab & "SALDO [SOLES]" & vbTab & "ENCARGADO" & vbTab & "IDPROVEEDOR"
        Grilla.AddItem cabecera
        For k = 0 To 11
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
       Next k
                            
        rst.MoveFirst
        
        in_cambio = Val(Me.txtTipo_cambio.Text)
        in_pago = 0
        For i = 0 To rst.RecordCount - 1
            If rst("id_doc") = "0007" Then
                in_pago = rst("pago") * -1
            Else
                in_pago = rst("pago")
            End If
            
            
            
            
            If rst("id_moneda") = "00002" Then
               If rst("reversion") = "no" Then
                    If rst("id_doc") = "0419" Then
                        nsaldo_soles = rst("total") * in_cambio
                        nsaldo_dolar = rst("total")
                    Else
                        nsaldo_soles = (rst("total") - in_pago) * in_cambio
                        nsaldo_dolar = rst("total") - in_pago
                    End If
               Else
                    nsaldo_soles = (rst("total") - in_pago) * in_cambio
                    nsaldo_dolar = rst("total") - in_pago
                    
               End If
            Else
               If rst("reversion") = "no" Then
                  If rst("id_doc") = "0419" Then
                        nsaldo_soles = (rst("total") - in_pago)
                        nsaldo_dolar = (rst("total") - in_pago) / in_cambio
                  Else
                       
                        nsaldo_soles = (rst("total") - in_pago)
                        nsaldo_dolar = (rst("total") - in_pago) / in_cambio
                  End If
            Else
                nsaldo_soles = (rst("total") - in_pago)
                nsaldo_dolar = (rst("total") - in_pago) / in_cambio
               End If
            End If
            
            
            tSaldo_soles = tSaldo_soles + nsaldo_soles
            tSaldo_dolares = tSaldo_dolares + nsaldo_dolar
            
            
            
            Fila = rst("id_compra") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & Format(rst("fecha_cancelacion"), "dd-mm-YYYY") & vbTab & rst("comprobante") & Space(2) & "[" & rst("id_compra") & "]" & vbTab & rst("id_proveedor") & vbTab & Mid(UCase(rst("nproveedor")), 1, 40) & vbTab & rst("moneda") & vbTab & Format(rst("tc"), "#,##0.0000") & vbTab & Format(rst("total"), "#,##0.00") & vbTab & Format(nsaldo_dolar, "#,##0.00") & vbTab & Format(nsaldo_soles, "#,##0.00") & vbTab & Mid(rst("nombre_completo"), 1, 25) & vbTab & rst("id_proveedor")
            Grilla.AddItem Fila
            If nsaldo_soles <> 0 Then
            For k = 8 To 10
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H80C0FF
            Next k
            End If
            rst.MoveNext
        Next i
        cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(tSaldo_dolares, "#,##0.00") & vbTab & Format(tSaldo_soles, "#,##0.00") & vbTab & ""
        Grilla.AddItem cabecera
                            For k = 8 To 10
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
                            Next k
    
    
    
 ' Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub
Public Sub llenarGrid_Proveedor()
On Error GoTo salir
Dim tTotal As Double
tTotal = 0

strCadena = "SELECT  DocumentoCompra.idCompra, DocumentoCompra.Alm_cod, DocumentoCompra.dEmisionCompra, DocumentoCompra.dVencimiento," & _
" (Comprobantes.doc_abrev + ':'+ DocumentoCompra.sSerie +'-'+ DocumentoCompra.cDocumentoCompra) as DOCUMENTO,DocumentoCompra.Persona, " & _
"DocumentoCompra.moneda , DocumentoCompra.tc, DocumentoCompra.nTotalCompra, DocumentoCompra.Saldo,Anulado,DocumentoCompra.cPersona,seleccion " & _
"FROM DocumentoCompra INNER JOIN Comprobantes ON DocumentoCompra.doc_cod = Comprobantes.doc_cod WHERE (DocumentoCompra.doc_cod='0001' OR DocumentoCompra.doc_cod='0003') AND saldo>0 AND DocumentoCompra.Persona LIKE '%" & Trim(Me.TxtcodProveedor.Text) & "%' ORDER BY DocumentoCompra.dVencimiento ASC "

Call ConfiguraRst(strCadena)
 
If rst.RecordCount < 1 Then
    Me.HfgFacturas.Rows = 1
    HfgFacturas.Clear
    Exit Sub

End If
  
  N = 1
  
   If Me.HfgFacturas.Rows > 0 Then
   HfgFacturas.Clear
   HfgFacturas.Refresh
   HfgFacturas.Rows = 0
   End If
   
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           HfgFacturas.ColWidth(0) = 600
           HfgFacturas.ColWidth(1) = 1000
           HfgFacturas.ColWidth(2) = 1000
           HfgFacturas.ColWidth(3) = 2300
           HfgFacturas.ColWidth(4) = 4000
           HfgFacturas.ColWidth(5) = 500
           HfgFacturas.ColWidth(6) = 500
           HfgFacturas.ColWidth(7) = 1500
           HfgFacturas.ColWidth(8) = 1500
           HfgFacturas.ColWidth(9) = 0
           HfgFacturas.ColWidth(10) = 0
           
           
        Next
         cabecera = "ITEM" & vbTab & "EMISION" & vbTab & "VENCIMIENTO" & vbTab & "COMPROBANTE" & vbTab & "PROVEEDOR" & vbTab & "M" & vbTab & "TC" & vbTab & "TOTAL" & vbTab & "SALDO" & vbTab & "codigounico" & vbTab & "cPersona"
        HfgFacturas.AddItem cabecera
         For k = 0 To 9
                                HfgFacturas.col = k
                                HfgFacturas.Row = 0
                                HfgFacturas.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If (rst("moneda") = "0001") Then
               ' moneda = "S/."
            Else
               ' moneda = "US$."
            End If
            'fila = fila & Str(i + 1) & vbTab & rst("dEmisionCompra") & vbTab & rst("dVencimiento") & vbTab & rst(4) & vbTab & rst("Persona") & vbTab & moneda & vbTab & rst("tc") & vbTab & Format(rst("nTotalCompra"), "#,##0.00") & vbTab & Format(rst("Saldo"), "#,##0.00") & vbTab & rst("idCompra") & vbTab & rst("cPersona")
            If (Fila = "") Then
                X = 1
            End If
          HfgFacturas.AddItem Fila
               
                    
                    
                       If (Trim(rst("Anulado")) = "V") Then
                            For k = 0 To 9
                                HfgFacturas.col = k
                                HfgFacturas.Row = i
                                HfgFacturas.CellBackColor = &H8080FF
                            Next k
                        Else
                        
                        tTotal = tTotal + rst("saldo")
                        End If
                        
                        If (Trim(rst("seleccion")) = "si") Then
                            For k = 0 To 9
                                HfgFacturas.col = k
                                HfgFacturas.Row = i + 1
                                HfgFacturas.CellBackColor = &H80FF80
                            Next k
                       End If
                   
            
            
            Fila = ""
            rst.MoveNext
             
        Next i
    
    
    
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub
Public Sub llenarGrid_ProveedorRUC()
On Error GoTo salir
Dim tTotal As Double
tTotal = 0

strCadena = "SELECT     DocumentoCompra.idCompra, DocumentoCompra.Alm_cod, DocumentoCompra.dEmisionCompra, DocumentoCompra.dVencimiento, " & _
"                      Comprobantes.doc_abrev, DocumentoCompra.sSerie, DocumentoCompra.cDocumentoCompra, DocumentoCompra.Persona, " & _
"                      DocumentoCompra.moneda, DocumentoCompra.tc, DocumentoCompra.nTotalCompra, DocumentoCompra.saldo, " & _
"                      DocumentoCompra.Anulado ,   DocumentoCompra.cPersona,seleccion FROM         DocumentoCompra INNER JOIN " & _
"                      Comprobantes ON DocumentoCompra.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
"                      Persona ON DocumentoCompra.cPersona = Persona.cPersona WHERE (DocumentoCompra.doc_cod='0001' OR DocumentoCompra.doc_cod='0003') AND saldo>0 AND Persona.Per_Ruc LIKE '%" & Trim(Me.txtRuc.Text) & "%' ORDER BY DocumentoCompra.dVencimiento ASC "


Call ConfiguraRst(strCadena)
 
If rst.RecordCount < 1 Then
    Me.HfgFacturas.Rows = 1
    HfgFacturas.Clear
    Exit Sub

End If
  
  N = 1
  
   HfgFacturas.Clear
   HfgFacturas.Refresh
   HfgFacturas.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           HfgFacturas.ColWidth(0) = 600
           HfgFacturas.ColWidth(1) = 1000
           HfgFacturas.ColWidth(2) = 1000
           HfgFacturas.ColWidth(3) = 2300
           HfgFacturas.ColWidth(4) = 4000
           HfgFacturas.ColWidth(5) = 500
           HfgFacturas.ColWidth(6) = 500
           HfgFacturas.ColWidth(7) = 1500
           HfgFacturas.ColWidth(8) = 1500
           HfgFacturas.ColWidth(9) = 0
           
           
           
        Next
         cabecera = "ITEM" & vbTab & "EMISION" & vbTab & "VENCIMIENTO" & vbTab & "COMPROBANTE" & vbTab & "PROVEEDOR" & vbTab & "M" & vbTab & "TC" & vbTab & "TOTAL" & vbTab & "SALDO" & vbTab & "codigounico" & vbTab & "cpersona"
        HfgFacturas.AddItem cabecera
         For k = 0 To 9
                                HfgFacturas.col = k
                                HfgFacturas.Row = 0
                                HfgFacturas.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If (rst("moneda") = "0001") Then
               ' moneda = "S/."
            Else
               ' moneda = "US$."
            End If
            If IsNull(rst("dVencimiento")) = True Then
                vencimiento = rst("dEmisionCompra")
            End If
            'fila = fila & Str(i + 1) & vbTab & rst("dEmisionCompra") & vbTab & vencimiento & vbTab & rst(4) & vbTab & rst("Persona") & vbTab & moneda & vbTab & rst("tc") & vbTab & Format(rst("nTotalCompra"), "#,##0.00") & vbTab & Format(rst("Saldo"), "#,##0.00") & vbTab & rst("idCompra") & vbTab & rst("cPersona")
            If (Fila = "") Then
                X = 1
            End If
          HfgFacturas.AddItem Fila
               
                    
                    
                       If (Trim(rst("Anulado")) = "V") Then
                            For k = 0 To 9
                                HfgFacturas.col = k
                                HfgFacturas.Row = i
                                HfgFacturas.CellBackColor = &H8080FF
                            Next k
                        Else
                        
                        tTotal = tTotal + rst("saldo")
                        End If
                        
                         If (Trim(rst("seleccion")) = "si") Then
                            For k = 0 To 9
                                HfgFacturas.col = k
                                HfgFacturas.Row = i - 1
                                HfgFacturas.CellBackColor = &H80FF80
                            Next k
                       End If
                        
                   
            
            
            Fila = ""
            rst.MoveNext
             
        Next i
    
    
    
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub

Sub llenarGrid_PagosLetras(ByVal Grilla As MSHFlexGrid, ByVal NumeroVenta As String, ByVal serie As String, ByVal Persona As String)

strCadena = "SELECT id_detalle as Codigo, FechaPago, Monto, Operacion FROM  DetallePagos WHERE (DetallePagos.numero='" & NumeroVenta & "'  AND DetallePagos.serie='" & serie & "' AND DetallePagos.cPersona='" & Persona & "')"
On Error GoTo salir
  Call ConfiguraRst(strCadena)

  Grilla.Clear
  Set Grilla.Recordset = rst
  Grilla.Rows = rst.RecordCount
  Grilla.ColWidth(0) = 1500
  Grilla.ColWidth(1) = 2000
  Grilla.ColWidth(2) = 1300
  Grilla.ColWidth(3) = 1300
  Grilla.ColWidth(4) = 0
Call DarFormatoFecha(Grilla, 1)
Call DarFormato(Grilla, 2)


Set rst = Nothing

  'Me.TlbAcciones.Buttons(KEY_PRINT).Enabled = True
  'Me.TlbAcciones.Buttons(KEY_EXIT).Enabled = True
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub



Private Sub ChameleonBtn1_Click()

End Sub

Private Sub chk_ajuste_canje_anticipo_Click()
If Me.chk_ajuste_canje_anticipo.Value = 1 Then
   Me.frmajuste_anticipo.Visible = True
Else
    Me.frmajuste_anticipo.Visible = False
End If
End Sub

Private Sub chk_ajustecontable_Click()

If Me.chk_ajustecontable.Value = 1 Then
  Me.frmajustescontables.Visible = True
Else
  Me.frmajustescontables.Visible = True
End If



End Sub

Private Sub chk_cta_contable_Click()
If Me.chk_cta_contable.Value = 1 Then
    Me.txtAjusteporCuenta.Visible = True
Else
    Me.txtAjusteporCuenta.Visible = False
End If
End Sub

Private Sub chk_porvencerse_Click()
If Me.chk_porvencerse.Value = 1 Then
   Me.chkVencidas.Value = 0
End If

End Sub

Private Sub chkVencidas_Click()
If Me.chkVencidas.Value = 1 Then
   Me.chk_porvencerse.Value = 0
End If
End Sub

Private Sub cmd_Click()

End Sub

Private Sub cmd_agregar_anticipo_Click()
If Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) > 0 Then
    strCadena = "SELECT `id_compra`,`id_moneda`,`tc`,`total`, function_pago_factura(id_compra,'" & KEY_FECHA & "',id_moneda,ruc) as pago,reversion,id_doc  FROM view_cuentas_cobrar WHERE id_doc='0419' and  id_compra='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If rst("reversion") = "no" Then
            in_saldo = rst("total") - rst("pago")
        Else
            in_saldo = rst("total") - rst("pago")
        End If
        strCadena = "call put_vincular_anticipo('" & rst("id_compra") & "','" & rst("id_doc") & "','" & in_saldo & "','si','" & KEY_USUARIO & "','" & KEY_RUC & "') "
        CnBd.Execute (strCadena)
        strCadena = "SELECT * FROM view_anticipo_temporal WHERE anticipo='si' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
        Call Me.llenar_anticipo(Me.HfAnticipos)
    End If
End If
End Sub

Private Sub cmd_agregar_factura_Click()
If Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) > 0 Then
    strCadena = "SELECT `id_compra`,`id_moneda`,`tc`,`total`, function_pago_factura(id_compra,'" & KEY_FECHA & "',id_moneda,ruc) as pago,reversion,id_doc  FROM view_cuentas_cobrar WHERE id_compra='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        in_saldo = rst("total") - rst("pago")
        strCadena = "call put_vincular_anticipo('" & rst("id_compra") & "','" & rst("id_doc") & "','" & in_saldo & "','no','" & KEY_USUARIO & "','" & KEY_RUC & "') "
        CnBd.Execute (strCadena)
        strCadena = "SELECT * FROM view_anticipo_temporal WHERE anticipo='no' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
        Call Me.llenar_anticipo(Me.HfCompras)
    End If
End If
End Sub

Private Sub cmd_eliminar_anticipo_Click()
If Val(Me.HfAnticipos.Rows) > 0 Then
strCadena = "call put_limpiar_anticipo('" & Val(HfAnticipos.TextMatrix(Me.HfAnticipos.Row, 0)) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
strCadena = "SELECT * FROM view_anticipo_temporal WHERE anticipo='si' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"

Call Me.llenar_anticipo(Me.HfAnticipos)
End If
End Sub

Private Sub cmdagregar_factura_Click()
Call seleccionar_factura(Me.txtid_compranueva.Text)
strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,monto_pagar,reversion FROM view_cuentas_cobrar WHERE dni_save_pago='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
Call Me.llenar_facturas(Me.HfListadoFacturas)
End Sub

Private Sub cmdAjuste_Click()
 strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodoAjuste)
  

Me.frmajusteCobrar.Visible = True




Exit Sub



If MsgBox("Esta seguro de Realizar el ajuste [TC] para:" + Chr(13) + Chr(13) + "TIPO CAMBIO:" & str(Me.txtTipo_cambio.Text) + Chr(13) + "PERIODO AJUSTE:" + get_periodo_descripcion(get_periodo_actual(Me.DtpInicio.Value)), vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then
        
    strCadena = "SELECT DISTINCT id_proveedor FROM view_cuentas_cobrar WHERE id_moneda='00002' and  (total-function_pago_factura(id_compra,'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "',id_moneda,ruc)) >0   and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       For i = 0 To rst.RecordCount - 1
           
           
           strCadena = "call CON_InsertaAsiento_AjusteTC_Empresa('" & KEY_RUC & "','" & get_periodo_actual(Me.DtpInicio.Value) & "','4212','" & Trim(rst("id_proveedor")) & "','" & Val(Me.txtTipo_cambio.Text) & "','" & KEY_USUARIO & "')"
           CnBd.Execute (strCadena)
           
           rst.MoveNext
       Next i
    End If
        MsgBox "AJUSTE DE TIPO CAMBIO REALIZADO", vbInformation, KEY_VENDEDOR
End If


End Sub

Private Sub cmdamortizar_Click()

strCadena = "SELECT  (total-function_pago_factura(id_compra,'" & KEY_FECHA & "',id_moneda,ruc)) as saldon,id_moneda,tc,reversion FROM view_cuentas_cobrar WHERE id_compra='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   strCadena = "UPDATE movimiento_compra SET seleccion='no',dni_save_pago='0' WHERE id_compra<>'" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) & "'  and dni_save_pago='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
    If rst("tc") <= 0 Then
        in_tc = KEY_CAMBIO
    Else
        in_tc = Val(Me.txtTipo_cambio.Text)
    End If
    
   If KEY_PAIS = KEY_PERU Then
   If rst("id_moneda") = "00002" Then
        saldon = rst("saldon") * in_tc
   Else
        saldon = rst("saldon")
   End If
   Else
    saldon = rst("saldon")
   End If
   strCadena = "UPDATE movimiento_compra SET seleccion='si',dni_save_pago='" & KEY_USUARIO & "',monto_pagar='" & saldon & "' WHERE id_compra='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
   
   strCadena = "DELETE FROM persona_gasto WHERE id_venta='0' and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
   
   

 
 
End If



      
      FrmComprasPagos.Show
      Exit Sub
End Sub

Private Sub cmdAnular_Click()

Call anular_letras(Me.txtid_canje.Text, Me.DtcTipoCanje.BoundText)

End Sub

Private Sub anular_letras(ByVal in_canje As String, ByVal in_doc As String)

If MsgBox("Esta seguro de Anular EL CANJE", vbInformation + vbYesNo, KEY_VENDEDOR) = vbYes Then
    
    
    strCadena = "call CON_InsertaAsiento_CanjeLetra_Extorno('" & Val(in_canje) & "')"
    CnBd.Execute (strCadena)
    
    strCadena = "SELECT * FROM movimiento_compra_canje_letra WHERE  id_canje='" & Val(in_canje) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       For i = 0 To rst.RecordCount - 1
           strCadena = "DELETE FROM movimiento_compra_canje_letra WHERE id_detalle='" & rst("id_detalle") & "' and ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
           
           
           
           
           rst.MoveNext
       Next i
    End If
    
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & in_doc & "' and  id_canje='" & Val(in_canje) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       For i = 0 To rst.RecordCount - 1
           strCadena = "DELETE FROM mis_cuentas_det_detalle WHERE id_detalle='" & rst("id_venta") & "'"
           CnBd.Execute (strCadena)
           
            strCadena = "DELETE  FROM movimiento_venta WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
    
           rst.MoveNext
       Next i
    End If
    
    
    
    strCadena = "DELETE  FROM movimiento_venta WHERE id_doc='" & in_doc & "' and  id_canje='" & Val(in_canje) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    strCadena = "DELETE  FROM movimiento_compra WHERE id_doc='" & in_doc & "' and id_canje='" & Val(in_canje) & "' and  ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    
    strCadena = "SELECT * FROM  movimiento_compra WHERE id_canje='" & Val(in_canje) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       For i = 0 To rst.RecordCount - 1
           
        strCadena = "UPDATE movimiento_compra SET id_canje='0',reversion='no' WHERE id_compra='" & rst("id_compra") & "'"
           CnBd.Execute (strCadena)
           rst.MoveNext
       Next i
    End If
    
    MsgBox "Proceso Exitoso.", vbInformation
    
    
End If


End Sub


Private Sub cmdAnularReversion_Click()
If Val(Me.txtId_canje_anticipo.Text) > 0 Then
    If MsgBox("Esta Seguro de Eliminar la Reversion", vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then
    strCadena = "call CON_InsertaAsiento_CanjeAnticipo_Extorno('" & Val(Me.txtId_canje_anticipo.Text) & "')"
    CnBd.Execute (strCadena)
    
    
    strCadena = "SELECT * FROM movimiento_compra_anticipo WHERE id_canje_anticipo='" & Val(Me.txtId_canje_anticipo.Text) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       For i = 0 To rst.RecordCount - 1
                
                strCadena = "UPDATE movimiento_compra SET reversion='no' WHERE id_compra='" & rst("id_compra") & "'"
                CnBd.Execute (strCadena)
                
                strCadena = "DELETE FROM mis_cuentas_det_detalle WHERE id_movimiento='" & rst("id_compra") & "' and monto_pagado='" & rst("monto_compra") & "'"
                CnBd.Execute (strCadena)
                
                strCadena = "DELETE FROM movimiento_compra_anticipo WHERE id_canje_anticipo='" & Val(Me.txtId_canje_anticipo.Text) & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                
                
                rst.MoveNext
       Next i
    End If
End If
MsgBox "Reversion realizada", vbInformation
End If
End Sub

Private Sub cmdBuscar_Click()

Dim in_operacion As String


If Me.chk_emision.Value = 1 Then
    in_operacion = 2
Else
    in_operacion = 4
End If

If Me.chk_periodo.Value = 1 Then
     Me.txtTipo_cambio.Text = get_tipo_cambio_dia(CVDate(get_ultimo_dia_periodo(Me.DtcPeriodo.BoundText)), "valor_venta")
    in_operacion = 5
End If




strCadena = "CALL CON_CuentaPagarV2_LST('" & in_operacion & "','" & Trim(Me.txtRuc.Text) & "','" & Trim(TxtcodProveedor.Text) & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & Me.DtcComprobante.BoundText & "','" & Trim(Me.txtNumero.Text) & "','','" & Val(Me.txtTipo_cambio.Text) & "','" & Me.DtcBusquedaPeriodo.BoundText & "','" & KEY_RUC & "')"
Call llenarGrid(Me.HfgFacturas)

    
    
End Sub

Private Sub cmdCerrar_Click()

End Sub
Private Sub limpiar_seleccion()

strCadena = "DELETE FROM movimiento_compra_letra WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


strCadena = "SELECT * FROM movimiento_compra WHERE dni_save_pago='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "UPDATE movimiento_compra SET dni_save_pago='0',seleccion='no' WHERE id_compra='" & rst("id_compra") & "'"
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If

End Sub

Private Sub seleccionar_factura(ByVal in_compra As String)
strCadena = "SELECT  (total-function_pago_factura(id_compra,'" & KEY_FECHA & "',id_moneda,ruc)) as saldon,id_moneda,tc,reversion FROM view_cuentas_cobrar WHERE id_compra='" & Val(in_compra) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
  If rst("tc") <= 0 Then
        in_tc = KEY_CAMBIO
    Else
        in_tc = rst("tc")
    End If
    
   If rst("id_moneda") = "00002" Then
        saldon = rst("saldon") * in_tc
   Else
        saldon = rst("saldon")
   End If
   strCadena = "UPDATE movimiento_compra SET seleccion='si',dni_save_pago='" & KEY_USUARIO & "',monto_pagar='" & saldon & "' WHERE id_compra='" & Val(in_compra) & "' AND ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
   Me.txtid_compranueva.Text = 0
End If
End Sub

Private Function verificar_canje(ByVal in_compra As String) As Boolean
Dim in_canje As String
strCadena = "SELECT id_canje FROM movimiento_compra WHERE id_compra='" & Val(in_compra) & "' "
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
If rstLocal("id_canje") > 0 Then
   verificar_canje = True
   Me.txtid_canje.Text = rstLocal("id_canje")
   strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,monto_pagar,reversion FROM view_cuentas_cobrar WHERE id_doc<>'0417' and  id_canje='" & Val(txtid_canje.Text) & "' and  ruc='" & KEY_RUC & "'"
   Call llenar_facturas(Me.HfListadoFacturas)
   
   strCadena = "SELECT * FROM movimiento_compra_canje_letra WHERE id_canje='" & Val(Me.txtid_canje.Text) & "' and ruc='" & KEY_RUC & "' ORDER BY id_detalle DESC LIMIT 1"
   Call ConfiguraRstlocal(strCadena)
   If rstLocal.RecordCount > 0 Then
      Me.DtcTipoCanje.BoundText = rstLocal("id_doc")
   End If
   
   
   strCadena = "SELECT id_detalle as id_letra,emision as fecha,vencimiento as vencimiento,monto as monto,serie,numero,id_doc FROM movimiento_compra_canje_letra WHERE id_doc='" & Me.DtcTipoCanje.BoundText & "' and  id_canje='" & Val(Me.txtid_canje.Text) & "' and ruc='" & KEY_RUC & "'"
   Call llenar_letras(Me.HfLetras)
   
   Me.cmdAnular.Visible = True
   Me.cmdProcesar.Visible = False
   Me.frmcanjeLetras.Visible = True
   Me.cmdupdate.Visible = False
Else
   verificar_canje = False
   Me.cmdAnular.Visible = False
   Me.cmdProcesar.Visible = True
   Me.cmdupdate.Visible = True
   Me.txtid_canje.Text = 0
   strCadena = "SELECT * FROM  movimiento_compra_letra WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' ORDER BY fecha ASC"
   Call llenar_letras(Me.HfLetras)
End If
End If
End Function

Private Sub cmdCanjeLetra_Click()
Me.txtTc.Text = get_tipo_cambio_dia(Me.DtpFechaLetra.Value, "valor_venta")
strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM comprobantes ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComprobante)
Me.DtcComprobante.BoundText = "0001"

If Me.HfgFacturas.Rows > 0 Then

Call limpiar_seleccion
Me.lbl_idcompra.Caption = Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0))
If verificar_canje(Me.lbl_idcompra.Caption) = True Then
   Exit Sub
End If

Call seleccionar_factura(Val(Me.lbl_idcompra.Caption))

Me.frmcanjeLetras.Visible = True
strCadena = "SELECT id_compra as id_letra,fecha_emision as fecha,fecha_cancelacion as vencimiento,total as monto,serie,numero FROM movimiento_compra WHERE id_doc='0417' and  id_comprobante_referencia='" & Val(Me.lbl_idcompra.Caption) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
    strCadena = "SELECT id_compra as id_letra,fecha_emision as fecha,fecha_cancelacion as vencimiento,total as monto,serie,numero,id_doc FROM movimiento_compra WHERE id_doc='0417' and  id_comprobante_referencia='" & Val(Me.lbl_idcompra.Caption) & "' and ruc='" & KEY_RUC & "'"
    Call llenar_letras(Me.HfLetras)
   Me.cmdAnular.Visible = True
   Me.cmdDelete.Enabled = False
Else

   Me.lblComprobante.Caption = Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 1)
    Me.cmdAnular.Visible = False
    Me.cmdDelete.Enabled = True
End If

strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,monto_pagar,reversion FROM view_cuentas_cobrar WHERE dni_save_pago='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
Call llenar_facturas(Me.HfListadoFacturas)

End If



End Sub

Private Sub cmdcerrarEstado_Click()
Me.frmEstadoCuenta.Visible = False
End Sub

Private Sub cmddelete_Click()
If Val(Me.HfLetras.TextMatrix(Me.HfLetras.Row, 0)) > 0 Then
    strCadena = "DELETE FROM movimiento_compra_letra WHERE id_letra='" & Val(Me.HfLetras.TextMatrix(Me.HfLetras.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    strCadena = "SELECT * FROM  movimiento_compra_letra WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' ORDER BY fecha ASC"
    Call Me.llenar_letras(Me.HfLetras)
End If
End Sub

Private Sub cmdDEletefactura_Click()
If Me.HfListadoFacturas.Rows > 0 Then
    If Val(Me.HfListadoFacturas.TextMatrix(Me.HfListadoFacturas.Row, 0)) > 0 Then
        strCadena = "UPDATE movimiento_compra SET seleccion='no',dni_save_pago='0' WHERE id_compra='" & Val(Me.HfListadoFacturas.TextMatrix(Me.HfListadoFacturas.Row, 0)) & "'"
        CnBd.Execute (strCadena)
        strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,monto_pagar,reversion FROM view_cuentas_cobrar WHERE dni_save_pago='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
        Call Me.llenar_facturas(Me.HfListadoFacturas)
    End If
End If

End Sub



Private Sub cmdDetallado_Click()
If Trim(Me.txtRuc.Text) = "" Then
   Me.txtRuc.Text = Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 4)
End If



Me.frmEstadoCuenta.Visible = True
End Sub

Private Sub cmdEliminar_factura_Click()
If Me.HfCompras.Rows > 0 Then

strCadena = "call put_limpiar_anticipo('" & Val(Me.HfCompras.TextMatrix(Me.HfCompras.Row, 0)) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
strCadena = "SELECT * FROM view_anticipo_temporal WHERE anticipo='no' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
Call Me.llenar_anticipo(Me.HfCompras)
End If
End Sub

Private Sub cmdEstadoCuenta_Click()
Dim arr(0 To 2, 1 To 2) As String
Dim param As Variant
Dim in_fecha As Date
Dim in_ruc As String


If Trim(Me.txtRuc.Text) = "" Then
   Me.txtRuc.Text = Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 4)
End If

If Len(Trim(Me.txtRuc.Text)) = 8 Then
   in_ruc = "10" & Trim(Me.txtRuc.Text)
   in_ruc = DigitoVerificadorRUC(Trim(in_ruc))

End If

arr(0, 1) = "in_ruc"
arr(0, 2) = in_ruc
param = arr()


If Me.chk_emision.Value = 1 Then
   in_fecha = Format(Me.DtpInicio.Value, "YYYY-mm-dd")
   in_parametro = " and fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "'"
Else
   in_fecha = KEY_FECHA
   in_parametro = ""
End If


If Trim(Me.txtRuc.Text) <> "" Then
    strCadena = "SELECT id_compra,fecha_emision,fecha_cancelacion,dias,'-',numero,comprobante,id_proveedor,nproveedor,'-',celular,total, if (id_doc='0007',(total-function_pago_factura(id_compra,'" & Format(in_fecha, "YYYY-mm-dd") & "',id_moneda,ruc))*-1,(total-function_pago_factura(id_compra,'" & Format(in_fecha, "YYYY-mm-dd") & "',id_moneda,ruc)))   as saldo  ,anulado,id_moneda,ruc,'" & Val(Me.txtTipo_cambio.Text) & "',id_alm,id_doc,'-',moneda,simbolo FROM view_cuentas_cobrar WHERE (total-function_pago_factura(id_compra,'" & Format(in_fecha, "YYYY-mm-dd") & "',id_moneda,ruc))<>0  and  anulado='no'  and id_proveedor LIKE '%" & Trim(Me.txtRuc.Text) & "%' AND ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT id_compra,fecha_emision,fecha_cancelacion,dias,'-',numero,comprobante,id_proveedor,nproveedor,'-',celular,total, if (id_doc='0007',(total-function_pago_factura(id_compra,'" & Format(in_fecha, "YYYY-mm-dd") & "',id_moneda,ruc))*-1,(total-function_pago_factura(id_compra,'" & Format(in_fecha, "YYYY-mm-dd") & "',id_moneda,ruc)))   as saldo  ,anulado,id_moneda,ruc,tc,id_alm,id_doc,'-',moneda,simbolo FROM view_cuentas_cobrar WHERE (total-function_pago_factura(id_compra,'" & Format(in_fecha, "YYYY-mm-dd") & "',id_moneda,ruc))<>0 and   anulado='no'  and nproveedor LIKE '%" & Trim(Me.TxtCliente.Text) & "%' AND ruc='" & KEY_RUC & "'"
End If

Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_cliente_detalle_cobranza_letras", param, App.Path + "\Reportes\")


End Sub

Private Sub CmdEstadoCuentaProveedor_Click()
Dim arr(0 To 2, 1 To 2) As String
Dim in_doc As String
Dim param As Variant

arr(0, 1) = "fecha_ini"
arr(1, 1) = "fecha_fin"
arr(2, 1) = "telefono"

arr(0, 2) = get_direccion(Trim(Me.txtRuc.Text))
arr(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
arr(2, 2) = get_telefono(Trim(Me.txtRuc.Text))
param = arr()

If Me.chk_pendiente_pago.Value = 1 Then
    strCadena = "call ADM_EstadoCuenta('4','" & Trim(Me.txtRuc.Text) & "','" & Val(Me.txtTipo_cambio.Text) & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptEstadodeCuenta", param, App.Path + "\Reportes\")
    Exit Sub
End If

If Me.chk_credito.Value = 1 Then
    strCadena = "call ADM_EstadoCuenta('2','" & Trim(Me.txtRuc.Text) & "','" & Val(Me.txtTipo_cambio.Text) & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptEstadodeCuenta", param, App.Path + "\Reportes\")
End If

If Me.chk_todos.Value = 1 Then
    strCadena = "call ADM_EstadoCuenta('1','" & Trim(Me.txtRuc.Text) & "','" & Val(Me.txtTipo_cambio.Text) & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptEstadodeCuenta", param, App.Path + "\Reportes\")
End If



End Sub

Private Sub cmdGeneral_Click()
Dim cam3(0 To 2, 1 To 2)  As String


Dim param As Variant


                    cam3(0, 1) = "fecha_ini"
                    cam3(1, 1) = "fecha_fin"
                    cam3(2, 1) = "cambio"
                    
                   If Me.Check1.Value = 1 Then
                        cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
                   Else
                        cam3(0, 2) = "INICIO"
                   End If
                   
                   cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
                   cam3(2, 2) = KEY_VENDEDOR
                   param = cam3()
                    
                    



If Me.chk_emision.Value = 1 Then
    in_operacion = "1"
Else
    in_operacion = "3"
End If

strCadena = "CALL CON_CuentaPagarV2_LST('" & in_operacion & "','" & Trim(Me.txtRuc.Text) & "','" & Me.TxtcodProveedor.Text & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & Me.DtcComprobante.BoundText & "','" & Trim(Me.txtNumero.Text) & "','" & in_serie & "','" & Val(Me.txtTipo_cambio.Text) & "','" & Me.DtcBusquedaPeriodo.BoundText & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_cliente_detalle_cobranza", param, App.Path + "\Reportes\")
Exit Sub
       



'strCadena = "SELECT fecha_emision,fecha_cancelacion,comprobante,id_proveedor,nproveedor,total,(total-function_pago_factura(id_compra,'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "',id_moneda,ruc)) as saldo,id_moneda,'" & Val(Me.txtTipo_cambio.Text) & "',id_doc,doc_des FROM view_cuentas_cobrar WHERE  id_proveedor='20100675537' and  (total-function_pago_factura(id_compra,'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "',id_moneda,ruc))<>0 and    fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and id_proveedor LIKE '%" & Trim(tXTrUC.Text) & "%' and ruc='" & KEY_RUC & "' "
strCadena = "SELECT fecha_emision,fecha_cancelacion,comprobante,id_proveedor,nproveedor,total,(total-function_pago_factura(id_compra,'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "',id_moneda,ruc)) as saldo,id_moneda,'" & Val(Me.txtTipo_cambio.Text) & "',id_doc,doc_des FROM view_cuentas_cobrar WHERE   (total-function_pago_factura(id_compra,'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "',id_moneda,ruc))<>0 and    fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and id_proveedor LIKE '%" & Trim(txtRuc.Text) & "%' and ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)

Ans = ShowMultiReport(rst, "rpt_cliente_detalle_cobranza", param, App.Path + "\Reportes\")



End Sub

Private Sub cmdhistorial_Click()
'Dim arr(0 To 2, 1 To 2) As String
'Dim param As Variant
'arr(0, 1) = "in_cliente"
'arr(0, 2) = Trim(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 6))
'param = arr()
  

Dim arr(0 To 2, 1 To 2) As String
Dim param As Variant

arr(0, 1) = "in_cliente"
arr(1, 1) = "id_moneda"
arr(0, 2) = Trim(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 5))
arr(1, 2) = Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 6)

param = arr()

          
strCadena = "SELECT fecha_emision,id_proveedor,documento,fecha_origen,recibo,id_moneda,monto_pagado,tc,nombre_completo,forma_pago,id_tarjeta_operacion,total FROM view_historial_pago_compra WHERE id_venta='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptHistorial_venta", param, App.Path + "\Reportes\")


'Exit Sub

'Dim arr(0 To 2, 1 To 2) As String
'Dim param As Variant

'arr(0, 1) = "in_cliente"
'arr(1, 1) = "id_moneda"
'arr(0, 2) = Trim(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 6))
'arr(1, 2) = Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 7)

'param = arr()
  
       'strCadena = "SELECT fecha_emision,id_cliente,documento,fecha_origen,recibo,id_moneda,monto_pagado,tc,nombre_completo,forma_pago,id_tarjeta_operacion,total FROM view_historial_pago_v2 WHERE id_venta='" & Val(Me.HfgFacturas.TextMatrix(HfgFacturas.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
        '  Call ConfiguraRst(strCadena)
        '  Ans = ShowMultiReport(rst, "RptHistorial_venta", param, App.Path + "\Reportes\")
          
          
  
 
  
  
End Sub
Private Sub get_letra_cancelacion(ByVal in_serie As String, ByVal in_numero As String, ByVal in_ruc As String, ByVal in_monto As Single, ByVal in_emision As Date, ByVal in_factura As String, ByVal in_canje As String)
                    Documento = "LP:" & ":" & in_serie & "-" & in_numero
                    
                    strCadena = "CALL P_insert_venta('0417','" & KEY_ALM & "','0','" & Me.DtcMoneda.BoundText & "','no'," & _
                    "'" & in_serie & "','" & in_numero & "','" & in_ruc & "','" & get_persona(in_ruc) & "','0','0','0','" & in_monto & "','0'," & _
                    "'" & in_monto & "','0','" & Format(in_emision, "YYYY-mm-dd") & "','" & Format(in_emision, "YYYY-mm-dd") & "','" & id_tipo_factura & "','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & Val(Me.txtTc.Text) & "','no','" & formato_item(Month(in_emision), 2) & "','" & Year(in_emision) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstP(strCadena)
                    id_venta = rstP(0)
                    
                    in_detalle = "CANJE :" & in_factura
                    
                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & id_venta & "','00','" & in_detalle & "','-','1','" & in_monto & "','0','" & in_monto & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & id_venta & "','" & Val(Me.lbl_idcompra.Caption) & "','" & in_monto & "','" & in_monto & "','" & Me.DtcMoneda.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtTc.Text) & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "UPDATE movimiento_venta SET id_referencia='" & Val(Me.lbl_idcompra.Caption) & "' WHERE id_venta='" & id_venta & "'"
                    CnBd.Execute (strCadena)
                    
                   ' strCadena = "INSERT INTO movimiento_venta_monto(id_venta,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,banco,cheque,cuenta_contable,forma_pago_contable,flujo_caja,id_cuenta_origen,cta_redondeo,cta_anticipo,monto_redondeo,monto_anticipo,ruc)VALUES " & _
                    "('" & id_venta & "','01','" & get_forma_pago_anterior(Me.DtcMoneda.BoundText) & "','" & Val(Me.TxtMontoPago.Text) & "','" & Val(Me.TxtMontoPago.Text) * -1 & "','00','-','" & Trim(Me.txtOperacion.Text) & "','-','" & Me.DtcCheque.BoundText & "','" & get_cuenta_contable_cuenta(Me.DtcCuentas.BoundText) & "','" & DtcFormaPago.BoundText & "','" & Me.DtcFlujo.BoundText & "','" & Me.DtcCuentas.BoundText & "','" & Trim(Me.txtCuenta_redondeo.Text) & "','" & Trim(Me.txtCuenta_anticipo.Text) & "','" & Val(Me.txtMontoRedondeo.Text) & "','" & Val(Me.txtMontoAnticipo.Text) & "','" & KEY_RUC & "')"
                   ' CnBd.Execute (strCadena)
                   
                   
                   
                   'Put procesar letras de facturas
                   
                   'strCadena = "SELECT * FROM movimiento_compra WHERE id_canje='" & in_canje & "'"
                   
                   
                   
                   
                   
                   
End Sub
Private Function get_id_canje() As Double
strCadena = "SELECT id_canje FROM movimiento_compra_canje_letra ORDER BY id_canje DESC LIMIT 1"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   get_id_canje = rstLocal("id_canje") + 1
Else
    get_id_canje = 1
End If
End Function


Private Sub put_factura_involucrada(ByVal in_canje As String)
strCadena = "SELECT id_doc,`id_compra`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`ruc`, function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,monto_pagar,reversion,id_proveedor,fecha_emision,fecha_cancelacion,serie, numero FROM view_cuentas_cobrar WHERE seleccion='si' and  dni_save_pago='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
   For i = 0 To rstLocal.RecordCount - 1
   
            
        
        
        If rstLocal("id_moneda") = "00002" Then
           If Me.DtcMoneda.BoundText = rstLocal("id_moneda") Then
             in_monto_pagar = rstLocal("monto_pagar") / rstLocal("tc")
           Else
             in_monto_pagar = rstLocal("monto_pagar") * rstLocal("tc")
           End If
        Else
          in_monto_pagar = rstLocal("monto_pagar")
        
       End If
        strCadena = "INSERT INTO movimiento_compra_canje_letra(id_canje,id_compra,id_doc,id_moneda,monto,saldo_interno,id_proveedor,emision,vencimiento,serie,numero,cta_redondeo,monto_redondeo,ruc)VALUES " & _
        "('" & Val(in_canje) & "','" & rstLocal("id_compra") & "','" & rstLocal("id_doc") & "','" & rstLocal("id_moneda") & "','" & in_monto_pagar & "','" & in_monto_pagar & "','" & rstLocal("id_proveedor") & "','" & Format(rstLocal("fecha_emision"), "YYYY-mm-dd") & "','" & Format(rstLocal("fecha_cancelacion"), "YYYY-mm-dd") & "','" & rstLocal("serie") & "','" & rstLocal("numero") & "','" & Trim(Me.txtCuenta_redondeo.Text) & "','" & Val(Me.txtMontoRedondeo.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
      
        strCadena = "UPDATE movimiento_compra SET id_canje='" & Val(in_canje) & "' WHERE id_compra='" & rstLocal("id_compra") & "'"
        CnBd.Execute (strCadena)
        rstLocal.MoveNext
   Next i
End If
End Sub




Private Sub cmdProcesar_Click()
Dim in_canje As Double
Dim in_exonerado As Single
in_canje = get_id_canje
   
    
    
    
    strCadena = "SELECT * FROM movimiento_compra_letra WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        
        'INVOLUCRAR FACTURAS
        Call put_factura_involucrada(in_canje)
        
        For i = 0 To rst.RecordCount - 1
            
            If Me.DtcTipoCanje.BoundText = "0417" Then
                If Me.DtcMoneda.BoundText = "00001" Then
                    in_cta_compra = KEY_CTA_LETRA_PAGAR_SOLES
                End If
            
                If Me.DtcMoneda.BoundText = "00002" Then
                    in_cta_compra = KEY_CTA_LETRA_PAGAR_DOLARES
                End If
            Else
                
                
            If Me.DtcMoneda.BoundText = "00001" Then
                    in_cta_compra = KEY_CTA_FET_SOLES
                End If
            
                If Me.DtcMoneda.BoundText = "00002" Then
                    in_cta_compra = KEY_CTA_FET_DOLARES
                End If
            End If
            
        
           
           in_total = rst("monto")
           in_valor_venta = 0
           in_igv = 0
           in_exonerado = 0
        
        strCadena = "call P_insert_compra_ultimate('" & Me.DtcTipoCanje.BoundText & "','" & KEY_ALM & "','" & Format(rst("fecha"), "YYYY-mm-dd") & "','" & Format(rst("vencimiento"), "YYYY-mm-dd") & "','02'," & _
        "'03','--','" & Me.DtcMoneda.BoundText & "','" & formato_item(Month(rst("fecha")), 2) & "','" & Year(rst("fecha")) & "','" & rst("serie") & "'," & _
        "'" & Format(Trim(rst("numero")), "00000000") & "','6','" & rst("id_proveedor") & "','" & get_persona(rst("id_proveedor")) & "','" & Trim(Me.txtTc.Text) & "'," & _
        "'0','" & Val(in_valor_venta) & "','" & Val(in_igv) & "','0','0','0','0','" & Val(in_exonerado) & "','0','" & Val(in_total) & "','" & in_total & "'," & _
        " '" & KEY_USUARIO & "','LETRA POR PAGAR ','01','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','0','0','0','0','0','" & KEY_RUC & "')"
        Call ConfiguraRstP(strCadena)
        id_compra = rstP(0)
        
        Call put_generar_letra(in_canje, id_compra, rst("id_letra"), Me.DtcTipoCanje.BoundText)
        
        
        If Val(Me.txtMontoRedondeo.Text) <> 0 Then
            Me.txtMontoRedondeo.Text = 0
        End If
        
        rst.MoveNext
    Next i
      
      Call put_historial_pago(in_canje)
      
      
      strCadena = "call CON_InsertaAsiento_CanjeLetra('" & in_canje & "')"
      CnBd.Execute (strCadena)
      
      
     
      
      
      
   End If
    
    
   MsgBox "Canje Realizado con EXITO", vbInformation



End Sub
Private Sub put_generar_letra(ByVal in_canje As String, ByVal in_compra As String, ByVal in_letra_temporal As String, ByVal in_doc As String)
strCadena = "SELECT * FROM movimiento_compra_letra WHERE id_letra='" & Val(in_letra_temporal) & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   strCadena = "INSERT INTO movimiento_compra_canje_letra(id_canje,id_compra,id_doc,id_moneda,monto,saldo_interno,id_proveedor,emision,vencimiento,serie,numero,cta_redondeo,monto_redondeo,ruc)VALUES " & _
   "('" & in_canje & "','" & in_compra & "','" & in_doc & "','" & Me.DtcMoneda.BoundText & "','" & rstL("monto") & "','" & rstL("monto") & "','" & rstL("id_proveedor") & "','" & Format(rstL("fecha"), "YYYY-mm-dd") & "','" & Format(rstL("vencimiento"), "YYYY-mm-dd") & "','" & rstL("serie") & "','" & rstL("numero") & "','" & Trim(Me.txtCuenta_redondeo.Text) & "','" & Val(Me.txtMontoRedondeo.Text) & "','" & KEY_RUC & "')"
   CnBd.Execute (strCadena)
      
   strCadena = "UPDATE movimiento_compra SET id_canje='" & Val(in_canje) & "' WHERE id_compra='" & Val(in_compra) & "'"
   CnBd.Execute (strCadena)
   
End If



End Sub
Private Function put_generar_letra_venta(ByVal in_canje As String, ByVal in_compra As String, ByVal in_total As Single, ByVal in_tc As Single, ByVal in_doc As String) As Single
Dim in_residuo As Single
Dim in_referencia As String
strCadena = "SELECT * FROM movimiento_compra_canje_letra WHERE id_doc='" & in_doc & "'and  saldo_interno>0 and  id_canje='" & in_canje & "'  and ruc='" & KEY_RUC & "'"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   rstP.MoveFirst
   in_residuo = 0
   For i = 0 To rstP.RecordCount - 1
                         
                    
                    
                    If in_total > 0 Then
                    
                    
                        If in_total <= rstP("monto") Then
                            in_monto = in_total
                        Else
                            in_monto = rstP("monto")
                        End If
                    
                    Documento = "LP:" & ":" & rstP("serie") & "-" & rstP("numero")
                    
                    strCadena = "CALL P_insert_venta_cancelacion('" & in_doc & "','" & KEY_ALM & "','0','" & rstP("id_moneda") & "','no'," & _
                    "'" & rstP("serie") & "','" & rstP("numero") & "','" & rstP("id_proveedor") & "','" & get_persona(rstP("id_proveedor")) & "','0','0','0','" & in_monto & "','0'," & _
                    "'" & in_monto & "','0','" & Format(rstP("emision"), "YYYY-mm-dd") & "','" & Format(rstP("vencimiento"), "YYYY-mm-dd") & "','01','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & Val(Me.txtTc.Text) & "','no','" & formato_item(Month(rstP("emision")), 2) & "','" & Year(rstP("emision")) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstlocal(strCadena)
                    id_venta = rstLocal(0)
                    
                    strCadena = "UPDATE movimiento_venta SET numero='" & rstP("numero") & "',id_canje='" & Val(in_canje) & "' WHERE id_venta='" & id_venta & "'"
                    CnBd.Execute (strCadena)
                    
                    in_detalle = "CANJE :" & get_comprobante_compra(in_compra)
                    
                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & id_venta & "','00','" & in_detalle & "','-','1','" & in_monto & "','0','" & in_monto & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & id_venta & "','" & Val(in_compra) & "','" & in_monto & "','" & in_monto & "','" & Me.DtcMoneda.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtTc.Text) & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "UPDATE movimiento_compra SET id_canje='" & in_canje & "' WHERE id_compra='" & Val(in_compra) & "'"
                    CnBd.Execute (strCadena)
                    in_total = in_total - rstP("monto")
                    
                    
                    strCadena = "UPDATE movimiento_compra_canje_letra SET saldo_interno=saldo_interno-'" & in_monto & "' WHERE id_detalle='" & rstP("id_detalle") & "'"
                    CnBd.Execute (strCadena)
                    
                    put_generar_letra_venta = in_total
                   
                       
                   End If
                    
    
            rstP.MoveNext
   Next i
 Else
    
                    in_monto = in_total
                    strCadena = "SELECT * FROM movimiento_compra_canje_letra WHERE id_canje='" & in_canje & "'  and ruc='" & KEY_RUC & "' ORDER BY monto DESC LIMIT 1"
                    Call ConfiguraRstP(strCadena)
                    If rstP.RecordCount > 0 Then
                    
                    Documento = "LP:" & ":" & rstP("serie") & "-" & rstP("numero")
                    strCadena = "CALL P_insert_venta('" & in_doc & "','" & KEY_ALM & "','0','" & rstP("id_moneda") & "','no'," & _
                    "'" & rstP("serie") & "','" & rstP("numero") & "','" & rstP("id_proveedor") & "','" & get_persona(rstP("id_proveedor")) & "','0','0','0','" & in_monto & "','0'," & _
                    "'" & in_monto & "','0','" & Format(rstP("emision"), "YYYY-mm-dd") & "','" & Format(rstP("vencimiento"), "YYYY-mm-dd") & "','01','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & Val(Me.txtTc.Text) & "','no','" & formato_item(Month(rstP("emision")), 2) & "','" & Year(rstP("emision")) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstlocal(strCadena)
                    id_venta = rstLocal(0)
                    
                    strCadena = "UPDATE movimiento_venta SET numero='" & rstP("numero") & "',id_canje='" & Val(in_canje) & "' WHERE id_venta='" & id_venta & "'"
                    CnBd.Execute (strCadena)
                    
                    in_detalle = "CANJE :" & get_comprobante_compra(in_compra)
                    
                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & id_venta & "','00','" & in_detalle & "','-','1','" & in_monto & "','0','" & in_monto & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & id_venta & "','" & Val(in_compra) & "','" & in_monto & "','" & in_monto & "','" & Me.DtcMoneda.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtTc.Text) & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "UPDATE movimiento_compra SET id_canje='" & in_canje & "' WHERE id_compra='" & Val(in_compra) & "'"
                    CnBd.Execute (strCadena)
                    in_total = in_total - rstP("monto")
                    
                    
                    strCadena = "UPDATE movimiento_compra_canje_letra SET saldo_interno='" & in_total & "' WHERE id_detalle='" & rstP("id_detalle") & "'"
                    CnBd.Execute (strCadena)
                    
                        
                        
                    End If
                    
                    
     
       
End If

End Function

Private Function get_comprobante_compra(ByVal in_compra As String)
strCadena = "SELECT comprobante FROM view_cuentas_cobrar WHERE id_compra='" & Val(in_compra) & "'"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
    get_comprobante_compra = rstLocal("comprobante")
Else
    get_comprobante_compra = "-"
End If

End Function

Private Sub put_historial_pago(ByVal in_canje As String)
Dim in_saldo As Single
strCadena = "SELECT id_compra,monto_pagar,tc,id_moneda,id_doc FROM movimiento_compra WHERE seleccion='si' and dni_save_pago='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
   rstT.MoveFirst
   For i = 0 To rstT.RecordCount - 1
      If rstT("monto_pagar") > 0 Then
            If rstT("id_moneda") = "00002" Then
                in_monto_pagar = rstT("monto_pagar") / rstT("tc")
            Else
                in_monto_pagar = rstT("monto_pagar")
            End If
            
            in_saldo = put_generar_letra_venta(in_canje, rstT("id_compra"), in_monto_pagar, rstT("tc"), Me.DtcTipoCanje.BoundText)
       End If
       rstT.MoveNext
   Next i
End If
End Sub


Private Sub cmdProcesarAjuste_Click()

If MsgBox("Esta seguro de Realizar el ajuste [TC] para:" + Chr(13) + Chr(13) + "TIPO CAMBIO:" & str(Me.TxtTipoCambio.Text) + Chr(13) + "PERIODO AJUSTE:" + Me.DtcPeriodoAjuste.Text, vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then

If KEY_RUC = "20128836251" Then
    strCadena = "SELECT * FROM con_periodo WHERE id='" & Me.DtcPeriodoAjuste.BoundText & "' LIMIT 1"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
       
       
    
    If Trim(Me.TxtCliente.Text) <> "" Then
       strCadena = "SELECT id_proveedor as id_cliente,id_doc,serie,numero,comprobante FROM view_cuentas_cobrar WHERE fecha_emision>='" & Format(rst("FechaInicio"), "YYYY-mm-dd") & "' and fecha_emision<='" & Format(rst("FechaFin"), "YYYY-mm-dd") & "' and id_proveedor='" & Trim(Me.TxtCliente.Text) & "' and id_moneda='00002' and  ruc='" & KEY_RUC & "'"
    Else
       strCadena = "SELECT id_proveedor as id_cliente, id_doc,serie,numero,comprobante FROM view_cuentas_cobrar WHERE fecha_emision>='" & Format(rst("FechaInicio"), "YYYY-mm-dd") & "' and fecha_emision<='" & Format(rst("FechaFin"), "YYYY-mm-dd") & "' and  id_moneda='00002'  and  ruc='" & KEY_RUC & "'"
    End If
  End If
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        Me.prg_avance.Min = 0
        Me.prg_avance.Max = rst.RecordCount - 1
        For i = 0 To rst.RecordCount - 1
            
            strCadena = "call CON_InsertaAsiento_AjusteTC_Persona_comprobante('" & KEY_RUC & "','" & Me.DtcPeriodoAjuste.BoundText & "','" & Trim(Me.txtCuentaPrincipal.Text) & "','" & Trim(rst("id_cliente")) & "','" & Val(Me.TxtTipoCambio.Text) & "','" & Val(rst("id_doc")) & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("comprobante") & "','" & Trim(Me.txtCuentaPerdida.Text) & "','" & Trim(Me.txtCuentaGanancia.Text) & "','" & KEY_USUARIO & "')"
            CnBd.Execute (strCadena)
            rst.MoveNext
            prg_avance.Value = i
            DoEvents
        Next i
   
    End If


Else
    strCadena = "SELECT * FROM con_periodo WHERE id='" & Me.DtcPeriodoAjuste.BoundText & "' LIMIT 1"
    Call ConfiguraRstK(strCadena)
       
    If Me.chk_cta_contable.Value = 1 Then
        If Trim(Me.txtAjusteporCuenta.Text) <> "" And Trim(TxtCliente.Text) <> "" Then
            strCadena = "call CON_InsertaAsiento_AjusteTC_Empresa('" & KEY_RUC & "','" & Me.DtcPeriodoAjuste.BoundText & "','" & Trim(Me.txtAjusteporCuenta.Text) & "','" & Trim(TxtCliente.Text) & "','" & Val(Me.TxtTipoCambio.Text) & "','" & KEY_USUARIO & "')"
            ConfiguraRstK (strCadena)
        End If
        MsgBox "AJUSTE DE TIPO CAMBIO REALIZADO", vbInformation, KEY_VENDEDOR
        Exit Sub
    End If
    
    If Trim(TxtCliente.Text) <> "" Then
        strCadena = "SELECT DISTINCT id_proveedor FROM view_cuentas_cobrar WHERE id_proveedor='" & Trim(Me.TxtCliente.Text) & "' and  id_moneda='00002' and  (total-function_pago_factura(id_compra,'" & Format(rstK("FechaFin"), "YYYY-mm-dd") & "',id_moneda,ruc)) >0.1  and fecha_emision<= '" & Format(rstK("FechaFin"), "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'"
    Else
        strCadena = "SELECT DISTINCT id_proveedor FROM view_cuentas_cobrar WHERE id_moneda='00002' and  (total-function_pago_factura(id_compra,'" & Format(rstK("FechaFin"), "YYYY-mm-dd") & "',id_moneda,ruc)) >0.1  and fecha_emision<= '" & Format(rstK("FechaFin"), "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'"
    End If
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       For i = 0 To rst.RecordCount - 1
           
           
           strCadena = "call CON_InsertaAsiento_AjusteTC_Empresa('" & KEY_RUC & "','" & Me.DtcPeriodoAjuste.BoundText & "','4212','" & Trim(rst("id_proveedor")) & "','" & Val(Me.TxtTipoCambio.Text) & "','" & KEY_USUARIO & "')"
           CnBd.Execute (strCadena)
           
           rst.MoveNext
       Next i
    End If
End If





MsgBox "AJUSTE DE TIPO CAMBIO REALIZADO", vbInformation, KEY_VENDEDOR
End If


End Sub

Private Sub Command1_Click()
strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='0090' and  ruc='" & KEY_RUC & "' ORDER BY id_compra ASC"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
   rstIN.MoveFirst
   For i = 0 To rstIN.RecordCount - 1
        
        If rstIN("id_doc") = "0090" Then
           in_doc = "SALALMA:"
        End If
         
        in_glosa = Trim(in_doc & rstIN("serie") & "-" & rstIN("numero") & Space(1) & rstIN("nproveedor"))
        
        Call delete_asientos(in_glosa)
        
        strCadena = "CON_InsertaAsiento_salida_internacional('" & rstIN("id_compra") & "')"
        CnBd.Execute (strCadena)
        rstIN.MoveNext
        DoEvents
   Next i
End If
End Sub
Private Sub delete_asientos(ByVal in_glosa As String)
strCadena = "SELECT * FROM con_asiento WHERE Activo='1' and  Glosa LIKE '%" & Trim(in_glosa) & "%' and IdTipoAsiento IN('1CIX000000000145') and idEmpresaSis='" & KEY_RUC & "'"
       Call ConfiguraRstL(strCadena)
       If rstL.RecordCount > 0 Then
          rstL.MoveFirst
          For j = 0 To rstL.RecordCount - 1
                strCadena = "DELETE FROM  con_asiento  where id = '" & rstL("id") & "'"
                CnBd.Execute (strCadena)
                
                strCadena = "SELECT * FROM con_asientomovimiento where IdAsiento = '" & rstL("id") & "'"
                Call ConfiguraRstA(strCadena)
                If rstA.RecordCount > 0 Then
                   rstA.MoveFirst
                   For k = 0 To rstA.RecordCount - 1
                       strCadena = "DELETE FROM  con_asientomovimiento  where Id = '" & rstA("id") & "' and idEmpresaSis='" & KEY_RUC & "'"
                       CnBd.Execute (strCadena)
                       
                       strCadena = "DELETE FROM  con_asientomovimiento_documento  where IdAsientoMovimiento = '" & rstA("id") & "' and idEmpresaSis='" & KEY_RUC & "'"
                       CnBd.Execute (strCadena)
                       
                      
                       rstA.MoveNext
                   Next k
                End If
                rstL.MoveNext
     Next j
    End If
End Sub
Private Sub DtcPeriodoAjuste_Change()
 

Dim in_fecha As Date
strCadena = "SELECT * FROM con_periodo WHERE id='" & Me.DtcPeriodoAjuste.BoundText & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    in_fecha = DateSerial(rst("Ejercicio"), rst("mes") + 1, 0)
    Me.TxtTipoCambio.Text = get_tipo_cambio_dia(in_fecha, "valor_venta")
End If



End Sub

Private Sub DtcTipoComprobante_KeyPress(KeyAscii As Integer)
Dim in_operacion As String
If KeyAscii = 13 Then
    

If Me.chk_emision.Value = 1 Then
    in_operacion = 2
Else
    in_operacion = 4
End If



strCadena = "CALL CON_CuentaPagarV2_LST('" & in_operacion & "','" & Trim(Me.txtRuc.Text) & "','" & Trim(TxtcodProveedor.Text) & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & Me.DtcComprobante.BoundText & "','" & Trim(Me.txtNumero.Text) & "','','" & Val(Me.txtTipo_cambio.Text) & "','" & Me.DtcBusquedaPeriodo.BoundText & "','" & KEY_RUC & "')"
Call llenarGrid(Me.HfgFacturas)

End If
End Sub

Private Sub Image3_Click()
Me.frmajusteCobrar.Visible = False
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Procedencia = Selecionar
   FrmPersona.Show
   Exit Sub
End If
End Sub

Private Sub txtCuentaPerdida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = seleccionar_per
    FrmPlanContableCuentas.Show
    Exit Sub
End If
End Sub

Private Sub txtCuentaGanancia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = seleccionar_soldadura
    FrmPlanContableCuentas.Show
    Exit Sub
End If
End Sub


Private Sub cmdRegenerar_Click()




strCadena = "CALL p_insert_compra_emitido_premiun('" & Val(HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) & "')"
Call ConfiguraRst(strCadena)

MsgBox "Proceso Correcto", vbInformation

End Sub

Private Sub txtCuentaPrincipal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = buscar
    FrmPlanContableCuentas.Show
    Exit Sub
End If
End Sub


Private Sub cmdrevertir_anticipo_Click()
Dim in_anticipo As Double

strCadena = "SELECT * FROM movimiento_compra_anticipo order by id_canje_anticipo DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   in_anticipo = rst("id_canje_anticipo") + 1
Else
   in_anticipo = 1
End If
strCadena = "SELECT * FROM movimiento_compra_anticipo_temporal WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       If rst("monto_compra") < 0 Then
          in_monto_compra = rst("monto_compra") * -1
       Else
          in_monto_compra = rst("monto_compra")
       End If
       strCadena = "INSERT INTO movimiento_compra_anticipo(id_doc,id_canje_anticipo,id_compra,monto_compra,monto_saldo,anticipo,cta_redondeo,monto_redondeo,dni_save,ruc)VALUES " & _
       "('" & rst("id_doc") & "','" & in_anticipo & "','" & rst("id_compra") & "','" & in_monto_compra & "','" & in_monto_compra & "','" & rst("anticipo") & "','" & Trim(Me.txtcuenta_redondeo_anticipo.Text) & "','" & Val(Me.txtmonto_anticipo.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       If Trim(Me.txtcuenta_redondeo_anticipo.Text) <> "" Then
          Me.txtcuenta_redondeo_anticipo.Text = ""
          Me.txtmonto_anticipo.Text = 0
       End If
       rst.MoveNext
   Next i
   
End If


Call put_revertir_factura_anticipo(in_anticipo)


If KEY_PAIS <> KEY_PERU Then
    strCadena = "call CON_InsertaAsiento_CanjeAnticipo_Internacional('" & in_anticipo & "')"
Else
    strCadena = "call CON_InsertaAsiento_CanjeAnticipo('" & in_anticipo & "')"
End If
CnBd.Execute (strCadena)



MsgBox "Proceso Realizado con Exito", vbInformation
End Sub


Private Sub put_revertir_factura_anticipo(ByVal in_canje_anticipo As String)
' CANCELAR LAS FACTURAS CON LOS ANTICIPOS
Dim in_saldo_factura As Double
strCadena = "SELECT id_compra,monto_compra FROM movimiento_compra_anticipo WHERE id_doc<>'0419' and   id_canje_anticipo='" & in_canje_anticipo & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        in_saldo_factura = put_update_saldo_letra(in_canje_anticipo, rst("id_compra"), rst("monto_compra"))
        rst.MoveNext
   Next i
End If


strCadena = "SELECT id_compra as id_anticipo,id_compra_referencia as id_compra,monto_saldo FROM movimiento_compra_anticipo WHERE monto_saldo<>0 and id_doc='0419' and id_canje_anticipo='" & in_canje_anticipo & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        in_saldo_factura = put_update_saldo_letra_finish(in_canje_anticipo, rst("id_compra"), rst("id_anticipo"), Abs(rst("monto_saldo")))
        rst.MoveNext
   Next i
End If



End Sub


Function put_update_saldo_letra(ByVal in_anticipo As String, ByVal in_venta As String, ByVal in_monto_factura As Double) As Single

strCadena = "SELECT * FROM movimiento_compra_anticipo WHERE monto_saldo>0 and  id_doc='0419' and   id_canje_anticipo='" & in_anticipo & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   rstP.MoveFirst
   For i = 0 To rstP.RecordCount - 1
        If in_monto_factura > 0 Then
                If in_monto_factura > rstP("monto_saldo") Then
                    in_monto_cancelar = rstP("monto_saldo")
                Else
                    in_monto_cancelar = in_monto_factura
                End If
                strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & rstP("id_compra") & "'"
                Call ConfiguraRstT(strCadena)
                If rstT.RecordCount > 0 Then
                    in_ref = get_documento_abrev(rstT("id_doc")) & ":" & rstT("serie") & "-" & rstT("numero")
                    Call put_cancelar_comprobante(rstT("serie"), rstT("numero"), rstT("id_doc"), rstT("id_proveedor"), rstT("id_moneda"), in_monto_cancelar, rstT("tc"), in_venta, in_ref, rstT("fecha_emision"), rstT("fecha_cancelacion"), in_anticipo)
                    
                    Call put_cancelar_comprobante(rstT("serie"), rstT("numero"), rstT("id_doc"), rstT("id_proveedor"), rstT("id_moneda"), in_monto_cancelar, rstT("tc"), rstP("id_compra"), in_ref, rstT("fecha_emision"), rstT("fecha_cancelacion"), in_anticipo)
                    
                    
                    
                    strCadena = "UPDATE movimiento_compra_anticipo SET id_compra_referencia='" & in_venta & "',monto_saldo=monto_saldo-'" & Val(in_monto_cancelar) & "' WHERE id='" & rstP("id") & "' "
                    CnBd.Execute (strCadena)
                    
                    strCadena = "UPDATE movimiento_compra_anticipo SET id_compra_referencia='" & rstP("id_compra") & "', monto_saldo=monto_saldo-'" & Val(in_monto_cancelar) & "' WHERE id_compra='" & in_venta & "' "
                    CnBd.Execute (strCadena)
                    
                    
                    strCadena = "UPDATE movimiento_compra SET reversion='si' WHERE id_compra='" & rstP("id_compra") & "' "
                    CnBd.Execute (strCadena)
                End If
                in_monto_factura = in_monto_factura - in_monto_cancelar
                put_update_saldo_letra = in_monto_factura
    End If
    rstP.MoveNext
   Next i
End If
End Function


Function put_update_saldo_letra_finish(ByVal in_anticipo As String, ByVal in_venta As String, ByVal id_anticipo As String, ByVal in_monto_factura As Double) As Single
strCadena = "SELECT * FROM movimiento_compra_anticipo WHERE monto_saldo>0 and  id_doc='0419' and   id_canje_anticipo='" & in_anticipo & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   rstP.MoveFirst
   For i = 0 To rstP.RecordCount - 1
        If in_monto_factura > 0 Then
                If in_monto_factura > rstP("monto_saldo") Then
                    in_monto_cancelar = rstP("monto_saldo")
                Else
                    in_monto_cancelar = in_monto_factura
                End If
                strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & in_venta & "'"
                Call ConfiguraRstT(strCadena)
                If rstT.RecordCount > 0 Then
                    in_ref = get_documento_abrev(rstT("id_doc")) & ":" & rstT("serie") & "-" & rstT("numero")
                    Call put_cancelar_comprobante(rstT("serie"), rstT("numero"), rstT("id_doc"), rstT("id_proveedor"), rstT("id_moneda"), in_monto_cancelar, rstT("tc"), id_anticipo, in_ref, rstT("fecha_emision"), rstT("fecha_cancelacion"), in_anticipo)
                    
                    strCadena = "UPDATE movimiento_compra_anticipo SET id_compra_referencia='" & in_venta & "',monto_compra=monto_compra-'" & Val(in_monto_cancelar) & "',monto_saldo=monto_saldo-'" & Val(in_monto_cancelar) & "' WHERE id='" & rstP("id") & "' "
                    CnBd.Execute (strCadena)
                    
                    'strCadena = "UPDATE movimiento_compra_anticipo SET id_compra_referencia='" & rstP("id_compra") & "', monto_saldo=monto_saldo-'" & Val(in_monto_cancelar) & "' WHERE id_compra='" & in_venta & "' "
                    'CnBd.Execute (strCadena)
                    
                    
                    strCadena = "UPDATE movimiento_compra SET reversion='si' WHERE id_compra='" & rstP("id_compra") & "' "
                    CnBd.Execute (strCadena)
                End If
                in_monto_factura = in_monto_factura - in_monto_cancelar
                put_update_saldo_letra_finish = in_monto_factura
    End If
    rstP.MoveNext
   Next i
End If
End Function

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdupdate_Click()
If Val(Me.txtMontoLetra.Text) > 0 And Val(Me.lbl_idcompra.Caption) > 0 And Val(Me.txtTc.Text) > 0 And Me.DtcProveedorLetra.BoundText <> "" Then
   Me.cmdAnular.Visible = False
    Me.cmdDelete.Enabled = True
    strCadena = "INSERT INTO movimiento_compra_letra (fecha,vencimiento,serie,numero,monto,dni_save,id_proveedor,ruc)VALUES " & _
    "('" & Format(Me.DtpFechaLetra.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpVence.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtserie_letra.Text) & "','" & Format(Val(Me.TxtNumeroLetra.Text), "00000000") & "','" & Val(Me.txtMontoLetra.Text) & "','" & KEY_USUARIO & "','" & Me.DtcProveedorLetra.BoundText & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
    strCadena = "SELECT * FROM  movimiento_compra_letra WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' ORDER BY fecha ASC"
    Call llenar_letras(Me.HfLetras)
Else
    MsgBox "INGRESE DATOS VALIDOS", vbInformation
End If
End Sub



    







Private Function validar_reversion_anticipo(ByVal in_compra As String) As Boolean

strCadena = "SELECT * FROM movimiento_compra_anticipo WHERE id_compra='" & Val(in_compra) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtId_canje_anticipo.Text = rst("id_canje_anticipo")
   Me.cmdrevertir_anticipo.Visible = False
   Me.cmdAnularReversion.Visible = True
   validar_reversion_anticipo = True
   Me.frm_canje_anticipos.Visible = True
Else
    Me.cmdrevertir_anticipo.Visible = True
    Me.cmdAnularReversion.Visible = False
    validar_reversion_anticipo = False
End If

End Function


Private Sub cmdVincularAnticipo_Click()




If Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) > 0 Then
    
    
    
    
    strCadena = "call put_limpiar_anticipo('0','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
    
    If validar_reversion_anticipo(Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0))) = True Then
        Exit Sub
    End If
    
    
    

    strCadena = "SELECT `id_compra`,no_domiciliada,`id_moneda`,`tc`,`total`, function_pago_factura(id_compra,'" & KEY_FECHA & "',id_moneda,ruc) as pago,id_doc,reversion  FROM view_cuentas_cobrar WHERE id_compra='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        in_saldo = rst("total") - rst("pago")
        
        If rst("no_domiciliada") = "si" Then
            in_saldo = in_saldo '- in_saldo * 30 / 100
        End If
        
        strCadena = "call put_vincular_anticipo('" & rst("id_compra") & "','" & rst("id_doc") & "','" & in_saldo & "','no','" & KEY_USUARIO & "','" & KEY_RUC & "') "
        CnBd.Execute (strCadena)
        
        strCadena = "SELECT * FROM view_anticipo_temporal WHERE anticipo='no' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
        Call Me.llenar_anticipo(Me.HfCompras)
    End If
    Me.frm_canje_anticipos.Visible = True
End If
End Sub



Private Sub DtcBusquedaPeriodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "CALL CON_CuentaPagarV2_LST('6','" & Trim(Me.txtRuc.Text) & "','" & Trim(TxtcodProveedor.Text) & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & Me.DtcComprobante.BoundText & "','" & Trim(Me.txtNumero.Text) & "','','" & Val(Me.txtTipo_cambio.Text) & "','" & Me.DtcBusquedaPeriodo.BoundText & "','" & KEY_RUC & "')"

    Call llenarGrid(Me.HfgFacturas)
End If
End Sub

Private Sub DtcComprobante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtserie_factura)
End If
End Sub

Private Sub dtcProveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Me.DtcComprobante.BoundText & "' and serie='" & Trim(Me.txtserie_factura.Text) & "' and numero='" & Trim(Me.txtnumero_factura.Text) & "' and id_proveedor='" & Me.DtcProveedor.BoundText & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.txtid_compranueva.Text = rst("id_compra")
       Me.cmdagregar_factura.SetFocus
    End If
    
End If
End Sub

Private Sub DtpFechaLetra_Change()
Me.txtTc.Text = get_tipo_cambio_dia(CVDate(Me.DtpFechaLetra.Value), "valor_venta")
Me.DtcPeriodo.BoundText = get_periodo_actual(Me.DtpFechaLetra.Value)
End Sub

Private Sub DtpFin_Change()

Me.txtTipo_cambio.Text = get_tipo_cambio_dia(CVDate(Me.DtpFin.Value), "valor_venta")

End Sub

Private Sub Form_Load()
  CenterForm Me
  Me.Top = 50
  Me.DtpInicio.Value = KEY_FECHA
  Me.DtpFin.Value = KEY_FECHA
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  Me.DtcAlmacen.Enabled = False
  
  
  strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM moneda  ORDER BY descripcion "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMoneda)
  Me.DtcMoneda.BoundText = "00001"
  
  
  
  strCadena = "SELECT id_doc as Codigo, doc_abrev as Descripcion FROM comprobantes WHERE id_doc In('0417','0418')  ORDER BY descripcion "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoCanje)
  Me.DtcTipoCanje.BoundText = "0417"
  
  
  
  strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM comprobantes  ORDER BY doc_des "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoComprobante)
  Me.DtcTipoComprobante.BoundText = "0001"
  
  
  
  
  strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodo)
  
    strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcBusquedaPeriodo)
  Me.DtcBusquedaPeriodo.BoundText = get_periodo_actual(KEY_FECHA)
    
    
  Me.txtTipo_cambio.Text = get_tipo_cambio_dia(CVDate(Me.DtpFin.Value), "valor_venta")
  
  
  Call facturas

  
End Sub


Private Sub HfgFacturas_KeyPress(KeyAscii As Integer)
Dim facturas As String
Dim total_facturas As Single
If Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 7)) > 0 Then
    If KeyAscii = 13 Then
        If frmNuevoComprobante.Procedencia = buscar Then
            
            facturas = ""
            total_facturas = 0
            strCadena = "UPDATE DocumentoCompra set seleccion='F' WHERE cPersona='" & Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 8) & "' AND Ruc='" & KEY_RUC & "'"
            Call Execute_Sql(strCadena)
            strCadena = "UPDATE DocumentoCompra set seleccion='V' WHERE idCompra='" & Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 7) & "'"
            Call Execute_Sql(strCadena)
            strCadena = "SELECT * FROM DocumentoCompra WHERE idCompra='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 7)) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                frmNuevoComprobante.FrameComprobante.Visible = True
                frmNuevoComprobante.FrameFacturas.Visible = False
                frmNuevoComprobante.TxtTD.Text = rst("doc_cod")
                frmNuevoComprobante.txtserie.Text = rst("sSerie")
                frmNuevoComprobante.TxtMontoFactura.Text = rst("nTotalCompra")
                frmNuevoComprobante.txtNumero.Text = rst("cDocumentoCompra")
                frmNuevoComprobante.txtCodPersona.Text = rst("cPersona")
                frmNuevoComprobante.TxtGlosa.Text = "PAGO: " + Mid(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 2), 1, 3) + ":" + Mid(rst("sSerie"), 2, 4) + "-" + Mid(rst("cDocumentoCompra"), 5, 10)
                frmNuevoComprobante.TxtMonto1.Text = Format(rst("saldo"), "###0.00")
                frmNuevoComprobante.TlbAcciones.Buttons(KEY_SAVE).Enabled = True
                Call frmNuevoComprobante.precionar
                Call frmNuevoComprobante.LlenarVinculados(frmNuevoComprobante.HfVinculados, Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 8)))
                
                Call Resalta(frmNuevoComprobante.TxtMonto1)
                Unload Me
                frmNuevoComprobante.Procedencia = Neutro
                
                Exit Sub
            End If
        End If
    End If
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub




Private Sub HfgFacturas_SelChange()

If Me.HfgFacturas.Rows > 0 Then
   If Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 0)) > 0 Then
       Me.cmdamortizar.Enabled = True
   Else
       Me.cmdamortizar.Enabled = False
   End If
   Me.cmdhistorial.Enabled = True
Else
    Me.cmdamortizar.Enabled = False
    Me.cmdhistorial.Enabled = False
End If


End Sub
Public Sub LlenarPagos(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT * FROM  view_pagos_compra WHERE id_movimiento='" & idVenta & "' ORDER BY fecha_emision DESC"
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
            Grilla.ColWidth(2) = 1300
            Grilla.ColWidth(3) = 4000
            Grilla.ColWidth(4) = 1500
            Grilla.ColWidth(5) = 3500
            Grilla.ColWidth(6) = 3500
        Next
        cabecera = "CODIGO" & vbTab & "ITEM" & vbTab & "FECHA" & vbTab & "DOCUMENTO" & vbTab & "MONTO" & vbTab & "SUCURSAL" & vbTab & "RESPONSABLE"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_venta") & vbTab & Format((i + 1), "00") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & rst("documento") & vbTab & Format(rst("monto_inicial"), "#,##0.00") & vbTab & Format(rst("monto_pagado"), "#,##0.00") & vbTab & rst("nombre_completo")
            Grilla.AddItem Fila
            If rst("anulado") = "si" Then
                For k = 0 To 6
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
            Else
            tTotal = tTotal + rst("monto_pagado")
            End If
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL CANCELADO:" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
       For k = 0 To 3
            Grilla.col = 3
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
    
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Public Sub llenar_letras(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim tTotal As Double
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
      Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1400
            Grilla.ColWidth(2) = 1400
            Grilla.ColWidth(3) = 2000
            Grilla.ColWidth(4) = 2000
        Next
        cabecera = "CODIGO" & vbTab & "EMISION" & vbTab & "VENCIMIENTO" & vbTab & "NUMERO" & vbTab & "MONTO"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
       
        
        For i = 0 To rst.RecordCount - 1
            
            Fila = rst("id_letra") & vbTab & Format(rst("fecha"), "dd-mm-YYYY") & vbTab & Format(rst("vencimiento"), "dd-mm-YYYY") & vbTab & "LETRA  :  " & rst("serie") & "-" & rst("numero") & vbTab & Format(rst("monto"), "#,##0.00")
            Grilla.AddItem Fila
            
            tTotal = tTotal + rst("monto")
            
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL LETRAS:" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
       For k = 3 To 4
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H80FF&
      Next k
    
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub


Public Sub llenar_anticipo(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim tTotal As Double
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
            Grilla.ColWidth(2) = 2400
            Grilla.ColWidth(3) = 1400
            
        Next
        cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & " SALDO  "
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & rst("comprobante") & vbTab & Format(rst("monto_compra"), "#,##0.00")
            Grilla.AddItem Fila
            
            tTotal = tTotal + rst("monto_compra")
            If rst("id_doc") <> "0419" Then
                 Me.txtTotal_factura_reversion.Text = tTotal
            Else
                 Me.txttotal_anticipo_reversion.Text = tTotal
            End If
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "TOTAL :" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
       For k = 2 To 3
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H80FF&
      Next k
   
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub


Public Sub llenar_facturas(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim tTotal As Double

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
            Grilla.ColWidth(2) = 2100
            Grilla.ColWidth(3) = 3000
            Grilla.ColWidth(4) = 1200
            Grilla.ColWidth(5) = 1000
            Grilla.ColWidth(6) = 1300
        Next
        
        cabecera = "IDCOMPRA" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "PROVEEDOR" & vbTab & "MONEDA" & vbTab & " SALDO " & vbTab & " SALDO [ SOLES ]"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        
        rst.MoveFirst
        tTotal = 0
        tTotaldolares = 0
        For i = 0 To rst.RecordCount - 1
            If rst("id_moneda") = "00002" Then ' DOLARES
               in_saldo = rst("monto_pagar") / rst("tc")
            Else
              in_saldo = rst("monto_pagar")
            End If
            Fila = rst("id_compra") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & rst("comprobante") & vbTab & rst("nproveedor") & vbTab & rst("moneda") & vbTab & Format(in_saldo, "#,##0.00") & vbTab & Format(rst("monto_pagar"), "#,##0.00")
            Grilla.AddItem Fila
            tTotal = tTotal + rst("monto_pagar")
            If rst("id_moneda") = "00002" Then ' DOLARES
               tTotaldolares = tTotaldolares + rst("monto_pagar") / rst("tc")
            Else
               tTotaldolares = tTotaldolares + rst("monto_pagar")
            End If
            
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(tTotaldolares, "###0.00") & vbTab & Format(tTotal, "###0.00")
      Grilla.AddItem Fila
      For k = 5 To 6
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H80FF&
      Next k
    
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub




Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_NEW
    Case KEY_PRINT
          strCadena = "SELECT * FROM DocumentoCompra WHERE idCompra='" & Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 7)) & "'"
          Call ConfiguraRst(strCadena)
          cod_doc = rst("doc_cod")
          serie_doc = rst("sSerie")
          numero_doc = rst("cDocumentoCompra")
          FrmCompras.DtcAlmacen.Enabled = True
          FrmCompras.DtcTipoDoc.Enabled = True
          FrmCompras.txtserie.Enabled = True
          FrmCompras.TxtNumeroDoc.Enabled = True
          FrmCompras.DtcTipoDoc.BoundText = cod_doc
          FrmCompras.Txtdoc_cod.Text = cod_doc
          FrmCompras.txtserie.Text = serie_doc
          FrmCompras.TxtNumeroDoc.Text = numero_doc
          Procedencia = buscar
          Call FrmCompras.buscar_comprobante(Val(Me.HfgFacturas.TextMatrix(Me.HfgFacturas.Row, 7)))
          FrmCompras.Top = 50
          Exit Sub
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_DELETE
          
    Case KEY_PRINT
         ' If Val(Me.Hfpagos.TextMatrix(Me.Hfpagos.Row, 4)) > 0 Then
        'strCadena = "SELECT     Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, movimiento_caja.cPersona, " & _
        "movimiento_caja.descripcion_per, Persona.sDireccionCliente1, Persona.Per_Ruc, movimiento_caja.fecha_valor," & _
        "movimiento_caja.cambio , movimiento_caja.glosa, centro_costos.descripcion, movimiento_caja.Monto,movimiento_caja.monto_letras " & _
        "FROM movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
        "centro_costos ON movimiento_caja.id_costo = centro_costos.id_costo INNER JOIN " & _
        "Persona ON movimiento_caja.cPersona = Persona.cPersona WHERE movimiento_caja.codigo='" & Val(Me.Hfpagos.TextMatrix(Me.Hfpagos.Row, 4)) & "'"
        'Call ConfiguraRst(strCadena)
        'Ans = ShowMultiReport(rst, "RptReciboIngreso", , App.Path + "\Reportes\")
        'Set rst = Nothing
         ' End If
    Case KEY_EXIT
        Unload Me
     
  End Select
  Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub


Private Sub Image1_Click()
Me.frm_canje_anticipos.Visible = False
End Sub

Private Sub img_cerrar_Click()
Me.frmcanjeLetras.Visible = False
End Sub

Private Sub txtBusqueda_proveedor_Change()
    
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE nombre_completo LIKE '%" & Trim(Me.txtBusqueda_proveedor.Text) & "%' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcProveedorLetra)
    
End Sub

Private Sub TxtcodProveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,reversion,id_doc FROM view_cuentas_cobrar WHERE nproveedor LIKE '%" & Trim(Me.TxtcodProveedor.Text) & "%' and ruc='" & KEY_RUC & "'"
    Call llenarGrid(Me.HfgFacturas)
End If
End Sub

Private Sub txtcuenta_redondeo_anticipo_KeyPress(KeyAscii As Integer)
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

Private Sub txtMontoLetra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdupdate.SetFocus
    
End If
End Sub

Private Sub txtnumero_factura_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Me.txtnumero_factura.Text = Format(Me.txtnumero_factura.Text, "00000000")
    strCadena = "SELECT id_proveedor as Codigo,nproveedor as Descripcion FROM view_cuentas_cobrar WHERE id_doc='" & Me.DtcComprobante.BoundText & "' and serie='" & Trim(Me.txtserie_factura.Text) & "' and numero='" & Trim(Me.txtnumero_factura.Text) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcProveedor)
    
    If Me.DtcProveedor.Enabled = True Then
       Me.DtcProveedor.SetFocus
    End If
    
    strCadena = "SELECT id_proveedor as Codigo,nproveedor as Descripcion FROM view_cuentas_cobrar WHERE id_doc='" & Me.DtcComprobante.BoundText & "' and serie='" & Trim(Me.txtserie_factura.Text) & "' and numero='" & Trim(Me.txtnumero_factura.Text) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcProveedorLetra)
End If
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    
    strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,reversion,id_doc FROM view_cuentas_cobrar WHERE comprobante LIKE '%" & Trim(Me.txtNumero.Text) & "%' and ruc='" & KEY_RUC & "'"
    Call llenarGrid(Me.HfgFacturas)
    
End If
End Sub

Private Sub TxtNumeroLetra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.TxtNumeroLetra.Text = Format(Me.TxtNumeroLetra.Text, "00000000")
   Call Resalta(Me.txtMontoLetra)
End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        
    strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,`id_moneda`,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,reversion,id_doc FROM view_cuentas_cobrar WHERE id_proveedor LIKE '%" & Trim(Me.txtRuc.Text) & "%' and ruc='" & KEY_RUC & "'"
    Call llenarGrid(Me.HfgFacturas)

End If
End Sub

Private Sub txtserie_factura_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call Resalta(Me.txtnumero_factura)
   Exit Sub
End If
End Sub

Private Sub txtserie_letra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtNumeroLetra)
End If
End Sub


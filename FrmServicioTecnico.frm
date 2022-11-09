VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmServicioTecnico 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   17535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DETALLE DE SERVICIO TECNICO"
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
      Height          =   8055
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   17295
      Begin VB.CheckBox chk_vinculado 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "VINCULADO"
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
         Height          =   315
         Left            =   3960
         TabIndex        =   62
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text3 
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
         Left            =   2160
         TabIndex        =   57
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox TXTDNIRUC 
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
         Left            =   2190
         TabIndex        =   54
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox TxtSerie 
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
         Left            =   2160
         TabIndex        =   22
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtProducto 
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
         Left            =   2160
         TabIndex        =   21
         Top             =   4920
         Width           =   6975
      End
      Begin VB.TextBox txtContador 
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
         Left            =   2160
         TabIndex        =   20
         Top             =   5880
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "REGISTRO DE CAMBIO DE ACEITE"
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
         Height          =   7455
         Left            =   9480
         TabIndex        =   5
         Top             =   240
         Width           =   7695
         Begin VB.TextBox txtBuscarTecnico 
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
            Left            =   6240
            TabIndex        =   61
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox txtobservacion 
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
            Height          =   555
            Left            =   2160
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   1200
            Width           =   4935
         End
         Begin VB.TextBox TxtContadorServicio 
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
            Left            =   2160
            TabIndex        =   7
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtContadorProximo 
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
            Left            =   2160
            TabIndex        =   6
            Top             =   2880
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker DtpFechaServicio 
            Height          =   345
            Left            =   2160
            TabIndex        =   9
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   42926081
            CurrentDate     =   43349
         End
         Begin MSDataListLib.DataCombo DtcTecnico 
            Height          =   315
            Left            =   2160
            TabIndex        =   10
            Top             =   1920
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   4194304
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
         Begin VitekeySoft.ChameleonBtn cmdsalir_detalle 
            Height          =   615
            Left            =   5160
            TabIndex        =   11
            Top             =   6600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
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
            MICON           =   "FrmServicioTecnico.frx":0000
            PICN            =   "FrmServicioTecnico.frx":001C
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
            Height          =   615
            Left            =   2040
            TabIndex        =   12
            Top             =   6600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1085
            BTYPE           =   5
            TX              =   "SAVE"
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
            MICON           =   "FrmServicioTecnico.frx":3043
            PICN            =   "FrmServicioTecnico.frx":305F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker dtpProximoServicio 
            Height          =   345
            Left            =   2160
            TabIndex        =   13
            Top             =   2400
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   42926081
            CurrentDate     =   43349
         End
         Begin MSDataListLib.DataCombo DtcTipoMantenimiento 
            Height          =   315
            Left            =   2160
            TabIndex        =   44
            Top             =   3480
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   4194304
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
         Begin MSDataListLib.DataCombo DtcEstado 
            Height          =   315
            Left            =   2160
            TabIndex        =   46
            Top             =   4080
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   4194304
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfImpresoras 
            Height          =   1695
            Left            =   240
            TabIndex        =   48
            Top             =   4800
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   2990
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
         Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
            Height          =   735
            Left            =   6480
            TabIndex        =   50
            Top             =   4800
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1296
            BTYPE           =   5
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
            MICON           =   "FrmServicioTecnico.frx":66A7
            PICN            =   "FrmServicioTecnico.frx":66C3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn ChameleonBtn2 
            Height          =   735
            Left            =   6480
            TabIndex        =   51
            Top             =   5760
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1296
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
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmServicioTecnico.frx":9459
            PICN            =   "FrmServicioTecnico.frx":9475
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdImprimir 
            Height          =   615
            Left            =   3600
            TabIndex        =   63
            Top             =   6600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1085
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
            MICON           =   "FrmServicioTecnico.frx":B8BF
            PICN            =   "FrmServicioTecnico.frx":B8DB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REPUESTOS UTILIZADOS"
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
            Left            =   2040
            TabIndex        =   49
            Top             =   4560
            Width           =   1830
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ESTADO :"
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
            TabIndex        =   47
            Top             =   4080
            Width           =   600
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TIPO MANTENIMIENTO:"
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
            TabIndex        =   45
            Top             =   3480
            Width           =   1560
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA SERVICIO :"
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
            Left            =   765
            TabIndex        =   19
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label13 
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
            Left            =   885
            TabIndex        =   18
            Top             =   1320
            Width           =   1035
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TECNICO :"
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
            Left            =   1260
            TabIndex        =   17
            Top             =   2040
            Width           =   660
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PROXIMO SERVICIO:"
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
            TabIndex        =   16
            Top             =   2400
            Width           =   1350
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CONTADOR :"
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
            Left            =   1080
            TabIndex        =   15
            Top             =   840
            Width           =   840
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CONTADOR PROX SERVICIO:"
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
            Left            =   60
            TabIndex        =   14
            Top             =   3000
            Width           =   1860
         End
      End
      Begin VitekeySoft.ChameleonBtn cmdConsultar 
         Height          =   735
         Left            =   7080
         TabIndex        =   23
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1296
         BTYPE           =   5
         TX              =   "CONSULTAR"
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
         MICON           =   "FrmServicioTecnico.frx":DEAC
         PICN            =   "FrmServicioTecnico.frx":DEC8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker Dtpfechaventa 
         Height          =   345
         Left            =   2160
         TabIndex        =   39
         Top             =   5400
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   42926081
         CurrentDate     =   43349
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfMaquinas 
         Height          =   2175
         Left            =   240
         TabIndex        =   56
         Top             =   2640
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3836
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
      Begin MSDataListLib.DataCombo DtcTipoSalida 
         Height          =   315
         Left            =   2160
         TabIndex        =   59
         Top             =   6960
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   4194304
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
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO SALIDA :"
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
         Left            =   1080
         TabIndex        =   60
         Top             =   7080
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CELULAR :"
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
         Left            =   1320
         TabIndex        =   58
         Top             =   2280
         Width           =   690
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI-RUC :"
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
         Left            =   1200
         TabIndex        =   55
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label16 
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
         Left            =   855
         TabIndex        =   41
         Top             =   6480
         Width           =   1155
      End
      Begin VB.Label lblComprobante 
         BackColor       =   &H00808080&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   480
         Left            =   2160
         TabIndex        =   40
         Top             =   6360
         Width           =   5145
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA DE VENTA :"
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
         Left            =   780
         TabIndex        =   38
         Top             =   5400
         Width           =   1230
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   7335
         Left            =   120
         Top             =   360
         Width           =   9255
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NRO SERIE :"
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
         Left            =   1170
         TabIndex        =   28
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCION :"
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
         Left            =   1020
         TabIndex        =   27
         Top             =   4920
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTADOR :"
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
         Left            =   1155
         TabIndex        =   26
         Top             =   5880
         Width           =   855
      End
      Begin VB.Label lblpropietario 
         BackColor       =   &H00808080&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   2160
         TabIndex        =   25
         Top             =   1560
         Width           =   6825
      End
      Begin VB.Label lblIdServicio 
         Caption         =   "IDSERVICIO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.TextBox txtRazonSocial 
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
      Left            =   8400
      TabIndex        =   53
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtServicio 
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
      Left            =   1320
      TabIndex        =   43
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtSerieBusqueda 
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
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtClienteBusqueda 
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
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VitekeySoft.ChameleonBtn mdBuscar 
      Height          =   345
      Left            =   14520
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "BUSCAR"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmServicioTecnico.frx":E1E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpfechaInicio 
      Height          =   300
      Left            =   11400
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
      Format          =   42926081
      CurrentDate     =   43495
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   16200
      TabIndex        =   29
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "FrmServicioTecnico.frx":E1FE
      PICN            =   "FrmServicioTecnico.frx":E21A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfmensualidad 
      Height          =   7575
      Left            =   240
      TabIndex        =   30
      Top             =   600
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   13361
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   12582912
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
   Begin VitekeySoft.ChameleonBtn cmdeliminar 
      Height          =   855
      Left            =   16200
      TabIndex        =   31
      Top             =   2565
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
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
      MICON           =   "FrmServicioTecnico.frx":E66C
      PICN            =   "FrmServicioTecnico.frx":E688
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
      Left            =   16200
      TabIndex        =   32
      Top             =   3555
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmServicioTecnico.frx":10AD2
      PICN            =   "FrmServicioTecnico.frx":10AEE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdUpdate 
      Height          =   855
      Left            =   16200
      TabIndex        =   33
      Top             =   1575
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MODIFICAR"
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
      MICON           =   "FrmServicioTecnico.frx":10EDE
      PICN            =   "FrmServicioTecnico.frx":10EFA
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
      Height          =   300
      Left            =   12840
      TabIndex        =   34
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
      Format          =   42926081
      CurrentDate     =   43495
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAZON SOCIAL :"
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
      Left            =   7320
      TabIndex        =   52
      Top             =   120
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N° SERVICIO:"
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
      TabIndex        =   42
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SERIE :"
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
      Left            =   2640
      TabIndex        =   37
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHAS:"
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
      Left            =   10800
      TabIndex        =   36
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI/RUC :"
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
      Left            =   4920
      TabIndex        =   35
      Top             =   120
      Width           =   675
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8280
      Left            =   0
      Top             =   0
      Width           =   17535
   End
End
Attribute VB_Name = "FrmServicioTecnico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub consultar_serie(ByVal in_serie As String)

strCadena = "call ADM_servicio_tecnico('1','" & Trim(Me.TXTDNIRUC.Text) & "','','','','','','','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.Txtproducto.Text = rst("detalle")
   Me.txtContador.Text = rst("contador")
   Me.Dtpfechaventa.Value = rst("fecha_emision")
   
   Me.lblpropietario.Caption = rst("ncliente")
End If


End Sub


Private Sub chk_vinculado_Click()

If Me.chk_vinculado.Value = 1 Then
   Me.TXTDNIRUC.Locked = False
Else
    Me.TXTDNIRUC.Locked = True
End If


End Sub

Private Sub cmdConsultar_Click()


If Trim(Me.TXTDNIRUC.Text) <> "" Then
    Call llenar_maquinas(Me.HfMaquinas, Trim(Me.TXTDNIRUC.Text))
    
    
End If

End Sub

Private Sub cmdEliminar_Click()

If MsgBox("Desea Anular el Servicio Tecnico", vbYesNo + vbQuestion) = vbYes Then
    
    strCadena = "call PRO_servicio_tecnico('8','" & Val(Me.hfmensualidad.TextMatrix(Me.hfmensualidad.Row, 0)) & "','" & KEY_FECHA & "','0','" & in_dni & "','','','','','','','','','no','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
    
    
     strCadena = "call PRO_servicio_tecnico('2','','" & KEY_FECHA & "','0','','','" & Trim(Me.txtserie.Text) & "','','','','','','','no','" & KEY_USUARIO & "','" & KEY_RUC & "')"
     Call llenarGrid(Me.hfmensualidad)
    
End If

End Sub

Private Sub cmdNuevo_Click()




strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcTecnico)
Me.frmdetalle.Visible = True
Me.txtserie.Text = ""
Me.Txtproducto.Text = ""
Me.txtContador.Text = ""
Me.Dtpfechaventa.Value = KEY_FECHA

Me.lblpropietario.Caption = ""

Me.DtpFechaServicio.Value = KEY_FECHA
Me.dtpProximoServicio.Value = KEY_FECHA

Me.TxtContadorServicio.Text = ""
Me.txtContadorProximo.Text = ""
Me.Txtproducto.Text = ""
Me.Txtproducto.Tag = ""
Me.Dtpfechaventa.Value = KEY_FECHA
Me.txtContador.Text = ""
Me.txtserie.Text = ""
Me.lblComprobante.Caption = ""
Me.lblComprobante.Tag = ""
Me.txtObservacion.Text = ""
Me.txtContadorProximo.Text = ""
Me.TxtContadorServicio.Text = ""
Me.TXTDNIRUC.Text = ""
Me.TXTDNIRUC.Locked = False
Me.chk_vinculado.Value = 0


End Sub

Private Sub llenar_maquinas(ByVal Grilla As MSHFlexGrid, ByVal in_dni As String)

On Error GoTo salir


strCadena = "call PRO_servicio_tecnico('6','" & Val(Me.lblIdServicio.Caption) & "','" & KEY_FECHA & "','0','" & in_dni & "','','" & Trim(Me.txtserie.Text) & "','','','','','','','','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
  
       Me.TXTDNIRUC.Locked = True
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 900
           Grilla.ColWidth(2) = 1800
           Grilla.ColWidth(3) = 3500
           Grilla.ColWidth(4) = 1500
           Grilla.ColWidth(5) = 900
           
           
       Next
         cabecera = "IDVENTA" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "MAQUINA" & vbTab & "SERIE" & vbTab & "CONTADOR"
         Grilla.AddItem cabecera
         For k = 0 To 5
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = rst("id_venta") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & rst("documento") & vbTab & rst("detalle") & vbTab & rst("nro_chasis") & vbTab & rst("contador")
             Grilla.AddItem Fila
             
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Call ConfiguraRst(strCadena)

 
If rst.RecordCount < 1 Then
    Grilla.Rows = 0

    Exit Sub

End If
  
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 600
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 3000
           Grilla.ColWidth(4) = 2000
           Grilla.ColWidth(5) = 900
           Grilla.ColWidth(6) = 2000
           Grilla.ColWidth(7) = 1500
           Grilla.ColWidth(8) = 1200
           Grilla.ColWidth(9) = 2500
           Grilla.ColWidth(10) = 2500
           
       Next
         cabecera = "N° SERV" & vbTab & "FECHA" & vbTab & "DNI/RUC" & vbTab & "CLIENTE" & vbTab & "COMPROBANTE" & vbTab & "CODIGO" & vbTab & "PRODUCTO" & vbTab & "SERIE" & vbTab & "CONTADOR" & vbTab & "TECNICO" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 10
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = Format(rst("id"), "0000") & vbTab & Format(rst("fecha_solicitud"), "dd-mm-YYYY") & vbTab & rst("dni_cliente") & vbTab & rst("ncliente") & vbTab & rst("documento") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("serie") & vbTab & rst("contador") & vbTab & rst("tecnico") & vbTab & rst("estado")
             Grilla.AddItem Fila
             
             
            If rst("id_estado") = 4 Then
            For k = 8 To 10
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &H8080FF
            Next k
            End If
                            
             
            rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub cmdProcesar_Click()
If Trim(Me.TXTDNIRUC.Text) <> "" And Trim(Me.Txtproducto.Tag) <> "" And Trim(Me.lblComprobante.Tag) <> "" Then
    
    If Me.chk_vinculado.Value = 1 Then
        in_vinculado = "si"
    Else
        in_vinculado = "no"
    End If
    
    strCadena = "call PRO_servicio_tecnico('1','0','" & Format(Me.DtpFechaServicio.Value, "YYYY-mm-dd") & "','" & Me.lblComprobante.Tag & "','" & Trim(Me.TXTDNIRUC.Text) & "','" & Trim(Me.Txtproducto.Tag) & "','" & Trim(Me.txtserie.Text) & "','" & Me.TxtContadorServicio.Text & "','" & Me.DtcEstado.BoundText & "','" & DtcTipoSalida.BoundText & "','" & Me.DtcTipoMantenimiento.BoundText & "','" & Trim(Me.txtObservacion.Text) & "','" & Me.DtcTecnico.BoundText & "','" & in_vinculado & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    
    
    MsgBox "Servicio Tecnico Procesado", vbInformation
    Me.frmdetalle.Visible = False
    
 
    
    
    strCadena = "call PRO_servicio_tecnico('2','','" & KEY_FECHA & "','0','','','','','','','','','','','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    Call llenarGrid(Me.hfmensualidad)
    
    
    
    
    Exit Sub
    
    
    
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdsalir_detalle_Click()

Me.frmdetalle.Visible = False

End Sub

Private Sub get_servicio(ByVal in_servicio As String)

strCadena = "call PRO_servicio_tecnico('9','" & in_servicio & "','" & KEY_FECHA & "','0','','','','','','','','','','','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and  ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcTecnico)
    
    Me.DtcTecnico.BoundText = rst("id_tecnico")
    Me.DtcTipoMantenimiento.BoundText = rst("id_tipo")
    Me.DtcTipoSalida.BoundText = rst("id_tipo_salida")
    
    
   Me.lblIdServicio.Caption = in_servicio
   Me.txtserie.Text = rst("serie")
   Me.TXTDNIRUC.Text = rst("dni_cliente")
   Me.lblpropietario.Caption = get_persona(rst("dni_cliente"))
   Me.Txtproducto.Text = get_producto(rst("id_producto"))
   Me.Txtproducto.Tag = rst("id_producto")
   Me.lblComprobante.Tag = rst("id_venta")
   Me.lblComprobante.Caption = rst("documento")
   Me.txtObservacion.Text = rst("detalle")
  
   Me.TxtContadorServicio.Text = rst("contador")
   
   
   
   
   Me.frmdetalle.Visible = True
   
End If
End Sub


Private Sub cmdupdate_Click()
    If Val(Me.hfmensualidad.TextMatrix(Me.hfmensualidad.Row, 0)) > 0 Then
        Call get_servicio(Val(Me.hfmensualidad.TextMatrix(Me.hfmensualidad.Row, 0)))
    End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 150




strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM servicio_tecnico_tipo_mantenimiento "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoMantenimiento)


strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM adm_servicio_tecnico_estado "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstado)


strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM servicio_tecnico_tipo_salida "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoSalida)


strCadena = "call PRO_servicio_tecnico('2','','','0','','','','','','','','','','','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call llenarGrid(Me.hfmensualidad)
    

End Sub
Private Sub get_comprobante(ByVal in_venta As String)


strCadena = "call PRO_servicio_tecnico('7','" & Val(Me.lblIdServicio.Caption) & "','" & KEY_FECHA & "','" & in_venta & "','0','','0','','','','','','','no','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.Txtproducto.Text = rst("detalle")
   Me.Txtproducto.Tag = rst("id_producto")
   Me.Dtpfechaventa.Value = rst("fecha_emision")
   Me.txtContador.Text = rst("contador")
   Me.txtserie.Text = rst("nro_chasis")
   Me.lblComprobante.Caption = rst("documento")
   Me.lblComprobante.Tag = rst("id_venta")
Else
    Me.Txtproducto.Text = ""
   Me.Dtpfechaventa.Value = KEY_FECHA
   Me.txtContador.Text = ""
   Me.txtserie.Text = ""
   Me.lblComprobante.Caption = ""
   
End If

End Sub
Private Sub HfMaquinas_SelChange()
If Val(Me.HfMaquinas.TextMatrix(Me.HfMaquinas.Row, 0)) > 0 Then
    Call get_comprobante(Val(Me.HfMaquinas.TextMatrix(Me.HfMaquinas.Row, 0)))
End If
End Sub

Private Sub txtBuscarTecnico_Change()
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE nombre_completo LIKE '%" & Trim(Me.txtBuscarTecnico.Text) & "%' and  id_personal='si' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcTecnico)
End Sub

Private Sub TXTDNIRUC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Procedencia = Selecionar
   FrmPersona.Show
   Exit Sub
End If
End Sub

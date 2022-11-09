VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmLibroDetalle 
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
   Begin VB.Frame frmIngreso 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NUMEROS DE INGRESO"
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
      Height          =   4095
      Left            =   10080
      TabIndex        =   69
      Top             =   2040
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame frmDetalleIngreso 
         BackColor       =   &H00FFFFFF&
         Height          =   3375
         Left            =   240
         TabIndex        =   74
         Top             =   480
         Visible         =   0   'False
         Width           =   5415
         Begin VB.TextBox txtNumero_ingreso 
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
            Left            =   2400
            MaxLength       =   500
            TabIndex        =   75
            Top             =   720
            Width           =   1815
         End
         Begin VitekeySoft.ChameleonBtn cmdprocesarIngreso 
            Height          =   795
            Left            =   2400
            TabIndex        =   77
            Top             =   1320
            Width           =   885
            _ExtentX        =   1561
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
            MICON           =   "frmLibroDetalle.frx":0000
            PICN            =   "frmLibroDetalle.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdCerrarIngreso 
            Height          =   795
            Left            =   3360
            TabIndex        =   78
            Top             =   1320
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   1402
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
            MICON           =   "frmLibroDetalle.frx":3664
            PICN            =   "frmLibroDetalle.frx":3680
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblid_ingreso 
            Height          =   255
            Left            =   2400
            TabIndex        =   90
            Top             =   2280
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° INGRESO :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   165
            Left            =   1185
            TabIndex        =   76
            Top             =   780
            Width           =   975
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfIngreso 
         Height          =   3375
         Left            =   240
         TabIndex        =   70
         Top             =   480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5953
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
      Begin VitekeySoft.ChameleonBtn cmdnuevoIngreso 
         Height          =   795
         Left            =   4680
         TabIndex        =   71
         Top             =   480
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1402
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
         MICON           =   "frmLibroDetalle.frx":66A7
         PICN            =   "frmLibroDetalle.frx":66C3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdModificarIngreso 
         Height          =   795
         Left            =   4680
         TabIndex        =   72
         Top             =   1320
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1402
         BTYPE           =   5
         TX              =   "EDITAR"
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
         MICON           =   "frmLibroDetalle.frx":6B15
         PICN            =   "frmLibroDetalle.frx":6B31
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdEliminarIngreso 
         Height          =   795
         Left            =   4680
         TabIndex        =   73
         Top             =   2160
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1402
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
         MICON           =   "frmLibroDetalle.frx":9E07
         PICN            =   "frmLibroDetalle.frx":9E23
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdcerraringresodetalle 
         Height          =   795
         Left            =   4680
         TabIndex        =   79
         Top             =   3000
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1402
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
         MICON           =   "frmLibroDetalle.frx":C26D
         PICN            =   "frmLibroDetalle.frx":C289
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame frmcompatibilidad 
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
      Height          =   3015
      Left            =   10800
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox txtCodCompatible 
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
         MaxLength       =   80
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TxtCompatible 
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
         MaxLength       =   80
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VitekeySoft.ChameleonBtn CmdprocesarCompatibilidad 
         Height          =   400
         Left            =   720
         TabIndex        =   85
         Top             =   2160
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "PROCESAR COMPATIBILIDAD"
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
         MICON           =   "frmLibroDetalle.frx":F2B0
         PICN            =   "frmLibroDetalle.frx":F2CC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO :"
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
         TabIndex        =   7
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
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
         Left            =   225
         TabIndex        =   6
         Top             =   840
         Width           =   1005
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdAgregarCompatibilidad 
      Height          =   375
      Left            =   10800
      TabIndex        =   83
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "     AGREGAR  "
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
      MICON           =   "frmLibroDetalle.frx":118B1
      PICN            =   "frmLibroDetalle.frx":118CD
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
      Height          =   855
      Left            =   17160
      TabIndex        =   80
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "frmLibroDetalle.frx":13EB2
      PICN            =   "frmLibroDetalle.frx":13ECE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdIngresos 
      Height          =   315
      Left            =   9000
      TabIndex        =   68
      Top             =   5760
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "N° INGRESO"
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
      MICON           =   "frmLibroDetalle.frx":17516
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtContenido 
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
      Height          =   2655
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   66
      Top             =   6360
      Width           =   7575
   End
   Begin VB.TextBox txtCantidad 
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
      MaxLength       =   500
      TabIndex        =   64
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox txtAnio 
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
      MaxLength       =   500
      TabIndex        =   63
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox txtPrecio 
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
      MaxLength       =   500
      TabIndex        =   62
      Top             =   4850
      Width           =   1935
   End
   Begin VB.TextBox TxtTamalo 
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
      MaxLength       =   500
      TabIndex        =   61
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox txtEdicion 
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
      Left            =   1440
      MaxLength       =   500
      TabIndex        =   60
      Top             =   5760
      Width           =   2415
   End
   Begin VB.TextBox txtTomo 
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
      Left            =   1440
      MaxLength       =   500
      TabIndex        =   59
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox txtPaginas 
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
      Left            =   1440
      MaxLength       =   500
      TabIndex        =   58
      Top             =   4850
      Width           =   2415
   End
   Begin VB.TextBox txtciudad 
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
      Left            =   1440
      MaxLength       =   500
      TabIndex        =   50
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox TxtBuscarProveedor 
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
      Left            =   5745
      MaxLength       =   80
      TabIndex        =   36
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox TxtBuscamarca 
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
      Left            =   5760
      MaxLength       =   80
      TabIndex        =   35
      Top             =   3165
      Width           =   735
   End
   Begin VB.TextBox TxtBuscaLinea 
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
      Left            =   5760
      MaxLength       =   80
      TabIndex        =   34
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox TxtDescripcion 
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
      Left            =   1515
      MaxLength       =   500
      TabIndex        =   33
      Top             =   1680
      Width           =   8415
   End
   Begin VB.CommandButton cmdCodBarra 
      Caption         =   "GENERAR BARRA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6825
      TabIndex        =   28
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton CmdQuitar 
      Caption         =   "-"
      Height          =   255
      Left            =   3825
      TabIndex        =   27
      Top             =   915
      Width           =   375
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "+"
      Height          =   255
      Left            =   3465
      TabIndex        =   26
      Top             =   915
      Width           =   375
   End
   Begin VB.TextBox TxtCodBarra 
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
      Height          =   315
      Left            =   1545
      MaxLength       =   80
      TabIndex        =   25
      Top             =   915
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MATRIZ UBICACION FISICA"
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
      Height          =   2295
      Left            =   15120
      TabIndex        =   8
      Top             =   5760
      Width           =   4815
      Begin VB.TextBox TxtSector 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtPiso 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox TxtAndamio 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_x 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   10
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Txt_y 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   9
         Top             =   1800
         Width           =   975
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   330
         Left            =   1200
         TabIndex        =   14
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632319
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
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BIBLIOTECA"
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
         TabIndex        =   19
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SECTOR :"
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
         Left            =   450
         TabIndex        =   18
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PISO :"
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
         TabIndex        =   17
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ANDAMIO :"
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
         TabIndex        =   16
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CASILLERO :"
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
         TabIndex        =   15
         Top             =   1800
         Width           =   825
      End
   End
   Begin VB.Frame FrmCaracteristicas 
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
      Height          =   1575
      Left            =   10680
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtCaracteristica 
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
         Height          =   765
         Left            =   240
         MaxLength       =   80
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   3735
      End
      Begin VitekeySoft.ChameleonBtn ChameleonBtn3 
         Height          =   380
         Left            =   240
         TabIndex        =   88
         Top             =   1100
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "PROCESAR COMPATIBILIDAD"
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
         MICON           =   "frmLibroDetalle.frx":17532
         PICN            =   "frmLibroDetalle.frx":1754E
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
   Begin VB.PictureBox Picthumbnail1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   15120
      ScaleHeight     =   4905
      ScaleMode       =   0  'User
      ScaleWidth      =   4785
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16080
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCompatibilidad 
      Height          =   3015
      Left            =   10800
      TabIndex        =   20
      Top             =   360
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5318
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCaracteristica 
      Height          =   1455
      Left            =   10680
      TabIndex        =   21
      Top             =   5760
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2566
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgBarras 
      Height          =   855
      Left            =   6825
      TabIndex        =   29
      Top             =   180
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1508
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
   Begin MSDataListLib.DataCombo DtcArea 
      Height          =   330
      Left            =   1485
      TabIndex        =   37
      Top             =   2160
      Width           =   4215
      _ExtentX        =   7435
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
   Begin MSDataListLib.DataCombo DtcEditorial 
      Height          =   330
      Left            =   1485
      TabIndex        =   38
      Top             =   3165
      Width           =   4215
      _ExtentX        =   7435
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
   Begin MSDataListLib.DataCombo DtcProcedencia 
      Height          =   330
      Left            =   7005
      TabIndex        =   39
      Top             =   3735
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
   Begin MSDataListLib.DataCombo DTcTipoLibro 
      Height          =   330
      Left            =   1485
      TabIndex        =   40
      Top             =   3735
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
   Begin MSDataListLib.DataCombo DtcAutor 
      Height          =   330
      Left            =   1485
      TabIndex        =   46
      Top             =   2640
      Width           =   4215
      _ExtentX        =   7435
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
   Begin MSDataListLib.DataCombo DtcUnidad 
      Height          =   330
      Left            =   7680
      TabIndex        =   49
      Top             =   2160
      Width           =   2295
      _ExtentX        =   4048
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
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   855
      Left            =   18600
      TabIndex        =   81
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "frmLibroDetalle.frx":19B33
      PICN            =   "frmLibroDetalle.frx":19B4F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdeliminar 
      Height          =   375
      Left            =   13200
      TabIndex        =   84
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "     ELIMINAR  "
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
      MICON           =   "frmLibroDetalle.frx":19F3F
      PICN            =   "frmLibroDetalle.frx":19F5B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn CmdAgregarCaracteristica 
      Height          =   375
      Left            =   12360
      TabIndex        =   86
      Top             =   5220
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "AGREGAR  "
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
      MICON           =   "frmLibroDetalle.frx":1CE0F
      PICN            =   "frmLibroDetalle.frx":1CE2B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdeliminarcaracteritica 
      Height          =   375
      Left            =   13680
      TabIndex        =   87
      Top             =   5220
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "ELIMINAR  "
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
      MICON           =   "frmLibroDetalle.frx":1F410
      PICN            =   "frmLibroDetalle.frx":1F42C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn CmdFoto 
      Height          =   375
      Left            =   15120
      TabIndex        =   89
      Top             =   5160
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "             SUBIR CATALOGO"
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
      MICON           =   "frmLibroDetalle.frx":222E0
      PICN            =   "frmLibroDetalle.frx":222FC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COINSIDENCIAS:"
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
      Left            =   10785
      TabIndex        =   82
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label LblObservacion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTENIDO :"
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
      TabIndex        =   67
      Top             =   6360
      Width           =   915
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2940
      Left            =   120
      Top             =   6240
      Width           =   9975
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   5895
      TabIndex        =   65
      Top             =   5880
      Width           =   885
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AÑO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   6345
      TabIndex        =   57
      Top             =   5400
      Width           =   435
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   6105
      TabIndex        =   56
      Top             =   4920
      Width           =   675
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAMAÑO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   6015
      TabIndex        =   55
      Top             =   4440
      Width           =   765
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EDICION :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   450
      TabIndex        =   54
      Top             =   5760
      Width           =   765
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOMO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   660
      TabIndex        =   53
      Top             =   5400
      Width           =   555
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAGINAS :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   450
      TabIndex        =   52
      Top             =   4920
      Width           =   765
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CIUDAD :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   510
      TabIndex        =   51
      Top             =   4500
      Width           =   705
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1860
      Left            =   120
      Top             =   4320
      Width           =   9975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNIDAD:"
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
      Left            =   6960
      TabIndex        =   48
      Top             =   2250
      Width           =   675
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AUTOR :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   825
      TabIndex        =   47
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO LIBRO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   495
      TabIndex        =   45
      Top             =   3765
      Width           =   945
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROCEDENCIA :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   5535
      TabIndex        =   44
      Top             =   3765
      Width           =   1185
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EDITORIAL :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   555
      TabIndex        =   43
      Top             =   3225
      Width           =   885
   End
   Begin VB.Label LblLaboratorio 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AREA :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   945
      TabIndex        =   42
      Top             =   2220
      Width           =   495
   End
   Begin VB.Label LblDescripcion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TITULO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   795
      TabIndex        =   41
      Top             =   1740
      Width           =   645
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.CLASIFICACION  :"
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
      TabIndex        =   32
      Top             =   915
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO INTERNO:"
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
      TabIndex        =   31
      Top             =   435
      Width           =   1245
   End
   Begin VB.Label LblCodigoProducto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   1530
      TabIndex        =   30
      Top             =   195
      Width           =   4305
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3540
      Left            =   10560
      Top             =   5640
      Width           =   4335
   End
   Begin VB.Label lblError 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   10560
      TabIndex        =   24
      Top             =   4650
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblCabecera 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CODIGO DE BARRAS YA REGISTRADO"
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
      Height          =   315
      Left            =   10560
      TabIndex        =   23
      Top             =   4200
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3900
      Left            =   10560
      Top             =   75
      Width           =   4335
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARACTERISTICAS :"
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
      Left            =   10665
      TabIndex        =   22
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1260
      Left            =   120
      Top             =   120
      Width           =   9975
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2700
      Left            =   120
      Top             =   1560
      Width           =   9975
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "frmLibroDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrCodTabla As String
Dim StrCodProducto As String
Dim img As String
Dim strAfectaStock As String * 2
Dim StrPercepcion As String * 2
Dim strAfectoIGV As String * 2
Dim RstAlmProd As New ADODB.Recordset
Dim sub_producto As String
Public Procedencia As EnumProcede
Dim strCombo As String
Dim PrtFoto As String
Dim FlagFoto As Boolean






Private Sub ChkDesabilitado_Click()

End Sub








Private Sub ChameleonBtn1_Click()

End Sub

Private Sub ChameleonBtn2_Click()

End Sub

Private Sub CmdAgregar_Click()
Call agrega_barra(Trim(Me.LblCodigoProducto.Caption))
End Sub
Private Sub agrega_barra(ByVal codigo As String)
If Trim(Me.TxtCodBarra.Text) <> "" Then
    strCadena = "SELECT * FROM producto_barras WHERE cod_barra='" & Trim(Me.TxtCodBarra.Text) & "' AND id_producto='" & Trim(codigo) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        MsgBox "Codigo de barras ya registrado", vbInformation, KEY_EMPRESA
    Else
        strCadena = "INSERT INTO producto_barras VALUES('" & Trim(codigo) & "','" & Trim(Me.TxtCodBarra.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
    End If
    Set rst = Nothing
    Call llena_barra
End If
End Sub


Private Sub CmdAgregarCaracteristica_Click()
If Me.FrmCaracteristicas.Visible = False Then
   Me.FrmCaracteristicas.Visible = True
Else
    Me.FrmCaracteristicas.Visible = False
End If
End Sub

Private Sub cmdAgregarCompatibilidad_Click()
If Me.frmcompatibilidad.Visible = False Then
   Me.frmcompatibilidad.Visible = True
Else
    Me.frmcompatibilidad.Visible = False
End If
End Sub

Private Sub cmdCerrarIngreso_Click()
Me.frmDetalleIngreso.Visible = False
End Sub

Private Sub cmdcerraringresodetalle_Click()
Me.frmIngreso.Visible = False
End Sub

Private Sub cmdCodBarra_Click()
Me.TxtCodBarra.Text = Trim(Me.LblCodigoProducto.Caption)
Call agrega_barra(Trim(Me.LblCodigoProducto.Caption))
End Sub



Private Sub cmdeliminar_Click()
If Val(Me.HfCompatibilidad.TextMatrix(Me.HfCompatibilidad.Row, 0)) > 0 Then
    strCadena = "DELETE FROM producto_compatibilidad WHERE id_producto_compatible ='" & Trim(Me.HfCompatibilidad.TextMatrix(Me.HfCompatibilidad.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Call llenarCompatibilidad(Me.HfCompatibilidad, Trim(Me.LblCodigoProducto.Caption))
End If
End Sub

Private Sub cmdeliminarcaracteritica_Click()
If Val(Me.HfCaracteristica.TextMatrix(Me.HfCaracteristica.Row, 0)) > 0 Then
    strCadena = "DELETE FROM producto_caracteristicas WHERE id_detalle ='" & Trim(Me.HfCaracteristica.TextMatrix(Me.HfCaracteristica.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Call llenarCaracteristica(Me.HfCaracteristica, Trim(Me.LblCodigoProducto.Caption))
End If

End Sub

Private Sub cmdEliminarIngreso_Click()

strCadena = "DELETE FROM biblio_libro_ingreso WHERE id='" & Val(Me.HfIngreso.TextMatrix(Me.HfIngreso.Row, 0)) & "'"
CnBd.Execute (strCadena)
Call llenar_ingreso(Me.HfIngreso, Trim(Me.LblCodigoProducto.Caption))

End Sub

Private Sub CmdFoto_Click()
Dim ext As String
On Error GoTo finish
Me.CommonDialog1.Filter = "*.Jpg"
Me.CommonDialog1.ShowOpen
Me.Picthumbnail1.Picture = LoadPicture(Me.CommonDialog1.FileName)
PrtFoto = Trim(Me.CommonDialog1.FileName)

img = Trim(str(Me.LblCodigoProducto.Caption) & Trim(Right(PrtFoto, 4)))
strCadena = "SELECT * FROM producto_foto WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    img = Trim(str(Me.LblCodigoProducto.Caption)) & "_" + str(rstK.RecordCount + 1) & Trim(Right(PrtFoto, 4))
    
End If
Me.CmdFoto.Caption = "Archivos de Imagen" + Space(1) + "[ " + str(rstK.RecordCount) + " ]"
Call Copiar_Archivo(PrtFoto, App.Path + "\archivos\" & KEY_RUC & "\" + img)
    strCadena = "INSERT INTO producto_foto (id_producto,foto,ruc)VALUES('" & Trim(Me.LblCodigoProducto.Caption) & "','" & img & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
FlagFoto = True

Exit Sub
finish: MsgBox "La Imagen que Intenta Subir tiene que ser .JPG", vbInformation, "Imagen no Compatible"
End Sub

Private Sub cmdImpresionBarras_Click()

End Sub









Private Sub cmdNext_Click()

End Sub

Private Sub cmdIngresos_Click()
Me.frmIngreso.Visible = True
Call llenar_ingreso(Me.HfIngreso, Trim(Me.LblCodigoProducto.Caption))
End Sub

Private Sub cmdModificarIngreso_Click()
Me.frmDetalleIngreso.Visible = True
Me.lblid_ingreso.Caption = Val(Me.HfIngreso.TextMatrix(Me.HfIngreso.Row, 0))
Me.txtNumero_ingreso.Text = Trim(Me.HfIngreso.TextMatrix(Me.HfIngreso.Row, 2))

End Sub

Private Sub cmdnuevoIngreso_Click()
Me.frmDetalleIngreso.Visible = True
Me.lblid_ingreso.Caption = 0
Me.txtNumero_ingreso.Text = ""
End Sub

Private Sub cmdprocesar_Click()
      On Error GoTo error
      Call Save
      FlagFoto = False
      Exit Sub
       
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
End Sub

Private Sub CmdprocesarCaracteristica_Click()
    strCadena = "INSERT INTO producto_caracteristicas (id_producto,caracteristica,ruc)VALUES('" & Trim(Me.LblCodigoProducto.Caption) & "','" & Trim(Me.txtCaracteristica.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    Me.FrmCaracteristicas.Visible = False
    Call llenarCaracteristica(Me.HfCaracteristica, Trim(Me.LblCodigoProducto.Caption))
    Exit Sub
End Sub

Private Sub CmdprocesarCompatibilidad_Click()
strCadena = "SELECT * FROM producto_compatibilidad WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND id_producto_compatible='" & Trim(Me.txtCodCompatible.Text) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    strCadena = "INSERT INTO producto_compatibilidad (id_producto,id_producto_compatible,ruc)VALUES('" & Trim(Me.LblCodigoProducto.Caption) & "','" & Trim(Me.txtCodCompatible.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    Me.frmcompatibilidad.Visible = False
    Call llenarCompatibilidad(Me.HfCompatibilidad, Trim(Me.LblCodigoProducto.Caption))
    Exit Sub
End If
End Sub

Private Sub cmdprocesarIngreso_Click()
If Val(Me.txtNumero_ingreso.Text) > 0 Then
    strCadena = "call put_libro_ingreso('" & Val(Me.lblid_ingreso.Caption) & "','" & Trim(Me.LblCodigoProducto.Caption) & "','" & Trim(Me.txtNumero_ingreso.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Call llenar_ingreso(Me.HfIngreso, Trim(Me.LblCodigoProducto.Caption))
    Me.frmDetalleIngreso.Visible = False
Else
    MsgBox "Ingrese un N° Ingreso Valido", vbInformation
End If
End Sub

Private Sub CmdQuitar_Click()
strCadena = "DELETE  FROM producto_barras WHERE cod_barra='" & Trim(Me.HfgBarras.TextMatrix(Me.HfgBarras.Row, 1)) & "' AND id_producto='" & Trim(Me.HfgBarras.TextMatrix(Me.HfgBarras.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
 
Call llena_barra
End Sub

Private Sub CmdRelacionados_Click()

FrmProductoRelacionado.Show
End Sub


Private Sub ingreso_relacionados()

End Sub




Private Sub Command1_Click()

End Sub



Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub DtcAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM almacen_producto WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txt_x.Text = rst("casillero_x")
    Me.Txt_y.Text = rst("casillero_y")
    Me.TxtAndamio.Text = rst("andamio")
    Me.TxtPiso.Text = rst("piso")
    Me.TxtSector.Text = rst("sector")
Else
    Me.txt_x.Text = ""
    Me.Txt_y.Text = ""
    Me.TxtAndamio.Text = ""
    Me.TxtPiso.Text = ""
    Me.TxtSector.Text = ""
End If
End If
End Sub


Private Sub llenarCompatibilidad(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
strCadena = "SELECT C.id_producto_compatible,P.nombre_prod FROM producto_compatibilidad C,producto P WHERE C.id_producto_compatible=P.id_producto AND C.id_producto='" & id_producto & "' AND P.ruc='" & KEY_RUC & "' AND C.ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 3200
        Next
         cabecera = "CODIGO" & vbTab & "NOMBRE PRODUCTO"
         Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id_producto_compatible") & vbTab & UCase(rst("nombre_prod"))
             Grilla.AddItem Fila
             Fila = ""
             rst.MoveNext
        Next i
  Exit Sub
End Sub
Private Sub llenarCaracteristica(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
strCadena = "SELECT * FROM producto_caracteristicas WHERE id_producto='" & id_producto & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 4000
        Next
         cabecera = "CODIGO" & vbTab & "CARACTERISTICAS"
         Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id_detalle") & vbTab & UCase(rst("caracteristica"))
             Grilla.AddItem Fila
             Fila = ""
             rst.MoveNext
        Next i
  Exit Sub
End Sub

Private Sub llenar_ingreso(ByVal Grilla As MSHFlexGrid, ByVal in_libro As String)
strCadena = "SELECT * FROM biblio_libro_ingreso WHERE id_libro='" & in_libro & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Exit Sub
End If
   
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 2000
        Next
         cabecera = "ID" & vbTab & "N° ORDEN" & vbTab & "N° INGRESO"
         Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id") & vbTab & Format(i + 1, "000") & vbTab & rst("ingreso")
             Grilla.AddItem Fila
             rst.MoveNext
        Next i
  Exit Sub
End Sub

Private Sub llenarcriterio(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
strCadena = "SELECT * FROM producto_busqueda WHERE id_producto='" & id_producto & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 4000
        Next
         cabecera = "CODIGO" & vbTab & "CRITERIO BUSQUEDA"
         Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id_criterio") & vbTab & UCase(rst("criterio"))
             Grilla.AddItem Fila
             Fila = ""
             rst.MoveNext
        Next i
  Exit Sub
End Sub









Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
    Exit Sub
  End If
  If KeyCode = 27 Then
    Unload Me
  End If
  
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
strCombo = "no"
'Me.Picthumbnail1.BordeEstilo = Borde4

FlagFoto = False
'---------Llenar  Combos------------------------
  strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcArea)
  
  
  
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  
 
  
  strCadena = "SELECT id_editorial as Codigo, descripcion as Descripcion FROM editorial WHERE ruc='" & KEY_RUC & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcEditorial)
  
  
  strCadena = "SELECT id_procedencia as Codigo, descripcion as Descripcion FROM biblio_procedencia WHERE ruc='" & KEY_RUC & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcProcedencia)
  
  

  strCadena = "SELECT id_und as Codigo, abreviatura as Descripcion FROM unidad WHERE id_usu='" & KEY_RUC & "' " & _
  " ORDER BY abreviatura"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcUnidad)
  
  
  strCadena = "SELECT id_tipoproducto as Codigo,descripcion as Descripcion FROM tipo_producto WHERE ruc='" & KEY_RUC & "' ORDER BY id_tipoproducto"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DTcTipoLibro)
  
  strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' AND id_proveedor='si' "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAutor)
  
  
  

Select Case frmLibro.Procedencia
        Case nuevo
            Me.CmdAgregar.Enabled = False
            Me.CmdQuitar.Enabled = False
         
         
    Case Modificar
        Me.CmdAgregar.Enabled = True
        Me.CmdQuitar.Enabled = True
      Call LLENA
      Call llenarCompatibilidad(Me.HfCompatibilidad, Trim(Me.LblCodigoProducto.Caption))
      Call llenarCaracteristica(Me.HfCaracteristica, Trim(Me.LblCodigoProducto.Caption))
  End Select
'-------------------------------------------
  
End Sub

Private Sub LLENA()
  strCadena = "SELECT * FROM view_libro WHERE id_libro = '" & frmLibro.HfdGrilla.TextMatrix(frmLibro.HfdGrilla.Row, 0) & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' LIMIT 1"
  Call ConfiguraRstK(strCadena)
  If rstK.RecordCount > 0 Then
    StrCodTabla = rstK("id_libro")
    Me.TxtCodBarra.Text = rstK("codigo_libro")
    Me.LblCodigoProducto.Caption = rstK("id_libro")
    Me.txtdescripcion.Text = rstK("titulo")
    Me.DtcUnidad.BoundText = rstK("id_unidad")
    Me.DtcArea.BoundText = rstK("id_area")
    Me.DtcAutor.BoundText = rstK("id_autor")
    Me.DtcEditorial.BoundText = rstK("id_editorial")
    Me.DTcTipoLibro.BoundText = rstK("id_tipo")
    Me.DtcProcedencia.BoundText = rstK("id_procedencia")
   
    Me.txtciudad.Text = rstK("ciudad")
    Me.txtPaginas.Text = rstK("paginas")
    Me.txtTomo.Text = rstK("tomo")
    Me.txtEdicion.Text = rstK("edicion")
    Me.TxtTamalo.Text = rstK("tamanio")
    Me.txtPrecio.Text = rstK("precio")
    Me.txtAnio.Text = rstK("anio")
    Me.txtCantidad.Text = rstK("cantidad")
    Me.TxtContenido.Text = rstK("contenido")
    
    
    
    
    
    
    
    
    
  Me.TxtSector.Text = rstK("sector")
  Me.TxtPiso.Text = rstK("piso")
  Me.TxtAndamio.Text = rstK("andamio")
  Me.txt_x.Text = rstK("casillero_x")
  Me.Txt_y.Text = rstK("casillero_y")
  Call llena_barra
  '--------- foto--------
    
    If IsNull(rstK("imagen")) = False And Len(rstK("imagen")) > 5 Then
    If VerificarArchivo(App.Path & "\archivos\" & KEY_RUC & "\" & rstK("imagen")) = True Then
        
        'Me.Image1.Visible = True
        'Me.Image1.Picture = LoadPicture(App.Path + "\archivos\" + KEY_RUC + "\" + Trim(RstEjecuta!imagen))
        Me.Picthumbnail1.Picture = LoadPicture(App.Path + "\archivos\" + KEY_RUC + "\" + Trim(rstK("imagen")))
        img = Trim(RstEjecuta!imagen)
    Else
        img = ""
        'Me.Picthumbnail1 = Nothing
        'Me.Image1 = Nothing
    End If
End If
        Call llenarCompatibilidad(Me.HfCompatibilidad, Trim(Me.LblCodigoProducto.Caption))
        Call llenarCaracteristica(Me.HfCaracteristica, Trim(Me.LblCodigoProducto.Caption))
       
  End If
  

End Sub
Private Sub CargarLogo()
Dim sql As String
Dim sw As String
Dim imagen As String
imagen = "Invierno.jpg"
 'Me.Image1.Picture = LoadPicture("C:\" + imagen)
'sql = "select imagen From Producto Where cProducto='" & Trim(StrCodTabla) & "'"
'Call ConfiguraRst(sql)
'If rst.RecordCount > 0 Then

'If IsNull(rst(0)) = False Then


'Image1.Picture = Leer_Imagen(CnBd, sql, "imagen")
'End If
'End If
'Set rst = Nothing
End Sub
Sub llena_barra()
Dim x As Integer
strCadena = "SELECT id_producto,cod_barra as CODIGO_BARRAS FROM producto_barras WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND ruc='" & KEY_RUC & "'"
   Call ConfiguraRst(strCadena)
    Set Me.HfgBarras.Recordset = rst
    Me.HfgBarras.Rows = rst.RecordCount + 1
    Me.HfgBarras.ColWidth(0) = 0
    Me.HfgBarras.ColWidth(1) = 2000
   
End Sub

Private Sub Save()
    Dim error As Boolean
    If Val(Me.DtcArea.BoundText) < 1 Then
       MsgBox "Ingrese una [ AREA ] Correcta", vbInformation, KEY_VENDEDOR
       Exit Sub
    End If
    If Val(Me.DtcUnidad.BoundText) < 1 Then
       MsgBox "Ingrese una [ UNIDAD ] Correcta", vbInformation, KEY_VENDEDOR
       Exit Sub
    End If
    If Val(Me.DtcEditorial.BoundText) < 1 Then
       MsgBox "Ingrese una [ EDITORIAL ] Correcta", vbInformation, KEY_VENDEDOR
       Exit Sub
    End If
    If Me.DtcAutor.BoundText = "" Then
       MsgBox "Ingrese una [ AUTOR ] Correcta", vbInformation, KEY_VENDEDOR
       Exit Sub
    End If
    
    If Val(Me.DtcProcedencia.BoundText) < 1 Then
       MsgBox "Ingrese una [ PROCEDENCIA ] Correcta", vbInformation, KEY_VENDEDOR
       Exit Sub
    End If
    
    
    If Val(Me.LblCodigoProducto.Caption) < 1 Then
       in_libro = formato_item(ConsultaUltimoRegistro("biblio_libro", "id_libro", "ruc", KEY_RUC), 5)
    Else
       in_libro = Trim(Me.LblCodigoProducto.Caption)
    End If
       
       
       strCadena = "call put_procesar_libro('" & in_libro & "','" & Trim(Me.TxtCodBarra.Text) & "','" & Me.DTcTipoLibro.BoundText & "','" & Me.DtcUnidad.BoundText & "'," & _
       " '" & UCase(Trim(Me.txtdescripcion.Text)) & "','" & Me.DtcArea.BoundText & "','" & Me.DtcAutor.BoundText & "','" & Me.DtcEditorial.BoundText & "','" & Trim(Me.txtciudad.Text) & "'," & _
       " '" & Val(Me.txtPaginas.Text) & "','" & Trim(Me.txtTomo.Text) & "','-','" & Trim(Me.txtEdicion.Text) & "','" & Me.DtcProcedencia.BoundText & "'," & _
       " '" & Me.TxtTamalo.Text & "','" & Val(Me.txtCantidad.Text) & "','" & Trim(Me.TxtContenido.Text) & "','" & Val(Me.txtPrecio.Text) & "','" & Me.txtAnio.Text & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       
       If Val(Me.LblCodigoProducto.Caption) < 1 Then
           strCadena = "SELECT * FROM almacen WHERE stock='si' and  ruc='" & KEY_RUC & "' ORDER BY id_alm ASC"
           Call ConfiguraRstK(strCadena)
           If rstK.RecordCount < 1 Then
                MsgBox "No hay Ningun Almacen registrado", vbInformation
                MsgBox "Producto NO Grabado, Cree un Almacen", vbInformation
                strCadena = "DELETE FROM producto WHERE id_libro='" & in_libro & "' AND ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                Exit Sub
           End If
           rstK.MoveFirst
           For i = 0 To rstK.RecordCount - 1
             strCadena = "INSERT INTO almacen_producto(id_alm,id_producto,ruc) VALUES ('" & rstK("id_alm") & "','" & Trim(in_libro) & "','" & Trim(KEY_RUC) & "')"
             CnBd.Execute (strCadena)
             rstK.MoveNext
           Next i
       End If
       
       Unload Me
       Call frmLibro.actualizar_update(in_libro)
       
       Exit Sub
    
            
 


End Sub


Private Sub HfgBarras_Click()
If Me.HfgBarras.Rows > 0 Then
    Me.CmdQuitar.Visible = True
Else
    Me.CmdQuitar.Visible = False
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
     
    Case KEY_CANCEL
        Unload Me
  End Select
 
  
  Exit Sub
error:
  MsgBox "A sucedido un Inconveniente", vbInformation
End Sub







Private Sub TxtCodBarra_Change()
strCadena = "SELECT P.nombre_prod FROM producto_barras B,producto P WHERE B.id_producto=P.id_producto AND P.ruc='" & KEY_RUC & "' AND B.ruc='" & KEY_RUC & "'AND B.cod_barra = '" & Trim(Me.TxtCodBarra.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.lblError.Visible = True
    Me.lblCabecera.Visible = True
    Me.lblError.Caption = Trim(rst(0))
    Me.cmdprocesar.Enabled = True
Else
    Me.lblError.Visible = False
    Me.lblCabecera.Visible = False
    Me.cmdprocesar.Enabled = True
End If
Set rst = Nothing
End Sub

Private Sub TxtCodBarra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtdescripcion)
End If
End Sub

Private Sub txtCodCompatible_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtCodCompatible.Text = formato_item(Me.txtCodCompatible.Text, 5)
    strCadena = "SELECT * FROM producto WHERE id_producto='" & Trim(Me.txtCodCompatible.Text) & "' AND ruc='" & KEY_RUC & "' AND id_producto<>'" & Trim(Me.txtCodCompatible.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.TxtCompatible.Text = rst("nombre_prod")
    Else
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
    End If
End If
End Sub














VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmmantenimientos 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   17760
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmmantenimiento 
      BackColor       =   &H00FFFFFF&
      Height          =   7455
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   16245
      Begin VB.TextBox txtbuscartrabajador 
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
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   15240
         TabIndex        =   46
         Top             =   6480
         Width           =   735
      End
      Begin VB.Frame frminsumo 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   975
         Left            =   9960
         TabIndex        =   37
         Top             =   2580
         Visible         =   0   'False
         Width           =   6015
         Begin VB.TextBox txtid_insumo 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1320
            TabIndex        =   43
            Top             =   360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CheckBox chkfacturar 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "FACTURAR"
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
            Left            =   2280
            TabIndex        =   42
            Top             =   60
            Width           =   1215
         End
         Begin VitekeySoft.ChameleonBtn cmdagregar 
            Height          =   735
            Left            =   4800
            TabIndex        =   40
            Top             =   120
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1296
            BTYPE           =   5
            TX              =   "AGREGAR"
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmantenimientos.frx":0000
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txtcantidad 
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
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   1305
            TabIndex        =   38
            Text            =   "1"
            Top             =   60
            Width           =   735
         End
         Begin VB.Label lblproducto 
            BackStyle       =   0  'Transparent
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
            Height          =   330
            Left            =   90
            TabIndex        =   41
            Top             =   480
            Width           =   4335
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CANTIDAD :"
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
            Left            =   150
            TabIndex        =   39
            Top             =   105
            Width           =   915
         End
      End
      Begin VB.TextBox txtrecomendaciones 
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
         ForeColor       =   &H00C00000&
         Height          =   645
         Left            =   9960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   1440
         Width           =   6015
      End
      Begin VB.TextBox txtobservacion 
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
         ForeColor       =   &H00C00000&
         Height          =   645
         Left            =   9960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   480
         Width           =   6015
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfmantenimientos 
         Height          =   4935
         Left            =   1800
         TabIndex        =   24
         Top             =   2280
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   8705
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfinsumo 
         Height          =   2775
         Left            =   9960
         TabIndex        =   27
         Top             =   3600
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4895
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
      Begin VitekeySoft.ChameleonBtn cmdagergarinsumo 
         Height          =   300
         Left            =   9960
         TabIndex        =   28
         Top             =   2160
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "AGREGAR INSUMO"
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
         MICON           =   "frmmantenimientos.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdcerrarframe 
         Height          =   405
         Left            =   14640
         TabIndex        =   29
         Top             =   6960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "CERRAR "
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
         MICON           =   "frmmantenimientos.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesarmantenimiento 
         Height          =   405
         Left            =   13200
         TabIndex        =   30
         Top             =   6960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   714
         BTYPE           =   3
         TX              =   "GRABAR"
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
         MICON           =   "frmmantenimientos.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcResponsable 
         Height          =   330
         Left            =   11160
         TabIndex        =   45
         Top             =   6480
         Width           =   3975
         _ExtentX        =   7011
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
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RESPONSABLE :"
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
         Left            =   9915
         TabIndex        =   44
         Top             =   6520
         Width           =   1185
      End
      Begin VB.Label lblplaca 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         TabIndex        =   36
         Top             =   1440
         Width           =   4455
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLACA :"
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
         Left            =   1110
         TabIndex        =   35
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label lblvehiculo 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         TabIndex        =   34
         Top             =   1080
         Width           =   4455
      End
      Begin VB.Label lblmotor 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         TabIndex        =   33
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label lblchasis 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1800
         TabIndex        =   32
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VEHICULO :"
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
         Left            =   810
         TabIndex        =   31
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RECOMENDACIONES"
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
         Left            =   9900
         TabIndex        =   25
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   9600
         X2              =   9600
         Y1              =   840
         Y2              =   7320
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LISTADO DE MANTENIMIENTOS "
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
         Left            =   1770
         TabIndex        =   23
         Top             =   1920
         Width           =   2445
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   9840
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SERIE MOTOR:"
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
         Left            =   570
         TabIndex        =   20
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº CHASIS :"
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
         Left            =   810
         TabIndex        =   19
         Top             =   360
         Width           =   885
      End
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
      Left            =   1800
      TabIndex        =   6
      Top             =   200
      Width           =   1335
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
      Left            =   1800
      TabIndex        =   5
      Top             =   555
      Width           =   1335
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
      Left            =   5625
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.CheckBox chkTipoComprobante 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TIPO COMPROBANTE"
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
      Height          =   255
      Left            =   11880
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VitekeySoft.ChameleonBtn cmdamortizar 
      Height          =   975
      Left            =   16440
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "NUEVO"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmantenimientos.frx":0070
      PICN            =   "frmmantenimientos.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo dtcalmacen 
      Height          =   330
      Left            =   8520
      TabIndex        =   2
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
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
   Begin VitekeySoft.ChameleonBtn cmdBuscar 
      Height          =   345
      Left            =   14040
      TabIndex        =   3
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "BUSCAR                        "
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
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmantenimientos.frx":2976
      PICN            =   "frmmantenimientos.frx":2992
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpDesde 
      Height          =   315
      Left            =   8520
      TabIndex        =   7
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   192479233
      CurrentDate     =   41251
   End
   Begin MSComCtl2.DTPicker DtpHasta 
      Height          =   315
      Left            =   10440
      TabIndex        =   8
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   192479233
      CurrentDate     =   41251
   End
   Begin MSDataListLib.DataCombo dtcComprobante 
      Height          =   330
      Left            =   11880
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   7455
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   13150
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
   Begin VitekeySoft.ChameleonBtn cmdhistorial 
      Height          =   975
      Left            =   16440
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "HISTORIAL"
      ENAB            =   0   'False
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmantenimientos.frx":2F2C
      PICN            =   "frmmantenimientos.frx":2F48
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
      Height          =   975
      Left            =   16440
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "CERRAR"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmantenimientos.frx":6551
      PICN            =   "frmmantenimientos.frx":656D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº PLACA :"
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
      Left            =   735
      TabIndex        =   17
      Top             =   240
      Width           =   825
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI/ RUC :"
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
      Left            =   735
      TabIndex        =   16
      Top             =   600
      Width           =   825
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHAS :"
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
      Left            =   7515
      TabIndex        =   15
      Top             =   285
      Width           =   705
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE / RAZON SOCIAL :"
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
      Left            =   3450
      TabIndex        =   14
      Top             =   285
      Width           =   2085
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUCURSAL :"
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
      Left            =   7470
      TabIndex        =   13
      Top             =   720
      Width           =   915
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   915
      Left            =   120
      Top             =   60
      Width           =   17415
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   8655
      Left            =   0
      Top             =   0
      Width           =   17760
   End
End
Attribute VB_Name = "frmmantenimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub chkTipoComprobante_Click()
If Me.chkTipoComprobante.Value = 1 Then
    Me.dtcComprobante.Visible = True
Else
    Me.dtcComprobante.Visible = False
End If
End Sub

Private Sub ClbAcciones_HeightChanged(ByVal NewHeight As Single)

End Sub

Private Sub cmdagergarinsumo_Click()
Procedencia = seleccionar_insumo
FrmProducto.Show
Exit Sub
End Sub

Private Sub cmdagregar_Click()
Dim facturar As String
If Val(Me.txtcantidad.Text) > 0 Then
    If Me.chkfacturar.Value = 1 Then
        facturar = "si"
    Else
        facturar = "no"
    End If
    
    strCadena = "INSERT INTO movimiento_venta_mantenimiento_insumos(`id_listado`,`id_producto`,`cantidad`,`pagado`,`ruc`)VALUES " & _
    "('" & Val(Me.hfmantenimientos.TextMatrix(Me.hfmantenimientos.Row, 0)) & "','" & Trim(Me.txtid_insumo.Text) & "','" & Val(Me.txtcantidad.Text) & "','" & facturar & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    Me.frminsumo.Visible = False
    Call llenar_insumo(Me.hfinsumo, Val(Me.hfmantenimientos.TextMatrix(Me.hfmantenimientos.Row, 0)))
    
End If
End Sub





Private Sub Command2_Click()

End Sub



Private Sub Command3_Click()

End Sub

Private Sub cmdCliente_Click()

End Sub

Private Sub cmdfechas_Click()
 
End Sub







Private Sub cmdcerrarframe_Click()
Me.frmmantenimiento.Visible = False
End Sub

Private Sub cmdCerrarpantalla_Click()
Unload Me
End Sub

Private Sub cmdcuentasporcobrar_Click()
  

End Sub

Private Sub cmdgenerarreporte_Click()
strCadena = "SELECT fecha_emision,hora,documento,id_cliente,ncliente,total FROM movimiento_venta v WHERE v.id_comprobante='" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' and id_recibo='0' and v.ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "Rptpagos_realizados", , App.Path + "\Reportes\")
End Sub

Private Sub cmdhistorial_Click()
Me.frmmantenimiento.Visible = True
strCadena = "SELECT * FROM view_mantenimientos WHERE id_mantenimiento='" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.lblchasis.Caption = rst("nro_chasis")
    Me.lblmotor.Caption = rst("nro_motor")
    Me.lblvehiculo.Caption = rst("nombre_prod")
    Me.lblplaca.Caption = rst("placa")
    Call listado_mantenimientos(Me.hfmantenimientos, Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)))
End If

End Sub

Private Sub cmdprocesarmantenimiento_Click()

strCadena = "UPDATE movimiento_venta_mantenimiento_listado SET observacion='" & Trim(Me.txtobservacion.Text) & "',recomendacion='" & Trim(Me.txtrecomendaciones.Text) & "',dni_responsable='" & Me.DtcResponsable.BoundText & "', id_estado='02',fecha_mantenimiento='" & KEY_FECHA & "' WHERE id='" & Val(Me.hfmantenimientos.TextMatrix(Me.hfmantenimientos.Row, 0)) & "'"
CnBd.Execute (strCadena)
 
Me.cmdprocesarmantenimiento.Enabled = False

End Sub


Private Sub Form_Load()
CenterForm Me
Me.Top = 100
Me.DtpDesde.Value = KEY_FECHA
Me.DtpHasta.Value = KEY_FECHA

strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE id_tipoentidad<>'00012' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcalmacen)
Me.dtcalmacen.BoundText = KEY_ALM

strCadena = "SELECT  DISTINCT a.id_doc as Codigo,c.doc_des as Descripcion FROM almacen_comprobante a, comprobantes c WHERE  a.id_doc=c.id_doc and   ruc='" & KEY_RUC & "' AND a.venta='si'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcComprobante)

strCadena = "SELECT  cod_unico as Codigo,p.nombre_completo as Descripcion FROM entidad_empresa e,persona p WHERE  e.cod_unico=p.dni and e.id_personal='si' and e.id_empresa='" & KEY_RUC & "' ORDER BY p.nombre_completo"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcResponsable)
Me.DtcResponsable.BoundText = 0



Call actualizar
End Sub
Private Sub actualizar()
    strCadena = "SELECT * FROM view_mantenimientos ORDER BY id_mantenimiento DESC  "
    Call llenar_grid(Me.HfdPersona)
End Sub
Private Sub llenar_grid(ByVal Grilla As MSHFlexGrid)
Dim nsaldo As Double
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Me.cmdamortizar.Enabled = False
    
    Exit Sub
End If
   
    Grilla.Clear
    Grilla.Rows = 0
   
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1800
           Grilla.ColWidth(3) = 3500
           Grilla.ColWidth(4) = 1500
           Grilla.ColWidth(5) = 1500
           Grilla.ColWidth(6) = 1400
           Grilla.ColWidth(7) = 2500
           Grilla.ColWidth(8) = 1200
           
        Next
         cabecera = "ID_MANTENIMIENTO" & vbTab & "F.EMISION" & vbTab & "COMPROBANTE" & vbTab & "VEHICULO" & vbTab & "CHASIS" & vbTab & "Nº MOTOR" & vbTab & "PLACA" & vbTab & "CLIENTE" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 1 To 8
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If rst("id_estado") = "01" Then
                nestado = "PENDIENTE"
            Else
                nestado = "COMPLETADO"
            End If
             Fila = rst("id_mantenimiento") & vbTab & Format(rst("fecha_venta"), "dd-mm-YYYY") & vbTab & rst("documento") & vbTab & rst("nombre_prod") & vbTab & rst("nro_chasis") & vbTab & rst("nro_motor") & vbTab & rst("placa") & vbTab & rst("ncliente") & vbTab & nestado
             Grilla.AddItem Fila
            
            
            If rst("id_estado") = "01" Then
                For k = 1 To 8
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
            End If
            
        rst.MoveNext
        Next i
        
 ' Grilla.Row = 1
 ' Grilla.col = 0
 ' Grilla.ColSel = 1
 ' Grilla.RowSel = 1
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub
Private Sub llenar_insumo(ByVal Grilla As MSHFlexGrid, ByVal in_listado As Integer)
Dim nsaldo As Double
On Error GoTo salir
strCadena = "SELECT * FROM movimiento_venta_mantenimiento_listado WHERE id='" & in_listado & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txtobservacion.Text = rst("observacion")
    Me.txtrecomendaciones.Text = rst("recomendacion")
    Me.DtcResponsable.BoundText = rst("dni_responsable")
    If rst("id_estado") = "01" Then ' pendiente
       Me.cmdprocesarmantenimiento.Enabled = True
       Me.cmdagergarinsumo.Enabled = True
    Else
        
        Me.cmdprocesarmantenimiento.Enabled = False
        Me.cmdagergarinsumo.Enabled = False
    End If
End If


strCadena = "SELECT * FROM view_mantenimiento_insumo WHERE id_listado='" & in_listado & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Me.cmdamortizar.Enabled = False
    
    Exit Sub
End If
   
    Grilla.Clear
    Grilla.Rows = 0
   
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 4200
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 800
           
        Next
         cabecera = "IDINSUMO" & vbTab & "DESCRIPCION" & vbTab & "CANTIDAD" & vbTab & "PRECIO"
         Grilla.AddItem cabecera
         For k = 1 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        Acumulado = 0
        For i = 0 To rst.RecordCount - 1
            
             
             If rst("pagado") = "si" Then
                precio = rst("precio_venta")
             Else
                precio = 0
             End If
             Acumulado = Acumulado + precio
             Fila = rst("id") & vbTab & rst("nombre_prod") & vbTab & rst("cantidad") & vbTab & Format(precio, "###0.00")
             
             Grilla.AddItem Fila
             If rst("pagado") = "no" Then
                For k = 1 To 3
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
            End If
            
            
            
        rst.MoveNext
        Next i
    Fila = "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(Acumulado, "###0.00")
    Grilla.AddItem Fila
                For k = 2 To 3
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &HC0FFC0
                Next k
                
 ' Grilla.Row = 1
 ' Grilla.col = 0
 ' Grilla.ColSel = 1
 ' Grilla.RowSel = 1
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub


Public Sub listado_mantenimientos(ByVal Grilla As MSHFlexGrid, ByVal in_mantenimiento As String)
Dim nsaldo As Double
On Error GoTo salir
strCadena = "SELECT * FROM view_mantenimientos_listado WHERE id_mantenimiento='" & in_mantenimiento & "' ORDER BY id ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    
    Exit Sub
End If
   
    Grilla.Clear
   Grilla.Rows = rst.RecordCount - 1
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1300
           Grilla.ColWidth(2) = 1300
           Grilla.ColWidth(3) = 3300
           Grilla.ColWidth(4) = 1100
           
         
           
           
           
        Next
         
         cabecera = "IDMANTENIMIENTO" & vbTab & "F. MANTENIMIENTO" & vbTab & "F. REALIZACION" & vbTab & "RESPONSABLE" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         
         For k = 1 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
            rst.MoveFirst
  
        For i = 1 To rst.RecordCount
             Fila = rst("id") & vbTab & Format(rst("fecha_aproximada"), "dd-mm-YYYY") & vbTab & Format(rst("fecha_mantenimiento"), "dd-mm-YYYY") & vbTab & rst("nombre_completo") & vbTab & rst("estado")
             Grilla.AddItem Fila
             rst.MoveNext
             Grilla.RowHeight(i) = 350
        Next i
       

         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub









Private Sub HfdPersona_SelChange()
If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
    Me.cmdhistorial.Enabled = True
Else
    Me.cmdhistorial.Enabled = False
End If
End Sub

Private Sub hfmantenimientos_SelChange()
If Val(Me.hfmantenimientos.TextMatrix(Me.hfmantenimientos.Row, 0)) > 0 Then
    Call llenar_insumo(Me.hfinsumo, Val(Me.hfmantenimientos.TextMatrix(Me.hfmantenimientos.Row, 0)))
End If
End Sub

Private Sub txtbuscartrabajador_Change()
strCadena = "SELECT  cod_unico as Codigo,P.nombre_completo as Descripcion FROM entidad_empresa e,persona p WHERE  e.cod_unico=p.dni and e.id_personal='si' and e.id_empresa='" & KEY_RUC & "' AND p.nombre_completo LIKE '%" & Trim(Me.txtbuscartrabajador.Text) & "%' ORDER BY p.nombre_completo"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcResponsable)
End Sub

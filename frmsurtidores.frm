VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmsurtidores 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   18015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   8055
      Left            =   7800
      TabIndex        =   3
      Top             =   120
      Width           =   9855
      Begin VB.Frame frmframelectura 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TOMA DE LECTURA"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   5415
         Left            =   120
         TabIndex        =   39
         Top             =   2520
         Visible         =   0   'False
         Width           =   8535
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "LECTURA  [SOLES ]"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   240
            TabIndex        =   58
            Top             =   2880
            Width           =   8175
            Begin VB.TextBox txtlectura_fin_soles 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial Rounded MT Bold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1920
               TabIndex        =   60
               Top             =   1050
               Width           =   1455
            End
            Begin VB.TextBox txtlectura_ini_soles 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial Rounded MT Bold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1920
               TabIndex        =   59
               Top             =   600
               Width           =   1455
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfsoles 
               Height          =   1215
               Left            =   3480
               TabIndex        =   68
               Top             =   240
               Width           =   4575
               _ExtentX        =   8070
               _ExtentY        =   2143
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
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "LECTURA ANTERIOR :"
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
               Height          =   255
               Left            =   240
               TabIndex        =   66
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label lbllecturaanterior_soles 
               BeginProperty Font 
                  Name            =   "Arial Rounded MT Bold"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   300
               Left            =   1920
               TabIndex        =   65
               Top             =   200
               Width           =   1455
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "LECTURA FINAL :"
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
               Height          =   255
               Left            =   360
               TabIndex        =   62
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label19 
               BackStyle       =   0  'Transparent
               Caption         =   "LECTURA INICIAL :"
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
               Height          =   255
               Left            =   240
               TabIndex        =   61
               Top             =   720
               Width           =   1215
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "LECTURA  [GALONES]"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1550
            Left            =   240
            TabIndex        =   53
            Top             =   1080
            Width           =   8175
            Begin VB.TextBox txtlectura_ini 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial Rounded MT Bold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1920
               TabIndex        =   55
               Top             =   720
               Width           =   1455
            End
            Begin VB.TextBox txtlectura_fin 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial Rounded MT Bold"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1920
               TabIndex        =   54
               Top             =   1100
               Width           =   1455
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfCombustible 
               Height          =   1215
               Left            =   3480
               TabIndex        =   67
               Top             =   240
               Width           =   4575
               _ExtentX        =   8070
               _ExtentY        =   2143
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
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "LECTURA ANTERIOR :"
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
               Height          =   255
               Left            =   240
               TabIndex        =   64
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label lbllectura_anterior 
               BeginProperty Font 
                  Name            =   "Arial Rounded MT Bold"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   300
               Left            =   1920
               TabIndex        =   63
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "LECTURA INICIAL :"
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
               Height          =   255
               Left            =   240
               TabIndex        =   57
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "LECTURA FINAL :"
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
               Height          =   255
               Left            =   360
               TabIndex        =   56
               Top             =   1080
               Width           =   1215
            End
         End
         Begin VitekeySoft.ChameleonBtn cmdprocesar_lectura 
            Height          =   795
            Left            =   7440
            TabIndex        =   40
            Top             =   4560
            Width           =   975
            _ExtentX        =   1720
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
            FCOL            =   4194304
            FCOLO           =   4194304
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmsurtidores.frx":0000
            PICN            =   "frmsurtidores.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lbllado 
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   4080
            TabIndex        =   76
            Top             =   285
            Width           =   1935
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "LADO :"
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
            Height          =   255
            Left            =   3480
            TabIndex        =   75
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblsurtidor 
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   300
            Left            =   1320
            TabIndex        =   43
            Top             =   600
            Width           =   4695
         End
         Begin VB.Image Image4 
            Height          =   240
            Left            =   8160
            Picture         =   "frmsurtidores.frx":3664
            Top             =   240
            Width           =   240
         End
         Begin VB.Label lblidsurtidor 
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   1320
            TabIndex        =   42
            Top             =   280
            Width           =   1575
         End
         Begin VB.Label Label12 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   360
            TabIndex        =   41
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame frmdetalle_reporte 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GENERAR REPORTE"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   5415
         Left            =   120
         TabIndex        =   45
         Top             =   2520
         Visible         =   0   'False
         Width           =   8535
         Begin MSDataListLib.DataCombo DtcIslareporte 
            Height          =   330
            Left            =   1200
            TabIndex        =   69
            Top             =   1800
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
         Begin MSComCtl2.DTPicker DtpIni 
            Height          =   350
            Left            =   1200
            TabIndex        =   50
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Rounded MT Bold"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   189005825
            CurrentDate     =   43443
         End
         Begin VitekeySoft.ChameleonBtn cmdGenerarReporte 
            Height          =   555
            Left            =   1200
            TabIndex        =   46
            Top             =   2760
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   979
            BTYPE           =   5
            TX              =   "GENERAR REPORTE"
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
            FCOL            =   4194304
            FCOLO           =   4194304
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmsurtidores.frx":6508
            PICN            =   "frmsurtidores.frx":6524
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker DtpFin 
            Height          =   345
            Left            =   3000
            TabIndex        =   51
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Rounded MT Bold"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   189005825
            CurrentDate     =   43443
         End
         Begin MSDataListLib.DataCombo DtcTurno 
            Height          =   330
            Left            =   1200
            TabIndex        =   70
            Top             =   2280
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
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "TURNO :"
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
            Height          =   255
            Left            =   480
            TabIndex        =   72
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ISLA :"
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
            Height          =   255
            Left            =   480
            TabIndex        =   71
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA INI:"
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
            Height          =   255
            Left            =   -120
            TabIndex        =   49
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Image Image5 
            Height          =   240
            Left            =   6840
            Picture         =   "frmsurtidores.frx":9728
            Top             =   240
            Width           =   240
         End
         Begin VB.Label lblid_surtidor_reporte 
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
            Height          =   255
            Left            =   360
            TabIndex        =   48
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   360
            TabIndex        =   47
            Top             =   840
            Width           =   3855
         End
      End
      Begin VB.Frame frmdetallesurtidor 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DETALLE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   120
         TabIndex        =   28
         Top             =   2520
         Visible         =   0   'False
         Width           =   8535
         Begin MSDataListLib.DataCombo dtcIsla 
            Height          =   315
            Left            =   1920
            TabIndex        =   37
            Top             =   720
            Width           =   4695
            _ExtentX        =   8281
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
         Begin VB.TextBox txtdescripcionsurtidor 
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
            Left            =   1920
            TabIndex        =   30
            Top             =   1320
            Width           =   4695
         End
         Begin VB.TextBox txtcodigoproducto_surtidor 
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
            Left            =   1920
            TabIndex        =   29
            Top             =   1920
            Width           =   975
         End
         Begin VitekeySoft.ChameleonBtn cmdprocesarsurtidor 
            Height          =   795
            Left            =   6120
            TabIndex        =   31
            Top             =   3000
            Width           =   975
            _ExtentX        =   1720
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
            FCOL            =   4194304
            FCOLO           =   4194304
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmsurtidores.frx":C5CC
            PICN            =   "frmsurtidores.frx":C5E8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DtcLado 
            Height          =   315
            Left            =   1920
            TabIndex        =   74
            Top             =   2520
            Width           =   2055
            _ExtentX        =   3625
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
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LADO :"
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
            TabIndex        =   73
            Top             =   2640
            Width           =   450
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "ISLA  :"
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
            Height          =   255
            Left            =   360
            TabIndex        =   38
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label15 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   360
            TabIndex        =   36
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblcodigosurtidor 
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
            Height          =   255
            Left            =   1920
            TabIndex        =   35
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label13 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCTO :"
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
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblproductosurtidor 
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
            Height          =   375
            Left            =   3000
            TabIndex        =   32
            Top             =   1860
            Width           =   3615
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   6840
            Picture         =   "frmsurtidores.frx":FC30
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.TextBox txtid_isla 
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
         Left            =   7320
         TabIndex        =   52
         Top             =   2040
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfIslas 
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2566
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfSurtidores 
         Height          =   3900
         Left            =   240
         TabIndex        =   6
         Top             =   3480
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   6879
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
      Begin VitekeySoft.ChameleonBtn cmdNuevoSurtidor 
         Height          =   780
         Left            =   8760
         TabIndex        =   7
         Top             =   3495
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1376
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
         FCOL            =   4194304
         FCOLO           =   4194304
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsurtidores.frx":12AD4
         PICN            =   "frmsurtidores.frx":12AF0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdUpdateSurtidor 
         Height          =   780
         Left            =   8760
         TabIndex        =   8
         Top             =   4275
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1376
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
         FCOL            =   4194304
         FCOLO           =   4194304
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsurtidores.frx":12F42
         PICN            =   "frmsurtidores.frx":12F5E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdEliminarSurtidor 
         Height          =   780
         Left            =   8760
         TabIndex        =   9
         Top             =   5055
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1376
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
         FCOL            =   4194304
         FCOLO           =   4194304
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsurtidores.frx":15597
         PICN            =   "frmsurtidores.frx":155B3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdLecturaSurtidor 
         Height          =   780
         Left            =   8760
         TabIndex        =   10
         Top             =   5835
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1376
         BTYPE           =   5
         TX              =   "LECTURA"
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
         FCOL            =   4194304
         FCOLO           =   4194304
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsurtidores.frx":179FD
         PICN            =   "frmsurtidores.frx":17A19
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdReporte 
         Height          =   780
         Left            =   8760
         TabIndex        =   44
         Top             =   6615
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1376
         BTYPE           =   5
         TX              =   "REPORTES"
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
         FCOL            =   4194304
         FCOLO           =   4194304
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsurtidores.frx":17E8F
         PICN            =   "frmsurtidores.frx":17EAB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "SURTIDORES"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
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
         TabIndex        =   11
         Top             =   2160
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ISLAS DE COMBUSTIBLE"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
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
         TabIndex        =   4
         Top             =   240
         Width           =   2250
      End
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H00400000&
      Height          =   7575
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7455
      Begin VB.Frame frmtanquedetalle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DETALLE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   360
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   5775
         Begin VB.TextBox txtcodigoproductotanque 
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
            Left            =   1920
            TabIndex        =   26
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtmaxima 
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
            Left            =   1920
            TabIndex        =   23
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox txtminima 
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
            Left            =   1920
            TabIndex        =   22
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox txtdescripciontanque 
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
            Left            =   1920
            TabIndex        =   19
            Top             =   600
            Width           =   3615
         End
         Begin VitekeySoft.ChameleonBtn cmdprocesarTanque 
            Height          =   795
            Left            =   4560
            TabIndex        =   24
            Top             =   1560
            Width           =   975
            _ExtentX        =   1720
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
            FCOL            =   4194304
            FCOLO           =   4194304
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmsurtidores.frx":1A47C
            PICN            =   "frmsurtidores.frx":1A498
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   5280
            Picture         =   "frmsurtidores.frx":1DAE0
            Top             =   240
            Width           =   240
         End
         Begin VB.Label lblproductotanque 
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
            Height          =   375
            Left            =   3000
            TabIndex        =   27
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCTO :"
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
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "M.MAXIMA:"
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
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "M.MINIMA:"
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
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label5 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   360
            TabIndex        =   18
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblcodigotanque 
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
            Height          =   255
            Left            =   1920
            TabIndex        =   17
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label4 
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
            ForeColor       =   &H00400000&
            Height          =   255
            Left            =   360
            TabIndex        =   16
            Top             =   360
            Width           =   735
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfTanques 
         Height          =   2415
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4260
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
      Begin VitekeySoft.ChameleonBtn cmdNuevoTanque 
         Height          =   795
         Left            =   6240
         TabIndex        =   12
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
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
         FCOL            =   4194304
         FCOLO           =   4194304
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsurtidores.frx":20984
         PICN            =   "frmsurtidores.frx":209A0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdUpdateTanque 
         Height          =   795
         Left            =   6240
         TabIndex        =   13
         Top             =   1515
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1402
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
         FCOL            =   4194304
         FCOLO           =   4194304
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsurtidores.frx":20DF2
         PICN            =   "frmsurtidores.frx":20E0E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdDeleteTanque 
         Height          =   795
         Left            =   6240
         TabIndex        =   14
         Top             =   2325
         Width           =   855
         _ExtentX        =   1508
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
         FCOL            =   4194304
         FCOLO           =   4194304
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmsurtidores.frx":23447
         PICN            =   "frmsurtidores.frx":23463
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Cisterna_vacia 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3375
         Index           =   4
         Left            =   5640
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Shape Cisterna_llena 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   3615
         Index           =   4
         Left            =   5640
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "TANQUES DE COMBUSTIBLE"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2625
      End
      Begin VB.Shape Cisterna_vacia 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3375
         Index           =   3
         Left            =   4320
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Shape Cisterna_vacia 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3375
         Index           =   2
         Left            =   3000
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Shape Cisterna_vacia 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3375
         Index           =   1
         Left            =   1680
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Shape Cisterna_llena 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   3615
         Index           =   3
         Left            =   4320
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Shape Cisterna_llena 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   3615
         Index           =   2
         Left            =   3000
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Shape Cisterna_llena 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   3615
         Index           =   1
         Left            =   1680
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Shape Cisterna_vacia 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3375
         Index           =   0
         Left            =   360
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Shape Cisterna_llena 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   3615
         Index           =   0
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   1215
      End
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   17640
      Picture         =   "frmsurtidores.frx":258AD
      Top             =   120
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8295
      Left            =   0
      Top             =   0
      Width           =   18015
   End
End
Attribute VB_Name = "frmsurtidores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub nuevo_tanque()
Me.frmtanquedetalle.Visible = True
Me.lblcodigotanque.Caption = ""
Me.txtdescripciontanque.Text = ""
Me.txtcodigoproductotanque.Text = ""
Me.lblproductotanque.Caption = ""
Me.txtmaxima.Text = 0
Me.txtminima.Text = "0"

Call Resalta(Me.txtdescripciontanque)


End Sub
Private Sub update_tanque(ByVal in_tanque As String)
strCadena = "SELECT * FROM tanque where id_tanque='" & Val(in_tanque) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.frmtanquedetalle.Visible = True
   Me.lblcodigotanque.Caption = Format(rst("id_tanque"), "00000")
   Me.txtdescripciontanque.Text = rst("descripcion")
   Me.txtcodigoproductotanque.Text = rst("id_producto")
   Me.lblproductotanque.Caption = get_producto(rst("id_producto"))
   Me.txtminima.Text = rst("minimo")
   Me.txtmaxima.Text = rst("maxima")
End If
End Sub
Private Sub nuevo_surtidor()
Me.frmdetallesurtidor.Visible = True
Me.lblcodigosurtidor.Caption = ""
Me.txtdescripcionsurtidor.Text = ""
Me.txtcodigoproducto_surtidor.Text = ""
Me.lblproductosurtidor.Caption = ""
Me.dtcIsla.BoundText = Me.HfIslas.TextMatrix(Me.HfIslas.Row, 0)

End Sub
Private Sub nuevo_reporte()
On Error GoTo salir
Me.frmdetalle_reporte.Visible = True
Me.lblid_surtidor_reporte.Caption = Me.HfSurtidores.TextMatrix(Me.HfSurtidores.Row, 0)
Me.Label18.Caption = Me.HfSurtidores.TextMatrix(Me.HfSurtidores.Row, 1)
Me.DtpIni.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA
Me.DtcIslareporte.BoundText = Me.txtid_isla.Text

Exit Sub
salir:
End Sub

Private Sub update_surtidor(ByVal in_surtidor As String)
strCadena = "SELECT * FROM surtidor where id_surtidor='" & Val(in_surtidor) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.frmdetallesurtidor.Visible = True
    Me.lblcodigosurtidor.Caption = Format(in_surtidor)
    Me.txtdescripcionsurtidor.Text = rst("descripcion")
    Me.txtcodigoproducto_surtidor.Text = rst("id_producto")
    Me.DtcLado.BoundText = rst("id_lado")
    Me.lblproductosurtidor.Caption = get_producto(rst("id_producto"))
     
End If
End Sub


Private Sub load_surtidor(ByVal in_surtidor As String)

strCadena = "SELECT * FROM view_lectura_surtidor where id_surtidor='" & Val(in_surtidor) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.frmframelectura.Visible = True
    Me.lbllado.Caption = rst("lado")
    Me.lblidsurtidor.Caption = Format(in_surtidor, "00000")
    Me.lblsurtidor.Caption = rst("descripcion")
    Me.lbllectura_anterior.Caption = Format(rst("lectura_anterior"), "###00.000")
    
    Me.txtlectura_ini.Text = Format(rst("lectura_anterior"), "###00.000")
    
    Me.lbllecturaanterior_soles.Caption = Format(rst("lectura_anterior_soles"), "###00.000")
    Me.txtlectura_ini_soles.Text = Format(rst("lectura_anterior_soles"), "###00.000")
    Me.txtlectura_fin.Text = 0
    Me.txtlectura_fin_soles.Text = 0
    Me.DtcIslareporte.BoundText = Trim(Me.txtid_isla.Text)
    Call llenarGrid_combustible(Me.hfCombustible, in_surtidor)
    
    Call llenarGrid_soles(Me.hfsoles, in_surtidor)
End If
End Sub


Private Sub ChameleonBtn1_Click()

End Sub

Private Sub cmdDeleteTanque_Click()
strCadena = "CALL put_tanques('" & Val(Me.lblcodigotanque.Caption) & "','" & Trim(Me.txtdescripciontanque.Text) & "','" & Trim(Me.txtcodigoproductotanque.Text) & "','" & Val(Me.txtminima.Text) & "','" & Val(Me.txtmaxima.Text) & "','si','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
Me.frmtanquedetalle.Visible = False
Call Me.llenarGrid(Me.HfTanques)

End Sub


Private Sub cmdEliminarSurtidor_Click()
If Val(Me.HfSurtidores.TextMatrix(Me.HfSurtidores.Row, 0)) > 0 Then

If MsgBox("Desea Eliminar el Surtidor", vbYesNo, KEY_VENDEDOR) = vbYes Then
    strCadena = "CALL put_surtidor('" & Val(Me.HfSurtidores.TextMatrix(Me.HfSurtidores.Row, 0)) & "','" & Me.dtcIsla.BoundText & "','" & Trim(Me.txtdescripcionsurtidor.Text) & "','" & Trim(Me.txtcodigoproducto_surtidor.Text) & "','si','0','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Me.frmdetallesurtidor.Visible = False
    Call Me.llenarGrid_surtidor(Me.HfSurtidores, Me.HfIslas.TextMatrix(Me.HfIslas.Row, 0))
End If
End If
End Sub

Private Sub cmdgenerarreporte_Click()
'strCadena = "SELECT id_surtidor,descripcion,lectura_ini,lectura_fin,lectura_anterior,fecha,hora,hora_cadena,func_get_precio(id_producto,'" & KEY_ALM & "',ruc) as precio,operador,almacen FROM view_lectura_surtidor WHERE fecha>='" & Format(Me.DtpIni.Value, "YYYY-mm-dd") & "' and fecha<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  id_isla='" & Val(Me.txtid_isla.Text) & "' and ruc='" & KEY_RUC & "'"
'Call ConfiguraRst(strCadena)
'Ans = ShowMultiReport(rst, "rptlectura_grifo", , App.Path + "\Reportes\")

Call impresion_cierre_grifo(Me.DtcIslareporte.BoundText, Me.DtcTurno.BoundText, Format(Me.DtpIni.Value, "YYYY-mm-dd"), Format(Me.DtpFin.Value, "YYYY-mm-dd"))

End Sub



Private Sub cmdLecturaSurtidor_Click()
Call load_surtidor(Me.HfSurtidores.TextMatrix(Me.HfSurtidores.Row, 0))
End Sub

Private Sub cmdNuevoSurtidor_Click()
Call nuevo_surtidor
End Sub

Private Sub cmdNuevoTanque_Click()
Call nuevo_tanque
End Sub

Private Sub cmdprocesar_lectura_Click()
Dim in_surtidor As Double
Dim in_inicial As Double
Dim in_final As Double
Dim in_hora_actual As Variant
Dim in_inicial_soles As Double
Dim in_final_soles As Double


in_surtidor = Val(lblidsurtidor.Caption)

in_inicial = Format(Val(Me.txtlectura_ini.Text), "###00.000")
in_final = Format(Val(Me.txtlectura_fin.Text), "###00.000")

in_inicial_soles = Format(Val(Me.txtlectura_ini_soles.Text), "###00.000")
in_final_soles = Format(Val(Me.txtlectura_fin_soles.Text), "###00.000")


strCadena = "call put_lectura_surtidor('" & in_surtidor & "','" & in_inicial & "','" & in_final & "','" & in_inicial_soles & "','" & in_final_soles & "','" & KEY_USUARIO & "','" & KEY_FECHA & "',CURTIME(),'" & KEY_TURNO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

MsgBox "LECTURA DE SURTIDOR REALIZADA.!!!", vbInformation
Call load_surtidor(Me.HfSurtidores.TextMatrix(Me.HfSurtidores.Row, 0))
'Me.frmframelectura.Visible = False
End Sub

Private Sub cmdprocesarsurtidor_Click()

strCadena = "CALL put_surtidor('" & Val(Me.lblcodigosurtidor.Caption) & "','" & Me.dtcIsla.BoundText & "','" & Trim(Me.txtdescripcionsurtidor.Text) & "','" & Trim(Me.txtcodigoproducto_surtidor.Text) & "','no','" & Me.DtcLado.BoundText & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
Me.frmdetallesurtidor.Visible = False
Call Me.llenarGrid_surtidor(Me.HfSurtidores, Me.HfIslas.TextMatrix(Me.HfIslas.Row, 0))



End Sub

Private Sub cmdprocesarTanque_Click()

strCadena = "CALL put_tanques('" & Val(Me.lblcodigotanque.Caption) & "','" & Trim(Me.txtdescripciontanque.Text) & "','" & Trim(Me.txtcodigoproductotanque.Text) & "','" & Val(Me.txtminima.Text) & "','" & Val(Me.txtmaxima.Text) & "','no','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
Me.frmtanquedetalle.Visible = False
Call Me.llenarGrid(Me.HfTanques)

End Sub

Private Sub cmdReporte_Click()
Call nuevo_reporte
End Sub

Private Sub cmdUpdateSurtidor_Click()
Call update_surtidor(Me.HfSurtidores.TextMatrix(Me.HfSurtidores.Row, 0))
End Sub

Private Sub cmdUpdateTanque_Click()
Call update_tanque(Me.HfTanques.TextMatrix(Me.HfTanques.Row, 0))
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100

strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE id_tipoentidad='00012' and id_sucursal='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcIsla)

strCadena = "SELECT id_lado as Codigo,descripcion as Descripcion FROM surtidor_lado ORDER BY id_lado ASC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcLado)


strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE id_tipoentidad='00012' and id_sucursal='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcIslareporte)



strCadena = "SELECT id_turno as Codigo,descripcion as Descripcion FROM turno WHERE  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTurno)
Me.DtcTurno.BoundText = KEY_TURNO




Call Me.llenarGrid(Me.HfTanques)
Call llenarGrid_isla(Me.HfIslas)

If KEY_CARGO = "00004" Then
   Me.cmdNuevoSurtidor.Enabled = True
   Me.cmdEliminarSurtidor.Enabled = True
   Me.cmdUpdateSurtidor.Enabled = True
Else
    Me.cmdNuevoSurtidor.Enabled = False
   Me.cmdEliminarSurtidor.Enabled = False
   Me.cmdUpdateSurtidor.Enabled = False
End If

End Sub



Private Sub HfIslas_SelChange()
Call Me.llenarGrid_surtidor(Me.HfSurtidores, Me.HfIslas.TextMatrix(Me.HfIslas.Row, 0))
End Sub

Private Sub HfSurtidores_SelChange()
If Val(Me.HfSurtidores.TextMatrix(Me.HfSurtidores.Row, 0)) > 0 Then
    If KEY_CARGO = "00004" Then
    
    Me.cmdNuevoSurtidor.Enabled = True
    Me.cmdEliminarSurtidor.Enabled = True
    Me.cmdUpdateSurtidor.Enabled = True
     Me.cmdLecturaSurtidor.Enabled = True
     Me.cmdReporte.Enabled = True
Else
    Me.cmdLecturaSurtidor.Enabled = True
    Me.cmdGenerarReporte.Enabled = True
    Me.cmdReporte.Enabled = True
    End If
Else
    
    Me.cmdEliminarSurtidor.Enabled = False
    Me.cmdUpdateSurtidor.Enabled = False
    Me.cmdLecturaSurtidor.Enabled = False
    Me.cmdReporte.Enabled = False
End If
End Sub

Private Sub HfTanques_SelChange()
Call load_tanque(Trim(Me.HfTanques.TextMatrix(Me.HfTanques.Row, 1)))
End Sub
Private Sub load_tanque(ByVal in_producto As String)

For i = 0 To 4
    Me.Cisterna_llena(i).Visible = False
    Me.Cisterna_vacia(i).Visible = False
   ' Me.lblCisterna(i).Visible = False
Next i


strCadena = "SELECT descripcion,minimo,maxima,funct_stock(id_producto,'" & KEY_ALM & "',ruc) as stock FROM view_producto_tanque WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
  
   For i = 0 To rst.RecordCount - 1
        Me.Cisterna_llena(i).Visible = True
        Me.Cisterna_vacia(i).Visible = True
        
       'Me.lblCisterna(i).Caption = rst("descripcion")
      ' Me.lblmin.Caption = rst("minimo")
      ' Me.lblmedio.Caption = (rst("minimo") + rst("maxima")) / 2
       'Me.lblmax.Caption = rst("maxima")
       If rst("stock") > 0 Then
       
       If rst("stock") > rst("maxima") Then
          Me.Cisterna_vacia(i) = 0
       Else
         Me.Cisterna_vacia(i).Height = 3855 - rst("stock") * 3855 / rst("maxima")

       End If
       Else
        Me.Cisterna_vacia(i).Height = 3855
       End If
       
       
       rst.MoveNext
   Next i
End If


End Sub


Private Sub Image1_Click()
Me.frmtanquedetalle.Visible = False
End Sub

Private Sub llenarGrid_combustible(ByVal Grilla As MSHFlexGrid, ByVal in_surtidor As String)
Dim in_precio As String
On Error GoTo salir
strCadena = "SELECT * FROM view_lectura_surtidor WHERE id_surtidor='" & in_surtidor & "' and id_turno='" & KEY_TURNO & "' and  ruc='" & KEY_RUC & "' LIMIT 2"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
  
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 850
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 1200
    Next
         cabecera = "CODIGO" & vbTab & "HORA" & vbTab & "INICIAL" & vbTab & "FINAL" & vbTab & "OPERADOR"
         Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
            Fila = rst("id_lectura") & vbTab & rst("hora_cadena") & vbTab & Format(rst("lectura_ini"), "#,##0.000") & vbTab & Format(rst("lectura_fin"), "#,##0.000") & vbTab & rst("operador")
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub llenarGrid_soles(ByVal Grilla As MSHFlexGrid, ByVal in_surtidor As String)
Dim in_precio As String
On Error GoTo salir
strCadena = "SELECT * FROM view_lectura_surtidor WHERE id_surtidor='" & in_surtidor & "' and id_turno='" & KEY_TURNO & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
  
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 850
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 1200
    Next
         cabecera = "CODIGO" & vbTab & "HORA" & vbTab & "INICIAL" & vbTab & "FINAL" & vbTab & "OPERADOR"
         Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
            
            Fila = rst("id_lectura") & vbTab & rst("hora_cadena") & vbTab & Format(rst("lectura_ini_soles"), "#,##0.000") & vbTab & Format(rst("lectura_fin_soles"), "#,##0.000") & vbTab & rst("operador")
            Grilla.AddItem Fila
            
            rst.MoveNext
        Next i
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
Dim in_precio As String
On Error GoTo salir
strCadena = "SELECT * FROM tanque WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
  
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 3500
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1000
    Next
         cabecera = "CODIGO" & vbTab & "PRODUCTO" & vbTab & "DESCRIPCION" & vbTab & "MINIMA" & vbTab & "MAXIMA"
         Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
            Fila = rst("id_tanque") & vbTab & rst("id_producto") & vbTab & rst("descripcion") & vbTab & rst("minimo") & vbTab & rst("maxima")
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub




Sub llenarGrid_isla(ByVal Grilla As MSHFlexGrid)
Dim in_precio As String
On Error GoTo salir
strCadena = "SELECT id_alm as Codigo,descripcion,`funct_estado_almacen`(dni_save,id_alm) as operador  FROM almacen WHERE id_tipoentidad='00012' and id_sucursal='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
  
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 3500
           Grilla.ColWidth(2) = 2500
        
    Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "RESPONSABLE"
         Grilla.AddItem cabecera
         For k = 0 To 2
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("Codigo") & vbTab & rst("descripcion") & vbTab & rst("operador")
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Sub llenarGrid_surtidor(ByVal Grilla As MSHFlexGrid, ByVal in_isla As String)
Dim in_precio As String
On Error GoTo salir
Me.txtid_isla.Text = in_isla
strCadena = "SELECT *  FROM view_surtidor WHERE id_isla='" & in_isla & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
  
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 1500
           Grilla.ColWidth(2) = 2500
           Grilla.ColWidth(3) = 3000
    Next
         cabecera = "CODIGO" & vbTab & "LADO" & vbTab & "DESCRIPCION" & vbTab & "PRODUCTO"
         Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
         Next k
        
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = Format(rst("id_surtidor"), "00000") & vbTab & rst("LADO") & vbTab & rst("descripcion") & vbTab & get_producto(rst("id_producto"))
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub



Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Image3_Click()
Me.frmdetallesurtidor.Visible = False
End Sub

Private Sub Image4_Click()
Me.frmframelectura.Visible = False
End Sub

Private Sub Image5_Click()
Me.frmdetalle_reporte.Visible = False
End Sub

Private Sub txtcodigoproducto_surtidor_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    
    Procedencia = buscar
    FrmProducto.Show
    Exit Sub
    
End If

End Sub

Private Sub txtcodigoproductotanque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
End If
End Sub

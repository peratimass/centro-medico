VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCompras 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameCaracteristicas 
      BackColor       =   &H00FFFFFF&
      Height          =   7215
      Left            =   12825
      TabIndex        =   105
      Top             =   165
      Visible         =   0   'False
      Width           =   7335
      Begin VB.Frame frmip 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   400
         Left            =   240
         TabIndex        =   221
         Top             =   5640
         Visible         =   0   'False
         Width           =   4335
         Begin VB.TextBox txtIP 
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
            Height          =   315
            Left            =   1215
            TabIndex        =   222
            Top             =   40
            Width           =   2655
         End
         Begin VB.Label lblip 
            BackStyle       =   0  'Transparent
            Caption         =   "N� IP           :"
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
            Height          =   165
            Left            =   0
            TabIndex        =   223
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmpoliza 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   400
         Left            =   240
         TabIndex        =   218
         Top             =   5160
         Visible         =   0   'False
         Width           =   4335
         Begin VB.TextBox txtPoliza 
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
            Height          =   315
            Left            =   1215
            TabIndex        =   219
            Top             =   40
            Width           =   2655
         End
         Begin VB.Label lblpoliza 
            BackStyle       =   0  'Transparent
            Caption         =   "POLIZA       :"
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
            Height          =   165
            Left            =   0
            TabIndex        =   220
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmitem 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   400
         Left            =   240
         TabIndex        =   149
         Top             =   6120
         Visible         =   0   'False
         Width           =   4335
         Begin VB.TextBox txtitemdua 
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
            Height          =   315
            Left            =   1215
            TabIndex        =   150
            Top             =   45
            Width           =   2655
         End
         Begin VB.Label lblitem 
            BackStyle       =   0  'Transparent
            Caption         =   "ITEM          :"
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
            Height          =   165
            Left            =   0
            TabIndex        =   151
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame frmmotor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   400
         Left            =   240
         TabIndex        =   145
         Top             =   4680
         Visible         =   0   'False
         Width           =   4335
         Begin VB.TextBox txtseriemotor 
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
            Height          =   315
            Left            =   1215
            TabIndex        =   147
            Top             =   40
            Width           =   2655
         End
         Begin VB.TextBox txtMotorG 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4455
            TabIndex        =   146
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label lblmotor 
            BackStyle       =   0  'Transparent
            Caption         =   "N� MOTOR :"
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
            Height          =   165
            Left            =   0
            TabIndex        =   148
            Top             =   120
            Width           =   900
         End
      End
      Begin VB.Frame frmchasis 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   400
         Left            =   240
         TabIndex        =   141
         Top             =   4200
         Visible         =   0   'False
         Width           =   4335
         Begin VB.TextBox txtseriechasis 
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
            Height          =   315
            Left            =   1215
            TabIndex        =   143
            Top             =   45
            Width           =   2655
         End
         Begin VB.TextBox txtChasisG 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4470
            TabIndex        =   142
            Top             =   0
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label lblchasis 
            BackStyle       =   0  'Transparent
            Caption         =   "N� CHASIS :"
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
            Height          =   165
            Left            =   0
            TabIndex        =   144
            Top             =   90
            Width           =   915
         End
      End
      Begin VB.Frame frmvim 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   400
         Left            =   240
         TabIndex        =   138
         Top             =   3720
         Visible         =   0   'False
         Width           =   4335
         Begin VB.TextBox txtnumeroserie 
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
            Height          =   315
            Left            =   1215
            TabIndex        =   139
            Top             =   45
            Width           =   2655
         End
         Begin VB.Label lblvin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VIN         :"
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
            Left            =   120
            TabIndex        =   140
            Top             =   60
            Width           =   720
         End
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   345
         Left            =   1440
         TabIndex        =   106
         Top             =   6720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompra1.frx":0000
         PICN            =   "FrmCompra1.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfItem 
         Height          =   3375
         Left            =   240
         TabIndex        =   107
         Top             =   240
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   5953
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   -2147483635
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin VitekeySoft.ChameleonBtn cmdcerrarpantalla 
         Height          =   345
         Left            =   2880
         TabIndex        =   108
         Top             =   6720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompra1.frx":05B6
         PICN            =   "FrmCompra1.frx":05D2
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
   Begin VB.Frame FRameiva 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IVA"
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
      Height          =   975
      Left            =   11120
      TabIndex        =   249
      Top             =   5880
      Visible         =   0   'False
      Width           =   1550
      Begin MSDataListLib.DataCombo DtcIva 
         Height          =   315
         Left            =   600
         TabIndex        =   250
         Top             =   195
         Width           =   855
         _ExtentX        =   1508
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
      Begin MSDataListLib.DataCombo DtcRetencionFuente 
         Height          =   315
         Left            =   600
         TabIndex        =   253
         Top             =   600
         Width           =   855
         _ExtentX        =   1508
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
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
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
         Index           =   17
         Left            =   120
         TabIndex        =   255
         Top             =   240
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
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
         Index           =   16
         Left            =   105
         TabIndex        =   254
         Top             =   480
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R (FTE):"
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
         Index           =   15
         Left            =   40
         TabIndex        =   252
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R (IVA):"
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
         Index           =   14
         Left            =   40
         TabIndex        =   251
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox txtNumeroAutorizacion 
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
      Left            =   8400
      MaxLength       =   80
      TabIndex        =   247
      Top             =   8850
      Width           =   4095
   End
   Begin VB.TextBox txtalmacen 
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
      Left            =   7410
      MaxLength       =   80
      TabIndex        =   246
      Top             =   720
      Width           =   735
   End
   Begin VB.Frame frmCuentaDescargo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CUENTA DESCARGO"
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
      Height          =   975
      Left            =   12120
      TabIndex        =   243
      Top             =   2520
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdcerrarPlanilla 
         Height          =   255
         Left            =   4800
         Picture         =   "FrmCompra1.frx":35E7
         Style           =   1  'Graphical
         TabIndex        =   245
         Top             =   240
         Width           =   255
      End
      Begin MSDataListLib.DataCombo DtcCuentaDescargo 
         Height          =   345
         Left            =   240
         TabIndex        =   244
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
   End
   Begin VB.Frame frmProrrateo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "UPDATE PRORRATEO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12840
      TabIndex        =   239
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtid_detallecompra 
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
         Left            =   2040
         MaxLength       =   80
         TabIndex        =   242
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtMontoProrrateo 
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
         Left            =   960
         MaxLength       =   80
         TabIndex        =   241
         Top             =   300
         Width           =   1095
      End
      Begin VB.Image cmdcerrar 
         Height          =   240
         Left            =   2400
         Picture         =   "FrmCompra1.frx":648B
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO:"
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
         Index           =   13
         Left            =   270
         TabIndex        =   240
         Top             =   360
         Width           =   585
      End
   End
   Begin VB.CheckBox chk_obsequio 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12240
      TabIndex        =   224
      ToolTipText     =   "OBSEQUIO"
      Top             =   7320
      Width           =   255
   End
   Begin VB.Frame framedua 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   11400
      TabIndex        =   115
      Top             =   2880
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox txtseriedua 
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
         Left            =   2160
         TabIndex        =   206
         Top             =   795
         Width           =   495
      End
      Begin VB.TextBox txta�omodelo 
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
         Height          =   285
         Left            =   2145
         TabIndex        =   119
         Top             =   1515
         Width           =   1455
      End
      Begin VB.TextBox txtaniodua 
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
         Height          =   285
         Left            =   2145
         TabIndex        =   118
         Top             =   1155
         Width           =   1455
      End
      Begin VB.TextBox txtnumero_dua 
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
         Height          =   285
         Left            =   2760
         TabIndex        =   117
         Top             =   795
         Width           =   855
      End
      Begin VB.TextBox txta�ofabricacion 
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
         Height          =   285
         Left            =   2145
         TabIndex        =   116
         Top             =   360
         Width           =   1455
      End
      Begin VitekeySoft.ChameleonBtn cmdcerrardua 
         Height          =   345
         Left            =   1680
         TabIndex        =   125
         Top             =   1920
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         BTYPE           =   5
         TX              =   "CERRAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
         MICON           =   "FrmCompra1.frx":932F
         PICN            =   "FrmCompra1.frx":934B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "A�O MODELO :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   705
         TabIndex        =   123
         Top             =   1515
         Width           =   1140
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A�O DUA  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   990
         TabIndex        =   122
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N� DUA :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   1200
         TabIndex        =   121
         Top             =   795
         Width           =   645
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "A�O FABRICACION :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   240
         TabIndex        =   120
         Top             =   435
         Width           =   1605
      End
   End
   Begin VB.Frame frmvinculadas 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "IMPORTACIONES VINCULADAS"
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
      Height          =   3255
      Left            =   7320
      TabIndex        =   196
      Top             =   2040
      Visible         =   0   'False
      Width           =   12735
      Begin VB.Frame frmvinculada_detalle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "VINCULADA"
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
         Height          =   2295
         Left            =   120
         TabIndex        =   199
         Top             =   360
         Visible         =   0   'False
         Width           =   12375
         Begin VitekeySoft.ChameleonBtn cmdAgregar_vinculacion 
            Height          =   495
            Left            =   3600
            TabIndex        =   205
            Top             =   1560
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "AGREGAR GASTO VINCULADO"
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
            MICON           =   "FrmCompra1.frx":C360
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txtMonto_asignado 
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
            Left            =   1680
            MaxLength       =   80
            TabIndex        =   204
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox txtMonto_porcentaje 
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
            Left            =   1680
            MaxLength       =   80
            TabIndex        =   203
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox TxtMontoTotal_vinculado 
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
            Left            =   1680
            MaxLength       =   80
            TabIndex        =   202
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtComprobante_vinculado 
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
            Left            =   1680
            MaxLength       =   80
            TabIndex        =   200
            Top             =   400
            Width           =   2775
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MONTO ASIGNADO :"
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
            Index           =   12
            Left            =   120
            TabIndex        =   238
            Top             =   1920
            Width           =   1365
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
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
            Index           =   11
            Left            =   540
            TabIndex        =   237
            Top             =   1440
            Width           =   945
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MONTO TOTAL :"
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
            Index           =   10
            Left            =   420
            TabIndex        =   236
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COMPROBANTE:"
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
            Index           =   9
            Left            =   360
            TabIndex        =   235
            Top             =   480
            Width           =   1125
         End
         Begin VB.Label lblid_compra_vinculada 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   6240
            TabIndex        =   201
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VitekeySoft.ChameleonBtn cmdAgregarVinculacion 
         Height          =   350
         Left            =   240
         TabIndex        =   198
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "AGREGAR GASTO VINCULADO"
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
         MICON           =   "FrmCompra1.frx":C37C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfImportaciones 
         Height          =   2295
         Left            =   240
         TabIndex        =   197
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   4048
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         ForeColorSel    =   16777215
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
      Begin VB.Image cmdcerrar_compartir 
         Height          =   240
         Left            =   12405
         Picture         =   "FrmCompra1.frx":C398
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.CheckBox chk_valor_venta 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "VALOR VENTA"
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
      Height          =   375
      Left            =   10920
      TabIndex        =   217
      ToolTipText     =   "CONVERTIR A MONEDA NACIONAL"
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtLote 
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
      Left            =   16560
      MaxLength       =   80
      TabIndex        =   216
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdInventario 
      Caption         =   "INVENTARIO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   214
      Top             =   8280
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.Frame Frame_Relacionado 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "COMPROBANTE RELACIONADO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2055
      Left            =   16020
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtid_comprobante_relacionado 
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
         Left            =   0
         MaxLength       =   80
         TabIndex        =   213
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame frmtiponota 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   495
         Left            =   120
         TabIndex        =   210
         Top             =   1440
         Visible         =   0   'False
         Width           =   3855
         Begin MSDataListLib.DataCombo DtcTipoNota 
            Height          =   315
            Left            =   1080
            TabIndex        =   211
            Top             =   120
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
         Begin VB.Label Label57 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MOTIVO NOTA:"
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
            Left            =   0
            TabIndex        =   212
            Top             =   120
            Width           =   1065
         End
      End
      Begin VB.TextBox txtBuscarproveedor 
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
         Left            =   3160
         MaxLength       =   80
         TabIndex        =   192
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox TxtFechaEmision 
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
         Left            =   1440
         MaxLength       =   80
         TabIndex        =   99
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox TxtNumeroR 
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
         Left            =   2685
         MaxLength       =   80
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtSerieR 
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
         Left            =   1920
         MaxLength       =   80
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DtcRelacionado 
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSDataListLib.DataCombo DtcProveedor 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   3015
         _ExtentX        =   5318
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
      Begin VitekeySoft.ChameleonBtn cmdImportar 
         Height          =   300
         Left            =   3150
         TabIndex        =   209
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "IMPORTAR"
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
         MICON           =   "FrmCompra1.frx":F23C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image cmdcerrarnota 
         Height          =   240
         Left            =   3840
         Picture         =   "FrmCompra1.frx":F258
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMPROBANTE RELACIONADO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   195
         TabIndex        =   191
         Top             =   120
         Width           =   2085
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA EMISION :"
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
         Left            =   270
         TabIndex        =   98
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lblProveedor 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   4020
      End
   End
   Begin VB.CheckBox chk_afecto_costo 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "AFECTO AL COSTO"
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
      Height          =   260
      Left            =   12120
      TabIndex        =   207
      Top             =   960
      Width           =   4395
   End
   Begin VB.TextBox txtNumero_orden 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Height          =   300
      Left            =   5775
      MaxLength       =   80
      TabIndex        =   195
      Top             =   885
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.TextBox TxtSerie_orden 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      Height          =   300
      Left            =   5220
      MaxLength       =   80
      TabIndex        =   194
      Top             =   885
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CheckBox ChkExtraer 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "EXTRAER"
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
      Height          =   300
      Left            =   2700
      TabIndex        =   193
      Top             =   885
      Width           =   880
   End
   Begin VB.TextBox TxtTotalRetencion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   186
      Text            =   "0.00"
      Top             =   8860
      Width           =   1455
   End
   Begin VB.TextBox txtBuscar 
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
      Left            =   10920
      MaxLength       =   80
      TabIndex        =   185
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame frmgastos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1260
      Left            =   16560
      TabIndex        =   178
      Top             =   7880
      Width           =   3440
      Begin VB.CheckBox chk_cantidad_gastos 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "CANTIDAD"
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
         Height          =   200
         Left            =   1800
         TabIndex        =   183
         Top             =   700
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chk_valor_venta_gasto 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "VALOR VENTA"
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
         Height          =   200
         Left            =   1800
         TabIndex        =   182
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox lblgastos 
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
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   179
         Text            =   "0.00"
         Top             =   360
         Width           =   1335
      End
      Begin VitekeySoft.ChameleonBtn cmdProrrateoGastos 
         Height          =   380
         Left            =   1800
         TabIndex        =   181
         Top             =   45
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "GENERAR PRORRATEO"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompra1.frx":120FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdGastos 
         Height          =   375
         Left            =   120
         TabIndex        =   184
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "AGREGAR"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompra1.frx":12118
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdAplicarGastos 
         Height          =   300
         Left            =   1800
         TabIndex        =   190
         Top             =   920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "APLICAR PRORRATEO"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompra1.frx":12134
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OTROS GASTOS"
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
         Left            =   195
         TabIndex        =   180
         Top             =   120
         Width           =   1185
      End
   End
   Begin VB.Frame frmImportacion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   1260
      Left            =   12600
      TabIndex        =   166
      Top             =   7880
      Width           =   3855
      Begin VB.CheckBox chk_cantidad_importacion 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "CANTIDAD"
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
         Height          =   200
         Left            =   2160
         TabIndex        =   177
         Top             =   700
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkValorVenta_importacion 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Caption         =   "VALOR VENTA"
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
         Height          =   200
         Left            =   2160
         TabIndex        =   176
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox TxtCif 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   960
         TabIndex        =   174
         Text            =   "0.00"
         Top             =   940
         Width           =   1095
      End
      Begin VB.TextBox TxtFlete 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   960
         TabIndex        =   173
         Text            =   "0.00"
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtSeguro 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   960
         TabIndex        =   172
         Text            =   "0.00"
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox txtFob 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   960
         TabIndex        =   171
         Text            =   "0.00"
         Top             =   30
         Width           =   1095
      End
      Begin VitekeySoft.ChameleonBtn cmdProrrateoImportacion 
         Height          =   380
         Left            =   2160
         TabIndex        =   175
         Top             =   45
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "GENERAR PRORRATEO"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompra1.frx":12150
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdAplicarImportacion 
         Height          =   300
         Left            =   2160
         TabIndex        =   189
         Top             =   920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "APLICAR PRORRATEO"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompra1.frx":1216C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.I.F :"
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
         Left            =   390
         TabIndex        =   170
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.O.B :"
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
         Left            =   330
         TabIndex        =   169
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label38 
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
         Left            =   330
         TabIndex        =   168
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEGURO :"
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
         TabIndex        =   167
         Top             =   410
         Width           =   645
      End
   End
   Begin VB.CheckBox chkresponsable 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "RESPONSABLE:"
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
      Left            =   10920
      TabIndex        =   165
      Top             =   1725
      Width           =   1300
   End
   Begin VB.TextBox txttotal_final 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9960
      TabIndex        =   164
      Text            =   "0.00"
      Top             =   8040
      Visible         =   0   'False
      Width           =   375
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
      Left            =   15600
      MaxLength       =   80
      TabIndex        =   163
      Top             =   1660
      Width           =   855
   End
   Begin VB.Frame frmRetencion 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   270
      Left            =   12120
      TabIndex        =   159
      Top             =   600
      Visible         =   0   'False
      Width           =   4395
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
         TabIndex        =   160
         Top             =   20
         Width           =   2415
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdsave 
      Height          =   900
      Left            =   120
      TabIndex        =   157
      Top             =   8310
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "FrmCompra1.frx":12188
      PICN            =   "FrmCompra1.frx":121A4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkproyecto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "PROYECTO"
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
      Left            =   240
      TabIndex        =   136
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtvalidacion_chasis 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   135
      Text            =   "no"
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox chkConvertir 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "CONVERTIR"
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
      Height          =   255
      Left            =   9600
      TabIndex        =   110
      ToolTipText     =   "CONVERTIR A MONEDA NACIONAL"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VitekeySoft.ChameleonBtn CmdQuitar 
      Height          =   375
      Left            =   18000
      TabIndex        =   102
      ToolTipText     =   "ELIMINAR ITEM"
      Top             =   6240
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      MICON           =   "FrmCompra1.frx":157EC
      PICN            =   "FrmCompra1.frx":15808
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdActivar 
      Height          =   375
      Left            =   4560
      TabIndex        =   101
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "ACTIVAR"
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
      MICON           =   "FrmCompra1.frx":15DA2
      PICN            =   "FrmCompra1.frx":15DBE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtIdCompra 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   100
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox lblotros 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   92
      Text            =   "0.000"
      Top             =   8640
      Width           =   1110
   End
   Begin VB.TextBox lblExonerado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8400
      TabIndex        =   89
      Text            =   "0.00"
      Top             =   8550
      Width           =   1455
   End
   Begin VB.TextBox TxtValorventa 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   7425
      MaxLength       =   80
      TabIndex        =   88
      Text            =   "0.00"
      Top             =   7320
      Width           =   975
   End
   Begin VB.CheckBox ChkRecalcular 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Recalcular Precios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   15480
      TabIndex        =   86
      Top             =   7395
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtCosto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   9735
      MaxLength       =   80
      TabIndex        =   56
      Top             =   6300
      Width           =   735
   End
   Begin VB.TextBox TxtUnitario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   240
      MaxLength       =   80
      TabIndex        =   55
      Text            =   "0.00"
      Top             =   7320
      Width           =   975
   End
   Begin VB.TextBox TxtDireccion 
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
      MaxLength       =   80
      TabIndex        =   54
      Top             =   1750
      Width           =   5175
   End
   Begin VB.TextBox TxtCodProducto 
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
      Left            =   240
      MaxLength       =   80
      TabIndex        =   53
      Top             =   6300
      Width           =   1095
   End
   Begin VB.TextBox TxtDescripcionProducto 
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
      Left            =   5340
      MaxLength       =   80
      TabIndex        =   52
      Top             =   6300
      Width           =   4335
   End
   Begin VB.TextBox TxtSerie 
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
      Left            =   8180
      MaxLength       =   80
      TabIndex        =   51
      Top             =   720
      Width           =   855
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
      Left            =   9045
      MaxLength       =   80
      TabIndex        =   50
      Top             =   720
      Width           =   1815
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      MaxLength       =   80
      TabIndex        =   49
      Top             =   885
      Width           =   1215
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   48
      Top             =   1300
      Width           =   5175
   End
   Begin VB.TextBox TxtValoerNeto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   4040
      MaxLength       =   80
      TabIndex        =   47
      Text            =   "0.00"
      Top             =   7320
      Width           =   855
   End
   Begin VB.TextBox TxtISC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   5660
      MaxLength       =   80
      TabIndex        =   46
      Text            =   "0.00"
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox TxtIGV 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   8415
      MaxLength       =   80
      TabIndex        =   45
      Text            =   "0.00"
      Top             =   7320
      Width           =   975
   End
   Begin VB.TextBox TxtRetencion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   6800
      MaxLength       =   80
      TabIndex        =   44
      Text            =   "0.00"
      Top             =   7320
      Width           =   620
   End
   Begin VB.TextBox TxtPrecioVenta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   10035
      MaxLength       =   80
      TabIndex        =   43
      Text            =   "0.00"
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox TxtCostoAnt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   12840
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   42
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox TxtUtilidadAnt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   14640
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   41
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox txtPrecioVentaAnt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   15480
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   40
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox TxtCostoHoy 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   12840
      MaxLength       =   80
      TabIndex        =   39
      Top             =   7080
      Width           =   855
   End
   Begin VB.TextBox txtUtilidadhoy 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   14640
      MaxLength       =   80
      TabIndex        =   38
      Top             =   7080
      Width           =   735
   End
   Begin VB.TextBox TxtventaHoy 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   15480
      MaxLength       =   80
      TabIndex        =   37
      Top             =   7080
      Width           =   975
   End
   Begin VB.CheckBox chkModo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11760
      TabIndex        =   36
      Top             =   7320
      Width           =   255
   End
   Begin VB.TextBox txtIgv_Porcentaje 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   9405
      MaxLength       =   80
      TabIndex        =   35
      Text            =   "0.00"
      Top             =   7320
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DESCT (S/.)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   1320
      TabIndex        =   33
      Top             =   6960
      Width           =   975
      Begin VB.TextBox TxtDctoSoles 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
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
         Left            =   120
         MaxLength       =   80
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DESCT (%)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   2400
      TabIndex        =   31
      Top             =   6960
      Width           =   975
      Begin VB.TextBox TxtDstoporcentaje 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
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
         Left            =   120
         MaxLength       =   80
         TabIndex        =   32
         Text            =   "0.00"
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CANTIDADES"
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
      Height          =   700
      Left            =   1440
      TabIndex        =   28
      Top             =   5990
      Width           =   3735
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
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
         Left            =   50
         MaxLength       =   80
         TabIndex        =   30
         Top             =   320
         Width           =   650
      End
      Begin VB.TextBox TxtUnidades 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
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
         Left            =   1980
         MaxLength       =   80
         TabIndex        =   29
         Text            =   "1"
         Top             =   320
         Width           =   375
      End
      Begin MSDataListLib.DataCombo DtcUnidad 
         Height          =   315
         Left            =   720
         TabIndex        =   260
         Top             =   315
         Width           =   1240
         _ExtentX        =   2196
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
      Begin MSDataListLib.DataCombo DtcUnidadFinal 
         Height          =   315
         Left            =   2400
         TabIndex        =   261
         Top             =   315
         Visible         =   0   'False
         Width           =   1240
         _ExtentX        =   2196
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
      Begin VB.Label lblTrae 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRAE"
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
         Left            =   1170
         TabIndex        =   262
         Top             =   360
         Width           =   345
      End
   End
   Begin VB.TextBox lblIMPBruto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   27
      Text            =   "0.00"
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox lblDescuento 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   26
      Text            =   "0.00"
      Top             =   8230
      Width           =   1455
   End
   Begin VB.TextBox lblValorVenta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8400
      TabIndex        =   25
      Text            =   "0.00"
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox lblIgv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8400
      TabIndex        =   24
      Text            =   "0.00"
      Top             =   8230
      Width           =   1455
   End
   Begin VB.TextBox LblPercepcion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11040
      TabIndex        =   23
      Text            =   "0.00"
      Top             =   8325
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11040
      TabIndex        =   22
      Text            =   "0.00"
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CheckBox chkIGV 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11280
      TabIndex        =   21
      Top             =   7320
      Width           =   200
   End
   Begin VB.TextBox txtOtros 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   3400
      MaxLength       =   80
      TabIndex        =   20
      Text            =   "0.00"
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox TxtTotalDescuento 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   4900
      MaxLength       =   80
      TabIndex        =   19
      Text            =   "0.00"
      Top             =   7320
      Width           =   735
   End
   Begin VB.CheckBox ChkPercepcion 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "PERCEP"
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
      Left            =   2370
      TabIndex        =   18
      Top             =   7920
      Width           =   855
   End
   Begin VB.TextBox TxtPecepcion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   17
      Text            =   "0.000"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox txtRedondeo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   16
      Text            =   "0.000"
      Top             =   8280
      Width           =   1110
   End
   Begin VB.TextBox TxtObservacion 
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
      Left            =   13440
      MaxLength       =   2500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   2020
      Width           =   5055
   End
   Begin VB.TextBox TxtISC_p 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
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
      Left            =   6285
      MaxLength       =   80
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox Txtdoc_cod 
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtSeriaGuardada 
      Height          =   375
      Left            =   10440
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtNumeroGuardado 
      Height          =   375
      Left            =   12120
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtAlmacenGuardado 
      Height          =   375
      Left            =   13800
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox TxtProveedorGuardado 
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtTC 
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   8445
      MaxLength       =   5
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox lblISC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   8550
      Width           =   1455
   End
   Begin VB.CommandButton cmdActualizar 
      BackColor       =   &H008080FF&
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7395
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   330
      Left            =   210
      TabIndex        =   57
      Top             =   240
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   345
      Left            =   7410
      TabIndex        =   58
      Top             =   240
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      Style           =   2
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
   Begin MSDataListLib.DataCombo DtTipoCompra 
      Height          =   330
      Left            =   8445
      TabIndex        =   60
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSDataListLib.DataCombo DtcMoneda 
      Height          =   330
      Left            =   8445
      TabIndex        =   61
      Top             =   1920
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSComCtl2.DTPicker Dtpproducto_vencimiento 
      Height          =   300
      Left            =   16560
      TabIndex        =   96
      Top             =   6360
      Width           =   1245
      _ExtentX        =   2196
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
      Format          =   177209345
      CurrentDate     =   40963
   End
   Begin VitekeySoft.ChameleonBtn CmdAgregar 
      Height          =   375
      Left            =   18000
      TabIndex        =   103
      ToolTipText     =   "AGREGAR ITEM"
      Top             =   6720
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      MICON           =   "FrmCompra1.frx":183A3
      PICN            =   "FrmCompra1.frx":183BF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdRevertir 
      Height          =   375
      Left            =   18000
      TabIndex        =   104
      ToolTipText     =   "APLICAR CAMBIO"
      Top             =   7200
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      MICON           =   "FrmCompra1.frx":18959
      PICN            =   "FrmCompra1.frx":18975
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdIngresoSerie 
      Height          =   495
      Left            =   10560
      TabIndex        =   109
      ToolTipText     =   "INGRESO DE N� SERIES"
      Top             =   6255
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   ""
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      MICON           =   "FrmCompra1.frx":18F0F
      PICN            =   "FrmCompra1.frx":18F2B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtPorcentajeGastos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   240
      Left            =   13800
      MaxLength       =   80
      TabIndex        =   114
      ToolTipText     =   "% GASTOS ADMINISTRATIVOS"
      Top             =   7400
      Width           =   735
   End
   Begin VB.TextBox TxtGastoAdminHoy 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   13800
      MaxLength       =   80
      TabIndex        =   112
      Top             =   7080
      Width           =   735
   End
   Begin VB.TextBox TxtGastoAdminAnterior 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Left            =   13800
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   111
      Top             =   6600
      Width           =   735
   End
   Begin VitekeySoft.ChameleonBtn cmdDatosdua 
      Height          =   495
      Left            =   11400
      TabIndex        =   124
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "DATOS DUA"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCompra1.frx":1C052
      PICN            =   "FrmCompra1.frx":1C06E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   3015
      Left            =   120
      TabIndex        =   59
      Top             =   2640
      Width           =   18495
      _ExtentX        =   32623
      _ExtentY        =   5318
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      ForeColorSel    =   16777215
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
   Begin MSComCtl2.DTPicker dtpFechaRegistro 
      Height          =   315
      Left            =   12120
      TabIndex        =   127
      Top             =   290
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
      Format          =   177209345
      CurrentDate     =   40963
   End
   Begin MSDataListLib.DataCombo DtcTipo 
      Height          =   330
      Left            =   8445
      TabIndex        =   128
      Top             =   1480
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSDataListLib.DataCombo Dtcperiodo 
      Height          =   330
      Left            =   12300
      TabIndex        =   130
      Top             =   1320
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
   Begin MSMask.MaskEdBox TxtFecha_emision 
      Height          =   300
      Left            =   13800
      TabIndex        =   132
      ToolTipText     =   "dd/mm/yyyy"
      Top             =   290
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtfecha_Vencimiento 
      Height          =   300
      Left            =   15360
      TabIndex        =   133
      ToolTipText     =   "dd/mm/yyyy"
      Top             =   285
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSDataListLib.DataCombo DtcProyecto 
      Height          =   330
      Left            =   1440
      TabIndex        =   137
      Top             =   2160
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   33023
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
   Begin VitekeySoft.ChameleonBtn cmdEliminar 
      Height          =   1020
      Left            =   18960
      TabIndex        =   152
      Top             =   3360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1799
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
      MICON           =   "FrmCompra1.frx":1EF07
      PICN            =   "FrmCompra1.frx":1EF23
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   900
      Left            =   18960
      TabIndex        =   153
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "FrmCompra1.frx":2136D
      PICN            =   "FrmCompra1.frx":21389
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
      Height          =   900
      Left            =   18960
      TabIndex        =   154
      Top             =   1530
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
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
      MICON           =   "FrmCompra1.frx":217DB
      PICN            =   "FrmCompra1.frx":217F7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdBuscar 
      Height          =   1020
      Left            =   18960
      TabIndex        =   155
      Top             =   5500
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1799
      BTYPE           =   5
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCompra1.frx":21B11
      PICN            =   "FrmCompra1.frx":21B2D
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
      Height          =   1020
      Left            =   18960
      TabIndex        =   156
      Top             =   6600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1799
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
      MICON           =   "FrmCompra1.frx":21E47
      PICN            =   "FrmCompra1.frx":21E63
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdmodificar 
      Height          =   900
      Left            =   1200
      TabIndex        =   158
      Top             =   8310
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
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
      MICON           =   "FrmCompra1.frx":22253
      PICN            =   "FrmCompra1.frx":2226F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCompratir 
      Height          =   900
      Left            =   18960
      TabIndex        =   161
      Top             =   2460
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
      BTYPE           =   5
      TX              =   "COMPARTIR"
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
      MICON           =   "FrmCompra1.frx":248A8
      PICN            =   "FrmCompra1.frx":248C4
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
      Left            =   12300
      TabIndex        =   162
      Top             =   1665
      Width           =   3255
      _ExtentX        =   5741
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
   Begin VitekeySoft.ChameleonBtn cmdImprimir 
      Height          =   1020
      Left            =   18960
      TabIndex        =   208
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1799
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
      MICON           =   "FrmCompra1.frx":24D3A
      PICN            =   "FrmCompra1.frx":24D56
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcComprobante_orden 
      Height          =   315
      Left            =   3600
      TabIndex        =   256
      Top             =   885
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
   Begin MSComctlLib.ProgressBar progresbar_kardex 
      Height          =   225
      Left            =   120
      TabIndex        =   257
      Top             =   8060
      Width           =   2065
      _ExtentX        =   3651
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker DtpKardex 
      Height          =   315
      Left            =   15280
      TabIndex        =   258
      Top             =   1320
      Width           =   1220
      _ExtentX        =   2143
      _ExtentY        =   556
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
      Format          =   177209345
      CurrentDate     =   40963
   End
   Begin VB.Label lblAnulado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   3840
      TabIndex        =   81
      Top             =   4995
      Visible         =   0   'False
      Width           =   3270
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F.KARDEX:"
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
      Left            =   14430
      TabIndex        =   259
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N� AUTORIZACION :"
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
      Left            =   6900
      TabIndex        =   248
      Top             =   8940
      Width           =   1335
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COD-PRODUCTO"
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
      Index           =   8
      Left            =   220
      TabIndex        =   234
      Top             =   6000
      Width           =   1125
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.UNITARIO"
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
      Index           =   7
      Left            =   360
      TabIndex        =   233
      Top             =   6960
      Width           =   825
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL VENTA"
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
      Index           =   6
      Left            =   10230
      TabIndex        =   232
      Top             =   6960
      Width           =   915
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV (%)"
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
      Index           =   5
      Left            =   9510
      TabIndex        =   231
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV (S/.)"
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
      Index           =   4
      Left            =   8820
      TabIndex        =   230
      Top             =   6960
      Width           =   555
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "V.VENTA"
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
      Index           =   3
      Left            =   7605
      TabIndex        =   229
      Top             =   6960
      Width           =   585
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISC"
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
      Left            =   5985
      TabIndex        =   228
      Top             =   6960
      Width           =   225
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCT"
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
      Left            =   5220
      TabIndex        =   227
      Top             =   6960
      Width           =   435
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBS"
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
      Index           =   2
      Left            =   12225
      TabIndex        =   226
      Top             =   6960
      Width           =   285
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIV"
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
      Index           =   1
      Left            =   11760
      TabIndex        =   225
      Top             =   6960
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   735
      Index           =   2
      Left            =   12120
      Top             =   6915
      Width           =   420
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N� LOTE :"
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
      Index           =   5
      Left            =   16635
      TabIndex        =   215
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label lblidCompra 
      Alignment       =   2  'Center
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
      Height          =   280
      Left            =   120
      TabIndex        =   188
      Top             =   7770
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RETENCION:"
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
      Left            =   4425
      TabIndex        =   187
      Top             =   8950
      Width           =   825
   End
   Begin VB.Label Label60 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GLOSA :"
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
      Left            =   12810
      TabIndex        =   134
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label59 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "PERIODO CONT:"
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
      Height          =   300
      Left            =   10890
      TabIndex        =   131
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO SERVICIO :"
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
      Left            =   7125
      TabIndex        =   129
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F. VENCIMIENTO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   15030
      TabIndex        =   126
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G.[ADMIN]"
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
      Left            =   13800
      TabIndex        =   113
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENCIMIENTO:"
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
      Index           =   4
      Left            =   16560
      TabIndex        =   97
      Top             =   6120
      Width           =   1005
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION :"
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
      Left            =   390
      TabIndex        =   95
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAZON SOCIAL :"
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
      Left            =   120
      TabIndex        =   94
      Top             =   1395
      Width           =   1215
   End
   Begin VB.Label Label45 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/DNI :"
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
      Index           =   0
      Left            =   450
      TabIndex        =   93
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OTROS(FISE) :"
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
      Left            =   2295
      TabIndex        =   91
      Top             =   8760
      Width           =   915
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXONERADO  :"
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
      Left            =   7260
      TabIndex        =   90
      Top             =   8670
      Width           =   975
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F. REGISTRO "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   12105
      TabIndex        =   87
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label lblpercep_titulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL+PERC:"
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
      Left            =   10125
      TabIndex        =   85
      Top             =   8400
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IM.BRUTO :"
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
      Left            =   4425
      TabIndex        =   84
      Top             =   7920
      Width           =   795
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COSTO"
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
      Left            =   9705
      TabIndex        =   83
      Top             =   6000
      Width           =   465
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION PRODUCTO"
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
      Left            =   5340
      TabIndex        =   82
      Top             =   6000
      Width           =   1725
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F. EMISION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   13740
      TabIndex        =   79
      Top             =   60
      Width           =   885
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESC :"
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
      Left            =   4785
      TabIndex        =   78
      Top             =   8265
      Width           =   435
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR VENTA :"
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
      Left            =   7170
      TabIndex        =   77
      Top             =   7920
      Width           =   1065
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV  :"
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
      Left            =   7860
      TabIndex        =   76
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL  :"
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
      Left            =   10485
      TabIndex        =   75
      Top             =   7920
      Width           =   555
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IMP.BRUTO"
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
      Left            =   4185
      TabIndex        =   74
      Top             =   6960
      Width           =   795
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RET(8%)"
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
      Left            =   6870
      TabIndex        =   73
      Top             =   6960
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   4
      X1              =   12840
      X2              =   16440
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.VENTA"
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
      Index           =   3
      Left            =   15420
      TabIndex        =   72
      Top             =   6360
      Width           =   585
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UTIL[%]"
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
      Left            =   14640
      TabIndex        =   71
      Top             =   6360
      Width           =   525
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.COSTO"
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
      Left            =   12840
      TabIndex        =   70
      Top             =   6360
      Width           =   615
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV"
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
      Left            =   11280
      TabIndex        =   69
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OTROS"
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
      TabIndex        =   68
      Top             =   6960
      Width           =   465
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ROUND :"
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
      Left            =   2565
      TabIndex        =   67
      Top             =   8400
      Width           =   615
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISC(%)"
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
      Left            =   6405
      TabIndex        =   66
      Top             =   6960
      Width           =   435
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO COMPRA:"
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
      Left            =   7245
      TabIndex        =   65
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONEDA:"
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
      Left            =   7605
      TabIndex        =   64
      Top             =   1920
      Width           =   765
   End
   Begin VB.Label lblcambio 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T.CAMBIO:"
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
      Left            =   7545
      TabIndex        =   63
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISC :"
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
      Left            =   4905
      TabIndex        =   62
      Top             =   8640
      Width           =   315
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   855
      Left            =   195
      Top             =   5880
      Width           =   10335
   End
   Begin VB.Shape Shape10 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000080&
      BorderStyle     =   6  'Inside Solid
      Height          =   1815
      Left            =   12720
      Top             =   5880
      Width           =   5895
   End
   Begin VB.Label LblCantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   10560
      TabIndex        =   80
      Top             =   5880
      Width           =   495
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   660
      Left            =   120
      Top             =   45
      Width           =   6615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1875
      Left            =   120
      Top             =   735
      Width           =   6615
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   1380
      Left            =   2280
      Top             =   7800
      Width           =   17775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2595
      Left            =   6840
      Top             =   15
      Width           =   11775
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   735
      Index           =   0
      Left            =   11160
      Top             =   6915
      Width           =   420
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   735
      Index           =   1
      Left            =   11640
      Top             =   6915
      Width           =   420
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1980
      Left            =   120
      Top             =   5760
      Width           =   18495
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9225
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doc_Tienda As String * 1
Dim cod_doc As String
Dim RstDetCompra As New ADODB.Recordset
Dim rstTemporal As New ADODB.Recordset
Dim StrCodDetCompra As String
Dim StrCodReferencia As Double
Dim Referencia As Boolean
Public Procedencia As EnumProcede
Public ProcendenciaGuia As EnumGuia
Public ProcedenciaFactura As EnumFactura
Public codigoP As String
Dim precio_unit As Single
Public rever As Boolean
Dim DescuentoT As Single
Dim Pcosto As Single
Dim descuentoRecorrido As Single
Dim descuentoPersonal As Single
Dim descuentoViaticos As Single
Dim idOrden As Double


Public Sub get_unidad(ByVal in_producto As String, ByVal in_agranel As String)
    
    
    If in_agranel = "si" Then
        strCadena = "SELECT id_unidad as Codigo,descripcion as Descripcion FROM view_unidad_producto WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    Else
        strCadena = "SELECT id_unidad as Codigo,descripcion as Descripcion FROM view_unidad WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    End If
    
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcUnidad)
    
    
    
End Sub

Public Sub get_unidad_final(ByVal in_producto As String, ByVal in_unidad As String)
   
   
    
        strCadena = "SELECT id_unidad as Codigo,descripcion as Descripcion FROM view_unidad_producto WHERE id_producto='" & in_producto & "' and id_unidad='" & in_unidad & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstT(strCadena)
        Call LlenaDataComboT(Me.DtcUnidadFinal)
        
        If Me.DtcUnidad.BoundText = Me.DtcUnidadFinal.BoundText Then
            Me.DtcUnidadFinal.Visible = False
        End If
   
    
    
End Sub



Public Sub LlenarGastos(ByVal Grilla As MSHFlexGrid, ByVal cCompra As Double)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT     DocumentoCompraGastos.IdGasto, (Comprobantes.doc_abrev+':'+DocumentoCompraGastos.serie+'-'+ DocumentoCompraGastos.numero) as Numero, " & _
" DocumentoCompraGastos.Monto , DocumentoCompraGastos.detalle FROM DocumentoCompraGastos INNER JOIN " & _
" Comprobantes ON DocumentoCompraGastos.doc_cod = Comprobantes.doc_cod  WHERE idCompra='" & cCompra & "'"
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
            Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 2500
           Grilla.ColWidth(2) = 2300
           Grilla.ColWidth(3) = 2900
                     
         Next
        cabecera = "ID" & vbTab & "COMPROBANTE" & vbTab & "DETALLE" & vbTab & "MONTO"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("IdGasto") & vbTab & rst("Numero") & vbTab & rst("detalle") & vbTab & Format(rst("Monto"), "###0.00")
            Grilla.AddItem Fila
            tTotal = tTotal + rst("Monto")
            Fila = ""
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & Format(tTotal, "###0.00")
      Grilla.AddItem Fila
       For k = 5 To 3
            Grilla.col = 5
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
 
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub
Private Sub AgregarGrilla()
If Val(Me.txtCantidad.Text) > 0 And Trim(Me.TxtCodProducto.Text) <> "" Then
    Call AgregarTemporal
    'Me.TxtCodProducto.Text = "0000"
    'Me.TxtDescripcionProducto.Text = ""
    'Me.TxtCantidad.Text = "0"
     
    Call Resalta(Me.TxtCodProducto)
Else
    Call Resalta(Me.txtCantidad)
End If
End Sub
Private Sub AgregarTemporal()
    strCadena = "SELECT cProducto,Cantidad FROM Temporal_Compras WHERE cProducto='" & Trim(codigoP) & "' AND cDocumentoCompra='" & Trim(Me.TxtNumeroDoc.Text) & "'"
    Call ConfiguraRst(strCadena)
    'Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
    Me.cmdsave.Enabled = True
    Call AgregarNuevo
End Sub
Private Sub get_glosa_servicio(ByVal in_producto As String)
strCadena = "SELECT * FROM producto WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   If Val(rst("cta_contable")) > 0 Then
        Me.txtObservacion.Text = Trim(Me.TxtDescripcionProducto.Text)
        Me.chkresponsable.Value = 1
   
       
   End If
End If
End Sub
Sub AgregarNuevo()
Dim exonerado As Double, valor_venta As Double, igv As Double
    If Val(Me.txtprecioventa.Text) = 0 Then
        Me.ChkRecalcular.Visible = True
    Else
        Me.ChkRecalcular.Value = 0
        Me.ChkRecalcular.Visible = False
    End If
    
    If Me.DtTipoCompra.BoundText = "01" Then
        valor_venta = Val(Me.txtprecioventa.Text)
        igv = 0
        in_total = Val(Me.txtprecioventa.Text)
    Else
        If KEY_CON_IGV = "si" And Me.chkigv.Value = 1 Then
            
            If chk_valor_venta = 1 Then
                exonerado = 0
                in_total = Val(Me.txtprecioventa.Text) '* (1 + KEY_IGV)
                valor_venta = Val(Me.txtValorVenta.Text)    'Val(Me.txtprecioventa.Text)
                igv = (in_total - valor_venta)
                'valor_venta = Val(Me.TxtValorventa.Text)
            Else
                exonerado = 0
                in_total = Val(Me.txtprecioventa.Text)
                igv = Val(Me.TxtIgv.Text)
                valor_venta = Val(Me.txtValorVenta.Text)
            End If
            
            
            
        Else
            
           If KEY_PAIS = KEY_PERU Then
            
                    If Me.chkigv.Value = 1 Then
                        exonerado = 0
                        igv = Val(Me.TxtIgv.Text)
                        valor_venta = Val(Me.txtValorVenta.Text)
                        in_total = Val(Me.txtprecioventa.Text)
                    Else
                        exonerado = Val(Me.txtprecioventa.Text)
                        igv = 0
                        valor_venta = 0
                        in_total = Val(Me.txtprecioventa.Text)
                    End If
           Else
                If Me.chkigv.Value = 1 Then
                        exonerado = 0
                        igv = Val(Me.TxtIgv.Text)
                        valor_venta = Val(Me.txtValorVenta.Text)
                        in_total = Val(Me.txtprecioventa.Text)
                    Else
                        exonerado = Val(Me.txtprecioventa.Text)
                        igv = 0
                        valor_venta = Val(Me.txtprecioventa.Text)
                        in_total = Val(Me.txtprecioventa.Text)
                    End If
            
           End If
        End If
        
        
        
        
    End If
    
    
    
    
    
    If Me.DtcTipoDoc.BoundText = "0050" Then
        exonerado = 0
        igv = Val(Me.txtprecioventa.Text)
        valor_venta = 0
    End If
    
    If Me.DtcTipoDoc.BoundText = "0002" Then
        exonerado = 0
        igv = 0
        valor_venta = Val(Me.txtprecioventa.Text)
    End If
    
    If Me.DtcTipoDoc.BoundText = "0200" Then
        exonerado = 0
        igv = 0
        valor_venta = Val(Me.txtprecioventa.Text)
    End If
    
    
    Call get_glosa_servicio(Trim(Me.TxtCodProducto.Text))
    
    If Me.chk_obsequio.Value = 1 Then
        in_obsequio = "si"
    Else
        in_obsequio = "no"
    End If
    
    If KEY_PAIS = KEY_PERU Then
        in_cuenta_contable = "0"
        in_porcentaje_retencion = 0
    Else
        If Me.DtcTipoDoc.BoundText = "0020" Or Me.DtcTipoDoc.BoundText = "0427" Then
            in_cuenta_contable = Trim(Me.Label21(17).Caption)
            in_porcentaje_retencion = Val(Label21(16).Caption)
            
        Else
            in_cuenta_contable = "0"
            in_porcentaje_retencion = 0
        End If
    End If
    
    strCadena = "INSERT INTO movimiento_compra_temporal(id_doc,serie,numero,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento," & _
    "valor_neto,isc,igv,retencion,otros,percepcion,exonerado,valor_venta,precio_venta,p_venta,p_costo,dni_save,id_alm,detalle,fecha_vencimiento,numero_lote,obsequio,id_unidad,cuenta_contable,porcentaje_retencion,ruc) VALUES " & _
    "('" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Trim(Me.txtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(codigoP) & "'," & _
    "'" & Val(Me.txtCantidad.Text) & "','" & Val(Me.TxtUnitario.Text) & "','" & Val(Me.TxtDctoSoles.Text) & "'," & _
    "'" & Val(Me.TxtDstoporcentaje.Text) & "','" & Val(Me.TxtTotalDescuento.Text) & "','" & Val(Me.TxtValoerNeto.Text) & "','" & Val(Me.txtisc.Text) & "','" & igv & "'," & _
    "'" & Val(Me.txtRetencion.Text) & "','" & Val(Me.txtOtros.Text) & "','" & Val(percepcion) & "','" & exonerado & "','" & valor_venta & "','" & Val(in_total) & "','" & Val(Me.TxtventaHoy.Text) & "'," & _
    "'" & Val(Me.TxtCostoHoy.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.DtcAlmacen.BoundText) & "','" & Trim(Me.TxtDescripcionProducto.Text) & "','" & Format(Me.Dtpproducto_vencimiento.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtLote.Text) & "','" & in_obsequio & "','" & Me.DtcUnidad.BoundText & "','" & in_cuenta_contable & "','" & in_porcentaje_retencion & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
     
     
     
    strCadena = "UPDATE producto SET vencimiento='" & Format(Me.Dtpproducto_vencimiento.Value, "YYYY-mm-dd") & "' WHERE id_producto='" & Trim(codigoP) & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    
    Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcTipoDoc.BoundText, Me.txtSerie.Text)
    Me.TxtCodProducto.Text = ""
    Me.TxtCodProducto.Text = ""
    Me.txtCantidad.Text = 0
    Me.TxtDescripcionProducto.Text = ""
    Me.DtcUnidad.BoundText = 0
    Me.txtcosto.Text = 0
    Me.TxtUnitario.Text = 0
    Me.TxtDctoSoles.Text = 0
    Me.TxtDstoporcentaje.Text = 0
    Me.txtOtros.Text = 0
    Me.TxtValoerNeto.Text = 0
    Me.TxtTotalDescuento.Text = 0
    Me.txtisc.Text = 0
    Me.TxtISC_p.Text = 0
    Me.txtRetencion.Text = 0
    Me.txtValorVenta.Text = 0
    Me.TxtIgv.Text = 0
    Me.txtIgv_Porcentaje.Text = 0
    Me.txtprecioventa.Text = 0
    Me.txtPrecioVentaAnt.Text = 0
    Me.TxtCostoAnt.Text = 0
    Me.TxtCostoHoy.Text = 0
    Me.TxtUtilidadAnt.Text = 0
    Me.txtUtilidadhoy.Text = 0
    Me.TxtventaHoy.Text = 0
    Me.TxtCostoHoy.Text = 0
    Me.chkigv.Value = 0
    Me.chkModo.Value = 0
    Me.chk_obsequio.Value = 0
    Me.TxtUnidades.Text = 1
    
End Sub
Private Sub VerificaDocumento(ByVal TipoDoc As String)
If Trim(Me.DtcTipoDoc.BoundText) = "0009" Then
    'Me.TlbGrabar.Buttons(KEY_GUIAREMISION).Enabled = True
End If
End Sub
Sub ModificarCantidad(ByVal Can_Previa As Integer)
    Dim Can_actual As Integer
'    Can_actual = Can_Previa + Val(Me.TxtCantidad.Text)
 '   StrCadena = "UPDATE Temporal_Compras SET cantidad='" & Can_actual & "',Total='" & Can_actual * Val(Me.TxtPrecio.Text) & "' WHERE cProducto='" & Trim(Me.TxtCodProducto.Text) & "' "
  '  Call EjecutaRST(StrCadena)
   ' Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcTipoDoc.BoundText, Me.TxtSerie.Text)
End Sub

Function GeneraCodTemporal() As Integer
Dim Codtemporal As Integer
Call ConfiguraRst(strCadena)
    If rst.EOF = True Then
        Codtemporal = 1
    Else
        Codtemporal = rst(0) + 1
    End If
  GeneraCodTemporal = Codtemporal
  Set rst = Nothing
End Function
Function GeneraCodTemporalCompras() As Integer
Dim Codtemporal As Integer
Call ConfiguraRst(strCadena)
    If rst.EOF = True Then
        Codtemporal = 1
    Else
        Codtemporal = rst(0) + 1
    End If
  GeneraCodTemporalCompras = Codtemporal
  Set rst = Nothing
End Function
Function GeneraCodReferencia() As Integer
Dim CodReferencia As Integer
strCadena = "SELECT IdReferencia FROM DocReferencia_Compra ORDER BY IdReferencia DESC "
Call ConfiguraRst(strCadena)
    If rst.EOF = True Then
        CodReferencia = 1
        
    Else
        CodReferencia = rst(0) + 1

    End If
  GeneraCodReferencia = CodReferencia
  
  
  Set rst = Nothing
End Function
Sub llenarGrid_det(ByVal Grilla As MSHFlexGrid, ByVal numero As String, ByVal TipoDoc As String, ByVal serie As String)
On Error GoTo salir
Dim Total As Double, SUBTOTAL As Double, igv As Single, tpercepcion As Single
Dim in_pventa As Double
Dim in_pcosto As Double

strCadena = "SELECT * FROM view_temporal_compra_ii WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
   
   Grilla.Rows = 0
  ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 700
            Grilla.ColWidth(2) = 4300
            Grilla.ColWidth(3) = 500
            Grilla.ColWidth(4) = 500
            Grilla.ColWidth(5) = 850
            Grilla.ColWidth(6) = 850
            Grilla.ColWidth(7) = 800
            Grilla.ColWidth(8) = 600
            Grilla.ColWidth(9) = 800
            Grilla.ColWidth(10) = 800
            Grilla.ColWidth(11) = 800
            Grilla.ColWidth(12) = 800
            Grilla.ColWidth(13) = 700
            Grilla.ColWidth(14) = 700
            Grilla.ColWidth(15) = 800
            Grilla.ColWidth(16) = 1100
            Grilla.ColWidth(17) = 1000
            Grilla.ColWidth(18) = 900
            Grilla.ColWidth(19) = 1100
            
    Next
  
 
             
             Fila = "IDTEMPORAL" & vbTab & "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "UND" & vbTab & "CANT" & vbTab & "P.UNIT" & vbTab & "V.NETO" & vbTab & "OTROS" & vbTab & "T.DSC" & vbTab & "ISC" & vbTab & "RET(8%)" & vbTab & "P.VENTA" & vbTab & "P.COSTO" & vbTab & "INC[NETO]" & vbTab & "INC[%]" & vbTab & "G.VINCU" & vbTab & "VALOR VENTA" & vbTab & "IGV" & vbTab & "EXONERADO" & vbTab & "TOTAL"
             Grilla.AddItem Fila
             For k = 1 To 19
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
             Fila = ""
             cantidad = 0
             unit = 0
             neto = 0
             desS = 0
             desP = 0
             Otros = 0
             Tdes = 0
             isc = 0
             igv = 0
             retencion = 0
             pventa = 0
             valor_venta = 0
             in_exonerado = 0
             in_pventa = 0
             in_pcosto = 0
             in_incremento = 0
             in_valor_neto = 0
             in_incremento_neto = 0
             incremento_neto_gasto = 0
        For i = 0 To rst.RecordCount - 1

             
             Fila = rst("id_temporal") & vbTab & rst("id_producto") & vbTab & UCase(rst("nombre_prod")) & vbTab & rst("abreviatura") & vbTab & rst("cantidad") & vbTab & Format(rst("c_unitario"), "###0.00") & vbTab & Format(rst("valor_neto"), "###0.00") & vbTab & Format(rst("otros"), "###0.00") & vbTab & Format(rst("total_descuento"), "###0.00") & vbTab & Format(rst("isc"), "###0.00") & vbTab & Format(rst("retencion"), "###0.00") & vbTab & Format(rst("p_venta"), "###0.00") & vbTab & Format(rst("p_costo"), "###0.00") & vbTab & Format(rst("incremento_neto"), "###0.00") & vbTab & Format(rst("incremento"), "###0.00") & " %" & vbTab & Format(rst("incremento_neto_gasto"), "###0.00") & vbTab & Format(rst("valor_venta"), "###0.00") & vbTab & Format(rst("igv"), "###0.00") & vbTab & Format(rst("exonerado"), "###0.00") & vbTab & Format(rst("precio_venta"), "###0.00")
             Grilla.AddItem Fila
           
                                Grilla.col = 19
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
           
           
             cantidad = cantidad + rst("cantidad")
             unit = unit + rst("c_unitario")
             neto = neto + rst("valor_neto")
             desS = desS + rst("dsto_soles")
             desP = desP + rst("dsto_procentaje")
             Otros = Otros + rst("otros")
             Tdes = Tdes + rst("total_descuento")
             isc = isc + rst("isc")
             igv = igv + rst("igv")
             retencion = retencion + rst("retencion")
             valor_venta = valor_venta + rst("valor_venta")
             in_exonerado = in_exonerado + rst("exonerado")
             pventa = pventa + rst("precio_venta")
             in_pventa = in_pventa + rst("p_venta")
             in_pcosto = in_pcosto + rst("p_costo")
             in_valor_neto = in_valor_neto + rst("valor_neto")
             in_incremento = in_incremento + rst("incremento")
             in_incremento_neto = in_incremento_neto + rst("incremento_neto")
             incremento_neto_gasto = incremento_neto_gasto + rst("incremento_neto_gasto")
             
             rst.MoveNext
        Next i

         Fila = "" & vbTab & "" & vbTab & " [ ::::::::::        T  O  T  A  L  E  S        ::::::::::::: ] " & vbTab & "" & vbTab & Format(cantidad, "###0.00") & vbTab & Format(unit, "###0.00") & vbTab & Format(neto, "###0.00") & vbTab & Format(Otros, "###0.00") & vbTab & Format(Tdes, "###0.00") & vbTab & Format(isc, "###0.00") & vbTab & Format(retencion, "###0.00") & vbTab & Format(in_pventa, "###0.00") & vbTab & Format(in_pcosto, "###0.00") & vbTab & Format(in_incremento_neto, "###0.00") & vbTab & Format(in_incremento, "###0.00") & " %" & vbTab & Format(incremento_neto_gasto, "###0.00") & vbTab & Format(valor_venta, "###0.00") & vbTab & Format(igv, "###0.00") & vbTab & Format(in_exonerado, "###0.00") & vbTab & Format(pventa, "###0.00")
         Grilla.AddItem Fila
                        For k = 10 To 19
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
                            Next k
Me.lblCantidad.Caption = Trim(rst.RecordCount)

Call Resalta(Me.TxtCodProducto)


If Me.ChkPercepcion.Value = 1 Then
    Me.lblPercepcion.Text = Format(Val(Me.TxtPecepcion.Text), "###0.00")
    tpercepcion = Me.lblPercepcion.Text
Else
    tpercepcion = 0
    Me.lblPercepcion.Text = "0.00"
End If

Me.lblIMPBruto.Text = Format(in_valor_neto, "###0.00")
Me.lblDescuento.Text = Format(Tdes, "###0.00")
Me.lblExonerado.Text = Format(in_exonerado, "###0.00")
Me.LblValorVenta.Text = Format(valor_venta, "###0.00")
Me.LblIgv.Text = Format(igv, "###0.00")
Me.lblTotal.Text = Format(pventa, "###0.00")
Me.txttotal_final.Text = Val(Me.lblTotal.Text)
Me.lblISC.Text = Format(isc, "###0.00")
Me.TxtTotalRetencion.Text = Format(retencion, "###0.00")
Me.lblPercepcion.Text = Format(Val(Me.TxtPecepcion.Text) + Val(Me.lblTotal.Text), "###0.00")
Me.lblotros.Text = Format(Otros, "###0.00")
  
  Me.cmdAnular.Enabled = False
  Me.cmdEliminar.Enabled = False
  Me.cmdSalir.Enabled = True
  Me.cmdsave.Enabled = True
  'Grilla.Row = 1
  'Grilla.col = 0
  'Grilla.ColSel = 1
  'Grilla.RowSel = 1
  Exit Sub

salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Sub llenarGrid_revertir(ByVal Grilla As MSHFlexGrid, ByVal numero As String, ByVal TipoDoc As String, ByVal serie As String)
On Error GoTo salir
Dim Total As Double, SUBTOTAL As Double, igv As Single, tpercepcion As Single
Dim in_pventa As Double
Dim in_pcosto As Double

strCadena = "SELECT * FROM view_temporal_compra WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
   
   Grilla.Rows = 0
  ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 700
            Grilla.ColWidth(2) = 4000
            Grilla.ColWidth(3) = 700
            Grilla.ColWidth(4) = 700
            Grilla.ColWidth(5) = 1000
            Grilla.ColWidth(6) = 1000
            Grilla.ColWidth(7) = 800
            Grilla.ColWidth(8) = 600
            Grilla.ColWidth(9) = 900
            Grilla.ColWidth(10) = 1000
            Grilla.ColWidth(11) = 1000
            Grilla.ColWidth(12) = 1000
            Grilla.ColWidth(13) = 1000
            Grilla.ColWidth(14) = 1200
            Grilla.ColWidth(15) = 1000
            Grilla.ColWidth(16) = 1200
            Grilla.ColWidth(17) = 700
    Next
  
 
             
             Fila = "IDTEMPORAL" & vbTab & "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "UND" & vbTab & "CANT" & vbTab & "P.UNIT" & vbTab & "V.NETO" & vbTab & "OTROS" & vbTab & "T.DSC" & vbTab & "ISC" & vbTab & "RETEN" & vbTab & "P.VENTA" & vbTab & "P.COSTO" & vbTab & "INC[%]" & vbTab & "VALOR VENTA" & vbTab & "IGV" & vbTab & "TOTAL" & vbTab & "ESTADO"
             Grilla.AddItem Fila
             For k = 1 To 17
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
             
             cantidad = 0
             unit = 0
             neto = 0
             desS = 0
             desP = 0
             Otros = 0
             Tdes = 0
             isc = 0
             igv = 0
             retencion = 0
             pventa = 0
             valor_venta = 0
             in_exonerado = 0
             in_pventa = 0
             in_pcosto = 0
             in_incremento = 0
             in_valor_neto = 0
        For i = 0 To rst.RecordCount - 1

             
             
             If rst("seleccionado") = "si" Then
                estado = Chr(254)
            Else
                estado = Chr(168)
            End If
            
             
             
             Fila = rst("id_temporal") & vbTab & rst("id_producto") & vbTab & UCase(rst("nombre_prod")) & Space(2) & rst("color") & vbTab & rst("abreviatura") & vbTab & rst("cantidad") & vbTab & Format(rst("c_unitario"), "###0.00") & vbTab & Format(rst("valor_neto"), "###0.00") & vbTab & Format(rst("otros"), "###0.00") & vbTab & Format(rst("total_descuento"), "###0.00") & vbTab & Format(rst("isc"), "###0.00") & vbTab & Format(rst("retencion"), "###0.00") & vbTab & Format(rst("p_venta"), "###0.00") & vbTab & Format(rst("p_costo"), "###0.00") & vbTab & Format(rst("incremento"), "###0.00") & " %" & vbTab & Format(rst("valor_venta"), "###0.00") & vbTab & Format(rst("igv"), "###0.00") & vbTab & Format(rst("precio_venta"), "###0.00") & vbTab & estado
             Grilla.AddItem Fila
           
                                Grilla.col = 16
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
            
            With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 5 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
            End With
      
            
            
            If rst("seleccionado") = "si" Then
                For k = 10 To 17
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
                Next k
            End If
            
           
           
             cantidad = cantidad + rst("cantidad")
             unit = unit + rst("c_unitario")
             neto = neto + rst("valor_neto")
             desS = desS + rst("dsto_soles")
             desP = desP + rst("dsto_procentaje")
             Otros = Otros + rst("otros")
             Tdes = Tdes + rst("total_descuento")
             isc = isc + rst("isc")
             igv = igv + rst("igv")
             retencion = retencion + rst("retencion")
             valor_venta = valor_venta + rst("valor_venta")
             in_exonerado = in_exonerado + rst("exonerado")
             pventa = pventa + rst("precio_venta")
             in_pventa = in_pventa + rst("p_venta")
             in_pcosto = in_pcosto + rst("p_costo")
             in_valor_neto = in_valor_neto + rst("valor_neto")
             in_incremento = in_incremento + rst("incremento")
             Fila = ""
             rst.MoveNext
        Next i

         Fila = "" & vbTab & "" & vbTab & " [ ::::::::::        T  O  T  A  L  E  S        ::::::::::::: ] " & vbTab & "" & vbTab & Format(cantidad, "###0.00") & vbTab & Format(unit, "###0.00") & vbTab & Format(neto, "###0.00") & vbTab & Format(Otros, "###0.00") & vbTab & Format(Tdes, "###0.00") & vbTab & Format(isc, "###0.00") & vbTab & Format(retencion, "###0.00") & vbTab & Format(in_pventa, "###0.00") & vbTab & Format(in_pcosto, "###0.00") & vbTab & Format(in_incremento, "###0.00") & " %" & vbTab & Format(valor_venta, "###0.00") & vbTab & Format(igv, "###0.00") & vbTab & Format(pventa, "###0.00")
         Grilla.AddItem Fila
                        For k = 0 To 16
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
                            Next k
Me.lblCantidad.Caption = Trim(rst.RecordCount)

'Me.TxtCantidad.Text = 0

'strCadena = "SELECT sum(valor_neto) as valorneto,sum(valor_venta) as valorventa,sum(total_descuento) as descuento,sum(igv) as igv,sum(exonerado) as exonerado,sum(precio_venta) as precioventa,sum(isc) as isc,sum(otros) as otros FROM movimiento_compra_temporal WHERE id_alm='" & KEY_ALM & "' and  dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
'Call ConfiguraRst(strCadena)


If Me.ChkPercepcion.Value = 1 Then
    Me.lblPercepcion.Text = Format(Val(Me.TxtPecepcion.Text), "###0.00")
    tpercepcion = Me.lblPercepcion.Text
Else
    tpercepcion = 0
    Me.lblPercepcion.Text = "0.00"
End If

Me.lblIMPBruto.Text = Format(in_valor_neto, "###0.00")
Me.lblDescuento.Text = Format(Tdes, "###0.00")
Me.lblExonerado.Text = Format(in_exonerado, "###0.00")
Me.LblValorVenta.Text = Format(valor_venta, "###0.00")
Me.LblIgv.Text = Format(igv, "###0.00")
Me.lblTotal.Text = Format(pventa, "###0.00")
Me.lblISC.Text = Format(isc, "###0.00")
Me.lblPercepcion.Text = Format(Val(Me.TxtPecepcion.Text) + Val(Me.lblTotal.Text), "###0.00")
Me.lblotros.Text = Format(Otros, "###0.00")
  
  Me.cmdAnular.Enabled = False
  Me.cmdEliminar.Enabled = False
  Me.cmdSalir.Enabled = True
  Me.cmdsave.Enabled = True
  'Grilla.Row = 1
  'Grilla.col = 0
  'Grilla.ColSel = 1
  'Grilla.RowSel = 1
  Exit Sub

salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Sub llenarGrid_Comprobante(ByVal Grilla As MSHFlexGrid, ByVal id_compra As Double)
 On Error GoTo salir
Dim Total As Double, SUBTOTAL As Double, igv As Single, tpercepcion As Single
Dim prorrateo_imp As String
Dim prorrateo_gas As String


strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & id_compra & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   prorrateo_imp = rst("prorrateo_importacion")
   prorrateo_gas = rst("prorrateo_gastos")
   
Else
  prorrateo_imp = "no"
  prorrateo_gas = "no"
End If

strCadena = "SELECT * FROM view_detalle_compra WHERE id_compra='" & id_compra & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & id_compra & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.lblExonerado.Text = Format(rst("exonerado"), "###0.00")
        Me.LblValorVenta.Text = Format(rst("valor_venta"), "###0.00")
        Me.lblISC.Text = Format(rst("isc"), "###0.00")
        Me.LblIgv.Text = Format(rst("igv"), "###0.00")
        
        
       
        Me.lblTotal.Text = Format(rst("total"), "###0.00")
        Me.txttotal_final.Text = Val(Me.lblTotal.Text)
        Me.lblPercepcion.Text = Format(rst("percepcion"), "###0.00")
        Me.txtObservacion.Text = rst("observacion")
        
    End If
    Grilla.Rows = 0
    Exit Sub
End If
   
  Grilla.Rows = 0
  ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 700
            Grilla.ColWidth(2) = 3500
            Grilla.ColWidth(3) = 500
            Grilla.ColWidth(4) = 500
            Grilla.ColWidth(5) = 850
            Grilla.ColWidth(6) = 850
            Grilla.ColWidth(7) = 600
            Grilla.ColWidth(8) = 600
            Grilla.ColWidth(9) = 600
            Grilla.ColWidth(10) = 600
            Grilla.ColWidth(11) = 900
            Grilla.ColWidth(12) = 900
            Grilla.ColWidth(13) = 900
            Grilla.ColWidth(14) = 1000
            Grilla.ColWidth(15) = 1000
            Grilla.ColWidth(16) = 1000
            Grilla.ColWidth(17) = 1000
            Grilla.ColWidth(18) = 850
            Grilla.ColWidth(19) = 1200
    Next
  
 
             
             Fila = "IDTEMPORAL" & vbTab & "CODIGO" & vbTab & "DETALLE PRODUCTO/SERVICIO" & vbTab & "UND" & vbTab & "CANT" & vbTab & "P.UNIT" & vbTab & "V.NETO" & vbTab & "OTROS" & vbTab & "T.DSC" & vbTab & "ISC" & vbTab & "IVAP" & vbTab & "P.VENTA" & vbTab & "P.COSTO" & vbTab & "INC [NETO]" & vbTab & "INC [%]" & vbTab & "GAS.VIN" & vbTab & "[%GASTOS]" & vbTab & "VALOR VENTA" & vbTab & "IGV" & vbTab & "TOTAL"
             Grilla.AddItem Fila
             For k = 0 To 19
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
             Fila = ""
             cantidad = 0
             unit = 0
             neto = 0
             desS = 0
             desP = 0
             Otros = 0
             Tdes = 0
             isc = 0
             igv = 0
             ivap = 0
             pventa = 0
             valor_venta = 0
             in_exonerado = 0
             in_incremento_neto = 0
             in_incremento_neto_gasto = 0
             in_acumulado = 0
             in_retencion = 0
        For i = 0 To rst.RecordCount - 1

             If prorrateo_imp = "si" Then
                pventa = pventa + rst("precio_venta") + rst("incremento_neto")
                in_venta = rst("precio_venta") + rst("incremento_neto")
             Else
                pventa = pventa + rst("precio_venta")
                in_venta = rst("precio_venta")
             End If
             
             If prorrateo_gas = "si" Then
                pventa = pventa + rst("incremento_neto_gasto")
                in_venta = rst("precio_venta") + rst("incremento_neto_gasto") + rst("incremento_neto")
             Else
                pventa = pventa
                in_venta = in_venta
             End If
             
             Fila = rst("id_detalle_compra") & vbTab & rst("id_producto") & vbTab & UCase(rst("detalle")) & vbTab & rst("abreviatura") & vbTab & rst("cantidad") & vbTab & Format(rst("c_unitario"), "###0.00") & vbTab & Format(rst("valor_neto"), "###0.00") & vbTab & Format(rst("otros"), "###0.00") & vbTab & Format(rst("total_descuento"), "###0.00") & vbTab & Format(rst("isc"), "###0.00") & vbTab & Format(rst("ivap"), "###0.00") & vbTab & Format(rst("p_venta"), "###0.00") & vbTab & Format(rst("p_costo"), "###0.00") & vbTab & Format(rst("incremento_neto"), "###0.000") & vbTab & Format(rst("incremento"), "###0.000000") & vbTab & Format(rst("incremento_neto_gasto"), "###0.000") & vbTab & Format(rst("incremento_gasto"), "###0.000") & vbTab & Format(rst("valor_venta"), "###0.00") & vbTab & Format(rst("igv"), "###0.00") & vbTab & Format(in_venta, "###0.00")
             Grilla.AddItem Fila
              
                                Grilla.col = 19
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
                           
           
             cantidad = cantidad + rst("cantidad")
             unit = unit + rst("c_unitario")
             neto = neto + rst("valor_neto")
             desS = desS + rst("dsto_soles")
             desP = desP + rst("dsto_procentaje")
             Otros = Otros + rst("otros")
             Tdes = Tdes + rst("total_descuento")
             isc = isc + rst("isc")
             igv = igv + rst("igv")
             in_retencion = in_retencion + rst("retencion")
             ivap = ivap + rst("ivap")
             in_pventa = in_pventa + rst("p_venta")
             in_pcosto = in_pcosto + rst("p_costo")
             in_acumulado = in_acumulado + in_venta
             in_exonerado = in_exonerado + rst("exonerado")
             valor_venta = valor_venta + rst("valor_venta")
             in_incremento = in_incremento + rst("incremento")
             in_incremento_neto = in_incremento_neto + rst("incremento_neto")
             in_incremento_neto_gasto = in_incremento_neto_gasto + rst("incremento_neto_gasto")
             in_incremento_neto_gasto_porcentaje = in_incremento_neto_gasto_porcentaje + rst("incremento_gasto")
             rst.MoveNext
        Next i

         Fila = "" & vbTab & "" & vbTab & " [      :::::::::::::::: T  O  T  A  L  E  S ::::::::::::::::: ]" & vbTab & "" & vbTab & Format(cantidad, "###0.00") & vbTab & Format(unit, "###0.00") & vbTab & Format(neto, "###0.00") & vbTab & Format(Otros, "###0.00") & vbTab & Format(Tdes, "###0.00") & vbTab & Format(isc, "###0.00") & vbTab & Format(ivap, "###0.00") & vbTab & Format(in_pventa, "###0.00") & vbTab & Format(in_pcosto, "###0.00") & vbTab & Format(in_incremento_neto, "###0.000") & vbTab & Format(in_incremento, "###0.000") & vbTab & Format(in_incremento_neto_gasto, "###0.000") & vbTab & Format(in_incremento_neto_gasto_porcentaje, "###0.000") & vbTab & Format(valor_venta, "###0.00") & vbTab & Format(igv, "###0.00") & vbTab & Format(in_acumulado, "###0.00")
         Grilla.AddItem Fila
                        For k = 0 To 19
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
                            Next k
Me.lblCantidad.Caption = Trim(rst.RecordCount)
If Me.ChkPercepcion.Value = 1 Then
    Me.lblPercepcion.Text = Format(Val(Me.TxtPecepcion.Text), "###0.00")
    tpercepcion = Me.lblPercepcion.Text
Else
    tpercepcion = 0
    Me.lblPercepcion.Text = "0.00"
End If

Me.lblIMPBruto.Text = Format(neto, "###0.00")
Me.lblDescuento.Text = Format(Tdes, "###0.00")
Me.lblExonerado.Text = Format(in_exonerado, "###0.00")
Me.LblValorVenta.Text = Format(valor_venta, "###0.00")
Me.LblIgv.Text = Format(igv, "###0.00")
Me.lblTotal.Text = Format(pventa, "###0.00")
 

Me.TxtTotalRetencion.Text = Format(in_retencion, "###0.00")
If Val(Me.TxtTotalRetencion.Text) > 0 Then
    Me.chk_suspencion_retencion.Value = 0
Else
    Me.chk_suspencion_retencion.Value = 1
End If



Me.txttotal_final.Text = Val(Me.lblTotal.Text)
Me.lblISC.Text = Format(isc, "###0.00")
Me.lblPercepcion.Text = Format(Val(Me.TxtPecepcion.Text) + Val(Me.lblTotal.Text), "###0.00")
  
 Call llenar_totales(id_compra)
  
  
  Me.cmdAnular.Enabled = True
  Me.cmdEliminar.Enabled = True
  Me.cmdsave.Enabled = False
  Exit Sub

salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub ChameleonBtn2_Click()

End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub chk_cantidad_importacion_Click()
If Me.chk_cantidad_importacion.Value = 1 Then
   Me.chkValorVenta_importacion.Value = 0
End If
End Sub

Private Sub chk_obsequio_Click()
If Me.chk_obsequio.Value = 1 Then
    txtprecioventa.Text = 0
    TxtIgv.Text = 0
    txtValorVenta.Text = 0

End If
End Sub

Private Sub chk_suspencion_retencion_Click()

If Me.chk_suspencion_retencion.Value = 1 Then
   If Val(Me.TxtTotalRetencion.Text) > 0 Then
        
   End If
End If



End Sub

Private Sub ChkExtraer_Click()
If Me.ChkExtraer.Value = 1 Then
    
    Me.TxtSerie_orden.Visible = True
    Me.txtNumero_orden.Visible = True
    Me.DtcComprobante_orden.Visible = True
    Call Resalta(Me.TxtSerie_orden)
    Exit Sub
Else
    
    Me.TxtSerie_orden.Visible = False
    Me.txtNumero_orden.Visible = False
    Me.DtcComprobante_orden.Visible = False
End If
End Sub

Private Sub chkIGV_Click()
    If Val(Me.txtCantidad.Text) > 0 And Val(Me.txtprecioventa.Text) > 0 Then
        Call calculo_igv
    End If
End Sub
Private Sub calculo_igv()
Dim utilidad As Single
    Dim costo As Single
    Dim venta As Single
    Dim unidades As Single
    Dim val_igv As Single
    Dim cantidad As Single
    Dim isc_v As Single
    Dim ivap As Single
    Dim Descuento As Single
    
    venta = Val(Me.txtPrecioVentaAnt.Text)
    unidades = Val(Me.TxtUnidades.Text)
    precio_unit = Val(Me.TxtUnitario.Text)
    cantidad = Val(Me.txtCantidad.Text)
    isc_v = Val(Me.txtisc.Text)
    ivap = Val(Me.txtRetencion.Text)
    Descuento = Val(Me.TxtTotalDescuento.Text)
    
    
    If Me.chkigv.Value = 1 Then
                    Me.txtIgv_Porcentaje.Text = KEY_IGV * 100
                    val_igv = KEY_IGV + 1
                    If ivap > 0 Then
                        Me.txtValorVenta.Text = Format(Val(Me.txtprecioventa.Text) / val_igv, "###0.00")
                        Else
                    
                            If Me.chk_valor_venta.Value = 1 Then
                                Me.TxtIgv.Text = Format(Val(Me.txtprecioventa.Text) * KEY_IGV, "###0.00")
                                Me.txtprecioventa.Text = Format(Val(TxtIgv.Text) + Val(Me.txtValorVenta.Text), "###0.00")
                       
                            Else
                               
                                Me.txtValorVenta.Text = Format((Val(Me.txtprecioventa.Text)) / val_igv, "###0.00")
                                Me.TxtIgv.Text = Format(Val(Me.txtprecioventa.Text) - Val(Me.txtValorVenta.Text), "###0.00")
                            End If
        
        
                    End If
          
          If Val(Me.TxtUnitario.Text) = 0 Then
              Me.TxtUnitario.Text = Val(Me.txtprecioventa.Text) / (Val(Me.txtCantidad.Text) * Val(Me.TxtUnidades.Text))
          End If
          
          costo = (Val(Me.TxtUnitario.Text) - Val(Me.TxtTotalDescuento.Text) / Val(Me.txtCantidad.Text)) + Val(Me.txtOtros.Text) / (Val(Me.txtCantidad.Text) * Val(Me.TxtUnidades.Text))
          
    Else ' Sin Check de IGV
        Me.txtIgv_Porcentaje.Text = "0"
        Me.TxtIgv.Text = 0#
        val_igv = 1
        costo = (precio_unit - Descuento / cantidad) + Val(Me.txtOtros.Text) / (cantidad * unidades)
        If Me.DtcTipoDoc.BoundText <> "0002" Then
            Me.txtprecioventa.Text = Format((precio_unit * cantidad + Val(Me.txtOtros.Text)) * val_igv + isc_v + ivap - Descuento, "###0.00")
            Me.txtValorVenta.Text = Format(Val(Me.txtprecioventa.Text), "###0.00")
        End If
    End If
    If Val(Me.txtCantidad.Text) <= 0 Then
        Exit Sub
    End If
  
    If Me.DtcMoneda.BoundText = "00002" And Me.DtcMoneda.BoundText <> KEY_MONEDA Then
        Me.TxtCostoHoy.Text = Format(costo * Val(Me.txtTc.Text), "###0.00")
    Else
        Me.TxtCostoHoy.Text = Format(costo, "###0.00")
    End If
    
    
    If costo > 0 Then
        utilidad = 15 '(costo - costo) * 100 / costo
        Me.txtUtilidadhoy.Text = Format(utilidad, "###0.00")
        costo = Val(Me.TxtCostoHoy.Text)
        utilidad = Val(Me.txtUtilidadhoy.Text)
        venta = costo + costo * utilidad / 100
        If Val(Me.txtPrecioVentaAnt.Text) = 0 Then
            Me.TxtventaHoy.Text = Format(venta, "###0.00")
        Else
            Me.TxtventaHoy.Text = Format(Me.TxtventaHoy.Text, "###0.00")
        End If
        
    End If
   
   Call Resalta(Me.txtUtilidadhoy)
  
End Sub
Private Sub chkModo_Click()
   Dim unitario As Single
    Dim precio_venta As Single
    Dim cantidad As Single
    Dim unidades As Single
    Dim utilidad As Single
    Dim p_venta As Single
    Dim val_igv As Single
    If (Val(Me.txtIgv_Porcentaje.Text) > 0) Then
        val_igv = KEY_IGV + 1
    Else
        val_igv = 1
    End If
    unidades = Val(Me.TxtUnidades.Text)
    cantidad = Val(Me.txtCantidad.Text)
    precio_venta = Me.txtprecioventa.Text

If Me.chkModo.Value = 1 Then
    
    If cantidad > 0 And unidades > 0 Then
        unitario = precio_venta / (cantidad * unidades)
       ' Me.TxtUnitario.Text = Format(unitario, "#,##0.00")
        If Val(Me.TxtDctoSoles.Text) > 0 Then
            Me.TxtCostoHoy.Text = Format(unitario, "###0.00")
        Else
            Me.TxtCostoHoy.Text = Format(unitario, "###0.00")
        End If
        utilidad = Val(Me.txtUtilidadhoy.Text)
        p_venta = unitario * val_igv + (utilidad * unitario) / 100
        Me.TxtventaHoy.Text = Format(p_venta, "###0.00")
        Call Resalta(Me.TxtventaHoy)
    End If
Else
        
    unitario = Val(Me.TxtUnitario.Text)
    Me.TxtCostoHoy.Text = Format(unitario, "###0.00")
   
    
End If
End Sub

Private Sub ChkPercepcion_Click()
If Me.ChkPercepcion.Value = 1 Then
    Me.TxtPecepcion.Visible = True
    Me.lblpercep_titulo.Visible = True
    Me.lblPercepcion.Visible = True
    Me.TxtPecepcion.Text = 0
    Call Resalta(Me.TxtPecepcion)
    
    Me.lblPercepcion.Text = Format(Val(Me.lblTotal.Text) + Val(Me.TxtPecepcion.Text), "###0.00")
    
   
Else
    Me.lblpercep_titulo.Visible = False
    Me.lblPercepcion.Visible = False
    Me.TxtPecepcion.Visible = False
End If
End Sub





Private Sub chkresponsable_Click()
If Me.chkresponsable.Value = 1 Then
   Me.DtcResponsable.Visible = True
   Me.txtBuscarresponsable.Visible = True
Else
   Me.DtcResponsable.Visible = False
   Me.txtBuscarresponsable.Visible = False
End If
End Sub

Private Sub chkValorVenta_importacion_Click()
If Me.chkValorVenta_importacion.Value = 1 Then
   Me.chk_cantidad_importacion.Value = 0
End If
End Sub

Private Sub cmdactivar_Click()
 

 Call put_fecha_kardex
 
 Call Prender
 
 
End Sub

Private Sub put_fecha_kardex()
strCadena = "SELECT * FROM movimiento_compra WHERE ruc='" & KEY_RUC & "' and  fecha_kardex is null ORDER BY id_compra DESC "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       
       strCadena = "UPDATE movimiento_compra SET fecha_kardex='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' WHERE id_compra='" & rst("id_compra") & "' LIMIT 1"
       CnBd.Execute (strCadena)
       rst.MoveNext
       DoEvents
   Next i
End If

End Sub


Private Sub CmdActualizar_Click()
If Val(Me.TxtCostoHoy.Text) > 0 And Val(Me.TxtventaHoy.Text) > 0 And Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    If MsgBox("ESTA SEGURO DE CAMBIAR LOS PRECIOS", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
        strCadena = "UPDATE producto SET precio_compra='" & Val(Me.TxtCostoHoy.Text) & "',precio_venta='" & Val(Me.TxtventaHoy.Text) & "' WHERE id_producto='" & Trim(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
    End If
End If
End Sub

Private Sub cmdagregar_Click()
    
    Call AgregarNuevo
    Exit Sub


End Sub

Private Sub CmdAgregarG_Click()
Dim IdCompras As Double
'If Me.DtcDocumento.BoundText = "" Or Me.TxtSerieG.Text = "" Or Me.TxtNumeroG.Text = "" Or Val(Me.TxtMontoG.Text) <= 0 Then
 '   MsgBox "Ingrese Datos Correctos", vbInformation, "mensaje para el Usuario"
'Else
 '   IdCompras = IdInsert("DocumentoCompra")
  '  strCadena = "INSERT INTO DocumentoCompraGastos(idCompra,doc_cod,serie,numero,monto,detalle) VALUES " & _
   ' "('" & IdCompras & "','" & Me.DtcDocumento.BoundText & "','" & Me.TxtSerieG.Text & "','" & Me.TxtNumeroG.Text & "','" & Val(Me.TxtMontoG.Text) & "','" & Me.TXtDetalleG.Text & "')"
    'CnBd.Execute (strCadena)
    'Call LlenarGastos(Me.MfGasto, IdCompras)
    'Me.TxtSerieG.Text = ""
    'Me.TxtNumeroG.Text = ""
    'Me.TxtMontoG.Text = 0#
    'Me.TXtDetalleG.Text = ""
    'Me.DtcDocumento.SetFocus
'End If
End Sub

Private Sub cmdCerrar_Click()
Me.frmProrrateo.Visible = False
End Sub

Private Sub cmdAgregar_vinculacion_Click()
strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & Val(Me.lblid_compra_vinculada.Caption) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   If rst("igv") > 0 Then
        afecto_igv = "si"
   Else
        afecto_igv = "no"
   End If
    strCadena = "INSERT INTO movimiento_compra_gasto(id_compra,id_persona,id_doc,serie,numero,monto,monto_total,porcentaje,fecha,descripcion,tc,id_moneda,id_compra_gasto,afecto_igv,ruc)VALUES " & _
     " ('" & Val(Me.txtIdCompra.Text) & "','" & rst("id_proveedor") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & Val(Me.txtMonto_asignado.Text) & "','" & Val(rst("total")) & "','" & Val(Me.txtMonto_porcentaje.Text) & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','-','" & rst("tc") & "','" & rst("id_moneda") & "','" & Val(Me.lblid_compra_vinculada.Caption) & "','" & afecto_igv & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
End If
frmvinculada_detalle.Visible = False
Call llenar_gastos(Me.HfImportaciones, Val(FrmCompras.txtIdCompra.Text))


End Sub

Private Sub cmdAgregarVinculacion_Click()
If Me.frmvinculada_detalle.Visible = True Then
   Me.frmvinculada_detalle.Visible = False
Else
   Me.frmvinculada_detalle.Visible = True
   Me.txtMonto_asignado.Text = ""
   Me.txtMonto_porcentaje.Text = ""
   Me.TxtMontoTotal_vinculado.Text = ""
   Me.txtComprobante_vinculado.Text = ""
   Call Resalta(Me.txtComprobante_vinculado)
   Exit Sub
End If
End Sub

Private Sub cmdAnular_Click()
If MsgBox(MSGANULAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
            
            Procedencia = anular
            FrmSeguridad.Show
            
End If
End Sub

Private Sub cmdAplicarGastos_Click()
Call aplicar_prorrateo_gasto(Val(Me.txtIdCompra.Text), Me.DtcMoneda.BoundText)
End Sub

Private Sub cmdAplicarImportacion_Click()
Call aplicar_prorrateo(Me.txtIdCompra.Text, Me.DtcMoneda.BoundText)
End Sub

Private Sub cmdBuscar_Click()
Procedencia = buscar
         FrmBusquedaCompras.Show
End Sub

Private Sub cmdcerrar_compartir_Click()
Me.frmvinculadas.Visible = False
End Sub

Private Sub cmdcerrardua_Click()
Me.framedua.Visible = False
End Sub

Private Sub cmdcerrarnota_Click()
Frame_Relacionado.Visible = False
End Sub

Private Sub cmdCerrarpantalla_Click()
Me.FrameCaracteristicas.Visible = False
End Sub

Private Sub CMDCONSULTADUA_Click()
frmconsultadua.Show
Exit Sub
End Sub

Private Sub cmdcerrarPlanilla_Click()
Me.frmCuentaDescargo.Visible = False
End Sub

Private Sub cmdCompratir_Click()
Call llenar_gastos(Me.HfImportaciones, Val(FrmCompras.txtIdCompra.Text))
Me.frmvinculadas.Visible = True
End Sub

Private Sub cmdDatosdua_Click()
Me.framedua.Visible = True
End Sub

Private Sub cmdEditable_Click()

End Sub
Private Sub llenar_gastos(ByVal Grilla As MSHFlexGrid, ByVal id_compra As Double)
Dim Total As Double
strCadena = "SELECT * FROM view_factura_vinculada_gasto WHERE id_compra='" & id_compra & "' AND ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    FrmCompras.lblgastos.Text = Format(0, "###0.00")
    Grilla.Rows = 0
    
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 2500
           Grilla.ColWidth(4) = 900
           Grilla.ColWidth(5) = 800
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 1200
           Grilla.ColWidth(8) = 1200
           Grilla.ColWidth(9) = 1200
        Next
         cabecera = "IDGASTO" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "CLIENTE/PROVEEDOR" & vbTab & "MONEDA" & vbTab & "TC" & vbTab & "MONTO" & vbTab & "VALOR VENTA [S/.]" & vbTab & "TOTAL [S/.]" & vbTab & "PORCENTAJE "
         Grilla.AddItem cabecera
         For k = 0 To 9
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
             
             Fila = rst("id_gasto") & vbTab & rst("fecha") & vbTab & rst("comprobante") & vbTab & rst("nombre_completo") & vbTab & rst("moneda") & vbTab & Format(rst("tc"), "#,##0.0000") & vbTab & Format(rst("monto"), "#,##0.0000") & vbTab & Format(in_parcial_venta, "###0.000") & vbTab & Format(in_parcial, "###0.000") & vbTab & Format(rst("porcentaje"), "###0.000")
             Grilla.AddItem Fila
             
        rst.MoveNext
        Next i
         cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "=========================" & vbTab & Format(in_valor_venta, "#,##0.0000") & vbTab & Format(in_total, "#,##0.0000")
         Grilla.AddItem cabecera
          For k = 0 To 9
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &HC0C0FF
                            Next k
                            FrmCompras.lblgastos.Text = Format(in_valor_venta, "###0.00")
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub
Private Sub cmdEliminar_Click()
If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Procedencia = Eliminar
        FrmSeguridad.Show
     End If
End Sub

Private Sub cmdgastos_Click()


strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Me.txtSerie.Text & "' AND  numero='" & Me.TxtNumeroDoc.Text & "' AND id_proveedor='" & Me.txtRuc.Text & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Call disabled_form(Me)
    FrmComprasGastos.txtIdCompra.Text = rst("id_compra")
    FrmComprasGastos.Show
    If DtTipoCompra.BoundText <> "01" Then
        FrmComprasGastos.SSTab1.TabVisible(0) = False
    Else
        FrmComprasGastos.SSTab1.TabVisible(0) = True
        FrmComprasGastos.SSTab1.Tab = 0
    End If
    
Else
    MsgBox "ANTES DE INGRESAR GASTOS, GUARDE ESTE COMPROBANTE", vbInformation, KEY_EMPRESA
End If
Set rst = Nothing

End Sub

Private Sub cmdgastosImportacion_Click()

End Sub

Private Sub CmdImportar_Click()
Call importar_compra
End Sub
Private Sub importar_compra()
Dim in_compra As String
strCadena = "SELECT * FROM movimiento_compra WHERE id_proveedor='" & Me.DtcProveedor.BoundText & "' and serie='" & Trim(TxtSerieR.Text) & "' and numero='" & Trim(Me.TxtNumeroR.Text) & "' and ruc='" & KEY_RUC & "' limit 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   in_compra = rst("id_compra")
   Me.txtid_comprobante_relacionado.Text = in_compra
   Me.DtTipoCompra.BoundText = rst("id_tipo_compra")
   Me.DtTipoCompra.Locked = True
   Me.DtcTipo.BoundText = rst("id_tipo")
   Me.DtcMoneda.BoundText = rst("id_moneda")
   If Me.DtcTipoDoc.BoundText = "0040" Then   ' PERCEPCION
    Me.DtcMoneda.Locked = False
   Else
        Me.DtcMoneda.Locked = True
   End If
   
   Me.txtTc.Text = rst("tc")
   Me.txtObservacion.Text = rst("observacion")
   Call LlenarDatosCliente(Me.DtcProveedor.BoundText)
   strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & in_compra & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRstT(strCadena)
   If rstT.RecordCount > 0 Then
      rstT.MoveFirst
      For i = 0 To rstT.RecordCount - 1
        strCadena = "INSERT INTO movimiento_compra_temporal(id_doc,serie,numero,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento," & _
        "valor_neto,isc,igv,retencion,otros,percepcion,exonerado,valor_venta,precio_venta,p_venta,p_costo,dni_save,id_alm,detalle,ruc) VALUES " & _
        "('" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Trim(Me.txtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & rstT("id_producto") & "'," & _
        "'" & rstT("cantidad") & "','" & rstT("c_unitario") & "','" & rstT("dsto_soles") & "'," & _
        "'" & rstT("dsto_procentaje") & "','" & rstT("total_descuento") & "','" & rstT("valor_neto") & "','" & rstT("isc") & "','" & rstT("igv") & "'," & _
        "'" & rstT("retencion") & "','" & rstT("otros") & "','" & rstT("percepcion") & "','" & rstT("exonerado") & "','" & rstT("valor_venta") & "','" & rstT("total") & "','" & rstT("p_venta") & "'," & _
        "'" & rstT("p_costo") & "','" & KEY_USUARIO & "','" & Trim(Me.DtcAlmacen.BoundText) & "','" & rstT("detalle") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        rstT.MoveNext
      Next i
   End If
   Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcTipoDoc.BoundText, Me.txtSerie.Text)
      
        
        
      
        
        
End If

End Sub
Private Sub cmdImprimir_Click()
strCadena = "SELECT id_compra,comprobante,fecha_registro,fecha_emision,fecha_cancelacion,id_proveedor,nproveedor,id_producto,nombre_prod,unidad,cantidad,c_unitario,dsto_soles,total,total_factura,vv_factura,igv_factura,operador,ruc  FROM view_compra_vista WHERE id_compra='" & Val(Me.txtIdCompra.Text) & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptCompra", , App.Path + "\Reportes\")


End Sub

Private Sub cmdIngresoSerie_Click()

Call parametro_importacion
Call llenar_series_producto(Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)), Me.HfItem)


End Sub

Private Sub cmdInventario_Click()
If Me.DtcTipoDoc.BoundText = "0089" Then
        strCadena = "call p_insert_compra_emitido_premiun('" & Val(Me.txtIdCompra.Text) & "')"
        CnBd.Execute (strCadena)
        Call put_ajuste_ingreso(Val(Me.txtIdCompra.Text))
End If
End Sub

Private Sub cmdModificar_Click()
     
     Procedencia = revertir
     FrmSeguridad.Show
     rever = True
     
     
End Sub

Private Sub cmdnuevo_Click()
 Call nuevo
End Sub

Private Sub cmdProcesar_Click()
Dim estado As String
If Trim(Me.txtnumeroserie.Text) <> "" And Trim(Me.txtseriechasis.Text) <> "" Then
    estado = "si"
Else
    estado = "no"
End If


If Trim(txtnumeroserie) = "" Then
    MsgBox "Ingrese un Valor Valido...", vbInformation
    Exit Sub
End If

strCadena = "SELECT * FROM imp_producto_detalle WHERE nro_motor='" & Trim(Me.txtnumeroserie.Text) & "' and nro_chasis='" & Trim(Me.txtseriechasis.Text) & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
    If MsgBox("ESTA SERIE YA ESTA REGISTRADA" + Chr(13) + "DESEA MODIFICARLA", vbYesNo + vbQuestion) = vbYes Then
        strCadena = "UPDATE imp_producto_detalle SET serie='" & Trim(Me.txtnumeroserie.Text) & "',anio_fabricacion='" & Trim(Me.txtA�oFabricacion.Text) & "',anio_contenedor='" & Trim(Me.TxtAnioDua.Text) & "',nro_contenedor='" & Trim(Me.txtnumero_dua.Text) & "',nro_chasis='" & Trim(Me.txtseriechasis.Text) & "',nro_motor='" & Trim(Me.txtseriemotor.Text) & "',anio_modelo='" & Trim(Me.txta�omodelo.Text) & "',item='" & Trim(Me.txtitemdua.Text) & "',dni_save='" & KEY_USUARIO & "',fecha_mod=CURDATE(),hora_mod=CURTIME(),serie_asignada='" & estado & "' WHERE id_detalle='" & Val(Me.HfItem.TextMatrix(Me.HfItem.Row, 1)) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        Exit Sub
    End If
    
End If




strCadena = "UPDATE imp_producto_detalle SET poliza='" & Trim(Me.txtPoliza.Text) & "',ip='" & Trim(Me.txtIP.Text) & "', serie='" & Trim(Me.txtnumeroserie.Text) & "',anio_fabricacion='" & Trim(Me.txtA�oFabricacion.Text) & "',anio_contenedor='" & Trim(Me.TxtAnioDua.Text) & "',nro_contenedor='" & Trim(Me.txtnumero_dua.Text) & "',nro_chasis='" & Trim(Me.txtseriechasis.Text) & "',nro_motor='" & Trim(Me.txtseriemotor.Text) & "',anio_modelo='" & Trim(Me.txta�omodelo.Text) & "',item='" & Trim(Me.txtitemdua.Text) & "',dni_save='" & KEY_USUARIO & "',fecha_mod=CURDATE(),hora_mod=CURTIME(),serie_asignada='" & estado & "' WHERE id_detalle='" & Val(Me.HfItem.TextMatrix(Me.HfItem.Row, 1)) & "' AND ruc='" & KEY_RUC & "'"


'strCadena = "UPDATE imp_producto_detalle SET serie='" & Trim(Me.txtnumeroserie.Text) & "',anio_fabricacion='" & Trim(Me.txtA�oFabricacion.Text) & "',anio_contenedor='" & Trim(Me.TxtAnioDua.Text) & "',nro_contenedor='" & Trim(Me.txtnumero_dua.Text) & "',nro_chasis='" & Trim(Me.txtseriechasis.Text) & "',nro_motor='" & Trim(Me.txtseriemotor.Text) & "',anio_modelo='" & Trim(Me.txta�omodelo.Text) & "',item='" & Trim(Me.txtitemdua.Text) & "',dni_save='" & KEY_USUARIO & "',fecha_mod=CURDATE(),hora_mod=CURTIME(),serie_asignada='" & estado & "' WHERE nro_chasis='" & Trim(Me.txtChasisG.Text) & "' and nro_motor='" & Trim(Me.txtMotorG.Text) & "' AND ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
 

strCadena = "SELECT * FROM kardex where nro_chasis='" & Trim(Me.txtChasisG.Text) & "' and nro_motor='" & Trim(Me.txtMotorG.Text) & "'"


If estado = "si" Then
           
           Me.HfItem.TextMatrix(Me.HfItem.Row, 6) = Chr(254)
            For j = 0 To 6
                Me.HfItem.col = j
                HfItem.Row = Me.HfItem.Row
                HfItem.CellBackColor = &HC0FFC0
            Next j
        Else
          
            Me.HfItem.TextMatrix(Me.HfItem.Row, 6) = Chr(168)
            For j = 0 To 6
                HfItem.col = j
                HfItem.Row = Me.HfItem.Row
                HfItem.CellBackColor = &HFFFFFF
            Next j
        End If
    If Val(Me.HfItem.TextMatrix(Me.HfItem.Row, 0)) < Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 4)) Then
        HfItem.Row = HfItem.Row + 1
        HfItem.col = 0
        HfItem.ColSel = 6
        HfItem.RowSel = HfItem.Row
        Me.txtitemdua.Text = Format(Val(Me.txtitemdua.Text) + 1, "0000")
        'Me.txtnumeroserie.Text=""
        'Me.txtseriechasis.Text = ""
        'Me.txtseriemotor.Text = ""
        'Call Resalta(Me.txtnumeroserie)
        If Me.txtseriechasis.Visible = True Then
            Call Resalta(Me.txtseriechasis)
        Else
            If Me.txtseriechasis.Visible = True Then
                Call Resalta(Me.txtseriechasis)
            Else
                If Me.txtnumeroserie.Visible = True Then
                   Call Resalta(Me.txtnumeroserie)
                End If
            End If
        End If
        
        
        
        
        Exit Sub
    End If

End Sub

Private Sub cmdProrrateoGastos_Click()

Call prorratear_gastos(Me.txtIdCompra.Text)
Me.cmdAplicarGastos.Enabled = True

End Sub

Public Sub aplicar_prorrateo(ByVal in_compra As String, ByVal in_moneda As String)
Dim in_costo As Double

If get_periodo_cierre(Me.DtcPeriodo.BoundText, "compras") = True Then
    MsgBox "PERIODO DE COMPRAS " + Space(2) + Me.DtcPeriodo.Text + Space(2) + "CERRADO" + Chr(13) + "CONSULTE CON EL AREA CONTABLE", vbInformation, KEY_VENDEDOR
    Exit Sub
End If

If get_periodo_cierre_fecha(Me.DtpKardex.Value) = True Then
        MsgBox "PERIODO DE LA FECHA KARDEX QUE INTENTA INGRESAR.!!!" + Chr(13) + "YA ESTA CERRADO.", vbInformation, KEY_VENDEDOR
        Exit Sub
End If
      
      


strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & Val(in_compra) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "UPDATE movimiento_compra_detalle SET p_venta=p_venta+'" & rst("incremento_neto") / rst("cantidad") & "',p_costo=p_costo+'" & rst("incremento_neto") / rst("cantidad") & "' WHERE id_detalle_compra='" & rst("id_detalle_compra") & "'"
        CnBd.Execute (strCadena)
        
        If in_moneda = "00002" Then
            If KEY_CON_IGV = "si" Then
                in_costo = (rst("valor_venta") / rst("cantidad") + rst("incremento_neto") / rst("cantidad")) * Val(Me.txtTc.Text)
            Else
                in_costo = (rst("total") / rst("cantidad") + rst("incremento_neto") / rst("cantidad")) * Val(Me.txtTc.Text)
            End If
            
        Else
            
            If KEY_CON_IGV = "si" Then
                in_costo = rst("valor_venta") / rst("cantidad") + rst("incremento_neto") / rst("cantidad")
            Else
                in_costo = rst("total") / rst("cantidad") + rst("incremento_neto") / rst("cantidad")
            End If
            
        End If
        
        Call put_actualizar_kardex(rst("id_producto"), in_compra, in_costo, Me.DtcAlmacen.BoundText)
        
        rst.MoveNext
   Next i
   
   
   
   
       
        
        
        
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   strCadena = "UPDATE movimiento_compra SET prorrateo_importacion='si' WHERE id_compra='" & Val(in_compra) & "'"
   CnBd.Execute (strCadena)
   Me.cmdAplicarImportacion.Enabled = False
   
   Call Me.llenarGrid_Comprobante(Me.HfdDetalle, Val(Me.txtIdCompra.Text))
   
   
   
End If





End Sub

Public Sub aplicar_prorrateo_gasto(ByVal in_compra As String, ByVal in_moneda As String)
Dim in_costo As Double
   
If get_periodo_cierre(Me.DtcPeriodo.BoundText, "compras") = True Then
    MsgBox "PERIODO DE COMPRAS " + Space(2) + Me.DtcPeriodo.Text + Space(2) + "CERRADO" + Chr(13) + "CONSULTE CON EL AREA CONTABLE", vbInformation, KEY_VENDEDOR
    Exit Sub
End If

If get_periodo_cierre_fecha(Me.DtpKardex.Value) = True Then
        MsgBox "PERIODO DE LA FECHA KARDEX QUE INTENTA INGRESAR.!!!" + Chr(13) + "YA ESTA CERRADO.", vbInformation, KEY_VENDEDOR
        Exit Sub
End If

   
   
   
   strCadena = "UPDATE movimiento_compra SET prorrateo_gastos='si' WHERE id_compra='" & Val(in_compra) & "'"
   CnBd.Execute (strCadena)
   
   MsgBox "SE VA A PROCEDER A ACTUALIZAR EL KARDEX" + Chr(13) + "DE LOS " + Space(2) + str(Me.lblCantidad.Caption) + Space(2) + "PRODUCTOS", vbInformation, "UN MOMENTO..."
   
   
   strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & Val(in_compra) & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRstUpdate(strCadena)
      If rstUpdate.RecordCount > 0 Then
         rstUpdate.MoveFirst
         Me.progresbar_kardex.Min = 0
         Me.progresbar_kardex.Max = rstUpdate.RecordCount
         For i = 0 To rstUpdate.RecordCount - 1
             
            Call update_kardex_update(rstUpdate("id_producto"), Me.DtpKardex.Value)
            rstUpdate.MoveNext
            DoEvents
            Me.progresbar_kardex.Value = i
         Next i
         
         
      End If
      
  MsgBox "ACTUALIZACION DE KARDEX CORRECTA", vbInformation
   
   
   
   
   Me.cmdAplicarImportacion.Enabled = False
   
   Call Me.llenarGrid_Comprobante(Me.HfdDetalle, Val(Me.txtIdCompra.Text))
   
   
   




strCadena = "UPDATE movimiento_compra SET prorrateo_importacion='si' WHERE id_compra='" & Val(in_compra) & "'"
CnBd.Execute (strCadena)


End Sub


Public Sub prorratear_importacion(ByVal in_compra As String)
Dim in_monto_fs As Single

If Me.chkValorVenta_importacion.Value = 0 And Me.chk_cantidad_importacion.Value = 0 Then
   MsgBox "Debe Seleccionar el tipo de Prorrateo [CANTIDAD - VALOR VENTA]"
   Exit Sub
End If

strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & Val(in_compra) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   
   in_monto_fs = Val(Me.TxtFlete.Text) + Val(Me.txtSeguro.Text)
   
   For i = 0 To rst.RecordCount - 1
        
        If KEY_CON_IGV = "si" Then
            in_monto_parcial = rst("total") / (Val(Me.LblValorVenta.Text) - Val(Me.TxtFlete.Text) - Val(Me.txtSeguro.Text)) * in_monto_fs
        Else
            in_monto_parcial = rst("total") / (Val(Me.txtFob.Text)) * in_monto_fs
        End If
        in_monto_porcentaje = in_monto_parcial * 100 / in_monto_fs
        
        
        
        strCadena = "UPDATE movimiento_compra_detalle SET incremento_neto='" & in_monto_parcial & "',incremento='" & in_monto_porcentaje & "'  WHERE id_detalle_compra='" & rst("id_detalle_compra") & "'"
        CnBd.Execute (strCadena)
        
       
        
        
        rst.MoveNext
   Next i
   
   Call Me.llenarGrid_Comprobante(Me.HfdDetalle, Val(in_compra))
   Call llenar_totales(in_compra)
  
End If


End Sub
Public Sub prorratear_gastos(ByVal in_compra As String)
Dim in_monto_gasto As Single
Dim in_prorrateo_importacion As String

in_monto_parcial = 0
If Me.chk_valor_venta_gasto.Value = 0 And Me.chk_cantidad_gastos.Value = 0 Then
   MsgBox "Debe Seleccionar el tipo de Prorrateo [CANTIDAD - VALOR VENTA]", vbInformation
   Exit Sub
End If

strCadena = "SELECT prorrateo_importacion FROM movimiento_compra WHERE id_compra='" & Val(in_compra) & "' LIMIT 1 "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   in_prorrateo_importacion = rst("prorrateo_importacion")
End If


strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & Val(in_compra) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   
   in_monto_gasto = Val(Me.lblgastos.Text)
   If Me.DtcMoneda.BoundText = "00002" And Me.DtcMoneda.BoundText <> KEY_MONEDA Then
      in_monto_gasto = in_monto_gasto / Val(Me.txtTc.Text)
      
   End If
   
   
   
   
   For i = 0 To rst.RecordCount - 1
        
        If KEY_CON_IGV = "si" Then
            
            
          If in_prorrateo_importacion = "si" Then
             in_monto_parcial = (rst("valor_venta") + rst("incremento_neto")) / (Val(Me.TxtCif.Text)) * in_monto_gasto
          Else
             in_monto_parcial = rst("valor_venta") / (Val(Me.LblValorVenta.Text)) * in_monto_gasto
          End If
            
            
            
           
        
        
        
        Else
           
            If in_prorrateo_importacion = "si" Then
             in_monto_parcial = (rst("total") + rst("incremento_neto")) / (Val(Me.TxtCif.Text)) * in_monto_gasto
          Else
             in_monto_parcial = rst("total") / (Val(Me.lblTotal.Text)) * in_monto_gasto
          End If
          
                
                
           
            
        End If
        
        
        
        in_monto_porcentaje = in_monto_parcial * 100 / in_monto_gasto
        
        
        strCadena = "UPDATE movimiento_compra_detalle SET incremento_neto_gasto='" & in_monto_parcial & "',incremento_gasto='" & in_monto_porcentaje & "'  WHERE id_detalle_compra='" & rst("id_detalle_compra") & "'"
        CnBd.Execute (strCadena)
        rst.MoveNext
   Next i
   
   Call Me.llenarGrid_Comprobante(Me.HfdDetalle, Val(in_compra))
  
End If


End Sub



Private Sub cmdProrrateoImportacion_Click()

Call prorratear_importacion(Val(Me.txtIdCompra.Text))

End Sub

Private Sub CmdQuitar_Click()
If Me.HfdDetalle.Rows > 0 Then
Me.HfdDetalle.col = 0
Call Quitar(Me.HfdDetalle.Text)
End If
End Sub
Private Sub Quitar(ByVal codigo As String)
If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
strCadena = "DELETE FROM movimiento_compra_temporal WHERE id_temporal='" & Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "' AND dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
 
Call llenarGrid_det(Me.HfdDetalle, Trim(Me.TxtNumeroDoc.Text), Trim(Me.DtcTipoDoc.BoundText), Trim(Me.txtSerie.Text))
End If
End Sub



Private Sub cmdRecalcular_Click()
End Sub

Private Sub cmdrecarlcaular_importaciones_Click()
End Sub


Private Sub cmdRevertir_Click()
If Me.TxtCodProducto.Text <> "" Then
    If Me.chkigv.Value = 1 Then
        exonerado = 0
        igv = Val(Me.TxtIgv.Text)
        valor_venta = Val(Me.txtValorVenta.Text)
        
    Else
        exonerado = Val(Me.txtprecioventa.Text)
        igv = 0
        valor_venta = 0
    
    
    End If
    
    strCadena = "UPDATE movimiento_compra_temporal SET id_producto='" & Trim(Me.TxtCodProducto.Text) & "',cantidad='" & Val(Me.txtCantidad.Text) * Val(Me.TxtUnidades.Text) & "'," & _
    "c_unitario='" & Val(Me.TxtUnitario.Text) & "',dsto_soles='" & Val(Me.TxtDctoSoles.Text) & "',dsto_procentaje='" & Val(Me.TxtDstoporcentaje.Text) & "',total_descuento='" & Val(Me.TxtTotalDescuento.Text) & "'," & _
    "valor_neto='" & Val(Me.TxtValoerNeto.Text) & "',isc='" & Val(Me.txtisc.Text) & "',igv='" & igv & "',ivap='" & Val(Me.txtRetencion.Text) & "',otros='" & Val(Me.txtOtros.Text) & "',percepcion='" & Val(Me.txtOtros.Text) & "'," & _
    "exonerado='" & exonerado & "',valor_venta='" & valor_venta & "',precio_venta='" & Val(Me.txtprecioventa.Text) & "',p_venta='" & Val(Me.TxtventaHoy.Text) & "',p_costo='" & Val(Me.TxtCostoHoy.Text) & "' WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' AND id_temporal='" & Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "'"
    CnBd.Execute (strCadena)
     
    
    Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcTipoDoc.BoundText, Me.txtSerie.Text)
End If
End Sub



Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub put_saldo_stock(ByVal in_producto As String)

strCadena = "SELECT DISTINCT id_producto,id_alm FROM kardex where id_producto='" & Trim(in_producto) & "' and  ruc='" & KEY_RUC & "' ORDER BY id_producto ASC,id_alm ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    in_costo = 0
    in_saldo = 0
    For i = 0 To rst.RecordCount - 1
    
    
    
ini:
      in_costo = 0
      in_saldo = 0
        'strCadena = "call put_crear_kardex_id_producto_internacional('" & rst("id_producto") & "','" & rst("id_alm") & "','" & KEY_RUC & "')"
        'CnBd.Execute (strCadena)
        strCadena = "SELECT * FROM kardex WHERE id_alm='" & rst("id_alm") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then

            rstK.MoveFirst
            For j = 0 To rstK.RecordCount - 1
                in_saldo = in_saldo + rstK("cantidad_real")
                If Val(in_saldo) <> rstK("saldo_stock") Then
                    strCadena = "UPDATE kardex SET saldo_stock='" & Val(in_saldo) & "' WHERE id_producto='" & rst("id_producto") & "' and  id_kardex='" & rstK("id_kardex") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                    CnBd.Execute (strCadena)
                    GoTo ini
                End If
                
                If j = rstK.RecordCount - 1 Then
                    strCadena = "SELECT sum(cantidad_real) FROM kardex WHERE id_alm='" & rst("id_alm") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstL(strCadena)
                    If rstL(0) <> Val(in_saldo) Then
                        MsgBox "CANTIDAD REAL:" & rstL(0) + Chr(13) + "CANTIDAD FINAL:" & str(in_saldo)
                    End If
                End If
                rstK.MoveNext
                
            Next j
        End If
        in_costo = 0
        in_saldo = 0
        

        
       
        
       
        rst.MoveNext
        DoEvents
        
    Next i
End If
End Sub

Public Function get_verifica_fecha_periodo(ByVal in_periodo As String, ByVal in_fecha As Date) As Boolean

strCadena = "SELECT * FROM con_periodo WHERE Id='" & in_periodo & "' LIMIT 1"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
    If Format(Me.DtpKardex.Value, "dd-mm-YYYY") < Format(rstIN("FechaInicio"), "dd-mm-YYYY") Then
       MsgBox "ESTIMADO USUARIO.  " + Chr(13) + "La fecha de Ingreso a Kardex de estar contenida en" + Chr(13) + Chr(13) + Format(rst("FechaInicio"), "dd/mm/YYYY") & " - " & Format(rst("FechaFin"), "dd/mm/YYYY"), vbInformation
       get_verifica_fecha_periodo = False
    Else
       get_verifica_fecha_periodo = True
    End If
    
End If

End Function

Private Sub cmdsave_Click()
      
      
     If get_periodo_cierre(Me.DtcPeriodo.BoundText, "compras") = True Then
        
        MsgBox "PERIODO DE COMPRAS CERRARDO.!!!", vbInformation, KEY_VENDEDOR
        Exit Sub
        
     End If
    If get_periodo_cierre_fecha(Me.DtpKardex.Value) = True Then
        MsgBox "PERIODO DE LA FECHA KARDEX QUE INTENTA INGRESAR.!!!" + Chr(13) + "YA ESTA CERRADO.", vbInformation, KEY_VENDEDOR
        Exit Sub
    End If
      
      
      
      
      
      
      
      
      Me.cmdsave.Enabled = False
      Dim in_mes As Integer
      Dim in_anio As String
      in_mes = Month(CVDate(Me.TxtFecha_emision.Text))
      in_anio = Year(CVDate(Me.TxtFecha_emision.Text))
      
      
      If KEY_PAIS <> KEY_PERU Then
        If Trim(Me.txtNumeroAutorizacion.Text) = "" Then
        
            MsgBox "INGRESE UN NUMERO DE AUTORIZACION VALIDO", vbInformation, KEY_VENDEDOR
            Me.cmdsave.Enabled = True
            Exit Sub
        End If
      End If
      
      
If Me.DtTipoCompra.BoundText = "01" Then

If get_verifica_fecha_periodo(Me.DtcPeriodo.BoundText, Me.DtpKardex.Value) = False Then
     Me.cmdsave.Enabled = True
     Exit Sub
End If
End If
      
      
      
    If Format(TxtFecha_emision.Text, "YYYY-mm-dd") >= Format(get_fecha_periodo_abierto_compras(Me.TxtFecha_emision.Text), "YYYY-mm-dd") Then
    
    Else
        MsgBox "PERIODO CERRADO COORDINE CON EL AREA CONTABLE", vbInformation
         Me.cmdsave.Enabled = True
        Exit Sub
    End If
      
      
      
      
      
      
      If Trim(Me.DtcTipoDoc.BoundText) <> "" And Trim(Me.txtSerie.Text) <> "" And Trim(Me.TxtNumeroDoc.Text) <> "" Then
      
      Else
         MsgBox "DATOS OBLIGATORIOS :" + Chr(13) + "[1] Tipo de Comprobante." + Chr(13) + "[2] Numero de Serie y  Numero."
         Call Resalta(Me.txtSerie)
         Exit Sub
      End If
      
      
      
      If KEY_CONTABILIDAD = "si" Then
        If validar_contabilidad(Me.DtTipoCompra.BoundText, Me.DtcTipo.BoundText) = False Then
            Exit Sub
        End If
      End If
      
      
      
      
      If validar_periodo(in_mes, in_anio, Me.DtcPeriodo.BoundText) = False Then
         MsgBox "INCONSISTENCIA EN PERIODO Y FECHA DE EMISION", vbInformation, KEY_VENDEDOR
         Exit Sub
      End If
        
      If Me.DtcTipoDoc.BoundText = "0002" Then
         If Len(Me.txtRuc.Text) < 11 Then
            MsgBox "Ingrese un Ruc Valido" + Chr(13) + "RECIBO POR HONORARIOS", vbInformation
            Call Resalta(Me.txtRuc)
            Exit Sub
         End If
      End If
        
      
      
      
      
      Me.cmdgastos.Visible = True
      
      If Me.DtTipoCompra.BoundText = "01" Then
         If validar_fua = True Then
            Call Save
            
            Call llenarGrid_Comprobante(Me.HfdDetalle, Val(Me.txtIdCompra.Text))
            
            Me.cmdgastos.Enabled = True
         End If
      Else
            Call Save
            Call llenarGrid_Comprobante(Me.HfdDetalle, Val(Me.txtIdCompra.Text))
            Me.cmdgastos.Enabled = True
      End If
      
     
       If Me.DtcTipoDoc.BoundText = "0422" Then
         Call put_pago_automatico_planilla(Val(Me.txtIdCompra.Text), Val(Me.lblTotal.Text))
      End If
      
      If Me.DtcTipoDoc.BoundText = "0020" Then
        Call put_pago_automatico_retencion(get_comprobante_reten(DtcRelacionado.BoundText, Me.TxtSerieR.Text, Me.TxtNumeroR.Text, Me.DtcProveedor.BoundText), Me.txtIdCompra.Text, Val(Me.lblTotal.Text))
        Call put_pago_automatico_retencion(Me.txtIdCompra.Text, Me.txtIdCompra.Text, Val(Me.lblTotal.Text))
      End If
      
      
      
      
End Sub

Private Sub put_pago_automatico_planilla(ByVal in_compra As String, ByVal in_monto As Double)
                    KEY_VENCIMIENTO = KEY_FECHA
                    id_tipo_factura = "0002"
                    igv = "si"
                    dfac = "no"
                    
                    strCadena = "SELECT numero,serie FROM  movimiento_venta WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and  id_doc='0097' and serie='001'  and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
                    Call ConfiguraRstZ(strCadena)
                    If rstZ.RecordCount > 0 Then
                        in_numero = Format(Val(rstZ("numero")) + 1, "000000")
                        in_serie = rstZ("serie")
                    Else
                        in_numero = Format(1, "000000")
                        in_serie = "001"
                    End If
                    
                    
                    horario = Format(Time, "hh:mm")
                    If horario >= "07:00" And horario <= "13:00" Then
                        turno = "M"
                    Else
                        turno = "T"
                    End If
                    
                    Documento = "RECIBO EGRESO" & ":" & Trim(in_serie) & "-" & Trim(in_numero)
                    strCadena = "call P_insert_venta('0097','" & KEY_ALM & "','" & get_pago_anterior(Me.DtcAlmacen.BoundText) & "','00001','no'," & _
                    "'" & in_serie & "','" & in_numero & "','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.TxtProveedor.Text) & "','0','0','0','" & Val(Me.lblTotal.Text) & "','0'," & _
                    "'" & Val(Me.lblTotal.Text) & "','0','" & Format(TxtFecha_emision.Text, "YYYY-mm-dd") & "','" & Format(Me.dtpFechaRegistro.Value, "YYYY-mm-dd") & "','" & id_tipo_factura & "','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & Val(Me.txtTc.Text) & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstP(strCadena)
                    
                    id_venta = rstP(0)
                   
                    strCadena = "UPDATE movimiento_venta SET  observacion='" & Trim(Me.txtObservacion.Text) & "' WHERE id_venta='" & id_venta & "'"
                    CnBd.Execute (strCadena)
                    
                    
                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & id_venta & "','00','" & Trim(Me.txtObservacion.Text) & "','-','1','" & Val(Me.lblTotal.Text) & "','0','" & Val(Me.lblTotal.Text) & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                               
                   
                    
                   
                    
                    strCadena = "INSERT INTO movimiento_venta_monto(id_venta,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,banco,cheque,cuenta_contable,forma_pago_contable,flujo_caja,id_cuenta_origen,cta_redondeo,cta_anticipo,monto_redondeo,monto_anticipo,ruc)VALUES " & _
                    "('" & id_venta & "','01','" & get_pago_anterior(Me.DtcAlmacen.BoundText) & "','" & Val(Me.lblTotal.Text) & "','" & Val(Me.lblTotal.Text) * -1 & "','00','-','-','-','-','-','01','-','-','-','-','0','0','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    
                    
                    strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(Me.TxtNumeroDoc.Text + 1), "000000") & "' WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.txtSerie.Text) & "' AND ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    in_documento = "PLANILLA MOV" & ":" & Trim(Me.txtSerie.Text) & "-" & Trim(Me.TxtNumeroDoc.Text)
                    in_codigo_det = procesar_transaccion(KEY_ALM, Me.DtcCuentaDescargo.BoundText, Format(TxtFecha_emision.Text, "YYYY-mm-dd"), "00002", Trim(Me.txtRuc.Text), Trim(Me.TxtProveedor.Text), Trim(Me.txtObservacion.Text), Val(Me.lblTotal.Text), "0", "0", Val(Me.txtIdCompra.Text), in_documento, Val(Me.txtTc.Text), "", 0, 0, Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                    
                    strCadena = "CALL p_insert_pago_factura_vitekey('" & Val(id_venta) & "','" & Val(in_compra) & "','" & Val(Me.lblTotal.Text) & "','" & Val(Me.lblTotal.Text) & "','" & Me.DtcMoneda.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtTc.Text) & "','" & in_codigo_det & "')"
                    CnBd.Execute (strCadena)
                    
                    

End Sub
Private Sub put_pago_automatico_retencion(ByVal in_compra As String, ByVal in_retencion As String, ByVal in_monto As Double)
                    
                    KEY_VENCIMIENTO = KEY_FECHA
                    id_tipo_factura = "0002"
                    igv = "si"
                    dfac = "no"
                    
                    
                    strCadena = "SELECT numero,serie FROM  movimiento_venta WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and  id_doc='0097' and serie='001'  and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
                    Call ConfiguraRstZ(strCadena)
                    If rstZ.RecordCount > 0 Then
                        in_numero = Format(Val(rstZ("numero")) + 1, "000000")
                        in_serie = rstZ("serie")
                    Else
                        in_numero = Format(1, "000000")
                        in_serie = "001"
                    End If
                    
                    horario = Format(Time, "hh:mm")
                    If horario >= "07:00" And horario <= "13:00" Then
                        turno = "M"
                    Else
                        turno = "T"
                    End If
                    
                    Documento = "RECIBO EGRESO" & ":" & Trim(in_serie) & "-" & Trim(in_numero)
                    strCadena = "call P_insert_venta('0097','" & KEY_ALM & "','" & get_pago_anterior(Me.DtcAlmacen.BoundText) & "','" & Me.DtcMoneda.BoundText & "','no'," & _
                    "'" & in_serie & "','" & in_numero & "','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.TxtProveedor.Text) & "','0','0','0','" & Val(Me.lblTotal.Text) & "','0'," & _
                    "'" & Val(Me.lblTotal.Text) & "','0','" & Format(TxtFecha_emision.Text, "YYYY-mm-dd") & "','" & Format(Me.dtpFechaRegistro.Value, "YYYY-mm-dd") & "','" & id_tipo_factura & "','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & Val(Me.txtTc.Text) & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstP(strCadena)
                    
                    id_venta = rstP(0)
                   
                    strCadena = "UPDATE movimiento_venta SET  observacion='" & Trim(Me.txtObservacion.Text) & "' WHERE id_venta='" & id_venta & "'"
                    CnBd.Execute (strCadena)
                    
                    
                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & id_venta & "','00','" & Trim(Me.txtObservacion.Text) & "','-','1','" & Val(Me.lblTotal.Text) & "','0','" & Val(Me.lblTotal.Text) & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                               
                   
                    
                   
                    
                    strCadena = "INSERT INTO movimiento_venta_monto(id_venta,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,banco,cheque,cuenta_contable,forma_pago_contable,flujo_caja,id_cuenta_origen,cta_redondeo,cta_anticipo,monto_redondeo,monto_anticipo,ruc)VALUES " & _
                    "('" & id_venta & "','01','" & get_pago_anterior(Me.DtcAlmacen.BoundText) & "','" & Val(Me.lblTotal.Text) & "','" & Val(Me.lblTotal.Text) * -1 & "','00','-','-','-','-','-','01','-','-','-','-','0','0','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    
                    
                    strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(Me.TxtNumeroDoc.Text + 1), "000000") & "' WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.txtSerie.Text) & "' AND ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    in_documento = "RETENCION" & ":" & Trim(Me.txtSerie.Text) & "-" & Trim(Me.TxtNumeroDoc.Text)
                    in_codigo_det = procesar_transaccion(KEY_ALM, Me.DtcCuentaDescargo.BoundText, Format(TxtFecha_emision.Text, "YYYY-mm-dd"), "00002", Trim(Me.txtRuc.Text), Trim(Me.TxtProveedor.Text), Trim(Me.txtObservacion.Text), Val(Me.lblTotal.Text), "0", "0", Val(Me.txtIdCompra.Text), in_documento, Val(Me.txtTc.Text), "", 0, 0, Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                    
                    strCadena = "CALL p_insert_pago_factura_vitekey('" & Val(id_venta) & "','" & Val(in_compra) & "','" & Val(Me.lblTotal.Text) & "','" & Val(Me.lblTotal.Text) & "','" & Me.DtcMoneda.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtTc.Text) & "','" & in_codigo_det & "')"
                    CnBd.Execute (strCadena)
                    
                    

End Sub

Private Function get_pago_anterior(ByVal in_alm As String)
strCadena = "SELECT id_registro FROM forma_pago_detalle WHERE id_alm='" & in_alm & "' and id_detalle='10' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
    get_pago_anterior = rstA("id_registro")
Else
    get_pago_anterior = "0"
End If

End Function
Private Function validar_contabilidad(ByVal in_tipo_compra As String, ByVal in_tipo As String) As Boolean
validar_contabilidad = True

If in_tipo_compra = "01" Then ' IMPORTACION
   If in_tipo = "01" Then ' material
      strCadena = "SELECT * FROM view_validar_conta_material WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "'"
      Call ConfiguraRst(strCadena)
      If rst.RecordCount > 0 Then
         rst.MoveFirst
         For i = 0 To rst.RecordCount - 1
             If rst("nro_cuenta_importacion") = "" Then
                MsgBox "Configure CTA IMPORTACION " + Chr(13) + "Linea:" + rst("descripcion"), vbInformation, "Consulte con el Area Contable"
                validar_contabilidad = False
                Me.cmdsave.Enabled = True
                Exit For
             End If
             rst.MoveNext
         Next i
      End If
   Else
    strCadena = "SELECT * FROM view_validar_conta_servicio WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
      If rst.RecordCount > 0 Then
         rst.MoveFirst
         For i = 0 To rst.RecordCount - 1
             If rst("cta_contable") = "0" Then
                MsgBox "CONFIGURE CUENTA PARA SERVICIO: " + rst("nombre_prod"), vbInformation, "CONSULTE CON CONTABILIDAD"
                validar_contabilidad = False
                Me.cmdsave.Enabled = True
                Exit For
             End If
             rst.MoveNext
         Next i
      End If
   End If
End If


If in_tipo_compra <> "01" Then ' IMPORTACION
   If in_tipo = "01" Then ' material
      strCadena = "SELECT * FROM view_validar_conta_material WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "'"
      Call ConfiguraRst(strCadena)
      If rst.RecordCount > 0 Then
         rst.MoveFirst
         For i = 0 To rst.RecordCount - 1
             If rst("nro_cuenta") = "" Then
                MsgBox "Configure Cta Contable. " + Chr(13) + "LINEA:" + rst("descripcion"), vbInformation, "CONSULTE CON CONTABILIDAD"
                validar_contabilidad = False
                Me.cmdsave.Enabled = True
                Exit For
             End If
             rst.MoveNext
         Next i
      End If
   Else
    strCadena = "SELECT * FROM view_validar_conta_servicio WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
      If rst.RecordCount > 0 Then
         rst.MoveFirst
         For i = 0 To rst.RecordCount - 1
             If rst("cta_contable") = "0" Then
                MsgBox "CONFIGURE CUENTA PARA SERVICIO: " + rst("nombre_prod"), vbInformation, "CONSULTE CON CONTABILIDAD"
                validar_contabilidad = False
                Me.cmdsave.Enabled = True
                Exit For
             End If
             rst.MoveNext
         Next i
      End If
   End If
End If







End Function

Private Sub Command2_Click()

End Sub

Private Sub DtcAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcTipoDoc.SetFocus
End If
End Sub






Private Sub DtcDocumento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Call Resalta(Me.TxtSerieG)
End If
End Sub

Private Sub DtcIva_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call put_renta_iva(DtcRelacionado.BoundText, Me.TxtSerieR.Text, Me.TxtNumeroR.Text, Me.DtcProveedor.BoundText)
    Label21(16).Caption = Val(Me.DtcIva.Text)
    Me.Label21(17).Caption = Me.DtcIva.BoundText
End If
End Sub

Private Sub DtcMoneda_Change()
If Me.DtcMoneda.BoundText = "00002" Then
   Me.lblcambio.Visible = True
   
   Me.chkConvertir.Visible = True
Else
    Me.lblcambio.Visible = True
   
    Me.chkConvertir.Visible = False
End If
End Sub

Private Sub DtcMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtTc)
End If
End Sub

Private Sub DtcPeriodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    strCadena = "SELECT * FROM con_periodo WHERE Id='" & Me.DtcPeriodo.BoundText & "' LIMIT 1"
    Call ConfiguraRstIN(strCadena)
    If rstIN.RecordCount > 0 Then
       Me.DtpKardex.Value = Format(Me.TxtFecha_emision.Text, "dd-mm-YYYY")
    
    End If


    Call Resalta(Me.txtRuc)
End If
End Sub

Private Sub DtcProveedor_Change()
'Me.TxtCodProveedor.text = Me.DtcProveedor.BoundText
'Call buscarcliente
End Sub

Private Sub DtcRelacionado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerieR.SetFocus
End If
End Sub

Private Sub DtcRetencionFuente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call put_renta(DtcRelacionado.BoundText, Me.TxtSerieR.Text, Me.TxtNumeroR.Text, Me.DtcProveedor.BoundText)
    Label21(16).Caption = Val(Me.DtcRetencionFuente.Text)
    Me.Label21(17).Caption = Me.DtcRetencionFuente.BoundText
End If
End Sub

Private Sub DtcTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.DtcMoneda.SetFocus
End If
End Sub

Private Sub DtcTipoDoc_Change()
If Me.DtcTipoDoc.BoundText = "0002" Then
   Me.frmretencion.Visible = True
Else
   Me.frmretencion.Visible = False
End If

If Me.DtcTipoDoc.BoundText = "0007" Or Me.DtcTipoDoc.BoundText = "0008" Or Me.DtcTipoDoc.BoundText = "0009" Or Me.DtcTipoDoc.BoundText = "0020" Or Me.DtcTipoDoc.BoundText = "0427" Or Me.DtcTipoDoc.BoundText = "0040" Or Me.DtcTipoDoc.BoundText = "0091" Or Trim(Me.DtcTipoDoc.Text) = "DETRACCION" Then
    Me.Frame_Relacionado.Visible = True
    Me.cmdsave.Enabled = True
    strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_proveedor='si' ORDER BY nombre_completo LIMIT 10"
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcProveedor)
    If Me.DtcTipoDoc.BoundText = "0007" Or Me.DtcTipoDoc.BoundText = "0008" Then
       Me.frmtiponota.Visible = True
       
       If Me.DtcTipoDoc.BoundText = "0007" Then
          Call load_tipo_nota
       End If
       
       If Me.DtcTipoDoc.BoundText = "0008" Then
          Call load_tipo_debito
       End If
        
        
        
    Else
        Me.frmtiponota.Visible = False
    End If
Else
    Me.Frame_Relacionado.Visible = False
End If
End Sub
Private Sub load_tipo_debito()
strCadena = "SELECT id_tipo_nota as Codigo,CONCAT('[',id_tipo_nota,']:',descripcion) as Descripcion FROM tipo_nota_debito ORDER BY id_tipo_nota"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcTipoNota)

End Sub
Private Sub DtcTipoDoc_Click(Area As Integer)

doc_cod = Me.DtcTipoDoc.BoundText


End Sub

Private Sub DtcTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Me.DtcAlmacen.SetFocus
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.txtSerie)
End If
End Sub
Private Sub load_cuenta_pago(ByVal in_doc As String)

If in_doc = "0422" Then
  
  strCadena = "SELECT id_cuenta as Codigo,cuenta as Descripcion FROM view_mis_cuentas_contable WHERE ruc='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCuentaDescargo)
  Me.frmCuentaDescargo.Visible = True
Else
     Me.frmCuentaDescargo.Visible = False
End If
End Sub


Private Sub DtcTipoDoc_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    
    If Me.DtcTipoDoc.BoundText = "0007" Then
        Me.cmdsave.Enabled = True
        
    End If
    
    
    If KEY_PAIS <> KEY_PERU Then
       Me.TxtAlmacen.Visible = True
       Call Resalta(Me.TxtAlmacen)
       Exit Sub
    Else
       Me.TxtAlmacen.Visible = False
    End If
    
    
    If (Me.DtcTipoDoc.BoundText = "0089" Or Me.DtcTipoDoc.BoundText = "0090" Or Me.DtcTipoDoc.BoundText = "0419" Or Me.DtcTipoDoc.BoundText = "0422") Then
        
        Call load_cuenta_pago(Me.DtcTipoDoc.BoundText)
        
        
        strCadena = "SELECT numero,serie FROM  movimiento_compra WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "'  AND  id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND ruc='" & Trim(KEY_RUC) & "' ORDER BY numero DESC LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
           strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "' and id_alm='" & KEY_ALM & "' and id_doc='" & Me.DtcTipoDoc.BoundText & "' LIMIT 1"
           Call ConfiguraRstK(strCadena)
           If rstK.RecordCount > 0 Then
                Me.txtSerie.Text = rstK("serie")
                Me.TxtNumeroDoc.Text = formato_item(rstK("numero"), 8)
           Else
                MsgBox "USTED NO CUENTA CON NINGUNA SERIE ASIGNADA", vbInformation, "MENSAJE PARA EL USUARIO"
                Exit Sub
           End If
        Else
            Me.txtSerie.Text = rst("serie")
            Me.TxtNumeroDoc.Text = formato_item(rst("numero") + 1, 8)
        End If
        
        
        
        
        
       
    
    strCadena = "SELECT serie,numero FROM movimiento_compra WHERE (numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND serie='" & Trim(Me.txtSerie.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Set rst = Nothing
        strCadena = "SELECT * FROM movimiento_compra_temporal WHERE (serie='" & Trim(Me.txtSerie.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND ruc='" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
        Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcTipoDoc.BoundText, Me.txtSerie.Text)
        
        Me.cmdsave.Enabled = True
        End If
        Set rst = Nothing
        
        
        If IsDate(Me.TxtFecha_emision.Text) = False Then
            Me.TxtFecha_emision.Text = CVDate(KEY_FECHA)
        End If
        
        
        
        If Format(CVDate(Me.TxtFecha_emision.Text), "YYYY-mm-dd") <> KEY_FECHA Then
            If MsgBox("la Fecha no Coincide con la Fecha del Documento...Desea Continuar", vbInformation + vbYesNo, KEY_EMPRESA) = vbYes Then
                Me.txtRuc.Text = ""
                Call Resalta(Me.txtRuc)
            Else
                Call Resalta(Me.TxtFecha_emision)
            End If
        Else
                Me.txtRuc.Text = ""
                Call Resalta(Me.txtRuc)
        ProcendenciaGuia = NuevaGuia
    End If
End If
        '---
        
         Me.TxtFecha_emision.SetFocus
        Set rst = Nothing
    Else
    Me.txtSerie.Text = "0000"
    Me.TxtNumeroDoc.Text = "00000000"
    Call Resalta(Me.txtSerie)
    End If
End If
End Sub








Private Sub DtcUnidad_Change()
Call get_cantidad_agranel_final(Me.TxtCodProducto.Text, Me.DtcUnidad.BoundText)
End Sub

Private Sub DtcUnidad_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Call Resalta(Me.TxtUnitario)
End If


End Sub

Private Sub DtTipoCompra_Change()
If Me.DtTipoCompra.BoundText = "01" Then
   Me.frmImportacion.Visible = True
Else
   Me.frmImportacion.Visible = False
End If
        
        
End Sub
Private Sub get_cantidad_agranel_final(ByVal in_producto As String, ByVal in_unidad As String)

        strCadena = "call ADM_unidad_agranel('2','" & in_producto & "','" & in_unidad & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        If rst("in_agranel") = "si" Then
            Me.TxtUnidades.Visible = True
            Me.DtcUnidadFinal.Visible = True
            Me.TxtUnidades.Text = rst("in_cantidad")
            Call get_unidad_final(in_producto, rst("in_unidad"))
            
        Else
            Me.TxtUnidades.Visible = False
            Me.DtcUnidadFinal.Visible = False
            Me.TxtUnidades.Text = 1
            
        End If
        
        
        
End Sub
Private Sub get_cantidad_agranel_ini(ByVal in_producto As String, ByVal in_unidad As String)

        
        
        strCadena = "call ADM_unidad_agranel('1','" & in_producto & "','" & in_unidad & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        If rst("in_agranel") = "si" Then
            
            Me.TxtUnidades.Visible = True
            Me.DtcUnidadFinal.Visible = True
            Me.TxtUnidades.Text = rst("in_cantidad")
            Call get_cantidad_agranel_final(Me.TxtCodProducto.Text, in_unidad)
            
           
        Else
            Me.TxtUnidades.Visible = False
            Me.DtcUnidadFinal.Visible = False
            Me.TxtUnidades.Text = 1
            
        End If
        
        
        
End Sub

Private Sub get_cantidad_agranel_fin(ByVal in_producto As String, ByVal in_unidad As String)

        strCadena = "call ADM_unidad_agranel('1','" & in_producto & "','" & in_unidad & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        If rst("in_agranel") = "si" Then
            Me.TxtUnidades.Visible = True
            Me.DtcUnidadFinal.Visible = True
            Call get_unidad_final(in_producto, in_unidad)
           
        Else
            Me.TxtUnidades.Visible = False
            Me.DtcUnidadFinal.Visible = False
            Me.TxtUnidades.Text = 1
            
        End If
        
        
        
End Sub


Private Sub DtTipoCompra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    
   Me.DtcTipo.SetFocus
End If
End Sub
Private Sub actualizar_totales()
If Me.DtTipoCompra.BoundText = "01" Then
    Me.lblTotal.Text = Format(Val(Me.txttotal_final.Text), "###0.00")
    Me.LblValorVenta.Text = Format(Val(Me.lblTotal.Text), "###0.00")
    Me.lblIMPBruto.Text = Format(Val(Me.lblTotal.Text), "###0.00")

End If
End Sub
Private Sub Form_Activate()

FrmCompras.DtcTipoDoc.Locked = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If Shift = 2 And KeyCode = Asc("G") Then
    If Int(Me.lblTotal.Text) > 1 Then
        Call Save
        Exit Sub
    Else
        If MsgBox("Esta Intentado Grabar una Factura con Monto CERO" + Chr(13) + "Desea Continuar", vbQuestion + vbYesNo) = vbYes Then
            Call Save
            Exit Sub
        End If
  End If
End If

If KeyCode = 122 Then
    Me.DtTipoCompra.SetFocus
End If
End Sub


Private Sub load_proyecto()
strCadena = "SELECT id_proyecto as Codigo,descripcion as Descripcion FROM mis_proyectos WHERE  finalizado='no' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcProyecto)
Me.chkproyecto.Value = 1
   
Me.chkproyecto.Visible = True
Me.DtcProyecto.Visible = True
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50



Me.HfItem.SelectionMode = flexSelectionByRow
HfItem.FocusRect = flexFocusNone

If KEY_CON_IGV = "si" Then
   Me.chk_afecto_costo.Value = 0
   Me.chk_afecto_costo.Visible = False
End If


If KEY_PROYECTO = "si" Then
   Call load_proyecto
End If


Me.DtcUnidadFinal.Visible = False
Me.Dtpproducto_vencimiento.Value = KEY_FECHA
Me.DtpKardex.Value = KEY_FECHA
Me.DtcResponsable.Visible = False
Me.DtcResponsable.BoundText = 0
Me.txtBuscarresponsable.Visible = False

Me.TxtUnidades.Visible = False

Me.txtTc.Text = KEY_CAMBIO_VENTA
Me.TxtPecepcion.Text = 0
precio_unit = 0
rever = False
    Me.TxtFecha_emision.Mask = ""
    Me.TxtFecha_emision.Text = ""
    Me.TxtFecha_emision.Mask = "##/##/####"
    Me.txtFecha_vencimiento.Mask = ""
    Me.txtFecha_vencimiento.Text = ""
    Me.txtFecha_vencimiento.Mask = "##/##/####"


Me.dtpFechaRegistro.Value = KEY_FECHA
Me.dtpFechaRegistro.Enabled = False
Me.ChkRecalcular.Value = 0
Me.Dtpproducto_vencimiento.Value = KEY_FECHA
  
  
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "'  ORDER BY id_alm"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  Me.DtcAlmacen.Enabled = False
  
  
  strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_personal='si'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcResponsable)
  
  
  'strCadena = "SELECT codigo as Codigo,Descripcion  FROM view_periodo order by codigo"
  strCadena = "SELECT Id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion  FROM con_periodo a order by a.`Ejercicio` DESC, a.`Mes` DESC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodo)
  Me.DtcPeriodo.BoundText = get_periodo_actual(KEY_FECHA)
  
  
   
  
  
  If KEY_PAIS = KEY_PERU Then
    strCadena = "SELECT id_tipo as Codigo,descripcion as Descripcion FROM tipo_producto_contable"
  Else
    strCadena = "SELECT id_tipo as Codigo,descripcion as Descripcion FROM tipo_producto_contable_internacional"
  End If
  
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipo)
  Me.DtcTipo.BoundText = "01"
   
  
  strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM moneda ORDER BY id_moneda ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMoneda)
  Me.DtcMoneda.BoundText = KEY_MONEDA

  
  strCadena = "SELECT tipo_compra as Codigo, descripcion as Descripcion FROM tipo_compra   ORDER BY tipo_compra ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtTipoCompra)
  Me.DtTipoCompra.BoundText = "03"
   
  strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM comprobantes ORDER BY doc_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
  Me.DtcTipoDoc.BoundText = "0001"
  
  strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM comprobantes ORDER BY doc_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcComprobante_orden)
  Me.DtcComprobante_orden.BoundText = "0089"
   
 
  
   strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM comprobantes  ORDER BY doc_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcRelacionado)
  Me.DtcRelacionado.BoundText = "0000"
  
  
  Me.DtcTipoDoc.Enabled = False
  Me.txtSerie.Enabled = False
  Me.TxtNumeroDoc.Enabled = False
  
  Referencia = True
  
  Me.cmdsave.Enabled = False
  Me.cmdModificar.Enabled = False
  
   
  
  
  
  If KEY_USUARIO = "46947665" Or KEY_USUARIO = "42546269" Then
     Me.cmdInventario.Visible = True
  Else
    Me.cmdInventario.Visible = False
  End If
  
  If KEY_PAIS <> KEY_PERU Then
    Me.Label16(0).Caption = "IVA"
    strCadena = "SELECT cuenta as Codigo,porcentaje as Descripcion FROM parametro_iva WHERE iva='si' and  ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcIva)
    
    strCadena = "SELECT cuenta as Codigo,porcentaje as Descripcion FROM parametro_iva WHERE iva='no' and  ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcRetencionFuente)
 Else
    Me.TxtAlmacen.Visible = False
  End If
  
  Me.cmdEliminar.Enabled = True
  
End Sub



Private Sub LblTotalParcial_KeyPress(KeyAscii As Integer)
Dim PrecioCompra As Double
Dim TotalParcial As Double
If KeyAscii = 13 Then
   ' TotalParcial = Me.LblTotalParcial.Text
    PrecioCompra = TotalParcial / Val(Me.txtCantidad.Text)
    'Me.TxtPrecio.Text = Format(PrecioCompra, "#,##0.00")
    Me.cmdAgregar.SetFocus

End If
End Sub








Private Sub parametro_importacion()

strCadena = "SELECT * FROM parametros_produccion WHERE habilitado='si' and  ruc='" & KEY_RUC & "' ORDER BY id asc"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       If rst("codigo") = "vin" Then
          Me.lblvin.Caption = rst("descripcion")
          Me.frmvim.Visible = True
       End If
       If rst("codigo") = "chasis" Then
          Me.lblchasis.Caption = rst("descripcion")
          Me.frmchasis.Visible = True
       End If
       If rst("codigo") = "motor" Then
          Me.lblmotor.Caption = rst("descripcion")
          Me.frmmotor.Visible = True
       End If
       If rst("codigo") = "poliza" Then
          Me.lblpoliza.Caption = rst("descripcion")
          Me.frmpoliza.Visible = True
       End If
       If rst("codigo") = "ip" Then
          Me.lblip.Caption = rst("descripcion")
          Me.frmip.Visible = True
       End If
       
       If rst("codigo") = "item" Then
          Me.lblitem.Caption = rst("descripcion")
          Me.frmitem.Visible = True
       End If
       
       
       
    rst.MoveNext
   Next i
End If

End Sub
Public Sub llenar_series_producto(ByVal id_detalle, ByVal Grilla As MSHFlexGrid)
Dim porcentaje As Single
On Error GoTo salir

Iniciar:
strCadena = "SELECT I.id_detalle,I.serie_asignada,D.id_producto,D.id_alm FROM imp_producto_detalle I,movimiento_compra_detalle D WHERE I.id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' and  I.id_detalle_compra=D.id_detalle_compra AND I.id_detalle_compra='" & id_detalle & "' "
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
    
    If MsgBox("ITEMS INGRESADOS NO CUENTAN CON CADENA DE PRODUCCION" + Chr(13) + "DESEA INGRESARLOS ", vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then
        strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & Val(Me.txtIdCompra.Text) & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            For i = 0 To rst.RecordCount - 1
                strCadena = "INSERT INTO imp_producto_detalle(`id_compra`,`id_detalle_compra`,`id_producto`,`id_alm`,`ruc`)VALUES " & _
                "('" & Val(Me.txtIdCompra.Text) & "','" & rst("id_detalle_compra") & "','" & rst("id_producto") & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                rst.MoveNext
            Next i
            GoTo Iniciar
        End If
        
    End If
  Else
 '   Exit Sub
    
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       
        For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 400
           Grilla.ColWidth(1) = 0
           Grilla.ColWidth(2) = 2800
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 800
           Grilla.ColWidth(5) = 700
           Grilla.ColWidth(6) = 350
        Next
        cabecera = "ID" & vbTab & "IDPRODUCTO" & vbTab & "DESCRIPCION" & vbTab & "MODELO" & vbTab & "COLOR" & vbTab & "UNIDAD"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & rstT("id_alm") & "' AND id_producto='" & rstT("id_producto") & "'"
        Call ConfiguraRstL(strCadena)
        
            c = 6
            NumeroCampo = 6
            
        For i = 0 To rstT.RecordCount - 1
          If rstT("serie_asignada") = "si" Then
              estado = Chr(254)
          Else
              estado = Chr(168)
          End If
          Fila = Format(i + 1, "00") & vbTab & rstT("id_detalle") & vbTab & rstL("nombre_prod") & vbTab & rstL("modelo") & vbTab & rstL("color") & vbTab & rstL("unidad") & vbTab & estado
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
          If rstT("serie_asignada") = "si" Then
            For j = 0 To 6
                Grilla.col = j
                Grilla.Row = i + 1
                Grilla.CellBackColor = &HC0FFC0
            Next j
        End If
          rstT.MoveNext
      Next i
      
      Me.FrameCaracteristicas.Visible = True
      
      
      Exit Sub
salir:
      
End Sub

Private Sub HfdDetalle_DblClick()

If Val(Me.lblgastos.Text) > 0 Then
    Me.frmProrrateo.Visible = True
    Me.txtid_detallecompra.Text = Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))
    
End If


End Sub

Private Sub HfdDetalle_SelChange()

On Error GoTo salir

Dim utilidad As Single
If Val(Me.txtIdCompra.Text) > 0 Then
    If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
        Me.cmdIngresoSerie.Enabled = True
    Else
        Me.cmdIngresoSerie.Enabled = False
    End If
End If


If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) < 1 Then
    Me.TxtCodProducto.Text = ""
    Me.txtCantidad.Text = 0
    Me.TxtDescripcionProducto.Text = ""
    Me.DtcUnidad.BoundText = 0
    Me.txtcosto.Text = 0
    Me.TxtUnitario.Text = 0
    Me.TxtDctoSoles.Text = 0
    Me.TxtDstoporcentaje.Text = 0
    Me.txtOtros.Text = 0
    Me.TxtValoerNeto.Text = 0
    Me.TxtTotalDescuento.Text = 0
    Me.txtisc.Text = 0
    Me.TxtISC_p.Text = 0
    Me.txtRetencion.Text = 0
    Me.txtValorVenta.Text = 0
    Me.TxtIgv.Text = 0
    Me.txtIgv_Porcentaje.Text = 0
    Me.txtprecioventa.Text = 0
    Me.txtPrecioVentaAnt.Text = 0
    Me.TxtCostoAnt.Text = 0
    Me.TxtCostoHoy.Text = 0
    Me.TxtUtilidadAnt.Text = 0
    Me.txtUtilidadhoy.Text = 0
    Me.TxtventaHoy.Text = 0
    Me.TxtCostoHoy.Text = 0
    Me.chkigv.Value = 0
    Me.cmdAgregar.Enabled = True
    'If rever = True Then
    Me.cmdAgregar.Enabled = True
    Me.CmdQuitar.Enabled = True
    Me.cmdRevertir.Visible = False
    'End If
    If Me.cmdsave.Enabled = True Then
    'If Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True Then
    Call Resalta(Me.TxtCodProducto)
    End If
    Exit Sub
End If

       Me.TxtCodProducto.Enabled = True
       Me.txtCantidad.Enabled = True
       Me.TxtDescripcionProducto.Enabled = True
       Me.cmdAgregar.Enabled = False
       If Me.cmdsave.Enabled = True Then
            Me.cmdRevertir.Visible = True
       End If
       TxtCodProducto.Text = Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 1)
       
       strCadena = "SELECT * FROM view_producto WHERE id_producto='" & Trim(Me.TxtCodProducto.Text) & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
       Call ConfiguraRst(strCadena)
       
       TxtDescripcionProducto.Text = Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 2)
       txtcosto.Text = Format(rst("precio_compra"), "###0.00")
       
       
       
       
       TxtCostoAnt.Text = Format(rst("precio_compra"), "###0.00")
       txtPrecioVentaAnt.Text = Format(rst("precio_venta"), "###0.00")
       If rst("precio_compra") = 0 Then
        Me.TxtUtilidadAnt.Text = Format(0, "###0.00")
       Else
        Me.TxtUtilidadAnt.Text = Format(Val(rst("precio_venta") - rst("precio_compra")) * 100 / rst("precio_compra"), "###0.00")
       End If
       
       Call Me.get_unidad(rst("id_producto"), rst("agranel"))
       
       
       TxtventaHoy.Text = Format(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 11), "###0.00")
       Me.TxtCostoHoy.Text = Format(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 12), "###0.00")
       If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 12)) = 0 Then
            utilidad = Format(Val(Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 11)) - 0), "#,##0.00")
       Else
            utilidad = Format(Val(Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 11)) - Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 12))) * 100 / Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 12)), "#,##0.00")
       End If
       Me.txtUtilidadhoy.Text = Format(utilidad, "###0.00")
       
       Me.txtCantidad.Text = Format(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 4), "###0.00")
       Me.TxtUnidades.Text = 1
       Me.TxtUnitario.Text = Format(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 5), "###0.00")
       Me.TxtDctoSoles.Text = Format(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 8), "###0.00")
       If Val(Me.TxtUnitario.Text) = 0 Then
            Me.TxtDstoporcentaje.Text = Format(0, "#,##0.000")
       Else
            Me.TxtDstoporcentaje.Text = Format(Val(Me.TxtDctoSoles.Text) * 100 / Val(Me.TxtUnitario.Text), "#,##0.000")
       End If
       
       'Val (Me.TxtDctoSoles.text) * 100 / Val(Me.TxtUnitario.text)
       Me.txtOtros.Text = Format(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 7), "###0.00")
       Me.TxtValoerNeto.Text = Format(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 6), "###0.00")
       Me.txtValorVenta.Text = Format(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 16), "###0.00")
       Me.txtisc.Text = Format(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 9), "###0.00")
       Me.TxtIgv.Text = Format(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 17), "###0.00")
       If Val(Me.TxtIgv.Text) > 0 Then
            Me.chkigv.Value = 1
            Me.txtIgv_Porcentaje.Text = KEY_IGV * 100
        Else
            Me.chkigv.Value = 0
            Me.txtIgv_Porcentaje.Text = 0
        End If
       
       Me.txtRetencion.Text = Format(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 10), "###0.00")
       Me.txtprecioventa.Text = Format(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 19), "###0.00")
      ' Call Resalta(Me.TxtCodProducto)
       'Set rst = Nothing
    
    
    Exit Sub
salir:
    
    
    

End Sub



Private Sub MfGasto_SelChange()
'If Val(Me.MfGasto.TextMatrix(Me.MfGasto.Row, 0)) > 0 Then
 '   Me.CmdQuitarGasto.Visible = True
'Else
 '  Me.CmdQuitarGasto.Visible = False
'End If
End Sub

 Private Sub ActualizarImagen(ByVal id_detalle As Double, ByVal Grilla As MSHFlexGrid)
     Dim estado As String
      
      strCadena = "SELECT * FROM imp_producto_detalle WHERE id_detalle='" & id_detalle & "'  AND ruc='" & KEY_RUC & "'"
      Call ConfiguraRst(strCadena)
      If rst.RecordCount > 0 Then
        If rst("serie_asignada") = "si" Then
            estado = "si"
           Me.HfItem.TextMatrix(Me.HfItem.Row, 6) = Chr(254)
            
            '   For j = 0 To 6
            '   Me.HfItem.col = j
            '   HfItem.Row = Me.HfItem.Row
            '   HfItem.CellBackColor = &HFFFFFF
            '   HfItem.CellBackColor = &HC0FFC0
            '   Next j
            
        Else
            estado = "si"
            Me.HfItem.TextMatrix(Me.HfItem.Row, 6) = Chr(254)
           '    For j = 0 To 6
           '    HfItem.col = j
           '    HfItem.Row = Me.HfItem.Row
           '    HfItem.CellBackColor = &HC0FFC0
           '  Next j
        End If
        
      End If
      
     strCadena = "UPDATE imp_producto_detalle SET serie_asignada='" & estado & "' WHERE id_detalle='" & Val(Me.HfItem.TextMatrix(Me.HfItem.Row, 1)) & "' AND ruc='" & KEY_RUC & "'"
     CnBd.Execute (strCadena)
     
     Me.txtseriechasis.Text = ""
     Me.txtseriemotor.Text = ""
     Me.txtitemdua.Text = ""
     Me.txtnumeroserie.Text = rst("serie")
     Me.txtseriechasis.Text = rst("nro_chasis")
     Me.txtseriemotor.Text = rst("nro_motor")
     Me.txtPoliza.Text = rst("poliza")
     Me.txtIP.Text = rst("ip")
     Me.txtChasisG.Text = rst("nro_chasis")
     Me.txtMotorG.Text = rst("nro_motor")
      
     Me.txtitemdua.Text = rst("item")
     
     If Me.txtnumeroserie.Visible = True Then
        Call Resalta(Me.txtnumeroserie)
     End If
     
     
      
    End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Procedencia = Selecionar
  FrmBusquedaCompras.Show
End If
End Sub

Private Sub HfItem_Click()
If Val(Me.HfItem.TextMatrix(Me.HfItem.Row, 1)) > 0 Then
    Call ActualizarImagen(Me.HfItem.TextMatrix(Me.HfItem.Row, 1), Me.HfItem)
End If
End Sub

Private Sub HfPersonal_Click()

End Sub

Private Sub lblImportacion_Change()
Call actualizar_totales
End Sub

Private Sub lblotros_Change()
Me.lblTotal.Text = Val(Me.lblTotal.Text)
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_NEW
        Call nuevo
    Case KEY_UPDATE
    Case KEY_ANULAR
         
    Case "(Buscar)"
         
    Case KEY_DELETE
      
    Case KEY_EXIT
      Procedencia = Neutro
      Unload Me
      Exit Sub
  End Select
End Sub
Public Sub nuevo()
Dim cCompra As Double
    strCadena = "DELETE FROM movimiento_compra_temporal WHERE dni_save='" & Trim(KEY_USUARIO) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
     
   ' cCompra = IdInsert("DocumentoCompra")
    
    If Me.DtcTipoDoc.BoundText = "0089" Then
        strCadena = "SELECT * FROM  almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND  serie='" & Trim(Me.txtSerie.Text) & "' AND ruc='" & Trim(KEY_RUC) & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "'"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
        Me.txtSerie.Text = rstZ("serie")
        Me.TxtNumeroDoc.Text = formato_item(rstZ("numero"), 8)
        End If
    Else
    
    Me.txtSerie.Text = "000"
    Me.TxtNumeroDoc.Text = "00000000"
    Me.lblidCompra.Caption = ""
    End If
    Me.HfdDetalle.Rows = 0
    rever = False
    
    Me.txtid_comprobante_relacionado.Text = ""
    Me.txtvalidacion_chasis.Text = "no"
    Me.TxtFecha_emision.Mask = ""
    Me.TxtFecha_emision.Text = ""
    Me.TxtFecha_emision.Mask = "##/##/####"
    Me.txtFecha_vencimiento.Mask = ""
    Me.txtFecha_vencimiento.Text = ""
    Me.txtFecha_vencimiento.Mask = "##/##/####"
    Me.cmdProrrateoImportacion.Enabled = False
    Me.chkValorVenta_importacion.Enabled = False
    Me.chk_cantidad_importacion.Enabled = False
    Me.cmdRevertir.Visible = False
    Me.cmdactualizar.Visible = False
    Me.txtSerie.Enabled = True
    Me.txtnumero_dua.Text = ""
    Me.txtA�oFabricacion.Text = ""
    Me.TxtAnioDua.Text = ""
    Me.txta�omodelo.Text = ""
    Me.txtnumero_dua.Locked = False
    Me.txtA�oFabricacion.Locked = False
    Me.TxtAnioDua.Locked = False
    Me.txta�omodelo.Locked = False
    Me.TxtAnioDua.Locked = False
    Me.txtA�oFabricacion.Locked = False
    Me.txta�omodelo.Locked = False
    Me.TxtNumeroDoc.Enabled = True
    Me.DtTipoCompra.Locked = False
    Me.DtcMoneda.Locked = False
    Me.DtcTipoDoc.Enabled = True
    Me.cmdgastos.Visible = False
    Me.DtcTipoDoc.SetFocus
    Me.txtRuc.Text = ""
    Me.Dtpproducto_vencimiento.Value = KEY_FECHA
    Me.txtLote.Text = ""
    Me.TxtProveedor.Text = "CLIENTE"
    Me.txtDireccion.Text = KEY_DIR_PUBLIC
    Me.lblIMPBruto.Text = ""
    Me.txtObservacion.Text = ""
    Me.TxtDescripcionProducto.Text = ""
    Me.ChkPercepcion.Value = 0
    Me.txtisc.Text = 0#
    Me.lblISC.Text = 0#
    
    Me.ChkExtraer.Value = 0
    Me.DtcComprobante_orden.Visible = False
    Me.TxtSerie_orden.Text = ""
    Me.txtNumero_orden.Text = ""
    Me.TxtSerie_orden.Visible = False
    Me.txtNumero_orden.Visible = False
    
    
    Me.TxtTotalRetencion.Text = 0#
    Me.lblgastos.Text = 0#
    Me.TxtISC_p.Text = 0#
    Me.lblExonerado.Text = 0#
    Me.txtRedondeo.Text = 0#
    Me.lblTotal.Text = 0#
    Me.TxtUtilidadAnt.Text = 0#
    Me.txtUtilidadhoy.Text = 0#
    Me.txtPrecioVentaAnt.Text = 0#
    Me.TxtventaHoy.Text = 0#
    Me.lblDescuento.Text = 0#
    Me.TxtCodProducto = "00000"
    Me.TxtValoerNeto.Text = 0
    Me.txtValorVenta.Text = 0
    Me.txtprecioventa.Text = 0
    Me.chkigv.Value = 0
    Me.TxtDescripcionProducto.Text = ""
    Me.TxtDctoSoles.Text = 0#
    Me.TxtDstoporcentaje.Text = 0#
    Me.txtOtros.Text = 0#
    Me.TxtTotalDescuento.Text = 0#
    Me.txtcosto.Text = 0#
    Me.TxtCostoAnt.Text = 0#
    Me.TxtCostoHoy.Text = 0#
    Me.txtIgv_Porcentaje.Text = 0#
    Me.txtCantidad.Text = 0
    Me.TxtIgv.Text = 0
    Me.txtIgv_Porcentaje.Text = 0
    Me.TxtUnitario.Text = 0#
    Me.LblIgv.Text = 0#
    Me.TxtUnidades.Text = 1
    Me.lblPercepcion.Text = ""
    Me.lblTotal.Text = ""
    Me.LblValorVenta.Text = ""
    Me.lblCantidad.Caption = "0"
    Me.txtFob.Text = ""
    Me.TxtFlete.Text = ""
    Me.txtSeguro.Text = ""
    Me.TxtCif.Text = ""
    Me.txtNumeroAutorizacion.Text = ""
    'Me.TxtSerieG.Text = ""
    'Me.TxtNumeroG.Text = ""
    'Me.TxtTransporte.Text = ""
    'Me.TxtChofer.Text = ""
    'Me.TxtLicencia.Text = ""
    'Me.TxtInicial.Text = ""
    'Me.TxtFinal.Text = ""
    'Me.Txtrecorrido.Text = ""
    'Me.TxtDias.Text = ""
    'Me.HfPersonal.Rows = 1
    'Me.HfPersonal.Clear
    'Me.HfGastos.Rows = 1
    'Me.HfGastos.Clear
    Me.TxtCodProducto.Enabled = True
    Me.TxtDescripcionProducto.Enabled = True
    Me.txtCantidad.Enabled = True
    Me.cmdAgregar.Enabled = True
    Me.CmdQuitar.Enabled = True
    Me.chkresponsable.Value = 0
    Me.lblAnulado.Visible = False
    Me.cmdsave.Enabled = False
    Me.cmdModificar.Enabled = False
    
    Me.HfdDetalle.Rows = 0
    
    
    
    
End Sub

Sub verifica(ByVal doc_deta As String)
    Select Case Val(doc_deta)
        Case 1
           ' Call Doc_Referencia(True, Val(doc_deta))
        Case 3
           ' Call Doc_Referencia(False, Val(doc_deta))
        Case 7
           ' Call Doc_Referencia(True, Val(doc_deta))
        Case 8
            'Call Doc_Referencia(True, Val(doc_deta))
        Case 9
            
            'Call Doc_Referencia(True, Val(doc_deta))
        Case 88
            'Call Doc_Referencia(False, Val(doc_deta))
        Case 89
            'Call Doc_Referencia(False, Val(doc_deta))
        Case 90
            'Call Doc_Referencia(False, Val(doc_deta))
        Case 95
            'Call Doc_Referencia(False, Val(doc_deta))
        Case 96
            'Call Doc_Referencia(True, Val(doc_deta))
    End Select
    
End Sub



Private Sub TlbAgregar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_AGREGAR
        Call AgregarGrilla
    Case KEY_QUITAR
        'Call Quitar
    
  End Select
End Sub

Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
      
    Case KEY_PRINT
       MsgBox "Ingrese el Formato de Impresion", vbInformation, "Mensaje para el Usuario"
    Case KEY_REVERTIR
     
  End Select
  Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Set rst = Nothing
Exit Sub
End Sub
Private Function validar_fua() As Boolean

If Trim(Me.txtvalidacion_chasis.Text) = "si" And Trim(Me.txtnumero_dua.Text) = "" And Trim(Me.txtA�oFabricacion.Text) = "" Then
   validar_fua = False
   MsgBox "INGRESE DATOS DE LA DUA", vbInformation, KEY_EMPRESA
Else
    validar_fua = True
End If


End Function

Private Function CodigoKardex() As String
strCadena = "SELECT int_Kardex FROM Kardex ORDER BY int_kardex DESC"
    Call ConfiguraRst(strCadena)
    CodigoKardex = GeneraCodigos()
    Set rst = Nothing
End Function
Private Sub Imprimir()
If Me.DtcTipoDoc.BoundText = KEY_INGALMA Then
Dim i As Integer, j As Integer
Dim laVenta, espacios
Dim MES As String
Dim Ans As Boolean
Dim cantidad As String, Und As String, descripcion As String, precio As String
Dim Total As String, SUBTOTAL As String, igv As String
Dim totalPar As String
Dim Descuento As String
Dim GranTotal As String
Dim totalletras As String
Dim Peso As Double
Dim inc As Single
Dim codigo As String, Unidad As String, PesoTotal As Double
Dim Toneladas As String
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    Printer.Print Tab(20); "PROVEEDOR:" + Space(1) + Me.TxtProveedor.Text
    Printer.Print Tab(20); "DIRECCION:" + Space(1) + Me.txtDireccion.Text
    Printer.Print Tab(20); "RUC      :" + Space(1) + Me.txtRuc.Text
   ' Printer.Print Tab(20); "FACTURA"; Space(2); Mid(Me.TxtSerie_Ref.Text + Space(50), 1, 4) & Space(2) & "-" & Space(2) & Me.TxtNumero_Ref.Text + Space(6) + "INGALMA"; Space(2); Mid(Me.TxtSerie.Text + Space(50), 1, 4) & Space(2) & "-" & Space(2) & Me.TxtNumeroDoc.Text
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
strCadena = "SELECT Detalle_DocumentoCompra.cProducto, Detalle_DocumentoCompra.cantidad, Unidad.sAbreviatura, Producto.DescripcionProducto, " & _
            "Detalle_DocumentoCompra.precio , Detalle_DocumentoCompra.TOTAL, DocumentoCompra.nTotalCompra FROM DocumentoCompra INNER JOIN " & _
            "Detalle_DocumentoCompra ON DocumentoCompra.cDocumentoCompra = Detalle_DocumentoCompra.cDocumentoCompra AND " & _
            "DocumentoCompra.Alm_cod = Detalle_DocumentoCompra.Alm_cod AND " & _
            "DocumentoCompra.doc_cod = Detalle_DocumentoCompra.doc_cod AND " & _
            "DocumentoCompra.sSerie = Detalle_DocumentoCompra.sSerie INNER JOIN " & _
            "Producto ON Detalle_DocumentoCompra.cProducto = Producto.cProducto INNER JOIN " & _
            "Unidad ON Producto.cUnidad = Unidad.cUnidad " & _
            "WHERE (Detalle_DocumentoCompra.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND Detalle_DocumentoCompra.sSerie='" & Trim(Me.txtSerie.Text) & "' " & _
            "AND Detalle_DocumentoCompra.cDocumentoCompra='" & Trim(Me.TxtNumeroDoc.Text) & "')"
    Call ConfiguraRst(strCadena)
    rst.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.2
            For j = 0 To rst.RecordCount - 1
                codigo = Mid(rst(0) + Space(50), 1, 4)
                cantidad = Mid(str(rst(1)) + Space(10), 1, 4)
                Und = rst(2)
                descripcion = Mid(rst(3) + Space(80), 1, 48)
                precio = Mid(Format(str(rst(4)), "###0.00") + Space(4), 1, 8)
                totalPar = Mid(Format(str(rst(5)), "###0.00") + Space(4), 1, 8)
                'Printer.Print Tab(2); Codigo & Space(7) & Cantidad & Space(1) & Und & Space(6) & descripcion & precio & Space(4) & totalPar
                Printer.Print Tab(5); codigo & Space(2) & cantidad & Space(1) & Und & Space(2) & descripcion & precio & Space(4) & totalPar
                Printer.CurrentY = Printer.CurrentY + 0.2
                rst.MoveNext
            Next j
            inc = 0.5
            Do While (Val(Printer.CurrentY) <= 19)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    rst.MoveFirst
    Total = Format(str(rst(6)), "###0.00")
    Descuento = Format(str(KEY_DSCTO), "###0.00")
    totalletras = UCase(EnLetras(Total))
    Set rst = Nothing
    '---- fin totales
    Printer.CurrentY = Printer.CurrentY + 0.8
    Printer.Print Tab(10); Mid(totalletras + Space(100), 1, 100)
    Printer.Print Tab(55); Mid(Total & Space(20), 1, 13) & Descuento & Space(15) & Total
    Printer.EndDoc
    
    Exit Sub
End If
End Sub

Public Sub Prender()


Me.DtcAlmacen.Enabled = True
Me.DtcTipoDoc.Enabled = True
Me.txtSerie.Enabled = True
Me.TxtNumeroDoc.Enabled = True
Call llenarGrid_det(Me.HfdDetalle, "", "", "")
Me.DtcTipoDoc.SetFocus
End Sub

Private Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub

Private Sub txta�ocontenedor_Change()

End Sub

Private Sub txta�ocontenedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txta�omodelo)
End If
End Sub

Private Sub txtAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtAlmacen.Text = Format(Trim(Me.TxtAlmacen.Text), "000")
    Call Resalta(Me.txtSerie)
End If
End Sub

Private Sub TxtAnioDua_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(txta�omodelo)
End If
End Sub

Private Sub txta�ofabricacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtnumero_dua)
End If
End Sub

Private Sub txtBuscar_Change()
  strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM comprobantes WHERE doc_des LIKE '%" & Trim(Me.txtBuscar.Text) & "%' ORDER BY doc_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
  
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   Me.DtcTipoDoc.SetFocus
End If


End Sub

Private Sub txtBuscarproveedor_Change()
 strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and nombre_completo LIKE '%" & Trim(Me.TxtBuscarProveedor.Text) & "%'  ORDER by nombre_completo"
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcProveedor)
End Sub

Private Sub txtBuscarresponsable_Change()
 strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_personal='si' and nombre_completo LIKE '%" & Trim(Me.txtBuscarresponsable.Text) & "%'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcResponsable)
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
Dim TotalP As Single
If KeyAscii = 13 And Val(Me.txtCantidad.Text) > 0 Then
   
   
    If Me.DtcTipoDoc.BoundText = "0089" Or Me.DtcTipoDoc.BoundText = "0090" Then
       Me.TxtUnitario.Text = Val(Me.TxtCostoAnt.Text)
    End If
 Me.DtcUnidad.SetFocus
End If
End Sub





Private Sub TxtChofer_Change()

End Sub

Private Sub txtComprobante_vinculado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmBusquedaCompras.Show
   
    Exit Sub
End If
End Sub

Private Sub TxtCostoHoy_Change()
Dim costo As Single
If rever = False Then
    costo = Val(Me.TxtCostoHoy.Text)
    Me.TxtventaHoy.Text = Format(costo + costo * Val(Me.txtUtilidadhoy.Text) / 100, "###0.00")
Else
    If Val(Me.TxtCostoHoy.Text) > 0 Then
    utilidad = Val(Me.txtUtilidadhoy.Text)  '(Val(Me.TxtventaHoy.Text) - Val(Me.TxtCostoHoy.Text)) * 100 / Val(Me.TxtCostoHoy.Text)
    Me.txtUtilidadhoy.Text = Format(utilidad, "###0.00")
    Me.TxtventaHoy.Text = Format(Val(Me.TxtCostoHoy.Text) + Val(Me.TxtCostoHoy.Text) * Val(Me.txtUtilidadhoy.Text) / 100, "###0.00")
    End If
End If

End Sub

Private Sub TxtCostoHoy_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtCostoHoy.Text = Format(Me.TxtCostoHoy.Text, "###0.00")
    Call Resalta(Me.txtUtilidadhoy)
End If
End Sub









Private Sub TxtDctoSoles_KeyPress(KeyAscii As Integer)
Dim totalneto As Single
totalneto = Me.txtprecioventa.Text
If KeyAscii = 13 Then
   
    If Val(Me.TxtDctoSoles.Text) > 0 Then
        FrmDescuentos.lblUnitario.Caption = Format(Val(Me.TxtDctoSoles.Text), "###0.00")
        FrmDescuentos.lblTotal.Caption = Format(Val(Me.TxtDctoSoles.Text), "###0.00")
        FrmDescuentos.Show
        Exit Sub
    Else
    Me.TxtDctoSoles.Text = Format(Val(Me.TxtDctoSoles.Text), "###0.00")
    Me.TxtDstoporcentaje.Text = Format(Val(Me.TxtDctoSoles.Text), "###0.00")
    Call Resalta(Me.TxtDstoporcentaje)
    End If
End If
End Sub

Public Sub llenar_descuento()
totalneto = Val(Me.TxtValoerNeto.Text)



If FrmDescuentos.Procedencia = unitario Then
    Me.TxtDctoSoles.Text = Format(Val(FrmDescuentos.lblUnitario.Caption), "###0.00")
   Me.TxtTotalDescuento.Text = Format(Val(FrmDescuentos.lblUnitario.Caption) * Val(Me.txtCantidad.Text), "###0.00")
Else
    
   Me.TxtTotalDescuento.Text = Format(Val(Me.TxtDctoSoles.Text), "###0.00")
End If
        

If Me.chkigv.Value = 1 Then
            
            If Val(Me.TxtUnitario.Text) > 0 Then
                Me.TxtDstoporcentaje.Text = Format(Val(Me.TxtDctoSoles.Text) * 100 / Val(Me.TxtUnitario.Text), "##0.00")
                
            Else
                Me.TxtDstoporcentaje.Text = Format(Val(Me.TxtDctoSoles.Text) * 100, "##0.00")
            End If
            Me.TxtCostoHoy.Text = Format(Val(Me.TxtCostoHoy.Text) * (1 + KEY_IGV) - Val(Me.TxtTotalDescuento.Text) / Val(Me.txtCantidad.Text), "##0.00")
            Me.txtprecioventa.Text = Format(totalneto * (1 + KEY_IGV) - Val(Me.TxtTotalDescuento.Text), "###0.00")
            Me.txtValorVenta.Text = Format(Val(Me.TxtValoerNeto.Text) - Val(Me.TxtTotalDescuento.Text), "##0.00")
    Else
        
      
        If Val(Me.TxtUnitario.Text) > 0 Then
            If Val(TxtDstoporcentaje) = 0 Then
                Me.TxtDstoporcentaje.Text = Format(Val(Me.TxtDctoSoles.Text) * 100 / (Val(Me.TxtUnitario.Text) * Val(Me.txtCantidad.Text)), "##0.00")
            End If
        Else
            Me.TxtDstoporcentaje.Text = Format(Val(Me.TxtDctoSoles.Text) * 100 / (Val(Me.TxtUnitario.Text) * Val(Me.txtCantidad.Text)), "###0.00")
        End If
        
        
        
        Me.TxtCostoHoy.Text = Format(Val(Me.TxtCostoHoy.Text) - Val(Me.TxtTotalDescuento.Text) / Val(Me.txtCantidad.Text), "##0.00")
        Me.txtValorVenta.Text = Format(Val(Me.TxtValoerNeto.Text) - Val(Me.TxtTotalDescuento.Text), "###0.00")
        Me.txtprecioventa.Text = Format(totalneto - Val(Me.TxtTotalDescuento.Text), "###0.00")
    End If
    
    Me.TxtDstoporcentaje.Text = Format(Val(Me.TxtDstoporcentaje.Text), "#,##0.00")
    Call Resalta(Me.txtOtros)
End Sub


Private Sub TXtDetalleG_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Me.CmdAgregarG.SetFocus
End If
End Sub

Private Sub TxtDstoporcentaje_KeyPress(KeyAscii As Integer)
Dim totalneto As Single
totalneto = Val(Me.TxtValoerNeto.Text)
If KeyAscii = 13 Then
If Me.chkigv.Value = 1 Then
          Me.TxtDctoSoles.Text = Format(Val(Me.TxtUnitario.Text) * Val(Me.TxtDstoporcentaje.Text) / 100, "###0.00")
          Me.TxtDstoporcentaje.Text = Format(Me.TxtDstoporcentaje.Text, "#,#0.0000")
                If FrmDescuentos.Procedencia = unitario Then
                    Me.TxtTotalDescuento.Text = Format(Val(Me.TxtDctoSoles.Text), "#,##0.000")
                Else
                    Me.TxtTotalDescuento.Text = Format(Val(Me.TxtDctoSoles.Text) * Val(Me.txtCantidad.Text), "###0.00")
                End If
            Me.txtprecioventa.Text = Format(totalneto * (1 + KEY_IGV) - Val(Me.TxtTotalDescuento.Text), "###0.000")
            Me.TxtCostoHoy.Text = Format(Val(Me.TxtUnitario.Text) * (1 + KEY_IGV) - Val(Me.TxtDctoSoles.Text) + Val(Me.txtOtros.Text) / (Val(Me.txtCantidad.Text) * Val(Me.TxtUnidades.Text)), "###0.00")
    Else
         FrmDescuentos.lblUnitario.Caption = Format(Val(Me.TxtUnitario.Text) * Val(Me.TxtDstoporcentaje.Text) / 100, "###0.000")
         FrmDescuentos.lblTotal.Caption = Format((Val(Me.TxtUnitario.Text) * Val(Me.TxtDstoporcentaje.Text) / 100) * Val(Me.txtCantidad.Text), "###0.000")
         FrmDescuentos.Show
         Exit Sub
        Me.TxtDctoSoles.Text = Format(Val(Me.TxtUnitario.Text) * Val(Me.TxtDstoporcentaje.Text) / 100, "###0.00")
        Me.TxtDstoporcentaje.Text = Format(Me.TxtDstoporcentaje.Text, "###0.00")
                If FrmDescuentos.Procedencia = unitario Then
                    Me.TxtTotalDescuento.Text = Format(Val(Me.TxtDctoSoles.Text), "###0.00")
                    Me.TxtCostoHoy.Text = Format(Val(Me.TxtUnitario.Text) - Val(Me.TxtDctoSoles.Text) / Val(Me.txtCantidad.Text), "###0.00")
                    Me.txtValorVenta.Text = Format(Val(Me.TxtValoerNeto.Text) - Val(Me.TxtTotalDescuento.Text), "###0.00")
                Else
                    Me.TxtTotalDescuento.Text = Format(Val(Me.TxtDctoSoles.Text) * Val(Me.txtCantidad.Text), "###0.00")
                    Me.TxtCostoHoy.Text = Format(Val(Me.TxtUnitario.Text) - Val(Me.TxtDctoSoles.Text) + Val(Me.txtOtros.Text) / (Val(Me.txtCantidad.Text) * Val(Me.TxtUnidades.Text)), "###0.00")
                End If
                Me.txtprecioventa.Text = Format(totalneto - Val(Me.TxtTotalDescuento.Text), "###0.00")
                
       
    End If
    FrmDescuentos.Procedencia = vacio
    Call Resalta(Me.txtOtros)
    If KEY_CON_IGV = "si" Then
        Me.chkigv.Value = 1
    End If
End If
End Sub

Private Sub TxtFecha_emision_KeyPress(KeyAscii As Integer)
On Error GoTo salir
If KeyAscii = 13 Then
     If IsDate(Trim(Me.TxtFecha_emision.Text)) = False Then
        Me.TxtFecha_emision.Text = CVDate(KEY_FECHA)
        
     End If
    Me.DtpKardex.Value = Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd")
    Me.txtTc.Text = cambio_venta(CVDate(Me.TxtFecha_emision.Text))
    Me.txtFecha_vencimiento.SetFocus
End If
Exit Sub
salir:
    MsgBox "Ingrese Una Fecha Correcta", vbInformation, KEY_USUARIO
    Exit Sub
End Sub

Private Sub txtFecha_vencimiento_KeyPress(KeyAscii As Integer)
On Error GoTo salir
If KeyAscii = 13 Then
    If IsDate(Trim(Me.txtFecha_vencimiento.Text)) = False Then
        Me.txtFecha_vencimiento.Text = CVDate(Me.TxtFecha_emision.Text)
        If Me.DtcPeriodo.Enabled = True Then
            Me.DtcPeriodo.SetFocus
        End If
        Exit Sub
    Else
        If Me.DtcPeriodo.Enabled = True Then
            Me.DtcPeriodo.SetFocus
        End If
    End If
End If
Exit Sub
salir:
MsgBox "Ingrese Una Fecha Correcta", vbInformation, KEY_USUARIO
    Exit Sub
End Sub

Private Sub TxtFlete_Change()
Me.TxtCif.Text = Val(Me.txtFob.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtSeguro.Text)
If Val(Me.txtIdCompra.Text) < 1 Then
    If Me.DtTipoCompra.BoundText = "01" Then ' IMPORTACION
       Me.lblTotal.Text = Val(Me.txttotal_final.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtSeguro.Text)
       Me.LblValorVenta.Text = Val(Me.lblTotal.Text)
       Me.lblIMPBruto.Text = Val(Me.lblTotal.Text)
       Me.LblIgv.Text = 0
    End If
    
End If


End Sub

Private Sub txtfob_Change()
Me.TxtCif.Text = Val(Me.txtFob.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtSeguro.Text)



End Sub

Private Sub TxtIGV_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim igv_t As Single
    Me.TxtIgv.Text = Val(Me.TxtIgv.Text)

    Call Resalta(Me.txtIgv_Porcentaje)
End If
End Sub

Private Sub TxtISC_KeyPress(KeyAscii As Integer)
On Error GoTo salir
Dim isc_v As Single
Dim tprecioventa As Single
Dim igv_v As Single
If KeyAscii = 13 Then
    
    
    isc_v = Val(Me.txtisc.Text)
    If isc_v > 0 Then
     tprecioventa = Val(Me.TxtValoerNeto.Text) - Val(Me.TxtTotalDescuento.Text)
    igv_v = Me.TxtIgv.Text
    Me.txtisc.Text = Format(isc_v, "###0.00")
    
    
    If (Me.chkigv.Value = 1) Then
        Me.TxtIgv.Text = Format(igv_v + KEY_IGV * isc_v, "###0.00")
        Me.txtprecioventa.Text = Format(tprecioventa + isc_v + KEY_IGV * isc_v + igv_v, "###0.00")
        Me.TxtISC_p.Text = Format(isc_v * 100 / Val(Me.txtprecioventa.Text), "###0.00")
    Else
        
        Me.txtprecioventa.Text = Format(tprecioventa + isc_v, "###0.00")
        If (Val(tprecioventa) > 0) Then
            Me.TxtISC_p.Text = Format(isc_v * 100 / Val(tprecioventa), "###0.00")
        Else
            Me.TxtISC_p.Text = Format(isc_v * 100, "###0.00")
        End If
    End If
    End If
    
    Call Resalta(Me.TxtISC_p)
End If
Exit Sub
salir: MsgBox "Debe Ingresar un Valor Positivo"
End Sub

Private Sub TxtISC_p_KeyPress(KeyAscii As Integer)
Dim valorneto As Single
Dim isc_v As Single
Dim isc_p As Single
If KeyAscii = 13 Then
    
    valorneto = Val(Me.TxtValoerNeto.Text) - Val(Me.TxtTotalDescuento.Text)
    isc_p = Val(Me.TxtISC_p.Text)
    If isc_p > 0 Then
    isc_v = Format(valorneto * isc_p / 100, "###0.00")
    Me.txtisc.Text = isc_v
    igv_v = Me.TxtIgv.Text
        
    If (Me.chkigv.Value = 1) Then
        Me.TxtIgv.Text = Format(igv_v + KEY_IGV * isc_v, "#,#00.000")
        Me.txtprecioventa.Text = Format(valorneto + isc_v + KEY_IGV * isc_v + igv_v, "###0.00")
    Else
        
        Me.txtprecioventa.Text = Format(valorneto + isc_v, "###0.00")
    End If
    End If
    Call Resalta(Me.txtRetencion)
    
    
End If
End Sub

Private Sub txtitemdua_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtitemdua.Text = Format(Trim(Me.txtitemdua.Text), "0000")
    Me.cmdProcesar.SetFocus
End If
End Sub

Private Sub txtMonto_porcentaje_Change()
If Val(Me.txtMonto_porcentaje.Text) > 0 Then
   Me.txtMonto_asignado.Text = Val(Me.txtMonto_porcentaje.Text) * Val(Me.TxtMontoTotal_vinculado.Text) / 100
End If
End Sub

Private Sub txtMontoProrrateo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
       strCadena = "SELECT * FROM  movimiento_compra  WHERE id_compra='" & Val(Me.txtIdCompra.Text) & "' and ruc='" & KEY_RUC & "'"
       Call ConfiguraRstA(strCadena)
       If rstA.RecordCount > 0 Then
           If rstA("prorrateo_gastos") = "no" Then
            strCadena = "UPDATE movimiento_compra_detalle SET incremento_neto_gasto='" & Val(Me.txtMontoProrrateo.Text) & "'  WHERE id_compra='" & Val(Me.txtIdCompra.Text) & "' and  id_detalle_compra='" & Val(Me.txtid_detallecompra.Text) & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            Call Me.llenarGrid_Comprobante(Me.HfdDetalle, Val(Me.txtIdCompra.Text))
            Me.frmProrrateo.Visible = False
            Else
                MsgBox "ESTE GASTO YA FUE PRORRATEADO EN EL KARDEX" + Chr(13) + "COORDINE CON EL AREA DE SISTEMAS"
            End If
       End If
    
    
End If
End Sub

Private Sub txtNumero_orden_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtNumero_orden.Text = Format(Me.txtNumero_orden.Text, "00000000")
    'Me.TxtSerie_orden.Text = Format(Me.TxtSerie_orden.Text, "0000")
    If MsgBox("Esta Seguro de realizar la Operacion", vbYesNo + vbQuestion, KEY_VENDEDOR) = vbYes Then
        Dim in_igv As String
        strCadena = "DELETE FROM movimiento_compra_temporal WHERE dni_save='" & Trim(KEY_USUARIO) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
    
        strCadena = "SELECT * FROM movimiento_compra WHERE  id_doc='" & Me.DtcComprobante_orden.BoundText & "' and  serie='" & Trim(Me.TxtSerie_orden.Text) & "' and numero='" & Trim(Me.txtNumero_orden.Text) & "' and id_proveedor='" & Trim(Me.txtRuc.Text) & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
         If rst("igv") > 0 Then
            in_igv = "si"
         Else
            in_igv = "no"
         End If
         
         
         Me.txtRuc.Text = rst("id_proveedor")
         Me.TxtProveedor.Text = get_persona(rst("id_proveedor"))
         
         
        strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & rst("id_compra") & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            For i = 0 To rst.RecordCount - 1
               
                
                strCadena = "INSERT INTO movimiento_compra_temporal(id_doc,serie,numero,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento," & _
                "valor_neto,isc,igv,retencion,otros,percepcion,exonerado,valor_venta,precio_venta,p_venta,p_costo,dni_save,id_alm,detalle,ruc) VALUES " & _
                "('" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Trim(Me.txtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & rst("id_producto") & "'," & _
                "'" & rst("cantidad") & "','" & rst("c_unitario") & "','0'," & _
                "'0','0','" & rst("valor_venta") & "','0','" & rst("igv") & "'," & _
                "'0','0','0','" & rst("exonerado") & "','" & rst("valor_venta") & "','" & rst("total") & "','" & get_precio_venta_hoy(rst("id_producto")) & "'," & _
                "'" & rst("p_costo") & "','" & KEY_USUARIO & "','" & Trim(Me.DtcAlmacen.BoundText) & "','" & get_producto(rst("id_producto")) & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
     
     
     
     
    
     
    
    
                rst.MoveNext
            Next i
                Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcTipoDoc.BoundText, Me.txtSerie.Text)
        End If
        End If
    End If
    
    
    
End If
End Sub
Private Function get_precio_venta_hoy(ByVal in_producto As String) As Single
strCadena = "SELECT * FROM almacen_producto where id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   get_precio_venta_hoy = rstP("precio_venta")
End If
End Function

Private Sub TxtRetencion_KeyPress(KeyAscii As Integer)
Dim totalneto As Single
totalneto = Val(Me.txtprecioventa.Text)
If KeyAscii = 13 Then
    Me.txtRetencion.Text = Format(Val(Me.txtRetencion.Text), "###0.00")
    Me.txtprecioventa.Text = Format(totalneto + Val(Me.txtRetencion.Text), "###0.00")
    Call Resalta(Me.txtprecioventa)
End If
End Sub
Public Sub Ordencompra()
Dim fecha_inicial As Date
Dim fecha_final As Date
Dim Orden As Double
' Me.TxtNumeroG.Text = formato_item(Me.TxtNumeroG.Text, 10)
'  strCadena = "SELECT * FROM OrdenCompra WHERE serie='" & Trim(Me.TxtSerieG.Text) & "' AND numero='" & Trim(Me.TxtNumeroG.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        strCadena = "SELECT * FROM MiTransporte WHERE id_transporte='" & Val(rst("id_transporte")) & "' AND Ruc='" & KEY_RUC & "'"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
           ' Me.TxtTransporte.Text = rstT("marca") + Space(1) + "PLACA:" + rstT("placa")
            If IsNull(rst("id_Cisterna")) = False Then
                strCadena = "SELECT * FROM MiTransporte WHERE id_transporte='" & Val(rst("id_Cisterna")) & "' AND Ruc='" & KEY_RUC & "'"
                Call ConfiguraRstT(strCadena)
                If rstT.RecordCount > 0 Then
                    'Me.TxtTransporte.Text = Me.TxtTransporte.Text + "/" + rstT("placa")
                End If
            End If
            
            strCadena = "SELECT * FROM Persona WHERE cPersona='" & rst("cConductor") & "'"
            Call ConfiguraRstT(strCadena)
            If rst.RecordCount > 0 Then
               ' Me.txtchofer.Text = rstT("NombrePersona")
               ' Me.TxtLicencia.Text = rstT("licencia")
            End If
            strCadena = "SELECT * FROM Ocurrencias WHERE idOrden='" & rst("idOrden") & "' AND idsalidaingreso='1'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
                  '  Me.txtInicial.Text = Format(rstT("kilometraje"), "###0.00")
                  '  fecha_inicial = CVDate(rstT("fecha"))
            End If
            strCadena = "SELECT * FROM Ocurrencias WHERE idOrden='" & rst("idOrden") & "' AND idsalidaingreso='2'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
                  '  Me.txtfinal.Text = Format(rstT("kilometraje"), "###0.00")
                  '  fecha_final = CVDate(rstT("fecha"))
                  '  Me.txtdias.Text = DateDiff("d", fecha_inicial, fecha_final)
                  '  Me.Txtrecorrido.Text = Format(Val(Me.txtfinal.Text) - Val(Me.txtInicial.Text), "###0.00")
                     If rever = True Then
                        '  descuentoRecorrido = Val(Me.Txtrecorrido.Text) / Val(Me.txtcantidad.Text)
                        '  Me.TxtCostoHoy.Text = Format(Val(Me.TxtCostoAnt.Text) + descuentoRecorrido, "###0.00")
                    End If
            End If
          
        End If
        Orden = rst("idorden")
        'Call llenarGridUsuarios(Me.HfPersonal, Orden)
        'Call llenarGridGastos(Me.HfGastos, Orden)
    Else
        MsgBox "Orden de Compra no registrada", vbInformation, Mid(KEY_EMPRESA, 1, 35)
        Exit Sub
    End If
End Sub






Private Sub txtnumero_dua_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtAnioDua)
End If
End Sub

Private Sub TxtNumeroG_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call Ordencompra
End If
End Sub
Private Sub llenarGridUsuarios(ByVal Grilla As MSHFlexGrid, ByVal idOrden As Double)
Dim sueldodia As Single
On Error GoTo salir
strCadena = "SELECT Persona.cPersona, Persona.NombrePersona, Persona.sueldo_mensual FROM  OcurrenciasPersonal INNER JOIN " & _
"Persona ON OcurrenciasPersonal.cPersona = Persona.cPersona INNER JOIN " & _
"Ocurrencias ON OcurrenciasPersonal.idOcurrencia = Ocurrencias.idOcurrencia WHERE idOrden='" & idOrden & "' AND idSalidaingreso='1'"
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
           Grilla.ColWidth(0) = 500
           Grilla.ColWidth(1) = 2500
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 500
           Grilla.ColWidth(4) = 800
       Next
         cabecera = "COD" & vbTab & "DESCRIPCION" & vbTab & "SUELDO" & vbTab & "DIAS" & vbTab & "TOTAL"
         Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
        sueldodia = 0
        rst.MoveFirst
        For i = 1 To rst.RecordCount
   '          Fila = Fila & rst("cPersona") & vbTab & rst("NombrePersona") & vbTab & rst("sueldo_mensual") & vbTab & Val(Me.txtdias.Text) & vbTab & Format(rst("sueldo_mensual") / 30 * Val(Me.txtdias.Text), "###0.00")
            If (Fila = "") Then
                X = 1
            End If
    '        sueldodia = sueldodia + rst("sueldo_mensual") / 30 * Val(Me.txtdias.Text)
           Grilla.AddItem Fila

        Fila = ""
        rst.MoveNext
        Next i
         Fila = "" & vbTab & "TOTAL PAGO PERSONAL" & vbTab & "" & vbTab & "" & vbTab & "" & Format(sueldodia, "###0.00")
         Grilla.AddItem Fila
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &HC0FFFF
        Next k
        If rever = True Then
     '   descuentoPersonal = sueldodia / Val(Me.TxtCantidad.Text)
      '  Me.TxtCostoHoy.Text = Format(Val(Me.TxtCostoAnt.Text) + descuentoRecorrido + descuentoPersonal, "###0.00")
        End If
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub llenarGridGastos(ByVal Grilla As MSHFlexGrid, ByVal idOrden As Double)
Dim totalGasto As Single
On Error GoTo salir
strCadena = "SELECT     (Comprobantes.doc_abrev +':'+ SolicitudViaticosComprobantes.serie +'-'+ SolicitudViaticosComprobantes.numero) AS comprobante, " & _
" SolicitudViaticosComprobantes.total FROM SolicitudViaticos INNER JOIN SolicitudViaticosComprobantes ON SolicitudViaticos.id_Solicitud = SolicitudViaticosComprobantes.idSolicitud INNER JOIN " & _
" Comprobantes ON SolicitudViaticosComprobantes.doc_cod = Comprobantes.doc_cod WHERE idOrden='" & idOrden & "' AND  Ruc='" & KEY_RUC & "'"
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
           Grilla.ColWidth(0) = 500
           Grilla.ColWidth(1) = 2500
           Grilla.ColWidth(2) = 800
         
          
       Next
         cabecera = "COD" & vbTab & "COMPROBANTE" & vbTab & "MONTO"
         Grilla.AddItem cabecera
         For k = 0 To 2
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
        totalGasto = 0
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = Fila & i & vbTab & rst("comprobante") & vbTab & Format(rst("total"), "###0.00")
            If (Fila = "") Then
                X = 1
            End If
            totalGasto = totalGasto + rst("total")
           Grilla.AddItem Fila

        Fila = ""
        rst.MoveNext
        Next i
         Fila = "" & vbTab & "" & vbTab & "" & Format(totalGasto, "###0.00")
         Grilla.AddItem Fila
         For k = 0 To 2
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &HC0FFFF
        Next k
        If rever = True Then
        descuentoViaticos = totalGasto / Val(Me.txtCantidad.Text)
        Me.TxtCostoHoy.Text = Format(Val(Me.TxtCostoAnt.Text) + descuentoRecorrido + descuentoPersonal + descuentoViaticos, "###0.00")
        End If
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub TxtNumeroR_KeyPress(KeyAscii As Integer)
Dim cPersona As String
If KeyAscii = 13 Then
    Me.TxtNumeroR.Text = formato_item(Me.TxtNumeroR.Text, 8)
    
    strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Trim(Me.DtcRelacionado.BoundText) & "' AND serie='" & Trim(Me.TxtSerieR.Text) & "' AND  numero='" & Trim(Me.TxtNumeroR.Text) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.txtRuc.Text = rst("id_proveedor")
        Me.TxtFechaEmision.Text = str(rst("fecha_emision"))
        strCadena = "SELECT DISTINCT id_proveedor as Codigo,nproveedor as Descripcion FROM movimiento_compra WHERE id_doc='" & Trim(Me.DtcRelacionado.BoundText) & "' AND serie='" & Trim(Me.TxtSerieR.Text) & "' AND  numero='" & Trim(Me.TxtNumeroR.Text) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcProveedor)
        strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
            Me.TxtProveedor.Text = UCase(rstT("nombre_completo"))
            Me.txtDireccion.Text = UCase(rstT("direccion"))
        End If
    Else
        Me.DtcProveedor.Text = "COMPROBANTE NO REGISTRADO"
    End If
    Set rst = Nothing
   ' Me.TxtCodProducto.SetFocus
End If
End Sub

Private Sub txtnumeroserie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Me.txtseriechasis.Text = Trim(Me.txtnumeroserie.Text)
    Call Resalta(Me.txtseriemotor)
    
    Exit Sub
    
End If
End Sub

Private Sub txtOtros_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.txtOtros.Text = Format(Val(Me.txtOtros.Text), "###0.00")
 If Me.chkigv.Value = 0 Then
    Me.txtprecioventa.Text = Format(Val(Me.txtprecioventa.Text) + Val(Me.txtOtros.Text), "###0.00")
Else
    Me.txtprecioventa.Text = Format(Val(Me.txtprecioventa.Text) + (Val(Me.txtOtros.Text)) * (KEY_IGV + 1), "###0.00")
 End If
    Call Resalta(Me.TxtValoerNeto)
   
End If
End Sub

Private Sub TxtPecepcion_Change()
Dim valortotal As Single
Dim total_percepcion As Single
Dim percepcion As Single
valortotal = Val(Me.lblTotal.Text)
percepcion = Val(Me.TxtPecepcion.Text)
Me.lblPercepcion.Text = Format(valortotal + percepcion, "###0.00")

If Me.chkigv.Value = 1 And Val(Me.txtCantidad.Text) > 0 Then
    costo = (Val(Me.TxtUnitario.Text) - Val(Me.TxtTotalDescuento.Text) / Val(Me.txtCantidad.Text)) * (1 + KEY_IGV) + Val(percepcion) / Val(Me.txtCantidad.Text)
Else
   If Val(Me.txtCantidad.Text) > 0 Then
     costo = (Val(Me.TxtUnitario.Text) - Val(Me.TxtTotalDescuento.Text) / Val(Me.txtCantidad.Text)) + Val(percepcion) / Val(Me.txtCantidad.Text)
   End If
End If

Me.TxtCostoHoy.Text = Format(costo, "###0.00")

End Sub

Private Sub TxtPecepcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
End If
End Sub

Private Sub txtPorcentajeGastos_Change()
Dim costo As Single
Dim venta As Single
Dim utilidad As Single
Dim venta_actual As Single
Me.TxtGastoAdminHoy.Text = Format(Val(Me.TxtCostoHoy.Text) * Val(Me.txtPorcentajeGastos.Text) / 100, "###0.00")

costo = Val(Me.TxtCostoHoy.Text)
utilidad = Val(Me.txtUtilidadhoy.Text)
venta = Val(Me.TxtGastoAdminHoy.Text) + costo + costo * utilidad / 100

If rever = False Then
    Me.TxtventaHoy.Text = Format(venta, "###0.00")
Else
    If Val(Me.TxtCostoHoy.Text) > 0 Then
    Me.TxtventaHoy.Text = Format(Val(Me.TxtCostoHoy.Text) + Val(Me.TxtCostoHoy.Text) * Val(Me.txtUtilidadhoy.Text) / 100, "#,##0.00")
    'Me.txtUtilidadhoy.text = Format(utilidad, "###0.00")
    End If
End If

End Sub

Private Sub txtPorcentajeGastos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtUtilidadhoy)
End If
End Sub

Private Sub TxtPrecioVenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If KEY_CON_IGV = "si" Then
        If Me.DtcTipoDoc.BoundText = "0002" Then
                Me.chkigv.Value = 0
            Else
                Me.chkigv.Value = 1
        End If
        
    Else
        If Me.cmdRevertir.Visible = False Then
            Me.chkigv.Value = 0
        Else
            Call calculo_igv
        End If
    End If
    
    If Me.DtcTipoDoc.BoundText = "0200" Or Me.DtcTipoDoc.BoundText = "0003" Then
       Me.chkigv.Value = 0
    End If
    
    If Val(Me.TxtUnidades.Text) > 1 Then
        Me.chkModo.Value = 1
    End If
    
    If Val(Me.txtprecioventa.Text) = 0 Then
        Me.TxtCostoHoy.Text = Format(Val(Me.TxtCostoAnt.Text), "#,##0.00")
        Me.TxtventaHoy.Text = Format(Val(Me.txtPrecioVentaAnt.Text), "###0.00")
    End If
    Me.TxtventaHoy.Text = Format(Val(Me.txtPrecioVentaAnt.Text), "###0.00")
    
    
    
   
    Call Resalta(Me.txtPorcentajeGastos)
    
    
    
End If




End Sub

Private Sub TxtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.txtRuc)
End If

End Sub

Private Sub TxtProveedor_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
        Call Resalta(Me.txtDireccion)
                
End If
End Sub

Private Sub TxtCodProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtNumeroDoc)
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.TxtProveedor)
End If
End Sub
Private Sub buscarcliente()
If (Trim(Me.txtRuc.Text) = "") Then
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
End If


    strCadena = "SELECT *  FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.txtRuc.Text = Trim(Me.txtRuc.Text)
        FrmDetallePersona.chkProveedor.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.txtRuc.Text = rst("dni")
        Me.TxtProveedor.Text = rst("nombre_completo")
        Me.txtDireccion.Text = rst("direccion")
        Me.DtTipoCompra.SetFocus
        strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Me.txtSerie.Text & "' AND numero='" & Me.TxtNumeroDoc.Text & "' AND id_proveedor='" & Trim(Me.txtRuc.Text) & "' ANd ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            MsgBox "Documento ya Registrado. ", vbInformation, KEY_EMPRESA
            Call buscar_comprobante(rst("id_compra"))
            Exit Sub
        End If
        Exit Sub
       
    End If

End Sub

Private Sub TxtcodProveedor_KeyPress(KeyAscii As Integer)


End Sub

Private Sub TxtCodProducto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.txtCantidad)
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.txtObservacion)
End If
End Sub

Private Sub TxtCodProducto_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then
    If Trim(Me.TxtCodProducto.Text) = "" Then
     Procedencia = Selecionar
        FrmProducto.Show
     Exit Sub
     End If
    Me.TxtDctoSoles.Text = 0
    Me.TxtDstoporcentaje.Text = 0
    Me.TxtIgv.Text = 0
    Me.txtIgv_Porcentaje.Text = 0
    Me.TxtValoerNeto.Text = 0
    Me.txtprecioventa.Text = 0
    Me.txtisc.Text = 0
    Me.TxtISC_p.Text = 0
    Me.txtRetencion.Text = 0
    Me.txtOtros.Text = 0
    Me.txtValorVenta.Text = 0
    Me.TxtventaHoy.Text = 0#
    Me.TxtTotalDescuento.Text = 0
    Me.TxtUnitario.Text = 0
    Me.TxtCostoHoy.Text = 0
    Me.TxtventaHoy.Text = 0
    Me.TxtCodProducto.Text = formato_item(Trim(Me.TxtCodProducto.Text), 5)
    Call put_producto
    
    
End If
End Sub

Public Sub put_producto()
On Error GoTo nsalir


If KEY_PAIS <> KEY_PERU And (Me.DtcTipoDoc.BoundText = "0020" Or Me.DtcTipoDoc.BoundText = "0427") Then
        Me.FRameiva.Visible = True
    End If
    
    
    If Me.chkigv.Value = 1 Then
        Me.chkigv.Value = 0
    End If
    
    If KEY_BARRAS = "si" Then
        strCadena = "SELECT B.cod_barra,P.nombre_prod,P.precio_compra,P.precio_venta,U.abreviatura,B.id_producto,P.id_unidad FROM producto_barras B,producto P,unidad U WHERE B.id_producto=P.id_producto AND B.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND B.cod_barra='" & Trim(Me.TxtCodProducto.Text) & "'"
    Else
      strCadena = "SELECT A.id_producto,P.nombre_prod,U.descripcion as abreviatura,A.stock,A.precio_compra,A.precio_venta,P.agranel,P.numero_procedimientos,P.id_unidad FROM almacen_producto A,producto P,unidad U WHERE A.id_producto=P.id_producto AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND A.id_alm='" & Trim(FrmCompras.DtcAlmacen.BoundText) & "' AND A.id_producto='" & Trim(Me.TxtCodProducto.Text) & "'"
    End If
    
    
    
    Call ConfiguraRst(strCadena)
    
    
    
    
    If rst.RecordCount > 0 Then
        codigoP = rst("id_producto")
        Me.TxtDescripcionProducto.Text = UCase(rst("nombre_prod"))
        Me.txtcosto.Text = rst("precio_compra")
        Me.txtPrecioVentaAnt.Text = rst("precio_venta")
        Me.TxtCostoAnt.Text = rst("precio_compra")
        Me.TxtUnidades.Text = rst("numero_procedimientos")
        If Val(Me.TxtUnidades.Text) < 1 Then
            Me.TxtUnidades.Text = 1
        End If
        
        If rst("precio_compra") > 0 Then
            Me.TxtUtilidadAnt.Text = Format((rst("precio_venta") - rst("precio_compra")) * 100 / rst("precio_compra"), "#,##0.00")
        End If
        
       Call get_unidad(rst("id_producto"), rst("agranel"))
        
       
       
       Call get_cantidad_agranel_ini(Trim(Me.TxtCodProducto.Text), Me.DtcUnidad.BoundText)
       
       
       Me.txtCantidad.Text = 1
       Call Resalta(Me.txtCantidad)
       
        
    Else
        Procedencia = Selecionar
        FrmProducto.Show
    End If
    
   Exit Sub
   
nsalir:
   MsgBox "Configure, el producto Correctamente", vbInformation
    
    
End Sub




Private Sub TxtDireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Me.DtTipoCompra.SetFocus
End If
End Sub


Private Sub Llenar_Temporal()
'Dim RstTemporal As New ADODB.Recordset
'Dim rstDetalle As New ADODB.Recordset
'Dim i As Integer
'StrCadena = "SELECT * FROM Detalle_DocumentoCompra WHERE (cDocumentoCompra='" & Trim(Me.TxtNumero_guia.Text) & "' AND doc_cod='" & Trim(Me.DtcComprobanteGuia.BoundText) & "' AND " & _
'"sSerie='" & Trim(Me.TxtSeri_guia.Text) & "') "
'rstDetalle.Open StrCadena, CnBd, adOpenKeyset, adLockOptimistic
'StrCadena = "SELECT * FROM Temporal_Compras"
'RstTemporal.Open StrCadena, CnBd, adOpenKeyset, adLockOptimistic
'rstDetalle.MoveFirst

 'For i = 0 To rstDetalle.RecordCount - 1
  ' StrCadena = "SELECT cTemporal FROM Temporal_Compras ORDER BY cTemporal DESC "
   ' RstTemporal.AddNew
    'RstTemporal.Fields(0) = GeneraCodTemporal
    'RstTemporal.Fields(1) = Trim(Me.TxtNumeroDoc.Text)
    ''RstTemporal.Fields(2) = Trim(Me.DtcTipoDoc.BoundText)
    'RstTemporal.Fields(3) = Trim(Me.TxtSerie.Text)
    'RstTemporal.Fields(4) = rstDetalle.Fields(4)
    'RstTemporal.Fields(5) = rstDetalle.Fields(5)
    'RstTemporal.Fields(6) = rstDetalle.Fields(6)
    'RstTemporal.Fields(7) = rstDetalle.Fields(7)
    'RstTemporal.Update
    'rstDetalle.MoveNext
       
 'Next i
'Set RstTemporal = Nothing
'Set rstDetalle = Nothing
End Sub

Private Sub TxtNumeroDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       Me.TxtNumeroDoc.Text = formato_item(Me.TxtNumeroDoc.Text, 8)
       Me.TxtFecha_emision.SetFocus
        'Call llenarGrid_det(Me.HfdDetalle, "", "", "")
End If

End Sub


Public Sub buscar_comprobante(Optional idCompra As Double)
On Error GoTo validar
    Me.txtIdCompra.Text = idCompra
    Me.TxtNumeroDoc.Text = Format(Me.TxtNumeroDoc.Text, "00000000")
    strCadena = "SELECT * FROM movimiento_compra WHERE (numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND serie='" & Trim(Me.txtSerie.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND ruc='" & KEY_RUC & "' AND id_proveedor='" & Trim(Me.txtRuc.Text) & "')"
    If FrmBusquedaCompras.Procedencia = buscar Then
        strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & idCompra & "' AND ruc='" & KEY_RUC & "'"
        FrmBusquedaCompras.Procedencia = Neutro
    End If
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        strCadena = "SELECT * FROM movimiento_compra_temporal WHERE (serie='" & Trim(Me.txtSerie.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
        Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcTipoDoc.BoundText, Me.txtSerie.Text)
        
        Me.cmdsave.Enabled = True
        End If
        Set rst = Nothing
        
        
        
        'Me.TxtRuc.text = "00000000"
        Me.TxtProveedor.Text = "PUBLICO EN GENERAL"
        Me.txtDireccion.Text = KEY_DIR_PUBLIC
        If Trim(Me.DtcTipoDoc.BoundText) = "0007" Then
            Me.DtcRelacionado.SetFocus
        Else
            Call Resalta(Me.txtRuc)
        End If
        'ProcendenciaGuia = NuevaGuia
  
    Else
        
        
        
     If FrmListadoFacturasCompra.Procedencia = buscar Then
        FrmListadoFacturasCompra.Procedencia = Neutro
        GoTo saltara
     End If
     
    ' If MsgBox("Comprobante ya Registrado !!" + Chr(13) + Chr(13) + "SI  ..... Para Visualizarlo" + Chr(13) + "NO ..... Para Crear Nuevo Comprobante", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
       If idCompra = 0 Then
            idCompra = rst("id_compra")
        End If
       Me.lblidCompra.Caption = rst("id_compra")
       If rst("anulado") = "si" Then
            Me.lblAnulado.Visible = True
        Else
            Me.lblAnulado.Visible = False
        
        End If
saltara:
        Me.DtcAlmacen.BoundText = rst("id_alm")
        Me.DtTipoCompra.BoundText = rst("id_tipo_compra")
        Me.DtTipoCompra.Locked = True
        Me.DtcTipo.BoundText = rst("id_tipo")
        
        If Me.DtTipoCompra.BoundText = "01" Then
           If rst("prorrateo_importacion") = "no" Then
              Me.cmdProrrateoImportacion.Enabled = True
              Me.chkValorVenta_importacion.Visible = True
              Me.chk_cantidad_importacion.Visible = True
              Me.chkValorVenta_importacion.Enabled = True
              Me.chk_cantidad_importacion.Enabled = True
              Me.cmdAplicarImportacion.Enabled = True
              Me.cmdAplicarImportacion.Enabled = True
              
              
           Else
              Me.cmdProrrateoImportacion.Enabled = False
              Me.chkValorVenta_importacion.Enabled = False
              Me.chk_cantidad_importacion.Enabled = False
              Me.cmdAplicarImportacion.Enabled = False
           End If
            
        End If
        If rst("prorrateo_gastos") = "no" Then
           Me.cmdProrrateoGastos.Enabled = True
           Me.chk_valor_venta_gasto.Visible = True
           Me.chk_cantidad_gastos.Visible = True
           Me.chk_valor_venta_gasto.Enabled = True
           Me.chk_cantidad_gastos.Enabled = True
           
        Else
           Me.cmdProrrateoGastos.Enabled = True
           Me.chk_valor_venta_gasto.Visible = True
           Me.chk_cantidad_gastos.Visible = True
           Me.chk_valor_venta_gasto.Value = 1
        End If
        Me.cmdAplicarGastos.Enabled = True
        Me.DtcMoneda.BoundText = rst("id_moneda")
        Me.DtcMoneda.Locked = True
        Me.txtTc.Text = rst("tc")
        Me.DtcResponsable.BoundText = rst("id_responsable")
        Me.txtObservacion.Text = rst("observacion")
        
        
        
       
        
        '----- datos dua
        Me.txtA�oFabricacion.Text = rst("anio_fabricacion")
        Me.txtA�oFabricacion.Locked = True
        
        Me.txtseriedua.Text = rst("serie_dua")
        Me.txtnumero_dua.Text = rst("numero_dua")
        
        Me.txtnumero_dua.Locked = True
        Me.TxtAnioDua.Text = rst("anio_dua")
        Me.TxtAnioDua.Locked = True
        txta�omodelo.Text = rst("anio_modelo")
        Me.txta�omodelo.Locked = True
        Me.txtFob.Text = rst("fob")
        Me.TxtFlete.Text = rst("flete")
        Me.TxtCif.Text = rst("cif")
        Me.txtSeguro.Text = rst("seguro")
        Me.cmdgastos.Enabled = True
        Me.TxtAlmacen.Text = rst("id_almacen")
        Me.txtNumeroAutorizacion.Text = rst("autorizacion")
        '------ fin datos dua
        
        Me.dtpFechaRegistro.Value = rst("fecha_registro")
        Me.TxtFecha_emision.Text = CVDate(rst("fecha_emision"))
        Me.DtpKardex.Value = rst("fecha_kardex")
        Me.txtFecha_vencimiento.Text = CVDate(rst("fecha_cancelacion"))
        If rst("percepcion") > 0 Then
            Me.ChkPercepcion.Value = 1
            Me.TxtPecepcion.Text = Format(rst("percepcion"), "###0.00")
        Else
            Me.ChkPercepcion.Value = 0
        End If
        
        If rst("retencion") > 0 Then
           Me.TxtTotalRetencion.Text = Format(rst("retencion"), "###0.00")
           Me.chk_suspencion_retencion.Value = 0
        Else
          Me.TxtTotalRetencion.Text = 0
          Me.chk_suspencion_retencion.Value = 1
        End If
        
        If Me.DtcTipoDoc.BoundText = "0008" Then
          Call load_tipo_debito
        
        End If
        
        Me.DtcTipoNota.BoundText = rst("id_tipo_nota")
        Me.DtcPeriodo.BoundText = rst("id_periodo")
        
        
        Call LlenarDatosCliente(Me.txtRuc.Text)
        
        Call llenarGrid_Comprobante(Me.HfdDetalle, idCompra)
        Call VerificaAnulado(idCompra)
        Call otros_gastos(idCompra)
       
        
        
        
        
        Me.Txtdoc_cod.Text = Me.DtcTipoDoc.BoundText
        Me.TxtSeriaGuardada.Text = Me.txtSerie.Text
        Me.TxtNumeroGuardado.Text = Me.TxtNumeroDoc.Text
        Me.TxtAlmacenGuardado.Text = Me.DtcAlmacen.BoundText
        Me.TxtProveedorGuardado.Text = Me.txtRuc.Text
        Me.DtcTipoDoc.Enabled = False
        Me.txtSerie.Enabled = False
        Me.TxtNumeroDoc.Enabled = False
        Me.TxtCodProducto.Enabled = False
        Me.TxtDescripcionProducto.Enabled = False
        Me.cmdAgregar.Enabled = False
        Me.CmdQuitar.Enabled = False
        
        If KEY_RUC = "20480516771" Then
        
        If Me.DtcTipoDoc.BoundText = "0089" Or Me.DtcTipoDoc.BoundText = "0090" Then
          Me.cmdEliminar.Enabled = False
        Else
          Me.cmdEliminar.Enabled = True
        End If
        
        
        End If
        Me.cmdModificar.Enabled = True
        Me.txtCantidad.Enabled = False
        Me.cmdgastos.Visible = True
        Me.cmdgastos.Enabled = True
        Me.lblidCompra.Caption = get_periodo_detalle(Me.DtcPeriodo.BoundText, idCompra)
    
    
    End If
    
    
    Exit Sub
validar:
        MsgBox "Ocurrio un problema en la lectura del comprobante", vbInformation
End Sub
Public Sub llenar_totales(ByVal in_compra As String)
strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & in_compra & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.lblExonerado.Text = Format(rst("exonerado"), "###0.00")
        Me.LblValorVenta.Text = Format(rst("valor_venta"), "###0.00")
        Me.lblISC.Text = Format(rst("isc"), "###0.00")
        Me.LblIgv.Text = Format(rst("igv"), "###0.00")
         
        'Me.lblDescuento.Text = Format(rst("descuento_global"), "###0.00")
        Me.lblISC.Text = Format(rst("isc"), "###0.00")
        TxtTotalRetencion.Text = Format(rst("retencion"), "###0.00")
        
        Me.lblTotal.Text = Format(rst("total"), "###0.00")
        Me.txttotal_final.Text = Val(Me.lblTotal.Text)
        Me.lblPercepcion.Text = Format(rst("percepcion") + Val(Me.lblTotal.Text), "###0.00")
        Me.txtObservacion.Text = rst("observacion")
        
    End If
End Sub
Public Sub otros_gastos(ByVal id_compra As Double)
Dim in_gasto_total As Double
Dim in_valor_venta As Double
Dim in_monto_parcial As Double

strCadena = "SELECT * FROM movimiento_compra_gasto WHERE id_compra='" & id_compra & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    
    in_valor_venta = 0
    For i = 0 To rstT.RecordCount - 1
        If rstT("id_moneda") = "00002" And rstT("id_moneda") <> KEY_MONEDA Then
           in_monto_parcial = rstT("monto") * rstT("tc")
        Else
           in_monto_parcial = rstT("monto")
        End If
        
        
        If KEY_CON_IGV = "si" Then
            If rstT("afecto_igv") = "si" Then
                in_valor_venta = in_valor_venta + in_monto_parcial / (1 + KEY_IGV)
            Else
                in_valor_venta = in_valor_venta + in_monto_parcial
            
            End If
        Else
                in_valor_venta = in_valor_venta + in_monto_parcial
        End If
        
        rstT.MoveNext
    Next i
End If




Me.lblgastos.Text = Format(in_valor_venta, "###0.00")






End Sub

Private Sub LlenarDatosCliente(ByVal in_ruc As String)

strCadena = "SELECT * FROM persona WHERE dni='" & in_ruc & "'"
Call ConfiguraRstT(strCadena)
    
        If rstT.RecordCount > 0 Then
        Me.TxtProveedor.Text = rstT("nombre_completo")
        Me.txtDireccion.Text = rstT("direccion")
        
    End If
  Set rstT = Nothing
End Sub
Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    If Me.TxtCodProducto.Enabled = True Then
        Call Resalta(Me.TxtCodProducto)
    End If
    
End If
End Sub



Private Sub TxtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtDescripcionProducto)
       
End If
If KeyCode = vbKeyRight Then
     Me.cmdAgregar.SetFocus
End If
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
Dim TotalP As Single
If KeyAscii = 13 Then
        
       ' TotalP = Val(Me.TxtCantidad.Text) * Val(Me.TxtPrecio.Text)
        'Me.LblTotalParcial.Text = Format(TotalP, "#,##0.000")
        Me.cmdAgregar.SetFocus
End If
End Sub
Private Sub load_tipo_nota()
strCadena = "SELECT id_tipo_nota as Codigo,CONCAT('[',id_tipo_nota,'] -',descripcion) as Descripcion FROM tipo_nota_credito WHERE id_tipo_nota IN('05','07','09') ORDER BY id_tipo_nota"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcTipoNota)
Me.frmtiponota.Visible = True
End Sub
Private Sub put_retencion_automatica(ByVal in_alm As String)
strCadena = "SELECT * FROM movimiento_venta WHERE  id_doc='0097' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY numero ASC LIMIT 1"
                    Call ConfiguraRstPP(strCadena)
                    If rstPP.RecordCount > 0 Then
                       in_numero = Format(Val(rstPP("numero")) + 1, "000000")
                       Documento = "RETENCION" & ":" & rstPP("serie") & "-" & in_numero
                   
                    
                    
                    strCadena = "P_insert_venta('0097','" & in_alm & "','0','" & Me.DtcMoneda.BoundText & "','no'," & _
                    "'" & rstPP("serie") & "','" & in_numero & "','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.TxtProveedor.Text) & "','0','0','0','" & Val(Me.TxtTotalRetencion.Text) & "','0'," & _
                    "'" & Val(Me.TxtTotalRetencion.Text) & "','0','" & Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd") & "','" & Format(Me.txtFecha_vencimiento.Text, "YYYY-mm-dd") & "','00001','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & Val(Me.txtTc.Text) & "','no','" & formato_item(Month(Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd")), 2) & "','" & Year(Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd")) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstP(strCadena)
                    
                    id_venta = rstP(0)
                    
                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & id_venta & "','00','RETENCION','-','1','" & Val(Me.TxtTotalRetencion.Text) & "','0','" & Val(Me.TxtTotalRetencion.Text) & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                               
                    
                    in_numero = Format(Val(in_numero) + 1, "000000")
                    strCadena = "UPDATE almacen_comprobante SET numero='" & in_numero & "' WHERE id_doc='0097' AND serie='" & rstPP("serie") & "' AND ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "INSERT INTO mis_cuentas_det_detalle(id_detalle,id_cuenta_det,monto_inicial,monto_pagado,id_movimiento,id_tipo)VALUES " & _
                    "('" & id_venta & "','0','" & Val(Me.TxtTotalRetencion.Text) & "','" & Val(Me.TxtTotalRetencion.Text) & "','" & Val(Me.txtIdCompra.Text) & "','02')"
                    CnBd.Execute (strCadena)
           
           
           
     End If
           
           
           
           
End Sub

Private Sub put_ajuste_ingreso(ByVal in_compra As String)
Dim in_asiento As String
Dim in_producto As String
Dim in_alm As String

strCadena = "SELECT CONCAT(p.id_producto,'-',p.nombre_prod) as glosa,d.id_producto,d.id_alm FROM movimiento_compra_detalle d,producto p WHERE d.id_producto=p.id_producto and d.ruc=p.ruc and d.id_compra='" & in_compra & "' and d.ruc='" & KEY_RUC & "'"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
    in_glosa = rstA("glosa")
    in_producto = rstA("id_producto")
    in_alm = rstA("id_alm")
Else
    in_glosa = "Aqui va la glosa"
    in_producto = ""
    in_alm = ""
End If

       
        strCadena = "SELECT ifnull(sum(cantidad_real),0) FROM kardex WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstZ(strCadena)

        strCadena = "UPDATE almacen_producto SET stock='" & rstZ(0) & "' WHERE id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        



End Sub

Private Sub Save()
'On Error GoTo salir
Dim i As Integer
Dim CodReferencia As String, cod_identidad As String * 1
Dim tSaldo As Double
Dim TotalFactura As Double
Dim ValorCompra As Double
Dim igv As Double
Dim Descuento As Single
Dim percepcion As Single
Dim rstVerifica As New ADODB.Recordset
Dim Ordencompra As Double
Dim in_saldo As Double

If IsDate(CVDate(Me.TxtFecha_emision.Text)) = False Then
   Me.TxtFecha_emision.Text = CVDate(KEY_FECHA)
End If

If IsDate(Me.txtFecha_vencimiento.Text) = False Then
   Me.txtFecha_vencimiento.Text = CVDate(KEY_FECHA)
End If

TotalFactura = Val(Me.lblTotal.Text) + Val(Me.TxtPecepcion.Text)

If Len(Trim(Me.txtRuc.Text)) = 8 Then
    cod_identidad = 1
End If

If Len(Trim(Me.txtRuc.Text)) = 11 Then
    cod_identidad = 6
End If


If Len(Trim(Me.txtRuc.Text)) <> 8 And Len(Trim(Me.txtRuc.Text)) <> 11 Then
    cod_identidad = 0
End If
ValorCompra = Val(Me.LblValorVenta.Text)
igv = Val(Me.LblIgv.Text)
Descuento = Val(Me.lblDescuento.Text)
If Me.ChkPercepcion.Value = 1 Then
    percepcion = Me.TxtPecepcion.Text
Else
    percepcion = 0
End If


in_saldo = Val(Me.lblTotal.Text)

If Me.DtcTipoDoc.BoundText = "0002" Then
   If Me.chk_suspencion_retencion.Value = 1 Then
      in_saldo = Val(Me.lblTotal.Text)
   Else
      in_saldo = Val(Me.lblTotal.Text) - Val(Me.TxtTotalRetencion.Text)
   End If
End If


If Me.DtcTipoDoc.BoundText = "0091" Then
    GoTo esquivar
End If

If (Trim(Me.txtRuc.Text) = "" Or Trim(Me.TxtProveedor.Text) = "") Then
    MsgBox "INGRESE UN PROVEEDOR DE LA LISTA DE PROVEEDORES", vbInformation, "Documento No Guardado"
    Call Resalta(txtRuc)
    Exit Sub
End If

If Len(Me.TxtNumeroDoc.Text) <= 0 Then
    MsgBox "INGRESE UN COMPROBANTE VALIDO", vbInformation, "Mensaje para el Usuario"
    Call Resalta(Me.txtSerie)
    Exit Sub
End If
esquivar:

'01----------------guardar en Documento Compra---------------------

If rever = True Then
strCadena = "SELECT * FROM movimiento_compra WHERE numero='" & Trim(Me.TxtNumeroGuardado.Text) & "' AND serie='" & Trim(Me.TxtSeriaGuardada.Text) & "' " & _
" AND id_doc='" & Trim(Me.Txtdoc_cod.Text) & "' AND  id_alm='" & Trim(Me.TxtAlmacenGuardado.Text) & "' AND id_proveedor='" & Trim(Me.TxtProveedorGuardado.Text) & "' AND ruc='" & KEY_RUC & "'"
Else
strCadena = "SELECT * FROM movimiento_compra WHERE numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND serie='" & Trim(Me.txtSerie.Text) & "' " & _
" AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND  id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND id_proveedor='" & Trim(Me.txtRuc.Text) & "' AND ruc='" & KEY_RUC & "'"
End If

Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   idCompra = rst("id_compra")
    If MsgBox("ESTA SEGURO DE MODIFICAR ESTE COMPROBANTE DE COMPRA", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
    
    
    
    
    If (Trim(Me.Txtdoc_cod.Text) <> Trim(Me.DtcTipoDoc.BoundText) Or Trim(Me.TxtSeriaGuardada.Text) <> Trim(Me.txtSerie.Text) Or Trim(Me.TxtNumeroGuardado.Text) <> Trim(Me.TxtNumeroDoc.Text) Or Trim(Me.TxtProveedorGuardado.Text) <> Trim(Me.txtRuc.Text)) Then
        strCadena = "SELECT * FROM movimiento_compra WHERE numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND serie='" & Trim(Me.txtSerie.Text) & "' " & _
        " AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND  id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND id_proveedor='" & Trim(txtRuc.Text) & "' and ruc='" & KEY_RUC & "'"
        rstVerifica.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
        If rstVerifica.RecordCount > 0 Then
            MsgBox "Imposible Grabar Este Documento ya esta Registrado para Este Proveedor", vbInformation, KEY_EMPRESA
            Set rstVerifica = Nothing
            Exit Sub
        End If
        Set rstVerifica = Nothing
    End If
     
     strCadena = "UPDATE movimiento_compra SET fecha_kardex='" & Format(Me.DtpKardex.Value, "YYYY-mm-dd") & "',prorrateo_importacion='no', id_periodo='" & Me.DtcPeriodo.BoundText & "', fecha_emision='" & Format(CVDate(Me.TxtFecha_emision.Text), "YYYY-mm-dd") & "',fecha_cancelacion='" & Format(CVDate(Me.txtFecha_vencimiento.Text), "YYYY-mm-dd") & "',id_tipo_compra='" & Me.DtTipoCompra.BoundText & "', " & _
     "id_moneda='" & Me.DtcMoneda.BoundText & "',anio_fabricacion='" & Trim(Me.txtA�oFabricacion.Text) & "',serie_dua='" & Trim(Me.txtseriedua.Text) & "',numero_dua='" & Trim(Me.txtnumero_dua.Text) & "',id_mes='" & formato_item(Month(KEY_FECHA), 2) & "',id_anio='" & Year(KEY_FECHA) & "',id_alm='" & Me.DtcAlmacen.BoundText & "',id_doc='" & Me.DtcTipoDoc.BoundText & "'," & _
     "serie='" & Me.txtSerie.Text & "',numero='" & Me.TxtNumeroDoc.Text & "',tipo_doc_identidad='" & cod_identidad & "',id_proveedor='" & Me.txtRuc.Text & "',anio_modelo='" & Trim(Me.txta�omodelo.Text) & "',anio_dua='" & Trim(Me.TxtAnioDua.Text) & "',nproveedor='" & Trim(Me.TxtProveedor.Text) & "'," & _
     "tc='" & Val(Me.txtTc.Text) & "',retencion='" & Val(Me.TxtTotalRetencion.Text) & "',valor_venta='" & Val(Me.LblValorVenta.Text) & "',igv='" & Val(Me.LblIgv.Text) & "',isc='" & Val(Me.lblISC.Text) & "',percepcion='" & Val(Me.TxtPecepcion.Text) & "',exonerado='" & Val(Me.lblExonerado.Text) & "',total='" & Val(Me.lblTotal.Text) & "',saldo='" & Val(Me.lblTotal.Text) & "',fob='" & Val(Me.txtFob.Text) & "',seguro='" & Val(Me.txtSeguro.Text) & "',flete='" & Val(Me.TxtFlete.Text) & "',cif='" & Val(Me.TxtCif.Text) & "',anulado='no' WHERE id_compra='" & idCompra & "' AND ruc='" & KEY_RUC & "'"
     CnBd.Execute (strCadena)
     
    strCadena = "UPDATE movimiento_compra_temporal SET numero='" & Trim(Me.TxtNumeroDoc.Text) & "',id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "',serie='" & Trim(Me.txtSerie.Text) & "'" & _
     " WHERE numero='" & Trim(Me.TxtNumeroGuardado.Text) & "' AND serie='" & Trim(Me.TxtSeriaGuardada.Text) & "' AND id_doc='" & Trim(Me.Txtdoc_cod.Text) & "' AND dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
     CnBd.Execute (strCadena)
     
     
     strCadena = "DELETE FROM movimiento_compra_detalle WHERE id_compra='" & idCompra & "' AND ruc='" & KEY_RUC & "'"
     CnBd.Execute (strCadena)
     
     Call SaveDetalleDocumentoCompra(idCompra)
     
     strCadena = "SELECT * FROM imp_producto_detalle_temp WHERE id_compra='" & idCompra & "'"
     Call ConfiguraRstZ(strCadena)
     If rstZ.RecordCount > 0 Then
        rstZ.MoveFirst
        For m = 0 To rstZ.RecordCount - 1
            strCadena = "UPDATE imp_producto_detalle SET  serie='" & rstZ("serie") & "',id_estado='" & rstZ("id_estado") & "',id_estado_detalle='" & rstZ("id_estado_detalle") & "',id_alm='" & rstZ("id_alm") & "',anio_fabricacion='" & Trim(Me.txtA�oFabricacion.Text) & "',anio_contenedor='" & Trim(Me.TxtAnioDua.Text) & "',nro_contenedor='" & Trim(Me.txtnumero_dua.Text) & "',nro_chasis='" & rstZ("nro_chasis") & "',nro_motor='" & rstZ("nro_motor") & "',anio_modelo='" & Trim(Me.TxtAnioDua.Text) & "',item='" & rstZ("item") & "',serie_asignada='" & rstZ("serie_asignada") & "',vendido='" & rstZ("vendido") & "' WHERE id_compra='" & idCompra & "' and id_orden='" & rstZ("id_orden") & "' and id_producto='" & rstZ("id_producto") & "'  "
            CnBd.Execute (strCadena)
             
            rstZ.MoveNext
        Next m
     End If
     
     strCadena = "DELETE FROM imp_producto_detalle_temp where id_compra='" & idCompra & "'"
     CnBd.Execute (strCadena)
     
     
     If KEY_CONTABILIDAD = "si" And Me.DtcTipoDoc.BoundText <> "0089" And Me.DtcTipoDoc.BoundText <> "0009" And Me.DtcTipoDoc.BoundText <> "0031" Then
        
        
        strCadena = "DELETE FROM mis_cuentas_det_detalle WHERE id_movimiento='" & Val(Me.txtIdCompra.Text) & "' and monto_pagado='" & Val(Me.TxtTotalRetencion.Text) & "' LIMIT 1"
        CnBd.Execute (strCadena)
        
        strCadena = "Call CON_Asiento_EliminarCompra('" & idCompra & "', '" & KEY_USUARIO & "', '" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
     
        
        
        
    If Me.DtcTipoDoc.BoundText = "0089" Then
        If KEY_PAIS = KEY_PERU Then
            strCadena = "call p_insert_compra_emitido_premiun('" & idCompra & "')"
        Else
            strCadena = "call p_insert_compra_emitido_premiun_internacional('" & idCompra & "')"
        End If
        CnBd.Execute (strCadena)
        
        Call put_ajuste_ingreso(idCompra)
    Else
        
        
        If Me.DtcTipoDoc.BoundText <> "0089" And Me.DtcTipoDoc.BoundText <> "0090" Then
        
        If KEY_PAIS = KEY_PERU Then
            strCadena = "call p_insert_compra_emitido_premiun('" & idCompra & "')"
        Else
            strCadena = "call p_insert_compra_emitido_internacional('" & idCompra & "')"
        End If
            CnBd.Execute (strCadena)
        End If
        
    End If
    
    
    
    
        
        
        
        
        
        
        Me.lblidCompra.Caption = get_periodo_detalle(Me.DtcPeriodo.BoundText, id_compra)
        MsgBox "NUMERO DE VOUCHER GENERADO  : " & Trim(idCompra), vbInformation, KEY_VENDEDOR
        
        
        
        
End If


If KEY_CONTABILIDAD = "si" And Me.DtcTipoDoc.BoundText = "0090" Then
    If KEY_PAIS <> KEY_PERU Then ' PERU
            strCadena = "call CON_InsertaAsiento_salida_internacional('" & idCompra & "')"
            CnBd.Execute (strCadena)
    End If
End If

If Format(Me.DtpKardex.Value, "YYYY-mm-dd") < KEY_FECHA Then
    
   If (Me.DtTipoCompra.BoundText = "02" Or Me.DtTipoCompra.BoundText = "03") Then
       If Me.DtcTipo.BoundText = "01" Then ' material
        Call actualizar_kardex_item(idCompra)
        End If
    End If
End If

     
     Me.TxtCodProducto.Enabled = False
     Me.TxtDescripcionProducto.Enabled = False
     Me.txtCantidad.Enabled = False
     Me.txtcosto.Enabled = False
     Me.cmdAgregar.Enabled = False
     Me.CmdQuitar.Enabled = False
     
     Me.TxtAnioDua.Locked = True
     Me.txtA�oFabricacion.Locked = True
     Me.txta�omodelo.Locked = True
     Me.txtnumero_dua.Locked = True
     
     Me.cmdsave.Enabled = False
     Me.cmdRevertir.Enabled = True
     Me.cmdModificar.Enabled = True
     
     
     
     rever = False
     Exit Sub
     Else
        Exit Sub
   End If
    rever = False
 
End If
        
         
          If Me.DtcTipo.BoundText = "02" Then
           in_cta_compra = KEY_CTA_PAGAR_SERVICIO
        End If
         
         If Me.DtcMoneda.BoundText = "00001" Then
                in_cta_compra = KEY_CTA_COMPRA_SOLES
         Else
                in_cta_compra = KEY_CTA_COMPRA_SOLES
         End If
           
        If Me.DtcTipoDoc.BoundText = "0002" Then
             in_cta_compra = KEY_CTA_COMPRA_RH
        End If
        
        
        If Me.DtcTipoDoc.BoundText = "0417" And Me.DtcMoneda.BoundText = "00001" Then
            in_cta_compra = KEY_CTA_LETRA_PAGAR_SOLES
        End If
        
        If Me.DtcTipoDoc.BoundText = "0417" And Me.DtcMoneda.BoundText = "00002" Then
            in_cta_compra = KEY_CTA_LETRA_PAGAR_DOLARES
        End If
        
        If Me.DtcTipoDoc.BoundText = "0418" And Me.DtcMoneda.BoundText = "00001" Then
            in_cta_compra = KEY_CTA_FET_SOLES
        End If
        
        If Me.DtcTipoDoc.BoundText = "0418" And Me.DtcMoneda.BoundText = "00002" Then
            in_cta_compra = KEY_CTA_FET_DOLARES
        End If
        
        If Me.DtcTipoDoc.BoundText = "0419" And Me.DtcMoneda.BoundText = "00001" Then
           in_cta_compra = KEY_CTA_ANT_SOLES
        End If
        If Me.DtcTipoDoc.BoundText = "0419" And Me.DtcMoneda.BoundText = "00002" Then
           in_cta_compra = KEY_CTA_ANT_DOLARES
        End If
        
        
       
        
        
        
        If KEY_CONTABILIDAD = "si" Then
           
           
        
           If put_verifica_cuenta_contable(Me.DtcTipoDoc.BoundText, Trim(Me.txtSerie.Text), Trim(Me.TxtNumeroDoc.Text), in_cta_compra, Me.DtTipoCompra.BoundText) = False Then
              Exit Sub
           End If
           
        End If
       
       
        
        If Me.chkresponsable.Value = 1 Then
           in_responsable = Me.DtcResponsable.BoundText
        Else
           in_responsable = "0"
        End If
        
        If Me.DtcTipoDoc.BoundText = "0089" Or Me.DtcTipoDoc.BoundText = "0090" Then
            Me.TxtNumeroDoc.Text = get_correlativo(Me.DtcTipoDoc.BoundText, Trim(Me.txtSerie.Text))
        End If
        
        
        If KEY_PAIS = KEY_PERU Then
            strCadena = "call P_insert_compra_ultimate('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Format(CVDate(Me.TxtFecha_emision.Text), "YYYY-mm-dd") & "','" & Format(CVDate(Me.txtFecha_vencimiento.Text), "YYYY-mm-dd") & "','02'," & _
            "'" & Me.DtTipoCompra.BoundText & "','--','" & Me.DtcMoneda.BoundText & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Trim(Me.txtSerie.Text) & "'," & _
            "'" & Format(Trim(Me.TxtNumeroDoc.Text), "00000000") & "','" & cod_identidad & "','" & Trim(Me.txtRuc.Text) & "','" & UCase(Me.TxtProveedor.Text) & "','" & Trim(Me.txtTc.Text) & "'," & _
            "'0','" & Val(Me.LblValorVenta.Text) & "','" & Val(Me.LblIgv.Text) & "','" & Val(Me.lblISC.Text) & "','0','" & Val(Me.TxtPecepcion.Text) & "','0','" & Val(Me.lblExonerado.Text) & "','0','" & Val(Me.lblTotal.Text) & "','" & in_saldo & "'," & _
            " '" & KEY_USUARIO & "','" & Trim(Me.txtObservacion.Text) & "','" & Me.DtcTipo.BoundText & "','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & in_responsable & "','" & Val(Me.txtFob.Text) & "','" & Val(Me.txtSeguro.Text) & "','" & Val(Me.TxtFlete.Text) & "','" & Val(Me.TxtCif.Text) & "','" & KEY_RUC & "')"
        Else
            strCadena = "call P_insert_compra_ultimate_internacional_ii('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Format(CVDate(Me.TxtFecha_emision.Text), "YYYY-mm-dd") & "','" & Format(CVDate(Me.txtFecha_vencimiento.Text), "YYYY-mm-dd") & "','02'," & _
            "'" & Me.DtTipoCompra.BoundText & "','--','" & Me.DtcMoneda.BoundText & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Trim(Me.txtSerie.Text) & "'," & _
            "'" & Format(Trim(Me.TxtNumeroDoc.Text), "00000000") & "','" & cod_identidad & "','" & Trim(Me.txtRuc.Text) & "','" & UCase(Me.TxtProveedor.Text) & "','" & Trim(Me.txtTc.Text) & "'," & _
            "'0','" & Val(Me.LblValorVenta.Text) & "','" & Val(Me.LblIgv.Text) & "','" & Val(Me.lblISC.Text) & "','0','" & Val(Me.TxtPecepcion.Text) & "','0','" & Val(Me.lblExonerado.Text) & "','0','" & Val(Me.lblTotal.Text) & "','" & in_saldo & "'," & _
            " '" & KEY_USUARIO & "','" & Trim(Me.txtObservacion.Text) & "','" & Me.DtcTipo.BoundText & "','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & in_responsable & "','" & Val(Me.txtFob.Text) & "','" & Val(Me.txtSeguro.Text) & "','" & Val(Me.TxtFlete.Text) & "','" & Val(Me.TxtCif.Text) & "','" & Trim(Me.TxtAlmacen.Text) & "','" & Trim(Me.txtNumeroAutorizacion.Text) & "','" & KEY_RUC & "')"
        End If
        
        Call ConfiguraRstP(strCadena)
        id_compra = rstP(0)
        Me.txtIdCompra.Text = id_compra
        
        
        If Val(lblISC) > 0 Then
            strCadena = "UPDATE movimiento_compra SET isc='" & Val(Me.lblISC.Text) & "' WHERE id_compra='" & Val(Me.txtIdCompra.Text) & "'"
            CnBd.Execute (strCadena)
        End If
        
        
        If Me.DtcTipoDoc.BoundText = "0020" Then
            strCadena = "UPDATE movimiento_compra SET id_comprobante='" & get_comprobante_reten(DtcRelacionado.BoundText, Me.TxtSerieR.Text, Me.TxtNumeroR.Text, Me.DtcProveedor.BoundText) & "'  WHERE id_compra='" & id_compra & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
        End If
        
        
        
       
        
        
        
        If Me.DtcTipoDoc.BoundText = "0007" Then
            strCadena = "UPDATE movimiento_compra SET id_comprobante='" & Val(Me.txtid_comprobante_relacionado.Text) & "' , id_tipo_nota='" & Me.DtcTipoNota.BoundText & "',valor_venta='" & Val(Me.LblValorVenta.Text) & "',igv='" & Val(Me.LblIgv.Text) & "',isc='" & Val(Me.lblISC.Text) & "',percepcion='" & Val(Me.TxtPecepcion.Text) & "',exonerado='" & Val(Me.lblExonerado.Text) & "',total='" & Val(Me.lblTotal.Text) * -1 & "',fecha_fact='" & Format(Me.TxtFechaEmision.Text, "YYYY-mm-dd") & "',id_doc_fact='" & Me.DtcRelacionado.BoundText & "',serie_fact='" & Trim(Me.TxtSerieR.Text) & "',numero_fact='" & Trim(Me.TxtNumeroR.Text) & "' WHERE id_compra='" & id_compra & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            If KEY_RUC <> "20128836251" Then
            Call generar_recibo_egreso(Me.TxtFecha_emision.Text, Trim(Me.txtRuc.Text), Val(Me.lblTotal.Text), Val(Me.txtTc.Text), Me.DtcMoneda.BoundText, "-", 0, Val(Me.txtid_comprobante_relacionado.Text), Val(Me.txtIdCompra.Text))
            End If
                                    
                                    
            
        End If
        
        If Me.DtcTipoDoc.BoundText = "0008" Then
            strCadena = "UPDATE movimiento_compra SET id_comprobante='" & Val(Me.txtid_comprobante_relacionado.Text) & "' , id_tipo_nota='" & Me.DtcTipoNota.BoundText & "',valor_venta='" & Val(Me.LblValorVenta.Text) & "',igv='" & Val(Me.LblIgv.Text) & "',isc='" & Val(Me.lblISC.Text) & "',percepcion='" & Val(Me.TxtPecepcion.Text) & "',exonerado='" & Val(Me.lblExonerado.Text) & "',total='" & Val(Me.lblTotal.Text) & "',fecha_fact='" & Format(Me.TxtFechaEmision.Text, "YYYY-mm-dd") & "',id_doc_fact='" & Me.DtcRelacionado.BoundText & "',serie_fact='" & Trim(Me.TxtSerieR.Text) & "',numero_fact='" & Trim(Me.TxtNumeroR.Text) & "' WHERE id_compra='" & id_compra & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
        End If
        
        
        Call SaveDetalleDocumentoCompra(id_compra)
        strCadena = "call p_update_proveedor('" & Trim(Me.txtRuc.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        If Me.Frame_Relacionado.Visible = True Then
            strCadena = "UPDATE movimiento_compra SET fecha_fact='" & Format(CVDate(Me.TxtFechaEmision.Text), "YYYY-mm-dd") & "',numero_fact='" & Trim(Me.TxtNumeroR.Text) & "',serie_fact='" & Trim(Me.TxtSerieR.Text) & "',id_doc_fact='" & DtcRelacionado.BoundText & "',id_proveedor_relacionado='" & DtcProveedor.BoundText & "' WHERE id_compra='" & Val(Me.txtIdCompra.Text) & "'"
            CnBd.Execute (strCadena)
        End If
        
        in_afecto_costo = "no"
        If Me.chk_afecto_costo.Value = 1 Then
            in_afecto_costo = "si"
        Else
            in_afecto_costo = "no"
        End If
        
        If Me.frmtiponota.Visible = True Then
            in_tipo_nota = Me.DtcTipoNota.BoundText
        Else
            in_tipo_nota = 0
        End If
               
        strCadena = "UPDATE movimiento_compra SET fecha_kardex='" & Format(Me.DtpKardex.Value, "YYYY-mm-dd") & "' ,   afecta_costo='" & in_afecto_costo & "',retencion='" & Val(Me.TxtTotalRetencion.Text) & "', fecha_registro=CURDATE() ,anio_fabricacion='" & Trim(Me.txtA�oFabricacion.Text) & "',serie_dua='" & Trim(Me.txtseriedua.Text) & "',`numero_dua`='" & Trim(Me.txtnumero_dua.Text) & "',`anio_modelo`='" & Me.TxtAnioDua.Text & "',anio_dua='" & Trim(Me.TxtAnioDua.Text) & "' WHERE id_compra='" & Val(Me.txtIdCompra.Text) & "'  and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
        
        If Me.DtcTipoDoc.BoundText = "0002" And Val(Me.TxtTotalRetencion.Text) > 0 And Me.chk_suspencion_retencion.Value = 0 Then ' recibo Honorario
            Call put_retencion_automatica(Me.DtcAlmacen.BoundText)
        End If
        
        
        'NOTA DE CREDITO
        
        If Me.DtTipoCompra.BoundText = "01" Then
            strCadena = "INSERT INTO movimiento_compra_importacion(id_compra,ruc)VALUES('" & id_compra & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
        End If
        
        
         If KEY_PAIS <> KEY_PERU And Me.DtcTipoDoc.BoundText = "0427" Then
            Call generar_recibo_egreso(Me.TxtFecha_emision.Text, Trim(Me.txtRuc.Text), Val(Me.lblTotal.Text), Val(Me.txtTc.Text), Me.DtcMoneda.BoundText, "-", 0, Val(Me.txtid_comprobante_relacionado.Text), Val(Me.txtIdCompra.Text))
        End If
        
        
'02----------------guardar en detalle documento Compra-----------
 
If KEY_CONTABILIDAD = "si" Then
    If Me.DtcTipoDoc.BoundText = "0009" Or Me.DtcTipoDoc.BoundText = "0031" Then
        strCadena = "call CON_InsertaAsiento_GuiaRR('" & id_compra & "')"
        CnBd.Execute (strCadena)
        Me.lblidCompra.Caption = get_periodo_detalle(Me.DtcPeriodo.BoundText, id_compra)
       
    End If
End If

 
 
If KEY_CONTABILIDAD = "si" And Me.DtcTipoDoc.BoundText <> "0009" And Me.DtcTipoDoc.BoundText <> "0031" Then
    
    
    If Me.DtcTipoDoc.BoundText = "0089" Then
        If KEY_PAIS = KEY_PERU Then
            strCadena = "call p_insert_compra_emitido_premiun('" & id_compra & "')"
        Else
            strCadena = "call p_insert_compra_emitido_premiun_internacional('" & id_compra & "')"
        End If
        CnBd.Execute (strCadena)
        
        Call put_ajuste_ingreso(id_compra)
    Else
        
        
        If Me.DtcTipoDoc.BoundText <> "0089" And Me.DtcTipoDoc.BoundText <> "0090" Then
        
        If KEY_PAIS = KEY_PERU Then
            strCadena = "call p_insert_compra_emitido_premiun('" & id_compra & "')"
        Else
            strCadena = "call p_insert_compra_emitido_internacional('" & id_compra & "')"
        End If
            CnBd.Execute (strCadena)
        End If
        
    End If
    
    
    If KEY_CONTABILIDAD = "si" And Me.DtcTipoDoc.BoundText = "0090" Then
    If KEY_PAIS <> KEY_PERU Then ' PERU
            strCadena = "call CON_InsertaAsiento_salida_internacional('" & id_compra & "')"
            Call ConfiguraRstK(strCadena)
    End If
End If
    If Me.DtcTipoDoc.BoundText = "0419" Then
        strCadena = "UPDATE con_documento SET CuentaContable='422' WHERE IdReferencia='" & id_compra & "'"
        CnBd.Execute (strCadena)
    End If
    
    
    
    
    
    
    
    
   
End If



If Format(Me.DtpKardex.Value, "YYYY-mm-dd") < KEY_FECHA Then
    
   If (Me.DtTipoCompra.BoundText = "02" Or Me.DtTipoCompra.BoundText = "03") Then
       If Me.DtcTipo.BoundText = "01" Then ' material
          Call actualizar_kardex_item(id_compra)
       End If
    End If
End If





 Me.lblidCompra.Caption = get_periodo_detalle(Me.DtcPeriodo.BoundText, id_compra)
MsgBox "NUMERO DE VOUCHER GENERADO  : " & Trim(Me.lblidCompra.Caption), vbInformation, KEY_VENDEDOR

If Trim(Me.DtcTipoDoc.BoundText) = "0089" Then
    num = formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6)
    strCadena = "UPDATE almacen_comprobante SET numero='" & num & "' WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.txtSerie.Text) & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
End If

If Trim(Me.DtcTipoDoc.BoundText) = "0090" Then
    num = formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6)
    strCadena = "UPDATE almacen_comprobante SET numero='" & num & "' WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.txtSerie.Text) & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
    CnBd.Execute (strCadena)
End If
    
 
 If KEY_PROYECTO = "si" Then
    If Val(Me.DtcProyecto.BoundText) > 0 Then
        strCadena = "call sp_update_movimiento_compra_proyecto('" & Val(id_compra) & "','" & Val(Me.DtcProyecto.BoundText) & "')"
        CnBd.Execute (strCadena)
    End If
 End If
 
 
'02-------------------------------------------------------------
                Me.TxtCodProducto.Enabled = False
                Me.TxtDescripcionProducto.Enabled = False
                Me.txtCantidad.Enabled = False
                Me.txtcosto.Enabled = False
                 Me.cmdAgregar.Enabled = False
                Me.CmdQuitar.Enabled = False
                'Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
                Me.cmdsave.Enabled = False
                
                Call Me.llenarGrid_Comprobante(Me.HfdDetalle, id_compra)
                
                If Me.DtTipoCompra.BoundText = "01" Then
                    Me.cmdProrrateoImportacion.Enabled = True
                    chkValorVenta_importacion.Enabled = True
                    Me.chk_cantidad_importacion.Enabled = True
                End If
                
                
                
                Exit Sub
'salir:
    MsgBox "No se Guardo este Documento, Disculpe las Molestias", vbInformation, KEY_EMPRESA


End Sub

Public Sub actualizar_kardex_item(ByVal in_compra As String)
MsgBox "SE VA A PROCEDER A ACTUALIZAR KARDEX" + Chr(13) + Chr(13) + "PULSE ACEPTAR Y DEJE QUE TERMINE DE ACTUALIZAR.", vbInformation
    
    
    
   strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & Val(in_compra) & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRstIN(strCadena)
      If rstIN.RecordCount > 0 Then
         Me.progresbar_kardex.Min = 0
         Me.progresbar_kardex.Max = rstIN.RecordCount
         rstIN.MoveFirst
         
         
         For i = 0 To rstIN.RecordCount - 1
            If KEY_RUC = "20128836251" Then
               Call update_kardex_Vargas_modulo_compra(rstIN("id_producto"), Format(Me.DtpKardex.Value, "YYYY-mm-dd"))
            Else
               
               
                    If KEY_PAIS = KEY_PERU Then
                        Call update_kardex_update(rstIN("id_producto"), Format(Me.DtpKardex.Value, "YYYY-mm-dd"))
                    Else
                        Call update_kardex_internacional(rstIN("id_producto"), Format(Me.DtpKardex.Value, "YYYY-mm-dd"))
                    End If
            End If
            
            
            rstIN.MoveNext
            Me.progresbar_kardex.Value = i
            DoEvents
         Next i
      
      
      
      End If
      
      MsgBox "Proceso Actualizacion Kardex Correcto.", vbInformation
      
      
      
End Sub

Private Function get_correlativo(ByVal in_doc As String, ByVal in_serie As String) As String

strCadena = "SELECT numero FROM movimiento_compra WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and  id_doc='" & in_doc & "' and serie='" & in_serie & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    get_correlativo = Format(Val(rstL("numero")) + 1, "00000000")
Else
    get_correlativo = Format(1, "00000000")
End If



End Function


Private Sub save_seriales(ByVal in_compra As String, ByVal in_detalle_compra As String, ByVal in_producto As String, ByVal in_cantidad As Double)
strCadena = "SELECT * FROM producto p, linea l WHERE l.produccion='si' and  p.id_producto='" & in_producto & "' and  p.id_linea=l.id_linea and p.ruc=l.id_usu and p.ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    For i = 0 To in_cantidad - 1
        strCadena = "INSERT INTO imp_producto_detalle (`id_compra`,`id_detalle_compra`,`id_producto`,`id_alm`, " & _
        "`anio_fabricacion`,`anio_contenedor`,`nro_contenedor`,`anio_modelo`,`ruc`)VALUES " & _
        "('" & in_compra & "','" & in_detalle_compra & "','" & in_producto & "','" & Me.DtcAlmacen.BoundText & "','" & Trim(Me.txtA�oFabricacion.Text) & "'," & _
        "'" & Trim(Me.TxtAnioDua.Text) & "','" & Trim(Me.txtnumero_dua.Text) & "','" & Trim(Me.txta�omodelo.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
    Next i
End If
End Sub
Private Sub SaveDetalleDocumentoCompra(ByVal id_compra As Double)
'******************************* INICIO PROCESO*************************
Dim in_cantidad As Double

Dim in_costo_unitario As Double
Dim in_monto_fs As Single

If Me.DtTipoCompra.BoundText = "01" Then
    in_monto_fs = Val(Me.txtSeguro.Text) + Val(Me.TxtFlete.Text)
Else
    in_monto_fs = Val(Me.lblgastos.Text)
End If

If Me.DtcTipoDoc.BoundText = "0007" Then
   strCadena = "DELETE FROM kardex WHERE id_movimiento='" & Val(Me.txtIdCompra.Text) & "' and id_doc='" & Me.DtcTipoDoc.BoundText & "' and id_serie='" & Trim(Me.txtSerie.Text) & "' and id_numero='" & Trim(Me.TxtNumeroDoc.Text) & "' and ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
End If



 strCadena = "SELECT * FROM movimiento_compra_temporal WHERE  id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.txtSerie.Text) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "' ORDER BY id_temporal ASC"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       For i = 0 To rstT.RecordCount - 1
           If Me.DtTipoCompra.BoundText = "01" Then
                in_monto_parcial = rstT("valor_venta") / (Val(Me.LblValorVenta.Text) - in_monto_fs) * in_monto_fs
           Else
                in_monto_parcial = 0
           End If
           
           
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,fecha_vencimiento,numero_lote,obsequio,id_unidad,cuenta_contable,porcentaje_retencion,ruc) VALUES ('" & id_compra & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("c_unitario") & "'," & _
           "'" & rstT("dsto_soles") & "','" & rstT("dsto_procentaje") & "','" & rstT("total_descuento") & "','" & rstT("valor_neto") & "','" & rstT("isc") & "','" & rstT("igv") & "', " & _
           "'" & rstT("retencion") & "','" & rstT("otros") & "','" & rstT("percepcion") & "','" & rstT("valor_venta") & "','" & rstT("exonerado") & "','" & rstT("precio_venta") & "','" & rstT("p_venta") & "','" & rstT("p_costo") & "','" & rstT("id_alm") & "','" & rstT("detalle") & "','" & in_monto_parcial & "','" & Format(rstT("fecha_vencimiento"), "YYYY-mm-dd") & "','" & rstT("numero_lote") & "','" & rstT("obsequio") & "','" & rstT("id_unidad") & "','" & rstT("cuenta_contable") & "','" & rstT("porcentaje_retencion") & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
          If get_servicio(rstT("id_producto")) = "no" Then
            If Me.DtcTipoDoc.BoundText <> "0419" Then
                
                
                    If Me.DtcMoneda.BoundText = "00002" And Me.DtcMoneda.BoundText <> KEY_MONEDA Then
                       in_costo_unitario = (rstT("c_unitario")) * Val(Me.txtTc.Text)
                    Else
                       If KEY_CON_IGV = "si" Then
                          in_costo_unitario = rstT("valor_venta") / rstT("cantidad")
                       Else
                          in_costo_unitario = rstT("precio_venta") / rstT("cantidad")
                       End If
                       
                    End If
                    
                               
                 in_cantidad = get_cantidad_agranel(rstT("id_producto"), rstT("id_unidad")) * rstT("cantidad")
                 If rstT("cantidad") <> in_cantidad Then
                    in_costo_unitario = in_costo_unitario / in_cantidad
                 End If
                 
                
                strCadena = "UPDATE almacen_producto SET precio_venta='" & Val(rstT("p_venta")) & "' WHERE id_producto='" & rstT("id_producto") & "' and id_alm='" & rstT("id_alm") & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                 
                
                If KEY_PAIS = KEY_PERU Then
                    If KEY_RUC = "10419736444" Then  ' 5 DECIMALES DE CANTIDAD
                        strCadena = "call put_kardex_stock_vitekey_v5('02','" & Format(Me.DtpKardex.Value, "YYYY-mm-dd") & "','" & Val(id_compra) & "','" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.txtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.txtRuc.Text) & "','" & rstT("id_producto") & "','" & in_cantidad & "','" & in_costo_unitario & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & rstT("obsequio") & "','" & KEY_RUC & "')"
                    Else
                        strCadena = "call put_kardex_stock_vitekey_v1('02','" & Format(Me.DtpKardex.Value, "YYYY-mm-dd") & "','" & Val(id_compra) & "','" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.txtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.txtRuc.Text) & "','" & rstT("id_producto") & "','" & in_cantidad & "','" & in_costo_unitario & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & rstT("obsequio") & "','" & KEY_RUC & "')"
                    End If
                    
                Else
                    If Me.DtcTipoDoc.BoundText = "0090" Then
                        nntipo = "01"
                    Else
                        nntipo = "02"
                    End If
                    strCadena = "call put_kardex_stock_vitekey_internacional('" & nntipo & "','" & Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd") & "','" & Val(id_compra) & "','" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.txtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.txtRuc.Text) & "','" & rstT("id_producto") & "','" & in_cantidad & "','" & in_costo_unitario & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                End If
                CnBd.Execute (strCadena)
                
                
           End If
          End If
           
           
           strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & id_compra & "' and  ruc='" & KEY_RUC & "' ORDER BY id_detalle_compra DESC LIMIT 1"
           Call ConfiguraRst(strCadena)
           If rst.RecordCount > 0 Then
            Call save_seriales(id_compra, rst("id_detalle_compra"), rstT("id_producto"), rstT("cantidad"))
          End If
          
          If KEY_RUC <> "20487473881" Then
          
          If Me.DtcTipoDoc.BoundText = "0089" Or Me.DtcTipoDoc.BoundText = "0090" Then
               
               
                Call put_saldo_stock(rstT("id_producto"))
               
           End If
          End If
          
           rstT.MoveNext
        Next i
    End If
    
    strCadena = "DELETE FROM movimiento_compra_temporal WHERE (numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.txtSerie.Text) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "')"
    Call ConfiguraRstT(strCadena)
    
'******************************* FINISH PROCESO*************************
End Sub
Private Function CodigoDetalleCompra() As String
    strCadena = "SELECT int_det_documentoCompra FROM Detalle_DocumentoCompra ORDER BY int_det_documentoCompra DESC"
    Call ConfiguraRst(strCadena)
    CodigoDetalleCompra = GeneraCodigos()
    Set rst = Nothing
End Function
Private Sub save_preciocompra(ByVal codigo As String, ByVal precio As Double)
Dim consecutivo As String
strCadena = "SELECT * FROM Producto_precio WHERE cProducto='" & Trim(codigo) & "'"
Call ConfiguraRst(strCadena)
consecutivo = GeneraCodigo(10)



End Sub

Private Sub txtRedondeo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim redondeo As Single
Dim Total As Single
redondeo = Val(Me.txtRedondeo.Text)
Total = Val(Me.LblIgv.Text) + Val(Me.LblValorVenta.Text) + redondeo
Me.lblTotal.Text = Format(Total, "###0.00")
End If
End Sub

Private Sub put_renta_iva(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_proveedor As String)

strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & in_doc & "' and serie='" & Trim(in_serie) & "' and numero='" & Trim(in_numero) & "' and id_proveedor='" & in_proveedor & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.TxtDescripcionProducto.Text = "EJERCICIO:" & Year(Me.TxtFecha_emision.Text) & Space(2) & "MONTO:[" & str(rst("igv")) & "]" & Space(2) & "IVA :" & Trim(Me.DtcIva.Text) & "%" & Space(2) & str(rst("igv") * Val(Me.DtcIva.Text) / 100)
    Me.TxtUnitario.Text = rst("igv") * Val(Me.DtcIva.Text) / 100
    Call Resalta(Me.TxtUnitario)
End If


End Sub


Private Sub put_renta(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_proveedor As String)

strCadena = "SELECT if(valor_venta=0,exonerado,valor_venta) as valor_venta FROM movimiento_compra WHERE id_doc='" & in_doc & "' and serie='" & Trim(in_serie) & "' and numero='" & Trim(in_numero) & "' and id_proveedor='" & in_proveedor & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.TxtDescripcionProducto.Text = "EJERCICIO:" & Year(Me.TxtFecha_emision.Text) & Space(2) & "MONTO:[" & str(rst("valor_venta")) & " ]" & Space(2) & "RETENCION:" & Trim(Me.DtcRetencionFuente.Text) & "%" & Space(2) & str(rst("valor_venta") * Val(Me.DtcRetencionFuente.Text) / 100)
    Me.TxtUnitario.Text = rst("valor_venta") * Val(Me.DtcRetencionFuente.Text) / 100
    Call Resalta(Me.TxtUnitario)
    
End If


End Sub

Private Function get_comprobante_reten(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_proveedor As String) As String

strCadena = "SELECT id_compra FROM movimiento_compra WHERE id_doc='" & in_doc & "' and serie='" & Trim(in_serie) & "' and numero='" & Trim(in_numero) & "' and id_proveedor='" & in_proveedor & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    get_comprobante_reten = rst("id_compra")
Else
    get_comprobante_reten = 0
End If


End Function

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
On Error GoTo errohandler
 If KeyAscii = 13 Then
   Call buscarcliente

End If
Exit Sub
errohandler: MsgBox "Hubo un Error Digite Nuevamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub txtSeguro_Change()
Me.TxtCif.Text = Val(Me.txtFob.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtSeguro.Text)
If Val(Me.txtIdCompra.Text) < 1 Then
    If Me.DtTipoCompra.BoundText = "01" Then ' IMPORTACION
       Me.lblTotal.Text = Val(Me.txttotal_final.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtSeguro.Text)
       Me.LblValorVenta.Text = Val(Me.lblTotal.Text)
       Me.lblIMPBruto.Text = Val(Me.lblTotal.Text)
       Me.LblIgv.Text = 0
    End If
    
End If
End Sub

Private Sub TxtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Me.DtcTipoDoc.SetFocus
End If
If KeyCode = vbKeyRight Then
        Call Resalta(Me.TxtNumeroDoc)
End If
End Sub
Private Sub txtSerie_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    If (Me.DtcTipoDoc.BoundText = "0089" Or Me.DtcTipoDoc.BoundText = "0090") Then
    
    
    
    If KEY_PAIS <> KEY_PERU Then
        Me.txtSerie.Text = formato_item(Me.txtSerie.Text, 3)
    Else
        Me.txtSerie.Text = formato_item(Me.txtSerie.Text, 4)
    End If
    
    strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "' AND serie='" & Trim(Me.txtSerie.Text) & "' ORDER BY numero DESC LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.txtSerie.Text = rst("serie")
       Me.TxtNumeroDoc.Text = Format(Val(rst("numero")) + 1, "000000000")
       Call Resalta(Me.TxtNumeroDoc)
    Else
        strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "' AND serie='" & Trim(Me.txtSerie.Text) & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            Me.TxtNumeroDoc.Text = formato_item(rst("numero"), 8)
            Call Resalta(Me.txtRuc)
        End If
    End If
    
    Else
        Me.txtSerie.Text = formato_item(Me.txtSerie.Text, 4)
        Call Resalta(Me.TxtNumeroDoc)
   End If
    
        
    End If

End Sub

Private Sub VerificaAnulado(ByVal idCompra As String)
'strCadena = "Select anulado FROM movimiento_compra WHERE  id_compra='" & idCompra & "' AND ruc='" & KEY_RUC & "'"
'Call ConfiguraRstT(strCadena)
'If rst.RecordCount > 0 Then
 '   If Trim(rstT("anulado")) = "si" Then
  '      Me.lblAnulado.Visible = True
        
   '     Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
    '    Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
    'Else
     '   Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
    'End If
'End If
'Set rstT = Nothing
End Sub







Private Sub TxtSerieG_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  '  Me.TxtSerieG.Text = formato_item(Me.TxtSerieG.Text, 4)
  '  Call Resalta(Me.TxtNumeroG)
End If
End Sub

Private Sub TxtSerie_orden_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Call Resalta(Me.txtNumero_orden)
    
End If
End Sub

Private Sub txtseriemotor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtitemdua)
End If

End Sub

Private Sub TxtSerieR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If KEY_PAIS = KEY_PERU Then
        Me.TxtSerieR.Text = UCase(formato_item(Me.TxtSerieR.Text, 4))
    Else
        Me.TxtSerieR.Text = UCase(formato_item(Me.TxtSerieR.Text, 3))
    End If
     
    Me.TxtNumeroR.SetFocus
End If
End Sub

Private Sub TxtTc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtObservacion)
End If
End Sub

Private Sub TxtTotalDescuento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        
        Call Resalta(Me.txtisc)
End If
End Sub

Private Sub TxtUnidades_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If KEY_PAIS = KEY_PERU Then
        Call Resalta(Me.TxtUnitario)
    Else
        If Me.DtcTipoDoc.BoundText = "0020" Then
            Me.DtcRetencionFuente.SetFocus
        Else
            Call Resalta(Me.TxtUnitario)
        End If
    End If
    
    
End If
End Sub

Private Sub TxtUnitario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim utilidad As Single
    Dim costo As Single
    Dim venta As Single
    Dim unidades As Single
    
    If Val(Me.txtCantidad.Text) < 0.01 Then
            MsgBox "Ingrese UNA CANTIDAD VALIDA", vbInformation, KEY_VENDEDOR
            Call Resalta(Me.txtCantidad)
            Exit Sub
        End If
    
    
    venta = Val(Me.txtPrecioVentaAnt.Text)
    unidades = Val(Me.TxtUnidades.Text)
    precio_unit = Val(Me.TxtUnitario.Text)
    
    
    If Val(Me.TxtUnitario.Text) <= 0 Then
        Call Resalta(Me.TxtUnitario)
        Exit Sub
    End If
    
    If Me.DtcMoneda.BoundText = "00002" And Me.chkConvertir.Value = 1 Then
        Me.TxtUnitario.Text = Format(Val(Me.TxtUnitario.Text) * Val(Me.txtTc.Text), "###0.0000")
    
    End If
    
     costo = Val(Me.TxtUnitario.Text)
     
     Me.TxtUnitario.Text = costo
     If Me.DtcMoneda.BoundText = "00002" And Me.DtcMoneda.BoundText <> KEY_MONEDA Then
        Me.TxtCostoHoy.Text = Format(costo * Val(Me.txtTc.Text), "###0.00")
     Else
        Me.TxtCostoHoy.Text = Format(costo, "###0.00")
     End If
     
     
    If Me.DtcTipoDoc.BoundText = "0002" Then
                Me.TxtValoerNeto.Text = Format(Val(Me.txtCantidad.Text) * costo, "###0.00")
                Me.txtRetencion.Text = 0
                Me.txtValorVenta.Text = 0
                Me.TxtIgv.Text = 0
                Me.txtisc.Text = 0
                If Me.chk_suspencion_retencion.Value = 1 Then
                   Me.txtprecioventa.Text = Format(Val(Me.txtCantidad.Text) * costo, "###0.00")
                   utilidad = 0
                   
                   Me.txtUtilidadhoy.Text = Format(utilidad, "###0.00")
                   costo = Val(Me.TxtCostoHoy.Text)
                   Me.TxtCostoHoy.Text = Format(costo, "###0.00")
                   utilidad = Val(Me.txtUtilidadhoy.Text)
                   venta = costo
                   Me.TxtventaHoy.Text = Format(venta, "###0.00")
                Else
                   Me.txtprecioventa.Text = Format(Val(Me.txtCantidad.Text) * costo, "###0.00")
                   Me.txtRetencion.Text = Format(Val(Me.txtprecioventa.Text) * 8 / 100, "###0.00")
                   Me.txtprecioventa.Text = Format(Val(Me.txtCantidad.Text) * costo, "###0.00")
                   utilidad = 0
                   Me.txtUtilidadhoy.Text = Format(utilidad, "###0.00")
                   costo = Val(Me.TxtCostoHoy.Text)
                   Me.TxtCostoHoy.Text = Format(costo, "###0.00")
                   utilidad = Val(Me.txtUtilidadhoy.Text)
                   venta = costo
                   Me.TxtventaHoy.Text = Format(venta, "###0.00")
                End If
                Me.txtValorVenta.Text = Format(Val(Me.txtprecioventa.Text), "###0.00")
                
    Else
            
            
            
            
            If Me.chkigv.Value = 1 Then
                
                Me.txtprecioventa.Text = Format(Val(Me.txtCantidad.Text) * costo, "###0.00")
                Me.TxtValoerNeto.Text = Format(Val(Me.txtprecioventa.Text) / (1 + KEY_IGV), "###0.00")
                Me.txtValorVenta.Text = Format(Val(Me.TxtValoerNeto.Text) - Val(Me.TxtTotalDescuento.Text), "##0.00")
                utilidad = 15 '(costo - costo) * 100 / costo
                Me.txtUtilidadhoy.Text = Format(utilidad, "###0.00")
                        If Me.DtcMoneda.BoundText = "00002" Then
                            costo = Val(Me.TxtCostoHoy.Text) * (1 + KEY_IGV) * Val(Me.txtTc.Text)
                        Else
                            costo = Val(Me.TxtCostoHoy.Text) * (1 + KEY_IGV)
                        End If
                
                Me.TxtCostoHoy.Text = Format(costo, "###0.00")
                utilidad = Val(Me.txtUtilidadhoy.Text)
                venta = costo + costo * utilidad / 100
                Me.TxtventaHoy.Text = Format(venta, "###0.00")
                Call calculo_igv
            Else
                Me.TxtValoerNeto.Text = Format(Val(Me.txtCantidad.Text) * costo, "###0.00")
                Me.txtprecioventa.Text = Format(Val(Me.txtCantidad.Text) * costo, "###0.00")
                Me.txtValorVenta.Text = Format(Val(Me.TxtValoerNeto.Text) - Val(Me.TxtTotalDescuento.Text), "##0.00")
                utilidad = 15 '(costo - costo) * 100 / costo
                Me.txtUtilidadhoy.Text = Format(utilidad, "###0.00")
                costo = Val(Me.TxtCostoHoy.Text)
                Me.TxtCostoHoy.Text = Format(costo, "###0.00")
                utilidad = Val(Me.txtUtilidadhoy.Text)
                venta = costo + costo * utilidad / 100
                Me.TxtventaHoy.Text = Format(venta, "###0.00")
            End If
   End If
    Call Resalta(Me.txtprecioventa)
End If




End Sub

Private Sub txtUtilidadhoy_Change()
    Call put_utilidad_compra
End Sub
Private Sub put_utilidad_compra()
Dim costo As Single
Dim venta As Single
Dim utilidad As Single
Dim venta_actual As Single


costo = Val(Me.TxtCostoHoy.Text)
utilidad = Val(Me.txtUtilidadhoy.Text)
venta = Val(Me.TxtGastoAdminHoy.Text) + costo + costo * utilidad / 100
If rever = False Then
    Me.TxtventaHoy.Text = Format(venta, "###0.00")
Else
    If Val(Me.TxtCostoHoy.Text) > 0 Then
    Me.TxtventaHoy.Text = Format(Val(Me.TxtCostoHoy.Text) + Val(Me.TxtCostoHoy.Text) * Val(Me.txtUtilidadhoy.Text) / 100, "#,##0.00")
    
    End If
End If

End Sub


Private Sub txtUtilidadhoy_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Me.TxtventaHoy.Text) = 0 Then
        Call put_utilidad_compra
    End If
    Call Resalta(Me.TxtventaHoy)
End If
End Sub

Private Sub TxtValoerNeto_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 37 Then
       Call Resalta(Me.txtisc)
 End If
End Sub

Private Sub TxtValoerNeto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.TxtValoerNeto.Text = Format(Val(Me.TxtValoerNeto.Text), "###0.00")
    Call Resalta(Me.TxtTotalDescuento)
End If
End Sub

Private Sub precio_venta_automatico()
Dim costo As Single
Dim utilidad As Single
costo = Val(Me.TxtCostoHoy.Text) + Val(Me.TxtGastoAdminHoy.Text)
venta = Val(Me.TxtventaHoy.Text)
If costo > 0 Then
    utilidad = (venta - costo) * 100 / costo
    Me.txtUtilidadhoy.Text = Format(utilidad, "###0.00")
    Me.TxtventaHoy.Text = venta
    Me.cmdAgregar.Enabled = True
    Me.cmdAgregar.SetFocus
End If

End Sub

Private Sub TxtventaHoy_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Val(Me.TxtventaHoy.Text) = 0 Then
        MsgBox "Ingrese un Monto Valido", vbInformation
        Exit Sub
    End If
    
    If Me.DtcTipo.BoundText = "01" Then
    If Val(Me.TxtventaHoy.Text) <> Val(Me.txtPrecioVentaAnt.Text) And Val(Me.txtPrecioVentaAnt.Text) <> 0 Then
        Procedencia = modificar_precio
        frmsegurity.Show
        Exit Sub
    Else
        Call precio_venta_automatico
    End If
    
    Else
        Me.cmdAgregar.Enabled = True
        Me.cmdAgregar.SetFocus
    End If
End If

End Sub




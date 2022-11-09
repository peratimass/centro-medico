VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmInventario 
   BorderStyle     =   0  'None
   Caption         =   "Inventario"
   ClientHeight    =   7350
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   11520
      TabIndex        =   84
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame frmVencimiento 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   8880
      TabIndex        =   69
      Top             =   4440
      Width           =   4575
      Begin VB.TextBox txtcantidad 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   1080
         MaxLength       =   80
         TabIndex        =   77
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtcantidad 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   1080
         MaxLength       =   80
         TabIndex        =   76
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtcantidad 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   1080
         MaxLength       =   80
         TabIndex        =   75
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtcantidad 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   1080
         MaxLength       =   80
         TabIndex        =   71
         Top             =   240
         Width           =   1215
      End
      Begin MSMask.MaskEdBox dtpCaduca 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   79
         ToolTipText     =   "dd/mm/yyyy"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin MSMask.MaskEdBox dtpCaduca 
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   80
         ToolTipText     =   "dd/mm/yyyy"
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin MSMask.MaskEdBox dtpCaduca 
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   81
         ToolTipText     =   "dd/mm/yyyy"
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin MSMask.MaskEdBox dtpCaduca 
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   82
         ToolTipText     =   "dd/mm/yyyy"
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin VB.Label lbltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   1080
         TabIndex        =   83
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Image cmdcerrar 
         Height          =   240
         Left            =   4200
         Picture         =   "FrmInventario.frx":0000
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD 4 :"
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
         TabIndex        =   74
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD 3 :"
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
         TabIndex        =   73
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD 2 :"
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
         TabIndex        =   72
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD 1 :"
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
         TabIndex        =   70
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.OptionButton opt_stock 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "CON STOCK"
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
      Height          =   195
      Left            =   11040
      TabIndex        =   68
      Top             =   2040
      Width           =   1575
   End
   Begin VB.OptionButton opt_todos 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "TODOS"
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
      Height          =   195
      Left            =   9720
      TabIndex        =   67
      Top             =   2040
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtOferta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6960
      MaxLength       =   80
      TabIndex        =   65
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtLote 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6960
      MaxLength       =   80
      TabIndex        =   61
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox txtCodigoBarra 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6960
      MaxLength       =   80
      TabIndex        =   60
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CheckBox chk_all 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "TODOS LAS SUCURSALES"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9720
      TabIndex        =   58
      Top             =   1320
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker DtpKardex 
      Height          =   315
      Left            =   10680
      TabIndex        =   55
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Format          =   245694465
      CurrentDate     =   43412
   End
   Begin VB.CheckBox chk_kardex 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "KARDEX"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9720
      TabIndex        =   54
      Top             =   1680
      Width           =   855
   End
   Begin VB.Frame FrameCaracteristicas 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   14760
      TabIndex        =   21
      Top             =   -120
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtitemdua 
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
         Left            =   2040
         TabIndex        =   29
         Top             =   4680
         Width           =   3735
      End
      Begin VB.TextBox txtañomodelo 
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
         Left            =   2040
         TabIndex        =   28
         Top             =   4320
         Width           =   3735
      End
      Begin VB.TextBox txtañocontenedor 
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
         Left            =   2040
         TabIndex        =   27
         Top             =   3960
         Width           =   3735
      End
      Begin VB.TextBox txtcontenedor 
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
         Left            =   2040
         TabIndex        =   26
         Top             =   3600
         Width           =   3735
      End
      Begin VB.TextBox txtañofabricacion 
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
         Left            =   2040
         TabIndex        =   25
         Top             =   3240
         Width           =   3735
      End
      Begin VB.TextBox txtseriechasis 
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
         Left            =   2040
         TabIndex        =   24
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtseriemotor 
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
         Left            =   2040
         TabIndex        =   23
         Top             =   2520
         Width           =   3735
      End
      Begin VB.TextBox txtnumeroserie 
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
         Left            =   2040
         TabIndex        =   22
         Top             =   2160
         Width           =   3735
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   345
         Left            =   2040
         TabIndex        =   30
         Top             =   5040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         BTYPE           =   5
         TX              =   "PROCESAR"
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
         MICON           =   "FrmInventario.frx":2EA4
         PICN            =   "FrmInventario.frx":2EC0
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
         Height          =   1815
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   3201
         _Version        =   393216
         ForeColor       =   8388608
         Rows            =   0
         FixedRows       =   0
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
      Begin VitekeySoft.ChameleonBtn cmdcerrarpantalla 
         Height          =   345
         Left            =   3960
         TabIndex        =   32
         Top             =   5040
         Width           =   1815
         _ExtentX        =   3201
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
         MICON           =   "FrmInventario.frx":345A
         PICN            =   "FrmInventario.frx":3476
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM :"
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
         Left            =   360
         TabIndex        =   40
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label Label56 
         BackStyle       =   0  'Transparent
         Caption         =   "AÑO MODELO :"
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
         Left            =   360
         TabIndex        =   39
         Top             =   4440
         Width           =   1140
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "AÑO CONTENEDOR :"
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
         Left            =   360
         TabIndex        =   38
         Top             =   4080
         Width           =   1530
      End
      Begin VB.Label Label54 
         BackStyle       =   0  'Transparent
         Caption         =   "N° CONTENEDOR :"
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
         Left            =   360
         TabIndex        =   37
         Top             =   3720
         Width           =   1365
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "AÑO FABRICACION :"
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
         Left            =   360
         TabIndex        =   36
         Top             =   3360
         Width           =   1605
      End
      Begin VB.Label Label52 
         BackStyle       =   0  'Transparent
         Caption         =   "N° CHASIS :"
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
         Left            =   360
         TabIndex        =   35
         Top             =   3000
         Width           =   915
      End
      Begin VB.Label Label51 
         BackStyle       =   0  'Transparent
         Caption         =   "N° MOTOR :"
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
         Left            =   360
         TabIndex        =   34
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° SERIE :"
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
         Left            =   360
         TabIndex        =   33
         Top             =   2280
         Width           =   765
      End
   End
   Begin VB.TextBox txtObservacion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   450
      Left            =   6960
      MaxLength       =   300
      TabIndex        =   49
      Top             =   6000
      Width           =   3615
   End
   Begin VB.CheckBox chk_sucursales 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "ACTUALIZAR TODAS LAS SUCURSALES"
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
      Height          =   495
      Left            =   8880
      TabIndex        =   46
      Top             =   3600
      Width           =   1935
   End
   Begin VitekeySoft.ChameleonBtn cmdStock 
      Height          =   555
      Left            =   1320
      TabIndex        =   19
      Top             =   6600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   "   ACTUALIZAR            "
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
      MICON           =   "FrmInventario.frx":648B
      PICN            =   "FrmInventario.frx":64A7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtid_producto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   7560
      MaxLength       =   80
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtStock_factura 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   405
      Left            =   2160
      MaxLength       =   80
      TabIndex        =   16
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox TxtVenta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6960
      MaxLength       =   80
      TabIndex        =   13
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox TxtCosto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   6960
      MaxLength       =   80
      TabIndex        =   12
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox TxtStock_nuevo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   405
      Left            =   2160
      MaxLength       =   80
      TabIndex        =   11
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox TxtStck_actual 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   405
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox TxtUnidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   8250
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox TxtDescripcionProducto 
      Appearance      =   0  'Flat
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
      Left            =   2070
      MaxLength       =   80
      TabIndex        =   1
      Top             =   1200
      Width           =   6135
   End
   Begin VB.TextBox TxtCodProducto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   360
      MaxLength       =   80
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   330
      Left            =   2040
      TabIndex        =   14
      Top             =   480
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   8421631
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
   Begin VitekeySoft.ChameleonBtn Command1 
      Height          =   555
      Left            =   3720
      TabIndex        =   20
      Top             =   6600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   "CERRAR PANTALLA"
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
      MICON           =   "FrmInventario.frx":8D91
      PICN            =   "FrmInventario.frx":8DAD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdSeriales 
      Height          =   555
      Left            =   6120
      TabIndex        =   41
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   "INGRESAR SERIALES"
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
      MICON           =   "FrmInventario.frx":90C7
      PICN            =   "FrmInventario.frx":90E3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcClasificacion 
      Height          =   330
      Left            =   2070
      TabIndex        =   43
      Top             =   1680
      Width           =   3975
      _ExtentX        =   7011
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
   Begin MSDataListLib.DataCombo DtcModelo 
      Height          =   330
      Left            =   2070
      TabIndex        =   44
      Top             =   2160
      Width           =   3975
      _ExtentX        =   7011
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
   Begin MSDataListLib.DataCombo DtcPeriodo 
      Height          =   330
      Left            =   2070
      TabIndex        =   47
      Top             =   2640
      Width           =   3975
      _ExtentX        =   7011
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
   Begin VitekeySoft.ChameleonBtn cmdReporte 
      Height          =   675
      Left            =   6360
      TabIndex        =   51
      Top             =   2280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1191
      BTYPE           =   5
      TX              =   "REPORTE FALTANTES SOBRANTES"
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
      MICON           =   "FrmInventario.frx":93FD
      PICN            =   "FrmInventario.frx":9419
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdstockfactura 
      Height          =   1035
      Left            =   9840
      TabIndex        =   52
      Top             =   240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1826
      BTYPE           =   5
      TX              =   "MIGRAR SALDO STOCK FACTURA"
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
      MICON           =   "FrmInventario.frx":B9EA
      PICN            =   "FrmInventario.frx":BA06
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdInventario10 
      Height          =   675
      Left            =   9720
      TabIndex        =   53
      Top             =   2340
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1191
      BTYPE           =   5
      TX              =   "REPORTE DE INVENTARIO             "
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
      MICON           =   "FrmInventario.frx":F13D
      PICN            =   "FrmInventario.frx":F159
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
      Height          =   555
      Left            =   11040
      TabIndex        =   56
      Top             =   5640
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   "ACTUALIZAR STOCK"
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
      MICON           =   "FrmInventario.frx":1172A
      PICN            =   "FrmInventario.frx":11746
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn ChameleonBtn2 
      Height          =   555
      Left            =   11040
      TabIndex        =   57
      Top             =   4920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   "LOCAL"
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
      MICON           =   "FrmInventario.frx":11A60
      PICN            =   "FrmInventario.frx":11A7C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSMask.MaskEdBox txtvencimiento 
      Height          =   330
      Left            =   6960
      TabIndex        =   64
      ToolTipText     =   "dd/mm/yyyy"
      Top             =   5640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
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
   Begin VitekeySoft.ChameleonBtn cmdDetallado 
      Height          =   315
      Left            =   8880
      TabIndex        =   78
      Top             =   5640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
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
      BCOL            =   33023
      BCOLO           =   33023
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmInventario.frx":11D96
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.OFERTA :"
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
      Left            =   6000
      TabIndex        =   66
      Top             =   4440
      Width           =   825
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENCIMIENTO  :"
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
      Left            =   5625
      TabIndex        =   63
      Top             =   5640
      Width           =   1200
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N° LOTE   :"
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
      Left            =   6045
      TabIndex        =   62
      Top             =   5280
      Width           =   780
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COD BARRA :"
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
      Left            =   5835
      TabIndex        =   59
      Top             =   4920
      Width           =   990
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F.FARMACOLOGICA :"
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
      Left            =   5280
      TabIndex        =   50
      Top             =   6120
      Width           =   1545
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODO :"
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
      Left            =   795
      TabIndex        =   48
      Top             =   2760
      Width           =   780
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MODELO :"
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
      Left            =   795
      TabIndex        =   45
      Top             =   2280
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLASIFICACION :"
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
      Left            =   360
      TabIndex        =   42
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK CONTABLE:"
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
      Left            =   360
      TabIndex        =   17
      Top             =   4440
      Width           =   1500
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   4200
      Y1              =   3720
      Y2              =   4560
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALMACEN"
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
      Left            =   480
      TabIndex        =   15
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.VENTA  :"
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
      Left            =   6045
      TabIndex        =   9
      Top             =   3960
      Width           =   780
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.COSTO :"
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
      Left            =   6090
      TabIndex        =   8
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK FÍSICO :"
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
      Left            =   675
      TabIndex        =   7
      Top             =   3960
      Width           =   1185
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "STOCK SISTEMA :"
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
      Left            =   495
      TabIndex        =   6
      Top             =   3480
      Width           =   1365
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   3975
      Left            =   240
      Top             =   3240
      Width           =   13335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNIDAD"
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
      Left            =   8280
      TabIndex        =   5
      Top             =   960
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE PRODUCTO"
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
      TabIndex        =   4
      Top             =   960
      Width           =   1590
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
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
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   3015
      Left            =   240
      Top             =   120
      Width           =   13335
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderWidth     =   4
      Height          =   7320
      Left            =   0
      Top             =   0
      Width           =   13770
   End
End
Attribute VB_Name = "FrmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede


Private Sub ChameleonBtn1_Click()
    
    strCadena = "SELECT * FROM view_producto WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and  ruc='" & KEY_RUC & "'  ORDER BY id_producto ASC,id_alm ASC"
    '5305
    'strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "'  ORDER BY id_producto ASC,id_alm ASC"
    Call ConfiguraRstIN(strCadena)
    If rstIN.RecordCount > 0 Then
       rstIN.MoveFirst
       For i = 0 To rstIN.RecordCount - 1
       DoEvents
        Me.txtid_producto.Text = rstIN("id_producto")
        Me.TxtCodProducto.Text = rstIN("id_producto")
        Me.TxtDescripcionProducto.Text = rstIN("nombre_prod")
        Me.TxtUnidad.Text = rstIN("unidad")
        Me.DtcAlmacen.BoundText = rstIN("id_alm")
        DoEvents
        Me.TxtStck_actual.Text = get_stock_kardex(Trim(Me.txtid_producto.Text), KEY_FECHA, rstIN("id_alm"))
        'Me.txtStock_factura.Text = get_stock_kardex_contable(Trim(Me.txtid_producto.Text), "2019-11-30", rstIN("id_alm"))
        
        
        
       
        
        
        Me.TxtVenta.Text = rstIN("precio_venta")
        Me.txtcosto.Text = rstIN("precio_compra")
        
        
        
        'strCadena = "SELECT * FROM producto_migrar WHERE id_producto='" & rstIN("id_producto") & "' and id_alm='" & rstIN("id_alm") & "'"
        'Call ConfiguraRstlocal(strCadena)
        'If rstLocal.RecordCount > 0 Then
        
            Me.TxtStock_nuevo.Text = 0 'rstLocal("stock_fisico")
            Me.txtStock_factura.Text = 0 'rstLocal("stock_contable")
            
            If Val(Me.TxtStck_actual.Text) <> 0 Then
                Call put_inventario
                
            End If
        
        
        
        Me.ChameleonBtn1.Caption = str(i) & Space(2) & str(rstIN.RecordCount)
        
       
        rstIN.MoveNext
        DoEvents
        
      Next i
    End If
        
End Sub
Private Sub migrar_maquina_local()
Dim sys_ConString2 As String
Dim stock_actual As Integer

strRuta_ini = App.Path & "\comparar_percy\producto.txt"

FileName = Dir(strRuta_ini)
fnum = FreeFile
   
   
    
    
    
strCadena = "SELECT * FROM view_producto WHERE  id_alm='00001' and   ruc='" & KEY_RUC & "' ORDER BY id_producto ASC "
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           Open strRuta_ini For Output As #1
           For j = 0 To rst.RecordCount - 1
              
              
              Print #1, rst("id_producto") & "," & rst("id_producto") & "," & Replace(rst("nombre_prod"), ",", " ") & "," & rst("precio_venta") & "," & "B"
               
        
   
                Me.ChameleonBtn1.Caption = str(rst("id_producto"))
                rst.MoveNext
                DoEvents
           Next j
           Close #1
          
        End If
End Sub




Private Sub ChameleonBtn2_Click()



Call migrar_maquina_local
End Sub

Private Sub chk_kardex_Click()
If Me.chk_kardex.Value = 1 Then
    Me.DtpKardex.Visible = True
Else
    Me.DtpKardex.Visible = False
End If
End Sub

Private Sub cmdCerrar_Click()
Me.frmvencimiento.Visible = False
End Sub

Private Sub cmdCerrarpantalla_Click()
Me.FrameCaracteristicas.Visible = False
End Sub

Private Sub cmdDetallado_Click()
Me.frmvencimiento.Visible = True
Call Resalta(Me.txtCantidad(0))
End Sub

Private Sub cmdInventario10_Click()

Dim param As Variant
Dim arr(0 To 1, 1 To 2) As String
arr(0, 1) = "fecha_ini"
arr(1, 1) = "fecha_fin"
arr(0, 2) = Format(Me.DtpKardex.Value, "dd-mm-YYYY")
arr(1, 2) = KEY_VENDEDOR
param = arr()

If Me.chk_kardex.Value = 1 Then
    
   
  
       If Me.chk_all.Value = 1 Then
            in_almacen = "TODAS LAS SUCURSALES"
            
       Else
            in_almacen = Me.DtcAlmacen.Text
            
       End If
       
       
       strCadena = "DELETE FROM producto_kardex WHERE dni_save='" & KEY_USUARIO & "'  and ruc='" & KEY_RUC & "'"
       CnBd.Execute (strCadena)
       
     
       
       
       strCadena = "call PUT_kardex_valorizado_itemv2('" & Format(Me.DtpKardex.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       If Me.opt_todos.Value = True Then
          in_parametro = ""
       Else
          in_parametro = " and stock<>0"
       End If
       
       
       
       If Me.chk_all.Value = 1 Then
            strCadena = "SELECT `id_producto`,`nombre_prod`,'" & in_almacen & "',sector,linea,`stock`,precio_costo,total FROM view_producto_kardex WHERE dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'" & in_parametro
       Else
            strCadena = "SELECT `id_producto`,`nombre_prod`,'" & in_almacen & "',sector,linea,`stock`,precio_costo,total FROM view_producto_kardex WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and  dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'" & in_parametro
       End If
       Call ConfiguraRst(strCadena)
       Ans = ShowMultiReport(rst, "RptInventario", param, App.Path + "\Reportes\")
       Exit Sub
       
       
       
       For i = 0 To rst.RecordCount - 1
            
            'strCadena = "call PUT_kardex_valorizado_item('" & rst("id_producto") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
            'CnBd.Execute (strCadena)
            'strCadena = "SELECT IFNULL(sum(cantidad_real),0) FROM kardex WHERE fecha_emision<='" & Format(Me.DtpKardex.Value, "YYYY-mm-dd") & "' and  id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC,id_kardex DESC LIMIT 1"
            'Call ConfiguraRstA(strCadena)
            
            'strCadena = "SELECT saldo_stock,costo_promedio FROM kardex WHERE fecha_emision<='" & Format(Me.DtpKardex.Value, "YYYY-mm-dd") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC,id_kardex DESC LIMIT 1"
            'Call ConfiguraRstL(strCadena)
            'If rstL.RecordCount > 0 Then
            '    in_costo_promedio = rstL(0)
            'Else
            '    in_costo_promedio = 0
            'End If
            
            
            
            'If rstA.RecordCount > 0 Then
             '   strCadena = "INSERT INTO producto_kardex(`id_producto`,`precio_costo`,`stock`,`dni_save`,`id_alm`,`ruc`)VALUES " & _
             '   "('" & rst("id_producto") & "','" & Val(in_costo_promedio) & "','" & rstA(0) & "','" & KEY_USUARIO & "','','" & KEY_RUC & "')"
             '   CnBd.Execute (strCadena)
            'End If
            rst.MoveNext
            DoEvents
            Me.cmdInventario10.Caption = str(i) & Space(2) & str(rst.RecordCount - 1)
       Next i
    End If
    
    strCadena = "SELECT `id_producto`,`nombre_prod`,'" & in_almacen & "','-',linea,`stock`,precio_costo,total FROM view_producto_kardex WHERE dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptInventario", param, App.Path + "\Reportes\")
    
    
'Else
    strCadena = "SELECT `id_producto`,`nombre_prod`,'" & Me.DtcAlmacen.Text & "',`sector`,linea,`stock`,precio_compra FROM view_producto WHERE stock>0 and id_linea not in('00009','00017') and  id_alm='" & Me.DtcAlmacen.BoundText & "' and  ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptInventario", , App.Path + "\Reportes\")
'End If
End Sub

Private Sub cmdReporte_Click()

strCadena = "SELECT `fecha_emision`,`id_producto`,`nombre_prod`,`Stock_sistema`,`Stock_conteo`,`sobrante`,`faltante`,`comentario`,`nombre_completo` FROM view_inventario WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "rpt_inventario_fisico", , App.Path + "\Reportes\")




End Sub

Private Sub cmdSeriales_Click()
'select l.`produccion` INTO produccion from producto p,linea l where p.`id_producto`=NEW.id_producto AND  p.`id_linea`=l.`id_linea` and p.`ruc`=l.`id_usu` AND p.`ruc`=new.ruc;
'if produccion='si' then
'my_loop:    Loop
 '       INSERT INTO `imp_producto_detalle`(id_compra,id_detalle_compra,id_alm,ruc)values(NEW.id_compra,NEW.id_detalle_compra,NEW.id_alm,NEW.ruc);
  '      if counter=new.cantidad then
   '         LEAVE my_loop;
    '    end IF;
        
     '   set counter=`counter`+1;
    'END LOOP my_loop;

'end if;

End Sub

Private Sub cmdstockfactura_Click()

strCadena = "SELECT * FROM almacen_producto WHERE ruc='" & KEY_RUC & "' ORDER BY id_producto ASC,id_alm ASC"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
   rstZ.MoveFirst
   For i = 0 To rstZ.RecordCount - 1
        
         Call put_saldo_factura(rstZ("id_producto"), rstZ("id_alm"))
         
         Call put_stock_factura(rstZ("id_producto"), rstZ("id_alm"))
         
         rstZ.MoveNext
         
         DoEvents
         Me.cmdstockfactura.Caption = str(i) & Space(2) & rstZ.RecordCount
   Next i
End If
End Sub
Private Sub put_saldo_factura(ByVal in_producto As String, ByVal in_alm As String)
 strCadena = "SELECT * FROM kardex k WHERE k.id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC"
 Call ConfiguraRstL(strCadena)
 If rstL.RecordCount > 0 Then
    rstL.MoveFirst
    For i = 0 To rstL.RecordCount - 1
        If rstL("cantidad_real") < 0 Then
            strCadena = "SELECT * FROM movimiento_venta where id_venta='" & rstL("id_movimiento") & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
                If rstT("afecta_factura") = "no" And rstT("id_doc") <> "0054" And rstT("id_doc") <> "0099" Then
                    strCadena = "UPDATE kardex SET cantidad_factura='" & rstL("cantidad_real") & "' WHERE id_kardex='" & rstL("id_kardex") & "'"
                    CnBd.Execute (strCadena)
                End If
            Else
                strCadena = "UPDATE kardex SET cantidad_factura='" & rstL("cantidad_real") & "' WHERE id_kardex='" & rstL("id_kardex") & "'"
                CnBd.Execute (strCadena)
            End If
        Else
            
            If rstL("id_doc") = "0089" Or rstL("id_doc") = "0001" Or rstL("id_doc") = "0003" Or rstL("id_doc") = "0009" Or rstL("id_doc") = "0007" Then
                strCadena = "UPDATE kardex SET cantidad_factura='" & rstL("cantidad_real") & "' WHERE id_kardex='" & rstL("id_kardex") & "'"
                CnBd.Execute (strCadena)
            Else
                X = 0
            End If
        End If
            
    
        rstL.MoveNext
       
    Next i
 End If
End Sub
Private Sub put_stock_factura(ByVal in_producto As String, ByVal in_alm As String)

strCadena = "SELECT * FROM view_producto WHERE id_producto='" & Trim(in_producto) & "' AND id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    cod_articulo = rst("id_producto")
    
    strCadena = "SELECT ifnull(sum(cantidad_real),0) FROM kardex WHERE  id_tipo_movimiento<>'10' and id_doc in('0001','0003','0007','0009','0054','0089','0090') and  id_producto='" & rst("id_producto") & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    stock_actual = rstT(0)
    
    strCadena = "UPDATE almacen_producto SET stock='" & Val(stock_actual) & "' WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    strCadena = "SELECT ifnull(sum(cantidad_factura),0) FROM kardex WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    stock_factura = rstT(0)
    
    strCadena = "UPDATE almacen_producto SET stock_factura='" & Val(stock_factura) & "' WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    
    Exit Sub
    
    strInventario = formato_item(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), 6)
    strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
    "('" & strInventario & "','" & cod_articulo & "','0106','001','" & strInventario & "','" & rst("precio_compra") & "','" & KEY_FECHA & "','" & in_almacen & "','" & Val(stock_actual) & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    in_cta_compra = KEY_CTA_COMPRA_SOLES
    
    If Val(Abs(stock_actual)) > 0 Then
       strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0089' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             in_numero = Format(1, "00000000")
             in_serie = "0001"
       End If
            in_cantidad = Val(stock_actual)
           
            strCadena = "call P_insert_compra_ultimate('0089','" & in_alm & "',CURDATE(),CURDATE(),'02'," & _
            "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
            "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
            "'0','0','0','0','0','0','0','0','0','0','0'," & _
            " '" & KEY_USUARIO & "','OBSERVACION','01','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
            Call ConfiguraRstP(strCadena)
            id_compra = rstP(0)
           
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(in_producto) & "','" & in_cantidad & "','0'," & _
           "'0','0','0','" & in_cantidad * Val(rst("precio_compra")) & "','0','0', " & _
           "'0','0','0','" & in_cantidad * Val(rst("precio_compra")) & "','0','" & Val(rst("precio_compra")) * in_cantidad & "','" & rst("precio_venta") & "','" & rst("precio_compra") & "','" & in_almacen & "','" & rst("nombre_prod") & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           strCadena = "call put_kardex_stock_vitekey_inventario('10','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0089','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(in_producto) & "','" & in_cantidad & "','" & Val(rst("precio_compra")) & "','" & in_alm & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
        
   
        
    
    'in_comentario = "INVENTARIO:" & Trim(Me.txtObservacion.Text) & Space(2) & KEY_VENDEDOR + Chr(13) + "CONTEO FISICO :" + str(Me.TxtStck_actual.Text) + Chr(13) + "AJUSTE :" + str(in_cantidad)
    
    strCadena = "UPDATE producto SET  inventario='si' WHERE id_producto='" & Trim(in_producto) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    
'If Me.chk_sucursales.Value = 0 Then
        'strCadena = "UPDATE almacen_producto SET precio_venta='" & rst("") & "',precio_compra='" & Val(Me.TxtCosto.Text) & "' WHERE id_producto='" & Trim(Me.TxtCodProducto.Text) & "' AND id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
    'Else
        'strCadena = "UPDATE almacen_producto SET precio_venta='" & Val(Me.TxtVenta.Text) & "',precio_compra='" & Val(Me.TxtCosto.Text) & "' WHERE id_producto='" & Trim(Me.TxtCodProducto.Text) & "' and ruc='" & KEY_RUC & "'"
    'End If
    
'CnBd.Execute (strCadena)

 End If
End If
End Sub


Private Sub Command1_Click()
Procedencia = Neutro
Unload Me
End Sub
Public Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub

Private Sub CmdPrecios_Click()

End Sub

Private Sub cmdStock_Click()

Call put_inventario

End Sub

Private Function get_stock_kardex(ByVal in_producto As String, ByVal in_fecha As Date, ByVal in_alm As String)

         
strCadena = "SELECT ifnull(sum(cantidad_real),0) FROM kardex WHERE fecha_emision<='" & Format(in_fecha, "YYYY-mm-dd") & "' and  id_tipo_movimiento<>'10' and id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstlocal(strCadena)
get_stock_kardex = rstLocal(0)

End Function

Private Function get_stock_kardex_contable(ByVal in_producto As String, ByVal in_fecha As Date, ByVal in_alm As String)
strCadena = "SELECT ifnull(sum(cantidad_factura),0) FROM kardex WHERE fecha_emision<='" & Format(in_fecha, "YYYY-mm-dd") & "' and  id_doc IN('0001','0003','0007','0009','0089','0090') and  id_producto='" & in_producto & "' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstlocal(strCadena)
get_stock_kardex_contable = rstLocal(0)
End Function


Private Sub put_inventario()
Dim strInventario As String
Dim in_cantidad As Double
Dim in_cantidad_contable As Double
Dim in_numero As String
    'If KEY_BARRAS = "si" Then
     '   strCadena = "SELECT A.id_producto,U.abreviatura,A.stock,P.nombre_prod,P.precio_venta,P.precio_compra FROM producto_barras B,producto P,unidad U,almacen_producto A WHERE B.id_producto=P.id_producto AND P.id_unidad=U.id_und AND B.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND A.id_producto=P.id_producto AND A.ruc='" & KEY_RUC & "' AND A.id_producto=P.id_producto AND B.id_producto=A.id_producto AND B.cod_barra='" & Trim(Me.TxtCodProducto.text) & "' AND A.id_alm='" & Me.DtcAlmacen.BoundText & "'"
    'Else


'If get_periodo_cierre(Me.DtcPeriodo.BoundText, "compras") = True Then
        
 '       MsgBox "PERIODO DE COMPRAS CERRARDO.!!!", vbInformation, KEY_VENDEDOR
 '       Exit Sub
        
 '    End If
     

strCadena = "SELECT A.id_producto,U.abreviatura,A.stock,P.nombre_prod,A.precio_venta,A.precio_compra FROM almacen_producto A,producto P,unidad U WHERE A.id_producto=P.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND A.id_producto='" & Trim(Me.txtid_producto.Text) & "' AND A.id_alm='" & Me.DtcAlmacen.BoundText & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 And Val(Me.TxtStock_nuevo.Text) >= 0 Then
    cod_articulo = rst("id_producto")
    stock_actual = Val(Me.TxtStck_actual.Text)
    
    strInventario = formato_item(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), 6)
    
    strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
    "('" & strInventario & "','" & cod_articulo & "','0106','001','" & strInventario & "','" & Val(Me.txtcosto.Text) & "','" & KEY_FECHA & "','" & Me.DtcAlmacen.BoundText & "','" & Val(Me.TxtStock_nuevo.Text) & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    in_cta_compra = KEY_CTA_COMPRA_SOLES
    
    If Len(Trim(Me.TxtCodigoBarra.Text)) > 2 Then
        strCadena = "DELETE  FROM producto_barras WHERE id_producto='" & Trim(Me.TxtCodProducto.Text) & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        strCadena = "INSERT INTO producto_barras(`id_producto`,`cod_barra`,`ruc`)VALUES('" & Trim(Me.TxtCodProducto.Text) & "','" & Trim(Me.TxtCodigoBarra.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
    End If
    
    
    strCadena = "UPDATE producto SET codigo_barra='" & Trim(Me.TxtCodigoBarra.Text) & "', vencimiento='" & Format(Me.txtvencimiento.Text, "YYYY-mm-dd") & "',lote='" & Trim(Me.txtLote.Text) & "',forma_farmacologica='" & Trim(UCase(txtObservacion.Text)) & "' WHERE id_producto='" & Trim(Me.TxtCodProducto.Text) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    
    
    If Val(Me.TxtStock_nuevo.Text) > Val(Me.TxtStck_actual) Then
       strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0089' and id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             
             strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "' and id_alm='" & Me.DtcAlmacen.BoundText & "' and id_doc='0089' LIMIT 1"
             Call ConfiguraRstlocal(strCadena)
             If rstLocal.RecordCount > 0 Then
               in_serie = rstLocal("serie")
               in_numero = formato_item(rstLocal("numero"), 8)
             Else
                MsgBox "CREE EL COMPROBANTE INGRESO A ALMACEN" + Chr(13) + "PARA ESTA SUCURSAL", vbInformation
                Exit Sub
             End If
                
             
       End If
            in_cantidad = Val(Me.TxtStock_nuevo.Text) - Val(Me.TxtStck_actual.Text)
           
             strCadena = "select funct_costo_ini('3','" & cod_articulo & "','" & KEY_ALM & "','" & KEY_RUC & "')"
             Call ConfiguraRstlocal(strCadena)
             If IsNull(rstLocal(0)) = True Then
                Me.txtcosto.Text = 0
            Else
                Me.txtcosto.Text = rstLocal(0)
             End If
             
             
             
             
             
        
            strCadena = "call P_insert_compra_ultimate('0089','" & Me.DtcAlmacen.BoundText & "','" & KEY_FECHA & "','" & KEY_FECHA & "','02'," & _
            "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
            "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
            "'0','0','0','0','0','0','0','0','0','0','0'," & _
            " '" & KEY_USUARIO & "','OBSERVACION','01','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
            Call ConfiguraRstP(strCadena)
            id_compra = rstP(0)
           
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(cod_articulo) & "','" & in_cantidad & "','" & Val(Me.txtcosto.Text) & "'," & _
           "'0','0','0','" & in_cantidad * Val(Me.txtcosto.Text) & "','0','0', " & _
           "'0','0','0','" & in_cantidad * Val(Me.txtcosto.Text) & "','0','" & Val(Me.txtcosto.Text) * in_cantidad & "','" & Val(Me.TxtVenta.Text) & "','" & Val(Me.txtcosto.Text) & "','" & Me.DtcAlmacen.BoundText & "','" & Trim(Me.TxtDescripcionProducto.Text) & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           strCadena = "call put_kardex_stock_v16('02','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & id_compra & "','0089','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & Trim(cod_articulo) & "','" & Val(Abs(in_cantidad)) & "','" & Val(Me.txtcosto.Text) & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','no','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
           
        
    Else
        strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0090' and id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
        Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             
             strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "' and id_alm='" & Me.DtcAlmacen.BoundText & "' and id_doc='0090' LIMIT 1"
             Call ConfiguraRstlocal(strCadena)
             If rstLocal.RecordCount > 0 Then
               in_serie = rstLocal("serie")
               in_numero = formato_item(rstLocal("numero"), 8)
             Else
                MsgBox "CREE EL COMPROBANTE SALIDA A ALMACEN" + Chr(13) + "PARA ESTA SUCURSAL", vbInformation
                Exit Sub
             End If
             
        End If
            
            in_cantidad = Val(Me.TxtStck_actual.Text) - Val(Me.TxtStock_nuevo.Text)
            
            
            If in_cantidad <> 0 Then
                
                strCadena = "select funct_costo_ini('3','" & cod_articulo & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_RUC & "')"
                Call ConfiguraRstlocal(strCadena)
                Me.txtcosto.Text = rstLocal(0)
             
                strCadena = "call P_insert_compra_ultimate('0090','" & Me.DtcAlmacen.BoundText & "','" & KEY_FECHA & "','" & KEY_FECHA & "','02'," & _
                "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
                "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
                "'0','0','0','0','0','0','0','0','0','0','0'," & _
                " '" & KEY_USUARIO & "','OBSERVACION','01','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
                Call ConfiguraRstP(strCadena)
                id_compra = rstP(0)
                strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
                "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(cod_articulo) & "','" & in_cantidad & "','" & Val(Me.txtcosto.Text) & "'," & _
                "'0','0','0','" & in_cantidad * Val(Me.txtcosto.Text) & "','0','0', " & _
                "'0','0','0','" & in_cantidad * Val(Me.txtcosto.Text) & "','0','" & Val(Me.txtcosto.Text) * in_cantidad & "','" & Val(Me.TxtVenta.Text) & "','" & Val(Me.txtcosto.Text) & "','" & Me.DtcAlmacen.BoundText & "','" & Trim(Me.TxtDescripcionProducto.Text) & "','0','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
           
                 strCadena = "call put_kardex_stock_v16('02','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & id_compra & "','0090','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & Trim(cod_articulo) & "','" & Val(Abs(in_cantidad)) & "','" & Val(Me.txtcosto.Text) & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','no','" & KEY_RUC & "')"
                 CnBd.Execute (strCadena)
            End If
            
            
            
        End If
        
        
  End If
   
    
    in_comentario = "INVENTARIO:" & Trim(Me.txtObservacion.Text) & Space(2) & KEY_VENDEDOR + Chr(13) + "CONTEO FISICO :" + str(Me.TxtStck_actual.Text) + Chr(13) + "AJUSTE :" + str(in_cantidad)
    strCadena = "UPDATE producto SET  inventario='si',comentario='" & in_comentario & "' WHERE id_producto='" & Trim(cod_articulo) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Call put_vencimiento(Trim(Me.TxtCodProducto.Text))
    
    
If Me.chk_sucursales.Value = 0 Then
        strCadena = "UPDATE almacen_producto SET precio_mayor='" & Val(Me.txtOferta.Text) & "',precio_venta='" & Val(Me.TxtVenta.Text) & "',precio_compra='" & Val(Me.txtcosto.Text) & "' WHERE id_producto='" & Trim(cod_articulo) & "' AND id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
    Else
        strCadena = "UPDATE almacen_producto SET precio_mayor='" & Val(Me.txtOferta.Text) & "',precio_venta='" & Val(Me.TxtVenta.Text) & "',precio_compra='" & Val(Me.txtcosto.Text) & "' WHERE id_producto='" & Trim(cod_articulo) & "' and ruc='" & KEY_RUC & "'"
    End If
    
    CnBd.Execute (strCadena)
    
    
 
    
    
  If KEY_SKFACTURA = "no" Then
    GoTo fin
  End If
  
 
    
    
    
strCadena = "SELECT A.id_producto,U.abreviatura,A.stock_factura as stock,P.nombre_prod,A.precio_venta,A.precio_compra FROM almacen_producto A,producto P,unidad U WHERE A.id_producto=P.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND A.id_producto='" & Trim(Me.txtid_producto.Text) & "' AND A.id_alm='" & Me.DtcAlmacen.BoundText & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    cod_articulo = rst("id_producto")
    
    stock_actual = get_stock_kardex_contable(cod_articulo, KEY_FECHA, Me.DtcAlmacen.BoundText)
    
    If Val(Me.txtStock_factura.Text) < 0 Then
        Exit Sub
    End If
    
    If Val(stock_actual) = Val(Me.txtStock_factura.Text) Then
        GoTo fin
        Exit Sub
    End If

    
    strInventario = formato_item(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), 6)
    
    strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
    "('" & strInventario & "','" & cod_articulo & "','0106','001','" & strInventario & "','" & Val(Me.txtcosto.Text) & "','" & KEY_FECHA & "','" & Me.DtcAlmacen.BoundText & "','" & Val(Me.txtStock_factura.Text) & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    in_cta_compra = KEY_CTA_COMPRA_SOLES
    
    If Val(Me.txtStock_factura.Text) > Val(stock_actual) Then
       strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0089' and id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             in_numero = Format(1, "00000000")
             in_serie = "0001"
       End If
            in_cantidad = Val(Me.txtStock_factura.Text) - Val(stock_actual)
             strCadena = "select funct_costo_ini('2','" & cod_articulo & "','" & KEY_ALM & "','" & KEY_RUC & "')"
             Call ConfiguraRstlocal(strCadena)
             Me.txtcosto.Text = rstLocal(0)
            
            strCadena = "call P_insert_compra_ultimate_v2('0089','" & Me.DtcAlmacen.BoundText & "','" & KEY_FECHA & "','" & KEY_FECHA & "','02'," & _
            "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
            "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
            "'0','0','0','0','0','0','0','0','0','0','0'," & _
            " '" & KEY_USUARIO & "','OBSERVACION','01','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','si','" & KEY_RUC & "')"
            Call ConfiguraRstP(strCadena)
            id_compra = rstP(0)
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(cod_articulo) & "','" & in_cantidad & "','" & Val(Me.txtcosto.Text) & "'," & _
           "'0','0','0','" & in_cantidad * Val(Me.txtcosto.Text) & "','0','0', " & _
           "'0','0','0','" & in_cantidad * Val(Me.txtcosto.Text) & "','0','" & Val(Me.txtcosto.Text) * in_cantidad & "','" & Val(Me.TxtVenta.Text) & "','" & Val(Me.txtcosto.Text) & "','" & Me.DtcAlmacen.BoundText & "','" & Trim(Me.TxtDescripcionProducto.Text) & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           strCadena = "call put_kardex_stock_v16('10','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & id_compra & "','0089','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & Trim(cod_articulo) & "','" & Val(Abs(in_cantidad)) & "','" & Val(Me.txtcosto.Text) & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','no','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
           
        
    Else
        strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0090' and id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
        Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             in_numero = Format(1, "00000000")
             in_serie = "0001"
       End If
            in_cantidad = Abs(Val(Me.txtStock_factura.Text) - Val(stock_actual))
            If in_cantidad <> 0 Then
               strCadena = "select funct_costo_ini('2','" & cod_articulo & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_RUC & "')"
               Call ConfiguraRstlocal(strCadena)
               Me.txtcosto.Text = rstLocal(0)
                strCadena = "call P_insert_compra_ultimate_v2('0090','" & Me.DtcAlmacen.BoundText & "','" & KEY_FECHA & "','" & KEY_FECHA & "','02'," & _
                "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
                "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
                "'0','0','0','0','0','0','0','0','0','0','0'," & _
                " '" & KEY_USUARIO & "','OBSERVACION','01','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','si','" & KEY_RUC & "')"
                Call ConfiguraRstP(strCadena)
                id_compra = rstP(0)
                
                strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
                "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(cod_articulo) & "','" & in_cantidad & "','" & Val(Me.txtcosto.Text) & "'," & _
                "'0','0','0','" & in_cantidad * Val(Me.txtcosto.Text) & "','0','0', " & _
                "'0','0','0','" & in_cantidad * Val(Me.txtcosto.Text) & "','0','" & Val(Me.txtcosto.Text) * in_cantidad & "','" & Val(Me.TxtVenta.Text) & "','" & Val(Me.txtcosto.Text) & "','" & Me.DtcAlmacen.BoundText & "','" & Trim(Me.TxtDescripcionProducto.Text) & "','0','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
           
                strCadena = "call put_kardex_stock_v16('10','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & id_compra & "','0090','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & Trim(cod_articulo) & "','" & Val(Abs(in_cantidad)) & "','" & Val(Me.txtcosto.Text) & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','no','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
            End If
            
            
            
        End If
        
        
  End If
   
    
    If Me.frmvencimiento.Visible = True Then
        
        
    End If
    
    
    in_comentario = "INVENTARIO:" & Trim(Me.txtObservacion.Text) & Space(2) & KEY_VENDEDOR + Chr(13) + "CONTEO FISICO :" + str(Me.TxtStck_actual.Text) + Chr(13) + "AJUSTE :" + str(in_cantidad)
    
    strCadena = "UPDATE producto SET  inventario='si',comentario='" & in_comentario & "' WHERE id_producto='" & Trim(cod_articulo) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    Call put_vencimiento(Trim(Me.TxtCodProducto.Text))
    
    
    
If Me.chk_sucursales.Value = 0 Then
        strCadena = "UPDATE almacen_producto SET precio_mayor='" & Val(Me.txtOferta.Text) & "', precio_venta='" & Val(Me.TxtVenta.Text) & "',precio_compra='" & Val(Me.txtcosto.Text) & "' WHERE id_producto='" & Trim(cod_articulo) & "' AND id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
    Else
        strCadena = "UPDATE almacen_producto SET precio_mayor='" & Val(Me.txtOferta.Text) & "', precio_venta='" & Val(Me.TxtVenta.Text) & "',precio_compra='" & Val(Me.txtcosto.Text) & "' WHERE id_producto='" & Trim(cod_articulo) & "' and ruc='" & KEY_RUC & "'"
    End If
    
    CnBd.Execute (strCadena)
    
fin:
    
  Me.TxtCodProducto.Text = ""
  Me.TxtDescripcionProducto.Text = ""
  Me.TxtUnidad.Text = ""
  Me.txtcosto.Text = ""
  Me.TxtVenta.Text = ""
  Me.txtOferta.Text = ""
  Me.TxtStck_actual.Text = ""
  Me.TxtStock_nuevo.Text = 0
  Me.txtStock_factura.Text = 0
  Me.txtObservacion.Text = ""
  Me.TxtCodigoBarra.Text = ""
  Me.txtLote.Text = ""
  
  Me.txtvencimiento.Mask = ""
  Me.txtvencimiento.Text = ""
  Me.txtvencimiento.Mask = "##/##/####"
  
  Me.cmdStock.Enabled = False
  Call Resalta(Me.TxtCodProducto)

End Sub

Private Sub put_vencimiento(ByVal in_producto As String)

For i = 0 To Me.txtCantidad.Count - 1
    If Val(Me.txtCantidad(i).Text) > 0 Then
        strCadena = "INSERT INTO producto_vencimiento(`id_producto`,`cantidad`,`vencimiento`,`ruc`)VALUES ('" & in_producto & "','" & Val(Me.txtCantidad(i).Text) & "','" & Format(Me.dtpCaduca(i).Text, "YYYY-mm-dd") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
    End If
        
Next i

strCadena = "SELECT vencimiento FROM producto_vencimiento WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' ORDER BY vencimiento ASC LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    strCadena = "UPDATE producto SET  vencimiento='" & Format(rstL("vencimiento"), "YYYY-mm-dd") & "' WHERE id_producto='" & Trim(in_producto) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
End If


For i = 0 To Me.txtCantidad.Count - 1
    Me.txtCantidad(i).Text = ""
    Me.dtpCaduca(i).Mask = ""
    Me.dtpCaduca(i).Text = ""
    Me.dtpCaduca(i).Mask = "##/##/####"
        
Next i


Me.frmvencimiento.Visible = False

End Sub

Private Sub Command2_Click()
strCadena = "SELECT * FROM producto_demo "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       
       in_habilitado = "si"
       
       If rst("status") = "NO VENDER" Then
          in_habilitado = "no"
       End If
       
       'strCadena = "UPDATE almacen_producto SET precio_compra='" & rst("costo") & "',precio_venta='" & rst("precio") & "',precio_alterno_a='" & rst("mercado") & "',precio_mayor='" & rst("mayorista") & "',habilitado='" & in_habilitado & "' WHERE id_producto='" & Format(rst("id_producto"), "00000") & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
       'CnBd.Execute (strCadena)
       
       
       strCadena = "UPDATE producto SET nombre_prod='" & UCase(Trim(rst("real"))) & "',nombre_comercial='" & UCase(Trim(rst("comercial"))) & "' WHERE id_producto='" & Format(rst("id_producto"), "00000") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
       CnBd.Execute (strCadena)
       
       DoEvents
       
       rst.MoveNext
       
       
       
   Next i
End If


End Sub

Private Sub dtpCaduca_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo salir


If KeyAscii = 13 Then
    
    If (Index + 1) < 3 Then
        Call Resalta(Me.txtCantidad(Index + 1))
    End If
    

End If
Exit Sub
salir:

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500

Me.frmvencimiento.Visible = False

strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'" & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  
strCadena = "SELECT id_linea as Codigo,descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcClasificacion)

strCadena = "SELECT id_tipo as Codigo,descripcion as Descripcion FROM linea_sub WHERE id_linea='" & Trim(Me.DtcClasificacion.BoundText) & "' AND  id_usu='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcModelo)


strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodo)
  Me.DtcPeriodo.BoundText = get_periodo_actual(KEY_FECHA)
  
  Me.DtpKardex.Value = KEY_FECHA
  
 
    Me.txtvencimiento.Mask = ""
    Me.txtvencimiento.Text = ""
    Me.txtvencimiento.Mask = "##/##/####"
  
  
  If KEY_USUARIO = "42546269" Then
    ChameleonBtn1.Visible = True

  End If
  
  
End Sub

Private Sub txtcantidad_Change(Index As Integer)
Dim Total As Single
Total = 0
If Val(Me.txtCantidad(Index).Text) > 0 Then
    For i = 0 To Me.txtCantidad.Count - 1
       Total = Total + Val(Me.txtCantidad(i).Text)
    Next i
    Me.lblTotal.Caption = Format(Total, "###0.00")
End If




End Sub

Private Sub txtCantidad_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 13 Then
    Me.dtpCaduca(Index).SetFocus
End If


End Sub

Private Sub TxtCodigoBarra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtLote)
End If
End Sub

Private Sub TxtCodProducto_KeyPress(KeyAscii As Integer)
Dim cantidad As Single
If KeyAscii = 13 Then
    
    If (Len(Me.TxtCodProducto.Text) = 0) Then
        
        Call Resalta(Me.TxtCodProducto)
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
    End If
      
   If KEY_BARRAS = "si" Then
        If Trim(Mid(Me.TxtCodProducto.Text, 1, 2)) = "00" And Len(Me.TxtCodProducto.Text) > 8 Then
            cantidad = Val(Mid(Trim(Me.TxtCodProducto.Text), 8, 4) / 1000)
            Me.TxtCodProducto.Text = Mid(Me.TxtCodProducto, 3, 5)
        End If
    End If
    
    If KEY_BARRAS = "si" Then
        strCadena = "SELECT A.id_producto,U.abreviatura,A.stock,P.nombre_prod,P.precio_venta,P.precio_compra,A.stock_factura FROM producto_barras B,producto P,unidad U,almacen_producto A WHERE B.id_producto=P.id_producto AND P.id_unidad=U.id_und AND B.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND A.id_producto=P.id_producto AND A.ruc='" & KEY_RUC & "' AND A.id_producto=P.id_producto AND B.id_producto=A.id_producto AND B.cod_barra='" & Trim(Me.TxtCodProducto.Text) & "' AND A.id_alm='" & Me.DtcAlmacen.BoundText & "'"
    Else
        Me.TxtCodProducto.Text = formato_item(Me.TxtCodProducto.Text, 5)
        If KEY_RUBRO = "00003" Then
            strCadena = "SELECT * FROM view_producto_farmacia WHERE  ruc='" & KEY_RUC & "' and  (id_producto='" & Trim(Me.TxtCodProducto.Text) & "' or codigo_barra='" & Trim(Me.TxtCodProducto.Text) & "') AND id_alm='" & Me.DtcAlmacen.BoundText & "'"
        Else
            strCadena = "SELECT * FROM view_producto WHERE  ruc='" & KEY_RUC & "' and  id_producto='" & Trim(Me.TxtCodProducto.Text) & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "'"
        End If
    End If
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.cmdStock.Enabled = True
        Me.txtid_producto.Text = rst("id_producto")
        Me.TxtDescripcionProducto.Text = rst("nombre_prod")
        Me.TxtUnidad.Text = rst("unidad")
        
        Me.TxtStck_actual.Text = get_stock_kardex(Trim(Me.txtid_producto.Text), KEY_FECHA, rst("id_alm"))
        Me.txtStock_factura.Text = get_stock_kardex_contable(Trim(Me.txtid_producto.Text), KEY_FECHA, rst("id_alm"))
        
        
        Me.txtcosto.Text = rst("precio_compra")
        Me.TxtVenta.Text = rst("precio_venta")
        Me.TxtStock_nuevo.Text = 0
        Call Resalta(Me.TxtStock_nuevo)
        Set rst = Nothing
        
    Else
        
        Call Resalta(Me.TxtCodProducto)
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
    End If
End If

End Sub

Private Sub txtcosto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    If KEY_RUBRO = "00003" Then
       Me.TxtVenta.Text = Val(Me.txtcosto.Text) * 1.34
       Me.txtOferta.Text = Val(Me.txtcosto.Text) * 1.34
    End If
    Call Resalta(Me.TxtVenta)
End If
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsDate(Trim(Me.txtvencimiento.Text)) = False Then
        Me.txtvencimiento.Text = CVDate(KEY_FECHA)
     End If
     Me.txtvencimiento.SetFocus
    
End If
End Sub

Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdStock.SetFocus
End If
End Sub

Private Sub txtOferta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtCodigoBarra)
End If
End Sub

Private Sub TxtStock_nuevo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtcosto)
End If
End Sub

Private Sub txtvencimiento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtObservacion)
    
End If
End Sub

Private Sub TxtVenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtOferta)
End If
End Sub

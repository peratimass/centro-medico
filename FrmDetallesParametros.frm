VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDetallesParametros 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   20055
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtVersion 
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
      Left            =   16800
      TabIndex        =   219
      Top             =   3080
      Width           =   495
   End
   Begin VB.CommandButton cmdupdateruta 
      Caption         =   "OK"
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
      Left            =   19320
      TabIndex        =   218
      Top             =   3060
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtRutaActualizar 
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
      Left            =   17400
      TabIndex        =   217
      Top             =   3080
      Width           =   1815
   End
   Begin VB.Frame frm_importacion 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5175
      Left            =   3720
      TabIndex        =   105
      Top             =   3045
      Visible         =   0   'False
      Width           =   11055
      Begin VB.CheckBox chk_adicionar_producto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "ADICIONAR"
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
         Left            =   120
         TabIndex        =   132
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox chk_sucursal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "SUCURSAL"
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
         Left            =   120
         TabIndex        =   107
         Top             =   120
         Width           =   1815
      End
      Begin VitekeySoft.ChameleonBtn Command12 
         Height          =   375
         Left            =   7080
         TabIndex        =   106
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
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
         MICON           =   "FrmDetallesParametros.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   315
         Left            =   2040
         TabIndex        =   108
         Top             =   120
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
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
      Begin VitekeySoft.ChameleonBtn cmdcollege 
         Height          =   375
         Left            =   9000
         TabIndex        =   109
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "IMPORTAR COLLEGE"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetallesParametros.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfproductos 
         Height          =   4095
         Left            =   120
         TabIndex        =   110
         Top             =   960
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   7223
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
      Begin VitekeySoft.ChameleonBtn cmdLoadvitekey 
         Height          =   375
         Left            =   5160
         TabIndex        =   111
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "LOAD VITEKEY"
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
         MICON           =   "FrmDetallesParametros.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdgetkeyfacil 
         Height          =   375
         Left            =   9000
         TabIndex        =   228
         Top             =   520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "GET_KEYFACIL"
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
         MICON           =   "FrmDetallesParametros.frx":0054
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
         Left            =   10680
         Picture         =   "FrmDetallesParametros.frx":0070
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.TextBox txtNombreComercial 
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
      Left            =   4680
      MaxLength       =   100
      TabIndex        =   196
      Top             =   720
      Width           =   2775
   End
   Begin VitekeySoft.ChameleonBtn cmdProcesar 
      Height          =   555
      Left            =   12120
      TabIndex        =   189
      Top             =   8040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   979
      BTYPE           =   3
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetallesParametros.frx":2F14
      PICN            =   "FrmDetallesParametros.frx":2F30
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   300
      Left            =   7800
      TabIndex        =   71
      Top             =   11160
      Visible         =   0   'False
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   529
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Forma Pago"
      TabPicture(0)   =   "FrmDetallesParametros.frx":6578
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "HfFormapago"
      Tab(0).Control(1)=   "DtcFormapago"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Hora de Trabajo"
      TabPicture(1)   =   "FrmDetallesParametros.frx":6594
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label23"
      Tab(1).Control(1)=   "txt"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Import Base Externa"
      TabPicture(2)   =   "FrmDetallesParametros.frx":65B0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label20"
      Tab(2).Control(1)=   "Label17"
      Tab(2).Control(2)=   "Label19"
      Tab(2).Control(3)=   "Label18"
      Tab(2).Control(4)=   "Command5"
      Tab(2).Control(5)=   "cmdimportartabla"
      Tab(2).Control(6)=   "Command2"
      Tab(2).Control(7)=   "TxtNombreTablaDestino"
      Tab(2).Control(8)=   "TxtCriterioOrigen"
      Tab(2).Control(9)=   "CmdImportar"
      Tab(2).Control(10)=   "TxtNombreBaseOrigen"
      Tab(2).Control(11)=   "TxtNombreTablaOrigen"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "IMPORTAR DATOS"
      TabPicture(3)   =   "FrmDetallesParametros.frx":65CC
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label28"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label27"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label26"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "ProgressBar1"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdHBS"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Command8"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Command7"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Command6"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "txtserver1"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "txttablaorigen"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "cmdimportarTabla1"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "txtbaseOrigen1"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).ControlCount=   12
      Begin VB.TextBox TxtNombreTablaOrigen 
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
         Left            =   -74160
         TabIndex        =   89
         Text            =   "Unidad"
         Top             =   1140
         Width           =   1335
      End
      Begin VB.TextBox TxtNombreBaseOrigen 
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
         Left            =   -74160
         TabIndex        =   88
         Text            =   "Base"
         Top             =   780
         Width           =   1335
      End
      Begin VB.CommandButton CmdImportar 
         Caption         =   "Importar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70800
         TabIndex        =   87
         Top             =   660
         Width           =   1335
      End
      Begin VB.TextBox TxtCriterioOrigen 
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
         Left            =   -74160
         TabIndex        =   86
         Text            =   "cUnidad"
         Top             =   1860
         Width           =   1335
      End
      Begin VB.TextBox TxtNombreTablaDestino 
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
         Left            =   -74160
         TabIndex        =   85
         Text            =   "unidad"
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Recalcular Productos Relacionados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72720
         TabIndex        =   84
         Top             =   1740
         Width           =   4215
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72840
         TabIndex        =   83
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CommandButton cmdimportartabla 
         Caption         =   "IMPORTAR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72720
         TabIndex        =   82
         Top             =   780
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "IMPORTAR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72720
         TabIndex        =   81
         Top             =   1260
         Width           =   1695
      End
      Begin VB.TextBox txtbaseOrigen1 
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
         Left            =   20
         TabIndex        =   80
         Text            =   "Base"
         Top             =   20
         Width           =   1335
      End
      Begin VB.CommandButton cmdimportarTabla1 
         Caption         =   "IMPORTAR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   79
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txttablaorigen 
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
         Left            =   1680
         TabIndex        =   78
         Text            =   "Tabla"
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txtserver1 
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
         Left            =   1680
         TabIndex        =   77
         Text            =   "Base"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   75
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   74
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   375
         Left            =   5520
         TabIndex        =   73
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdHBS 
         BackColor       =   &H008080FF&
         Caption         =   "MIGRAR HBS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   960
         Width           =   1575
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   960
         TabIndex        =   76
         Top             =   2160
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin MSDataListLib.DataCombo DtcFormapago 
         Height          =   315
         Left            =   -74880
         TabIndex        =   90
         Top             =   660
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfFormapago 
         Height          =   1455
         Left            =   -73080
         TabIndex        =   91
         Top             =   660
         Width           =   5295
         _ExtentX        =   9340
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
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base :"
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
         Left            =   -74670
         TabIndex        =   99
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla :"
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
         Left            =   -74730
         TabIndex        =   98
         Top             =   1140
         Width           =   525
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Criterio :"
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
         Left            =   -74850
         TabIndex        =   97
         Top             =   1860
         Width           =   645
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destino :"
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
         Left            =   -74880
         TabIndex        =   96
         Top             =   1500
         Width           =   675
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORAS DIARIAS:"
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
         Left            =   -74640
         TabIndex        =   95
         Top             =   1020
         Width           =   1365
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "BASE ORIGEN :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   390
         TabIndex        =   94
         Top             =   960
         Width           =   1140
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "TABLA :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   960
         TabIndex        =   93
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "SERVIDOR :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   92
         Top             =   600
         Width           =   885
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   5340
      Left            =   7800
      TabIndex        =   140
      Top             =   2640
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9419
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CONTABILIDAD"
      TabPicture(0)   =   "FrmDetallesParametros.frx":65E8
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frmcontable"
      Tab(0).Control(1)=   "chkContador"
      Tab(0).Control(2)=   "FrameContador"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "GRUPO EMPRESARIAL"
      TabPicture(1)   =   "FrmDetallesParametros.frx":6604
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmgrupoempresarial"
      Tab(1).Control(1)=   "cmdNuevo"
      Tab(1).Control(2)=   "hfgrupoempresarial"
      Tab(1).Control(3)=   "cmdEliminar"
      Tab(1).Control(4)=   "Label1"
      Tab(1).Control(5)=   "Shape10"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "+ PARAMETROS"
      TabPicture(2)   =   "FrmDetallesParametros.frx":6620
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Shape11"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "LblRazonSocial(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "chk_alerta_corte"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "chk_bonificaciones"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "chk_alarma_stock"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "chk_planes"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "chksegmentacion_precio"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame5"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "chk_sinfecto_caja"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "chk_grupo_empresarial"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "chk_mostrar_direccion"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "chk_incremento_zona"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtIncrementoPrecioZona"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txt_nota_credito_user"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtproveedor_servicio"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Chk_detalle_combo"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "chk_impresion_proformas"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "chk_bolsa_plastica"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txtImpuesto_bolsa"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "chk_tiendaOnline"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "chk_pago_efectivo"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "chk_pago_visa"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "chk_pago_mstercard"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "chk_pago_yape"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "chk_stock_reservado"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "chk_descuentos"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).ControlCount=   26
      Begin VB.CheckBox chk_descuentos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "DESCUENTOS"
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
         Height          =   375
         Left            =   4680
         TabIndex        =   242
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CheckBox chk_stock_reservado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "RESERVA DE STOCK"
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
         Height          =   375
         Left            =   4680
         TabIndex        =   240
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CheckBox chk_pago_yape 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "YAPE (BCP)"
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
         Height          =   250
         Left            =   3120
         TabIndex        =   238
         Top             =   4575
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chk_pago_mstercard 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "MASTERCARD"
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
         Height          =   250
         Left            =   3120
         TabIndex        =   237
         Top             =   4275
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chk_pago_visa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "VISA"
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
         Height          =   250
         Left            =   3120
         TabIndex        =   236
         Top             =   3975
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chk_pago_efectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "CONTADO"
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
         Height          =   250
         Left            =   3120
         TabIndex        =   235
         Top             =   3675
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chk_tiendaOnline 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "TIENDA ONLINE"
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
         Height          =   250
         Left            =   3120
         TabIndex        =   234
         Top             =   3360
         Width           =   2655
      End
      Begin VB.TextBox txtImpuesto_bolsa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   227
         Top             =   3050
         Width           =   855
      End
      Begin VB.CheckBox chk_bolsa_plastica 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         Caption         =   "IMP. BOLSA PLATICA"
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
         Left            =   3120
         TabIndex        =   226
         Top             =   3080
         Width           =   1695
      End
      Begin VB.CheckBox chk_impresion_proformas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "IMPRESION DE PROFORMAS VENDEDOR"
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
         Height          =   360
         Left            =   3120
         TabIndex        =   224
         Top             =   2640
         Width           =   2655
      End
      Begin VB.CheckBox Chk_detalle_combo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "DETALLE INSUMO COMBO"
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
         Left            =   3120
         TabIndex        =   220
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox txtproveedor_servicio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   215
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Frame frmcontable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "GENERAR CONTABILIDAD"
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
         Height          =   855
         Left            =   -74760
         TabIndex        =   210
         Top             =   4320
         Width           =   2775
         Begin VB.CommandButton Command3 
            BackColor       =   &H008080FF&
            Caption         =   "PROCESAR"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   211
            Top             =   600
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker DtpCini 
            Height          =   315
            Left            =   120
            TabIndex        =   212
            Top             =   240
            Width           =   1215
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
            Format          =   140247041
            CurrentDate     =   43262
         End
         Begin MSComCtl2.DTPicker DtpCfin 
            Height          =   315
            Left            =   1320
            TabIndex        =   213
            Top             =   240
            Width           =   1215
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
            Format          =   140247041
            CurrentDate     =   43262
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   2520
            Picture         =   "FrmDetallesParametros.frx":663C
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.TextBox txt_nota_credito_user 
         Height          =   285
         Left            =   -1080
         TabIndex        =   209
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtIncrementoPrecioZona 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   208
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox chk_incremento_zona 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         Caption         =   "INC PRECIO X ZONA"
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
         Left            =   3120
         TabIndex        =   207
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox chk_mostrar_direccion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         Caption         =   "MOSTRAR DIRECCION"
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
         Left            =   3120
         TabIndex        =   206
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CheckBox chk_grupo_empresarial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "GRUPO EMPRESARIAL"
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
         Left            =   3120
         TabIndex        =   204
         Top             =   840
         Width           =   2655
      End
      Begin VB.CheckBox chk_sinfecto_caja 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "SIN EFECTO CAJA Y BANCOS"
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
         Left            =   3120
         TabIndex        =   203
         Top             =   480
         Width           =   2655
      End
      Begin VB.Frame Frame5 
         Caption         =   "CIERRE KARDEX"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   240
         TabIndex        =   197
         Top             =   2280
         Width           =   2655
         Begin VB.TextBox txtCorrelativa 
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
            Left            =   600
            MaxLength       =   100
            TabIndex        =   201
            Top             =   600
            Width           =   1815
         End
         Begin MSDataListLib.DataCombo DtpPeriodo 
            Height          =   315
            Left            =   120
            TabIndex        =   198
            Top             =   960
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
         Begin VitekeySoft.ChameleonBtn cmdgenerarCierre 
            Height          =   345
            Left            =   120
            TabIndex        =   199
            Top             =   1320
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   609
            BTYPE           =   3
            TX              =   "GENERAR CIERRE"
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDetallesParametros.frx":94E0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComctlLib.ProgressBar progress_kardex 
            Height          =   220
            Left            =   120
            TabIndex        =   200
            Top             =   280
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   397
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VitekeySoft.ChameleonBtn ChameleonBtn3 
            Height          =   345
            Left            =   120
            TabIndex        =   202
            Top             =   1680
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   609
            BTYPE           =   3
            TX              =   "UPDATE SALDO"
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDetallesParametros.frx":94FC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn ChameleonBtn4 
            Height          =   345
            Left            =   120
            TabIndex        =   205
            Top             =   2160
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   609
            BTYPE           =   3
            TX              =   "CORREGIR RECEPCIONES"
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
            MICON           =   "FrmDetallesParametros.frx":9518
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
      Begin VB.CheckBox chksegmentacion_precio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "SEGMENTACION DE PRECIOS"
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
         Left            =   240
         TabIndex        =   194
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CheckBox chk_planes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "EMPRESA CON PLANES"
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
         Left            =   240
         TabIndex        =   193
         Top             =   1560
         Width           =   2655
      End
      Begin VB.CheckBox chk_alarma_stock 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "ALERTA DE STOCK BAJO"
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
         Left            =   240
         TabIndex        =   192
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CheckBox chk_bonificaciones 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "BONIFICACIONES"
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
         Left            =   240
         TabIndex        =   191
         Top             =   840
         Width           =   2655
      End
      Begin VB.CheckBox chk_alerta_corte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "ALERTA DE CORTE"
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
         Left            =   240
         TabIndex        =   188
         Top             =   480
         Width           =   2655
      End
      Begin VB.Frame frmgrupoempresarial 
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
         Height          =   2295
         Left            =   -74760
         TabIndex        =   184
         Top             =   840
         Visible         =   0   'False
         Width           =   5895
         Begin VB.TextBox txtrucVinculado 
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
            Left            =   960
            MaxLength       =   100
            TabIndex        =   185
            Top             =   480
            Width           =   1455
         End
         Begin VitekeySoft.ChameleonBtn cmdprocesarvinculado 
            Height          =   495
            Left            =   3960
            TabIndex        =   187
            Top             =   1320
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            BTYPE           =   3
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDetallesParametros.frx":9534
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label LblRazonSocial 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RUC :"
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
            Left            =   240
            TabIndex        =   230
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblgrupoempresa 
            BackColor       =   &H00C0C0C0&
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
            Left            =   960
            TabIndex        =   186
            Top             =   840
            Width           =   4665
         End
      End
      Begin VitekeySoft.ChameleonBtn cmdNuevo 
         Height          =   615
         Left            =   -68760
         TabIndex        =   182
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         BTYPE           =   3
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetallesParametros.frx":9550
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox chkContador 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -72960
         TabIndex        =   141
         Top             =   360
         Width           =   255
      End
      Begin VB.Frame FrameContador 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CONTABILIDAD INCLUIDA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   3855
         Left            =   -74880
         TabIndex        =   142
         Top             =   375
         Width           =   6975
         Begin VB.TextBox TxtRucContador 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            TabIndex        =   161
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtCuenta_cobrar_producto 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            TabIndex        =   160
            Text            =   "12121"
            Top             =   675
            Width           =   1215
         End
         Begin VB.TextBox txtCuenta_Cobrar_servicio 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5640
            TabIndex        =   159
            Text            =   "16811"
            Top             =   675
            Width           =   1095
         End
         Begin VB.TextBox txtCuenta_ingreso_producto 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2175
            TabIndex        =   158
            Text            =   "70111"
            Top             =   1635
            Width           =   1215
         End
         Begin VB.TextBox txtCuenta_ingreso_servicio 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5535
            TabIndex        =   157
            Text            =   "70411"
            Top             =   1635
            Width           =   1215
         End
         Begin VB.TextBox txtCompra_cta_pagar_soles 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1710
            TabIndex        =   156
            Text            =   "42121"
            Top             =   2100
            Width           =   975
         End
         Begin VB.TextBox txtCompra_cta_pagar_dolar 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1710
            TabIndex        =   155
            Text            =   "42122"
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox TxtCuenta_cobrar_rh 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5910
            TabIndex        =   154
            Text            =   "424"
            Top             =   2160
            Width           =   855
         End
         Begin VB.CheckBox chk_cta_pagar_asiento_global 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "CTAS PAGAR ASIENTO GLOBAL"
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
            Height          =   280
            Left            =   960
            TabIndex        =   153
            Top             =   3480
            Width           =   2655
         End
         Begin VB.CheckBox chk_linea_credito 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "INABILITAR LINEA DE CREDITO"
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
            Height          =   280
            Left            =   3720
            TabIndex        =   152
            Top             =   3480
            Width           =   3135
         End
         Begin VB.TextBox txtCuena_letraPagar_soles 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1710
            TabIndex        =   151
            Text            =   "0"
            Top             =   2820
            Width           =   975
         End
         Begin VB.TextBox txtCuena_letraPagar_dolares 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1710
            TabIndex        =   150
            Text            =   "0"
            Top             =   3120
            Width           =   975
         End
         Begin VB.TextBox txtCuena_fet_soles 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5910
            TabIndex        =   149
            Text            =   "0"
            Top             =   2580
            Width           =   855
         End
         Begin VB.TextBox txtCuena_fet_dolares 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5910
            TabIndex        =   148
            Text            =   "0"
            Top             =   2880
            Width           =   855
         End
         Begin VB.TextBox txtCuenta_cobrar_letra_soles 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            TabIndex        =   147
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtCuenta_cobrar_letra_dolares 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            TabIndex        =   146
            Top             =   1240
            Width           =   1215
         End
         Begin VB.TextBox txtCuenta_igv_ventas 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5655
            TabIndex        =   145
            Text            =   "40111"
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtCuentaigv_servicio 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3840
            TabIndex        =   144
            Text            =   "40111"
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txtcuenta_pagar_servicio 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3870
            TabIndex        =   143
            Text            =   "42121"
            Top             =   2520
            Width           =   855
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RUC  :"
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
            TabIndex        =   179
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblRazonContador 
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
            Height          =   300
            Left            =   2400
            TabIndex        =   178
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA COBRAR [PRODUCTO ] :"
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
            TabIndex        =   177
            Top             =   680
            Width           =   1845
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA COBRAR [SERVICIO ] :"
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
            TabIndex        =   176
            Top             =   705
            Width           =   1695
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA INGRESO [PRODUCTO ] :"
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
            TabIndex        =   175
            Top             =   1605
            Width           =   1905
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA INGRESO [SERVICIO ] :"
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
            Left            =   3570
            TabIndex        =   174
            Top             =   1605
            Width           =   1755
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA PAGAR [SOLES ] :"
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
            TabIndex        =   173
            Top             =   2160
            Width           =   1395
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA PAGAR [DOLAR ] :"
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
            TabIndex        =   172
            Top             =   2415
            Width           =   1455
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA PAGAR RH :"
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
            Left            =   4785
            TabIndex        =   171
            Top             =   2175
            Width           =   1065
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LETRA PAGAR [DOLAR]"
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
            Left            =   90
            TabIndex        =   170
            Top             =   3195
            Width           =   1515
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LETRA PAGAR [SOLES]"
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
            TabIndex        =   169
            Top             =   2820
            Width           =   1455
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FET [DOLARES] :"
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
            Left            =   4785
            TabIndex        =   168
            Top             =   3015
            Width           =   1065
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FET [SOLES] :"
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
            Left            =   4995
            TabIndex        =   167
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA COBRAR LETRA[SOLES]:"
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
            TabIndex        =   166
            Top             =   960
            Width           =   1845
         End
         Begin VB.Label Label39 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA COBRAR LETRA[DOLAR]:"
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
            TabIndex        =   165
            Top             =   1245
            Width           =   1905
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA IGV SERV :"
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
            Left            =   2850
            TabIndex        =   164
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA IGV:"
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
            Left            =   4515
            TabIndex        =   163
            Top             =   1110
            Width           =   585
         End
         Begin VB.Label Label41 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA PAG SERV :"
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
            Left            =   2850
            TabIndex        =   162
            Top             =   2520
            Width           =   1035
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   940
            Left            =   120
            Top             =   600
            Width           =   6735
         End
         Begin VB.Shape Shape9 
            BackColor       =   &H008080FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   1410
            Left            =   120
            Top             =   2040
            Width           =   6735
         End
         Begin VB.Shape Shape8 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   405
            Left            =   120
            Top             =   1560
            Width           =   6735
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgrupoempresarial 
         Height          =   2295
         Left            =   -74760
         TabIndex        =   180
         Top             =   840
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4048
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
      Begin VitekeySoft.ChameleonBtn cmdEliminar 
         Height          =   615
         Left            =   -68760
         TabIndex        =   183
         Top             =   1560
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         BTYPE           =   3
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetallesParametros.frx":956C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LblRazonSocial 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C000&
         Caption         =   "PROV SERV:"
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
         Index           =   2
         Left            =   3135
         TabIndex        =   214
         Top             =   1920
         Width           =   795
      End
      Begin VB.Shape Shape11 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   4935
         Left            =   120
         Top             =   360
         Width           =   6975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPRESAS VINCULADAS"
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
         Left            =   -74685
         TabIndex        =   181
         Top             =   480
         Width           =   1695
      End
      Begin VB.Shape Shape10 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   3735
         Left            =   -74880
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MIGRACION CONTABLE"
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
      Height          =   1820
      Left            =   15120
      TabIndex        =   133
      Top             =   6720
      Width           =   4815
      Begin VB.CheckBox chk_servicio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "VALIDAR SERVICIO"
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
         Height          =   375
         Left            =   3555
         TabIndex        =   225
         Top             =   240
         Width           =   1140
      End
      Begin VB.CheckBox chk_incompletos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "INCOMPLETOS"
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
         Height          =   375
         Left            =   3555
         TabIndex        =   223
         Top             =   1200
         Width           =   1140
      End
      Begin VB.CheckBox chk_sin_conta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "SIN ASIENTO"
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
         Height          =   375
         Left            =   3555
         TabIndex        =   222
         Top             =   720
         Width           =   1140
      End
      Begin VB.TextBox txtidventa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   139
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtRutaMigracion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   138
         Top             =   960
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DtpInicio_migracion 
         Height          =   315
         Left            =   120
         TabIndex        =   134
         Top             =   240
         Width           =   1250
         _ExtentX        =   2196
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
         Format          =   140247041
         CurrentDate     =   43262
      End
      Begin MSComCtl2.DTPicker DtpFin_migracion 
         Height          =   315
         Left            =   1450
         TabIndex        =   135
         Top             =   240
         Width           =   1250
         _ExtentX        =   2196
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
         Format          =   140247041
         CurrentDate     =   43262
      End
      Begin VitekeySoft.ChameleonBtn cmdrealizarmigracion 
         Height          =   450
         Left            =   120
         TabIndex        =   136
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   794
         BTYPE           =   5
         TX              =   "IMPORTAR EMPRESA"
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
         MICON           =   "FrmDetallesParametros.frx":9588
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
         Left            =   120
         TabIndex        =   137
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VitekeySoft.ChameleonBtn ChameleonBtn5 
         Height          =   1455
         Left            =   2715
         TabIndex        =   221
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   2566
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetallesParametros.frx":95A4
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
   Begin VB.CheckBox chk_Precio_cliente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "PRECIOS X CLIENTE"
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
      Left            =   4440
      TabIndex        =   131
      Top             =   5280
      Width           =   3015
   End
   Begin VB.CheckBox chk_moneda 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Left            =   4440
      TabIndex        =   128
      Top             =   5760
      Width           =   975
   End
   Begin VB.CheckBox chk_guia_fraccionada 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "GUIAS FRACCIONADAS"
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
      Left            =   4440
      TabIndex        =   127
      Top             =   1240
      Width           =   3015
   End
   Begin VB.CheckBox chk_codigo_universal_impresion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "COD UNIVERSAL EN IMPRESION:"
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
      Height          =   250
      Left            =   240
      TabIndex        =   125
      Top             =   1850
      Width           =   3735
   End
   Begin VB.CheckBox chk_keyfacil 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      Caption         =   "SERVIDOR KEYFACIL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Left            =   15240
      TabIndex        =   124
      Top             =   480
      Width           =   4110
   End
   Begin VB.TextBox txtToken_sucursal 
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
      Left            =   16800
      TabIndex        =   122
      Top             =   1980
      Width           =   2415
   End
   Begin VB.CheckBox chk_skin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "APLICACION DE SKIN'S"
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
      Left            =   4440
      TabIndex        =   120
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AUDITORIA GINSAC"
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
      Left            =   15240
      TabIndex        =   115
      Top             =   4200
      Width           =   4095
      Begin VB.CheckBox chk_update_kardex 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "FECHA"
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
         TabIndex        =   231
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H000080FF&
         Caption         =   "UPDATE STOCK,CONTABLE"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CommandButton Command32 
         Caption         =   "UPDATE COSTO"
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
         Left            =   2640
         TabIndex        =   121
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command27 
         Caption         =   "UPDATE SALDO"
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
         Left            =   2640
         TabIndex        =   119
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command30 
         Caption         =   "SALDO STOCK"
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
         Left            =   120
         TabIndex        =   118
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtidproducto 
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
         Left            =   2640
         TabIndex        =   117
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command26 
         Caption         =   "KARDEX ID"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   116
         Top             =   240
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DtpKardex 
         Height          =   315
         Left            =   1200
         TabIndex        =   232
         Top             =   840
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   140247041
         CurrentDate     =   43262
      End
   End
   Begin VB.CheckBox chk_referencia_comprobantes 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "REFERENCIA [CHASIS-MOTOR]"
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
      Left            =   4440
      TabIndex        =   114
      Top             =   7680
      Width           =   3015
   End
   Begin VB.CheckBox chk_grifo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "GRIFO COMBUSTIBLE"
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
      Left            =   4440
      TabIndex        =   113
      Top             =   1520
      Width           =   3015
   End
   Begin VB.CommandButton cmdventas_conta 
      Caption         =   "ACTIIVAR VENTAS CONTA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      TabIndex        =   104
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdcomprascontables 
      Caption         =   "ACTIVAR COMPRAS CONTA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   103
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdcorregirComprobantes 
      BackColor       =   &H008080FF&
      Caption         =   "MIGRAR STOCK Y ACTUALIZAR KARDEX"
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   7920
      Width           =   3015
   End
   Begin VB.CheckBox chk_stock_contable 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "STOCK CONTABLE  +  NO CONTABLE"
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
      Left            =   4440
      TabIndex        =   100
      Top             =   7440
      Width           =   3015
   End
   Begin VB.CheckBox chkStock_global 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   " STOCK GLOBAL EN CATALOGO"
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
      Left            =   4440
      TabIndex        =   70
      Top             =   7200
      Width           =   3015
   End
   Begin VB.TextBox txtPorcentajeCredito 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6120
      TabIndex        =   68
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CheckBox chk_notra_credito 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "NOTAS DE CREDITO [ADMIN]"
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
      Left            =   4440
      TabIndex        =   67
      Top             =   6525
      Width           =   3015
   End
   Begin VB.TextBox txttoken_local 
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
      Left            =   16800
      TabIndex        =   66
      Top             =   2355
      Width           =   2415
   End
   Begin VB.CheckBox chkServidorcloud 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "SERVIDOR CLOUD"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Left            =   15240
      TabIndex        =   65
      Top             =   160
      Width           =   4110
   End
   Begin VB.CheckBox chkGranel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "VENTAS A GRANEL"
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
      Left            =   4440
      TabIndex        =   64
      Top             =   6240
      Width           =   3015
   End
   Begin VB.CheckBox chk_update_proformas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "MODIFICAR PROFORMAS"
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
      Height          =   250
      Left            =   240
      TabIndex        =   63
      Top             =   4420
      Width           =   3735
   End
   Begin VB.CheckBox chk_precio_costo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "MOSTRAR PRECIO COSTO [PRODUCTO]"
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
      Height          =   250
      Left            =   240
      TabIndex        =   62
      Top             =   3800
      Width           =   3735
   End
   Begin VB.CheckBox chk_precio_mayor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "MOSTRAR PRECIO MAYOR [PRODUCTO]"
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
      Height          =   250
      Left            =   240
      TabIndex        =   61
      Top             =   3480
      Width           =   3735
   End
   Begin VB.CheckBox chk_mora_mensualidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "MONTO MORA [1 DIA ]"
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
      Height          =   280
      Left            =   240
      TabIndex        =   60
      Top             =   8220
      Width           =   2775
   End
   Begin VB.TextBox txtMora_monto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      TabIndex        =   59
      Top             =   8220
      Width           =   855
   End
   Begin VB.TextBox txtDiasCredito 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      TabIndex        =   58
      Top             =   7920
      Width           =   855
   End
   Begin VB.CheckBox chk_producto_duplicado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "PRODUCTO  [REPETIDO] VENTAS"
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
      Height          =   250
      Left            =   240
      TabIndex        =   56
      Top             =   7605
      Width           =   3735
   End
   Begin VB.CheckBox chk_entrega_mercaderia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "COTROL DE ENTREGA DE MERCADERIA"
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
      Height          =   250
      Left            =   240
      TabIndex        =   55
      Top             =   7275
      Width           =   3735
   End
   Begin VB.CheckBox chkMensaualidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "GENERACION MENSUALIDAD"
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
      Height          =   250
      Left            =   240
      TabIndex        =   54
      Top             =   6960
      Width           =   3735
   End
   Begin VB.CheckBox chkTransporte_migra 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "GUIAS REMISION [TRANSPORTISTA ]"
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
      Height          =   250
      Left            =   240
      TabIndex        =   53
      Top             =   6645
      Width           =   3735
   End
   Begin VB.CheckBox chk_envio_sunarp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "ENVIO XML SUNARP [CHASIS-MOTOR]"
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
      Height          =   250
      Left            =   240
      TabIndex        =   52
      Top             =   6330
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NIVEL EMPRESA"
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
      Left            =   4440
      TabIndex        =   48
      Top             =   3960
      Width           =   3015
      Begin VB.OptionButton OptPremiun 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Caption         =   "PAQUETE PREMIUN"
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
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   860
         Width           =   2655
      End
      Begin VB.OptionButton OptProfesional 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Caption         =   "PAQUETE PROFESIONAL"
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
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   540
         Width           =   2655
      End
      Begin VB.OptionButton OptStandart 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "PAQUETE STANDART"
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
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   220
         Width           =   2655
      End
   End
   Begin VB.CheckBox chk_cambio_precio_clave 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "CAMBIO PRECIO [ PASSWORD ADMIN]"
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
      Height          =   250
      Left            =   240
      TabIndex        =   47
      Top             =   6000
      Width           =   3735
   End
   Begin VB.CheckBox chk_proyecto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "PROYECTOS DE INVERSION"
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
      Height          =   250
      Left            =   240
      TabIndex        =   46
      Top             =   5685
      Width           =   3735
   End
   Begin VB.CheckBox chkvalidacion_clientes 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "VALIDACION EXTREMA [CLIENTES]"
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
      Height          =   250
      Left            =   240
      TabIndex        =   45
      Top             =   5370
      Width           =   3735
   End
   Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
      Height          =   315
      Left            =   7800
      TabIndex        =   44
      Top             =   8295
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "IMPORTAR EMPRESA"
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
      MICON           =   "FrmDetallesParametros.frx":95C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtprocentaje_detraccion 
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
      Left            =   18600
      TabIndex        =   43
      Top             =   2700
      Width           =   615
   End
   Begin VB.TextBox txtdetraccion 
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
      Left            =   16800
      TabIndex        =   42
      Top             =   2715
      Width           =   1455
   End
   Begin VB.CheckBox chk_seguro_venta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "SEGUROS EN VENTAS"
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
      Height          =   250
      Left            =   240
      TabIndex        =   40
      Top             =   4740
      Width           =   3735
   End
   Begin VB.CommandButton Command16 
      Caption         =   "PASAR TODAS LAS COMPRAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15120
      TabIndex        =   39
      Top             =   8160
      Width           =   4575
   End
   Begin VB.CommandButton Command15 
      Caption         =   "PASAR TODAS LAS VENTAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15120
      TabIndex        =   38
      Top             =   7800
      Width           =   4575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "LEER EXCEL FORMATO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15240
      TabIndex        =   37
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "MODIFICAR PRECIO VENTA A VALOR VENTA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15120
      TabIndex        =   36
      Top             =   7440
      Width           =   4575
   End
   Begin VB.TextBox txttoken 
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
      Left            =   16800
      TabIndex        =   34
      Top             =   1670
      Width           =   2415
   End
   Begin VB.CommandButton cmdmigrar 
      Caption         =   "MIGRAR BASE DE DATOS COMPLETA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15120
      TabIndex        =   32
      Top             =   7080
      Width           =   4575
   End
   Begin VB.CheckBox chkfacturacion_electronica 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "FACTURACION ELECTRONICA"
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
      Height          =   250
      Left            =   15240
      TabIndex        =   30
      Top             =   960
      Width           =   4110
   End
   Begin VB.TextBox txtresolucion 
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
      Left            =   16800
      TabIndex        =   29
      Top             =   1305
      Width           =   2415
   End
   Begin VB.CheckBox chkmodelo_color 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "MOSTRAR MODELO Y COLOR [PRODUCTO]"
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
      Height          =   250
      Left            =   240
      TabIndex        =   28
      Top             =   4110
      Width           =   3735
   End
   Begin VB.CheckBox chktrackin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "TRACKING VENTAS"
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
      Height          =   250
      Left            =   240
      TabIndex        =   27
      Top             =   1540
      Width           =   3735
   End
   Begin VB.CheckBox chkCajaIndependiente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "CAJA INDEPENDIENTE :"
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
      Left            =   11280
      TabIndex        =   26
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CheckBox chktramitedocumentario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "TRAMITE DOCUMENTARIO :"
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
      Left            =   8040
      TabIndex        =   25
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CheckBox chkActivacion 
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
      Height          =   255
      Left            =   12480
      TabIndex        =   23
      Top             =   300
      Width           =   255
   End
   Begin VB.Frame Frame1 
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
      Height          =   1335
      Left            =   12240
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CheckBox chkActivacionPermanente 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Activacion Permanente"
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
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DtpInstalacion 
         Height          =   300
         Left            =   1035
         TabIndex        =   20
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Format          =   140247041
         CurrentDate     =   41285
      End
      Begin MSComCtl2.DTPicker DtpCaducidad 
         Height          =   300
         Left            =   1035
         TabIndex        =   21
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Format          =   140247041
         CurrentDate     =   41285
      End
   End
   Begin VB.CheckBox ChkHuellaDigital 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "HUELLA DIGITAL"
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
      Left            =   8040
      TabIndex        =   16
      Top             =   510
      Width           =   2145
   End
   Begin VB.TextBox txtDireccionPublico 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4680
      TabIndex        =   15
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CheckBox chkfotoproducto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "FOTO PRODUCTO"
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
      Height          =   250
      Left            =   240
      TabIndex        =   13
      Top             =   3135
      Width           =   3735
   End
   Begin VB.CheckBox ChkUpdatePrecios 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "UPDATE PRECIOS BONIFICACIONES"
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
      Height          =   250
      Left            =   240
      TabIndex        =   12
      Top             =   5055
      Width           =   3735
   End
   Begin VB.OptionButton OptInventario 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   6000
      TabIndex        =   11
      Top             =   2895
      Width           =   1335
   End
   Begin VB.OptionButton OptContabilidad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CONTABILIDAD"
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
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   2895
      Width           =   1455
   End
   Begin VB.CheckBox ChkCerveceria 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "MANUAL"
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
      Left            =   8040
      TabIndex        =   9
      Top             =   180
      Width           =   2145
   End
   Begin VB.CheckBox ChkAutomatico 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "VENTAS AUTOMATICAS"
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
      Height          =   250
      Left            =   240
      TabIndex        =   8
      Top             =   2805
      Width           =   3735
   End
   Begin VB.CheckBox ChkBarras 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "CODIGO DE BARRA"
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
      Height          =   250
      Left            =   240
      TabIndex        =   6
      Top             =   2475
      Width           =   3735
   End
   Begin VB.CheckBox ckkFacturas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "STOCK FACTURAS"
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
      Height          =   250
      Left            =   240
      TabIndex        =   5
      Top             =   2145
      Width           =   3735
   End
   Begin VB.TextBox TxtEmpresa 
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
      Left            =   1605
      MaxLength       =   100
      TabIndex        =   3
      Top             =   380
      Width           =   3015
   End
   Begin VB.TextBox TxtRuc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1605
      MaxLength       =   20
      TabIndex        =   2
      Top             =   50
      Width           =   1455
   End
   Begin VB.TextBox TxtDireccion 
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
      Left            =   1605
      MaxLength       =   100
      TabIndex        =   1
      Top             =   705
      Width           =   3015
   End
   Begin VB.CheckBox chkigv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "AFECTO A IGV"
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
      Height          =   250
      Left            =   240
      TabIndex        =   0
      Top             =   1245
      Width           =   3735
   End
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   315
      Left            =   4560
      TabIndex        =   7
      Top             =   1995
      Width           =   2775
      _ExtentX        =   4895
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
   Begin MSDataListLib.DataCombo DtcTipoLetra 
      Height          =   315
      Left            =   8040
      TabIndex        =   17
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgSeries 
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   2355
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
   Begin ComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   15120
      TabIndex        =   33
      Top             =   6800
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VitekeySoft.ChameleonBtn ChameleonBtn2 
      Height          =   255
      Left            =   9720
      TabIndex        =   102
      Top             =   8040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      BTYPE           =   5
      TX              =   "CORREGIR SALDOS"
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
      MICON           =   "FrmDetallesParametros.frx":95DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcPais 
      Height          =   315
      Left            =   4680
      TabIndex        =   112
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VitekeySoft.ChameleonBtn cmdUpdate_stock_migracion 
      Height          =   345
      Left            =   9720
      TabIndex        =   126
      Top             =   8260
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "ACTUALIZAR SALDO EMPRESA MIGRADA"
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
      MICON           =   "FrmDetallesParametros.frx":95F8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo dtcMoneda 
      Height          =   315
      Left            =   5520
      TabIndex        =   129
      Top             =   5640
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
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   555
      Left            =   13560
      TabIndex        =   190
      Top             =   8040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   979
      BTYPE           =   3
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetallesParametros.frx":9614
      PICN            =   "FrmDetallesParametros.frx":9630
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   8040
      TabIndex        =   233
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "TOKEN SUCURSAL:"
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
      Index           =   1
      Left            =   15240
      TabIndex        =   241
      Top             =   2400
      Width           =   1425
   End
   Begin VB.Label LblRazonSocial 
      Alignment       =   2  'Center
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
      Index           =   5
      Left            =   240
      TabIndex        =   239
      Top             =   480
      Width           =   1065
   End
   Begin VB.Label LblRazonSocial 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   18360
      TabIndex        =   229
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label LblRazonSocial 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ACTUALIZADOR:"
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
      Left            =   15345
      TabIndex        =   216
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION FISCAL :"
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
      TabIndex        =   195
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOKEN SUCURSAL:"
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
      Index           =   0
      Left            =   15285
      TabIndex        =   123
      Top             =   2040
      Width           =   1425
   End
   Begin VB.Label lbldni_encargado 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "% CREDITO :"
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
      Height          =   225
      Left            =   4440
      TabIndex        =   69
      Top             =   6840
      Width           =   1545
   End
   Begin VB.Label Label25 
      BackColor       =   &H000080FF&
      Caption         =   "DIAS DE CREDITO [CREDITO] :"
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
      Height          =   270
      Left            =   240
      TabIndex        =   57
      Top             =   7920
      Width           =   2835
   End
   Begin VB.Label Label2 
      BackColor       =   &H008080FF&
      Caption         =   "CTA DETRACCION:"
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
      Height          =   735
      Left            =   15240
      TabIndex        =   41
      Top             =   2685
      Width           =   4575
   End
   Begin VB.Label Label30 
      BackColor       =   &H008080FF&
      Caption         =   "TOKEN CLOUD:"
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
      Height          =   675
      Left            =   15240
      TabIndex        =   35
      Top             =   1635
      Width           =   4095
   End
   Begin VB.Label Label50 
      BackColor       =   &H008080FF&
      Caption         =   "RESOLUCION:"
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
      Height          =   345
      Left            =   15240
      TabIndex        =   31
      Top             =   1260
      Width           =   4095
   End
   Begin VB.Shape Shape21 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8450
      Left            =   15000
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "ACTIVACION"
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
      Height          =   230
      Left            =   12690
      TabIndex        =   24
      Top             =   315
      Width           =   1815
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direccion Publico en General"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4650
      TabIndex        =   14
      Top             =   3360
      Width           =   2325
   End
   Begin VB.Shape Shape15 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   600
      Left            =   4440
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   400
      Left            =   4440
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   975
      Left            =   4440
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label LblRazonSocial 
      Alignment       =   2  'Center
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
      Index           =   0
      Left            =   450
      TabIndex        =   4
      Top             =   180
      Width           =   345
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1100
      Left            =   120
      Top             =   50
      Width           =   7575
   End
   Begin VB.Shape Shape17 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   375
      Left            =   12240
      Top             =   240
      Width           =   2535
   End
   Begin VB.Shape Shape19 
      BorderColor     =   &H00808080&
      Height          =   495
      Left            =   7920
      Top             =   2040
      Width           =   6855
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2475
      Left            =   7800
      Top             =   120
      Width           =   7335
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   7395
      Left            =   120
      Top             =   1200
      Width           =   7575
   End
   Begin VB.Shape Shape22 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   8640
      Left            =   0
      Top             =   0
      Width           =   20055
   End
End
Attribute VB_Name = "FrmDetallesParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Dim strCodAlmacen As String
Dim strSerie As String
Public cnbd1 As New ADODB.Connection


Private Sub put_cambio()
strCadena = "SELECT * from tipo_cambio WHERE id_creador='20487725286' and fecha>='2018-05-01'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "INSERT INTO tipo_cambio(`descripcion`,`fecha`,`valor`,`valor_venta`,`valor_compra`,`valor_venta1`,`valor_local`,`id_creador`)VALUES " & _
       "('" & rst("descripcion") & "','" & Format(rst("fecha"), "YYYY-mm-dd") & "','" & rst("valor") & "','" & rst("valor_venta") & "','" & rst("valor_compra") & "','" & rst("valor_venta1") & "','" & rst("valor_local") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If
End Sub

Private Sub put_inventarionn()

strCadena = "SELECT * FROM  producto WHERE ruc='" & KEY_RUC & "' ORDER BY id_producto"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   For i = 0 To rstA.RecordCount - 1
        
        strCadena = "SELECT * FROM inventario WHERE id_producto='" & rstA("id_producto") & "' and  fecha>='2018-11-27' and ruc='" & KEY_RUC & "' and id_alm='00003' LIMIT 1"
        Call ConfiguraRstlocal(strCadena)
        If rstLocal.RecordCount < 1 Then
            Call put_inventario(rstA("id_producto"), "00003", 0)
        End If
        DoEvents
   
        rstA.MoveNext
   Next i
   
   
End If

End Sub


Private Sub update_kardex_olivos()

strCadena = "SELECT * FROM almacen_producto WHERE ruc='" & KEY_RUC & "' ORDER BY id_producto ASC"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "UPDATE almacen_producto SET precio_compra='" & rst2("precio_compra") & "'  WHERE id_producto='" & rst2("id_producto") & "' and id_alm='" & rst2("id_alm") & "' and ruc='" & KEY_RUC & "' "
       CnBd.Execute (strCadena)
       
       strCadena = "UPDATE kardex SET costo_unitario='" & rst2("precio_compra") & "',costo_promedio='" & rst2("precio_compra") & "'  WHERE id_producto='" & rst2("id_producto") & "' and id_alm='" & rst2("id_alm") & "' and ruc='" & KEY_RUC & "' "
       CnBd.Execute (strCadena)
       
       
       rst2.MoveNext
   Next i
End If

End Sub
Private Sub put_cliente(ByVal in_dni As String, ByVal in_cliente As String)
        
        If in_cliente = "A    N    U    L    A    D    O" Then
            Exit Sub
        End If
        strCadena = "SELECT * FROM persona WHERE dni='" & Trim(in_dni) & "' LIMIT 1"
        Call ConfiguraRstT(strCadena)
            If rstT.RecordCount < 1 Then
                
                strCadena = "call P_insert_persona_ii('" & Trim(in_dni) & "' " & _
                ",'-', " & _
                "'-' " & _
                ",'-' " & _
                ",'" & Replace(UCase(Trim(in_cliente)), "'", " ") & "' " & _
                ",'-' " & _
                ",'-' " & _
                ",'-'" & _
                ",'no' " & _
                ",'no'" & _
                ",'si' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                 
   End If
                
                    
        strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & in_dni & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
                    Call ConfiguraRstT(strCadena)
                    If rstT.RecordCount < 1 Then
                        strCadena = "INSERT INTO entidad_empresa(cod_unico,id_empresa,id_almacen,passwordaccesso)VALUES ('" & in_dni & "','" & KEY_RUC & "','no','" & in_dni & "')"
                        CnBd.Execute (strCadena)
                    End If
      
End Sub

Private Sub importar_ventas_n1()

GoTo migrar
strCadena = "select a.`documento`,a.`fecha_emision`,id_venta from movimiento_venta a Where id_doc IN('0001','0003','0007') and a.`ruc`='20487473881' and a.`fecha_emision`>='2019-11-01' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For j = 0 To rst.RecordCount - 1
      
                    
                    strCadena = "DELETE FROM movimiento_venta WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                    CnBd.Execute (strCadena)
                    
                    
                    
                    
                    strCadena = "select Id from con_documento   where IdEmpresaSis='" & KEY_RUC & "' and  IdReferencia = '" & rst("id_venta") & "' LIMIT 1"
                    Call ConfiguraRstL(strCadena)
                    If rstL.RecordCount > 0 Then
                        
                        strCadena = "UPDATE con_documento set Activo = 0 where id = '" & rstL("id") & "' and IdEmpresaSis='" & KEY_RUC & "'"
                        CnBd.Execute (strCadena)
                        
                        strCadena = "update con_venta set Activo = 0   where IdDocumento = '" & rstL("id") & "' and IdEmpresaSis='" & KEY_RUC & "'"
                        CnBd.Execute (strCadena)
                        
                    End If
                    
                    
                   
                   
            
   
   rst.MoveNext
   DoEvents
   Me.ChameleonBtn1.Caption = str(j) & Space(5) & str(rst.RecordCount)
   Next j
   
End If

Exit Sub
migrar:

'1CIX000000000137
strCadena = "select * from con_asiento a where glosa like '%COBRO%' and  Activo='1' and a.`IdTipoAsiento`='1CIX000000000053' and a.`IdEmpresaSis`='20487473881' and a.`IdPeriodo`='1CIX000000000050' ORDER BY id"
'strCadena = "select * from con_asiento a where   Activo='1' and a.`IdTipoAsiento`='1CIX000000000137' and a.`IdEmpresaSis`='20487473881' and a.`IdPeriodo`='1CIX000000000050' ORDER BY id"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        
        
        
        
        strCadena = "DELETE FROM con_asiento WHERE  id='" & rst("id") & "' and IdEmpresaSis='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
        
        strCadena = "SELECT id FROM con_asientomovimiento WHERE idAsiento='" & rst("id") & "' and IdEmpresaSis='" & KEY_RUC & "' "
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
           rstK.MoveFirst
           For j = 0 To rstK.RecordCount - 1
               'strCadena = "UPDATE CON_MovimientoCajaBanco SET FechaModifica = NOW(), UsuarioModifica = '" & KEY_USUARIO & "',   Activo = 0 WHERE IdAsientoMovimiento='" & rstK("id") & "' and IdEmpresaSis='" & KEY_RUC & "' "
               ' CnBd.Execute (strCadena)
        
                'strCadena = "UPDATE CON_AsientoMovimiento_Documento SET Activo = 0, FechaModifica = NOW(), UsuarioModifica = '" & KEY_USUARIO & "'  WHERE IdAsientoMovimiento='" & rstK("id") & "' and IdEmpresaSis='" & KEY_RUC & "' "
                'CnBd.Execute (strCadena)
               
                strCadena = "DELETE FROM con_asientomovimiento WHERE id='" & rstK("id") & "' and IdEmpresaSis='" & KEY_RUC & "' LIMIT 1"
                CnBd.Execute (strCadena)
               
               rstK.MoveNext
           Next j
        End If
        
        
        
        
        
        rst.MoveNext
        DoEvents
        Me.ChameleonBtn1.Caption = str(i) & Space(5) & str(rst.RecordCount)
   Next i
End If

Exit Sub


End Sub




Private Sub ChameleonBtn1_Click()
    
MsgBox "Importar ventas"
 
 Call migrar_empresa_local
 Exit Sub
 
 Call migrar_empresa_online
'Call importar_ventas_n1
Exit Sub
    
    
MsgBox "Importar Clientes"
    
    strCadena = "SELECT * FROM producto_sabar"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       For i = 0 To rst.RecordCount - 1
           Call put_cliente(rst("id_proveedor"), rst("proveedor"))
           rst.MoveNext
       Next i
       
    End If
    


Exit Sub
MsgBox "Importando Fotos"
    
 strCadena = "SELECT * FROM producto WHERE ruc='" & KEY_RUC & "'"
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        strCadena = "UPDATE producto SET imagen='" & Trim(rst("imagen")) & "' WHERE id_producto='" & rst("id_producto") & "' and ruc IN('20539068262','10411203536','10175307678')"
        CnBd.Execute (strCadena)
        rst.MoveNext
        
    Next i
 End If
    Exit Sub
    
If MsgBox("Desea Realizar el proceso", vbQuestion + vbYesNo, KEY_VENDEDOR) = vbNo Then
    Exit Sub
End If

   ' Call put_inventarionn
   ' Call migrar_empresa_local
   'Call put_cambio
    
    Call migrar_empresa_online
    
    'Call delete_espacio

Exit Sub
'[1] importar sus Almacenes.
strCadena = "DELETE FROM almacen WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM almacen WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO almacen(`id_alm`,`ruc`,`descripcion`,`abreviatura`,`direccion`,`ubicacion_interna`,`id_responsable`,`id_tipoentidad`,`piso`,`pabellon`,`id_especialidad`,`observacion`,`stock`,`stock_personalizado`,`hora_inicio`,`hora_fin`,`horas`,`defecto`,`activo`,`id_atension`,`id_actividad`,`dni_save`,`ocupado`,`id_sucursal`,`facturacion_detallada`,`facturacion_centralizada`,`caja_independiente`,`cloud`,`comprobante_adicional`)VALUES " & _
       "('" & rst2("id_alm") & "','" & rst2("ruc") & "','" & rst2("descripcion") & "','" & rst2("abreviatura") & "','" & rst2("direccion") & "','" & rst2("ubicacion_interna") & "','" & rst2("id_responsable") & "','" & rst2("id_tipoentidad") & "','" & rst2("piso") & "','" & rst2("pabellon") & "','" & rst2("id_especialidad") & "','" & rst2("observacion") & "','" & rst2("stock") & "', " & _
       "'" & rst2("stock_personalizado") & "','" & rst2("hora_inicio") & "','" & rst2("hora_fin") & "','" & rst2("horas") & "','" & rst2("defecto") & "','" & rst2("activo") & "','" & rst2("id_atension") & "','" & rst2("id_actividad") & "','" & rst2("dni_save") & "','" & rst2("ocupado") & "','" & rst2("id_sucursal") & "','" & rst2("facturacion_detallada") & "','" & rst2("facturacion_centralizada") & "','" & rst2("caja_independiente") & "','" & rst2("cloud") & "','" & rst2("comprobante_adicional") & "')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If


'[2] importar sus Comprobantes.

strCadena = "DELETE FROM almacen_comprobantes WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM almacen_comprobantes WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO almacen_comprobantes(`ruc`,`id_alm`,`id_doc`,`serie`,`numero`,`igv`,`defecto`,`id_moneda`,`venta`,`id_formato_impresion`,`serial`,`afecta_caja`,`tipo_movimiento`,`id_usuario`,`numero_caracteres`,`electronico`,`firmado_online`,`produccion`,`online`)VALUES " & _
       "('" & rst2("ruc") & "','" & rst2("id_alm") & "','" & rst2("id_doc") & "','" & rst2("serie") & "','" & rst2("numero") & "','" & rst2("igv") & "','" & rst2("defecto") & "','" & rst2("id_moneda") & "','" & rst2("venta") & "','" & rst2("id_formato_impresion") & "','" & rst2("serial") & "','" & rst2("afecta_caja") & "','" & rst2("tipo_movimiento") & "', " & _
       "'" & rst2("id_usuario") & "','" & rst2("numero_caracteres") & "','" & rst2("electronico") & "','" & rst2("firmado_online") & "','" & rst2("produccion") & "','" & rst2("online") & "')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If

'[3] importar sus Lineas.
strCadena = "DELETE FROM linea WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM linea WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO linea(`id_linea`,`descripcion`,`consulta_externa`,`procedimientos`,`imagen`,`id_tipo`,`afecto_garantia`,`planilla`,`produccion`,`id_usu`,`garantia`,`mantenimientos`,`nro_cuenta`)VALUES " & _
       "('" & rst2("id_linea") & "','" & rst2("descripcion") & "','" & rst2("consulta_externa") & "','" & rst2("procedimientos") & "','" & rst2("imagen") & "','" & rst2("id_tipo") & "','" & rst2("afecto_garantia") & "','" & rst2("planilla") & "','" & rst2("produccion") & "','" & rst2("garantia") & "','" & rst2("mantenimientos") & "','" & rst2("nro_cuenta") & "')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If
'[4] importar sus Sub Lineas.
strCadena = "DELETE FROM linea_sub WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM linea_sub WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO linea_sub(`id_tipo`,`id_linea`,`descripcion`,`id_usu`)VALUES " & _
       "('" & rst2("id_tipo") & "','" & rst2("id_linea") & "','" & rst2("descripcion") & "','" & rst2("id_usu") & "')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If

'[5] importar sus Plan Contable.
strCadena = "DELETE FROM con_cuentacontable WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM con_cuentacontable WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO con_cuentacontable(`Id`,`IdEmpresaSis`,`IdSucursal`,`IdNaturaleza`,`Ejercicio`,`NroCuenta`,`Descripcion`,`MonedaExtranjera`,`IndCuentaDependiente`,`IdCuentaContableDepende`,`CtaCtbleDepende`,`IndMovimiento`,`DigitoSubfijo`,`CuentaSUNAT`,`IndFlujoCaja`,`IndConciliacion`,`IndDocumento`,`IndObligacion`,`IndDebe`,`IndHaber`,`IndGastoFuncion`,`IndItemGasto`,`IndCentroCosto`,`IndTrabajador`,`IndTracto`,`IndRuta`,`IndBanco`,`Analisis01`,`Analisis02`,`Tesoreria`,`Activo`,`UsuarioCrea`,`FechaCrea`,`UsuarioModifica`,`FechaModifica`)VALUES " & _
       "('" & rst2("id_tipo") & "','" & rst2("id_linea") & "','" & rst2("descripcion") & "','" & rst2("id_usu") & "')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If

'[6] importar Clientes.
strCadena = "SELECT * FROM view_entidad WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "SELECT * FROM persona WHERE dni='" & rst2("dni") & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount < 1 Then
          Call get_cliente(rst2("dni"))
       End If
       rst2.MoveNext
       DoEvents
   Next i
End If


strCadena = "SELECT * FROM persona  ORDER BY dni"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       
       If Val(rst2("dni")) > 10 Then
       If Len(rst2("dni")) = 8 Or Len(rst2("dni")) = 11 Then
nnnn:
       strCadena = "SELECT * FROM persona WHERE   dni='" & Trim(rst2("dni")) & "'  LIMIT 1"
       Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                
                strCadena = "call P_insert_persona_ii('" & rst2("dni") & "' " & _
                ",'" & rst2("a_paterno") & "', " & _
                "'" & rst2("a_materno") & "' " & _
                ",'" & rst2("nombres") & "' " & _
                ",'" & rst2("nombre_completo") & "' " & _
                ",'" & rst2("direccion") & "' " & _
                ",'" & rst2("celular") & "' " & _
                ",'" & rst2("mail") & "'" & _
                ",'no' " & _
                ",'no'" & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                GoTo nnnn
       Else
                strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' and cod_unico='" & Trim(rst2("dni")) & "' LIMIT 1 "
                Call ConfiguraRstlocal(strCadena)
                If rstLocal.RecordCount < 1 Then
                    strCadena = "INSERT INTO entidad_empresa(cod_unico,id_cliente,id_empresa)VALUES ('" & rst2("dni") & "','si','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                End If
       End If
         
       End If
       End If
         rst2.MoveNext
         '  DoEvents
   Next i
End If



proveedorr:

strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='20487725286' and id_proveedor='si' "
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    For i = 0 To rstL.RecordCount - 1
        If Len(rstL("cod_unico")) <> 8 Then
        strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & rstL("cod_unico") & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount < 1 Then
          '  strCadena = "INSERT INTO entidad_empresa(cod_unico,id_cliente,id_proveedor,id_empresa)VALUES ('" & rstL("cod_unico") & "','si','si','" & KEY_RUC & "')"
          '  CnBd.Execute (strCadena)
            
        End If
        End If
        rstL.MoveNext
    Next i
        
        
End If


'[2] importar sus Comprobantes.
'[2] importar sus Comprobantes.
'[2] importar sus Comprobantes.
'[2] importar sus Comprobantes.
'[2] importar sus Comprobantes.
'[2] importar sus Comprobantes.
'[2] importar sus Comprobantes.
'[2] importar sus Comprobantes.
'[2] importar sus Comprobantes.
'[2] importar sus Comprobantes.
'[2] importar sus Comprobantes.


End Sub

Private Function firma_electronica(ByVal in_doc As String, ByVal in_extranjero As String, ByVal in_observacion As String, ByVal in_venta As String, ByVal numero As String, ByVal in_serie As String, ByVal in_dni As String, ByVal in_alumno As String, ByVal in_direccion As String)
Dim in_moneda As String
Call disabled_form(Me)
FrmLoad_web_service.Show
FrmLoad_web_service.nom_prcedimiento = "procesar_firma_electronica_local"
Set FrmLoad_web_service.FormPadre = Me

Select Case in_doc
    Case "0003"
         in_tipo_doc = "1"
         If Trim(in_extranjero) = "si" Then
            in_tipo_doc = "4"
         End If
    Case "0001"
        in_tipo_doc = "6"
         If Trim(in_extranjero) = "si" Then
            in_tipo_doc = "4"
         End If
         
    Case "0002"
        in_tipo_doc = "1"
End Select

    id_motivo_nota = ""
    motivo_nota = ""
    in_serie_afectado = ""
    in_numero_afectado = ""
    in_observacion = Replace(Trim(in_observacion), "'", " ")
    
    id_motivo_nota = ""
    motivo_nota = ""
    in_serie_afectado = ""
    in_numero_afectado = ""
    in_moneda = "PEN"



If get_comprobante_produccion(in_doc, in_serie) = "si" Then
    in_numero = Trim(numero)
    If KEY_SERVIDOR_CLOUD = "si" Then
        
        ' EMPRESAS ANTIGUAS
        'Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_CLOUD & "', x-api-produccion: 'yes'}")
    Else
        'Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_LOCAL & "', x-api-produccion: 'yes'}")
    End If
Else
    in_numero = Trim(numero)
    If KEY_SERVIDOR_CLOUD = "si" Then
       Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, "", in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
    Else
       Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_LOCAL & "'}")
    End If
End If



End Function


Private Sub get_cliente(ByVal in_dni As String)
strCadena = "SELECT * FROM persona where dni='" & in_dni & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
 '   strCadena = "call P_insert_persona_ii('" & Trim(rst("dni")) & "' " & _
                ",'" & Replace(UCase(Me.txtPaterno.Text), "'", " ") & "', " & _
                "'" & Replace(UCase(Me.txtMaterno.Text), "'", " ") & "' " & _
                ",'" & Replace(UCase(Trim(Me.txtNombre.Text)), "'", " ") & "' " & _
                ",'" & Replace(UCase(Trim(Me.txtrazonsocial.Text)), "'", " ") & "' " & _
                ",'" & Trim(Me.TxtDireccion1.Text) & "' " & _
                ",'" & Trim(Me.TxtTelefono.Text) & "' " & _
                ",'" & Me.TxtEmail.Text & "'" & _
                ",'" & StrTransporte & "' " & _
                ",'" & StrContable & "'" & _
                ",'" & strProveedor & "' " & _
                ",'" & StrPersonal & "' " & _
                ",'" & StrAuspiciador & "' " & _
                ",'" & StrAlmacen & "' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
  '              CnBd.Execute (strCadena)
End If
End Sub



Private Sub ChameleonBtn2_Click()


strCadena = "SELECT v.id_venta,v.id_comprobante,v.total,c.documento FROM movimiento_venta v,movimiento_venta c WHERE v.id_comprobante=c.id_venta and  v.id_comprobante>0 and  v.ruc='" & KEY_RUC & "' ORDER BY v.fecha_emision ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
nuevo2:
       strCadena = "SELECT sum(monto_pagado) FROM mis_cuentas_det_detalle WHERE id_movimiento='" & rst("id_comprobante") & "' "
       Call ConfiguraRstL(strCadena)
       If IsNull(rstL(0)) = True Then
          in_pago = 0
        Else
          in_pago = rstL(0)
       End If
       
       If in_pago > 0 Then
            strCadena = "DELETE FROM mis_cuentas_det_detalle WHERE id_movimiento='" & rst("id_comprobante") & "'  LIMIT 1 "
            CnBd.Execute (strCadena)
            GoTo nuevo2
       End If
       
      
            
            
            
            in_saldo = rst("total")
            
            
            
            
  
       If in_saldo > 0 Then
           strCadena = "INSERT INTO mis_cuentas_det_detalle(id_detalle,id_cuenta_det,monto_inicial,monto_pagado,id_movimiento,id_tipo)VALUES " & _
           "('" & rst("id_venta") & "','0','" & in_saldo & "','" & in_saldo & "','" & rst("id_comprobante") & "','01')"
           CnBd.Execute (strCadena)
    End If
siguiente:
       rst.MoveNext
       'DoEvents
   Next i
End If






' IN SP
Exit Sub

strCadena = "SELECT v.total,m.monto_caja,v.id_venta,v.documento FROM movimiento_venta v, movimiento_venta_monto m WHERE  v.id_doc IN ('0001','0003','0007','0054')  and v.id_venta=m.id_venta and m.forma_pago='02'  and v.ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
nuevo:
       strCadena = "SELECT sum(monto_pagado) FROM mis_cuentas_det_detalle WHERE id_movimiento='" & rst("id_venta") & "' "
       Call ConfiguraRstL(strCadena)
       If IsNull(rstL(0)) = True Then
          in_pago = 0
        Else
          in_pago = rstL(0)
       End If
       
       If in_pago > 0 Then
            strCadena = "DELETE FROM mis_cuentas_det_detalle WHERE id_movimiento='" & rst("id_venta") & "'  LIMIT 1 "
            CnBd.Execute (strCadena)
            GoTo nuevo
       End If
       
       If in_pago = rst("total") Then
            
          
            GoTo siguiente2
       Else
            
            
            
            in_saldo = Abs(rst("monto_caja"))
            
            
            
            
            
       End If
       If in_saldo > 0 Then
           strCadena = "INSERT INTO mis_cuentas_det_detalle(id_detalle,id_cuenta_det,monto_inicial,monto_pagado,id_movimiento,id_tipo)VALUES " & _
           "('" & rst("id_venta") & "','0','" & in_saldo & "','" & in_saldo & "','" & rst("id_venta") & "','01')"
           CnBd.Execute (strCadena)
    End If
siguiente2:
       rst.MoveNext
       DoEvents
   Next i
End If

















strCadena = "SELECT v.total,m.monto_caja,v.id_venta,v.documento FROM movimiento_venta v, movimiento_venta_monto m WHERE  v.id_doc IN ('0001','0003','0007','0054')  and v.id_venta=m.id_venta and m.forma_pago='01'  and v.ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
'nuevo:
       strCadena = "SELECT sum(monto_pagado) FROM mis_cuentas_det_detalle WHERE id_movimiento='" & rst("id_venta") & "' "
       Call ConfiguraRstL(strCadena)
       If IsNull(rstL(0)) = True Then
          in_pago = 0
        Else
          in_pago = rstL(0)
       End If
       
       If in_pago > 0 Then
            strCadena = "DELETE FROM mis_cuentas_det_detalle WHERE id_movimiento='" & rst("id_venta") & "' and id_detalle='" & rst("id_venta") & "' LIMIT 1 "
            CnBd.Execute (strCadena)
            GoTo nuevo
       End If
       
       If in_pago = rst("total") Then
            
          
            GoTo siguiente
       Else
            
            
            
            in_saldo = Abs(rst("monto_caja"))
            
            
            
            
            
       End If
       If in_saldo > 0 Then
           strCadena = "INSERT INTO mis_cuentas_det_detalle(id_detalle,id_cuenta_det,monto_inicial,monto_pagado,id_movimiento,id_tipo)VALUES " & _
           "('" & rst("id_venta") & "','0','" & in_saldo & "','" & in_saldo & "','" & rst("id_venta") & "','01')"
           CnBd.Execute (strCadena)
    End If
'siguiente:
       rst.MoveNext
       DoEvents
   Next i
End If




strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='0002' and ruc='" & KEY_RUC & "' and retencion>0   ORDER BY fecha_emision ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
   
       strCadena = "DELETE FROM mis_cuentas_det_detalle WHERE id_movimiento='" & rst("id_compra") & "' and monto_pagado='" & rst("retencion") & "' LIMIT 1 "
       CnBd.Execute (strCadena)
       strCadena = "DELETE FROM movimiento_venta WHERE  fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' and  id_cliente='" & rst("id_proveedor") & "' and  ruc='" & KEY_RUC & "' and id_doc='0097' and total='" & rst("retencion") & "' "
       CnBd.Execute (strCadena)
       
       strCadena = "SELECT * FROM mis_cuentas_det_detalle WHERE id_movimiento='" & rst("id_compra") & "' and monto_pagado='" & rst("retencion") & "' "
       Call ConfiguraRstL(strCadena)
       If rstL.RecordCount < 1 Then
           
                    strCadena = "SELECT * FROM movimiento_venta WHERE  id_doc='0097' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' ORDER BY numero ASC"
                    Call ConfiguraRstPP(strCadena)
                    If rstPP.RecordCount > 0 Then
                       in_numero = Format(Val(rstPP("numero")) + 1, "000000")
                       Documento = "RETENCION" & ":" & rstPP("serie") & "-" & in_numero
                   
                    
                    
                    strCadena = "P_insert_venta('0097','" & rst("id_alm") & "','0','" & rst("id_moneda") & "','no'," & _
                    "'" & rstPP("serie") & "','" & in_numero & "','" & rst("id_proveedor") & "','" & rst("nproveedor") & "','0','0','0','" & rst("retencion") & "','0'," & _
                    "'" & rst("retencion") & "','0','" & Format(rst("fecha_Emision"), "YYYY-mm-dd") & "','" & Format(rst("fecha_Emision"), "YYYY-mm-dd") & "','00001','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & rst("tc") & "','" & dfac & "','" & formato_item(Month(rst("fecha_emision")), 2) & "','" & Year(rst("fecha_emision")) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstP(strCadena)
                    
                    id_venta = rstP(0)
                    
                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & id_venta & "','00','RETENCION','-','1','" & Val(rst("retencion")) & "','0','" & Val(rst("retencion")) & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                               
                    
                    in_numero = Format(Val(in_numero) + 1, "000000")
                    strCadena = "UPDATE almacen_comprobante SET numero='" & in_numero & "' WHERE id_doc='0097' AND serie='" & rstPP("serie") & "' AND ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
           
           
           
           
           
           
           
           
           strCadena = "INSERT INTO mis_cuentas_det_detalle(id_detalle,id_cuenta_det,monto_inicial,monto_pagado,id_movimiento,id_tipo)VALUES " & _
           "('" & id_venta & "','0','" & rst("retencion") & "','" & rst("retencion") & "','" & rst("id_compra") & "','02')"
           CnBd.Execute (strCadena)
       End If
       End If
       rst.MoveNext
      ' DoEvents
   Next i
End If


End Sub

Private Sub ChameleonBtn3_Click()

strCadena = "CALL put_crear_kardex_general_v2('" & KEY_FECHA & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
strCadena = "SELECT * FROM almacen_producto WHERE ruc='" & KEY_RUC & "' order by id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.progress_kardex.Min = 1
   Me.progress_kardex.Max = rst.RecordCount
   For i = 0 To rst.RecordCount - 1
       strCadena = "SELECT saldo_stock FROM tmp_kardex_general  WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC , id_kardex DESC LIMIT 1"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
          strCadena = "UPDATE almacen_producto set stock='" & rstK("saldo_stock") & "' WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
          CnBd.Execute (strCadena)
       End If
       rst.MoveNext
       DoEvents
      ChameleonBtn3.Caption = str(i) & Space(2) & str(rst.RecordCount)
  Next i
 End If
End Sub

Private Sub ChameleonBtn4_Click()

Dim in_fecha As Date
strCadena = "SELECT * FROM con_periodo WHERE id='" & Me.DtpPeriodo.BoundText & "'"
Call ConfiguraRst(strCadena)
in_fecha = Format(rst("FechaFin"), "YYYY-mm-dd")
in_fechaINI = Format(DateAdd("d", 1, rst("FechaFin")), "YYYY-mm-dd")

strCadena = "SELECT * FROM orden_compra WHERE monto_flete>0 and  id_doc='0414' and fecha>='" & in_fecha_ini & "' and fecha<='" & in_fecha_fin & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    rstK.MoveFirst
    For i = 0 To rstK.RecordCount
        ' Call put_flete_orden(rstK("id_recepcion"), rstK("monto_flete"), rstK("valor_venta"))
        ' Call actualizar_kardex(rstK("id_recepcion"))
         
         
         
         
         If i = 0 Then
         If MsgBox("Esta correcto el calculo, para Proceder", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
         End If
         End If
         rstK.MoveNext
    Next i
    
    'aqui tenemos que hacer distinct y recorrer todos los productos
    strCadena = "SELECT * FROM orden "
    
    
    
End If



End Sub

Private Sub validar_servicio()
strCadena = "SELECT * FROM movimiento_venta WHERE id_tipo='02' and  fecha_emision>='" & Format(Me.DtpInicio_migracion.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_migracion.Value, "YYYY-mm-dd") & "' and id_doc IN('0001','0003','0007','0008') and ruc='" & KEY_RUC & "' and anulado='no' ORDER BY fecha_emision ASC,id_doc ASC,numero ASC "
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
   rstIN.MoveFirst
   For i = 0 To rstIN.RecordCount - 1
       strCadena = "SELECT d.id_venta FROM movimiento_venta_detalle d,producto p WHERE p.id_tipo='01' and   d.id_venta='" & rstIN("id_venta") & "' and  d.id_producto=p.id_producto and d.ruc=p.ruc and p.ruc='" & KEY_RUC & "' LIMIT 1"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
                
                strCadena = "UPDATE movimiento_venta SET id_tipo='01' WHERE id_venta='" & rstIN("id_venta") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                CnBd.Execute (strCadena)
                cmdrealizarmigracion.Caption = str(i) & Space(2) & str(rstIN.RecordCount) & Space(2) & rstIN("documento") & Space(2) & rstIN("fecha_emision")
                
                Call delete_asiento(rstIN("id_venta"), Trim(in_doc & rstIN("serie") & rstIN("numero")))
                DoEvents
                
                strCadena = "call P_insert_venta_asiento_contable('" & Val(rstIN("id_venta")) & "')"
                CnBd.Execute (strCadena)
       Else
            DoEvents
            cmdrealizarmigracion.Caption = str(i) & Space(2) & str(rstIN.RecordCount) & Space(2) & rstIN("fecha_emision")
       End If
        DoEvents
       rstIN.MoveNext
   Next i
End If

End Sub


Private Sub ChameleonBtn5_Click()


If Me.chk_servicio.Value = 1 Then
    Call validar_servicio
    Exit Sub
End If


If Me.chk_incompletos.Value = 1 Then
    strCadena = "select v.`id_venta`,v.`documento`,v.`fecha_emision`,v.`id_doc`,v.serie,v.`numero`,a.`Id` as idasiento,v.diferida from movimiento_venta v, " & _
    "con_documento d,con_asiento a Where v.`id_venta`=d.`IdReferencia` and d.`Id`=a.`IdReferencia` and v.fecha_emision>='" & Format(Me.DtpInicio_migracion.Value, "YYYY-mm-dd") & "' and v.fecha_emision<='" & Format(Me.DtpFin_migracion.Value, "YYYY-mm-dd") & "' and " & _
    " v.ruc=d.`IdEmpresaSis` and d.`IdEmpresaSis`=a.`IdEmpresaSis` and v.id_doc IN('0001','0003','0007','0008') and  v.ruc='" & KEY_RUC & "' " & _
    " and a.`IdTipoAsiento`='1CIX000000000137' and d.`Activo`=a.`Activo` and a.`Activo`='1' ORDER BY v.fecha_emision ASC,v.id_doc ASC,v.numero ASC"
Else
    strCadena = "SELECT * FROM movimiento_venta WHERE fecha_emision>='" & Format(Me.DtpInicio_migracion.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_migracion.Value, "YYYY-mm-dd") & "' and id_doc IN('0001','0003','0007','0008') and ruc='" & KEY_RUC & "' and anulado='no' ORDER BY fecha_emision ASC,id_doc ASC,numero ASC"
End If




Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
   rstIN.MoveFirst
   Me.prg_avance.Min = 0
   Me.prg_avance.Max = rstIN.RecordCount
   For i = 0 To rstIN.RecordCount - 1
   
        
        in_doc = ""
        If rstIN("id_doc") = "0001" Then
           in_doc = "FACTURA:"
        End If
         If rstIN("id_doc") = "0003" Then
           in_doc = "BOLETA:"
        End If
         If rstIN("id_doc") = "0007" Then
         in_doc = "NC:"
        End If
        
        If Me.chk_sin_conta.Value = 1 Then
            strCadena = "SELECT Id FROM con_documento WHERE CuentaContable<>'CuentaCont' and idReferencia='" & rstIN("id_venta") & "' and Activo='1' LIMIT 1"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount < 1 Then
            
                cmdrealizarmigracion.Caption = str(i) & Space(2) & str(rstIN.RecordCount) & Space(2) & rstIN("documento") & Space(2) & rstIN("fecha_emision")
                Call delete_asiento(rstIN("id_venta"), Trim(in_doc & rstIN("serie") & rstIN("numero")))
                DoEvents
                strCadena = "call P_insert_venta_asiento_contable('" & Val(rstIN("id_venta")) & "')"
                CnBd.Execute (strCadena)
                
                
                
            Else
                DoEvents
                cmdrealizarmigracion.Caption = str(i) & Space(2) & str(rstIN.RecordCount) & Space(2) & rstIN("fecha_emision")
            End If
            
            
            
            
      End If

      
      If Me.chk_incompletos.Value = 1 Then
            strCadena = "SELECT Id FROM con_asientomovimiento WHERE idEmpresaSis='" & KEY_RUC & "' and  idAsiento='" & rstIN("idasiento") & "' and Activo='1'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount < 4 Then
                
                If rstIN("diferida") = "si" Then
                    If rstT.RecordCount < 2 Then
                        cmdrealizarmigracion.Caption = str(i) & Space(2) & str(rstIN.RecordCount) & Space(2) & rstIN("documento") & Space(2) & rstIN("fecha_emision")
                        Call delete_asiento(rstIN("id_venta"), Trim(in_doc & rstIN("serie") & rstIN("numero")))
                        DoEvents
                        strCadena = "call P_insert_venta_asiento_contable('" & Val(rstIN("id_venta")) & "')"
                        CnBd.Execute (strCadena)
                    End If
                Else
                   cmdrealizarmigracion.Caption = str(i) & Space(2) & str(rstIN.RecordCount) & Space(2) & rstIN("documento") & Space(2) & rstIN("fecha_emision")
                   Call delete_asiento(rstIN("id_venta"), Trim(in_doc & rstIN("serie") & rstIN("numero")))
                   DoEvents
                   strCadena = "call P_insert_venta_asiento_contable('" & Val(rstIN("id_venta")) & "')"
                   CnBd.Execute (strCadena)
                  
                End If
                
                
           Else
                DoEvents
                cmdrealizarmigracion.Caption = str(i) & Space(2) & str(rstIN.RecordCount) & Space(2) & rstIN("fecha_emision")
            End If
      End If
       
            
            DoEvents
            Me.prg_avance.Value = i
            DoEvents
            rstIN.MoveNext
            DoEvents
   Next i
End If
End Sub

Private Sub chk_sucursal_Click()
If Me.chk_sucursal.Value = 1 Then
    Me.DtcAlmacen.Visible = True
    strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcAlmacen)
    
End If
End Sub

Private Sub chk_tiendaOnline_Click()

If Me.chk_tiendaOnline.Value = 1 Then
   Me.chk_pago_efectivo.Visible = True
   Me.chk_pago_visa.Visible = True
   Me.chk_pago_mstercard.Visible = True
   Me.chk_pago_yape.Visible = True
Else
    Me.chk_pago_efectivo.Visible = False
   Me.chk_pago_visa.Visible = False
   Me.chk_pago_mstercard.Visible = False
   Me.chk_pago_yape.Visible = False
End If


End Sub

Private Sub chk_update_kardex_Click()
If Me.chk_update_kardex.Value = 1 Then
   Me.DtpKardex.Value = KEY_FECHA
   Me.DtpKardex.Visible = True
Else
    Me.DtpKardex.Visible = False
End If
End Sub

Private Sub chkActivacion_Click()
If Me.chkActivacion.Value = 1 Then
    Me.Frame1.Visible = True
Else
    Me.Frame1.Visible = False
End If
End Sub

Private Sub chkContador_Click()
If Me.chkContador.Value = 1 Then
    Me.FrameContador.Enabled = True
    Call Resalta(Me.TxtRucContador)
Else
    Me.FrameContador.Enabled = False
End If
End Sub

Private Sub cmdBlanquear_Click()


End Sub


Private Sub cmdcollege_Click()

If MsgBox("Desea migrar los alumnos" + Chr(13) + "RECUERDE QUE ES PARA HSB", vbInformation + vbYesNo) = vbYes Then
   Call migrar_alumnos
End If

End Sub
Private Sub migrar_alumnos()
For i = 0 To 15000
        
        If Val(Me.hfproductos.TextMatrix(i, 0)) >= 0 Then
        
registrar:
        in_grado = Val(Me.hfproductos.TextMatrix(i, 0))
        in_matricula = Format(Trim(Me.hfproductos.TextMatrix(i, 1)), "00000")
        in_pension = Format(Trim(Me.hfproductos.TextMatrix(i, 2)), "00000")
        dni_alumno = Format(Trim(Me.hfproductos.TextMatrix(i, 3)), "00000000")
        in_beca = Me.hfproductos.TextMatrix(i, 17)
        in_media_beca = Me.hfproductos.TextMatrix(i, 18)
        in_estado = "01"
        in_certificado = "no"
        'VERIFICACION DE LA EXISTENCIA DEL ALUMNO
        strCadena = "SELECT cod_unico FROM entidad_empresa WHERE cod_unico='" & dni_alumno & "' and id_empresa='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            
             'in_detalle = rstL("nombre_prod") & Space(1) & "PERIODO [2017]"
             'MATRICULA
             in_unico_pago = "si"
             in_pago_anual = "si"
             in_pago_mensual = "no"
             
            
             strCadena = "DELETE FROM college_matricula WHERE dni='" & dni_alumno & "' and id_periodo='2' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
             
             
            If Trim(in_matricula) <> "" Then
                 strCadena = "DELETE FROM persona_plan_servicio WHERE dni='" & dni_alumno & "' and pago_anual='si' and ruc='" & KEY_RUC & "'"
             CnBd.Execute (strCadena)
                
                strCadena = "call put_plan_servicio_ii('" & in_matricula & "','" & dni_alumno & "','" & in_unico_pago & "','" & in_pago_mensual & "','" & KEY_USUARIO & "','" & KEY_ALM & "','" & Format("01-01-2018", "YYYY-mm-dd") & "','" & in_certificado & "','" & Format("01-01-2018", "YYYY-mm-dd") & "','" & Format("30-12-2018", "YYYY-mm-dd") & "','" & in_pago_anual & "','0','" & KEY_MORA_MONTO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
            
            'PENSION
             in_unico_pago = "no"
             in_pago_anual = "no"
             in_pago_mensual = "si"
            If Trim(in_pension) <> "" Then
                 strCadena = "DELETE FROM persona_plan_servicio WHERE dni='" & dni_alumno & "' and pago_mensual='si' and ruc='" & KEY_RUC & "'"
                 CnBd.Execute (strCadena)
                strCadena = "call put_plan_servicio_ii('" & in_pension & "','" & dni_alumno & "','" & in_unico_pago & "','" & in_pago_mensual & "','" & KEY_USUARIO & "','" & KEY_ALM & "','" & Format("01-01-2018", "YYYY-mm-dd") & "','" & in_certificado & "','" & Format("01-01-2018", "YYYY-mm-dd") & "','" & Format("30-12-2018", "YYYY-mm-dd") & "','" & in_pago_anual & "','0','" & KEY_MORA_MONTO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
            
            
            strCadena = "SELECT * FROM nivel_educativo_grado WHERE id_grado='" & in_grado & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRstL(strCadena)
            If rstL.RecordCount > 0 Then
                strCadena = "call put_matricula_college('0','" & dni_alumno & "','" & in_grado & "','" & in_beca & "','" & in_media_beca & "','" & in_matricula & "','" & in_pension & "','" & in_estado & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
             End If
                
        
        Else
           ' Call put_estudiante(Trim(dni_alumno))
            MsgBox "REGISTRE A ESTE ALUMNO" + Chr(13) + dni_alumno, vbInformation
            'GoTo registrar
        End If
        
        
        End If
        DoEvents
Next i



End Sub



Private Sub cmdcomprascontables_Click()

strCadena = "SELECT * FROM movimiento_venta WHERE  id_doc='0007' and id_tipo_nota IN('04','05','08','09') and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_detalle_serie>0 and  id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
           rstK.MoveFirst
           
           For j = 0 To rstK.RecordCount - 1
                MsgBox rstK("id_producto") & Space(2) & rstK("detalle") & Space(2) & rstK("nro_chasis") + Chr(13) + rst("id_cliente") & Space(2) & rst("ncliente") + Chr(13) + rst("documento")
                X = 0
                
                rstK.MoveNext
                DoEvents
           Next j
        End If
        rst.MoveNext
   Next i
End If






Exit Sub
'
strCadena = "SELECT * FROM movimiento_compra c WHERE  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "UPDATE movimiento_compra SET id_alm='00001' WHERE id_compra='" & rst("id_compra") & "'"
        CnBd.Execute (strCadena)
        rst.MoveNext
   Next i
End If



strCadena = "SELECT * FROM movimiento_compra c WHERE id_doc<>'0417' and id_doc<>'0418' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       
       
         If rst("id_moneda") = "00001" Then
                in_cta_compra = KEY_CTA_COMPRA_SOLES
         Else
                in_cta_compra = KEY_CTA_COMPRA_DOLARES
         End If
        
        If rst("id_doc") = "0002" Then
             in_cta_compra = KEY_CTA_COMPRA_RH
        End If
        
       If rst("id_doc") <> "0089" Then
       strCadena = "UPDATE movimiento_compra SET cta_pagar='" & in_cta_compra & "' WHERE id_compra='" & rst("id_compra") & "'"
       CnBd.Execute (strCadena)
       strCadena = "call p_insert_compra_emitido_xd('" & rst("id_compra") & "')"
       CnBd.Execute (strCadena)
       End If
       
       rst.MoveNext
   Next i
End If
End Sub

Private Sub cmdCorregir_forma_pago_Click()
End Sub
Private Function get_registro(ByVal in_forma_pago As String) As Integer
strCadena = "SELECT id_registro FROM forma_pago_detalle WHERE id_detalle='" & in_forma_pago & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    get_registro = rstK("id_registro")
Else
    get_registro = 0
End If
End Function


Private Sub ventas(ByVal in_fecha As Date)
strCadena = "SELECT id_venta,id_tipo,fecha_emision,id_doc,serie,numero,id_cliente,id_producto,cantidad,id_alm,dni_save,ruc,id_orden_compra,id_recepcion FROM vargas_kardex_ventas where fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'  ORDER BY fecha_emision ASC,id_doc ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       in_tipo = "01"
       strCadena = "call put_kardex_stock_vitekey('" & in_tipo & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & rst("id_venta") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_cliente") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','0','" & rst("id_alm") & "','" & rst("dni_save") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If
End Sub
Private Sub transferencia_ingreso(ByVal in_fecha As Date)
strCadena = "SELECT * FROM view_vargas_transferencia_ingreso where fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_doc ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       in_tipo = "03"
       
       strCadena = "SELECT * FROM almacen_producto WHERE ruc='" & KEY_RUC & "' and id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "'"
       Call ConfiguraRstL(strCadena)
       If rstL.RecordCount < 1 Then
             strCadena = "INSERT INTO almacen_producto(id_alm,precio_venta,precio_compra,id_producto,ruc) VALUES ('" & rst("id_alm") & "','" & get_precio_producto(rst("id_producto"), "00001") & "','" & get_costo_producto(rst("id_producto")) & "','" & rst("id_producto") & "','" & Trim(KEY_RUC) & "')"
             CnBd.Execute (strCadena)
       End If
       
       
       strCadena = "call put_kardex_stock_vitekey('03','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & rst("id_compra") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_remitente") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','0','" & rst("id_alm") & "','" & rst("dni_save") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       
       rst.MoveNext
   Next i
End If
End Sub

Private Sub transferencia_salida(ByVal in_fecha As Date)
strCadena = "SELECT * FROM view_vargas_transferencia_salida where fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_doc ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       in_tipo = "03"
       
       strCadena = "SELECT * FROM almacen_producto WHERE ruc='" & KEY_RUC & "' and id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "'"
       Call ConfiguraRstL(strCadena)
       If rstL.RecordCount < 1 Then
             strCadena = "INSERT INTO almacen_producto(id_alm,precio_venta,precio_compra,id_producto,ruc) VALUES ('" & rst("id_alm") & "','" & get_precio_producto(rst("id_producto"), "00001") & "','" & get_costo_producto(rst("id_producto")) & "','" & rst("id_producto") & "','" & Trim(KEY_RUC) & "')"
             CnBd.Execute (strCadena)
       End If
       
       
       strCadena = "call put_kardex_stock_vitekey('03','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & rst("id_transferencia") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_destinatario") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','0','" & rst("id_alm") & "','" & rst("dni_save") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       
       rst.MoveNext
   Next i
End If
End Sub


Private Sub compras(ByVal in_fecha As Date, ByVal in_producto As String)
'strCadena = "SELECT id_compra,id_tipo,fecha_emision,id_doc,serie,numero,id_proveedor,id_producto,cantidad,c_unitario,id_alm,dni_save,ruc,id_orden_compra,id_recepcion,id_moneda,tc FROM vargas_kardex_compra where fecha_emision='" & Format(In_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_doc ASC"
strCadena = "SELECT c.id_compra,fecha_emision,id_doc,serie,numero,id_proveedor,id_producto,cantidad,c_unitario,c.id_alm,dni_save,c.ruc,id_orden_compra,id_recepcion,id_moneda,tc FROM movimiento_compra_detalle d,movimiento_compra c where d.id_compra=c.id_compra and d.ruc=c.ruc and  d.id_producto='" & in_producto & "' and  fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and c.ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_doc ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       in_tipo = "02"
       If rst("id_moneda") = "00002" Then
          in_unitario = rst("c_unitario") * rst("tc")
       Else
          in_unitario = rst("c_unitario")
       End If
       strCadena = "call put_kardex_stock_vitekey('" & in_tipo & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & rst("id_compra") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_proveedor") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','" & in_unitario & "','" & rst("id_alm") & "','" & rst("dni_save") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       
       rst.MoveNext
   Next i
End If
End Sub


Private Sub notas(ByVal in_fecha As Date)
strCadena = "SELECT id_venta,id_tipo,fecha_emision,id_doc,serie,numero,id_cliente,id_producto,cantidad,id_alm,dni_save,ruc,id_orden_compra,id_recepcion FROM vargas_kardex_notas where fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_doc ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
      in_tipo = "07"
      strCadena = "call put_kardex_stock_vitekey('" & in_tipo & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & rst("id_venta") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_cliente") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','0','" & rst("id_alm") & "','" & rst("dni_save") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       
       rst.MoveNext
   Next i
End If
End Sub


       
Private Function actualizar_kardex_orden_compra(ByVal in_recepcion As String, ByVal in_estado As String, ByVal in_orden As String) As Integer
    Dim in_costo_igv As Single
    Dim in_afecto_igv As String
    Dim in_moneda As String
    Dim in_factor As Single
   strCadena = "SELECT * FROM orden_compra WHERE id_recepcion='" & Val(in_orden) & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRstL(strCadena)
   If rstL.RecordCount > 0 Then
        'ACTUALIZAR KARDEX CON GUIA
        in_moneda = rstL("id_moneda")
        in_afecto_igv = rstL("afecto_igv")
        
        If rstL.RecordCount = 1 And in_estado = 2 Then
            ' aqui ya se agrego con la factura
            actualizar_kardex_orden_compra = 1
        Else
            strCadena = "SELECT * FROM orden_compra_detalle WHERE id_orden='" & Val(in_recepcion) & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRstL(strCadena)
            If rstL.RecordCount > 0 Then
               rstL.MoveFirst
               If in_moneda = "00001" Then
                   in_factor = 1
               Else
                   in_factor = KEY_CAMBIO
               End If
               For i = 0 To rstL.RecordCount - 1
                    'Call insertar_kardex_producto(in_recepcion, "0009", Me.txtGuia_serie.Text, Me.TxtGuia_numero.Text, Trim(Me.TxtRuc.Text), rstL("id_producto"), Me.DtcAlmacen.BoundText, Me.DtpPedido.Value, rstL("cantidad"), rstL("precio") + rstL("incremento_neto"))
                    in_monto_neto = (rstL("precio") * in_factor + rstL("incremento_neto"))
                    If in_afecto_igv = "si" Then
                        in_costo_igv = rstL("precio") * in_factor + rstL("precio") * KEY_IGV * in_factor + rstL("incremento_neto")
                    Else
                        in_costo_igv = in_monto_neto
                    End If
                    '****** MODIFICAR EL PERIODO
                    strCadena = "call put_kardex_stock('04','" & in_periodo & "','" & Val(in_recepcion) & "','0009','" & Trim(SERIE_GUIA) & "','" & Trim(numero_guia) & "','" & Trim(in_ruc_proveedor) & "','" & rstL("id_producto") & "','" & rstL("cantidad") & "','" & in_costo_igv & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    rstL.MoveNext
               Next i
            End If
        End If
    End If
End Function

Private Sub cmdcorregirComprobantes_Click()

'excel migrar proveedores


strCadena = "SELECT * FROM producto_sabar order by id_producto DESC "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM persona WHERE dni='" & rst("id_proveedor") & "'"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount < 1 Then
            strCadena = "call P_insert_persona_ii('" & rst("id_proveedor") & "' " & _
                ",'-', " & _
                "'-' " & _
                ",'-' " & _
                ",'" & rst("proveedor") & "' " & _
                ",'-' " & _
                ",'-' " & _
                ",'-'" & _
                ",'no' " & _
                ",'no'" & _
                ",'si' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
        End If
        
        strCadena = "UPDATE producto SET id_proveedor='" & rst("id_proveedor") & "' WHERE id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        rst.MoveNext
        DoEvents
        
   Next i
End If


End Sub

Private Sub cmdEliminar_Click()
strCadena = "DELETE  FROM  grupo_empresarial WHERE ruc_vinculado='" & Trim(Me.hfgrupoempresarial.TextMatrix(Me.hfgrupoempresarial.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Call llenar_empresas(hfgrupoempresarial)


End Sub



Private Sub cmdfaltantes_Click()

End Sub

Private Sub cmdgenerarCierre_Click()

Dim in_fecha As Date
strCadena = "SELECT * FROM con_periodo WHERE id='" & Me.DtpPeriodo.BoundText & "'"
Call ConfiguraRst(strCadena)
in_fecha = Format(rst("FechaFin"), "YYYY-mm-dd")
in_fechaINI = Format(DateAdd("d", 1, rst("FechaFin")), "YYYY-mm-dd")
in_cta_compra = KEY_CTA_COMPRA_SOLES

strCadena = "SELECT * FROM almacen_producto WHERE ruc='" & KEY_RUC & "' order by id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.progress_kardex.Min = 1
   Me.progress_kardex.Max = rst.RecordCount
   For i = 0 To rst.RecordCount - 1
       
       If i < Val(Me.txtCorrelativa.Text) Then
            GoTo nn
       End If
       
       
       
       strCadena = "SELECT id_kardex FROM kardex  WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount < 1 Then
            GoTo nn
       End If
       
       strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0424' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "'"
       Call ConfiguraRstIN(strCadena)
       in_serie = rstIN("serie")
       in_numero = rstIN("numero")
       
       strCadena = "SELECT precio,costo_promedio,costo_unitario,saldo_stock FROM kardex  WHERE fecha_emision<='" & Format(in_fecha, "YYYY-mm-dd") & "' and  id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC, id_kardex DESC LIMIT 1"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
          ':::::Generar Salida Cierre
            in_cantidad = Val(rstK("saldo_stock"))
            in_precio_venta = get_precio_producto(rst("id_producto"), rst("id_alm"))
            n_producto = get_producto(rst("id_producto"), KEY_RUC)
            
            strCadena = "call P_insert_compra_ultimate('0424','" & rst("id_alm") & "','" & Format(in_fecha, "YYYY-mm-dd") & "','" & Format(in_fecha, "YYYY-mm-dd") & "','02'," & _
            "'03','--','00001','" & formato_item(Month(in_fecha), 2) & "','" & Year(in_fecha) & "','" & in_serie & "'," & _
            "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "',3.37," & _
            "'0','0','0','0','0','0','0','0','0','0','0'," & _
            " '" & KEY_USUARIO & "','OBSERVACION','01','" & Me.DtpPeriodo.BoundText & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
           Call ConfiguraRstP(strCadena)
           id_compra = rstP(0)
           
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(rst("id_producto")) & "','" & rstK("saldo_stock") & "','" & rstK("costo_promedio") & "'," & _
           "'0','0','0','" & rstK("saldo_stock") * rstK("costo_promedio") & "','0','0', " & _
           "'0','0','0','" & rstK("saldo_stock") * rstK("costo_promedio") & "','0','" & rstK("costo_promedio") * rstK("saldo_stock") & "','" & in_precio_venta & "','" & rstK("costo_promedio") & "','" & rst("id_alm") & "','" & n_producto & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           strCadena = "call put_kardex_stock_inventario('20','" & Format(in_fecha, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0424','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(rst("id_producto")) & "','" & rstK("saldo_stock") & "','" & rstK("costo_promedio") & "','" & rst("id_alm") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(in_numero) + 1, "000000") & "' WHERE id_doc='0424' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
           CnBd.Execute (strCadena)
            
           strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0106' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "'"
           Call ConfiguraRstIN(strCadena)
           in_serie = rstIN("serie")
           in_numero = rstIN("numero")
           in_fechaFIN = Format(DateAdd("d", 1, in_fecha), "YYYY-mm-dd")
            strCadena = "call P_insert_compra_ultimate('0106','" & rst("id_alm") & "','" & Format(in_fechaFIN, "YYYY-mm-dd") & "','" & Format(in_fechaFIN, "YYYY-mm-dd") & "','02'," & _
            "'03','--','00001','" & formato_item(Month(in_fechaFIN), 2) & "','" & Year(in_fechaFIN) & "','" & in_serie & "'," & _
            "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "',3.37," & _
            "'0','0','0','0','0','0','0','0','0','0','0'," & _
            " '" & KEY_USUARIO & "','OBSERVACION','01','" & Me.DtpPeriodo.BoundText & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
           Call ConfiguraRstP(strCadena)
           id_compra = rstP(0)
           
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(rst("id_producto")) & "','" & rstK("saldo_stock") & "','" & rstK("costo_promedio") & "'," & _
           "'0','0','0','" & rstK("saldo_stock") * rstK("costo_promedio") & "','0','0', " & _
           "'0','0','0','" & rstK("saldo_stock") * rstK("costo_promedio") & "','0','" & rstK("costo_promedio") * rstK("saldo_stock") & "','" & in_precio_venta & "','" & rstK("costo_promedio") & "','" & rst("id_alm") & "','" & n_producto & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           strCadena = "call put_kardex_stock_inventario('02','" & Format(in_fechaFIN, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0106','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(rst("id_producto")) & "','" & rstK("saldo_stock") & "','" & rstK("costo_promedio") & "','" & rst("id_alm") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(in_numero) + 1, "000000") & "' WHERE id_doc='0106' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
           CnBd.Execute (strCadena)
           
           
       End If
nn:
       rst.MoveNext
       DoEvents
       Me.progress_kardex.Value = i + 1
       Me.cmdgenerarCierre.Caption = str(i) & Space(5) & str(rst.RecordCount)
   Next i
End If






End Sub

Private Sub cmdgetkeyfacil_Click()
'Call disabled_form(Me)
'FrmLoad_web_service.Show
'FrmLoad_web_service.nom_prcedimiento = "get_producto_keyfacil"
'Set FrmLoad_web_service.FormPadre = Me
'Call FrmLoad_web_service.get_producto_keyfacil("https://api.vitekey.com/keyfact/erp/products/list?api_key=fd235235-e97a-4db6-8f50-fa84145c3f5d", "POST", json_crear_get_producto(), "{x-api-token: '" & KEY_TOKEN_CLOUD & "', x-api-produccion: 'yes'}")
 Call get_importar_productos_keyfacil



End Sub



Public Sub get_producto_keyfacil(ByVal strHtml As String)

Dim in_error As Boolean
Dim in_hash As String
Dim in_total_comprobante As Double
Dim in_emision As Date
Dim in_fecha As String
Dim in_alm As String
Dim in_fecha_actual As String
Dim in_fecha_comprobante As Date
Dim in_productos() As String
Dim json_r As Object

Set json_r = JSON.parse(strHtml)









If json_r("response").Count >= 1 Then
   
For i = 1 To json_r("response").Count  ' recorro la cantidad de comprobantes
    in_producto = Format(json_r("response")(i).Item("code"), "00000")
    
    If json_r("response")(1).Item("photos_urls").Count > 0 Then
        in_foto = ""
    End If
    
    
    
            
                
     
        DoEvents
Next i
End If



End Sub


Private Sub cmdHBS_Click()
Dim sys_ConString2 As String
Dim stock_actual As Integer

sys_Server2 = Trim(Me.txtserver1.Text)
sys_DataBase2 = Trim(Me.txtbaseOrigen1.Text)   'ConfigRead("DataBase")
sys_SUser2 = "user_cord" 'DecryptString(ConfigRead("SUser"))
sys_SPassword2 = "123456" 'DecryptString(ConfigRead("SPassword"))
db_port = "3306"
sys_ConString2 = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server2 & ";" & _
            "Database=" & sys_DataBase2 & ";" & _
            "UID=" & sys_SUser2 & ";" & _
            "PWD=" & sys_SPassword2 & ";" & _
            " PORT=" & db_port & ";OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
CnBd2.ConnectionString = sys_ConString2
CnBd2.Open

strCadena = "SELECT * FROM view_entidad WHERE id_cliente='si' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       
       strCadena = "SELECT * FROM view_entidad WHERE id_cliente='si' and  dni='" & Trim(rst2("dni")) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
       Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                
                strCadena = "call P_insert_persona_ii('" & rst2("cod_unico") & "' " & _
                ",'" & rstCloud("a_paterno") & "', " & _
                "'" & rstCloud("a_materno") & "' " & _
                ",'" & rstCloud("nombres") & "' " & _
                ",'" & rstCloud("nombre_completo") & "' " & _
                ",'" & rstCloud("direccion") & "' " & _
                ",'" & rstCloud("celular") & "' " & _
                ",'" & rstCloud("mail") & "'" & _
                ",'no' " & _
                ",'no'" & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                Call familiares(rst2("cod_unico"))
       End If
       
       strCadena = "UPDATE persona SET id_pais='" & rst2("id_pais") & "',peso='" & Val(rst2("peso")) & "',estatura='" & Trim(rst2("estatura")) & "',id_mes='" & rst2("id_mes") & "',id_anio='" & rst2("id_anio") & "'," & _
       "id_departamento='" & rst2("id_departamento") & "',id_provincia='" & rst2("id_provincia") & "',id_distrito='" & rst2("id_distrito") & "' WHERE dni='" & rst2("dni") & "'"
       CnBd.Execute (strCadena)
       ' EXPED
       Call put_estudiante(Trim(rst2("dni")))
       'Call put_matricula(Me.DtcPeriodo.BoundText, Me.DtcNivel.BoundText, Me.DtcGrado.BoundText, Trim(Me.TxtRuc.Text), Me.DtcServicio.BoundText)
                
       '--
                    
                    
   rst2.MoveNext
   DoEvents
   Next i
   
   strCadena = "SELECT * FROM movimiento_venta WHERE ruc='" & KEY_RUC & "' ORDER BY id_venta DESC"
   Call ConfiguraRstCloud(strCadena)
   If rstCloud.RecordCount > 1 Then
      rstCloud.MoveFirst
      For i = 0 To rstCloud.RecordCount - 1
          strCadena = "SELECT * FROM movimiento_venta WHERE serie='" & rstCloud("serie") & "' and numero='" & rstCloud("numero") & "' and id_doc='" & rstCloud("id_doc") & "' and ruc='" & KEY_RUC & "'"
          Call ConfiguraRst(strCadena)
          If rst.RecordCount < 1 Then
            Call put_venta(rstCloud("id_venta"))
          End If
          rstCloud.MoveNext
      Next i
   End If
   
   
End If
End Sub
Public Sub update_cuenta_contable()
Dim in_cuenta As String
strCadena = "SELECT * FROM con_cuentacontable WHERE IdEmpresaSis='" & KEY_RUC & "' order by id ASC" ' ginsac
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        in_cuenta = ""
        For j = 1 To Len(rst("NroCuenta"))
            If j = 1 Then
                in_cuenta = Mid(rst("NroCuenta"), j, 1)
            Else
                If Mid(rst("NroCuenta"), j, 1) = " " Then
                    Exit For
                End If
                in_cuenta = in_cuenta & Mid(rst("NroCuenta"), j, 1)
            End If
        Next j
   
   
       strCadena = "UPDATE con_cuentacontable SET NroCuenta ='" & Trim(in_cuenta) & "' WHERE id='" & rst("id") & "' and IdEmpresaSis='" & KEY_RUC & "' LIMIT 1 "
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If
End Sub


Public Sub migrar_empresa_local()
Dim sys_ConString2 As String
Dim stock_actual As Integer
'Call update_cuenta_contable



'Exit Sub
sys_Server2 = "localhost"
sys_DataBase2 = "bd_vitekey_repos_ii"
sys_SUser2 = "root" 'DecryptString(ConfigRead("SUser"))
sys_SPassword2 = "@02021974abc2016@123@cord" 'DecryptString(ConfigRead("SPassword"))
sys_SPassword2 = "02021974abc2016" 'DecryptString(ConfigRead("SPassword"))
'sys_SPassword2 = "vitekey2018"
db_port = "3306"
sys_ConString2 = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server2 & ";" & _
            "Database=" & sys_DataBase2 & ";" & _
            "UID=" & sys_SUser2 & ";" & _
            "PWD=" & sys_SPassword2 & ";" & _
            " PORT=" & db_port & ";OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
CnBd2.ConnectionString = sys_ConString2
CnBd2.Open

GoTo migra_venta
'[ELIMINAR VENTAS]
strCadena = "DELETE FROM movimiento_venta where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE FROM movimiento_venta_detalle where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE from movimiento_venta_monto where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE from movimiento_venta_cuotas where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


'[ELIMINAR COMPRAS]
strCadena = "DELETE FROM movimiento_compra where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
strCadena = "SELECT * FROM movimiento_compra_detalle where ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "DELETE from movimiento_compra_detalle where  ruc='" & KEY_RUC & "' LIMIT 100"
        CnBd.Execute (strCadena)
        rst.MoveNext
        DoEvents
   Next i
End If


strCadena = "DELETE from imp_producto_detalle where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


'[ELIMINAR GUIAS DE REMISION]
strCadena = "DELETE FROM movimiento_transferencia where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE from movimiento_transferencia_detalle where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


'[ELIMINAR PRODUCTOS]
strCadena = "DELETE FROM producto WHERE    ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
strCadena = "DELETE FROM almacen_producto WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
strCadena = "DELETE FROM almacen_producto_precio WHERE  ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


'[ELIMINAR LINEAS]
strCadena = "DELETE FROM linea where id_usu='" & KEY_RUC & "' "
CnBd.Execute (strCadena)


strCadena = "DELETE FROM linea_sub WHERE id_usu='" & KEY_RUC & "' "
CnBd.Execute (strCadena)


'[ ELIMINAR ALMACENES ].
strCadena = "DELETE FROM almacen WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM almacen WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO almacen(`id_alm`,`ruc`,`descripcion`,`abreviatura`,`direccion`,`ubicacion_interna`,`id_responsable`,`id_tipoentidad`,`piso`,`pabellon`,`id_especialidad`,`observacion`,`stock`,`stock_personalizado`,`hora_inicio`,`hora_fin`,`horas`,`defecto`,`activo`,`id_atension`,`id_actividad`,`dni_save`,`ocupado`,`id_sucursal`,`facturacion_detallada`,`facturacion_centralizada`,`caja_independiente`,`cloud`,`comprobante_adicional`)VALUES " & _
       "('" & rst2("id_alm") & "','" & rst2("ruc") & "','" & rst2("descripcion") & "','" & rst2("abreviatura") & "','" & rst2("direccion") & "','" & rst2("ubicacion_interna") & "','" & rst2("id_responsable") & "','" & rst2("id_tipoentidad") & "','" & rst2("piso") & "','" & rst2("pabellon") & "','" & rst2("id_especialidad") & "','" & rst2("observacion") & "','" & rst2("stock") & "', " & _
       "'" & rst2("stock_personalizado") & "','" & Format(rst2("hora_inicio"), "HH:mm:ss") & "','" & Format(rst2("hora_fin"), "HH:mm:ss") & "','" & Format(rst2("horas"), "HH:mm:ss") & "','" & rst2("defecto") & "','" & rst2("activo") & "','" & rst2("id_atension") & "','" & rst2("id_actividad") & "','" & rst2("dni_save") & "','" & rst2("ocupado") & "','" & rst2("id_sucursal") & "','" & rst2("facturacion_detallada") & "','no','" & rst2("caja_independiente") & "','" & rst2("cloud") & "','no')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If


'[2] importar sus Comprobantes.

strCadena = "DELETE FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO almacen_comprobante(`ruc`,`id_alm`,`id_doc`,`serie`,`numero`,`igv`,`defecto`,`id_moneda`,`venta`,`id_formato_impresion`,`serial`,`afecta_caja`,`tipo_movimiento`,`id_usuario`,`numero_caracteres`,`electronico`,`firmado_online`,`produccion`,`online`)VALUES " & _
       "('" & rst2("ruc") & "','" & rst2("id_alm") & "','" & rst2("id_doc") & "','" & rst2("serie") & "','" & rst2("numero") & "','" & rst2("igv") & "','" & rst2("defecto") & "','" & rst2("id_moneda") & "','" & rst2("venta") & "','" & rst2("id_formato_impresion") & "','" & rst2("serial") & "','" & rst2("afecta_caja") & "','" & rst2("tipo_movimiento") & "', " & _
       "'" & rst2("id_usuario") & "','" & rst2("numero_caracteres") & "','" & rst2("electronico") & "','" & rst2("firmado_online") & "','" & rst2("produccion") & "','" & rst2("online") & "')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If

'[3] importar sus Marcas.
strCadena = "DELETE FROM marca WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM marca WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
      strCadena = "INSERT INTO marca(id_marca,descripcion,id_usu)VALUES('" & rst2("id_marca") & "','" & rst2("descripcion") & "','" & KEY_RUC & "')"
      CnBd.Execute (strCadena)
      rst2.MoveNext
   Next i
End If


'[3] importar sus Lineas.
strCadena = "DELETE FROM linea WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM linea WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO linea(`id_linea`,`descripcion`,`id_tipo`,`afecto_garantia`,`planilla`,`produccion`,`id_usu`,`garantia`,`mantenimientos`,`nro_cuenta`)VALUES " & _
       "('" & rst2("id_linea") & "','" & rst2("descripcion") & "','" & rst2("id_tipo") & "','" & rst2("afecto_garantia") & "','" & rst2("planilla") & "','" & rst2("produccion") & "','" & KEY_RUC & "','" & rst2("garantia") & "','" & rst2("mantenimientos") & "','60111')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If

strCadena = "DELETE FROM unidad WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM unidad WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO unidad(id_und,abreviatura,descripcion,id_usu)VALUES " & _
       "('" & rst2("id_und") & "','" & rst2("abreviatura") & "','" & rst2("descripcion") & "','" & rst2("id_usu") & "')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If
'[4] importar sus Sub Lineas.
strCadena = "DELETE FROM linea_sub WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM linea_sub WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO linea_sub(`id_tipo`,`id_linea`,`descripcion`,`id_usu`)VALUES " & _
       "('" & rst2("id_tipo") & "','" & rst2("id_linea") & "','" & rst2("descripcion") & "','" & rst2("id_usu") & "')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If


Producto:

'[IMPORTAR PRODUCTOS]
        strCadena = "SELECT * FROM producto WHERE ruc='" & KEY_RUC & "' order by id_producto ASC"
        Call ConfiguraRst2(strCadena)
        If rst2.RecordCount > 0 Then
           rst2.MoveFirst
           For i = 0 To rst2.RecordCount - 1
               strCadena = "SELECT * FROM producto WHERE id_producto='" & rst2("id_producto") & "' and ruc='" & KEY_RUC & "'"
               Call ConfiguraRstA(strCadena)
               If rstA.RecordCount < 1 Then
                    strCadena = "INSERT INTO producto (`id_producto`,id_tipo,`id_linea`,`id_sublinea`,`id_moneda`,`id_color`,`nombre_prod`,`id_unidad`,`nombre_comercial`,`id_marca`,`id_igv`,`dni_save`,`ruc`) VALUES " & _
                    "('" & rst2("id_producto") & "','01','" & rst2("id_linea") & "','" & rst2("id_sublinea") & "','00001','" & rst2("id_color") & "','" & rst2("nombre_prod") & "','" & rst2("id_unidad") & "','" & rst2("nombre_comercial") & "','" & rst2("id_marca") & "','si','" & rst2("dni_save") & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                Else
                    If rst2("nombre_prod") <> rstA("nombre_prod") Then
                        strCadena = "UPDATE producto SET nombre_prod='" & rst2("nombre_prod") & "' WHERE id_producto='" & rst2("id_producto") & "' and ruc='" & KEY_RUC & "'"
                        CnBd.Execute (strCadena)
                    End If
                End If
                
                
                
               rst2.MoveNext
               
           Next i
        End If
        
        
stock:
        strCadena = "DELETE FROM almacen_producto WHERE ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)

        strCadena = "SELECT * FROM almacen_producto WHERE id_alm='00001' and   ruc='" & KEY_RUC & "'"
        Call ConfiguraRst2(strCadena)
        If rst2.RecordCount > 0 Then
           rst2.MoveFirst
           For j = 0 To rst2.RecordCount - 1
               strCadena = "INSERT INTO almacen_producto(`id_alm`,`id_producto`,precio_venta,precio_compra,`ruc`,`habilitado`,stock,stock_contable) VALUES ('" & rst2("id_alm") & "','" & rst2("id_producto") & "','" & rst2("precio_venta") & "','" & rst2("precio_compra") & "','" & KEY_RUC & "','si','" & rst2("stock") & "','" & rst2("stock_factura") & "')"
               CnBd.Execute (strCadena)
               Call put_kardex_inventario(rst2("id_producto"), rst2("id_alm"), rst2("stock"), rst2("precio_compra"), rst2("precio_venta"), "1CIX000000000038")
               rst2.MoveNext
           Next j
        End If



'[5] importar sus Plan Contable.

Dim in_inicial As String
strCadena = "DELETE FROM con_cuentacontable WHERE IdEmpresaSis='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM con_cuentacontable WHERE IdEmpresaSis='20487725286' order by id ASC" ' ginsac
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   in_inicial = Mid(KEY_RUC, 9, 2) & Format(Trim(Val(KEY_ALM)), "00")
   For i = 0 To rstA.RecordCount - 1
       in_id = Trim(in_inicial + Mid(rstA("id"), 5, 16))
       strCadena = "INSERT INTO con_cuentacontable(`Id`,`IdEmpresaSis`,`IdSucursal`,`IdNaturaleza`,`Ejercicio`,`NroCuenta`,`Descripcion`,`MonedaExtranjera`,`IndCuentaDependiente`,`IdCuentaContableDepende`,`CtaCtbleDepende`," & _
       "`IndMovimiento`,`DigitoSubfijo`,`CuentaSUNAT`,`IndFlujoCaja`,`IndConciliacion`,`IndDocumento`,`IndObligacion`,`IndDebe`,`IndHaber`,`IndGastoFuncion`,`IndItemGasto`,`IndCentroCosto`,`IndTrabajador`,`IndTracto`," & _
       "`IndRuta`,`IndBanco`,`Analisis01`,`Analisis02`,`Tesoreria`,`Activo`,`UsuarioCrea`,`FechaCrea`,`UsuarioModifica`,`FechaModifica`)VALUES " & _
       "('" & in_id & "','" & KEY_RUC & "','1CIX124','" & rstA("IdNaturaleza") & "','" & rstA("Ejercicio") & "','" & rstA("NroCuenta") & "','" & rstA("Descripcion") & "'," & _
       "'" & rstA("MonedaExtranjera") & "','" & rstA("IndCuentaDependiente") & "','" & rstA("IdCuentaContableDepende") & "','" & rstA("CtaCtbleDepende") & "','" & rstA("IndMovimiento") & "'," & _
       "'" & rstA("DigitoSubfijo") & "','" & rstA("CuentaSUNAT") & "','" & rstA("IndFlujoCaja") & "','" & rstA("IndConciliacion") & "','" & rstA("IndDocumento") & "','" & rstA("IndObligacion") & "'," & _
       "'" & rstA("IndDebe") & "','" & rstA("IndHaber") & "','" & rstA("IndGastoFuncion") & "','" & rstA("IndItemGasto") & "','" & rstA("IndCentroCosto") & "','" & rstA("IndTrabajador") & "','" & rstA("IndTracto") & "','" & rstA("IndRuta") & "','" & rstA("IndBanco") & "','" & rstA("Analisis01") & "' " & _
       ",'" & rstA("Analisis02") & "','" & rstA("Tesoreria") & "','" & rstA("Activo") & "','" & KEY_USUARIO & "','" & rstA("FechaCrea") & "','" & rstA("UsuarioModifica") & "','" & rstA("FechaModifica") & "')"
       CnBd.Execute (strCadena)
       rstA.MoveNext
   Next i
End If

'strCadena = "DELETE FROM con_cuentaasociada WHERE IdEmpresaSis='2048aaa7725286'"
'CnBd.Execute (strCadena)

strCadena = "SELECT * FROM con_cuentaasociada WHERE IdEmpresaSis='20487725286'"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   in_inicial = "1CIX000000001073"
   For i = 0 To rstA.RecordCount - 1
   
       in_id = Trim("1CIX" + Format(1742 + i, "000000000000"))
       
       'strCadena = "INSERT ITO con_cuentaasociada(`Id`,`IdEmpresaSis`,`IdSucursal`,`CuentaContable`,`CuentaAsociada1`,`DebeHaber1`,`CuentaAsociada2`,`Porcentaje1`,`DebeHaber2`,`CuentaAsociada3`,`Porcentaje3`,`DebeHaber3`,`Depreciacion`,`UsuarioCrea`,`FechaCrea`,`UsuarioModifica`,`FechaModifica`,`Activo`"
       
       
       strCadena = "INSERT INTO con_cuentaasociada(`Id`,`IdEmpresaSis`,`IdSucursal`,`CuentaContable`,`CuentaAsociada1`,`DebeHaber1`,`CuentaAsociada2`,Porcentaje1,`DebeHaber2`,`CuentaAsociada3`,Porcentaje3,`DebeHaber3`,Depreciacion,`UsuarioCrea`,`FechaCrea`,`UsuarioModifica`,`FechaModifica`,`Activo`)VALUES " & _
       "('" & in_id & "','" & KEY_RUC & "','1CIX124','" & rstA("CuentaContable") & "','" & rstA("CuentaAsociada1") & "','" & rstA("DebeHaber1") & "','" & rstA("CuentaAsociada2") & "','" & rstA("Porcentaje1") & "','" & rstA("DebeHaber2") & "','" & rstA("CuentaAsociada3") & "','" & rstA("Porcentaje3") & "','" & rstA("DebeHaber3") & "','" & rstA("Depreciacion") & "','" & KEY_USUARIO & "',CURDATE(),'0',CURDATE(),'" & rstA("Activo") & "')"
       CnBd.Execute (strCadena)
       rstA.MoveNext
       
   Next i
End If



strCadena = "SELECT cod_unico as dni FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "'  ORDER BY cod_unico DESC "
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       
       If Len(rst2("dni")) = 11 Then
       If Len(Trim(rst2("dni"))) = 8 Or Len(Trim(rst2("dni"))) = 11 Then
nnnn:
       strCadena = "SELECT * FROM persona WHERE   dni='" & Trim(rst2("dni")) & "'  LIMIT 1"
       Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                If Len(Trim(rst2("dni"))) = 8 Then
                 If get_dni_reniec_iii(Trim(rst2("dni")), KEY_DEPARTAMENTO, KEY_PROVINCIA, KEY_DISTRITO) = True Then
                    GoTo nnnn
                 End If
                    GoTo N
                End If
                
                strCadena = "SELECT * FROM persona WHERE   dni='" & Trim(rst2("dni")) & "'  LIMIT 1"
                Call ConfiguraRst3(strCadena)
                If rst3.RecordCount > 0 Then
                strCadena = "call P_insert_persona_ii('" & rst2("dni") & "' " & _
                ",'" & rst3("a_paterno") & "', " & _
                "'" & rst3("a_materno") & "' " & _
                ",'" & rst3("nombres") & "' " & _
                ",'" & rst3("nombre_completo") & "' " & _
                ",'" & rst3("direccion") & "' " & _
                ",'" & rst3("celular") & "' " & _
                ",'" & rst3("mail") & "'" & _
                ",'no' " & _
                ",'no'" & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                End If
                GoTo nn
       Else
nn:
                
                strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' and cod_unico='" & Trim(rst2("dni")) & "' LIMIT 1 "
                Call ConfiguraRstlocal(strCadena)
                If rstLocal.RecordCount < 1 Then
                    strCadena = "INSERT INTO entidad_empresa(cod_unico,id_cliente,id_empresa)VALUES ('" & rst2("dni") & "','si','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                End If
       End If
N:
       End If
       End If
         rst2.MoveNext
         DoEvents
         
   Next i
   
  


   
  strCadena = "DELETE FROM movimiento_venta WHERE ruc='" & KEY_RUC & "' "
  CnBd.Execute (strCadena)
   
 

   strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='0007' and ruc='" & KEY_RUC & "' ORDER BY id_venta ASC"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      rst.MoveFirst
      For i = 0 To rst.RecordCount - 1
           
           strCadena = "call P_insert_venta_agenda_test(" & rst("id_venta") & ")"
           CnBd.Execute (strCadena)
           rst.MoveNext
           DoEvents
      Next i
   End If
   


migra_venta:
  
   strCadena = "SELECT * FROM movimiento_venta WHERE id_venta>=659172 and   id_doc IN ('0001','0003','0007') and   ruc='" & KEY_RUC & "' and fecha_emision>='2019-01-01' and fecha_emision<='2019-01-31'  ORDER BY fecha_emision ASC,id_venta ASC"
   Call ConfiguraRstCloud(strCadena)
   If rstCloud.RecordCount > 0 Then
      rstCloud.MoveFirst
      Me.ProgressBar2.Min = 0
      Me.ProgressBar2.Max = rstCloud.RecordCount
      For i = 0 To rstCloud.RecordCount - 1
                      
            
            Call insertar_item_venta(rstCloud("id_venta"), rstCloud("id_cliente"), rstCloud("id_alm"), rstCloud("id_doc"), rstCloud("serie"), rstCloud("numero"), rstCloud("dni_save"))
            
            strCadena = "call CON_InsertaPeriodoNuevo('" & Format(rstCloud("fecha_emision"), "YYYY-mm-dd") & "','" & KEY_RUC & "','42546269')"
            CnBd.Execute (strCadena)
           
            
            strCadena = "DELETE from movimiento_venta_monto_temporal where  ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
        
            strCadena = "SELECT * FROM movimiento_venta_monto WHERE id_venta='" & rstCloud("id_venta") & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRst2(strCadena)
            If rst2.RecordCount > 0 Then
               rst2.MoveFirst
               For j = 0 To rst2.RecordCount - 1
                        forma_pago = "01"
                        id_forma = "89"
                   
                   strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,cuotas,id_usuario,id_alm,fecha,cuenta_contable,ruc)VALUES " & _
                   "('" & rstCloud("id_doc") & "','" & rstCloud("serie") & "','" & rstCloud("numero") & "','" & forma_pago & "','" & id_forma & "','" & rstCloud("id_moneda") & "','" & rst2("monto") & "','" & rst2("monto_caja") & "','00','0','" & rstCloud("dni_save") & "','" & rstCloud("id_alm") & "','" & Format(rstCloud("fecha_emision"), "YYYY-mm-dd") & "','10111','" & KEY_RUC & "')"
                   CnBd.Execute (strCadena)
                   rst2.MoveNext
               Next j
            Else
                forma_pago = "01"
                id_forma = "89"
                strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,cuotas,id_usuario,id_alm,fecha,cuenta_contable,ruc)VALUES " & _
                "('" & rstCloud("id_doc") & "','" & rstCloud("serie") & "','" & rstCloud("numero") & "','" & forma_pago & "','" & id_forma & "','" & rstCloud("id_moneda") & "','" & rstCloud("total") & "','" & rstCloud("total") & "','00','0','" & rstCloud("dni_save") & "','" & rstCloud("id_alm") & "','" & Format(rstCloud("fecha_emision"), "YYYY-mm-dd") & "','10111','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
            
            If rstCloud("id_comprobante") > 0 Then
               in_comprobante = get_id_comprobante_aurora(rstCloud("id_comprobante"))
            Else
                in_comprobante = 0
            End If
            
            
            'strCadena = "call p_insert_venta_cabecera_premiun('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcFormapago.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
            "'" & Trim(Me.DtcSerieDoc.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Me.TxtCodCliente.Text & "','" & Me.txtcliente.Text & "','" & SUBTOTAL & "','" & igv & "','" & exonerado & "','" & TotalVenta & "','" & Saldo & "'," & _
            "'" & Val(Me.lblPago.Caption) & "','" & Val(Me.lblVuelto.Caption) & "','" & KEY_FECHA & "','" & KEY_VENCIMIENTO & "','" & id_tipo_factura & "','" & Me.Dtcvendedor.BoundText & "','" & KEY_USUARIO & "','" & Val(KEY_CAMBIO_COMPRA) & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "'" & _
            ",'" & Documento & "','" & horario & "','T','" & Trim(Me.TxtDireccion.Text) & "','" & strconyugue & "','" & Trim(Me.txt_hash.Text) & "','" & Trim(Me.txt_sunat_key.Text) & "','" & Trim(Me.DtcTipoNota.BoundText) & "','" & Trim(Me.txtmotivo_nota.Text) & "','" & id_guia & "','" & in_guia & "','" & KEY_VENTANILLA & "','" & Trim(Me.txt_tipo.Text) & "','" & in_seguro & "','" & Trim(Me.txtobservacion.Text) & "','" & Trim(Me.txteditable.Text) & "','" & KEY_CONTABILIDAD & "','" & in_cta_cobrar & "','" & in_cta_ingreso & "','" & Val(Me.TxtDescuento_global.Text) & "','" & Val(Me.TxtCuotas.Text) & "','" & in_interes & "','" & Val(Me.txtid_venta_ref.Text) & "','" & in_diferida & "','" & KEY_RUC & "')"
            'Call ConfiguraRstPP(strCadena)
            'id_venta = rstPP("in_venta")
          '  If rstCloud("fecha_emision") < "16-01-2019" Then
           '     n_fecha = "2019-01-16"
          '  Else
                n_fecha = rstCloud("fecha_emision")
          '  End If
            strCadena = "call p_insert_venta_cabecera_premiun('" & rstCloud("id_doc") & "','" & rstCloud("id_alm") & "','" & rstCloud("id_forma_pago") & "','" & rstCloud("id_moneda") & "','" & rstCloud("id_delivery") & "'," & _
            "'" & rstCloud("serie") & "','" & rstCloud("numero") & "','" & rstCloud("id_cliente") & "','" & rstCloud("ncliente") & "','" & rstCloud("valor_venta") & "','" & rstCloud("igv") & "','" & rstCloud("exonerado") & "','" & rstCloud("total") & "','" & rstCloud("saldo") & "', " & _
            "'" & rstCloud("monto_pago") & "','" & rstCloud("monto_vuelto") & "','" & Format(n_fecha, "YYYY-mm-dd") & "','" & Format(n_fecha, "YYYY-mm-dd") & "','" & rstCloud("id_tipo_factura") & "','" & rstCloud("id_vendedor") & "','" & rstCloud("dni_save") & "','" & rstCloud("tc") & "','no','" & Format(Month(rstCloud("fecha_emision")), "00") & "','" & Year(rstCloud("fecha_emision")) & "'" & _
            ",'" & rstCloud("documento") & "','" & rstCloud("hora") & "','" & rstCloud("turno") & "','" & rstCloud("direccion") & "','0','-','-','" & rstCloud("id_tipo_nota") & "','" & rstCloud("motivo_nota") & "','" & rstCloud("id_guia") & "','" & rstCloud("nguia") & "','" & rstCloud("id_ventanilla") & "','01', " & _
            "'0','" & rstCloud("observacion") & "','no','si','1212','70111','0','0','0','" & in_comprobante & "','no','" & KEY_RUC & "')"
            Call ConfiguraRstPP(strCadena)
            id_venta = rstPP("in_venta")
            If IsNull(rstCloud("sunat_key")) = True Then
                strCadena = "UPDATE movimiento_venta SET sunat_key='" & rstCloud("sunat_key") & "',sunat_hash='" & rstCloud("sunat_hash") & "' WHERE id_venta='" & Val(id_venta) & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                'If KEY_FACTURACION_ELECTRONICA = "si" Then
                'If get_firma_online(in_doc, in_serie) = "si" Then
                '   Call firma_electronica(rstCloud("id_doc"), "no", " ", id_venta, rstCloud("numero"), rstCloud("serie"), rstCloud("id_cliente"), rstCloud("ncliente"), rstCloud("direccion"))
                   
                ' End If
                ' End If
           
            End If
            
            StrNumero = Format(Trim(str(Val(rstCloud("numero"))) + 1), "000000")
            strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE  id_doc='" & rstCloud("id_doc") & "' AND serie='" & rstCloud("serie") & "'  AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)

          rstCloud.MoveNext
          
          DoEvents
          
          
      Next i
   End If
   MsgBox "Yaaaaa"
Exit Sub
    '[MIGRAR COMPRAS]
    
plan_compras:
    
   in_periodo_victor = "1CIX000000000003"
   in_periodo_aurora = "1CIX000000000005"
   strCadena = "SELECT * FROM movimiento_compra WHERE   ruc='" & KEY_RUC & "' ORDER BY id_compra asc"
   Call ConfiguraRstCloud(strCadena)
   If rstCloud.RecordCount > 1 Then
      rstCloud.MoveFirst
      For i = 0 To rstCloud.RecordCount - 1
            
        Call verificar_existencia_cliente(rstCloud("id_proveedor"))
        
        strCadena = "call P_insert_compra_test('" & rstCloud("id_doc") & "','" & rstCloud("id_alm") & "','" & Format(rstCloud("fecha_emision"), "YYYY-mm-dd") & "','" & Format(rstCloud("fecha_cancelacion"), "YYYY-mm-dd") & "','" & rstCloud("id_forma_pago") & "'," & _
        "'" & rstCloud("id_tipo_compra") & "','" & rstCloud("anio_dua") & "','" & rstCloud("id_moneda") & "','" & Format(Month(rstCloud("fecha_emision")), "00") & "','" & Year(rstCloud("fecha_emision")) & "','" & rstCloud("serie") & "'," & _
        "'" & rstCloud("numero") & "','" & rstCloud("tipo_doc_identidad") & "','" & rstCloud("id_proveedor") & "','" & rstCloud("nproveedor") & "','" & Val(rstCloud("tc")) & "'," & _
        "'0','" & rstCloud("valor_venta") & "','" & rstCloud("igv") & "','" & rstCloud("isc") & "','" & rstCloud("ivap") & "','" & rstCloud("percepcion") & "','" & rstCloud("retencion") & "','" & rstCloud("exonerado") & "','" & rstCloud("otros") & "','" & rstCloud("total") & "','" & rstCloud("saldo") & "','" & rstCloud("dni_save") & "','" & rstCloud("observacion") & "','01','" & in_periodo_aurora & "','42121','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        'strCadena = "call P_insert_compra_test('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Format(CVDate(Me.TxtFecha_emision.Text), "YYYY-mm-dd") & "','" & Format(CVDate(Me.txtfecha_Vencimiento.Text), "YYYY-mm-dd") & "','02'," & _
        "'" & Me.DtTipoCompra.BoundText & "','--','" & Me.DtcMoneda.BoundText & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Trim(Me.TxtSerie.Text) & "'," & _
        "'" & Format(Trim(Me.TxtNumeroDoc.Text), "00000000") & "','" & cod_identidad & "','" & Trim(Me.TxtRuc.Text) & "','" & UCase(Me.TxtProveedor.Text) & "','" & Trim(Me.TxtTc.Text) & "'," & _
        "'0','" & Val(Me.LblValorVenta.Text) & "','" & Val(Me.LblIgv.Text) & "','" & Val(Me.lblISC.Text) & "','0','" & Val(Me.TxtPecepcion.Text) & "','0','" & Val(Me.lblExonerado.Text) & "','0','" & Val(Me.lblTotal.Text) & "','" & Val(Me.lblTotal.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtObservacion.Text) & "','" & Me.DtcTipo.BoundText & "','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & KEY_RUC & "')"
       ' CnBd.Execute (strCadena)
        
        id_compra = LastRegistroRUC("movimiento_compra", "id_compra")
        
        strCadena = "p_update_proveedor('" & rstCloud("id_proveedor") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        'NOTA DE CREDITO
        If rstCloud("id_doc") = "0007" Then
            strCadena = "UPDATE movimiento_compra SET valor_venta='" & rstCloud("valor_venta") * -1 & "',igv='" & rstCloud("igv") * -1 & "',isc='" & rstCloud("isc") * -1 & "',percepcion='" & Format("percepcion") * -1 & "',exonerado='" & rstCloud("exonerado") * -1 & "',total='" & rstCloud("total") * -1 & "',fecha_fact='" & Format(rstCloud("fecha_fact"), "YYYY-mm-dd") & "',id_doc_fact='" & rstCloud("id_doc_fact") & "',serie_fact='" & rstCloud("serie_fact") & "',numero_fact='" & rstCloud("numero_fact") & "' WHERE id_compra='" & id_compra & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
        End If
        If rstCloud("id_tipo_compra") = "01" Then
            strCadena = "INSERT INTO movimiento_compra_importacion(id_compra,ruc)VALUES('" & id_compra & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
        End If
        Call SaveDetalleDocumentoCompra(id_compra, rstCloud("id_compra"), rstCloud("fecha_emision"), rstCloud("id_doc"), rstCloud("serie"), rstCloud("numero"), rstCloud("id_proveedor"), rstCloud("id_alm"), Val(rstCloud("tc")))
        If KEY_CONTABILIDAD = "si" And rstCloud("id_doc") <> "0089" Then
            strCadena = "call p_insert_compra_emitido_ii('" & id_compra & "')"
            Call Execute_Sql(strCadena)
        End If
        If rstCloud("id_doc") = "0089" Then
            num = Format(Val(rstCloud("numero")) + 1, "000000")
            strCadena = "UPDATE almacen_comprobante SET numero='" & num & "' WHERE id_doc='" & rstCloud("id_doc") & "' AND serie='" & rstCloud("serie") & "' AND id_alm='" & rstCloud("id_alm") & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
    
        End If
            
            
          rstCloud.MoveNext
          DoEvents
          
      Next i
   End If
   
   
End If

End Sub



Public Sub migrar_ventas(ByVal in_ruta As String)
Dim sys_ConString2 As String
Dim stock_actual As Integer
'Call update_cuenta_contable



'Exit Sub
sys_Server2 = in_ruta
'sys_DataBase2 = "bd_vitekey_repos_lenin"
sys_DataBase2 = "bd_vitekey_aurora"
sys_DataBase2 = "bd_demo"

sys_SUser2 = "user_cord" 'DecryptString(ConfigRead("SUser"))
'sys_SPassword2 = "@02021974abc2016@123@cord" 'DecryptString(ConfigRead("SPassword"))
'sys_SPassword2 = "02021974abc2016" 'DecryptString(ConfigRead("SPassword"))
sys_SPassword2 = "vitekey2018"

db_port = "3306"
sys_ConString2 = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server2 & ";" & _
            "Database=" & sys_DataBase2 & ";" & _
            "UID=" & sys_SUser2 & ";" & _
            "PWD=" & sys_SPassword2 & ";" & _
            " PORT=" & db_port & ";OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
CnBd2.ConnectionString = sys_ConString2
CnBd2.Open

GoTo migra_venta

CnBd.Execute (strCadena)
strCadena = "SELECT * FROM movimiento_compra_detalle where ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "DELETE from movimiento_compra_detalle where  ruc='" & KEY_RUC & "' LIMIT 100"
        CnBd.Execute (strCadena)
        rst.MoveNext
        DoEvents
   Next i
End If


strCadena = "DELETE from imp_producto_detalle where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


'[ELIMINAR GUIAS DE REMISION]
strCadena = "DELETE FROM movimiento_transferencia where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE from movimiento_transferencia_detalle where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


'[ELIMINAR PRODUCTOS]
strCadena = "DELETE FROM producto WHERE    ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
strCadena = "DELETE FROM almacen_producto WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
strCadena = "DELETE FROM almacen_producto_precio WHERE  ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


'[ELIMINAR LINEAS]
strCadena = "DELETE FROM linea where id_usu='" & KEY_RUC & "' "
CnBd.Execute (strCadena)


strCadena = "DELETE FROM linea_sub WHERE id_usu='" & KEY_RUC & "' "
CnBd.Execute (strCadena)


'[ ELIMINAR ALMACENES ].
strCadena = "DELETE FROM almacen WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM almacen WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO almacen(`id_alm`,`ruc`,`descripcion`,`abreviatura`,`direccion`,`ubicacion_interna`,`id_responsable`,`id_tipoentidad`,`piso`,`pabellon`,`id_especialidad`,`observacion`,`stock`,`stock_personalizado`,`hora_inicio`,`hora_fin`,`horas`,`defecto`,`activo`,`id_atension`,`id_actividad`,`dni_save`,`ocupado`,`id_sucursal`,`facturacion_detallada`,`facturacion_centralizada`,`caja_independiente`,`cloud`,`comprobante_adicional`)VALUES " & _
       "('" & rst2("id_alm") & "','" & rst2("ruc") & "','" & rst2("descripcion") & "','" & rst2("abreviatura") & "','" & rst2("direccion") & "','" & rst2("ubicacion_interna") & "','" & rst2("id_responsable") & "','" & rst2("id_tipoentidad") & "','" & rst2("piso") & "','" & rst2("pabellon") & "','" & rst2("id_especialidad") & "','" & rst2("observacion") & "','" & rst2("stock") & "', " & _
       "'" & rst2("stock_personalizado") & "','" & Format(rst2("hora_inicio"), "HH:mm:ss") & "','" & Format(rst2("hora_fin"), "HH:mm:ss") & "','" & Format(rst2("horas"), "HH:mm:ss") & "','" & rst2("defecto") & "','" & rst2("activo") & "','" & rst2("id_atension") & "','" & rst2("id_actividad") & "','" & rst2("dni_save") & "','" & rst2("ocupado") & "','" & rst2("id_sucursal") & "','" & rst2("facturacion_detallada") & "','no','" & rst2("caja_independiente") & "','" & rst2("cloud") & "','no')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If


'[2] importar sus Comprobantes.

strCadena = "DELETE FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO almacen_comprobante(`ruc`,`id_alm`,`id_doc`,`serie`,`numero`,`igv`,`defecto`,`id_moneda`,`venta`,`id_formato_impresion`,`serial`,`afecta_caja`,`tipo_movimiento`,`id_usuario`,`numero_caracteres`,`electronico`,`firmado_online`,`produccion`,`online`)VALUES " & _
       "('" & rst2("ruc") & "','" & rst2("id_alm") & "','" & rst2("id_doc") & "','" & rst2("serie") & "','" & rst2("numero") & "','" & rst2("igv") & "','" & rst2("defecto") & "','" & rst2("id_moneda") & "','" & rst2("venta") & "','" & rst2("id_formato_impresion") & "','" & rst2("serial") & "','" & rst2("afecta_caja") & "','" & rst2("tipo_movimiento") & "', " & _
       "'" & rst2("id_usuario") & "','" & rst2("numero_caracteres") & "','" & rst2("electronico") & "','" & rst2("firmado_online") & "','" & rst2("produccion") & "','" & rst2("online") & "')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If

'[3] importar sus Marcas.
strCadena = "DELETE FROM marca WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM marca WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
      strCadena = "INSERT INTO marca(id_marca,descripcion,id_usu)VALUES('" & rst2("id_marca") & "','" & rst2("descripcion") & "','" & KEY_RUC & "')"
      CnBd.Execute (strCadena)
      rst2.MoveNext
   Next i
End If


'[3] importar sus Lineas.
strCadena = "DELETE FROM linea WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM linea WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO linea(`id_linea`,`descripcion`,`id_tipo`,`afecto_garantia`,`planilla`,`produccion`,`id_usu`,`garantia`,`mantenimientos`,`nro_cuenta`)VALUES " & _
       "('" & rst2("id_linea") & "','" & rst2("descripcion") & "','" & rst2("id_tipo") & "','" & rst2("afecto_garantia") & "','" & rst2("planilla") & "','" & rst2("produccion") & "','" & KEY_RUC & "','" & rst2("garantia") & "','" & rst2("mantenimientos") & "','60111')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If

strCadena = "DELETE FROM unidad WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM unidad WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO unidad(id_und,abreviatura,descripcion,id_usu)VALUES " & _
       "('" & rst2("id_und") & "','" & rst2("abreviatura") & "','" & rst2("descripcion") & "','" & rst2("id_usu") & "')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If
'[4] importar sus Sub Lineas.
strCadena = "DELETE FROM linea_sub WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM linea_sub WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "INSERT INTO linea_sub(`id_tipo`,`id_linea`,`descripcion`,`id_usu`)VALUES " & _
       "('" & rst2("id_tipo") & "','" & rst2("id_linea") & "','" & rst2("descripcion") & "','" & rst2("id_usu") & "')"
       CnBd.Execute (strCadena)
       rst2.MoveNext
   Next i
End If


Producto:

'[IMPORTAR PRODUCTOS]
        strCadena = "SELECT * FROM producto WHERE ruc='" & KEY_RUC & "' order by id_producto ASC"
        Call ConfiguraRst2(strCadena)
        If rst2.RecordCount > 0 Then
           rst2.MoveFirst
           For i = 0 To rst2.RecordCount - 1
               strCadena = "SELECT * FROM producto WHERE id_producto='" & rst2("id_producto") & "' and ruc='" & KEY_RUC & "'"
               Call ConfiguraRstA(strCadena)
               If rstA.RecordCount < 1 Then
                    strCadena = "INSERT INTO producto (`id_producto`,id_tipo,`id_linea`,`id_sublinea`,`id_moneda`,`id_color`,`nombre_prod`,`id_unidad`,`nombre_comercial`,`id_marca`,`id_igv`,`dni_save`,`ruc`) VALUES " & _
                    "('" & rst2("id_producto") & "','01','" & rst2("id_linea") & "','" & rst2("id_sublinea") & "','00001','" & rst2("id_color") & "','" & rst2("nombre_prod") & "','" & rst2("id_unidad") & "','" & rst2("nombre_comercial") & "','" & rst2("id_marca") & "','si','" & rst2("dni_save") & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                Else
                    If rst2("nombre_prod") <> rstA("nombre_prod") Then
                        strCadena = "UPDATE producto SET nombre_prod='" & rst2("nombre_prod") & "' WHERE id_producto='" & rst2("id_producto") & "' and ruc='" & KEY_RUC & "'"
                        CnBd.Execute (strCadena)
                    End If
                End If
                
                
                
               rst2.MoveNext
               
           Next i
        End If
        
        
stock:
        strCadena = "DELETE FROM almacen_producto WHERE ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)

        strCadena = "SELECT * FROM almacen_producto WHERE id_alm='00001' and   ruc='" & KEY_RUC & "'"
        Call ConfiguraRst2(strCadena)
        If rst2.RecordCount > 0 Then
           rst2.MoveFirst
           For j = 0 To rst2.RecordCount - 1
               strCadena = "INSERT INTO almacen_producto(`id_alm`,`id_producto`,precio_venta,precio_compra,`ruc`,`habilitado`,stock,stock_contable) VALUES ('" & rst2("id_alm") & "','" & rst2("id_producto") & "','" & rst2("precio_venta") & "','" & rst2("precio_compra") & "','" & KEY_RUC & "','si','" & rst2("stock") & "','" & rst2("stock_factura") & "')"
               CnBd.Execute (strCadena)
               Call put_kardex_inventario(rst2("id_producto"), rst2("id_alm"), rst2("stock"), rst2("precio_costo"), rst2("precio_venta"), "1CIX000000000038")
               rst2.MoveNext
           Next j
        End If



'[5] importar sus Plan Contable.

Dim in_inicial As String
strCadena = "DELETE FROM con_cuentacontable WHERE IdEmpresaSis='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM con_cuentacontable WHERE IdEmpresaSis='20487725286' order by id ASC" ' ginsac
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   in_inicial = Mid(KEY_RUC, 9, 2) & Format(Trim(Val(KEY_ALM)), "00")
   For i = 0 To rstA.RecordCount - 1
       in_id = Trim(in_inicial + Mid(rstA("id"), 5, 16))
       strCadena = "INSERT INTO con_cuentacontable(`Id`,`IdEmpresaSis`,`IdSucursal`,`IdNaturaleza`,`Ejercicio`,`NroCuenta`,`Descripcion`,`MonedaExtranjera`,`IndCuentaDependiente`,`IdCuentaContableDepende`,`CtaCtbleDepende`," & _
       "`IndMovimiento`,`DigitoSubfijo`,`CuentaSUNAT`,`IndFlujoCaja`,`IndConciliacion`,`IndDocumento`,`IndObligacion`,`IndDebe`,`IndHaber`,`IndGastoFuncion`,`IndItemGasto`,`IndCentroCosto`,`IndTrabajador`,`IndTracto`," & _
       "`IndRuta`,`IndBanco`,`Analisis01`,`Analisis02`,`Tesoreria`,`Activo`,`UsuarioCrea`,`FechaCrea`,`UsuarioModifica`,`FechaModifica`)VALUES " & _
       "('" & in_id & "','" & KEY_RUC & "','1CIX124','" & rstA("IdNaturaleza") & "','" & rstA("Ejercicio") & "','" & rstA("NroCuenta") & "','" & rstA("Descripcion") & "'," & _
       "'" & rstA("MonedaExtranjera") & "','" & rstA("IndCuentaDependiente") & "','" & rstA("IdCuentaContableDepende") & "','" & rstA("CtaCtbleDepende") & "','" & rstA("IndMovimiento") & "'," & _
       "'" & rstA("DigitoSubfijo") & "','" & rstA("CuentaSUNAT") & "','" & rstA("IndFlujoCaja") & "','" & rstA("IndConciliacion") & "','" & rstA("IndDocumento") & "','" & rstA("IndObligacion") & "'," & _
       "'" & rstA("IndDebe") & "','" & rstA("IndHaber") & "','" & rstA("IndGastoFuncion") & "','" & rstA("IndItemGasto") & "','" & rstA("IndCentroCosto") & "','" & rstA("IndTrabajador") & "','" & rstA("IndTracto") & "','" & rstA("IndRuta") & "','" & rstA("IndBanco") & "','" & rstA("Analisis01") & "' " & _
       ",'" & rstA("Analisis02") & "','" & rstA("Tesoreria") & "','" & rstA("Activo") & "','" & KEY_USUARIO & "','" & rstA("FechaCrea") & "','" & rstA("UsuarioModifica") & "','" & rstA("FechaModifica") & "')"
       CnBd.Execute (strCadena)
       rstA.MoveNext
   Next i
End If

'strCadena = "DELETE FROM con_cuentaasociada WHERE IdEmpresaSis='2048aaa7725286'"
'CnBd.Execute (strCadena)

strCadena = "SELECT * FROM con_cuentaasociada WHERE IdEmpresaSis='20487725286'"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   in_inicial = "1CIX000000001073"
   For i = 0 To rstA.RecordCount - 1
   
       in_id = Trim("1CIX" + Format(1742 + i, "000000000000"))
       
       'strCadena = "INSERT ITO con_cuentaasociada(`Id`,`IdEmpresaSis`,`IdSucursal`,`CuentaContable`,`CuentaAsociada1`,`DebeHaber1`,`CuentaAsociada2`,`Porcentaje1`,`DebeHaber2`,`CuentaAsociada3`,`Porcentaje3`,`DebeHaber3`,`Depreciacion`,`UsuarioCrea`,`FechaCrea`,`UsuarioModifica`,`FechaModifica`,`Activo`"
       
       
       strCadena = "INSERT INTO con_cuentaasociada(`Id`,`IdEmpresaSis`,`IdSucursal`,`CuentaContable`,`CuentaAsociada1`,`DebeHaber1`,`CuentaAsociada2`,Porcentaje1,`DebeHaber2`,`CuentaAsociada3`,Porcentaje3,`DebeHaber3`,Depreciacion,`UsuarioCrea`,`FechaCrea`,`UsuarioModifica`,`FechaModifica`,`Activo`)VALUES " & _
       "('" & in_id & "','" & KEY_RUC & "','1CIX124','" & rstA("CuentaContable") & "','" & rstA("CuentaAsociada1") & "','" & rstA("DebeHaber1") & "','" & rstA("CuentaAsociada2") & "','" & rstA("Porcentaje1") & "','" & rstA("DebeHaber2") & "','" & rstA("CuentaAsociada3") & "','" & rstA("Porcentaje3") & "','" & rstA("DebeHaber3") & "','" & rstA("Depreciacion") & "','" & KEY_USUARIO & "',CURDATE(),'0',CURDATE(),'" & rstA("Activo") & "')"
       CnBd.Execute (strCadena)
       rstA.MoveNext
       
   Next i
End If



strCadena = "SELECT * FROM persona  ORDER BY dni"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       
       If Val(rst2("dni")) > 10 Then
       If Len(rst2("dni")) = 8 Or Len(rst2("dni")) = 11 Then
nnnn:
       strCadena = "SELECT * FROM persona WHERE   dni='" & Trim(rst2("dni")) & "'  LIMIT 1"
       Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                
                strCadena = "call P_insert_persona_ii('" & rst2("dni") & "' " & _
                ",'" & rst2("a_paterno") & "', " & _
                "'" & rst2("a_materno") & "' " & _
                ",'" & rst2("nombres") & "' " & _
                ",'" & rst2("nombre_completo") & "' " & _
                ",'" & rst2("direccion") & "' " & _
                ",'" & rst2("celular") & "' " & _
                ",'" & rst2("mail") & "'" & _
                ",'no' " & _
                ",'no'" & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                GoTo nnnn
       Else
                strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' and cod_unico='" & Trim(rst2("dni")) & "' LIMIT 1 "
                Call ConfiguraRstlocal(strCadena)
                If rstLocal.RecordCount < 1 Then
                    strCadena = "INSERT INTO entidad_empresa(cod_unico,id_cliente,id_empresa)VALUES ('" & rst2("dni") & "','si','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                End If
       End If
         
       End If
       End If
         rst2.MoveNext
         '  DoEvents
   Next i
   
  


   
  strCadena = "DELETE FROM movimiento_venta WHERE ruc='" & KEY_RUC & "' "
  CnBd.Execute (strCadena)
   
 

   strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='0007' and ruc='" & KEY_RUC & "' ORDER BY id_venta ASC"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      rst.MoveFirst
      For i = 0 To rst.RecordCount - 1
           
           strCadena = "call P_insert_venta_agenda_test(" & rst("id_venta") & ")"
           CnBd.Execute (strCadena)
           rst.MoveNext
           DoEvents
      Next i
   End If
   


migra_venta:
  
   strCadena = "SELECT * FROM movimiento_venta WHERE id_venta>='" & Val(Me.txtidVenta.Text) & "' and  id_doc IN ('0001','0003','0007') and   ruc='" & KEY_RUC & "' and fecha_emision>='" & Format(Me.DtpInicio_migracion.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin_migracion.Value, "YYYY-mm-dd") & "'  ORDER BY fecha_emision ASC,id_venta ASC"
   Call ConfiguraRstCloud(strCadena)
   If rstCloud.RecordCount > 0 Then
      rstCloud.MoveFirst
      Me.ProgressBar2.Min = 0
      Me.ProgressBar2.Max = rstCloud.RecordCount
      For i = 0 To rstCloud.RecordCount - 1
                      
            
            
            strCadena = "SELECT * FROM persona WHERE   dni='" & Trim(rstCloud("id_cliente")) & "'  LIMIT 1"
            Call ConfiguraRstIN(strCadena)
            If rstIN.RecordCount < 1 Then
                strCadena = "call P_insert_persona_ii('" & rstCloud("id_cliente") & "' " & _
                ",'-', " & _
                "'-' " & _
                ",'-' " & _
                ",'" & rstCloud("ncliente") & "' " & _
                ",'" & rstCloud("direccion") & "' " & _
                ",'-' " & _
                ",'-'" & _
                ",'no' " & _
                ",'no'" & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
             Else
                strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' and cod_unico='" & Trim(rstCloud("id_cliente")) & "' LIMIT 1 "
                Call ConfiguraRstIN(strCadena)
                If rstIN.RecordCount < 1 Then
                    strCadena = "INSERT INTO entidad_empresa(cod_unico,id_cliente,id_empresa)VALUES ('" & rstCloud("id_cliente") & "','si','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                End If
            End If
       
       
            
            Call insertar_item_venta(rstCloud("id_venta"), rstCloud("id_cliente"), rstCloud("id_alm"), rstCloud("id_doc"), rstCloud("serie"), rstCloud("numero"), rstCloud("dni_save"))
            
            strCadena = "call CON_InsertaPeriodoNuevo('" & Format(rstCloud("fecha_emision"), "YYYY-mm-dd") & "','" & KEY_RUC & "','42546269')"
            CnBd.Execute (strCadena)
           
            
            strCadena = "DELETE from movimiento_venta_monto_temporal where  ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
        
            strCadena = "SELECT * FROM movimiento_venta_monto WHERE id_venta='" & rstCloud("id_venta") & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRst2(strCadena)
            If rst2.RecordCount > 0 Then
               rst2.MoveFirst
               For j = 0 To rst2.RecordCount - 1
                        forma_pago = "01"
                        id_forma = "89"
                   
                   strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,cuotas,id_usuario,id_alm,fecha,cuenta_contable,ruc)VALUES " & _
                   "('" & rstCloud("id_doc") & "','" & rstCloud("serie") & "','" & rstCloud("numero") & "','" & forma_pago & "','" & id_forma & "','" & rstCloud("id_moneda") & "','" & rst2("monto") & "','" & rst2("monto_caja") & "','00','0','" & rstCloud("dni_save") & "','" & rstCloud("id_alm") & "','" & Format(rstCloud("fecha_emision"), "YYYY-mm-dd") & "','10111','" & KEY_RUC & "')"
                   CnBd.Execute (strCadena)
                   rst2.MoveNext
               Next j
            Else
                forma_pago = "01"
                id_forma = "89"
                strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,cuotas,id_usuario,id_alm,fecha,cuenta_contable,ruc)VALUES " & _
                "('" & rstCloud("id_doc") & "','" & rstCloud("serie") & "','" & rstCloud("numero") & "','" & forma_pago & "','" & id_forma & "','" & rstCloud("id_moneda") & "','" & rstCloud("total") & "','" & rstCloud("total") & "','00','0','" & rstCloud("dni_save") & "','" & rstCloud("id_alm") & "','" & Format(rstCloud("fecha_emision"), "YYYY-mm-dd") & "','10111','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
            
            If rstCloud("id_comprobante") > 0 Then
               in_comprobante = get_id_comprobante_aurora(rstCloud("id_comprobante"))
            Else
                in_comprobante = 0
            End If
            
            
            'strCadena = "call p_insert_venta_cabecera_premiun('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcFormapago.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
            "'" & Trim(Me.DtcSerieDoc.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Me.TxtCodCliente.Text & "','" & Me.txtcliente.Text & "','" & SUBTOTAL & "','" & igv & "','" & exonerado & "','" & TotalVenta & "','" & Saldo & "'," & _
            "'" & Val(Me.lblPago.Caption) & "','" & Val(Me.lblVuelto.Caption) & "','" & KEY_FECHA & "','" & KEY_VENCIMIENTO & "','" & id_tipo_factura & "','" & Me.Dtcvendedor.BoundText & "','" & KEY_USUARIO & "','" & Val(KEY_CAMBIO_COMPRA) & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "'" & _
            ",'" & Documento & "','" & horario & "','T','" & Trim(Me.TxtDireccion.Text) & "','" & strconyugue & "','" & Trim(Me.txt_hash.Text) & "','" & Trim(Me.txt_sunat_key.Text) & "','" & Trim(Me.DtcTipoNota.BoundText) & "','" & Trim(Me.txtmotivo_nota.Text) & "','" & id_guia & "','" & in_guia & "','" & KEY_VENTANILLA & "','" & Trim(Me.txt_tipo.Text) & "','" & in_seguro & "','" & Trim(Me.txtobservacion.Text) & "','" & Trim(Me.txteditable.Text) & "','" & KEY_CONTABILIDAD & "','" & in_cta_cobrar & "','" & in_cta_ingreso & "','" & Val(Me.TxtDescuento_global.Text) & "','" & Val(Me.TxtCuotas.Text) & "','" & in_interes & "','" & Val(Me.txtid_venta_ref.Text) & "','" & in_diferida & "','" & KEY_RUC & "')"
            'Call ConfiguraRstPP(strCadena)
            'id_venta = rstPP("in_venta")
          '  If rstCloud("fecha_emision") < "16-01-2019" Then
           '     n_fecha = "2019-01-16"
          '  Else
                n_fecha = rstCloud("fecha_emision")
          '  End If
          'If rstCloud("id_tipo_nota") Then
          'End If
          
          
           ' strCadena = "call p_insert_venta_cabecera_premiun('" & rstCloud("id_doc") & "','" & rstCloud("id_alm") & "','" & rstCloud("id_forma_pago") & "','" & rstCloud("id_moneda") & "','" & rstCloud("id_delivery") & "'," & _
            "'" & rstCloud("serie") & "','" & rstCloud("numero") & "','" & rstCloud("id_cliente") & "','" & rstCloud("ncliente") & "','" & rstCloud("valor_venta") & "','" & rstCloud("igv") & "','" & rstCloud("exonerado") & "','" & rstCloud("total") & "','" & rstCloud("saldo") & "', " & _
            "'" & rstCloud("monto_pago") & "','" & rstCloud("monto_vuelto") & "','" & Format(n_fecha, "YYYY-mm-dd") & "','" & Format(n_fecha, "YYYY-mm-dd") & "','" & rstCloud("id_tipo_factura") & "','" & rstCloud("id_vendedor") & "','" & rstCloud("dni_save") & "','" & rstCloud("tc") & "','no','" & Format(Month(rstCloud("fecha_emision")), "00") & "','" & Year(rstCloud("fecha_emision")) & "'" & _
            ",'" & rstCloud("documento") & "','" & rstCloud("hora") & "','" & rstCloud("turno") & "','" & rstCloud("direccion") & "','0','-','-','" & rstCloud("id_tipo_nota") & "','" & rstCloud("motivo_nota") & "','" & rstCloud("id_guia") & "','" & rstCloud("nguia") & "','" & rstCloud("id_ventanilla") & "','01', " & _
            "'0','" & rstCloud("observacion") & "','no','si','1212','70111','0','0','0','" & in_comprobante & "','no','" & KEY_RUC & "')"
            
            strCadena = "call p_insert_venta_cabecera_migracion('" & rstCloud("id_doc") & "','" & rstCloud("id_alm") & "','" & rstCloud("id_forma_pago") & "','" & rstCloud("id_moneda") & "','" & rstCloud("id_delivery") & "'," & _
            "'" & rstCloud("serie") & "','" & rstCloud("numero") & "','" & rstCloud("id_cliente") & "','" & rstCloud("ncliente") & "','" & rstCloud("valor_venta") & "','" & rstCloud("igv") & "','" & rstCloud("exonerado") & "','" & rstCloud("total") & "','" & rstCloud("saldo") & "', " & _
            "'" & rstCloud("monto_pago") & "','" & rstCloud("monto_vuelto") & "','" & Format(n_fecha, "YYYY-mm-dd") & "','" & Format(n_fecha, "YYYY-mm-dd") & "','" & rstCloud("id_tipo_factura") & "','" & rstCloud("id_vendedor") & "','" & rstCloud("dni_save") & "','" & rstCloud("tc") & "','no','" & Format(Month(rstCloud("fecha_emision")), "00") & "','" & Year(rstCloud("fecha_emision")) & "'" & _
            ",'" & rstCloud("documento") & "','" & rstCloud("hora") & "','" & rstCloud("turno") & "','" & rstCloud("direccion") & "','0','-','-','0','-','0','-','0','01', " & _
            "'0','" & rstCloud("observacion") & "','no','si','1212','70111','0','0','0','" & in_comprobante & "','no','" & KEY_RUC & "')"
            
            
            Call ConfiguraRstPP(strCadena)
            id_venta = rstPP("in_venta")
            
            strCadena = "call P_insert_venta_agenda_test('" & id_venta & "')"
            CnBd.Execute (strCadena)
            'If IsNull(rstCloud("sunat_key")) = True Then
                'strCadena = "UPDATE movimiento_venta SET sunat_key='" & rstCloud("sunat_key") & "',sunat_hash='" & rstCloud("sunat_hash") & "' WHERE id_venta='" & Val(id_venta) & "' and ruc='" & KEY_RUC & "'"
                'CnBd.Execute (strCadena)
                'If KEY_FACTURACION_ELECTRONICA = "si" Then
                'If get_firma_online(in_doc, in_serie) = "si" Then
                '   Call firma_electronica(rstCloud("id_doc"), "no", " ", id_venta, rstCloud("numero"), rstCloud("serie"), rstCloud("id_cliente"), rstCloud("ncliente"), rstCloud("direccion"))
                   
                ' End If
                ' End If
           
            'End If
            
            StrNumero = Format(Trim(str(Val(rstCloud("numero"))) + 1), "000000")
            strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE  id_doc='" & rstCloud("id_doc") & "' AND serie='" & rstCloud("serie") & "'  AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)

          rstCloud.MoveNext
          
          DoEvents
          Me.cmdrealizarmigracion.Caption = str(i) & Space(3) & str(rstCloud.RecordCount)
          
      Next i
   End If
   MsgBox "Migrado Exitosamente", vbInformation
Exit Sub
    '[MIGRAR COMPRAS]
    
plan_compras:
    
   in_periodo_victor = "1CIX000000000003"
   in_periodo_aurora = "1CIX000000000005"
   strCadena = "SELECT * FROM movimiento_compra WHERE   ruc='" & KEY_RUC & "' ORDER BY id_compra asc"
   Call ConfiguraRstCloud(strCadena)
   If rstCloud.RecordCount > 1 Then
      rstCloud.MoveFirst
      For i = 0 To rstCloud.RecordCount - 1
            
        Call verificar_existencia_cliente(rstCloud("id_proveedor"))
        
        strCadena = "call P_insert_compra_test('" & rstCloud("id_doc") & "','" & rstCloud("id_alm") & "','" & Format(rstCloud("fecha_emision"), "YYYY-mm-dd") & "','" & Format(rstCloud("fecha_cancelacion"), "YYYY-mm-dd") & "','" & rstCloud("id_forma_pago") & "'," & _
        "'" & rstCloud("id_tipo_compra") & "','" & rstCloud("anio_dua") & "','" & rstCloud("id_moneda") & "','" & Format(Month(rstCloud("fecha_emision")), "00") & "','" & Year(rstCloud("fecha_emision")) & "','" & rstCloud("serie") & "'," & _
        "'" & rstCloud("numero") & "','" & rstCloud("tipo_doc_identidad") & "','" & rstCloud("id_proveedor") & "','" & rstCloud("nproveedor") & "','" & Val(rstCloud("tc")) & "'," & _
        "'0','" & rstCloud("valor_venta") & "','" & rstCloud("igv") & "','" & rstCloud("isc") & "','" & rstCloud("ivap") & "','" & rstCloud("percepcion") & "','" & rstCloud("retencion") & "','" & rstCloud("exonerado") & "','" & rstCloud("otros") & "','" & rstCloud("total") & "','" & rstCloud("saldo") & "','" & rstCloud("dni_save") & "','" & rstCloud("observacion") & "','01','" & in_periodo_aurora & "','42121','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        'strCadena = "call P_insert_compra_test('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Format(CVDate(Me.TxtFecha_emision.Text), "YYYY-mm-dd") & "','" & Format(CVDate(Me.txtfecha_Vencimiento.Text), "YYYY-mm-dd") & "','02'," & _
        "'" & Me.DtTipoCompra.BoundText & "','--','" & Me.DtcMoneda.BoundText & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Trim(Me.TxtSerie.Text) & "'," & _
        "'" & Format(Trim(Me.TxtNumeroDoc.Text), "00000000") & "','" & cod_identidad & "','" & Trim(Me.TxtRuc.Text) & "','" & UCase(Me.TxtProveedor.Text) & "','" & Trim(Me.TxtTc.Text) & "'," & _
        "'0','" & Val(Me.LblValorVenta.Text) & "','" & Val(Me.LblIgv.Text) & "','" & Val(Me.lblISC.Text) & "','0','" & Val(Me.TxtPecepcion.Text) & "','0','" & Val(Me.lblExonerado.Text) & "','0','" & Val(Me.lblTotal.Text) & "','" & Val(Me.lblTotal.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtObservacion.Text) & "','" & Me.DtcTipo.BoundText & "','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & KEY_RUC & "')"
       ' CnBd.Execute (strCadena)
        
        id_compra = LastRegistroRUC("movimiento_compra", "id_compra")
        
        strCadena = "p_update_proveedor('" & rstCloud("id_proveedor") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        'NOTA DE CREDITO
        If rstCloud("id_doc") = "0007" Then
            strCadena = "UPDATE movimiento_compra SET valor_venta='" & rstCloud("valor_venta") * -1 & "',igv='" & rstCloud("igv") * -1 & "',isc='" & rstCloud("isc") * -1 & "',percepcion='" & Format("percepcion") * -1 & "',exonerado='" & rstCloud("exonerado") * -1 & "',total='" & rstCloud("total") * -1 & "',fecha_fact='" & Format(rstCloud("fecha_fact"), "YYYY-mm-dd") & "',id_doc_fact='" & rstCloud("id_doc_fact") & "',serie_fact='" & rstCloud("serie_fact") & "',numero_fact='" & rstCloud("numero_fact") & "' WHERE id_compra='" & id_compra & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
        End If
        If rstCloud("id_tipo_compra") = "01" Then
            strCadena = "INSERT INTO movimiento_compra_importacion(id_compra,ruc)VALUES('" & id_compra & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
        End If
        Call SaveDetalleDocumentoCompra(id_compra, rstCloud("id_compra"), rstCloud("fecha_emision"), rstCloud("id_doc"), rstCloud("serie"), rstCloud("numero"), rstCloud("id_proveedor"), rstCloud("id_alm"), Val(rstCloud("tc")))
        If KEY_CONTABILIDAD = "si" And rstCloud("id_doc") <> "0089" Then
            strCadena = "call p_insert_compra_emitido_ii('" & id_compra & "')"
            Call Execute_Sql(strCadena)
        End If
        If rstCloud("id_doc") = "0089" Then
            num = Format(Val(rstCloud("numero")) + 1, "000000")
            strCadena = "UPDATE almacen_comprobante SET numero='" & num & "' WHERE id_doc='" & rstCloud("id_doc") & "' AND serie='" & rstCloud("serie") & "' AND id_alm='" & rstCloud("id_alm") & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
    
        End If
            
            
          rstCloud.MoveNext
          DoEvents
          
      Next i
   End If
   
   
End If

End Sub



Public Sub update_precio()
strCadena = "SELECT * FROM almacen_producto WHERE ruc='20487376338' and id_alm='00001'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "UPDATE almacen_producto SET precio_venta='" & rst("precio_venta") & "',precio_compra='" & rst("precio_compra") & "',precio_mayor='" & rst("precio_mayor") & "' WHERE ruc='" & KEY_RUC & "' and id_producto='" & rst("id_producto") & "' LIMIT 1"
        CnBd.Execute (strCadena)
        rst.MoveNext
   Next i
   
End If
End Sub
Public Sub migrar_empresa_online()

'Call update_cuenta_contable
'Call update_precio
''
GoTo put_producto

strCadena = "DELETE FROM movimiento_venta where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE FROM movimiento_venta_detalle where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE from movimiento_venta_monto where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE from movimiento_venta_cuotas where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


'[ELIMINAR COMPRAS]
strCadena = "DELETE FROM movimiento_compra where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
strCadena = "SELECT * FROM movimiento_compra_detalle where ruc='" & KEY_RUC & "'"

strCadena = "DELETE from imp_producto_detalle where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


'[ELIMINAR GUIAS DE REMISION]
strCadena = "DELETE FROM movimiento_transferencia where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE from movimiento_transferencia_detalle where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


'[ELIMINAR PRODUCTOS]
strCadena = "DELETE FROM producto WHERE    ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
strCadena = "DELETE FROM almacen_producto WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
strCadena = "DELETE FROM almacen_producto_precio WHERE  ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


'[ELIMINAR LINEAS]
strCadena = "DELETE FROM linea where id_usu='" & KEY_RUC & "' "
CnBd.Execute (strCadena)


strCadena = "DELETE FROM linea_sub WHERE id_usu='" & KEY_RUC & "' "
CnBd.Execute (strCadena)


'[ ELIMINAR ALMACENES ].

'[2] importar sus Comprobantes.

'[3] importar sus Marcas.

put_producto:

strCadena = "DELETE FROM marca WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM marca WHERE id_usu='20393769468'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
      strCadena = "INSERT INTO marca(id_marca,descripcion,id_usu)VALUES('" & rst("id_marca") & "','" & rst("descripcion") & "','" & KEY_RUC & "')"
      CnBd.Execute (strCadena)
      rst.MoveNext
   Next i
End If


'[3] importar sus Lineas.
strCadena = "DELETE FROM linea WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM linea WHERE id_usu='20393769468'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "INSERT INTO linea(`id_linea`,`descripcion`,`id_tipo`,`afecto_garantia`,`planilla`,`produccion`,`id_usu`,`garantia`,`mantenimientos`,`nro_cuenta`)VALUES " & _
       "('" & rst("id_linea") & "','" & rst("descripcion") & "','" & rst("id_tipo") & "','" & rst("afecto_garantia") & "','" & rst("planilla") & "','" & rst("produccion") & "','" & KEY_RUC & "','" & rst("garantia") & "','" & rst("mantenimientos") & "','60111')"
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If



strCadena = "DELETE FROM unidad WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM unidad WHERE id_usu='20393769468'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "INSERT INTO unidad(id_und,abreviatura,descripcion,id_usu)VALUES " & _
       "('" & rst("id_und") & "','" & rst("abreviatura") & "','" & rst("descripcion") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If
'[4] importar sus Sub Lineas.
strCadena = "DELETE FROM linea_sub WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM linea_sub WHERE id_usu='20393769468'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "INSERT INTO linea_sub(`id_tipo`,`id_linea`,`descripcion`,`id_usu`)VALUES " & _
       "('" & rst("id_tipo") & "','" & rst("id_linea") & "','" & Replace(rst("descripcion"), "'", "") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If

'[IMPORTAR PRODUCTOS]
        strCadena = "SELECT * FROM producto WHERE ruc='" & KEY_RUC & "' order by id_producto DESC LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            in_producto = rst("id_producto")
        End If
        
        strCadena = "SELECT * FROM producto WHERE ruc='20393769468' order by id_producto ASC"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           
           For i = 0 To rst.RecordCount - 1
                 in_producto = Format(Val(in_producto) + 1, "00000")
                 strCadena = "INSERT INTO producto (`id_producto`,id_tipo,`id_linea`,`id_sublinea`,`id_moneda`,`id_color`,`nombre_prod`,`id_unidad`,`nombre_comercial`,`id_marca`,`id_igv`,icbper,codigo_alterno,`dni_save`,`ruc`) VALUES " & _
                 "('" & in_producto & "','01','" & rst("id_linea") & "','" & rst("id_sublinea") & "','00001','" & rst("id_color") & "','" & rst("nombre_prod") & "','" & rst("id_unidad") & "','" & Replace(rst("nombre_comercial"), "'", "") & "','" & rst("id_marca") & "','si','" & rst("icbper") & "','" & rst("id_producto") & "','" & rst("dni_save") & "','" & KEY_RUC & "')"
                 CnBd.Execute (strCadena)
                 
                 
              
               rst.MoveNext
               
           Next i
           
        End If
        
productos:
        
        strCadena = "DELETE FROM almacen_producto WHERE   ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)

        strCadena = "SELECT * FROM producto WHERE   ruc='20393769468' ORDER BY id_producto"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           
           
           For j = 0 To rst.RecordCount - 1
                
                strCadena = "SELECT * FROM almacen_producto WHERE id_alm='00001' and    id_producto='" & rst("id_producto") & "'  and   ruc='20393769468'  LIMIT 1 "
                Call ConfiguraRstK(strCadena)
                If rstK.RecordCount > 0 Then
                    in_precio_venta = rstK("precio_venta")
                    in_precio_compra = 0.01 'rstK("precio_compra")
                    in_precio_mayor = rstK("precio_mayor")
                    
               
                    
               End If
              
                    
                    
                   
                    
                    
                    strCadena = "INSERT INTO almacen_producto(`id_alm`,`id_producto`,precio_venta,precio_compra,`ruc`,`habilitado`,precio_mayor) VALUES " & _
                    " ('00001','" & rst("id_producto") & "','" & Val(in_precio_venta) & "','" & Val(in_precio_compra) & "','" & KEY_RUC & "','si','" & Val(in_precio_mayor) & "')"
                   CnBd.Execute (strCadena)
            
               
               
               
             
               rst.MoveNext
           Next j
        End If





                

'[5] importar sus Plan Contable.
   MsgBox "Yaaaaa"
Exit Sub
End Sub
Private Sub delete_espacio()
strCadena = "SELECT * FROM con_cuentacontable WHERE IdEmpresaSis='" & KEY_RUC & "' ORDER BY id ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       in_descripcion = ""
       in_descripcion = Trim(rst("Descripcion"))
       strCadena = "UPDATE con_cuentacontable SET Descripcion='" & in_descripcion & "' WHERE id='" & rst("id") & "' and IdEmpresaSis='" & KEY_RUC & "'"
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If
End Sub
Private Function get_id_comprobante_aurora(ByVal in_venta_victor As String) As Double
strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & in_venta_victor & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & rst2("id_doc") & "' and serie='" & rst2("serie") & "' and numero='" & rst2("numero") & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstlocal(strCadena)
    If rstLocal.RecordCount > 0 Then
       get_id_comprobante_aurora = rstLocal("id_venta")
    Else
        get_id_comprobante_aurora = 0
    End If
End If

End Function
Private Sub verificar_existencia_cliente(ByVal in_dni As String)

strCadena = "SELECT * FROM persona WHERE  dni='" & in_dni & "' LIMIT 1"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
          
       strCadena = "SELECT * FROM persona WHERE   dni='" & Trim(rst2("dni")) & "'  LIMIT 1"
       Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                
                strCadena = "call P_insert_persona_ii('" & rst2("dni") & "' " & _
                ",'" & rst2("a_paterno") & "', " & _
                "'" & rst2("a_materno") & "' " & _
                ",'" & rst2("nombres") & "' " & _
                ",'" & rst2("nombre_completo") & "' " & _
                ",'" & rst2("direccion") & "' " & _
                ",'" & rst2("celular") & "' " & _
                ",'" & rst2("mail") & "'" & _
                ",'no' " & _
                ",'no'" & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
       Else
                strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' and cod_unico='" & rst2("dni") & "' LIMIT 1 "
                Call ConfiguraRstlocal(strCadena)
                If rstLocal.RecordCount < 1 Then
                    strCadena = "INSERT INTO entidad_empresa(cod_unico,id_cliente,id_empresa,id_almacen)VALUES ('" & rst2("dni") & "','si','" & KEY_RUC & "','00001')"
                    CnBd.Execute (strCadena)
                End If
       End If
   
         '  DoEvents
   
   End If

End Sub
Private Function get_comprobante_nota(ByVal in_comprobante As String) As Double
strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & Val(in_comprobante) & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & rst2("id_doc") & "' and serie='" & rst2("serie") & "' and numero='" & rst2("numero") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
   Call ConfiguraRstP(strCadena)
   If rstP.RecordCount > 0 Then
      get_comprobante_nota = rstP("id_venta")
   Else
        get_comprobante_nota = 0
   End If
End If

End Function


Public Sub insertar_item_venta(ByVal in_venta As String, ByVal in_cliente As String, ByVal in_alm As String, ByVal in_tipo_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_save As String)
        
        
        strCadena = "DELETE from temporal_ventas where ruc='" & KEY_RUC & "' and dni_save='" & in_save & "'"
        CnBd.Execute (strCadena)
        strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & in_venta & "' "
        Call ConfiguraRst2(strCadena)
        If rst2.RecordCount > 0 Then
           rst2.MoveFirst
           For i = 0 To rst2.RecordCount - 1
                strCadena = "INSERT INTO temporal_ventas(ruc,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save) VALUES " & _
                "('" & KEY_RUC & "','" & in_cliente & "','" & Format(in_alm, "00000") & "','" & in_tipo_doc & "','" & in_serie & "','" & in_numero & "','" & rst2("id_producto") & "','" & rst2("cantidad") & "'," & _
                "'" & rst2("precio") & " ','" & rst2("total") & "','" & rst2("peso") & "','si','" & rst2("detalle") & "','" & in_save & "')"
                CnBd.Execute (strCadena)
            rst2.MoveNext
           Next i
        End If
        
        
      
    

End Sub
Private Sub put_estudiante(ByVal in_dni As String)
strCadena = "SELECT * FROM persona_estudiante WHERE dni='" & in_dni & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
            
            strCadena = "SELECT * FROM persona_estudiante WHERE dni='" & in_dni & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount < 1 Then
                strCadena = "call p_insert_estudiante('" & in_dni & "','" & rstLocal("id_nivel") & "','" & rstLocal("id_grado") & "','" & rstLocal("procedencia") & "','" & rstLocal("promovido") & "','" & rstLocal("tercio_estudiantil") & "','" & rstLocal("habilidad") & "','" & rstLocal("requiere_recuperacion") & "','" & rstLocal("id_tipo_nacimiento") & "','" & rstLocal("enfermedades") & "','" & rstLocal("vacunas") & "','" & rstLocal("alergias") & "','" & rstLocal("id_seguro") & "','" & rstLocal("vive_papa") & "','" & rstLocal("vive_mama") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            Else
                in_recuperacion = rst3("requiere_recuperacion")
                strCadena = "call p_update_estudiante('" & in_dni & "','" & rst3("id_nivel") & "','" & rst3("id_grado") & "','" & rst3("procedencia") & "','" & rst3("promovido") & "','" & rst3("tercio_estudiantil") & "','" & rst3("habilidad") & "','" & rst3("requiere_recuperacion") & "','" & rst3("id_tipo_nacimiento") & "','" & rst3("enfermedades") & "','" & rst3("vacunas") & "','" & rst3("alergias") & "','" & rst3("id_seguro") & "','" & rst3("vive_papa") & "','" & rst3("vive_mama") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If




End If












'If Me.chkseguro.Value = 1 Then
'   If Me.chk_estadoseguro.Value = 1 Then
'      in_estado_seguro = "si"
'   Else
'      in_estado_seguro = "no"
'   End If

strCadena = "SELECT * FROM persona_seguro WHERE dni='" & in_dni & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRst3(strCadena)
If rst3.RecordCount > 0 Then
   rst3.MoveFirst
   in_estado_seguro = rst3("activo")
   strCadena = "CALL p_put_seguro_persona('" & in_dni & "','" & rst3("id_seguro") & "','" & rst3("numero") & "','" & Format(rst3("expedicion"), "YYYY-mm-dd") & "','" & Format(rst3("expiracion"), "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & in_estado_seguro & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
End If


'End If
End Sub
Private Sub put_venta(ByVal in_venta As String)
            
            
        
            
            
            
        
        
End Sub



Private Sub familiares(ByVal dni As String)
    strCadena = "SELECT * FROM persona_accidentes WHERE dni='" & dni & "'"
    Call ConfiguraRstCloud(strCadena)
    If rstCloud.RecordCount > 0 Then
       rstCloud.MoveFirst
       For i = 0 To rstCloud.RecordCount - 1
           strCadena = "SELECT * FROM persona_accidentes WHERE dni='" & dni & "' and dni_familia='" & rstCloud("dni_familia") & "'"
           Call ConfiguraRst3(strCadena)
           If rst3.RecordCount < 1 Then
              strCadena = "INSERT INTO persona_accidentes(dni,dni_familia,id_parentesco,telefono,direccion,id_ocupacion,id_grado) " & _
             " VALUES('" & dni & "','" & rst3("dni_familia") & "','" & rst3("id_parentesco") & "','" & rst3("telefono") & "','" & rst3("direccion") & "','" & rst3("id_ocupacion") & "','" & rst3("id_grado") & "')"
           Else
              strCadena = "UPDATE persona_accidentes SET telefono='" & rst3("telefono") & "',id_parentesco='" & rst3("id_parentesco") & "',id_ocupacion='" & rst3("id_ocupacion") & "',id_grado='" & rst3("id_grado") & "',direccion='" & rst3("direccion") & "' WHERE dni_familia='" & rst3("dni_familia") & "' AND dni='" & dni & "' and id_parentesco='" & rst3("id_parentesco") & "'"
           End If
           CnBd.Execute (strCadena)
       rstCloud.MoveNext
       Next i
    End If
    
    
    
    
End Sub



Private Sub CmdImportar_Click()
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim rst4 As New ADODB.Recordset
Dim RstAlmProd As New ADODB.Recordset
Dim Base As String, Tabla As String, Criterio As String, TablaDestino As String
Base = Trim(Me.TxtNombreBaseOrigen.Text)
Tabla = Trim(Me.TxtNombreTablaOrigen.Text)
Criterio = Trim(Me.TxtCriterioOrigen.Text)
TablaDestino = Trim(Me.TxtNombreTablaDestino.Text)
Dim idProducto As String
cnbd1.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog='" & Base & "'" ';Data Source=192.168.1.33"
    
'GoTo compras
    
    
    strCadena = "DELETE FROM unidad WHERE id_usu='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    strCadena = "SELECT * FROM " & Tabla & " ORDER BY " & Criterio & " ASC"
    rst1.CursorLocation = adUseClient
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    If rst1.RecordCount > 0 Then
       rst1.MoveFirst
       For i = 0 To rst1.RecordCount - 1
             strCadena = "INSERT INTO " & TablaDestino & "(id_und,descripcion,abreviatura,id_usu)VALUES('" & formato_item(rst1("cUnidad"), 5) & "','" & rst1("sDescripcion") & "','" & Trim(rst1("sAbreviatura")) & "','" & KEY_RUC & "')"
             CnBd.Execute (strCadena)
              
             rst1.MoveNext
       Next i
       MsgBox "MIGRACION EXITOSA" + Space(2) + TablaDestino, vbInformation, KEY_EMPRESA
    End If
    Set rst1 = Nothing
    strCadena = "DELETE  FROM linea WHERE id_usu='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    
    strCadena = "SELECT * FROM Linea ORDER BY cLinea"
    rst1.CursorLocation = adUseClient
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    If rst1.RecordCount > 0 Then
       rst1.MoveFirst
       For i = 0 To rst1.RecordCount - 1
             strCadena = "INSERT INTO linea (id_linea,descripcion,id_usu)VALUES('" & formato_item(rst1("cLinea"), 5) & "','" & rst1("sDescripcion") & "','" & KEY_RUC & "')"
             CnBd.Execute (strCadena)
              
             rst1.MoveNext
       Next i
       MsgBox "MIGRACION EXITOSA" + Space(2) + "LINEA", vbInformation, KEY_EMPRESA
    End If
    Set rst1 = Nothing
    strCadena = "DELETE FROM marca WHERE id_usu='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    
    strCadena = "SELECT * FROM Marcas ORDER BY cMarca"
    rst1.CursorLocation = adUseClient
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    If rst1.RecordCount > 0 Then
       rst1.MoveFirst
       For i = 0 To rst1.RecordCount - 1
             strCadena = "INSERT INTO marca (id_marca,descripcion,id_usu)VALUES('" & formato_item(rst1("cMarca"), 5) & "','" & rst1("Mar_des") & "','" & KEY_RUC & "')"
             CnBd.Execute (strCadena)
              
             rst1.MoveNext
       Next i
       MsgBox "MIGRACION EXITOSA" + Space(2) + "MARCA", vbInformation, KEY_EMPRESA
    End If
    Dim Proveedor As String
    Set rst1 = Nothing
    strCadena = "SELECT * FROM Persona WHERE Per_Ruc<>'' ORDER BY cPersona"
    rst1.CursorLocation = adUseClient
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    If rst1.RecordCount > 0 Then
       rst1.MoveFirst
       For i = 0 To rst1.RecordCount - 1
            Proveedor = rst1("proveedor")
            If Trim(Proveedor) = "V" Then
                Proveedor = "si"
            Else
                Proveedor = "no"
            End If
            
            strCadena = "SELECT * FROM persona where dni='" & rst1("Per_Ruc") & "'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount < 0 Then
                strCadena = "P_insert_persona('" & Trim(rst1("Per_Ruc")) & "','-','-','-','" & Trim(rst1("NombrePersona")) & "','" & Trim(rst1("sDireccionCliente1")) & "','" & Trim(rst1("Telefono1")) & "','-','no','no','" & Proveedor & "','no','no','no','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                 
            Else
                
                strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & rst1("Per_Ruc") & "' AND id_empresa='" & KEY_RUC & "'"
                Call ConfiguraRstT(strCadena)
                If rstT.RecordCount < 1 Then
                    strCadena = "INSERT INTO entidad_empresa(cod_unico,id_empresa,id_cliente)VALUES('" & rst1("Per_Ruc") & "','" & KEY_RUC & "','si')"
                    CnBd.Execute (strCadena)
                     
                End If
            End If
             
             rst1.MoveNext
       Next i
       MsgBox "MIGRACION EXITOSA" + Space(2) + "PERSONA RUC", vbInformation, KEY_EMPRESA
    End If
    
    Set rst1 = Nothing
    
    strCadena = "SELECT * FROM Persona WHERE Per_Ruc='' and cPersona<>'00004' ORDER BY cPersona"
    rst1.CursorLocation = adUseClient
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    If rst1.RecordCount > 0 Then
       rst1.MoveFirst
       For i = 0 To rst1.RecordCount - 1
            Proveedor = rst1("proveedor")
            If Trim(Proveedor) = "V" Then
                Proveedor = "si"
            Else
                Proveedor = "no"
            End If
            
            strCadena = "DELETE FROM persona WHERE dni='" & formato_item(Trim(rst1("cPersona")), 8) & "'"
            CnBd.Execute (strCadena)
             
             strCadena = "P_insert_persona('" & formato_item(Trim(rst1("cPersona")), 8) & "','-','-','-','" & Trim(rst1("NombrePersona")) & "','" & Trim(rst1("sDireccionCliente1")) & "','" & Trim(rst1("Telefono1")) & "','-','no','no','" & Proveedor & "','no','no','no','" & KEY_RUC & "')"
             CnBd.Execute (strCadena)
              
             rst1.MoveNext
             DoEvents
       Next i
       MsgBox "MIGRACION EXITOSA" + Space(2) + "PERSONA RUC", vbInformation, KEY_EMPRESA
    End If
Set rst1 = Nothing

strCadena = "DELETE FROM producto WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
 
strCadena = "DELETE FROM almacen_producto WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
 

    strCadena = "SELECT * FROM Producto P,Almacen_Productos A WHERE P.cProducto=A.cProducto   ORDER BY A.cProducto"
    rst1.CursorLocation = adUseClient
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    If rst1.RecordCount > 0 Then
       rst1.MoveFirst
       For i = 0 To rst1.RecordCount - 1
             strCadena = "INSERT INTO producto (id_producto, id_unidad, id_linea, id_marca,nombre_prod,stock_total,stock_minimo,peso,id_percepcion,comentario,id_igv,id_sub_producto," & _
           "id_proveedor,id_auspiciador,id_combo,ruc,precio_delivery,imagen,id_tipo) VALUES ('" & formato_item(rst1("cProducto"), 5) & "','" & formato_item(rst1("cUnidad"), 5) & "','" & formato_item(rst1("cLinea"), 5) & "','" & formato_item(rst1("cMarca"), 5) & "'," & _
           "'" & Trim(rst1("DescripcionProducto")) & "','0','2','" & rst1("prod_peso") & "','no'," & _
           "'--','si','no','0','0','no','" & KEY_RUC & "'," & _
           "'0','','01')"
           CnBd.Execute (strCadena)
            
           strCadena = "UPDATE producto set precio_venta='" & rst1("PrecioVenta") & "',precio_compra='" & rst1("PrecioCompra") & "' WHERE id_producto='" & formato_item(rst1("cProducto"), 5) & "' AND ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
            
           
           strCadena = "SELECT * FROM almacen WHERE ruc='" & KEY_RUC & "'"
           
           RstAlmProd.CursorLocation = adUseClient
           RstAlmProd.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
           If RstAlmProd.RecordCount <= 0 Then
                MsgBox "No hay Ningun Almacen registrado", vbInformation
                MsgBox "Producto NO Grabado, Cree un Almacen", vbInformation
                Exit Sub
           End If
           RstAlmProd.MoveFirst
           For j = 0 To RstAlmProd.RecordCount - 1
             strCadena = "INSERT INTO almacen_producto(id_alm,id_producto,stock,stock_factura,ruc) VALUES ('" & formato_item(RstAlmProd("id_alm"), 5) & "','" & formato_item(rst1("cProducto"), 5) & "','" & rst1("Stock") & "','" & rst1("Stock_factura") & "','" & Trim(KEY_RUC) & "')"
             CnBd.Execute (strCadena)
              
             RstAlmProd.MoveNext
           Next j
           Set RstAlmProd = Nothing
             rst1.MoveNext
       Next i
       MsgBox "MIGRACION EXITOSA" + Space(2) + "PRODUCTOS", vbInformation, KEY_EMPRESA
    End If
    'GoTo inventario
    Set rst1 = Nothing
    '-------documento ventas
    
    'strCadena = "DELETE FROM movimiento_venta WHERE ruc='" & KEY_RUC & "'"
    'CnBd.Execute (strCadena)
'compras:
    fecha = CVDate("01-01-2000")
    
    strCadena = "DELETE FROM movimiento_venta WHERE ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    strCadena = "SELECT * FROM DocumentoVenta WHERE (doc_cod='0003' or doc_cod='0001') AND dEmisionVenta>='" & fecha & "' AND sSerie<>''  ORDER BY cDocumentoVenta"
    rst1.CursorLocation = adUseClient
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    If rst1.RecordCount > 0 Then
       rst1.MoveFirst
       For i = 0 To rst1.RecordCount - 1
            strCadena = "SELECT Per_Ruc FROM Persona WHERE cPersona='" & rst1("cPersona") & "'"
            rst4.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
            If rst4.RecordCount > 0 Then
            
            If rst4.RecordCount > 0 And rst4("Per_Ruc") <> "" Then
                If rst1("cPersona") = "00004" Or rst1("cPersona") = "00129" Then
                    dni = "00000000"
                Else
                    dni = rst4("Per_Ruc")
                End If
            Else
                dni = BDBuscarCampo("persona", "dni", "dni", formato_item(rst1("cPersona"), 8))
            End If
            Else
               dni = "00000000"
               If rst1("doc_cod") = "0001" Then
                GoTo siguiente
               End If
            End If
            Set rst4 = Nothing
                        
            'registro de ventas
            rmes = formato_item(Month(rst1("dEmisionVenta")), 2)
            ranio = formato_item(Year(rst1("dEmisionVenta")), 4)
            strCadena = "SELECT * FROM registro_ventas WHERE ruc='" & KEY_RUC & "' AND mes='" & rmes & "' AND anio='" & ranio & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                desVentas = "REGISTRO VENTAS :" + Space(3) + nombre_mes(rmes)
                strCadena = "INSERT INTO registro_ventas(ruc,mes,anio,descripcion)VALUES('" & KEY_RUC & "','" & rmes & "','" & ranio & "','" & desVentas & "')"
                CnBd.Execute (strCadena)
                 
            End If

    

            'fin registro ventas
            
            
            
            
            
            
            strCadena = "P_insert_venta('" & formato_item(rst1("doc_cod"), 4) & "','" & formato_item(rst1("Alm_cod"), 5) & "','01','00001','no'," & _
            "'" & Trim(formato_item(rst1("sSerie"), 3)) & "','" & Trim(formato_item(rst1("cDocumentoVenta"), 6)) & "','" & dni & "','" & Trim(rst1("Persona")) & "','" & Val(rst1("nSubTotal")) & "','" & Val(rst1("nIgv")) & "','0','" & Val(rst1("nTotalVenta")) & "','0'," & _
            "'" & Val(rst1("nTotalVenta")) & "','0','" & Format(rst1("dEmisionVenta"), "YYYY-mm-dd") & "','" & Format(rst1("dVencimiento"), "YYYY-mm-dd") & "','00001','" & KEY_USUARIO & "','" & KEY_CAMBIO & "','" & dfac & "','" & formato_item(Month(Format(rst1("dVencimiento"), "YYYY-mm-dd")), 2) & "','" & Year(Format(rst1("dVencimiento"), "YYYY-mm-dd")) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
            
            id_venta = LastRegistroRUC("movimiento_venta", "id_venta")
128:
            Call SaveDetalleDocumentoVenta(id_venta, Trim(rst1("cDocumentoVenta")), Trim(rst1("sSerie")), rst1("doc_cod"))
            If Trim(rst1("Anulado")) = "V" Then
                On Error GoTo kardex1
                strCadena = "UPDATE movimiento_venta SET anulado='si' WHERE id_venta='" & id_venta & "' AND ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                 
kardex1:
                strCadena = "DELETE FROM kardex WHERE id_doc='" & formato_item(rst1("doc_cod"), 4) & "' AND id_serie='" & Trim(formato_item(rst1("sSerie"), 3)) & "' AND id_numero='" & Trim(formato_item(rst1("cDocumentoVenta"), 6)) & "' AND ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                 
            End If
siguiente:
            rst1.MoveNext
            DoEvents
       Next i
       MsgBox "MIGRACION EXITOSA" + Space(2) + "REGISTRO VENTAS", vbInformation, KEY_EMPRESA
    End If
    
 '-------documento Compras
'compras:
   Set rst1 = Nothing
 
   strCadena = "DELETE FROM movimiento_compra WHERE ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
    
   
    strCadena = "SELECT * FROM DocumentoCompra WHERE cPersona<>'00004' AND Anulado='F'  ORDER BY cDocumentoCompra"
    rst1.CursorLocation = adUseClient
    rst1.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
    If rst1.RecordCount > 0 Then
       rst1.MoveFirst
       
       
       
       For i = 0 To rst1.RecordCount - 1
           'dni = obtenerDNI(rst1("cPersona"))
           strCadena = "SELECT Per_Ruc FROM Persona WHERE cPersona='" & rst1("cPersona") & "'"
            rst4.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
            If rst4.RecordCount > 0 And rst4("Per_Ruc") <> "" Then
                If rst1("cPersona") = "00004" Then
                    dni = "00000000"
                Else
                    dni = rst4("Per_Ruc")
                End If
            Else
                dni = BDBuscarCampo("persona", "dni", "dni", formato_item(rst1("cPersona"), 8))
            End If
            Set rst4 = Nothing
            
           
            
        If Len(Trim(dni)) = 11 Then
            cod_identidad = 6
        End If
        If Len(Trim(dni)) <> 8 And Len(Trim(Me.txtRuc.Text)) <> 11 Then
            cod_identidad = 0
        End If
            
       'registro compras
            rmes = formato_item(Month(rst1("dEmisionCompra")), 2)
            ranio = formato_item(Year(rst1("dEmisionCompra")), 4)
            strCadena = "SELECT * FROM registro_compras WHERE ruc='" & KEY_RUC & "' AND mes='" & rmes & "' AND anio='" & ranio & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                desCompras = "REGISTRO COMPRAS :" + Space(3) + nombre_mes(rmes)
                strCadena = "INSERT INTO registro_compras(ruc,mes,anio,descripcion)VALUES('" & KEY_RUC & "','" & rmes & "','" & ranio & "','" & desCompras & "')"
                CnBd.Execute (strCadena)
                 
            End If
       'fin registro compras
       
       
        strCadena = "SELECT * FROM Docreferencia_Compra WHERE IdReferencia='" & rst1("IdReferencia") & "'" ' solo para don Daniel Olivos
        rst4.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
        If rst4.RecordCount > 0 Then
            nnumero = formato_item(rst4("cDocumentoCompra"), 8) ' aqui es para als demas bases de datos
            serie = formato_item(rst4("sSerie"), 3)
            doc_cod = formato_item(rst4("doc_cod"), 4)
        End If
        Set rst4 = Nothing
        strCadena = "P_insert_compra('" & doc_cod & "','" & KEY_ALM & "','" & Format(rst1("dEmisionCompra"), "YYYY-mm-dd") & "','" & Format(rst1("dVencimiento"), "YYYY-mm-dd") & "','02'," & _
        "'03','','00001','" & formato_item(Month(rst1("dEmisionCompra")), 2) & "','" & Year(rst1("dEmisionCompra")) & "','" & serie & "'," & _
        "'" & nnumero & "','" & cod_identidad & "','" & Trim(dni) & "','" & UCase(rst1("Persona")) & "','" & KEY_CAMBIO & "'," & _
        "'0','" & Val(rst1("nSubTotal")) & "','" & Val(rst1("nIgv")) & "','0','0','0','0','0','0','" & Val(rst1("nTotalCompra")) & "','" & Val(rst1("nTotalCompra")) & "','" & KEY_USUARIO & "','" & Trim(rst1("Observacion")) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        
        id_compra = LastRegistroRUC("movimiento_compra", "id_compra")

'02----------------guardar en detalle documento Compra-----------
        'Call SaveDetalleDocumentoCompra(id_compra)
        rst1.MoveNext
        DoEvents
       Next i
       
     Exit Sub
 'ACTUALIZAR STOCK PRODUCTO
inventario:

Dim strInventario As String
strCadena = "SELECT A.cProducto,A.Stock,P.Stock_factura,A.Alm_cod FROM Almacen_productos A,Producto P WHERE A.cProducto=P.cProducto"
rst3.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic

If rst3.RecordCount > 0 Then
rst3.MoveFirst
For i = 0 To rst3.RecordCount - 1
    cod_articulo = formato_item(rst3("cProducto"), 5)
    stock_actual = BDBuscarCampoRuc("almacen_producto", "stock", "id_producto", cod_articulo)
    stock_nuevo = rst3("Stock")
    Pcosto = BDBuscarCampoRuc("producto", "precio_compra", "id_producto", cod_articulo)
    
    strInventario = formato_item(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), 6)
    strCadena = "INSERT INTO inventario(id_inventario,id_producto,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES('" & strInventario & "','" & cod_articulo & "','" & Val(Pcosto) & "','" & KEY_FECHA & "','" & formato_item(rst3("Alm_cod"), 5) & "','" & Val(stock_nuevo) & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    strCadena = "UPDATE almacen_producto SET stock_factura='" & Val(rst3("Stock_factura")) & "' WHERE id_producto='" & cod_articulo & "' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    'strCadena = "UPDATE producto SET precio_venta='" & Val(Me.TxtVenta.text) & "',precio_compra='" & Val(Pcosto) & "' WHERE id_producto='" & cod_articulo & "' AND ruc='" & KEY_RUC & "'"
    'CnBd.Execute (strCadena)
    rst3.MoveNext
    DoEvents
Next i
  
  
End If
       
       
       
       MsgBox "MIGRACION EXITOSA" + Space(2) + "PERSONA RUC", vbInformation, KEY_EMPRESA
    End If
    
End Sub
Private Sub SaveDetalleDocumentoCompra(ByVal id_compra_nuevo As Double, ByVal in_compra_viejo As Double, ByVal in_fecha As String, ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String, ByVal in_in_proveedor As String, ByVal in_alm As String, ByVal in_tc As Single)
   strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & in_compra_viejo & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRst2(strCadena)
   If rst2.RecordCount > 0 Then
      rst2.MoveFirst
      For i = 0 To rst2.RecordCount - 1
            strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,fecha_vencimiento,numero_lote,ruc) VALUES ('" & id_compra_nuevo & "','" & rst2("id_producto") & "','" & rst2("cantidad") & "','" & rst2("c_unitario") & "'," & _
           "'" & rst2("dsto_soles") & "','" & rst2("dsto_procentaje") & "','" & rst2("total_descuento") & "','" & rst2("valor_neto") & "','" & rst2("isc") & "','" & rst2("igv") & "', " & _
           "'0','" & rst2("otros") & "','" & rst2("percepcion") & "','" & rst2("valor_venta") & "','" & rst2("exonerado") & "','" & rst2("total") & "','" & rst2("p_venta") & "','" & rst2("p_costo") & "','" & rst2("id_alm") & "','" & get_producto(rst2("id_producto"), KEY_RUC) & "','0','" & KEY_FECHA & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
           'If Me.DtcMoneda.BoundText = "00002" Then
           '     in_costo_unitario = (rst2("valor_venta") / rst2("cantidad")) * Val(in_tc)
           ' Else
                in_costo_unitario = (rst2("valor_venta") / rst2("cantidad"))
           ' End If
            strCadena = "call put_kardex_stock_vitekey('02','" & Format(in_fecha, "YYYY-mm-dd") & "','" & Val(id_compra_nuevo) & "','" & in_doc & "','" & Trim(in_serie) & "','" & Trim(in_numero) & "','" & Trim(in_in_proveedor) & "','" & rst2("id_producto") & "','" & rst2("cantidad") & "','" & in_costo_unitario & "','" & in_alm & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            rst2.MoveNext
      Next i
   End If
   
   
   
   
   
   
   
   
   
   
   
  ' strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,fecha_vencimiento,numero_lote,ruc) VALUES ('" & id_compra & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("c_unitario") & "'," & _
           "'" & rstT("dsto_soles") & "','" & rstT("dsto_procentaje") & "','" & rstT("total_descuento") & "','" & rstT("valor_neto") & "','" & rstT("isc") & "','" & rstT("igv") & "', " & _
           "'" & rstT("retencion") & "','" & rstT("otros") & "','" & rstT("percepcion") & "','" & rstT("valor_venta") & "','" & rstT("exonerado") & "','" & rstT("precio_venta") & "','" & rstT("p_venta") & "','" & rstT("p_costo") & "','" & rstT("id_alm") & "','" & rstT("detalle") & "','" & in_monto_parcial & "','" & Format(rstT("fecha_vencimiento"), "YYYY-mm-dd") & "','" & rstT("numero_lote") & "','" & KEY_RUC & "')"
   '        CnBd.Execute (strCadena)
           
    '      If get_servicio(rstT("id_producto")) = "no" Then
    '        If Me.DtcTipoDoc.BoundText <> "0419" Then
                'If chk_valor_venta.Value = 1 Then
                '   in_costo_unitario = rstT("c_unitario") * (1 + KEY_IGV)
                'Else
                '   in_costo_unitario = rstT("c_unitario")
                'End If
                
                   ' If Me.DtcMoneda.BoundText = "00002" Then
                  '     in_costo_unitario = (rstT("valor_venta") / rstT("cantidad")) * Val(Me.TxtTc.Text)
                 '   Else
                '    End If
                    
                               
                
                
               ' strCadena = "call put_kardex_stock_vitekey('02','" & Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd") & "','" & Val(id_compra) & "','" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.txtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.txtruc.Text) & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & in_costo_unitario & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
             '   CnBd.Execute (strCadena)
     '      End If
      '    End If
   
End Sub
Private Sub SaveDetalleDocumentoVenta(ByVal idVenta As Double, ByVal numero As String, serie As String, id_doc As String)
   Dim rst6 As New ADODB.Recordset
   Set rst6 = Nothing
   rst6.CursorLocation = adUseClient
   strCadena = "SELECT * FROM Detalle_DocumentoVenta WHERE (cDocumentoVenta='" & Trim(numero) & "' AND doc_cod='" & Trim(id_doc) & "' AND sSerie='" & Trim(serie) & "')"
   rst6.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
 
    If rst6.RecordCount > 0 Then
       rst6.MoveFirst
       For i = 0 To rst6.RecordCount - 1
           strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,cantidad,precio,peso,total,ruc) VALUES ('" & idVenta & "','" & formato_item(rst6("cProducto"), 5) & "','" & rst6("cantidad") & "','" & rst6("Precio") & "','" & rst6("Peso") & "','" & rst6("Total") & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
            
           rst6.MoveNext
        Next i
    End If
End Sub
Private Function obtenerDNI(ByVal cPersona As String) As String
obtenerDNI = BDBuscarCampo("persona", "dni", "cod_interno", cPersona)
End Function


Public Sub migrar_base(ByVal in_server As String, ByVal in_base As String)
Dim sys_ConString2 As String
sys_Server2 = in_server
sys_DataBase2 = in_base 'ConfigRead("DataBase")
sys_SUser2 = "user_vitekey" 'DecryptString(ConfigRead("SUser"))
sys_SPassword2 = "02021974abc2014@" 'DecryptString(ConfigRead("SPassword"))
db_port = "3306"
sys_ConString2 = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server2 & ";" & _
            "Database=" & sys_DataBase2 & ";" & _
            "UID=" & sys_SUser2 & ";" & _
            "PWD=" & sys_SPassword2 & ";" & _
            " PORT=" & db_port & ";OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
CnBd2.ConnectionString = sys_ConString2
CnBd2.Open

strCadena = "SELECT * FROM turno  WHERE ruc='20479779598'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
    
        strCadena = "SELECT * FROM turno  WHERE id_turno='" & rst("id_turno") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstP(strCadena)
        If rstP.RecordCount < 1 Then
        strCadena = "INSERT INTO turno(`id_turno`,`descripcion`,`hora_inicio`,`hora_final`,`ruc`)VALUES('" & rst("id_turno") & "','" & rst("descripcion") & "','" & Format(rst("hora_inicio"), "HH:mm:ss") & "','" & Format(rst("hora_final"), "HH:mm:ss") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        End If
        '
        rst.MoveNext
    Next i
End If





strCadena = "SELECT * FROM persona_cargos  WHERE id_empresa='20479779598'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM persona_cargos  WHERE id_cargo='" & rst("id_cargo") & "' and ruc='" & rst("ruc") & "' and id_empresa='" & KEY_RUC & "'"
        Call ConfiguraRstP(strCadena)
        If rstP.RecordCount < 1 Then
            strCadena = "INSERT INTO persona_cargos(`id_cargo`,`descripcion`,`ruc`,`id_empresa`)VALUES('" & rst("id_cargo") & "','" & rst("descripcion") & "','" & rst("ruc") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        End If
            rst.MoveNext
    Next i
End If

MsgBox "cargos Ingresados Correctamente"


strCadena = "SELECT * FROM unidad  WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
    rst2.MoveFirst
    For i = 0 To rst2.RecordCount - 1
         strCadena = "SELECT * FROM unidad  WHERE id_und='" & rst2("id_und") & "' and id_usu='" & KEY_RUC & "'"
        Call ConfiguraRstP(strCadena)
        If rstP.RecordCount < 1 Then
            strCadena = "INSERT INTO unidad(`id_und`,`abreviatura`,`descripcion`,`id_usu`)VALUES('" & rst2("id_und") & "','" & rst2("abreviatura") & "','" & rst2("descripcion") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
        End If
        rst2.MoveNext
    Next i
End If

MsgBox "Unidades Ingresados Correctamente"

strCadena = "SELECT * FROM linea  WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
    rst2.MoveFirst
    For i = 0 To rst2.RecordCount - 1
        strCadena = "SELECT * FROM linea  WHERE id_linea='" & rst2("id_linea") & "' and id_usu='" & KEY_RUC & "'"
        Call ConfiguraRstP(strCadena)
        If rstP.RecordCount < 1 Then
            strCadena = "INSERT INTO linea(`id_linea`,`descripcion`,`id_usu`)VALUES('" & rst2("id_linea") & "','" & rst2("descripcion") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        End If
        rst2.MoveNext
    Next i
End If

MsgBox "Linea Ingresados Correctamente"

strCadena = "SELECT * FROM marca  WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
    rst2.MoveFirst
    For i = 0 To rst2.RecordCount - 1
        strCadena = "SELECT * FROM marca  WHERE id_marca='" & rst2("id_marca") & "' and id_usu='" & KEY_RUC & "'"
        Call ConfiguraRstP(strCadena)
        If rstP.RecordCount < 1 Then
                    strCadena = "INSERT INTO marca(`id_marca`,`descripcion`,`id_usu`)VALUES('" & rst2("id_marca") & "','" & rst2("descripcion") & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                     
         End If
        rst2.MoveNext
    Next i
End If

MsgBox "Marcas Ingresados Correctamente"


strCadena = "SELECT * FROM linea  WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
    rst2.MoveFirst
      For i = 0 To rst2.RecordCount - 1
         strCadena = "SELECT * FROM linea_sub  WHERE id_tipo='" & Format(i + 1, "00000") & "' and id_linea='" & rst2("id_linea") & "' and id_usu='" & KEY_RUC & "'"
        Call ConfiguraRstP(strCadena)
        If rstP.RecordCount < 1 Then
                    strCadena = "INSERT INTO linea_sub(`id_tipo`,`id_linea`,`descripcion`,`id_usu`)VALUES('" & Format(i + 1, "00000") & "','" & rst2("id_linea") & "','SIN MODELO','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                     
        End If
        rst2.MoveNext
      Next i
End If

MsgBox "sub-lineas Ingresados Correctamente"


strCadena = "SELECT p.* FROM entidad_empresa e,persona p WHERE e.cod_unico=p.dni and e.id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
   strCadena = "SELECT * FROM persona WHERE dni='" & rst2("dni") & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount < 1 Then
      strCadena = "INSERT INTO persona(`dni`,`a_paterno`,`a_materno`,`nombres`,`nombre_completo`,`direccion`,`id_dia`,`id_mes`,`id_anio`,`sexo`,`peso`, " & _
      "`estatura`,`estado_civil`,`mail`,`id_pais`,`celular`,`fecha_ingreso`,`id_departamento`,`id_provincia`,`id_distrito`,`id_urbanizacion`,`id_zona`,`grupo_votacion`," & _
      "`id_religion`,`foto`,`fecha_registro`, " & _
      "`licencia`" & _
      ")VALUES(" & _
      "'" & rst2("dni") & "','" & rst2("a_paterno") & "','" & rst2("a_materno") & "','" & rst2("nombres") & "','" & rst2("nombre_completo") & "','" & rst2("direccion") & "', " & _
      "'" & rst2("id_dia") & "','" & rst2("id_mes") & "','" & rst2("id_anio") & "','" & rst2("sexo") & "','" & rst2("peso") & "','" & rst2("estatura") & "','" & rst2("estado_civil") & "' " & _
      ",'" & rst2("mail") & "','" & rst2("id_pais") & "','" & rst2("celular") & "','" & Format(rst2("fecha_ingreso"), "YYYY-mm-dd") & "','" & rst2("id_departamento") & "','" & rst2("id_provincia") & "','" & rst2("id_distrito") & "'," & _
      "'" & rst2("id_urbanizacion") & "','" & rst2("id_zona") & "','" & rst2("grupo_votacion") & "','" & rst2("id_religion") & "','" & rst2("foto") & "','" & Format(rst2("fecha_registro"), "YYYY-mm-dd") & "','" & rst2("licencia") & "' " & _
      ")"
      CnBd.Execute (strCadena)
       
   End If
   rst2.MoveNext
   DoEvents
   Next i
   
End If

MsgBox "persona Ingresados Correctamente"

strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico<>'00000000' and id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
    rst2.MoveFirst
    For i = 0 To rst2.RecordCount - 1
        strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & rst2("cod_unico") & "' and id_empresa='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            If IsNull(rst2("fecha_ingreso")) = True Then
                in_fecha_ingreso = KEY_FECHA
            Else
                in_fecha_ingreso = rst2("fecha_ingreso")
            End If
            strCadena = "INSERT INTO entidad_empresa(`cod_unico`,`id_empresa`,`id_tipo_per`,`id_cargo`,`id_cliente`,`id_personal`,`id_transporte`,`id_proveedor`,`id_contable`,  " & _
            "`id_auspeciador`,`id_almacen`,`id_area`,`id_condicion`,`password`,`passwordaccesso`,`habilitado`,`fecha_ingreso`," & _
            "`id_moneda`,`sueldo`,`id_planilla`,`id_afp`,`rta_quinta`,`asig_familiar`,`bonificacion_extraordinaria`,`cuspp`,`essalud`,`sndp`,`id_sucursal`," & _
            " `id_credito`,`monto_credito`,`id_estado`,`observacion`)VALUES('" & rst2("cod_unico") & "','" & KEY_RUC & "','" & rst2("id_tipo_per") & "','" & rst2("id_cargo") & "','" & rst2("id_cliente") & "','" & rst2("id_personal") & "'," & _
            "'" & rst2("id_transporte") & "','" & rst2("id_proveedor") & "','" & rst2("id_contable") & "','" & rst2("id_auspeciador") & "','" & rst2("id_almacen") & "','" & rst2("id_area") & "','" & rst2("id_condicion") & "','" & rst2("password") & "','" & rst2("passwordaccesso") & "'," & _
            "'" & rst2("habilitado") & "','" & Format(in_fecha_ingreso, "YYYY-mm-dd") & "','" & rst2("id_moneda") & "','" & rst2("sueldo") & "','" & rst2("id_planilla") & "','" & rst2("id_afp") & "','" & rst2("rta_quinta") & "','" & rst2("asig_familiar") & "','" & rst2("bonificacion_extraordinaria") & "', " & _
            "'" & rst2("cuspp") & "','" & rst2("essalud") & "','" & rst2("sndp") & "','" & rst2("id_sucursal") & "','" & rst2("id_credito") & "','" & rst2("monto_credito") & "','" & rst2("id_estado") & "','" & rst2("observacion") & "')"
            CnBd.Execute (strCadena)
             
        End If
        rst2.MoveNext
        DoEvents
    Next i
End If

MsgBox "entidad empresa Ingresados Correctamente"
strCadena = "SELECT * FROM almacen WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "SELECT * FROM almacen WHERE id_alm='" & rst2("id_alm") & "' and ruc='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount < 1 Then
          strCadena = "INSERT INTO almacen(`id_alm`,`ruc`,`descripcion`,`abreviatura`,`direccion`,`stock`,`defecto`,`activo`,`dni_save`)VALUES " & _
          "('" & rst2("id_alm") & "','" & KEY_RUC & "','" & rst2("descripcion") & "','" & rst2("descripcion") & "','" & rst2("direccion") & "','si','si','si','" & KEY_USUARIO & "')"
          CnBd.Execute (strCadena)
       End If
       rst2.MoveNext
   Next i
End If
MsgBox "Almacen Ingresado Correctamente"

strCadena = "DELETE FROM producto WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM producto WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
    rst2.MoveFirst
    For i = 0 To rst2.RecordCount - 1
        strCadena = "SELECT * FROM marca WHERE id_marca='" & rst2("id_marca") & "' and id_usu='" & KEY_RUC & "'"
        Call ConfiguraRstK(strCadena)
        
        If rst2("id_marca") = "" Then
           in_marca = "00001"
           nmarca = ""
        Else
           in_marca = rst2("id_marca")
           nmarca = rstK("descripcion")
        End If
        strCadena = "SELECT * FROM producto WHERE id_producto='" & rst2("id_producto") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstP(strCadena)
        If rstP.RecordCount < 1 Then
        strCadena = "INSERT INTO producto(`id_producto`,`id_proveedor`,`id_auspiciador`,`id_linea`,`id_moneda`," & _
        "`nombre_prod`,`id_unidad`,`stock_minimo`,`peso`,`imagen`,`nombre_comercial`,`id_marca`,`marca`,`comentario`,`id_percepcion`,`id_combo`,`id_igv`,`dni_save`,`ruc`)VALUES(" & _
        "'" & rst2("id_producto") & "','" & rst2("id_proveedor") & "','" & rst2("id_auspiciador") & "','" & rst2("id_linea") & "','" & rst2("id_moneda") & "','" & rst2("nombre_prod") & "','" & rst2("id_unidad") & "'," & _
        "'" & rst2("stock_minimo") & "','" & rst2("peso") & "','" & rst2("imagen") & "','" & rst2("nombre_comercial") & "','" & in_marca & "','" & nmarca & "','" & rst2("comentario") & "','" & rst2("id_percepcion") & "','" & rst2("id_combo") & "','" & rst2("id_igv") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        
        strCadena = "SELECT * FROM almacen WHERE ruc='" & KEY_RUC & "'"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            rstK.MoveFirst
            For j = 0 To rstK.RecordCount - 1
            
                strCadena = "INSERT INTO almacen_producto(`id_alm`,`id_producto`,`ruc`,`precio_venta`," & _
                "`precio_compra`)VALUES('" & rstK("id_alm") & "','" & rst2("id_producto") & "'," & _
                "'" & KEY_RUC & "','" & rst2("precio_venta") & "','" & rst2("precio_compra") & "')"
                CnBd.Execute (strCadena)
                 
                
                
                rstK.MoveNext
            Next j
        End If
        End If
        rst2.MoveNext
        
    Next i
End If



strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "SELECT * FROM almacen_comprobante WHERE id_alm_com='" & rst2("id_alm_com") & "' and ruc='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount < 1 Then
          strCadena = "INSERT INTO almacen_comprobante(`id_alm_com`,`ruc`,`id_alm`,`id_doc`,`serie`,`numero`,`igv`,`defecto`,`id_moneda`,`venta`,`id_formato_impresion`)VALUES " & _
          "('" & rst2("id_alm_com") & "','" & KEY_RUC & "','" & rst2("id_alm") & "','" & rst2("id_doc") & "','" & rst2("serie") & "','" & rst2("numero") & "','" & rst2("igv") & "','" & rst2("defecto") & "','" & rst2("id_moneda") & "','" & rst2("venta") & "','" & rst2("id_formato_impresion") & "')"
          CnBd.Execute (strCadena)
       End If
       rst2.MoveNext
   Next i
End If

MsgBox "Almacen Ingresado Correctamente"




strCadena = "SELECT * FROM movimiento_venta WHERE fecha_emision>='2015-01-01' and  ruc='" & KEY_RUC & "' ORDER BY id_venta ASC"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   Me.ProgressBar1.Min = 0
   Me.ProgressBar1.Max = rst2.RecordCount
   For i = 0 To rst2.RecordCount - 1
            
            strCadena = "SELECT * FROM comprobantes WHERE id_doc='" & rst2("id_doc") & "'"
            Call ConfiguraRstP(strCadena)
            If rstP.RecordCount > 0 Then
                in_doc = rstP("doc_abrev") & ":" & rst2("serie") & "-" & rst2("numero")
            Else
                X = 1
            End If
            
            
            id_tipo_factura = "00001"
            strCadena = "P_insert_venta_v2('" & rst2("id_doc") & "','" & rst2("id_alm") & "','" & rst2("id_forma_pago") & "','" & rst2("id_moneda") & "','" & rst2("id_delivery") & "'," & _
            "'" & Trim(rst2("serie")) & "','" & Trim(rst2("numero")) & "','" & rst2("id_cliente") & "','" & rst2("ncliente") & "','" & rst2("valor_venta") & "','" & rst2("igv") & "','" & rst2("exonerado") & "','" & rst2("total") & "','" & rst2("saldo") & "'," & _
            "'" & Val(rst2("total")) & "','" & Val(rst2("monto_vuelto")) & "','" & Format(rst2("fecha_emision"), "YYYY-mm-dd") & "','" & Format(rst2("fecha_emision"), "YYYY-mm-dd") & "','" & id_tipo_factura & "','" & rst2("id_vendedor") & "','" & KEY_USUARIO & "','" & rst2("tc") & "','" & rst2("afecta_factura") & "','" & formato_item(Month(rst2("fecha_emision")), 2) & "','" & Year(rst2("fecha_emision")) & "','" & in_doc & "',CURTIME(),'T','--','no','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
            id_venta = LastRegistroRUC("movimiento_venta", "id_venta")
            
            strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & rst2("id_venta") & "' and ruc='" & KEY_RUC & "' ORDER BY id_detalle_venta"
            Call ConfiguraRstCloud(strCadena)
            If rstCloud.RecordCount > 0 Then
                rstCloud.MoveFirst
                For j = 0 To rstCloud.RecordCount - 1
                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,cantidad,precio,peso,total,ruc) VALUES ('" & id_venta & "','" & rstCloud("id_producto") & "','" & get_producto(rstCloud("id_producto"), KEY_RUC) & "','" & rstCloud("cantidad") & "','" & rstCloud("precio") & "','" & rstCloud("peso") & "','" & rstCloud("total") & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    rstCloud.MoveNext
                Next j
            End If
            strCadena = "INSERT INTO movimiento_venta_monto(id_venta,id_forma_pago,monto,monto_caja,id_tarjeta,ruc)VALUES('" & id_venta & "','01','" & rst2("total") & "','" & rst2("total") & "','00','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
            rst2.MoveNext
            Me.ProgressBar1.Value = i
            DoEvents
    Next i
End If


strCadena = "SELECT * FROM movimiento_compra WHERE ruc='" & KEY_RUC & "' and fecha_emision>='2015-01-01' ORDER BY id_compra ASC"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   Me.ProgressBar1.Min = 0
   Me.ProgressBar1.Max = rst2.RecordCount
   For i = 0 To rst2.RecordCount - 1
           
            strCadena = "P_insert_compra('" & rst2("id_doc") & "','" & rst2("id_alm") & "','" & Format(rst2("fecha_emision"), "YYYY-mm-dd") & "','" & Format(rst2("fecha_cancelacion"), "YYYY-mm-dd") & "','02'," & _
        "'" & rst2("id_tipo_compra") & "','','" & rst2("id_moneda") & "','" & formato_item(Month(rst2("fecha_emision")), 2) & "','" & Year(rst2("fecha_emision")) & "','" & Trim(rst2("serie")) & "'," & _
        "'" & Format(Trim(rst2("numero")), "00000000") & "','" & rst2("tipo_doc_identidad") & "','" & Trim(rst2("id_proveedor")) & "','" & rst2("nproveedor") & "','" & rst2("tc") & "'," & _
        "'0','" & rst2("valor_venta") & "','" & rst2("igv") & "','" & rst2("isc") & "','0','" & rst2("percepcion") & "','0','" & rst2("exonerado") & "','0','" & rst2("total") & "','" & rst2("saldo") & "','" & KEY_USUARIO & "','-','" & KEY_RUC & "')"
        
        CnBd.Execute (strCadena)
            
            id_compra = LastRegistroRUC("movimiento_compra", "id_compra")
            
            strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & rst2("id_compra") & "' ORDER BY id_detalle_compra ASC"
            Call ConfiguraRstCloud(strCadena)
            If rstCloud.RecordCount > 0 Then
            rstCloud.MoveFirst
                For p = 0 To rstCloud.RecordCount - 1
                    strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,ivap,otros,percepcion, " & _
                    "valor_venta,exonerado,total,p_venta,p_costo,id_alm,ruc) VALUES ('" & id_compra & "','" & rstCloud("id_producto") & "','" & rstCloud("cantidad") & "','" & rstCloud("c_unitario") & "'," & _
                    "'" & rstCloud("dsto_soles") & "','" & rstCloud("dsto_procentaje") & "','" & rstCloud("total_descuento") & "','" & rstCloud("valor_neto") & "','" & rstCloud("isc") & "','" & rstCloud("igv") & "', " & _
                    "'" & rstCloud("ivap") & "','" & rstCloud("otros") & "','" & rstCloud("percepcion") & "','" & rstCloud("valor_venta") & "','" & rstCloud("exonerado") & "','" & rstCloud("p_venta") & "','" & rstCloud("p_venta") & "','" & rstCloud("p_costo") & "','" & rstCloud("id_alm") & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    rstCloud.MoveNext
                Next p
    End If
            Me.ProgressBar1.Value = i
            rst2.MoveNext
            DoEvents
    Next i
End If


strCadena = "SELECT * FROM almacen_producto a,producto p WHERE  a.id_producto=p.id_producto and a.id_alm='00001' and a.ruc=p.ruc and  a.ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
    rst2.MoveFirst
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = rst2.RecordCount
    For i = 0 To rst2.RecordCount - 1
        cod_articulo = rst2("id_producto")
        
            strInventario = formato_item(LastRegistroCloud("inventario", "id_inventario"), 6)
            strCadena = "INSERT INTO inventario(id_inventario,id_producto,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
            "('" & strInventario & "','" & cod_articulo & "','" & KEY_FECHA & "','" & rst2("id_alm") & "','" & rst2("stock") & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
            
    
            strCadena = "UPDATE almacen_producto SET stock_factura ='" & rst2("stock_factura") & "',precio_venta='" & rst2("precio_venta") & "',precio_compra='" & rst2("precio_compra") & "' WHERE id_producto='" & Trim(cod_articulo) & "' AND id_alm='" & rst2("id_alm") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            rst2.MoveNext
            Me.ProgressBar1.Value = i
            DoEvents
    Next i
End If



MsgBox "Listo"








End Sub

Public Function get_producto(ByVal in_producto As String, ByVal in_ruc As String)
strCadena = "SELECT * FROM producto where id_producto='" & in_producto & "' and ruc='" & in_ruc & "'"
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
    get_producto = rstAux("nombre_prod")
Else
    get_producto = "--"
End If
End Function

Private Sub cmdimportarTabla1_Click()
Call migrar_base(Trim(Me.txtserver1.Text), Trim(Me.txtbaseOrigen1.Text))
End Sub

Private Sub cmdLoadvitekey_Click()
Dim codigo As Integer
Dim habilitado As String


If chk_sucursal.Value = 1 Then
   strCadena = "DELETE FROM inventario_inicial WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and  ruc='" & KEY_RUC & "'"
Else
   strCadena = "DELETE FROM inventario_inicial WHERE   ruc='" & KEY_RUC & "'"
End If
CnBd.Execute (strCadena)


Dim in_cod As String
'On Error GoTo salir
For i = 0 To Me.hfproductos.Rows - 1
        'in_caracter = "3"
        
        If Len(Trim(Me.hfproductos.TextMatrix(i, 1))) > 1 Then
            in_producto = Format(Trim(Me.hfproductos.TextMatrix(i, 0)), "00000") & in_caracter
        If in_producto = "CODIGO" Then
            GoTo abc
        End If
        
        
        If Val(Me.hfproductos.TextMatrix(i, 0)) = 0 Then
        
           codigo = i + 1
           in_producto = Format(codigo, "00000") '& in_caracter
        Else
            If Val(Trim(Me.hfproductos.TextMatrix(i, 0))) > 0 Then
                in_producto = Format(Trim(Me.hfproductos.TextMatrix(i, 0)), "00000")
            Else
                in_producto = Trim(Me.hfproductos.TextMatrix(i, 0))
            End If
            
        End If
        in_cod = Trim(Me.hfproductos.TextMatrix(i, 0))
        
        in_unidad_abrev = UCase(Trim(Me.hfproductos.TextMatrix(i, 7)))
        in_unidad = UCase(Trim(Me.hfproductos.TextMatrix(i, 8)))
        in_unidad_compra = UCase(Trim(Me.hfproductos.TextMatrix(i, 10)))
        
        
        nombre_producto = Trim(Replace(UCase(Trim(Me.hfproductos.TextMatrix(i, 1))), "'", " "))
        nombre_comercial = Trim(Replace(UCase(Trim(Me.hfproductos.TextMatrix(i, 2))), "'", " "))
        If Trim(Replace(UCase(Trim(Me.hfproductos.TextMatrix(i, 3))), "'", " ")) = "HABILITADO" Then
            habilitado = "si"
        Else
            habilitado = "no"
        End If
        in_linea = UCase(Trim(Me.hfproductos.TextMatrix(i, 4)))
        in_sublinea = Trim(Me.hfproductos.TextMatrix(i, 5))
        in_modelo = Trim(Me.hfproductos.TextMatrix(i, 6))
        in_unidad = UCase(Trim(Me.hfproductos.TextMatrix(i, 8)))
        in_unidad_trae = Val(Me.hfproductos.TextMatrix(i, 9))
        
        
               
        in_marca = Trim(Me.hfproductos.TextMatrix(i, 11))
        in_color = Trim(Me.hfproductos.TextMatrix(i, 12))
        in_precio_costo = Val(Format(Me.hfproductos.TextMatrix(i, 13), "###0.00"))
        in_precio_venta = Val(Format(Me.hfproductos.TextMatrix(i, 14), "###0.00"))
        
        
        in_precio_a = Val(Format(Me.hfproductos.TextMatrix(i, 15), "###0.00"))
        in_precio_mayor = Val(Format(Me.hfproductos.TextMatrix(i, 16), "###0.00"))
        in_genero = UCase(Me.hfproductos.TextMatrix(i, 17))
        in_talla = UCase(Me.hfproductos.TextMatrix(i, 18))
        in_stock_contable = (Me.hfproductos.TextMatrix(i, 19))
        in_stock_nocontable = (Me.hfproductos.TextMatrix(i, 20))
        in_peso = Val(Me.hfproductos.TextMatrix(i, 21))
        
        If chk_sucursal.Value = 1 Then
            in_alm = Me.DtcAlmacen.BoundText
        Else
            in_alm = Format(Me.hfproductos.TextMatrix(i, 20), "00000")
        End If
        codigo_universal = Replace(UCase(Trim(Me.hfproductos.TextMatrix(i, 22))), "'", " ")
        codigo_barra = Format(Replace(UCase(Trim(Me.hfproductos.TextMatrix(i, 23))), "'", " "), "00000000000")
        codigo_proveedor = Format(Replace(UCase(Trim(Me.hfproductos.TextMatrix(i, 25))), "'", " "), "00000000000")
        razon_social = Trim(Replace(UCase(Trim(Me.hfproductos.TextMatrix(i, 26))), "'", " "))
        direccion_fiscal = Trim(Replace(UCase(Trim(Me.hfproductos.TextMatrix(i, 27))), "'", " "))
        
        
        
        
        strCadena = "INSERT INTO inventario_inicial " & _
        "(`id_producto`,`nombre_producto`,nombre_comercial,`linea`,sub_linea,`modelo`,unidad_abrev,`unidad`,unidad_compra,`marca`,`color`,`precio_costo`,`precio_venta`," & _
        "`precio_mayor`,precio_alterno_a,`genero`,`talla`,`stock_fisico`,`stock_contable`,`peso`,`id_alm`,codigo_universal,codigo_proveedor,razon_social," & _
        " direccion_proveedor,codigo_barra,cantidad_trae,habilitado,`ruc`)VALUES" & _
        "('" & in_producto & "','" & nombre_producto & "','" & nombre_comercial & "','" & in_linea & "','" & in_sublinea & "','" & in_modelo & "','" & in_unidad_abrev & "','" & in_unidad & "','" & in_unidad_compra & "'," & _
        " '" & in_marca & "','" & in_color & "','" & in_precio_costo & "','" & in_precio_venta & "','" & in_precio_mayor & "','" & in_precio_a & "','" & in_genero & "','" & in_talla & "','" & in_stock_contable & "', " & _
        "'" & in_stock_nocontable & "','" & in_peso & "','" & in_alm & "','" & codigo_universal & "','" & Trim(codigo_proveedor) & "','" & razon_social & "','" & direccion_fiscal & "','" & codigo_barra & "'," & _
        " '" & Val(in_unidad_trae) & "','" & habilitado & "','" & Trim(Me.txtRuc.Text) & "')"
        CnBd.Execute (strCadena)
       
       
       
        
       
        
ns:
        
        DoEvents
        Me.cmdLoadvitekey.Caption = str(i)
    End If
abc:
Next i
'salir:
Call llenar_productos(Me.hfproductos)


End Sub
Public Sub llenar_productos(ByVal Grilla As MSHFlexGrid)


If chk_sucursal.Value = 1 Then
    strCadena = "SELECT * FROM inventario_inicial WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and  ruc='" & Trim(Me.txtRuc.Text) & "'"
Else
    strCadena = "SELECT * FROM inventario_inicial WHERE ruc='" & Trim(Me.txtRuc.Text) & "'"
End If

Call ConfiguraRstT(strCadena)


If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
  
    Exit Sub
End If
  
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
        For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 1200 'codigo
           Grilla.ColWidth(1) = 3500 'producto
           Grilla.ColWidth(2) = 2000 'linea
           Grilla.ColWidth(3) = 2000 'modelo
           Grilla.ColWidth(4) = 1200 'unidad
           Grilla.ColWidth(5) = 1000 'marca
           Grilla.ColWidth(6) = 1000 'color
           Grilla.ColWidth(7) = 1200 'costo
           Grilla.ColWidth(8) = 1200 'venta
           Grilla.ColWidth(9) = 1200 'mayor
           Grilla.ColWidth(10) = 1000 'genero
           Grilla.ColWidth(11) = 600 'talla
           Grilla.ColWidth(12) = 1000 'stock contable
           Grilla.ColWidth(13) = 1000 'stock no contable
           Grilla.ColWidth(14) = 1200 'peso
           Grilla.ColWidth(15) = 700 'almacen
        Next
        cabecera = "CODIGO" & vbTab & "PRODUCTO" & vbTab & "LINEA" & vbTab & "MODELO" & vbTab & "UNIDAD" & vbTab & "MARCA" & vbTab & "COLOR" & vbTab & "COSTO" & vbTab & "VENTA" & vbTab & "MAYOR" & vbTab & "GENERO" & vbTab & "TALLA" & vbTab & "S.CONTABLE" & vbTab & "S.NOCANTABLE" & vbTab & "PESO" & vbTab & "ALMACEN"
        Grilla.AddItem cabecera
         For k = 0 To 15
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
       

            Fila = rstT("id_producto") & vbTab & rstT("nombre_producto") & vbTab & rstT("linea") & vbTab & rstT("modelo") & vbTab & rstT("unidad") & vbTab & rstT("marca") & vbTab & rstT("color") & vbTab & rstT("precio_costo") & vbTab & rstT("precio_venta") & vbTab & rstT("precio_mayor") & vbTab & rstT("genero") & vbTab & rstT("talla") & vbTab & rstT("stock_fisico") & vbTab & rstT("stock_contable") & vbTab & rstT("peso") & vbTab & rstT("id_alm")
            Grilla.AddItem Fila
            rstT.MoveNext
      Next i
      Me.Command12.Enabled = True
    
     
End Sub

Private Sub cmdmigrar_Click()
Dim sys_ConString2 As String
Dim stock_actual As Integer

sys_Server2 = "52.33.74.33"
sys_DataBase2 = "bd_vitekey_repos_ii"   'ConfigRead("DataBase")
sys_SUser2 = "user_cord" 'DecryptString(ConfigRead("SUser"))
sys_SPassword2 = "123456" 'DecryptString(ConfigRead("SPassword"))
db_port = "3306"
sys_ConString2 = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server2 & ";" & _
            "Database=" & sys_DataBase2 & ";" & _
            "UID=" & sys_SUser2 & ";" & _
            "PWD=" & sys_SPassword2 & ";" & _
            " PORT=" & db_port & ";OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
CnBd2.ConnectionString = sys_ConString2
CnBd2.Open


strCadena = "SELECT * FROM view_entidad WHERE id_cliente='si' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
        strCadena = "SELECT * FROM persona WHERE dni='" & Trim(rst2("dni")) & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                strCadena = "call P_insert_persona_ii('" & rst2("dni") & "' " & _
                ",'" & rst2("a_paterno") & "', " & _
                "'" & rst2("a_materno") & "' " & _
                ",'" & rst2("nombres") & "' " & _
                ",'" & rst2("nombre_completo") & "' " & _
                ",'" & rst2("direccion") & "' " & _
                ",'" & rst2("celular") & "' " & _
                ",'" & rst2("mail") & "'" & _
                ",'no' " & _
                ",'no'" & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                strCadena = "UPDATE persona SET id_pais='" & rst2("id_pais") & "', peso='" & Val(rst2("peso")) & "',estatura='" & rst2("estatura") & "',id_dia='" & Trim(rst2("id_dia")) & "',id_mes='" & Trim(rst2("id_mes")) & "',id_anio='" & Trim(rst2("id_anio")) & "'," & _
                "id_departamento='" & rst2("id_departamento") & "',id_provincia='" & rst2("id_provincia") & "',id_distrito='" & rst2("id_distrito") & "' WHERE dni='" & Trim(rst2("dni")) & "'"
                CnBd.Execute (strCadena)
                
            End If
                
                Call put_estudiante(Trim(rst2("dni")))
                strCadena = "SELECT * FROM college_matricula where dni='" & rst2("dni") & "'  and ruc='" & KEY_RUC & "'"
                Call ConfiguraRstCloud(strCadena)
                If rstCloud.RecordCount > 0 Then
                   strCadena = "SELECT * FROM college_servicio_persona WHERE dni='" & rst2("dni") & "' LIMIT 1 "
                   Call ConfiguraRst3(strCadena)
                   Call put_matricula(rstCloud("id_periodo"), rstCloud("id_nivel"), rstCloud("id_grado"), Trim(rst2("dni")), rst3("id_servicio"))
                End If
                
                strCadena = "SELECT * FROM persona_accidentes WHERE dni='" & rst2("dni") & "' "
                Call ConfiguraRstCloud(strCadena)
                If rstCloud.RecordCount > 0 Then
                    rstCloud.MoveFirst
                    For k = 0 To rstCloud.RecordCount - 1
                        strCadena = "SELECT * FROM persona_accidentes WHERE dni_familia='" & rstCloud("dni_familia") & "' and dni='" & rst2("dni") & "' "
                        Call ConfiguraRstK(strCadena)
                        If rstK.RecordCount < 1 Then
                           strCadena = "INSERT INTO persona_accidentes(dni,dni_familia,id_parentesco,telefono,direccion,id_ocupacion,id_grado)VALUES('" & rst2("dni") & "','" & rstCloud("dni_familia") & "','" & rstCloud("id_parentesco") & "','" & rstCloud("telefono") & "','" & rstCloud("direccion") & "','" & rstCloud("id_ocupacion") & "','" & rstCloud("id_grado") & "')"
                           CnBd.Execute (strCadena)
                           
                           strCadena = "SELECT * FROM  persona where dni='" & Trim(rstCloud("dni_familia")) & "'"
                           Call ConfiguraRstP(strCadena)
                           If rstP.RecordCount < 1 Then
                              strCadena = "SELECT * FROM persona WHERE dni='" & Trim(rstCloud("dni_familia")) & "'"
                              Call ConfiguraRst3(strCadena)
                              If rst3.RecordCount > 0 Then
                                strCadena = "P_insert_persona('" & rstCloud("dni_familia") & "','" & rst3("a_paterno") & "','" & rst3("a_materno") & "','" & rst3("nombres") & "','" & rst3("nombre_completo") & "','" & rst3("direccion") & "','" & rst3("celular") & "','--','no','no','no','no','no','0','')"
                                CnBd.Execute (strCadena)
                              End If
                            End If
                        End If
                        rstCloud.MoveNext
                    Next k
                End If
                
                
                
          DoEvents
           
        rst2.MoveNext
   Next i
End If
'GoTo 100
 
'MIGRAR UNIDADES DE MEDIDA

'strCadena = "SELECT * FROM almacen_producto where ruc='" & KEY_RUC & "' and id_alm='00001' ORDER BY id_producto ASC"
'Call ConfiguraRst2(strCadena)
'If rst2.RecordCount > 0 Then
    'rst2.MoveFirst
   ' For i = 0 To rst2.RecordCount - 1
   '     Call get_stock(rst2("id_producto"), rst2("stock"), KEY_FECHA)
   '     rst2.MoveNext
   '     DoEvents
  '  Next i
'End If


'strCadena = "SELECT id_producto,d.cantidad,v.fecha_emision FROM movimiento_venta v, movimiento_venta_detalle d WHERE   v.id_venta=d.id_venta and v.ruc='" & KEY_RUC & "' and v.fecha_emision>='2017-01-03'"
'Call ConfiguraRst(strCadena)
'If rst.RecordCount > 0 Then
'    rst.MoveFirst
'    For i = 0 To rst.RecordCount - 1
'        Call get_update_stock(rst("id_producto"), rst("cantidad"), rst("fecha_emision"))
'        rst.MoveNext
'        DoEvents
'    Next i
'End If

Exit Sub



strCadena = "DELETE FROM unidad WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM unidad WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   Me.ProgressBar2.Min = 0
   Me.ProgressBar2.Max = rst2.RecordCount
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       
          strCadena = "INSERT INTO unidad(`id_und`,`abreviatura`,`descripcion`,`id_usu`)VALUES('" & rst2("id_und") & "','" & rst2("abreviatura") & "','" & rst2("descripcion") & "','" & KEY_RUC & "')"
          CnBd.Execute (strCadena)
       
       rst2.MoveNext
       DoEvents
       Me.ProgressBar2.Value = i
   Next i
   
End If



'MIGRAR LINEAS
strCadena = "DELETE FROM linea WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM linea WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   Me.ProgressBar2.Min = 0
   Me.ProgressBar2.Max = rst2.RecordCount
   rst2.MoveFirst
   
   For i = 0 To rst2.RecordCount - 1
       strCadena = "SELECT * FROM linea WHERE id_linea='" & rst2("id_linea") & "' and id_usu='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount < 1 Then
          strCadena = "INSERT INTO linea(`id_linea`,`descripcion`,`id_tipo`,`planilla`,`produccion`,`id_usu`)VALUES " & _
          "('" & rst2("id_linea") & "','" & rst2("descripcion") & "','" & rst2("id_tipo") & "','" & rst2("planilla") & "','" & rst2("planilla") & "','" & KEY_RUC & "')"
          CnBd.Execute (strCadena)
       End If
       rst2.MoveNext
   Next i
   DoEvents
   Me.ProgressBar2.Value = i
End If


'MIGRAR SUB LINEA SUB

strCadena = "DELETE FROM linea_sub WHERE id_usu='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM linea_sub WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   Me.ProgressBar2.Min = 0
   Me.ProgressBar2.Max = rst2.RecordCount
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "SELECT * FROM linea_sub WHERE id_tipo='" & rst2("id_tipo") & "' and id_usu='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount < 1 Then
          strCadena = "INSERT INTO linea_sub(`id_tipo`,`id_linea`,`descripcion`,`id_usu`)VALUES " & _
          "('" & rst2("id_tipo") & "','" & rst2("id_linea") & "','" & rst2("descripcion") & "','" & KEY_RUC & "')"
          CnBd.Execute (strCadena)
       End If
       rst2.MoveNext
   Next i
   DoEvents
   Me.ProgressBar2.Value = i
End If

'MIGRAR SUB ALMACEN

strCadena = "DELETE FROM almacen WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM almacen WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   Me.ProgressBar2.Min = 0
   Me.ProgressBar2.Max = rst2.RecordCount
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       strCadena = "SELECT * FROM almacen WHERE ruc='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount < 1 Then
          strCadena = "INSERT INTO almacen(`id_alm`,`ruc`,`descripcion`,`direccion`,`id_responsable`,`stock`,`stock_personalizado`,`hora_inicio`,`hora_fin`,`horas`,`defecto`,`activo`,`id_sucursal`,`facturacion_detallada`,`caja_independiente`)VALUES " & _
          "('" & rst2("id_alm") & "','" & KEY_RUC & "','" & rst2("descripcion") & "','" & rst2("direccion") & "','" & rst2("id_responsable") & "','" & rst2("stock") & "','" & rst2("stock_personalizado") & "'," & _
          " '" & Format(rst2("hora_inicio"), "HH:mm") & "','" & Format(rst2("hora_fin"), "HH:mm") & "','" & Format(rst2("horas"), "HH:mm") & "','" & rst2("defecto") & "','" & rst2("activo") & "','" & rst2("id_sucursal") & "','" & rst2("facturacion_detallada") & "','" & rst2("caja_independiente") & "')"
          CnBd.Execute (strCadena)
       End If
       rst2.MoveNext
       DoEvents
       Me.ProgressBar2.Value = i
   Next i
End If

'MIGRAR ALMACEN COMPROBANTES

strCadena = "DELETE FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       
          strCadena = "INSERT INTO almacen_comprobante(`ruc`,`id_alm`,`id_doc`,`serie`,`numero`,`igv`,`defecto`,`id_moneda`,`venta`,`id_formato_impresion`,`serial`,`afecta_caja`,`tipo_movimiento`,`id_usuario`)VALUES " & _
          "('" & KEY_RUC & "','" & rst2("id_alm") & "','" & rst2("id_doc") & "','" & rst2("serie") & "','" & rst2("numero") & "','" & rst2("igv") & "','" & rst2("defecto") & "'," & _
          " '" & rst2("id_moneda") & "','" & rst2("venta") & "','" & rst2("id_formato_impresion") & "','" & rst2("serial") & "','" & rst2("afecta_caja") & "','" & rst2("tipo_movimiento") & "','" & rst2("id_usuario") & "')"
          CnBd.Execute (strCadena)
       
       rst2.MoveNext
   Next i
End If




'MIGRAR SUB PERSONA



strCadena = "SELECT * FROM persona p, entidad_empresa e where p.dni=e.cod_unico and e.id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   Me.ProgressBar2.Min = 0
   Me.ProgressBar2.Max = rst2.RecordCount
   rst2.MoveFirst
   
   For i = 0 To rst2.RecordCount - 1
       strCadena = "SELECT * FROM persona WHERE dni='" & rst2("dni") & "' LIMIT 1"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount < 1 Then
          strCadena = "DELETE FROM entidad_empresa WHERE cod_unico='" & rst2("dni") & "' and id_empresa='" & KEY_RUC & "'"
          CnBd.Execute (strCadena)
          strCadena = "call P_insert_persona('" & Trim(rst2("dni")) & "' " & _
                ",'" & UCase(rst2("a_paterno")) & "', " & _
                "'" & UCase(rst2("a_materno")) & "' " & _
                ",'" & UCase(rst2("nombres")) & "' " & _
                ",'" & UCase(rst2("nombre_completo")) & "' " & _
                ",'" & rst2("direccion") & "' " & _
                ",'" & rst2("celular") & "' " & _
                ",'" & rst2("mail") & "'" & _
                ",'" & rst2("id_transporte") & "' " & _
                ",'" & rst2("id_contable") & "'" & _
                ",'" & rst2("id_proveedor") & "' " & _
                ",'" & rst2("id_personal") & "' " & _
                ",'" & rst2("id_auspeciador") & "' " & _
                ",'" & rst2("id_almacen") & "' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
       End If
       rst2.MoveNext
       DoEvents
       Me.ProgressBar2.Value = i
   Next i
End If



'MIGRAR PRODUCTO

strCadena = "DELETE FROM producto WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM producto WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   Me.ProgressBar2.Min = 0
   Me.ProgressBar2.Max = rst2.RecordCount
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
       
          strCadena = "INSERT INTO producto(`id_producto`,`id_proveedor`,`id_auspiciador`,`id_linea`,`id_sublinea`,`id_categoria`,`id_categoria1`," & _
          "`id_categoria2`,`id_color`,`nombre_prod`,`id_unidad`,`stock_total`,`stock_minimo`,`peso`,`nombre_comercial`,`id_marca`,`id_percepcion`,`id_combo`, " & _
          " `id_igv`,`id_sub_producto`,`id_relacionado`,`cantidad_afecta_stock`,`id_tipo`,`dni_save`,`ruc`)VALUES " & _
          "('" & rst2("id_producto") & "','" & rst2("id_proveedor") & "','" & rst2("id_auspiciador") & "','" & rst2("id_linea") & "','" & rst2("id_sublinea") & "','" & rst2("id_categoria") & "','" & rst2("id_categoria1") & "'," & _
          " '" & rst2("id_categoria2") & "','" & rst2("id_color") & "','" & rst2("nombre_prod") & "','" & rst2("id_unidad") & "','" & rst2("stock_total") & "','" & rst2("stock_minimo") & "','" & rst2("peso") & "','" & rst2("nombre_comercial") & "','" & rst2("id_marca") & "','" & rst2("id_percepcion") & "'" & _
          ",'" & rst2("id_combo") & "','" & rst2("id_igv") & "','" & rst2("id_sub_producto") & "','" & rst2("id_relacionado") & "','" & rst2("cantidad_afecta_stock") & "','" & rst2("id_tipo") & "','" & rst2("dni_save") & "','" & KEY_RUC & "')"
          CnBd.Execute (strCadena)
       
       rst2.MoveNext
       DoEvents
       Me.ProgressBar2.Value = i
   Next i
End If



'MIGRAR ALMACEN COMPROBANTES


strCadena = "DELETE FROM almacen_producto WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
strCadena = "SELECT * FROM almacen_producto WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   Me.ProgressBar2.Min = 0
   Me.ProgressBar2.Max = rst2.RecordCount
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
         ' Call reiniciar_conexion
          strCadena = "INSERT INTO almacen_producto(`id_alm`,`id_producto`,`stock`,`stock_factura`,`stock_contable`,`ruc`,`sector`,`piso`,`andamio`," & _
          "`casillero_x`,`casillero_y`,`precio_venta`,`precio_compra`,precio_mayor,`habilitado`)VALUES " & _
          "('" & rst2("id_alm") & "','" & rst2("id_producto") & "','" & rst2("stock") & "','" & rst2("stock_factura") & "','" & rst2("stock_contable") & "'," & _
          "'" & KEY_RUC & "','" & rst2("sector") & "'," & _
          " '" & rst2("piso") & "','" & rst2("andamio") & "','" & rst2("casillero_x") & "','" & rst2("casillero_y") & "','" & rst2("precio_venta") & "','" & rst2("precio_compra") & "','" & rst2("precio_mayor") & "','" & rst2("habilitado") & "')"
          CnBd.Execute (strCadena)

       rst2.MoveNext
       DoEvents
       Me.ProgressBar2.Value = i
   Next i
End If

'MIGRAR ALMACEN COMPROBANTES

'100:



strCadena = "SELECT * FROM almacen_producto where ruc='" & KEY_RUC & "'"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   Me.ProgressBar2.Min = 0
   Me.ProgressBar2.Max = rst2.RecordCount
   rst2.MoveFirst
   For i = 0 To rst2.RecordCount - 1
        strInventario = formato_item(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), 6)
        strCadena = "INSERT INTO inventario(id_inventario,id_producto,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
        "('" & strInventario & "','" & rst2("id_producto") & "','" & KEY_FECHA & "','" & rst2("id_alm") & "','" & rst2("stock") & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        rst2.MoveNext
        DoEvents
        Me.ProgressBar2.Value = i
   Next i
End If
MsgBox "MIGRACION EXITOSA"
    
    
    
    

End Sub

Private Sub cmdmigrartabla_Click()
End Sub



Private Sub cmdnuevo_Click()
Me.frmgrupoempresarial.Visible = True
Call Resalta(Me.txtrucVinculado)
End Sub

Private Sub cmdProcesar_Click()
Call Save
End Sub

Private Sub cmdprocesarvinculado_Click()
strCadena = "SELECT * FROM grupo_empresarial WHERE ruc='" & KEY_RUC & "' and ruc_vinculado='" & Trim(Me.txtrucVinculado.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    strCadena = "INSERT INTO grupo_empresarial(ruc,ruc_vinculado)VALUES('" & KEY_RUC & "','" & Trim(Me.txtrucVinculado.Text) & "')"
    CnBd.Execute (strCadena)
    
    Call llenar_empresas(hfgrupoempresarial)
    
End If
Me.frmgrupoempresarial.Visible = False
End Sub

Private Sub cmdrealizarmigracion_Click()
    
    Call migrar_ventas(Trim(Me.txtRutaMigracion.Text))
    
    
End Sub

Private Sub cmdSalir_Click()
Unload Me
Exit Sub
End Sub

Private Sub cmdUpdate_stock_migracion_Click()
Dim sys_ConString2 As String
Dim stock_actual As Integer

sys_Server2 = "7.247.54.16"
sys_DataBase2 = "bd_vitekey_repos_ii"
sys_SUser2 = "user_cord" 'DecryptString(ConfigRead("SUser"))
sys_SPassword2 = "@02021974abc2016@123@cord" 'DecryptString(ConfigRead("SPassword"))
db_port = "3306"
sys_ConString2 = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server2 & ";" & _
            "Database=" & sys_DataBase2 & ";" & _
            "UID=" & sys_SUser2 & ";" & _
            "PWD=" & sys_SPassword2 & ";" & _
            " PORT=" & db_port & ";OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
CnBd2.ConnectionString = sys_ConString2
CnBd2.Open






strCadena = "SELECT * FROM almacen_producto WHERE ruc='" & KEY_RUC & "' ORDER BY id_producto"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
   rst2.MoveFirst
   in_periodo = get_periodo_actual(KEY_FECHA)
   
   For i = 0 To rst2.RecordCount - 1
       
       strCadena = "UPDATE almacen_producto SET precio_compra='" & rst2("precio_compra") & "',precio_venta='" & rst2("precio_venta") & "' WHERE id_producto='" & rst2("id_producto") & "' and id_alm='" & rst2("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
       CnBd.Execute (strCadena)
   
       Call put_actualizar_kardex_inventario(rst2("id_producto"), rst2("id_alm"), rst2("stock"), rst2("stock_factura"), in_periodo)
       Call put_kardex_ventas(rst2("id_producto"))
       rst2.MoveNext
       DoEvents
       Me.cmdUpdate_stock_migracion.Caption = str(i)
   Next i
End If


End Sub
Private Sub put_kardex_ventas(ByVal in_producto As String)
strCadena = "SELECT * FROM view_venta_detalle_producto WHERE fecha_emision>='2018-11-18' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' ORDER BY id_venta ASC"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   For j = 0 To rstA.RecordCount - 1
        If rstA("afecta_factura") = "no" Then
        strCadena = "call put_kardex_stock_inventario_v2('01','" & Format(rstA("fecha_emision"), "YYYY-mm-dd") & "','" & rstA("id_venta") & "','" & rstA("id_doc") & "','" & rstA("serie") & "','" & rstA("numero") & "','" & rstA("id_cliente") & "','" & Trim(in_producto) & "','" & rstA("cantidad") & "','0','" & rstA("id_alm") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        rstA.MoveNext
        Else
            strCadena = "UPDATE almacen_producto SET stock_factura=stock_factura-'" & rstA("cantidad") & "' WHERE id_producto='" & rstA("id_producto") & "' and id_alm='" & rstA("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            CnBd.Execute (strCadena)
        End If
        
    
   Next j
End If
End Sub


Private Sub cmdupdateruta_Click()

strCadena = "SELECT * FROM entidad_parametros WHERE actualizador_diferente='no' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "SELECT * FROM version_empresa WHERE ruc='" & rst("cod_unico") & "'"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
          strCadena = "UPDATE version_empresa SET version='" & Val(Me.txtVersion.Text) & "',descarga='" & Trim(Me.txtRutaActualizar.Text) & "',fecha='" & Format(KEY_FECHA, "YYYY-mm-dd") & "' WHERE ruc='" & rst("cod_unico") & "'"
       Else
            strCadena = "INSERT INTO version_empresa(`fecha`,`version`,`descarga`,`actualizado`,`ruc`)VALUES('" & KEY_FECHA & "','" & Val(Me.txtVersion.Text) & "','" & Trim(Me.txtRutaActualizar.Text) & "','si','" & rst("cod_unico") & "')"
       End If
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
   MsgBox "Actualizador Realizado" + Chr(13) + "EMPRESAS:" + str(rst.RecordCount), vbInformation
End If


End Sub

Private Sub cmdventas_conta_Click()


strCadena = "select c.`id_doc`,a.`Serie`,a.`Numero`,a.`IdClienteProveedor`,a.`IdSucursal`,a.`SubTotal`,a.`IdMoneda`,a.`Impuesto` from con_documento a INNER JOIN con_asiento aa ON aa.`IdReferencia`=a.`Id` INNER JOIN `comprobantes` c ON (a.`IdTipoDocumento`=c.`IdTipoDoc`) Where a.`IdEmpresaSis`='20600549490' and a.`IdPeriodo`='1CIX000000000039' and aa.`IdTipoAsiento`='1CIX000000000137'  and  a.`Activo`=aa.`Activo` and a.`Activo`='1' ORDER BY c.`id_doc`,a.`Serie`,a.`Numero`"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       
       strCadena = "SELECT * FROM"
       
       rst.MoveNext
   Next i
End If


Exit Sub






strCadena = "SELECT * FROM producto_sabar"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
   rstIN.MoveFirst
   For i = 0 To rstIN.RecordCount - 1
        
        strCadena = "SELECT * FROM linea WHERE descripcion = '" & Trim(rstP("linea")) & "' AND id_usu='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY id_linea DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_linea = "00001"
            Else
                id_linea = Format(Val(rst("id_linea")) + 1, "00000")
            End If
            
            strCadena = "INSERT INTO linea(id_linea,descripcion,id_usu)VALUES('" & id_linea & "','" & Trim(rstP("linea")) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
            id_linea = rst("id_linea")
        End If
        id_sublinea = "00001"
        'END  LINEA *********
        
        
         '----- MODELO *********
        in_modelo = rstP("sub_linea")
        strCadena = "SELECT * FROM linea_sub WHERE descripcion = '" & in_modelo & "' AND  id_linea='" & id_linea & "' AND id_usu='" & KEY_RUC & "' limit 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM linea_sub where id_usu='" & KEY_RUC & "' ORDER BY id_tipo DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_sublinea = "00001"
            Else
                id_sublinea = Format(Val(rst("id_tipo")) + 1, "00000")
            End If
            strCadena = "INSERT INTO linea_sub(id_tipo,id_linea,descripcion,id_usu)VALUES('" & id_sublinea & "','" & id_linea & "','" & in_modelo & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
          id_sublinea = rst("id_tipo")
        End If
        'END  MODELO *********
        
        
        
        in_modelo = rstP("modelo")
        strCadena = "SELECT * FROM linea_modelo WHERE id_linea='" & id_linea & "' and id_sublinea='" & id_sublinea & "' and  descripcion = '" & in_modelo & "'  AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM linea_modelo WHERE ruc ='" & KEY_RUC & "' ORDER BY id DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_modelo = "00001"
            Else
                id_modelo = rst("id") + 1
            End If
            strCadena = "INSERT INTO linea_modelo(`id_linea`,`id_sublinea`,`descripcion`,`ruc`)VALUES('" & id_linea & "','" & id_sublinea & "','" & in_modelo & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            strCadena = "SELECT * FROM linea_modelo WHERE ruc ='" & KEY_RUC & "' ORDER BY id DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            id_modelo = rst("id")
            
            
        Else
            id_modelo = rst("id")
        End If
        
        '----- MARCA *********
        in_marca = rstP("marca")
        strCadena = "SELECT * FROM marca WHERE descripcion = '" & in_marca & "'  AND id_usu='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM marca  where id_usu ='" & KEY_RUC & "' ORDER BY id_marca DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_marca = "00001"
            Else
                id_marca = Format(Val(rst("id_marca")) + 1, "00000")
            End If
            strCadena = "INSERT INTO marca(id_marca,descripcion,id_usu)VALUES('" & id_marca & "','" & in_marca & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
            id_marca = rst("id_marca")
        End If
        'END  MARCA *********
        rstIN.MoveNext
     Next i
End If

Exit Sub



strCadena = "SELECT id_venta,tc,fecha_emision,id_doc,id_moneda FROM movimiento_venta WHERE  ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       If rst("tc") <= 0 Then
          X = 0
          in_tc = get_tipo_cambio_dia(rst("fecha_emision"), "valor_compra")
          strCadena = "UPDATE movimiento_venta SET tc='" & in_tc & "' WHERE id_venta='" & rst("id_venta") & "' LIMIT 1"
          CnBd.Execute (strCadena)
       End If
       
       in_cta_cobrar = KEY_CTA_COBRAR_PRODUCTO
         
        
         If rst("id_doc") = "0412" Then
             If rst("id_moneda") = "00001" Then
                in_cta_cobrar = "1230101"
            Else
                in_cta_cobrar = "1230102"
            End If
        Else
            X = 0
            
         End If
         
        
       
       strCadena = "UPDATE movimiento_venta SET cta_cobrar='" & Trim(in_cta_cobrar) & "' WHERE id_venta='" & rst("id_venta") & "'"
       CnBd.Execute (strCadena)
       
       
       strCadena = "call P_insert_venta_agenda_xd('" & rst("id_venta") & "')"
       CnBd.Execute (strCadena)
       
       
       rst.MoveNext
       DoEvents
   Next i
End If


End Sub



Private Sub Command1_Click()

End Sub

Private Sub Command10_Click()
Dim Archivo As String
Archivo = Trim("Producto Format" & KEY_RUC) & ".xls"
      'Dim obj As New get_excel
      Set Me.hfproductos.DataSource = Leer_Excel(App.Path & "\comparar_percy\" & Archivo, "Sheet1")
      Me.frm_importacion.Visible = True
      Me.frm_importacion.Top = Me.Command10.Top - 300
      'Set obj = Nothing
      
      

End Sub
Private Function get_registro_producto(ByVal in_producto As String) As Boolean
strCadena = "SELECT id_producto FROM producto WHERE nombre_prod LIKE '%" & Trim(in_producto) & "%' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_registro_producto = True
Else
   get_registro_producto = False
End If

End Function
Private Sub Command11_Click()
' Modulo de Compras
strCadena = "SELECT * FROM kardex k,movimiento_compra c WHERE k.id_movimiento=c.id_compra and k.ruc=c.ruc and c.anulado='si' and c.ruc='" & KEY_RUC & "'  ORDER BY id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM movimiento_compra WHERE anulado='si' and  serie='" & rst("id_serie") & "' and numero='" & rst("id_numero") & "' and  id_compra='" & rst("id_movimiento") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            strCadena = "DELETE FROM kardex WHERE id_producto='" & rst("id_producto") & "' and id_doc='" & rst("id_doc") & "' and id_serie='" & rst("serie") & "' and id_numero='" & rst("id_numero") & "' and id_movimiento='" & rstL("id_compra") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            GoTo upd
        End If
        
        strCadena = "SELECT * FROM movimiento_compra WHERE serie='" & rst("id_serie") & "' and numero='" & rst("id_numero") & "' and  id_compra='" & rst("id_movimiento") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount < 1 Then
            strCadena = "DELETE FROM kardex WHERE id_producto='" & rst("id_producto") & "' and id_doc='" & rst("id_doc") & "' and id_serie='" & rst("serie") & "' and id_numero='" & rst("id_numero") & "' and id_movimiento='" & rstL("id_compra") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
        End If
        
upd:
        strCadena = "SELECT ifnull(sum(cantidad_real),0),ifnull(sum(cantidad_pendiente),0)  FROM kardex WHERE  id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstK(strCadena)
        
        strCadena = "SELECT * FROM almacen_producto WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstA(strCadena)
        If rstA("stock") <> rstK(0) Then
            X = 0
        End If
         If rstA("stock_contable") <> rstK(1) Then
            X = 1
        End If
        
        strCadena = "UPDATE almacen_producto SET stock='" & rstK(0) & "',stock_contable='" & rstK(1) & "' WHERE   id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        rst.MoveNext
        Me.Command11.Caption = str(i)
        DoEvents
   Next i
End If


' Modulo de Transferencias
strCadena = "SELECT * FROM kardex  WHERE id_doc='0009' and ruc='" & KEY_RUC & "' ORDER BY id_producto ASC "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT * FROM movimiento_transferencia WHERE anulado='si' and  serie='" & rst("id_serie") & "' and numero='" & rst("id_numero") & "' and  id_transferencia='" & rst("id_movimiento") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            strCadena = "DELETE FROM kardex WHERE id_producto='" & rst("id_producto") & "' and id_doc='" & rst("id_doc") & "' and id_serie='" & rst("serie") & "' and id_numero='" & rst("id_numero") & "' and id_movimiento='" & rstL("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            GoTo upd2
        End If
        
        strCadena = "SELECT * FROM movimiento_transferencia WHERE serie='" & rst("id_serie") & "' and numero='" & rst("id_numero") & "' and  id_transferencia='" & rst("id_movimiento") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount < 1 Then
            strCadena = "DELETE FROM kardex WHERE id_producto='" & rst("id_producto") & "' and id_doc='" & rst("id_doc") & "' and id_serie='" & rst("serie") & "' and id_numero='" & rst("id_numero") & "' and id_movimiento='" & rstL("id_transferencia") & "' and ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
        End If
        
upd2:
        strCadena = "SELECT ifnull(sum(cantidad_real),0),ifnull(sum(cantidad_pendiente),0)  FROM kardex WHERE  id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstK(strCadena)
        
        strCadena = "SELECT * FROM almacen_producto WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstA(strCadena)
        If rstA("stock") <> rstK(0) Then
            X = 0
        End If
         If rstA("stock_contable") <> rstK(1) Then
            X = 1
        End If
        
        
        strCadena = "UPDATE almacen_producto SET stock='" & rstK(0) & "',stock_contable='" & rstK(1) & "' WHERE   id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        rst.MoveNext
        Me.Command11.Caption = str(i)
        DoEvents
   Next i
End If




End Sub
Private Sub put_kardex_inventario(ByVal in_producto As String, ByVal in_alm As String, ByVal in_cantidad As Double, ByVal in_costo As Single, ByVal in_venta As Single, ByVal in_periodo As String)

           
           strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0089' and id_alm='" & in_alm & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
           Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             
             strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "' and id_alm='" & in_alm & "' and id_doc='0089' LIMIT 1"
             Call ConfiguraRstlocal(strCadena)
             If rstLocal.RecordCount > 0 Then
               in_serie = rstLocal("serie")
               in_numero = formato_item(rstLocal("numero"), 8)
             Else
                MsgBox "CREE EL COMPROBANTE INGRESO A ALMACEN" + Chr(13) + "PARA ESTA SUCURSAL", vbInformation
                Exit Sub
             End If
       End If
             in_cta_compra = KEY_CTA_COMPRA_SOLES
           
           
            strCadena = "call P_insert_compra_ultimate('0089','" & in_alm & "',CURDATE(),CURDATE(),'02'," & _
            "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
            "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
            "'0','0','0','0','0','0','0','0','0','0','0'," & _
            " '" & KEY_USUARIO & "','OBSERVACION','01','" & in_periodo & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
            Call ConfiguraRstPP(strCadena)
            id_compra = rstPP(0)
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(in_producto) & "','" & in_cantidad & "','0'," & _
           "'0','0','0','" & in_cantidad * in_costo & "','0','0', " & _
           "'0','0','0','" & in_cantidad * in_costo & "','0','" & in_costo * in_cantidad & "','" & in_venta & "','" & in_costo & "','" & in_alm & "','" & get_producto(in_producto, KEY_RUC) & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
           strCadena = "call put_kardex_stock_inventario_v2('02','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0089','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & in_producto & "','" & in_cantidad & "','" & in_costo & "','" & in_alm & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
End Sub
Private Sub Command12_Click()
 
       If Me.chk_adicionar_producto.Value = 1 Then
          GoTo avanzar
       End If
        
        If MsgBox("SE VA ELIMINAR TODOS LOS REGISTROS DE PRODUCTOS", vbInformation + vbYesNo, KEY_USUARIO) = vbNo Then
           Exit Sub
        End If
        
        
        If chk_sucursal.Value = 0 Then
           strCadena = "DELETE FROM producto WHERE ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
           
           strCadena = "DELETE FROM producto_proveedor WHERE ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
           
           strCadena = "DELETE FROM producto_unidad WHERE ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
           
           strCadena = "DELETE FROM almacen_producto WHERE ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
        
           strCadena = "DELETE FROM kardex WHERE ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
           
           strCadena = "DELETE FROM linea WHERE id_usu='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
           
           strCadena = "DELETE FROM linea_sub WHERE id_usu='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
         
        Else
           strCadena = "DELETE FROM almacen_producto WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and   ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
           
           strCadena = "DELETE FROM producto_unidad WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and  ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
        
           strCadena = "DELETE FROM kardex WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and  ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
        End If
        
        
        
         strCadena = "SELECT DISTINCT unidad_abrev,unidad FROM inventario_inicial WHERE   ruc='" & Trim(KEY_RUC) & "' order by id_producto asc"
         Call ConfiguraRstP(strCadena)
         If rstP.RecordCount > 0 Then
            rstP.MoveFirst
            For i = 0 To rstP.RecordCount - 1
                in_unidad = rstP("unidad_abrev")
                strCadena = "SELECT * FROM unidad WHERE abreviatura = '" & in_unidad & "' AND id_usu='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRst(strCadena)
                If rst.RecordCount < 1 Then
                    strCadena = "SELECT * FROM unidad WHERE id_usu='" & KEY_RUC & "' ORDER BY id_und DESC LIMIT 1"
                    Call ConfiguraRst(strCadena)
                    If rst.RecordCount > 0 Then
                        id_unidad = Format(Val(rst("id_und") + 1), "00000")
                    Else
                        id_unidad = "00001"
                    End If
                    strCadena = "INSERT INTO unidad(id_und,descripcion,abreviatura,id_usu)VALUES('" & id_unidad & "','" & rstP("unidad") & "','" & in_unidad & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
               
                End If
        'END  MARCA *********
                rstP.MoveNext
            Next i
         End If
        
     
        
        
        
avanzar:
        
        strCadena = "SELECT * FROM con_periodo WHERE Ejercicio='" & Year(KEY_FECHA) & "' and Mes='" & Month(KEY_FECHA) & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            in_periodo = rst("id")
        End If
        
         If chk_sucursal.Value = 0 Then
            strCadena = "SELECT * FROM inventario_inicial WHERE   ruc='" & Trim(KEY_RUC) & "' order by id_producto asc"
         Else
            strCadena = "SELECT * FROM inventario_inicial WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and  ruc='" & Trim(KEY_RUC) & "'"
         End If
        
        Call ConfiguraRstP(strCadena)
        If rstP.RecordCount > 0 Then
           rstP.MoveFirst
            
            
           
        For i = 0 To rstP.RecordCount - 1
            
        '*****LINEA
       
        
        
        strCadena = "SELECT * FROM linea WHERE descripcion = '" & Trim(rstP("linea")) & "' AND id_usu='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY id_linea DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_linea = "00001"
            Else
                id_linea = Format(Val(rst("id_linea")) + 1, "00000")
            End If
            
            strCadena = "INSERT INTO linea(id_linea,descripcion,id_usu)VALUES('" & id_linea & "','" & Trim(rstP("linea")) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
            id_linea = rst("id_linea")
        End If
        id_sublinea = "00001"
        'END  LINEA *********
        
        
         '----- MODELO *********
        in_modelo = rstP("sub_linea")
        strCadena = "SELECT * FROM linea_sub WHERE descripcion = '" & in_modelo & "' AND  id_linea='" & id_linea & "' AND id_usu='" & KEY_RUC & "' limit 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM linea_sub where id_usu='" & KEY_RUC & "' ORDER BY id_tipo DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_sublinea = "00001"
            Else
                id_sublinea = Format(Val(rst("id_tipo")) + 1, "00000")
            End If
            strCadena = "INSERT INTO linea_sub(id_tipo,id_linea,descripcion,id_usu)VALUES('" & id_sublinea & "','" & id_linea & "','" & in_modelo & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
          id_sublinea = rst("id_tipo")
        End If
        'END  MODELO *********
        
        
        
        in_modelo = rstP("modelo")
        strCadena = "SELECT * FROM linea_modelo WHERE id_linea='" & id_linea & "' and id_sublinea='" & id_sublinea & "' and  descripcion = '" & in_modelo & "'  AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM linea_modelo WHERE ruc ='" & KEY_RUC & "' ORDER BY id DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_modelo = "00001"
            Else
                id_modelo = rst("id") + 1
            End If
            strCadena = "INSERT INTO linea_modelo(`id_linea`,`id_sublinea`,`descripcion`,`ruc`)VALUES('" & id_linea & "','" & id_sublinea & "','" & in_modelo & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            strCadena = "SELECT * FROM linea_modelo WHERE ruc ='" & KEY_RUC & "' ORDER BY id DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            id_modelo = rst("id")
            
            
        Else
            id_modelo = rst("id")
        End If
        
        '----- MARCA *********
        in_marca = rstP("marca")
        strCadena = "SELECT * FROM marca WHERE descripcion = '" & in_marca & "'  AND id_usu='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM marca  where id_usu ='" & KEY_RUC & "' ORDER BY id_marca DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_marca = "00001"
            Else
                id_marca = Format(Val(rst("id_marca")) + 1, "00000")
            End If
            strCadena = "INSERT INTO marca(id_marca,descripcion,id_usu)VALUES('" & id_marca & "','" & in_marca & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
            id_marca = rst("id_marca")
        End If
        'END  MARCA *********
        
        
        
        'COLOR *********
       in_color = rstP("color")
       strCadena = "SELECT * FROM imp_color WHERE descripcion = '" & in_color & "' LIMIT 1"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM imp_color ORDER BY id_color DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_color = "0000"
            Else
                id_color = Format(Val(rst("id_color")) + 1, "0000")
            End If
            strCadena = "INSERT INTO imp_color(id_color,descripcion)VALUES('" & id_color & "','" & in_color & "')"
            CnBd.Execute (strCadena)
        Else
            id_color = rst("id_color")
        End If
       'END  MARCA *********
       
       
       
       'UNIDAD MEDIDA *****************
        in_unidad = rstP("unidad_abrev")
        
        strCadena = "SELECT * FROM unidad WHERE abreviatura = '" & in_unidad & "' AND id_usu='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
       
            strCadena = "SELECT * FROM unidad WHERE id_usu='" & KEY_RUC & "' ORDER BY id_und DESC LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                id_unidad = Format(Val(rst("id_und") + 1), "00000")
            Else
               id_unidad = "00001"
            End If
            strCadena = "INSERT INTO unidad(id_und,descripcion,abreviatura,id_usu)VALUES('" & id_unidad & "','" & rstP("unidad") & "','" & in_unidad & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
            id_unidad = rst("id_und")
        End If
        'END  MARCA *********
        
        
      strCadena = "SELECT * FROM persona WHERE dni='" & rstP("codigo_proveedor") & "' LIMIT 1"
      Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                
                strCadena = "call P_insert_persona_ii('" & rstP("codigo_proveedor") & "' " & _
                ",'', " & _
                "'' " & _
                ",'' " & _
                ",'" & Replace(UCase(Trim(rstP("razon_social"))), "'", " ") & "' " & _
                ",'" & Trim(rstP("direccion_proveedor")) & "' " & _
                ",' ' " & _
                ",' '" & _
                ",'no' " & _
                ",'no'" & _
                ",'si' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'si' " & _
                ",'" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
        Else
          
          strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' and  cod_unico='" & rstP("codigo_proveedor") & "' LIMIT 1"
          Call ConfiguraRst(strCadena)
          If rst.RecordCount < 1 Then
            strCadena = "INSERT INTO entidad_empresa(cod_unico,id_empresa,id_cliente,id_proveedor)VALUES('" & rstP("codigo_proveedor") & "','" & KEY_RUC & "','si','si')"
            CnBd.Execute (strCadena)
          End If
     End If
        
        
       
       'strCadena = "SELECT * FROM producto WHERE id_producto='" & rstP("id_producto") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
       'Call ConfiguraRstC(strCadena)
       'If rstc.RecordCount < 1 Then
            strCadena = "INSERT INTO producto (`id_producto`,id_universal,codigo_proveedor,codigo_barra,`id_linea`,`id_sublinea`,`id_moneda`,`id_color`,`nombre_prod`,`id_unidad`,`nombre_comercial`,`id_marca`,`id_igv`,`dni_save`,talla,genero,peso,`ruc`) VALUES " & _
            "('" & rstP("id_producto") & "','" & rstP("codigo_universal") & "','" & rstP("codigo_proveedor") & "','" & rstP("codigo_barra") & "','" & id_linea & "','" & id_sublinea & "','00001','" & id_color & "','" & rstP("nombre_producto") & "','" & id_unidad & "','" & rstP("nombre_comercial") & "','" & id_marca & "','si','" & KEY_USUARIO & "','" & rstP("talla") & "','" & rstP("genero") & "','" & rstP("peso") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
            strCadena = "INSERT INTO producto_proveedor(id_producto,id_proveedor,ruc) VALUES ('" & rstP("id_producto") & "','" & rstP("codigo_proveedor") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
      ' End If
        
ggg:
        
        
        'Unidad2
        If Len(Trim(rstP("unidad_compra"))) > 1 Then
            in_unidad2 = rstP("unidad_compra")
            strCadena = "SELECT * FROM unidad WHERE abreviatura = '" & in_unidad2 & "' AND id_usu='" & KEY_RUC & "' LIMIT 1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
              strCadena = "SELECT * FROM unidad WHERE id_usu='" & KEY_RUC & "' ORDER BY id_und DESC LIMIT 1"
              Call ConfiguraRst(strCadena)
              If rst.RecordCount > 0 Then
                id_unidad2 = Format(Val(rst("id_und") + 1), "00000")
              Else
               id_unidad2 = "00001"
             End If
            
            strCadena = "INSERT INTO unidad(id_und,descripcion,abreviatura,id_usu)VALUES('" & id_unidad2 & "','" & in_unidad2 & "','" & in_unidad2 & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
            id_unidad2 = rst("id_und")
        End If
        End If
        
        
        
        
        
        If Me.chk_sucursal.Value = 1 Then
        
             If Len(Trim(rstP("unidad2"))) > 1 Then
                strCadena = "call put_unidad_producto('" & rstP("id_producto") & "','" & id_unidad2 & "','" & rstP("cantidad_trae2") & "','" & Me.DtcAlmacen.BoundText & "','" & Val(rstP("costo2")) & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                strCadena = "UPDATE producto SET agranel='si' WHERE id_producto='" & rstP("id_producto") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                CnBd.Execute (strCadena)
                
            End If
            strCadena = "SELECT * FROM almacen WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "'"
        Else
            If Trim(rstP("unidad_compra")) <> Trim(in_unidad) And Len(Trim(rstP("unidad_compra"))) > 1 Then
                strCadena = "call put_unidad_producto('" & rstP("id_producto") & "','" & id_unidad & "','1','00001','" & Val(rstP("precio_venta")) & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                strCadena = "call put_unidad_producto('" & rstP("id_producto") & "','" & id_unidad2 & "','" & rstP("cantidad_trae") & "','00001','" & Val(rstP("precio_costo")) & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                strCadena = "UPDATE producto SET agranel='si' WHERE id_producto='" & rstP("id_producto") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                CnBd.Execute (strCadena)
                
            End If
            
            
            strCadena = "SELECT * FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "' ORDER BY id_alm ASC"
        End If
        
        Call ConfiguraRstA(strCadena)
        If rstA.RecordCount > 0 Then
           rstA.MoveFirst
           
           For j = 0 To rstA.RecordCount - 1
               strCadena = "INSERT INTO almacen_producto(`id_alm`,`id_producto`,precio_venta,precio_compra,precio_mayor,precio_alterno_a,`ruc`,`habilitado`,stock) VALUES " & _
               "('" & rstA("id_alm") & "','" & rstP("id_producto") & "','" & rstP("precio_venta") & "','" & rstP("precio_costo") & "','" & rstP("precio_mayor") & "','" & rstP("precio_alterno_a") & "','" & KEY_RUC & "','" & rstP("habilitado") & "','0')"
               CnBd.Execute (strCadena)
              ' If Me.chk_sucursal.Value = 0 Then
              '      in_fisico = rstP("stock_fisico")
                    
              '      If in_fisico > 0 Then
              '         Call put_kardex_inventario(rstP("id_producto"), rstP("id_alm"), Val(in_fisico), rstP("precio_costo"), rstP("precio_venta"), in_periodo)
              '    End If
              'Else
                '   in_fisico = rstP("stock_fisico")
                '   If in_fisico > 0 Then
                '   Call put_kardex_inventario(rstP("id_producto"), rstA("id_alm"), Val(in_fisico), rstP("precio_costo"), rstP("precio_venta"), in_periodo)
                '   End If
               'End If
               
               
               rstA.MoveNext
           Next j
           
           strCadena = "SELECT * FROM inventario_inicial WHERE ruc='" & KEY_RUC & "' and id_producto='" & rstP("id_producto") & "'"
           Call ConfiguraRstIN(strCadena)
           If rstIN.RecordCount > 0 Then
              rstIN.MoveFirst
              For l = 0 To rstIN.RecordCount - 1
                   in_fisico = rstIN("stock_fisico")
                    
                    If in_fisico > 0 Then
                        Call put_kardex_inventario(rstP("id_producto"), rstIN("id_alm"), Val(in_fisico), rstP("precio_costo"), rstP("precio_venta"), in_periodo)
                    End If
                    rstIN.MoveNext
              Next l
           End If
           
           
        End If
        
        
        
        
        
            
nnn:
            
            rstP.MoveNext
            DoEvents
            Me.Command12.Caption = str(i) & "-" & Space(1) & rstP.RecordCount
        Next i
        
        End If
        
        
        
        
        
        
        
        
        
       
        
        
     
        
        
        
        nombre_producto = ""
        in_producto = UCase(Trim(Mid(UCase(Replace(Trim(Me.hfproductos.TextMatrix(i, 1)), "'", " ")), 1, 120)))
        'nombre_producto = in_producto & Space(1) & "  [ " & UCase(in_modelo) & " ]"
        'nombre_producto = UCase(in_modelo) & Space(1) & "  [ " & Trim(in_producto) & " ]"
        nombre_producto = Trim(in_producto)
        'strCadena = "SELECT * FROM producto where ruc='" & KEY_RUC & "' ORDER BY id_producto DESC LIMIT 1"
        'Call ConfiguraRstK(strCadena)
        'If rstK.RecordCount > 0 Then
        '    id_producto = Format(Val(rstK("id_producto") + 1), "00000")
        'Else
         '   id_producto = Format(1, "00000")
        'End If
        
        
        
        
        
        
    
    
      
      






'finnnnn
Exit Sub
GoTo SS



strCadena = "DELETE FROM movimiento_venta where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "delete from movimiento_venta_monto where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT id_producto FROM producto p,linea l  WHERE l.produccion='no'  and  p.id_linea=l.id_linea and p.ruc=l.id_usu and   p.ruc='" & KEY_RUC & "'  "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
    strCadena = "DELETE FROM producto WHERE id_producto='" & rst("id_producto") & "'  and   ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    strCadena = "DELETE FROM almacen_producto_precio WHERE id_producto='" & rst("id_producto") & "' and   ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    rst.MoveNext
   ' DoEvents
    
   Next i
End If

'Unload Me


SS:
strCadena = "DELETE FROM producto WHERE ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE FROM almacen_producto where ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "DELETE FROM linea where id_usu='" & KEY_RUC & "' "
CnBd.Execute (strCadena)

strCadena = "DELETE FROM linea_sub WHERE id_usu='" & KEY_RUC & "' "
CnBd.Execute (strCadena)

Me.ProgressBar2.Min = 0
Me.ProgressBar2.Max = 10000
For i = 0 To 10000
        On Error GoTo salirs
        If Me.hfproductos.TextMatrix(i, 1) <> "" Then
        
        in_producto = Format(Trim(Me.hfproductos.TextMatrix(i, 0)), "00000")
        nombre_producto = UCase(Trim(Me.hfproductos.TextMatrix(i, 1)))
        in_linea = UCase(Trim(Me.hfproductos.TextMatrix(i, 2)))
        in_modelo = Trim(Me.hfproductos.TextMatrix(i, 3))
        in_unidad = UCase(Trim(Me.hfproductos.TextMatrix(i, 4)))
        in_marca = Trim(Me.hfproductos.TextMatrix(i, 5))
        in_color = Trim(Me.hfproductos.TextMatrix(i, 6))
        in_precio_costo = Format(Me.hfproductos.TextMatrix(i, 7), "###0.00")
        in_precio_venta = Format(Me.hfproductos.TextMatrix(i, 8), "###0.00")
        in_precio_mayor = Format(Me.hfproductos.TextMatrix(i, 9), "###0.00")
        in_genero = UCase(Me.hfproductos.TextMatrix(i, 10))
        in_talla = UCase(Me.hfproductos.TextMatrix(i, 11))
        in_stock_contable = (Me.hfproductos.TextMatrix(i, 12))
        in_stock_nocontable = (Me.hfproductos.TextMatrix(i, 13))
        in_peso = Val(Me.hfproductos.TextMatrix(i, 14))
        in_alm = Format(Me.hfproductos.TextMatrix(i, 15), "00000")

        
        
        
        
        If Trim(in_linea) <> "" Then  ' verificar que tenga contenido
        
        '----- CLASIFICACION *********
        
        strCadena = "SELECT * FROM linea WHERE descripcion = '" & in_linea & "' AND id_usu='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY id_linea DESC LIMIT 0,1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_linea = "00001"
            Else
                id_linea = Format(Val(rst("id_linea")) + 1, "00000")
            End If
            
            
            strCadena = "INSERT INTO linea(id_linea,descripcion,id_usu)VALUES('" & id_linea & "','" & in_linea & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
            id_sublinea = "00001"
            
            in_modelo = Trim(Me.hfproductos.TextMatrix(i, 3))
            If Mid(in_modelo, 1, 1) = "Z" Then
               in_modelo = Trim(Mid(in_modelo, 2, 30))
            End If
            
            
            strCadena = "INSERT INTO linea_sub(id_tipo,id_linea,descripcion,id_usu)VALUES('" & id_sublinea & "','" & id_linea & "','" & in_modelo & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
        Else
            id_linea = rst("id_linea")
        End If
        id_sublinea = "00001"
        'END  CLASIFICACION *********
        
        
        
        
        
        
        If Mid(in_modelo, 1, 1) = "Z" Then
               in_modelo = Trim(Mid(in_modelo, 2, 30))
         End If
            
         '----- MODELO *********
        strCadena = "SELECT * FROM linea_sub WHERE descripcion = '" & in_modelo & "' AND  id_linea='" & id_linea & "' AND id_usu='" & KEY_RUC & "' limit 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM linea_sub where id_usu='" & KEY_RUC & "' ORDER BY id_tipo DESC LIMIT 0,1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_sublinea = "00001"
            Else
                id_sublinea = Format(Val(rst("id_tipo")) + 1, "00000")
            End If
            strCadena = "INSERT INTO linea_sub(id_tipo,id_linea,descripcion,id_usu)VALUES('" & id_sublinea & "','" & id_linea & "','" & in_modelo & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
          id_sublinea = rst("id_tipo")
        End If
        'END  CLASIFICACION *********
        
        '----- MARCA *********
        
        strCadena = "SELECT * FROM marca WHERE descripcion = '" & in_marca & "'  AND id_usu='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "SELECT * FROM marca  where id_usu ='" & KEY_RUC & "' ORDER BY id_marca DESC LIMIT 0,1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                id_marca = "00001"
            Else
                id_marca = Format(Val(rst("id_marca")) + 1, "00000")
            End If
            strCadena = "INSERT INTO marca(id_marca,descripcion,id_usu)VALUES('" & id_marca & "','" & in_marca & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
            id_marca = rst("id_marca")
        End If
        'END  MARCA *********
        
        '----- COLOR *********
        id_color = "0000"
     
       ' strCadena = "SELECT * FROM imp_color WHERE descripcion LIKE '%" & in_color & "%'"
       ' Call ConfiguraRst(strCadena)
       ' If rst.RecordCount < 1 Then
       '     strCadena = "SELECT * FROM imp_color ORDER BY id_color DESC LIMIT 0,1"
       '     Call ConfiguraRst(strCadena)
       '     If rst.RecordCount < 1 Then
       '         id_color = "0000"
       '     Else
       '         id_color = Format(Val(rst("id_color")) + 1, "0000")
       '     End If
       '     strCadena = "INSERT INTO imp_color(id_color,descripcion)VALUES('" & id_color & "','" & in_color & "')"
       '     CnBd.Execute (strCadena)
       ' Else
       '     id_color = rst("id_color")
       ' End If
       
        
        
        
        'END  MARCA *********
        '----- Unidad de medida *********
        
        strCadena = "SELECT * FROM unidad WHERE abreviatura = '" & in_unidad & "' AND id_usu='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
       
            strCadena = "SELECT * FROM unidad WHERE id_usu='" & KEY_RUC & "' ORDER BY id_und DESC LIMIT 0,1"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                id_unidad = Format(Val(rst("id_und") + 1), "00000")
            Else
               id_unidad = "00001"
            End If
            strCadena = "INSERT INTO unidad(id_und,descripcion,abreviatura,id_usu)VALUES('" & id_unidad & "','" & in_unidad & "','" & in_unidad & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        Else
            id_unidad = rst("id_und")
        End If
        'END  MARCA *********
        nombre_producto = ""
        in_producto = UCase(Trim(Mid(UCase(Replace(Trim(Me.hfproductos.TextMatrix(i, 1)), "'", " ")), 1, 120)))
        'nombre_producto = in_producto & Space(1) & "  [ " & UCase(in_modelo) & " ]"
        'nombre_producto = UCase(in_modelo) & Space(1) & "  [ " & Trim(in_producto) & " ]"
        nombre_producto = Trim(in_producto)
        'strCadena = "SELECT * FROM producto where ruc='" & KEY_RUC & "' ORDER BY id_producto DESC LIMIT 1"
        'Call ConfiguraRstK(strCadena)
        'If rstK.RecordCount > 0 Then
        '    id_producto = Format(Val(rstK("id_producto") + 1), "00000")
        'Else
         '   id_producto = Format(1, "00000")
        'End If
        
        
        
        
        strCadena = "INSERT INTO producto (`id_producto`,`id_linea`,`id_sublinea`,`id_moneda`,`id_color`,`nombre_prod`,`id_unidad`,`nombre_comercial`,`id_marca`,`id_igv`,`dni_save`,talla,genero,peso,`ruc`) VALUES " & _
        "('" & id_producto & "','" & id_linea & "','" & id_sublinea & "','00001','" & id_color & "','" & nombre_producto & "','" & id_unidad & "','" & nombre_producto & "','" & id_marca & "','si','" & KEY_USUARIO & "','" & in_talla & "','" & in_genero & "','" & in_peso & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        If Me.chk_sucursal.Value = 1 Then
            strCadena = "SELECT * FROM almacen WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "'"
        Else
            strCadena = "SELECT * FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "'"
        End If
        
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
           rstK.MoveFirst
           For j = 0 To rstK.RecordCount - 1
               strCadena = "INSERT INTO almacen_producto(`id_alm`,`id_producto`,precio_venta,precio_compra,precio_mayor,`ruc`,`habilitado`) VALUES ('" & rstK("id_alm") & "','" & id_producto & "','" & in_precio_venta & "','" & in_precio_costo & "','" & in_precio_mayor & "','" & KEY_RUC & "','si')"
               CnBd.Execute (strCadena)
               
               strCadena = "INSERT INTO almacen_producto_precio(`id_alm`,`id_producto`,`precio`,`cant_ini`,`cant_fin`,`ruc`)VALUES " & _
               "('" & rstK("id_alm") & "','" & id_producto & "','" & in_precio_mayor & "','2','5','" & KEY_RUC & "')"
               CnBd.Execute (strCadena)
               rstK.MoveNext
           Next j
        End If
        
    
    
      End If
      End If
siguiente:
      DoEvents
      Me.ProgressBar2.Value = i
      
    Next i
        
salirs:
 MsgBox "SE HA SUBIDO EL DOCUMENTO EXCEL"


End Sub

Private Sub get_update_stock(ByVal in_producto As String, ByVal in_cantidad As Single, ByVal in_fecha As String)
Dim in_stock As Single
strCadena = "SELECT * FROM almacen_producto where id_producto='" & in_producto & "' and  ruc='" & KEY_RUC & "' and id_alm='00001'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   in_stock = rstK("stock") - in_cantidad
   in_fecha = Format(in_fecha, "YYYY-mm-dd")
        strInventario = formato_item(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), 6)
        strCadena = "INSERT INTO inventario(id_inventario,id_producto,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
        "('" & strInventario & "','" & rstK("id_producto") & "','" & in_fecha & "','" & rstK("id_alm") & "','" & in_stock & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
End If
        
   
End Sub
Private Sub get_stock(ByVal in_producto As String, ByVal in_cantidad As Single, ByVal in_fecha As String)
Dim in_stock As Single

   in_stock = in_cantidad
   in_fecha = Format(in_fecha, "YYYY-mm-dd")
        strInventario = formato_item(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), 6)
        strCadena = "INSERT INTO inventario(id_inventario,id_producto,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
        "('" & strInventario & "','" & in_producto & "','" & in_fecha & "','00001','" & in_stock & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        '02516
        
   
End Sub

Private Function get_celular(ByVal in_dni As String) As String
strCadena = "SELECT * FROM persona_telefono WHERE dni='" & in_dni & "' ORDER BY id_telefono DESC LIMIT 1"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    strCadena = "UPDATE persona set celular='" & rstT("telefono") & "' where dni='" & in_dni & "'"
    CnBd.Execute (strCadena)
    
Else
    get_celular = ""
End If
End Function



Private Sub Command14_Click()


End Sub

Private Sub Command13_Click()

End Sub

Private Sub Command15_Click()
       
       
       
       
       
       
       
'pagar en efectivo

strCadena = "SELECT id_venta,fecha_emision,id_cliente,ncliente,documento,tc,id_doc,(total-function_pago_factura(id_venta,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc)) as saldo FROM movimiento_venta WHERE (total-function_pago_factura(id_venta,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc)) >0 and id_doc IN('0003','0001') and id_forma_pago='01' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstlocal(strCadena)
If rstLocal.RecordCount > 0 Then
        For j = 0 To rstLocal.RecordCount - 1
        
        strCadena = "SELECT * FROM movimiento_venta_monto WHERE id_venta='" & Val(rstLocal("id_venta")) & "' and forma_pago='01' "
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            rstK.MoveFirst
            For i = 0 To rstK.RecordCount - 1
                 in_flujo = "1CIX000000000078"
                 in_glosa = "COBRO:" & rstLocal("documento")
                  strCadena = "SELECT * FROM mis_cuentas_det_detalle WHERE id_movimiento='" & rstLocal("id_venta") & "' and monto_pagado='" & rstK("monto_caja") & "' "
                  Call ConfiguraRstP(strCadena)
                  If rstP.RecordCount < 1 Then
                 'Call procesar_transaccion_venta(rstK("id_forma_pago"), KEY_ALM, get_cuenta_pago(rstK("id_forma_pago")), rstLocal("fecha_emision"), "00001", rstLocal("id_cliente"), rstLocal("ncliente"), in_glosa, rstK("monto_caja"), "0", rstLocal("id_venta"), "0", rstLocal("documento"), rstLocal("tc"), rstK("id_tarjeta_operacion"), "1CIX000000000174", in_flujo, KEY_USUARIO, rstLocal("id_doc"), KEY_RUC)
                 Call put_realizar_pago(Val(rstLocal("id_venta")), Val(rstLocal("id_venta")), rstK("monto_caja"), rstLocal("id_doc"), rstLocal("tc"), Val(in_mis_cuentas_det), "01")
                 End If
                 rstK.MoveNext
            Next i
        End If
          
          rstLocal.MoveNext
       Next j
End If
        
       
       
       
       
       
       strCadena = "SELECT * FROM movimiento_venta where anulado='no' and id_doc='0412' and ruc='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
          rst.MoveFirst
          For i = 0 To rst.RecordCount - 1
               strCadena = "SELECT * FROM mis_cuentas_det_detalle WHERE id_detalle='" & rst("id_venta") & "' and  id_movimiento='" & rst("id_referencia") & "'"
               Call ConfiguraRstZ(strCadena)
               If rstZ.RecordCount < 1 Then
                   strCadena = "call sp_mis_cuentas_det_detalle('" & rst("id_venta") & "','0','" & rst("total") & "','" & rst("total") & "','" & rst("id_referencia") & "')"
                   CnBd.Execute (strCadena)
               
               End If
               rst.MoveNext
                
          Next i
       End If
       
            
            
       
   
       
   
   
   
   Exit Sub
   
   
   
   
                
                
   'REGULARIZAR RECIBO
   strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,hora,numero,comprobante,id_cliente,ncliente,total,saldo,anulado,id_moneda,tc,id_alm,id_doc," & _
   " id_proyecto,nombre_completo,descripcion,simbolo,factor,seguro,detraccion,garantia,guia,tseguro,id_forma_pago,referencia,function_pago_factura(id_venta,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc) as pago " & _
   " FROM view_listado_comprobante_ultimate WHERE  id_doc IN('0054')  AND ruc='" & KEY_RUC & "'"
   Call ConfiguraRstT(strCadena)
   If rstT.RecordCount > 0 Then
      rstT.MoveFirst
      For i = 0 To rstT.RecordCount - 1
                
                in_monto_cancelar = rstT("total") - rstT("saldo")
                If in_monto_cancelar > 0 Then
                    in_ref = rstT("comprobante")
                    
                    strCadena = "SELECT * FROM movimiento_venta WHERE  id_doc='0054' AND ruc='" & KEY_RUC & "' order by numero DESC LIMIT 1"
                    Call ConfiguraRst(strCadena)
                    If rst.RecordCount > 0 Then
                    in_numero = Format(Val(rst("numero")) + 1, "000000")
                    Call put_cancelar_comprobante(rst("serie"), in_numero, "0054", rstT("id_cliente"), rstT("id_moneda"), in_monto_cancelar, rstT("tc"), rstT("id_venta"), in_ref, rstT("fecha_emision"), rstT("fecha_emision"), 0)
                    strCadena = "UPDATE almacen_comprobante SET numero='" & Val(in_numero + 1) & "' WHERE id_doc='0054' and ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    End If
                End If
                
                rstT.MoveNext
      Next i
   
   End If
   
   
   Exit Sub
   
   
   
   
                
                

End Sub

Private Sub Command16_Click()

strCadena = "UPDATE movimiento_venta SET cobranza_dudosa='no'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM cobranza_dudosa WHERE estado='no'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        
        in_numero = rst("numero")
buscar:
       strCadena = "SELECT * FROM movimiento_venta WHERE    documento LIKE '%" & in_numero & "%'  and ruc='" & KEY_RUC & "'"
       Call ConfiguraRstZ(strCadena)
       If rstZ.RecordCount > 1 Then
            X = 0
       End If
       
       If rstZ.RecordCount > 0 Then
            in_dudosa = Format(CVDate("01-10" & "-" & rst("anio")), "YYYY-mm-dd")
           strCadena = "UPDATE movimiento_venta SET cobranza_dudosa='si',fecha_cobranza_dudosa='" & in_dudosa & "' WHERE id_venta='" & rstZ("id_venta") & "'"
           CnBd.Execute (strCadena)
           strCadena = "UPDATE cobranza_dudosa SET estado='si' WHERE id_detalle='" & rst("id_detalle") & "' "
           CnBd.Execute (strCadena)
        Else
           

            in_numero = Format(Mid(Trim(rst("numero")), 2, 6), "000000")
            GoTo buscar
       End If
       rst.MoveNext
        
   Next i
End If

End Sub




Private Function get_vendedor(ByVal in_comprobante As String)

strCadena = "select * from movimiento_venta WHERE id_venta='" & Val(in_comprobante) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   get_vendedor = rstP("id_vendedor")
   
End If

End Function

Private Sub Command17_Click()

End Sub

Private Sub Command19_Click()

        in_producto = "68716"
        Call put_kardex_producto(in_producto)
        Exit Sub


For i = 0 To 15000
    '   On Error GoTo Saltar
        If Val(Me.hfproductos.TextMatrix(i, 0)) >= 0 Then
        'in_producto = Format(Trim(Me.hfproductos.TextMatrix(i, 0)), "00000")
       
        
        
        
        strCadena = "SELECT * FROM kardex WHERE id_doc<>'0106' and id_producto='" & in_producto & "' and  ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            rst.MoveFirst
           ' For i = 0 To rst.RecordCount - 1
           '     strCadena = "DELETE FROM kardex WHERE id_kardex='" & rst("id_kardex") & "'"
           '     CnBd.Execute (strCadena)
                rst.MoveNext
           ' Next i
        End If



        
        
        in_costo = Format(Trim(Me.hfproductos.TextMatrix(i, 8)), "###0.00000000")
        in_stock_cix = Format(Trim(Me.hfproductos.TextMatrix(i, 3)), "###0.0000")
        in_stock_piura = Format(Trim(Me.hfproductos.TextMatrix(i, 4)), "###0.0000")
        in_alm1 = "00001"
        in_alm2 = "00002"
        
        
        
        'Almacen 1
        strInventario = formato_item(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), 6)
        in_fecha = "2018-01-01"
        strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
        "('" & strInventario & "','" & in_producto & "','0106','001','" & strInventario & "','" & in_costo & "','2018-01-01','" & in_alm1 & "','" & in_stock_cix & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        strCadena = "call put_kardex_stock_vitekey('06','" & Format(in_fecha, "YYYY-mm-dd") & "','" & Val(strInventario) & "','0106','001','" & strInventario & "','" & KEY_RUC & "','" & in_producto & "','" & in_stock_cix & "','" & in_costo & "','" & in_alm1 & "','" & KEY_RUC & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        'Almacen 2
        strInventario = formato_item(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), 6)
        in_fecha = "2018-01-01"
        
        strCadena = "INSERT INTO inventario(id_inventario,id_producto,id_doc,id_serie,id_numero,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES " & _
        "('" & strInventario & "','" & in_producto & "','0106','001','" & strInventario & "','" & in_costo & "','2018-01-01','" & in_alm2 & "','" & in_stock_piura & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        strCadena = "call put_kardex_stock_vitekey('06','" & Format(in_fecha, "YYYY-mm-dd") & "','" & Val(strInventario) & "','0106','001','" & strInventario & "','" & KEY_RUC & "','" & in_producto & "','" & in_stock_piura & "','" & in_costo & "','" & in_alm2 & "','" & KEY_RUC & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        Call put_kardex_producto(in_producto)
        
        
        
        
       End If
       
   Next i
   
Exit Sub




End Sub

Private Sub put_kardex_producto(ByVal in_producto As String)

Dim in_fechai As Date
in_fechai = "01-01-2018"
For i = 0 To 120
    Call compras(in_fechai, in_producto)
    Call transferencia_ingreso(in_fechai)
    Call transferencia_salida(in_fechai)
    Call ventas(in_fechai)
    Call notas(in_fechai)
    in_fechai = DateAdd("d", 1, in_fechai)
Next i
'End If

End Sub


Private Sub Command2_Click()
Dim nncantidad As Single
Dim ncantidadafecta As Single
strCadena = "SELECT * FROM producto WHERE id_relacionado<>'0' AND cantidad_afecta_stock>0 AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    
    For i = 0 To rst.RecordCount - 1
        ncantidadafecta = rst("cantidad_afecta_stock")
        strCadena = "SELECT sum(cantidad_real) FROM kardex WHERE id_producto='" & rst("id_relacionado") & "' AND id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstT(strCadena)
        
        If IsNull(rstT(0)) = True Then
           nncantidad = 0
        Else
           nncantidad = rstT(0)
        End If
        strCadena = "UPDATE almacen_producto SET stock='" & Val(nncantidad / ncantidadafecta) & "' WHERE id_producto='" & rst("id_producto") & "' AND id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        rst.MoveNext
    Next i
End If

End Sub

Private Sub Command20_Click()


End Sub




Private Sub sin_stock_inicial()
Dim in_fechai As Date
Dim encontro As String
strCadena = "SELECT * FROM producto WHERE ruc='" & KEY_RUC & "' ORDER BY id_producto DESC "
Call ConfiguraRstPP(strCadena)
If rstPP.RecordCount > 0 Then
    rstPP.MoveFirst
    For j = 0 To rstPP.RecordCount - 1
        strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_producto='" & rstPP("id_producto") & "' and ruc='" & KEY_RUC & "' LIMIT 1 "
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
           GoTo siguiente
        End If
            in_fechai = "01-01-2018"
            For m = 0 To 136
                
                ' compras
                strCadena = "select * from view_kardex_compra_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call compras_producto(in_fechai, in_producto)
                End If
                               
                ' transferencias ingreso
                
                strCadena = "select * from view_transferencia_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                  Call transferencia_ingreso_producto(in_fechai, in_producto)
                  Call transferencia_salida_producto(in_fechai, in_producto)
                End If
               'ventas salida
               
                strCadena = "select * from view_kardex_ventas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                
                If rstL.RecordCount > 0 Then
                    Call ventas_producto(in_fechai, in_producto)
                End If
              
                'notas salida
                 
                strCadena = "select * from view_kardex_notas_existe WHERE fecha_emision='" & Format(in_fechai, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                    Call notas_producto(in_fechai, in_producto)
                End If
                
               
                
                in_fechai = DateAdd("d", 1, in_fechai)
                DoEvents
                
            Next m
    DoEvents
        'Me.Command20.Caption = in_producto & "ALM:" & in_alm1
siguiente:
    rstPP.MoveNext
  Next j
End If
           
            
           
        
    MsgBox "INGRESADO"

   

End Sub






Private Sub Command23_Click()

End Sub
Private Function get_id_mes(ByVal in_abrev As String) As String

strCadena = "SELECT * FROM meses WHERE abreviatura='" & Trim(in_abrev) & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_id_mes = rstL("id_mes")
Else
   get_id_mes = 0
End If

End Function

Private Sub Command24_Click()


End Sub



Private Sub Command25_Click()

End Sub



Private Sub Command26_Click()
Dim in_kardex_ini As Double
Dim in_kardex_next As Double
Dim id1 As Double
Dim id2 As Double
Dim in_doc1 As String
Dim in_doc2 As String
Dim in_serie1 As String
Dim in_serie2 As String

strCadena = "SELECT * FROM producto WHERE    ruc='" & KEY_RUC & "' ORDER BY id_producto ASC"
Call ConfiguraRstI(strCadena)
If rstI.RecordCount > 0 Then
   rstI.MoveFirst
   For i = 0 To rstI.RecordCount - 1
       
       
inicializar:
       
       strCadena = "SELECT DISTINCT fecha_emision,id_producto FROM kardex where id_producto='" & rstI("id_producto") & "' and ruc='" & KEY_RUC & "' and id_tipo_movimiento='01' ORDER BY fecha_emision ASC"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
          rstK.MoveFirst
          For j = 0 To rstK.RecordCount - 1
                strCadena = "SELECT * FROM kardex WHERE id_tipo_movimiento='01' and id_producto='" & rstK("id_producto") & "' and fecha_emision='" & Format(rstK("fecha_emision"), "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' ORDER BY id_kardex ASC"
                Call ConfiguraRstL(strCadena)
                If rstL.RecordCount > 0 Then
                   rstL.MoveFirst
                   For k = 0 To rstL.RecordCount - 1
                        in_kardex_ini = rstL("id_movimiento")
                        id1 = rstL("id_kardex")
                        in_doc1 = rstL("id_doc")
                        in_serie1 = rstL("id_serie")
                        If k < rstL.RecordCount - 1 Then
                            rstL.MoveNext
                            in_kardex_next = rstL("id_movimiento")
                            id2 = rstL("id_kardex")
                             in_doc2 = rstL("id_doc")
                             in_serie2 = rstL("id_serie")
                            rstL.MovePrevious
                        Else
                            GoTo s
                        End If
                        
                        If in_kardex_ini > in_kardex_next And in_doc1 = in_doc2 And in_serie1 = in_serie2 Then
                           
                          ' strCadena = "update kardex SET ruc='" & in_kardex_ini & "' WHERE  id_tipo_movimiento='01' and id_producto='" & rstI("id_producto") & "' and fecha_emision='" & Format(rstK("fecha_emision"), "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "' "
                          ' CnBd.Execute (strCadena)
                          ' Call ventas_producto_(rstK("fecha_emision"), rstI("id_producto"))
                          'strCadena = "SELECT * FROM ruc='" & in_kardex_ini & "' and id_tipo_movimiento='01' and id_producto='" & rstI("id_producto") & "' and fecha_emision='" & Format(rstK("fecha_emision"), "YYYY-mm-dd") & "' ORDER BY id_kardex ASC"
                          'Call ConfiguraRstPP(strCadena)
                          'If rstPP.RecordCount > 0 Then
                          '   rstPP.MoveFirst
                          '   For h = 0 To rstPP.RecordCount - 1
                          '      strCadena = "UPDATE kardex SET saldo_stock='" & rstPP("saldo_stock") & "',costo_promedio='" & rstPP("costo_promedio") & "' WHERE id_tipo_movimiento='01' and id_movimiento='" & rstPP("id_movimiento") & "' and fecha_emision='" & Format(rstK("fecha_emision"), "YYYY-mm-dd") & "' and id_producto='" & rstI("id_producto") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                          '      CnBd.Execute (strCadena)
                          '      rstPP.MoveNext
                          '   Next h
                          'End If
                           Call put_kardex_dell(id1, id2)
                            Me.Command26.Caption = rstI("id_producto") & Space(2) & rstL("fecha_emision")
                            GoTo inicializar
                           
                        End If
                        
s:
                        
                        rstL.MoveNext
                   Next k
                End If
                rstK.MoveNext
                DoEvents
          Next j
       End If
       rstI.MoveNext
       DoEvents
       
   Next i
End If


MsgBox "LISTO"
End Sub
Public Sub put_actualizars()
strCadena = "SELECT * FROM kardex WHERE id_producto='" & in_producto & "' and fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and "
End Sub
Public Sub ventas_producto_(ByVal in_fecha As Date, ByVal in_producto As String)
strCadena = "SELECT id_venta,id_tipo,fecha_emision,id_doc,serie,numero,id_cliente,id_producto,cantidad,id_alm,dni_save,ruc,id_orden_compra,id_recepcion FROM vargas_kardex_ventas where diferida='no' and  id_producto='" & in_producto & "' and  fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and ruc='" & KEY_RUC & "'  ORDER BY fecha_emision ASC,id_doc ASC,numero ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       in_tipo = "01"
       strCadena = "call put_kardex_stock_vitekey('" & in_tipo & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & rst("id_venta") & "','" & rst("id_doc") & "','" & rst("serie") & "','" & rst("numero") & "','" & rst("id_cliente") & "','" & rst("id_producto") & "','" & Val(rst("cantidad")) & "','0','" & rst("id_alm") & "','" & rst("dni_save") & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       
       On Error GoTo sit
        
sit:
       rst.MoveNext
   Next i
End If
End Sub
Private Sub put_kardex_dell(ByVal in_kardex_ini As String, ByVal in_kardex_next As String)

Dim in_nuevo As Double
Dim in_costo As Double
Dim in_saldo As Double
Dim in_nuevo2 As Double


strCadena = "INSERT INTO kardex (id_movimiento,fecha_emision,id_doc,dni_save,ruc)VALUES('" & Val(in_kardex_next) & "',CURDATE(),'0000','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
    
strCadena = "SELECT * FROM kardex WHERE id_movimiento='" & Val(in_kardex_next) & "' and dni_save='" & KEY_USUARIO & "'"
Call ConfiguraRstPP(strCadena)
in_nuevo = rstPP("id_kardex")

strCadena = "delete from kardex WHERE id_kardex='" & in_nuevo & "'"
CnBd.Execute (strCadena)

'CAMBIAMOS EL KARDEX AL INI A MENOR
strCadena = "UPDATE kardex SET id_kardex='" & in_nuevo & "' WHERE id_kardex='" & Val(in_kardex_ini) & "' LIMIT 1 "
CnBd.Execute (strCadena)







'strCadena = "INSERT INTO kardex (id_movimiento,fecha_emision,id_doc,dni_save,ruc)VALUES('" & Val(in_kardex_next) & "',CURDATE(),'0000','" & KEY_USUARIO & "','" & KEY_RUC & "')"
'CnBd.Execute (strCadena)
    
'strCadena = "SELECT * FROM kardex WHERE id_movimiento='" & Val(in_kardex_next) & "' and dni_save='" & KEY_USUARIO & "'"
'Call ConfiguraRstPP(strCadena)
'in_nuevo2 = rstPP("id_kardex")

'strCadena = "delete from kardex WHERE id_kardex='" & in_nuevo2 & "'"
'CnBd.Execute (strCadena)

'strCadena = "UPDATE kardex SET id_kardex='" & in_nuevo2 & "' WHERE id_kardex='" & Val(in_kardex_next) & "' LIMIT 1 "
'CnBd.Execute (strCadena)



' ACTUALIZAR

strCadena = "SELECT * FROM kardex WHERE id_kardex='" & in_nuevo & "' "
Call ConfiguraRstPP(strCadena)
If rstPP.RecordCount > 0 Then
    in_costo = rstPP("costo_promedio")
    in_saldo = rstPP("saldo_stock")
End If


strCadena = "SELECT * FROM kardex WHERE id_kardex='" & in_kardex_next & "' "
Call ConfiguraRstPP(strCadena)
If rstPP.RecordCount > 0 Then
    in_costo2 = rstPP("costo_promedio")
    in_saldo2 = rstPP("saldo_stock")
End If

strCadena = "UPDATE kardex SET saldo_stock='" & Val(in_saldo2) & "',costo_promedio='" & Val(in_costo2) & "' WHERE id_kardex='" & in_nuevo & "' LIMIT 1 "
CnBd.Execute (strCadena)

strCadena = "UPDATE kardex SET saldo_stock='" & Val(in_saldo) & "',costo_promedio='" & Val(in_costo) & "' WHERE id_kardex='" & in_kardex_next & "' LIMIT 1 "
CnBd.Execute (strCadena)










End Sub


Private Sub Command27_Click()
Call put_update_kardex


End Sub





Private Sub Command28_Click()

End Sub

Private Sub Command3_Click()

If MsgBox("Desea Realizar el proceso", vbQuestion + vbYesNo, KEY_VENDEDOR) = vbNo Then
    Exit Sub
End If

Dim in_costo As Single
GoTo conta
strCadena = "SELECT id_venta,id_alm FROM movimiento_venta WHERE id_doc IN('0001','0003','0007') and  ruc='" & KEY_RUC & "' ORDER BY id_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   
   For i = 0 To rst.RecordCount - 1
    
   strCadena = "SELECT costo_comprobante('" & rst("id_venta") & "','" & rst("id_alm") & "','" & KEY_RUC & "')"
   Call ConfiguraRstT(strCadena)
   If rstT.RecordCount > 0 Then
       strCadena = "UPDATE `movimiento_venta` SET precio_costo='" & rstT(0) & "' WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
       CnBd.Execute (strCadena)
   End If
    rst.MoveNext
   Next
   
End If

strCadena = "SELECT * FROM movimiento_venta_detalle WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "UPDATE movimiento_venta_detalle SET precio_costo='" & get_precio_costo(rst("id_producto")) & "' WHERE id_detalle_venta='" & rst("id_detalle_venta") & "' and ruc='" & KEY_RUC & "'"
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
End If


conta:
' nota item 60 ya pasados

strCadena = "SELECT * FROM movimiento_venta WHERE  fecha_emision>='" & Format(Me.DtpCini.Value, "YYYY-mm-dd") & "' and anulado='no' and fecha_emision<='" & Format(Me.DtpCfin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "' and id_doc IN('0001','0003') order by fecha_emision,id_doc,serie,numero ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
           On Error GoTo siguiente
            If KEY_RUC = "20128836251" Then
                strCadena = "call _insert_venta_agenda_v2('" & rst("id_venta") & "')"
            Else
                strCadena = "call P_insert_venta_agenda_test('" & rst("id_venta") & "')"
            End If
            
            CnBd.Execute (strCadena)
siguiente:
            rst.MoveNext
            DoEvents
   Next i
End If

strCadena = "SELECT * FROM movimiento_venta WHERE fecha_emision>='" & Format(Me.DtpCini.Value, "YYYY-mm-dd") & "' and anulado='no' and fecha_emision<='" & Format(Me.DtpCfin.Value, "YYYY-mm-dd") & "'  and ruc='" & KEY_RUC & "' and id_doc IN('0007') order by fecha_emision,id_doc,serie,numero ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
           
            If KEY_RUC = "20128836251" Then
                strCadena = "call _insert_venta_agenda_v2('" & rst("id_venta") & "')"
            Else
                strCadena = "call P_insert_venta_agenda_test('" & rst("id_venta") & "')"
            End If
            CnBd.Execute (strCadena)
            rst.MoveNext
            DoEvents
   Next i
End If

MsgBox "CONTABILIDAD INGRESADA CORRECTAMENTE", vbInformation

Exit Sub

End Sub







Private Sub Command30_Click()
If Val(Me.txtidproducto.Text) > 0 Then
    strCadena = "SELECT DISTINCT id_producto,id_alm FROM kardex where id_producto='" & Trim(Me.txtidproducto.Text) & "' and  ruc='" & KEY_RUC & "' ORDER BY id_producto ASC,id_alm ASC"
Else
    strCadena = "SELECT DISTINCT id_producto,id_alm FROM kardex where  ruc='" & KEY_RUC & "' ORDER BY id_producto ASC,id_alm ASC"
End If

Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    in_costo = 0
    in_saldo = 0
    For i = 0 To rst.RecordCount - 1
    
    
    
ini:
      in_costo = 0
      in_saldo = 0
        
    strCadena = "call put_crear_kardex_id_producto('" & rst("id_producto") & "','" & rst("id_alm") & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
        
        
        If chk_update_kardex.Value = 1 Then
                strCadena = "SELECT * FROM tmp_kardex_producto WHERE fecha_emision<'" & Format(Me.DtpKardex.Value, "YYYY-mm-dd") & "' and  id_alm='" & rst("id_alm") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC,id_kardex DESC LIMIT 1"
                Call ConfiguraRstAux(strCadena)
                If rstAux.RecordCount > 0 Then
                    in_saldo = rstAux("saldo_stock")
                Else
                    in_saldo = 0
                End If
            strCadena = "SELECT * FROM tmp_kardex_producto WHERE fecha_emision>='" & Format(Me.DtpKardex.Value, "YYYY-mm-dd") & "' and id_alm='" & rst("id_alm") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC"
        Else
            in_saldo = 0
            strCadena = "SELECT * FROM tmp_kardex_producto WHERE id_alm='" & rst("id_alm") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC,id_kardex ASC"
       End If
            
        
        
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then

            rstK.MoveFirst
            
            For j = 0 To rstK.RecordCount - 1
                in_saldo = in_saldo + rstK("cantidad_real")
                If Val(in_saldo) <> rstK("saldo_stock") Then
                    strCadena = "UPDATE kardex SET saldo_stock='" & Val(in_saldo) & "' WHERE id_producto='" & rst("id_producto") & "' and  id_kardex='" & rstK("id_kardex") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                    CnBd.Execute (strCadena)
                     DoEvents
                    GoTo ini
                End If
                
                If j = rstK.RecordCount - 1 Then
                    strCadena = "SELECT sum(cantidad_real) FROM tmp_kardex_producto WHERE id_alm='" & rst("id_alm") & "' and id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstL(strCadena)
                    If rstL(0) <> Val(in_saldo) Then
                        strCadena = "UPDATE almacen_producto SET stock='" & rstL(0) & "' WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                        CnBd.Execute (strCadena)
                         DoEvents
                    End If
                End If
                rstK.MoveNext
                
            Next j
        End If
        in_costo = 0
        in_saldo = 0
        

        
       
         DoEvents
        Command30.Caption = rst("id_producto") & Space(2) & rst("id_alm") & Space(2) & rst.RecordCount
        rst.MoveNext
        DoEvents
        
    Next i
End If
End Sub
Private Sub put_update_kardex()


If Trim(Me.txtidproducto.Text) > 0 Then
    strCadena = "SELECT * FROM almacen_producto WHERE id_producto='" & Trim(Me.txtidproducto.Text) & "' and   ruc='" & KEY_RUC & "' ORDER BY id_producto DESC"
Else
    strCadena = "SELECT * FROM almacen_producto WHERE ruc='" & KEY_RUC & "' ORDER BY id_producto DESC"
End If

Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        
        strCadena = "SELECT ifnull(sum(cantidad_real),0) FROM kardex WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstZ(strCadena)

        strCadena = "UPDATE almacen_producto SET stock='" & rstZ(0) & "' WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        
        rst.MoveNext
        
        If i Mod 10 = 0 Then
            DoEvents
            Command28.Caption = rst("id_producto")
       End If
        
   Next i
   
   MsgBox "LISTO"
End If






End Sub


Private Sub Command31_Click()
End Sub

Private Sub Command32_Click()
strCadena = "SELECT * FROM almacen_producto WHERE   ruc='" & KEY_RUC & "' ORDER BY id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       
       strCadena = "SELECT costo_promedio FROM kardex WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC, id_kardex DESC LIMIT 1"
       Call ConfiguraRstA(strCadena)
       If rstA.RecordCount > 0 Then
          If Round(rstA("costo_promedio"), 2) <> Round(rst("precio_compra"), 2) Then
            strCadena = "UPDATE almacen_producto SET precio_compra='" & rstA("costo_promedio") & "' WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & rst("id_alm") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            CnBd.Execute (strCadena)
            
            
          
          End If
          
          
          
       End If
       
       
       rst.MoveNext
       DoEvents
       Me.Command32.Caption = str(i) & Space(3) & str(rst.RecordCount)
       
   Next i
End If
End Sub



Private Sub delete_asiento(ByVal in_venta As String, ByVal in_numero As String)

strCadena = "SELECT id FROM con_documento WHERE Activo='1' and  idReferencia='" & in_venta & "'  and idEmpresaSis='" & KEY_RUC & "' "
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   rstK.MoveFirst
   For i = 0 To rstK.RecordCount - 1
       strCadena = "DELETE FROM  con_venta  where IdDocumento = '" & rstK("id") & "' AND Activo='1'"
       CnBd.Execute (strCadena)
       
       strCadena = "DELETE FROM  con_documento  where id = '" & rstK("id") & "' AND Activo='1'"
       CnBd.Execute (strCadena)
       
       strCadena = "SELECT * FROM con_asiento WHERE Glosa LIKE '%" & Trim(in_numero) & "%' and IdTipoAsiento IN('1CIX000000000137','1CIX000000000053','1CIX000000000055') "
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
                       strCadena = "DELETE FROM  con_asientomovimiento  where Id = '" & rstA("id") & "'"
                       CnBd.Execute (strCadena)
                       
                       strCadena = "DELETE FROM  con_asientomovimiento_documento  where IdAsientoMovimiento = '" & rstA("id") & "'"
                       CnBd.Execute (strCadena)
                       
                       strCadena = "DELETE FROM  CON_MovimientoCajaBanco  where IdAsientoMovimiento = '" & rstA("id") & "'"
                       CnBd.Execute (strCadena)
                       rstA.MoveNext
                   Next k
                End If
                
                
                
            rstL.MoveNext
                
          Next j
       End If
       rstK.MoveNext
   Next i
End If


End Sub


Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

strCadena = "select * from movimiento_venta_detalle d,producto p WHERE d.id_producto=p.id_producto and d.ruc=p.ruc and   d.detalle='-- and p.ruc='" & KEY_RUC & "'"

Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    rstK.MoveFirst
    For i = 0 To rstK.RecordCount - 1
        strCadena = "UPDATE movimiento_venta_detalle SET detalle='" & rstK("nombre_prod") & "' WHERE id_detalle_venta='" & rstK("id_detalle_venta") & "'"
        CnBd.Execute (strCadena)
        rstK.MoveNext
        DoEvents
    Next i
End If


End Sub

Private Sub Command6_Click()
strCadena = "SELECT * FROM movimiento_venta_detalle WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = rst.RecordCount
    For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT nombre_prod FROM producto WHERE id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
                productot = rstT("nombre_prod")
        Else
                productot = "--"
        End If
        
        strCadena = "UPDATE movimiento_venta_detalle SET detalle='" & productot & "' WHERE  id_detalle_venta='" & rst("id_detalle_venta") & "'"
        CnBd.Execute (strCadena)
        Me.ProgressBar1.Value = i
                rst.MoveNext
                DoEvents
    Next i
End If
End Sub

Private Sub Command7_Click()
strCadena = "SELECT * FROM movimiento_compra WHERE ruc='" & KEY_RUC & "' order by numero ASC "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
       rst.MoveFirst
       Me.ProgressBar1.Min = 0
       Me.ProgressBar1.Max = rst.RecordCount
       For i = 0 To rst.RecordCount - 1
           
            strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & rst("id_doc") & "' and serie='" & rst("serie") & "' and numero='" & Trim(rst("numero")) & "' and id_proveedor='" & rst("id_proveedor") & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 1 Then
                strCadena = "DELETE FROM movimiento_compra WHERE id_compra='" & rst("id_compra") & "'"
                CnBd.Execute (strCadena)
                
            End If
            Me.ProgressBar1.Value = i
            DoEvents
            rst.MoveNext
       Next i
End If

MsgBox "ok"
End Sub

Private Sub Command8_Click()
Dim sys_ConString2 As String
Dim stock_actual As Integer
sys_Server2 = Trim(Me.txtserver1.Text)
sys_DataBase2 = Trim(Me.txtbaseOrigen1.Text)   'ConfigRead("DataBase")
sys_SUser2 = "user_vitekey" 'DecryptString(ConfigRead("SUser"))
sys_SPassword2 = "02021974abc2014@" 'DecryptString(ConfigRead("SPassword"))
db_port = "3306"
sys_ConString2 = "" & _
            "DRIVER={MySQL ODBC 5.1 Driver};" & _
            "Server=" & sys_Server2 & ";" & _
            "Database=" & sys_DataBase2 & ";" & _
            "UID=" & sys_SUser2 & ";" & _
            "PWD=" & sys_SPassword2 & ";" & _
            " PORT=" & db_port & ";OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384 & ";"
CnBd2.ConnectionString = sys_ConString2
CnBd2.Open


strCadena = "SELECT * FROM almacen_producto WHERE id_alm='00001' and ruc='" & KEY_RUC & "'  order by id_producto"
Call ConfiguraRst2(strCadena)
If rst2.RecordCount > 0 Then
    rst2.MoveFirst
    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Max = rst2.RecordCount
    For i = 0 To rst2.RecordCount - 1
        stock_actual = rst2("stock") + stock_comprado(rst2("id_producto")) - stock_vendido(rst2("id_producto"))
        
        DoEvents
        Me.ProgressBar1.Value = i
        rst2.MoveNext
    Next i
End If



End Sub
Public Function stock_vendido(ByVal in_producto As String) As Integer
    strCadena = "SELECT sum(d.cantidad) FROM movimiento_venta v, movimiento_venta_detalle d WHERE v.afecta_factura='no' and   d.id_producto='" & in_producto & "' and   v.id_venta=d.id_venta and v.anulado='no' and v.ruc='" & KEY_RUC & "' and v.id_venta>'383212'"
    Call ConfiguraRstT(strCadena)
    If IsNull(rstT(0)) = True Then
        stock_vendido = 0
    Else
        stock_vendido = rstT(0)
    End If
End Function
Public Function stock_comprado(ByVal in_producto As String) As Integer
    strCadena = "SELECT sum(d.cantidad) FROM movimiento_compra v, movimiento_compra_detalle d WHERE d.id_producto='" & in_producto & "' and   v.id_compra=d.id_compra and v.anulado='no' and v.ruc='" & KEY_RUC & "' and v.id_compra>'18619'"
    Call ConfiguraRstT(strCadena)
    If IsNull(rstT(0)) = True Then
        stock_comprado = 0
    Else
        stock_comprado = rstT(0)
    End If
End Function


Public Sub insert_inventario(ByVal in_poducto As String, ByVal in_almacen As String, ByVal in_stock_nuevo As Single, ByVal in_stock_factura As Single)
Dim strInventario As String
    
strCadena = "SELECT * FROM almacen_producto A WHERE A.ruc='" & KEY_RUC & "' AND A.id_producto='" & Trim(in_poducto) & "' AND A.id_alm='" & in_almacen & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then

    cod_articulo = rst("id_producto")
    stock_actual = rst("stock")
    strInventario = formato_item(ConsultaUltimoRegistro("inventario", "id_inventario", "ruc", KEY_RUC), 6)
    
    strCadena = "INSERT INTO inventario(id_inventario,id_producto,precio_costo,fecha,id_alm,cantidad,id_usuario,nusuario,ruc)VALUES('" & strInventario & "','" & in_poducto & "','" & rst("precio_compra") & "','" & KEY_FECHA & "','" & in_almacen & "','" & in_stock_nuevo & "','" & KEY_USUARIO & "','" & KEY_VENDEDOR & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
    'strCadena = "UPDATE almacen_producto SET stock_factura ='" & rst("stock_factura") & "' WHERE id_producto='" & in_poducto & "' AND id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
    'CnBd.Execute (strCadena)
    
End If
End Sub

Private Sub Command9_Click()


strCadena = "SELECT * FROM "





Exit Sub
Dim in_producto As String

strCadena = "SELECT * FROM producto_desabilitado WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
  For i = 0 To rst.RecordCount - 1
        in_producto = Format(rst("id_producto"), "00000")
       strCadena = "SELECT id_producto,fecha_emision FROM view_venta_detalle WHERE id_producto='" & in_producto & "' and  fecha_emision>='2018-01-01' and ruc='" & KEY_RUC & "' LIMIT 1"
       Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
          strCadena = "UPDATE producto_desabilitado SET habilitado='si' WHERE id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "' "
          CnBd.Execute (strCadena)
          strCadena = "UPDATE almacen_producto SET habilitado='si' WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
          CnBd.Execute (strCadena)
       Else
          strCadena = "UPDATE almacen_producto SET habilitado='no' WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
          CnBd.Execute (strCadena)
        End If
       rst.MoveNext
       DoEvents
  Next i
End If




End Sub


Private Sub DtcFormaPago_Change()
Call llenarGrid(Me.HfFormapago, Me)
End Sub

Private Sub DtcTipoDoc_Change()
Call llenar_serie(Me.hfgSeries)
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500

Me.DtpInicio_migracion.Value = KEY_FECHA
Me.DtpFin_migracion.Value = KEY_FECHA


If KEY_USUARIO = "42546269" Then
Me.cmdupdateruta.Visible = True
End If


Me.DtpInstalacion.Value = KEY_FECHA
Me.DtpCaducidad.Value = KEY_FECHA
  strCadena = "SELECT  DISTINCT A.id_doc as Codigo, doc_des as Descripcion FROM comprobantes C,almacen_comprobante A WHERE C.id_doc=A.id_doc AND A.ruc='" & KEY_RUC & "' AND A.venta='si' ORDER BY doc_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
  
  strCadena = "SELECT  DISTINCT A.id_doc as Codigo, doc_des as Descripcion FROM comprobantes C,almacen_comprobante A WHERE C.id_doc=A.id_doc AND A.ruc='" & KEY_RUC & "' AND A.venta='si' ORDER BY doc_des"
  Call ConfiguraRst(strCadena)
  'Call LlenaDataCombo(Me.DtcComprobante_buscar)
  
  strCadena = "SELECT id_pais as Codigo,descripcion as Descripcion FROM pais ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPais)
  
  strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMoneda)
  
  
  strCadena = "SELECT Id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion  FROM con_periodo order by codigo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtpPeriodo)
  Me.DtpPeriodo.BoundText = get_periodo_actual(KEY_FECHA)
  
   
  strCadena = "SELECT id_tipo_letra as Codigo,descripcion as Descripcion FROM tipo_letra_impresion ORDER BY id_tipo_letra"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoLetra)
  
    
  
Select Case FrmParametrosEmpresa.Procedencia
    Case nuevo
        Me.txtRuc.Enabled = True
    Case modificar
        Call LLENA
        Call llenarforma_pago
        Call llenar_empresas(hfgrupoempresarial)
    End Select
End Sub
Private Sub llenarforma_pago()
strCadena = "SELECT  id as Codigo, descripcion as Descripcion FROM forma_pago ORDER BY id"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcFormaPago)
  Call llenarGrid(Me.HfFormapago, Me)
End Sub
Private Sub llenar_serie(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
strCadena = "SELECT C.doc_abrev,CONCAT(A.serie,'-',A.numero)as numero,A.serie,defecto FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND  A.id_doc='" & Me.DtcTipoDoc.BoundText & "' AND id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
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
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 1200
       Next
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
             Fila = rstT("doc_abrev") & vbTab & rstT("numero")
             Grilla.AddItem Fila
             If rstT("defecto") = "si" Then
                         For k = 0 To 1
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &H8080FF
                         Next k
             End If
             Fila = ""
             rstT.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub llenar_empresas(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Exit Sub
strCadena = "SELECT * FROM view_grupo_empresarial WHERE  ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub
End If
 
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 4000
       Next
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
             Fila = rstT("ruc_vinculado") & vbTab & rstT("nombre_completo")
             Grilla.AddItem Fila
            
             rstT.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
strCadena = "SELECT * FROM forma_pago_detalle WHERE id='" & Me.DtcFormaPago.BoundText & "' AND ruc='" & KEY_RUC & "'"
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
           Grilla.ColWidth(1) = 500
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 800
      Next
         cabecera = "ID" & vbTab & "COD" & vbTab & "DESCRIPCION" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
         Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id_registro") & vbTab & rst("id_detalle") & vbTab & rst("descripcion") & vbTab & rst("estado")
             Grilla.AddItem Fila
            If (Trim(rst("estado")) = "no") Then
                            For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                            Next k
        End If
        Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub Save()
Dim igv As String
Dim Factura As String
Dim Barras As String, permanente As String
If Me.DtcTipoDoc.BoundText <> "" Then
    KEY_COMPROBANTE = Me.DtcTipoDoc.BoundText
End If

If Me.chkmodelo_color.Value = 1 Then
    in_modelo_color = "si"
Else
    in_modelo_color = "no"
End If
If Me.chktrackin.Value = 1 Then
   KEY_TRACKING = "si"
Else
   KEY_TRACKING = "no"
End If


If Me.chk_linea_credito.Value = 1 Then
    KEY_LINEA_CREDITO = "si"
Else
    KEY_LINEA_CREDITO = "no"
End If

If Me.chk_stock_reservado.Value = 1 Then
    KEY_RESERVA_STOCK = "si"
Else
    KEY_RESERVA_STOCK = "no"
End If


If Me.chk_stock_contable.Value = 1 Then
   KEY_STOCK_CONTABLE = "si"
Else
   KEY_STOCK_CONTABLE = "no"
End If


If Me.chk_grifo.Value = 1 Then
   KEY_GRIFO = "si"
Else
   KEY_GRIFO = "no"
End If


If Me.chk_descuentos.Value = 1 Then
    KEY_DESCUENTO_LINEA = "si"
Else
    KEY_DESCUENTO_LINEA = "no"
End If




If Me.chk_cta_pagar_asiento_global.Value = 1 Then
    KEY_ASIENTO_GLOBAL_CTA_PAGAR = "si"
Else
    KEY_ASIENTO_GLOBAL_CTA_PAGAR = "no"
End If


If Me.chk_bolsa_plastica.Value = 1 Then
   KEY_IMPUESTO_BOLSAS = "si"
Else
   KEY_IMPUESTO_BOLSAS = "no"
End If



If Me.chkServidorcloud.Value = 1 Then
   KEY_SERVIDOR_CLOUD = "si"
Else
   KEY_SERVIDOR_CLOUD = "no"
End If


If Me.chk_keyfacil.Value = 1 Then
    KEY_SERVIDOR_KEYFACIL = "si"
    KEY_TOKEN_SUCURSAL = Trim(Me.txtToken_sucursal.Text)
Else
    KEY_SERVIDOR_KEYFACIL = "no"
End If



If Me.chk_codigo_universal_impresion.Value = 1 Then
    KEY_CODIGO_UNIVERSAL_IMPRESION = "si"
Else
    KEY_CODIGO_UNIVERSAL_IMPRESION = "no"
End If


If Me.chk_skin.Value = 1 Then
    KEY_SKIN = "si"
Else
    KEY_SKIN = "no"
End If



If Me.chkGranel.Value = 1 Then
    KEY_AGRANEL = "si"

Else
    KEY_AGRANEL = "no"
End If

If Me.chk_update_proformas.Value = 1 Then
   KEY_UPDATE_PROFORM = "si"
Else
   KEY_UPDATE_PROFORM = "no"
End If

If Me.chk_cambio_precio_clave.Value = 1 Then
    KEY_CAMBIO_PRECIO_PASS = "si"
Else
    KEY_CAMBIO_PRECIO_PASS = "no"
End If

KEY_PAQUETE_EMPRESARIAL = "01"
If Me.OptStandart.Value = True Then
   KEY_PAQUETE_EMPRESARIAL = "01"
End If

If Me.OptProfesional.Value = True Then
   KEY_PAQUETE_EMPRESARIAL = "02"
End If

If Me.OptPremiun.Value = True Then
    KEY_PAQUETE_EMPRESARIAL = "03"
End If


If Me.chk_impresion_proformas.Value = 1 Then
   KEY_IMPRESION_PROFORMA = "si"
Else
   KEY_IMPRESION_PROFORMA = "no"
End If



If Me.chk_proyecto.Value = 1 Then
   KEY_PROYECTO = "si"
Else
   KEY_PROYECTO = "no"
End If

If Me.chk_alarma_stock.Value = 1 Then
    KEY_ALARMA_STOCK = "si"
Else
    KEY_ALARMA_STOCK = "no"
End If


If Me.chk_seguro_venta.Value = 1 Then
   KEY_SEGURO_VENTA = "si"
Else
   KEY_SEGURO_VENTA = "no"
End If


If Me.chkStock_global.Value = 1 Then
   KEY_STOCK_GLOBAL = "si"
Else
   KEY_STOCK_GLOBAL = "no"
End If


If (Me.chkigv.Value = 1) Then
    KEY_CON_IGV = "si"
Else
    KEY_CON_IGV = "no"
End If
If (Me.ChkAutomatico.Value = 1) Then
    KEY_AUTOMATICO = "si"
Else
    KEY_AUTOMATICO = "no"
End If
If Me.chkfotoproducto.Value = 1 Then
    KEY_FOTO = "si"
Else
    KEY_FOTO = "no"
End If
If Me.chkActivacionPermanente.Value = 1 Then
    permanente = "si"
Else
    permanente = "no"
End If

If Me.chkContador.Value = 1 Then
    KEY_CONTABILIDAD = "si"
Else
    KEY_CONTABILIDAD = "no"
End If

If (Me.ChkCerveceria.Value = 1) Then
    KEY_CERVECERIA = "si"
Else
    KEY_CERVECERIA = "no"
End If
If (Me.ckkFacturas.Value = 1) Then
    Factura = "si"
Else
    Factura = "no"
End If

If (Me.ChkBarras.Value = 1) Then
    Barras = "si"
Else
    Barras = "no"
End If

If (Me.ChkUpdatePrecios.Value = 1) Then
    KEY_UPDATE_PRECIOS = "si"
Else
    KEY_UPDATE_PRECIOS = "no"
End If

If Me.chkfotoproducto.Value = 1 Then
    KEY_FOTO = "si"
Else
    KEY_FOTO = "no"
End If
If Me.ChkHuellaDigital.Value = 1 Then
    KEY_FINGERPRINT = "si"
Else
    KEY_FINGERPRINT = "no"
End If

If Me.chktramitedocumentario.Value = 1 Then
    KEY_TRAMITE = "si"
Else
    KEY_TRAMITE = "no"
End If
If Me.chkCajaIndependiente.Value = 1 Then
    KEY_CAJA_INDEPENDIENTE = "si"
Else
    KEY_CAJA_INDEPENDIENTE = "no"
End If

 If Me.chk_referencia_comprobantes.Value = 1 Then
    KEY_REFERENCIA_COMPROBANTE = "si"
Else
    KEY_REFERENCIA_COMPROBANTE = "no"
 End If
 
 
 If Me.chkfacturacion_electronica.Value = 1 Then
    facturacion_electronica = "si"
    KEY_FACTURACION_ELECTRONICA = "si"
    KEY_TOKEN_CLOUD = Trim(Me.txttoken.Text)
    KEY_TOKEN_LOCAL = Trim(Me.txttoken_local.Text)
    
    
 Else
    facturacion_electronica = "no"
    KEY_FACTURACION_ELECTRONICA = "no"
 End If
 
 
If Me.chksegmentacion_precio.Value = 1 Then
    KEY_SEGMENTACION_PRECIO = "si"
Else
    KEY_SEGMENTACION_PRECIO = "no"
End If


If Me.chk_alerta_corte.Value = 1 Then
    KEY_ALERTA_CORTE = "si"
Else
    KEY_ALERTA_CORTE = "no"
End If


If Me.chkvalidacion_clientes.Value = 1 Then
   KEY_VALIDACION_EXTREMA = "si"
Else
   KEY_VALIDACION_EXTREMA = "no"
End If

If Me.chkMensaualidad.Value = 1 Then
    KEY_GENERADOR_MENSUALIDAD = "si"
Else
    KEY_GENERADOR_MENSUALIDAD = "no"
End If
 
If Me.chk_guia_fraccionada.Value = 1 Then
    KEY_GUIA_FRACCIONADA = "si"
Else
    KEY_GUIA_FRACCIONADA = "no"
End If

KEY_MONEDA = Me.DtcMoneda.BoundText


If Me.chk_envio_sunarp.Value = 1 Then
    KEY_ENVIO_SUNARP = "si"
Else
    KEY_ENVIO_SUNARP = "no"
End If

If Me.chkTransporte_migra.Value = 1 Then
   KEY_TRANSPORTE_MIGRA = "si"
Else
   KEY_TRANSPORTE_MIGRA = "no"
End If
 
 
If Me.chk_entrega_mercaderia.Value = 1 Then
   KEY_CONTROL_MERCADERIA = "si"
Else
   KEY_CONTROL_MERCADERIA = "no"
End If

If Me.chk_producto_duplicado.Value = 1 Then
    KEY_PRODUCTO_REPETIDO = "si"
Else
    KEY_PRODUCTO_REPETIDO = "no"
End If

If Me.chk_planes.Value = 1 Then
    KEY_EMPRESA_PLAN = "si"
Else
    KEY_EMPRESA_PLAN = "no"
End If

If Me.chk_sinfecto_caja.Value = 1 Then
    KEY_SIN_EFECTO_CAJA = "si"
Else
    KEY_SIN_EFECTO_CAJA = "no"
End If


If Me.chk_mostrar_direccion.Value = 1 Then
    KEY_MOSTRAR_SURCURSAL = "si"
Else
    KEY_MOSTRAR_SURCURSAL = "no"
End If













KEY_CTA_COMPRA_SOLES = Trim(Me.txtCompra_cta_pagar_soles.Text)
KEY_CTA_COMPRA_DOLARES = Trim(Me.txtCompra_cta_pagar_dolar.Text)
KEY_CTA_COMPRA_RH = Trim(Me.TxtCuenta_cobrar_rh.Text)


KEY_CTA_PAGAR_SERVICIO = Trim(Me.txtcuenta_pagar_servicio.Text)
KEY_CTA_IGV_VENTA = Trim(Me.txtCuenta_igv_ventas.Text)
KEY_CTA_IGV_SERVICIO_COMPRA = Trim(Me.txtCuentaigv_servicio.Text)






KEY_DIAS_CREDITO = Val(Me.txtDiasCredito.Text)

 If Me.chk_mora_mensualidad.Value = 1 Then
    KEY_MORA = "si"
    KEY_MORA_MONTO = Val(Me.txtMora_monto.Text)
 Else
    KEY_MORA = "no"
    KEY_MORA_MONTO = 0
 End If
 
 If Me.chk_precio_mayor.Value = 1 Then
    KEY_MOSTRAR_PRECIO_MAYOR = "si"
 Else
    KEY_MOSTRAR_PRECIO_MAYOR = "no"
 End If
 If Me.chk_precio_costo.Value = 1 Then
    KEY_MOSTRAR_PRECIO_COSTO = "si"
 Else
    KEY_MOSTRAR_PRECIO_COSTO = "no"
 End If
 
 If Me.chk_bonificaciones.Value = 1 Then
    KEY_BONIFICACIONES = "si"
 Else
    KEY_BONIFICACIONES = "no"
 End If
 
 
 If Me.chk_grupo_empresarial.Value = 1 Then
    KEY_GRUPO_EMPRESARIAL = "si"
 Else
    KEY_GRUPO_EMPRESARIAL = "no"
 End If
 
 If chk_tiendaOnline.Value = 1 Then
    in_tienda_online = "si"
 Else
    in_tienda_online = "no"
 End If
 
 
 If Me.Chk_detalle_combo.Value = 1 Then
    KEY_DETALLE_COMBO = "si"
 Else
    KEY_DETALLE_COMBO = "no"
 End If
 
 
 If Me.chk_notra_credito.Value = 1 Then
    KEY_NOTA_CREDITO_ADMIN = "si"
    KEY_NOTA_CREDITO_USER = Trim(Me.txt_nota_credito_user.Text)
 Else
    KEY_NOTA_CREDITO_ADMIN = "no"
    KEY_NOTA_CREDITO_USER = Trim(Me.txt_nota_credito_user.Text)
 End If
 
 KEY_PAIS = Me.DtcPais.BoundText
 
 KEY_PROVEEDOR = Trim(Me.txtproveedor_servicio.Text)
 
 KEY_NOMBRE_COMERCIAL = Trim(Me.TxtNombrecomercial.Text)
 
 KEY_PORCENTAJE_INTERES = Val(Me.txtPorcentajeCredito.Text)
 KEY_PORCENTAJE_ZONA = Val(Me.txtIncrementoPrecioZona.Text)
        
  If Me.txtRuc.Text = "" Or Me.TxtEmpresa.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
   
 Else
    
        strCadena = "UPDATE entidad_parametros SET descuentos='" & KEY_DESCUENTO_LINEA & "',reserva_stock='" & KEY_RESERVA_STOCK & "',tienda_online='" & in_tienda_online & "',valor_impuesto_bolsa='" & Val(Me.txtImpuesto_bolsa.Text) & "',impuesto_bolsas='" & KEY_IMPUESTO_BOLSAS & "',impresion_proforma='" & KEY_IMPRESION_PROFORMA & "',detalle_consumo_combo='" & KEY_DETALLE_COMBO & "',id_proveedor_servicio='" & Trim(Me.txtproveedor_servicio.Text) & "',porcentaje_incremento_zona='" & Val(txtIncrementoPrecioZona.Text) & "' " & _
        ",porcentaje_interes='" & Val(Me.txtPorcentajeCredito.Text) & "',mostrar_direccion_sucursal='" & KEY_MOSTRAR_SURCURSAL & "',grupo_empresarial='" & KEY_GRUPO_EMPRESARIAL & "',sin_efecto_caja='" & KEY_SIN_EFECTO_CAJA & "',nombre_comercial='" & UCase(Me.TxtNombrecomercial.Text) & "',empresa_planes='" & KEY_EMPRESA_PLAN & "',bonificaciones='" & KEY_BONIFICACIONES & "',alerta_cobranza='" & KEY_ALERTA_CORTE & "',alarma_stock='" & KEY_ALARMA_STOCK & "',cuenta_pagar_servicio='" & KEY_CTA_PAGAR_SERVICIO & "',cuenta_igv_venta='" & KEY_CTA_IGV_VENTA & "',cuenta_igv_compra_servicio='" & KEY_CTA_IGV_SERVICIO_COMPRA & "',id_moneda='" & Me.DtcMoneda.BoundText & "', " & _
        " segmentacion_precio='" & KEY_SEGMENTACION_PRECIO & "',codigo_pais='" & Me.DtcPais.BoundText & "', guia_fraccionada='" & KEY_GUIA_FRACCIONADA & "', codigo_universal_impresion='" & KEY_CODIGO_UNIVERSAL_IMPRESION & "', servidor_keyfacil='" & KEY_SERVIDOR_KEYFACIL & "'," & _
        "token_sucursal='" & KEY_TOKEN_SUCURSAL & "' ,skin='" & KEY_SKIN & "',referencia_comprobante='" & KEY_REFERENCIA_COMPROBANTE & "',grifo='" & KEY_GRIFO & "', linea_credito='" & KEY_LINEA_CREDITO & "',asiento_global_cta_pagar='" & KEY_ASIENTO_GLOBAL_CTA_PAGAR & "',stock_contable='" & KEY_STOCK_CONTABLE & "',stock_global='" & KEY_STOCK_GLOBAL & "',nota_credito_admin='" & KEY_NOTA_CREDITO_ADMIN & "',nota_credito_user='" & KEY_NOTA_CREDITO_USER & "',servidor_cloud='" & KEY_SERVIDOR_CLOUD & "',agranel='" & KEY_AGRANEL & "',cuenta_compra_pagar_rh='" & KEY_CTA_COMPRA_RH & "', modificar_proforma='" & KEY_UPDATE_PROFORM & "', mostrar_precio_costo='" & KEY_MOSTRAR_PRECIO_COSTO & "'," & _
        " mostrar_precio_mayor='" & KEY_MOSTRAR_PRECIO_MAYOR & "',cuenta_compra_pagar_soles='" & Trim(Me.txtCompra_cta_pagar_soles.Text) & "',cuenta_compra_pagar_dolar='" & Trim(Me.txtCompra_cta_pagar_dolar.Text) & "',mora_mensualidad='" & KEY_MORA & "',mora_monto='" & KEY_MORA_MONTO & "', " & _
        "dias_credito='" & Val(Me.txtDiasCredito.Text) & "',producto_repetido='" & KEY_PRODUCTO_REPETIDO & "',control_salida_mercaderia='" & KEY_CONTROL_MERCADERIA & "',cuenta_cobrar_producto='" & Trim(Me.txtCuenta_cobrar_producto.Text) & "',cuenta_cobrar_servicio='" & Trim(Me.txtCuenta_Cobrar_servicio.Text) & "',cuenta_ingreso_producto='" & Trim(Me.txtCuenta_ingreso_producto.Text) & "',cuenta_ingreso_servicio='" & Trim(Me.txtCuenta_ingreso_servicio.Text) & "',generador_mensualidad='" & KEY_GENERADOR_MENSUALIDAD & "'," & _
        "transporte_integrado='" & KEY_TRANSPORTE_MIGRA & "',envio_sunarp_xml='" & KEY_ENVIO_SUNARP & "',id_paquete_empresarial='" & KEY_PAQUETE_EMPRESARIAL & "', " & _
        "cambio_precio_clave='" & KEY_CAMBIO_PRECIO_PASS & "',validacion_extrema_cliente='" & KEY_VALIDACION_EXTREMA & "',cuenta_detraccion='" & Trim(Me.txtdetraccion.Text) & "',porcentaje_detraccion='" & Val(Me.txtprocentaje_detraccion.Text) & "',servicio_seguro='" & KEY_SEGURO_VENTA & "',tracking='" & KEY_TRACKING & "',token_local='" & Trim(Me.txttoken_local.Text) & "', token='" & Trim(Me.txttoken.Text) & "',resolucion_electronica='" & Trim(Me.txtresolucion.Text) & "',facturacion_electronica='" & facturacion_electronica & "', sub_linea_color='" & in_modelo_color & "', igv='" & KEY_CON_IGV & "',factura='" & Trim(Factura) & "',barras='" & Trim(Barras) & "',doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "'," & _
        "automatico='" & KEY_AUTOMATICO & "',caja_independiente='" & KEY_CAJA_INDEPENDIENTE & "',tramite_documentario='" & KEY_TRAMITE & "',cerveceria='" & KEY_CERVECERIA & "',contabilidad='" & KEY_CONTABILIDAD & "',update_precios='" & KEY_UPDATE_PRECIOS & "'," & _
        "foto_producto='" & KEY_FOTO & "',fingerprint='" & KEY_FINGERPRINT & "',id_tipo_letra='" & Trim(Me.DtcTipoLetra.BoundText) & "',instalacion='" & Format(Me.DtpInstalacion.Value, "YYYY-mm-dd") & "',caducidad='" & Format(Me.DtpCaducidad.Value, "YYYY-mm-dd") & "',activacion_permanente='" & permanente & "',proyectos_inversion='" & KEY_PROYECTO & "' WHERE cod_unico='" & Trim(Me.txtRuc.Text) & "'"
        CnBd.Execute (strCadena)
        
        
        
        KEY_TIPO_LETRA = BDBuscarCampo("tipo_letra_impresion", "descripcion", "id_tipo_letra", Trim(Me.DtcTipoLetra.BoundText))
        
        If Me.chkContador.Value = 1 Then
                If Trim(Me.TxtRucContador.Text) <> "" And Len(Trim(Me.TxtRucContador.Text)) > 7 And Len(Trim(Me.TxtRucContador.Text)) < 12 Then
                    strCadena = "UPDATE entidad_empresa SET id_contador='" & Trim(Me.TxtRucContador.Text) & "' WHERE cod_unico='" & KEY_RUC & "' AND id_empresa='0'"
                    CnBd.Execute (strCadena)
                    
                    '******* ACTUALIZAR DATA CONTABLE *******
                    
                    '****************************************
                    
                    
                    strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & KEY_RUC & "' AND id_empresa='" & Trim(Me.TxtRucContador.Text) & "' AND id_cliente='si'"
                    Call ConfiguraRst(strCadena)
                    If rst.RecordCount < 1 Then
                        strCadena = "INSERT INTO entidad_empresa(cod_unico,id_empresa,id_cliente)VALUES('" & KEY_RUC & "','" & Trim(Me.TxtRucContador.Text) & "','si')"
                        CnBd.Execute (strCadena)
                        
                        strCadena = "INSERT INTO entidad_empresa(cod_unico,id_empresa,id_contable)VALUES('" & Trim(Me.TxtRucContador.Text) & "','" & KEY_RUC & "','si')"
                        CnBd.Execute (strCadena)
                        
                    End If
                End If
        Else
                strCadena = "DELETE FROM entidad_empresa WHERE cod_unico='" & KEY_RUC & "' AND id_contador='" & Trim(Me.TxtRucContador.Text) & "' AND id_empresa='0'"
                CnBd.Execute (strCadena)
                
        End If
        
        
        
        strCadena = "SELECT * FROM persona_publico WHERE ruc='" & KEY_RUC & "' AND dni='00000000'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 And Trim(Me.txtDireccionPublico.Text) <> "" Then
            
            strCadena = "UPDATE persona_publico SET direccion='" & Trim(Me.txtDireccionPublico.Text) & "' WHERE ruc='" & KEY_RUC & "' AND dni='00000000'"
            KEY_DIR_PUBLIC = Trim(Me.txtDireccionPublico.Text)
        Else
            
            
            strCadena = "INSERT INTO persona_publico(ruc,dni,direccion)VALUES('" & KEY_RUC & "','00000000','" & Trim(Me.txtDireccionPublico.Text) & "')"
            KEY_DIR_PUBLIC = Trim(Me.txtDireccionPublico.Text)
        End If
        CnBd.Execute (strCadena)
        
        
        KEY_SKFACTURA = Factura
        KEY_BARRAS = Barras
        Call FrmParametrosEmpresa.llenarGrid
        FrmParametrosEmpresa.Procedencia = Neutro
        Unload Me
        
   
  End If
End Sub
Private Sub LLENA()

strCadena = "SELECT * FROM entidad_empresa E,entidad_parametros P,persona U WHERE E.cod_unico=P.cod_unico AND P.cod_unico='" & Trim(FrmParametrosEmpresa.HfgMarcas.TextMatrix(FrmParametrosEmpresa.HfgMarcas.Row, 0)) & "' AND E.cod_unico=U.dni AND P.cod_unico=U.dni AND E.id_empresa='0'"
Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
    strSerie = rst("serie")
    Me.txtRuc.Text = rst("dni")
    Me.txttoken.Text = rst("token")
    Me.txttoken_local.Text = rst("token_local")
    Me.txtDireccion.Text = rst("direccion")
    Me.TxtEmpresa.Text = rst("nombre_completo")
    Me.DtcTipoDoc.BoundText = rst("doc_cod")
    Me.txtCompra_cta_pagar_soles.Text = rst("cuenta_compra_pagar_soles")
    Me.txtCompra_cta_pagar_dolar.Text = rst("cuenta_compra_pagar_dolar")
    Me.TxtCuenta_cobrar_rh.Text = rst("cuenta_compra_pagar_rh")
    Me.txtImpuesto_bolsa.Text = rst("valor_impuesto_bolsa")
    Me.txtproveedor_servicio.Text = rst("id_proveedor_servicio")
    
    Me.txtCuena_letraPagar_soles.Text = rst("cuenta_letra_pagar_soles")
    Me.txtCuena_letraPagar_dolares.Text = rst("cuenta_letra_pagar_dolares")
    
    Me.txtCuena_fet_soles.Text = rst("cuenta_pagar_fet_soles")
    Me.txtCuena_fet_dolares.Text = rst("cuenta_pagar_fet_dolares")
    Me.DtcPais.BoundText = rst("codigo_pais")
    
    '% Credito
    Me.txtPorcentajeCredito.Text = rst("porcentaje_interes")
    Me.txtIncrementoPrecioZona.Text = rst("porcentaje_incremento_zona")
    '******************************************************
    
    If rst("nombre_comercial") = "-" Then
        Me.TxtNombrecomercial.Text = rst("nombre_completo")
    End If
    
    
    If rst("tienda_online") = "si" Then
       Me.chk_tiendaOnline.Value = 1
    Else
       Me.chk_tiendaOnline.Value = 0
    End If
    
     If rst("reserva_stock") = "si" Then
       Me.chk_stock_reservado.Value = 1
    Else
       Me.chk_stock_reservado.Value = 0
    End If
    
    
    If rst("impuesto_bolsas") = "si" Then
       Me.chk_bolsa_plastica.Value = 1
    Else
       Me.chk_bolsa_plastica.Value = 0
    End If
    
    
    If rst("descuentos") = "si" Then
        Me.chk_descuentos.Value = 1
    Else
        Me.chk_descuentos.Value = 0
    End If
    
    
    
    If rst("grifo") = "si" Then
       Me.chk_grifo.Value = 1
    Else
       Me.chk_grifo.Value = 0
    End If
    
    
    If rst("impresion_proforma") = "si" Then
        Me.chk_impresion_proformas.Value = 1
    Else
        Me.chk_impresion_proformas.Value = 0
    End If
    
    
    If rst("grupo_empresarial") = "si" Then
        Me.chk_grupo_empresarial.Value = 1
    Else
        Me.chk_grupo_empresarial.Value = 0
    End If
    
    
    If rst("bonificaciones") = "si" Then
       Me.chk_bonificaciones.Value = 1
    Else
        Me.chk_bonificaciones.Value = 0
    End If
    
    
    If rst("detalle_consumo_combo") = "si" Then
       Me.Chk_detalle_combo.Value = 1
    Else
       Me.Chk_detalle_combo.Value = 0
    End If
    
    
    
    Me.chk_moneda.Value = 1
    Me.DtcMoneda.BoundText = rst("id_moneda")
    
    If rst("segmentacion_precio") = "si" Then
       Me.chksegmentacion_precio.Value = 1
    Else
       Me.chksegmentacion_precio.Value = 0
    End If
    
    
    
    If rst("skin") = "si" Then
       Me.chk_skin.Value = 1
    Else
        Me.chk_skin.Value = 0
    End If
    
    
    If rst("alarma_stock") = "si" Then
       Me.chk_alarma_stock.Value = 1
    Else
       Me.chk_alarma_stock.Value = 0
    End If
    
    If rst("guia_fraccionada") = "si" Then
        Me.chk_guia_fraccionada.Value = 1
    Else
        Me.chk_guia_fraccionada.Value = 0
    End If
    
    
    If rst("nota_credito_admin") = "si" Then
       Me.chk_notra_credito.Value = 1
    Else
       Me.chk_notra_credito.Value = 0
    End If
    If rst("linea_credito") = "si" Then
        Me.chk_linea_credito.Value = 1
    Else
        Me.chk_linea_credito.Value = 0
    End If
    
    
    Me.txt_nota_credito_user.Text = rst("nota_credito_user")
    
    If rst("stock_global") = "si" Then
        Me.chkStock_global.Value = 1
    Else
        Me.chkStock_global.Value = 0
    End If
    
    If rst("referencia_comprobante") = "si" Then
        Me.chk_referencia_comprobantes.Value = 1
    Else
        Me.chk_referencia_comprobantes.Value = 0
    End If
    
    
    If rst("codigo_universal_impresion") = "si" Then
        Me.chk_codigo_universal_impresion.Value = 1
    Else
        Me.chk_codigo_universal_impresion.Value = 0
    End If
    
    If rst("stock_contable") = "si" Then
       Me.chk_stock_contable.Value = 1
    Else
        Me.chk_stock_contable.Value = 0
    End If
    
    
    If rst("cambio_precio_clave") = "si" Then
       Me.chk_cambio_precio_clave.Value = 1
    Else
       Me.chk_cambio_precio_clave.Value = 0
    End If
    
    If rst("agranel") = "si" Then
       Me.chkGranel.Value = 1
    Else
       Me.chkGranel.Value = 0
    End If
    
    If rst("servidor_cloud") = "si" Then
       Me.chkServidorcloud.Value = 1
    Else
       Me.chkServidorcloud.Value = 0
    End If
    
    If rst("envio_sunarp_xml") = "si" Then
       Me.chk_envio_sunarp.Value = 1
    Else
       Me.chk_envio_sunarp.Value = 0
    End If
    
    If rst("modificar_proforma") = "si" Then
       Me.chk_update_proformas.Value = 1
    Else
       Me.chk_update_proformas.Value = 0
    End If
    
    If rst("sin_efecto_caja") = "si" Then
       Me.chk_sinfecto_caja.Value = 1
    Else
       Me.chk_sinfecto_caja.Value = 0
    End If
    
    
    If rst("servicio_seguro") = "si" Then
       Me.chk_seguro_venta.Value = 1
    Else
        Me.chk_seguro_venta.Value = 0
    End If
    If rst("contabilidad") = "si" Then
        Me.chkContador.Value = 1
    Else
        Me.chkContador.Value = 0
    End If
    
    If rst("tracking") = "si" Then
       Me.chktrackin.Value = 1
    Else
       Me.chktrackin.Value = 0
    End If
    
    If rst("validacion_extrema_cliente") = "si" Then
        Me.chkvalidacion_clientes.Value = 1
    Else
        Me.chkvalidacion_clientes.Value = 0
    End If
    
    If rst("proyectos_inversion") = "si" Then
       Me.chk_proyecto.Value = 1
    Else
       Me.chk_proyecto.Value = 0
    End If
    
    If rst("servidor_keyfacil") = "si" Then
        Me.chk_keyfacil.Value = 1
        Me.txtToken_sucursal.Text = rst("token_sucursal")
    Else
        Me.chk_keyfacil.Value = 0
    End If
    
    
    
    
    
    If rst("sub_linea_color") = "si" Then
       chkmodelo_color.Value = 1
    Else
       chkmodelo_color.Value = 0
    End If
    
    Me.txtdetraccion.Text = rst("cuenta_detraccion")
    Me.txtprocentaje_detraccion.Text = rst("porcentaje_detraccion")
    
    If rst("tramite_documentario") = "si" Then
        Me.chktramitedocumentario.Value = 1
    Else
        Me.chktramitedocumentario.Value = 0
    End If
    
    If rst("caja_independiente") = "si" Then
        Me.chkCajaIndependiente.Value = 1
    Else
        Me.chkCajaIndependiente.Value = 0
    End If
    
    If rst("facturacion_electronica") = "si" Then
       Me.chkfacturacion_electronica.Value = 1
    Else
       Me.chkfacturacion_electronica.Value = 0
    End If
    
    Me.txtresolucion.Text = rst("resolucion_electronica")
    If IsNull(rst("instalacion")) = False Then
        Me.DtpInstalacion.Value = rst("instalacion")
    Else
        Me.DtpInstalacion.Value = KEY_FECHA
    End If
    If IsNull(rst("caducidad")) = False Then
        Me.DtpCaducidad.Value = rst("caducidad")
    Else
        Me.DtpCaducidad.Value = KEY_FECHA
    End If
    If rst("activacion_permanente") = "si" Then
        Me.chkActivacionPermanente.Value = 1
    Else
        Me.chkActivacionPermanente.Value = 0
    End If
    
    If rst("id_paquete_empresarial") = "01" Then
       Me.OptStandart.Value = True
    End If
    If rst("id_paquete_empresarial") = "02" Then
       Me.OptProfesional.Value = True
    End If
    If rst("id_paquete_empresarial") = "03" Then
       Me.OptPremiun.Value = True
    End If
    
    If rst("transporte_integrado") = "si" Then
       Me.chkTransporte_migra.Value = 1
    Else
       Me.chkTransporte_migra.Value = 0
    End If
    
    If rst("control_salida_mercaderia") = "si" Then
       Me.chk_entrega_mercaderia.Value = 1
    Else
       Me.chk_entrega_mercaderia.Value = 0
    End If
    
    If rst("producto_repetido") = "si" Then
        Me.chk_producto_duplicado.Value = 1
    Else
        Me.chk_producto_duplicado.Value = 0
    End If
    
    If rst("mora_mensualidad") = "si" Then
       Me.chk_mora_mensualidad.Value = 1
       Me.txtMora_monto.Text = rst("mora_monto")
    Else
        Me.chk_mora_mensualidad.Value = 0
       Me.txtMora_monto.Text = 0
    End If
    
    
    If rst("asiento_global_cta_pagar") = "si" Then
        chk_cta_pagar_asiento_global.Value = 1
    Else
        chk_cta_pagar_asiento_global.Value = 0
    End If
    
    If rst("empresa_planes") = "si" Then
        Me.chk_planes.Value = 1
    Else
        Me.chk_planes.Value = 0
    End If
    
    
    If rst("alerta_cobranza") = "si" Then
        Me.chk_alerta_corte.Value = 1
    Else
        Me.chk_alerta_corte.Value = 0
    End If
    
    
    
    Me.txtCuenta_cobrar_producto.Text = rst("cuenta_cobrar_producto")
    Me.txtCuenta_Cobrar_servicio.Text = rst("cuenta_cobrar_servicio")
    Me.txtCuenta_ingreso_producto.Text = rst("cuenta_ingreso_producto")
    Me.txtCuenta_ingreso_servicio.Text = rst("cuenta_ingreso_servicio")
    Me.txtDiasCredito.Text = rst("dias_credito")
    
    '---
    txtcuenta_pagar_servicio.Text = rst("cuenta_pagar_servicio")
    txtCuenta_igv_ventas.Text = rst("cuenta_igv_venta")
    txtCuentaigv_servicio.Text = rst("cuenta_igv_compra_servicio")
    
    
    
    'Me.TxtAlmacen.text = rst("id_alm")
    Me.DtcTipoLetra.BoundText = rst("id_tipo_letra")
    If Len(rst("id_contador")) > 7 Then
        Me.chkContador.Value = 1
        strCadena = "SELECT * FROM persona WHERE dni='" & rst("id_contador") & "'"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
            Me.TxtRucContador.Text = rst("id_contador")
            Me.lblRazonContador.Caption = rstT("nombre_completo")
            
        Else
            Me.lblRazonContador.Caption = "--------"
            
        End If
    
         
         Me.FrameContador.Enabled = False
    End If
    
    If rst("igv") = "si" Then
        Me.chkigv.Value = 1
    Else
        Me.chkigv.Value = 0
    End If
    
    If rst("automatico") = "si" Then
        Me.ChkAutomatico.Value = 1
    Else
        Me.ChkAutomatico.Value = 0
    End If
    If rst("contabilidad") = "si" Then
        Me.OptContabilidad.Value = True
    Else
        Me.OptInventario.Value = True
    End If
    
    If rst("cerveceria") = "si" Then
       Me.ChkCerveceria.Value = 1
    Else
        Me.ChkCerveceria.Value = 0
    End If
    
    If rst("factura") = "si" Then
        Me.ckkFacturas.Value = 1
    Else
        Me.ckkFacturas.Value = 0
    End If
     If rst("barras") = "si" Then
        Me.ChkBarras.Value = 1
    Else
        Me.ChkBarras.Value = 0
    End If
    If rst("update_precios") = "si" Then
        Me.ChkUpdatePrecios.Value = 1
    Else
        Me.ChkUpdatePrecios.Value = 0
    End If
    
    If rst("foto_producto") = "si" Then
        Me.chkfotoproducto.Value = 1
    Else
        Me.chkfotoproducto.Value = 0
    End If
    If rst("fingerprint") = "si" Then
        Me.ChkHuellaDigital.Value = 1
    Else
        Me.ChkHuellaDigital.Value = 0
    End If
    
    Set rst = Nothing
    
    strCadena = "SELECT * FROM persona_publico WHERE dni='00000000' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            Me.txtDireccionPublico.Text = rst("direccion")
        Else
            Me.txtDireccionPublico.Text = ""
        End If
        Set rst = Nothing
  End If
End Sub

Private Sub HfFormapago_DblClick()
If Val(Me.HfFormapago.TextMatrix(Me.HfFormapago.Row, 0)) > 0 Then
    strCadena = "SELECT * FROM forma_pago_detalle WHERE id_registro='" & Val(Me.HfFormapago.TextMatrix(Me.HfFormapago.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If rst("estado") = "si" Then
            strCadena = "UPDATE forma_pago_detalle set estado='no' WHERE id_registro='" & Me.HfFormapago.TextMatrix(Me.HfFormapago.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
        Else
            strCadena = "UPDATE forma_pago_detalle set estado='si' WHERE id_registro='" & Me.HfFormapago.TextMatrix(Me.HfFormapago.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
        End If
        CnBd.Execute (strCadena)
        Call llenarGrid(Me.HfFormapago, Me)
    End If
End If
End Sub



Private Sub Image1_Click()
Me.frm_importacion.Visible = False
End Sub

Private Sub Image3_Click()
Me.frmcontable.Visible = False
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub



Private Sub TxtEmpresa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtDireccion.SetFocus
End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtEmpresa.SetFocus
End If
End Sub

Private Sub TxtRucContador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = buscar
    FrmPersona.Show
    Exit Sub
End If
End Sub

Private Sub txtrucVinculado_Change()
strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtrucVinculado.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.lblgrupoempresa.Caption = rst("nombre_completo")
Else
    Me.lblgrupoempresa.Caption = ""
End If
End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmgeneracionmensualidad 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   15735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GENERACION DE MENSUALIDAD"
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
      Height          =   8055
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   15735
      Begin VB.CheckBox chk_demarcar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "DESMARCAR TODOS"
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
         Left            =   13560
         TabIndex        =   30
         Top             =   1150
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   30
         Left            =   360
         TabIndex        =   28
         Top             =   2520
         Visible         =   0   'False
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
         _Version        =   393216
         Format          =   170655745
         CurrentDate     =   43639
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   615
         Left            =   11880
         TabIndex        =   9
         Top             =   7320
         Width           =   1695
         _extentx        =   2990
         _extenty        =   1085
         btype           =   5
         tx              =   "PROCESAR"
         enab            =   -1  'True
         font            =   "frmgeneracionmensualidad.frx":0000
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   15790320
         bcolo           =   15790320
         fcol            =   8388608
         fcolo           =   8388608
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmgeneracionmensualidad.frx":0028
         picn            =   "frmgeneracionmensualidad.frx":0046
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin MSDataListLib.DataCombo dtcmes 
         Height          =   330
         Left            =   1080
         TabIndex        =   7
         Top             =   1080
         Width           =   5415
         _ExtentX        =   9551
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   11865
         TabIndex        =   8
         Top             =   6960
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VitekeySoft.ChameleonBtn cmdsalirdetalle 
         Height          =   615
         Left            =   13680
         TabIndex        =   10
         Top             =   7320
         Width           =   1695
         _extentx        =   2990
         _extenty        =   1085
         btype           =   5
         tx              =   "SALIR"
         enab            =   -1  'True
         font            =   "frmgeneracionmensualidad.frx":368E
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   15790320
         bcolo           =   15790320
         fcol            =   8388608
         fcolo           =   8388608
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmgeneracionmensualidad.frx":36B6
         picn            =   "frmgeneracionmensualidad.frx":36D4
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcAnio 
         Height          =   330
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPlan 
         Height          =   5415
         Left            =   1080
         TabIndex        =   29
         Top             =   1440
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   9551
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AÑO :"
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
         Left            =   570
         TabIndex        =   12
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIODO  :"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   720
      End
   End
   Begin VB.Frame frmPeriodo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CREACION DE PERIODO"
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
      Height          =   3975
      Left            =   5520
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   8775
      Begin VB.Frame frmdetalle_periodo 
         BackColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   7455
         Begin VB.TextBox txtDescripcion 
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
            Left            =   1800
            TabIndex        =   23
            Top             =   1560
            Width           =   5415
         End
         Begin MSDataListLib.DataCombo DtcAnioPeriodo 
            Height          =   330
            Left            =   1800
            TabIndex        =   19
            Top             =   600
            Width           =   2175
            _ExtentX        =   3836
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
         Begin MSDataListLib.DataCombo DtcMesPeriodo 
            Height          =   330
            Left            =   1800
            TabIndex        =   22
            Top             =   1080
            Width           =   2175
            _ExtentX        =   3836
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
         Begin VitekeySoft.ChameleonBtn cmdSalirPeriodo 
            Height          =   735
            Left            =   6480
            TabIndex        =   25
            Top             =   2160
            Width           =   735
            _extentx        =   1296
            _extenty        =   1296
            btype           =   5
            tx              =   "SALIR"
            enab            =   -1  'True
            font            =   "frmgeneracionmensualidad.frx":66FC
            coltype         =   2
            focusr          =   -1  'True
            bcol            =   16777215
            bcolo           =   16777215
            fcol            =   8388608
            fcolo           =   8388608
            mcol            =   12632256
            mptr            =   1
            micon           =   "frmgeneracionmensualidad.frx":6724
            picn            =   "frmgeneracionmensualidad.frx":6742
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdSavePeriodo 
            Height          =   735
            Left            =   5640
            TabIndex        =   26
            Top             =   2160
            Width           =   735
            _extentx        =   1296
            _extenty        =   1296
            btype           =   5
            tx              =   "SAVE"
            enab            =   -1  'True
            font            =   "frmgeneracionmensualidad.frx":976A
            coltype         =   2
            focusr          =   -1  'True
            bcol            =   16777215
            bcolo           =   16777215
            fcol            =   8388608
            fcolo           =   8388608
            mcol            =   12632256
            mptr            =   1
            micon           =   "frmgeneracionmensualidad.frx":9792
            picn            =   "frmgeneracionmensualidad.frx":97B0
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   2
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin VB.Label Label7 
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
            Left            =   360
            TabIndex        =   24
            Top             =   1560
            Width           =   990
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MES :"
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
            Left            =   960
            TabIndex        =   21
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AÑO :"
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
            Left            =   960
            TabIndex        =   20
            Top             =   600
            Width           =   390
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPeriodo 
         Height          =   3135
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   5530
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
      Begin VitekeySoft.ChameleonBtn cmdnuevo_periodo 
         Height          =   855
         Left            =   7920
         TabIndex        =   16
         Top             =   720
         Width           =   735
         _extentx        =   1296
         _extenty        =   1508
         btype           =   5
         tx              =   "NUEVO"
         enab            =   -1  'True
         font            =   "frmgeneracionmensualidad.frx":CDF8
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   16777215
         bcolo           =   16777215
         fcol            =   8388608
         fcolo           =   8388608
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmgeneracionmensualidad.frx":CE20
         picn            =   "frmgeneracionmensualidad.frx":CE3E
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   2
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn ChameleonBtn2 
         Height          =   855
         Left            =   7920
         TabIndex        =   17
         Top             =   1635
         Width           =   735
         _extentx        =   1296
         _extenty        =   1508
         btype           =   5
         tx              =   "UPDATE"
         enab            =   -1  'True
         font            =   "frmgeneracionmensualidad.frx":F572
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   16777215
         bcolo           =   16777215
         fcol            =   8388608
         fcolo           =   8388608
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmgeneracionmensualidad.frx":F59A
         picn            =   "frmgeneracionmensualidad.frx":F5B8
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   2
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   8400
         Picture         =   "frmgeneracionmensualidad.frx":11BF4
         Top             =   240
         Width           =   240
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   14520
      TabIndex        =   0
      Top             =   480
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      btype           =   5
      tx              =   "Generar"
      enab            =   -1  'True
      font            =   "frmgeneracionmensualidad.frx":14A98
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   8388608
      fcolo           =   8388608
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmgeneracionmensualidad.frx":14AC0
      picn            =   "frmgeneracionmensualidad.frx":14ADE
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfmensualidad 
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   12938
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
   Begin VitekeySoft.ChameleonBtn cmdeliminar 
      Height          =   855
      Left            =   14520
      TabIndex        =   2
      Top             =   2320
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      btype           =   5
      tx              =   "Eliminar"
      enab            =   -1  'True
      font            =   "frmgeneracionmensualidad.frx":1711A
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   8388608
      fcolo           =   8388608
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmgeneracionmensualidad.frx":17142
      picn            =   "frmgeneracionmensualidad.frx":17160
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   855
      Left            =   14520
      TabIndex        =   3
      Top             =   4150
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      btype           =   5
      tx              =   "Salir"
      enab            =   -1  'True
      font            =   "frmgeneracionmensualidad.frx":195AC
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   8388608
      fcolo           =   8388608
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmgeneracionmensualidad.frx":195D4
      picn            =   "frmgeneracionmensualidad.frx":195F2
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdPeriodo 
      Height          =   855
      Left            =   14520
      TabIndex        =   13
      Top             =   3240
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      btype           =   5
      tx              =   "Periodo"
      enab            =   -1  'True
      font            =   "frmgeneracionmensualidad.frx":199E2
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   8388608
      fcolo           =   8388608
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmgeneracionmensualidad.frx":19A0A
      picn            =   "frmgeneracionmensualidad.frx":19A28
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
      Height          =   855
      Left            =   14520
      TabIndex        =   27
      Top             =   1400
      Width           =   975
      _extentx        =   1720
      _extenty        =   1508
      btype           =   5
      tx              =   "Afiliados"
      enab            =   -1  'True
      font            =   "frmgeneracionmensualidad.frx":1BC84
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   16777215
      bcolo           =   16777215
      fcol            =   8388608
      fcolo           =   8388608
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmgeneracionmensualidad.frx":1BCAC
      picn            =   "frmgeneracionmensualidad.frx":1BCCA
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   2
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GENERACION MENSUALIDADES"
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
      TabIndex        =   4
      Top             =   120
      Width           =   2370
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8040
      Left            =   0
      Top             =   0
      Width           =   15735
   End
End
Attribute VB_Name = "frmgeneracionmensualidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdmodificar_Click()

End Sub

Private Sub chk_demarcar_Click()
If Me.chk_demarcar.Value = 1 Then
   For i = 0 To Me.HfPlan.Rows - 1
        If Val(Me.HfPlan.TextMatrix(i, 0)) > 0 Then
            Me.HfPlan.TextMatrix(i, 6) = Chr(168)
        End If
   Next i
End If
End Sub

Private Sub cmdNuevo_Click()


Call Me.llenarGrid_plan(Me.HfPlan)
Me.DtcAnio.BoundText = Me.hfmensualidad.TextMatrix(Me.hfmensualidad.Row, 3)
Me.dtcmes.BoundText = Val(Me.hfmensualidad.TextMatrix(Me.hfmensualidad.Row, 0))
Me.frmdetalle.Visible = True
End Sub

Private Sub cmdnuevo_periodo_Click()
frmdetalle_periodo.Visible = True
End Sub

Private Sub cmdPeriodo_Click()
strCadena = "SELECT id_mes as Codigo,descripcion as Descripcion from mes order by id_mes "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMesPeriodo)
Me.DtcMesPeriodo.BoundText = Month(KEY_FECHA)

strCadena = "SELECT anio as Codigo,anio as Descripcion from anio  order by anio ASC "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAnioPeriodo)
Me.DtcAnioPeriodo.BoundText = Year(KEY_FECHA)
Call llenarGridPeriodo(Me.HfPeriodo)
Me.frmPeriodo.Visible = True
End Sub


Private Sub put_fecha_corte(ByVal in_ruc As String)

strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & in_ruc & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   in_fecha_fin = DateSerial(Year(KEY_FECHA), Val(Month(KEY_FECHA)) + 1, 1 - 1)
   in_fecha_corte = DateAdd("d", rstL("dias_prorroga"), in_fecha_fin)
   
   strCadena = "UPDATE entidad_empresa SET fecha_corte='" & Format(in_fecha_corte, "YYYY-mm-dd") & "' WHERE cod_unico='" & in_ruc & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
   CnBd.Execute (strCadena)
   
   strCadena = "UPDATE entidad_parametros SET caducidad='" & Format(in_fecha_corte, "YYYY-mm-dd") & "' WHERE cod_unico='" & in_ruc & "' LIMIT 1"
   CnBd.Execute (strCadena)
   
   
End If
End Sub



Private Sub cmdprocesar_Click()
Dim in_usuarios As Integer
Dim in_mes As Integer
Dim in_dias As Integer
If Val(Me.dtcmes.BoundText) > 0 And Val(Me.DtcAnio.BoundText) > 0 Then
   strCadena = "SELECT * FROM cobranza_periodo WHERE id_periodo='" & Val(Me.dtcmes.BoundText) & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
    in_mes = rst("id_mes")
    in_fecha_fin = DateSerial(Me.DtcAnio.BoundText, Val(rst("id_mes")) + 1, 1 - 1)
    in_fecha_ini = "01-" & Format(Val(rst("id_mes")), "00") & "-" & Trim(Me.DtcAnio.BoundText)
    in_dias = DateDiff("d", in_fecha_ini, in_fecha_fin)
   End If

   
  
      Me.ProgressBar1.Min = 0
      Me.ProgressBar1.Max = Me.HfPlan.Rows
      in_usuarios = 0
      For i = 0 To Me.HfPlan.Rows - 2
          If Me.HfPlan.TextMatrix(i, 6) = Chr(254) Then

               in_ruc = Me.HfPlan.TextMatrix(i, 1)
               in_producto = Me.HfPlan.TextMatrix(i, 7)
               in_producto_des = get_producto(in_producto) & Space(1) & "[" & Trim(Me.dtcmes.Text) & "]"
               in_precio = Format(Me.HfPlan.TextMatrix(i, 5), "###0.00")
               
               strCadena = "call sp_insert_servicio_mensual('" & Me.dtcmes.BoundText & "','" & in_ruc & "','" & in_producto & "','" & in_producto_des & "','" & in_precio & "','" & KEY_USUARIO & "','" & KEY_ALM & "','" & Format(in_fecha_ini, "YYYY-mm-dd") & "','" & Format(in_fecha_fin, "YYYY-mm-dd") & "','" & KEY_RUC & "')"
               CnBd.Execute (strCadena)
               
               Call put_fecha_corte(in_ruc)
               
               
               in_usuarios = in_usuarios + 1
            End If
            
          
          
          
         
          Me.ProgressBar1.Value = i
          DoEvents
    Next i
   End If
   
   strCadena = "UPDATE cobranza_periodo SET procesado='si',usuarios='" & in_usuarios & "' WHERE id_periodo='" & Val(Me.dtcmes.BoundText) & "' and ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
   Call llenarGrid_periodos(hfmensualidad)
   Me.frmdetalle.Visible = False
   

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdsalirdetalle_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub cmdSalirPeriodo_Click()
Me.frmdetalle_periodo.Visible = False
End Sub

Private Sub cmdSavePeriodo_Click()
Dim in_descripcion As String
Dim in_fecha_fin As Date
Dim in_fecha_ini As Date
If Trim(Me.txtDescripcion.Text) <> "" Then
   in_descripcion = UCase(Trim(Me.txtDescripcion.Text)) + Space(1) + "[ " + Mid(Me.DtcMesPeriodo.Text, 1, 3) + Space(1) + Trim(Me.DtcAnioPeriodo.Text) + " ]"
   in_fecha_fin = DateSerial(Me.DtcAnio.BoundText, Val(Me.DtcMesPeriodo.BoundText) + 1, 1 - 1)
   in_fecha_ini = "01-" & Format(Me.DtcMesPeriodo.BoundText, "00") & "-" & Trim(Me.DtcAnioPeriodo.BoundText)
   strCadena = "put_mensualidad_cobranza('" & Val(Me.DtcMesPeriodo.BoundText) & "','" & Me.DtcAnioPeriodo.BoundText & "','" & in_descripcion & "','" & KEY_USUARIO & "','" & Format(in_fecha_ini, "YYYY-mm-dd") & "','" & Format(in_fecha_fin, "YYYY-mm-dd") & "','" & KEY_RUC & "')"
   CnBd.Execute (strCadena)
   Call llenarGridPeriodo(Me.HfPeriodo)
   
   strCadena = "SELECT id_periodo as Codigo,descripcion as Descripcion from cobranza_periodo WHERE ruc='" & KEY_RUC & "' order by id_periodo "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcmes)

   Me.dtcmes.BoundText = Month(KEY_FECHA)
   Me.frmdetalle_periodo.Visible = False
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50

strCadena = "SELECT id_periodo as Codigo,descripcion as Descripcion from cobranza_periodo WHERE ruc='" & KEY_RUC & "' order by id_periodo "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcmes)

strCadena = "SELECT anio as Codigo,anio as Descripcion from anio  order by anio ASC "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAnio)
Me.DtcAnio.BoundText = Year(KEY_FECHA)
Call llenarGrid_periodos(hfmensualidad)

'Call llenarGrid(Me.hfmensualidad, Month(KEY_FECHA))


End Sub
Private Sub llenar_periodo(ByVal Grilla As MSHFlexGrid, ByVal in_periodo As Integer)
'On Error GoTo salir
strCadena = "SELECT *FROM view_afiliados_colegio Where ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub

End If
   
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 1400
           Grilla.ColWidth(4) = 1000
        Next
         cabecera = "CODIGO" & vbTab & "NIVEL" & vbTab & "GRADO" & vbTab & "AFILIADOS" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        
        
        For i = 1 To rst.RecordCount
            
             Fila = rst("id_nivel") & vbTab & rst("nivel") & vbTab & rst("grado") & vbTab & rst("afiliados") & vbTab & Chr(168)
             Grilla.AddItem Fila
              
              With Grilla
                            .Row = i  ' se posiciona en la fila
                            .col = 4 '  .. en la columna
                            ' cambia la fuente para esta celda
                            
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            ' edita la celda
                            ' If rstT("estado") = "no" Then
                             '   estado = Chr(168)
                            'Else
                             '   estado = Chr(254)
                            'End If
                            
                        End With
            rst.MoveNext
        Next i
Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal in_periodo As Integer)
'On Error GoTo salir
strCadena = "SELECT *FROM view_mensualidad_periodo WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub

End If
   
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 3500
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1200
        Next
         cabecera = "ID" & vbTab & "PERIODO" & vbTab & "AÑO" & vbTab & "DESCRIPCION" & vbTab & "REGISTRO" & vbTab & "OPERADOR"
         Grilla.AddItem cabecera
         For k = 1 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        
        
        For i = 1 To rst.RecordCount
            
             Fila = rst("id_periodo") & vbTab & rst("periodo") & vbTab & rst("anio") & vbTab & rst("descripcion") & vbTab & rst("fecha_registro") & vbTab & rst("nombre_completo")
             Grilla.AddItem Fila
            rst.MoveNext
        Next i
Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub llenarGrid_periodos(ByVal Grilla As MSHFlexGrid)
'On Error GoTo salir
strCadena = "SELECT *FROM cobranza_periodo WHERE ruc='" & KEY_RUC & "' order by id_anio DESC, id_mes DESC "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub

End If
   
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 3500
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 3000
           Grilla.ColWidth(5) = 1800
           Grilla.ColWidth(6) = 1800
        Next
         cabecera = "CODIGO" & vbTab & "PERIODO" & vbTab & "REGISTRO" & vbTab & "AÑO" & vbTab & "USUARIOS" & vbTab & "OPERADOR" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        
        
        For i = 1 To rst.RecordCount
             If rst("procesado") = "si" Then
                in_estado = "GENERADO"
             Else
                in_estado = "PENDIENTE"
             End If
             Fila = Format(rst("id_periodo"), "00000") & vbTab & rst("descripcion") & vbTab & Format(rst("fecha_registro"), "dd-mm-YYYY") & vbTab & rst("id_anio") & vbTab & rst("usuarios") & vbTab & get_persona(rst("dni_save")) & vbTab & in_estado
             Grilla.AddItem Fila
             If rst("procesado") = "si" Then
             For k = 4 To 6
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &H80FF&
              Next k
             End If
            rst.MoveNext
        Next i
Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Public Sub llenarGrid_plan(ByVal Grilla As MSHFlexGrid)
Dim in_monto_acumulado As Double
Dim in_monto_deuda As Double

On Error GoTo salir
strCadena = "SELECT * FROM view_empresa_planes WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
 
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
  
   Grilla.Rows = 0
            ReDim arrColWidth(1 To rst.Fields.Count)
            For Each Campo In rst.Fields
                Grilla.ColWidth(0) = 0
                Grilla.ColWidth(1) = 1300
                Grilla.ColWidth(2) = 5500
                Grilla.ColWidth(3) = 3000
                Grilla.ColWidth(4) = 1800
                Grilla.ColWidth(5) = 1500
                Grilla.ColWidth(6) = 500
                Grilla.ColWidth(7) = 0
            Next
            cabecera = "ID" & vbTab & "DNI/RUC" & vbTab & "NOMBRE CLIENTE" & vbTab & "PLAN CONTRATADO" & vbTab & "TIPO PLAN" & vbTab & "PRECIO SERVICIO" & vbTab & "ST" & vbTab & "IDPRODUCTO"
        
         Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        in_monto_acumulado = 0
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If rst("habilitado") = "si" Then
                in_habilitado = "ACTIVO"
            Else
                in_habilitado = "INACTIVO"
            End If
            
            If rst("pago_mensual") = "si" Then
               in_forma = "MENSUAL"
            End If
            If rst("pago_3meses") = "si" Then
               in_forma = "[3] MESES"
            End If
            If rst("pago_6meses") = "si" Then
               in_forma = "[6] MESES"
            End If
            If rst("pago_anual") = "si" Then
               in_forma = "[12] MESES"
            End If
            in_estado = Chr(254)
            
             Fila = rst("id") & vbTab & rst("dni") & vbTab & Trim(rst("nombre_completo")) & vbTab & rst("descripcion") & vbTab & in_forma & vbTab & Format(rst("monto"), "#,##0.00") & vbTab & in_estado & vbTab & rst("id_producto")
             Grilla.AddItem Fila
             If rst("habilitado") = "no" Then
                For k = 1 To 5
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
               Next k
            Else
                in_monto_acumulado = in_monto_acumulado + rst("monto")
                
            End If
            
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 6 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
                            
                        End With
        
        
            in_monto_deuda = in_monto_deuda
        rst.MoveNext
        Next i
        
         Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(in_monto_acumulado, "#,##0.00")
         Grilla.AddItem Fila
             
                For k = 5 To 5
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
               Next k
  
  Grilla.ColAlignment(0) = 1
  Grilla.ColAlignment(1) = 1
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
         
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub llenar_listado(ByVal Grilla As MSHFlexGrid, ByVal in_periodo As String)
'On Error GoTo salir
Dim in_habilitado As String
Dim in_acumulado As Single
strCadena = "SELECT *FROM view_cobranza_lista WHERE id_periodo='" & Val(in_periodo) & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub

End If
   
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 500
           Grilla.ColWidth(1) = 1300
           Grilla.ColWidth(2) = 4000
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1000
           Grilla.ColWidth(6) = 1300
        Next
         cabecera = " ID " & vbTab & "USUARIO" & vbTab & "DESCRIPCION" & vbTab & "PERIODO" & vbTab & "FECHA INICIO" & vbTab & "MONTO" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        in_acumulado = 0
        
        For i = 0 To rst.RecordCount - 1
             If rst("habilitado") = "si" Then
                in_habilitado = "ACTIVO"
                in_acumulado = in_acumulado + rst("precio_venta")
             Else
                in_habilitado = "INACTIVO"
             End If
             
             Fila = Format(i + 1, "000") & vbTab & rst("dni") & vbTab & rst("nombre_completo") & vbTab & rst("fecha_inscripcion") & vbTab & Format(rst("fecha_inicio_cobranza"), "dd-mm-YYYY") & vbTab & Format(rst("precio_venta"), "#,##0.00") & vbTab & in_habilitado
             Grilla.AddItem Fila
             If rst("habilitado") = "no" Then
             For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
             Next k
             
             End If
            rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(in_acumulado, "#,##0.00") & vbTab & in_habilitado
             Grilla.AddItem Fila
Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub llenarGridPeriodo(ByVal Grilla As MSHFlexGrid)
'On Error GoTo salir
strCadena = "SELECT *FROM view_periodo_cobranza WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub

End If
   
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 2800
           Grilla.ColWidth(4) = 1000
           Grilla.ColWidth(5) = 1200
        Next
         cabecera = "ID" & vbTab & "PERIODO" & vbTab & "AÑO" & vbTab & "DESCRIPCION" & vbTab & "REGISTRO" & vbTab & "OPERADOR"
         Grilla.AddItem cabecera
         For k = 1 To 5
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        
        
        For i = 0 To rst.RecordCount - 1
            
             Fila = rst("id_periodo") & vbTab & rst("mes") & vbTab & rst("id_anio") & vbTab & rst("periodo") & vbTab & Format(rst("fecha_registro"), "dd-mm-YYYY") & vbTab & rst("nombre_completo")
             Grilla.AddItem Fila
            rst.MoveNext
        Next i
Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub HfPlan_Click()
If Val(Me.HfPlan.TextMatrix(Me.HfPlan.Row, 0)) > 0 Then
    If Me.HfPlan.TextMatrix(Me.HfPlan.Row, 6) = Chr(168) Then
       Me.HfPlan.TextMatrix(Me.HfPlan.Row, 6) = Chr(254)
    Else
       Me.HfPlan.TextMatrix(Me.HfPlan.Row, 6) = Chr(168)
    End If
End If
End Sub

Private Sub Image1_Click()
Me.frmPeriodo.Visible = False
End Sub

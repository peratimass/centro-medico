VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmmisproyectos 
   BorderStyle     =   0  'None
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   20055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DETALLE PROYECTO"
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
      Height          =   7935
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   18735
      Begin VB.TextBox txtgarantia 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   5240
         Width           =   2055
      End
      Begin VB.TextBox txtsaldo 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   5760
         Width           =   2055
      End
      Begin VB.Frame frm_cobros 
         Appearance      =   0  'Flat
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   5920
         TabIndex        =   41
         Top             =   960
         Visible         =   0   'False
         Width           =   12735
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCobros 
            Height          =   5655
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   9975
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
         Begin VitekeySoft.ChameleonBtn cmdcerrarcobros 
            Height          =   225
            Left            =   12360
            TabIndex        =   43
            Top             =   60
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   397
            BTYPE           =   3
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
            COLTYPE         =   3
            FOCUSR          =   -1  'True
            BCOL            =   8421631
            BCOLO           =   8421631
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmisproyectos.frx":0000
            PICN            =   "frmmisproyectos.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LISTADO DE COMPROBANTES"
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
            Height          =   210
            Left            =   200
            TabIndex        =   44
            Top             =   60
            Width           =   2220
         End
      End
      Begin VB.TextBox txt_idproyecto 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   8880
         TabIndex        =   36
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtdescripcion 
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
         ForeColor       =   &H00C00000&
         Height          =   675
         Left            =   3360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   1540
         Width           =   4215
      End
      Begin VB.Frame frmpersonal 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "PERSONAL INVOLUCRADO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   9240
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   6735
         Begin VB.TextBox txtsueldobase 
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
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   1440
            TabIndex        =   40
            Top             =   840
            Width           =   1575
         End
         Begin VB.CheckBox chk_coordinador 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "COORDINADOR"
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
            Height          =   290
            Left            =   5040
            TabIndex        =   38
            Top             =   360
            Width           =   1530
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPersonal 
            Height          =   2415
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   4260
            _Version        =   393216
            ForeColor       =   8388608
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   8388608
            BackColorBkg    =   16777215
            GridColor       =   12632064
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
         Begin MSDataListLib.DataCombo DtcPersonal 
            Height          =   330
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   4575
            _ExtentX        =   8070
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
         Begin VitekeySoft.ChameleonBtn cmdadd 
            Height          =   300
            Left            =   5040
            TabIndex        =   30
            Top             =   720
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   529
            BTYPE           =   8
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
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmisproyectos.frx":2ED0
            PICN            =   "frmmisproyectos.frx":2EEC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmddell 
            Height          =   300
            Left            =   5040
            TabIndex        =   31
            Top             =   1080
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   529
            BTYPE           =   8
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
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmmisproyectos.frx":54D1
            PICN            =   "frmmisproyectos.frx":54ED
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SALARIO BASE:"
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
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   1140
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PERSONAL ASIGNADO AL PROYECTO"
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
            TabIndex        =   37
            Top             =   45
            Width           =   2730
         End
      End
      Begin VB.TextBox txtcliente 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   7680
         TabIndex        =   25
         Top             =   2700
         Width           =   855
      End
      Begin VB.TextBox txtcontratista 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   7680
         TabIndex        =   24
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtcobros 
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
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   4680
         Width           =   2055
      End
      Begin VB.TextBox txtgastos 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox txtvalorizacion_inicial 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   3360
         TabIndex        =   18
         Top             =   3480
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo DtcContratista 
         Height          =   330
         Left            =   3360
         TabIndex        =   16
         Top             =   2280
         Width           =   4215
         _ExtentX        =   7435
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
      Begin MSComCtl2.DTPicker DtpInicio 
         Height          =   375
         Left            =   3360
         TabIndex        =   14
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
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
         Format          =   61210624
         CurrentDate     =   42997
      End
      Begin MSComCtl2.DTPicker DtpFin 
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         Top             =   1080
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
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
         Format          =   61210624
         CurrentDate     =   42997
      End
      Begin MSDataListLib.DataCombo dtcCliente 
         Height          =   330
         Left            =   3360
         TabIndex        =   17
         Top             =   2700
         Width           =   4215
         _ExtentX        =   7435
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
      Begin VitekeySoft.ChameleonBtn cmdcancelar 
         Height          =   855
         Left            =   15000
         TabIndex        =   22
         Top             =   5400
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
         MICON           =   "frmmisproyectos.frx":86F8
         PICN            =   "frmmisproyectos.frx":8714
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdsave 
         Height          =   855
         Left            =   13920
         TabIndex        =   23
         Top             =   5400
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
         MICON           =   "frmmisproyectos.frx":8B04
         PICN            =   "frmmisproyectos.frx":8B20
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdcomprobantes 
         Height          =   405
         Left            =   5520
         TabIndex        =   32
         Top             =   4680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "..."
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
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmmisproyectos.frx":C168
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdgastos 
         Height          =   405
         Left            =   5520
         TabIndex        =   33
         Top             =   4080
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "..."
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
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmmisproyectos.frx":C184
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GARANTIA [5%]"
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
         Left            =   1440
         TabIndex        =   49
         Top             =   5280
         Width           =   1185
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OPERADOR :"
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
         TabIndex        =   47
         Top             =   7320
         Width           =   945
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SALDO POR COBRAR :"
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
         Left            =   1440
         TabIndex        =   45
         Top             =   5880
         Width           =   1635
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCION :"
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
         Left            =   1980
         TabIndex        =   35
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         X1              =   8640
         X2              =   8640
         Y1              =   1560
         Y2              =   5190
      End
      Begin VB.Label lblencargado 
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
         Height          =   270
         Left            =   3360
         TabIndex        =   26
         Top             =   7320
         Width           =   4185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMPROBANTES COBRADOS :"
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
         TabIndex        =   20
         Top             =   4800
         Width           =   2265
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GASTOS REGISTRADOS A LA FECHA :"
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
         Left            =   330
         TabIndex        =   13
         Top             =   4200
         Width           =   2745
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALORIZACION INICIAL :"
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
         Left            =   1290
         TabIndex        =   12
         Top             =   3600
         Width           =   1785
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE FINAL :"
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
         Left            =   1905
         TabIndex        =   11
         Top             =   2760
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTRATISTA :"
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
         Left            =   1935
         TabIndex        =   10
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA CULMINACION ESTIMADA :"
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
         Left            =   525
         TabIndex        =   9
         Top             =   1080
         Width           =   2550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INICIO  :"
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
         Left            =   1935
         TabIndex        =   8
         Top             =   600
         Width           =   1140
      End
   End
   Begin VB.TextBox TxtApellido 
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
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1995
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   7935
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   18735
      _ExtentX        =   33046
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
   Begin VitekeySoft.ChameleonBtn cmdexit 
      Height          =   855
      Left            =   18960
      TabIndex        =   2
      Top             =   5610
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
      MICON           =   "frmmisproyectos.frx":C1A0
      PICN            =   "frmmisproyectos.frx":C1BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmddelete 
      Height          =   855
      Left            =   18960
      TabIndex        =   3
      Top             =   2850
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmisproyectos.frx":C5AC
      PICN            =   "frmmisproyectos.frx":C5C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdupdate 
      Height          =   855
      Left            =   18960
      TabIndex        =   4
      Top             =   1965
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "DETALLE"
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
      MICON           =   "frmmisproyectos.frx":EA12
      PICN            =   "frmmisproyectos.frx":EA2E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   18960
      TabIndex        =   5
      Top             =   1080
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmisproyectos.frx":ED48
      PICN            =   "frmmisproyectos.frx":ED64
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdreporte 
      Height          =   855
      Left            =   18960
      TabIndex        =   50
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "REPORTE"
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
      MICON           =   "frmmisproyectos.frx":F1B6
      PICN            =   "frmmisproyectos.frx":F1D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MIS PROYECTOS :"
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
      Left            =   450
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   120
      Top             =   180
      Width           =   16215
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9135
      Left            =   0
      Top             =   0
      Width           =   20055
   End
End
Attribute VB_Name = "frmmisproyectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub cmdadd_Click()
Dim in_coordinador As String

If Len(Me.DtcPersonal.BoundText) > 2 Then

If Me.chk_coordinador.Value = 1 Then
   in_coordinador = "si"
Else
   in_coordinador = "no"
End If
strCadena = "sp_put_personal_proyecto('" & Val(Me.txt_idproyecto.Text) & "','" & Me.DtcPersonal.BoundText & "','" & Val(Me.txtsueldobase.Text) & "','" & in_coordinador & "')"
CnBd.Execute (strCadena)

Call load_personal(Val(Me.txt_idproyecto.Text), Me.HfPersonal)
End If
End Sub

Private Sub cmdcancelar_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub nuevo()
Me.txt_idproyecto.Text = ""
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA
Me.txtvalorizacion_inicial.Text = ""
Me.txtgastos.Text = ""
Me.txtcobros.Text = ""
Me.txtdescripcion.Text = ""


Me.frmpersonal.Visible = False

Me.frmdetalle.Visible = True








End Sub
Private Function validar() As Boolean
Dim in_mensaje As String




End Function

Private Sub cmdcerrarcobros_Click()
Me.frm_cobros.Visible = False
End Sub

Private Sub cmdcomprobantes_Click()

strCadena = "SELECT * FROM view_listado_comprobanteii WHERE id_proyecto='" & Val(Me.txt_idproyecto.Text) & "' and  ruc='" & KEY_RUC & "' "
Call llenar_cobros(Me.HfCobros, Val(Me.txt_idproyecto.Text))

End Sub

Private Sub cmddelete_Click()
Procedencia = Eliminar
frmsegurity.Show
End Sub

Private Sub cmddell_Click()
strCadena = "sp_delete_personal_proyec('" & Trim(Me.HfPersonal.TextMatrix(Me.HfPersonal.Row, 0)) & "','" & Val(Me.txt_idproyecto.Text) & "')"
CnBd.Execute (strCadena)
   Call load_personal(Val(Me.txt_idproyecto.Text), Me.HfPersonal)

End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdgastos_Click()

strCadena = "SELECT * FROM view_gastos_proyecto WHERE id_proyecto='" & Val(Me.txt_idproyecto.Text) & "' and  ruc='" & KEY_RUC & "'"

Call llenar_gastos(Me.HfCobros, Val(Me.txt_idproyecto.Text))

End Sub

Private Sub cmdnuevo_Click()
Call nuevo
End Sub
Private Sub load_proyecto(ByVal in_proyecto As String)

strCadena = "SELECT * FROM view_proyectos where id_proyecto='" & Val(in_proyecto) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txt_idproyecto.Text = rst("id_proyecto")
   Me.txtdescripcion.Text = rst("descripcion")
   Me.dtcCliente.BoundText = rst("id_cliente")
   Me.DtcContratista.BoundText = rst("id_contratista")
   Me.txtvalorizacion_inicial.Text = Format(rst("presupuesto"), "#,##0.00")
   Me.lblencargado.Caption = get_persona(rst("dni_save"))
   Me.txtcobros.Text = Format(rst("monto_cobrado"), "#,##0.00")
   Me.txtgastos.Text = Format(rst("monto_gasto"), "#,##0.00")
   Me.txtsaldo.Text = Format(rst("presupuesto") - rst("monto_cobrado"), "#,##0.00")
   Me.txtgarantia.Text = Format(rst("monto_garantia"), "#,##0.00")
   Call load_personal(rst("id_proyecto"), Me.HfPersonal)
   
   
   Me.DtcPersonal.BoundText = 0
   Me.frmdetalle.Visible = True
   
End If
End Sub
Private Sub load_personal(ByVal in_proyecto As String, ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Me.frmpersonal.Visible = True
strCadena = "SELECT * FROM  view_proyecto_personal WHERE id_proyecto='" & Val(in_proyecto) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
  
       Grilla.Rows = 0
      
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
        Next
        cabecera = "DNI" & vbTab & "PERSONAL" & vbTab & "SALARIO" & vbTab & "CORDINADOR"
        Grilla.AddItem cabecera
                            For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If rst("encargado") = "si" Then
               in_estado = "     [  X  ]"
            Else
               in_estado = ""
            End If
            
            Fila = rst("dni") & vbTab & rst("nombre_completo") & vbTab & Format(rst("salario"), "#,##0.00") & vbTab & in_estado
            Grilla.AddItem Fila
            
            If rst("encargado") = "si" Then
                For l = 0 To 3
                                Grilla.col = l
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
                Next l
            End If
            
            rst.MoveNext
    Next i
    
    
    
    
    
    
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub llenar_cobros(ByVal Grilla As MSHFlexGrid, ByVal in_proyecto As String)
Dim nsaldo As Double
Dim in_operador As String
On Error GoTo salir
Me.frm_cobros.Visible = True

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
           Grilla.ColWidth(2) = 2500
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 2500
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 1200
           Grilla.ColWidth(8) = 1200
        Next
         cabecera = "IDVENTA" & vbTab & "F.EMISION" & vbTab & "COMPROBANTE" & vbTab & "DNI CLIENTE" & vbTab & "CLIENTE" & vbTab & "TOTAL" & vbTab & "SALDO" & vbTab & "GARANTIA(5%)" & vbTab & "DETRACCION"
         Grilla.AddItem cabecera
         For k = 0 To 8
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        
        in_acumulado = 0
        in_garantia = 0
        in_detraccion = 0
        
        For i = 0 To rst.RecordCount - 1
        
             
             in_acumulado = in_acumulado + rst("total")
             in_garantia = in_garantia + rst("garantia")
             in_detraccion = in_detraccion + rst("detraccion")
             Fila = rst("id_venta") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & rst("comprobante") & vbTab & rst("id_cliente") & vbTab & rst("ncliente") & vbTab & rst("simbolo") & Space(3) & Format(rst("total"), "#,##0.00") & vbTab & rst("simbolo") & Space(3) & Format(rst("saldo"), "#,##0.00") & vbTab & rst("simbolo") & Space(1) & Format(rst("garantia"), "#,##0.00") & vbTab & rst("simbolo") & Space(1) & Format(rst("detraccion"), "#,##0.00")
             Grilla.AddItem Fila
             
                
                If rst("id_moneda") = "00002" Then
                    in_saldo = rst("saldo") * rst("tc")
                    in_factor = in_tseguro * rst("tc")
                Else
                    in_saldo = rst("saldo")
                    in_factor = in_tseguro
                End If
                nsaldo = nsaldo + in_saldo
                nfactor = nfactor + in_tseguro
            If rst("saldo") > 0 Then
            For k = 0 To 8
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H80C0FF
            Next k
            End If
            
            If rst("anulado") = "si" Then
           
            For k = 0 To 8
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
            Next k
            End If
            rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "SALDO A COBRAR:" & vbTab & "S/.  " & Format(in_acumulado, "#,##0.00") & vbTab & "S/.  " & Format(nsaldo, "#,##0.00") & vbTab & "S/.  " & Format(in_garantia, "#,##0.00") & vbTab & "S/.  " & Format(in_detraccion, "#,##0.00")
        Grilla.AddItem Fila
        For k = 4 To 8
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
            Next k
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub llenar_gastos(ByVal Grilla As MSHFlexGrid, ByVal in_proyecto As String)
Dim nsaldo As Double
Dim in_operador As String
On Error GoTo salir
Me.frm_cobros.Visible = True

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
           Grilla.ColWidth(2) = 2500
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 2500
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1200
        Next
         cabecera = "IDCOMPRA" & vbTab & "F.EMISION" & vbTab & "COMPROBANTE" & vbTab & "DNI CLIENTE" & vbTab & "CLIENTE" & vbTab & "TOTAL" & vbTab & "SALDO"
         Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        
        in_acumulado = 0
        For i = 0 To rst.RecordCount - 1
        
             
             in_acumulado = in_acumulado + rst("total")
             Fila = rst("id_compra") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & rst("comprobante") & vbTab & rst("id_proveedor") & vbTab & rst("nombre_completo") & vbTab & Format(rst("total"), "#,##0.00") & vbTab & Format(rst("saldo"), "#,##0.00")
             Grilla.AddItem Fila
             
                
                If rst("id_moneda") = "00002" Then
                    in_saldo = rst("saldo") * rst("tc")
                    in_factor = in_tseguro * rst("tc")
                Else
                    in_saldo = rst("saldo")
                    in_factor = in_tseguro
                End If
                nsaldo = nsaldo + in_saldo
                nfactor = nfactor + in_tseguro
            If rst("saldo") > 0 Then
            For k = 0 To 6
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H80C0FF
            Next k
            End If
            
            If rst("anulado") = "si" Then
           
            For k = 0 To 6
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
            Next k
            End If
            rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "SALDO A COBRAR:" & vbTab & "S/.  " & Format(in_acumulado, "#,##0.00") & vbTab & "S/.  " & Format(nsaldo, "#,##0.00")
        Grilla.AddItem Fila
        For k = 4 To 6
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
            Next k
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub




Private Sub cmdreporte_Click()
strCadena = "SELECT id_proyecto,fecha_inicio,fecha_fin,descripcion,cliente,presupuesto,fecha_emision,documento,id_linea,linea,id_producto,nombre_prod,cantidad,precio FROM view_proyecto_reporte WHERE id_proyecto='" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "'"
Call ConfiguraRst(strCadena)
 Ans = ShowMultiReport(rst, "RptLiquidacionObra", , App.Path + "\Reportes\")
End Sub

Private Sub cmdsave_Click()

strCadena = "sp_put_proyecto('" & Val(Me.txt_idproyecto.Text) & "','" & UCase(Trim(Me.txtdescripcion.Text)) & "','" & Me.DtcContratista.BoundText & "','" & Me.dtcCliente.BoundText & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & Val(Me.txtvalorizacion_inicial.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
Call llenarGrid(Me.HfdPersona)
Me.frmdetalle.Visible = False




End Sub
Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir

strCadena = "SELECT * FROM  view_proyectos WHERE ruc='" & KEY_RUC & "'"
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
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 3500
           Grilla.ColWidth(4) = 2800
           Grilla.ColWidth(5) = 2800
           Grilla.ColWidth(6) = 1500
           Grilla.ColWidth(7) = 1500
           Grilla.ColWidth(8) = 1500
           Grilla.ColWidth(9) = 1400
        Next
        cabecera = "CODIGO" & vbTab & "FE-INICIO" & vbTab & "FE-ENTREGA" & vbTab & "DESCRIPCION" & vbTab & "CLIENTE" & vbTab & "CONTRATISTA" & vbTab & "VALORIZACION" & vbTab & "T. GASTO" & vbTab & "T.FACTURADO" & vbTab & "ESTADO"
        Grilla.AddItem cabecera
                            For k = 0 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If rst("finalizado") = "si" Then
               in_estado = "FINALIZADO"
            Else
               in_estado = "PENDIENTE"
            End If
            
            Fila = Format(rst("id_proyecto"), "00000") & vbTab & Format(rst("fecha_inicio"), "dd-mm-YYYY") & vbTab & Format(rst("fecha_fin"), "dd-mm-YYYY") & vbTab & rst("descripcion") & vbTab & rst("cliente") & vbTab & rst("contratista") & vbTab & Format(rst("presupuesto"), "#,##0.00") & vbTab & Format(rst("monto_gasto"), "#,##.00") & vbTab & Format(rst("monto_cobrado"), "#,##0.00") & vbTab & in_estado
            Grilla.AddItem Fila
            
            If rst("finalizado") = "no" Then
                For l = 9 To 9
                                Grilla.col = l
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
                Next l
            End If
            
            rst.MoveNext
    Next i
    
    
    Me.HfdPersona.ColAlignment(4) = 1
    
    
    
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub cmdupdate_Click()
If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
    Call load_proyecto(Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)))
End If


End Sub

Private Sub DtcPersonal_Change()
strCadena = "SELECT sueldo FROM view_entidad WHERE dni='" & Me.DtcPersonal.BoundText & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   Me.txtsueldobase.Text = rstK("sueldo")
Else
   Me.txtsueldobase.Text = 0
End If
End Sub

Private Sub Form_Load()

CenterForm Me
Me.Top = 100
strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_cliente='si'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcCliente)

'----

strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcContratista)
Call llenarGrid(Me.HfdPersona)


strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_personal='si'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPersonal)
Me.DtcPersonal.BoundText = 0



Call llenarGrid(Me.HfdPersona)


End Sub

Private Sub txtcliente_Change()
strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and nombre_completo LIKE '%" & Trim(Me.txtcliente.Text) & "%'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.dtcCliente)
End Sub

Private Sub txtcontratista_Change()
strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "'  AND nombre_completo LIKE '%" & Trim(Me.txtcontratista.Text) & "%'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcContratista)
End Sub

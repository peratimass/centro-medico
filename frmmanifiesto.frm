VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmanifiesto 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17970
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   17970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmdetalle 
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
      Height          =   8340
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   17370
      Begin VB.Frame frmManifiestoDetalle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LISTADO DE COMPROBANTES"
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
         Height          =   4935
         Left            =   0
         TabIndex        =   46
         Top             =   3360
         Visible         =   0   'False
         Width           =   16095
         Begin VB.CheckBox chk_proforma 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "PROFORMA"
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
            Height          =   345
            Left            =   11520
            TabIndex        =   59
            Top             =   285
            Width           =   1215
         End
         Begin VB.CheckBox chk_vendedor 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "VENDEDOR :"
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
            Height          =   345
            Left            =   5160
            TabIndex        =   56
            Top             =   360
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk_sinmanifiesto 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "SIN MANIFIESTO:"
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
            Left            =   13440
            TabIndex        =   54
            Top             =   480
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin MSDataListLib.DataCombo DtcVendedor 
            Height          =   315
            Left            =   6480
            TabIndex        =   53
            Top             =   360
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
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
         Begin VB.CommandButton cmdCerrarDetalle 
            Height          =   255
            Left            =   15720
            Picture         =   "frmmanifiesto.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   240
            Width           =   255
         End
         Begin MSComCtl2.DTPicker DtpInicio 
            Height          =   315
            Left            =   1320
            TabIndex        =   48
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   110428161
            CurrentDate     =   42712
         End
         Begin MSComCtl2.DTPicker DtpFin 
            Height          =   315
            Left            =   3555
            TabIndex        =   50
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   110428161
            CurrentDate     =   42712
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfVentas 
            Height          =   3975
            Left            =   240
            TabIndex        =   51
            Top             =   840
            Width           =   15615
            _ExtentX        =   27543
            _ExtentY        =   7011
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
         Begin VitekeySoft.ChameleonBtn cmdbuscarmanifiesto 
            Height          =   420
            Left            =   10320
            TabIndex        =   55
            Top             =   285
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   741
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
            MICON           =   "frmmanifiesto.frx":2EA4
            PICN            =   "frmmanifiesto.frx":2EC0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA FIN:"
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
            Left            =   2745
            TabIndex        =   49
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   405
            TabIndex        =   47
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame frmdireccion 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2175
         Left            =   9000
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   8295
         Begin VB.CommandButton cmdcerrardireccion 
            Height          =   255
            Left            =   8000
            Picture         =   "frmmanifiesto.frx":54A5
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   120
            Width           =   255
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfdireccion 
            Height          =   1935
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   3413
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
      End
      Begin VB.TextBox txtcertificado_inscripcion 
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
         Left            =   9255
         TabIndex        =   41
         Text            =   " "
         Top             =   2940
         Width           =   2295
      End
      Begin VB.TextBox txtmarca 
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
         TabIndex        =   39
         Text            =   " "
         Top             =   2085
         Width           =   2055
      End
      Begin VB.TextBox txtmtc 
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
         TabIndex        =   37
         Text            =   " "
         Top             =   2940
         Width           =   2055
      End
      Begin VB.TextBox txtid_manifiesto 
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
         Left            =   5640
         TabIndex        =   36
         Text            =   " "
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtdireccion 
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   " "
         Top             =   1320
         Width           =   4935
      End
      Begin VB.CheckBox chk_direccion 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   14160
         TabIndex        =   31
         Top             =   1320
         Width           =   300
      End
      Begin VB.TextBox txtruc 
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
         Left            =   9120
         TabIndex        =   26
         Text            =   " "
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtpropietario 
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   " "
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox txtplaca 
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
         TabIndex        =   24
         Text            =   " "
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtbrevete 
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
         Left            =   1800
         TabIndex        =   23
         Text            =   " "
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtchofer 
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
         Left            =   1800
         TabIndex        =   21
         Text            =   " "
         Top             =   1680
         Width           =   5535
      End
      Begin VB.TextBox txtdni_chofer 
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
         Left            =   1800
         TabIndex        =   20
         Text            =   " "
         Top             =   1200
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DtoFecha 
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   110428161
         CurrentDate     =   42712
      End
      Begin VB.TextBox txtnumero 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   18
         Text            =   " "
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtanio 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         Text            =   " "
         Top             =   240
         Width           =   735
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   780
         Left            =   10560
         TabIndex        =   27
         Top             =   7320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1376
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
         MICON           =   "frmmanifiesto.frx":8349
         PICN            =   "frmmanifiesto.frx":8365
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdexit 
         Height          =   780
         Left            =   14880
         TabIndex        =   28
         Top             =   7320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1376
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
         MICON           =   "frmmanifiesto.frx":B8F7
         PICN            =   "frmmanifiesto.frx":B913
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfListado 
         Height          =   3735
         Left            =   1560
         TabIndex        =   29
         Top             =   3480
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   6588
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
      Begin VitekeySoft.ChameleonBtn cmdimprimir 
         Height          =   780
         Left            =   13800
         TabIndex        =   30
         Top             =   7320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1376
         BTYPE           =   5
         TX              =   "MAN.GUIAS"
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
         MICON           =   "frmmanifiesto.frx":BD03
         PICN            =   "frmmanifiesto.frx":BD1F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdDelete 
         Height          =   780
         Left            =   16320
         TabIndex        =   44
         Top             =   3480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1376
         BTYPE           =   5
         TX              =   "DELETE"
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
         MICON           =   "frmmanifiesto.frx":E2F0
         PICN            =   "frmmanifiesto.frx":E30C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdAgregar 
         Height          =   780
         Left            =   16320
         TabIndex        =   45
         Top             =   4320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1376
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
         MICON           =   "frmmanifiesto.frx":10756
         PICN            =   "frmmanifiesto.frx":10772
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdImprimirManifiestoventa 
         Height          =   780
         Left            =   12720
         TabIndex        =   57
         Top             =   7320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1376
         BTYPE           =   5
         TX              =   "MAN.VENTAS"
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
         MICON           =   "frmmanifiesto.frx":144B3
         PICN            =   "frmmanifiesto.frx":144CF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
         Height          =   780
         Left            =   11640
         TabIndex        =   58
         Top             =   7320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1376
         BTYPE           =   5
         TX              =   "LIQUIDACION"
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
         MICON           =   "frmmanifiesto.frx":16AA0
         PICN            =   "frmmanifiesto.frx":16ABC
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
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CERT.INSCRIPCION :"
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
         Left            =   7560
         TabIndex        =   42
         Top             =   3000
         Width           =   1605
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MARCA :"
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
         Left            =   4395
         TabIndex        =   40
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MTC :"
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
         Left            =   4515
         TabIndex        =   38
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE CHOFER :"
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
         Left            =   105
         TabIndex        =   22
         Top             =   1680
         Width           =   1545
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUC :"
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
         Left            =   8610
         TabIndex        =   16
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCION :"
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
         Left            =   8070
         TabIndex        =   15
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROPIETARIO :"
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
         Left            =   7860
         TabIndex        =   14
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLACA :"
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
         Left            =   4455
         TabIndex        =   13
         Top             =   2520
         Width           =   675
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BREVETE :"
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
         Left            =   825
         TabIndex        =   12
         Top             =   2160
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI CHOFER :"
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
         Left            =   525
         TabIndex        =   11
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA REGISTRO :"
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
         Left            =   165
         TabIndex        =   10
         Top             =   720
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° MANIFIESTO :"
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
         Left            =   285
         TabIndex        =   9
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.TextBox txtbuscar 
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
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   7830
      Width           =   3375
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3360
      Top             =   2670
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanifiesto.frx":1908D
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanifiesto.frx":194E1
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanifiesto.frx":19801
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanifiesto.frx":19C55
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanifiesto.frx":1A0A9
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanifiesto.frx":1A3C9
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanifiesto.frx":1A6E9
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanifiesto.frx":1AA09
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmanifiesto.frx":1AD29
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfmanifiesto 
      Height          =   7215
      Left            =   240
      TabIndex        =   1
      Top             =   390
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   12726
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
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   855
      Left            =   16560
      TabIndex        =   4
      Top             =   360
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
      MICON           =   "frmmanifiesto.frx":1B049
      PICN            =   "frmmanifiesto.frx":1B065
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
      Height          =   855
      Left            =   16560
      TabIndex        =   5
      Top             =   1215
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmmanifiesto.frx":1B4B7
      PICN            =   "frmmanifiesto.frx":1B4D3
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
      Height          =   855
      Left            =   16560
      TabIndex        =   6
      Top             =   2115
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
      MICON           =   "frmmanifiesto.frx":1B7ED
      PICN            =   "frmmanifiesto.frx":1B809
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdSalir 
      Height          =   855
      Left            =   16560
      TabIndex        =   7
      Top             =   3945
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
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
      MICON           =   "frmmanifiesto.frx":1BB23
      PICN            =   "frmmanifiesto.frx":1BB3F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdimprimirmani 
      Height          =   855
      Left            =   16560
      TabIndex        =   43
      Top             =   3000
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
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
      MICON           =   "frmmanifiesto.frx":1EB66
      PICN            =   "frmmanifiesto.frx":1EB82
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NUMERO :"
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
      Left            =   465
      TabIndex        =   3
      Top             =   7830
      Width           =   825
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO DE MANIFIESTO"
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
      Left            =   255
      TabIndex        =   2
      Top             =   120
      Width           =   2025
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   240
      Top             =   7710
      Width           =   16095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8415
      Left            =   0
      Top             =   0
      Width           =   17970
   End
End
Attribute VB_Name = "frmmanifiesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub ChameleonBtn1_Click()
strCadena = "SELECT id_manifiesto,fecha,numero,marca,placa,dni_chofer,'" & Trim(Me.txtchofer.Text) & "','" & Trim(Me.txtbrevete.Text) & "',direccion,documento,total,id_forma_pago,ncliente,dir_cliente FROM view_manifiesto_liquidacion WHERE id_manifiesto='" & Val(Me.txtid_manifiesto.Text) & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptManifiesto_liquidacion", , App.Path + "\Reportes\")


End Sub

Private Sub chk_direccion_Click()
If Me.chk_direccion.Value = 1 Then
   Call llenar_direccion(Me.hfdireccion, Trim(Me.txtRuc.Text))
   Me.frmdireccion.Visible = True
End If
End Sub

Private Sub cmdagregar_Click()
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA

If Me.chk_sinmanifiesto.Value = 1 Then
    strCadena = "SELECT * FROM view_venta_manifiesto WHERE id_manifiesto='0' and  id_doc IN('0003','0001') and  fecha_emision>='" & Format(KEY_FECHA, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(KEY_FECHA, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT * FROM view_venta_manifiesto WHERE id_doc IN('0003','0001') and  fecha_emision>='" & Format(KEY_FECHA, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(KEY_FECHA, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
End If

Call llenar_ventas(HfVentas, KEY_FECHA, KEY_FECHA)



End Sub

Private Sub cmdAnular_Click()
Procedencia = anular
frmsegurity.Show
Call disabled_form(Me)
Exit Sub
End Sub

Private Sub cmdBuscar_Click()
Call load_manifiesto(Me.hfmanifiesto.TextMatrix(Me.hfmanifiesto.Row, 0))
Call listo_guias(Me.HfListado, Val(Me.hfmanifiesto.TextMatrix(Me.hfmanifiesto.Row, 0)))
Call llenar_manifiesto_venta(Me.HfListado, Me.txtid_manifiesto.Text)
Me.frmdetalle.Visible = True
End Sub

Private Sub cmdbuscarmanifiesto_Click()
Dim in_vendedor As String

If Me.chk_vendedor.Value = 1 Then
    in_vendedor = Me.DtcVendedor.BoundText
Else
    in_vendedor = ""
End If


If Me.chk_sinmanifiesto.Value = 1 Then
    If Me.chk_proforma.Value = 1 Then
        strCadena = "SELECT * FROM view_venta_manifiesto WHERE  id_vendedor LIKE  '%" & in_vendedor & "%' and  id_manifiesto='0' and  id_doc IN('0099') and  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
    Else
        strCadena = "SELECT * FROM view_venta_manifiesto WHERE  id_vendedor LIKE  '%" & in_vendedor & "%' and  id_manifiesto='0' and  id_doc IN('0003','0001') and  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
    End If
    
Else
    If Me.chk_proforma.Value = 1 Then
        strCadena = "SELECT * FROM view_venta_manifiesto WHERE id_vendedor LIKE '%" & in_vendedor & "%' and id_doc IN('0099') and  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
    Else
        strCadena = "SELECT * FROM view_venta_manifiesto WHERE id_vendedor LIKE '%" & in_vendedor & "%' and id_doc IN('0003','0001') and  fecha_emision>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
    End If
    
    
End If

Call llenar_ventas(HfVentas, Me.DtpInicio.Value, Me.DtpFin.Value)


End Sub

Private Sub cmdcerrardetalle_Click()

Call llenar_manifiesto_venta(Me.HfListado, Me.txtid_manifiesto.Text)
frmManifiestoDetalle.Visible = False
End Sub

Private Sub cmdcerrardireccion_Click()
Me.frmdireccion.Visible = False
End Sub

Private Sub cmdexit_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub cmdImprimir_Click()

strCadena = "SELECT id_manifiesto,id_anio,id_numero,fecha,ruc_propietario,'PROPIETARIO',direccion_propietario,'CHOFER','LICENCIA',marca,placa,distrito,documento,guia,remitente,destinatario,cantidad_total,peso_total,total,'1','2','3','4','5','6',ruc FROM view_manifiesto_produccion WHERE id_manifiesto='" & Val(Me.txtid_manifiesto.Text) & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptManifiesto", , App.Path + "\Reportes\")


'Call impresion_manifiesto(Me.txtid_manifiesto.Text)
End Sub

Private Sub cmdimprimirmani_Click()
Dim in_manifiesto As String
Dim in_propietario As String
Dim in_chofer As String
Dim in_licencia As String

in_manifiesto = Me.hfmanifiesto.TextMatrix(Me.hfmanifiesto.Row, 0)
strCadena = "SELECT * FROM transferencia_manifiesto WHERE id_manifiesto='" & Val(in_manifiesto) & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   in_propietario = get_persona(rst("ruc_propietario"))
   in_chofer = get_persona(rst("dni_chofer"))
   in_licencia = get_licencia(rst("dni_chofer"))
End If
strCadena = "SELECT id_manifiesto,id_anio,id_numero,fecha,ruc_propietario,'" & in_propietario & "',direccion_propietario,'" & in_chofer & "','" & in_licencia & "',marca,placa,distrito,documento,guia,remitente,destinatario,cantidad_total,peso_total,total,'1','2','3','4','5','6',ruc FROM view_manifiesto_produccion WHERE id_manifiesto='" & Val(in_manifiesto) & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptManifiesto", , App.Path + "\Reportes\")
'll impresion_manifiesto(Me.hfmanifiesto.TextMatrix(Me.hfmanifiesto.Row, 0))
End Sub



Private Sub cmdImprimirManifiestoventa_Click()

strCadena = "CALL ADM_manifiesto('" & Val(txtid_manifiesto.Text) & "','" & KEY_RUC & "')"

'strCadena = "SELECT id_manifiesto,fecha,numero,marca,placa,dni_chofer,'" & Trim(Me.txtchofer.Text) & "','" & Trim(Me.txtbrevete.Text) & "',direccion,linea,id_producto,nombre_prod,cantidad,unidad FROM view_manifiesto_reporte_v2 WHERE id_manifiesto='" & Val(Me.txtid_manifiesto.Text) & "'"
Call ConfiguraRst(strCadena)

Ans = ShowMultiReport(rst, "RptManifiesto_venta", , App.Path + "\Reportes\")



End Sub

Private Sub cmdNuevo_Click()
Call nuevo
End Sub
Public Sub load_manifiesto(ByVal in_manifiesto As String)
strCadena = "SELECT * FROM view_manifiesto where id_manifiesto='" & Val(in_manifiesto) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtid_manifiesto.Text = rst("id_manifiesto")
   Me.txtAnio.Text = rst("id_anio")
   Me.txtNumero.Text = rst("id_numero")
   Me.DtoFecha.Value = rst("fecha")
   Me.TxtPlaca.Text = rst("placa")
   Me.txtmtc.Text = rst("placa2")
   Me.txtcertificado_inscripcion.Text = rst("certificado")
   Me.TxtMarca.Text = rst("marca")
   Me.txtRuc.Text = rst("ruc_propietario")
   Me.txtpropietario.Text = rst("propietario")
   Me.txtDireccion.Text = rst("direccion")
   
   Call Me.BuscarChofer(rst("dni_chofer"))
   
End If
End Sub
Private Sub nuevo()
Me.txtid_manifiesto.Text = ""
Me.txtAnio.Text = Year(KEY_FECHA)
Me.txtNumero.Text = get_manifiesto(Year(KEY_FECHA))
Me.txtdni_chofer.Text = ""
Me.txtchofer.Text = ""
Me.TxtPlaca.Text = ""
Me.txtmtc.Text = ""
Me.TxtMarca.Text = ""
Me.txtbrevete.Text = ""
Me.txtpropietario.Text = ""
Me.txtDireccion.Text = ""
Me.txtRuc.Text = ""
Call get_propietario(KEY_RUC)
Me.frmdetalle.Visible = True
End Sub



Private Sub cmdProcesar_Click()
If Val(Me.txtid_manifiesto.Text) < 1 Then
    
    strCadena = "call p_inserta_manifiesto('" & Trim(Year(KEY_FECHA)) & "','" & Trim(Me.txtNumero.Text) & "','" & Format(Me.DtoFecha.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtRuc.Text) & "','" & Trim(Me.txtDireccion.Text) & "','" & Trim(Me.txtdni_chofer.Text) & "','" & Trim(Me.txtbrevete.Text) & "','" & Trim(Me.TxtPlaca.Text) & "','" & Trim(Me.txtmtc.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.TxtMarca.Text) & "','" & Trim(Me.txtcertificado_inscripcion.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Call llenar_manifiesto(Me.hfmanifiesto)
    Me.frmdetalle.Visible = False
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
Me.DtoFecha.Value = KEY_FECHA

strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcVendedor)


Call llenar_manifiesto(Me.hfmanifiesto)



End Sub
Private Function get_manifiesto(ByVal in_anio As String) As String
strCadena = "SELECT * FROM transferencia_manifiesto WHERE id_anio='" & Trim(in_anio) & "' and  ruc='" & KEY_RUC & "' ORDER BY id_numero DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   get_manifiesto = Format(Val(rst("id_numero")) + 1, "000000")
Else
   get_manifiesto = Format(1, "000000")
End If

End Function

Public Sub llenar_manifiesto(ByVal Grilla As MSHFlexGrid)
Dim tTotal As Double, ccostos As String
strCadena = "SELECT * FROM view_manifiesto WHERE  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 3000
           Grilla.ColWidth(4) = 2500
           Grilla.ColWidth(5) = 2500
           Grilla.ColWidth(6) = 1300
           Grilla.ColWidth(7) = 1300
           Grilla.ColWidth(8) = 1300
       Next
        cabecera = "IDMANIFIESTO" & vbTab & "FECHA" & vbTab & "N° MANIFIESTO" & vbTab & "PROPIETARIO" & vbTab & "DIRECCION" & vbTab & "CHOFER" & vbTab & "BREVETE" & vbTab & "PLACA N°1" & vbTab & " PLACA N°2"
        Grilla.AddItem cabecera
         For k = 1 To 8
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
          
          Fila = rst("id_manifiesto") & vbTab & Format(rst("fecha"), "YYYY-mm-dd") & vbTab & rst("id_anio") & "-" & rst("id_numero") & vbTab & rst("propietario") & vbTab & rst("direccion") & vbTab & rst("chofer") & vbTab & rst("licencia") & vbTab & rst("placa") & vbTab & rst("placa2")
          Grilla.AddItem Fila
          
          
          rst.MoveNext
      Next i
        
       
      
End Sub

Public Sub llenar_ventas(ByVal Grilla As MSHFlexGrid, ByVal in_fecha_ini As Date, ByVal in_fecha_fin As Date)
Dim tTotal As Double, ccostos As String

frmManifiestoDetalle.Visible = True


Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 3000
           Grilla.ColWidth(4) = 3000
           Grilla.ColWidth(5) = 1500
           Grilla.ColWidth(6) = 2000
           Grilla.ColWidth(7) = 500
       Next
        cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "CLIENTE" & vbTab & "VENDEDOR" & vbTab & "TOTAL" & vbTab & "MANIFIESTO" & vbTab & "ESTADO"
        Grilla.AddItem cabecera
         For k = 1 To 7
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
          If rst("manifiesto") = "-" Then
            in_estado = Chr(168)
          Else
            in_estado = Chr(254)
          End If
          Fila = rst("id_venta") & vbTab & Format(rst("fecha_emision"), "YYYY-mm-dd") & vbTab & rst("documento") & vbTab & rst("ncliente") & vbTab & rst("vendedor") & vbTab & Format(rst("total"), "#,##0.00") & vbTab & rst("manifiesto") & vbTab & in_estado
          Grilla.AddItem Fila
          
           With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 7 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
            End With
           ' If rst("manifiesto") = "-" Then
           '     For k = 1 To 6
           '         Grilla.col = k
           '         Grilla.Row = i + 1
           '         Grilla.CellBackColor = &H8080FF
           '     Next k
           ' End If
            
          
          
          
          rst.MoveNext
      Next i
        
       
      
End Sub


Public Sub llenar_manifiesto_venta(ByVal Grilla As MSHFlexGrid, ByVal in_manifiesto As String)
Dim tTotal As Double, ccostos As String



strCadena = "SELECT * FROM view_venta_manifiesto WHERE id_doc IN('0003','0001','0007') and  id_manifiesto='" & Val(in_manifiesto) & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 3000
           Grilla.ColWidth(4) = 3000
           Grilla.ColWidth(5) = 1500
           Grilla.ColWidth(6) = 2000
           Grilla.ColWidth(7) = 500
       Next
        cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "CLIENTE" & vbTab & "VENDEDOR" & vbTab & "TOTAL" & vbTab & "MANIFIESTO" & vbTab & "ESTADO"
        Grilla.AddItem cabecera
         For k = 1 To 7
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
          If rst("manifiesto") = "-" Then
            in_estado = Chr(168)
          Else
            in_estado = Chr(254)
          End If
          Fila = rst("id_venta") & vbTab & Format(rst("fecha_emision"), "YYYY-mm-dd") & vbTab & rst("documento") & vbTab & rst("ncliente") & vbTab & rst("vendedor") & vbTab & Format(rst("total"), "#,##0.00") & vbTab & rst("manifiesto") & vbTab & in_estado
          Grilla.AddItem Fila
          
           With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 7 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
            End With
            If rst("manifiesto") = "-" Then
                For k = 1 To 6
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
            End If
            
          
          
          
          rst.MoveNext
      Next i
        
       
      
End Sub


Public Sub listo_guias(ByVal Grilla As MSHFlexGrid, ByVal in_manifiesto As String)
Dim tTotal As Double, ccostos As String
strCadena = "SELECT * FROM view_manifiesto_produccion WHERE id_manifiesto='" & in_manifiesto & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 1500
           Grilla.ColWidth(1) = 1800
           Grilla.ColWidth(2) = 1300
           Grilla.ColWidth(3) = 2400
           Grilla.ColWidth(4) = 2400
           Grilla.ColWidth(5) = 800
           Grilla.ColWidth(6) = 1100
           Grilla.ColWidth(7) = 1100
           
       Next
        cabecera = "DESTINO" & vbTab & "FACTURA/BOLETA" & vbTab & "GUIA REMISION" & vbTab & "REMITENTE " & vbTab & " DESTINATARIO" & vbTab & "BULTO" & vbTab & "PESO(KG)" & vbTab & "FLETE"
        Grilla.AddItem cabecera
         For k = 0 To 7
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
          
          Fila = UCase(rst("distrito")) & vbTab & rst("documento") & vbTab & rst("guia") & vbTab & rst("remitente") & vbTab & rst("destinatario") & vbTab & rst("cantidad_total") & vbTab & rst("peso_total") & vbTab & rst("total")
          Grilla.AddItem Fila
          
          
          rst.MoveNext
      Next i
        
       
      
End Sub
Public Sub get_propietario(ByVal in_dni As String)

strCadena = "SELECT * FROM persona where dni='" & in_dni & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtRuc.Text = rst("dni")
   Me.txtpropietario.Text = rst("nombre_completo")
   Me.txtDireccion.Text = rst("direccion")
Else
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
End If

End Sub
Public Sub BuscarChofer(ByVal ruc As String)
If (Trim(ruc) = "") Then
    Procedencia = seleccionar_per
    FrmPersona.Show
    Exit Sub
End If
   
buscar_nuevamente:
    strCadena = "SELECT *  FROM persona WHERE dni='" & ruc & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        If get_dni_reniec(ruc) = True Then
            GoTo buscar_nuevamente
        End If
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.txtRuc.Text = ruc
        FrmDetallePersona.ChkTransporte.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.txtdni_chofer.Text = rst("dni")
        Me.txtchofer.Text = rst("nombre_completo")
        If IsNull(rst("licencia")) = True Then
            Me.txtbrevete.Text = ""
        Else
            Me.txtbrevete.Text = rst("licencia")
        End If
        
         Call Resalta(Me.txtbrevete)
        Exit Sub
       
    End If

End Sub

Private Sub hfdireccion_Click()
Call select_direccion
End Sub
Private Sub select_direccion()
Me.hfdireccion.TextMatrix(Me.hfdireccion.Row, 3) = Chr(254)
Me.txtDireccion.Text = Me.hfdireccion.TextMatrix(Me.hfdireccion.Row, 2)
Me.frmdireccion.Visible = False
End Sub




Private Sub put_manifiesto(ByVal in_venta As String, ByVal in_manifiesto As String)

strCadena = "SELECT id_manifiesto FROM movimiento_venta WHERE id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   If Val(rst("id_manifiesto")) < 1 Then
        
        strCadena = "UPDATE movimiento_venta SET id_manifiesto='" & Val(in_manifiesto) & "' WHERE id_venta='" & Val(in_venta) & "'"
        CnBd.Execute (strCadena)
        
        Me.HfVentas.TextMatrix(Me.HfVentas.Row, 6) = Trim(Me.txtAnio.Text) & "-" & Trim(Me.txtNumero.Text)
        Me.HfVentas.TextMatrix(Me.HfVentas.Row, 7) = Chr(254)
        
    Else
        If rst("id_manifiesto") <> Val(Me.txtid_manifiesto.Text) Then
        MsgBox "YA ESTA ASIGNADA A UN MANIFIESTO", vbInformation
        If MsgBox("DESDE ASIGNARLE ESTE MANIFIESTO", vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then
            strCadena = "UPDATE movimiento_venta SET id_manifiesto='" & Val(in_manifiesto) & "' WHERE id_venta='" & Val(in_venta) & "'"
            CnBd.Execute (strCadena)
            Me.HfVentas.TextMatrix(Me.HfVentas.Row, 6) = Trim(Me.txtAnio.Text) & "-" & Trim(Me.txtNumero.Text)
            Me.HfVentas.TextMatrix(Me.HfVentas.Row, 7) = Chr(254)
        End If
        Else
            strCadena = "UPDATE movimiento_venta SET id_manifiesto='0' WHERE id_venta='" & Val(in_venta) & "'"
            CnBd.Execute (strCadena)
            Me.HfVentas.TextMatrix(Me.HfVentas.Row, 6) = "-"
            Me.HfVentas.TextMatrix(Me.HfVentas.Row, 7) = Chr(168)
        End If
   End If
    
End If

End Sub

Private Sub HfVentas_DblClick()
Call put_manifiesto(Val(Me.HfVentas.TextMatrix(Me.HfVentas.Row, 0)), Val(Me.txtid_manifiesto.Text))
End Sub

Private Sub txtdni_chofer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call BuscarChofer(Me.txtdni_chofer.Text)
End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call get_propietario(Trim(Me.txtRuc.Text))
    
    
    Exit Sub
End If
End Sub

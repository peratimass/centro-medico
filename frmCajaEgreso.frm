VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCajaEgreso 
   BorderStyle     =   0  'None
   Caption         =   "CAJA EGRESO"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtComprobante 
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
      Left            =   6600
      TabIndex        =   88
      Top             =   680
      Width           =   1215
   End
   Begin VB.TextBox txtid_moneda 
      Height          =   285
      Left            =   18960
      TabIndex        =   87
      Text            =   "idmoneda"
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtbuscardni 
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
      Left            =   10800
      TabIndex        =   71
      Top             =   680
      Width           =   1335
   End
   Begin VB.TextBox TxtBuscaEntidad 
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
      Left            =   14640
      TabIndex        =   68
      Top             =   680
      Width           =   2535
   End
   Begin VB.TextBox txtOperacionBusqueda 
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
      Left            =   10800
      TabIndex        =   62
      Top             =   300
      Width           =   1335
   End
   Begin VB.TextBox txtMontoBusqueda 
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
      Left            =   6600
      TabIndex        =   60
      Top             =   300
      Width           =   1215
   End
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
      Height          =   7455
      Left            =   840
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   17415
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "COMPROBANTES RELACIONADOS"
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
         Height          =   3375
         Left            =   8880
         TabIndex        =   75
         Top             =   3120
         Width           =   8415
         Begin VB.Frame frmmonto_nuevo 
            BackColor       =   &H00404040&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   855
            Left            =   4560
            TabIndex        =   79
            Top             =   960
            Visible         =   0   'False
            Width           =   2535
            Begin VB.TextBox txtMontoNuevo 
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
               Left            =   960
               TabIndex        =   81
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label lblid_temporal 
               Caption         =   "Label24"
               Height          =   135
               Left            =   240
               TabIndex        =   82
               Top             =   240
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Image Image1 
               Height          =   240
               Left            =   2200
               Picture         =   "frmCajaEgreso.frx":0000
               Top             =   60
               Width           =   240
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MONTO :"
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
               Left            =   240
               TabIndex        =   80
               Top             =   480
               Width           =   600
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HgPagados 
            Height          =   2775
            Left            =   240
            TabIndex        =   76
            Top             =   360
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4895
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
         Begin VitekeySoft.ChameleonBtn cmdEliminarPago 
            Height          =   855
            Left            =   7440
            TabIndex        =   77
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1508
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
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCajaEgreso.frx":2EA4
            PICN            =   "frmCajaEgreso.frx":2EC0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdModificar_pago 
            Height          =   855
            Left            =   7440
            TabIndex        =   78
            Top             =   1320
            Width           =   855
            _ExtentX        =   1508
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
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCajaEgreso.frx":530A
            PICN            =   "frmCajaEgreso.frx":5326
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
      Begin VB.TextBox txtid 
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
         Left            =   6120
         TabIndex        =   53
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtTipoFlujo 
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
         Left            =   6480
         TabIndex        =   52
         Top             =   4920
         Width           =   975
      End
      Begin VB.TextBox txtFormaPago 
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
         Left            =   6480
         TabIndex        =   51
         Top             =   4320
         Width           =   975
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "COMPROBANTE"
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
         Height          =   2055
         Left            =   8880
         TabIndex        =   39
         Top             =   960
         Width           =   8415
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfContenido 
            Height          =   735
            Left            =   1440
            TabIndex        =   46
            Top             =   1080
            Width           =   6735
            _ExtentX        =   11880
            _ExtentY        =   1296
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
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MONTO :"
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
            Left            =   555
            TabIndex        =   45
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F.EMISION :"
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
            Left            =   6015
            TabIndex        =   44
            Top             =   360
            Width           =   780
         End
         Begin VB.Label lblmonto 
            Caption         =   " "
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
            Left            =   1440
            TabIndex        =   43
            Top             =   720
            Width           =   3750
         End
         Begin VB.Label lblfecha_emision 
            Caption         =   " "
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
            Left            =   6960
            TabIndex        =   42
            Top             =   285
            Width           =   1365
         End
         Begin VB.Label lblcomprobante 
            Caption         =   " "
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
            Height          =   405
            Left            =   1440
            TabIndex        =   41
            Top             =   240
            Width           =   3765
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COMPROBANTE :"
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
            Left            =   30
            TabIndex        =   40
            Top             =   360
            Width           =   1125
         End
      End
      Begin VB.TextBox TxtOperacion 
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
         Left            =   2385
         TabIndex        =   37
         Top             =   5520
         Width           =   2175
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
         Height          =   675
         Left            =   2400
         TabIndex        =   31
         Top             =   6000
         Width           =   5055
      End
      Begin VB.TextBox txtrucproveedor 
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
         TabIndex        =   20
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtcambio 
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
         Left            =   6120
         TabIndex        =   18
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtMonto 
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
         TabIndex        =   1
         Top             =   1320
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo DtcMoneda 
         Height          =   330
         Left            =   2400
         TabIndex        =   15
         Top             =   855
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
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   795
         Left            =   15000
         TabIndex        =   22
         Top             =   6600
         Width           =   1095
         _ExtentX        =   1931
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
         MICON           =   "frmCajaEgreso.frx":5640
         PICN            =   "frmCajaEgreso.frx":565C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdimprimir 
         Height          =   795
         Left            =   16200
         TabIndex        =   23
         Top             =   6600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1402
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
         MICON           =   "frmCajaEgreso.frx":8CA4
         PICN            =   "frmCajaEgreso.frx":8CC0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcTipomovimiento 
         Height          =   330
         Left            =   2400
         TabIndex        =   25
         Top             =   2880
         Width           =   3015
         _ExtentX        =   5318
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
      Begin MSDataListLib.DataCombo DtcCuentaOrigen 
         Height          =   330
         Left            =   2400
         TabIndex        =   27
         Top             =   3360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
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
      Begin MSDataListLib.DataCombo DtcCuentaDestino 
         Height          =   330
         Left            =   2400
         TabIndex        =   29
         Top             =   3840
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
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
      Begin VitekeySoft.ChameleonBtn cmdsalir 
         Height          =   315
         Left            =   16920
         TabIndex        =   36
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   ""
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
         MICON           =   "frmCajaEgreso.frx":B291
         PICN            =   "frmCajaEgreso.frx":B2AD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcFormaPago 
         Height          =   330
         Left            =   2400
         TabIndex        =   49
         Top             =   4320
         Width           =   3975
         _ExtentX        =   7011
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
      Begin MSDataListLib.DataCombo DtcFlujo 
         Height          =   330
         Left            =   2400
         TabIndex        =   50
         Top             =   4920
         Width           =   3975
         _ExtentX        =   7011
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
      Begin MSMask.MaskEdBox TxtFecha_emision 
         Height          =   345
         Left            =   2400
         TabIndex        =   0
         ToolTipText     =   "dd/mm/yyyy"
         Top             =   360
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
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
      Begin VB.Label lblOperador 
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
         Left            =   2400
         TabIndex        =   86
         Top             =   6840
         Width           =   5100
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OPERADOR :"
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
         Left            =   1200
         TabIndex        =   85
         Top             =   6840
         Width           =   825
      End
      Begin VB.Label txtid_recibo 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   9000
         TabIndex        =   84
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblid_recibo 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   9000
         TabIndex        =   83
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO DE FLUJO :"
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
         Left            =   1065
         TabIndex        =   48
         Top             =   5040
         Width           =   1050
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FORMA DE PAGO :"
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
         Left            =   915
         TabIndex        =   47
         Top             =   4440
         Width           =   1200
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NRO OPERACION :"
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
         Left            =   915
         TabIndex        =   38
         Top             =   5520
         Width           =   1200
      End
      Begin VB.Label lblhora_registro 
         Caption         =   " "
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
         Left            =   12240
         TabIndex        =   35
         Top             =   600
         Width           =   2100
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H.REGISTRO :"
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
         Left            =   11160
         TabIndex        =   34
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lblfecha_registro 
         Caption         =   " "
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
         Left            =   12240
         TabIndex        =   33
         Top             =   240
         Width           =   2100
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.REGISTRO :"
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
         Left            =   11160
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label12 
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
         Left            =   1080
         TabIndex        =   30
         Top             =   6240
         Width           =   1035
      End
      Begin VB.Label lbldestino 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUENTA DESTINO :"
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
         TabIndex        =   28
         Top             =   3885
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUENTA ORIGEN :"
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
         Left            =   945
         TabIndex        =   26
         Top             =   3450
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO MOVIMIENTO  :"
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
         Left            =   750
         TabIndex        =   24
         Top             =   3000
         Width           =   1365
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   2400
         X2              =   7560
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label lblproveedor 
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
         Left            =   2400
         TabIndex        =   21
         Top             =   2280
         Width           =   5220
      End
      Begin VB.Label lblproveedorn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROVEEDOR :"
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
         Left            =   1215
         TabIndex        =   19
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T.CAMBIO :"
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
         Left            =   5280
         TabIndex        =   17
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO  :"
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
         Left            =   1485
         TabIndex        =   16
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONEDA   :"
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
         Left            =   1365
         TabIndex        =   14
         Top             =   915
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA OPERACION :"
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
         Left            =   780
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
   End
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   300
      Left            =   1080
      TabIndex        =   8
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
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
      CalendarBackColor=   8421631
      CalendarForeColor=   8388608
      CalendarTitleForeColor=   8388608
      Format          =   50331649
      CurrentDate     =   42975
   End
   Begin VB.TextBox txtIdcuenta 
      Height          =   375
      Left            =   18960
      TabIndex        =   3
      Text            =   "idcuenta"
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Hfdetalle 
      Height          =   7935
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   18495
      _ExtentX        =   32623
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
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9.75
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   9000
      Top             =   2550
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
            Picture         =   "frmCajaEgreso.frx":E161
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajaEgreso.frx":E5B5
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajaEgreso.frx":E8D5
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajaEgreso.frx":ED29
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajaEgreso.frx":F17D
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajaEgreso.frx":F49D
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajaEgreso.frx":F7BD
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajaEgreso.frx":FADD
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajaEgreso.frx":FDFD
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   18840
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "NUEVA "
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
      MICON           =   "frmCajaEgreso.frx":1011D
      PICN            =   "frmCajaEgreso.frx":10139
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmddetalle 
      Height          =   855
      Left            =   18840
      TabIndex        =   5
      Top             =   1950
      Width           =   1095
      _ExtentX        =   1931
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCajaEgreso.frx":1058B
      PICN            =   "frmCajaEgreso.frx":105A7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrar 
      Height          =   855
      Left            =   18840
      TabIndex        =   6
      Top             =   7140
      Width           =   1095
      _ExtentX        =   1931
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCajaEgreso.frx":12800
      PICN            =   "frmCajaEgreso.frx":1281C
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
      Left            =   2640
      TabIndex        =   10
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
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
      CalendarBackColor=   8421631
      CalendarForeColor=   8388608
      CalendarTitleForeColor=   8388608
      Format          =   50331649
      CurrentDate     =   42975
   End
   Begin VitekeySoft.ChameleonBtn cmdBuscar 
      Height          =   345
      Left            =   4080
      TabIndex        =   11
      Top             =   300
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "   BUSCAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "frmCajaEgreso.frx":15843
      PICN            =   "frmCajaEgreso.frx":1585F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdReporte 
      Height          =   855
      Left            =   18840
      TabIndex        =   54
      Top             =   2820
      Width           =   1095
      _ExtentX        =   1931
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCajaEgreso.frx":17D8A
      PICN            =   "frmCajaEgreso.frx":17DA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrarCaja 
      Height          =   855
      Left            =   18840
      TabIndex        =   56
      Top             =   4560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "CERRAR CAJA"
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
      MICON           =   "frmCajaEgreso.frx":1A377
      PICN            =   "frmCajaEgreso.frx":1A393
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdExtornar 
      Height          =   855
      Left            =   18840
      TabIndex        =   57
      Top             =   6280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "EXTORNAR"
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
      MICON           =   "frmCajaEgreso.frx":1A7E5
      PICN            =   "frmCajaEgreso.frx":1A801
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdeditar 
      Height          =   855
      Left            =   18840
      TabIndex        =   58
      Top             =   5420
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCajaEgreso.frx":1AB1B
      PICN            =   "frmCajaEgreso.frx":1AB37
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdBuscarMonto 
      Height          =   345
      Left            =   7920
      TabIndex        =   63
      Top             =   300
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "   BUSCAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "frmCajaEgreso.frx":1DE0D
      PICN            =   "frmCajaEgreso.frx":1DE29
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdBuscarOperacion 
      Height          =   345
      Left            =   12240
      TabIndex        =   64
      Top             =   300
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "   BUSCAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "frmCajaEgreso.frx":20354
      PICN            =   "frmCajaEgreso.frx":20370
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdBuscartipomov 
      Height          =   345
      Left            =   17280
      TabIndex        =   65
      Top             =   300
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "   BUSCAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "frmCajaEgreso.frx":2289B
      PICN            =   "frmCajaEgreso.frx":228B7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcTipomov 
      Height          =   330
      Left            =   14640
      TabIndex        =   67
      Top             =   300
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VitekeySoft.ChameleonBtn cmdBuscarEntidad 
      Height          =   345
      Left            =   17280
      TabIndex        =   69
      Top             =   680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "   BUSCAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "frmCajaEgreso.frx":24DE2
      PICN            =   "frmCajaEgreso.frx":24DFE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdbuscardni 
      Height          =   345
      Left            =   12240
      TabIndex        =   73
      Top             =   680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "   BUSCAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "frmCajaEgreso.frx":27329
      PICN            =   "frmCajaEgreso.frx":27345
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdOrigen_pago 
      Height          =   855
      Left            =   18840
      TabIndex        =   74
      Top             =   3690
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "ORIGEN PAGO"
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
      MICON           =   "frmCajaEgreso.frx":29870
      PICN            =   "frmCajaEgreso.frx":2988C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdBuscarNumero 
      Height          =   345
      Left            =   7920
      TabIndex        =   89
      Top             =   680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "   BUSCAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
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
      MICON           =   "frmCajaEgreso.frx":2CA90
      PICN            =   "frmCajaEgreso.frx":2CAAC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPROBANTE :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5475
      TabIndex        =   90
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI/RUC :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   9960
      TabIndex        =   72
      Top             =   765
      Width           =   630
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTIDAD :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   13800
      TabIndex        =   70
      Top             =   720
      Width           =   630
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO MOV:"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   13800
      TabIndex        =   66
      Top             =   360
      Width           =   645
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OPERACION :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   9840
      TabIndex        =   61
      Top             =   300
      Width           =   825
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONTO :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6000
      TabIndex        =   59
      Top             =   300
      Width           =   525
   End
   Begin VB.Label lblcuenta 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   360
      TabIndex        =   55
      Top             =   675
      Width           =   4905
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   915
      Left            =   240
      Top             =   120
      Width           =   18495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AL"
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
      Left            =   2445
      TabIndex        =   9
      Top             =   360
      Width           =   165
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " FECHAS :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   300
      Width           =   630
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "frmCajaEgreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub DtcCaja_Click(Area As Integer)
'Call actualizar
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DtcCuentas_Change()
'Call llenarGrid(Me.HfDetalle, Me)
End Sub

Private Sub CmdRecientes_Click()
Call recientes(Val(Me.TxtidCuenta.Text), Me.DtpInicio.Value, Me.DtpFin.Value, get_moneda_cuenta(Me.TxtidCuenta.Text))
End Sub
Public Sub recientes(ByVal id_cuenta As Double, ByVal in_fecha_ini As Date, ByVal in_fecha_fin As Date, ByVal in_moneda As String)
Dim Saldo As Double

strCadena = "select SUM(montoreal) FROM mis_cuentas_det WHERE  anulado='no' and id_cuenta='" & id_cuenta & "'  AND fecha<'" & Format(in_fecha_ini, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If IsNull(rst(0)) = True Then
    Saldo = 0
Else
    Saldo = rst(0)
End If



strCadena = "SELECT * FROM view_detalle_caja WHERE  id_cuenta='" & id_cuenta & "' AND ruc='" & KEY_RUC & "' AND date(fecha)>='" & Format(in_fecha_ini, "YYYY-mm-dd") & "'AND date(fecha)<='" & Format(in_fecha_fin, "YYYY-mm-dd") & "'"
If in_moneda = "00001" Then
    
   Call llenarGrid_mn(Me.HfDetalle, Val(Me.TxtidCuenta.Text), Saldo)
Else
    Call llenarGrid_me(Me.HfDetalle, Val(Me.TxtidCuenta.Text), Saldo)
End If

End Sub

Private Sub CmdTodos_Click()
Dim Saldo As Double
strCadena = "select SUM(montoreal) FROM mis_cuentas_det WHERE id_cuenta='" & id_cuenta & "' AND anulado='no' AND ruc='" & KEY_RUC & "' AND fecha_sys<'2000-01-01'"
Call ConfiguraRst(strCadena)
If IsNull(rst(0)) = True Then
    Saldo = 0
Else
    Saldo = rst(0)
End If
Set rst = Nothing

strCadena = "SELECT id,fecha_sys,fecha,documento,P.nombre_completo,G.plan_des,monto,montoreal,M.ccostos,M.anulado FROM mis_cuentas_det M,persona P,plan_contable_det G " & _
"WHERE M.ccostos=G.pc_codigo AND  M.id_persona=P.dni AND id_cuenta='" & Val(Me.TxtidCuenta.Text) & "' AND M.ruc='" & KEY_RUC & "' AND M.dni_save='" & KEY_USUARIO & "' ORDER BY id ASC"
Call llenarGrid_mn(Me.HfDetalle, Val(Me.TxtidCuenta.Text), Saldo)
End Sub

Private Sub Command3_Click()

End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub cmdBuscar_Click()
Call recientes(Val(Me.TxtidCuenta.Text), Me.DtpInicio.Value, Me.DtpFin.Value, get_moneda_cuenta(Me.TxtidCuenta.Text))
End Sub

Private Sub cmdbuscardni_Click()
Dim Saldo As Double

strCadena = "select SUM(montoreal) FROM mis_cuentas_det WHERE  id_persona = '" & Trim(Me.TxtBuscarDNI.Text) & "' and anulado='no' and id_cuenta='" & Val(Me.TxtidCuenta.Text) & "'  AND fecha<'" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If IsNull(rst(0)) = True Then
    Saldo = 0
Else
    Saldo = rst(0)
End If


strCadena = "SELECT * FROM view_detalle_caja WHERE  id_persona = '" & Trim(Me.TxtBuscarDNI.Text) & "' and id_cuenta='" & Val(Me.TxtidCuenta.Text) & "' AND ruc='" & KEY_RUC & "' AND date(fecha)>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "'AND date(fecha)<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "'"
Call llenarGrid_mn(Me.HfDetalle, Val(Me.TxtidCuenta.Text), Saldo)

End Sub

Private Sub cmdBuscarEntidad_Click()
Dim Saldo As Double

strCadena = "select SUM(montoreal) FROM mis_cuentas_det WHERE  npersona LIKE '%" & Trim(Me.TxtBuscaEntidad.Text) & "%' and anulado='no' and id_cuenta='" & Val(Me.TxtidCuenta.Text) & "'  AND fecha<'" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If IsNull(rst(0)) = True Then
    Saldo = 0
Else
    Saldo = rst(0)
End If


strCadena = "SELECT * FROM view_detalle_caja WHERE  npersona LIKE '%" & Trim(Me.TxtBuscaEntidad.Text) & "%' and id_cuenta='" & Val(Me.TxtidCuenta.Text) & "' AND ruc='" & KEY_RUC & "' AND date(fecha)>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "'AND date(fecha)<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "'"
Call llenarGrid_mn(Me.HfDetalle, Val(Me.TxtidCuenta.Text), Saldo)

End Sub

Private Sub cmdBuscarMonto_Click()
Dim Saldo As Double

strCadena = "select SUM(montoreal) FROM mis_cuentas_det WHERE  monto LIKE '%" & Val(Me.txtMontoBusqueda.Text) & "%' and anulado='no' and id_cuenta='" & Val(Me.TxtidCuenta.Text) & "'  AND fecha<'" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If IsNull(rst(0)) = True Then
    Saldo = 0
Else
    Saldo = rst(0)
End If


strCadena = "SELECT * FROM view_detalle_caja WHERE  monto LIKE '%" & Val(Me.txtMontoBusqueda.Text) & "%' and id_cuenta='" & Val(Me.TxtidCuenta.Text) & "' AND ruc='" & KEY_RUC & "' AND date(fecha)>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "'AND date(fecha)<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "'"
Call llenarGrid_mn(Me.HfDetalle, Val(Me.TxtidCuenta.Text), Saldo)
End Sub

Private Sub cmdBuscarNumero_Click()
Dim Saldo As Double

strCadena = "select SUM(montoreal) FROM mis_cuentas_det WHERE  documento LIKE '%" & Trim(Me.txtComprobante.Text) & "%' and anulado='no' and id_cuenta='" & Val(Me.TxtidCuenta.Text) & "'  AND fecha<'" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If IsNull(rst(0)) = True Then
    Saldo = 0
Else
    Saldo = rst(0)
End If


strCadena = "SELECT * FROM view_detalle_caja WHERE  documento LIKE '%" & Trim(Me.txtComprobante.Text) & "%' and id_cuenta='" & Val(Me.TxtidCuenta.Text) & "' AND ruc='" & KEY_RUC & "' AND date(fecha)>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "'AND date(fecha)<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "'"
Call llenarGrid_mn(Me.HfDetalle, Val(Me.TxtidCuenta.Text), Saldo)
End Sub

Private Sub cmdBuscarOperacion_Click()
Dim Saldo As Double


strCadena = "select SUM(montoreal) FROM mis_cuentas_det WHERE  operacion LIKE '%" & Trim(Me.TxtOperacion.Text) & "%' and anulado='no' and id_cuenta='" & Val(Me.TxtidCuenta.Text) & "'  AND fecha<'" & Format(in_fecha_ini, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If IsNull(rst(0)) = True Then
    Saldo = 0
Else
    Saldo = rst(0)
End If


strCadena = "SELECT * FROM view_detalle_caja WHERE  operacion LIKE '%" & Trim(Me.TxtOperacion.Text) & "%' and  id_cuenta='" & Val(Me.TxtidCuenta.Text) & "' AND ruc='" & KEY_RUC & "' "
Call llenarGrid_mn(Me.HfDetalle, Val(Me.TxtidCuenta.Text), Saldo)
End Sub

Private Sub cmdBuscartipomov_Click()
Dim Saldo As Double

strCadena = "select SUM(montoreal) FROM mis_cuentas_det WHERE  id_tipo_movimiento = '" & Me.DtcTipomov.BoundText & "' and anulado='no' and id_cuenta='" & Val(Me.TxtidCuenta.Text) & "'  AND fecha<'" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If IsNull(rst(0)) = True Then
    Saldo = 0
Else
    Saldo = rst(0)
End If


strCadena = "SELECT * FROM view_detalle_caja WHERE  id_tipo_movimiento LIKE '%" & Me.DtcTipomov.BoundText & "%' and id_cuenta='" & Val(Me.TxtidCuenta.Text) & "' AND ruc='" & KEY_RUC & "' AND date(fecha)>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "'AND date(fecha)<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "'"
Call llenarGrid_mn(Me.HfDetalle, Val(Me.TxtidCuenta.Text), Saldo)

End Sub

Private Sub cmdCerrar_Click()
FrmMiscuentas.actualizar
Unload Me
End Sub

Public Sub get_movimiento(ByVal in_id As String)
Dim in_cuenta As Integer
Dim in_recibo As String
Dim in_operador As String
strCadena = "SELECT * FROM mis_cuentas_det where id='" & Val(in_id) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   in_cuenta = rst("id_cuenta")
   Me.TxtFecha_emision.Text = CVDate(rst("fecha"))
   Me.txtMonto.Text = rst("monto")
   Me.txtrucproveedor.Text = rst("id_persona")
   Me.DtcTipomovimiento.BoundText = rst("id_tipo_movimiento")
   Me.DtcCuentaOrigen.BoundText = in_cuenta
   Me.DtcCuentaDestino.BoundText = rst("id_cuenta_destino")
   Me.TxtOperacion.Text = rst("operacion")
   Me.txtObservacion.Text = rst("glosa")
   Me.lblhora_registro.Caption = rst("hora_registro")
   Me.lblfecha_registro.Caption = rst("fecha_sys")
   Me.lblcomprobante.Caption = rst("documento")
   Me.lblfecha_emision.Caption = rst("fecha")
   Me.lblmonto.Caption = rst("monto")
   Me.txtcambio.Text = rst("tc")
   in_operador = rst("dni_save")
   Me.lblProveedor.Caption = get_persona(Trim(Me.txtrucproveedor.Text))
   Me.DtcFormaPago.BoundText = rst("id_forma_pago")
   Me.DtcFlujo.BoundText = rst("id_tipo_flujo")
   Me.txtid.Text = Val(in_id)
   in_recibo = rst("id_venta")
   Call llenar(Me.HfContenido, rst("id_venta"), rst("id_compra"))
   If in_recibo > 0 And Me.DtcTipomovimiento.BoundText = "00002" Then
      Call put_listar_pagos_temporal(in_recibo)
   End If
   Me.lblOperador.Caption = get_persona(in_operador)
   
   Me.cmdProcesar.Enabled = False
   Me.frmdetalle.Visible = True
End If


End Sub

Private Sub put_listar_pagos_temporal(ByVal in_recibo As String)
strCadena = "SELECT * FROM view_movimiento_compra_pagos WHERE id_recibo='" & in_recibo & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.lblid_recibo.Caption = in_recibo
   strCadena = "call put_delete_pagos_temporal('" & KEY_USUARIO & "','" & KEY_RUC & "')"
   CnBd.Execute (strCadena)
   
   For i = 0 To rst.RecordCount - 1
       strCadena = "INSERT INTO movimiento_compra_pago_temporal(id_factura,monto_pago,id_recibo,dni_save,ruc)VALUES " & _
       "('" & rst("id_compra") & "','" & rst("monto_pagado") & "','" & rst("id_recibo") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       rst.MoveNext
   Next i
 Call Llenar_Temporal(Me.HgPagados, in_recibo)
   
End If
End Sub

Sub llenar(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double, ByVal idCompra As String)
On Error GoTo salir
Dim tTotal As Double
Dim texonerado As Double
Dim tafecto As Double


If idVenta > 0 Then
    strCadena = "SELECT * FROM movimiento_venta_detalle  WHERE id_venta='" & idVenta & "'"
Else
    If idCompra = 0 Then
        Exit Sub
    End If
    strCadena = "SELECT * FROM movimiento_compra_detalle WHERE id_compra='" & idCompra & "'"
End If



Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
   
    Grilla.Rows = 0
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 500
           Grilla.ColWidth(3) = 600
           Grilla.ColWidth(4) = 700
           
        Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION " & vbTab & "CANT" & vbTab & "PRECIO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
     
        For i = 0 To rst.RecordCount - 1
            
            
                Fila = rst("id_producto") & vbTab & UCase(rst("detalle")) & vbTab & rst("cantidad") & vbTab & Format(rst("total") / rst("cantidad"), "###0.00") & vbTab & Format(rst("total"), "###0.00")
            
            
            Grilla.AddItem Fila
            rst.MoveNext
    Next i
  


Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Sub Llenar_Temporal(ByVal Grilla As MSHFlexGrid, ByVal in_recibo As Double)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT * FROM view_movimiento_compra_pagos_temp WHERE id_recibo='" & in_recibo & "' and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
   
    Grilla.Rows = 0
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 3500
           Grilla.ColWidth(2) = 1600
           
           
        Next
        cabecera = "CODIGO" & vbTab & "COMPROBANTE " & vbTab & "MONTO PAGADO"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        in_monto_pagado = 0
        For i = 0 To rst.RecordCount - 1
            in_monto_pagado = in_monto_pagado + rst("monto_pago")
            Fila = rst("id_compra") & vbTab & rst("comprobante") & vbTab & rst("monto_pago")
            
            Grilla.AddItem Fila
            rst.MoveNext
    Next i
    cabecera = "" & vbTab & "TOTAL PAGADO:" & vbTab & Format(in_monto_pagado, "###0.0000")
        Grilla.AddItem cabecera
    For k = 0 To 2
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H80FF&
        Next k
       If Me.cmdProcesar.Enabled = True Then
          Me.txtMonto.Text = in_monto_pagado
       End If
        
        
        


Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Private Sub cmdCerrarCaja_Click()
Procedencia = cerrarcaja
Call disabled_form(FrmMiscuentas)
Call disabled_form(Me)
FrmSeguridad.Show
Exit Sub

End Sub

Private Sub cmdDetalle_Click()

Call get_movimiento(Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)))
Me.cmdEliminarPago.Enabled = False
Me.cmdModificar_pago.Enabled = False
End Sub


Private Sub cmdEliminar_Click()

End Sub

Private Sub cmdEditar_Click()
Procedencia = modificar
Call enabled_form(frmCajaEgreso)
Call enabled_form(FrmMiscuentas)
frmsegurity.Show
Exit Sub
End Sub

Private Sub cmdEliminarPago_Click()
If Val(Me.HgPagados.TextMatrix(Me.HgPagados.Row, 0)) > 0 Then
   
   strCadena = "DELETE FROM movimiento_compra_pago_temporal WHERE id_factura='" & Val(Me.HgPagados.TextMatrix(Me.HgPagados.Row, 0)) & "' and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
   Call Me.Llenar_Temporal(Me.HgPagados, Val(Me.lblid_recibo.Caption))
   
End If
End Sub

Private Sub cmdExtornar_Click()

frmsegurity.Show
Call disabled_form(Me)
Procedencia = extornar
Exit Sub

End Sub

Private Sub cmdImprimir_Click()
If Me.DtcTipomovimiento.BoundText = "00003" Then
strCadena = "SELECT id,fecha,id_tipo_movimiento,flujo,medio_pago,operacion,sucursal,operador,glosa,ruc FROM view_transferencia  WHERE id='" & Val(Me.txtid.Text) & "'"
Call ConfiguraRst(strCadena)
strCadena = "SELECT descripcion,simbolo,monto,tc,id_moneda,operador FROM view_transferencia_reporte WHERE id_origen='" & Val(Me.txtid.Text) & "'"
Call ConfiguraRstK(strCadena)
Ans = ShowMultiReport(rst, "RptVoucher_transferencia", , App.Path + "\Reportes\", , , , , rstK, "RptVoucher_transferencia_detalle")
Exit Sub
End If

If Me.DtcTipomovimiento.BoundText = "00002" Then
strCadena = "SELECT id,fecha,id_tipo_movimiento,flujo,medio_pago,operacion,sucursal,operador,glosa,ruc,tc,id_persona,npersona,monto,monto_letras FROM view_transferencia  WHERE id='" & Val(Me.txtid.Text) & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptVoucher_egreso", , App.Path + "\Reportes\")


End If



End Sub

Private Sub cmdModificar_pago_Click()
If Val(Me.HgPagados.TextMatrix(Me.HgPagados.Row, 0)) > 0 Then
   
   Me.frmmonto_nuevo.Visible = True
   strCadena = "SELECT * FROM movimiento_compra_pago_temporal WHERE id_factura='" & Val(Me.HgPagados.TextMatrix(Me.HgPagados.Row, 0)) & "' and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      Me.lblid_temporal.Caption = rst("id_factura")
      Me.txtMontoNuevo.Text = rst("monto_pago")
   End If
   
   Call Resalta(Me.txtMontoNuevo)
   
   Exit Sub
End If
End Sub

Private Sub cmdnuevo_Click()

Me.txtid.Text = 0
Me.txtMonto.Text = ""
Me.txtrucproveedor.Text = ""
Me.lblProveedor.Caption = ""
Me.DtcTipomovimiento.BoundText = "00001"
Me.TxtOperacion.Text = ""
Me.txtObservacion.Text = ""
Me.lblcomprobante.Caption = ""
Me.lblmonto.Caption = ""
Me.lblfecha_emision.Caption = ""

    
    Me.TxtFecha_emision.Mask = ""
    Me.TxtFecha_emision.Text = ""
    Me.TxtFecha_emision.Mask = "##/##/####"
    Me.TxtFecha_emision.Text = CVDate(KEY_FECHA)
    
    
 
    
Me.cmdProcesar.Enabled = True
Me.cmdImprimir.Enabled = False
Me.DtcMoneda.Locked = True
Me.DtcMoneda.BoundText = get_moneda_cuenta(Me.TxtidCuenta.Text)



Me.frmdetalle.Visible = True
Me.TxtFecha_emision.TabIndex = 0
Call Resalta(Me.txtMonto)


End Sub

Private Sub cmdOrigen_pago_Click()
       Dim in_recibo As String
       Dim in_total As String
       
       strCadena = "SELECT * FROM view_comprobante_origen WHERE id='" & Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) & "' LIMIT 1"
       Call ConfiguraRstT(strCadena)
       If rstT.RecordCount > 0 Then
            in_recibo = rstT("id_venta")
            in_total = rstT("total")
       Else
            Exit Sub
       End If
        
        
strCadena = "SELECT id_venta,fecha_emision,hora,documento,id_cliente,ncliente,total,forma_pago,flujo,cuenta_origen,observacion,nombre_completo,tc,'" & UCase(EnLetras(in_total)) & "',operacion,monto_redondeo,monto_anticipo,cta_redondeo,cta_anticipo,ruc FROM view_compra_pago WHERE id_venta='" & in_recibo & "'"
Call ConfiguraRst(strCadena)

strCadena = "SELECT fecha_emision,id_proveedor,nproveedor,comprobante,tc,id_moneda,monto_inicial,monto_pagado FROM view_compra_pago_detalle WHERE id_venta='" & in_recibo & "'"
Call ConfiguraRstK(strCadena)
Ans = ShowMultiReport(rst, "RptVoucher_pago", , App.Path + "\Reportes\", , , , , rstK, "RptVoucher_pago_detalle")

End Sub
Private Sub cmdProcesar_Click()

If get_periodo_cierre(get_periodo_actual(Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd")), "caja") = True Then
        
        MsgBox "PERIODO DE COMPRAS CERRARDO.!!!", vbInformation, KEY_VENDEDOR
        Exit Sub
        
     End If


If verificar_cierre_caja(Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd")) = 1 Then
    MsgBox "AVISO IMPORTANTE..." + Chr(13) + Chr(13) + "CAJA CONTABLE YA CERRADA.", vbInformation, KEY_VENDEDOR
    Exit Sub
End If
'-

If Val(Me.txtid.Text) < 1 Then
    Me.txtid.Text = procesar_transaccion_caja(KEY_ALM, Me.DtcCuentaOrigen.BoundText, Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd"), Me.DtcTipomovimiento.BoundText, Trim(Me.txtrucproveedor.Text), Trim(Me.lblProveedor.Caption), Trim(Me.txtObservacion.Text), Val(Me.txtMonto.Text), Me.DtcCuentaDestino.BoundText, "0", "0", Trim(lblcomprobante.Caption), Val(Me.txtcambio.Text), Trim(Me.TxtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
    
Else
    
    If Val(Me.lblid_recibo.Caption) > 0 Then
        Call delete_recibo_egreso(Me.lblid_recibo.Caption, Me.txtid.Text)
       
    Else
        strCadena = "UPDATE mis_cuentas_det SET fecha='" & Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd") & "',glosa='" & UCase(Trim(Me.txtObservacion.Text)) & "' WHERE id='" & Val(Me.txtid.Text) & "'"
        CnBd.Execute (strCadena)
    
        strCadena = "UPDATE con_movimientocajabanco SET Fecha='" & Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd") & "', Glosa='" & UCase(Trim(Me.txtObservacion.Text)) & "' WHERE NroImpresion = '" & Val(Me.txtid.Text) & "'"
        CnBd.Execute (strCadena)
    End If
    
    
    
End If



MsgBox "Proceso Realizado con Exito", vbInformation, KEY_VENDEDOR

Call recientes(Val(Me.TxtidCuenta.Text), Me.DtpInicio.Value, Me.DtpFin.Value, get_moneda_cuenta(Me.TxtidCuenta.Text))
Me.cmdProcesar.Enabled = False
Me.cmdImprimir.Enabled = True
End Sub

Private Sub delete_recibo_egreso(ByVal in_recibo As String, ByVal in_id_detalle_det As String)

strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & Val(in_recibo) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    
    Call AnularVentas(Trim(rst("id_doc")), rst("serie"), rst("numero"), rst("id_alm"))
              
    If KEY_CONTABILIDAD = "si" Then
       strCadena = "call CON_InsertaAsiento_PagoGlobal_Extorno('" & Val(in_recibo) & "')"
       CnBd.Execute (strCadena)
    End If
    Call save_modificar(in_recibo)
End If


End Sub
Private Sub save_modificar(ByVal in_recibo As String)

Dim monto_pago As Double, saldof As Double, comprobante As String, monto_pagado As Double, Saldo As Double, id_moneda As String
Dim in_registro As Double
Dim in_tipo As String

strCadena = "SELECT * FROM movimiento_compra_pago_temporal WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   rstL.MoveFirst
   For i = 0 To rstL.RecordCount - 1
    strCadena = "UPDATE movimiento_compra SET monto_pagar='" & rstL("monto_pago") & "' , dni_save_pago='" & KEY_USUARIO & "' , seleccion='si' WHERE id_compra='" & rstL("id_factura") & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    rstL.MoveNext
   Next i
End If




saldof = 0
       
        'actualizado todo por pagar
        
        monto_pago = Val(Me.txtMonto.Text)
        in_recibo = generar_recibo_update
        
        
        monto_pagado = Val(Me.txtMonto.Text)
                    
                    strCadena = "SELECT `id_compra`,`fecha_emision`,`fecha_cancelacion`,`comprobante`,`id_proveedor`,`nproveedor`,id_moneda,`simbolo`,`moneda`,`tc`,`total`,`saldo`,`nombre_completo`,`id_alm`,`ruc`, function_pago_factura(id_compra,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc) as pago,seleccion,monto_pagar FROM view_cuentas_cobrar WHERE dni_save_pago='" & KEY_USUARIO & "' and seleccion='si' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRst(strCadena)
                    If rst.RecordCount > 0 Then
                        rst.MoveFirst
                        
                        For i = 0 To rst.RecordCount - 1
                           in_saldo = rst("monto_pagar")
                           in_saldo_inicial = in_saldo
                           
                           If rst("id_moneda") = "00002" Then
                              If rst("id_moneda") = Me.DtcMoneda.BoundText Then
                                 saldof = in_saldo
                              Else
                                saldof = in_saldo * Val(Me.txtcambio.Text)
                              End If
                               
                           Else
                              If rst("id_moneda") = Me.DtcMoneda.BoundText Then
                                 saldof = in_saldo
                              Else
                                saldof = in_saldo / Val(Me.txtcambio.Text)
                              End If
                              
                            End If
                            
                            If monto_pago > saldof Then
                               
                               If KEY_ASIENTO_GLOBAL_CTA_PAGAR = "no" Then
                                    Call procesar_transaccion(KEY_ALM, Me.DtcCuentaOrigen.BoundText, Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd"), "00002", rst("id_proveedor"), rst("nproveedor"), in_glosa, saldof, "0", "0", rst("id_compra"), rst("comprobante"), Val(Me.txtcambio.Text), Trim(Me.TxtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                               End If
                                    strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & Val(Me.txtid_recibo.Caption) & "','" & rst("id_compra") & "','" & in_saldo_inicial & "','" & saldof & "','" & rst("id_moneda") & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtcambio.Text) & "')"
                                    CnBd.Execute (strCadena)
                                    monto_pago = monto_pago - saldof
                                    strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & Val(Me.txtid_recibo.Caption) & "','00','" & Trim(rst("comprobante")) & "','-','1','" & saldof & "','0','" & saldof & "','" & KEY_RUC & "')"
                                    CnBd.Execute (strCadena)
                            Else
                                
                                in_glosa = "[" & UCase(Trim(Me.txtObservacion.Text)) & "]"
                                
                                If KEY_ASIENTO_GLOBAL_CTA_PAGAR = "no" Then
                                    Call procesar_transaccion(KEY_ALM, Me.DtcCuentaOrigen.BoundText, Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd"), "00002", rst("id_proveedor"), rst("nproveedor"), in_glosa, monto_pago, "0", "0", rst("id_compra"), rst("comprobante"), Val(Me.txtcambio.Text), Trim(Me.TxtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                                End If
                                strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & Val(Me.txtid_recibo.Caption) & "','" & rst("id_compra") & "','" & in_saldo_inicial & "','" & monto_pago & "','" & rst("id_moneda") & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.txtcambio.Text) & "')"
                                CnBd.Execute (strCadena)
                                
                                strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,referencia,cantidad,precio,peso,total,ruc) VALUES ('" & Val(Me.txtid_recibo.Caption) & "','00','" & Trim(rst("comprobante")) & "','-','1','" & in_saldo_inicial & "','0','" & in_saldo_inicial & "','" & KEY_RUC & "')"
                                CnBd.Execute (strCadena)
                                monto_pago = 0
                                GoTo siguiente
                            End If
                        rst.MoveNext
                        Next i
                          
                          
siguiente:
                          
                          If KEY_ASIENTO_GLOBAL_CTA_PAGAR = "si" Then
                            
                            strCadena = "call CON_InsertaAsiento_PagoGlobal('" & Val(Me.txtid_recibo.Caption) & "')"
                            CnBd.Execute (strCadena)
                            Call procesar_transaccion_egreso(Me.DtcFormaPago.BoundText, KEY_ALM, Me.DtcCuentaOrigen.BoundText, Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd"), "00002", Trim(Me.txtrucproveedor.Text), Trim(Me.lblProveedor.Caption), Trim(Me.txtObservacion.Text), Val(Me.txtMonto.Text), "", 0, Val(txtid_recibo.Caption), in_recibo, Val(Me.txtcambio.Text), Trim(Me.TxtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, KEY_USUARIO, KEY_RUC)
                            
                            
                          
                          End If


                      
                    End If
                    
        Exit Sub
    

End Sub
Private Function generar_recibo_update() As String
                    KEY_VENCIMIENTO = KEY_FECHA
                    id_tipo_factura = "0002"
                    igv = "si"
                    dfac = "no"
                    
                    strCadena = "SELECT numero,serie FROM  movimiento_venta WHERE id_doc='0097' and id_alm='" & KEY_ALM & "'  and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
                    Call ConfiguraRstZ(strCadena)
                    If rstZ.RecordCount > 0 Then
                        in_serie = rstZ("serie")
                        in_numero = Format(Val(rstZ("numero")) + 1, "000000")
                    Else
                        in_serie = "001"
                        in_numero = Format(1, "000000")
                    End If
                    
                    horario = Format(Time, "hh:mm")
                    If horario >= "07:00" And horario <= "13:00" Then
                        turno = "M"
                    Else
                        turno = "T"
                    End If
                    Documento = "RECIBO EGRESO:" & ":" & in_serie & "-" & in_numero
                    strCadena = "P_insert_venta('0097','" & KEY_ALM & "','" & get_forma_pago(Me.DtcCuentaOrigen.BoundText) & "','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
                    "'" & in_serie & "','" & in_numero & "','" & Trim(Me.txtrucproveedor.Text) & "','" & Me.lblProveedor.Caption & "','0','0','0','" & Val(Me.txtMonto.Text) & "','0'," & _
                    "'" & Val(Me.txtMonto.Text) & "','0','" & Format(TxtFecha_emision.Text, "YYYY-mm-dd") & "','" & Format(Me.TxtFecha_emision.Text, "YYYY-mm-dd") & "','" & id_tipo_factura & "','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & Val(Me.txtcambio.Text) & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Documento & "','" & horario & "','" & turno & "','--','" & KEY_RUC & "')"
                    Call ConfiguraRstP(strCadena)
                    
                    id_venta = rstP(0)
                    Me.txtid_recibo.Caption = id_venta
                  
                    
                    strCadena = "UPDATE movimiento_venta SET  observacion='" & Trim(Me.txtObservacion.Text) & "',operacion='" & Trim(Me.TxtOperacion.Text) & "' WHERE id_venta='" & id_venta & "'"
                    CnBd.Execute (strCadena)
                    
                    
                   
                               
                               
                    
                    strCadena = "INSERT INTO movimiento_venta_monto(id_venta,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,banco,cheque,cuenta_contable,forma_pago_contable,flujo_caja,id_cuenta_origen,cta_redondeo,cta_anticipo,monto_redondeo,monto_anticipo,ruc)VALUES " & _
                    "('" & id_venta & "','01','" & get_forma_pago_anterior(Me.DtcMoneda.BoundText) & "','" & Val(Me.txtMonto.Text) & "','" & Val(Me.txtMonto.Text) * -1 & "','00','-','" & Trim(Me.TxtOperacion.Text) & "','-','0','" & get_cuenta_contable_cuenta(Me.DtcCuentaOrigen.BoundText) & "','" & DtcFormaPago.BoundText & "','" & Me.DtcFlujo.BoundText & "','" & Me.DtcCuentaOrigen.BoundText & "','0','0','0','0','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    
                    strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(in_numero + 1), "000000") & "' WHERE id_doc='0097' AND serie='" & Trim(in_serie) & "' AND ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    generar_recibo_update = "RECIBO EGRESO:" & Trim(in_serie) & "-" & in_numero
                    


      
      Exit Function

End Function


Private Sub cmdReporte_Click()

strCadena = "select SUM(montoreal) FROM mis_cuentas_det WHERE id_cuenta='" & Val(Me.TxtidCuenta.Text) & "' AND fecha<'" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If IsNull(rst(0)) = True Then
    Saldo = 0
Else
    Saldo = rst(0)
End If


strCadena = "SELECT '" & Trim(Me.lblcuenta.Caption) & "','" & Format(Me.DtpInicio.Value, "dd-mm-YYYY") & "','" & Format(Me.DtpFin.Value, "dd-mm-YYYY") & "',fecha,'" & Saldo & "',id_tipo_movimiento,tipo_movimiento,id_persona,left(npersona,20),glosa,nombre_completo,montoreal,'" & Mid(KEY_VENDEDOR, 1, 20) & "',operacion FROM view_detalle_caja WHERE  id_cuenta='" & Val(Me.TxtidCuenta.Text) & "' AND ruc='" & KEY_RUC & "' AND date(fecha)>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "'AND date(fecha)<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptCuenta_detalle", , App.Path + "\Reportes\")
   
End Sub

Private Sub cmdSalir_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub DtcCuentaDestino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcFormaPago.SetFocus
End If
End Sub

Private Sub DtcCuentaOrigen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcCuentaDestino.Visible = True Then
        Me.DtcCuentaDestino.SetFocus
    Else
        Me.DtcFormaPago.SetFocus
    End If
End If
End Sub

Private Sub DtcFlujo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtOperacion)
End If
End Sub

Private Sub DtcFormaPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcFlujo.SetFocus
End If
End Sub

Private Sub DtcTipomovimiento_Change()
If Me.DtcTipomovimiento.BoundText = "00003" Then
   strCadena = "SELECT id_cuenta as Codigo, CONCAT(descripcion,'-',numero_cuenta,'    [',moneda,']') as Descripcion FROM view_cuenta_banco WHERE id_cuenta<>'" & Me.DtcCuentaOrigen.BoundText & "' and  ruc='" & KEY_RUC & "' ORDER BY id_cuenta"
   Call ConfiguraRstT(strCadena)
   Call LlenaDataComboT(Me.DtcCuentaDestino)
   Me.DtcCuentaDestino.Visible = True
   Me.lbldestino.Visible = True
   Me.lblproveedorn.Visible = False
   Me.txtrucproveedor.Visible = False
Else
   Me.DtcCuentaDestino.Visible = False
   Me.lbldestino.Visible = False
   Me.lblproveedorn.Visible = True
   Me.txtrucproveedor.Visible = True
End If
End Sub

Private Sub DtcTipomovimiento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcCuentaDestino.Visible = True Then
        Me.DtcCuentaDestino.SetFocus
    Else
        Me.DtcFormaPago.SetFocus
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 27) Then
    Unload Me
End If
End Sub

Private Function get_moneda(ByVal in_cuenta As String) As String

strCadena = "SELECT id_moneda FROM mis_cuentas WHERE id_cuenta='" & Val(in_cuenta) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtid_moneda.Text = rst("id_moneda")
End If
End Function
Private Sub Form_Load()
CenterForm Me
Me.Top = 100
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA


strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM moneda ORDER BY id_moneda"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)

strCadena = "SELECT id_mov as Codigo, descripcion as Descripcion FROM tipo_movimiento_caja ORDER BY id_mov"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipomovimiento)

Me.TxtidCuenta.Text = Val(FrmMiscuentas.HfgDetalle.TextMatrix(FrmMiscuentas.HfgDetalle.Row, 0))

strCadena = "SELECT id_cuenta as Codigo, CONCAT(descripcion,'-',numero_cuenta,'    [',moneda,']') as Descripcion FROM view_cuenta_banco WHERE ruc='" & KEY_RUC & "' ORDER BY id_cuenta"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcCuentaOrigen)
Me.DtcCuentaOrigen.BoundText = Val(Me.TxtidCuenta.Text)
strCadena = "SELECT id as Codigo,Descripcion  as Descripcion FROM vw_mediopago_nombre  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFormaPago)

strCadena = "SELECT id as Codigo,Nombre  as Descripcion FROM adm_flujocaja ORDER BY Nombre  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFlujo)

strCadena = "select id_mov as Codigo, descripcion as Descripcion from tipo_movimiento_caja ORDER BY id_mov"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipomov)




Me.txtcambio.Text = Format(KEY_CAMBIO_COMPRA, "#,##.0000")




Call recientes(Val(Me.TxtidCuenta.Text), KEY_FECHA, KEY_FECHA, get_moneda_cuenta(Me.TxtidCuenta.Text))
End Sub
Private Sub Transferencia(ByVal in_tipo As String)

If in_tipo = "00003" Then ' transferencia
   Me.lbldestino.Visible = True
   Me.DtcCuentaDestino.Visible = True
Else
   Me.lbldestino.Visible = False
   Me.DtcCuentaDestino.Visible = False
End If

End Sub
Private Sub llenarGrid_me(ByVal Grilla As MSHFlexGrid, ByVal id_cuenta As Double, ByVal Saldo As Double)
On Error GoTo salir
Dim tTotal As Double, denominacion As String
Grilla.Font.Size = 8
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
           Grilla.Rows = 0
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 3000
           Grilla.ColWidth(5) = 3000
           Grilla.ColWidth(6) = 1200
           
           Grilla.ColWidth(7) = 1200
           Grilla.ColWidth(8) = 1200
           Grilla.ColWidth(9) = 1300
           
           Grilla.ColWidth(10) = 1200
           Grilla.ColWidth(11) = 1200
           Grilla.ColWidth(12) = 1300
           
        cabecera = "ID_DETALLE" & vbTab & "EMISION" & vbTab & "MOVIMIENTO" & vbTab & "DNI/RUC" & vbTab & "NOMBRE/RAZON SOCIAL" & vbTab & "CONCEPTO" & vbTab & "OPERADOR" & vbTab & "DEBE" & vbTab & "HABER" & vbTab & "SALDO"
        Grilla.AddItem cabecera
         For k = 1 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
        cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "============   SALDO INICIAL  ============" & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & Format(Saldo, "#,##0.00")
        Grilla.AddItem cabecera
        For k = 7 To 9
                                Grilla.col = k
                                Grilla.Row = 1
                                Grilla.CellBackColor = &H80FF&
         Next k
    Exit Sub
End If
 
   
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 3000
           Grilla.ColWidth(5) = 3000
           Grilla.ColWidth(6) = 1200
           
           Grilla.ColWidth(7) = 1200
           Grilla.ColWidth(8) = 1200
           Grilla.ColWidth(9) = 1300
           
           Grilla.ColWidth(10) = 1200
           Grilla.ColWidth(11) = 1200
           Grilla.ColWidth(12) = 1300
       Next
        cabecera = "ID_DETALLE" & vbTab & "EMISION" & vbTab & "MOVIMIENTO" & vbTab & "DNI/RUC" & vbTab & "NOMBRE/RAZON SOCIAL" & vbTab & "CONCEPTO" & vbTab & " T.C" & vbTab & "DEBE" & vbTab & "HABER" & vbTab & "SALDO" & vbTab & "DEBE" & vbTab & "HABER" & vbTab & "SALDO"
        Grilla.AddItem cabecera
         For k = 1 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
        cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "============   SALDO INICIAL  ============" & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & Format(Saldo, "#,##0.00")
        Grilla.AddItem cabecera
         For k = 7 To 9
                                Grilla.col = k
                                Grilla.Row = 1
                                Grilla.CellBackColor = &H80FF&
         Next k
                           
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
        
        If rst("montoreal") > 0 Then
            tdebe = rst("monto")
            thaber = 0
        Else
            tdebe = 0
            thaber = rst("monto")
        End If
            
            If rst("anulado") = "no" Then
                Saldo = Saldo + rst("montoreal")
            End If
            
            
            If rst("id_tipo_movimiento") = "00003" Then
                If rst("id_cuenta_destino") <> 0 Then
                   npersona = rst("cuenta_destino") 'get_nombre_cuenta(rst("id_cuenta_destino"))
                Else
                   npersona = get_nombre_cuenta(rst("id_cuenta"))
                End If
            Else
                npersona = rst("npersona")
            End If
            
            Fila = rst("id") & vbTab & Format(rst("fecha"), "dd-mm-YYYY") & vbTab & rst("tipo_movimiento") & vbTab & rst("id_persona") & vbTab & npersona & vbTab & rst("glosa") & vbTab & Mid(rst("nombre_completo"), 1, 25) & vbTab & Format(tdebe, "#,##0.00") & vbTab & Format(thaber, "#,##0.00") & vbTab & Format(Saldo, "#,##0.00")
            Grilla.AddItem Fila
            
            If rst("anulado") = "si" Then
                
            
                For k = 1 To 9
                                Grilla.col = k
                                Grilla.Row = i + 2
                                Grilla.CellBackColor = &HDFDFE0
                Next k
            End If
            
            For l = 7 To 9
                                Grilla.col = l
                                Grilla.Row = i + 2
                                Grilla.CellBackColor = &H80FF&
            Next l
          
            rst.MoveNext
           
        Next i
    
    
   Me.cmdDetalle.Enabled = False
   
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub

Private Sub llenarGrid_mn(ByVal Grilla As MSHFlexGrid, ByVal id_cuenta As Double, ByVal Saldo As Double)
On Error GoTo salir
Dim tTotal As Double, denominacion As String
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
           Grilla.Rows = 0
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1500
           Grilla.ColWidth(3) = 1500
           Grilla.ColWidth(4) = 3500
           Grilla.ColWidth(5) = 4700
           Grilla.ColWidth(6) = 1200
           
           Grilla.ColWidth(7) = 1400
           Grilla.ColWidth(8) = 1400
           Grilla.ColWidth(9) = 1500
        cabecera = "ID_DETALLE" & vbTab & "EMISION" & vbTab & "MOVIMIENTO" & vbTab & "DNI/RUC" & vbTab & "NOMBRE/RAZON SOCIAL" & vbTab & "CONCEPTO" & vbTab & "T.C" & vbTab & "DEBE" & vbTab & "HABER" & vbTab & "SALDO"
        Grilla.AddItem cabecera
         For k = 1 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
        cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "============   SALDO INICIAL  ============" & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & Format(Saldo, "#,##0.00")
        Grilla.AddItem cabecera
        For k = 7 To 9
                                Grilla.col = k
                                Grilla.Row = 1
                                Grilla.CellBackColor = &H80FF&
         Next k
    Exit Sub
End If
 
   
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 1800
           Grilla.ColWidth(3) = 1400
           Grilla.ColWidth(4) = 3200
           Grilla.ColWidth(5) = 5200
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 1400
           Grilla.ColWidth(8) = 1400
           Grilla.ColWidth(9) = 1400
           
           
     
        cabecera = "ID_DETALLE" & vbTab & "EMISION" & vbTab & "MOVIMIENTO" & vbTab & "DNI/RUC" & vbTab & "NOMBRE/RAZON SOCIAL" & vbTab & "CONCEPTO" & vbTab & " T.C" & vbTab & "DEBE" & vbTab & "HABER" & vbTab & "SALDO"
        Grilla.AddItem cabecera
         For k = 1 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
        cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "============   SALDO INICIAL  ============" & vbTab & " " & vbTab & " " & vbTab & " " & vbTab & Format(Saldo, "#,##0.00")
        Grilla.AddItem cabecera
         For k = 5 To 9
                                Grilla.col = k
                                Grilla.Row = 1
                                Grilla.CellBackColor = &H80FF&
         Next k
                           
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
        
        If rst("montoreal") > 0 Then
            tdebe = rst("monto")
            thaber = 0
        Else
            tdebe = 0
            thaber = rst("monto")
        End If
            
            If rst("anulado") = "no" Then
                Saldo = Saldo + rst("montoreal")
            End If
            
            
            If rst("id_tipo_movimiento") = "00003" Then
                If rst("id_cuenta_destino") <> 0 Then
                   npersona = rst("cuenta_destino") 'get_nombre_cuenta(rst("id_cuenta_destino"))
                Else
                   npersona = get_nombre_cuenta(rst("id_cuenta"))
                End If
            Else
                npersona = rst("npersona")
            End If
            
            Fila = rst("id") & vbTab & Format(rst("fecha"), "dd-mm-YYYY") & vbTab & rst("tipo_movimiento") & vbTab & rst("id_persona") & vbTab & npersona & vbTab & rst("glosa") & vbTab & Format(rst("tc"), "#,##0.0000") & vbTab & Format(tdebe, "#,##0.00") & vbTab & Format(thaber, "#,##0.00") & vbTab & Format(Saldo, "#,##0.00")
            Grilla.AddItem Fila
            
            If rst("anulado") = "si" Then
                
            
                For k = 1 To 9
                                Grilla.col = k
                                Grilla.Row = i + 2
                                Grilla.CellBackColor = &HDFDFE0
                Next k
            End If
            
            'For l = 7 To 9
            '                    Grilla.col = l
             '                   Grilla.Row = i + 2
             '                   Grilla.CellBackColor = &H80FF&
            'Next l
          
            rst.MoveNext
           
        Next i
        
        For k = 5 To 9
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H80FF&
         Next k
    
   Me.cmdDetalle.Enabled = False
   
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub


Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Error
  Select Case Button.key
    Case KEY_NEW
        Procedencia = nuevo
        frmNuevoComprobante.Show
    Case KEY_DELETE
       If MsgBox(MSGANULAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Procedencia = anular
        FrmSeguridad.Show
       End If
    Case KEY_EXIT
        Unload Me
'Error:
 ' MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  
  Exit Sub
End Select
End Sub

Private Sub HfDetalle_SelChange()
If Val(Me.HfDetalle.TextMatrix(Me.HfDetalle.Row, 0)) > 0 Then
   Me.cmdDetalle.Enabled = True
   
Else
    Me.cmdDetalle.Enabled = False
   
End If
End Sub

Private Sub Image1_Click()
Me.frmmonto_nuevo.Visible = False
End Sub

Private Sub TxtFecha_emision_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtMonto)
End If


End Sub

Private Sub txtFormaPago_Change()
strCadena = "SELECT id as Codigo,Descripcion  as Descripcion FROM vw_mediopago_nombre WHERE Descripcion LIKE '%" & Trim(Me.txtFormaPago.Text) & "%'  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFormaPago)
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.txtrucproveedor.Visible = True Then
       Call Resalta(Me.txtrucproveedor)
    End If
    
End If
End Sub

Private Sub txtMontoNuevo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    strCadena = "UPDATE movimiento_compra_pago_temporal SET monto_pago='" & Val(Me.txtMontoNuevo.Text) & "' WHERE id_factura='" & Val(Me.lblid_temporal.Caption) & "' and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Me.frmmonto_nuevo.Visible = False
    Call Me.Llenar_Temporal(Me.HgPagados, Val(Me.lblid_recibo.Caption))
End If
End Sub

Private Sub TxtOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtObservacion)
End If
End Sub

Private Sub txtrucproveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
buscar_nuevamente:
   strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtrucproveedor.Text) & "' LIMIT 1"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      Me.lblProveedor.Caption = rst("nombre_completo")
      Me.DtcTipomovimiento.SetFocus
   Else
        If Len(Trim(Me.txtrucproveedor.Text)) = 8 Then
            If get_dni_reniec_iii(Trim(Me.txtrucproveedor.Text), KEY_DEPARTAMENTO, KEY_PROVINCIA, KEY_DISTRITO) = True Then
                GoTo buscar_nuevamente
                Exit Sub
            End If
        End If
        
      Procedencia = Selecionar
      FrmPersona.Show
      Exit Sub
   End If
End If
End Sub

Private Sub txtTipoFlujo_Change()
strCadena = "SELECT id as Codigo,Nombre  as Descripcion FROM adm_flujocaja WHERE Nombre LIKE '%" & Trim(Me.txtTipoFlujo.Text) & "%' ORDER BY Nombre  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFlujo)
End Sub

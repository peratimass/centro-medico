VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmMemorandun 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   16710
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DETALLE MEMORANDUM"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   15135
      Begin VB.TextBox txtIdmemo 
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
         Left            =   6720
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txtBuscar 
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
         Left            =   6240
         TabIndex        =   33
         Top             =   840
         Width           =   855
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
         Height          =   315
         Left            =   3645
         TabIndex        =   31
         Top             =   1845
         Width           =   1215
      End
      Begin VB.TextBox TxtRazon 
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
         Left            =   9240
         TabIndex        =   24
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox TxtDni 
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
         Left            =   9240
         TabIndex        =   23
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox TxtDetalle 
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
         Height          =   675
         Left            =   1005
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   2280
         Width           =   6135
      End
      Begin MSComCtl2.DTPicker DtpFecha 
         Height          =   375
         Left            =   1005
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   67108865
         CurrentDate     =   43165
      End
      Begin VB.TextBox TxtAsunto 
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
         Left            =   1005
         TabIndex        =   15
         Top             =   1320
         Width           =   5175
      End
      Begin VB.TextBox TxtSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Height          =   330
         Left            =   10920
         MaxLength       =   50
         TabIndex        =   7
         Top             =   735
         Width           =   855
      End
      Begin VB.TextBox TxtNumero 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Height          =   330
         Left            =   12120
         MaxLength       =   50
         TabIndex        =   6
         Top             =   735
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo DtcComrpobante 
         Height          =   345
         Left            =   7920
         TabIndex        =   8
         Top             =   735
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   345
         Left            =   7920
         TabIndex        =   9
         Top             =   360
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo DtcEmisor 
         Height          =   330
         Left            =   1005
         TabIndex        =   10
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
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
      Begin MSDataListLib.DataCombo DtcReceptor 
         Height          =   330
         Left            =   1005
         TabIndex        =   12
         Top             =   840
         Width           =   5175
         _ExtentX        =   9128
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfComprobantes 
         Height          =   3375
         Left            =   1005
         TabIndex        =   20
         Top             =   3240
         Width           =   13120
         _ExtentX        =   23151
         _ExtentY        =   5953
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
      Begin VitekeySoft.ChameleonBtn cmdImprimir 
         Height          =   900
         Left            =   12315
         TabIndex        =   26
         Top             =   6735
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1588
         BTYPE           =   5
         TX              =   "IMPRIMIR"
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMemorandun.frx":0000
         PICN            =   "frmMemorandun.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdProcesar 
         Height          =   900
         Left            =   11400
         TabIndex        =   27
         Top             =   6750
         Width           =   855
         _ExtentX        =   1508
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
         MICON           =   "frmMemorandun.frx":25ED
         PICN            =   "frmMemorandun.frx":2609
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdCerrarpantalla 
         Height          =   900
         Left            =   13320
         TabIndex        =   28
         Top             =   6735
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1588
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
         MICON           =   "frmMemorandun.frx":5C51
         PICN            =   "frmMemorandun.frx":5C6D
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
         Height          =   900
         Left            =   10440
         TabIndex        =   29
         Top             =   6750
         Width           =   855
         _ExtentX        =   1508
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
         MICON           =   "frmMemorandun.frx":8C94
         PICN            =   "frmMemorandun.frx":8CB0
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
         Height          =   420
         Left            =   14160
         TabIndex        =   30
         Top             =   3240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   741
         BTYPE           =   5
         TX              =   ""
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMemorandun.frx":9102
         PICN            =   "frmMemorandun.frx":911E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcMoneda 
         Height          =   330
         Left            =   4920
         TabIndex        =   35
         Top             =   1845
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label9 
         Alignment       =   2  'Center
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
         Left            =   2955
         TabIndex        =   32
         Top             =   1845
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LISTADO DE COMPROBANTES :"
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
         Left            =   990
         TabIndex        =   25
         Top             =   3000
         Width           =   2025
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE/ RAZON :"
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
         Left            =   7890
         TabIndex        =   22
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI / RUC :"
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
         Left            =   7890
         TabIndex        =   21
         Top             =   1440
         Width           =   765
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   795
         Left            =   7560
         Shape           =   4  'Rounded Rectangle
         Top             =   1320
         Width           =   6375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETALLE:"
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
         Left            =   165
         TabIndex        =   18
         Top             =   2400
         Width           =   645
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA :"
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
         Left            =   285
         TabIndex        =   16
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ASUNTO :"
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
         Left            =   165
         TabIndex        =   14
         Top             =   1335
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARA :"
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
         Left            =   345
         TabIndex        =   13
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DE:"
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
         TabIndex        =   11
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape ShpDatos 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00E0E0E0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   1035
         Left            =   7560
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   6375
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdBuscar 
      Height          =   375
      Left            =   13800
      TabIndex        =   41
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMemorandun.frx":96B8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtBuscarDNI 
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
      Left            =   960
      TabIndex        =   38
      Top             =   520
      Width           =   1935
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   975
      Left            =   15360
      TabIndex        =   0
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
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
      MICON           =   "frmMemorandun.frx":96D4
      PICN            =   "frmMemorandun.frx":96F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfMemorandum 
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   12303
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
      Height          =   975
      Left            =   15360
      TabIndex        =   2
      Top             =   2960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
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
      MICON           =   "frmMemorandun.frx":BE23
      PICN            =   "frmMemorandun.frx":BE3F
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
      Height          =   975
      Left            =   15360
      TabIndex        =   3
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
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
      MICON           =   "frmMemorandun.frx":E289
      PICN            =   "frmMemorandun.frx":E2A5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdVisualizar 
      Height          =   975
      Left            =   15360
      TabIndex        =   36
      Top             =   1960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "VISUALIZAR"
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
      MICON           =   "frmMemorandun.frx":E695
      PICN            =   "frmMemorandun.frx":E6B1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   375
      Left            =   9960
      TabIndex        =   39
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   67108865
      CurrentDate     =   43165
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   375
      Left            =   12000
      TabIndex        =   40
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   67108865
      CurrentDate     =   43165
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
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
      Left            =   11670
      TabIndex        =   42
      Top             =   600
      Width           =   195
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI/ RUC :"
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
      TabIndex        =   37
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LISTADO DE MEMORANDUM'S"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2445
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8160
      Left            =   0
      Top             =   0
      Width           =   16710
   End
End
Attribute VB_Name = "frmMemorandun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub cmdBuscar_Click()
strCadena = "SELECT * FROM view_memorandum_externo where fecha>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and   ruc='" & KEY_RUC & "' ORDER BY id_memo DESC LIMIT 20"
Call Me.llenar_memorandum(Me.HfMemorandum)
End Sub

Private Sub cmdCerrarpantalla_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub cmddelete_Click()
strCadena = "call sp_memorandum_delete_temporal('" & Val(Me.HfComprobantes.TextMatrix(Me.HfComprobantes.Row, 0)) & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
Call llenar_estado_cuenta(Trim(Me.txtDni.Text))
End Sub

Private Sub cmdEliminar_Click()
Procedencia = anular
Call disabled_form(Me)
frmsegurity.Show

End Sub

Private Sub cmdImprimir_Click()
Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant

arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"



arr(0, 2) = "00001"
arr(1, 2) = UCase(EnLetras(Val(Me.txtMonto.Text))) & Space(1) & Trim(Me.DtcMoneda.Text)


param = arr()


    


strCadena = "SELECT * FROM view_memorandum_externo WHERE id_memo='" & Val(Me.txtIdmemo.Text) & "'"
Call ConfiguraRst(strCadena)
strCadena = "SELECT * FROM view_memorandum_detalle WHERE id_memo='" & Val(Me.txtIdmemo.Text) & "'"
Call ConfiguraRstK(strCadena)
Ans = ShowMultiReport(rst, "Rptmemorandum_externo", param, App.Path + "\Reportes\", , , True, , rstK, "Rptmemorandum_detalle")



End Sub

Private Sub cmdNuevo_Click()

strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE id_personal='si' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcReceptor)


strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE dni='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEmisor)




Me.frmdetalle.Visible = True

End Sub

Private Sub cmdProcesar_Click()


strCadena = "call sp_memorandum_save('" & Me.DtcEmisor.BoundText & "','" & Me.DtcReceptor.BoundText & "','" & Trim(Me.TxtAsunto.Text) & "','" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "','" & UCase(Trim(Me.TxtDetalle.Text)) & "','" & Val(Me.txtMonto.Text) & "','" & Trim(Me.txtDni.Text) & "','" & Me.DtcComrpobante.BoundText & "','" & Trim(Me.txtserie.Text) & "','" & Trim(Me.txtNumero.Text) & "','" & Me.DtcMoneda.BoundText & "','" & KEY_RUC & "')"
Call ConfiguraRstP(strCadena)
Me.txtIdmemo.Text = rstP(0)
Call save_detalle_memo(Val(Me.txtIdmemo.Text))

Me.cmdImprimir.Enabled = True
Me.cmdProcesar.Enabled = False

End Sub
Private Sub save_detalle_memo(ByVal in_memo As String)

strCadena = "select * from view_memorandum_temporal WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        strCadena = "call put_momorandum_detalle('" & Val(in_memo) & "','" & rst("id_venta") & "','" & rst("total") & "','" & rst("saldo") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        rst.MoveNext
    Next
End If



End Sub



Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdVisualizar_Click()
Call load_memorandum(Val(Me.HfMemorandum.TextMatrix(Me.HfMemorandum.Row, 0)))
End Sub

Private Sub DtcComrpobante_Change()
Call get_comprobante(Me.DtcComrpobante.BoundText)
End Sub
Private Sub get_comprobante(ByVal in_doc As String)
strCadena = "SELECT * FROM memorandun WHERE id_doc='" & in_doc & "' and   ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
     Me.txtserie.Text = rst("serie")
     Me.txtNumero.Text = Format(Val(rst("numero")) + 1, "000000")
Else
    Me.txtserie.Text = "001"
    Me.txtNumero.Text = "000001"
End If
End Sub

Private Sub Form_Load()
CenterForm Me

Me.Top = 50
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA

Me.Dtpfecha.Value = KEY_FECHA
strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' and id_tipoentidad='0' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM

strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM view_almacen_comprobante_ultimate WHERE id_doc in ('0415','0416') and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComrpobante)


strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion  FROM moneda"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)
Me.DtcMoneda.BoundText = "00001"


Call actualizar
    
    
    






End Sub
Public Sub actualizar()

strCadena = "SELECT * FROM view_memorandum_externo where ruc='" & KEY_RUC & "' ORDER BY id_memo DESC LIMIT 20"
Call Me.llenar_memorandum(Me.HfMemorandum)

End Sub
Public Sub llenar_grid(ByVal Grilla As MSHFlexGrid)
Dim nsaldo As Double
Dim in_operador As String
On Error GoTo salir
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
           Grilla.ColWidth(3) = 2500
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 3500
           Grilla.ColWidth(6) = 1500
           Grilla.ColWidth(7) = 1500
        Next
         
         cabecera = "IDVENTA" & vbTab & "EMISION" & vbTab & "VENCIMIENTO" & vbTab & "COMPROBANTE" & vbTab & "DNI CLIENTE" & vbTab & "DATOS CLIENTE" & vbTab & "TOTAL" & vbTab & "SALDO" & vbTab & "REFERENCIA" & vbTab & "%SEGURO"
         Grilla.AddItem cabecera
         
         For k = 1 To 7
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
         
        rst.MoveFirst
        nfactor = 0
        in_acumulado = 0
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & Format(rst("fecha_vencimiento"), "dd-mm-YYYY") & vbTab & rst("comprobante") & vbTab & rst("id_cliente") & vbTab & rst("ncliente") & vbTab & rst("simbolo") & Space(1) & Format(rst("total"), "#,##0.00") & vbTab & rst("simbolo") & Space(2) & Format(rst("saldo"), "#,##0.00") & vbTab & rst("referencia")
             Grilla.AddItem Fila
               If rst("id_moneda") = "00002" Then
                    in_saldo = rst("saldo") * rst("tc")
                  
                Else
                    in_saldo = rst("saldo")
                    
                End If
                nsaldo = nsaldo + in_saldo
                
            If rst("saldo") > 0 Then
            For k = 6 To 7
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H80C0FF
            Next k
            End If
            
            If rst("anulado") = "si" Then
           
            For k = 1 To 7
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
            Next k
            End If
            rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "MONTO A COBRAR:" & vbTab & "" & vbTab & "S/.  " & Format(nsaldo, "#,##0.00")
        Grilla.AddItem Fila
   Grilla.ColAlignment(3) = 1
  Grilla.ColAlignment(5) = 1
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Public Sub llenar_memorandum(ByVal Grilla As MSHFlexGrid)
Dim nsaldo As Double
Dim in_operador As String
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2500
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 3000
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 2800
           Grilla.ColWidth(7) = 2800
        Next
         
         cabecera = "ID_MEMO" & vbTab & "MEMORANDUM" & vbTab & "FECHA" & vbTab & "DNI/RUC" & vbTab & "CLIENTE" & vbTab & "MONTO" & vbTab & "DE" & vbTab & "PARA"
         Grilla.AddItem cabecera
         
         For k = 1 To 7
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
         
        rst.MoveFirst
        in_monto = 0
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id_memo") & vbTab & rst("numero") & vbTab & Format(rst("fecha"), "dd-mm-YYYY") & vbTab & rst("in_dni_cliente") & vbTab & rst("cliente") & vbTab & Format(rst("monto"), "#,##0.00") & vbTab & rst("emisor") & vbTab & rst("receptor")
             Grilla.AddItem Fila
               If rst("id_moneda") = "00002" Then
                    in_monto = rst("monto") * KEY_CAMBIO
                    
                Else
                    in_monto = rst("monto")
                    
                End If
                in_total = in_total + in_monto
            If rst("anulado") = "si" Then
            For k = 1 To 7
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
            Next k
            End If
            
            
            rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "MONTO A COBRAR:" & vbTab & "" & vbTab & "S/.  " & Format(nsaldo, "#,##0.00")
        Grilla.AddItem Fila
  Grilla.ColAlignment(3) = 1
  Grilla.ColAlignment(5) = 1
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub
Public Sub load_memorandum(ByVal in_memo As String)

strCadena = "SELECT * FROM memorandun WHERE id_memo='" & Val(in_memo) & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
        Me.txtIdmemo.Text = rstK("id_memo")
        strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE dni='" & rstK("id_receptor") & "' and  id_personal='si' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcReceptor)
        
        strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE dni='" & rstK("id_emisor") & "' and  dni='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcEmisor)
        
        Me.DtcComrpobante.BoundText = rstK("id_doc")
        Me.txtserie.Text = rstK("serie")
        Me.txtNumero.Text = rstK("numero")
        Me.txtDni.Text = rstK("in_dni_cliente")
        Me.TxtRazon.Text = get_persona(rstK("in_dni_cliente"))
        Me.Dtpfecha.Value = rstK("fecha")
        Me.TxtAsunto.Text = rstK("asunto")
        Me.TxtDetalle.Text = rstK("detalle")
        Me.txtMonto.Text = rstK("monto")
        Me.DtcMoneda.BoundText = rstK("id_moneda")
        strCadena = "SELECT * FROM view_memorandum_detalle_listar WHERE id_memo='" & Val(Me.txtIdmemo.Text) & "'"
        Call llenar_grid(Me.HfComprobantes)
        Me.cmdProcesar.Enabled = False
        Me.cmdImprimir.Enabled = True
        Me.frmdetalle.Visible = True

End If

End Sub


Private Sub HfComprobantes_SelChange()
If Val(Me.HfComprobantes.Rows) > 0 Then
   Me.cmdDelete.Enabled = True
Else
   Me.cmdDelete.Enabled = False
End If
End Sub

Private Sub txtBuscar_Change()
strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE nombre_completo LIKE '%" & Trim(Me.txtBuscar.Text) & "%' and  id_personal='si' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcReceptor)
End Sub

Public Sub put_estado_cuenta_temp(ByVal in_dni As String)
strCadena = "DELETE FROM memorandum_detalle_temp WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT id_venta FROM view_listado_comprobante_vargas_real WHERE total>pago and id_cliente='" & cPersona & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision"

strCadena = "SELECT id_venta FROM view_listado_comprobante_ultimate WHERE saldo>0 and   id_cliente = '" & in_dni & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    For i = 0 To rst.RecordCount - 1
    strCadena = "call put_momorandum_temp('" & rst("id_venta") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    rst.MoveNext
    Next i
End If
Call llenar_estado_cuenta(Trim(Me.txtDni.Text))
End Sub
Public Sub llenar_estado_cuenta(ByVal in_dni As String)


strCadena = "SELECT * FROM view_memorandum_temporal WHERE id_cliente = '" & in_dni & "' and dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
Call llenar_grid(Me.HfComprobantes)
End Sub

Private Sub TxtBuscarDNI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_memorandum_externo where in_dni_cliente='" & Trim(Me.TxtBuscarDNI.Text) & "'  and   ruc='" & KEY_RUC & "' ORDER BY id_memo DESC LIMIT 20"
    Call Me.llenar_memorandum(Me.HfMemorandum)
End If
End Sub
Public Sub LlenarVinculados(ByVal Grilla As MSHFlexGrid, ByVal cPersona As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,comprobante,descripcion,total,pago,seleccion FROM view_listado_comprobante_vargas_real WHERE total>pago and id_cliente='" & cPersona & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2100
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 400
         Next
        cabecera = "CODIGO" & vbTab & "EMISION  - VENCE" & vbTab & "COMPROBANTE" & vbTab & "MONEDA" & vbTab & "MONTO TOTAL" & vbTab & "SALDO" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 1 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        in_total_pago = 0
        For i = 0 To rst.RecordCount - 1
            If rst("seleccion") = "si" Then
                estado = Chr(254)
            Else
                estado = Chr(168)
            End If
            in_saldo = rst("total") - rst("pago")
            Fila = rst("id_venta") & vbTab & "[" & Format(rst("fecha_emision"), "dd-mm-YYYY") & " - " & Format(rst("fecha_vencimiento"), "dd-mm-YYYY") & "]" & vbTab & rst("comprobante") & vbTab & rst("descripcion") & vbTab & Format(rst("total"), "#,##0.00") & vbTab & Format(in_saldo, "#,##0.00") & vbTab & estado
            Grilla.AddItem Fila
            in_total_pago = in_total_pago + in_saldo
            With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 6 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
            End With
      
            
            
            If rst("seleccion") = "si" Then
                tTotal = tTotal + rst("total") - rst("pago")
                For k = 1 To 5
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
                Next k
            End If
            
            Fila = ""
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "SALDO TOTAL :::>" & vbTab & Format(in_total_pago, "#,##0.00")
      Grilla.AddItem Fila
      
            Grilla.col = 5
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      
            
  Exit Sub
salir:    MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Private Sub txtDni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtDni.Text) & "' "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.TxtRazon.Text = rst("nombre_completo")
       Call put_estado_cuenta_temp(Trim(Me.txtDni.Text))
       
    Else
        Procedencia = Selecionar
        FrmPersona.Show
        Exit Sub
    
    End If
    
End If
End Sub

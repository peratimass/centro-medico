VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmSolicitudViaticos 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18690
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   18690
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmdetalle 
      BackColor       =   &H00FFFFFF&
      Height          =   7360
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   18255
      Begin VB.TextBox txtNumero 
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
         TabIndex        =   40
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtSerie 
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
         Left            =   3720
         TabIndex        =   39
         Top             =   360
         Width           =   855
      End
      Begin VitekeySoft.ChameleonBtn CmdQuitar 
         Height          =   855
         Left            =   12720
         TabIndex        =   30
         Top             =   3120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "QUITAR"
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
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmSolicitudViaticos.frx":0000
         PICN            =   "FrmSolicitudViaticos.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox TxtObservacion 
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
         Height          =   555
         Left            =   2160
         MaxLength       =   80
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   6600
         Width           =   10575
      End
      Begin VB.TextBox TxtMontoMotivo 
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
         Left            =   11160
         MaxLength       =   80
         TabIndex        =   27
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox TxtDetalle 
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
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   26
         Top             =   2640
         Width           =   8775
      End
      Begin VB.TextBox TxtResumen 
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
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   25
         Top             =   2160
         Width           =   10455
      End
      Begin VB.TextBox TxtMonto 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
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
         Height          =   315
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   24
         Top             =   6120
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DtpFechaSolicitud 
         Height          =   315
         Left            =   2160
         TabIndex        =   21
         Top             =   360
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
         CalendarBackColor=   16777215
         CalendarForeColor=   -2147483635
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   -2147483635
         Format          =   186777601
         CurrentDate     =   37091
      End
      Begin MSComCtl2.DTPicker DtpInicio 
         Height          =   315
         Left            =   2160
         TabIndex        =   22
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
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
         CalendarBackColor=   16777215
         CalendarForeColor=   -2147483635
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   -2147483635
         Format          =   186777601
         CurrentDate     =   37091
      End
      Begin MSComCtl2.DTPicker DtpFin 
         Height          =   315
         Left            =   3720
         TabIndex        =   23
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
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
         CalendarBackColor=   16777215
         CalendarForeColor=   -2147483635
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   -2147483635
         Format          =   186777601
         CurrentDate     =   37091
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDetalle 
         Height          =   2895
         Left            =   2160
         TabIndex        =   29
         Top             =   3105
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5106
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
      Begin VitekeySoft.ChameleonBtn cmdcerrarviaticos 
         Height          =   375
         Left            =   17760
         TabIndex        =   31
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmSolicitudViaticos.frx":2466
         PICN            =   "FrmSolicitudViaticos.frx":2482
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar_detalle 
         Height          =   825
         Left            =   14760
         TabIndex        =   32
         Top             =   6180
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1455
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
         MICON           =   "FrmSolicitudViaticos.frx":5336
         PICN            =   "FrmSolicitudViaticos.frx":5352
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdSalir_detalle 
         Height          =   825
         Left            =   15840
         TabIndex        =   33
         Top             =   6180
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1455
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
         MICON           =   "FrmSolicitudViaticos.frx":899A
         PICN            =   "FrmSolicitudViaticos.frx":89B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcSolicitante 
         Height          =   315
         Left            =   2160
         TabIndex        =   35
         Top             =   960
         Width           =   6135
         _ExtentX        =   10821
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NUMERO"
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
         Left            =   4680
         TabIndex        =   38
         Top             =   165
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SERIE"
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
         Left            =   3720
         TabIndex        =   37
         Top             =   165
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOLICITANTE  :"
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
         Left            =   840
         TabIndex        =   36
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label lblidsolicitud 
         Height          =   375
         Left            =   10440
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Image Image1 
         Height          =   4500
         Left            =   13680
         Picture         =   "FrmSolicitudViaticos.frx":8DA6
         Stretch         =   -1  'True
         Top             =   240
         Width           =   4170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA SOLICITUD :"
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
         Left            =   510
         TabIndex        =   20
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIODO DE GASTOS :"
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
         Left            =   240
         TabIndex        =   19
         Top             =   1560
         Width           =   1725
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DETALLE :"
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
         Left            =   1230
         TabIndex        =   18
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACION :"
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
         Left            =   765
         TabIndex        =   17
         Top             =   6720
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RESUMEN :"
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
         Left            =   1110
         TabIndex        =   16
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL A SOLICITAR :"
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
         Left            =   405
         TabIndex        =   15
         Top             =   6120
         Width           =   1560
      End
   End
   Begin MSDataListLib.DataCombo DtcTrabajador 
      Height          =   315
      Left            =   8640
      TabIndex        =   13
      Top             =   405
      Width           =   5175
      _ExtentX        =   9128
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
   Begin VB.TextBox txtDescripcion 
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
      TabIndex        =   11
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtsolicitud 
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
      Left            =   1320
      TabIndex        =   9
      Top             =   440
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   7335
      Left            =   240
      TabIndex        =   0
      Top             =   870
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   12938
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
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   825
      Left            =   17520
      TabIndex        =   2
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
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
      MICON           =   "FrmSolicitudViaticos.frx":194A5
      PICN            =   "FrmSolicitudViaticos.frx":194C1
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
      Height          =   825
      Left            =   17520
      TabIndex        =   3
      Top             =   1740
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
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
      MICON           =   "FrmSolicitudViaticos.frx":19913
      PICN            =   "FrmSolicitudViaticos.frx":1992F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdDeclarar 
      Height          =   825
      Left            =   17520
      TabIndex        =   4
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
      BTYPE           =   5
      TX              =   "DECLARAR"
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
      MICON           =   "FrmSolicitudViaticos.frx":19C49
      PICN            =   "FrmSolicitudViaticos.frx":19C65
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdAtender 
      Height          =   825
      Left            =   17520
      TabIndex        =   5
      Top             =   3540
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
      BTYPE           =   5
      TX              =   "ATENDER"
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
      MICON           =   "FrmSolicitudViaticos.frx":19F7F
      PICN            =   "FrmSolicitudViaticos.frx":19F9B
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
      Height          =   825
      Left            =   17520
      TabIndex        =   6
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
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
      MICON           =   "FrmSolicitudViaticos.frx":1D271
      PICN            =   "FrmSolicitudViaticos.frx":1D28D
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
      Height          =   825
      Left            =   17520
      TabIndex        =   7
      Top             =   5340
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
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
      MICON           =   "FrmSolicitudViaticos.frx":1D5A7
      PICN            =   "FrmSolicitudViaticos.frx":1D5C3
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
      BackStyle       =   0  'Transparent
      Caption         =   "TRABAJADOR:"
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
      Left            =   7560
      TabIndex        =   12
      Top             =   480
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3600
      TabIndex        =   10
      Top             =   480
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N° SOLICITUD :"
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
      TabIndex        =   8
      Top             =   480
      Width           =   990
   End
   Begin VB.Label LblEmpresa 
      BackStyle       =   0  'Transparent
      Caption         =   "SOLICITUD DE VIATICOS"
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
      TabIndex        =   1
      Top             =   100
      Width           =   2265
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8370
      Left            =   0
      Top             =   0
      Width           =   18690
   End
End
Attribute VB_Name = "FrmSolicitudViaticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede







Private Sub ChameleonBtn1_Click()

End Sub

Private Sub cmdAnular_Click()
Procedencia = anular
frmsegurity.Show
Call disabled_form(Me)
End Sub

Private Sub cmdAtender_Click()
FrmSolicitudViaticosAtender.Show
End Sub

Private Sub cmdcerrarviaticos_Click()

Me.frmdetalle.Visible = False

End Sub
Private Function get_numero_solicitud(ByVal in_alm As String)

End Function

Private Sub Save()
Dim id_solicitud As Double, numero As String
  If Val(Me.TxtMonto.Text) <= 0 Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    Select Case FrmSolicitudViaticos.Procedencia
      Case nuevo
       
       strCadena = "SELECT * FROM solicitud_dinero WHERE dni='" & Me.DtcSolicitante.BoundText & "' AND ruc='" & KEY_RUC & "'  AND finalizado='no' and anulado='no'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 2 Then
            MsgBox "USUARIO CUENTA CON " + str(rst.RecordCount) + Space(1) + "SOLICITUDES PENDIENTE", vbInformation, "Mensaje para el Usuario"
            Exit Sub
       End If
        
       
      strCadena = "SELECT * FROM solicitud_dinero WHERE id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
      Call ConfiguraRst(strCadena)
      If rst.RecordCount > 0 Then
         Me.txtSerie.Text = rst("serie")
         Me.TxtNumero.Text = Format(Val(rst("numero")) + 1, "000000")
      Else
         strCadena = "SELECT serie,numero FROM almacen_comprobante WHERE id_doc='0423' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
         Call ConfiguraRst(strCadena)
         If rst.RecordCount > 0 Then
            Me.txtSerie.Text = rst("serie")
            Me.TxtNumero.Text = rst("numero")
      
        End If
      End If
      
        
        strCadena = "INSERT INTO solicitud_dinero (serie,numero,resumen,monto_solicitado,saldo,fecha_solicitud,hora_solicitud,fecha_inicio,fecha_fin,observacion,dni,dni_save,id_alm,ruc) VALUES " & _
        " ('" & Trim(Me.txtSerie.Text) & "','" & Trim(Me.TxtNumero.Text) & "','" & UCase(Trim(Me.TxtResumen.Text)) & "','" & Val(Me.TxtMonto.Text) & "','" & Val(Me.TxtMonto.Text) & "','" & Format(Me.DtpFechaSolicitud.Value, "YYYY-mm-dd") & "','" & str(Time) & "','" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "'," & _
        "'" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "','" & Trim(Me.TxtObservacion.Text) & "','" & Me.DtcSolicitante.BoundText & "','" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        id_solicitud = LastRegistro("solicitud_dinero", "id_solicitud")
        
        strCadena = "SELECT * FROM solicitud_dinero_temporal WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
           rst.MoveFirst
           For i = 0 To rst.RecordCount - 1
                strCadena = "INSERT INTO solicitud_dinero_detalle(id_solicitud,descripcion,monto,ruc)VALUES('" & id_solicitud & "','" & rst("detalle") & "','" & rst("monto") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                rst.MoveNext
           Next i
           
           strCadena = "DELETE FROM solicitud_dinero_temporal WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
           
           strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(Me.TxtNumero.Text) + 1, "000000") & "' WHERE id_doc='0423' and id_alm='" & KEY_ALM & "'  and  ruc='" & KEY_RUC & "'"
           CnBd.Execute (strCadena)
        End If
        
        strCadena = "call P_insert_venta_agenda_xd_viaticos('" & Val(id_solicitud) & "')"
        CnBd.Execute (strCadena)
        
        Me.frmdetalle.Visible = False
        Call FrmSolicitudViaticos.actualizar
        Me.cmdprocesar_detalle.Enabled = False
        
       
       End Select
   End If
End Sub

Private Sub cmdDeclarar_Click()
        FrmSolicitudViaticosDeclarar.Show
        
End Sub

Private Sub cmdNuevo_Click()
      Procedencia = nuevo
      strCadena = "call sp_solicitud_viaticos_new('" & KEY_USUARIO & "','" & KEY_RUC & "')"
      CnBd.Execute (strCadena)
      
      strCadena = "SELECT * FROM solicitud_dinero WHERE id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
      Call ConfiguraRst(strCadena)
      If rst.RecordCount > 0 Then
         Me.txtSerie.Text = rst("serie")
         Me.TxtNumero.Text = Format(Val(rst("numero")) + 1, "000000")
      Else
         strCadena = "SELECT serie,numero FROM almacen_comprobante WHERE id_doc='0423' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
         Call ConfiguraRst(strCadena)
         If rst.RecordCount > 0 Then
            Me.txtSerie.Text = rst("serie")
            Me.TxtNumero.Text = rst("numero")
        Else
            MsgBox "NO CUENTA CON ESTE COMPROBANTE", vbInformation
            Exit Sub
        End If
      End If
      
      Me.DtpFechaSolicitud.Value = KEY_FECHA
      Me.DtpFin.Value = KEY_FECHA
      Me.DtpInicio.Value = KEY_FECHA
      Me.TxtResumen.Text = ""
      Me.TxtObservacion.Text = ""
      Me.lblidsolicitud.Caption = 0
      Me.HfDetalle.Rows = 0
      Me.TxtMonto.Text = 0
      frmdetalle.Visible = True
      Call Resalta(Me.TxtResumen)
      
End Sub

Private Sub cmdprocesar_detalle_Click()
Call Save
End Sub

Private Sub CmdQuitar_Click()
strCadena = "SELECT * FROM solicitud_dinero_temporal WHERE id='" & Val(HfDetalle.TextMatrix(HfDetalle.Row, 0)) & "'"
CnBd.Execute (strCadena)
Call llenarGrid_temp(Me.HfDetalle)

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdsalir_detalle_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub cmdVisualizar_Click()
        Call load_viatico(Val(HfgDetalle.TextMatrix(HfgDetalle.Row, 0)))
        Me.cmdprocesar_detalle.Enabled = False
End Sub

Public Sub actualizar()
strCadena = "SELECT * FROM view_solicitud_viatico WHERE  ruc='" & KEY_RUC & "'"
Call llenarGrid(Me.HfgDetalle, Me)
End Sub

Public Sub load_viatico(ByVal in_viatico As String)
strCadena = "SELECT * FROM view_solicitud_viatico WHERE id_solicitud='" & Val(in_viatico) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.lblidsolicitud.Caption = in_viatico
    Me.DtpFechaSolicitud.Value = rst("fecha_solicitud")
    Me.DtpInicio.Value = rst("fecha_inicio")
    Me.DtpFin.Value = rst("fecha_fin")
    
    
    Me.TxtResumen.Text = rst("resumen")
    
    If rst("dni") = KEY_USUARIO And rst("anulado") = "no" And rst("atendido") = "no" Then
        Me.cmdprocesar_detalle.Enabled = False
        
    End If
    Call llenarGrid_detalle(Me.HfDetalle, in_viatico)
    Me.frmdetalle.Visible = True
Else
   Me.lblidsolicitud.Caption = 0
End If
End Sub

Private Sub HfgMarcas_Click()
If HfgMarcas.Row > 0 Then
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    TlbAcciones.Buttons(KEY_DELETE).Enabled = True
  End If
End Sub
Private Sub LLENA(ByVal id_solicitud As Double)
 strCadena = "SELECT * FROM solicitud_dinero WHERE id_solicitud='" & id_solicitud & "' AND ruc='" & KEY_RUC & "'"
 Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    
 End If
End Sub

Private Sub DtcTrabajador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   strCadena = "SELECT * FROM view_solicitud_viatico WHERE dni ='" & Me.DtcTrabajador.BoundText & "' and  ruc='" & KEY_RUC & "'"
   Call llenarGrid(Me.HfgDetalle, Me)
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100

strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE  id_personal='si' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTrabajador)

strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE  id_personal='si' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcSolicitante)
Me.DtcSolicitante.BoundText = KEY_USUARIO
Call actualizar





End Sub

Private Sub HfgDetalle_SelChange()
If Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) > 0 Then
    
    
    If Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 10) = "ANULADO" Then
       Me.cmdAtender.Enabled = False
       Me.cmdAnular.Enabled = False
       Me.cmdDeclarar.Enabled = False
       Me.cmdVisualizar.Enabled = True
       Exit Sub
    End If
    
   If Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 8) = "no" Then
        Me.cmdAtender.Enabled = True
        Me.cmdAnular.Enabled = True
        Me.cmdDeclarar.Enabled = False
        Me.cmdVisualizar.Enabled = True
        Exit Sub
   End If
   If Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 8) = "si" Then
        Me.cmdAtender.Enabled = False
        Me.cmdAnular.Enabled = False
        Me.cmdDeclarar.Enabled = True
        Me.cmdVisualizar.Enabled = True
        
    End If
    
    
    If Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 9) = "si" Then
        Me.cmdAtender.Enabled = False
        Me.cmdAnular.Enabled = False
        Me.cmdDeclarar.Enabled = True
        Me.cmdVisualizar.Enabled = True
        
   
        
    End If
    
    
   
Else
   Me.cmdAnular.Enabled = False
   Me.cmdAtender.Enabled = False
   Me.cmdDeclarar.Enabled = False
   Me.cmdVisualizar.Enabled = False
End If
End Sub



Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case KEY_NEW
      
      
      
     Case KEY_ATENDER
            
     Case KEY_UPDATE
        
        
    Case "(Declarar)"
        
   
    Case KEY_DELETE
        If MsgBox("ESTA SEGURO DE ANULAR ESTA SOLICITUD", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
            strCadena = "UPDATE solicitud_dinero SET anulado='si' WHERE id_solicitud='" & Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            Call actualizar
            
        End If
    Case KEY_EXIT
        Unload Me
  End Select
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
'On Error GoTo salir
Dim tTotal As Double, ccostos As String
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
           Grilla.ColWidth(2) = 1100
           Grilla.ColWidth(3) = 3000
           Grilla.ColWidth(4) = 3000
           Grilla.ColWidth(5) = 1300
           Grilla.ColWidth(6) = 1300
           Grilla.ColWidth(7) = 1300
           Grilla.ColWidth(8) = 2500
           Grilla.ColWidth(9) = 0
           Grilla.ColWidth(10) = 0
           Grilla.ColWidth(11) = 1400
        Next
        cabecera = "IDSOLICITUD" & vbTab & "NUMERO" & vbTab & "FECHA" & vbTab & "SOLICITANTE" & vbTab & "MOTIVO" & vbTab & " M.SOLICITADO" & vbTab & " SALDO" & vbTab & "DECLARADO" & vbTab & "OPERADOR" & vbTab & "atendido" & vbTab & "finalizado" & vbTab & "ESTADO"
        Grilla.AddItem cabecera
         For k = 0 To 11
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
          
          
          
          If rst("atendido") = "si" Then
             in_estado = "ATENDIDO"
          Else
             in_estado = "PENDIENTE"
          End If
          
          If rst("finalizado") = "si" Then
             in_estado = "FINALIZADO"
          End If
          
          If rst("anulado") = "si" Then
             in_estado = "ANULADO"
          End If
          
          Fila = rst("id_solicitud") & vbTab & rst("solicitud") & vbTab & rst("fecha_solicitud") & vbTab & rst("empleado") & vbTab & rst("resumen") & vbTab & Format(rst("monto_solicitado"), "#,##0.00") & vbTab & Format(rst("monto_solicitado") - rst("declarado"), "#,##0.00") & vbTab & Format(rst("declarado"), "#,##0.00") & vbTab & Mid(rst("operador"), 1, 25) & vbTab & rst("atendido") & vbTab & rst("finalizado") & vbTab & in_estado
          Grilla.AddItem Fila
          If rst("anulado") = "si" Then
                        For k = 1 To 11
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
                        Next k
                        GoTo siguiente
          End If
          
          If rst("finalizado") = "si" Then
                        For k = 5 To 11
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H80FF&
                        Next k
                        GoTo siguiente
          End If
          If rst("atendido") = "no" Then
                        For k = 5 To 11
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                        Next k
         Else
                        
                        For k = 5 To 11
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H80C0FF
                        Next k
                        
          
                        
          End If
          
siguiente:
        rst.MoveNext
             
        Next i
    
' Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub llenarGrid_temp(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim tTotal As Double
tTotal = 0
strCadena = "SELECT * FROM solicitud_dinero_temporal WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.TxtMonto.Text = 0
    Grilla.Rows = 0
    Me.CmdQuitar.Visible = False
    Grilla.Rows = 0
    Exit Sub

End If
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 7500
           Grilla.ColWidth(3) = 1500
       Next
         cabecera = "IDITEM" & vbTab & "ITEM" & vbTab & "DESCRIPCION" & vbTab & "MONTO"
         Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        
        For i = 1 To rst.RecordCount
        tTotal = tTotal + rst("monto")
             Fila = rst("id") & vbTab & formato_item(i, 2) & vbTab & rst("detalle") & vbTab & Format(rst("monto"), "#,##0.00")
             Grilla.AddItem Fila
             Fila = ""
        rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "***********  TOTAL A SOLICITAR  ***********" & vbTab & Format(tTotal, "###0.00")
       Grilla.AddItem Fila
       Me.TxtMonto.Text = Format(tTotal, "###0.00")
       For k = 0 To 3
            Grilla.col = k
            Grilla.Row = i
            Grilla.CellBackColor = &HC0C0FF
       Next k
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub llenarGrid_detalle(ByVal Grilla As MSHFlexGrid, ByVal in_solicitud As String)
On Error GoTo salir
Dim tTotal As Double
tTotal = 0
strCadena = "SELECT * FROM solicitud_dinero_detalle WHERE id_solicitud='" & Val(in_solicitud) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.TxtMonto.Text = 0
    Grilla.Rows = 0
    Me.CmdQuitar.Visible = False
    Grilla.Rows = 0
    Exit Sub

End If
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 7500
           Grilla.ColWidth(3) = 1500
       Next
         cabecera = "IDITEM" & vbTab & "ITEM" & vbTab & "DESCRIPCION" & vbTab & "MONTO"
         Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        
        For i = 1 To rst.RecordCount
        tTotal = tTotal + rst("monto")
             Fila = rst("id_detalle") & vbTab & formato_item(i, 2) & vbTab & rst("descripcion") & vbTab & Format(rst("monto"), "#,##0.00")
             Grilla.AddItem Fila
             Fila = ""
        rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "***********  TOTAL A SOLICITAR  ***********" & vbTab & Format(tTotal, "###0.00")
       Grilla.AddItem Fila
       Me.TxtMonto.Text = Format(tTotal, "###0.00")
       For k = 0 To 3
            Grilla.col = k
            Grilla.Row = i
            Grilla.CellBackColor = &HC0C0FF
       Next k
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub







Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)


End Sub

Private Sub txtdescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_solicitud_viatico WHERE resumen LIKE '%" & Trim(Me.txtDescripcion.Text) & "%' and  ruc='" & KEY_RUC & "'"
    Call llenarGrid(Me.HfgDetalle, Me)
End If
End Sub

Private Sub TxtDetalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtMontoMotivo)
End If
End Sub

Private Sub TxtMontoMotivo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
     If Trim(Me.TxtDetalle.Text) <> "" And Val(Me.TxtMontoMotivo.Text) > 0 Then
        strCadena = "INSERT INTO solicitud_dinero_temporal(detalle,monto,dni_save,ruc)VALUES('" & UCase(Trim(Me.TxtDetalle.Text)) & "','" & Val(Me.TxtMontoMotivo.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        Call llenarGrid_temp(Me.HfDetalle)
        Me.TxtDetalle.Text = ""
        Me.TxtMontoMotivo.Text = 0
        Call Resalta(Me.TxtDetalle)
        Me.cmdprocesar_detalle.Enabled = True
    
    
    End If
End If


End Sub

Private Sub TxtResumen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtDetalle)
End If
End Sub

Private Sub txtsolicitud_KeyPress(KeyAscii As Integer)
strCadena = "SELECT * FROM view_solicitud_viatico WHERE solicitud LIKE '%" & Trim(Me.txtsolicitud.Text) & "%' ruc='" & KEY_RUC & "'"
Call llenarGrid(Me.HfgDetalle, Me)
End Sub

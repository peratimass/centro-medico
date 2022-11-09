VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmDetalleventa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   13290
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmFormaPago 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7800
      TabIndex        =   24
      Top             =   3120
      Visible         =   0   'False
      Width           =   5055
      Begin MSDataListLib.DataCombo DtcFormaPagodetalle 
         Height          =   315
         Left            =   960
         TabIndex        =   26
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   12582912
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
      Begin VitekeySoft.ChameleonBtn cmdProcesar 
         Height          =   330
         Left            =   4005
         TabIndex        =   27
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         BTYPE           =   5
         TX              =   "CAMBIAR"
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
         BCOL            =   8454016
         BCOLO           =   8454016
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleventa.frx":0000
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
         Height          =   315
         Left            =   960
         TabIndex        =   29
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   12582912
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
      Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
         Height          =   255
         Left            =   4680
         TabIndex        =   30
         Top             =   120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
         MICON           =   "FrmDetalleventa.frx":001C
         PICN            =   "FrmDetalleventa.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblid_registro 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4080
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   30
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "TIPO    :"
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
         TabIndex        =   28
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "DETALLE :"
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
         TabIndex        =   25
         Top             =   720
         Width           =   645
      End
   End
   Begin VB.TextBox txtkey 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   2160
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtmail 
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
      Height          =   405
      Left            =   7800
      TabIndex        =   21
      Top             =   6480
      Width           =   3255
   End
   Begin VB.TextBox txtObservacion 
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
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   6000
      Width           =   6495
   End
   Begin VitekeySoft.ChameleonBtn cmdimprimir 
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   4920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
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
      MICON           =   "FrmDetalleventa.frx":2EEC
      PICN            =   "FrmDetalleventa.frx":2F08
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
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   5106
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPagos 
      Height          =   1095
      Left            =   7800
      TabIndex        =   1
      Top             =   4800
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1931
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
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   375
      Left            =   12840
      TabIndex        =   13
      Top             =   60
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      MICON           =   "FrmDetalleventa.frx":54D9
      PICN            =   "FrmDetalleventa.frx":54F5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdEnviarMail 
      Height          =   495
      Left            =   11160
      TabIndex        =   20
      Top             =   6480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "ENVIAR MAIL"
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
      MICON           =   "FrmDetalleventa.frx":83A9
      PICN            =   "FrmDetalleventa.frx":83C5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdRegenerar 
      Height          =   615
      Left            =   2160
      TabIndex        =   23
      Top             =   4920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "RE-GENERAR ASIENTO CONTABLE"
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
      BCOL            =   8454016
      BCOLO           =   8454016
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetalleventa.frx":C8E8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcVendedor 
      Height          =   315
      Left            =   7800
      TabIndex        =   32
      Top             =   6000
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   12582912
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
   Begin VitekeySoft.ChameleonBtn cmdCambiarVendedor 
      Height          =   345
      Left            =   11160
      TabIndex        =   33
      Top             =   6000
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "CAMBIAR"
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
      MICON           =   "FrmDetalleventa.frx":C904
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdprocesarVendedor 
      Height          =   345
      Left            =   12000
      TabIndex        =   34
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetalleventa.frx":C920
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblmoneda 
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
      Height          =   300
      Left            =   10920
      TabIndex        =   19
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "MONEDA  :"
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
      Left            =   9510
      TabIndex        =   18
      Top             =   480
      Width           =   750
   End
   Begin VB.Label lblcomprobante 
      BackColor       =   &H000080FF&
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
      Height          =   300
      Left            =   9360
      TabIndex        =   17
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "OBSERVACION :"
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
      Left            =   240
      TabIndex        =   16
      Top             =   5760
      Width           =   1185
   End
   Begin VB.Label lblVenta 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   7440
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "DNI / RUC :"
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
      Left            =   600
      TabIndex        =   11
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "DIRECCION :"
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
      Left            =   480
      TabIndex        =   10
      Top             =   840
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "RAZON SOCIAL :"
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
      Left            =   180
      TabIndex        =   9
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "FECHA EMISION :"
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
      Left            =   9105
      TabIndex        =   8
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "T.CAMBIO :"
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
      Left            =   9480
      TabIndex        =   7
      Top             =   1200
      Width           =   780
   End
   Begin VB.Label lblruc 
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
      Height          =   300
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblrazonsocial 
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
      Height          =   300
      Left            =   1680
      TabIndex        =   5
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label lbldireccion 
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
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label lblfecha 
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
      Height          =   300
      Left            =   10920
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblcambio 
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
      Height          =   300
      Left            =   10920
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   4575
      Left            =   0
      Top             =   0
      Width           =   13290
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   2415
      Left            =   0
      Top             =   4680
      Width           =   13290
   End
End
Attribute VB_Name = "FrmDetalleventa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede


Private Sub ChameleonBtn1_Click()
Me.FrmFormaPago.Visible = False
End Sub

Private Sub cmdCambiarVendedor_Click()

Procedencia = modificar_credito
Call disabled_form(Me)
frmsegurity.Show
Exit Sub

End Sub

Private Sub cmdEnviarMail_Click()
Dim IN_asunto As String
If InStr(1, Trim(Me.txtmail.Text), "@") > 0 Then
    IN_asunto = "KEYFACIL :" & "COMPROBANTE ELECTRONICO :" & Trim(Me.lblComprobante.Caption)
    Call enviar_mail_facturacion(Trim(Me.txtkey.Text), IN_asunto, Trim(Me.txtmail.Text))
Else
    MsgBox "Ingrese un mail Valido.!!", vbInformation, KEY_VENDEDOR
End If
End Sub
Public Sub enviar_mail_electronico(ByVal strHtml As String)

Call enabled_form(Me)

End Sub


Public Function enviar_mail_facturacion(ByVal in_key As String, ByVal IN_asunto As String, ByVal in_mail As String)
Call disabled_form(Me)
Procedencia = mailenviar
FrmLoad_web_service.Show
FrmLoad_web_service.nom_prcedimiento = "enviar_mail_electronico"
Set FrmLoad_web_service.FormPadre = Me
If KEY_SERVIDOR_CLOUD = "si" Then
     
     
     If Len(in_key) > 36 Then
        in_abc = "5e8836bb8b3445ac15f0c0c5a22815d31cd7e84df87b5b60f40ff612cff7d41c"
        Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/enviarxml", "POST", json_facturacion_electronica_mail(KEY_RUC, in_key, IN_asunto, in_mail), "{x-api-token: '" & in_abc & "', x-api-produccion: 'yes'}")
     Else
        Call FrmLoad_web_service.crear_json_facturacion_electronica("https://api.vitekey.com/keyfact/erp/invoice-send-email?password=vitekey2018", "POST", json_facturacion_electronica_mail(KEY_RUC, in_key, IN_asunto, in_mail), "{x-api-token: '" & KEY_TOKEN_CLOUD & "', x-api-produccion: 'yes'}")
     
     End If
     
    
Else
    '-
    
    Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/enviarxml", "POST", json_facturacion_electronica_mail(KEY_RUC, in_key, IN_asunto, in_mail), "{x-api-token: '" & KEY_TOKEN_LOCAL & "', x-api-produccion: 'yes'}")
End If
End Function



Private Sub cmdImprimir_Click()
strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & Val(Me.lblVenta.Caption) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Call Orden_Impresion(rst("id_doc"), Trim(rst("serie")), Trim(rst("numero")), rst("id_tipo_factura"), Val(Me.lblVenta.Caption))
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
       
       
       
       strCadena = "SELECT * FROM con_asiento WHERE Glosa LIKE '%" & Trim(in_numero) & "%' and IdTipoAsiento IN('1CIX000000000137','1CIX000000000053','1CIX000000000055') and idEmpresaSis='" & KEY_RUC & "' "
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
                       strCadena = "DELETE FROM  con_asientomovimiento  where Id = '" & rstA("id") & "' and idEmpresaSis='" & KEY_RUC & "'"
                       CnBd.Execute (strCadena)
                       
                       strCadena = "DELETE FROM  con_asientomovimiento_documento  where IdAsientoMovimiento = '" & rstA("id") & "' and idEmpresaSis='" & KEY_RUC & "'"
                       CnBd.Execute (strCadena)
                       
                       strCadena = "DELETE FROM  CON_MovimientoCajaBanco  where IdAsientoMovimiento = '" & rstA("id") & "' and idEmpresaSis='" & KEY_RUC & "'"
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

Private Sub cmdProcesar_Click()

Call put_procesar

End Sub

Public Sub put_procesar()

strCadena = "UPDATE movimiento_venta_monto SET forma_pago='" & Me.DtcFormaPago.BoundText & "',id_forma_pago='" & Me.DtcFormapagodetalle.BoundText & "'  WHERE id_detalle='" & Me.lblid_registro.Caption & "' LIMIT 1"
CnBd.Execute (strCadena)



Call llena_pagosVenta(Me.HfPagos, FrmReporteRegistroVentas.HfdPersona.TextMatrix(FrmReporteRegistroVentas.HfdPersona.Row, 0))

Me.FrmFormaPago.Visible = False



End Sub


Private Sub put_vendedor()
    strCadena = "UPDATE movimiento_venta SET id_vendedor='" & Me.DtcVendedor.BoundText & "' WHERE id_venta='" & Val(lblVenta.Caption) & "' LIMIT 1"
    CnBd.Execute (strCadena)
    MsgBox "Cambio Realizado", vbInformation
    Me.cmdprocesarVendedor.Visible = False
End Sub


Private Sub cmdprocesarVendedor_Click()
Call put_vendedor
End Sub

Private Sub cmdRegenerar_Click()
strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & Val(Me.lblVenta.Caption) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    
       
    
    in_doc = ""
        If rst("id_doc") = "0001" Then
           in_doc = "FACTURA:"
        End If
         If rst("id_doc") = "0003" Then
           in_doc = "BOLETA:"
        End If
         If rst("id_doc") = "0007" Then
         in_doc = "NC:"
        End If
        in_glosa = Trim(in_doc & rst("serie") & rst("numero") & Space(1) & rst("ncliente"))
        
        Call delete_asiento(rst("id_venta"), in_glosa)
        
        strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
           rstL.MoveFirst
           For j = 0 To rstL.RecordCount - 1
                
                strCadena = "SELECT costo_promedio FROM kardex WHERE fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' and id_producto='" & rstL("id_producto") & "' and id_movimiento='" & rst("id_venta") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                Call ConfiguraRstA(strCadena)
                If rstA.RecordCount > 0 Then
                    in_costo_promedio = rstA(0)
                Else
                    in_costo_promedio = 0
                End If
            
               strCadena = "UPDATE movimiento_venta_detalle SET precio_costo='" & in_costo_promedio & "' WHERE id_detalle_venta='" & rstL("id_detalle_venta") & "' "
               CnBd.Execute (strCadena)
               
               rstL.MoveNext
           Next j
        End If
        
        
        strCadena = "call P_insert_venta_agenda_test('" & Val(rst("id_venta")) & "')"
        CnBd.Execute (strCadena)
        MsgBox "Proceso Realizado", vbInformation
        
End If

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub


Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub DtcFormaPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     
    strCadena = "SELECT id_registro as Codigo, CONCAT(descripcion,'-',observacion) as Descripcion FROM forma_pago_detalle WHERE  id_alm='" & KEY_ALM & "'  and  id='" & Me.DtcFormaPago.BoundText & "' AND ruc='" & KEY_RUC & "' AND estado='si' ORDER BY id_detalle"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcFormapagodetalle)
     
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 200
Me.lblVenta.Caption = 0
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad  WHERE  id_personal='si' and habilitado='si' and  ruc='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcVendedor)
  Me.DtcVendedor.BoundText = 0





If FrmReporteRegistroVentas.Procedencia = Selecionar Then
    Call llenar(FrmReporteRegistroVentas.HfdPersona.TextMatrix(FrmReporteRegistroVentas.HfdPersona.Row, 0))
    Call llena_pagosVenta(Me.HfPagos, FrmReporteRegistroVentas.HfdPersona.TextMatrix(FrmReporteRegistroVentas.HfdPersona.Row, 0))
End If



If frmtracking.Procedencia = buscar Then
        Call llenar(frmtracking.HfdPersona.TextMatrix(frmtracking.HfdPersona.Row, 0))
        Call llena_pagosVenta(Me.HfPagos, frmtracking.HfdPersona.TextMatrix(frmtracking.HfdPersona.Row, 0))
End If


End Sub
Public Sub llena_pagosVenta(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)

On Error GoTo salir
Dim tpago As Double
strCadena = "SELECT M.id_detalle,CONCAT(F.descripcion,'-',F.observacion) as descripcion,M.monto FROM movimiento_venta_monto M,forma_pago_detalle F WHERE M.id_forma_pago=F.id_registro AND id_venta='" & idVenta & "' AND M.ruc='" & KEY_RUC & "' AND F.ruc='" & KEY_RUC & "' "
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
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 1500
       Next
        cabecera = "CODIGO" & vbTab & "FORMA PAGO" & vbTab & "MONTO"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tpago = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_detalle") & vbTab & rst("descripcion") & vbTab & Format(rst("monto"), "###0.00")
            Grilla.AddItem Fila
            tpago = rst("monto") + tpago
            rst.MoveNext
    Next i
    
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Private Sub llenar(ByVal id_venta As Double)
Me.lblVenta.Caption = id_venta
strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & id_venta & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Me.lblComprobante.Caption = rst("documento") & Space(2) & rst("id_venta")
Me.lblruc.Caption = rst("id_cliente")
Me.TxtObservacion.Text = rst("observacion")
Me.lblrazonsocial.Caption = rst("ncliente")
Me.lbldireccion.Caption = BDBuscarCampo("persona", "direccion", "dni", rst("id_cliente"))
Me.lblfecha.Caption = str(rst("fecha_emision"))
Me.lblcambio.Caption = str(rst("tc"))
If IsNull(rst("sunat_key")) = False Then
    Me.txtkey.Text = rst("sunat_key")
Else
    Me.txtkey.Text = ""
End If

If rst("id_moneda") = "00002" Then
   Me.lblmoneda.Caption = "DOLARES"
Else
    Me.lblmoneda.Caption = "SOLES"
End If
Me.DtcVendedor.BoundText = rst("id_vendedor")

Call llenarGrid_Comprobante(Me.HfdDetalle, rst("id_venta"))
Call llena_pagosVenta(Me.HfPagos, id_venta)
End Sub
Sub llenarGrid_Comprobante(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo salir
Dim tTotal As Double
Dim texonerado As Double
Dim tafecto As Double
Dim in_obsequio As Single
strCadena = "SELECT * FROM view_detalle_venta WHERE id_venta='" & idVenta & "' and ruc='" & KEY_RUC & "' ORDER BY id_detalle_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    
    Grilla.Rows = 0
    in_obsequio = 0
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 5500
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 0
           'Grilla.ColAlignment(4) = 7
        Next
        cabecera = "IDDETALLE" & vbTab & "CODIGO" & vbTab & "DESCRIPCION " & vbTab & "UND " & vbTab & "CANTIDAD" & vbTab & "PRECIO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        in_obsequio = 0
        For i = 0 To rst.RecordCount - 1
            If rst("id_producto") = KEY_COD_PER Then
               in_producto = ""
               in_unidad = ""
               If rst("cantidad") = 0 Then
                  in_cantidad = ""
                Else
                  in_cantidad = Format(rst("cantidad"), "###0.00")
               End If
               
               If rst("precio") = 0 Then
                  in_precio = ""
                Else
                  in_precio = Format(rst("precio"), "###0.00")
               End If
               
            Else
              in_producto = rst("id_producto")
              in_unidad = rst("abreviatura")
              in_cantidad = Format(rst("cantidad"), "###0.00")
              in_precio = Format(rst("precio"), "###0.00")
            End If
            
            
            
            Fila = rst("id_detalle_venta") & vbTab & in_producto & vbTab & rst("detalle") & vbTab & in_unidad & vbTab & in_cantidad & vbTab & in_precio & vbTab & Format(rst("total"), "###0.00")
            Grilla.AddItem Fila
            If (Trim(rst("id_igv")) = "no") Then
                            texonerado = texonerado + rst("total")
                            For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0FFFF
                            Next k
             Else
                            tafecto = tafecto + rst("total")
             End If
             If rst("obsequio") = "si" Then
                in_obsequio = in_obsequio + in_precio * in_cantidad
                For k = 3 To 6
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
             End If
             
            
            rst.MoveNext
    Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1


Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Sub llenarGridDetalle(ByVal Grilla As MSHFlexGrid, ByVal id_venta As Double)
'On Error GoTo salir
    strCadena = "SELECT D.id_producto as CODIGO,U.abreviatura AS UND,P.nombre_prod AS DESCRIPCION,D.cantidad as CANT,D.precio as PRECIO,D.total as TOTAL FROM movimiento_venta_detalle D,producto P,unidad U WHERE U.id_und=P.id_unidad AND U.id_usu='" & KEY_RUC & "' AND  D.id_producto=P.id_producto AND D.id_venta='" & id_venta & "' AND D.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "'"
    
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.Rows = rst.RecordCount
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 800
  Grilla.ColWidth(1) = 1200
  Grilla.ColWidth(2) = 4500
  Grilla.ColWidth(3) = 1200
  Grilla.Refresh
   'Me.HfgDetalle.SetFocus
 ' Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub HfPagos_DblClick()

Procedencia = modificar
Call disabled_form(Me)
frmsegurity.Show
Exit Sub


End Sub


Public Sub put_formapago(ByVal in_registro As String)
    
    strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM forma_pago ORDER BY id ASC "
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcFormaPago)
    
    
    
    strCadena = "SELECT id_registro as Codigo, CONCAT(descripcion,'-',observacion) as Descripcion FROM forma_pago_detalle WHERE  id_alm='" & KEY_ALM & "'  and  id='" & Me.DtcFormaPago.BoundText & "' AND ruc='" & KEY_RUC & "' AND estado='si' ORDER BY id_detalle"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcFormapagodetalle)
    Me.FrmFormaPago.Visible = True
    
    
    
End Sub



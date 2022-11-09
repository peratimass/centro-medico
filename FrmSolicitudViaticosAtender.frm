VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmSolicitudViaticosAtender 
   BorderStyle     =   0  'None
   Caption         =   "MOVIMIENTOS"
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   18690
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   18690
   ShowInTaskbar   =   0   'False
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
      Left            =   5685
      TabIndex        =   50
      Top             =   5160
      Width           =   975
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
      Left            =   5685
      TabIndex        =   49
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox Txtid_solicitud 
      Height          =   285
      Left            =   360
      TabIndex        =   45
      Top             =   720
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.TextBox TxtDireccion 
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
      Left            =   1605
      MaxLength       =   80
      TabIndex        =   28
      Top             =   2200
      Width           =   5535
   End
   Begin VB.TextBox TxtCliente 
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
      Left            =   1605
      MaxLength       =   80
      TabIndex        =   27
      Top             =   1700
      Width           =   5535
   End
   Begin VB.TextBox TxtRuc 
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
      Left            =   1620
      MaxLength       =   80
      TabIndex        =   26
      Top             =   1260
      Width           =   1815
   End
   Begin VB.TextBox TxtSerie 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   8445
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   25
      Text            =   "000"
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox TxtNumeroDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   10125
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   24
      Text            =   "000000"
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox TxtSaldo 
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
      Left            =   1605
      MaxLength       =   80
      TabIndex        =   23
      Top             =   3135
      Width           =   1935
   End
   Begin VB.TextBox TxtMontoPago 
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
      Left            =   1605
      MaxLength       =   80
      TabIndex        =   0
      Top             =   3660
      Width           =   1935
   End
   Begin VB.Frame FrmCheque 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PAGAR CON CHEQUE"
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
      Height          =   1815
      Left            =   1485
      TabIndex        =   18
      Top             =   6480
      Width           =   5655
      Begin VB.CommandButton cmdCargarCheque 
         Caption         =   "CARGAR CHEQUE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton OptChequeNO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NO"
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
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton OptChequeSi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SI"
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
         Left            =   840
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin MSDataListLib.DataCombo DtcCheque 
         Height          =   315
         Left            =   1560
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CUENTA DESTINO PROVEEDOR"
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
      Left            =   10005
      TabIndex        =   7
      Top             =   1800
      Width           =   8055
      Begin VB.TextBox TxtCuentaBancaria 
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
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   10
         Top             =   2445
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   4200
         TabIndex        =   9
         Top             =   2685
         Width           =   375
      End
      Begin VB.TextBox TxtOperacion 
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
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   8
         Top             =   2880
         Width           =   2415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MshCuentasBancarias 
         Height          =   1215
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2143
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
      Begin MSDataListLib.DataCombo DtcBanco 
         Height          =   315
         Left            =   1680
         TabIndex        =   12
         Top             =   1545
         Width           =   2415
         _ExtentX        =   4260
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
      Begin MSDataListLib.DataCombo DtcMonedaCuenta 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   1995
         Width           =   2415
         _ExtentX        =   4260
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUENTA BANCARIA:"
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
         Left            =   75
         TabIndex        =   17
         Top             =   2460
         Width           =   1515
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N.OPERACION :"
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
         TabIndex        =   16
         Top             =   2940
         Width           =   1185
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BANCO   :"
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
         Left            =   855
         TabIndex        =   15
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONEDA :"
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
         Left            =   795
         TabIndex        =   14
         Top             =   2040
         Width           =   825
      End
   End
   Begin VB.TextBox TxtTc 
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
      Height          =   285
      Left            =   5205
      MaxLength       =   80
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GLOSA"
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
      Height          =   1695
      Left            =   9960
      TabIndex        =   3
      Top             =   5280
      Width           =   8175
      Begin VB.TextBox TxtObservacion 
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
         Height          =   1275
         Left            =   120
         MaxLength       =   80
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   7695
      End
   End
   Begin VB.Frame frmitf 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ITF"
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
      Left            =   4245
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox TxtItf 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   600
         MaxLength       =   80
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComCtl2.DTPicker DtpEmision 
      Height          =   300
      Left            =   10920
      TabIndex        =   5
      Top             =   1280
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
      Format          =   142213121
      CurrentDate     =   41130
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   330
      Left            =   285
      TabIndex        =   29
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
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
   Begin MSDataListLib.DataCombo DtcCuentas 
      Height          =   330
      Left            =   1605
      TabIndex        =   30
      Top             =   4680
      Width           =   5535
      _ExtentX        =   9763
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
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   330
      Left            =   8445
      TabIndex        =   31
      Top             =   240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   582
      _Version        =   393216
      Locked          =   -1  'True
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
      ForeColor       =   8388608
      ListField       =   "0000º"
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
      Left            =   1605
      TabIndex        =   32
      Top             =   2655
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSComCtl2.DTPicker DtpValor 
      Height          =   300
      Left            =   13200
      TabIndex        =   33
      Top             =   1275
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
      Format          =   142278657
      CurrentDate     =   41130
   End
   Begin VitekeySoft.ChameleonBtn cmdprocesar_detalle 
      Height          =   825
      Left            =   15000
      TabIndex        =   46
      Top             =   7320
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
      MICON           =   "FrmSolicitudViaticosAtender.frx":0000
      PICN            =   "FrmSolicitudViaticosAtender.frx":001C
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
      Left            =   17160
      TabIndex        =   47
      Top             =   7320
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
      MICON           =   "FrmSolicitudViaticosAtender.frx":3664
      PICN            =   "FrmSolicitudViaticosAtender.frx":3680
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdImprimir 
      Height          =   825
      Left            =   16080
      TabIndex        =   48
      Top             =   7320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
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
      MICON           =   "FrmSolicitudViaticosAtender.frx":3A70
      PICN            =   "FrmSolicitudViaticosAtender.frx":3A8C
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
      Left            =   1605
      TabIndex        =   51
      Top             =   5160
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
      Left            =   1605
      TabIndex        =   52
      Top             =   5760
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
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FORMA DE PAGO :"
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
      Left            =   90
      TabIndex        =   54
      Top             =   5280
      Width           =   1395
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO DE FLUJO :"
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
      Left            =   270
      TabIndex        =   53
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONEDA :"
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
      Left            =   660
      TabIndex        =   44
      Top             =   2640
      Width           =   825
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/DNI :"
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
      Left            =   690
      TabIndex        =   43
      Top             =   1260
      Width           =   795
   End
   Begin VB.Label lblAnulado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   885
      Left            =   14520
      TabIndex        =   42
      Top             =   120
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUENTA ORIGEN:"
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
      Left            =   150
      TabIndex        =   41
      Top             =   4700
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAZON SOCIAL:"
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
      Left            =   300
      TabIndex        =   40
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION :"
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
      TabIndex        =   39
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SALDO :"
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
      Left            =   870
      TabIndex        =   38
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONTO A PAGAR :"
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
      Left            =   60
      TabIndex        =   37
      Top             =   3720
      Width           =   1425
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T.C:"
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
      Left            =   4770
      TabIndex        =   36
      Top             =   2700
      Width           =   345
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMISION:"
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
      Left            =   10065
      TabIndex        =   35
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR:"
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
      Left            =   12600
      TabIndex        =   34
      Top             =   1320
      Width           =   585
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   975
      Left            =   165
      Top             =   120
      Width           =   12615
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8370
      Left            =   0
      Top             =   0
      Width           =   18690
   End
End
Attribute VB_Name = "FrmSolicitudViaticosAtender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mostrar As Boolean
Dim codigo_P As String
Public Procedencia As EnumProcede
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Dim tTotal As Double
tTotal = 0

strCadena = "SELECT      movimiento_caja.doc_cod,movimiento_caja.fecha_valor, Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, " & _
"movimiento_caja.descripcion_per , movimiento_caja.Monto,movimiento_caja.anulado " & _
"FROM         movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod " & _
" WHERE  movimiento_caja.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "'  ORDER BY  movimiento_caja.numero DESC"
Call ConfiguraRst(strCadena)
 
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
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
           Grilla.ColWidth(0) = 600
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 600
           Grilla.ColWidth(4) = 800
           Grilla.ColWidth(5) = 3000
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 0
           Grilla.ColWidth(8) = 0
          Next
         cabecera = "ITEM" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "SERIE" & vbTab & "NUMERO" & vbTab & "PERSONA" & vbTab & "MONTO" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 8
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = Fila & str(rst.RecordCount - i + 1) & vbTab & rst("fecha_valor") & vbTab & rst("doc_abrev") & vbTab & rst("serie") & vbTab & rst("numero") & vbTab & rst("descripcion_per") & vbTab & Format(rst("monto"), "#,##0.00")
            If (Fila = "") Then
                X = 1
            End If
            
          Grilla.AddItem Fila
            If (Trim(rst("anulado")) = "si") Then
                            For k = 0 To 8
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &H8080FF
                            Next k
        End If
        Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub ChkAdelantado_Click()

End Sub



Private Sub ChameleonBtn1_Click()

End Sub

Private Sub cmdCargarCheque_Click()
Dim glosa As String
Procedencia = 1
If Val(Me.TxtMontoPago.Text) > 0 Then
        strCadena = "DELETE FROM cheque_detalle WHERE id_cheque='" & Val(Me.DtcCheque.BoundText) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        strCadena = "DELETE FROM cheque_factura WHERE id_cheque='" & Val(Me.DtcCheque.BoundText) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        
        strCadena = "INSERT INTO cheque_detalle(id_cheque,detalle,monto,ruc)VALUES('" & Val(Me.DtcCheque.BoundText) & "','" & Trim(Me.txtObservacion.Text) & "','" & Val(Me.TxtMontoPago.Text) & "','" & KEY_RUC & "')"
        Call CnBd.Execute(strCadena)
        strCadena = "SELECT * FROM solicitud_dinero WHERE dni='" & Me.txtruc.Text & "' AND ruc='" & KEY_RUC & "' AND saldo>0 AND anulado='no'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            rst.MoveFirst
            For i = 0 To rst.RecordCount - 1
                strCadena = "INSERT INTO cheque_factura(id_cheque,id_compra,ruc)VALUES('" & Me.DtcCheque.BoundText & "','" & rst("id_solicitud") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                 
                rst.MoveNext
            Next i
        End If
        FrmChequeNuevo.Txtcentrocosto.Text = "42121"
        FrmChequeNuevo.lblcostos.Text = "FACTURAS POR PAGAR"
        FrmChequeNuevo.txtMotivo.Text = ""
        FrmChequeNuevo.TxtMontoMotivo.Text = ""
        FrmChequeNuevo.txtruc.Text = Me.txtruc.Text
        FrmChequeNuevo.txtrazonsocial.Text = Me.txtCliente.Text
        FrmChequeNuevo.txtDireccion.Text = Me.txtDireccion.Text
        Call Resalta(FrmChequeNuevo.txtMotivo)
        
    End If

'FrmChequeNuevo.Show
'FrmChequeNuevo.TxtidCheque.text = Me.DtcCheque.BoundText
'Call FrmChequeNuevo.llenar_cheque(Me.DtcCheque.BoundText)

End Sub

Private Sub cmdImprimir_Click()
strCadena = "SELECT     Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, movimiento_caja.cPersona, " & _
        "movimiento_caja.descripcion_per, Persona.sDireccionCliente1, Persona.Per_Ruc, movimiento_caja.fecha_valor," & _
        "movimiento_caja.cambio , movimiento_caja.glosa, centro_costos.descripcion, movimiento_caja.Monto,movimiento_caja.monto_letras " & _
        "FROM movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
        "centro_costos ON movimiento_caja.id_costo = centro_costos.id_costo INNER JOIN " & _
        "Persona ON movimiento_caja.cPersona = Persona.cPersona WHERE movimiento_caja.serie='" & Trim(Me.txtSerie.Text) & "' AND movimiento_caja.numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND movimiento_caja.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND Ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptReciboIngreso", , App.Path + "\Reportes\")
        Set rst = Nothing
        
        strCadena = "SELECT     Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, movimiento_caja.cPersona, " & _
        "movimiento_caja.descripcion_per, Persona.sDireccionCliente1, Persona.Per_Ruc, movimiento_caja.fecha_valor," & _
        "movimiento_caja.cambio , movimiento_caja.glosa, centro_costos.descripcion, movimiento_caja.Monto,movimiento_caja.monto_letras " & _
        "FROM movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
        "centro_costos ON movimiento_caja.id_costo = centro_costos.id_costo INNER JOIN " & _
        "Persona ON movimiento_caja.cPersona = Persona.cPersona WHERE movimiento_caja.serie='" & Trim(Me.txtSerie.Text) & "' AND movimiento_caja.numero='" & Trim(Me.TxtNumeroDoc.Text) & "'"
        
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptReciboIngreso", , App.Path + "\Reportes\")
    
End Sub

Private Sub cmdprocesar_detalle_Click()
Call Save
End Sub

Private Sub cmdsalir_detalle_Click()
Unload Me
End Sub

Private Sub Command1_Click()
If Len(Me.TxtCuentaBancaria.Text) > 5 Then
    strCadena = "SELECT * FROM persona_cuentabancaria WHERE dni='" & Trim(Me.txtruc.Text) & "' AND cuenta='" & Trim(Me.TxtCuentaBancaria.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        strCadena = "INSERT INTO persona_cuentabancaria(dni,id_banco,id_moneda,cuenta)VALUES('" & Trim(Me.txtruc.Text) & "','" & Me.DtcBanco.BoundText & "','" & Me.DtcMonedaCuenta.BoundText & "','" & Trim(Me.TxtCuentaBancaria.Text) & "') "
        CnBd.Execute (strCadena)
         
        Call llenar_cuentas(Me.MshCuentasBancarias, Trim(Me.txtruc.Text))
    Else
    MsgBox "CUENTA YA REGISTRADA", vbInformation, KEY_EMPRESA
    End If
End If
End Sub



Private Sub DtcCuentas_Change()
Dim ssaldo As Double, residuo As Single, sitf As Single
ssaldo = Val(Me.txtsaldo.Text)
strCadena = "SELECT * FROM mis_cuentas WHERE id_cuenta='" & Me.DtcCuentas.BoundText & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst("id_moneda") <> Me.DtcMoneda.BoundText Then
    If rst("id_moneda") = "00001" Then
        Me.txtsaldo.Text = Format(ssaldo * KEY_CAMBIO, "###0.00")
        Me.TxtMontoPago.Text = Format(ssaldo * KEY_CAMBIO, "###0.00")
        Me.DtcMoneda.BoundText = rst("id_moneda")
    Else
        Me.txtsaldo.Text = Format(ssaldo / KEY_CAMBIO, "###0.00")
        Me.TxtMontoPago.Text = Format(ssaldo / KEY_CAMBIO, "###0.00")
        Me.DtcMoneda.BoundText = rst("id_moneda")
    End If
End If
residuo = Val(Me.TxtMontoPago.Text) Mod 1000

If (Val(Me.TxtMontoPago.Text) - residuo) > 0 Then
    sitf = (Val(Me.TxtMontoPago.Text) - residuo) * 0.005 / 100
Else
    itf = 0#
End If
If rst("id_tipo") = "01" Then
    FrmCheque.Enabled = False
    Me.frmitf.Visible = False
Else
    FrmCheque.Enabled = True
    Me.frmitf.Visible = True
    Me.TxtItf.Text = Format(sitf, "#,##0.00")
End If
End Sub

Private Sub DtcTipoDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Me.DtcTipoDoc.BoundText) <> "0001" And Trim(Me.DtcTipoDoc.BoundText) <> "0003" Then
        Call Resalta(Me.txtSerie)
    Else
        MsgBox "Srta:" + Space(1) + KEY_VENDEDOR + Space(1) + "Solo Salida de Dinero", vbInformation
        Me.DtcTipoDoc.SetFocus
    End If
    
    
End If
End Sub

Public Sub nuevo()
    
End Sub

Private Sub ActualizarAdelanto(ByVal TotalPedido As Double)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Dim id_solicitud As Double
id_solicitud = Val(FrmSolicitudViaticos.HfgDetalle.TextMatrix(FrmSolicitudViaticos.HfgDetalle.Row, 0))
Me.Txtid_solicitud.Text = id_solicitud
Me.cmdCargarCheque.Visible = False
Me.DtpEmision.Value = KEY_FECHA
Me.DtpValor.Value = KEY_FECHA
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
      
  strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0097' AND ruc='" & KEY_RUC & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount < 1 Then
        strCadena = "INSERT INTO almacen_comprobante(ruc,id_alm,id_doc,serie,numero)VALUES ('" & KEY_RUC & "','" & Me.DtcAlmacen.BoundText & "','0097','001','000001')"
        CnBd.Execute (strCadena)
         
        strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='0097' AND ruc='" & KEY_RUC & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "'"
        Call ConfiguraRst(strCadena)
        Me.txtSerie.Text = rst("serie")
        Me.TxtNumeroDoc.Text = rst("numero")
  Else
      Me.txtSerie.Text = rst("serie")
      Me.TxtNumeroDoc.Text = rst("numero")
  End If
  
  strCadena = "SELECT A.id_doc as Codigo,C.doc_abrev as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' ORDER BY C.doc_abrev  "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
  Me.DtcTipoDoc.BoundText = "0097"
   
  
  strCadena = "SELECT * FROM solicitud_dinero S,persona P WHERE id_solicitud='" & id_solicitud & "' AND S.dni=P.dni AND S.ruc='" & KEY_RUC & "' "
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
        Me.txtruc.Text = rst("dni")
        Me.txtCliente.Text = UCase(rst("nombre_completo"))
        Me.txtDireccion.Text = UCase(rst("direccion"))
       
        Me.txtObservacion.Text = rst("resumen")
  End If

  
  strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM moneda  ORDER BY id_moneda ASC"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtcMoneda)
  Me.DtcMoneda.BoundText = rst("id_moneda")
  Me.DtcTipoDoc.Enabled = False
  Me.txtsaldo.Text = Format(rst("saldo"), "###0.00")
  Me.TxtMontoPago.Text = Format(rst("saldo"), "###0.00")
  Call llenar_cuentas(Me.MshCuentasBancarias, rst("dni"))
  'Call llenar_facturas(Me.MshFacturas, Trim(Me.TxtRuc.text))
  strCadena = "SELECT id_cuenta as Codigo,CONCAT(C.descripcion,'-',M.descripcion,'  ',C.numero_cuenta) as Descripcion FROM mis_cuentas C,moneda M WHERE C.id_moneda=M.id_moneda AND ruc='" & KEY_RUC & "' "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCuentas)
  Me.txtTc.Text = KEY_CAMBIO
  strCadena = "SELECT id_banco as Codigo,abreviatura as Descripcion FROM banco ORDER BY abreviatura"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtcBanco)
  strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMonedaCuenta)
   
   ' Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
   ' Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
   Me.cmdprocesar_detalle.Enabled = True
    'Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
    'Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
    strCadena = "SELECT id as Codigo,Descripcion  as Descripcion FROM vw_mediopago_nombre  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFormaPago)

strCadena = "SELECT id as Codigo,Nombre  as Descripcion FROM adm_flujocaja ORDER BY Nombre  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFlujo)
End Sub
Public Sub llenar_cuentas(ByVal Grilla As MSHFlexGrid, ByVal dni As String)
On Error GoTo salir
strCadena = "SELECT B.id_banco,B.abreviatura,M.descripcion,PB.cuenta,M.id_moneda FROM persona_cuentabancaria PB,banco B,moneda M WHERE PB.id_banco=B.id_banco AND PB.id_moneda=M.id_moneda AND PB.dni='" & dni & "'"
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
            Grilla.ColWidth(1) = 1200
            Grilla.ColWidth(2) = 1000
            Grilla.ColWidth(3) = 2000
            Grilla.ColWidth(4) = 0
        Next
        cabecera = "BCO" & vbTab & "MONEDA" & vbTab & "CUENTA"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_banco") & vbTab & rst("abreviatura") & vbTab & rst("descripcion") & vbTab & rst("cuenta") & vbTab & rst("id_moneda")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
        Next i
      
     
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Private Sub LlenarDatosCliente(ByVal numero As String, ByVal Documento As String, ByVal serie As String, ByVal Almacen As String)

End Sub
Private Sub Save()
Dim in_mis_cuenta_det As String
Dim monto_pago As Double, saldof As Double, comprobante As String, monto_pagado As Double, Saldo As Double, id_moneda As String, Documento As String
monto_pago = Val(Me.TxtMontoPago.Text)
saldof = 0
        If Me.OptChequeSi.Value = False Then
            strCadena = "SELECT * FROM solicitud_dinero WHERE id_solicitud='" & Val(Me.Txtid_solicitud.Text) & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                Documento = "S-" + Format(rst("id_solicitud"), "000000")
                '------ VERIFICAR MONEDA
                id_moneda = Trim(BDBuscarCampo("mis_cuentas", "id_moneda", "id_cuenta", Me.DtcCuentas.BoundText))
                If rst("id_moneda") <> id_moneda Then
                        If rst("id_moneda") = "00001" Then
                               Saldo = rst("saldo") / KEY_CAMBIO
                        Else
                               Saldo = rst("saldo") * KEY_CAMBIO
                        End If
                  Else
                    Saldo = rst("saldo")
                End If
                '-------END
                
                
                For i = 0 To rst.RecordCount - 1
                    If monto_pago > 0 Then
                        If monto_pago >= Saldo Then
                            saldof = 0
                            monto_pagado = Saldo
                            monto_pago = monto_pago - Saldo
                        Else
                            saldof = Saldo - monto_pago
                            monto_pagado = monto_pago
                            monto_pago = 0
                        End If
                    
                    
                    
                    in_mis_cuenta_det = procesar_transaccion_caja(KEY_ALM, Me.DtcCuentas.BoundText, Format(Me.DtpEmision.Value, "YYYY-mm-dd"), "00002", Trim(Me.txtruc.Text), Trim(Me.txtCliente.Text), Trim(Me.txtObservacion.Text), monto_pagado, "0", "0", Val(Me.Txtid_solicitud.Text), Documento, Val(Me.txtTc.Text), Trim(Me.txtOperacion.Text), Me.DtcFormaPago.BoundText, Me.DtcFlujo.BoundText, Me.DtcMoneda.BoundText, KEY_USUARIO, KEY_RUC)
                    
                    strCadena = "call sp_cuentas_rendir('" & in_mis_cuenta_det & "', '" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                     
                    If rst("id_moneda") <> id_moneda Then
                        If rst("id_moneda") = "00001" Then
                               saldof = saldof * KEY_CAMBIO
                        Else
                            saldof = saldof / KEY_CAMBIO
                        End If
                    End If
                    strCadena = "UPDATE solicitud_dinero SET saldo='" & saldof & "',atendido='si',fecha_confirmacion='" & KEY_FECHA & "',hora_confirmacion='" & str(Time) & "' WHERE id_solicitud='" & Val(Me.Txtid_solicitud.Text) & "' AND ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                     
                    
                   End If
                   rst.MoveNext
                Next i
                Call FrmSolicitudViaticos.actualizar
            End If
            
        End If
        
        
        
        nuevo_numero = formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6)
        strCadena = "UPDATE  almacen_comprobante SET numero='" & Trim(nuevo_numero) & "'  WHERE (serie='" & Trim(Me.txtSerie.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "')"
        CnBd.Execute (strCadena)

        Me.cmdprocesar_detalle.Enabled = False
        Me.cmdImprimir.Enabled = True
        
        
        

        Exit Sub
    

End Sub




Private Sub MshCuentasBancarias_SelChange()
If Trim(Me.MshCuentasBancarias.TextMatrix(Me.MshCuentasBancarias.Row, 2) <> "") Then
    Me.DtcBanco.BoundText = Me.MshCuentasBancarias.TextMatrix(Me.MshCuentasBancarias.Row, 0)
    Me.DtcMonedaCuenta.BoundText = Me.MshCuentasBancarias.TextMatrix(Me.MshCuentasBancarias.Row, 4)
    Me.TxtCuentaBancaria.Text = Me.MshCuentasBancarias.TextMatrix(Me.MshCuentasBancarias.Row, 3)
    Call Resalta(Me.txtOperacion)
Else
    Me.TxtCuentaBancaria.Text = ""
End If
End Sub

Private Sub OptChequeNO_Click()
Me.DtcCheque.Visible = False
Me.cmdCargarCheque.Visible = False
End Sub

Private Sub OptChequeSi_Click()
Me.DtcCheque.Visible = True
strCadena = "SELECT C.id_cheque as Codigo,CONCAT('CHEQUE ',':',C.numero) as Descripcion FROM cheque C,mis_cuentas M,chequera CH  WHERE C.id_chequera=CH.id_chequera AND CH.id_cuenta=M.id_cuenta AND C.seleccionado='si' AND C.ruc='" & KEY_RUC & "' AND M.ruc='" & KEY_RUC & "' AND CH.ruc='" & KEY_RUC & "' AND M.id_cuenta='" & Me.DtcCuentas.BoundText & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
     Call LlenaDataCombo(Me.DtcCheque)
     Me.DtcCheque.SetFocus
     Me.cmdCargarCheque.Visible = True
Else
    Me.cmdCargarCheque.Visible = False
End If
End Sub



Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
      
      
    Case KEY_PRINT
        
  Exit Sub
error:
  
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  
  Exit Sub
  Case KEY_EXIT
    Unload Me
End Select
End Sub

Private Sub Imprimir(ByVal TipoDoc As String, ByVal CodAlm As String, ByVal serie As String, ByVal numero As String)
Dim i As Integer, j As Integer
Dim totalletras As String

    Printer.ScaleMode = vbCharacters 'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "Draft 17cpi"
    Printer.Font.Size = 10
       
If Me.DtcTipoDoc.BoundText = KEY_SALDINER Then
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.PaperSize = 1
    Printer.Print ""
    Printer.Print ""
    'Printer.Print Tab(15); (CVDate(Me.DtpActual.Value))
    Printer.Print ""
    Printer.Print ""
   'Printer.Print Tab(5); Mid(Me.TxtCliente.text + Space(80), 1, 65)
   ' Printer.Print Tab(5); Mid(Me.TxtDireccion.text + Space(80), 1, 65)
   ' Printer.Print Tab(5); Mid(Me.TxtRuc.Text + Space(50), 1, 40) & "SALDINER"; Space(1); Mid(Me.TxtSerie.Text + Space(50), 1, 4) & Space(1) & "-" & Me.TxtNumeroDoc.Text
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); Mid(Trim(KEY_VENDEDOR), 1, 10)
    Printer.Print ""
    Printer.Print ""
   ' Printer.Print Tab(15); "Monto Efectivo:" & "=============" & Space(20) & Me.TxtMontoIngresar.text
    Printer.CurrentY = Printer.CurrentY + 10
    'totalletras = UCase(EnLetras(Me.TxtMontoIngresar.text))
    Set rst = Nothing
    '---- fin totales
    'Printer.Print Tab(10); Mid(totalletras + Space(100), 1, 60)
    Printer.CurrentY = Printer.CurrentY + 0.2
    'Printer.Print Tab(60); Me.TxtMontoIngresar.text
    Printer.EndDoc
    'Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
    Exit Sub
End If
End Sub


Private Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub






Private Sub TxtMontoPago_Change()
Dim sitf As Single, residuo As Single
residuo = Val(Me.TxtMontoPago.Text) Mod 1000

If (Val(Me.TxtMontoPago.Text) - residuo) > 0 Then
    sitf = (Val(Me.TxtMontoPago.Text) - residuo) * 0.005 / 100
Else
    sitf = 0#
End If
Me.TxtItf.Text = Format(sitf, "#,##0.00")
End Sub

Private Sub TxtMontoPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Me.DtcCuentas.SetFocus
End If
End Sub

Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Call Resalta(Me.TxtMontoIngresar)
End If
End Sub


VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmSolicitudViaticosDeclarar 
   BorderStyle     =   0  'None
   Caption         =   "REPORTE DE GASTOS"
   ClientHeight    =   8370
   ClientLeft      =   5055
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
   Begin VB.TextBox txtid_servicio 
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
      Left            =   3645
      TabIndex        =   38
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox TxtLugar 
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
      Left            =   10920
      TabIndex        =   26
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox TxtRazonsocial 
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
      Left            =   14715
      TabIndex        =   25
      Top             =   3720
      Width           =   3555
   End
   Begin VB.TextBox TxtRuc 
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
      Left            =   16125
      TabIndex        =   17
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox TxtNumero 
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
      Left            =   15000
      TabIndex        =   15
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox TxtSerie 
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
      Left            =   14355
      TabIndex        =   14
      Top             =   3360
      Width           =   615
   End
   Begin MSDataListLib.DataCombo DtcComprobante 
      Height          =   330
      Left            =   12525
      TabIndex        =   13
      Top             =   3360
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VB.TextBox TxtMonto 
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
      Left            =   9900
      TabIndex        =   12
      Top             =   3360
      Width           =   975
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   4640
      TabIndex        =   1
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DETALLE SOLICITADO"
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
      Height          =   2760
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   18135
      Begin VB.TextBox TxtId_solicitud 
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
         Height          =   330
         Left            =   2560
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtResumen 
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
         Height          =   555
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox TxtMontoEntregado 
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
         Left            =   2040
         TabIndex        =   11
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox TxtMontoSolicitud 
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
         Left            =   2040
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
         Height          =   2445
         Left            =   8400
         TabIndex        =   3
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4313
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
      Begin MSComCtl2.DTPicker DtpEmision 
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   477822977
         CurrentDate     =   41138
      End
      Begin MSComCtl2.DTPicker DtpAtencion 
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   477822977
         CurrentDate     =   41138
      End
      Begin MSComCtl2.DTPicker DtpfechaReversion 
         Height          =   315
         Left            =   5295
         TabIndex        =   49
         Top             =   960
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
         Format          =   477822977
         CurrentDate     =   41138
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA REVERSION:"
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
         Index           =   19
         Left            =   3960
         TabIndex        =   50
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label lblnombre 
         BackColor       =   &H000080FF&
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
         Height          =   260
         Index           =   0
         Left            =   5280
         TabIndex        =   47
         Top             =   600
         Width           =   3030
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRABAJADOR :"
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
         Index           =   18
         Left            =   4200
         TabIndex        =   46
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lbldni 
         BackColor       =   &H000080FF&
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
         Height          =   260
         Index           =   0
         Left            =   5280
         TabIndex        =   45
         Top             =   240
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DNI TRABAJADOR :"
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
         Index           =   17
         Left            =   3960
         TabIndex        =   44
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "N? SOLICITUD:"
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
         Index           =   0
         Left            =   480
         TabIndex        =   35
         Top             =   260
         Width           =   1080
      End
      Begin VB.Label lblsolicitud 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2040
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RESUMEN :"
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
         Index           =   12
         Left            =   840
         TabIndex        =   31
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO ENTREGADO :"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO SOLICITUD :"
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
         Index           =   2
         Left            =   225
         TabIndex        =   6
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA ATENCION :"
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
         Left            =   345
         TabIndex        =   5
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA EMISION :"
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
         Left            =   450
         TabIndex        =   4
         Top             =   600
         Width           =   1125
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfGastos 
      Height          =   4125
      Left            =   360
      TabIndex        =   16
      Top             =   4080
      Width           =   16695
      _ExtentX        =   29448
      _ExtentY        =   7276
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
   Begin MSMask.MaskEdBox TxtFecha 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "dd/mm/yyyy"
      Top             =   3360
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
   Begin VitekeySoft.ChameleonBtn cmdeliminar 
      Height          =   855
      Left            =   17160
      TabIndex        =   28
      Top             =   4965
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "FrmSolicitudViaticosDeclarar.frx":0000
      PICN            =   "FrmSolicitudViaticosDeclarar.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdfinalizar 
      Height          =   825
      Left            =   17160
      TabIndex        =   29
      Top             =   4095
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1455
      BTYPE           =   5
      TX              =   "FINALIZAR"
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
      MICON           =   "FrmSolicitudViaticosDeclarar.frx":2466
      PICN            =   "FrmSolicitudViaticosDeclarar.frx":2482
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
      Left            =   17160
      TabIndex        =   30
      Top             =   5835
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
      MICON           =   "FrmSolicitudViaticosDeclarar.frx":4D6C
      PICN            =   "FrmSolicitudViaticosDeclarar.frx":4D88
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcPeriodo 
      Height          =   330
      Left            =   1560
      TabIndex        =   37
      Top             =   3360
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSDataListLib.DataCombo DtcIgv 
      Height          =   330
      Left            =   9120
      TabIndex        =   40
      Top             =   3360
      Width           =   735
      _ExtentX        =   1296
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
   Begin MSDataListLib.DataCombo DtcMoneda 
      Height          =   330
      Left            =   7920
      TabIndex        =   42
      Top             =   3360
      Width           =   1160
      _ExtentX        =   2037
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
   Begin VitekeySoft.ChameleonBtn cmdAgregar 
      Height          =   345
      Left            =   18000
      TabIndex        =   48
      Top             =   3300
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
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
      BCOL            =   32768
      BCOLO           =   32768
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSolicitudViaticosDeclarar.frx":5178
      PICN            =   "FrmSolicitudViaticosDeclarar.frx":5194
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "MONEDA"
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
      Index           =   16
      Left            =   7920
      TabIndex        =   43
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "IGV"
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
      Index           =   15
      Left            =   9120
      TabIndex        =   41
      Top             =   3120
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "DESCRIPCION"
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
      Index           =   14
      Left            =   4680
      TabIndex        =   39
      Top             =   3120
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "PERIODO"
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
      Index           =   13
      Left            =   1440
      TabIndex        =   36
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "LUGAR"
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
      Index           =   11
      Left            =   10920
      TabIndex        =   27
      Top             =   3120
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Index           =   10
      Left            =   14970
      TabIndex        =   24
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Index           =   9
      Left            =   14280
      TabIndex        =   23
      Top             =   3120
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "RUC"
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
      Index           =   8
      Left            =   16125
      TabIndex        =   22
      Top             =   3120
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Height          =   195
      Index           =   7
      Left            =   12600
      TabIndex        =   21
      Top             =   3120
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "MONTO"
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
      Index           =   6
      Left            =   9900
      TabIndex        =   20
      Top             =   3120
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "GASTO"
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
      Left            =   3600
      TabIndex        =   19
      Top             =   3120
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   18
      Top             =   3120
      Width           =   435
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   5295
      Left            =   240
      Top             =   3000
      Width           =   18375
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   8370
      Left            =   0
      Top             =   0
      Width           =   18690
   End
End
Attribute VB_Name = "FrmSolicitudViaticosDeclarar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_NEW
      Procedencia = nuevo
      FrmSolicitudViaticosDet.Show
     Case KEY_UPDATE
        
        strCadena = "SELECT * FROM movimiento_caja WHERE "
        
         strCadena = "SELECT     Comprobantes.doc_abrev, movimiento_caja.serie, movimiento_caja.numero, movimiento_caja.cPersona, " & _
        "movimiento_caja.descripcion_per, Persona.sDireccionCliente1, Persona.Per_Ruc, movimiento_caja.fecha_valor," & _
        "movimiento_caja.cambio , movimiento_caja.glosa, centro_costos.descripcion, movimiento_caja.Monto,movimiento_caja.monto_letras " & _
        "FROM movimiento_caja INNER JOIN Comprobantes ON movimiento_caja.doc_cod = Comprobantes.doc_cod INNER JOIN " & _
        "centro_costos ON movimiento_caja.id_costo = centro_costos.id_costo INNER JOIN " & _
        "Persona ON movimiento_caja.cPersona = Persona.cPersona WHERE codigo='" & idRecibo & "'"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptReciboCaja", , App.Path + "\Reportes\")
    Case "(Declarar)"
        FrmSolicitudViaticosDeclarar.Show
   Case KEY_DELETE
      If MsgBox(MSGELIMINAR + Chr(13) + "Se Eliminaran los cheques Relacionados a esta Chequera", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        
    End If
    Case KEY_EXIT
        Unload Me
  End Select
End Sub
Private Sub llenar_detalle(ByVal Grilla As MSHFlexGrid, ByVal id_solicitud As Double)
On Error GoTo salir
Dim tTotal As Double
tTotal = 0
strCadena = "SELECT * FROM solicitud_dinero_detalle WHERE id_solicitud='" & id_solicitud & "' AND ruc='" & KEY_RUC & "'"
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
           Grilla.ColWidth(2) = 5000
           Grilla.ColWidth(3) = 1400
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
        Fila = "" & vbTab & "" & vbTab & "***********  TOTAL SOLICITADO ***********" & vbTab & Format(tTotal, "###0.00")
        Grilla.AddItem Fila
       
      For k = 0 To 3
            Grilla.col = k
            Grilla.Row = i
            Grilla.CellBackColor = &HC0C0FF
      Next k

Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub cmdagregar_Click()
If Trim(Me.TxtRuc.Text) <> "" And Trim(Me.TxtRazonsocial.Text) <> "" Then
     Call put_gasto
End If
End Sub
Private Sub put_gasto()
Dim cod_identidad As String * 1
Dim valor_venta As Double
Dim igv As Double
Dim Total As Double


strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Me.DtcComprobante.BoundText & "' AND serie='" & Me.TxtSerie.Text & "' AND numero='" & Me.TxtNumero.Text & "' AND id_proveedor='" & Me.TxtRuc.Text & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
If Len(Trim(Me.TxtRuc.Text)) = 8 Then
    cod_identidad = 1
End If
If Len(Trim(Me.TxtRuc.Text)) = 11 Then
    cod_identidad = 6
End If

If Len(Trim(Me.TxtRuc.Text)) <> 8 And Len(Trim(Me.TxtRuc.Text)) <> 11 Then
    cod_identidad = 0
End If


If Me.DtcIgv.BoundText = "SI" Then
    exonerado = 0
    valor_venta = Val(Me.TxtMonto.Text) / (KEY_IGV + 1)
    igv = Val(Me.TxtMonto.Text) - valor_venta
Else
    
    igv = 0
    exonerado = Val(Me.TxtMonto.Text)
    valor_venta = 0
    
End If


        
        If Me.DtcMoneda.BoundText = "00001" Then
                in_cta_compra = KEY_CTA_COMPRA_SOLES
           Else
                in_cta_compra = KEY_CTA_COMPRA_SOLES
           End If
           
            If Me.DtcComprobante.BoundText = "0002" Then
             in_cta_compra = KEY_CTA_COMPRA_RH
        End If
        
        
          in_cta_compra = KEY_CTA_PAGAR_SERVICIO
        
        If KEY_CONTABILIDAD = "si" Then
           If put_verifica_cuenta_contable(Me.DtcComprobante.BoundText, Trim(Me.TxtSerie.Text), Trim(Me.TxtNumero.Text), in_cta_compra, "02") = False Then
              Exit Sub
           End If
        End If
       
      
       
        
        
       
        
        If KEY_PAIS = KEY_PERU Then
            strCadena = "call P_insert_compra_ultimate('" & Me.DtcComprobante.BoundText & "','" & KEY_ALM & "','" & Format(Me.TxtFecha.Text, "YYYY-mm-dd") & "','" & Format(TxtFecha.Text, "YYYY-mm-dd") & "','02'," & _
            "'02','--','" & Me.DtcMoneda.BoundText & "','" & formato_item(Month(Me.TxtFecha.Text), 2) & "','" & Year(Me.TxtFecha.Text) & "','" & Trim(Me.TxtSerie.Text) & "'," & _
            "'" & Format(Trim(Me.TxtNumero.Text), "00000000") & "','" & cod_identidad & "','" & Trim(Me.TxtRuc.Text) & "','" & UCase(Me.TxtRazonsocial.Text) & "','" & KEY_CAMBIO & "'," & _
            "'0','" & valor_venta & "','" & igv & "','0','0','0','" & Val(in_retencion) & "','" & exonerado & "','0','" & Val(Me.TxtMonto.Text) & "','" & Val(Me.TxtMonto.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.TxtDetalle.Text) & "','02','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & Trim(Me.lbldni(0).Caption) & "','0','0','0','0','" & KEY_RUC & "')"
        Else
        
            strCadena = "call P_insert_compra_ultimate_internacional('" & Me.DtcComprobante.BoundText & "','" & KEY_ALM & "','" & Format(Me.TxtFecha.Text, "YYYY-mm-dd") & "','" & Format(DtpVencimiento.Value, "YYYY-mm-dd") & "','02'," & _
            "'02','--','" & Me.DtcMoneda.BoundText & "','" & formato_item(Month(Me.TxtFecha.Text), 2) & "','" & Year(Me.TxtFecha.Text) & "','" & Trim(Me.TxtSerie.Text) & "'," & _
            "'" & Format(Trim(Me.TxtNumero.Text), "00000000") & "','" & cod_identidad & "','" & Trim(Me.lbldni(0).Caption) & "','" & UCase(Me.lblnombre(0).Caption) & "','" & KEY_CAMBIO & "'," & _
            "'0','" & valor_venta & "','" & igv & "','0','0','0','" & in_retencion & "','" & exonerado & "','0','" & Val(Me.TxtMonto.Text) & "','" & Val(Me.TxtMonto.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.TxtDetalle.Text) & "','02','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & Trim(Me.lbldni(0).Caption) & "','0','0','0','0','" & KEY_RUC & "')"
        End If
        Call ConfiguraRstP(strCadena)
        id_compra = rstP(0)
        
       
        
        
        
        If Me.DtcIgv.BoundText = "SI" Then
            in_afecto = "si"
            in_exonerado = 0
            valor_venta = Val(Me.TxtMonto.Text) / (KEY_IGV + 1)
            igv = Val(Me.TxtMonto.Text) - valor_venta
        Else
            in_afecto = "no"
            igv = 0
            in_exonerado = Val(Me.TxtMonto.Text)
            valor_venta = 0
    
        End If
        
        
       
        
        
        strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,detalle,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,ivap,otros,percepcion, " & _
        "valor_venta,exonerado,total,p_venta,p_costo,id_alm,retencion,ruc) VALUES ('" & id_compra & "','" & Trim(Me.txtid_servicio.Text) & "','" & Trim(Me.TxtDetalle.Text) & "','1','" & Val(Me.TxtMonto.Text) & "'," & _
        "'0','0','0','" & valor_venta & "','0','" & igv & "', " & _
        "'0','0','0','" & valor_venta & "','" & in_exonerado & "','" & Val(Me.TxtMonto.Text) & "','" & Val(Me.TxtMonto.Text) & "','" & get_precio_costo(Trim(Me.txtid_servicio.Text)) & "','" & KEY_ALM & "','" & Val(in_retencion) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
     


        
'02----------------guardar en detalle documento Compra-----------
 If KEY_CONTABILIDAD = "si" And Me.DtcComprobante.BoundText <> "0089" Then
    If KEY_PAIS = KEY_PERU Then
        strCadena = "call p_insert_compra_emitido_ii('" & id_compra & "')"
    Else
         strCadena = "call p_insert_compra_emitido_internacional('" & id_compra & "')"
    End If
    Call Execute_Sql(strCadena)
 End If
 
        
     strCadena = "INSERT INTO solicitud_dinero_declarar(id_solicitud,id_moneda,afecto_igv,id_compra,valor_venta,igv,total,descripcion,fecha_gasto,monto,id_doc,serie,numero,id_proveedor,lugar_gasto,dni_save,ruc)VALUES " & _
    "('" & Val(Me.TxtId_solicitud.Text) & "','" & Me.DtcMoneda.BoundText & "','" & Me.DtcIgv.BoundText & "','" & id_compra & "','" & valor_venta & "','" & igv & "','" & Val(Me.TxtMonto.Text) & "','" & UCase(Me.TxtDetalle.Text) & "','" & Format(Me.TxtFecha.Text, "YYYY-mm-dd") & "','" & Val(Me.TxtMonto.Text) & "'," & _
    "'" & Me.DtcComprobante.BoundText & "','" & Me.TxtSerie.Text & "','" & Me.TxtNumero.Text & "','" & Me.TxtRuc.Text & "','" & Me.TxtLugar.Text & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    
    
        
        


 in_compra = get_periodo_detalle(Me.DtcPeriodo.BoundText, id_compra)
 MsgBox "NUMERO DE VOUCHER GENERADO  : " & Trim(in_compra), vbInformation, KEY_VENDEDOR
  Call llenar_gastos(Me.HfGastos, Val(Me.TxtId_solicitud.Text))
    
    Me.TxtDetalle.Text = ""
    Me.TxtMonto.Text = ""
    Me.TxtLugar.Text = ""
    Me.DtcComprobante.Text = ""
    Me.TxtSerie.Text = ""
    Me.TxtNumero.Text = ""
    Me.TxtRuc.Text = ""
    Me.TxtRazonsocial.Text = ""
    Me.TxtFecha.SetFocus



Exit Sub
Else
    MsgBox "COMPROBANTE YA REGISTRADO, IMPOSIBLE GUARDAR ", vbInformation, KEY_EMPRESA
End If
Set rst = Nothing

End Sub

Public Sub llenar_gastos(ByVal Grilla As MSHFlexGrid, ByVal id_solicitud As Double)
On Error GoTo salir
Dim tTotal As Double, Treembolso As Double, comprobante As String, razonsocial As String
tTotal = 0
Treembolso = 0
strCadena = "SELECT * FROM view_viatico_detaclarar WHERE id_solicitud='" & id_solicitud & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
   Grilla.Rows = 0
   Exit Sub
End If

   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 700
           Grilla.ColWidth(2) = 1100
           Grilla.ColWidth(3) = 2000
           Grilla.ColWidth(4) = 3000
           Grilla.ColWidth(5) = 3500
           Grilla.ColWidth(6) = 1500
           Grilla.ColWidth(7) = 1500
           Grilla.ColWidth(8) = 1000
           Grilla.ColWidth(9) = 1000
           Grilla.ColWidth(10) = 1000
       Next
         cabecera = "IDDETALLE" & vbTab & "COMPRA" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "RAZON SOCIAL" & vbTab & "DETALLE GASTO" & vbTab & "LUGAR" & vbTab & "MONEDA" & vbTab & "V.VENTA" & vbTab & "IGV" & vbTab & "TOTAL"
         Grilla.AddItem cabecera
         For k = 0 To 10
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        in_valor_venta = 0
        in_igv = 0
        in_total = 0
        For i = 0 To rst.RecordCount - 1
        
        in_valor_venta = in_valor_venta + rst("valor_venta")
        in_igv = in_igv + rst("igv")
        in_total = in_total + rst("total")
        If rst("id_moneda") = "00001" Then
            in_moneda = "SOLES"
        Else
            in_moneda = "DOLARES"
        End If
            
             Fila = rst("id_detalle") & vbTab & rst("id_compra") & vbTab & rst("fecha_gasto") & vbTab & rst("comprobante") & vbTab & rst("nombre_completo") & vbTab & rst("descripcion") & vbTab & rst("lugar_gasto") & vbTab & in_moneda & vbTab & Format(rst("valor_venta"), "#,##0.00") & vbTab & Format(rst("igv"), "#,##0.00") & vbTab & Format(rst("total"), "#,##0.00")
             Grilla.AddItem Fila
             
        rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "T. DECLARADO :" & vbTab & Format(in_valor_venta, "###0.00") & vbTab & Format(in_igv, "###0.00") & vbTab & Format(in_total, "###0.00")
        Grilla.AddItem Fila
       
      For k = 7 To 10
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0C0FF
      Next k
        
      
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub cmdEliminar_Click()
If MsgBox("ESTA SEGURO DE ELIMINAR ESTE REGISTRO ", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
    Call disabled_form(Me)
    Procedencia = Eliminar
    frmsegurity.Show
    Exit Sub
End If
End Sub

Private Sub cmdfinalizar_Click()
If MsgBox("ESTA SEGURO DE FINALIZAR SU REPORTE DE SOLICITUD", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
            strCadena = "UPDATE solicitud_dinero SET finalizado='si',fecha_reversion='" & Format(Me.DtpfechaReversion.Value, "YYYY-mm-dd") & "' WHERE id_solicitud='" & Val(Me.TxtId_solicitud.Text) & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
            strCadena = "call CON_InsertaAsiento_CanjeViaticos('" & Val(Me.TxtId_solicitud.Text) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            Call FrmSolicitudViaticos.actualizar
            Unload Me
        End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub DtcComprobante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtSerie)
End If
End Sub

Private Sub DtcIgv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtMonto)
End If
End Sub

Private Sub DtcMoneda_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Me.DtcIgv.SetFocus
End If

End Sub

Private Sub DtcPeriodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtid_servicio)
    
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
Me.DtpfechaReversion.Value = KEY_FECHA

Dim id_solicitud As Double

Me.TxtFecha.Text = Format(KEY_FECHA, "dd/mm/YYYY")
Me.TxtId_solicitud.Text = FrmSolicitudViaticos.HfgDetalle.TextMatrix(FrmSolicitudViaticos.HfgDetalle.Row, 0)
id_solicitud = Val(Me.TxtId_solicitud.Text)
strCadena = "SELECT * FROM solicitud_dinero S WHERE id_solicitud='" & id_solicitud & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.lblsolicitud.Caption = rst("numero")
    Me.DtpEmision.Value = rst("fecha_solicitud")
    If IsNull(rst("fecha_reversion")) = True Then
        Me.DtpfechaReversion.Value = KEY_FECHA
    Else
        Me.DtpfechaReversion.Value = rst("fecha_reversion")
    End If
    If IsNull(rst("fecha_confirmacion")) = False Then
        Me.DtpAtencion.Value = rst("fecha_confirmacion")
    Else
        Me.DtpAtencion.Value = KEY_FECHA
    End If
    
    Me.TxtMontoSolicitud.Text = Format(rst("monto_solicitado"), "###0.00")
    Me.TxtMontoEntregado.Text = Format(rst("monto_entregado"), "###0.00")
    Me.txtResumen.Text = rst("resumen")
    
    Me.lbldni(0).Caption = rst("dni")
    Me.lblnombre(0).Caption = get_persona(rst("dni"))
    
    Call llenar_detalle(Me.HfdDetalle, id_solicitud)
    Call llenar_gastos(Me.HfGastos, id_solicitud)
End If
strCadena = "SELECT id_doc as Codigo,doc_des as Descripcion FROM comprobantes ORDER BY doc_des"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComprobante)
Me.DtcComprobante.BoundText = "0003"



strCadena = "SELECT igv as Codigo,igv as Descripcion FROM afecto_igv ORDER BY igv"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcIgv)



strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda ORDER BY id_moneda"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)
Me.DtcMoneda.BoundText = KEY_MONEDA



  strCadena = "SELECT Id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion  FROM con_periodo order by codigo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodo)
  Me.DtcPeriodo.BoundText = get_periodo_actual(KEY_FECHA)
  


End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case "(Finalizar)"
        
    Case KEY_DELETE
         
    Case KEY_EXIT
        Unload Me
End Select
End Sub

Private Sub HfGastos_SelChange()
If Val(Me.HfGastos.TextMatrix(Me.HfGastos.Row, 0)) > 0 Then
   Me.cmdeliminar.Enabled = True
Else
   Me.cmdeliminar.Enabled = False
End If
End Sub

Private Sub TxtDetalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcMoneda.SetFocus
End If
End Sub

Private Sub TxtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcPeriodo.SetFocus
End If
End Sub

Private Sub txtid_servicio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
   strCadena = "SELECT nombre_prod FROM producto WHERE id_producto='" & Format(Trim(Me.txtid_servicio.Text), "00000") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      Me.txtid_servicio.Text = Format(Me.txtid_servicio.Text, "00000")
      Me.TxtDetalle.Text = rst("nombre_prod")
      Me.DtcMoneda.SetFocus
  Else
   
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
End If
End If
End Sub

Private Sub TxtLugar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtLugar.Text = UCase(Me.TxtLugar.Text)
    Me.DtcComprobante.SetFocus
End If
End Sub

Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtMonto.Text = Format(Val(Me.TxtMonto.Text), "###0.00")
    Call Resalta(Me.TxtLugar)
End If
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtNumero.Text = formato_item(Me.TxtNumero.Text, 6)
    Call Resalta(Me.TxtRuc)
End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
        Call buscarcliente(Trim(Me.TxtRuc.Text))
   
    
End If
End Sub
Private Sub buscarcliente(ByVal ruc As String)
If ruc = "" Then
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
End If
    strCadena = "SELECT *  FROM persona WHERE dni='" & ruc & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.TxtRuc.Text = Trim(Me.TxtRuc.Text)
        FrmDetallePersona.chkProveedor.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.TxtRuc.Text = rst("dni")
        Me.TxtRazonsocial.Text = rst("nombre_completo")
        Me.cmdAgregar.SetFocus
        Exit Sub
       
    End If

End Sub
Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie.Text = formato_item(Me.TxtSerie.Text, 3)
    Call Resalta(Me.TxtNumero)
End If
End Sub

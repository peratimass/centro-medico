VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmRegistroVentasList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmCerrarPeriodo 
      BackColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   3960
      TabIndex        =   41
      Top             =   3840
      Visible         =   0   'False
      Width           =   5175
      Begin MSDataListLib.DataCombo DtcPeriodo 
         Height          =   360
         Left            =   1200
         TabIndex        =   42
         Top             =   840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   4194304
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
         Height          =   705
         Left            =   1200
         TabIndex        =   45
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1244
         BTYPE           =   3
         TX              =   "CERRAR PERIODO"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight Condensed"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmRegistroVentasList.frx":0000
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CERRAR PERIODO"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiBold SemiConden"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1200
         TabIndex        =   44
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "PERIODO :"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   240
         TabIndex        =   43
         Top             =   840
         Width           =   750
      End
   End
   Begin VB.TextBox txtCliente 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      TabIndex        =   38
      Top             =   450
      Width           =   1695
   End
   Begin VB.Frame frmDetalle 
      BackColor       =   &H00FFFFFF&
      Height          =   6615
      Left            =   9960
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   8895
      Begin MSComCtl2.DTPicker DtpEmision 
         Height          =   300
         Left            =   1680
         TabIndex        =   51
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   163708929
         CurrentDate     =   44642
      End
      Begin VB.TextBox txtObservaciones 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   5325
         Width           =   5055
      End
      Begin VB.TextBox txtObservacion 
         Appearance      =   0  'Flat
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
         Height          =   500
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   4245
         Width           =   4935
      End
      Begin VB.TextBox txtMontoCobrado 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1680
         TabIndex        =   32
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CheckBox chk_verificado 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "VERIFICADO"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiBold SemiConden"
            Size            =   14.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   3120
         TabIndex        =   29
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox txtComision 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtOperacionVenta 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1680
         TabIndex        =   26
         Top             =   3840
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DtcVendedorComprobante 
         Height          =   330
         Left            =   1680
         TabIndex        =   24
         Top             =   975
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcCuentaBancaria 
         Height          =   330
         Left            =   1680
         TabIndex        =   25
         Top             =   1440
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   465
         Left            =   1680
         TabIndex        =   27
         Top             =   4800
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "PROCESAR"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmRegistroVentasList.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdSiguiente 
         Height          =   465
         Left            =   4680
         TabIndex        =   30
         Top             =   4800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   820
         BTYPE           =   3
         TX              =   "SIGUIENTE"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmRegistroVentasList.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcFlujo 
         Height          =   315
         Left            =   3120
         TabIndex        =   36
         Top             =   3360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMISION :"
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
         Left            =   600
         TabIndex        =   50
         Top             =   680
         Width           =   615
      End
      Begin VB.Label lbl_saldo 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   3120
         TabIndex        =   49
         Top             =   2880
         Width           =   1290
      End
      Begin VB.Label lbl_pagado 
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   3120
         TabIndex        =   48
         Top             =   2520
         Width           =   1290
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO SALDO    :"
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
         Left            =   1800
         TabIndex        =   47
         Top             =   2880
         Width           =   1140
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO PAGADO  :"
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
         Left            =   1800
         TabIndex        =   46
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   1680
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Width           =   4815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMISION BANCO :"
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
         Left            =   135
         TabIndex        =   37
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACION :"
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
         TabIndex        =   35
         Top             =   4440
         Width           =   990
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO FACTURADO :"
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
         Left            =   150
         TabIndex        =   31
         Top             =   1920
         Width           =   1350
      End
      Begin VB.Image cmdclose 
         Height          =   240
         Left            =   8400
         Picture         =   "FrmRegistroVentasList.frx":0054
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENDEDOR  :"
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
         Left            =   555
         TabIndex        =   23
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° OPERACION  :"
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
         Left            =   300
         TabIndex        =   22
         Top             =   3960
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUENTA BANCARIA :"
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
         Left            =   45
         TabIndex        =   21
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Label lblcomprobante 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMPROBANTE :"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiBold"
            Size            =   11.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   1710
      End
   End
   Begin MSDataListLib.DataCombo DtcVendedor 
      Height          =   315
      Left            =   9840
      TabIndex        =   17
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtOperacionBusqueda 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7680
      TabIndex        =   15
      Top             =   120
      Width           =   1215
   End
   Begin VitekeySoft.ChameleonBtn cmdBuscarfecha 
      Height          =   375
      Left            =   15000
      TabIndex        =   12
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRegistroVentasList.frx":2EF8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   345
      Left            =   13665
      TabIndex        =   10
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
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
      Format          =   163708929
      CurrentDate     =   44576
   End
   Begin VB.TextBox txtNumero 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7800
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   18855
      _ExtentX        =   33258
      _ExtentY        =   14631
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   1
      Cols            =   11
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin VitekeySoft.ChameleonBtn cmdSalir 
      Height          =   780
      Left            =   19080
      TabIndex        =   2
      Top             =   7260
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1376
      BTYPE           =   5
      TX              =   "SALIR"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRegistroVentasList.frx":2F14
      PICN            =   "FrmRegistroVentasList.frx":2F30
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   780
      Left            =   19080
      TabIndex        =   3
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1376
      BTYPE           =   5
      TX              =   "NUEVO"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRegistroVentasList.frx":3320
      PICN            =   "FrmRegistroVentasList.frx":333C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdEliminar 
      Height          =   780
      Left            =   19080
      TabIndex        =   4
      Top             =   2400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1376
      BTYPE           =   5
      TX              =   "ELIMINAR"
      ENAB            =   0   'False
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRegistroVentasList.frx":378E
      PICN            =   "FrmRegistroVentasList.frx":37AA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdVerificar 
      Height          =   780
      Left            =   19080
      TabIndex        =   5
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1376
      BTYPE           =   5
      TX              =   "VERIFICAR"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRegistroVentasList.frx":5BF4
      PICN            =   "FrmRegistroVentasList.frx":5C10
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdImportar 
      Height          =   780
      Left            =   19080
      TabIndex        =   6
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1376
      BTYPE           =   5
      TX              =   "IMPORTAR"
      ENAB            =   0   'False
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRegistroVentasList.frx":9951
      PICN            =   "FrmRegistroVentasList.frx":996D
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
      Height          =   780
      Left            =   19080
      TabIndex        =   7
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1376
      BTYPE           =   5
      TX              =   "ANULAR"
      ENAB            =   0   'False
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRegistroVentasList.frx":BFA6
      PICN            =   "FrmRegistroVentasList.frx":BFC2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdPendientePago 
      Height          =   375
      Left            =   18120
      TabIndex        =   14
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRegistroVentasList.frx":E40C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrarPeriodo 
      Height          =   1020
      Left            =   19080
      TabIndex        =   40
      Top             =   6240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1799
      BTYPE           =   5
      TX              =   "CERRAR PERIODO"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRegistroVentasList.frx":E428
      PICN            =   "FrmRegistroVentasList.frx":E444
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
      Caption         =   "NOMBRE CLIENTE :"
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
      Left            =   3000
      TabIndex        =   39
      Top             =   480
      Width           =   1185
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENDEDOR:"
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
      Left            =   9120
      TabIndex        =   18
      Top             =   180
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N° OPERACION :"
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
      Left            =   6600
      TabIndex        =   16
      Top             =   180
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "PENDIENTE DE VERIFICACION"
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
      Height          =   285
      Left            =   15960
      TabIndex        =   13
      Top             =   165
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA:"
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
      Left            =   13080
      TabIndex        =   11
      Top             =   180
      Width           =   450
   End
   Begin VB.Label Label1 
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
      Left            =   3120
      TabIndex        =   8
      Top             =   180
      Width           =   1050
   End
   Begin VB.Label lblMes 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Ventas Mensual:"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiCondensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   195
      Width           =   2835
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   750
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Width           =   18855
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmRegistroVentasList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim igv As String, id_alm As String
Public Procedencia As EnumProcede
Dim correlativo As Double



Private Sub cmdagregar_Click()

End Sub
Private Sub formatearGrilla(ByVal Grilla As MSHFlexGrid)
         Grilla.Clear

   Grilla.Rows = 0

       ReDim arrColWidth(1 To rst.Fields.Count)
       For i = 0 To 0

           Grilla.ColWidth(0) = 600
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 600
           Grilla.ColWidth(3) = 600
           Grilla.ColWidth(4) = 1000
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 3800
           Grilla.ColWidth(7) = 1200
           Grilla.ColWidth(8) = 1100
           Grilla.ColWidth(9) = 1200
           Grilla.ColWidth(10) = 1450
           Grilla.ColWidth(11) = 1000
           Grilla.ColWidth(12) = 0
           
        Next i
         cabecera = "ITEM" & vbTab & "FECHA" & vbTab & "TD" & vbTab & "SERIE" & vbTab & "NUMERO" & vbTab & "RUC" & vbTab & "CLIENTE" & vbTab & "AFECTO" & vbTab & "EXONERADO" & vbTab & "IGV" & vbTab & "TOTAL" & vbTab & "RETENCION" & vbTab & "CODIGO"
        Grilla.AddItem cabecera
         For k = 0 To 12
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
End Sub
Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
'On Error GoTo SALIR
Dim tafecto As Double, texonerado As Double, tigv As Double, tTotal As Double, tRetencion As Double
tafecto = 0
texonerado = 0
tigv = 0
tTotal = 0
tRetencion = 0
correlativo = 0
in_comision = 0


Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
 N = 1
   
   Grilla.Rows = 0

       ReDim arrColWidth(1 To rst.Fields.Count)
       

           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 600
           Grilla.ColWidth(2) = 1100
           Grilla.ColWidth(3) = 500
           Grilla.ColWidth(4) = 500
           Grilla.ColWidth(5) = 800
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 2500
           Grilla.ColWidth(8) = 1000
           Grilla.ColWidth(9) = 1000
           Grilla.ColWidth(10) = 1000
           Grilla.ColWidth(11) = 1000
           Grilla.ColWidth(12) = 2100
           Grilla.ColWidth(13) = 2300
           Grilla.ColWidth(14) = 800
           Grilla.ColWidth(15) = 1000
           Grilla.ColWidth(16) = 1000
       
        cabecera = "IDVENTA" & vbTab & "ITEM" & vbTab & "FECHA" & vbTab & "TD" & vbTab & "SERIE" & vbTab & "NUMERO" & vbTab & "DNI-RUC" & vbTab & "CLIENTE" & vbTab & "AFECTO" & vbTab & "EXONERADO" & vbTab & "IGV" & vbTab & "TOTAL" & vbTab & "VENDEDOR" & vbTab & "CUENTA BANCARIA" & vbTab & "COMISION" & vbTab & "OPERACION" & vbTab & "ESTADO"
        Grilla.AddItem cabecera
         For k = 0 To 16
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
        Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            
                Fila = str(rst("id_venta")) & vbTab & Format(str(i + 1), "000") & vbTab & rst("fecha_emision") & vbTab & rst("id_doc") & vbTab & rst("serie") & vbTab & rst("numero") & vbTab & rst("id_cliente") & vbTab & rst("ncliente") & vbTab & Format(rst("valor_venta"), "#,##0.00") & vbTab & Format(rst("exonerado"), "#,##0.00") & vbTab & Format(rst("igv"), "#,##0.00") & vbTab & Format(rst("total"), "#,##0.00") & vbTab & rst("vendedor") & vbTab & rst("cuenta_bancaria") & vbTab & Format(rst("tc_local"), "#,##0.00") & vbTab & rst("operacion") & vbTab & rst("estado")
          
            
            Grilla.AddItem Fila
            correlativo = correlativo + 1
            If (Trim(rst("anulado")) = "si") Then
                            For k = 0 To 11
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                            Next k
            Else
                            If ("id_doc") = "0007" Then
                                tafecto = tafecto + rst("valor_venta") * -1
                                texonerado = texonerado + rst("exonerado") * -1
                                tigv = tigv + rst("igv") * -1
                                tTotal = tTotal + rst("total") * -1
                                tRetencion = tRetencion + rst("retencion") * -1
                            Else
                                tafecto = tafecto + rst("valor_venta")
                                texonerado = texonerado + rst("exonerado")
                                tigv = tigv + rst("igv")
                                tTotal = tTotal + rst("total")
                                tRetencion = tRetencion + rst("retencion")
                            End If
            End If
            in_comision = in_comision + rst("tc_local")
            Grilla.col = 16
            Grilla.Row = i + 1
            If rst("pendiente") = "si" Then
                Grilla.TextMatrix(i + 1, 16) = "PENDIENTE"
                Grilla.CellBackColor = &H8080FF
            Else
                Grilla.TextMatrix(i + 1, 16) = "VERIFICADO"
                Grilla.CellBackColor = &H80FF80
            End If
                        
                   
            
            
            
            rst.MoveNext
             
        Next i
       Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(tafecto, "#,##0.00") & vbTab & Format(texonerado, "#,##0.00") & vbTab & Format(tigv, "#,##0.00") & vbTab & Format(tTotal, "#,##0.00") & vbTab & "" & vbTab & "" & vbTab & Format(in_comision, "#,##0.00")
       Grilla.AddItem Fila
       For k = 8 To 14
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H80FF&
        Next k
                            
   
  Exit Sub
'SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub









Private Sub CmdBuscarFecha_Click()
strCadena = "CALL ADM_reportes_generales('34','" & Format(Me.DtpFecha.Tag, "YYYY-mm-dd") & "','" & Format(Me.DtpFecha.Value, "YYYY-mm-dd") & "','','','','','','','','" & KEY_RUC & "')"
Call llenarGrid(Me.HfdPersona)

End Sub

Private Sub cmdClose_Click()

Me.frmdetalle.Visible = False
Me.HfdPersona.Enabled = True
End Sub



Private Sub cmdPendientePago_Click()
    strCadena = "CALL ADM_reportes_generales('36','" & Format(Me.DtpFecha.Tag, "YYYY-mm-dd") & "','','','','','','','','','" & KEY_RUC & "')"
    Call llenarGrid(Me.HfdPersona)
End Sub

Private Sub cmdProcesar_Click()
Call put_comprobante(Me.lblcomprobante.Tag)
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdsiguiente_Click()

in_fila = Val(Me.CmdSiguiente.Tag) + 1
If in_fila < Me.HfdPersona.Rows Then
    
    in_venta = Me.HfdPersona.TextMatrix(in_fila, 0)
     Call get_comprobante(in_venta)
    Me.CmdSiguiente.Tag = in_fila
End If



End Sub

Private Sub cmdVerificar_Click()
 
      Call get_comprobante(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
End Sub





Public Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub

Private Sub DtcVendedor_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    strCadena = "CALL ADM_reportes_generales('37','" & Format(Me.DtpFecha.Tag, "YYYY-mm-dd") & "','','" & Me.DtcVendedor.BoundText & "','','','','','','" & Trim(Me.TxtCliente.Text) & "','" & KEY_RUC & "')"
    Call llenarGrid(Me.HfdPersona)
End If

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
correlativo = 0
Me.DtpFecha.Value = KEY_FECHA




Me.lblMes.Caption = FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 2) + Space(2) + FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3)
in_anio = FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3)
  
in_fecha = Format("01-" + Format(FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 1), "00") + "-" + in_anio, "dd-mm-YYYY")
  

strCadena = "SELECT id as Codigo,Nombre  as Descripcion FROM adm_flujocaja ORDER BY Nombre  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFlujo)


strCadena = "CALL ADM_servicios_generales('16','dni','','','','','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcVendedor)

strCadena = "CALL ADM_servicios_generales('16','dni','','','','','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcVendedorComprobante)

strCadena = "CALL ADM_servicios_generales('17','dni','','','','','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcCuentaBancaria)
Me.DtpFecha.Tag = Format(in_fecha, "YYYY-mm-dd")


strCadena = "CALL ADM_reportes_generales('33','" & Format(in_fecha, "YYYY-mm-dd") & "','','','','','','','','','" & KEY_RUC & "')"
Call llenarGrid(Me.HfdPersona)
  
  
  
End Sub
Private Sub put_insertar_abono()
'----
End Sub
Private Function get_forma_pago_cuenta(ByVal in_cuenta As String) As String
strCadena = "SELECT id_registro FROM  view_forma_pago_conta WHERE id_cuenta='" & in_cuenta & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    get_forma_pago_cuenta = rstT("id_registro")
Else
    get_forma_pago_cuenta = "0"
End If

End Function

Private Sub put_comprobante(ByVal in_venta As String)
Dim in_total_comprobante As Double
in_observacion = "[" + KEY_FECHA + Space(2) + str(Time) + "]  " + Mid(KEY_VENDEDOR, 1, 30) + " : " + UCase(Me.txtObservacion.Text) & vbCrLf & Trim(Me.txtObservaciones.Text)

'monto_pago = monto comision
'orden_compra= flujo de pago
in_validador = True

If Val(Me.DtcCuentaBancaria.BoundText) <= 0 Then
   in_validador = False
End If

If Trim(Me.txtOperacionVenta.Text) = "" Then
   in_validador = False
End If

If Trim(Me.txtOperacionVenta.Text) = "-" Then
   in_validador = False
End If


If Me.chk_verificado.Value = 1 Then
   in_pendiente = "no"
   If in_validador = False Then
      in_pendiente = "si"
      MsgBox "PARA MARCAR VERIFICADO DEBE CUMPLIR CON" + Chr(13) + "1. CUENTA BANCARIA." + Chr(13) + "2. NUMERO OPERACION", vbInformation, "RECUERDA: " + KEY_VENDEDOR
      Exit Sub
   End If
Else
   in_pendiente = "si"
End If

strCadena = "UPDATE movimiento_venta SET tc_local='" & Val(Me.txtComision.Text) & "',orden_compra='" & Me.DtcFlujo.BoundText & "',id_orden_salida='" & Me.DtcCuentaBancaria.BoundText & "',operacion='" & Trim(Me.txtOperacionVenta.Text) & "',id_vendedor='" & Me.DtcVendedorComprobante.BoundText & "',pendiente='" & in_pendiente & "',observacion='" & in_observacion & "' WHERE id_venta='" & in_venta & "'"
Call ConfiguraRst(strCadena)

strCadena = "call ADM_impresion_comprobantes('19','" & in_venta & "','" & KEY_RUC & "')"
Call ConfiguraRstK(strCadena)
in_documento = rstK("documento")


strCadena = "call ADM_impresion_comprobantes('18','" & in_venta & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst(0) < rstK("total") And Val(Me.txtMontoCobrado.Text) > 0 Then
    '*********REALIZAR EL ABONO A LA CUENTA
    strCadena = "call ADM_impresion_comprobantes('19','" & in_venta & "','" & KEY_RUC & "')"
    Call ConfiguraRstK(strCadena)
    If rstK.RecordCount > 0 Then
        in_glosa = "COBRO:" + in_documento
        'strCadena = "call sp_procesar_transaccion_caja_venta('" & KEY_ALM & "','" & Me.DtcCuentaBancaria.BoundText & "','" & Format(rstK("fecha_emision"), "YYYY-mm-dd") & "','00001','" & rstK("id_cliente") & "','" & in_glosa & "','" & Format(Me.txtMontoCobrado.Text, "###0.00") & "','','" & rstK("tc") & "','" & rstK("operacion") & "','" & KEY_USUARIO & "','" & rstK("id_venta") & "','0','" & rstK("ncliente") & "','" & rstK("documento") & "','1CIX000000000025','" & rstK("id_forma_pago") & "','" & KEY_RUC & "')"
        'Call ConfiguraRstPP(strCadena)
        '--- generacion de asiento
        
        in_mis_cuentas_det = procesar_transaccion(KEY_ALM, Me.DtcCuentaBancaria.BoundText, Format(Me.DtpEmision.Value, "YYYY-mm-dd"), "00001", rstK("id_cliente"), rstK("ncliente"), in_glosa, Format(Me.txtMontoCobrado.Text, "###0.00"), "0", rstK("id_venta"), "0", rstK("documento"), KEY_CAMBIO, Me.txtOperacionVenta.Text, get_forma_pago_cuenta(Me.DtcCuentaBancaria.BoundText), "1CIX000000000025", "00001", KEY_USUARIO, KEY_RUC)
        '--- insert mis cuentas_det detalle
        Call put_realizar_pago(in_venta, in_venta, Format(Me.txtMontoCobrado.Text, "###0.00"), rstK("id_doc"), KEY_CAMBIO, Val(in_mis_cuentas_det))
        
        
        
    End If
End If
    
    
    '*********REALIZAR EL EGRESO A LA CUENTA
    'strCadena = "call ADM_impresion_comprobantes('20','" & in_venta & "','" & KEY_RUC & "')"
    'Call ConfiguraRstA(strCadena)
    'If rstA.RecordCount < 1 Then
    '##-------------------------------------------------------------------------
        If Val(Me.txtComision.Text) > 0 Then
            in_glosa = Me.DtcFlujo.Text + Space(2) + "[ " + in_documento + " ]"
            strCadena = "call sp_procesar_transaccion_caja_venta('" & KEY_ALM & "','" & Me.DtcCuentaBancaria.BoundText & "','" & Format(Me.DtpEmision.Value, "YYYY-mm-dd") & "','00002','" & rstK("id_cliente") & "','" & in_glosa & "','" & Val(Me.txtComision.Text) & "','','" & rstK("tc") & "','" & rstK("operacion") & "','" & KEY_USUARIO & "','" & rstK("id_venta") & "','0','" & rstK("ncliente") & "','" & rstK("documento") & "','" & Me.DtcFlujo.BoundText & "','" & rstK("id_forma_pago") & "','" & KEY_RUC & "')"
            Call ConfiguraRstPP(strCadena)
            
            strCadena = "call sp_comision_bancario('" & rstPP("in_id") & "')"
            CnBd.Execute (strCadena)
        End If
     '##------------------------------------------------------------------------
    'End If
    '*********************************************
  


strCadena = "call ADM_servicios_generales('29','" & in_venta & "','','','','','" & KEY_RUC & "')"
Call ConfiguraRstK(strCadena)
in_total_comprobante = rstK(0)

strCadena = "call ADM_servicios_generales('21','" & in_venta & "','','','','','" & KEY_RUC & "')"
Call ConfiguraRstK(strCadena)
If rstK(0) < in_total_comprobante Then
    strCadena = "UPDATE movimiento_venta SET pendiente='si' WHERE id_venta='" & in_venta & "'"
    Call ConfiguraRst(strCadena)
End If







Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 12) = Me.DtcVendedorComprobante.Text
Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 13) = Me.DtcCuentaBancaria.Text
Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 14) = Format(Val(Me.txtComision.Text), "#,##0.00")
Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 15) = Me.txtOperacionVenta.Text
Me.HfdPersona.col = 16
If Me.chk_verificado.Value = 1 Then
    HfdPersona.TextMatrix(Me.HfdPersona.Row, 16) = "VERIFICADO"
    Me.HfdPersona.CellBackColor = &H80FF80
    
Else
    Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 16) = "PENDIENTE"
    Me.HfdPersona.CellBackColor = &H8080FF
End If

Call get_comprobante(in_venta)

End Sub


Private Sub get_comprobante(ByVal in_venta As String)

strCadena = "call ADM_servicios_generales('21','" & in_venta & "','','','','','" & KEY_RUC & "')"
Call ConfiguraRstK(strCadena)
Me.lbl_pagado.Caption = rstK(0)


strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & in_venta & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.lblcomprobante.Tag = in_venta
   Me.CmdSiguiente.Tag = Me.HfdPersona.Row
   Me.lblcomprobante.Caption = rst("documento")
   Me.DtpEmision.Value = rst("fecha_emision")
   
   If rst("pendiente") = "si" Then
      Me.chk_verificado.Value = 0
   Else
      Me.chk_verificado.Value = 1
   End If
   
   If rst("tc_local") > 0 Then
      Me.txtComision.Text = Format(rst("tc_local"), "#,##0.00")
      Me.DtcFlujo.BoundText = rst("orden_compra")
   Else
      If rst("anulado") = "si" Then
           Me.txtComision.Text = Format(0, "#,##0.00")
           Me.DtcFlujo.BoundText = "1CIX000000000025"
      Else
            Me.txtComision.Text = Format(0, "#,##0.00")
            Me.DtcFlujo.BoundText = "1CIX000000000025"
      End If
      
   End If
   
   Me.DtcVendedorComprobante.BoundText = rst("id_vendedor")
   Me.DtcCuentaBancaria.BoundText = rst("id_orden_salida")
   Me.txtMontoCobrado.Text = Format(rst("total"), "#,##0.00")
   
   Me.lbl_saldo.Caption = Val(rst("total") - Val(Me.lbl_pagado.Caption))
   
   
   
   Me.txtObservacion.Text = ""
   Me.txtOperacionVenta.Text = rst("operacion")
   If rst("observacion") = "-" Then
      Me.txtObservaciones.Text = ""
   Else
      Me.txtObservaciones.Text = UCase(rst("observacion"))
   End If
   Me.HfdPersona.Enabled = False
   Me.frmdetalle.Visible = True
   Call Resalta(Me.txtComision)
End If
End Sub



Private Sub HfdPersona_DblClick()

Call get_comprobante(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
  

End If
End Sub

Private Sub HfdPersona_GotFocus()
HookForm Me.HfdPersona
End Sub

Private Sub HfdPersona_LostFocus()
UnHookForm Me.HfdPersona
End Sub


Private Sub txtCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "CALL ADM_reportes_generales('35','" & Format(Me.DtpFecha.Tag, "YYYY-mm-dd") & "','','','','','','','','" & Trim(Me.TxtCliente.Text) & "','" & KEY_RUC & "')"
    Call llenarGrid(Me.HfdPersona)
End If
End Sub





Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "CALL ADM_reportes_generales('35','" & Format(Me.DtpFecha.Tag, "YYYY-mm-dd") & "','','" & Trim(Me.TxtNumero.Text) & "','','','','','','','" & KEY_RUC & "')"
    Call llenarGrid(Me.HfdPersona)
End If
End Sub

























Private Sub txtOperacionBusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "CALL ADM_reportes_generales('35','" & Format(Me.DtpFecha.Tag, "YYYY-mm-dd") & "','','','" & Trim(Me.txtOperacionBusqueda.Text) & "','','','','','','" & KEY_RUC & "')"
    Call llenarGrid(Me.HfdPersona)
End If

End Sub

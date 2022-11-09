VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmventaslistado 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13005
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmTienda 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MI TIENDA"
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
      Height          =   1335
      Left            =   10680
      TabIndex        =   42
      Top             =   4200
      Width           =   2295
      Begin VB.TextBox Text1 
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
         Left            =   680
         TabIndex        =   43
         Top             =   320
         Width           =   1455
      End
      Begin VitekeySoft.ChameleonBtn cmdImportarTienda 
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "IMPORTAR"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmventaslistado.frx":0000
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
         Caption         =   "CODIGO:"
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
         Left            =   40
         TabIndex        =   44
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Frame frmmasivo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FACTURACION MASIVA"
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
      Height          =   8175
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   10215
      Begin VB.TextBox txtForma_pago 
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
         Left            =   7080
         TabIndex        =   30
         Top             =   6120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Timer timer_masivo 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   6480
         Top             =   3720
      End
      Begin VB.TextBox txtCantidadFacturada 
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
         Left            =   6600
         TabIndex        =   29
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox TxtIdVenta 
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
         Left            =   7080
         TabIndex        =   28
         Top             =   6840
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txt_sunat_key 
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
         Left            =   7080
         TabIndex        =   27
         Top             =   7560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txt_hash 
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
         Left            =   7080
         TabIndex        =   26
         Top             =   7200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtPrecio 
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
         Left            =   2400
         TabIndex        =   25
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtDireccion 
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
         Left            =   2400
         TabIndex        =   23
         Top             =   1440
         Width           =   5175
      End
      Begin VB.TextBox txtcantidad_maxima 
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
         Left            =   2400
         TabIndex        =   20
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtcantidad_facturar 
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
         Left            =   2400
         TabIndex        =   18
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtproductomasivo 
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
         Left            =   3840
         TabIndex        =   16
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtid_productomasivo 
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
         Left            =   2400
         TabIndex        =   15
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtclientemasivo 
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
         Left            =   3840
         TabIndex        =   13
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox txtdnimasivo 
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
         Left            =   2400
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DtpEmisionMasivo 
         Height          =   300
         Left            =   2400
         TabIndex        =   10
         Top             =   600
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   194772993
         CurrentDate     =   43618
      End
      Begin MSComctlLib.ProgressBar prog_indicador 
         Height          =   225
         Left            =   2400
         TabIndex        =   21
         Top             =   3720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VitekeySoft.ChameleonBtn cmdMasivo 
         Height          =   615
         Left            =   2400
         TabIndex        =   36
         Top             =   4080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "INICIAR LA FACTURACION        "
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
         MICON           =   "frmventaslistado.frx":001C
         PICN            =   "frmventaslistado.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image cmdCerrar_masivo 
         Height          =   240
         Left            =   9840
         Picture         =   "frmventaslistado.frx":3D79
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD FACTURADA :"
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
         Left            =   4800
         TabIndex        =   31
         Top             =   2880
         Width           =   1620
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO :"
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
         Left            =   1470
         TabIndex        =   24
         Top             =   2280
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCION :"
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
         TabIndex        =   22
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANT.MAX POR DOCUMENTO:"
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
         TabIndex        =   19
         Top             =   3240
         Width           =   2010
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD A FACTURAR:"
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
         Left            =   435
         TabIndex        =   17
         Top             =   2880
         Width           =   1605
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCTO :"
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
         TabIndex        =   14
         Top             =   2040
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE   :"
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
         TabIndex        =   11
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label3 
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
         Left            =   915
         TabIndex        =   9
         Top             =   600
         Width           =   1125
      End
   End
   Begin VB.CheckBox chk_all 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "SELECC TODOS"
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
      Left            =   8880
      TabIndex        =   40
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtComprobante 
      Height          =   285
      Left            =   10560
      TabIndex        =   39
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer_masivo_todo 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   12480
      Top             =   5880
   End
   Begin VitekeySoft.ChameleonBtn cmdFacturadorMasivo 
      Height          =   615
      Left            =   10560
      TabIndex        =   33
      Top             =   7680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "FACTURAR MASIVO"
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
      MICON           =   "frmventaslistado.frx":6C1D
      PICN            =   "frmventaslistado.frx":6C39
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcVendedor 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   4335
      _ExtentX        =   7646
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
   Begin VB.CheckBox Check1 
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
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   10560
      ScaleHeight     =   915
      ScaleWidth      =   2235
      TabIndex        =   5
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtcliente 
      Appearance      =   0  'Flat
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
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   4335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
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
      CalendarForeColor=   8388608
      CalendarTitleForeColor=   8388608
      Format          =   194772993
      CurrentDate     =   42236
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPendientes 
      Height          =   7095
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   12515
      _Version        =   393216
      ForeColor       =   8388608
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
   Begin MSComctlLib.ProgressBar prg_mobil 
      Height          =   225
      Left            =   10560
      TabIndex        =   32
      Top             =   2175
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VitekeySoft.ChameleonBtn cmdImportarMovil 
      Height          =   615
      Left            =   10560
      TabIndex        =   34
      Top             =   1560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "IMPORTAR MOBILE"
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
      BCOL            =   33023
      BCOLO           =   33023
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmventaslistado.frx":A97A
      PICN            =   "frmventaslistado.frx":A996
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmddownload 
      Height          =   615
      Left            =   10560
      TabIndex        =   35
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "DOWNLOAD DOCUMENT"
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
      MICON           =   "frmventaslistado.frx":D7B7
      PICN            =   "frmventaslistado.frx":D7D3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdconsultar 
      Height          =   615
      Left            =   6120
      TabIndex        =   37
      Top             =   480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "CONSULTAR"
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
      MICON           =   "frmventaslistado.frx":105F4
      PICN            =   "frmventaslistado.frx":10610
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdSeleccionados 
      Height          =   615
      Left            =   10560
      TabIndex        =   38
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "FACTURAR SELECCIONADOS"
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
      MICON           =   "frmventaslistado.frx":12EFA
      PICN            =   "frmventaslistado.frx":12F16
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   300
      Left            =   3480
      TabIndex        =   41
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
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
      CalendarForeColor=   8388608
      CalendarTitleForeColor=   8388608
      Format          =   194772993
      CurrentDate     =   42236
   End
   Begin VB.Image cmdsalir 
      Height          =   240
      Left            =   12600
      Picture         =   "frmventaslistado.frx":16C57
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CLIENTE  :"
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
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FECHA     :"
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
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   8505
      Left            =   0
      Top             =   0
      Width           =   13005
   End
End
Attribute VB_Name = "frmventaslistado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private obj_Word As Object
Public Procedencia As EnumProcede

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub chk_all_Click()
If Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) > 0 Then
    For i = 0 To Me.HfPendientes.Rows - 1
    
    strCadena = "SELECT seleccion,referencia FROM movimiento_venta WHERE id_venta='" & Val(Me.HfPendientes.TextMatrix(i, 0)) & "' and ruc='" & KEY_RUC & "'  LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       If Len(rst("referencia")) < 5 Then
            If rst("seleccion") = "no" Then
                strCadena = "UPDATE movimiento_venta SET seleccion='si' WHERE   id_venta='" & Val(Me.HfPendientes.TextMatrix(i, 0)) & "' and ruc='" & KEY_RUC & "'   LIMIT 1"
                in_char = Chr(254)
            Else
                strCadena = "UPDATE movimiento_venta SET seleccion='no' WHERE id_venta='" & Val(Me.HfPendientes.TextMatrix(i, 0)) & "' and  ruc='" & KEY_RUC & "'   LIMIT 1"
                in_char = Chr(168)
            End If
               
      Else
          strCadena = "UPDATE movimiento_venta SET seleccion='no' WHERE id_venta='" & Val(Me.HfPendientes.TextMatrix(i, 0)) & "' and  ruc='" & KEY_RUC & "'   LIMIT 1"
          in_char = Chr(168)
        End If
            CnBd.Execute (strCadena)
            Me.HfPendientes.TextMatrix(i, 6) = in_char
    End If
            
   
    Next i
    
    
End If
End Sub

Private Sub cmdCerrar_masivo_Click()
Me.frmmasivo.Visible = False
End Sub

Private Sub cmdConsultar_Click()

Dim strpersona As String
strpersona = ""
strpersona = Trim(Me.TxtCliente.Text)

strCadena = "SELECT id_venta,ncliente,documento,total,referencia,vendedor,seleccion FROM view_listado_pendientes_ref WHERE ncliente LIKE '%" & strpersona & "%' AND  fecha_emision>='" & Format(Me.DTPicker1.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DTPicker2.Value, "YYYY-mm-dd") & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' ORDER BY vendedor,id_venta"
Call llenar_pendientes(Me.HfPendientes)
End Sub
Private Sub get_foto(ByVal in_venta As String)
strCadena = "SELECT imagen,id_producto FROM view_movimiento_venta_detalle WHERE id_venta='" & in_venta & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    '--------- foto--------
strCadena = "SELECT id_producto FROM producto_foto WHERE id_producto='" & rst("id_producto") & "' AND ruc='" & KEY_RUC & "' "
Call ConfiguraRstK(strCadena)
If IsNull(rst("imagen")) = False And Len(rst("imagen")) > 5 Then
    If VerificarArchivo(App.Path & "\archivos\" & KEY_RUC & "\" & rst("imagen")) = True Then
        Me.Picture1.Picture = LoadPicture(App.Path + "\archivos\" + KEY_RUC + "\" + Trim(rst("imagen")))
    Else
        Me.Picture1.Picture = LoadPicture(App.Path + "\imagenes\no_disponible.jpg")
    End If
Else
     Me.Picture1.Picture = LoadPicture(App.Path + "\imagenes\no_disponible.jpg")
End If
'-------------------------
End If
End Sub
Private Sub cmddownload_Click()
Call get_foto(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0))
Call put_plantilla(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0))

End Sub

Private Sub put_plantilla(ByVal in_comprobante As String)

Dim Ret As Boolean
Dim Destino As String
Dim Origen As String
Dim in_archivo As String
in_archivo = "\" & in_comprobante & ".doc"
Destino = App.Path & "\archivos\" & KEY_RUC & in_archivo 'Vos pones el directorio destino en el textbox
Origen = App.Path & "\archivos\plantilla.doc" 'Vos pones el archivo a copiar en el textbox
FileCopy Origen, Destino
Ret = Imagen_a_Word(Destino, _
                        Picture1.Picture, _
                        "Foto", in_comprobante)
    ' Si devuelve True no hubo error
    
End Sub

' Función que pega el gráfico en el marcador del documento de word
'*******************************************************************
Function Imagen_a_Word(Path_Word As String, Grafico As Picture, marcador As String, ByVal in_venta As String) As Boolean

On Local Error GoTo ErrFunction
    
    
    ' Nueva instancia de Word
    Set obj_Word = CreateObject("Word.Application")
    
    
    ' Abre el documento
    obj_Word.Documents.Open _
        FileName:=Path_Word, _
        ConfirmConversions:=False, _
        ReadOnly:=False, _
        AddToRecentFiles:=False, _
        PasswordDocument:="", _
        PasswordTemplate:="", _
        Revert:=False, _
        WritePasswordDocument:="", _
        WritePasswordTemplate:="", _
        Format:=0

    ' Ubica la selección en el marcador del documento de word
    obj_Word.Selection.Goto _
        What:=-1, _
        name:=marcador

    ' Limpia el Clipboard
    Clipboard.Clear

    ' Pasa el gráfico al portapapeles
    Clipboard.SetData Grafico, vbCFBitmap

    ' Pega la imagen en la selección
    obj_Word.Selection.Paste
    
    strCadena = "SELECT * FROM view_movimiento_venta_detalle WHERE id_venta='" & in_venta & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    With obj_Word.ActiveDocument.Bookmarks
        
        .Item("Datoscliente").Range.Text = Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 2)
        
        .Item("Precio").Range.Text = Format(rst("precio"), "#,##00.00")
        .Item("Anio").Range.Text = "2019"
        .Item("Linea").Range.Text = rst("linea")
        .Item("Marca").Range.Text = rst("marca")
        .Item("Marcaii").Range.Text = rst("marca")
        .Item("Modelo").Range.Text = rst("modelo")
    
        strCadena = "SELECT * FROM producto_importaciones WHERE id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
           rstL.MoveFirst
           .Item("Potencia").Range.Text = rstL("potencia_motor")
        End If
       
    End With
   End If




    ' Cierra el word, guarda los cambios y elimina la referencia
  Descargar_Word (in_venta)
    Imagen_a_Word = True
    
    Exit Function
    
' Error
ErrFunction:

MsgBox Err.Description, vbCritical
On Error Resume Next
Descargar_Word (in_venta)
    
End Function
Private Sub Descargar_Word(ByVal in_venta As String)

    Dim Destino As String
    If Not obj_Word Is Nothing Then
        ' Cierra el word y guarda los cambios
        obj_Word.Quit True
        ' Elimina la referencia
         Set obj_Word = Nothing
    End If
    
    in_archivo = "\" & in_venta & ".doc"
    Destino = App.Path & "\archivos\" & KEY_RUC & in_archivo

    MsgBox "Se ha generado Satisfactoriamente", vbInformation
        Call AbreArchivo(Destino)
    
End Sub

Private Sub cmdFacturadorMasivo_Click()
Me.DtpEmisionMasivo.Value = KEY_FECHA
Me.txtdnimasivo.Text = "00000000"
Me.txtclientemasivo.Text = get_persona(Trim(Me.txtdnimasivo.Text))
Me.TxtDireccion.Text = KEY_DIRECCION_ALM
Call Resalta(Me.txtid_productomasivo)
Me.frmmasivo.Visible = True



End Sub

Private Sub facturar_masivo()
Dim cantidad_facturada As Double
Dim in_cantidad As Single
Dim in_monto As Single
Dim in_forma_pago As Double
Dim in_acumulado As Double

    Me.timer_masivo.Enabled = False
    
    in_cantidad = Aleatorio(1, Val(Me.txtcantidad_maxima.Text))
    If in_cantidad > 0 Then
    
    
    If Val(Me.txtCantidadFacturada.Text) <= Val(Me.txtcantidad_facturar.Text) And in_cantidad > 0 Then
        If (Val(Me.txtCantidadFacturada.Text) + in_cantidad) < Val(Me.txtcantidad_facturar.Text) Then
            in_cantidad = in_cantidad
        Else
            in_cantidad = Val(Me.txtcantidad_facturar.Text) - Val(Me.txtCantidadFacturada.Text)
        End If
        If in_cantidad > 0 Then
            Me.txtCantidadFacturada.Text = Val(Me.txtCantidadFacturada) + in_cantidad
            Call put_factura(Trim(Me.txtdnimasivo.Text), Trim(Me.txtclientemasivo.Text), Trim(Me.TxtDireccion.Text), Val(Me.txtprecio.Text), "-", Val(Me.txtForma_pago.Text), Trim(Me.txtid_productomasivo.Text), in_cantidad, Trim(Me.txtproductomasivo.Text))
            Me.timer_masivo.Enabled = False
        Else
            Me.timer_masivo.Enabled = False
        End If
   
    End If
        Me.prog_indicador.Value = Val(Me.txtCantidadFacturada.Text) - 1
    End If
End Sub




Private Function Aleatorio(Minimo As Long, Maximo As Long) As Long
    Randomize ' inicializar la semilla
    Aleatorio = CLng((Minimo - Maximo) * Rnd + Maximo)
End Function

Private Sub cmdImportarMovil_Click()

Call get_importar_pedido(Me.DTPicker1.Value, DateAdd("d", 2, Me.DTPicker1.Value))

strCadena = "SELECT id_venta,ncliente,documento,total,referencia,vendedor,seleccion FROM view_listado_pendientes_ref WHERE ncliente LIKE '%" & strpersona & "%' AND  fecha_emision='" & Format(Me.DTPicker1.Value, "YYYY-mm-dd") & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' ORDER BY id_venta ASC"
Call llenar_pendientes(Me.HfPendientes)
End Sub

Private Sub cmdImportarTienda_Click()
Call get_importar_MITIENDA(Trim(Me.Text1.Text))
End Sub

Private Sub cmdMasivo_Click()

Me.txtForma_pago.Text = get_forma_pago_detalle_contado
Me.txtCantidadFacturada.Text = 0
Me.timer_masivo.Enabled = True
Me.prog_indicador.Min = 0
Me.prog_indicador.Max = Val(Me.txtcantidad_facturar.Text) + 1


End Sub

Private Sub cmdSalir_Click()
Call enabled_form(FrmVentas)
Unload Me
End Sub
Public Function get_importar_pedido(ByVal in_fecha_ini As String, ByVal in_fecha_fin As String) As Boolean
Dim strHtml As String
Set DomDoc = New XMLHTTP
     'urlstr = "https://api.vitekey.com/keyfact/utils/reporte-ventas?password=vitekey2018&ruc=" & KEY_RUC & "&date_start=" & Format(in_fecha_ini, "MM/dd/YYYY") & "&date_end=" & Format(in_fecha_fin, "MM/dd/YYYY")
     
     urlstr = "https://api.vitekey.com/keyfact/erp/proformas?password=vitekey2018&ruc=" & KEY_RUC & "&date_start=" & Format(in_fecha_ini, "MM/dd/YYYY") & "&date_end=" & Format(in_fecha_fin, "MM/dd/YYYY")
     
     Set DomDoc = New XMLHTTP
     'Parámetros en formato URLEncode
     'Metodo a usar, url, y true en caso de manejar la respuesta en modo asíncrono
     DomDoc.Open "GET", urlstr, False
     'encabezados
     DomDoc.setRequestHeader "Connection", "close"
     DomDoc.send params
     'La respuesta, en caso de existir, está en responseBody.
    'También puedes especificar responseXml si tu aplicación devolviese XML
     strHtml = StrConv(DomDoc.responseBody, vbUnicode)
     
     Call procesar_pedidos(strHtml)

End Function

Public Function get_importar_MITIENDA(ByVal in_codigo As String) As Boolean
Dim strHtml As String
Set DomDoc = New XMLHTTP
     
     
     urlstr = "https://api.mitienda.pe/v1/mitienda/order?code=" & in_codigo
     
     Set DomDoc = New XMLHTTP
     'Parámetros en formato URLEncode
     'Metodo a usar, url, y true en caso de manejar la respuesta en modo asíncrono
     DomDoc.Open "GET", urlstr, False
     
     'encabezados
     
     
     
     
     DomDoc.setRequestHeader "Content-type", "application/json"
     DomDoc.setRequestHeader "token", "EPMkXspWiaSIWmOrRKXXMuGrAhwgaorE"
     DomDoc.setRequestHeader "Connection", "close"
     DomDoc.send params
     'La respuesta, en caso de existir, está en responseBody.
    'También puedes especificar responseXml si tu aplicación devolviese XML
     strHtml = StrConv(DomDoc.responseBody, vbUnicode)
     
     Call procesar_pedido_mitienda(strHtml)

End Function



Public Sub procesar_pedido_mitienda(ByVal strHtml As String)
Dim in_error As Boolean
Dim in_hash As String
Dim in_total_comprobante As Double
Dim in_temporal As Integer
Dim in_numero() As String
Dim json_r As Object
Set json_r = JSON.parse(strHtml)

If json_r.Count >= 1 Then
   
    in_doc = "0099"
    in_serie = "004"
    
    
    in_nume = get_numero_comprobante(in_doc, in_serie)
    in_dni = json_r.Item("billing_info").Item("doc_number")
    in_cliente = UCase(json_r.Item("billing_info").Item("name")) & Space(1) & UCase(json_r.Item("billing_info").Item("last_name"))
    in_direccion = UCase(json_r.Item("billing_info").Item("billing_address").Item("address_line"))
    
    in_mail = json_r.Item("billing_info").Item("email")
    in_celular = json_r.Item("billing_info").Item("phone_number")
    
    in_fecha = Format(json_r.Item("date_created"), "YYYY-mm-dd")
    
    in_ubigeo = Format(json_r.Item("billing_info").Item("billing_address").Item("state").Item("id"), "00")
    in_ubigeo = in_ubigeo + Format(json_r.Item("billing_info").Item("billing_address").Item("city").Item("id"), "00")
    in_ubigeo = in_ubigeo + Format(json_r.Item("billing_info").Item("billing_address").Item("distric").Item("id"), "00")
     
     
     
     strCadena = "call p_nueva_venta_v11('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
     CnBd.Execute (strCadena)
       
       
    strCadena = "SELECT * FROM persona WHERE dni='" & in_dni & "' LIMIT 1"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount < 1 Then
        strCadena = " call p_insert_persona_iii('" & Trim(in_dni) & "','-','-','-','" & Replace(Trim(in_cliente), "'", "") & "','" & Replace(in_direccion, "'", "") & "','','-','no','no','no','no','no','no','si','" & KEY_DEPARTAMENTO & "','" & KEY_PROVINCIA & "','" & KEY_DISTRITO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
    End If
    
    
    strCadena = "UPDATE persona SET codigo_ubigeo_sunat='" & in_ubigeo & "',celular='" & in_celular & "',mail='" & in_mail & "' WHERE dni='" & in_dni & "' LIMIT 1"
    CnBd.Execute (strCadena)
    
    
    strCadena = "SELECT cod_unico FROM entidad_empresa WHERE cod_unico='" & in_dni & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount < 1 Then
        strCadena = "INSERT INTO entidad_empresa (cod_unico,id_empresa,id_cliente)VALUES('" & in_dni & "','" & KEY_RUC & "','si')"
        CnBd.Execute (strCadena)
    End If
    
    in_vendedor = KEY_USUARIO
    'in_vendedor = get_dni(UCase(json_r(i).Item("seller_name")))
    
    If get_existe_comprobante(in_doc, in_serie, in_nume) = False Then
Iniciar:
       in_documento = "P:" & in_serie & "-" & in_nume
    
   
    IN_TOTAL_TOTAL = json_r.Item("total_amount")
    in_descuento_global = 0
    If IN_TOTAL_TOTAL = 0 And in_descuento_global > 0 Then
        in_obsequio = "si"
    Else
        in_obsequio = "no"
    End If
    
    If KEY_APLICA_IGV = "si" Then
        in_valor_venta = IN_TOTAL_TOTAL / (1 + KEY_IGV)
    Else
        in_valor_venta = IN_TOTAL_TOTAL
    End If
    
    
    in_total_exonerado = 0
    in_total_igv = IN_TOTAL_TOTAL - in_valor_venta
    in_total_comprobante = 0
    in_temporal = 0
    in_costo_envio = json_r.Item("shipping").Item("cost")
    in_observacion = "CODIGO:" + json_r.Item("code") + Chr(13) + "COSTO ENVIO:" + in_costo_envio
    
    
    For j = 1 To json_r.Item("order_items").Count ' recorro la cantidad de productos
        in_producto = Format(json_r.Item("order_items")(j).Item("sku"), "00000")
        in_cantidad = json_r.Item("order_items")(j).Item("quantity")
        in_precio = json_r.Item("order_items")(j).Item("unit_price")
        in_detalle = get_producto_comercial(in_producto) 'json_r(i).Item("items")(j).Item("description")
        If in_obsequio = "si" Then
            in_total = 0
        Else
            in_total = Val(in_precio) * Val(in_cantidad)
        End If
        in_total_comprobante = in_total_comprobante + in_total
        
        
       ' If control_stock_pedido(in_producto, in_cantidad, in_documento) = True Then'
                    
            strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,servicio,obsequio,costo) VALUES " & _
            "('" & KEY_RUC & "','" & get_unidad_producto(in_producto) & "','" & in_dni & "','" & KEY_ALM & "','" & in_doc & "','" & in_serie & "','" & in_nume & "','" & in_producto & "','" & Val(in_cantidad) & "'," & _
            "'" & Val(in_precio) & " ','" & in_total & "','0','" & KEY_APLICA_IGV & "','" & in_detalle & "','" & KEY_USUARIO & "','no','" & in_obsequio & "','" & get_precio_costo(in_producto) & "')"
            CnBd.Execute (strCadena)
            in_temporal = in_temporal + 1
         'End If
        
    Next j
    
              If Round(Val(in_total_comprobante), 2) = Round(Val(IN_TOTAL_TOTAL), 2) Then
                If in_temporal > 0 Then
                    'Call get_auto_pago_main(in_doc, in_serie, in_nume, in_total_total)
                
                    strCadena = "call p_insert_venta_cabecera_premiun_ref('" & in_doc & "','" & KEY_ALM & "','01','00001','no'," & _
                    "'" & in_serie & "','" & in_nume & "','" & in_dni & "','" & Replace(in_cliente, "'", "") & "','" & in_valor_venta & "','" & in_total_igv & "','" & in_total_exonerado & "','" & IN_TOTAL_TOTAL & "','0'," & _
                    "'" & Val(in_total_exonerado) & "','0','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','00001','" & in_vendedor & "','" & KEY_USUARIO & "','" & Val(KEY_CAMBIO_COMPRA) & "','no','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "'" & _
                    ",'" & in_documento & "',CURTIME(),'T','" & in_direccion & "','no','-','-',' ',' ',' ',' ','" & KEY_VENTANILLA & "','01','" & in_seguro & "','" & in_observacion & "','no','" & KEY_CONTABILIDAD & "','" & in_cta_cobrar & "','" & in_cta_ingreso & "','" & in_descuento_global & "','0','0','0','no','0','" & KEY_RUC & "')"
                    Call ConfiguraRstPP(strCadena)
                    id_venta = rstPP("in_venta")
                    Call put_correlativo_venta(in_doc, in_serie, in_nume)
                End If
             Else
                
                If MsgBox("Ha Ocurrido una Inconsistencia con este Pedido" + Space(2) + in_documento + Chr(13) + "Desea Pasarlo de Nuevo.", vbInformation + vbYesNo) = vbYes Then
                    strCadena = "call p_nueva_venta_v11('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    GoTo Iniciar
                End If
             End If
       End If
       DoEvents

   
End If
strCadena = "call p_nueva_venta_v11('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

MsgBox "Proceso Correcto", vbInformation


End Sub




Public Sub procesar_pedidos(ByVal strHtml As String)
Dim in_error As Boolean
Dim in_hash As String
Dim in_total_comprobante As Double
Dim in_temporal As Integer
Dim in_numero() As String
Dim json_r As Object
Set json_r = JSON.parse(strHtml)

If json_r.Count >= 1 Then
   Me.prg_mobil.Min = 1
   Me.prg_mobil.Max = json_r.Count + 1
For i = 1 To json_r.Count  ' recorro la cantidad de comprobantes
    in_doc = "0099"
    in_serie = "007"
    in_nume = Format(json_r(i).Item("number"), "000000")
    in_dni = json_r(i).Item("client_docid")
    in_cliente = json_r(i).Item("client_name")
    in_direccion = json_r(i).Item("client_address")
    
    
     
     
     strCadena = "call p_nueva_venta_v11('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
     CnBd.Execute (strCadena)
       
       
    strCadena = "SELECT * FROM persona WHERE dni='" & in_dni & "'"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount < 1 Then
        strCadena = " call p_insert_persona_iii('" & Trim(in_dni) & "','-','-','-','" & Replace(Trim(in_cliente), "'", "") & "','" & Replace(in_direccion, "'", "") & "','','-','no','no','no','no','no','no','si','" & KEY_DEPARTAMENTO & "','" & KEY_PROVINCIA & "','" & KEY_DISTRITO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
    Else
        If Len(in_direccion) > 10 Then
            strCadena = "UPDATE persona SET direccion='" & Replace(in_direccion, "'", "") & "' WHERE dni='" & in_dni & "' LIMIT 1"
            CnBd.Execute (strCadena)
        Else
            in_direccion = rstA("direccion")
        End If
    End If
    
    
    strCadena = "SELECT cod_unico FROM entidad_empresa WHERE cod_unico='" & in_dni & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount < 1 Then
        strCadena = "INSERT INTO entidad_empresa (cod_unico,id_empresa,id_cliente)VALUES('" & in_dni & "','" & KEY_RUC & "','si')"
        CnBd.Execute (strCadena)
    End If
    
    
    
    
    
    
    in_direccion = get_direccion(in_dni)
    in_vendedor = get_dni_keyfacil(json_r(i).Item("account_id"))
    'in_vendedor = get_dni(UCase(json_r(i).Item("seller_name")))
    
    If get_existe_comprobante(in_doc, in_serie, in_nume) = False Then
Iniciar:
       in_documento = "P:" & in_serie & "-" & in_nume
    
    in_observacion = "-"
    IN_TOTAL_TOTAL = json_r(i).Item("computed").Item("total")
    in_descuento_global = json_r(i).Item("computed").Item("total_discount")
    If IN_TOTAL_TOTAL = 0 And in_descuento_global > 0 Then
        in_obsequio = "si"
    Else
        in_obsequio = "no"
    End If
    
    If KEY_APLICA_IGV = "si" Then
        in_valor_venta = IN_TOTAL_TOTAL / (1 + KEY_IGV)
    Else
        in_valor_venta = IN_TOTAL_TOTAL
    End If
    
    
    in_total_exonerado = json_r(i).Item("computed").Item("total_exonerated")
    in_total_igv = json_r(i).Item("computed").Item("total_igv")
    in_total_comprobante = 0
    in_temporal = 0
    
    For j = 1 To json_r(i).Item("items").Count ' recorro la cantidad de productos
        in_producto = Format(json_r(i).Item("items")(j).Item("code"), "00000")
        in_cantidad = json_r(i).Item("items")(j).Item("quantity")
        in_precio = json_r(i).Item("items")(j).Item("unit_price")
        in_detalle = get_producto(in_producto) 'json_r(i).Item("items")(j).Item("description")
        If in_obsequio = "si" Then
            in_total = 0
        Else
            in_total = Val(in_precio) * Val(in_cantidad)
        End If
        in_total_comprobante = in_total_comprobante + in_total
        
        
        If control_stock_pedido(in_producto, in_cantidad, in_documento) = True Then
                    
            strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,servicio,obsequio,costo) VALUES " & _
            "('" & KEY_RUC & "','" & get_unidad_producto(in_producto) & "','" & in_dni & "','" & KEY_ALM & "','" & in_doc & "','" & in_serie & "','" & in_nume & "','" & in_producto & "','" & Val(in_cantidad) & "'," & _
            "'" & Val(in_precio) & " ','" & in_total & "','0','" & KEY_APLICA_IGV & "','" & in_detalle & "','" & KEY_USUARIO & "','no','" & in_obsequio & "','" & get_precio_costo(in_producto) & "')"
            CnBd.Execute (strCadena)
            in_temporal = in_temporal + 1
         End If
        
    Next j
    
              If Round(Val(in_total_comprobante), 2) = Round(Val(IN_TOTAL_TOTAL), 2) Then
                If in_temporal > 0 Then
                    Call get_auto_pago_main(in_doc, in_serie, in_nume, IN_TOTAL_TOTAL)
                
                    strCadena = "call p_insert_venta_cabecera_premiun_ref('" & in_doc & "','" & KEY_ALM & "','01','00001','no'," & _
                    "'" & in_serie & "','" & in_nume & "','" & in_dni & "','" & Replace(in_cliente, "'", "") & "','" & in_valor_venta & "','" & in_total_igv & "','" & in_total_exonerado & "','" & IN_TOTAL_TOTAL & "','0'," & _
                    "'" & Val(in_total_exonerado) & "','0','" & Format(Me.DTPicker1.Value, "YYYY-mm-dd") & "','" & Format(Me.DTPicker1.Value, "YYYY-mm-dd") & "','00001','" & in_vendedor & "','" & KEY_USUARIO & "','" & Val(KEY_CAMBIO_COMPRA) & "','no','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "'" & _
                    ",'" & in_documento & "',CURTIME(),'T','" & in_direccion & "','no','-','-',' ',' ',' ',' ','" & KEY_VENTANILLA & "','01','" & in_seguro & "','" & in_observacion & "','no','" & KEY_CONTABILIDAD & "','" & in_cta_cobrar & "','" & in_cta_ingreso & "','" & in_descuento_global & "','0','0','0','no','0','" & KEY_RUC & "')"
                    Call ConfiguraRstPP(strCadena)
                    id_venta = rstPP("in_venta")
                    Call put_correlativo_venta(in_doc, in_serie, in_nume)
                End If
             Else
                
                If MsgBox("Ha Ocurrido una Inconsistencia con este Pedido" + Space(2) + in_documento + Chr(13) + "Desea Pasarlo de Nuevo.", vbInformation + vbYesNo) = vbYes Then
                    strCadena = "call p_nueva_venta_v11('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    GoTo Iniciar
                End If
             End If
       End If
       DoEvents

       Me.prg_mobil.Value = i
Next i
End If
strCadena = "call p_nueva_venta_v11('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)


End Sub

Private Sub cmdSeleccionados_Click()

If MsgBox("ESTA SEGURO DE FACTURAR TODOS" + Chr(13) + "PEDIDOS SELECCIONADOS", vbYesNo + vbQuestion) = vbYes Then
    Me.txtForma_pago.Text = get_forma_pago_detalle_contado
    Me.prog_indicador.Min = 0
    Me.prog_indicador.Max = Me.HfPendientes.Rows - 1
    Me.Timer_masivo_todo.Enabled = True

End If


End Sub

Private Sub DtcVendedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    strCadena = "SELECT id_venta FROM movimiento_venta WHERE id_vendedor<>'" & Me.DtcVendedor.BoundText & "' and  referencia='0' and seleccion='si' and id_doc='0099'  and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            strCadena = "UPDATE movimiento_venta SET seleccion='no' WHERE id_venta='" & rst("id_venta") & "' LIMIT 1"
            Call ConfiguraRstK(strCadena)
            rst.MoveNext
        Next i
    End If
    
    strCadena = "SELECT id_venta,ncliente,documento,total,referencia,vendedor,seleccion FROM view_listado_pendientes_ref WHERE id_vendedor='" & Me.DtcVendedor.BoundText & "' and  fecha_emision='" & KEY_FECHA & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' ORDER BY id_venta DESC"
    Call llenar_pendientes(Me.HfPendientes)
End If
End Sub
Private Sub Form_Load()
On Error GoTo salir
Dim pantalla As Single
CenterForm Me
'pantalla = Screen.Width
'pantalla1 = (pantalla - FrmVentas.Width) / 2
Me.DTPicker1.Value = KEY_FECHA
Me.DTPicker2.Value = KEY_FECHA
'pantalla3 = FrmVentas.Width - pantalla1
'Me.Left = pantalla3 - 500
Me.Top = 50

strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad  WHERE  id_personal='si' and habilitado='si' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcVendedor)
Me.DtcVendedor.BoundText = 0
  
strCadena = "SELECT id_venta FROM movimiento_venta WHERE referencia='0' and seleccion='si' and id_doc='0099'  and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "UPDATE movimiento_venta SET seleccion='no' WHERE id_venta='" & rst("id_venta") & "' LIMIT 1"
       Call ConfiguraRstK(strCadena)
       rst.MoveNext
   Next i
End If



  

strCadena = "SELECT id_venta,ncliente,documento,total,referencia,vendedor,seleccion FROM view_listado_pendientes_ref WHERE  fecha_emision='" & KEY_FECHA & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' ORDER BY id_venta DESC"
Call llenar_pendientes(Me.HfPendientes)


Exit Sub
salir:

End Sub
Public Sub llenar_pendientes(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir

Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 2800
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 2000
           Grilla.ColWidth(5) = 2000
           Grilla.ColWidth(6) = 500
           
        Next
        cabecera = "IDVENTA" & vbTab & "PROFORMA" & vbTab & "CLIENTE" & vbTab & "MONTO" & vbTab & "VENDEDOR" & vbTab & "REFERENCIA" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 1 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
        c = 6
        NumeroCampo = 6
            
        For i = 0 To rstI.RecordCount - 1
          
          If rstI("seleccion") = "si" Then
            estado = Chr(254)
          Else
            estado = Chr(168)
          End If
          
          descripcion = ""
            ndocumento = Split(rstI("documento"), ":")
            nproforma = "P:" & ndocumento(1)
            
          Fila = rstI("id_venta") & vbTab & nproforma & vbTab & rstI("ncliente") & vbTab & Format(rstI("total"), "#,##0.00") & vbTab & Mid(rstI("vendedor"), 1, 20) & vbTab & rstI("referencia") & vbTab & estado
          Grilla.AddItem Fila
          
        
        If c = NumeroCampo Then
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = c '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
                            
                        End With
        End If
        Fila = ""
          
          rstI.MoveNext
      Next i
      
      
      
      Exit Sub
salir:
    
End Sub

Private Sub HfPendientes_Click()

If Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) > 0 Then
    
    strCadena = "SELECT seleccion FROM movimiento_venta WHERE id_venta='" & Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) & "'  LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       If rst("seleccion") = "no" Then
          strCadena = "UPDATE movimiento_venta SET seleccion='si' WHERE id_venta='" & Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) & "'  LIMIT 1"
          in_char = Chr(254)
       Else
          strCadena = "UPDATE movimiento_venta SET seleccion='no' WHERE id_venta='" & Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) & "'  LIMIT 1"
          in_char = Chr(168)
       End If
       CnBd.Execute (strCadena)
       Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 6) = in_char
    End If
End If


End Sub

Private Sub HfPendientes_DblClick()

If Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) > 0 Then
    
    FrmVentas.txt_id_pendiente.Text = Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0))
    
    Call FrmVentas.get_comprobante(Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)))
    FrmVentas.timer_pendientes.Enabled = False
    
    Unload Me
    Call enabled_form(FrmVentas)
End If

End Sub
Public Sub put_factura(ByVal in_dni As String, ByVal in_nombre As String, ByVal in_direccion As String, ByVal in_monto As Single, ByVal in_operacion As String, ByVal in_forma_pago As String, ByVal in_producto As String, ByVal in_cantidad As Single, ByVal in_detalle As String)
Dim in_doc As String
Dim in_serie As String
Dim in_numero As String
Dim in_subtotal As Single
Dim in_igv As Single
Dim in_exonerado As Single
Dim in_total As Single
strCadena = "call P_nueva_venta('" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

If Len(in_dni) = 8 Then
    in_doc = "0003"
Else
    in_doc = "0001"
End If


strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='" & in_doc & "' and electronico='si' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   in_serie = rst("serie")
   strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & in_doc & "' and serie='" & rst("serie") & "' and id_alm='" & KEY_ALM & "' and  ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      in_numero = Format(Val(rst("numero")) + 1, "000000")
   Else
      in_numero = "000001"
   End If
      'inserta temporal
      
      
      strCadena = "INSERT INTO temporal_ventas(ruc,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,costo,servicio) VALUES " & _
      "('" & KEY_RUC & "','" & in_dni & "','" & KEY_ALM & "','" & in_doc & "','" & in_serie & "','" & in_numero & "','" & in_producto & "','" & in_cantidad & "'," & _
      "'" & in_monto & "','" & in_monto * in_cantidad & "','0','no','" & in_detalle & "','" & KEY_USUARIO & "','1','no')"
      CnBd.Execute (strCadena)
        
        
       
        
     

    strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,cuenta_contable,cuotas,id_usuario,id_recibo,detalle,banco,cheque,fecha,id_alm,serie_nota,numero_nota,ruc) VALUES " & _
    " ('" & in_doc & "','" & in_serie & "','" & in_numero & "','01','" & in_forma_pago & "','00001','" & in_monto * in_cantidad & "','" & in_monto * in_cantidad & "','00','','" & in_operacion & "','" & get_cuenta_contable_caja(in_forma_pago) & "','0','" & KEY_USUARIO & "','0','','-','-','" & KEY_FECHA & "','" & KEY_ALM & "','0','0','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    in_cta_cobrar = KEY_CTA_COBRAR_SERVICIO
    in_cta_ingreso = KEY_CTA_INGRESO_SERVICIO

If KEY_CON_IGV = "si" Then
    in_total = in_monto * in_cantidad
    in_valorventa = in_total / (1 + KEY_IGV)
    in_igv = in_total - in_valorventa
    in_exonerado = 0
    
Else
    in_valorventa = in_monto * in_cantidad
    in_igv = 0
    in_exonerado = in_valorventa
    in_total = in_valorventa
End If

in_documento = "BOLETA:" & in_serie & "-" & in_numero
    
    strCadena = "call p_insert_venta_cabecera_premiun_demo('" & in_doc & "','" & KEY_ALM & "','01','00001','no'," & _
    "'" & Trim(in_serie) & "','" & Trim(in_numero) & "','" & in_dni & "','" & in_nombre & "','" & in_valorventa & "','" & in_igv & "','" & in_exonerado & "','" & in_total & "','0'," & _
    "'" & in_total & "','0','" & KEY_FECHA & "','" & KEY_FECHA & "','00001','" & KEY_USUARIO & "','" & KEY_USUARIO & "','" & Val(KEY_CAMBIO_COMPRA) & "','no','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "'" & _
    ",'" & in_documento & "',CURDATE(),'T','" & Trim(in_direccion) & "','no','" & Trim(Me.txt_hash.Text) & "','" & Trim(Me.txt_sunat_key.Text) & "',' ',' ',' ',' ','" & KEY_VENTANILLA & "','01','0','-','no','" & KEY_CONTABILIDAD & "','" & in_cta_cobrar & "','" & in_cta_ingreso & "','0','0','0','0','no','" & KEY_RUC & "')"
    
    
    
    Call ConfiguraRstPP(strCadena)
    id_venta = rstPP("in_venta")
    Me.TxtIdVenta.Text = id_venta
    
   
    
    
    Call put_correlativo_venta(in_doc, in_serie, in_numero)
    
        If KEY_FACTURACION_ELECTRONICA = "si" Then
                If get_firma_online(in_doc, in_serie) = "si" Then
                   Call firma_electronica(in_doc, "no", " ", id_venta, in_numero, in_serie, in_dni, in_nombre, in_direccion)
                   Exit Sub
                End If
           End If
     End If
      
End Sub

Private Sub put_deseleccionar(ByVal in_pedido As String)

End Sub
Public Sub put_masivo_seleccionados()
Dim in_doc As String
Dim in_serie As String
Dim in_numero As String
Dim in_subtotal As Single
Dim in_igv As Single
Dim in_exonerado As Single
Dim in_total As Single
Dim codigoP As String
Dim in_item As Integer

Me.Timer_masivo_todo.Enabled = False
strCadena = "call P_nueva_venta('" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)


strCadena = "SELECT id_venta,id_doc,id_cliente,ncliente,direccion,id_vendedor FROM movimiento_venta WHERE seleccion='si' and referencia='0' and id_doc='0099' and ruc='" & KEY_RUC & "' ORDER BY id_venta ASC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
        in_dni = rst("id_cliente")
        in_nombre = rst("ncliente")
        in_direccion = rst("direccion")
        in_vendedor = rst("id_vendedor")
        in_pedido = rst("id_venta")
        If Len(in_dni) = 11 Then
            in_doc = "0001"
        Else
            in_doc = "0003"
        End If
Else
    strCadena = "SELECT id_venta,ncliente,documento,total,referencia,vendedor,seleccion FROM view_listado_pendientes_ref WHERE  fecha_emision='" & KEY_FECHA & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' ORDER BY id_venta DESC"
   Call llenar_pendientes(Me.HfPendientes)
   Exit Sub
End If


strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='" & in_doc & "' and electronico='si' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   in_serie = rst("serie")
   
   strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & in_doc & "' and serie='" & rst("serie") & "' and id_alm='" & KEY_ALM & "' and  ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      in_numero = Format(Val(rst("numero")) + 1, "000000")
   Else
      in_numero = "000001"
   End If
      
      '*****inserta temporal
      
      in_total = 0
      in_item = 0
      strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & in_pedido & "' and ruc='" & KEY_RUC & "'"
      ConfiguraRstK (strCadena)
      If rstK.RecordCount > 0 Then
         rstK.MoveFirst
         
         For i = 0 To rstK.RecordCount - 1
            codigoP = rstK("id_producto")
           If control_stock(Trim(codigoP), rstK("cantidad")) = True Then
           
           If KEY_SEGMENTACION_PRECIO = "si" Then
               in_precio = get_precio_segmentacion(codigoP, in_dni)
           Else
               in_precio = get_precio_venta_now(codigoP)
           End If
            in_item = in_item + 1
            strCadena = "INSERT INTO temporal_ventas(ruc,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,costo,servicio) VALUES " & _
            "('" & KEY_RUC & "','" & in_dni & "','" & KEY_ALM & "','" & in_doc & "','" & in_serie & "','" & in_numero & "','" & codigoP & "','" & rstK("cantidad") & "'," & _
            "'" & in_precio & "','" & in_precio * rstK("cantidad") & "','0','no','" & rstK("detalle") & "','" & KEY_USUARIO & "','1','no')"
            CnBd.Execute (strCadena)
            
            
            If KEY_BONIFICACIONES = "si" Then
                strCadena = "CALL get_idTemporalventas('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
                Call ConfiguraRstC(strCadena)
                in_idVenta = rstc(0)
                strCadena = "CALL put_bonificacion_linea('" & codigoP & "','" & in_dni & "','" & KEY_USUARIO & "','" & KEY_ALM & "','" & in_idVenta & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                Call put_verificar_bonificacion_monto(codigoP, rstK("cantidad"), in_dni, in_doc, in_serie)
                Call put_verificar_bonificacion_cruzada_v2(codigoP, rstK("cantidad"), in_dni, in_doc, in_serie)
            End If
            in_total = in_total + in_precio * rstK("cantidad")
            End If
            rstK.MoveNext
         Next i
         
      End If
      
 
     If in_item = 0 Then
        strCadena = "call ADM_venta_operacion('1','" & in_pedido & "')"
        Call ConfiguraRstAux(strCadena)
        GoTo terminar
     End If
     

    strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,cuenta_contable,cuotas,id_usuario,id_recibo,detalle,banco,cheque,fecha,id_alm,serie_nota,numero_nota,ruc) VALUES " & _
    " ('" & in_doc & "','" & in_serie & "','" & in_numero & "','01','" & Trim(Me.txtForma_pago.Text) & "','00001','" & in_total & "','" & in_total & "','00','','" & in_operacion & "','" & get_cuenta_contable_caja(Me.txtForma_pago.Text) & "','0','" & KEY_USUARIO & "','0','','-','-','" & KEY_FECHA & "','" & KEY_ALM & "','0','0','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    in_cta_cobrar = KEY_CTA_COBRAR_SERVICIO
    in_cta_ingreso = KEY_CTA_INGRESO_SERVICIO

If KEY_CON_IGV = "si" Then
    in_total = in_total
    in_valorventa = in_total / (1 + KEY_IGV)
    in_igv = in_total - in_valorventa
    in_exonerado = 0
    
Else
    in_valorventa = in_total
    in_igv = 0
    in_exonerado = in_valorventa
    in_total = in_valorventa
End If

    If in_doc = "0003" Then
       in_documento = "BOLETA:" & in_serie & "-" & in_numero
    Else
       in_documento = "FACTURA:" & in_serie & "-" & in_numero
    End If
    
    
    strCadena = "call p_insert_venta_cabecera_v15('" & in_doc & "','" & KEY_ALM & "','01','00001','no'," & _
    "'" & Trim(in_serie) & "','" & Trim(in_numero) & "','" & in_dni & "','" & in_nombre & "','" & in_valorventa & "','" & in_igv & "','" & in_exonerado & "','" & in_total & "','0'," & _
    "'" & in_total & "','0','" & KEY_FECHA & "','" & KEY_FECHA & "','00001','" & in_vendedor & "','" & KEY_USUARIO & "','" & Val(KEY_CAMBIO_COMPRA) & "','no','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "'" & _
    ",'" & in_documento & "',CURDATE(),'T','" & Trim(in_direccion) & "','no','" & Trim(Me.txt_hash.Text) & "','" & Trim(Me.txt_sunat_key.Text) & "','0','',' ',' ','" & KEY_VENTANILLA & "','01','0','-','no','" & KEY_CONTABILIDAD & "','" & in_cta_cobrar & "','" & in_cta_ingreso & "','0','0','0','0','no','0','si',0,'" & KEY_RUC & "')"
    Call ConfiguraRstPP(strCadena)
    
    'strCadena = "call p_insert_venta_cabecera_v15('" & in_doc & "','" & KEY_ALM & "','01','00001','no'," & _
                "'" & Trim(in_serie) & "','" & Trim(in_numero) & "','" & in_dni & "','" & in_nombre & "','" & in_valorventa & "','" & in_igv & "','" & in_exonerado & "','" & in_total & "','0'," & _
                "'" & in_total & "','0','" & KEY_FECHA & "','" & KEY_FECHA & "','00001','" & in_vendedor & "','" & KEY_USUARIO & "','" & Val(KEY_CAMBIO_COMPRA) & "','no','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "'" & _
                ",'" & in_documento & "',CURDATE(),'T','" & Trim(in_direccion) & "','no','" & Trim(Me.txt_hash.Text) & "','" & Trim(Me.txt_sunat_key.Text) & "','" & Trim(Me.DtcTipoNota.BoundText) & "','" & Trim(Me.txtmotivo_nota.Text) & "','" & id_guia & "','" & in_guia & "','" & KEY_VENTANILLA & "'," & _
                " '" & Trim(Me.txt_tipo.Text) & "','" & in_seguro & "','" & Trim(Me.txtObservacion.Text) & "','" & Trim(Me.txteditable.Text) & "','" & KEY_CONTABILIDAD & "','" & in_cta_cobrar & "','" & in_cta_ingreso & "','" & Val(Me.TxtDescuento_global.Text) & "','" & Val(Me.TxtCuotas.Text) & "','" & in_interes & "','" & Val(Me.txtid_venta_ref.Text) & "','" & in_diferida & "','" & Val(Me.txt_id_pendiente.Text) & "','" & KEY_SIN_EFECTO_CAJA & "','" & Val(Me.lblicbper.Caption) & "','" & KEY_RUC & "')"
              
              
    
    
    id_venta = rstPP("in_venta")
    
    
    
    Me.TxtIdVenta.Text = id_venta
    Me.txtComprobante.Text = in_documento
     
    strCadena = "UPDATE movimiento_venta SET referencia='" & Trim(in_documento) & "' WHERE id_venta='" & Val(in_pedido) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    CnBd.Execute (strCadena)
    
    Call put_correlativo_venta(in_doc, in_serie, in_numero)
    
        If KEY_FACTURACION_ELECTRONICA = "si" Then
                If get_firma_online(in_doc, in_serie) = "si" Then
                   Call firma_electronica_masiva(in_doc, "no", " ", id_venta, in_numero, in_serie, in_dni, in_nombre, in_direccion)
                   Exit Sub
                End If
           End If
     End If
      
terminar:
Me.Timer_masivo_todo.Enabled = True
End Sub

Private Function control_stock(ByVal in_producto As String, ByVal in_cantidad As Double) As Boolean
On Error GoTo salir


strCadena = "SELECT stock FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
    If rstAux("stock") < in_cantidad And KEY_MOVIMIENTO_SIN_STOCK = "no" Then
                
            MsgBox "PRODUCTO NO CUENTA CON STOCK." + Chr(13) + Chr(13) + in_producto + Space(1) + get_producto(in_producto) + Space(2) + Chr(13) + Chr(13) + "STOCK ACTUAL : " + str(rstAux("stock")) + Chr(13) + "TOTAL PEDIDO :" + str(in_cantidad) + Chr(13) + Chr(13) + "Consulte con el Area de Almacen.", vbInformation, KEY_EMPRESA
            control_stock = False
    Else
            control_stock = True
    End If
End If

Exit Function
salir:


End Function

Private Function firma_electronica(ByVal in_doc As String, ByVal in_extranjero As String, ByVal in_observacion As String, ByVal in_venta As String, ByVal numero As String, ByVal in_serie As String, ByVal in_dni As String, ByVal in_alumno As String, ByVal in_direccion As String)
Dim in_moneda As String
Call disabled_form(Me)
FrmLoad_web_service.Show
FrmLoad_web_service.nom_prcedimiento = "procesar_firma_electronica_local"
Set FrmLoad_web_service.FormPadre = Me

Select Case in_doc
    Case "0003"
         in_tipo_doc = "1"
         If Trim(in_extranjero) = "si" Then
            in_tipo_doc = "4"
         End If
    Case "0001"
        in_tipo_doc = "6"
         If Trim(in_extranjero) = "si" Then
            in_tipo_doc = "4"
         End If
         
    Case "0002"
        in_tipo_doc = "1"
End Select

    id_motivo_nota = ""
    motivo_nota = ""
    in_serie_afectado = ""
    in_numero_afectado = ""
    in_observacion = Replace(Trim(in_observacion), "'", " ")
    
    id_motivo_nota = ""
    motivo_nota = ""
    in_serie_afectado = ""
    in_numero_afectado = ""
    in_moneda = "PEN"



If get_comprobante_produccion(in_doc, in_serie) = "si" Then
    in_numero = Trim(numero)
    If KEY_SERVIDOR_KEYFACIL = "si" Then
        Call FrmLoad_web_service.crear_json_facturacion_electronica("https://api.vitekey.com/keyfact/utils/erp-invoice?password=vitekey2018&company_id=" & KEY_TOKEN_CLOUD & "", "POST", json_facturacion_electronica_firmar_id_venta_keyfacil(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, "", in_serie_afectado, in_numero_afectado, in_moneda, in_observacion, "no", ""), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
    Else
        If KEY_SERVIDOR_CLOUD = "si" Then
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, "", in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
        Else
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_LOCAL & "'}")
        End If
    End If


Else
    in_numero = Trim(numero)
    
    
    
    If KEY_SERVIDOR_KEYFACIL = "si" Then
        Call FrmLoad_web_service.crear_json_facturacion_electronica("https://api.vitekey.com/keyfact/utils/erp-invoice?password=vitekey2018&company_id=" & KEY_TOKEN_CLOUD & "", "POST", json_facturacion_electronica_firmar_id_venta_keyfacil(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, "", in_serie_afectado, in_numero_afectado, in_moneda, in_observacion, "no", ""), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
    Else
        If KEY_SERVIDOR_CLOUD = "si" Then
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, "", in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
        Else
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_LOCAL & "'}")
        End If
    End If
    End If



End Function

Public Sub procesar_firma_electronica_local(ByVal strHtml As String)
On Error GoTo procesar_nuevamente
Dim in_error As Boolean
Dim in_hash As String
Dim in_numero() As String
Dim json_r As Object
Me.txt_hash.Text = ""
in_hash = ""
Set json_r = JSON.parse(strHtml)
in_error = json_r.Item("error")
If in_error = True Then

Else
     
     If KEY_SERVIDOR_KEYFACIL = "si" Then
        in_hash = Trim(json_r.Item("response").Item("id"))
        in_key = Trim(json_r.Item("response").Item("id"))
        'get_numero = Trim(json_r.Item("response").Item("numero"))
     Else
        in_hash = Trim(json_r.Item("response").Item("digest_value"))
        in_key = Trim(json_r.Item("response").Item("key"))
        get_numero = Trim(json_r.Item("response").Item("numero"))
     End If
     
     Me.txt_hash.Text = Trim(in_hash)
     Me.txt_sunat_key.Text = Trim(in_key)
     
   
     
     strCadena = "UPDATE movimiento_venta SET sunat_key='" & Trim(in_key) & "',sunat_hash='" & Trim(in_hash) & "' WHERE id_venta='" & Val(Me.TxtIdVenta.Text) & "'"
     CnBd.Execute (strCadena)
     
     Me.txt_sunat_key.Text = ""
     Me.txt_hash.Text = ""
     Me.timer_masivo.Enabled = True
     Me.Enabled = True
     
     
     Exit Sub
     
     'Call procesar_comprobante
     
End If
Exit Sub
procesar_nuevamente:
MsgBox "SE PRESENTO UN PROBLEMA CON EL INTERNET" + Chr(13) + Chr(13) + "INTENTENTALO NUEVAMENTE.", vbInformation, KEY_USUAURIO
Me.Enabled = True

End Sub
Private Function firma_electronica_masiva(ByVal in_doc As String, ByVal in_extranjero As String, ByVal in_observacion As String, ByVal in_venta As String, ByVal numero As String, ByVal in_serie As String, ByVal in_dni As String, ByVal in_alumno As String, ByVal in_direccion As String)
Dim in_moneda As String
Call disabled_form(Me)
FrmLoad_web_service.Show
FrmLoad_web_service.nom_prcedimiento = "procesar_firma_electronica_masiva"
Set FrmLoad_web_service.FormPadre = Me

Select Case in_doc
    Case "0003"
         in_tipo_doc = "1"
         If Trim(in_extranjero) = "si" Then
            in_tipo_doc = "4"
         End If
    Case "0001"
        in_tipo_doc = "6"
         If Trim(in_extranjero) = "si" Then
            in_tipo_doc = "4"
         End If
         
    Case "0002"
        in_tipo_doc = "1"
End Select

    id_motivo_nota = ""
    motivo_nota = ""
    in_serie_afectado = ""
    in_numero_afectado = ""
    in_observacion = Replace(Trim(in_observacion), "'", " ")
    
    id_motivo_nota = ""
    motivo_nota = ""
    in_serie_afectado = ""
    in_numero_afectado = ""
    in_moneda = "PEN"



If get_comprobante_produccion(in_doc, in_serie) = "si" Then
    in_numero = Trim(numero)
    If KEY_SERVIDOR_KEYFACIL = "si" Then
        Call FrmLoad_web_service.crear_json_facturacion_electronica("https://api.vitekey.com/keyfact/utils/erp-invoice?password=vitekey2018&company_id=" & KEY_TOKEN_CLOUD & "", "POST", json_facturacion_electronica_firmar_id_venta_keyfacil(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, "", in_serie_afectado, in_numero_afectado, in_moneda, in_observacion, "no", ""), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
    Else
        If KEY_SERVIDOR_CLOUD = "si" Then
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, "", in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
        Else
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_LOCAL & "'}")
        End If
    End If


Else
    in_numero = Trim(numero)
    
    
    
    If KEY_SERVIDOR_KEYFACIL = "si" Then
        Call FrmLoad_web_service.crear_json_facturacion_electronica("https://api.vitekey.com/keyfact/utils/erp-invoice?password=vitekey2018&company_id=" & KEY_TOKEN_CLOUD & "", "POST", json_facturacion_electronica_firmar_id_venta_keyfacil(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, "", in_serie_afectado, in_numero_afectado, in_moneda, in_observacion, "no", ""), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
    Else
        If KEY_SERVIDOR_CLOUD = "si" Then
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, "", in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
        Else
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(in_venta), Format(Val(in_doc), "00"), Trim(in_serie), in_numero, KEY_FECHA, Trim(in_dni), Trim(in_alumno), Trim(in_direccion), in_tipo_doc, 0, KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_LOCAL & "'}")
        End If
    End If
    End If



End Function


Public Sub procesar_firma_electronica_masiva(ByVal strHtml As String)
On Error GoTo procesar_nuevamente
Dim in_error As Boolean
Dim in_hash As String
Dim in_numero() As String
Dim json_r As Object
Me.txt_hash.Text = ""
in_hash = ""
Set json_r = JSON.parse(strHtml)
in_error = json_r.Item("error")
If in_error = True Then

Else
     
     If KEY_SERVIDOR_KEYFACIL = "si" Then
        in_hash = Trim(json_r.Item("response").Item("id"))
        in_key = Trim(json_r.Item("response").Item("id"))
        'get_numero = Trim(json_r.Item("response").Item("numero"))
     Else
        in_hash = Trim(json_r.Item("response").Item("digest_value"))
        in_key = Trim(json_r.Item("response").Item("key"))
        get_numero = Trim(json_r.Item("response").Item("numero"))
     End If
     
     Me.txt_hash.Text = Trim(in_hash)
     Me.txt_sunat_key.Text = Trim(in_key)
     
   
     
     strCadena = "UPDATE movimiento_venta SET sunat_key='" & Trim(in_key) & "',sunat_hash='" & Trim(in_hash) & "' WHERE id_venta='" & Val(Me.TxtIdVenta.Text) & "'"
     CnBd.Execute (strCadena)
     
     Me.txt_sunat_key.Text = ""
     Me.txt_hash.Text = ""
     
    
     strCadena = "SELECT id_doc,serie,numero,id_tipo_factura FROM movimiento_venta WHERE id_venta='" & Val(Me.TxtIdVenta.Text) & "' LIMIT 1"
     Call ConfiguraRst(strCadena)
     If rst.RecordCount > 0 Then
           ndoc = rst("id_doc")
           nser = rst("serie")
           nnumero = rst("numero")
           ntipofactura = rst("id_tipo_factura")
        Call Orden_Impresion(ndoc, nser, nnumero, ntipofactura, Val(Me.TxtIdVenta.Text))
        Call Orden_Impresion(ndoc, nser, nnumero, ntipofactura, Val(Me.TxtIdVenta.Text))
      End If
     
     Me.Timer_masivo_todo.Enabled = True
     Me.Enabled = True
     
     
     Exit Sub
     
     'Call procesar_comprobante
     
End If
Exit Sub
procesar_nuevamente:
MsgBox "SE PRESENTO UN PROBLEMA CON EL INTERNET" + Chr(13) + Chr(13) + "INTENTENTALO NUEVAMENTE.", vbInformation, KEY_USUAURIO
Me.Enabled = True

End Sub



Private Sub timer_masivo_Timer()
Call facturar_masivo
End Sub

Private Sub Timer_masivo_todo_Timer()

Call put_masivo_seleccionados

End Sub

Private Sub txtid_productomasivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Procedencia = Selecionar
   FrmProducto.Show
End If
End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmProducto 
   BorderStyle     =   0  'None
   Caption         =   "E"
   ClientHeight    =   9240
   ClientLeft      =   915
   ClientTop       =   600
   ClientWidth     =   20145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame framemayor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "QUIEBRE DE STOCK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1455
      Left            =   16080
      TabIndex        =   14
      Top             =   5880
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox txtprecioventa 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         Text            =   "20"
         Top             =   480
         Width           =   1695
      End
      Begin VitekeySoft.ChameleonBtn cmdReport 
         Height          =   405
         Left            =   1200
         TabIndex        =   51
         ToolTipText     =   "Reporte"
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   714
         BTYPE           =   3
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmProducto.frx":0000
         PICN            =   "FrmProducto.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   3240
         Picture         =   "FrmProducto.frx":00A9
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MARGEN :"
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
         Left            =   255
         TabIndex        =   16
         Top             =   480
         Width           =   675
      End
   End
   Begin VB.Frame frmFoto 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   14880
      TabIndex        =   66
      Top             =   4320
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Image Image1 
         Height          =   240
         Left            =   4800
         Picture         =   "FrmProducto.frx":2F4D
         Top             =   120
         Width           =   240
      End
      Begin VB.Image img_foto 
         Height          =   4650
         Left            =   45
         Stretch         =   -1  'True
         Top             =   120
         Width           =   5055
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   14160
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfGrupoEmpresarial 
      Height          =   1575
      Left            =   14880
      TabIndex        =   65
      Top             =   4440
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2778
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
   Begin VB.Frame frm_ubicacion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MATRIZ UBICACION FISICA"
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
      Height          =   3135
      Left            =   14880
      TabIndex        =   17
      Top             =   4320
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtcodigo_barra 
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   60
         Top             =   2640
         Width           =   1935
      End
      Begin VB.TextBox txtFormafarmacologica 
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   59
         Top             =   2200
         Width           =   3735
      End
      Begin VB.TextBox TxtSector 
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   22
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtPiso 
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   21
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox TxtAndamio 
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   20
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_x 
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   19
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Txt_y 
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
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   18
         Top             =   1800
         Width           =   975
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   315
         Left            =   1200
         TabIndex        =   23
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632319
         ForeColor       =   8388608
         Text            =   ""
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
      Begin VitekeySoft.ChameleonBtn cmdProcesar 
         Height          =   360
         Left            =   3600
         TabIndex        =   61
         ToolTipText     =   "Stock Bajo"
         Top             =   2640
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   635
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmProducto.frx":5DF1
         PICN            =   "FrmProducto.frx":5E0D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COD.BARRA :"
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
         Index           =   2
         Left            =   75
         TabIndex        =   58
         Top             =   2640
         Width           =   1005
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.FARMA :"
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
         Index           =   1
         Left            =   285
         TabIndex        =   57
         Top             =   2200
         Width           =   795
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ALMACEN :"
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
         Left            =   225
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SECTOR :"
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
         Left            =   375
         TabIndex        =   27
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PISO :"
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
         Left            =   615
         TabIndex        =   26
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ANDAMIO :"
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
         Left            =   195
         TabIndex        =   25
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CASILLERO :"
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
         Left            =   165
         TabIndex        =   24
         Top             =   1800
         Width           =   915
      End
   End
   Begin VB.Frame frmbusquedafarmacia 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   760
      Left            =   6840
      TabIndex        =   52
      Top             =   8380
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox txtprincipioActivo 
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
         Left            =   930
         TabIndex        =   63
         Top             =   50
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker Dtpvencimiento 
         Height          =   300
         Left            =   1680
         TabIndex        =   62
         Top             =   405
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Format          =   171966465
         CurrentDate     =   43577
      End
      Begin VB.TextBox txtforma_farmacologica 
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
         Left            =   4680
         TabIndex        =   54
         Top             =   50
         Width           =   2055
      End
      Begin VitekeySoft.ChameleonBtn cmdvencimiento 
         Height          =   360
         Left            =   3200
         TabIndex        =   55
         ToolTipText     =   "Stock Bajo"
         Top             =   375
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   635
         BTYPE           =   5
         TX              =   "VENCIMIENTO"
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
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmProducto.frx":8161
         PICN            =   "FrmProducto.frx":817D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P.ACTIVO :"
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
         Left            =   105
         TabIndex        =   64
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIAS X VENCERSE :"
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
         Left            =   90
         TabIndex        =   56
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.FARMACOLOGICA :"
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
         Left            =   3150
         TabIndex        =   53
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame frmCisterna 
      BackColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   14880
      TabIndex        =   44
      Top             =   3960
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Label lblmax 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "100 GL"
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
         Left            =   100
         TabIndex        =   50
         Top             =   740
         Width           =   585
      End
      Begin VB.Label lblmedio 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "100 GL"
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
         Left            =   100
         TabIndex        =   49
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label lblmin 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0 GL"
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
         Left            =   100
         TabIndex        =   48
         Top             =   4560
         Width           =   585
      End
      Begin VB.Label lblCisterna 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "CODIGO :"
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
         Left            =   3840
         TabIndex        =   47
         Top             =   360
         Width           =   645
      End
      Begin VB.Label lblCisterna 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "CODIGO :"
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
         Left            =   2400
         TabIndex        =   46
         Top             =   360
         Width           =   645
      End
      Begin VB.Shape Cisterna_vacia 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3855
         Index           =   2
         Left            =   3600
         Top             =   840
         Width           =   1335
      End
      Begin VB.Shape Cisterna_llena 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   4095
         Index           =   2
         Left            =   3600
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   1335
      End
      Begin VB.Shape Cisterna_vacia 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3855
         Index           =   1
         Left            =   2160
         Top             =   840
         Width           =   1335
      End
      Begin VB.Shape Cisterna_llena 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   4095
         Index           =   1
         Left            =   2160
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   1335
      End
      Begin VB.Shape Cisterna_vacia 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   3855
         Index           =   0
         Left            =   720
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblCisterna 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "CODIGO :"
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
         Index           =   0
         Left            =   1080
         TabIndex        =   45
         Top             =   360
         Width           =   645
      End
      Begin VB.Shape Cisterna_llena 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000040C0&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   4095
         Index           =   0
         Left            =   720
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.TextBox txtCuenta_contable 
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
      Left            =   1275
      TabIndex        =   42
      Top             =   8760
      Width           =   1335
   End
   Begin VB.Frame frmcompatible 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COMPATIBLE CON"
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
      Height          =   3975
      Left            =   6000
      TabIndex        =   38
      Top             =   2280
      Visible         =   0   'False
      Width           =   7695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCompatible 
         Height          =   3135
         Left            =   240
         TabIndex        =   39
         Top             =   600
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5530
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
      Begin VitekeySoft.ChameleonBtn cmdCerrar 
         Height          =   260
         Left            =   7320
         TabIndex        =   40
         ToolTipText     =   "Reporte"
         Top             =   240
         Width           =   260
         _ExtentX        =   450
         _ExtentY        =   450
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmProducto.frx":A4D1
         PICN            =   "FrmProducto.frx":A4ED
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.TextBox txtMarca 
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
      Left            =   12480
      TabIndex        =   36
      Top             =   8820
      Width           =   975
   End
   Begin VB.CheckBox chkmarca 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "MARCA                 :"
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
      Height          =   240
      Left            =   7320
      TabIndex        =   35
      Top             =   8860
      Width           =   1455
   End
   Begin VitekeySoft.ChameleonBtn cmdexit 
      Height          =   855
      Left            =   13800
      TabIndex        =   33
      Top             =   7425
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
      MICON           =   "FrmProducto.frx":D3A1
      PICN            =   "FrmProducto.frx":D3BD
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
      Left            =   13800
      TabIndex        =   32
      Top             =   2010
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
      MICON           =   "FrmProducto.frx":D7AD
      PICN            =   "FrmProducto.frx":D7C9
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
      Left            =   13800
      TabIndex        =   31
      Top             =   1120
      Width           =   975
      _ExtentX        =   1720
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmProducto.frx":FC13
      PICN            =   "FrmProducto.frx":FC2F
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
      Left            =   13800
      TabIndex        =   30
      Top             =   240
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
      MICON           =   "FrmProducto.frx":FF49
      PICN            =   "FrmProducto.frx":FF65
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkLinea 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "CLASIFICACION :"
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
      Height          =   240
      Left            =   7320
      TabIndex        =   13
      Top             =   8490
      Width           =   1455
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   12480
      TabIndex        =   12
      Top             =   8420
      Width           =   975
   End
   Begin VitekeySoft.ChameleonBtn cmdStockBajo 
      Height          =   405
      Left            =   13800
      TabIndex        =   11
      ToolTipText     =   "Stock Bajo"
      Top             =   5640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      BTYPE           =   5
      TX              =   "ST(-)"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmProducto.frx":103B7
      PICN            =   "FrmProducto.frx":103D3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton CmdSiguiente 
      Caption         =   "SIGUIENTE >>>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   18480
      TabIndex        =   9
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton CmdAnterior 
      Caption         =   "<<< ANTERIOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14880
      TabIndex        =   8
      Top             =   8880
      Width           =   1575
   End
   Begin VB.TextBox TxtCod 
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
      Left            =   1275
      TabIndex        =   3
      Top             =   8400
      Width           =   1335
   End
   Begin VB.TextBox TxtProducto 
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
      Left            =   3885
      TabIndex        =   1
      Top             =   8520
      Width           =   2775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   14208
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgAlmacen 
      Height          =   3135
      Left            =   14880
      TabIndex        =   6
      Top             =   1200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5530
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
   Begin MSDataListLib.DataCombo DtcLinea 
      Height          =   330
      Left            =   8880
      TabIndex        =   7
      Top             =   8415
      Width           =   3495
      _ExtentX        =   6165
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
   Begin VitekeySoft.ChameleonBtn cmdlogistica 
      Height          =   525
      Left            =   13800
      TabIndex        =   29
      ToolTipText     =   "Ubicacion"
      Top             =   6120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   926
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmProducto.frx":135DE
      PICN            =   "FrmProducto.frx":135FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcMarca 
      Height          =   330
      Left            =   8880
      TabIndex        =   34
      Top             =   8820
      Width           =   3495
      _ExtentX        =   6165
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
   Begin VitekeySoft.ChameleonBtn cmdcompatibility 
      Height          =   855
      Left            =   13800
      TabIndex        =   37
      Top             =   2890
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "COMPATIBLE"
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
      MICON           =   "FrmProducto.frx":15D17
      PICN            =   "FrmProducto.frx":15D33
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCaracteristicas 
      Height          =   855
      Left            =   13800
      TabIndex        =   41
      Top             =   3780
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "CARACTERIS"
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
      MICON           =   "FrmProducto.frx":19A74
      PICN            =   "FrmProducto.frx":19A90
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdFoto 
      Height          =   855
      Left            =   13800
      TabIndex        =   67
      Top             =   4680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "IMAGEN"
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
      MICON           =   "FrmProducto.frx":1CD66
      PICN            =   "FrmProducto.frx":1CD82
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CTA CONTABLE :"
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
      Left            =   120
      TabIndex        =   43
      Top             =   8790
      Width           =   1125
   End
   Begin VB.Label lblFotos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION :"
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
      Left            =   16800
      TabIndex        =   10
      Top             =   8880
      Width           =   1155
   End
   Begin VB.Label LblProducto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   920
      Left            =   14880
      TabIndex        =   5
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO :"
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
      Left            =   615
      TabIndex        =   4
      Top             =   8430
      Width           =   645
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION :"
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
      Left            =   2775
      TabIndex        =   2
      Top             =   8550
      Width           =   1005
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   120
      Top             =   8340
      Width           =   13575
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
Attribute VB_Name = "FrmProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Dim strLinea As Boolean
Dim strMostrarTodos As Boolean
Dim KEY_VENTA_MONEDA As String


Private Sub ChameleonBtn1_Click()
    
End Sub

Private Sub ChameleonBtn2_Click()

End Sub

Private Sub CmdAnterior_Click()
On Error GoTo salir
If rstI.EOF = True Or rstI.BOF = True Then
    If rstI.RecordCount < 1 Then
        Exit Sub
    End If
    rstI.MoveFirst
Else
    rstI.MovePrevious
    If rstI.EOF = True Or rstI.BOF = True Then
        rstI.MoveLast
    End If
End If
If IsNull(rstI("foto")) = False And Len(rstI("foto")) > 5 Then
    If VerificarArchivo(App.Path & "\archivos\" & KEY_RUC & "\" & rstI("foto")) = True Then
        'Me.Image1.Visible = True
        Image1.Picture = LoadPicture(App.Path + "\archivos\" + KEY_RUC + "\" + Trim(rstI("foto")))
    Else
        'Me.Image1 = Nothing
    End If
End If
Exit Sub
salir:

End Sub

Private Sub CmdLinea_Click()
End Sub

Private Sub cmdCaracteristicas_Click()
Me.frmcompatible.Caption = "CARACTERISTICAS"
Call load_caracteristicas(Me.HfCompatible, Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))
frmcompatible.Visible = True

End Sub

Private Sub cmdCerrar_Click()
Me.frmcompatible.Visible = False
End Sub

Private Sub cmdcompatibility_Click()

Me.frmcompatible.Caption = "COMPATIBILIDAD"
Call load_compatibility(Me.HfCompatible, Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))
frmcompatible.Visible = True

End Sub

Private Sub cmddelete_Click()
 If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Procedencia = Eliminar
        Call disabled_form(Me)
        frmsegurity.Show
        Exit Sub
        
          
          
            Call ActualizarProd
          '  Call ActualizarAlm
        End If
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub CmdFoto_Click()
On Error GoTo salir
Dim in_ruta As String



strCadena = "SELECT imagen FROM producto WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
   
   If Len(rstIN("imagen")) > 10 Then
      in_ruta = rstIN("imagen")
   Else
      in_ruta = ""
      Me.frmFoto.Visible = False
      Exit Sub
   End If
   
End If

With Inet1
    .AccessType = icUseDefault
    .Url = in_ruta
    .Execute , "GET"
End With
Exit Sub
salir:
Exit Sub

End Sub

Private Sub cmdlogistica_Click()
'strCadena = "SELECT * FROM almacen_producto a WHERE id_alm='" & KEY_ALM & "' and  precio_mayor>0 and  ruc='" & KEY_RUC & "'"
'Call ConfiguraRst(strCadena)
'If rst.RecordCount > 0 Then
 '   rst.MoveFirst
  '  For i = 0 To rst.RecordCount - 1
   '     strCadena = "SELECT * FROM almacen_producto_precio WHERE id_alm='" & KEY_ALM & "' and  id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
    '    Call ConfiguraRstT(strCadena)
     '   If rstT.RecordCount < 1 Then
      '      strCadena = "INSERT INTO almacen_producto_precio(id_producto,id_alm,precio,cant_ini,cant_fin,ruc)VALUES('" & rst("id_producto") & "','" & KEY_ALM & "','" & Val(rst("precio_mayor")) & "','1','1','" & KEY_RUC & "')"
       '     CnBd.Execute (strCadena)

      '  End If
      '  rst.MoveNext
      '  DoEvents
   ' Next i
'End If
   
   If Me.frm_ubicacion.Visible = True Then
      Me.frm_ubicacion.Visible = False
   Else
    Me.frm_ubicacion.Visible = True
   End If
End Sub


Private Sub cmdModificar_Click()

End Sub

Private Sub cmdNuevo_Click()
      Procedencia = nuevo
      FrmDetalleProducto.Show
      FrmDetalleProducto.DtcColor.BoundText = "0000"
End Sub

Private Sub cmdProcesar_Click()

strCadena = "UPDATE almacen_producto SET sector='" & Trim(Me.TxtSector.Text) & "',piso='" & Me.TxtPiso.Text & "',andamio='" & Me.TxtAndamio.Text & "',casillero_x='" & Me.txt_x.Text & "',casillero_y='" & Me.Txt_y.Text & "' WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "' "
CnBd.Execute (strCadena)

strCadena = "UPDATE producto SET codigo_barra='" & Trim(Me.txtcodigo_barra.Text) & "', forma_farmacologica='" & Trim(txtFormafarmacologica.Text) & "'  WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


'Call agrega_barra(Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))






End Sub
Private Sub agrega_barra(ByVal codigo As String)
If Trim(Me.txtcodigo_barra.Text) <> "" Then
    strCadena = "SELECT * FROM producto_barras WHERE cod_barra='" & Trim(Me.txtcodigo_barra.Text) & "' AND id_producto='" & Trim(codigo) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        MsgBox "Codigo de barras ya registrado", vbInformation, KEY_EMPRESA
    Else
        strCadena = "INSERT INTO producto_barras VALUES('" & Trim(codigo) & "','" & Trim(Me.txtcodigo_barra.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
    End If
    
End If
End Sub

Private Sub cmdReport_Click()
If Me.chkLinea.Value = 1 Then
    strCadena = "SELECT id_producto,nombre_prod,linea,modelo,color,unidad,stock,precio_compra,precio_venta FROM view_producto WHERE stock<=stock_minimo and  ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND id_linea='" & Trim(Me.DtcLinea.BoundText) & "' ORDER BY nombre_prod"
Else
    strCadena = "SELECT id_producto,nombre_prod,linea,modelo,color,unidad,stock,precio_compra,precio_venta FROM view_producto WHERE stock<=stock_minimo and ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY id_linea,nombre_prod"
End If
    Call ConfiguraRst(strCadena)
    Ans = ShowMultiReport(rst, "RptStockMinimo", , App.Path + "\Reportes\")

End Sub

Private Sub cmdsiguiente_Click()
On Error GoTo salir
If rstI.EOF = True Or rstI.BOF = True Then
    If rstI.RecordCount < 1 Then
        Exit Sub
    End If
    rstI.MoveFirst
Else
    rstI.MoveNext
    If rstI.EOF = True Or rstI.BOF = True Then
        rstI.MoveFirst
    End If
End If
If IsNull(rstI("foto")) = False And Len(rstI("foto")) > 5 Then
    If VerificarArchivo(App.Path & "\archivos\" & KEY_RUC & "\" & rstI("foto")) = True Then
        'Me.Image1.Visible = True
        Image1.Picture = LoadPicture(App.Path + "\archivos\" + KEY_RUC + "\" + Trim(rstI("foto")))
    Else
        'Me.Image1 = Nothing
    End If
End If

Exit Sub
salir:


End Sub



Private Sub cmdStockBajo_Click()


If strLinea = False Then
 If KEY_SKFACTURA = "no" Then
    strCadena = "SELECT * FROM view_producto WHERE stock<=stock_minimo and  ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY nombre_prod LIMIT 100 "
  Else
    strCadena = "SELECT * FROM view_producto WHERE stock<=stock_minimo and  ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY nombre_prod LIMIT 100"
  End If
  
  
  
  If KEY_SKFACTURA = "no" Then
        If KEY_STOCK_GLOBAL = "no" Then
            Call llenarGrid(Me.HfdGrilla, strCadena)
        Else
            Call llenarGrid_stock(Me.HfdGrilla, strCadena)
        End If
        Exit Sub
    Else
        Call LlenarGrid_Factura(Me.HfdGrilla, strCadena)
        Exit Sub
    End If

End If





If strLinea = True Then
    If KEY_SKFACTURA = "no" Then
        strCadena = "SELECT * FROM view_producto WHERE stock<=stock_minimo and  ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND id_linea = '" & Trim(Me.DtcLinea.BoundText) & "' ORDER BY nombre_prod LIMIT 0,50 "
    Else
        strCadena = "SELECT * FROM view_producto WHERE stock<=stock_minimo and  ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND id_linea = '" & Trim(Me.DtcLinea.BoundText) & "' ORDER BY nombre_prod LIMIT 0,50 "
    End If
    
    
    
    If KEY_SKFACTURA = "no" Then
        If KEY_STOCK_GLOBAL = "no" Then
            Call llenarGrid(Me.HfdGrilla, strCadena)
        Else
            Call llenarGrid_stock(Me.HfdGrilla, strCadena)
        End If
        Exit Sub
    Else
        Call LlenarGrid_Factura(Me.HfdGrilla, strCadena)
        Exit Sub
    End If
    
    
End If
End Sub

Private Sub cmdupdate_Click()
Procedencia = modificar
FrmDetalleProducto.Show

End Sub

Private Sub cmdvencimiento_Click()

'strCadena = "SELECT * FROM view_producto_farmacia WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  vencimiento <='" & Format(Me.DtpVencimiento.Value, "YYYY-mm-dd") & "' ORDER BY vencimiento"
'Call llenarGrid_farmacia(Me.HfdGrilla, strCadena)



 strCadena = "SELECT id_producto,nombre_prod,linea,vencimiento,color,unidad,stock,stock_minimo,precio_venta FROM view_producto_farmacia WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  vencimiento <='" & Format(Me.DtpVencimiento.Value, "YYYY-mm-dd") & "' ORDER BY vencimiento"
 Call ConfiguraRst(strCadena)
 Ans = ShowMultiReport(rst, "RptStockMinimo", , App.Path + "\Reportes\")


End Sub

Private Sub Command1_Click()

End Sub

Private Sub Image1_Click()
Me.frmFoto.Visible = False
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
On Error GoTo salir
Dim vtData As Variant, nomArchivo As String
Dim bDone As Boolean, tempArray() As Byte
nomArchivo = Right(Inet1.Url, Len(Inet1.Url) - InStrRev(Inet1.Url, "/"))

Select Case State
    Case icResponseCompleted
         bDone = False
         fileSize = Inet1.GetHeader("Content-length")
         contentType = Inet1.GetHeader("Content-type")
         
         Open App.Path & "\" & nomArchivo For Binary As #1
         vtData = Inet1.GetChunk(1024, icByteArray)

         DoEvents

         If Len(vtData) = 0 Then
            bDone = True
         End If
           
    Do While Not bDone

       tempArray = vtData

       Put #1, , tempArray

      

       vtData = Inet1.GetChunk(1024, icByteArray)
       DoEvents

       If Len(vtData) = 0 Then
          bDone = True
       End If
    Loop

    Close #1

'Carga la imagen
Me.frmFoto.Visible = True
Me.img_foto.Picture = LoadPicture(App.Path & "\" & nomArchivo)
'Image1.Picture = LoadPicture(App.Path & "\" & nomArchivo)

If Check1 Then Kill App.Path & "\" & nomArchivo



End Select
Exit Sub
salir:


End Sub


Private Sub DtcAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcAlmacen.BoundText <> KEY_ALM Then
    strCadena = "SELECT * FROM view_producto_almacen WHERE id_alm ='" & Me.DtcAlmacen.BoundText & "' and id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
                Me.TxtSector.Text = rst("sector")
                    Me.TxtPiso.Text = rst("piso")
                    Me.TxtAndamio.Text = rst("andamio")
                    Me.txt_x.Text = rst("casillero_x")
                    Me.Txt_y.Text = rst("casillero_y")
                    Me.txtcodigo_barra.Text = rst("codigo_barra")
                    Me.txtFormafarmacologica.Text = rst("forma_farmacologica")
    End If
End If
End If
End Sub

Private Sub DtcLinea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      If Me.chkLinea.Value = 1 Then
        strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'  and nombre_prod LIKE '%" & Trim(Me.Txtproducto.Text) & "%' and  id_linea ='" & Trim(Me.DtcLinea.BoundText) & "' ORDER BY nombre_prod "
      Else
        strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' and  id_linea ='" & Trim(Me.DtcLinea.BoundText) & "' ORDER BY nombre_prod "
      End If
      
      Call llenarGrid(Me.HfdGrilla, strCadena)
    
End If
End Sub

Private Sub DtcMarca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Me.chkmarca.Value = 1 Then
    strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'   and nombre_prod LIKE '%" & Trim(Me.Txtproducto.Text) & "%' and  marca LIKE '%" & Trim(Me.DtcMarca.Text) & "%' ORDER BY nombre_prod "
Else
    strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND    id_marca ='" & Trim(Me.DtcMarca.Text) & "' ORDER BY nombre_prod"
End If
Call llenarGrid(Me.HfdGrilla, strCadena)
    End If
End Sub

Private Sub Form_Activate()
Me.Txtproducto.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 27) Then
    Unload Me
End If
If Shift = 2 And KeyCode = Asc("R") Then
    If MsgBox("QUIERE RECOMEDAR ESTE PRODUCTO", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Dim codigo_r As String
        strCadena = "SELECT * FROM Producto_Recomendado ORDER BY id_producto DESC"
        Call ConfiguraRst(strCadena)
        codigo_r = GeneraCodigo(4)
        Set rst = Nothing
        strCadena = "INSERT INTO Producto_Recomendado VALUES ('" & Trim(codigo_r) & "','" & Trim(Me.Txtproducto.Text) & "') "
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Call Resalta(Me.Txtproducto)
        Exit Sub
       End If
End If
End Sub

Private Sub Form_Load()
On Error GoTo salir
  CenterForm Me
  Me.Top = 50

  Me.DtpVencimiento.Value = KEY_FECHA
  strLinea = False
  strMostrarTodos = False
  strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLinea)
  
  strCadena = "SELECT id_marca as Codigo, descripcion as Descripcion FROM marca WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMarca)
  
  
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  
  
  
  
  
  
'Me.Label2.Caption = KEY_EMPRESA
If KEY_RUBRO = "00003" Then
    Me.frmbusquedafarmacia.Visible = True
    Call Actualizar_farmacia
Else
    Me.frmbusquedafarmacia.Visible = False
    Call ActualizarProd

End If

  

'Call ActualizarAlm

If KEY_CARGO = "00001" Or KEY_CARGO = "00009" Or KEY_CARGO = "00004" Or KEY_CARGO = "00057" Then
    Me.cmdNuevo.Enabled = True
    Me.cmdupdate.Enabled = True
    Me.cmdDelete.Enabled = True
    Me.cmdcompatibility.Enabled = True
    Me.CmdFoto.Enabled = True
    Me.cmdCaracteristicas.Enabled = True
Else
    Me.cmdNuevo.Enabled = False
    Me.cmdupdate.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.CmdFoto.Enabled = True
    Me.cmdcompatibility.Enabled = True
    Me.cmdCaracteristicas.Enabled = True
End If

    
salir:
    
End Sub


Private Sub CargarLogo(ByVal cproducto As String, ByVal id_producto As String)
 On Error GoTo salir
    If VerificarArchivo(App.Path & "\archivos\" & KEY_RUC & "\" & cproducto) = True Then
       Me.Image1.Visible = True
       Me.lblFotos.Caption = "1"
       Me.Image1.Picture = LoadPicture(App.Path + "\archivos\" + KEY_RUC + "\" + Trim(cproducto))
       strCadena = "SELECT * FROM producto_foto WHERE id_producto='" & id_producto & "' AND ruc='" & KEY_RUC & "'"
       Call ConfiguraRstI(strCadena)
       If rstI.RecordCount > 0 Then
            Me.CmdAnterior.Enabled = True
            Me.CmdSiguiente.Enabled = True
            Me.lblFotos.Caption = str(rstI.RecordCount)
       Else
            Me.CmdAnterior.Enabled = False
            Me.CmdSiguiente.Enabled = False
       End If
    Else
        Me.Image1.Visible = False
        Me.Image1.Picture = Nothing
    End If
Exit Sub
salir:
    Me.Image1.Picture = Nothing
    Exit Sub
End Sub

Private Sub HfCompatible_DblClick()
If Me.HfCompatible.Rows > 0 Then
    If Val(Me.HfCompatible.TextMatrix(Me.HfCompatible.Row, 0)) > 0 Then
        strLinea = False
        Me.Txtproducto.Text = Trim(Me.HfCompatible.TextMatrix(Me.HfCompatible.Row, 1))
        Call busqueda
        Me.frmcompatible.Visible = False
    End If
End If
End Sub

Private Sub HfdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
        Me.HfdGrilla.AllowBigSelection = True
'        Me.HfdGrilla.SetFocus
End If
If KeyCode = vbKeyUp Then
     Me.HfdGrilla.AllowBigSelection = True
End If
End Sub



Private Sub HfdGrilla_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And FrmVentas.Procedencia = Selecionar Then
     
    strCadena = "SELECT * FROM view_producto_selec WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        
        If get_producto_habilitado(rst("habilitado")) = False Then
           Exit Sub
       End If
        
        FrmVentas.codigoP = Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0))
        FrmVentas.TxtCodProducto.Text = Trim(rst("id_producto"))
        FrmVentas.TxtDescripcionProducto.Text = rst("nombre_prod")
        FrmVentas.txtagranel.Text = rst("agranel")
        FrmVentas.IN_ICBPER = rst("icbper")
        
        Call FrmVentas.get_unidad(Trim(rst("id_producto")), rst("agranel"))
        
        If rst("agranel") = "si" Then
           in_precio = get_precio_unidad(rst("id_producto"), FrmVentas.DtcUnidad.BoundText)
        Else
           
           
           If KEY_SEGMENTACION_PRECIO = "si" Then
               in_precio = get_precio_segmentacion(rst("id_producto"), FrmVentas.TxtCodCliente.Text)
           Else
               
               If KEY_GRIFO = "si" Then
                    in_precio = get_precio_propio(rst("id_producto"), FrmVentas.TxtCodCliente.Text, rst("precio_venta"))
                    
               Else
                    in_precio = rst("precio_venta")
               End If
               
           End If
           
           
        End If
        
        
        
        
        
        If FrmVentas.DtcMoneda.BoundText = "00002" Then
            If KEY_CONVERSION_CAMBIO = "si" Then
                FrmVentas.txtprecio.Text = Format(in_precio, "###0.00")
                FrmVentas.txtpreciooriginal.Text = Format(in_precio, "###0.00")
            Else
                FrmVentas.txtprecio.Text = Format(in_precio / (FrmVentas.TxtTipoCambio.Text), "###0.00")
                FrmVentas.txtpreciooriginal.Text = Format(in_precio / (FrmVentas.TxtTipoCambio.Text), "###0.00")
            End If
            
        Else
            
            If KEY_CONVERSION_CAMBIO = "si" Then
                FrmVentas.txtprecio.Text = in_precio * KEY_CAMBIO_LOCAL
                FrmVentas.txtpreciooriginal.Text = in_precio * KEY_CAMBIO_LOCAL
            Else
                FrmVentas.txtpreciooriginal.Text = in_precio
                FrmVentas.txtprecio.Text = in_precio
            End If
        End If
        
        
        
           
        FrmVentas.txtServicio.Text = rst("servicio")
        If rst("servicio") = "si" Then
           If rst("icbper") = "si" And FrmVentas.txt_tipo.Text = "01" Then
              FrmVentas.txt_tipo.Text = "01"
           Else
              FrmVentas.txt_tipo.Text = "02"
           End If
        Else
           FrmVentas.txt_tipo.Text = "01"
        End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        FrmVentas.txtpeso.Text = rst("peso")
        
       
    
      
        
        FrmVentas.TxtIgv.Text = rst("id_igv")
        If Val(FrmVentas.txtCantidad.Text) <> 0 Then
            FrmVentas.txtCantidad.Text = Val(FrmVentas.txtCantidad.Text)
        Else
            FrmVentas.txtCantidad.Text = 1
        End If
        FrmVentas.chkPrecios.Enabled = True
        FrmVentas.chkPrecios.Value = 1
        
        Call FrmVentas.mostrar_precios
        
       If FrmVentas.txt_tipo.Text <> "02" And IN_ICBPER = "no" Then ' INGRESO SE SERVICIOS
        If rst("stock") <= 0 And FrmVentas.chk_venta_diferida.Value = 0 And KEY_MOVIMIENTO_SIN_STOCK = "no" Then
            MsgBox "PRODUCTO NO CUENTA CON STOCK." + Chr(13) + Chr(13) + "Consulte con el Area de Almacen.", vbInformation, KEY_EMPRESA
            If FrmVentas.DtcTipoDoc.BoundText <> "0099" Then
                
            
                Call Resalta(Me.Txtproducto)
                Exit Sub
            End If
            
            
        End If
        End If
        
         FrmVentas.lblhistorial.Caption = FrmVentas.get_ultimo_precio(Trim(FrmVentas.TxtCodCliente.Text), Trim(Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0))))
        
        
        
        
        
   If FrmVentas.OptAuto.Value = True Then
            Call FrmVentas.Agregar_directo
     Else
        FrmVentas.txtprecio.Locked = False
        Call FrmVentas.Resalta(FrmVentas.txtCantidad)
        'Call FrmVentas.Resalta(FrmVentas.txtprecio)
        'Call FrmVentas.mostrar_precios
        End If
        FrmVentas.Procedencia = Neutro
        Unload Me
        Set rst = Nothing
    End If
    Exit Sub
End If

If FrmDetalleProducto.Procedencia = Selecionar Then
   FrmDetalleProducto.Procedencia = Neutro
   FrmDetalleProducto.txtCodCompatible.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmDetalleProducto.TxtCompatible.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   Unload Me
   Exit Sub
End If

If FrmDetalleAlmacen.Procedencia = Selecionar Then
   FrmDetalleAlmacen.Procedencia = Neutro
   FrmDetalleAlmacen.txtcodigo_producto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmDetalleAlmacen.lblproducto.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   Unload Me
   Exit Sub
End If


If FrmCambioAceite.Procedencia = Selecionar Then
   FrmCambioAceite.Procedencia = Neutro
   FrmCambioAceite.txtCodigo.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmCambioAceite.txtDescripcionRepuesto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   FrmCambioAceite.txtprecio.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 7)
   FrmCambioAceite.txtCantidad.Text = 1
   Call Resalta(FrmCambioAceite.txtCantidad)
   Unload Me
   Exit Sub
End If



If frmpersonaDeuda.Procedencia = Selecionar Then
   frmpersonaDeuda.Procedencia = Neutro
   frmpersonaDeuda.txtId_producto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   frmpersonaDeuda.txtServicio.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   Call Resalta(frmpersonaDeuda.txtServicio)
   Unload Me
   Exit Sub
End If



If frmHotelInfraestructura.Procedencia = Selecionar Then
   frmHotelInfraestructura.Procedencia = Neutro
   frmHotelInfraestructura.txtId_producto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   frmHotelInfraestructura.Txtproducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   frmHotelInfraestructura.txtdescripcionhabitacion.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   frmHotelInfraestructura.txtprecio.Text = get_precio_producto(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0), KEY_ALM)
   Unload Me
   Exit Sub
End If


If frmventaslistado.Procedencia = Selecionar Then
   frmventaslistado.Procedencia = Neutro
   frmventaslistado.txtid_productomasivo.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   frmventaslistado.txtproductomasivo.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   frmventaslistado.txtprecio.Text = get_precio_producto(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0), KEY_ALM)
   Unload Me
   Exit Sub
End If


If frmHotel.Procedencia = Selecionar Then
   frmHotel.txtid_producto_habitacion.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   frmHotel.lblproducto_habitacion(0).Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   frmHotel.txtPrecio_habitacion.Text = get_precio_producto(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0), KEY_ALM)
   frmHotel.Procedencia = Neutro
   frmHotel.txtcantidad_habitacion.Text = 1
   Call Resalta(frmHotel.txtcantidad_habitacion)
   
   Unload Me
   Exit Sub
End If



If frmHotel.Procedencia = seleccionar_otro Then
   frmHotel.txtidproducto_resta.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   frmHotel.lblproducto_resta(0).Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   frmHotel.txtprecio_rest.Text = get_precio_producto(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0), KEY_ALM)
   frmHotel.Procedencia = Neutro
   frmHotel.txtcantidad_resta.Text = 1
   Call Resalta(frmHotel.txtcantidad_resta)
   Unload Me
   Exit Sub
End If



If FrmBonificacion.Procedencia = Selecionar Then
   FrmBonificacion.Procedencia = Neutro
   FrmBonificacion.txtId_producto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmBonificacion.lblproducto.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   
   strCadena = "SELECT agranel FROM producto WHERE id_producto='" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
   Call ConfiguraRstA(strCadena)
   If rstA.RecordCount > 0 Then
      Call FrmBonificacion.get_unidad(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0), rstA("agranel"), FrmBonificacion.DtcUnidad)
   End If
   Unload Me
   Exit Sub
End If


If frmPlanesServicio.Procedencia = Selecionar Then
   frmPlanesServicio.Procedencia = Neutro
   frmPlanesServicio.txtId_producto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   frmPlanesServicio.lblproducto.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   
   Unload Me
   Exit Sub
End If




If FrmSolicitudViaticosDeclarar.Procedencia = Selecionar Then
   FrmSolicitudViaticosDeclarar.Procedencia = Neutro
   FrmSolicitudViaticosDeclarar.txtid_servicio.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmSolicitudViaticosDeclarar.TxtDetalle.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   Call FrmSolicitudViaticosDeclarar.DtcMoneda.SetFocus
   Unload Me
   
   Exit Sub
End If


If FrmBonificacion.Procedencia = buscar Then
   FrmBonificacion.Procedencia = Neutro
   FrmBonificacion.txtidproductocruzado.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmBonificacion.lblproductocruzada.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   strCadena = "SELECT agranel FROM producto WHERE id_producto='" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
   Call ConfiguraRstA(strCadena)
   If rstA.RecordCount > 0 Then
      Call FrmBonificacion.get_unidad(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0), rstA("agranel"), FrmBonificacion.DtcUnidad)
   End If
   Unload Me
   Exit Sub
   Unload Me
   Exit Sub
End If


If FrmBonificacion.Procedencia = relacionar Then
   FrmBonificacion.Procedencia = Neutro
   FrmBonificacion.txtid_producto_boni_cruzada.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmBonificacion.lblproducto_boni_cruzada.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   strCadena = "SELECT agranel FROM producto WHERE id_producto='" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
   Call ConfiguraRstA(strCadena)
   If rstA.RecordCount > 0 Then
      Call FrmBonificacion.get_unidad(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0), rstA("agranel"), FrmBonificacion.DtcUnidadBoni)
   End If
   Unload Me
   Exit Sub
End If




If frmVariacionCostes.Procedencia = Selecionar Then
   frmVariacionCostes.Procedencia = Neutro
   Unload Me
End If


If frmsurtidores.Procedencia = Selecionar Then
   frmsurtidores.Procedencia = Neutro
   frmsurtidores.txtcodigoproductotanque.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   frmsurtidores.lblproductotanque.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   Unload Me
   Exit Sub
End If


If frmsurtidores.Procedencia = buscar Then
   frmsurtidores.Procedencia = Neutro
   frmsurtidores.txtcodigoproducto_surtidor.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   frmsurtidores.lblproductosurtidor.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   Unload Me
   Exit Sub
End If



If FrmReporteProducto.Procedencia = Selecionar Then
    FrmReporteProducto.Procedencia = Neutro
    FrmReporteProducto.DtcProductogen.BoundText = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
    Unload Me
    Exit Sub
End If


If FrmComprasGastos.Procedencia = Selecionar Then
   FrmComprasGastos.Procedencia = Neutro
   FrmComprasGastos.lblcuenta_contable.Caption = get_cuenta_contable_producto(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0))
   FrmComprasGastos.lblcuenta_detalle.Caption = get_cuenta(FrmComprasGastos.lblcuenta_contable.Caption)
   FrmComprasGastos.txtcodigoprod.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmComprasGastos.Txtproducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
 
   Unload Me
   Exit Sub
End If

If FrmOrdenCompraDet.Procedencia = buscar Then
   FrmOrdenCompraDet.Procedencia = Neutro
   FrmOrdenCompraDet.lblcuenta_contable.Caption = get_cuenta_contable_producto(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0))
   FrmOrdenCompraDet.lblcuenta_detalle.Caption = get_cuenta(FrmOrdenCompraDet.lblcuenta_contable.Caption)
   FrmOrdenCompraDet.txtcodigoprod.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmOrdenCompraDet.Txtproducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   Unload Me
   Exit Sub
End If










If KeyAscii = 13 And FrmProductosRelacionados.Procedencia = relacionar Then
    FrmProductosRelacionados.Procedencia = Neutro
    strCadena = "SELECT id_producto,nombre_prod FROM producto WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' AND ruc='" & KEY_RUC & "'  "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmProductosRelacionados.txtCodSubProducto.Text = rst(0)
        FrmProductosRelacionados.txtDescripcionSubproducto.Text = rst(1)
        Call Resalta(FrmProductosRelacionados.txtCantidad)
        Unload Me
        
        Set rst = Nothing
    End If
    Exit Sub
End If

If KeyAscii = 13 And FrmProductoSubproducto.Procedencia = relacionar Then
    
        FrmProductoSubproducto.txtCodSubProducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
        FrmProductoSubproducto.txtDescripcionSubproducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
        Call Resalta(FrmProductoSubproducto.txtCantidad)
        Unload Me
        FrmProductoSubproducto.Procedencia = Neutro
     
    
    Exit Sub
End If


If frmCorProcesos.Procedencia = seleccionar_soldadura Then
        strCadena = "INSERT INTO imp_producto_insumo(id_producto_detalle,id_producto,id_linea,ruc)VALUES('" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "','" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "','" & frmCorProcesos.Txtid_estado.Text & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        Call frmCorProcesos.llena_insumos(frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0), frmCorProcesos.Txtid_estado.Text, frmCorProcesos.HfSoldaduraInsumo)
        frmCorProcesos.Procedencia = Neutro
        Unload Me
        
        Exit Sub
End If


If frmCorProcesos.Procedencia = seleccionar_ensamblaje Then
        frmCorProcesos.Procedencia = Neutro
        strCadena = "INSERT INTO imp_producto_insumo(id_producto_detalle,id_producto,id_linea,ruc)VALUES('" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "','" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "','" & frmCorProcesos.Txtid_estado.Text & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
         
        Unload Me
        Call frmCorProcesos.llena_insumos(frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0), frmCorProcesos.Txtid_estado.Text, frmCorProcesos.HfEnsambladoInsumo)
        frmCorProcesos.Procedencia = Neutro
        Exit Sub
End If

If FrmDetalleLinea.Procedencia = seleccionar_insumo Then
   FrmDetalleLinea.txtid_insumo.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmDetalleLinea.lblinsumo.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   Call Resalta(FrmDetalleLinea.txtCantidad)
   FrmDetalleLinea.Procedencia = Neutro
   Unload Me
   Exit Sub
End If

If frmmantenimientos.Procedencia = seleccionar_insumo Then
   frmmantenimientos.frminsumo.Visible = True
   frmmantenimientos.txtid_insumo.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   frmmantenimientos.lblproducto.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   Call Resalta(frmmantenimientos.txtCantidad)
   frmmantenimientos.Procedencia = Neutro
   Unload Me
   Exit Sub
End If

If FrmOrdenCompraDet.Procedencia = Selecionar Then
   FrmOrdenCompraDet.Procedencia = Neutro
    strCadena = "SELECT * FROM view_producto WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmOrdenCompraDet.TxtCodProducto.Text = rst("id_producto")
        FrmOrdenCompraDet.TxtDescripcionProducto.Text = rst("nombre_prod")
        FrmOrdenCompraDet.TxtUnidad.Text = rst("unidad")
        FrmOrdenCompraDet.txtcosto.Text = rst("precio_compra")
        Call Resalta(FrmOrdenCompraDet.txtCantidad)
    End If
    Unload Me
    Exit Sub
End If



If FrmDetallePedido.Procedencia = Selecionar Then
   FrmDetallePedido.Procedencia = Neutro
    strCadena = "SELECT * FROM view_producto WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmDetallePedido.TxtCodProducto.Text = rst("id_producto")
        FrmDetallePedido.TxtDescripcionProducto.Text = rst("nombre_prod")
        FrmDetallePedido.TxtUnidad.Text = rst("unidad")
        FrmDetallePedido.txtcosto.Text = rst("precio_compra")
        Call Resalta(FrmDetallePedido.txtCantidad)
    End If
    Unload Me
    Exit Sub
End If







If frmCorProcesos.Procedencia = seleccionar_tapiz Then
        strCadena = "INSERT INTO imp_producto_insumo(id_producto_detalle,id_producto,id_linea,ruc)VALUES('" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "','" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "','" & frmCorProcesos.Txtid_estado.Text & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
         
        Unload Me
        Call frmCorProcesos.llena_insumos(frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0), frmCorProcesos.Txtid_estado.Text, frmCorProcesos.HfTapizadoInsumo)
        frmCorProcesos.Procedencia = Neutro
        Exit Sub
End If

If frmCorProcesos.Procedencia = seleccionar_otro Then
        strCadena = "INSERT INTO imp_producto_insumo(id_producto_detalle,id_producto,id_linea,ruc)VALUES('" & frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0) & "','" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "','" & frmCorProcesos.Txtid_estado.Text & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
         
        Unload Me
        Call frmCorProcesos.llena_insumos(frmCorProcesos.gridDetalle.TextMatrix(frmCorProcesos.gridDetalle.Row, 0), frmCorProcesos.Txtid_estado.Text, frmCorProcesos.HfInsumoTercero)
        frmCorProcesos.Procedencia = Neutro
        Exit Sub
End If




If KeyAscii = 13 And FrmKardexdeProductos.Procedencia = buscar Then
        FrmKardexdeProductos.DtcProducto.BoundText = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
        Call FrmKardexdeProductos.presionar
         
        Unload Me
        FrmKardexdeProductos.Procedencia = Neutro
    Exit Sub
End If

If KeyAscii = 13 And FrmGeneradorBarras.Procedencia = Selecionar Then
        FrmGeneradorBarras.Procedencia = Neutro
        FrmGeneradorBarras.txtCodigo.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
        Call FrmGeneradorBarras.precionar(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0))
         
        Unload Me
       
    Exit Sub
End If

If KeyAscii = 13 And FrmVentasPersonalizada.Procedencia = Selecionar Then
   FrmVentasPersonalizada.TxtCodProducto(FrmVentasPersonalizada.numeroItem).Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   FrmVentasPersonalizada.txtDescripcion(FrmVentasPersonalizada.numeroItem).Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
   FrmVentasPersonalizada.txtprecio(FrmVentasPersonalizada.numeroItem).Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 8)
   FrmVentasPersonalizada.TxtUnidad(FrmVentasPersonalizada.numeroItem).Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 5)
   Call Resalta(FrmVentasPersonalizada.txtCantidad(FrmVentasPersonalizada.numeroItem))
   FrmVentasPersonalizada.Procedencia = Neutro
   Unload Me
   Exit Sub
End If





If frmVentasPagos.Procedencia = Selecionar Then
   frmVentasPagos.txtId_producto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   frmVentasPagos.txtObservacion.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1) & "-" & "COLOR :" & "- " & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 4) & "-" & Trim(frmVentasPagos.txtObservacion.Text) & Space(2)
   frmVentasPagos.framevehiculo.Visible = True
   frmVentasPagos.txtmontovehiculo.Text = Format(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 8), "###0.00")
   frmVentasPagos.TxtMontoReal.Text = Format(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 8), "###0.00")
   frmVentasPagos.txtsaldo.Text = Format(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 8), "###0.00")
   frmVentasPagos.Procedencia = Neutro
   Unload Me
   Exit Sub
End If

'If FrmReporteProducto.Procedencia = buscar Then
  ' FrmReporteProducto.Procedencia = Neutro
   'FrmReporteProducto.txtcodigo.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   'FrmReporteProducto.txtcodigo.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
 '  Unload Me
'   Exit Sub
'End If

If KeyAscii = 13 And FrmBusquedaDocumentos.Procedencia = Selecionar Then
    FrmBusquedaDocumentos.TxtCodigoInterno.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
    FrmBusquedaDocumentos.txtCodigoProducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
    FrmBusquedaDocumentos.TxtUnidad.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 2)
    FrmBusquedaDocumentos.txtDescripcion.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
    FrmBusquedaDocumentos.cmdBuscar_producto.Enabled = True
    FrmBusquedaDocumentos.cmdBuscar_producto.SetFocus
    FrmBusquedaDocumentos.Procedencia = Neutro
    Exit Sub
End If


If KeyAscii = 13 And FrmDetalleLinea.Procedencia = Selecionar Then
    FrmDetalleLinea.txtId_producto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
    FrmDetalleLinea.lblproducto.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
    FrmDetalleLinea.lblcosto.Caption = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 7)
    
    FrmDetalleLinea.Procedencia = Neutro
    Unload Me
    Exit Sub
End If

'****************MERMAS******************
If KeyAscii = 13 And FrmProductoMermas.Procedencia = mermas Then
     strCadena = "SELECT P.id_producto,P.nombre_prod, U.abreviatura,A.stock,P.precio_compra FROM producto P,almacen_producto A ,unidad U WHERE A.id_alm='" & Trim(FrmProductoMermas.DtcAlmacen.BoundText) & "'" & _
    " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND P.id_producto=A.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "' AND P.id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmProductoMermas.cod_producto = rst("id_producto")
        FrmProductoMermas.txtId_producto.Text = rst("id_producto")
        FrmProductoMermas.Txtproducto.Text = rst("nombre_prod")
        FrmProductoMermas.lblStock.Caption = rst("stock")
        FrmProductoMermas.txtcosto.Text = rst("precio_compra")
        'FrmProductoMermas.DtcMerma.SetFocus
        FrmProductoMermas.Procedencia = Neutro
        Unload Me
       
        Set rst = Nothing
    End If
    Exit Sub
End If


If KeyAscii = 13 And FrmProductoTransformaciones.Procedencia = transformaciones And FrmProductoTransformaciones.prodA = True Then
    strCadena = "SELECT P.id_producto,P.nombre_prod, U.abreviatura,A.stock,P.precio_compra,P.precio_venta FROM producto P,almacen_producto A ,unidad U WHERE A.id_alm='" & Trim(FrmProductoTransformaciones.DtcAlmacen.BoundText) & "'" & _
    " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND P.id_producto=A.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "' AND P.id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "'"
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmProductoTransformaciones.txtCodigo.Text = rst("id_producto")
        FrmProductoTransformaciones.Txtproducto.Text = rst("nombre_prod")
        FrmProductoTransformaciones.lblStock.Caption = rst("stock")
        FrmProductoTransformaciones.txtcosto.Text = rst("precio_compra")
        FrmProductoTransformaciones.LblUnidad.Caption = rst("abreviatura")
        FrmProductoTransformaciones.TxtPVenta.Text = rst("precio_venta")
        Unload Me
        Set rst = Nothing
    End If
    FrmProductoTransformaciones.Procedencia = Neutro
    FrmProductoTransformaciones.prodA = False
    Exit Sub
End If
If KeyAscii = 13 And FrmProductoTransformaciones.Procedencia = transformaciones And FrmProductoTransformaciones.prodB = True Then
    strCadena = "SELECT P.id_producto,P.nombre_prod, U.abreviatura,A.stock,P.precio_compra,P.precio_venta FROM producto P,almacen_producto A ,unidad U WHERE A.id_alm='" & Trim(FrmProductoTransformaciones.DtcAlmacen.BoundText) & "'" & _
    " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND P.id_producto=A.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "' AND P.id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "'"
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmProductoTransformaciones.TxtCodigobarraB.Text = rst("id_producto")
        FrmProductoTransformaciones.txtDescripcionB.Text = rst("nombre_prod")
        FrmProductoTransformaciones.LblStockB.Caption = rst("stock")
        FrmProductoTransformaciones.txtCostoB.Text = rst("precio_compra")
        FrmProductoTransformaciones.lblUndB.Caption = rst("abreviatura")
        FrmProductoTransformaciones.TxtVentaB.Text = rst("precio_venta")
        Unload Me
        Set rst = Nothing
    End If
     FrmProductoTransformaciones.Procedencia = Neutro
    FrmProductoTransformaciones.prodB = False
    Exit Sub
End If






If KeyAscii = 13 And FrmProductoTransformaciones.Procedencia = transformaciones Then
    Me.HfdGrilla.col = 0
    
   strCadena = "SELECT Producto_barras.cProducto, Producto_barras.cod_barra, Producto.DescripcionProducto, Unidad.sAbreviatura, " & _
    "Almacen_Productos.Stock,Producto.PrecioCompra,Producto.PrecioVenta FROM Producto_barras INNER JOIN Producto ON Producto_barras.cProducto = Producto.cProducto INNER JOIN " & _
    "Almacen_Productos ON Producto_barras.cProducto = Almacen_Productos.cProducto INNER JOIN Unidad ON Producto.cUnidad = Unidad.cUnidad " & _
    "WHERE Almacen_Productos.cProducto='" & Trim(Me.HfdGrilla.Text) & "' AND Almacen_Productos.Alm_cod='" & Trim(FrmProductoTransformaciones.DtcAlmacen.BoundText) & "'"
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        FrmProductoTransformaciones.cod_producto = rst(0)
        FrmProductoTransformaciones.txtCodigo.Text = rst(1)
        FrmProductoTransformaciones.Txtproducto.Text = rst(2)
        FrmProductoTransformaciones.lblStock.Caption = rst(4)
        FrmProductoTransformaciones.txtcosto.Text = rst(5)
        FrmProductoTransformaciones.LblUnidad.Caption = rst("sAbreviatura")
        FrmProductoTransformaciones.TxtPVenta.Text = rst("PrecioVenta")
        Unload Me
        Set rst = Nothing
    End If
    Exit Sub
End If

If KeyAscii = 13 And FrmInventario.Procedencia = Selecionar Then
       
       strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "'  and id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' AND id_alm='" & Trim(FrmInventario.DtcAlmacen.BoundText) & "'"
       Call ConfiguraRst(strCadena)
    
    If rst.RecordCount > 0 Then
         FrmInventario.TxtCodProducto.Text = rst("id_producto")
         FrmInventario.txtId_producto.Text = rst("id_producto")
         FrmInventario.TxtDescripcionProducto.Text = rst("nombre_prod")
         FrmInventario.TxtStck_actual.Text = rst("stock")
         FrmInventario.txtStock_factura.Text = rst("stock_factura")
         FrmInventario.TxtUnidad.Text = rst("unidad")
         FrmInventario.TxtVenta.Text = rst("precio_venta")
         FrmInventario.txtcosto.Text = rst("precio_compra")
         FrmInventario.DtcClasificacion.BoundText = rst("id_linea")
         FrmInventario.DtcModelo.BoundText = rst("id_sublinea")
         FrmInventario.cmdStock.Enabled = True
         If rst("produccion") = "si" Then
            FrmInventario.cmdSeriales.Visible = True
         Else
            FrmInventario.cmdSeriales.Visible = False
         End If
         Call FrmInventario.Resalta(FrmInventario.TxtStock_nuevo)
        
         Set rst = Nothing
    End If
        FrmInventario.Procedencia = Neutro
         Unload Me
    Exit Sub
End If




If KeyAscii = 13 And FrmOrdenCompraDet.Procedencia = Selecionar Then
       FrmOrdenCompraDet.Procedencia = Neutro
       strCadena = "SELECT * FROM view_producto WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' AND id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
         FrmOrdenCompraDet.TxtCodProducto.Text = rst("id_producto")
         FrmOrdenCompraDet.TxtDescripcionProducto.Text = rst("nombre_prod")
         FrmOrdenCompraDet.TxtUnidad.Text = rst("unidad")
         FrmOrdenCompraDet.txtcosto.Text = rst("precio_compra")
         Call Resalta(FrmOrdenCompraDet.txtCantidad)
    End If
   
    Unload Me
    Exit Sub
End If


If KeyAscii = 13 And frmReportesGenerales.Procedencia = Selecionar Then
    frmReportesGenerales.Procedencia = Neutro
    frmReportesGenerales.DtcProducto.BoundText = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
    Unload Me
    Exit Sub
End If


If KeyAscii = 13 And FrmCompras.Procedencia = Selecionar Then
   Dim utilidad As Single
   FrmCompras.Procedencia = Neutro
    
   
   FrmCompras.TxtCodProducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
   
   Call FrmCompras.put_producto
   Unload Me
   
   Exit Sub
    
    strCadena = "SELECT A.id_producto,P.nombre_prod,U.descripcion as abreviatura,A.stock,A.precio_compra,A.precio_venta,P.id_linea,P.numero_procedimientos,P.agranel FROM almacen_producto A,producto P,unidad U WHERE A.id_producto=P.id_producto AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND A.id_alm='" & Trim(FrmCompras.DtcAlmacen.BoundText) & "' AND A.id_producto='" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
         FrmCompras.codigoP = Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0))
         FrmCompras.TxtCodProducto.Text = Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0))
         FrmCompras.TxtDescripcionProducto.Text = UCase(rst("nombre_prod"))
         
         Call FrmCompras.get_unidad(rst("id_producto"), rst("agranel"))
         FrmCompras.txtCantidad.Text = 1
         FrmCompras.TxtUnidades.Text = rst("numero_procedimientos")
         
         If KEY_CON_IGV = "si" Then
            FrmCompras.TxtCostoAnt.Text = rst("precio_compra") * (1 + KEY_IGV)
            FrmCompras.txtcosto.Text = rst("precio_compra") * (1 + KEY_IGV)
         Else
            FrmCompras.TxtCostoAnt.Text = rst("precio_compra")
            FrmCompras.txtcosto.Text = rst("precio_compra")
         End If
         
         
         FrmCompras.txtPrecioVentaAnt.Text = rst("precio_venta")
         FrmCompras.TxtventaHoy.Text = rst("precio_venta")
         
         If get_produccion(rst("id_linea")) = True Then
            FrmCompras.txtvalidacion_chasis.Text = "si"
         Else
            FrmCompras.txtvalidacion_chasis.Text = "no"
         End If
         
         FrmCompras.TxtUnidades.Text = get_cantidad_agranel(Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)), FrmCompras.DtcUnidad.BoundText)
         
         
         If rst("precio_compra") > 0 Then
            utilidad = (rst("precio_venta") - rst("precio_compra")) * 100 / rst("precio_compra")
            FrmCompras.TxtUtilidadAnt.Text = Format(utilidad, "#,##0.00")
         End If
         
         If KEY_PAIS <> KEY_PERU And FrmCompras.DtcTipoDoc.BoundText = "0020" Then
            FrmCompras.FRameiva.Visible = True
         
         End If
         
         
        Call Resalta(FrmCompras.txtCantidad)
        
        Unload Me
        
    End If
    Exit Sub
End If




If KeyAscii = 13 And FrmTransferencias.Procedencia = Selecionar Then
        FrmTransferencias.Procedencia = Neutro
        strCadena = "SELECT * FROM producto WHERE id_producto='" & Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            FrmTransferencias.TxtCodProducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
            FrmTransferencias.cprod = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
            FrmTransferencias.TxtDescripcionProducto.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 1)
            FrmTransferencias.TxtUnidad.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 5)
            FrmTransferencias.txtpeso.Text = rst("peso")
            FrmTransferencias.TxtUnidad.Text = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 3)
            FrmTransferencias.txtCantidad.Enabled = True
            FrmTransferencias.txtCantidad.Text = ""
            Call Resalta(FrmTransferencias.txtCantidad)
            Unload Me
            Set rst = Nothing
    End If
    
    Exit Sub
End If





End Sub
Private Function get_produccion(ByVal in_linea As String) As Boolean
strCadena = "SELECT produccion FROM linea WHERE id_linea='" & in_linea & "' and id_usu='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   If rstK("produccion") = "si" Then
      get_produccion = True
   Else
      get_produccion = False
   End If
   
End If


End Function
Private Sub HfdGrilla_SelChange()
  
  
  'strCadena = "SELECT * FROM view_producto_almacen WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
  'Call ConfiguraRstP(strCadena)
  'If rstP.RecordCount > 0 Then
  '   rstP.MoveFirst
     Me.lblproducto.Caption = get_producto(Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))
     'img = rstP("imagen")
    If Trim(Me.lblproducto.Caption) <> "" Then
     If KEY_STOCK_CONTABLE = "no" Then
        Call ActualizarAlm(Me.HfgAlmacen, Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))
    Else
        strCadena = "SELECT * FROM view_producto_almacen WHERE id_producto='" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstP(strCadena)
        Call Me.llenar_almacen_contable(Me.HfgAlmacen, Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))
    End If
    
    
    
 '   If KEY_FOTO = "si" And Len(img) > 0 Then
 '       Me.Image1.Visible = True
 '       Call CargarLogo(img, Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))
    
 '       Else
 '       Me.Image1.Visible = False
 '   End If

If Me.frmFoto.Visible = True Then
   Me.frmFoto.Visible = False
End If

If KEY_GRUPO_EMPRESARIAL = "si" Then
    Me.hfgrupoempresarial.Visible = True
    Call load_grupo_empresarial(Me.hfgrupoempresarial, Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))
Else
    Me.hfgrupoempresarial.Visible = False
End If



If KEY_CARGO = "00001" Or KEY_CARGO = "00009" Or KEY_CARGO = "00004" Or KEY_CARGO = "00057" Then
    Me.cmdNuevo.Enabled = True
    Me.cmdupdate.Enabled = True
    Me.cmdDelete.Enabled = True
    Me.cmdcompatibility.Enabled = True
    Me.CmdFoto.Enabled = True
    Me.cmdCaracteristicas.Enabled = True
Else
    Me.cmdNuevo.Enabled = False
    Me.cmdupdate.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.CmdFoto.Enabled = True
    Me.cmdcompatibility.Enabled = True
    Me.cmdCaracteristicas.Enabled = True
End If
Else
    Me.HfgAlmacen.Rows = 0
    Me.cmdNuevo.Enabled = True
    Me.cmdupdate.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdCaracteristicas.Enabled = False
    Me.cmdcompatibility.Enabled = False
    Me.CmdFoto.Enabled = False
End If


If KEY_GRIFO = "si" Then
    Call load_tanque(Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))
End If



End Sub


Private Sub load_tanque(ByVal in_producto As String)

For i = 0 To 2
    Me.Cisterna_llena(i).Visible = False
    Me.Cisterna_vacia(i).Visible = False
    Me.lblCisterna(i).Visible = False
Next i


strCadena = "SELECT descripcion,minimo,maxima,funct_stock(id_producto,'" & KEY_ALM & "',ruc) as stock FROM view_tanque_producto WHERE estado='si' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.frmCisterna.Visible = True
   For i = 0 To rst.RecordCount - 1
        Me.Cisterna_llena(i).Visible = True
        Me.Cisterna_vacia(i).Visible = True
        Me.lblCisterna(i).Visible = True
       Me.lblCisterna(i).Caption = rst("descripcion")
       Me.lblmin.Caption = rst("minimo")
       Me.lblmedio.Caption = (rst("minimo") + rst("maxima")) / 2
       Me.lblmax.Caption = rst("maxima")
       If rst("stock") > 0 Then
       
       If rst("stock") > rst("maxima") Then
          Me.Cisterna_vacia(i) = 0
       Else
         Me.Cisterna_vacia(i).Height = 3855 - rst("stock") * 3855 / rst("maxima")

       End If
       Else
        Me.Cisterna_vacia(i).Height = 3855
       End If
       
       
       rst.MoveNext
   Next i
End If


End Sub


Sub llenarGrid_prod(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
Dim X As Integer
On Error GoTo salir
  Call ConfiguraRst(strCadena)
  Grilla.Rows = 1
  Set Grilla.Recordset = rst
  Grilla.Rows = rst.RecordCount
  Grilla.ColWidth(0) = 600
  Grilla.ColWidth(1) = 5500
  Grilla.ColWidth(2) = 700
  Grilla.ColWidth(3) = 1100
  Grilla.ColAlignment(3) = 7
  Grilla.ColWidth(4) = 1100
  Grilla.ColAlignment(4) = 7
  Grilla.ColWidth(5) = 1100
  Grilla.ColAlignment(5) = 7
  Formulario.TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
  Formulario.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal sql As String)
Dim in_precio As Single
On Error GoTo salir

Call ConfiguraRst(sql)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Me.cmdCaracteristicas.Enabled = False
    Me.cmdcompatibility.Enabled = False
    Me.cmdupdate.Enabled = False
    Me.cmdDelete.Enabled = False
    Exit Sub
End If
  
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       If KEY_MODELO_COLOR = "si" Then
            For Each Campo In rst.Fields
                Grilla.ColWidth(0) = 750
                Grilla.ColWidth(1) = 4100
                Grilla.ColWidth(2) = 1200
                Grilla.ColWidth(3) = 1200
                Grilla.ColWidth(4) = 1100
                Grilla.ColWidth(5) = 1000
                Grilla.ColWidth(6) = 1000
                Grilla.ColWidth(7) = 950
                Grilla.ColWidth(8) = 950
                Grilla.ColWidth(9) = 1000
            Next
                If KEY_CONVERSION_CAMBIO = "si" Then
                    cabecera = "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "CLASIFICACION" & vbTab & "MODELO" & vbTab & "COLOR" & vbTab & "UNIDAD" & vbTab & "MARCA" & vbTab & "P.MAYOR" & vbTab & "VENTA [US $]" & vbTab & "VENTA [S/.]"
                Else
                    cabecera = "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "CLASIFICACION" & vbTab & "MODELO" & vbTab & "COLOR" & vbTab & "UNIDAD" & vbTab & "MARCA" & vbTab & "P.MAYOR" & vbTab & "P.COSTO" & vbTab & "P.VENTA"
                End If
                Grilla.AddItem cabecera
                For k = 0 To 9
                    Grilla.col = k
                    Grilla.Row = 0
                    Grilla.CellBackColor = &HDFDFE0
                Next k
        Else
            For Each Campo In rst.Fields
                Grilla.ColWidth(0) = 700
                If KEY_RUC = "20480516771" Then
                    Grilla.ColWidth(1) = 7500
                    Grilla.ColWidth(5) = 0
                Else
                    Grilla.ColWidth(1) = 6500
                    Grilla.ColWidth(5) = 1000
                End If
                Grilla.ColWidth(2) = 1200
                Grilla.ColWidth(3) = 700
                Grilla.ColWidth(4) = 1100
               
                Grilla.ColWidth(6) = 1000
                Grilla.ColWidth(7) = 1000
            Next
                If KEY_CONVERSION_CAMBIO = "si" Then
                    cabecera = "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "CLASIFICACION" & vbTab & "UND" & vbTab & "MARCA" & vbTab & "P.MAYOR" & vbTab & "VENTA [US$]" & vbTab & "VENTA [S/.]"
                Else
                    cabecera = "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "CLASIFICACION" & vbTab & "UND" & vbTab & "MARCA" & vbTab & "P.MAYOR" & vbTab & "P.COSTO" & vbTab & "P.VENTA"
                End If
                Grilla.AddItem cabecera
                For k = 0 To 7
                    Grilla.col = k
                    Grilla.Row = 0
                    Grilla.CellBackColor = &HDFDFE0
                Next k
        End If
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            If rst("presentacion") > 1 Then
               in_nombre = rst("nombre_prod") & " : TRAE [" & rst("presentacion") & "]"
            Else
               in_nombre = rst("nombre_prod")
            End If
            
            
            If rst("id_igv") = "si" Then
               in_precio = rst("precio_venta")
            Else
               in_precio = rst("precio_venta")
            End If
            
            
            If KEY_CARGO = "00052" Or KEY_CARGO = "00008" Or KEY_CARGO = "00001" Then
                in_precio_costo = "[***]"
            Else
                in_precio_costo = Format(rst("precio_compra"), "#,##0.00")
            End If
            
            
            
            
            If KEY_CONVERSION_CAMBIO = "si" Then
               in_precio_costo = Format(in_precio, "#,##0.00")
               in_precio = in_precio_costo * KEY_CAMBIO_LOCAL
            End If
            If KEY_MODELO_COLOR = "si" Then
                Fila = rst("id_producto") & vbTab & in_nombre & vbTab & UCase(rst("linea")) & vbTab & UCase(rst("modelo")) & vbTab & rst("color") & vbTab & rst("unidad") & vbTab & rst("marca") & vbTab & Format(rst("precio_mayor"), "#,##0.00") & vbTab & in_precio_costo & vbTab & Format(in_precio, "#,##0.00")
            Else
                Fila = rst("id_producto") & vbTab & in_nombre & vbTab & UCase(rst("linea")) & vbTab & rst("unidad") & vbTab & rst("marca") & vbTab & Format(rst("precio_mayor"), "#,##0.00") & vbTab & in_precio_costo & vbTab & Format(in_precio, "#,##0.00")
            End If
            Grilla.AddItem Fila
            rst.MoveNext
    Next i
  Grilla.ColAlignment(1) = 1
  Grilla.ColAlignment(3) = 1
  Grilla.ColAlignment(5) = 1
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Public Sub llenarGrid_farmacia(ByVal Grilla As MSHFlexGrid, ByVal sql As String)
Dim in_precio As Single
On Error GoTo salir

Call ConfiguraRst(sql)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Me.cmdCaracteristicas.Enabled = False
    Me.cmdcompatibility.Enabled = False
    Me.cmdupdate.Enabled = False
    Me.cmdDelete.Enabled = False
    Exit Sub
End If
  
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       If KEY_MODELO_COLOR = "si" Then
            For Each Campo In rst.Fields
                Grilla.ColWidth(0) = 750
                Grilla.ColWidth(1) = 3500
                Grilla.ColWidth(2) = 3000
                Grilla.ColWidth(3) = 1200
                Grilla.ColWidth(4) = 1100
                Grilla.ColWidth(5) = 1000
                Grilla.ColWidth(6) = 1000
                Grilla.ColWidth(7) = 950
                Grilla.ColWidth(8) = 950
                Grilla.ColWidth(9) = 1000
            Next
                If KEY_CONVERSION_CAMBIO = "si" Then
                    cabecera = "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "PRINCIPIO ACTIVO" & vbTab & "MODELO" & vbTab & "COLOR" & vbTab & "UNIDAD" & vbTab & "MARCA" & vbTab & "P.MAYOR" & vbTab & "VENTA [US $]" & vbTab & "VENTA [S/.]"
                Else
                    cabecera = "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "PRINCIPIO ACTIVO" & vbTab & "MODELO" & vbTab & "COLOR" & vbTab & "UNIDAD" & vbTab & "MARCA" & vbTab & "P.MAYOR" & vbTab & "P.COSTO" & vbTab & "P.VENTA"
                End If
                Grilla.AddItem cabecera
                For k = 0 To 9
                    Grilla.col = k
                    Grilla.Row = 0
                    Grilla.CellBackColor = &HDFDFE0
                Next k
        Else
            For Each Campo In rst.Fields
                Grilla.ColWidth(0) = 700
                Grilla.ColWidth(1) = 4000
                Grilla.ColWidth(2) = 2200
                Grilla.ColWidth(3) = 2100
                Grilla.ColWidth(4) = 1100
                Grilla.ColWidth(5) = 1000
                Grilla.ColWidth(6) = 1000
                Grilla.ColWidth(7) = 1000
                
            Next
                
                cabecera = "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "PRINCIPIO ACTIVO" & vbTab & "ACCION.FARMACOLOGICA" & vbTab & "VENCE" & vbTab & "UND" & vbTab & "P.COSTO" & vbTab & "P.VENTA"
                Grilla.AddItem cabecera
                For k = 0 To 7
                    Grilla.col = k
                    Grilla.Row = 0
                    Grilla.CellBackColor = &HDFDFE0
                Next k
        End If
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            If rst("presentacion") > 1 Then
               in_nombre = rst("nombre_prod") & " : TRAE [" & rst("presentacion") & "]"
            Else
               in_nombre = rst("nombre_prod")
            End If
            
            
            If rst("id_igv") = "si" Then
               in_precio = rst("precio_venta")
            Else
               in_precio = rst("precio_venta")
            End If
            
            
            If KEY_CARGO = "00052" Or KEY_CARGO = "00008" Then
                in_precio_costo = "[***]"
            Else
                in_precio_costo = Format(rst("precio_compra"), "#,##0.00")
            End If
            
            If KEY_CONVERSION_CAMBIO = "si" Then
               in_precio_costo = Format(in_precio, "#,##0.00")
               in_precio = in_precio_costo * KEY_CAMBIO_LOCAL
            End If
           
            If IsNull(rst("vencimiento")) = True Then
                in_vencimiento = "SIN REGISTRO "
            Else
                in_vencimiento = Format(rst("vencimiento"), "dd-mm-YYYY")
            End If
            
            Fila = rst("id_producto") & vbTab & in_nombre & vbTab & UCase(rst("principio_activo")) & vbTab & rst("forma_farmacologica") & vbTab & in_vencimiento & vbTab & rst("unidad") & vbTab & in_precio_costo & vbTab & Format(in_precio, "#,##0.00")
            Grilla.AddItem Fila
            
            If Format(in_vencimiento, "YYYY-mm-dd") <= KEY_FECHA Then
                    Grilla.col = 4
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
             End If
            
            rst.MoveNext
    Next i
  Grilla.ColAlignment(1) = 1
  Grilla.ColAlignment(3) = 1
  Grilla.ColAlignment(5) = 1
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub


Public Sub llenarGrid_stock(ByVal Grilla As MSHFlexGrid, ByVal sql As String)
Dim in_precio As Single
On Error GoTo salir

Call ConfiguraRst(sql)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Me.cmdCaracteristicas.Enabled = False
    Me.cmdcompatibility.Enabled = False
    Me.cmdupdate.Enabled = False
    Me.cmdDelete.Enabled = False
    Exit Sub
End If
  
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       If KEY_MODELO_COLOR = "si" Then
            For Each Campo In rst.Fields
                Grilla.ColWidth(0) = 750
                Grilla.ColWidth(1) = 3500
                Grilla.ColWidth(2) = 1200
                Grilla.ColWidth(3) = 1200
                Grilla.ColWidth(4) = 1100
                Grilla.ColWidth(5) = 1000
                Grilla.ColWidth(6) = 1000
                Grilla.ColWidth(7) = 1000
                Grilla.ColWidth(8) = 1000
                Grilla.ColWidth(9) = 1000
                Grilla.ColWidth(10) = 1000
            Next
                cabecera = "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "CLASIFICACION" & vbTab & "MODELO" & vbTab & "COLOR" & vbTab & "UNIDAD" & vbTab & "MARCA" & vbTab & "P.MAYOR" & vbTab & "P.COSTO" & vbTab & "P.VENTA" & vbTab & "STOCK"
                Grilla.AddItem cabecera
                For k = 0 To 10
                    Grilla.col = k
                    Grilla.Row = 0
                    Grilla.CellBackColor = &HDFDFE0
                Next k
        Else
            For Each Campo In rst.Fields
                Grilla.ColWidth(0) = 700
                Grilla.ColWidth(1) = 5500
                Grilla.ColWidth(2) = 1200
                Grilla.ColWidth(3) = 700
                Grilla.ColWidth(4) = 1100
                Grilla.ColWidth(5) = 1000
                Grilla.ColWidth(6) = 1000
                Grilla.ColWidth(7) = 1000
                Grilla.ColWidth(8) = 1000
            Next
                cabecera = "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "CLASIFICACION" & vbTab & "UND" & vbTab & "MARCA" & vbTab & "P.MAYOR" & vbTab & "P.COSTO" & vbTab & "P.VENTA" & vbTab & "STOCK"
                Grilla.AddItem cabecera
                For k = 0 To 8
                    Grilla.col = k
                    Grilla.Row = 0
                    Grilla.CellBackColor = &HDFDFE0
                Next k
        End If
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            If rst("presentacion") > 1 Then
               in_nombre = rst("nombre_prod") & " : TRAE [" & rst("presentacion") & "]"
            Else
               in_nombre = rst("nombre_prod")
            End If
            
            
            If rst("id_igv") = "si" Then
               in_precio = rst("precio_venta")
            Else
               in_precio = rst("precio_venta")
            End If
            If KEY_MODELO_COLOR = "si" Then
                Fila = rst("id_producto") & vbTab & in_nombre & vbTab & UCase(rst("linea")) & vbTab & UCase(rst("modelo")) & vbTab & rst("color") & vbTab & rst("unidad") & vbTab & rst("marca") & vbTab & Format(rst("precio_mayor"), "#,##0.00") & vbTab & Format(rst("precio_compra"), "#,##0.00") & vbTab & Format(in_precio, "#,##0.00") & vbTab & Format(rst("stock"), "#,##0.00")
            Else
                Fila = rst("id_producto") & vbTab & in_nombre & vbTab & UCase(rst("linea")) & vbTab & rst("unidad") & vbTab & rst("marca") & vbTab & Format(rst("precio_mayor"), "#,##0.00") & vbTab & Format(rst("precio_compra"), "#,##0.00") & vbTab & Format(in_precio, "#,##0.00") & vbTab & Format(rst("stock"), "#,##0.00")
            End If
            Grilla.AddItem Fila
                
                Grilla.col = 8
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H80FF&
                
            rst.MoveNext
    Next i
  Grilla.ColAlignment(3) = 1
  Grilla.ColAlignment(5) = 1
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub

Public Sub llenarGridColor(ByVal Grilla As MSHFlexGrid, ByVal sql As String)



On Error GoTo salir
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
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 5200
           Grilla.ColWidth(2) = 700
           Grilla.ColWidth(3) = 900
           Grilla.ColWidth(4) = 1300
           Grilla.ColWidth(5) = 1300
         Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION ARTICULO" & vbTab & "UND" & vbTab & "STOCK" & vbTab & "P.COSTO" & vbTab & "P.VENTA"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & Format(rst("stock"), "#,##0.00") & vbTab & Format(rst("precio_compra"), "#,##0.00") & vbTab & Format(rst("precio_venta"), "#,##0.00")
            If (Fila = "") Then
                X = 1
            End If
          Grilla.AddItem Fila
                        
                        If (Trim(rst("Stock")) < 2) Then
                            For k = 0 To 5
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                            Next k
                      End If
            Fila = ""
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

Public Sub LlenarGrid_Factura(ByVal Grilla As MSHFlexGrid, ByVal sql As String)

On Error GoTo salir
Call ConfiguraRst(sql)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 4000
           Grilla.ColWidth(2) = 1900
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 1100
           Grilla.ColWidth(5) = 700
           Grilla.ColWidth(6) = 1000
           Grilla.ColWidth(7) = 1000
           Grilla.ColWidth(8) = 1000
         Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "CLASIFICACION" & vbTab & "UNIDAD" & vbTab & "MARCA" & vbTab & "ST.FISICO" & vbTab & "ST.CONT" & vbTab & "P.COSTO" & vbTab & "P.VENTA"
        Grilla.AddItem cabecera
         For k = 0 To 8
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If KEY_RUC = "20561193291" Or KEY_RUC = "20479798894" Then
                in_costo = rst("precio_compra") * (1 + KEY_IGV)
            Else
                in_costo = rst("precio_compra")
            End If
            
            Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & UCase(rst("linea")) & vbTab & rst("unidad") & vbTab & rst("marca") & vbTab & rst("stock") & vbTab & rst("stock_factura") & vbTab & Format(in_costo, "#,##0.00") & vbTab & Format(rst("precio_venta"), "#,##0.00")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
    Next i
  Grilla.ColAlignment(1) = 1
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub



Public Sub LlenarGridcolorFactura(ByVal Grilla As MSHFlexGrid, ByVal sql As String)
On Error GoTo salir


Call ConfiguraRst(sql)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 4800
           Grilla.ColWidth(2) = 700
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1000
           Grilla.ColWidth(5) = 1000
           Grilla.ColWidth(6) = 1000
         Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION ARTICULO" & vbTab & "UND" & vbTab & "STOCK" & vbTab & "S-FACTURA" & vbTab & "P.VENTA" & vbTab & "P.COSTO"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & rst("stock") & vbTab & rst("stock_factura") & vbTab & rst("precio_compra") & vbTab & rst("precio_venta")
           
          Grilla.AddItem Fila
                        
                        If (Trim(rst("stock")) < 2) Then
                            For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                            Next k
                      End If
            Fila = ""
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
Public Sub actualizar()
    Call ActualizarProd
    'Call ActualizarAlm
End Sub
Public Sub ActualizarProd()
If KEY_ALM = "" Then
   KEY_ALM = "00001"
End If
 
 If KEY_CARGO = "00008" Or KEY_CARGO = "00052" Then
    If KEY_STOCK_GLOBAL = "no" Then
        strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' and cta_contable='0' ORDER BY nombre_prod LIMIT 30 "
    Else
        strCadena = "SELECT * FROM view_producto_stock WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' and cta_contable='0' ORDER BY nombre_prod LIMIT 30 "
    End If
 Else
    If KEY_STOCK_GLOBAL = "no" Then
        strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY nombre_prod LIMIT 30 "
    Else
        strCadena = "SELECT * FROM view_producto_stock WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY nombre_prod LIMIT 30 "
    End If
 End If
 
  If KEY_SKFACTURA = "no" Then
        If KEY_STOCK_GLOBAL = "no" Then
            Call llenarGrid(Me.HfdGrilla, strCadena)
        Else
            Call llenarGrid_stock(Me.HfdGrilla, strCadena)
        End If
        Exit Sub
 Else
        Call LlenarGrid_Factura(Me.HfdGrilla, strCadena)
        Exit Sub
 End If

End Sub
Public Sub Actualizar_farmacia()
If KEY_ALM = "" Then
   KEY_ALM = "00001"
End If
 
 If KEY_CARGO = "00008" Or KEY_CARGO = "00052" Then
    If KEY_STOCK_GLOBAL = "no" Then
        strCadena = "SELECT * FROM view_producto_farmacia WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' and cta_contable='0' ORDER BY nombre_prod LIMIT 30 "
    Else
        strCadena = "SELECT * FROM view_producto_stock WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' and cta_contable='0' ORDER BY nombre_prod LIMIT 30 "
    End If
 Else
    If KEY_STOCK_GLOBAL = "no" Then
        strCadena = "SELECT * FROM view_producto_farmacia WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY nombre_prod LIMIT 30 "
    Else
        strCadena = "SELECT * FROM view_producto_stock WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY nombre_prod LIMIT 30 "
    End If
 End If
 
  If KEY_SKFACTURA = "no" Then
        If KEY_STOCK_GLOBAL = "no" Then
            Call llenarGrid_farmacia(Me.HfdGrilla, strCadena)
        Else
            Call llenarGrid_stock(Me.HfdGrilla, strCadena)
        End If
        Exit Sub
 Else
        Call LlenarGrid_Factura(Me.HfdGrilla, strCadena)
        Exit Sub
 End If

End Sub

Public Sub actualizar_update(ByVal in_producto As String)
 strCadena = "SELECT * FROM view_producto WHERE id_producto='" & in_producto & "' and  ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
 
 
  If KEY_SKFACTURA = "no" Then
        Call llenarGrid(Me.HfdGrilla, strCadena)
        Exit Sub
 Else
        Call LlenarGrid_Factura(Me.HfdGrilla, strCadena)
        Exit Sub
 End If

End Sub


Public Sub llenar_almacen_contable(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
On Error GoTo salir
Dim in_stock As Single
If rstP.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstP.Fields.Count)
       For Each Campo In rstP.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2000
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1000
           
         Next
        cabecera = "CODIGO" & vbTab & "ALMACEN" & vbTab & "CONTABLE" & vbTab & "NO CONTAB" & vbTab & "TOTAL "
        Grilla.AddItem cabecera
         For k = 1 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstP.MoveFirst
        in_stock = 0
        For i = 0 To rstP.RecordCount - 1
            Fila = rstP("id_alm") & vbTab & rstP("descripcion") & vbTab & rstP("stock") & vbTab & rstP("stock_no_contable") & vbTab & rstP("stock") + rstP("stock_no_contable")
            Grilla.AddItem Fila
            in_stock = in_stock + rstP("stock") + rstP("stock_contable")
            If rstP("id_alm") = KEY_ALM Then
               For k = 1 To 4
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
                Next k
                 Me.TxtSector.Text = rstP("sector")
                 Me.TxtPiso.Text = rstP("piso")
                 Me.TxtAndamio.Text = rstP("andamio")
                 Me.txt_x.Text = rstP("casillero_x")
                 Me.Txt_y.Text = rstP("casillero_y")
  
            End If
            rstP.MoveNext
    Next i
    Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & in_stock
    Grilla.AddItem Fila
    
                Grilla.col = 4
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H80FF&
    
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rstP = Nothing

End Sub
Public Sub ActualizarAlm(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
On Error GoTo salir
Dim in_stock As Single

If KEY_RESERVA_STOCK = "si" Then
    strCadena = "call ADM_almacen_producto('1','" & id_producto & "','" & KEY_FECHA & "','" & KEY_ALM & "','" & KEY_RUC & "')"
Else
    strCadena = "call ADM_almacen_producto('2','" & id_producto & "','" & KEY_FECHA & "','" & KEY_ALM & "','" & KEY_RUC & "')"
End If

Call ConfiguraRstP(strCadena)
If rstP.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstP.Fields.Count)
       For Each Campo In rstP.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2000
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1000
           
         Next
        cabecera = "CODIGO" & vbTab & "ALMACEN" & vbTab & "STOCK " & vbTab & "PENDIENTE " & vbTab & "TOTAL "
        Grilla.AddItem cabecera
         For k = 1 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstP.MoveFirst
        in_stock = 0
        For i = 0 To rstP.RecordCount - 1
            
            'strCadena = "CALL ADM_stock_separado('1','" & rstP("id_producto") & "','" & rstP("id_alm") & "','" & KEY_RUC & "')"
            'Call ConfiguraRstL(strCadena)
            
            Fila = rstP("id_alm") & vbTab & rstP("descripcion") & vbTab & rstP("stock") & vbTab & rstP("_separado") & vbTab & Val(rstP("stock") - rstP("_separado"))
            Grilla.AddItem Fila
            in_stock = in_stock + rstP("stock") - rstP("_separado")
            If rstP("id_alm") = KEY_ALM Then
               For k = 1 To 4
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
                Next k
                 
                 
                 If KEY_ALM = rstP("id_alm") Then
                   ' Me.TxtSector.Text = rstP("sector")
                   ' Me.TxtPiso.Text = rstP("piso")
                   ' Me.TxtAndamio.Text = rstP("andamio")
                   ' Me.txt_x.Text = rstP("casillero_x")
                   ' Me.Txt_y.Text = rstP("casillero_y")
                   ' Me.txtcodigo_barra.Text = rstP("codigo_barra")
                   ' Me.txtFormafarmacologica.Text = rstP("forma_farmacologica")
                    
                End If
                
  
            End If
            rstP.MoveNext
    Next i
    Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & in_stock
    Grilla.AddItem Fila
    
                Grilla.col = 4
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H80FF&
    
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rstP = Nothing

End Sub


Public Sub load_compatibility(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
On Error GoTo salir
strCadena = "SELECT * FROM view_compatibilidad WHERE id_padre='" & Trim(id_producto) & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 4000
           Grilla.ColWidth(2) = 1500
           
           
         Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "STOCK "
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("stock")
            Grilla.AddItem Fila
             
            rst.MoveNext
    Next i
        
     
        
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rstP = Nothing

End Sub

Public Sub load_caracteristicas(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
On Error GoTo salir
strCadena = "SELECT * FROM producto_caracteristicas WHERE id_producto='" & Trim(id_producto) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 5500
           
       Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION"
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_producto") & vbTab & rst("caracteristica")
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
        
     
        
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"


End Sub

Public Sub load_grupo_empresarial(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
On Error GoTo salir
strCadena = "SELECT id_producto,nombre_completo,sum(stock) as stock FROM view_producto_grupo_empresarial WHERE id_producto='" & Trim(id_producto) & "' and ruc='" & KEY_RUC & "' GROUP BY id_producto,ruc_vinculado"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
   
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 3200
           Grilla.ColWidth(2) = 1000
       Next
        cabecera = "CODIGO" & vbTab & "EMPRESA" & vbTab & "STOCK"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_producto") & vbTab & rst("nombre_completo") & vbTab & rst("stock")
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
        
     
        
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"


End Sub

Sub llenarGrid_alm(ByVal Grilla As MSHFlexGrid, ByVal sql As String)

  Call ConfiguraRst(sql)
  Grilla.Clear
  Grilla.Rows = rst.RecordCount
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 0
  Grilla.ColWidth(1) = 3900
  Grilla.Enabled = False
Grilla.Refresh

End Sub

Private Sub Image2_Click()
Me.framemayor.Visible = False
End Sub

Private Sub txtBuscar_Change()

  strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' AND descripcion like '%" & Trim(Me.txtBuscar.Text) & "%' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLinea)
  
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcLinea.SetFocus
End If
End Sub

Private Sub TxtCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If KEY_RUBRO = "00003" Then
        Call buscar_codigo_farmacia
    Else
        Call buscar_codigo
    End If
End If
End Sub
Private Sub buscar_codigo()

Dim Registros As Integer
Dim Criterio As String
  If Me.TxtCod.Text = "" Then
    Call ActualizarProd
    Exit Sub
  Else
  If Len(Me.TxtCod.Text) > 0 Then
  If KEY_BARRAS = "no" Then
   
        Criterio = "id_producto LIKE '%" & Trim(Me.TxtCod.Text) & "%'"
    
    
  Else
 
    
    Criterio = "(cod_barra LIKE '%" & Trim(Me.TxtCod.Text) & "%' or id_producto LIKE  '%" & Trim(Me.TxtCod.Text) & "%' or  codigo_proveedor LIKE '%" & Trim(Me.TxtCod.Text) & "%' or  id_universal LIKE '%" & Trim(Me.TxtCod.Text) & "%' or  codigo_alterno LIKE '%" & Trim(Me.TxtCod.Text) & "%')"
  End If
  
   If KEY_SKFACTURA = "no" Then
      If KEY_BARRAS = "si" Then
        strCadena = "SELECT * FROM view_producto_barras WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY nombre_prod LIMIT 150 "
        Call llenarGrid(Me.HfdGrilla, strCadena)
     Else
        
        strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY nombre_prod LIMIT 150 "
        Call llenarGrid(Me.HfdGrilla, strCadena)
    End If
   
    Else
      If KEY_BARRAS = "si" Then
    
    strCadena = "SELECT * FROM view_producto_barras WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY nombre_prod LIMIT 150 "
        
    
    'strCadena = "SELECT A.id_producto,P.nombre_prod,U.abreviatura,A.stock,A.stock_factura,P.precio_compra,P.precio_venta FROM almacen_producto A,producto P,unidad U,producto_barras B WHERE A.id_producto=P.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "'" & _
    " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND A.id_alm='" & KEY_ALM & "' AND A.id_producto=B.id_producto AND B.ruc='" & KEY_RUC & "' AND P.id_producto=B.id_producto AND  " & Criterio & "ORDER BY nombre_prod LIMIT 150"
 
    Else
        strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY nombre_prod LIMIT 150 "
    End If
    Call LlenarGrid_Factura(Me.HfdGrilla, strCadena)
  End If
End If
  
    Me.TxtCod.SetFocus
  End If


End Sub

Private Sub buscar_codigo_farmacia()

Dim Registros As Integer
Dim Criterio As String
  If Me.TxtCod.Text = "" Then
    Call ActualizarProd
    Exit Sub
  Else
  If Len(Me.TxtCod.Text) > 0 Then
  If KEY_BARRAS = "no" Then
   
        Criterio = "(codigo_barra LIKE '%" & Trim(Me.TxtCod.Text) & "%' or  id_producto LIKE '%" & Trim(Me.TxtCod.Text) & "%')"
    
    
  Else
 
    
    Criterio = "(cod_barra LIKE '%" & Trim(Me.TxtCod.Text) & "%' or id_producto LIKE  '%" & Trim(Me.TxtCod.Text) & "%' or  codigo_proveedor LIKE '%" & Trim(Me.TxtCod.Text) & "%' or  id_universal LIKE '%" & Trim(Me.TxtCod.Text) & "%' or  codigo_alterno LIKE '%" & Trim(Me.TxtCod.Text) & "%')"
  End If
  
   If KEY_SKFACTURA = "no" Then
      If KEY_BARRAS = "si" Then
        strCadena = "SELECT * FROM view_producto_barras WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY nombre_prod LIMIT 150 "
        Call llenarGrid(Me.HfdGrilla, strCadena)
      Else
        
        strCadena = "SELECT * FROM view_producto_farmacia WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY nombre_prod LIMIT 150 "
        Call llenarGrid_farmacia(Me.HfdGrilla, strCadena)
    End If
   
    Else
      If KEY_BARRAS = "si" Then
    
    strCadena = "SELECT * FROM view_producto_barras WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY nombre_prod LIMIT 150 "
        
    
    'strCadena = "SELECT A.id_producto,P.nombre_prod,U.abreviatura,A.stock,A.stock_factura,P.precio_compra,P.precio_venta FROM almacen_producto A,producto P,unidad U,producto_barras B WHERE A.id_producto=P.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "'" & _
    " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND A.id_alm='" & KEY_ALM & "' AND A.id_producto=B.id_producto AND B.ruc='" & KEY_RUC & "' AND P.id_producto=B.id_producto AND  " & Criterio & "ORDER BY nombre_prod LIMIT 150"
 
    Else
        strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY nombre_prod LIMIT 150 "
    End If
    Call LlenarGrid_Factura(Me.HfdGrilla, strCadena)
  End If
End If
  
    Me.TxtCod.SetFocus
  End If


End Sub

Private Sub txtcuenta_contable_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Criterio = " cta_contable LIKE '%" & Trim(Me.txtcuenta_contable.Text) & "%'"
    strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY nombre_prod LIMIT 0,50 "
    Call llenarGrid(Me.HfdGrilla, strCadena)
End If
End Sub

Private Sub txtforma_farmacologica_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    strLinea = False
    If KEY_RUBRO = "00003" Then
    Call busqueda_farmacia_farmacologica
    Else
     Call busqueda
    End If
End If

End Sub

Private Sub TxtMarca_Change()
  
  strCadena = "SELECT id_marca as Codigo, descripcion as Descripcion FROM marca WHERE id_usu='" & KEY_RUC & "' AND descripcion like '%" & Trim(Me.TxtMarca.Text) & "%' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMarca)

End Sub

Private Sub txtprincipioActivo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strLinea = False
    If KEY_RUBRO = "00003" Then
        Call busqueda_farmacia_principio
     End If
End If
End Sub

Private Sub TxtProducto_Change()
'Me.chkStockBajo.Value = 0

End Sub
Public Sub busqueda()
Dim parametros() As String
Dim Criterio As String

                StrNombre = InStr(1, Trim(Me.Txtproducto.Text), "'", vbTextCompare)
                If Val(StrNombre) > 0 Then
                   MsgBox "Caracter Incorrecto Gracias", vbInformation, "No Dijite este Signo"
                   Exit Sub
                End If
                
                    parametros = Split(Trim(Me.Txtproducto.Text), " ")
                    Criterio = ""
                    For i = 0 To UBound(parametros)
                        If Criterio <> "" Then
                            Criterio = Trim(Criterio & "%" & Trim(parametros(i)))
                        Else
                            Criterio = Trim(parametros(i))
                        End If
                        
                    Next i
                 
                StrNombre = InStr(1, Trim(Me.Txtproducto.Text), "'", vbTextCompare)
                If Val(StrNombre) > 0 Then
                   MsgBox "Caracter Incorrecto Gracias", vbInformation, "No Dijite este Signo"
                   Exit Sub
                End If
    
 
 
   strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  nombre_prod LIKE '%" & Criterio & "%' ORDER BY nombre_prod LIMIT 500"
    If KEY_SKFACTURA = "no" Then
        Call llenarGrid(Me.HfdGrilla, strCadena)
    Else
        Call LlenarGrid_Factura(Me.HfdGrilla, strCadena)
    End If
  
  Me.Txtproducto.SetFocus
End Sub

Public Sub busqueda_farmacia()
Dim parametros() As String
Dim Criterio As String

                StrNombre = InStr(1, Trim(Me.Txtproducto.Text), "'", vbTextCompare)
                If Val(StrNombre) > 0 Then
                   MsgBox "Caracter Incorrecto Gracias", vbInformation, "No Dijite este Signo"
                   Exit Sub
                End If
                
                    parametros = Split(Trim(Me.Txtproducto.Text), " ")
                    Criterio = ""
                    For i = 0 To UBound(parametros)
                        If Criterio <> "" Then
                            Criterio = Trim(Criterio & "%" & Trim(parametros(i)))
                        Else
                            Criterio = Trim(parametros(i))
                        End If
                        
                    Next i
                 
                StrNombre = InStr(1, Trim(Me.Txtproducto.Text), "'", vbTextCompare)
                If Val(StrNombre) > 0 Then
                   MsgBox "Caracter Incorrecto Gracias", vbInformation, "No Dijite este Signo"
                   Exit Sub
                End If
    
 
 
    strCadena = "SELECT * FROM view_producto_farmacia WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  nombre_producto LIKE '%" & Criterio & "%' ORDER BY nombre_prod LIMIT 500"
    If KEY_SKFACTURA = "no" Then
        Call llenarGrid_farmacia(Me.HfdGrilla, strCadena)
    Else
        Call LlenarGrid_Factura(Me.HfdGrilla, strCadena)
    End If
  
  Me.Txtproducto.SetFocus
End Sub
Public Sub busqueda_farmacia_farmacologica()
Dim parametros() As String
Dim Criterio As String

                StrNombre = InStr(1, Trim(Me.txtforma_farmacologica.Text), "'", vbTextCompare)
                If Val(StrNombre) > 0 Then
                   MsgBox "Caracter Incorrecto Gracias", vbInformation, "No Dijite este Signo"
                   Exit Sub
                End If
                
                    parametros = Split(Trim(Me.txtforma_farmacologica.Text), " ")
                    Criterio = ""
                    For i = 0 To UBound(parametros)
                        If Criterio <> "" Then
                            Criterio = Trim(Criterio & "%" & Trim(parametros(i)))
                        Else
                            Criterio = Trim(parametros(i))
                        End If
                        
                    Next i
                 
                StrNombre = InStr(1, Trim(Me.txtforma_farmacologica.Text), "'", vbTextCompare)
                If Val(StrNombre) > 0 Then
                   MsgBox "Caracter Incorrecto Gracias", vbInformation, "No Dijite este Signo"
                   Exit Sub
                End If
    
 
 
    strCadena = "SELECT * FROM view_producto_farmacia WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  forma_farmacologica LIKE '%" & Criterio & "%' ORDER BY nombre_prod LIMIT 500"
    If KEY_SKFACTURA = "no" Then
        Call llenarGrid_farmacia(Me.HfdGrilla, strCadena)
    Else
        Call LlenarGrid_Factura(Me.HfdGrilla, strCadena)
    End If
  
  Me.txtforma_farmacologica.SetFocus
End Sub
Public Sub busqueda_farmacia_principio()
Dim parametros() As String
Dim Criterio As String

                StrNombre = InStr(1, Trim(Me.txtprincipioActivo.Text), "'", vbTextCompare)
                If Val(StrNombre) > 0 Then
                   MsgBox "Caracter Incorrecto Gracias", vbInformation, "No Dijite este Signo"
                   Exit Sub
                End If
                
                    parametros = Split(Trim(Me.txtprincipioActivo.Text), " ")
                    Criterio = ""
                    For i = 0 To UBound(parametros)
                        If Criterio <> "" Then
                            Criterio = Trim(Criterio & "%" & Trim(parametros(i)))
                        Else
                            Criterio = Trim(parametros(i))
                        End If
                        
                    Next i
                 
                StrNombre = InStr(1, Trim(Me.txtprincipioActivo.Text), "'", vbTextCompare)
                If Val(StrNombre) > 0 Then
                   MsgBox "Caracter Incorrecto Gracias", vbInformation, "No Dijite este Signo"
                   Exit Sub
                End If
    
 
 
    strCadena = "SELECT * FROM view_producto_farmacia WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  principio_activo LIKE '%" & Criterio & "%' ORDER BY nombre_prod LIMIT 500"
    If KEY_SKFACTURA = "no" Then
        Call llenarGrid_farmacia(Me.HfdGrilla, strCadena)
    Else
        Call LlenarGrid_Factura(Me.HfdGrilla, strCadena)
    End If
  
  Me.txtprincipioActivo.SetFocus
End Sub


Private Sub TxtProducto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        
        Me.HfdGrilla.SetFocus
    End If
End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    strLinea = False
    If KEY_RUBRO = "00003" Then
        Call busqueda_farmacia
    Else
        Call busqueda
    End If
End If
End Sub


Private Sub Resalta(ByVal Texto As TextBox)
On Error GoTo Saltar
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
Saltar:
Exit Sub
End Sub

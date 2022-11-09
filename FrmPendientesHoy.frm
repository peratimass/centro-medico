VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmPendientesHoy 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   17595
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "CERRAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15840
      TabIndex        =   29
      Top             =   170
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2535
      Left            =   8160
      TabIndex        =   19
      Top             =   3120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4471
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "SOLICITUDES"
      TabPicture(0)   =   "FrmPendientesHoy.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "HfgSolicitud"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "PRESTAMOS A TERCEROS"
      TabPicture(1)   =   "FrmPendientesHoy.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSHFlexGrid3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "PAGO PROVEEDORES"
      TabPicture(2)   =   "FrmPendientesHoy.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "HfgPagar"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgSolicitud 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   23
         Top             =   480
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3413
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   24
         Top             =   480
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3413
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgPagar 
         Height          =   1935
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3413
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
   Begin VB.Frame FrmServiciotecnico 
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
      Height          =   9015
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   7695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgStockMinimo 
         Height          =   1695
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2990
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgVencerce 
         Height          =   1695
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2990
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgTransferencias 
         Height          =   1215
         Left            =   240
         TabIndex        =   13
         Top             =   5640
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2143
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgPedidos 
         Height          =   1335
         Left            =   240
         TabIndex        =   16
         Top             =   7560
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2355
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "PEDIDOS PENDIENTES DE REALIZAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   840
         TabIndex        =   15
         Top             =   7080
         Width           =   3975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   6720
         TabIndex        =   14
         Top             =   7080
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   255
         Left            =   360
         Picture         =   "FrmPendientesHoy.frx":0054
         Stretch         =   -1  'True
         Top             =   7080
         Width           =   255
      End
      Begin VB.Image Image7 
         Height          =   255
         Left            =   360
         Picture         =   "FrmPendientesHoy.frx":05DE
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   255
      End
      Begin VB.Image Image6 
         Height          =   255
         Left            =   360
         Picture         =   "FrmPendientesHoy.frx":0B68
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   255
      End
      Begin VB.Image Image2 
         Height          =   255
         Left            =   360
         Picture         =   "FrmPendientesHoy.frx":10F2
         Stretch         =   -1  'True
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblcomprobantescobrar 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   6720
         TabIndex        =   10
         Top             =   5160
         Width           =   375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSFERENCIAS ENTRE ESTABLECIMIENTOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   840
         TabIndex        =   9
         Top             =   5160
         Width           =   3975
      End
      Begin VB.Shape ShapeInspecciones 
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         Height          =   495
         Left            =   240
         Top             =   5040
         Width           =   7215
      End
      Begin VB.Label lblvencerse 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   6720
         TabIndex        =   6
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label lblstockminimo 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   6720
         TabIndex        =   5
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCTOS POR VENCERCE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   840
         TabIndex        =   4
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Shape ShapeAnexos 
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         Height          =   495
         Left            =   240
         Top             =   2640
         Width           =   7215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCTOS CON STOCK MINIMO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
      Begin VB.Shape ShapeInstalaciones 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         Height          =   495
         Left            =   240
         Top             =   240
         Width           =   7215
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         Height          =   495
         Left            =   240
         Top             =   6960
         Width           =   7215
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2775
      Left            =   8160
      TabIndex        =   22
      Top             =   6240
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CUENTAS X COBRAR"
      TabPicture(0)   =   "FrmPendientesHoy.frx":167C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "HfgCuentasCobrar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "LETRAS DE CAMBIO"
      TabPicture(1)   =   "FrmPendientesHoy.frx":1698
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSHFlexGrid6"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgCuentasCobrar 
         Height          =   2175
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3836
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid6 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3413
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
   End
   Begin VB.Label lblcargo 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   210
      Left            =   11280
      TabIndex        =   28
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Image Image8 
      Height          =   255
      Left            =   10860
      Picture         =   "FrmPendientesHoy.frx":16B4
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPROBANTES X COBRAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8760
      TabIndex        =   21
      Top             =   5820
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   14640
      TabIndex        =   20
      Top             =   5820
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   8280
      Picture         =   "FrmPendientesHoy.frx":1C3E
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPROBANTES X PAGAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8760
      TabIndex        =   18
      Top             =   2700
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   14640
      TabIndex        =   17
      Top             =   2700
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   8280
      Picture         =   "FrmPendientesHoy.frx":21C8
      Stretch         =   -1  'True
      Top             =   2700
      Width           =   255
   End
   Begin VB.Image Image11 
      Height          =   255
      Left            =   10860
      Picture         =   "FrmPendientesHoy.frx":2752
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   255
   End
   Begin VB.Image Image10 
      Height          =   255
      Left            =   10860
      Picture         =   "FrmPendientesHoy.frx":2CDC
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   255
   End
   Begin VB.Image Image9 
      Height          =   255
      Left            =   10860
      Picture         =   "FrmPendientesHoy.frx":3266
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblDni 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   210
      Left            =   11280
      TabIndex        =   8
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   10800
      TabIndex        =   7
      Top             =   240
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   8280
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblfecha 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   210
      Left            =   11280
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label lblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Height          =   210
      Left            =   11280
      TabIndex        =   0
      Top             =   840
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   2415
      Left            =   8160
      Top             =   120
      Width           =   9255
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   345
      Left            =   8160
      Top             =   2640
      Width           =   9255
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   345
      Left            =   8160
      Top             =   5760
      Width           =   7815
   End
End
Attribute VB_Name = "FrmPendientesHoy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.Label6.Caption = "BIENVENIDO A" + Space(1) + KEY_EMPRESA
Me.lblUsuario.Caption = BDBuscarCampo("persona", "nombre_completo", "dni", KEY_USUARIO)
Me.lblcargo.Caption = "CARGO:" + Space(1) + BDBuscarCampo("persona_cargos", "descripcion", "id_cargo", KEY_CARGO)
Me.LblFecha.Caption = "FECHA :" + Space(1) + KEY_FECHA
Me.lbldni.Caption = "DNI" + Space(1) + KEY_USUARIO




strCadena = "SELECT * FROM persona WHERE dni='" & KEY_USUARIO & "'"
Call ConfiguraRst(strCadena)
On Error GoTo sli
If IsNull(rst("foto")) = False And Len(rst("foto")) > 5 Then
            If VerificarFichero(App.Path & "\archivos\" & rst("dni")) = True Then
               Me.Image1.Picture = LoadPicture(App.Path + "\archivos\" + rst("dni") + "\" + Trim(rst("foto")))
            Else
                Me.Image1.Picture = LoadPicture(App.Path + "\archivos\no_photo.jpg")
            End If
        Else
            Me.Image1.Picture = LoadPicture(App.Path + "\archivos\no_photo.jpg")
        End If
sli:
Call ProductosStockMinimo(Me.HfgStockMinimo)
Call vencidos(Me.HfgVencerce)
Call Transferencia(Me.HfgTransferencias)
Call pedidos(Me.HfgPedidos, Me)
Call CuentasCobrar(Me.HfgCuentasCobrar)
Call CuentasPagar(Me.HfgPagar, Me)
End Sub
Private Sub llenar_solicitudes(ByVal Grilla As MSHFlexGrid, ByVal formulario As Form)
'On Error GoTo salir
Dim tTotal As Double, ccostos As String
strCadena = "SELECT * FROM solicitud_dinero S,persona P WHERE S.dni=P.dni AND S.ruc='" & KEY_RUC & "' AND S.anulado='no' AND (atendido='no' OR finalizado='no') ORDER BY id_solicitud"
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
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 1100
           Grilla.ColWidth(3) = 3000
           Grilla.ColWidth(4) = 3000
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 1200
           Grilla.ColWidth(8) = 0
           Grilla.ColWidth(9) = 0
           Grilla.ColWidth(10) = 0
        Next
        cabecera = "IDSOLICITUD" & vbTab & "NºSOLICITUD" & vbTab & "FECHA" & vbTab & "SOLICITANTE" & vbTab & "MOTIVO" & vbTab & " MONTO" & vbTab & " SALDO" & vbTab & "DECLARADO" & vbTab & "atendido" & vbTab & "finalizado" & vbTab & "dni"
        Grilla.AddItem cabecera
         For k = 0 To 7
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            
          strCadena = "SELECT sum(monto) FROM solicitud_dinero_declarar WHERE id_solicitud='" & rst("id_solicitud") & "' AND ruc='" & KEY_RUC & "'"
          Call ConfiguraRstT(strCadena)
          If IsNull(rstT(0)) = True Then
            declarado = 0
          Else
            declarado = rstT(0)
           End If
          Fila = rst("id_solicitud") & vbTab & rst("numero") & vbTab & rst("fecha_solicitud") & vbTab & rst("nombre_completo") & vbTab & rst("resumen") & vbTab & Format(rst("monto_solicitado"), "#,##0.00") & vbTab & Format(rst("saldo"), "#,##0.00") & vbTab & Format(declarado, "#,##0.00") & vbTab & rst("atendido") & vbTab & rst("finalizado") & vbTab & rst("dni")
          Grilla.AddItem Fila
          If rst("finalizado") = "si" Then
                        For k = 0 To 7
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H80C0FF
                        Next k
          End If
          If rst("atendido") = "no" Then
                        For k = 0 To 7
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                        Next k
          End If
          
        Fila = ""
        rst.MoveNext
             
        Next i
  End Sub


Private Sub CuentasPagar(ByVal Grilla As MSHFlexGrid, ByVal formulario As Form)
On Error GoTo SALIR
strCadena = "SELECT MC.id_compra,MC.fecha_emision,MC.fecha_cancelacion,CONCAT(C.doc_abrev,':',MC.serie,'-',MC.numero) as comprobante,PR.nombre_completo as nproveedor,MC.tc,MC.total,MC.saldo,P.nombre_completo,M.simbolo,MC.id_moneda,MC.id_proveedor FROM movimiento_compra MC,comprobantes C,persona P,moneda M,persona PR WHERE MC.id_proveedor=PR.dni AND MC.id_doc=C.id_doc AND MC.ruc='" & KEY_RUC & "' AND MC.id_moneda=M.id_moneda AND MC.dni_save=P.dni AND MC.saldo>0 AND MC.anulado='no' ORDER BY MC.fecha_emision ASC LIMIT 0,50  "
Dim tTotal As Double, tSaldo As Double, nsaldo As Double
tTotal = 0
tSaldo = 0
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
  
  N = 1
  
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 2000
           Grilla.ColWidth(4) = 2500
           Grilla.ColWidth(5) = 250
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 1300
           Grilla.ColWidth(8) = 900
           Grilla.ColWidth(9) = 1400
           Grilla.ColWidth(10) = 0
           
        Next
        cabecera = "IDCOMPRA" & vbTab & "EMISION" & vbTab & "VENCIMIENTO" & vbTab & "COMPROBANTE" & vbTab & "PROVEEDOR" & vbTab & "" & vbTab & "FACTURADO" & vbTab & "SALDO  (S/.)" & vbTab & "TC" & vbTab & "VENDEDOR" & vbTab & "IDPROVEEDOR"
        Grilla.AddItem cabecera
         For k = 0 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If rst("id_moneda") = "00002" Then
                nsaldo = rst("saldo") * rst("tc")
            Else
                nsaldo = rst("saldo")
            End If
            tSaldo = tSaldo + nsaldo
            Fila = rst("id_compra") & vbTab & rst("fecha_emision") & vbTab & rst("fecha_cancelacion") & vbTab & rst("comprobante") & vbTab & UCase(rst("nproveedor")) & vbTab & rst("simbolo") & vbTab & Format(rst("total"), "#,##0.00") & vbTab & Format(nsaldo, "#,##0.00") & vbTab & Format(rst("tc"), "#,##0.000") & vbTab & rst("nombre_completo") & vbTab & rst("id_proveedor")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
        Next i
        cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "**** S/. ****" & vbTab & Format(tSaldo, "#,##0.00") & vbTab & "************" & vbTab & "********************"
        Grilla.AddItem cabecera
          For k = 6 To 9
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
                            Next k
    
    
    
  Exit Sub
SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub

Public Sub CuentasCobrar(ByVal Grilla As MSHFlexGrid)
On Error GoTo SALIR
Dim tTotal As Double
strCadena = "SELECT M.id_venta,M.fecha_emision,M.fecha_vencimiento,C.doc_abrev,M.serie,M.numero,P.nombre_completo,M.total,M.saldo FROM movimiento_venta M,comprobantes C,persona P WHERE M.id_doc=C.id_doc AND M.id_cliente=P.dni AND M.ruc='" & KEY_RUC & "' AND anulado='no' AND id_forma_pago='02' AND M.saldo>0   ORDER BY id_venta ASC "
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
           Grilla.ColWidth(1) = 1150
           Grilla.ColWidth(2) = 1150
           Grilla.ColWidth(3) = 1800
           Grilla.ColWidth(4) = 2500
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1200
         Next
        cabecera = "CODIGO" & vbTab & "EMISION" & vbTab & "VENCIMIENTO" & vbTab & "COMPROBANTE" & vbTab & "CLIENTE" & vbTab & "TOTAL" & vbTab & "SALDO"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_venta") & vbTab & rst("fecha_emision") & vbTab & rst("fecha_vencimiento") & vbTab & rst("doc_abrev") + ":" + rst("serie") + "-" + rst("numero") & vbTab & UCase(rst("nombre_completo")) & vbTab & Format(rst("total"), "#,##0.00") & vbTab & Format(rst("saldo"), "#,##0.00")
            Grilla.AddItem Fila
            tTotal = tTotal + rst("saldo")
            rst.MoveNext
        Next i
     
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(tTotal, "#,##0.00")
      Grilla.AddItem Fila
       For k = 6 To 6
            Grilla.col = 6
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFFF
      Next k
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
  
  Exit Sub
SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub



Private Sub pedidos(ByVal Grilla As MSHFlexGrid, ByVal formulario As Form)
On Error GoTo SALIR
strCadena = "SELECT PE.id_pedido,PE.fecha,CONCAT(C.doc_abrev,':',PE.serie,'-',PE.numero) as comprobante,P.nombre_completo,A.descripcion,atendido FROM movimiento_pedido PE,comprobantes C,persona P,almacen A WHERE PE.id_alm=A.id_alm AND A.ruc='" & KEY_RUC & "' AND PE.id_doc=C.id_doc AND PE.dni_save=P.dni AND PE.ruc='" & KEY_RUC & "' AND PE.atendido='no' AND PE.anulado='no' ORDER BY PE.fecha,PE.serie,PE.numero "
Dim tTotal As Double, tSaldo As Double, nsaldo As Double
tTotal = 0
tSaldo = 0
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
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 1700
           Grilla.ColWidth(3) = 2300
           Grilla.ColWidth(4) = 2300
      Next
        cabecera = "IDPEDIDO" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "PERSONAL" & vbTab & "ALMACEN"
        Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_pedido") & vbTab & rst("fecha") & vbTab & rst("comprobante") & vbTab & UCase(rst("nombre_completo")) & vbTab & rst("descripcion")
            Grilla.AddItem Fila
           If rst("atendido") = "no" Then
            For k = 0 To 4
                Grilla.col = k
                Grilla.Row = i + 1
                Grilla.CellBackColor = &H8080FF
            Next k
           End If
            Fila = ""
            rst.MoveNext
        Next i
Exit Sub
SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub
Public Sub ProductosStockMinimo(ByVal Grilla As MSHFlexGrid)
On Error GoTo SALIR

strCadena = "SELECT A.id_producto,P.nombre_prod,U.abreviatura,A.stock FROM almacen_producto A,producto P,unidad U WHERE P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND A.id_producto=P.id_producto AND P.ruc='" & KEY_RUC & "' AND A.ruc='" & KEY_RUC & "' AND A.stock<=P.stock_minimo AND A.id_alm='" & KEY_ALM & "' ORDER BY P.id_linea,nombre_prod LIMIT 0,20"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
            Grilla.ColWidth(0) = 1000
            Grilla.ColWidth(1) = 2800
            Grilla.ColWidth(2) = 1200
            Grilla.ColWidth(3) = 1100
        Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "UND" & vbTab & "STOCK"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstT.MoveFirst
          For i = 0 To rstT.RecordCount - 1
            Fila = rstT("id_producto") & vbTab & rstT("nombre_prod") & vbTab & rstT("abreviatura") & vbTab & Format(rstT("stock"), "#,##0.00")
            Grilla.AddItem Fila
            Fila = ""
            rstT.MoveNext
        Next i
 Exit Sub
SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rstT = Nothing
End Sub
Public Sub Transferencia(ByVal Grilla As MSHFlexGrid)
On Error GoTo SALIR
strCadena = "SELECT * FROM movimiento_transferencia T,comprobantes C,almacen A WHERE T.id_doc=C.id_doc AND T.id_alm_origen=A.id_alm AND A.ruc='" & KEY_RUC & "' AND  finalizado='no' AND T.id_alm_destino='" & KEY_ALM & "' AND A.ruc='" & KEY_RUC & "' AND T.ruc='" & KEY_RUC & "' LIMIT 0,50"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 1200
            Grilla.ColWidth(2) = 2000
            Grilla.ColWidth(3) = 2500
        Next
        cabecera = "IDTRANSFERENCIA" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "ALMACEN ORIGEN"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstT.MoveFirst
          For i = 0 To rstT.RecordCount - 1
            Fila = rstT("id_transferencia") & vbTab & rstT("fecha") & vbTab & rstT("doc_abrev") & ":" & rstT("serie") & "-" & rstT("numero") & vbTab & rstT("descripcion")
            Grilla.AddItem Fila
            Fila = ""
            rstT.MoveNext
        Next i
 Exit Sub
SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rstT = Nothing
End Sub
Public Sub vencidos(ByVal Grilla As MSHFlexGrid)
On Error GoTo SALIR
Dim fecha As String
fecha = Format(DateAdd("d", 5, KEY_FECHA), "YYYY-mm-dd")
strCadena = "SELECT P.id_producto,P.nombre_prod,U.abreviatura,P.vencimiento FROM producto P,unidad U WHERE P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND vencimiento<='" & fecha & "' LIMIT 0,50"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       For Each Campo In rstT.Fields
            Grilla.ColWidth(0) = 1000
            Grilla.ColWidth(1) = 3500
            Grilla.ColWidth(2) = 1200
            Grilla.ColWidth(3) = 1200
            
        Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "UNIDAD" & vbTab & "VENCIMIENTO"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstT.MoveFirst
          For i = 0 To rstT.RecordCount - 1
            Fila = rstT("id_producto") & vbTab & rstT("nombre_prod") & vbTab & rstT("abreviatura") & vbTab & rstT("vencimiento")
            Grilla.AddItem Fila
            Fila = ""
            rstT.MoveNext
        Next i
 Exit Sub
SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rstT = Nothing
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
If Me.SSTab1.Tab = 0 Then
    Call llenar_solicitudes(Me.HfgSolicitud, Me)
End If
If Me.SSTab1.Tab = 2 Then
    Call CuentasPagar(Me.HfgPagar, Me)
End If

End Sub



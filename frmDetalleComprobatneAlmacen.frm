VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmDetalleComprobanteAlmacen 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   14025
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   470
      Left            =   2160
      TabIndex        =   56
      Top             =   2640
      Width           =   1455
      Begin VB.OptionButton Opt_afecta_stock_no 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         TabIndex        =   58
         Top             =   195
         Width           =   615
      End
      Begin VB.OptionButton opt_afecta_stock_si 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   57
         Top             =   195
         Width           =   495
      End
   End
   Begin VB.CheckBox chk_firmadodigital 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "ONLINE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2160
      TabIndex        =   53
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   2160
      TabIndex        =   48
      Top             =   5280
      Width           =   1455
      Begin VB.CheckBox chk_produccion 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "PRODUCCION"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   600
         Width           =   1305
      End
      Begin VB.OptionButton opt_electronica_si 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   120
         TabIndex        =   50
         Top             =   200
         Width           =   550
      End
      Begin VB.OptionButton opt_electronica_no 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   250
         Left            =   720
         TabIndex        =   49
         Top             =   200
         Value           =   -1  'True
         Width           =   550
      End
   End
   Begin VB.CheckBox chk_online 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "ONLINE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2175
      TabIndex        =   47
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox TxtCaracteres 
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
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   45
      Text            =   "000000"
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Height          =   800
      Left            =   2160
      TabIndex        =   39
      Top             =   3120
      Width           =   1455
      Begin VB.OptionButton opt_egreso 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "EGRESOS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   220
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton opt_ingreso 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "INGRESOS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   220
         Left            =   120
         TabIndex        =   40
         Top             =   200
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   33
      Top             =   2160
      Width           =   1455
      Begin VB.OptionButton op_caja_si 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton op_caja_no 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         TabIndex        =   34
         Top             =   200
         Width           =   615
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   29
      Top             =   4320
      Width           =   1455
      Begin VB.OptionButton OptDefaulsi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton OptDefaultno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   720
         TabIndex        =   30
         Top             =   200
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FORMATO (centimetros)"
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
      Height          =   4575
      Left            =   6000
      TabIndex        =   20
      Top             =   1200
      Width           =   7335
      Begin VB.CheckBox chk_sin_impresion 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "SIN IMPRESION"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4080
         TabIndex        =   60
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox chkDetallada 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "DETALLADA"
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
         Left            =   240
         TabIndex        =   43
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Frame FrameSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   660
         Left            =   5280
         TabIndex        =   36
         Top             =   360
         Width           =   1695
         Begin VB.TextBox txtseriemaquina 
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
            Left            =   180
            MaxLength       =   20
            TabIndex        =   37
            Top             =   240
            Width           =   1365
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SERIAL"
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
            Left            =   555
            TabIndex        =   38
            Top             =   0
            Width           =   525
         End
      End
      Begin VB.OptionButton Opt_0 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "OTRO"
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
         Left            =   4080
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Opt_4 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "TICKET"
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
         Left            =   4080
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Opt_3 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "A4"
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
         Left            =   2880
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Opt_2 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "1/2  A4"
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
         Left            =   1560
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Opt_1 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "1/4  A4"
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
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DtcFuente 
         Height          =   315
         Left            =   4080
         TabIndex        =   61
         Top             =   1440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
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
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FUENTE:"
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
         Left            =   3480
         TabIndex        =   62
         Top             =   1500
         Width           =   585
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ejm : Nº Serie,Modelo,etc."
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
         TabIndex        =   44
         Top             =   2400
         Width           =   2115
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "22.9 x 14.0"
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
         Left            =   1575
         TabIndex        =   26
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " 11.8 x 15.3"
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
         Left            =   180
         TabIndex        =   25
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   460
      Left            =   2160
      TabIndex        =   16
      Top             =   1680
      Width           =   1455
      Begin VB.OptionButton OptVno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         TabIndex        =   18
         Top             =   200
         Width           =   615
      End
      Begin VB.OptionButton OptVsi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   200
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   2160
      TabIndex        =   13
      Top             =   1200
      Width           =   1455
      Begin VB.OptionButton OptIgvsi 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   200
         Width           =   495
      End
      Begin VB.OptionButton OptIvgno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         TabIndex        =   14
         Top             =   200
         Width           =   615
      End
   End
   Begin VB.TextBox TxtNumero 
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
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   8
      Top             =   870
      Width           =   1485
   End
   Begin VB.TextBox Text1 
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
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   3
      Top             =   510
      Width           =   1485
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   8280
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleComprobatneAlmacen.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleComprobatneAlmacen.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleComprobatneAlmacen.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleComprobatneAlmacen.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleComprobatneAlmacen.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleComprobatneAlmacen.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleComprobatneAlmacen.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleComprobatneAlmacen.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleComprobatneAlmacen.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleComprobatneAlmacen.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleComprobatneAlmacen.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalleComprobatneAlmacen.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   11460
      TabIndex        =   0
      Top             =   5970
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1875
      _CBHeight       =   870
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   810
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   1429
         ButtonWidth     =   1402
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Grabar"
               Key             =   "(Grabar)"
               Object.ToolTipText     =   "Grabar Ctrl+G"
               ImageKey        =   "(Grabar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "(Cancelar)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Cancelar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   330
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
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
   Begin MSDataListLib.DataCombo DtcAsignadoa 
      Height          =   330
      Left            =   2160
      TabIndex        =   6
      Top             =   7200
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
   Begin MSDataListLib.DataCombo DtcMoneda 
      Height          =   315
      Left            =   2160
      TabIndex        =   12
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   16777215
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
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AFECTA STOCK :"
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
      Left            =   735
      TabIndex        =   59
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FIRMADO DIGITAL :"
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
      Left            =   525
      TabIndex        =   54
      Top             =   6360
      Width           =   1485
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ELECTRONICO:"
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
      Left            =   900
      TabIndex        =   52
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENVIO SUNAT:"
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
      Left            =   870
      TabIndex        =   51
      Top             =   6720
      Width           =   1125
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FORMATO :"
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
      Left            =   1110
      TabIndex        =   46
      Top             =   5040
      Width           =   885
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO DE MOVIMIENTO :"
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
      Left            =   210
      TabIndex        =   42
      Top             =   3480
      Width           =   1785
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AFECTA A CAJA :"
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
      Left            =   750
      TabIndex        =   32
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEFAULT :"
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
      Left            =   1200
      TabIndex        =   28
      Top             =   4560
      Width           =   795
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MODULO DE VENTAS :"
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
      Left            =   270
      TabIndex        =   19
      Top             =   1800
      Width           =   1725
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONEDA :"
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
      Left            =   1170
      TabIndex        =   11
      Top             =   4080
      Width           =   825
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV:"
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
      Left            =   1650
      TabIndex        =   10
      Top             =   1320
      Width           =   345
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NUMERO :"
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
      Left            =   1170
      TabIndex        =   9
      Top             =   870
      Width           =   825
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASIGNADO A:"
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
      Left            =   960
      TabIndex        =   7
      Top             =   7200
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SERIE :"
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
      Left            =   1470
      TabIndex        =   5
      Top             =   510
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPROBANTE :"
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
      Left            =   690
      TabIndex        =   4
      Top             =   120
      Width           =   1305
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   7560
      Left            =   0
      Top             =   0
      Width           =   14025
   End
End
Attribute VB_Name = "frmDetalleComprobanteAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrCodTabla As String
Dim StrCodMarca As String
Dim strigv As String
Dim strVenta As String
Dim strTipo As Integer
Dim strDefault As String * 2

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 800
  strCadena = "SELECT id_doc as Codigo,doc_des as Descripcion FROM comprobantes ORDER BY doc_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoDoc)
  
  strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda ORDER BY id_moneda ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMoneda)
 
  
  strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM persona P,entidad_empresa E WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' AND id_personal='si' ORDER BY nombre_completo "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAsignadoa)
  
  
  
  strCadena = "SELECT id_tipo_letra as Codigo, descripcion as Descripcion FROM tipo_letra_impresion  ORDER BY id_tipo_letra "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcFuente)
  
  
  
  
  
 
  

  Select Case FrmAlmacenes.Procedencia
    Case modificar
      Call LLENA
  End Select
End Sub

Private Sub LLENA()
strCadena = "SELECT * FROM almacen_comprobante WHERE id_alm_com='" & Val(FrmAlmacenes.HfgComprobante.TextMatrix(FrmAlmacenes.HfgComprobante.Row, 0)) & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
Me.DtcTipoDoc.BoundText = rst("id_doc")
Me.Text1.Text = rst("serie")
Me.TxtNumero.Text = rst("numero")
If rst("igv") = "si" Then
    Me.OptIgvsi.Value = True
Else
    Me.OptIvgno.Value = True
End If
If rst("venta") = "si" Then
    Me.OptVsi.Value = True
Else
    Me.OptVno.Value = True
End If
'If rst("detallada") = "si" Then
 '   Me.chkDetallada.Value = 1
'Else
 '   Me.chkDetallada.Value = 0
'End If

If rst("impresion") = "si" Then
   Me.chk_sin_impresion.Value = 0
   
Else
    Me.chk_sin_impresion.Value = 1
End If

If rst("tipo_movimiento") = "01" Then
    Me.opt_ingreso.Value = True
Else
    Me.opt_egreso.Value = True
End If

If rst("produccion") = "si" Then
    Me.chk_produccion.Value = 1
Else
    Me.chk_produccion.Value = 0
End If


If rst("afecta_stock") = "si" Then
   Me.opt_afecta_stock_si.Value = True
   Me.Opt_afecta_stock_no.Value = False
Else
   Me.Opt_afecta_stock_no = False
   Me.opt_afecta_stock_si.Value = True
   
End If




If rst("firmado_online") = "si" Then
   Me.chk_firmadodigital.Value = 1
Else
   Me.chk_firmadodigital.Value = 0
End If

If rst("defecto") = "si" Then
    Me.OptDefaulsi.Value = True
    strDefault = "si"
Else
    Me.OptDefaultno.Value = True
    strDefault = "no"
End If

If rst("afecta_caja") = "si" Then
    Me.op_caja_si.Value = True
Else
    Me.op_caja_no.Value = True
End If

If rst("electronico") = "si" Then
   Me.opt_electronica_si.Value = True
Else
   Me.opt_electronica_no.Value = True
End If

If rst("online") = "si" Then
   Me.chk_online.Value = 1
Else
   Me.chk_online.Value = 0
End If

Me.DtcFuente.Text = rst("fuente")


Me.FrameSerie.Visible = False
 Me.txtseriemaquina.Text = rst("serial")
Select Case rst("id_formato_impresion")
    Case 0
        Me.Opt_0.Value = True
    Case 1
        Me.Opt_1.Value = True
    Case 2
        Me.Opt_2.Value = True
    Case 3
        Me.Opt_3.Value = True
    Case 4
        Me.Opt_4.Value = True
        Me.FrameSerie.Visible = True
        Me.txtseriemaquina.Text = rst("serial")
End Select

Me.DtcMoneda.BoundText = rst("id_moneda")
End If
Set rst = Nothing
End Sub

Private Sub Save()
Dim strdetallada As String
Dim str_afecta_caja As String
Dim id_comprobante As String
Dim str_caja As String
Dim op_ingreso_egreso As String

Dim in_produccion As String

If Me.op_caja_si.Value = True Then
    str_caja = "si"
Else
    str_caja = "no"
End If

If Me.opt_ingreso.Value = True Then
    op_ingreso_egreso = "01"
Else
    op_ingreso_egreso = "02"
End If

If Me.chk_firmadodigital.Value = 1 Then
   in_firmado_online = "si"
Else
   in_firmado_online = "no"
End If


If Me.chkDetallada.Value = 1 Then
   strdetallada = "si"
Else
    strdetallada = "no"
End If

If Me.chk_produccion.Value = 1 Then
   in_produccion = "si"
Else
   in_produccion = "no"
End If


If Me.opt_afecta_stock_si.Value = True Then
   in_afecta_stock = "si"
Else
   in_afecta_stock = "no"
End If

If Me.chk_sin_impresion.Value = 1 Then
   in_impresion = "no"
Else
   in_impresion = "si"
End If


If Me.opt_electronica_si.Value = True Then
               in_electronico = "si"
Else
               in_electronico = "no"
End If
            
            If Me.chk_online.Value = 1 Then
               in_online = "si"
            Else
               in_online = "no"
            End If
            
            


  If Me.Text1.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
    
    Select Case FrmAlmacenes.Procedencia
      
      Case modificar
     
       
       strCadena = "UPDATE almacen_comprobante SET fuente='" & Trim(Me.DtcFuente.Text) & "',impresion='" & in_impresion & "', afecta_stock='" & in_afecta_stock & "',produccion='" & in_produccion & "', firmado_online='" & in_firmado_online & "',online='" & in_online & "',electronico='" & in_electronico & "',tipo_movimiento='" & op_ingreso_egreso & "',afecta_caja='" & str_caja & "', defecto='" & strDefault & "',id_formato_impresion='" & strTipo & "',numero='" & Trim(Me.TxtNumero.Text) & "',igv='" & strigv & "',venta='" & strVenta & "',id_moneda='" & Me.DtcMoneda.BoundText & "',id_usuario='" & Me.DtcAsignadoa.BoundText & "',serial='" & Trim(Me.txtseriemaquina.Text) & "' WHERE id_alm_com = '" & Val(FrmAlmacenes.HfgComprobante.TextMatrix(FrmAlmacenes.HfgComprobante.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
       CnBd.Execute (strCadena)
   
       
       If strDefault = "si" Then
            strCadena = "SELECT * FROM almacen_comprobante WHERE id_alm_com='" & Val(FrmAlmacenes.HfgComprobante.TextMatrix(FrmAlmacenes.HfgComprobante.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRst(strCadena)
            
            If rst.RecordCount > 0 Then
                KEY_COMPROBANTE = rst("id_doc")
            End If
            
            strCadena = "UPDATE almacen_comprobante SET defecto='no' WHERE id_alm='" & FrmAlmacenes.HfgAlmacen.TextMatrix(FrmAlmacenes.HfgAlmacen.Row, 0) & "' AND ruc='" & KEY_RUC & "' AND id_alm_com <> '" & Val(FrmAlmacenes.HfgComprobante.TextMatrix(FrmAlmacenes.HfgComprobante.Row, 0)) & "'"
            CnBd.Execute (strCadena)
            
           
            strCadena = "UPDATE entidad_parametros SET doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' WHERE cod_unico='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
       End If
       
       
       FrmAlmacenes.Procedencia = Neutro
       Call FrmAlmacenes.actualizacomp(FrmAlmacenes.HfgComprobante, FrmAlmacenes.HfgAlmacen.TextMatrix(FrmAlmacenes.HfgAlmacen.Row, 0))
       Unload Me
     End Select
  End If
End Sub

Private Sub Opt_0_Click()
If Me.Opt_0.Value = True Then
  strTipo = 0

End If
End Sub

Private Sub Opt_1_Click()
If Me.Opt_1.Value = True Then
    strTipo = 1
End If
End Sub

Private Sub Opt_2_Click()
If Me.Opt_2.Value = True Then
    strTipo = 2
End If
End Sub

Private Sub Opt_3_Click()
If Me.Opt_3.Value = True Then
   strTipo = 3
End If
End Sub

Private Sub Opt_4_Click()

    If Me.Opt_4.Value = True Then
    strTipo = 4
    Me.FrameSerie.Visible = True
End If

End Sub

Private Sub OptDefaulsi_Click()
If Me.OptDefaulsi.Value = True Then
    strDefault = "si"
Else
    strDefault = "no"
End If
End Sub

Private Sub OptDefaultno_Click()
If Me.OptDefaultno.Value = True Then
    strDefault = "no"
Else
    strDefault = "si"
End If
End Sub

Private Sub OptIgvsi_Click()
If Me.OptIgvsi.Value = True Then
    strigv = "si"
Else
    strigv = "no"
End If
End Sub

Private Sub OptIvgno_Click()
If Me.OptIvgno.Value = True Then
    strigv = "no"
Else
    strigv = "si"
End If
End Sub

Private Sub OptVno_Click()
If Me.OptVno.Value = True Then
    strVenta = "no"
Else
    strVenta = "si"
End If
End Sub

Private Sub OptVsi_Click()
If Me.OptVsi.Value = True Then
    strVenta = "si"
Else
    strVenta = "no"
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
    Case KEY_CANCEL
      Unload Me
  End Select
  Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub





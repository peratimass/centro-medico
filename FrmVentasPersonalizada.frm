VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmVentasPersonalizada 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "VENTA PERSONALIZADA"
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12705
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   12705
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtOperacion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6000
      MaxLength       =   80
      TabIndex        =   90
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox TxtCuotas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   6000
      MaxLength       =   80
      TabIndex        =   89
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox TxtCreditoDisponible 
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
      Left            =   4680
      MaxLength       =   80
      TabIndex        =   88
      ToolTipText     =   "TELEFONO"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frm_tarjeta 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   660
      Left            =   10480
      TabIndex        =   85
      Top             =   1200
      Visible         =   0   'False
      Width           =   2150
      Begin VB.TextBox TxtNumeroTargeta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         MaxLength       =   80
         TabIndex        =   86
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSDataListLib.DataCombo DtTargeta 
         Height          =   315
         Left            =   120
         TabIndex        =   87
         Top             =   20
         Visible         =   0   'False
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox TxtMontoPagado 
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
      Height          =   315
      Left            =   11400
      TabIndex        =   84
      Top             =   1540
      Width           =   1215
   End
   Begin VB.TextBox txtid_venta 
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
      Height          =   315
      Left            =   4680
      TabIndex        =   81
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkconyuge 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "CONYUGE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3600
      TabIndex        =   78
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtBuscarVendedor 
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
      Height          =   315
      Left            =   4440
      TabIndex        =   77
      Top             =   1800
      Width           =   1215
   End
   Begin VitekeySoft.ChameleonBtn cmdImprimir 
      Height          =   780
      Left            =   1320
      TabIndex        =   72
      Top             =   6720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1376
      BTYPE           =   3
      TX              =   "Imprimir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmVentasPersonalizada.frx":0000
      PICN            =   "FrmVentasPersonalizada.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtExonerado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   10680
      TabIndex        =   70
      Top             =   6180
      Width           =   1695
   End
   Begin VB.TextBox TxtCodProducto 
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
      Index           =   6
      Left            =   240
      TabIndex        =   66
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox TxtCodProducto 
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
      Index           =   5
      Left            =   240
      TabIndex        =   65
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox TxtCodProducto 
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
      Index           =   4
      Left            =   240
      TabIndex        =   64
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox TxtCodProducto 
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
      Index           =   3
      Left            =   240
      TabIndex        =   63
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox TxtCodProducto 
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
      Index           =   2
      Left            =   240
      TabIndex        =   62
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox TxtCodProducto 
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
      Index           =   1
      Left            =   240
      TabIndex        =   61
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TxtCodProducto 
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
      Index           =   0
      Left            =   240
      TabIndex        =   60
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtpreciototal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   10680
      TabIndex        =   57
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox txtigv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   10680
      TabIndex        =   52
      Top             =   6780
      Width           =   1695
   End
   Begin VB.TextBox txtvalorventa 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Left            =   10680
      TabIndex        =   51
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox txttotal 
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
      Index           =   6
      Left            =   10680
      TabIndex        =   50
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txttotal 
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
      Index           =   5
      Left            =   10680
      TabIndex        =   49
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox txttotal 
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
      Index           =   4
      Left            =   10680
      TabIndex        =   48
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox txttotal 
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
      Index           =   3
      Left            =   10680
      TabIndex        =   47
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txttotal 
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
      Index           =   2
      Left            =   10680
      TabIndex        =   46
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txttotal 
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
      Index           =   1
      Left            =   10680
      TabIndex        =   45
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox txttotal 
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
      Index           =   0
      Left            =   10680
      TabIndex        =   44
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txtprecio 
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
      Index           =   6
      Left            =   9240
      TabIndex        =   43
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox txtprecio 
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
      Index           =   5
      Left            =   9240
      TabIndex        =   42
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox txtprecio 
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
      Index           =   4
      Left            =   9240
      TabIndex        =   41
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtprecio 
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
      Index           =   3
      Left            =   9240
      TabIndex        =   40
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtprecio 
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
      Index           =   2
      Left            =   9240
      TabIndex        =   39
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtprecio 
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
      Index           =   1
      Left            =   9240
      TabIndex        =   38
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtprecio 
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
      Index           =   0
      Left            =   9240
      TabIndex        =   37
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtdescripcion 
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
      Index           =   6
      Left            =   3720
      TabIndex        =   36
      Top             =   5640
      Width           =   5415
   End
   Begin VB.TextBox txtdescripcion 
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
      Index           =   5
      Left            =   3720
      TabIndex        =   35
      Top             =   5280
      Width           =   5415
   End
   Begin VB.TextBox txtdescripcion 
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
      Index           =   4
      Left            =   3720
      TabIndex        =   34
      Top             =   4920
      Width           =   5415
   End
   Begin VB.TextBox txtdescripcion 
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
      Index           =   3
      Left            =   3720
      TabIndex        =   33
      Top             =   4560
      Width           =   5415
   End
   Begin VB.TextBox txtdescripcion 
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
      Index           =   2
      Left            =   3720
      TabIndex        =   32
      Top             =   4200
      Width           =   5415
   End
   Begin VB.TextBox txtdescripcion 
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
      Index           =   1
      Left            =   3720
      TabIndex        =   31
      Top             =   3840
      Width           =   5415
   End
   Begin VB.TextBox txtdescripcion 
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
      Index           =   0
      Left            =   3720
      TabIndex        =   30
      Top             =   3480
      Width           =   5415
   End
   Begin VB.TextBox txtunidad 
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
      Index           =   6
      Left            =   2640
      TabIndex        =   29
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtunidad 
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
      Index           =   5
      Left            =   2640
      TabIndex        =   28
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtunidad 
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
      Index           =   4
      Left            =   2640
      TabIndex        =   27
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtunidad 
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
      Index           =   3
      Left            =   2640
      TabIndex        =   26
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtunidad 
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
      Index           =   2
      Left            =   2640
      TabIndex        =   25
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtunidad 
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
      Index           =   1
      Left            =   2640
      TabIndex        =   24
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtunidad 
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
      Index           =   0
      Left            =   2640
      TabIndex        =   23
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtcantidad 
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
      Index           =   6
      Left            =   1560
      TabIndex        =   22
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox txtcantidad 
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
      Index           =   5
      Left            =   1560
      TabIndex        =   21
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txtcantidad 
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
      Index           =   4
      Left            =   1560
      TabIndex        =   20
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox txtcantidad 
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
      Index           =   3
      Left            =   1560
      TabIndex        =   19
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtcantidad 
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
      Index           =   2
      Left            =   1560
      TabIndex        =   18
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox txtcantidad 
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
      Index           =   1
      Left            =   1560
      TabIndex        =   17
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtcantidad 
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
      Index           =   0
      Left            =   1560
      TabIndex        =   16
      Top             =   3480
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpfecha 
      Height          =   300
      Left            =   5160
      TabIndex        =   10
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      Format          =   184221697
      CurrentDate     =   41159
   End
   Begin VB.TextBox TxtNumeroDoc 
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
      Left            =   9360
      TabIndex        =   9
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtserie 
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
      Left            =   7920
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DtcTipoDoc 
      Height          =   315
      Left            =   7920
      TabIndex        =   7
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin VB.TextBox txtdireccion 
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
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   4215
   End
   Begin VB.TextBox txtRazonsocial 
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
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox txtRuc 
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
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   1440
      TabIndex        =   59
      Top             =   330
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin MSDataListLib.DataCombo DtcMoneda 
      Height          =   315
      Left            =   9360
      TabIndex        =   68
      Top             =   840
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin VitekeySoft.ChameleonBtn cmdGuardar 
      Height          =   780
      Left            =   240
      TabIndex        =   73
      Top             =   6720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1376
      BTYPE           =   3
      TX              =   "Grabar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmVentasPersonalizada.frx":3220
      PICN            =   "FrmVentasPersonalizada.frx":323C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdcerrar 
      Height          =   780
      Left            =   2400
      TabIndex        =   74
      Top             =   6720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1376
      BTYPE           =   3
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmVentasPersonalizada.frx":3B16
      PICN            =   "FrmVentasPersonalizada.frx":3B32
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo Dtcvendedor 
      Height          =   315
      Left            =   1440
      TabIndex        =   75
      Top             =   1800
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
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
   Begin MSDataListLib.DataCombo DtcFormaPago 
      Height          =   315
      Left            =   8640
      TabIndex        =   79
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgTipoPagos 
      Height          =   1215
      Left            =   7920
      TabIndex        =   82
      Top             =   1875
      Width           =   4720
      _ExtentX        =   8334
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
   Begin MSDataListLib.DataCombo DtcFormapagodetalle 
      Height          =   315
      Left            =   7920
      TabIndex        =   83
      Top             =   1545
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F.PAGO :"
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
      Left            =   7920
      TabIndex        =   80
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENDEDOR :"
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
      Left            =   120
      TabIndex        =   76
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXONERADO :"
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
      Left            =   9270
      TabIndex        =   71
      Top             =   6240
      Width           =   1035
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONEDA :"
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
      Left            =   7920
      TabIndex        =   69
      Top             =   840
      Width           =   1350
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
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
      Left            =   360
      TabIndex        =   67
      Top             =   3240
      Width           =   615
   End
   Begin VB.Image imgfoto 
      Height          =   1935
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   195
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ALMACEN :"
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
      Left            =   120
      TabIndex        =   58
      Top             =   360
      Width           =   810
   End
   Begin VB.Label lblletras 
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
      Left            =   120
      TabIndex        =   56
      Top             =   6120
      Width           =   8550
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR VENTA :"
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
      Left            =   9180
      TabIndex        =   55
      Top             =   6600
      Width           =   1125
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV :"
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
      Left            =   9945
      TabIndex        =   54
      Top             =   6840
      Width           =   360
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL :"
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
      Left            =   9720
      TabIndex        =   53
      Top             =   7200
      Width           =   585
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   1335
      Left            =   8760
      Top             =   6120
      Width           =   3855
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.IMPORTE TOTAL"
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
      Left            =   10680
      TabIndex        =   15
      Top             =   3240
      Width           =   1350
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.UNITARIO"
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
      Left            =   9240
      TabIndex        =   14
      Top             =   3240
      Width           =   885
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONCEPTO"
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
      Left            =   5340
      TabIndex        =   13
      Top             =   3240
      Width           =   825
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNIDAD"
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
      Left            =   2640
      TabIndex        =   12
      Top             =   3240
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
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
      Left            =   1560
      TabIndex        =   11
      Top             =   3240
      Width           =   795
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   2955
      Left            =   120
      Top             =   3120
      Width           =   12495
   End
   Begin VB.Label lblruc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   0
      Width           =   3105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION :"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAZON SOCIAL :"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DNI/RUC :"
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
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   750
   End
End
Attribute VB_Name = "FrmVentasPersonalizada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Public numeroItem As Integer

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub CmdGuardar_Click()
If Len(Me.txtRuc.Text) < 11 And Len(Me.txtRuc.Text) > 11 Then
        MsgBox "INGRESE UN RUC PARA EL CLIENTE", vbInformation, KEY_EMPRESA
        Call Resalta(Me.txtRuc)
        Exit Sub
     End If
      
      
        Call Save
End Sub

Private Sub cmdImprimir_Click()
Call OrdenImpresion
End Sub

Private Sub DtcTipoDoc_Change()
    strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen  WHERE ruc='" & KEY_RUC & "'   ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  Me.DtcAlmacen.Enabled = False
  
strCadena = "SELECT serie, numero,igv FROM almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND ruc='" & KEY_RUC & "' ORDER BY serie ASC"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    Me.txtserie.Text = rstT("serie")
    Me.TxtNumeroDoc.Text = rstT("numero")
    KEY_APLICA_IGV = rstT("igv")
    If serieA <> "" And numeroA <> "" Then
    strCadena = "UPDATE temporal_ventas  SET id_doc ='" & Trim(Me.DtcTipoDoc.BoundText) & "',id_serie='" & Trim(Me.txtserie.Text) & "',numero='" & Trim(Me.TxtNumeroDoc.Text) & "'  WHERE id_serie='" & Trim(serieA) & "' AND numero='" & Trim(numeroA) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "'"
    CnBd.Execute (strCadena)
     
    End If
    If (Trim(Me.DtcTipoDoc.BoundText) = "0001") Then
        Call Resalta(Me.txtRuc)
    Else
        If (Me.DtcAlmacen.Enabled = True) Then
            Call Resalta(Me.txtcantidad(0))
        End If
    End If
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.Dtpfecha.Value = KEY_FECHA

strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM moneda  ORDER BY id_moneda ASC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)
  
strCadena = "SELECT E.cod_unico as Codigo,P.nombre_completo as Descripcion FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni and E.id_personal='si' and E.id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcVendedor)
Me.DtcVendedor.BoundText = 0
  
strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM forma_pago ORDER BY id ASC "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFormaPago)
Me.DtcFormaPago.BoundText = "01"

strCadena = "SELECT id as Codigo, descripcion as Descripcion FROM targeta WHERE id<>'00' ORDER BY id ASC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtTargeta)
  
 strCadena = "SELECT id_detalle as Codigo, descripcion as Descripcion FROM forma_pago_detalle  WHERE id='" & Me.DtcFormaPago.BoundText & "' AND ruc='" & KEY_RUC & "' AND estado='si' ORDER BY id_detalle"
 Call ConfiguraRstT(strCadena)
 Call LlenaDataComboT(Me.DtcFormapagodetalle)
    
  
strCadena = "SELECT DISTINCT A.id_doc as Codigo, C.doc_abrev as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' AND A.venta='si' AND A.id_alm='" & KEY_ALM & "' ORDER BY doc_abrev"
Call ConfiguraRst(strCadena)
 If rst.RecordCount > 0 Then
    Call LlenaDataCombo(Me.DtcTipoDoc)
    Me.DtcTipoDoc.BoundText = Trim(FrmVentas.DtcTipoDoc.BoundText)
  End If

Me.lblruc.Caption = "RUC :" & KEY_RUC
End Sub

Private Sub txtBuscarVendedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  strCadena = "SELECT E.cod_unico as Codigo,P.nombre_completo as Descripcion FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni and E.id_personal='si' and E.id_empresa='" & KEY_RUC & "' and P.nombre_completo LIKE '%" & Trim(Me.txtBuscarVendedor.Text) & "%'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcVendedor)
  
End If
End Sub
Public Sub OrdenImpresion()
Dim impresiones As Integer, id_venta As Double
       strCadena = "SELECT * FROM movimiento_venta WHERE numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND serie='" & Trim(Me.txtserie.Text) & "' AND ruc='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
          If rst("impresiones") < 1 Then
              impresiones = rst("impresiones") + 1
              id_venta = rst("id_venta")
              'Call Imprimir_Tiketera(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcAlmacen.BoundText), Trim(Me.txtSerie.Text), Trim(Me.TxtNumeroDoc.Text))
              'Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Trim(Me.txtserie.Text), Trim(Me.TxtNumeroDoc.Text), "00001")
              strCadena = "UPDATE movimiento_venta SET impresiones='" & impresiones & "' WHERE id_venta='" & id_venta & "' AND ruc='" & KEY_RUC & "'"
              CnBd.Execute (strCadena)
               
           Else
              If MsgBox("ESTE DOCUMENTO YA FUE IMPRESO:" + Space(2) + str(rst("impresiones")) + Space(1) + "IMPRESIONES" + Chr(13) + "DESEA IMPRIMIR NUEVAMENTE ?", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
                    Procedencia = imprimir_s
                    FrmSeguridad.Show
              End If
          End If
      End If

End Sub

Private Sub txtcantidad_Change(Index As Integer)

    Me.TxtTotal(Index).Text = Format(Val(Me.txtprecio(Index).Text) * Val(Me.txtcantidad(Index).Text), "###0.00")

End Sub

Private Sub TxtCodProducto_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscar(Index)
End If
End Sub
Private Sub buscar(ByVal num As Integer)
 numeroItem = num
If (Len(Me.TxtCodProducto(num).Text) = 0) Or Val(Me.TxtCodProducto(num).Text) = 0 Then
        Call Resalta(Me.TxtCodProducto(num))
       
        Procedencia = Selecionar
        FrmProducto.Show
        
        Exit Sub
    End If
    
 
    If Trim(Mid(Me.TxtCodProducto(num).Text, 1, 2)) = "00" And Len(Me.TxtCodProducto(num).Text) > 8 Then
       Me.txtcantidad(num).Text = Val(Mid(Trim(Me.TxtCodProducto(num).Text), 8, 4) / 1000)
       Me.TxtCodProducto(num).Text = Mid(Me.TxtCodProducto(num), 3, 5)
    End If
    
    If KEY_BARRAS = "si" Then
        strCadena = "SELECT B.id_producto,P.nombre_prod,P.precio_venta,P.peso,P.id_igv FROM producto_barras B,producto P ,unidad U WHERE B.id_producto=P.id_producto AND B.ruc='" & KEY_RUC & "' " & _
        "AND P.ruc='" & KEY_RUC & "' AND B.cod_barra='" & Trim(Me.TxtCodProducto(num).Text) & "'"
    Else
        Me.TxtCodProducto(num).Text = FormatosCeros(Me.TxtCodProducto(num).Text, 5)
        strCadena = "SELECT A.id_producto, P.nombre_prod,P.precio_venta,P.peso,P.id_igv,U.abreviatura FROM almacen_producto A,producto P ,unidad U WHERE P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND A.id_producto=P.id_producto AND A.id_alm='" & KEY_ALM & "' AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND A.id_producto='" & Trim(Me.TxtCodProducto(num).Text) & "'"
    End If
        
    
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        codigoP = rst("id_producto")
        Me.txtDescripcion(num).Text = rst("nombre_prod")
        Me.txtprecio(num).Text = rst("precio_venta")
        Me.TxtUnidad(num).Text = rst("abreviatura")
        If Trim(Me.txtcantidad(num).Text) > 0 Then
            Me.txtcantidad(num).Text = Me.txtcantidad(num).Text
         Else
          Me.txtcantidad(num).Text = 1
        End If
        
        Call Resalta(Me.txtcantidad(num))
End If
End Sub

Private Sub txtdescripcion_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index <> 6 Then
        Call Resalta(Me.txtprecio(Index))
    End If
End If
End Sub

Private Sub TxtMontoPagado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call realizar_ingreso_pago
End If
End Sub

Private Sub txtPrecio_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index <> 6 Then
       ' Me.txttotal(Index).Text=Val(Me.txtcantidad(Index)
        'Call Resalta(Me.txtTotal(Index).Text * Val(Me.TxtCantidad(Index).Text))
    End If
End If
End Sub
Private Sub realizar_ingreso_pago()
Dim monto_pagado As Double
Dim cod_pago As String
Dim nuevo_monto As Double
Dim strTarjeta As String
Dim Saldo_deudor As Single
Dim vencimiento As String
Dim nrecibo As String

     vencimiento = Format(KEY_FECHA, "YYYY-mm-dd")
    If Me.DtcTipoDoc.BoundText = "0007" Then
        GoTo sigy
    End If
    
    If (Val(Me.TxtMontoPagado.Text) > 0 And Val(Me.txtpreciototal.Text) > 0) Then
sigy:
     monto_pagado = Val(Me.TxtMontoPagado.Text)
     If Me.DtcMoneda.BoundText = "00002" Then
        monto_pagado = monto_pagado '* Val(Me.TxtTipoCambio.Text)
     End If
     
     
        
        If Me.DtTargeta.Visible = False Then
            strTarjeta = "00"
        Else
            strTarjeta = Me.DtTargeta.BoundText
        End If
    
    strCadena = "SELECT * FROM movimiento_venta_monto_temporal WHERE id_usuario='" & KEY_USUARIO & "' AND id_forma_pago='" & Trim(Me.DtcFormapagodetalle.BoundText) & "' and id_moneda='" & Trim(Me.DtcMoneda.BoundText) & "'  AND ruc='" & KEY_RUC & "' AND id_tarjeta LIKE '%" & strTarjeta & "%'  ORDER BY id_monto DESC"
    Call ConfiguraRst(strCadena)
        
    If rst.RecordCount < 1 Then
        
       If Me.DtcFormaPago.BoundText = "02" Then
          strCadena = "SELECT monto_credito FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' AND cod_unico='" & Trim(Me.txtRuc.Text) & "'"
          Call ConfiguraRstT(strCadena)
          If IsNull(rstT(0)) = False Then
             If (Val(Me.TxtCreditoDisponible.Text) < Val(Me.TxtMontoPagado.Text)) Then
                 MsgBox "EL MONTO EXCEDE AL CREDITO ACTUAL EN " + Space(1) + str(Format(Val(Me.TxtMontoPagado.Text) - Val(rstT(0)), "#,##0.00")), vbInformation, KEY_EMPRESA
                 Call Resalta(Me.TxtMontoPagado)
                 Exit Sub
             Else
                If (Me.TxtCuotas.Visible = True And Val(Me.TxtCuotas.Text) > 0) Then
                    For k = 1 To Val(Me.TxtCuotas.Text)
                        vencimiento = Format(DateAdd("m", 1, vencimiento), "YYYY-mm-dd")
                        strCadena = "INSERT INTO movimiento_venta_cuotas_temporal(id_cuota,id_doc,serie,numero,monto,saldo,vencimiento,id_usuario,ruc)VALUES " & _
                        "('" & formato_item(k, 2) & "','" & Me.DtcTipoDoc.BoundText & "','" & Me.txtserie.Text & "','" & Me.TxtNumeroDoc.Text & "','" & Val(Me.TxtMontoPagado.Text) / Val(Me.TxtCuotas.Text) & "','" & Val(Me.TxtMontoPagado.Text) / Val(Me.TxtCuotas.Text) & "','" & vencimiento & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                        CnBd.Execute (strCadena)
                         
                    Next k
                End If
             End If
          End If
       End If
       
      
       
       strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,id_forma_pago,id_moneda,monto,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,cuotas,id_usuario,id_recibo,detalle,fecha,id_alm,ruc) VALUES " & _
       " ('" & Me.DtcTipoDoc.BoundText & "','" & Me.txtserie.Text & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.DtcFormapagodetalle.BoundText) & "','" & Me.DtcMoneda.BoundText & "','" & monto_pagado & "','" & strTarjeta & "','" & Me.TxtNumeroTargeta.Text & "','" & Me.txtOperacion.Text & "','" & Val(Me.TxtCuotas.Text) & "','" & KEY_USUARIO & "','-','" & nrecibo & "',CURDATE(),'" & KEY_ALM & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
        
       
    Else
    If Me.DtcMoneda.BoundText = "00001" Then
        nuevo_monto = Val(Me.TxtMontoPagado.Text)
    Else
        nuevo_monto = Val(Me.TxtMontoPagado.Text) '* Val(Me.TxtTipoCambio.Text)
    End If
    strCadena = "UPDATE movimiento_venta_monto_temporal SET monto='" & nuevo_monto & "',id_tarjeta='" & strTarjeta & "',id_tarjeta_numero='" & Me.TxtNumeroTargeta.Text & "',id_tarjeta_operacion='" & Me.txtOperacion.Text & "',id_recibo='0',detalle='" & nrecibo & "' WHERE id_moneda='" & Me.DtcMoneda.BoundText & "' and id_usuario='" & KEY_USUARIO & "' AND id_forma_pago='" & Trim(Me.DtcFormapagodetalle.BoundText) & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.txtserie.Text) & "' AND numero='" & Me.TxtNumeroDoc.Text & "' AND id_tarjeta LIKE '%" & strTarjeta & "%' AND ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    End If
    Me.TxtNumeroTargeta.Text = ""
    Me.txtOperacion.Text = ""
    Call llena_pagos(Me.HfgTipoPagos, Me.TxtNumeroDoc.Text)
    'Call DisplayTextoCom("TOTAL : S/." & AlineaString(Me.lblTotal.Caption, 9, pAlnDerecha) & _
                            "VUELTO: S/." & AlineaString(Me.lblVuelto.Caption, 9, pAlnDerecha), mscConecta)
    If (Me.TxtCuotas.Visible = True And Val(Me.TxtCuotas.Text) > 0) Then
         FrmVentasCuotas.Show
    End If
    Call Resalta(Me.TxtCodProducto)
    End If

End Sub
Public Sub llena_pagos(ByVal Grilla As MSHFlexGrid, ByVal idVenta As String)
On Error GoTo salir
Dim tpago As Double
Dim strTarjeta As String
strCadena = "SELECT * FROM movimiento_venta_monto_temporal M,forma_pago_detalle F WHERE M.id_forma_pago=F.id_detalle AND id_usuario='" & KEY_USUARIO & "' AND numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Me.txtserie.Text & "' AND M.ruc='" & KEY_RUC & "' AND F.ruc='" & KEY_RUC & "' "
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
           Grilla.ColWidth(1) = 2800
           Grilla.ColWidth(2) = 1200
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
            Select Case rst("id_moneda")
                Case "00001"
                    strmoneda = " [ S/.]"
                Case "00002"
                    strmoneda = " [ USS/.]"
            End Select
            
            strCadena = "SELECT * FROM targeta WHERE id='" & rst("id_tarjeta") & "'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
                If rst("id_tarjeta") = "00" Then
                    strTarjeta = rst("descripcion") + Space(1) + "[" + rst("detalle") & "]"
                Else
                    strTarjeta = rst("descripcion") + Space(1) + rstT("descripcion")
                End If
               
            Else
                strTarjeta = strmoneda & Space(1) & rst("descripcion")
            End If
            Fila = rst("id_monto") & vbTab & strTarjeta & vbTab & Format(rst("monto"), "###0.00")
            Grilla.AddItem Fila
            tpago = rst("monto") + tpago
            rst.MoveNext
    Next i
    Dim tventa As Double
   
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub
Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If Trim(Me.txtRuc.Text) = "" Then
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
 End If
  
 If Trim(Me.DtcTipoDoc.BoundText) = "0001" And (Trim(Me.txtRuc.Text) = "00000000" Or Trim(Me.txtRuc.Text) = "" Or Len(Trim(Me.txtRuc.Text)) <> 11) Then
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
End If

If (Len(Trim(Me.txtRuc.Text)) = 8 And Trim(Me.txtRuc.Text) = "00000000") Then
    Me.txtRuc.Text = "PUBLICO EN GENERAL"
    Me.TxtDireccion.Text = KEY_DIR_PUBLIC
    Call Resalta(Me.txtrazonsocial)
    Exit Sub
End If


 If Trim(Me.DtcTipoDoc.BoundText) = "0003" And (Trim(Me.txtRuc.Text) = "") Then
    Me.txtRuc.Text = "00000000"
    Me.txtrazonsocial.Text = "PUBLICO EN GENERAL"
    Me.TxtDireccion.Text = KEY_DIR_PUBLIC
    Call Resalta(Me.txtrazonsocial)
    Exit Sub
End If
If Len(Trim(Me.txtRuc.Text)) = 8 Or Len(Trim(Me.txtRuc.Text)) = 11 Then
    strCadena = "SELECT dni,nombre_completo,direccion,foto  FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        FrmDetallePersona.txtRuc.Text = Trim(Me.txtRuc.Text)
        FrmDetallePersona.ChkCliente.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.imgFoto.Visible = True
        Me.txtrazonsocial.Text = UCase(rst("nombre_completo"))
        Me.TxtDireccion.Text = UCase(rst("direccion"))
        If IsNull(rst("foto")) = False And Len(rst("foto")) > 5 Then
            If VerificarFichero(App.Path & "\archivos\" & rst("dni")) = True Then
                Me.imgFoto.Picture = LoadPicture(App.Path + "\archivos\" + rst("dni") + "\" + Trim(rst("foto")))
            Else
                Me.imgFoto.Picture = LoadPicture(App.Path + "\archivos\no_photo.jpg")
            End If
        Else
            Me.imgFoto.Picture = LoadPicture(App.Path + "\archivos\no_photo.jpg")
        End If
        Call Resalta(Me.TxtCodProducto(0))
    End If
End If
   
End If
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtserie.Text = formato_item(Trim(Me.txtserie.Text), 3)
    strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.txtserie.Text) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.TxtNumeroDoc.Text = rst("numero")
        Call Resalta(Me.txtRuc)
    Else
        Call Resalta(Me.txtserie)
        Me.TxtNumeroDoc.Text = ""
    End If
End If
End Sub

Private Sub txttotal_Change(Index As Integer)
Dim tTotal As Double
For i = 0 To 6
    tTotal = Val(Me.TxtTotal(i).Text) + tTotal
Next i
If KEY_APLICA_IGV = "si" Then
    SUBTOTAL = tTotal / (1 + KEY_IGV)
    igv = tTotal - SUBTOTAL
    texonerado = 0
Else
     texonerado = tTotal + texonerado
    SUBTOTAL = 0
    igv = 0
End If
Me.txtexonerado.Text = Format(texonerado, "###0.00")
Me.txtValorVenta.Text = Format(SUBTOTAL, "###0.00")
Me.TxtIgv.Text = Format(igv, "###0.00")
Me.txtpreciototal.Text = Format(tTotal, "###0.00")
End Sub
Public Sub Save()
'On Error GoTo Error
Dim i As Integer, anul As String * 2, MontoActual As Double, TotalVenta As Double
Dim igv As Double, SUBTOTAL As Double, exonerado As Double, dfac As String, Monto_descuento As Single
Dim monto_pagado As Double, Monto_Vuelto As Double, Monto_Sobrante As Double, saldo_f As Double, estado_f As String
Dim id_venta  As Double, CodReferencia As String, KEY_VENCIMIENTO As String, cod_cliente As String, rst1 As New ADODB.Recordset, p As Integer
Dim horario As String, turno As String
Dim id_tipo_factura As String



If Trim(Me.DtcVendedor.Text) = "" Then
    MsgBox "DEBE SELECCIONAR UN VENDEDOR. ", vbInformation
    Me.DtcVendedor.SetFocus
    Exit Sub
End If



horario = Format(Time, "hh:mm")
If horario >= "07:00" And horario <= "13:00" Then
   turno = "M"
Else
   turno = "T"
End If

If KEY_CERVECERIA = "no" Then
    KEY_MONTOENVASE = 0#
    KEY_ENVASE = "no"
    
Else
    KEY_DETALLE = KEY_DETALLE
    KEY_MONTOENVASE = KEY_MONTOENVASE
    KEY_ENVASE = KEY_ENVASE
End If
If Trim(Me.txtRuc.Text) = "" Then
    Me.txtRuc.Text = "00000000"
End If
If Trim(Me.DtcTipoDoc.BoundText) = "0001" And Len(Me.txtRuc.Text) <> 11 Then
    MsgBox "INGRESE RUC VALIDO PARA EL CLIENTE", vbInformation, KEY_EMPRESA
    Call Resalta(Me.txtRuc)
    Exit Sub
End If

SUBTOTAL = Val(Me.txtValorVenta.Text)
igv = Val(Me.TxtIgv.Text)
exonerado = Val(Me.txtexonerado.Text)
TotalVenta = Val(Format(Me.txtpreciototal.Text, "###0.000"))

Monto_descuento = 0
monto_pagado = Val(TotalVenta)
Monto_Vuelto = 0
Monto_Sobrante = 0

If Me.chkconyuge.Value = 1 Then
    strconyugue = "si"
Else
    strconyugue = "no"
End If

If (Trim(Me.DtcFormaPago.BoundText) = "05") Then
    If (Trim(Me.txtRuc.Text) = "00000000") Then
        MsgBox "Elija un Cliente Registrado, para dar Credito", vbInformation, "Mensaje de Administracion"
        Call Resalta(Me.txtRuc)
        Exit Sub
    End If
    saldo_f = TotalVenta
    estado_f = "Credito"
Else
    saldo_f = KEY_NULO
End If
cod_cliente = Trim(txtRuc.Text)


strCadena = "SELECT * FROM movimiento_venta_monto_temporal WHERE numero='" & Trim(Me.TxtNumeroDoc.Text) & "' and id_usuario='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Me.txtserie.Text & "' ORDER BY id_monto ASC"
rst1.CursorLocation = adUseClient
rst1.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
If rst1.RecordCount > 0 Then
   rst1.MoveFirst
       'strCadena = "SELECT * FROM movimiento_venta WHERE numero='" & Trim(Me.TxtNumeroDoc.text) & "' AND serie='" & Trim(Me.TxtSerie.text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND ruc='" & KEY_RUC & "'"
       'Call ConfiguraRst(strCadena)
       
       If Val(Monto_Vuelto) < 0 Then
            MsgBox "El Monto Pagado es Inferior al Monto Total", vbInformation, "Mensaje para el Usuario"
            Call Resalta(Me.txtpreciototal)
            Exit Sub
        End If
        If Me.DtcFormaPago.BoundText = "05" Then
            'KEY_VENCIMIENTO = Format(Me.DtpFechaReferencia.Value, "yyyy-mm-dd")
        Else
            rst1.MoveFirst
            Saldo = 0
            For i = 0 To rst1.RecordCount - 1
                If rst1("id_forma_pago") = "08" Then
                    Saldo = rst1("monto")
                End If
                rst1.MoveNext
            Next i
            KEY_VENCIMIENTO = KEY_FECHA
        End If
            
    If strEspecial > 100 Then
    '      Call save_especial
     '     Exit Sub
    Else
            
            
            Documento = Trim(Me.DtcTipoDoc.Text) & ":" & Trim(Me.txtserie.Text) & "-" & Trim(Me.TxtNumeroDoc.Text)
            
            'If Me.cmdSeriales.Visible = True Then
            'If KEY_TRAMITE = "si" Then
            'If trim = "si" Then
                
                id_tipo_factura = "00002"
            'Else
                'id_tipo_factura = Trim(Me.txttipofactura.Text)
            'End If
            
            
            strCadena = "P_insert_venta_v2('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcFormaPago.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
            "'" & Trim(Me.txtserie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Me.txtRuc.Text & "','" & Me.txtrazonsocial.Text & "','" & SUBTOTAL & "','" & igv & "','" & exonerado & "','" & TotalVenta & "','" & Saldo & "'," & _
            "'" & Val(Me.txtpreciototal.Text) & "','0.00','" & KEY_FECHA & "','" & KEY_VENCIMIENTO & "','" & id_tipo_factura & "','" & Me.DtcVendedor.BoundText & "','" & KEY_USUARIO & "','" & KEY_CAMBIO & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Documento & "','" & horario & "','T','--','" & strconyugue & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
            
            
            id_venta = LastRegistroRUC("movimiento_venta", "id_venta")
            Me.txtid_venta.Text = id_venta
            Call SaveDetalleDocumentoVenta(id_venta)
     End If
        
            
        
        
        
        rst1.MoveFirst
        For k = 0 To rst1.RecordCount - 1
            strCadena = "INSERT INTO movimiento_venta_monto(id_venta,id_forma_pago,monto,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,ruc)VALUES('" & id_venta & "','" & rst1("id_forma_pago") & "','" & rst1("monto") & "','" & rst1("id_tarjeta") & "','" & rst1("id_tarjeta_numero") & "','" & rst1("id_tarjeta_operacion") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
               rst1.MoveNext
        Next k
       
       
        
        
End If
StrNumero = FormatosCeros(Trim(str(Val(Me.TxtNumeroDoc.Text)) + 1), 6)
strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE  id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.txtserie.Text) & "'  AND ruc='" & Trim(KEY_RUC) & "'"
CnBd.Execute (strCadena)
 
                
                
                
            
                dfactura = False
                
                Exit Sub
        Exit Sub
'Error:
 ' MsgBox "Uso Incorrecto del Sistema", vbInformation, KEY_EMPRESA
  'MsgBox "PULSE F2 PARA SALVAR LOS PRODUCTOS", vbInformation, KEY_EMPRESA
  
End Sub

Private Sub SaveDetalleDocumentoVenta(ByVal idVenta As Double)
Dim codigo As String, cantidad As Single, descripcion As String, punitario As Double, ptotal As Double, ppeso As Double
   
       For i = 0 To 6
           If Me.txtDescripcion(i).Text <> "" Then
                
                If Me.TxtCodProducto(i).Text = "" Then
                    codigo = 0
                    Peso = 0
                Else
                    codigo = Me.TxtCodProducto(i).Text
                    Peso = BDBuscarCampo("producto", "peso", "id_producto", codigo)
                End If
                If Me.txtcantidad(i).Text = "" Then
                    cantidad = 0
                Else
                    cantidad = Me.txtcantidad(i).Text
                End If
                If Me.TxtTotal(i).Text = "" Then
                    ptotal = 0
                Else
                    ptotal = Me.TxtTotal(i).Text
                End If
                If Me.txtprecio(i).Text = "" Then
                   punitario = 0
                Else
                    punitario = Me.txtprecio(i)
                
                End If
                
                strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,cantidad,precio,peso,total,detalle,ruc) VALUES ('" & idVenta & "','" & codigo & "','" & cantidad & "','" & punitario & "','" & Peso & "','" & ptotal & "','" & Replace(Trim(Me.txtDescripcion(i).Text), "'", "") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                 
           Else
                Exit Sub
           End If
        Next i
   
End Sub


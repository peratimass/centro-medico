VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmBusquedaCompras 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtruc 
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6345
      TabIndex        =   36
      Top             =   480
      Width           =   1695
   End
   Begin VB.Frame frmDetraccion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DATOS DETRACCION"
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
      Height          =   2925
      Left            =   12840
      TabIndex        =   21
      Top             =   2040
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtconstancia 
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
         Height          =   285
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   25
         Top             =   1500
         Width           =   1455
      End
      Begin VB.TextBox txtnumerodetraccion 
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
         Height          =   285
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   24
         Top             =   1845
         Width           =   1455
      End
      Begin VB.TextBox txtmontodetraccion 
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
         Height          =   285
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   23
         Top             =   1170
         Width           =   1455
      End
      Begin VB.TextBox txtprocentajedetraccion 
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
         Height          =   285
         Left            =   1680
         MaxLength       =   80
         TabIndex        =   22
         Top             =   840
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker Dtpfechadetraccion 
         Height          =   300
         Left            =   4080
         TabIndex        =   26
         Top             =   1800
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
         Format          =   54263809
         CurrentDate     =   42752
      End
      Begin MSDataListLib.DataCombo DtcTiposervicio 
         Height          =   315
         Left            =   1680
         TabIndex        =   27
         Top             =   360
         Width           =   3975
         _ExtentX        =   7011
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
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   375
         Left            =   1680
         TabIndex        =   34
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
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
         MICON           =   "FrmBusquedaCompras.frx":0000
         PICN            =   "FrmBusquedaCompras.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdcerrar 
         Height          =   375
         Left            =   3600
         TabIndex        =   35
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "CERRAR"
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
         MICON           =   "FrmBusquedaCompras.frx":2601
         PICN            =   "FrmBusquedaCompras.frx":261D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:"
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
         Left            =   3360
         TabIndex        =   33
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T.SERVICIO :"
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
         Left            =   645
         TabIndex        =   32
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° CONSTANCIA :"
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
         Left            =   285
         TabIndex        =   31
         Top             =   1500
         Width           =   1140
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° DE OPERACION :"
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
         TabIndex        =   30
         Top             =   1920
         Width           =   1275
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO DETRACC:"
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
         Left            =   210
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PORCENTAJE :"
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
         Left            =   540
         TabIndex        =   28
         Top             =   840
         Width           =   930
      End
   End
   Begin VitekeySoft.ChameleonBtn CmdBuscar 
      Height          =   330
      Left            =   17280
      TabIndex        =   17
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmBusquedaCompras.frx":54D1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcTipoCompra 
      Height          =   315
      Left            =   8520
      TabIndex        =   16
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
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
   Begin VitekeySoft.ChameleonBtn cmdexportar 
      Height          =   975
      Left            =   19080
      TabIndex        =   10
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "EXPORTAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
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
      MICON           =   "FrmBusquedaCompras.frx":54ED
      PICN            =   "FrmBusquedaCompras.frx":5509
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
      Height          =   975
      Left            =   19080
      TabIndex        =   9
      Top             =   5160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "SALIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
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
      MICON           =   "FrmBusquedaCompras.frx":595B
      PICN            =   "FrmBusquedaCompras.frx":5977
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpFecha 
      Height          =   315
      Left            =   14040
      TabIndex        =   5
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   54263809
      CurrentDate     =   41251
   End
   Begin VB.TextBox TxtNumero 
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1755
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox TxtEmpresa 
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4155
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   7935
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   13996
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      GridColor       =   12582912
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
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   12960
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusquedaCompras.frx":5D67
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusquedaCompras.frx":61BB
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusquedaCompras.frx":64DB
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusquedaCompras.frx":692F
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusquedaCompras.frx":6D83
            Key             =   "(RCompras)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBusquedaCompras.frx":84F5
            Key             =   "(RVentas)"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DtpHasta 
      Height          =   315
      Left            =   15840
      TabIndex        =   7
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   54263809
      CurrentDate     =   41251
   End
   Begin MSComCtl2.DTPicker DtpInicioRegistro 
      Height          =   315
      Left            =   14040
      TabIndex        =   12
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   54263809
      CurrentDate     =   41251
   End
   Begin MSComCtl2.DTPicker DtpFinRegistro 
      Height          =   315
      Left            =   15840
      TabIndex        =   13
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   54263809
      CurrentDate     =   41251
   End
   Begin VitekeySoft.ChameleonBtn cmdgastosImportacion 
      Height          =   900
      Left            =   19080
      TabIndex        =   15
      Top             =   3120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
      BTYPE           =   5
      TX              =   "GASTOS"
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
      MICON           =   "FrmBusquedaCompras.frx":8947
      PICN            =   "FrmBusquedaCompras.frx":8963
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdbuscarRegistro 
      Height          =   330
      Left            =   17280
      TabIndex        =   18
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmBusquedaCompras.frx":AF34
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdTipoCompra 
      Height          =   330
      Left            =   11640
      TabIndex        =   19
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmBusquedaCompras.frx":AF50
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdDetraccion 
      Height          =   900
      Left            =   19080
      TabIndex        =   20
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
      BTYPE           =   5
      TX              =   "DETRACCION"
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
      MICON           =   "FrmBusquedaCompras.frx":AF6C
      PICN            =   "FrmBusquedaCompras.frx":AF88
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
      Height          =   1020
      Left            =   19080
      TabIndex        =   38
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1799
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
      MICON           =   "FrmBusquedaCompras.frx":D5C1
      PICN            =   "FrmBusquedaCompras.frx":D5DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC :"
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
      Left            =   5925
      TabIndex        =   37
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AL"
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
      Left            =   15480
      TabIndex        =   14
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA REGISTRO :"
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
      Left            =   12690
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AL"
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
      Left            =   15450
      TabIndex        =   8
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Left            =   12690
      TabIndex        =   6
      Top             =   300
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N° COMPROBANTE:"
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
      Left            =   405
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR :"
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
      Left            =   3090
      TabIndex        =   3
      Top             =   480
      Width           =   945
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   915
      Left            =   240
      Top             =   120
      Width           =   18735
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   9225
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmBusquedaCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub Command1_Click()

End Sub

Private Sub cmdBuscar_Click()
    
    strCadena = "SELECT * FROM view_compras WHERE id_proveedor LIKE '%" & Trim(Me.txtRuc.Text) & "%' and nproveedor LIKE '%" & Trim(Me.TxtEmpresa.Text) & "%' and  ruc='" & KEY_RUC & "' and  fecha_emision>='" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "'"
    Call llenar_grid(Me.HfdPersona)
    
    
    
End Sub
Private Sub cmdbuscarRegistro_Click()
    
    strCadena = "SELECT * FROM view_compras WHERE ruc='" & KEY_RUC & "' and  fecha_registro>='" & Format(Me.DtpInicioRegistro.Value, "YYYY-mm-dd") & "' and fecha_registro<='" & Format(Me.DtpFinRegistro.Value, "YYYY-mm-dd") & "'"
    Call llenar_grid(Me.HfdPersona)
    
End Sub

Private Sub cmdCerrar_Click()
Me.frmdetraccion.Visible = False
End Sub

Private Sub get_detraccion(ByVal in_compra As String)
strCadena = "SELECT codigo as Codigo,CONCAT(codigo,'-',descripcion) as Descripcion FROM tipo_detraccion order by codigo ASC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTiposervicio)
Me.Dtpfechadetraccion.Value = KEY_FECHA
Me.frmdetraccion.Visible = True

strCadena = "SELECT * FROM movimiento_compra_detraccion WHERE id_compra='" & Val(in_compra) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txtconstancia.Text = rst("constancia")
    Me.txtmontodetraccion.Text = rst("monto")
    Me.txtnumerodetraccion.Text = rst("numero")
    Me.DtcTiposervicio.BoundText = rst("id_servicio")
    
    Me.Dtpfechadetraccion.Value = rst("fecha_detraccion")
Else
    Me.txtconstancia.Text = ""
    Me.txtmontodetraccion.Text = 0
    Me.txtnumerodetraccion.Text = ""
    Me.DtcTiposervicio.BoundText = 0
    Me.Dtpfechadetraccion.Value = KEY_FECHA
End If
End Sub

Private Sub cmdDetraccion_Click()
Call get_detraccion(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))

End Sub

Private Sub cmdexportar_Click()
Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptBusquedaCompra", , App.Path + "\Reportes\")
End Sub

Private Sub cmdgastosImportacion_Click()
strCadena = "SELECT G.fecha,'-',CONCAT(C.doc_abrev,':',serie,'-',numero) as comprobante,P.dni,P.nombre_completo,G.monto,G.descripcion FROM movimiento_compra_gasto G,comprobantes C,persona P WHERE id_compra='" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND G.ruc='" & KEY_RUC & "' AND G.id_doc=C.id_doc AND G.id_persona=P.dni"
Call ConfiguraRst(strCadena)

   Ans = ShowMultiReport(rst, "RptComprobante_Gasto", , App.Path + "\Reportes\")
End Sub

Private Sub cmdImprimir_Click()
strCadena = "SELECT * FROM view_compra_vista WHERE id_compra='" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptCompra", , App.Path + "\Reportes\")
End Sub

Private Sub cmdprocesar_Click()

Call put_detraccion(Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)), Trim(Me.txtconstancia.Text), Me.DtcTiposervicio.BoundText, Me.Dtpfechadetraccion.Value)


End Sub
Private Sub put_detraccion(ByVal in_compra As String, ByVal in_constancia As String, ByVal in_tipo_servicio As String, ByVal in_fecha As String)

            strCadena = "SELECT * FROM movimiento_compra_detraccion WHERE id_compra='" & Val(in_compra) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                strCadena = "call p_insert_detraccion('" & in_compra & "','" & Trim(Me.txtnumerodetraccion.Text) & "','" & Val(Me.txtmontodetraccion.Text) & "','--','" & in_tipo_servicio & "','" & in_constancia & "','" & Format(in_fecha, "YYYY-mm-dd") & "')"
                CnBd.Execute (strCadena)
            Else
                strCadena = "UPDATE movimiento_compra_detraccion SET numero='" & Trim(Me.txtnumerodetraccion.Text) & "',fecha_detraccion='" & Format(Me.Dtpfechadetraccion.Value, "YYYY-mm-dd") & "',monto='" & Val(Me.txtmontodetraccion.Text) & "',constancia='" & Trim(Me.txtconstancia.Text) & "',id_servicio='" & Me.DtcTiposervicio.BoundText & "' WHERE id_compra='" & Val(in_compra) & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
            End If
            Me.frmdetraccion.Visible = False
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdTipoCompra_Click()
    strCadena = "SELECT * FROM view_compras WHERE ruc='" & KEY_RUC & "' and  fecha_emision>='" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "' and id_tipo_compra='" & Me.DtcTipoCompra.BoundText & "'"
    Call llenar_grid(Me.HfdPersona)

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = Me.Top = 10
Me.Dtpfecha.Value = KEY_FECHA
Me.DtpHasta.Value = KEY_FECHA

Me.DtpInicioRegistro.Value = KEY_FECHA
Me.DtpFinRegistro.Value = KEY_FECHA

strCadena = "SELECT tipo_compra as Codigo, descripcion as Descripcion FROM tipo_compra ORDER By tipo_compra "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoCompra)



Call actualizar
End Sub
Private Sub actualizar()
    
    
    If KEY_PAIS = KEY_PERU Then
        strCadena = "SELECT * FROM view_compras WHERE ruc='" & KEY_RUC & "'  LIMIT 25"
       
    Else
        strCadena = "SELECT * FROM view_compras_internacional WHERE ruc='" & KEY_RUC & "'  LIMIT 25"
       
    End If
    
     Call llenar_grid(Me.HfdPersona)
    
End Sub
Private Sub llenar_grid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Me.cmdexportar.Enabled = False
    Exit Sub
End If
   Me.cmdexportar.Enabled = True
    
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 2400
           Grilla.ColWidth(5) = 1400
           Grilla.ColWidth(6) = 3000
           Grilla.ColWidth(7) = 1500
           Grilla.ColWidth(8) = 1000
           Grilla.ColWidth(9) = 1200
           Grilla.ColWidth(10) = 2500
           Grilla.ColWidth(11) = 1200
        Next
         cabecera = "CODIGO" & vbTab & "REGISTRO" & vbTab & "EMISION" & vbTab & "CANCELACION" & vbTab & "COMPROBANTE" & vbTab & "RUC" & vbTab & "PROVEEDOR" & vbTab & "TIPO COMPRA" & vbTab & "MONEDA" & vbTab & "TOTAL" & vbTab & "ALMACEN" & vbTab & "OPERADOR"
         Grilla.AddItem cabecera
         For k = 0 To 11
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = Format(rst("id_compra"), "000000") & vbTab & Format(rst("fecha_registro"), "dd-mm-YYYY") & vbTab & Format(rst("fecha_emision"), "dd-mm-YYYY") & vbTab & Format(rst("fecha_cancelacion"), "dd-mm-YYYY") & vbTab & rst("comprobante") & vbTab & rst("id_proveedor") & vbTab & rst("nproveedor") & vbTab & rst("tipocompra") & vbTab & rst("moneda") & vbTab & Format(rst("total"), "#,##0.00") & vbTab & rst("almacen") & vbTab & rst("nombre_completo")
             Grilla.AddItem Fila
           
        rst.MoveNext
        Next i
 ' Grilla.Row = 1
 ' Grilla.col = 0
 ' Grilla.ColSel = 1
 ' Grilla.RowSel = 1
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub HfdPersona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If FrmCompras.Procedencia = Selecionar Then
       FrmCompras.Procedencia = Neutro
       strCadena = "SELECT * FROM movimiento_compra where id_compra='" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
            FrmCompras.lblid_compra_vinculada.Caption = rst("id_compra")
            FrmCompras.txtComprobante_vinculado.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 4)
            
            FrmCompras.TxtMontoTotal_vinculado.Text = rst("total")
            FrmCompras.txtMonto_porcentaje.Text = "100"
            FrmCompras.txtMonto_asignado.Text = rst("total")
       End If
       
       Unload Me
       Exit Sub
    End If
    
    
    
    
    If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
        Call FrmCompras.nuevo
        strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        FrmCompras.DtcAlmacen.Enabled = True
        FrmCompras.DtcTipoDoc.Enabled = True
        FrmCompras.txtserie.Enabled = True
        FrmCompras.TxtNumeroDoc.Enabled = True
        
        FrmCompras.DtcTipoDoc.BoundText = rst("id_doc")
        FrmCompras.txtRuc.Text = rst("id_proveedor")
        FrmCompras.TxtProveedor.Text = rst("nproveedor")
        FrmCompras.Txtdoc_cod.Text = rst("id_doc")
        FrmCompras.txtserie.Text = rst("serie")
        FrmCompras.TxtNumeroDoc.Text = rst("numero")
        FrmCompras.TxtIdCompra.Text = rst("id_compra")
        FrmCompras.TxtFecha_emision.Mask = ""
        FrmCompras.TxtFecha_emision.Text = ""
        FrmCompras.TxtFecha_emision.Mask = "##/##/####"
        FrmCompras.txtFecha_vencimiento.Mask = ""
        FrmCompras.txtFecha_vencimiento.Text = ""
        FrmCompras.txtFecha_vencimiento.Mask = "##/##/####"
        FrmCompras.TxtFecha_emision.Text = rst("fecha_emision")
        
        If FrmCompras.Frame_Relacionado.Visible = True Then
           FrmCompras.DtcRelacionado.BoundText = rst("id_doc_fact")
           FrmCompras.TxtSerieR.Text = rst("serie_fact")
           FrmCompras.TxtNumeroR.Text = rst("numero_fact")
           strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE dni='" & rst("id_proveedor_relacionado") & "' and   ruc='" & KEY_RUC & "'"
           Call ConfiguraRstT(strCadena)
           Call LlenaDataComboT(FrmCompras.DtcProveedor)
    
           FrmCompras.DtcProveedor.BoundText = rst("id_proveedor_relacionado")
           If IsNull(FrmCompras.TxtFechaEmision.Text) = True Then
                FrmCompras.TxtFechaEmision.Text = ""
           Else
                If IsNull(rst("fecha_fact")) = True Then
                    FrmCompras.TxtFechaEmision.Text = ""
                Else
                    FrmCompras.TxtFechaEmision.Text = rst("fecha_fact")
                End If
                
           End If
           
        End If
    
        
        
        
        
        Procedencia = buscar
        Call FrmCompras.buscar_comprobante(Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)))
        
        Unload Me
        Exit Sub
    End If
    
    
    
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
   Case "(Exportar)"
        
        
    Case KEY_EXIT
        Unload Me

End Select
End Sub







Private Sub TxtEmpresa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_compras WHERE ruc='" & KEY_RUC & "' and  nproveedor LIKE '%" & Trim(Me.TxtEmpresa.Text) & "%'"
    Call llenar_grid(Me.HfdPersona)
End If
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_compras WHERE ruc='" & KEY_RUC & "' and  numero LIKE '%" & Trim(Me.txtNumero.Text) & "%'"
    Call llenar_grid(Me.HfdPersona)
End If
End Sub


Private Sub TxtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_compras WHERE ruc='" & KEY_RUC & "' and  id_proveedor LIKE '%" & Trim(Me.txtRuc.Text) & "%'"
    Call llenar_grid(Me.HfdPersona)
End If

End Sub

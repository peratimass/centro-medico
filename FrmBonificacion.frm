VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBonificacion 
   BorderStyle     =   0  'None
   ClientHeight    =   8880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   18480
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OptDescuentos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "DESCUENTOS"
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
      Height          =   250
      Left            =   5040
      TabIndex        =   80
      Top             =   1000
      Width           =   2775
   End
   Begin TabDlg.SSTab frmdetalle 
      Height          =   6855
      Left            =   2760
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   12091
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "BONIFICACION LINEA"
      TabPicture(0)   =   "FrmBonificacion.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape4"
      Tab(0).Control(1)=   "IMG_CLOSE"
      Tab(0).Control(2)=   "lblmonto"
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(5)=   "lblproducto"
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(11)=   "DtcModelo"
      Tab(0).Control(12)=   "DtcCobertura"
      Tab(0).Control(13)=   "DtpBonificacionFin"
      Tab(0).Control(14)=   "DtcSubFamilia"
      Tab(0).Control(15)=   "DtcLinea"
      Tab(0).Control(16)=   "cmdprocesar"
      Tab(0).Control(17)=   "DtpBonificacionIni"
      Tab(0).Control(18)=   "txtcantidad"
      Tab(0).Control(19)=   "txtid_producto"
      Tab(0).Control(20)=   "txtcantidadbonificacion"
      Tab(0).Control(21)=   "txtlinea"
      Tab(0).Control(22)=   "txtDescripcionBonificacion"
      Tab(0).Control(23)=   "txtidbonificacion"
      Tab(0).Control(24)=   "chk_bonificacion_monto"
      Tab(0).Control(25)=   "chk_Familia"
      Tab(0).Control(26)=   "chk_subfamilia"
      Tab(0).Control(27)=   "chk_modelo"
      Tab(0).Control(28)=   "chk_all_canal_linea"
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "BONIFICACION CUZADA"
      TabPicture(1)   =   "FrmBonificacion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape3"
      Tab(1).Control(1)=   "Image1"
      Tab(1).Control(2)=   "lblproductocruzada"
      Tab(1).Control(3)=   "Label20"
      Tab(1).Control(4)=   "lblproducto_boni_cruzada"
      Tab(1).Control(5)=   "Label16"
      Tab(1).Control(6)=   "Label15"
      Tab(1).Control(7)=   "Label14"
      Tab(1).Control(8)=   "Label13"
      Tab(1).Control(9)=   "Label3"
      Tab(1).Control(10)=   "DtcUnidad"
      Tab(1).Control(11)=   "cmdEliminarobsequio"
      Tab(1).Control(12)=   "cmdagregar_boni"
      Tab(1).Control(13)=   "HfObsequios"
      Tab(1).Control(14)=   "cmdEliminar"
      Tab(1).Control(15)=   "cmdagregar"
      Tab(1).Control(16)=   "HfCruzada"
      Tab(1).Control(17)=   "DtcTipoCoberturaCruzada"
      Tab(1).Control(18)=   "DtpFinCruzada"
      Tab(1).Control(19)=   "cmdprocesarcruzadda"
      Tab(1).Control(20)=   "DtpInicruzada"
      Tab(1).Control(21)=   "txtidproductocruzado"
      Tab(1).Control(22)=   "txtcantidad_cruzada"
      Tab(1).Control(23)=   "txtid_producto_boni_cruzada"
      Tab(1).Control(24)=   "txtcantidad_boni_cruzada"
      Tab(1).Control(25)=   "txtdescripcioncruzada"
      Tab(1).Control(26)=   "chk_all_canal"
      Tab(1).Control(27)=   "DtcUnidadBoni"
      Tab(1).Control(28)=   "chk_bloqueo"
      Tab(1).Control(29)=   "txtVentas"
      Tab(1).ControlCount=   30
      TabCaption(2)   =   "PROMOCIONES Y DESCUENTOS"
      TabPicture(2)   =   "FrmBonificacion.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Shape5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Image2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label5"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label17"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label18"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "prg_avance"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "DtcAlmacen"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdProcesarDescuentos"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "DtcCategoria"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "DtpDescuento_fin"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "DtpDescuento_ini"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "chk_categoria_descuento"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtdescuentoCategoria"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).ControlCount=   13
      Begin VB.TextBox txtdescuentoCategoria 
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
         Left            =   6960
         TabIndex        =   79
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox chk_categoria_descuento 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "DESCUENTO CATEGORIA :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   840
         TabIndex        =   76
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txtVentas 
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
         Left            =   -65580
         TabIndex        =   68
         Top             =   1380
         Width           =   495
      End
      Begin VB.CheckBox chk_bloqueo 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "BLOQUEO A :"
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
         Height          =   300
         Left            =   -66840
         TabIndex        =   67
         Top             =   1380
         Width           =   1215
      End
      Begin VB.CheckBox chk_all_canal_linea 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "ALL CHANEL'S"
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
         Height          =   300
         Left            =   -68400
         TabIndex        =   66
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CheckBox chk_modelo 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "MODELO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -74400
         TabIndex        =   65
         Top             =   3420
         Width           =   1455
      End
      Begin VB.CheckBox chk_subfamilia 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "SUB FAMILIA  :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -74400
         TabIndex        =   64
         Top             =   2940
         Width           =   1455
      End
      Begin VB.CheckBox chk_Familia 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "FAMILIA   :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -74400
         TabIndex        =   63
         Top             =   2505
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo DtcUnidadBoni 
         Height          =   315
         Left            =   -66360
         TabIndex        =   61
         Top             =   5580
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
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
      Begin VB.CheckBox chk_all_canal 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "ALL CHANEL'S"
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
         Height          =   300
         Left            =   -68280
         TabIndex        =   60
         Top             =   1380
         Width           =   1335
      End
      Begin VB.CheckBox chk_bonificacion_monto 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "BONIFICACION POR MONTO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -67200
         TabIndex        =   57
         Top             =   2460
         Width           =   2295
      End
      Begin VB.TextBox txtidbonificacion 
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
         Left            =   -65880
         TabIndex        =   52
         Top             =   5100
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtDescripcionBonificacion 
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
         Left            =   -72810
         TabIndex        =   23
         Top             =   1980
         Width           =   4335
      End
      Begin VB.TextBox txtlinea 
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
         Left            =   -68370
         TabIndex        =   22
         Top             =   2505
         Width           =   1095
      End
      Begin VB.TextBox txtcantidadbonificacion 
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
         Left            =   -72810
         TabIndex        =   21
         Top             =   5580
         Width           =   1455
      End
      Begin VB.TextBox txtid_producto 
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
         Left            =   -72810
         TabIndex        =   20
         Top             =   5100
         Width           =   1455
      End
      Begin VB.TextBox txtcantidad 
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
         Left            =   -72810
         TabIndex        =   19
         Top             =   3900
         Width           =   1455
      End
      Begin VB.TextBox txtdescripcioncruzada 
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
         Left            =   -72690
         TabIndex        =   18
         Top             =   1620
         Width           =   4335
      End
      Begin VB.TextBox txtcantidad_boni_cruzada 
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
         Left            =   -65010
         TabIndex        =   17
         Top             =   5580
         Width           =   495
      End
      Begin VB.TextBox txtid_producto_boni_cruzada 
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
         Left            =   -72690
         TabIndex        =   16
         Top             =   5580
         Width           =   975
      End
      Begin VB.TextBox txtcantidad_cruzada 
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
         Left            =   -65010
         TabIndex        =   15
         Top             =   3660
         Width           =   495
      End
      Begin VB.TextBox txtidproductocruzado 
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
         Left            =   -72690
         TabIndex        =   14
         Top             =   3660
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DtpBonificacionIni 
         Height          =   300
         Left            =   -72810
         TabIndex        =   24
         Top             =   1020
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   136183809
         CurrentDate     =   43546
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   615
         Left            =   -65850
         TabIndex        =   25
         Top             =   5820
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmBonificacion.frx":0054
         PICN            =   "FrmBonificacion.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcLinea 
         Height          =   330
         Left            =   -72810
         TabIndex        =   26
         Top             =   2505
         Width           =   4335
         _ExtentX        =   7646
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
      Begin MSDataListLib.DataCombo DtcSubFamilia 
         Height          =   330
         Left            =   -72810
         TabIndex        =   27
         Top             =   2940
         Width           =   4335
         _ExtentX        =   7646
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
      Begin MSComCtl2.DTPicker DtpBonificacionFin 
         Height          =   300
         Left            =   -69930
         TabIndex        =   28
         Top             =   1020
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   136183809
         CurrentDate     =   43546
      End
      Begin MSDataListLib.DataCombo DtcCobertura 
         Height          =   330
         Left            =   -72810
         TabIndex        =   29
         Top             =   1500
         Width           =   4335
         _ExtentX        =   7646
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
      Begin MSComCtl2.DTPicker DtpInicruzada 
         Height          =   300
         Left            =   -72690
         TabIndex        =   30
         Top             =   660
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   136183809
         CurrentDate     =   43546
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesarcruzadda 
         Height          =   495
         Left            =   -64440
         TabIndex        =   31
         Top             =   6060
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmBonificacion.frx":36B8
         PICN            =   "FrmBonificacion.frx":36D4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpFinCruzada 
         Height          =   300
         Left            =   -69810
         TabIndex        =   32
         Top             =   660
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   136183809
         CurrentDate     =   43546
      End
      Begin MSDataListLib.DataCombo DtcTipoCoberturaCruzada 
         Height          =   330
         Left            =   -72690
         TabIndex        =   33
         Top             =   1140
         Width           =   4335
         _ExtentX        =   7646
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCruzada 
         Height          =   1575
         Left            =   -72690
         TabIndex        =   34
         Top             =   2025
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2778
         _Version        =   393216
         ForeColor       =   8388608
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
      Begin VitekeySoft.ChameleonBtn cmdagregar 
         Height          =   345
         Left            =   -64485
         TabIndex        =   35
         Top             =   3660
         Width           =   450
         _ExtentX        =   794
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmBonificacion.frx":6D1C
         PICN            =   "FrmBonificacion.frx":6D38
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdEliminar 
         Height          =   345
         Left            =   -63960
         TabIndex        =   53
         Top             =   2220
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
         BTYPE           =   5
         TX              =   "ELIMINAR"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmBonificacion.frx":908C
         PICN            =   "FrmBonificacion.frx":90A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfObsequios 
         Height          =   1215
         Left            =   -72690
         TabIndex        =   54
         Top             =   4260
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2143
         _Version        =   393216
         ForeColor       =   8388608
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
      Begin VitekeySoft.ChameleonBtn cmdagregar_boni 
         Height          =   345
         Left            =   -64440
         TabIndex        =   55
         Top             =   5580
         Width           =   450
         _ExtentX        =   794
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmBonificacion.frx":B3FC
         PICN            =   "FrmBonificacion.frx":B418
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdEliminarobsequio 
         Height          =   345
         Left            =   -63960
         TabIndex        =   56
         Top             =   4620
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   609
         BTYPE           =   5
         TX              =   "ELIMINAR"
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmBonificacion.frx":D76C
         PICN            =   "FrmBonificacion.frx":D788
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcModelo 
         Height          =   330
         Left            =   -72810
         TabIndex        =   59
         Top             =   3420
         Width           =   4335
         _ExtentX        =   7646
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
      Begin MSDataListLib.DataCombo DtcUnidad 
         Height          =   315
         Left            =   -66360
         TabIndex        =   62
         Top             =   3660
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
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
      Begin MSComCtl2.DTPicker DtpDescuento_ini 
         Height          =   300
         Left            =   3000
         TabIndex        =   72
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   136183809
         CurrentDate     =   43546
      End
      Begin MSComCtl2.DTPicker DtpDescuento_fin 
         Height          =   300
         Left            =   5535
         TabIndex        =   73
         Top             =   840
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
         Format          =   136183809
         CurrentDate     =   43546
      End
      Begin MSDataListLib.DataCombo DtcCategoria 
         Height          =   330
         Left            =   3030
         TabIndex        =   77
         Top             =   1680
         Width           =   3855
         _ExtentX        =   6800
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
      Begin VitekeySoft.ChameleonBtn cmdProcesarDescuentos 
         Height          =   615
         Left            =   4560
         TabIndex        =   78
         Top             =   3120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmBonificacion.frx":FADC
         PICN            =   "FrmBonificacion.frx":FAF8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   330
         Left            =   3000
         TabIndex        =   81
         Top             =   2160
         Width           =   3855
         _ExtentX        =   6800
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
      Begin MSComctlLib.ProgressBar prg_avance 
         Height          =   255
         Left            =   4560
         TabIndex        =   83
         Top             =   3840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUCURSAL :"
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
         Left            =   2040
         TabIndex        =   82
         Top             =   2280
         Width           =   780
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INICIAL :"
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
         Left            =   1800
         TabIndex        =   75
         Top             =   885
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FINAL :"
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
         Left            =   4575
         TabIndex        =   74
         Top             =   885
         Width           =   930
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   11880
         Picture         =   "FrmBonificacion.frx":13140
         Top             =   720
         Width           =   240
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   6135
         Left            =   120
         Top             =   480
         Width           =   12255
      End
      Begin VB.Label Label3 
         BackColor       =   &H008080FF&
         Caption         =   "VENTAS"
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
         Height          =   300
         Left            =   -65040
         TabIndex        =   69
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO COBERTURA:"
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
         Left            =   -74430
         TabIndex        =   51
         Top             =   1500
         Width           =   1200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FINAL :"
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
         Left            =   -71010
         TabIndex        =   50
         Top             =   1020
         Width           =   930
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INICIAL :"
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
         Left            =   -74265
         TabIndex        =   49
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCION :"
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
         Left            =   -74220
         TabIndex        =   48
         Top             =   1980
         Width           =   990
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD :"
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
         Left            =   -74010
         TabIndex        =   47
         Top             =   5580
         Width           =   780
      End
      Begin VB.Label lblproducto 
         BackColor       =   &H00808080&
         Caption         =   " "
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
         Height          =   315
         Left            =   -71250
         TabIndex        =   46
         Top             =   5100
         Width           =   5250
      End
      Begin VB.Label Label7 
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
         Left            =   -74070
         TabIndex        =   45
         Top             =   5100
         Width           =   840
      End
      Begin VB.Label Label6 
         BackColor       =   &H008080FF&
         Caption         =   "PRODUCTO A BONIFICAR"
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
         Left            =   -72810
         TabIndex        =   44
         Top             =   4740
         Width           =   6780
      End
      Begin VB.Label lblmonto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD :"
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
         Left            =   -74010
         TabIndex        =   43
         Top             =   3900
         Width           =   780
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO COBERTURA:"
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
         Left            =   -74310
         TabIndex        =   42
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA FINAL :"
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
         Left            =   -70890
         TabIndex        =   41
         Top             =   660
         Width           =   930
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA INICIAL :"
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
         Left            =   -74145
         TabIndex        =   40
         Top             =   660
         Width           =   1035
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRIPCION :"
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
         Left            =   -74100
         TabIndex        =   39
         Top             =   1620
         Width           =   990
      End
      Begin VB.Label lblproducto_boni_cruzada 
         BackColor       =   &H00808080&
         Caption         =   " "
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
         Height          =   315
         Left            =   -71640
         TabIndex        =   38
         Top             =   5580
         Width           =   5250
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCTO A BONIFICAR"
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
         Left            =   -72690
         TabIndex        =   37
         Top             =   4020
         Width           =   1680
      End
      Begin VB.Label lblproductocruzada 
         BackColor       =   &H00808080&
         Caption         =   " "
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
         Height          =   315
         Left            =   -71640
         TabIndex        =   36
         Top             =   3660
         Width           =   5250
      End
      Begin VB.Image IMG_CLOSE 
         Height          =   240
         Left            =   -63000
         Picture         =   "FrmBonificacion.frx":15FE4
         Top             =   780
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   -63000
         Picture         =   "FrmBonificacion.frx":18E88
         Top             =   780
         Width           =   240
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   6285
         Left            =   -74880
         Top             =   420
         Width           =   12255
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   6135
         Left            =   -74880
         Top             =   600
         Width           =   12255
      End
   End
   Begin VB.OptionButton OptBonificacionCruzada 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "BONIFICACIONES CRUZADA"
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
      Height          =   250
      Left            =   5040
      TabIndex        =   12
      Top             =   700
      Width           =   2775
   End
   Begin VB.OptionButton optBonificacion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "BONIFICACIONES POR LINEA"
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
      Height          =   250
      Left            =   5040
      TabIndex        =   11
      Top             =   410
      Value           =   -1  'True
      Width           =   2775
   End
   Begin VB.TextBox txtDescripcion 
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
      Left            =   1560
      TabIndex        =   9
      Top             =   600
      Width           =   2535
   End
   Begin VitekeySoft.ChameleonBtn cmdbuscar 
      Height          =   345
      Left            =   15360
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
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
      MICON           =   "FrmBonificacion.frx":1BD2C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   345
      Left            =   12000
      TabIndex        =   1
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   136183809
      CurrentDate     =   43141
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   7215
      Left            =   120
      TabIndex        =   2
      Top             =   1470
      Width           =   17055
      _ExtentX        =   30083
      _ExtentY        =   12726
      _Version        =   393216
      ForeColor       =   8388608
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
   Begin VitekeySoft.ChameleonBtn cmdAnularOrden 
      Height          =   900
      Left            =   17280
      TabIndex        =   3
      Top             =   3330
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
      BTYPE           =   5
      TX              =   "ANULAR "
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
      MICON           =   "FrmBonificacion.frx":1BD48
      PICN            =   "FrmBonificacion.frx":1BD64
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
      Height          =   900
      Left            =   17280
      TabIndex        =   4
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmBonificacion.frx":1C07E
      PICN            =   "FrmBonificacion.frx":1C09A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdEditable 
      Height          =   900
      Left            =   17280
      TabIndex        =   5
      Top             =   2385
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
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
      MICON           =   "FrmBonificacion.frx":1C4EC
      PICN            =   "FrmBonificacion.frx":1C508
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrarpantalla 
      Height          =   900
      Left            =   17280
      TabIndex        =   6
      Top             =   5205
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
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
      MICON           =   "FrmBonificacion.frx":1C822
      PICN            =   "FrmBonificacion.frx":1C83E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DtpFin 
      Height          =   345
      Left            =   13680
      TabIndex        =   7
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   136183809
      CurrentDate     =   43141
   End
   Begin VitekeySoft.ChameleonBtn cmdEliminarBonificacion 
      Height          =   900
      Left            =   17280
      TabIndex        =   58
      Top             =   4275
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1588
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmBonificacion.frx":1F865
      PICN            =   "FrmBonificacion.frx":1F881
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcBusquedaCobertura 
      Height          =   330
      Left            =   8880
      TabIndex        =   70
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COBERTURA :"
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
      Height          =   195
      Left            =   7920
      TabIndex        =   71
      Top             =   720
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION :"
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
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BONIFICACIONES"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   45
      Width           =   1680
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   975
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   17055
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   8880
      Left            =   0
      Top             =   0
      Width           =   18480
   End
End
Attribute VB_Name = "FrmBonificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub ChameleonBtn2_Click()

End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub chk_bonificacion_monto_Click()
If Me.chk_bonificacion_monto.Value = 1 Then
   Me.lblmonto.Caption = "MONTO :"
Else
   Me.lblmonto.Caption = "CANTIDAD :"
End If
End Sub

Private Sub cmdagregar_boni_Click()
If Val(Me.txtidbonificacion.Text) = 0 Then
    MsgBox "Primero Guarde la Bonificacion" + Chr(13) + "Luego se Ingresa los productos.", vbInformation
    Exit Sub
End If

If Trim(Me.txtid_producto_boni_cruzada.Text) <> "" And Val(Me.txtcantidad_boni_cruzada.Text) > 0 Then
    strCadena = "INSERT INTO bonificacion_cruzada_detalle(`id_bonificacion`,`id_producto`,detalle,id_unidad,`cantidad`,`ruc`)VALUES " & _
    "('" & Val(Me.txtidbonificacion.Text) & "','" & Trim(Me.txtid_producto_boni_cruzada.Text) & "','" & Trim(Me.lblproducto_boni_cruzada.Caption) & "','" & Me.DtcUnidadBoni.BoundText & "','" & Val(Me.txtcantidad_boni_cruzada.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Call bonificacion(Me.HfObsequios)
    
End If
End Sub

Private Sub cmdagregar_Click()






If Val(Me.txtidbonificacion.Text) = 0 Then
    MsgBox "Primero Guarde la Bonificacion" + Chr(13) + "Luego se Ingresa los productos.", vbInformation
    Exit Sub
End If

If Trim(Me.txtidproductocruzado.Text) <> "" And Val(Me.txtcantidad_cruzada.Text) > 0 Then
    strCadena = "INSERT INTO bonificacion_detalle(`id_bonificacion`,`id_producto`,id_unidad,`cantidad`,`ruc`)VALUES " & _
    "('" & Val(Me.txtidbonificacion.Text) & "','" & Trim(Me.txtidproductocruzado.Text) & "','" & Me.DtcUnidad.BoundText & "','" & Val(Me.txtcantidad_cruzada.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Call detalle(Me.HfCruzada)
    
End If


End Sub

Private Sub cmdAnularOrden_Click()
If MsgBox("Desea Anular esta Bonificacion", vbYesNo + vbQuestion, KEY_VENDEDOR) = vbYes Then
    strCadena = "UPDATE bonificacion SET anulado='si' WHERE id_bonificacion='" & Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     strCadena = "SELECT * FROM view_bonificacion WHERE ruc='" & KEY_RUC & "'"
    Call actualizar(Me.HfgDetalle)
End If
End Sub

Private Sub cmdBuscar_Click()
 
 
 
 
If Me.optBonificacion.Value = True Then
 
    strCadena = "SELECT * FROM view_bonificacion WHERE fecha_ini>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_fin<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and   ruc='" & KEY_RUC & "'"
 
 Else
    strCadena = "SELECT * FROM view_bonificacion_cruzada WHERE fecha_ini>='" & Format(Me.DtpInicio.Value, "YYYY-mm-dd") & "' and fecha_fin<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  ruc='" & KEY_RUC & "'"
 
 End If
  Call actualizar(Me.HfgDetalle)
End Sub

Private Sub cmdCerrarpantalla_Click()
Unload Me

End Sub

Public Sub get_unidad(ByVal in_producto As String, ByVal in_agranel As String, ByVal in_combo As DataCombo)
    If in_agranel = "si" Then
        strCadena = "SELECT id_unidad as Codigo,descripcion as Descripcion FROM view_unidad_producto WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    Else
        strCadena = "SELECT id_unidad as Codigo,descripcion as Descripcion FROM view_unidad WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    End If
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(in_combo)
End Sub

Private Sub cmdEditable_Click()
Call load_bonificacion(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0))
End Sub

Private Sub cmdEliminar_Click()

strCadena = "DELETE FROM bonificacion_detalle WHERE id='" & Val(Me.HfCruzada.TextMatrix(Me.HfCruzada.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
Call detalle(Me.HfCruzada)



End Sub

Private Sub cmdEliminarBonificacion_Click()
If MsgBox("Desea Eliminar esta Bonificacion", vbYesNo + vbQuestion, KEY_VENDEDOR) = vbYes Then
    strCadena = "DELETE FROM  bonificacion WHERE id_bonificacion='" & Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    strCadena = "SELECT * FROM view_bonificacion WHERE ruc='" & KEY_RUC & "'"
    Call actualizar(Me.HfgDetalle)
    
End If
End Sub

Private Sub cmdEliminarobsequio_Click()
strCadena = "DELETE FROM bonificacion_cruzada_detalle WHERE id='" & Val(Me.HfObsequios.TextMatrix(Me.HfObsequios.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
Call bonificacion(Me.HfObsequios)

End Sub

Private Sub cmdNuevo_Click()
Call nuevo
Me.frmdetalle.Visible = True
End Sub

Private Sub cmdProcesar_Click()

If Trim(Me.txtid_producto.Text) <> "" And Val(Me.txtcantidad.Text) > 0 And Val(Me.txtcantidadbonificacion.Text) > 0 Then
        
        
        If Me.chk_bonificacion_monto.Value = 1 Then
           in_por_monto = "si"
           in_por_cantidad = "no"
        Else
            in_por_monto = "no"
            in_por_cantidad = "si"
        End If
        
        If Me.chk_Familia.Value = 1 Then
           in_linea = Me.DtcLinea.BoundText
        Else
           in_linea = 0
        End If
        If Me.chk_subfamilia.Value = 1 Then
           in_subfamilia = Me.DtcSubFamilia.BoundText
        Else
           in_subfamilia = 0
        End If
        
        If Me.chk_modelo.Value = 1 Then
           in_modelo = Me.DtcModelo.BoundText
        Else
           in_modelo = 0
        End If
        
        If Me.chk_all_canal_linea.Value = 1 Then
            in_all_canal = "si"
        Else
            in_all_canal = "no"
        End If
        
        
        
        
        strCadena = "call put_bonificacion_v3('" & Val(Me.txtidbonificacion.Text) & "','" & Format(Me.DtpBonificacionIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpBonificacionFin.Value, "YYYY-mm-dd") & "','" & in_linea & "','" & in_subfamilia & "','" & in_modelo & "','" & Trim(Me.txtDescripcionBonificacion.Text) & "','" & Val(Me.txtcantidad.Text) & "','" & Trim(Me.txtid_producto.Text) & "','" & Me.DtcCobertura.BoundText & "','" & Val(Me.txtcantidadbonificacion.Text) & "','" & KEY_USUARIO & "','no','" & in_por_cantidad & "','" & in_por_monto & "','" & in_all_canal & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        Me.frmdetalle.Visible = False
        strCadena = "SELECT * FROM view_bonificacion WHERE ruc='" & KEY_RUC & "'"
        Call actualizar(Me.HfgDetalle)
        
        
End If


End Sub
Private Sub nuevo()

Me.DtpInicruzada.Value = KEY_FECHA
Me.DtpFinCruzada.Value = KEY_FECHA

Me.txtidbonificacion.Text = ""
Me.txtDescripcionBonificacion.Text = ""
Me.txtcantidad.Text = 0
Me.txtcantidadbonificacion.Text = 0
Me.txtid_producto.Text = ""
Me.txtcantidad_boni_cruzada.Text = 0
Me.txtid_producto_boni_cruzada.Text = ""
Me.lblproducto.Caption = ""
Me.txtdescripcioncruzada.Text = ""
Me.HfCruzada.Rows = 0
Me.HfObsequios.Rows = 0
If Me.optBonificacion.Value = True Then
    Me.frmdetalle.Tab = 0
Else
    Me.frmdetalle.Tab = 1
End If


End Sub
Private Sub load_bonificacion(ByVal in_bonificacion As Double)


If Me.optBonificacion.Value = True Then
strCadena = "SELECT * FROM view_bonificacion WHERE id_bonificacion='" & Val(in_bonificacion) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   
   Me.DtcTipoCoberturaCruzada.BoundText = rst("id_cobertura")
   Me.txtidbonificacion.Text = rst("id_bonificacion")
   Me.DtpBonificacionIni.Value = rst("fecha_ini")
   Me.DtpBonificacionFin.Value = rst("fecha_fin")
   Me.txtDescripcionBonificacion.Text = rst("descripcion")
   Me.DtcLinea.BoundText = rst("id_linea")
   If rst("id_linea") <> "0" Then
        Me.chk_Familia.Value = 1
   Else
        Me.chk_Familia.Value = 0
   End If
   
   If rst("limite_ventas") > 0 Then
      Me.chk_bloqueo.Value = 1
      Me.txtVentas.Text = rst("limite_ventas")
   Else
      Me.chk_bloqueo.Value = 0
      Me.txtVentas.Text = 0
   End If
   
   If rst("id_subfamilia") <> "0" Then
      Me.chk_subfamilia.Value = 1
   Else
      Me.chk_subfamilia.Value = 0
   End If
   
   If rst("id_modelo") <> "0" Then
      Me.chk_modelo.Value = 1
   Else
      Me.chk_modelo.Value = 0
   End If
   
   
   Call load_sub(Me.DtcLinea.BoundText, rst("id_subfamilia"))
   Me.DtcModelo.BoundText = rst("id_modelo")
   Me.txtcantidad.Text = rst("cantidad")
   Me.txtid_producto.Text = rst("id_producto")
   Me.lblproducto.Caption = rst("nombre_prod")
   If rst("all_canal") = "si" Then
      Me.chk_all_canal_linea.Value = 1
   Else
     Me.chk_all_canal_linea.Value = 0
   End If
   
   Me.txtcantidadbonificacion.Text = rst("cantidad_bonificacion")
   Me.frmdetalle.Tab = 0
   
   If rst("por_cantidad") = "si" Then
      Me.chk_bonificacion_monto.Value = 0
   Else
     Me.chk_bonificacion_monto.Value = 1
   End If
   
   Me.frmdetalle.Visible = True
End If

End If



If Me.OptBonificacionCruzada.Value = True Then
strCadena = "SELECT * FROM view_bonificacion_cruzada WHERE id_bonificacion='" & Val(in_bonificacion) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.txtidbonificacion.Text = rst("id_bonificacion")
   Me.DtcTipoCoberturaCruzada.BoundText = rst("id_cobertura")
   Me.DtpInicruzada.Value = rst("fecha_ini")
   Me.DtpFinCruzada.Value = rst("fecha_fin")
   Me.txtdescripcioncruzada.Text = rst("descripcion")
   
   Me.txtid_producto_boni_cruzada.Text = rst("id_producto")
   Me.lblproducto_boni_cruzada.Caption = rst("nombre_prod")
   
   If rst("all_canal") = "si" Then
      Me.chk_all_canal.Value = 1
   Else
      Me.chk_all_canal.Value = 0
   End If
   
   If rst("limite_ventas") > 0 Then
      Me.chk_bloqueo.Value = 1
      Me.txtVentas.Text = rst("limite_ventas")
   Else
      Me.chk_bloqueo.Value = 0
      Me.txtVentas.Text = 0
   End If
   
   Me.txtcantidad_boni_cruzada.Text = rst("cantidad_bonificacion")
   Me.frmdetalle.Tab = 1
   Call detalle(Me.HfCruzada)
   Call bonificacion(Me.HfObsequios)
   
   Me.frmdetalle.Visible = True
End If
End If

If Me.OptDescuentos.Value = True Then
   
   strCadena = "CALL put_bonificacion('4','" & Val(in_bonificacion) & "','0','0','" & KEY_USUARIO & "','" & KEY_RUC & "')"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
    Me.txtidbonificacion.Text = rst("id_bonificacion")
    Me.chk_categoria_descuento.Value = 1
    Me.DtcTipoCoberturaCruzada.BoundText = rst("id_cobertura")
    Me.DtpDescuento_ini.Value = rst("fecha_ini")
    Me.DtpDescuento_fin.Value = rst("fecha_fin")
    Me.DtcCategoria.BoundText = rst("id_linea")
    Me.txtdescuentoCategoria.Text = rst("porcentaje_descuento")
    Me.DtcAlmacen.BoundText = rst("id_alm")
    If rst("all_canal") = "si" Then
      Me.chk_all_canal.Value = 1
    Else
      Me.chk_all_canal.Value = 0
    End If
   
   
   Me.frmdetalle.Visible = True

   Me.frmdetalle.Tab = 2
   End If
End If

End Sub
Public Sub actualizar(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub

End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 2500
           Grilla.ColWidth(4) = 1500
           Grilla.ColWidth(5) = 1500
           Grilla.ColWidth(6) = 1500
           Grilla.ColWidth(7) = 900
           Grilla.ColWidth(8) = 3500
           Grilla.ColWidth(9) = 900
           Grilla.ColWidth(10) = 1500
        Next
         cabecera = "CODIGO" & vbTab & "FECHA INI" & vbTab & "FECHA FIN" & vbTab & "DESCRIPCION" & vbTab & "CANAL" & vbTab & "LINEA" & vbTab & "SUB-LINEA" & vbTab & "CANTIDAD" & vbTab & "PRODUCTO" & vbTab & "CANTIDAD" & vbTab & "OPERADOR"
         Grilla.AddItem cabecera
         For k = 0 To 10
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
         Next k
       
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
        
            If rst("all_canal") = "si" Then
               in_canal = "ALL CHANEL'S"
            Else
                in_canal = rst("canal")
            End If
                
                
            Fila = Format(rst("id_bonificacion"), "00000") & vbTab & Format(rst("fecha_ini"), "dd-mm-YYYY") & vbTab & Format(rst("fecha_fin"), "dd-mm-YYYY") & vbTab & rst("descripcion") & vbTab & in_canal & vbTab & rst("linea") & vbTab & rst("modelo") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & rst("nombre_prod") & vbTab & Format(rst("cantidad_bonificacion"), "#,##0.00") & vbTab & rst("nombre_completo")
            Grilla.AddItem Fila
            
            If rst("anulado") = "si" Then
               
               
                For k = 6 To 9
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
            
             
            End If
            
            
            rst.MoveNext
        Next i
        
       
         
          
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub

Public Sub detalle(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Dim in_acumulado As Single
strCadena = "SELECT * FROM view_bonificacion_detalle WHERE id_bonificacion='" & Val(Me.txtidbonificacion.Text) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub

End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 4500
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 800
           
        Next
        cabecera = "ID" & vbTab & "CODIGO" & vbTab & "PRODUCTO" & vbTab & "UNIDAD" & vbTab & "CANTIDAD"
        Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
          Next k
       
        rst.MoveFirst
        in_acumulado = 0
        For i = 0 To rst.RecordCount - 1
            
            Fila = rst("id") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("unidad") & vbTab & rst("cantidad")
            Grilla.AddItem Fila
            in_acumulado = in_acumulado + rst("cantidad")
            rst.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "" & vbTab & "ACUMULADO:" & vbTab & Format(in_acumulado, "#,##0.00")
        Grilla.AddItem Fila
        Me.txtcantidad_cruzada.Text = in_acumulado
       
         
          
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub
Public Sub bonificacion(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
strCadena = "SELECT * FROM view_bonificacion_detalle_cruzada  WHERE id_bonificacion='" & Val(Me.txtidbonificacion.Text) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub

End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 4500
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 800
           
        Next
        cabecera = "ID" & vbTab & "CODIGO" & vbTab & "PRODUCTO" & vbTab & "UNIDAD" & vbTab & "CANTIDAD"
        Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
          Next k
       
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            
            Fila = rst("id") & vbTab & rst("id_producto") & vbTab & rst("detalle") & vbTab & rst("unidad") & vbTab & rst("cantidad")
            Grilla.AddItem Fila
            
            rst.MoveNext
        Next i
        
       
         
          
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub


Private Sub cmdprocesarcruzadda_Click()

If Me.chk_all_canal.Value = 1 Then
    in_all_canal = "si"
Else
    in_all_canal = "no"
End If

If Trim(Me.txtid_producto_boni_cruzada.Text) <> "" And Val(Me.txtcantidad_boni_cruzada.Text) > 0 Then
        
        strCadena = "call put_bonificacion_v3('" & Val(Me.txtidbonificacion.Text) & "','" & Format(Me.DtpInicruzada.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFinCruzada.Value, "YYYY-mm-dd") & "','0','0','0','" & Trim(Me.txtdescripcioncruzada.Text) & "','" & Me.txtcantidad_cruzada.Text & "','" & Trim(Me.txtid_producto_boni_cruzada.Text) & "','" & Me.DtcTipoCoberturaCruzada.BoundText & "','" & Val(Me.txtcantidad_boni_cruzada.Text) & "','" & KEY_USUARIO & "','si','no','no','" & in_all_canal & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        If Val(rst("_bonificacion")) > 0 Then
            If Me.chk_bloqueo.Value = 1 And Val(Me.txtVentas.Text) > 0 Then
                strCadena = "CALL put_bonificacion('1','" & rst("_bonificacion") & "','" & Val(Me.txtVentas.Text) & "','0','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            Else
                strCadena = "CALL put_bonificacion('1','" & rst("_bonificacion") & "','0','0','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            End If
        End If
        
        Me.frmdetalle.Visible = False
        
        
        
        Me.OptBonificacionCruzada.Value = True
        strCadena = "SELECT * FROM view_bonificacion_cruzada WHERE ruc='" & KEY_RUC & "'"
        Call actualizar(Me.HfgDetalle)
End If







End Sub


Private Sub load_sub(ByVal in_linea As String, ByVal in_subfamilia As String)
strCadena = "SELECT id_tipo as Codigo,descripcion as Descripcion FROM linea_sub WHERE id_linea='" & in_linea & "' and   id_usu='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcSubFamilia)
    Me.DtcSubFamilia.BoundText = in_subfamilia
    Call load_modelo(in_linea, Me.DtcSubFamilia.BoundText)
End Sub

Private Sub load_modelo(ByVal in_linea As String, ByVal in_sublinea As String)
strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM linea_modelo WHERE id_linea='" & in_linea & "' and id_sublinea='" & in_sublinea & "' and   ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcModelo)
End Sub

Private Sub cmdProcesarDescuentos_Click()

in_all_canal = "si"

in_categoria_descuento = ""

If Me.chk_categoria_descuento.Value = 1 Then
   in_categoria_descuento = Me.DtcCategoria.BoundText
   in_descripcion = Me.DtcCategoria.Text + Space(2) + "DESCUENTO " + str(Me.txtdescuentoCategoria.Text) + "%"
End If



If in_categoria_descuento = "" Then
    MsgBox "Marque una opcion de Descuento", vbInformation
    Exit Sub
End If



        
        strCadena = "call put_bonificacion_v3('" & Val(Me.txtidbonificacion.Text) & "','" & Format(Me.DtpBonificacionIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpBonificacionFin.Value, "YYYY-mm-dd") & "','" & in_categoria_descuento & "','0','0','" & in_descripcion & "','0','0','" & in_all_canal & "','0','" & KEY_USUARIO & "','no','no','no','" & in_all_canal & "','" & KEY_RUC & "')"
        Call ConfiguraRst(strCadena)
        If Val(rst("_bonificacion")) > 0 Then
            strCadena = "UPDATE bonificacion SET id_alm='" & Me.DtcAlmacen.BoundText & "',descuento='si',descripcion='" & in_descripcion & "',porcentaje_descuento='" & Val(Me.txtdescuentoCategoria.Text) & "' WHERE id_bonificacion='" & Val(rst("_bonificacion")) & "'"
            CnBd.Execute (strCadena)
        End If
        
       
        
        
        
        strCadena = "CALL put_bonificacion('3','0','0','0','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Call actualizar(Me.HfgDetalle)
        
        
        
     
          
           strCadena = "SELECT * FROM view_producto WHERE id_alm='" & Me.DtcAlmacen.BoundText & "' and  id_linea='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "'"
           Call ConfiguraRst(strCadena)
           If rst.RecordCount > 0 Then
              rst.MoveFirst
              Me.prg_avance.Min = 0
              Me.prg_avance.Max = rst.RecordCount
              For i = 0 To rst.RecordCount - 1
                    If Format(Me.DtpDescuento_fin.Value, "YYYY-mm-dd") < KEY_FECHA Then
                        in_precio_venta = rst("precio_venta") + rst("precio_venta") * Val(Me.txtdescuentoCategoria.Text) / 100
                    Else
                        in_precio_venta = rst("precio_venta") - rst("precio_venta") * Val(Me.txtdescuentoCategoria.Text) / 100
                    End If
                    strCadena = "UPDATE almacen_producto SET precio_venta='" & in_precio_venta & "' WHERE id_producto='" & rst("id_producto") & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    rst.MoveNext
                    DoEvents
                    Me.prg_avance.Value = i
              Next i
              MsgBox "Descuento Realizado Correctamente", vbInformation
           End If
     
       Me.frmdetalle.Visible = False
          
        

End Sub

Private Sub DtcBusquedaCobertura_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   
   If Me.optBonificacion.Value = True Then
        strCadena = "SELECT * FROM view_bonificacion WHERE id_cobertura='" & Me.DtcBusquedaCobertura.BoundText & "' and  ruc='" & KEY_RUC & "'"
   Else
        strCadena = "SELECT * FROM view_bonificacion_cruzada WHERE id_cobertura='" & Me.DtcBusquedaCobertura.BoundText & "' and  ruc='" & KEY_RUC & "'"
   End If
   
   Call actualizar(Me.HfgDetalle)
   
End If
End Sub

Private Sub DtcLinea_Change()
 strCadena = "SELECT id_tipo as Codigo, descripcion as Descripcion FROM linea_sub WHERE id_usu='" & KEY_RUC & "' AND id_linea='" & Me.DtcLinea.BoundText & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtcSubFamilia)
End Sub

Private Sub DtcSubFamilia_Change()
If Me.DtcSubFamilia.BoundText <> "" Then
    strCadena = "SELECT id as Codigo, descripcion as Descripcion FROM linea_modelo WHERE ruc='" & KEY_RUC & "' AND id_linea='" & Me.DtcLinea.BoundText & "' and id_sublinea='" & Me.DtcSubFamilia.BoundText & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtcModelo)
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 150

Me.DtpBonificacionIni.Value = KEY_FECHA
Me.DtpBonificacionFin.Value = KEY_FECHA
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA

Me.DtpBonificacionIni.Value = KEY_FECHA
Me.DtpBonificacionFin.Value = KEY_FECHA

strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen  WHERE id_sucursal='0' and  ruc='" & KEY_RUC & "'  ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM
  
  
  
strCadena = "SELECT id_linea as Codigo,descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcLinea)

strCadena = "SELECT id_linea as Codigo,descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcCategoria)



 strCadena = "SELECT id_tipo_cliente as Codigo,descripcion as Descripcion FROM tipo_cliente ORDER BY id_tipo_cliente ASC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcBusquedaCobertura)


 
 strCadena = "SELECT id_tipo_cliente as Codigo,descripcion as Descripcion FROM tipo_cliente ORDER BY id_tipo_cliente ASC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(DtcCobertura)


 strCadena = "SELECT id_tipo_cliente as Codigo,descripcion as Descripcion FROM tipo_cliente ORDER BY id_tipo_cliente ASC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoCoberturaCruzada)

 
 strCadena = "SELECT * FROM view_bonificacion WHERE ruc='" & KEY_RUC & "'"
 Call actualizar(Me.HfgDetalle)




End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub HfCruzada_SelChange()
If Me.HfCruzada.Rows > 0 Then
If Val(Me.HfCruzada.TextMatrix(Me.HfCruzada.Row, 0)) > 0 Then
    Me.cmdEliminar.Enabled = True
Else
    Me.cmdEliminar.Enabled = False
End If
End If
End Sub

Private Sub HfgDetalle_SelChange()
If Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) > 0 Then
    Me.cmdEditable.Enabled = True
Else
    Me.cmdEditable.Enabled = False
End If
End Sub

Private Sub HfObsequios_SelChange()
If Val(Me.HfObsequios.TextMatrix(Me.HfObsequios.Row, 0)) > 0 Then
    Me.cmdEliminarobsequio.Enabled = True
    
Else
    Me.cmdEliminarobsequio.Enabled = False
End If
End Sub

Private Sub Image1_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub Image2_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub IMG_CLOSE_Click()
Me.frmdetalle.Visible = False
End Sub

Private Sub Text2_Change()

End Sub

Private Sub optBonificacion_Click()
 strCadena = "SELECT * FROM view_bonificacion WHERE ruc='" & KEY_RUC & "'"
 Call actualizar(Me.HfgDetalle)
End Sub

Private Sub OptBonificacionCruzada_Click()
 strCadena = "SELECT * FROM view_bonificacion_cruzada WHERE ruc='" & KEY_RUC & "'"
 Call actualizar(Me.HfgDetalle)
End Sub

Private Sub OptDescuentos_Click()
 strCadena = "CALL put_bonificacion('3','0','0','0','" & KEY_USUARIO & "','" & KEY_RUC & "')"
 Call actualizar(Me.HfgDetalle)
End Sub

Private Sub txtid_producto_boni_cruzada_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = relacionar
    FrmProducto.Show
    Exit Sub
End If
End Sub

Private Sub txtid_producto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
End If
End Sub

Private Sub txtidpro_Change()

End Sub

Private Sub txtidproducto_Change()

End Sub

Private Sub txtidproducto_KeyPress(KeyAscii As Integer)
End Sub

Private Sub txtidproductocruzado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = buscar
    FrmProducto.Show
    Exit Sub
End If

End Sub

Private Sub TxtLinea_Change()
strCadena = "SELECT id_linea as Codigo,descripcion as Descripcion FROM linea WHERE descripcion LIKE '%" & Trim(Me.txtlinea.Text) & "%' and  id_usu='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcLinea)
End Sub

Private Sub txtlinea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcLinea.SetFocus
End If
End Sub

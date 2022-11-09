VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmDetalleLinea 
   BorderStyle     =   0  'None
   Caption         =   "Detalle Clasificacion"
   ClientHeight    =   9135
   ClientLeft      =   270
   ClientTop       =   105
   ClientWidth     =   10905
   ControlBox      =   0   'False
   Icon            =   "FrmDetalleLinea.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk_activado_online 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "ACTIVO ONLINE"
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
      Left            =   4800
      TabIndex        =   83
      Top             =   1560
      Width           =   1695
   End
   Begin MSDataListLib.DataCombo DtcClasificacion 
      Height          =   330
      Left            =   2760
      TabIndex        =   81
      Top             =   240
      Width           =   4125
      _ExtentX        =   7276
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
   Begin VitekeySoft.ChameleonBtn cmdAjustes 
      Height          =   1095
      Left            =   120
      TabIndex        =   75
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetalleLinea.frx":058A
      PICN            =   "FrmDetalleLinea.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frmConfiguracion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "CONFIGURACION"
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
      Height          =   2775
      Left            =   1080
      TabIndex        =   73
      Top             =   5160
      Visible         =   0   'False
      Width           =   6495
      Begin VB.Frame FrmConfiguracionDetalle 
         BackColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   3720
         TabIndex        =   76
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
         Begin VB.CheckBox chk_Habilitado 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "HABILITADO"
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
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   1005
            Width           =   1935
         End
         Begin VB.TextBox txtDescripcionConfiguracion 
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
            Left            =   120
            MaxLength       =   50
            TabIndex        =   78
            Top             =   600
            Width           =   1965
         End
         Begin VitekeySoft.ChameleonBtn cmdProcesarLinea 
            Height          =   375
            Left            =   120
            TabIndex        =   79
            Top             =   1320
            Width           =   1935
            _ExtentX        =   3413
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
            BCOL            =   8421631
            BCOLO           =   8421631
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDetalleLinea.frx":3DFC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   2040
            Picture         =   "FrmDetalleLinea.frx":3E18
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPCION "
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
            Left            =   480
            TabIndex        =   77
            Top             =   240
            Width           =   975
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfConfiguracion 
         Height          =   2175
         Left            =   240
         TabIndex        =   74
         Top             =   480
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3836
         _Version        =   393216
         ForeColor       =   8388608
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         ForeColorFixed  =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   -2147483635
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
      Begin VB.Image Image2 
         Height          =   240
         Left            =   6240
         Picture         =   "FrmDetalleLinea.frx":6CBC
         Top             =   120
         Width           =   240
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   1200
      TabIndex        =   58
      Top             =   1920
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2990
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "CONTABILIDAD"
      TabPicture(0)   =   "FrmDetalleLinea.frx":9B60
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "COSTO"
      TabPicture(1)   =   "FrmDetalleLinea.frx":9B7C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CONTABILIDAD"
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
         Height          =   1215
         Left            =   -74880
         TabIndex        =   66
         Top             =   360
         Width           =   6135
         Begin VB.TextBox txtCostoHaber 
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
            Left            =   1545
            MaxLength       =   50
            TabIndex        =   68
            Top             =   720
            Width           =   1365
         End
         Begin VB.TextBox txtCostoDebe 
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
            Left            =   1545
            MaxLength       =   50
            TabIndex        =   67
            Top             =   300
            Width           =   1365
         End
         Begin VB.Label lblcuentahabercosto 
            Caption         =   " "
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
            Left            =   3120
            TabIndex        =   72
            Top             =   675
            Width           =   2775
         End
         Begin VB.Label lblcuentadebecosto 
            Caption         =   " "
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
            Left            =   3120
            TabIndex        =   71
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA COSTO DEBE  :"
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
            Left            =   195
            TabIndex        =   70
            Top             =   360
            Width           =   1245
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA COSTO HABER :"
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
            TabIndex        =   69
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CONTABILIDAD"
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
         Height          =   1215
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   6135
         Begin VB.TextBox txtcodigo_contable 
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
            Left            =   1785
            MaxLength       =   50
            TabIndex        =   61
            Top             =   180
            Width           =   1245
         End
         Begin VB.TextBox txtCuentaContableImportacion 
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
            Left            =   1785
            MaxLength       =   50
            TabIndex        =   60
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA CONBALE  IMPORT :"
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
            Left            =   75
            TabIndex        =   65
            Top             =   720
            Width           =   1665
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CTA CONTABLE  :"
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
            Left            =   525
            TabIndex        =   64
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lblcuenta_contable 
            Caption         =   " "
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
            Left            =   3120
            TabIndex        =   63
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label lblcuenta_contable_import 
            Caption         =   " "
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
            Left            =   3120
            TabIndex        =   62
            Top             =   720
            Width           =   2775
         End
      End
   End
   Begin VB.CheckBox chk_motor 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "MOTOR"
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
      Left            =   2760
      TabIndex        =   57
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame frmplantilla 
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
      Height          =   2535
      Left            =   1200
      TabIndex        =   43
      Top             =   3960
      Visible         =   0   'False
      Width           =   6420
      Begin VB.Frame frmplantilla_detalle 
         BackColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   5295
         Begin VB.TextBox txtidPlantilla 
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
            Left            =   3240
            MaxLength       =   50
            TabIndex        =   56
            Top             =   960
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox txtProcentaje_detalle 
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
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   51
            Top             =   960
            Width           =   1245
         End
         Begin VB.TextBox txtDescripcion_plantilla 
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
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   50
            Top             =   360
            Width           =   3165
         End
         Begin VitekeySoft.ChameleonBtn cmdprocesar_plantilla 
            Height          =   375
            Left            =   1800
            TabIndex        =   52
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
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
            BCOL            =   8421631
            BCOLO           =   8421631
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDetalleLinea.frx":9B98
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdcerrar_plantilla 
            Height          =   375
            Left            =   3240
            TabIndex        =   53
            Top             =   1440
            Width           =   1215
            _ExtentX        =   2143
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
            BCOL            =   8421631
            BCOLO           =   8421631
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDetalleLinea.frx":9BB4
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
            Alignment       =   2  'Center
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
            Left            =   390
            TabIndex        =   49
            Top             =   1080
            Width           =   945
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPCION "
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
            Left            =   375
            TabIndex        =   48
            Top             =   360
            Width           =   975
         End
      End
      Begin VitekeySoft.ChameleonBtn cmdnueva_plantilla 
         Height          =   375
         Left            =   5565
         TabIndex        =   45
         Top             =   360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
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
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleLinea.frx":9BD0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPlantilla 
         Height          =   2055
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3625
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   -2147483635
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
      Begin VitekeySoft.ChameleonBtn cmdEliminar 
         Height          =   375
         Left            =   5565
         TabIndex        =   46
         Top             =   1365
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
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
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleLinea.frx":9BEC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdUpdate 
         Height          =   375
         Left            =   5565
         TabIndex        =   55
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "UPDATE"
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
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleLinea.frx":9C08
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
   Begin VB.CheckBox chk_plantilla 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "PLANTILLA"
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
      Height          =   255
      Left            =   2280
      TabIndex        =   42
      Top             =   8280
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   1080
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   6495
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
         Left            =   1560
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   255
         Left            =   5400
         TabIndex        =   10
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleLinea.frx":9C24
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
         Height          =   255
         Left            =   5400
         TabIndex        =   14
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleLinea.frx":9C40
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblcosto 
         Caption         =   " "
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
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblproducto 
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
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COSTO :"
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
         Left            =   690
         TabIndex        =   9
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label2 
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
         Left            =   225
         TabIndex        =   8
         Top             =   600
         Width           =   1005
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
         Left            =   630
         TabIndex        =   7
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Frame frminsumos 
      BackColor       =   &H00FFFFFF&
      Height          =   3675
      Left            =   1080
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   6420
      Begin VB.TextBox txtdias 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1080
         TabIndex        =   37
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtid_insumo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1920
         TabIndex        =   35
         Top             =   1395
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtcantidad 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1080
         TabIndex        =   33
         Top             =   1395
         Width           =   735
      End
      Begin VB.CheckBox chk_pagado 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "PAGADO"
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
         Left            =   3960
         TabIndex        =   28
         Top             =   1425
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfInsumos 
         Height          =   1575
         Left            =   240
         TabIndex        =   26
         Top             =   1800
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2778
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   -2147483635
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
      Begin VitekeySoft.ChameleonBtn cmdeliminarinsumo 
         Height          =   285
         Left            =   5760
         TabIndex        =   27
         Top             =   1800
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "FrmDetalleLinea.frx":9C5C
         PICN            =   "FrmDetalleLinea.frx":9C78
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdagregar 
         Height          =   465
         Left            =   5175
         TabIndex        =   29
         Top             =   1305
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   820
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
         MICON           =   "FrmDetalleLinea.frx":A212
         PICN            =   "FrmDetalleLinea.frx":A22E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdagregarinsumo 
         Height          =   285
         Left            =   240
         TabIndex        =   34
         Top             =   285
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   503
         BTYPE           =   5
         TX              =   "AGREGAR PRODUCTO"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleLinea.frx":A7C8
         PICN            =   "FrmDetalleLinea.frx":A7E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdcerrarinsumo 
         Height          =   285
         Left            =   4440
         TabIndex        =   36
         Top             =   360
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   503
         BTYPE           =   5
         TX              =   "CERRAR PANTALLA"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleLinea.frx":AD7E
         PICN            =   "FrmDetalleLinea.frx":AD9A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdupdatear 
         Height          =   280
         Left            =   1920
         TabIndex        =   39
         Top             =   720
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
         MICON           =   "FrmDetalleLinea.frx":DDAF
         PICN            =   "FrmDetalleLinea.frx":DDCB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº DIAS :"
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
         Height          =   165
         Left            =   240
         TabIndex        =   38
         Top             =   765
         Width           =   765
      End
      Begin VB.Label lblinsumo 
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
         Height          =   195
         Left            =   1080
         TabIndex        =   32
         Top             =   1125
         Width           =   3885
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "CANT :"
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
         Height          =   165
         Left            =   240
         TabIndex        =   31
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "INSUMO :"
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
         Height          =   165
         Left            =   240
         TabIndex        =   30
         Top             =   1125
         Width           =   765
      End
   End
   Begin VB.CheckBox chkgarantia 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "AFECTO A GARANTIA."
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
      Left            =   4800
      TabIndex        =   19
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame frmgarantia 
      BackColor       =   &H00FFFFFF&
      Height          =   3645
      Left            =   1200
      TabIndex        =   15
      Top             =   3600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txtkilometros 
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
         Left            =   2310
         TabIndex        =   40
         Top             =   480
         Width           =   735
      End
      Begin VitekeySoft.ChameleonBtn cmdgenrar 
         Height          =   285
         Left            =   3240
         TabIndex        =   23
         Top             =   735
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         BTYPE           =   5
         TX              =   "GEN MANT."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
         MICON           =   "FrmDetalleLinea.frx":E365
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtnumero_mantenimientos 
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
         Left            =   2310
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtperiodogarantia 
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
         Left            =   2310
         TabIndex        =   17
         Top             =   120
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Hfmantenimientos 
         Height          =   2175
         Left            =   240
         TabIndex        =   22
         Top             =   1320
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3836
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColor       =   -2147483635
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
      Begin VitekeySoft.ChameleonBtn cmdquitar 
         Height          =   285
         Left            =   5955
         TabIndex        =   24
         Top             =   1320
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   503
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         MICON           =   "FrmDetalleLinea.frx":E381
         PICN            =   "FrmDetalleLinea.frx":E39D
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
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº KILOMETROS :"
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
         Left            =   915
         TabIndex        =   41
         Top             =   525
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº DE MANTENIMIENTOS :"
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
         Left            =   195
         TabIndex        =   21
         Top             =   885
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MESES"
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
         Height          =   165
         Left            =   3240
         TabIndex        =   18
         Top             =   165
         Width           =   525
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIODO GARANTIA :"
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
         Left            =   585
         TabIndex        =   16
         Top             =   165
         Width           =   1665
      End
   End
   Begin VB.CheckBox chkproduccion 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "LINEA DE PRODUCCION"
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
      Left            =   2760
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   360
      Top             =   3720
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
            Picture         =   "FrmDetalleLinea.frx":E937
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLinea.frx":EC53
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLinea.frx":F0B3
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLinea.frx":F513
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLinea.frx":F82F
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLinea.frx":FC8F
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLinea.frx":FFAB
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLinea.frx":1040B
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLinea.frx":1086B
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLinea.frx":1114B
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLinea.frx":11467
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleLinea.frx":11783
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtDescripcion 
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
      Left            =   2760
      MaxLength       =   100
      TabIndex        =   0
      Top             =   720
      Width           =   4125
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   5925
      TabIndex        =   1
      Top             =   8130
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
      _Version        =   "6.7.9782"
      Child1          =   "TlbAcciones"
      MinHeight1      =   810
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   2
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfItem 
      Height          =   1215
      Left            =   1320
      TabIndex        =   5
      Top             =   1920
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2143
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      GridColor       =   -2147483635
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
   Begin VitekeySoft.ChameleonBtn cmdSiguienteImagen 
      Height          =   375
      Left            =   7680
      TabIndex        =   84
      Top             =   3480
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "SIGUIENTE"
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetalleLinea.frx":11A9F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdImagen 
      Height          =   280
      Left            =   6580
      TabIndex        =   85
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
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
      BCOL            =   33023
      BCOLO           =   33023
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetalleLinea.frx":11ABB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8880
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image img_foto 
      Height          =   3255
      Left            =   7750
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2980
   End
   Begin VB.Shape shape_image 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      Height          =   3375
      Left            =   7680
      Top             =   45
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   1395
      TabIndex        =   82
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label lblid_linea 
      BackColor       =   &H008080FF&
      Height          =   255
      Left            =   8040
      TabIndex        =   54
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LblDescripcion 
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
      Left            =   1545
      TabIndex        =   3
      Top             =   765
      Width           =   1005
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   9135
      Left            =   0
      Top             =   0
      Width           =   10905
   End
End
Attribute VB_Name = "FrmDetalleLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Dim StrCodTabla As String
Dim strCodLinea As String

Private Sub cmdImagen_Click()
   Me.img_foto.Visible = True
   Me.cmdSiguienteImagen.Visible = True
   Me.shape_image.Visible = True
   
   
   Call get_image_standart("00001")
   Me.cmdSiguienteImagen.Tag = "00001"
   
End Sub

Private Sub cmdSiguienteImagen_Click()

Call get_image_standart(Format(Val(Me.cmdSiguienteImagen.Tag) + 1, "00000"))
Me.cmdSiguienteImagen.Tag = Format(Val(Me.cmdSiguienteImagen.Tag) + 1, "00000")



End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)

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

Me.img_foto.Picture = LoadPicture(App.Path & "\" & nomArchivo)


'Image1.Picture = LoadPicture(App.Path & "\" & nomArchivo)

 Kill App.Path & "\" & nomArchivo



End Select



End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub chk_activado_online_Click()
If Me.chk_activado_online.Value = 1 Then
   Me.cmdImagen.Visible = True
Else
   Me.cmdImagen.Visible = False
End If


End Sub
Private Sub get_image(ByVal in_codigo As String)
On Error GoTo salir
Dim in_ruta As String

strCadena = "SELECT imagen FROM linea WHERE id_linea='" & Trim(in_codigo) & "' and id_usu='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
   If Len(rstIN("imagen")) > 10 Then
      in_ruta = rstIN("imagen")
   Else
      in_ruta = ""
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

Private Sub get_image_standart(ByVal in_codigo As String)
On Error GoTo salir
Dim in_ruta As String

strCadena = "SELECT standart,logo FROM linea_online WHERE id_linea='" & Trim(in_codigo) & "' LIMIT 1"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
   If Len(rstIN("standart")) > 10 Then
      in_ruta = rstIN("standart")
      Me.cmdImagen.Tag = Trim(in_ruta)
      Me.cmdImagen.ToolTipText = rstIN("logo")
   Else
      Me.cmdImagen.Tag = ""
      Me.cmdImagen.ToolTipText = ""
      in_ruta = ""
      Exit Sub
   End If
Else
    Me.cmdSiguienteImagen.Tag = "00001"
    Exit Sub
End If

With Inet1
    .AccessType = icUseDefault
    .Url = in_ruta
    .Execute , "GET"
End With
Exit Sub
salir:
MsgBox "Descargando Archivos", vbInformation
Exit Sub
End Sub

Private Sub chk_plantilla_Click()
If Me.chk_plantilla.Value = 1 Then
    Call llenar_plantilla(Me.HfPlantilla)
    Me.frmplantilla.Visible = True
    Me.frmplantilla_detalle.Visible = False
Else
    Me.frmplantilla.Visible = False
End If
End Sub

Private Sub chkgarantia_Click()
If Me.chkgarantia.Value = 1 Then
    Me.frmgarantia.Visible = True
     Call Me.llenar_mantenimientos(FrmLinea.HfgLinea.TextMatrix(FrmLinea.HfgLinea.Row, 0), Me.Hfmantenimientos)
Else
    Me.frmgarantia.Visible = False
End If
End Sub

Private Sub chkproduccion_Click()
Dim id_linea As String
id_linea = Trim(FrmLinea.HfgLinea.TextMatrix(FrmLinea.HfgLinea.Row, 0))

If Me.chkproduccion.Value = 1 Then
    strCadena = "SELECT id_produccion from linea_produccion WHERE ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       rst.MoveFirst
       For i = 0 To rst.RecordCount - 1
            strCadena = "SELECT count(*) FROM linea_produccion_detalle WHERE id_linea='" & id_linea & "' AND id_produccion='" & rst("id_produccion") & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRstT(strCadena)
            If rstT(0) < 1 Then
                strCadena = "INSERT INTO linea_produccion_detalle (id_linea,id_produccion,ruc)VALUES('" & id_linea & "','" & rst("id_produccion") & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                 
            End If
            rst.MoveNext
       Next i
    End If
    Call llenar_fases_produccion(id_linea, Me.HfItem)
    Me.HfItem.Visible = True
    Me.chk_motor.Visible = True
   
Else
    Me.HfItem.Visible = False
    Me.chk_motor.Visible = False
End If
End Sub
 Private Sub ActualizarImagen(ByVal id_produccion As String, ByVal id_linea As String)
     Dim estado As String
      strCadena = "SELECT estado FROM linea_produccion_detalle WHERE id_linea='" & id_linea & "' AND id_produccion='" & id_produccion & "' AND ruc='" & KEY_RUC & "'"
      Call ConfiguraRst(strCadena)
      If rst.RecordCount > 0 Then
        If rst("estado") = "si" Then
            estado = "no"
           Me.HfItem.TextMatrix(Me.HfItem.Row, 2) = Chr(168)
            For j = 0 To 2
                Me.HfItem.col = j
                HfItem.Row = Me.HfItem.Row
                HfItem.CellBackColor = &HFFFFFF
            Next j
            strCadena = "UPDATE linea_produccion_detalle SET estado='" & estado & "' WHERE id_linea='" & id_linea & "' AND id_produccion='" & id_produccion & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
            
            Me.Frame1.Visible = False
            Exit Sub
        Else
            estado = "si"
            Me.HfItem.TextMatrix(Me.HfItem.Row, 2) = Chr(254)
            For j = 0 To 2
                HfItem.col = j
                HfItem.Row = Me.HfItem.Row
                HfItem.CellBackColor = &HC0FFC0
            Next j
        End If
        strCadena = "UPDATE linea_produccion_detalle SET estado='" & estado & "' WHERE id_linea='" & id_linea & "' AND id_produccion='" & id_produccion & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
        Me.Frame1.Visible = True
      End If
      
      
      
    End Sub

Private Sub cmdagregar_Click()
Dim strpagado As String
If Me.chk_pagado.Value = 1 Then
    strpagado = "si"
Else
    strpagado = "no"
End If

strCadena = "SELECT * FROM linea_mantenimiento_detalle WHERE id_producto='" & Trim(Me.txtid_insumo.Text) & "' and id_mantenimiento='" & Val(FrmDetalleLinea.Hfmantenimientos.TextMatrix(FrmDetalleLinea.Hfmantenimientos.Row, 0)) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    strCadena = "INSERT INTO linea_mantenimiento_detalle (id_mantenimiento,id_producto,cantidad,pagado,ruc)VALUES('" & Val(Me.Hfmantenimientos.TextMatrix(Me.Hfmantenimientos.Row, 0)) & "','" & Trim(Me.txtid_insumo.Text) & "','" & Val(Me.txtcantidad.Text) & "','" & strpagado & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
     
    strCadena = "UPDATE linea_mantenimiento SET insumos='si' WHERE id_mantenimiento='" & Val(Me.Hfmantenimientos.TextMatrix(Me.Hfmantenimientos.Row, 0)) & "'"
    CnBd.Execute (strCadena)
     
     
    strCadena = "UPDATE linea_mantenimiento SET dias='" & Val(Me.txtdias.Text) & "' WHERE id_mantenimiento='" & Val(Me.Hfmantenimientos.TextMatrix(Me.Hfmantenimientos.Row, 0)) & "'"
    CnBd.Execute (strCadena)
     
     
    Call Me.llenar_insumos(Val(Me.Hfmantenimientos.TextMatrix(Me.Hfmantenimientos.Row, 0)), Me.HfInsumos)
End If
End Sub

Private Sub cmdagregarinsumo_Click()
Procedencia = seleccionar_insumo
FrmProducto.Show
Exit Sub
End Sub

Private Sub cmdAjustes_Click()

Me.frmConfiguracion.Visible = True
Call llenar_configuracion(Me.HfConfiguracion)

End Sub





Private Sub cmdCerrar_Click()
Me.Frame1.Visible = False
End Sub

Private Sub cmdcerrar_plantilla_Click()
Me.frmplantilla_detalle.Visible = False
End Sub

Private Sub cmdcerrarinsumo_Click()
Me.frminsumos.Visible = False
End Sub

Private Sub cmdEliminar_Click()
If Val(Me.HfPlantilla.TextMatrix(Me.HfPlantilla.Row, 0)) > 0 Then
    strCadena = "DELETE FROM linea_plantilla WHERE id='" & Val(Me.HfPlantilla.TextMatrix(Me.HfPlantilla.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Call Me.llenar_plantilla(Me.HfPlantilla)
End If
End Sub

Private Sub cmdeliminarinsumo_Click()
Procedencia = eliminar_insumo
FrmSeguridad.Show
Exit Sub
End Sub

Private Sub cmdgenrar_Click()
strCadena = "SELECT * FROM linea_mantenimiento WHERE id_linea='" & Trim(FrmLinea.HfgLinea.TextMatrix(FrmLinea.HfgLinea.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount <= Val(Me.txtnumero_mantenimientos.Text) Then
    For i = 0 To Val(Me.txtnumero_mantenimientos.Text) - rst.RecordCount - 1
        strCadena = "INSERT INTO linea_mantenimiento(id_linea,dias,kilometros,ruc)VALUES('" & Trim(FrmLinea.HfgLinea.TextMatrix(FrmLinea.HfgLinea.Row, 0)) & "','" & Val(Me.txtperiodogarantia.Text) & "','" & Val(Me.txtkilometros.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
         
    Next i

End If
Call llenar_mantenimientos(Trim(FrmLinea.HfgLinea.TextMatrix(FrmLinea.HfgLinea.Row, 0)), FrmDetalleLinea.Hfmantenimientos)
End Sub

Private Sub cmdnueva_plantilla_Click()
Me.frmplantilla_detalle.Visible = True
Me.txtDescripcion_plantilla.Text = ""
Me.txtProcentaje_detalle.Text = ""
Me.txtidPlantilla.Text = 0
Me.frmplantilla_detalle.Visible = True
End Sub

Private Sub cmdProcesar_Click()
If Trim(Me.txtid_producto.Text) <> "" And Me.LblDescripcion.Caption <> "" Then
    strCadena = "UPDATE linea_produccion_detalle SET id_producto='" & Trim(Me.txtid_producto.Text) & "' WHERE id_linea='" & FrmLinea.HfgLinea.TextMatrix(FrmLinea.HfgLinea.Row, 0) & "' AND id_produccion='" & Me.HfItem.TextMatrix(Me.HfItem.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
     CnBd.Execute (strCadena)
      
     Me.HfItem.TextMatrix(Me.HfItem.Row, 3) = Trim(Me.lblproducto.Caption)
     Me.Frame1.Visible = False
End If
End Sub

Private Sub cmdprocesar_plantilla_Click()
If Trim(Me.txtDescripcion_plantilla.Text) <> "" And Val(Me.txtProcentaje_detalle.Text) > 0 Then
    strCadena = "call put_linea_plantilla('" & Val(Me.txtidPlantilla.Text) & "','" & Trim(Me.lblid_linea.Caption) & "','" & Trim(Me.txtDescripcion_plantilla.Text) & "','" & Val(Me.txtProcentaje_detalle.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Me.frmplantilla_detalle.Visible = False
    Call Me.llenar_plantilla(Me.HfPlantilla)
End If
End Sub

Private Sub cmdProcesarLinea_Click()
If Me.chk_Habilitado.Value = 0 Then
       in_habilitado = "no"
       
    Else
       in_habilitado = "si"
       
    End If
    strCadena = "UPDATE parametros_produccion SET descripcion ='" & Trim(Me.txtDescripcionConfiguracion.Text) & "',habilitado='" & in_habilitado & "' WHERE id='" & Val(Me.HfConfiguracion.TextMatrix(Me.HfConfiguracion.Row, 0)) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    CnBd.Execute (strCadena)
    
    Me.FrmConfiguracionDetalle.Visible = False
    Call Me.llenar_configuracion(Me.HfConfiguracion)
End Sub

Private Sub CmdQuitar_Click()
Procedencia = Eliminar
FrmSeguridad.Show
End Sub

Private Sub cmdupdate_Click()
Me.txtidPlantilla.Text = Me.HfPlantilla.TextMatrix(Me.HfPlantilla.Row, 0)
Me.txtDescripcion_plantilla.Text = Trim(Me.HfPlantilla.TextMatrix(Me.HfPlantilla.Row, 1))
Me.txtProcentaje_detalle.Text = Val(Me.HfPlantilla.TextMatrix(Me.HfPlantilla.Row, 2))
Me.frmplantilla_detalle.Visible = True
End Sub

Private Sub cmdupdatear_Click()
strCadena = "UPDATE linea_mantenimiento SET dias='" & Val(Me.txtdias.Text) & "' WHERE id_mantenimiento='" & Val(Me.Hfmantenimientos.TextMatrix(Me.Hfmantenimientos.Row, 0)) & "'"
CnBd.Execute (strCadena)
 
Me.Hfmantenimientos.TextMatrix(Me.Hfmantenimientos.Row, 2) = Val(Me.txtdias.Text)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
  End If
End Sub
Public Sub llenar_fases_produccion(ByVal id_linea As String, ByVal Grilla As MSHFlexGrid)
Dim porcentaje As Single


'Me.HfPrecios.MergeCells = flexMergeFree
strCadena = "SELECT P.id_produccion,P.descripcion,D.estado,D.id_producto FROM linea_produccion P,linea_produccion_detalle D  WHERE   D.id_linea='" & id_linea & "' AND P.id_produccion=D.id_produccion AND P.ruc=D.ruc AND P.ruc='" & KEY_RUC & "'ORDER BY id_produccion ASC"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       
        For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2500
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 2800
        Next
        cabecera = "ID_PRODUCCION" & vbTab & "LINEA PRODUCCION" & vbTab & "ESTADO" & vbTab & "PRODUCCION"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
            c = 2
            NumeroCampo = 2
            
        For i = 0 To rstT.RecordCount - 1
          If rstT("estado") = "si" Then
              estado = Chr(254)
          Else
              estado = Chr(168)
          End If
          nnproducto = ""
          If Val(rstT("id_producto")) > 0 Then
                strCadena = "SELECT nombre_prod FROM producto where id_producto='" & rstT("id_producto") & "' and ruc='" & KEY_RUC & "'"
                Call ConfiguraRstZ(strCadena)
                If rstZ.RecordCount > 0 Then
                    nnproducto = rstZ("nombre_prod")
                Else
                    nnproducto = "-"
                End If
          End If
          
          
          Fila = rstT("id_produccion") & vbTab & rstT("descripcion") & vbTab & estado & vbTab & nnproducto
          Grilla.AddItem Fila
        
        If c = NumeroCampo Then
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = c '  .. en la columna
                            ' cambia la fuente para esta celda
                            
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            ' edita la celda
                            ' If rstT("estado") = "no" Then
                             '   estado = Chr(168)
                            'Else
                             '   estado = Chr(254)
                            'End If
                            
                        End With
        End If
        Fila = ""
          If rstT("estado") = "si" Then
            For j = 0 To 2
                Grilla.col = j
                Grilla.Row = i + 1
                Grilla.CellBackColor = &HC0FFC0
            Next j
        End If
          rstT.MoveNext
      Next i
    
End Sub
Public Sub llenar_insumos(ByVal in_mantenimiento As String, ByVal Grilla As MSHFlexGrid)
Dim porcentaje As Single
strCadena = "SELECT d.id_detalle,p.id_producto,p.nombre_prod,d.cantidad,u.abreviatura,d.pagado FROM linea_mantenimiento_detalle d,producto p,unidad u  WHERE d.id_producto=p.id_producto and p.id_unidad=u.id_und and p.ruc=u.id_usu AND d.id_mantenimiento='" & in_mantenimiento & "'  AND d.ruc=p.ruc and  d.ruc='" & KEY_RUC & "'ORDER BY id_detalle ASC"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       
        For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 800
           Grilla.ColWidth(2) = 2200
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1000
        Next
        cabecera = "IDDETALLE" & vbTab & "COD" & vbTab & "DESCRIPCION" & vbTab & " CANTIDAD" & vbTab & " PAGADO"
        Grilla.AddItem cabecera
         For k = 1 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
           rstT.MoveFirst
            c = 4
            NumeroCampo = 4
            
        For i = 0 To rstT.RecordCount - 1
          If rstT("pagado") = "si" Then
              estado = Chr(254)
          Else
              estado = Chr(168)
          End If
          nnproducto = ""
          
          Fila = rstT("id_detalle") & vbTab & rstT("id_producto") & vbTab & rstT("nombre_prod") & vbTab & rstT("cantidad") & vbTab & estado
          Grilla.AddItem Fila
        
        If c = NumeroCampo Then
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = c '  .. en la columna
                            ' cambia la fuente para esta celda
                            
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            ' edita la celda
                            ' If rstT("estado") = "no" Then
                             '   estado = Chr(168)
                            'Else
                             '   estado = Chr(254)
                            'End If
                            
                        End With
        End If
        Fila = ""
          If rstT("pagado") = "si" Then
            For j = 1 To 4
                Grilla.col = j
                Grilla.Row = i + 1
                Grilla.CellBackColor = &HC0FFC0
            Next j
        End If
          rstT.MoveNext
      Next i
    
End Sub
Public Sub llenar_mantenimientos(ByVal id_linea As String, ByVal Grilla As MSHFlexGrid)
Dim porcentaje As Single


'Me.HfPrecios.MergeCells = flexMergeFree
strCadena = "SELECT * FROM linea_mantenimiento  WHERE   id_linea='" & id_linea & "'  AND ruc='" & KEY_RUC & "'ORDER BY id_mantenimiento ASC"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       
        For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 800
           Grilla.ColWidth(2) = 2200
           Grilla.ColWidth(3) = 1000
        Next
        cabecera = "IDDETALLE" & vbTab & "ITEM" & vbTab & "Nº MESES" & vbTab & " INSUMOS"
        Grilla.AddItem cabecera
         For k = 1 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
           rstT.MoveFirst
            c = 3
            NumeroCampo = 3
            
        For i = 0 To rstT.RecordCount - 1
          If rstT("insumos") = "si" Then
              estado = Chr(254)
          Else
              estado = Chr(168)
          End If
          nnproducto = ""
          
          Fila = rstT("id_mantenimiento") & vbTab & Format(i + 1, "00") & vbTab & rstT("dias") & vbTab & estado
          Grilla.AddItem Fila
        
        If c = NumeroCampo Then
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = c '  .. en la columna
                            ' cambia la fuente para esta celda
                            
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            ' edita la celda
                            ' If rstT("estado") = "no" Then
                             '   estado = Chr(168)
                            'Else
                             '   estado = Chr(254)
                            'End If
                            
                        End With
        End If
        Fila = ""
          If rstT("insumos") = "si" Then
            For j = 0 To 3
                Grilla.col = j
                Grilla.Row = i + 1
                Grilla.CellBackColor = &HC0FFC0
            Next j
        End If
          rstT.MoveNext
      Next i
    
End Sub
Public Sub llenar_plantilla(ByVal Grilla As MSHFlexGrid)
Dim porcentaje As Single
strCadena = "SELECT * FROM linea_plantilla WHERE id_linea='" & Trim(Me.lblid_linea.Caption) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
      Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       
        For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 3800
           Grilla.ColWidth(2) = 700
           
        Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION ITEM" & vbTab & " [ % ]"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
           rstT.MoveFirst
          
            
        For i = 0 To rstT.RecordCount - 1
         
          
          Fila = Format(rstT("id"), "00") & vbTab & rstT("descripcion") & vbTab & Format(rstT("porcentaje"), "#,##0.00")
          Grilla.AddItem Fila
        porcentaje = porcentaje + rstT("porcentaje")
        
        
       
          rstT.MoveNext
      Next i
       Fila = "" & vbTab & "TOTAL PORCENTAJE " & vbTab & Format(porcentaje, "#,##0.00")
       Grilla.AddItem Fila
       For k = 1 To 2
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &H80FF&
         Next k
End Sub
Private Sub put_configuraion()

strCadena = "SELECT * FROM parametros_produccion WHERE ruc='0' ORDER BY id ASC"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   rstK.MoveFirst
   For i = 0 To rstK.RecordCount - 1
        
        strCadena = "INSERT INTO parametros_produccion(`codigo`,`descripcion`,`habilitado`,`referencia`,`ruc`)VALUES " & _
        "('" & rstK("codigo") & "','" & rstK("descripcion") & "','no','" & rstK("referencia") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        rstK.MoveNext
   Next i
End If
Call llenar_configuracion(Me.HfConfiguracion)

End Sub

Public Sub llenar_configuracion(ByVal Grilla As MSHFlexGrid)
Dim porcentaje As Single

strCadena = "SELECT * FROM parametros_produccion WHERE   ruc='" & KEY_RUC & "' ORDER BY id asc"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    
    Grilla.Rows = 0
    Call put_configuraion
    
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       
        For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2500
           Grilla.ColWidth(2) = 1500
           Grilla.ColWidth(3) = 1000
        Next
        cabecera = "IDDETALLE" & vbTab & "DESCRIPCION" & vbTab & "REFERENCIA" & vbTab & " HABILITADO"
        Grilla.AddItem cabecera
         For k = 1 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
           rstT.MoveFirst
            c = 3
            NumeroCampo = 3
            
        For i = 0 To rstT.RecordCount - 1
          If rstT("habilitado") = "si" Then
              habilitado = Chr(254)
          Else
              habilitado = Chr(168)
          End If
          nnproducto = ""
          
          Fila = rstT("id") & vbTab & rstT("descripcion") & vbTab & rstT("referencia") & vbTab & habilitado
          Grilla.AddItem Fila
        
        If c = NumeroCampo Then
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = c '  .. en la columna
                            ' cambia la fuente para esta celda
                            
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            ' edita la celda
                            ' If rstT("estado") = "no" Then
                             '   estado = Chr(168)
                            'Else
                             '   estado = Chr(254)
                            'End If
                            
                        End With
        End If
        
        
          rstT.MoveNext
      Next i
    
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500

strCadena = "SELECT id_linea as Codigo,descripcion as Descripcion FROM linea_online ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcClasificacion)






  Select Case FrmLinea.Procedencia
    Case modificar
      Call LLENA
  End Select
End Sub

Private Sub LLENA()
  StrCodTabla = FrmLinea.HfgLinea.TextMatrix(FrmLinea.HfgLinea.Row, 0)
  strCadena = "SELECT * FROM linea where id_linea ='" & StrCodTabla & "' AND id_usu='" & KEY_RUC & "'"
  Call ConfiguraRstP(strCadena)
  If rstP.RecordCount > 0 Then
        Me.lblid_linea.Caption = StrCodTabla
        TxtDescripcion.Text = FrmLinea.HfgLinea.TextMatrix(FrmLinea.HfgLinea.Row, 1)
        Me.txtnumero_mantenimientos.Text = rstP("mantenimientos")
        Me.txtkilometros.Text = rstP("mantenimientos")
        Me.DtcClasificacion.BoundText = rstP("id_clasificacion")
        If rstP("produccion") = "si" Then
            Me.txtperiodogarantia.Text = rstP("garantia")
            Me.chkproduccion.Value = 1
            Me.HfItem.Visible = True
            Me.frmgarantia.Visible = True
            If rstP("motor") = "si" Then
                Me.chk_motor.Value = 1
            Else
                Me.chk_motor.Value = 0
            End If
            
        Else
            Me.chkproduccion.Value = 0
            Me.HfItem.Visible = False
             Me.frmgarantia.Visible = True
        End If
        
        If rstP("activo") = "si" Then
           Me.chk_activado_online.Value = 1
           Me.cmdImagen.Visible = True
           Call get_image(StrCodTabla)
        Else
          Me.chk_activado_online.Value = 0
          Me.img_foto.Visible = False
        End If
        
        
        
        
        
        If rstP("afecto_garantia") = "si" Then
            Me.chkgarantia.Value = 1
            Me.frmgarantia.Visible = True
            Me.txtnumero_mantenimientos.Text = rstP("mantenimientos")
            Call Me.llenar_mantenimientos(FrmLinea.HfgLinea.TextMatrix(FrmLinea.HfgLinea.Row, 0), Me.Hfmantenimientos)
        Else
            Me.chkgarantia.Value = 0
            Me.frmgarantia.Visible = False
        End If
        Me.txtcodigo_contable.Text = rstP("nro_cuenta")
        Me.txtCuentaContableImportacion.Text = rstP("nro_cuenta_importacion")
        Me.txtCostoDebe.Text = rstP("cuenta_costodebe")
        Me.txtCostoHaber.Text = rstP("cuenta_costohaber")
        Me.lblcuentadebecosto.Caption = UCase(get_cuenta(Trim(Me.txtCostoDebe.Text)))
        Me.lblcuentahabercosto.Caption = UCase(get_cuenta(Trim(Me.txtCostoHaber.Text)))
        
        
        Me.lblcuenta_contable.Caption = UCase(get_cuenta(Trim(Me.txtcodigo_contable.Text)))
        Me.lblcuenta_contable_import.Caption = UCase(get_cuenta(Trim(Me.txtCuentaContableImportacion.Text)))
  End If
  
End Sub
Private Sub Save()
Dim produccion As String
Dim strgarantia As String

  If TxtDescripcion.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
  Else
  
  If Me.chkproduccion.Value = 1 Then
    produccion = "si"
  Else
    produccion = "no"
  End If
  
  If Me.chkgarantia.Value = 1 Then
    strgarantia = "si"
  Else
    strgarantia = "no"
  End If
  
  If Me.chk_motor.Value = 1 Then
     in_motor = "si"
  Else
     in_motor = "no"
  End If
  
  If Me.chk_activado_online.Value = 1 Then
    in_habilitado = "si"
  Else
    in_habilitado = "no"
  End If
  
  
  
    Select Case FrmLinea.Procedencia
      Case nuevo
        strCadena = "SELECT * FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY id_linea DESC"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            strCodLinea = formato_item(Val(rst("id_linea")) + 1, 5)
        Else
            strCodLinea = formato_item(1, 5)
        End If
        
        strCadena = "INSERT INTO linea (id_linea,id_clasificacion,descripcion,produccion,motor,garantia,afecto_garantia,mantenimientos,nro_cuenta,nro_cuenta_importacion,cuenta_costodebe,cuenta_costohaber,activo,imagen,imagen_png,id_usu) VALUES " & _
        " ('" & strCodLinea & "','" & Me.DtcClasificacion.BoundText & "','" & TxtDescripcion.Text & "','" & produccion & "','" & in_motor & "','" & Val(Me.txtperiodogarantia.Text) & "','" & strgarantia & "','" & Val(Me.txtnumero_mantenimientos.Text) & "','" & Trim(Me.txtcodigo_contable.Text) & "','" & Trim(Me.txtCuentaContableImportacion.Text) & "','" & Trim(Me.txtCostoDebe.Text) & "','" & Trim(Me.txtCostoHaber.Text) & "','" & in_habilitado & "','" & Trim(Me.cmdImagen.Tag) & "','" & Trim(Me.cmdImagen.ToolTipText) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
         
        Call FrmLinea.actualizar
        FrmLinea.Procedencia = Neutro
        Unload Me
      Case modificar
        
        strCadena = "UPDATE linea SET activo='" & in_habilitado & "',id_clasificacion='" & Me.DtcClasificacion.BoundText & "',cuenta_costodebe='" & Trim(Me.txtCostoDebe.Text) & "',cuenta_costohaber='" & Trim(Me.txtCostoHaber.Text) & "',motor='" & in_motor & "', nro_cuenta_importacion='" & Trim(Me.txtCuentaContableImportacion.Text) & "',nro_cuenta='" & Trim(Me.txtcodigo_contable.Text) & "', descripcion='" & TxtDescripcion.Text & "',mantenimientos='" & Val(Me.txtnumero_mantenimientos.Text) & "',afecto_garantia='" & strgarantia & "',produccion='" & produccion & "',garantia='" & Val(Me.txtperiodogarantia.Text) & "',imagen='" & Trim(Me.cmdImagen.Tag) & "',imagen_png='" & Trim(Me.cmdImagen.ToolTipText) & "' WHERE id_linea = '" & StrCodTabla & "' AND id_usu='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
        FrmLinea.HfgLinea.TextMatrix(FrmLinea.HfgLinea.Row, 1) = Trim(Me.TxtDescripcion.Text)
        FrmLinea.Procedencia = Neutro
        Unload Me
    End Select
   
  End If
End Sub

Private Sub HfConfiguracion_DblClick()
If Val(Me.HfConfiguracion.TextMatrix(Me.HfConfiguracion.Row, 0)) > 0 Then
    If Me.HfConfiguracion.TextMatrix(Me.HfConfiguracion.Row, 3) = Chr(254) Then
       Me.chk_Habilitado.Value = 1
    Else
       Me.chk_Habilitado.Value = 0
    End If
    
    Me.txtDescripcionConfiguracion.Text = Me.HfConfiguracion.TextMatrix(Me.HfConfiguracion.Row, 1)
    
    Me.FrmConfiguracionDetalle.Visible = True
    
End If
End Sub

Private Sub HfItem_Click()
If Me.HfItem.Rows > 0 Then
    If Val(Me.HfItem.TextMatrix(Me.HfItem.Row, 0)) > 0 Then
        Call ActualizarImagen(Me.HfItem.TextMatrix(Me.HfItem.Row, 0), FrmLinea.HfgLinea.TextMatrix(FrmLinea.HfgLinea.Row, 0))
    End If
End If
End Sub

Private Sub Hfmantenimientos_DblClick()
If Val(Me.Hfmantenimientos.TextMatrix(Me.Hfmantenimientos.Row, 0)) > 0 Then
    Me.frminsumos.Visible = True
    Me.txtdias.Text = Val(Me.Hfmantenimientos.TextMatrix(Me.Hfmantenimientos.Row, 2))
    Call Me.llenar_insumos(Val(Me.Hfmantenimientos.TextMatrix(Me.Hfmantenimientos.Row, 0)), Me.HfInsumos)
End If
End Sub

Private Sub Image1_Click()
Me.FrmConfiguracionDetalle.Visible = False
End Sub

Private Sub Image2_Click()
Me.frmConfiguracion.Visible = False
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


Private Sub txtcodigo_contable_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Procedencia = Selecionar
   FrmPlanContableCuentas.Show
   Exit Sub
End If
End Sub

Private Sub txtCostoDebe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = seleccionar_soldadura
    FrmPlanContableCuentas.Show
   Exit Sub
End If
End Sub

Private Sub txtCuentaContableImportacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = seleccionar_otro
    FrmPlanContableCuentas.Show
    Exit Sub
End If
End Sub

Private Sub txtdescripcion_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub

Private Sub txtid_producto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
End If
End Sub

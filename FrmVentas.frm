VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmVentas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Seccion Ventas"
   ClientHeight    =   9240
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   20145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtExtranjero 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   230
      Text            =   "no"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chk_venta_diferida 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "VENTA DIFERIDA"
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
      Left            =   12120
      TabIndex        =   221
      Top             =   710
      Width           =   2055
   End
   Begin VB.TextBox txtServicio 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   202
      Text            =   "no"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrarpantalla 
      Height          =   1020
      Left            =   14520
      TabIndex        =   187
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1799
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
      MICON           =   "FrmVentas.frx":0000
      PICN            =   "FrmVentas.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdActivar 
      Height          =   405
      Left            =   4920
      TabIndex        =   183
      Top             =   195
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   714
      BTYPE           =   5
      TX              =   "ACTIVAR FACTURACION"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   33023
      BCOLO           =   33023
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmVentas.frx":3043
      PICN            =   "FrmVentas.frx":305F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frmprincipal 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   11175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20175
      Begin VB.TextBox txtid_agenda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   247
         Top             =   6120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   15120
         ScaleHeight     =   315
         ScaleWidth      =   195
         TabIndex        =   246
         Top             =   6240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chk_OrdenCompra 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "ORDEN COMPRA"
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
         Height          =   300
         Left            =   15600
         TabIndex        =   245
         Top             =   2820
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox txtOrdenCompra 
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
         Height          =   285
         Left            =   17160
         TabIndex        =   244
         Text            =   "4700004202"
         Top             =   2820
         Width           =   1935
      End
      Begin VB.CheckBox chk_detraccion 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "AFECTO DETRACCION"
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
         Left            =   12120
         TabIndex        =   243
         Top             =   940
         Width           =   2055
      End
      Begin VB.TextBox txtmail 
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
         Left            =   3480
         MaxLength       =   80
         TabIndex        =   242
         ToolTipText     =   "DNI / RUC"
         Top             =   2500
         Width           =   3375
      End
      Begin VB.TextBox txtTelefono 
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
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   241
         ToolTipText     =   "DNI / RUC"
         Top             =   2500
         Width           =   1280
      End
      Begin VB.Frame frmvencimiento 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   720
         Left            =   11400
         TabIndex        =   217
         Top             =   1210
         Visible         =   0   'False
         Width           =   1960
         Begin MSMask.MaskEdBox txtFecha_vencimiento 
            Height          =   285
            Left            =   60
            TabIndex        =   218
            ToolTipText     =   "dd/mm/yyyy"
            Top             =   240
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
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
         Begin VitekeySoft.ChameleonBtn cmdgenerarLetras 
            Height          =   450
            Left            =   1260
            TabIndex        =   220
            Top             =   120
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   794
            BTYPE           =   2
            TX              =   "GENERAR LETRAS"
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
            MICON           =   "FrmVentas.frx":558A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VENCIMIENTO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   45
            TabIndex        =   219
            Top             =   0
            Width           =   1005
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   238
         Text            =   "no"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPrecios 
         Height          =   1155
         Left            =   7695
         TabIndex        =   121
         Top             =   7680
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   2037
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
      Begin VB.Frame frmalmacen_entrega 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "ENTREGA EN ALMACEN"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Index           =   1
         Left            =   11400
         TabIndex        =   223
         Top             =   2295
         Visible         =   0   'False
         Width           =   3015
         Begin VB.TextBox txtNumero_nota 
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
            Left            =   1320
            TabIndex        =   226
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtserie_nota 
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
            Left            =   360
            TabIndex        =   225
            Top             =   480
            Width           =   855
         End
         Begin VitekeySoft.ChameleonBtn cmdCerrarEntrega 
            Height          =   180
            Index           =   1
            Left            =   2760
            TabIndex        =   224
            Top             =   120
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   318
            BTYPE           =   5
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmVentas.frx":55A6
            PICN            =   "FrmVentas.frx":55C2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SALDO :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   4
            Left            =   480
            TabIndex        =   229
            Top             =   960
            Width           =   660
         End
         Begin VB.Label Label22 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   300
            Index           =   3
            Left            =   1320
            TabIndex        =   228
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "INGRESE SU NOTA DE CREDITO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   227
            Top             =   120
            Width           =   2445
         End
      End
      Begin VB.Frame FrameSerieModelo 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   2685
         Left            =   2640
         TabIndex        =   132
         Top             =   3990
         Visible         =   0   'False
         Width           =   8655
         Begin VB.TextBox txtid_temporal_serie 
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
            Left            =   6360
            MaxLength       =   80
            TabIndex        =   231
            ToolTipText     =   "TELEFONO"
            Top             =   1680
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtModelo 
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
            Left            =   1920
            MaxLength       =   80
            TabIndex        =   141
            ToolTipText     =   "DNI / RUC"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtA�oFabricacion 
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
            Left            =   4515
            MaxLength       =   80
            TabIndex        =   140
            ToolTipText     =   "DNI / RUC"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtbusquedamotor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   4515
            MaxLength       =   80
            TabIndex        =   138
            ToolTipText     =   "DNI / RUC"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox TxtColor 
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
            Left            =   1920
            MaxLength       =   80
            TabIndex        =   137
            ToolTipText     =   "DNI / RUC"
            Top             =   2085
            Width           =   1455
         End
         Begin VB.TextBox txtBuscarSerie 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   4515
            MaxLength       =   80
            TabIndex        =   136
            ToolTipText     =   "DNI / RUC"
            Top             =   285
            Width           =   1575
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
            Left            =   1920
            MaxLength       =   80
            TabIndex        =   135
            ToolTipText     =   "DNI / RUC"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtdua 
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
            Left            =   4515
            MaxLength       =   80
            TabIndex        =   134
            ToolTipText     =   "DNI / RUC"
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtitem 
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
            Left            =   4515
            MaxLength       =   80
            TabIndex        =   133
            ToolTipText     =   "DNI / RUC"
            Top             =   2160
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo DtcSerie 
            Height          =   330
            Left            =   1920
            TabIndex        =   139
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   ""
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
         Begin VitekeySoft.ChameleonBtn cmdcerrar 
            Height          =   375
            Left            =   6360
            TabIndex        =   142
            Top             =   2040
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "           CERRAR "
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
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
            MICON           =   "FrmVentas.frx":8476
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo DtcMotor 
            Height          =   330
            Left            =   1920
            TabIndex        =   143
            Top             =   720
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   8388608
            Text            =   ""
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
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MODELO :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1065
            TabIndex        =   151
            Top             =   1680
            Width           =   705
         End
         Begin VB.Label lblmotor 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MOTOR :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1140
            TabIndex        =   150
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblchasis 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CHASIS :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1155
            TabIndex        =   149
            Top             =   360
            Width           =   585
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A�O MOD:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   3600
            TabIndex        =   148
            Top             =   1320
            Width           =   765
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COLOR  :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1155
            TabIndex        =   147
            Top             =   2085
            Width           =   585
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MARCA :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1125
            TabIndex        =   146
            Top             =   1200
            Width           =   645
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NRO DUA :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   3690
            TabIndex        =   145
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NRO ITEM :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   3630
            TabIndex        =   144
            Top             =   2160
            Width           =   765
         End
      End
      Begin VB.Frame FrameReferencia 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1695
         Left            =   10320
         TabIndex        =   32
         Top             =   1185
         Visible         =   0   'False
         Width           =   3975
         Begin VB.TextBox txtdocreferencia 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   930
            Left            =   120
            MaxLength       =   80
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   645
            Width           =   3735
         End
         Begin VB.TextBox txtid_venta_ref 
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
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   120
            MaxLength       =   80
            TabIndex        =   33
            Top             =   645
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COMPROBANTE REFERENCIA"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   720
            TabIndex        =   35
            Top             =   180
            Width           =   2310
         End
      End
      Begin VB.Frame frmcredito 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   7080
         TabIndex        =   55
         Top             =   1305
         Visible         =   0   'False
         Width           =   3255
         Begin VB.TextBox txtmontocredito 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   56
            Top             =   120
            Width           =   1335
         End
         Begin VitekeySoft.ChameleonBtn cmdactualizarcredito 
            Height          =   285
            Left            =   1680
            TabIndex        =   57
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            BTYPE           =   3
            TX              =   "ACTUALIZAR"
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
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmVentas.frx":8492
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MONTO CREDITO :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   120
            TabIndex        =   58
            Top             =   120
            Width           =   1395
         End
      End
      Begin VB.TextBox txtagranel 
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
         Left            =   14520
         MaxLength       =   80
         TabIndex        =   222
         ToolTipText     =   "TELEFONO"
         Top             =   6720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame frmbanco 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2745
         Left            =   8460
         TabIndex        =   45
         Top             =   2360
         Visible         =   0   'False
         Width           =   4260
         Begin VB.TextBox txtbuscarbanco 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   48
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox txtCheque 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   47
            Top             =   2400
            Width           =   2415
         End
         Begin VB.TextBox txtBanco 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   46
            Top             =   2040
            Width           =   2415
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfBancos 
            Height          =   1575
            Left            =   45
            TabIndex        =   49
            Top             =   45
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   2778
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
         Begin VitekeySoft.ChameleonBtn cmdcerrarbanco 
            Height          =   180
            Left            =   3960
            TabIndex        =   50
            Top             =   120
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   318
            BTYPE           =   5
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmVentas.frx":84AE
            PICN            =   "FrmVentas.frx":84CA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N� CHEQUE:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   0
            TabIndex        =   52
            Top             =   2400
            Width           =   915
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BANCO :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   285
            TabIndex        =   51
            Top             =   2040
            Width           =   630
         End
      End
      Begin VB.Frame frmdireccion 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2175
         Left            =   7080
         TabIndex        =   11
         Top             =   2400
         Visible         =   0   'False
         Width           =   8295
         Begin VB.CommandButton cmdcerrardireccion 
            Height          =   255
            Left            =   8000
            Picture         =   "FrmVentas.frx":B37E
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   255
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfdireccion 
            Height          =   1935
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   7815
            _ExtentX        =   13785
            _ExtentY        =   3413
            _Version        =   393216
            ForeColor       =   8388608
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   8388608
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
      End
      Begin VB.Frame fraApp 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   8325
         TabIndex        =   15
         Top             =   2355
         Visible         =   0   'False
         Width           =   5655
         Begin VB.CommandButton cmdvincular 
            Caption         =   "VINCULAR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3960
            TabIndex        =   19
            Top             =   480
            Width           =   1575
         End
         Begin VB.CommandButton cmddesvincular 
            Caption         =   "DESVINCULAR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3960
            TabIndex        =   18
            Top             =   960
            Width           =   1575
         End
         Begin VB.CommandButton cmdsalir 
            Caption         =   "SALIR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3960
            TabIndex        =   17
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox txtrecibo_anterior 
            Height          =   285
            Left            =   4200
            TabIndex        =   16
            Text            =   "0"
            Top             =   1440
            Visible         =   0   'False
            Width           =   975
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfRecibos 
            Height          =   1935
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   3413
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
      End
      Begin VB.Frame PanelCredito 
         BackColor       =   &H00FFFFFF&
         Height          =   750
         Left            =   11400
         TabIndex        =   81
         Top             =   1200
         Visible         =   0   'False
         Width           =   1960
         Begin VB.TextBox TxtClaveRandonCredito 
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
            Left            =   840
            MaxLength       =   80
            TabIndex        =   83
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox TxtCuotas 
            Alignment       =   1  'Right Justify
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
            Left            =   840
            MaxLength       =   80
            TabIndex        =   82
            Top             =   420
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblclave 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CLAVE  :"
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
            TabIndex        =   85
            Top             =   120
            Width           =   555
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CUOTAS :"
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
            TabIndex        =   84
            Top             =   480
            Width           =   645
         End
      End
      Begin VB.TextBox TxtDescuento_global 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   210
         Top             =   8865
         Width           =   1000
      End
      Begin VB.TextBox TxtDescuento_porcentaje 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   209
         Top             =   8865
         Width           =   600
      End
      Begin VB.CheckBox chk_descuento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "DESCUENTO GLOBAL"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   260
         Left            =   4000
         Picture         =   "FrmVentas.frx":E222
         TabIndex        =   208
         ToolTipText     =   "OBSEQUIO"
         Top             =   8880
         Width           =   1935
      End
      Begin VB.CheckBox chkObsequio 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Height          =   300
         Left            =   10620
         Picture         =   "FrmVentas.frx":E7AC
         TabIndex        =   203
         ToolTipText     =   "OBSEQUIO"
         Top             =   6880
         Width           =   375
      End
      Begin VB.Frame lblAnulado 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   5520
         TabIndex        =   37
         Top             =   5385
         Visible         =   0   'False
         Width           =   3975
         Begin VB.Label lblAnulado1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ANULADO"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   36
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   870
            Left            =   240
            TabIndex        =   38
            Top             =   -120
            Width           =   3570
         End
      End
      Begin VB.CommandButton cmdCuotas 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13380
         Picture         =   "FrmVentas.frx":ED36
         Style           =   1  'Graphical
         TabIndex        =   200
         ToolTipText     =   "LISTADO CUOTAS"
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame frmcopropietario 
         BackColor       =   &H00404040&
         Caption         =   "DATOS COPROPIETARIO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   855
         Left            =   6960
         TabIndex        =   39
         Top             =   1665
         Visible         =   0   'False
         Width           =   5895
         Begin VB.TextBox txtdni_copropietario 
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
            Left            =   600
            MaxLength       =   80
            TabIndex        =   41
            ToolTipText     =   "DNI / RUC"
            Top             =   360
            Width           =   1335
         End
         Begin VB.CommandButton cmdokcopropietario 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5280
            TabIndex        =   40
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DNI  :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   210
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   420
         End
         Begin VB.Label lblcopropietario 
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
            Left            =   1965
            TabIndex        =   42
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame frmimpresoras 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2655
         Left            =   15600
         TabIndex        =   2
         Top             =   -1650
         Visible         =   0   'False
         Width           =   4420
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfImpresoras 
            Height          =   2055
            Left            =   45
            TabIndex        =   3
            Top             =   120
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   3625
            _Version        =   393216
            ForeColor       =   8388608
            Rows            =   0
            FixedRows       =   0
            FixedCols       =   0
            ForeColorFixed  =   8388608
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
         Begin VitekeySoft.ChameleonBtn cmdcerrarimpresora 
            Height          =   180
            Left            =   4080
            TabIndex        =   4
            Top             =   120
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   318
            BTYPE           =   5
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmVentas.frx":1107A
            PICN            =   "FrmVentas.frx":11096
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdenviarimpresion 
            Height          =   345
            Left            =   45
            TabIndex        =   5
            Top             =   2235
            Visible         =   0   'False
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   609
            BTYPE           =   3
            TX              =   "ENVIAR A IMPRESION"
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   33023
            BCOLO           =   33023
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmVentas.frx":13F4A
            PICN            =   "FrmVentas.frx":13F66
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   1
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame frm_motivo_nota 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   15600
         TabIndex        =   25
         Top             =   -15
         Visible         =   0   'False
         Width           =   4455
         Begin VB.TextBox txtmotivo_nota 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   645
            Left            =   840
            MaxLength       =   80
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            ToolTipText     =   "INGRESE UN MOTIVO"
            Top             =   600
            Width           =   3345
         End
         Begin VB.CommandButton cmdcerrarmotivonota 
            Height          =   255
            Left            =   4200
            Picture         =   "FrmVentas.frx":13FF3
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   40
            Width           =   255
         End
         Begin MSDataListLib.DataCombo DtcTipoNota 
            Height          =   330
            Left            =   840
            TabIndex        =   28
            Top             =   165
            Width           =   3345
            _ExtentX        =   5900
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
            Text            =   ""
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
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MOTIVO :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   2
            Left            =   0
            TabIndex        =   30
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "T.NOTA:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame frminteres 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   13365
         TabIndex        =   196
         Top             =   1110
         Visible         =   0   'False
         Width           =   1140
         Begin VB.TextBox txtporcentaje_interes 
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
            Left            =   60
            TabIndex        =   198
            Top             =   380
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox chk_interes 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "%INTERES"
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
            Height          =   230
            Left            =   60
            TabIndex        =   197
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.CheckBox chkseguro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "SEGURO "
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
         Height          =   250
         Left            =   2160
         TabIndex        =   194
         Top             =   3165
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame frmmanual 
         BackColor       =   &H008080FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   14040
         TabIndex        =   190
         Top             =   8160
         Width           =   1455
         Begin VB.CommandButton Command1 
            Caption         =   "ELIMINAR"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   220
            Left            =   120
            TabIndex        =   192
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chk_manual 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "MANUALES"
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
            Height          =   195
            Left            =   120
            TabIndex        =   191
            Top             =   80
            Width           =   1215
         End
      End
      Begin VB.CheckBox chkimprimira4 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "FORMATO A4"
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
         Left            =   14040
         TabIndex        =   188
         Top             =   8820
         Width           =   1455
      End
      Begin VitekeySoft.ChameleonBtn cmdEliminar 
         Height          =   945
         Left            =   14520
         TabIndex        =   186
         Top             =   1965
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1667
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
         MICON           =   "FrmVentas.frx":16E97
         PICN            =   "FrmVentas.frx":16EB3
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
         Height          =   950
         Left            =   14520
         TabIndex        =   184
         Top             =   45
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1667
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
         MICON           =   "FrmVentas.frx":192FD
         PICN            =   "FrmVentas.frx":19319
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CommandButton CmdVisualizar 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "Command1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   7800
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox TxtCliente 
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
         MaxLength       =   300
         TabIndex        =   153
         Top             =   1780
         Width           =   4695
      End
      Begin VB.Frame frameCajaIndependiente 
         BackColor       =   &H00FFFFFF&
         Height          =   2300
         Left            =   15600
         TabIndex        =   125
         Top             =   3600
         Width           =   4455
         Begin VB.Timer timer_pendientes 
            Enabled         =   0   'False
            Interval        =   10000
            Left            =   10
            Top             =   10
         End
         Begin VB.TextBox txtnumeropendientes 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3120
            TabIndex        =   127
            Top             =   1080
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txt_id_pendiente 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   13440
            TabIndex        =   126
            Top             =   8160
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPendientes 
            Height          =   1695
            Left            =   120
            TabIndex        =   128
            Top             =   150
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   2990
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
         Begin VitekeySoft.ChameleonBtn cmddescartar 
            Height          =   360
            Left            =   120
            TabIndex        =   129
            Top             =   1880
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   635
            BTYPE           =   5
            TX              =   "DESC.  ATENCION"
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
            BCOL            =   8421631
            BCOLO           =   8421631
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmVentas.frx":1976B
            PICN            =   "FrmVentas.frx":19787
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdlistado 
            Height          =   360
            Left            =   2280
            TabIndex        =   130
            Top             =   1880
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   635
            BTYPE           =   5
            TX              =   "LISTADO"
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
            BCOL            =   8421631
            BCOLO           =   8421631
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmVentas.frx":19D21
            PICN            =   "FrmVentas.frx":19D3D
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
      Begin VB.CheckBox chkPrecios 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "PRECIOS  MAYOR"
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
         Height          =   270
         Left            =   7680
         TabIndex        =   122
         Top             =   7360
         Width           =   1450
      End
      Begin VB.CheckBox ChkExtraer 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "EXTRAER"
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
         Height          =   250
         Left            =   240
         TabIndex        =   106
         Top             =   950
         Width           =   1215
      End
      Begin VB.TextBox TxtSeri_guia 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4200
         MaxLength       =   80
         TabIndex        =   105
         Top             =   875
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TxtNumero_guia 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5325
         MaxLength       =   80
         TabIndex        =   104
         Top             =   875
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtDireccion 
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
         Height          =   285
         Left            =   2160
         MaxLength       =   300
         TabIndex        =   103
         Top             =   2150
         Width           =   4695
      End
      Begin VB.TextBox TxtDescripcionProducto 
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
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   1000
         TabIndex        =   102
         Top             =   6905
         Width           =   5055
      End
      Begin VB.TextBox TxtCodCliente 
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
         Height          =   285
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   101
         ToolTipText     =   "DNI / RUC"
         Top             =   1430
         Width           =   1335
      End
      Begin VB.TextBox TxtNumeroDoc 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   8610
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   100
         Top             =   590
         Width           =   1650
      End
      Begin VB.CheckBox chk_factura 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "STOCK FACT"
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
         Left            =   12120
         TabIndex        =   99
         Top             =   475
         Width           =   2055
      End
      Begin VB.TextBox TxtMontoPagado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   11400
         MaxLength       =   80
         TabIndex        =   98
         Top             =   1950
         Width           =   1935
      End
      Begin VB.TextBox TxtNumeroTargeta 
         Alignment       =   1  'Right Justify
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
         Left            =   11400
         MaxLength       =   80
         TabIndex        =   97
         Top             =   1625
         Visible         =   0   'False
         Width           =   950
      End
      Begin VB.CheckBox chkDelivery 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "DELIVERY"
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
         Left            =   960
         TabIndex        =   95
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox TxtPuntos 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   94
         Text            =   "0"
         Top             =   4910
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuitarMonto 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13380
         Picture         =   "FrmVentas.frx":1C091
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   2330
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton OptAuto 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "AUTO"
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
         Left            =   14080
         TabIndex        =   92
         Top             =   7550
         Width           =   1365
      End
      Begin VB.OptionButton OptManual 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "MANUAL"
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
         Height          =   210
         Left            =   14080
         TabIndex        =   91
         Top             =   7790
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.CheckBox chkconsultar 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "CONSULTAR"
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
         Height          =   315
         Left            =   10320
         TabIndex        =   90
         Top             =   590
         Width           =   1215
      End
      Begin VB.TextBox txtpeso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1560
         MaxLength       =   80
         TabIndex        =   88
         Text            =   "0.00"
         Top             =   105
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtMontoPagovitekey 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   11400
         MaxLength       =   80
         TabIndex        =   87
         Top             =   1950
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtOperacion 
         Alignment       =   1  'Right Justify
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
         Left            =   12405
         MaxLength       =   80
         TabIndex        =   86
         Top             =   1625
         Visible         =   0   'False
         Width           =   950
      End
      Begin VB.TextBox TxtIgv 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   80
         Text            =   "si"
         Top             =   345
         Visible         =   0   'False
         Width           =   495
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
         Left            =   14520
         MaxLength       =   80
         TabIndex        =   79
         ToolTipText     =   "TELEFONO"
         Top             =   5040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TxtTipoCambio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Height          =   315
         Left            =   11445
         MaxLength       =   80
         TabIndex        =   78
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox TxtIdVenta 
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
         Left            =   14520
         MaxLength       =   80
         TabIndex        =   76
         ToolTipText     =   "TELEFONO"
         Top             =   5385
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame frameTramite 
         BackColor       =   &H00FFFFFF&
         Height          =   1755
         Left            =   15600
         TabIndex        =   70
         Top             =   300
         Visible         =   0   'False
         Width           =   4455
         Begin VitekeySoft.ChameleonBtn cmdConstancia 
            Height          =   330
            Left            =   360
            TabIndex        =   71
            Top             =   165
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            BTYPE           =   5
            TX              =   "  CONSTANCIA         "
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
            MICON           =   "FrmVentas.frx":1C61B
            PICN            =   "FrmVentas.frx":1C637
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdDeclaracion 
            Height          =   330
            Left            =   360
            TabIndex        =   72
            Top             =   525
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   582
            BTYPE           =   5
            TX              =   " DECLARACION JURADA                                                   "
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
            MICON           =   "FrmVentas.frx":1E98B
            PICN            =   "FrmVentas.frx":1E9A7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdSolicitud 
            Height          =   345
            Left            =   360
            TabIndex        =   73
            Top             =   885
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   609
            BTYPE           =   5
            TX              =   "SOLICITUD AP        "
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
            MICON           =   "FrmVentas.frx":20CFB
            PICN            =   "FrmVentas.frx":20D17
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdgenerarmantenimientos 
            Height          =   345
            Left            =   2280
            TabIndex        =   74
            Top             =   885
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   609
            BTYPE           =   5
            TX              =   "CARTA PODER             "
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
            MICON           =   "FrmVentas.frx":2306B
            PICN            =   "FrmVentas.frx":23087
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdpedidosoendientes 
            Height          =   330
            Left            =   2280
            TabIndex        =   75
            Top             =   165
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            BTYPE           =   5
            TX              =   "PEDIDOS PENDIEN  "
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
            MICON           =   "FrmVentas.frx":253DB
            PICN            =   "FrmVentas.frx":253F7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdSolicitudcredito 
            Height          =   345
            Left            =   360
            TabIndex        =   214
            Top             =   1320
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   609
            BTYPE           =   5
            TX              =   "SOLITUD CREDITO   "
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
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmVentas.frx":2774B
            PICN            =   "FrmVentas.frx":27767
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
      Begin VB.TextBox txtpreciooriginal 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   69
         Top             =   4550
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkVincular 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "VINCULAR"
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
         Left            =   12120
         TabIndex        =   68
         Top             =   475
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtBuscarVendedor 
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
         Left            =   6240
         MaxLength       =   80
         TabIndex        =   67
         Top             =   2865
         Width           =   615
      End
      Begin VB.CheckBox chkconyuge 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "CONYUGE"
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
         Height          =   285
         Left            =   3600
         TabIndex        =   66
         Top             =   1430
         Width           =   1095
      End
      Begin VB.TextBox txttipofactura 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   65
         Text            =   "00001"
         Top             =   4190
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   1450
         TabIndex        =   64
         Top             =   6905
         Width           =   585
      End
      Begin VB.TextBox TxtCodProducto 
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
         Left            =   240
         TabIndex        =   63
         Top             =   6905
         Width           =   1185
      End
      Begin VB.TextBox txtprecio 
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
         Left            =   8505
         MaxLength       =   80
         TabIndex        =   62
         Top             =   6905
         Width           =   855
      End
      Begin VB.TextBox txteditable 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   61
         Text            =   "no"
         Top             =   5270
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtafectacaja 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   60
         Text            =   "no"
         Top             =   350
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtserial 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   54
         Top             =   830
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txt_tipo_movimiento 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   53
         Text            =   "01"
         Top             =   590
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtformato_impresion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   36
         Top             =   110
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txt_hash 
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
         Left            =   14520
         TabIndex        =   24
         ToolTipText     =   "TELEFONO"
         Top             =   5745
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtnumeroguia 
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
         Height          =   300
         Left            =   17760
         TabIndex        =   23
         Top             =   2075
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtserieguia 
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
         Height          =   300
         Left            =   17175
         TabIndex        =   22
         Top             =   2075
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox txt_sunat_key 
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
         Left            =   14520
         MaxLength       =   80
         TabIndex        =   21
         ToolTipText     =   "TELEFONO"
         Top             =   6480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chk_direccion 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6885
         TabIndex        =   14
         Top             =   2150
         Width           =   255
      End
      Begin VB.CheckBox chk_nueva_guia 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "NUEVA GUIA"
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
         Height          =   300
         Left            =   15600
         TabIndex        =   10
         Top             =   2075
         Width           =   1455
      End
      Begin VB.CheckBox chk_seleccionar_guia 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "SELECC GUIA "
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
         Height          =   300
         Left            =   15600
         TabIndex        =   9
         Top             =   2440
         Width           =   1455
      End
      Begin VB.TextBox txtmayor 
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
         Left            =   14520
         MaxLength       =   80
         TabIndex        =   8
         Text            =   "no"
         ToolTipText     =   "TELEFONO"
         Top             =   6105
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtObservacion 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
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
         Height          =   480
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   7310
         Width           =   3735
      End
      Begin VB.TextBox txtguia_referencia 
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
         Height          =   315
         Left            =   19120
         TabIndex        =   6
         Top             =   2430
         Visible         =   0   'False
         Width           =   820
      End
      Begin VB.TextBox txt_tipo 
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
         Left            =   14520
         MaxLength       =   80
         TabIndex        =   1
         ToolTipText     =   "TELEFONO"
         Top             =   7065
         Visible         =   0   'False
         Width           =   975
      End
      Begin VitekeySoft.ChameleonBtn cmdprocesar 
         Height          =   780
         Left            =   120
         TabIndex        =   31
         Top             =   8400
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1376
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
         MICON           =   "FrmVentas.frx":29ABB
         PICN            =   "FrmVentas.frx":29AD7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdcopropietario 
         Height          =   405
         Left            =   5400
         TabIndex        =   44
         Top             =   1350
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   714
         BTYPE           =   5
         TX              =   "CO-PROPIETARIO"
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
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas.frx":2D11F
         PICN            =   "FrmVentas.frx":2D13B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdcredito 
         Height          =   405
         Left            =   4800
         TabIndex        =   59
         Top             =   1350
         Visible         =   0   'False
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   714
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas.frx":2FF2B
         PICN            =   "FrmVentas.frx":2FF47
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn CmdAgregar 
         Height          =   345
         Left            =   11050
         TabIndex        =   77
         ToolTipText     =   "AGREGAR ITEM"
         Top             =   6870
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MICON           =   "FrmVentas.frx":330BD
         PICN            =   "FrmVentas.frx":330D9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpActual 
         Height          =   340
         Left            =   12120
         TabIndex        =   89
         Top             =   105
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   171966465
         CurrentDate     =   40579
      End
      Begin MSComCtl2.DTPicker DtpFechaReferencia 
         Height          =   255
         Left            =   9240
         TabIndex        =   96
         Top             =   5025
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
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
         Format          =   171966465
         CurrentDate     =   40579
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   405
         Left            =   120
         TabIndex        =   107
         Top             =   195
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   714
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DtcTipoDoc 
         Height          =   330
         Left            =   7440
         TabIndex        =   108
         Top             =   105
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo DtcFormaPago 
         Height          =   330
         Left            =   8445
         TabIndex        =   109
         Top             =   1590
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   8454143
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
      Begin MSComCtl2.DTPicker DTPDetracion 
         Height          =   375
         Left            =   12120
         TabIndex        =   110
         Top             =   105
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   0
         CalendarForeColor=   255
         Format          =   171966465
         CurrentDate     =   39535
      End
      Begin MSDataListLib.DataCombo DtcComprobanteGuia 
         Height          =   330
         Left            =   1560
         TabIndex        =   111
         Top             =   870
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
      Begin MSDataListLib.DataCombo DtTargeta 
         Height          =   330
         Left            =   11400
         TabIndex        =   112
         Top             =   1230
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   8454143
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
         Left            =   8445
         TabIndex        =   113
         Top             =   1230
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         BackColor       =   8454143
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgTipoPagos 
         Height          =   1095
         Left            =   8445
         TabIndex        =   114
         Top             =   2340
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   1931
         _Version        =   393216
         ForeColor       =   8388608
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         ForeColorFixed  =   8388608
         ForeColorSel    =   16777215
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
      Begin MSDataListLib.DataCombo DtcFormapagodetalle 
         Height          =   330
         Left            =   8445
         TabIndex        =   115
         Top             =   1950
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   8454143
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
      Begin TabDlg.SSTab SSTab1 
         Height          =   3255
         Left            =   15600
         TabIndex        =   116
         Top             =   5940
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   5741
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         BackColor       =   16777215
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
         TabCaption(0)   =   "ESTADO COMPROBANTES"
         TabPicture(0)   =   "FrmVentas.frx":33673
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "web_estado_sunat"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "COMPROBANTES"
         TabPicture(1)   =   "FrmVentas.frx":3368F
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "DtpFin"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "HfFacturas"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "DTPIni"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmdBuscarFecha"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         Begin SHDocVwCtl.WebBrowser web_estado_sunat 
            Height          =   2775
            Left            =   -74955
            TabIndex        =   195
            Top             =   360
            Visible         =   0   'False
            Width           =   4320
            ExtentX         =   7620
            ExtentY         =   4895
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
         Begin VB.CommandButton cmdBuscarFecha 
            BackColor       =   &H008080FF&
            Caption         =   "BUSCAR"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   117
            Top             =   360
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPIni 
            Height          =   300
            Left            =   120
            TabIndex        =   118
            Top             =   360
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
            Format          =   171966465
            CurrentDate     =   41751
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfFacturas 
            Height          =   2475
            Left            =   120
            TabIndex        =   119
            Top             =   720
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   4366
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
         Begin MSComCtl2.DTPicker DtpFin 
            Height          =   300
            Left            =   1560
            TabIndex        =   120
            Top             =   360
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
            Format          =   171966465
            CurrentDate     =   41751
         End
      End
      Begin VitekeySoft.ChameleonBtn CmdQuitar 
         Height          =   345
         Left            =   11880
         TabIndex        =   123
         ToolTipText     =   "ELIMINAR ITEM"
         Top             =   6870
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MICON           =   "FrmVentas.frx":336AB
         PICN            =   "FrmVentas.frx":336C7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdSeriales 
         Height          =   345
         Left            =   12285
         TabIndex        =   124
         ToolTipText     =   "VISUALIZAR SERIES"
         Top             =   6870
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MICON           =   "FrmVentas.frx":33C61
         PICN            =   "FrmVentas.frx":33C7D
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
         Left            =   3480
         TabIndex        =   131
         Top             =   2820
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
         Height          =   3135
         Left            =   120
         TabIndex        =   152
         Top             =   3585
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   5530
         _Version        =   393216
         ForeColor       =   8388608
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         ForeColorFixed  =   8388608
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
      Begin VitekeySoft.ChameleonBtn cmdimprimir 
         Height          =   780
         Left            =   1320
         TabIndex        =   155
         Top             =   8400
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1376
         BTYPE           =   5
         TX              =   "IMPRIMIR"
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
         MICON           =   "FrmVentas.frx":34217
         PICN            =   "FrmVentas.frx":34233
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdImprimirGuia 
         Height          =   300
         Left            =   19560
         TabIndex        =   156
         Top             =   2070
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas.frx":36804
         PICN            =   "FrmVentas.frx":36820
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdGrabarGuia 
         Height          =   300
         Left            =   19140
         TabIndex        =   157
         Top             =   2070
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   529
         BTYPE           =   5
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas.frx":368AD
         PICN            =   "FrmVentas.frx":368C9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcGuia 
         Height          =   315
         Left            =   17160
         TabIndex        =   158
         Top             =   2430
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
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
      Begin VitekeySoft.ChameleonBtn cmdAnular 
         Height          =   945
         Left            =   14520
         TabIndex        =   185
         Top             =   1000
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "ANULAR"
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
         MICON           =   "FrmVentas.frx":36E63
         PICN            =   "FrmVentas.frx":36E7F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcSerieDoc 
         Height          =   330
         Left            =   7440
         TabIndex        =   189
         Top             =   585
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo Dtcseguro 
         Height          =   315
         Left            =   3480
         TabIndex        =   193
         Top             =   3180
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   16777215
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
      Begin VitekeySoft.ChameleonBtn cmdModificar 
         Height          =   1020
         Left            =   14520
         TabIndex        =   215
         Top             =   2925
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1799
         BTYPE           =   5
         TX              =   "MODIFICAR"
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
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas.frx":37199
         PICN            =   "FrmVentas.frx":371B5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcUnidad 
         Height          =   330
         Left            =   7155
         TabIndex        =   216
         Top             =   6900
         Width           =   1310
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VitekeySoft.ChameleonBtn cmdPdf 
         Height          =   780
         Left            =   2520
         TabIndex        =   239
         Top             =   8400
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1376
         BTYPE           =   5
         TX              =   "PDF SUNAT"
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
         MICON           =   "FrmVentas.frx":397EE
         PICN            =   "FrmVentas.frx":3980A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdConversion 
         Height          =   345
         Left            =   11460
         TabIndex        =   240
         ToolTipText     =   "ELIMINAR ITEM"
         Top             =   6870
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
         BTYPE           =   5
         TX              =   ""
         ENAB            =   0   'False
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
         MICON           =   "FrmVentas.frx":39C5C
         PICN            =   "FrmVentas.frx":39C78
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "ICBPER :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   250
         Index           =   6
         Left            =   5025
         TabIndex        =   237
         Top             =   8550
         Width           =   765
      End
      Begin VB.Label lblicbper 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   6000
         TabIndex        =   236
         Top             =   8550
         Width           =   1605
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL VENTA :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   9030
         TabIndex        =   235
         Top             =   7720
         Width           =   1605
      End
      Begin VB.Label lblPercepcion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   12720
         TabIndex        =   234
         Tag             =   "no"
         Top             =   8850
         Width           =   1245
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERCEPCION :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   11250
         TabIndex        =   233
         Top             =   8880
         Width           =   1425
      End
      Begin VB.Label lblhistorial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9240
         TabIndex        =   232
         Top             =   7365
         Width           =   4725
      End
      Begin VB.Label lblConversion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   2
         Left            =   12720
         TabIndex        =   213
         Top             =   8460
         Width           =   1245
      End
      Begin VB.Label lblConversion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   12720
         TabIndex        =   212
         Top             =   8070
         Width           =   1245
      End
      Begin VB.Label lblConversion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   12720
         TabIndex        =   211
         Top             =   7680
         Width           =   1245
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IGV :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   3
         Left            =   5325
         TabIndex        =   207
         Top             =   8295
         Width           =   465
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GRATUITAS:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   2
         Left            =   4635
         TabIndex        =   206
         Top             =   7980
         Width           =   1155
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR VENTA:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   1
         Left            =   4395
         TabIndex        =   205
         Top             =   7680
         Width           =   1395
      End
      Begin VB.Label lblGratuitas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   6000
         TabIndex        =   204
         Top             =   7950
         Width           =   1605
      End
      Begin VB.Label lblpeso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   12720
         TabIndex        =   201
         Top             =   6870
         Width           =   975
      End
      Begin VB.Label lblSaldodisponible 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   585
         Left            =   8445
         TabIndex        =   165
         Top             =   1950
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.Label Label35 
         BackColor       =   &H008080FF&
         Caption         =   "T. CAMBIO :"
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
         Height          =   310
         Index           =   5
         Left            =   10440
         TabIndex        =   199
         Top             =   120
         Width           =   975
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Height          =   2250
         Left            =   120
         Top             =   1305
         Width           =   7095
      End
      Begin VB.Shape ShpDatos 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   120
         Top             =   45
         Width           =   7095
      End
      Begin VB.Shape ShaGuia 
         BackColor       =   &H00DFDFE0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   120
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label LblCantidad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   13725
         TabIndex        =   182
         Top             =   6870
         Width           =   495
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   10680
         TabIndex        =   181
         Top             =   7680
         Width           =   1965
      End
      Begin VB.Label LblValorVenta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   6000
         TabIndex        =   180
         Top             =   7650
         Width           =   1605
      End
      Begin VB.Label lblExonerado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   6000
         TabIndex        =   179
         Top             =   7350
         Width           =   1605
      End
      Begin VB.Label LblComprobante_DR 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FORMA PAGO:"
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
         Left            =   7350
         TabIndex        =   178
         Top             =   1665
         Width           =   1035
      End
      Begin VB.Label LblTotalLetras 
         BackColor       =   &H80000007&
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   480
         TabIndex        =   177
         Top             =   9345
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label LblIgv 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   6000
         TabIndex        =   176
         Top             =   8250
         Width           =   1605
      End
      Begin VB.Label lblPago 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   10680
         TabIndex        =   175
         Top             =   8070
         Width           =   1965
      End
      Begin VB.Label lblVuelto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   10680
         TabIndex        =   174
         Top             =   8460
         Width           =   1965
      End
      Begin VB.Label lbltargeta 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TARJETA :"
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
         Left            =   7680
         TabIndex        =   173
         Top             =   2025
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOBRANTE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   8655
         TabIndex        =   172
         Top             =   9705
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblSobrante 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   10560
         TabIndex        =   171
         Top             =   9780
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONEDA :"
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
         Left            =   7665
         TabIndex        =   170
         Top             =   1305
         Width           =   735
      End
      Begin VB.Label lblPendientes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   360
         TabIndex        =   169
         Top             =   7545
         Width           =   45
      End
      Begin VB.Label lblfactura_masiva 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "no"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   360
         TabIndex        =   168
         Top             =   3840
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   615
         Left            =   14040
         Top             =   7455
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PAGO :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9150
         TabIndex        =   167
         Top             =   8100
         Width           =   1500
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VUELTO :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9675
         TabIndex        =   166
         Top             =   8475
         Width           =   975
      End
      Begin VB.Image imgFoto 
         Height          =   1935
         Left            =   240
         Picture         =   "FrmVentas.frx":3A212
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXONERADO:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   0
         Left            =   4545
         TabIndex        =   164
         Top             =   7350
         Width           =   1245
      End
      Begin VB.Label lblDisponible 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   120
         TabIndex        =   163
         Top             =   8120
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label lblContabilidad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INGRESADO POR AREA CONTABLE"
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
         Left            =   7680
         TabIndex        =   162
         Top             =   8925
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "VENDEDOR :"
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
         Index           =   0
         Left            =   2160
         TabIndex        =   161
         Top             =   2865
         Width           =   1095
      End
      Begin VB.Label lblregistradopor 
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
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   7620
         TabIndex        =   160
         Top             =   8880
         Width           =   3600
      End
      Begin VB.Label LblTotalParcial 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   9390
         TabIndex        =   159
         Top             =   6900
         Width           =   1155
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00400000&
         FillStyle       =   0  'Solid
         Height          =   1890
         Left            =   3960
         Top             =   7305
         Width           =   10040
      End
      Begin VB.Shape ShapeDR 
         BackColor       =   &H00DFDFE0&
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   2310
         Left            =   7320
         Top             =   1185
         Width           =   6975
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   435
         Left            =   120
         Top             =   6825
         Width           =   14175
      End
      Begin VB.Shape Shape9 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   345
         Left            =   15600
         Top             =   30
         Width           =   4455
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         FillColor       =   &H000080FF&
         Height          =   9240
         Left            =   0
         Top             =   0
         Width           =   20175
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1050
         Left            =   7320
         Top             =   30
         Width           =   6975
      End
   End
End
Attribute VB_Name = "FrmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doc_Tienda As String * 1
Dim cod_doc As String
Dim rstTemporal As New ADODB.Recordset
Dim StrCodDetVenta As String
Dim StrCodReferencia As Double
Dim Referencia As Boolean
Dim dfactura As String
Dim total_descuento As Single
Dim delivery As String
Dim Descuento As Single
Dim strEspecial As Integer
Dim cQrCode As ClsQrCode
Dim in_chasis As String
Dim in_motor As String
Dim in_total_documento As Double

Public numeroItem As Integer
Public Procedencia As EnumProcede
Public ProcendenciaGuia As EnumGuia
Public codigoP As String
Public in_codigo_habitacion As String
Public IN_ICBPER As String


Public Sub facturar_agenda(ByVal in_agenda As String)
   
   Call Me.activar
   Me.txtid_agenda.Text = in_agenda
   strCadena = "CALL procedure_agenda('9','" & Val(in_agenda) & "','','','','','','','','','','" & KEY_RUC & "')"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
        in_detalle = rst("nombre_prod")
        Me.TxtCodCliente.Text = rst("dni_cliente")
        
        
        strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,servicio,obsequio,costo) VALUES " & _
        "('" & KEY_RUC & "','" & rst("id_unidad") & "','" & rst("dni_cliente") & "','" & KEY_ALM & "','" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcSerieDoc.BoundText & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & rst("id_producto") & "','1'," & _
        "'" & rst("precio_venta") & " ','" & rst("precio_venta") & "','0','" & KEY_APLICA_IGV & "','" & in_detalle & "','" & KEY_USUARIO & "','no','no','0')"
         CnBd.Execute (strCadena)
         Call precionar_cliente
   End If
                    
    Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcSerieDoc.BoundText, Me.DtcTipoDoc.BoundText, Trim(Me.txtformato_impresion.Text))
       
         
        
   

End Sub



Private Sub parametro_importacion()
Me.lblmotor.Tag = "no"

strCadena = "SELECT * FROM parametros_produccion WHERE habilitado='si' and  ruc='" & KEY_RUC & "' ORDER BY id asc"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   rstK.MoveFirst
   For i = 0 To rstK.RecordCount - 1
       
       If rstK("codigo") = "vin" Then
         in_chasis = rstK("descripcion")
         Me.lblchasis.Caption = in_chasis
         
       End If
       
       If rstK("codigo") = "chasis" Then
         in_chasis = rstK("descripcion")
         Me.lblchasis.Caption = in_chasis
       End If
       
       
       If rstK("codigo") = "motor" Then
          in_motor = rstK("descripcion") & ":"
          Me.lblmotor.Caption = in_motor
          Me.lblmotor.Tag = "si"
       End If
    
      
    
    
    
    rstK.MoveNext
   Next i
   
   If Me.lblmotor.Tag = "no" Then
      Me.lblmotor.Visible = False
      Me.DtcMotor.Visible = False
   End If
   
End If

End Sub
Private Function get_repetido(ByVal in_producto As String) As Boolean
strCadena = "SELECT id_producto FROM temporal_ventas WHERE obsequio='no' and  id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   MsgBox "ESTE PRODUCTO YA ESTA EN LA ORDEN DE VENTA !!!", vbInformation, "Modulo de ITEM'S DUPLICADOS [ACTIVADO]"
   get_repetido = True
Else
   get_repetido = False
End If
End Function
Public Sub insertar_item()
Dim in_total_parcial As Double
Dim in_obsequio As String

If Val(Me.txtprecio.Text) <= 0 Then
   MsgBox "INGRESE UN PRECIO DE VENTA", vbInformation, KEY_USUARIO
   Call Resalta(Me.txtprecio)
   Exit Sub
End If

If Me.chkObsequio.Value = 1 Then
    in_obsequio = "si"
Else
    in_obsequio = "no"
End If


If Val(Me.txtCantidad.Text) > 0 And Trim(Me.TxtCodProducto.Text) <> "" Then
        Me.txtmayor.Text = "no"
        If KEY_RUC = "20219520281" And KEY_ALM = "00002" Then
            in_total_parcial = Val(Me.txtprecio.Text)
            Me.txtprecio.Text = Val(Me.txtprecio.Text) / Val(Me.txtCantidad.Text)
        Else
            in_total_parcial = Val(Me.txtprecio.Text) * Val(Me.txtCantidad.Text)
        End If
        
        If in_obsequio = "si" Then
            in_total_parcial = 0
        End If
        
        If KEY_PRODUCTO_REPETIDO = "no" Then
           If get_repetido(Trim(codigoP)) = True Then
              Exit Sub
           End If
       End If
        
       If get_costo_producto(codigoP) <= 0 And Me.chk_venta_diferida.Value = 0 Then
            MsgBox "PRODUCTO con PRECIO COSTO =[  0.00  ]" + Chr(13) + "Coordine con Centro de Costos", vbInformation, KEY_VENDEDOR
            Exit Sub
       End If
      
        
       ' If KEY_DESCUENTO_LINEA = "si" Then
       '    in_precio = Val(Me.txtprecio.Text) - Val(Me.txtprecio.Text) * (put_descuento_categoria(codigoP, Val(Me.txtprecio.Text))) / 100
       '    in_total_parcial = Val(in_precio) * Val(Me.txtCantidad.Text)
       ' Else
           in_precio = Val(Me.txtprecio.Text)
       ' End If
        
            strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,servicio,obsequio,costo,agranel,icbper) VALUES " & _
            "('" & KEY_RUC & "','" & Me.DtcUnidad.BoundText & "','" & Trim(Me.TxtCodCliente.Text) & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.DtcSerieDoc.BoundText) & "','" & Me.TxtNumeroDoc.Text & "','" & codigoP & "','" & Val(Me.txtCantidad.Text) & "'," & _
            "'" & Val(in_precio) & " ','" & in_total_parcial & "','" & Val(Me.txtpeso.Text) & "','" & Trim(Me.TxtIgv.Text) & "','" & UCase(Trim(Me.TxtDescripcionProducto.Text)) & "','" & KEY_USUARIO & "','" & Trim(Me.txtServicio.Text) & "','" & in_obsequio & "','" & get_precio_costo(codigoP) & "','" & Trim(Me.txtagranel.Text) & "','" & IN_ICBPER & "')"
            CnBd.Execute (strCadena)
        
        
        '**************VERIFICA BONIFICACIONES
        If KEY_BONIFICACIONES = "si" Then
            strCadena = "CALL get_idTemporalventas('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
            Call ConfiguraRst(strCadena)
            in_idVenta = rst(0)
            
            
            strCadena = "CALL put_bonificacion_linea('" & codigoP & "','" & Trim(Me.TxtCodCliente.Text) & "','" & KEY_USUARIO & "','" & KEY_ALM & "','" & in_idVenta & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
            Call put_verificar_bonificacion_monto(codigoP, Me.txtCantidad.Text, Trim(Me.TxtCodCliente.Text), Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText)
            Call put_verificar_bonificacion_cruzada_v2(codigoP, Me.txtCantidad.Text, Trim(Me.TxtCodCliente.Text), Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText)
            
            
           End If
        
        Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcSerieDoc.BoundText, Me.DtcTipoDoc.BoundText, Trim(Me.txtformato_impresion.Text))
        Call VerificaDocumento(Trim(Me.DtcTipoDoc.BoundText))
        strCadena = "SELECT L.produccion FROM producto P,linea L WHERE P.id_linea=L.id_linea AND P.ruc=L.id_usu AND P.id_producto='" & Trim(codigoP) & "' AND P.ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
    
    If rstL("produccion") = "si" Then
        strCadena = "SELECT * FROM temporal_ventas WHERE ruc='" & KEY_RUC & "' and id_doc='" & Me.DtcTipoDoc.BoundText & "' and id_serie='" & Me.DtcSerieDoc.BoundText & "' and dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' ORDER BY id DESC LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
           Me.txtid_temporal_serie.Text = rstL("id")
        End If
        
        strCadena = "SELECT modelo,color,marca FROM view_producto WHERE id_producto='" & Trim(Me.TxtCodProducto.Text) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        Me.txtModelo.Text = rstL("modelo")
        Me.txtcolor.Text = rstL("color")
        Me.TxtMarca.Text = rstL("marca")
        Me.txttipofactura.Text = "00002"
        
        strCadena = "SELECT Codigo,nro_chasis as Descripcion FROM view_producto_serie WHERE  id_alm='" & KEY_ALM & "' and   vendido='no'  and id_producto='" & Trim(Me.TxtCodProducto.Text) & "' AND  ruc='" & KEY_RUC & "'"
        Call ConfiguraRstT(strCadena)
        Call LlenaDataComboT(Me.DtcSerie)
        Me.FrameSerieModelo.Visible = True
        Me.cmdSeriales.Visible = True
        Me.cmdProcesar.Enabled = True
        Me.TxtCodProducto.Text = "00000"
        Me.txtCantidad.Text = "0"
        Me.TxtDescripcionProducto.Text = ""
        Me.txtprecio.Text = ""
        Me.DtcUnidad.BoundText = 0
        Me.LblTotalParcial.Caption = ""
        chkPrecios.Enabled = False
        Me.HfPrecios.Visible = False
        Call Resalta(Me.txtBuscarSerie)
        Call parametro_importacion
        Exit Sub
    Else
        Me.txttipofactura.Text = "00001"
    End If
    
    
    'BONIFICACIONES
    
    
    Me.cmdProcesar.Enabled = True
    Me.TxtCodProducto.Text = "00000"
    Me.txtCantidad.Text = "0"
    Me.TxtDescripcionProducto.Text = ""
    Me.txtprecio.Text = ""
        Me.DtcUnidad.BoundText = 0
    Me.LblTotalParcial.Caption = ""
    Me.chkObsequio.Value = 0
    chkPrecios.Enabled = False
    Me.HfPrecios.Visible = False
    Call Resalta(Me.TxtCodProducto)
   'Call DisplayTextoCom("TOTAL : S/." & AlineaString(Me.lblTotal.Caption, 9, pAlnDerecha) & _
                            "VUELTO: S/." & AlineaString(Me.lblVuelto.Caption, 9, pAlnDerecha), mscConecta)
    
Else
    Call Resalta(Me.txtCantidad)
End If

End Sub
Public Sub insertar_item_save()
Dim in_total_parcial As Double
Dim in_obsequio As String

If Val(Me.txtprecio.Text) <= 0 Then
   MsgBox "INGRESE UN PRECIO DE VENTA", vbInformation, KEY_USUARIO
   Call Resalta(Me.txtprecio)
   Exit Sub
End If

If Me.chkObsequio.Value = 1 Then
    in_obsequio = "si"
Else
    in_obsequio = "no"
End If


If Val(Me.txtCantidad.Text) > 0 And Trim(Me.TxtCodProducto.Text) <> "" Then
        Me.txtmayor.Text = "no"
        If KEY_RUC = "20219520281" And KEY_ALM = "00002" Then
            in_total_parcial = Val(Me.txtprecio.Text)
            Me.txtprecio.Text = Val(Me.txtprecio.Text) / Val(Me.txtCantidad.Text)
        Else
            in_total_parcial = Val(Me.txtprecio.Text) * Val(Me.txtCantidad.Text)
        End If
        
        If in_obsequio = "si" Then
                        in_total_parcial = 0
        End If
        
        If KEY_PRODUCTO_REPETIDO = "no" Then
           If get_repetido(Trim(Me.TxtCodProducto.Text)) = True Then
              Exit Sub
           End If
       End If
      
       If IN_ICBPER = "si" Then
          strCadena = "UPDATE movimiento_venta SET icbper='" & Val(Me.txtCantidad.Text) * KEY_VALOR_BOLSA & "' WHERE id_venta='" & Val(Me.txtidVenta.Text) & "'"
          CnBd.Execute (strCadena)
       End If
        
        
        strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,cantidad,precio,total,peso,igv,detalle,obsequio,icbper,ruc) VALUES " & _
        "('" & Val(Me.txtidVenta.Text) & "','" & codigoP & "','" & Val(Me.txtCantidad.Text) & "','" & Val(Me.txtprecio.Text) & "','" & in_total_parcial & "','" & Val(Me.txtpeso.Text) & "','" & Trim(Me.TxtIgv.Text) & "','" & Trim(Me.TxtDescripcionProducto.Text) & "','" & in_obsequio & "','" & IN_ICBPER & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        Call Me.llenarGrid_Comprobante_edit(Me.HfdDetalle, Val(Me.txtidVenta.Text))
        
    
    
    Me.cmdProcesar.Enabled = True
    Me.TxtCodProducto.Text = "00000"
    Me.txtCantidad.Text = "0"
    Me.TxtDescripcionProducto.Text = ""
    Me.txtprecio.Text = ""
    Me.DtcUnidad.Text = ""
    Me.LblTotalParcial.Caption = ""
    Me.chkObsequio.Value = 0
    chkPrecios.Enabled = False
    Me.HfPrecios.Visible = False
    Call Resalta(Me.TxtCodProducto)
   
    
Else
    Call Resalta(Me.txtCantidad)
End If

End Sub



Private Sub VerificaDocumento(ByVal TipoDoc As String)
'If Trim(Me.DtcTipoDoc.BoundText) = "0009" Then
'    Me.TlbGrabar.Buttons(KEY_GUIAREMISION).Enabled = True
'End If
End Sub
Sub ModificarCantidad(ByVal Can_Previa As Integer, ByVal PesoProd As Double, ByVal serie As String, ByVal TipoDoc As String, ByVal numero As String)
    Dim Can_actual As Integer
    Can_actual = Can_Previa + Val(Me.txtCantidad.Text)
    strCadena = "UPDATE Temporal_Ventas SET cantidad='" & Can_actual & "',Total='" & (Can_actual * Val(Me.txtprecio.Text)) & "'," & _
                "Peso='" & PesoProd & "' WHERE cProducto='" & Trim(Me.TxtCodProducto.Text) & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "' AND " & _
                "cDocumentoVenta='" & numero & "'"
    Call EjecutaRST(strCadena)
    Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText, Me.txtformato_impresion.Text)
End Sub

Function GeneraCodTemporal() As Integer
Dim Codtemporal As Integer
strCadena = "SELECT cTemporal FROM Temporal_Ventas ORDER BY cTemporal DESC"
Call ConfiguraRst(strCadena)
    If rst.EOF Or rst.BOF = True Then
        Codtemporal = 1
    Else
        Codtemporal = rst(0) + 1
    End If
  GeneraCodTemporal = Codtemporal
  Set rst = Nothing
End Function
Function GeneraCodReferencia() As Integer
Dim CodReferencia As Integer
strCadena = "SELECT IdReferencia FROM DocReferencia_Venta ORDER BY IdReferencia DESC "
Call ConfiguraRst(strCadena)
    If rst.EOF = True Then
        CodReferencia = 1
        
    Else
        CodReferencia = rst(0) + 1

    End If
  GeneraCodReferencia = CodReferencia
  
  
  Set rst = Nothing
End Function

Public Sub put_interes_fraccionado(ByVal in_monto_tarjeta As Double)

strCadena = "SELECT ifnull(sum(monto_caja),0) FROM movimiento_venta_monto_temporal WHERE ruc='" & KEY_RUC & "' and id_alm='" & KEY_ALM & "' and dni_save='" & KEY_USUARIO & "'"
Call ConfiguraRstA(strCadena)



End Sub
Public Sub put_incrementar_interes_tarjeta(ByVal in_porcentaje As Single, ByVal in_monto As Double)
Dim in_precio As Double
Dim in_total As Double
Dim in_porcen As Single

strCadena = "SELECT * FROM temporal_ventas WHERE ruc='" & KEY_RUC & "' and id_doc='" & Me.DtcTipoDoc.BoundText & "' and id_serie='" & Me.DtcSerieDoc.BoundText & "' and dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' ORDER BY id DESC"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   in_porcen = (in_monto * in_porcentaje / 100) / rstA.RecordCount
   
   For i = 0 To rstA.RecordCount - 1
        
        in_precio = rstA("precio") + in_porcen / rstA("cantidad")
        in_total = in_precio * rstA("cantidad")
        
        strCadena = "UPDATE temporal_ventas SET precio_neutro='" & rstA("precio") & "',precio='" & in_precio & "',total='" & in_total & "' WHERE id='" & rstA("id") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        rstA.MoveNext
   Next i
End If

Me.TxtMontoPagado.Text = in_monto + in_monto * in_porcentaje / 100

Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcSerieDoc.BoundText, Me.DtcTipoDoc.BoundText, Trim(Me.txtformato_impresion.Text))

End Sub


Public Sub llenarGrid_det(ByVal Grilla As MSHFlexGrid, ByVal id_numero As String, ByVal id_serie As String, ByVal id_doc As String, ByVal in_tipo_impresion As String)
On Error GoTo salir
Dim tTotal As Double
Dim texonerado As Double
Dim tafecto As Double
Dim in_peso As Single
Dim in_servicio As String
Dim in_contador As Integer
Dim in_cantidad_vendida As Integer

in_contador = 0
in_cantidad_vendida = 0
in_total_documento = 0
in_total_icbper = 0



strCadena = "SELECT * FROM view_venta_temporal_ultimate WHERE id_doc='" & id_doc & "' and id_serie='" & id_serie & "'  and dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.lblVuelto.Caption = 0#
    
    
    Me.lblCantidad.Caption = "0"
    Me.lblExonerado.Caption = ""
    Me.LblValorVenta.Caption = ""
    Me.LblIgv.Caption = ""
    Me.TxtDescuento_global.Text = ""
    Me.lblGratuitas.Caption = ""
    Me.lblTotal.Caption = ""
    Me.lblPago.Caption = ""
    Me.lblpeso.Caption = ""
    Me.txtServicio.Text = "no"
    Me.lblCantidad.Caption = 0
    Me.lblPercepcion.Caption = 0#
    Me.lblicbper.Caption = ""
    
    
    
    Grilla.Rows = 0
    Exit Sub
End If
   
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 6100
           Grilla.ColWidth(3) = 1500
           Grilla.ColWidth(4) = 1300
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1300
           Grilla.ColWidth(7) = 1300
       Next
        cabecera = "IDTEMPORAL" & vbTab & "CODIGO" & vbTab & "DESCRIPCION " & vbTab & "MARCA " & vbTab & "UND " & vbTab & "CANTIDAD" & vbTab & "PRECIO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 1 To 7
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        texonerado = 0
        tafecto = 0
        strEspecial = 0
        in_peso = 0
        in_gratuita = 0
        in_total_documento = 0
        Me.lblCantidad.Caption = rst.RecordCount
        in_contador = 0
       
        
        For i = 0 To rst.RecordCount - 1
            If rst("id_producto") = KEY_COD_PER Then
                in_marca = ""
                in_abreviatura = ""
            Else
                in_marca = rst("marca")
                in_abreviatura = rst("unidad")
            End If
            If rst("obsequio") = "si" Then
                in_total_parcial = 0
            Else
                in_total_parcial = rst("total")
            End If
            Fila = rst("id") & vbTab & rst("id_producto") & vbTab & rst("detalle") & vbTab & in_marca & vbTab & in_abreviatura & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("precio"), "#,##0.00") & vbTab & Format(in_total_parcial, "#,##0.00")
            Grilla.AddItem Fila
            in_cantidad_vendida = in_cantidad_vendida + rst("cantidad")
            
            If KEY_CON_IGV = "no" Then
                texonerado = texonerado + in_total_parcial
                
                If KEY_IMPUESTO_BOLSAS = "si" Then
                   If rst("icbper") = "si" Then
                      in_total_icbper = in_total_icbper + KEY_VALOR_BOLSA * rst("cantidad")
                   End If
                End If
            Else
                If KEY_APLICA_IGV = "no" Then
                    texonerado = texonerado + in_total_parcial
                Else
                    tafecto = tafecto + in_total_parcial
                    If KEY_IMPUESTO_BOLSAS = "si" Then
                        If rst("icbper") = "si" Then
                            in_total_icbper = in_total_icbper + KEY_VALOR_BOLSA * rst("cantidad")
                        End If
                    End If
                End If
            End If
            
            
            
            
            If rst("servicio") = "no" Then
                in_contador = in_contador + 1
            End If
             
             If rst("obsequio") = "si" Then
                in_gratuita = in_gratuita + rst("precio") * rst("cantidad")
                For k = 5 To 7
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
             End If
             
             
             
             in_peso = in_peso + rst("peso") * rst("cantidad")
            
            rst.MoveNext
    Next i
    
    If in_contador > 0 Then
       txt_tipo.Text = "01"
       Me.txtServicio.Text = "no"
    Else
       txt_tipo.Text = "02"
       Me.txtServicio.Text = "si"
    End If
    
    
  tTotal = texonerado + tafecto
  Me.lblPercepcion.Caption = 0
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 5
  Grilla.RowSel = 1





If KEY_CON_IGV = "si" Then
    If KEY_APLICA_IGV = "no" Then
        texonerado = texonerado
        SUBTOTAL = texonerado
        igv = 0
        tTotal = SUBTOTAL
    Else
        tTotal = tafecto
        SUBTOTAL = tafecto / (1 + KEY_IGV)
        igv = tafecto - SUBTOTAL
        
        tafecto = SUBTOTAL
    End If
    
    
Else
    texonerado = texonerado
    SUBTOTAL = texonerado
    igv = 0
    tTotal = SUBTOTAL
End If




in_total_documento = tTotal
If texonerado > 0 Then
    Me.lblExonerado.Caption = Format(texonerado, "###0.00")
End If
Me.lblicbper.Caption = Format(in_total_icbper, "###0.00")
Me.lblpeso.Caption = "PESO:" & Format(in_peso, "#,##0.00")
Me.lblTotal.Caption = Format(tTotal + Val(Me.lblPercepcion.Caption) + Val(Me.lblicbper.Caption), "###0.00")
Me.lblVuelto.Caption = Format(Val(Me.lblPago.Caption) - tTotal - Val(Me.lblPercepcion.Caption) - Val(Me.lblicbper.Caption), "###0.00")
Me.LblIgv.Caption = Format(igv, "###0.00")
Me.lblExonerado.Caption = Format(texonerado, "###0.00")
Me.lblGratuitas.Caption = Format(in_gratuita, "###0.00")
Me.LblValorVenta.Caption = Format(SUBTOTAL, "###0.00")
Me.txtCantidad.Text = 0


   
   If KEY_RUC = "20561358550" And Me.lblPercepcion.Tag = "si" Then
      If Me.DtcTipoDoc.BoundText = "0003" And in_cantidad_vendida > 2 Then
         Me.lblPercepcion.Caption = Format(tTotal * 2 / 100, "#,##0.00")
      End If
       If Me.DtcTipoDoc.BoundText = "0001" Then
         Me.lblPercepcion.Caption = Format(tTotal * 2 / 100, "#,##0.00")
      End If
      Me.lblTotal.Caption = tTotal + Val(Me.lblPercepcion.Caption) + Val(Me.lblicbper.Caption)
      Me.lblVuelto.Caption = Format(Val(Me.lblPago.Caption) - tTotal - Val(Me.lblPercepcion.Caption) - Val(Me.lblicbper.Caption), "###0.00")
   Else
      Me.lblPercepcion.Caption = "0.00"
   End If



Me.cmdAnular.Enabled = False
Me.cmdEliminar.Enabled = False
Me.cmdProcesar.Enabled = True



Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Public Sub llenar_impresoras(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
   Grilla.Clear
   Grilla.Rows = 0
   If Printers.Count > 0 Then
      Me.frmimpresoras.Visible = True
      Me.cmdenviarimpresion.Visible = True
    Else
        Me.frmimpresoras.Visible = False
      Me.cmdenviarimpresion.Visible = False
      Exit Sub
   End If
  

       ReDim arrColWidth(1 To Printers.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 400
           Grilla.ColWidth(1) = 3100
           
       Next
        cabecera = "N�" & vbTab & "IMPRESORA"
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        
            For i = 0 To Printers.Count - 1
                Fila = Format(i + 1, "00") & vbTab & UCase(Trim(Printers(i).DeviceName))
                Grilla.AddItem Fila
            Next i
salir:
Exit Sub
End Sub

Private Sub SaveReferencia(ByVal codigo As String, ByVal TipoDoc As String, ByVal serie As String, ByVal numero As String, ByVal fecha As Date, ByVal Almacen As String)
strCadena = "INSERT INTO DocReferencia_Venta(IdReferencia,doc_cod,sSerie,cDocumentoVenta,FechaProceso,Alm_Cod) VALUES " & _
            "('" & codigo & "','" & TipoDoc & "','" & serie & "','" & numero & "','" & fecha & "','" & Almacen & "')"
            Call EjecutaRST(strCadena)
            Set rst = Nothing
End Sub
Sub llenarGrid_Comprobante(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo salir
Dim tTotal As Double
Dim texonerado As Double
Dim tafecto As Double
Dim in_obsequio As Single
'strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & idVenta & "'"
If KEY_AGRANEL = "si" Then
    strCadena = "SELECT * FROM view_detalle_venta_agranel WHERE id_venta='" & idVenta & "' and ruc='" & KEY_RUC & "' ORDER BY id_detalle_venta ASC"
Else
    strCadena = "SELECT * FROM view_detalle_venta WHERE id_venta='" & idVenta & "' and ruc='" & KEY_RUC & "' ORDER BY id_detalle_venta ASC"
End If
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.lblContabilidad.Visible = True
    Grilla.Rows = 0
    in_obsequio = 0
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 6600
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1400
           Grilla.ColWidth(5) = 1400
           Grilla.ColWidth(6) = 1400
           Grilla.ColWidth(7) = 0
           'Grilla.ColAlignment(4) = 7
        Next
        cabecera = "IDDETALLE" & vbTab & "CODIGO" & vbTab & "DESCRIPCION " & vbTab & "UND " & vbTab & "CANTIDAD" & vbTab & "PRECIO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        in_obsequio = 0
        For i = 0 To rst.RecordCount - 1
            If rst("id_producto") = KEY_COD_PER Then
               in_producto = ""
               in_unidad = ""
               If rst("cantidad") = 0 Then
                  in_cantidad = ""
                Else
                  in_cantidad = Format(rst("cantidad"), "###0.00")
               End If
               
               If rst("precio") = 0 Then
                  in_precio = ""
                Else
                  in_precio = Format(rst("precio"), "###0.00")
               End If
               
            Else
              in_producto = rst("id_producto")
              in_unidad = rst("abreviatura")
              in_cantidad = Format(rst("cantidad"), "###0.00")
              in_precio = Format(rst("precio"), "###0.00")
            End If
            
            
            
            Fila = rst("id_detalle_venta") & vbTab & in_producto & vbTab & rst("detalle") & vbTab & in_unidad & vbTab & in_cantidad & vbTab & in_precio & vbTab & Format(rst("total"), "###0.00")
            Grilla.AddItem Fila
            If (Trim(rst("id_igv")) = "no") Then
                            texonerado = texonerado + rst("total")
                            For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0FFFF
                            Next k
             Else
                            tafecto = tafecto + rst("total")
             End If
             
            
             
             If rst("obsequio") = "si" Then
                in_obsequio = in_obsequio + in_precio * in_cantidad
                For k = 3 To 6
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
            Else
                 tTotal = tTotal + rst("total")
             End If
             
            
            rst.MoveNext
    Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1

Me.lblCantidad.Caption = Trim(rst.RecordCount)
Me.lblGratuitas.Caption = Format(in_obsequio, "###0.00")
'Me.lblTotal.Caption = Format(tTotal, "###0.000")
'Me.LblVuelto.Caption = Format(Val(Me.lblPago.Caption) - tTotal, "###0.000")

If KEY_CON_IGV = "si" Then
    SUBTOTAL = tafecto / (1 + KEY_IGV)
    igv = tafecto - SUBTOTAL
Else
    texonerado = tafecto
    SUBTOTAL = 0
    igv = 0
End If





If KEY_RUC = "20561358550" And Val(Me.lblPercepcion.Caption) > 0 Then
     
      Me.lblTotal.Caption = tTotal + Val(Me.lblPercepcion.Caption)
      Me.lblVuelto.Caption = Format(Val(Me.lblPago.Caption) - tTotal - Val(Me.lblPercepcion.Caption), "###0.00")
   Else
      Me.lblPercepcion.Caption = "0.00"
   End If


'Me.LblIgv.Caption = Format(igv, "###0.000")
'Me.LblValorVenta.Caption = Format(SUBTOTAL, "###0.000")
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub
Sub llenarGrid_Comprobante_edit(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo salir
Dim tTotal As Double
Dim texonerado As Double
Dim tafecto As Double
Dim in_obsequio As Single
'strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & idVenta & "'"

strCadena = "SELECT * FROM view_detalle_venta WHERE id_venta='" & idVenta & "' and ruc='" & KEY_RUC & "' ORDER BY id_detalle_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.lblContabilidad.Visible = True
    Grilla.Rows = 0
    in_obsequio = 0
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 6600
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1400
           Grilla.ColWidth(5) = 1400
           Grilla.ColWidth(6) = 1400
           Grilla.ColWidth(7) = 0
           'Grilla.ColAlignment(4) = 7
        Next
        cabecera = "IDDETALLE" & vbTab & "CODIGO" & vbTab & "DESCRIPCION " & vbTab & "UND " & vbTab & "CANTIDAD" & vbTab & "PRECIO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        in_obsequio = 0
        texonerado = 0
        For i = 0 To rst.RecordCount - 1
            If rst("id_producto") = KEY_COD_PER Then
               in_producto = ""
               in_unidad = ""
               If rst("cantidad") = 0 Then
                  in_cantidad = ""
                Else
                  in_cantidad = Format(rst("cantidad"), "###0.00")
               End If
               
               If rst("precio") = 0 Then
                  in_precio = ""
                Else
                  in_precio = Format(rst("precio"), "###0.00")
               End If
               
            Else
              in_producto = rst("id_producto")
              in_unidad = rst("abreviatura")
              in_cantidad = Format(rst("cantidad"), "###0.00")
              in_precio = Format(rst("precio"), "###0.00")
            End If
            
            nTotal = nTotal + rst("total")
            
            Fila = rst("id_detalle_venta") & vbTab & in_producto & vbTab & rst("detalle") & vbTab & in_unidad & vbTab & in_cantidad & vbTab & in_precio & vbTab & Format(rst("total"), "###0.00")
            Grilla.AddItem Fila
            If (Trim(rst("id_igv")) = "no") Then
                         '   texonerado = texonerado + rst("total")
                            For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0FFFF
                            Next k
             Else
                            tafecto = tafecto + rst("total")
             End If
             
             If KEY_CON_IGV = "no" Then
                texonerado = texonerado + rst("total")
             End If
             
             If rst("obsequio") = "si" Then
                in_obsequio = in_obsequio + in_precio * in_cantidad
                For k = 3 To 6
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
             End If
             
            
            rst.MoveNext
    Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1

Me.lblCantidad.Caption = Trim(rst.RecordCount)
Me.lblGratuitas.Caption = Format(in_obsequio, "###0.00")
Me.lblTotal.Caption = Format(nTotal, "###0.000")
Me.lblVuelto.Caption = ""
Me.lblExonerado.Caption = Format(texonerado, "###0.00")
If KEY_CON_IGV = "si" Then
    SUBTOTAL = tafecto / (1 + KEY_IGV)
    igv = tafecto - SUBTOTAL
Else
    texonerado = tafecto
    SUBTOTAL = 0
    igv = 0
End If


Me.LblIgv.Caption = Format(igv, "###0.000")
Me.LblValorVenta.Caption = Format(SUBTOTAL, "###0.000")
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Private Sub menudes1_linkClick()
    Dim i As Integer
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        MsgBox rst(0) + Chr(13) + "Debe" + Space(2) + str(rst(1)), vbInformation, "Mensaje para el Usuario"
        rst.MoveNext
    Next i
    Set rst = Nothing
End Sub

Private Sub chk_descuento_Click()
If Me.chk_descuento.Value = 1 And Val(Me.txtidVenta.Text) < 1 Then
    
    Procedencia = relacionar
    
    
    
    Me.TxtDescuento_porcentaje.Locked = False
    Me.TxtDescuento_global.Locked = False
    Call Resalta(Me.TxtDescuento_porcentaje)
    Exit Sub
Else
    
    Me.lblTotal.Caption = Format(in_total_documento, "###0.00")
    
    Me.TxtDescuento_porcentaje.Text = 0
    Me.TxtDescuento_global.Text = 0
    Me.TxtDescuento_porcentaje.Locked = True
    Me.TxtDescuento_global.Locked = True
End If
End Sub

Private Sub chk_direccion_Click()
If Me.chk_direccion.Value = 1 Then
   Call llenar_direccion(Me.hfdireccion, Trim(Me.TxtCodCliente.Text))
   Me.frmdireccion.Visible = True
End If
End Sub

Private Sub chk_factura_Click()
If (Me.chk_factura.Value = 1) Then
    dfactura = "si"
Else
    dfactura = "no"
End If
End Sub

Private Sub chk_interes_Click()
If Me.chk_interes.Value = 1 Then
   Me.txtporcentaje_interes.Visible = True
Else
   Me.txtporcentaje_interes.Visible = False
End If
End Sub

Private Sub chk_nueva_guia_Click()
If Me.chk_nueva_guia.Value = 1 Then
   Me.chk_seleccionar_guia.Value = 0
If Val(Me.txtidVenta.Text) > 0 Then
    strCadena = "SELECT serie,numero FROM movimiento_transferencia WHERE id_venta='" & Val(Me.txtidVenta.Text) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstZ(strCadena)
    If rstZ.RecordCount > 0 Then
       Me.txtserieguia.Text = rstZ("serie")
       Me.txtnumeroguia.Text = rstZ("numero")
       Me.txtserieguia.Locked = False
       Me.txtnumeroguia.Locked = False
       Me.txtserieguia.Visible = True
       Me.txtnumeroguia.Visible = True
       Me.cmdImprimirGuia.Visible = True
       Me.cmdGrabarGuia.Visible = True
    Else
        strCadena = "SELECT numero,serie,id_doc FROM almacen_comprobante WHERE id_doc='0009' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
                    
                    Me.txtserieguia.Text = rstT("serie")
                    Me.txtnumeroguia.Text = rstT("numero")
                    Me.txtserieguia.Visible = True
                    Me.txtnumeroguia.Visible = True
                    Me.cmdImprimirGuia.Visible = False
                    Me.cmdGrabarGuia.Visible = True
                    Exit Sub
        End If
    End If
End If
End If
End Sub

Private Sub chk_OrdenCompra_Click()

If Me.chk_OrdenCompra.Value = 1 Then
    Me.txtOrdenCompra.Visible = True
Else
    Me.txtOrdenCompra.Visible = False
End If


End Sub

Private Sub chk_seleccionar_guia_Click()
If Me.chk_seleccionar_guia.Value = 1 Then
   Me.chk_nueva_guia.Value = 0
  
   Call load_guia(Me.txtidVenta.Text)
   Me.txtguia_referencia.Visible = True
Else
  
    Me.DtcGuia.Visible = False
    Me.txtguia_referencia.Visible = False
End If
End Sub
Private Sub load_guia(ByVal in_venta As String)
On Error GoTo salir


Me.txtguia_referencia.Visible = True
strCadena = "SELECT id_transferencia as Codigo, CONCAT('GUIA:',serie,'-',numero) as Descripcion FROM movimiento_transferencia WHERE anulado='no' and  ruc='" & KEY_RUC & "' ORDER BY id_transferencia DESC LIMIT 5"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcGuia)
Me.DtcGuia.Visible = True
Call Resalta(Me.txtguia_referencia)

Exit Sub
salir:


End Sub



Private Sub chk_venta_diferida_Click()

  If Val(Me.txtidVenta.Text) = 0 And Me.DtcTipoDoc.BoundText = "0099" Then
   If KEY_CARGO = "00052" Or KEY_CARGO = "00004" Then
        
        MsgBox "Usted No esta Autorizado Hacer VENTA DIFERIDA" + Chr(13) + Chr(13) + "COORDINE CON LOGISTICA", vbInformation, KEY_VENDEDOR
        Procedencia = diferida
        frmsegurity.Show
        Call disabled_form(Me)
        Exit Sub
   End If
 End If

End Sub
Private Sub load_almacen_entrega()
'strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "'"
'Call ConfiguraRstT(strCadena)
'Call LlenaDataComboT(Me.DtcAlmacen_entrega)
End Sub
Private Sub chkConsultar_Click()
If Me.DtcAlmacen.Enabled = True Then

If (Me.chkconsultar.Value = 1) Then
    Dim in_serie As String
    in_serie = Trim(Me.DtcSerieDoc.BoundText)
    Call get_serie_comprobante_all(Me.DtcSerieDoc, Me.DtcTipoDoc.BoundText)
    If in_serie <> "" Then
        Me.DtcSerieDoc.BoundText = in_serie
    End If
    Me.DtcSerieDoc.Locked = False
    Me.TxtNumeroDoc.Locked = False
    Call Resalta(Me.TxtNumeroDoc)

Else
   Call get_serie_comprobante_alm(Me.DtcSerieDoc, Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText)
   Me.TxtNumeroDoc.Locked = True
End If
End If
End Sub

Private Sub chkDelivery_Click()
If Me.chkDelivery.Value = 1 Then
    
    delivery = "si"
    
    Call Resalta(Me.TxtCliente)
    'Call Resalta(Me.TxtMontoPagado)
Else
    
    delivery = "no"
    
End If
End Sub

Private Sub ChkExtraer_Click()
On Error GoTo salir

If Me.ChkExtraer.Value = 1 Then
    Me.ShaGuia.Width = 7095
     strCadena = "SELECT DISTINCT A.id_doc as Codigo, C.doc_des as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' AND A.venta='si' ORDER BY doc_abrev"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    Call LlenaDataCombo(Me.DtcComprobanteGuia)
    Me.DtcComprobanteGuia.BoundText = KEY_COMPROBANTE
  End If
    Me.DtcComprobanteGuia.Visible = True
    Me.TxtSeri_guia.Visible = True
    Me.TxtNumero_guia.Visible = True
    Me.DtcComprobanteGuia.SetFocus
    
Else
    Me.ShaGuia.Width = 1455
    Me.DtcComprobanteGuia.Visible = False
    Me.TxtSeri_guia.Visible = False
    Me.TxtNumero_guia.Visible = False
    Referencia = False
End If


Exit Sub
salir:

End Sub



Public Sub activar()
On Error GoTo reintentar
Dim in_doc As String
Me.frmprincipal.Enabled = True
Me.DtcAlmacen.Enabled = True
Me.DtcTipoDoc.Enabled = True

Me.TxtNumeroDoc.Enabled = True
Me.DtcTipoDoc.Enabled = False

If KEY_COMPROBANTES_PROPIOS = "si" Then
    strCadena = "SELECT id_doc,serie,numero,igv FROM almacen_comprobante WHERE  defecto='si' AND  id_alm='" & KEY_VENTANILLA & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
Else
    strCadena = "SELECT id_doc,serie,numero,igv FROM almacen_comprobante WHERE  defecto='si' AND  id_alm='" & KEY_ALM & "'  and ruc='" & KEY_RUC & "' LIMIT 1"
End If
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
in_doc = rst("id_doc")
KEY_APLICA_IGV = rst("igv")
Me.DtcTipoDoc.BoundText = in_doc


  
    
    Call comprobante(Me.DtcTipoDoc.BoundText)
    Call nuevo
    Me.timer_pendientes.Enabled = True
    Exit Sub
End If
    
    
 Exit Sub
reintentar:
    

End Sub
Public Sub put_load_proyecto(ByVal in_recibo As String, ByVal in_dni As String)
Call activar
Me.TxtCodCliente.Text = Trim(in_dni)
Call precionar_cliente



strCadena = "call cursor_put_proyecto_cabecera('" & Val(in_recibo) & "','" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcSerieDoc.BoundText & "','" & in_dni & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & Me.DtcMoneda.BoundText & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcSerieDoc.BoundText, Me.DtcTipoDoc.BoundText, Trim(Me.txtformato_impresion.Text))

End Sub


Public Sub put_load_recibos(ByVal in_dni As String)
Call activar
Me.TxtCodCliente.Text = Trim(in_dni)
Me.lblfactura_masiva.Caption = "si"
Call precionar_cliente


strCadena = "call cursor_put_recibo_cabecera('" & Val(in_recibo) & "','" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcSerieDoc.BoundText & "','" & in_dni & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & Me.DtcMoneda.BoundText & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcSerieDoc.BoundText, Me.DtcTipoDoc.BoundText, Trim(Me.txtformato_impresion.Text))

End Sub

Private Sub chkPrecios_Click()
If Me.chkPrecios.Value = 1 Then
    'Call llena_precios(codigoP, Me.HfPrecios)
    Me.HfPrecios.Visible = True
Else
    Me.HfPrecios.Visible = False
End If
End Sub
Public Sub mostrar_precios()
 Call llena_precios(codigoP, Me.HfPrecios)
    
End Sub



Private Sub cmdactivar_Click()
Dim in_serie As String
Call activar
If KEY_FACTURACION_ELECTRONICA = "si" Then
    in_serie = Mid(Trim(Me.DtcSerieDoc.BoundText), 2, 3)
    If KEY_SERVIDOR_CLOUD = "si" Then
       If KEY_SERVIDOR_KEYFACIL = "si" Then
            in_ruta = "https://api.vitekey.com/keyfact/erp/summary?password=vitekey2018&company_id=" & KEY_TOKEN_CLOUD
       Else
            in_ruta = "http://facturacion.vitekey.com/api/comprobantes/listar_estados?serie=" & in_serie & "&token=" & KEY_TOKEN_CLOUD
       End If
        
    Else
       in_ruta = "http://192.168.1.241:3030/api/comprobantes/listar_estados?serie=" & in_serie & "&token=" & KEY_TOKEN_LOCAL
    End If
    
    
    
    Me.web_estado_sunat.Navigate2 in_ruta
    Me.web_estado_sunat.Visible = True
End If
Call Resalta(Me.TxtCodCliente)
Exit Sub
End Sub

Private Sub cmdactualizarcredito_Click()
Dim in_monto_credito As Double
strCadena = "SELECT monto_credito FROM entidad_empresa WHERE cod_unico='" & Trim(Me.TxtCodCliente.Text) & "' and id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    in_monto_credito = rst("monto_credito")
Else
    in_monto_credito = 0
End If

in_monto_credito = in_monto_credito + Val(Me.txtmontocredito.Text)
strCadena = "UPDATE entidad_empresa SET id_credito='si',monto_credito='" & in_monto_credito & "' WHERE cod_unico='" & Trim(Me.TxtCodCliente.Text) & "' and id_empresa='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

Me.frmcredito.Visible = False
Call Me.precionar_cliente
End Sub

Private Sub cmdagregar_Click()


Call agregar_item_grilla


End Sub
Public Sub agregar_item_grilla()
'If control_stock(Trim(Me.TxtCodProducto.Text), Val(Me.txtCantidad.Text)) = True Then
    Call insertar_codigo
'End If
End Sub
Public Sub insertar_codigo()

If Me.DtcTipoDoc.BoundText = "0099" And KEY_UPDATE_PROFORM = "si" And Val(Me.txtidVenta.Text) > 0 And Me.cmdModificar.Enabled = False Then
        If Me.DtcTipoDoc.BoundText = "0099" And KEY_UPDATE_PROFORM = "si" And Val(Me.txtidVenta.Text) > 0 And Me.cmdModificar.Enabled = False Then
           Call Me.insertar_item_save
        End If
Else
    If control_stock(Trim(codigoP), Val(Me.txtCantidad.Text)) = True Then
        Call Me.insertar_item
    End If
End If

End Sub
Private Sub cmdAgregarA_Click()

End Sub


Private Sub cmdAnular_Click()



If MsgBox(MSGANULAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
               Procedencia = anular
               FrmSeguridad.Show
               Exit Sub
        End If
End Sub

Private Sub CmdBuscarFecha_Click()
 Call llenarGrid_Facturas_all(Me.HfFacturas, Trim(Me.TxtCodCliente.Text))
End Sub

Private Sub cmdCerrar_Click()
Me.FrameSerieModelo.Visible = False
End Sub

Private Sub cmdcerrars_Click()
Me.fraApp.Visible = False
End Sub

Private Sub cmdcerrarbanco_Click()
Me.frmbanco.Visible = False
End Sub

Private Sub cmdcerrardireccion_Click()
Me.frmdireccion.Visible = False
End Sub

Private Sub cmdcerraredicion_Click()
Me.txteditable.Text = "no"

End Sub



Private Sub cmdCerrarEntrega_Click(Index As Integer)
If Index = 1 Then
    Me.frmalmacen_entrega(1).Visible = False
End If
End Sub

Private Sub cmdcerrarimpresora_Click()
Me.frmimpresoras.Visible = False
End Sub

Private Sub cmdcerrarmotivonota_Click()
Me.frm_motivo_nota.Visible = False
End Sub

Private Sub cmdCerrarpantalla_Click()
Procedencia = Neutro
Unload Me
End Sub

Private Sub cmdConstancia_Click()

strCadena = "select p.`nombre_completo` as cliente, p.`dni`, CONCAT(p.`direccion`,'-',funct_ubigueo(p.id_departamento,p.id_provincia,p.id_distrito)) as direccion,  v.`numero`, v.`fecha_emision`, " & _
"d.`anio_modelo`, d.`nro_chasis`, d.`serie`, " & _
"pr.`nombre_prod`, pr.`marca`, i.`descripcion` as color, c.`descripcion` as nom_marca , " & _
"yo.dni as ruc, yo.`nombre_completo` as miempresa,  ss.descripcion as modelo " & _
" from `movimiento_venta_detalle` d , `movimiento_venta` v , " & _
"persona p, producto pr, `marca` c, `imp_color` i, persona yo, linea_sub ss " & _
"where v.`id_cliente` = p.`dni` and v.`id_venta` = d.`id_venta` and " & _
"pr.`id_producto` = d.`id_producto` and pr.`ruc` = v.`ruc` and pr.id_sublinea = ss.id_tipo and ss.id_usu = v.ruc and " & _
"pr.`id_marca` = c.`id_marca` and c.`id_usu` = v.`ruc` " & _
"and i.`id_color` = pr.`id_color` and v.`ruc` = yo.`dni` " & _
"and v.`id_venta` = '" & Val(Me.txtidVenta.Text) & "'"

  Call ConfiguraRstK(strCadena)
      
  Ans = ShowMultiReport(rstK, "CorConstanciaVenta", , App.Path + "\Reportes\")

End Sub

Private Sub cmdConsultar_Click()

End Sub

Private Sub cmdconsultaringreso_Click()
Call activar
Exit Sub
End Sub

Private Sub cmdConversion_Click()
Call enabled_form(Me)
Procedencia = transformaciones
frmventa_conversion.Show
frmventa_conversion.lblIdtemporal.Caption = Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)

Exit Sub
End Sub

Private Sub cmdcopropietario_Click()
Me.frmcopropietario.Visible = True
End Sub

Private Sub cmdCredito_Click()

'strCadena = "SELECT id_detalle_venta,p.nombre_prod FROM movimiento_venta_detalle d,producto p WHERE  d.id_producto=p.id_producto and d.ruc=p.ruc and d.ruc='" & KEY_RUC & "' ORDER BY id_detalle_venta DESC "
'Call ConfiguraRst(strCadena)
'If rst.RecordCount > 0 Then
 '   rst.MoveFirst
 '   For i = 0 To rst.RecordCount - 1
  '      strCadena = "UPDATE movimiento_venta_detalle SET detalle='" & rst("nombre_prod") & "' WHERE id_detalle_venta='" & rst("id_detalle_venta") & "'"
   '     CnBd.Execute (strCadena)
   ''
    '    rst.MoveNext
     '   DoEvents
   ' Next i
'End If
Procedencia = modificar_credito
Call disabled_form(Me)
FrmSeguridad.Show

Exit Sub
End Sub

Private Sub cmdCuotas_Click()
FrmVentasCuotas.Show
Call FrmVentasCuotas.llenarGrid_cuotas(Val(Me.txtidVenta.Text), FrmVentasCuotas.HfdDetalle)

End Sub

Private Sub cmdDeclaracion_Click()
  
'strCadena = "SELECT ncliente,id_cliente,marca,nro_chasis,nro_motor FROM view_declaracion WHERE id_venta='" & Val(Me.TxtIdVenta.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
'Call ConfiguraRstK(strCadena)
'Ans = ShowMultiReport(rstK, "CorDeclaracion", , App.Path + "\Reportes\")

strCadena = "SELECT nombre_completo,id_cliente,id_copropietario,coopropietario,documento,cheque,fecha_emision,id_forma_pago FROM view_forma_pago WHERE id_venta='" & Val(Me.txtidVenta.Text) & "'"
 
 'strCadena = "select p.`nombre_completo` as cliente , p.`dni`, m.`serie`, m.`numero`,v.`doc_des` as documento,v.`doc_des` as documento, " & _
" v.`doc_des` as documento, m.`fecha_emision` " & _
"from persona p, movimiento_venta m, `comprobantes` v " & _
"where p.`dni` = m.`id_cliente` " & _
"and m.`id_doc` = v.`id_doc` " & _
"and m.`id_venta` = '" & Val(Me.TxtIdVenta.Text) & "'"
 
 
  Call ConfiguraRstK(strCadena)
  strCadena = "SELECT id_venta,descripcion,monto_caja FROM view_forma_pago_detalle WHERE id_venta='" & Val(Me.txtidVenta.Text) & "'"
  Call ConfiguraRstL(strCadena)
  
  strCadena = "SELECT id_venta,nro_chasis,serie,marca,modelo,anio_fabricacion,color,nro_dua,nro_item FROM view_datos_chasis_motor WHERE id_venta='" & Val(Me.txtidVenta.Text) & "'"
  Call ConfiguraRstP(strCadena)
 Ans = ShowMultiReport(rstK, "CorDeclaracion", , App.Path + "\Reportes\", , , , , rstL, "rpt_declaracion_detalle_pago", rstP, "rpt_datos_vehiculo")
 
 'Ans = ShowMultiReport(rstK, "CorDeclaracion", , App.Path + "\Reportes\")
End Sub

Private Sub cmddelivery_Click()

End Sub
Public Sub llena_precios(ByVal id_producto As String, ByVal Grilla As MSHFlexGrid)
Dim porcentaje As Single
strCadena = "SELECT * FROM almacen_producto_precio  WHERE id_producto='" & id_producto & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY precio DESC"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Me.chkPrecios.Value = 0
    Me.HfPrecios.Visible = False
    Grilla.Rows = 0
    
    Exit Sub
End If
   Grilla.Visible = True
   Grilla.Rows = 0
   
       ReDim arrColWidth(1 To rstT.Fields.Count)
       
        For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 300
           Grilla.ColWidth(1) = 850
           Grilla.ColWidth(2) = 1050
           Grilla.ColWidth(3) = 0
        Next
        cabecera = "" & vbTab & "PRECIO" & vbTab & "CANTIDADES" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        c = 0
            NumeroCampo = 0
            
        For i = 0 To rstT.RecordCount - 1
          estado = Chr(168)
          descripcion = ""
          descripcion = "   [ " & rstT("cant_ini") & Space(1) & "-" & Space(1) & rstT("cant_fin") & " ]"
          Fila = estado & vbTab & Format(rstT("precio"), "#,##0.00") & vbTab & descripcion & vbTab & rstT("cant_ini")
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
          
          rstT.MoveNext
      Next i
    
End Sub




Public Sub buscar_pendientes()
On Error GoTo salir

strCadena = "SELECT funct_proformas_pendientes('" & KEY_ALM & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & KEY_RUC & "')"
Call ConfiguraRstI(strCadena)
If rstI(0) > 0 Then
    PlaySound App.Path & "\sonidos\dingding.wav"
    If rstI(0) <> Val(Me.txtnumeropendientes.Text) Then
        Call llenar_pendientes(Me.HfPendientes)
    End If
End If
Exit Sub
salir:

End Sub
Public Sub llenar_pendientes(ByVal Grilla As MSHFlexGrid)
Dim porcentaje As Single
Dim ndocumento() As String
strCadena = "SELECT id_venta,ncliente,documento,total FROM view_listado_pendientes_prod WHERE pendiente='si' and fecha_emision='" & KEY_FECHA & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRstI(strCadena)
Me.txtnumeropendientes.Text = rstI.RecordCount
If rstI.RecordCount < 1 Then
    
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 1700
           Grilla.ColWidth(3) = 700
           Grilla.ColWidth(4) = 400
           
        Next
        cabecera = "IDVENTA" & vbTab & "PROFORMA" & vbTab & "CLIENTE" & vbTab & "MONTO" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 1 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
        c = 4
        NumeroCampo = 4
            
        For i = 0 To rstI.RecordCount - 1
          estado = Chr(168)
          descripcion = ""
            ndocumento = Split(rstI("documento"), ":")
            nproforma = "P:" & ndocumento(1)
            
          Fila = rstI("id_venta") & vbTab & nproforma & vbTab & rstI("ncliente") & vbTab & Format(rstI("total"), "#,##0.00") & vbTab & estado
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
    
End Sub

Public Sub llenar_bancos(ByVal Grilla As MSHFlexGrid)
Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
   Grilla.Rows = 0
   Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 400
        Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
        c = 2
        NumeroCampo = 2
            
        For i = 0 To rstI.RecordCount - 1
          estado = Chr(168)
          descripcion = ""
            
            
          Fila = rstI("Codigo") & vbTab & rstI("Descripcion") & vbTab & estado
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
    
End Sub

Private Sub cmdEnviar_Click()
strCadena = "select d.`nro_chasis`, d.`serie` , s.`descripcion` as modelo, " & _
"m.`descripcion` as `marca`, CONCAT( pr.`nombres`, ' ', pr.`a_paterno`, ' ', pr.`a_materno`) as paciente," & _
"pr.`direccion` , pr.`dni`, v.`fecha_emision` " & _
"from movimiento_venta v , `movimiento_venta_detalle` d, producto p, " & _
"`linea_sub` s, marca m, persona pr " & _
"where v.`id_venta` = d.`id_venta` and d.`id_producto` = p.`id_producto` and v.`ruc` = p.`ruc` " & _
"and p.`id_sublinea` = s.`id_tipo` and v.ruc = s.`id_usu` " & _
"and p.`id_marca` = m.`id_marca` and v.`ruc` = m.`id_usu` " & _
"and v.`id_cliente` = pr.`dni` and v.`id_venta` ='" & Val(Me.txtidVenta.Text) & "'"

Call ConfiguraRstK(strCadena)
strCadena = "select dni, nombre from  imp_reporte_app where `id_venta` = " & Val(Me.txtidVenta.Text)
Call ConfiguraRstL(strCadena)
      
  'Ans = ShowMultiReport(rstK, "CorApp", , App.Path + "\Reportes\", , , , , rstL, "CorAppNombres")
  'Ans = ShowMultiReport(rstK, "CorConstanciaVenta", , App.Path + "\Reportes\")
  Ans = ShowMultiReport(rstK, "CorApp", , App.Path + "\Reportes\")
'Ans = ShowMultiReport(rstK, "CorApp", , App.Path + "\Reportes\")
   
  
End Sub

Private Sub cmddescartar_Click()
On Error GoTo salir
If MsgBox("Esta seguro de Quitar este documento", vbQuestion + vbYesNo) = vbYes Then
   strCadena = "UPDATE movimiento_venta set pendiente='no' WHERE id_venta='" & Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) & "' and ruc"
   CnBd.Execute (strCadena)
    
    
   Me.HfPendientes.RemoveItem (Me.HfPendientes.Row)
   Me.txtnumeropendientes.Text = Val(Me.txtnumeropendientes.Text) - 1
End If
salir:
Exit Sub
End Sub





Private Sub cmdEliminar_Click()
 If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Procedencia = Eliminar
        FrmSeguridad.Show
End If
End Sub

Private Sub cmdenviarimpresion_Click()
    Dim in_transferencia As String
    Call Establecer_Impresora_predeterminada(Trim(Me.HfImpresoras.TextMatrix(Me.HfImpresoras.Row, 1)))
    
    strCadena = "SELECT  * FROM movimiento_transferencia WHERE id_doc='0009' and serie='" & Trim(Me.txtserieguia.Text) & "' and numero='" & Trim(txtnumeroguia.Text) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       in_transferencia = rst("id_transferencia")
    End If
    
                strCadena = "SELECT id_transferencia,fecha,fecha_traslado,id_doc,comprobante,id_remitente,remitente,id_destinatario,destinatario,direccion_origen,direccion_destino,ubigeo,dni_atencion,atencion,almacen_origen,id_alm_destino,id_transporte,transporte,marca_placa,placa,mtc,certificado,id_chofer,chofer,licencia,id_motivo,peso_total,valor_mercaderia,observacion2,observacion3,venta,ruc FROM view_tranferencia_cabecera WHERE id_transferencia='" & Val(in_transferencia) & "' and ruc='" & KEY_RUC & "'"
                Call ConfiguraRst(strCadena)
                If rst.RecordCount > 0 Then
                    strCadena = "SELECT * FROM view_transferencia_detallado WHERE id_transferencia='" & Val(in_transferencia) & "'"
                    Call ConfiguraRstK(strCadena)
                   
                    Ans = ShowMultiReport(rst, "rpt_guia_remision", , App.Path + "\Reportes\", , , , , rstK, "rptguia_detalle")
                End If
                
    
                Exit Sub
                
       
    
    Set Printer = Printers(Val(Me.HfImpresoras.TextMatrix(Me.HfImpresoras.Row, 0)) - 1)
    Printer.Font.name = "Draft 17cpi"
   ' Call Orden_Impresion("0009", Trim(Me.txtserieguia.Text), Trim(Me.txtnumeroguia.Text), Trim(Me.txttipofactura.Text), Val(Me.TxtIdVenta.Text), Trim(Me.txtdireccion.Text))
    Call impresion_formato(rst("id_formato_impresion"), "0009", Trim(Me.txtserieguia.Text), Trim(Me.txtnumeroguia.Text), Trim(Me.txttipofactura.Text), Trim(Me.txtDireccion.Text))
    
    Me.frmimpresoras.Visible = False

End Sub

Private Sub cmdgenerarLetras_Click()
        
        Me.frmvencimiento.Visible = False
        Me.PanelCredito.Visible = True
        Me.TxtCuotas.Visible = True
        Me.TxtCuotas.Text = "1"
        Call Resalta(Me.TxtCuotas)
        


Exit Sub
End Sub

Private Sub cmdgenerarmantenimientos_Click()
'Call generar_mantenimientos(Val(Me.TxtIdVenta.Text), Trim(Me.TxtCodCliente.Text), KEY_FECHA)

Call generar_carta_poder(Me.txtidVenta.Text)
End Sub
Private Sub generar_carta_poder(ByVal in_venta As String)
Dim in_dni As String
Dim in_persona As String
Dim in_direccion As String
in_dni = ""
in_persona = ""
in_direccion = ""


strCadena = "select p.`nombre_completo` as cliente, p.`dni`, CONCAT(p.`direccion`,'-',funct_ubigueo(p.id_departamento,p.id_provincia,p.id_distrito)) as direccion,  v.`numero`, v.`fecha_emision`, " & _
"d.`anio_modelo`, d.`nro_chasis`, d.`serie`, " & _
"pr.`nombre_prod`, pr.`marca`, i.`descripcion` as color, c.`descripcion` as nom_marca , " & _
"yo.dni as ruc, yo.`nombre_completo` as miempresa,  ss.descripcion as modelo " & _
" from `movimiento_venta_detalle` d , `movimiento_venta` v , " & _
"persona p, producto pr, `marca` c, `imp_color` i, persona yo, linea_sub ss " & _
"where v.`id_cliente` = p.`dni` and v.`id_venta` = d.`id_venta` and " & _
"pr.`id_producto` = d.`id_producto` and pr.`ruc` = v.`ruc` and pr.id_sublinea = ss.id_tipo and ss.id_usu = v.ruc and " & _
"pr.`id_marca` = c.`id_marca` and c.`id_usu` = v.`ruc` " & _
"and i.`id_color` = pr.`id_color` and v.`ruc` = yo.`dni` " & _
"and v.`id_venta` = '" & Val(Me.txtidVenta.Text) & "'"

  Call ConfiguraRstK(strCadena)
      
  Ans = ShowMultiReport(rstK, "CorCartaPoder", , App.Path + "\Reportes\")
  
  Exit Sub
  


If Len(Trim(Me.TxtCodCliente.Text)) = 11 And Mid(Trim(Me.TxtCodCliente.Text), 1, 2) <> "10" Then
     strCadena = "SELECT p.dni,p.nombre_completo,p.direccion FROM persona_accidentes a,persona p where a.dni='" & Trim(Me.TxtCodCliente.Text) & "' and a.dni_familia=p.dni LIMIT 1"
     Call ConfiguraRstK(strCadena)
     If rstK.RecordCount > 0 Then
         strCadena = "SELECT direccion FROM persona where dni='" & Trim(Me.TxtCodCliente.Text) & "'"
         Call ConfiguraRstL(strCadena)
         If rstL.RecordCount > 0 Then
            in_direccion = rstL("direccion")
        End If
         strCadena = "SELECT '" & rstK("dni") & "','" & rstK("nombre_completo") & "','" & in_direccion & "'" & " ,d.nro_chasis as serie FROM movimiento_venta v,movimiento_venta_detalle d WHERE   v.id_venta=d.id_venta and  v.id_venta='" & Val(in_venta) & "' and v.ruc='" & KEY_RUC & "'"
     End If
     
Else
       strCadena = "SELECT id_cliente,ncliente,direccion,d.nro_chasis as serie FROM movimiento_venta v,movimiento_venta_detalle d WHERE   v.id_venta=d.id_venta and  v.id_venta='" & Val(in_venta) & "' and v.ruc='" & KEY_RUC & "'"
End If
  
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "cartapoder", , App.Path + "\Reportes\")






strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   If Len(rstK("id_copropietario")) > 1 Then
      strCadena = "SELECT '" & rstK("id_copropietario") & "','" & get_persona(rstK("id_copropietario")) & "','" & get_direccion(rstK("id_copropietario")) & "'" & " ,d.nro_chasis as serie FROM movimiento_venta v,movimiento_venta_detalle d WHERE   v.id_venta=d.id_venta and  v.id_venta='" & Val(in_venta) & "' and v.ruc='" & KEY_RUC & "'"
      Call ConfiguraRst(strCadena)
      Ans = ShowMultiReport(rst, "cartapoder", , App.Path + "\Reportes\")
   End If
End If




End Sub
Private Sub cmdGrabarGuia_Click()
    strCadena = "SELECT count(*) FROM  movimiento_transferencia WHERE id_doc='0009' and serie='" & Trim(Me.txtserieguia.Text) & "' and numero='" & Trim(Me.txtnumeroguia.Text) & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstZ(strCadena)
                    If rstZ(0) > 0 Then
                        MsgBox "Guia  ya generada verifique su correlativo", vbInformation, KEY_EMPRESA
                        Call Resalta(Me.txtnumeroguia)
                        Exit Sub
                    Else
                        Call Llenar_Temporal_transferencias(Val(Me.txtidVenta.Text))
                    End If
                    
    
End Sub
Private Sub Llenar_Temporal_transferencias(ByVal idVenta As Double)
Dim total_temp As Double
Dim rstTemporal As New ADODB.Recordset
Dim rstDetalle As New ADODB.Recordset
Dim i As Integer

strCadena = "SELECT * FROM movimiento_venta_detalle D WHERE id_venta='" & idVenta & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then

strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "' AND id_doc='0009' "
CnBd.Execute (strCadena)

strCadena = "DELETE FROM movimiento_transferencia_series WHERE ruc='" & KEY_RUC & "'  AND id_doc='0009' and serie='" & Trim(Me.txtserieguia.Text) & "' and numero='" & Trim(Me.txtnumeroguia.Text) & "' "
CnBd.Execute (strCadena)

total_temp = 0
rst.MoveFirst

For i = 0 To rst.RecordCount - 1
    strCadena = "INSERT INTO movimiento_transferencia_temporal(id_doc,serie,numero,id_producto,detalle,cantidad,peso,total,dni_save,ruc) VALUES " & _
    "('0009','" & Trim(Me.txtserieguia.Text) & "','" & Trim(Me.txtnumeroguia.Text) & "','" & rst("id_producto") & "','" & rst("detalle") & "','" & rst("cantidad") & "','" & rst("peso") & "'," & _
    "'" & Val(rst("peso")) * Val(rst("cantidad")) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
     
    strCadena = "INSERT INTO movimiento_transferencia_series(id_doc,serie,numero,id_producto,chasis,motor,anio_fabricacion,nro_dua,nro_item,ruc)  VALUES " & _
    "('0009','" & Trim(Me.txtserieguia.Text) & "','" & Trim(Me.txtnumeroguia.Text) & "','" & rst("id_producto") & "','" & rst("nro_chasis") & "','" & rst("serie") & "','" & rst("anio_fabricacion") & "','" & rst("nro_dua") & "','" & rst("nro_item") & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)

     
    rst.MoveNext
Next i
   Call save_guia
   
End If

 

End Sub
Private Sub savedetalle(ByVal id_transferencia As Double, ByVal nfinalizado As String)
    strCadena = "SELECT * FROM movimiento_transferencia_temporal WHERE (numero='" & Trim(Me.txtnumeroguia.Text) & "' AND id_doc='0009' AND serie='" & Trim(Me.txtserieguia.Text) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "')"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       For i = 0 To rstT.RecordCount - 1
            If nfinalizado = "si" Then
                strCadena = "DELETE FROM movimiento_transferencia_detalle WHERE id_producto='" & rstT("id_producto") & "' and id_transferencia='" & id_transferencia & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                
                 
            End If
           strCadena = "INSERT INTO movimiento_transferencia_detalle(id_transferencia,id_producto,detalle,cantidad,recibido,peso,total,ruc) VALUES ('" & id_transferencia & "','" & rstT("id_producto") & "','" & rstT("detalle") & "','" & rstT("cantidad") & "','" & rstT("cantidad") & "','" & rstT("peso") & "','" & rstT("total") & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           rstT.MoveNext
        Next i
        strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE numero='" & Trim(Me.txtnumeroguia.Text) & "' AND id_doc='0009' AND serie='" & Trim(Me.txtserieguia.Text) & "'AND  dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
         
    End If
End Sub

Private Sub save_guia()
        
        strCadena = "INSERT INTO movimiento_transferencia(id_doc,id_tipo_guia,serie,numero,fecha,fecha_traslado,id_destinatario,destinatario,direccion,direccion_destino,ubigeo_destino,id_alm_origen,id_alm_destino,id_motivo,motivo_otros,observacion,id_venta,dni_save,ruc) " & _
        "VALUES('0009','" & Trim(Me.txttipofactura.Text) & "','" & Trim(Me.txtserieguia.Text) & "','" & Trim(Me.txtnumeroguia.Text) & "','" & KEY_FECHA & "','" & KEY_FECHA & "','" & Trim(Me.TxtCodCliente.Text) & "','" & Trim(Me.TxtCliente.Text) & "'," & _
        "'" & KEY_DIRECCION_ALM & "','" & Trim(Me.txtDireccion.Text) & "','" & get_ubigeo_sunat_persona(Me.TxtCodCliente.Text) & "','" & KEY_ALM & "','" & KEY_ALM & "','1','','" & Trim(Me.txtObservacion.Text) & "','" & Val(Me.txtidVenta.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        id_transferencia = LastRegistroRUC("movimiento_transferencia", "id_transferencia")
        
         
        
        
        strCadena = "UPDATE movimiento_transferencia_series SET id_transferencia='" & id_transferencia & "' WHERE id_doc='0009' and serie='" & Trim(Me.txtserieguia.Text) & "' and numero='" & Trim(Me.txtnumeroguia.Text) & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
         
        
        Call savedetalle(id_transferencia, "no")
        StrNumero = FormatosCeros(Trim(str(Val(Me.txtnumeroguia.Text)) + 1), 6)
        strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE id_alm='" & KEY_ALM & "' AND id_doc='0009' AND serie='" & Trim(Me.txtserieguia.Text) & "' AND ruc='" & Trim(KEY_RUC) & "'"
        CnBd.Execute (strCadena)
        
         
        
        Me.cmdImprimirGuia.Visible = True
        Me.cmdGrabarGuia.Visible = False
        Me.txtserieguia.Locked = True
        Me.txtnumeroguia.Locked = True
End Sub

Private Sub cmdGuiaRemision_Click()


End Sub

Private Sub cmdHistorial_Click()

End Sub

Private Sub cmdImprimir_Click()

'Call printer_barcode("Hola")
'    Exit Sub
    

'Call generar_codificacion
       ' If strEspecial > 1 Then
       '     Call OrdenImpresionEspecial
       ' Else
           If Me.chkimprimira4.Value = 1 Then
              Call put_impresion_a4(Me.txtidVenta.Text, Val(Me.lblTotal.Caption), Me.DtcMoneda.BoundText)
              Exit Sub
           End If
           
           
           
           
           Call OrdenImpresion(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text))
           
           
       ' End If
End Sub
Private Sub generar_codificacion()

'Dim comprobante As String
'comprobante = "http://facturacion.vitekey.com/api/comprobantes/pdf/" & Trim(Me.txt_sunat_key.Text) & ".pdf"
'Pic_firma.Picture = cQrCode.GetPictureQrCode("123", 200, 200)
'If Picture1.Picture Is Nothing Then MsgBox "Error!"
'Picture1.Picture = cQrCode.GetPictureQrCode(Text1.Text, 200, 200, "UTF-8", "L", vbRed, vbBlue, 3)
   
   
End Sub
Private Sub cmdImprimirGuia_Click()

Call llenar_impresoras(Me.HfImpresoras)

End Sub

Private Sub put_editable()
  
End Sub

Public Sub facturar()
    
    
    If Val(Me.txtidVenta.Text) > 0 Then
        If KEY_FACTURACION_ELECTRONICA = "si" Then
                If get_firma_online(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText)) = "si" Then
                   If KEY_SERVIDOR_KEYFACIL = "si" Then
                        Call firma_electronica_eliminar
                   Else
                        Call firma_electronica_eliminar_save
                   End If
                   
                   Call firma_electronica
                   
                   Exit Sub
                End If
           End If
        Exit Sub
    End If
    
    
    
    If Me.chkconsultar.Value = 1 And KEY_USUARIO <> "42546269" Then
       MsgBox "Para PROCESAR un comprobante." + Chr(13) + Chr(13) + "Debe Desactivar MODO CONSULTA.", vbInformation, KEY_VENDEDOR
       Exit Sub
    End If
    
    
    If Val(Me.txtidVenta.Text) > 0 Then
        If KEY_FACTURACION_ELECTRONICA = "si" Then
                If get_firma_online(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText)) = "si" Then
                   Call firma_electronica_eliminar_save
                   Call firma_electronica
                   Exit Sub
                End If
           End If
        Exit Sub
    End If
    
    
   
    
   
    
    
    
    If Me.DtcSerieDoc.BoundText = "" Then
        MsgBox "LA " & Me.DtcTipoDoc.Text & " debe tener una SERIE valida" + Chr(13) + Chr(13) + "Configure su comprobante.", vbInformation
        Me.DtcTipoDoc.SetFocus
        Exit Sub
    End If
    
    
   ' If Trim(Me.DtcTipoDoc.BoundText) = "0001" And Len(Me.TxtCodCliente.Text) <> 11 Then
   '     MsgBox "INGRESE RUC VALIDO PARA EL CLIENTE", vbInformation, KEY_EMPRESA
   '     Call Resalta(Me.TxtCodCliente)
   '     Exit Sub
   ' End If

    
    If Trim(Me.DtcVendedor.Text) = "" Then
        MsgBox "DEBE SELECCIONAR UN VENDEDOR PARA ESTA VENTA." + Chr(13) + Chr(13) + "GRACIAS : " & KEY_VENDEDOR, vbInformation
        Exit Sub
    End If

    
    If Trim(Me.DtcTipoDoc.BoundText) = "0007" Then ' NOTA DE CREDITO
       
       If Mid(Me.DtcSerieDoc.BoundText, 1, 1) = "F" And Len(Trim(Me.TxtCodCliente.Text)) <> 11 Then
          MsgBox "INGRESE UN RUC VALIDO", vbInformation, KEY_VENDEDOR
          Call Resalta(Me.TxtCodCliente)
          Exit Sub
       End If
       
       If Mid(Me.DtcSerieDoc.BoundText, 1, 1) = "B" And Trim(Me.TxtCodCliente.Text) = "00000000" Then
            
          
         If MsgBox("INGRESE UN DNI VALIDO." + Chr(13) + "NOTA DE CREDITO [OBLIGATORIO DNI]" + Chr(13) + Chr(13) + "DESEA REALIZAR DE TODOS MODOS ?", vbInformation + vbYesNo, KEY_VENDEDOR) = vbYes Then
         
         Else
            Call Resalta(Me.TxtCodCliente)
            Exit Sub
         End If
        End If
       
       
       
    End If
     
  Call put_editable
    
    
    If Trim(Me.txtafectacaja.Text) = "no" Then
          '  FrmVentas.Enabled = False
          '  Procedencia = seleccionar_vendedor
          '  FrmSeguridad.Show
          '  Exit Sub
    End If
        
      
      Call procesar_comprobante
      Exit Sub
    
    
    
    
  


   
End Sub
Public Sub procesar_comprobante()
    strCadena = "SELECT * FROM movimiento_venta_monto_temporal WHERE id_usuario='" & KEY_USUARIO & "' AND id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If Me.DtcTipoDoc.BoundText = "0007" Then
        GoTo grabarnota
    End If
    
    If (Val(Me.lblPago.Caption) >= 0 And rst.RecordCount > 0) Then
grabarnota:
        
        
        
        If Save = True Then
           If Me.chk_manual.Value = 1 Then
               Call nuevo
               Exit Sub
           End If
           
           If KEY_FACTURACION_ELECTRONICA = "si" Then
                If get_firma_online(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText)) = "si" Then
                   Call firma_electronica
                   Call get_pendiente
                   Exit Sub
                End If
           End If
                
        If Me.DtcTipoDoc.BoundText <> "0099" Then
            
            If KEY_RUC = "20487473881" And Me.DtcTipoDoc.BoundText = "0054" Then
                Me.cmdProcesar.Enabled = False
                Me.cmdImprimir.Enabled = True
                Exit Sub
            Else
            Call OrdenImpresion(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text))
            End If
            
            
            
            Me.cmdProcesar.Enabled = False
            Me.cmdImprimir.Enabled = True
        Else
                Me.cmdProcesar.Enabled = False
                Me.cmdImprimir.Enabled = True
                Exit Sub
        End If
        
        If Val(Me.txt_id_pendiente.Text) > 0 And (KEY_CARGO = "00008" Or KEY_CARGO = "00004") Then
           Call get_pendiente
        End If
        
        If KEY_COMPROBANTE_ADICIONAL = "si" Then
            Call enabled_form(Me)
            Exit Sub
        End If
            Call nuevo
            Exit Sub
        End If
        
    Else
        MsgBox "INGRESE UN MONTO VALIDO", vbExclamation, "Mensaje para la Cajera"
        If Me.TxtMontoPagado.Visible = True Then
            Call Resalta(Me.TxtMontoPagado)
        Else
            Exit Sub
        End If
    End If
End Sub
Private Sub get_pendiente()
 On Error GoTo salir
 strCadena = "UPDATE movimiento_venta SET pendiente='no' WHERE id_venta='" & Val(txt_id_pendiente.Text) & "' AND ruc='" & KEY_RUC & "'"
 CnBd.Execute (strCadena)
 Me.HfPendientes.RemoveItem (Me.HfPendientes.Row)
 Me.txtnumeropendientes.Text = Val(Me.txtnumeropendientes.Text) - 1
 Exit Sub
salir:
End Sub

Private Sub cmdImprimirRecibo_Click()

End Sub

Private Sub cmdListado_Click()
Call disabled_form(Me)
frmventaslistado.Show
Exit Sub
End Sub

Private Sub cmdModificar_Click()
Procedencia = modificar
Call disabled_form(Me)
FrmSeguridad.Show
Exit Sub
End Sub

Private Sub cmdNuevo_Click()
    Call nuevo
    Exit Sub
End Sub

Private Sub cmdokcopropietario_Click()
Me.frmcopropietario.Visible = False
End Sub

Private Sub cmdPdf_Click()
 MsgBox "Para Visualizar el PDF Sunat, debe tener Iniciado Keyfacil", vbInformation

Shell "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --start-maximized --url https://api.vitekey.com/keyfact/companies/" & KEY_TOKEN_CLOUD & "/sales/invoices/" & Trim(Me.txt_hash.Text) & "/pdf?format=A4"

End Sub

Private Sub cmdpedidosoendientes_Click()
frmventas_pedidos.Show
Exit Sub
End Sub




Private Sub cmdProcesar_Click()

Call validar_facturacion




End Sub
Private Function get_nota_credito_user() As Boolean

'strCadena = "SELECT habilitado_nota_credito FROM entidad_empresa WHERE cod_unico='" & KEY_USUARIO & "' and id_empresa='" & KEY_RUC & "' and id_personal='si' LIMIT 1 "
'Call ConfiguraRstP(strCadena)
'If rstP.RecordCount > 0 Then
'   If rstP("habilitado_nota_credito") = "si" Then
'        get_nota_credito_user = True
'   Else
'        get_nota_credito_user = False
'   End If
'Else
'    get_nota_credito_user = False
'End If

If KEY_HABILITADO_NOTACREDITO = "si" Then
   get_nota_credito_user = True
Else
   KEY_HABILITADO_NOTACREDITO = False

End If


End Function

Public Sub validar_facturacion()
Dim in_diferidan As String
Dim in_periodo As String



If get_periodo_cierre(get_periodo_actual(KEY_FECHA), "ventas") = True Then
    MsgBox "PERIODO DE VENTAS. CERRADO...." + Chr(13) + Chr(13) + "CONSULTE CON CONTABILIDAD", vbInformation, KEY_VENDEDOR
    
    Exit Sub
End If



If Trim(Me.TxtCodCliente.Text) = "" Then
   MsgBox "INGRESE, UN DNI/RUC   ........   VALIDO.", vbInformation
   Exit Sub
End If













If KEY_GRIFO = "si" Then
    If Trim(Me.txtObservacion.Text) = "PLACA N�:" Or Trim(Me.txtObservacion) = "" Then
        MsgBox "INGRESE PLACA DE VEHICULO.", vbInformation, KEY_VENDEDOR
        Me.txtObservacion.Text = "PLACA N�:"
        Call Resalta(Me.txtObservacion)
        Exit Sub
    End If
End If




If KEY_NOTA_CREDITO_ADMIN = "si" And Me.DtcTipoDoc.BoundText = "0007" Then
   If get_nota_credito_user = False Then
    'If get_nota_credito_admin = False Then
        MsgBox "USTED NO ESTA AUTORIZADO PARA REALIZAR" + Chr(13) + Chr(13) + "NOTAS DE CREDITO", vbInformation, KEY_VENDEDOR
        Exit Sub
   ' End If
    End If
    
End If

If put_dni_boleta(Trim(Me.TxtCodCliente.Text), Val(Me.lblTotal.Caption), Me.DtcTipoDoc.BoundText) = False Then
   Call Resalta(Me.TxtCodCliente)
  Exit Sub
End If


If Me.DtcTipoDoc.BoundText = "0099" And KEY_UPDATE_PROFORM = "si" And Val(Me.txtidVenta.Text) > 0 And Me.cmdModificar.Enabled = False Then
    If Me.chk_venta_diferida.Value = 1 Then
        in_diferidan = "si"
    Else
        in_diferidan = "no"
    End If


    strCadena = "UPDATE movimiento_venta SET diferida='" & in_diferidan & "', exonerado='" & Val(Me.lblExonerado.Caption) & "', valor_venta='" & Val(Me.LblValorVenta.Caption) & "',igv='" & Val(Me.LblIgv.Caption) & "',total='" & Val(Me.lblTotal.Caption) & "',monto_pago='" & Val(Me.lblTotal.Caption) & "' WHERE id_venta='" & Val(Me.txtidVenta.Text) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
    strCadena = "UPDATE movimiento_venta_monto SET monto='" & Val(Me.lblTotal.Caption) & "',monto_caja='" & Val(Me.lblTotal.Caption) & "' WHERE id_venta='" & Val(Me.txtidVenta.Text) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Me.cmdProcesar.Enabled = False
    Exit Sub
End If


If Me.DtcTipoDoc.BoundText = "0007" Then
    
    
    If Val(Me.txtid_venta_ref.Text) = 0 Then
        strCadena = "SELECT * FROM movimiento_venta WHERE (numero='" & Trim(Me.TxtNumero_guia.Text) & "' AND id_doc='" & Trim(Me.DtcComprobanteGuia.BoundText) & "' AND serie='" & Trim(Me.TxtSeri_guia.Text) & "' AND ruc='" & KEY_RUC & "')"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
            Me.txtid_venta_ref.Text = rstL("id_venta")
        End If
    End If
    
    

    If validar_dni_nota_credito(Trim(Me.TxtCodCliente.Text), Me.txtid_venta_ref.Text) = False Then
        MsgBox "ADVERTENCIA. [  INCONSISTENCIA  ]." + Chr(13) + "DNI/RUC DOCUMENTO RELACIONADO DIFIERE AL ACTUAL", vbInformation, KEY_VENDEDOR
        
        Exit Sub
    End If
End If


If Me.txtServicio.Text = "si" And Len(KEY_CTA_DETRACCION) > 5 And Me.chk_detraccion.Value = 0 Then
    If MsgBox("ESTA SEGURO DE EMITIR SIN DETRACCION" + Chr(13) + "CONFIRME PORFAVOR.", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
End If







If validar_comprobante_electronico(Me.DtcTipoDoc.BoundText, Trim(Me.TxtCodCliente.Text)) = True Then
    Call facturar
End If

End Sub
Private Function validar_dni_nota_credito(ByVal in_dni As String, ByVal in_referencia As String) As Boolean
strCadena = "select * from movimiento_venta WHERE id_cliente='" & Trim(in_dni) & "' and  id_venta='" & Val(in_referencia) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
   validar_dni_nota_credito = False
Else
   validar_dni_nota_credito = True
End If
End Function
Public Function firma_electronica_eliminar()
Call disabled_form(Me)
FrmLoad_web_service.Show
FrmLoad_web_service.nom_prcedimiento = "eliminar_firma_electronica"
Set FrmLoad_web_service.FormPadre = Me
If KEY_SERVIDOR_CLOUD = "si" Then

    If KEY_SERVIDOR_KEYFACIL = "si" Then
        Call FrmLoad_web_service.crear_json_facturacion_electronica("https://api.vitekey.com/keyfact/utils/erp-invoice-delete?password=vitekey2018&company_id=" & KEY_TOKEN_CLOUD & "", "POST", json_facturacion_electronica_eliminar(Format(Val(Me.DtcTipoDoc.BoundText), "00"), Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text)), "{x-api-token: '" & KEY_TOKEN_CLOUD & "', x-api-produccion: 'yes'}")
    Else
        Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/eliminar", "POST", json_facturacion_electronica_eliminar(Format(Val(Me.DtcTipoDoc.BoundText), "00"), Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text)), "{x-api-token: '" & KEY_TOKEN_CLOUD & "', x-api-produccion: 'yes'}")
    End If
    
Else
    Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/eliminar", "POST", json_facturacion_electronica_eliminar(Format(Val(Me.DtcTipoDoc.BoundText), "00"), Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text)), "{x-api-token: '" & KEY_TOKEN_LOCAL & "', x-api-produccion: 'yes'}")
End If

End Function
Public Function firma_electronica_eliminar_save()
Call disabled_form(Me)
FrmLoad_web_service.Show
FrmLoad_web_service.nom_prcedimiento = "eliminar_comprobante_electronico"
Set FrmLoad_web_service.FormPadre = Me
If KEY_SERVIDOR_CLOUD = "si" Then
   Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/eliminar", "POST", json_facturacion_electronica_eliminar(Format(Val(Me.DtcTipoDoc.BoundText), "00"), Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text)), "{x-api-token: '" & KEY_TOKEN_CLOUD & "', x-api-produccion: 'yes'}")
Else
    Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/eliminar", "POST", json_facturacion_electronica_eliminar(Format(Val(Me.DtcTipoDoc.BoundText), "00"), Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text)), "{x-api-token: '" & KEY_TOKEN_LOCAL & "', x-api-produccion: 'yes'}")
End If

End Function

Public Function eliminar_comprobante_electronico() As Boolean
eliminar_comprobante_electronico = True
End Function
Public Sub eliminar_firma_electronica(ByVal strHtml As String)


Call EliminarVentas(Trim(DtcTipoDoc.BoundText), Trim(DtcSerieDoc.BoundText), Trim(TxtNumeroDoc.Text), Trim(DtcAlmacen.BoundText))

                    
End Sub
Private Function firma_electronica()

Call disabled_form(Me)
FrmLoad_web_service.Show
FrmLoad_web_service.nom_prcedimiento = "procesar_firma_electronica"

Set FrmLoad_web_service.FormPadre = Me

Select Case Me.DtcTipoDoc.BoundText
    Case "0003"
         in_tipo_doc = "1"
         If Trim(Me.txtExtranjero.Text) = "si" Then
            in_tipo_doc = "4"
         End If
    
    
    Case "0001"
        in_tipo_doc = "6"
         If Trim(Me.txtExtranjero.Text) = "si" Then
            in_tipo_doc = "0"
         End If
         
    Case "0002"
        in_tipo_doc = "1"
        
End Select
    id_motivo_nota = ""
    motivo_nota = ""
    in_serie_afectado = ""
    in_numero_afectado = ""
    in_observacion = Replace(Trim(Me.txtObservacion.Text), "'", " ")


If Me.DtcTipoDoc.BoundText = "0007" Then
    Select Case Me.DtcComprobanteGuia.BoundText
    Case "0003"
         If Trim(Me.TxtCodCliente.Text) = "00000000" Then
            in_tipo_doc = "1"
         Else
            in_tipo_doc = "1"
         End If
         
         If Trim(Me.txtExtranjero.Text) = "si" Then
            in_tipo_doc = "4"
         End If
         
         in_tipo_doc_nota = "03"
         
    Case "0001"
        in_tipo_doc = "6"
        If Trim(Me.txtExtranjero.Text) = "si" Then
            in_tipo_doc = "0"
         End If
        in_tipo_doc_nota = "01"
        
    Case "0002"
        n_tipo_doc = "1"
End Select
    
    
    
    
    id_motivo_nota = Me.DtcTipoNota.BoundText
    motivo_nota = Trim(Me.txtmotivo_nota.Text)
    in_serie_afectado = Trim(Me.TxtSeri_guia.Text)
    in_numero_afectado = Trim(Me.TxtNumero_guia.Text)
End If

If Me.DtcTipoDoc.BoundText = "0008" Then
    Select Case Me.DtcComprobanteGuia.BoundText
    Case "0003"
         If Trim(Me.TxtCodCliente.Text) = "00000000" Then
            in_tipo_doc = "1"
         Else
            in_tipo_doc = "1"
         End If
         If Trim(Me.txtExtranjero.Text) = "si" Then
            in_tipo_doc = "4"
         End If
         
         in_tipo_doc_nota = "03"
         
    Case "0001"
        in_tipo_doc = "6"
        If Trim(Me.txtExtranjero.Text) = "si" Then
            in_tipo_doc = "0"
        End If
        in_tipo_doc_nota = "01"
        
    Case "0002"
        n_tipo_doc = "1"
        
     Case "0007"
      If Mid(Me.TxtSeri_guia.Text, 1, 1) = "B" Then
         If Trim(Me.txtExtranjero.Text) = "si" Then
            in_tipo_doc = "4"
         End If
         
         in_tipo_doc = "1"
      Else
         in_tipo_doc = "6"
      End If
       
       in_tipo_doc_nota = "07"
    End Select
    
    
        
    id_motivo_nota = Me.DtcTipoNota.BoundText
    motivo_nota = Trim(Me.txtmotivo_nota.Text)
    in_serie_afectado = Trim(Me.TxtSeri_guia.Text)
    in_numero_afectado = Trim(Me.TxtNumero_guia.Text)
End If

Dim in_moneda As String
in_moneda = "PEN"

If Me.chk_detraccion.Value = 1 Then
   in_detraccion = "si"
Else
   in_detraccion = "no"
End If


If Me.DtcMoneda.BoundText <> "00001" Then
    
    If get_moneda_documento(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText)) = "00002" Then
       in_moneda = "USD"
    End If
    If get_moneda_documento(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText)) = "00003" Then
       in_moneda = "EUR"
    End If
End If



If get_comprobante_produccion(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText)) = "si" Then
    If Me.frmmanual.Visible = True And Me.chk_manual.Value = 1 Then
        in_numero = Trim(Me.TxtNumeroDoc.Text)
    Else
        in_numero = Trim(Me.TxtNumeroDoc.Text)
    End If
    
    
    If KEY_SERVIDOR_KEYFACIL = "si" Then
        Call FrmLoad_web_service.crear_json_facturacion_electronica("https://api.vitekey.com/keyfact/utils/erp-invoice?password=vitekey2018&company_id=" & KEY_TOKEN_CLOUD & "", "POST", json_facturacion_electronica_firmar_id_venta_keyfacil(Val(Me.txtidVenta.Text), Format(Val(Me.DtcTipoDoc.BoundText), "00"), Trim(Me.DtcSerieDoc.BoundText), in_numero, Format(Me.DtpActual.Value, "YYYY-mm-dd"), Trim(Me.TxtCodCliente.Text), Trim(Me.TxtCliente.Text), Trim(Me.txtDireccion.Text), in_tipo_doc, Val(Me.TxtDescuento_porcentaje.Text), KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion, in_detraccion, Trim(Me.txtOrdenCompra.Text), Val(Me.lblPercepcion.Caption)), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
    Else
        If KEY_SERVIDOR_CLOUD = "si" Then
        Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(Me.txtidVenta.Text), Format(Val(Me.DtcTipoDoc.BoundText), "00"), Trim(Me.DtcSerieDoc.BoundText), in_numero, KEY_FECHA, Trim(Me.TxtCodCliente.Text), Trim(Me.TxtCliente.Text), Trim(Me.txtDireccion.Text), in_tipo_doc, Val(Me.TxtDescuento_global.Text), KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_CLOUD & "', x-api-produccion: 'yes'}")
        'KEY_COMPANY = "374de97b-4a82-4c23-b9d4-a65d7417a1db"
        'KEY_TOKEN_CLOUD = "374de97b-4a82-4c23-b9d4-a65d7417a1db"
        'Call FrmLoad_web_service.crear_json_facturacion_electronica("https://api.vitekey.com/keyfact/utils/erp-invoice?password=vitekey2018&company_id=" & KEY_TOKEN_CLOUD & "", "POST", json_facturacion_electronica_firmar_id_venta_keyfacil(Val(Me.TxtIdVenta.Text), Format(Val(Me.DtcTipoDoc.BoundText), "00"), Trim(Me.DtcSerieDoc.BoundText), in_numero, KEY_FECHA, Trim(Me.TxtCodCliente.Text), Trim(Me.txtcliente.Text), "", in_tipo_doc, Val(Me.TxtDescuento_porcentaje.Text), KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
        
        Else
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(Me.txtidVenta.Text), Format(Val(Me.DtcTipoDoc.BoundText), "00"), Trim(Me.DtcSerieDoc.BoundText), in_numero, KEY_FECHA, Trim(Me.TxtCodCliente.Text), Trim(Me.TxtCliente.Text), Trim(Me.txtDireccion.Text), in_tipo_doc, Val(Me.TxtDescuento_porcentaje.Text), KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_LOCAL & "', x-api-produccion: 'yes'}")
        End If
    End If
    
    
    
    
Else
    in_numero = Trim(Me.TxtNumeroDoc.Text)
    
    
    If KEY_SERVIDOR_KEYFACIL = "si" Then
            Call FrmLoad_web_service.crear_json_facturacion_electronica("https://api.vitekey.com/keyfact/utils/erp-invoice?password=vitekey2018&company_id='" & KEY_TOKEN_CLOUD & "'", "POST", json_facturacion_electronica_firmar_id_venta_keyfacil(Val(Me.txtidVenta.Text), Format(Val(Me.DtcTipoDoc.BoundText), "00"), Trim(Me.DtcSerieDoc.BoundText), in_numero, Format(Me.DtpActual.Value, "YYYY-mm-dd"), Trim(Me.TxtCodCliente.Text), Trim(Me.TxtCliente.Text), Trim(Me.txtDireccion.Text), in_tipo_doc, Val(Me.TxtDescuento_porcentaje.Text), KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion, in_detraccion, Trim(Me.txtOrdenCompra.Text), Val(Me.lblPercepcion.Caption)), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
    Else
        If KEY_SERVIDOR_CLOUD = "si" Then
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://facturacion.vitekey.com/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(Me.txtidVenta.Text), Format(Val(Me.DtcTipoDoc.BoundText), "00"), Trim(Me.DtcSerieDoc.BoundText), in_numero, KEY_FECHA, Trim(Me.TxtCodCliente.Text), Trim(Me.TxtCliente.Text), Trim(Me.txtDireccion.Text), in_tipo_doc, Val(Me.TxtDescuento_global.Text), KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_CLOUD & "'}")
        Else
            Call FrmLoad_web_service.crear_json_facturacion_electronica("http://192.168.1.241:3030/api/comprobantes/enviar", "POST", json_facturacion_electronica_firmar_id_venta(Val(Me.txtidVenta.Text), Format(Val(Me.DtcTipoDoc.BoundText), "00"), Trim(Me.DtcSerieDoc.BoundText), in_numero, KEY_FECHA, Trim(Me.TxtCodCliente.Text), Trim(Me.TxtCliente.Text), Trim(Me.txtDireccion.Text), in_tipo_doc, Val(Me.TxtDescuento_porcentaje.Text), KEY_IGV, id_motivo_nota, motivo_nota, in_tipo_doc_nota, in_serie_afectado, in_numero_afectado, in_moneda, in_observacion), "{x-api-token: '" & KEY_TOKEN_LOCAL & "'}")
        End If
    End If
End If



End Function
Private Function facturacion_electronica()
Call disabled_form(Me)
FrmLoad_web_service.Show
FrmLoad_web_service.nom_prcedimiento = "procesar_facturacion_electronica"
Set FrmLoad_web_service.FormPadre = Me
Select Case Me.DtcTipoDoc.BoundText
    Case "0003"
         in_tipo_doc = "1"
    Case "0001"
        in_tipo_doc = "6"
    Case "0002"
        in_tipo_doc = "1"
    
End Select

If Me.DtcTipoDoc.BoundText = "0007" Then
    Select Case Me.DtcComprobanteGuia.BoundText
    Case "0003"
         in_tipo_doc = "1"
    Case "0001"
        in_tipo_doc = "6"
    Case "0002"
        n_tipo_doc = "1"
    End Select
End If

Call FrmLoad_web_service.crear_json_facturacion_electronica("http://sunat.vitekey.com/web/send-invoice", "POST", json_facturacion_electronica(Format(Val(Me.DtcTipoDoc.BoundText), "00"), Trim(Me.DtcSerieDoc.BoundText), Val(Me.TxtNumeroDoc.Text), KEY_FECHA, Trim(Me.TxtCodCliente.Text), Trim(Me.TxtCliente.Text), in_tipo_doc, 0), "")
End Function
Public Function duplicado(ByVal in_doc As String, ByVal in_serie As String, ByVal in_numero As String) As Boolean
     strCadena = "SELECT count(*) FROM movimiento_venta WHERE  id_doc='" & in_doc & "' and serie='" & in_serie & "' and numero='" & in_numero & "' and ruc='" & KEY_RUC & "'"
     Call ConfiguraRst(strCadena)
     If rst(0) < 1 Then
        duplicado = False
    Else
        duplicado = True
    End If
End Function
Private Sub update_nota(ByVal in_venta As String, ByVal in_tipo_doc As String, ByVal in_serie As String, ByVal in_numero As String)

strCadena = "UPDATE movimiento_venta SET id_doc_fact='" & in_tipo_doc & "',serie_fact='" & in_serie & "',numero_fact='" & in_numero & "' WHERE id_venta='" & Val(in_venta) & "'"
CnBd.Execute (strCadena)



strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & in_tipo_doc & "' and serie='" & in_serie & "' and numero='" & in_numero & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstI(strCadena)
If rstI.RecordCount > 0 Then
    in_comprobante = rstI("id_venta")
    strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & Val(rstI("id_venta")) & "' and ruc='" & KEY_RUC & "' "
    Call ConfiguraRstI(strCadena)
    If rstI.RecordCount > 0 Then
        rstI.MoveFirst
        For i = 0 To rstI.RecordCount - 1
            
            'Solo en el Caso de Anulacion y/o Devolucion va a proceder a estar disponible
            If Me.DtcTipoNota.BoundText = "01" Or Me.DtcTipoNota.BoundText = "02" Or Me.DtcTipoNota.BoundText = "06" Or Me.DtcTipoNota.BoundText = "07" Then
                
                strCadena = "UPDATE imp_producto_detalle SET vendido='no' WHERE id_detalle='" & rstI("id_detalle_serie") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
                CnBd.Execute (strCadena)
                
            End If
            
            
            rstI.MoveNext
        Next i
    End If
    
    
    Call revertir_letras_nota(in_venta, in_comprobante, "0007")
End If







End Sub
Private Sub revertir_letras_nota(ByVal in_nota As String, ByVal in_venta As String, ByVal in_doc As String)
Dim in_monto As Double
Dim in_monto_pago As Double

If in_doc = "0007" Then

      strCadena = "SELECT ifnull(total,0) as total FROM movimiento_venta WHERE id_venta='" & in_nota & "' and ruc='" & KEY_RUC & "'"
      Call ConfiguraRstlocal(strCadena)
      in_monto = rstLocal("total")
      
      strCadena = "SELECT id_venta,id_doc,tc,monto_interes,(total-function_pago_factura(id_venta,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc)) as saldo FROM movimiento_venta WHERE id_referencia='" & in_venta & "' and id_doc='0412' and ruc='" & KEY_RUC & "'"
      Call ConfiguraRstlocal(strCadena)
      If rstLocal.RecordCount > 0 Then
          rstLocal.MoveFirst
          
          For i = 0 To rstLocal.RecordCount - 1
               If in_monto > 0 Then
              strCadena = "UPDATE movimiento_venta SET interes_revertido='" & rstLocal("monto_interes") * -1 & "' WHERE id_venta='" & rstLocal("id_venta") & "'"
              CnBd.Execute (strCadena)
             
              If in_monto > rstLocal("saldo") Then
                 in_monto_pago = rstLocal("saldo")
                 in_monto = in_monto - in_monto_pago
              Else
                in_monto_pago = in_monto
                in_monto = 0
              End If
              Call put_realizar_pago(rstLocal("id_venta"), in_nota, in_monto_pago, rstLocal("id_doc"), rstLocal("tc"), Val(in_mis_cuentas_det))
              Call put_realizar_pago(in_nota, rstLocal("id_venta"), in_monto_pago, rstLocal("id_doc"), rstLocal("tc"), Val(in_mis_cuentas_det))
              End If
              rstLocal.MoveNext
              
          Next i
       End If
End If

End Sub



Public Sub procesar_ws(ByVal strHtml As String)
Dim in_error As Boolean
Dim StrNumero As String
Dim in_numero() As String
Dim json_r As Object
Set json_r = JSON.parse(strHtml)
in_error = json_r.Item("error")
If in_error = True Then

Else
     
     
     StrNumero = Trim(json_r.Item("response").Item("num_recibo"))
     in_numero = Split(Trim(StrNumero), "-")
     
    Call update_temporal(DtcTipoDoc.BoundText, Trim(in_numero(0)), Val(in_numero(1)), Trim(Me.TxtCodCliente.Text), txtSerie, TxtNumeroDoc)
  
   Call procesar
End If
End Sub

Public Sub procesar()
Dim in_fua_actual As String
Dim strestado As String


strCadena = "SELECT count(*) FROM temporal_ventas WHERE id_dni='" & Trim(Me.TxtCodCliente.Text) & "' and  dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst(0) < 1 Then
    MsgBox "Imposible Guardar este comprobante." + Chr(13) + "Debe contener algun detalle", vbInformation, KEY_VENDEDOR
    Exit Sub
End If

If duplicado(Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText, Me.TxtNumeroDoc.Text) = True Then
   MsgBox "ESTE COMPROBANTE YA HA ESTA REGISTRADO" + Chr(13) + Chr(13) + "VERIFIQUE EL CORRELATIVO DE SU SERIE.", vbInformation
   Call enabled_form(Me)
   Exit Sub
End If

If Save = True Then
   Me.cmdProcesar.Enabled = False
   Me.cmdImprimir.Enabled = True
   
        
        Call OrdenImpresion(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text))
   
   
   
        
        If Val(Me.txt_id_pendiente.Text) > 0 And (KEY_CARGO = "00008" Or KEY_CARGO = "00004") Then
           Call get_pendiente
        End If
        Call nuevo
        Exit Sub
End If
        Exit Sub
     
     
     
     
     
     If Len(Me.TxtCodCliente.Text) < 11 And Len(Me.TxtCodCliente.Text) > 11 Then
        MsgBox "INGRESE UN RUC PARA EL CLIENTE", vbInformation, KEY_EMPRESA
        Call Resalta(Me.TxtCodCliente)
        Exit Sub
     End If
      
    If (Val(Me.lblPago.Caption) > 1 Or Me.DtcTipoDoc.BoundText = KEY_GUIA) Then
        Call get_auto_pago(Me.DtcTipoDoc.BoundText)
        Call Save
        Call nuevo
        
    Else
        If Me.DtcTipoDoc.BoundText = "0099" Then ' proforma
            Call disabled_form(Me)
            Procedencia = seleccionar_vendedor
            FrmSeguridad.Show
            
            Exit Sub
        End If
        MsgBox "INGRESE UN MONTO PARA EL COMPROBANTE", vbInformation, KEY_EMPRESA
        If Me.DtcFormaPago.BoundText = "01" Then
           Me.DtcMoneda.SetFocus
           Exit Sub
        Else
            Me.DtcFormaPago.SetFocus
        End If
    End If




End Sub

Public Sub procesar_firma_electronica(ByVal strHtml As String)
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
     'Me.TxtNumeroDoc.Text = Trim(get_numero)
     
     strCadena = "UPDATE movimiento_venta SET sunat_key='" & Trim(in_key) & "',sunat_hash='" & Trim(in_hash) & "' WHERE id_venta='" & Val(Me.txtidVenta.Text) & "'"
     CnBd.Execute (strCadena)
     Call OrdenImpresion(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text))
     Call next_save
      If Val(Me.txt_id_pendiente.Text) > 0 And (KEY_CARGO = "00008" Or KEY_CARGO = "00004") Then
                        Call get_pendiente
     End If
                    
     Me.Enabled = True
     Exit Sub
     
     'Call procesar_comprobante
     
End If
Exit Sub
procesar_nuevamente:
MsgBox "SE PRESENTO UN PROBLEMA CON EL INTERNET" + Chr(13) + Chr(13) + "INTENTENTALO NUEVAMENTE.", vbInformation, KEY_USUAURIO
Me.Enabled = True
Me.cmdProcesar.Enabled = True
End Sub

Public Sub procesar_firma_electronica_eliminar(ByVal strHtml As String)
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
     in_hash = Trim(json_r.Item("response").Item("comprobante").Item("digest_value"))
     in_key = Trim(json_r.Item("response").Item("comprobante").Item("key"))
     Me.txt_hash.Text = Trim(in_hash)
     Me.txt_sunat_key.Text = Trim(in_key)
     Call procesar_comprobante
     
End If
End Sub
Public Sub procesar_firma_electronica_reenvio(ByVal strHtml As String)
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
     in_hash = Trim(json_r.Item("response").Item("comprobante").Item("digest_value"))
     in_key = Trim(json_r.Item("response").Item("comprobante").Item("key"))
     Me.txt_hash.Text = Trim(in_hash)
     Me.txt_sunat_key.Text = Trim(in_key)
     Call procesar_comprobante
     
End If
End Sub

Public Sub procesar_facturacion_electronica(ByVal strHtml As String)
Dim in_error As Boolean
Dim StrNumero As String
Dim in_numero() As String
Dim json_r As Object
Set json_r = JSON.parse(strHtml)
in_error = json_r.Item("error")
If in_error = True Then

Else
     StrNumero = Trim(json_r.Item("response").Item("data").Item("numero"))
     strSerie = Trim(json_r.Item("response").Item("data").Item("serie"))
     Call update_temporal(DtcTipoDoc.BoundText, strSerie, Format(StrNumero, "000000"), Trim(Me.TxtCodCliente.Text), txtSerie, TxtNumeroDoc)
  
   Call procesar
End If
End Sub

Public Sub listar_estado_comprobantes(ByVal strHtml As String)
Dim in_error As Boolean
Dim StrNumero As String
Dim in_numero() As String
Dim json_r As Object
Set json_r = JSON.parse(strHtml)
in_error = json_r.Item("error")
If in_error = True Then

Else
     StrNumero = Trim(json_r.Item("response").Item("data").Item("numero"))
     strSerie = Trim(json_r.Item("response").Item("data").Item("serie"))
     Call update_temporal(DtcTipoDoc.BoundText, strSerie, Format(StrNumero, "000000"), Trim(Me.TxtCodCliente.Text), txtSerie, TxtNumeroDoc)
  
   Call procesar
End If
End Sub

Private Sub CmdQuitar_Click()
If Me.HfdDetalle.Rows > 0 Then
If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    
    If Me.DtcTipoDoc.BoundText = "0099" And KEY_UPDATE_PROFORM = "si" And Val(Me.txtidVenta.Text) > 0 And Me.cmdModificar.Enabled = False Then
        Call Quitar_save(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))
    Else
        Call Quitar(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))
    End If
    
End If
End If
End Sub
Public Sub quitar_bonificacion_cruzada(ByVal in_producto As String, ByVal in_cliente As String, ByVal in_temporal As String)
        
   
        
        strCadena = "DELETE FROM temporal_ventas WHERE obsequio='si' and id_persona_analisis='" & Val(in_temporal) & "' and  ruc='" & KEY_RUC & "' and dni_save='" & KEY_USUARIO & "'"
        CnBd.Execute (strCadena)
  
    
    
    
  Exit Sub
    strCadena = "SELECT DISTINCT id_bonificacion,all_canal,id_cobertura FROM view_bonificacion_cruzada_venta WHERE  cruzada='si' and  id_producto='" & in_producto & "' and  dni_save='" & KEY_USUARIO & "' and id_dni='" & in_cliente & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRstC(strCadena)
    If rstc.RecordCount > 0 Then
    rstc.MoveFirst
     
     For R = 0 To rstc.RecordCount - 1
        
        If rstc("all_canal") = "si" Then
        
        Else
           If rstc("id_cobertura") <> get_tipo_cobertura(in_cliente) Then
            GoTo siguiente_bonificacion
           End If
        End If
        
        
        
     
     
        in_bonificcion = False
        
        strCadena = "SELECT DISTINCT id_producto,sum(cantidad_temp) as cantidad_temp,id_bonificacion FROM view_bonificacion_cruzada_venta WHERE  id_bonificacion='" & rstc("id_bonificacion") & "'and  dni_save='" & KEY_USUARIO & "' and id_dni='" & in_cliente & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' GROUP BY id_bonificacion,id_producto"
        Call ConfiguraRstA(strCadena)
        
        strCadena = "SELECT * FROM bonificacion_detalle WHERE  id_bonificacion='" & rstc("id_bonificacion") & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        
        If rstA.RecordCount = rstL.RecordCount And rstA.RecordCount > 0 Then
            in_multiplo = 1
            rstA.MoveFirst
            in_pedido_anterior = rstA("cantidad_temp")
            codigo_bonificacion = rstA("id_bonificacion")
           
        
            
        'verificacion multiplos
        
        
        strCadena = "SELECT * FROM bonificacion_cruzada_detalle WHERE  id_bonificacion='" & codigo_bonificacion & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstA(strCadena)
        If rstA.RecordCount > 0 Then
           rstA.MoveFirst
        For m = 0 To rstA.RecordCount - 1
            in_unidad = rstA("id_unidad")
       
        in_cantidad = in_multiplo * rstA("cantidad")
        strCadena = "SELECT id_linea,id_sublinea,id_unidad,agranel FROM producto WHERE id_producto='" & rstA("id_producto") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
        
          strCadena = "DELETE FROM temporal_ventas WHERE id_unidad='" & in_unidad & "' and  ruc='" & KEY_RUC & "' and id_alm='" & KEY_ALM & "' and dni_save='" & KEY_USUARIO & "' and obsequio='si' and tipo_bonificacion='02' and  id_producto='" & rstA("id_producto") & "' LIMIT 1 "
         CnBd.Execute (strCadena)
                
        
        
        
        
       
        End If
        rstA.MoveNext
        Next m
        
        
        End If
        End If
siguiente_bonificacion:
      rstc.MoveNext
     Next R


End If



End Sub



Private Sub Quitar(ByVal codigo As Double)
On Error GoTo salir
Dim in_codigo As String
If Val(codigo) > 0 Then
    
    
    If KEY_BONIFICACIONES = "si" Then
        
        strCadena = "SELECT * FROM temporal_ventas WHERE id='" & codigo & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            in_dni = rst("id_dni")
            in_producto = rst("id_producto")
        End If
        Call quitar_bonificacion_linea(codigo)
        
        
        
        'strCadena = "SELECT * FROM temporal_ventas WHERE id='" & Val(codigo) & "'"
        'Call ConfiguraRst(strCadena)
        'If rst.RecordCount > 0 Then
            Call quitar_bonificacion_cruzada(in_producto, Trim(Me.TxtCodCliente.Text), codigo)
        'End If
        strCadena = "DELETE FROM temporal_ventas WHERE id='" & Trim(codigo) & "'  "
        CnBd.Execute (strCadena)
    Else
        
        strCadena = "DELETE FROM temporal_ventas WHERE id='" & Trim(codigo) & "'  "
        CnBd.Execute (strCadena)
    End If
    

End If


        
         
         
        
    
    
    
    Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcSerieDoc.BoundText, Me.DtcTipoDoc.BoundText, Me.txtformato_impresion.Text)
    
    
    Exit Sub
salir:

End Sub

Private Sub Quitar_save(ByVal codigo As Double)
If Trim(codigo) <> "" Then
    strCadena = "DELETE FROM movimiento_venta_detalle WHERE id_detalle_venta='" & Trim(codigo) & "' and ruc='" & KEY_RUC & "'  "
    CnBd.Execute (strCadena)
    Call Me.llenarGrid_Comprobante_edit(Me.HfdDetalle, Val(Me.txtidVenta.Text))
     
    
End If
End Sub



Private Sub cmdQuitarMonto_Click()
If Me.HfgTipoPagos.Rows > 0 Then
If Val(Me.HfgTipoPagos.TextMatrix(Me.HfgTipoPagos.Row, 0)) > 0 Then
    strCadena = "DELETE  FROM movimiento_venta_monto_temporal WHERE id_monto='" & Trim(Me.HfgTipoPagos.TextMatrix(Me.HfgTipoPagos.Row, 0)) & "' "
    CnBd.Execute (strCadena)
    
    strCadena = "DELETE FROM movimiento_venta_targeta_temporal WHERE id_temporal='" & Trim(Me.HfgTipoPagos.TextMatrix(Me.HfgTipoPagos.Row, 0)) & "' AND id_usuario='" & KEY_USUARIO & "' ANd ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    
     
    Call llena_pagos(Me.HfgTipoPagos, Me.TxtNumeroDoc.Text)
    
End If
End If
End Sub

Private Sub DtaGasto_KeyPress(KeyAscii As Integer)

End Sub

Private Sub cmdRecibo_Click()

End Sub

Private Sub cmdSalir_Click()
Me.fraApp.Visible = False
End Sub


Public Sub automatico_fac()
Call mostrar_comprobante(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text))
Call facturar
End Sub


Private Sub llenar_montos(ByVal id_venta As Double)
strCadena = "SELECT * FROM movimiento_venta_monto WHERE id_venta='" & Val(Me.txtidVenta.Text) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    rstL.MoveFirst
    For i = 0 To rstL.RecordCount - 1
        strCadena = "INSERT INTO movimiento_venta_monto(id_venta,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,ruc)VALUES('" & id_venta & "','" & rstL("id_forma_pago") & "','" & rstL("monto") & "','" & rstL("monto") & "','" & rstL("id_tarjeta") & "','" & rstL("id_tarjeta_numero") & "','" & rstL("id_tarjeta_operacion") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
    
         rstL.MoveNext
    Next i
End If


End Sub
Private Sub cmdSeriales_Click()
If Val(Me.txtidVenta.Text) > 0 And Me.cmdProcesar.Enabled = False Then
    strCadena = "SELECT D.serie,D.anio_fabricacion,D.anio_modelo,D.nro_chasis,D.id_producto,D.nro_dua,D.nro_item FROM movimiento_venta_detalle D WHERE  D.id_detalle_venta='" & Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "' AND id_venta='" & Val(Me.txtidVenta.Text) & "'"
    Me.txtid_temporal_serie.Text = 0
Else
    strCadena = "SELECT * FROM temporal_ventas WHERE id='" & Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "' "
    Me.txtid_temporal_serie.Text = Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))
End If
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.DtcSerie.Text = rst("nro_chasis")
    Me.DtcMotor.Text = rst("serie")
    Me.txtA�oFabricacion.Text = rst("anio_fabricacion")
    Me.txtModelo.Text = rst("anio_modelo")
    Me.txtdua.Text = rst("nro_dua")
    Me.txtitem = rst("nro_item")
    
    strCadena = "SELECT color,marca,modelo FROM view_producto WHERE id_producto='" & rst("id_producto") & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRstT(strCadena)
    Me.txtcolor.Text = rstT("color")
    Me.TxtMarca.Text = rstT("marca")
    Me.txtModelo.Text = rstT("modelo")
    Me.FrameSerieModelo.Visible = True
    
    Call parametro_importacion
    
    Exit Sub
End If
End Sub

Private Sub cmdSolicitud_Click()

Dim in_direccion As String
If Len(Trim(Me.TxtCodCliente.Text)) = 11 And Mid(Trim(Me.TxtCodCliente.Text), 1, 2) <> "10" Then
    
    strCadena = "SELECT p.dni,p.nombre_completo,p.direccion,a.dni_familia FROM persona_accidentes a,persona p where a.dni='" & Trim(Me.TxtCodCliente.Text) & "' and a.dni_familia=p.dni LIMIT 1"
     Call ConfiguraRstK(strCadena)
     If rstK.RecordCount > 0 Then
        strCadena = "UPDATE persona SET direccion='" & Trim(Me.txtDireccion.Text) & "' WHERE dni='" & rstK("dni_familia") & "'"
        CnBd.Execute (strCadena)
    End If
        
        
        
    
    
    
    

strCadena = "select d.`nro_chasis`, d.`serie` , s.`descripcion` as modelo, m.`descripcion` as `marca` " & _
",aa.`nombre_completo` as paciente,aa.`direccion`  as direccion , " & _
" a.`dni_familia`, v.`fecha_emision`from movimiento_venta v , " & _
" `movimiento_venta_detalle` d,producto p,`linea_sub` s,marca m,persona pr,`persona_accidentes` a ,persona aa where v.`id_venta` = d.`id_venta` and d.`id_producto` = p.`id_producto` and " & _
" v.`ruc` = p.`ruc` and p.`id_sublinea` = s.`id_tipo` and v.ruc = s.`id_usu` and p.`id_marca` = m.`id_marca` and v.`ruc` = m.`id_usu` and v.`id_cliente` = pr.`dni` and v.`id_venta` ='" & Val(Me.txtidVenta.Text) & "' and " & _
" v.`id_cliente`=a.`dni` and a.`dni_familia`=aa.`dni` "

Else
 strCadena = "select d.`nro_chasis`, d.`serie` , s.`descripcion` as modelo, " & _
"m.`descripcion` as `marca`, pr.nombre_completo as paciente,CONCAT(pr.`direccion`,'-',funct_ubigueo(pr.id_departamento,pr.id_provincia,pr.id_distrito)) as direccion , pr.`dni`, v.`fecha_emision` " & _
"from movimiento_venta v , `movimiento_venta_detalle` d, producto p, " & _
"`linea_sub` s, marca m, persona pr " & _
"where v.`id_venta` = d.`id_venta` and d.`id_producto` = p.`id_producto` and v.`ruc` = p.`ruc` " & _
"and p.`id_sublinea` = s.`id_tipo` and v.ruc = s.`id_usu` " & _
"and p.`id_marca` = m.`id_marca` and v.`ruc` = m.`id_usu` " & _
"and v.`id_cliente` = pr.`dni` and v.`id_venta` ='" & Val(Me.txtidVenta.Text) & "'"
End If


  

  Call ConfiguraRstK(strCadena)
  
  strCadena = "select dni, nombre from  imp_reporte_app where `id_venta` = " & Val(Me.txtidVenta.Text)
  
  Call ConfiguraRstL(strCadena)
  strCadena = "SELECT id_venta,nro_chasis,serie,marca,modelo,anio_fabricacion,color,nro_dua,nro_item FROM view_datos_chasis_motor WHERE id_venta='" & Val(Me.txtidVenta.Text) & "'"
  Call ConfiguraRstP(strCadena)
  Ans = ShowMultiReport(rstK, "CorApp", , App.Path + "\Reportes\", , , , , rstP, "rpt_datos_vehiculo")
  'Ans = ShowMultiReport(rstK, "CorConstanciaVenta", , App.Path + "\Reportes\")
  'Ans = ShowMultiReport(rstK, "CorApp", , App.Path + "\Reportes\")
'Ans = ShowMultiReport(rstK, "CorApp", , App.Path + "\Reportes\")






  Exit Sub
  

End Sub
Public Sub cargarParientes(ByVal Grilla As MSHFlexGrid)

strCadena = " select * from imp_reporte_app where id_venta = '" & Val(Me.txtidVenta.Text) & "'"
                  
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Grilla.Cols = 3
    Grilla.Refresh
    Grilla.Clear
    
    cabecera = "" & vbTab & "DNI" & vbTab & "PACIENTE" & vbTab & ""
    Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
         
    Grilla.ColWidth(0) = 0
    Grilla.ColWidth(1) = 1200
    Grilla.ColWidth(2) = 3500
    Grilla.ColWidth(3) = 0
           
           
    Exit Sub
    
End If


  N = 1
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 3500
           Grilla.ColWidth(3) = 0
          
           
        Next
         cabecera = "" & vbTab & "DNI" & vbTab & "PACIENTE" & vbTab & ""
         Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
        
             estado = Chr(168)
        
             Fila = rst("id_detalle") & vbTab & rst("dni") & vbTab & rst("nombre") & vbTab & ""
             Grilla.AddItem Fila
             
             'With Grilla
             
                 '.row = i + 1 ' se posiciona en la fila
                 '.col = 1 '  .. en la columna
                 
                 ' cambia la fuente para esta celda
                            
                 '.CellFontName = "Wingdings"
                 '.CellFontSize = 14
                 '.CellAlignment = flexAlignCenterCenter
    
             'End With
             
             
             Fila = ""
             rst.MoveNext
        Next i
        
        
Exit Sub


  
End Sub


Private Sub cmdSolicitudcredito_Click()
Call disabled_form(Me)
frmventa_solicitud.Show
Call frmventa_solicitud.put_llenar(Me.txtidVenta.Text)
Exit Sub

End Sub

Private Sub cmdvincular_Click()
Me.TxtMontoPagado.Text = Format(Me.HfRecibos.TextMatrix(Me.HfRecibos.Row, 3), "###0.00")
Me.txtrecibo_anterior.Text = Me.HfRecibos.TextMatrix(Me.HfRecibos.Row, 0)
Me.fraApp.Visible = False
Call realizar_ingreso_pago
Exit Sub
End Sub

Private Sub cmdVisualizar_Click()
FrmReporteRegistroVentas.Show

FrmReporteRegistroVentas.txtRuc.Text = Trim(Me.TxtCodCliente.Text)
strCadena = "SELECT * FROM view_listado_comprobante_ultimate WHERE saldo<>0 and  id_cliente LIKE '%" & Trim(Me.TxtCodCliente.Text) & "%' AND ruc='" & KEY_RUC & "'"
Call FrmReporteRegistroVentas.llenar_grid(FrmReporteRegistroVentas.HfdPersona)


Exit Sub


End Sub


Public Function enviar_post(ByVal in_venta As String) As Boolean

Dim Nombre As String
    On Error GoTo salir_cancel


    Dim strHtml As String
    urlstr = "https://192.168.3.250/facturacion/web/"
     Set DomDoc = New XMLHTTP
     'Par�metros en formato URLEncode
    ' ventanilla = Split(rstaux("ventanilla"), "-")
     
     params = "idVenta=" & in_venta
     'Metodo a usar, url, y true en caso de manejar la respuesta en modo as�ncrono
     DomDoc.Open "POST", urlstr, False
     'encabezados
     DomDoc.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
     DomDoc.setRequestHeader "Content-length", Len(params)
     DomDoc.setRequestHeader "Connection", "close"
     DomDoc.send params
     'La respuesta, en caso de existir, est� en responseBody.
    'Tambi�n puedes especificar responseXml si tu aplicaci�n devolviese XML
     strHtml = StrConv(DomDoc.responseBody, vbUnicode)
     
     Dim p As Object
     Set p = JSON.parse(strHtml)
     Nombre = Trim(p.Item("apellidoPaterno")) & Space(1) & Trim(p.Item("apellidoMaterno")) & Space(1) & Trim(p.Item("nombres"))
     If Trim(Nombre) <> "" Then
        strCadena = " call P_insert_persona('" & Trim(in_dni) & "','" & Trim(p.Item("apellidoPaterno")) & "','" & Trim(p.Item("apellidoMaterno")) & "','" & Trim(p.Item("nombres")) & "','" & Trim(Nombre) & "','-','','-','no','no','no','no','no','no','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        'get_dni_reniec== True
    Else
        'get_dni_reniec = False
    End If
    


     
     
     
     Exit Function
  'get_dni_reniec = False
salir_cancel:

End Function
Private Sub put_disparar_electronica()

strCadena = "SELECT * FROM movimiento_venta WHERE sunat_key='-' and  id_doc IN('0001','0003','0007') and  fecha_emision>='2019-01-08' and  ruc='" & KEY_RUC & "' ORDER BY fecha_emision ASC, id_venta ASC LIMIT 1"
Call ConfiguraRstUpdate(strCadena)
If rstUpdate.RecordCount > 0 Then
   rstUpdate.MoveFirst
   For i = 0 To rstUpdate.RecordCount - 1
      
        Call nuevo
        Call mostrar_comprobante(rstUpdate("id_doc"), Trim(rstUpdate("serie")), Trim(rstUpdate("numero")))
        Call firma_electronica
        
      
   Next i
End If
End Sub





Private Sub Command1_Click()




Call Me.firma_electronica_eliminar


Call EliminarVentas(Trim(FrmVentas.DtcTipoDoc.BoundText), Trim(FrmVentas.DtcSerieDoc.BoundText), Trim(FrmVentas.TxtNumeroDoc.Text), Trim(FrmVentas.DtcAlmacen.BoundText))
                   
                    

End Sub

Private Sub DtcAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcTipoDoc.SetFocus
End If
End Sub

Private Sub DtcComprobanteGuia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSeri_guia.Text = Trim(Me.DtcSerieDoc.BoundText)
    Call Resalta(Me.TxtSeri_guia)
End If
End Sub




Private Sub forma_pago()
On Error GoTo salir


Dim Total As Double
Dim pagado As Double
Dim tcredito As Double


If Me.HfgTipoPagos.Rows < 1 Then
    Me.lblPago.Caption = Format(0, "###0.000")
End If

If (Trim(Me.DtcFormaPago.BoundText) = "01") Then
    Me.frmvencimiento.Visible = False
    Me.lbltargeta.Visible = False
    Me.DtTargeta.Visible = False
    Me.TxtNumeroTargeta.Visible = False
    Me.DtpFechaReferencia.Visible = False
    
    Me.PanelCredito.Visible = False
    Me.lblSaldodisponible.Visible = False
    Me.DtcFormapagodetalle.Visible = True
    Me.TxtMontoPagado.Enabled = True
    Me.frminteres.Visible = False
    
   ' If Me.DtcMoneda.BoundText = "00002" Then
   '     Me.TxtMontoPagado.Text = Format(Val(Format(Me.lblTotal.Caption, "###0.000")) / (KEY_CAMBIO) - Val(Format(Me.lblPago.Caption, "###0.000")), "###0.000")
   ' Else
        Me.TxtMontoPagado.Text = Format(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000")), "###0.000")
   ' End If
    
    
ElseIf (Trim(Me.DtcFormaPago.BoundText) = "02") Then
    
    Me.lbltargeta.Visible = False
    Me.DtTargeta.Visible = False
    Me.frminteres.Visible = True
    Me.txtporcentaje_interes.Text = 0
    Me.TxtNumeroTargeta.Visible = False
    Me.DtpFechaReferencia.Visible = False
    Me.TxtMontoPagado.Text = Format(Val(Format(Me.lblTotal.Caption, "###0.000")) + Val(Format(Me.lblPercepcion.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000")), "###0.000")
    
    If KEY_LINEA_CREDITO = "si" Then
    strCadena = "SELECT cod_unico FROM entidad_empresa WHERE cod_unico='" & Trim(Me.TxtCodCliente.Text) & "' AND id_empresa='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
    
    
    
    'strCadena = "SELECT id_venta,fecha_emision,fecha_vencimiento,hora,numero,comprobante,id_cliente,ncliente,total,saldo,anulado,id_moneda,tc,id_alm,id_doc," & _
    '" id_proyecto,vendedor as nombre_completo,descripcion,simbolo,id_forma_pago,referencia,function_pago_factura(id_venta,'" & Format(Me.DtpHasta.Value, "YYYY-mm-dd") & "',id_moneda,ruc) as pago " & _
    " FROM view_listado_comprobante_vitekey WHERE anulado='no' and  id_cliente LIKE '%" & Trim(Me.txtruc.Text) & "%' AND ruc='" & KEY_RUC & "'"
    
    
      If Me.DtcTipoDoc.BoundText = "0007" Then ' nota credito
          GoTo procesarnota
      End If
         
        
        
        tcredito = load_credito_dispobible(Trim(Me.TxtCodCliente.Text))
        If tcredito < Val(Me.TxtMontoPagado.Text) Then
            Me.lblSaldodisponible.Visible = True
            Me.lblSaldodisponible.ForeColor = &HC0&
            Me.lblSaldodisponible.Caption = "SOBREPASA EL LIMITE DE CREDITO"
            Me.DtcFormapagodetalle.Visible = False
            Me.TxtMontoPagado.Enabled = False
            Exit Sub
        Else
            Me.lblSaldodisponible.Visible = False
            Me.DtcFormapagodetalle.Visible = True
            Me.TxtMontoPagado.Enabled = True
        End If
        
    Else
procesarnota:
        strCadena = "call put_entidad_empresa('" & Trim(Me.TxtCodCliente.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        Me.lblSaldodisponible.ForeColor = &HFF&
        Me.lblSaldodisponible.Caption = "NO AUTORIZADO"
        
    End If
    Set rstT = Nothing
End If
  End If
    
    If KEY_CONTABILIDAD = "si" Then
       strCadena = "SELECT id_registro as Codigo, CONCAT(descripcion,'-',observacion) as Descripcion FROM forma_pago_detalle  WHERE id_alm='" & KEY_ALM & "' and  id_moneda='" & Me.DtcMoneda.BoundText & "' and  id='" & Me.DtcFormaPago.BoundText & "' AND ruc='" & KEY_RUC & "' AND estado='si' ORDER BY id_detalle"
    Else
       strCadena = "SELECT id_registro as Codigo, CONCAT(descripcion,'-',observacion) as Descripcion FROM forma_pago_detalle WHERE  id_alm='" & KEY_ALM & "' and id_moneda='" & Me.DtcMoneda.BoundText & "' and  id='" & Me.DtcFormaPago.BoundText & "' AND ruc='" & KEY_RUC & "' AND estado='si' ORDER BY id_detalle"
    End If
    
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcFormapagodetalle)
    Call Formapagodetalle

Exit Sub
salir:


End Sub
Private Sub DtcFormaPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Call put_forma_pago
    
End If
End Sub
Private Sub put_forma_pago()
    Call forma_pago
    If Me.DtcFormapagodetalle.Visible = True Then
        If Me.DtcFormapagodetalle.Enabled = True Then
            Me.DtcFormapagodetalle.SetFocus
        Else
            MsgBox "No tiene Configurado Formas de Pago. " + Chr(13) + "Mantenimiento Formas pago.", vbInformation
        End If
    Else
        Me.DtcFormaPago.SetFocus
    End If
End Sub

Private Sub Formapagodetalle()
Dim tTotal As Double
If (Trim(Me.DtcFormapagodetalle.BoundText) = "01") Then
    Me.TxtOperacion.Visible = False
    
    
    Me.lbltargeta.Visible = False
    Me.DtTargeta.Visible = False
    Me.TxtNumeroTargeta.Visible = False
    Me.DtpFechaReferencia.Visible = False
    Me.TxtMontoPagado.BackColor = &H80FFFF
    Me.TxtMontoPagado.Text = Format(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000")), "###0.000")
    
    Exit Sub
ElseIf (Trim(Me.DtcFormapagodetalle.BoundText) = "03" Or Trim(Me.DtcFormapagodetalle.BoundText) = "04") Then
    tTotal = Val(Format(Me.lblTotal.Caption, "###0.000"))
    pagado = Val(Format(Me.lblPago.Caption, "###0.000"))
    
    Me.TxtMontoPagado.Visible = True
    
    Me.lbltargeta.Visible = True
    Me.DtTargeta.Visible = True
    Me.TxtNumeroTargeta.Visible = True
    Me.DtpFechaReferencia.Visible = False
    Me.TxtMontoPagado.Text = Format(tTotal - pagado, "###0.000")
    Me.DtpFechaReferencia.Value = Date
    
    Exit Sub
ElseIf (Trim(Me.DtcFormapagodetalle.BoundText) = "02") Then
    Me.TxtOperacion.Visible = False
    Me.TxtMontoPagado.Visible = False
    Me.lbltargeta.Visible = False
    Me.DtTargeta.Visible = False
    Me.TxtNumeroTargeta.Visible = False
    Me.DtpFechaReferencia.Visible = False
    strCadena = "SELECT sum(monto_real) FROM gigabanck WHERE dni='" & Trim(Me.TxtCodCliente.Text) & "'"
    Call ConfiguraRst(strCadena)
    If IsNull(rst(0)) = False Then
        
       
       
    
    Exit Sub
    Else
    Exit Sub
    End If
ElseIf (Trim(Me.DtcFormapagodetalle.BoundText) = "07") Then
    strCadena = "SELECT * FROM  entidad_empresa WHERE cod_unico='" & Trim(Me.TxtCodCliente.Text) & "' ANd id_empresa='" & KEY_RUC & "' AND id_credito='si'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        
        
        
        
    Else
        
        
        
        
        Me.TxtMontoPagado.Visible = False
        'Me.cmdConsultar.Picture = LoadPicture(App.Path + "/Imagenes/noprocede.jpg")
    End If

ElseIf (Trim(Me.DtcFormapagodetalle.BoundText) = "08") Then
        strCadena = "SELECT * FROM  entidad_empresa WHERE cod_unico='" & Trim(Me.TxtCodCliente.Text) & "' ANd id_empresa='" & KEY_RUC & "' AND id_credito='si'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            
            
            
            Me.TxtMontoPagado.Visible = False
            Me.PanelCredito.Visible = False
        Else
        Me.PanelCredito.Visible = True
        Me.lblclave.Visible = False
        Me.TxtClaveRandonCredito.Visible = False
        Me.TxtCuotas.Visible = True
        
        
        
        Me.TxtOperacion.Visible = False
        
        Me.lbltargeta.Visible = False
        Me.DtTargeta.Visible = False
        Me.TxtNumeroTargeta.Visible = False
        Me.DtpFechaReferencia.Visible = False
        End If
        

End If

End Sub

Private Sub DtcFormapagodetalle_KeyPress(KeyAscii As Integer)
Dim in_codigo As String
Me.frmalmacen_entrega(1).Visible = False
If KeyAscii = 13 Then
    in_codigo = get_forma_pago_detalle(Me.DtcFormapagodetalle.BoundText)
    
    If in_codigo = "01" Then
        Me.TxtMontoPagado.Visible = True
        Me.TxtMontoPagovitekey.Visible = False
       ' strCadena = "SELECT sum(monto) FROM movimiento_venta_monto_temporal WHERE id_forma_pago='" & Me.DtcFormapagodetalle.BoundText & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Me.DtcSerieDoc.BoundText & "' ANd numero='" & Me.TxtNumeroDoc.Text & "' AND ruc='" & KEY_RUC & "'"
       ' Call ConfiguraRst(strCadena)
       ' If IsNull(rst(0)) = False Then
        '    Me.TxtMontoPagado.Text = Format(Val(Format(rst(0), "###0.000")) + Format(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000")), "###0.00"))
       ' Else
         '   Me.TxtMontoPagado.Text = Format(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000")), "###0.00")
       ' End If
        Call Resalta(Me.TxtMontoPagado)
        Exit Sub
    End If
    
    If in_codigo = "02" Then
    
    Me.TxtOperacion.Visible = False
    Me.TxtMontoPagado.Visible = False
    Me.lbltargeta.Visible = False
    Me.DtTargeta.Visible = False
    Me.TxtNumeroTargeta.Visible = False
    Me.DtpFechaReferencia.Visible = False
    strCadena = "SELECT sum(monto_real) FROM gigabanck WHERE dni='" & Trim(Me.TxtCodCliente.Text) & "'"
    Call ConfiguraRst(strCadena)
    If IsNull(rst(0)) = False Then
   
        Exit Sub
    Else
        Exit Sub
    End If
    End If
    
    If in_codigo = "03" Or in_codigo = "04" Then
        Me.DtTargeta.Visible = True
        Me.DtTargeta.Enabled = True
        Me.DtTargeta.SetFocus
        Exit Sub
    End If
    
    If in_codigo = "08" Then
        Me.frmvencimiento.Visible = True
        Me.PanelCredito.Visible = False
        
        Me.txtFecha_vencimiento.Text = DateAdd("d", KEY_DIAS_CREDITO, KEY_FECHA)
        'Call Resalta(Me.TxtMontoPagado)
        'Call Resalta(Me.txtFecha_vencimiento)
        Me.txtFecha_vencimiento.SetFocus
        Exit Sub
    End If
    If in_codigo = "06" Then
        Call Resalta(Me.TxtMontoPagado)
        Exit Sub
    End If
    If in_codigo = "09" Then
        
        Me.TxtOperacion.Visible = True
        Call Resalta(Me.TxtOperacion)
        'Call Resalta(Me.TxtMontoPagado)
        Exit Sub
    End If
    
    If in_codigo = "10" Then
    
        Call llenarGrid_recibos(Me.HfRecibos, Trim(Me.TxtCodCliente.Text))
        Me.fraApp.Visible = True
        Exit Sub
    End If
    
    If in_codigo = "12" Then
        strCadena = "SELECT * FROM entidadfinanciera   ORDER BY descripcion"
        Call llenar_bancos(Me.HfBancos)
        Me.frmbanco.Visible = True
        Exit Sub
    End If
    
    If in_codigo = "13" Then
       Me.frmalmacen_entrega(1).Visible = True
       Me.txtserie_nota.Text = ""
       Me.txtNumero_nota.Text = ""
       Me.Label22(3).Caption = 0
       Call Resalta(Me.txtserie_nota)
       Exit Sub
    End If
    
End If
End Sub




Private Sub DtcMoneda_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    Me.DtcFormaPago.SetFocus
End If
End Sub



Private Sub DtcMotor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT nro_chasis,anio_fabricacion FROM imp_producto_detalle WHERE id_detalle='" & Val(Me.DtcMotor.BoundText) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.txtA�oFabricacion.Text = rst("anio_fabricacion")
        
        
        strCadena = "SELECT Codigo,Descripcion FROM view_producto_serie WHERE motor='" & Trim(Me.DtcMotor.Text) & "'"
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcSerie)
        
        
        
        
        Exit Sub
    End If
End If
End Sub
Private Sub parametros_busqueda()
strCadena = "SELECT * FROM parametros_produccion WHERE (codigo='motor' or codigo='chasis') and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       If rst("codigo") = "motor" Then
          Me.lblmotor.Caption = rst("descripcion") & " :"
       End If
       If rst("codigo") = "chasis" Then
           Me.lblchasis.Caption = rst("descripcion") & ":"
       End If
 
       rst.MoveNext
   Next i
   
End If
End Sub
Private Sub DtcSerie_KeyPress(KeyAscii As Integer)
Dim nanio_modelo As String
Dim in_detalle As String
If KeyAscii = 13 Then
   
   strCadena = "select id_producto,id_detalle_serie,detalle from temporal_ventas WHERE id='" & Val(Me.txtid_temporal_serie.Text) & "' and ruc='" & KEY_RUC & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      in_producto = rst("id_producto")
      in_detalle_serie = rst("id_detalle_serie")
      in_detalle = rst("detalle")
   End If
    
    strCadena = "SELECT nro_chasis,anio_fabricacion,nro_contenedor,item,anio_modelo,anio_contenedor,poliza,ip FROM imp_producto_detalle WHERE id_detalle='" & Val(Me.DtcSerie.BoundText) & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.txtA�oFabricacion.Text = rst("anio_fabricacion")
        Me.txtdua.Text = rst("nro_contenedor")
        Me.txtitem.Text = rst("item")
        nanio_modelo = rst("anio_modelo")
        nanio_dua = rst("anio_contenedor")
        npoliza = rst("poliza")
        nip = rst("ip")
        strCadena = "SELECT Codigo,motor as Descripcion FROM view_producto_serie WHERE codigo='" & Trim(Me.DtcSerie.BoundText) & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcMotor)
        
        If KEY_ENVIO_SUNARP = "si" Then
            If KEY_RUC = "20480516771" Then
                in_detalle = get_producto(in_producto) & Space(1) & "  [ " & in_chasis & Trim(Me.DtcSerie.Text) & " ]"
            Else
                in_detalle = get_producto(in_producto) & Space(1) & "  [ " & in_chasis & Trim(Me.DtcSerie.Text) & Space(1) & in_motor & Trim(Me.DtcMotor.Text) & " ]"
            End If
        Else
           If in_detalle = "" Then
                in_detalle = get_producto(in_producto)
           End If
           
        End If
        
        strCadena = "SELECT * FROM temporal_ventas WHERE ruc='" & KEY_RUC & "' and dni_save='" & KEY_USUARIO & "' and id_doc='" & Me.DtcTipoDoc.BoundText & "' and  id_detalle_serie='" & Val(Me.DtcSerie.BoundText) & "' LIMIT 1"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            MsgBox "SERIE YA INGRESADA PARA ESTA VENTA." + Chr(13) + "SELECCIONE OTRA SERIE", vbInformation, KEY_VENDEDOR
        Else
             strCadena = "UPDATE temporal_ventas SET id_detalle_serie='" & Val(Me.DtcSerie.BoundText) & "',detalle='" & in_detalle & "', serie='" & Trim(Me.DtcMotor.Text) & "',anio_fabricacion='" & Trim(Me.txtA�oFabricacion.Text) & "',nro_chasis='" & Me.DtcSerie.Text & "',anio_modelo='" & Trim(nanio_modelo) & "',anio_dua='" & nanio_dua & "',nro_dua='" & Trim(Me.txtdua.Text) & "',nro_item='" & Trim(Me.txtitem.Text) & "',poliza='" & npoliza & "',ip='" & nip & "' WHERE id='" & Val(Me.txtid_temporal_serie.Text) & "'"
             CnBd.Execute (strCadena)
        End If
            Call Me.llenarGrid_det(Me.HfdDetalle, Trim(Me.TxtNumeroDoc.Text), Trim(Me.DtcSerieDoc.BoundText), Me.DtcTipoDoc.BoundText, Trim(Me.txtformato_impresion.Text))
            Me.DtcFormaPago.SetFocus
            Exit Sub
    End If
End If
End Sub

Private Sub put_update_temporal(ByVal in_moneda_ini As String, ByVal in_moneda_fin As String)
On Error GoTo salir
strCadena = "call put_update_temporal_v15('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcSerieDoc.BoundText & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & in_moneda_ini & "','" & in_moneda_fin & "','" & Val(Me.TxtTipoCambio.Text) & "','" & Me.txt_tipo_movimiento.Text & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

If in_moneda_ini <> in_moneda_fin Then
   Call llena_pagos(Me.HfgTipoPagos, Me.TxtNumeroDoc.Text)
   Call Me.llenarGrid_det(Me.HfdDetalle, Trim(Me.TxtNumeroDoc.Text), Trim(Me.DtcSerieDoc.BoundText), Me.DtcTipoDoc.BoundText, Trim(Me.txtformato_impresion.Text))
End If

Exit Sub
salir:


End Sub
Private Sub comprobante(ByVal id_doc As String)
Dim serieA As String
Dim id_docA As String
Dim comproa As String

Me.DtcTipoDoc.Enabled = True
serieA = Trim(Me.DtcSerieDoc.BoundText)
id_docA = Me.DtcTipoDoc.BoundText

If KEY_COMPROBANTES_PROPIOS = "si" Then
    strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND  id_alm='" & KEY_VENTANILLA & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
Else
    strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND  id_alm='" & KEY_ALM & "' AND   ruc='" & KEY_RUC & "' ORDER BY serie ASC LIMIT 1"
End If
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    Call get_serie_comprobante(Me.DtcSerieDoc, Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText)
    Me.DtcSerieDoc.BoundText = rstT("serie")
    Me.TxtNumeroDoc.Text = rstT("numero")
    Me.txtafectacaja.Text = rstT("afecta_caja")
    
    Me.txt_tipo_movimiento.Text = rstT("tipo_movimiento")
    Me.txtformato_impresion.Text = rstT("id_formato_impresion")
    Me.txtserial.Text = rstT("serial")
    Me.DtcMoneda.BoundText = rstT("id_moneda")
    
    
    
    If (Trim(Me.DtcTipoDoc.BoundText) = "0001") Then
        If Me.TxtCodCliente.Locked = True Then
            Me.TxtCodCliente.Locked = False
            Call Resalta(Me.TxtCodCliente)
            Exit Sub
        Else
            Me.TxtCodCliente.Locked = False
        End If
        Call Resalta(Me.TxtCodCliente)
            Exit Sub
    Else
        If (Me.DtcAlmacen.Enabled = True) Then
            If Me.TxtCodProducto.Enabled = True Then
               Call Resalta(Me.TxtCodCliente)
            End If
        End If
    End If
Else
                Me.TxtNumeroDoc.Text = ""
                 
                Call get_serie_comprobante(Me.DtcSerieDoc, Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText)
                
                If (Trim(Me.DtcTipoDoc.BoundText) = "0001") Then
                    If Me.TxtCodCliente.Locked = True Then
                        Call Resalta(Me.TxtCodCliente)
                    End If
                Else
                    If (Me.DtcAlmacen.Enabled = True) Then
                        If Me.TxtCodProducto.Enabled = True Then
                            Call Resalta(Me.TxtCodProducto)
                        End If
                    End If
                End If
                
    End If

End Sub

Private Sub DtcSerieDoc_Change()

Dim in_moneda_inicio As String
Dim in_moneda_destino As String

in_moneda_inicio = Me.DtcMoneda.BoundText

Me.TxtNumeroDoc.Text = get_comprobante_numero_v2(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText))

in_moneda_destino = Me.DtcMoneda.BoundText

Call put_update_temporal(in_moneda_inicio, in_moneda_destino)
End Sub


Public Function get_comprobante_numero_v2(ByVal in_doc As String, ByVal in_serie As String) As String
On Error GoTo salir


strCadena = "SELECT numero,id_moneda,afecta_caja,id_formato_impresion,igv FROM almacen_comprobante WHERE id_doc='" & in_doc & "' and serie='" & in_serie & "' and ruc='" & KEY_RUC & "' limit 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
   get_comprobante_numero_v2 = rstL("numero")
   Me.DtcMoneda.BoundText = rstL("id_moneda")
   Me.txtafectacaja.Text = rstL("afecta_caja")
   Me.txtformato_impresion.Text = rstL("id_formato_impresion")
   KEY_APLICA_IGV = rstL("igv")
Else
   get_comprobante_numero_v2 = ""
End If

Exit Function
salir:

End Function


Private Sub DtcTipoDoc_Change()
Dim in_moneda_ini As String
Dim in_moneda_fin As String

If Me.DtcTipoDoc.Enabled = True Then
       in_moneda_ini = Me.DtcMoneda.BoundText
       Call comprobante(Me.DtcTipoDoc.BoundText)
       If Me.DtcTipoDoc.BoundText = "0007" Or Me.DtcTipoDoc.BoundText = "0008" Then
          
            If Me.DtcTipoDoc.BoundText = "0007" Then
                Call load_tipo_nota
            End If
          
          
            If Me.DtcTipoDoc.BoundText = "0008" Then
                Call load_tipo_debito
            End If
          
        Else
            Me.frm_motivo_nota.Visible = False
       End If
       in_moneda_fin = Me.DtcMoneda.BoundText
       
       Call put_update_temporal(in_moneda_ini, in_moneda_fin)
 End If
End Sub
Private Sub load_tipo_nota()
strCadena = "SELECT id_tipo_nota as Codigo,descripcion as Descripcion FROM tipo_nota_credito ORDER BY id_tipo_nota"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoNota)
Me.frm_motivo_nota.Visible = True
End Sub
Private Sub load_tipo_debito()
strCadena = "SELECT id_tipo_nota as Codigo,descripcion as Descripcion FROM tipo_nota_debito ORDER BY id_tipo_nota"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoNota)
Me.frm_motivo_nota.Visible = True
End Sub

Private Sub DtcTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Me.DtcAlmacen.SetFocus
End If
If KeyCode = vbKeyRight Then
    'Call Resalta(Me.TxtSerie)
End If
End Sub

Private Sub DtcTipoDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
  Call comprobante(Me.DtcTipoDoc)
End If

End Sub






Private Sub DtcUnidad_Change()

Call precio_unidad(Trim(Me.txtagranel.Text))

End Sub
Public Sub precio_unidad(ByVal in_agranel As String)
 
 If in_agranel = "si" Then
    Me.txtpreciooriginal.Text = get_precio_unidad(Trim(Me.TxtCodProducto.Text), FrmVentas.DtcUnidad.BoundText)
    Me.txtprecio.Text = Val(Me.txtpreciooriginal.Text)
    Me.LblTotalParcial.Caption = Format(Val(Me.txtCantidad.Text) * Val(Me.txtprecio.Text), "###0.00")
 End If


End Sub

Private Sub DtcUnidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtprecio)
End If
End Sub

Private Sub DtTargeta_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        Me.TxtNumeroTargeta.Visible = True
        Me.TxtMontoPagado.Text = Format(Val(Format(Me.lblTotal.Caption, "###0.00") - Val(Format(Me.lblPago.Caption, "###0.00"))), "###0.00")
        Me.TxtOperacion.Visible = True
        Call Resalta(Me.TxtNumeroTargeta)
        Me.frminteres.Visible = True
        
End If
End Sub

Sub save_temporal()
Dim codigo As String
Dim hora As Date
Dim id_codigo As Double
strCadena = "INSERT INTO temporal_venta_guardado(fecha,hora,cliente,dni_save,monto_guardado,ruc)VALUES('" & KEY_FECHA & "','" & str(Time) & "','" & Trim(Me.TxtCliente.Text) & "','" & KEY_USUARIO & "','" & Val(Me.lblTotal.Caption) & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
'
id_codigo = LastRegistro("temporal_venta_guardado", "id_codigo")

strCadena = "UPDATE temporal_ventas SET save='si',id_medic='" & id_codigo & "' WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' AND id_medic='0'"
CnBd.Execute (strCadena)
'
Call nuevo
End Sub


Private Sub Form_Activate()



If KEY_AUTOMATICO = "si" Then
    If Me.OptAuto.Value = True Then
        Me.OptAuto.Value = True
    End If
Else
    Me.OptManual.Value = True
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 
 If (KeyCode = 122) Then
     'Me.DtcMoneda.SetFocus
     Me.DtcFormaPago.SetFocus
     Exit Sub
 End If
 If KeyCode = 123 Then
    Call Resalta(Me.TxtCodCliente)
    Exit Sub
 End If
 If (KeyCode = 112) Then
    Call Resalta(Me.TxtCodCliente)
    Exit Sub
 End If
 
 If KeyCode = 115 Then
    If Me.DtcTipoDoc.Enabled = True Then
        Me.DtcTipoDoc.SetFocus
        Exit Sub
    End If
 End If
 If (KeyCode = 114) Then
    Procedencia = buscar
    frmBuscardoc.Show
     Exit Sub
 End If
 
 If (KeyCode = 113) Then
    If Val(Me.lblTotal.Caption) > 0 Then
     Call save_temporal
    
    End If
     Exit Sub
 End If
 
 
 

If KeyCode = 120 Then
    
    Call validar_facturacion
    'If put_dni_boleta(Trim(Me.TxtCodCliente.Text), Val(Me.lblTotal.Caption), Me.DtcTipoDoc.BoundText) = False Then
    '    Call Resalta(Me.TxtCodCliente)
    '    Exit Sub
    'End If

    'If validar_comprobante_electronico(Me.DtcTipoDoc.BoundText, Trim(Me.TxtCodCliente.Text)) = True Then
    '    Call facturar
    'End If
    
    Exit Sub
  End If


If KeyCode = 117 Then
    Call nuevo
    Exit Sub
  End If

If Shift = 2 And KeyCode = Asc("A") Then
    If MsgBox(MSGANULAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Procedencia = anular
        FrmSeguridad.Show
        Exit Sub
       End If
End If
  
If Shift = 2 And KeyCode = Asc("E") Then
    If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Procedencia = Eliminar
        FrmSeguridad.Show
       End If
End If


If Shift = 2 And KeyCode = Asc("D") Then
    If Me.chkDelivery.Value = 1 Then
        Me.chkDelivery.Value = 0
    Else
    Me.chkDelivery.Value = 1
    End If
    Exit Sub
End If
If KeyCode = 119 Then
     
     If strEspecial > 1 Then
        Call OrdenImpresionEspecial
        Else
            Call OrdenImpresion(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text))
        End If
End If
'If Shift = 2 And KeyCode = Asc("I") Then
'     Call Imprimir(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcAlmacen.BoundText), Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text))
'End If
  
End Sub
Private Sub Form_Load()
CenterForm Me




Me.Top = 50
 delivery = "no"
 Me.DtpIni.Value = KEY_FECHA
 Me.DtpFin.Value = KEY_FECHA
 
  Me.TxtTipoCambio.Text = KEY_CAMBIO_LOCAL
 
 If KEY_USUARIO = "42546269" Or KEY_USUARIO = "46947665" Or KEY_USUARIO = "900001" Then
    Me.frmmanual.Visible = True
 Else
    Me.frmmanual.Visible = False
 End If
 
 If KEY_SKFACTURA = "si" Then
    Me.chk_factura.Visible = True
  Else
    Me.chk_factura.Visible = False
 End If
 
 If KEY_COMPROBANTE_ADICIONAL = "si" Then
    Me.frameTramite.Visible = True
 Else
    Me.frameCajaIndependiente.Visible = False
 End If
 
 If KEY_CAJA_INDEPENDIENTE = "si" Then
    Me.frameCajaIndependiente.Visible = True
 Else
    Me.frameCajaIndependiente.Visible = False
 End If
 
 If KEY_VALIDACION_EXTREMA = "si" Then
    Me.TxtDescripcionProducto.Locked = True
 Else
        Me.TxtDescripcionProducto.Locked = False
 End If
 
If KEY_GRIFO = "si" Then
    Me.txtObservacion.Text = "PLACA N�:"
 End If
dfactura = False
Me.DtpActual.Value = CVDate(KEY_FECHA)
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen  WHERE id_sucursal='0' and  ruc='" & KEY_RUC & "'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  Me.DtcAlmacen.Enabled = False
  Me.TxtCodProducto.Enabled = False
 
  
  
  
    strCadena = "SELECT id as Codigo, descripcion as Descripcion FROM targeta WHERE id<>'00' ORDER BY id ASC"
  
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtTargeta)
  
  strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM moneda  ORDER BY id_moneda ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMoneda)
 
  strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad  WHERE  id_personal='si' and habilitado='si' and  ruc='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcVendedor)
  Me.DtcVendedor.BoundText = 0
  

strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM forma_pago ORDER BY id ASC "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFormaPago)
Me.DtcFormaPago.BoundText = "01"



If KEY_SEGURO_VENTA = "si" Then
   Me.chkseguro.Visible = True
   Me.Dtcseguro.Visible = True
End If




  If KEY_COMPROBANTES_PROPIOS = "si" Then
    strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM view_almacen_comprobante WHERE id_alm='" & KEY_VENTANILLA & "' and ruc='" & KEY_RUC & "'"
  Else
    strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM view_almacen_comprobante WHERE id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
  End If
  
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
        Call LlenaDataCombo(Me.DtcTipoDoc)
        Me.DtcTipoDoc.Enabled = False
        Me.DtcTipoDoc.BoundText = 0
  End If
  
  
  
  
  Me.DtcTipoDoc.Enabled = False
  Me.DtcSerieDoc.Enabled = False
  Me.TxtNumeroDoc.Enabled = False
  
  Me.DTPDetracion.Value = CVDate(KEY_FECHA)
  Me.cmdProcesar.Enabled = False
  Me.cmdImprimir.Enabled = False
  Me.cmdEliminar.Enabled = False
  Me.cmdAnular.Enabled = False
 
  Set cQrCode = New ClsQrCode
  Call parametro_importacion
 
End Sub

Private Sub HfBancos_Click()
    Me.txtBanco.Text = Me.HfBancos.TextMatrix(Me.HfBancos.Row, 1)
     Call Resalta(Me.txtCheque)
    Exit Sub
End Sub

Private Sub HfBancos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call Resalta(Me.TxtMontoPagado)
   Exit Sub
End If
End Sub

Private Sub HfdDetalle_DblClick()
If Me.HfdDetalle.Rows > 0 Then


        
        
        
        If Me.DtcTipoDoc.BoundText = "0099" And KEY_UPDATE_PROFORM = "si" And Val(Me.txtidVenta.Text) > 0 And Me.cmdModificar.Enabled = False Then
            If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
            FrmVentaCantidad.Show
            strCadena = "SELECT * FROM temporal_ventas WHERE id='" & Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
            FrmVentaCantidad.lbl_obsequio.Caption = rst("obsequio")
            FrmVentaCantidad.lblprecio_original.Caption = rst("precio")
            End If
            
            
            Exit Sub
        End If
        End If
        
        If (Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))) > 0 And (Me.cmdProcesar.Enabled = True) Then
            FrmVentaCantidad.Show
            strCadena = "SELECT * FROM temporal_ventas WHERE id='" & Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
            FrmVentaCantidad.lbl_obsequio.Caption = rst("obsequio")
            FrmVentaCantidad.lblprecio_original.Caption = rst("precio")
            End If
            Exit Sub
        End If
        
        
        

End If
End Sub

Private Sub HfdDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
   If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    Call Quitar(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))
    Call Resalta(Me.TxtCodProducto)
End If

End If
End Sub

Private Sub HfdDetalle_SelChange()
On Error GoTo salir

If Me.HfdDetalle.Rows > 0 Then

If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    Me.chkPrecios.Enabled = True
     Me.cmdConversion.Enabled = True
    strCadena = "SELECT produccion FROM producto P,linea L WHERE P.id_linea=L.id_linea AND P.id_producto='" & Trim(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 1)) & "' AND P.ruc=L.id_usu AND P.ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.cmdSeriales.Visible = True
    Else
        cmdSeriales.Visible = False
    End If
    Exit Sub
Else
    Me.cmdSeriales.Visible = False
    Me.cmdConversion.Enabled = False
    Exit Sub
End If
   
Else
    Me.cmdConversion.Enabled = False
End If
Exit Sub
salir:


End Sub

Private Sub hfdireccion_Click()
Call select_direccion
End Sub
Private Sub select_direccion()
Me.hfdireccion.TextMatrix(Me.hfdireccion.Row, 3) = Chr(254)
Me.txtDireccion.Text = Me.hfdireccion.TextMatrix(Me.hfdireccion.Row, 2)
Me.frmdireccion.Visible = False
End Sub

Private Sub HfFacturas_DblClick()
If Me.HfFacturas.Rows > 0 Then
    Call disabled_form(Me)
    Procedencia = buscar
    frmdetalle.Show
    Exit Sub
End If
End Sub

Private Sub HfgTipoPagos_Click()
If Me.HfgTipoPagos.Rows > 0 Then
If Val(Me.HfgTipoPagos.TextMatrix(Me.HfgTipoPagos.Row, 0)) > 0 Then
    Me.cmdQuitarMonto.Visible = True
Else
    Me.cmdQuitarMonto.Visible = False
End If
End If
End Sub

Private Sub HfgTipoPagos_SelChange()
If Val(Me.txtidVenta.Text) > 0 Then
    Me.cmdQuitarMonto.Visible = False
Else
    Me.cmdQuitarMonto.Visible = True
End If

End Sub
Public Function get_ultimo_precio(ByVal in_cliente As String, ByVal in_producto As String) As String
strCadena = "SELECT * FROM view_producto_ultima_venta WHERE id_cliente='" & in_cliente & "' and id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   get_ultimo_precio = Format(rstA("fecha_emision"), "dd-mm-YYYY") & " | " & rstA("documento") & "     |     " & Format(rstA("precio"), "#,##0.00")
Else
   get_ultimo_precio = "..."
End If
End Function


Private Sub mscConecta_OnComm()

End Sub

Private Sub ActualizarImagen(ByVal Grilla As MSHFlexGrid, ByVal Fila As Integer)
     Dim estado As String
      
            For i = 1 To Grilla.Rows - 1
                Grilla.TextMatrix(i, 0) = Chr(168)
            Next i
            
               If Val(Grilla.TextMatrix(Grilla.Row, 3)) > Val(Me.txtCantidad.Text) Then
                    MsgBox "APLICA ESTE PRECIO APARTIR DE:" + Space(1) + str(Grilla.TextMatrix(Grilla.Row, 3)) + Chr(13) + "RECUERDE ESTA SIENDO NOTIFICADO.", vbInformation, KEY_VENDEDOR
                    Exit Sub
               End If
               
            
               Grilla.TextMatrix(Grilla.Row, 0) = Chr(254)
               Me.txtprecio.Text = Format(Me.HfPrecios.TextMatrix(Me.HfPrecios.Row, 1), "###0.00")
               Me.LblTotalParcial.Caption = Format(Val(Me.txtCantidad.Text) * Val(Me.txtprecio.Text), "###0.00")
               Me.txtprecio.Locked = False
               Me.txtmayor.Text = "si"
               
               Call insertar_codigo
                'Me.CmdAgregar.SetFocus
               'Call Resalta(Me.txtprecio)
                Exit Sub
            'For j = 0 To 2
            '    HfLinea.col = j
             '   HfLinea.Row = Me.HfLinea.Row
              '  HfLinea.CellBackColor = &HC0FFC0
            'Next j
        
       ' strCadena = "UPDATE linea_medico SET estado='" & estado & "' WHERE id_linea='" & id_linea & "' AND dni='" & dni & "' AND ruc='" & KEY_RUC & "'"
        'CnBd.Execute (strCadena)  '
      
      
      
      
End Sub

Private Sub ActualizarPendiente(ByVal Grilla As MSHFlexGrid, ByVal Fila As Integer)
     Dim estado As String
     Dim in_fila As String
     in_fila = Me.HfPendientes.Row
            For i = 1 To Me.HfPendientes.Rows - 1
                Grilla.TextMatrix(i, 4) = Chr(168)
                For j = 1 To 4
                Grilla.col = j
                Grilla.Row = i
                Grilla.CellBackColor = &HFFFFFF
                Next j
        
            Next i
            
            
            
            
            Grilla.TextMatrix(in_fila, 4) = Chr(254)
                
        
            For j = 1 To 4
                Grilla.col = j
                Grilla.Row = in_fila
                Grilla.CellBackColor = &H80FF&
            Next j
        
        
      
      
     '&HFFFFFF
      
End Sub

Private Sub HfPendientes_DblClick()

If Me.HfPendientes.Rows > 0 Then
    Me.txt_id_pendiente.Text = Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0))
    If Val(txt_id_pendiente) > 0 Then
        Call ActualizarPendiente(Me.HfPendientes, 3)
        Call get_comprobante(Val(txt_id_pendiente))
        Call llena_pagos(Me.HfgTipoPagos, Me.TxtNumeroDoc.Text)
        Me.timer_pendientes.Enabled = False
    End If
End If
End Sub

Private Sub HfPendientes_SelChange()
If Me.HfPendientes.Rows > 0 Then
If Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) > 0 Then
    Me.cmddescartar.Enabled = True
Else
    Me.cmddescartar.Enabled = False
End If
End If
End Sub

Private Sub HfPrecios_Click()
Call ActualizarImagen(Me.HfPrecios, Me.HfPrecios.Row)
End Sub

Private Sub lblPago_Change()
If Val(Me.TxtTipoCambio.Text) > 0 Then
    If Me.DtcMoneda.BoundText = "00001" Then
       Me.lblConversion(1).Caption = Format(Val(Me.lblPago.Caption) / Val(Me.TxtTipoCambio.Text), "###0.00")
    Else
       Me.lblConversion(1).Caption = Format(Val(Me.lblPago.Caption) * Val(Me.TxtTipoCambio.Text), "###0.00")
    End If
End If
End Sub

Private Sub lblTotal_Change()
If Val(Me.TxtTipoCambio.Text) > 0 Then
If Me.DtcMoneda.BoundText = "00001" Then
   Me.lblConversion(0).Caption = Format(Val(Me.lblTotal.Caption) / Val(Me.TxtTipoCambio.Text), "#,##0.00")
Else
   Me.lblConversion(0).Caption = Format(Val(Me.lblTotal.Caption) * Val(Me.TxtTipoCambio.Text), "#,##0.00")
End If
End If
End Sub

Private Sub lblVuelto_Change()

If Val(Me.TxtTipoCambio.Text) > 0 Then
        If Me.DtcMoneda.BoundText = "00001" Then
           Me.lblConversion(2).Caption = Format(Val(Me.lblVuelto.Caption) / Val(Me.TxtTipoCambio.Text), "#,##0.00")
        Else
           Me.lblConversion(2).Caption = Format(Val(Me.lblVuelto.Caption) * Val(Me.TxtTipoCambio.Text), "#,##0.00")
        End If
End If

End Sub

Private Sub OptAuto_Click()
If TxtCodProducto.Enabled = True Then
    Call Resalta(Me.TxtCodProducto)
End If
End Sub

Private Sub OptManual_Click()
If Me.TxtCodProducto.Enabled = True Then
    Me.OptManual.Value = True
    Call Resalta(Me.TxtCodProducto)
End If
End Sub

Private Sub sendmail1_Click()

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If Me.SSTab1.Tab = 1 Then
    Call llenarGrid_Facturas(Me.HfFacturas, Trim(Me.TxtCodCliente.Text))
End If
End Sub

Private Sub timer_pendientes_Timer()

If KEY_CAJA_INDEPENDIENTE = "si" And (KEY_CARGO = "00008" Or KEY_CARGO = "00004") Then
    Call buscar_pendientes
End If
End Sub


Public Sub nuevo()
On Error GoTo salir
    
    'VALIDAR SI SE VA A GENERAR UNA NUEVA NOTA DE CREDITO
    'DE UN COMPROBANTE CON NOTA DE CREDITO
    If Me.DtcTipoDoc.BoundText = "0007" Then
       If Me.FrameReferencia.Visible = True Then
          If MsgBox("Desea realizar otra: NOTA DE CREDITO" + Chr(13) + "Para este comprobante. ?", vbQuestion + vbYesNo) = vbYes Then
             Me.FrameReferencia.Visible = False
             Me.txtidVenta.Text = 0
             Me.cmdProcesar.Enabled = True
             Exit Sub
            
          End If
       End If
    End If
    
    
    Me.txtObservacion.Text = ""
    Me.Enabled = True
    Me.lblPercepcion.Caption = ""
    lblfactura_masiva.Caption = "no"
    Me.chk_venta_diferida.Value = 0
    Me.txtFecha_vencimiento.Mask = ""
    Me.txtFecha_vencimiento.Text = ""
    Me.txtFecha_vencimiento.Mask = "##/##/####"
    Me.frmvencimiento.Visible = False
    Me.lblhistorial.Caption = ""
    Me.txt_id_pendiente.Text = ""
    Me.chk_detraccion.Value = 0
    Me.txtOrdenCompra.Text = ""
    Me.chkObsequio.Value = 0
    Me.HfgTipoPagos.Rows = 0
    Me.txtTelefono.Text = ""
    Me.txtmail.Text = ""
    Me.FrameReferencia.Visible = False
    Me.imgFoto.Visible = False
    Me.cmdcredito.Visible = False
    Me.txtmontocredito.Text = 0
    Me.TxtOperacion.Visible = False
    Me.lblpeso.Caption = ""
    Me.lblicbper.Caption = ""
    Me.lblGratuitas.Caption = ""
    cmdCuotas.Visible = False
    Me.lblConversion(0).Caption = ""
    Me.lblConversion(1).Caption = ""
    Me.lblConversion(2).Caption = ""
    Me.TxtTipoCambio.Text = KEY_CAMBIO_LOCAL
    Me.chkconsultar.Value = 0
    Me.chk_direccion.Value = 0
    Me.DtcUnidad.BoundText = 0
    Me.DtcUnidad.Text = ""
    Me.txtmayor.Text = "no"
    Me.IN_ICBPER = "no"
    Me.txtExtranjero.Text = "no"
    Me.TxtDescuento_global.Text = 0
    Me.TxtDescuento_porcentaje.Text = 0
    Me.chk_descuento.Value = 0
    Me.frmdireccion.Visible = False
    Me.DtcFormaPago.BoundText = "01"
    Me.DtcFormapagodetalle.BoundText = "01"
    Me.TxtOperacion.Text = ""
    Me.txt_sunat_key.Text = ""
    Me.txt_hash.Text = ""
    Me.txtidVenta.Text = ""
    Me.txtid_venta_ref.Text = 0
    Me.txttipofactura.Text = "00001"
    Me.txtrecibo_anterior.Text = 0
    Me.timer_pendientes.Enabled = True
    Me.txtA�oFabricacion.Text = ""
    Me.txtModelo.Text = ""
    Me.txtcolor.Text = ""
    in_total_documento = 0
    Me.txtdni_copropietario.Text = ""
    Me.lblcopropietario.Caption = ""
    Me.frmcopropietario.Visible = False
    Me.chkconyuge.Value = 0
    Me.txtbusquedamotor.Text = ""
    Me.DtcSerie.BoundText = ""
    Me.DtcVendedor.BoundText = "0"
    Me.PanelCredito.Visible = False
    Me.TxtMontoPagovitekey.Visible = False
    Me.lblContabilidad.Visible = False
    
    Me.lblDisponible.Caption = ""
    Me.lblDisponible.Visible = False
    Me.DtTargeta.Visible = False
    Me.cmdSeriales.Visible = False
    Me.chkvincular.Visible = False
    
    frmalmacen_entrega(1).Visible = False
    
    
    
    
    Me.fraApp.Visible = False
    Me.txtserieguia.Visible = False
    Me.txtnumeroguia.Visible = False
    Me.cmdGrabarGuia.Visible = False
    Me.cmdImprimirGuia.Visible = False
    Me.DtcVendedor.Locked = False
    Me.txteditable.Text = "no"
    Me.lblregistradopor.Caption = ""
    Me.FrameSerieModelo.Visible = False
    Me.frmcredito.Visible = False
    Me.frminteres.Visible = False
    Me.HfFacturas.Rows = 0
    Me.chk_factura.Value = 0
    Me.txtid_agenda.Text = 0
 
    If KEY_CARGO = "00001" And KEY_RUC = "20603698852" Then
        chkObsequio.Enabled = False
    End If
    
 
    If Me.DtcAlmacen.Enabled = True Then
    
    
     Me.chkDelivery.Value = 0
     Me.cmdVisualizar.Visible = False
     Me.lblPendientes.Caption = ""
     
     strCadena = "call p_nueva_venta_v11('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
     CnBd.Execute (strCadena)
     
     Me.HfgTipoPagos.Rows = 0
     Me.cmdQuitarMonto.Visible = False
     Call get_serie_comprobante(Me.DtcSerieDoc, Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText)
     
     
        If Me.chkconsultar.Value = 1 Then
            Call get_serie_comprobante_alm(Me.DtcSerieDoc, Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText)
        End If
    
     strCadena = "SELECT igv,serie,numero,igv, id_formato_impresion,id_moneda FROM almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Me.DtcSerieDoc.BoundText & "'  AND ruc='" & KEY_RUC & "' LIMIT 1"
     Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        KEY_APLICA_IGV = rst("igv")
        Me.DtcSerieDoc.BoundText = rst("serie")
        Me.TxtNumeroDoc.Text = rst("numero")
        Me.DtcMoneda.BoundText = rst("id_moneda")
        Me.txtformato_impresion.Text = rst("id_formato_impresion")
    End If
    Else
        MsgBox "LA SERIE DE " + Chr(32) & Me.DtcTipoDoc.Text & Space(1) & "NO ESTA REGISTRADA." + Chr(13) + Chr(13) + "REGISTRE EN EL MODULO DE COMPROBANTES �" + Chr(13) + "CAMBIE DE SERIE.", vbInformation, KEY_EMPRESA
        
        Me.cmdAnular.Enabled = False
        Me.cmdEliminar.Enabled = False
        Me.cmdProcesar.Enabled = False
        Me.cmdImprimir.Enabled = False
        Me.HfdDetalle.Rows = 0
        Call Resalta(Me.TxtNumeroDoc)
        Exit Sub
    End If
        Me.HfPrecios.Visible = False
        chkPrecios.Enabled = False
        Me.TxtCodCliente.Text = "00000000"
        Me.TxtCliente.Text = "CLIENTE"
        Me.txtDireccion.Text = KEY_DIR_PUBLIC
        Me.DtcAlmacen.BoundText = KEY_ALM
        Me.DtpActual.Value = KEY_FECHA
        
        Me.txtObservacion.Text = ""
        Me.TxtDescripcionProducto.Text = ""
        Me.TxtCodProducto = "00000"
        Me.TxtDescripcionProducto.Text = ""
        Me.txtprecio.Text = ""
        Me.txtCantidad.Text = 1
        Me.LblTotalParcial.Caption = ""
        Me.lblCantidad.Caption = "0"
        Me.LblTotalLetras.Caption = ""
        Me.lblPago.Caption = ""
        Me.lblVuelto.Caption = ""
        Me.lblSobrante.Caption = ""
        Me.DtcFormaPago.BoundText = "01"
        Me.TxtMontoPagado.BackColor = &HFFFFFF
        Me.TxtMontoPagado.Text = ""
        
        If Me.DtcTipoDoc.BoundText <> "0007" Then
           Me.frm_motivo_nota.Visible = False
        End If
        Me.txtmotivo_nota.Text = ""
        
        chk_descuento.Enabled = True
        Me.lblExonerado.Caption = ""
        Me.TxtMontoPagado.Text = ""
        Me.TxtNumeroTargeta.Text = ""
        Me.TxtCodProducto.Enabled = True
        Me.TxtDescripcionProducto.Enabled = True
        Me.txtCantidad.Enabled = True
        Me.txtprecio.Enabled = True
        Me.cmdAgregar.Enabled = True
        Me.CmdQuitar.Enabled = True
        'Me.cmdEditable.Enabled = True
        Me.cmdAnular.Enabled = False
        Me.cmdEliminar.Enabled = False
        Me.cmdProcesar.Enabled = False
        Me.cmdImprimir.Enabled = False
        Me.HfdDetalle.Rows = 0
        Call Resalta(Me.TxtCodCliente)
        If Me.ChkExtraer.Value = 1 Then
            Me.ChkExtraer.Value = 0
            Me.TxtSeri_guia.Text = ""
            Me.TxtNumero_guia.Text = ""
        End If
    Me.LblIgv.Caption = ""
    Me.LblTotalParcial.Caption = ""
    Me.lblTotal.Caption = ""
    Me.LblValorVenta.Caption = ""
    Me.lblAnulado.Visible = False
    If Me.DtcTipoDoc.BoundText = "0001" Then
        Call Resalta(Me.TxtCodCliente)
    End If
    Me.DtcVendedor.BoundText = KEY_USUARIO
    If KEY_PROYECTO = "si" Then
        chkseguro.Value = 0
         
    End If
    'Call DisplayTextoCom(Space(5) + "BIENVENIDO A:" & AlineaString(Me.lblTotal.Caption, 8, pAlnDerecha) & _
                            "---- " & AlineaString(Me.lblTotal.Caption, 8, pAlnDerecha), mscConecta)
                            
    'Call DisplayTextoCom(AlineaString("PEPES", 20, pAlnCentro, "*") & String$(20, " "), mscConecta)

   ' MsgBox "Active el Almacen Correspondiente", vbInformation, KEY_EMPRESA

Exit Sub
Set rst = Nothing
salir:
MsgBox "Disculpe las Molestias SR." & KEY_VENDEDOR + Chr(13) + Chr(13) + "Ha Ocurrido un Fallo en su Dispositivo.", vbInformation, KEY_EMPRESA

End Sub

Public Sub get_unidad(ByVal in_producto As String, ByVal in_agranel As String)
    If in_agranel = "si" Then
        strCadena = "SELECT id_unidad as Codigo,descripcion as Descripcion FROM view_unidad_producto WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    Else
        strCadena = "SELECT id_unidad as Codigo,descripcion as Descripcion FROM view_unidad WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
    End If
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcUnidad)
End Sub


Sub verifica(ByVal doc_deta As String)
    Select Case Val(doc_deta)
        Case 1
'            Call Doc_Referencia(True, Val(doc_deta))
        Case 3
          '  Call Doc_Referencia(False, Val(doc_deta))
        Case 7
           ' Call Doc_Referencia(True, Val(doc_deta))
        Case 8
            'Call Doc_Referencia(True, Val(doc_deta))
        Case 9
            
            'Call Doc_Referencia(True, Val(doc_deta))
        Case 88
            'Call Doc_Referencia(False, Val(doc_deta))
        Case 89
            'Call Doc_Referencia(False, Val(doc_deta))
        Case 90
            'Call Doc_Referencia(False, Val(doc_deta))
        Case 95
            'Call Doc_Referencia(False, Val(doc_deta))
        Case 96
            'Call Doc_Referencia(True, Val(doc_deta))
    End Select
    
End Sub

Public Sub get_auto_pago(ByVal in_doc As String)
Dim in_cuenta_caja As String
Dim in_forma_pago_detalle As Integer
If in_doc = "0099" Then

strCadena = "SELECT * FROM movimiento_venta_monto_temporal WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' and serie='" & Trim(Me.DtcSerieDoc.BoundText) & "' and id_usuario='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount < 1 Then
'strCadena = "DELETE FROM movimiento_venta_monto_temporal WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' and serie='" & Trim(Me.DtcSerieDoc.BoundText) & "' and id_usuario='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
'CnBd.Execute (strCadena)
in_forma_pago_detalle = get_forma_pago_contado
in_cuenta_caja = get_cuenta_contable_caja(in_forma_pago_detalle)


strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,cuenta_contable,cuotas,id_usuario,id_recibo,detalle,banco,cheque,fecha,id_alm,ruc) VALUES " & _
       " ('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcSerieDoc.BoundText & "','" & Trim(Me.TxtNumeroDoc.Text) & "','01','" & in_forma_pago_detalle & "','" & Me.DtcMoneda.BoundText & "','" & Val(Me.lblTotal.Caption) & "','" & Val(Me.lblTotal.Caption) & "','00','-','-','" & in_cuenta_caja & "','0','" & KEY_USUARIO & "','0','-','-','-','" & KEY_FECHA & "','" & KEY_ALM & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       
'strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,cuotas,id_usuario,fecha,id_alm,ruc) VALUES " & _
       " ('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcSerieDoc.BoundText & "','" & Trim(Me.TxtNumeroDoc.Text) & "','01','" & get_forma_pago_contado & "','" & Val(Me.lblTotal.Caption) & "','" & Val(Me.lblTotal.Caption) & "','00','-','-','0','" & KEY_USUARIO & "','" & KEY_FECHA & "','" & KEY_ALM & "','" & KEY_RUC & "')"
'CnBd.Execute (strCadena)
End If

Me.lblPago.Caption = Format(Val(Me.lblTotal.Caption), "###0.000")
Me.lblVuelto.Caption = Format(0, "###0.00")
End If


End Sub
Public Sub OrdenImpresion(ByVal ndoc As String, ByVal nserie As String, nnumero As String)
On Error GoTo salirr
Dim X As Integer
Dim impresiones As Integer, id_venta As Double

       If KEY_IMPRESION_PROFORMA = "no" And ndoc = "0099" Then
         Exit Sub
       End If



       strCadena = "SELECT * FROM movimiento_venta WHERE numero='" & nnumero & "' AND id_doc='" & ndoc & "' AND id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND serie='" & nserie & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
            If ndoc = "0054" Then
                GoTo imprimir_n
            End If
          If rst("impresiones") < 1 Then
imprimir_n:
              impresiones = rst("impresiones") + 1
              id_venta = rst("id_venta")
              'Call Imprimir_Tiketera(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcAlmacen.BoundText), Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text))
              
              
              If KEY_RUC = "20566449383" And ndoc = "0054" And rst("id_forma_pago") = "02" Then
                Call Orden_Impresion(ndoc, nserie, nnumero, rst("id_tipo_factura"), rst("id_venta"))
                Call Orden_Impresion(ndoc, nserie, nnumero, rst("id_tipo_factura"), rst("id_venta"))
              Else
                
             If KEY_RUC = "20566449383" And Me.DtcTipoDoc.BoundText = "0054" Then
                    If MsgBox("Desea Imprimir este Comprobante ?", vbYesNo + vbQuestion, KEY_VENDEDOR) = vbYes Then
                        Call Orden_Impresion(ndoc, nserie, nnumero, rst("id_tipo_factura"), rst("id_venta"))
                    End If
             Else
                    
                    If KEY_RUC = "10749269729" Then  ' solo en ventas
                        Call Orden_Impresion(ndoc, nserie, nnumero, rst("id_tipo_factura"), rst("id_venta"))
                        Call Orden_Impresion(ndoc, nserie, nnumero, rst("id_tipo_factura"), rst("id_venta"))
                    Else
                    Call Orden_Impresion(ndoc, nserie, nnumero, rst("id_tipo_factura"), rst("id_venta"))
                    End If
                
             End If
             End If
            '  Call Orden_Impresion(ndoc, nserie, nnumero, rst("id_tipo_factura"), rst("id_venta"))
              
              
              strCadena = "UPDATE movimiento_venta SET impresiones='" & impresiones & "' WHERE id_venta='" & id_venta & "' AND ruc='" & KEY_RUC & "'"
              CnBd.Execute (strCadena)
              
              If KEY_RUC = "20601689546" Or ndoc = "20601402433" Then
                Call Orden_Impresion(ndoc, nserie, nnumero, rst("id_tipo_factura"), rst("id_venta"))
              End If
             
               
           Else
              If MsgBox("ESTE DOCUMENTO YA FUE IMPRESO:" + Space(2) + str(rst("impresiones")) + Space(1) + "IMPRESIONES" + Chr(13) + "DESEA IMPRIMIR NUEVAMENTE ?", vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then
                    Procedencia = imprimir_s
                    Call disabled_form(FrmVentas)
                    FrmSeguridad.Show
              End If
          End If
      End If
salirr: X = 1
End Sub
Public Sub OrdenImpresion___()
On Error GoTo salirr
Dim X As Integer
Dim impresiones As Integer, id_venta As Double
       strCadena = "SELECT * FROM movimiento_venta WHERE numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND serie='" & Trim(Me.DtcSerieDoc.BoundText) & "' AND ruc='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
          If rst("impresiones") < 1 Then
              impresiones = rst("impresiones") + 1
              id_venta = rst("id_venta")
              'Call Imprimir_Tiketera(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcAlmacen.BoundText), Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text))
              Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text), rst("id_tipo_factura"), Trim(Me.txtDireccion.Text))
              
              strCadena = "UPDATE movimiento_venta SET impresiones='" & impresiones & "' WHERE id_venta='" & id_venta & "' AND ruc='" & KEY_RUC & "'"
              CnBd.Execute (strCadena)
              '
               
           Else
              If MsgBox("ESTE DOCUMENTO YA FUE IMPRESO:" + Space(2) + str(rst("impresiones")) + Space(1) + "IMPRESIONES" + Chr(13) + "DESEA IMPRIMIR NUEVAMENTE ?", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
                    Procedencia = imprimir_s
                    FrmSeguridad.Show
              End If
          End If
      End If
salirr: X = 1
End Sub

Public Sub OrdenImpresionEspecial()
Dim impresiones As Integer, id_venta As Double
       strCadena = "SELECT * FROM movimiento_venta WHERE numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND serie='" & Trim(Me.DtcSerieDoc.BoundText) & "' AND ruc='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
          If rst("impresiones") < 1 Then
              impresiones = rst("impresiones") + 1
              id_venta = rst("id_venta")
             ' Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text), "00002")
              strCadena = "UPDATE movimiento_venta SET impresiones='" & impresiones & "' WHERE id_venta='" & id_venta & "' AND ruc='" & KEY_RUC & "'"
              CnBd.Execute (strCadena)
               
               
              strCadena = "SELECT * FROM movimiento_venta WHERE numero='" & formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND serie='" & Trim(Me.DtcSerieDoc.BoundText) & "' AND ruc='" & KEY_RUC & "'"
              Call ConfiguraRst(strCadena)
              impresiones = rst("impresiones") + 1
              id_venta = rst("id_venta")
              Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText), formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6), "00002", 0)
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

Public Sub abrir_caja()
Open "COM1" For Output As #1 Len = 1
Write #1, Chr(13)
Close #1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

  

Me.DtcAlmacen.Enabled = True
Me.DtcTipoDoc.Enabled = True

Me.TxtNumeroDoc.Enabled = True
Call nuevo
Me.DtcTipoDoc.SetFocus
End Sub




Private Sub txtbuscarbanco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM entidadfinanciera  WHERE descripcion LIKE '%" & Trim(Me.txtbuscarbanco.Text) & "%'  ORDER BY descripcion"
     Call llenar_bancos(Me.HfBancos)
End If
End Sub

Private Sub txtBuscarSerie_Change()
strCadena = "SELECT Codigo,Descripcion FROM view_producto_serie WHERE id_alm='" & KEY_ALM & "' and  vendido='no' and transferencia='no' and  id_producto='" & Trim(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 1)) & "' and Descripcion LIKE '%" & Trim(Me.txtBuscarSerie.Text) & "%' AND  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcSerie)

End Sub

Private Sub txtBuscarSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcSerie.Enabled = True Then
        Me.DtcSerie.SetFocus
    End If
End If
End Sub

Private Sub txtBuscarVendedor_KeyPress(KeyAscii As Integer)
On Error GoTo salir

If KeyAscii = 13 Then
  strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad  WHERE nombre_completo LIKE '%" & Trim(Me.txtBuscarVendedor.Text) & "%' and  habilitado='si' and id_personal='si' and ruc='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcVendedor)
  
End If
Exit Sub
salir:

End Sub

Private Sub txtbusquedamotor_Change()
strCadena = "SELECT Codigo,motor as Descripcion FROM view_producto_serie WHERE vendido='no' and  id_producto='" & Trim(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 1)) & "' AND  motor LIKE '%" & Trim(Me.txtbusquedamotor.Text) & "%' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcMotor)
End Sub

Private Sub txtbusquedamotor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcMotor.Enabled = True Then
        Me.DtcMotor.SetFocus
    End If
End If
End Sub

Private Sub TxtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtCodProducto)
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.txtprecio)
End If
End Sub
Public Sub Resalta(ByVal Texto As TextBox)
On Error GoTo salir
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
salir:
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
Dim TotalP As Single
If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 And KeyAscii <> 13 Then
        KeyAscii = 0
End If

If KeyAscii = 13 Then
    If Me.OptAuto.Value = True Then
        Call Agregar_directo
    Else
        strCadena = "SELECT * FROM almacen_producto A,producto P WHERE A.id_producto=P.id_producto AND A.ruc='" & KEY_ALM & "' AND P.ruc='" & KEY_RUC & "' AND A.id_producto='" & Trim(codigoP) & "'"
        Call ConfiguraRst(strCadena)
        'Me.txtprecio.Locked = True
        
        
    If Val(Me.txtprecio.Text) = 0 Then
        MsgBox "Este Producto no Cuenta con un Precio de Venta", vbExclamation
        Call Resalta(Me.TxtCodProducto)
       Exit Sub
    End If
    TotalP = Val(Me.txtCantidad.Text) * Val(Me.txtprecio.Text)
    Me.LblTotalParcial.Caption = Format(TotalP, "#,##0.00")
    
    'Me.ChkPrecioAlterno.Enabled = True
    If Me.OptAuto.Value = True Then
        Call cmdagregar_Click
    End If
    End If
    
    If Trim(Me.txtagranel.Text) = "si" Then
            Me.DtcUnidad.SetFocus
    Else
            Call Resalta(Me.txtprecio)
    End If
    
    
    Set rst = Nothing
End If
End Sub
Public Sub Agregar_directo()
  '  strCadena = "SELECT     almacen_Producto.stock,producto.precio_venta " & _
   ' "FROM  almacen_productos INNER JOIN producto ON almacen_producto.id_proucto = Producto.cProducto WHERE (Almacen_Productos.cProducto='" & Trim(codigoP) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "')"
    'Call ConfiguraRst(strCadena)
    Call Resalta(Me.txtprecio)
    If Val(Me.txtprecio.Text) = 0 Then
        MsgBox "Este Producto no Cuenta con un Precio de Venta", vbExclamation
        Call Resalta(Me.TxtCodProducto)
       Exit Sub
    End If
    
    TotalP = Val(Me.txtCantidad.Text) * Val(Me.txtprecio.Text)
    Me.LblTotalParcial.Caption = Format(TotalP, "#,##0.00")
    'Me.ChkPrecioAlterno.Enabled = True
    Call cmdagregar_Click
    Set rst = Nothing
End Sub



Private Sub txtCheque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtMontoPagado)
End If
End Sub

Private Sub Txtclaverandon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtMontoPagovitekey.Visible = True
    Me.TxtMontoPagovitekey.Text = Format((Val(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000")))), "###0.00")
    
    Call Resalta(Me.TxtMontoPagovitekey)
    
    Exit Sub
End If
End Sub

Private Sub txtcliente_Change()
If (Trim(Me.TxtCodCliente.Text)) = "00000000" Then
      Me.TxtCliente.Locked = False
    Else
      Me.TxtCliente.Locked = True
 End If
End Sub

Private Sub TxtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtCodCliente)
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.TxtCodProducto)
End If
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
        If Me.chkDelivery.Value = 1 Then
            Me.TxtMontoPagado.Text = "0.00"
            Call Resalta(Me.TxtMontoPagado)
            Exit Sub
        End If
        Call Resalta(Me.txtDireccion)
                
End If
End Sub

Private Sub txtCodCliente_Change()
'If Len(Me.TxtCodCliente.Text) > 7 Then
 '   Me.FRMGastos.Visible = True
'Else
 '    Me.FRMGastos.Visible = False
'End If
End Sub

Private Sub TxtCodCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtNumeroDoc)
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.TxtCliente)
End If
End Sub
Public Sub put_mensualidad()
Dim in_mora_dias As Integer
'strCadena = "SELECT mora_dias FROM persona_plan_servicio WHERE pago_mensual='si' and dni='" & Trim(Me.TxtCodCliente.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
'Call ConfiguraRstK(strCadena)
'If rstK.RecordCount > 0 Then
 '   in_mora_dias = rstK("mora_dias")
    strCadena = "call put_temporal_cobranza_ii('" & Trim(Me.TxtCodCliente.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.DtcTipoDoc.BoundText) & "','" & Me.DtcSerieDoc.BoundText & "','" & KEY_ALM & "','" & in_mora_dias & "','" & KEY_MORA_MONTO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
'End If
Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcSerieDoc.BoundText, Me.DtcTipoDoc.BoundText, Trim(Me.txtformato_impresion.Text))


End Sub

Private Sub TxtCodCliente_KeyPress(KeyAscii As Integer)
On Error GoTo errohandler
 If KeyAscii = 13 Then
    Call precionar_cliente
    Exit Sub
 End If
    

If (KeyAscii = 66 Or KeyAscii = 98) Then
    Procedencia = Selecionar
    FrmPersona.Show
End If
Exit Sub
errohandler: MsgBox "Hubo un Error Digite Nuevamente", vbInformation, "Mensaje para el usuario"
End Sub
Public Sub llenarGrid_Facturas(ByVal Grilla As MSHFlexGrid, ByVal dni As String)
On Error GoTo salir

Dim Anulado As String
If dni = "00000000" Then
    GoTo sigt
End If
strCadena = "SELECT id_venta,fecha_emision,documento,total,anulado FROM movimiento_venta WHERE id_cliente='" & dni & "' AND ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC,numero DESC LIMIT 10"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount < 1 Then
sigt:
    Grilla.Rows = 0
    
    Exit Sub
End If
  
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstL.Fields.Count)
       For Each Campo In rstL.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 950
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 1000
           
       Next
        cabecera = "IDVENTA" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstL.MoveFirst
        in_total = 0
        For i = 0 To rstL.RecordCount - 1
            Fila = rstL("id_venta") & vbTab & Format(rstL("fecha_emision"), "dd-mm-YYYY") & vbTab & rstL("documento") & vbTab & Format(rstL("total"), "#,##0.00")
            Grilla.AddItem Fila
            If rstL("anulado") = "si" Then
                For k = 0 To 3
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
            Else
                in_total = in_total + rstL("total")
            End If
                rstL.MoveNext
        Next i
        Fila = "" & vbTab & "" & vbTab & "TOTAL PAGADO" & vbTab & Format(in_total, "#,##0.00")
            Grilla.AddItem Fila
        Exit Sub
salir:
  
End Sub
Public Sub llenarGrid_Facturas_all(ByVal Grilla As MSHFlexGrid, ByVal dni As String)
On Error GoTo salir

Dim Anulado As String
If dni = "00000000" Then
    GoTo sigt
End If
strCadena = "SELECT id_venta,fecha_emision,documento,total,anulado FROM movimiento_venta WHERE fecha_emision>='" & Format(Me.DtpIni.Value, "YYYY-mm-dd") & "' and fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' and  id_cliente='" & dni & "' AND ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC,numero DESC"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount < 1 Then
sigt:
    Grilla.Rows = 0
    
    Exit Sub
End If
  
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstL.Fields.Count)
       For Each Campo In rstL.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 950
           Grilla.ColWidth(2) = 1950
           Grilla.ColWidth(3) = 1100
           
       Next
        cabecera = "IDVENTA" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstL.MoveFirst
        For i = 0 To rstL.RecordCount - 1
            Fila = rstL("id_venta") & vbTab & Format(rstL("fecha_emision"), "dd-mm-YYYY") & vbTab & rstL("documento") & vbTab & Format(rstL("total"), "#,##0.00")
            Grilla.AddItem Fila
            If rstL("anulado") = "si" Then
                For k = 0 To 3
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
            End If
                rstL.MoveNext
        Next i
                            
        Exit Sub
salir:
    
  
End Sub
Public Sub llenarGrid_recibos(ByVal Grilla As MSHFlexGrid, ByVal dni As String)
Dim Anulado As String
If dni = "00000000" Then
    GoTo sigt
End If
strCadena = "SELECT id_venta,fecha_emision,documento,total,anulado FROM movimiento_venta WHERE id_doc='0054' and id_cliente='" & dni & "' and anulado='no' AND ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC,numero DESC LIMIT 0,10"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount < 1 Then
sigt:
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstL.Fields.Count)
       For Each Campo In rstL.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 950
           Grilla.ColWidth(2) = 1950
           Grilla.ColWidth(3) = 800
           
       Next
        cabecera = "IDVENTA" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstL.MoveFirst
        For i = 0 To rstL.RecordCount - 1
            Fila = rstL("id_venta") & vbTab & rstL("fecha_emision") & vbTab & rstL("documento") & vbTab & Format(rstL("total"), "#,##0.00")
            Grilla.AddItem Fila
            If rstL("anulado") = "si" Then
                For k = 0 To 3
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
            End If
                rstL.MoveNext
        Next i
                            
        
    
  
End Sub


Public Sub llenarGrid_Facturas_FECHA(ByVal Grilla As MSHFlexGrid, ByVal dni As String)
Dim Anulado As String
Dim Acumulado As Double
strCadena = "SELECT * FROM movimiento_venta WHERE id_cliente='" & dni & "' AND ruc='" & KEY_RUC & "' AND fecha_emision>='" & Format(Me.DtpIni.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' ORDER BY fecha_emision DESC,numero DESC LIMIT 0,100"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstL.Fields.Count)
       For Each Campo In rstL.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 950
           Grilla.ColWidth(2) = 1950
           Grilla.ColWidth(3) = 800
           
       Next
        cabecera = "IDVENTA" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rstL.MoveFirst
        Acumulado = 0
        For i = 0 To rstL.RecordCount - 1
            Fila = rstL("id_venta") & vbTab & rstL("fecha_emision") & vbTab & rstL("documento") & vbTab & Format(rstL("total"), "#,##0.00")
            Grilla.AddItem Fila
            If rstL("anulado") = "no" Then
                Acumulado = Acumulado + rstL("total")
            End If
            
            If rstL("anulado") = "si" Then
                For k = 0 To 3
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
                Next k
            End If
                rstL.MoveNext
        Next i
                            
        cabecera = "" & vbTab & "" & vbTab & "ACUMULADO TOTAL" & vbTab & Format(Acumulado, "#,##0.00")
        Grilla.AddItem cabecera
    
  
End Sub

Public Sub precionar_cliente()

If Trim(Me.TxtCodCliente.Text) = "" Then
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
 End If
  
 If KEY_PAIS = "9589" Then
 If Len(Trim(Me.TxtCodCliente.Text)) > 11 Then
    strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.TxtCodCliente.Text) & "' and extranjero='si'"
    Call ConfiguraRstK(strCadena)
    If rstK.RecordCount > 0 Then
        GoTo siguiente
    End If
End If
End If





If KEY_PAIS = "9589" Then
If Me.DtcTipoDoc.BoundText = "0001" Then
     
     If Len(Trim(Me.TxtCodCliente.Text)) <> 11 Then
        If get_extrangero(Trim(Me.TxtCodCliente.Text)) = "no" Then
        
            MsgBox "Debe Ingresar un Ruc para este comprobante", vbInformation
            Call Resalta(Me.TxtCodCliente)
            Exit Sub
        End If
    
     End If
     
 End If
End If
 


If KEY_PAIS = "9859" Then
If Trim(Me.DtcTipoDoc.BoundText) = "0001" And (Trim(Me.TxtCodCliente.Text) = "00000000" Or Trim(Me.TxtCodCliente.Text) = "" Or Len(Trim(Me.TxtCodCliente.Text)) <> 11) Then
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
End If
Else
    If Trim(Me.DtcTipoDoc.BoundText) = "0001" And (Trim(Me.TxtCodCliente.Text) = "00000000" Or Trim(Me.TxtCodCliente.Text) = "") Then
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
    End If
End If



siguiente:
If (Len(Trim(Me.TxtCodCliente.Text)) = 8 And Trim(Me.TxtCodCliente.Text) = "00000000") Then
    Me.TxtCliente.Text = "CLIENTE"
    Me.txtDireccion.Text = KEY_DIR_PUBLIC
    Call Resalta(Me.TxtCliente)
    Exit Sub
End If




 If Trim(Me.DtcTipoDoc.BoundText) = "0003" And (Trim(Me.TxtCodCliente.Text) = "") Then
    Me.TxtCodCliente.Text = "00000000"
    Me.TxtCliente.Text = "PUBLICO EN GENERAL"
    Me.txtDireccion.Text = KEY_DIR_PUBLIC
    Call Resalta(Me.TxtCliente)
    
    Exit Sub
End If
buscar_nuevamente:
Me.TxtCodCliente.Text = Trim(Me.TxtCodCliente.Text)

If Len(Trim(Me.TxtCodCliente.Text)) > 1 Then
    
    strCadena = "SELECT * FROM  persona WHERE  dni='" & Trim(Me.TxtCodCliente.Text) & "' LIMIT 1 "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
      If get_dni_reniec_iii(Trim(Me.TxtCodCliente.Text), KEY_DEPARTAMENTO, KEY_PROVINCIA, KEY_DISTRITO) = True Then
                GoTo buscar_nuevamente
            End If
        
        Procedencia = 1
        FrmDetallePersona.Show
        
        If KEY_PAIS = "9589" Then
        
        If Len(Trim(Me.TxtCodCliente.Text)) = 8 Then
            nruc = "10" & Trim(Me.TxtCodCliente.Text)
            FrmDetallePersona.txtRuc.Text = DigitoVerificadorRUC(Trim(nruc))
            FrmDetallePersona.ChkCliente.Value = 1
            Call FrmDetallePersona.precionar
            Exit Sub
        Else
            FrmDetallePersona.txtRuc.Text = Trim(Me.TxtCodCliente.Text)
            FrmDetallePersona.ChkCliente.Value = 1
            Call FrmDetallePersona.precionar
            Exit Sub
        End If
        Else
             FrmDetallePersona.ChkCliente.Value = 1
             Call Resalta(FrmDetallePersona.txtRuc)
             Exit Sub
        End If
    Else
        If KEY_PROYECTO = "no" Then
            If KEY_SEGURO_VENTA = "si" Then
                Call load_seguro(Trim(Me.TxtCodCliente.Text))
            End If
        End If
        
        
        
        Me.cmdcredito.Visible = True
        Me.imgFoto.Visible = True
        Me.TxtCliente.Text = UCase(rst("nombre_completo"))
        Me.txtDireccion.Text = UCase(rst("direccion"))
        Me.txtExtranjero.Text = rst("extranjero")
        
        If rst("afecto_percepcion") = "si" Then
           Me.lblPercepcion.Tag = "si"
        Else
           Me.lblPercepcion.Tag = "no"
        End If
        
        On Error GoTo nsalir
        If IsNull(rst("foto")) = False And Len(rst("foto")) > 5 Then
            If VerificarFichero(App.Path & "\archivos\" & rst("dni")) = True Then
                Me.imgFoto.Picture = LoadPicture(App.Path + "\archivos\" + rst("dni") + "\" + Trim(rst("foto")))
            Else
nsalir:
                Me.imgFoto.Picture = LoadPicture(App.Path + "\archivos\no_photo.jpg")
            End If
        Else
            If rst("sexo") = "M" Then
                Me.imgFoto.Picture = LoadPicture(App.Path + "\archivos\img_men.jpg")
            Else
                Me.imgFoto.Picture = LoadPicture(App.Path + "\archivos\img_dama.jpg")
            End If
        End If
        
    End If
End If

Call Resalta(Me.TxtCodProducto)

       
    
    Call load_credito_dispobible(Trim(Me.TxtCodCliente.Text))
    
    
    If KEY_GENERADOR_MENSUALIDAD = "si" Then
        strCadena = "P_nueva_venta_temporal('" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        Call put_mensualidad
     End If
     
     
     If KEY_PROYECTO = "si" And Trim(Me.TxtCodCliente.Text) <> "00000000" Then
      Me.chkseguro.Caption = "PROYECTO"
      Me.chkseguro.Visible = True
      Me.Dtcseguro.Visible = True
      Call load_proyecto(Trim(Me.TxtCodCliente.Text))
    End If
    

End Sub
Private Function load_credito_dispobible(ByVal in_dni As String) As Double
        strCadena = "SELECT func_credito_persona('" & in_dni & "','" & KEY_RUC & "')"
        Call ConfiguraRstT(strCadena)
        in_credito_persona = rstT(0)
        If Val(in_credito_persona) > 0 Then
             'Saldo
             strCadena = "SELECT funct_total_saldo('" & Trim(in_dni) & "','" & KEY_RUC & "')"
             Call ConfiguraRstT(strCadena)
             in_consumo_persona = rstT(0)
            
           
                Me.cmdVisualizar.Visible = True
                Me.cmdVisualizar.Caption = "TIENE" + Space(2) + Format(in_consumo_persona, "#,##0.00") + Space(2) + "DE CONSUMO"
            If in_credito_persona > 0 Then
                lblDisponible.Visible = True
                lblDisponible.Caption = "CREDITO DISPONIBLE :" + Space(2) + Format(Val(in_credito_persona - in_consumo_persona), "#,##0.00")
                TxtCreditoDisponible.Text = Val(in_credito_persona - in_consumo_persona)
                load_credito_dispobible = Val(in_credito_persona - in_consumo_persona)
            Else
                load_credito_dispobible = 0
            End If
        Else
               If in_credito_persona > 0 Then
                    Me.lblDisponible.Visible = True
                    Me.lblDisponible.Caption = "CREDITO DISPONIBLE :" + Space(2) + Format(Val(in_credito_persona), "#,##0.00")
                    Me.TxtCreditoDisponible.Text = Val(in_credito_persona)
                Else
                    Me.lblDisponible.Visible = False
                End If
               load_credito_dispobible = (in_credito_persona)
             Me.cmdVisualizar.Visible = False
             
End If


End Function



Private Sub TxtCodProducto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.txtCantidad)
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.txtCantidad)
End If
If KeyCode = 38 Then
    If Me.HfdDetalle.Rows > 0 Then
        Me.HfdDetalle.SetFocus
    End If
End If

End Sub
Private Sub TxtCodProducto_KeyPress(KeyAscii As Integer)
Dim Criterio As String
If KeyAscii = 13 Then
    
   If (Len(Me.TxtCodProducto.Text) = 0) Or Val(Me.TxtCodProducto.Text) = 0 Then
        chkPrecios.Enabled = False
        Call Resalta(Me.TxtCodProducto)
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
    End If
    
 
  ' If (Trim(Mid(Me.TxtCodProducto.Text, 2, 2)) = "00" Or Trim(Mid(Me.TxtCodProducto.Text, 1, 2)) = "20") And Len(Me.TxtCodProducto.Text) > 8 And Trim(Mid(Me.TxtCodProducto.Text, 1, 1)) <> "9" Then
  '     Me.txtcantidad.Text = Val(Mid(Trim(Me.TxtCodProducto.Text), 8, 5)) / 1000
  '     Me.TxtCodProducto.Text = formato_item(Mid(Me.TxtCodProducto.Text, 2, 6), 5)
  '     GoTo pesable
  '  End If
    
    If KEY_BARRAS = "si" Then
        strCadena = "SELECT B.id_producto,P.nombre_prod,P.precio_venta,P.peso,P.id_igv,A.stock,P.icbper FROM producto_barras B,producto P,almacen_producto A WHERE P.id_producto=A.id_producto AND A.ruc='" & KEY_RUC & "' AND B.id_producto=P.id_producto AND B.ruc='" & KEY_RUC & "' " & _
        "AND P.ruc='" & KEY_RUC & "' AND B.cod_barra='" & Trim(Me.TxtCodProducto.Text) & "'"
    Else
pesable:
        Me.TxtCodProducto.Text = Format(Me.TxtCodProducto.Text, "00000")
        If KEY_RUBRO = "00003" Then
            strCadena = "SELECT * FROM view_producto_selec WHERE (id_producto='" & Trim(Me.TxtCodProducto.Text) & "' or codigo_barra ='" & Trim(Me.TxtCodProducto.Text) & "') and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Else
            strCadena = "SELECT * FROM view_producto_selec WHERE id_producto='" & Trim(Me.TxtCodProducto.Text) & "'  and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
        End If
        
    End If
    
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       
       If get_producto_habilitado(rst("habilitado")) = False Then
           Me.TxtCodProducto.Text = ""
           Call Resalta(Me.TxtCodProducto)
           Exit Sub
       End If
        
        
        
        
        
        
        Me.txtagranel.Text = rst("agranel")
        Me.IN_ICBPER = rst("icbper")
        
        Call get_unidad(Trim(rst("id_producto")), rst("agranel"))
        If rst("agranel") = "si" Then
           in_precio = get_precio_unidad(rst("id_producto"), DtcUnidad.BoundText)
        Else
           
           If KEY_SEGMENTACION_PRECIO = "si" Then
               in_precio = get_precio_segmentacion(rst("id_producto"), Me.TxtCodCliente.Text)
           Else
               
                If KEY_GRIFO = "si" Then
                    in_precio = get_precio_propio(rst("id_producto"), FrmVentas.TxtCodCliente.Text, rst("precio_venta"))
                    
               Else
                    in_precio = rst("precio_venta")
               End If
               
           
           End If
           
           
        End If
        
        
        
        If DtcMoneda.BoundText = "00002" Then
           
           If KEY_CONVERSION_CAMBIO = "si" Then
              txtprecio.Text = Format(in_precio, "###0.00")
              txtpreciooriginal.Text = Format(in_precio, "###0.00")
            Else
              txtprecio.Text = Format(in_precio / (TxtTipoCambio.Text), "###0.00")
              txtpreciooriginal.Text = Format(in_precio / (TxtTipoCambio.Text), "###0.00")
            End If
           
           
        Else
           If KEY_CONVERSION_CAMBIO = "si" Then
              txtprecio.Text = in_precio * KEY_CAMBIO_LOCAL
              txtpreciooriginal.Text = in_precio * KEY_CAMBIO_LOCAL
            Else
              txtpreciooriginal.Text = in_precio
              txtprecio.Text = in_precio
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
        
        
        
        
         
        
        If rst("servicio") = "si" Then
           If rst("icbper") = "si" And txt_tipo.Text = "01" Then
              txt_tipo.Text = "01"
           Else
              txt_tipo.Text = "02"
           End If
           Me.txtServicio.Text = "si"
        Else
           Me.txt_tipo.Text = "01"
           Me.txtServicio.Text = "no"
        End If
        
        
        
        
        
        
        
        codigoP = rst("id_producto")
        Me.TxtDescripcionProducto.Text = rst("nombre_prod")
        Me.TxtIgv.Text = rst("id_igv")
        
        Me.txtpeso.Text = rst("peso")
        
        Me.txtprecio.Locked = False
        If Trim(Me.txtCantidad.Text) > 0 Then
            Me.txtCantidad.Text = Me.txtCantidad.Text
         Else
          Me.txtCantidad.Text = 1
        End If
        
        Call Resalta(Me.txtCantidad)
        
        If Me.OptAuto.Value = True Then
            Call txtCantidad_KeyPress(13)
        Else
            Call Me.mostrar_precios
        End If
        'Me.ChkPrecioAlterno.Enabled = True
        Set rst = Nothing
        
        
        Me.lblhistorial.Caption = get_ultimo_precio(Trim(Me.TxtCodCliente.Text), Trim(codigoP))
        
    Else
        chkPrecios.Enabled = False
        Call Resalta(Me.TxtCodProducto)
        Procedencia = Selecionar
        FrmProducto.Show
    End If
End If
End Sub


Public Function control_stock(ByVal in_producto As String, ByVal in_cantidad As Double) As Boolean

On Error GoTo salir
Dim in_reserva As Single

If Me.txtServicio.Text = "si" Then
    control_stock = True
Else

If KEY_RESERVA_STOCK = "si" Then
    strCadena = "select ALM_almacen_stock('1','" & in_producto & "','" & KEY_ALM & "','" & KEY_RUC & "')"
    Call ConfiguraRstAux(strCadena)
    in_reserva = rstAux(0)
Else
    in_reserva = 0
End If


strCadena = "SELECT stock FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then
    If (rstAux("stock") - in_reserva) < in_cantidad And KEY_MOVIMIENTO_SIN_STOCK = "no" Then
                
                MsgBox "NOTA:======= PRODUCTO NO CUENTA CON STOCK." + Chr(13) + Chr(13) + "[ " + in_producto + " ]" + Space(1) + get_producto(in_producto) + Space(2) + Chr(13) + "=======================================" + Chr(13) + "STOCK ACTUAL : " + str((rstAux("stock") - in_reserva)) + Chr(13) + "TOTAL PEDIDO :" + str(in_cantidad) + Chr(13) + Chr(13) + "Consulte con el Area de Almacen.", vbInformation, KEY_EMPRESA
                
                
                    Call Resalta(Me.TxtCodProducto)
                    control_stock = False
                
                
    Else
        control_stock = True
    End If
End If
End If

Exit Function
salir:


End Function


Public Function control_stock_cantidad(ByVal in_producto As String, ByVal in_cantidad As Double) As Single

On Error GoTo salir


If Me.txtServicio.Text = "si" Then
    control_stock_cantidad = in_cantidad
Else
    strCadena = "SELECT stock FROM almacen_producto WHERE id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstAux(strCadena)
If rstAux.RecordCount > 0 Then

    If rstAux("stock") < in_cantidad And chk_venta_diferida.Value = 0 And KEY_MOVIMIENTO_SIN_STOCK = "no" Then
                
                If MsgBox("NOTA:======= PRODUCTO NO CUENTA CON STOCK." + Chr(13) + Chr(13) + "[ " + in_producto + " ]" + Space(1) + get_producto(in_producto) + Space(2) + Chr(13) + "=======================================" + Chr(13) + "STOCK ACTUAL : " + str(rstAux("stock")) + Chr(13) + "TOTAL PEDIDO :" + str(in_cantidad) + Chr(13) + Chr(13) + "Desea Ingresar SOLO LA CANTIDAD DISPONIBLE.", vbInformation + vbYesNo, KEY_EMPRESA) = vbYes Then
                   control_stock_cantidad = rstAux("stock")
                Else
                    If Me.DtcTipoDoc.BoundText = "0099" Then
                        control_stock_cantidad = in_cantidad
                    Else
                        control_stock_cantidad = 0
                    End If
               End If
    Else
        control_stock_cantidad = in_cantidad
    End If

End If
End If
Exit Function
salir:


End Function



Private Sub TxtCuotas_KeyPress(KeyAscii As Integer)
Dim vencimiento As String
If KeyAscii = 13 Then
        vencimiento = Format(Date, "YYYY-mm-dd")
        strCadena = "DELETE FROM movimiento_venta_cuotas_temporal WHERE id_usuario='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
         
         Me.TxtMontoPagado.Visible = True
        If Val(Me.TxtCuotas.Text) < 1 Then
            Me.TxtCuotas.Text = 1
        End If
        
        For k = 1 To Val(Me.TxtCuotas.Text)
            vencimiento = Format(DateAdd("m", 1, vencimiento), "YYYY-mm-dd")
            Me.TxtMontoPagado.Text = Format(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000")), "###0.000")
            If Val(Me.TxtCreditoDisponible.Text) < Val(Me.TxtMontoPagado.Text) Then
                 MsgBox "EL MONTO EXCEDE AL CREDITO ACTUAL EN " + Space(1) + Format(Val(Me.TxtMontoPagado.Text) - Val(Me.TxtCreditoDisponible.Text), "#,##0.00"), vbInformation, KEY_EMPRESA
                 Call Resalta(Me.TxtMontoPagado)
                 Exit Sub
            'Else
                
           ' strCadena = "INSERT INTO movimiento_venta_cuotas_temporal(id_cuota,id_doc,serie,numero,monto,saldo,vencimiento,id_usuario,ruc)VALUES " & _
            "('" & formato_item(k, 2) & "','" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcSerieDoc.BoundText & "','" & Me.TxtNumeroDoc.text & "','" & Val(Me.TxtMontoPagado.text) / Val(Me.TxtCuotas.text) & "','" & Val(Me.TxtMontoPagado.text) / Val(Me.TxtCuotas.text) & "','" & vencimiento & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
            'CnBd.Execute (strCadena)  '
        
            End If
            
            
        Next k
        
   
    Call Resalta(Me.TxtMontoPagado)
   
    Exit Sub
End If
End Sub

Private Sub put_descuento_porcentaje()
If in_total_documento > 0 Then
  Me.TxtDescuento_global.Text = Format(in_total_documento * Val(Me.TxtDescuento_porcentaje.Text) / 100, "###0.00")
   Me.txtObservacion.Text = "DESCUENTO DE :" & Trim(Format(Val(Me.TxtDescuento_porcentaje.Text), "###0.00")) & "%"
   Me.lblTotal.Caption = Format(in_total_documento - Val(Me.TxtDescuento_global.Text), "###0.00")
   
   
   If KEY_CON_IGV = "si" Then
      Me.LblValorVenta.Caption = Format(Val(Me.lblTotal.Caption) / (1 + KEY_IGV), "###0.00")
      Me.LblIgv.Caption = Format(Val(Me.LblValorVenta.Caption) * KEY_IGV, "###0.00")
   Else
      Me.LblValorVenta.Caption = Format(Val(Me.lblTotal.Caption), "###0.00")
      Me.LblIgv.Caption = Format(0, "###0.00")
      Me.lblExonerado.Caption = Format(Val(Me.lblTotal.Caption), "###0.00")
   End If
   
   Me.lblVuelto.Caption = Format(in_total_documento - Val(Me.TxtDescuento_global.Text) - Val(Me.lblPago.Caption), "###0.00")
End If
End Sub
Private Sub put_descuento_total()
If in_total_documento > 0 Then
   Me.TxtDescuento_porcentaje.Text = Val(Me.TxtDescuento_global.Text) * 100 / in_total_documento
   Me.txtObservacion.Text = "DESCUENTO DE :" & Trim(Format(Val(Me.TxtDescuento_porcentaje.Text), "###0.00")) & "%"
   Me.lblTotal.Caption = Format(in_total_documento - Val(Me.TxtDescuento_global.Text), "###0.00")
   
   If KEY_CON_IGV = "si" Then
      Me.LblValorVenta.Caption = Format(Val(Me.lblTotal.Caption) / (1 + KEY_IGV), "###0.00")
      Me.LblIgv.Caption = Format(Val(Me.LblValorVenta.Caption) * KEY_IGV, "###0.00")
      
   Else
      Me.LblValorVenta.Caption = Format(Val(Me.lblTotal.Caption), "###0.00")
      Me.LblIgv.Caption = Format(0, "###0.00")
      Me.lblExonerado.Caption = Format(Val(Me.lblTotal.Caption), "###0.00")
   End If
   
   Me.lblVuelto.Caption = Format(in_total_documento - Val(Me.TxtDescuento_global.Text) - Val(Me.lblPago.Caption), "###0.00")
   End If
End Sub

Private Sub TxtDescripcionProducto_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Call Resalta(Me.txtprecio)
End If
End Sub

Private Sub TxtDescuento_global_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call put_descuento_total
End If
End Sub

Private Sub TxtDescuento_porcentaje_Change()

Call put_descuento_porcentaje

End Sub

Private Sub TxtDireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtCodProducto)
End If
End Sub


Private Sub txtdni_copropietario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT nombre_completo FROM persona p where dni='" & Trim(Me.txtdni_copropietario.Text) & "'"
    Call ConfiguraRstF(strCadena)
    If rstF.RecordCount > 0 Then
        Me.lblcopropietario.Caption = rstF("nombre_completo")
    Else
        Procedencia = seleccionar_otro
        FrmPersona.Show
        Exit Sub
    End If
End If
End Sub

Private Sub txtFecha_vencimiento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtMontoPagado)
End If
End Sub

Private Sub txtguia_referencia_Change()
On Error GoTo salir

strCadena = "SELECT id_transferencia as Codigo, CONCAT('GUIA:',serie,'-',numero) as Descripcion FROM movimiento_transferencia WHERE numero LIKE '%" & Trim(Me.txtguia_referencia.Text) & "%' and  anulado='no' and  ruc='" & KEY_RUC & "' ORDER BY id_transferencia DESC LIMIT 10"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcGuia)

Exit Sub
salir:


End Sub

Private Sub TxtMontoPagado_KeyPress(KeyAscii As Integer)

If (KeyAscii = 13) Then
    
    Call realizar_ingreso_pago
    
End If
End Sub

Private Function get_monto_caja(ByVal m_total As Single, m_pagado As Single, ByVal in_tipomovimiento As String) As Single
Dim in_factor As Integer
Dim in_vuelto As Double
in_vuelto = Format(lblVuelto.Caption, "###0.00")

If in_tipomovimiento = "01" Then
    in_factor = 1
Else
    in_factor = -1
End If


If m_pagado <= m_total Then
        
        If Val(in_vuelto) < 0 Then
            If m_pagado >= Val(in_vuelto) * -1 Then
               get_monto_caja = Format(Format(in_vuelto, "###0.00") * -1, "###0.00") * in_factor
            Else
               get_monto_caja = Format(m_pagado, "###0.00") * in_factor
            End If
        Else
            get_monto_caja = Format(m_pagado, "###0.00") * in_factor
        End If
        
Else
        get_monto_caja = Format(m_total, "###0.00") * in_factor
End If
End Function
Private Function get_id_nota(ByVal in_serie As String, ByVal in_numero As String)
    
    strCadena = "SELECT id_venta FROM movimiento_venta WHERE id_doc='0007' and serie='" & in_serie & "' and numero='" & in_numero & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRstA(strCadena)
    If rstA.RecordCount > 0 Then
       get_id_nota = rstA("id_venta")
    Else
       get_id_nota = 0
    End If
    
End Function
Private Sub realizar_ingreso_pago()
Dim monto_pagado As Single
Dim monto_caja As Single
Dim cod_pago As String
Dim nuevo_monto As Double
Dim strTarjeta As String
Dim Saldo_deudor As Single
Dim vencimiento As String
Dim nrecibo As String

If Me.frmbanco.Visible = True And (Trim(Me.txtCheque.Text) = "" Or Trim(Me.txtBanco.Text) = "") Then
   MsgBox "Seleccione Un Banco y un N� Cheque", vbInformation, "PAGO CON CHEQUE"
   Call Resalta(Me.txtCheque)
   Exit Sub
End If
    
    vencimiento = Format(KEY_FECHA, "YYYY-mm-dd")
    If Me.DtcTipoDoc.BoundText = "0007" Then
        GoTo sigy
    End If
    
    If Val(Me.TxtDescuento_porcentaje.Text) = 100 Then
        GoTo sigy
    End If
    
    If Val(Me.lblGratuitas.Caption) > 0 Then
        GoTo sigy
    End If
    
    
    If (Val(Me.TxtMontoPagado.Text) > 0 And Val(Me.lblTotal.Caption) > 0) Then
sigy:
     monto_pagado = Val(Me.TxtMontoPagado.Text)
     monto_caja = get_monto_caja(Format(Me.lblTotal.Caption, "###0.00"), monto_pagado, Trim(Me.txt_tipo_movimiento.Text))
        
        
     
     If Me.DtcMoneda.BoundText = "00002" Then
        If get_moneda_documento(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText)) = "00002" Then
            monto_pagado = monto_pagado
        Else
            monto_pagado = monto_pagado * Val(Me.TxtTipoCambio.Text)
        End If
        
        
     End If
     Me.lblPago.Caption = Format(monto_pagado, "###0.000")
     Me.lblVuelto.Caption = Format(Me.lblPago.Caption - Me.lblTotal.Caption, "###0.000")
        
        If Me.DtTargeta.Visible = False Then
            strTarjeta = "00"
        Else
            strTarjeta = Me.DtTargeta.BoundText
        End If
    
    strCadena = "SELECT * FROM movimiento_venta_monto_temporal WHERE id_tarjeta_operacion='" & Trim(Me.TxtOperacion.Text) & "' and  id_usuario='" & KEY_USUARIO & "' AND id_forma_pago='" & Trim(Me.DtcFormapagodetalle.BoundText) & "' and id_moneda='" & Trim(Me.DtcMoneda.BoundText) & "'  AND fecha='" & KEY_FECHA & "' and id_alm='" & KEY_ALM & "' and  ruc='" & KEY_RUC & "' AND id_tarjeta='" & strTarjeta & "'  ORDER BY id_monto DESC"
    Call ConfiguraRst(strCadena)
        
    If rst.RecordCount < 1 Then
        
       If Me.DtcFormaPago.BoundText = "02" Then
           TxtCreditoDisponible.Text = get_credito_disponible_persona(Me.TxtCodCliente.Text)
           
           If IsDate(Me.txtFecha_vencimiento.Text) = False Then
               MsgBox "Debe Ingresar una Fecha de Vencimiento Correcto.", vbInformation
               Exit Sub
           End If
        
        
        
        'CREDITO DISPONIBLE
        
       
             If (Val(Me.TxtCreditoDisponible.Text) < Val(Me.TxtMontoPagado.Text)) And Me.DtcTipoDoc.BoundText <> "0007" Then
                If KEY_LINEA_CREDITO = "si" Then
                 MsgBox "EL MONTO EXCEDE AL CREDITO ACTUAL EN " + Space(1) + str(Format(Val(Me.TxtMontoPagado.Text) - Val(Me.TxtCreditoDisponible.Text), "#,##0.00")), vbInformation, KEY_EMPRESA
                 Call Resalta(Me.TxtMontoPagado)
                 Exit Sub
                 End If
             Else
                If (Me.TxtCuotas.Visible = True And Val(Me.TxtCuotas.Text) > 0) Then
                    For k = 1 To Val(Me.TxtCuotas.Text)
                        vencimiento = Format(DateAdd("d", KEY_DIAS_CREDITO, vencimiento), "YYYY-mm-dd")
                        strCadena = "INSERT INTO movimiento_venta_cuotas_temporal(id_cuota,id_doc,serie,numero,monto,saldo,vencimiento,id_alm,id_usuario,ruc)VALUES " & _
                        "('" & formato_item(k, 2) & "','" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcSerieDoc.BoundText & "','" & Me.TxtNumeroDoc.Text & "','" & Val(Me.TxtMontoPagado.Text) / Val(Me.TxtCuotas.Text) & "','" & Val(Me.TxtMontoPagado.Text) / Val(Me.TxtCuotas.Text) & "','" & vencimiento & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                        CnBd.Execute (strCadena)
                    Next k
                End If
             End If
          End If
      
       
       If Val(Me.txtrecibo_anterior.Text) > 0 Then
            nrecibo = Me.HfRecibos.TextMatrix(Me.HfRecibos.Row, 2)
       Else
            nrecibo = " "
       End If
       
       If Me.DtcFormapagodetalle.BoundText = "12" Then
            Me.txtCheque.Text = Trim(Me.txtBanco.Text) & "-" & Trim(Me.txtCheque.Text)
       Else
           If strTarjeta <> "00" Then
                Me.txtCheque.Text = "PAGO TARJETA CON REF:" & Trim(Me.TxtOperacion.Text)
           Else
                Me.txtCheque.Text = ""
           End If
       End If
       
       If Me.frmalmacen_entrega(1).Visible = True And Val(Me.Label22(3).Caption) > 0 Then
       
            in_serie_nota = Trim(Me.txtserie_nota.Text)
            in_numero_nota = Trim(Me.txtNumero_nota.Text)
            in_id_nota = get_id_nota(in_serie_nota, in_numero_nota)
        Else
            in_serie_nota = 0
            in_numero_nota = 0
            in_id_nota = 0
       End If
            
       strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,cuenta_contable,cuotas,id_usuario,id_recibo,detalle,banco,cheque,fecha,id_alm,id_nota_credito,serie_nota,numero_nota,ruc) VALUES " & _
       " ('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcSerieDoc.BoundText & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.DtcFormaPago.BoundText) & "','" & Trim(Me.DtcFormapagodetalle.BoundText) & "','" & Me.DtcMoneda.BoundText & "','" & monto_pagado & "','" & monto_caja & "','" & strTarjeta & "','" & Me.TxtNumeroTargeta.Text & "','" & Me.TxtOperacion.Text & "','" & get_cuenta_contable_caja(Me.DtcFormapagodetalle.BoundText) & "','" & Val(Me.TxtCuotas.Text) & "','" & KEY_USUARIO & "','" & Val(Me.txtrecibo_anterior.Text) & "','" & nrecibo & "','-','" & Trim(Me.txtCheque.Text) & "','" & KEY_FECHA & "','" & KEY_ALM & "','" & in_id_nota & "','" & in_serie_nota & "','" & in_numero_nota & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
    Else
    
    If Me.DtcMoneda.BoundText = "00001" Then
        nuevo_monto = Val(Me.TxtMontoPagado.Text)
        monto_caja = get_monto_caja(Format(Me.lblTotal.Caption, "###0.00"), Val(Me.TxtMontoPagado.Text), Trim(Me.txt_tipo_movimiento.Text))
    Else
        nuevo_monto = Val(Me.TxtMontoPagado.Text) * Val(Me.TxtTipoCambio.Text)
        monto_caja = get_monto_caja(Format(Me.lblTotal.Caption, "###0.00"), Val(Me.TxtMontoPagado.Text) * Val(Me.TxtTipoCambio.Text), Trim(Me.txt_tipo_movimiento.Text))
    End If
       
        
    
    strCadena = "UPDATE movimiento_venta_monto_temporal SET monto='" & nuevo_monto & "',monto_caja='" & monto_caja & "',id_tarjeta='" & strTarjeta & "',id_tarjeta_numero='" & Me.TxtNumeroTargeta.Text & "',id_tarjeta_operacion='" & Me.TxtOperacion.Text & "',id_recibo='" & Val(Me.txtrecibo_anterior.Text) & "',detalle='" & nrecibo & "' WHERE id_usuario='" & KEY_USUARIO & "' AND id_forma_pago='" & Trim(Me.DtcFormapagodetalle.BoundText) & "'  AND id_alm='" & KEY_ALM & "' and  ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
        
    End If
    Me.TxtNumeroTargeta.Text = ""
    Me.TxtOperacion.Text = ""
    Me.txtBanco.Text = ""
    Me.txtCheque.Text = ""
    Me.frmbanco.Visible = False
    
    
    Call llena_pagos(Me.HfgTipoPagos, Me.TxtNumeroDoc.Text)
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
Dim tventa As Double

strCadena = "SELECT * FROM view_pago_temporal_vitekey WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' and  id_usuario='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Me.lblPago.Caption = "0.00"
    Me.lblVuelto.Caption = "0.00"
    Exit Sub
    
End If
   
       Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3400
           Grilla.ColWidth(2) = 1300
           
       Next
        cabecera = "CODIGO" & vbTab & "FORMA PAGO" & vbTab & "MONTO"
        Grilla.AddItem cabecera
         For k = 1 To 2
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
                    strmoneda = " [ US$/.]"
            End Select
            
            
                If rst("id_tarjeta") = "00" Then
                    strTarjeta = rst("descripcion") + Space(2) + rst("observacion")
                Else
                    strTarjeta = rst("descripcion") + Space(2) + rst("observacion") + Space(2) + "[" + rst("tarjeta") + "]"
                End If
               
            
            Fila = rst("id_monto") & vbTab & strTarjeta & vbTab & Format(rst("monto"), "###0.00")
            Grilla.AddItem Fila
            tpago = rst("monto") + tpago
            rst.MoveNext
    Next i
    
    tventa = Val(Format(Me.lblTotal.Caption, "###0.00"))
    Me.lblPago.Caption = Format(tpago, "###0.00")
    Me.lblVuelto.Caption = Format(tpago - tventa, "#,##0.00")
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub
Public Sub llena_pagosVenta(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo salir
Dim tpago As Double

strCadena = "SELECT * FROM view_venta_pago_ultimate WHERE id_venta='" & idVenta & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    If Me.chkconsultar.Value = 0 Then
        Me.lblPago.Caption = "0.00"
    End If
    Me.lblVuelto.Caption = "0.00"
    Exit Sub
    
End If
     Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3400
           Grilla.ColWidth(2) = 1300
       Next
        cabecera = "CODIGO" & vbTab & "FORMA PAGO" & vbTab & "MONTO"
        Grilla.AddItem cabecera
         For k = 1 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tpago = 0
        For i = 0 To rst.RecordCount - 1
            If rst("id_tarjeta") = "00" Then
                    strTarjeta = rst("descripcion") + Space(2) + rst("observacion")
            Else
                    strTarjeta = rst("descripcion") + Space(2) + "[" + rst("tarjeta") + "]"
            End If
            
            
            
            
            Fila = rst("id_detalle") & vbTab & strTarjeta & vbTab & Format(rst("monto"), "###0.00")
            
            Grilla.AddItem Fila
            tpago = rst("monto") + tpago
            rst.MoveNext
    Next i
    Dim tventa As Double
    tventa = Val(Format(Me.lblTotal.Caption, "###0.000"))
    Me.lblTotal.Caption = Format(tventa, "###0.000")
    Me.lblPago.Caption = Format(tpago, "###0.000")
    Me.lblVuelto.Caption = Format(tpago - tventa, "#,##0.000")
Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub


Sub llena_pagos_g(ByVal Grilla As MSHFlexGrid)
    Grilla.Clear
  Set Grilla.Recordset = rst

  Grilla.ColWidth(0) = 2000
  Grilla.ColWidth(1) = 1000

Call DarFormato(Grilla, 1)
Set rst = Nothing

End Sub


Private Sub TxtNumero_guia_KeyPress(KeyAscii As Integer)
Dim idVenta As Double
If KeyAscii = 13 Then
    
    
    
    Me.TxtNumero_guia.Text = FormatosCeros(Me.TxtNumero_guia.Text, 6)
    
    strCadena = "SELECT * FROM movimiento_venta WHERE (numero='" & Trim(Me.TxtNumero_guia.Text) & "' AND id_doc='" & Trim(Me.DtcComprobanteGuia.BoundText) & "' AND serie='" & Trim(Me.TxtSeri_guia.Text) & "' AND ruc='" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Me.HfdDetalle.Clear
         MsgBox "DOCUMENTO NO REGISTRADO ", vbInformation, KEY_EMPRESA
         Call Resalta(Me.TxtNumero_guia)
         Exit Sub
      
    Else
            idVenta = rst("id_venta")
            
            If MsgBox("ESTA SEGURO DE REALIZAR ESTA OPERACI�N", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
mostrar:
                Me.txtid_venta_ref.Text = idVenta
               
                'Validar proforma
                If Me.DtcComprobanteGuia.BoundText = "0099" Then
                    If validar_orden_proforma(idVenta) = False Then
                    Exit Sub
                    End If
                End If
                
                
                Me.txttipofactura.Text = rst("id_tipo_factura")
                Me.txt_tipo.Text = rst("id_tipo")
                Me.DtcVendedor.BoundText = rst("id_vendedor")
                If rst("diferida") = "si" Then
                    chk_venta_diferida.Value = 1
                Else
                    chk_venta_diferida.Value = 0
                End If
                Call LlenarDatosCliente_referencia(idVenta)
                Call llenar_comprobante_referencia(idVenta)
                Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcSerieDoc.BoundText, Me.DtcTipoDoc.BoundText, Trim(Me.txtformato_impresion.Text))
                Me.TxtCodProducto.Enabled = False
                Me.TxtDescripcionProducto.Enabled = False
                Me.txtprecio.Enabled = False
                Me.cmdAgregar.Enabled = False
                Me.CmdQuitar.Enabled = False
                Call Resalta(TxtNumero_guia)
                Referencia = True
                If get_verificar_tiene_nota(idVenta) = True Then
                '   Me.cmdProcesar.Enabled = False
                Else
                 '  Me.cmdProcesar.Enabled = True
                End If
                Me.cmdProcesar.Enabled = True
                Me.cmdImprimir.Enabled = False
                Me.TxtCodProducto.Enabled = True
                Me.txtprecio.Enabled = True
                Me.cmdAgregar.Enabled = True
                Me.CmdQuitar.Enabled = True
            End If
    End If
End If
Set rst = Nothing
End Sub

Public Function get_verificar_tiene_nota(ByVal in_venta As String) As Boolean
Dim in_comprobantes As String

strCadena = "SELECT id_venta FROM movimiento_venta WHERE id_doc='0007' and id_comprobante='" & Val(in_venta) & "' and anulado='no' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
            rstK.MoveFirst
            in_comprobantes = ""
            For i = 0 To rstK.RecordCount - 1
            strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & rstK("id_venta") & "'"
            Call ConfiguraRstL(strCadena)
            If rstL.RecordCount > 0 Then
                
               
                If rstL("id_doc") = "0007" Then
                    If in_comprobantes = "" Then
                        in_comprobantes = rstL("documento")
                    Else
                        in_comprobantes = rstL("documento") & Chr(13) & in_comprobantes
                    End If
                    
                    
                    Me.txtmotivo_nota.Text = rstL("motivo_nota")
                    Me.cmdProcesar.Enabled = False
                    get_verificar_tiene_nota = True
                    Me.FrameReferencia.Visible = True
                Else
                    Me.cmdProcesar.Enabled = True
                    get_verificar_tiene_nota = False
                    Me.FrameReferencia.Visible = False
                End If
                
            End If
           
           rstK.MoveNext
           Next i
           Me.txtdocreferencia.Text = in_comprobantes


Else
            Me.txtdocreferencia.Text = ""
            Me.FrameReferencia.Visible = False
            get_verificar_tiene_nota = False
   
End If


End Function
Private Function get_pagos_pedido(ByVal in_venta As String) As Boolean
get_pagos_pedido = False

strCadena = "DELETE FROM movimiento_venta_monto_temporal WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' and serie='" & Me.DtcSerieDoc.BoundText & "' and id_usuario='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM movimiento_venta_monto WHERE id_venta='" & Val(in_venta) & "' ORDER BY id_detalle ASC"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   rstK.MoveFirst
   get_pagos_pedido = True
   For i = 0 To rstK.RecordCount - 1
        strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,forma_pago,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,cuenta_contable,cuotas,id_usuario,id_recibo,banco,cheque,fecha,id_alm,ruc) VALUES " & _
       " ('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcSerieDoc.BoundText & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & rstK("forma_pago") & "','" & rstK("id_forma_pago") & "','" & Me.DtcMoneda.BoundText & "','" & rstK("monto") & "','" & rstK("monto_caja") & "','" & rstK("id_tarjeta") & "','" & rstK("id_tarjeta_numero") & "','" & rstK("id_tarjeta_operacion") & "','" & rstK("cuenta_contable") & "','" & Val(Me.TxtCuotas.Text) & "','" & KEY_USUARIO & "','" & rstK("id_recibo") & "','" & rstK("banco") & "','" & rstK("cheque") & "','" & KEY_FECHA & "','" & KEY_ALM & "','" & KEY_RUC & "')"
       
       
       
       CnBd.Execute (strCadena)
       rstK.MoveNext
   Next i
   
End If

strCadena = "DELETE FROM movimiento_venta_cuotas_temporal WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' and serie='" & Me.DtcSerieDoc.BoundText & "'  and id_usuario='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
Me.cmdCuotas.Visible = False
Me.TxtCuotas.Text = 0
strCadena = "SELECT * FROM movimiento_venta_cuotas WHERE id_venta='" & Val(in_venta) & "' ORDER BY id ASC"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   Me.TxtCuotas.Text = rstK.RecordCount
   rstK.MoveFirst
   Me.cmdCuotas.Visible = True
   For i = 0 To rstK.RecordCount - 1
        strCadena = "INSERT INTO movimiento_venta_cuotas_temporal(id_cuota,id_doc,serie,numero,monto,saldo,vencimiento,id_usuario,ruc) VALUES " & _
       " ('" & rstK("id_cuota") & "','" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.DtcSerieDoc.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & rstK("monto") & "','" & rstK("saldo") & "','" & Format(rstK("vencimiento"), "YYYY-mm-dd") & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       rstK.MoveNext
   Next i
   
End If

       
End Function
Public Function validar_orden_proforma(ByVal in_venta As String) As Boolean
strCadena = "SELECT id_producto,cantidad FROM movimiento_venta_detalle WHERE icbper='no' and  id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    rstL.MoveFirst
    For i = 0 To rstL.RecordCount - 1
        If control_stock(rstL("id_producto"), rstL("cantidad")) = True Then
           validar_orden_proforma = True
        Else
            validar_orden_proforma = False
            Exit Function
        End If
        rstL.MoveNext
    Next i
End If
 
End Function


Public Sub get_comprobante(ByVal i_venta As Double)
           Dim int_venta As Double
           Dim in_descuento As Double
           
           strCadena = "SELECT id_venta,id_tipo_factura,id_vendedor,observacion,id_copropietario,id_cliente,diferida,fecha_emision,descuento,icbper FROM movimiento_venta where id_venta='" & i_venta & "' AND ruc='" & KEY_RUC & "'"
           Call ConfiguraRst(strCadena)
           If rst.RecordCount > 0 Then
                int_venta = rst("id_venta")
                in_descuento = rst("descuento")
                If DateDiff("d", Format(rst("fecha_emision"), "YYYY-mm-dd"), KEY_FECHA) > 7 Then
                    MsgBox "PROFORMA EXCEDIO EL LIMITE DE DIAS [7]." + Chr(13) + Chr(13) + "FECHA EMISION:" + str(rst("fecha_emision")) + Chr(13), vbInformation
                    Exit Sub
                End If
                
                Me.txttipofactura.Text = rst("id_tipo_factura")
                Me.DtcVendedor.BoundText = rst("id_vendedor")
                Me.txtObservacion.Text = rst("observacion")
                Me.txtdni_copropietario.Text = rst("id_copropietario")
                Me.TxtCodCliente.Text = rst("id_cliente")
                Me.txtTelefono.Text = get_telefono(Me.TxtCodCliente.Text)
                Me.txtmail.Text = get_mail(Me.TxtCodCliente.Text)
                
                lblcopropietario.Caption = get_persona(rst("id_copropietario"))
                
                
                If rst("diferida") = "si" Then
                    chk_venta_diferida.Value = 1
                Else
                   chk_venta_diferida.Value = 0
                End If
                Call LlenarDatosPedido(int_venta)
                'If Me.chk_venta_diferida.Value = 0 Then
                    
                '    If validar_orden_proforma(i_venta) = False Then
                '        Exit Sub
                '    End If
                'End If
                
                '---
                
                Call Llenar_Temporal(int_venta)
                '--
                
                Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcSerieDoc.BoundText, Me.DtcTipoDoc.BoundText, Trim(Me.txtformato_impresion.Text))
                If in_descuento > 0 Then
                    Me.TxtDescuento_global.Text = in_descuento
                     Call put_descuento_total
                End If
                
                Me.TxtCodProducto.Enabled = False
                Me.TxtDescripcionProducto.Enabled = False
                Me.txtprecio.Enabled = False
                Me.cmdAgregar.Enabled = False
                Me.CmdQuitar.Enabled = False
                Me.cmdProcesar.Enabled = True
                Me.cmdImprimir.Enabled = False
                Me.TxtCodProducto.Enabled = True
                Me.txtprecio.Enabled = True
                Me.cmdAgregar.Enabled = True
                Me.CmdQuitar.Enabled = True
          End If
          
         ' Call get_pagos_pedido(i_venta)
            
End Sub
Private Function get_cantidad_nota(ByVal in_producto As String, ByVal in_venta As String) As Single


strCadena = "SELECT * FROM movimiento_venta_detalle "


strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_producto='" & in_producto & "' and  id_venta='" & in_venta & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   get_cantidad_nota = rstP("cantidad")
End If


End Function

Private Sub Llenar_Temporal(ByVal idVenta As Double)
Dim total_temp As Double
Dim i As Integer
Dim in_precio_venta As Double
Dim codigoP As String
Dim cantidadP As Single
Dim n_cantidad As Single

strCadena = "SELECT * FROM view_detalle_comprobante WHERE id_venta='" & idVenta & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then

strCadena = "DELETE FROM temporal_ventas WHERE ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "' AND id_alm='" & KEY_ALM & "' "
CnBd.Execute (strCadena)

total_temp = 0
rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        codigoP = rst("id_producto")
        cantidadP = rst("cantidad")
        
        cantidadP = control_stock_cantidad(Trim(codigoP), rst("cantidad"))
        
        If cantidadP = 0 Then
            GoTo siguiente
        End If
        
        If KEY_SEGMENTACION_PRECIO = "si" Then
               in_precio_venta = get_precio_segmentacion(codigoP, TxtCodCliente.Text)
        Else
               in_precio_venta = rst("precio")
        End If
        
        
        
        
        strCadena = "INSERT INTO temporal_ventas(ruc,id_alm,id_dni,id_doc,id_serie,numero,id_producto,detalle,cantidad,precio,total,peso,dni_save,id_detalle_serie,serie,anio_fabricacion,nro_chasis,anio_modelo,nro_dua,nro_item,servicio,costo,obsequio,icbper) VALUES " & _
        "('" & KEY_RUC & "','" & Me.DtcAlmacen.BoundText & "','" & Trim(Me.TxtCodCliente.Text) & "','" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcSerieDoc.BoundText & "','" & Me.TxtNumeroDoc.Text & "','" & rst("id_producto") & "','" & rst("detalle") & "','" & cantidadP & "'," & _
        "'" & in_precio_venta & " ','" & in_precio_venta * cantidadP & "','" & rst("peso") & "','" & KEY_USUARIO & "','" & rst("id_detalle_serie") & "','" & rst("serie") & "','" & rst("anio_fabricacion") & "','" & rst("nro_chasis") & "','" & rst("anio_modelo") & "','" & rst("nro_dua") & "','" & rst("nro_item") & "','" & get_servicio(rst("id_producto")) & "','" & get_costo_producto(rst("id_producto")) & "','" & rst("obsequio") & "','" & rst("icbper") & "')"
       CnBd.Execute (strCadena)
       
        If KEY_BONIFICACIONES = "si" Then
            If Me.DtcTipoDoc.BoundText <> "0007" Then
                strCadena = "CALL get_idTemporalventas('" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
                Call ConfiguraRstC(strCadena)
                in_idVenta = rstc(0)
            
                strCadena = "CALL put_bonificacion_linea('" & codigoP & "','" & Trim(Me.TxtCodCliente.Text) & "','" & KEY_USUARIO & "','" & KEY_ALM & "','" & in_idVenta & "','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
            
                Call put_verificar_bonificacion_monto(codigoP, cantidadP, Trim(Me.TxtCodCliente.Text), Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText)
                Call put_verificar_bonificacion_cruzada_v2(codigoP, cantidadP, Trim(Me.TxtCodCliente.Text), Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText)
            End If
        End If
        
         
        total_temp = total_temp + in_precio_venta * cantidadP
siguiente:
        rst.MoveNext
    Next i
End If

 

End Sub

Private Sub llenar_comprobante_referencia(ByVal idVenta As Double)
Dim total_temp As Double
Dim i As Integer
Dim in_precio_venta As Double
Dim codigoP As String
Dim cantidadP As Single


strCadena = "SELECT * FROM view_detalle_comprobante WHERE id_venta='" & idVenta & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then

strCadena = "DELETE FROM temporal_ventas WHERE ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "' AND id_alm='" & KEY_ALM & "' "
CnBd.Execute (strCadena)
'
total_temp = 0
rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        codigoP = rst("id_producto")
        cantidadP = rst("cantidad")
        in_precio_venta = rst("precio")
        strCadena = "INSERT INTO temporal_ventas(ruc,id_alm,id_dni,id_doc,id_serie,numero,id_producto,detalle,cantidad,precio,total,peso,dni_save,id_detalle_serie,serie,anio_fabricacion,nro_chasis,anio_modelo,nro_dua,nro_item,servicio,costo,obsequio,icbper) VALUES " & _
        "('" & KEY_RUC & "','" & Me.DtcAlmacen.BoundText & "','" & Trim(Me.TxtCodCliente.Text) & "','" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcSerieDoc.BoundText & "','" & Me.TxtNumeroDoc.Text & "','" & rst("id_producto") & "','" & rst("detalle") & "','" & rst("cantidad") & "'," & _
        "'" & in_precio_venta & " ','" & in_precio_venta * rst("cantidad") & "','" & rst("peso") & "','" & KEY_USUARIO & "','" & rst("id_detalle_serie") & "','" & rst("serie") & "','" & rst("anio_fabricacion") & "','" & rst("nro_chasis") & "','" & rst("anio_modelo") & "','" & rst("nro_dua") & "','" & rst("nro_item") & "','" & get_servicio(rst("id_producto")) & "','" & get_costo_producto(rst("id_producto")) & "','" & rst("obsequio") & "','" & rst("icbper") & "')"
        CnBd.Execute (strCadena)
        total_temp = total_temp + in_precio_venta * rst("cantidad")
        rst.MoveNext
    Next i
End If

 

End Sub






Private Function get_servicio(ByVal in_producto As String) As String
strCadena = "SELECT servicio FROM view_tipo_servicio WHERE   id_producto='" & in_producto & "'   and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
    get_servicio = rstA("servicio")


    
End If


End Function
Private Sub txtNumero_nota_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtNumero_nota.Text = FormatosCeros(txtNumero_nota, 6)
    Me.TxtMontoPagado.Text = get_saldo_nota_credito(Trim(Me.txtserie_nota.Text), Trim(Me.txtNumero_nota.Text), Trim(Me.TxtCodCliente.Text))
    If Val(Me.TxtMontoPagado.Text) > 0 Then
         Call realizar_ingreso_pago
    End If
End If

End Sub
Private Function get_saldo_nota_credito(ByVal in_serie As String, ByVal in_numero As String, ByVal in_dni As String) As Single
strCadena = "SELECT (total-function_pago_factura(id_venta,'" & Format(KEY_FECHA, "YYYY-mm-dd") & "',id_moneda,ruc))as saldo,fecha_emision FROM movimiento_venta WHERE id_cliente='" & Trim(Me.TxtCodCliente.Text) & "' and  id_doc='0007' and  serie='" & Trim(Me.txtserie_nota.Text) & "' and numero='" & Trim(Me.txtNumero_nota.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    get_saldo_nota_credito = True
   
      Me.Label22(3).Caption = Format(rstL("saldo"), "###0.00")
      get_saldo_nota_credito = rstL("saldo")
      
 
      
  
   
   
Else
    MsgBox "NOTA DE CREDITO NO REGISTRADA", vbInformation, KEY_VENDEDOR
End If

End Function

Private Sub TxtNumeroDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        'Call Resalta(Me.TxtSerie)
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.TxtCodCliente)
End If
End Sub

Private Sub TxtNumeroDoc_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    
    
    Call mostrar_comprobante(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text))
'    Me.cmdeliminar.Enabled = True
End If

 
 
 
End Sub
Public Sub get_update_proform()
Call enabled_form(FrmVentas)
Me.cmdProcesar.Enabled = True
Me.cmdModificar.Enabled = False
Me.TxtCodProducto.Enabled = True
Me.txtCantidad.Enabled = True
Me.txtprecio.Enabled = True
Me.cmdAgregar.Enabled = True
Me.CmdQuitar.Enabled = True


End Sub

Public Sub get_comprobante_electronico(ByVal in_doc As String, ByVal in_serie As String)

strCadena = "SELECT electronico FROM almacen_comprobante WHERE   id_doc='" & in_doc & "' and serie='" & in_serie & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
    If rstA("electronico") = "si" Then
        If KEY_USUARIO = "42546269" Or KEY_USUARIO = "46947665" Or KEY_USUARIO = "900001" Then
            Me.cmdAnular.Enabled = True
            Me.cmdEliminar.Enabled = True
        Else
            If KEY_CARGO = "00004" Then
                Me.cmdAnular.Enabled = True
            Else
                Me.cmdAnular.Enabled = False
            End If
        End If
        
        
    Else
        Me.cmdAnular.Enabled = True
        
         If KEY_USUARIO = "42546269" Or KEY_USUARIO = "46947665" Or KEY_USUARIO = "900001" Then
            Me.cmdEliminar.Enabled = True
         Else
            Me.cmdEliminar.Enabled = False
         End If
        
    End If
End If


End Sub


Public Sub mostrar_comprobante(ByVal n_doc As String, ByVal nserie As String, ByVal numerot As String)
    Dim montot As Double
    Dim idVenta As Double
    Dim nnumero As String

    Me.DtcTipoDoc.BoundText = n_doc
    Me.DtcSerieDoc.BoundText = Format(nserie, "000")
    Me.TxtNumeroDoc.Text = FormatosCeros(numerot, 6)
    nnumero = Trim(Me.TxtNumeroDoc.Text)
    
    strCadena = "SELECT * FROM movimiento_venta WHERE (numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND serie='" & Trim(Me.DtcSerieDoc.BoundText) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND ruc='" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        If correlativo_comprobante(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText), Trim(Me.TxtNumeroDoc.Text)) = True Then
            MsgBox "Comprobante Consultado No esta Emitido :" + Chr(13) + "Imposible Ingresar un comprobante con fecha ANTERIOR" + Chr(13) + Chr(13) + "Ingrese a la fecha de trabajo Correcta.", vbInformation
            Call nuevo
            Exit Sub
        End If
        Call nuevo
        
        Me.TxtNumeroDoc.Text = nnumero
        Exit Sub
   
    Else
        
        
        
        MsgBox "VISUALIZAR DOCUMENTO", vbInformation, KEY_VENDEDOR
        Me.txtidVenta.Text = rst("id_venta")
        
        
        If get_cuotas(Me.txtidVenta.Text) = True Then
           Me.cmdCuotas.Visible = True
        End If
        Me.txtObservacion.Text = rst("observacion")
        Me.lblPercepcion.Caption = Format(rst("percepcion"), "#,##0.00")
        Me.DtcVendedor.BoundText = rst("id_vendedor")
        If Me.DtcTipoDoc.BoundText = "0099" And KEY_UPDATE_PROFORM = "si" Then
           Me.cmdModificar.Enabled = True
        Else
            Me.cmdModificar.Enabled = False
        End If
        If rst("diferida") = "si" Then
            Me.chk_venta_diferida.Value = 1
        Else
            Me.chk_venta_diferida.Value = 0
        End If
        
        
        If rst("afecto_detraccion") = "si" Then
           Me.chk_detraccion.Value = 1
        Else
           Me.chk_detraccion.Value = 0
        End If
        
        If Len(Trim(rst("orden_compra"))) > 2 Then
           Me.chk_OrdenCompra.Value = 1
           Me.txtOrdenCompra.Text = rst("orden_compra")
        Else
            Me.chk_OrdenCompra.Value = 0
           Me.txtOrdenCompra.Text = ""
        End If
        
        
        
        
        
        get_verificar_tiene_nota (Val(Me.txtidVenta.Text))

        
        If KEY_SERVIDOR_KEYFACIL = "si" Then
            Me.txt_hash.Text = rst("sunat_hash")
            If get_firma_online(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText)) = "si" And Len(Trim(rst("sunat_hash"))) < 5 Then
                Me.cmdProcesar.Enabled = True
            Else
                Me.cmdProcesar.Enabled = False
            End If
            
        Else
        If get_firma_online(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText)) = "si" And Len(Trim(rst("sunat_key"))) < 5 Then
           Me.cmdProcesar.Enabled = True
           
        Else
            Me.cmdProcesar.Enabled = False
            
            
            Me.cmdAnular.Enabled = True
        End If
        End If
        
        
        
        
        If get_dar_baja(Me.txtidVenta.Text) = True Then
              Me.cmdAnular.Caption = "DAR BAJA"
              Me.cmdAnular.Enabled = True
         End If
        Me.chkvincular.Visible = True
        Me.lblregistradopor.Caption = "[" & Trim(Me.txtidVenta.Text) & "]  " & Mid(get_persona(rst("dni_save")), 1, 20) & Space(2) & Format(rst("hora"), "HH:mm:ss AM/PM")
        Me.txttipofactura.Text = rst("id_tipo_factura")
        
        If rst("id_recibo") > 0 Then
            strCadena = "SELECT serie,numero,id_vendedor,id_tipo_factura FROM movimiento_venta WHERE id_venta='" & rst("id_recibo") & "' AND ruc='" & KEY_RUC & "' ORDER BY id_venta DESC LIMIT 1"
            Call ConfiguraRstK(strCadena)
            'Me.txtSerieRecibo.Text = rstK("serie")
           ' Me.txtNumeroRecibo.Text = rstK("numero")
            Me.DtcVendedor.BoundText = rstK("id_vendedor")
            
            Me.DtcVendedor.Locked = True
          '  Me.cmdGrabarRecibo.Visible = True
          ' Me.cmdImprimirRecibo.Visible = True
          '  Me.txtSerieRecibo.Visible = True
          '  Me.txtNumeroRecibo.Visible = True
         '   Me.txtSerieRecibo.Locked = False
           ' Me.txtNumeroRecibo.Locked = False
        End If
        
         If rst("id_doc") = "0007" Then
            Me.txtmotivo_nota.Text = rst("motivo_nota")
            Me.DtcTipoNota.BoundText = rst("id_tipo_nota")
            strCadena = "SELECT documento FROM movimiento_venta WHERE id_venta='" & rst("id_comprobante") & "' "
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount > 0 Then
                Me.txtdocreferencia.Text = rstK("documento")
                Me.FrameReferencia.Visible = True
            End If
          End If
          
          If rst("id_doc") = "0008" Then
            Me.txtmotivo_nota.Text = rst("motivo_nota")
            Me.DtcTipoNota.BoundText = rst("id_tipo_nota")
            strCadena = "SELECT documento FROM movimiento_venta WHERE id_venta='" & rst("id_comprobante") & "'  and ruc='" & KEY_RUC & "' "
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount > 0 And rst("id_comprobante") > 0 Then
                Me.txtdocreferencia.Text = rstK("documento")
                Me.FrameReferencia.Visible = True
            End If
          End If
        
        If rst("id_doc") = "0054" Then
            Me.lblContabilidad.Caption = rst("observacion")
        End If
        Me.lblContabilidad.Visible = False
        idVenta = rst("id_venta")
        montot = 0
        Me.TxtCodCliente.Text = rst("id_cliente")
       
        If rst("descuento") > 0 Then
            Me.chk_descuento.Value = 1
        End If
        
        If rst("anulado") = "si" Then
            Me.lblAnulado.Visible = True
            Me.cmdAnular.Enabled = False
            
            If Me.DtcTipoDoc.BoundText = "0099" Then
               Me.cmdEliminar.Enabled = False
            Else
                Me.cmdEliminar.Enabled = True
            End If
            
           
            
            
        Else
            Me.lblAnulado.Visible = False
            
            If KEY_FACTURACION_ELECTRONICA = "si" Then
               ' Me.cmdAnular.Enabled = False
                'Me.cmdeliminar.Enabled = False
            Else
                Me.cmdAnular.Enabled = True
                Me.cmdEliminar.Enabled = True
            End If
                'Me.cmdEditable.Enabled = False
            
            
            End If
        
        
                
            
        End If
        If rst("afecta_factura") = "si" Then
            Me.chkvincular.Visible = False
            Me.chk_factura.Value = 1
        Else
            Me.chk_factura.Value = 0
        End If
        
       If Len(rst("id_copropietario")) > 1 Then
           Me.txtdni_copropietario.Text = rst("id_copropietario")
           Me.lblcopropietario.Caption = get_persona(rst("id_copropietario"))
       End If
        in_total_documento = 0
        Me.lblTotal.Caption = Format(rst("total"), "###0.00")
        Me.DtcAlmacen.BoundText = rst("id_alm")
        Me.lblPago.Caption = Format(rst("monto_pago"), "###0.00")
        Me.lblVuelto.Caption = Format(rst("monto_vuelto"), "###0.00")
        Me.lblExonerado.Caption = Format(rst("exonerado"), "###0.00")
        Me.LblValorVenta.Caption = Format(rst("valor_venta"), "#,##0.00")
        Me.LblIgv.Caption = Format(rst("igv"), "#,##0.00")
        Me.lblicbper.Caption = Format(rst("icbper"), "#,##0.00")
        Me.TxtDescuento_global.Text = Format(rst("descuento"), "###0.00")
        'Me.TxtDescuento_porcentaje.Enabled = False
        
        
        
        Me.TxtDescuento_porcentaje.Text = Format(rst("descuento_porcentaje"), "#,##0.00000")
        
        Me.chk_descuento.Enabled = False
        Me.DtcMoneda.BoundText = rst("id_moneda")
        
        
                
        
        Call LlenarDatosCliente(idVenta)
        
        Call llenarGrid_Comprobante(Me.HfdDetalle, idVenta)
        Call llena_pagosVenta(Me.HfgTipoPagos, idVenta)
        Call get_comprobante_electronico(Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.Text)
        
        If Val(Me.TxtDescuento_global.Text) > 0 Then
            
        End If
        
        
        
        Me.TxtCodProducto.Enabled = False
        Me.TxtDescripcionProducto.Enabled = False
        Me.txtprecio.Enabled = False
        Me.cmdAgregar.Enabled = False
        Me.CmdQuitar.Enabled = False
        Me.txtCantidad.Enabled = False
        
        Me.TxtCodProducto.Enabled = False
        Me.HfdDetalle.SetFocus
End Sub
Private Sub VerificaAnulado(ByVal idVenta As Double)
strCadena = "Select Anulado FROM DocumentoVenta WHERE idVenta='" & idVenta & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    If Trim(rst(0)) = "V" Then
        Me.lblAnulado.Visible = True
        Me.cmdAnular.Enabled = False
        Me.cmdEliminar.Enabled = False
        
    Else
        Me.cmdAnular.Enabled = True
        
    End If
End If
Set rst = Nothing
End Sub
Private Sub LLenarDatosReferencia(ByVal tipo_doc As String, ByVal numero As String, ByVal serie As String)
strCadena = "SELECT * FROM movimiento_venta_targeta WHERE id_doc='" & tipo_doc & "' AND numero='" & numero & "' AND serie='" & Trim(serie) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.DtTargeta.BoundText = rst("id_targeta")
    Me.TxtNumeroTargeta.Text = Trim(rst("numero_tarjeta"))
End If
Set rst = Nothing
End Sub
Private Sub LlenarDatosCliente(ByVal idVenta As Double)
Dim CodPersona As String
Dim Nombre As String
strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & idVenta & "' AND ruc='" & KEY_RUC & "' LIMIT 1 "
Call ConfiguraRst(strCadena)



    
    Me.DtpActual.Value = CVDate(rst("fecha_emision"))
    Me.DtpFechaReferencia.Value = CVDate(rst("fecha_vencimiento"))
    CodPersona = rst("id_cliente")
    Nombre = rst("ncliente")
    
    Me.TxtCodCliente.Text = rst("id_cliente")
    Me.TxtCliente.Text = rst("ncliente")
    Me.txtDireccion.Text = rst("direccion")
    Me.txtTelefono.Text = get_telefono(Me.TxtCodCliente.Text)
    Me.txtmail.Text = get_mail(Me.TxtCodCliente.Text)
    
    Me.DtcFormaPago.BoundText = rst("id_forma_pago")
    txtExtranjero.Text = get_extrangero(Me.TxtCodCliente.Text)
       
    
    
    
    If Trim(CodPersona) = "00000000" Or CodPersona = "" Then
        Me.TxtCodCliente.Text = "00000000"
        Me.TxtCliente.Text = UCase(Nombre)
        Me.txtDireccion.Text = KEY_DIR_PUBLIC
        Me.cmdImprimir.Enabled = True
        
        Exit Sub
    End If

Me.cmdImprimir.Enabled = True
End Sub


Private Sub LlenarDatosPedido(ByVal idVenta As Double)
Dim CodPersona As String
Dim Nombre As String
strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & idVenta & "' AND ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
    
    Me.DtpActual.Value = KEY_FECHA
    Me.DtpFechaReferencia.Value = KEY_FECHA
    CodPersona = rst("id_cliente")
    Me.TxtCodCliente.Text = rst("id_cliente")
    Me.TxtCliente.Text = rst("ncliente")
    Me.txtDireccion.Text = get_direccion(rst("id_cliente"))
    Me.DtcFormaPago.BoundText = rst("id_forma_pago")
    If Trim(CodPersona) = "00000000" Or CodPersona = "" Then
        Me.TxtCodCliente.Text = "00000000"
        Me.TxtCliente.Text = UCase(rst("ncliente"))
        Me.txtDireccion.Text = KEY_DIR_PUBLIC
        Me.cmdImprimir.Enabled = True
        
        Exit Sub
    End If

Me.cmdImprimir.Enabled = True
End Sub
Private Sub LlenarDatosCliente_referencia(ByVal idVenta As Double)
Dim CodPersona As String
Dim Nombre As String
strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & idVenta & "' AND ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    CodPersona = rst("id_cliente")
    Me.TxtCodCliente.Text = rst("id_cliente")
    Me.TxtCliente.Text = rst("ncliente")
    Me.txtDireccion.Text = rst("direccion")
    Me.txtObservacion.Text = rst("observacion")
    Me.DtcFormaPago.BoundText = rst("id_forma_pago")
    If Trim(CodPersona) = "00000000" Or CodPersona = "" Then
        Me.TxtCodCliente.Text = "00000000"
        Me.TxtCliente.Text = UCase(rst("ncliente"))
        Me.txtDireccion.Text = KEY_DIR_PUBLIC
        Me.cmdImprimir.Enabled = True
        
        Exit Sub
    End If
End If

End Sub

Private Sub txtNumeroRecibo_Change()

End Sub

Private Sub TxtNumeroTargeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.TxtNumeroTargeta.Text) <> "" Then
            If Me.frminteres.Visible = True Then
              Call put_incrementar_interes_tarjeta(Val(Me.txtporcentaje_interes.Text), Val(Me.TxtMontoPagado.Text))
            End If
            Call Resalta(Me.TxtOperacion)
        Else
            Call Resalta(Me.TxtNumeroTargeta)
        End If
    End If
End Sub



Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    
    
   Call Resalta(Me.TxtCodProducto)
    
End If
End Sub



Private Sub TxtOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Me.TxtOperacion.Text) <> "" Then
        Call Resalta(Me.TxtMontoPagado)
        
    Else
        Call Resalta(Me.TxtOperacion)
    End If
End If
End Sub

Private Sub put_incremento_interes()



Call put_interes_venta_credito(Val(Me.TxtMontoPagado.Text), Val(Me.lblTotal.Caption), Val(Me.lblPago.Caption), Val(Me.txtporcentaje_interes.Text))

Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcSerieDoc.BoundText, Me.DtcTipoDoc.BoundText, Me.txtformato_impresion.Text)

End Sub


Private Sub txtporcentaje_interes_KeyPress(KeyAscii As Integer)
Dim in_forma_detalle As String
Dim in_porcentaje As Single
If KeyAscii = 13 Then
        in_forma_detalle = Me.DtcFormapagodetalle.BoundText
        If Val(Me.TxtCuotas.Text) = 0 Then
            Me.TxtCuotas.Text = 1
        End If
        in_porcentaje = Val(Me.txtporcentaje_interes.Text) * Val(Me.TxtCuotas.Text)
        Me.txtporcentaje_interes.Text = in_porcentaje
        Call put_incremento_interes
        Call put_forma_pago
        Me.DtcFormapagodetalle.BoundText = in_forma_detalle
        Me.txtporcentaje_interes.Text = in_porcentaje
        Call Resalta(TxtCuotas)

   
   
End If
End Sub
Private Sub put_interes_nota_credito()

End Sub

Private Sub TxtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtDescripcionProducto)
       
End If
If KeyCode = vbKeyRight Then
     Me.cmdAgregar.SetFocus
End If
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
Dim TotalP As Single
If KeyAscii = 13 Then
        
        
        If Me.chkPrecios.Value = 1 And Me.HfPrecios.Rows > 0 And Trim(Me.txtmayor.Text) = "si" Then
           If validar_precio(Format(Me.HfPrecios.TextMatrix(Me.HfPrecios.Row, 1), "###00.00"), Val(Me.txtprecio.Text)) = True Then
                Exit Sub
            Else
                Procedencia = modificar_precio
                frmsegurity.Show
                frmsegurity.txtMotivo.Text = "AUTORIZACION CAMBIO PRECIO"
                Exit Sub
            End If
        Else
           
           If KEY_GRIFO = "si" And Me.DtcTipoDoc.BoundText <> "0108" Then
                If Val(Me.txtCantidad.Text) = 1 Then
                    Me.txtCantidad = Val(Me.txtprecio.Text) / Val(Me.txtpreciooriginal.Text)
                    Me.txtprecio.Text = Val(Me.txtpreciooriginal.Text)
                End If
                
                GoTo insertar
           Else
           If validar_precio(Val(Me.txtpreciooriginal.Text), Val(Me.txtprecio.Text)) = True Then
               
               GoTo insertar
           Else
                Call disabled_form(Me)
                Procedencia = modificar_precio
                frmsegurity.Show
                Exit Sub
           End If
           End If
           
           
           
           
        End If
        
insertar:
            TotalP = Val(Me.txtCantidad.Text) * Val(Me.txtprecio.Text)
            Me.LblTotalParcial.Caption = Format(TotalP, "###0.00")
            Call insertar_codigo
        
End If
End Sub




Function validar_precio(ByVal precio_base As Single, ByVal precio_venta As Single) As Boolean
        If precio_venta < precio_base Then
            If KEY_CARGO = "00004" Then
               If MsgBox("PRECIO INGRESADO ES MENOR DEL MINIMO." + Chr(13) + Chr(13) + "DESEA CONTINUAR .", vbQuestion + vbYesNo, "SR. " & Space(1) & KEY_VENDEDOR) = vbYes Then
                  Me.txtprecio.Text = precio_venta
                  Call Resalta(Me.txtprecio)
                  validar_precio = False
                  Exit Function
                  
               Else
                  Me.txtprecio.Text = precio_base
                  Call Resalta(Me.txtprecio)
                  validar_precio = True
               End If
                
            Else
                If MsgBox("PRECIO INGRESADO ES MENOR DEL MINIMO." + Chr(13) + Chr(13) + "Necesita Password [ ADMINISTRADOR ]" + Chr(13) + "DESEA CONTINUAR.", vbInformation + vbYesNo, "SR. " & Space(1) & KEY_VENDEDOR) = vbYes Then
                    validar_precio = False
                    Exit Function
                Else
                        Me.txtprecio.Text = precio_base
                        Call Resalta(Me.txtprecio)
                        validar_precio = True
                End If
                
            End If
            
        Else
            validar_precio = True
        End If
    
End Function
Private Sub save_targeta(ByVal id_doc As String, ByVal Documento As String, ByVal serie As String, ByVal id_targeta As String, ByVal numero As String)
    strCadena = "INSERT INTO DocumentoVenta_Targeta VALUES ('" & Trim(id_doc) & "','" & Trim(Documento) & "','" & Trim(serie) & "','" & id_targeta & "','" & numero & "')"
    CnBd.Execute (strCadena)
    '
     
End Sub

Public Sub put_cancelar_recibo(ByVal in_venta_ref As String)


Call put_realizar_pago(Val(Me.txtidVenta.Text), Val(Me.txtid_venta_ref.Text), Val(Me.lblPago.Caption), Me.DtcTipoDoc.BoundText, Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))

End Sub

Public Sub put_cancelar_recibo_masivos()

strCadena = "SELECT id_venta,documento,(total-pago) as saldo FROM view_recibo_pago WHERE seleccion='si' and dni_use='" & KEY_USUARIO & "' and  id_doc='0054' and id_cliente='" & Trim(Me.TxtCodCliente.Text) & "'  and ruc='" & KEY_RUC & "'"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
   rstIN.MoveFirst
   For i = 0 To rstIN.RecordCount - 1
        
        Call put_realizar_pago(Val(Me.txtidVenta.Text), rstIN("id_venta"), rstIN("saldo"), Me.DtcTipoDoc.BoundText, Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
        rstIN.MoveNext
   Next i
End If


End Sub

Public Function Save() As Boolean
'On Error GoTo error
Dim i As Integer, anul As String * 2, MontoActual As Double, TotalVenta As Double
Dim igv As Double, SUBTOTAL As Double, exonerado As Double, dfac As String, Monto_descuento As Single
Dim monto_pagado As Double, Monto_Vuelto As Double, Monto_Sobrante As Double, saldo_f As Double, estado_f As String
Dim id_venta  As Double, CodReferencia As String, KEY_VENCIMIENTO As String, cod_cliente As String, rst1 As New ADODB.Recordset, p As Integer
Dim horario As String, turno As String
Dim id_tipo_factura As String
Dim in_monto_saldo As Double
Dim in_factor_descuento As Double
Dim in_descuento_parcial As Single
Dim in_total_temporal As Double
Save = False


If KEY_MOSTRAR_SUCURSAL = "si" Then
    If Trim(Me.txtObservacion.Text) <> "" Then
        If KEY_DIRECCION <> KEY_DIRECCION_ALM Then
            Me.txtObservacion.Text = KEY_DIRECCION_ALM & Space(2) & Trim(Me.txtObservacion.Text)
                    
        End If
        
    Else
        If KEY_DIRECCION <> KEY_DIRECCION_ALM Then
            Me.txtObservacion.Text = "SUC:" & KEY_DIRECCION_ALM
                    
        End If
    End If
End If



horario = Format(Time, "hh:mm")
If horario >= "07:00" And horario <= "13:00" Then
   turno = "M"
Else
   turno = "T"
End If

If Me.chk_detraccion.Value = 1 Then
   in_detraccion = "si"
Else
   in_detraccion = "no"
End If



If Trim(Me.TxtCodCliente.Text) = "" Then
    Me.TxtCodCliente.Text = "00000000"
End If

SUBTOTAL = Val(Me.LblValorVenta.Caption)
igv = Val(Me.LblIgv.Caption)
exonerado = Val(Me.lblExonerado.Caption)
TotalVenta = Val(Format(Me.lblTotal.Caption, "###0.000"))
Monto_descuento = Val(Me.TxtDescuento_global.Text)
monto_pagado = Val(Me.lblPago.Caption)
Monto_Vuelto = Val(Me.lblVuelto.Caption)
Monto_Sobrante = 0

If Me.chkconyuge.Value = 1 Then
    strconyugue = "si"
Else
    strconyugue = "no"
End If
If KEY_SKFACTURA = "si" Then
    If Me.chk_factura.Value = 1 Then
        dfac = "si"
    Else
        dfac = "no"
    End If
Else
        dfac = "no"
End If
If (Trim(Me.DtcFormaPago.BoundText) = "05") Then
    If (Trim(Me.TxtCodCliente.Text) = "00000000") Then
        MsgBox "Elija un Cliente Registrado, para dar Credito", vbInformation, "Mensaje de Administracion"
        Call Resalta(Me.TxtCodCliente)
        Exit Function
    End If
    saldo_f = TotalVenta
    estado_f = "Credito"
Else
    saldo_f = KEY_NULO
End If
cod_cliente = Trim(Me.TxtCodCliente.Text)

strCadena = "SELECT * FROM movimiento_venta_monto_temporal WHERE id_alm='" & KEY_ALM & "' and fecha='" & KEY_FECHA & "' and  id_usuario='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'  ORDER BY id_monto ASC"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
        rstK.MoveFirst
        If Me.chk_descuento.Value = 1 And Val(Me.TxtDescuento_porcentaje.Text) = 100 Then
            GoTo avanzar2
        End If
        If Val(Monto_Vuelto) < 0 And Me.DtcTipoDoc.BoundText <> "0007" Then
            MsgBox "El Monto Pagado es Inferior al Monto Total", vbInformation, "Mensaje para el Usuario"
            Call Resalta(Me.TxtMontoPagado)
            Exit Function
        End If
        
avanzar2:
        If Me.DtcFormaPago.BoundText = "05" Then
            KEY_VENCIMIENTO = Format(Me.DtpFechaReferencia.Value, "yyyy-mm-dd")
        Else
            rstK.MoveFirst
            Saldo = 0
            For i = 0 To rstK.RecordCount - 1
                If rstK("forma_pago") = "08" Then
                    Saldo = rstK("monto")
                End If
                rstK.MoveNext
            Next i
            KEY_VENCIMIENTO = KEY_FECHA
        End If
            
            
            id_tipo_factura = Trim(Me.txttipofactura.Text)
            If Me.txteditable.Text = "si" Then
                   id_tipo_factura = "00003"
            End If
            
            If Trim(Me.txtformato_impresion.Text) <> "4" Then
                 If Me.chk_manual.Value = 1 Then
                   Me.TxtNumeroDoc.Text = Trim(Me.TxtNumeroDoc.Text)
                Else
                    Me.TxtNumeroDoc.Text = get_numero_comprobante(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText))
                End If
                
                
                Documento = Trim(Me.DtcTipoDoc.Text) & ":" & Trim(Me.DtcSerieDoc.BoundText) & "-" & Trim(Me.TxtNumeroDoc.Text)
            Else
                If Me.chk_manual.Value = 1 Then
                   Me.TxtNumeroDoc.Text = Trim(Me.TxtNumeroDoc.Text)
                Else
                    Me.TxtNumeroDoc.Text = Trim(Me.TxtNumeroDoc.Text)
                End If
                
            End If

            If Me.chk_seleccionar_guia.Value = 1 Then
               in_guia = Me.DtcGuia.Text
               id_guia = Me.DtcGuia.BoundText
            Else
               in_guia = 0
               id_guia = 0
            End If
            
            If txt_tipo.Text = "" Then
                txt_tipo.Text = "01"
            End If
            
            
            If Me.chkseguro.Value = 1 Then
               in_seguro = Me.Dtcseguro.BoundText
            Else
               in_seguro = "0"
            End If
            
            
            Call validar_cliente(Trim(Me.TxtCodCliente.Text))
            
            
            If Trim(Me.txtServicio.Text) = "no" Then
                in_cta_cobrar = KEY_CTA_COBRAR_PRODUCTO
                in_cta_ingreso = KEY_CTA_INGRESO_PRODUCTO
            Else
                in_cta_cobrar = KEY_CTA_COBRAR_SERVICIO
                in_cta_ingreso = KEY_CTA_INGRESO_SERVICIO
            End If
            
            If Me.chk_interes.Value = 1 Then
                in_interes = Val(Me.txtporcentaje_interes.Text)
            Else
                in_interes = 0
            End If
            If Me.chk_venta_diferida.Value = 1 Then
               in_diferida = "si"
            Else
               in_diferida = "no"
            End If
            
          
            
            If Me.DtcFormaPago.BoundText = "02" Then ' credito
                Saldo = TotalVenta
            End If
            
            
            
            If Me.DtcFormaPago.BoundText = "02" And Val(Me.TxtCuotas.Text) < 1 Then
                KEY_VENCIMIENTO = Format(Me.txtFecha_vencimiento.Text, "YYYY-mm-dd")
            End If
            
            
Iniciar:
            
            
            
            If Me.chk_manual.Value = 1 Then
                   Me.TxtNumeroDoc.Text = Trim(Me.TxtNumeroDoc.Text)
                   If verificar_duplicado(Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText, Trim(Me.TxtNumeroDoc.Text)) = True Then
                        MsgBox "COMPROBANTE YA REGISTRADO, VERIFIQUE EL CORRELATIVO" + Chr(13) + Chr(13) + "OBSERVE EL COMPROBANTE FISICO.", vbInformation, KEY_EMPRESA
                        Save = False
                        Exit Function
                    End If
            Else
                   Me.TxtNumeroDoc.Text = get_numero_comprobante(Me.DtcTipoDoc.BoundText, Trim(Me.DtcSerieDoc.BoundText))
            End If
            
            Documento = Trim(Me.DtcTipoDoc.Text) & ":" & Me.DtcSerieDoc.BoundText & "-" & Trim(Me.TxtNumeroDoc.Text)
                
             
             
    
            
            If KEY_PAIS = KEY_PERU Then
                
                'strCadena = "call p_insert_venta_cabecera_v11('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcFormaPago.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
                "'" & Trim(Me.DtcSerieDoc.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Me.TxtCodCliente.Text & "','" & Me.txtCliente.Text & "','" & SUBTOTAL & "','" & igv & "','" & exonerado & "','" & TotalVenta & "','" & Saldo & "'," & _
                "'" & Val(Me.lblPago.Caption) & "','" & Val(Me.lblVuelto.Caption) & "','" & KEY_FECHA & "','" & KEY_VENCIMIENTO & "','" & id_tipo_factura & "','" & Me.DtcVendedor.BoundText & "','" & KEY_USUARIO & "','" & Val(KEY_CAMBIO_COMPRA) & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "'" & _
                ",'" & Documento & "','" & horario & "','T','" & Trim(Me.txtdireccion.Text) & "','" & strconyugue & "','" & Trim(Me.txt_hash.Text) & "','" & Trim(Me.txt_sunat_key.Text) & "','" & Trim(Me.DtcTipoNota.BoundText) & "','" & Trim(Me.txtmotivo_nota.Text) & "','" & id_guia & "','" & in_guia & "','" & KEY_VENTANILLA & "','" & Trim(Me.txt_tipo.Text) & "','" & in_seguro & "','" & Trim(Me.txtObservacion.Text) & "','" & Trim(Me.txteditable.Text) & "','" & KEY_CONTABILIDAD & "','" & in_cta_cobrar & "','" & in_cta_ingreso & "','" & Val(Me.TxtDescuento_global.Text) & "','" & Val(Me.TxtCuotas.Text) & "','" & in_interes & "','" & Val(Me.txtid_venta_ref.Text) & "','" & in_diferida & "','" & Val(Me.txt_id_pendiente.Text) & "','" & KEY_RUC & "')"
                
                strCadena = "call p_insert_venta_cabecera_v16('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcFormaPago.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
                "'" & Trim(Me.DtcSerieDoc.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Me.TxtCodCliente.Text & "','" & Me.TxtCliente.Text & "','" & SUBTOTAL & "','" & igv & "','" & exonerado & "','" & TotalVenta & "','" & Saldo & "'," & _
                "'" & Val(Me.lblPago.Caption) & "','" & Val(Me.lblVuelto.Caption) & "','" & KEY_FECHA & "','" & KEY_VENCIMIENTO & "','" & id_tipo_factura & "','" & Me.DtcVendedor.BoundText & "','" & KEY_USUARIO & "','" & Val(KEY_CAMBIO_VENTA) & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "'" & _
                ",'" & Documento & "','" & horario & "','T','" & Trim(Me.txtDireccion.Text) & "','" & strconyugue & "','" & Trim(Me.txt_hash.Text) & "','" & Trim(Me.txt_sunat_key.Text) & "','" & Trim(Me.DtcTipoNota.BoundText) & "','" & Trim(Me.txtmotivo_nota.Text) & "','" & id_guia & "','" & in_guia & "','" & KEY_VENTANILLA & "'," & _
                " '" & Trim(Me.txt_tipo.Text) & "','" & in_seguro & "','" & Trim(Me.txtObservacion.Text) & "','" & Trim(Me.txteditable.Text) & "','" & KEY_CONTABILIDAD & "','" & in_cta_cobrar & "','" & in_cta_ingreso & "','" & Val(Me.TxtDescuento_global.Text) & "','" & Val(Me.TxtCuotas.Text) & "','" & in_interes & "','" & Val(Me.txtid_venta_ref.Text) & "','" & in_diferida & "','" & Val(Me.txt_id_pendiente.Text) & "','" & KEY_SIN_EFECTO_CAJA & "','" & Val(Me.lblicbper.Caption) & "','" & Trim(Me.txtOrdenCompra.Text) & "','" & in_detraccion & "','" & KEY_RUC & "')"
                
                
            Else
                strCadena = "call p_insert_venta_cabecera_premiun_internacional('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcFormaPago.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
                "'" & Trim(Me.DtcSerieDoc.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Me.TxtCodCliente.Text & "','" & Me.TxtCliente.Text & "','" & SUBTOTAL & "','" & igv & "','" & exonerado & "','" & TotalVenta & "','" & Saldo & "'," & _
                "'" & Val(Me.lblPago.Caption) & "','" & Val(Me.lblVuelto.Caption) & "','" & KEY_FECHA & "','" & KEY_VENCIMIENTO & "','" & id_tipo_factura & "','" & Me.DtcVendedor.BoundText & "','" & KEY_USUARIO & "','" & Val(KEY_CAMBIO_COMPRA) & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "'" & _
                ",'" & Documento & "','" & horario & "','T','" & Trim(Me.txtDireccion.Text) & "','" & strconyugue & "','" & Trim(Me.txt_hash.Text) & "','" & Trim(Me.txt_sunat_key.Text) & "','" & Trim(Me.DtcTipoNota.BoundText) & "','" & Trim(Me.txtmotivo_nota.Text) & "','" & id_guia & "','" & in_guia & "','" & KEY_VENTANILLA & "','" & Trim(Me.txt_tipo.Text) & "','" & in_seguro & "','" & Trim(Me.txtObservacion.Text) & "','" & Trim(Me.txteditable.Text) & "','" & KEY_CONTABILIDAD & "','" & in_cta_cobrar & "','" & in_cta_ingreso & "','" & Val(Me.TxtDescuento_global.Text) & "','" & Val(Me.TxtCuotas.Text) & "','" & in_interes & "','" & Val(Me.txtid_venta_ref.Text) & "','" & in_diferida & "','" & Val(Me.txt_id_pendiente.Text) & "','" & KEY_RUC & "')"
            End If
            
            Call ConfiguraRstPP(strCadena)
            id_venta = rstPP("in_venta")
            
           
            
            If Abs(Round(get_total_detalle(id_venta), 2) + Val(Me.lblicbper.Caption) - Round(TotalVenta + Val(Me.TxtDescuento_global.Text), 2)) > 0.2 Then
                
                If MsgBox("Ha ocurrido un Inprevisto con el Envio de comprobante" + Chr(13) + "Se va a proceder a enviar Nuevamente", vbYesNo + vbQuestion, "Importante") = vbYes Then
                    Call eliminar_error
                    GoTo Iniciar
                    Exit Function
                Else
                    Call eliminar_error
                    MsgBox "Este comprobante tiene Inconsistencia." + Chr(13) + "Consulte con el Area de Sistemas �" + Chr(13) + "Genere Nuevamente.", vbInformation
                    Exit Function
                End If
                
            End If
            
            
            If KEY_RUC = "20561358550" Then
                strCadena = "UPDATE movimiento_venta SET percepcion='" & Val(Me.lblPercepcion.Caption) & "' WHERE id_venta='" & id_venta & "' LIMIT 1"
                CnBd.Execute (strCadena)
            End If
            
            
            '*******
            If KEY_RUC = "20487936813" Then
                strCadena = "UPDATE movimiento_venta SET id_referencia='" & Val(Me.txtid_venta_ref.Text) & "' WHERE id_venta='" & id_venta & "' LIMIT 1"
                CnBd.Execute (strCadena)
            End If
            
            
          
            Me.txtidVenta.Text = id_venta
            
            If Val(Me.txtid_agenda.Text) > 0 Then
                strCadena = "CALL procedure_agenda('10','" & Val(Me.txtid_agenda.Text) & "','','','','','','','" & Val(id_venta) & "','','','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                
            End If
            
            If Val(Me.TxtCuotas.Text) > 0 Then
                Call put_letras_cobrar(id_venta)
                If KEY_CONTABILIDAD = "si" Then
                    strCadena = "call P_insert_venta_asiento_contable('" & id_venta & "')"
                    CnBd.Execute (strCadena)
                End If
                
            End If
            
            
            
            
            Call put_correlativo_venta(Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText, Me.TxtNumeroDoc.Text)
            
            
            
          
            If Me.chk_direccion.Value = 1 Then
                in_direccion = Val(Me.hfdireccion.TextMatrix(Me.hfdireccion.Row, 0))
                strCadena = "UPDATE movimiento_venta SET id_direccion='" & in_direccion & "' WHERE id_venta='" & id_venta & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
            End If
            
            
            
            If Me.chk_factura.Value = 1 Then
               Call put_referencia(Me.txtid_venta_ref.Text, Me.txtidVenta.Text)
            End If
            
            If KEY_TRACKING = "si" Then
                Call put_tracking(id_venta, "01", "REGISTRO DE PAGO")
            End If
                         
            If Trim(Me.txtdni_copropietario.Text) <> "" Then
               strCadena = "UPDATE movimiento_venta SET id_copropietario='" & Trim(Me.txtdni_copropietario.Text) & "' WHERE id_venta='" & id_venta & "'"
                CnBd.Execute (strCadena)
            End If
            '---
            If Me.DtcTipoDoc.BoundText = "0007" Then
                    Call update_nota(id_venta, Me.DtcComprobanteGuia.BoundText, Trim(Me.TxtSeri_guia.Text), Trim(Me.TxtNumero_guia.Text))
            End If
            Save = True
        Else
            MsgBox "INGRESE UN MONTO DE PAGO VALIDO", vbInformation, KEY_EMPRESA
            Call Resalta(Me.TxtMontoPagado)
            Exit Function
     End If
        
            
        
        
        in_monto_saldo = 0
        rstK.MoveFirst
        For k = 0 To rstK.RecordCount - 1
            If Trim(Me.txt_tipo_movimiento.Text) = "01" Then
                nmonto_caja = Format(rstK("monto_caja"), "###0.00")
            Else
                If rstK("monto_caja") > 0 Then
                    nmonto_caja = Format(rstK("monto_caja") * -1, "###0.00")
                Else
                    nmonto_caja = Format(rstK("monto_caja"), "###0.00")
                End If
                
            End If
            
            If Val(Me.TxtDescuento_global.Text) > 0 Then
                in_factor_descuento = Val(Me.lblPago.Caption) - Val(Me.TxtDescuento_global.Text)
                If in_factor_descuento > 0 Then
                    in_descuento_parcial = rstK("monto_caja") / (in_factor_descuento) * Val(Me.TxtDescuento_global.Text)
                Else
                    in_descuento_parcial = rstK("monto_caja")
                End If
               
            End If
            
            
            If rstK("forma_pago") = "02" Then
               in_monto_saldo = in_monto_saldo + rstK("monto_caja")
               
               
               If KEY_RUC = "20603698852" Then
                       If Me.DtcTipoDoc.BoundText = "0007" Then
                            Call put_realizar_pago(Val(Me.txtidVenta.Text), Val(Me.txtid_venta_ref.Text), Abs(nmonto_caja), Me.DtcTipoDoc.BoundText, Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
                            Call put_realizar_pago(Val(Me.txtid_venta_ref.Text), Val(Me.txtidVenta.Text), Abs(nmonto_caja), Me.DtcTipoDoc.BoundText, Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
                       End If
               Else
                   If Me.DtcTipoDoc.BoundText = "0007" Then
                       
                       in_saldo_ref = get_saldo_comprobante(Me.txtid_venta_ref.Text)
                       
                       If in_saldo_ref > 0 Then
                          If Abs(nmonto_caja) > in_saldo_ref Then
                             nmonto_caja = in_saldo_ref
                          End If
                          
                          Call put_realizar_pago(Val(Me.txtidVenta.Text), Val(Me.txtid_venta_ref.Text), Abs(nmonto_caja), Me.DtcTipoDoc.BoundText, Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
                          Call put_realizar_pago(Val(Me.txtid_venta_ref.Text), Val(Me.txtidVenta.Text), Abs(nmonto_caja), Me.DtcTipoDoc.BoundText, Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
                        
                       End If
                       
                       
                       
                       'Call put_realizar_pago(Val(Me.txtidVenta.Text), Val(Me.txtid_venta_ref.Text), Abs(nmonto_caja), Me.DtcTipoDoc.BoundText, Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
                       'Call put_realizar_pago(Val(Me.txtid_venta_ref.Text), Val(Me.txtidVenta.Text), Abs(nmonto_caja), Me.DtcTipoDoc.BoundText, Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
                       
                       If Me.DtcFormaPago.BoundText = "01" Then
                            Call put_realizar_pago(Val(Me.txtid_venta_ref.Text), Val(Me.txtidVenta.Text), Abs(nmonto_caja), Me.DtcTipoDoc.BoundText, Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
                       End If
                    End If
               End If
               
            Else
               in_glosa = "PAGO :" & Documento
               If Trim(Me.txtServicio.Text) = "si" Then
                  in_flujo = "1CIX000000000077"
               Else
                  in_flujo = "1CIX000000000078"
               End If
               
               
               
               
               
               If Me.DtcTipoDoc.BoundText <> "0099" Then
                    If KEY_SIN_EFECTO_CAJA = "no" Then
                        
                        in_mis_cuentas_det = procesar_transaccion_venta(rstK("id_forma_pago"), KEY_ALM, get_cuenta_pago(rstK("id_forma_pago")), KEY_FECHA, "00001", Trim(Me.TxtCodCliente.Text), Trim(Me.TxtCliente.Text), in_glosa, nmonto_caja, "0", id_venta, "0", Documento, Val(Me.TxtTipoCambio.Text), rstK("id_tarjeta_operacion"), "1CIX000000000174", in_flujo, KEY_USUARIO, Me.DtcTipoDoc.BoundText, KEY_RUC)
                    End If
                    
                    If Me.DtcTipoDoc.BoundText = "0007" Then
                       
                       in_saldo_ref = get_saldo_comprobante(Me.txtid_venta_ref.Text)
                       
                       If in_saldo_ref > 0 Then
                          Call put_realizar_pago(Val(Me.txtidVenta.Text), Val(Me.txtid_venta_ref.Text), Abs(in_saldo_ref) - Abs(nmonto_caja), Me.DtcTipoDoc.BoundText, Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
                          nmonto_caja = Abs(in_saldo_ref) - Abs(nmonto_caja)
                          
                          If nmonto_caja <> 0 Then
                                Call put_realizar_pago(Val(Me.txtid_venta_ref.Text), Val(Me.txtidVenta.Text), Abs(nmonto_caja), Me.DtcTipoDoc.BoundText, Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
                          End If
                     
                       
                       End If
                        
                       
                       
                    Else
                      
                      If get_forma_pago_detalle(rstK("id_forma_pago")) = "13" Then   ' ****  pagar con nota de credito
                                in_id_nota = get_id_nota(Trim(Me.txtserie_nota.Text), Trim(Me.txtNumero_nota.Text))
                                Call put_realizar_pago(in_id_nota, Val(Me.txtidVenta.Text), Abs(nmonto_caja), "0007", Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
                                Call put_realizar_pago(Val(Me.txtidVenta.Text), in_id_nota, Abs(nmonto_caja), "0007", Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
                                If Me.DtcTipoDoc.BoundText = "0097" Then
                                    strCadena = "call CON_InsertaAsiento_CobroGlobal('" & Val(Me.txtidVenta.Text) & "')"
                                    CnBd.Execute (strCadena)
                                    
                                End If
                       Else
                                If KEY_SIN_EFECTO_CAJA = "no" Then
                                Call put_realizar_pago(Val(Me.txtidVenta.Text), Val(Me.txtidVenta.Text), Abs(nmonto_caja), Me.DtcTipoDoc.BoundText, Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
                                End If
                       End If
                      
                    End If
                    
                    
                    
                   
                    
                    '--
                    If Me.DtcComprobanteGuia.BoundText = "0007" And Val(Me.txtid_venta_ref.Text) > 0 Then
                        strCadena = "update movimiento_venta SET id_referencia='" & Val(Me.txtid_venta_ref.Text) & "' WHERE id_venta='" & Val(Me.txtidVenta.Text) & "'"
                        CnBd.Execute (strCadena)
                    End If
                    '---Finalizar
        
               End If
            End If
            
            'strCadena = "INSERT INTO movimiento_venta_monto(id_venta,forma_pago,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,banco,cheque,cuenta_contable,ruc)VALUES('" & id_venta & "','" & rstK("forma_pago") & "','" & rstK("id_forma_pago") & "','" & rstK("monto") & "','" & nmonto_caja & "','" & rstK("id_tarjeta") & "','" & rstK("id_tarjeta_numero") & "','" & rstK("id_tarjeta_operacion") & "','" & rstK("banco") & "','" & rstK("cheque") & "','" & rstK("cuenta_contable") & "','" & KEY_RUC & "')"
            'CnBd.Execute (strCadena)
            rstK.MoveNext
        Next k
        
        
     '   If get_forma_pago_detalle(Me.DtcForma_pago_detalle.BoundText) = "13" Then
            
      '  End If
        
        
        
        If in_monto_saldo > 0 And Me.DtcTipoDoc.BoundText <> "0099" Then
            Call put_cuenta_cobrar(id_venta, in_monto_saldo)
        End If
                
        'If Val(Me.TxtCuotas.Text) > 0 And Me.DtcFormaPago.BoundText = "02" Then
       '     strCadena = "SELECT * FROM movimiento_venta_cuotas_temporal WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Me.DtcSerieDoc.BoundText & "' AND numero='" & Me.TxtNumeroDoc.Text & "' AND id_usuario='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
        '    Call ConfiguraRstT(strCadena)
         '   If rstT.RecordCount > 0 Then
          '      rstT.MoveFirst
           '     For N = 0 To rstT.RecordCount - 1
            '        strCadena = "INSERT INTO movimiento_venta_cuotas(id_cuota,id_venta,vencimiento,monto,saldo,ruc)VALUES('" & rstT("id_cuota") & "','" & id_venta & "','" & Format(rstT("vencimiento"), "YYYY-mm-dd") & "','" & rstT("monto") & "','" & rstT("saldo") & "','" & KEY_RUC & "')"
             '       CnBd.Execute (strCadena)
              '      rstT.MoveNext
              '  Next N
            'End If
        'End If
        
          
If Me.chk_manual.Value = 1 Then
    GoTo siguientemanual
End If
'StrNumero = FormatosCeros(Trim(str(Val(Me.TxtNumeroDoc.Text)) + 1), 6)
'strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE  id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.DtcSerieDoc.BoundText) & "'  AND ruc='" & Trim(KEY_RUC) & "'"
'CnBd.Execute (strCadena)

siguientemanual:

If KEY_GENERADOR_MENSUALIDAD = "si" Then
        Call put_pagar_servicio_cobranza(id_venta, Val(Me.lblTotal.Caption), KEY_RUC)
End If


If Me.chkseguro.Value = 1 And KEY_PROYECTO = "si" Then
   strCadena = "call sp_update_movimiento_venta_proyecto('" & Val(Me.txtidVenta.Text) & "','" & Val(Me.Dtcseguro.BoundText) & "')"
   CnBd.Execute (strCadena)
End If

If Me.lblfactura_masiva.Caption = "si" Then
    Call put_cancelar_recibo_masivos
Else
    If KEY_RUC = "20566449383" And Val(Me.txtid_venta_ref.Text) > 0 Then
        Call put_cancelar_recibo(Me.txtid_venta_ref.Text)
    End If
End If

Exit Function
        
'error:
'MsgBox "HA OCURRIDO UN PROBLEMA CON LA RED" + Chr(13) + Err.Description, vbInformation, KEY_EMPRESA
  
  
End Function
Private Function get_total_detalle(ByVal in_venta As String) As Double
strCadena = "SELECT sum(total) as in_total_detalle FROM movimiento_venta_detalle  WHERE id_venta='" & in_venta & "'"
Call ConfiguraRstIN(strCadena)
If IsNull(rstIN("in_total_detalle")) = True Then
    get_total_detalle = 0
Else
    get_total_detalle = rstIN("in_total_detalle") + Val(Me.lblPercepcion.Caption)
End If

End Function
Private Sub eliminar_error()

                    Call EliminarVentas_error(Trim(DtcTipoDoc.BoundText), Trim(DtcSerieDoc.BoundText), Trim(TxtNumeroDoc.Text), Trim(DtcAlmacen.BoundText))
                   
                  
          
End Sub
Private Sub put_letras_cobrar(ByVal in_venta As String)
 Dim in_numero As String
 Dim in_letra As String
 Dim in_interes As Single
 
 If Trim(Me.txtServicio.Text) = "no" Then
                in_cta_cobrar = KEY_CTA_COBRAR_PRODUCTO
                in_cta_ingreso = KEY_CTA_INGRESO_PRODUCTO
 Else
                in_cta_cobrar = KEY_CTA_COBRAR_SERVICIO
                in_cta_ingreso = KEY_CTA_INGRESO_SERVICIO
 End If
            
            
 strCadena = "SELECT total FROM movimiento_venta_detalle WHERE id_producto='06225' and  id_venta='" & Val(in_venta) & "' and ruc='" & KEY_RUC & "'"
 Call ConfiguraRstP(strCadena)
 If rstP.RecordCount > 0 Then
    in_interes = rstP("total")
    strCadena = "UPDATE movimiento_venta SET monto_interes='" & in_interes & "' WHERE id_venta='" & in_venta & "'"
    CnBd.Execute (strCadena)
 Else
    in_interes = 0
 End If
 
 
 strCadena = "SELECT id_cuota,monto,saldo,vencimiento FROM movimiento_venta_cuotas_temporal  WHERE id_alm='" & KEY_ALM & "' and  id_usuario='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "' ORDER BY id_cuota ASC"
 Call ConfiguraRstP(strCadena)
 If rstP.RecordCount > 0 Then
    rstP.MoveFirst
    in_numero_credito = get_numero_credito
    For i = 0 To rstP.RecordCount - 1
        in_numero = Format(in_numero_credito, "0000") & Format(i + 1, "00")
        in_documento = "L.COBRAR :" + in_numero
        in_monto_interes = in_interes / rstP.RecordCount
        
        If Val(Me.LblIgv.Caption) > 0 Then
            in_valor_venta = rstP("monto") / (1 + KEY_IGV)
            in_igv = in_valor_venta * (KEY_IGV)
            in_exonerado = 0
        Else
            in_valor_venta = rstP("monto")
            in_igv = 0
            in_exonerado = in_valor_venta
        End If
        
                
        strCadena = "INSERT INTO movimiento_venta (id_doc,id_alm,id_forma_pago,id_moneda,serie,numero,id_cliente,ncliente,valor_venta,igv,exonerado,total,saldo,monto_pago,monto_vuelto,fecha_emision," & _
        "fecha_vencimiento,id_tipo_factura,id_vendedor,dni_save,tc,afecta_factura,id_mes,id_anio,documento,hora,turno,direccion,id_ventanilla,id_tipo,observacion,cta_cobrar,cta_ingreso,id_referencia,monto_interes,interes_revertido,ruc) " & _
        "VALUES('0412','" & KEY_ALM & "','02','" & Me.DtcMoneda.BoundText & "','001','" & in_numero & "','" & Trim(Me.TxtCodCliente.Text) & "','" & Trim(Me.TxtCliente.Text) & "'," & _
        "'" & in_valor_venta & "','" & in_igv & "','" & in_exonerado & "','" & rstP("monto") & "','" & rstP("monto") & "','" & rstP("monto") & "',0,'" & KEY_FECHA & "','" & Format(rstP("vencimiento"), "YYYY-mm-dd") & "','00001'" & _
        ",'" & Me.DtcVendedor.BoundText & "','" & KEY_USUARIO & "','" & Val(Me.TxtTipoCambio.Text) & "','no','" & Format(Month(KEY_FECHA)) & "','" & Year(KEY_FECHA) & "','" & in_documento & "',CURTIME(),'-', " & _
        "'" & Trim(Me.txtDireccion.Text) & "','" & KEY_VENTANILLA & "','01','-','1230101','" & in_cta_ingreso & "','" & in_venta & "','" & in_monto_interes & "','" & in_monto_interes & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        strCadena = "SELECT id_venta FROM movimiento_venta WHERE ruc='" & KEY_RUC & "' and id_doc='0412' ORDER BY id_venta LIMIT 1"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
           in_letra = rstZ("id_venta")
           strCadena = "INSERT INTO movimiento_venta_detalle (id_venta,id_producto,cantidad,valor_neto,igv,precio,total,detalle,ruc)VALUES('" & Val(in_letra) & "','00000',1,'" & in_valor_venta & "','" & in_igv & "','" & rstP("monto") & "','" & rstP("monto") & "','" & in_documento & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           'strCadena = "CALL p_insert_pago_factura_ultimate_ii('" & in_letra & "','" & in_venta & "','" & rstP("monto") & "','" & rstP("monto") & "','" & rst("id_moneda") & "','" & rstP("id_moneda") & "','" & rst("tc") & "')"
           'CnBd.Execute (strCadena)
           
           Call put_realizar_pago(in_letra, in_venta, rstP("monto"), "0412", Val(Me.TxtTipoCambio.Text), Val(in_mis_cuentas_det))
           
           
           
           
 
        strCadena = "select id_forma_pago  from movimiento_venta_monto WHERE id_venta='" & Val(in_venta) & "' and forma_pago='02' LIMIT 1"
        Call ConfiguraRstZ(strCadena)
        If rstZ.RecordCount > 0 Then
            strCadena = "INSERT INTO movimiento_venta_monto(id_venta,id_forma_pago,forma_pago,monto,monto_caja,ruc) VALUES ('" & in_letra & "','" & rstZ("id_forma_pago") & "','02','" & rstP("monto") & "','" & rstP("monto") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
        End If
        
       If KEY_CONTABILIDAD = "si" Then
          strCadena = "call P_insert_venta_agenda_test('" & in_letra & "')"
          CnBd.Execute (strCadena)
       End If
       End If
        
        rstP.MoveNext
    Next i
 
 End If
 


End Sub


Private Sub put_cuenta_cobrar(ByVal in_venta As String, ByVal in_saldo As String)
On Error GoTo sit
If Val(Me.TxtCuotas.Text) < 1 Then
strCadena = "UPDATE movimiento_venta SET saldo='" & in_saldo & "' WHERE id_venta='" & Val(in_venta) & "'"
CnBd.Execute (strCadena)
End If
Exit Sub
sit:
End Sub
Private Sub validar_cliente(ByVal in_dni As String)
On Error GoTo salir
strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & in_dni & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
   strCadena = "INSERT INTO entidad_empresa(cod_unico,id_empresa,id_almacen)VALUES ('" & in_dni & "','" & KEY_RUC & "','" & KEY_ALM & "')"
   CnBd.Execute (strCadena)
End If
Exit Sub
salir:

End Sub
Private Function put_dni_boleta(ByVal in_dni As String, ByVal in_monto As Single, ByVal in_doc As String) As Boolean

If in_doc <> "0003" Then
    put_dni_boleta = True
    Exit Function
    
End If


If in_doc = "0003" And in_monto > 700 And in_dni = "00000000" Then
    If MsgBox("Ingrese un DNI VALIDO el monto supera los 700" + Chr(13) + "Configuracion SUNAT" + Chr(13) + Chr(13) + "Desea Continuar De Todos Modos.", vbInformation + vbYesNo, KEY_VENDEDOR) = vbYes Then
        put_dni_boleta = True
         Exit Function
    Else
        put_dni_boleta = False
        Exit Function
    End If
    
    
End If

If in_doc = "0003" And in_monto <= 90000 Then
     put_dni_boleta = True
    Exit Function
End If

If in_monto >= 90000 And in_doc = "0003" And in_dni <> "00000000" And Len(in_dni) = 8 Then
    put_dni_boleta = True
    Exit Function
Else
    
    MsgBox "Ingrese un DNI VALIDO el monto supera los 700" + Chr(13) + "Configuracion SUNAT", vbInformation, KEY_VENDEDOR
    put_dni_boleta = False
    Exit Function
End If
put_dni_boleta = True


End Function
Private Sub next_save()
                If KEY_COMPROBANTE_ADICIONAL = "no" Then
                    Call nuevo
                    Exit Sub
                End If
                Me.TxtCodProducto.Enabled = False
                Me.TxtDescripcionProducto.Enabled = False
                Me.txtCantidad.Enabled = False
                Me.txtprecio.Enabled = False
                Me.cmdAgregar.Enabled = False
                Me.CmdQuitar.Enabled = False
                chkPrecios.Enabled = False
                Me.cmdAnular.Enabled = True
               ' Me.cmdEditable.Enabled = False
                If (KEY_CARGO = "00001" Or KEY_CARGO = "00004" And KEY_FACTURACION_ELECTRONICA = "no") Then
                    Me.cmdEliminar.Enabled = True
                Else
                    
                    Me.cmdEliminar.Enabled = False
                End If
                Me.cmdProcesar.Enabled = False
                Me.cmdImprimir.Enabled = True
                dfactura = False
                Me.chk_factura.Value = 0
End Sub
Private Sub put_referencia(ByVal in_referencia As String, ByVal in_venta As String)
strCadena = "UPDATE movimiento_venta SET id_recibo='" & Val(in_referencia) & "' WHERE id_venta='" & Val(in_venta) & "'"
CnBd.Execute (strCadena)
End Sub
Private Sub ActualizarAdelanto(ByVal TotalPedido As Double)
Dim MontoAnterior As Double
strCadena = "SELECT MontoAdelantado FROM Persona WHERE cPersona='" & Trim(Me.TxtCodCliente.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    MontoAnterior = rst(0)
    Set rst = Nothing
End If
strCadena = "UPDATE Persona SET MontoAdelantado='" & (MontoAnterior - TotalPedido) & "' WHERE cPersona='" & Trim(Me.TxtCodCliente.Text) & "'"
Call EjecutaRST(strCadena)
Set RstEjecuta = Nothing
End Sub
Private Sub save_especial()
            Dim vuelto As Double, pago As Double, saldo1 As Double, saldo2 As Double
            dfac = "no"
            
            strCadena = "UPDATE temporal_ventas SET numero='" & formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6) & "' WHERE igv='no' AND dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
            CnBd.Execute (strCadena)
            '
            
             
            strCadena = "SELECT sum(total) FROM temporal_ventas WHERE igv='si' AND dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND numero='" & Trim(Me.TxtNumeroDoc.Text) & "'"
            Call ConfiguraRstT(strCadena)
            
            strCadena = "SELECT sum(total) FROM temporal_ventas WHERE igv='no' AND dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
            Call ConfiguraRst(strCadena)
            vuelto = Val(Me.lblPago.Caption) - (Val(rstT(0)) + Val(rst(0)))
            pago = Val(Me.lblPago.Caption) - Val(rst(0))
            If Me.DtcFormaPago.BoundText = "05" Then
                KEY_VENCIMIENTO = Format(Me.DtpFechaReferencia.Value, "yyyy-mm-dd")
            Else
                If Me.DtcFormaPago.BoundText = "01" Then
                    saldo1 = 0#
                    saldo2 = 0#
                Else
                    saldo1 = rstT(0)
                    saldo2 = rstTemporal(0)
                End If
            KEY_VENCIMIENTO = KEY_FECHA
        End If
        
            'CON IGV----
            strCadena = "P_insert_venta('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcFormaPago.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
            "'" & Trim(Me.DtcSerieDoc.BoundText) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Me.TxtCodCliente.Text & "','" & Me.TxtCliente.Text & "','" & Val(Me.LblValorVenta.Caption) & "','" & Val(Me.LblIgv.Caption) & "','0','" & rstT(0) & "','" & saldo1 & "'," & _
            "'" & Val(pago) & "','" & Val(vuelto) & "','" & KEY_FECHA & "','" & KEY_VENCIMIENTO & "','00001','" & KEY_USUARIO & "','" & Val(Me.TxtTipoCambio.Text) & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            '
             
           
            id_venta = LastRegistroRUC("movimiento_venta", "id_venta")
            Call SaveDetalleDocumentoVenta(id_venta, Trim(Me.txteditable.Text), Me.DtcTipoDoc.BoundText, Me.DtcSerieDoc.BoundText)
            
                      
            strCadena = "P_insert_venta('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcFormaPago.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
            "'" & Trim(Me.DtcSerieDoc.BoundText) & "','" & formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6) & "','" & Me.TxtCodCliente.Text & "','" & Me.TxtCliente.Text & "','0','0','" & Val(Me.lblExonerado.Caption) & "','" & rst(0) & "','" & saldo2 & "'," & _
            "'" & rst(0) & "','0','" & KEY_FECHA & "','" & KEY_VENCIMIENTO & "','00001','" & Me.DtcVendedor.BoundText & "','" & KEY_USUARIO & "','" & Val(Me.TxtTipoCambio.Text) & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
            id_venta = LastRegistroRUC("movimiento_venta", "id_venta")
            '
            Call SaveDetalleDocumentoVentaEspecial(id_venta, formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6))
            
            
StrNumero = FormatosCeros(Trim(str(Val(Me.TxtNumeroDoc.Text)) + 2), 6)
strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.DtcSerieDoc.BoundText) & "'  AND ruc='" & Trim(KEY_RUC) & "'"
CnBd.Execute (strCadena)
'
                Me.TxtCodProducto.Enabled = False
                Me.TxtDescripcionProducto.Enabled = False
                Me.txtCantidad.Enabled = False
                Me.txtprecio.Enabled = False
                Me.cmdAgregar.Enabled = False
                Me.CmdQuitar.Enabled = False
                Me.cmdAnular.Enabled = True
                If (KEY_CARGO = "00001" Or KEY_CARGO = "00004") Then
                    
                    Me.cmdEliminar.Enabled = True
                Else
                    
                    Me.cmdEliminar.Enabled = False
                End If
                
                Me.cmdProcesar.Enabled = False
                
                Me.cmdImprimir.Enabled = True
                dfactura = False
                Me.chk_factura.Value = 0
                Exit Sub
            
End Sub

Private Function SaveDetalleDocumentoVenta(ByVal idVenta As Double, ByVal in_editable As String, ByVal in_doc As String, ByVal in_serie As String) As Boolean
Dim in_tipo_factura As String
Dim in_codigo As String
On Error GoTo Saltar
    SaveDetalleDocumentoVenta = False
    If in_editable = "si" Then
                
    Else
           strCadena = "SELECT * FROM view_venta_temporal_ii WHERE id_doc='" & in_doc & "' and id_serie='" & in_serie & "' and dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and  ruc='" & KEY_RUC & "'"
           Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
               rstT.MoveFirst
               For i = 0 To rstT.RecordCount - 1
                        strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,cantidad,precio,peso,total,serie,anio_fabricacion,nro_chasis,anio_modelo,nro_dua,nro_item,id_detalle_serie,detalle,ruc) VALUES " & _
                        "('" & idVenta & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("precio") & "','" & rstT("peso") & "','" & rstT("total") & "','" & rstT("serie") & "','" & rstT("anio_fabricacion") & "','" & rstT("nro_chasis") & "','" & rstT("anio_modelo") & "','" & rstT("nro_dua") & "','" & rstT("nro_item") & "','" & rstT("id_detalle_serie") & "','" & rstT("detalle") & "','" & KEY_RUC & "')"
                        CnBd.Execute (strCadena)
                     
                        If rstT("nro_chasis") <> "-" Then
                            strCadena = "UPDATE imp_producto_detalle SET vendido='si' WHERE nro_chasis='" & rstT("nro_chasis") & "' and ruc='" & KEY_RUC & "'"
                            CnBd.Execute (strCadena)
                            
                            strCadena = "INSERT INTO imp_tramite(`id_venta`,`ruc`)VALUES('" & idVenta & "','" & KEY_RUC & "')"
                            CnBd.Execute (strCadena)
                            
                         End If
                         SaveDetalleDocumentoVenta = True
                   rstT.MoveNext
                Next i
            End If
    End If
    Exit Function
Saltar:
SaveDetalleDocumentoVenta = False
End Function

Private Sub SaveDetalleDocumentoVentaRecibo(ByVal idVenta As Double)

    strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & Val(Me.txtidVenta.Text) & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       
       rstT.MoveFirst
       For i = 0 To rstT.RecordCount - 1
           strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,detalle,id_producto,cantidad,precio,peso,total,serie,anio_fabricacion,nro_chasis,anio_modelo,nro_dua,nro_item,ruc) VALUES ('" & idVenta & "','" & rstT("detalle") & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("precio") & "','" & rstT("peso") & "','" & rstT("total") & "','" & rstT("serie") & "','" & rstT("anio_fabricacion") & "','" & rstT("nro_chasis") & "','" & rstT("anio_modelo") & "','" & rstT("nro_dua") & "','" & rstT("nro_item") & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           rstT.MoveNext
       Next i
    
    End If
End Sub

Private Sub SaveDetalleDocumentoVentaEspecial(ByVal idVenta As Double, ByVal numero As String)

   strCadena = "SELECT * FROM temporal_ventas WHERE (numero='" & Trim(numero) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_serie='" & Trim(Me.DtcSerieDoc.BoundText) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "')"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       For i = 0 To rstT.RecordCount - 1
           strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,cantidad,precio,peso,total,ruc) VALUES ('" & idVenta & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("precio") & "','" & rstT("peso") & "','" & rstT("total") & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
            
            
           rstT.MoveNext
        Next i
    End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtCodProducto)
End If
End Sub



Private Sub TxtSeri_guia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSeri_guia.Text = UCase(Me.TxtSeri_guia.Text)
    Me.TxtNumero_guia.SetFocus
End If
End Sub

Private Sub TxtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Me.DtcTipoDoc.SetFocus
End If
If KeyCode = vbKeyRight Then
        Call Resalta(Me.TxtNumeroDoc)
End If
End Sub


Private Function ConvertirMes(ByVal numero As Integer) As String
Select Case numero
    Case 1
        ConvertirMes = "ENERO"
    Case 2
        ConvertirMes = "FEBRERO"
    Case 3
        ConvertirMes = "MARZO"
    Case 4
        ConvertirMes = "ABRIL"
    Case 5
        ConvertirMes = "MAYO"
    Case 6
        ConvertirMes = "JUNIO"
    Case 7
        ConvertirMes = "JULIO"
    Case 8
        ConvertirMes = "AGOSTO"
    Case 9
        ConvertirMes = "SETIEMBRE"
    Case 10
        ConvertirMes = "OCTUBRE"
    Case 11
        ConvertirMes = "MOVIEMBRE"
    Case 12
        ConvertirMes = "DICIEMBRE"
    
End Select
End Function


Private Sub txtVueltoDelivery_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtCodProducto)
End If
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Me.DtcSerieDoc.BoundText) <> "" Then
        Me.DtcSerieDoc.BoundText = formato_item(Me.DtcSerieDoc.BoundText, 3)
        strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.DtcSerieDoc.BoundText) & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            Me.TxtNumeroDoc.Text = rst("numero")
            Me.DtcMoneda.BoundText = rst("id_moneda")
            KEY_APLICA_IGV = rst("igv")
            Me.TxtCodProducto.Enabled = True
            Call Resalta(Me.TxtCodProducto)
        Else
            MsgBox "SERIE NO ASIGNADA A ESTA SUCURSAL", vbInformation, KEY_EMPRESA
            Call Resalta(Me.TxtNumeroDoc)
            Exit Sub
        End If
    End If
End If
End Sub



Private Sub txtSerie_nota_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtserie_nota.Text = UCase(Me.txtserie_nota.Text)
    Call Resalta(Me.txtNumero_nota)
End If
End Sub


Private Sub load_seguro(ByVal in_dni As String)

strCadena = "SELECT id_detalle as Codigo, CONCAT(descripcion,' [ ',descuento,'%] ') as Descripcion FROM persona_seguro WHERE  activo='si' and dni='" & in_dni & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.Dtcseguro)
If Me.Dtcseguro.Enabled = True Then
   Me.chkseguro.Value = 1
End If
End Sub

Private Sub load_proyecto(ByVal in_cliente As String)
strCadena = "SELECT id_proyecto as Codigo,descripcion as Descripcion FROM mis_proyectos WHERE id_cliente='" & in_cliente & "' and  finalizado='no' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.Dtcseguro)

   Me.chkseguro.Value = 1
   

End Sub


Public Function validar_comprobante_electronico(ByVal in_doc As String, ByVal in_dni As String) As Boolean
 
 validar_comprobante_electronico = True
 
 If Len(Trim(Me.txtDireccion.Text)) < 2 Then
    MsgBox "DISCULPE SR(A): " & KEY_VENDEDOR + Chr(13) + Chr(13) + "INGRESE UNA DIRECCION VALIDA.", vbInformation, KEY_EMPRESA
                Call Resalta(Me.txtDireccion)
                validar_comprobante_electronico = False
 End If
 
 
 If verifica_existencia_persona(Trim(in_dni)) = False Then
                MsgBox "DISCULPE SR(A): " & KEY_VENDEDOR + Chr(13) + Chr(13) + "DNI/RUC no esta REGISTRADO.", vbInformation, KEY_EMPRESA
                Call Resalta(Me.TxtCodCliente)
                validar_comprobante_electronico = False
                Exit Function
 End If
    
   
    
Select Case in_doc
    Case "0003"
        
        
         If Me.chk_descuento.Value = 1 Then
            If Val(Me.TxtDescuento_porcentaje.Text) = 100 And Me.HfdDetalle.Rows > 0 Then
                GoTo avanzar2
            End If
         End If
         
         If KEY_APLICA_IGV = "si" And Val(Me.LblIgv.Caption) = 0 And Val(Me.TxtDescuento_porcentaje.Text) < 100 And Val(Me.lblTotal.Caption) > 0 Then
            MsgBox "Configuracion INCORRECTA DEL IGV", vbExclamation, KEY_VENDEDOR
            validar_comprobante_electronico = False
            Exit Function
         End If
         
         
         
         If KEY_APLICA_IGV = "no" And Val(Me.LblIgv.Caption) > 0 Then
            MsgBox "Configuracion INCORRECTA DEL IGV" + Chr(13) + "ZONA SELVA.", vbExclamation, KEY_VENDEDOR
            validar_comprobante_electronico = False
            Exit Function
         End If
         
avanzar2:
         If Len(Trim(Me.TxtCodCliente.Text)) <> 8 Then
            If Trim(Me.txtExtranjero.Text) = "no" Then
                MsgBox "Nota:" + Chr(13) + "Una BOLETA necesita un DNI CORRECTO.", vbExclamation
                validar_comprobante_electronico = False
                Exit Function
           
            End If
         End If
         
         If Val(Me.lblCantidad.Caption) < 1 Then
            MsgBox "Nota:" + Chr(13) + "DEBE TENER UN ITEM como m�nimo.", vbExclamation
            validar_comprobante_electronico = False
            Exit Function
         End If
         
    Case "0001"
        If KEY_APLICA_IGV = "si" And Val(Me.LblIgv.Caption) = 0 And Val(Me.TxtDescuento_porcentaje.Text) <> 100 Then
            MsgBox "Configuracion INCORRECTA DEL IGV", vbExclamation, KEY_VENDEDOR
            validar_comprobante_electronico = False
            Exit Function
         End If
         
         If KEY_APLICA_IGV = "no" And Val(Me.LblIgv.Caption) > 0 Then
            MsgBox "Configuracion INCORRECTA DEL IGV" + Chr(13) + "ZONA SELVA.", vbExclamation, KEY_VENDEDOR
            validar_comprobante_electronico = False
            Exit Function
         End If
         
         
         If Len(Trim(Me.TxtCodCliente.Text)) <> 11 Then
            If KEY_PAIS = "9589" Then
            If get_extranjero(Trim(Me.TxtCodCliente.Text)) = "no" Then
                MsgBox "Nota:" + Chr(13) + "Una FACTURA necesita un RUC CORRECTO.", vbExclamation
                validar_comprobante_electronico = False
            Else
                validar_comprobante_electronico = True
            End If
            Exit Function
            End If
         End If
         
         If Val(Me.lblCantidad.Caption) < 1 Then
            MsgBox "Nota:" + Chr(13) + "DEBE TENER UN ITEM como m�nimo.", vbExclamation
            validar_comprobante_electronico = False
            Exit Function
         End If
         
    Case "0007"
        
        If Trim(Me.txtmotivo_nota.Text) = "" Then
            MsgBox "Ingrese un MOTIVO para la NOTA CREDITO", vbExclamation, KEY_VENDEDOR
            Me.frm_motivo_nota.Visible = True
            validar_comprobante_electronico = False
            Exit Function
        End If
        
     '   If KEY_APLICA_IGV = "si" And Val(Me.LblIgv.Caption) = 0 Then
     '       MsgBox "Configuracion INCORRECTA DEL IGV", vbExclamation, KEY_VENDEDOR
     '       validar_comprobante_electronico = False
      '      Exit Function
      '  End If
         
        If KEY_APLICA_IGV = "no" And Val(Me.LblIgv.Caption) > 0 Then
            MsgBox "Configuracion INCORRECTA DEL IGV" + Chr(13) + "ZONA SELVA.", vbExclamation, KEY_VENDEDOR
            validar_comprobante_electronico = False
            Exit Function
        End If
        
         
         If Val(Me.lblCantidad.Caption) < 1 Then
            MsgBox "Nota:" + Chr(13) + "DEBE TENER UN ITEM como m�nimo.", vbExclamation
            validar_comprobante_electronico = False
            Exit Function
            
         End If
        
         
         If get_electronico(Me.DtcComprobanteGuia.BoundText, Trim(TxtSeri_guia.Text)) = "si" Then
            
           If Me.DtcComprobanteGuia.BoundText = "0001" Then
              If Mid(Trim(Me.TxtSeri_guia.Text), 1, 1) <> Mid(Trim(Me.DtcSerieDoc.BoundText), 1, 1) Then
                 MsgBox "ADVERTENCIA." + Chr(13) + "SERIE no corresponde a una FACTURA" + Chr(13) + "", vbExclamation, KEY_VENDEDOR
                 validar_comprobante_electronico = False
                 Exit Function
              End If
           End If
           
         If Me.DtcComprobanteGuia.BoundText = "0003" Then
              If Mid(Trim(Me.TxtSeri_guia.Text), 1, 1) <> Mid(Trim(Me.DtcSerieDoc.BoundText), 1, 1) Then
                 MsgBox "ADVERTENCIA." + Chr(13) + "SERIE no corresponde a una BOLETA", vbExclamation, KEY_VENDEDOR
                 validar_comprobante_electronico = False
                 Exit Function
              End If
         End If
            
        End If
        
               
         
    
    
End Select
End Function




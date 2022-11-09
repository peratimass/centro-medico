VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmCorProcesos 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   18990
   ShowInTaskbar   =   0   'False
   Begin VB.Frame framehistorial 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   8400
      TabIndex        =   185
      Top             =   120
      Visible         =   0   'False
      Width           =   10575
      Begin VitekeySoft.ChameleonBtn cmdcerrarhistorial 
         Height          =   390
         Left            =   8160
         TabIndex        =   186
         Top             =   4440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   688
         BTYPE           =   3
         TX              =   "CERRAR HISTORIAL"
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
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfHistorial 
         Height          =   2055
         Left            =   2160
         TabIndex        =   210
         Top             =   2280
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   3625
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
      Begin VitekeySoft.ChameleonBtn cmdMarcarVendido 
         Height          =   390
         Left            =   2160
         TabIndex        =   211
         Top             =   4440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   688
         BTYPE           =   3
         TX              =   "MARCAR VENDIDO:"
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
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblid_detalle_serie 
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
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
         Height          =   255
         Left            =   840
         TabIndex        =   212
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label lblalmacen 
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
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
         Height          =   300
         Left            =   2160
         TabIndex        =   194
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label lblfecha_transfer 
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
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
         Height          =   300
         Left            =   2160
         TabIndex        =   193
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   240
         Picture         =   "frmCorProcesos.frx":0038
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   240
         Picture         =   "frmCorProcesos.frx":05C2
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   240
         Picture         =   "frmCorProcesos.frx":0B4C
         Top             =   240
         Width           =   240
      End
      Begin VB.Label lbldoc_transfer 
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
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
         Height          =   300
         Left            =   2160
         TabIndex        =   192
         Top             =   1800
         Width           =   4695
      End
      Begin VB.Label lbldoc_ingreso 
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
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
         Height          =   300
         Left            =   2160
         TabIndex        =   191
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label lblfecha_ingreso 
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
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
         Height          =   255
         Left            =   2160
         TabIndex        =   190
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label61 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "VENTA  :"
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
         Height          =   225
         Left            =   600
         TabIndex        =   189
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label60 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSFERENCIA"
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
         Height          =   225
         Left            =   600
         TabIndex        =   188
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label59 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "INGRESO  :"
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
         Left            =   600
         TabIndex        =   187
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frmconvertir 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   11640
      TabIndex        =   198
      Top             =   5280
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox txtProducto 
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
         Left            =   840
         TabIndex        =   202
         Top             =   840
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo DtcProducto 
         Height          =   315
         Left            =   120
         TabIndex        =   199
         Top             =   360
         Width           =   5535
         _ExtentX        =   9763
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
      Begin VitekeySoft.ChameleonBtn cmdProcesar 
         Height          =   300
         Left            =   3600
         TabIndex        =   201
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "REALIZAR CONVERSION"
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
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":10D6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label64 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BUSCAR :"
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
         Left            =   135
         TabIndex        =   209
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label63 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCTO "
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
         TabIndex        =   200
         Top             =   120
         Width           =   915
      End
   End
   Begin VB.Frame FrameTercero 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   4140
      Left            =   3960
      TabIndex        =   172
      Top             =   1200
      Visible         =   0   'False
      Width           =   14895
      Begin VB.TextBox txtBuscarInsumo 
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
         Left            =   960
         TabIndex        =   183
         Top             =   3750
         Width           =   5295
      End
      Begin MSDataListLib.DataCombo DtcEmpresaTercero 
         Height          =   315
         Left            =   1200
         TabIndex        =   173
         Top             =   240
         Width           =   5100
         _ExtentX        =   8996
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HgTrabajadorOtro 
         Height          =   1095
         Left            =   120
         TabIndex        =   174
         Top             =   600
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   1931
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
      Begin VitekeySoft.ChameleonBtn cmdCerrarOtro 
         Height          =   285
         Left            =   14520
         TabIndex        =   175
         Top             =   150
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":10F2
         PICN            =   "frmCorProcesos.frx":110E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfInsumoTercero 
         Height          =   1815
         Left            =   120
         TabIndex        =   176
         Top             =   1800
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   3201
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
      Begin VitekeySoft.ChameleonBtn cmdquitarTercero 
         Height          =   285
         Left            =   14520
         TabIndex        =   177
         Top             =   600
         Visible         =   0   'False
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":16A8
         PICN            =   "frmCorProcesos.frx":16C4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfListadoInsumos 
         Height          =   3015
         Left            =   6360
         TabIndex        =   180
         Top             =   600
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   5318
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
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BUSCAR :"
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
         Left            =   135
         TabIndex        =   182
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LISTADO DE INSUMOS"
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
         Left            =   6765
         TabIndex        =   181
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label Label57 
         Caption         =   "-"
         Height          =   255
         Left            =   720
         TabIndex        =   179
         Top             =   2400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label58 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPRESA :"
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
         Left            =   120
         TabIndex        =   178
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox txtreiniciar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10320
      TabIndex        =   170
      Text            =   "no"
      Top             =   5160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Txtid_estado 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10320
      TabIndex        =   154
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame fraTapFin 
      BackColor       =   &H00FFFFFF&
      Height          =   1480
      Left            =   13200
      TabIndex        =   146
      Top             =   7320
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txtObservacionTap2 
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
         Height          =   525
         Left            =   1320
         TabIndex        =   147
         Top             =   240
         Width           =   3615
      End
      Begin VitekeySoft.ChameleonBtn cmdConfFinTap 
         Height          =   375
         Left            =   3120
         TabIndex        =   148
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "CONFIRMAR"
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
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":1C5E
         PICN            =   "frmCorProcesos.frx":1C7A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdCerrarTap2 
         Height          =   285
         Left            =   5160
         TabIndex        =   149
         Top             =   240
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":2214
         PICN            =   "frmCorProcesos.frx":2230
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpFinTapizado 
         Height          =   345
         Left            =   1320
         TabIndex        =   207
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
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
         CalendarBackColor=   16777215
         Format          =   173408257
         CurrentDate     =   42314
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMENTARIO :"
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
         Left            =   120
         TabIndex        =   150
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Frame fraTap 
      BackColor       =   &H00FFFFFF&
      Height          =   3300
      Left            =   13200
      TabIndex        =   126
      Top             =   5520
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txtObservacionTap 
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
         Height          =   375
         Left            =   720
         TabIndex        =   127
         Top             =   2760
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo cboEmpresaTap 
         Height          =   315
         Left            =   1200
         TabIndex        =   128
         Top             =   240
         Width           =   4020
         _ExtentX        =   7091
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridTrabajadorTap 
         Height          =   1095
         Left            =   120
         TabIndex        =   129
         Top             =   600
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   1931
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
      Begin VitekeySoft.ChameleonBtn cmdConfirmarTap 
         Height          =   375
         Left            =   4080
         TabIndex        =   130
         Top             =   2760
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "CONFIRMAR"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":27CA
         PICN            =   "frmCorProcesos.frx":27E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdCerrarTap 
         Height          =   285
         Left            =   5280
         TabIndex        =   131
         Top             =   150
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":2D80
         PICN            =   "frmCorProcesos.frx":2D9C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfTapizadoInsumo 
         Height          =   855
         Left            =   120
         TabIndex        =   155
         Top             =   1800
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   1508
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
      Begin VitekeySoft.ChameleonBtn cmdagregartapiz 
         Height          =   285
         Left            =   5280
         TabIndex        =   156
         Top             =   1800
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":3336
         PICN            =   "frmCorProcesos.frx":3352
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdquitartapiz 
         Height          =   285
         Left            =   5280
         TabIndex        =   157
         Top             =   2160
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":38EC
         PICN            =   "frmCorProcesos.frx":3908
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpInicioTapizado 
         Height          =   345
         Left            =   2640
         TabIndex        =   208
         Top             =   2760
         Width           =   1455
         _ExtentX        =   2566
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
         CalendarBackColor=   16777215
         Format          =   173408257
         CurrentDate     =   42314
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBS :"
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
         Left            =   135
         TabIndex        =   135
         Top             =   2880
         Width           =   435
      End
      Begin VB.Label lblNombreTempTap 
         Caption         =   "-"
         Height          =   255
         Left            =   3000
         TabIndex        =   134
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblDniTempTap 
         Caption         =   "-"
         Height          =   255
         Left            =   720
         TabIndex        =   133
         Top             =   2400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPRESA :"
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
         Left            =   120
         TabIndex        =   132
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CheckBox chkEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ESTADO  :"
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
      Left            =   9600
      TabIndex        =   125
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtModelo 
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
      Left            =   7800
      TabIndex        =   123
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame fraEnsamblaje 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   4155
      Left            =   -9855
      TabIndex        =   82
      Top             =   0
      Visible         =   0   'False
      Width           =   9855
      Begin VB.Frame fraEns2Fin 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   960
         TabIndex        =   93
         Top             =   2520
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton cmdCerrarEns2_2 
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   96
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton cmdConfFinEns2 
            Caption         =   "CONFIRMAR"
            Height          =   375
            Left            =   1440
            TabIndex        =   95
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtObservacionEns2_2 
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
            TabIndex        =   94
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COMENTARIO :"
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
            Left            =   0
            TabIndex        =   97
            Top             =   120
            Width           =   1185
         End
      End
      Begin VB.Frame fraEns2 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   960
         TabIndex        =   85
         Top             =   600
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton cmdCerrarEns_2 
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   88
            Top             =   0
            Width           =   255
         End
         Begin VB.TextBox txtObservacionEns_2 
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
            Left            =   1320
            TabIndex        =   87
            Top             =   1920
            Width           =   3255
         End
         Begin VB.CommandButton cmdConfirmarEns2 
            Caption         =   "CONFIRMAR"
            Height          =   375
            Left            =   1440
            TabIndex        =   86
            Top             =   2280
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo cboEmpresaEns2 
            Height          =   315
            Left            =   960
            TabIndex        =   120
            Top             =   120
            Width           =   3420
            _ExtentX        =   6033
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridTrabajadorEns2 
            Height          =   1215
            Left            =   240
            TabIndex        =   121
            Top             =   600
            Width           =   4365
            _ExtentX        =   7699
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
         Begin VB.Label Label45 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EMPRESA :"
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
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label46 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COMENTARIO :"
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
            Left            =   120
            TabIndex        =   91
            Top             =   1920
            Width           =   1185
         End
         Begin VB.Label lblNombreTempEns2 
            Caption         =   "-"
            Height          =   255
            Left            =   3000
            TabIndex        =   90
            Top             =   2280
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblDniTempEns2 
            Caption         =   "-"
            Height          =   255
            Left            =   600
            TabIndex        =   89
            Top             =   2400
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdFinEns2 
         Caption         =   "FINALIZAR"
         Height          =   375
         Left            =   3840
         TabIndex        =   84
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton cmdIniEns2 
         Caption         =   "INICIAR"
         Height          =   375
         Left            =   1680
         TabIndex        =   83
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INICIO"
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
         Left            =   1680
         TabIndex        =   114
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPRESA :"
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
         Left            =   1800
         TabIndex        =   113
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA :"
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
         Left            =   1320
         TabIndex        =   112
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label lblEns2FechaIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2415
         TabIndex        =   111
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA :"
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
         Left            =   1320
         TabIndex        =   110
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label lblEns2HoraIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2415
         TabIndex        =   109
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label lblEns2Empresa 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3120
         TabIndex        =   108
         Top             =   1920
         Width           =   1275
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRABAJADOR :"
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
         Left            =   1800
         TabIndex        =   107
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Label lblEns2Trabajador 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3120
         TabIndex        =   106
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Left            =   1800
         TabIndex        =   105
         Top             =   2400
         Width           =   1245
      End
      Begin VB.Label lblEns2Observacion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3120
         TabIndex        =   104
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label Label53 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FIN"
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
         Left            =   3870
         TabIndex        =   103
         Top             =   840
         Width           =   285
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA :"
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
         Left            =   3600
         TabIndex        =   102
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label lblEns2FechaFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   4455
         TabIndex        =   101
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label55 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA :"
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
         Left            =   3600
         TabIndex        =   100
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label lblEns2HoraFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   4455
         TabIndex        =   99
         Top             =   1560
         Width           =   675
      End
      Begin VB.Label lblEnsamblaje 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENSAMBLAJE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   2040
         TabIndex        =   98
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame fraCambiar 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3600
      TabIndex        =   78
      Top             =   4680
      Visible         =   0   'False
      Width           =   7335
      Begin MSDataListLib.DataCombo cboCambiar 
         Height          =   315
         Left            =   2640
         TabIndex        =   81
         Top             =   190
         Width           =   3495
         _ExtentX        =   6165
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
      Begin VitekeySoft.ChameleonBtn cmdOkCambiar 
         Height          =   375
         Left            =   6240
         TabIndex        =   118
         Top             =   190
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":3EA2
         PICN            =   "frmCorProcesos.frx":3EBE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label56 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPEZAR CON EL PROCESO : "
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
         Left            =   165
         TabIndex        =   79
         Top             =   240
         Width           =   2265
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdSalir 
      Height          =   375
      Left            =   18480
      TabIndex        =   69
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCorProcesos.frx":4458
      PICN            =   "frmCorProcesos.frx":4474
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtChasis 
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
      Left            =   5040
      TabIndex        =   66
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtMotor 
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
      Left            =   5040
      TabIndex        =   65
      Top             =   120
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridDetalle 
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   820
      Width           =   18645
      _ExtentX        =   32888
      _ExtentY        =   6800
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
   Begin MSDataListLib.DataCombo cboEstado 
      Height          =   285
      Left            =   10800
      TabIndex        =   64
      Top             =   240
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      ListField       =   "c"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VitekeySoft.ChameleonBtn cmdCambiar 
      Height          =   495
      Left            =   240
      TabIndex        =   117
      Top             =   4800
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "CAMBIAR LINEA DE PROCESOS"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCorProcesos.frx":7328
      PICN            =   "frmCorProcesos.frx":7344
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdTerminar 
      Height          =   375
      Left            =   16440
      TabIndex        =   119
      Top             =   4800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "TERMINAR PRODUCTO"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCorProcesos.frx":78DE
      PICN            =   "frmCorProcesos.frx":78FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdreportes 
      Height          =   400
      Left            =   13440
      TabIndex        =   122
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   714
      BTYPE           =   5
      TX              =   "REPORTE DE PRODUCCION"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCorProcesos.frx":7E94
      PICN            =   "frmCorProcesos.frx":7EB0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraEnsamblado 
      BackColor       =   &H00FFFFFF&
      Height          =   3300
      Left            =   240
      TabIndex        =   136
      Top             =   5520
      Visible         =   0   'False
      Width           =   6375
      Begin MSComCtl2.DTPicker DtpFechaInicio 
         Height          =   345
         Left            =   3600
         TabIndex        =   203
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
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
         CalendarBackColor=   16777215
         Format          =   173408257
         CurrentDate     =   42314
      End
      Begin VB.TextBox txtObservacion 
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
         Height          =   375
         Left            =   960
         TabIndex        =   137
         Top             =   2880
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo cboEmpresa 
         Height          =   315
         Left            =   1800
         TabIndex        =   138
         Top             =   240
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   8388608
         ListField       =   "c"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridTrabajador 
         Height          =   1215
         Left            =   480
         TabIndex        =   139
         Top             =   600
         Width           =   5445
         _ExtentX        =   9604
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
      Begin VitekeySoft.ChameleonBtn cmdConfirmar 
         Height          =   340
         Left            =   5040
         TabIndex        =   140
         Top             =   2880
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BTYPE           =   5
         TX              =   "CONFIRMAR"
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
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":7F3D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdCerrarEns 
         Height          =   285
         Left            =   6045
         TabIndex        =   141
         Top             =   150
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":7F59
         PICN            =   "frmCorProcesos.frx":7F75
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfSoldaduraInsumo 
         Height          =   855
         Left            =   480
         TabIndex        =   158
         Top             =   1920
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   1508
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
      Begin VitekeySoft.ChameleonBtn cmdagregarSoldadura 
         Height          =   285
         Left            =   6000
         TabIndex        =   159
         Top             =   1920
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":850F
         PICN            =   "frmCorProcesos.frx":852B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdquitarSoldadura 
         Height          =   285
         Left            =   6000
         TabIndex        =   160
         Top             =   2280
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":8AC5
         PICN            =   "frmCorProcesos.frx":8AE1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBS :"
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
         Left            =   495
         TabIndex        =   145
         Top             =   3000
         Width           =   435
      End
      Begin VB.Label lblNombreTemp 
         Caption         =   "-"
         Height          =   255
         Left            =   3240
         TabIndex        =   144
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblDniTemp 
         Caption         =   "-"
         Height          =   255
         Left            =   720
         TabIndex        =   143
         Top             =   2760
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPRESA :"
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
         Left            =   660
         TabIndex        =   142
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraDetalle 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   4035
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   18800
      Begin VB.Frame fraEnsambladoFin 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   240
         TabIndex        =   161
         Top             =   2160
         Visible         =   0   'False
         Width           =   6150
         Begin VB.TextBox txtObservacion2 
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
            Height          =   405
            Left            =   1320
            TabIndex        =   162
            Top             =   240
            Width           =   4215
         End
         Begin VitekeySoft.ChameleonBtn cmdConfFinEns 
            Height          =   375
            Left            =   3720
            TabIndex        =   163
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "CONFIRMAR"
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
            BCOL            =   8421631
            BCOLO           =   8421631
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCorProcesos.frx":907B
            PICN            =   "frmCorProcesos.frx":9097
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdCerrarEns2 
            Height          =   285
            Left            =   5760
            TabIndex        =   164
            Top             =   240
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCorProcesos.frx":9631
            PICN            =   "frmCorProcesos.frx":964D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker DtpFinalizarSoldadura 
            Height          =   375
            Left            =   1320
            TabIndex        =   204
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
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
            CalendarBackColor=   16777215
            Format          =   173408257
            CurrentDate     =   42314
         End
         Begin VB.Label Label35 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COMENTARIO :"
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
            Left            =   120
            TabIndex        =   165
            Top             =   360
            Width           =   1185
         End
      End
      Begin VB.Frame fraSolFin 
         BackColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   6600
         TabIndex        =   61
         Top             =   1920
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox txtObservacionSol2 
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
            Height          =   525
            Left            =   1320
            TabIndex        =   62
            Top             =   240
            Width           =   4455
         End
         Begin VitekeySoft.ChameleonBtn cmdConfFinSol 
            Height          =   375
            Left            =   4200
            TabIndex        =   72
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "CONFIRMAR"
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
            BCOL            =   8421631
            BCOLO           =   8421631
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCorProcesos.frx":9BE7
            PICN            =   "frmCorProcesos.frx":9C03
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdCerrarSol2 
            Height          =   285
            Left            =   6000
            TabIndex        =   116
            Top             =   120
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCorProcesos.frx":A19D
            PICN            =   "frmCorProcesos.frx":A1B9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker DtpFinEnsablaje 
            Height          =   345
            Left            =   1320
            TabIndex        =   205
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
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
            CalendarBackColor=   16777215
            Format          =   173408257
            CurrentDate     =   42314
         End
         Begin VB.Label Label37 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COMENTARIO :"
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
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   1185
         End
      End
      Begin VB.Frame fraSol 
         BackColor       =   &H00FFFFFF&
         Height          =   3300
         Left            =   6600
         TabIndex        =   53
         Top             =   120
         Visible         =   0   'False
         Width           =   6375
         Begin VB.TextBox txtObservacionSol 
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
            Height          =   375
            Left            =   600
            TabIndex        =   54
            Top             =   2880
            Width           =   2655
         End
         Begin MSDataListLib.DataCombo cboEmpresaSol 
            Height          =   315
            Left            =   1800
            TabIndex        =   55
            Top             =   240
            Width           =   4020
            _ExtentX        =   7091
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridTrabajadorSol 
            Height          =   1215
            Left            =   480
            TabIndex        =   56
            Top             =   600
            Width           =   5325
            _ExtentX        =   9393
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
         Begin VitekeySoft.ChameleonBtn cmdConfirmarSol 
            Height          =   375
            Left            =   4920
            TabIndex        =   77
            Top             =   2880
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "CONFIRMAR"
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
            BCOL            =   8421631
            BCOLO           =   8421631
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCorProcesos.frx":A753
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdCerrarSol 
            Height          =   285
            Left            =   6000
            TabIndex        =   115
            Top             =   240
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCorProcesos.frx":A76F
            PICN            =   "frmCorProcesos.frx":A78B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfEnsambladoInsumo 
            Height          =   855
            Left            =   480
            TabIndex        =   151
            Top             =   1920
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   1508
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
         Begin VitekeySoft.ChameleonBtn cmdagregarsol 
            Height          =   285
            Left            =   6000
            TabIndex        =   152
            Top             =   1920
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCorProcesos.frx":AD25
            PICN            =   "frmCorProcesos.frx":AD41
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdquitarsol 
            Height          =   285
            Left            =   6000
            TabIndex        =   153
            Top             =   2280
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   503
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
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCorProcesos.frx":B2DB
            PICN            =   "frmCorProcesos.frx":B2F7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker DtpInicioEnsamblaje 
            Height          =   345
            Left            =   3360
            TabIndex        =   206
            Top             =   2880
            Width           =   1455
            _ExtentX        =   2566
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
            CalendarBackColor=   16777215
            Format          =   173408257
            CurrentDate     =   42314
         End
         Begin VB.Label Label33 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "EMPRESA :"
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
            TabIndex        =   60
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label36 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OBS:"
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
            Left            =   165
            TabIndex        =   59
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label lblNombreTempSol 
            Caption         =   "-"
            Height          =   255
            Left            =   1920
            TabIndex        =   58
            Top             =   2880
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblDniTempSol 
            Caption         =   "-"
            Height          =   255
            Left            =   600
            TabIndex        =   57
            Top             =   2880
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VitekeySoft.ChameleonBtn cmdIniEns 
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   3600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "  INICIAR PROCESO"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":B891
         PICN            =   "frmCorProcesos.frx":B8AD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdFinEns 
         Height          =   375
         Left            =   4320
         TabIndex        =   71
         Top             =   3600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "  FINALIZAR PROCESO"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":BE47
         PICN            =   "frmCorProcesos.frx":BE63
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdIniSol 
         Height          =   375
         Left            =   6600
         TabIndex        =   73
         Top             =   3480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "  INICIAR PROCESO"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":C3FD
         PICN            =   "frmCorProcesos.frx":C419
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdIniTap 
         Height          =   375
         Left            =   13080
         TabIndex        =   74
         Top             =   3480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "  INICIAR PROCESO"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":C9B3
         PICN            =   "frmCorProcesos.frx":C9CF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdFinSol 
         Height          =   375
         Left            =   10800
         TabIndex        =   75
         Top             =   3480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "  FINALIZAR PROCESO"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":CF69
         PICN            =   "frmCorProcesos.frx":CF85
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdFinTap 
         Height          =   375
         Left            =   16560
         TabIndex        =   76
         Top             =   3480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "  FINALIZAR PROCESO"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":D51F
         PICN            =   "frmCorProcesos.frx":D53B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdreiniciarSoldadura 
         Height          =   375
         Left            =   2520
         TabIndex        =   167
         Top             =   3600
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "REINICIAR"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":DAD5
         PICN            =   "frmCorProcesos.frx":DAF1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdreiniciarensamblaje 
         Height          =   375
         Left            =   9000
         TabIndex        =   168
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "REINICIAR"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":E08B
         PICN            =   "frmCorProcesos.frx":E0A7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdreiniciartapiz 
         Height          =   375
         Left            =   15120
         TabIndex        =   169
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "REINICIAR"
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCorProcesos.frx":E641
         PICN            =   "frmCorProcesos.frx":E65D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TAPICERIA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   15000
         TabIndex        =   52
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INICIO"
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
         Left            =   13440
         TabIndex        =   51
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPRESA :"
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
         Left            =   13200
         TabIndex        =   50
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA :"
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
         Left            =   16560
         TabIndex        =   49
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label lblTaFechaIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   14295
         TabIndex        =   48
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA :"
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
         Left            =   16560
         TabIndex        =   47
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label lblTaHoraIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   14295
         TabIndex        =   46
         Top             =   1560
         Width           =   75
      End
      Begin VB.Label lblTaEmpresa 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   14520
         TabIndex        =   45
         Top             =   2160
         Width           =   4125
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRABAJADOR :"
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
         Left            =   13200
         TabIndex        =   44
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Label lblTaTrabajador 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   14520
         TabIndex        =   43
         Top             =   2640
         Width           =   4125
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Left            =   13200
         TabIndex        =   42
         Top             =   3120
         Width           =   1245
      End
      Begin VB.Label lblTaObservacion 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   14520
         TabIndex        =   41
         Top             =   3120
         Width           =   4125
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FIN"
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
         Left            =   16560
         TabIndex        =   40
         Top             =   840
         Width           =   285
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA :"
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
         Left            =   13440
         TabIndex        =   39
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label lblTaFechaFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   17415
         TabIndex        =   38
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA :"
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
         Left            =   13440
         TabIndex        =   37
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label lblTaHoraFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   17415
         TabIndex        =   36
         Top             =   1560
         Width           =   75
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOLDADURA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   2160
         TabIndex        =   35
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INICIO"
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
         Left            =   7200
         TabIndex        =   34
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPRESA :"
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
         Left            =   6960
         TabIndex        =   33
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA :"
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
         Left            =   7200
         TabIndex        =   32
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label lblSoFechaIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   8175
         TabIndex        =   31
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA :"
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
         Left            =   7200
         TabIndex        =   30
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label lblSoHoraIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   8175
         TabIndex        =   29
         Top             =   1560
         Width           =   75
      End
      Begin VB.Label lblSoEmpresa 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         Height          =   255
         Left            =   8520
         TabIndex        =   28
         Top             =   2160
         Width           =   4245
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRABAJADOR :"
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
         Left            =   6960
         TabIndex        =   27
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Label lblSoTrabajador 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         Height          =   255
         Left            =   8520
         TabIndex        =   26
         Top             =   2640
         Width           =   4245
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Left            =   6960
         TabIndex        =   25
         Top             =   3120
         Width           =   1245
      End
      Begin VB.Label lblSoObservacion 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         Height          =   255
         Left            =   8520
         TabIndex        =   24
         Top             =   3120
         Width           =   4245
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FIN"
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
         Left            =   10080
         TabIndex        =   23
         Top             =   840
         Width           =   285
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA :"
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
         Left            =   10080
         TabIndex        =   22
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label lblSoFechaFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   11055
         TabIndex        =   21
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA :"
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
         Left            =   10080
         TabIndex        =   20
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label lblSoHoraFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   11055
         TabIndex        =   19
         Top             =   1560
         Width           =   75
      End
      Begin VB.Label lblEmHoraFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3615
         TabIndex        =   18
         Top             =   1560
         Width           =   75
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA :"
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
         Left            =   2880
         TabIndex        =   17
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label lblEmFechaFin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3615
         TabIndex        =   16
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA :"
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
         Left            =   2880
         TabIndex        =   15
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FIN"
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
         Left            =   3000
         TabIndex        =   14
         Top             =   840
         Width           =   285
      End
      Begin VB.Label lblEmObservacion 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1800
         TabIndex        =   13
         Top             =   3120
         Width           =   4575
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACIÓN :"
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
         Left            =   360
         TabIndex        =   12
         Top             =   3120
         Width           =   1245
      End
      Begin VB.Label lblEmTrabajador 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1800
         TabIndex        =   11
         Top             =   2640
         Width           =   4575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRABAJADOR :"
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
         Left            =   360
         TabIndex        =   10
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Label lblEmEmpresa 
         AutoSize        =   -1  'True
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1800
         TabIndex        =   9
         Top             =   2160
         Width           =   4575
      End
      Begin VB.Label lblEmHoraIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1215
         TabIndex        =   8
         Top             =   1560
         Width           =   75
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA :"
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
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label lblEmFechaIni 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1215
         TabIndex        =   6
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA :"
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
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPRESA :"
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
         Left            =   360
         TabIndex        =   4
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label INICIO 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INICIO"
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
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ENSAMBLAJE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   8640
         TabIndex        =   2
         Top             =   120
         Width           =   1935
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   1935
         Left            =   13080
         Top             =   120
         Width           =   5655
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   1335
         Left            =   13080
         Top             =   2085
         Width           =   5655
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   1935
         Left            =   6600
         Top             =   120
         Width           =   6375
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   1335
         Left            =   6600
         Top             =   2085
         Width           =   6375
      End
      Begin VB.Shape Shape7 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   1335
         Left            =   120
         Top             =   2085
         Width           =   6375
      End
      Begin VB.Shape Shape6 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   1935
         Left            =   120
         Top             =   120
         Width           =   6375
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdreportepagos 
      Height          =   400
      Left            =   15360
      TabIndex        =   166
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
      BTYPE           =   5
      TX              =   "REPORTE DE PAGOS"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCorProcesos.frx":EBF7
      PICN            =   "frmCorProcesos.frx":EC13
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdAdicionarProceso 
      Height          =   375
      Left            =   14040
      TabIndex        =   171
      Top             =   4800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "ADICIONAR PROCESO"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCorProcesos.frx":ECA0
      PICN            =   "frmCorProcesos.frx":ECBC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdhistorial 
      Height          =   400
      Left            =   16920
      TabIndex        =   184
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
      BTYPE           =   5
      TX              =   "HISTORIAL"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCorProcesos.frx":F256
      PICN            =   "frmCorProcesos.frx":F272
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
      Height          =   285
      Left            =   240
      TabIndex        =   195
      Top             =   360
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      ListField       =   "c"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VitekeySoft.ChameleonBtn cmdconvertir 
      Height          =   375
      Left            =   11640
      TabIndex        =   197
      Top             =   4800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "CONVERTIR"
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
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCorProcesos.frx":1232F
      PICN            =   "frmCorProcesos.frx":1234B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label62 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUCURSAL"
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
      Left            =   285
      TabIndex        =   196
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MODELO :"
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
      Left            =   6990
      TabIndex        =   124
      Top             =   240
      Width           =   795
   End
   Begin VB.Label lblProcesoCambiado 
      Caption         =   "-"
      Height          =   255
      Left            =   10800
      TabIndex        =   80
      Top             =   4665
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblchasis 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHASIS : "
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
      Left            =   4230
      TabIndex        =   68
      Top             =   480
      Width           =   705
   End
   Begin VB.Label lblmotor 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOTOR : "
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
      Left            =   4215
      TabIndex        =   67
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   9495
      Left            =   0
      Top             =   0
      Width           =   18990
   End
End
Attribute VB_Name = "frmCorProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Public Sub diagnostico(ByVal Grilla As MSHFlexGrid, param As String)

Dim color As String, edad As Double

 
                  
                  
Select Case param
  Case "0"
    strCadena = strCadena & " and d.id_estado in ('" & Me.cboEstado.BoundText & "')"

  Case "1"
    If Me.chkestado.Value = 1 Then
        strCadena = strCadena & " and d.id_estado in ('" & Me.cboEstado.BoundText & "') and p.`nombre_prod` like '%" & Trim(Me.txtModelo.Text) & "%' "
    Else
        strCadena = strCadena & " and p.`nombre_prod` like '%" & Trim(Me.txtModelo.Text) & "%' "
    End If
    
  
  Case "2"
    If Me.chkestado.Value = 1 Then
        strCadena = strCadena & " and d.id_estado in ('" & Me.cboEstado.BoundText & "') and d.`nro_motor` like '%" & Me.TxtMotor.Text & "%' "
    Else
        strCadena = strCadena & " and d.`nro_motor` like '%" & Me.TxtMotor.Text & "%' "
    End If
    
  Case "3"
    If Me.chkestado.Value = 1 Then
        strCadena = strCadena & " and d.id_estado in ('" & Me.cboEstado.BoundText & "') and d.`nro_chasis` like '%" & Me.txtchasis.Text & "%' "
    Else
        strCadena = strCadena & " and d.`nro_chasis` like '%" & Me.txtchasis.Text & "%' "
    End If
End Select
                  
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If


  N = 1
 
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 0
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 4500
           Grilla.ColWidth(4) = 2500
           
           Grilla.ColWidth(5) = 2500
          
           Grilla.ColWidth(6) = 2000
           Grilla.ColWidth(7) = 1000
           Grilla.ColWidth(8) = 1000
           Grilla.ColWidth(9) = 850
           Grilla.ColWidth(10) = 1500
           Grilla.ColWidth(11) = 1500
          
           
           
        Next
         cabecera = "" & vbTab & vbTab & "" & "CODIGO" & vbTab & "PRODUCTO" & vbTab & Me.lblchasis.Caption & vbTab & Me.lblmotor.Caption & vbTab & "MARCA" & vbTab & "AÑO.FAB" & vbTab & "CONTENEDOR" & vbTab & "ITEM" & vbTab & "ESTADO" & vbTab & "UBICACION"
         Grilla.AddItem cabecera
         For k = 1 To 11
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        Grilla.ColAlignment(4) = 7
        Grilla.ColAlignment(5) = 7
        Grilla.ColAlignment(6) = 4
        Grilla.ColAlignment(7) = 4
        Grilla.ColAlignment(8) = 4
        Grilla.ColAlignment(9) = 4
        For i = 0 To rst.RecordCount - 1
        
             estado = Chr(168)
        
             Fila = rst("id_detalle") & vbTab & est & vbTab & rst("id_producto") & vbTab & rst("producto") & Space(2) & "[" & rst("color") & "]" & vbTab & rst("nro_chasis") & vbTab & rst("nro_motor") & vbTab & rst("marca") & vbTab & rst("anio_fabricacion") & vbTab & rst("nro_contenedor") & vbTab & rst("item") & vbTab & rst("estado") & vbTab & rst("almacen")
             Grilla.AddItem Fila
              
             With Grilla
                 .Row = i + 1 ' se posiciona en la fila
                 .col = 1 '  .. en la columna
                 ' cambia la fuente para esta celda
                            
                 .CellFontName = "Wingdings"
                 .CellFontSize = 14
                 .CellAlignment = flexAlignCenterCenter
    
              End With
             
             
             Fila = ""
             rst.MoveNext
        Next i
        
        
Exit Sub


End Sub




Public Sub actualizaGrid(ByVal Grilla As MSHFlexGrid, tipo As String, id_detalle As Integer)
Dim color As String, edad As Double
  
Me.cboEstado.BoundText = tipo
 
strCadena = " select d.`id_detalle`, d.`id_detalle_compra` , p.`id_producto`, " & _
  " p.`nombre_prod` as producto, d.`serie`, d.`anio_fabricacion`, " & _
  " d.`anio_contenedor`, d.`anio_modelo`, d.`nro_chasis`, d.`nro_contenedor`, " & _
  " d.`nro_motor`, d.`item`, e.`descripcion` as estado, e.`id_estado`,cc.descripcion as color from `imp_producto_detalle` d, movimiento_compra_detalle c, " & _
  " producto p , imp_estado e ,imp_color cc  where p.id_color=cc.id_color and d.`id_detalle_compra` = c.`id_detalle_compra` and " & _
  " c.`id_producto` = p.`id_producto` and " & _
  " p.`ruc` = '" & KEY_RUC & "' and " & _
  " e.`id_estado` = d.`id_estado` and d.id_estado in ('" & Me.cboEstado.BoundText & "')  "
  
  
  
 
   
   
   
            
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
           Grilla.ColWidth(1) = 0
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 5500
           Grilla.ColWidth(4) = 2500
           Grilla.ColWidth(5) = 2500
           Grilla.ColWidth(6) = 1000
           Grilla.ColWidth(7) = 1000
           Grilla.ColWidth(8) = 1500
           Grilla.ColWidth(9) = 800
           Grilla.ColWidth(10) = 2500
           
        Next
         
         cabecera = "" & vbTab & vbTab & "" & "CODIGO" & vbTab & "PRODUCTO" & vbTab & "MOTOR" & vbTab & "CHASIS" & vbTab & "AÑO FAB." & vbTab & "AÑO CONTENEDOR" & vbTab & "CONTENEDOR" & vbTab & "ITEM" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 10
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        
        
        Dim Indice As Integer
        Indice = 0
             
        For i = 0 To rst.RecordCount - 1
        
             estado = Chr(168)
        
             
             Fila = rst("id_detalle") & vbTab & est & vbTab & rst("id_producto") & vbTab & rst("producto") & Space(2) & "[" & rst("color") & "]" & vbTab & rst("nro_motor") & vbTab & rst("nro_chasis") & vbTab & rst("anio_fabricacion") & vbTab & rst("anio_contenedor") & vbTab & rst("nro_contenedor") & vbTab & rst("item") & vbTab & rst("estado")
             
             Grilla.AddItem Fila
             
             With Grilla
                 .Row = i + 1 ' se posiciona en la fila
                 .col = 1 '  .. en la columna
                 ' cambia la fuente para esta celda
                            
                 .CellFontName = "Wingdings"
                 .CellFontSize = 14
                 .CellAlignment = flexAlignCenterCenter
    
              End With
             
             
             
             If rst("id_detalle") = id_detalle Then
                Indice = Grilla.Row
                
                For j = 0 To 10
                Grilla.col = j
                Grilla.CellBackColor = &HC0FFC0
                Next j
                
             End If
             
             Fila = ""
             rst.MoveNext
        Next i
        
        Grilla.Row = Indice


End Sub





Private Sub btnBuscar_Click()
    
End Sub

Private Sub cboEmpresa_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call llena_trabajadores("123", Me.gridTrabajador, Me.cboEmpresa.BoundText)
  End If
End Sub




Private Sub cboEmpresaEns2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call llena_trabajadores("123", Me.gridTrabajadorEns2, Me.cboEmpresaEns2)
  End If
End Sub

Private Sub cboEmpresaSol_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call llena_trabajadores("123", Me.gridTrabajadorSol, Me.cboEmpresaSol.BoundText)
  End If
End Sub

Private Sub cboEmpresaTap_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Call llena_trabajadores("123", Me.gridTrabajadorTap, Me.cboEmpresaTap.BoundText)
     
  End If
End Sub


Private Sub cboEstado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Dim clausula As String
  clausula = ""
       
  'kkk
  
  
    If txtModelo.Text <> "" Then
        clausula = clausula & " and p.`nombre_prod` like '%" & Trim(Me.txtModelo.Text) & "%' "
    End If
    If TxtMotor.Text <> "" Then
        clausula = clausula & " and d.`nro_motor` like '%" & Me.TxtMotor.Text & "%' "
    End If
    If txtchasis.Text <> "" Then
        clausula = clausula & " and d.`nro_chasis` like '%" & Me.txtchasis.Text & "%' "
    End If
    
  strCadena = " select d.`id_detalle`, d.`id_detalle_compra` , p.`id_producto`, " & _
  " p.`nombre_prod` as producto, d.`serie`, d.`anio_fabricacion`, " & _
  " d.`anio_contenedor`, d.`anio_modelo`, d.`nro_chasis`, d.`nro_contenedor`, " & _
  " d.`nro_motor`, d.`item`, e.`descripcion` as estado, e.`id_estado`,cc.descripcion as color,m.descripcion as marca,a.descripcion as almacen from `imp_producto_detalle` d, movimiento_compra_detalle c, " & _
  " producto p , imp_estado e,imp_color cc,marca m ,almacen a where  a.id_alm=d.id_alm and a.ruc=d.ruc and m.id_marca=p.id_marca and m.id_usu=p.ruc and   p.id_color=cc.id_color and   d.`id_detalle_compra` = c.`id_detalle_compra` and " & _
  " c.`id_producto` = p.`id_producto`  and  d.id_alm='" & Me.DtcAlmacen.BoundText & "' and  " & _
  " p.`ruc` = '" & KEY_RUC & "' and " & _
  " e.`id_estado` = d.`id_estado`  " & clausula
  
    Call diagnostico(Me.gridDetalle, "0")
End If
End Sub

Private Sub ChameleonBtn4_Click()

End Sub

Private Sub ChameleonBtn3_Click()

End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub cmdAdicionarProceso_Click()
    Me.Txtid_estado.Text = "05"
    Me.FrameTercero.Visible = True
    Call frmCorProcesos.llena_insumos_otro(gridDetalle.TextMatrix(gridDetalle.Row, 0), Txtid_estado.Text, Me.HfListadoInsumos)
    gridDetalle.Enabled = False
    
        
 End Sub

Private Sub cmdagregarsol_Click()
Procedencia = seleccionar_ensamblaje
FrmProducto.Show
Exit Sub
End Sub

Private Sub cmdagregarSoldadura_Click()
Procedencia = seleccionar_soldadura
FrmProducto.Show
Exit Sub
End Sub

Private Sub cmdagregartapiz_Click()
    
    Procedencia = seleccionar_tapiz
    FrmProducto.Show
    Exit Sub
    
End Sub

Private Sub cmdagregarTercero_Click()
    
    Procedencia = seleccionar_otro
    FrmProducto.Show
    Exit Sub
    
End Sub

Private Sub cmdCambiar_Click()
  If Me.fraCambiar.Visible = True Then
    Me.fraCambiar.Visible = False
  Else
    Me.fraCambiar.Visible = True
  End If
  
  
End Sub

Private Sub cmdCerrarEns_Click()
  Me.fraEnsamblado.Visible = False
End Sub

Private Sub cmdCerrarEns2_Click()
   Me.fraEnsambladoFin.Visible = False
End Sub

Private Sub cmdcerrarhistorial_Click()
    Me.framehistorial.Visible = False
End Sub

Private Sub cmdCerrarOtro_Click()
    Me.FrameTercero.Visible = False
    gridDetalle.Enabled = True
End Sub

Private Sub cmdCerrarSol_Click()
   Me.fraSol.Visible = False
End Sub

Private Sub cmdCerrarSol2_Click()
    Me.fraSolFin.Visible = False
End Sub

Private Sub cmdCerrarTap_Click()
   Me.fraTap.Visible = False
End Sub


Private Sub cmdCerrarTap2_Click()
   Me.fraTapFin.Visible = False
End Sub

Private Sub cmdConfFinEns_Click()
    rowTemp = Me.gridDetalle.Row
    
   strCadena = " select id_mov from imp_producto_movimiento where id_detalle = " & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0) & _
               " and id_proceso = '01' "
   Call ConfiguraRstT(strCadena)
   
   strCadena = "update imp_producto_movimiento set fecha_salida = '" & Format(Me.DtpFinalizarSoldadura.Value, "YYYY-mm-dd") & "',hora_salida = '" & Format(Now(), "HH:mm:ss") & "', estado = '1', observacion = '" & Me.txtObservacion2.Text & "' " & _
               " where id_mov =" & rstT("id_mov")
   CnBd.Execute strCadena
               
   
   'Call diagnostico(Me.gridDetalle)
   Call actualizaGrid(Me.gridDetalle, "02", Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0))
   
   Me.gridDetalle.Row = rowTemp
   actualizarProcesos
End Sub

Private Sub cmdConfFinEns2_Click()
   rowTemp = Me.gridDetalle.Row
    
   strCadena = " select id_mov from imp_producto_movimiento where id_detalle = " & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0) & _
               " and id_proceso = '02' "
   Call ConfiguraRstT(strCadena)
   
   strCadena = "update imp_producto_movimiento set fecha_salida = '" & KEY_FECHA & "',hora_salida = '" & Format(Now(), "HH:mm:ss") & "', estado = '1', observacion = '" & Me.txtObservacionEns2_2.Text & "' " & _
               " where id_mov =" & rstT("id_mov")
   CnBd.Execute strCadena
               
   strCadena = "update imp_producto_detalle set id_estado = '03' where id_detalle = " & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0)
   CnBd.Execute strCadena
  
   'Call diagnostico(Me.gridDetalle)
   Call actualizaGrid(Me.gridDetalle, "03", Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0))
   
   'Me.gridDetalle.row = rowTemp
   actualizarProcesos
End Sub

Private Sub cmdConfFinSol_Click()
   rowTemp = Me.gridDetalle.Row
    
   strCadena = " select id_mov from imp_producto_movimiento where id_detalle = " & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0) & _
               " and id_proceso = '02' "
   Call ConfiguraRstT(strCadena)
   
   strCadena = "update imp_producto_movimiento set fecha_salida = '" & Format(Me.DtpFinEnsablaje.Value, "YYYY-mm-dd") & "',hora_salida = '" & Format(Now(), "HH:mm:ss") & "', estado = '1', observacion = '" & Me.txtObservacionSol2.Text & "' " & _
               " where id_mov =" & rstT("id_mov")
   CnBd.Execute strCadena
               
   
   'Call diagnostico(Me.gridDetalle)
   Call actualizaGrid(Me.gridDetalle, "02", Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0))
   
   Me.gridDetalle.Row = rowTemp
   actualizarProcesos
End Sub

Private Sub cmdConfFinTap_Click()
  rowTemp = Me.gridDetalle.Row
    
   strCadena = " select id_mov from imp_producto_movimiento where id_detalle = " & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0) & _
               " and id_proceso = '03' limit 0,1 "
   Call ConfiguraRstT(strCadena)
   If rstT.RecordCount > 0 Then
   strCadena = "update imp_producto_movimiento set fecha_salida = '" & Format(Me.DtpFinTapizado.Value, "YYYY-mm-dd") & "',hora_salida = '" & Format(Now(), "HH:mm:ss") & "', estado = '1', observacion = '" & Me.txtObservacionTap2.Text & "' " & _
               " where id_mov =" & rstT("id_mov")
   CnBd.Execute strCadena
   End If
   strCadena = "update imp_producto_detalle set id_estado = '03' where id_detalle = " & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0)
   
   CnBd.Execute strCadena
  
   'Call diagnostico(Me.gridDetalle)
   Call actualizaGrid(Me.gridDetalle, "03", Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0))
   
   'Me.gridDetalle.row = rowTemp
   actualizarProcesos
End Sub

Private Sub cmdConfirmar_Click()
   errores = ""
   
   If Me.lblDniTemp.Caption = "-" Then
     errores = "-Debe seleccionar trabajador."
   End If
   
   If Not errores = "" Then
     MsgBox errores
     Exit Sub
   End If

   rowTemp = Me.gridDetalle.Row


   strCadena = "insert into imp_producto_movimiento (id_detalle,id_proceso,fecha_entrada,hora_entrada,estado, observacion, id_autor, ruc_empresa, ruc) " & _
               " values (" & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0) & ",'01', '" & Format(Me.DtpfechaInicio.Value, "YYYY-mm-dd") & "', '" & Format(Now(), "HH:mm:ss") & "', '0','" & Me.TxtObservacion.Text & "','" & Me.lblDniTemp.Caption & "','" & Me.cboEmpresa.BoundText & "','" & KEY_RUC & "')"
   CnBd.Execute strCadena
                         
   strCadena = "update imp_producto_detalle set id_estado = '02', id_estado_detalle = '01' where id_detalle = " & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0)
   CnBd.Execute strCadena
   
   'Call diagnostico(Me.gridDetalle)
   Call actualizaGrid(Me.gridDetalle, "02", Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0))
   
   'Me.gridDetalle.row = rowTemp
   actualizarProcesos
End Sub

Private Sub cmdConfirmarEns2_Click()
  errores = ""
   
   If Me.lblDniTempEns2.Caption = "-" Then
     errores = "-Debe seleccionar trabajador."
   End If
   
   If Not errores = "" Then
     MsgBox errores
     Exit Sub
   End If

   rowTemp = Me.gridDetalle.Row


   strCadena = "insert into imp_producto_movimiento (id_detalle,id_proceso,fecha_entrada,hora_entrada,estado, observacion, id_autor, ruc_empresa, ruc) " & _
               " values (" & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0) & ",'02', '" & KEY_FECHA & "', '" & Format(Now(), "HH:mm:ss") & "', '0','" & Me.txtObservacionEns_2.Text & "','" & Me.lblDniTempEns2.Caption & "','" & Me.cboEmpresaEns2.BoundText & "','" & KEY_RUC & "')"
   CnBd.Execute strCadena
                         
   strCadena = "update imp_producto_detalle set id_estado = '02', id_estado_detalle = '02' where id_detalle = " & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0)
   CnBd.Execute strCadena
   
   'Call diagnostico(Me.gridDetalle)
   Call actualizaGrid(Me.gridDetalle, "02", Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0))
   
   'Me.gridDetalle.row = rowTemp
   actualizarProcesos
   
End Sub

Private Sub cmdConfirmarSol_Click()
   errores = ""
   
   If Me.lblDniTempSol.Caption = "-" Then
     errores = "-Debe seleccionar trabajador."
   End If
   
   If Not errores = "" Then
     MsgBox errores
     Exit Sub
   End If
   

   rowTemp = Me.gridDetalle.Row


   strCadena = "insert into imp_producto_movimiento (id_detalle,id_proceso,fecha_entrada,hora_entrada,estado, observacion, id_autor, ruc_empresa, ruc) " & _
               " values (" & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0) & ",'02', '" & Format(Me.DtpInicioEnsamblaje.Value, "YYYY-mm-dd") & "', '" & Format(Now(), "HH:mm:ss") & "', '0','" & Me.TxtObservacion.Text & "','" & Me.lblDniTempSol.Caption & "','" & Me.cboEmpresaSol.BoundText & "','" & KEY_RUC & "')"
   CnBd.Execute strCadena
               
   strCadena = "update imp_producto_detalle set id_estado = '02', id_estado_detalle = '02' where id_detalle = " & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0)
   CnBd.Execute strCadena
   'Call diagnostico(Me.gridDetalle)
   Call actualizaGrid(Me.gridDetalle, "02", Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0))
   
   
   'Me.gridDetalle.row = rowTemp
   actualizarProcesos
End Sub

Private Sub cmdConfirmarTap_Click()
    errores = ""
   
   If Me.lblDniTempTap.Caption = "-" Then
     errores = "-Debe seleccionar trabajador."
   End If
   
   If Not errores = "" Then
     MsgBox errores
     Exit Sub
   End If
   
   rowTemp = Me.gridDetalle.Row

   strCadena = "insert into imp_producto_movimiento (id_detalle,id_proceso,fecha_entrada,hora_entrada,estado, observacion, id_autor, ruc_empresa, ruc) " & _
               " values (" & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0) & ",'03', '" & Format(Me.DtpInicioTapizado.Value, "YYYY-mm-dd") & "', '" & Format(Now(), "HH:mm:ss") & "', '0','" & Me.TxtObservacion.Text & "','" & Me.lblDniTempTap.Caption & "','" & Me.cboEmpresaTap.BoundText & "','" & KEY_RUC & "')"
   CnBd.Execute strCadena
               
   strCadena = "update imp_producto_detalle set id_estado = '02', id_estado_detalle = '03' where id_detalle = " & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0)
   CnBd.Execute strCadena
   
   'Call diagnostico(Me.gridDetalle)
   Call actualizaGrid(Me.gridDetalle, "02", Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0))
   
   
   'Me.gridDetalle.row = rowTemp
   
   actualizarProcesos
End Sub

Private Sub cmdconvertir_Click()

strCadena = "SELECT id_producto as Codigo,nombre_prod as Descripcion FROM producto  WHERE ruc='" & KEY_RUC & "' LIMIT 20"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcProducto)
Me.frmconvertir.Visible = True
End Sub

Private Sub cmdFinEns_Click()
   Me.fraEnsamblado.Visible = False
   Me.fraEnsambladoFin.Visible = True
   
   'Me.cmdConfirmar.Visible = False
   Me.cmdConfFinEns.Visible = True
End Sub

Private Sub cmdFinEns2_Click()
   'Me.fraSol.Visible = True
   Me.fraEns2Fin.Visible = True
   'Me.cmdConfirmarSol.Visible = False
   Me.cmdConfFinEns2.Visible = True
End Sub

Private Sub cmdFinSol_Click()
   'Me.fraSol.Visible = True
   Me.fraSolFin.Visible = True
   'Me.cmdConfirmarSol.Visible = False
   Me.cmdConfFinSol.Visible = True
End Sub

Private Sub cmdFinTap_Click()
   'Me.fraTap.Visible = True
   Me.fraTapFin.Visible = True
   'Me.cmdConfirmarTap.Visible = False
   Me.cmdConfFinTap.Visible = True
End Sub

Private Sub cmdhistorial_Click()
Dim nro_chasis As String
Dim nro_motor As String
Dim in_producto As String

If Me.gridDetalle.Rows < 1 Then
   Exit Sub
End If

nro_chasis = Trim(gridDetalle.TextMatrix(gridDetalle.Row, 4))
nro_motor = Trim(gridDetalle.TextMatrix(gridDetalle.Row, 5))
in_producto = Trim(gridDetalle.TextMatrix(gridDetalle.Row, 2))
Me.lblid_detalle_serie.Caption = Val(gridDetalle.TextMatrix(gridDetalle.Row, 0))


strCadena = "SELECT CONCAT(D.doc_abrev,':',C.serie,'-',C.numero) as compra,C.nproveedor,C.fecha_emision,A.descripcion  FROM imp_producto_detalle I,movimiento_compra C,comprobantes D,almacen A WHERE  C.id_alm=A.id_alm and C.ruc=A.ruc and C.id_doc=D.id_doc and   I.id_compra=C.id_compra AND I.id_detalle='" & Val(Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0)) & "'"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
   Me.lblfecha_ingreso.Caption = Format(rstZ("fecha_emision"), "YYYY-mm-dd")
   Me.lbldoc_ingreso.Caption = rstZ("compra") & Space(1)
   Me.lblalmacen.Caption = rstZ("descripcion")
Else
    Me.lbldoc_ingreso.Caption = ""
End If

strCadena = "SELECT * FROM movimiento_transferencia_series S,movimiento_transferencia T WHERE S.id_transferencia=T.id_transferencia AND S.chasis='" & nro_chasis & "' and  T.ruc='" & KEY_RUC & "'"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
    Me.lblfecha_transfer.Caption = Format(rstZ("fecha"), "YYYY-mm-dd")
    Me.lbldoc_transfer.Caption = "GUIA REMISION:" & rstZ("serie") & "-" & rstZ("numero")
Else
    Me.lblfecha_transfer.Caption = ""
    Me.lbldoc_transfer.Caption = ""
End If

Call llena_historial_motor(in_producto, nro_chasis, nro_motor, HfHistorial)


Me.framehistorial.Visible = True

End Sub

Private Sub cmdIniEns_Click()
  
   Me.DtpfechaInicio.Value = KEY_FECHA
   Me.fraEnsamblado.Visible = True
   Me.cmdConfirmar.Visible = True
   Me.cmdConfFinEns.Visible = False
   Me.Txtid_estado.Text = "01"
   
   strCadena = "SELECT id_linea from producto where id_producto='" & Trim(Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 2)) & "' and ruc='" & KEY_RUC & "' "
   Call ConfiguraRstZ(strCadena)
   
   If rstZ.RecordCount > 0 Then
        strCadena = "SELECT id_producto FROM linea_produccion_detalle WHERE estado='si' and  id_linea='" & rstZ("id_linea") & "' and id_produccion='" & Trim(Me.Txtid_estado.Text) & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
            Call agregar_proceso(rstT("id_producto"), Trim(Me.Txtid_estado.Text))
        End If
   End If
   
   
   Call Me.llena_insumos(Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0), Trim(Me.Txtid_estado.Text), Me.HfSoldaduraInsumo)
   
End Sub

Private Sub cmdIniEns2_Click()
  Me.fraEns2.Visible = True
  Me.cmdConfirmarEns2.Visible = True
  Me.cmdConfFinEns2.Visible = False
End Sub

Private Sub cmdIniSol_Click()
   Dim in_producto As String
   Me.fraSol.Visible = True
   Me.cmdConfirmarSol.Visible = True
   Me.cmdConfFinSol.Visible = False
   Me.Txtid_estado.Text = "02"
   If Trim(Me.txtreiniciar.Text) = "si" Then
      strCadena = "SELECT * FROM imp_producto_movimiento WHERE id_detalle='" & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0) & "' and id_proceso='" & Trim(Me.Txtid_estado.Text) & "'"
      Call ConfiguraRstZ(strCadena)
      Call llena_trabajadores("123", Me.gridTrabajadorSol, rstZ("ruc_empresa"))
      Call Me.llena_insumos(Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0), Trim(Me.Txtid_estado.Text), Me.HfEnsambladoInsumo)
      Exit Sub
   End If
   
   
   
   strCadena = "SELECT id_linea from producto where id_producto='" & Trim(Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 2)) & "' and ruc='" & KEY_RUC & "' "
   Call ConfiguraRstZ(strCadena)
   
   If rstZ.RecordCount > 0 Then
        strCadena = "SELECT id_producto FROM linea_produccion_detalle WHERE estado='si' and  id_linea='" & rstZ("id_linea") & "' and id_produccion='02' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstT(strCadena)
        Call agregar_proceso(rstT("id_producto"), "02")
   End If
   
   
   Call Me.llena_insumos(Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0), Trim(Me.Txtid_estado.Text), Me.HfEnsambladoInsumo)

End Sub
Private Sub agregar_proceso(ByVal id_producto As String, ByVal id_estado As String)

strCadena = "SELECT precio_compra FROM almacen_producto where id_producto='" & id_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
    strCadena = "INSERT INTO imp_producto_insumo(id_producto_detalle,id_producto,cantidad,id_linea,precio,ruc) " & _
    "VALUES('" & Val(Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0)) & "','" & id_producto & "','1','" & Trim(Me.Txtid_estado.Text) & "','" & rstZ("precio_compra") & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
End If

End Sub
Private Sub cmdIniTap_Click()
   Me.fraTap.Visible = True
   Me.cmdConfirmarTap.Visible = True
   Me.cmdConfFinTap.Visible = False
   Me.Txtid_estado.Text = "03"
   
   strCadena = "SELECT id_linea from producto where id_producto='" & Trim(Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 2)) & "' and ruc='" & KEY_RUC & "' "
   Call ConfiguraRstZ(strCadena)
   
   If rstZ.RecordCount > 0 Then
        strCadena = "SELECT id_producto FROM linea_produccion_detalle WHERE estado='si' and  id_linea='" & rstZ("id_linea") & "' and id_produccion='" & Trim(Me.Txtid_estado.Text) & "' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
            Call agregar_proceso(rstT("id_producto"), Trim(Me.Txtid_estado.Text))
        End If
   End If
   
   
   Call Me.llena_insumos(Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0), Trim(Me.Txtid_estado.Text), Me.HfTapizadoInsumo)
End Sub

Private Sub cmdMarcarVendido_Click()
Dim in_producto As String

in_producto = gridDetalle.TextMatrix(gridDetalle.Row, 2)

strCadena = "UPDATE imp_producto_detalle SET vendido='si' WHERE  id_producto='" & in_producto & "' and  id_detalle='" & Val(Me.lblid_detalle_serie.Caption) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
MsgBox "Proceso Realizado.", vbInformation

End Sub

Private Sub cmdOkCambiar_Click()
  Call ocultarFrames
  Me.lblProcesoCambiado.Caption = Me.cboCambiar.BoundText
  
  Call evaluarProcesoInicial(Me.lblProcesoCambiado.Caption)
End Sub

Private Sub evaluarProcesoInicial(ByVal id_proceso As String)
   
   Select Case id_proceso
   
   Case "01"
     Call ocultarBotones
     cmdIniEns.Visible = True
     
   Case "02"
     Call ocultarBotones
     cmdIniSol.Visible = True
     
   Case "03"
     Call ocultarBotones
     cmdIniTap.Visible = True
     
   End Select
  
End Sub

Private Sub cmdquitarsol_Click()
If MsgBox("Esta Seguro de quitar este Insumo", vbInformation + vbYesNo) = vbYes Then
    strCadena = "DELETE FROM imp_producto_insumo WHERE id='" & Me.HfEnsambladoInsumo.TextMatrix(Me.HfEnsambladoInsumo.Row, 0) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    Call Me.llena_insumos(Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0), Trim(Me.Txtid_estado.Text), Me.HfEnsambladoInsumo)
End If
End Sub

Private Sub cmdquitarSoldadura_Click()
If MsgBox("Esta Seguro de quitar este Insumo", vbInformation + vbYesNo) = vbYes Then
    strCadena = "DELETE FROM imp_producto_insumo WHERE id='" & Me.HfSoldaduraInsumo.TextMatrix(Me.HfSoldaduraInsumo.Row, 0) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    Call Me.llena_insumos(Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0), Trim(Me.Txtid_estado.Text), Me.HfSoldaduraInsumo)
End If
End Sub

Private Sub cmdquitartapiz_Click()
If MsgBox("Esta Seguro de quitar este Insumo", vbInformation + vbYesNo) = vbYes Then
    strCadena = "DELETE FROM imp_producto_insumo WHERE id='" & Me.HfTapizadoInsumo.TextMatrix(Me.HfTapizadoInsumo.Row, 0) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    Call Me.llena_insumos(Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0), Trim(Me.Txtid_estado.Text), Me.HfTapizadoInsumo)
End If

End Sub

Private Sub cmdquitarTercero_Click()
If MsgBox("Esta Seguro de quitar este Insumo", vbInformation + vbYesNo) = vbYes Then
    strCadena = "DELETE FROM imp_producto_insumo WHERE id='" & Me.HfListadoInsumos.TextMatrix(Me.HfListadoInsumos.Row, 0) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    Call llena_insumos_otro(gridDetalle.TextMatrix(gridDetalle.Row, 0), Txtid_estado.Text, Me.HfListadoInsumos)
    Me.cmdquitarTercero.Visible = False
End If

End Sub

Private Sub cmdreiniciarensamblaje_Click()
Me.Txtid_estado.Text = "02"
Procedencia = modificar
FrmSeguridad.Show
Exit Sub
End Sub

Private Sub cmdreiniciarSoldadura_Click()
Me.Txtid_estado.Text = "01"
Procedencia = modificar
FrmSeguridad.Show
Exit Sub
End Sub

Private Sub cmdreiniciartapiz_Click()
Me.Txtid_estado.Text = "03"
Procedencia = modificar
FrmSeguridad.Show
Exit Sub

End Sub

Private Sub cmdreportepagos_Click()
frmKorea.Show
End Sub

Private Sub cmdreportes_Click()
frmCorReportes.Show


End Sub

Private Sub cmdsalir_Click()
     Unload Me
End Sub

Private Sub DtcEmpresaTercero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call llena_trabajadores("123", Me.HgTrabajadorOtro, Me.DtcEmpresaTercero.BoundText)
    
    strCadena = "SELECT a.id_producto,p.nombre_prod,c.descripcion as color,a.precio_compra FROM producto p,almacen_producto a,imp_color c WHERE  p.id_color=c.id_color and   p.id_producto=a.id_producto and a.id_alm='" & KEY_ALM & "' and id_proveedor='" & Me.DtcEmpresaTercero.BoundText & "' and  a.ruc=p.ruc and  a.ruc='" & KEY_RUC & "'"
    Call llena_productos(HfInsumoTercero, Me.DtcEmpresaTercero.BoundText)
    
End If
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)


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
Private Sub Form_Load()
CenterForm Me
Me.Top = 100
 Call parametros_busqueda

Me.cmdreiniciarensamblaje.Visible = False
Me.cmdreiniciarSoldadura.Visible = False
Me.cmdreiniciartapiz.Visible = False

Me.DtpfechaInicio.Value = KEY_FECHA
Me.DtpFinalizarSoldadura.Value = KEY_FECHA
Me.DtpFinEnsablaje.Value = KEY_FECHA
Me.DtpFinTapizado.Value = KEY_FECHA
Me.DtpInicioEnsamblaje.Value = KEY_FECHA
Me.DtpInicioTapizado.Value = KEY_FECHA




strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE id_tipoentidad='0' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM


Call llenarEstados(Me.cboEstado)

strCadena = " select d.`id_detalle`, d.`id_detalle_compra` , p.`id_producto`, " & _
  " p.`nombre_prod` as producto, d.`serie`, d.`anio_fabricacion`, " & _
  " d.`anio_contenedor`, d.`anio_modelo`, d.`nro_chasis`, d.`nro_contenedor`, " & _
  " d.`nro_motor`, d.`item`, e.`descripcion` as estado, e.`id_estado`,cc.descripcion as color,m.descripcion as marca,a.descripcion as almacen from `imp_producto_detalle` d, movimiento_compra_detalle c, " & _
  " producto p , imp_estado e,imp_color cc,marca m,almacen a  where  d.id_alm=a.id_alm and d.ruc=a.ruc and p.id_marca=m.id_marca and m.id_usu=p.ruc and   p.id_color=cc.id_color and   d.`id_detalle_compra` = c.`id_detalle_compra` and " & _
  " c.`id_producto` = p.`id_producto` and " & _
  " p.`ruc` = '" & KEY_RUC & "' and " & _
  " e.`id_estado` = d.`id_estado` and d.id_alm='" & Me.DtcAlmacen.BoundText & "'  ORDER BY d.id_estado LIMIT 0,15 "
  
     Call diagnostico(Me.gridDetalle, "")
   
     Call llenarComboEmpresa(Me.cboEmpresa)
     Call llenarComboEmpresa(Me.cboEmpresaSol)
     Call llenarComboEmpresa(Me.cboEmpresaTap)
     
          
      Call llenarComboEmpresa(Me.DtcEmpresaTercero)
     
     'Call llenarComboEmpresa(Me.cboEmpresaEns2)
     
     Call llenarProcesos(Me.cboCambiar)
     
End Sub
Private Sub llenarComboEmpresa(ByVal cbo As DataCombo)
    'strCadena = " select codigo as Codigo, concat(codigo,' ',descripcion) as Descripcion from prestaciones "
    'strCadena = " select e.`cod_unico` as Codigo, p.`nombre_completo` as Descripcion from `entidad_empresa` e, persona p " & _
                " where  e.`cod_unico` = p.`dni` and e.`id_empresa` = '20531516045'"
                
    
    strCadena = " select e.cod_unico as Codigo, p.`nombre_completo` as Descripcion from entidad_empresa e, persona p " & _
                " where e.id_empresa='" & KEY_RUC & "' and e.id_almacen='si' and p.`dni` = e.`cod_unico`"
    
    Call ConfiguraRstT(strCadena)
    
    Call LlenaDataComboT(cbo)
    
End Sub



Public Sub llena_trabajadores(ByVal id_ruc As String, ByVal Grilla As MSHFlexGrid, ByVal ruc_empresa As String)
Dim porcentaje As Single
Grilla.SelectionMode = flexSelectionFree
Grilla.MergeCells = flexMergeFree
'strCadena = "select e.`cod_unico`, p.`nombre_completo` from `entidad_empresa` e, persona p " & _
           "where  e.`cod_unico` = p.`dni` and e.`id_empresa` = '20531516045' limit 5"
 strCadena = " select e.`cod_unico`, p.`nombre_completo` from `entidad_empresa` e, persona p " & _
 " where  e.`cod_unico` = p.`dni` and id_empresa='" & KEY_RUC & "' and id_empresa_rel= '" & ruc_empresa & "'"
           
Call ConfiguraRstT(strCadena)
'Call Cargar_FlexGrid(Me.HfActividades, 8, rstT)
porcentaje = 0
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
                           
       ' edita la celda
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 3500
           Grilla.ColWidth(2) = 500
           
        Next
        cabecera = "DNI" & vbTab & "TRABAJADOR" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
            c = 2
            NumeroCampo = 2
        
        
        est = "no"
        If est = "no" Then
           estado = Chr(168)
        Else
           estado = Chr(254)
        End If
        
        
          Fila = rstT("cod_unico") & vbTab & rstT("nombre_completo") & vbTab & estado
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
                            
                            
                             If est = "no" Then
                                estado = Chr(168)
                            Else
                                estado = Chr(254)
                            End If
                            
                        End With
        End If
        Fila = ""
        
        
          If est = "si" Then
            For j = 0 To 2
                Grilla.col = j
                Grilla.Row = i + 1
                Grilla.CellBackColor = &HC0FFC0
            Next j
        End If
          rstT.MoveNext
      Next i
    
End Sub
Public Sub llena_productos(ByVal Grilla As MSHFlexGrid, ByVal ruc_empresa As String)
Dim porcentaje As Single
Grilla.SelectionMode = flexSelectionFree
Grilla.MergeCells = flexMergeFree
'strCadena = "select e.`cod_unico`, p.`nombre_completo` from `entidad_empresa` e, persona p " & _
           "where  e.`cod_unico` = p.`dni` and e.`id_empresa` = '20531516045' limit 5"
 
 'strCadena = " select e.`cod_unico`, p.`nombre_completo` from `entidad_empresa` e, persona p " & _
 " where  e.`cod_unico` = p.`dni` and id_empresa='" & KEY_RUC & "' and id_empresa_rel= '" & ruc_empresa & "'"
           
Call ConfiguraRstT(strCadena)
'Call Cargar_FlexGrid(Me.HfActividades, 8, rstT)
porcentaje = 0
If rstT.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
                           
       ' edita la celda
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 4000
           Grilla.ColWidth(2) = 600
           Grilla.ColWidth(3) = 500
        Next
        cabecera = "CODIDO" & vbTab & "DESCRIPCION" & vbTab & "COSTO" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
            c = 3
            NumeroCampo = 3
        
        
        est = "no"
        
           estado = Chr(168)
        
        
          Fila = rstT("id_producto") & vbTab & rstT("nombre_prod") & Space(2) & rstT("color") & vbTab & rstT("precio_compra") & vbTab & estado
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
        End With
        End If
        Fila = ""
        
        
        
          rstT.MoveNext
      Next i
    
End Sub

Public Sub llena_insumos(ByVal id_detalle As Double, ByVal id_linea As String, ByVal Grilla As MSHFlexGrid)
Dim nTotal As Single
strCadena = "SELECT I.id,I.id_producto,P.nombre_prod,I.cantidad,I.precio FROM imp_producto_insumo I,producto P WHERE I.id_producto=P.id_producto and I.ruc=P.ruc and P.ruc='" & KEY_RUC & "' and I.id_producto_detalle='" & id_detalle & "' and I.id_linea='" & Trim(id_linea) & "'"
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
                           
       ' edita la celda
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 700
           Grilla.ColWidth(2) = 3300
           Grilla.ColWidth(3) = 400
           Grilla.ColWidth(4) = 500
        Next
        cabecera = "ID" & vbTab & "CODIGO" & vbTab & "INSUMO" & vbTab & "CANT" & vbTab & "COSTO"
        Grilla.AddItem cabecera
         For k = 1 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
            
        
        
          Fila = rstT("id") & vbTab & rstT("id_producto") & vbTab & rstT("nombre_prod") & vbTab & rstT("cantidad") & vbTab & rstT("precio")
          Grilla.AddItem Fila
          nTotal = nTotal + rstT("precio")
          Fila = ""
          rstT.MoveNext
      Next i
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(nTotal, "#,##0.00")
          Grilla.AddItem Fila
    
End Sub
Public Sub llena_historial_motor(ByVal in_producto As String, ByVal in_chasis As String, ByVal nro_motor As String, ByVal Grilla As MSHFlexGrid)


Dim nTotal As Single
strCadena = "SELECT v.fecha_emision,v.documento,ncliente,v.total FROM movimiento_venta v,movimiento_venta_detalle d WHERE v.anulado='no' and  v.id_venta=d.id_venta and d.id_producto='" & in_producto & "' and   (d.nro_chasis IN('" & in_chasis & "','" & nro_motor & "') OR d.nro_motor IN('" & in_chasis & "','" & nro_motor & "' )) and v.ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
                           
       ' edita la celda
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 2600
           Grilla.ColWidth(2) = 2400
           Grilla.ColWidth(3) = 900
           
        Next
        cabecera = "F.EMISION" & vbTab & "COMPROBANTE" & vbTab & "CLIENTE" & vbTab & "MONTO"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
            
        
        
          Fila = Format(rstT("fecha_emision"), "dd-mm-YYYY") & vbTab & rstT("documento") & vbTab & rstT("ncliente") & vbTab & Format(rstT("total"), "#,##0.00")
          Grilla.AddItem Fila
          
          
          rstT.MoveNext
      Next i
      
    
End Sub
Public Sub llena_historial(ByVal nro_chasis As String, ByVal Grilla As MSHFlexGrid)
Dim nTotal As Single
strCadena = "SELECT I.id,I.id_producto,P.nombre_prod,I.cantidad,I.precio FROM imp_producto_insumo I,producto P WHERE I.id_producto=P.id_producto and I.ruc=P.ruc and P.ruc='" & KEY_RUC & "' and I.id_producto_detalle='" & id_detalle & "' and I.id_linea='" & Trim(id_linea) & "'"
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
                           
       ' edita la celda
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 700
           Grilla.ColWidth(2) = 3300
           Grilla.ColWidth(3) = 400
           Grilla.ColWidth(4) = 500
        Next
        cabecera = "ID" & vbTab & "CODIGO" & vbTab & "INSUMO" & vbTab & "CANT" & vbTab & "COSTO"
        Grilla.AddItem cabecera
         For k = 1 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
            
        
        
          Fila = rstT("id") & vbTab & rstT("id_producto") & vbTab & rstT("nombre_prod") & vbTab & rstT("cantidad") & vbTab & rstT("precio")
          Grilla.AddItem Fila
          nTotal = nTotal + rstT("precio")
          Fila = ""
          rstT.MoveNext
      Next i
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(nTotal, "#,##0.00")
          Grilla.AddItem Fila
    
End Sub
Public Sub llena_insumos_otro(ByVal id_detalle As Double, ByVal id_linea As String, ByVal Grilla As MSHFlexGrid)
Dim nTotal As Single
strCadena = "SELECT I.id,I.id_producto,P.nombre_prod,I.cantidad,I.precio,PP.nombre_completo FROM imp_producto_insumo I,producto P,persona PP WHERE  I.ruc_empresa=PP.dni and  I.id_producto=P.id_producto and I.ruc=P.ruc and P.ruc='" & KEY_RUC & "' and I.id_producto_detalle='" & id_detalle & "' and I.id_linea='" & Trim(id_linea) & "'"
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
                           
       ' edita la celda
       For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 700
           Grilla.ColWidth(2) = 3500
           Grilla.ColWidth(3) = 2400
           Grilla.ColWidth(4) = 500
           Grilla.ColWidth(5) = 800
        Next
        cabecera = "ID" & vbTab & "CODIGO" & vbTab & "INSUMO" & vbTab & "PROVEEDOR" & vbTab & "CANT" & vbTab & "COSTO"
        Grilla.AddItem cabecera
         For k = 1 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
            
        
        
          Fila = rstT("id") & vbTab & rstT("id_producto") & vbTab & rstT("nombre_prod") & vbTab & rstT("nombre_completo") & vbTab & rstT("cantidad") & vbTab & rstT("precio")
          Grilla.AddItem Fila
          nTotal = nTotal + rstT("precio")
          Fila = ""
          rstT.MoveNext
      Next i
      Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(nTotal, "#,##0.00")
          Grilla.AddItem Fila
    
End Sub









Private Sub ActualizarTrabajador(ByVal Grilla As MSHFlexGrid, ByVal id_proc As String)
        
        Dim indRow As Integer
        indRow = Grilla.Row
        
        i = 1
        Do While i < Grilla.Rows
          Grilla.TextMatrix(i, 2) = Chr(168)
          For j = 0 To 2
                Grilla.col = j
                Grilla.Row = i
                Grilla.CellBackColor = &HFFFFFF
          Next j
          
          i = i + 1
          
        Loop
        
        Grilla.Row = indRow
        
        Select Case id_proc
        
        Case "01"
        Me.lblDniTemp.Caption = Grilla.TextMatrix(Grilla.Row, 0)
        Me.lblNombreTemp.Caption = Grilla.TextMatrix(Grilla.Row, 1)
        
        Case "02"
        Me.lblDniTempSol.Caption = Grilla.TextMatrix(Grilla.Row, 0)
        Me.lblNombreTempSol.Caption = Grilla.TextMatrix(Grilla.Row, 1)
                
        Case "03"
        Me.lblDniTempTap.Caption = Grilla.TextMatrix(Grilla.Row, 0)
        Me.lblNombreTempTap.Caption = Grilla.TextMatrix(Grilla.Row, 1)
        
        Case "04"
            Me.lblDniTempEns2.Caption = Grilla.TextMatrix(Grilla.Row, 0)
            Me.lblNombreTempEns2.Caption = Grilla.TextMatrix(Grilla.Row, 1)
        
        Case "05"
                
        End Select
        
        If Grilla.TextMatrix(Grilla.Row, 2) = Chr(254) Then
            Grilla.TextMatrix(Grilla.Row, 2) = Chr(168)
            For j = 0 To 2
                Grilla.col = j
                Grilla.Row = Grilla.Row
                Grilla.CellBackColor = &HFFFFFF
            Next j
        Else
            Grilla.TextMatrix(Grilla.Row, 2) = Chr(254)
            For j = 0 To 2
                Grilla.col = j
                Grilla.Row = Grilla.Row
                Grilla.CellBackColor = &HC0FFC0
            Next j
        End If

      
End Sub
Private Sub ActualizarTrabajador_otro(ByVal Grilla As MSHFlexGrid, ByVal id_proc As String)
        
        Dim indRow As Integer
        indRow = Grilla.Row
        
        i = 1
        Do While i < Grilla.Rows
          Grilla.TextMatrix(i, 2) = Chr(168)
          For j = 0 To 1
                Grilla.col = j
                Grilla.Row = i
                Grilla.CellBackColor = &HFFFFFF
          Next j
          
          i = i + 1
          
        Loop
        
        Grilla.Row = indRow
        
        
        
        If Grilla.TextMatrix(Grilla.Row, 2) = Chr(254) Then
            Grilla.TextMatrix(Grilla.Row, 2) = Chr(168)
            For j = 0 To 1
                Grilla.col = j
                Grilla.Row = Grilla.Row
                Grilla.CellBackColor = &HFFFFFF
            Next j
        Else
            Grilla.TextMatrix(Grilla.Row, 2) = Chr(254)
            For j = 0 To 1
                Grilla.col = j
                Grilla.Row = Grilla.Row
                Grilla.CellBackColor = &HC0FFC0
            Next j
        End If

      
End Sub


Private Sub gridDetalle_Click()
Call seleccion
End Sub
Public Sub seleccion()
  limpiarDatos
  fraDetalle.Visible = True
  
  If Me.gridDetalle.Row > 0 Then
  strCadena = "select DATE_FORMAT(m.`fecha_entrada`,'%d/%m/%Y')  as fecha_entrada,  DATE_FORMAT(m.`fecha_salida`,'%d/%m/%Y') as fecha_salida, DATE_FORMAT(m.`hora_entrada`, '%H:%i') as hora_entrada , " & _
" DATE_FORMAT(m.`hora_salida`, '%H:%i')  as hora_salida , p.`nombre_completo` as trabajador , e.`nombre_completo` as empresa,  " & _
" m.`id_proceso`, m.`estado`, m.`observacion` " & _
" from `imp_producto_detalle` i , `imp_producto_movimiento` m, persona e, persona p " & _
" where i.`id_detalle` = m.`id_detalle` and i.`id_detalle` = " & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0) & " and " & _
" m.`ruc_empresa` = e.`dni` and p.`dni` = m.`id_autor`"
     
      Call ConfiguraRstT(strCadena)
      
      Call ocultarBotones
      
      Dim tipoProceso As Integer
      
      tipoProceso = "01"
      
      If tipoProceso = "02" Then
          
          fraDetalle.Visible = False
          fraEnsamblaje.Visible = True
          
          If rstT.RecordCount = 0 Then
               Me.cmdIniEns2.Visible = True
          End If
          
          
          For i = 1 To rstT.RecordCount
          
          Select Case rstT("id_proceso")
         
           Case "02"
             lblEns2FechaIni.Caption = rstT("fecha_entrada")
             lblEns2HoraIni.Caption = rstT("hora_entrada")
             
             If Not IsNull(rstT("fecha_salida")) Then
              lblEns2FechaFin.Caption = rstT("fecha_salida")
             End If
             
             If Not IsNull(rstT("hora_salida")) Then
              lblEns2HoraFin.Caption = rstT("hora_salida")
             End If
             
             lblEns2Empresa.Caption = rstT("empresa")
             lblEns2Trabajador.Caption = rstT("trabajador")
             lblEns2Observacion.Caption = rstT("observacion")
             
             Call evaluarBotones_v2(rstT("id_proceso"), rstT("estado"))
             
            End Select
            
            rstT.MoveNext
            
            Next i
            
            
      
          Exit Sub
      End If
      
      
      
      
      limpiarCambioProceso
      
      If rstT.RecordCount = 0 Then
         Me.cmdCambiar.Visible = True
         Me.cmdIniEns.Visible = True

         Else
         
         Me.cmdCambiar.Visible = False
      End If
      

      For i = 1 To rstT.RecordCount

         Select Case rstT("id_proceso")
          Case "01"
             lblEmFechaIni.Caption = rstT("fecha_entrada")
             lblEmHoraIni.Caption = rstT("hora_entrada")
             If Not IsNull(rstT("fecha_salida")) Then
              lblEmFechaFin.Caption = rstT("fecha_salida")
             End If
             If Not IsNull(rstT("hora_salida")) Then
             lblEmHoraFin.Caption = rstT("hora_salida")
             End If
             lblEmEmpresa.Caption = rstT("empresa")
             lblEmTrabajador.Caption = rstT("trabajador")
             lblEmObservacion.Caption = rstT("observacion")
             
             Call evaluarBotones(rstT("id_proceso"), rstT("estado"))
          
          Case "02"
             lblSoFechaIni.Caption = rstT("fecha_entrada")
             lblSoHoraIni.Caption = rstT("hora_entrada")
             If Not IsNull(rstT("fecha_salida")) Then
              lblSoFechaFin.Caption = rstT("fecha_salida")
             End If
             If Not IsNull(rstT("hora_salida")) Then
             lblSoHoraFin.Caption = rstT("hora_salida")
             End If
             lblSoEmpresa.Caption = rstT("empresa")
             lblSoTrabajador.Caption = rstT("trabajador")
             lblSoObservacion.Caption = rstT("observacion")
             
             Call evaluarBotones(rstT("id_proceso"), rstT("estado"))
          
          Case "03"
             lblTaFechaIni.Caption = rstT("fecha_entrada")
             lblTaHoraIni.Caption = rstT("hora_entrada")
             If Not IsNull(rstT("fecha_salida")) Then
              lblTaFechaFin.Caption = rstT("fecha_salida")
             End If
             If Not IsNull(rstT("hora_salida")) Then
             lblTaHoraFin.Caption = rstT("hora_salida")
             End If
             lblTaEmpresa.Caption = rstT("empresa")
             lblTaTrabajador.Caption = rstT("trabajador")
             lblTaObservacion.Caption = rstT("observacion")
             
             Call evaluarBotones(rstT("id_proceso"), rstT("estado"))
         
         End Select
         
        rstT.MoveNext
            
      Next i
       
      
       
    End If

End Sub
Private Sub limpiarDatos()
             lblEmFechaIni.Caption = "-"
             lblEmHoraIni.Caption = "-"
             
              lblEmFechaFin.Caption = "-"
             
            
             lblEmHoraFin.Caption = "-"
             
             lblEmEmpresa.Caption = "-"
             lblEmTrabajador.Caption = "-"
             lblEmObservacion.Caption = "-"
          
         
             lblSoFechaIni.Caption = "-"
             lblSoHoraIni.Caption = "-"
             
             lblSoFechaFin.Caption = "-"
             
             
             lblSoHoraFin.Caption = "-"
             
             lblSoEmpresa.Caption = "-"
             lblSoTrabajador.Caption = "-"
             lblSoObservacion.Caption = "-"
          
          
             lblTaFechaIni.Caption = "-"
             lblTaHoraIni.Caption = "-"
             
             lblTaFechaFin.Caption = "-"
             
             lblTaHoraFin.Caption = "-"
             
             lblTaEmpresa.Caption = "-"
             lblTaTrabajador.Caption = "-"
             lblTaObservacion.Caption = "-"

End Sub


Private Sub gridDetalle_SelChange()
Call seleccion
End Sub

Private Sub gridInsumoSol_Click()

End Sub



Private Sub gridTrabajador_Click()
   If gridTrabajador.Row > 0 Then
     Call ActualizarTrabajador(gridTrabajador, "01")
   End If
   
   
           
End Sub

Private Sub ocultarBotones()
             Me.cmdFinEns.Visible = False
             Me.cmdIniEns.Visible = False
             Me.cmdIniSol.Visible = False
             Me.cmdFinSol.Visible = False
             Me.cmdIniTap.Visible = False
             Me.cmdFinTap.Visible = False
             
             Me.cmdIniEns2.Visible = False
             Me.cmdFinEns2.Visible = False
End Sub

Private Sub mostrarBotones()
             Me.cmdFinEns.Visible = True
             Me.cmdIniEns.Visible = True
             Me.cmdIniSol.Visible = True
             Me.cmdFinSol.Visible = True
             Me.cmdIniTap.Visible = True
             Me.cmdFinTap.Visible = True
End Sub




Private Sub gridTrabajadorEns2_Click()
   If gridTrabajadorEns2.Row > 0 Then
     Call ActualizarTrabajador(gridTrabajadorEns2, "04")
   End If
End Sub

Private Sub gridTrabajadorSol_Click()
   If gridTrabajadorSol.Row > 0 And Trim(Me.gridTrabajadorSol.TextMatrix(Me.gridTrabajadorSol.Row, 0)) <> "" Then
     Call ActualizarTrabajador(gridTrabajadorSol, "02")
   End If
   
   
End Sub

Private Sub gridTrabajadorTap_Click()
   If gridTrabajadorTap.Row > 0 Then
     Call ActualizarTrabajador(gridTrabajadorTap, "03")
   End If
End Sub


Private Sub actualizarProcesos()
  limpiarDatos
  limpiarCambioProceso
  
  Me.fraEnsamblado.Visible = False
  Me.fraSol.Visible = False
  Me.fraTap.Visible = False
  
  Me.fraEnsambladoFin.Visible = False
  Me.fraSolFin.Visible = False
  Me.fraTapFin.Visible = False
  
  Me.fraEns2.Visible = False
  Me.fraEns2Fin.Visible = False
  
  fraDetalle.Visible = True
  
  
  If Me.gridDetalle.Row > 0 Then
  
  strCadena = "select DATE_FORMAT(m.`fecha_entrada`,'%d/%m/%Y')  as fecha_entrada,  DATE_FORMAT(m.`fecha_salida`,'%d/%m/%Y') as fecha_salida, DATE_FORMAT(m.`hora_entrada`, '%H:%i') as hora_entrada , " & _
" DATE_FORMAT(m.`hora_salida`, '%H:%i')  as hora_salida , p.`nombre_completo` as trabajador , e.`nombre_completo` as empresa,  " & _
" m.`id_proceso`, m.`estado`, m.`observacion` " & _
" from `imp_producto_detalle` i , `imp_producto_movimiento` m, persona e, persona p " & _
" where i.`id_detalle` = m.`id_detalle` and i.`id_detalle` = " & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0) & " and " & _
" m.`ruc_empresa` = e.`dni` and p.`dni` = m.`id_autor` ORDER BY m.id_mov DESC LIMIT 1"
     
      Call ConfiguraRstT(strCadena)
      
       ocultarBotones
       
       
       
       Dim tipoProceso As String
      
      tipoProceso = "01"
      
      
      If tipoProceso = "02" Then
          
          
          fraDetalle.Visible = False
          fraEnsamblaje.Visible = True
          
          If rstT.RecordCount = 0 Then
           Me.cmdIniEns2.Visible = True
          End If
          
          
          For i = 1 To rstT.RecordCount
          
          Select Case rstT("id_proceso")
         
           Case "02"
             lblEns2FechaIni.Caption = rstT("fecha_entrada")
             lblEns2HoraIni.Caption = rstT("hora_entrada")
             If Not IsNull(rstT("fecha_salida")) Then
              lblEns2FechaFin.Caption = rstT("fecha_salida")
             End If
             If Not IsNull(rstT("hora_salida")) Then
              lblEns2HoraFin.Caption = rstT("hora_salida")
             End If
             lblEns2Empresa.Caption = rstT("empresa")
             lblEns2Trabajador.Caption = rstT("trabajador")
             lblEns2Observacion.Caption = rstT("observacion")
             
             Call evaluarBotones_v2(rstT("id_proceso"), rstT("estado"))
             
            End Select
            
            rstT.MoveNext
            
            Next i
            
            
      
          Exit Sub
      End If
       
       
       
       
       
       
       
       
       
       
       
       limpiarCambioProceso
      
       If rstT.RecordCount = 0 Then
         Me.cmdCambiar.Visible = True
         Me.cmdIniEns.Visible = True

         Else
         
         Me.cmdCambiar.Visible = False
       End If
       
       For i = 1 To rstT.RecordCount

         Select Case rstT("id_proceso")
          Case "01"
             lblEmFechaIni.Caption = rstT("fecha_entrada")
             lblEmHoraIni.Caption = rstT("hora_entrada")
             If Not IsNull(rstT("fecha_salida")) Then
              lblEmFechaFin.Caption = rstT("fecha_salida")
             End If
             If Not IsNull(rstT("hora_salida")) Then
             lblEmHoraFin.Caption = rstT("hora_salida")
             End If
             lblEmEmpresa.Caption = rstT("empresa")
             lblEmTrabajador.Caption = rstT("trabajador")
             lblEmObservacion.Caption = rstT("observacion")
             
             Call evaluarBotones(rstT("id_proceso"), rstT("estado"))
          
          Case "02"
             lblSoFechaIni.Caption = rstT("fecha_entrada")
             lblSoHoraIni.Caption = rstT("hora_entrada")
             If Not IsNull(rstT("fecha_salida")) Then
              lblSoFechaFin.Caption = rstT("fecha_salida")
             End If
             If Not IsNull(rstT("hora_salida")) Then
             lblSoHoraFin.Caption = rstT("hora_salida")
             End If
             lblSoEmpresa.Caption = rstT("empresa")
             lblSoTrabajador.Caption = rstT("trabajador")
             lblSoObservacion.Caption = rstT("observacion")
             
             Call evaluarBotones(rstT("id_proceso"), rstT("estado"))
          
          Case "03"
             lblTaFechaIni.Caption = rstT("fecha_entrada")
             lblTaHoraIni.Caption = rstT("hora_entrada")
             If Not IsNull(rstT("fecha_salida")) Then
              lblTaFechaFin.Caption = rstT("fecha_salida")
             End If
             If Not IsNull(rstT("hora_salida")) Then
             lblTaHoraFin.Caption = rstT("hora_salida")
             End If
             lblTaEmpresa.Caption = rstT("empresa")
             lblTaTrabajador.Caption = rstT("trabajador")
             lblTaObservacion.Caption = rstT("observacion")
             
             Call evaluarBotones(rstT("id_proceso"), rstT("estado"))
         
         End Select
         
        rstT.MoveNext
            
      Next i
       
    End If
     
End Sub


Public Sub evaluarBotones(ByVal id_proceso As String, ByVal id_estado As String)
       ocultarBotones
       
       Select Case id_proceso
            Case "01"
               If id_estado = "0" Then
                 cmdFinEns.Visible = True
               Else
                 cmdIniSol.Visible = True
                 Me.cmdreiniciarSoldadura.Visible = True
               End If
            
            Case "02"
               If id_estado = "0" Then
                 Me.cmdFinSol.Visible = True
               Else
                 cmdIniTap.Visible = True
                 Me.cmdreiniciarensamblaje.Visible = True
               End If
            
            Case "03"
               If id_estado = "0" Then
                 Me.cmdFinTap.Visible = True
                Else
                    Me.cmdreiniciartapiz.Visible = True
               End If
               
       End Select
End Sub
Public Sub evaluarBotonesReiniciar(ByVal id_proceso As String, ByVal id_estado As String)
       ocultarBotones
       
       Select Case id_proceso
            Case "01"
               If id_estado = "0" Then
                  cmdFinEns.Visible = True
                  Me.cmdIniEns.Visible = True
                  Me.cmdreiniciarSoldadura.Visible = False
              
               End If
            
            Case "02"
               If id_estado = "0" Then
                 Me.cmdFinSol.Visible = True
                 Me.cmdIniSol.Visible = True
                 Me.cmdreiniciarensamblaje.Visible = False
               End If
            
            Case "03"
               If id_estado = "0" Then
                 Me.cmdFinTap.Visible = True
                 Me.cmdIniTap.Visible = True
                 Me.cmdreiniciartapiz.Visible = False
               End If
               
       End Select
End Sub



Private Sub evaluarBotones_v2(ByVal id_proceso As String, ByVal id_estado As String)
       Call ocultarBotones
       
       Select Case id_proceso
                       
            Case "02"
               If id_estado = "0" Then
                 Me.cmdFinEns2.Visible = True
               End If
               
       End Select
End Sub

Private Sub llenarEstados(ByVal cbo As DataCombo)
   
    strCadena = "select e.id_estado as Codigo, e.`descripcion` as Descripcion from imp_estado e"
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(cbo)
    
End Sub

Private Sub llenarProcesos(ByVal cbo As DataCombo)
   
    strCadena = "select e.id_produccion as Codigo, e.`descripcion` as Descripcion from linea_produccion e"
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(cbo)
    
End Sub


Private Sub cmdTerminar_Click()
     intResponse = MsgBox("¿Está seguro que desea finalizar producción de producto?", vbYesNo + vbQuestion, "Confirmar")
  
     If intResponse = vbYes Then
      strCadena = " update imp_producto_detalle set id_estado = '03',dni_save='" & KEY_USUARIO & "',nsave='" & KEY_VENDEDOR & "' where id_detalle =" & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0)
      CnBd.Execute (strCadena)
       
      
      'strCadena = "SELECT * FROM imp_producto_detalle I,movimiento_compra_detalle D WHERE I.id_detalle_compra=D.id_detalle_compra AND  I.id_detalle='" & Val(gridDetalle.TextMatrix(Me.gridDetalle.Row, 0)) & "' "
      'Call ConfiguraRst(strCadena)
      'If rst.RecordCount > 0 Then
      ' strCadena = "UPDATE kardex SET cantidad_contable='0',cantidad_real='1' WHERE id_movimiento='" & rst("id_compra") & "' AND id_producto='" & rst("id_producto") & "' AND ruc='" & KEY_RUC & "'"
      'CnBd.Execute (strCadena)
      'End If
      
      Call actualizaGrid(Me.gridDetalle, "03", Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0))
      
     End If
End Sub



Private Sub txtSerie_KeyPress(KeyAscii As Integer)
End Sub

Private Sub HfInsumoTercero_Click()
Call agregar_insumo_otro
End Sub

Private Sub HfListadoInsumos_SelChange()
If Val(Me.HfListadoInsumos.TextMatrix(Me.HfListadoInsumos.Row, 0)) > 0 Then
    Me.cmdquitarTercero.Visible = True
Else
    Me.cmdquitarTercero.Visible = False
End If
End Sub

Private Sub HgTrabajadorOtro_Click()
 If HfInsumoTercero.Row > 0 Then
     Call ActualizarTrabajador_otro(HgTrabajadorOtro, "05")
 End If

End Sub

Private Sub txtBuscarInsumo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT a.id_producto,p.nombre_prod,c.descripcion as color,a.precio_compra FROM producto p,almacen_producto a,imp_color c WHERE  p.id_color=c.id_color and   p.id_producto=a.id_producto and a.id_alm='" & KEY_ALM & "' and id_proveedor='" & Me.DtcEmpresaTercero.BoundText & "' and  a.ruc=p.ruc and  a.ruc='" & KEY_RUC & "' and (p.nombre_prod LIKE '%" & Trim(Me.txtBuscarInsumo.Text) & "%' or c.descripcion LIKE '%" & Trim(Me.txtBuscarInsumo.Text) & "%')"
    Call llena_productos(HfInsumoTercero, Me.DtcEmpresaTercero.BoundText)
End If
End Sub
Private Sub agregar_insumo()
        HfInsumoTercero.TextMatrix(HfInsumoTercero.Row, 3) = Chr(254)
        For j = 0 To 2
                HfInsumoTercero.col = j
                HfInsumoTercero.Row = HfInsumoTercero.Row
                HfInsumoTercero.CellBackColor = &HC0FFC0
        Next j

    strCadena = "insert into imp_producto_movimiento (id_detalle,id_proceso,fecha_entrada,hora_entrada,estado, observacion, id_autor, ruc_empresa, ruc) " & _
               " values (" & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0) & ",'05', '" & KEY_FECHA & "', '" & Format(Now(), "HH:mm:ss") & "', '0','-','" & Me.HgTrabajadorOtro.TextMatrix(Me.HgTrabajadorOtro.Row, 0) & "','" & Me.DtcEmpresaTercero.BoundText & "','" & KEY_RUC & "')"
    CnBd.Execute strCadena
               
        
        strCadena = "INSERT INTO imp_producto_insumo(id_producto_detalle,id_producto,precio,id_linea,dni_autor,ruc_empresa,ruc)VALUES('" & gridDetalle.TextMatrix(gridDetalle.Row, 0) & "','" & Me.HfInsumoTercero.TextMatrix(Me.HfInsumoTercero.Row, 0) & "','" & Me.HfInsumoTercero.TextMatrix(Me.HfInsumoTercero.Row, 2) & "','" & Txtid_estado.Text & "','" & Me.HgTrabajadorOtro.TextMatrix(Me.HgTrabajadorOtro.Row, 0) & "','" & Me.DtcEmpresaTercero.BoundText & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        
        Call frmCorProcesos.llena_insumos(gridDetalle.TextMatrix(gridDetalle.Row, 0), Txtid_estado.Text, Me.HfListadoInsumos)
        frmCorProcesos.Procedencia = Neutro
        Exit Sub
End Sub
Private Sub agregar_insumo_otro()
        HfInsumoTercero.TextMatrix(HfInsumoTercero.Row, 3) = Chr(254)
        For j = 0 To 2
                HfInsumoTercero.col = j
                HfInsumoTercero.Row = HfInsumoTercero.Row
                HfInsumoTercero.CellBackColor = &HC0FFC0
        Next j

    strCadena = "insert into imp_producto_movimiento (id_detalle,id_proceso,fecha_entrada,hora_entrada,estado, observacion, id_autor, ruc_empresa, ruc) " & _
               " values (" & Me.gridDetalle.TextMatrix(Me.gridDetalle.Row, 0) & ",'05', '" & KEY_FECHA & "', '" & Format(Now(), "HH:mm:ss") & "', '0','-','" & Me.HgTrabajadorOtro.TextMatrix(Me.HgTrabajadorOtro.Row, 0) & "','" & Me.DtcEmpresaTercero.BoundText & "','" & KEY_RUC & "')"
    CnBd.Execute strCadena
               
        
        strCadena = "INSERT INTO imp_producto_insumo(id_producto_detalle,id_producto,precio,id_linea,dni_autor,ruc_empresa,ruc)VALUES('" & gridDetalle.TextMatrix(gridDetalle.Row, 0) & "','" & Me.HfInsumoTercero.TextMatrix(Me.HfInsumoTercero.Row, 0) & "','" & Me.HfInsumoTercero.TextMatrix(Me.HfInsumoTercero.Row, 2) & "','" & Txtid_estado.Text & "','" & Me.HgTrabajadorOtro.TextMatrix(Me.HgTrabajadorOtro.Row, 0) & "','" & Me.DtcEmpresaTercero.BoundText & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        
        Call frmCorProcesos.llena_insumos_otro(gridDetalle.TextMatrix(gridDetalle.Row, 0), Txtid_estado.Text, Me.HfListadoInsumos)
        frmCorProcesos.Procedencia = Neutro
        Exit Sub
End Sub

Private Sub txtModelo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim clausula As String
    clausula = ""
    
    
    If txtModelo.Text = "" Then
        clausula = ""
    End If
 

   strCadena = " select d.`id_detalle`, d.`id_detalle_compra` , p.`id_producto`, " & _
  " p.`nombre_prod` as producto, d.`serie`, d.`anio_fabricacion`, " & _
  " d.`anio_contenedor`, d.`anio_modelo`, d.`nro_chasis`, d.`nro_contenedor`, " & _
  " d.`nro_motor`, d.`item`, e.`descripcion` as estado, e.`id_estado`,cc.descripcion as color,m.descripcion as marca,a.descripcion as almacen from `imp_producto_detalle` d, movimiento_compra_detalle c, " & _
  " producto p , imp_estado e,imp_color cc,marca m ,almacen a where a.id_alm=d.id_alm and a.ruc=d.ruc and   m.id_marca=p.id_marca and m.id_usu=p.ruc and   p.id_color=cc.id_color and   d.`id_detalle_compra` = c.`id_detalle_compra` and " & _
  " c.`id_producto` = p.`id_producto` and " & _
  " p.`ruc` = '" & KEY_RUC & "' and  " & _
  " e.`id_estado` = d.`id_estado` and d.id_alm='" & Me.DtcAlmacen.BoundText & "' "
    Call diagnostico(Me.gridDetalle, "1")
End If
End Sub

Private Sub txtMotor_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      strCadena = " select d.`id_detalle`, d.`id_detalle_compra` , p.`id_producto`, " & _
  " p.`nombre_prod` as producto, d.`serie`, d.`anio_fabricacion`, " & _
  " d.`anio_contenedor`, d.`anio_modelo`, d.`nro_chasis`, d.`nro_contenedor`, " & _
  " d.`nro_motor`, d.`item`, e.`descripcion` as estado, e.`id_estado`,cc.descripcion as color,m.descripcion as marca,a.descripcion as almacen from `imp_producto_detalle` d, movimiento_compra_detalle c, " & _
  " producto p , imp_estado e,imp_color cc,marca m,almacen a  where d.id_alm=a.id_alm and a.ruc=d.ruc and   m.id_marca=p.id_marca and m.id_usu=p.ruc  and   p.id_color=cc.id_color and   d.`id_detalle_compra` = c.`id_detalle_compra`  and " & _
  " c.`ruc`=p.`ruc` and c.`id_producto` = p.`id_producto` and " & _
  " p.`ruc` = '" & KEY_RUC & "' and " & _
  " e.`id_estado` = d.`id_estado` and d.id_alm='" & Me.DtcAlmacen.BoundText & "' "
  
      Call diagnostico(Me.gridDetalle, "2")
  End If
End Sub

Private Sub txtChasis_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      strCadena = " select d.`id_detalle`, d.`id_detalle_compra` , p.`id_producto`, " & _
  " p.`nombre_prod` as producto, d.`serie`, d.`anio_fabricacion`, " & _
  " d.`anio_contenedor`, d.`anio_modelo`, d.`nro_chasis`, d.`nro_contenedor`, " & _
  " d.`nro_motor`, d.`item`, e.`descripcion` as estado, e.`id_estado`,cc.descripcion as color,m.descripcion as marca,a.descripcion as almacen from `imp_producto_detalle` d, movimiento_compra_detalle c, " & _
  " producto p , imp_estado e,imp_color cc ,marca m,almacen a where d.id_alm=a.id_alm and d.ruc=a.ruc and   m.id_marca=p.id_marca and m.id_usu=p.ruc and    p.id_color=cc.id_color and   d.`id_detalle_compra` = c.`id_detalle_compra`  and  " & _
  " c.`ruc`=p.`ruc` and c.`id_producto` = p.`id_producto` and " & _
  " p.`ruc` = '" & KEY_RUC & "' and " & _
  " e.`id_estado` = d.`id_estado` and d.id_alm='" & Me.DtcAlmacen.BoundText & "' "
  
      Call diagnostico(Me.gridDetalle, "3")
  End If
End Sub

Private Sub limpiarCambioProceso()
  Me.lblProcesoCambiado.Caption = "-"
  Me.fraCambiar.Visible = False
End Sub

Private Sub ocultarFrames()
    fraEnsamblado.Visible = False
    fraEnsambladoFin.Visible = False
    fraSol.Visible = False
    fraSolFin.Visible = False
    fraTap.Visible = False
    fraTapFin.Visible = False
End Sub


Private Sub txtProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT id_producto as Codigo, nombre_prod as Descripcion FROM producto  WHERE nombre_prod LIKE '%" & Trim(Me.txtproducto.Text) & "%' and   ruc='" & KEY_RUC & "' LIMIT 20"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcProducto)
End If
End Sub

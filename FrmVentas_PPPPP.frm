VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmVentas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Seccion Ventas"
   ClientHeight    =   8925
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   18990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   18990
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtserial 
      Height          =   375
      Left            =   10800
      TabIndex        =   200
      Top             =   7560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fraApp 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   6480
      TabIndex        =   112
      Top             =   2280
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txtrecibo_anterior 
         Height          =   285
         Left            =   4200
         TabIndex        =   148
         Text            =   "0"
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
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
         TabIndex        =   147
         Top             =   1920
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
         TabIndex        =   129
         Top             =   960
         Width           =   1575
      End
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
         TabIndex        =   128
         Top             =   480
         Width           =   1575
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfRecibos 
         Height          =   1935
         Left            =   120
         TabIndex        =   146
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
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
   Begin VB.Frame frmcredito 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4800
      TabIndex        =   196
      Top             =   1800
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox txtmontocredito 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   199
         Top             =   120
         Width           =   1335
      End
      Begin VitekeySoft.ChameleonBtn cmdactualizarcredito 
         Height          =   285
         Left            =   1680
         TabIndex        =   198
         Top             =   480
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "ACTUALIZAR"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas_.frx":0000
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
         TabIndex        =   197
         Top             =   120
         Width           =   1395
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdcredito 
      Height          =   405
      Left            =   4800
      TabIndex        =   195
      Top             =   1360
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmVentas_.frx":001C
      PICN            =   "FrmVentas_.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frm_editable 
      BackColor       =   &H00FFFFFF&
      Height          =   3045
      Left            =   120
      TabIndex        =   150
      Top             =   3600
      Visible         =   0   'False
      Width           =   13215
      Begin VitekeySoft.ChameleonBtn cmdcerraredicion 
         Height          =   300
         Left            =   10800
         TabIndex        =   191
         Top             =   2520
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "CERRAR EDICION"
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
         MICON           =   "FrmVentas_.frx":31AE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox TxtCodProducto_per 
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
         Index           =   4
         Left            =   120
         TabIndex        =   188
         Top             =   1920
         Width           =   1185
      End
      Begin VB.TextBox TxtCodProducto_per 
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
         Index           =   3
         Left            =   120
         TabIndex        =   187
         Top             =   1560
         Width           =   1185
      End
      Begin VB.TextBox TxtCodProducto_per 
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
         Index           =   2
         Left            =   120
         TabIndex        =   186
         Top             =   1200
         Width           =   1185
      End
      Begin VB.TextBox TxtCodProducto_per 
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
         Index           =   1
         Left            =   120
         TabIndex        =   185
         Top             =   840
         Width           =   1185
      End
      Begin VB.TextBox TxtCodProducto_per 
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
         Index           =   0
         Left            =   120
         TabIndex        =   184
         Top             =   480
         Width           =   1185
      End
      Begin VB.TextBox txtCantidadPer 
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
         Index           =   4
         Left            =   1440
         TabIndex        =   182
         Top             =   1920
         Width           =   825
      End
      Begin VB.TextBox txtCantidadPer 
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
         Index           =   3
         Left            =   1440
         TabIndex        =   181
         Top             =   1560
         Width           =   825
      End
      Begin VB.TextBox txtCantidadPer 
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
         Index           =   2
         Left            =   1440
         TabIndex        =   180
         Top             =   1200
         Width           =   825
      End
      Begin VB.TextBox txtCantidadPer 
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
         Index           =   1
         Left            =   1440
         TabIndex        =   179
         Top             =   840
         Width           =   825
      End
      Begin VB.TextBox txtCantidadPer 
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
         Index           =   0
         Left            =   1440
         TabIndex        =   178
         Top             =   480
         Width           =   825
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
         Left            =   2400
         TabIndex        =   170
         Top             =   480
         Width           =   1215
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
         Left            =   2400
         TabIndex        =   169
         Top             =   840
         Width           =   1215
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
         Left            =   2400
         TabIndex        =   168
         Top             =   1200
         Width           =   1215
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
         Left            =   2400
         TabIndex        =   167
         Top             =   1560
         Width           =   1215
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
         Left            =   2400
         TabIndex        =   166
         Top             =   1920
         Width           =   1215
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
         TabIndex        =   165
         Top             =   480
         Width           =   6255
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
         TabIndex        =   164
         Top             =   840
         Width           =   6255
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
         TabIndex        =   163
         Top             =   1200
         Width           =   6255
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
         TabIndex        =   162
         Top             =   1560
         Width           =   6255
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
         TabIndex        =   161
         Top             =   1920
         Width           =   6255
      End
      Begin VB.TextBox txtprecio_per 
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
         Left            =   10080
         TabIndex        =   160
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtprecio_per 
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
         Left            =   10080
         TabIndex        =   159
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtprecio_per 
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
         Left            =   10080
         TabIndex        =   158
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtprecio_per 
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
         Left            =   10080
         TabIndex        =   157
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtprecio_per 
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
         Left            =   10080
         TabIndex        =   156
         Top             =   1920
         Width           =   1215
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
         Left            =   11400
         TabIndex        =   155
         Top             =   480
         Width           =   1215
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
         Left            =   11400
         TabIndex        =   154
         Top             =   840
         Width           =   1215
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
         Left            =   11400
         TabIndex        =   153
         Top             =   1200
         Width           =   1215
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
         Left            =   11400
         TabIndex        =   152
         Top             =   1560
         Width           =   1215
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
         Left            =   11400
         TabIndex        =   151
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANT"
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
         Left            =   1440
         TabIndex        =   176
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label28 
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
         Left            =   2400
         TabIndex        =   175
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label27 
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
         Left            =   5100
         TabIndex        =   174
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P.UNIT"
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
         Left            =   10080
         TabIndex        =   173
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMP TOTAL"
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
         Left            =   11400
         TabIndex        =   172
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label23 
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
         Left            =   120
         TabIndex        =   171
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtafectacaja 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10440
      TabIndex        =   192
      Text            =   "no"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txteditable 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   190
      Text            =   "no"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
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
      Left            =   8700
      MaxLength       =   80
      TabIndex        =   189
      Top             =   6795
      Width           =   1095
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
      Height          =   285
      Left            =   240
      TabIndex        =   183
      Top             =   6795
      Width           =   1425
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
      Height          =   285
      Left            =   1750
      TabIndex        =   177
      Top             =   6795
      Width           =   700
   End
   Begin VB.Frame FrameReferencia 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1100
      Left            =   4200
      TabIndex        =   142
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
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
         TabIndex        =   145
         Top             =   645
         Visible         =   0   'False
         Width           =   150
      End
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
         Height          =   315
         Left            =   240
         MaxLength       =   80
         TabIndex        =   144
         Top             =   645
         Width           =   2655
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COMPROBANTE REFERENCIA"
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
         Left            =   240
         TabIndex        =   143
         Top             =   180
         Width           =   2640
      End
   End
   Begin VB.TextBox txttipofactura 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   139
      Text            =   "00001"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   13440
      TabIndex        =   138
      Top             =   8760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkconyuge 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "CONYUGE"
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
      Height          =   285
      Left            =   3600
      TabIndex        =   137
      Top             =   1440
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
      Height          =   285
      Left            =   6240
      MaxLength       =   80
      TabIndex        =   136
      Top             =   2520
      Width           =   855
   End
   Begin VB.CheckBox chkVincular 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "VINCULAR"
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
      Left            =   11760
      TabIndex        =   127
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
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
      TabIndex        =   126
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame frameTramite 
      BackColor       =   &H00FFFFFF&
      Height          =   2480
      Left            =   14520
      TabIndex        =   108
      Top             =   300
      Visible         =   0   'False
      Width           =   4455
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
         Left            =   2400
         TabIndex        =   121
         Top             =   1980
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   1815
         TabIndex        =   120
         Top             =   1980
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.TextBox txtNumeroRecibo 
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
         Left            =   2400
         TabIndex        =   119
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtSerieRecibo 
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
         Left            =   1815
         TabIndex        =   118
         Top             =   1440
         Visible         =   0   'False
         Width           =   550
      End
      Begin VitekeySoft.ChameleonBtn cmdConstancia 
         Height          =   350
         Left            =   360
         TabIndex        =   109
         Top             =   165
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "  CONSTANCIA                                             "
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas_.frx":31CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdDeclaracion 
         Height          =   350
         Left            =   360
         TabIndex        =   110
         Top             =   570
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   " DECLARACION JURADA                              "
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas_.frx":31E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdSolicitud 
         Height          =   345
         Left            =   360
         TabIndex        =   111
         Top             =   975
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "SOLICITUD AP"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas_.frx":3202
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdGuiaRemision 
         Height          =   350
         Left            =   360
         TabIndex        =   116
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "G.REMISION"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas_.frx":321E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdRecibo 
         Height          =   350
         Left            =   360
         TabIndex        =   117
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "RECIBO       "
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas_.frx":323A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdImprimirRecibo 
         Height          =   300
         Left            =   3960
         TabIndex        =   122
         Top             =   1440
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
         MICON           =   "FrmVentas_.frx":3256
         PICN            =   "FrmVentas_.frx":3272
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdImprimirGuia 
         Height          =   300
         Left            =   3960
         TabIndex        =   123
         Top             =   1980
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
         MICON           =   "FrmVentas_.frx":32FF
         PICN            =   "FrmVentas_.frx":331B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdGrabarRecibo 
         Height          =   300
         Left            =   3540
         TabIndex        =   124
         Top             =   1440
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
         MICON           =   "FrmVentas_.frx":33A8
         PICN            =   "FrmVentas_.frx":33C4
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
         Left            =   3540
         TabIndex        =   125
         Top             =   1980
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
         MICON           =   "FrmVentas_.frx":395E
         PICN            =   "FrmVentas_.frx":397A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdgenerarmantenimientos 
         Height          =   345
         Left            =   2160
         TabIndex        =   194
         Top             =   975
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   " MANTENIMIENTOS"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   8388608
         FCOLO           =   8388608
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas_.frx":3F14
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdActivar 
      Height          =   520
      Left            =   6000
      TabIndex        =   105
      Top             =   195
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   926
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmVentas_.frx":3F30
      PICN            =   "FrmVentas_.frx":3F4C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Left            =   13440
      MaxLength       =   80
      TabIndex        =   104
      ToolTipText     =   "TELEFONO"
      Top             =   8400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VitekeySoft.ChameleonBtn CmdAgregar 
      Height          =   345
      Left            =   11520
      TabIndex        =   101
      ToolTipText     =   "AGREGAR ITEM"
      Top             =   6770
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
      MICON           =   "FrmVentas_.frx":6836
      PICN            =   "FrmVentas_.frx":6852
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtTipoCambio 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   11760
      MaxLength       =   80
      TabIndex        =   77
      Top             =   1320
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
      Left            =   13440
      MaxLength       =   80
      TabIndex        =   76
      ToolTipText     =   "TELEFONO"
      Top             =   8040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox sendmail1 
      Height          =   480
      Left            =   13440
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   70
      Top             =   9360
      Width           =   480
   End
   Begin VB.CommandButton CmdVisualizar 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Command1"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   7450
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox TxtIgv 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12000
      TabIndex        =   68
      Text            =   "si"
      Top             =   7440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame PanelCredito 
      Height          =   735
      Left            =   11040
      TabIndex        =   63
      Top             =   1245
      Visible         =   0   'False
      Width           =   2175
      Begin VB.TextBox TxtCuotas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   840
         MaxLength       =   80
         TabIndex        =   65
         Top             =   420
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtClaveRandonCredito 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   840
         MaxLength       =   80
         TabIndex        =   64
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUOTAS:"
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
         Left            =   90
         TabIndex        =   67
         Top             =   480
         Width           =   705
      End
      Begin VB.Label lblclave 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLAVE:"
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
         Left            =   90
         TabIndex        =   66
         Top             =   120
         Width           =   555
      End
   End
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
      Left            =   12000
      MaxLength       =   80
      TabIndex        =   61
      Top             =   1635
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox TxtMontoPagovitekey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   11880
      MaxLength       =   80
      TabIndex        =   60
      Top             =   1950
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtMontoVitekey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
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
      Left            =   11880
      MaxLength       =   80
      TabIndex        =   57
      Top             =   1245
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txtclaverandon 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   11880
      MaxLength       =   80
      TabIndex        =   56
      Top             =   1590
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdConsultar 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   2550
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.TextBox txtpeso 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   12360
      MaxLength       =   80
      TabIndex        =   54
      Text            =   "0.00"
      Top             =   7440
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DtpActual 
      Height          =   400
      Left            =   11040
      TabIndex        =   50
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16711681
      CurrentDate     =   40579
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
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   10320
      TabIndex        =   49
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton OptManual 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "MANUAL"
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
      Height          =   195
      Left            =   13440
      TabIndex        =   48
      Top             =   7560
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton OptAuto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "AUTO"
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
      Height          =   195
      Left            =   13440
      TabIndex        =   47
      Top             =   7320
      Width           =   855
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
      Left            =   12780
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtPuntos 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   44
      Text            =   "0"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelivery 
      BackColor       =   &H00DFDFE0&
      Height          =   615
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2880
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CheckBox chkDelivery 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3240
      TabIndex        =   42
      Top             =   3015
      Width           =   255
   End
   Begin MSComCtl2.DTPicker DtpFechaReferencia 
      Height          =   255
      Left            =   9240
      TabIndex        =   37
      Top             =   5040
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
      Format          =   16711681
      CurrentDate     =   40579
   End
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
      Left            =   10920
      MaxLength       =   80
      TabIndex        =   33
      Top             =   1635
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TxtMontoPagado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   10920
      MaxLength       =   80
      TabIndex        =   31
      Top             =   1965
      Width           =   1935
   End
   Begin VB.CheckBox chk_factura 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "STOCK FACT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   11760
      TabIndex        =   26
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox TxtSerie 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   7440
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox TxtNumeroDoc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   8490
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   8
      Top             =   600
      Width           =   1770
   End
   Begin VB.TextBox TxtCodCliente 
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
      MaxLength       =   80
      TabIndex        =   7
      ToolTipText     =   "DNI / RUC"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox TxtDescripcionProducto 
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
      Left            =   2490
      MaxLength       =   80
      TabIndex        =   5
      Top             =   6795
      Width           =   5055
   End
   Begin VB.TextBox TxtObservacion 
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   7680
      MaxLength       =   80
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox TxtDireccion 
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
      MaxLength       =   80
      TabIndex        =   3
      Top             =   2160
      Width           =   4935
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
      TabIndex        =   2
      Top             =   885
      Visible         =   0   'False
      Width           =   1455
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
      TabIndex        =   1
      Top             =   885
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox ChkExtraer 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "EXTRAER"
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
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   360
      Left            =   330
      TabIndex        =   10
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      TabIndex        =   11
      Top             =   120
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
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcFormaPago 
      Height          =   315
      Left            =   8565
      TabIndex        =   12
      Top             =   1605
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   8454143
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
   Begin MSComCtl2.DTPicker DTPDetracion 
      Height          =   375
      Left            =   11040
      TabIndex        =   13
      Top             =   120
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
      Format          =   16711681
      CurrentDate     =   39535
   End
   Begin MSDataListLib.DataCombo DtcComprobanteGuia 
      Height          =   315
      Left            =   1560
      TabIndex        =   14
      Top             =   885
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
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
   Begin MSDataListLib.DataCombo DtTargeta 
      Height          =   315
      Left            =   10920
      TabIndex        =   32
      Top             =   1245
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   8454143
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
   Begin MSDataListLib.DataCombo DtcMoneda 
      Height          =   315
      Left            =   8565
      TabIndex        =   38
      Top             =   1245
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   8454143
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
      Height          =   1095
      Left            =   8520
      TabIndex        =   40
      Top             =   2355
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1931
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
         Name            =   "Arial"
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
      Left            =   8565
      TabIndex        =   58
      Top             =   1965
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      BackColor       =   8454143
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   14520
      TabIndex        =   78
      Top             =   5400
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ESTADO DE CUENTA"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "COMPROBANTES"
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DtpFin"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "HfFacturas"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DTPIni"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdBuscarFecha"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdBuscarFecha 
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         TabIndex        =   79
         Top             =   360
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPIni 
         Height          =   300
         Left            =   360
         TabIndex        =   80
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
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
         Format          =   16711681
         CurrentDate     =   41751
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfFacturas 
         Height          =   2720
         Left            =   120
         TabIndex        =   81
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4789
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
      Begin MSComCtl2.DTPicker DtpFin 
         Height          =   300
         Left            =   2040
         TabIndex        =   82
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
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
         Format          =   16711681
         CurrentDate     =   41751
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "AL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   1680
         TabIndex        =   83
         Top             =   405
         Width           =   165
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPrecios 
      Height          =   1215
      Left            =   7920
      TabIndex        =   85
      Top             =   7560
      Width           =   2535
      _ExtentX        =   4471
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
   Begin VB.CheckBox chkPrecios 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "DESCUENTOS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   7680
      TabIndex        =   86
      Top             =   7560
      Width           =   1575
   End
   Begin VitekeySoft.ChameleonBtn CmdQuitar 
      Height          =   345
      Left            =   11940
      TabIndex        =   102
      ToolTipText     =   "ELIMINAR ITEM"
      Top             =   6770
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
      MICON           =   "FrmVentas_.frx":6DEC
      PICN            =   "FrmVentas_.frx":6E08
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
      Left            =   12360
      TabIndex        =   103
      ToolTipText     =   "VISUALIZAR SERIES"
      Top             =   6770
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
      MICON           =   "FrmVentas_.frx":73A2
      PICN            =   "FrmVentas_.frx":73BE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frameCajaIndependiente 
      BackColor       =   &H00FFFFFF&
      Height          =   2640
      Left            =   14520
      TabIndex        =   113
      Top             =   2730
      Width           =   4455
      Begin VB.TextBox txtnumeropendientes 
         Height          =   285
         Left            =   2040
         TabIndex        =   141
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Timer timer_pendientes 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   2280
         Top             =   480
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPendientes 
         Height          =   1935
         Left            =   120
         TabIndex        =   114
         Top             =   195
         Width           =   4215
         _ExtentX        =   7435
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
      Begin VitekeySoft.ChameleonBtn cmddescartar 
         Height          =   390
         Left            =   120
         TabIndex        =   140
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   688
         BTYPE           =   5
         TX              =   "DESC.  ATENCION"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas_.frx":7958
         PICN            =   "FrmVentas_.frx":7974
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
         Height          =   390
         Left            =   2280
         TabIndex        =   193
         Top             =   2160
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   688
         BTYPE           =   5
         TX              =   "LISTADO"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas_.frx":7F0E
         PICN            =   "FrmVentas_.frx":7F2A
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
   Begin MSDataListLib.DataCombo DtcVendedor 
      Height          =   315
      Left            =   3240
      TabIndex        =   134
      Top             =   2520
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
   Begin VB.Frame FrameSerieModelo 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   2560
      Left            =   2280
      TabIndex        =   88
      Top             =   3765
      Visible         =   0   'False
      Width           =   8655
      Begin VB.TextBox txtitem 
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
         Left            =   4515
         MaxLength       =   80
         TabIndex        =   132
         ToolTipText     =   "DNI / RUC"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox txtdua 
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
         Left            =   4515
         MaxLength       =   80
         TabIndex        =   130
         ToolTipText     =   "DNI / RUC"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtMarca 
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
         Left            =   1920
         MaxLength       =   80
         TabIndex        =   107
         ToolTipText     =   "DNI / RUC"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtBuscarSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   4515
         MaxLength       =   80
         TabIndex        =   99
         ToolTipText     =   "DNI / RUC"
         Top             =   285
         Width           =   1575
      End
      Begin VB.TextBox TxtColor 
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
         Left            =   1920
         MaxLength       =   80
         TabIndex        =   98
         ToolTipText     =   "DNI / RUC"
         Top             =   2085
         Width           =   1455
      End
      Begin VB.TextBox txtbusquedamotor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   4515
         MaxLength       =   80
         TabIndex        =   97
         ToolTipText     =   "DNI / RUC"
         Top             =   720
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo DtcSerie 
         Height          =   315
         Left            =   1920
         TabIndex        =   96
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
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
      Begin VB.TextBox txtA�oFabricacion 
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
         Left            =   4515
         MaxLength       =   80
         TabIndex        =   95
         ToolTipText     =   "DNI / RUC"
         Top             =   1200
         Width           =   1575
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1920
         MaxLength       =   80
         TabIndex        =   94
         ToolTipText     =   "DNI / RUC"
         Top             =   1680
         Width           =   1455
      End
      Begin VitekeySoft.ChameleonBtn cmdcerrar 
         Height          =   375
         Left            =   6360
         TabIndex        =   100
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         BTYPE           =   5
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmVentas_.frx":84C4
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
         Height          =   315
         Left            =   1920
         TabIndex        =   115
         Top             =   720
         Width           =   2415
         _ExtentX        =   4260
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
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NRO ITEM :"
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
         Left            =   3570
         TabIndex        =   133
         Top             =   2160
         Width           =   885
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NRO DUA :"
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
         Left            =   3645
         TabIndex        =   131
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MARCA :"
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
         Left            =   1095
         TabIndex        =   106
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COLOR  :"
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
         Left            =   1095
         TabIndex        =   93
         Top             =   2085
         Width           =   705
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A�O MOD:"
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
         Left            =   3570
         TabIndex        =   92
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHASIS :"
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
         Left            =   1095
         TabIndex        =   91
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOTOR :"
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
         Left            =   1095
         TabIndex        =   90
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MODELO :"
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
         Left            =   1035
         TabIndex        =   89
         Top             =   1680
         Width           =   765
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   3015
      Left            =   120
      TabIndex        =   74
      Top             =   3600
      Width           =   13215
      _ExtentX        =   23310
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
   Begin VB.TextBox TxtCliente 
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
      MaxLength       =   80
      TabIndex        =   6
      Top             =   1800
      Width           =   4935
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   720
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":84E0
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":C243
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":FB2B
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":13B51
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":173FA
            Key             =   "(Salir)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgIconos1 
      Left            =   1200
      Top             =   2280
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
            Picture         =   "FrmVentas_.frx":1AE79
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1B195
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1B5F5
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1BA55
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1BD71
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1C1D1
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1C4ED
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1C94D
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1CDAD
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1D68D
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1D9A9
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1DCC5
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1DFE1
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1E8BB
            Key             =   "(GuiaRemision)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas_.frx":1EBD5
            Key             =   "(Imprimir)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones1 
      Height          =   6225
      Left            =   13440
      TabIndex        =   201
      Top             =   0
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   10980
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImageList2"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   960
      _CBHeight       =   6225
      _Version        =   "6.7.9782"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   900
      Width1          =   495
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   9180
         Left            =   30
         TabIndex        =   202
         Top             =   420
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   16193
         ButtonWidth     =   1614
         ButtonHeight    =   1799
         Style           =   1
         ImageList       =   "ImageList2"
         DisabledImageList=   "ImageList2"
         HotImageList    =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Nuevo (F6)"
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Anular (A)"
               Key             =   "(Anular)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Anular)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Editable"
               Key             =   "(Editable)"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   840
      Left            =   120
      TabIndex        =   203
      Top             =   8050
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   1482
      BandCount       =   1
      ForeColor       =   -2147483635
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   3750
      _CBHeight       =   840
      _Version        =   "6.7.9782"
      Child1          =   "TlbAcciones"
      MinHeight1      =   780
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbGrabar 
         Height          =   780
         Left            =   120
         TabIndex        =   204
         Top             =   15
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   1376
         ButtonWidth     =   1561
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Grabar"
               Key             =   "(Grabar)"
               Object.ToolTipText     =   "Grabar Ctrl+G"
               ImageKey        =   "(Grabar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Imprimir"
               Key             =   "(Imprimir)"
               Object.ToolTipText     =   "Grabar Ctrl+I"
               ImageKey        =   "(Imprimir)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Guia Remis"
               Key             =   "(GuiaRemision)"
               ImageKey        =   "(GuiaRemision)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Label lblAnulado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ANULADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   585
      Left            =   5280
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Label lblregistradopor 
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
      Height          =   315
      Left            =   120
      TabIndex        =   149
      Top             =   7170
      Width           =   3735
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENDEDOR :"
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
      Left            =   2160
      TabIndex        =   135
      Top             =   2640
      Width           =   945
   End
   Begin VB.Label lblContabilidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INGRESADO POR AREA DE CONTABILIDAD"
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
      Height          =   585
      Left            =   3600
      TabIndex        =   73
      Top             =   2880
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblunidad 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   7560
      TabIndex        =   87
      Top             =   6795
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DOCUMENTOS RELACIONADOS"
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
      Left            =   15600
      TabIndex        =   84
      Top             =   50
      Width           =   2340
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   14760
      Picture         =   "FrmVentas_.frx":1EC62
      Stretch         =   -1  'True
      Top             =   15
      Width           =   285
   End
   Begin VB.Label lblDisponible 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   75
      Top             =   7780
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   6  'Inside Solid
      Height          =   1965
      Left            =   225
      Top             =   1425
      Width           =   1845
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXONERADO:"
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
      Height          =   240
      Left            =   4290
      TabIndex        =   72
      Top             =   7560
      Width           =   1305
   End
   Begin VB.Label lblDescuento 
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
      Left            =   5640
      TabIndex        =   71
      Top             =   8430
      Width           =   1965
   End
   Begin VB.Label lblsincredito 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SIN LINEA DE CREDITO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   10920
      TabIndex        =   62
      Top             =   1605
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image imgFoto 
      Height          =   1935
      Left            =   240
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1815
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
      Left            =   10920
      TabIndex        =   59
      Top             =   1245
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VUELTO:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9360
      TabIndex        =   53
      Top             =   8280
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAGO    :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9420
      TabIndex        =   52
      Top             =   7920
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9435
      TabIndex        =   51
      Top             =   7560
      Width           =   1005
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   855
      Left            =   13410
      Top             =   7080
      Width           =   990
   End
   Begin VB.Label lblDelivery 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   4800
      TabIndex        =   46
      Top             =   3240
      Visible         =   0   'False
      Width           =   585
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
      TabIndex        =   41
      Top             =   7680
      Width           =   45
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
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
      Left            =   7770
      TabIndex        =   39
      Top             =   1320
      Width           =   765
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
      TabIndex        =   36
      Top             =   9795
      Visible         =   0   'False
      Width           =   1965
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
      TabIndex        =   35
      Top             =   9720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lbltargeta 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TARJETA :"
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
      Left            =   7740
      TabIndex        =   34
      Top             =   2040
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblVuelto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   10560
      TabIndex        =   30
      Top             =   8385
      Width           =   2565
   End
   Begin VB.Label lblPago 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   10560
      TabIndex        =   29
      Top             =   7935
      Width           =   2565
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCUENTO :"
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
      Height          =   240
      Left            =   4245
      TabIndex        =   28
      Top             =   8475
      Width           =   1335
   End
   Begin VB.Label LblIgv 
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
      Left            =   5640
      TabIndex        =   27
      Top             =   8115
      Width           =   1965
   End
   Begin VB.Label LblTotalLetras 
      BackColor       =   &H80000007&
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   480
      TabIndex        =   25
      Top             =   9360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label LblComprobante_DR 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FORMA PAGO:"
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
      Left            =   7440
      TabIndex        =   24
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR VENTA:"
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
      Height          =   240
      Left            =   4125
      TabIndex        =   23
      Top             =   7905
      Width           =   1455
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV   :"
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
      Height          =   240
      Left            =   4965
      TabIndex        =   22
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label lblExonerado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5640
      TabIndex        =   21
      Top             =   7485
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
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5640
      TabIndex        =   20
      Top             =   7800
      Width           =   1965
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   10560
      TabIndex        =   19
      Top             =   7485
      Width           =   2565
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DELIVERY :"
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
      Left            =   2205
      TabIndex        =   18
      Top             =   3120
      Width           =   825
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
      Left            =   12765
      TabIndex        =   17
      Top             =   6770
      Width           =   495
   End
   Begin VB.Label LblTotalParcial 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   9870
      TabIndex        =   16
      Top             =   6795
      Width           =   1560
   End
   Begin VB.Shape ShaGuia 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   120
      Top             =   825
      Width           =   1455
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   1380
      Left            =   4080
      Top             =   7440
      Width           =   9255
   End
   Begin VB.Shape ShapeDR 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2310
      Left            =   7320
      Top             =   1200
      Width           =   6015
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
      Top             =   50
      Width           =   6015
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00E0E0E0&
      Height          =   720
      Left            =   120
      Top             =   50
      Width           =   7095
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   2175
      Left            =   120
      Top             =   1320
      Width           =   7095
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   120
      Top             =   6720
      Width           =   13215
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   300
      Left            =   14520
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "FrmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doc_Tienda As String * 1
Dim cod_doc As String
Dim RstDetVenta As New ADODB.Recordset
Dim rstTemporal As New ADODB.Recordset
Dim StrCodDetVenta As String
Dim StrCodReferencia As Double
Dim Referencia As Boolean
Public Procedencia As EnumProcede
Public ProcendenciaGuia As EnumGuia
Dim DbAdelanto As Double
Public cKardex As String
Dim dfactura As String
Public codigoP As String
Dim total_descuento As Single
Dim delivery As String
Dim Descuento As Single
Dim strEspecial As Integer
Public numeroItem As Integer

Private Sub AgregarGrilla()
If Val(Me.txtcantidad.Text) > 0 And Trim(Me.TxtCodProducto.Text) <> "" Then
    'strCadena = "SELECT COUNT(DISTINCT igv) INTO ncantidad FROM temporal_ventas WHERE ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "' AND id_alm='" & KEY_ALM & "'"
    
    strCadena = "INSERT INTO temporal_ventas(ruc,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save) VALUES " & _
        "('" & KEY_RUC & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.TxtSerie.Text) & "','" & Me.TxtNumeroDoc.Text & "','" & codigoP & "','" & Val(Me.txtcantidad.Text) & "'," & _
        "'" & Val(Me.txtprecio.Text) & " ','" & Val(Me.txtprecio.Text) * Val(Me.txtcantidad.Text) & "','" & Val(Me.txtpeso.Text) & "','" & Trim(Me.TxtIgv.Text) & "','" & Trim(Me.TxtDescripcionProducto.Text) & "','" & KEY_USUARIO & "')"
       CnBd.Execute (strCadena)
    
    Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.TxtSerie.Text, Me.DtcTipoDoc.BoundText)
    Call VerificaDocumento(Trim(Me.DtcTipoDoc.BoundText))
   
    strCadena = "SELECT L.produccion FROM producto P,linea L WHERE P.id_linea=L.id_linea AND P.ruc=L.id_usu AND P.id_producto='" & Trim(Me.TxtCodProducto.Text) & "' AND P.ruc='" & KEY_RUC & "'"
    Call ConfiguraRstL(strCadena)
    If rstL("produccion") = "si" Then
        strCadena = "SELECT modelo,color,marca FROM view_producto WHERE id_producto='" & Trim(Me.TxtCodProducto.Text) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        Me.txtModelo.Text = rstL("modelo")
        Me.TxtColor.Text = rstL("color")
        Me.txtMarca.Text = rstL("marca")
        Me.txttipofactura.Text = "00002"
        
        
        
        strCadena = "SELECT Codigo,Descripcion FROM view_producto_serie WHERE  id_alm='" & KEY_ALM & "' and   vendido='no' and id_producto='" & Trim(Me.TxtCodProducto.Text) & "' AND  ruc='" & KEY_RUC & "'"
        Call ConfiguraRstT(strCadena)
        Call LlenaDataComboT(Me.DtcSerie)
        
        
        Me.FrameSerieModelo.Visible = True
        Me.cmdSeriales.Visible = True
        
        Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
        Me.TxtCodProducto.Text = "00000"
        Me.txtcantidad.Text = "0"
        Me.TxtDescripcionProducto.Text = ""
        Me.txtprecio.Text = ""
        Me.LblTotalParcial.Caption = ""
        chkPrecios.Enabled = False
        Me.HfPrecios.Visible = False
        Call Resalta(Me.txtBuscarSerie)
        Exit Sub
    Else
        Me.txttipofactura.Text = "00001"
    End If
    
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
    'Me.ChkPrecioAlterno.Value = 0
    Me.TxtCodProducto.Text = "00000"
    Me.txtcantidad.Text = "0"
    Me.TxtDescripcionProducto.Text = ""
    Me.txtprecio.Text = ""
    Me.LblTotalParcial.Caption = ""
    chkPrecios.Enabled = False
    Me.HfPrecios.Visible = False
    Call Resalta(Me.TxtCodProducto)
   ' Call DisplayTextoCom("TOTAL : S/." & AlineaString(Me.lblTotal.Caption, 9, pAlnDerecha) & _
                            "VUELTO: S/." & AlineaString(Me.lblVuelto.Caption, 9, pAlnDerecha), mscConecta)
    
Else
    Call Resalta(Me.txtcantidad)
End If
End Sub

Private Sub VerificaDocumento(ByVal TipoDoc As String)
If Trim(Me.DtcTipoDoc.BoundText) = "0009" Then
    Me.TlbGrabar.Buttons(KEY_GUIAREMISION).Enabled = True
End If
End Sub
Sub ModificarCantidad(ByVal Can_Previa As Integer, ByVal PesoProd As Double, ByVal serie As String, ByVal TipoDoc As String, ByVal Numero As String)
    Dim Can_actual As Integer
    Can_actual = Can_Previa + Val(Me.txtcantidad.Text)
    strCadena = "UPDATE Temporal_Ventas SET cantidad='" & Can_actual & "',Total='" & (Can_actual * Val(Me.txtprecio.Text)) & "'," & _
                "Peso='" & PesoProd & "' WHERE cProducto='" & Trim(Me.TxtCodProducto.Text) & "' AND sSerie='" & serie & "' AND doc_cod='" & TipoDoc & "' AND " & _
                "cDocumentoVenta='" & Numero & "'"
    Call EjecutaRST(strCadena)
    Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.DtcTipoDoc.BoundText, Me.TxtSerie.Text)
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
Public Sub llenarGrid_det(ByVal Grilla As MSHFlexGrid, ByVal id_numero As String, ByVal id_serie As String, ByVal id_doc As String)
On Error GoTo SALIR
Dim tTotal As Double
Dim texonerado As Double
Dim tafecto As Double
strCadena = "SELECT T.id,T.id_producto,T.detalle,ma.descripcion as marca,U.abreviatura,T.cantidad,T.precio,T.igv,T.total FROM temporal_ventas T,producto P,unidad U,marca ma WHERE P.id_marca=ma.id_marca and ma.id_usu=P.ruc and T.save='no' AND T.id_producto=P.id_producto AND T.id_alm='" & Me.DtcAlmacen.BoundText & "' AND T.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND T.dni_save='" & KEY_USUARIO & "' AND T.id_doc='" & id_doc & "' AND id_serie='" & id_serie & "' AND numero='" & id_numero & "' ORDER By T.id DESC"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.LblCantidad.Caption = "0"
    Me.lblExonerado.Caption = "0.00"
    Me.LblValorVenta.Caption = "0.00"
    Me.LblIgv.Caption = "0.00"
    Me.lblDescuento.Caption = "0.00"
    Me.lblTotal.Caption = "0.00"
    Me.lblPago.Caption = "0.00"
    Me.LblCantidad.Caption = 0
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 6100
           Grilla.ColWidth(3) = 1500
           Grilla.ColWidth(4) = 800
           Grilla.ColWidth(5) = 900
           Grilla.ColWidth(6) = 1100
           Grilla.ColWidth(7) = 1100
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
        Me.LblCantidad.Caption = rst.RecordCount
        For i = 0 To rst.RecordCount - 1
            If rst("id_producto") = KEY_COD_PER Then
                in_marca = ""
                in_abreviatura = ""
            Else
                in_marca = rst("marca")
                in_abreviatura = rst("abreviatura")
            End If
            
            Fila = rst("id") & vbTab & rst("id_producto") & vbTab & rst("detalle") & vbTab & in_marca & vbTab & in_abreviatura & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("precio"), "#,##0.00") & vbTab & Format(rst("total"), "#,##0.00")
            Grilla.AddItem Fila
             If (Trim(rst("igv")) = "no") Then
                            texonerado = texonerado + rst("total")
                            If KEY_APLICA_IGV = "si" Then
                                strEspecial = strEspecial + 1
                            End If
                            
             Else
                            tafecto = tafecto + rst("total")
             End If
            
            rst.MoveNext
    Next i
  tTotal = texonerado + tafecto
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 5
  Grilla.RowSel = 1


Me.lblTotal.Caption = Format(tTotal, "###0.000")
Me.lblVuelto.Caption = Format(Val(Me.lblPago.Caption) - tTotal, "###0.000")
If KEY_APLICA_IGV = "si" Then
    SUBTOTAL = tafecto / (1 + KEY_IGV)
    igv = tafecto - SUBTOTAL
Else
     texonerado = tafecto + texonerado
    SUBTOTAL = 0
    igv = 0
End If

If texonerado > 0 Then
    Me.lblExonerado.Caption = Format(texonerado, "###0.000")
End If

Me.LblIgv.Caption = Format(igv, "###0.000")
Me.LblValorVenta.Caption = Format(SUBTOTAL, "###0.000")
Me.txtcantidad.Text = 0
Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
Exit Sub
SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Private Sub SaveReferencia(ByVal codigo As String, ByVal TipoDoc As String, ByVal serie As String, ByVal Numero As String, ByVal fecha As Date, ByVal Almacen As String)
strCadena = "INSERT INTO DocReferencia_Venta(IdReferencia,doc_cod,sSerie,cDocumentoVenta,FechaProceso,Alm_Cod) VALUES " & _
            "('" & codigo & "','" & TipoDoc & "','" & serie & "','" & Numero & "','" & fecha & "','" & Almacen & "')"
            Call EjecutaRST(strCadena)
            Set rst = Nothing
End Sub
Sub llenarGrid_aa(ByVal Grilla As MSHFlexGrid, ByVal Numero As String, ByVal TipoDoc As String, ByVal serie As String)
Dim total As Double
Dim Descuento As Single
Dim SUBTOTAL As Double
Dim valor_venta As Double
Dim valor_igv As Double
Dim i As Integer
Dim registrso As Integer

strCadena = "SELECT cTemporal,Temporal_Ventas.cProducto as Codigo,Producto.DescripcionProducto as Producto,Unidad.sAbreviatura as Unidad,Temporal_Ventas.Cantidad as Cantidad,Temporal_Ventas.Precio as Precio,Temporal_Ventas.Total as Total " & _
    "FROM Temporal_Ventas INNER JOIN (Producto INNER JOIN Unidad ON Producto.cunidad=Unidad.cunidad) ON Temporal_Ventas.cProducto=Producto.cProducto WHERE (Temporal_Ventas.cDocumentoVenta='" & Numero & "' AND Temporal_Ventas.doc_cod='" & TipoDoc & "' AND Temporal_Ventas.sSerie='" & serie & "') ORDER BY cTemporal DESC"
On Error GoTo SALIR
  Call ConfiguraRst(strCadena)
  If rst.RecordCount < 1 Then
    Grilla.Clear
    Call Resalta(Me.TxtCodProducto)
    Exit Sub
  End If
  Grilla.Clear
  Grilla.Rows = 1
  Set Grilla.Recordset = rst
  Grilla.Rows = rst.RecordCount
  Grilla.ColWidth(0) = 0
  Grilla.ColWidth(1) = 900
  Grilla.ColWidth(2) = 6600
  Grilla.ColWidth(3) = 800
  Grilla.ColAlignment(3) = 7
  Grilla.ColWidth(4) = 1200
  Grilla.ColAlignment(4) = 7
  Grilla.ColWidth(5) = 1200
  Grilla.ColAlignment(5) = 7
  Grilla.ColWidth(6) = 1300
  Grilla.ColAlignment(6) = 7


  Exit Sub
SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Sub llenarGrid_Comprobante(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo SALIR
Dim tTotal As Double
Dim texonerado As Double
Dim tafecto As Double
strCadena = "SELECT D.id_detalle_venta,D.id_producto,U.abreviatura,D.cantidad,D.precio,D.total,P.id_igv,D.detalle FROM movimiento_venta_detalle D,producto P,unidad U WHERE D.id_producto=P.id_producto AND D.id_venta='" & idVenta & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "' ORDER BY id_detalle_venta ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Me.lblContabilidad.Visible = True
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 900
           Grilla.ColWidth(2) = 6600
           Grilla.ColWidth(3) = 800
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1300
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
        For i = 0 To rst.RecordCount - 1
            If rst("id_producto") = KEY_COD_PER Then
               in_producto = ""
               in_unidad = ""
               If rst("cantidad") = 0 Then
                  in_cantidad = ""
                Else
                  in_cantidad = Format(rst("cantidad"), "#,##0.00")
               End If
               
               If rst("precio") = 0 Then
                  in_precio = ""
                Else
                  in_precio = Format(rst("precio"), "#,##0.00")
               End If
               
            Else
              in_producto = rst("id_producto")
              in_unidad = rst("abreviatura")
              in_cantidad = Format(rst("cantidad"), "#,##0.00")
              in_precio = Format(rst("precio"), "#,##0.00")
            End If
            
            
            
            Fila = rst("id_detalle_venta") & vbTab & in_producto & vbTab & rst("detalle") & vbTab & in_unidad & vbTab & in_cantidad & vbTab & in_precio & vbTab & Format(rst("total"), "#,##0.00")
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
            
            rst.MoveNext
    Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1

Me.LblCantidad.Caption = Trim(rst.RecordCount)
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
If texonerado > 0 Then
    Me.lblExonerado.Caption = Format(texonerado, "###0.000")
End If

Me.LblIgv.Caption = Format(igv, "###0.000")
Me.LblValorVenta.Caption = Format(SUBTOTAL, "###0.000")
Exit Sub
SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub
Private Sub menudes1_linkClick()
    Dim i As Integer
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        MsgBox rst(0) + Chr(13) + "Debe" + Space(2) + Str(rst(1)), vbInformation, "Mensaje para el Usuario"
        rst.MoveNext
    Next i
    Set rst = Nothing
End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub chk_factura_Click()
If (Me.chk_factura.Value = 1) Then
    dfactura = "si"
Else
    dfactura = "no"
End If
End Sub

Private Sub chkConsultar_Click()
If Me.DtcAlmacen.Enabled = True Then

If (Me.chkconsultar.Value = 1) Then
    Me.TxtSerie.Locked = False
    Me.TxtNumeroDoc.Locked = False
    Call Resalta(Me.TxtSerie)
Else
    Me.TxtSerie.Locked = True
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
If Me.ChkExtraer.Value = 1 Then
    Me.ShaGuia.Width = 7095
     strCadena = "SELECT DISTINCT A.id_doc as Codigo, C.doc_abrev as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' AND A.venta='si' ORDER BY doc_abrev"
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

End Sub



Public Sub activar()
Me.TlbAcciones.Enabled = True
Me.DtcAlmacen.Enabled = True
Me.DtcTipoDoc.Enabled = True
Me.TxtSerie.Enabled = True
Me.TxtNumeroDoc.Enabled = True
Me.DtcTipoDoc.Enabled = False
   If KEY_CARGO = "00001" Then ' agente de ventas
       ' If Val(KEY_VENTANILLA) > 0 Then
        '    strCadena = "SELECT id_doc,serie,numero,igv FROM almacen_comprobante WHERE  id_doc='0099' AND  id_alm='" & KEY_VENTANILLA & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
       ' Else
        '    strCadena = "SELECT id_doc,serie,numero,igv FROM almacen_comprobante WHERE  id_doc='0099' AND  id_alm'" & KEY_ALM & "'  AND ruc='" & KEY_RUC & "'"
       ' End If
       Me.DtcTipoDoc.BoundText = "0099"
   Else
     '   strCadena = "SELECT id_doc,serie,numero,igv FROM almacen_comprobante WHERE  id_doc='" & KEY_COMPROBANTE & "' AND  id_alm='" & KEY_ALM & "' AND defecto='si' AND ruc='" & KEY_RUC & "'"
     
     Me.DtcTipoDoc.BoundText = KEY_COMPROBANTE
   End If
    
    Call comprobante(Me.DtcTipoDoc.BoundText)
    Call Nuevo
    Me.timer_pendientes.Enabled = True
    Exit Sub
    
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        KEY_APLICA_IGV = rst("igv")
        Me.DtcTipoDoc.BoundText = rst("id_doc")
        Me.TxtSerie.Text = rst("serie")
        Me.TxtNumeroDoc.Text = rst("numero")
        
    End If
    
 
    
Call Nuevo
Me.timer_pendientes.Enabled = True
End Sub

Private Sub chkPrecios_Click()
If Me.chkPrecios.Value = 1 Then
    Call llena_precios(codigoP, Me.HfPrecios)
    Me.HfPrecios.Visible = True
Else
    Me.HfPrecios.Visible = False
End If
End Sub
Public Sub mostrar_precios()
 Call llena_precios(codigoP, Me.HfPrecios)
    
End Sub



Private Sub cmdactivar_Click()
Call activar
Exit Sub
End Sub

Private Sub cmdactualizarcredito_Click()
strCadena = "UPDATE entidad_empresa SET id_credito='si',monto_credito='" & Val(Me.txtmontocredito.Text) & "' WHERE cod_unico='" & Trim(Me.TxtCodCliente.Text) & "' and id_empresa='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
Me.frmcredito.Visible = False
Call Me.precionar_cliente
End Sub

Private Sub CmdAgregar_Click()
  Call AgregarGrilla
End Sub




Private Sub cmdAgregarA_Click()

End Sub

Private Sub cmdcerrar_Click()
Me.FrameSerieModelo.Visible = False
End Sub

Private Sub cmdcerrars_Click()
Me.fraApp.Visible = False
End Sub

Private Sub cmdcerraredicion_Click()
Me.txteditable.Text = "no"
Me.frm_editable.Visible = False
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
"and v.`id_venta` = '" & Val(Me.TxtIdVenta.Text) & "'"

  Call ConfiguraRstK(strCadena)
      
  Ans = ShowMultiReport(rstK, "CorConstanciaVenta", , App.Path + "\Reportes\")

End Sub

Private Sub cmdconsultar_Click()
strCadena = "SELECT mail,nombre_completo FROM persona WHERE dni='" & Trim(Me.TxtCodCliente.Text) & "'"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
     Me.Txtclaverandon.Text = ""
     'Call enviar(Trim(rstT("mail")), rstT("nombre_completo"))
     Me.Txtclaverandon.Visible = True
    Call Resalta(Me.Txtclaverandon)
Else
   Me.Txtclaverandon.Visible = False
   MsgBox "Usuario no tiene mail para el envio", vbInformation, KEY_EMPRESA
   Exit Sub
End If
End Sub

Private Sub cmdcredito_Click()

'strCadena = "SELECT id_detalle_venta,p.nombre_prod FROM movimiento_venta_detalle d,producto p WHERE  d.id_producto=p.id_producto and d.ruc=p.ruc and d.ruc='" & KEY_RUC & "' ORDER BY id_detalle_venta DESC "
'Call ConfiguraRst(strCadena)
'If rst.RecordCount > 0 Then
 '   rst.MoveFirst
 '   For i = 0 To rst.RecordCount - 1
  '      strCadena = "UPDATE movimiento_venta_detalle SET detalle='" & rst("nombre_prod") & "' WHERE id_detalle_venta='" & rst("id_detalle_venta") & "'"
   '     CnBd.Execute (strCadena)
    '    rst.MoveNext
     '   DoEvents
   ' Next i
'End If
Procedencia = modificar_credito
FrmSeguridad.Show
Exit Sub
End Sub

Private Sub cmdDeclaracion_Click()
  
    strCadena = "select p.`nombre_completo` as cliente , p.`dni`, m.`serie`, m.`numero`,v.`doc_des` as documento,v.`doc_des` as documento, " & _
" v.`doc_des` as documento, m.`fecha_emision` " & _
"from persona p, movimiento_venta m, `comprobantes` v " & _
"where p.`dni` = m.`id_cliente` " & _
"and m.`id_doc` = v.`id_doc` " & _
"and m.`id_venta` = '" & Val(Me.TxtIdVenta.Text) & "'"
 
 
  Call ConfiguraRstK(strCadena)
      
  Ans = ShowMultiReport(rstK, "CorDeclaracion", , App.Path + "\Reportes\")

End Sub

Private Sub cmdDelivery_Click()
FrmDelivery.Show
End Sub
Public Sub llena_precios(ByVal id_producto As String, ByVal Grilla As MSHFlexGrid)
Dim porcentaje As Single


'Me.HfPrecios.MergeCells = flexMergeFree
strCadena = "SELECT * FROM almacen_producto_precio  WHERE id_producto='" & id_producto & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY precio DESC"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount < 1 Then
    Me.chkPrecios.Value = 0
    Me.HfPrecios.Visible = False
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Visible = True
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
       
        For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 300
           Grilla.ColWidth(1) = 850
           Grilla.ColWidth(2) = 1050
           
        Next
        cabecera = "" & vbTab & "PRECIO" & vbTab & "CANTIDADES"
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
          Fila = estado & vbTab & Format(rstT("precio"), "#,##0.00") & vbTab & descripcion
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
strCadena = "SELECT funct_proformas_pendientes('" & KEY_ALM & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & KEY_RUC & "')"
Call ConfiguraRstI(strCadena)
If rstI(0) > 0 Then
    PlaySound App.Path & "\sonidos\dingding.wav"
    If rstI(0) <> Val(Me.txtnumeropendientes.Text) Then
        Call llenar_pendientes(Me.HfPendientes)
    End If
End If
End Sub
Public Sub llenar_pendientes(ByVal Grilla As MSHFlexGrid)
Dim porcentaje As Single
Dim ndocumento() As String
strCadena = "SELECT id_venta,ncliente,documento,total FROM view_listado_pendientes WHERE pendiente='si' and  fecha_emision='" & KEY_FECHA & "' and id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRstI(strCadena)
Me.txtnumeropendientes.Text = rstI.RecordCount
If rstI.RecordCount < 1 Then
    
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
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


Private Sub cmdEnviar_Click()
strCadena = "select d.`nro_chasis`, d.`serie` , s.`descripcion` as modelo, " & _
"m.`descripcion` as `marca`, CONCAT( pr.`nombres`, ' ', pr.`a_paterno`, ' ', pr.`a_materno`) as paciente," & _
"pr.`direccion` , pr.`dni`, v.`fecha_emision` " & _
"from movimiento_venta v , `movimiento_venta_detalle` d, producto p, " & _
"`linea_sub` s, marca m, persona pr " & _
"where v.`id_venta` = d.`id_venta` and d.`id_producto` = p.`id_producto` and v.`ruc` = p.`ruc` " & _
"and p.`id_sublinea` = s.`id_tipo` and v.ruc = s.`id_usu` " & _
"and p.`id_marca` = m.`id_marca` and v.`ruc` = m.`id_usu` " & _
"and v.`id_cliente` = pr.`dni` and v.`id_venta` ='" & Val(Me.TxtIdVenta.Text) & "'"

Call ConfiguraRstK(strCadena)
strCadena = "select dni, nombre from  imp_reporte_app where `id_venta` = " & Val(Me.TxtIdVenta.Text)
Call ConfiguraRstL(strCadena)
      
  'Ans = ShowMultiReport(rstK, "CorApp", , App.Path + "\Reportes\", , , , , rstL, "CorAppNombres")
  'Ans = ShowMultiReport(rstK, "CorConstanciaVenta", , App.Path + "\Reportes\")
  Ans = ShowMultiReport(rstK, "CorApp", , App.Path + "\Reportes\")
'Ans = ShowMultiReport(rstK, "CorApp", , App.Path + "\Reportes\")
   
  
End Sub

Private Sub cmddescartar_Click()
If MsgBox("Esta seguro de Quitar este documento", vbQuestion + vbYesNo) = vbYes Then
   strCadena = "UPDATE movimiento_venta set pendiente='no' WHERE id_venta='" & Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) & "' and ruc"
   CnBd.Execute (strCadena)
   Me.HfPendientes.RemoveItem (Me.HfPendientes.Row)
   Me.txtnumeropendientes.Text = Val(Me.txtnumeropendientes.Text) - 1
End If
End Sub



Private Sub cmdgenerarmantenimientos_Click()
MsgBox "ESTAMOS TRABAJANDO EN ESTE MODULO", vbInformation
'Call generar_mantenimientos(Val(Me.TxtIdVenta.Text), Trim(Me.TxtCodCliente.Text), KEY_FECHA)
End Sub

Private Sub cmdGrabarGuia_Click()
    strCadena = "SELECT count(*) FROM  movimiento_transferencia WHERE id_doc='0009' and serie='" & Trim(Me.txtserieguia.Text) & "' and numero='" & Trim(Me.txtnumeroguia.Text) & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstZ(strCadena)
                    If rstZ(0) > 0 Then
                        MsgBox "Guia  ya generada verifique su correlativo", vbInformation, KEY_EMPRESA
                        Call Resalta(Me.txtnumeroguia)
                        Exit Sub
                    Else
                        Call Llenar_Temporal_transferencias(Val(Me.TxtIdVenta.Text))
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
    strCadena = "INSERT INTO movimiento_transferencia_temporal(id_doc,serie,numero,id_producto,cantidad,peso,total,dni_save,ruc) VALUES " & _
    "('0009','" & Trim(Me.txtserieguia.Text) & "','" & Trim(Me.txtnumeroguia.Text) & "','" & rst("id_producto") & "','" & rst("cantidad") & "','" & rst("peso") & "'," & _
    "'" & Val(rst("peso")) * Val(rst("cantidad")) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
    strCadena = "INSERT INTO movimiento_transferencia_series(id_doc,serie,numero,id_producto,chasis,motor,nro_dua,nro_item,ruc)  VALUES " & _
    "('0009','" & Trim(Me.txtserieguia.Text) & "','" & Trim(Me.txtnumeroguia.Text) & "','" & rst("id_producto") & "','" & rst("nro_chasis") & "','" & rst("serie") & "','" & rst("nro_dua") & "','" & rst("nro_item") & "','" & KEY_RUC & "')"
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
           strCadena = "INSERT INTO movimiento_transferencia_detalle(id_transferencia,id_producto,cantidad,recibido,peso,total,ruc) VALUES ('" & id_transferencia & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("cantidad") & "','" & rstT("peso") & "','" & rstT("total") & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           rstT.MoveNext
        Next i
        strCadena = "DELETE FROM movimiento_transferencia_temporal WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
    End If
End Sub

Private Sub save_guia()
        
        strCadena = "INSERT INTO movimiento_transferencia(id_doc,id_tipo_guia,serie,numero,fecha,id_destinatario,destinatario,id_alm_origen,id_alm_destino,id_motivo,motivo_otros,observacion,id_venta,dni_save,ruc) " & _
        "VALUES('0009','" & Trim(Me.txttipofactura.Text) & "','" & Trim(Me.txtserieguia.Text) & "','" & Trim(Me.txtnumeroguia.Text) & "','" & KEY_FECHA & "','" & Trim(Me.TxtCodCliente.Text) & "','" & Trim(Me.TxtCliente.Text) & "'," & _
        "'" & KEY_ALM & "','" & KEY_ALM & "','1','','" & Trim(Me.TxtObservacion.Text) & "','" & Val(Me.TxtIdVenta.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        id_transferencia = LastRegistro("movimiento_transferencia", "id_transferencia")
        strCadena = "UPDATE movimiento_transferencia_series SET id_transferencia='" & id_transferencia & "' WHERE id_doc='0009' and serie='" & Trim(Me.txtserieguia.Text) & "' and numero='" & Trim(Me.txtnumeroguia.Text) & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
        Call savedetalle(id_transferencia, "no")
        StrNumero = FormatosCeros(Trim(Str(Val(Me.txtnumeroguia.Text)) + 1), 6)
        strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE id_alm='" & KEY_ALM & "' AND id_doc='0009' AND serie='" & Trim(Me.txtserieguia.Text) & "' AND ruc='" & Trim(KEY_RUC) & "'"
        CnBd.Execute (strCadena)
        Me.cmdImprimirGuia.Visible = True
        Me.cmdGrabarGuia.Visible = False
        Me.txtserieguia.Locked = True
        Me.txtnumeroguia.Locked = True
End Sub
Private Sub cmdGrabarRecibo_Click()
strCadena = "SELECT * FROM movimiento_venta where  id_venta='" & Val(Me.TxtIdVenta.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
        If rst("id_recibo") = 0 Then
                    KEY_VENCIMIENTO = KEY_FECHA
                    id_tipo_factura = rst("id_tipo_factura")
                    igv = "si"
                    dfac = rst("afecta_factura")
                    
                    strCadena = "SELECT count(*) FROM  movimiento_venta WHERE id_doc='0054' and serie='" & Trim(Me.txtSerieRecibo.Text) & "' and numero='" & Trim(Me.txtNumeroRecibo.Text) & "' and ruc='" & KEY_RUC & "'"
                    Call ConfiguraRstZ(strCadena)
                    If rstZ(0) > 0 Then
                        MsgBox "Recibo ya generado verifique su correlativo", vbInformation, KEY_EMPRESA
                        Call Resalta(Me.txtNumeroRecibo)
                        Exit Sub
                    End If
                    
                    
                    Documento = "RECIBO" & ":" & Trim(Me.txtSerieRecibo.Text) & "-" & Trim(Me.txtNumeroRecibo.Text)
                    strCadena = "P_insert_venta('0054','" & KEY_ALM & "','" & Me.DtcFormaPago.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
                    "'" & Trim(Me.txtSerieRecibo.Text) & "','" & Trim(Me.txtNumeroRecibo.Text) & "','" & Me.TxtCodCliente.Text & "','" & Me.TxtCliente.Text & "','" & rst("valor_venta") & "','" & rst("igv") & "','" & rst("exonerado") & "','" & rst("total") & "','" & rst("saldo") & "'," & _
                    "'" & rst("monto_pago") & "','" & rst("saldo") & "','" & KEY_FECHA & "','" & KEY_VENCIMIENTO & "','" & id_tipo_factura & "','" & Me.DtcVendedor.BoundText & "','" & KEY_USUARIO & "','" & KEY_CAMBIO & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Documento & "','" & horario & "','T','--','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                                       
                    strCadena = "SELECT LAST_INSERT_ID() as ultimo"
                    Call ConfiguraRstL(strCadena)
                    id_venta = rstL(0)
                    strCadena = "UPDATE movimiento_venta SET id_recibo='" & id_venta & "',id_comprobante='" & Val(Me.TxtIdVenta.Text) & "' WHERE id_venta='" & Val(Me.TxtIdVenta.Text) & "'"
                    CnBd.Execute (strCadena)
                    strCadena = "UPDATE almacen_comprobante SET numero='" & Format(Val(Me.txtNumeroRecibo.Text + 1), "000000") & "' WHERE id_doc='0054' AND serie='" & Trim(Me.txtSerieRecibo.Text) & "' AND ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    
                    Call SaveDetalleDocumentoVentaRecibo(id_venta)
                    Call llenar_montos(id_venta)
                    Me.cmdGrabarRecibo.Visible = False
                    Me.cmdImprimirRecibo.Visible = True
            
        End If
    End If
End Sub

Private Sub cmdGuiaRemision_Click()

If Val(Me.TxtIdVenta.Text) > 0 Then
    strCadena = "SELECT serie,numero FROM movimiento_transferencia WHERE id_venta='" & Val(Me.TxtIdVenta.Text) & "' and ruc='" & KEY_RUC & "'"
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

End Sub

Private Sub cmdhistorial_Click()

End Sub

Private Sub cmdImprimirGuia_Click()
Call Orden_Impresion("0009", Trim(Me.txtserieguia.Text), Trim(Me.txtnumeroguia.Text), Trim(Me.txttipofactura.Text))
End Sub

Private Sub cmdImprimirRecibo_Click()
Call OrdenImpresion("0054", Trim(Me.txtSerieRecibo.Text), Trim(Me.txtNumeroRecibo.Text))
Exit Sub
End Sub

Private Sub cmdprocesar_Click()




End Sub

Private Sub cmdpersonalizada_Click()

End Sub

Private Sub cmdListado_Click()
frmventaslistado.Show
Exit Sub
End Sub

Private Sub CmdQuitar_Click()
If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    Call Quitar(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))
End If
End Sub
Private Sub Quitar(ByVal codigo As Double)
If Trim(codigo) <> "" Then
    strCadena = "DELETE FROM temporal_ventas WHERE id='" & Trim(codigo) & "'  "
    CnBd.Execute (strCadena)
    Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.TxtSerie.Text, Me.DtcTipoDoc.BoundText)
End If
End Sub



Private Sub cmdQuitarMonto_Click()
If Val(Me.HfgTipoPagos.TextMatrix(Me.HfgTipoPagos.Row, 0)) > 0 Then
    strCadena = "DELETE  FROM movimiento_venta_monto_temporal WHERE id_monto='" & Trim(Me.HfgTipoPagos.TextMatrix(Me.HfgTipoPagos.Row, 0)) & "' "
    CnBd.Execute (strCadena)
    strCadena = "DELETE FROM movimiento_venta_targeta_temporal WHERE id_temporal='" & Trim(Me.HfgTipoPagos.TextMatrix(Me.HfgTipoPagos.Row, 0)) & "' AND id_usuario='" & KEY_USUARIO & "' ANd ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Call llena_pagos(Me.HfgTipoPagos, Me.TxtNumeroDoc.Text)
    
End If
End Sub

Private Sub DtaGasto_KeyPress(KeyAscii As Integer)

End Sub

Private Sub cmdsalir_Click()
Me.fraApp.Visible = False
End Sub

Private Sub Command1_Click()
'FrmVentasCuotas.Show

strCadena = "SELECT * FROM imp_producto_detalle WHERE vendido='si'"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
    rstZ.MoveFirst
    For i = 0 To rstZ.RecordCount - 1
        strCadena = "SELECT * FROM movimiento_venta_detalle d,movimiento_venta m WHERE  (m.id_doc='0001' or m.id_doc='0003') and d.id_venta=m.id_venta and d.nro_chasis='" & rstZ("nro_chasis") & "' and m.ruc='" & KEY_RUC & "'"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount < 1 Then
            strCadena = "UPDATE imp_producto_detalle SET vendido='no' where id_detalle='" & rstZ("id_detalle") & "'"
            CnBd.Execute (strCadena)
        End If
        rstZ.MoveNext
    Next i
End If

End Sub

Private Sub cmdRecibo_Click()
strCadena = "SELECT * FROM movimiento_venta where  id_venta='" & Val(Me.TxtIdVenta.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
        If rst("id_recibo") = 0 Then
            
       
            strCadena = "SELECT numero,serie,id_doc FROM almacen_comprobante WHERE id_doc='0054' AND ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
                    Me.txtSerieRecibo.Text = rstT("serie")
                    Me.txtNumeroRecibo.Text = rstT("numero")
                    Me.txtSerieRecibo.Visible = True
                    Me.txtNumeroRecibo.Visible = True
                    Me.cmdImprimirRecibo.Visible = False
                    Me.cmdGrabarRecibo.Visible = True
                    
                    Exit Sub
            End If
        End If
End If
End Sub
Private Sub llenar_montos(ByVal id_venta As Double)
strCadena = "SELECT * FROM movimiento_venta_monto WHERE id_venta='" & Val(Me.TxtIdVenta.Text) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
    rstL.MoveFirst
    For i = 0 To rstL.RecordCount - 1
        strCadena = "INSERT INTO movimiento_venta_monto(id_venta,id_forma_pago,monto,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,ruc)VALUES('" & id_venta & "','" & rstL("id_forma_pago") & "','" & rstL("monto") & "','" & rstL("id_tarjeta") & "','" & rstL("id_tarjeta_numero") & "','" & rstL("id_tarjeta_operacion") & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        rstL.MoveNext
    Next i
End If


End Sub
Private Sub cmdSeriales_Click()
If Val(Me.TxtIdVenta.Text) > 0 And Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False Then
    strCadena = "SELECT D.serie,D.anio_fabricacion,D.anio_modelo,D.nro_chasis,D.id_producto,D.nro_dua,D.nro_item FROM movimiento_venta_detalle D WHERE  D.id_detalle_venta='" & Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "' AND id_venta='" & Val(Me.TxtIdVenta.Text) & "'"
Else
    strCadena = "SELECT * FROM temporal_ventas WHERE id='" & Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "' "
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
    Me.TxtColor.Text = rstT("color")
    Me.txtMarca.Text = rstT("marca")
    Me.txtModelo.Text = rstT("modelo")
    Me.FrameSerieModelo.Visible = True
    Exit Sub
End If
End Sub

Private Sub cmdSolicitud_Click()
  strCadena = "select d.`nro_chasis`, d.`serie` , s.`descripcion` as modelo, " & _
"m.`descripcion` as `marca`, pr.nombre_completo as paciente," & _
"CONCAT(pr.`direccion`,'-',funct_ubigueo(pr.id_departamento,pr.id_provincia,pr.id_distrito)) as direccion , pr.`dni`, v.`fecha_emision` " & _
"from movimiento_venta v , `movimiento_venta_detalle` d, producto p, " & _
"`linea_sub` s, marca m, persona pr " & _
"where v.`id_venta` = d.`id_venta` and d.`id_producto` = p.`id_producto` and v.`ruc` = p.`ruc` " & _
"and p.`id_sublinea` = s.`id_tipo` and v.ruc = s.`id_usu` " & _
"and p.`id_marca` = m.`id_marca` and v.`ruc` = m.`id_usu` " & _
"and v.`id_cliente` = pr.`dni` and v.`id_venta` ='" & Val(Me.TxtIdVenta.Text) & "'"

  Call ConfiguraRstK(strCadena)
  
  strCadena = "select dni, nombre from  imp_reporte_app where `id_venta` = " & Val(Me.TxtIdVenta.Text)
  
  Call ConfiguraRstL(strCadena)
      
  'Ans = ShowMultiReport(rstK, "CorApp", , App.Path + "\Reportes\", , , , , rstL, "CorAppNombres")
  'Ans = ShowMultiReport(rstK, "CorConstanciaVenta", , App.Path + "\Reportes\")
  Ans = ShowMultiReport(rstK, "CorApp", , App.Path + "\Reportes\")
'Ans = ShowMultiReport(rstK, "CorApp", , App.Path + "\Reportes\")


  Exit Sub
  

End Sub
Public Sub cargarParientes(ByVal Grilla As MSHFlexGrid)

strCadena = " select * from imp_reporte_app where id_venta = '" & Val(Me.TxtIdVenta.Text) & "'"
                  
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


Private Sub cmdvincular_Click()
Me.TxtMontoPagado.Text = Format(Me.HfRecibos.TextMatrix(Me.HfRecibos.Row, 3), "###0.00")
Me.txtrecibo_anterior.Text = Me.HfRecibos.TextMatrix(Me.HfRecibos.Row, 0)
Me.fraApp.Visible = False
Call realizar_ingreso_pago
Exit Sub
End Sub

Private Sub CmdVisualizar_Click()
FrmCuentasxCobrar.txtruc.Text = Trim(Me.TxtCodCliente.Text)
FrmCuentasxCobrar.Show
Exit Sub
End Sub

Private Sub Command2_Click()


End Sub

Private Sub DtcAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcTipoDoc.SetFocus
End If
End Sub

Private Sub DtcComprobanteGuia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtSeri_guia)
End If
End Sub



Private Sub forma_pago()
Dim total As Double
Dim pagado As Double
Dim tcredito As Double

If (Trim(Me.DtcFormaPago.BoundText) = "01") Then
    Me.cmdConsultar.Visible = False
    Me.lbltargeta.Visible = False
    Me.DtTargeta.Visible = False
    Me.TxtNumeroTargeta.Visible = False
    Me.DtpFechaReferencia.Visible = False
    Me.lblsincredito.Visible = False
    Me.PanelCredito.Visible = False
    Me.lblSaldodisponible.Visible = False
    Me.DtcFormapagodetalle.Visible = True
    Me.TxtMontoPagado.Enabled = True
  
    Me.TxtMontoPagado.Text = Format(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000")), "###0.000")
        
    
    
    Me.TxtMontoVitekey.Visible = False
    
ElseIf (Trim(Me.DtcFormaPago.BoundText) = "02") Then
    Me.cmdConsultar.Visible = False
    Me.lbltargeta.Visible = False
    Me.DtTargeta.Visible = False
    Me.TxtNumeroTargeta.Visible = False
    Me.DtpFechaReferencia.Visible = False
    Me.TxtMontoPagado.Text = Format(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000")), "###0.000")
    Me.TxtMontoVitekey.Visible = False
    strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & Trim(Me.TxtCodCliente.Text) & "' AND id_empresa='" & KEY_RUC & "' AND id_credito='si'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        strCadena = "SELECT sum(C.saldo) FROM movimiento_venta V,movimiento_venta_cuotas C WHERE V.id_venta=C.id_venta AND V.ruc='" & KEY_RUC & "' AND C.ruc='" & KEY_RUC & "' AND C.saldo>0 AND V.id_cliente='" & Trim(Me.TxtCodCliente.Text) & "'"
        Call ConfiguraRst(strCadena)
        If IsNull(rst(0)) = True Then
            tcredito = 0
          
        Else
            If Me.TxtNumeroDoc.Enabled = True Then
                total = 0
            Else
                total = Val(Me.lblTotal.Caption)
            End If
            tcredito = rst(0)
        End If
        If (tcredito + total) > rstT("monto_credito") Then
            Me.lblSaldodisponible.Visible = True
            Me.lblSaldodisponible.ForeColor = &HC0&
            Me.lblSaldodisponible.Caption = "SOBREPASA EL LIMITE DE CREDITO"
            Me.DtcFormapagodetalle.Visible = False
            Me.TxtMontoPagado.Enabled = False
            Exit Sub
        Else
        Me.DtcFormapagodetalle.Visible = True
        Me.TxtMontoPagado.Enabled = True
        End If
    Else
        Me.lblSaldodisponible.ForeColor = &HFF&
        Me.lblSaldodisponible.Caption = "NO AUTORIZADO"
        
    End If
    Set rstT = Nothing
End If
  
    strCadena = "SELECT id_detalle as Codigo, descripcion as Descripcion FROM forma_pago_detalle  WHERE id='" & Me.DtcFormaPago.BoundText & "' AND ruc='" & KEY_RUC & "' AND estado='si' ORDER BY id_detalle"
    Call ConfiguraRstT(strCadena)
    Call LlenaDataComboT(Me.DtcFormapagodetalle)
    Call Formapagodetalle

End Sub
Private Sub DtcFormaPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call forma_pago
    If Me.DtcFormapagodetalle.Visible = True Then
        Me.DtcFormapagodetalle.SetFocus
    Else
        Me.DtcFormaPago.SetFocus
    End If
End If
End Sub

Private Sub Formapagodetalle()
Dim tTotal As Double
If (Trim(Me.DtcFormapagodetalle.BoundText) = "01") Then
    Me.txtOperacion.Visible = False
    Me.cmdConsultar.Visible = False
    Me.Txtclaverandon.Visible = False
    Me.lbltargeta.Visible = False
    Me.DtTargeta.Visible = False
    Me.TxtNumeroTargeta.Visible = False
    Me.DtpFechaReferencia.Visible = False
    Me.TxtMontoPagado.BackColor = &H80FFFF
    Me.TxtMontoPagado.Text = Format(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000")), "###0.000")
    Me.TxtMontoVitekey.Visible = False
    Exit Sub
ElseIf (Trim(Me.DtcFormapagodetalle.BoundText) = "03" Or Trim(Me.DtcFormapagodetalle.BoundText) = "04") Then
    tTotal = Val(Format(Me.lblTotal.Caption, "###0.000"))
    pagado = Val(Format(Me.lblPago.Caption, "###0.000"))
    Me.cmdConsultar.Visible = False
    Me.TxtMontoPagado.Visible = True
    Me.Txtclaverandon.Visible = False
    Me.lbltargeta.Visible = True
    Me.DtTargeta.Visible = True
    Me.TxtNumeroTargeta.Visible = True
    Me.DtpFechaReferencia.Visible = False
    Me.TxtMontoPagado.Text = Format(tTotal - pagado, "###0.000")
    Me.DtpFechaReferencia.Value = Date
    Me.TxtMontoVitekey.Visible = False
    Exit Sub
ElseIf (Trim(Me.DtcFormapagodetalle.BoundText) = "02") Then
    Me.txtOperacion.Visible = False
    Me.TxtMontoPagado.Visible = False
    Me.lbltargeta.Visible = False
    Me.DtTargeta.Visible = False
    Me.TxtNumeroTargeta.Visible = False
    Me.DtpFechaReferencia.Visible = False
    strCadena = "SELECT sum(monto_real) FROM gigabanck WHERE dni='" & Trim(Me.TxtCodCliente.Text) & "'"
    Call ConfiguraRst(strCadena)
    If IsNull(rst(0)) = False Then
        If rst(0) <= Val(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000"))) Then
            Me.Txtclaverandon.Visible = True
            Me.TxtMontoVitekey.Visible = True
            Me.TxtMontoVitekey.Text = Str(rst(0))
            Me.cmdConsultar.Picture = LoadPicture(App.Path + "/Imagenes/noprocede.jpg")
        Else
           Me.cmdConsultar.Picture = LoadPicture(App.Path + "/Imagenes/procede.jpg")
           Me.TxtMontoVitekey.Visible = False
        End If
       Me.cmdConsultar.Enabled = True
       Me.cmdConsultar.Visible = True
    
    Exit Sub
    Else
    Exit Sub
    End If
ElseIf (Trim(Me.DtcFormapagodetalle.BoundText) = "07") Then
    strCadena = "SELECT * FROM  entidad_empresa WHERE cod_unico='" & Trim(Me.TxtCodCliente.Text) & "' ANd id_empresa='" & KEY_RUC & "' AND id_credito='si'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        Me.cmdConsultar.Picture = LoadPicture(App.Path + "/Imagenes/procede.jpg")
        Me.TxtMontoVitekey.Visible = False
        Me.cmdConsultar.Enabled = True
        Me.cmdConsultar.Visible = True
    Else
        Me.Txtclaverandon.Visible = False
        Me.cmdConsultar.Visible = False
        Me.cmdConsultar.Enabled = False
        Me.lblsincredito.Visible = True
        Me.TxtMontoPagado.Visible = False
        'Me.cmdConsultar.Picture = LoadPicture(App.Path + "/Imagenes/noprocede.jpg")
    End If

ElseIf (Trim(Me.DtcFormapagodetalle.BoundText) = "08") Then
        strCadena = "SELECT * FROM  entidad_empresa WHERE cod_unico='" & Trim(Me.TxtCodCliente.Text) & "' ANd id_empresa='" & KEY_RUC & "' AND id_credito='si'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            Me.cmdConsultar.Visible = False
            Me.cmdConsultar.Enabled = False
            Me.lblsincredito.Visible = True
            Me.TxtMontoPagado.Visible = False
            Me.PanelCredito.Visible = False
        Else
        Me.PanelCredito.Visible = True
        Me.lblclave.Visible = False
        Me.TxtClaveRandonCredito.Visible = False
        Me.TxtCuotas.Visible = True
        Me.TxtMontoVitekey.Visible = False
        Me.cmdConsultar.Enabled = False
        Me.cmdConsultar.Visible = False
        Me.txtOperacion.Visible = False
        Me.Txtclaverandon.Visible = False
        Me.lbltargeta.Visible = False
        Me.DtTargeta.Visible = False
        Me.TxtNumeroTargeta.Visible = False
        Me.DtpFechaReferencia.Visible = False
        End If
        

End If

End Sub
Private Sub DtcFormapagodetalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    
    If Me.DtcFormapagodetalle.BoundText = "01" Then
        Me.TxtMontoPagado.Visible = True
        Me.TxtMontoPagovitekey.Visible = False
        strCadena = "SELECT sum(monto) FROM movimiento_venta_monto_temporal WHERE id_forma_pago='" & Me.DtcFormapagodetalle.BoundText & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Me.TxtSerie.Text & "' ANd numero='" & Me.TxtNumeroDoc.Text & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If IsNull(rst(0)) = False Then
            Me.TxtMontoPagado.Text = Format(Val(Format(rst(0), "###0.000")) + Format(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000")), "###0.00"))
        Else
            Me.TxtMontoPagado.Text = Format(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000")), "###0.00")
        End If
        Call Resalta(Me.TxtMontoPagado)
        Exit Sub
    End If
    
    If Me.DtcFormapagodetalle.BoundText = "02" Then
    
    Me.txtOperacion.Visible = False
    Me.TxtMontoPagado.Visible = False
    Me.lbltargeta.Visible = False
    Me.DtTargeta.Visible = False
    Me.TxtNumeroTargeta.Visible = False
    Me.DtpFechaReferencia.Visible = False
    strCadena = "SELECT sum(monto_real) FROM gigabanck WHERE dni='" & Trim(Me.TxtCodCliente.Text) & "'"
    Call ConfiguraRst(strCadena)
    If IsNull(rst(0)) = False Then
        If rst(0) <= Val(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000"))) Then
            Me.Txtclaverandon.Visible = True
            Me.TxtMontoVitekey.Visible = True
            Me.TxtMontoVitekey.Text = Str(rst(0))
            Me.cmdConsultar.Picture = LoadPicture(App.Path + "/Imagenes/noprocede.jpg")
        Else
           Me.cmdConsultar.Picture = LoadPicture(App.Path + "/Imagenes/procede.jpg")
           Me.TxtMontoVitekey.Visible = False
        End If
       Me.cmdConsultar.Enabled = True
       Me.cmdConsultar.Visible = True
    
    Exit Sub
    Else
    Exit Sub
    End If
    End If
    
    
    
    
    If Me.DtcFormapagodetalle.BoundText = "03" Or Me.DtcFormapagodetalle.BoundText = "04" Then
        Me.DtTargeta.Visible = True
        Me.DtTargeta.Enabled = True
        Me.DtTargeta.SetFocus
        Exit Sub
    End If
    
    If Me.DtcFormapagodetalle.BoundText = "08" Then
        Me.PanelCredito.Visible = True
        Me.TxtCuotas.Visible = True
        Me.TxtCuotas.Text = "1"
        Call Resalta(Me.TxtCuotas)
        Exit Sub
    End If
    If Me.DtcFormapagodetalle.BoundText = "06" Then
        Call Resalta(Me.TxtMontoPagado)
        Exit Sub
    End If
    If Me.DtcFormapagodetalle.BoundText = "09" Then
        Call Resalta(Me.TxtMontoPagado)
        Exit Sub
    End If
    
    If Me.DtcFormapagodetalle.BoundText = "10" Then
    
        Call llenarGrid_recibos(Me.HfRecibos, Trim(Me.TxtCodCliente.Text))
        Me.fraApp.Visible = True
        Exit Sub
    End If
    
End If
End Sub
Private Sub llenar_recibos(ByVal dni As String)

End Sub


Private Sub DtcMoneda_Change()
If Trim(Me.DtcMoneda.Text) <> "SOLES" Then
    Me.TxtTipoCambio.Visible = True
Else
    Me.TxtTipoCambio.Visible = False
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

Private Sub DtcSerie_KeyPress(KeyAscii As Integer)
Dim nanio_modelo As String
If KeyAscii = 13 Then
    strCadena = "SELECT nro_chasis,anio_fabricacion,nro_contenedor,item,anio_modelo FROM imp_producto_detalle WHERE id_detalle='" & Val(Me.DtcSerie.BoundText) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.txtA�oFabricacion.Text = rst("anio_fabricacion")
        Me.txtdua.Text = rst("nro_contenedor")
        Me.txtitem.Text = rst("item")
        nanio_modelo = rst("anio_modelo")
        
        strCadena = "SELECT Codigo,motor as Descripcion FROM view_producto_serie WHERE codigo='" & Trim(Me.DtcSerie.BoundText) & "'"
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcMotor)
        stracdena = "UPDATE temporal_ventas SET id_detalle_serie='" & Val(Me.DtcSerie.BoundText) & "', serie='" & Trim(Me.DtcMotor.Text) & "',anio_fabricacion='" & Trim(Me.txtA�oFabricacion.Text) & "',nro_chasis='" & Me.DtcSerie.Text & "',anio_modelo='" & Trim(nanio_modelo) & "',nro_dua='" & Trim(Me.txtdua.Text) & "',nro_item='" & Trim(Me.txtitem.Text) & "' WHERE id='" & Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "'"
        CnBd.Execute (stracdena)
        Me.DtcFormaPago.SetFocus
        
        Exit Sub
    End If
End If
End Sub

Private Sub comprobante(ByVal id_doc As String)
Dim serieA As String
Dim numeroA As String
Dim comproa As String

Me.DtcTipoDoc.Enabled = True

serieA = Trim(Me.TxtSerie.Text)
numeroA = Trim(Me.TxtNumeroDoc.Text)

If Val(KEY_VENTANILLA) > 0 And Me.DtcTipoDoc.BoundText = "0099" Then
    strCadena = "SELECT serie, numero,afecta_caja,serial FROM almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND  id_alm='" & KEY_VENTANILLA & "' AND ruc='" & KEY_RUC & "' LIMIT 0,1"
Else
    strCadena = "SELECT serie, numero,afecta_caja,serial FROM almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND  id_alm='" & KEY_ALM & "' AND   ruc='" & KEY_RUC & "' LIMIT 0,1"
End If
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    Me.TxtSerie.Text = rstT("serie")
    Me.TxtNumeroDoc.Text = rstT("numero")
    Me.txtafectacaja.Text = rstT("afecta_caja")
    Me.txtserial.Text = rstT("serial")
    If serieA <> "" And numeroA <> "" Then
    strCadena = "UPDATE temporal_ventas  SET id_doc ='" & Trim(Me.DtcTipoDoc.BoundText) & "',id_serie='" & Trim(Me.TxtSerie.Text) & "',numero='" & Trim(Me.TxtNumeroDoc.Text) & "'  WHERE id_serie='" & Trim(serieA) & "' AND numero='" & Trim(numeroA) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "'"
    CnBd.Execute (strCadena)
    End If
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
                'Call Resalta(Me.TxtCodProducto)
                Call Resalta(Me.TxtCodCliente)
            End If
        End If
    End If
Else
    serieA = Trim(Me.TxtSerie.Text)
    numeroA = Trim(Me.TxtNumeroDoc.Text)
    If Val(KEY_VENTANILLA) > 0 Then
        strCadena = "SELECT serie, numero,afecta_caja,serial FROM almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND  id_alm='" & KEY_VENTANILLA & "' AND ruc='" & KEY_RUC & "' LIMIT 0,1"
    Else
        strCadena = "SELECT serie, numero,afecta_caja,serial FROM almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND  id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' LIMIT 0,1"
    End If
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        
        Me.TxtSerie.Text = rst("serie")
        Me.TxtNumeroDoc.Text = rst("numero")
        Me.txtafectacaja.Text = rst("afecta_caja")
        Me.txtserial.Text = rst("serial")
         If serieA <> "" And numeroA <> "" Then
    strCadena = "UPDATE temporal_ventas  SET id_doc ='" & Trim(Me.DtcTipoDoc.BoundText) & "',id_serie='" & Trim(Me.TxtSerie.Text) & "',numero='" & Trim(Me.TxtNumeroDoc.Text) & "'  WHERE id_serie='" & Trim(serieA) & "' AND numero='" & Trim(numeroA) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "'"
    CnBd.Execute (strCadena)
    End If
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
End If
End Sub

Private Sub DtcTipoDoc_Change()
If Me.DtcTipoDoc.Enabled = True Then
    Call comprobante(Me.DtcTipoDoc.BoundText)
 End If
End Sub

Private Sub DtcTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Me.DtcAlmacen.SetFocus
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.TxtSerie)
End If
End Sub

Private Sub DtcTipoDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
  Call comprobante(Me.DtcTipoDoc)
End If

End Sub






Private Sub DtTargeta_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
        Me.TxtNumeroTargeta.Visible = True
        Me.TxtMontoPagado.Text = Format(Val(Format(Me.lblTotal.Caption, "###0.00") - Val(Format(Me.lblPago.Caption, "###0.00"))), "###0.00")
        Me.txtOperacion.Visible = True
        Call Resalta(Me.TxtNumeroTargeta)
End If
End Sub

Sub save_temporal()
Dim codigo As String
Dim hora As Date
Dim id_codigo As Double
strCadena = "INSERT INTO temporal_venta_guardado(fecha,hora,cliente,dni_save,monto_guardado,ruc)VALUES('" & KEY_FECHA & "','" & Str(Time) & "','" & Trim(Me.TxtCliente.Text) & "','" & KEY_USUARIO & "','" & Val(Me.lblTotal.Caption) & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
id_codigo = LastRegistro("temporal_venta_guardado", "id_codigo")

strCadena = "UPDATE temporal_ventas SET save='si',id_medic='" & id_codigo & "' WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' AND id_medic='0'"
CnBd.Execute (strCadena)
Call Nuevo
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
    
    'If Me.DtcTipoDoc.BoundText = "0099" Then ' proforma
     If Trim(Me.txtafectacaja.Text) = "no" Then
            Procedencia = seleccionar_vendedor
            FrmSeguridad.Show
            Exit Sub
    End If
        
    
    
    
    strCadena = "SELECT * FROM movimiento_venta_monto_temporal WHERE id_usuario='" & KEY_USUARIO & "' AND numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND serie='" & Me.TxtSerie.Text & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If Me.DtcTipoDoc.BoundText = "0007" Then
        GoTo grabarnota
    End If
    If (Val(Me.lblPago.Caption) > 0 And rst.RecordCount > 0) Then
            
         
grabarnota:
        Call Save
        
        Call OrdenImpresion(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.TxtSerie.Text), Trim(Me.TxtNumeroDoc.Text))
        'Call Nuevo
        'If KEY_TRAMITE = "no" Then
            If Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) > 0 And KEY_CARGO = "00008" Then
              
                    strCadena = "UPDATE movimiento_venta SET pendiente='no' WHERE id_venta='" & Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
                    CnBd.Execute (strCadena)
                    Me.HfPendientes.RemoveItem (Me.HfPendientes.Row)
                    Me.txtnumeropendientes.Text = Val(Me.txtnumeropendientes.Text) - 1
                
            End If
           ' Call Nuevo
        'End If
        Exit Sub
    Else
        MsgBox "INGRESE UN MONTO VALIDO", vbExclamation, "Mensaje para la Cajera"
        If Me.TxtMontoPagado.Visible = True Then
            Call Resalta(Me.TxtMontoPagado)
        Else
            Exit Sub
        End If
    End If
  End If
If KeyCode = 117 Then
    Call Nuevo
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
            Call OrdenImpresion(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.TxtSerie.Text), Trim(Me.TxtNumeroDoc.Text))
        End If
End If
'If Shift = 2 And KeyCode = Asc("I") Then
'     Call Imprimir(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcAlmacen.BoundText), Trim(Me.TxtSerie.Text), Trim(Me.TxtNumeroDoc.Text))
'End If
  
End Sub
Private Sub Form_Load()
CenterForm Me
Me.Top = 50
 delivery = "no"
 If KEY_SKFACTURA = "si" Then
    Me.chk_factura.Visible = True
  Else
    Me.chk_factura.Visible = False
 End If
 
 If KEY_TRAMITE = "si" Then
    Me.frameTramite.Visible = True
 Else
    Me.frameCajaIndependiente.Visible = False
 End If
 If KEY_CAJA_INDEPENDIENTE = "si" Then
    Me.frameCajaIndependiente.Visible = True
Else
    Me.frameCajaIndependiente.Visible = False
 End If
 
dfactura = False
Me.DtpActual.Value = CVDate(KEY_FECHA)
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen  WHERE ruc='" & KEY_RUC & "'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  Me.DtcAlmacen.Enabled = False
  Me.TxtCodProducto.Enabled = False
  Me.TxtTipoCambio.Text = Format(KEY_CAMBIO, "#,##0.00")
  
  
    
  strCadena = "SELECT id as Codigo, descripcion as Descripcion FROM targeta ORDER BY id ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtTargeta)

  
    
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


  strCadena = "SELECT DISTINCT A.id_doc as Codigo, C.doc_abrev as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND A.ruc='" & KEY_RUC & "' AND A.venta='si' AND A.id_alm='" & KEY_ALM & "' ORDER BY doc_abrev"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
    Call LlenaDataCombo(Me.DtcTipoDoc)
    Me.DtcTipoDoc.Enabled = False
    Me.DtcTipoDoc.BoundText = 0
    'Me.DtcTipoDoc.BoundText = KEY_COMPROBANTE
  End If
  Me.DtcTipoDoc.Enabled = False
  Me.TxtSerie.Enabled = False
  Me.TxtNumeroDoc.Enabled = False
  'Me.ChkPrecioAlterno.Enabled = True
  Me.TlbAcciones.Buttons(KEY_EXIT).Enabled = True
  Me.DTPDetracion.Value = CVDate(KEY_FECHA)
  Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
  Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
  Me.TlbGrabar.Buttons(KEY_GUIAREMISION).Enabled = False
  
  
        Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
        Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
        Me.TlbAcciones.Buttons("(Editable)").Enabled = False
  
 
End Sub

Private Sub HfdDetalle_DblClick()
If (Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))) > 0 And (Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True) Then
    FrmVentaCantidad.Show
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
If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    Me.chkPrecios.Enabled = True
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
    Exit Sub
End If
End Sub

Private Sub HfFacturas_DblClick()
If Me.HfFacturas.Rows > 0 Then
    Procedencia = buscar
    FrmDetalle.Show
End If
End Sub

Private Sub HfgTipoPagos_Click()
If Val(Me.HfgTipoPagos.TextMatrix(Me.HfgTipoPagos.Row, 0)) > 0 Then
    Me.cmdQuitarMonto.Visible = True
Else
    Me.cmdQuitarMonto.Visible = False
End If
End Sub

Private Sub HfgTipoPagos_SelChange()
If Val(Me.HfgTipoPagos.TextMatrix(Me.HfgTipoPagos.Row, 0)) > 0 Then
    Me.cmdQuitarMonto.Visible = True
Else
    Me.cmdQuitarMonto.Visible = False
End If
End Sub



Private Sub mscConecta_OnComm()

End Sub

Private Sub ActualizarImagen(ByVal Grilla As MSHFlexGrid, ByVal Fila As Integer)
     Dim estado As String
      
            For i = 1 To Grilla.Rows - 1
                Grilla.TextMatrix(i, 0) = Chr(168)
            Next i
            
            
            
            
               Grilla.TextMatrix(Grilla.Row, 0) = Chr(254)
               Me.txtprecio.Text = Format(Me.HfPrecios.TextMatrix(Me.HfPrecios.Row, 1), "###0.00")
               Me.LblTotalParcial.Caption = Format(Val(Me.txtcantidad.Text) * Val(Me.txtprecio.Text), "###0.00")
               Me.txtprecio.Locked = False
                'Me.CmdAgregar.SetFocus
               Call Resalta(Me.txtprecio)
                Exit Sub
            'For j = 0 To 2
            '    HfLinea.col = j
             '   HfLinea.Row = Me.HfLinea.Row
              '  HfLinea.CellBackColor = &HC0FFC0
            'Next j
        
       ' strCadena = "UPDATE linea_medico SET estado='" & estado & "' WHERE id_linea='" & id_linea & "' AND dni='" & dni & "' AND ruc='" & KEY_RUC & "'"
        'CnBd.Execute (strCadena)
      
      
      
      
End Sub

Private Sub ActualizarPendiente(ByVal Grilla As MSHFlexGrid, ByVal Fila As Integer)
     Dim estado As String
      
            For i = 1 To Me.HfPendientes.Rows - 1
                Grilla.TextMatrix(i, 4) = Chr(168)
            Next i
            
            
            
            
                Grilla.TextMatrix(Grilla.Row, 4) = Chr(254)
                
        
            For j = 1 To 4
                Grilla.col = j
                Grilla.Row = Grilla.Row
                Grilla.CellBackColor = &HC0FFC0
            Next j
        
        
      
      
      
      
End Sub

Private Sub HfPendientes_DblClick()
If Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) > 0 Then
    Call ActualizarPendiente(Me.HfPendientes, 3)
    Call get_comprobante(Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)))
    Me.timer_pendientes.Enabled = False
End If
End Sub

Private Sub HfPendientes_SelChange()
If Val(Me.HfPendientes.TextMatrix(Me.HfPendientes.Row, 0)) > 0 Then
    Me.cmddescartar.Enabled = True
Else
    Me.cmddescartar.Enabled = False
End If
End Sub

Private Sub HfPrecios_Click()
Call ActualizarImagen(Me.HfPrecios, Me.HfPrecios.Row)
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

Private Sub timer_pendientes_Timer()
'If KEY_CAJA_INDEPENDIENTE = "si" And KEY_CARGO = "00008" Then
    Call buscar_pendientes
'End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo SALIR
Select Case Button.Key
    Case KEY_NEW
        Call Nuevo
    Case KEY_UPDATE
         FrmVentasPersonalizada.Show
    Case KEY_ANULAR
        
        If MsgBox(MSGANULAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
               Procedencia = anular
               FrmSeguridad.Show
               Exit Sub
        End If
    Case KEY_PENDIENT
            Procedencia = Modificar
            FrmSeguridad.Show
            Exit Sub
    Case "(Editable)"
            
            
            Me.txteditable.Text = "si"
            strCadena = "SELECT * FROM view_temporal WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' ORDER BY id DESC "
            Call ConfiguraRst(strCadena)
                    If rst.RecordCount > 0 Then
                       rst.MoveFirst
                           If Me.FrameSerieModelo.Visible = True Then
                                Me.TxtCodProducto_per(0).Text = rst("id_producto")
                                Me.txtprecio_per(0).Text = rst("precio")
                                Me.txtCantidadPer(0).Text = rst("cantidad")
                                
                                Me.txtdescripcion(0).Text = rst("nombre_prod")
                                Me.txtdescripcion(1).Text = rst("nro_chasis")
                                Me.txtdescripcion(2).Text = rst("serie")
                                Me.txtdescripcion(3).Text = rst("anio_modelo")
                                Me.txtdescripcion(4).Text = rst("nro_dua")
                            Else
                                
                                        Me.TxtCodProducto_per(0).Text = rst("id_producto")
                                        Me.txtprecio_per(0).Text = rst("precio")
                                        Me.txtCantidadPer(0).Text = rst("cantidad")
                                        Me.txtdescripcion(0).Text = rst("nombre_prod")
                                        Me.txtdescripcion(1).Text = rst("nro_chasis")
                                        Me.txtdescripcion(2).Text = rst("serie")
                                        Me.txtdescripcion(3).Text = rst("anio_modelo")
                                        Me.txtdescripcion(4).Text = rst("nro_dua")
                                
                            End If
                    
                    End If
             Me.frm_editable.Visible = True
             Call Resalta(Me.TxtCodProducto_per(0))
            Exit Sub
    Case KEY_DELETE
      If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Procedencia = Eliminar
        FrmSeguridad.Show
       End If
    Case KEY_EXIT
        Procedencia = Neutro
      Unload Me
  End Select
  Exit Sub
SALIR:
  MsgBox "Uso Incorrecto del Sistema", vbInformation, KEY_EMPRESA
  
End Sub
Public Sub Nuevo()
On Error GoTo SALIR
    Me.FrameReferencia.Visible = False
    Me.imgFoto.Visible = False
    Me.cmdcredito.Visible = False
    Me.txtmontocredito.Text = 0
    Me.txtOperacion.Visible = False
    Me.chkconsultar.Value = 0
    Me.DtcFormaPago.BoundText = "01"
    Me.DtcFormapagodetalle.BoundText = "01"
    Me.txtOperacion.Text = ""
    Me.TxtIdVenta.Text = ""
    Me.txttipofactura.Text = "00001"
    Me.txtrecibo_anterior.Text = 0
    Me.timer_pendientes.Enabled = True
    Me.txtA�oFabricacion.Text = ""
    Me.txtModelo.Text = ""
    Me.TxtColor.Text = ""
    Me.chkconyuge.Value = 0
    Me.txtbusquedamotor.Text = ""
    Me.DtcSerie.BoundText = ""
    Me.DtcVendedor.BoundText = "0"
    Me.PanelCredito.Visible = False
    Me.TxtMontoPagovitekey.Visible = False
    Me.lblContabilidad.Visible = False
    Me.lblsincredito.Visible = False
    Me.lblDisponible.Caption = ""
    Me.lblDisponible.Visible = False
    Me.DtTargeta.Visible = False
    Me.cmdSeriales.Visible = False
    Me.chkVincular.Visible = False
    Me.txtSerieRecibo.Locked = False
    Me.txtNumeroRecibo.Locked = False
    Me.txtSerieRecibo.Visible = False
    Me.txtNumeroRecibo.Visible = False
    Me.cmdGrabarRecibo.Visible = False
    Me.cmdImprimirRecibo.Visible = False
    Me.fraApp.Visible = False
    Me.txtserieguia.Visible = False
    Me.txtnumeroguia.Visible = False
    Me.cmdGrabarGuia.Visible = False
    Me.cmdImprimirGuia.Visible = False
    Me.DtcVendedor.Locked = False
    Me.txteditable.Text = "no"
    Me.lblregistradopor.Caption = ""
    Me.FrameSerieModelo.Visible = False
    
    For i = 0 To Me.TxtCodProducto_per.Count - 1
        Me.TxtCodProducto_per(i).Text = ""
        Me.txtCantidadPer(i).Text = ""
        Me.txtdescripcion(i).Text = ""
        Me.txtprecio_per(i).Text = ""
        Me.txttotal(i).Text = ""
    Next i
    Me.frm_editable.Visible = False
    If Me.DtcAlmacen.Enabled = True Then
    
    
  '  strCadena = "SELECT count(*) FROM movimiento_venta WHERE id_delivery='si' AND ruc='" & KEY_RUC & "' AND fecha_emision='" & KEY_FECHA & "' AND id_vendedor='" & KEY_USUARIO & "'"
  '  Call ConfiguraRst(strCadena)
  '  If rst.RecordCount > 0 Then
   '     Me.cmdDelivery.Visible = True
    '    Me.lblDelivery.Visible = True
     '   Me.lblDelivery.Caption = Str(rst(0)) & Space(1) & "Delivery"
   ' Else
    '    Me.cmdDelivery.Visible = False
     '   Me.lblDelivery.Visible = False
   ' End If
     Me.chkDelivery.Value = 0
     
     
     strCadena = "DELETE FROM temporal_ventas WHERE  dni_save='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
     CnBd.Execute (strCadena)
     strCadena = "DELETE FROM movimiento_venta_cuotas_temporal WHERE id_usuario='" & KEY_USUARIO & "' ANd ruc='" & KEY_RUC & "'"
     CnBd.Execute (strCadena)
     Me.CmdVisualizar.Visible = False
     Me.lblPendientes.Caption = ""
     strCadena = "DELETE FROM movimiento_venta_targeta_temporal WHERE id_usuario='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' "
     CnBd.Execute (strCadena)
     strCadena = "DELETE FROM movimiento_venta_monto_temporal WHERE id_usuario='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
     CnBd.Execute (strCadena)
     Me.HfgTipoPagos.Clear
     Me.cmdQuitarMonto.Visible = False
     
    If Val(KEY_VENTANILLA) > 0 And Me.DtcTipoDoc.BoundText = "0099" Then
        strCadena = "SELECT igv,serie,numero FROM almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND  id_alm='" & KEY_VENTANILLA & "' AND ruc='" & KEY_RUC & "' LIMIT 0,1"
    Else
        strCadena = "SELECT igv,serie,numero FROM almacen_comprobante WHERE id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND  id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' LIMIT 0,1"
    End If
    Call ConfiguraRst(strCadena)
    KEY_APLICA_IGV = rst("igv")
    Me.TxtSerie.Text = rst("serie")
    Me.TxtNumeroDoc.Text = rst("numero")
    Me.HfPrecios.Visible = False
    
    chkPrecios.Enabled = False
    Me.TxtCodCliente.Text = "00000000"
    Me.TxtCliente.Text = "PUBLICO EN GENERAL"
    Me.TxtDireccion.Text = KEY_DIR_PUBLIC
 
    
    
    Me.DtcMoneda.BoundText = "00001"
    Me.TxtTipoCambio.Visible = False
    Me.TxtObservacion.Text = ""
    Me.TxtDescripcionProducto.Text = ""
    Me.TxtCodProducto = "00000"
    Me.TxtDescripcionProducto.Text = ""
    Me.txtprecio.Text = ""
    Me.txtcantidad.Text = 1
    Me.LblTotalParcial.Caption = "0.00"
    Me.LblCantidad.Caption = "0"
    Me.LblTotalLetras.Caption = ""
    Me.lblPago.Caption = ""
    Me.lblVuelto.Caption = ""
    Me.lblSobrante.Caption = ""
    Me.DtcFormaPago.BoundText = "01"
    Me.TxtMontoPagado.BackColor = &HFFFFFF
    Me.TxtMontoPagado.Text = ""
    Me.Label5.Caption = "DESCUENTO"
    Me.lblDescuento.Caption = "0.00"
    Me.lblExonerado.Caption = "0.00"
    Me.TxtMontoPagado.Text = ""
    Me.TxtNumeroTargeta.Text = ""
    
    Me.TxtCodProducto.Enabled = True
    Me.TxtDescripcionProducto.Enabled = True
    Me.txtcantidad.Enabled = True
    Me.txtprecio.Enabled = True
    Me.CmdAgregar.Enabled = True
    Me.CmdQuitar.Enabled = True
    
    Me.TlbAcciones.Buttons("(Editable)").Enabled = True
   ' Me.TlbAcciones.Buttons("(Pendiente)").Enabled = True
   ' Me.TlbAcciones.Buttons(KEY_PENDIENT).Enabled = True
    Me.TlbAcciones.Buttons(KEY_NEW).Enabled = True
    Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
    Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
    
    Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
    Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
    Me.TlbGrabar.Buttons(KEY_GUIAREMISION).Enabled = False
    If Trim(Me.TxtCodCliente.Text) <> "00000000" Then
        Call llenarGrid_Facturas(Me.HfFacturas, Trim(Me.TxtCodCliente.Text))
        
    End If
    
    Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.TxtSerie.Text, Me.DtcTipoDoc.BoundText)
    'Call Resalta(Me.TxtCodProducto)
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
    'Call DisplayTextoCom(Space(5) + "BIENVENIDO A:" & AlineaString(Me.lblTotal.Caption, 8, pAlnDerecha) & _
                            "---- " & AlineaString(Me.lblTotal.Caption, 8, pAlnDerecha), mscConecta)
                            
       'Call DisplayTextoCom(AlineaString("PEPES", 20, pAlnCentro, "*") & String$(20, " "), mscConecta)
Else
    MsgBox "Active el Almacen Correspondiente", vbInformation, KEY_EMPRESA
End If
Exit Sub
Set rst = Nothing
SALIR:
MsgBox "Cree el Comprobante Seleccionado por Defecto, no hay series asignadas", vbInformation, KEY_EMPRESA
End Sub
Private Sub buscar_per(ByVal num As Integer)
 numeroItem = num
If (Len(Me.TxtCodProducto_per(num).Text) = 0) Or Val(Me.TxtCodProducto_per(num).Text) = 0 Then
        Call Resalta(Me.TxtCodProducto_per(num))
       
        Procedencia = seleccionar_per
        FrmProducto.Show
        Exit Sub
    End If
    
 
    If Trim(Mid(Me.TxtCodProducto_per(num).Text, 1, 2)) = "00" And Len(Me.TxtCodProducto_per(num).Text) > 8 Then
       Me.txtCantidadPer(num).Text = Val(Mid(Trim(Me.TxtCodProducto_per(num).Text), 8, 4) / 1000)
       Me.TxtCodProducto_per(num).Text = Mid(Me.TxtCodProducto_per(num), 3, 5)
    End If
    
    If KEY_BARRAS = "si" Then
        strCadena = "SELECT B.id_producto,P.nombre_prod,P.precio_venta,P.peso,P.id_igv FROM producto_barras B,producto P ,unidad U WHERE B.id_producto=P.id_producto AND B.ruc='" & KEY_RUC & "' " & _
        "AND P.ruc='" & KEY_RUC & "' AND B.cod_barra='" & Trim(Me.TxtCodProducto_per(num).Text) & "'"
    Else
        Me.TxtCodProducto_per(num).Text = FormatosCeros(Me.TxtCodProducto_per(num).Text, 5)
        strCadena = "SELECT A.id_producto, P.nombre_prod,P.precio_venta,P.peso,P.id_igv,U.abreviatura FROM almacen_producto A,producto P ,unidad U WHERE P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND A.id_producto=P.id_producto AND A.id_alm='" & KEY_ALM & "' AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND A.id_producto='" & Trim(Me.TxtCodProducto_per(num).Text) & "'"
    End If
        
    
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        codigoP = rst("id_producto")
        Me.txtdescripcion(num).Text = rst("nombre_prod")
        Me.txtprecio_per(num).Text = rst("precio_venta")
        Me.txtunidad(num).Text = rst("abreviatura")
        If Trim(Me.txtCantidadPer(num).Text) > 0 Then
            Me.txtCantidadPer(num).Text = Me.txtCantidadPer(num).Text
         Else
          Me.txtCantidadPer(num).Text = 1
        End If
        
        Call Resalta(Me.txtCantidadPer(num))
End If
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



Private Sub TlbAgregar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_AGREGAR
        Call AgregarGrilla
    Case KEY_QUITAR
        'Call Quitar
    
  End Select
End Sub
Public Sub get_auto_pago(ByVal in_doc As String)
If in_doc = "0099" Then

strCadena = "DELETE FROM movimiento_venta_monto_temporal WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' and serie='" & Trim(Me.TxtSerie.Text) & "' and numero='" & Trim(Me.TxtNumeroDoc.Text) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,cuotas,id_usuario,ruc) VALUES " & _
       " ('" & Me.DtcTipoDoc.BoundText & "','" & Me.TxtSerie.Text & "','" & Trim(Me.TxtNumeroDoc.Text) & "','01','" & Val(Me.lblTotal.Caption) & "','" & Val(Me.lblTotal.Caption) & "','00','-','-','0','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

Me.lblPago.Caption = Format(Val(Me.lblTotal.Caption), "###0.000")
Me.lblVuelto.Caption = Format(0, "###0.00")
End If
End Sub
Private Sub TlbGrabar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error

  Select Case Button.Key
    Case KEY_SAVE
      
     If Len(Me.TxtCodCliente.Text) < 11 And Len(Me.TxtCodCliente.Text) > 11 Then
        MsgBox "INGRESE UN RUC PARA EL CLIENTE", vbInformation, KEY_EMPRESA
        Call Resalta(Me.TxtCodCliente)
        Exit Sub
     End If
      
      
      
      If (Val(Me.lblPago.Caption) > 1 Or Me.DtcTipoDoc.BoundText = KEY_GUIA) Then
        Call get_auto_pago(Me.DtcTipoDoc.BoundText)
        Call Save
        Call Nuevo
        
    Else
        If Me.DtcTipoDoc.BoundText = "0099" Then ' proforma
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
   
    
    Case KEY_PRINT
        If strEspecial > 1 Then
        Call OrdenImpresionEspecial
        Else
            Call OrdenImpresion(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.TxtSerie.Text), Trim(Me.TxtNumeroDoc.Text))
        End If
    Case KEY_GUIAREMISION
        FrmDetalleGuia.Show
  End Select
  Exit Sub
Error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  
  Exit Sub
End Sub
Public Sub OrdenImpresion(ByVal ndoc As String, ByVal nserie As String, nnumero As String)
On Error GoTo salirr
Dim X As Integer
Dim impresiones As Integer, id_venta As Double
       strCadena = "SELECT * FROM movimiento_venta WHERE numero='" & nnumero & "' AND id_doc='" & ndoc & "' AND id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND serie='" & nserie & "' AND ruc='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
            If ndoc = "0054" Then
                GoTo imprimir_n
            End If
          If rst("impresiones") < 1 Then
imprimir_n:
              impresiones = rst("impresiones") + 1
              id_venta = rst("id_venta")
              'Call Imprimir_Tiketera(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcAlmacen.BoundText), Trim(Me.txtSerie.Text), Trim(Me.TxtNumeroDoc.Text))
              Call Orden_Impresion(ndoc, nserie, nnumero, rst("id_tipo_factura"), Trim(Me.TxtDireccion.Text))
              
              strCadena = "UPDATE movimiento_venta SET impresiones='" & impresiones & "' WHERE id_venta='" & id_venta & "' AND ruc='" & KEY_RUC & "'"
              CnBd.Execute (strCadena)
           Else
              If MsgBox("ESTE DOCUMENTO YA FUE IMPRESO:" + Space(2) + Str(rst("impresiones")) + Space(1) + "IMPRESIONES" + Chr(13) + "DESEA IMPRIMIR NUEVAMENTE ?", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
                    Procedencia = imprimir_s
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
       strCadena = "SELECT * FROM movimiento_venta WHERE numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND ruc='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
          If rst("impresiones") < 1 Then
              impresiones = rst("impresiones") + 1
              id_venta = rst("id_venta")
              'Call Imprimir_Tiketera(Trim(Me.DtcTipoDoc.BoundText), Trim(Me.DtcAlmacen.BoundText), Trim(Me.txtSerie.Text), Trim(Me.TxtNumeroDoc.Text))
              Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Trim(Me.TxtSerie.Text), Trim(Me.TxtNumeroDoc.Text), rst("id_tipo_factura"), Trim(Me.TxtDireccion.Text))
              
              strCadena = "UPDATE movimiento_venta SET impresiones='" & impresiones & "' WHERE id_venta='" & id_venta & "' AND ruc='" & KEY_RUC & "'"
              CnBd.Execute (strCadena)
           Else
              If MsgBox("ESTE DOCUMENTO YA FUE IMPRESO:" + Space(2) + Str(rst("impresiones")) + Space(1) + "IMPRESIONES" + Chr(13) + "DESEA IMPRIMIR NUEVAMENTE ?", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
                    Procedencia = imprimir_s
                    FrmSeguridad.Show
              End If
          End If
      End If
salirr: X = 1
End Sub

Public Sub OrdenImpresionEspecial()
Dim impresiones As Integer, id_venta As Double
       strCadena = "SELECT * FROM movimiento_venta WHERE numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND ruc='" & KEY_RUC & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
          If rst("impresiones") < 1 Then
              impresiones = rst("impresiones") + 1
              id_venta = rst("id_venta")
              Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Trim(Me.TxtSerie.Text), Trim(Me.TxtNumeroDoc.Text), "00002")
              strCadena = "UPDATE movimiento_venta SET impresiones='" & impresiones & "' WHERE id_venta='" & id_venta & "' AND ruc='" & KEY_RUC & "'"
              CnBd.Execute (strCadena)
              strCadena = "SELECT * FROM movimiento_venta WHERE numero='" & formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND ruc='" & KEY_RUC & "'"
              Call ConfiguraRst(strCadena)
              impresiones = rst("impresiones") + 1
              id_venta = rst("id_venta")
              Call Orden_Impresion(Me.DtcTipoDoc.BoundText, Trim(Me.TxtSerie.Text), formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6), "00002")
              strCadena = "UPDATE movimiento_venta SET impresiones='" & impresiones & "' WHERE id_venta='" & id_venta & "' AND ruc='" & KEY_RUC & "'"
              CnBd.Execute (strCadena)
           Else
              If MsgBox("ESTE DOCUMENTO YA FUE IMPRESO:" + Space(2) + Str(rst("impresiones")) + Space(1) + "IMPRESIONES" + Chr(13) + "DESEA IMPRIMIR NUEVAMENTE ?", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
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

Public Sub Imprimir_Tiketera(ByVal TipoDoc As String, ByVal CodAlm As String, ByVal serie As String, ByVal Numero As String)
Dim RstDoc As New ADODB.Recordset
Dim i As Integer, j As Integer
Dim laVenta, espacios
Dim MES As String
Dim Ans As Boolean
Dim cantidad As String, Und As String, descripcion As String, precio As String
Dim total As String, SUBTOTAL As String, igv As String
Dim totalPar As String
Dim Descuento As String
Dim GranTotal As String
Dim totalletras As String
Dim Peso As Double
Dim inc As Single
Dim codigo As String, Unidad As String, PesoTotal As Double
Dim Toneladas As String
Dim doc_identidad As String
Dim tTotal As Double
Dim tdescuento As Double
Dim tpago As Double
Dim tvuelto As Double
Dim cod_unico As String
Dim id_cliente As String
Dim fecha_doc As Date
Dim nimpresiones As Integer
  If Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False Then
    Exit Sub
  End If
    Call CargaDefConfigEpsonTM
    Printer.ScaleMode = vbCharacters  'establezco caracteres para controlar la impresion
    Printer.TrackDefault = True 'siempre apunta a la impresora predeter
    Printer.Font.name = "FontB11"
    Printer.Font.Size = "10"
    m_iTamLineaImpresion = Fix(Printer.ScaleWidth / Printer.TextWidth(" "))
    
    strCadena = "SELECT * FROM DocumentoVenta WHERE cDocumentoVenta='" & Trim(Me.TxtNumeroDoc.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "'"
    Call ConfiguraRst(strCadena)
    cod_unico = rst(0)
    
    nimpresiones = rst("impresiones") + 1
    fecha_doc = rst("dEmisionVenta")
    If rst("estado") = "Pendiente" Then
        rst("estado") = "Cancelado"
        rst.Update
    End If
    id_cliente = rst("cPersona")
    Set rst = Nothing
    
    strCadena = "UPDATE DocumentoVenta SET estado='Cancelado',impresiones='" & nimpresiones & "' WHERE cDocumentoVenta='" & Trim(Me.TxtNumeroDoc.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "'"
    CnBd.Execute (strCadena)
    
    
    If Me.DtcTipoDoc.BoundText = KEY_FACTURA Then
       Printer.CurrentX = 0
    Printer.CurrentY = 0
 'Printer.PaperSize = vbPRPSLetter
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    'Printer.Print Tab(70); (CVDate(Me.DtpActual.Value))
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(99); "FACTURA"; Space(1); Mid(Me.TxtSerie.Text + Space(50), 1, 4) & Space(1) & "-" & Me.TxtNumeroDoc.Text
    Printer.CurrentY = Printer.CurrentY + 0.7
    Printer.Print Tab(18); Mid(Me.TxtCliente.Text + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Print Tab(18); Mid(Me.TxtDireccion.Text + Space(80), 1, 75) & (CVDate(Me.DtpActual.Value))
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(18); Mid(Me.TxtCodCliente.Text + Space(100), 1, 85)
    Printer.Print ""
    Printer.Print ""
    strCadena = "SELECT Detalle_DocumentoVenta.cProducto, Detalle_DocumentoVenta.cantidad, Unidad.sAbreviatura, Producto.DescripcionProducto, " & _
            "Detalle_DocumentoVenta.Precio, Detalle_DocumentoVenta.Total, DocumentoVenta.nSubTotal, DocumentoVenta.nIgv," & _
            "DocumentoVenta.nTotalVenta FROM DocumentoVenta INNER JOIN Detalle_DocumentoVenta ON DocumentoVenta.cDocumentoVenta = Detalle_DocumentoVenta.cDocumentoVenta AND " & _
            "DocumentoVenta.doc_cod = Detalle_DocumentoVenta.doc_cod AND DocumentoVenta.Alm_cod = Detalle_DocumentoVenta.Alm_Cod AND " & _
            "DocumentoVenta.sSerie = Detalle_DocumentoVenta.sSerie INNER JOIN Seguridad ON DocumentoVenta.IdUsuario = Seguridad.IdUsuario INNER JOIN " & _
            "Producto ON Detalle_DocumentoVenta.cProducto = Producto.cProducto INNER JOIN Unidad ON Producto.cUnidad = Unidad.cUnidad " & _
            "WHERE (Detalle_DocumentoVenta.Alm_Cod='" & CodAlm & "' AND Detalle_DocumentoVenta.doc_cod='" & TipoDoc & "' AND Detalle_DocumentoVenta.sSerie='" & serie & "' " & _
             "AND Detalle_DocumentoVenta.cDocumentoVenta='" & Numero & "')"
    Call ConfiguraRst(strCadena)
    rst.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.8
            For j = 0 To rst.RecordCount - 1
                codigo = Mid(FormatosCeros(rst(0), 4) + Space(50), 1, 4)
                cantidad = Mid(Str(rst(1)) + Space(10), 1, 4)
                Und = rst(2)
                descripcion = Mid(rst(3) + Space(80), 1, 85)
                precio = Mid(Format(Str(rst(4)), "#,##0.00") + Space(4), 1, 7)
                Descuento = Mid(Format(Str(KEY_DSCTO), "#,##0.00") + Space(4), 1, 6)
                totalPar = Mid(Format(Str(rst(5)), "#,##0.00") + Space(4), 1, 7)
                Printer.Print Tab(4); cantidad & Space(10) & descripcion & "S/." & precio & Space(4) & "S/." & totalPar
                Printer.CurrentY = Printer.CurrentY + 0.4
                rst.MoveNext
            Next j
            inc = 0.5
           ' Printer.FontBold = True
            Printer.Print Tab(14); "NUEVA DF:" & KEY_DIRECCION
            'Printer.FontBold = False
            Do While (Val(Printer.CurrentY) <= 29)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    rst.MoveFirst
    total = rst(8)
    SUBTOTAL = Format(Str(Me.LblValorVenta.Caption), "#,##0.00")
    igv = Format(total - total / 1.18, "#,##0.00")
    totalletras = UCase(EnLetras(Me.lblTotal.Caption))
    Descuento = Mid(Format(Str(KEY_DSCTO), "#,##0.00") + Space(4), 1, 6)
    Set rst = Nothing
    '---- fin totales
    Printer.CurrentY = Printer.CurrentY + 1.5
    Printer.Print ""
    Printer.Print Tab(4); Mid(Me.LblTotalLetras.Caption + Space(100), 1, 66) & Space(40) & "S/." & SUBTOTAL
    Printer.CurrentY = Printer.CurrentY + 1
    'Printer.Print ""
    Printer.Print Tab(110); " S/." & igv
    Printer.CurrentY = Printer.CurrentY + 1
    'Printer.Print ""
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(110); "S/." & Format(total, "#,##0.00")
    'Printer.Print Tab(10); "ATENDIDO POR:" & Mid(Trim(KEY_VENDEDOR), 1, 10)
    Printer.EndDoc
    Exit Sub

End If
    
If Me.DtcTipoDoc.BoundText = KEY_BOLETA Then
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    If nimpresiones > 1 Then
        Printer.Print Tab(1); "==================================="
        Printer.Print Tab(5); "Copia de Original:" + Space(1) + Str(nimpresiones) + Space(1) + "Impresiones"
        Printer.Print Tab(1); "==================================="
    End If
    Printer.Print Tab(5); "PEPE'S AUTOSERVICIOS S.A.C"
    Printer.Print Tab(5); "JR. RAMON CASTILLA N� 155"
    Printer.Print Tab(5); "DELIVERY:521559  RPM:#913647"
    Printer.Print Tab(5); "Tarapoto - San Martin - San Martin"
    Printer.Print Tab(5); "RUC:20493899229"
    Printer.Print Tab(2); "-----------------------------------"
    Printer.Print Tab(4); "TIKET BOLT:"; Space(2); Mid(Me.TxtSerie.Text + Space(50), 1, 4) & "-" & Me.TxtNumeroDoc.Text & Space(2) & Trim(fecha_doc)
    Printer.Print Tab(4); "CLIENTE:" + Space(2); Mid(Me.TxtCliente.Text + Space(80), 1, 25)
    Printer.Print Tab(1); "===================================="
    Printer.Print Tab(1); "CANT" + Space(2) + "DESCRIPCION" + Space(10) + "PV" + Space(5) + "Total"
    Printer.CurrentY = Printer.CurrentY + 0.1
    Printer.Print Tab(1); "===================================="
    
    strCadena = "SELECT Detalle_DocumentoVenta.cProducto as Codigo,Detalle_DocumentoVenta.Cantidad as Cantidad,Unidad.sAbreviatura as Unidad,Producto.DescripcionProducto as Producto,Detalle_DocumentoVenta.Precio as Precio,Detalle_DocumentoVenta.Total as Total " & _
   "FROM Detalle_DocumentoVenta INNER JOIN (Producto INNER JOIN Unidad ON Producto.cunidad=Unidad.cunidad) ON Detalle_DocumentoVenta.cProducto=Producto.cProducto WHERE (Detalle_DocumentoVenta.cDocumentoVenta='" & Numero & "' AND Detalle_DocumentoVenta.doc_cod='" & TipoDoc & "' AND Detalle_DocumentoVenta.sSerie='" & serie & "' AND Detalle_DocumentoVenta.Alm_Cod='" & CodAlm & "')"
    Call ConfiguraRst(strCadena)
    rst.MoveFirst
               For j = 0 To rst.RecordCount - 1
                codigo = Mid(FormatosCeros(rst(0), 4) + Space(50), 1, 5)
                cantidad = Mid(Str(rst(1)) + Space(10), 1, 4)
                Und = rst(2)
                descripcion = Mid(rst(3) + Space(80), 1, 40)
                precio = Mid(Format(Str(rst(4)), "#,##0.00") + Space(4), 1, 6)
                totalPar = Mid(Format(Str(rst(5)), "#,##0.00") + Space(4), 1, 7)
                Printer.Print Tab(0); descripcion
                Printer.Print Tab(3); Format(cantidad, "#,##0.000") & Space(2) + Und + Space(10) & precio & Space(10) & totalPar
                Printer.CurrentY = Printer.CurrentY + 0.1
                rst.MoveNext
            Next j
            rst.MoveFirst
    
    tdescuento = Me.lblDescuento.Caption
    tpago = Me.lblPago.Caption
    tvuelto = Me.lblVuelto.Caption
    Printer.Print Tab(1); "==================================="
    Dim TTventa As Double
    TTventa = Me.lblTotal.Caption
    Printer.Print Tab(1); "DESCUENTO        S/." + Space(1) + Format(tdescuento, "#,##0.00")
    Printer.Print Tab(1); "TOTAL            S/." + Space(1) + Format(TTventa, "#,##0.00")
    strCadena = "SELECT * FROM DocumentoVenta_montos WHERE cDocumentoVenta='" & Trim(Me.TxtNumeroDoc.Text) & "' AND id_usuario='" & Trim(KEY_USUARIO) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
        If rst("id_formapago") = "0001" Then
    Printer.Print Tab(1); "EFECTIVO         S/." + Space(1) + Format(rst("monto"), "#,##0.00")
            
        End If
        If rst("id_formapago") = "0002" Then
    Printer.Print Tab(1); "TARJETA CREDITO  S/." + Space(1) + Format(rst("monto"), "#,##0.00")
            
        End If
        If rst("id_formapago") = "0003" Then
    Printer.Print Tab(1); "TARJETA DEBITO   S/." + Space(1) + Format(rst("monto"), "#,##0.00")
        
        End If
    If rst("id_formapago") = "0004" Then
    Printer.Print Tab(1); "CREDITO          S/." + Space(1) + Format(rst("monto"), "#,##0.00")
        End If
        rst.MoveNext
    Next i
    Else
     strCadena = "SELECT idFormaPago,nTotalVenta FROM DocumentoVenta WHERE  cDocumentoVenta='" & Trim(Me.TxtNumeroDoc.Text) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND " & _
     "doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "' AND id_usuario='" & Trim(KEY_USUARIO) & "'"
     Call ConfiguraRst(strCadena)
     If rst.RecordCount > 1 Then
     rst.MoveFirst
     For i = 0 To rst.RecordCount - 1
        If rst("idFormaPago") = "0001" Then
    Printer.Print Tab(1); "EFECTIVO         S/." + Space(1) + Format(rst("nTotalVenta"), "#,##0.00")
            
        End If
        If rst("idFormaPago") = "0002" Then
    Printer.Print Tab(1); "TARJETA CREDITO  S/." + Space(1) + Format(rst("nTotalVenta"), "#,##0.00")
            
        End If
        If rst("idFormaPago") = "0003" Then
    Printer.Print Tab(1); "TARJETA DEBITO   S/." + Space(1) + Format(rst("nTotalVenta"), "#,##0.00")
        
        End If
    If rst("idFormaPago") = "0004" Then
    Printer.Print Tab(1); "CREDITO          S/." + Space(1) + Format(rst("nTotalVenta"), "#,##0.00")
        End If
        rst.MoveNext
    Next i
     Else
    Printer.Print Tab(1); "EFECTIVO         S/." + Space(1) + Format(tpago, "#,##0.00")
     End If
    
    End If
    Printer.Print Tab(1); "VUELTO           S/." + Space(1) + Format(tvuelto, "#,##0.00")
    Printer.Print Tab(1); "==================================="
    Printer.Print Tab(0); "LO ATENDIO:" + Space(1) + KEY_VENDEDOR + Space(1) + "A LAS:" + Str(Time)
    Printer.Print Tab(20); "INT." & Space(2) + "ID:" & cod_unico
    Printer.Print Tab(3); " BIENES TRANSFERIDOS EN LA AMAZONIA"
    Printer.Print Tab(3); "  PARA SER CONSUMIDOS EN LA MISMA"
    Printer.Print Tab(3); "  GRACIAS POR SU COMPRA"
    Printer.Print Tab(3); "     REGRESE PRONTO  "
    Call AbreGaveta
    Printer.EndDoc
    Call Nuevo
    Exit Sub
End If
If Me.DtcTipoDoc.BoundText = KEY_NOTAPED Then
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.3
    Printer.Print Tab(4); "COD      :" + Space(1); Me.TxtCodCliente.Text
    Printer.Print Tab(4); "CLIENTE  :" + Space(1); Mid(Me.TxtCliente.Text + Space(80), 1, 40) & Space(1) & CVDate(Me.DtpActual.Value)
    Printer.Print ""
    Printer.Print Tab(4); "DIRECCI�N:" + Space(1); Mid(Me.TxtDireccion.Text + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0
    Printer.Print Tab(40); "COTIZA:"; Space(2); Mid(Me.TxtSerie.Text + Space(50), 1, 4) & Space(2) & "-" & Space(2) & Me.TxtNumeroDoc.Text
    Printer.Print "LO ATENDIO:" + Space(1) + KEY_VENDEDOR + Space(2) + "A LAS:" + Space(1) + Str(Time) + Space(3) + "FORMA PAGO:" + Space(1) + Me.DtcFormaPago.Text
    Printer.Print "----------------------------------------------------------------------"
    Printer.CurrentY = Printer.CurrentY + 0.5
strCadena = "SELECT Detalle_DocumentoVenta.cProducto, Detalle_DocumentoVenta.cantidad, Unidad.sAbreviatura, Producto.DescripcionProducto, " & _
                   "Detalle_DocumentoVenta.precio , Detalle_DocumentoVenta.TOTAL, DocumentoVenta.nTotalVenta, DocumentoVenta.nDescuento " & _
                   "FROM Detalle_DocumentoVenta INNER JOIN DocumentoVenta ON Detalle_DocumentoVenta.cDocumentoVenta = DocumentoVenta.cDocumentoVenta AND " & _
                   "Detalle_DocumentoVenta.doc_cod = DocumentoVenta.doc_cod AND Detalle_DocumentoVenta.Alm_Cod = DocumentoVenta.Alm_cod AND " & _
                   "Detalle_DocumentoVenta.sSerie = DocumentoVenta.sSerie INNER JOIN Seguridad ON DocumentoVenta.IdUsuario = Seguridad.IdUsuario INNER JOIN " & _
                   "Producto ON Detalle_DocumentoVenta.cProducto = Producto.cProducto INNER JOIN Unidad ON Producto.cUnidad = Unidad.cUnidad " & _
                   "WHERE (Detalle_DocumentoVenta.Alm_Cod='" & CodAlm & "' AND Detalle_DocumentoVenta.doc_cod='" & TipoDoc & "' AND Detalle_DocumentoVenta.sSerie='" & serie & "' " & _
                    "AND Detalle_DocumentoVenta.cDocumentoVenta='" & Numero & "')"
    Call ConfiguraRst(strCadena)
    rst.MoveFirst
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print "----------------------------------------------------------------------"
    Printer.Print Tab(1); "CODIGO" & Space(1) & "CANT" & Space(1) & "UND" & Space(2) & "DESCRIPCION" & Space(28) & "PRECIO" & Space(8) & "TOTAL"
    Printer.Print "----------------------------------------------------------------------"
            For j = 0 To rst.RecordCount - 1
                codigo = Mid(FormatosCeros(rst(0), 4) + Space(50), 1, 4)
                cantidad = rst(1)
                Und = rst(2)
                descripcion = Mid(rst(3) + Space(80), 1, 42)
                precio = Mid(Format(Str(rst(4)), "#,##0.00") + Space(4), 1, 8)
                totalPar = Mid(Format(Str(rst(5)), "#,##0.00") + Space(4), 1, 8)
                Printer.Print Tab(1); codigo & Space(2) & cantidad & Space(3) & Und & Space(2) & descripcion & precio & Space(4) & totalPar
                Printer.CurrentY = Printer.CurrentY + 0.2
                rst.MoveNext
            Next j
            inc = 0.5
            Do While (Val(Printer.CurrentY) <= 19)
                Printer.CurrentY = Printer.CurrentY + inc
            Loop
    rst.MoveFirst
    total = rst(6)
    Descuento = Format(Str(KEY_DSCTO), "#,##0.00")
    totalletras = UCase(EnLetras(Me.lblTotal.Caption))
    Set rst = Nothing
    '---- fin totales
    Printer.CurrentY = Printer.CurrentY + 0.4
    Printer.Print Tab(10); Mid(Me.LblTotalLetras.Caption + Space(100), 1, 100)
     Printer.CurrentY = Printer.CurrentY + 1
    Printer.Print Tab(45); Mid(total & Space(20), 1, 11) & Descuento & Space(8) & Format(total, "#,##0.00")
    Printer.Print "----------------------------------------------------------------------"
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(5); KEY_DIRECCION
    Printer.EndDoc
    
    Exit Sub
End If
If Me.DtcTipoDoc.BoundText = KEY_GUIA Then
    Dim RucEmpTrans As String * 11
    Dim RazonSocialTrans As String
    Dim DomicilioTrans As String
    Dim strMTC As String
    Dim marca As String, Placa As String
    Dim Licencia As String, Chofer As String
    Dim PesoFormato As String
    Dim PesoTotalForm As String
    Printer.CurrentX = 0
    Printer.CurrentY = 0
    Printer.PaperSize = 1
    Printer.Print "" 'Tab(10); "1 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "2 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "3 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "4 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "5 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "6 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "7 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "8 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "9 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "10---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "11---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "12---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "13 ---------------------------------------------------------------------"
    'Printer.Print "" 'Tab(10); "14 ---------------------------------------------------------------------"
    Printer.CurrentY = Printer.CurrentY + 0.5
    strCadena = "SELECT sOrigen, sDestino, sRazonDestinatario, sRucDestinatario, sRucTransporte, sEmpresaTransporte," & _
    "sDireccionTransporte, MTC, marca, placa, slicencia,Chofer FROM DetalleGuia WHERE DetalleGuia.sSerieGuia='" & Trim(Me.TxtSerie.Text) & "' " & _
            "AND DetalleGuia.sNumeroGuia='" & Trim(Me.TxtNumeroDoc.Text) & "'AND DetalleGuia.doc_cod='" & Me.DtcTipoDoc.BoundText & "' " & _
            "AND DetalleGuia.Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount <= 0 Then
                MsgBox "Guia Imcompleta favor de llenar Bien los Datos", vbInformation, KEY_EMPRESA
                Exit Sub
            End If
    ruc_transporte = rst(4)
    razon_transporte = rst(5)
    dir_transporte = rst(6)
    marca = rst(6)
    'MTC = rst(7)
    marca = rst(8)
    Placa = rst(9)
    Licencia = rst(10)
    Chofer = rst(11)
    Printer.Print Tab(12); Str(Me.DtpActual.Value) & Space(35) & Str(Me.DtpActual.Value) & Space(20) & "GUIAREM" & Space(1) & Me.TxtSerie.Text & "-" & Me.TxtNumeroDoc
    Printer.Print "" 'Tab(10); "16 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "17 ---------------------------------------------------------------------"
    Printer.Print Tab(1); Mid(rst(0) + Space(80), 1, 65) & Space(10) & Mid(rst(1) + Space(80), 1, 65)
    Printer.Print "" 'Tab(10); "19 ---------------------------------------------------------------------"
    'Printer.Print "" 'Tab(10); "20 ---------------------------------------------------------------------"
    Printer.Print "" ' Tab(10); "21 ---------------------------------------------------------------------"
    'Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(80); Mid(marca + Space(80), 1, 20) & Space(5) & Mid(Placa + Space(80), 1, 20)
   Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(10); Mid(rst(2) + Space(80), 1, 65) & Space(20) & Mid(MTC + Space(80), 1, 20)
    Printer.Print Tab(10); Mid(rst(3) + Space(80), 1, 20)
    Printer.Print Tab(85); Mid(Licencia + Space(80), 1, 20)
    Printer.Print Tab(80); Mid(Chofer + Space(80), 1, 20)
    Printer.Print "" ' Tab(10); "27---------------------------------------------------------------------"
    Printer.Print Tab(10); 'Tab(10); "28---------------------------------------------------------------------"
    strCadena = "SELECT Detalle_DocumentoVenta.cProducto, Unidad.sAbreviatura, Producto.DescripcionProducto, Detalle_DocumentoVenta.cantidad," & _
        "Detalle_DocumentoVenta.Peso FROM Detalle_DocumentoVenta INNER JOIN Producto ON Detalle_DocumentoVenta.cProducto = Producto.cProducto INNER JOIN " & _
        "Unidad ON Producto.cUnidad = Unidad.cUnidad WHERE (Detalle_DocumentoVenta.sSerie = '" & Trim(Me.TxtSerie.Text) & "') AND (Detalle_DocumentoVenta.cDocumentoVenta = '" & Trim(Me.TxtNumeroDoc.Text) & "') AND " & _
        "(Detalle_DocumentoVenta.doc_cod = '" & Trim(Me.DtcTipoDoc.BoundText) & "') AND (Detalle_DocumentoVenta.Alm_Cod = '" & Trim(Me.DtcAlmacen.BoundText) & "')"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount <= 0 Then
            MsgBox "No Hay Productos Registrados", vbInformation, KEY_EMPRESA
            Exit Sub
        End If
        rst.MoveFirst
        Peso = 0
            For j = 0 To rst.RecordCount - 1
                codigo = Mid((rst(0)) + Space(50), 1, 4)
                Und = Mid((rst(1)) + Space(50), 1, 4)
                descripcion = Mid(rst(2) + Space(80), 1, 70)
                cantidad = Mid(Format(Str(rst(3)), "###,##0.00") + Space(10), 1, 4)
                Toneladas = rst(4)
                'Format(Str((Rst(3) * Val(Cantidad) / 1000)), "###,##0.00")
                Peso = Peso + Toneladas
                PesoFormato = Format(Toneladas, "#,##0.00")
                Printer.Print Tab(-10); codigo & Space(4) & descripcion & Space(17) & Und & Space(4) & PesoFormato & Space(9) & cantidad
                Printer.CurrentY = Printer.CurrentY + 0.2
                rst.MoveNext
            Next j
            inc = 0.5
            Do While (Val(Printer.CurrentY) <= 35)
                Printer.CurrentY = Printer.CurrentY + inc
                            Loop
    rst.MoveFirst
    PesoTotalForm = Format(Str(Peso), "#,##0.00")
    Printer.Print Tab(50); "PESO TOTAL ->"; Space(55) & PesoTotalForm + Space(2) + "Kg."
    Printer.Print "" '29---------------------------------------------------------------------"
    Printer.Print "" '30---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "31---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "32---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "33---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "33 ---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "34---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "35---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "36---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "37---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "38---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "39---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "40---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "41---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "42---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "43---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "44---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "45---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "46---------------------------------------------------------------------"
    ''Printer.Print "" 'Tab(10); "47---------------------------------------------------------------------"
    'Printer.Print "" 'Tab(10); "48---------------------------------------------------------------------"
    ' Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(10); Mid(razon_transporte + Space(80), 1, 65); '49-------
    Printer.Print Tab(10); Mid(ruc_transporte + Space(80), 1, 65);   '50-------
    Printer.CurrentY = Printer.CurrentY + 0.2
    Printer.Print Tab(2); Mid(dir_transporte + Space(80), 1, 40);   '51-------
    Printer.Print Tab(2); Mid(dir_transporte + Space(80), 41, 50);   '51-------
    Printer.Print "" 'Tab(10); "52---------------------------------------------------------------------"
    Printer.Print "" 'Tab(10); "53---------------------------------------------------------------------"
    'Printer.Print "" 'Tab(10); "54---------------------------------------------------------------------"
    Printer.CurrentY = Printer.CurrentY + 0.5
    
    
    
    
    Printer.EndDoc
    Exit Sub
    
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    
      RucEmpTrans = Trim(rst(4))
      RazonSocialTrans = rst(5)
      DomicilioTrans = rst(6)
      strMTC = rst(7)
      marca = rst(8)
      Placa = rst(9)
      Licencia = rst(10)
      Chofer = rst(11)
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(15); Mid(rst(0) + Space(80), 1, 65)
    Printer.Print Tab(15); Mid(rst(1) + Space(80), 1, 65)
    Printer.CurrentY = Printer.CurrentY + 0.5
    Printer.Print Tab(15); Mid(rst(2) + Space(80), 1, 70)
    Printer.Print Tab(15); Mid(rst(3) + Space(80), 1, 60) & Format(KEY_DSCTO, "###,##0.00") & Space(20) & Str(Me.DtpActual.Value)
    Set rst = Nothing
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentY = Printer.CurrentY + 0.8
    
    'Printer.Print Tab(50); Me.DtcTipoDoc_Ref.Text & ":" & Me.TxtSerie_Ref.Text & "-" & Mid(Me.TxtNumero_Ref.Text + Space(100), 1, 70) & CVDate(Me.DtpFechaReferencia.Value)
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(15); RucEmpTrans & Space(30) & RazonSocialTrans
    Printer.Print Tab(50); DomicilioTrans
    Printer.Print ""
    Printer.Print ""
    Printer.Print Tab(15); strMTC & Space(30) & marca & Space(5) & Placa
    Printer.Print Tab(15); Licencia & Space(5) & Chofer
    Set rst = Nothing
    Printer.EndDoc
    Exit Sub
End If
    '----------------imprime guia------------------

End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

  
Me.TlbAcciones.Enabled = True
Me.DtcAlmacen.Enabled = True
Me.DtcTipoDoc.Enabled = True
Me.TxtSerie.Enabled = True
Me.TxtNumeroDoc.Enabled = True
Call Nuevo
Me.DtcTipoDoc.SetFocus
End Sub


Private Sub txtBuscarSerie_Change()
strCadena = "SELECT Codigo,Descripcion FROM view_producto_serie WHERE vendido='no' and id_producto='" & Trim(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 1)) & "' and Descripcion LIKE '%" & Trim(Me.txtBuscarSerie.Text) & "%' AND  ruc='" & KEY_RUC & "'"
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
If KeyAscii = 13 Then
  strCadena = "SELECT E.cod_unico as Codigo,P.nombre_completo as Descripcion FROM entidad_empresa E,persona P WHERE E.cod_unico=P.dni and E.id_personal='si' and E.id_empresa='" & KEY_RUC & "' and P.nombre_completo LIKE '%" & Trim(Me.txtBuscarVendedor.Text) & "%'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcVendedor)
  
End If
End Sub

Private Sub txtbusquedamotor_Change()
strCadena = "SELECT Codigo,motor as Descripcion FROM view_producto_serie WHERE vendido='no' and  id_producto='" & Trim(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 1)) & "' AND  motor LIKE '%" & Trim(Me.txtbusquedamotor.Text) & "%' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstT(strCadena)
Call LlenaDataComboT(Me.DtcMotor)
End Sub

Private Sub TxtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtCodProducto)
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.txtprecio)
End If
End Sub
Public Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub

Private Sub txtcantidad_KeyPress(KeyAscii As Integer)
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
    Me.txtprecio.Locked = True
    Call Resalta(Me.txtprecio)
    If Val(Me.txtprecio.Text) = 0 Then
        MsgBox "Este Producto no Cuenta con un Precio de Venta", vbExclamation
        Call Resalta(Me.TxtCodProducto)
       Exit Sub
    End If
    TotalP = Val(Me.txtcantidad.Text) * Val(Me.txtprecio.Text)
    Me.LblTotalParcial.Caption = Format(TotalP, "#,##0.00")
    
    'Me.ChkPrecioAlterno.Enabled = True
    If Me.OptAuto.Value = True Then
        Call CmdAgregar_Click
    End If
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
    
    TotalP = Val(Me.txtcantidad.Text) * Val(Me.txtprecio.Text)
    Me.LblTotalParcial.Caption = Format(TotalP, "#,##0.00")
    'Me.ChkPrecioAlterno.Enabled = True
    Call CmdAgregar_Click
    Set rst = Nothing
End Sub

Private Sub txtCantidadPer_Change(Index As Integer)
Me.txttotal(Index).Text = Format(Val(Me.txtprecio_per(Index).Text) * Val(Me.txtCantidadPer(Index).Text), "###0.00")
End Sub

Private Sub Txtclaverandon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtMontoPagovitekey.Visible = True
    Me.TxtMontoPagovitekey.Text = Format((Val(Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000")))), "###0.00")
    
    Call Resalta(Me.TxtMontoPagovitekey)
    
    Exit Sub
End If
End Sub

Private Sub TxtCliente_Change()
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
        Call Resalta(Me.TxtDireccion)
                
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
Dim Anulado As String
If dni = "00000000" Then
    GoTo sigt
End If
strCadena = "SELECT id_venta,fecha_emision,documento,total,anulado FROM movimiento_venta WHERE id_cliente='" & dni & "' AND ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC,numero DESC LIMIT 0,5"
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
strCadena = "SELECT * FROM movimiento_venta WHERE id_cliente='" & dni & "' AND ruc='" & KEY_RUC & "' AND fecha_emision>='" & Format(Me.DTPIni.Value, "YYYY-mm-dd") & "' AND fecha_emision<='" & Format(Me.DtpFin.Value, "YYYY-mm-dd") & "' ORDER BY fecha_emision DESC,numero DESC LIMIT 0,100"
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
  
 If Me.DtcTipoDoc.BoundText = "0001" And Len(Trim(Me.TxtCodCliente.Text)) <> 11 Then
     MsgBox "Debe Ingresar un Ruc para este comprobante", vbInformation
     Call Resalta(Me.TxtCodCliente)
     Exit Sub
 End If
 
 If Trim(Me.DtcTipoDoc.BoundText) = "0001" And (Trim(Me.TxtCodCliente.Text) = "00000000" Or Trim(Me.TxtCodCliente.Text) = "" Or Len(Trim(Me.TxtCodCliente.Text)) <> 11) Then
    
    Procedencia = Selecionar
    FrmPersona.Show
    
    Exit Sub
End If

If (Len(Trim(Me.TxtCodCliente.Text)) = 8 And Trim(Me.TxtCodCliente.Text) = "00000000") Then
    Me.TxtCliente.Text = "PUBLICO EN GENERAL"
    Me.TxtDireccion.Text = KEY_DIR_PUBLIC
    Call Resalta(Me.TxtCliente)
    Exit Sub
End If




 If Trim(Me.DtcTipoDoc.BoundText) = "0003" And (Trim(Me.TxtCodCliente.Text) = "") Then
    Me.TxtCodCliente.Text = "00000000"
    Me.TxtCliente.Text = "PUBLICO EN GENERAL"
    Me.TxtDireccion.Text = KEY_DIR_PUBLIC
    Call Resalta(Me.TxtCliente)
    
    Exit Sub
End If

If Len(Trim(Me.TxtCodCliente.Text)) = 8 Or Len(Trim(Me.TxtCodCliente.Text)) = 11 Then
    strCadena = "SELECT dni,nombre_completo,direccion,foto ,sexo FROM persona WHERE dni='" & Trim(Me.TxtCodCliente.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Procedencia = 1
        FrmDetallePersona.Show
        nruc = "10" & Trim(Me.TxtCodCliente.Text)
        FrmDetallePersona.txtruc.Text = DigitoVerificadorRUC(Trim(nruc))
        FrmDetallePersona.ChkCliente.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
        Me.cmdcredito.Visible = True
        Me.imgFoto.Visible = True
        Me.TxtCliente.Text = UCase(rst("nombre_completo"))
        Me.TxtDireccion.Text = UCase(rst("direccion"))
        If IsNull(rst("foto")) = False And Len(rst("foto")) > 5 Then
            If VerificarFichero(App.Path & "\archivos\" & rst("dni")) = True Then
                Me.imgFoto.Picture = LoadPicture(App.Path + "\archivos\" + rst("dni") + "\" + Trim(rst("foto")))
            Else
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
Me.Label5.Caption = "DESCUENTO(" + Str(total_descuento) + "%)"
      
        'strCadena = "SELECT sum(saldo) FROM movimiento_venta V,movimiento_venta_cuotas C WHERE V.id_venta=C.id_venta AND V.ruc='" & KEY_RUC & "' AND C.ruc='" & KEY_RUC & "' AND C.saldo>0 AND V.id_cliente='" & Trim(Me.TxtCodCliente.Text) & "' AND V.anulado='no'"
        strCadena = "SELECT sum(saldo) FROM movimiento_venta WHERE saldo>0 AND id_cliente='" & Trim(Me.TxtCodCliente.Text) & "' AND anulado='no'"
        'strCadena = "SELECT * FROM view_saldo_cliente WHERE id_cliente='" & Trim(Me.TxtCodCliente.Text) & "'"
        Call ConfiguraRst(strCadena)
        If IsNull(rst(0)) = False Then
            Me.CmdVisualizar.Visible = True
            Me.CmdVisualizar.Caption = "TIENE" + Space(2) + Format((rst(0)), "#,##0.00") + Space(2) + "DE CONSUMO"
            strCadena = "SELECT monto_credito FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' AND cod_unico='" & Trim(Me.TxtCodCliente.Text) & "'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
                 Me.lblDisponible.Visible = True
                Me.lblDisponible.Caption = "CREDITO DISPONIBLE :" + Space(2) + Format(Val(rstT("monto_credito") - rst(0)), "#,##0.00")
                Me.TxtCreditoDisponible.Text = Val(rstT("monto_credito") - rst(0))
            End If
        Else
            
            strCadena = "SELECT sum(C.saldo) FROM movimiento_venta V,movimiento_venta_cuotas C WHERE V.id_venta=C.id_venta AND V.ruc='" & KEY_RUC & "' AND C.ruc='" & KEY_RUC & "' AND C.saldo>0 AND V.id_cliente='" & Trim(Me.TxtCodCliente.Text) & "' AND V.anulado='no'"
            Call ConfiguraRst(strCadena)
            If IsNull(rst(0)) = False Then
                Me.CmdVisualizar.Visible = True
                Me.lblDisponible.Visible = True
                Me.CmdVisualizar.Caption = "TIENE" + Space(2) + Format((rst(0)), "#,##0.00") + Space(2) + "DE CONSUMO"
                strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' AND cod_unico='" & Trim(Me.TxtCodCliente.Text) & "'"
                Call ConfiguraRstT(strCadena)
                If rstT.RecordCount > 0 Then
                    Me.lblDisponible.Caption = "CREDITO DISPONIBLE :" + Space(2) + Format(Val(rstT("monto_credito") - rst(0)), "#,##0.00")
                    Me.TxtCreditoDisponible.Text = Val(rstT("monto_credito") - rst(0))
                End If
            Else
                strCadena = "SELECT * FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' AND cod_unico='" & Trim(Me.TxtCodCliente.Text) & "'"
                Call ConfiguraRstT(strCadena)
                If rstT.RecordCount > 0 Then
                    Me.lblDisponible.Visible = True
                    Me.lblDisponible.Caption = "CREDITO DISPONIBLE :" + Space(2) + Format(Val(rstT("monto_credito")), "#,##0.00")
                    Me.TxtCreditoDisponible.Text = Val(rstT("monto_credito"))
                Else
                    Me.lblDisponible.Visible = False
                End If
                
             Me.CmdVisualizar.Visible = False
             
            End If
            
            
        End If
        Set rst = Nothing
      Call llenarGrid_Facturas(Me.HfFacturas, Trim(Me.TxtCodCliente.Text))
      
      If Val(Me.lblTotal.Caption) > 0 Then
       Me.lblDescuento.Caption = Format(Me.lblTotal.Caption * total_descuento / 100, "#,##0.00")
       Descuento = Me.lblDescuento.Caption
       Me.lblTotal.Caption = Format(Me.lblTotal.Caption - Descuento, "#,##0.000")
        Set rst = Nothing
       
      End If

End Sub

Private Sub TxtCodProducto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.txtcantidad)
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.txtcantidad)
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
    
 
    If (Trim(Mid(Me.TxtCodProducto.Text, 2, 2)) = "00" Or Trim(Mid(Me.TxtCodProducto.Text, 1, 2)) = "20") And Len(Me.TxtCodProducto.Text) > 8 And Trim(Mid(Me.TxtCodProducto.Text, 1, 1)) <> "9" Then
       Me.txtcantidad.Text = Val(Mid(Trim(Me.TxtCodProducto.Text), 8, 5)) / 1000
       'Me.TxtCodProducto.text = formato_item(Mid(Me.TxtCodProducto.text, 5, 3), 5)
       Me.TxtCodProducto.Text = formato_item(Mid(Me.TxtCodProducto.Text, 2, 6), 5)
       GoTo pesable
    End If
    
    If KEY_BARRAS = "si" Then
        strCadena = "SELECT B.id_producto,P.nombre_prod,P.precio_venta,P.peso,P.id_igv,A.stock FROM producto_barras B,producto P,almacen_producto A WHERE P.id_producto=A.id_producto AND A.ruc='" & KEY_RUC & "' AND B.id_producto=P.id_producto AND B.ruc='" & KEY_RUC & "' " & _
        "AND P.ruc='" & KEY_RUC & "' AND B.cod_barra='" & Trim(Me.TxtCodProducto.Text) & "'"
    Else
pesable:
        Me.TxtCodProducto.Text = FormatosCeros(Me.TxtCodProducto.Text, 5)
        strCadena = "SELECT A.id_producto, P.nombre_prod,A.precio_venta,P.peso,P.id_igv,A.stock FROM almacen_producto A,producto P WHERE A.id_producto=P.id_producto AND A.id_alm='" & KEY_ALM & "' AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND A.id_producto='" & Trim(Me.TxtCodProducto.Text) & "'"
    End If
        
    
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If rst("stock") <= 0 Then
            strCadena = "SELECT * FROM producto WHERE id_relacionado='" & rst("id_producto") & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount < 1 Then
                If MsgBox("PRODUCTO NO CUENTA CON STOCK." + Chr(13) + "DESEA CONTINUAR ?", vbQuestion + vbYesNo, KEY_EMPRESA) = vbNo Then
                       Call Resalta(Me.TxtCodProducto)
                End If
            End If
        End If
        
        
        codigoP = rst("id_producto")
        Me.TxtDescripcionProducto.Text = rst("nombre_prod")
        Me.TxtIgv.Text = rst("id_igv")
        Me.txtprecio.Text = rst("precio_venta")
        Me.txtpreciooriginal.Text = rst("precio_venta")
        Me.txtprecio.Locked = False
        If Trim(Me.txtcantidad.Text) > 0 Then
            Me.txtcantidad.Text = Me.txtcantidad.Text
         Else
          Me.txtcantidad.Text = 1
        End If
        
        Call Resalta(Me.txtcantidad)
        
        If Me.OptAuto.Value = True Then
            Call txtcantidad_KeyPress(13)
        Else
            Call Me.mostrar_precios
        End If
        'Me.ChkPrecioAlterno.Enabled = True
        Set rst = Nothing
        
    Else
        chkPrecios.Enabled = False
        Call Resalta(Me.TxtCodProducto)
        Procedencia = Selecionar
        FrmProducto.Show
    End If
End If
End Sub

Private Sub TxtCodProducto_per_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call buscar_per(Index)
End If
End Sub

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
            "('" & formato_item(k, 2) & "','" & Me.DtcTipoDoc.BoundText & "','" & Me.TxtSerie.text & "','" & Me.TxtNumeroDoc.text & "','" & Val(Me.TxtMontoPagado.text) / Val(Me.TxtCuotas.text) & "','" & Val(Me.TxtMontoPagado.text) / Val(Me.TxtCuotas.text) & "','" & vencimiento & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
            'CnBd.Execute (strCadena)
        
            End If
            
            
        Next k
        
   
    Call Resalta(Me.TxtMontoPagado)
   
    Exit Sub
End If
End Sub

Private Sub TxtDireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtCodProducto)
End If
End Sub


Private Sub TxtMontoPagado_KeyPress(KeyAscii As Integer)

If (KeyAscii = 13) Then
    Call realizar_ingreso_pago
End If
End Sub

Private Function get_monto_caja(ByVal m_total As Single, m_pagado As Single) As Single

    If m_pagado <= m_total Then
       get_monto_caja = m_pagado
    Else
        get_monto_caja = m_total
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

     vencimiento = Format(KEY_FECHA, "YYYY-mm-dd")
    If Me.DtcTipoDoc.BoundText = "0007" Then
        GoTo sigy
    End If
    
    If (Val(Me.TxtMontoPagado.Text) > 0 And Val(Me.lblTotal.Caption) > 0) Then
sigy:
     monto_pagado = Val(Me.TxtMontoPagado.Text)
     monto_caja = get_monto_caja(Val(Me.lblTotal.Caption), monto_pagado)
     If Me.DtcMoneda.BoundText = "00002" Then
        monto_pagado = monto_pagado * Val(Me.TxtTipoCambio.Text)
     End If
     
     
     
     
     
     Me.lblPago.Caption = Format(monto_pagado, "###0.000")
     Me.lblVuelto.Caption = Format(Me.lblPago.Caption - Me.lblTotal.Caption, "###0.000")
        
        If Me.DtTargeta.Visible = False Then
            strTarjeta = "00"
        Else
            strTarjeta = Me.DtTargeta.BoundText
        End If
    
    strCadena = "SELECT * FROM movimiento_venta_monto_temporal WHERE id_usuario='" & KEY_USUARIO & "' AND id_forma_pago='" & Trim(Me.DtcFormapagodetalle.BoundText) & "' and id_moneda='" & Trim(Me.DtcMoneda.BoundText) & "'  AND ruc='" & KEY_RUC & "' AND id_tarjeta LIKE '%" & strTarjeta & "%'  ORDER BY id_monto DESC"
    Call ConfiguraRst(strCadena)
        
    If rst.RecordCount < 1 Then
        
       If Me.DtcFormaPago.BoundText = "02" Then
          strCadena = "SELECT monto_credito FROM entidad_empresa WHERE id_empresa='" & KEY_RUC & "' AND cod_unico='" & Trim(Me.TxtCodCliente.Text) & "'"
          Call ConfiguraRstT(strCadena)
          If IsNull(rstT(0)) = False Then
             If (Val(Me.TxtCreditoDisponible.Text) < Val(Me.TxtMontoPagado.Text)) Then
                 MsgBox "EL MONTO EXCEDE AL CREDITO ACTUAL EN " + Space(1) + Str(Format(Val(Me.TxtMontoPagado.Text) - Val(rstT(0)), "#,##0.00")), vbInformation, KEY_EMPRESA
                 Call Resalta(Me.TxtMontoPagado)
                 Exit Sub
             Else
                If (Me.TxtCuotas.Visible = True And Val(Me.TxtCuotas.Text) > 0) Then
                    For k = 1 To Val(Me.TxtCuotas.Text)
                        vencimiento = Format(DateAdd("m", 1, vencimiento), "YYYY-mm-dd")
                        strCadena = "INSERT INTO movimiento_venta_cuotas_temporal(id_cuota,id_doc,serie,numero,monto,saldo,vencimiento,id_usuario,ruc)VALUES " & _
                        "('" & formato_item(k, 2) & "','" & Me.DtcTipoDoc.BoundText & "','" & Me.TxtSerie.Text & "','" & Me.TxtNumeroDoc.Text & "','" & Val(Me.TxtMontoPagado.Text) / Val(Me.TxtCuotas.Text) & "','" & Val(Me.TxtMontoPagado.Text) / Val(Me.TxtCuotas.Text) & "','" & vencimiento & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                        CnBd.Execute (strCadena)
                    Next k
                End If
             End If
          End If
       End If
       
       If Val(Me.txtrecibo_anterior.Text) > 0 Then
            nrecibo = Me.HfRecibos.TextMatrix(Me.HfRecibos.Row, 2)
        Else
            nrecibo = " "
       End If
       
       strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,id_forma_pago,id_moneda,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,cuotas,id_usuario,id_recibo,detalle,ruc) VALUES " & _
       " ('" & Me.DtcTipoDoc.BoundText & "','" & Me.TxtSerie.Text & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.DtcFormapagodetalle.BoundText) & "','" & Me.DtcMoneda.BoundText & "','" & monto_pagado & "','" & monto_caja & "','" & strTarjeta & "','" & Me.TxtNumeroTargeta.Text & "','" & Me.txtOperacion.Text & "','" & Val(Me.TxtCuotas.Text) & "','" & KEY_USUARIO & "','" & Val(Me.txtrecibo_anterior.Text) & "','" & nrecibo & "','" & KEY_RUC & "')"
       CnBd.Execute (strCadena)
       
    Else
    
    If Me.DtcMoneda.BoundText = "00001" Then
        nuevo_monto = Val(Me.TxtMontoPagado.Text)
        monto_caja = get_monto_caja(Val(Me.lblTotal.Caption), Val(Me.TxtMontoPagado.Text))
    Else
        nuevo_monto = Val(Me.TxtMontoPagado.Text) * Val(Me.TxtTipoCambio.Text)
        monto_caja = get_monto_caja(Val(Me.lblTotal.Caption), Val(Me.TxtMontoPagado.Text) * Val(Me.TxtTipoCambio.Text))
    End If
    strCadena = "UPDATE movimiento_venta_monto_temporal SET monto='" & nuevo_monto & "',monto_caja='" & monto_caja & "',id_tarjeta='" & strTarjeta & "',id_tarjeta_numero='" & Me.TxtNumeroTargeta.Text & "',id_tarjeta_operacion='" & Me.txtOperacion.Text & "',id_recibo='" & Val(Me.txtrecibo_anterior.Text) & "',detalle='" & nrecibo & "' WHERE id_moneda='" & Me.DtcMoneda.BoundText & "' and id_usuario='" & KEY_USUARIO & "' AND id_forma_pago='" & Trim(Me.DtcFormapagodetalle.BoundText) & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Me.TxtNumeroDoc.Text & "' AND id_tarjeta LIKE '%" & strTarjeta & "%' AND ruc='" & KEY_RUC & "'"
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
On Error GoTo SALIR
Dim tpago As Double
Dim strTarjeta As String
strCadena = "SELECT * FROM movimiento_venta_monto_temporal M,forma_pago_detalle F WHERE M.id_forma_pago=F.id_detalle AND id_usuario='" & KEY_USUARIO & "' AND numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Me.TxtSerie.Text & "' AND M.ruc='" & KEY_RUC & "' AND F.ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Me.lblPago.Caption = "0.00"
    Me.lblVuelto.Caption = "0.00"
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
    tventa = Val(Format(Me.lblTotal.Caption, "###0.000"))
    Me.lblPago.Caption = Format(tpago, "###0.000")
    Me.lblVuelto.Caption = Format(tpago - tventa, "#,##0.000")
Exit Sub
SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub
Public Sub llena_pagosVenta(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo SALIR
Dim tpago As Double
strCadena = "SELECT * FROM movimiento_venta_monto M,forma_pago_detalle F WHERE M.id_forma_pago=F.id_detalle AND id_venta='" & idVenta & "' AND M.ruc='" & KEY_RUC & "' AND F.ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    If Me.chkconsultar.Value = 0 Then
        Me.lblPago.Caption = "0.00"
    End If
    Me.lblVuelto.Caption = "0.00"
    Exit Sub
    
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2000
           Grilla.ColWidth(2) = 1500
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
            Fila = rst("id_detalle") & vbTab & rst("descripcion") & vbTab & Format(rst("monto"), "###0.00")
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
SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
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

Private Sub TxtMontoPagovitekey_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Call agregar_monto_pago
End If
End Sub
Private Sub agregar_monto_pago()
Dim nrecibo As String
strCadena = "SELECT * FROM empresa_random WHERE dni='" & Trim(Me.TxtCodCliente.Text) & "' AND id_empresa='" & KEY_RUC & "' ORDER BY id DESC"
  Call ConfiguraRstT(strCadena)
  If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    If (Trim(Me.Txtclaverandon.Text) = Trim(rstT("codigo"))) Then
       strCadena = "SELECT * FROM movimiento_venta_monto_temporal WHERE id_usuario='" & KEY_USUARIO & "' AND id_forma_pago='" & Trim(Me.DtcFormapagodetalle.BoundText) & "' AND ruc='" & KEY_RUC & "'  ORDER BY id_monto DESC"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount < 1 Then
        If Me.TxtMontoVitekey.Visible = True Then
            strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,id_forma_pago,monto,id_usuario,id_recibo,ruc) VALUES ('" & Me.DtcTipoDoc.BoundText & "','" & Me.TxtSerie.Text & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.DtcFormapagodetalle.BoundText) & "','" & Val(Me.TxtMontoVitekey.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        Else
            strCadena = "INSERT INTO movimiento_venta_monto_temporal(id_doc,serie,numero,id_forma_pago,monto,id_usuario,ruc) VALUES ('" & Me.DtcTipoDoc.BoundText & "','" & Me.TxtSerie.Text & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Trim(Me.DtcFormapagodetalle.BoundText) & "','" & Val(Me.TxtMontoPagovitekey.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
        End If
       Else
       If Me.DtcMoneda.BoundText = "00001" Then
            If Me.TxtMontoVitekey.Visible = True Then
               ' nuevo_monto = Val(Me.TxtMontoVitekey.text)
            Else
                nuevo_monto = Val(Format(Me.lblTotal.Caption, "###0.000")) - Val(Format(Me.lblPago.Caption, "###0.000"))
            End If
        Else
            If Me.TxtMontoVitekey.Visible = True Then
               ' nuevo_monto = Val(Me.TxtMontoVitekey.text) * Val(Me.LblTipoCambio.Caption)
            Else
               ' nuevo_monto = Val(Format(Me.lblTotal.Caption, "###0.000")) * Val(Me.LblTipoCambio.Caption)
            End If
            
            
        End If
        strCadena = "UPDATE movimiento_venta_monto_temporal SET monto='" & nuevo_monto & "' WHERE id_usuario='" & KEY_USUARIO & "' AND id_forma_pago='" & Trim(Me.DtcFormapagodetalle.BoundText) & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Me.TxtNumeroDoc.Text & "' AND ruc='" & KEY_RUC & "'"
        End If
        CnBd.Execute (strCadena)
        Call llena_pagos(Me.HfgTipoPagos, Me.TxtNumeroDoc.Text)
     Else
      MsgBox "Clave Incorrecta", vbInformation, KEY_EMPRESA
      
    End If
  End If
End Sub
Private Sub TxtNumero_guia_KeyPress(KeyAscii As Integer)
Dim idVenta As Double
If KeyAscii = 13 Then
    
    Me.TxtNumero_guia.Text = FormatosCeros(Me.TxtNumero_guia.Text, 6)
    strCadena = "SELECT * FROM movimiento_venta WHERE (numero='" & Trim(Me.TxtNumero_guia.Text) & "' AND id_doc='" & Trim(Me.DtcComprobanteGuia.BoundText) & "' AND serie='" & Trim(Me.TxtSeri_guia.Text) & "' AND id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND ruc='" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Me.HfdDetalle.Clear
         MsgBox "DOCUMENTO NO REGISTRADO ", vbInformation, KEY_EMPRESA
         Call Resalta(Me.TxtNumero_guia)
         Exit Sub
      
    Else
            idVenta = rst("id_venta")
            If MsgBox("ESTA SEGURO DE REALIZAR ESTA OPERACI�N", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
                Me.TxtIdVenta.Text = idVenta
                Me.txttipofactura.Text = rst("id_tipo_factura")
                Call LlenarDatosCliente(idVenta)
                Call Llenar_Temporal(idVenta)
                Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.TxtSerie.Text, Me.DtcTipoDoc.BoundText)
                Me.TxtCodProducto.Enabled = False
                Me.TxtDescripcionProducto.Enabled = False
                Me.txtprecio.Enabled = False
                Me.CmdAgregar.Enabled = False
                Me.CmdQuitar.Enabled = False
                Call Resalta(TxtNumero_guia)
                Referencia = True
                Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
                Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
                Me.TxtCodProducto.Enabled = True
                Me.txtprecio.Enabled = True
                Me.CmdAgregar.Enabled = True
                Me.CmdQuitar.Enabled = True
            End If
    End If
End If
Set rst = Nothing
End Sub
Public Sub get_comprobante(ByVal i_venta As Double)
           Dim int_venta As Double
           strCadena = "SELECT id_venta,id_tipo_factura,id_vendedor FROM movimiento_venta where id_venta='" & i_venta & "' AND ruc='" & KEY_RUC & "'"
           Call ConfiguraRst(strCadena)
           If rst.RecordCount > 0 Then
                int_venta = rst("id_venta")
                Me.txttipofactura.Text = rst("id_tipo_factura")
                Me.DtcVendedor.BoundText = rst("id_vendedor")
                Call LlenarDatosCliente(int_venta)
                Call Llenar_Temporal(int_venta)
                Call llenarGrid_det(Me.HfdDetalle, Me.TxtNumeroDoc.Text, Me.TxtSerie.Text, Me.DtcTipoDoc.BoundText)
                Me.TxtCodProducto.Enabled = False
                Me.TxtDescripcionProducto.Enabled = False
                Me.txtprecio.Enabled = False
                Me.CmdAgregar.Enabled = False
                Me.CmdQuitar.Enabled = False
                Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True
                Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = False
                Me.TxtCodProducto.Enabled = True
                Me.txtprecio.Enabled = True
                Me.CmdAgregar.Enabled = True
                Me.CmdQuitar.Enabled = True
            End If
End Sub
Public Sub llenar_doc(ByVal Numero As String)
    Call llenarGrid_guardado(Me.HfdDetalle, Numero)
End Sub
Sub llenarGrid_guardado(ByVal Grilla As MSHFlexGrid, ByVal Numero As String)
Dim total As Double
Dim Descuento As Single
Dim SUBTOTAL As Double
Dim valor_venta As Double
Dim valor_igv As Double
Dim i As Integer

'StrCadena = "DELETE FROM Temporal_Venta_Guardado WHERE id_guardado='" & Numero & "' AND id_usuario='" & KEY_USUARIO & "'"
'Call EjecutaRST(StrCadena)
'Set RstEjecuta = Nothing
strCadena = "UPDATE Temporal_Ventas SET cDocumentoVenta='" & Trim(Me.TxtNumeroDoc.Text) & "' , sSerie='" & Trim(Me.TxtSerie.Text) & "' , doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' WHERE cDocumentoVenta='" & Trim(Numero) & "' AND id_usuario='" & KEY_USUARIO & "'"
Call EjecutaRST(strCadena)
Set RstEjecuta = Nothing
strCadena = "SELECT Temporal_Ventas.cProducto as Codigo,Producto.DescripcionProducto as Producto,Unidad.sAbreviatura as Unidad,Temporal_Ventas.Cantidad as Cantidad,Temporal_Ventas.Precio as Precio,Temporal_Ventas.Total as Total " & _
    "FROM Temporal_Ventas INNER JOIN (Producto INNER JOIN Unidad ON Producto.cunidad=Unidad.cunidad) ON Temporal_Ventas.cProducto=Producto.cProducto WHERE (Temporal_Ventas.cDocumentoVenta='" & Trim(Me.TxtNumeroDoc.Text) & "' AND Temporal_Ventas.doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND Temporal_Ventas.sSerie='" & Trim(Me.TxtSerie.Text) & "')"
On Error GoTo SALIR
  Call ConfiguraRst(strCadena)
  'Grilla.Clear
  'Grilla.Row = 0
  'Grilla.Rows = Rst.RecordCount
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 1600
  Grilla.ColWidth(1) = 5300
  Grilla.ColWidth(2) = 800
  Grilla.ColWidth(3) = 1500
  Grilla.ColWidth(4) = 1500
  Grilla.ColWidth(5) = 1600
Call DarFormato(Grilla, 3)
Call DarFormato(Grilla, 4)
Call DarFormato(Grilla, 5)
Me.LblCantidad.Caption = Trim(rst.RecordCount)
Set rst = Nothing

strCadena = "SELECT SUM(total) FROM Temporal_Ventas WHERE (cDocumentoVenta='" & Trim(Me.TxtNumeroDoc.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "')"
Call ConfiguraRst(strCadena)
total = rst(0)
Me.lblDescuento.Caption = Format(total * total_descuento / 100, "#,##0.00")
Descuento = Me.lblDescuento.Caption
Me.lblTotal.Caption = Format(rst(0) - Descuento, "#,##0.000")
Set rst = Nothing
valor_venta = 0
valor_igv = 0
 strCadena = "SELECT * FROM Temporal_Ventas WHERE  (cDocumentoVenta='" & Trim(Me.TxtNumeroDoc.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "')"
 Call ConfiguraRst(strCadena)
 rst.MoveFirst
  For i = 0 To rst.RecordCount - 1
    If (rst("igv") = "V") Then
        valor_venta = valor_venta + rst("Total") / 1.19
    Else
        valor_venta = valor_venta + rst("Total")
    End If
    rst.MoveNext
  Next i
 Set rst = Nothing
    SUBTOTAL = valor_venta
    igv = total - SUBTOTAL
    Me.LblIgv.Caption = Format(igv, "#,##0.000")
    Me.LblValorVenta.Caption = Format(SUBTOTAL, "#,##0.000")
    
    Me.txtcantidad.Text = 0
    Set rst = Nothing
  Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
  Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  Exit Sub
SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub ColocarDatosReferencia()
'On Error GoTo Error
'strCadena = "SELECT cDocumentoVenta,doc_cod,sSerie,FechaProceso FROM DocumentoVenta WHERE (cDocumentoVenta='" & Trim(Me.TxtNumero_guia.Text) & "' AND doc_cod='" & Trim(Me.DtcComprobanteGuia.BoundText) & "' AND sSerie='" & Trim(Me.TxtSeri_guia.Text) & "')"
'Call ConfiguraRst(strCadena)



'Me.DtpFechaReferencia.Value = rst(3)
'Set rst = Nothing

End Sub
Private Sub Llenar_Temporal(ByVal idVenta As Double)
Dim total_temp As Double
Dim rstTemporal As New ADODB.Recordset
Dim rstDetalle As New ADODB.Recordset
Dim i As Integer
strCadena = "SELECT * FROM movimiento_venta_detalle D WHERE id_venta='" & idVenta & "' AND ruc='" & KEY_RUC & "' ORDER BY id_detalle_venta DESC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
strCadena = "DELETE FROM temporal_ventas WHERE ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' "
CnBd.Execute (strCadena)
total_temp = 0
rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
    
    
        If Me.DtcTipoDoc.BoundText = "0007" Then ' nota de credito
        strCadena = "INSERT INTO temporal_ventas(ruc,id_alm,id_doc,id_serie,numero,id_producto,detalle,cantidad,precio,total,peso,dni_save,id_detalle_serie,serie,anio_fabricacion,nro_chasis,anio_modelo,nro_dua,nro_item) VALUES " & _
        "('" & KEY_RUC & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.TxtSerie.Text) & "','" & Me.TxtNumeroDoc.Text & "','" & rst("id_producto") & "','" & rst("detalle") & "','" & rst("cantidad") & "'," & _
        "'" & rst("precio") * -1 & " ','" & rst("total") * -1 & "','" & rst("peso") & "','" & KEY_USUARIO & "','" & rst("id_detalle_serie") & "','" & rst("serie") & "','" & rst("anio_fabricacion") & "','" & rst("nro_chasis") & "','" & rst("anio_modelo") & "','" & rst("nro_dua") & "','" & rst("nro_item") & "')"
        Else
            strCadena = "INSERT INTO temporal_ventas(ruc,id_alm,id_doc,id_serie,numero,id_producto,detalle,cantidad,precio,total,peso,dni_save,id_detalle_serie,serie,anio_fabricacion,nro_chasis,anio_modelo,nro_dua,nro_item) VALUES " & _
        "('" & KEY_RUC & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcTipoDoc.BoundText & "','" & Trim(Me.TxtSerie.Text) & "','" & Me.TxtNumeroDoc.Text & "','" & rst("id_producto") & "','" & rst("detalle") & "','" & rst("cantidad") & "'," & _
        "'" & rst("precio") & " ','" & rst("total") & "','" & rst("peso") & "','" & KEY_USUARIO & "','" & rst("id_detalle_serie") & "','" & rst("serie") & "','" & rst("anio_fabricacion") & "','" & rst("nro_chasis") & "','" & rst("anio_modelo") & "','" & rst("nro_dua") & "','" & rst("nro_item") & "')"
        End If
        
        CnBd.Execute (strCadena)
        total_temp = total_temp + rst("total")
        rst.MoveNext
    Next i
End If

 

End Sub
Private Sub TxtNumeroDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtSerie)
End If
If KeyCode = vbKeyRight Then
    Call Resalta(Me.TxtCodCliente)
End If
End Sub

Private Sub TxtNumeroDoc_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    
    Call mostrar_comprobante(Me.DtcTipoDoc.BoundText, Trim(Me.TxtSerie.Text), Trim(Me.TxtNumeroDoc.Text))
End If

 
 
 
End Sub

Public Sub mostrar_comprobante(ByVal n_doc As String, ByVal nserie As String, ByVal numerot As String)
    Dim montot As Double
    Dim idVenta As Double
    Dim nnumero As String

    Me.TxtNumeroDoc.Text = FormatosCeros(numerot, 6)
    Me.TxtSerie.Text = Format(nserie, "000")
    nnumero = Trim(Me.TxtNumeroDoc.Text)
    
    strCadena = "SELECT * FROM movimiento_venta WHERE (numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND ruc='" & KEY_RUC & "')"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        Call Nuevo
        
        Me.TxtNumeroDoc.Text = nnumero
        Exit Sub
   
    Else
        
        
        
        MsgBox "VISUALIZAR DOCUMENTO", vbInformation, KEY_EMPRESA
        Me.TxtIdVenta.Text = rst("id_venta")
        Me.chkVincular.Visible = True
        Me.lblregistradopor.Caption = "REG:  " & get_persona(rst("dni_save")) & Space(2) & Format(rst("hora"), "HH:mm:ss AM/PM")
         Me.txttipofactura.Text = rst("id_tipo_factura")
        If rst("id_recibo") > 0 Then
            strCadena = "SELECT serie,numero,id_vendedor,id_tipo_factura FROM movimiento_venta WHERE id_venta='" & rst("id_recibo") & "' AND ruc='" & KEY_RUC & "' ORDER BY id_venta DESC LIMIT 1"
            Call ConfiguraRstK(strCadena)
            Me.txtSerieRecibo.Text = rstK("serie")
            Me.txtNumeroRecibo.Text = rstK("numero")
            Me.DtcVendedor.BoundText = rstK("id_vendedor")
            
            Me.DtcVendedor.Locked = True
            Me.cmdGrabarRecibo.Visible = True
            Me.cmdImprimirRecibo.Visible = True
            Me.txtSerieRecibo.Visible = True
            Me.txtNumeroRecibo.Visible = True
            Me.txtSerieRecibo.Locked = False
            Me.txtNumeroRecibo.Locked = False
        End If
        
         If rst("id_doc") = "0007" Then
            strCadena = "SELECT documento FROM movimiento_venta WHERE id_venta='" & rst("id_comprobante") & "' "
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount > 0 Then
                Me.txtdocreferencia.Text = rstK("documento")
                Me.FrameReferencia.Visible = True
            End If
          Else
            Me.txtdocreferencia.Text = ""
            Me.FrameReferencia.Visible = False
        End If
        
        If rst("id_doc") = "0054" Then
            Me.lblContabilidad.Caption = rst("observacion")
        End If
        Me.lblContabilidad.Visible = False
        idVenta = rst("id_venta")
        montot = 0
        Me.TxtCodCliente.Text = rst("id_cliente")
       
      
        
        If rst("anulado") = "si" Then
            Me.lblAnulado.Visible = True
            Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
            Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True
            
            End If
        
                Me.TlbAcciones.Buttons("(Editable)").Enabled = False
                'Me.TlbAcciones.Buttons("(Pendiente)").Enabled = False
                Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
                Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True
                
            
        End If
        If rst("afecta_factura") = "si" Then
            Me.chk_factura.Value = 1
        Else
            Me.chk_factura.Value = 0
        End If
        
       
        
        Me.lblTotal.Caption = Format(rst("total"), "###0.00")
        Me.DtcAlmacen.BoundText = rst("id_alm")
        Me.lblPago.Caption = Format(rst("monto_pago"), "###0.00")
        Me.lblVuelto.Caption = Format(rst("monto_vuelto"), "###0.00")
        Me.lblExonerado.Caption = Format(rst("exonerado"), "###0.00")
        Me.LblValorVenta.Caption = Format(rst("valor_venta"), "#,##0.00")
        Me.LblIgv.Caption = Format(rst("igv"), "#,##0.00")
        Me.lblDescuento.Caption = Format(0, "#,##0.00")
        Me.DtcMoneda.BoundText = rst("id_moneda")
        Call LlenarDatosCliente(idVenta)
        Call llenarGrid_Comprobante(Me.HfdDetalle, idVenta)
        Call llena_pagosVenta(Me.HfgTipoPagos, idVenta)
        Me.TxtCodProducto.Enabled = False
        Me.TxtDescripcionProducto.Enabled = False
        Me.txtprecio.Enabled = False
        Me.CmdAgregar.Enabled = False
        Me.CmdQuitar.Enabled = False
        Me.txtcantidad.Enabled = False
        If Trim(Me.DtcTipoDoc.BoundText) = "0009" Then
            Me.TlbGrabar.Buttons(KEY_GUIAREMISION).Enabled = True
            ProcendenciaGuia = MostrarGuia
        Else
        Me.TlbGrabar.Buttons(KEY_GUIAREMISION).Enabled = False
        End If
        Me.TxtCodProducto.Enabled = False
        Me.HfdDetalle.SetFocus
End Sub
Private Sub VerificaAnulado(ByVal idVenta As Double)
strCadena = "Select Anulado FROM DocumentoVenta WHERE idVenta='" & idVenta & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    If Trim(rst(0)) = "V" Then
        Me.lblAnulado.Visible = True
        
        Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
        Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
    Else
        Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
    End If
End If
Set rst = Nothing
End Sub
Private Sub LLenarDatosReferencia(ByVal tipo_doc As String, ByVal Numero As String, ByVal serie As String)
strCadena = "SELECT * FROM movimiento_venta_targeta WHERE id_doc='" & tipo_doc & "' AND numero='" & Numero & "' AND serie='" & Trim(serie) & "' AND ruc='" & KEY_RUC & "'"
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
strCadena = "SELECT * FROM movimiento_venta WHERE id_venta='" & idVenta & "' AND ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
    
    Me.DtpActual.Value = CVDate(rst("fecha_emision"))
    Me.DtpFechaReferencia.Value = CVDate(rst("fecha_vencimiento"))
    CodPersona = rst("id_cliente")
    Me.DtcFormaPago.BoundText = rst("id_forma_pago")
    If Trim(CodPersona) = "00000000" Then
        Me.TxtCodCliente.Text = "00000000"
        Me.TxtCliente.Text = UCase(rst("ncliente"))
        Me.TxtDireccion.Text = KEY_DIR_PUBLIC
        Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
        Exit Sub
    End If
strCadena = "SELECT dni,nombre_completo,direccion FROM persona WHERE (dni ='" & CodPersona & "' )"
Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        If Trim(rst("dni")) <> "" Then
            Me.TxtCodCliente.Text = rst("dni")
        Else
            Me.TxtCodCliente.Text = rst("dni")
        End If
        Me.TxtCliente.Text = UCase(rst("nombre_completo"))
        Me.TxtDireccion.Text = UCase(rst("direccion"))
        
    End If
Set rst = Nothing
Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
End Sub

Private Sub TxtNumeroTargeta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Me.TxtNumeroTargeta.Text) <> "" Then
            Call Resalta(Me.txtOperacion)
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
    If Trim(Me.txtOperacion.Text) <> "" Then
        Call Resalta(Me.TxtMontoPagado)
    Else
        Call Resalta(Me.txtOperacion)
    End If
End If
End Sub

Private Sub TxtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then
        Call Resalta(Me.TxtDescripcionProducto)
       
End If
If KeyCode = vbKeyRight Then
     Me.CmdAgregar.SetFocus
End If
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
Dim TotalP As Single
If KeyAscii = 13 Then
        
        
        
        If Me.chkPrecios.Value = 1 And Me.HfPrecios.Rows > 0 Then
           If validar_precio(Format(Me.HfPrecios.TextMatrix(Me.HfPrecios.Row, 1), "###00.00"), Val(Me.txtprecio.Text)) = True Then
            Exit Sub
            End If
        Else
           If validar_precio(Val(Me.txtpreciooriginal.Text), Val(Me.txtprecio.Text)) = True Then
                Exit Sub
           End If
        End If
        TotalP = Val(Me.txtcantidad.Text) * Val(Me.txtprecio.Text)
        Me.LblTotalParcial.Caption = Format(TotalP, "###0.00")
        
        
        Call CmdAgregar_Click
        
End If
End Sub
Function validar_precio(ByVal precio_base As Single, ByVal precio_venta As Single) As Boolean
  
        
        If precio_venta < precio_base Then
            MsgBox "PRECIO INGRESADO NO ES VALIDO" + Chr(13) + Chr(13) + "Esta siendo NOTIFICADO", vbInformation, "VITEKEY"
            Me.txtprecio.Text = precio_base
            Call Resalta(Me.txtprecio)
            validar_precio = True
        Else
            validar_precio = False
        End If
    
End Function
Private Sub save_targeta(ByVal id_doc As String, ByVal Documento As String, ByVal serie As String, ByVal id_targeta As String, ByVal Numero As String)
    strCadena = "INSERT INTO DocumentoVenta_Targeta VALUES ('" & Trim(id_doc) & "','" & Trim(Documento) & "','" & Trim(serie) & "','" & id_targeta & "','" & Numero & "')"
    CnBd.Execute (strCadena)
End Sub
Private Function GeneraCodigoVenta(ByVal longitud As Integer) As String
Dim X As Integer
Dim rst_v As New ADODB.Recordset

strCadena = "SELECT intDocumentoVenta FROM DocumentoVenta WHERE doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND Alm_cod='" & Trim(Me.DtcAlmacen.BoundText) & "' AND sSerie='" & Trim(Me.TxtSerie.Text) & "' AND id_usuario='" & Trim(KEY_USUARIO) & "' ORDER BY intDocumentoVenta DESC "
        
rst_v.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic

        
Dim Formato As String
  Formato = ""
  For X = 1 To longitud
    Formato = Formato + "0"
  Next X
   
  If (rst_v.BOF And rst_v.EOF) Then
    StrNumero = Format(Str(Val(Formato) + 1), Formato)
  Else
    StrNumero = Format(Trim(Str(Val(Right(rst_v(0), longitud + 1))) + 1), Formato)
  End If
  Set rst = Nothing
  GeneraCodigoVenta = Gencodigo + StrNumero
  Gencodigo = ""
Set rst_v = Nothing
End Function
Public Sub Save()
On Error GoTo Error
Dim i As Integer, anul As String * 2, MontoActual As Double, TotalVenta As Double
Dim igv As Double, SUBTOTAL As Double, exonerado As Double, dfac As String, Monto_descuento As Single
Dim monto_pagado As Double, Monto_Vuelto As Double, Monto_Sobrante As Double, saldo_f As Double, estado_f As String
Dim id_venta  As Double, CodReferencia As String, KEY_VENCIMIENTO As String, cod_cliente As String, rst1 As New ADODB.Recordset, p As Integer
Dim horario As String, turno As String
Dim id_tipo_factura As String


If Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False Then
    Exit Sub
End If

If Trim(Me.DtcVendedor.Text) = "" Then
    MsgBox "Debe Seleccionar un Vendedor para Esta Operacion. " + Chr(13) + Chr(13) + "Gracias : " & KEY_VENDEDOR, vbInformation
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
    KEY_DETALLE = Trim(TxtObservacion.Text)
Else
    KEY_DETALLE = KEY_DETALLE
    KEY_MONTOENVASE = KEY_MONTOENVASE
    KEY_ENVASE = KEY_ENVASE
End If
If Trim(Me.TxtCodCliente.Text) = "" Then
    Me.TxtCodCliente.Text = "00000000"
End If
If Trim(Me.DtcTipoDoc.BoundText) = "0001" And Len(Me.TxtCodCliente.Text) <> 11 Then
    MsgBox "INGRESE RUC VALIDO PARA EL CLIENTE", vbInformation, KEY_EMPRESA
    Call Resalta(Me.TxtCodCliente)
    Exit Sub
End If

SUBTOTAL = Val(Me.LblValorVenta.Caption)
igv = Val(Me.LblIgv.Caption)
exonerado = Val(Me.lblExonerado.Caption)
TotalVenta = Val(Format(Me.lblTotal.Caption, "###0.000"))
Monto_descuento = Val(Me.lblDescuento.Caption)
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
        Exit Sub
    End If
    saldo_f = TotalVenta
    estado_f = "Credito"
Else
    saldo_f = KEY_NULO
End If
cod_cliente = Trim(Me.TxtCodCliente.Text)
Set rst1 = Nothing
strCadena = "SELECT * FROM movimiento_venta_monto_temporal WHERE numero='" & Trim(Me.TxtNumeroDoc.Text) & "' and id_usuario='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "' AND id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Me.TxtSerie.Text & "' ORDER BY id_monto ASC"
rst1.CursorLocation = adUseClient
rst1.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
If rst1.RecordCount > 0 Then
   rst1.MoveFirst
       'strCadena = "SELECT * FROM movimiento_venta WHERE numero='" & Trim(Me.TxtNumeroDoc.text) & "' AND serie='" & Trim(Me.TxtSerie.text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND ruc='" & KEY_RUC & "'"
       'Call ConfiguraRst(strCadena)
       
       If Val(Monto_Vuelto) < 0 Then
            MsgBox "El Monto Pagado es Inferior al Monto Total", vbInformation, "Mensaje para el Usuario"
            Call Resalta(Me.TxtMontoPagado)
            Exit Sub
        End If
        If Me.DtcFormaPago.BoundText = "05" Then
            KEY_VENCIMIENTO = Format(Me.DtpFechaReferencia.Value, "yyyy-mm-dd")
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
            
            
            Documento = Trim(Me.DtcTipoDoc.Text) & ":" & Trim(Me.TxtSerie.Text) & "-" & Trim(Me.TxtNumeroDoc.Text)
            
            'If Me.cmdSeriales.Visible = True Then
            'If KEY_TRAMITE = "si" Then
            'If trim = "si" Then
                'id_tipo_factura = "00002"
            'Else
                id_tipo_factura = Trim(Me.txttipofactura.Text)
                If Me.txteditable.Text = "si" Then
                   id_tipo_factura = "00003"
                End If
            'End If
            
            
            strCadena = "P_insert_venta_v2('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcFormaPago.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
            "'" & Trim(Me.TxtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Me.TxtCodCliente.Text & "','" & Me.TxtCliente.Text & "','" & SUBTOTAL & "','" & igv & "','" & exonerado & "','" & TotalVenta & "','" & Saldo & "'," & _
            "'" & Val(Me.lblPago.Caption) & "','" & Val(Me.lblVuelto.Caption) & "','" & KEY_FECHA & "','" & KEY_VENCIMIENTO & "','" & id_tipo_factura & "','" & Me.DtcVendedor.BoundText & "','" & KEY_USUARIO & "','" & KEY_CAMBIO & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Documento & "','" & horario & "','T','--','" & strconyugue & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
            
            Me.txtid_venta_ref = Val(Me.TxtIdVenta.Text)
            strCadena = "SELECT LAST_INSERT_ID() as ultimo"
            Call ConfiguraRstT(strCadena)
            id_venta = rstT(0)
            Me.TxtIdVenta.Text = id_venta
            
            
            If Me.DtcTipoDoc.BoundText = "0007" Then
                strCadena = "SELECT id_doc,fecha_emision,serie,numero FROM movimiento_venta WHERE id_venta='" & Val(Me.txtid_venta_ref.Text) & "'"
                Call ConfiguraRstT(strCadena)
                
                strCadena = "UPDATE movimiento_venta SET id_comprobante='" & Val(txtid_venta_ref.Text) & "',fecha_fact='" & Format(rstT("fecha_emision"), "YYYY-mm-dd") & "',id_doc_fact='" & rstT("id_doc") & "',serie_fact='" & rstT("serie") & "',numero_fact='" & rstT("numero") & "' WHERE id_venta='" & id_venta & "'"
                CnBd.Execute (strCadena)
            End If
            
            
            Call SaveDetalleDocumentoVenta(id_venta, Trim(Me.txteditable.Text))
     End If
        
            
        
        
        
        rst1.MoveFirst
        For k = 0 To rst1.RecordCount - 1
            strCadena = "INSERT INTO movimiento_venta_monto(id_venta,id_forma_pago,monto,monto_caja,id_tarjeta,id_tarjeta_numero,id_tarjeta_operacion,ruc)VALUES('" & id_venta & "','" & rst1("id_forma_pago") & "','" & rst1("monto") & "','" & rst1("monto_caja") & "','" & rst1("id_tarjeta") & "','" & rst1("id_tarjeta_numero") & "','" & rst1("id_tarjeta_operacion") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
               rst1.MoveNext
        Next k
       
        If Saldo > 0 And Val(Me.TxtCuotas.Text) > 0 Then
            strCadena = "SELECT * FROM movimiento_venta_cuotas_temporal WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Me.TxtSerie.Text & "' AND numero='" & Me.TxtNumeroDoc.Text & "' AND id_usuario='" & KEY_USUARIO & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
                rstT.MoveFirst
                For N = 0 To rstT.RecordCount - 1
                    strCadena = "INSERT INTO movimiento_venta_cuotas(id_cuota,id_venta,vencimiento,monto,saldo,ruc)VALUES('" & rstT("id_cuota") & "','" & id_venta & "','" & Format(rstT("vencimiento"), "YYYY-mm-dd") & "','" & rstT("monto") & "','" & rstT("saldo") & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                    rstT.MoveNext
                Next N
            End If
        End If
        
        
End If
StrNumero = FormatosCeros(Trim(Str(Val(Me.TxtNumeroDoc.Text)) + 1), 6)
strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE  id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "'  AND ruc='" & Trim(KEY_RUC) & "'"
CnBd.Execute (strCadena)
                Me.TxtCodProducto.Enabled = False
                Me.TxtDescripcionProducto.Enabled = False
                Me.txtcantidad.Enabled = False
                Me.txtprecio.Enabled = False
                Me.CmdAgregar.Enabled = False
                Me.CmdQuitar.Enabled = False
                chkPrecios.Enabled = False
                Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
                If (KEY_CARGO = "00001" Or KEY_CARGO = "00004") Then
                    Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True
                Else
                    Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
                End If
                Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
                Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
            
                dfactura = False
                Me.chk_factura.Value = 0
                Exit Sub
        Exit Sub
Error:
  MsgBox "Uso Incorrecto del Sistema", vbInformation, KEY_EMPRESA
  MsgBox "PULSE F2 PARA SALVAR LOS PRODUCTOS", vbInformation, KEY_EMPRESA
  
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
            "'" & Trim(Me.TxtSerie.Text) & "','" & Trim(Me.TxtNumeroDoc.Text) & "','" & Me.TxtCodCliente.Text & "','" & Me.TxtCliente.Text & "','" & Val(Me.LblValorVenta.Caption) & "','" & Val(Me.LblIgv.Caption) & "','0','" & rstT(0) & "','" & saldo1 & "'," & _
            "'" & Val(pago) & "','" & Val(vuelto) & "','" & KEY_FECHA & "','" & KEY_VENCIMIENTO & "','00001','" & KEY_USUARIO & "','" & KEY_CAMBIO & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            strCadena = "SELECT LAST_INSERT_ID() as ultimo"
            Call ConfiguraRstT(strCadena)
            id_venta = rstT(0)
            Call SaveDetalleDocumentoVenta(id_venta, Trim(Me.txteditable.Text))
            
                      
            strCadena = "P_insert_venta('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Me.DtcFormaPago.BoundText & "','" & Me.DtcMoneda.BoundText & "','" & delivery & "'," & _
            "'" & Trim(Me.TxtSerie.Text) & "','" & formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6) & "','" & Me.TxtCodCliente.Text & "','" & Me.TxtCliente.Text & "','0','0','" & Val(Me.lblExonerado.Caption) & "','" & rst(0) & "','" & saldo2 & "'," & _
            "'" & rst(0) & "','0','" & KEY_FECHA & "','" & KEY_VENCIMIENTO & "','00001','" & Me.DtcVendedor.BoundText & "','" & KEY_USUARIO & "','" & KEY_CAMBIO & "','" & dfac & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            strCadena = "SELECT LAST_INSERT_ID() as ultimo"
            Call ConfiguraRstT(strCadena)
            id_venta = rstT(0)
            Call SaveDetalleDocumentoVentaEspecial(id_venta, formato_item(Val(Me.TxtNumeroDoc.Text) + 1, 6))
            
            
StrNumero = FormatosCeros(Trim(Str(Val(Me.TxtNumeroDoc.Text)) + 2), 6)
strCadena = "UPDATE almacen_comprobante SET numero='" & StrNumero & "' WHERE id_alm='" & Trim(Me.DtcAlmacen.BoundText) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND serie='" & Trim(Me.TxtSerie.Text) & "'  AND ruc='" & Trim(KEY_RUC) & "'"
CnBd.Execute (strCadena)
                Me.TxtCodProducto.Enabled = False
                Me.TxtDescripcionProducto.Enabled = False
                Me.txtcantidad.Enabled = False
                Me.txtprecio.Enabled = False
                Me.CmdAgregar.Enabled = False
                Me.CmdQuitar.Enabled = False
                
                Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
                If (KEY_CARGO = "00001" Or KEY_CARGO = "00004") Then
                    Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True
                Else
                    Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
                End If
                Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = False
                Me.TlbGrabar.Buttons(KEY_PRINT).Enabled = True
            
                dfactura = False
                Me.chk_factura.Value = 0
                Exit Sub
            
End Sub
Private Sub SaveDetalleDocumentoVenta(ByVal idVenta As Double, ByVal in_editable As String)

Dim in_tipo_factura As String

Dim in_codigo As String
    
    If in_editable = "si" Then
                For i = 0 To Me.TxtCodProducto_per.Count - 1
                    If Trim(Me.txtdescripcion(i).Text) <> "-" And Trim(Me.txtdescripcion(i).Text) <> "" Then
                            If Trim(Me.TxtCodProducto_per(i).Text) <> "" Then
                                in_codigo = Trim(Me.TxtCodProducto_per(i).Text)
                            Else
                                in_codigo = "02484"
                            End If
                               strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,detalle,cantidad,precio,peso,total,ruc) VALUES ('" & idVenta & "','" & in_codigo & "','" & Trim(Me.txtdescripcion(i).Text) & "','" & Val(Me.txtCantidadPer(i).Text) & "','" & Val(Me.txtprecio_per(i).Text) & "','0','" & Val(Me.txttotal(i).Text) & "','" & KEY_RUC & "')"
                               CnBd.Execute (strCadena)
                    End If
                Next i
                Exit Sub
    Else
            strCadena = "SELECT * FROM temporal_ventas WHERE (numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_serie='" & Trim(Me.TxtSerie.Text) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "' and save='no')"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
               rstT.MoveFirst
               For i = 0 To rstT.RecordCount - 1
                   
                        strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,cantidad,precio,peso,total,serie,anio_fabricacion,nro_chasis,anio_modelo,nro_dua,nro_item,id_detalle_serie,detalle,ruc) VALUES ('" & idVenta & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("precio") & "','" & rstT("peso") & "','" & rstT("total") & "','" & rstT("serie") & "','" & rstT("anio_fabricacion") & "','" & rstT("nro_chasis") & "','" & rstT("anio_modelo") & "','" & rstT("nro_dua") & "','" & rstT("nro_item") & "','" & rstT("id_detalle_serie") & "','" & rstT("detalle") & "','" & KEY_RUC & "')"
                        CnBd.Execute (strCadena)
                   
                        strCadena = "UPDATE imp_producto_detalle SET vendido='si' WHERE nro_chasis='" & rstT("nro_chasis") & "' and ruc='" & KEY_RUC & "'"
                        CnBd.Execute (strCadena)
                   
                   rstT.MoveNext
                Next i
            End If
    End If
End Sub

Private Sub SaveDetalleDocumentoVentaRecibo(ByVal idVenta As Double)

   strCadena = "SELECT * FROM movimiento_venta_detalle WHERE id_venta='" & Val(Me.TxtIdVenta.Text) & "'"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
       rstT.MoveFirst
       For i = 0 To rstT.RecordCount - 1
           strCadena = "INSERT INTO movimiento_venta_detalle(id_venta,id_producto,cantidad,precio,peso,total,serie,anio_fabricacion,nro_chasis,anio_modelo,nro_dua,nro_item,ruc) VALUES ('" & idVenta & "','" & rstT("id_producto") & "','" & rstT("cantidad") & "','" & rstT("precio") & "','" & rstT("peso") & "','" & rstT("total") & "','" & rstT("serie") & "','" & rstT("anio_fabricacion") & "','" & rstT("nro_chasis") & "','" & rstT("anio_modelo") & "','" & rstT("nro_dua") & "','" & rstT("nro_item") & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           rstT.MoveNext
        Next i
    End If
End Sub

Private Sub SaveDetalleDocumentoVentaEspecial(ByVal idVenta As Double, ByVal Numero As String)

   strCadena = "SELECT * FROM temporal_ventas WHERE (numero='" & Trim(Numero) & "' AND id_doc='" & Trim(Me.DtcTipoDoc.BoundText) & "' AND id_serie='" & Trim(Me.TxtSerie.Text) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "')"
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

Private Sub TxtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtCodProducto)
End If
End Sub

Private Sub txtprecio_per_Change(Index As Integer)
If Val(Me.txtCantidadPer(Index).Text) = 0 Then
    Me.txttotal(Index).Text = Format(Me.txtprecio_per(Index).Text, "##00.00")
Else
    Me.txttotal(Index).Text = Format(Val(Me.txtprecio_per(Index).Text) * Val(Me.txtCantidadPer(Index).Text), "##00.00")
End If

  

End Sub

Private Sub TxtSeri_guia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSeri_guia.Text = FormatosCeros(Me.TxtSeri_guia.Text, 3)
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


Private Function ConvertirMes(ByVal Numero As Integer) As String
Select Case Numero
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

Private Sub TxtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(Me.TxtSerie.Text) <> "" Then
        Me.TxtSerie.Text = formato_item(Me.TxtSerie.Text, 3)
        strCadena = "SELECT * FROM almacen_comprobante WHERE id_doc='" & Me.DtcTipoDoc.BoundText & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            Me.TxtNumeroDoc.Text = rst("numero")
            KEY_APLICA_IGV = rst("igv")
            Me.TxtCodProducto.Enabled = True
            Call Resalta(Me.TxtCodProducto)
        Else
            MsgBox "SERIE NO ASIGNADA A ESTA SUCRSAL", vbInformation, KEY_EMPRESA
            Call Resalta(Me.TxtSerie)
            Exit Sub
        End If
    End If
End If
End Sub



Private Sub txttotal_Change(Index As Integer)
Dim tTotal As Double
For i = 0 To 4
    tTotal = Val(Me.txttotal(i).Text) + tTotal
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

Me.lblTotal.Caption = Format(tTotal, "###0.000")
Me.lblVuelto.Caption = Format(Val(Me.lblPago.Caption) - tTotal, "###0.000")
If KEY_APLICA_IGV = "si" Then
    SUBTOTAL = tTotal / (1 + KEY_IGV)
    igv = tTotal - SUBTOTAL
Else
     texonerado = tTotal + texonerado
    SUBTOTAL = 0
    igv = 0
End If

If texonerado > 0 Then
    Me.lblExonerado.Caption = Format(texonerado, "###0.000")
End If

Me.LblIgv.Caption = Format(igv, "###0.000")
Me.LblValorVenta.Caption = Format(SUBTOTAL, "###0.000")

Me.TlbGrabar.Buttons(KEY_SAVE).Enabled = True


End Sub

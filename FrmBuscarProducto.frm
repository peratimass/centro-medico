VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmComprasGastos 
   BorderStyle     =   0  'None
   Caption         =   "GASTOS COMPRA"
   ClientHeight    =   9225
   ClientLeft      =   945
   ClientTop       =   630
   ClientWidth     =   17205
   Icon            =   "FrmBuscarProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   17205
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtIdCompra 
      Height          =   285
      Left            =   12600
      TabIndex        =   59
      Text            =   "IdCompra"
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   240
      TabIndex        =   27
      Top             =   120
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "IMPORTACIONES"
      TabPicture(0)   =   "FrmBuscarProducto.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(5)=   "Frame6"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "OTROS GASTOS"
      TabPicture(1)   =   "FrmBuscarProducto.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Shape3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ChameleonBtn1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
         Height          =   375
         Left            =   15960
         TabIndex        =   90
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   ""
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmBuscarProducto.frx":0342
         PICN            =   "FrmBuscarProducto.frx":035E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LISTADO DE GASTOS ADICIONALES"
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
         Height          =   2895
         Left            =   720
         TabIndex        =   63
         Top             =   5880
         Width           =   15135
         Begin VB.CommandButton Command1 
            Caption         =   "REGENERAR"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   14040
            TabIndex        =   118
            Top             =   2160
            Width           =   855
         End
         Begin VitekeySoft.ChameleonBtn cmdupdate 
            Height          =   855
            Left            =   14040
            TabIndex        =   98
            Top             =   1200
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1508
            BTYPE           =   5
            TX              =   "Modificar"
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmBuscarProducto.frx":3212
            PICN            =   "FrmBuscarProducto.frx":322E
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
            Left            =   14040
            TabIndex        =   99
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1508
            BTYPE           =   5
            TX              =   "Eliminar"
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmBuscarProducto.frx":5867
            PICN            =   "FrmBuscarProducto.frx":5883
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshGastos 
            Height          =   2535
            Left            =   120
            TabIndex        =   103
            Top             =   240
            Width           =   13575
            _ExtentX        =   23945
            _ExtentY        =   4471
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
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "COMPROBANTE"
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
         Height          =   5295
         Left            =   720
         TabIndex        =   60
         Top             =   480
         Width           =   15135
         Begin VB.Frame frm_domiciliada 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NO DOMICILIADA"
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
            Height          =   1815
            Left            =   11400
            TabIndex        =   111
            Top             =   720
            Visible         =   0   'False
            Width           =   3615
            Begin VB.TextBox txtTotal 
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
               Left            =   1440
               TabIndex        =   117
               Top             =   1320
               Width           =   1575
            End
            Begin VB.TextBox txtRetencion 
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
               Left            =   1440
               TabIndex        =   116
               Top             =   840
               Width           =   1575
            End
            Begin VB.TextBox txtSubtotal 
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
               Left            =   1440
               TabIndex        =   115
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL :"
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
               TabIndex        =   114
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "RETENCION :"
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
               TabIndex        =   113
               Top             =   840
               Width           =   840
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SUB TOTAL :"
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
               Left            =   180
               TabIndex        =   112
               Top             =   480
               Width           =   795
            End
         End
         Begin VB.CheckBox chk_nodomiciliada 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "NO DOMICILIADA"
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
            Height          =   255
            Left            =   8640
            TabIndex        =   110
            Top             =   540
            Width           =   2655
         End
         Begin VB.TextBox txtBuscarresponsable 
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
            Left            =   8520
            MaxLength       =   80
            TabIndex        =   107
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkresponsable 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "RESPONSABLE :"
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
            Height          =   320
            Left            =   3600
            TabIndex        =   106
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.Frame frmRetencion 
            BackColor       =   &H008080FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   280
            Left            =   8640
            TabIndex        =   104
            Top             =   240
            Visible         =   0   'False
            Width           =   2655
            Begin VB.CheckBox chk_suspencion_retencion 
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               Caption         =   "SUSPENCION RETENCION"
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
               Height          =   255
               Left            =   0
               TabIndex        =   105
               Top             =   10
               Width           =   2415
            End
         End
         Begin VitekeySoft.ChameleonBtn CmdAgregar 
            Height          =   470
            Left            =   1800
            TabIndex        =   97
            Top             =   4680
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   820
            BTYPE           =   5
            TX              =   "AGREGAR GASTO"
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
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmBuscarProducto.frx":7CCD
            PICN            =   "FrmBuscarProducto.frx":7CE9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txtproducto 
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
            TabIndex        =   93
            Top             =   3360
            Width           =   5055
         End
         Begin VB.TextBox txtcodigoprod 
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
            Left            =   1800
            TabIndex        =   92
            Top             =   3360
            Width           =   1575
         End
         Begin VB.TextBox txtTipoCambio 
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
            Left            =   3960
            TabIndex        =   72
            Top             =   1440
            Width           =   1000
         End
         Begin VB.TextBox Txtdni 
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
            Left            =   1800
            TabIndex        =   65
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtmonto 
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
            Left            =   1800
            TabIndex        =   24
            Top             =   1080
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtpfecha 
            Height          =   325
            Left            =   1800
            TabIndex        =   25
            Top             =   2160
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
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
            Format          =   62390273
            CurrentDate     =   41126
         End
         Begin VB.TextBox txtserie 
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
            Height          =   315
            Left            =   5880
            TabIndex        =   22
            Top             =   240
            Width           =   1000
         End
         Begin VB.TextBox txtdescripcion 
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
            Height          =   525
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Top             =   4080
            Width           =   6735
         End
         Begin VB.TextBox txtnumero 
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
            Height          =   315
            Left            =   6960
            TabIndex        =   23
            Top             =   240
            Width           =   1455
         End
         Begin MSDataListLib.DataCombo DtcComprobante 
            Height          =   315
            Left            =   1800
            TabIndex        =   68
            Top             =   240
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
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
         Begin MSComCtl2.DTPicker DtpVencimiento 
            Height          =   330
            Left            =   4200
            TabIndex        =   70
            Top             =   2160
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
            Format          =   62390273
            CurrentDate     =   41126
         End
         Begin MSDataListLib.DataCombo DtcAfectoIgv 
            Height          =   315
            Left            =   1800
            TabIndex        =   73
            Top             =   1440
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   8388608
            ListField       =   ""
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
         Begin MSDataListLib.DataCombo DtcMoneda 
            Height          =   315
            Left            =   1800
            TabIndex        =   74
            Top             =   1800
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
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
         Begin MSDataListLib.DataCombo DtcTipoCompra 
            Height          =   315
            Left            =   1800
            TabIndex        =   79
            Top             =   2895
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
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
         Begin MSDataListLib.DataCombo DtcPeriodo 
            Height          =   315
            Left            =   1800
            TabIndex        =   94
            Top             =   2520
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
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
         Begin MSDataListLib.DataCombo DtcResponsable 
            Height          =   330
            Left            =   4980
            TabIndex        =   108
            Top             =   1080
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
         Begin VB.Label lblidCompra 
            BackColor       =   &H008080FF&
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
            Height          =   495
            Left            =   3720
            TabIndex        =   109
            Top             =   4680
            Width           =   1815
         End
         Begin VB.Label lblcuenta_detalle 
            BackColor       =   &H008080FF&
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
            Height          =   255
            Left            =   4560
            TabIndex        =   96
            Top             =   3720
            Width           =   3975
         End
         Begin VB.Label lblcuenta_contable 
            BackColor       =   &H008080FF&
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
            Height          =   255
            Left            =   3480
            TabIndex        =   95
            Top             =   3720
            Width           =   975
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SERVICIO :"
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
            TabIndex        =   91
            Top             =   3345
            Width           =   690
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TIPO GASTO :"
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
            Left            =   360
            TabIndex        =   80
            Top             =   3000
            Width           =   870
         End
         Begin VB.Label Label33 
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
            Left            =   240
            TabIndex        =   78
            Top             =   4200
            Width           =   990
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PERIODO :"
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
            TabIndex        =   77
            Top             =   2640
            Width           =   690
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "VENCIMI:"
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
            TabIndex        =   76
            Top             =   2200
            Width           =   630
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "T.C :"
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
            Left            =   3600
            TabIndex        =   75
            Top             =   1440
            Width           =   285
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AFECTO IGV :"
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
            TabIndex        =   71
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA  :"
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
            Left            =   690
            TabIndex        =   69
            Top             =   2200
            Width           =   540
         End
         Begin VB.Label lblcliente 
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
            Left            =   3600
            TabIndex        =   67
            Top             =   720
            Width           =   4875
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RUC/DNI  :"
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
            Left            =   525
            TabIndex        =   66
            Top             =   720
            Width           =   705
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MONTO  :"
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
            Left            =   600
            TabIndex        =   64
            Top             =   1080
            Width           =   630
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MONEDA :"
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
            TabIndex        =   62
            Top             =   1800
            Width           =   690
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DOCUMENTO :"
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
            TabIndex        =   61
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PRODUCTOS IMPORTACION"
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
         Height          =   1695
         Left            =   -74520
         TabIndex        =   58
         Top             =   6600
         Width           =   12975
         Begin VitekeySoft.ChameleonBtn cmdProcesar 
            Height          =   975
            Left            =   10560
            TabIndex        =   100
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
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
            MICON           =   "FrmBuscarProducto.frx":B331
            PICN            =   "FrmBuscarProducto.frx":B34D
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
            Left            =   11760
            TabIndex        =   101
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
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
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmBuscarProducto.frx":E995
            PICN            =   "FrmBuscarProducto.frx":E9B1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfVinculados 
            Height          =   1335
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   2355
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TOTALES"
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
         Height          =   1695
         Left            =   -67680
         TabIndex        =   53
         Top             =   4800
         Width           =   6135
         Begin VB.TextBox txtigvcomision 
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
            Left            =   2775
            TabIndex        =   21
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtcomision 
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
            Left            =   2775
            TabIndex        =   20
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lbltotalgeneral 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   2760
            TabIndex        =   57
            Top             =   960
            Width           =   1485
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IGV  :"
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
            Left            =   1785
            TabIndex        =   56
            Top             =   600
            Width           =   405
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "COMISION  :"
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
            Left            =   1260
            TabIndex        =   55
            Top             =   240
            Width           =   930
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL GENERAL  :"
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
            Left            =   840
            TabIndex        =   54
            Top             =   960
            Width           =   1350
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DATOS IMPORTACION"
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
         Height          =   1935
         Left            =   -74400
         TabIndex        =   49
         Top             =   4560
         Width           =   5895
         Begin VB.CheckBox chkConvertirSoles 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Caption         =   "CONV.MN"
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
            Left            =   2760
            TabIndex        =   87
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox txtFleteImportacion 
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
            Left            =   2775
            TabIndex        =   85
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtseguro 
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
            Left            =   2775
            TabIndex        =   81
            Top             =   825
            Width           =   1455
         End
         Begin VB.TextBox txttc 
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
            Left            =   615
            TabIndex        =   18
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtfob 
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
            Left            =   2775
            TabIndex        =   17
            Top             =   390
            Width           =   1455
         End
         Begin VB.TextBox txtcif 
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
            Left            =   2775
            TabIndex        =   19
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FLETE :"
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
            Left            =   1560
            TabIndex        =   86
            Top             =   1200
            Width           =   540
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SEGURO :"
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
            Left            =   1365
            TabIndex        =   82
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "T.C  :"
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
            TabIndex        =   52
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FOB  :"
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
            Left            =   1635
            TabIndex        =   51
            Top             =   480
            Width           =   450
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C.I.F  :"
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
            Left            =   1560
            TabIndex        =   50
            Top             =   1560
            Width           =   525
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "GASTOS"
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
         Height          =   4215
         Left            =   -67680
         TabIndex        =   39
         Top             =   480
         Width           =   6135
         Begin VB.TextBox TxtNotadebito 
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
            Left            =   2760
            TabIndex        =   88
            Top             =   3480
            Width           =   1455
         End
         Begin VB.TextBox txtcerfgases 
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
            Left            =   2760
            TabIndex        =   16
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox txtalmacenaje 
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
            Left            =   2760
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtgastosoperativos 
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
            Left            =   2760
            TabIndex        =   9
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox txtterminal 
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
            Left            =   2760
            TabIndex        =   10
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtgremiosmaritimos 
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
            Left            =   2760
            TabIndex        =   11
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtciasnaviera 
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
            Left            =   2760
            TabIndex        =   12
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtconduccion 
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
            Left            =   2760
            TabIndex        =   13
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtgastosdepacho 
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
            Left            =   2760
            TabIndex        =   14
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox txtflete 
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
            Left            =   2760
            TabIndex        =   15
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N.DEBITO :"
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
            Left            =   1365
            TabIndex        =   89
            Top             =   3600
            Width           =   825
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL GASTOS :"
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
            Left            =   840
            TabIndex        =   84
            Top             =   3880
            Width           =   1350
         End
         Begin VB.Label lbltotalgastos 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2790
            TabIndex        =   83
            Top             =   3840
            Width           =   1455
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ALMACENAJE DESCARGA :"
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
            Left            =   255
            TabIndex        =   48
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GASTOS OPERATIVOS :"
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
            Left            =   480
            TabIndex        =   47
            Top             =   720
            Width           =   1710
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TERMINAL :"
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
            Left            =   1335
            TabIndex        =   46
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GREMIOS MARITIMOS :"
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
            Left            =   480
            TabIndex        =   45
            Top             =   1440
            Width           =   1710
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CIAS NAVIERAS GASTOS :"
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
            Left            =   285
            TabIndex        =   44
            Top             =   1800
            Width           =   1905
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CONDUCCION :"
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
            Left            =   1050
            TabIndex        =   43
            Top             =   2160
            Width           =   1140
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GASTOS DESPACHO :"
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
            Left            =   630
            TabIndex        =   42
            Top             =   2520
            Width           =   1560
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FLETE :"
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
            Left            =   1650
            TabIndex        =   41
            Top             =   2880
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GASTOS VARIOS :"
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
            Left            =   870
            TabIndex        =   40
            Top             =   3240
            Width           =   1320
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DERECHOS"
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
         Height          =   4095
         Left            =   -74400
         TabIndex        =   28
         Top             =   480
         Width           =   5895
         Begin VB.TextBox txtderechoespecifico 
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
            Left            =   2760
            TabIndex        =   7
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox txtpercepcionigv 
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
            Left            =   2760
            TabIndex        =   6
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txtservdespacho 
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
            Left            =   2760
            TabIndex        =   5
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txtipm 
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
            Left            =   2760
            TabIndex        =   4
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtigv 
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
            Left            =   2760
            TabIndex        =   3
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtisc 
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
            Left            =   2760
            TabIndex        =   2
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtsobretasa 
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
            Left            =   2760
            TabIndex        =   1
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtadvalorem 
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
            Left            =   2760
            TabIndex        =   0
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lbltotalderechos 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   2760
            TabIndex        =   38
            Top             =   3600
            Width           =   1485
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "TOTAL DERECHOS :"
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
            Left            =   570
            TabIndex        =   37
            Top             =   3600
            Width           =   1545
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DERECHO ESPECIFICO :"
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
            Left            =   360
            TabIndex        =   36
            Top             =   2880
            Width           =   1770
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PERCEPCION IGV :"
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
            Left            =   765
            TabIndex        =   35
            Top             =   2520
            Width           =   1365
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SERV. DESPACHO :"
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
            Left            =   735
            TabIndex        =   34
            Top             =   2160
            Width           =   1395
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "IPM :"
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
            Left            =   1755
            TabIndex        =   33
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label Label4 
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
            Left            =   1770
            TabIndex        =   32
            Top             =   1440
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ISC :"
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
            Left            =   1770
            TabIndex        =   31
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SOBRETASA :"
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
            Left            =   1140
            TabIndex        =   30
            Top             =   720
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AD VALOREM :"
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
            Left            =   1065
            TabIndex        =   29
            Top             =   360
            Width           =   1065
         End
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   8535
         Left            =   120
         Top             =   360
         Width           =   16335
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   8445
         Left            =   -74880
         Top             =   375
         Width           =   16455
      End
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   9225
      Left            =   0
      Top             =   0
      Width           =   17205
   End
End
Attribute VB_Name = "FrmComprasGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Dim totalderechos As Double
Dim totalgeneral As Double
Private Sub llenarImportacion(ByVal Grilla As MSHFlexGrid, ByVal id_compra As Double)
On Error GoTo salir
strCadena = "SELECT * FROM movimiento_compra_importacion WHERE id_compra='" & id_compra & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txtadvalorem.Text = Format(rst("ad_valorem"), "###0.00")
    Me.txtsobretasa.Text = Format(rst("sobretasa"), "###0.00")
    Me.txtisc.Text = Format(rst("isc"), "###0.00")
    Me.TxtIgv.Text = Format(rst("igv"), "###0.00")
    Me.txtipm.Text = Format(rst("ipm"), "###0.00")
    Me.txtservdespacho.Text = Format(rst("serv_despacho"), "###0.00")
    Me.txtpercepcionigv.Text = Format(rst("percepcion_igv"), "###0.00")
    Me.txtderechoespecifico.Text = Format(rst("derecho_especifico"), "###0.00")
    Me.txtalmacenaje.Text = Format(rst("almacenaje_descarga"), "###0.00")
    Me.txtgastosoperativos.Text = Format(rst("gastos_operativos"), "###0.00")
    Me.txtterminal.Text = Format(rst("terminal"), "###0.00")
    Me.txtgremiosmaritimos.Text = Format(rst("gremios_maritimos"), "###0.00")
    Me.txtciasnaviera.Text = Format(rst("cias_navieras_gastos"), "###0.00")
    Me.txtconduccion.Text = Format(rst("conduccion"), "###0.00")
    Me.txtgastosdepacho.Text = Format(rst("gastos_despacho"), "###0.00")
    Me.TxtFlete.Text = Format(rst("flete"), "###0.00")
    Me.txtcerfgases.Text = Format(rst("certifi_gases"), "###0.00")
    Me.txtcomision.Text = Format(rst("comision"), "###0.00")
    Me.txtigvcomision.Text = Format(rst("igv_comision"), "###0.00")
    Me.txtFob.Text = Format(rst("fob"), "###0.00")
    Me.txtSeguro.Text = Format(rst("seguro"), "###0.00")
    Me.TxtCif.Text = Format(rst("cif"), "###0.00")
    Me.txtTc.Text = Format(rst("tc"), "###0.00")
  Else
    Me.txtTc.Text = Format(KEY_CAMBIO, "###0.00")
End If

Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Private Sub ChameleonBtn1_Click()

Call enabled_form(FrmCompras)
Call FrmCompras.otros_gastos(Val(FrmCompras.txtIdCompra.Text))

Unload Me
Exit Sub
End Sub

Private Sub chk_nodomiciliada_Click()
If Me.chk_nodomiciliada.Value = 1 Then
   Me.frm_domiciliada.Visible = True
Else
   Me.frm_domiciliada.Visible = False
End If
End Sub

Private Sub chkConvertirSoles_Click()
If Me.chkConvertirSoles.Value = 1 Then
   Me.txtFob.Text = Val(Me.txtFob.Text) * Val(Me.txtTc.Text)
   Me.txtSeguro.Text = Val(Me.txtSeguro.Text) * Val(Me.txtTc.Text)
   Me.TxtFlete.Text = Val(Me.TxtFlete.Text) * Val(Me.txtTc.Text)
 Else
   Me.txtFob.Text = Val(Me.txtFob.Text) / Val(Me.txtTc.Text)
   Me.txtSeguro.Text = Val(Me.txtSeguro.Text) / Val(Me.txtTc.Text)
   Me.TxtFlete.Text = Val(Me.TxtFlete.Text) / Val(Me.txtTc.Text)
End If
End Sub

Private Sub cmdagregar_Click()
Dim cod_identidad As String * 1
Dim valor_venta As Double
Dim igv As Double
Dim Total As Double
in_nodomiciliada = "no"
in_retencion = 0
If Me.chk_nodomiciliada.Value = 1 Then
   in_nodomiciliada = "si"
End If

If Me.cmdagregar.Caption = "Modificar" Then
    strCadena = "UPDATE movimiento_compra_gasto SET no_domiciliada='" & in_nodomiciliada & "',id_doc='" & Me.DtcComprobante.BoundText & "',serie='" & Me.TxtSerie.Text & "',numero='" & Me.txtNumero.Text & "', id_persona='" & Me.txtDni.Text & "' WHERE id_gasto='" & Me.mshGastos.TextMatrix(Me.mshGastos.Row, 0) & "' ANd ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    Call llenar_gastos(Me.mshGastos, Val(Me.txtIdCompra.Text))
    Me.cmdagregar.Caption = "Agregar"
    Me.TxtSerie.Text = ""
    Me.txtNumero.Text = ""
    Me.txtDni.Text = ""
    Me.lblcliente.Caption = ""
    Me.txtMonto.Text = 0#
    Me.Dtpfecha.Value = Me.Dtpfecha.Value
    Me.txtDescripcion.Text = ""
    Me.DtcComprobante.SetFocus
    Exit Sub
End If


strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Me.DtcComprobante.BoundText & "' AND serie='" & Me.TxtSerie.Text & "' AND numero='" & Me.txtNumero.Text & "' AND id_proveedor='" & Me.txtDni.Text & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
If Len(Trim(Me.txtDni.Text)) = 8 Then
    cod_identidad = 1
End If
If Len(Trim(Me.txtDni.Text)) = 11 Then
    cod_identidad = 6
End If
If Len(Trim(Me.txtDni.Text)) <> 8 And Len(Trim(Me.txtDni.Text)) <> 11 Then
    cod_identidad = 0
End If

If Me.DtcAfectoIgv.BoundText = "SI" Then
    exonerado = 0
    valor_venta = Val(Me.txtMonto.Text) / (KEY_IGV + 1)
    igv = Val(Me.txtMonto.Text) - valor_venta
Else
    
    
        If Me.DtcComprobante.BoundText = "0002" Then
            igv = 0
            exonerado = 0
            valor_venta = Val(Me.txtMonto.Text)
        Else
            igv = 0
            exonerado = Val(Me.txtMonto.Text)
            valor_venta = 0
        End If
        
   
    
End If


        
        If Me.DtcMoneda.BoundText = "00001" Then
                in_cta_compra = KEY_CTA_COMPRA_SOLES
           Else
                in_cta_compra = KEY_CTA_COMPRA_SOLES
           End If
           
            If Me.DtcComprobante.BoundText = "0002" Then
             in_cta_compra = KEY_CTA_COMPRA_RH
        End If
        
        
        If KEY_CONTABILIDAD = "si" Then
           If put_verifica_cuenta_contable(Me.DtcComprobante.BoundText, Trim(Me.TxtSerie.Text), Trim(Me.txtNumero.Text), in_cta_compra, Me.DtcTipoCompra.BoundText) = False Then
              Exit Sub
           End If
        End If
       
      
        If Me.chk_nodomiciliada.Value = 1 Then
            exonerado = 0
            valor_venta = Val(Me.txtSubtotal.Text)
            igv = 0
            Me.txtMonto.Text = Val(Me.TxtTotal.Text)
            in_retencion = Val(Me.txtRetencion.Text)
        
        End If
        
        If Me.DtcTipoCompra.BoundText = "01" Then
            exonerado = 0
            igv = 0
            valor_venta = Val(Me.txtMonto.Text)
           'MsgBox "Gasttos Importacion : Valor Venta:0.00, IGV :" & Format(igv, "#,##0.00") & " TOTAL :0.00", vbInformation, KEY_VENDEDOR
            
       End If
       
        
        If KEY_PAIS = KEY_PERU Then
            strCadena = "call P_insert_compra_ultimate('" & Me.DtcComprobante.BoundText & "','" & KEY_ALM & "','" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "','" & Format(DtpVencimiento.Value, "YYYY-mm-dd") & "','02'," & _
            "'" & Me.DtcTipoCompra.BoundText & "','--','" & Me.DtcMoneda.BoundText & "','" & formato_item(Month(Me.Dtpfecha.Value), 2) & "','" & Year(Me.Dtpfecha.Value) & "','" & Trim(Me.TxtSerie.Text) & "'," & _
            "'" & Format(Trim(Me.txtNumero.Text), "00000000") & "','" & cod_identidad & "','" & Trim(Me.txtDni.Text) & "','" & UCase(Me.lblcliente.Caption) & "','" & Trim(Me.TxtTipoCambio.Text) & "'," & _
            "'0','" & valor_venta & "','" & igv & "','0','0','0','" & in_retencion & "','" & exonerado & "','0','" & Val(Me.txtMonto.Text) & "','" & Val(Me.txtMonto.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtDescripcion.Text) & "','02','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & Me.DtcResponsable.BoundText & "','0','0','0','0','" & KEY_RUC & "')"
        Else
        
            strCadena = "call P_insert_compra_ultimate_internacional('" & Me.DtcComprobante.BoundText & "','" & KEY_ALM & "','" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "','" & Format(DtpVencimiento.Value, "YYYY-mm-dd") & "','02'," & _
            "'" & Me.DtcTipoCompra.BoundText & "','--','" & Me.DtcMoneda.BoundText & "','" & formato_item(Month(Me.Dtpfecha.Value), 2) & "','" & Year(Me.Dtpfecha.Value) & "','" & Trim(Me.TxtSerie.Text) & "'," & _
            "'" & Format(Trim(Me.txtNumero.Text), "00000000") & "','" & cod_identidad & "','" & Trim(Me.txtDni.Text) & "','" & UCase(Me.lblcliente.Caption) & "','" & Trim(Me.TxtTipoCambio.Text) & "'," & _
            "'0','" & valor_venta & "','" & igv & "','0','0','0','" & in_retencion & "','" & exonerado & "','0','" & Val(Me.txtMonto.Text) & "','" & Val(Me.txtMonto.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtDescripcion.Text) & "','02','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & Me.DtcResponsable.BoundText & "','0','0','0','0','" & KEY_RUC & "')"
        End If
        Call ConfiguraRstP(strCadena)
        id_compra = rstP(0)
        
        strCadena = "UPDATE movimiento_compra SET fecha_kardex='" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "' WHERE id_compra='" & id_compra & "' LIMIT 1"
        CnBd.Execute (strCadena)
        
        
        
        If Me.chk_nodomiciliada.Value = 1 Then
            strCadena = "UPDATE movimiento_compra SET no_domiciliada='si' WHERE id_compra='" & id_compra & "'"
            CnBd.Execute (strCadena)
        End If
        
        
        
        If Me.DtcAfectoIgv.BoundText = "SI" Then
            in_afecto = "si"
            in_exonerado = 0
            valor_venta = Val(Me.txtMonto.Text) / (KEY_IGV + 1)
            igv = Val(Me.txtMonto.Text) - valor_venta
        Else
            
            
            in_afecto = "no"
            igv = 0
            in_exonerado = Val(Me.txtMonto.Text)
            valor_venta = 0
    
        End If
        
        
        If Me.chk_nodomiciliada.Value = 1 Then
            in_exonerado = 0
            valor_venta = Val(Me.txtSubtotal.Text)
            igv = 0
            Me.txtMonto.Text = valor_venta
            in_retencion = Val(Me.txtRetencion.Text)
        End If
        
        If Me.DtcTipoCompra.BoundText = "01" Then
            exonerado = 0
            igv = 0
            valor_venta = Val(Me.txtMonto.Text)
        End If
        
        
        strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,detalle,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,ivap,otros,percepcion, " & _
        "valor_venta,exonerado,total,p_venta,p_costo,id_alm,retencion,ruc) VALUES ('" & id_compra & "','" & Trim(Me.txtcodigoprod.Text) & "','" & Me.lblcuenta_detalle.Caption & "','1','" & Val(Me.txtMonto.Text) & "'," & _
        "'0','0','0','" & valor_venta & "','0','" & igv & "', " & _
        "'0','0','0','" & valor_venta & "','" & in_exonerado & "','" & Val(Me.txtMonto.Text) & "','" & Val(Me.txtMonto.Text) & "','" & get_precio_costo(Trim(Me.txtcodigoprod.Text)) & "','" & KEY_ALM & "','" & in_retencion & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
        

        strCadena = "INSERT INTO movimiento_compra_gasto(id_compra,id_persona,id_doc,serie,numero,monto,fecha,descripcion,tc,id_moneda,id_compra_gasto,afecto_igv,ruc)VALUES " & _
        " ('" & Val(Me.txtIdCompra.Text) & "','" & Me.txtDni.Text & "','" & Me.DtcComprobante.BoundText & "','" & Me.TxtSerie.Text & "','" & Me.txtNumero.Text & "','" & Val(Me.txtMonto.Text) & "','" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "','" & Me.txtDescripcion.Text & "','" & Val(Me.TxtTipoCambio.Text) & "','" & Me.DtcMoneda.BoundText & "','" & id_compra & "','" & in_afecto & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)


        
'02----------------guardar en detalle documento Compra-----------
 If KEY_CONTABILIDAD = "si" And Me.DtcComprobante.BoundText <> "0089" Then
    If KEY_PAIS = KEY_PERU Then
        'strCadena = "call p_insert_compra_emitido_ii('" & id_compra & "')"
        
            strCadena = "call p_insert_compra_emitido_premiun('" & id_compra & "')"
        
           
        
    Else
         'strCadena = "call p_insert_compra_emitido_internacional('" & id_compra & "')"
         strCadena = "call p_insert_compra_emitido_internacional_gasto('" & id_compra & "')"
         
    End If
    Call Execute_Sql(strCadena)
 End If
 
        
        
        
        


 Me.lblidCompra.Caption = get_periodo_detalle(Me.DtcPeriodo.BoundText, id_compra)
MsgBox "NUMERO DE VOUCHER GENERADO  : " & Trim(Me.lblidCompra.Caption), vbInformation, KEY_VENDEDOR
    

Call llenar_gastos(Me.mshGastos, Val(Me.txtIdCompra.Text))
Me.TxtSerie.Text = ""
Me.txtNumero.Text = ""
Me.txtDni.Text = ""
Me.lblcliente.Caption = ""
Me.txtMonto.Text = 0#
Me.Dtpfecha.Value = KEY_FECHA
Me.txtDescripcion.Text = ""
Me.DtcComprobante.SetFocus
Me.lblcuenta_contable.Caption = ""
Me.lblcuenta_detalle.Caption = ""



Exit Sub
Else
    id_compra = rst("id_compra")
    If MsgBox("COMPROBANTE YA REGISTRADO, DESEA COMPARTIR GASTOS ", vbInformation + vbYesNo, KEY_EMPRESA) = vbYes Then
        strCadena = "INSERT INTO movimiento_compra_gasto(id_compra,id_persona,id_doc,serie,numero,monto,fecha,descripcion,tc,id_moneda,id_compra_gasto,afecto_igv,ruc)VALUES " & _
        " ('" & Val(Me.txtIdCompra.Text) & "','" & Me.txtDni.Text & "','" & Me.DtcComprobante.BoundText & "','" & Me.TxtSerie.Text & "','" & Me.txtNumero.Text & "','" & Val(Me.txtMonto.Text) & "','" & Format(Me.Dtpfecha.Value, "YYYY-mm-dd") & "','" & Me.txtDescripcion.Text & "','" & Val(Me.TxtTipoCambio.Text) & "','" & Me.DtcMoneda.BoundText & "','" & id_compra & "','" & in_afecto & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        Call llenar_gastos(Me.mshGastos, Val(Me.txtIdCompra.Text))
    End If
End If
Set rst = Nothing
End Sub




Private Sub cmddelete_Click()
If MsgBox("Esta Seguro de eliminar este Registro", vbYesNo + vbQuestion, KEY_EMPRESA) = vbYes Then
            
            strCadena = "SELECT * FROM movimiento_compra_gasto WHERE id_gasto='" & Val(Me.mshGastos.TextMatrix(Me.mshGastos.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
                strCadena = "DELETE FROM movimiento_compra_gasto WHERE id_gasto='" & Val(Me.mshGastos.TextMatrix(Me.mshGastos.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                
                strCadena = "Call CON_Asiento_EliminarCompra('" & rstT("id_compra_gasto") & "', '" & KEY_USUARIO & "', '" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                strCadena = "DELETE FROM movimiento_compra WHERE id_compra='" & rstT("id_compra_gasto") & "' and ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                
            End If
            
             
        End If
        Call llenar_gastos(Me.mshGastos, Val(Me.txtIdCompra.Text))
End Sub

Private Sub cmdProcesar_Click()
  strCadena = "SELECT * FROM movimiento_compra_importacion WHERE id_compra='" & Val(Me.txtIdCompra.Text) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "INSERT INTO movimiento_compra_importacion (id_compra,ruc)VALUES('" & Val(Me.txtIdCompra.Text) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
        End If
        strCadena = "UPDATE movimiento_compra_importacion SET ad_valorem='" & Val(Me.txtadvalorem.Text) & "',sobretasa='" & Val(Me.txtsobretasa.Text) & "',isc='" & Val(Me.txtisc.Text) & "'," & _
        " igv='" & Val(Me.TxtIgv.Text) & "',ipm='" & Val(Me.txtipm.Text) & "',serv_despacho='" & Val(Me.txtservdespacho.Text) & "',percepcion_igv='" & Val(Me.txtpercepcionigv.Text) & "'," & _
        "derecho_especifico='" & Val(Me.txtderechoespecifico.Text) & "',almacenaje_descarga='" & Val(Me.txtalmacenaje.Text) & "',gastos_operativos='" & Val(Me.txtgastosoperativos.Text) & "'," & _
        "terminal='" & Val(Me.txtterminal.Text) & "',gremios_maritimos='" & Val(Me.txtgremiosmaritimos.Text) & "',cias_navieras_gastos='" & Val(Me.txtciasnaviera.Text) & "'," & _
        "conduccion='" & Val(Me.txtconduccion.Text) & "',gastos_despacho='" & Val(Me.txtgastosdepacho.Text) & "',flete='" & Val(Me.TxtFlete.Text) & "',certifi_gases='" & Val(Me.txtcerfgases.Text) & "'," & _
        "comision='" & Val(Me.txtcomision.Text) & "',igv_comision='" & Val(Me.txtigvcomision.Text) & "',tc='" & Val(Me.txtTc.Text) & "',fob='" & Val(Me.txtFob.Text) & "',seguro='" & Val(Me.txtSeguro.Text) & "',cif='" & Val(Me.TxtCif.Text) & "' WHERE id_compra='" & Val(Me.txtIdCompra.Text) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        
End Sub

Private Sub cmdSalir_Click()
Call enabled_form(FrmCompras)
Unload Me

End Sub

Private Sub cmdupdate_Click()
If MsgBox("Esta Seguro de Modificar este Registro", vbYesNo + vbQuestion, KEY_EMPRESA) = vbYes Then
           strCadena = "SELECT * FROM movimiento_compra WHERE id_gasto='" & Val(Me.mshGastos.TextMatrix(Me.mshGastos.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
           Call ConfiguraRst(strCadena)
           If rst.RecordCount > 0 Then
                Me.DtcComprobante.BoundText = rst("id_doc")
                Me.TxtSerie.Text = rst("serie")
                Me.txtNumero.Text = rst("numero")
                Me.txtDni.Text = rst("id_persona")
                Me.txtMonto.Text = Format(rst("monto"), "#,##0.00")
                Me.Dtpfecha.Value = rst("fecha")
                Me.txtDescripcion.Text = rst("descripcion")
                Me.lblcliente.Caption = BDBuscarCampo("persona", "nombre_completo", "dni", Trim(Me.txtDni.Text))
                Me.cmdagregar.Caption = "Modificar"
                Me.cmdagregar.SetFocus
            
                
           End If
        End If
End Sub

Private Sub Command1_Click()

strCadena = "SELECT * FROM movimiento_compra_gasto WHERE id_gasto='" & Val(mshGastos.TextMatrix(Me.mshGastos.Row, 0)) & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    
    strCadena = "call p_insert_compra_emitido_internacional_gasto('" & rst("id_compra_gasto") & "')"
    CnBd.Execute (strCadena)
    MsgBox "OK"
End If



End Sub

Private Sub DtcAfectoIgv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcMoneda.SetFocus
End If
End Sub

Private Sub DtcComprobante_Change()
If Me.DtcComprobante.BoundText = "0002" Then
   Me.frmretencion.Visible = True
Else
   Me.frmretencion.Visible = False
End If
End Sub

Private Sub DtcComprobante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtSerie)
End If
End Sub

Private Sub DtcMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcPeriodo.SetFocus
End If
End Sub

Private Sub DtcPeriodo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtcodigoprod)
End If
End Sub

Private Sub DtcTipoCompra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtDescripcion)
End If
End Sub

Private Sub DtpFecha_Change()
Me.TxtTipoCambio.Text = cambio_venta(CVDate(Me.Dtpfecha.Value))
Me.DtpVencimiento.Value = Me.Dtpfecha.Value
End Sub

Private Sub DtpFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtDescripcion)
End If
End Sub

Private Sub Form_Load()
Dim nombreperiodo As String
CenterForm Me
Me.Top = 100
Me.Dtpfecha.Value = KEY_FECHA
Me.DtpVencimiento.Value = KEY_FECHA

Me.txtTc.Text = KEY_CAMBIO
strCadena = "SELECT id_doc as Codigo,doc_des as Descripcion FROM comprobantes ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComprobante)
Me.DtcComprobante.BoundText = "0001"

strCadena = "SELECT tipo_compra as Codigo,descripcion as Descripcion FROM tipo_compra  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoCompra)

strCadena = "SELECT igv as Codigo,igv as Descripcion FROM afecto_igv ORDER BY igv"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAfectoIgv)

strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda ORDER BY id_moneda"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)
Me.DtcMoneda.BoundText = KEY_MONEDA

strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPeriodo)
Me.DtcPeriodo.BoundText = get_periodo_actual(KEY_FECHA)

strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_personal='si'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcResponsable)
  
strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & FrmCompras.DtcTipoDoc.BoundText & "' AND serie='" & FrmCompras.TxtSerie.Text & "' AND numero='" & FrmCompras.TxtNumeroDoc.Text & "' AND id_proveedor='" & FrmCompras.txtRuc.Text & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txtIdCompra.Text = rst("id_compra")
End If
Call llenar_gastos(Me.mshGastos, Val(Me.txtIdCompra.Text))
strCadena = "SELECT * FROM registro_compras WHERE mes='" & formato_item(Month(KEY_FECHA), 2) & "' AND anio='" & Year(KEY_FECHA) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    nombreperiodo = "REGISTRO COMPRAS:" + Space(1) + nombre_mes(Month(KEY_FECHA))
    strCadena = "INSERT INTO registro_compras(ruc,mes,anio,descripcion,id_estado)VALUES('" & KEY_RUC & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & nombreperiodo & "','01')"
    CnBd.Execute (strCadena)
    
End If

Me.TxtTipoCambio.Text = Format(KEY_CAMBIO_VENTA, "#,##0.0000")

If FrmCompras.DtTipoCompra.BoundText = "01" Then
    Call llenarImportacion(Me.HfVinculados, Val(Me.txtIdCompra.Text))
    Call llenar_productos(HfVinculados, Val(Me.txtIdCompra.Text))
End If
End Sub
Private Sub llenar_gastos(ByVal Grilla As MSHFlexGrid, ByVal id_compra As Double)
Dim Total As Double
strCadena = "SELECT * FROM view_factura_vinculada_gasto WHERE id_compra='" & id_compra & "' AND ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    FrmCompras.lblgastos.Text = Format(0, "###0.00")
    Grilla.Rows = 0
    
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 3000
           Grilla.ColWidth(3) = 2500
           Grilla.ColWidth(4) = 900
           Grilla.ColWidth(5) = 800
           Grilla.ColWidth(6) = 1500
           Grilla.ColWidth(7) = 1500
           Grilla.ColWidth(8) = 1500
        Next
         cabecera = "IDGASTO" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "CLIENTE/PROVEEDOR" & vbTab & "MONEDA" & vbTab & "TC" & vbTab & "MONTO" & vbTab & "VALOR VENTA [S/.]" & vbTab & "TOTAL [S/.]"
         Grilla.AddItem cabecera
         For k = 0 To 8
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        
             in_dolares = 0
             in_valor_venta = 0
             in_parcial = 0
             in_total = 0
             in_parcial_venta = 0
             in_total = 0
        For i = 1 To rst.RecordCount
             
            If rst("id_moneda") = "00002" Then
                 in_dolares = in_dolares + rst("monto")
                in_parcial = rst("monto") * rst("tc")
            Else
               in_parcial = rst("monto")
               in_dolares = in_dolares + rst("monto") / rst("tc")
            End If
             
            If KEY_CON_IGV = "si" Then
                If rst("afecto_igv") = "si" Then
                    in_valor_venta = in_valor_venta + in_parcial / (1 + KEY_IGV)
                    in_parcial_venta = in_parcial / (1 + KEY_IGV)
                Else
                    in_valor_venta = in_valor_venta + in_parcial
                    in_parcial_venta = in_parcial
                End If
            Else
                    in_valor_venta = in_valor_venta + in_parcial
                    in_parcial_venta = in_parcial
            End If
             
             in_total = in_total + in_parcial
             
             Fila = rst("id_gasto") & vbTab & rst("fecha") & vbTab & rst("comprobante") & vbTab & rst("nombre_completo") & vbTab & rst("moneda") & vbTab & Format(rst("tc"), "#,##0.0000") & vbTab & Format(rst("monto"), "#,##0.0000") & vbTab & Format(in_parcial_venta, "###0.000") & vbTab & Format(in_parcial, "###0.000")
             Grilla.AddItem Fila
             
        rst.MoveNext
        Next i
         cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(in_dolares, "#,##0.0000") & vbTab & Format(in_valor_venta, "#,##0.0000") & vbTab & Format(in_total, "#,##0.0000")
         Grilla.AddItem cabecera
          For k = 0 To 8
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &HC0C0FF
                            Next k
                           
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub
Private Sub lbltotalderechos_Change()
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lblTotalgeneral.Caption = Format(totalgeneral, "###0.00")
End Sub
Private Sub llenar_productos(ByVal Grilla As MSHFlexGrid, ByVal id_compra As Double)
strCadena = "SELECT D.id_producto,P.nombre_prod,U.abreviatura,D.cantidad FROM movimiento_compra_detalle D,producto P,unidad U WHERE D.id_producto=P.id_producto AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND D.id_compra='" & id_compra & "' AND D.ruc='" & KEY_RUC & "'"
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
           Grilla.ColWidth(0) = 750
           Grilla.ColWidth(1) = 3500
           Grilla.ColWidth(2) = 850
           Grilla.ColWidth(3) = 850
           
        Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION PRODUCTO" & vbTab & "UNIDAD" & vbTab & "CANTIDAD"
         Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
            
             Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & rst("cantidad")
             Grilla.AddItem Fila
             Fila = ""
        rst.MoveNext
        Next i
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub


Private Sub mshgastos_SelChange()
If Val(Me.mshGastos.TextMatrix(Me.mshGastos.Row, 0)) > 0 Then
   ' Me.Toolbar1.Enabled = True
Else
  '  Me.Toolbar1.Enabled = False
End If
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_DELETE
        
    Case KEY_UPDATE
        
End Select
End Sub

Private Sub txtadvalorem_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.TxtIgv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")
End Sub

Private Sub txtalmacenaje_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text)
Me.lblTotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcTipoCompra.SetFocus
End If
End Sub

Private Sub txtBuscarresponsable_Change()
 strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_personal='si' and nombre_completo LIKE '%" & Trim(Me.txtBuscarresponsable.Text) & "%'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcResponsable)

End Sub

Private Sub txtcerfgases_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lblTotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtciasnaviera_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lblTotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub TxtcodigoProd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM producto where id_producto='" & Trim(Me.txtcodigoprod.Text) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.txtcodigoprod.Text = Trim(Me.txtcodigoprod.Text)
       Me.txtProducto.Text = rst("nombre_prod")
    Else
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
    End If
End If
End Sub

Private Sub txtcomision_Change()
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lblTotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtconduccion_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lblTotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtderechoespecifico_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.TxtIgv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub

Private Sub txtdescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdagregar.SetFocus
End If
End Sub
Private Sub get_comprobante()
strCadena = "SELECT * FROM movimiento_compra WHERE id_proveedor='" & Trim(Me.txtDni.Text) & "' and id_doc='" & Me.DtcComprobante.BoundText & "' and serie='" & Trim(Me.TxtSerie.Text) & "' and numero='" & Trim(Me.txtNumero.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    Me.txtMonto.Text = rstK("total")
    If rstK("igv") > 0 Then
       Me.DtcAfectoIgv.BoundText = "si"
    Else
       Me.DtcAfectoIgv.BoundText = "no"
    End If
    
    Me.DtcMoneda.BoundText = rstK("id_moneda")
    Me.DtcPeriodo.BoundText = rstK("id_periodo")
    
End If

End Sub
Private Sub txtDni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtDni.Text) & "'"
   Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.lblcliente.Caption = UCase(rst("nombre_completo"))
        
        Call get_comprobante
        Call Resalta(Me.txtMonto)
        
    Else
        If Len(Me.txtDni.Text) > 7 And Len(Me.txtDni.Text) < 12 Then
        Procedencia = nuevo
        FrmDetallePersona.Show
        FrmDetallePersona.txtRuc.Text = Trim(Me.txtDni)
        FrmDetallePersona.chkProveedor.Value = 1
        FrmDetallePersona.ChkCliente.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
        Else
            Procedencia = Selecionar
            FrmPersona.Show
            Exit Sub
        End If
End If

End If
End Sub

Private Sub TxtFlete_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lblTotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtFleteImportacion_Change()
Me.TxtCif.Text = Format(Val(Me.txtFob.Text) + Val(Me.txtFleteImportacion.Text) + Val(Me.txtSeguro.Text), "###0.00")
End Sub

Private Sub txtfob_Change()
Me.TxtCif.Text = Format(Val(Me.txtFob.Text) + Val(Me.txtFleteImportacion.Text) + Val(Me.txtSeguro.Text), "###0.00")
End Sub

Private Sub txtgastosdepacho_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lblTotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtgastosoperativos_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lblTotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtgremiosmaritimos_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")

totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lblTotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtIgv_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.TxtIgv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub

Private Sub txtigvcomision_Change()
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lblTotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtipm_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.TxtIgv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub

Private Sub txtisc_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.TxtIgv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub



Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcAfectoIgv.SetFocus
End If
End Sub

Private Sub TxtNotadebito_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text)
Me.lblTotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtNumero.Text = formato_item(Me.txtNumero.Text, 8)
    Call Resalta(Me.txtDni)
End If
End Sub

Private Sub txtpercepcionigv_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.TxtIgv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub

Private Sub txtSeguro_Change()
Me.TxtCif.Text = Format(Val(Me.txtFob.Text) + Val(Me.txtFleteImportacion.Text) + Val(Me.txtSeguro.Text), "###0.00")
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie.Text = formato_item(Me.TxtSerie.Text, 3)
    Call Resalta(Me.txtNumero)
End If
End Sub

Private Sub txtservdespacho_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.TxtIgv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub

Private Sub txtsobretasa_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.TxtIgv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub

Private Sub txtterminal_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.TxtFlete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lblTotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

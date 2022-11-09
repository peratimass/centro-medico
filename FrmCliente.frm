VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPersona 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Clientes-Proveedores-Transportistas"
   ClientHeight    =   9240
   ClientLeft      =   240
   ClientTop       =   75
   ClientWidth     =   20145
   Icon            =   "FrmCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmActividades 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DETALLE DE ACTIVIDADES"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1935
      Left            =   14040
      TabIndex        =   37
      Top             =   4560
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CheckBox chk_actividad_fecha 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "FECHAS"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   320
         Left            =   360
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
      Begin VitekeySoft.ChameleonBtn cmdEstadoCliente 
         Height          =   855
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "ESTADO DE CLIENTES"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
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
         MICON           =   "FrmCliente.frx":030A
         PICN            =   "FrmCliente.frx":0326
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdActividades 
         Height          =   855
         Left            =   2280
         TabIndex        =   39
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "ACTIVIDADES DE PERSONAL"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
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
         MICON           =   "FrmCliente.frx":3558
         PICN            =   "FrmCliente.frx":3574
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpActividad_ini 
         Height          =   345
         Left            =   1320
         TabIndex        =   40
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   141623297
         CurrentDate     =   44568
      End
      Begin MSComCtl2.DTPicker DtpActividad_fin 
         Height          =   345
         Left            =   2760
         TabIndex        =   41
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9.75
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   141623297
         CurrentDate     =   44568
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   3975
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   4320
         Picture         =   "FrmCliente.frx":66FC
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.Frame frmValidar 
      BackColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   10680
      TabIndex        =   23
      Top             =   2280
      Width           =   7935
      Begin VB.TextBox txtObservacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1005
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   2400
         Width           =   4695
      End
      Begin VB.CheckBox chk_verificado 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "VERIFICADO "
         BeginProperty Font 
            Name            =   "Bahnschrift SemiBold SemiConden"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   1440
         TabIndex        =   32
         Top             =   2400
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DtpInicio 
         Height          =   345
         Left            =   1440
         TabIndex        =   30
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   141623297
         CurrentDate     =   44568
      End
      Begin VB.CheckBox chk_fechas 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "FECHAS"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   320
         Left            =   240
         TabIndex        =   29
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chk_periodo 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "PERIODO"
         BeginProperty Font 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   320
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo DtcPeriodoContable 
         Height          =   330
         Left            =   1440
         TabIndex        =   25
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   4194304
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VitekeySoft.ChameleonBtn cmdGenerarContabilidad 
         Height          =   495
         Left            =   1455
         TabIndex        =   26
         Top             =   1680
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "VALIDAR FACTURACION ELECTRONICA"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
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
         MICON           =   "FrmCliente.frx":95A0
         PICN            =   "FrmCliente.frx":95BC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar Prg_Contabilidad 
         Height          =   225
         Left            =   1440
         TabIndex        =   27
         Top             =   1440
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComCtl2.DTPicker DtpFin 
         Height          =   345
         Left            =   3720
         TabIndex        =   31
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   141623297
         CurrentDate     =   44568
      End
      Begin VitekeySoft.ChameleonBtn cmdVerificar 
         Height          =   525
         Left            =   1440
         TabIndex        =   33
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   926
         BTYPE           =   5
         TX              =   "PROCESAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   8.25
            Charset         =   0
            Weight          =   350
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
         MICON           =   "FrmCliente.frx":BD43
         PICN            =   "FrmCliente.frx":BD5F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Hfacciones 
         Height          =   2415
         Left            =   240
         TabIndex        =   35
         Top             =   3480
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4260
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
            Name            =   "Bahnschrift SemiLight SemiConde"
            Size            =   9
            Charset         =   0
            Weight          =   350
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblempresa 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Bahnschrift SemiCondensed"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1320
         TabIndex        =   36
         Top             =   150
         Width           =   4785
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   7560
         Picture         =   "FrmCliente.frx":E344
         Top             =   240
         Width           =   240
      End
   End
   Begin VB.CheckBox chk_tiendas_online 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "EMPRESAS ONLINE"
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
      Height          =   315
      Left            =   12120
      TabIndex        =   19
      Top             =   540
      Width           =   1575
   End
   Begin VitekeySoft.ChameleonBtn cmdBuscar 
      Height          =   315
      Left            =   16920
      TabIndex        =   18
      Top             =   195
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
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
      MICON           =   "FrmCliente.frx":111E8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcPlanes 
      Height          =   315
      Left            =   13800
      TabIndex        =   17
      Top             =   200
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   4194304
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
   Begin VB.CheckBox chk_planes_empresa 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "PLANES EMPRESA"
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
      Height          =   315
      Left            =   12120
      TabIndex        =   15
      Top             =   200
      Width           =   1575
   End
   Begin VB.CheckBox chk_proveedor 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "PROVEEDOR"
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
      Height          =   250
      Left            =   8400
      TabIndex        =   14
      Top             =   510
      Width           =   1815
   End
   Begin VB.CheckBox chk_cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "CLIENTE"
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
      Height          =   250
      Left            =   8400
      TabIndex        =   13
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox pbImageFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   -2040
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TxtApellido 
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
      Left            =   1515
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox TxtRuc 
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
      Left            =   5355
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   7935
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   13996
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
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   9
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VitekeySoft.ChameleonBtn cmdexit 
      Height          =   855
      Left            =   18960
      TabIndex        =   6
      Top             =   8130
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCliente.frx":11204
      PICN            =   "FrmCliente.frx":11220
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
      Left            =   18960
      TabIndex        =   7
      Top             =   2850
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCliente.frx":11610
      PICN            =   "FrmCliente.frx":1162C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdupdate 
      Height          =   855
      Left            =   18960
      TabIndex        =   8
      Top             =   1965
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MODIFICAR"
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
      MICON           =   "FrmCliente.frx":13A76
      PICN            =   "FrmCliente.frx":13A92
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   855
      Left            =   18960
      TabIndex        =   9
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCliente.frx":13DAC
      PICN            =   "FrmCliente.frx":13DC8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdmatricula 
      Height          =   855
      Left            =   18960
      TabIndex        =   10
      Top             =   5565
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "PAGOS"
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
      MICON           =   "FrmCliente.frx":1421A
      PICN            =   "FrmCliente.frx":14236
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdficha 
      Height          =   855
      Left            =   18960
      TabIndex        =   11
      Top             =   4680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "ACTIVIDADES"
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
      MICON           =   "FrmCliente.frx":16B2E
      PICN            =   "FrmCliente.frx":16B4A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdContrato 
      Height          =   855
      Left            =   18960
      TabIndex        =   12
      Top             =   6480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "HISTORIA"
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
      MICON           =   "FrmCliente.frx":18DA3
      PICN            =   "FrmCliente.frx":18DBF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcRubro 
      Height          =   315
      Left            =   13800
      TabIndex        =   20
      Top             =   540
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   4194304
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
   Begin VitekeySoft.ChameleonBtn cmdBuscarOnline 
      Height          =   315
      Left            =   16920
      TabIndex        =   21
      Top             =   540
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
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
      MICON           =   "FrmCliente.frx":1B390
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdvalidar_sunat 
      Height          =   855
      Left            =   18960
      TabIndex        =   24
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "VALIDAR"
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
      MICON           =   "FrmCliente.frx":1B3AC
      PICN            =   "FrmCliente.frx":1B3C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblcount_online 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   17760
      TabIndex        =   22
      Top             =   540
      Width           =   1575
   End
   Begin VB.Label lblacount 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   17760
      TabIndex        =   16
      Top             =   200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAZON SOCIAL :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   345
      TabIndex        =   3
      Top             =   360
      Width           =   1065
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/DNI :"
      BeginProperty Font 
         Name            =   "Bahnschrift SemiLight SemiConde"
         Size            =   8.25
         Charset         =   0
         Weight          =   350
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   645
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   800
      Left            =   120
      Top             =   100
      Width           =   19815
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20055
   End
End
Attribute VB_Name = "FrmPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public EnumFrmCliente As EnumCliente
Public Procedencia As EnumProcede

Private Sub ChkMondoAdelantado_Click()

End Sub

Private Sub cmdvehiculos_Click()

End Sub



Private Sub Check1_Click()

End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub chk_fechas_Click()

If Me.chk_fechas.Value = 1 Then
    
  Me.DtpInicio.Visible = True
  Me.DtpFin.Visible = True
Else
    Me.DtpInicio.Visible = False
  Me.DtpFin.Visible = False
End If

End Sub

Private Sub chk_planes_empresa_Click()
If Me.chk_planes_empresa.Value = 1 Then
    
    strCadena = "SELECT id_plan as Codigo,descripcion as Descripcion FROM plan_servicio WHERE  ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcPlanes)
    Me.DtcPlanes.Visible = True
    Me.cmdBuscar.Visible = True
Else
    Me.DtcPlanes.Visible = False
    Me.cmdBuscar.Visible = False
    
End If
End Sub

Private Sub chk_tiendas_online_Click()

If Me.chk_tiendas_online.Value = 1 Then
    strCadena = "SELECT codigo as Codigo,descripcion as Descripcion FROM persona_rubro ORDER BY descripcion"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcRubro)
    Me.DtcRubro.Visible = True
    Me.cmdBuscarOnline.Visible = True
Else
    Me.DtcRubro.Visible = False
    Me.cmdBuscarOnline.Visible = False
End If


End Sub

Private Sub cmdActividades_Click()
Dim cam3(0 To 5, 1 To 5)  As String
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(3, 1) = "empresa"
    cam3(4, 1) = "direccion"
    cam3(5, 1) = "titulo"
    
    strCadena = "SELECT CURDATE()"
    Call ConfiguraRstA(strCadena)
    
   
    
    cam3(0, 2) = Format(KEY_FECHA, "dd-mm-YYYY")
    cam3(1, 2) = Format(KEY_FECHA, "dd-mm-YYYY")
    cam3(2, 2) = ""
    cam3(3, 2) = KEY_EMPRESA
    cam3(4, 2) = KEY_DIRECCION_ALM
    cam3(5, 2) = "SALDO X PERIODO"
    param = cam3()
    
strCadena = "call ADM_auditoria_empresa('7','" & Me.DtcPlanes.BoundText & "','','','','','" & Format(Me.DtpActividad_ini.Value, "YYYY-mm-dd") & "','','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptInformeActividad_det", param, App.Path + "\Reportes\")



End Sub

Private Sub cmdBuscar_Click()

Call Me.llenarGrid_plan(Me.HfdPersona, "")

End Sub

Private Sub cmdBuscarOnline_Click()





Call Me.llenarGrid_plan(Me.HfdPersona, Me.DtcRubro.BoundText)
End Sub

Private Sub cmdContrato_Click()
in_dni = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
Call generar_hc(Trim(in_dni))

strCadena = "UPDATE entidad_empresa SET nuevo='no' WHERE cod_unico='" & in_dni & "' AND id_empresa='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)

Exit Sub
Dim cam3(0 To 5, 1 To 5)  As String
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(3, 1) = "empresa"
    cam3(4, 1) = "direccion"
    cam3(5, 1) = "titulo"
    
    cam3(0, 2) = Format(Me.DtpInicio.Value, "dd-mm-YYYY")
    cam3(1, 2) = Format(Me.DtpFin.Value, "dd-mm-YYYY")
    cam3(2, 2) = in_almacen
    cam3(3, 2) = KEY_EMPRESA
    cam3(4, 2) = KEY_DIRECCION_ALM
    cam3(5, 2) = "AUDITORIA INTERNA"
    param = cam3()
    
strCadena = "call ADM_auditoria_empresa('5','" & Me.DtcPlanes.BoundText & "','','" & KEY_USUARIO & "','','','','','" & KEY_RUC & "')"

Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptInformeAuditoria", param, App.Path + "\Reportes\")
Exit Sub



End Sub

Private Sub cmddelete_Click()
If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        strCadena = "SELECT * FROM movimiento_venta WHERE id_cliente='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' and ruc='" & KEY_RUC & "' "
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
         strCadena = "call p_delete_persona('" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "','" & KEY_RUC & "')"
         CnBd.Execute (strCadena)
          
         Call actualizar
         Else
            MsgBox "Imposible Eliminar a este Usuario, esta Vinculado a Movimientos"
         End If
      End If
End Sub

Private Sub cmdEstadoCliente_Click()
Dim cam3(0 To 5, 1 To 5)  As String
    cam3(0, 1) = "fecha_ini"
    cam3(1, 1) = "fecha_fin"
    cam3(2, 1) = "almacen"
    cam3(3, 1) = "empresa"
    cam3(4, 1) = "direccion"
    cam3(5, 1) = "titulo"
    
    strCadena = "SELECT CURDATE()"
    Call ConfiguraRstA(strCadena)
    
   
    
    cam3(0, 2) = Format(KEY_FECHA, "dd-mm-YYYY")
    cam3(1, 2) = Format(KEY_FECHA, "dd-mm-YYYY")
    cam3(2, 2) = ""
    cam3(3, 2) = KEY_EMPRESA
    cam3(4, 2) = KEY_DIRECCION_ALM
    cam3(5, 2) = "SALDO X PERIODO"
    param = cam3()
    
strCadena = "call ADM_auditoria_empresa('6','" & Me.DtcPlanes.BoundText & "','','','','','','','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptInformeActividad", param, App.Path + "\Reportes\")




End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdficha_Click()
Me.DtpActividad_ini.Value = KEY_FECHA
Me.DtpActividad_fin.Value = KEY_FECHA
Me.frmActividades.Visible = True


End Sub

Private Sub cmdGenerarContabilidad_Click()
Dim strHtml As String
    Set DomDoc = New XMLHTTP
    
    Dim bandera As Integer
    
    bandera = 0
    If Me.chk_periodo.Value = 1 Then
        strCadena = "SELECT * FROM con_periodo WHERE id='" & Me.DtcPeriodoContable.BoundText & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            in_fecha_inicio = rst("FechaInicio")
            in_fecha_fin = rst("FechaFin")
        End If
        bandera = 1
    End If
    
    If Me.chk_fechas.Value = 1 Then
        bandera = 1
        in_fecha_inicio = Me.DtpInicio.Value
        in_fecha_fin = Me.DtpFin.Value
    End If
    
    If bandera = 0 Then
        MsgBox "SELECCIONE UNA OPCION :" + Chr(13) + "1.- PERIODO" + Chr(13) + "2.- RANGO DE FECHAS", vbInformation
        Exit Sub
    End If
    
    urlstr = "https://api.vitekey.com/keyfact/utils/reporte-ventas?password=vitekey2018&ruc=" & Trim(FrmPersona.HfdPersona.TextMatrix(FrmPersona.HfdPersona.Row, 0)) & "&date_start=" & Format(in_fecha_inicio, "MM/dd/YYYY") & "&date_end=" & Format(in_fecha_fin, "MM/dd/YYYY")
    
    Set DomDoc = New XMLHTTP
     'Parámetros en formato URLEncode
     'Metodo a usar, url, y true en caso de manejar la respuesta en modo asíncrono
     DomDoc.Open "GET", urlstr, False
     'encabezados
     DomDoc.setRequestHeader "Connection", "close"
     DomDoc.send params
     'La respuesta, en caso de existir, está en responseBody.
    'También puedes especificar responseXml si tu aplicación devolviese XML
     strHtml = StrConv(DomDoc.responseBody, vbUnicode)
     Call verificar_ventas_keyfacil(strHtml, Trim(FrmPersona.HfdPersona.TextMatrix(FrmPersona.HfdPersona.Row, 0)))
     
     Call put_registrar_acciones(FrmPersona.HfdPersona.TextMatrix(FrmPersona.HfdPersona.Row, 0), "", "01")
    
     
     
     
End Sub
Private Sub put_registrar_acciones(ByVal in_ruc As String, ByVal in_descripcion As String, ByVal in_tipo As String)

If in_tipo = "01" Then
        If Me.chk_periodo.Value = 1 Then
           descripcion = Mid(KEY_VENDEDOR, 1, 18) + Space(2) + ": " + "F.ELEC:  [" + Me.DtcPeriodoContable.Text + "]"
        Else
           descripcion = Mid(KEY_VENDEDOR, 1, 18) + Space(2) + ": " + "F.ELEC: [" + Format(Me.DtpInicio.Value, "dd-mm-YYYY") + " - " + Format(Me.DtpFin.Value, "dd-mm-YYYY") + "]"
        End If
         Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 9) = "PENDIENTE"
         For k = 9 To 11
            Me.HfdPersona.col = k
            Me.HfdPersona.CellBackColor = &HDFDFE0
        Next k
        in_estado = "01"
Else
    If Me.chk_verificado.Value = 1 Then
        strCadena = "SELECT DATE_SUB(NOW(), INTERVAL 5 HOUR);"
        Call ConfiguraRst(strCadena)
        Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 10) = Format(rst(0), "dd-mm-YYYY")
        Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 11) = Format(rst(0), "HH:mm:ss")
        Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 9) = "VERIFICADO"
        
        For k = 9 To 11
            Me.HfdPersona.col = k
            Me.HfdPersona.CellBackColor = &H80FF80
        Next k
        descripcion = Mid(KEY_VENDEDOR, 1, 18) + Space(2) + ": " + "VERIFICACION CORRECTA" + Space(1) + Trim(in_descripcion)
        in_estado = "02"
        
    Else
        MsgBox "SELECCIONE EL CHECK DE VERIFICADO.", vbInformation, Mid(KEY_VENDEDOR, 1, 20)
        Exit Sub
    End If
    
End If
 
strCadena = "call ADM_auditoria_empresa('2','','" & in_ruc & "','" & KEY_USUARIO & "','" & in_estado & "','" & descripcion & "','','','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

Call Me.llenar_acciones(Me.Hfacciones, in_ruc)

End Sub
Private Sub cmdmatricula_Click()
Dim in_dni As String
'Call FrmVentas.Show
Call FrmVentas.activar
strCadena = "P_nueva_venta('" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
in_dni = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
FrmVentas.TxtCodCliente.Text = in_dni
Call FrmVentas.precionar_cliente
strCadena = "SELECT * FROM college_servicio_persona WHERE dni='" & in_dni & "' and saldo>0 and ruc='" & KEY_RUC & "' ORDER BY id_detalle ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
       strCadena = "INSERT INTO temporal_ventas(ruc,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save) VALUES " & _
        "('" & KEY_RUC & "','" & in_dni & "','" & KEY_ALM & "','" & FrmVentas.DtcTipoDoc.BoundText & "','" & Trim(FrmVentas.DtcSerieDoc.BoundText) & "','" & Trim(FrmVentas.TxtNumeroDoc.Text) & "','" & rst("id_servicio") & "','1'," & _
        "'" & Val(rst("saldo")) & " ','" & Val(rst("saldo")) & "','0','si','" & rst("detalle") & "','" & KEY_USUARIO & "')"
        CnBd.Execute (strCadena)
        rst.MoveNext
   Next i
End If
'strCadena = "SELECT * FROM almacen_producto WHERE id_producto='00001' and  ruc='" & KEY_RUC & "' limit 1"
'Call ConfiguraRstK(strCadena)
'If rstK.RecordCount > 0 Then
'       in_detalle = "MENSUALIDAD [ENERO -2017 ]"
'       strCadena = "INSERT INTO temporal_ventas(ruc,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save) VALUES " & _
'        "('" & KEY_RUC & "','" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "','" & KEY_ALM & "','" & FrmVentas.DtcTipoDoc.BoundText & "','" & Trim(frmventas.DtcSerieDoc.BoundText ) & "','" & Trim(FrmVentas.TxtNumeroDoc.Text) & "','00001','1'," & _
'        "'" & Val(rstK("precio_venta")) & " ','" & Val(rstK("precio_venta")) & "','0','si','" & in_detalle & "','" & KEY_USUARIO & "')"
'        CnBd.Execute (strCadena)
'End If

 Call FrmVentas.llenarGrid_det(FrmVentas.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText, Trim(FrmVentas.txtformato_impresion.Text))
       
    
End Sub

Private Sub cmdnuevo_Click()
 Procedencia = nuevo
      FrmDetallePersona.Show
      'Call Resalta(FrmDetallePersona.TxtRuc)
      Exit Sub
End Sub

Private Sub cmdupdate_Click()
 Procedencia = modificar
      FrmDetallePersona.Show
End Sub

Private Sub cmdvalidar_sunat_Click()
 strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by Ejercicio DESC,mes DESC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodoContable)
  
  Me.chk_fechas.Value = 0
  Me.DtpInicio.Value = KEY_FECHA
  Me.DtpFin.Value = KEY_FECHA
  Me.DtpInicio.Visible = False
  Me.DtpFin.Visible = False
    
  Me.HfdPersona.Enabled = False
  Me.frmValidar.Visible = True
  Me.LblEmpresa.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
  Call Me.llenar_acciones(Me.Hfacciones, Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
  

End Sub

Private Sub cmdVerificar_Click()
Call put_registrar_acciones(FrmPersona.HfdPersona.TextMatrix(FrmPersona.HfdPersona.Row, 0), Trim(Me.txtObservacion.Text), "02")

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    FrmDetallesParametros.Procedencia = Neutro
    Unload Me
    
End If
If KeyCode = 40 Then
    Me.HfdPersona.SetFocus
End If
End Sub
Public Sub actualizarPersonal()
  strCadena = "SELECT cPersona,NombrePersona, sDireccionCliente1 ,Per_Ruc,puntos FROM Persona WHERE personal='V' ORDER BY NombrePersona ASC"
  Call llenarGrid(Me.HfdPersona)

  Exit Sub
End Sub
Public Sub actualizar()

If KEY_RUBRO = "00026" Then ' INVESTIGACION
   strCadena = "SELECT * FROM view_investigacion WHERE ruc='" & KEY_RUC & "' limit 30"
   Call Me.llenarGrid_investigacion(Me.HfdPersona)
   Exit Sub
End If





If KEY_RUBRO = "00025" Then
   strCadena = "SELECT * FROM view_estudiante WHERE id_empresa='" & KEY_RUC & "' limit 30"
Else
   strCadena = "SELECT * FROM view_cliente WHERE id_cliente='si' and  id_empresa='" & KEY_RUC & "' limit 30"
End If

Call llenarGrid(Me.HfdPersona)
End Sub




Public Sub actualizar_contadores()
strCadena = "SELECT P.dni,P.nombre_completo,P.direccion,P.id_departamento FROM entidad_empresa E,persona P WHERE  E.cod_unico=P.dni AND E.id_tipo_per='00022' AND id_empresa='0' ORDER BY nombre_completo"
Call llenarGridContador(Me.HfdPersona, Me)
End Sub
Private Sub Form_Load()
 CenterForm Me
 Me.Top = 50
 Me.frmValidar.Visible = False
 If KEY_RUBRO = "00025" Then
    Me.cmdmatricula.Enabled = True
    Me.cmdficha.Enabled = True
    Me.cmdContrato.Caption = "CONTRATO"
Else
    Me.cmdmatricula.Enabled = False
    Me.cmdficha.Enabled = False
    Me.cmdContrato.Caption = "HISTORIA"
 End If
 
 If FrmDetallesParametros.Procedencia = buscar Then
    strCadena = "SELECT * FROM entidad_empresa WHERE id_tipo_per='00022' AND id_empresa='0'"
    Call ConfiguraRst(strCadena)
    
    Call actualizar_contadores
    Exit Sub
 End If
 
 'strCadena = "SELECT count(*) FROM entidad_empresa WHERE id_cliente='si' and  id_empresa='" & KEY_RUC & "'"
 'Call ConfiguraRst(strCadena)
 'Me.lblAcoount.Caption = str(rst(0)) + Space(2) + "REGISTRADOS"
 If FrmDetalleAlmacen.Procedencia = Selecionar Then
    Call actualizarPersonal
 Else
 Call actualizar
 End If
End Sub
Private Sub OptApellido_Click()
End Sub



Private Sub HfdPersona_Click()
'If Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Or Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) = "00000000" Then
    
  

  
'     Me.cmdNuevo.Enabled = True
     
'     Me.cmdDelete.Enabled = True
'     Exit Sub
'Else
  
'     Me.cmdupdate.Enabled = False
'     Me.cmdDelete.Enabled = False

'End If
End Sub

Private Sub HfdPersona_DblClick()

'Call disabled_form(Me)
frmpersonaDeuda.Show
Call frmpersonaDeuda.llenarGrid_deuda(frmpersonaDeuda.HfdPersona, Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
Exit Sub

End Sub

Private Sub HfdPersona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If FrmVentas.Procedencia = Selecionar Then
          FrmVentas.TxtCodCliente.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
          Call FrmVentas.precionar_cliente
          FrmVentas.Procedencia = Neutro
          Unload Me
          Exit Sub
    End If
     
   If FrmVentas.Procedencia = seleccionar_otro Then
      FrmVentas.txtdni_copropietario.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
      FrmVentas.lblcopropietario.Caption = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
      FrmVentas.Procedencia = Neutro
      Unload Me
      Exit Sub
   End If
    
    If frmCajaEgreso.Procedencia = Selecionar Then
       frmCajaEgreso.txtrucproveedor.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
       frmCajaEgreso.lblProveedor.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
       frmCajaEgreso.Procedencia = Neutro
       Unload Me
       Exit Sub
    End If
    
    
    If frmMemorandun.Procedencia = Selecionar Then
       frmMemorandun.Procedencia = Neutro
       frmMemorandun.txtDni.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
       frmMemorandun.TxtRazon.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
       Call frmMemorandun.put_estado_cuenta_temp(Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)))
       Unload Me
       Exit Sub
    End If
    
    
    If frmHotel.Procedencia = Selecionar Then
       frmHotel.Procedencia = Neutro
       frmHotel.txtDni.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
       frmHotel.TxtCliente.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
       Unload Me
       Exit Sub
    End If
    
    
    If FrmReporteRegistroVentas.Procedencia = Selecionar Then
       FrmReporteRegistroVentas.Procedencia = Neutro
       FrmReporteRegistroVentas.TxtCliente.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
       Unload Me
       Exit Sub
    End If
    
    If FrmListadoFacturasCompra.Procedencia = Selecionar Then
       FrmListadoFacturasCompra.Procedencia = Neutro
       FrmListadoFacturasCompra.TxtCliente.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
       Unload Me
       Exit Sub
    End If
    
    If frmAnalisisporCuenta.Procedencia = Selecionar Then
       frmAnalisisporCuenta.Procedencia = Neutro
       frmAnalisisporCuenta.txtRuc.Text = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0))
       frmAnalisisporCuenta.lblcliente.Caption = Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
       Unload Me
       Exit Sub
    End If
    
    
    
    If FrmOrdenCompraDet.Procedencia = Selecionar Then
       FrmOrdenCompraDet.Procedencia = Neutro
       strCadena = "SELECT * FROM view_entidad WHERE dni='" & Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0) & "' and ruc='" & KEY_RUC & "'"
       Call ConfiguraRstL(strCadena)
       If rstL.RecordCount > 0 Then
          FrmOrdenCompraDet.txtRuc.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
          FrmOrdenCompraDet.TxtProveedor.Text = rstL("nombre_completo")
          If rstL("afecto_igv") = "si" Then
             FrmOrdenCompraDet.chk_igv.Value = 1
          Else
             FrmOrdenCompraDet.chk_igv.Value = 0
          End If
          
       End If
       
       
       
       Unload Me
       Exit Sub
    End If
    
    If FrmDetallePedido.Procedencia = Selecionar Then
       FrmDetallePedido.Procedencia = Neutro
       FrmDetallePedido.txtRuc.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
       FrmDetallePedido.TxtProveedor.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
       Unload Me
       Exit Sub
    End If
    
    
    
    If FrmCambioAceite.Procedencia = Selecionar Then
       FrmCambioAceite.txtDni.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
       FrmCambioAceite.lblpropietario.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
       FrmCambioAceite.Procedencia = Neutro
       Unload Me
       Exit Sub
       
    End If
    
    If FrmTransferencias.Procedencia = seleccionar_atencion Then
       FrmTransferencias.txt_dni_atencion.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
       FrmTransferencias.txt_atencion.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
       FrmTransferencias.Procedencia = Neutro
       Unload Me
       Exit Sub
    End If
    If FrmTransferencias.Procedencia = seleccionar_otro Then
       FrmTransferencias.Procedencia = Neutro
       FrmTransferencias.txt_idremitente.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
       FrmTransferencias.txtremitente.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
       Unload Me
       Exit Sub
    End If
    
    
    If FrmServiciotecnico.Procedencia = Selecionar Then
       FrmServiciotecnico.Procedencia = Neutro
       FrmServiciotecnico.TXTDNIRUC.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
       FrmServiciotecnico.lblpropietario.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
       Unload Me
       Exit Sub
    End If
    
        
    If frmmanifiesto.Procedencia = Selecionar Then
       frmmanifiesto.Procedencia = Neutro
       frmmanifiesto.txtRuc.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
       frmmanifiesto.txtpropietario.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
       frmmanifiesto.txtDireccion.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 2)
       Unload Me
       Exit Sub
    End If
    
    If frmmanifiesto.Procedencia = seleccionar_per Then
       frmmanifiesto.Procedencia = Neutro
       Call frmmanifiesto.BuscarChofer(Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)))
       Unload Me
       Exit Sub
    End If
    
    If frmCorTramite.Procedencia = Selecionar Then
       frmCorTramite.txtDni.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
       frmCorTramite.lblNombre.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
       frmCorTramite.Procedencia = Neutro
       Call frmCorTramite.actualizaLista(frmCorTramite.gridTramites, "01")
       
       Unload Me
       Exit Sub
    End If
    
    
    
    If FrmVentasPersonalizada.Procedencia = Selecionar Then
        strCadena = "SELECT * FROM persona WHERE dni='" & Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0) & "'"
        Call ConfiguraRst(strCadena)
        FrmVentasPersonalizada.txtRuc.Text = rst("dni")
        FrmVentasPersonalizada.TxtRazonSocial.Text = rst("nombre_completo")
        FrmVentasPersonalizada.txtDireccion.Text = rst("direccion")
        
        If IsNull(rst("foto")) = False And Len(rst("foto")) > 5 Then
                If VerificarFichero(App.Path & "\archivos\" & rst("dni")) = True Then
                    FrmVentasPersonalizada.imgFoto.Picture = LoadPicture(App.Path + "\archivos\" + rst("dni") + "\" + Trim(rst("foto")))
                Else
                    FrmVentasPersonalizada.imgFoto.Picture = LoadPicture(App.Path + "\archivos\no_photo.jpg")
                End If
        Else
            FrmVentasPersonalizada.imgFoto.Picture = LoadPicture(App.Path + "\archivos\no_photo.jpg")
        End If
        FrmVentasPersonalizada.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    
    
    
    If FrmDetallesParametros.Procedencia = buscar Then
        FrmDetallesParametros.TxtRucContador.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmDetallesParametros.lblRazonContador.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        
        FrmDetallesParametros.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    
    If frmVentasPagos.Procedencia = Selecionar Then
        frmVentasPagos.txtRuc.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        frmVentasPagos.TxtCliente.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        frmVentasPagos.txtDireccion.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 2)
        Call Resalta(frmVentasPagos.TxtMontoPago)
        frmVentasPagos.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    
    If FrmTransferencias.Procedencia = Selecionar Then
        FrmTransferencias.TxtRucDestino.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmTransferencias.TxtNombreDestino.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        FrmTransferencias.txtdireccionfiscal.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 2)
        FrmTransferencias.txtDireccionLlegada.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 2)
        Call get_ubigeo_sunat(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0), FrmTransferencias.txtIdUbigeoDestino, FrmTransferencias.txtUbigeoDestino)
        Call Resalta(FrmTransferencias.TxtRucTransporte)
        FrmTransferencias.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    
    
    
    
    
    If FrmTransferencias.Procedencia = buscar Then
        FrmTransferencias.TxtRucTransporte.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmTransferencias.lblRazonTransporte.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        Call Resalta(FrmTransferencias.TxtMarcayPlaca)
        FrmTransferencias.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    If FrmTransferencias.Procedencia = modificar Then
        FrmTransferencias.TxtRucChofer.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmTransferencias.lblRazonChofer.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        strCadena = "SELECT * FROM persona WHERE dni='" & Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0) & "'"
        Call ConfiguraRst(strCadena)
        If IsNull(rst("licencia")) = True Then
            FrmTransferencias.TxtLicencia.Text = ""
        Else
            FrmTransferencias.TxtLicencia.Text = rst("licencia")
        End If
        Call Resalta(FrmTransferencias.TxtLicencia)
        FrmTransferencias.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
    
     If FrmTransferencias.Procedencia = buscar Then
        FrmTransferencias.TxtRucTransporte.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmTransferencias.lblRazonTransporte.Caption = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        Call Resalta(FrmTransferencias.TxtMarcayPlaca)
        FrmTransferencias.Procedencia = Neutro
        Unload Me
        Exit Sub
    End If
     If FrmChequeNuevo.Procedencia = buscar Then
        FrmChequeNuevo.txtRuc.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmChequeNuevo.TxtRazonSocial.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        FrmChequeNuevo.txtDireccion.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 2)
        FrmChequeNuevo.Procedencia = Neutro
        Call Resalta(FrmChequeNuevo.Txtcentrocosto)
        Unload Me
        Exit Sub
    End If
    
      If FrmSolicitudViaticosDeclarar.Procedencia = Selecionar Then
        FrmSolicitudViaticosDeclarar.txtRuc.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        FrmSolicitudViaticosDeclarar.TxtRazonSocial.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
        FrmSolicitudViaticosDeclarar.Procedencia = Neutro
       ' FrmSolicitudViaticosDeclarar.cmdagregar.SetFocus
        Unload Me
        Exit Sub
    End If
     
  
    
    
    If FrmDetalleAlmacen.Procedencia = Selecionar Then
       strCadena = "SELECT cPersona,NombrePersona,sDireccionCliente1,Per_Ruc FROM " & _
       " Persona WHERE cPersona='" & Trim(Me.HfdPersona.Text) & "'"
       Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
          FrmDetalleAlmacen.TxtCodCliente.Text = rst("cPersona")
          FrmDetalleAlmacen.txtEncargado.Text = rst("NombrePersona")
        End If
        FrmDetalleAlmacen.Procedencia = Neutro
       Unload Me
       Set rst = Nothing
       Exit Sub
    End If
    
     
    
'    If FrmOrdenCompraDet.Procedencia = buscar Then
 '      strCadena = "SELECT cPersona,NombrePersona  FROM " & _
  '     " Persona WHERE cPersona='" & Trim(Me.HfdPersona.text) & "'"
   '    Call ConfiguraRst(strCadena)
    '    If rst.RecordCount > 0 Then
     '     FrmOrdenCompraDet.TxtcodProveedor.text = rst("cPersona")
      '    FrmOrdenCompraDet.TxtProveedor.text = rst("NombrePersona")
       '   FrmOrdenCompraDet.DtcTipoTransporte.SetFocus
        'End If
 '       FrmOrdenCompraDet.Procedencia = Neutro
  '     Unload Me
   '    Set rst = Nothing
    '   Exit Sub
    'End If
     
     
   
    
    
    
     
     
     'If FrmOrdenCompraDet.Procedencia = Selecionar Then
     '  strCadena = "SELECT cPersona,NombrePersona,licencia FROM " & _
     '  " Persona WHERE cPersona='" & Trim(Me.HfdPersona.text) & "'"
     '  Call ConfiguraRst(strCadena)
     '   If rst.RecordCount > 0 Then
     '     FrmOrdenCompraDet.TxtCPersona.text = rst("cPersona")
     '     FrmOrdenCompraDet.TxtChofer.text = rst("NombrePersona")
     '     FrmOrdenCompraDet.TxtLicencia.text = rst("licencia")
     '     Call Resalta(FrmOrdenCompraDet.TxtLicencia)
     '   End If
     '   FrmOrdenCompraDet.Procedencia = Neutro
     '  Unload Me
     '  Set rst = Nothing
     '  Exit Sub
    'End If
    
    If FrmCompras.Procedencia = Selecionar Then
          FrmCompras.txtRuc.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
          Call FrmCompras.buscar_comprobante
          FrmCompras.TxtProveedor.Text = UCase(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
          FrmCompras.txtDireccion.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 2)
          Unload Me
          FrmCompras.DtTipoCompra.SetFocus
          FrmCompras.Procedencia = Neutro
          Exit Sub
    End If
     If FrmComprasGastos.Procedencia = Selecionar Then
          FrmComprasGastos.txtDni.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
          FrmComprasGastos.lblcliente.Caption = UCase(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
          Call Resalta(FrmComprasGastos.txtMonto)
          FrmComprasGastos.Procedencia = Neutro
          Unload Me
          Exit Sub
    End If
    
    
     If FrmOrdenCompraDet.Procedencia = buscar Then
          FrmOrdenCompraDet.txtDni.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
          FrmOrdenCompraDet.lblcliente.Caption = UCase(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))
          Call Resalta(FrmOrdenCompraDet.txtMonto)
          FrmOrdenCompraDet.Procedencia = Neutro
          Unload Me
          Exit Sub
    End If
    
    
  If frmNuevoComprobante.Procedencia = buscar Then
       strCadena = "SELECT Per_Ruc FROM  Persona WHERE Per_Ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "'"
       Call ConfiguraRst(strCadena)
       If rst.RecordCount > 0 Then
                
            frmNuevoComprobante.txtRuc.Text = rst(0)
            
       End If
       Set rst = Nothing
       Unload Me
       frmNuevoComprobante.txtRuc.SetFocus
       frmNuevoComprobante.Procedencia = Neutro
       Exit Sub
    End If
    
If FrmListadoFacturasCompra.Procedencia = buscar Then
       strCadena = "SELECT cPersona FROM  Persona WHERE cPersona='" & Trim(Me.HfdPersona.Text) & "'"
       Call ConfiguraRst(strCadena)
       FrmListadoFacturasCompra.TxtcodProveedor.Text = Trim(rst(0))
       Call FrmListadoFacturasCompra.llenarGrid_Proveedor
       FrmListadoFacturasCompra.Procedencia = Neutro
       Unload Me
       Set rst = Nothing
        Exit Sub
End If
If FrmBusquedaDocumentos.Procedencia = buscar Then
       FrmBusquedaDocumentos.txtDni.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
       FrmBusquedaDocumentos.TxtCliente.Text = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)
       FrmBusquedaDocumentos.Procedencia = Neutro
       FrmBusquedaDocumentos.cmdBuscarCliente.Enabled = True
       FrmBusquedaDocumentos.cmdBuscarCliente.SetFocus
       Unload Me
       Exit Sub
End If


    
    If FrmDetalleGuia.Procedencia = Selecionar Then
       strCadena = "SELECT cPersona,NombrePersona,sDireccionCliente1,Per_Ruc FROM " & _
       " Persona WHERE cPersona='" & Trim(Me.HfdPersona.Text) & "'"
       Call ConfiguraRst(strCadena)
       
          FrmDetalleGuia.TxtCodigoEmpresaTransporte.Text = rst(0)
          FrmDetalleGuia.TxtNombreEmpresaTransporte.Text = rst(1)
          FrmDetalleGuia.TxtDireccionTransporte.Text = rst(2)
          FrmDetalleGuia.TxtRuc_Transportes.Text = rst(3)
          Unload Me
       
       Set rst = Nothing
       FrmDetalleGuia.Procedencia = Neutro
       Exit Sub
    End If
    If FrmAdelantoPersonal.Procedencia = Selecionar Then
       strCadena = "SELECT dni,nombre_completo,direccion FROM " & _
       " persona WHERE dni='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "'"
       Call ConfiguraRst(strCadena)
       
      
          'FrmAdelantoPersonal.TxtCodCliente.text = rst("cPersona")
          FrmAdelantoPersonal.TxtCliente.Text = rst("nombre_completo")
          FrmAdelantoPersonal.txtDireccion.Text = rst("direccion")
          FrmAdelantoPersonal.txtRuc.Text = rst("dni")
          FrmAdelantoPersonal.Procedencia = Neutro
          'Call FrmAdelantoPersonal.Resalta(FrmAdelantoPersonal.TxtObservacion)
      Set rst = Nothing
       Unload Me
       
       
      
       Exit Sub
    End If
    If FrmreciboIngresos.Procedencia = Selecionar Then
       strCadena = "SELECT dni,nombre_completo,direccion FROM " & _
       " persona WHERE dni='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "'"
       Call ConfiguraRst(strCadena)
       
      
          'FrmreciboIngresos.TxtCodCliente.text = rst(0)
          FrmreciboIngresos.TxtCliente.Text = rst(1)
          FrmreciboIngresos.txtDireccion.Text = rst("direccion")
          FrmreciboIngresos.txtRuc.Text = rst("dni")
          FrmreciboIngresos.txtObservacion.Text = ""
          FrmreciboIngresos.Procedencia = Neutro
          FrmreciboIngresos.TxtMontoPago.SetFocus
          
      Set rst = Nothing
       Unload Me
       FrmreciboIngresos.Procedencia = Neutro
       
      
       Exit Sub
    End If
      If FrmIngresoDinero.Procedencia = Selecionar Then
       strCadena = "SELECT cPersona,NombrePersona,sDireccionCliente1,Per_Ruc,Observacion FROM " & _
       " Persona WHERE cPersona='" & Trim(Me.HfdPersona.Text) & "'"
       Call ConfiguraRst(strCadena)
      
          FrmIngresoDinero.TxtCodCliente.Text = rst(0)
          FrmIngresoDinero.TxtCliente.Text = rst(1)
          FrmIngresoDinero.txtDireccion.Text = rst(2)
          FrmIngresoDinero.txtRuc.Text = rst(3)
          FrmIngresoDinero.txtObservacion.Text = rst(4)
          
      
       Unload Me
       FrmIngresoDinero.txtObservacion.SetFocus
       Set rst = Nothing
       FrmIngresoDinero.Procedencia = Neutro
       Exit Sub
    End If
    If FrmDetalleAdelanto.Procedencia = Selecionar Then
       strCadena = "SELECT cPersona,NombrePersona,sRazonSocial,sDireccionCliente1,Per_Ruc,Observacion,MontoAdelantado FROM " & _
       " Persona WHERE cPersona='" & Trim(Me.HfdPersona.Text) & "'"
       Call ConfiguraRst(strCadena)
       If Trim(rst(6)) = "N" Then
          FrmDetalleAdelanto.TxtCodCliente.Text = rst(0)
          FrmDetalleAdelanto.TxtCliente.Text = rst(1)
          FrmDetalleAdelanto.txtDireccion.Text = rst(3)
          FrmDetalleAdelanto.txtRuc.Text = rst(4)
          FrmDetalleAdelanto.txtsaldo.Text = rst(7)
        Else
          FrmDetalleAdelanto.TxtCodCliente.Text = rst(0)
          FrmDetalleAdelanto.TxtCliente.Text = rst(2)
          FrmDetalleAdelanto.txtDireccion.Text = rst(3)
          FrmDetalleAdelanto.txtRuc.Text = rst(4)
          FrmDetalleAdelanto.txtsaldo.Text = rst(7)
          
       End If
       Unload Me
       
       Set rst = Nothing
       FrmDetalleAdelanto.Procedencia = Neutro
       Exit Sub
    End If
End If
End Sub

Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
  
   Grilla.Rows = 0
  
         
         If KEY_RUBRO = "00025" Then ' COLEGIO
           ReDim arrColWidth(1 To rst.Fields.Count)
           For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 1200
            Grilla.ColWidth(1) = 5000
            Grilla.ColWidth(2) = 4500
            Grilla.ColWidth(3) = 1800
            Grilla.ColWidth(4) = 1800
            Grilla.ColWidth(5) = 2000
            Grilla.ColWidth(6) = 2000
          Next
            cabecera = "DNI" & vbTab & "NOMBRE ESTUDIANTE" & vbTab & "DIRECCION CLIENTE" & vbTab & "TELEFONO" & vbTab & "NIVEL" & vbTab & "GRADO ACADEMICO" & vbTab & "ESTADO"
            Grilla.AddItem cabecera
            For k = 0 To 6
                Grilla.col = k
                Grilla.Row = 0
                Grilla.CellBackColor = &HDFDFE0
            Next k
            rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            
            If rst("habilitado") = "si" Then
                in_habilitado = "ACTIVO"
            Else
                in_habilitado = "INACTIVO"
            End If
             Fila = rst("dni") & vbTab & UCase(rst("nombre_completo")) & vbTab & UCase(rst("direccion")) & vbTab & rst("celular") & vbTab & rst("nivel") & vbTab & rst("grado") & vbTab & in_habilitado
             Grilla.AddItem Fila
            
        rst.MoveNext
        Next i
  Else
              ReDim arrColWidth(1 To rst.Fields.Count)
            For Each Campo In rst.Fields
                Grilla.ColWidth(0) = 1500
                Grilla.ColWidth(1) = 5500
                Grilla.ColWidth(2) = 5500
                Grilla.ColWidth(3) = 2000
                Grilla.ColWidth(4) = 2500
                Grilla.ColWidth(5) = 1200
            Next
            cabecera = "DNI/RUC" & vbTab & "NOMBRE CLIENTE" & vbTab & "DIRECCION CLIENTE" & vbTab & "CELULAR" & vbTab & "E-MAIL" & vbTab & "ESTADO"
        
         Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If rst("habilitado") = "si" Then
                in_habilitado = "ACTIVO"
            Else
                in_habilitado = "INACTIVO"
            End If
            
             Fila = rst("dni") & vbTab & Trim(rst("nombre_completo")) & vbTab & UCase(rst("direccion")) & vbTab & rst("celular") & vbTab & rst("mail") & vbTab & in_habilitado
             Grilla.AddItem Fila
             If rst("habilitado") = "no" Then
                For k = 1 To 5
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
               Next k
             End If
        rst.MoveNext
        Next i
  End If
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
   
     Me.cmdupdate.Enabled = False
     Me.cmdDelete.Enabled = False
     
  Grilla.ColAlignment(0) = 1
  Grilla.ColAlignment(1) = 1
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
         
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Public Sub llenar_acciones(ByVal Grilla As MSHFlexGrid, ByVal in_ruc As String)
On Error GoTo salir
strCadena = "call ADM_auditoria_empresa('3','','" & in_ruc & "','','','','','','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
  
   Grilla.Rows = 0
  
        
              ReDim arrColWidth(1 To rst.Fields.Count)
            
                Grilla.ColWidth(0) = 0
                Grilla.ColWidth(1) = 1000
                Grilla.ColWidth(2) = 700
                Grilla.ColWidth(3) = 4300
            
            cabecera = "ID" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "DESCRIPCIOPN"
        
         Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id") & vbTab & Format(rst("fecha"), "dd-mm-YYYY") & vbTab & Format(rst("hora"), "HH:mm:ss") & vbTab & Trim(rst("observacion"))
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
  
  
     
  
         
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Public Sub llenarGrid_plan(ByVal Grilla As MSHFlexGrid, ByVal in_parametro As String)
Dim in_monto_acumulado As Double
Dim in_monto_deuda As Double

On Error GoTo salir

If Me.chk_tiendas_online.Value = 1 Then
        strCadena = "SELECT * FROM view_empresa_online WHERE ruc='" & KEY_RUC & "' and id_rubro='" & in_parametro & "' "
Else
    If in_parametro <> "" Then
        strCadena = "SELECT * FROM view_empresa_planes WHERE ruc='" & KEY_RUC & "' and (dni LIKE '%" & in_parametro & "%' OR nombre_completo LIKE '%" & in_parametro & "%' OR id_plan LIKE '%" & in_parametro & "%' OR id_rubro='" & in_parametro & "') "
    Else
        
        'strCadena = "SELECT * FROM view_empresa_planes WHERE id_plan='" & Me.DtcPlanes.BoundText & "' and  ruc='" & KEY_RUC & "'"
        strCadena = "call ADM_auditoria_empresa('1','" & Me.DtcPlanes.BoundText & "','','','','','','','" & KEY_RUC & "')"
    End If
End If

Call ConfiguraRst(strCadena)
 If Me.chk_tiendas_online.Value = 1 Then
    Me.lblcount_online.Caption = "EMP.AFILIADAS :  " & rst.RecordCount
 Else
    Me.lblacount.Caption = "EMP.AFILIADAS :  " & rst.RecordCount
 End If
 
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
  
   Grilla.Rows = 0
            ReDim arrColWidth(1 To rst.Fields.Count)
            
                Grilla.ColWidth(0) = 1200
                Grilla.ColWidth(1) = 2500
                Grilla.ColWidth(2) = 2800
                Grilla.ColWidth(3) = 1200
                Grilla.ColWidth(4) = 1200
                Grilla.ColWidth(5) = 1200
                Grilla.ColWidth(6) = 1200
                Grilla.ColWidth(7) = 1100
                Grilla.ColWidth(8) = 2100
                Grilla.ColWidth(9) = 1500
                Grilla.ColWidth(10) = 1300
                Grilla.ColWidth(11) = 1050
            
         cabecera = "DNI/RUC" & vbTab & "NOMBRE CLIENTE" & vbTab & "PLAN CONTRATADO" & vbTab & "CELULAR" & vbTab & "PERIODO" & vbTab & "PRECIO SERVICIO" & vbTab & "DEUDA FECHA" & vbTab & "FECHA CORTE" & vbTab & "ENCARGADO" & vbTab & "ESTADO" & vbTab & "F.VERIFICACION" & vbTab & "HORA."
         Grilla.AddItem cabecera
         For k = 0 To 11
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        in_monto_acumulado = 0
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            
            
            
             
             Fila = rst("dni") & vbTab & Trim(rst("nombre_completo")) & vbTab & rst("descripcion") & vbTab & rst("celular") & vbTab & rst("forma") & vbTab & Format(rst("monto"), "#,##0.00") & vbTab & Format(rst("saldo"), "#,##0.00") & vbTab & rst("fecha_corte") & vbTab & rst("operador") & vbTab & rst("estado") & vbTab & Format(rst("fecha_verificacion"), "dd-mm-YYYY") & vbTab & rst("hora_verificacion")
             Grilla.AddItem Fila
             
             Grilla.Row = i + 1
             If rst("id_estado") = "01" Then
                For k = 9 To 11
                    Grilla.col = k
                    Grilla.CellBackColor = &H8080FF
                Next k
                
                
             Else
             
                For k = 9 To 11
                    Grilla.col = k
                    Grilla.CellBackColor = &H80FF80
                Next k
            End If
            
            in_monto_acumulado = in_monto_acumulado + rst("monto")
            in_acumulado_deuda = in_acumulado_deuda + rst("saldo")
        rst.MoveNext
        
        Next i
        
         Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(in_monto_acumulado, "#,##0.00") & vbTab & Format(in_acumulado_deuda, "#,##0.00") & vbTab & "" & vbTab & "" & vbTab & ""
         Grilla.AddItem Fila
             
                For k = 5 To 6
                    Grilla.col = k
                    Grilla.Row = i + 1
                    Grilla.CellBackColor = &H8080FF
               Next k
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
     Me.cmdficha.Enabled = True
     Me.cmdupdate.Enabled = False
     Me.cmdDelete.Enabled = False
     
  Grilla.ColAlignment(0) = 1
  Grilla.ColAlignment(1) = 1
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
         
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Public Sub llenarGrid_investigacion(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
  
   Grilla.Rows = 0
  
         
 
           ReDim arrColWidth(1 To rst.Fields.Count)
           For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 1200
            Grilla.ColWidth(1) = 3500
            Grilla.ColWidth(2) = 7500
            Grilla.ColWidth(3) = 1500
            Grilla.ColWidth(4) = 1500
            Grilla.ColWidth(5) = 3000
            
          Next
            cabecera = "DNI" & vbTab & "NOMBRE TESISTA" & vbTab & "DESCRIPCION TESIS" & vbTab & "TIPO" & vbTab & "% AVANCE" & vbTab & "ASESOR"
            Grilla.AddItem cabecera
            For k = 0 To 5
                Grilla.col = k
                Grilla.Row = 0
                Grilla.CellBackColor = &HDFDFE0
            Next k
            rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            
            If rst("habilitado") = "si" Then
                in_habilitado = "ACTIVO"
            Else
                in_habilitado = "INACTIVO"
            End If
             Fila = rst("dni") & vbTab & UCase(rst("nombre_completo")) & vbTab & UCase(rst("nombre_prod")) & vbTab & rst("descripcion") & vbTab & "" & vbTab & rst("asesor")
             Grilla.AddItem Fila
            
        rst.MoveNext
        Next i
     Me.cmdupdate.Enabled = False
     Me.cmdDelete.Enabled = False
     
  Grilla.ColAlignment(0) = 1
  Grilla.ColAlignment(1) = 1
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
         
         
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub llenarGridContador(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
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
           Grilla.ColWidth(0) = 1100
           Grilla.ColWidth(1) = 5000
           Grilla.ColWidth(2) = 6000
           Grilla.ColWidth(3) = 1100
           
          Next
         cabecera = "DNI/RUC" & vbTab & "NOMBRE/RAZON SOCIAL" & vbTab & "DIRECCION FISCAL" & vbTab & "DEPARTAMENTO"
         Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             strCadena = "SELECT * FROM departamentos WHERE id_depa='" & rst("id_departamento") & "'"
             Call ConfiguraRstT(strCadena)
             If rstT.RecordCount > 0 Then
                departamento = UCase(rstT("descripcion"))
            Else
                departamento = "NO REGISTRADO"
             End If
             Fila = rst("dni") & vbTab & UCase(rst("nombre_completo")) & vbTab & UCase(rst("direccion")) & vbTab & departamento
             Grilla.AddItem Fila
            Fila = ""
        rst.MoveNext
        Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
         TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
         TlbAcciones.Buttons(KEY_NEW).Enabled = False
         TlbAcciones.Buttons(KEY_DELETE).Enabled = False
         TlbAcciones.Buttons(KEY_MAIL).Enabled = False
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub MATRICULA_Click()





End Sub

Private Sub HfdPersona_SelChange()
If Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0) <> "DNI/RUC" Then
    Me.cmdupdate.Enabled = True
Else
    Me.cmdupdate.Enabled = False
End If

End Sub

Private Sub Image1_Click()
Me.HfdPersona.Enabled = True
Me.frmValidar.Visible = False

End Sub

Private Sub Image2_Click()

Me.frmActividades.Visible = False

End Sub

Private Sub TxtApellido_KeyPress(KeyAscii As Integer)
    Call Mayusculas(KeyAscii)
in_cliente = "no"
    in_proveedor = "no"
    
    
    If Me.chk_cliente.Value = 1 Then
        in_cliente = "si"
    End If
    
    If Me.chk_proveedor.Value = 1 Then
        in_proveedor = "si"
    End If
    
    
    If in_cliente = "no" And in_proveedor = "no" Then
        in_cliente = "si"
    in_proveedor = "si"
    End If
    
    If KeyAscii = 13 Then
       If KEY_RUBRO = "00025" Then
            strCadena = "SELECT * FROM view_estudiante WHERE nombre_completo LIKE '%" & Trim(Me.TxtApellido.Text) & "%' and  id_empresa='" & KEY_RUC & "'"
       Else
            strCadena = "SELECT * FROM view_cliente WHERE nombre_completo LIKE '%" & Trim(Me.TxtApellido.Text) & "%' and (id_cliente='" & in_cliente & "' or  id_proveedor='" & in_proveedor & "')  and id_empresa='" & KEY_RUC & "'"
       End If
        
        Call llenarGrid(Me.HfdPersona)
    End If
End Sub



Private Sub TxtRazonSocial_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
Dim nruc As String
If KeyAscii = 13 Then

If Len(Me.txtRuc.Text) > 0 Then
      If KEY_RUBRO = "00025" Then
            strCadena = "SELECT * FROM view_estudiante WHERE dni= '" & Trim(Me.txtRuc.Text) & "' AND id_empresa='" & KEY_RUC & "' "
      Else
            strCadena = "SELECT * FROM view_cliente WHERE dni  LIKE  '" & Trim(Me.txtRuc.Text) & "%' and  id_cliente='si' AND id_empresa='" & KEY_RUC & "'"
      End If
      
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
       
        If Len(Trim(Me.txtRuc.Text)) = 8 Then
        If get_dni_reniec(Trim(Me.txtRuc.Text)) = True Then
            GoTo siguiente
        End If
        End If
siguiente:
        Procedencia = nuevo
        FrmDetallePersona.Show
       
        If Len(Trim(Me.txtRuc.Text)) = 8 Then
            strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
            Call ConfiguraRstK(strCadena)
            If rstK.RecordCount > 0 Then
               FrmDetallePersona.txtRuc.Text = rstK("dni")
               Call FrmDetallePersona.LLENA(rstK("dni"))
               Exit Sub
            Else
            nruc = "10" & Trim(Me.txtRuc.Text)
            FrmDetallePersona.txtRuc.Text = DigitoVerificadorRUC(Trim(nruc))
            End If
        Else
        FrmDetallePersona.txtRuc.Text = Trim(Me.txtRuc.Text)
        End If
        
        FrmDetallePersona.ChkCliente.Value = 1
        Call FrmDetallePersona.precionar
        Exit Sub
    Else
       ' If KEY_RUBRO = "00025" Then
       '       strCadena = "SELECT * FROM view_estudiante WHERE dni LIKE  '%" & Trim(Me.TxtRuc.Text) & "%' AND id_empresa='" & KEY_RUC & "' ORDER BY nombre_completo"
       ' Else
       '       strCadena = "SELECT * FROM view_cliente WHERE dni LIKE  '" & Trim(Me.TxtRuc.Text) & "%' AND id_empresa='" & KEY_RUC & "' ORDER BY nombre_completo LIMIT 100"
       ' End If
       Call llenarGrid(Me.HfdPersona)
       Exit Sub
    End If
Else
    
        If KEY_RUBRO = "00025" Then
            strCadena = "SELECT * FROM view_estudiante WHERE dni LIKE '%" & Trim(Me.txtRuc.Text) & "%' AND id_empresa='" & KEY_RUC & "' ORDER BY nombre_completo"
      Else
            strCadena = "SELECT * FROM view_cliente WHERE dni= '%" & Trim(Me.txtRuc.Text) & "' AND id_empresa='" & KEY_RUC & "' ORDER BY nombre_completo"
      End If
        Call llenarGrid(Me.HfdPersona)
   Exit Sub
End If
End If
End Sub


Private Sub TxtTelefono_KeyPress(KeyAscii As Integer)

End Sub

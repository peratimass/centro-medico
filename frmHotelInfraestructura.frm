VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frmHotelInfraestructura 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   16875
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmhabitacion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DETALLE HABITACION"
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
      Height          =   4815
      Left            =   8400
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox txtaccion_habitacion 
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
         Left            =   5280
         TabIndex        =   30
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtCodigoHabitacion 
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
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtprecio 
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
         Left            =   1680
         TabIndex        =   22
         Top             =   1800
         Width           =   2415
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
         Left            =   2640
         TabIndex        =   20
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtid_producto 
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
         Left            =   1680
         TabIndex        =   19
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtdescripcionhabitacion 
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
         Left            =   1680
         TabIndex        =   14
         Top             =   1320
         Width           =   3255
      End
      Begin VitekeySoft.ChameleonBtn cmd_save_habitacion 
         Height          =   615
         Left            =   1680
         TabIndex        =   15
         Top             =   4080
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1085
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
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHotelInfraestructura.frx":0000
         PICN            =   "frmHotelInfraestructura.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmd_cancelar_habitacion 
         Height          =   615
         Left            =   2880
         TabIndex        =   16
         Top             =   4080
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1085
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
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHotelInfraestructura.frx":3664
         PICN            =   "frmHotelInfraestructura.frx":3680
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtcEstadoHabitacion 
         Height          =   315
         Left            =   1680
         TabIndex        =   28
         Top             =   2760
         Width           =   2415
         _ExtentX        =   4260
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
      Begin MSDataListLib.DataCombo DtcPiso 
         Height          =   315
         Left            =   1680
         TabIndex        =   32
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
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
      Begin MSDataListLib.DataCombo DtcTipoHabitacion 
         Height          =   315
         Left            =   1680
         TabIndex        =   34
         Top             =   2280
         Width           =   2415
         _ExtentX        =   4260
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
      Begin MSDataListLib.DataCombo DtcSucursal 
         Height          =   315
         Left            =   1680
         TabIndex        =   35
         Top             =   3240
         Width           =   3855
         _ExtentX        =   6800
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
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "SUCURSAL:"
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
         Left            =   720
         TabIndex        =   36
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO  :"
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
         Left            =   600
         TabIndex        =   33
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PISO         :"
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
         TabIndex        =   31
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO :"
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
         Left            =   600
         TabIndex        =   29
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO  :"
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
         TabIndex        =   21
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SERVICIO  :"
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
         TabIndex        =   18
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame frmpiso 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DETALLE PISO"
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
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txtaccion 
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
         Left            =   2520
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin MSDataListLib.DataCombo DtcEstadoPiso 
         Height          =   315
         Left            =   1320
         TabIndex        =   26
         Top             =   1320
         Width           =   2415
         _ExtentX        =   4260
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
      Begin VB.TextBox txtcodigo_piso 
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
         Left            =   1320
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtDescripcionPiso 
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
         Left            =   1320
         TabIndex        =   10
         Top             =   840
         Width           =   2415
      End
      Begin VitekeySoft.ChameleonBtn cmdsave_piso 
         Height          =   615
         Left            =   1320
         TabIndex        =   11
         Top             =   1920
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1085
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
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHotelInfraestructura.frx":399A
         PICN            =   "frmHotelInfraestructura.frx":39B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdCancelar_piso 
         Height          =   615
         Left            =   2520
         TabIndex        =   12
         Top             =   1920
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   1085
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
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHotelInfraestructura.frx":6FFE
         PICN            =   "frmHotelInfraestructura.frx":701A
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
         BackStyle       =   0  'Transparent
         Caption         =   "ESTADO :"
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
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   14280
      OleObjectBlob   =   "frmHotelInfraestructura.frx":7334
      Top             =   3120
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPisos 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   13573
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
   Begin VitekeySoft.ChameleonBtn cmddelete 
      Height          =   855
      Left            =   5760
      TabIndex        =   1
      Top             =   1890
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1508
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmHotelInfraestructura.frx":7568
      PICN            =   "frmHotelInfraestructura.frx":7584
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
      Left            =   5760
      TabIndex        =   2
      Top             =   1005
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1508
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmHotelInfraestructura.frx":99CE
      PICN            =   "frmHotelInfraestructura.frx":99EA
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
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   900
      _ExtentX        =   1588
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmHotelInfraestructura.frx":9D04
      PICN            =   "frmHotelInfraestructura.frx":9D20
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfHabitacion 
      Height          =   7695
      Left            =   7560
      TabIndex        =   4
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   13573
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
   Begin VitekeySoft.ChameleonBtn cmdEliminarHabitacion 
      Height          =   855
      Left            =   15240
      TabIndex        =   5
      Top             =   1890
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1508
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
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmHotelInfraestructura.frx":A172
      PICN            =   "frmHotelInfraestructura.frx":A18E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdModificarHabitacion 
      Height          =   855
      Left            =   15240
      TabIndex        =   6
      Top             =   1005
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   1508
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmHotelInfraestructura.frx":C5D8
      PICN            =   "frmHotelInfraestructura.frx":C5F4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdNuevoHabitacion 
      Height          =   855
      Left            =   15240
      TabIndex        =   7
      Top             =   120
      Width           =   900
      _ExtentX        =   1588
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmHotelInfraestructura.frx":C90E
      PICN            =   "frmHotelInfraestructura.frx":C92A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image close 
      Height          =   240
      Left            =   16440
      Picture         =   "frmHotelInfraestructura.frx":CD7C
      Top             =   120
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8070
      Left            =   0
      Top             =   0
      Width           =   16875
   End
End
Attribute VB_Name = "frmHotelInfraestructura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub close_Click()

    Unload Me
    
End Sub

Private Sub cmd_cancelar_habitacion_Click()
Me.frmHabitacion.Visible = False
End Sub

Private Sub cmd_save_habitacion_Click()
If Trim(Me.txtdescripcionhabitacion.Text) <> "" And Trim(Me.txtid_producto.Text) <> "" And Trim(Me.txtproducto.Text) <> "" Then
    
    strCadena = "CALL put_hotel_habitacion('" & Val(Me.DtcPiso.BoundText) & "','" & Val(Me.txtCodigoHabitacion.Text) & "','" & Trim(Me.txtdescripcionhabitacion.Text) & "','" & Me.txtid_producto.Text & "','" & Me.DtcEstadoHabitacion.BoundText & "','" & Me.DtcTipoHabitacion.BoundText & "','" & Val(Me.txtaccion_habitacion.Text) & "','" & Me.DtcSucursal.BoundText & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Me.frmHabitacion.Visible = False
    Call Me.llenar_habitacion(Me.HfHabitacion, Val(Me.DtcPiso.BoundText))
    
End If
End Sub

Private Sub cmdCancelar_piso_Click()
Me.frmpiso.Visible = False
End Sub

Private Sub cmddelete_Click()
Procedencia = Eliminar
Call disabled_form(Me)
frmsegurity.Show

End Sub

Private Sub cmdEliminarHabitacion_Click()
Procedencia = eliminar_informe
Call disabled_form(Me)
frmsegurity.Show
End Sub

Private Sub cmdModificarHabitacion_Click()
Call get_load_habitacion(Me.HfHabitacion.TextMatrix(Me.HfHabitacion.Row, 0))
End Sub

Private Sub cmdNuevo_Click()
Call nuevo_piso
End Sub
Private Sub nuevo_piso()

Me.txtDescripcionPiso.Text = ""
Me.txtcodigo_piso.Text = ""
Me.txtcodigo_piso.Text = ""
Me.txtaccion.Text = "1"
Me.txtcodigo_piso.Locked = True
Me.frmpiso.Visible = True


End Sub

Private Sub cmdNuevoHabitacion_Click()
Call nuevo_habitacion
End Sub
Private Sub nuevo_habitacion()

Me.txtdescripcionhabitacion.Text = ""
Me.txtCodigoHabitacion.Text = ""
Me.txtaccion_habitacion.Text = "1"
Me.txtid_producto.Text = ""
Me.txtproducto.Text = ""
Me.txtprecio.Text = 0
Me.DtcEstadoHabitacion.BoundText = "01"

strCadena = "SELECT id_piso as Codigo, descripcion as Descripcion FROM hotel_piso "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPiso)
Me.DtcPiso.BoundText = Val(Me.HfPisos.TextMatrix(Me.HfPisos.Row, 0))


Me.frmHabitacion.Visible = True


End Sub
Private Sub cmdsave_piso_Click()



If Trim(Me.txtDescripcionPiso.Text) <> "" Then
    strCadena = "CALL put_hotel_piso('" & Val(Me.txtcodigo_piso.Text) & "','" & Trim(Me.txtDescripcionPiso.Text) & "','" & Me.DtcEstadoPiso.BoundText & "','" & Val(Me.txtaccion.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Me.frmpiso.Visible = False
    Call llenar_piso(Me.HfPisos)
    
End If

End Sub
Public Sub put_delete_piso()

    strCadena = "CALL put_hotel_piso('" & Val(Me.HfPisos.TextMatrix(Me.HfPisos.Row, 0)) & "','','','3','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Call llenar_piso(Me.HfPisos)

End Sub


Public Sub put_delete_habitacion()

    strCadena = "CALL put_hotel_habitacion('0','" & Val(Me.HfHabitacion.TextMatrix(Me.HfHabitacion.Row, 0)) & "','','','','','3','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    Call Me.llenar_habitacion(Me.HfHabitacion, Val(Me.DtcPiso.BoundText))

End Sub
Private Sub cmdupdate_Click()
    Call get_load(Me.HfPisos.TextMatrix(Me.HfPisos.Row, 0))
End Sub
Private Sub get_load(ByVal In_Piso As String)
strCadena = "SELECT * FROM view_hotel_piso WHERE id_piso='" & Val(In_Piso) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txtcodigo_piso.Text = rst("id_piso")
    Me.txtDescripcionPiso.Text = rst("descripcion")
    Me.DtcEstadoPiso.BoundText = rst("id_estado")
    Me.txtaccion.Text = 2
    Me.frmpiso.Visible = True
End If
End Sub
Private Sub get_load_habitacion(ByVal in_habitacion As String)
strCadena = "SELECT id_piso as Codigo, descripcion as Descripcion FROM hotel_piso "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcPiso)
Me.DtcPiso.BoundText = Val(Me.HfPisos.TextMatrix(Me.HfPisos.Row, 0))

strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcSucursal)



strCadena = "SELECT * FROM view_habitacion WHERE id_habitacion='" & Val(in_habitacion) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.DtcPiso.BoundText = rst("id_piso")
    Me.txtCodigoHabitacion.Text = rst("id_habitacion")
    Me.txtdescripcionhabitacion.Text = rst("descripcion")
    Me.txtid_producto.Text = rst("id_producto")
    Me.txtproducto.Text = get_producto(rst("id_producto"))
    Me.txtprecio.Text = get_precio_producto(rst("id_producto"), KEY_ALM)
    Me.DtcEstadoHabitacion.BoundText = rst("id_estado")
    Me.DtcTipoHabitacion.BoundText = rst("id_tipo")
    Me.txtaccion_habitacion.Text = 2
    Me.DtcSucursal.BoundText = rst("id_alm")
    Me.frmHabitacion.Visible = True
End If
End Sub
Private Sub Form_Load()
CenterForm Me
Me.Top = 100
'Call load_skin("in_skin")

strCadena = "SELECT id_estado as Codigo, descripcion as Descripcion FROM hotel_estado "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstadoPiso)
Me.DtcEstadoPiso.BoundText = "01"

strCadena = "SELECT id_estado as Codigo, descripcion as Descripcion FROM hotel_estado "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstadoHabitacion)
Me.DtcEstadoHabitacion.BoundText = "01"


strCadena = "SELECT id_tipo_habitacion as Codigo, descripcion as Descripcion FROM hotel_habitacion_tipo "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoHabitacion)
Me.DtcTipoHabitacion.BoundText = "1"


Call llenar_piso(Me.HfPisos)

End Sub
Private Sub load_skin(ByVal in_skin As String)
Skin1.LoadSkin App.Path & "\Skins\" & in_skin & ".skn"
Skin1.ApplySkin Me.hwnd
End Sub
Public Sub llenar_piso(ByVal Grilla As MSHFlexGrid)
strCadena = "SELECT * FROM view_piso WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 2500
           Grilla.ColWidth(2) = 700
           
        Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "ESTADO"
        Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
 
        NumeroCampo = 6
            
        For i = 0 To rstI.RecordCount - 1
          
          Fila = Format(rstI("id_piso"), "00") & vbTab & rstI("descripcion") & vbTab & rstI("estado")
          Grilla.AddItem Fila
        
       
          
          rstI.MoveNext
      Next i
    
End Sub


Public Sub llenar_habitacion(ByVal Grilla As MSHFlexGrid, ByVal In_Piso As String)
strCadena = "SELECT * FROM view_habitacion WHERE  id_piso='" & Val(In_Piso) & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 1000
        Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "PRECIO" & vbTab & "ESTADO"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
 
      
            
        For i = 0 To rstI.RecordCount - 1
          
          Fila = Format(rstI("id_habitacion"), "00") & vbTab & rstI("descripcion") & vbTab & Format(get_precio_producto(rstI("id_producto"), KEY_ALM), "#,##0.00") & vbTab & rstI("estado")
          Grilla.AddItem Fila
          rstI.MoveNext
      Next i
    
End Sub


Private Sub HfHabitacion_SelChange()
If Val(Me.HfHabitacion.TextMatrix(Me.HfHabitacion.Row, 0)) > 0 Then
   Me.cmdModificarHabitacion.Enabled = True
   Me.cmdEliminarHabitacion.Enabled = True
Else
   Me.cmdModificarHabitacion.Enabled = False
   Me.cmdEliminarHabitacion.Enabled = False
End If
End Sub

Private Sub HfPisos_SelChange()
If Val(Me.HfPisos.TextMatrix(Me.HfPisos.Row, 0)) > 0 Then
   Me.cmdupdate.Enabled = True
   Me.cmddelete.Enabled = True
   Call llenar_habitacion(Me.HfHabitacion, Val(Me.HfPisos.TextMatrix(Me.HfPisos.Row, 0)))
Else
   Me.cmdupdate.Enabled = False
   Me.cmddelete.Enabled = False
   Me.HfHabitacion.Rows = 0
End If
End Sub

Private Sub txtid_producto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
End If
End Sub

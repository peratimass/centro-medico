VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPrecios 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Actualizacion de Precios"
   ClientHeight    =   9240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20595
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk_appvendedor 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "FUERZA DE VENTAS"
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
      Height          =   300
      Left            =   17280
      TabIndex        =   60
      Top             =   3360
      Width           =   2355
   End
   Begin VB.CheckBox chk_appdelivery 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "DELIVERY"
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
      Height          =   300
      Left            =   17280
      TabIndex        =   59
      Top             =   3000
      Width           =   2355
   End
   Begin VB.CheckBox chk_oferta 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "OFERTA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   17280
      TabIndex        =   58
      Top             =   2280
      Width           =   2280
   End
   Begin MSComCtl2.DTPicker Dtpkardex 
      Height          =   300
      Left            =   18600
      TabIndex        =   57
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
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
      Format          =   136249345
      CurrentDate     =   43655
   End
   Begin VB.CheckBox chk_kardex 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "AFECTAR KARDEX"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   17280
      TabIndex        =   56
      Top             =   2640
      Width           =   1275
   End
   Begin VB.TextBox txtprecio_mayor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   330
      Left            =   16035
      TabIndex        =   55
      Top             =   2600
      Width           =   1080
   End
   Begin VB.TextBox txtprecio_alterno_a 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   330
      Left            =   16035
      TabIndex        =   52
      Top             =   2160
      Width           =   1080
   End
   Begin VB.Frame frmagranel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRECIO A GRANEL"
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
      Height          =   2100
      Left            =   14520
      TabIndex        =   43
      Top             =   6360
      Width           =   5535
      Begin VB.Frame frmprecio_detalle 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   2400
         TabIndex        =   46
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
         Begin VB.TextBox txtPrecio_unidad 
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
            Left            =   1080
            TabIndex        =   48
            Top             =   240
            Width           =   1215
         End
         Begin VitekeySoft.ChameleonBtn cmdprocesar_unidad 
            Height          =   315
            Left            =   1080
            TabIndex        =   49
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
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
            MICON           =   "FrmPrecios.frx":0000
            PICN            =   "FrmPrecios.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO :"
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
            Left            =   105
            TabIndex        =   47
            Top             =   360
            Width           =   585
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfUnidades 
         Height          =   1575
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2778
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
      Begin VitekeySoft.ChameleonBtn cmdupdateprecio 
         Height          =   555
         Left            =   5040
         TabIndex        =   45
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   979
         BTYPE           =   5
         TX              =   ""
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
         MICON           =   "FrmPrecios.frx":2370
         PICN            =   "FrmPrecios.frx":238C
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
   Begin VB.CheckBox chk_sucursales 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "TODAS LAS SUCURSALES"
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
      Height          =   300
      Left            =   17280
      TabIndex        =   42
      Top             =   1920
      Width           =   2280
   End
   Begin VB.CheckBox chk_igv 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "AFECTO A IGV"
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
      Height          =   300
      Left            =   17280
      TabIndex        =   39
      Top             =   1560
      Width           =   2280
   End
   Begin VB.TextBox txtvalor_venta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   330
      Left            =   16035
      TabIndex        =   37
      Top             =   1680
      Width           =   1080
   End
   Begin VitekeySoft.ChameleonBtn cmdprocesar 
      Height          =   525
      Left            =   17040
      TabIndex        =   35
      Top             =   8520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   926
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
      MICON           =   "FrmPrecios.frx":48B7
      PICN            =   "FrmPrecios.frx":48D3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkHabilitado 
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
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   17280
      TabIndex        =   34
      Top             =   1200
      Width           =   2280
   End
   Begin VB.TextBox txtcant_fin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   2
      Left            =   19320
      TabIndex        =   20
      Top             =   5925
      Width           =   600
   End
   Begin VB.TextBox txtcant_ini 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   2
      Left            =   18000
      TabIndex        =   19
      Top             =   5925
      Width           =   600
   End
   Begin VB.TextBox txtcant_fin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   1
      Left            =   19320
      TabIndex        =   17
      Top             =   5205
      Width           =   600
   End
   Begin VB.TextBox txtcant_ini 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   1
      Left            =   18000
      TabIndex        =   16
      Top             =   5205
      Width           =   600
   End
   Begin VB.TextBox txtcant_fin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   0
      Left            =   19320
      TabIndex        =   14
      Top             =   4605
      Width           =   600
   End
   Begin VB.TextBox txtcant_ini 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   0
      Left            =   18000
      TabIndex        =   13
      Top             =   4605
      Width           =   600
   End
   Begin VB.TextBox TxtCosto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   330
      Left            =   16035
      TabIndex        =   23
      Top             =   1275
      Width           =   1080
   End
   Begin VB.TextBox TxtPrecio 
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
      Index           =   1
      Left            =   15960
      TabIndex        =   15
      Top             =   5205
      Width           =   1320
   End
   Begin VB.TextBox TxtPrecio 
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
      Index           =   2
      Left            =   15960
      TabIndex        =   18
      Top             =   5880
      Width           =   1320
   End
   Begin VB.TextBox TxtPrecioVenta 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   330
      Left            =   16035
      TabIndex        =   11
      Top             =   3030
      Width           =   1080
   End
   Begin VB.TextBox TxtDescripcion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   15915
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   750
      Width           =   3975
   End
   Begin VB.TextBox TxtPrecio 
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
      Index           =   0
      Left            =   15960
      TabIndex        =   12
      Top             =   4605
      Width           =   1320
   End
   Begin VB.TextBox TxtProducto 
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
      Left            =   4320
      TabIndex        =   1
      Top             =   8655
      Width           =   2055
   End
   Begin VB.TextBox TxtCod 
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
      Left            =   1020
      TabIndex        =   0
      Top             =   8655
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   8055
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   14208
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
   Begin MSDataListLib.DataCombo DtcClasificacion 
      Height          =   330
      Left            =   7965
      TabIndex        =   26
      Top             =   8655
      Width           =   4215
      _ExtentX        =   7435
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
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   555
      Left            =   18600
      TabIndex        =   36
      Top             =   8520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   979
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
      MICON           =   "FrmPrecios.frx":7F1B
      PICN            =   "FrmPrecios.frx":7F37
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdmodificar 
      Height          =   855
      Left            =   13380
      TabIndex        =   40
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPrecios.frx":8327
      PICN            =   "FrmPrecios.frx":8343
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
      Height          =   855
      Left            =   13380
      TabIndex        =   41
      Top             =   1200
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPrecios.frx":A97C
      PICN            =   "FrmPrecios.frx":A998
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcMoneda 
      Height          =   330
      Left            =   15915
      TabIndex        =   50
      Top             =   3720
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
   Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
      Height          =   405
      Left            =   14520
      TabIndex        =   61
      Top             =   8520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   714
      BTYPE           =   5
      TX              =   "LOAD AMAZON"
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
      MICON           =   "FrmPrecios.frx":AD88
      PICN            =   "FrmPrecios.frx":ADA4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar prog_indicador 
      Height          =   225
      Left            =   14520
      TabIndex        =   62
      Top             =   8880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.MAYORISTA :"
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
      Left            =   14625
      TabIndex        =   54
      Top             =   2640
      Width           =   1155
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.MERCADO:"
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
      Left            =   14805
      TabIndex        =   53
      Top             =   2175
      Width           =   975
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MONEDA :"
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
      Left            =   14955
      TabIndex        =   51
      Top             =   3840
      Width           =   825
   End
   Begin VB.Label lblpreciodelivery 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO DELIVERY :"
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
      Left            =   14580
      TabIndex        =   38
      Top             =   3045
      Width           =   1425
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTRE"
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
      Left            =   18735
      TabIndex        =   33
      Top             =   6000
      Width           =   525
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTRE"
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
      Left            =   18735
      TabIndex        =   32
      Top             =   5280
      Width           =   525
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENTRE"
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
      Left            =   18735
      TabIndex        =   31
      Top             =   4680
      Width           =   525
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "CANTIDADES"
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
      Left            =   18420
      TabIndex        =   30
      Top             =   4320
      Width           =   1005
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.ALTERNO III :"
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
      Left            =   14640
      TabIndex        =   29
      Top             =   5880
      Width           =   1125
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.ALTERNO II :"
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
      Left            =   14655
      TabIndex        =   28
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLASIFICACION :"
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
      Left            =   6675
      TabIndex        =   27
      Top             =   8655
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO COSTO :"
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
      Left            =   14565
      TabIndex        =   24
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.BARRA :"
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
      Left            =   17520
      TabIndex        =   22
      Top             =   360
      Width           =   765
   End
   Begin VB.Label lblbarra 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   18360
      TabIndex        =   21
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label LblDescripcion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION :"
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
      Left            =   14760
      TabIndex        =   10
      Top             =   780
      Width           =   1125
   End
   Begin VB.Label LblReorden 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR VENTA :"
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
      Left            =   14595
      TabIndex        =   9
      Top             =   1800
      Width           =   1185
   End
   Begin VB.Label LblCodigoProducto 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   15915
      TabIndex        =   8
      Top             =   405
      Width           =   1425
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO :"
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
      Left            =   15150
      TabIndex        =   7
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.ALTERNO I :"
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
      Left            =   14640
      TabIndex        =   6
      Top             =   4575
      Width           =   1035
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION :"
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
      Left            =   2955
      TabIndex        =   4
      Top             =   8655
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO :"
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
      Left            =   210
      TabIndex        =   3
      Top             =   8655
      Width           =   735
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MANTENIMIENTO PRECIOS"
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
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   120
      Top             =   8400
      Width           =   13215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3975
      Left            =   14520
      Top             =   240
      Width           =   5505
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   765
      Left            =   14520
      Top             =   4245
      Width           =   5505
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   14520
      Top             =   5040
      Width           =   5505
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   660
      Left            =   14520
      Top             =   5655
      Width           =   5505
   End
End
Attribute VB_Name = "FrmPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub ChameleonBtn1_Click()
Dim in_ruta As String
strCadena = "SELECT * FROM producto WHERE ruc='" & KEY_RUC & "' ORDER BY id_producto ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   Me.prog_indicador.Value = 0
   Me.prog_indicador.Max = rst.RecordCount
   For i = 0 To rst.RecordCount - 1
        in_ruta = "https://miparadita.s3.amazonaws.com/articulos/" + rst("id_producto") + ".jpg"
        strCadena = "UPDATE producto SET imagen='" & Trim(in_ruta) & "' WHERE id_producto='" & rst("id_producto") & "' and ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        DoEvents
        Me.prog_indicador.Value = i
         DoEvents
        rst.MoveNext
   Next i
End If
MsgBox "Carga Completa", vbInformation

End Sub

Private Sub chk_kardex_Click()
If Me.chk_kardex.Value = 1 Then
   Me.Dtpkardex.Visible = True
   Me.Dtpkardex.Value = KEY_FECHA
Else
  Me.Dtpkardex.Visible = False
End If

End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdModificar_Click()
 
 Procedencia = modificar
 FrmSeguridad.Show
 
 
End Sub

Private Sub put_costo_producto(ByVal in_producto As String, ByVal in_costo As Double)

strCadena = "SELECT * FROM almacen_producto WHERE ruc='" & KEY_RUC & "' and id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' LIMIT 1"
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
    strCadena = "SELECT id_kardex FROM kardex WHERE fecha_emision>='" & Format(Me.Dtpkardex.Value, "YYYY-mm-dd") & "' and  id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC , id_kardex DESC "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
        
        strCadena = "UPDATE kardex SET   costo_promedio='" & in_costo & "',costo_unitario='" & in_costo & "' WHERE id_kardex='" & rst("id_kardex") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        
        rst.MoveNext
        Next i
    Else
        strCadena = "SELECT id_kardex FROM kardex WHERE   id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' ORDER BY fecha_emision DESC , id_kardex DESC LIMIT 10"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
        strCadena = "UPDATE kardex SET   costo_promedio='" & in_costo & "',costo_unitario='" & in_costo & "' WHERE id_kardex='" & rst("id_kardex") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        CnBd.Execute (strCadena)
        rst.MoveNext
        Next i
        End If
        
    End If
End If


End Sub


Private Sub cmdProcesar_Click()
Call put_cambio_precio(Trim(LblCodigoProducto.Caption), Val(TxtPrecioVenta.Text), "MODULO DE CAMBIO DE PRECIOS [USUARIO Y CLAVE]")

Call Save

If KEY_RUC = "20480460538" Then
    Call put_costo_producto(Trim(Me.LblCodigoProducto.Caption), Val(Me.TxtCosto.Text))
End If


If KEY_RUC = "20479779598" Then
    Call put_costo_producto(Trim(Me.LblCodigoProducto.Caption), Val(Me.TxtCosto.Text))
End If


If KEY_RUC = "20522299953" Then
    Call put_costo_producto(Trim(Me.LblCodigoProducto.Caption), Val(Me.TxtCosto.Text))
End If



End Sub

Private Sub cmdprocesar_unidad_Click()
strCadena = "UPDATE producto_unidad SET precio='" & Val(Me.txtPrecio_unidad.Text) & "' WHERE id='" & Val(Me.HfUnidades.TextMatrix(Me.HfUnidades.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
Me.frmprecio_detalle.Visible = False
Call llenarUnidad(Me.HfUnidades, Trim(Me.LblCodigoProducto.Caption))
End Sub

Private Sub cmdSalir_Click()

Me.Width = 14475



Exit Sub





End Sub

Private Sub cmdupdateprecio_Click()
If Me.frmprecio_detalle.Visible = True Then
    
    Me.frmprecio_detalle.Visible = False
Else
    Me.txtPrecio_unidad.Text = Val(Me.HfUnidades.TextMatrix(Me.HfUnidades.Row, 3))
    Me.frmprecio_detalle.Visible = True
End If


End Sub

Private Sub Command1_Click()



End Sub

Private Sub DtcClasificacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

     strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  id_linea ='" & Trim(Me.DtcClasificacion.BoundText) & "' ORDER BY nombre_prod LIMIT 0,100 "
     Call llenarGrid_prod(Me.HfdDetalle, Me)


End If
End Sub

Private Sub Form_Load()
Call CenterForm(Me)
Me.Top = 50

  If KEY_RUC = "20511427721" Then
        If KEY_ALM = "00001" Then
                Me.lblpreciodelivery.Caption = "P.VENTA:"
        Else
                Me.lblpreciodelivery.Caption = "P.DELIVERY:"
        End If
  End If
  
  
  strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcClasificacion)
  
  strCadena = "SELECT id_moneda FROM almacen WHERE id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
        strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM moneda"
        Call ConfiguraRstT(strCadena)
        Call LlenaDataComboT(Me.DtcMoneda)
        Me.DtcMoneda.BoundText = rst("id_moneda")
        Me.DtcMoneda.Locked = True
  End If
  
  
  
  
  
Call ActualizarProd
End Sub
Sub ActualizarProd()
If frmVariacionCostes.Procedencia = buscar Then
Dim descri As String
descri = Mid(Trim(frmVariacionCostes.HfgFacturas.TextMatrix(frmVariacionCostes.HfgFacturas.Row, 2)), 1, 12)
  If KEY_BARRAS = "si" Then
   strCadena = "SELECT P.id_producto, PB.cod_barra,P.nombre_prod,U.abreviatura, P.precio_venta,P.precio_compra,P.id_igv " & _
  "FROM producto P, producto_barras PB,unidad U WHERE P.id_producto=PB.id_producto AND P.ruc='" & KEY_RUC & "' AND PB.ruc='" & KEY_RUC & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND   PB.cod_barra LIKE '%" & descri & "%' ORDER BY nombre_prod LIMIT 0,27"
  Else
  'strCadena = "SELECT P.id_producto,P.nombre_prod,U.abreviatura, P.precio_venta,P.precio_compra,P.id_igv " & _
  "FROM producto P,unidad U WHERE P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND P.id_producto LIKE '%" & descri & "%' ORDER BY nombre_prod LIMIT 0,27"
  strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY nombre_prod LIMIT 0,75 "
  'Call llenarGrid(Me.HfdGrilla, strCadena)
  End If
Else

 'strCadena = "SELECT P.id_producto,P.nombre_prod,U.abreviatura,P.precio_venta,P.precio_compra,P.id_igv " & _
  "FROM producto P,unidad U WHERE P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' ORDER BY P.nombre_prod LIMIT 0,27"
  strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' ORDER BY nombre_prod LIMIT 0,50 "

  End If
  Call llenarGrid_prod(Me.HfdDetalle, Me)
End Sub
Sub llenarGrid_alm(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.Rows = rst.RecordCount
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 0
  Grilla.ColWidth(1) = 3900
  Grilla.Enabled = False
  Grilla.Refresh
End Sub
Sub llenarGrid_prod(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
Dim in_precio As String
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Me.cmdmodificar.Enabled = False
    Exit Sub
End If
  
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 650
           Grilla.ColWidth(1) = 6900
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 0
           Grilla.ColWidth(4) = 0
           Grilla.ColWidth(5) = 1000
           Grilla.ColWidth(6) = 1000
           Grilla.ColWidth(7) = 900
           Grilla.ColWidth(8) = 1000
    Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "CLASIFICACION" & vbTab & "MODELO" & vbTab & "COLOR" & vbTab & "UND" & vbTab & "MARCA" & vbTab & "P.VENTA" & vbTab & "P.COSTO"
         Grilla.AddItem cabecera
         For k = 0 To 8
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
           
              
               in_precio = rst("precio_venta")
       
            
            Fila = rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & UCase(rst("linea")) & vbTab & UCase(rst("modelo")) & vbTab & rst("color") & vbTab & rst("unidad") & vbTab & rst("marca") & vbTab & Format(in_precio, "#,##0.00") & vbTab & Format(rst("precio_compra"), "#,##0.00")
            Grilla.AddItem Fila
            If (rst("id_igv") = "no") Then
                            For k = 0 To 5
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &H80FFFF
                           Next k
            End If
        
        rst.MoveNext
        Next i
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub HfdGrilla_Click()

End Sub

Private Sub HfdDetalle_SelChange()
Dim in_agranel As String
If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    
strCadena = "SELECT * FROM view_producto  WHERE id_producto='" & Trim(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "' AND id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)


Me.LblCodigoProducto.Caption = Trim(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))
If KEY_BARRAS = "si" Then
    Me.lblbarra.Caption = BDBuscarCampoRuc("producto_barras", "cod_barra", "id_producto", Trim(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)))
End If
in_agranel = rst("agranel")
Me.TxtDescripcion.Text = rst("nombre_prod")
Me.TxtCosto.Text = Format(rst("precio_compra"), "###0.00000000")
Me.TxtPrecioVenta.Text = Format(rst("precio_venta"), "###0.00")

If rst("id_igv") = "si" Then
   Me.chk_igv.Value = 1
  Me.txtvalor_venta.Text = Format(rst("precio_venta") / (1 + KEY_IGV), "###0.0000")
Else
   Me.chk_igv.Value = 0
   Me.TxtPrecioVenta.Text = Format(rst("precio_venta"), "###0.00")
End If

If rst("oferta") = "si" Then
   Me.chk_oferta.Value = 1
Else
   Me.chk_oferta.Value = 0
End If

If rst("app_delivery") = "si" Then
    Me.chk_appdelivery.Value = 1
Else
    Me.chk_appdelivery.Value = 0
End If


If rst("app_vendedor") = "si" Then
    Me.chk_appvendedor.Value = 1
Else
    Me.chk_appvendedor.Value = 0
End If







Me.txtprecio_alterno_a.Text = Format(rst("precio_alterno_a"), "###0.00")
Me.txtprecio_mayor.Text = Format(rst("precio_mayor"), "###0.00")

If rst("habilitado") = "si" Then
   Me.chkHabilitado.Value = 1
Else
   Me.chkHabilitado.Value = 0
End If

For i = 0 To 2
        Me.txtcant_ini(i).Text = ""
        Me.txtcant_fin(i).Text = ""
        Me.TxtPrecio(i).Text = ""
Next i


strCadena = "SELECT * FROM almacen_producto_precio WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' ORDER BY id_detalle ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        Me.txtcant_ini(i).Text = ""
        Me.txtcant_fin(i).Text = ""
        Me.TxtPrecio(i).Text = ""
        Me.TxtPrecio(i).Text = rst("precio")
        Me.txtcant_ini(i).Text = rst("cant_ini")
        Me.txtcant_fin(i).Text = rst("cant_fin")
        rst.MoveNext
    Next i
    
End If

Call Resalta(Me.TxtPrecioVenta)
If in_agranel = "si" Then
    
    Call llenarUnidad(Me.HfUnidades, Trim(Me.LblCodigoProducto.Caption))
    Me.frmagranel.Visible = True
Else
    Me.frmagranel.Visible = False
End If



Me.cmdmodificar.Enabled = True
Me.cmdprocesar.Enabled = True


Else
Me.cmdprocesar.Enabled = False
    
End If
End Sub
Private Sub llenarUnidad(ByVal Grilla As MSHFlexGrid, ByVal in_producto As String)
strCadena = "SELECT * FROM view_unidad_producto WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
   
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2000
           Grilla.ColWidth(2) = 1300
           Grilla.ColWidth(3) = 1300
        Next
         cabecera = "CODIGO" & vbTab & "UNIDAD" & vbTab & "EQUIVALE" & vbTab & "PRECIO"
         Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id") & vbTab & rst("descripcion") & vbTab & rst("cantidad") & vbTab & Format(rst("precio"), "###0.00")
             Grilla.AddItem Fila
             rst.MoveNext
        Next i
  Exit Sub
End Sub


 Sub Resize()
        Me.Width = 20145
        CenterForm Me
        Me.Top = 150
End Sub
Sub LLenaDatos()
strCadena = "SELECT * FROM view_producto WHERE id_producto='" & Trim(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "' AND id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Me.LblCodigoProducto.Caption = Trim(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))
Me.TxtDescripcion.Text = UCase(rst("nombre_prod"))
Me.TxtPrecioVenta.Text = Format(rst("precio_venta"), "###0.00")
Me.TxtCosto.Text = Format(rst("precio_compra"), "###0.00")
For i = 0 To 2
        Me.txtcant_ini(i).Text = ""
        Me.txtcant_fin(i).Text = ""
        Me.TxtPrecio(i).Text = ""
Next i


strCadena = "SELECT * FROM almacen_producto_precio WHERE id_producto='" & Trim(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "' AND id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "' ORDER BY id_detalle ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        Me.txtcant_ini(i).Text = ""
        Me.txtcant_fin(i).Text = ""
        Me.TxtPrecio(i).Text = ""
        Me.TxtPrecio(i).Text = rst("precio")
        Me.txtcant_ini(i).Text = rst("cant_ini")
        Me.txtcant_fin(i).Text = rst("cant_fin")
        rst.MoveNext
    Next i
 End If
Me.TxtPrecioVenta.SetFocus

End Sub


Private Sub Save()
Dim habilitado As String
If Me.chkHabilitado.Value = 1 Then
    habilitado = "si"
Else
    habilitado = "no"
End If

If Me.chk_igv.Value = 1 Then
   in_igv = "si"
Else
    in_igv = "no"
End If

If Me.chk_oferta.Value = 1 Then
   in_oferta = "si"
Else
   in_oferta = "no"
End If

If Me.chk_appdelivery.Value = 1 Then
    in_appdelivery = "si"
Else
    in_appdelivery = "no"

End If

If Me.chk_appvendedor.Value = 1 Then
   in_appvendedor = "si"
Else
   in_appvendedor = "no"
End If





strCadena = "UPDATE producto SET oferta='" & in_oferta & "',id_igv='" & in_igv & "' WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


'Verificacion precio venta bajo el costo

If Val(Me.TxtPrecioVenta.Text) < Val(Me.TxtCosto.Text) Then
    MsgBox "Precio de Venta NO DEBE SER MENOR QUE EL COSTO", vbInformation, KEY_VENDEDOR
    Exit Sub
End If

If Me.chk_sucursales.Value = 1 Then
    strCadena = "UPDATE almacen_producto SET app_delivery='" & in_appdelivery & "',app_vendedor='" & in_appvendedor & "',precio_venta='" & Format(Me.TxtPrecioVenta.Text, "###0.00") & "',precio_compra='" & Format(Me.TxtCosto.Text, "###0.00000000") & "',precio_mayor='" & Val(Me.TxtPrecio(0).Text) & "',habilitado='" & habilitado & "' WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "'  AND ruc='" & KEY_RUC & "'"
Else
    strCadena = "UPDATE almacen_producto SET app_delivery='" & in_appdelivery & "',app_vendedor='" & in_appvendedor & "', precio_venta='" & Format(Me.TxtPrecioVenta.Text, "###0.00") & "',precio_compra='" & Format(Me.TxtCosto.Text, "###0.00000000") & "',precio_mayor='" & Val(Me.txtprecio_mayor.Text) & "',precio_alterno_a='" & Val(Me.txtprecio_alterno_a.Text) & "',habilitado='" & habilitado & "' WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
End If
CnBd.Execute (strCadena)






strCadena = "DELETE FROM almacen_producto_precio WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND id_alm='" & KEY_ALM & "' AND ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


If Val(Me.txtcant_ini(0).Text) > 0 And Val(Me.txtcant_fin(0).Text) > 0 Then
    strCadena = "INSERT INTO almacen_producto_precio(id_producto,id_alm,precio,cant_ini,cant_fin,ruc)VALUES('" & Trim(Me.LblCodigoProducto.Caption) & "','" & KEY_ALM & "','" & Val(Me.TxtPrecio(0).Text) & "','" & Val(Me.txtcant_ini(0).Text) & "','" & Val(Me.txtcant_fin(0).Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)

End If

If Val(Me.txtcant_ini(1).Text) > 0 And Val(Me.txtcant_fin(1).Text) > 0 Then
    strCadena = "INSERT INTO almacen_producto_precio(id_producto,id_alm,precio,cant_ini,cant_fin,ruc)VALUES('" & Trim(Me.LblCodigoProducto.Caption) & "','" & KEY_ALM & "','" & Val(Me.TxtPrecio(1).Text) & "','" & Val(Me.txtcant_ini(1).Text) & "','" & Val(Me.txtcant_fin(1).Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)

End If

If Val(Me.txtcant_ini(2).Text) > 0 And Val(Me.txtcant_fin(2).Text) > 0 Then
    strCadena = "INSERT INTO almacen_producto_precio(id_producto,id_alm,precio,cant_ini,cant_fin,ruc)VALUES('" & Trim(Me.LblCodigoProducto.Caption) & "','" & KEY_ALM & "','" & Val(Me.TxtPrecio(2).Text) & "','" & Val(Me.txtcant_ini(2).Text) & "','" & Val(Me.txtcant_fin(2).Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)

End If


Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 7) = Format(Val(Me.TxtPrecioVenta.Text), "###0.00")
Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 8) = Format(Val(Me.TxtCosto.Text), "###0.00")




End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub HfUnidades_SelChange()
If Val(Me.HfUnidades.TextMatrix(Me.HfUnidades.Row, 0)) > 0 Then
   Me.cmdupdateprecio.Enabled = True
Else
   Me.cmdupdateprecio.Enabled = False
End If
End Sub

Private Sub TxtCod_KeyPress(KeyAscii As Integer)
 Dim Registros As Integer
Dim Criterio As String


If KeyAscii = 13 Then
        If Me.TxtCod.Text = "" Then
        Call ActualizarProd
        Exit Sub
  Else
    If Len(Me.TxtCod.Text) > 0 Then
        If KEY_BARRAS = "no" Then
    Criterio = " id_producto LIKE '%" & Trim(Me.TxtCod.Text) & "%'"
  Else
    Criterio = "(cod_barra LIKE  '%" & Trim(Me.TxtCod.Text) & "%' or id_producto LIKE '%" & Trim(Me.TxtCod.Text) & "%' or  codigo_proveedor LIKE '%" & Trim(Me.TxtCod.Text) & "%' or  id_universal LIKE '%" & Trim(Me.TxtCod.Text) & "%' or codigo_alterno LIKE '%" & Trim(Me.TxtCod.Text) & "%')"
  End If
  
   If KEY_SKFACTURA = "no" Then
      If KEY_BARRAS = "si" Then
        strCadena = "SELECT * FROM view_producto_barras WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY nombre_prod LIMIT 50 "
      Else
        strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY nombre_prod LIMIT 0,50 "
      End If
        Call llenarGrid_prod(Me.HfdDetalle, Me)
    Else
      If KEY_BARRAS = "si" Then
    strCadena = "SELECT * FROM view_producto_barras WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND  " & Criterio & " ORDER BY nombre_prod LIMIT 50 "
    Else
   
     strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND id_producto LIKE '%" & Trim(Me.TxtCod.Text) & "%' ORDER BY nombre_prod LIMIT 0,75 "
     Call Me.llenarGrid_prod(Me.HfdDetalle, Me)
    
    End If
    Call llenarGrid_prod(Me.HfdDetalle, Me)
  End If
End If
  
    Me.TxtCod.SetFocus
  End If
End If
End Sub

Private Sub TxtPrecioVenta_Change()
 Me.txtvalor_venta.Text = Format(Val(Me.TxtPrecioVenta.Text) / (1 + KEY_IGV), "###0.0000")
End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)
    Dim parametros() As String
Dim Criterio As String
If KeyAscii = 13 Then


             
                
                    parametros = Split(Trim(Me.TxtProducto.Text), " ")
                    Criterio = ""
                    For i = 0 To UBound(parametros)
                        If Criterio <> "" Then
                            Criterio = Trim(Criterio & "%" & Trim(parametros(i)))
                        Else
                            Criterio = Trim(parametros(i))
                        End If
                        
                    Next i
     strCadena = "SELECT * FROM view_producto WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "' AND nombre_prod LIKE '%" & Criterio & "%' ORDER BY nombre_prod LIMIT 0,75 "
     Call Me.llenarGrid_prod(Me.HfdDetalle, Me)
End If
End Sub

Private Sub UpDown4_Change()

End Sub



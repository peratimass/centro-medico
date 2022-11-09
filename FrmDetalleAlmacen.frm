VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmDetalleAlmacen 
   BorderStyle     =   0  'None
   Caption         =   "DETALLE ALMACEN"
   ClientHeight    =   8520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   15975
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPiePagina 
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
      Height          =   795
      Left            =   6120
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   56
      Top             =   6480
      Width           =   3975
   End
   Begin VB.Frame frmubigueo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "UBIGEO LOCAL"
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
      Height          =   1455
      Left            =   6720
      TabIndex        =   46
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtdistrito 
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
         Left            =   3520
         MaxLength       =   200
         TabIndex        =   53
         Top             =   240
         Width           =   735
      End
      Begin MSDataListLib.DataCombo DtcDistrito 
         Height          =   315
         Left            =   1440
         TabIndex        =   50
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
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
      Begin MSDataListLib.DataCombo DtcProvincia 
         Height          =   315
         Left            =   1440
         TabIndex        =   51
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
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
      Begin MSDataListLib.DataCombo DtcDepartamento 
         Height          =   315
         Left            =   1440
         TabIndex        =   52
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
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
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTAMENTO:"
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
         TabIndex        =   49
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROVINCIA :"
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
         Left            =   435
         TabIndex        =   48
         Top             =   675
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DISTRITO:"
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
         Left            =   615
         TabIndex        =   47
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame frmGrifo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "COMBUSTIBLE"
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
      Height          =   855
      Left            =   1680
      TabIndex        =   42
      Top             =   2400
      Visible         =   0   'False
      Width           =   8535
      Begin VB.TextBox txtcodigo_producto 
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
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   44
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblproducto 
         Alignment       =   2  'Center
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
         Left            =   2400
         TabIndex        =   45
         Top             =   360
         Width           =   6015
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
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
         Left            =   225
         TabIndex        =   43
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame frmcolor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ELIJE EL COLOR"
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
      Height          =   2415
      Left            =   13200
      TabIndex        =   35
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox txtbarra 
         Height          =   285
         Left            =   240
         TabIndex        =   41
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.OptionButton opt_color1 
         Height          =   315
         Index           =   4
         Left            =   240
         Picture         =   "FrmDetalleAlmacen.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         Tag             =   "line0005"
         Top             =   1800
         Width           =   2175
      End
      Begin VB.OptionButton opt_color1 
         Height          =   315
         Index           =   3
         Left            =   240
         Picture         =   "FrmDetalleAlmacen.frx":2572
         Style           =   1  'Graphical
         TabIndex        =   39
         Tag             =   "line0004"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton opt_color1 
         Height          =   315
         Index           =   2
         Left            =   240
         Picture         =   "FrmDetalleAlmacen.frx":48CB
         Style           =   1  'Graphical
         TabIndex        =   38
         Tag             =   "line0003"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton opt_color1 
         Height          =   315
         Index           =   1
         Left            =   240
         Picture         =   "FrmDetalleAlmacen.frx":6A0A
         Style           =   1  'Graphical
         TabIndex        =   37
         Tag             =   "line0002"
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton opt_color1 
         Height          =   315
         Index           =   0
         Left            =   240
         Picture         =   "FrmDetalleAlmacen.frx":892F
         Style           =   1  'Graphical
         TabIndex        =   36
         Tag             =   "line0001"
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CheckBox chk_color 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "COLOR DE BARRA"
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
      Left            =   11160
      TabIndex        =   34
      Top             =   960
      Width           =   2055
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   12840
      OleObjectBlob   =   "FrmDetalleAlmacen.frx":A636
      Top             =   480
   End
   Begin VB.CheckBox chk_skin 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "SKINS PERSONALIZABLES"
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
      Left            =   11160
      TabIndex        =   31
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CheckBox chk_conversion_moneda 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "CONVERSION DE MONEDA"
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
      Left            =   5880
      TabIndex        =   27
      Top             =   4560
      Width           =   2535
   End
   Begin VB.TextBox txtprefijo 
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
      Left            =   6870
      MaxLength       =   200
      TabIndex        =   25
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Frame frame_telefono 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TELEFONOS"
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
      Height          =   855
      Left            =   5880
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtTelefonos 
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
         Left            =   1080
         MaxLength       =   100
         TabIndex        =   24
         Top             =   320
         Width           =   3015
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TELEFONOS:"
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
         TabIndex        =   23
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.CheckBox chkcomprobantesPropios 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "COMPROBANTES PROPIOS VENTANILLA"
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
      Left            =   1680
      TabIndex        =   21
      Top             =   4480
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VitekeySoft.ChameleonBtn cmdsave 
      Height          =   855
      Left            =   12840
      TabIndex        =   19
      Top             =   6600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "PROCESAR"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetalleAlmacen.frx":A86A
      PICN            =   "FrmDetalleAlmacen.frx":A886
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chk_facturacion_centralizada 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "FACTURACION CENTRALIZADA"
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
      Left            =   1680
      TabIndex        =   18
      Top             =   5520
      Width           =   4095
   End
   Begin VB.CheckBox chk_comprobante_adicional 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "GENERACION COMPROBANTES ADICIONALES"
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
      Left            =   1680
      TabIndex        =   17
      Top             =   5160
      Width           =   4095
   End
   Begin MSDataListLib.DataCombo DtcSucursal 
      Height          =   315
      Left            =   1680
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "STOCK"
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
      Height          =   2175
      Left            =   1680
      TabIndex        =   13
      Top             =   6240
      Width           =   4095
      Begin VB.CheckBox chk_tiendaVirtual 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "TIENDA VIRTUAL"
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
         Height          =   300
         Left            =   120
         TabIndex        =   54
         Top             =   1680
         Width           =   3735
      End
      Begin VB.CheckBox chk_combustible 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "GRIFO [COMBUSTIBLE]"
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
         Height          =   300
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   3735
      End
      Begin VB.CheckBox chk_movimiento_sin_stock 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "PERMITIR MOVIMIENTOS SIN STOCK"
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
         Height          =   300
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   3735
      End
      Begin VB.CheckBox ChkAbilitado 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "HABILITADO"
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
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkstock 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "PERMITE TENER STOCK EN ESTA SUCURSAL"
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
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   3735
      End
   End
   Begin VB.CheckBox chk_facturacion_detallada 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "FACTURACION DETALLADA"
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
      Left            =   1680
      TabIndex        =   12
      Top             =   3360
      Width           =   4095
   End
   Begin VB.CheckBox chkVentanilla 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "VENTANILLA"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   3720
      Width           =   4095
   End
   Begin VB.CheckBox chkCajaIndependiente 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "CAJA INDEPENDIENTE"
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
      Left            =   1680
      TabIndex        =   10
      Top             =   4800
      Width           =   4095
   End
   Begin VB.TextBox txtCodCliente 
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
      Left            =   1725
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox chkDefault 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "DEFAULD"
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
      Left            =   1680
      TabIndex        =   8
      Top             =   5880
      Width           =   4095
   End
   Begin VB.TextBox txtEncargado 
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
      Left            =   3045
      MaxLength       =   50
      TabIndex        =   6
      Top             =   1200
      Width           =   3495
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
      Height          =   315
      Left            =   1725
      MaxLength       =   200
      TabIndex        =   4
      Top             =   840
      Width           =   4815
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1725
      MaxLength       =   80
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtAlmacen 
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
      Left            =   1725
      MaxLength       =   50
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
   Begin VitekeySoft.ChameleonBtn cmdexit 
      Height          =   855
      Left            =   14400
      TabIndex        =   20
      Top             =   6600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "SALIR"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetalleAlmacen.frx":DECE
      PICN            =   "FrmDetalleAlmacen.frx":DEEA
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
      Height          =   315
      Left            =   5880
      TabIndex        =   28
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Hfgskin 
      Height          =   4695
      Left            =   10440
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8281
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
   Begin VitekeySoft.ChameleonBtn cmdVisualizar 
      Height          =   375
      Left            =   14160
      TabIndex        =   33
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDetalleAlmacen.frx":E2DA
      PICN            =   "FrmDetalleAlmacen.frx":E2F6
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
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MENSAJE PIE DE PAGINA"
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
      Left            =   6120
      TabIndex        =   55
      Top             =   6240
      Width           =   1905
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PREFIJO :"
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
      Left            =   6000
      TabIndex        =   26
      Top             =   4260
      Width           =   705
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   14880
      Picture         =   "FrmDetalleAlmacen.frx":10821
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENCARGADO :"
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
      Left            =   345
      TabIndex        =   7
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DIRECCION :"
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
      Left            =   495
      TabIndex        =   5
      Top             =   900
      Width           =   945
   End
   Begin VB.Label LblRazonSocial 
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
      Left            =   705
      TabIndex        =   3
      Top             =   180
      Width           =   735
   End
   Begin VB.Label LblDireccion 
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
      Left            =   315
      TabIndex        =   2
      Top             =   540
      Width           =   1125
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   8520
      Left            =   0
      Top             =   0
      Width           =   15975
   End
End
Attribute VB_Name = "FrmDetalleAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strCodAlmacen As String
Public Procedencia As EnumProcede, defecto As String
Public in_skin_name As String

Private Sub chkCloud_Click()

End Sub

Private Sub Check1_Click()

End Sub

Private Sub chk_color_Click()

If Me.chk_color.Value = 1 Then
    Me.frmcolor.Visible = True
Else
    Me.frmcolor.Visible = False
End If

End Sub

Private Sub chk_skin_Click()

If Me.chk_skin.Value = 1 Then
    Call llenar_skin(Me.Hfgskin, in_skin_name)
    Me.Hfgskin.Visible = True
Else
    Me.Hfgskin.Visible = False
End If


End Sub

Private Sub chkDefault_Click()
If Me.chkDefault.Value = 1 Then
    defecto = "si"
Else
    defecto = "no"
End If
End Sub

Private Sub chkVentanilla_Click()
If Me.chkVentanilla.Value = 1 Then
    Me.DtcSucursal.Visible = True
    Me.chkcomprobantesPropios.Visible = True
Else
    Me.DtcSucursal.Visible = False
    Me.chkcomprobantesPropios.Visible = False
End If
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
On Error GoTo error
 Call Save

 Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub cmdVisualizar_Click()

Call load_skin(Me.Hfgskin.TextMatrix(Me.Hfgskin.Row, 1))

End Sub
Private Sub load_skin(ByVal in_skin As String)
Skin1.LoadSkin App.Path & "\Skins\" & in_skin & ".skn"
Skin1.ApplySkin Me.hwnd
End Sub



Private Sub DtcDistrito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcDistrito.BoundText <> "" Then
        strCadena = "SELECT id_provincia FROM distrito WHERE id_distrito='" & Me.DtcDistrito.BoundText & "'"
        Call ConfiguraTemporal(strCadena)
        If rstTemporal.RecordCount > 0 Then
            
            Me.DtcProvincia.Visible = True
            strCadena = "SELECT id_provincia as Codigo,descripcion as Descripcion FROM provincia WHERE id_provincia='" & rstTemporal("id_provincia") & "'"
            Call ConfiguraRst(strCadena)
            Call LlenaDataCombo(Me.DtcProvincia)
            Me.DtcProvincia.Enabled = False
        End If
    End If
End If
End Sub

Private Sub DtcProvincia_Change()
If Me.DtcProvincia.BoundText <> "" Then
    strCadena = "SELECT * FROM provincia WHERE id_provincia='" & Me.DtcProvincia.BoundText & "' "
    Call ConfiguraTemporal(strCadena)
    If rstTemporal.RecordCount > 0 Then
        
        Me.DtcDepartamento.Visible = True
        strCadena = "SELECT id_depa as Codigo,descripcion as Descripcion FROM departamentos WHERE id_depa='" & rstTemporal("id_departamento") & "'"
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcDepartamento)
        Set rst = Nothing
        Me.DtcDepartamento.Enabled = True
    End If
    Set rstTemporal = Nothing
End If
End Sub

Private Sub Form_Load()

CenterForm Me
Me.Top = 50

strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' and id_sucursal='0' ORDER BY id_alm"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcSucursal)
strCadena = "SELECT id_moneda as Codigo, descripcion as Descripcion FROM moneda"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)
Select Case FrmAlmacenes.Procedencia
    Case nuevo
        Me.txtCodigo.Enabled = False
         strCadena = "SELECT * FROM almacen WHERE ruc='" & Trim(KEY_RUC) & "'ORDER BY id_alm DESC"
         Call ConfiguraRst(strCadena)
         If rst.RecordCount > 0 Then
            rst.MoveFirst
            Me.txtCodigo.Text = formato_item(rst("id_alm") + 1, 5)
         Else
            Me.txtCodigo.Text = "00001"
         End If
         Me.DtcDistrito.BoundText = 0
         Me.DtcProvincia.BoundText = 0
         Me.DtcDepartamento.BoundText = 0
            
        Case modificar
            Call LLENA
    End Select
    
    If KEY_GRIFO = "si" Then
        Me.frmGrifo.Visible = True
    Else
        Me.frmGrifo.Visible = False
    End If

End Sub


Private Sub put_gen_sucursal(ByVal in_alm As String, ByVal in_profijo As String)
Dim in_principal As String
If Val(in_alm) = 1 Then
   in_principal = 1
Else
   in_principal = 0
End If
strCadena = "SELECT * FROM gen_sucursal WHERE IdEmpresaSis='" & KEY_RUC & "' and IdEmpresa='" & in_alm & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   strCadena = "UPDATE gen_sucursal SET Prefijo='" & in_profijo & "',Nombre='" & Trim(Me.txtAlmacen.Text) & "',Direccion='" & Trim(Me.TxtDireccion.Text) & "' ,IndPrincipal='" & in_principal & "'WHERE IdEmpresaSis='" & KEY_RUC & "' and IdEmpresa='" & in_alm & "'"
   CnBd.Execute (strCadena)
Else
   strCadena = "SELECT * FROM gen_sucursal ORDER BY id DESC LIMIT 1"
   Call ConfiguraRstK(strCadena)
   If rstK.RecordCount > 0 Then
       in_codigo = "1CIX" + Format(Val(Mid(rstK("id"), 5, 3)) + 1, "000")
   End If
   strCadena = "INSERT INTO gen_sucursal(`Id`,`Prefijo`,`IdEmpresaSis`,`IdEmpresa`,`Nombre`,`Abreviatura`,`Direccion`,`IndPrincipal`,`UsuarioCrea`,`FechaCrea`,`Activo`)VALUES " & _
   "('" & in_codigo & "','" & in_profijo & "','" & KEY_RUC & "','" & in_alm & "','" & Trim(Me.txtAlmacen.Text) & "','" & Trim(Me.txtAlmacen.Text) & "','" & Trim(Me.TxtDireccion.Text) & "','" & in_principal & "','" & KEY_USUARIO & "',CURDATE(),'1')"
   CnBd.Execute (strCadena)
End If

End Sub
Public Sub llenar_skin(ByVal Grilla As MSHFlexGrid, ByRef in_skin As String)

strCadena = "SELECT * FROM skin  order by descripcion"
Call ConfiguraRstT(strCadena)
'Call Cargar_FlexGrid(Me.HfActividades, 8, rstT)

If rstT.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstT.Fields.Count)
        For Each Campo In rstT.Fields
           Grilla.ColWidth(0) = 400
           Grilla.ColWidth(1) = 2100
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 300
           
        Next
        cabecera = "SKIN" & vbTab & "DESCRIPCION"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstT.MoveFirst
        For i = 0 To rstT.RecordCount - 1
        
        
        estado = Chr(168)
        Fila = Format(i, "00") & vbTab & rstT("descripcion") & vbTab & rstT("nombre") & vbTab & estado
          Grilla.AddItem Fila
           If c = NumeroCampo Then
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 3 '  .. en la columna
                            ' cambia la fuente para esta celda
                            
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
                            If rstT("descripcion") = in_skin Then
                                estado = Chr(254)
                                For k = 0 To 3
                                    Grilla.col = k
                                    Grilla.Row = i
                                    Grilla.CellBackColor = &HDFDFE0
                                 Next k
                            Else
                                estado = Chr(168)
                            End If
                            
                            
                            
                            
                        End With
                          
        End If
          Fila = ""
        
          rstT.MoveNext
      Next i
   
    
     
End Sub

Private Sub Save()
Dim idalmacen As Integer, StrAlmacen As String, strdetallada As String
Dim in_prefijo As String
  If Me.txtAlmacen.Text = "" Then
    MsgBox MSGFALTADATOS, vbCritical, MSGVALIDACION
Else
        If Me.chkVentanilla.Value = 1 Then
            ntipoentidad = "00012"
            nsucursal = Me.DtcSucursal.BoundText
        Else
            ntipoentidad = "0"
            nsucursal = "0"
        End If
        
        If Me.ChkAbilitado.Value = 1 Then
            activo = "si"
        Else
            activo = "no"
        End If
        
        If Me.chk_facturacion_detallada.Value = 1 Then
            strdetallada = "si"
        Else
            strdetallada = "no"
        End If
        
        If Me.chkstock.Value = 1 Then
            stock = "si"
        Else
            stock = "no"
        End If
        
        If Me.chk_comprobante_adicional.Value = 1 Then
            in_comprobante_adicional = "si"
        Else
            in_comprobante_adicional = "no"
        End If
        
        If Me.chk_facturacion_centralizada.Value = 1 Then
            in_centralizada = "si"
        Else
            in_centralizada = "no"
        End If
        
        If Me.chkcomprobantesPropios.Visible = True Then
           If Me.chkcomprobantesPropios.Value = 1 Then
               in_comprobantes_propios = "si"
           Else
               in_comprobantes_propios = "no"
           End If
           
        Else
            in_comprobantes_propios = "no"
        End If
        
        If Me.chk_conversion_moneda.Value = 1 Then
           in_conversion_dolares = "si"
        Else
           in_conversion_dolares = "no"
        End If
         
        If Me.chk_movimiento_sin_stock.Value = 1 Then
            KEY_MOVIMIENTO_SIN_STOCK = "si"
        Else
            KEY_MOVIMIENTO_SIN_STOCK = "no"
        End If
         
        
        
        If Me.chk_skin.Value = 1 Then
            in_skin = "si"
            skin_name = Me.Hfgskin.TextMatrix(Me.Hfgskin.Row, 1)
        Else
            in_skin = "no"
            skin_name = ""
        End If
        
        
        If Me.chk_color.Value = 1 Then
            in_color = "si"
            in_color_name = Trim(Me.txtbarra.Text)
        Else
            in_color = "no"
            in_color_name = ""
        End If
        
        
        If Me.chk_tiendaVirtual.Value = 1 Then
            in_tienda_virtual = "si"
        Else
            in_tienda_virtual = "no"
        End If
        
    
    
    If in_tienda_virtual = "si" Then
        strCadena = "SELECT  descripcion FROM almacen WHERE tienda_virtual='si' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            MsgBox "Desmarque Tienda Online...." + Chr(13) + "Ya existe una sucursal como tienda Online", vbInformation
            in_tienda_virtual = "no"
        End If
    End If
    
    
    
    
    Select Case FrmAlmacenes.Procedencia
      Case nuevo
        Dim rstAP As New ADODB.Recordset
        StrAlmacen = formato_item(Val(Me.txtCodigo.Text), 5)
       
        
        If Me.chkCajaIndependiente.Value = 1 Then
            str_cajai = "si"
        Else
            str_cajai = "no"
        End If
        If Me.chkDefault.Value = 1 Then
            defecto = "si"
        Else
            defecto = "no"
        End If
        
        
        
        
        
        
        in_prefijo = Mid(KEY_RUC, 9, 2) & Format(Trim(Val(Me.txtCodigo.Text)), "00")
        strCadena = "INSERT INTO almacen (id_alm,conversion_dolares,prefijo,descripcion,direccion,id_responsable,caja_independiente,id_sucursal,facturacion_detallada,facturacion_centralizada,activo,stock,defecto,id_tipoentidad,comprobante_adicional,comprobantes_propios,telefonos,id_moneda,movimiento_sin_stock,usa_skin,skin,color_barra,color,id_departamento,id_provincia,id_distrito,tienda_virtual,ruc) VALUES " & _
        " ('" & StrAlmacen & "','" & in_conversion_dolares & "','" & in_prefijo & "','" & Me.txtAlmacen.Text & "','" & Me.TxtDireccion.Text & "','" & Trim(Me.txtCodCliente.Text) & "','" & str_cajai & "','" & nsucursal & "','" & strdetallada & "','" & in_centralizada & "','" & activo & "','" & stock & "','" & defecto & "','" & ntipoentidad & "','" & in_comprobante_adicional & "','" & in_comprobantes_propios & "','" & Trim(Me.txtTelefonos.Text) & "','" & Me.DtcMoneda.BoundText & "','" & KEY_MOVIMIENTO_SIN_STOCK & "','" & in_skin & "','" & skin_name & "','" & in_color & "','" & in_color_name & "','" & Me.DtcDepartamento.BoundText & "','" & Me.DtcProvincia.BoundText & "','" & Me.DtcDistrito.BoundText & "','" & in_tienda_virtual & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        If defecto = "si" Then
            strCadena = "UPDATE almacen SET defecto='no' WHERE id_alm <> '" & Me.txtCodigo.Text & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
        End If
        
    If ntipoentidad = "0" And stock = "si" Then
        Call put_insert_producto_alm(StrAlmacen)
    End If
    
    FrmAlmacenes.Actualizar_Alm
    Unload Me
    Exit Sub
    
    Case modificar
        in_prefijo = Mid(KEY_RUC, 9, 2) & Format(Trim(Val(Me.txtCodigo.Text)), "00")
        If (Me.chkDefault.Value = 1) Then
            defecto = "si"
        Else
            defecto = "no"
        End If
        If Me.chkCajaIndependiente.Value = 1 Then
            str_cajai = "si"
        Else
            str_cajai = "no"
        End If
        
        strCadena = "UPDATE almacen SET pie_pagina='" & Trim(Me.txtPiePagina.Text) & "',tienda_virtual='" & in_tienda_virtual & "',id_departamento='" & Me.DtcDepartamento.BoundText & "',id_provincia='" & Me.DtcProvincia.BoundText & "',id_distrito='" & Me.DtcDistrito.BoundText & "',color_barra='" & in_color & "',color='" & in_color_name & "',usa_skin='" & in_skin & "',skin='" & skin_name & "', movimiento_sin_stock='" & KEY_MOVIMIENTO_SIN_STOCK & "', id_moneda='" & Me.DtcMoneda.BoundText & "'," & _
        " conversion_dolares='" & in_conversion_dolares & "',prefijo='" & in_prefijo & "',telefonos='" & Trim(Me.txtTelefonos.Text) & "',comprobantes_propios='" & in_comprobantes_propios & "',comprobante_adicional='" & in_comprobante_adicional & "',defecto='" & defecto & "',stock='" & stock & "',activo='" & activo & "', facturacion_detallada='" & strdetallada & "',facturacion_centralizada='" & in_centralizada & "', caja_independiente='" & str_cajai & "',descripcion='" & Me.txtAlmacen.Text & "', direccion='" & Me.TxtDireccion.Text & "',id_responsable='" & Trim(Me.txtCodCliente.Text) & "',id_tipoentidad='" & ntipoentidad & "',id_sucursal='" & nsucursal & "' WHERE id_alm = '" & Trim(Me.txtCodigo.Text) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
        
        If stock = "si" And nsucursal = "0" Then
           strCadena = "SELECT * FROM almacen_producto WHERE id_alm='" & Trim(Me.txtCodigo.Text) & "' and ruc='" & KEY_RUC & "'"
           Call ConfiguraRst(strCadena)
           If rst.RecordCount < 1 Then
              Call put_insert_producto_alm(Trim(Me.txtCodigo.Text))
           End If
        End If
        
        If defecto = "si" Then
            strCadena = "UPDATE almacen SET defecto='no' WHERE id_alm <> '" & Trim(Me.txtCodigo.Text) & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
        End If
        
        Call put_gen_sucursal(Trim(Me.txtCodigo.Text), in_prefijo)
        Call FrmAlmacenes.Actualizar_Alm
        
        
        
        Unload Me
        End Select
  End If
End Sub
Private Sub LLENA()

strCadena = "SELECT * FROM almacen WHERE id_alm='" & FrmAlmacenes.HfgAlmacen.TextMatrix(FrmAlmacenes.HfgAlmacen.Row, 0) & "'AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
Me.txtCodigo.Text = rst(0)
Me.DtcDistrito.BoundText = 0
Me.DtcProvincia.BoundText = 0
Me.DtcDepartamento.BoundText = 0
Me.txtAlmacen.Text = UCase(rst("descripcion"))
Me.TxtDireccion.Text = UCase(rst("direccion"))
Me.txtCodCliente.Text = rst("id_responsable")
Me.txtEncargado.Text = NombrePersona(rst("id_responsable"))
Me.txtPiePagina.Text = rst("pie_pagina")
    
    If rst("tienda_virtual") = "si" Then
        Me.chk_tiendaVirtual.Value = 1
    Else
        Me.chk_tiendaVirtual.Value = 0
    End If
    
    If (rst("defecto") = "si") Then
        Me.chkDefault.Value = 1
    Else
        Me.chkDefault.Value = 0
    End If
    Me.DtcMoneda.Visible = True
    If rst("facturacion_centralizada") = "si" Then
       Me.chk_facturacion_centralizada.Value = 1
    Else
       Me.chk_facturacion_centralizada.Value = 0
    End If
    
    If rst("conversion_dolares") = "si" Then
       Me.chk_conversion_moneda.Value = 1
    Else
        Me.chk_conversion_moneda.Value = 0
    End If
   Me.DtcMoneda.BoundText = rst("id_moneda")
   
    If rst("movimiento_sin_stock") = "si" Then
       Me.chk_movimiento_sin_stock.Value = 1
    Else
        Me.chk_movimiento_sin_stock.Value = 0
    End If
    
    If rst("skin") = "si" Then
        Me.chk_skin.Value = 1
        in_skin_name = rst("skin_name")
    Else
        Me.chk_skin.Value = 0
        in_skin_name = "-"
    End If
    
    If rst("color_barra") = "si" Then
       Me.chk_color.Value = 1
       Me.frmcolor.Visible = True
       
    Else
        Me.chk_color.Value = 0
    End If
    
    
    
    
    
    
    
    
    If rst("stock") = "si" Then
       Me.chkstock.Value = 1
    Else
       Me.chkstock.Value = 0
    End If
    If rst("activo") = "si" Then
       Me.ChkAbilitado.Value = 1
    Else
       Me.ChkAbilitado.Value = 0
    End If
    Me.txtTelefonos.Text = rst("telefonos")
    
    If rst("id_sucursal") <> "0" Then
        Me.chkVentanilla.Value = 1
        Me.DtcSucursal.Visible = True
        Me.DtcSucursal.BoundText = rst("id_sucursal")
        Me.frame_telefono.Visible = False
    Else
        Me.frame_telefono.Visible = True
        Me.chkVentanilla.Value = 0
        Me.DtcSucursal.Visible = False
    End If
    
    If rst("facturacion_detallada") = "si" Then
        Me.chk_facturacion_detallada.Value = 1
    Else
        Me.chk_facturacion_detallada.Value = 0
    End If
   
 If rst("caja_independiente") = "si" Then
    Me.chkCajaIndependiente.Value = 1
 Else
    Me.chkCajaIndependiente.Value = 0
 End If
 
 If rst("comprobante_adicional") = "si" Then
    Me.chk_comprobante_adicional.Value = 1
 Else
    Me.chk_comprobante_adicional.Value = 0
 End If
 
 If rst("comprobantes_propios") = "si" Then
    Me.chkcomprobantesPropios.Value = 1
 Else
    Me.chkcomprobantesPropios.Value = 0
 End If
 
 
 If KEY_GRIFO = "si" Then
    Me.txtcodigo_producto.Text = rst("id_producto")
    Me.lblproducto.Caption = get_producto(Me.txtcodigo_producto.Text)
 End If
    
    cDepartamento = rst("id_departamento")
    cProvincia = rst("id_provincia")
    cDistrito = rst("id_distrito")
    
 If Val(cDepartamento) > 0 Then
    strCadena = "SELECT id_depa as Codigo,descripcion as Descripcion FROM departamentos WHERE id_depa='" & cDepartamento & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcDepartamento)
    Me.DtcDepartamento.BoundText = cDepartamento
End If

If Val(cProvincia) > 0 Then
    strCadena = "SELECT id_provincia as Codigo,descripcion as Descripcion FROM provincia WHERE  id_departamento='" & cDepartamento & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcProvincia)
    Me.DtcProvincia.BoundText = cProvincia
End If
If Val(cDistrito) > 0 Then
    strCadena = "SELECT id_distrito as Codigo,descripcion as Descripcion FROM distrito WHERE  id_provincia='" & cProvincia & "'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcDistrito)
    Me.DtcDistrito.BoundText = cDistrito
End If

 
 
 
End If
End Sub



Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)

  Select Case Button.key
    Case KEY_SAVE
     
    Case KEY_CANCEL
      Unload Me
  End Select
 
End Sub

Private Sub opt_color1_Click(Index As Integer)
Me.txtbarra.Text = Me.opt_color1(Index).Tag
End Sub

Private Sub txtAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtDireccion)
End If
End Sub

Private Sub TxtCodCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Trim(Me.txtCodCliente.Text) <> "" Then
    strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtCodCliente.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.txtCodCliente.Text = rst("dni")
        Me.txtEncargado.Text = rst("nombre_completo")
    Else
        Procedencia = Selecionar
        FrmPersona.Show
        
    End If
Else
     Procedencia = Selecionar
     FrmPersona.Show
End If
End If
End Sub

Private Sub txtcodigo_producto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
End If
End Sub

Private Sub TxtDireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtCodCliente)
End If
End Sub

Private Sub TxtDistrito_Change()
If Trim(Me.txtdistrito.Text) <> "" Then
    
    strCadena = "SELECT id_distrito as Codigo,CONCAT(d.descripcion,' - ',p.descripcion) as Descripcion FROM distrito d,provincia p  WHERE  d.id_provincia=p.id_provincia and  d.descripcion LIKE '%" & Trim(Me.txtdistrito.Text) & "%'"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcDistrito)
    Set rst = Nothing
    
End If
End Sub

Private Sub TxtDistrito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcDistrito.BoundText <> "" Then
        Me.DtcDistrito.SetFocus
    End If
End If
End Sub

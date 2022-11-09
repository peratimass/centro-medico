VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHotel 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmcomandas 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   2520
      TabIndex        =   247
      Top             =   2040
      Visible         =   0   'False
      Width           =   14055
      Begin MSComCtl2.DTPicker DtpInicio_comanda 
         Height          =   300
         Left            =   8280
         TabIndex        =   255
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   187891713
         CurrentDate     =   43718
      End
      Begin VB.CheckBox chk_rango_fecha 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "FECHAS:"
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
         Height          =   300
         Left            =   7320
         TabIndex        =   254
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtCliente_busqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   253
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtHabitacion_busqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   251
         Top             =   480
         Width           =   1575
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfComandasListado 
         Height          =   4695
         Left            =   120
         TabIndex        =   248
         Top             =   1080
         Width           =   13815
         _ExtentX        =   24368
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
      Begin VitekeySoft.ChameleonBtn cmdImprimirComanda 
         Height          =   450
         Left            =   12480
         TabIndex        =   249
         Top             =   5880
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   794
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
         MICON           =   "frmHotel.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DtpFin_comanda 
         Height          =   300
         Left            =   9720
         TabIndex        =   256
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   187891713
         CurrentDate     =   43718
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE :"
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
         Left            =   3360
         TabIndex        =   252
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HABITACION:"
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
         TabIndex        =   250
         Top             =   480
         Width           =   885
      End
      Begin VB.Shape Shape11 
         BorderWidth     =   2
         Height          =   6495
         Left            =   0
         Top             =   0
         Width           =   14055
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   13560
         Picture         =   "frmHotel.frx":001C
         Top             =   240
         Width           =   240
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdComandas 
      Height          =   555
      Left            =   240
      TabIndex        =   246
      Top             =   8160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   "COMANDAS"
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
      MICON           =   "frmHotel.frx":2EC0
      PICN            =   "frmHotel.frx":2EDC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frmreserva 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6615
      Left            =   2520
      TabIndex        =   127
      Top             =   2040
      Visible         =   0   'False
      Width           =   14055
      Begin TabDlg.SSTab SSTab1 
         Height          =   6285
         Left            =   240
         TabIndex        =   128
         Top             =   120
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   11086
         _Version        =   393216
         Tabs            =   5
         Tab             =   4
         TabsPerRow      =   5
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
         TabCaption(0)   =   "DATOS DE RESERVA"
         TabPicture(0)   =   "frmHotel.frx":5407
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Shape4"
         Tab(0).Control(1)=   "lblPersona(20)"
         Tab(0).Control(2)=   "lblPersona(21)"
         Tab(0).Control(3)=   "lblPersona(22)"
         Tab(0).Control(4)=   "lblPersona(23)"
         Tab(0).Control(5)=   "lblPersona(24)"
         Tab(0).Control(6)=   "lblPersona(25)"
         Tab(0).Control(7)=   "lblfecha_ingreso"
         Tab(0).Control(8)=   "lblHoraIngreso"
         Tab(0).Control(9)=   "lblOperador"
         Tab(0).Control(10)=   "lblEstadoReserva"
         Tab(0).Control(11)=   "lblhabitacionnumero"
         Tab(0).Control(12)=   "Shape9"
         Tab(0).Control(13)=   "lblPersona(26)"
         Tab(0).Control(14)=   "cmdagregarComprobante"
         Tab(0).Control(15)=   "HfHistorialFactura"
         Tab(0).Control(16)=   "cmdEliminarVisita"
         Tab(0).Control(17)=   "cmdAgregarVisita"
         Tab(0).Control(18)=   "HfVisita"
         Tab(0).Control(19)=   "txtDni"
         Tab(0).Control(20)=   "txtCliente"
         Tab(0).Control(21)=   "txtDireccion"
         Tab(0).Control(22)=   "cmdgenerarReserva"
         Tab(0).Control(23)=   "Frame2"
         Tab(0).Control(24)=   "txtid_habitacion"
         Tab(0).Control(25)=   "chk_extranjero"
         Tab(0).Control(26)=   "txtid_producto"
         Tab(0).Control(27)=   "txtidReserva"
         Tab(0).Control(28)=   "Check1"
         Tab(0).Control(29)=   "txtruc"
         Tab(0).Control(30)=   "frmvisita"
         Tab(0).Control(31)=   "frmTraslado"
         Tab(0).Control(32)=   "chk_adicionar"
         Tab(0).Control(33)=   "txtSerie"
         Tab(0).Control(34)=   "TxtNumero"
         Tab(0).Control(35)=   "DtcTipoComprobante"
         Tab(0).Control(36)=   "txtidVenta"
         Tab(0).Control(37)=   "frmunificada"
         Tab(0).ControlCount=   38
         TabCaption(1)   =   "DIAS ALOJAMIENTO"
         TabPicture(1)   =   "frmHotel.frx":5423
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Shape6"
         Tab(1).Control(1)=   "Label1"
         Tab(1).Control(2)=   "Label2"
         Tab(1).Control(3)=   "cmdFacturarAlojamiento"
         Tab(1).Control(4)=   "cmdsalir_alojamiento"
         Tab(1).Control(5)=   "cmdEliminar_alojamiento"
         Tab(1).Control(6)=   "cmdAgregarAlojamiento"
         Tab(1).Control(7)=   "HfAlojamiento"
         Tab(1).Control(8)=   "DtpFechaAlojamiento"
         Tab(1).Control(9)=   "txtPrecio_Alojamiento"
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "CONSUMO HABITACION"
         TabPicture(2)   =   "frmHotel.frx":543F
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtid_alm"
         Tab(2).Control(1)=   "txtPrecio_habitacion"
         Tab(2).Control(2)=   "txtcantidad_habitacion"
         Tab(2).Control(3)=   "txtid_producto_habitacion"
         Tab(2).Control(4)=   "HfconsumoHabitacion"
         Tab(2).Control(5)=   "cmdEliminar_consumo_habitacion"
         Tab(2).Control(6)=   "cmdsalir_consumo_habitacion"
         Tab(2).Control(7)=   "cmdagregar_consumo_habitacion"
         Tab(2).Control(8)=   "cmdfacturar_consumo"
         Tab(2).Control(9)=   "lblfecha(25)"
         Tab(2).Control(10)=   "lblfecha(24)"
         Tab(2).Control(11)=   "lblproducto_habitacion(0)"
         Tab(2).Control(12)=   "lblfecha(20)"
         Tab(2).Control(13)=   "Shape7"
         Tab(2).ControlCount=   14
         TabCaption(3)   =   "RESTAURANT"
         TabPicture(3)   =   "frmHotel.frx":545B
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtprecio_rest"
         Tab(3).Control(1)=   "txtcantidad_resta"
         Tab(3).Control(2)=   "txtidproducto_resta"
         Tab(3).Control(3)=   "HfRestaurant"
         Tab(3).Control(4)=   "cmd_eliminar_consumo_restaurant"
         Tab(3).Control(5)=   "cmdeliminar_consumo_restaurant"
         Tab(3).Control(6)=   "cmdagregar_resta"
         Tab(3).Control(7)=   "cmdFacturar_rest"
         Tab(3).Control(8)=   "cmdImprimir"
         Tab(3).Control(9)=   "HfComandas"
         Tab(3).Control(10)=   "cmdreimprimir"
         Tab(3).Control(11)=   "lblfecha(23)"
         Tab(3).Control(12)=   "lblfecha(21)"
         Tab(3).Control(13)=   "lblproducto_resta(0)"
         Tab(3).Control(14)=   "lblfecha(22)"
         Tab(3).Control(15)=   "Shape8"
         Tab(3).ControlCount=   16
         TabCaption(4)   =   "HISTORIAL"
         TabPicture(4)   =   "frmHotel.frx":5477
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "Shape10"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "lblPersona(30)"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "cmdReporte"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "cmdConsultar"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "DtpFechaFin"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).Control(5)=   "HfHistorial"
         Tab(4).Control(5).Enabled=   0   'False
         Tab(4).Control(6)=   "DtpFechaIni"
         Tab(4).Control(6).Enabled=   0   'False
         Tab(4).ControlCount=   7
         Begin MSComCtl2.DTPicker DtpFechaIni 
            Height          =   375
            Left            =   1080
            TabIndex        =   242
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   187891713
            CurrentDate     =   43718
         End
         Begin VB.Frame frmunificada 
            BackColor       =   &H00FFFFFF&
            Caption         =   "QUE HABITACIONES DESEA CANCELAR"
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
            Height          =   5415
            Left            =   -68880
            TabIndex        =   234
            Top             =   660
            Visible         =   0   'False
            Width           =   3855
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfUnificada 
               Height          =   3735
               Left            =   120
               TabIndex        =   235
               Top             =   480
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   6588
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
            Begin VitekeySoft.ChameleonBtn cmdFacturaGlobal 
               Height          =   495
               Left            =   120
               TabIndex        =   236
               Top             =   4320
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   873
               BTYPE           =   5
               TX              =   "GENERAR FACTURA"
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
               MICON           =   "frmHotel.frx":5493
               PICN            =   "frmHotel.frx":54AF
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
               Left            =   3540
               Picture         =   "frmHotel.frx":7A94
               Top             =   120
               Width           =   240
            End
         End
         Begin VB.TextBox txtidVenta 
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
            Left            =   -74520
            TabIndex        =   232
            Top             =   5460
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSDataListLib.DataCombo DtcTipoComprobante 
            Height          =   315
            Left            =   -72520
            TabIndex        =   230
            Top             =   5540
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
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
         Begin VB.TextBox TxtNumero 
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
            Left            =   -70360
            TabIndex        =   229
            Top             =   5540
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtSerie 
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
            Left            =   -70920
            TabIndex        =   228
            Top             =   5540
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CheckBox chk_adicionar 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "ADD"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   260
            Left            =   -73200
            TabIndex        =   227
            Top             =   5580
            Width           =   615
         End
         Begin VB.TextBox txtid_alm 
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
            Left            =   -64320
            TabIndex        =   225
            Top             =   3540
            Width           =   855
         End
         Begin VB.Frame frmTraslado 
            BackColor       =   &H00FFFFFF&
            Caption         =   "HABITACION DESTINO TRASLADO"
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
            Height          =   5415
            Left            =   -68880
            TabIndex        =   222
            Top             =   660
            Visible         =   0   'False
            Width           =   3855
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfTraslado 
               Height          =   3735
               Left            =   120
               TabIndex        =   223
               Top             =   480
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   6588
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
            Begin VitekeySoft.ChameleonBtn cmdProcesarTranslado 
               Height          =   495
               Left            =   120
               TabIndex        =   224
               Top             =   4320
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   873
               BTYPE           =   5
               TX              =   "PROCESAR TRANSLADO"
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
               MICON           =   "frmHotel.frx":A938
               PICN            =   "frmHotel.frx":A954
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Image img_cerrar 
               Height          =   240
               Left            =   3540
               Picture         =   "frmHotel.frx":CF39
               Top             =   120
               Width           =   240
            End
         End
         Begin VB.Frame frmvisita 
            BackColor       =   &H00FFFFFF&
            Caption         =   "DETALLE VISITA"
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
            Height          =   3255
            Left            =   -68760
            TabIndex        =   214
            Top             =   1260
            Visible         =   0   'False
            Width           =   3615
            Begin VB.TextBox txtcelular_visita 
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
               Left            =   840
               TabIndex        =   220
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox txtdni_visita 
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
               Left            =   840
               TabIndex        =   216
               Top             =   360
               Width           =   1695
            End
            Begin VB.TextBox txtnombre_visita 
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
               Left            =   840
               TabIndex        =   215
               Top             =   840
               Width           =   2655
            End
            Begin VitekeySoft.ChameleonBtn cmdGuardarVisita 
               Height          =   375
               Left            =   840
               TabIndex        =   221
               Top             =   1800
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               BTYPE           =   5
               TX              =   "GUARDAR VISITA"
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
               MICON           =   "frmHotel.frx":FDDD
               PICN            =   "frmHotel.frx":FDF9
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label lblPersona 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NRO CELL:"
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
               Index           =   29
               Left            =   120
               TabIndex        =   219
               Top             =   1320
               Width           =   690
            End
            Begin VB.Label lblPersona 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DNI :"
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
               Index           =   28
               Left            =   285
               TabIndex        =   218
               Top             =   480
               Width           =   330
            End
            Begin VB.Label lblPersona 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DATOS:"
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
               Index           =   27
               Left            =   120
               TabIndex        =   217
               Top             =   900
               Width           =   495
            End
         End
         Begin VB.TextBox txtPrecio_habitacion 
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
            Left            =   -67680
            TabIndex        =   204
            Top             =   4700
            Width           =   855
         End
         Begin VB.TextBox txtPrecio_Alojamiento 
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
            Left            =   -68640
            TabIndex        =   203
            Top             =   4380
            Width           =   1215
         End
         Begin VB.TextBox txtprecio_rest 
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
            Left            =   -66600
            TabIndex        =   201
            Top             =   4095
            Width           =   735
         End
         Begin VB.TextBox txtruc 
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
            Left            =   -70440
            TabIndex        =   200
            Top             =   1260
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
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
            ForeColor       =   &H00FF0000&
            Height          =   260
            Left            =   -71640
            TabIndex        =   199
            Top             =   1260
            Width           =   1095
         End
         Begin VB.TextBox txtidReserva 
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
            Left            =   -70080
            TabIndex        =   194
            Top             =   2460
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtcantidad_resta 
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
            Left            =   -67680
            TabIndex        =   193
            Top             =   4095
            Width           =   735
         End
         Begin VB.TextBox txtidproducto_resta 
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
            Left            =   -73680
            TabIndex        =   190
            Top             =   4095
            Width           =   975
         End
         Begin VB.TextBox txtcantidad_habitacion 
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
            Left            =   -68640
            TabIndex        =   185
            Top             =   4700
            Width           =   855
         End
         Begin VB.TextBox txtid_producto_habitacion 
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
            Left            =   -73680
            TabIndex        =   182
            Top             =   4700
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker DtpFechaAlojamiento 
            Height          =   315
            Left            =   -73800
            TabIndex        =   174
            Top             =   4380
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   187891713
            CurrentDate     =   43602
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
            Height          =   315
            Left            =   -71160
            TabIndex        =   151
            Top             =   2820
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CheckBox chk_extranjero 
            Appearance      =   0  'Flat
            BackColor       =   &H000080FF&
            Caption         =   "EXTRANJERO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   260
            Left            =   -71640
            TabIndex        =   147
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txtid_habitacion 
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
            Left            =   -71160
            TabIndex        =   143
            Top             =   2460
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "DATOS ADICIONALES"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5295
            Left            =   -64920
            TabIndex        =   136
            Top             =   780
            Width           =   2415
            Begin VitekeySoft.ChameleonBtn cmdLimpieza 
               Height          =   495
               Left            =   120
               TabIndex        =   144
               Top             =   960
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   873
               BTYPE           =   5
               TX              =   "LIMPIEZA             "
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Cambria"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   2
               FOCUSR          =   -1  'True
               BCOL            =   12632064
               BCOLO           =   12632064
               FCOL            =   8388608
               FCOLO           =   8388608
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmHotel.frx":123DE
               PICN            =   "frmHotel.frx":123FA
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VitekeySoft.ChameleonBtn cmdDiponible 
               Height          =   495
               Left            =   120
               TabIndex        =   145
               Top             =   360
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   873
               BTYPE           =   5
               TX              =   "DISPONIBLE        "
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Cambria"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   2
               FOCUSR          =   -1  'True
               BCOL            =   8454016
               BCOLO           =   8454016
               FCOL            =   8388608
               FCOLO           =   8388608
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmHotel.frx":149DF
               PICN            =   "frmHotel.frx":149FB
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VitekeySoft.ChameleonBtn cmdMantenimiento 
               Height          =   495
               Left            =   120
               TabIndex        =   146
               Top             =   1560
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   873
               BTYPE           =   5
               TX              =   "MANTENIMIENTO"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Cambria"
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
               MICON           =   "frmHotel.frx":16FE0
               PICN            =   "frmHotel.frx":16FFC
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VitekeySoft.ChameleonBtn cmdCerrar 
               Height          =   495
               Left            =   120
               TabIndex        =   149
               Top             =   4680
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   873
               BTYPE           =   5
               TX              =   "CERRAR PANTALLA"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Cambria"
                  Size            =   9
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
               MICON           =   "frmHotel.frx":195E1
               PICN            =   "frmHotel.frx":195FD
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VitekeySoft.ChameleonBtn cmdFactura_unificada 
               Height          =   495
               Left            =   120
               TabIndex        =   195
               Top             =   2760
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   873
               BTYPE           =   5
               TX              =   "GENERAR FACTURA"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Cambria"
                  Size            =   9
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
               MICON           =   "frmHotel.frx":19917
               PICN            =   "frmHotel.frx":19933
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VitekeySoft.ChameleonBtn cmdTrasladoInterno 
               Height          =   495
               Left            =   120
               TabIndex        =   209
               Top             =   2160
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   873
               BTYPE           =   5
               TX              =   "TRASLADO              "
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Cambria"
                  Size            =   9
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
               MICON           =   "frmHotel.frx":1BF18
               PICN            =   "frmHotel.frx":1BF34
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VitekeySoft.ChameleonBtn cmdUnificada 
               Height          =   495
               Left            =   120
               TabIndex        =   233
               Top             =   4080
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   873
               BTYPE           =   5
               TX              =   "UNIFICADA"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Cambria"
                  Size            =   9
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
               MICON           =   "frmHotel.frx":1FC75
               PICN            =   "frmHotel.frx":1FC91
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
         Begin VitekeySoft.ChameleonBtn cmdgenerarReserva 
            Height          =   495
            Left            =   -73200
            TabIndex        =   135
            Top             =   3540
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   873
            BTYPE           =   5
            TX              =   "GENERAR RESERVA"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Cambria"
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
            MICON           =   "frmHotel.frx":22276
            PICN            =   "frmHotel.frx":22292
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txtDireccion 
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
            Left            =   -73200
            TabIndex        =   134
            Top             =   2040
            Width           =   3015
         End
         Begin VB.TextBox txtCliente 
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
            Left            =   -73200
            TabIndex        =   132
            Top             =   1680
            Width           =   3015
         End
         Begin VB.TextBox txtDni 
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
            Left            =   -73200
            TabIndex        =   130
            Top             =   1320
            Width           =   1455
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfAlojamiento 
            Height          =   3135
            Left            =   -74640
            TabIndex        =   172
            Top             =   900
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   5530
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
         Begin VitekeySoft.ChameleonBtn cmdAgregarAlojamiento 
            Height          =   495
            Left            =   -67080
            TabIndex        =   175
            Top             =   4260
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            BTYPE           =   5
            TX              =   "AGREGAR"
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
            MICON           =   "frmHotel.frx":24877
            PICN            =   "frmHotel.frx":24893
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdEliminar_alojamiento 
            Height          =   915
            Left            =   -65280
            TabIndex        =   176
            Top             =   900
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   1614
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
            MICON           =   "frmHotel.frx":26E78
            PICN            =   "frmHotel.frx":26E94
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdsalir_alojamiento 
            Height          =   915
            Left            =   -65280
            TabIndex        =   177
            Top             =   2820
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   1614
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
            MICON           =   "frmHotel.frx":292DE
            PICN            =   "frmHotel.frx":292FA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfconsumoHabitacion 
            Height          =   3495
            Left            =   -74760
            TabIndex        =   178
            Top             =   900
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   6165
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
         Begin VitekeySoft.ChameleonBtn cmdEliminar_consumo_habitacion 
            Height          =   795
            Left            =   -64200
            TabIndex        =   179
            Top             =   900
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   1402
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
            MICON           =   "frmHotel.frx":29614
            PICN            =   "frmHotel.frx":29630
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdsalir_consumo_habitacion 
            Height          =   795
            Left            =   -64200
            TabIndex        =   180
            Top             =   2580
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   1402
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
            MICON           =   "frmHotel.frx":2BA7A
            PICN            =   "frmHotel.frx":2BA96
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdagregar_consumo_habitacion 
            Height          =   375
            Left            =   -66720
            TabIndex        =   184
            Top             =   4635
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "AGREGAR"
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
            MICON           =   "frmHotel.frx":2BDB0
            PICN            =   "frmHotel.frx":2BDCC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfRestaurant 
            Height          =   2895
            Left            =   -74760
            TabIndex        =   186
            Top             =   900
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   5106
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
         Begin VitekeySoft.ChameleonBtn cmd_eliminar_consumo_restaurant 
            Height          =   810
            Left            =   -63960
            TabIndex        =   187
            Top             =   900
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   1429
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
            MICON           =   "frmHotel.frx":2E3B1
            PICN            =   "frmHotel.frx":2E3CD
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdeliminar_consumo_restaurant 
            Height          =   810
            Left            =   -63960
            TabIndex        =   188
            Top             =   3420
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   1429
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
            MICON           =   "frmHotel.frx":30817
            PICN            =   "frmHotel.frx":30833
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdagregar_resta 
            Height          =   375
            Left            =   -65400
            TabIndex        =   192
            Top             =   4035
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "AGREGAR"
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
            MICON           =   "frmHotel.frx":30B4D
            PICN            =   "frmHotel.frx":30B69
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdFacturarAlojamiento 
            Height          =   915
            Left            =   -65280
            TabIndex        =   196
            Top             =   1860
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   1614
            BTYPE           =   5
            TX              =   "FACTURAR"
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
            MICON           =   "frmHotel.frx":3314E
            PICN            =   "frmHotel.frx":3316A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdfacturar_consumo 
            Height          =   795
            Left            =   -64200
            TabIndex        =   197
            Top             =   1740
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   1402
            BTYPE           =   5
            TX              =   "FACTURAR"
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
            MICON           =   "frmHotel.frx":353C3
            PICN            =   "frmHotel.frx":353DF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdFacturar_rest 
            Height          =   810
            Left            =   -63960
            TabIndex        =   198
            Top             =   1740
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   1429
            BTYPE           =   5
            TX              =   "FACTURAR"
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
            MICON           =   "frmHotel.frx":37638
            PICN            =   "frmHotel.frx":37654
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfVisita 
            Height          =   3255
            Left            =   -68760
            TabIndex        =   211
            Top             =   1260
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   5741
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
         Begin VitekeySoft.ChameleonBtn cmdAgregarVisita 
            Height          =   375
            Left            =   -68760
            TabIndex        =   212
            Top             =   4980
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            BTYPE           =   5
            TX              =   "AGREGAR ..."
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
            MICON           =   "frmHotel.frx":398AD
            PICN            =   "frmHotel.frx":398C9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdEliminarVisita 
            Height          =   375
            Left            =   -66840
            TabIndex        =   213
            Top             =   4980
            Width           =   1695
            _ExtentX        =   2990
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
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmHotel.frx":3BEAE
            PICN            =   "frmHotel.frx":3BECA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfHistorialFactura 
            Height          =   1215
            Left            =   -73200
            TabIndex        =   226
            Top             =   4140
            Width           =   4095
            _ExtentX        =   7223
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
         Begin VitekeySoft.ChameleonBtn cmdagregarComprobante 
            Height          =   330
            Left            =   -69480
            TabIndex        =   231
            Top             =   5540
            Visible         =   0   'False
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   582
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
            FCOL            =   8388608
            FCOLO           =   8388608
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmHotel.frx":3E4AF
            PICN            =   "frmHotel.frx":3E4CB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdImprimir 
            Height          =   810
            Left            =   -63960
            TabIndex        =   237
            Top             =   2580
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   1429
            BTYPE           =   5
            TX              =   "COMANDA"
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
            MICON           =   "frmHotel.frx":4081F
            PICN            =   "frmHotel.frx":4083B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfComandas 
            Height          =   1500
            Left            =   -74760
            TabIndex        =   238
            Top             =   4620
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   2646
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
         Begin VitekeySoft.ChameleonBtn cmdreimprimir 
            Height          =   810
            Left            =   -63960
            TabIndex        =   239
            Top             =   4620
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   1429
            BTYPE           =   5
            TX              =   "RE-IMPR"
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
            MICON           =   "frmHotel.frx":42E0C
            PICN            =   "frmHotel.frx":42E28
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfHistorial 
            Height          =   5055
            Left            =   360
            TabIndex        =   240
            Top             =   960
            Width           =   12975
            _ExtentX        =   22886
            _ExtentY        =   8916
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
         Begin MSComCtl2.DTPicker DtpFechaFin 
            Height          =   375
            Left            =   2520
            TabIndex        =   243
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   187891713
            CurrentDate     =   43718
         End
         Begin VitekeySoft.ChameleonBtn cmdConsultar 
            Height          =   440
            Left            =   4080
            TabIndex        =   244
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   767
            BTYPE           =   5
            TX              =   "CONSULTAR"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Cambria"
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
            MICON           =   "frmHotel.frx":453F9
            PICN            =   "frmHotel.frx":45415
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VitekeySoft.ChameleonBtn cmdReporte 
            Height          =   440
            Left            =   5760
            TabIndex        =   245
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   767
            BTYPE           =   5
            TX              =   "IMPRIMIR"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Cambria"
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
            MICON           =   "frmHotel.frx":479FA
            PICN            =   "frmHotel.frx":47A16
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblPersona 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FECHAS :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   30
            Left            =   360
            TabIndex        =   241
            Top             =   520
            Width           =   645
         End
         Begin VB.Shape Shape10 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   5775
            Left            =   80
            Top             =   360
            Width           =   13500
         End
         Begin VB.Label lblPersona 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ACOMPAÑANTES"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   26
            Left            =   -67800
            TabIndex        =   210
            Top             =   900
            Width           =   1395
         End
         Begin VB.Shape Shape9 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   4695
            Left            =   -68880
            Top             =   780
            Width           =   3855
         End
         Begin VB.Label lblfecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   25
            Left            =   -67680
            TabIndex        =   208
            Top             =   4500
            Width           =   495
         End
         Begin VB.Label lblfecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CANTIDAD"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   24
            Left            =   -68640
            TabIndex        =   207
            Top             =   4500
            Width           =   735
         End
         Begin VB.Label lblfecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRECIO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   23
            Left            =   -66600
            TabIndex        =   206
            Top             =   3900
            Width           =   495
         End
         Begin VB.Label lblfecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CANTIDAD"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   21
            Left            =   -67680
            TabIndex        =   205
            Top             =   3900
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "PRECIO REF :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   -69840
            TabIndex        =   202
            Top             =   4420
            Width           =   1080
         End
         Begin VB.Label lblproducto_resta 
            BackColor       =   &H00C0C0C0&
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
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   -72660
            TabIndex        =   191
            Top             =   4095
            Width           =   4800
         End
         Begin VB.Label lblfecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCTO :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   22
            Left            =   -74640
            TabIndex        =   189
            Top             =   4140
            Width           =   840
         End
         Begin VB.Shape Shape8 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   615
            Left            =   -74760
            Top             =   3900
            Width           =   10695
         End
         Begin VB.Label lblproducto_habitacion 
            BackColor       =   &H00C0C0C0&
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
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   -72540
            TabIndex        =   183
            Top             =   4695
            Width           =   3840
         End
         Begin VB.Label lblfecha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCTO :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   20
            Left            =   -74640
            TabIndex        =   181
            Top             =   4740
            Width           =   840
         End
         Begin VB.Shape Shape7 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   615
            Left            =   -74760
            Top             =   4500
            Width           =   10335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "FECHA :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   -74520
            TabIndex        =   173
            Top             =   4380
            Width           =   780
         End
         Begin VB.Shape Shape6 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   735
            Left            =   -74640
            Top             =   4140
            Width           =   9255
         End
         Begin VB.Label lblhabitacionnumero 
            BackColor       =   &H000080FF&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   -73200
            TabIndex        =   150
            Top             =   900
            Width           =   1455
         End
         Begin VB.Label lblEstadoReserva 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   -73200
            TabIndex        =   148
            Top             =   3540
            Width           =   3135
         End
         Begin VB.Label lblOperador 
            BackColor       =   &H00C0C0C0&
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
            Left            =   -73200
            TabIndex        =   142
            Top             =   3135
            Width           =   3135
         End
         Begin VB.Label lblHoraIngreso 
            BackColor       =   &H00C0C0C0&
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
            Left            =   -73200
            TabIndex        =   141
            Top             =   2775
            Width           =   1575
         End
         Begin VB.Label lblfecha_ingreso 
            BackColor       =   &H00C0C0C0&
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
            Left            =   -73200
            TabIndex        =   140
            Top             =   2415
            Width           =   1575
         End
         Begin VB.Label lblPersona 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OPERADOR :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   25
            Left            =   -74400
            TabIndex        =   139
            Top             =   3180
            Width           =   975
         End
         Begin VB.Label lblPersona 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "H.INGRESO :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   24
            Left            =   -74370
            TabIndex        =   138
            Top             =   2820
            Width           =   945
         End
         Begin VB.Label lblPersona 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F.INGRESO :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   23
            Left            =   -74340
            TabIndex        =   137
            Top             =   2460
            Width           =   915
         End
         Begin VB.Label lblPersona 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DIRECCION :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   22
            Left            =   -74415
            TabIndex        =   133
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lblPersona 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DATOS CLIENTE :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   21
            Left            =   -74790
            TabIndex        =   131
            Top             =   1740
            Width           =   1365
         End
         Begin VB.Label lblPersona 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DNI CLIENTE :"
            BeginProperty Font 
               Name            =   "Cambria"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   20
            Left            =   -74550
            TabIndex        =   129
            Top             =   1320
            Width           =   1125
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0C0&
            Height          =   5415
            Left            =   -74880
            Top             =   660
            Width           =   5895
         End
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         Height          =   6495
         Left            =   120
         Top             =   0
         Width           =   13935
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   19
      Left            =   13200
      TabIndex        =   121
      Top             =   7080
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   19
         Left            =   840
         TabIndex        =   171
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   19
         Left            =   2400
         TabIndex        =   126
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   19
         Left            =   240
         TabIndex        =   125
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   124
         Top             =   1155
         Width           =   2385
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   19
         Left            =   2160
         TabIndex        =   123
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   19
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   19
         Left            =   120
         TabIndex        =   122
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   18
      Left            =   9600
      TabIndex        =   115
      Top             =   7080
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   18
         Left            =   840
         TabIndex        =   170
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   18
         Left            =   2400
         TabIndex        =   120
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   18
         Left            =   240
         TabIndex        =   119
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   118
         Top             =   1155
         Width           =   2745
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   18
         Left            =   2160
         TabIndex        =   117
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   18
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   18
         Left            =   120
         TabIndex        =   116
         Top             =   600
         Width           =   1890
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   17
      Left            =   6000
      TabIndex        =   109
      Top             =   7080
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   17
         Left            =   840
         TabIndex        =   169
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   17
         Left            =   2400
         TabIndex        =   114
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   17
         Left            =   240
         TabIndex        =   113
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   112
         Top             =   1155
         Width           =   2745
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   17
         Left            =   2160
         TabIndex        =   111
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   17
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   17
         Left            =   120
         TabIndex        =   110
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   16
      Left            =   2400
      TabIndex        =   103
      Top             =   7080
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   16
         Left            =   840
         TabIndex        =   168
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   16
         Left            =   2400
         TabIndex        =   108
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   16
         Left            =   240
         TabIndex        =   107
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   106
         Top             =   1155
         Width           =   2865
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   16
         Left            =   2160
         TabIndex        =   105
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   16
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   16
         Left            =   120
         TabIndex        =   104
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   15
      Left            =   13200
      TabIndex        =   97
      Top             =   5400
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   15
         Left            =   840
         TabIndex        =   167
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   15
         Left            =   2400
         TabIndex        =   102
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   15
         Left            =   240
         TabIndex        =   101
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   100
         Top             =   1155
         Width           =   2625
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   15
         Left            =   2160
         TabIndex        =   99
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   15
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   15
         Left            =   120
         TabIndex        =   98
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   14
      Left            =   9600
      TabIndex        =   91
      Top             =   5400
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   14
         Left            =   840
         TabIndex        =   166
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   14
         Left            =   2400
         TabIndex        =   96
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   14
         Left            =   240
         TabIndex        =   95
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   94
         Top             =   1155
         Width           =   2745
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   14
         Left            =   2160
         TabIndex        =   93
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   14
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   14
         Left            =   120
         TabIndex        =   92
         Top             =   600
         Width           =   1845
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   13
      Left            =   6000
      TabIndex        =   85
      Top             =   5400
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   13
         Left            =   840
         TabIndex        =   165
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   13
         Left            =   2400
         TabIndex        =   90
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   13
         Left            =   240
         TabIndex        =   89
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   88
         Top             =   1155
         Width           =   2745
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   13
         Left            =   2160
         TabIndex        =   87
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   13
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   13
         Left            =   120
         TabIndex        =   86
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   12
      Left            =   2400
      TabIndex        =   79
      Top             =   5400
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   12
         Left            =   840
         TabIndex        =   164
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   12
         Left            =   2400
         TabIndex        =   84
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   12
         Left            =   240
         TabIndex        =   83
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   82
         Top             =   1155
         Width           =   2745
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   12
         Left            =   2160
         TabIndex        =   81
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   12
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   12
         Left            =   120
         TabIndex        =   80
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   11
      Left            =   13200
      TabIndex        =   73
      Top             =   3720
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   11
         Left            =   840
         TabIndex        =   163
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   11
         Left            =   2400
         TabIndex        =   78
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   11
         Left            =   240
         TabIndex        =   77
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   76
         Top             =   1155
         Width           =   2625
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   11
         Left            =   2160
         TabIndex        =   75
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   11
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   11
         Left            =   120
         TabIndex        =   74
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   10
      Left            =   9600
      TabIndex        =   67
      Top             =   3720
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   10
         Left            =   840
         TabIndex        =   162
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   10
         Left            =   2400
         TabIndex        =   72
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   10
         Left            =   240
         TabIndex        =   71
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   70
         Top             =   1155
         Width           =   2745
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   10
         Left            =   2160
         TabIndex        =   69
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   10
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   10
         Left            =   120
         TabIndex        =   68
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   9
      Left            =   6000
      TabIndex        =   61
      Top             =   3720
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   9
         Left            =   840
         TabIndex        =   161
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   9
         Left            =   2400
         TabIndex        =   66
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   9
         Left            =   240
         TabIndex        =   65
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   64
         Top             =   1155
         Width           =   2745
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   9
         Left            =   2160
         TabIndex        =   63
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   9
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   9
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   8
      Left            =   2400
      TabIndex        =   55
      Top             =   3720
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   8
         Left            =   840
         TabIndex        =   160
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   8
         Left            =   2400
         TabIndex        =   60
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   8
         Left            =   240
         TabIndex        =   59
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   58
         Top             =   1155
         Width           =   2505
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   8
         Left            =   2160
         TabIndex        =   57
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   8
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   8
         Left            =   120
         TabIndex        =   56
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   7
      Left            =   13200
      TabIndex        =   49
      Top             =   2040
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   7
         Left            =   840
         TabIndex        =   159
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   7
         Left            =   2400
         TabIndex        =   54
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   7
         Left            =   240
         TabIndex        =   53
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   52
         Top             =   1155
         Width           =   2505
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   7
         Left            =   2160
         TabIndex        =   51
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   7
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   50
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   6
      Left            =   9600
      TabIndex        =   43
      Top             =   2040
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   6
         Left            =   840
         TabIndex        =   158
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   6
         Left            =   2400
         TabIndex        =   48
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   47
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   46
         Top             =   1155
         Width           =   2745
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   6
         Left            =   2160
         TabIndex        =   45
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   6
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   5
      Left            =   6000
      TabIndex        =   37
      Top             =   2040
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   5
         Left            =   840
         TabIndex        =   157
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   5
         Left            =   2400
         TabIndex        =   42
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   41
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   40
         Top             =   1155
         Width           =   2985
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   5
         Left            =   2160
         TabIndex        =   39
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   5
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   5
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   4
      Left            =   2400
      TabIndex        =   31
      Top             =   2040
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   4
         Left            =   840
         TabIndex        =   156
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   4
         Left            =   2400
         TabIndex        =   36
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   35
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   34
         Top             =   1155
         Width           =   2745
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   4
         Left            =   2160
         TabIndex        =   33
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   4
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   3
      Left            =   13200
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   3
         Left            =   840
         TabIndex        =   155
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   3
         Left            =   2400
         TabIndex        =   30
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   29
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   1155
         Width           =   2625
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   3
         Left            =   2160
         TabIndex        =   27
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   3
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   2
      Left            =   9600
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   2
         Left            =   840
         TabIndex        =   154
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   2
         Left            =   2400
         TabIndex        =   24
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   1155
         Width           =   2625
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   2
         Left            =   2160
         TabIndex        =   21
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   2
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   1
      Left            =   6000
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   1
         Left            =   840
         TabIndex        =   153
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   1
         Left            =   2400
         TabIndex        =   18
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1155
         Width           =   2745
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   1
         Left            =   2160
         TabIndex        =   15
         Top             =   720
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   1
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1770
      End
   End
   Begin VB.Frame frmHabitacion 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1365
      Index           =   0
      Left            =   2400
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label lblhora 
         BackStyle       =   0  'Transparent
         Caption         =   "10:00 PM"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   0
         Left            =   720
         TabIndex        =   152
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblfecha 
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA:10-10-2019"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1770
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   0
         X1              =   2040
         X2              =   2040
         Y1              =   120
         Y2              =   840
      End
      Begin VB.Label lbltipohabitacion 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMPLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Index           =   0
         Left            =   2160
         TabIndex        =   11
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lblPersona 
         BackStyle       =   0  'Transparent
         Caption         =   "PERCY R. ANTICONA MASABEL"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1155
         Width           =   2745
      End
      Begin VB.Label lblestado 
         BackStyle       =   0  'Transparent
         Caption         =   "DISPONIBLE"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label lblnumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Index           =   0
         Left            =   2400
         TabIndex        =   8
         Top             =   120
         Width           =   495
      End
   End
   Begin VitekeySoft.ChameleonBtn cmdpiso 
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "PISO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
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
      MICON           =   "frmHotel.frx":47E68
      PICN            =   "frmHotel.frx":47E84
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdpiso 
      Height          =   975
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "PISO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
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
      MICON           =   "frmHotel.frx":4A60B
      PICN            =   "frmHotel.frx":4A627
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdpiso 
      Height          =   975
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "PISO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
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
      MICON           =   "frmHotel.frx":4CDAE
      PICN            =   "frmHotel.frx":4CDCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdpiso 
      Height          =   975
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "PISO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
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
      MICON           =   "frmHotel.frx":4F551
      PICN            =   "frmHotel.frx":4F56D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdpiso 
      Height          =   975
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "PISO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
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
      MICON           =   "frmHotel.frx":51CF4
      PICN            =   "frmHotel.frx":51D10
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdpiso 
      Height          =   975
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "PISO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
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
      MICON           =   "frmHotel.frx":54497
      PICN            =   "frmHotel.frx":544B3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdpiso 
      Height          =   975
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "PISO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
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
      MICON           =   "frmHotel.frx":56C3A
      PICN            =   "frmHotel.frx":56C56
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image cmdclose 
      Height          =   240
      Left            =   19800
      Picture         =   "frmHotel.frx":593DD
      Top             =   360
      Width           =   240
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   8415
      Left            =   17040
      Top             =   240
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   8415
      Left            =   2280
      Top             =   240
      Width           =   14655
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      Height          =   8415
      Left            =   120
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public In_Piso As String

Public Procedencia As EnumProcede



Private Sub chk_adicionar_Click()
If Me.chk_adicionar.Value = 1 Then
   
   Me.DtcTipoComprobante.Visible = True
   Me.txtSerie.Visible = True
   Me.TxtNumero.Visible = True
   Me.cmdagregarComprobante.Visible = True
   
   
   strCadena = "SELECT id_doc as Codigo,doc_des as Descripcion FROM comprobantes where id_doc in('0001','0003','0054') ORDER BY descripcion"
   Call ConfiguraRst(strCadena)
   Call LlenaDataCombo(Me.DtcTipoComprobante)
   
   Call Resalta(Me.txtSerie)
   
 Else
 Me.DtcTipoComprobante.Visible = False
   Me.txtSerie.Visible = False
   Me.TxtNumero.Visible = False
   Me.cmdagregarComprobante.Visible = False
   
End If
End Sub

Private Sub cmd_eliminar_consumo_restaurant_Click()
strCadena = "call put_agregar_consumo('" & Val(Me.HfRestaurant.TextMatrix(Me.HfRestaurant.Row, 0)) & "','" & Val(Me.txtidReserva.Text) & "','" & Val(Me.txtid_habitacion.Text) & "','" & Trim(Me.txtid_producto.Text) & "','1','0','01','" & Format(Me.DtpFechaAlojamiento.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','2','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
Call delete_kardex(Me.HfRestaurant.TextMatrix(Me.HfRestaurant.Row, 3), Me.HfRestaurant.TextMatrix(Me.HfRestaurant.Row, 5), Me.txtid_alm.Text, Me.HfRestaurant.TextMatrix(Me.HfRestaurant.Row, 1))


Call llenar_consumo_rest(Me.HfRestaurant, Val(Me.txtid_habitacion.Text), Me.txtidReserva.Text, "03")


End Sub

Private Sub cmdagregar_consumo_habitacion_Click()

strCadena = "call put_agregar_consumo(0,'" & Val(Me.txtidReserva.Text) & "','" & Val(Me.txtid_habitacion.Text) & "','" & Trim(Me.txtid_producto_habitacion.Text) & "','" & Val(Me.txtcantidad_habitacion.Text) & "','" & Val(Me.txtPrecio_habitacion.Text) & "','02','" & KEY_FECHA & "','" & KEY_USUARIO & "','1','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)


Call put_kardex_habitacion(Me.txtid_alm.Text, Me.txtid_producto_habitacion.Text, Val(Me.txtcantidad_habitacion.Text), Val(Me.txtPrecio_habitacion.Text), rst("_id_consumo"))



Call llenar_consumo(Me.HfconsumoHabitacion, Val(Me.txtid_habitacion.Text), Me.txtidReserva.Text, "02")
Me.txtid_producto_habitacion.Text = ""
lblproducto_habitacion(0).Caption = ""
Me.txtcantidad_habitacion.Text = 1

End Sub
Private Sub put_kardex_habitacion(ByVal in_almacen As String, ByVal in_producto As String, ByVal in_cantidad As Single, ByVal in_precio As Single, ByVal in_movimiento As String)

strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "' and id_alm='" & KEY_ALM & "' and id_doc='0090' LIMIT 1"
             Call ConfiguraRstlocal(strCadena)
             If rstLocal.RecordCount > 0 Then
               in_serie = rstLocal("serie")
               in_numero = formato_item(rstLocal("numero"), 8)
             
             
                strCadena = "call P_insert_compra_ultimate('0090','" & in_almacen & "','" & KEY_FECHA & "','" & KEY_FECHA & "','02'," & _
                "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
                "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
                "'0','0','0','0','0','0','0','0','0','0','0'," & _
                " '" & KEY_USUARIO & "','OBSERVACION','01','" & get_periodo_actual(KEY_FECHA) & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
                Call ConfiguraRstP(strCadena)
                id_compra = rstP(0)
                
                strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
                "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & in_producto & "','" & in_cantidad & "','" & in_precio & "'," & _
                "'0','0','0','" & in_cantidad * in_precio & "','0','0', " & _
                "'0','0','0','" & in_cantidad * in_precio & "','0','" & in_cantidad * in_precio & "','" & in_precio & "','" & in_precio & "','" & in_alm & "','" & get_producto(in_producto) & "','0','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
           
                strCadena = "call put_kardex_stock_inventario('01','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0090','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & in_producto & "','" & in_cantidad & "','" & in_precio & "','" & in_almacen & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
                Call ConfiguraRstK(strCadena)
                
                
                
                
                
                
            Else
            MsgBox "NO TIENE REGISTRADO UN COMPROBANTE SALIDA", vbInformation
                End If
End Sub
Private Sub cmdagregar_resta_Click()




strCadena = "call put_agregar_consumo(0,'" & Val(Me.txtidReserva.Text) & "','" & Val(Me.txtid_habitacion.Text) & "','" & Trim(Me.txtidproducto_resta.Text) & "','" & Val(Me.txtcantidad_resta.Text) & "','" & Val(Me.txtprecio_rest.Text) & "','03','" & KEY_FECHA & "','" & KEY_USUARIO & "','1','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

Call put_kardex_habitacion(Me.txtid_alm.Text, Me.txtidproducto_resta.Text, Val(Me.txtcantidad_resta.Text), Val(Me.txtprecio_rest.Text), 1)

Call put_detalle_consumo_combo(Me.txtidproducto_resta.Text, Val(Me.txtcantidad_resta.Text))


Call llenar_consumo_rest(Me.HfRestaurant, Val(Me.txtid_habitacion.Text), Me.txtidReserva.Text, "03")


Me.txtidproducto_resta.Text = ""
lblproducto_resta(0).Caption = ""
Me.txtcantidad_resta.Text = 1


End Sub

Private Sub cmdAgregarAlojamiento_Click()
strCadena = "call put_agregar_consumo(0,'" & Val(Me.txtidReserva.Text) & "','" & Val(Me.txtid_habitacion.Text) & "','" & Trim(Me.txtid_producto.Text) & "','1','" & Val(Me.txtPrecio_Alojamiento.Text) & "','01','" & Format(Me.DtpFechaAlojamiento.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','1','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

Call llenar_consumo(Me.HfAlojamiento, Val(Me.txtid_habitacion.Text), Me.txtidReserva.Text, "01")



End Sub

Private Sub nueva_visita()
Me.frmvisita.Visible = True
Me.txtdni_visita.Text = ""
Me.txtnombre_visita.Text = ""
Me.txtcelular_visita.Text = ""
Call Resalta(Me.txtdni_visita)
Exit Sub
End Sub

Private Sub cmdagregarComprobante_Click()
If MsgBox("Esta seguro de Vincular este Comprobante", vbQuestion + vbYesNo) = vbYes Then
    
    strCadena = "UPDATE movimiento_venta SET id_referencia='" & Val(Me.txtidReserva.Text) & "' WHERE id_venta='" & Val(Me.txtidVenta.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    CnBd.Execute (strCadena)
    Call Me.llenar_comprobantes(Me.HfHistorialFactura, Me.txtidReserva.Text)
    
End If

End Sub

Private Sub cmdAgregarVisita_Click()
Call nueva_visita
End Sub

Private Sub cmdCerrar_Click()
Me.frmreserva.Visible = False
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub put_temporal(ByVal in_producto As String)
Dim in_precio As Single

in_precio = get_precio_producto(in_producto, KEY_ALM)
FrmVentas.Show
FrmVentas.activar
If Trim(Me.txtruc.Text) <> "" Then
    FrmVentas.TxtCodCliente.Text = Trim(Me.txtruc.Text)
Else
    FrmVentas.TxtCodCliente.Text = Trim(Me.txtDni.Text)
End If

FrmVentas.precionar_cliente
FrmVentas.TxtCodProducto.Text = in_producto
FrmVentas.codigoP = in_producto
FrmVentas.txtCantidad.Text = 1
FrmVentas.TxtDescripcionProducto.Text = get_producto(in_producto)
Call FrmVentas.get_unidad(Trim(in_producto), "no")

FrmVentas.txtprecio.Text = in_precio
FrmVentas.txtServicio.Text = "si"
FrmVentas.Agregar_directo

End Sub
Private Sub put_temporal_factura(ByVal in_reserva As String)
Dim in_descripcion As String
FrmVentas.Show
FrmVentas.activar
If Trim(Me.txtruc.Text) <> "0" And Trim(Me.txtruc.Text) <> "" Then
     
    FrmVentas.TxtCodCliente.Text = Trim(Me.txtruc.Text)
    FrmVentas.DtcTipoDoc.BoundText = "0001"
Else
    FrmVentas.TxtCodCliente.Text = Trim(Me.txtDni.Text)
End If

FrmVentas.txtid_venta_ref.Text = in_reserva
FrmVentas.precionar_cliente

strCadena = "SELECT * FROM view_habitacion_alojamiento WHERE   id_reserva='" & Val(in_reserva) & "' and  ruc='" & KEY_RUC & "' ORDER BY id_tipo"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        If rst("id_tipo") = "01" Then
            in_descripcion = rst("nombre_prod") & Space(2) & "[" & Format(rst("fecha"), "dd-mm-YYYY") & "]"
        Else
            in_descripcion = rst("nombre_prod") & Space(2) & "[" & Format(rst("fecha"), "dd-mm-YYYY") & "]"
        End If
        strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,servicio,obsequio,costo) VALUES " & _
        "('" & KEY_RUC & "','" & get_unidad_producto(rst("id_producto")) & "','" & Trim(Me.txtDni.Text) & "','" & KEY_ALM & "','" & FrmVentas.DtcTipoDoc.BoundText & "','" & FrmVentas.DtcSerieDoc.BoundText & "','" & Trim(FrmVentas.TxtNumeroDoc.Text) & "','" & rst("id_producto") & "','" & rst("cantidad") & "'," & _
        "'" & rst("precio_venta") & " ','" & rst("cantidad") * rst("precio_venta") & "','0','" & KEY_CON_IGV & "','" & in_descripcion & "','" & KEY_USUARIO & "','no','no','1')"
        CnBd.Execute (strCadena)
        rst.MoveNext
   Next i
End If


Call FrmVentas.llenarGrid_det(FrmVentas.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText, Trim(FrmVentas.txtformato_impresion.Text))




End Sub
Private Sub put_temporal_factura_unificada(ByVal in_reserva As String)
Dim in_descripcion As String
FrmVentas.Show
FrmVentas.activar
If Trim(Me.txtruc.Text) <> "0" And Trim(Me.txtruc.Text) <> "" Then
     
    FrmVentas.TxtCodCliente.Text = Trim(Me.txtruc.Text)
    FrmVentas.DtcTipoDoc.BoundText = "0001"
Else
    FrmVentas.TxtCodCliente.Text = Trim(Me.txtDni.Text)
End If

FrmVentas.txtid_venta_ref.Text = in_reserva
FrmVentas.precionar_cliente
For m = 1 To Me.HfUnificada.Rows - 1
If Me.HfUnificada.TextMatrix(m, 2) = Chr(254) Then
strCadena = "SELECT * FROM view_habitacion_alojamiento WHERE id_reserva='" & Val(Me.HfUnificada.TextMatrix(m, 0)) & "' and  ruc='" & KEY_RUC & "' ORDER BY id_tipo"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   For i = 0 To rst.RecordCount - 1
        If rst("id_tipo") = "01" Then
            in_descripcion = rst("nombre_prod") & Space(2) & "[" & Format(rst("fecha"), "dd-mm-YYYY") & "]"
        Else
            in_descripcion = rst("nombre_prod") & Space(2) & "[" & Format(rst("fecha"), "dd-mm-YYYY") & "]"
        End If
        strCadena = "INSERT INTO temporal_ventas(ruc,id_unidad,id_dni,id_alm,id_doc,id_serie,numero,id_producto,cantidad,precio,total,peso,igv,detalle,dni_save,servicio,obsequio,costo) VALUES " & _
        "('" & KEY_RUC & "','" & get_unidad_producto(rst("id_producto")) & "','" & Trim(Me.txtDni.Text) & "','" & KEY_ALM & "','" & FrmVentas.DtcTipoDoc.BoundText & "','" & FrmVentas.DtcSerieDoc.BoundText & "','" & Trim(FrmVentas.TxtNumeroDoc.Text) & "','" & rst("id_producto") & "','" & rst("cantidad") & "'," & _
        "'" & rst("precio_venta") & " ','" & rst("cantidad") * rst("precio_venta") & "','0','" & KEY_CON_IGV & "','" & in_descripcion & "','" & KEY_USUARIO & "','no','no','1')"
        CnBd.Execute (strCadena)
        rst.MoveNext
   Next i
End If
End If
Next m

Call FrmVentas.llenarGrid_det(FrmVentas.HfdDetalle, FrmVentas.TxtNumeroDoc.Text, FrmVentas.DtcSerieDoc.BoundText, FrmVentas.DtcTipoDoc.BoundText, Trim(FrmVentas.txtformato_impresion.Text))




End Sub


Private Sub put_reservar(ByVal in_dni As String, ByVal in_ruc As String, ByVal in_habitacion As String)
Dim in_reserva As Double

strCadena = "CALL put_reserva_habitacion_codigo('" & in_habitacion & "','" & in_dni & "','" & in_ruc & "','" & KEY_FECHA & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)

Me.txtidReserva.Text = rst(0)

strCadena = "call put_reserva_habitacion('" & in_habitacion & "','" & Trim(Me.txtid_producto.Text) & "','" & in_dni & "','2'," & Val(Me.txtidReserva.Text) & ",'" & KEY_USUARIO & "','" & Trim(Me.txtruc.Text) & "','" & KEY_FECHA & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)


If MsgBox("Desea Realizar su comprobante de Pago", vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then
    Call put_temporal(Me.txtid_producto.Text)
End If


Call llenar_consumo(Me.HfAlojamiento, in_habitacion, Val(Me.txtidReserva.Text), "01")
Call habitaciones(In_Piso)


End Sub



Private Sub cmdComandas_Click()
Me.frmcomandas.Visible = True
strCadena = "call ADM_comanda('1','" & in_habitacion & "','" & Format(Me.DtpFechaIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFechaFin.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtCliente_busqueda.Text) & "','" & KEY_RUC & "')"

Call historial_comandas(Me.HfComandasListado, Me.txtid_habitacion.Text)

End Sub

Private Sub cmdConsultar_Click()
Call historial_habitacion(HfHistorial, Me.txtid_habitacion.Text)
End Sub

Private Sub cmdDiponible_Click()
strCadena = "call put_reserva_habitacion('" & Val(Me.txtid_habitacion.Text) & "','" & Trim(Me.txtid_producto.Text) & "','" & Trim(Me.txtDni.Text) & "','1','" & Val(Me.txtidReserva.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtruc.Text) & "','" & KEY_FECHA & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

Call habitaciones(In_Piso)
End Sub

Private Sub cmdEliminar_alojamiento_Click()

strCadena = "call put_agregar_consumo('" & Val(Me.HfAlojamiento.TextMatrix(Me.HfAlojamiento.Row, 0)) & "','" & Val(Me.txtidReserva.Text) & "','" & Val(Me.txtid_habitacion.Text) & "','" & Trim(Me.txtid_producto.Text) & "','1','0','01','" & Format(Me.DtpFechaAlojamiento.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','2','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
Call llenar_consumo(Me.HfAlojamiento, Val(Me.txtid_habitacion.Text), Me.txtidReserva.Text, "01")

End Sub

Private Sub cmdEliminar_consumo_habitacion_Click()
strCadena = "call put_agregar_consumo('" & Val(Me.HfconsumoHabitacion.TextMatrix(Me.HfconsumoHabitacion.Row, 0)) & "','" & Val(Me.txtidReserva.Text) & "','" & Val(Me.txtid_habitacion.Text) & "','" & Trim(Me.txtid_producto.Text) & "','1','0','01','" & Format(Me.DtpFechaAlojamiento.Value, "YYYY-mm-dd") & "','" & KEY_USUARIO & "','2','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

Call delete_kardex(Me.HfconsumoHabitacion.TextMatrix(Me.HfconsumoHabitacion.Row, 3), Me.HfconsumoHabitacion.TextMatrix(Me.HfconsumoHabitacion.Row, 5), Me.txtid_alm.Text, Me.HfconsumoHabitacion.TextMatrix(Me.HfconsumoHabitacion.Row, 1))


Call llenar_consumo(Me.HfconsumoHabitacion, Val(Me.txtid_habitacion.Text), Me.txtidReserva.Text, "02")
End Sub
Private Sub delete_kardex(ByVal in_producto As String, ByVal in_cantidad As Single, ByVal in_almacen As String, ByVal in_fecha As String)

strCadena = "DELETE FROM kardex WHERE  fecha_emision='" & Format(in_fecha, "YYYY-mm-dd") & "' and  id_alm='" & in_almacen & "' and id_producto='" & in_producto & "' and cantidad='" & in_cantidad & "' and ruc='" & KEY_RUC & "' LIMIT 1"
CnBd.Execute (strCadena)



End Sub


Private Sub cmdeliminar_consumo_restaurant_Click()
Me.frmreserva.Visible = False
End Sub

Private Sub cmdEliminarVisita_Click()
If Val(Me.HfVisita.Rows) > 0 Then
   
strCadena = "DELETE FROM hotel_habitacion_visita WHERE id='" & Val(Me.HfVisita.TextMatrix(Me.HfVisita.Row, 0)) & "' and id_reserva='" & Val(Me.txtidReserva.Text) & "'"
CnBd.Execute (strCadena)

Call llenar_visita(Me.HfVisita, Val(Me.txtid_habitacion.Text), Val(Me.txtidReserva.Text))
 
End If
End Sub

Private Sub cmdFactura_unificada_Click()
Call put_temporal_factura(Me.txtidReserva.Text)



End Sub

Private Sub cmdFacturaGlobal_Click()
Call put_temporal_factura_unificada(Me.txtidReserva.Text)
End Sub

Private Sub cmdgenerarReserva_Click()
    
    If Trim(Me.txtDni.Text) <> "" And Trim(Me.txtDni.Text) <> "0" Then
        Call put_reservar(Trim(Me.txtDni.Text), Trim(Me.txtruc.Text), Val(Me.txtid_habitacion.Text))
    Else
        MsgBox "Ingrese un DNI Correcto", vbInformation, KEY_VENDEDOR
    End If
End Sub

Private Sub cmdGuardarVisita_Click()

strCadena = "INSERT INTO hotel_habitacion_visita(`id_habitacion`,id_reserva,`dni`,`nombre`,`cell`,`ruc`)VALUES('" & Val(Me.txtid_habitacion.Text) & "','" & Val(Me.txtidReserva.Text) & "','" & Trim(Me.txtdni_visita.Text) & "','" & Trim(Me.txtnombre_visita.Text) & "','" & Trim(Me.txtcelular_visita.Text) & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
Me.frmvisita.Visible = False

Call llenar_visita(Me.HfVisita, Val(Me.txtid_habitacion.Text), Val(Me.txtidReserva.Text))

End Sub

Private Sub cmdImprimir_Click()
If MsgBox("Esta Seguro de Generar la Comanda", vbQuestion + vbYesNo) = vbYes Then
  Call put_comanda(Val(Me.txtidReserva.Text))
  Call llenar_comanda(Me.HfComandas, Me.txtidReserva.Text)
End If
End Sub
Private Sub put_comanda(ByVal in_reserva As String)
Dim in_comanda As String
Dim in_acumulado As Double
strCadena = "call put_comanda('" & in_reserva & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
Call ConfiguraRst(strCadena)
in_comanda = rst(0)
in_acumulado = 0
For i = 0 To Me.HfRestaurant.Rows - 1
    If Val(Me.HfRestaurant.TextMatrix(i, 0)) > 0 And Me.HfRestaurant.TextMatrix(i, 8) = Chr(254) Then
        
        strCadena = "call put_comanda_detalle('" & Val(in_comanda) & "','" & Val(Me.HfRestaurant.TextMatrix(i, 0)) & "','" & Me.HfRestaurant.TextMatrix(i, 3) & "','" & Me.HfRestaurant.TextMatrix(i, 5) & "','" & Me.HfRestaurant.TextMatrix(i, 6) & "','" & Me.HfRestaurant.TextMatrix(i, 7) & "')"
        CnBd.Execute (strCadena)
        
    End If
Next i


If MsgBox("Desea Imprimir la Comanda:" + Format(in_comanda, "000000"), vbQuestion + vbYesNo) = vbYes Then
    Call impresion_comanda(in_comanda, Me.txtidReserva.Text)
End If


End Sub



Private Sub cmdImprimirComanda_Click()

Call impresion_comanda(Val(Me.HfComandasListado.TextMatrix(Me.HfComandasListado.Row, 0)), Val(Me.HfComandasListado.TextMatrix(Me.HfComandasListado.Row, 1)))


End Sub

Private Sub cmdLimpieza_Click()

strCadena = "SELECT * FROM hotel_habitacion WHERE id_habitacion='" & Val(Me.txtid_habitacion.Text) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    If rst("id_estado") = "03" Then
        If Len(rst("dni")) > 1 Then
            strCadena = "UPDATE hotel_habitacion SET id_estado='02' WHERE  id_habitacion='" & Val(Me.txtid_habitacion.Text) & "' and ruc='" & KEY_RUC & "'"
        Else
            strCadena = "UPDATE hotel_habitacion SET id_estado='01' WHERE  id_habitacion='" & Val(Me.txtid_habitacion.Text) & "' and ruc='" & KEY_RUC & "'"
        End If
        CnBd.Execute (strCadena)
    Else
        strCadena = "call put_reserva_habitacion('" & Val(Me.txtid_habitacion.Text) & "','" & Trim(Me.txtid_producto.Text) & "','" & Trim(Me.txtDni.Text) & "','3','" & Val(Me.txtidReserva.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtruc.Text) & "','" & KEY_FECHA & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
    End If
End If

'strCadena = "call put_reserva_habitacion('" & Val(Me.txtid_habitacion.Text) & "','" & Trim(Me.txtid_producto.Text) & "','" & Trim(Me.txtDni.Text) & "','3','" & Val(Me.txtidReserva.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtruc.Text) & "','" & KEY_RUC & "')"
'CnBd.Execute (strCadena)

Call habitaciones(In_Piso)

End Sub

Private Sub cmdMantenimiento_Click()
strCadena = "call put_reserva_habitacion('" & Val(Me.txtid_habitacion.Text) & "','" & Trim(Me.txtid_producto.Text) & "','" & Trim(Me.txtDni.Text) & "','4','" & Val(Me.txtidReserva.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtruc.Text) & "','" & KEY_FECHA & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

Call habitaciones(In_Piso)
End Sub

Private Sub cmdpiso_Click(Index As Integer)
In_Piso = Me.cmdpiso(Index).Tag
Call habitaciones(Me.cmdpiso(Index).Tag)

End Sub

Private Sub cmdsoloReservar_Click()
strCadena = "call put_reserva_habitacion('" & Val(Me.txtid_habitacion.Text) & "','" & Trim(Me.txtid_producto.Text) & "','" & Trim(Me.txtDni.Text) & "','1','" & Val(Me.txtidReserva.Text) & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

Call habitaciones(In_Piso)
End Sub

Private Sub cmdProcesarTranslado_Click()

If MsgBox("Esta Seguro de realizar el Traslado.", vbInformation + vbYesNo) = vbYes Then
    Call put_traslado(Me.txtid_habitacion.Text, Me.HfTraslado.TextMatrix(Me.HfTraslado.Row, 0))
End If


End Sub
Private Sub put_traslado(ByVal in_habitacion_actual As String, ByVal in_habitacion_destino As String)

strCadena = "UPDATE hotel_habitacion_reserva SET id_habitacion='" & Val(in_habitacion_destino) & "' WHERE   id_reserva='" & Val(Me.txtidReserva.Text) & "'"
CnBd.Execute (strCadena)

strCadena = "UPDATE hotel_habitacion_visita SET id_habitacion='" & Val(in_habitacion_destino) & "' WHERE id_reserva='" & Val(Me.txtidReserva.Text) & "'"
CnBd.Execute (strCadena)

strCadena = "UPDATE hotel_habitacion_alojamiento SET id_habitacion='" & Val(in_habitacion_destino) & "' WHERE id_reserva='" & Val(Me.txtidReserva.Text) & "'"
CnBd.Execute (strCadena)

strCadena = "SELECT * FROM hotel_habitacion WHERE id_habitacion='" & Val(in_habitacion_actual) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstIN(strCadena)
If rstIN.RecordCount > 0 Then
    strCadena = "UPDATE hotel_habitacion SET id_estado='02',dni='" & rstIN("dni") & "',fecha='" & Format(rstIN("fecha"), "YYYY-mm-dd") & "',hora='" & rstIN("hora") & "',ruc_empresa='" & rstIN("ruc_empresa") & "',id_reserva='" & Val(Me.txtidReserva.Text) & "' WHERE  id_habitacion='" & Val(in_habitacion_destino) & "' "
    CnBd.Execute (strCadena)
End If


strCadena = "call put_reserva_habitacion('" & Val(in_habitacion_actual) & "','0','" & Trim(Me.txtDni.Text) & "','1','" & Val(Me.txtidReserva.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtruc.Text) & "','" & KEY_FECHA & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

MsgBox "Traslado Realizado con Exito.", vbInformation
Call habitaciones(In_Piso)
Me.frmreserva.Visible = False
Me.frmTraslado.Visible = False



End Sub



Private Sub cmdreimprimir_Click()
Call impresion_comanda(Val(Me.HfComandas.TextMatrix(Me.HfComandas.Row, 0)), Me.txtidReserva.Text)
End Sub

Private Sub cmdReporte_Click()
Dim cam(0 To 1, 1 To 2)  As String
    
    cam(0, 1) = "in_fecha_ini"
    cam(1, 1) = "in_fecha_fin"

    cam(0, 2) = Format(Me.DtpFechaIni.Value, "dd-mm-YYYY")
    cam(1, 2) = Format(Me.DtpFechaFin.Value, "dd-mm-YYYY")
    param = cam()

   strCadena = "call ADM_habitacion('1','" & Val(Me.txtid_habitacion.Text) & "','" & Format(Me.DtpFechaIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFechaFin.Value, "YYYY-mm-dd") & "','" & KEY_RUC & "')"
   Call ConfiguraRst(strCadena)
   Ans = ShowMultiReport(rst, "RptHistorialHabitacion", param, App.Path + "\Reportes\")
   
   
End Sub

Private Sub cmdsalir_alojamiento_Click()
Me.frmreserva.Visible = False
End Sub

Private Sub cmdsalir_consumo_habitacion_Click()
Me.frmreserva.Visible = False
End Sub

Private Sub cmdTrasladoInterno_Click()
'Trasladar Cosnumo a Nueva Habitacion
Me.frmTraslado.Visible = True
Call llenar_traslado(Me.HfTraslado)

End Sub



Private Sub cmdUnificada_Click()
Call Me.llenar_ocupadas(Me.HfUnificada)
Me.frmunificada.Visible = True
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.DtpFechaAlojamiento.Value = KEY_FECHA
Call pisos
Call habitaciones(In_Piso)
End Sub

Private Sub pisos()
strCadena = "SELECT * FROM hotel_piso WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   In_Piso = rst("id_piso")
    For i = 0 To rst.RecordCount - 1
        Me.cmdpiso(i).Visible = True
        Me.cmdpiso(i).Caption = rst("descripcion")
        Me.cmdpiso(i).Tag = rst("id_piso")
        
        rst.MoveNext
    Next
End If
End Sub


Private Sub habitaciones(ByVal In_Piso As String)



For i = 0 To Me.frmHabitacion.Count - 1
    Me.frmHabitacion(i).Visible = False
Next i


'strCadena = "SELECT * FROM hotel_habitacion WHERE id_piso='" & Val(In_Piso) & "' and    ruc='" & KEY_RUC & "'"
'Call ConfiguraRst(strCadena)
'If rst.RecordCount > 0 Then
'    For i = 0 To rst.RecordCount - 1
'
'         If Me.frmHabitacion(i).Visible = True Then
'            Me.frmHabitacion(i).Visible = False
'         End If
'
    '     rst.MoveNext
'    Next i
'End If






strCadena = "SELECT * FROM view_habitacion_piso WHERE id_piso='" & Val(In_Piso) & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        Select Case rst("id_estado")
               Case "01" ' DISPONIBLE
                    Me.frmHabitacion(i).BackColor = &H80FF80
               Case "02" ' OCUPADO
                    Me.frmHabitacion(i).BackColor = &H8080FF
               Case "03" ' LIMPIEZA
                    Me.frmHabitacion(i).BackColor = &HC0C000
               Case "04" ' MANTENIMIENTO
                    Me.frmHabitacion(i).BackColor = &H40C0&
        End Select
        Me.lblestado(i).Caption = rst("estado")
        Me.lblnumero(i).Caption = rst("descripcion")
        Me.lblPersona(i).Caption = rst("cliente")
        If rst("dni") = "0" Then
           Me.lblfecha(i).Caption = ""
           Me.lblHoraIngreso.Caption = ""
           Me.lblfecha(i).Visible = False
           Me.lblhora(i).Visible = False
        Else
           Me.lblfecha(i).Caption = "INGRESO: " & Format(rst("fecha"), "dd-mm-YYYY")
           Me.lblhora(i).Caption = Format(rst("hora"), "HH:mm AM/PM")
           Me.lblfecha(i).Visible = True
           Me.lblhora(i).Visible = True
        End If
        
        Me.frmHabitacion(i).Tag = rst("id_habitacion")
        Me.lbltipohabitacion(i).Caption = rst("tipo")
    
        Me.frmHabitacion(i).Visible = True
        rst.MoveNext
    Next
End If
End Sub



Private Sub frmHabitacion_Click(Index As Integer)
Me.txtid_habitacion.Text = Me.frmHabitacion(Index).Tag

Me.HfconsumoHabitacion.Rows = 0
Me.HfAlojamiento.Rows = 0
Me.HfRestaurant.Rows = 0
Me.chk_adicionar.Value = 0
Me.DtcTipoComprobante.Visible = False
Me.txtSerie.Visible = False
Me.TxtNumero.Visible = False
Me.cmdagregarComprobante.Visible = False
Me.frmTraslado.Visible = False
Me.frmunificada.Visible = False
Call put_reserva(Me.txtid_habitacion.Text)

Me.frmreserva.Visible = True

If KEY_CARGO = "00001" Then
   SSTab1.TabVisible(1) = False
   SSTab1.TabVisible(2) = False
   Me.cmdDiponible.Enabled = False
   Me.cmdFactura_unificada.Enabled = False
   Me.cmdFacturaGlobal.Enabled = False
   Me.cmdTrasladoInterno.Enabled = False
   Me.cmdUnificada.Enabled = False
   Me.cmdMantenimiento.Enabled = False
   Me.cmdLimpieza.Enabled = False
   Me.cmdgenerarReserva.Enabled = False
End If




End Sub

Private Sub put_reserva(ByVal in_habitacion As String)

strCadena = "SELECT * FROM view_habitacion_piso WHERE id_habitacion='" & Val(in_habitacion) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.lblhabitacionnumero.Caption = rst("descripcion")
   Me.txtDni.Text = rst("dni")
   Me.txtCliente.Text = rst("cliente")
   Me.txtruc.Text = rst("ruc_empresa")
   Me.txtDireccion.Text = rst("direccion")
   Me.lblfecha_ingreso.Caption = rst("fecha")
   Me.lblHoraIngreso.Caption = rst("hora")
   Me.lblEstadoReserva.Caption = rst("estado")
   Me.txtid_producto.Text = rst("id_producto")
   Me.txtidReserva.Text = rst("id_reserva")
   Me.txtid_alm.Text = rst("id_alm")
   Me.HfAlojamiento.Rows = 0
   Me.HfconsumoHabitacion.Rows = 0
   Me.HfRestaurant.Rows = 0
   Me.HfHistorialFactura.Rows = 0
   Me.HfVisita.Rows = 0
   
   If Trim(Me.txtDni.Text) <> "0" Then
      Me.lblOperador.Caption = get_persona(rst("dni_save"))
      Me.cmdgenerarReserva.Enabled = False
      Call llenar_consumo(Me.HfAlojamiento, in_habitacion, Me.txtidReserva.Text, "01")
      Call llenar_consumo(Me.HfconsumoHabitacion, in_habitacion, Me.txtidReserva.Text, "02")
      
      Call llenar_consumo_rest(Me.HfRestaurant, in_habitacion, Me.txtidReserva.Text, "03")
      Call llenar_comanda(Me.HfComandas, Me.txtidReserva.Text)
      Call Me.llenar_visita(Me.HfVisita, Me.txtid_habitacion.Text, Me.txtidReserva.Text)
      Call llenar_comprobantes(Me.HfHistorialFactura, Me.txtidReserva.Text)
   Else
     Me.cmdgenerarReserva.Enabled = True
   End If
   
   Call Resalta(Me.txtDni)
   
   
   
   
   Exit Sub
End If


End Sub

Private Sub put_persona(ByVal in_dni As String)
buscar_nuevamente:
If Len(in_dni) > 1 Then
    strCadena = "SELECT * FROM  view_entidad WHERE  dni='" & Trim(in_dni) & "'LIMIT 1 "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        If Len(in_dni) = 8 Then
            If get_dni_reniec_iii(Trim(in_dni), KEY_DEPARTAMENTO, KEY_PROVINCIA, KEY_DISTRITO) = True Then
                GoTo buscar_nuevamente
            End If
        End If
        
        Procedencia = 1
        FrmDetallePersona.Show
        
        If KEY_PAIS = "9589" Then
        
        If Len(Trim(in_dni)) = 8 Then
            nruc = "10" & Trim(in_dni)
            FrmDetallePersona.txtruc.Text = DigitoVerificadorRUC(Trim(nruc))
            FrmDetallePersona.ChkCliente.Value = 1
            Call FrmDetallePersona.precionar
            Exit Sub
        Else
            FrmDetallePersona.txtruc.Text = Trim(in_dni)
            FrmDetallePersona.ChkCliente.Value = 1
            Call FrmDetallePersona.precionar
            Exit Sub
        End If
        Else
             FrmDetallePersona.ChkCliente.Value = 1
             Call Resalta(FrmDetallePersona.txtruc)
             Exit Sub
        End If
    Else
        Me.txtCliente.Text = UCase(rst("nombre_completo"))
        Me.txtDireccion.Text = UCase(rst("direccion"))
        
        If rst("extranjero") = "si" Then
           Me.chk_extranjero.Value = 1
        Else
           Me.chk_extranjero.Value = 0
        End If
        
        
    End If
Else
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
End If


End Sub

Private Sub put_acompaniante(ByVal in_dni As String)
buscar_nuevamente:
If Len(in_dni) > 1 Then
    strCadena = "SELECT * FROM  view_entidad WHERE  dni='" & Trim(in_dni) & "'LIMIT 1 "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount < 1 Then
        If Len(in_dni) = 8 Then
            If get_dni_reniec_iii(Trim(in_dni), KEY_DEPARTAMENTO, KEY_PROVINCIA, KEY_DISTRITO) = True Then
                GoTo buscar_nuevamente
            End If
        End If
        
        Procedencia = 1
        FrmDetallePersona.Show
        
        If KEY_PAIS = "9589" Then
        
        If Len(Trim(in_dni)) = 8 Then
            nruc = "10" & Trim(in_dni)
            FrmDetallePersona.txtruc.Text = DigitoVerificadorRUC(Trim(nruc))
            FrmDetallePersona.ChkCliente.Value = 1
            Call FrmDetallePersona.precionar
            Exit Sub
        Else
            FrmDetallePersona.txtruc.Text = Trim(in_dni)
            FrmDetallePersona.ChkCliente.Value = 1
            Call FrmDetallePersona.precionar
            Exit Sub
        End If
        Else
             FrmDetallePersona.ChkCliente.Value = 1
             Call Resalta(FrmDetallePersona.txtruc)
             Exit Sub
        End If
    Else
        txtnombre_visita.Text = UCase(rst("nombre_completo"))
        
        
        
        
        
    End If
Else
    Procedencia = Selecionar
    FrmPersona.Show
    Exit Sub
End If


End Sub


Private Sub HfRestaurant_Click()

If Val(Me.HfRestaurant.TextMatrix(Me.HfRestaurant.Row, 0)) > 0 Then
    
    If Me.HfRestaurant.TextMatrix(Me.HfRestaurant.Row, 8) = Chr(168) Then
       Me.HfRestaurant.TextMatrix(Me.HfRestaurant.Row, 8) = Chr(254)
    Else
       Me.HfRestaurant.TextMatrix(Me.HfRestaurant.Row, 8) = Chr(168)
    End If
    
    
End If

End Sub

Private Sub HfUnificada_Click()



If Me.HfUnificada.TextMatrix(Me.HfUnificada.Row, 2) = Chr(254) Then
   Me.HfUnificada.TextMatrix(Me.HfUnificada.Row, 2) = Chr(168)
   For j = 1 To 2
                        HfUnificada.col = j
                        HfUnificada.Row = HfUnificada.Row
                        HfUnificada.CellBackColor = &H80000005
                    Next j
                    
Else
    Me.HfUnificada.TextMatrix(Me.HfUnificada.Row, 2) = Chr(254)
                    For j = 1 To 2
                        HfUnificada.col = j
                        HfUnificada.Row = HfUnificada.Row
                        HfUnificada.CellBackColor = &H80FF&
                    Next j
End If

 


End Sub

Private Sub Image1_Click()
Me.frmunificada.Visible = False
End Sub

Private Sub Image2_Click()
Me.frmcomandas.Visible = False
End Sub

Private Sub img_cerrar_Click()
Me.frmTraslado.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If Me.SSTab1.Tab = 4 Then
    Call historial_habitacion(HfHistorial, Me.txtid_habitacion.Text)
End If
End Sub

Private Sub txtcantidad_habitacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdagregar_consumo_habitacion.SetFocus
End If
End Sub

Private Sub txtcantidad_resta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdagregar_resta.SetFocus
End If
End Sub

Private Sub txtCliente_busqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "call ADM_comanda('4','" & in_habitacion & "','" & Format(Me.DtpFechaIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFechaFin.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtCliente_busqueda.Text) & "','" & KEY_RUC & "')"
    Call historial_comandas(Me.HfComandasListado, Me.txtid_habitacion.Text)

End If
End Sub

Private Sub txtDni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call put_persona(Trim(Me.txtDni.Text))
End If
End Sub


                        
Public Sub llenar_consumo_rest(ByVal Grilla As MSHFlexGrid, ByVal in_habitacion As String, ByVal in_reserva As String, ByVal in_tipo As String)
Dim in_total As Double
strCadena = "SELECT * FROM view_habitacion_alojamiento WHERE   id_tipo='" & in_tipo & "' and  id_habitacion='" & Val(in_habitacion) & "' and id_reserva='" & Val(in_reserva) & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 700
           Grilla.ColWidth(4) = 4000
           Grilla.ColWidth(5) = 1000
           Grilla.ColWidth(6) = 1100
           Grilla.ColWidth(7) = 1100
           Grilla.ColWidth(8) = 700
        Next
        cabecera = "ID" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "CANTIDAD" & vbTab & "PRECIO" & vbTab & "TOTAL" & vbTab & "ESTADO"
        Grilla.AddItem cabecera
         For k = 1 To 8
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
        in_total = 0
      
            
        For i = 0 To rstI.RecordCount - 1
          
          If rstI("estado") = "si" Then
             in_estado = Chr(1)
          Else
             in_estado = Chr(168)
          End If
          
          Fila = rstI("id") & vbTab & Format(rstI("fecha"), "dd-mm-YYYY") & vbTab & Format(rstI("hora"), "HH:mm AM/PM") & vbTab & rstI("id_producto") & vbTab & rstI("nombre_prod") & vbTab & rstI("cantidad") & vbTab & Format(rstI("precio_venta"), "#,##0.00") & vbTab & Format(rstI("precio_venta") * rstI("cantidad"), "#,##0.00") & vbTab & in_estado
          Grilla.AddItem Fila
          
          With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 8 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
                            
                        End With
          
          
          
          in_total = in_total + rstI("precio_venta") * rstI("cantidad")
       
          
          rstI.MoveNext
      Next i
       Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(in_total, "#,##0.00")
       Grilla.AddItem Fila
       
       For k = 6 To 7
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HDFDFE0
         Next k
         
    
End Sub
                    
                        
                        
Public Sub llenar_consumo(ByVal Grilla As MSHFlexGrid, ByVal in_habitacion As String, ByVal in_reserva As String, ByVal in_tipo As String)
Dim in_total As Double
strCadena = "SELECT * FROM view_habitacion_alojamiento WHERE   id_tipo='" & in_tipo & "' and  id_habitacion='" & Val(in_habitacion) & "' and id_reserva='" & Val(in_reserva) & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 800
           Grilla.ColWidth(3) = 700
           Grilla.ColWidth(4) = 3000
           Grilla.ColWidth(5) = 1000
           Grilla.ColWidth(6) = 1000
           Grilla.ColWidth(7) = 1200
        Next
        cabecera = "ID" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "CANTIDAD" & vbTab & "PRECIO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 1 To 7
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
        in_total = 0
      
            
        For i = 0 To rstI.RecordCount - 1
          
          Fila = rstI("id") & vbTab & Format(rstI("fecha"), "dd-mm-YYYY") & vbTab & Format(rstI("hora"), "HH:mm AM/PM") & vbTab & rstI("id_producto") & vbTab & rstI("nombre_prod") & vbTab & rstI("cantidad") & vbTab & Format(rstI("precio_venta"), "#,##0.00") & vbTab & Format(rstI("precio_venta") * rstI("cantidad"), "#,##0.00")
          Grilla.AddItem Fila
          in_total = in_total + rstI("precio_venta") * rstI("cantidad")
       
          
          rstI.MoveNext
      Next i
       Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(in_total, "#,##0.00")
       Grilla.AddItem Fila
       
       For k = 6 To 7
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HDFDFE0
         Next k
         
    
End Sub
Public Sub historial_habitacion(ByVal Grilla As MSHFlexGrid, ByVal in_habitacion As String)
Dim in_total As Double
strCadena = "call ADM_habitacion('1','" & in_habitacion & "','" & Format(Me.DtpFechaIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFechaFin.Value, "YYYY-mm-dd") & "','" & KEY_RUC & "')"
Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1000
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 4000
           Grilla.ColWidth(5) = 2000
           Grilla.ColWidth(6) = 1200
           
        Next
        cabecera = "ID" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "DNI" & vbTab & "CLIENTE" & vbTab & "HABITACION" & vbTab & "CONSUMO"
        Grilla.AddItem cabecera
         For k = 1 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
        in_total = 0
      
            
        For i = 0 To rstI.RecordCount - 1
          
          Fila = rstI("id_reserva") & vbTab & Format(rstI("fecha"), "dd-mm-YYYY") & vbTab & Format(rstI("hora"), "HH:mm AM/PM") & vbTab & rstI("dni") & vbTab & rstI("nombre_completo") & vbTab & rstI("descripcion") & vbTab & Format(rstI("consumo"), "#,##0.00")
          Grilla.AddItem Fila
          in_total = in_total + rstI("consumo")
       
          
          rstI.MoveNext
      Next i
       Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(in_total, "#,##0.00")
       Grilla.AddItem Fila
       
       For k = 6 To 6
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HDFDFE0
       Next k
         
    
End Sub

Public Sub historial_comandas(ByVal Grilla As MSHFlexGrid, ByVal in_habitacion As String)
Dim in_total As Double
Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 0
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1000
           Grilla.ColWidth(5) = 2500
           Grilla.ColWidth(6) = 1000
           Grilla.ColWidth(7) = 1000
           Grilla.ColWidth(8) = 2700
           Grilla.ColWidth(9) = 1100
           Grilla.ColWidth(10) = 1200
        Next
        cabecera = "COMANDA" & vbTab & "RESERVA" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "DNI" & vbTab & "CLIENTE" & vbTab & "HABITACION" & vbTab & "CODIGO" & vbTab & "DETALLE" & vbTab & "TOTAL" & vbTab & "MOZO"
        Grilla.AddItem cabecera
         For k = 0 To 10
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
        in_total = 0
      
            
        For i = 0 To rstI.RecordCount - 1
          
          Fila = rstI("id_comanda") & vbTab & rstI("id_reserva") & vbTab & Format(rstI("fecha_hora"), "dd-mm-YYYY") & vbTab & Format(rstI("hora"), "HH:mm AM/PM") & vbTab & rstI("dni") & vbTab & rstI("nombre_completo") & vbTab & rstI("habitacion") & vbTab & rstI("id_producto") & vbTab & rstI("nombre_prod") & vbTab & Format(rstI("total"), "#,##0.00") & vbTab & rstI("mozo")
          Grilla.AddItem Fila
          
       in_total = in_total + rstI("total")
          
          rstI.MoveNext
      Next i
       Fila = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "TOTAL:" & vbTab & Format(in_total, "#,##0.00")
       Grilla.AddItem Fila
       
       For k = 6 To 6
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HDFDFE0
       Next k
         
    
End Sub


Public Sub llenar_visita(ByVal Grilla As MSHFlexGrid, ByVal in_habitacion As String, ByVal in_reserva As String)
Dim in_total As Double
strCadena = "SELECT * FROM hotel_habitacion_visita WHERE  id_habitacion='" & Val(in_habitacion) & "' and id_reserva='" & Val(in_reserva) & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 900
           Grilla.ColWidth(2) = 2200
           Grilla.ColWidth(3) = 1200
           
        Next
        cabecera = "ID" & vbTab & "DNI" & vbTab & "NOMBRE" & vbTab & "CELL"
        Grilla.AddItem cabecera
         For k = 1 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
      
      
            
        For i = 0 To rstI.RecordCount - 1
          
          Fila = rstI("id") & vbTab & rstI("dni") & vbTab & rstI("nombre") & vbTab & rstI("cell")
          Grilla.AddItem Fila
          
       
          
          rstI.MoveNext
      Next i
       
         
    
End Sub

Public Sub llenar_comanda(ByVal Grilla As MSHFlexGrid, ByVal in_reserva As String)
Dim in_total As Double
strCadena = "SELECT * FROM view_comanda WHERE   id_reserva='" & Val(in_reserva) & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 1100
           Grilla.ColWidth(1) = 1100
           Grilla.ColWidth(2) = 1100
           Grilla.ColWidth(3) = 3000
           Grilla.ColWidth(4) = 1200
           
        Next
        cabecera = "NUMERO" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "MOZO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 1 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
      
      
            
        For i = 0 To rstI.RecordCount - 1
          
          Fila = Format(rstI("id_comanda"), "000000") & vbTab & Format(rstI("fecha_hora"), "dd-mm-YYYY") & vbTab & rstI("hora") & vbTab & rstI("nombre_completo") & vbTab & Format(rstI("total"), "##0.00")
          Grilla.AddItem Fila
          
       
          
          rstI.MoveNext
      Next i
       
         
    
End Sub

Public Sub llenar_comprobantes(ByVal Grilla As MSHFlexGrid, ByVal in_reserva As String)
Dim in_total As Double

strCadena = "SELECT * FROM movimiento_venta WHERE id_referencia='" & Val(in_reserva) & "' and  ruc='" & KEY_RUC & "'"
Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

       Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 1000
           Grilla.ColWidth(1) = 2000
           Grilla.ColWidth(2) = 1000
           
           
        Next
        cabecera = "FECHA" & vbTab & "COMPROBANTE" & vbTab & "MONTO"
        Grilla.AddItem cabecera
         For k = 1 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
      
      
            
        For i = 0 To rstI.RecordCount - 1
          
          Fila = rstI("fecha_emision") & vbTab & rstI("documento") & vbTab & Format(rstI("total"), "#,##0.00")
          Grilla.AddItem Fila
          
       
          
          rstI.MoveNext
      Next i
       
         
    
End Sub

Public Sub llenar_traslado(ByVal Grilla As MSHFlexGrid)
Dim in_total As Double


strCadena = "SELECT * FROM hotel_habitacion WHERE  id_estado='01'  and  ruc='" & KEY_RUC & "' ORDER BY id_piso,descripcion"
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
           
           
        Next
        cabecera = "ID" & vbTab & "DESCRIPCION"
        Grilla.AddItem cabecera
         For k = 1 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
      
      
            
        For i = 0 To rstI.RecordCount - 1
          
          Fila = rstI("id_HABITACION") & vbTab & "            HAB : " & rstI("descripcion")
          Grilla.AddItem Fila
          
       
          
          rstI.MoveNext
      Next i
       
         
    
End Sub

Public Sub llenar_ocupadas(ByVal Grilla As MSHFlexGrid)
Dim in_total As Double


strCadena = "SELECT * FROM hotel_habitacion WHERE  id_estado='02'  and  ruc='" & KEY_RUC & "' ORDER BY id_piso,descripcion"
Call ConfiguraRstI(strCadena)
If rstI.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub
End If

      Grilla.Rows = 0
       ReDim arrColWidth(1 To rstI.Fields.Count)
       
        For Each Campo In rstI.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 2500
           Grilla.ColWidth(2) = 500
           
        Next
        cabecera = "ID" & vbTab & "DESCRIPCION"
        Grilla.AddItem cabecera
         For k = 1 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rstI.MoveFirst
      
      
            
        For i = 0 To rstI.RecordCount - 1
          If Val(Me.txtid_habitacion.Text) = rstI("id_habitacion") Then
             in_estado = Chr(254)
          Else
             in_estado = Chr(168)
          End If
          
          Fila = rstI("id_reserva") & vbTab & "HAB : " & rstI("descripcion") & vbTab & in_estado
          Grilla.AddItem Fila
          c = 2
        
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = c '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
                            
                        End With
            
                If in_estado = Chr(254) Then
                    
                    For j = 1 To 2
                        Grilla.col = j
                        Grilla.Row = i + 1
                        Grilla.CellBackColor = &H80FF&
                    Next j
                    
                End If
          
          rstI.MoveNext
      Next i
       
         
    
End Sub



Private Sub txtdni_visita_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call put_acompaniante(Trim(Me.txtdni_visita.Text))
    Call Resalta(Me.txtcelular_visita)
End If
End Sub

Private Sub txtHabitacion_busqueda_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Me.chk_rango_fecha.Value = 1 Then
        strCadena = "call ADM_comanda('3','" & Trim(Me.txtHabitacion_busqueda.Text) & "','" & Format(Me.DtpInicio_comanda.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFin_comanda.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtCliente_busqueda.Text) & "','" & KEY_RUC & "')"
    Else
        strCadena = "call ADM_comanda('2','" & Trim(Me.txtHabitacion_busqueda.Text) & "','" & Format(Me.DtpFechaIni.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpFechaFin.Value, "YYYY-mm-dd") & "','" & Trim(Me.txtCliente_busqueda.Text) & "','" & KEY_RUC & "')"
    End If
    
    
    Call historial_comandas(Me.HfComandasListado, Me.txtid_habitacion.Text)

End If

End Sub

Private Sub txtid_producto_habitacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
End If
End Sub

Private Sub txtidproducto_resta_KeyPress(KeyAscii As Integer)
    Procedencia = seleccionar_otro
    FrmProducto.Show
    Exit Sub
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtNumero.Text = Format(Me.TxtNumero.Text, "000000")
    strCadena = "SELECT * FROM movimiento_venta WHERE id_doc='" & Me.DtcTipoComprobante.BoundText & "' and serie='" & Trim(Me.txtSerie.Text) & "' and numero='" & Trim(Me.TxtNumero.Text) & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.txtidVenta.Text = rst("id_venta")
    End If
End If
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Me.txtSerie.Text = Format(Me.txtSerie.Text, "000")
    Call Resalta(Me.TxtNumero)
    
    
End If
End Sub

VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDetalleProducto 
   BorderStyle     =   0  'None
   Caption         =   "Detalle de Productos"
   ClientHeight    =   9240
   ClientLeft      =   2835
   ClientTop       =   2100
   ClientWidth     =   20145
   ControlBox      =   0   'False
   Icon            =   "FrmDetalleProducto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   20145
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmFarmaco 
      BackColor       =   &H00FFFFFF&
      Caption         =   "DETALLE FARMACO"
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
      Height          =   1935
      Left            =   10680
      TabIndex        =   178
      Top             =   6840
      Width           =   4095
      Begin MSComCtl2.DTPicker Dtpvencimiento 
         Height          =   300
         Left            =   1320
         TabIndex        =   186
         Top             =   600
         Width           =   1600
         _ExtentX        =   2831
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
         Format          =   199294977
         CurrentDate     =   43563
      End
      Begin VB.TextBox txtFormafarmacologica 
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
         Left            =   1320
         MaxLength       =   500
         TabIndex        =   181
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txtlote 
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
         Left            =   1320
         MaxLength       =   500
         TabIndex        =   180
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtsanitario 
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
         Left            =   1320
         MaxLength       =   500
         TabIndex        =   179
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label65 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F.FARMA:"
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
         Left            =   525
         TabIndex        =   185
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label66 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VENCIMIENTO :"
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
         Left            =   75
         TabIndex        =   184
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOTE :"
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
         Left            =   795
         TabIndex        =   183
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R.SANITARIO :"
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
         Left            =   165
         TabIndex        =   182
         Top             =   1320
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdeliminarcaracteritica 
      Caption         =   "ELIMINAR"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13560
      TabIndex        =   144
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdeliminar 
      Caption         =   "ELIMINAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13560
      TabIndex        =   143
      Top             =   1605
      Width           =   975
   End
   Begin VB.TextBox TxtPartidaArancelaria 
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
      Left            =   13320
      MaxLength       =   80
      TabIndex        =   126
      Top             =   240
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9135
      Left            =   45
      TabIndex        =   35
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   16113
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
      TabCaption(0)   =   "DETALLE PRODUCTO"
      TabPicture(0)   =   "FrmDetalleProducto.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Shape7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ShpDatos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label16"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Shape9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Shape6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "LblLaboratorio"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "LblCodigoProducto"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "LblObservacion"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "LblLinea"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "LblUnidad"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "LblReorden"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "LblDescripcion"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "LblPrecio"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "LblStockActual"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "LblStockMinimo"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label40"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label56"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblcuentacontable"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lblnuevo_codigo"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label59"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label60"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label61"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label62"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "DtcEspecialidad"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "DtcColor"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "DtcTipoProducto"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "DtaSublinea"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "DtcModelo"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "HfgBarras"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "dtcProveedor"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "DtcMarca"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "DtcLinea"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "DtcUnidad"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cmdCodBarra"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "TxtBuscarProveedor"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "TxtBuscamarca"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "TxtBuscaLinea"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "ChkRelacionados"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "CmdRelacionados"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "cmdSubProductos"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "chkSubproductos"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "TxtNombrecomercial"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "TxtTransporte"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "cmdActualizar"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "chkCombo"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "CmdQuitar"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "CmdAgregar"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "TxtCodBarra"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "cmdRelacionar"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "TxtPrecioCompra"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "Txt_Impuesto"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "ChkIGV"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "ChkPercepcion"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "TxtObservacion"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "TxtPeso"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "TxtDescripcion"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "TxtPrecio"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "TxtStockActual"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "TxtStockMinimo"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "cmddatosimportacion"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "txtBuscaSublinea"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "txttrae"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "chk_cuenta_contable"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "txtcuenta_contable"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "chkGranel"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "txtCodigoUniversal"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "txtCodAlterno"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "TxtCodProveedor"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "frmgrifo"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "txtCantidadDescuento"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "chk_icbper"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "txtModelo"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "frmUnidades"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).ControlCount=   85
      TabCaption(1)   =   "CLASIFICACION BUSQUEDA"
      TabPicture(1)   =   "FrmDetalleProducto.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label21"
      Tab(1).Control(1)=   "Label22"
      Tab(1).Control(2)=   "Label23"
      Tab(1).Control(3)=   "Label24"
      Tab(1).Control(4)=   "Label25"
      Tab(1).Control(5)=   "Label26"
      Tab(1).Control(6)=   "Label27"
      Tab(1).Control(7)=   "Label28"
      Tab(1).Control(8)=   "Label29"
      Tab(1).Control(9)=   "Label30"
      Tab(1).Control(10)=   "Label31"
      Tab(1).Control(11)=   "Label32"
      Tab(1).Control(12)=   "Label33"
      Tab(1).Control(13)=   "Shape1"
      Tab(1).Control(14)=   "Label34"
      Tab(1).Control(15)=   "Label36"
      Tab(1).Control(16)=   "Label37"
      Tab(1).Control(17)=   "Label38"
      Tab(1).Control(18)=   "Label39"
      Tab(1).Control(19)=   "DataCombo11"
      Tab(1).Control(20)=   "DataCombo10"
      Tab(1).Control(21)=   "DataCombo9"
      Tab(1).Control(22)=   "DataCombo8"
      Tab(1).Control(23)=   "DataCombo7"
      Tab(1).Control(24)=   "DataCombo6"
      Tab(1).Control(25)=   "DtcNivel4"
      Tab(1).Control(26)=   "DtcNivel3"
      Tab(1).Control(27)=   "DtcNivel2"
      Tab(1).Control(28)=   "DtcNivel1"
      Tab(1).Control(29)=   "Dtccategoria"
      Tab(1).Control(30)=   "txtCategoria"
      Tab(1).Control(31)=   "TxtCategoria1"
      Tab(1).Control(32)=   "Txtcategoria2"
      Tab(1).Control(33)=   "Text4"
      Tab(1).Control(34)=   "Text5"
      Tab(1).Control(35)=   "Text6"
      Tab(1).Control(36)=   "Text7"
      Tab(1).Control(37)=   "TxtMarca"
      Tab(1).Control(38)=   "Text10"
      Tab(1).Control(39)=   "Text11"
      Tab(1).Control(40)=   "Text12"
      Tab(1).Control(41)=   "Text13"
      Tab(1).Control(42)=   "TxtProducto"
      Tab(1).Control(43)=   "txtclave1"
      Tab(1).Control(44)=   "txtclave2"
      Tab(1).Control(45)=   "txtclave3"
      Tab(1).Control(46)=   "txtclave4"
      Tab(1).Control(47)=   "txtclave5"
      Tab(1).ControlCount=   48
      Begin VB.Frame frmUnidades 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1815
         Left            =   3840
         TabIndex        =   153
         Top             =   2880
         Visible         =   0   'False
         Width           =   4395
         Begin VB.Frame frmUnidadDetalle 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   1440
            Left            =   240
            TabIndex        =   155
            Top             =   200
            Visible         =   0   'False
            Width           =   3255
            Begin VB.TextBox txtprecio_granel 
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
               Left            =   870
               MaxLength       =   80
               TabIndex        =   176
               Top             =   780
               Width           =   1095
            End
            Begin VB.CheckBox chk_todas_sucursales 
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               Caption         =   "TODAS SUCURSALES"
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
               Left            =   840
               TabIndex        =   164
               Top             =   1155
               Width           =   2010
            End
            Begin VB.TextBox txtCantidadUnidad 
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
               Left            =   870
               MaxLength       =   80
               TabIndex        =   157
               Top             =   420
               Width           =   1095
            End
            Begin VitekeySoft.ChameleonBtn cmdagregarunidad 
               Height          =   570
               Left            =   2070
               TabIndex        =   158
               Top             =   420
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   582
               BTYPE           =   5
               TX              =   "ADD"
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
               MICON           =   "FrmDetalleProducto.frx":047A
               PICN            =   "FrmDetalleProducto.frx":0496
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin MSDataListLib.DataCombo DtcUnidad_detalle 
               Height          =   315
               Left            =   870
               TabIndex        =   161
               Top             =   75
               Width           =   1935
               _ExtentX        =   3413
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
            Begin VB.Image Image1 
               Height          =   240
               Left            =   3000
               Picture         =   "FrmDetalleProducto.frx":27EA
               Top             =   50
               Width           =   240
            End
            Begin VB.Label Label63 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PRECIO:"
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
               Left            =   285
               TabIndex        =   177
               Top             =   855
               Width           =   555
            End
            Begin VB.Label Label58 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "UNIDAD:"
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
               TabIndex        =   162
               Top             =   210
               Width           =   615
            End
            Begin VB.Label Label57 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TRAE :"
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
               Left            =   405
               TabIndex        =   156
               Top             =   555
               Width           =   435
            End
         End
         Begin VB.CommandButton cmddelete_unidad 
            Caption         =   "DELL"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   160
            Top             =   480
            Width           =   495
         End
         Begin VB.CommandButton cmdagregar_unidad 
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
            Height          =   255
            Left            =   3840
            TabIndex        =   159
            Top             =   120
            Width           =   495
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfUnidades 
            Height          =   1575
            Left            =   120
            TabIndex        =   154
            Top             =   120
            Width           =   3675
            _ExtentX        =   6482
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
         Left            =   4080
         MaxLength       =   80
         TabIndex        =   188
         Top             =   4160
         Width           =   855
      End
      Begin VB.CheckBox chk_icbper 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "ICBPER [BOLSA PLASTICO]"
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
         Height          =   375
         Left            =   7560
         TabIndex        =   187
         Top             =   7320
         Width           =   2295
      End
      Begin VB.TextBox txtCantidadDescuento 
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
         Left            =   4110
         MaxLength       =   80
         TabIndex        =   174
         Top             =   8495
         Width           =   735
      End
      Begin VB.Frame frmgrifo 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1575
         Left            =   5880
         TabIndex        =   171
         Top             =   5520
         Visible         =   0   'False
         Width           =   4215
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFTanques 
            Height          =   1335
            Left            =   60
            TabIndex        =   172
            Top             =   120
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   2355
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
         Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
            Height          =   180
            Left            =   3960
            TabIndex        =   173
            Top             =   120
            Width           =   200
            _ExtentX        =   344
            _ExtentY        =   318
            BTYPE           =   5
            TX              =   "X"
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
            MICON           =   "FrmDetalleProducto.frx":568E
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
      Begin VB.TextBox TxtCodProveedor 
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
         Left            =   5430
         MaxLength       =   80
         TabIndex        =   169
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtCodAlterno 
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
         Left            =   5430
         MaxLength       =   80
         TabIndex        =   167
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtCodigoUniversal 
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
         Left            =   1560
         MaxLength       =   80
         TabIndex        =   166
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CheckBox chkGranel 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "GRANEL"
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
         Left            =   9040
         TabIndex        =   163
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtcuenta_contable 
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
         Height          =   315
         Left            =   8640
         TabIndex        =   148
         Top             =   8160
         Width           =   1200
      End
      Begin VB.CheckBox chk_cuenta_contable 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "CUENTA CONTABLE:"
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
         Height          =   330
         Left            =   6360
         TabIndex        =   147
         Top             =   8160
         Width           =   1650
      End
      Begin VB.TextBox txttrae 
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
         Height          =   315
         Left            =   6720
         TabIndex        =   145
         Text            =   "1"
         Top             =   7320
         Width           =   720
      End
      Begin VB.TextBox txtBuscaSublinea 
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
         Left            =   4080
         MaxLength       =   80
         TabIndex        =   142
         Top             =   3675
         Width           =   855
      End
      Begin VitekeySoft.ChameleonBtn cmddatosimportacion 
         Height          =   375
         Left            =   6720
         TabIndex        =   141
         Top             =   5520
         Width           =   3135
         _ExtentX        =   4683
         _ExtentY        =   661
         BTYPE           =   5
         TX              =   "DATOS DE IMPORTACION"
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
         BCOL            =   33023
         BCOLO           =   33023
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDetalleProducto.frx":56AA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtclave5 
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
         Left            =   -72480
         MaxLength       =   80
         TabIndex        =   138
         Top             =   8400
         Width           =   3615
      End
      Begin VB.TextBox txtclave4 
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
         Left            =   -72480
         MaxLength       =   80
         TabIndex        =   137
         Top             =   8040
         Width           =   3615
      End
      Begin VB.TextBox txtclave3 
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
         Left            =   -72480
         MaxLength       =   80
         TabIndex        =   136
         Top             =   7680
         Width           =   3615
      End
      Begin VB.TextBox txtclave2 
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
         Left            =   -72480
         MaxLength       =   80
         TabIndex        =   135
         Top             =   7320
         Width           =   3615
      End
      Begin VB.TextBox txtclave1 
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
         Left            =   -72480
         MaxLength       =   80
         TabIndex        =   134
         Top             =   6960
         Width           =   3615
      End
      Begin VB.TextBox TxtProducto 
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
         Left            =   -73170
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   128
         Top             =   720
         Width           =   5655
      End
      Begin VB.TextBox Text13 
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
         Left            =   -68745
         MaxLength       =   80
         TabIndex        =   122
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox Text12 
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
         Left            =   -68745
         MaxLength       =   80
         TabIndex        =   119
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox Text11 
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
         Left            =   -68745
         MaxLength       =   80
         TabIndex        =   116
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox Text10 
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
         Left            =   -68745
         MaxLength       =   80
         TabIndex        =   113
         Top             =   4800
         Width           =   1215
      End
      Begin VB.TextBox TxtMarca 
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
         Left            =   -73170
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   112
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox Text7 
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
         Left            =   -68745
         MaxLength       =   80
         TabIndex        =   107
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox Text6 
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
         Left            =   -68745
         MaxLength       =   80
         TabIndex        =   104
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox Text5 
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
         Left            =   -68745
         MaxLength       =   80
         TabIndex        =   101
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text4 
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
         Left            =   -68745
         MaxLength       =   80
         TabIndex        =   98
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Txtcategoria2 
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
         Left            =   -68745
         MaxLength       =   80
         TabIndex        =   95
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox TxtCategoria1 
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
         Left            =   -68745
         MaxLength       =   80
         TabIndex        =   92
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtCategoria 
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
         Left            =   -68745
         MaxLength       =   80
         TabIndex        =   89
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox TxtStockMinimo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   3990
         TabIndex        =   60
         Text            =   "0"
         Top             =   6270
         Width           =   855
      End
      Begin VB.TextBox TxtStockActual 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "Sin Stock"
         Top             =   6270
         Width           =   1020
      End
      Begin VB.TextBox TxtPrecio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Left            =   1440
         TabIndex        =   58
         Text            =   "0.00"
         Top             =   6720
         Width           =   975
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
         Left            =   1440
         MaxLength       =   500
         TabIndex        =   0
         Top             =   1710
         Width           =   7575
      End
      Begin VB.TextBox TxtPeso 
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
         Left            =   5520
         TabIndex        =   57
         Text            =   "0.0"
         Top             =   6270
         Width           =   720
      End
      Begin VB.TextBox TxtObservacion 
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
         Height          =   500
         Left            =   1350
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   56
         Top             =   7230
         Width           =   2895
      End
      Begin VB.CheckBox ChkPercepcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "AFECTO A PERCEPCION"
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
         Height          =   330
         Left            =   510
         TabIndex        =   55
         Top             =   8115
         Width           =   2010
      End
      Begin VB.CheckBox ChkIGV 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "AFECTO A IGV"
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
         Height          =   330
         Left            =   510
         TabIndex        =   54
         Top             =   8400
         Width           =   2010
      End
      Begin VB.TextBox Txt_Impuesto 
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
         Left            =   4110
         MaxLength       =   80
         TabIndex        =   53
         Top             =   8180
         Width           =   735
      End
      Begin VB.TextBox TxtPrecioCompra 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
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
         Left            =   3990
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "0.00"
         Top             =   6720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdRelacionar 
         Caption         =   "Productos Combo"
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
         Left            =   8310
         TabIndex        =   51
         Top             =   4395
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtCodBarra 
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
         Left            =   1560
         MaxLength       =   80
         TabIndex        =   50
         Top             =   820
         Width           =   1575
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "+"
         Height          =   300
         Left            =   3150
         TabIndex        =   49
         Top             =   795
         Width           =   375
      End
      Begin VB.CommandButton CmdQuitar 
         Caption         =   "-"
         Height          =   300
         Left            =   3480
         TabIndex        =   48
         Top             =   795
         Width           =   375
      End
      Begin VB.CheckBox chkCombo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "COMBO"
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
         Height          =   330
         Left            =   6870
         TabIndex        =   47
         Top             =   4395
         Width           =   1170
      End
      Begin VB.CommandButton cmdActualizar 
         Height          =   375
         Left            =   4080
         Picture         =   "FrmDetalleProducto.frx":56C6
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   4680
         Width           =   375
      End
      Begin VB.TextBox TxtTransporte 
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
         Height          =   315
         Left            =   5430
         TabIndex        =   45
         Text            =   "0.0"
         Top             =   7275
         Width           =   720
      End
      Begin VB.TextBox TxtNombrecomercial 
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
         Left            =   1440
         MaxLength       =   500
         TabIndex        =   44
         Top             =   2115
         Width           =   7575
      End
      Begin VB.CheckBox chkSubproductos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "SUB-PRODUCTO"
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
         Height          =   330
         Left            =   6870
         TabIndex        =   43
         ToolTipText     =   "Productos Derivados"
         Top             =   4755
         Width           =   1410
      End
      Begin VB.CommandButton cmdSubProductos 
         Caption         =   "Sub Productos"
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
         Left            =   8310
         TabIndex        =   42
         Top             =   4755
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton CmdRelacionados 
         Caption         =   "Relacionados"
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
         Left            =   8310
         TabIndex        =   41
         Top             =   5115
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox ChkRelacionados 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "RELACIONADOS"
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
         Height          =   330
         Left            =   6870
         TabIndex        =   40
         ToolTipText     =   "Mismo Producto Diferente Unidad"
         Top             =   5115
         Width           =   1410
      End
      Begin VB.TextBox TxtBuscaLinea 
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
         Left            =   4110
         MaxLength       =   80
         TabIndex        =   39
         Top             =   3150
         Width           =   855
      End
      Begin VB.TextBox TxtBuscamarca 
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
         Left            =   4110
         MaxLength       =   80
         TabIndex        =   38
         Top             =   2550
         Width           =   855
      End
      Begin VB.TextBox TxtBuscarProveedor 
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
         Left            =   5430
         MaxLength       =   80
         TabIndex        =   37
         Top             =   5475
         Width           =   615
      End
      Begin VB.CommandButton cmdCodBarra 
         BackColor       =   &H008080FF&
         Caption         =   "GENERAR BARRA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6990
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1280
         Width           =   2895
      End
      Begin MSDataListLib.DataCombo DtcUnidad 
         Height          =   330
         Left            =   6240
         TabIndex        =   61
         Top             =   2520
         Width           =   2775
         _ExtentX        =   4895
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
      Begin MSDataListLib.DataCombo DtcLinea 
         Height          =   330
         Left            =   1440
         TabIndex        =   62
         Top             =   3150
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
      Begin MSDataListLib.DataCombo DtcMarca 
         Height          =   330
         Left            =   1440
         TabIndex        =   63
         Top             =   2550
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
      Begin MSDataListLib.DataCombo dtcProveedor 
         Height          =   330
         Left            =   1440
         TabIndex        =   64
         Top             =   5475
         Width           =   3975
         _ExtentX        =   7011
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgBarras 
         Height          =   855
         Left            =   6990
         TabIndex        =   65
         Top             =   420
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1508
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
      Begin MSDataListLib.DataCombo DtcModelo 
         Height          =   330
         Left            =   1440
         TabIndex        =   66
         Top             =   4155
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
      Begin MSDataListLib.DataCombo DtaSublinea 
         Height          =   330
         Left            =   1440
         TabIndex        =   67
         Top             =   3675
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
      Begin MSDataListLib.DataCombo DtcTipoProducto 
         Height          =   330
         Left            =   1440
         TabIndex        =   68
         Top             =   5085
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
      Begin MSDataListLib.DataCombo Dtccategoria 
         Height          =   315
         Left            =   -73170
         TabIndex        =   90
         Top             =   1440
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo DtcNivel1 
         Height          =   315
         Left            =   -73170
         TabIndex        =   93
         Top             =   1920
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo DtcNivel2 
         Height          =   315
         Left            =   -73170
         TabIndex        =   96
         Top             =   2400
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo DtcNivel3 
         Height          =   315
         Left            =   -73170
         TabIndex        =   99
         Top             =   2880
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo DtcNivel4 
         Height          =   315
         Left            =   -73170
         TabIndex        =   102
         Top             =   3360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo DataCombo6 
         Height          =   315
         Left            =   -73170
         TabIndex        =   105
         Top             =   3840
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo DataCombo7 
         Height          =   315
         Left            =   -73170
         TabIndex        =   108
         Top             =   4320
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo DataCombo8 
         Height          =   315
         Left            =   -73170
         TabIndex        =   114
         Top             =   4800
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo DataCombo9 
         Height          =   315
         Left            =   -73170
         TabIndex        =   117
         Top             =   5280
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo DataCombo10 
         Height          =   315
         Left            =   -73170
         TabIndex        =   120
         Top             =   5760
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo DataCombo11 
         Height          =   315
         Left            =   -73170
         TabIndex        =   123
         Top             =   6240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin MSDataListLib.DataCombo DtcColor 
         Height          =   330
         Left            =   1440
         TabIndex        =   139
         Top             =   4680
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
      Begin MSDataListLib.DataCombo DtcEspecialidad 
         Height          =   330
         Left            =   4080
         TabIndex        =   189
         Top             =   5085
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
      Begin VB.Label Label62 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD  :"
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
         Left            =   3120
         TabIndex        =   175
         Top             =   8520
         Width           =   825
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COD.PROVEEDOR:"
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
         Left            =   3855
         TabIndex        =   170
         Top             =   1275
         Width           =   1245
      End
      Begin VB.Label Label60 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COD.ALTERNO:"
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
         Left            =   4035
         TabIndex        =   168
         Top             =   915
         Width           =   1035
      End
      Begin VB.Label Label59 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO UNIVERSAL :"
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
         TabIndex        =   165
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label lblnuevo_codigo 
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3240
         TabIndex        =   152
         Top             =   435
         Width           =   975
      End
      Begin VB.Label lblcuentacontable 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   6360
         TabIndex        =   149
         Top             =   8520
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label56 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRAE:"
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
         Left            =   6285
         TabIndex        =   146
         Top             =   7335
         Width           =   405
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "COLOR :"
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
         Height          =   165
         Left            =   675
         TabIndex        =   140
         Top             =   4680
         Width           =   615
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLAVE BUSQUEDA 5:"
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
         Left            =   -74160
         TabIndex        =   133
         Top             =   8400
         Width           =   1545
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLAVE BUSQUEDA 4:"
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
         Left            =   -74160
         TabIndex        =   132
         Top             =   8040
         Width           =   1545
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLAVE BUSQUEDA 3:"
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
         Left            =   -74160
         TabIndex        =   131
         Top             =   7680
         Width           =   1545
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLAVE BUSQUEDA 2:"
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
         Left            =   -74160
         TabIndex        =   130
         Top             =   7320
         Width           =   1545
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLAVE BUSQUEDA 1:"
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
         Left            =   -74160
         TabIndex        =   129
         Top             =   6960
         Width           =   1545
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00DFDFE0&
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   2100
         Left            =   -74280
         Top             =   6720
         Width           =   6375
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 10 :"
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
         Left            =   -73965
         TabIndex        =   124
         Top             =   6300
         Width           =   765
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 9 :"
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
         Left            =   -73875
         TabIndex        =   121
         Top             =   5820
         Width           =   675
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 8 :"
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
         Left            =   -73875
         TabIndex        =   118
         Top             =   5340
         Width           =   675
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 7 :"
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
         Left            =   -73875
         TabIndex        =   115
         Top             =   4860
         Width           =   675
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MARCA :"
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
         Left            =   -73875
         TabIndex        =   111
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE PRODUCTO :"
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
         Left            =   -74835
         TabIndex        =   110
         Top             =   720
         Width           =   1635
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 6 :"
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
         Left            =   -73875
         TabIndex        =   109
         Top             =   4380
         Width           =   675
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 5 :"
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
         Left            =   -73875
         TabIndex        =   106
         Top             =   3900
         Width           =   675
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 4 :"
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
         Left            =   -73875
         TabIndex        =   103
         Top             =   3420
         Width           =   675
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 3 :"
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
         Left            =   -73875
         TabIndex        =   100
         Top             =   2940
         Width           =   675
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 2 :"
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
         Left            =   -73875
         TabIndex        =   97
         Top             =   2460
         Width           =   675
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIVEL 1 :"
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
         Left            =   -73875
         TabIndex        =   94
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CATEGORIA :"
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
         Left            =   -74205
         TabIndex        =   91
         Top             =   1500
         Width           =   1005
      End
      Begin VB.Label LblStockMinimo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK MINIMO :"
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
         Left            =   2745
         TabIndex        =   88
         Top             =   6360
         Width           =   1125
      End
      Begin VB.Label LblStockActual 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK ACTUAL :"
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
         Left            =   300
         TabIndex        =   87
         Top             =   6210
         Width           =   1095
      End
      Begin VB.Label LblPrecio 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO DE VENTA :"
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
         Left            =   90
         TabIndex        =   86
         Top             =   6690
         Width           =   1275
      End
      Begin VB.Label LblDescripcion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE REAL :"
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
         Height          =   165
         Left            =   165
         TabIndex        =   85
         Top             =   1770
         Width           =   1125
      End
      Begin VB.Label LblReorden 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PESO :"
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
         Left            =   5040
         TabIndex        =   84
         Top             =   6360
         Width           =   435
      End
      Begin VB.Label LblUnidad 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UNIDAD:"
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
         Left            =   5490
         TabIndex        =   83
         Top             =   2610
         Width           =   675
      End
      Begin VB.Label LblLinea 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FAMILIA :"
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
         Height          =   165
         Left            =   555
         TabIndex        =   82
         Top             =   3210
         Width           =   705
      End
      Begin VB.Label LblObservacion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACION :"
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
         Left            =   90
         TabIndex        =   81
         Top             =   7410
         Width           =   1065
      End
      Begin VB.Label LblCodigoProducto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   1560
         TabIndex        =   80
         Top             =   435
         Width           =   1575
      End
      Begin VB.Label LblLaboratorio 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MARCA :"
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
         Height          =   165
         Left            =   645
         TabIndex        =   79
         Top             =   2610
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO INTERNO:"
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
         TabIndex        =   78
         Top             =   555
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% ESPECIAL(DESC) :"
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
         Left            =   2640
         TabIndex        =   77
         Top             =   8235
         Width           =   1305
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Height          =   735
         Left            =   2550
         Top             =   8115
         Width           =   3735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIO COMPRA:"
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
         Left            =   2685
         TabIndex        =   76
         Top             =   6735
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PROVEEDOR :"
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
         Height          =   165
         Left            =   285
         TabIndex        =   75
         Top             =   5535
         Width           =   1005
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO BARRAS:"
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
         TabIndex        =   74
         Top             =   915
         Width           =   1185
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRANSPORTE :"
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
         Left            =   4440
         TabIndex        =   73
         Top             =   7335
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MODELO :"
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
         Height          =   165
         Left            =   495
         TabIndex        =   72
         Top             =   4215
         Width           =   735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUB-FAMILIA :"
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
         Height          =   165
         Left            =   230
         TabIndex        =   71
         Top             =   3675
         Width           =   1035
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   6750
         Top             =   4360
         Width           =   3135
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N. COMERCIAL :"
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
         Height          =   165
         Left            =   150
         TabIndex        =   70
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T.PRODUCTO :"
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
         Height          =   165
         Left            =   225
         TabIndex        =   69
         Top             =   5115
         Width           =   1065
      End
      Begin VB.Shape ShpDatos 
         BackColor       =   &H00DFDFE0&
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   4380
         Left            =   45
         Top             =   1590
         Width           =   9975
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00DFDFE0&
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1200
         Left            =   45
         Top             =   360
         Width           =   9975
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00DFDFE0&
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1020
         Left            =   45
         Top             =   6075
         Width           =   9975
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00DFDFE0&
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   660
         Left            =   45
         Top             =   7155
         Width           =   9975
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00DFDFE0&
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1020
         Left            =   45
         Top             =   7995
         Width           =   9975
      End
   End
   Begin VB.PictureBox Picthumbnail1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   15120
      ScaleHeight     =   4905
      ScaleMode       =   0  'User
      ScaleWidth      =   4785
      TabIndex        =   34
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "-->"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   19200
      TabIndex        =   33
      Top             =   5205
      Width           =   735
   End
   Begin VB.CommandButton CmdAgregarCaracteristica 
      BackColor       =   &H00808080&
      Caption         =   "AGREGAR"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   32
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame FrmCaracteristicas 
      BackColor       =   &H00FFFFFF&
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
      Height          =   1335
      Left            =   10680
      TabIndex        =   29
      Top             =   5400
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtCaracteristica 
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
         Height          =   525
         Left            =   240
         MaxLength       =   80
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton CmdprocesarCaracteristica 
         BackColor       =   &H000080FF&
         Caption         =   "PROCESAR CARACTERISTICA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   840
         Width           =   3495
      End
   End
   Begin VB.Frame frmcompatibilidad 
      BackColor       =   &H00FFFFFF&
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
      Height          =   2055
      Left            =   10680
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton CmdprocesarCompatibilidad 
         Caption         =   "PROCESAR COMPATIBILIDAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   26
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox TxtCompatible 
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
         MaxLength       =   80
         TabIndex        =   25
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtCodCompatible 
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
         MaxLength       =   80
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
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
         Left            =   225
         TabIndex        =   23
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CODIGO :"
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
         TabIndex        =   22
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdAgregarCompatibilidad 
      Caption         =   "AGREGAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12240
      TabIndex        =   20
      Top             =   1605
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MATRIZ UBICACION FISICA"
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
      Height          =   2655
      Left            =   15120
      TabIndex        =   6
      Top             =   5640
      Width           =   4815
      Begin VB.CheckBox chkEmpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "EMPRESA:"
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
         Left            =   120
         TabIndex        =   150
         Top             =   320
         Width           =   975
      End
      Begin VB.TextBox Txt_y 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         TabIndex        =   16
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txt_x 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   11
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox TxtAndamio 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   10
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox TxtPiso 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   9
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox TxtSector 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
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
         MaxLength       =   80
         TabIndex        =   8
         Top             =   1200
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo DtcAlmacen 
         Height          =   315
         Left            =   1200
         TabIndex        =   17
         Top             =   720
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12632319
         ForeColor       =   8388608
         Text            =   ""
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
      Begin MSDataListLib.DataCombo DtcEmpresa 
         Height          =   315
         Left            =   1200
         TabIndex        =   151
         Top             =   320
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   33023
         ForeColor       =   8388608
         Text            =   ""
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
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CASILLERO :"
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
         TabIndex        =   15
         Top             =   2280
         Width           =   825
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ANDAMIO :"
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
         Left            =   315
         TabIndex        =   14
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PISO :"
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
         Left            =   675
         TabIndex        =   13
         Top             =   1560
         Width           =   405
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SECTOR :"
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
         Left            =   450
         TabIndex        =   12
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ALMACEN :"
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
         Left            =   330
         TabIndex        =   7
         Top             =   720
         Width           =   765
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16080
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdFoto 
      Caption         =   "Seleccionar Imagen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15120
      TabIndex        =   3
      Top             =   5205
      Width           =   3495
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   17910
      TabIndex        =   1
      Top             =   8330
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1995
      _CBHeight       =   870
      _Version        =   "6.0.8169"
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
         Width           =   1875
         _ExtentX        =   3307
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
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCompatibilidad 
      Height          =   2055
      Left            =   10680
      TabIndex        =   19
      Top             =   1920
      Width           =   4095
      _ExtentX        =   7223
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfCaracteristica 
      Height          =   1335
      Left            =   10680
      TabIndex        =   28
      Top             =   5400
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2355
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
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   15240
      Top             =   120
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
            Picture         =   "FrmDetalleProducto.frx":5810
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleProducto.frx":5B2C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleProducto.frx":5F8C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleProducto.frx":63EC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleProducto.frx":6708
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleProducto.frx":6B68
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleProducto.frx":6E84
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleProducto.frx":72E4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleProducto.frx":7744
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleProducto.frx":8024
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleProducto.frx":8340
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDetalleProducto.frx":865C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblPartida 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      Height          =   705
      Left            =   10650
      TabIndex        =   127
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PARTIDA ARANCELARIA:"
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
      Left            =   10890
      TabIndex        =   125
      Top             =   240
      Width           =   1605
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1260
      Left            =   10560
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CARACTERISTICAS :"
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
      Left            =   10665
      TabIndex        =   27
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPATIBILIDAD :"
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
      Left            =   10800
      TabIndex        =   18
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2580
      Left            =   10560
      Top             =   1515
      Width           =   4335
   End
   Begin VB.Label lblCabecera 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CODIGO DE BARRAS YA REGISTRADO"
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
      Height          =   195
      Left            =   10560
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblError 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10560
      TabIndex        =   4
      Top             =   4410
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3540
      Left            =   10560
      Top             =   5280
      Width           =   4335
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   20145
   End
End
Attribute VB_Name = "FrmDetalleProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrCodTabla As String
Dim StrCodProducto As String
Dim img As String
Dim strAfectaStock As String * 2
Dim StrPercepcion As String * 2
Dim strAfectoIGV As String * 2
Dim RstAlmProd As New ADODB.Recordset
Dim sub_producto As String
Public Procedencia As EnumProcede
Dim strCombo As String
Dim PrtFoto As String
Dim FlagFoto As Boolean

Public Sub llenar_tanque(ByVal Grilla As MSHFlexGrid, ByVal in_producto As String)
On Error GoTo salir
Dim tTotal As Double
strCadena = "SELECT * FROM view_tanque_producto WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If

   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 500
           
         Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & ""
        Grilla.AddItem cabecera
         For k = 1 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            If rst("estado") = "si" Then
                estado = Chr(254)
            Else
                estado = Chr(168)
            End If
           
            Fila = rst("id") & vbTab & rst("descripcion") & vbTab & estado
            Grilla.AddItem Fila
            
            With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = 2 '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            
            End With
      
            
            
           
            rst.MoveNext
        Next i
     
  Exit Sub
salir:    MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub



Private Sub chk_cuenta_contable_Click()
If Me.chk_cuenta_contable.Value = 1 Then
   Me.txtcuenta_contable.Visible = True
   Me.lblcuentacontable.Visible = True
Else
   Me.txtcuenta_contable.Visible = False
   Me.lblcuentacontable.Visible = False
   Me.txtcuenta_contable.Text = "0"
End If
End Sub

Private Sub chkCombo_Click()
If Me.chkCombo.Value = 1 Then
    strCombo = "si"
    If Val(Me.LblCodigoProducto.Caption) > 0 Then
       Me.cmdRelacionar.Visible = True
    End If
    
Else
    strCombo = "no"
End If
End Sub

Private Sub ChkDesabilitado_Click()

End Sub


Private Sub chkEmpresa_Click()
If Me.chkEmpresa.Value = 1 Then

   strCadena = "SELECT ruc as Codigo,nombre_completo as Descripcion FROM view_empresas WHERE dni='" & KEY_USUARIO & "'"
   Call ConfiguraRstT(strCadena)
   Call LlenaDataComboT(Me.DtcEmpresa)
   
   
   Me.DtcEmpresa.Visible = True
Else
    Me.DtcEmpresa.Visible = False
End If
End Sub

Private Sub chkGranel_Click()
If Me.chkGranel.Value = 1 Then
   Me.frmUnidades.Visible = True
   Call llenarUnidad(Me.HfUnidades, Trim(Me.LblCodigoProducto.Caption))
Else
   Me.frmUnidades.Visible = False
End If
End Sub

Private Sub ChkRelacionados_Click()
If Me.ChkRelacionados.Value = 1 Then
    Me.CmdRelacionados.Visible = True
    
    Call ingreso_relacionados
Else
    Me.CmdRelacionados.Visible = False
End If
End Sub

Private Sub chkSubproductos_Click()
If Me.chkSubproductos.Value = 1 Then
   strCadena = "UPDATE producto SET id_sub_producto='si' WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
    
   Me.cmdSubProductos.Visible = True
Else
   strCadena = "UPDATE producto SET id_sub_producto='no' WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND ruc='" & KEY_RUC & "'"
   CnBd.Execute (strCadena)
    
   Me.cmdSubProductos.Visible = False
End If
End Sub

Private Sub CmdActualizar_Click()
 strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' " & _
  "  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLinea)
   
   strCadena = "SELECT id_marca as Codigo, descripcion as Descripcion FROM marca where id_usu='" & KEY_RUC & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMarca)

   strCadena = "SELECT id_und as Codigo, abreviatura as Descripcion FROM unidad WHERE id_usu='" & KEY_RUC & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcUnidad)
  
   strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM persona P,entidad_empresa E WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' AND id_proveedor='si' " & _
  " ORDER BY nombre_completo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcProveedor)
  
 
  
  strCadena = "SELECT id_tipo as Codigo, descripcion as Descripcion FROM linea_sub WHERE id_usu='" & KEY_RUC & "' AND id_linea='" & Me.DtcLinea.BoundText & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtaSublinea)
End Sub

Private Sub cmdagregar_Click()
Call agrega_barra(Trim(Me.LblCodigoProducto.Caption))
End Sub
Private Sub agrega_barra(ByVal codigo As String)
If Trim(Me.TxtCodBarra.Text) <> "" Then
    strCadena = "SELECT * FROM producto_barras WHERE cod_barra='" & Trim(Me.TxtCodBarra.Text) & "' AND id_producto='" & Trim(codigo) & "' AND ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        MsgBox "Codigo de barras ya registrado", vbInformation, KEY_EMPRESA
    Else
        strCadena = "INSERT INTO producto_barras VALUES('" & Trim(codigo) & "','" & Trim(Me.TxtCodBarra.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
    End If
    Set rst = Nothing
    Call llena_barra
End If
End Sub


Private Sub cmdagregar_unidad_Click()
  strCadena = "SELECT id_und as Codigo, CONCAT(abreviatura,':',descripcion) as Descripcion FROM unidad WHERE id_usu='" & KEY_RUC & "'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcUnidad_detalle)
  
frmUnidadDetalle.Visible = True
End Sub

Private Sub CmdAgregarCaracteristica_Click()
If Me.FrmCaracteristicas.Visible = False Then
   Me.FrmCaracteristicas.Visible = True
Else
    Me.FrmCaracteristicas.Visible = False
End If
End Sub

Private Sub cmdAgregarCompatibilidad_Click()
If Me.frmcompatibilidad.Visible = False Then
   Me.frmcompatibilidad.Visible = True
Else
    Me.frmcompatibilidad.Visible = False
End If
End Sub

Private Sub cmdAgregarUnd_Click()

End Sub

Private Sub cmdagregarunidad_Click()

If Val(Me.txtCantidadUnidad.Text) > 0 Then
   
   If Me.chk_todas_sucursales.Value = 1 Then
        strCadena = "SELECT * FROM almacen WHERE id_tipoentidad='0' and ruc='" & KEY_RUC & "'"
        Call ConfiguraRstT(strCadena)
        If rstT.RecordCount > 0 Then
           For i = 0 To rstT.RecordCount - 1
                 strCadena = "call put_unidad_producto('" & Trim(Me.LblCodigoProducto.Caption) & "','" & Me.DtcUnidad_detalle.BoundText & "','" & Val(Me.txtCantidadUnidad.Text) & "','" & rstT("id_alm") & "','" & Val(Me.txtprecio_granel.Text) & "','" & KEY_RUC & "')"
                 CnBd.Execute (strCadena)
                 rstT.MoveNext
           Next i
        End If
   Else
        strCadena = "call put_unidad_producto('" & Trim(Me.LblCodigoProducto.Caption) & "','" & Me.DtcUnidad_detalle.BoundText & "','" & Val(Me.txtCantidadUnidad.Text) & "','" & KEY_ALM & "','" & Val(Me.txtprecio_granel.Text) & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        If DtcUnidad_detalle.BoundText = Me.DtcUnidad.BoundText Then
            strCadena = "UPDATE almacen_producto SET precio_venta='" & Val(Me.txtprecio_granel.Text) & "' WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
            CnBd.Execute (strCadena)
        End If
        
   End If
   
   
   
   Call llenarUnidad(Me.HfUnidades, Trim(Me.LblCodigoProducto.Caption))
   Me.frmUnidadDetalle.Visible = False
End If

End Sub

Private Sub cmdCodBarra_Click()
Me.TxtCodBarra.Text = Trim(Me.LblCodigoProducto.Caption)
Call agrega_barra(Trim(Me.LblCodigoProducto.Caption))
End Sub



Private Sub cmddelete_unidad_Click()
If MsgBox("Desea quitar Esta UNIDAD", vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then
    Call delete_unidad(Me.HfUnidades.TextMatrix(Me.HfUnidades.Row, 0))
    Call llenarUnidad(Me.HfUnidades, Trim(Me.LblCodigoProducto.Caption))
End If
End Sub
Private Sub delete_unidad(ByVal in_id As String)
strCadena = "DELETE FROM producto_unidad WHERE id='" & Val(in_id) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
End Sub
Private Sub cmdEliminar_Click()
If Val(Me.HfCompatibilidad.TextMatrix(Me.HfCompatibilidad.Row, 0)) > 0 Then
    strCadena = "DELETE FROM producto_compatibilidad WHERE id_producto_compatible ='" & Trim(Me.HfCompatibilidad.TextMatrix(Me.HfCompatibilidad.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Call llenarCompatibilidad(Me.HfCompatibilidad, Trim(Me.LblCodigoProducto.Caption))
End If
End Sub

Private Sub cmdeliminarcaracteritica_Click()
If Val(Me.HfCaracteristica.TextMatrix(Me.HfCaracteristica.Row, 0)) > 0 Then
    strCadena = "DELETE FROM producto_caracteristicas WHERE id_detalle ='" & Trim(Me.HfCaracteristica.TextMatrix(Me.HfCaracteristica.Row, 0)) & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Call llenarCaracteristica(Me.HfCaracteristica, Trim(Me.LblCodigoProducto.Caption))
End If

End Sub

Private Sub CmdFoto_Click()
Dim ext As String
On Error GoTo finish
Me.CommonDialog1.Filter = "*.Jpg"
Me.CommonDialog1.ShowOpen
Me.Picthumbnail1.Picture = LoadPicture(Me.CommonDialog1.FileName)
PrtFoto = Trim(Me.CommonDialog1.FileName)

img = Trim(str(Me.LblCodigoProducto.Caption) & Trim(Right(PrtFoto, 4)))
strCadena = "SELECT * FROM producto_foto WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    img = Trim(str(Me.LblCodigoProducto.Caption)) & "_" + str(rstK.RecordCount + 1) & Trim(Right(PrtFoto, 4))
    
End If
Me.CmdFoto.Caption = "Archivos de Imagen" + Space(1) + "[ " + str(rstK.RecordCount) + " ]"
Call Copiar_Archivo(PrtFoto, App.Path + "\archivos\" & KEY_RUC & "\" + img)
    strCadena = "INSERT INTO producto_foto (id_producto,foto,ruc)VALUES('" & Trim(Me.LblCodigoProducto.Caption) & "','" & img & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
FlagFoto = True

Exit Sub
finish: MsgBox "La Imagen que Intenta Subir tiene que ser .JPG", vbInformation, "Imagen no Compatible"
End Sub

Private Sub cmdImpresionBarras_Click()

End Sub

Private Sub cmdmigrar_Click()

End Sub

Private Sub cmdNext_Click()


If rstI.EOF = True Or rstI.BOF = True Then
    If rstI.RecordCount < 1 Then
        Exit Sub
    End If
    rstI.MoveFirst
Else
    rstI.MoveNext
    If rstI.EOF = True Or rstI.BOF = True Then
        rstI.MoveFirst
    End If
End If
If IsNull(rstI("foto")) = False And Len(rstI("foto")) > 5 Then
    If VerificarArchivo(App.Path & "\archivos\" & KEY_RUC & "\" & rstI("foto")) = True Then
        'Me.Image1.Visible = True
        Me.Picthumbnail1.Picture = LoadPicture(App.Path + "\archivos\" + KEY_RUC + "\" + Trim(rstI("foto")))
    Else
        'Me.Image1 = Nothing
    End If
End If
End Sub







Private Sub CmdprocesarCaracteristica_Click()
    strCadena = "INSERT INTO producto_caracteristicas (id_producto,caracteristica,ruc)VALUES('" & Trim(Me.LblCodigoProducto.Caption) & "','" & Trim(Me.txtCaracteristica.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    Me.FrmCaracteristicas.Visible = False
    Call llenarCaracteristica(Me.HfCaracteristica, Trim(Me.LblCodigoProducto.Caption))
    Exit Sub
End Sub

Private Sub CmdprocesarCompatibilidad_Click()
strCadena = "SELECT * FROM producto_compatibilidad WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND id_producto_compatible='" & Trim(Me.txtCodCompatible.Text) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    strCadena = "INSERT INTO producto_compatibilidad (id_producto,id_producto_compatible,ruc)VALUES('" & Trim(Me.LblCodigoProducto.Caption) & "','" & Trim(Me.txtCodCompatible.Text) & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
     
    Me.frmcompatibilidad.Visible = False
    Call llenarCompatibilidad(Me.HfCompatibilidad, Trim(Me.LblCodigoProducto.Caption))
    Exit Sub
End If
End Sub

Private Sub CmdQuitar_Click()
strCadena = "DELETE  FROM producto_barras WHERE cod_barra='" & Trim(Me.HfgBarras.TextMatrix(Me.HfgBarras.Row, 1)) & "' AND id_producto='" & Trim(Me.HfgBarras.TextMatrix(Me.HfgBarras.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
 
Call llena_barra
End Sub

Private Sub CmdRelacionados_Click()

FrmProductoRelacionado.Show
End Sub

Private Sub cmdRelacionar_Click()
Procedencia = relacionar
FrmProductosRelacionados.Show

Call FrmProductosRelacionados.LLENA
End Sub
Private Sub ingreso_relacionados()
Dim cproducto As String
        If Trim(Me.LblCodigoProducto.Caption) = "" Then
        Call verifica
        cproducto = formato_item(ConsultaUltimoRegistro("producto", "id_producto", "ruc", KEY_RUC), 5)
         
        strCadena = "INSERT INTO producto (id_producto, id_unidad, id_linea, id_marca,nombre_prod,stock_total,stock_minimo,peso,id_percepcion,comentario,id_igv,id_sub_producto," & _
           "id_proveedor,id_auspiciador,id_combo,ruc,precio_delivery,imagen,id_tipo) VALUES ('" & cproducto & "','" & Me.DtcUnidad.BoundText & "','" & Me.DtcLinea.BoundText & "','" & Me.DtcMarca.BoundText & "'," & _
           "'" & Me.txtDescripcion.Text & "','" & Val(Me.TxtStockActual.Text) & "','" & Val(Me.TxtStockMinimo.Text) & "','" & Val(Me.txtpeso.Text) & "','" & StrPercepcion & "'," & _
           "'" & Me.txtObservacion.Text & "','" & strAfectoIGV & "','" & Trim(sub_producto) & "','" & Trim(Me.DtcProveedor.BoundText) & "','" & Me.DtcModelo.BoundText & "','" & Trim(strCombo) & "','" & KEY_RUC & "'," & _
           "'" & Val(Me.TxtTransporte.Text) & "','" & PrtFoto & "','" & Me.DtcTipoProducto.BoundText & "')"
           CnBd.Execute (strCadena)
            
           strCadena = "SELECT * FROM almacen WHERE ruc='" & KEY_RUC & "'"
           RstAlmProd.CursorLocation = adUseClient
           RstAlmProd.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
           If RstAlmProd.RecordCount <= 0 Then
                MsgBox "No hay Ningun Almacen registrado", vbInformation
                MsgBox "Producto NO Grabado, Cree un Almacen", vbInformation
                Exit Sub
           End If
           RstAlmProd.MoveFirst
           For i = 0 To RstAlmProd.RecordCount - 1
             strCadena = "INSERT INTO almacen_producto(id_alm,id_producto,ruc) VALUES ('" & RstAlmProd("id_alm") & "','" & Trim(cproducto) & "','" & Trim(KEY_RUC) & "')"
             CnBd.Execute (strCadena)
              
             RstAlmProd.MoveNext
           Next i
           Me.LblCodigoProducto.Caption = Trim(cproducto)
           Call agrega_barra(Trim(cproducto))
           Set RstAlmProd = Nothing
           Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = False
    End If
End Sub


Private Sub cmdSubProductos_Click()
Procedencia = relacionar
FrmProductoSubproducto.Show
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DtaSublinea_Change()
If Me.DtaSublinea.BoundText <> "" Then
    strCadena = "SELECT id as Codigo, descripcion as Descripcion FROM linea_modelo WHERE ruc='" & KEY_RUC & "' AND id_linea='" & Me.DtcLinea.BoundText & "' and id_sublinea='" & Me.DtaSublinea.BoundText & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtcModelo)
End If
End Sub

Private Sub DtaSublinea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Dim in_producto As String
    
    
    If Me.DtcModelo.Enabled = True Then
        Me.DtcModelo.SetFocus
    End If
    
    
    
   
    
    
    
End If
End Sub

Private Sub DtcAlmacen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM almacen_producto WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND ruc='" & KEY_RUC & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txt_x.Text = rst("casillero_x")
    Me.Txt_y.Text = rst("casillero_y")
    Me.TxtAndamio.Text = rst("andamio")
    Me.TxtPiso.Text = rst("piso")
    Me.TxtSector.Text = rst("sector")
Else
    Me.txt_x.Text = ""
    Me.Txt_y.Text = ""
    Me.TxtAndamio.Text = ""
    Me.TxtPiso.Text = ""
    Me.TxtSector.Text = ""
End If
End If
End Sub

Private Sub DtcAuspiciador_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcTipoProducto.SetFocus
End If
End Sub
Private Sub llenarCompatibilidad(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
strCadena = "SELECT * FROM view_compatibilidad WHERE id_producto='" & id_producto & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
   
    Exit Sub
End If
   
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 3200
        Next
         cabecera = "CODIGO" & vbTab & "NOMBRE PRODUCTO"
         Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id_producto_compatible") & vbTab & UCase(rst("nombre_prod"))
             Grilla.AddItem Fila
             Fila = ""
             rst.MoveNext
        Next i
  Exit Sub
End Sub
Private Sub llenarUnidad(ByVal Grilla As MSHFlexGrid, ByVal in_producto As String)
strCadena = "SELECT * FROM view_unidad_producto WHERE id_producto='" & in_producto & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub
End If
   Me.frmUnidades.Visible = True
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 600
           Grilla.ColWidth(3) = 1000
        Next
         cabecera = "CODIGO" & vbTab & "UNIDAD" & vbTab & "TRAE" & vbTab & "PRECIO"
         Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id") & vbTab & rst("descripcion") & vbTab & rst("cantidad") & vbTab & Format(rst("precio"), "#,##0.00")
             Grilla.AddItem Fila
             rst.MoveNext
        Next i
  Exit Sub
End Sub
Private Sub llenarCaracteristica(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
strCadena = "SELECT * FROM producto_caracteristicas WHERE id_producto='" & id_producto & "' AND ruc='" & KEY_RUC & "'"
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
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 4000
        Next
        If KEY_RUBRO = "00003" Then
            cabecera = "CODIGO" & vbTab & "F.FARMACOLOGICA"
        Else
            cabecera = "CODIGO" & vbTab & "CARACTERISTICAS"
        End If
         
         Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id_detalle") & vbTab & UCase(rst("caracteristica"))
             Grilla.AddItem Fila
             Fila = ""
             rst.MoveNext
        Next i
  Exit Sub
End Sub
Private Sub llenarcriterio(ByVal Grilla As MSHFlexGrid, ByVal id_producto As String)
strCadena = "SELECT * FROM producto_busqueda WHERE id_producto='" & id_producto & "' AND ruc='" & KEY_RUC & "'"
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
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 4000
        Next
         cabecera = "CODIGO" & vbTab & "CRITERIO BUSQUEDA"
         Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id_criterio") & vbTab & UCase(rst("criterio"))
             Grilla.AddItem Fila
             Fila = ""
             rst.MoveNext
        Next i
  Exit Sub
End Sub

Private Sub DtcColor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcTipoProducto.SetFocus
End If
End Sub

Private Sub DtcLinea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     strCadena = "SELECT id_tipo as Codigo, descripcion as Descripcion FROM linea_sub WHERE id_usu='" & KEY_RUC & "' AND id_linea='" & Me.DtcLinea.BoundText & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtaSublinea)
  
  
    If Me.DtaSublinea.Enabled = True Then
        Me.DtaSublinea.SetFocus
    End If
End If
End Sub

Private Sub DtcMarca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcUnidad.Enabled = True Then
        Me.DtcUnidad.SetFocus
    End If
End If
End Sub





Private Sub DtcModelo_KeyPress(KeyAscii As Integer)
Me.DtcColor.SetFocus
End Sub

Private Sub DtcTipoProducto_Change()

If Val(Me.DtcTipoProducto.BoundText) = "01" Then
    Me.DtcEspecialidad.Visible = False
Else
    Me.DtcEspecialidad.Visible = True
End If

End Sub

Private Sub DtcTipoProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtpeso)
End If
End Sub

Private Sub DtcUnidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcLinea.Enabled = True Then
        Call Resalta(Me.TxtBuscaLinea)
    End If
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If Shift = 2 And KeyCode = Asc("G") Then
    Call Save
    Exit Sub
  End If
  If KeyCode = 27 Then
    Unload Me
  End If
  
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
strCombo = "no"
'Me.Picthumbnail1.BordeEstilo = Borde4


If KEY_RUBRO = "00003" Then
    Me.Label10.Caption = "PRIN.ACTIVO:"
    Me.Label20.Caption = "F.FARMACOLOGICA"
    CmdprocesarCaracteristica.Caption = "AGREGAR F. FARMACOLOGICA"
End If


FlagFoto = False
'---------Llenar  Combos------------------------
  strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLinea)
  
  
  
  
  
  'Me.DtcLinea.BoundText = 0
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  
 
  
  strCadena = "SELECT id_marca as Codigo, descripcion as Descripcion FROM marca WHERE id_usu='" & KEY_RUC & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMarca)
  Me.DtcMarca.BoundText = 0
  
   strCadena = "SELECT id_color as Codigo, descripcion as Descripcion FROM imp_color  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcColor)
  Me.DtcColor.BoundText = 0

  strCadena = "SELECT id_und as Codigo, CONCAT(abreviatura,':',descripcion) as Descripcion FROM unidad WHERE id_usu='" & KEY_RUC & "' " & _
  " ORDER BY abreviatura"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcUnidad)

  
  
  strCadena = "SELECT id_tipoproducto as Codigo,descripcion as Descripcion FROM tipo_producto WHERE ruc='" & KEY_RUC & "' ORDER BY id_tipoproducto"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoProducto)
  
  strCadena = "SELECT dni as Codigo, nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' AND id_proveedor='si' "
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcProveedor)
  
  
  If KEY_CON_IGV = "si" Then
    Me.chkigv.Value = 1
  End If
  
'--------- llenar categorias-----------------
  strCadena = "SELECT id_categoria as Codigo, descripcion as Descripcion FROM categoria ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcCategoria)


'--------- fin categorias   -----------------

 strCadena = "SELECT id_especialidad as Codigo,descripcion as Descripcion FROM especialidad ORDER BY descripcion"
          Call ConfiguraRst(strCadena)
          Call LlenaDataCombo(Me.DtcEspecialidad)

'----------------------verificar procedencia---------------------
Select Case FrmProducto.Procedencia
        Case nuevo
            Me.cmdAgregar.Enabled = False
            Me.CmdQuitar.Enabled = False
            Me.lblnuevo_codigo.Caption = get_producto_nuevo
            Exit Sub
    Case modificar
        Me.cmdAgregar.Enabled = True
        Me.CmdQuitar.Enabled = True
      Call LLENA
      Call llenarCompatibilidad(Me.HfCompatibilidad, Trim(Me.LblCodigoProducto.Caption))
      Call llenarCaracteristica(Me.HfCaracteristica, Trim(Me.LblCodigoProducto.Caption))
      
  End Select
'-------------------------------------------
  
End Sub
Private Function get_producto_nuevo()
strCadena = "SELECT id_producto FROM producto WHERE ruc='" & KEY_RUC & "' order by id_producto DESC LIMIT 1"
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
   get_producto_nuevo = Format(Val(rstP("id_producto")) + 1, "00000")
Else
   get_producto_nuevo = Format(1, "00000")
End If
End Function
Private Sub LLENA()
  strCadena = "SELECT * FROM producto P,almacen_producto A WHERE P.id_producto=A.id_producto AND A.id_producto = '" & FrmProducto.HfdGrilla.TextMatrix(FrmProducto.HfdGrilla.Row, 0) & "' AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND A.id_alm='" & KEY_ALM & "' LIMIT 1"
  Call EjecutaRST(strCadena)
  StrCodTabla = RstEjecuta!id_producto
  Me.LblCodigoProducto.Caption = StrCodTabla
  Me.DtcUnidad.BoundText = RstEjecuta!id_unidad
  Me.DtcLinea.BoundText = RstEjecuta!id_linea
  Me.txttrae.Text = RstEjecuta!numero_procedimientos
  Me.txtCodigoUniversal.Text = RstEjecuta!id_universal
  Me.DtcEspecialidad.BoundText = RstEjecuta!clave5
  
  strCadena = "SELECT id_tipo as Codigo, descripcion as Descripcion FROM linea_sub WHERE id_usu='" & KEY_RUC & "' AND id_linea='" & RstEjecuta!id_linea & "' " & _
  " ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtaSublinea)
  
  Me.DtaSublinea.BoundText = RstEjecuta!id_sublinea
  Me.DtcModelo.BoundText = RstEjecuta!id_modelo
  Me.DtcColor.BoundText = RstEjecuta!id_color
  Me.DtcMarca.BoundText = RstEjecuta!id_marca
  Me.TxtMarca.Text = Me.DtcMarca.Text
  Me.Txtproducto.Text = RstEjecuta!nombre_prod
  Me.DtcTipoProducto.BoundText = RstEjecuta!id_tipo
  Me.txtDescripcion.Text = RstEjecuta!nombre_prod
  Me.TxtNombrecomercial.Text = RstEjecuta!nombre_comercial
  Me.TxtStockActual.Text = RstEjecuta!stock
  Me.TxtStockMinimo.Text = RstEjecuta!stock_minimo
  Me.txtpeso.Text = RstEjecuta!Peso
  Me.TxtPrecioCompra.Text = RstEjecuta!precio_compra
  Me.txtprecio.Text = RstEjecuta!precio_venta
  Me.TxtSector.Text = RstEjecuta!sector
  Me.TxtPiso.Text = RstEjecuta!piso
  Me.TxtAndamio.Text = RstEjecuta!andamio
  Me.txt_x.Text = RstEjecuta!casillero_x
  Me.Txt_y.Text = RstEjecuta!casillero_y
  
  
  '------- forma farmaceutica
 If KEY_RUBRO = "00003" Then
     
     frmFarmaco.Visible = True
     Me.txtFormafarmacologica.Text = RstEjecuta!forma_farmacologica
     Me.txtsanitario.Text = RstEjecuta!registro_sanitario
     Me.txtLote.Text = RstEjecuta!lote
     If IsNull(RstEjecuta!vencimiento) = True Then
        Me.DtpVencimiento.Value = KEY_FECHA
     Else
        Me.DtpVencimiento.Value = RstEjecuta!vencimiento
     End If
     
  End If
  
  
  Me.txtCodAlterno.Text = RstEjecuta!codigo_alterno
  Me.TxtcodProveedor.Text = RstEjecuta!codigo_proveedor
  
  If RstEjecuta!cta_contable <> "0" Then
     Me.chk_cuenta_contable.Value = 1
     Me.txtcuenta_contable.Text = RstEjecuta!cta_contable
     Me.lblcuentacontable.Caption = UCase(get_cuenta(Trim(Me.txtcuenta_contable.Text)))
  End If
  
  
  Me.DtcProveedor.BoundText = RstEjecuta!id_proveedor
  Me.TxtTransporte.Text = Format(RstEjecuta!precio_delivery, "#,##0.00")
  Me.txtObservacion.Text = RstEjecuta!comentario
  
  Call llena_barra
  If Trim(RstEjecuta!id_percepcion) = "V" Then
    Me.ChkPercepcion.Value = 1
  Else
    Me.ChkPercepcion.Value = 0
  End If
  
  If RstEjecuta!icbper = "si" Then
     Me.chk_icbper.Value = 1
  Else
    Me.chk_icbper.Value = 0
  End If
  
  
  If RstEjecuta!id_combo = "si" Then
      Me.chkCombo.Value = 1
      Me.cmdRelacionar.Visible = True
  Else
    Me.chkCombo.Value = 0
    Me.cmdRelacionar.Visible = False
  End If
  
  If RstEjecuta!id_sub_producto = "si" Then
     Me.chkSubproductos.Value = 1
  Else
     Me.chkSubproductos.Value = 0
  End If
  
   
  If Trim(RstEjecuta!id_igv) = "si" Then
    Me.chkigv.Value = 1
  Else
    Me.chkigv.Value = 0
  End If
  
  
'--------- foto--------
strCadena = "SELECT id_producto FROM producto_foto WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND ruc='" & KEY_RUC & "' "
Call ConfiguraRstK(strCadena)
Me.CmdFoto.Caption = "Archivos de Imagen" + Space(1) + "[ " + str(rstK.RecordCount) + " ]"
If IsNull(RstEjecuta!imagen) = False And Len(RstEjecuta!imagen) > 5 Then
    If VerificarArchivo(App.Path & "\archivos\" & KEY_RUC & "\" & RstEjecuta!imagen) = True Then
        Me.Picthumbnail1.Picture = LoadPicture(App.Path + "\archivos\" + KEY_RUC + "\" + Trim(RstEjecuta!imagen))
        img = Trim(RstEjecuta!imagen)
    Else
        img = ""
    End If
End If


strCadena = "SELECT id_producto FROM producto WHERE id_relacionado='" & Trim(RstEjecuta!id_producto) & "' AND ruc='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRstT(strCadena)
If rstT.RecordCount > 0 Then
    Me.ChkRelacionados.Value = 1
Else
    Me.ChkRelacionados.Value = 0
End If
  Me.TxtPartidaArancelaria.Text = RstEjecuta!partida_arancelaria
  Me.lblPartida.Caption = RstEjecuta!partida_arancelaria
  Me.DtcCategoria.BoundText = RstEjecuta!id_categoria
  Me.DtcNivel1.BoundText = RstEjecuta!id_categoria1
  Me.DtcNivel2.BoundText = RstEjecuta!id_categoria2
  Me.txtclave1.Text = RstEjecuta!clave1
  Me.txtclave2.Text = RstEjecuta!clave2
  Me.txtclave3.Text = RstEjecuta!clave3
  Me.txtclave4.Text = RstEjecuta!clave4
  Me.txtclave5.Text = RstEjecuta!clave5
  If RstEjecuta!agranel = "si" Then
     Me.chkGranel.Value = 1
     Call llenarUnidad(Me.HfUnidades, Trim(Me.LblCodigoProducto.Caption))
  End If
  
'Call llenarCompatibilidad(Me.HfCompatibilidad, Trim(Me.LblCodigoProducto.Caption))
'Call llenarCaracteristica(Me.HfCaracteristica, Trim(Me.LblCodigoProducto.Caption))

If KEY_GRIFO = "si" Then
    Call llenar_tanque(Me.HfTanques, Trim(Me.LblCodigoProducto.Caption))
    Me.frmGrifo.Visible = True
End If

Set RstEjecuta = Nothing

End Sub
Private Sub CargarLogo()
Dim sql As String
Dim sw As String
Dim imagen As String
imagen = "Invierno.jpg"
 'Me.Image1.Picture = LoadPicture("C:\" + imagen)
'sql = "select imagen From Producto Where cProducto='" & Trim(StrCodTabla) & "'"
'Call ConfiguraRst(sql)
'If rst.RecordCount > 0 Then

'If IsNull(rst(0)) = False Then


'Image1.Picture = Leer_Imagen(CnBd, sql, "imagen")
'End If
'End If
'Set rst = Nothing
End Sub
Sub llena_barra()
Dim X As Integer
strCadena = "SELECT id_producto,cod_barra as CODIGO_BARRAS FROM producto_barras WHERE id_producto='" & Trim(Me.LblCodigoProducto.Caption) & "' AND ruc='" & KEY_RUC & "'"
   Call ConfiguraRst(strCadena)
    Set Me.HfgBarras.Recordset = rst
    Me.HfgBarras.Rows = rst.RecordCount + 1
    Me.HfgBarras.ColWidth(0) = 0
    Me.HfgBarras.ColWidth(1) = 2000
   
End Sub
Private Function get_sublinea() As Boolean
strCadena = "SELECT * FROM linea_sub WHERE id_linea='" & Me.DtcLinea.BoundText & "' and id_tipo='" & Me.DtaSublinea.BoundText & "' and id_usu='" & KEY_RUC & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    get_sublinea = True
Else
    get_sublinea = False
    MsgBox "MI ESTIMADO USUARIO:" + Chr(13) + "SELECCIONE UNA SUB-LINEA CORRECTA", vbExclamation
End If
End Function

Private Sub save_grupo_empresarial(ByVal in_producto As String, ByVal in_empresa As String)
                      
           strCadena = "SELECT * FROM producto WHERE id_producto='" & in_producto & "' and ruc='" & in_empresa & "' LIMIT 1"
           Call ConfiguraRstP(strCadena)
           If rstP.RecordCount < 1 Then
           
           
           strCadena = "INSERT INTO producto (id_producto, id_unidad, id_linea,id_sublinea,id_color, id_marca,nombre_prod,stock_total,stock_minimo,peso,id_percepcion,comentario,id_igv,id_sub_producto," & _
           "id_proveedor,id_auspiciador,id_combo,ruc,precio_delivery,imagen,id_tipo,partida_arancelaria,id_categoria,id_categoria1,id_categoria2,marca,clave1,clave2,clave3,clave4,clave5,numero_procedimientos,cta_contable) VALUES ('" & in_producto & "','" & Me.DtcUnidad.BoundText & "','" & Me.DtcLinea.BoundText & "','" & Me.DtaSublinea.BoundText & "','" & Me.DtcColor.BoundText & "','" & Me.DtcMarca.BoundText & "'," & _
           "'" & Me.txtDescripcion.Text & "','" & Val(Me.TxtStockActual.Text) & "','" & Val(Me.TxtStockMinimo.Text) & "','" & Val(Me.txtpeso.Text) & "','" & StrPercepcion & "'," & _
           "'" & Me.txtObservacion.Text & "','" & strAfectoIGV & "','" & Trim(sub_producto) & "','" & Trim(Me.DtcProveedor.BoundText) & "','" & Me.DtcModelo.BoundText & "','" & Trim(strCombo) & "','" & in_empresa & "'," & _
           "'" & Val(Me.TxtTransporte.Text) & "','" & PrtFoto & "','" & Me.DtcTipoProducto.BoundText & "','" & Trim(Me.TxtPartidaArancelaria.Text) & "','" & Me.DtcCategoria.BoundText & "','" & Me.DtcNivel1.BoundText & "','" & Me.DtcNivel2.BoundText & "','" & Trim(Me.DtcMarca.Text) & "','" & Trim(Me.txtclave1.Text) & "','" & Trim(Me.txtclave2.Text) & "','" & Trim(Me.txtclave3.Text) & "','" & Trim(Me.txtclave4.Text) & "','" & Trim(Me.txtclave5.Text) & "','" & Val(Me.txttrae.Text) & "','" & Trim(Me.txtcuenta_contable.Text) & "')"
           CnBd.Execute (strCadena)
           
           strCadena = "SELECT * FROM almacen WHERE id_tipoentidad='0' and  ruc='" & in_empresa & "' ORDER BY id_alm ASC"
           Call ConfiguraRstK(strCadena)
           If rstK.RecordCount > 0 Then
              rstK.MoveFirst
              For i = 0 To rstK.RecordCount - 1
                    strCadena = "INSERT INTO almacen_producto(id_alm,id_producto,ruc) VALUES ('" & rstK("id_alm") & "','" & Trim(in_producto) & "','" & Trim(in_empresa) & "')"
                    CnBd.Execute (strCadena)
                    rstK.MoveNext
              Next i
              
           End If
         End If
End Sub

Public Function producto_nuevo_keyfacil()
Call disabled_form(Me)
FrmLoad_web_service.Show
FrmLoad_web_service.nom_prcedimiento = "mensaje_producto_keyfacil"
Set FrmLoad_web_service.FormPadre = Me
Call FrmLoad_web_service.crear_producto_keyfacil("https://api.vitekey.com/keyfact/erp/add-product?password=vitekey2018&ruc=" & KEY_RUC & "", "POST", json_crear_producto(Trim(Me.LblCodigoProducto.Caption), Trim(Me.txtDescripcion.Text), Val(Me.txtprecio.Text)), "{x-api-token: '" & KEY_TOKEN_CLOUD & "', x-api-produccion: 'yes'}")
    
End Function

Public Sub mensaje_producto_keyfacil(ByVal strHtml As String)



                    

End Sub
Private Sub Save()
Dim i As Integer
  
  If get_sublinea = False Then
    Exit Sub
  End If
  
  If Me.txtDescripcion = "" And Val(Me.DtaSublinea.BoundText) > 0 And Val(Me.DtcLinea.BoundText) And Val(Me.DtcUnidad.BoundText) Then
       MsgBox "COMPLETE LOS CAMPOS OBLIGATORIOS" + Chr(13) + "1.- DESCRIPCION" + Chr(13) + "2.- LINEA O CLASIFICACION" + Chr(13) + "3.- SUB LINEA O MODELO" + Chr(13) + "4.- UNIDAD DE MEDIDA.", vbCritical, MSGVALIDACION
       Exit Sub
  Else
    
    
       
       If Me.chkGranel.Value = 1 Then
          in_agranel = "si"
       Else
          in_agranel = "no"
       End If
       
       If Me.chk_icbper.Value = 1 Then
          IN_ICBPER = "si"
       Else
          IN_ICBPER = "no"
       End If
       
       If Me.DtcEspecialidad.Visible = True Then
          in_especialidad = Me.DtcEspecialidad.BoundText
       Else
          in_especialidad = ""
       End If
          
          
       Call verifica '--------llama al estado de los checks
       Select Case FrmProducto.Procedencia
         Case nuevo
          Dim cproducto As String
           'cProducto = formato_item(IdInsert("Producto"), 5)
           
           If Me.chkEmpresa.Value = 1 And Val(Me.DtcEmpresa.BoundText) > 0 Then
               cproducto = get_id_producto_grupo(Me.DtcEmpresa.BoundText)
           Else
               cproducto = get_producto_nuevo
           End If
           
           
           
           
           If Len(PrtFoto) > 0 And Val(cproducto) > 0 And FlagFoto = True Then
            Call Copiar_Archivo(PrtFoto, App.Path + "\archivos\" + img)
          End If
          
          
           strCadena = "INSERT INTO producto (id_producto,codigo_alterno,codigo_proveedor,id_universal,agranel, id_unidad, id_linea,id_sublinea,id_modelo,id_color, id_marca,nombre_prod,nombre_comercial,principio_activo,stock_total,stock_minimo,peso,id_percepcion,comentario,id_igv,id_sub_producto," & _
           "id_proveedor,id_auspiciador,id_combo,ruc,precio_delivery,imagen,id_tipo,partida_arancelaria,id_categoria,id_categoria1,id_categoria2,marca,clave1,clave2,clave3,clave4,clave5,numero_procedimientos,cta_contable,vencimiento,forma_farmacologica,lote,registro_sanitario,icbper) VALUES ('" & cproducto & "','" & Trim(Me.txtCodAlterno.Text) & "','" & Trim(Me.TxtcodProveedor.Text) & "','" & Trim(Me.txtCodigoUniversal.Text) & "','" & in_agranel & "','" & Me.DtcUnidad.BoundText & "','" & Me.DtcLinea.BoundText & "','" & Me.DtaSublinea.BoundText & "','" & Me.DtcModelo.BoundText & "','" & Me.DtcColor.BoundText & "','" & Me.DtcMarca.BoundText & "'," & _
           "'" & Me.txtDescripcion.Text & "','" & UCase(Me.TxtNombrecomercial.Text) & "','" & Trim(Me.TxtNombrecomercial.Text) & "','" & Val(Me.TxtStockActual.Text) & "','" & Val(Me.TxtStockMinimo.Text) & "','" & Val(Me.txtpeso.Text) & "','" & StrPercepcion & "'," & _
           "'" & Me.txtObservacion.Text & "','" & strAfectoIGV & "','" & Trim(sub_producto) & "','" & Trim(Me.DtcProveedor.BoundText) & "','" & Me.DtcModelo.BoundText & "','" & Trim(strCombo) & "','" & KEY_RUC & "'," & _
           "'" & Val(Me.TxtTransporte.Text) & "','" & PrtFoto & "','" & Me.DtcTipoProducto.BoundText & "','" & Trim(Me.TxtPartidaArancelaria.Text) & "','" & Me.DtcCategoria.BoundText & "','" & Me.DtcNivel1.BoundText & "','" & Me.DtcNivel2.BoundText & "','" & Trim(Me.DtcMarca.Text) & "','" & Trim(Me.txtclave1.Text) & "','" & Trim(Me.txtclave2.Text) & "','" & Trim(Me.txtclave3.Text) & "','" & Trim(Me.txtclave4.Text) & "','" & Trim(in_especialidad) & "','" & Val(Me.txttrae.Text) & "','" & Trim(Me.txtcuenta_contable.Text) & "'," & _
           "'" & Format(Me.DtpVencimiento.Value, "YYYY-mm-dd") & "','" & UCase(Me.txtFormafarmacologica.Text) & "','" & Trim(Me.txtLote.Text) & "','" & Trim(Me.txtsanitario.Text) & "','" & IN_ICBPER & "')"
           CnBd.Execute (strCadena)
           Me.LblCodigoProducto.Caption = cproducto
           
           strCadena = "SELECT * FROM almacen WHERE stock='si' and  ruc='" & KEY_RUC & "' ORDER BY id_alm ASC"
           Call ConfiguraRstK(strCadena)
           If rstK.RecordCount < 1 Then
                MsgBox "No hay Ningun Almacen registrado", vbInformation
                MsgBox "Producto NO Grabado, Cree un Almacen", vbInformation
                strCadena = "DELETE FROM producto WHERE id_producto='" & cproducto & "' AND ruc='" & KEY_RUC & "'"
                CnBd.Execute (strCadena)
                Exit Sub
           End If
           rstK.MoveFirst
           For i = 0 To rstK.RecordCount - 1
             strCadena = "INSERT INTO almacen_producto(id_alm,id_producto,ruc) VALUES ('" & rstK("id_alm") & "','" & Trim(cproducto) & "','" & Trim(KEY_RUC) & "')"
             CnBd.Execute (strCadena)
             rstK.MoveNext
           Next i
           
           If Me.chkEmpresa.Value = 1 And Val(Me.DtcEmpresa.BoundText) > 0 Then
              Call save_grupo_empresarial(cproducto, Me.DtcEmpresa.BoundText)
           End If
           
         
           
           Call agrega_barra(cproducto)
           
           
  '         If KEY_RUC = "20603698852" Then
                If Val(Me.txtprecio.Text) > 0 Then
                Call producto_nuevo_keyfacil
                End If
   '        End If
           
           
           Call FrmProducto.actualizar_update(cproducto)
           If KEY_GRIFO = "si" Then
               strCadena = "call cursor_put_producto_tanque('" & Trim(cproducto) & "','" & KEY_RUC & "')"
               CnBd.Execute (strCadena)
           End If
           
           
           
           Unload Me
           
           Exit Sub
         
         
         Case modificar
          If Len(img) > 0 And Val(Me.LblCodigoProducto.Caption) > 0 And FlagFoto = True Then
            ' strCadena = "SELECT * FROM producto_foto WHERE id_producto='" & StrCodTabla & "' AND foto='" & Trim(img) & "' AND ruc='" & KEY_RUC & "'"
             'Call ConfiguraRstT(strCadena)
             'If rstT.RecordCount > 0 Then
              '   strCadena = "UPDATE producto_foto SET foto='" & img & "' WHERE id_produco='" & StrCodTabla & "' AND ruc='" & KEY_RUC & "'"
               '  Call Copiar_Archivo(PrtFoto, App.Path + "\archivos\" & KEY_RUC & "\" + img)
            'Else
               
                
               
            'End If
            
            
            
          End If
           
           If Me.DtcCategoria.Enabled = False Then
                strCadena = "UPDATE producto SET clave5='" & in_especialidad & "',icbper='" & IN_ICBPER & "',vencimiento='" & Format(Me.DtpVencimiento.Value, "YYYY-mm-dd") & "',forma_farmacologica='" & Trim(Me.txtFormafarmacologica.Text) & "',principio_activo='" & Trim(Me.TxtNombrecomercial.Text) & "',lote='" & Trim(Me.txtLote.Text) & "',registro_sanitario='" & Trim(Me.txtsanitario.Text) & "', codigo_alterno='" & Trim(Me.txtCodAlterno.Text) & "',codigo_proveedor='" & Trim(Me.TxtcodProveedor.Text) & "', id_universal='" & Trim(Me.txtCodigoUniversal.Text) & "',agranel='" & in_agranel & "',numero_procedimientos='" & Val(Me.txttrae.Text) & "',cta_contable='" & Trim(Me.txtcuenta_contable.Text) & "', id_unidad = '" & Me.DtcUnidad.BoundText & "',id_proveedor='" & Trim(Me.DtcProveedor.BoundText) & "',id_sub_producto='" & Trim(sub_producto) & "',id_linea='" & Me.DtcLinea.BoundText & "',id_sublinea='" & Me.DtaSublinea.BoundText & "'," & _
                "id_color='" & Me.DtcColor.BoundText & "',id_marca='" & Me.DtcMarca.BoundText & "',nombre_prod=  '" & Me.txtDescripcion.Text & "',nombre_comercial='" & Trim(Me.TxtNombrecomercial.Text) & "',stock_minimo=" & Me.TxtStockMinimo.Text & ",peso='" & Me.txtpeso.Text & "'," & _
                "id_percepcion='" & StrPercepcion & "',id_igv='" & strAfectoIGV & "',comentario='" & Me.txtObservacion.Text & "',id_combo='" & Trim(strCombo) & "'," & _
                " precio_delivery='" & Val(Me.TxtTransporte.Text) & "',principio_activo='" & Trim(Me.TxtNombrecomercial.Text) & "',id_tipo='" & Me.DtcTipoProducto.BoundText & "',partida_arancelaria='" & Trim(Me.TxtPartidaArancelaria.Text) & "',clave1='" & Trim(Me.txtclave1.Text) & "',clave2='" & Trim(Me.txtclave2.Text) & "',clave3='" & Trim(Me.txtclave3.Text) & "',clave4='" & Trim(Me.txtclave4.Text) & "' WHERE id_producto= '" & StrCodTabla & "' AND Ruc='" & KEY_RUC & "'"
           Else
                strCadena = "UPDATE producto SET clave5='" & Trim(in_especialidad) & "',icbper='" & IN_ICBPER & "',vencimiento='" & Format(Me.DtpVencimiento.Value, "YYYY-mm-dd") & "',forma_farmacologica='" & Trim(Me.txtFormafarmacologica.Text) & "',principio_activo='" & Trim(Me.TxtNombrecomercial.Text) & "',lote='" & Trim(Me.txtLote.Text) & "',registro_sanitario='" & Trim(Me.txtsanitario.Text) & "',codigo_alterno='" & Trim(Me.txtCodAlterno.Text) & "',codigo_proveedor='" & Trim(Me.TxtcodProveedor.Text) & "',id_universal='" & Trim(Me.txtCodigoUniversal.Text) & "',agranel='" & in_agranel & "',cta_contable='" & Trim(Me.txtcuenta_contable.Text) & "', numero_procedimientos='" & Val(Me.txttrae.Text) & "',id_unidad = '" & Me.DtcUnidad.BoundText & "',id_proveedor='" & Trim(Me.DtcProveedor.BoundText) & "',id_sub_producto='" & Trim(sub_producto) & "',id_linea='" & Me.DtcLinea.BoundText & "',id_sublinea='" & Me.DtaSublinea.BoundText & "'," & _
                "id_modelo='" & Me.DtcModelo.BoundText & "', id_color='" & Me.DtcColor.BoundText & "',id_marca='" & Me.DtcMarca.BoundText & "',nombre_prod=  '" & Me.txtDescripcion.Text & "',nombre_comercial='" & Trim(Me.TxtNombrecomercial.Text) & "',stock_minimo=" & Me.TxtStockMinimo.Text & ",peso='" & Me.txtpeso.Text & "'," & _
                "id_percepcion='" & StrPercepcion & "',id_igv='" & strAfectoIGV & "',comentario='" & Me.txtObservacion.Text & "',id_combo='" & Trim(strCombo) & "'," & _
                " precio_delivery='" & Val(Me.TxtTransporte.Text) & "',principio_activo='" & Trim(Me.TxtNombrecomercial.Text) & "',id_tipo='" & Me.DtcTipoProducto.BoundText & "',partida_arancelaria='" & Trim(Me.TxtPartidaArancelaria.Text) & "',clave1='" & Trim(Me.txtclave1.Text) & "',clave2='" & Trim(Me.txtclave2.Text) & "',clave3='" & Trim(Me.txtclave3.Text) & "',clave4='" & Trim(Me.txtclave4.Text) & "' WHERE id_producto= '" & StrCodTabla & "' AND Ruc='" & KEY_RUC & "'"
           End If
           
           CnBd.Execute (strCadena)
            
           
           strCadena = "UPDATE almacen_producto SET sector='" & Trim(Me.TxtSector.Text) & "',piso='" & Me.TxtPiso.Text & "',andamio='" & Me.TxtAndamio.Text & "',casillero_x='" & Me.txt_x.Text & "',casillero_y='" & Me.Txt_y.Text & "' WHERE id_producto='" & Me.LblCodigoProducto.Caption & "' AND id_alm='" & Me.DtcAlmacen.BoundText & "' AND ruc='" & KEY_RUC & "' "
           CnBd.Execute (strCadena)
            
           If Me.chkEmpresa.Value = 1 And Val(Me.DtcEmpresa.BoundText) > 0 Then
              Call save_grupo_empresarial(StrCodTabla, Me.DtcEmpresa.BoundText)
           End If
           
           
           
           
           
           
           If Trim(FrmProducto.HfdGrilla.TextMatrix(FrmProducto.HfdGrilla.Row, 0)) <> "" Then
               FrmProducto.HfdGrilla.TextMatrix(FrmProducto.HfdGrilla.Row, 1) = Trim(Me.txtDescripcion.Text)
           End If
                Call FrmProducto.actualizar_update(Trim(Me.LblCodigoProducto.Caption))
                
                If KEY_GRIFO = "si" Then
                    strCadena = "call cursor_put_producto_tanque('" & Trim(LblCodigoProducto.Caption) & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                End If
                
           Unload Me
           Exit Sub
       End Select
     End If
    'Call FrmProducto.actualizar
    'FrmProducto.ActualizarProd
    'FrmProducto.ActualizarAlm
End Sub
Sub verifica()

If Me.ChkPercepcion.Value = 1 Then
    StrPercepcion = "si"
Else
    StrPercepcion = "no"
End If

If Me.chkSubproductos.Value = 1 Then
    sub_producto = "si"
Else
    sub_producto = "no"
End If
    
    If (Me.chkigv.Value = 1) Then
        strAfectoIGV = "si"
    Else
        strAfectoIGV = "no"
    End If

End Sub

Private Sub HfgBarras_Click()
If Me.HfgBarras.Rows > 0 Then
    Me.CmdQuitar.Visible = True
Else
    Me.CmdQuitar.Visible = False
End If
End Sub

Private Sub HFTanques_DblClick()
Dim in_estado As String
If Val(Me.HfTanques.TextMatrix(Me.HfTanques.Row, 0)) > 0 Then
    If Me.HfTanques.TextMatrix(Me.HfTanques.Row, 2) = Chr(168) Then
        in_estado = "si"
        Me.HfTanques.TextMatrix(Me.HfTanques.Row, 2) = Chr(254)
    Else
        in_estado = "no"
        Me.HfTanques.TextMatrix(Me.HfTanques.Row, 2) = Chr(168)
    End If
    
    strCadena = "call put_producto_tanque('" & Val(Me.HfTanques.TextMatrix(Me.HfTanques.Row, 0)) & "','" & in_estado & "')"
    CnBd.Execute (strCadena)
    
End If
End Sub

Private Sub Image1_Click()
frmUnidadDetalle.Visible = False
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
      FlagFoto = False
      Exit Sub
    Case KEY_CANCEL
        Unload Me
  End Select
  Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  
  Exit Sub
End Sub

Private Sub TxtBuscaLinea_Change()
 strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' AND descripcion LIKE '%" & Trim(Me.TxtBuscaLinea.Text) & "%'  ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLinea)
End Sub

Private Sub TxtBuscaLinea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.DtcLinea.Enabled = True Then
       Me.DtcLinea.SetFocus
    End If
End If
End Sub

Private Sub TxtBuscamarca_Change()
 strCadena = "SELECT id_marca as Codigo, descripcion as Descripcion FROM marca WHERE id_usu='" & KEY_RUC & "' AND descripcion like '%" & Trim(Me.TxtBuscamarca.Text) & "%' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMarca)
End Sub

Private Sub TxtBuscamarca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Me.DtcMarca.Enabled = True Then
      Me.DtcMarca.SetFocus
      Exit Sub
   End If
End If
End Sub



Private Sub TxtBuscarAuspiciador_Change()

End Sub

Private Sub txtBuscarproveedor_Change()
strCadena = "SELECT E.cod_unico as Codigo, nombre_completo as Descripcion FROM persona P, entidad_empresa E WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' AND E.id_proveedor='si' AND nombre_completo LIKE '%" & Trim(Me.TxtBuscarProveedor.Text) & "%'  ORDER BY nombre_completo"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcProveedor)
End Sub

Private Sub txtBuscaSublinea_Change()
 strCadena = "SELECT id_tipo as Codigo, descripcion as Descripcion FROM linea_sub WHERE id_usu='" & KEY_RUC & "'  " & _
  " and descripcion LIKE '%" & Trim(Me.txtBuscaSublinea.Text) & "%' AND id_linea='" & Me.DtcLinea.BoundText & "' ORDER BY descripcion"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtaSublinea)
End Sub

Private Sub txtCategoria_Change()
strCadena = "SELECT id_categoria as Codigo,descripcion as Descripcion FROM categoria WHERE descripcion LIKE '%" & Trim(Me.Txtcategoria.Text) & "%'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcCategoria)
End Sub

Private Sub TxtCategoria1_Change()
strCadena = "SELECT id_categoria1 as Codigo,descripcion as Descripcion FROM categoria_1 WHERE descripcion LIKE '%" & Trim(Me.TxtCategoria1.Text) & "%' AND id_categoria='" & Me.DtcCategoria.BoundText & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcNivel1)

End Sub

Private Sub Txtcategoria2_Change()
strCadena = "SELECT id_categoria1 as Codigo,descripcion as Descripcion FROM categoria_2 WHERE descripcion LIKE '%" & Trim(Me.Txtcategoria2.Text) & "%' AND id_categoria1='" & Me.DtcNivel1.BoundText & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcNivel2)

End Sub

Private Sub TxtCodBarra_Change()
strCadena = "SELECT P.nombre_prod FROM producto_barras B,producto P WHERE B.id_producto=P.id_producto AND P.ruc='" & KEY_RUC & "' AND B.ruc='" & KEY_RUC & "'AND B.cod_barra = '" & Trim(Me.TxtCodBarra.Text) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.lblError.Visible = True
    Me.lblCabecera.Visible = True
    Me.lblError.Caption = Trim(rst(0))
    Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = False
Else
    Me.lblError.Visible = False
    Me.lblCabecera.Visible = False
    Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = True
End If
Set rst = Nothing
End Sub

Private Sub TxtCodBarra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtDescripcion)
End If
End Sub

Private Sub txtCodCompatible_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtCodCompatible.Text = formato_item(Me.txtCodCompatible.Text, 5)
    strCadena = "SELECT * FROM producto WHERE id_producto='" & Trim(Me.txtCodCompatible.Text) & "' AND ruc='" & KEY_RUC & "' AND id_producto<>'" & Trim(Me.txtCodCompatible.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.TxtCompatible.Text = rst("nombre_prod")
    Else
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
    End If
End If
End Sub

Private Sub txtCodigoUniversal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtcodProveedor)
End If

End Sub

Private Sub TxtcodProveedor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtDescripcion)
End If
End Sub

Private Sub txtcuenta_contable_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Procedencia = Selecionar
    FrmPlanContableCuentas.Show
    Exit Sub
End If
End Sub

Private Sub txtdescripcion_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
If KeyAscii = 13 Then
    If Len(Me.TxtNombrecomercial.Text) < 1 Then
        Me.TxtNombrecomercial.Text = Trim(Me.txtDescripcion.Text)
        Me.Txtproducto.Text = Trim(Me.txtDescripcion.Text)
        Me.TxtMarca.Text = Me.DtcMarca.Text
    End If
    Call Resalta(Me.TxtNombrecomercial)
End If
End Sub

Private Sub TxtMarca_Change()
  'If Trim(Me.TxtMarca.Text) <> "" Then
  'strCadena = "SELECT id_marca as Codigo, descripcion as Descripcion FROM marca WHERE id_usu='" & KEY_RUC & "' AND descripcion LIKE '%" & Trim(Me.TxtBuscamarca.Text) & "%' ORDER BY descripcion"
  'Call ConfiguraRst(strCadena)
  'Call LlenaDataCombo(Me.DtcMarca)
  'End If
End Sub

Private Sub TxtLinea_Change()
 strCadena = "SELECT id_linea as Codigo, descripcion as Descripcion FROM linea WHERE id_usu='" & KEY_RUC & "' AND descripcion LIKE '%" & Trim(Me.TxtBuscaLinea.Text) & "%' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcLinea)
End Sub

Private Sub txtModelo_Change()
  strCadena = "SELECT id as Codigo, descripcion as Descripcion FROM linea_modelo WHERE ruc='" & KEY_RUC & "' AND id_linea='" & Me.DtcLinea.BoundText & "' and id_sublinea='" & Me.DtaSublinea.BoundText & "' and descripcion LIKE '%" & Trim(Me.txtModelo.Text) & "%' " & _
  " ORDER BY descripcion"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtcModelo)
End Sub

Private Sub TxtNombrecomercial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    If Me.DtcMarca.Enabled = True Then
        Call Resalta(Me.TxtBuscamarca)
        Exit Sub
    End If

End If
End Sub

Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
Call Mayusculas(KeyAscii)
End Sub

Private Sub get_partida_arancelaria(ByVal in_partida As String)
strCadena = "SELECT * FROM partida_arancelaria WHERE codigo LIKE '%" & Trim(Me.TxtPartidaArancelaria.Text) & "%'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.lblPartida.Caption = rst("descripcion")
Else
    Me.lblPartida.Caption = ""
End If
End Sub

Private Sub TxtPartidaArancelaria_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call get_partida_arancelaria(Trim(Me.TxtPartidaArancelaria.Text))
End If
End Sub

Private Sub TxtPeso_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
        Call Resalta(Me.txtObservacion)
End If
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   ' Me.TxtPrecioCompra.SetFocus
End If
End Sub

Private Sub TxtReorden_KeyPress(KeyAscii As Integer)
  
End Sub




Private Sub TxtPrecioCompra_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
    
End If
End Sub

Private Sub TxtStockActual_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Me.TxtStockMinimo.SetFocus
End If
End Sub
Private Sub TxtStockMinimo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
    Me.txtpeso.SetFocus
End If
End Sub

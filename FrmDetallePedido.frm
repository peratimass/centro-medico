VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmDetallePedido 
   BorderStyle     =   0  'None
   Caption         =   "Detalle de los Pedido"
   ClientHeight    =   9240
   ClientLeft      =   405
   ClientTop       =   255
   ClientWidth     =   16920
   Icon            =   "FrmDetallePedido.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   16920
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtObservacion 
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
      Height          =   555
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   51
      Top             =   7680
      Width           =   4575
   End
   Begin VB.TextBox TxtSerie 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   13800
      MaxLength       =   50
      TabIndex        =   19
      Top             =   520
      Width           =   855
   End
   Begin VB.TextBox TxtNumero 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   14760
      MaxLength       =   50
      TabIndex        =   18
      Top             =   520
      Width           =   1455
   End
   Begin VB.TextBox TxtRuc 
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
      Height          =   285
      Left            =   2160
      TabIndex        =   16
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox TxtProveedor 
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
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Top             =   1440
      Width           =   5295
   End
   Begin VB.TextBox TxtId_pedido 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8160
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtcosto 
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
      Left            =   12105
      MaxLength       =   80
      TabIndex        =   12
      Top             =   7035
      Width           =   975
   End
   Begin VB.TextBox TxtUnidad 
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
      Left            =   10740
      MaxLength       =   80
      TabIndex        =   11
      Top             =   7035
      Width           =   1215
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
      Left            =   2550
      MaxLength       =   80
      TabIndex        =   10
      Top             =   7035
      Width           =   6735
   End
   Begin VB.TextBox TxtCantidad 
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
      Left            =   1500
      MaxLength       =   80
      TabIndex        =   9
      Top             =   7035
      Width           =   975
   End
   Begin VB.TextBox TxtCodProducto 
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
      MaxLength       =   80
      TabIndex        =   8
      Top             =   7035
      Width           =   1215
   End
   Begin VB.TextBox txtValorVenta 
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
      Left            =   9360
      TabIndex        =   6
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox TxtIgv 
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
      Left            =   9360
      TabIndex        =   5
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox txtImporteBruto 
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
      Left            =   9360
      TabIndex        =   4
      Top             =   7605
      Width           =   1455
   End
   Begin VB.TextBox TxtTotal 
      Appearance      =   0  'Flat
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
      Left            =   9360
      TabIndex        =   3
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox txtTc 
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
      Left            =   6600
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtId_recepcion 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8160
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxtDescuentoParcial 
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
      Left            =   9360
      MaxLength       =   80
      TabIndex        =   0
      ToolTipText     =   "Descuento"
      Top             =   7035
      Width           =   1215
   End
   Begin VitekeySoft.ChameleonBtn CmdAgregar 
      Height          =   315
      Left            =   13320
      TabIndex        =   7
      Top             =   7035
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "ADD"
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
      MICON           =   "FrmDetallePedido.frx":058A
      PICN            =   "FrmDetallePedido.frx":05A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcTerminosEntrega 
      Height          =   330
      Left            =   2160
      TabIndex        =   15
      Top             =   2175
      Width           =   5295
      _ExtentX        =   9340
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
   Begin MSComCtl2.DTPicker DtpPedido 
      Height          =   315
      Left            =   2160
      TabIndex        =   17
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
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
      Format          =   183697409
      CurrentDate     =   40974
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   3975
      Left            =   240
      TabIndex        =   20
      Top             =   3000
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   7011
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
   Begin MSComCtl2.DTPicker DtpPago 
      Height          =   315
      Left            =   5880
      TabIndex        =   21
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
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
      Format          =   183697409
      CurrentDate     =   40974
   End
   Begin VitekeySoft.ChameleonBtn cmdImprimir 
      Height          =   900
      Left            =   15080
      TabIndex        =   22
      Top             =   8265
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1588
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
      MICON           =   "FrmDetallePedido.frx":2AD1
      PICN            =   "FrmDetallePedido.frx":2AED
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdProcesar 
      Height          =   900
      Left            =   14200
      TabIndex        =   23
      Top             =   8280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1588
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
      MICON           =   "FrmDetallePedido.frx":50BE
      PICN            =   "FrmDetallePedido.frx":50DA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdCerrarpantalla 
      Height          =   900
      Left            =   15960
      TabIndex        =   24
      Top             =   8250
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1588
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
      MICON           =   "FrmDetallePedido.frx":8722
      PICN            =   "FrmDetallePedido.frx":873E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn CmdQuitar 
      Height          =   315
      Left            =   15000
      TabIndex        =   25
      Top             =   7035
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "DELL"
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
      MICON           =   "FrmDetallePedido.frx":B765
      PICN            =   "FrmDetallePedido.frx":B781
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcCreador 
      Height          =   330
      Left            =   1920
      TabIndex        =   26
      Top             =   8340
      Width           =   4575
      _ExtentX        =   8070
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
   Begin MSDataListLib.DataCombo DtcEstadoPedido 
      Height          =   330
      Left            =   1920
      TabIndex        =   27
      Top             =   8850
      Width           =   4575
      _ExtentX        =   8070
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
   Begin VitekeySoft.ChameleonBtn cmdUpdate 
      Height          =   315
      Left            =   14160
      TabIndex        =   28
      Top             =   7035
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "UPD"
      ENAB            =   0   'False
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
      MICON           =   "FrmDetallePedido.frx":BD1B
      PICN            =   "FrmDetallePedido.frx":BD37
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdNuevo 
      Height          =   900
      Left            =   13320
      TabIndex        =   29
      Top             =   8280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1588
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
      MICON           =   "FrmDetallePedido.frx":E08B
      PICN            =   "FrmDetallePedido.frx":E0A7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo DtcComrpobante 
      Height          =   345
      Left            =   10320
      TabIndex        =   30
      Top             =   520
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   345
      Left            =   10320
      TabIndex        =   31
      Top             =   140
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcMoneda 
      Height          =   330
      Left            =   4680
      TabIndex        =   32
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
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
   Begin VitekeySoft.ChameleonBtn cmdActualizarEstado 
      Height          =   315
      Left            =   6525
      TabIndex        =   52
      Top             =   8880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "ACTUALIZAR"
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
      MICON           =   "FrmDetallePedido.frx":E4F9
      PICN            =   "FrmDetallePedido.frx":E515
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label14 
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
      Left            =   615
      TabIndex        =   50
      Top             =   7800
      Width           =   1065
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MODULOS DE PEDIDOS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   600
      Left            =   11160
      TabIndex        =   49
      Top             =   1440
      Width           =   6945
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA PEDIDO:"
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
      TabIndex        =   48
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/PROVEEDOR:"
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
      TabIndex        =   47
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label lblempresa 
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA EMISION :"
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
      Height          =   240
      Left            =   960
      TabIndex        =   46
      Top             =   200
      Width           =   7665
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA PAGO :"
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
      Left            =   4155
      TabIndex        =   45
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TERMINOS DE ENTREGA:"
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
      TabIndex        =   44
      Top             =   2160
      Width           =   1635
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ELABORADO POR :"
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
      TabIndex        =   43
      Top             =   8340
      Width           =   1245
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESTADO PEDIDO:"
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
      Left            =   570
      TabIndex        =   42
      Top             =   8850
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR :"
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
      TabIndex        =   41
      Top             =   1440
      Width           =   915
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
      Height          =   315
      Left            =   15765
      TabIndex        =   40
      Top             =   7035
      Width           =   945
   End
   Begin VB.Label lblid_detalle 
      Height          =   255
      Left            =   14160
      TabIndex        =   39
      Top             =   7680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTE BRUTO :"
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
      Left            =   8055
      TabIndex        =   38
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGV :"
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
      Left            =   8925
      TabIndex        =   37
      Top             =   8400
      Width           =   345
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR VENTA :"
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
      Left            =   8265
      TabIndex        =   36
      Top             =   8040
      Width           =   1005
   End
   Begin VB.Label lblruc 
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA EMISION :"
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
      Height          =   240
      Left            =   960
      TabIndex        =   35
      Top             =   500
      Width           =   7665
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
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
      Left            =   8790
      TabIndex        =   34
      Top             =   8760
      Width           =   525
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
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
      Left            =   3945
      TabIndex        =   33
      Top             =   1080
      Width           =   705
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   920
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   80
      Width           =   8775
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   920
      Left            =   9960
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   6495
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   1770
      Left            =   7920
      Shape           =   4  'Rounded Rectangle
      Top             =   7395
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   9240
      Left            =   0
      Top             =   0
      Width           =   16920
   End
End
Attribute VB_Name = "FrmDetallePedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub put_asiento_contable(ByVal in_compra As Double)
strCadena = "SELECT * FROM orden_compra WHERE id_recepcion='" & Val(Me.txtId_recepcion.Text) & "' AND id_estado<>'3' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount = 1 Then
If KEY_CONTABILIDAD = "si" Then
             strCadena = "call p_insert_compra_emitido_ii('" & in_compra & "')"
             CnBd.Execute (strCadena)
             MsgBox "NUMERO DE VOUCHER GENERADO  : " & Trim(in_compra), vbInformation, KEY_VENDEDOR
End If
End If
End Sub

Private Sub Save()
Dim in_pedido As Double
Dim in_finalizado As String
Dim in_finalizado_confirmar As String
If Trim(Me.txtRuc.Text) = "" Then
   MsgBox "INGRESE UN PROVEEDOR PARA SU ORDEN", vbInformation, KEY_VENDEDOR
   Exit Sub
End If
Me.cmdProcesar.Enabled = False


If Trim(Me.txtRuc.Text) = "" Then
   MsgBox "INGRESE UN PROVEEDOR PARA SU ORDEN", vbInformation, KEY_VENDEDOR
   Exit Sub
End If


strCadena = "call put_pedido('" & Me.DtcComrpobante.BoundText & "','" & Trim(Me.txtserie.Text) & "','" & get_nueva_orden(DtcComrpobante.BoundText) & "','" & Trim(Me.txtRuc.Text) & "','" & KEY_FECHA & "', " & _
" '" & KEY_USUARIO & "','" & KEY_ALM & "','" & Val(Me.TxtTotal.Text) & "','" & Me.DtcMoneda.BoundText & "','" & Trim(Me.txtObservacion.Text) & "','" & KEY_RUC & "')"
Call ConfiguraRstP(strCadena)
in_pedido = rstP("in_pedido")
Me.TxtId_pedido.Text = in_pedido



strCadena = "SELECT * FROM movimiento_pedido_detalle_temp WHERE id_doc='" & Me.DtcComrpobante.BoundText & "'and id_alm='" & Me.DtcAlmacen.BoundText & "' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   rst.MoveFirst
   in_finalizado = "si"
   For i = 0 To rst.RecordCount - 1
        strCadena = "INSERT INTO movimiento_pedido_detalle(`id_pedido`,`id_producto`,`cantidad`,`cantidad_pendiente`,`precio`,`total`,`dni_save`,`id_alm`,`ruc`) VALUES " & _
        "('" & in_pedido & "','" & rst("id_producto") & "','" & rst("cantidad") & "','" & rst("cantidad_pendiente") & "','" & rst("precio") & "','" & rst("total") & "','" & KEY_USUARIO & "','" & KEY_ALM & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        rst.MoveNext
   Next i
End If



     


strCadena = "DELETE FROM movimiento_pedido_detalle_temp WHERE id_doc='" & Me.DtcComrpobante.BoundText & "' and  dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)


Me.cmdProcesar.Enabled = False
Me.cmdImprimir.Enabled = True

End Sub
Private Sub verifica_unica_recepcion(ByVal in_orden_compra As String, ByVal in_recepcion As String, ByVal in_estado As String)

If in_estado = "2" Then
strCadena = "SELECT * FROM orden_compra WHERE id_recepcion='" & Val(in_orden_compra) & "' and id_estado='2' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstP(strCadena)
    If rstP.RecordCount = 1 Then
        strCadena = "call CON_InsertaAsiento_Recepcion('" & in_recepcion & "')"
        CnBd.Execute (strCadena)
    End If
Else
    If in_estado = "4" And verificar_periodo_recepcion(in_recepcion) = True Then
        strCadena = "call CON_InsertaAsiento_Recepcion('" & in_recepcion & "')"
        CnBd.Execute (strCadena)
    End If
End If
End Sub
Private Function verificar_periodo_recepcion(ByVal in_recepcion As String) As Boolean
Dim in_fecha_recepcion As Date
Dim in_fecha_compra As Date
strCadena = "SELECT * FROM orden_compra WHERE id_orden='" & in_recepcion & "' and ruc='" & KEY_RUC & "' "
Call ConfiguraRstP(strCadena)
If rstP.RecordCount > 0 Then
    in_fecha_recepcion = rstP("fecha_solicitud")
   If rstP("id_compra") > 0 Then
        strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & rstP("id_compra") & "' and ruc='" & KEY_RUC & "' LIMIT 1"
        Call ConfiguraRstL(strCadena)
        If rstL.RecordCount > 0 Then
           in_fecha_compra = rstL("fecha_emision")
        End If
        
        If Month(in_fecha_recepcion) <> Month(in_fecha_compra) Then
            verificar_periodo_recepcion = True
        Else
            verificar_periodo_recepcion = False
        End If
        
        
   End If
End If


End Function














Private Sub cmdActualizarEstado_Click()

strCadena = "UPDATE movimiento_pedido SET id_estado='" & Me.DtcEstadoPedido.BoundText & "' WHERE id_pedido='" & Val(Me.TxtId_pedido.Text) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
MsgBox "Procesado con Exito", vbInformation
End Sub

Private Sub cmdagregar_Click()

    strCadena = "call orden_pedido_temporal('" & Trim(Me.TxtCodProducto.Text) & "','" & Trim(Me.txtCantidad.Text) & "','" & Val(Me.txtcosto.Text) & "','" & KEY_ALM & "','" & KEY_USUARIO & "','" & Val(Me.lblid_detalle.Caption) & "','" & Me.DtcComrpobante.BoundText & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
Me.TxtCodProducto.Text = ""
Me.TxtUnidad.Text = ""
Me.txtcosto.Text = ""
Me.txtCantidad.Text = "1"
Me.TxtDescripcionProducto.Text = ""
Me.lblid_detalle.Caption = 0
Call Me.llenar_orden(Me.HfdDetalle, Val(Me.TxtId_pedido.Text))
Call Resalta(Me.TxtCodProducto)
End Sub

Private Sub cmdCerrarpantalla_Click()

Unload Me
Exit Sub
End Sub

Private Sub cmdEditable_Click()

End Sub

Private Sub cmddelete_Click()

End Sub





Private Sub cmdImprimir_Click()

Dim arr(0 To 1, 1 To 2) As String
Dim param As Variant
Dim in_total As String

arr(0, 1) = "moneda_ini"
arr(1, 1) = "moneda_fin"

arr(0, 2) = Me.DtcMoneda.BoundText
If Val(Me.TxtTotal.Text) = 0 Then
    arr(1, 2) = "CERO CON 00/100 SOLES"
Else
    arr(1, 2) = UCase(EnLetras(Val(Me.TxtTotal.Text))) & Space(1) & Me.DtcMoneda.Text
End If


param = arr()

strCadena = "SELECT id_pedido,comprobante,fecha,id_proveedor,'" & KEY_EMPRESA & "','" & KEY_DIRECCION_ALM & "','" & Me.DtcAlmacen.Text & "',id_producto,nombre_prod,marca,unidad,cantidad,precio,total,nombre_completo FROM view_pedido_print WHERE id_pedido='" & Val(Me.TxtId_pedido.Text) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptOrdenPedido", param, App.Path + "\Reportes\")
    



End Sub

Private Sub put_delete(ByVal in_detalle As String)
strCadena = "DELETE FROM movimiento_pedido_detalle_temp WHERE id_detalle='" & Val(in_detalle) & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)

Call Me.llenar_orden(Me.HfdDetalle, Val(Me.TxtId_pedido.Text))
End Sub




Private Sub cmdNuevo_Click()

If MsgBox("Desea Limpiar esta Orden TEMPORAL", vbYesNo + vbQuestion, KEY_VENDEDOR) = vbYes Then
    strCadena = "DELETE FROM movimiento_pedido_detalle_temp WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
    Call Me.llenar_orden(Me.HfdDetalle, Val(Me.TxtId_pedido.Text))
    Me.TxtId_pedido.Text = ""
    Me.txtId_recepcion.Text = ""
    Me.txtValorVenta.Text = ""
    Me.TxtIgv.Text = ""
    Me.TxtTotal.Text = ""
    
    Me.cmdProcesar.Enabled = True
End If


End Sub

Private Sub cmdProcesar_Click()


If Format(Me.DtpPedido.Value, "YYYY-mm-dd") >= Format(get_fecha_periodo_abierto, "YYYY-mm-dd") Then
    Call Save
    Call FrmPedido.actualizar
Else
    MsgBox "PERIODO CERRADO COORDINE CON EL AREA CONTABLE", vbInformation
    Exit Sub
End If




End Sub



Private Sub CmdQuitar_Click()
If Val(Me.HfdDetalle.Rows) > 0 Then
Call put_delete(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))


    Call Me.llenar_orden(Me.HfdDetalle, Val(Me.TxtId_pedido.Text))




End If
End Sub

Private Sub cmdupdate_Click()

strCadena = "SELECT * FROM view_orden_pedido_detalle_temp WHERE ruc='" & KEY_RUC & "' and  id_detalle='" & Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
    Me.TxtCodProducto.Text = rstK("id_producto")
    Me.txtCantidad.Text = rstK("cantidad")
    Me.TxtDescripcionProducto.Text = rstK("nombre_prod")
    Me.TxtUnidad.Text = rstK("unidad")
    Me.txtcosto.Text = rstK("precio")
    Me.lblid_detalle.Caption = Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0))
    Me.cmdupdate.Enabled = False
End If



End Sub

Private Sub Command1_Click()
End Sub



Private Sub DtcComrpobante_Change()
Call get_comprobante(Me.DtcComrpobante.BoundText)

End Sub







Private Sub Form_Load()
CenterForm Me
Me.Top = 50
strCadena = "SELECT id_termino as Codigo,descripcion as Descripcion FROM terminos_entrega ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTerminosEntrega)


strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)
Me.DtcMoneda.BoundText = "00001"




strCadena = "SELECT id_alm as Codigo,descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' and id_tipoentidad='0' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcAlmacen.BoundText = KEY_ALM

strCadena = "SELECT id_doc as Codigo, doc_des as Descripcion FROM view_almacen_comprobante_ultimate WHERE id_doc='0103' and ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComrpobante)


strCadena = "SELECT id_estado as Codigo, descripcion as Descripcion FROM estado_pedido ORDER BY id_estado"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcEstadoPedido)

    
    
    

Me.lblruc.Caption = KEY_RUC
Me.LblEmpresa.Caption = KEY_EMPRESA


strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and id_personal='si'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcCreador)
Me.DtcCreador.BoundText = KEY_USUARIO


Me.DtpPedido.Value = KEY_FECHA
Me.DtpPago.Value = KEY_FECHA


Me.txtTc.Text = KEY_CAMBIO_VENTA



End Sub
Private Sub get_comprobante(ByVal in_doc As String)
strCadena = "SELECT * FROM movimiento_pedido WHERE id_doc='" & in_doc & "' and   ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
     Me.txtserie.Text = rst("serie")
     Me.txtNumero.Text = Format(Val(rst("numero")) + 1, "000000")
Else
    Me.txtserie.Text = "001"
    Me.txtNumero.Text = "000001"
End If
End Sub
Private Function get_nueva_orden(ByVal in_doc As String) As String
strCadena = "SELECT * FROM movimiento_pedido WHERE id_doc='" & in_doc & "' and   ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
Call ConfiguraRstL(strCadena)
If rstL.RecordCount > 0 Then
     get_nueva_orden = Format(Val(rstL("numero")) + 1, "000000")
Else
     get_nueva_orden = "000001"
End If

End Function
Private Function get_factura_flete(ByVal in_flete As String)
strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & Val(in_flete) & "' and ruc='" & KEY_RUC & "'"
Call ConfiguraRstZ(strCadena)
If rstZ.RecordCount > 0 Then
      get_factura_flete = rstZ("id_doc") & "-" & rstZ("serie") & "-" & rstZ("numero")
Else
    get_factura_flete = ""
End If
End Function
Public Sub get_orden(ByVal in_orden As String)
Dim in_afecto As String
Dim in_save As String
Dim in_observacion As String

strCadena = "SELECT * FROM movimiento_pedido WHERE id_pedido='" & Val(in_orden) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   
   Me.TxtId_pedido.Text = rst("id_pedido")
   Me.DtcComrpobante.BoundText = rst("id_doc")
   Me.txtserie.Text = rst("serie")
   Me.txtNumero.Text = rst("numero")
   Me.txtRuc.Text = rst("id_proveedor")
   Me.TxtProveedor.Text = get_persona(rst("id_proveedor"))
   Me.DtpPedido.Value = rst("fecha")
   in_save = rst("dni_save")
   in_observacion = rst("observacion")
   Me.DtcMoneda.BoundText = "00001"
   Me.txtRuc.Text = rst("id_proveedor")
   
   
   Me.LblEmpresa.Caption = get_persona(rst("id_proveedor"))
   Me.DtcEstadoPedido.BoundText = rst("id_estado")
   
   
   strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and dni='" & rst("dni_save") & "'"
   Call ConfiguraRstT(strCadena)
   Call LlenaDataComboT(Me.DtcCreador)
   Call llenar_orden(Me.HfdDetalle, Val(Me.TxtId_pedido.Text))
   Me.DtcCreador.BoundText = in_save
   Me.txtObservacion.Text = in_observacion
   
   If KEY_CARGO = "00004" Then
      Me.cmdActualizarEstado.Visible = True
   Else
      Me.cmdActualizarEstado.Visible = False
   End If
   
   
   Me.cmdProcesar.Enabled = False
End If

End Sub
Public Sub llenar_orden(ByVal Grilla As MSHFlexGrid, ByVal id_orden As Double)
'On Error GoTo salir
Dim tTotal As Double
If Val(id_orden) > 0 Then
    strCadena = "SELECT * FROM view_pedido_detalle WHERE id_pedido='" & id_orden & "' and ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT * FROM view_orden_pedido_detalle_temp WHERE dni_save='" & KEY_USUARIO & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "'"
End If

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub

End If
   
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 5500
           Grilla.ColWidth(3) = 2000
           Grilla.ColWidth(4) = 2200
           Grilla.ColWidth(5) = 1300
           Grilla.ColWidth(6) = 1800
           Grilla.ColWidth(7) = 2000
        Next
        cabecera = "IDDETALLE" & vbTab & "CODIGO" & vbTab & "PRODUCTO" & vbTab & "UNIDAD" & vbTab & "CLASIFICACION" & vbTab & "CANTIDAD" & vbTab & "PRECIO COSTO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 7
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
          Next k
        tTotal = 0
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            tTotal = tTotal + rst("total")
            Fila = rst("id_detalle") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("unidad") & vbTab & rst("linea") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("precio"), "#,##0.0000") & vbTab & Format(rst("total"), "#,##0.00")
            Grilla.AddItem Fila
            
            rst.MoveNext
        Next i
        
        cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "===============" & vbTab & Format(tTotal, "#,##0.00")
        Grilla.AddItem cabecera
        For k = 6 To 7
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
          Next k
          Me.LblCantidad.Caption = rst.RecordCount
          
          in_bruto = Val(tTotal)
    
          
          Me.txtImporteBruto.Text = Format(tTotal, "###0.00")
          Me.txtValorVenta.Text = Format(Val(tTotal), "###0.00")
          Me.TxtIgv.Text = Format(Val(Me.txtValorVenta.Text) * KEY_IGV, "###0.00")
          Me.TxtTotal.Text = Format(Val(Me.txtValorVenta.Text) + Val(Me.TxtIgv.Text), "###0.00")
         
          
'Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub

Public Sub nuevo_registro()
Me.TxtId_pedido.Text = 0
strCadena = "SELECT * FROM movimiento_pedido WHERE id_doc='" & DtcComrpobante.BoundText & "' and  ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.txtserie.Text = rst("serie")
   Me.txtNumero.Text = Format(Val(rst("numero")) + 1, "000000")
   
   
   
Else
   Me.txtserie.Text = "001"
   Me.txtNumero.Text = "000001"
End If
Me.DtcComrpobante.Enabled = True
strCadena = "SELECT dni as Codigo,nombre_completo as Descripcion FROM view_entidad WHERE ruc='" & KEY_RUC & "' and dni='" & KEY_USUARIO & "'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcCreador)

Call Me.llenar_orden(Me.HfdDetalle, Val(Me.TxtId_pedido.Text))
End Sub



Private Sub HfdDetalle_SelChange()
If Val(Me.HfdDetalle.TextMatrix(Me.HfdDetalle.Row, 0)) > 0 Then
    Me.cmdupdate.Enabled = True
Else
    Me.cmdupdate.Enabled = False
End If

End Sub



Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtDescuentoParcial)
End If
End Sub

Private Sub TxtcodigoProd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM producto where id_producto='" & Trim(MeTxtCodProducto.Text) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    
       Me.TxtCodProducto.Text = Trim(Me.TxtCodProducto.Text)
       Me.TxtDescripcionProducto.Text = rst("nombre_prod")
    Else
        Procedencia = buscar
        FrmProducto.Show
        Exit Sub
    End If
End If
End Sub

Private Sub TxtCodProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtCodProducto.Text = Format(Trim(Me.TxtCodProducto.Text), "00000")
    strCadena = "SELECT * FROM view_producto WHERE id_producto = '" & Trim(Me.TxtCodProducto.Text) & "' AND ruc='" & KEY_RUC & "' and id_alm='" & KEY_ALM & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.TxtCodProducto.Text = rst("id_producto")
        Me.TxtDescripcionProducto.Text = rst("nombre_prod")
        Me.TxtUnidad.Text = rst("unidad")
        Call Resalta(Me.txtCantidad)
 
 Else
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
 End If
End If
End Sub

Private Sub txtcosto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.cmdagregar.SetFocus
End If
End Sub




Private Sub TxtDescuentoParcial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtcosto)
End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
            Me.TxtProveedor.Text = (rst("nombre_completo"))
            Call Resalta(Me.TxtCodProducto)
   Else
        Procedencia = Selecionar
        FrmPersona.Show
        Exit Sub
   End If
End If
End Sub



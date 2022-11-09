VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmRegistroComprasList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   18615
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dtpFechaFind 
      Height          =   345
      Left            =   15000
      TabIndex        =   51
      Top             =   165
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   173473793
      CurrentDate     =   42366
   End
   Begin VitekeySoft.ChameleonBtn cmdbuscar 
      Height          =   350
      Left            =   11760
      TabIndex        =   50
      Top             =   135
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      BTYPE           =   4
      TX              =   "BUSCAR"
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
      MICON           =   "FrmRegistroComprasList.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtnumerofind 
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
      Left            =   10560
      TabIndex        =   49
      Top             =   165
      Width           =   1095
   End
   Begin VB.TextBox txtseriefind 
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
      Left            =   9795
      TabIndex        =   48
      Top             =   165
      Width           =   735
   End
   Begin MSDataListLib.DataCombo DtcTipocomprobante 
      Height          =   330
      Left            =   7080
      TabIndex        =   47
      Top             =   165
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
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
   Begin VB.TextBox txtRucEmpresa 
      Height          =   285
      Left            =   20880
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TxtValorcompra 
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
      Left            =   13950
      TabIndex        =   21
      Text            =   "0.00"
      Top             =   8040
      Width           =   975
   End
   Begin VB.TextBox txtIgv 
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
      Left            =   15960
      TabIndex        =   20
      Text            =   "0.00"
      Top             =   8040
      Width           =   735
   End
   Begin VB.TextBox txtTotal 
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
      Left            =   16725
      TabIndex        =   19
      Text            =   "0.00"
      Top             =   8040
      Width           =   1215
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
      Height          =   285
      Left            =   5715
      TabIndex        =   18
      Top             =   8040
      Width           =   2655
   End
   Begin VB.TextBox TxtRuc 
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
      Left            =   4320
      MaxLength       =   11
      TabIndex        =   17
      Top             =   8040
      Width           =   1335
   End
   Begin VB.TextBox TxtNumero 
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
      Left            =   3225
      TabIndex        =   16
      Top             =   8040
      Width           =   975
   End
   Begin VB.TextBox txtSerie 
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
      Left            =   2700
      TabIndex        =   15
      Top             =   8040
      Width           =   495
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
      Height          =   285
      Left            =   13305
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   8040
      Width           =   615
   End
   Begin VB.TextBox TxtPercepcion 
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
      Left            =   12660
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   8040
      Width           =   615
   End
   Begin VB.TextBox TxtTipoComprobante 
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
      Left            =   2175
      TabIndex        =   12
      Top             =   8040
      Width           =   495
   End
   Begin VB.CommandButton CmdCentroCostos 
      Height          =   375
      Left            =   18030
      Picture         =   "FrmRegistroComprasList.frx":001C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Centro de Costos"
      Top             =   7725
      Width           =   375
   End
   Begin VB.CommandButton CmdAgregar 
      Height          =   375
      Left            =   18045
      Picture         =   "FrmRegistroComprasList.frx":2E34
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Agregar Registros"
      Top             =   8100
      Width           =   375
   End
   Begin VB.TextBox TxtTipoCambio 
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
      Height          =   310
      Left            =   9555
      TabIndex        =   9
      Top             =   8040
      Width           =   525
   End
   Begin VB.TextBox TxtValorCompraNoAfecta 
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
      Left            =   14955
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   8040
      Width           =   975
   End
   Begin VB.TextBox TxtAnioDua 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      TabIndex        =   7
      Top             =   8040
      Width           =   1335
   End
   Begin VB.TextBox TxtAnio 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   19560
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   12720
      Top             =   6900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroComprasList.frx":5BCA
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroComprasList.frx":601E
            Key             =   "(caja_chica)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroComprasList.frx":94BC
            Key             =   "(Cerrar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroComprasList.frx":9A56
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroComprasList.frx":9FF0
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroComprasList.frx":A310
            Key             =   "(Anular)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroComprasList.frx":A62A
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroComprasList.frx":AA7E
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroComprasList.frx":AED2
            Key             =   "(RCompras)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroComprasList.frx":C644
            Key             =   "(RVentas)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   12303
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
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   5985
      Left            =   17640
      TabIndex        =   2
      Top             =   660
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   10557
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   5985
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   3
         Top             =   375
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1429
         ButtonWidth     =   1614
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   " Modificar"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Anular"
               Key             =   "(Anular)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Anular)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Actualizar"
               Key             =   "(Actualizar)"
               ImageKey        =   "(RCompras)"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo DtcMes 
      Height          =   315
      Left            =   13680
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
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
   Begin MSDataListLib.DataCombo DtcTipoCompra 
      Height          =   315
      Left            =   11280
      TabIndex        =   22
      Top             =   8040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
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
   Begin MSMask.MaskEdBox TxtFecha 
      Height          =   285
      Left            =   240
      TabIndex        =   23
      ToolTipText     =   "dd/mm/yyyy"
      Top             =   8040
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSDataListLib.DataCombo DtcFormaPago 
      Height          =   315
      Left            =   10110
      TabIndex        =   24
      Top             =   8040
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
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
   Begin MSDataListLib.DataCombo DtcMoneda 
      Height          =   315
      Left            =   8415
      TabIndex        =   25
      Top             =   8040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   4194304
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
   Begin MSMask.MaskEdBox TxtCancelacion 
      Height          =   285
      Left            =   1200
      TabIndex        =   26
      ToolTipText     =   "dd/mm/yyyy"
      Top             =   8040
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VitekeySoft.ChameleonBtn cmdbuscarfecha 
      Height          =   345
      Left            =   16440
      TabIndex        =   52
      Top             =   135
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      BTYPE           =   4
      TX              =   "BUSCAR"
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
      MICON           =   "FrmRegistroComprasList.frx":CA96
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
      Caption         =   " COMPROBANTE :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   46
      Top             =   195
      Width           =   1455
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "SERIE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   2715
      TabIndex        =   44
      Top             =   7760
      Width           =   435
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "NUMERO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   3360
      TabIndex        =   43
      Top             =   7760
      Width           =   645
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "RUC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4125
      TabIndex        =   42
      Top             =   7760
      Width           =   315
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   5715
      TabIndex        =   41
      Top             =   7760
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "TD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   2160
      TabIndex        =   40
      Top             =   7760
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   240
      TabIndex        =   39
      Top             =   7760
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "T.COMPRA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   11280
      TabIndex        =   38
      Top             =   7760
      Width           =   795
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   16725
      TabIndex        =   37
      Top             =   7760
      Width           =   480
   End
   Begin VB.Label lblvalorigv 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "IGV"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   15960
      TabIndex        =   36
      Top             =   7760
      Width           =   255
   End
   Begin VB.Label lblvalorcompra 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "B.IMP(Afect)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   13950
      TabIndex        =   35
      Top             =   7760
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "ISC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   13320
      TabIndex        =   34
      Top             =   7760
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "PERCE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   12660
      TabIndex        =   33
      Top             =   7760
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "F.PAGO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   10110
      TabIndex        =   32
      Top             =   7760
      Width           =   570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "MONEDA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   8400
      TabIndex        =   31
      Top             =   7760
      Width           =   645
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "T.C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   9555
      TabIndex        =   30
      Top             =   7760
      Width           =   255
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "B.IMP(Inafe)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   14955
      TabIndex        =   29
      Top             =   7760
      Width           =   930
   End
   Begin VB.Label lblaniodua 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "AÑO DUA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   4440
      TabIndex        =   28
      Top             =   7760
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00DFDFE0&
      BackStyle       =   0  'Transparent
      Caption         =   "F.CANCEL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   1200
      TabIndex        =   27
      Top             =   7760
      Width           =   735
   End
   Begin VB.Label lblMes 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Compras Mensual:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   360
      TabIndex        =   6
      Top             =   180
      Width           =   4755
   End
   Begin VB.Label lblEmpresa 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Ventas Mensual:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   18600
      TabIndex        =   5
      Top             =   180
      Width           =   8955
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   540
      Left            =   120
      Top             =   60
      Width           =   18375
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00DFDFE0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   120
      Top             =   7680
      Width           =   18375
   End
End
Attribute VB_Name = "FrmRegistroComprasList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim igv As String
Public Procedencia As Calculadora
Public Procede As EnumProcede
Public strModificar As Boolean
Private Sub cmdagregar_Click()
Dim fechaI As String
Dim cPersona As String
Dim Saldo As Single
Dim Tipo_cambio As Single, cod_identidad As String * 1
KEY_ANULADO = "no"
If IsDate(Me.TxtFecha.Text) = False Or Me.TxtTipoComprobante.Text = "" Or Me.txtserie.Text = "" Or Me.txtNumero.Text = "" Then
    MsgBox "Parametros Incorrectos, llene los casilleros obligatorios", vbInformation, "Corrija los valores"
    Exit Sub
End If
fechaI = Format(CVDate(Me.TxtFecha.Text), "YYYY-mm-dd")

strCadena = "SELECT * FROM movimiento_compra WHERE serie='" & Trim(Me.txtserie.Text) & "' and numero='" & Trim(Me.txtNumero.Text) & "' and ruc='" & KEY_RUC & "'"

   If (Trim(Me.DtcFormaPago.BoundText) = "0001") Then
        Saldo = 0
    Else
        Saldo = Val(Me.TxtTotal.Text)
   End If
   
   If (Me.DtcMoneda.BoundText) = "00001" Then
        Tipo_cambio = 0
    Else
        Procedencia = MTC
        Tipo_cambio = Format(Val(Me.TxtTipoCambio.Text), "###0.00")
        Me.TxtPercepcion.Text = Round(Val(Me.TxtPercepcion.Text) * Tipo_cambio, 2)
        Me.txtisc.Text = Round(Val(Me.txtisc.Text) * Tipo_cambio, 2)
        Me.TxtValorcompra.Text = Round(Val(Me.TxtValorcompra.Text) * Tipo_cambio, 2)
        Me.TxtValorCompraNoAfecta.Text = Round(Val(Me.TxtValorCompraNoAfecta.Text) * Tipo_cambio, 2)
        Me.TxtIgv.Text = Round(Val(Me.TxtIgv.Text) * Tipo_cambio, 2)
        Me.TxtTotal.Text = Round(Val(Me.TxtTotal.Text) * Tipo_cambio, 2)
        Procedencia = Mneutro
   End If
   
   
    If Trim(Me.cmdagregar.Caption) = "M" Then
    If Trim(Me.TxtTipoComprobante.Text) = "0003" Then
        Me.txtCliente.Text = ""
    End If
    
    '--------------
    
    strCadena = "UPDATE movimiento_compra set id_tipo_compra='" & Me.DtcTipoCompra.BoundText & "' WHERE id_compra='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "'"
    CnBd.Execute (strCadena)
     
    GoTo llenar:
    If Me.DtcTipoCompra.BoundText = "01" Then
      
    strCadena = "UPDATE movimiento_venta SET fecha_emision='" & Format(Me.TxtFecha.Text, "YYYY-mm-dd") & "',fecha_cancelacion='" & CVDate(Me.TxtCancelacion.Text) & "',mes='" & Me.DtcMes.BoundText & "',anio='" & Trim(Me.txtAnio.Text) & "',doc_cod='" & Trim(Me.TxtTipoComprobante.Text) & "',serie='" & Trim(Me.txtserie.Text) & "', " & _
    "numero='" & Trim(Me.txtNumero.Text) & "',RucCliente='" & Trim(Me.txtRuc.Text) & "',NombreCliente='" & Trim(Me.txtCliente.Text) & "',moneda='" & Trim(Me.DtcMoneda.BoundText) & "'," & _
    "tipo_compra='" & Me.DtcTipoCompra.BoundText & "',idFormaPago='" & Me.DtcFormaPago.BoundText & "',isc='" & Val(Me.txtisc.Text) & "',grav1='" & Val(Me.TxtValorcompra.Text) & "'," & _
    "igv1='" & Val(Me.TxtIgv.Text) & "',nograv='" & Val(Me.TxtValorCompraNoAfecta.Text) & "',total='" & Val(Me.TxtTotal.Text) & "',saldo='" & Val(Saldo) & "',tc='" & Val(Tipo_cambio) & "' WHERE codigounico='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 29)) & "'"
   End If
   If Me.DtcTipoCompra.BoundText = "02" Then
    strCadena = "UPDATE RegistroComprasDetalle SET fecha='" & CVDate(Me.TxtFecha.Text) & "',fecha_cancelacion='" & CVDate(Me.TxtCancelacion.Text) & "',mes='" & Me.DtcMes.BoundText & "',anio='" & Trim(Me.txtAnio.Text) & "',doc_cod='" & Trim(Me.TxtTipoComprobante.Text) & "',serie='" & Trim(Me.txtserie.Text) & "', " & _
    "numero='" & Trim(Me.txtNumero.Text) & "',RucCliente='" & Trim(Me.txtRuc.Text) & "',NombreCliente='" & Trim(Me.txtCliente.Text) & "',moneda='" & Trim(Me.DtcMoneda.BoundText) & "'," & _
    "tipo_compra='" & Me.DtcTipoCompra.BoundText & "',idFormaPago='" & Me.DtcFormaPago.BoundText & "',isc='" & Val(Me.txtisc.Text) & "',grav2='" & Val(Me.TxtValorcompra.Text) & "'," & _
    "igv2='" & Val(Me.TxtIgv.Text) & "',nograv='" & Val(Me.TxtValorCompraNoAfecta.Text) & "',total='" & Val(Me.TxtTotal.Text) & "',saldo='" & Val(Saldo) & "',tc='" & Val(Tipo_cambio) & "' WHERE codigounico='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 29)) & "'"
   End If
   If Me.DtcTipoCompra.BoundText = "03" Then
    strCadena = "UPDATE RegistroComprasDetalle SET fecha='" & CVDate(Me.TxtFecha.Text) & "',fecha_cancelacion='" & CVDate(Me.TxtCancelacion.Text) & "',mes='" & Me.DtcMes.BoundText & "',anio='" & Trim(Me.txtAnio.Text) & "',doc_cod='" & Trim(Me.TxtTipoComprobante.Text) & "',serie='" & Trim(Me.txtserie.Text) & "', " & _
    "numero='" & Trim(Me.txtNumero.Text) & "',RucCliente='" & Trim(Me.txtRuc.Text) & "',NombreCliente='" & Trim(Me.txtCliente.Text) & "',moneda='" & Trim(Me.DtcMoneda.BoundText) & "'," & _
    "tipo_compra='" & Me.DtcTipoCompra.BoundText & "',idFormaPago='" & Me.DtcFormaPago.BoundText & "',isc='" & Val(Me.txtisc.Text) & "',grav3='" & Val(Me.TxtValorcompra.Text) & "'," & _
    "igv3='" & Val(Me.TxtIgv.Text) & "',nograv='" & Val(Me.TxtValorCompraNoAfecta.Text) & "',total='" & Val(Me.TxtTotal.Text) & "',saldo='" & Val(Saldo) & "',tc='" & Val(Tipo_cambio) & "' WHERE codigounico='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 29)) & "'"
   End If
   If Me.DtcTipoCompra.BoundText = "04" Then
    'StrCadena = "INSERT INTO RegistroComprasDetalle(Ruc,fecha,fecha_cancelacion,mes,anio,doc_cod,serie,numero,RucCliente,NombreCliente,moneda,tipo_compra,percepcion,nograv,total,anulado,saldo,idFormaPago,tc) VALUES ('" & Trim(FrmRegistroCompras.TxtRuc.Text) & "'," & _
    "'" & fechaI & "','" & CVDate(Me.TxtCancelacion.Text) & "','" & Trim(Me.DtcMes.BoundText) & "','" & Trim(Me.TxtAnio.Text) & "','" & Trim(Me.TxtTipoComprobante.Text) & "','" & Trim(Me.txtSerie.Text) & "','" & Trim(Me.TxtNumero.Text) & "','" & Trim(Me.TxtRuc.Text) & "','" & Me.TxtCliente.Text & "'," & _
    "'" & Trim(Me.DtcMoneda.BoundText) & "','" & Trim(Me.DtcTipoCompra.BoundText) & "','" & Val(Me.TxtPercepcion.Text) & "','" & Val(Me.txtTotal.Text) & "','" & Val(Me.txtTotal.Text) & "','F','" & saldo & "','" & Trim(Me.DtcFormaPago.BoundText) & "','" & Tipo_cambio & "')"
    
    strCadena = "UPDATE RegistroComprasDetalle SET fecha='" & CVDate(Me.TxtFecha.Text) & "',fecha_cancelacion='" & CVDate(Me.TxtCancelacion.Text) & "',mes='" & Me.DtcMes.BoundText & "',anio='" & Trim(Me.txtAnio.Text) & "',doc_cod='" & Trim(Me.TxtTipoComprobante.Text) & "',serie='" & Trim(Me.txtserie.Text) & "', " & _
    "numero='" & Trim(Me.txtNumero.Text) & "',RucCliente='" & Trim(Me.txtRuc.Text) & "',NombreCliente='" & Trim(Me.txtCliente.Text) & "',moneda='" & Trim(Me.DtcMoneda.BoundText) & "'," & _
    "tipo_compra='" & Me.DtcTipoCompra.BoundText & "',idFormaPago='" & Me.DtcFormaPago.BoundText & "',percepcion='" & Val(Me.TxtPercepcion.Text) & "',nograv='" & Val(Me.TxtTotal.Text) & "',total='" & Val(Me.TxtTotal.Text) & "',saldo='" & Val(Saldo) & "',tc='" & Val(Tipo_cambio) & "' WHERE codigounico='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 29)) & "'"
   End If
    '--------------
    
    CnBd.Execute (strCadena)
     
    If KEY_ANULADO = "V" Then
        strCadena = "UPDATE RegistroComprasDetalle SET anulado='V',NombreCliente='A N U L A D O',RucCliente='',grav1='0',igv1='0',grav2='0',igv2='0',grav3='0',igv3='0',percepcion='0',saldo='0' WHERE codigounico='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 29)) & "'"
        CnBd.Execute (strCadena)
         
        KEY_ANULADO = "F"
        
        
   End If
   Procedencia = Neutro
    Call grabarcc(Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 29)))
    Call llenarGrid(Me.HfdPersona, Me.txtRucEmpresa.Text)
    Me.txtNumero.Text = ""
    Me.txtRuc.Text = ""
    Me.txtCliente.Text = ""
    Me.TxtValorcompra.Text = 0#
    Me.TxtIgv.Text = 0#
    Me.txtisc.Text = 0#
    Me.TxtPercepcion.Text = 0#
    Me.TxtValorCompraNoAfecta.Text = 0#
    Me.TxtTotal.Text = 0#
    Me.cmdagregar.Caption = ""
    Me.cmdagregar.BackColor = &H8000000F
    Call Resalta(Me.txtserie)
    strModificar = False
   Exit Sub
 End If
   
        If Trim(Me.TxtTipoComprobante.Text) = "0050" Then
            Me.txtRuc.Text = "20131312955"
            Me.txtCliente.Text = "SUPERINTENDENCIA NACIONAL DE ADUANAS Y DE ADMINISTRACION TRIBUTARIA - SUNAT"
        
            
        End If
        If Len(Trim(Me.txtRuc.Text)) = 8 Then
            cod_identidad = 1
        End If
        If Len(Trim(Me.txtRuc.Text)) = 11 Then
            cod_identidad = 6
        End If
       
        
        strCadena = "P_insert_compra('" & Me.TxtTipoComprobante.Text & "','" & KEY_ALM & "','" & fechaI & "','" & Format(CVDate(Me.TxtCancelacion.Text), "YYYY-mm-dd") & "','02'," & _
        "'" & Me.DtcTipoCompra.BoundText & "','','" & Me.DtcMoneda.BoundText & "','" & formato_item(Me.DtcMes.BoundText, 2) & "','" & Trim(Me.txtAnio.Text) & "','" & Trim(Me.txtserie.Text) & "'," & _
        "'" & Trim(Me.txtNumero.Text) & "','" & cod_identidad & "','" & Trim(Me.txtRuc.Text) & "','" & UCase(Me.txtCliente.Text) & "','" & Val(Me.TxtTipoCambio.Text) & "'," & _
        "'0','" & Val(Me.TxtValorcompra.Text) & "','" & Val(Me.TxtIgv.Text) & "','" & Val(Me.txtisc.Text) & "','0','" & Val(Me.TxtPercepcion.Text) & "','0','" & Val(Me.TxtValorCompraNoAfecta.Text) & "','0','" & Val(Me.TxtTotal.Text) & "','" & Val(Me.TxtTotal.Text) & "','" & KEY_USUARIO & "','-','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
         
        
        id_compra = LastRegistroRUC("movimiento_compra", "id_compra")
         'MsgBox "REGISTRO UNICO:" + Str(rst(0)), vbInformation, "Numero de Registro"
   If KEY_ANULADO = "si" Then
        strCadena = "UPDATE movimiento_compra SET anulado='si',nproveedor='A N U L A D O',id_proveedor='',valor_venta='0',igv='0',isc='0',ivap='0',percepcion='0',retencion='0',exonerado='0',otros='0',total=0,saldo='0',isc='0',percepcion='0',otros='0' WHERE id_compra='" & Val(rst(0)) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
       KEY_ANULADO = "F"
   End If
llenar:
   Call grabarcc(id_compra)
   Call llenarGrid(Me.HfdPersona, Me.txtRucEmpresa.Text)
   Call nuevo

End Sub
Private Sub grabarcc(ByVal codigo As Double)
Dim ccosto As String

    If (cNaturaleza <> "") Then
        If strModificar = False Then
            strCadena = "INSERT INTO registrocomprascostosnaturaleza (codigounico,cnaturaleza)VALUES('" & Val(codigo) & "','" & Trim(cNaturaleza) & "')"
        Else
            strCadena = "UPDATE registrocomprascostosnaturaleza SET cnaturaleza='" & Trim(cNaturaleza) & "' WHERE codigounico='" & codigo & "'"
        End If
        CnBd.Execute (strCadena)
         
   End If
        If (ccostos1 <> "") Then
            If strModificar = False Then
                strCadena = "INSERT INTO registrocomprasnaturaleza_costos(codigounico,ccostos,monto)VALUES('" & Val(codigo) & "','" & Trim(ccostos1) & "','" & Val(cMonto1) & "') "
            Else
                strCadena = "SELECT * FROM registrocomprasnaturaleza_costos WHERE ccostos='" & Trim(ccostos1) & "' AND codigounico='" & Val(codigo) & "' AND ccostos='" & Trim(ccostos1G) & "'"
                ConfiguraTemporal (strCadena)
                If rstTemporal.RecordCount < 1 Then
                strCadena = "INSERT INTO registrocomprasnaturaleza_costos(codigounico,ccostos,monto)VALUES('" & Val(codigo) & "','" & Trim(ccostos1) & "','" & Val(cMonto1) & "') "
                Else
                strCadena = "UPDATE registrocomprasnaturaleza_costos SET ccostos='" & Trim(ccostos1) & "',monto='" & Val(cMonto1) & "' WHERE codigounico='" & Val(codigo) & "' AND ccostos='" & ccostos1G & "' "
                End If
            End If
            CnBd.Execute (strCadena)
             
            ccostos1 = ""
        End If
        If (ccostos2 <> "") Then
            If strModificar = False Then
                strCadena = "INSERT INTO registrocomprasnaturaleza_costos(codigounico,ccostos,monto)VALUES('" & Val(codigo) & "','" & Trim(ccostos2) & "','" & Val(cMonto2) & "') "
            Else
                strCadena = "SELECT * FROM registrocomprasnaturaleza_costos WHERE ccostos='" & Trim(ccostos2) & "' AND codigounico='" & Val(codigo) & "'"
                ConfiguraTemporal (strCadena)
                If rstTemporal.RecordCount < 1 Then
                strCadena = "INSERT INTO registrocomprasnaturaleza_costos(codigounico,ccostos,monto)VALUES('" & Val(codigo) & "','" & Trim(ccostos2) & "','" & Val(cMonto2) & "') "
                Else
                strCadena = "UPDATE registrocomprasnaturaleza_costos SET ccostos='" & Trim(ccostos2) & "',monto='" & Val(cMonto2) & "' WHERE codigounico='" & Val(codigo) & "' AND ccostos='" & ccostos2G & "' "
                End If
            End If
            CnBd.Execute (strCadena)
             
            ccostos2 = ""
        End If
        If (ccostos3 <> "") Then
            
            If strModificar = False Then
                strCadena = "INSERT INTO registrocomprasnaturaleza_costos(codigounico,ccostos,monto)VALUES('" & Val(codigo) & "','" & Trim(ccostos3) & "','" & Val(cMonto3) & "') "
            Else
                strCadena = "SELECT * FROM registrocomprasnaturaleza_costos WHERE ccostos='" & Trim(ccostos3) & "' AND codigounico='" & Val(codigo) & "'"
                ConfiguraTemporal (strCadena)
                If rstTemporal.RecordCount < 1 Then
                strCadena = "INSERT INTO RegistroComprasNaturaleza_Costos(codigounico,ccostos,monto)VALUES('" & Val(codigo) & "','" & Trim(ccostos3) & "','" & Val(cMonto3) & "') "
                Else
                strCadena = "UPDATE RegistroComprasNaturaleza_Costos SET ccostos='" & Trim(ccostos3) & "',monto='" & Val(cMonto3) & "' WHERE codigounico='" & Val(codigo) & "' AND ccostos='" & ccostos3G & "' "
                End If
            End If
            CnBd.Execute (strCadena)
             
            ccostos3 = ""
        End If
        If (ccostos4 <> "") Then
            If strModificar = False Then
                strCadena = "INSERT INTO registrocomprasnaturaleza_costos(codigounico,ccostos,monto)VALUES('" & Val(codigo) & "','" & Trim(ccostos4) & "','" & Val(cMonto4) & "') "
            Else
                
                strCadena = "SELECT * FROM registrocomprasnaturaleza_costos WHERE ccostos='" & Trim(ccostos4) & "' AND codigounico='" & Val(codigo) & "'"
                ConfiguraTemporal (strCadena)
                If rstTemporal.RecordCount < 1 Then
                strCadena = "INSERT INTO registrocomprasnaturaleza_costos(codigounico,ccostos,monto)VALUES('" & Val(codigo) & "','" & Trim(ccostos4) & "','" & Val(cMonto4) & "') "
                Else
                strCadena = "UPDATE registrocomprasnaturaleza_costos SET ccostos='" & Trim(ccostos4) & "',monto='" & Val(cMonto4) & "' WHERE codigounico='" & Val(codigo) & "' AND ccostos='" & ccostos4G & "' "
                End If
            End If
            CnBd.Execute (strCadena)
             
            ccostos4 = ""
        End If
    
End Sub

Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal ruc As String)
On Error GoTo salir
Dim tafecto As Double, texonerado As Double, tigv As Double, tTotal As Double, tpercepcion As Double, Totros As Double
Dim grav1 As Double, grav2 As Double, grav3 As Double, igv1 As Double, igv2 As Double, igv3 As Double
Dim in_periodo As String

tgrav1 = 0
tigv1 = 0
tgrav2 = 0
tigv2 = 0
tgrav3 = 0
tigv3 = 0
tngrav = 0
tTotal = 0
Totros = 0
tpercepcion = 0

strCadena = "SELECT * FROM con_periodo WHERE Mes='" & Me.DtcMes.BoundText & "' and Ejercicio='" & Val(Me.txtAnio.Text) & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   in_periodo = rst("Id")
End If


strCadena = "SELECT * FROM movimiento_compra WHERE ruc='" & Trim(ruc) & "' AND id_periodo='" & in_periodo & "'  and (id_doc<>'0089') ORDER BY fecha_emision,id_doc,serie,numero ASC"
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
           Grilla.ColWidth(1) = 950
           Grilla.ColWidth(2) = 950
           Grilla.ColWidth(3) = 500
           Grilla.ColWidth(4) = 800
           Grilla.ColWidth(5) = 500
           Grilla.ColWidth(6) = 800
           Grilla.ColWidth(7) = 1100
           Grilla.ColWidth(8) = 2800
           Grilla.ColWidth(9) = 1000
           Grilla.ColWidth(10) = 850
           Grilla.ColWidth(11) = 1000
           Grilla.ColWidth(12) = 850
           Grilla.ColWidth(13) = 1000
           Grilla.ColWidth(14) = 850
           Grilla.ColWidth(15) = 850
           Grilla.ColWidth(16) = 600
           Grilla.ColWidth(17) = 600
           Grilla.ColWidth(18) = 600
           Grilla.ColWidth(19) = 1100
           Grilla.ColWidth(20) = 1100
           Grilla.ColWidth(21) = 800
           Grilla.ColWidth(22) = 800
           Grilla.ColWidth(23) = 950
           Grilla.ColWidth(24) = 500
           Grilla.ColWidth(25) = 950
           Grilla.ColWidth(26) = 600
           Grilla.ColWidth(27) = 800
           Grilla.ColWidth(28) = 800
           
        Next
          
        rst.MoveFirst
        cabecera = "IDECOMPRA" & vbTab & "EMISION" & vbTab & "CANCELA" & vbTab & "TD" & vbTab & "AÑO DUA" & vbTab & "SERIE" & vbTab & "NUMERO" & vbTab & "RUC" & vbTab & "PROVEEDOR" & vbTab & "GRAV1" & vbTab & "IGV1" & vbTab & "GRAV2" & vbTab & "IGV2" & vbTab & "GRAV3" & vbTab & "IGV3" & vbTab & "NO GRAV" & vbTab & "ISC" & vbTab & "PERCE" & vbTab & "OTRO" & vbTab & "TOTAL" & vbTab & "TOTAL RUS" & vbTab & "N NDOM" & vbTab & "DETRAC" & vbTab & "EMISION" & vbTab & "TC" & vbTab & "FECHA" & vbTab & "TD" & vbTab & "SERIE" & vbTab & "NUMERO"
        Grilla.AddItem cabecera
         For k = 0 To 28
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        For i = 0 To rst.RecordCount - 1
            If (rst("id_doc") = "0003") Then
                trus = rst("total")
            Else
               trus = 0
            End If
            
            If IsDate(rst("fecha_cancelacion")) = True Then
                fcancelacion = rst("fecha_cancelacion")
            Else
                fcancelacion = rst("fecha_emision")
            End If
            If rst("id_tipo_compra") = "01" Then
                   grav1 = rst("valor_venta")
                   igv1 = rst("igv")
                   grav2 = 0
                   igv2 = 0
                   grav3 = 0
                   igv3 = 0
            End If
            If rst("id_tipo_compra") = "02" Then
                   grav2 = rst("valor_venta")
                   igv2 = rst("igv")
                   grav1 = 0
                   igv1 = 0
                   grav3 = 0
                   igv3 = 0
            End If
            If rst("id_tipo_compra") = "03" Then
                   grav3 = rst("valor_venta")
                   igv3 = rst("igv")
                   grav2 = 0
                   igv2 = 0
                   grav1 = 0
                   igv1 = 0
            End If
            
            Fila = rst("id_compra") & vbTab & rst("fecha_emision") & vbTab & fcancelacion & vbTab & rst("id_doc") & vbTab & rst("anio_dua") & vbTab & rst("serie") & vbTab & rst("numero") & vbTab & rst("id_proveedor") & vbTab & rst("nproveedor") & vbTab & Format(grav1, "###0.00") & vbTab & Format(igv1, "###0.00") & vbTab & Format(grav2, "###0.00") & vbTab & Format(igv2, "###0.00") & vbTab & Format(grav3, "###0.00") & vbTab & Format(igv3, "###0.00") & vbTab & Format(rst("exonerado"), "#,##0.00") & vbTab & Format(rst("isc"), "#,##0.00") & vbTab & Format(rst("percepcion"), "###0.00") & vbTab & Format(rst("otros"), "###0.00") & vbTab & Format(rst("total"), "###0.00") & vbTab & Format(trus, "###0.00") & vbTab & rst("numero_no_domiciliado") & vbTab & rst("numero_detrac") & vbTab & rst("fecha_detrac") & vbTab & rst("tc") & vbTab & rst("fecha_fact") & vbTab & rst("id_doc_fact") & vbTab & rst("serie_fact") & vbTab & rst("numero_fact") & vbTab & rst("id_compra")
            Grilla.AddItem Fila
                    If (Trim(rst("anulado")) = "si") Then
                            For k = 0 To 25
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                            Next k
                        Else
                        tgrav1 = tgrav1 + grav1
                        tigv1 = tigv1 + igv1
                        tgrav2 = tgrav2 + grav2
                        tigv2 = tigv2 + igv2
                        tgrav3 = tgrav3 + grav3
                        tigv3 = tigv3 + igv3
                        tngrav = tngrav + rst("exonerado")
                        Tisc = Tisc + rst("isc")
                        tpercepcion = tpercepcion + rst("percepcion")
                        Totros = Totros + rst("otros")
                        tTotal = tTotal + rst("total")
                        totrus = totrus + trus
                        End If
                        
                   
            
            
            Fila = ""
            rst.MoveNext
             
        Next i
     cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(tgrav1, "#,##0.00") & vbTab & Format(tigv1, "#,##0.00") & vbTab & Format(tgrav2, "#,##0.00") & vbTab & Format(tigv2, "#,##0.00") & vbTab & Format(tgrav3, "#,##0.00") & vbTab & Format(tigv3, "#,##0.00") & vbTab & Format(tngrav, "#,##0.00") & vbTab & Format(Tisc, "#,##0.00") & vbTab & Format(tpercepcion, "#,##0.00") & vbTab & Format(Totros, "#,##0.00") & vbTab & Format(tTotal, "#,##0.00") & vbTab & Format(totrus, "#,##0.00") & vbTab & ""
     Grilla.AddItem cabecera
     For k = 9 To 20
         Grilla.col = k
         Grilla.Row = i + 1
         Grilla.CellBackColor = &HC0FFFF
     Next k
    
    
    
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Public Sub llenarGridFind(ByVal Grilla As MSHFlexGrid, ByVal ruc As String)
On Error GoTo salir
Dim tafecto As Double, texonerado As Double, tigv As Double, tTotal As Double, tpercepcion As Double, Totros As Double
Dim grav1 As Double, grav2 As Double, grav3 As Double, igv1 As Double, igv2 As Double, igv3 As Double
tgrav1 = 0
tigv1 = 0
tgrav2 = 0
tigv2 = 0
tgrav3 = 0
tigv3 = 0
tngrav = 0
tTotal = 0
Totros = 0
tpercepcion = 0


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
           Grilla.ColWidth(1) = 950
           Grilla.ColWidth(2) = 950
           Grilla.ColWidth(3) = 500
           Grilla.ColWidth(4) = 800
           Grilla.ColWidth(5) = 500
           Grilla.ColWidth(6) = 800
           Grilla.ColWidth(7) = 1100
           Grilla.ColWidth(8) = 2800
           Grilla.ColWidth(9) = 1000
           Grilla.ColWidth(10) = 850
           Grilla.ColWidth(11) = 1000
           Grilla.ColWidth(12) = 850
           Grilla.ColWidth(13) = 1000
           Grilla.ColWidth(14) = 850
           Grilla.ColWidth(15) = 850
           Grilla.ColWidth(16) = 600
           Grilla.ColWidth(17) = 600
           Grilla.ColWidth(18) = 600
           Grilla.ColWidth(19) = 1100
           Grilla.ColWidth(20) = 1100
           Grilla.ColWidth(21) = 800
           Grilla.ColWidth(22) = 800
           Grilla.ColWidth(23) = 950
           Grilla.ColWidth(24) = 500
           Grilla.ColWidth(25) = 950
           Grilla.ColWidth(26) = 600
           Grilla.ColWidth(27) = 800
           Grilla.ColWidth(28) = 800
           
        Next
          
        rst.MoveFirst
        cabecera = "IDECOMPRA" & vbTab & "EMISION" & vbTab & "CANCELA" & vbTab & "TD" & vbTab & "AÑO DUA" & vbTab & "SERIE" & vbTab & "NUMERO" & vbTab & "RUC" & vbTab & "PROVEEDOR" & vbTab & "GRAV1" & vbTab & "IGV1" & vbTab & "GRAV2" & vbTab & "IGV2" & vbTab & "GRAV3" & vbTab & "IGV3" & vbTab & "NO GRAV" & vbTab & "ISC" & vbTab & "PERCE" & vbTab & "OTRO" & vbTab & "TOTAL" & vbTab & "TOTAL RUS" & vbTab & "N NDOM" & vbTab & "DETRAC" & vbTab & "EMISION" & vbTab & "TC" & vbTab & "FECHA" & vbTab & "TD" & vbTab & "SERIE" & vbTab & "NUMERO"
        Grilla.AddItem cabecera
         For k = 0 To 28
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        For i = 0 To rst.RecordCount - 1
            If (rst("id_doc") = "0003") Then
                trus = rst("total")
            Else
               trus = 0
            End If
            
            If IsDate(rst("fecha_cancelacion")) = True Then
                fcancelacion = rst("fecha_cancelacion")
            Else
                fcancelacion = rst("fecha_emision")
            End If
            If rst("id_tipo_compra") = "01" Then
                   grav1 = rst("valor_venta")
                   igv1 = rst("igv")
                   grav2 = 0
                   igv2 = 0
                   grav3 = 0
                   igv3 = 0
            End If
            If rst("id_tipo_compra") = "02" Then
                   grav2 = rst("valor_venta")
                   igv2 = rst("igv")
                   grav1 = 0
                   igv1 = 0
                   grav3 = 0
                   igv3 = 0
            End If
            If rst("id_tipo_compra") = "03" Then
                   grav3 = rst("valor_venta")
                   igv3 = rst("igv")
                   grav2 = 0
                   igv2 = 0
                   grav1 = 0
                   igv1 = 0
            End If
            
            Fila = rst("id_compra") & vbTab & rst("fecha_emision") & vbTab & fcancelacion & vbTab & rst("id_doc") & vbTab & rst("anio_dua") & vbTab & rst("serie") & vbTab & rst("numero") & vbTab & rst("id_proveedor") & vbTab & rst("nproveedor") & vbTab & Format(grav1, "###0.00") & vbTab & Format(igv1, "###0.00") & vbTab & Format(grav2, "###0.00") & vbTab & Format(igv2, "###0.00") & vbTab & Format(grav3, "###0.00") & vbTab & Format(igv3, "###0.00") & vbTab & Format(rst("exonerado"), "#,##0.00") & vbTab & Format(rst("isc"), "#,##0.00") & vbTab & Format(rst("percepcion"), "###0.00") & vbTab & Format(rst("otros"), "###0.00") & vbTab & Format(rst("total"), "###0.00") & vbTab & Format(trus, "###0.00") & vbTab & rst("numero_no_domiciliado") & vbTab & rst("numero_detrac") & vbTab & rst("fecha_detrac") & vbTab & rst("tc") & vbTab & rst("fecha_fact") & vbTab & rst("id_doc_fact") & vbTab & rst("serie_fact") & vbTab & rst("numero_fact") & vbTab & rst("id_compra")
            Grilla.AddItem Fila
                    If (Trim(rst("anulado")) = "si") Then
                            For k = 0 To 25
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                            Next k
                        Else
                        tgrav1 = tgrav1 + grav1
                        tigv1 = tigv1 + igv1
                        tgrav2 = tgrav2 + grav2
                        tigv2 = tigv2 + igv2
                        tgrav3 = tgrav3 + grav3
                        tigv3 = tigv3 + igv3
                        tngrav = tngrav + rst("exonerado")
                        Tisc = Tisc + rst("isc")
                        tpercepcion = tpercepcion + rst("percepcion")
                        Totros = Totros + rst("otros")
                        tTotal = tTotal + rst("total")
                        totrus = totrus + trus
                        End If
                        
                   
            
            
            Fila = ""
            rst.MoveNext
             
        Next i
     cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & Format(tgrav1, "#,##0.00") & vbTab & Format(tigv1, "#,##0.00") & vbTab & Format(tgrav2, "#,##0.00") & vbTab & Format(tigv2, "#,##0.00") & vbTab & Format(tgrav3, "#,##0.00") & vbTab & Format(tigv3, "#,##0.00") & vbTab & Format(tngrav, "#,##0.00") & vbTab & Format(Tisc, "#,##0.00") & vbTab & Format(tpercepcion, "#,##0.00") & vbTab & Format(Totros, "#,##0.00") & vbTab & Format(tTotal, "#,##0.00") & vbTab & Format(totrus, "#,##0.00") & vbTab & ""
     Grilla.AddItem cabecera
     For k = 9 To 20
         Grilla.col = k
         Grilla.Row = i + 1
         Grilla.CellBackColor = &HC0FFFF
     Next k
    
    
    
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub
Private Sub nuevo()
Me.TxtPercepcion.Text = 0#
Me.txtisc.Text = 0#
Me.TxtValorcompra.Text = 0#
Me.TxtValorCompraNoAfecta.Text = 0#
Me.TxtIgv.Text = 0#
Me.TxtTotal.Text = 0#
Me.txtRuc.Text = ""
Me.txtCliente.Text = ""
Me.TxtTipoComprobante.Text = ""
Me.txtserie.Text = ""
Me.txtNumero.Text = ""
Me.DtcFormaPago.BoundText = ""
 
Me.TxtFecha.Mask = ""
Me.TxtFecha.Text = ""
Me.TxtFecha.Mask = "##/##/####"
Me.TxtCancelacion.Mask = ""
Me.TxtCancelacion.Text = ""
Me.TxtCancelacion.Mask = "##/##/####"
Me.TxtFecha.SetFocus
'Me.txtRuc.SetFocus
'Call Resalta(Me.txtt)
End Sub
Private Sub DtComprobante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtserie)
End If
End Sub
Public Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub

Private Sub cmdBuscar_Click()
strCadena = "SELECT * FROM movimiento_compra WHERE ruc='" & Trim(KEY_RUC) & "' AND id_mes='" & Me.DtcMes.BoundText & "' AND id_anio='" & Trim(Me.txtAnio.Text) & "' and id_doc='" & Trim(Me.DtcTipoComprobante.BoundText) & "'  and serie='" & Trim(Me.txtseriefind.Text) & "' and numero LIKE '%" & Trim(Me.txtnumerofind.Text) & "%' ORDER BY fecha_emision,id_doc,serie,numero ASC"
Call Me.llenarGridFind(Me.HfdPersona, KEY_RUC)
End Sub

Private Sub CmdBuscarFecha_Click()
strCadena = "SELECT * FROM movimiento_compra WHERE ruc='" & Trim(KEY_RUC) & "' AND id_mes='" & Me.DtcMes.BoundText & "' AND id_anio='" & Trim(Me.txtAnio.Text) & "' and fecha_emision='" & Format(Me.dtpFechaFind.Value, "YYYY-mm-dd") & "'  ORDER BY fecha_emision,id_doc,serie,numero ASC"
Call Me.llenarGridFind(Me.HfdPersona, KEY_RUC)

End Sub

Private Sub CmdCentroCostos_Click()
FrmCCostos.Show
End Sub

Private Sub DtcFormaPago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcTipoCompra.SetFocus
End If
End Sub

Private Sub DtcMoneda_Change()
If Me.DtcMoneda.BoundText = "0002" Then
     Me.Label13.Visible = True
     Me.TxtTipoCambio.Visible = True
     strCadena = "SELECT * FROM Tipo_cambio WHERE fecha='" & CVDate(Date) & "'"
     Call ConfiguraRst(strCadena)
     Me.TxtTipoCambio.Text = rst("valor")
     Set rst = Nothing
     Call Resalta(Me.TxtTipoCambio)
Else
    Me.Label13.Visible = False
    Me.TxtTipoCambio.Visible = False
    
End If
End Sub

Private Sub DtcMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Me.DtcMoneda.BoundText = "0002" Then
    Call Resalta(Me.TxtTipoCambio)
Else
    Me.DtcTipoCompra.SetFocus
End If
End Sub

Private Sub DtcTipoCompra_Change()
If Me.DtcTipoCompra.BoundText = "04" Then
    Me.TxtIgv.Visible = False
    Me.lblvalorigv.Visible = False
    
    
Else
    Me.TxtIgv.Visible = True
    Me.lblvalorigv.Visible = True
End If
End Sub

Private Sub DtcTipoCompra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Call Resalta(Me.TxtPercepcion)
    If (Trim(Me.TxtTipoComprobante.Text) = "0003" Or Trim(Me.TxtTipoComprobante.Text) = "0040") Then
        Me.TxtTotal.Text = ""
        Me.TxtTotal.SetFocus
        Exit Sub
    End If
    Me.TxtPercepcion.Text = ""
    Me.TxtPercepcion.SetFocus
End If
End Sub

Private Sub Form_Activate()
'Call Resalta(Me.txtdia)
If FrmComprobantes.Procedencia = Selecionar Then
    Call Resalta(Me.txtserie)
    FrmComprobantes.Procedencia = Neutro
    Exit Sub
End If



End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
KEY_ANULADO = "F"
strModificar = False
igv = KEY_CON_IGV

Me.dtpFechaFind.Value = KEY_FECHA

Me.LblEmpresa.Caption = FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 6) + "***[" + Space(1) + FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 0) + Space(1) + "]***"
Me.lblMes.Caption = FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 2) + Space(2) + FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 3)
Me.TxtTipoCambio.Text = Format(KEY_CAMBIO, "#,##0.00")
Me.txtRucEmpresa.Text = FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 0)


strCadena = "SELECT id_mes as Codigo, descripcion as Descripcion FROM mes  ORDER BY id_mes ASC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMes)
  
strCadena = "SELECT id as Codigo,descripcion as Descripcion FROM forma_pago ORDER BY id DESC"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcFormaPago)
Me.DtcFormaPago.BoundText = "02"

strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda ORDER BY id_moneda ASC "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)

strCadena = "SELECT id_doc as Codigo,doc_des as Descripcion FROM comprobantes ORDER BY doc_abrev"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoComprobante)



  strCadena = "SELECT tipo_compra as Codigo, descripcion as Descripcion FROM tipo_compra ORDER BY tipo_compra ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcTipoCompra)
  Me.DtcTipoCompra.BoundText = "03"
  
  
  Me.DtcMes.BoundText = FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 1)
  Me.DtcMes.Enabled = False
  Me.txtAnio.Text = FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 3)
  
  Call llenarGrid(Me.HfdPersona, FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 0))
  
  
  
  Me.TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
  Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
  Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Me.txtserie.Text = formato_item(Me.txtserie.Text, 3)
    Call Resalta(Me.txtNumero)


End If
End Sub

Private Sub HfdPersona_Click()
If Me.HfdPersona.Rows > 0 Then
      Me.TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
      Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
      Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = True
End If
End Sub

Private Sub HfdPersona_GotFocus()
HookForm Me.HfdPersona
End Sub

Private Sub HfdPersona_LostFocus()
UnHookForm Me.HfdPersona
End Sub
Private Sub HfdPersona_DblClick()
If Me.HfdPersona.Rows > 0 Then
    
   ' FrmDetalleCompra.Show
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    'Case KEY_NEW
    '  Procedencia = nuevo
    ' FrmDetalleMarca.Show
    Case KEY_ANULAR
       If MsgBox("Esta Seguro de Anular este Comprobante", vbQuestion + vbYesNo, "Informacion para el Usuario") = vbYes Then
        strCadena = "UPDATE RegistroComprasDetalle SET anulado='V',NombreCliente='A N U L A D O',RucCliente='',grav1='0',igv1='0',grav2='0',igv2='0',grav3='0',igv3='0',nograv='0',total='0',isc='0',percepcion='0',otros='0' WHERE codigounico='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 29)) & "'"
        CnBd.Execute (strCadena)
         
        Call llenarGrid(Me.HfdPersona, Trim(Me.txtRucEmpresa.Text))
       End If
    Case KEY_UPDATE
       If MsgBox("Esta seguro de Modificar este Registro", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
           strModificar = True
            strCadena = "SELECT * FROM movimiento_compra WHERE  id_compra='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                Procede = modificar
                
                Me.TxtFecha.Text = CVDate(rst("fecha_emision"))
                Me.TxtCancelacion.Text = CVDate(rst("fecha_cancelacion"))
                Me.TxtTipoComprobante.Text = rst("id_doc")
                If rst("id_doc") <> "0050" Then
                    Me.TxtAnioDua.Visible = False
                    Me.txtRuc.Visible = True
                Else
                    Me.txtRuc.Visible = False
                    Me.TxtAnioDua.Visible = False
                End If
                Me.txtserie.Text = rst("serie")
                Me.txtNumero.Text = rst("numero")
                Me.TxtAnioDua.Text = rst("anio_dua")
                Me.txtRuc.Text = rst("id_proveedor")
                Me.txtCliente.Text = rst("nproveedor")
                Me.DtcMoneda.BoundText = rst("id_moneda")
                Me.TxtTipoCambio.Text = rst("tc")
                Me.DtcTipoCompra.BoundText = rst("id_tipo_compra")
                Me.DtcFormaPago.BoundText = rst("id_forma_pago")
                Me.TxtPercepcion.Text = rst("percepcion")
                Me.txtisc.Text = rst("isc")
                If Me.DtcTipoCompra.BoundText = "01" Then
                    Me.TxtValorcompra.Text = rst("valor_venta")
                    Me.TxtIgv.Text = rst("igv")
                End If
                If Me.DtcTipoCompra.BoundText = "02" Then
                    Me.TxtValorcompra.Text = rst("valor_venta")
                    Me.TxtIgv.Text = rst("igv")
                End If
                If Me.DtcTipoCompra.BoundText = "03" Then
                    Me.TxtValorcompra.Text = rst("valor_venta")
                    Me.TxtIgv.Text = rst("igv")
                End If
                If Me.DtcTipoCompra.BoundText = "04" Then
                    Me.TxtValorcompra.Text = rst("total")
                    
                End If
                Me.TxtTotal.Text = rst("total")
                Me.cmdagregar.Caption = "M"
                Me.cmdagregar.BackColor = &H8080FF
                Call Resalta(Me.txtRuc)
                Procede = Neutro
                Exit Sub
            End If
        End If
    Case KEY_DELETE
      If MsgBox("Esta Seguro de eliminar este registro", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        strCadena = "DELETE FROM movimiento_compra WHERE id_compra='" & Val(HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        Call llenarGrid(Me.HfdPersona, Trim(Me.txtRucEmpresa.Text))
      End If
      Case KEY_ACTUALIZAR
             Call llenarGrid(Me.HfdPersona, Trim(Me.txtRucEmpresa.Text))
    Case KEY_EXIT
        Call FrmRegistroCompras.actualizar
        Unload Me
  End Select
End Sub

Private Sub TxtAnioDua_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcMoneda.SetFocus
End If
End Sub

Private Sub TxtCancelacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsDate(Trim(Me.TxtCancelacion.Text)) = False Then
    Me.TxtCancelacion.Text = CVDate(Me.TxtFecha)
    Call Resalta(Me.TxtTipoComprobante)
    Exit Sub
End If
If IsDate(TxtCancelacion.Text) = True Then
    If CVDate(Me.TxtCancelacion.Text) <> CVDate(Me.TxtFecha.Text) And CVDate(Me.TxtCancelacion.Text) <> "" Then
        Me.DtcFormaPago.BoundText = "0004"
    Else
        Me.DtcFormaPago.BoundText = "0001"
    End If
    Call Resalta(Me.TxtTipoComprobante)
    
Else
    Me.TxtCancelacion.SetFocus
End If
End If
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
If Trim(Me.TxtTipoComprobante.Text) = "0001" And Me.txtCliente.Text = "" Then
    MsgBox "Es Obligatorio el Ruc del Cliente"
    Call Resalta(Me.txtRuc)

Else
    Me.DtcMoneda.SetFocus
End If
End Sub




Private Sub TxtFecha_KeyPress(KeyAscii As Integer)
If IsDate(Me.TxtFecha.Text) = True Then
    Me.TxtCancelacion.SetFocus
    
    
Else
    Me.TxtFecha.SetFocus
End If
End Sub

Private Sub txtIgv_Change()
'Dim valor  As Double
'If Trim(Me.TxtTipoComprobante.Text) <> "0003" Then
'valor = Format(Val(Me.TxtValorcompra.Text) * KEY_IGV, "###0.00")
 'If (Val(Me.txtIgv.Text) > 0 And Val(Me.txtIgv.Text) <> Val(valor)) Then
  '      Procedencia = Migv
   '     Calculator.Show
 'End If
 'End If
End Sub

Private Sub TxtIGV_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        Me.TxtTotal.Text = Format(Round(Val(Me.TxtValorcompra.Text) + Val(Me.TxtIgv.Text) + Val(Me.TxtValorCompraNoAfecta.Text) + Val(Me.TxtPercepcion.Text), 2), "###0.00")
        Call Resalta(Me.TxtTotal)
    
End If
End Sub

Private Sub txtisc_Change()
Dim Valor As Double
Valor = Val(Me.txtisc.Text)
If (Val(Me.txtisc.Text) > 0 And Val(Me.txtisc.Text) <> Val(Valor)) Then
    Procedencia = Misc
    Calculator.lblScreen.Caption = Val(Me.txtisc.Text)
    Calculator.Show
End If
End Sub

Private Sub TxtISC_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtisc.Text = Format(Val(Me.txtisc.Text), "###0.00")
    'Call Resalta(Me.TxtValorcompra)
    Me.TxtValorcompra.Text = ""
    Me.TxtValorcompra.SetFocus
End If

End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Me.txtNumero.Text = formato_item(Me.txtNumero.Text, 8)
    strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Trim(Me.TxtTipoComprobante.Text) & "' AND serie='" & Trim(Me.txtserie.Text) & "' AND numero='" & Trim(Me.txtNumero.Text) & "' AND ruc='" & Trim(FrmRegistroCompras.txtRuc.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        MsgBox "Comprobante ya Registrado", vbInformation, "Duplicidad de Comprobante"
        Call Resalta(Me.txtNumero)
        Exit Sub
    End If
    If Trim(Me.TxtTipoComprobante.Text) <> "0050" Then
        Call Resalta(Me.txtRuc)
    Else
        Call Resalta(Me.TxtAnioDua)
    End If
    Exit Sub
    FrmVerificarAnulado.Show
    
     
End If
End Sub

Private Sub TxtPercepcion_Change()
Dim Valor As Double
Valor = Val(Me.TxtPercepcion.Text)
  
If (Val(Me.TxtPercepcion.Text) > 0 And Val(Me.TxtPercepcion.Text) <> Val(Valor)) Then

    Procedencia = MOtro
    Calculator.lblScreen.Caption = Val(Me.TxtPercepcion.Text)
    Calculator.Show
End If
End Sub

Private Sub TxtPercepcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtPercepcion.Text = Format(Val(Me.TxtPercepcion.Text), "###0.00")
    Call Resalta(Me.txtisc)
   'Me.txtisc.Text = ""
   'Me.txtisc.SetFocus
End If
End Sub

Private Sub txtRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtRuc.Text) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.txtCliente.Text = rst("nombre_completo")
    Else
        If Trim(Me.TxtTipoComprobante.Text) = "0001" And Me.txtRuc.Text <> "" Then
            
            Procedencia = 1
            FrmDetallePersona.Show
            FrmDetallePersona.txtRuc.Text = Trim(Me.txtRuc.Text)
            FrmDetallePersona.ChkCliente.Value = 1
            Call FrmDetallePersona.precionar
            Exit Sub
            
         Else
         If Trim(Me.TxtTipoComprobante.Text) = "0001" Then
            Procedencia = buscar
            FrmPersona.Show
            Exit Sub
         Else
           If (Trim(Me.txtRuc.Text) <> "") Then
                    
                Procedencia = 1
                FrmDetallePersona.Show
                FrmDetallePersona.txtRuc.Text = Trim(Me.txtRuc.Text)
                FrmDetallePersona.ChkCliente.Value = 1
                Call FrmDetallePersona.precionar
                Exit Sub
                    
            Else
                
                    Procedencia = buscar
                    FrmPersona.Show
                    Exit Sub
                
            End If
        End If
    End If
    End If
    Call Resalta(Me.txtCliente)
End If
End Sub
Public Sub foco_ruc()
Call Resalta(Me.txtRuc)
End Sub
Public Sub foco_serie()
Call Resalta(Me.txtserie)
End Sub
Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtserie.Text = formato_item(Me.txtserie.Text, 3)
    Call Resalta(Me.txtNumero)
End If
End Sub

Private Sub TxtTipoCambio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcFormaPago.SetFocus
End If
End Sub

Private Sub TxtTipoComprobante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.TxtTipoComprobante.Text = "" Then
        Procedencia = buscar
        FrmComprobantes.Show
    Else
    
   
    
   Me.TxtTipoComprobante.Text = formato_item(Me.TxtTipoComprobante.Text, 4)
   If (Trim(Me.TxtTipoComprobante.Text) = "0050") Then
        Me.txtRuc.Visible = False
        Me.txtCliente.Visible = False
        Me.TxtAnioDua.Visible = True
        Me.Label10.Visible = False
        Me.Label9.Visible = False
        Me.lblaniodua.Visible = True
        Me.DtcTipoCompra.BoundText = "01"
        
    Else
        Me.txtRuc.Visible = True
        Me.txtCliente.Visible = True
        Me.TxtAnioDua.Visible = False
        Me.Label10.Visible = True
        Me.Label9.Visible = True
        Me.lblaniodua.Visible = False
   End If
   Call Resalta(Me.txtserie)
   
    End If
End If
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
Dim vigv As Single
Dim vsubtotal As Single
Dim Total As Single
Total = Val(Me.TxtTotal.Text)

If KeyAscii = 13 Then
Me.TxtTotal.Text = Format(Val(Me.TxtTotal.Text), "###0.000")
    If (Trim(Me.TxtTipoComprobante.Text) = "0003") Then
        If (igv = "si") Then
          '  Me.TxtValorcompra.Text = Format(Round(Val(Me.txtTotal.Text) / (1 + KEY_IGV), 2), "###0.00")
           ' Me.txtIgv.Text = Format(Round(Val(Me.txtTotal.Text) - Val(Me.TxtValorcompra.Text), 2), "###0.00")
             Me.TxtValorCompraNoAfecta.Text = Format(Me.TxtTotal.Text, "###0.000")
            If KEY_ANULADO = "F" And Val(Me.TxtTotal.Text) = 0 Then
                Me.TxtTotal.SetFocus
                Exit Sub
            End If
            Me.CmdCentroCostos.SetFocus
            Exit Sub
        End If
    End If
    If (Trim(Me.TxtTipoComprobante.Text) = "0040") Then
        Me.TxtPercepcion.Text = Format(Val(Me.TxtTotal.Text), "#,##0.00")
    End If
    Me.CmdCentroCostos.SetFocus
End If
End Sub

Private Sub TxtValorcompra_Change()


If Val(Me.TxtValorcompra.Text) > 0 Then
If Trim(Me.TxtTipoComprobante.Text) <> "0003" And Procedencia <> MTC And Procede <> modificar Then
     Procedencia = Mbimponible
     Calculator.lblScreen.Caption = Val(Me.TxtValorcompra.Text)
     Calculator.Show
End If
End If
End Sub

Private Sub TxtValorcompra_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Me.DtcTipoCompra.BoundText = "04" Then
            Me.TxtTotal.Text = Format(Val(Me.TxtValorcompra.Text), "###0.00")
            Call Resalta(Me.TxtTotal)
             
            Exit Sub
        End If
        Me.TxtValorcompra.Text = Format(Val(Me.TxtValorcompra.Text), "###0.00")
        Me.TxtIgv.Text = ""
        Me.TxtIgv.Text = Format(Round(Val(Me.TxtValorcompra.Text) * KEY_IGV, 2), "###0.00")
        Me.TxtTotal.Text = Format(Round(Val(Me.TxtValorcompra.Text) + Val(Me.TxtIgv.Text) + Val(Me.TxtValorCompraNoAfecta.Text) + Val(Me.TxtPercepcion.Text) + Val(Me.txtisc.Text), 2), "###0.00")
        Call Resalta(Me.TxtValorCompraNoAfecta)
        'Call Resalta(Me.txtIgv)
   
End If
End Sub

Private Sub TxtValorCompraNoAfecta_Change()
If Val(Me.TxtValorCompraNoAfecta.Text) > 0 Then
If Trim(Me.TxtTipoComprobante.Text) <> "0003" And Procedencia <> MTC And Procede <> modificar Then
     Procedencia = MbimponibleInafecta
     Calculator.lblScreen.Caption = Val(Me.TxtValorCompraNoAfecta.Text)
     Calculator.Show
End If
End If
End Sub

Private Sub TxtValorCompraNoAfecta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtValorCompraNoAfecta.Text = Format(Val(Me.TxtValorCompraNoAfecta.Text), "###0.00")
    Me.TxtTotal.Text = Format(Val(Me.TxtValorcompra.Text) + Val(Me.TxtIgv.Text) + Val(Me.TxtValorCompraNoAfecta.Text) + Val(Me.TxtPercepcion.Text), "###0.00")
    Call Resalta(Me.TxtIgv)
End If
End Sub



VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmComprasGastos 
   BorderStyle     =   0  'None
   Caption         =   "GASTOS COMPRA"
   ClientHeight    =   8865
   ClientLeft      =   945
   ClientTop       =   630
   ClientWidth     =   13050
   Icon            =   "FrmBuscarProducto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   13050
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtIdCompra 
      Height          =   285
      Left            =   10080
      TabIndex        =   61
      Text            =   "IdCompra"
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   240
      TabIndex        =   28
      Top             =   120
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "OTROS GASTOS"
      TabPicture(1)   =   "FrmBuscarProducto.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ChameleonBtn1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
         Height          =   615
         Left            =   10320
         TabIndex        =   96
         Top             =   7320
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         BTYPE           =   2
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
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmBuscarProducto.frx":0342
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
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2535
         Left            =   360
         TabIndex        =   65
         Top             =   5880
         Width           =   10815
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   10080
            Top             =   2040
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmBuscarProducto.frx":035E
                  Key             =   "(Borrar)"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmBuscarProducto.frx":08F8
                  Key             =   "(Modificar)"
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   1200
            Left            =   9960
            TabIndex        =   67
            Top             =   240
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   2117
            ButtonWidth     =   1111
            ButtonHeight    =   1005
            ImageList       =   "ImageList1"
            DisabledImageList=   "ImageList1"
            HotImageList    =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Borrar"
                  Key             =   "(Eliminar)"
                  ImageKey        =   "(Borrar)"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Modifi"
                  Key             =   "(Modificar)"
                  ImageKey        =   "(Modificar)"
               EndProperty
            EndProperty
            Enabled         =   0   'False
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshgastos 
            Height          =   2055
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   3625
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   8388608
            FixedCols       =   0
            ForeColorFixed  =   8388608
            BackColorBkg    =   16777215
            GridColor       =   8388608
            FocusRect       =   0
            GridLines       =   2
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
         Left            =   360
         TabIndex        =   62
         Top             =   480
         Width           =   10815
         Begin VB.TextBox txtproducto 
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
            Left            =   3480
            TabIndex        =   99
            Top             =   3480
            Width           =   5055
         End
         Begin VB.TextBox txtcodigoprod 
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
            Left            =   1800
            TabIndex        =   98
            Top             =   3480
            Width           =   1575
         End
         Begin VB.TextBox txtTipoCambio 
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
            Left            =   4320
            TabIndex        =   78
            Top             =   1080
            Width           =   1000
         End
         Begin VB.TextBox Txtdni 
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
            Left            =   1800
            TabIndex        =   69
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtmonto 
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
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   130678785
            CurrentDate     =   41126
         End
         Begin VB.CommandButton cmdagregar 
            BackColor       =   &H008080FF&
            Caption         =   "AGREGAR GASTO"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   4800
            Width           =   1575
         End
         Begin VB.TextBox txtserie 
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
            Left            =   5880
            TabIndex        =   22
            Top             =   240
            Width           =   1000
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
            Height          =   525
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Top             =   4200
            Width           =   8175
         End
         Begin VB.TextBox txtnumero 
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
            Left            =   6960
            TabIndex        =   23
            Top             =   240
            Width           =   1455
         End
         Begin MSDataListLib.DataCombo DtcComprobante 
            Height          =   315
            Left            =   1800
            TabIndex        =   72
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
               Name            =   "Tahoma"
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
            Left            =   4320
            TabIndex        =   76
            Top             =   2160
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
            Format          =   130678785
            CurrentDate     =   41126
         End
         Begin MSDataListLib.DataCombo DtcAfectoIgv 
            Height          =   315
            Left            =   1800
            TabIndex        =   79
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
            Left            =   1800
            TabIndex        =   80
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
            Left            =   1800
            TabIndex        =   85
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
               Name            =   "Tahoma"
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
            TabIndex        =   100
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
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            TabIndex        =   102
            Top             =   3840
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
            TabIndex        =   101
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SERVICIO :"
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
            TabIndex        =   97
            Top             =   3465
            Width           =   825
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TIPO GASTO :"
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
            TabIndex        =   86
            Top             =   3000
            Width           =   1020
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DESCRIPCION :"
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
            Left            =   240
            TabIndex        =   84
            Top             =   4200
            Width           =   1140
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PERIODO :"
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
            Left            =   585
            TabIndex        =   83
            Top             =   2640
            Width           =   795
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "VENCIMI:"
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
            Left            =   3480
            TabIndex        =   82
            Top             =   2280
            Width           =   690
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "T.C :"
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
            Left            =   3840
            TabIndex        =   81
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AFECTO IGV :"
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
            Left            =   375
            TabIndex        =   77
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FECHA  :"
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
            TabIndex        =   75
            Top             =   2160
            Width           =   645
         End
         Begin VB.Label lblcliente 
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
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3600
            TabIndex        =   71
            Top             =   720
            Width           =   4875
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RUC/DNI  :"
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
            Left            =   585
            TabIndex        =   70
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MONTO  :"
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
            Left            =   675
            TabIndex        =   68
            Top             =   1080
            Width           =   705
         End
         Begin VB.Label Label26 
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
            Left            =   630
            TabIndex        =   64
            Top             =   1800
            Width           =   750
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DOCUMENTO :"
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
            Left            =   315
            TabIndex        =   63
            Top             =   360
            Width           =   1065
         End
      End
      Begin VB.Frame Frame6 
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
         Height          =   1815
         Left            =   -74760
         TabIndex        =   59
         Top             =   6600
         Width           =   10335
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfVinculados 
            Height          =   1335
            Left            =   120
            TabIndex        =   60
            Top             =   360
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   2355
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   8388608
            FixedCols       =   0
            ForeColorFixed  =   8388608
            BackColorBkg    =   16777215
            GridColor       =   8388608
            FocusRect       =   0
            GridLines       =   2
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
         Begin MSComctlLib.ImageList ImgIconos 
            Left            =   7200
            Top             =   885
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
                  Picture         =   "FrmBuscarProducto.frx":0E92
                  Key             =   "(Aceptar)"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmBuscarProducto.frx":11AE
                  Key             =   "(Eliminar)"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmBuscarProducto.frx":160E
                  Key             =   "(Inicio)"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmBuscarProducto.frx":1A6E
                  Key             =   "(Modificar)"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmBuscarProducto.frx":1D8A
                  Key             =   "(Nuevo)"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmBuscarProducto.frx":21EA
                  Key             =   "(Quitar)"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmBuscarProducto.frx":2506
                  Key             =   "(Salir)"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmBuscarProducto.frx":2966
                  Key             =   "(Red)"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmBuscarProducto.frx":2DC6
                  Key             =   "(Grabar)"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmBuscarProducto.frx":36A6
                  Key             =   "(Agregar)"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmBuscarProducto.frx":39C2
                  Key             =   "(Buscar)"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmBuscarProducto.frx":3CDE
                  Key             =   "(Cancelar)"
               EndProperty
            EndProperty
         End
         Begin ComCtl3.CoolBar ClbAcciones 
            Height          =   870
            Left            =   7920
            TabIndex        =   73
            Top             =   600
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
               TabIndex        =   74
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
                     ImageKey        =   "(Cancelar)"
                  EndProperty
               EndProperty
               OLEDropMode     =   1
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TOTALES"
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
         Left            =   -69480
         TabIndex        =   54
         Top             =   4920
         Width           =   5055
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
            TabIndex        =   58
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
            TabIndex        =   57
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
            TabIndex        =   56
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
            TabIndex        =   55
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2055
         Left            =   -74760
         TabIndex        =   50
         Top             =   4560
         Width           =   5055
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
            TabIndex        =   93
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
            TabIndex        =   91
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
            TabIndex        =   87
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
            TabIndex        =   92
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
            TabIndex        =   88
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
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
         Left            =   -69480
         TabIndex        =   40
         Top             =   360
         Width           =   5055
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
            TabIndex        =   94
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
            TabIndex        =   95
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
            TabIndex        =   90
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
            TabIndex        =   89
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
            TabIndex        =   49
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
            TabIndex        =   48
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
            TabIndex        =   47
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
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
         Left            =   -74760
         TabIndex        =   29
         Top             =   360
         Width           =   5055
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
            Top             =   360
            Width           =   1065
         End
      End
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   8820
      Left            =   0
      Top             =   0
      Width           =   13020
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
    Me.txtigv.Text = Format(rst("igv"), "###0.00")
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
    Me.txtflete.Text = Format(rst("flete"), "###0.00")
    Me.txtcerfgases.Text = Format(rst("certifi_gases"), "###0.00")
    Me.txtcomision.Text = Format(rst("comision"), "###0.00")
    Me.txtigvcomision.Text = Format(rst("igv_comision"), "###0.00")
    Me.txtfob.Text = Format(rst("fob"), "###0.00")
    Me.txtseguro.Text = Format(rst("seguro"), "###0.00")
    Me.txtcif.Text = Format(rst("cif"), "###0.00")
    Me.txttc.Text = Format(rst("tc"), "###0.00")
  Else
    Me.txttc.Text = Format(KEY_CAMBIO, "###0.00")
End If

Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Private Sub ChameleonBtn1_Click()
Unload Me
End Sub

Private Sub chkConvertirSoles_Click()
If Me.chkConvertirSoles.Value = 1 Then
   Me.txtfob.Text = Val(Me.txtfob.Text) * Val(Me.txttc.Text)
   Me.txtseguro.Text = Val(Me.txtseguro.Text) * Val(Me.txttc.Text)
   Me.txtflete.Text = Val(Me.txtflete.Text) * Val(Me.txttc.Text)
 Else
   Me.txtfob.Text = Val(Me.txtfob.Text) / Val(Me.txttc.Text)
   Me.txtseguro.Text = Val(Me.txtseguro.Text) / Val(Me.txttc.Text)
   Me.txtflete.Text = Val(Me.txtflete.Text) / Val(Me.txttc.Text)
End If
End Sub

Private Sub CmdAgregar_Click()
Dim cod_identidad As String * 1
Dim valor_venta As Double
Dim igv As Double
Dim Total As Double
If Me.cmdagregar.Caption = "Modificar" Then
    strCadena = "UPDATE movimiento_compra_gasto SET id_doc='" & Me.DtcComprobante.BoundText & "',serie='" & Me.txtserie.Text & "',numero='" & Me.txtnumero.Text & "', id_persona='" & Me.Txtdni.Text & "' WHERE id_gasto='" & Me.mshgastos.TextMatrix(Me.mshgastos.Row, 0) & "' ANd ruc='" & KEY_RUC & "'"
    CnBd.Execute (strCadena)
     
    Call llenar_gastos(Me.mshgastos, Val(Me.txtIdCompra.Text))
    Me.cmdagregar.Caption = "Agregar"
    Me.txtserie.Text = ""
    Me.txtnumero.Text = ""
    Me.Txtdni.Text = ""
    Me.lblcliente.Caption = ""
    Me.txtmonto.Text = 0#
    Me.dtpfecha.Value = KEY_FECHA
    Me.txtdescripcion.Text = ""
    Me.DtcComprobante.SetFocus
    Exit Sub
End If

strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & Me.DtcComprobante.BoundText & "' AND serie='" & Me.txtserie.Text & "' AND numero='" & Me.txtnumero.Text & "' AND id_proveedor='" & Me.Txtdni.Text & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
If Len(Trim(Me.Txtdni.Text)) = 8 Then
    cod_identidad = 1
End If
If Len(Trim(Me.Txtdni.Text)) = 11 Then
    cod_identidad = 6
End If
If Len(Trim(Me.Txtdni.Text)) <> 8 And Len(Trim(Me.Txtdni.Text)) <> 11 Then
    cod_identidad = 0
End If

If Me.DtcAfectoIgv.BoundText = "SI" Then
    valor_venta = Val(Me.txtmonto.Text) / (KEY_IGV + 1)
    igv = Val(Me.txtmonto.Text) - valor_venta
Else
    valor_venta = 0
    igv = 0
End If


        
        If Me.DtcMoneda.BoundText = "00001" Then
                in_cta_compra = KEY_CTA_COMPRA_SOLES
           Else
                in_cta_compra = KEY_CTA_COMPRA_SOLES
           End If
           
        If KEY_CONTABILIDAD = "si" Then
           If put_verifica_cuenta_contable(Me.DtcComprobante.BoundText, Trim(Me.txtserie.Text), Trim(Me.txtnumero.Text), in_cta_compra) = False Then
              Exit Sub
           End If
        End If
       
       
        
        
       ' strCadena = "call P_insert_compra_test('" & Me.DtcTipoDoc.BoundText & "','" & Me.DtcAlmacen.BoundText & "','" & Format(CVDate(Me.TxtFecha_emision.Text), "YYYY-mm-dd") & "','" & Format(CVDate(Me.txtfecha_Vencimiento.Text), "YYYY-mm-dd") & "','02'," & _
        "'" & Me.DtTipoCompra.BoundText & "','--','" & Me.DtcMoneda.BoundText & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Trim(Me.txtserie.Text) & "'," & _
        "'" & Format(Trim(Me.TxtNumeroDoc.Text), "00000000") & "','" & cod_identidad & "','" & Trim(Me.TxtRuc.Text) & "','" & UCase(Me.TxtProveedor.Text) & "','" & Trim(Me.txttc.Text) & "'," & _
        "'0','" & Val(Me.LblValorVenta.Text) & "','" & Val(Me.LblIgv.Text) & "','" & Val(Me.lblISC.Text) & "','0','" & Val(Me.TxtPecepcion.Text) & "','0','" & Val(Me.lblExonerado.Text) & "','0','" & Val(Me.lblTotal.Text) & "','" & Val(Me.lblTotal.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtObservacion.Text) & "','" & Me.DtcTipo.BoundText & "','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & KEY_RUC & "')"
        'CnBd.Execute (strCadena)
        
        
        
        'strCadena = "P_insert_compra('" & Me.DtcComprobante.BoundText & "','" & KEY_ALM & "','" & Format(Me.dtpfecha.Value, "YYYY-mm-dd") & "','" & Format(Me.DtpVencimiento.Value, "YYYY-mm-dd") & "','02'," & _
        "'" & Me.DtcTipoCompra.BoundText & "','','" & Me.DtcMoneda.BoundText & "','" & formato_item(Me.txtmes.Text, 2) & "','" & formato_item(Me.txtanio.Text, 4) & "','" & Trim(Me.txtserie.Text) & "'," & _
        "'" & Trim(Me.txtnumero.Text) & "','" & cod_identidad & "','" & Trim(Me.Txtdni.Text) & "','" & UCase(Me.lblcliente.Caption) & "','" & Val(Me.txtTipoCambio.Text) & "'," & _
        "'0','" & valor_venta & "','" & igv & "','0','0','0','0','0','0','" & Val(Me.txtmonto.Text) & "','" & Val(Me.txtmonto.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtdescripcion.Text) & "','" & KEY_RUC & "')"
        'CnBd.Execute (strCadena)
        
        
        strCadena = "P_insert_compra_test('" & Me.DtcComprobante.BoundText & "','" & KEY_ALM & "','" & Format(Me.dtpfecha.Value, "YYYY-mm-dd") & "','" & Format(DtpVencimiento.Value, "YYYY-mm-dd") & "','02'," & _
        "'02','','" & Me.DtcMoneda.BoundText & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & Trim(Me.txtserie.Text) & "'," & _
        "'" & Format(Trim(Me.txtnumero.Text), "00000000") & "','" & cod_identidad & "','" & Trim(Me.Txtdni.Text) & "','" & UCase(Me.lblcliente.Caption) & "','" & Trim(Me.txtTipoCambio.Text) & "'," & _
        "'0','" & valor_venta & "','" & igv & "','0','0','0','0','0','0','" & Val(Me.txtmonto.Text) & "','" & Val(Me.txtmonto.Text) & "','" & KEY_USUARIO & "','" & Trim(Me.txtdescripcion.Text) & "','01','" & Me.DtcPeriodo.BoundText & "','" & in_cta_compra & "','" & KEY_RUC & "')"
        CnBd.Execute (strCadena)
        
        
           id_compra = LastRegistroRUC("movimiento_compra", "id_compra")
        
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,ivap,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,ruc) VALUES ('" & id_compra & "','" & Trim(Me.txtcodigoprod.Text) & "','1','" & Val(Me.txtmonto.Text) & "'," & _
           "'0','0','0','" & Val(Me.txtmonto.Text) / (1 + KEY_IGV) & "','0','" & Val(Me.txtmonto.Text) / (1 + KEY_IGV) * KEY_IGV & "', " & _
           "'0','0','0','" & Val(Me.txtmonto.Text) / (1 + KEY_IGV) & "','0','" & Val(Me.txtmonto.Text) & "','" & Val(Me.txtmonto.Text) & "','0','" & KEY_ALM & "','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
        
        

        
'02----------------guardar en detalle documento Compra-----------
 If KEY_CONTABILIDAD = "si" And Me.DtcComprobante.BoundText <> "0089" Then
    strCadena = "p_insert_compra_emitido_ii('" & id_compra & "')"
    Call Execute_Sql(strCadena)
 End If
 
 
 
 
 
 
        
        
        
        
        
strCadena = "INSERT INTO movimiento_compra_gasto(id_compra,id_persona,id_doc,serie,numero,monto,fecha,descripcion,ruc)VALUES " & _
" ('" & Val(Me.txtIdCompra.Text) & "','" & Me.Txtdni.Text & "','" & Me.DtcComprobante.BoundText & "','" & Me.txtserie.Text & "','" & Me.txtnumero.Text & "','" & Val(Me.txtmonto.Text) & "','" & Format(Me.dtpfecha.Value, "YYYY-mm-dd") & "','" & Me.txtdescripcion.Text & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)

Call llenar_gastos(Me.mshgastos, Val(Me.txtIdCompra.Text))
Me.txtserie.Text = ""
Me.txtnumero.Text = ""
Me.Txtdni.Text = ""
Me.lblcliente.Caption = ""
Me.txtmonto.Text = 0#
Me.dtpfecha.Value = KEY_FECHA
Me.txtdescripcion.Text = ""
Me.DtcComprobante.SetFocus
Me.lblcuenta_contable.Caption = ""
Me.lblcuenta_detalle.Caption = ""



Exit Sub
Else
    MsgBox "COMPROBANTE YA REGISTRADO, IMPOSIBLE GUARDAR ", vbInformation, KEY_EMPRESA
End If
Set rst = Nothing
End Sub




Private Sub DtcAfectoIgv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcMoneda.SetFocus
End If
End Sub

Private Sub DtcComprobante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtserie)
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
    Call Resalta(Me.txtdescripcion)
End If
End Sub

Private Sub DtpFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtdescripcion)
End If
End Sub

Private Sub Form_Load()
Dim nombreperiodo As String
CenterForm Me
Me.Top = 100
Me.dtpfecha.Value = KEY_FECHA


Me.txttc.Text = KEY_CAMBIO
strCadena = "SELECT id_doc as Codigo,doc_des as Descripcion FROM comprobantes ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComprobante)
Me.DtcComprobante.BoundText = "0001"

strCadena = "SELECT tipo_compra as Codigo,descripcion as Descripcion FROM tipo_compra WHERE tipo_compra<>'01'"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoCompra)

strCadena = "SELECT igv as Codigo,igv as Descripcion FROM afecto_igv ORDER BY igv"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAfectoIgv)

strCadena = "SELECT id_moneda as Codigo,descripcion as Descripcion FROM moneda ORDER BY id_moneda"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcMoneda)

strCadena = "SELECT id as Codigo,CONCAT(Nombre,'-',Ejercicio) as Descripcion FROM con_periodo order by id"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcPeriodo)
  Me.DtcPeriodo.BoundText = get_periodo_actual(KEY_FECHA)




strCadena = "SELECT * FROM movimiento_compra WHERE id_doc='" & FrmCompras.DtcTipoDoc.BoundText & "' AND serie='" & FrmCompras.txtserie.Text & "' AND numero='" & FrmCompras.TxtNumeroDoc.Text & "' AND id_proveedor='" & FrmCompras.TxtRuc.Text & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txtIdCompra.Text = rst("id_compra")
End If
Call llenar_gastos(Me.mshgastos, Val(Me.txtIdCompra.Text))
strCadena = "SELECT * FROM registro_compras WHERE mes='" & formato_item(Month(KEY_FECHA), 2) & "' AND anio='" & Year(KEY_FECHA) & "' AND ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    nombreperiodo = "REGISTRO COMPRAS:" + Space(1) + nombre_mes(Month(KEY_FECHA))
    strCadena = "INSERT INTO registro_compras(ruc,mes,anio,descripcion,id_estado)VALUES('" & KEY_RUC & "','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & nombreperiodo & "','01')"
    CnBd.Execute (strCadena)
    
End If

Me.txtTipoCambio.Text = Format(KEY_CAMBIO, "#,##0.00")

If FrmCompras.DtTipoCompra.BoundText = "01" Then
    Call llenarImportacion(Me.HfVinculados, Val(Me.txtIdCompra.Text))
    Call llenar_productos(HfVinculados, Val(Me.txtIdCompra.Text))
End If
End Sub
Private Sub llenar_gastos(ByVal Grilla As MSHFlexGrid, ByVal id_compra As Double)
Dim Total As Double
strCadena = "SELECT G.id_gasto,CONCAT(C.doc_abrev,':',serie,'-',numero) as comprobante,G.fecha,P.nombre_completo,G.monto,G.descripcion FROM movimiento_compra_gasto G,comprobantes C,persona P WHERE id_compra='" & id_compra & "' AND G.ruc='" & KEY_RUC & "' AND G.id_doc=C.id_doc AND G.id_persona=P.dni"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    FrmCompras.lblgastos.Text = Format(0, "###0.00")
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
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 2000
           Grilla.ColWidth(3) = 3000
           Grilla.ColWidth(4) = 800
           Grilla.ColWidth(5) = 2500
        Next
         cabecera = "IDGASTO" & vbTab & "FECHA" & vbTab & "COMPROBANTE" & vbTab & "CLIENTE/PROVEEDOR" & vbTab & "MONTO" & vbTab & "DESCRIPCION"
         Grilla.AddItem cabecera
         For K = 0 To 5
                                Grilla.col = K
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next K
                            
        rst.MoveFirst
        Total = 0
        For i = 1 To rst.RecordCount
            Total = Total + rst("monto")
             Fila = rst("id_gasto") & vbTab & rst("fecha") & vbTab & rst("comprobante") & vbTab & rst("nombre_completo") & vbTab & Format(rst("monto"), "#,##0.00") & vbTab & rst("descripcion")
             Grilla.AddItem Fila
             Fila = ""
        rst.MoveNext
        Next i
         cabecera = "" & vbTab & "" & vbTab & "" & vbTab & "=========================" & vbTab & Format(Total, "#,##0.00") & vbTab & ""
         Grilla.AddItem cabecera
          For K = 0 To 5
                                Grilla.col = K
                                Grilla.Row = i
                                Grilla.CellBackColor = &HC0C0FF
                            Next K
             FrmCompras.lblgastos.Text = Format(Total, "###0.00")
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub
Private Sub lbltotalderechos_Change()
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lbltotalgeneral.Caption = Format(totalgeneral, "###0.00")
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
         For K = 0 To 3
                                Grilla.col = K
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next K
                            
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
If Val(Me.mshgastos.TextMatrix(Me.mshgastos.Row, 0)) > 0 Then
    Me.Toolbar1.Enabled = True
Else
    Me.Toolbar1.Enabled = False
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_SAVE
        strCadena = "SELECT * FROM movimiento_compra_importacion WHERE id_compra='" & Val(Me.txtIdCompra.Text) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "INSERT INTO movimiento_compra_importacion (id_compra,ruc)VALUES('" & Val(Me.txtIdCompra.Text) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
        End If
        strCadena = "UPDATE movimiento_compra_importacion SET ad_valorem='" & Val(Me.txtadvalorem.Text) & "',sobretasa='" & Val(Me.txtsobretasa.Text) & "',isc='" & Val(Me.txtisc.Text) & "'," & _
        " igv='" & Val(Me.txtigv.Text) & "',ipm='" & Val(Me.txtipm.Text) & "',serv_despacho='" & Val(Me.txtservdespacho.Text) & "',percepcion_igv='" & Val(Me.txtpercepcionigv.Text) & "'," & _
        "derecho_especifico='" & Val(Me.txtderechoespecifico.Text) & "',almacenaje_descarga='" & Val(Me.txtalmacenaje.Text) & "',gastos_operativos='" & Val(Me.txtgastosoperativos.Text) & "'," & _
        "terminal='" & Val(Me.txtterminal.Text) & "',gremios_maritimos='" & Val(Me.txtgremiosmaritimos.Text) & "',cias_navieras_gastos='" & Val(Me.txtciasnaviera.Text) & "'," & _
        "conduccion='" & Val(Me.txtconduccion.Text) & "',gastos_despacho='" & Val(Me.txtgastosdepacho.Text) & "',flete='" & Val(Me.txtflete.Text) & "',certifi_gases='" & Val(Me.txtcerfgases.Text) & "'," & _
        "comision='" & Val(Me.txtcomision.Text) & "',igv_comision='" & Val(Me.txtigvcomision.Text) & "',tc='" & Val(Me.txttc.Text) & "',fob='" & Val(Me.txtfob.Text) & "',seguro='" & Val(Me.txtseguro.Text) & "',cif='" & Val(Me.txtcif.Text) & "' WHERE id_compra='" & Val(Me.txtIdCompra.Text) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        FrmCompras.lblImportacion.Text = Format(Val(Me.lbltotalgeneral.Caption), "###0.00")
        Me.TlbAcciones.Buttons(KEY_SAVE).Enabled = False
    Case KEY_CANCELAR
        Unload Me
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_DELETE
        If MsgBox("Esta Seguro de eliminar este Registro", vbYesNo + vbQuestion, KEY_EMPRESA) = vbYes Then
            strCadena = "DELETE FROM movimiento_compra_gasto WHERE id_gasto='" & Me.mshgastos.TextMatrix(Me.mshgastos.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
            CnBd.Execute (strCadena)
             
        End If
        Call llenar_gastos(Me.mshgastos, Val(Me.txtIdCompra.Text))
    Case KEY_UPDATE
        If MsgBox("Esta Seguro de Modificar este Registro", vbYesNo + vbQuestion, KEY_EMPRESA) = vbYes Then
           strCadena = "SELECT * FROM movimiento_compra WHERE id_gasto='" & Me.mshgastos.TextMatrix(Me.mshgastos.Row, 0) & "' AND ruc='" & KEY_RUC & "'"
           Call ConfiguraRst(strCadena)
           If rst.RecordCount > 0 Then
                Me.DtcComprobante.BoundText = rst("id_doc")
                Me.txtserie.Text = rst("serie")
                Me.txtnumero.Text = rst("numero")
                Me.Txtdni.Text = rst("id_persona")
                Me.txtmonto.Text = Format(rst("monto"), "#,##0.00")
                Me.dtpfecha.Value = rst("fecha")
                Me.txtdescripcion.Text = rst("descripcion")
                Me.lblcliente.Caption = BDBuscarCampo("persona", "nombre_completo", "dni", Trim(Me.Txtdni.Text))
                Me.cmdagregar.Caption = "Modificar"
                Me.cmdagregar.SetFocus
            Else
                
           End If
        End If
End Select
End Sub

Private Sub txtadvalorem_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.txtigv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")
End Sub

Private Sub txtalmacenaje_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text)
Me.lbltotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcTipoCompra.SetFocus
End If
End Sub

Private Sub txtcerfgases_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lbltotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtciasnaviera_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lbltotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub TxtcodigoProd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM producto where id_producto='" & Trim(Me.txtcodigoprod.Text) & "' and ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
       Me.txtcodigoprod.Text = Trim(Me.txtcodigoprod.Text)
       Me.txtproducto.Text = rst("nombre_prod")
    Else
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
    End If
End If
End Sub

Private Sub txtcomision_Change()
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lbltotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtconduccion_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lbltotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtderechoespecifico_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.txtigv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub

Private Sub txtdescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdagregar.SetFocus
End If
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.Txtdni.Text) & "'"
   Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.lblcliente.Caption = UCase(rst("nombre_completo"))
        Call Resalta(Me.txtmonto)
    Else
        If Len(Me.Txtdni.Text) > 7 And Len(Me.Txtdni.Text) < 12 Then
        Procedencia = nuevo
        FrmDetallePersona.Show
        FrmDetallePersona.TxtRuc.Text = Trim(Me.Txtdni)
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

Private Sub txtflete_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lbltotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtFleteImportacion_Change()
Me.txtcif.Text = Format(Val(Me.txtfob.Text) + Val(Me.txtFleteImportacion.Text) + Val(Me.txtseguro.Text), "###0.00")
End Sub

Private Sub txtfob_Change()
Me.txtcif.Text = Format(Val(Me.txtfob.Text) + Val(Me.txtFleteImportacion.Text) + Val(Me.txtseguro.Text), "###0.00")
End Sub

Private Sub txtgastosdepacho_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lbltotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtgastosoperativos_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lbltotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtgremiosmaritimos_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")

totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lbltotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtIgv_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.txtigv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub

Private Sub txtigvcomision_Change()
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lbltotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub txtipm_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.txtigv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub

Private Sub txtisc_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.txtigv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub



Private Sub TxtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.DtcAfectoIgv.SetFocus
End If
End Sub

Private Sub TxtNotadebito_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text)
Me.lbltotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtnumero.Text = formato_item(Me.txtnumero.Text, 8)
    Call Resalta(Me.Txtdni)
End If
End Sub

Private Sub txtpercepcionigv_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.txtigv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub

Private Sub txtseguro_Change()
Me.txtcif.Text = Format(Val(Me.txtfob.Text) + Val(Me.txtFleteImportacion.Text) + Val(Me.txtseguro.Text), "###0.00")
End Sub

Private Sub TxtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtserie.Text = formato_item(Me.txtserie.Text, 3)
    Call Resalta(Me.txtnumero)
End If
End Sub

Private Sub txtservdespacho_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.txtigv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub

Private Sub txtsobretasa_Change()
totalderechos = Val(Me.txtadvalorem.Text) + Val(Me.txtsobretasa.Text) + Val(Me.txtisc.Text) + Val(Me.txtigv.Text) + Val(Me.txtipm.Text) + Val(Me.txtservdespacho.Text) + Val(Me.txtpercepcionigv.Text) + Val(Me.txtderechoespecifico.Text)
Me.lbltotalderechos.Caption = Format(totalderechos, "###0.00")

End Sub

Private Sub txtterminal_Change()
Me.lbltotalgastos.Caption = Format(Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text) + Val(Me.TxtNotadebito.Text), "###0.00")
totalgeneral = Val(Me.lbltotalderechos.Caption) + Val(Me.txtalmacenaje.Text) + Val(Me.txtgastosoperativos.Text) + Val(Me.txtterminal.Text) + Val(Me.txtgremiosmaritimos.Text) + Val(Me.txtciasnaviera.Text) + Val(Me.txtconduccion.Text) + Val(Me.txtgastosdepacho.Text) + Val(Me.txtflete.Text) + Val(Me.txtcerfgases.Text) + Val(Me.txtcomision.Text) + Val(Me.txtigvcomision.Text)
Me.lbltotalgeneral.Caption = Format(totalgeneral, "###0.00")

End Sub
